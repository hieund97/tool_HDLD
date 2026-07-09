# -*- coding: utf-8 -*-
"""
Đổi tên PDF hợp đồng trong folder 'PLHD doi ten ok 3' theo rule:
    STT_PLHĐ_CCCD_Họ và tên     (ví dụ: 007_PLHĐ_064089007642_Lê Anh Dũng)

- Số CCCD lấy từ NỘI DUNG PDF (OCR dòng "Số CMND/CCCD" của Người lao động).
- Đối chiếu cột E (Căn cước) trong file Excel sheet ĐL02 để lấy:
    STT   = cột A
    Họ tên = cột C
- File KHÔNG bị đổi tên tại chỗ. Bản đã đổi tên được COPY sang folder 'output_PLHĐ'.
- In tiến trình ra terminal + tổng hợp kết quả vào 'ket_qua_doi_ten_PLHD3.csv'.

Chạy:  python3 doi_ten_PLHD3.py

Yêu cầu: pdftoppm (poppler), tesseract + gói tiếng Việt (vie),
         pip: pytesseract pillow openpyxl
"""
import os, re, subprocess, glob, csv, shutil
from collections import defaultdict, Counter
import pytesseract
from pytesseract import Output
from PIL import Image, ImageOps
import openpyxl

BASE   = os.path.dirname(os.path.abspath(__file__))
FOLDER = os.path.join(BASE, "PLHD doi ten ok 3")
XLSX   = os.path.join(BASE, "BM_ĐL06_Danh sách sale bán độc quyền Quỹ Vin của ĐL_V3 (1).xlsx")
OUTDIR = os.path.join(BASE, "output_PLHĐ")
OUTCSV = os.path.join(BASE, "ket_qua_doi_ten_PLHD3.csv")
SHEET  = "ĐL02"
EMP    = "040191008474"   # CCCD chủ DN (Bên A) - luôn bỏ qua
DPI    = 400
TMP    = "/tmp/rnplhd"


# ---------------------------------------------------------------- Excel
def load_excel():
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb[SHEET]
    m = {}
    for row in ws.iter_rows(min_row=9, max_col=5, values_only=True):
        stt, _, name, _, cccd = row
        if cccd is None:
            continue
        c = re.sub(r"\D", "", str(cccd))
        if not c:
            continue
        m[c] = ("" if stt is None else str(stt).strip(),
                "" if name is None else str(name).strip())
    wb.close()
    return m


# ---------------------------------------------------------------- OCR helpers
def render(pdf, dpi=DPI):
    for f in glob.glob(TMP + "-*.jpg"):
        os.remove(f)
    subprocess.run(["pdftoppm", "-r", str(dpi), "-jpeg", pdf, TMP],
                   check=True, stderr=subprocess.DEVNULL)
    return sorted(glob.glob(TMP + "-*.jpg"))


def oriented(im):
    """Tự xoay ảnh scan bị nghiêng/ngược bằng OSD của tesseract."""
    try:
        osd = pytesseract.image_to_osd(im)
        rot = int(re.search(r"Rotate: (\d+)", osd).group(1))
        if rot:
            return im.rotate(-rot, expand=True)
    except Exception:
        pass
    return im


def nums12(text):
    return [re.sub(r"\D", "", n)
            for n in re.findall(r"\d[\d ]{7,}\d", text)
            if len(re.sub(r"\D", "", n)) == 12]


def value_crops(im):
    """Trả về các vùng ảnh chứa PHẦN SỐ của dòng 'Số CMND/CCCD' (bỏ 'Ngày cấp')."""
    d = pytesseract.image_to_data(im, lang="vie", output_type=Output.DICT)
    n = len(d["text"]); W, H = im.size
    lines = defaultdict(list)
    for i in range(n):
        lines[(d["block_num"][i], d["par_num"][i], d["line_num"][i])].append(i)
    out = []
    for idxs in lines.values():
        idxs = sorted(idxs, key=lambda i: d["left"][i])
        lt = " ".join(d["text"][i] for i in idxs)
        if not re.search(r"CMND|CCCD", lt, re.I):
            continue
        y0 = min(d["top"][i] for i in idxs) - 6
        y1 = max(d["top"][i] + d["height"][i] for i in idxs) + 6
        x_start = None; x_end = W
        for i in idxs:
            if ":" in d["text"][i]:
                x_start = d["left"][i] + d["width"][i]
            if re.search(r"Ng[aà]y", d["text"][i], re.I):
                x_end = d["left"][i]; break
        if x_start is None:
            for i in idxs:
                if re.search(r"CCCD|CMND", d["text"][i], re.I):
                    x_start = d["left"][i] + d["width"][i]
        if x_start is None:
            x_start = int(W * 0.3)
        out.append((max(x_start - 5, 0), max(y0, 0),
                    min(x_end + 5, W), min(y1, H)))
    return out


def digitocr(crop):
    im = crop.convert("L")
    im = im.resize((im.width * 4, im.height * 4))
    im = ImageOps.autocontrast(im)
    res = []
    for psm in (7, 8, 6):
        s = pytesseract.image_to_string(
            im, config=f"--psm {psm} -c tessedit_char_whitelist=0123456789")
        res += re.findall(r"\d{12}", re.sub(r"[^\d]", " ", s))
    return res


def extract_worker(pdf, ccset):
    """Trả về (cccd_khớp_DS | None, danh_sách_worker, phương_pháp)."""
    imgs = render(pdf)
    pages_im, all12 = [], []
    for j in imgs:
        im = oriented(Image.open(j).convert("RGB"))
        pages_im.append(im)
        all12 += nums12(pytesseract.image_to_string(im, lang="vie"))
    all12 = list(dict.fromkeys(all12))
    workers = [c for c in all12 if c != EMP]
    inx = [c for c in workers if c in ccset]
    if len(inx) == 1:
        return inx[0], workers, "full-ocr"
    # fallback: OCR riêng vùng số
    cands = []
    for im in pages_im:
        for box in value_crops(im):
            cands += digitocr(im.crop(box))
    cands = [c for c in cands if c != EMP]
    workers2 = list(dict.fromkeys(workers + cands))
    inx2 = [c for c in workers2 if c in ccset]
    if len(inx2) == 1:
        return inx2[0], workers2, "digit-crop"
    return None, workers2, "none"


# ---------------------------------------------------------------- main
def main():
    mapping = load_excel()
    ccset = set(mapping)
    os.makedirs(OUTDIR, exist_ok=True)
    files = sorted(f for f in os.listdir(FOLDER) if f.lower().endswith(".pdf"))
    total = len(files)
    print(f"Đã nạp {len(ccset)} CCCD từ Excel. Tổng file PDF: {total}")
    print(f"Folder xuất: {OUTDIR}\n")

    # ---- Pha 1: OCR đọc CCCD từng file
    plan = []  # [src, cccd, stt, name, target, method, status]
    for idx, f in enumerate(files, 1):
        try:
            c, workers, method = extract_worker(os.path.join(FOLDER, f), ccset)
        except Exception as e:
            plan.append([f, "", "", "", "", "err", "LỖI_OCR:" + str(e)[:60]])
            print(f"[{idx}/{total}] ✗ {f}  -> LỖI OCR: {str(e)[:60]}")
            continue
        if c:
            stt, name = mapping[c]
            target = f"{stt}_PLHĐ_{c}_{name}.pdf"
            plan.append([f, c, stt, name, target, method, "OK"])
            print(f"[{idx}/{total}] ✓ {f}  ->  {target}   ({method})")
        else:
            nonemp = [w for w in workers if w != EMP]
            if nonemp:
                plan.append([f, "|".join(nonemp), "", "", "", method, "CCCD_KHÔNG_TRONG_DS"])
                print(f"[{idx}/{total}] ? {f}  -> đọc được {nonemp} nhưng KHÔNG có trong DS")
            else:
                plan.append([f, "", "", "", "", method, "KHÔNG_ĐỌC_ĐƯỢC"])
                print(f"[{idx}/{total}] ? {f}  -> KHÔNG đọc được CCCD")

    # ---- Phát hiện trùng CCCD (nhiều file cùng 1 người)
    tgt_counter = Counter(p[4] for p in plan if p[6] == "OK" and p[4])
    dup_targets = {t for t, n in tgt_counter.items() if n > 1}

    # ---- Pha 2: copy sang output_PLHĐ (chống ghi đè)
    results = []
    n_ok = n_dup = n_prob = 0
    for src, cccd, stt, name, target, method, status in plan:
        if status != "OK":
            results.append([status, src, cccd, stt, name, "", method])
            n_prob += 1
            continue
        is_dup = target in dup_targets
        final = target
        dst = os.path.join(OUTDIR, final)
        if os.path.exists(dst):
            stem, ext = os.path.splitext(target)
            i = 2
            while os.path.exists(os.path.join(OUTDIR, f"{stem}_({i}){ext}")):
                i += 1
            final = f"{stem}_({i}){ext}"
            dst = os.path.join(OUTDIR, final)
        try:
            shutil.copy2(os.path.join(FOLDER, src), dst)
            st = "TRÙNG_CCCD" if is_dup else "OK"
            if is_dup:
                n_dup += 1
            else:
                n_ok += 1
            results.append([st, src, cccd, stt, name, final, method])
        except Exception as e:
            results.append(["LỖI_COPY", src, cccd, stt, name, final, str(e)[:60]])
            n_prob += 1

    # ---- Ghi CSV
    results.sort(key=lambda r: (r[0] != "OK", r[0]))
    with open(OUTCSV, "w", newline="", encoding="utf-8-sig") as fp:
        w = csv.writer(fp)
        w.writerow(["Trạng thái", "Tên file gốc", "CCCD (đọc từ PDF)",
                    "STT", "Họ Tên", "Tên file mới", "Ghi chú/Phương pháp"])
        w.writerows(results)

    # ---- Tổng hợp ra terminal
    print("\n================ TỔNG HỢP ================")
    print(f"Tổng file            : {total}")
    print(f"Đổi tên OK           : {n_ok}")
    print(f"Trùng CCCD (>=2 file): {n_dup}")
    print(f"Cần kiểm tra tay     : {n_prob}")
    if dup_targets:
        print("\nCác CCCD bị TRÙNG (nhiều file cùng 1 người):")
        for t in sorted(dup_targets):
            srcs = [p[0] for p in plan if p[4] == t]
            print("  -", t, "<=", srcs)
    print(f"\nFile đã đổi tên nằm ở: {OUTDIR}")
    print(f"CSV tổng hợp        : {OUTCSV}")


if __name__ == "__main__":
    main()
