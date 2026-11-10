# -*- coding: utf-8 -*-
"""
Đổi tên PDF trong folder 'ilovepdf_split' theo cột AJ ("Tên đổi") của Google Sheet.

- Số CCCD lấy từ NỘI DUNG PDF (OCR dòng "Số CMND/CCCD" của Người lao động).
- Đối chiếu cột F ("Căn cước CD") trong tab 'Chi tiết' của Google Sheet để lấy
  cột AJ ("Tên đổi") -> tên file mới.
- File được đổi tên TẠI CHỖ. CSV log ghi đủ tên cũ/tên mới để hoàn tác.

Chạy:
    python3 doi_ten_pdf_tu_gsheet.py              # dry-run, chỉ in kế hoạch
    python3 doi_ten_pdf_tu_gsheet.py --apply      # đổi tên thật
    python3 doi_ten_pdf_tu_gsheet.py --undo       # hoàn tác theo CSV log

Yêu cầu: pdftoppm (poppler), tesseract + gói tiếng Việt (vie),
         pip: pytesseract pillow openpyxl
"""
import os, re, sys, csv, glob, shutil, tempfile, subprocess, urllib.request
from collections import defaultdict, Counter
from concurrent.futures import ThreadPoolExecutor

import openpyxl
import pytesseract
from pytesseract import Output
from PIL import Image, ImageOps

BASE    = os.path.dirname(os.path.abspath(__file__))
FOLDER  = os.path.join(BASE, "ilovepdf_split")
OUTCSV  = os.path.join(BASE, "ket_qua_doi_ten_ilovepdf_split.csv")

SHEET_ID  = "16W73Ubw46ICKt3HA0-RDFaUPbuDmQRqGVkYM7yTBSv0"
SHEET_TAB = "Chi tiết"
COL_CCCD  = 5    # cột F
COL_NAME  = 35   # cột AJ
XLSX_CACHE = os.path.join(tempfile.gettempdir(), f"gsheet_{SHEET_ID}.xlsx")

EMP     = "040191008474"   # CCCD đại diện Công ty (Bên A) - luôn bỏ qua
DPI     = 300              # DPI cho lượt OCR toàn trang
DPI_HI  = 400              # DPI cho lượt OCR cắt vùng số
WORKERS = max(1, (os.cpu_count() or 4) // 2)


# ---------------------------------------------------------------- Google Sheet
def load_sheet(use_cache=False):
    """Tải sheet (public link) -> dict: cccd(12 số) -> tên đổi (cột AJ)."""
    if not (use_cache and os.path.exists(XLSX_CACHE)):
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
        print(f"Tải Google Sheet ...")
        urllib.request.urlretrieve(url, XLSX_CACHE)
    wb = openpyxl.load_workbook(XLSX_CACHE, read_only=True, data_only=True)
    if SHEET_TAB not in wb.sheetnames:
        sys.exit(f"Không tìm thấy tab {SHEET_TAB!r}. Các tab có: {wb.sheetnames}")
    ws = wb[SHEET_TAB]
    m, skipped = {}, 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        raw  = row[COL_CCCD] if len(row) > COL_CCCD else None
        name = row[COL_NAME] if len(row) > COL_NAME else None
        cccd = re.sub(r"\D", "", str(raw or ""))
        name = str(name).strip() if name else ""
        if len(cccd) != 12 or not name:
            if raw or name:
                skipped += 1
            continue
        m[cccd] = name
    wb.close()
    print(f"Đã nạp {len(m)} CCCD từ tab {SHEET_TAB!r} (bỏ qua {skipped} dòng thiếu/sai dữ liệu).")
    return m


# ---------------------------------------------------------------- OCR helpers
def render(pdf, tmpdir, dpi=DPI):
    prefix = os.path.join(tmpdir, f"p{dpi}")
    subprocess.run(["pdftoppm", "-r", str(dpi), "-jpeg", pdf, prefix],
                   check=True, stderr=subprocess.DEVNULL)
    return sorted(glob.glob(prefix + "-*.jpg"))


def oriented(im):
    """Tự xoay ảnh scan bị nghiêng/ngược bằng OSD của tesseract."""
    try:
        rot = int(re.search(r"Rotate: (\d+)", pytesseract.image_to_osd(im)).group(1))
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
    """Các vùng ảnh chứa PHẦN SỐ của dòng 'Số CMND/CCCD' (bỏ phần 'Ngày cấp')."""
    d = pytesseract.image_to_data(im, lang="vie", output_type=Output.DICT)
    n = len(d["text"]); W, H = im.size
    lines = defaultdict(list)
    for i in range(n):
        lines[(d["block_num"][i], d["par_num"][i], d["line_num"][i])].append(i)
    out = []
    for idxs in lines.values():
        idxs = sorted(idxs, key=lambda i: d["left"][i])
        if not re.search(r"CMND|CCCD", " ".join(d["text"][i] for i in idxs), re.I):
            continue
        y0 = min(d["top"][i] for i in idxs) - 6
        y1 = max(d["top"][i] + d["height"][i] for i in idxs) + 6
        x_start, x_end = None, W
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
        out.append((max(x_start - 5, 0), max(y0, 0), min(x_end + 5, W), min(y1, H)))
    return out


def digitocr(crop):
    im = crop.convert("L")
    im = ImageOps.autocontrast(im.resize((im.width * 4, im.height * 4)))
    res = []
    for psm in (7, 8, 6):
        s = pytesseract.image_to_string(
            im, config=f"--psm {psm} -c tessedit_char_whitelist=0123456789")
        res += re.findall(r"\d{12}", re.sub(r"[^\d]", " ", s))
    return res


def extract_worker(pdf, ccset):
    """Trả về (cccd_khớp_DS | None, các_số_12_đọc_được, phương_pháp)."""
    with tempfile.TemporaryDirectory() as tmp:
        pages_im, all12 = [], []
        for j in render(pdf, tmp, DPI):
            im = oriented(Image.open(j).convert("RGB"))
            pages_im.append(im)
            all12 += nums12(pytesseract.image_to_string(im, lang="vie"))
        workers = [c for c in dict.fromkeys(all12) if c != EMP]
        inx = [c for c in workers if c in ccset]
        if len(inx) == 1:
            return inx[0], workers, "full-ocr"

        # fallback: OCR riêng vùng số, ảnh nét hơn
        cands = []
        for j in render(pdf, tmp, DPI_HI):
            im = oriented(Image.open(j).convert("RGB"))
            for box in value_crops(im):
                cands += digitocr(im.crop(box))
        workers2 = [c for c in dict.fromkeys(workers + cands) if c != EMP]
        inx2 = [c for c in workers2 if c in ccset]
        if len(inx2) == 1:
            return inx2[0], workers2, "digit-crop"
        return None, workers2, "nhiều-kết-quả" if len(inx2) > 1 else "không-khớp"


def fuzzy(cands, ccset):
    """Gợi ý CCCD trong DS lệch đúng 1 ký tự so với số OCR được (chỉ để tham khảo)."""
    out = []
    for c in cands:
        if len(c) != 12:
            continue
        for k in ccset:
            if sum(a != b for a, b in zip(c, k)) == 1:
                out.append(f"{c}->{k}")
    return sorted(set(out))


# ---------------------------------------------------------------- undo
def undo():
    if not os.path.exists(OUTCSV):
        sys.exit(f"Không có file log {OUTCSV}")
    n = 0
    with open(OUTCSV, encoding="utf-8-sig") as fp:
        for r in csv.DictReader(fp):
            old, new = r["Tên file gốc"], r["Tên file mới"]
            if r["Trạng thái"] != "ĐÃ ĐỔI TÊN" or not new:
                continue
            src, dst = os.path.join(FOLDER, new), os.path.join(FOLDER, old)
            if os.path.exists(src) and not os.path.exists(dst):
                os.rename(src, dst); n += 1
    print(f"Đã hoàn tác {n} file.")


# ---------------------------------------------------------------- main
def main():
    apply_   = "--apply" in sys.argv
    use_cache = "--cache" in sys.argv
    if "--undo" in sys.argv:
        return undo()

    mapping = load_sheet(use_cache)
    ccset = set(mapping)
    files = sorted(f for f in os.listdir(FOLDER) if f.lower().endswith(".pdf"))
    total = len(files)
    print(f"Folder: {FOLDER}  |  {total} file PDF  |  {WORKERS} luồng OCR")
    print("CHẾ ĐỘ: " + ("ĐỔI TÊN THẬT\n" if apply_ else "DRY-RUN (chưa đụng vào file, thêm --apply để chạy thật)\n"))

    # ---- Pha 1: OCR song song
    def job(f):
        try:
            return f, extract_worker(os.path.join(FOLDER, f), ccset), None
        except Exception as e:
            return f, (None, [], "lỗi"), str(e)[:80]

    plan = []   # [src, cccd, target, method, status, note]
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        for i, (f, (c, cands, method), err) in enumerate(ex.map(job, files), 1):
            if err:
                plan.append([f, "", "", method, "LỖI_OCR", err])
                print(f"[{i}/{total}] ✗ {f} -> LỖI OCR: {err}")
            elif c:
                target = f"{mapping[c]}.pdf"
                plan.append([f, c, target, method, "OK", ""])
                print(f"[{i}/{total}] ✓ {f} -> {target}   ({method})")
            else:
                note = "OCR: " + (", ".join(cands) if cands else "không đọc được số nào")
                fz = fuzzy(cands, ccset)
                if fz:
                    note += " | gợi ý lệch 1 số: " + ", ".join(fz)
                st = "CCCD_KHÔNG_TRONG_DS" if cands else "KHÔNG_ĐỌC_ĐƯỢC"
                plan.append([f, "|".join(cands), "", method, st, note])
                print(f"[{i}/{total}] ? {f} -> {st} ({note})")

    # ---- Phát hiện nhiều file cùng trỏ về 1 tên
    cnt = Counter(p[2] for p in plan if p[4] == "OK")
    dups = {t for t, n in cnt.items() if n > 1}

    # ---- Pha 2: đổi tên tại chỗ
    rows, n_ok, n_skip, n_prob = [], 0, 0, 0
    for src, cccd, target, method, status, note in plan:
        if status != "OK":
            rows.append([status, src, cccd, "", method, note]); n_prob += 1
            continue
        if src == target:
            rows.append(["ĐÃ ĐÚNG TÊN", src, cccd, target, method, ""]); n_skip += 1
            continue
        final = target
        if target in dups:
            note = "Nhiều file cùng CCCD - đã thêm hậu tố, cần kiểm tra tay"
        stem, ext = os.path.splitext(target)
        i = 2
        while os.path.exists(os.path.join(FOLDER, final)):
            final = f"{stem}_{i}{ext}"; i += 1
        if not apply_:
            rows.append(["SẼ ĐỔI TÊN", src, cccd, final, method, note]); n_ok += 1
            continue
        try:
            os.rename(os.path.join(FOLDER, src), os.path.join(FOLDER, final))
            rows.append(["ĐÃ ĐỔI TÊN", src, cccd, final, method, note]); n_ok += 1
        except Exception as e:
            rows.append(["LỖI_ĐỔI_TÊN", src, cccd, final, method, str(e)[:80]]); n_prob += 1

    # ---- CSV
    rows.sort(key=lambda r: (r[0] not in ("ĐÃ ĐỔI TÊN", "SẼ ĐỔI TÊN"), r[1]))
    with open(OUTCSV, "w", newline="", encoding="utf-8-sig") as fp:
        w = csv.writer(fp)
        w.writerow(["Trạng thái", "Tên file gốc", "CCCD (đọc từ PDF)",
                    "Tên file mới", "Phương pháp", "Ghi chú"])
        w.writerows(rows)

    print("\n================ TỔNG HỢP ================")
    print(f"Tổng file        : {total}")
    print(f"{'Đã' if apply_ else 'Sẽ'} đổi tên      : {n_ok}")
    print(f"Đã đúng tên sẵn  : {n_skip}")
    print(f"Cần kiểm tra tay : {n_prob}")
    if dups:
        print("\nTên đích bị TRÙNG (nhiều file cùng 1 CCCD):")
        for t in sorted(dups):
            print("  -", t, "<=", [p[0] for p in plan if p[2] == t])
    print(f"\nCSV: {OUTCSV}")
    if apply_:
        print("Hoàn tác: python3 doi_ten_pdf_tu_gsheet.py --undo")


if __name__ == "__main__":
    main()
