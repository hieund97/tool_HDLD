# -*- coding: utf-8 -*-
"""
Đổi tên PDF hợp đồng trong folder 'DỮ LIỆU CẦN ĐỔI' theo cú pháp:
    {STT hồ sơ}_{Số CCCD}_{Họ và tên}_{HĐLĐ|HĐTV|PHỤ LỤC SỬA ĐỔI THÔNG TIN}.pdf

- Loại văn bản: OCR tiêu đề viết HOA (HỢP ĐỒNG LAO ĐỘNG / HỢP ĐỒNG THỬ VIỆC / PHỤ LỤC).
- Số CCCD: OCR nội dung, bỏ CCCD đại diện công ty (EMP), đối chiếu Excel.
- STT hồ sơ / Họ và tên / Căn cước CD: tìm theo tiêu đề cột trong Excel đối chiếu.
- File có nhiều tiêu đề (nghi chứa nhiều văn bản) hoặc không khớp -> KHÔNG đổi, ghi CSV.
- Kết quả OCR được cache, chạy lại --apply không phải OCR lại.

Chạy:
    python3 doi_ten_hop_dong_vin.py            # dry-run
    python3 doi_ten_hop_dong_vin.py --apply    # đổi tên thật
    python3 doi_ten_hop_dong_vin.py --undo     # hoàn tác theo CSV
    python3 doi_ten_hop_dong_vin.py --no-cache # bỏ cache, OCR lại
    python3 doi_ten_hop_dong_vin.py --folder "DỮ LIỆU CẦN ĐỔI TÊN 2" [--apply|--undo]
    python3 doi_ten_hop_dong_vin.py --excel "<đường dẫn xlsx khác>"
"""
import os, re, sys, csv, glob, json, hashlib, tempfile, subprocess, unicodedata
from collections import Counter
from concurrent.futures import ThreadPoolExecutor

# 1 luồng / lần gọi tesseract: tránh tranh CPU khi đã chạy song song nhiều file
os.environ.setdefault("OMP_THREAD_LIMIT", "1")

import openpyxl
import pytesseract
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from doi_ten_pdf_tu_gsheet import oriented, nums12, value_crops, digitocr, fuzzy  # noqa: E402

def nfc(s):
    return unicodedata.normalize("NFC", s)


def strip_acc(s):
    """Bỏ dấu, GIỮ hoa/thường (để phân biệt tiêu đề viết HOA với câu thường)."""
    s = unicodedata.normalize("NFD", s).replace("Đ", "D").replace("đ", "d")
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


BASE   = os.path.dirname(os.path.abspath(__file__))
def _arg(name, default):
    for i, a in enumerate(sys.argv):
        if a == name and i + 1 < len(sys.argv):
            return sys.argv[i + 1]
        if a.startswith(name + "="):
            return a.split("=", 1)[1]
    return default


FOLDER = os.path.join(BASE, _arg("--folder", "DỮ LIỆU CẦN ĐỔI"))
EXCEL  = os.path.join(BASE, _arg("--excel", "DỮ LIỆU CẦN ĐỔI/Import HDLD.xlsx"))
_TAG   = re.sub(r"\W+", "_", strip_acc(os.path.basename(FOLDER.rstrip("/")))).strip("_")
OUTCSV = os.path.join(BASE, f"ket_qua_doi_ten_{_TAG}.csv")
CACHE  = os.path.join(tempfile.gettempdir(), f"ocr_cache_{_TAG}.json")

EMP     = "040191008474"
DPI     = 300
DPI_HI  = 400
WORKERS = max(1, (os.cpu_count() or 4) // 2)

LBL_HDLD = "HĐLĐ"
LBL_HDTV = "HĐTV"
LBL_PL   = "PHỤ LỤC SỬA ĐỔI THÔNG TIN"


# ---------------------------------------------------------------- Excel
def load_excel():
    """Tìm sheet có cột 'STT hồ sơ' + 'Họ và tên' + 'Căn cước CD' (theo tiêu đề, không cố định vị trí)."""
    wb = openpyxl.load_workbook(EXCEL, read_only=True, data_only=True)
    ws = cols = None
    for sh in wb.worksheets:
        head = [nfc(str(x or "")).strip().lower() for x in next(sh.iter_rows(max_row=1, values_only=True), ())]
        need = ("stt hồ sơ", "họ và tên", "căn cước cd")
        if all(h in head for h in need):
            ws, cols = sh, [head.index(h) for h in need]
            break
    if ws is None:
        sys.exit(f"Không thấy sheet có cột 'STT hồ sơ', 'Họ và tên', 'Căn cước CD' trong {EXCEL}")
    m, skipped = {}, 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) <= max(cols):
            continue
        stt, name, raw = (row[c] for c in cols)
        cccd = re.sub(r"\D", "", str(raw or ""))
        stt = str(stt).strip() if stt is not None else ""
        name = re.sub(r"\s+", " ", nfc(str(name or ""))).strip()
        if len(cccd) != 12 or not stt or not name:
            if stt or name or raw:
                skipped += 1
            continue
        if cccd in m and m[cccd] != (stt, name):
            print(f"  ! CCCD {cccd} xuất hiện nhiều lần trong Excel: {m[cccd]} / {(stt, name)}")
        m[cccd] = (stt, name)
    wb.close()
    print(f"Excel: {len(m)} CCCD từ sheet {ws.title!r} (bỏ qua {skipped} dòng thiếu dữ liệu)")
    return m


# ---------------------------------------------------------------- OCR
def render(pdf, tmp, dpi):
    prefix = os.path.join(tmp, f"p{dpi}")
    subprocess.run(["pdftoppm", "-r", str(dpi), "-jpeg", pdf, prefix],
                   check=True, stderr=subprocess.DEVNULL)
    return sorted(glob.glob(prefix + "-*.jpg"))


def classify_page(text):
    """Loại văn bản nếu trang có tiêu đề viết HOA, ngược lại None."""
    t = re.sub(r"\s+", " ", strip_acc(text))
    if re.search(r"PHU\s*LUC", t):
        return LBL_PL
    if re.search(r"HO+P\s*DO+NG\s*THU\s*VIEC", t):   # chịu OCR thừa chữ O ("HỌỢP")
        return LBL_HDTV
    if re.search(r"HO+P\s*DO+NG\s*LAO\s*DONG", t):
        return LBL_HDLD
    return None


def kind_from_number(text):
    """Dự phòng khi không đọc được tiêu đề: lấy từ số hợp đồng 'Số: MH..../HĐTV'."""
    t = strip_acc(text).upper()
    if re.search(r"MH\d+.{0,12}/\s*HDTV", t):
        return LBL_HDTV
    if re.search(r"MH\d+.{0,12}/\s*HDLD", t):
        return LBL_HDLD
    return None


def ocr_file(pdf):
    """OCR toàn bộ trang + vùng số CCCD. Trả về dict lưu được vào cache."""
    with tempfile.TemporaryDirectory() as tmp:
        pages = []
        for j in render(pdf, tmp, DPI):
            im = oriented(Image.open(j).convert("RGB"))
            pages.append(pytesseract.image_to_string(im, lang="vie"))
        crops = []
        for k, j in enumerate(render(pdf, tmp, DPI_HI)):
            im = oriented(Image.open(j).convert("RGB"))
            for x0, y0, x1, y1 in value_crops(im):
                if x1 - x0 < 20 or y1 - y0 < 5:   # vùng cắt lỗi (vd 'Ngày' đứng trước ':')
                    continue
                crops += [(k, c) for c in digitocr(im.crop((x0, y0, x1, y1)))]
    return {"pages": pages, "crops": crops}


def file_key(path):
    st = os.stat(path)
    return hashlib.md5(f"{st.st_size}-{int(st.st_mtime)}".encode()).hexdigest()


# ---------------------------------------------------------------- analysis
def analyse(ocr, ccset):
    """-> (loại, cccd|None, ứng viên CCCD, ghi chú)"""
    pages = ocr["pages"]
    titles = [(i + 1, classify_page(p)) for i, p in enumerate(pages)]
    titles = [(i, t) for i, t in titles if t]
    kinds = list(dict.fromkeys(t for _, t in titles))

    per_page = []
    for i, p in enumerate(pages):
        per_page += [(i, c) for c in nums12(p)]
    per_page += [tuple(x) for x in ocr["crops"]]
    cands = [c for c in dict.fromkeys(c for _, c in per_page) if c != EMP]
    inx = [c for c in cands if c in ccset]

    note = []
    if len(titles) > 1:
        note.append("tiêu đề ở trang " + ", ".join(f"{i}:{t}" for i, t in titles))
    kind = kinds[0] if len(kinds) == 1 else None
    if len(kinds) > 1:
        note.append("nhiều loại văn bản")
    if not kinds and pages and kind_from_number(pages[0]):
        kind = kind_from_number(pages[0])
        note.append("loại lấy theo số hợp đồng (không đọc được tiêu đề)")
    elif not kinds:
        note.append("không thấy tiêu đề")
    return kind, (inx[0] if len(inx) == 1 else None), cands, inx, len(titles), note


# ---------------------------------------------------------------- undo
def undo():
    if not os.path.exists(OUTCSV):
        sys.exit(f"Không có file log {OUTCSV}")
    n = 0
    with open(OUTCSV, encoding="utf-8-sig") as fp:
        for r in csv.DictReader(fp):
            if r["Trạng thái"] != "ĐÃ ĐỔI TÊN":
                continue
            src = os.path.join(FOLDER, r["Tên file mới"])
            dst = os.path.join(FOLDER, r["Tên file gốc"])
            if os.path.exists(src) and not os.path.exists(dst):
                os.rename(src, dst); n += 1
    print(f"Đã hoàn tác {n} file.")


# ---------------------------------------------------------------- main
def main():
    if "--undo" in sys.argv:
        return undo()
    apply_ = "--apply" in sys.argv

    mapping = load_excel()
    ccset = set(mapping)
    cache = {}
    if "--no-cache" not in sys.argv and os.path.exists(CACHE):
        with open(CACHE, encoding="utf-8") as fp:
            cache = json.load(fp)

    files = sorted(nfc(f) for f in os.listdir(FOLDER) if f.lower().endswith(".pdf"))
    total = len(files)
    print(f"Folder: {FOLDER} | {total} PDF | {WORKERS} luồng")
    print("CHẾ ĐỘ: " + ("ĐỔI TÊN THẬT" if apply_ else "DRY-RUN (thêm --apply để đổi thật)") + "\n")

    todo = [f for f in files if cache.get(f, {}).get("key") != file_key(os.path.join(FOLDER, f))]

    def job(f):
        try:
            return f, ocr_file(os.path.join(FOLDER, f)), None
        except Exception as e:
            return f, None, str(e)[:100]

    errors = {}
    if todo:
        print(f"OCR {len(todo)} file (cache sẵn {total - len(todo)}) ...")
        with ThreadPoolExecutor(max_workers=WORKERS) as ex:
            for i, (f, res, err) in enumerate(ex.map(job, todo), 1):
                if err:
                    errors[f] = err
                    print(f"  [{i}/{len(todo)}] ✗ {f}: {err}")
                    continue
                cache[f] = {"key": file_key(os.path.join(FOLDER, f)), **res}
                print(f"  [{i}/{len(todo)}] {f}")
                if i % 10 == 0:
                    with open(CACHE, "w", encoding="utf-8") as fp:
                        json.dump(cache, fp, ensure_ascii=False)
        with open(CACHE, "w", encoding="utf-8") as fp:
            json.dump(cache, fp, ensure_ascii=False)

    plan = []  # [status, src, cccd, target, kind, note]
    for f in files:
        if f in errors:
            plan.append(["LỖI_OCR", f, "", "", "", errors[f]]); continue
        kind, cccd, cands, inx, ntitle, note = analyse(cache[f], ccset)
        n_pages = len(cache[f]["pages"])
        note.insert(0, f"{n_pages} trang")
        if ntitle > 1 or len(inx) > 1:
            if len(inx) > 1:
                note.append("nhiều CCCD trong DS: " + ", ".join(inx))
            plan.append(["NGHI_NHIỀU_VĂN_BẢN", f, "|".join(inx or cands), "", kind or "", "; ".join(note)])
        elif not cccd:
            fz = fuzzy(cands, ccset)
            note.append("OCR: " + (", ".join(cands) if cands else "không đọc được số"))
            if fz:
                note.append("gợi ý lệch 1 số: " + ", ".join(fz))
            st = "CCCD_KHÔNG_TRONG_DS" if cands else "KHÔNG_ĐỌC_ĐƯỢC_CCCD"
            plan.append([st, f, "|".join(cands), "", kind or "", "; ".join(note)])
        elif not kind:
            stt, name = mapping[cccd]
            note.append(f"người: {stt}_{cccd}_{name}")
            plan.append(["KHÔNG_RÕ_LOẠI", f, cccd, "", "", "; ".join(note)])
        else:
            stt, name = mapping[cccd]
            plan.append(["OK", f, cccd, f"{stt}_{cccd}_{name}_{kind}.pdf", kind, "; ".join(note)])

    dups = {t for t, n in Counter(p[3] for p in plan if p[0] == "OK").items() if n > 1}

    rows = []
    for st, src, cccd, target, kind, note in plan:
        if st != "OK":
            rows.append([st, src, cccd, "", kind, note]); continue
        if src == target:
            rows.append(["ĐÃ ĐÚNG TÊN", src, cccd, target, kind, note]); continue
        if target in dups:
            note = (note + "; " if note else "") + "nhiều file cùng tên đích - đã thêm hậu tố, kiểm tra tay"
        stem, ext = os.path.splitext(target)
        final, i = target, 2
        taken = {r[3] for r in rows}
        while os.path.exists(os.path.join(FOLDER, final)) or final in taken:
            final = f"{stem}_{i}{ext}"; i += 1
        if not apply_:
            rows.append(["SẼ ĐỔI TÊN", src, cccd, final, kind, note]); continue
        try:
            os.rename(os.path.join(FOLDER, src), os.path.join(FOLDER, final))
            rows.append(["ĐÃ ĐỔI TÊN", src, cccd, final, kind, note])
        except Exception as e:
            rows.append(["LỖI_ĐỔI_TÊN", src, cccd, final, kind, str(e)[:100]])

    rows.sort(key=lambda r: (r[0] not in ("ĐÃ ĐỔI TÊN", "SẼ ĐỔI TÊN"), r[0], r[1]))
    with open(OUTCSV, "w", newline="", encoding="utf-8-sig") as fp:
        w = csv.writer(fp)
        w.writerow(["Trạng thái", "Tên file gốc", "CCCD", "Tên file mới", "Loại", "Ghi chú"])
        w.writerows(rows)

    c = Counter(r[0] for r in rows)
    print("================ TỔNG HỢP ================")
    for k, v in c.most_common():
        print(f"  {k:<22}: {v}")
    print("Theo loại (OK): ", dict(Counter(r[4] for r in rows if r[3])))
    for r in rows:
        if r[0] not in ("ĐÃ ĐỔI TÊN", "SẼ ĐỔI TÊN", "ĐÃ ĐÚNG TÊN") or "hậu tố" in r[5]:
            print(f"  - [{r[0]}] {r[1]} | {r[2]} | {r[5]}")
    print(f"\nCSV: {OUTCSV}")
    if apply_:
        print("Hoàn tác: python3 doi_ten_hop_dong_vin.py --undo")


if __name__ == "__main__":
    main()
