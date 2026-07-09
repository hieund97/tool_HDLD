#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ghép ảnh CCCD (2 mặt: trước + sau) thành 1 file PDF, đặt tên theo STT + tên trong Excel.

Quy trình:
  1. Đọc Excel (sheet 'ĐL02 Full'): STT (cột A), Tên CVKD (cột C), số CCCD (cột D).
  2. Gom ảnh trong thư mục CCCD thành cặp mặt trước / mặt sau theo tên file.
  3. OCR số CCCD trên ảnh, đối chiếu với danh sách CCCD trong Excel để lấy đúng STT + tên.
     - CHỈ khớp theo số CCCD. Nếu không đọc được CCCD -> ghi LỖI (không fallback theo tên).
  4. Ghép 2 mặt (trước ở trên, sau ở dưới) thành 1 file PDF 1 trang,
     đặt tên: "{STT}B - {Tên}.pdf"  (ví dụ: "51B - Nguyễn Văn A.pdf").
  5. Người thiếu mặt / không khớp -> bỏ qua và ghi vào báo cáo.
  6. Xuất ket_qua_tong_hop.csv + in tóm tắt Thành công / Lỗi ra màn hình.

Chạy lại được nhiều lần (resume): file PDF đã tạo sẽ được bỏ qua.

Cách dùng:
    python3 cccd_to_pdf.py

Phụ thuộc (đã có sẵn trên máy): pillow, pytesseract, openpyxl  + tesseract (ngôn ngữ 'vie').
"""

import os
import re
import csv
import sys
import unicodedata
from collections import defaultdict

from PIL import Image, ImageOps
import pytesseract
import openpyxl

# ----------------------------------------------------------------------------
# CẤU HÌNH — chỉnh đường dẫn ở đây nếu cần
# ----------------------------------------------------------------------------
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
# Thư mục ngoài "CCCD" đã xử lý xong ở lần chạy trước.
# Lần này CHỈ quét thư mục con "CCCD/CCCD_2".
# Dedup theo SỐ CCCD: OCR ra CCCD -> tra Excel ra {STT}+{Tên} -> nếu PDF đích đã có
# (đã tạo từ thư mục ngoài) thì bỏ qua "trùng số CCCD".
IMG_DIR    = os.path.join(BASE_DIR, "CCCD", "CCCD_2")
EXCEL_PATH = os.path.join(
    BASE_DIR,
    "ĐL02_Danh sách sale và dự án tham gia bán 11.6.26 Full toàn bộ dữ liệu gốc.xlsx",
)
SHEET_NAME = "ĐL02 Full"
HEADER_ROW = 9          # dòng tiêu đề; dữ liệu bắt đầu từ dòng 10
COL_STT    = 0          # cột A
COL_NAME   = 2          # cột C  (Tên CVKD)
COL_CCCD   = 3          # cột D  (Căn cước CD)

OUT_DIR    = os.path.join(BASE_DIR, "output_pdf")
CSV_PATH   = os.path.join(BASE_DIR, "ket_qua_tong_hop.csv")

IMG_EXTS   = (".jpg", ".jpeg", ".png")
OCR_LONG_EDGE = 2400    # phóng to cạnh dài tới ~2400px trước khi OCR
PDF_TARGET_W  = 1654    # bề rộng đích khi ghép (≈ A4 @200dpi)
PDF_MARGIN    = 40
PDF_GAP       = 40


# ----------------------------------------------------------------------------
# Tiện ích chuẩn hoá chuỗi (tiếng Việt)
# ----------------------------------------------------------------------------
def strip_accents(s: str) -> str:
    """Bỏ dấu tiếng Việt + đổi đ/Đ -> d/D (để so khớp gần đúng)."""
    s = s.replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s


def norm(s) -> str:
    """NFC + thường + gộp khoảng trắng."""
    s = unicodedata.normalize("NFC", str(s)).strip().lower()
    return re.sub(r"\s+", " ", s)


def fold(s) -> str:
    """Bỏ dấu hoàn toàn (để khớp gần đúng khi tên file sai dấu)."""
    return re.sub(r"\s+", " ", strip_accents(norm(s))).strip()


def only_digits(s: str) -> str:
    return re.sub(r"\D", "", s)


def edit_distance_le1(a: str, b: str) -> bool:
    """True nếu khoảng cách sửa (Levenshtein) giữa a và b <= 1 (nhanh, không tính đầy đủ)."""
    if a == b:
        return True
    la, lb = len(a), len(b)
    if abs(la - lb) > 1:
        return False
    if la == lb:  # tối đa 1 vị trí khác
        return sum(1 for x, y in zip(a, b) if x != y) <= 1
    # lệch 1 ký tự: thử 1 lần chèn/xoá
    if la > lb:
        a, b = b, a
    i = j = 0
    skipped = False
    while i < len(a) and j < len(b):
        if a[i] == b[j]:
            i += 1
            j += 1
        elif skipped:
            return False
        else:
            skipped = True
            j += 1
    return True


def sanitize_filename(name: str) -> str:
    name = re.sub(r'[/\\:*?"<>|]+', " ", name)
    return re.sub(r"\s+", " ", name).strip()


# ----------------------------------------------------------------------------
# 1) Đọc Excel
# ----------------------------------------------------------------------------
def load_excel(path):
    """Trả về:
        cccd2row     : {cccd-12-số: (stt, ten)}
        name_exact   : {norm(ten): [(stt, ten, cccd), ...]}
        name_fold    : {fold(ten): [(stt, ten, cccd), ...]}
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]

    cccd2row = {}
    name_exact = defaultdict(list)
    name_fold = defaultdict(list)

    for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
        stt = row[COL_STT] if len(row) > COL_STT else None
        name = row[COL_NAME] if len(row) > COL_NAME else None
        cccd = row[COL_CCCD] if len(row) > COL_CCCD else None
        if not name:
            continue
        name = str(name).strip()
        cccd_digits = only_digits(str(cccd)) if cccd else ""
        rec = (stt, name, cccd_digits)
        if len(cccd_digits) == 12:
            cccd2row[cccd_digits] = (stt, name)
        name_exact[norm(name)].append(rec)
        name_fold[fold(name)].append(rec)

    wb.close()
    return cccd2row, name_exact, name_fold


# ----------------------------------------------------------------------------
# 2) Gom ảnh thành cặp trước / sau
# ----------------------------------------------------------------------------
def parse_name_index(filename: str):
    """'Nguyễn Văn A1.jpg' -> ('Nguyễn Văn A', 1);  'Lê Quốc Vương.jpg' -> ('Lê Quốc Vương', None)"""
    stem = os.path.splitext(filename)[0]
    m = re.search(r"([12])\s*$", stem)
    if m:
        return stem[: m.start()].rstrip(), int(m.group(1))
    return stem.strip(), None


def build_groups(img_dir):
    """Trả về list các nhóm: mỗi nhóm = {'display', 'base', 'items':[(idx, filename), ...]}"""
    files = [f for f in os.listdir(img_dir) if f.lower().endswith(IMG_EXTS)]
    files.sort()

    raw = defaultdict(list)  # norm(base) -> [(idx, filename, base_goc)]
    for f in files:
        base, idx = parse_name_index(f)
        raw[norm(base)].append((idx, f, base))

    groups = []
    singletons = {}  # norm_base -> nhóm 1 ảnh (để gộp các ca lệch dấu)

    for nbase, items in raw.items():
        disp = items[0][2]  # tên gốc đại diện
        g = {"display": disp, "base": nbase, "items": [(i, f) for (i, f, _) in items]}
        if len(g["items"]) == 1:
            singletons[nbase] = g
        else:
            groups.append(g)

    # Gộp các nhóm 1-ảnh có tên bỏ-dấu trùng nhau và mặt bù nhau
    # (cứu: mác/mắc, chinh/chính, ... khi 2 mặt ghi khác dấu)
    by_fold = defaultdict(list)
    for nbase, g in singletons.items():
        by_fold[fold(nbase)].append(g)

    used = set()
    for fkey, glist in by_fold.items():
        if len(glist) == 2:
            (i1, f1) = glist[0]["items"][0]
            (i2, f2) = glist[1]["items"][0]
            # chỉ gộp nếu không phải cả hai cùng một mặt rõ ràng (vd cùng "1")
            if not (i1 is not None and i2 is not None and i1 == i2):
                merged = {
                    "display": glist[0]["display"],
                    "base": glist[0]["base"],
                    "items": [(i1, f1), (i2, f2)],
                    "merged_from": [glist[0]["display"], glist[1]["display"]],
                }
                groups.append(merged)
                used.add(id(glist[0]))
                used.add(id(glist[1]))

    # Gộp gần đúng: các nhóm 1-ảnh còn lại, tên bỏ-dấu chỉ lệch <=1 ký tự
    # và mặt bù nhau (cứu: phạm/phan thị kim liên 1+2).
    remain = [g for g in singletons.values() if id(g) not in used]
    for i in range(len(remain)):
        ga = remain[i]
        if id(ga) in used:
            continue
        ia, fa = ga["items"][0]
        for j in range(i + 1, len(remain)):
            gb = remain[j]
            if id(gb) in used:
                continue
            ib, fb = gb["items"][0]
            same_explicit = ia is not None and ib is not None and ia == ib
            if same_explicit:
                continue
            if edit_distance_le1(fold(ga["base"]), fold(gb["base"])):
                groups.append({
                    "display": ga["display"],
                    "base": ga["base"],
                    "items": [(ia, fa), (ib, fb)],
                    "merged_from": [ga["display"], gb["display"]],
                })
                used.add(id(ga))
                used.add(id(gb))
                break

    # Các nhóm 1-ảnh thực sự lẻ -> giữ nguyên (sẽ bị bỏ qua vì thiếu mặt)
    for g in singletons.values():
        if id(g) not in used:
            groups.append(g)

    groups.sort(key=lambda g: g["display"].lower())
    return groups


def assign_front_back(items):
    """items=[(idx, filename), ...] -> (front_filename, back_filename) | (x, None)"""
    by = {1: [], 2: [], None: []}
    for idx, f in sorted(items, key=lambda x: x[1]):
        by[idx].append(f)

    front = by[1][0] if by[1] else None
    back = by[2][0] if by[2] else None

    leftovers = by[None][:] + by[1][1:] + by[2][1:]
    if front is None and leftovers:
        front = leftovers.pop(0)
    if back is None and leftovers:
        back = leftovers.pop(0)
    return front, back


# ----------------------------------------------------------------------------
# 3) OCR số CCCD
# ----------------------------------------------------------------------------
def _prep_variants(path):
    im = ImageOps.exif_transpose(Image.open(path)).convert("L")
    w, h = im.size
    if max(w, h) < OCR_LONG_EDGE:
        s = OCR_LONG_EDGE / max(w, h)
        im = im.resize((int(w * s), int(h * s)))
    return [im, ImageOps.autocontrast(im)]


def ocr_cccd_candidates(path):
    """Trả về tập các chuỗi 12 chữ số khả dĩ đọc được từ ảnh."""
    cands = set()
    try:
        variants = _prep_variants(path)
    except Exception as e:
        print(f"    [!] Không mở được ảnh {os.path.basename(path)}: {e}")
        return cands

    texts = []
    for v in variants:
        for psm in (6, 11, 4):
            cfg = f"--psm {psm} -c tessedit_char_whitelist=0123456789"
            try:
                texts.append(pytesseract.image_to_string(v, config=cfg))
            except Exception:
                pass
        try:
            texts.append(pytesseract.image_to_string(v, lang="vie"))
        except Exception:
            pass
        v.close()  # giải phóng RAM ngay sau khi OCR xong biến thể này

    for t in texts:
        for line in t.splitlines():
            d = only_digits(line)
            for i in range(0, len(d) - 11):
                cands.add(d[i : i + 12])
    return cands


# ----------------------------------------------------------------------------
# 4) Khớp nhóm với Excel
# ----------------------------------------------------------------------------
def match_group(front_path, back_path, cccd_set):
    """OCR mặt trước trước; nếu chưa khớp duy nhất thì OCR thêm mặt sau."""
    cands = ocr_cccd_candidates(front_path)
    hits = cands & cccd_set
    if len(hits) != 1 and back_path:
        cands |= ocr_cccd_candidates(back_path)
        hits = cands & cccd_set
    return hits


def name_fallback(base, name_exact, name_fold):
    rows = name_exact.get(norm(base))
    if not rows:
        rows = name_fold.get(fold(base))
    if rows and len(rows) == 1:
        return rows[0], None              # (stt, ten, cccd), no-error
    if rows and len(rows) > 1:
        return None, "tên khớp nhiều dòng Excel (nhập nhằng)"
    return None, "không đọc được CCCD & tên không có trong Excel"


# ----------------------------------------------------------------------------
# 5) Ghép PDF
# ----------------------------------------------------------------------------
def make_pdf(front_path, back_path, out_path):
    imgs = []
    for p in (front_path, back_path):
        im = ImageOps.exif_transpose(Image.open(p)).convert("RGB")
        w, h = im.size
        nh = int(h * PDF_TARGET_W / w)
        imgs.append(im.resize((PDF_TARGET_W, nh)))

    total_h = PDF_MARGIN * 2 + sum(i.size[1] for i in imgs) + PDF_GAP * (len(imgs) - 1)
    canvas = Image.new("RGB", (PDF_TARGET_W + PDF_MARGIN * 2, total_h), "white")
    y = PDF_MARGIN
    for i in imgs:
        canvas.paste(i, (PDF_MARGIN, y))
        y += i.size[1] + PDF_GAP
    canvas.save(out_path, "PDF", resolution=200.0)


# ----------------------------------------------------------------------------
# Chương trình chính
# ----------------------------------------------------------------------------
def main():
    # In tiến trình ra ngay, không bị dồn buffer khi nối qua `| tee`
    try:
        sys.stdout.reconfigure(line_buffering=True)
    except Exception:
        pass

    # kiểm tra tesseract
    try:
        pytesseract.get_tesseract_version()
    except Exception as e:
        print("LỖI: không tìm thấy 'tesseract'. Cài bằng: brew install tesseract tesseract-lang")
        print("Chi tiết:", e)
        sys.exit(1)

    if not os.path.isfile(EXCEL_PATH):
        print("LỖI: không thấy file Excel:", EXCEL_PATH)
        sys.exit(1)
    if not os.path.isdir(IMG_DIR):
        print("LỖI: không thấy thư mục ảnh:", IMG_DIR)
        sys.exit(1)

    os.makedirs(OUT_DIR, exist_ok=True)

    print("Đang đọc Excel...")
    cccd2row, name_exact, name_fold = load_excel(EXCEL_PATH)
    cccd_set = set(cccd2row)
    print(f"  -> {len(cccd_set)} số CCCD hợp lệ (12 số) trong Excel.")

    groups = build_groups(IMG_DIR)
    print(f"Tìm thấy {len(groups)} nhóm ảnh (người).\n")

    success = []   # dict kết quả thành công
    failed = []    # dict kết quả lỗi/bỏ qua

    for n, g in enumerate(groups, 1):
        disp = g["display"]
        items = g["items"]
        front, back = assign_front_back(items)
        prefix = f"[{n}/{len(groups)}] {disp}"

        # Thiếu mặt -> bỏ qua
        if front is None or back is None:
            reason = f"thiếu mặt (chỉ có {len(items)} ảnh)"
            print(f"{prefix}: BỎ QUA - {reason}")
            failed.append({
                "nhom": disp, "trang_thai": "BỎ QUA", "phuong_thuc": "",
                "stt": "", "ten": "", "cccd": "",
                "file_pdf": "", "anh_truoc": front or "", "anh_sau": back or "",
                "ghi_chu": reason,
            })
            continue

        front_path = os.path.join(IMG_DIR, front)
        back_path = os.path.join(IMG_DIR, back)

        # Khớp theo CCCD (OCR)
        hits = match_group(front_path, back_path, cccd_set)
        method = ""
        stt = ten = cccd = None

        if len(hits) == 1:
            cccd = next(iter(hits))
            stt, ten = cccd2row[cccd]
            method = "CCCD"
        else:
            # CHỈ khớp theo CCCD. Không đọc được CCCD -> LỖI (không fallback theo tên).
            if len(hits) > 1:
                detail = f"OCR ra nhiều số CCCD khớp ({len(hits)}) - không xác định được"
            else:
                detail = "không đọc được số CCCD trên ảnh"
            print(f"{prefix}: LỖI - {detail}")
            failed.append({
                "nhom": disp, "trang_thai": "LỖI", "phuong_thuc": "",
                "stt": "", "ten": "", "cccd": "",
                "file_pdf": "", "anh_truoc": front, "anh_sau": back,
                "ghi_chu": detail,
            })
            continue

        # Tạo PDF (tên theo STT + tên Excel)
        out_name = sanitize_filename(f"{stt}B - {ten}") + ".pdf"
        out_path = os.path.join(OUT_DIR, out_name)

        note = ""
        if os.path.exists(out_path):
            note = "trùng số CCCD - đã có PDF, bỏ qua"
        else:
            try:
                make_pdf(front_path, back_path, out_path)
                note = "tạo mới"
            except Exception as e:
                print(f"{prefix}: LỖI tạo PDF - {e}")
                failed.append({
                    "nhom": disp, "trang_thai": "LỖI", "phuong_thuc": method,
                    "stt": stt, "ten": ten, "cccd": cccd,
                    "file_pdf": out_name, "anh_truoc": front, "anh_sau": back,
                    "ghi_chu": f"lỗi tạo PDF: {e}",
                })
                continue

        print(f"{prefix}: OK [{method}] -> {out_name}  ({note})")
        success.append({
            "nhom": disp, "trang_thai": "THÀNH CÔNG", "phuong_thuc": method,
            "stt": stt, "ten": ten, "cccd": cccd,
            "file_pdf": out_name, "anh_truoc": front, "anh_sau": back,
            "ghi_chu": note,
        })

    # --- Ghi CSV ---
    cols = ["nhom", "trang_thai", "phuong_thuc", "stt", "ten", "cccd",
            "file_pdf", "anh_truoc", "anh_sau", "ghi_chu"]
    with open(CSV_PATH, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=cols)
        w.writeheader()
        for r in success + failed:
            w.writerow(r)

    # --- In tóm tắt ---
    print("\n" + "=" * 70)
    print(f"=== THÀNH CÔNG ({len(success)}) ===")
    for r in success:
        print(f"  {r['file_pdf']:42s} | CCCD {r['cccd']} | khớp: {r['phuong_thuc']:4s}"
              f" | nguồn: {r['anh_truoc']} + {r['anh_sau']}")

    print(f"\n=== LỖI / BỎ QUA ({len(failed)}) ===")
    for r in failed:
        print(f"  [{r['trang_thai']:7s}] {r['nhom']:35s} | {r['ghi_chu']}")

    print("\n" + "=" * 70)
    print(f"Tổng nhóm: {len(groups)}  |  Thành công: {len(success)}  |  Lỗi/Bỏ qua: {len(failed)}")
    print(f"PDF lưu tại : {OUT_DIR}")
    print(f"Báo cáo CSV : {CSV_PATH}")


if __name__ == "__main__":
    main()
