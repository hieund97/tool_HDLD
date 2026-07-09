# -*- coding: utf-8 -*-
"""
Đổi tên ảnh trong folder editCCCD theo cấu trúc: STT_CCCD_HoVaTen.ext
- STT, Họ Tên, CCCD lấy từ Excel sheet ĐL02 (cột A, C, E).
- Tách CCCD từ tên file: lấy mọi chuỗi số, đối chiếu với tập CCCD (12 số) trong Excel.
- Chỉ xử lý file ảnh .jpg/.jpeg/.png ở thư mục gốc editCCCD. Thư mục con -> ghi lỗi.
- Xuất 1 file CSV kết quả gồm cả thành công và lỗi.
"""
import os
import re
import csv
import unicodedata
import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
FOLDER = os.path.join(BASE, "editCCCD")
EXCEL = os.path.join(FOLDER, "BM_ĐL06_Danh sách sale bán độc quyền Quỹ Vin của ĐL_V2 - MH Full CN.xlsx")
SHEET = "ĐL02"
CSV_OUT = os.path.join(BASE, "ket_qua_doi_ten_cccd.csv")

IMG_EXT = {".jpg", ".jpeg", ".png"}

def load_excel():
    """Trả về dict: cccd(str 12 số) -> (stt, ho_ten)."""
    wb = openpyxl.load_workbook(EXCEL, read_only=True, data_only=True)
    ws = wb[SHEET]
    mapping = {}
    for row in ws.iter_rows(min_row=9, max_col=5, values_only=True):
        stt, _, name, _, cccd = row
        if cccd is None:
            continue
        cccd = str(cccd).strip()
        if not cccd:
            continue
        stt = "" if stt is None else str(stt).strip()
        name = "" if name is None else str(name).strip()
        mapping[cccd] = (stt, name)
    wb.close()
    return mapping

def remove_diacritics(s):
    """Bỏ dấu tiếng Việt, xử lý đ/Đ riêng."""
    s = s.replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return unicodedata.normalize("NFC", s)

def sanitize(name):
    """Họ tên -> viết liền, không dấu, mỗi từ viết hoa chữ cái đầu.
    Ví dụ: 'Phạm Phương Thảo' -> 'PhamPhuongThao'."""
    name = remove_diacritics(name)
    # tách theo khoảng trắng và mọi ký tự không phải chữ/số, viết hoa đầu mỗi từ
    words = re.findall(r"[A-Za-z0-9]+", name)
    return "".join(w[:1].upper() + w[1:] for w in words)

def find_cccd(filename, cccd_set):
    """Tìm mọi chuỗi số trong tên file khớp với tập CCCD Excel."""
    digits = re.findall(r"\d+", filename)
    matches = [d for d in digits if d in cccd_set]
    # khử trùng lặp, giữ thứ tự
    seen = set()
    uniq = [d for d in matches if not (d in seen or seen.add(d))]
    return uniq

def unique_target(folder, target):
    """Tránh ghi đè: thêm hậu tố _2, _3... nếu trùng tên."""
    if not os.path.exists(os.path.join(folder, target)):
        return target
    stem, ext = os.path.splitext(target)
    i = 2
    while os.path.exists(os.path.join(folder, f"{stem}_{i}{ext}")):
        i += 1
    return f"{stem}_{i}{ext}"

def main():
    mapping = load_excel()
    cccd_set = set(mapping.keys())
    print(f"Đã nạp {len(cccd_set)} CCCD từ Excel.")

    rows = []  # cho CSV
    n_ok = n_err = 0

    for entry in sorted(os.listdir(FOLDER)):
        path = os.path.join(FOLDER, entry)

        # Bỏ qua thư mục con -> ghi lỗi
        if os.path.isdir(path):
            rows.append(["LỖI", entry, "", "", "", "", "Là thư mục con - bỏ qua"])
            n_err += 1
            continue

        ext = os.path.splitext(entry)[1].lower()
        if ext not in IMG_EXT:
            continue  # không phải ảnh jpg/png -> không liệt kê

        matches = find_cccd(entry, cccd_set)
        if not matches:
            rows.append(["LỖI", entry, "", "", "", "", "Không tìm thấy CCCD khớp trong Excel"])
            n_err += 1
            continue
        if len(matches) > 1:
            rows.append(["LỖI", entry, ",".join(matches), "", "", "",
                         "Nhiều CCCD khớp - cần kiểm tra tay"])
            n_err += 1
            continue

        cccd = matches[0]
        stt, ho_ten = mapping[cccd]
        ho_ten_clean = sanitize(ho_ten)
        # giữ đuôi gốc (kể cả .JPG -> chuẩn hóa về chữ thường)
        new_name = f"{stt}_{cccd}_{ho_ten_clean}{ext}"
        new_name = unique_target(FOLDER, new_name)

        if new_name == entry:
            rows.append(["THÀNH CÔNG", entry, cccd, stt, ho_ten, new_name, "Đã đúng tên, bỏ qua"])
            n_ok += 1
            continue

        try:
            os.rename(path, os.path.join(FOLDER, new_name))
            rows.append(["THÀNH CÔNG", entry, cccd, stt, ho_ten, new_name, ""])
            n_ok += 1
        except Exception as e:
            rows.append(["LỖI", entry, cccd, stt, ho_ten, new_name, f"Lỗi đổi tên: {e}"])
            n_err += 1

    # sắp xếp: lỗi xuống cuối cho dễ nhìn
    rows.sort(key=lambda r: (r[0] != "THÀNH CÔNG",))
    with open(CSV_OUT, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Trạng thái", "Tên file gốc", "CCCD", "STT", "Họ Tên (Excel)",
                    "Tên file mới", "Ghi chú"])
        w.writerows(rows)

    print(f"Thành công: {n_ok} | Lỗi/bỏ qua: {n_err}")
    print(f"CSV: {CSV_OUT}")

if __name__ == "__main__":
    main()
