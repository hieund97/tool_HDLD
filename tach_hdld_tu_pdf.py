#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tách file PDF "Danh sách HĐLĐ" (mỗi file scan chứa nhiều hợp đồng lao động)
thành từng file PDF riêng cho mỗi Người lao động.

Quy trình mỗi file SCAN*.pdf:
  1. OCR phần trên mỗi trang -> phát hiện TRANG ĐẦU của từng hợp đồng
     (dấu hiệu duy nhất: "SỐ: MH####/20##/HĐLĐ" + tiêu đề "HỢP ĐỒNG LAO ĐỘNG"
      + hai mục "NGƯỜI SỬ DỤNG LAO ĐỘNG" và "NGƯỜI LAO ĐỘNG").
  2. Mỗi hợp đồng = từ trang đầu của nó tới ngay trước trang đầu hợp đồng kế tiếp.
  3. Lấy số CCCD của NGƯỜI LAO ĐỘNG (mục 2, KHÔNG lấy CCCD người sử dụng lao động
     040191008474 vốn cố định trên mọi hợp đồng).
  4. Đối chiếu CCCD với cột D trong Excel (bỏ số 0 ở đầu để khớp vì OCR/định dạng
     thường mất 0 đầu) -> lấy STT (cột A) + Tên (cột C).
  5. Tách trang -> lưu PDF tên "{STT}A - {Tên}.pdf".
     - Không khớp được CCCD -> ghi LỖI, lưu vào thư mục con "_LOI" tên
       "LOI p{trang} {SỐ} - {tên OCR}.pdf" để kiểm tra thủ công.
  6. Ghi CSV + in chi tiết Thành công / Lỗi.

Phụ thuộc: pdf2image/poppler (pdftoppm), pytesseract+tesseract(vie), pypdf, openpyxl, pillow.
Cách dùng:
    python3 tach_hdld.py "Danh sach HĐLĐ/SCAN0036.pdf"   # 1 file
    python3 tach_hdld.py --all                            # tất cả SCAN*.pdf
"""
import os, re, sys, glob, csv, subprocess, unicodedata
from collections import defaultdict
from PIL import Image
import pytesseract
from pypdf import PdfReader, PdfWriter
import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HDLD_DIR = os.path.join(BASE_DIR, "Danh sach HĐLĐ")
EXCEL_PATH = os.path.join(BASE_DIR,
    "ĐL02_Danh sách sale và dự án tham gia bán 11.6.26 Full toàn bộ dữ liệu gốc.xlsx")
SHEET_NAME = "ĐL02 Full"
HEADER_ROW = 9
COL_STT, COL_NAME, COL_CCCD = 0, 2, 3   # A, C, D
OUT_DIR = os.path.join(BASE_DIR, "output_HDLD")
LOI_DIR = os.path.join(OUT_DIR, "_LOI")
CSV_PATH = os.path.join(BASE_DIR, "ket_qua_HDLD.csv")

EMP_CCCD = "040191008474"   # CCCD người sử dụng lao động (cố định) -> bỏ qua
DET_DPI = 200               # dpi OCR phát hiện trang đầu
HI_DPI  = 300               # dpi OCR lại trang đầu khi cần số CCCD chính xác
TMP = "/tmp/hdld_render"


# ---------- tiện ích chuỗi ----------
def fold(s):
    s = str(s).replace('đ','d').replace('Đ','D')
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn').lower()
    return re.sub(r'\s+',' ', s).strip()

def sanitize(name):
    name = re.sub(r'[/\\:*?"<>|]+', " ", str(name))
    return re.sub(r'\s+', " ", name).strip()


# ---------- Excel ----------
def load_excel(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    exact = {}
    stripm = defaultdict(list)
    for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
        stt  = row[COL_STT]  if len(row) > COL_STT  else None
        name = row[COL_NAME] if len(row) > COL_NAME else None
        cccd = row[COL_CCCD] if len(row) > COL_CCCD else None
        if not name:
            continue
        d = re.sub(r'\D', '', str(cccd or ''))
        if len(d) >= 9:
            exact[d] = (stt, str(name).strip())
            stripm[d.lstrip('0')].append((stt, str(name).strip(), d))
    wb.close()
    return exact, stripm


# ---------- render + OCR ----------
def render(pdf, pg, dpi):
    sub = os.path.join(TMP, re.sub(r'\W+', '_', os.path.basename(pdf)))
    os.makedirs(sub, exist_ok=True)
    pre = f"{sub}/p{pg}"
    for f in glob.glob(pre + "*"):
        os.remove(f)
    subprocess.run(["pdftoppm", "-f", str(pg), "-l", str(pg), "-r", str(dpi),
                    "-jpeg", pdf, pre], check=True, stderr=subprocess.DEVNULL)
    return sorted(glob.glob(pre + "*.jpg"))[0]

def ocr(path, lang='vie', config=''):
    return pytesseract.image_to_string(Image.open(path), lang=lang, config=config)

def is_first(ftext):
    has_so = re.search(r'mh\s*\d{3,5}\s*/\s*20\d\d\s*/\s*hdld', ftext) is not None
    has_title = 'hop dong lao dong' in ftext
    has_parties = ('nguoi su dung lao dong' in ftext) and ('nguoi lao dong' in ftext)
    return has_so or (has_title and has_parties and 'ho va ten' in ftext)

def extract_so(ftext):
    m = re.search(r'(mh\s*\d{3,5}\s*/\s*20\d\d\s*/\s*hdld[\w-]*)', ftext)
    return re.sub(r'\s+', '', m.group(1)).upper() if m else ''


# ---------- trích CCCD + tên NLĐ từ trang đầu ----------
def worker_section(raw):
    """Trả về (đoạn text mục 2 - NLĐ, vị trí). Cắt từ 'NGƯỜI LAO ĐỘNG' (mục 2)."""
    ft = fold(raw)
    i = ft.find('nguoi lao dong:')
    if i < 0:
        first = ft.find('nguoi su dung lao dong')
        i = ft.find('nguoi lao dong', (first + 5) if first >= 0 else 0)
    return raw[i:] if i >= 0 else raw

def extract_cccd_candidates(raw):
    seg = worker_section(raw)
    fseg = fold(seg)
    cands = []
    for m in re.finditer(r'(?:cccd|ho chieu|cmnd|can cuoc)[^0-9]{0,25}([0-9][0-9 ]{8,16}[0-9])', fseg):
        d = re.sub(r'\D', '', m.group(1))
        if 9 <= len(d) <= 12 and d != EMP_CCCD:
            cands.append(d)
    for m in re.finditer(r'\b(\d{9,12})\b', fseg):
        d = m.group(1)
        if d == EMP_CCCD:
            continue
        if len(d) in (10, 11) and d[0] == '0' and d[1] in '35789':  # giống SĐT
            continue
        if d not in cands:
            cands.append(d)
    # giữ thứ tự, ưu tiên 12 số
    cands.sort(key=lambda d: (0 if len(d) == 12 else 1))
    return cands

def extract_name(raw):
    seg = worker_section(raw)
    m = re.search(r'(?:Ông\s*/\s*bà|Ong\s*/\s*ba|Ơng|Ơ?ng/b[àa]|bà|ba)\s*[:\-]?\s*(.+?)\s*(?:Quốc|Quoc|Quéc|$)',
                  seg, re.I)
    if m:
        nm = re.sub(r'\s+', ' ', m.group(1)).strip(" :.-")
        # bỏ rác đầu (vd "f)lịgÍbà ")
        nm = re.sub(r'^[^A-Za-zÀ-ỹ]+', '', nm)
        return nm[:60]
    return ''

def match(cands, exact, stripm):
    """-> (stt, name, cccd_matched, used_cand) | None khi không khớp."""
    for d in cands:
        if d in exact:
            stt, name = exact[d]
            return stt, name, d, d
    for d in cands:
        rows = stripm.get(d.lstrip('0'))
        if rows and len(rows) == 1:
            stt, name, full = rows[0]
            return stt, name, full, d
    return None


# ---------- tách & lưu PDF ----------
def save_range(reader, start0, end0, out_path):
    w = PdfWriter()
    for p in range(start0, end0 + 1):
        w.add_page(reader.pages[p])
    with open(out_path, "wb") as fh:
        w.write(fh)


# ---------- xử lý 1 PDF ----------
def process_pdf(pdf_path, exact, stripm):
    name = os.path.basename(pdf_path)
    npg = int(subprocess.run(["pdfinfo", pdf_path], capture_output=True, text=True)
              .stdout.split("Pages:")[1].split()[0])
    print(f"\n### {name}: {npg} trang — đang phát hiện trang đầu hợp đồng...", flush=True)

    firsts = []   # (pg1based, so, raw_text_top)
    for pg in range(1, npg + 1):
        img = render(pdf_path, pg, DET_DPI)
        im = Image.open(img)
        w, h = im.size
        top = im.crop((0, 0, w, int(h * 0.72)))
        raw = pytesseract.image_to_string(top, lang='vie')
        if is_first(fold(raw)):
            firsts.append((pg, extract_so(fold(raw)), raw))
            print(f"  - trang {pg}: HỢP ĐỒNG {extract_so(fold(raw))}", flush=True)

    reader = PdfReader(pdf_path)
    results = []
    for k, (pg, so, raw) in enumerate(firsts):
        start0 = pg - 1
        end0 = (firsts[k + 1][0] - 2) if k + 1 < len(firsts) else (npg - 1)
        span = end0 - start0 + 1

        cands = extract_cccd_candidates(raw)
        m = match(cands, exact, stripm)
        if not m:  # OCR lại trang đầu ở dpi cao + whitelist số
            hi = render(pdf_path, pg, HI_DPI)
            raw2 = ocr(hi, 'vie')
            raw2 += "\n" + ocr(hi, 'eng', '--psm 6 -c tessedit_char_whitelist=0123456789')
            cands = list(dict.fromkeys(extract_cccd_candidates(raw) + extract_cccd_candidates(raw2)))
            m = match(cands, exact, stripm)
            if extract_name(raw) == '':
                raw = raw2

        ocr_name = extract_name(raw)
        if m:
            stt, ename, cccd_full, used = m
            out_name = sanitize(f"{stt}A - {ename}") + ".pdf"
            out_path = os.path.join(OUT_DIR, out_name)
            save_range(reader, start0, end0, out_path)
            print(f"    -> OK [{so}] CCCD {cccd_full} = {used} | {stt}A - {ename} "
                  f"(trang {pg}-{end0+1}, {span}t)", flush=True)
            results.append(dict(file_nguon=name, so_hd=so, trang_dau=pg, so_trang=span,
                                trang_thai="THÀNH CÔNG", stt=stt, ten=ename,
                                cccd=cccd_full, cccd_ocr=used, ten_ocr=ocr_name,
                                file_pdf=out_name, ghi_chu=""))
        else:
            tag = sanitize(f"LOI p{pg} {so} - {ocr_name}") or f"LOI p{pg}"
            out_path = os.path.join(LOI_DIR, tag[:80] + ".pdf")
            save_range(reader, start0, end0, out_path)
            note = (f"CCCD OCR={cands} không khớp Excel" if cands
                    else "không đọc được số CCCD")
            print(f"    -> LỖI [{so}] trang {pg}-{end0+1} | {note} | tên OCR='{ocr_name}'",
                  flush=True)
            results.append(dict(file_nguon=name, so_hd=so, trang_dau=pg, so_trang=span,
                                trang_thai="LỖI", stt="", ten="",
                                cccd="", cccd_ocr=";".join(cands), ten_ocr=ocr_name,
                                file_pdf=os.path.basename(out_path), ghi_chu=note))
    return results


def main():
    os.makedirs(TMP, exist_ok=True)
    os.makedirs(OUT_DIR, exist_ok=True)
    os.makedirs(LOI_DIR, exist_ok=True)
    try:
        sys.stdout.reconfigure(line_buffering=True)
    except Exception:
        pass

    args = sys.argv[1:]
    if not args:
        print(__doc__); sys.exit(0)
    if args == ["--all"]:
        pdfs = sorted(glob.glob(os.path.join(HDLD_DIR, "SCAN*.pdf")))
    else:
        pdfs = [a if os.path.isabs(a) else os.path.join(BASE_DIR, a) for a in args]

    exact, stripm = load_excel(EXCEL_PATH)
    print(f"Excel: {len(exact)} CCCD hợp lệ.")

    allres = []
    for pdf in pdfs:
        allres += process_pdf(pdf, exact, stripm)

    cols = ["file_nguon","so_hd","trang_dau","so_trang","trang_thai","stt","ten",
            "cccd","cccd_ocr","ten_ocr","file_pdf","ghi_chu"]
    csv_out = os.environ.get("CSV_OUT", CSV_PATH)
    # gộp với CSV cũ nếu chạy từng file (để không mất kết quả file trước)
    old = []
    if csv_out == CSV_PATH and len(pdfs) == 1 and os.path.exists(CSV_PATH):
        with open(CSV_PATH, encoding="utf-8-sig") as fh:
            old = [r for r in csv.DictReader(fh)
                   if r.get("file_nguon") != os.path.basename(pdfs[0])]
    with open(csv_out, "w", newline="", encoding="utf-8-sig") as fh:
        wr = csv.DictWriter(fh, fieldnames=cols)
        wr.writeheader()
        for r in old + allres:
            wr.writerow(r)

    ok = [r for r in allres if r["trang_thai"] == "THÀNH CÔNG"]
    err = [r for r in allres if r["trang_thai"] == "LỖI"]
    print("\n" + "=" * 72)
    print(f"=== THÀNH CÔNG ({len(ok)}) ===")
    for r in ok:
        print(f"  {r['file_pdf']:40s} | CCCD {r['cccd']} | {r['file_nguon']} tr{r['trang_dau']}")
    print(f"\n=== LỖI ({len(err)}) ===")
    for r in err:
        print(f"  [{r['file_nguon']} tr{r['trang_dau']} {r['so_hd']}] tênOCR='{r['ten_ocr']}'"
              f" | {r['ghi_chu']}")
    print("\n" + "=" * 72)
    print(f"Tổng hợp đồng: {len(allres)} | Thành công: {len(ok)} | Lỗi: {len(err)}")
    print(f"PDF: {OUT_DIR}\nLỗi: {LOI_DIR}\nCSV: {CSV_PATH}")


if __name__ == "__main__":
    main()
