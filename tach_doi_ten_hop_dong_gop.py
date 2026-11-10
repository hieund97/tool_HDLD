# -*- coding: utf-8 -*-
"""
Tách file PDF scan gộp nhiều hợp đồng thành từng file, đặt tên:
    {STT hồ sơ}_{Số CCCD}_{Họ và tên}_{HĐLĐ|HĐTV|PHỤ LỤC SỬA ĐỔI THÔNG TIN}.pdf

- OCR phần trên mỗi trang (tiêu đề + dòng CCCD Người lao động), song song theo trang.
- Trang có tiêu đề HỢP ĐỒNG LAO ĐỘNG / HỢP ĐỒNG THỬ VIỆC / PHỤ LỤC = trang đầu văn bản mới.
- CCCD đối chiếu Excel (mặc định 'DỮ LIỆU CẦN ĐỔI/Import HDLD.xlsx').
- File gốc KHÔNG bị sửa. Kết quả ghi vào <folder>/ĐÃ TÁCH, đoạn không xác định -> ĐÃ TÁCH/_LOI.
- Cache OCR theo (file, trang): chạy lại chỉ OCR trang chưa có.

Chạy:
    python3 tach_doi_ten_hop_dong_gop.py "ĐỔI TÊN LẦN 3 9.11"
    python3 tach_doi_ten_hop_dong_gop.py "<folder>" --excel "<xlsx>"
"""
import os
os.environ.setdefault("OMP_THREAD_LIMIT", "1")

import re, sys, csv, json, glob, shutil, tempfile, subprocess
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed

import pytesseract
from PIL import Image
from pypdf import PdfReader, PdfWriter

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import doi_ten_hop_dong_vin as H  # noqa: E402  (classify_page, kind_from_number, load_excel, ...)

DPI      = 200
TOP_FRAC = 0.6
WORKERS  = os.cpu_count() or 4
OUTDIR   = "ĐÃ TÁCH"


def ocr_page(pdf, page):
    """OCR phần trên 1 trang. Tự xoay 180° nếu đọc ra quá ít chữ."""
    with tempfile.TemporaryDirectory() as tmp:
        prefix = os.path.join(tmp, "p")
        subprocess.run(["pdftoppm", "-r", str(DPI), "-gray", "-jpeg", "-f", str(page), "-l", str(page),
                        pdf, prefix], check=True, stderr=subprocess.DEVNULL)
        im = Image.open(glob.glob(prefix + "*.jpg")[0])
        im.load()
    best = ""
    for rot in (0, 180):
        x = im.rotate(rot, expand=True) if rot else im
        txt = pytesseract.image_to_string(x.crop((0, 0, x.width, int(x.height * TOP_FRAC))), lang="vie")
        if len(re.findall(r"[A-Za-zÀ-ỹ]{3,}", txt)) > len(re.findall(r"[A-Za-zÀ-ỹ]{3,}", best)):
            best = txt
        if len(re.findall(r"[A-Za-zÀ-ỹ]{3,}", txt)) >= 40:
            break
    return best


def digit_rescue(pdf, page):
    """Đọc lại vùng số CCCD ở 400 DPI (cho trang tiêu đề chưa khớp CCCD)."""
    with tempfile.TemporaryDirectory() as tmp:
        prefix = os.path.join(tmp, "p")
        subprocess.run(["pdftoppm", "-r", "400", "-jpeg", "-f", str(page), "-l", str(page), pdf, prefix],
                       check=True, stderr=subprocess.DEVNULL)
        im = H.oriented(Image.open(glob.glob(prefix + "*.jpg")[0]).convert("RGB"))
    out = []
    for x0, y0, x1, y1 in H.value_crops(im):
        if x1 - x0 >= 20 and y1 - y0 >= 5:
            out += H.digitocr(im.crop((x0, y0, x1, y1)))
    return out


def _job(args):
    pdf, page = args
    try:
        return pdf, page, ocr_page(pdf, page), None
    except Exception as e:
        return pdf, page, "", str(e)[:100]


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        sys.exit(__doc__)
    folder = os.path.join(BASE, args[0])
    if "--excel" in sys.argv:
        H.EXCEL = os.path.join(BASE, sys.argv[sys.argv.index("--excel") + 1])
    mapping = H.load_excel()
    ccset = set(mapping)

    tag = re.sub(r"\W+", "_", H.strip_acc(os.path.basename(folder.rstrip("/")))).strip("_")
    cache_path = os.path.join(tempfile.gettempdir(), f"ocr_tach_{tag}.json")
    cache = json.load(open(cache_path, encoding="utf-8")) if os.path.exists(cache_path) else {}

    pdfs = sorted(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))
    npages = {f: len(PdfReader(os.path.join(folder, f)).pages) for f in pdfs}
    total = sum(npages.values())
    todo = [(os.path.join(folder, f), p) for f in pdfs for p in range(1, npages[f] + 1)
            if f"{f}|{p}" not in cache]
    print(f"{len(pdfs)} file, {total} trang | cần OCR {len(todo)} trang | {WORKERS} tiến trình")

    if todo:
        done = 0
        with ProcessPoolExecutor(max_workers=WORKERS) as ex:
            for fut in as_completed([ex.submit(_job, t) for t in todo]):
                pdf, page, txt, err = fut.result()
                done += 1
                if err:
                    print(f"  ✗ {os.path.basename(pdf)} trang {page}: {err}")
                    continue
                cache[f"{os.path.basename(pdf)}|{page}"] = txt
                if done % 50 == 0 or done == len(todo):
                    print(f"  OCR {done}/{len(todo)} trang", flush=True)
                    with open(cache_path, "w", encoding="utf-8") as fp:
                        json.dump(cache, fp, ensure_ascii=False)

    # ---------------------------------------------------------- tách đoạn
    out = os.path.join(folder, OUTDIR)
    if os.path.exists(out):
        shutil.rmtree(out)
    os.makedirs(os.path.join(out, "_LOI"))
    rows, used = [], Counter()

    for f in pdfs:
        segs = []  # [start, end, kind, cccd, note]
        for p in range(1, npages[f] + 1):
            txt = cache.get(f"{f}|{p}", "")
            kind = H.classify_page(txt)
            if kind is None and re.search(r"S[ỐO]\s*:?\s*MH\d+", H.strip_acc(txt).upper()):
                kind = H.kind_from_number(txt)
            nums = [c for c in dict.fromkeys(H.nums12(txt)) if c != H.EMP]
            inx = [c for c in nums if c in ccset]
            cccd = inx[0] if inx else None
            if kind:
                cur = segs[-1] if segs else None
                # trang tiêu đề lặp lại cùng người, cùng loại (scan trùng trang 1) -> gộp
                if cur and cur[2] == kind and cccd and cur[3] == cccd and p - cur[1] == 1:
                    cur[1] = p; cur[4].append(f"trang {p} lặp tiêu đề")
                    continue
                segs.append([p, p, kind, cccd, [] if cccd else [f"OCR: {','.join(nums) or 'không có số'}"]])
            elif segs:
                segs[-1][1] = p
                if not segs[-1][3] and inx:
                    segs[-1][3] = cccd
            else:
                segs.append([p, p, None, cccd, ["không có tiêu đề ở đầu file"]])

        # đọc lại vùng số cho đoạn chưa có CCCD
        for s in segs:
            if s[3] is None:
                c = [x for x in dict.fromkeys(digit_rescue(os.path.join(folder, f), s[0])) if x != H.EMP]
                hit = [x for x in c if x in ccset]
                if hit:
                    s[3] = hit[0]; s[4] = ["CCCD đọc lại ở 400 DPI"]
                elif c:
                    s[4].append("đọc lại: " + ",".join(c))

        reader = PdfReader(os.path.join(folder, f))
        for a, b, kind, cccd, note in segs:
            if kind and cccd:
                stt, name = mapping[cccd]
                base = f"{stt}_{cccd}_{name}_{kind}"
                used[base] += 1
                fname = base + (f"_{used[base]}" if used[base] > 1 else "") + ".pdf"
                if used[base] > 1:
                    note.append("trùng tên đích - kiểm tra tay")
                dst, st = os.path.join(out, fname), "OK"
            else:
                fname = f"{os.path.splitext(f)[0]}_trang_{a}-{b}.pdf"
                dst = os.path.join(out, "_LOI", fname)
                st = "KHÔNG_RÕ_LOẠI" if cccd else "KHÔNG_KHỚP_CCCD"
            w = PdfWriter()
            for i in range(a - 1, b):
                w.add_page(reader.pages[i])
            with open(dst, "wb") as fp:
                w.write(fp)
            rows.append([st, f, f"{a}-{b}", b - a + 1, kind or "", cccd or "", fname, "; ".join(note)])

    csv_path = os.path.join(BASE, f"ket_qua_tach_{tag}.csv")
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as fp:
        w = csv.writer(fp)
        w.writerow(["Trạng thái", "File gộp", "Trang", "Số trang", "Loại", "CCCD", "File tách", "Ghi chú"])
        w.writerows(rows)

    print("\n================ TỔNG HỢP ================")
    print(f"Văn bản tách được: {len(rows)}")
    for k, v in Counter(r[0] for r in rows).most_common():
        print(f"  {k:<18}: {v}")
    print("Theo loại:", dict(Counter(r[4] for r in rows if r[0] == "OK")))
    print("Số trang / văn bản:", dict(sorted(Counter((r[4], r[3]) for r in rows).items(), key=str)))
    for r in rows:
        if r[0] != "OK" or r[7]:
            print(f"  - [{r[0]}] {r[1]} trang {r[2]} | {r[4]} | {r[5]} | {r[7]}")
    print(f"\nKết quả: {out}\nCSV: {csv_path}")


if __name__ == "__main__":
    main()
