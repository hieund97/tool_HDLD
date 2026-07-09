#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Sửa số CCCD Người lao động in trên trang đầu của 1 file PDF hợp đồng.
Phủ trắng vùng số cũ và in đè số mới (phông Times New Roman ~ giống văn bản).
Chỉ render lại TRANG 1; các trang sau giữ nguyên (copy lossless).

Dùng: python3 sua_cccd_pdf.py "<file.pdf>" <CCCD_moi> [--apply]
  (không --apply: chỉ tạo bản nháp *_SUA.pdf + preview, KHÔNG đụng file gốc)
  (--apply: sao lưu gốc -> *.bak.pdf rồi ghi đè file gốc)
"""
import sys, os, re, glob, subprocess
from PIL import Image, ImageDraw, ImageFont
import pytesseract
from pytesseract import Output
from pypdf import PdfReader, PdfWriter
import img2pdf

FONT = "/System/Library/Fonts/Supplemental/Times New Roman.ttf"
EMP = "040191008474"
DPI = 300

def render_page(pdf, pg, dpi=DPI):
    for f in glob.glob("/tmp/sc-*.jpg"):
        os.remove(f)
    subprocess.run(["pdftoppm","-f",str(pg),"-l",str(pg),"-r",str(dpi),"-jpeg",pdf,"/tmp/sc"],
                   check=True, stderr=subprocess.DEVNULL)
    return sorted(glob.glob("/tmp/sc-*.jpg"))[0]

def find_worker_cccd_box(im):
    """Trả về (box, old_digits). box=(x0,y0,x1,y1) bao quanh PHẦN SỐ (bỏ dấu ':')."""
    W, H = im.size
    d = pytesseract.image_to_data(im, lang='vie', output_type=Output.DICT)
    cands = []
    for i in range(len(d['text'])):
        digs = re.sub(r'\D', '', d['text'][i])
        if len(digs) >= 9:
            y = d['top'][i] / H
            cands.append((y, digs, i, d['text'][i]))
    # NLĐ ở nửa dưới; loại CCCD chủ DN (EMP) và SĐT (bắt đầu 0[35789] dài 10)
    worker = None
    for y, digs, i, raw in sorted(cands):
        if digs == EMP:
            continue
        if len(digs) in (10, 11) and digs[0] == '0' and digs[1] in '35789':
            continue
        if 0.55 <= y <= 0.78:
            worker = (i, raw, digs); break
    if worker is None:
        raise SystemExit("Không định vị được số CCCD Người lao động trên trang.")
    i, raw, digs = worker
    x, y, w, h = d['left'][i], d['top'][i], d['width'][i], d['height'][i]
    # bỏ phần dấu ':' ở đầu token (nếu có) để chỉ phủ phần số
    nchar = max(len(raw), 1)
    lead = len(raw) - len(raw.lstrip(': '))   # số ký tự ':'/space đầu
    cw = w / nchar
    x0 = int(x + lead * cw)
    box = (x0 - 4, y - 6, x + w + 8, y + h + 6)
    return box, digs

def overlay(im, box, new_text):
    x0, y0, x1, y1 = box
    draw = ImageDraw.Draw(im)
    # màu nền: lấy mẫu ngay phía trên-phải vùng số (giấy)
    sx, sy = min(x1 + 15, im.size[0]-1), max(y0 - 2, 0)
    bg = im.getpixel((sx, sy))
    if isinstance(bg, int): bg = (bg, bg, bg)
    draw.rectangle([x0, y0, x1, y1], fill=bg)
    # phông vừa chiều cao số
    target_h = (y1 - y0) - 10
    fs = target_h
    for _ in range(40):
        font = ImageFont.truetype(FONT, fs)
        bb = draw.textbbox((0, 0), new_text, font=font)
        if (bb[3]-bb[1]) <= target_h and (bb[2]-bb[0]) <= (x1-x0-6):
            break
        fs -= 1
    font = ImageFont.truetype(FONT, max(fs, 10))
    bb = draw.textbbox((0, 0), new_text, font=font)
    tx = x0 + 2
    ty = y0 + ((y1-y0) - (bb[3]-bb[1]))//2 - bb[1]
    draw.text((tx, ty), new_text, fill=(0, 0, 0), font=font)
    return im

def main():
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(1)
    src = sys.argv[1]
    new_cccd = re.sub(r'\D', '', sys.argv[2])
    apply = "--apply" in sys.argv
    if not os.path.isfile(src):
        sys.exit(f"Không thấy file: {src}")

    img = render_page(src, 1)
    im = Image.open(img).convert("RGB")
    box, old = find_worker_cccd_box(im)
    print(f"Số cũ trên văn bản: {old}  | vùng {box}")
    print(f"Số mới in đè      : {new_cccd}")
    im = overlay(im, box, new_cccd)

    # preview: cắt vùng quanh số
    pv = im.crop((box[0]-60, box[1]-90, box[2]+260, box[3]+40))
    pv_path = "/tmp/preview_cccd.png"; pv.save(pv_path)
    print("Preview:", pv_path)

    # ảnh trang 1 đã sửa -> PDF 1 trang (đúng tỉ lệ A4)
    p1_jpg = "/tmp/p1_edited.jpg"; im.save(p1_jpg, quality=92)
    reader = PdfReader(src)
    mb = reader.pages[0].mediabox
    w_pt, h_pt = float(mb.width), float(mb.height)
    rot = (reader.pages[0].get('/Rotate') or 0)
    # nếu trang gốc xoay 90/270 thì kích thước hiển thị là (h,w)
    if rot in (90, 270): w_pt, h_pt = h_pt, w_pt
    layout = img2pdf.get_fixed_dpi_layout_fun((DPI, DPI))
    p1_pdf = "/tmp/p1_edited.pdf"
    with open(p1_pdf, "wb") as fh:
        fh.write(img2pdf.convert(p1_jpg, layout_fun=layout))

    out = src[:-4] + "_SUA.pdf"
    writer = PdfWriter()
    writer.append(PdfReader(p1_pdf))           # trang 1 đã sửa
    for p in reader.pages[1:]:                  # trang 2..n giữ nguyên
        writer.add_page(p)
    with open(out, "wb") as fh:
        writer.write(fh)
    print("Bản nháp:", out, f"({len(writer.pages)} trang)")

    if apply:
        bak = src[:-4] + ".bak.pdf"
        if not os.path.exists(bak):
            os.replace(src, bak)
        else:
            os.remove(src)
        os.replace(out, src)
        print(f"ĐÃ ÁP DỤNG. Gốc lưu ở: {bak}")

if __name__ == "__main__":
    main()
