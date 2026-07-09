# tool_HDLD — Bộ công cụ xử lý CCCD & Hợp đồng lao động

Tập hợp các script Python dùng để **xử lý hàng loạt** ảnh CCCD, hợp đồng lao động (HĐLĐ) và
phụ lục hợp đồng (PLHĐ), tự động **OCR số CCCD** rồi đối chiếu với danh sách trong file Excel
để đặt tên / tách / ghép file theo đúng người.

---

## Yêu cầu môi trường

Cần 3 thành phần: **Python 3**, **Tesseract OCR** (kèm gói tiếng Việt `vie`) và **Poppler**
(cung cấp `pdftoppm`, `pdfinfo`). Chọn hướng dẫn theo hệ điều hành bên dưới.

### macOS (Homebrew)

```bash
brew install python tesseract tesseract-lang poppler
pip3 install pillow pytesseract openpyxl pypdf img2pdf
```

### Linux (Ubuntu / Debian)

```bash
# Công cụ hệ thống
sudo apt update
sudo apt install -y python3 python3-pip tesseract-ocr tesseract-ocr-vie poppler-utils

# Thư viện Python
pip3 install pillow pytesseract openpyxl pypdf img2pdf
```

> Fedora / RHEL: `sudo dnf install python3 python3-pip tesseract tesseract-langpack-vie poppler-utils`
> Arch: `sudo pacman -S python python-pip tesseract tesseract-data-vie poppler`

### Windows

1. **Python 3** — tải tại <https://www.python.org/downloads/>, khi cài **tích chọn "Add Python to PATH"**.
2. **Tesseract OCR** — tải bộ cài từ <https://github.com/UB-Mannheim/tesseract/wiki>.
   - Trong lúc cài, ở bước *"Additional language data"* nhớ **chọn `Vietnamese`** (gói `vie`).
   - Mặc định cài vào `C:\Program Files\Tesseract-OCR`. Thêm đường dẫn này vào biến môi trường **PATH**.
3. **Poppler** — tải bản Windows tại <https://github.com/oschwartz10612/poppler-windows/releases>,
   giải nén rồi thêm thư mục `...\poppler\Library\bin` vào **PATH**.
4. **Thư viện Python** (mở PowerShell hoặc CMD):
   ```powershell
   pip install pillow pytesseract openpyxl pypdf img2pdf
   ```

> **Lưu ý Windows:** nếu đã thêm Tesseract vào PATH mà script vẫn báo không tìm thấy,
> mở đầu file `.py` thêm dòng chỉ đường dẫn tuyệt đối:
> ```python
> pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
> ```
> Ngoài ra, `sua_cccd_pdf.py` dùng phông `Times New Roman` theo đường dẫn macOS
> (`/System/Library/Fonts/...`) — trên Windows cần sửa biến `FONT` thành `C:\Windows\Fonts\times.ttf`.

### Kiểm tra đã cài đúng

```bash
python3 --version        # hoặc: python --version  (Windows)
tesseract --version
tesseract --list-langs   # phải thấy 'vie' trong danh sách
pdftoppm -v
```

---

## Danh sách công cụ

### 1. `ghep_anh_CCCD.py` — Ghép 2 mặt ảnh CCCD thành 1 PDF
Gom ảnh mặt trước + mặt sau của mỗi người, OCR số CCCD, đối chiếu Excel để lấy STT + tên,
rồi ghép thành 1 file PDF đặt tên `"{STT}B - {Tên}.pdf"`.

- Chỉ khớp theo **số CCCD** (không fallback theo tên); không đọc được CCCD → ghi LỖI.
- Tự gom cặp trước/sau kể cả khi tên file sai dấu (khớp gần đúng ≤ 1 ký tự).
- Chạy lại nhiều lần được (bỏ qua PDF đã tạo).
- Xuất `ket_qua_tong_hop.csv`.

```bash
python3 ghep_anh_CCCD.py
```

### 2. `tach_hdld_tu_pdf.py` — Tách file scan nhiều HĐLĐ thành từng file riêng
Một file scan chứa nhiều hợp đồng. Script OCR để phát hiện trang đầu mỗi hợp đồng
(dấu hiệu `SỐ: MH####/20##/HĐLĐ` + tiêu đề + hai bên ký), lấy CCCD **Người lao động**
(bỏ CCCD chủ doanh nghiệp cố định), đối chiếu Excel rồi tách thành `"{STT}A - {Tên}.pdf"`.

- File không khớp CCCD → lưu vào thư mục con `_LOI` để kiểm tra tay.
- Xuất `ket_qua_HDLD.csv`.

```bash
python3 tach_hdld_tu_pdf.py "Danh sach HĐLĐ/SCAN0036.pdf"   # 1 file
python3 tach_hdld_tu_pdf.py --all                            # tất cả SCAN*.pdf
```

### 3. `sua_cccd_pdf.py` — Sửa số CCCD in trên trang đầu 1 file PDF hợp đồng
Phủ trắng vùng số CCCD cũ của Người lao động và in đè số mới (phông Times New Roman).
Chỉ render lại **trang 1**, các trang sau giữ nguyên (copy lossless).

```bash
# Chỉ tạo bản nháp *_SUA.pdf + preview, KHÔNG đụng file gốc:
python3 sua_cccd_pdf.py "<file.pdf>" <CCCD_moi>

# Áp dụng đè lên file gốc (tự sao lưu *.bak.pdf):
python3 sua_cccd_pdf.py "<file.pdf>" <CCCD_moi> --apply
```

### 4. `doi_ten_cccd_ten_viet_lien.py` — Đổi tên ảnh CCCD theo `STT_CCCD_HọTênViếtLiền`
Tách số từ tên file, đối chiếu tập CCCD trong Excel, đổi tên ảnh thành
`STT_CCCD_HoVaTen.ext` (họ tên viết liền, không dấu, hoa chữ đầu mỗi từ).

- Xuất `ket_qua_doi_ten_cccd.csv`.

```bash
python3 doi_ten_cccd_ten_viet_lien.py
```

### 5. `doi_ten_dua_theo_cccd_ten_viet_hoa_co_dau_cach.py` — Đổi tên PDF phụ lục HĐ (PLHĐ)
OCR nội dung PDF để lấy CCCD Người lao động (dòng "Số CMND/CCCD"), đối chiếu Excel,
**copy** file sang thư mục `output_PLHĐ` với tên `STT_PLHĐ_CCCD_Họ và tên.pdf`.

- Không đổi tên tại chỗ (file gốc được giữ nguyên).
- Tự phát hiện file trùng CCCD, chống ghi đè.
- Xuất `ket_qua_doi_ten_PLHD3.csv`.

```bash
python3 doi_ten_dua_theo_cccd_ten_viet_hoa_co_dau_cach.py
```

---

## Ghi chú chung

- **Nguồn đối chiếu:** file Excel danh sách sale (sheet `ĐL02` / `ĐL02 Full`), lấy STT (cột A),
  Họ tên (cột C), số CCCD (cột D hoặc E tuỳ script). Đường dẫn Excel và tên thư mục dữ liệu
  được cấu hình ở đầu mỗi file (`BASE_DIR`, `EXCEL_PATH`, ...) — chỉnh lại cho phù hợp máy bạn.
- Mọi script đều xuất **báo cáo CSV** (mã hoá `utf-8-sig` để mở đúng trong Excel) và in
  tóm tắt Thành công / Lỗi ra terminal.
- Xử lý tiếng Việt: chuẩn hoá NFC/NFD, bỏ dấu, khớp gần đúng (Levenshtein ≤ 1) để chịu được
  sai lệch OCR và lỗi dấu trong tên file.
