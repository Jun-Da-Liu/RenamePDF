# HƯỚNG DẪN SỬ DỤNG CHI TIẾT
## Phần mềm Tự động Nhận diện OCR & Đặt tên Hàng loạt File PDF Hồ sơ (`PhanMemDoiTenPDF`)

---

## MỤC LỤC
1. [Giới thiệu tổng quan](#1-giới-thiệu-tổng-quan)
2. [Cấu trúc thư mục phần mềm](#2-cấu-trúc-thư-mục-phần-mềm)
3. [Cách khởi động phần mềm](#3-cách-khởi-động-phần-mềm)
4. [Quy trình sử dụng chuẩn (4 bước an toàn)](#4-quy-trình-sử-dụng-chuẩn-4-bước-an-toàn)
   - [Bước 1: Chọn thư mục hồ sơ PDF & Cấu hình thông số](#bước-1-chọn-thư-mục-hồ-sơ-pdf--cấu-hình-thông-số)
   - [Bước 2: Tải lên (Upload) hoặc Chỉnh sửa Danh mục Excel đặt tên](#bước-2-tải-lên-upload-hoặc-chỉnh-sửa-danh-mục-excel-đặt-tên)
   - [Bước 3: Quét & Dự kiến đổi tên (Dry-Run), Xem trước PDF và Sửa tên trực tiếp](#bước-3-quét--dự-kiến-đổi-tên-dry-run-xem-trước-pdf-và-sửa-tên-trực-tiếp)
   - [Bước 4: Thực thi đổi tên thật & Hoàn tác (Rollback)](#bước-4-thực-thi-đổi-tên-thật--hoàn-tác-rollback)
5. [Quy chuẩn định dạng File Excel Danh mục tải lên](#5-quy-chuẩn-định-dạng-file-excel-danh-mục-tải-lên)
6. [Ý nghĩa các trạng thái phân loại & Quy tắc đánh số trùng lặp](#6-ý-nghĩa-các-trạng-thái-phân-loại--quy-tắc-đánh-số-trùng-lặp)
7. [Danh sách các File Báo cáo Excel xuất tự động](#7-danh-sách-các-file-báo-cáo-excel-xuất-tự-động)
8. [Câu hỏi thường gặp & Xử lý sự cố (FAQ)](#8-câu-hỏi-thường-gặp--xử-lý-sự-cố-faq)
9. [Đóng gói chạy offline](#9-đóng-gói-chạy-offline)

---

## 1. Giới thiệu tổng quan

**`PhanMemDoiTenPDF`** là phần mềm chuyên dụng giúp tự động nhận diện chữ tiếng Việt (OCR) từ các file PDF scan hồ sơ (hồ sơ đảng viên, hồ sơ cán bộ, tài liệu lưu trữ...) và tự động đổi tên hàng loạt theo **Danh mục tên chuẩn trong file Excel**.

### Các ưu điểm cốt lõi:
- **Tải lên danh mục Excel linh hoạt:** Cho phép người dùng tải lên (Upload) bất kỳ file `.xlsx` danh mục tên mới nào hoặc thêm/sửa/xóa mục đặt tên ngay trên giao diện.
- **Bộ nhớ đệm thông minh (SQLite OCR Cache):** Mỗi file PDF chỉ cần chạy OCR 1 lần duy nhất. Khi bạn thay đổi file Excel danh mục đặt tên hoặc chỉnh sửa quy tắc, việc đối chiếu lại hàng chục nghìn file PDF chỉ mất **vài giây**.
- **Giao diện trực quan tích hợp Xem trước PDF:** Xem trực tiếp trang tài liệu PDF, đối chiếu nội dung OCR và chỉnh sửa tên dự kiến ngay trong một màn hình mà không cần mở phần mềm đọc PDF bên ngoài.
- **Tuyệt đối an toàn cho dữ liệu gốc:**
  - Luôn chạy chế độ **Dự kiến (Dry-Run)** trước để kiểm duyệt.
  - Khi đổi tên thật, sử dụng thuật toán **Đổi tên 2 pha (Two-Phase Rename)** qua tên tạm để không bao giờ xảy ra xung đột hoặc ghi đè mất file.
  - Hỗ trợ **Hoàn tác (Rollback) 1 chạm** để trả toàn bộ file về tên gốc ban đầu bất cứ lúc nào.

---

## 2. Cấu trúc thư mục phần mềm

Thư mục dự án bao gồm:

```text
PhanMemDoiTenPDF/
├── Chay_Phan_Mem_Doi_Ten.bat          # File khởi động nhanh (Nhấp đúp chuột để mở phần mềm)
├── app_gui.py                         # Mã nguồn Giao diện trực quan (Desktop GUI)
├── core_engine.py                     # Nhân xử lý OCR, đối chiếu danh mục Excel, Dry-Run & Rollback
├── config.json                        # File tự động lưu cấu hình lần chạy gần nhất
├── README.md                          # Tài liệu hướng dẫn sử dụng chi tiết (file này)
└── data\
    └── Danh_muc_tai_lieu_khong_dau.xlsx  # Thư mục lưu trữ file Excel danh mục đang áp dụng
```

Các file báo cáo Excel và bộ nhớ đệm OCR được lưu tại `Documents/PhanMemDoiTenPDF_Reports` theo mặc định.

---

## 3. Cách khởi động phần mềm

### Cách 1: Mở trực tiếp bằng file `.bat` (Khuyên dùng)
- Mở thư mục dự án trong File Explorer.
- Nhấp đúp chuột vào file **`Chay_Phan_Mem_Doi_Ten.bat`**.
- Cửa sổ giao diện chính của phần mềm sẽ hiển thị ngay lập tức.

### Cách 2: Chạy bằng dòng lệnh (PowerShell / Command Prompt)
```powershell
cd <đường-dẫn-thư-mục-dự-án>
py -3 -X utf8 app_gui.py
```

Khi chạy từ mã nguồn để OCR, cần đặt biến môi trường `TESSERACT_DIR` tới thư mục chứa `libtesseract-5.dll` và thư mục `tessdata` trước khi mở ứng dụng.

---

## 4. Quy trình sử dụng chuẩn (4 bước an toàn)

### Bước 1: Chọn thư mục hồ sơ PDF & Cấu hình thông số
Trên khung cấu hình ở phần đầu giao diện:
1. **📁 Thư mục hồ sơ PDF:**
   - Mặc định trỏ tới thư mục Documents của tài khoản Windows.
   - Bấm **`Chọn thư mục...`** để chọn nơi chứa PDF trên bất kỳ ổ đĩa nào. Phần mềm sẽ quét đệ quy toàn bộ thư mục con bên trong.
2. **Số luồng OCR:**
   - Mặc định là `10` (phù hợp với CPU nhiều nhân).
   - Nếu máy cấu hình thấp hoặc đang mở nhiều ứng dụng nặng, có thể giảm xuống `4` – `6`. Nếu máy cấu hình cao, có thể tăng lên `12` – `16`.
3. **Tùy chọn *"Đặt tên theo tiêu đề trang đầu nếu ngoài danh mục Excel"*:**
   - **Nếu bật (Tích chọn - Mặc định):** Các giấy tờ có tiêu đề rõ ràng nhưng không nằm trong danh mục Excel (ví dụ: *Mục lục tài liệu trong hồ sơ đảng viên*, *Trích biên bản sinh hoạt chi đoàn*, *Bản cam kết*...) sẽ được đặt tên không dấu theo đúng tiêu đề trang đầu và đánh dấu là `Ngoài Excel`.
   - **Nếu tắt (Bỏ tích):** Chỉ những file khớp với danh mục Excel mới được dự kiến đổi tên; tất cả các tài liệu ngoài danh mục Excel sẽ được chuyển sang trạng thái `Cần kiểm tra` (giữ nguyên tên gốc).

---

### Bước 2: Tải lên (Upload) hoặc Chỉnh sửa Danh mục Excel đặt tên

Bạn có thể thay đổi danh sách tên chuẩn bất cứ lúc nào theo 2 cách:

#### Cách A: Tải lên file Excel (`.xlsx`) từ máy tính
1. Bấm nút màu xanh **`📤 Tải lên File Excel danh mục mới (.xlsx)...`** trên thanh cấu hình (hoặc trong **Tab 2**).
2. Chọn file `.xlsx` chứa danh sách tên tài liệu của bạn.
3. Phần mềm sẽ tự động:
   - Kiểm tra và đọc toàn bộ danh mục trong file Excel.
   - Sao chép file Excel đó vào thư mục `PhanMemDoiTenPDF\data\` và lưu làm danh mục mặc định mới.
   - Hiển thị tổng số mục đã nạp ở nhãn thống kê (ví dụ: `103 mục đặt tên`).
   - Hỏi bạn có muốn **chạy đối chiếu lại ngay lập tức** với danh mục Excel vừa tải lên không. Nếu chọn **Yes**, phần mềm sẽ cập nhật lại toàn bộ kết quả dự kiến chỉ trong vài giây.

#### Cách B: Thêm / Sửa / Xóa trực tiếp trên phần mềm (Tab 2)
Chuyển sang tab **`📊 2. Quản lý Danh mục Excel đặt tên (Tải lên / Chỉnh sửa)`**:
- **Thêm mục mới:** Bấm **`➕ Thêm mục đặt tên mới`** -> Nhập *Tên tài liệu có dấu* (phần mềm sẽ tự động điền gợi ý *Tên đặt File không dấu an toàn*) -> Nhập *Từ khóa nhận diện* (nếu cần) -> Bấm **`💾 Lưu vào Danh mục Excel`**.
- **Sửa mục hiện có:** Nhấp đúp chuột vào một dòng trong bảng (hoặc chọn dòng và bấm **`✏️ Sửa mục đang chọn`**) -> Chỉnh sửa tên -> Bấm **Lưu**.
- **Xóa mục:** Chọn dòng cần xóa và bấm **`🗑️ Xóa mục đang chọn`**.
- **Xuất ra file Excel:** Bấm **`📥 Xuất danh mục ra file Excel...`** hoặc **`📥 Tải file Excel mẫu`** để lưu danh mục hiện tại ra file `.xlsx` phục vụ lưu trữ hoặc chỉnh sửa hàng loạt trên Microsoft Excel.

---

### Bước 3: Quét & Dự kiến đổi tên (Dry-Run), Xem trước PDF và Sửa tên trực tiếp

1. **Chạy quét & đối chiếu (Dry-Run):**
   - Bấm nút **`🔍 1. Quét & Dự kiến đổi tên (Dry-Run)`**.
   - Thanh tiến trình sẽ hiển thị trạng thái theo thời gian thực: số file đã có trong bộ nhớ đệm, số file đang chạy OCR mới, tốc độ xử lý (`file/giây`) và thời gian hoàn thành dự kiến (`ETA`).
   - Bạn có thể bấm nút **`⏹ Dừng`** bất cứ lúc nào nếu muốn tạm ngưng.
   - **Lưu ý quan trọng:** Ở bước này, phần mềm **CHƯA** đổi tên bất kỳ file PDF gốc nào trên ổ đĩa mà chỉ lập danh sách dự kiến để bạn kiểm tra.

2. **Tra cứu, Lọc & Xem trước trực tiếp trang PDF (Tab 1):**
   Tại tab **`📋 1. Danh sách File PDF & Xem trước trực quan`**:
   - **Bộ lọc nhanh:**
     - *Lọc trạng thái:* Chọn xem riêng nhóm `EXCEL - Khớp danh mục Excel`, `OUTSIDE - Ngoài danh mục Excel`, `CHECK - Cần kiểm tra thủ công`, hoặc `DUP - Có trùng lặp (.1, .2...)`.
     - *Phòng ban:* Lọc riêng từng Phòng ban / Đơn vị.
     - *Tìm kiếm:* Gõ tên cán bộ/đảng viên, tên file cũ, tên dự kiến hoặc từ khóa OCR rồi nhấn `Enter` (hoặc bấm `Lọc`).
   - **Xem trước trang PDF ngay trong phần mềm:**
     - Bấm vào bất kỳ dòng nào trong bảng bên trái, khung bên phải sẽ hiển thị ngay **hình ảnh trang đầu của file PDF** đó.
     - Dùng nút **`◀`** / **`▶`** ở góc trên khung ảnh để lật xem các trang tiếp theo của file PDF.
     - Bấm **`🔍 Mở PDF`** (hoặc nhấp đúp chuột vào dòng trong bảng) để mở file bằng phần mềm đọc PDF mặc định của Windows.
     - Bấm **`📂 Mở thư mục`** để mở ngay thư mục chứa file PDF đó trong Windows Explorer.

3. **Sửa hoặc gán tên thủ công cho từng file ngay trên giao diện:**
   - Với các file ở trạng thái `Cần kiểm tra` (hoặc bất kỳ file nào bạn muốn đổi sang tên khác):
     1. Chọn dòng file đó ở bảng bên trái để xem hình ảnh trang PDF bên phải.
     2. Tại mục **"Chọn nhanh từ danh mục Excel"**, bạn có thể gõ vài chữ cái để lọc nhanh hoặc chọn từ danh sách xổ xuống.
     3. Hoặc gõ trực tiếp tên mong muốn vào ô **"Tên file mới dự kiến"**.
     4. Bấm nút màu xanh lá **`💾 Áp dụng tên này`**.
   - Ngay lập tức, dòng đó sẽ được chuyển sang trạng thái đã duyệt (`100%`) và lưu tự động vào kế hoạch đổi tên cũng như cập nhật lại các file báo cáo Excel.

---

### Bước 4: Thực thi đổi tên thật & Hoàn tác (Rollback)

1. **Thực thi đổi tên thật (`✅ 2. Thực thi đổi tên thật`):**
   - Sau khi đã hài lòng với danh sách dự kiến ở Bước 3, bấm nút màu xanh lá **`✅ 2. Thực thi đổi tên thật`**.
   - Phần mềm sẽ hiển thị hộp thoại xác nhận tổng số file chuẩn bị đổi tên.
   - Khi xác nhận **Yes**:
     - Tất cả các file đã nhận diện được (`Khớp Excel` và `Ngoài Excel`, cùng các file bạn đã gán tên thủ công) sẽ được đổi sang tên mới qua **2 pha an toàn**.
     - Các file nằm trong nhóm `Cần kiểm tra (Chưa đổi tên)` sẽ được **giữ nguyên tên gốc** để đảm bảo an toàn.
     - Toàn bộ lịch sử đổi tên (tên cũ, tên mới, đường dẫn đầy đủ, thời gian) được ghi vào `Nhat_ky_doi_ten.xlsx` và `nhat_ky_doi_ten.json`.

2. **Hoàn tác về tên gốc (`⏪ 3. Hoàn tác (Rollback)`):**
   - Nếu sau khi đổi tên thật, bạn muốn trả toàn bộ các file đã đổi về lại tên gốc ban đầu (như trước khi chạy), chỉ cần bấm nút màu cam **`⏪ 3. Hoàn tác (Rollback)`** và chọn **Yes**.
   - Phần mềm sẽ đọc file nhật ký gần nhất và khôi phục nguyên trạng 100% tên file ban đầu.

---

## 5. Quy chuẩn định dạng File Excel Danh mục tải lên

Khi bạn chuẩn bị một file Excel (`.xlsx`) mới để tải lên phần mềm, phần mềm hỗ trợ tự động nhận diện **3 kiểu bố cục cột** dưới đây (trên Sheet đầu tiên của file Excel):

### Kiểu 1: Bố cục 2 cột chuẩn (Khuyên dùng)
| Cột A: `Tên đặt File/Folder an toàn (Bỏ dấu '/')` | Cột B: `Tên tài liệu (Có dấu gốc)` |
| :--- | :--- |
| `01. Ly lich cua nguoi vao Dang` | `Lý lịch của người vào Đảng` |
| `02. Ly lich dang vien` | `Lý lịch đảng viên` |
| `03. Phieu dang vien` | `Phiếu đảng viên` |

### Kiểu 2: Bố cục 3 cột (Có thêm cột Từ khóa nhận diện tùy chỉnh)
Nếu bạn muốn bổ sung thêm các loại tài liệu đặc thù mới và chỉ định sẵn các cụm từ khóa nhận diện khi quét OCR, hãy thêm **Cột C**:
| Cột A: `Tên đặt File (Không dấu)` | Cột B: `Tên tài liệu (Có dấu gốc)` | Cột C: `Từ khóa nhận diện bổ sung (Cách nhau dấu phẩy)` |
| :--- | :--- | :--- |
| `104. Phieu danh gia dang vien cuoi nam` | `Phiếu đánh giá đảng viên cuối năm` | `phieu danh gia cuoi nam, mau dg 01` |
| `105. Giay xac nhan nhan than` | `Giấy xác nhận nhân thân` | `xac nhan nhan than, cong an xa xac nhan` |

> **Cơ chế ưu tiên:** Khi quét OCR, nếu văn bản trang đầu chứa bất kỳ cụm từ khóa nào trong **Cột C**, phần mềm sẽ ưu tiên khớp ngay vào mục đó với độ tin cậy `98%`.

### Kiểu 3: Bố cục 1 cột (Tự động chuyển sang không dấu)
Nếu file Excel của bạn chỉ có duy nhất 1 cột chứa tên tài liệu tiếng Việt có dấu (ví dụ: `01. Lý lịch của người vào Đảng`), phần mềm sẽ **tự động loại bỏ dấu tiếng Việt và các ký tự cấm của Windows** (`\ / : * ? " < > |`) để tạo ra tên file không dấu an toàn cho bạn.

*(Mẹo: Bạn có thể bấm nút **`📥 Tải file Excel mẫu`** ngay trên phần mềm để lấy file chuẩn về chỉnh sửa).*

---

## 6. Ý nghĩa các trạng thái phân loại & Quy tắc đánh số trùng lặp

### 6.1. Các trạng thái phân loại
- 🟢 **`Khớp Excel` (`EXCEL`):** Nội dung OCR của file PDF khớp với một trong các mục trong file Excel danh mục hiện tại. File sẽ được đặt tên theo đúng cột *Tên đặt File không dấu* trong Excel.
- 🔵 **`Ngoài Excel` (`OUTSIDE`):** Tài liệu có tiêu đề rõ ràng ở đầu trang (ví dụ: *Mục lục tài liệu trong hồ sơ đảng viên*, *Trích biên bản sinh hoạt chi bộ*, *Đơn xin kết hôn*...) nhưng không nằm trong danh sách Excel. Nếu bật tùy chọn cho phép đặt tên ngoài Excel, file sẽ được đặt tên không dấu theo tiêu đề đó.
- 🔴 **`Cần kiểm tra` (`CHECK`):** Các trang tài liệu bị mờ, viết tay không rõ chữ, trang trắng, hoặc là trang lẻ bị tách rời khỏi bộ hồ sơ (không có tiêu đề ở trang 1). Phần mềm **giữ nguyên tên gốc** và liệt kê riêng để người dùng kiểm tra trực tiếp ở khung Xem trước PDF.

### 6.2. Quy tắc đánh số trùng lặp (`.1`, `.2`, `.3`...)
Trong thư mục của cùng một cá nhân, thường có nhiều file PDF thuộc cùng một loại tài liệu (ví dụ: nhiều tờ *Phiếu bổ sung hồ sơ đảng viên* qua các năm, hoặc nhiều *Quyết định điều động bổ nhiệm*, nhiều *Bản tự kiểm điểm*):
- Nếu trong thư mục của người đó chỉ có **1 file** thuộc loại đó:
  - Đặt tên chuẩn: `<Tên tài liệu>.pdf`
  - Ví dụ: `05. Quyet dinh ket nap dang vien.pdf`
- Nếu trong thư mục của người đó có **từ 2 file trở lên** cùng loại:
  - Phần mềm sắp xếp theo thứ tự tự nhiên của tên file scan gốc và đánh số hậu tố `.1`, `.2`, `.3`...:
  - Ví dụ:
    - `04. Phieu bo sung ho so dang vien.1.pdf`
    - `04. Phieu bo sung ho so dang vien.2.pdf`
    - `04. Phieu bo sung ho so dang vien.3.pdf`

---

## 7. Danh sách các File Báo cáo Excel xuất tự động

Mỗi khi chạy **Quét & Dự kiến đổi tên (Dry-Run)** hoặc **Thực thi đổi tên thật**, phần mềm tự động xuất báo cáo vào thư mục báo cáo (mặc định `Documents/PhanMemDoiTenPDF_Reports`; có thể mở bằng nút **`📂 Mở thư mục Báo cáo Excel`** hoặc từ **Tab 3**):

1. 📗 **`Danh_sach_du_kien_doi_ten.xlsx`**:
   - Bảng tổng hợp toàn bộ tất cả các file PDF đã quét gồm: *Phòng ban, Họ tên, Thư mục tương đối, Tên file gốc, Nội dung OCR nhận dạng, Tên Excel đối chiếu, Điểm tin cậy (%), Tên file dự kiến, Trùng lặp trong thư mục, Ghi chú/Lý do*.
2. 📕 **`Danh_sach_can_kiem_tra.xlsx`**:
   - Trích lọc riêng các file thuộc trạng thái `Cần kiểm tra` (trang mờ, trang lẻ tách rời...) kèm lý do cụ thể để cán bộ số hóa dễ dàng rà soát.
3. 📘 **`Thong_ke.xlsx`**:
   - Gồm 2 Sheet:
     - `Thong_ke_chung`: Tổng số file PDF, tổng số nhận dạng được, số khớp Excel, số ngoài Excel, số trùng lặp (`.1, .2...`), số cần kiểm tra.
     - `Theo_phong_ban`: Thống kê chi tiết số lượng và kết quả theo từng Phòng ban / Đơn vị.
4. 📙 **`Nhat_ky_doi_ten.xlsx`** & **`nhat_ky_doi_ten.json`**:
   - Xuất hiện sau khi bấm **`✅ 2. Thực thi đổi tên thật`**, lưu chi tiết từng file đã đổi từ tên cũ sang tên mới kèm mốc thời gian chính xác để phục vụ đối soát và hoàn tác (`Rollback`).

---

## 8. Câu hỏi thường gặp & Xử lý sự cố (FAQ)

### Câu 1: Khi tôi tải lên file Excel danh mục mới, tôi có phải chờ phần mềm chạy OCR lại từ đầu không?
> **Trả lời:** **Không.** Nhờ cơ chế lưu đệm kết quả OCR trong file `ocr_cache.sqlite`, những file PDF nào không thay đổi kích thước và ngày sửa đổi sẽ được lấy ngay văn bản OCR từ bộ nhớ đệm để đối chiếu với file Excel mới. Do đó, hơn 16.000 file PDF chỉ mất khoảng **2 – 5 giây** để cập nhật toàn bộ tên mới.

### Câu 2: Làm sao để buộc phần mềm quét OCR lại từ đầu cho toàn bộ file?
> **Trả lời:** Nếu muốn quét OCR lại từ đầu, mở thư mục báo cáo, xóa file `ocr_cache.sqlite`, rồi bấm **`🔍 1. Quét & Dự kiến đổi tên (Dry-Run)`**.

### Câu 3: Khi đang mở file báo cáo Excel bằng Microsoft Excel mà bấm "Quét Dry-Run" thì báo lỗi Permission Denied?
> **Trả lời:** Microsoft Excel khóa file `.xlsx` khi đang mở. Bạn chỉ cần **đóng các file báo cáo Excel** (`Danh_sach_du_kien_doi_ten.xlsx`, `Thong_ke.xlsx`...) đang mở trong Microsoft Excel rồi bấm chạy lại.

### Câu 4: Các file PDF đã từng được đổi tên trước đó, nếu chạy lại lần nữa có bị đổi sai hoặc đánh số lặp `.1.1` không?
> **Trả lời:** **Không.** Bộ phân tích của phần mềm tự động nhận diện tiền tố mã số và loại bỏ hậu tố `.1, .2...` cũ khi phân tích, đồng thời thuật toán đổi tên 2 bước kiểm tra các file đã đúng tên chuẩn để giữ nguyên, đảm bảo tính **idempotent** (chạy nhiều lần kết quả vẫn chuẩn xác như nhau).

---

## 9. Đóng gói chạy offline

Yêu cầu Python 3.14 và các thư viện `PyMuPDF`, `Pillow`, `RapidFuzz`, `PyInstaller` trên máy build. Đặt bộ Tesseract Windows có `libtesseract-5.dll` và `tessdata\vie.traineddata` trong một thư mục, rồi cấu hình trước khi build:

```powershell
$env:TESSERACT_DIR = "C:\duong-dan-den\Tesseract"
powershell -ExecutionPolicy Bypass -File .\build_offline.ps1
```

Gửi nguyên thư mục `dist\PhanMemDoiTenPDF` sang máy đích. Máy đích không cần cài Python hoặc Internet. Chạy `PhanMemDoiTenPDF.exe`; để xử lý tài liệu, chọn thư mục PDF trên máy đó trong giao diện.
