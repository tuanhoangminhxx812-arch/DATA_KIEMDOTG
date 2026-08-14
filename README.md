# ⚡ Hệ Thống Kiểm Dò & Đối Chiếu Tài Khoản Trung Gian (ERP)
### 🏢 Điện lực Vũng Tàu — Phòng Tài chính Kế toán (TCKT)

Hệ thống được phát triển nhằm mục đích **tự động hóa hoàn toàn** quy trình kiểm tra, đối chiếu và tìm kiếm các giao dịch chênh lệch (lệch số liệu) phát sinh trên các **Tài khoản Trung gian** từ dữ liệu Sổ cái (GL) xuất từ hệ thống ERP. 

Bằng việc kết hợp các giải pháp công nghệ hiện đại và thuật toán đối khớp đa tầng thông minh, hệ thống giúp tiết kiệm hàng chục giờ làm việc thủ công của kiểm toán viên và kế toán viên, tăng độ chính xác và giảm thiểu rủi ro sai sót số liệu.

---

## 🌟 Các Tính Năng Cốt Lõi

### 1. 🔍 Nhận diện cấu trúc dữ liệu tự động (Auto-Mapping)
Hệ thống có khả năng phân tích mọi cấu trúc tệp Excel Sổ cái kế toán được tải lên. Tự động nhận diện tên Sheet dữ liệu gốc và thiết lập vị trí các cột quan trọng dựa trên từ khóa linh hoạt:
*   **Tài khoản** (`taikhoan`, `tk`,...)
*   **Người tạo / Người lập** (`nguoitao`, `nguoilap`,...)
*   **Số tiền Nợ** (`noquydoi`, `no`,...)
*   **Số tiền Có** (`coquydoi`, `co`,...)
*   **Số chứng từ / Số giao dịch** (`sochungtu`, `sogiaodich`, `soctu`,...)
*   **Diễn giải / Nội dung** (`noidung`, `diengiai`,...)
*   **Ngày giao dịch** (`ngaygd`, `ngaychungtu`,...)

### 2. 🧠 Thuật toán đối ứng đa tầng thông minh (Reconciliation Engine)
Bộ não của hệ thống nằm ở tệp `loc_lech_taikhoan.py` với các tầng đối khớp liên tục:
*   **Creator-Level Difference Pairing (Bù trừ cấp Người tạo):** Tìm các nhóm người tạo có tổng số tiền lệch đối xứng chéo nhau (đối khớp 1-1, 1-nhiều, nhiều-1) thông qua thuật toán quy hoạch động giải bài toán tổng tập hợp con (Subset Sum).
*   **Cross-Creator Similarity Pairing (Đối chiếu chi tiết giao dịch):** Ghép đôi các giao dịch Nợ - Có lẻ giữa các người tạo dựa trên độ tương đồng nội dung diễn giải (sử dụng thuật toán **Jaccard Similarity** và chuẩn hóa chuỗi tiếng Việt), so khớp mã chứng từ độc lập, và áp dụng các bộ luật nghiệp vụ (heuristics) chuyên biệt của ngành Điện lực:
    *   *Nhóm Bảo hiểm:* BHXH, BHTN, BHYT, KPCD.
    *   *Nhóm Nhiên liệu / Vật tư:* Xăng, dầu, điều chuyển nội bộ giữa các đơn vị (PCSG, PCTT, PCVT,...).
    *   *Nhóm dịch vụ công ích:* Điện, nước, viễn thông.
    *   *Nhóm Thu hộ / Bảo lãnh hợp đồng:* Đối chiếu số PE/GK.
*   **Cross-Account Reconciliation (Đối chiếu chéo tài khoản):** Đối soát các giao dịch chênh lệch tồn đọng khác tài khoản trung gian nhưng có chung bản chất giao dịch và khớp giá trị số tiền.

### 3. 📊 Dashboard Đối chiếu Động ngay trong file Excel (`DoiChieu_TKTG`)
Sau khi đối soát, tệp Excel kết quả trả về không chỉ là danh sách phẳng mà được thiết kế như một **Báo cáo chuyên nghiệp chuẩn Corporate**:
*   Sử dụng phông chữ tiêu chuẩn `Segoe UI`, lưới hiển thị rõ ràng, định dạng số tiền dễ nhìn (`#,##0`).
*   **Bảng A (Tổng hợp trạng thái):** Cập nhật thời gian thực tình trạng Khớp/Lệch của toàn bộ các tài khoản trung gian chính.
*   **Dropdown Động (Data Validation tại ô C5):** Cho phép kế toán viên nhấp chọn nhanh tài khoản trung gian cần kiểm tra.
*   **Bảng B (Động hoàn toàn):** Tự động truy vấn và liệt kê chính xác các Người tạo đang gây lệch tiền kèm số tiền lệch chi tiết tương ứng với tài khoản được chọn ở ô C5.
*   **Tô màu cảnh báo (Conditional Formatting):** Màu Đỏ cảnh báo chênh lệch nguy hiểm, màu Xanh lá báo hiệu tài khoản đã cân đối hoàn hảo.

### 4. 💻 Giao diện Web Streamlit cực kỳ hiện đại (`app_streamlit.py`)
Cung cấp một cổng tương tác Web trực quan, tối ưu trải nghiệm người dùng với các thẻ chỉ số KPI nổi bật:
*   Độ lệch tổng thể hệ thống, Số lượng tài khoản chưa cân đối, Tổng phát sinh lệch Nợ/Có.
*   Trực quan hóa bảng đối chiếu tài khoản tổng thể và chi tiết từng giao dịch chưa khớp.
*   **Export & Tải xuống Báo cáo Tùy biến:** Kế toán có thể chọn xuất nhanh báo cáo Excel chi tiết giao dịch bị lệch cho *một tài khoản cụ thể* hoặc lọc theo *một người tạo cụ thể* trực tiếp trên trình duyệt.

---

## 📂 Danh Sách Tài Khoản Trung Gian Hỗ Trợ
Hệ thống tập trung đối chiếu chéo các tài khoản trung gian trọng điểm:
*   **`33191000000`** đến **`33195000000`** (Tài khoản công nợ trung gian mua sắm, dịch vụ)
*   **`33196100000`** / **`33198000000`**
*   **`24190000000`** (Tài khoản trung gian chi phí XDCB dở dang)
*   **`15110000000`** / **`15190000000`** (Tài khoản trung gian mua hàng, vật tư đi đường)

---

## 🛠️ Cấu Trúc Thư Mục Dự Án

```bash
KiemDoTKTG/
├── app_streamlit.py          # ⚡ File chạy giao diện chính Streamlit Dashboard
├── app.py                    # 🌐 File chạy giao diện phụ cổng Flask Web
├── loc_lech_taikhoan.py      # 🧠 Backend - Thuật toán đối chiếu đa tầng & Tạo file Excel kết quả
├── custom_ui/                # 🎨 Mã nguồn giao diện nhúng tùy biến của Streamlit
├── templates/ & static/      # 📂 Giao diện HTML/CSS/JS cho cổng Flask Web
├── requirements.txt          # 📌 Danh sách các thư viện Python cần thiết
├── Chay_Streamlit.bat        # 🚀 Script kích hoạt nhanh ứng dụng Streamlit trên Windows
├── Chay_Ung_Dung.bat         # 🚀 Script kích hoạt nhanh ứng dụng Flask trên Windows
└── README.md                 # 📖 Hướng dẫn sử dụng hệ thống
```

---

## 🚀 Hướng Dẫn Khởi Chạy Ứng Dụng (Dành Cho Máy Thành Viên)

Hệ thống đã được đóng gói kèm các script tự động hóa cài đặt môi trường. Bạn chỉ cần thực hiện các bước đơn giản sau:

### Cách 1: Sử dụng giao diện Dashboard Streamlit (Khuyên dùng)
Giao diện này hiện đại hơn, trực quan hơn và hỗ trợ tương tác lọc sâu dữ liệu.

1.  Nhấp đúp chuột vào file **`Chay_Streamlit.bat`**.
2.  Hệ thống sẽ tự động kiểm tra Python trên máy bạn, cài đặt các thư viện cần thiết (`streamlit`, `pandas`, `openpyxl`) nếu chưa có.
3.  Trình duyệt web sẽ tự động mở giao diện Dashboard tại địa chỉ `http://localhost:8501`.
4.  Kéo thả tệp Sổ cái Excel cần đối soát vào cổng tải lên ➔ Bấm **Phân Tích** ➔ Theo dõi KPI số liệu và tải về báo cáo tương ứng.

### Cách 2: Sử dụng giao diện Web Flask
Giao diện máy chủ web truyền thống tối giản.

1.  Nhấp đúp chuột vào file **`Chay_Ung_Dung.bat`**.
2.  Hệ thống cài đặt các thư viện liên quan (`flask`, `pandas`, `openpyxl`, `werkzeug`).
3.  Trình duyệt web tự động kết nối cổng `http://127.0.0.1:5000` để bắt đầu làm việc.

---

## 📝 Lưu Ý Quan Trọng Khi Sử Dụng
*   **Tính an toàn dữ liệu (Backup):** Khi chạy phân tích, hệ thống sẽ tự động tạo một tệp sao lưu có hậu tố `_backup.xlsx` tại cùng thư mục chứa tệp Excel gốc của bạn để bảo toàn nguyên vẹn mọi công thức và định dạng ban đầu của bạn. Không được xóa tệp backup này khi quá trình đối soát đang diễn ra.
*   **Yêu cầu môi trường:** Khuyên dùng **Python 3.9 trở lên**. Đảm bảo máy tính của bạn đã được cấu hình biến môi trường Path cho `python` hoặc trình khởi chạy `py`.
*   **Định dạng dữ liệu:** Sổ cái Excel tải lên nên giữ nguyên cấu trúc xuất ra từ ERP để hệ thống nhận diện tiêu đề cột chính xác nhất. Nếu hệ thống báo lỗi cột, vui lòng kiểm tra lại xem các tiêu đề cột (Tài khoản, Nợ, Có, Diễn giải, Người tạo) có bị ẩn hoặc viết sai chính tả quá nhiều hay không.

---

### 📬 Thông Tin Hỗ Trợ Kỹ Thuật
Dự án được xây dựng và hỗ trợ bởi **Phòng Tài chính Kế toán & Đội ngũ Phát triển Ứng dụng AI - Điện lực Vũng Tàu**. Mọi thắc mắc về giải thuật hoặc yêu cầu cải tiến tính năng vui lòng gửi yêu cầu phản hồi trực tiếp trên kho mã nguồn nội bộ.
