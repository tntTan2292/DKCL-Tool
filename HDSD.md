# HƯỚNG DẪN SỬ DỤNG TIỆN ÍCH "DKCL BÁO CÁO BĐTP HUẾ"

Tiện ích mở rộng **DKCL Báo cáo BĐTP HUẾ** (Chrome Extension) là công cụ hỗ trợ cán bộ, nhân viên VNPost nhanh chóng lấy dữ liệu chỉ số chất lượng KPI Bản đồ Hộ nghèo (BĐTP HUẾ) trực tiếp từ hệ thống `dkcl.vnpost.vn`, tự động tính toán, tổng hợp và xuất báo cáo ra file Excel.

---

## 1. Giới thiệu chức năng chính
* **Tự động hóa lấy dữ liệu:** Thay vì phải tải thủ công từng phần, công cụ tự động gọi API (curl) để thu thập dữ liệu chỉ số chất lượng:
  * Phát thành công TMĐT & Phát thành công Truyền thống (F4.1).
  * Thu gom bưu gửi đi liên tỉnh (F1.2).
  * Chất lượng toàn trình bưu gửi nội tỉnh (F1.1).
  * Chất lượng phát bưu gửi liên tỉnh (F1.3).
* **Cấu hình Trọng số động:** Cho phép tùy chỉnh trọng số (%) tính điểm bình quân cho các dịch vụ (TMĐT, Truyền thống, Thu gom).
* **Xuất báo cáo đa dạng (Hỗ trợ 2 phiên bản):**
  * **Phiên bản V1:** Xuất báo cáo theo Tỉnh, theo Bưu cục và báo cáo so sánh biến động 2 kỳ cấp tỉnh.
  * **Phiên bản V2:** Xuất báo cáo nâng cao (mẫu biểu mới `filebaocaov2.xlsx`) theo Tỉnh, theo Bưu cục và so sánh kỳ toàn diện F4.1/F1.1/F1.2/F1.3.
* **Giao diện Sidebar tiện lợi:** Tích hợp trực tiếp vào thanh bên phải trình duyệt Chrome khi truy cập `dkcl.vnpost.vn`, không che khuất nội dung chính của trang web.

---

## 2. Hướng dẫn cài đặt công cụ (Developer Mode)

Vì đây là tiện ích hỗ trợ nội bộ, bạn cài đặt trực tiếp vào trình duyệt Chrome theo các bước sau:

1. **Tải và giải nén:** Tải thư mục mã nguồn tiện ích (`BDHN_DKCL`) về máy tính và giải nén.
2. **Mở quản lý tiện ích trên Chrome:** Mở trình duyệt Chrome, truy cập địa chỉ:
   ```text
   chrome://extensions/
   ```
3. **Bật Chế độ dành cho nhà phát triển:** Nhìn góc trên bên phải màn hình, gạt công tắc kích hoạt **Chế độ dành cho nhà phát triển** (Developer mode).
4. **Tải tiện ích lên:**
   * Nhấp chọn nút **Tải tiện ích đã giải nén** (Load unpacked) ở góc trên bên trái.
   * Trỏ tới thư mục chứa mã nguồn tiện ích đã giải nén ở **Bước 1** (thư mục chứa tệp `manifest.json`) và bấm **Select Folder**.
5. **Hoàn tất:** Tiện ích **DKCL Báo cáo BĐTP HUẾ** sẽ xuất hiện trong danh sách các tiện ích đã cài đặt.

---

## 3. Quy trình sử dụng chi tiết

### Bước 1: Đăng nhập hệ thống VNPost
* Truy cập trang web: [https://dkcl.vnpost.vn](https://dkcl.vnpost.vn)
* Thực hiện đăng nhập bằng tài khoản nội bộ của bạn. 
* *Lưu ý: Tiện ích sẽ sử dụng phiên đăng nhập (Cookie/Token) hiện tại của trình duyệt để gọi dữ liệu, vì vậy bắt buộc phải đăng nhập trước.*

### Bước 2: Kích hoạt Sidebar báo cáo
* Sau khi đăng nhập và truy cập trang `dkcl.vnpost.vn`, một nút tròn màu cam **"DKCL"** sẽ tự động xuất hiện ở mép phải màn hình.
* Bấm vào nút **"DKCL"** này để mở thanh công cụ bên phải (Sidebar).
* Nếu muốn Sidebar luôn luôn mở cố định trên màn hình mà không tự ẩn đi khi click ra ngoài, hãy bấm nút **📌 Ghim** ở góc trên cùng của Sidebar.

### Bước 3: Thiết lập các tham số báo cáo
1. **Chọn khoảng thời gian:**
   * **Đầu kỳ (Từ ngày - Đến ngày):** Khoảng thời gian chính muốn lấy dữ liệu báo cáo.
   * **Kỳ so sánh (Từ ngày - Đến ngày):** Khoảng thời gian kỳ trước hoặc kỳ cần so sánh (chỉ cần thiết khi bạn sử dụng chức năng "So sánh kỳ" hoặc "Báo cáo so sánh kỳ V2").
   * *Mẹo: Bạn có thể nhập trực tiếp định dạng ngày `dd/MM/yyyy` hoặc nhấp vào biểu tượng lịch 📅 để chọn.*
2. **Thiết lập Trọng số bình quân cột P:**
   * Điều chỉnh phần trăm (%) trọng số cho 3 mảng dịch vụ: **TMĐT**, **Truyền thống**, **Thu gom** (Tổng phải bằng 100%).
   * Giá trị mặc định gợi ý là: TMĐT: `40%` | Truyền thống: `30%` | Thu gom: `30%`.

### Bước 4: Xuất báo cáo Excel
Tùy theo nhu cầu công việc, lựa chọn nút chức năng tương ứng trong Sidebar:

#### Nhóm Báo cáo V1:
* **Báo cáo theo tỉnh:** Lấy số liệu KPI phân theo cấp Tỉnh chấp nhận, xuất ra tệp Excel mẫu tiêu chuẩn V1.
* **Báo cáo theo bưu cục:** Lấy số liệu chi tiết tới từng Bưu cục phát/chấp nhận trên toàn hệ thống.
* **So sánh kỳ (Tỉnh):** Tải dữ liệu của cả 2 kỳ (Đầu kỳ & Kỳ so sánh), tự động tính toán thứ hạng (xếp hạng 36 đơn vị) và tính toán chênh lệch tăng/giảm chất lượng.

#### Nhóm Báo cáo V2 (Mẫu biểu mới):
* **Báo cáo theo tỉnh V2:** Sử dụng biểu mẫu nâng cao `filebaocaov2.xlsx` tổng hợp toàn diện chỉ số chất lượng tỉnh.
* **Báo cáo theo bưu cục V2:** Lấy dữ liệu chi tiết cho bưu cục thuộc mã tỉnh 53.
* **Báo cáo so sánh kỳ V2:** Tải dữ liệu 2 kỳ và so sánh chi tiết các chỉ số F4.1, F1.1, F1.2 và F1.3 để đánh giá biến động chất lượng dịch vụ toàn trình.

### Bước 5: Tải file kết quả
* Trong quá trình chạy, mục **Kết quả sau khi click** ở dưới cùng sẽ cập nhật trạng thái chi tiết (ví dụ: *Đang gọi curl V2 F4.1...*, *Đang xử lý xuất Excel...*).
* Khi tiến trình hoàn tất, tệp báo cáo Excel (`.xlsx`) sẽ tự động được tải về thiết bị của bạn.
* Trạng thái kết quả sẽ chuyển thành màu xanh lá hiển thị **"Hoàn thành"** cùng với mô tả ngắn gọn về số lượng dòng dữ liệu đã tải được.

---

## 4. Các lỗi thường gặp và cách xử lý

| Hiện tượng | Nguyên nhân | Cách khắc phục |
| :--- | :--- | :--- |
| **Báo lỗi 401 hoặc Không tải được dữ liệu** | Phiên đăng nhập trên hệ thống `dkcl.vnpost.vn` đã hết hạn. | Nhấp F5 tải lại trang web `dkcl.vnpost.vn`, đăng nhập lại tài khoản và thực hiện lại thao tác xuất báo cáo trên Sidebar. |
| **Nút "DKCL" không hiển thị** | Tiện ích chưa được bật hoặc trang web chưa tải xong. | Kiểm tra lại trong `chrome://extensions/` xem tiện ích đã được kích hoạt (Enabled) chưa. Nhấn tải lại trang web. |
| **Báo cáo so sánh kỳ không hiển thị cột chênh lệch** | Chưa nhập đầy đủ thông tin ngày của cả hai ô "Đầu kỳ" và "Kỳ so sánh". | Kiểm tra lại cấu hình ngày ở cả 2 khu vực thời gian. |
| **Trọng số bình quân cột P không chính xác** | Tổng trọng số TMĐT + Truyền thống + Thu gom khác 100%. | Nhập lại các giá trị sao cho tổng 3 ô bằng chính xác 100. |

---
*Chúc các bạn làm việc hiệu quả với công cụ hỗ trợ báo cáo KPI DKCL BĐTP HUẾ!*
