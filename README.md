# QHĐC 2027-2028 - Ứng Dụng Tổng Hợp & So Khớp Masterlist

Ứng dụng web chạy nội bộ hỗ trợ tổng hợp, kiểm tra số liệu, chuẩn hóa phân cấp cây danh mục đầu tư và so sánh với kịch bản Chiến lược 5 năm theo chuẩn Masterlist 2027-2028.

## 🚀 Hướng dẫn khởi chạy nhanh (1 chạm)

### Cách 1: Chạy bằng file Batch (Khuyến nghị trên Windows)
1. Nhấp đúp chuột vào file **`start_app.bat`**.
2. Hệ thống sẽ tự động khởi động server nội bộ và mở giao diện ứng dụng trên trình duyệt (`http://localhost:8080`).

### Cách 2: Khởi chạy thủ công bằng dòng lệnh
```bash
python server.py
```
Sau đó mở trình duyệt và truy cập: [http://localhost:8080](http://localhost:8080)

---

## 📂 Cấu trúc thư mục chính
- **`index.html`**: Giao diện chính của ứng dụng (4 Tab làm việc chuyên biệt).
- **`app.js`**: Toàn bộ logic xử lý dữ liệu, kiểm tra lỗi, tính toán Subtotal tự động và đồng bộ Profile.
- **`styles.css`**: Hệ thống giao diện Modern SaaS Light Mode.
- **`server.py`**: Server Python nhẹ (dùng thư viện chuẩn, không cần `pip install`).
- **`strategy_profiles.json`**: File lưu trữ các kịch bản số liệu Chiến lược 5 năm.
- **`Masterlist 2027-2028_Mau.xlsx`**: File phôi mẫu Masterlist chuẩn (ghi đè kết quả trực tiếp).
- **`xlsx.full.min.js`, `exceljs.min.js`, `jszip.min.js`**: Các thư viện xử lý bảng tính offline.

---

## 🔒 Tính bảo mật & Riêng tư
- Ứng dụng xử lý dữ liệu **100% trong bộ nhớ Client (RAM trình duyệt)**, không gửi file hay dữ liệu ra bất kỳ máy chủ Internet nào bên ngoài.
