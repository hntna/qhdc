# Thiết kế: Làm lại giao diện & kiến trúc điều hướng theo thương hiệu Viettel

- **Ngày:** 2026-09-25
- **Ứng dụng:** QHĐC 2027-2028 — Tổng hợp & So khớp Masterlist (SPA nội bộ, chạy offline qua `server.py`)
- **Mục tiêu:** Làm lại toàn bộ giao diện + tổ chức lại kiến trúc điều hướng (IA) cho chuyên nghiệp, đồng bộ bộ nhận diện Viettel (tông đỏ chủ đạo), **giữ nguyên 100% chức năng và logic JS**.

## 1. Bối cảnh & phạm vi

Ứng dụng hiện dùng design system "Modern SaaS Light Mode" tông indigo, điều hướng bằng **4 tab ngang**, và nhiều công cụ **nằm rải rác** (nút chuyển đổi VTB trên header; Fix group / Đánh lại chỉ mục / Tách Ví nằm trong tab Cấu hình nhóm). Cấu trúc file:
- `index.html` — giao diện chính. Các `tab-content`: `tabValidation`, `tabPreview`, `tabHierarchy`, `tabStrategy` (có sub-tab: `subtabTongHop`, `subtabMang`, `subtabDichVu`, `subtabStratDetail`).
- `styles.css` (~70KB, gồm tập tiện ích thay Tailwind) và `report.css` — giao diện.
- `app.js`, `report.js`, `vtb_converter.js` — toàn bộ logic. **KHÔNG thay đổi hành vi.**

**Phạm vi đợt này:** (a) tái cấu trúc điều hướng sang sidebar 6 khu; (b) reskin toàn bộ theo tông Viettel. **Không thêm tính năng nghiệp vụ mới** (trừ 1 placeholder "Sắp có", xem §3.5).

**Cách tiếp cận: tái dùng cấu trúc + CSS-first.**
- Giữ nguyên các khối `tab-content` và mọi `id`, `data-tab`, `data-subtab` mà JS đang truy vấn. Chỉ **thay lớp điều hướng** (sidebar bật/tắt đúng `tab-content` tương ứng) và **di chuyển/bọc lại markup** để gom công cụ.
- Viết lại design system trong `styles.css` + `report.css` (tokens, typography, spacing, shadow, radius, component).
- Gỡ các `style=""` inline gán màu rời rạc trên nút/badge.
- Không đổi JavaScript logic; nếu cần, chỉ thêm listener điều hướng sidebar (thuần trình bày, ánh xạ sang cơ chế tab hiện có).

## 2. Kiến trúc điều hướng (IA) — Sidebar 6 khu

Sidebar dọc bên trái (thu gọn được), liệt kê 6 khu; khu có nhiều trang thì dùng sub-tab ngang phía trên nội dung.

| # | Khu (sidebar) | Nội dung / ánh xạ hiện có |
|---|---|---|
| 🏠 | **Trang chủ** | Dashboard: bảng "Xem trước Masterlist Tổng hợp" (`tabPreview`) + hàng thẻ thống kê nhanh (số dòng, số lỗi, số nhóm, tiến độ 7 mảng) |
| 📥 | **Nạp dữ liệu** | Khu nạp 7 mảng nghiệp vụ + "Nạp file có sẵn" + "Bản lưu trên Server" (khối `uploadCard` hiện tại) |
| ✅ | **Kiểm tra dữ liệu** | `tabValidation` — bộ lọc phân loại lỗi + bảng kiểm tra + tìm kiếm + sao chép |
| 🧰 | **Công cụ** | Trang tập hợp các thẻ hành động (xem §3.5) |
| ⚙️ | **Cấu hình** | `tabHierarchy` — Cấu hình nhóm hạng mục & Subtotal |
| 📊 | **Tổng hợp** | `tabStrategy` với sub-tab: Tổng hợp chung · theo Mảng · theo Dịch vụ · So sánh CL 5 năm |

- Sidebar hiển thị logo Viettel + tên app ở đầu; mục đang chọn nhấn nền đỏ nhạt + gạch/thanh đỏ bên trái.
- Header trên cùng thu gọn lại: chừa nút hành động chính (Lưu lên Server, File tổng hợp, Làm mới) — không còn chứa nút công cụ.
- Trạng thái điều hướng ánh xạ 1-1 sang `data-tab`/`data-subtab` sẵn có; JS chuyển tab không đổi.

## 3. Hệ thống màu, thương hiệu & component

### 3.1 Màu & thương hiệu (Viettel) — cập nhật `:root`
| Token | Giá trị | Ghi chú |
|---|---|---|
| `--primary` | `#EE0033` | Viettel Red |
| `--primary-hover` | `#C4002A` | đỏ đậm khi hover |
| `--primary-soft` | `#FFF1F3` | nền active/nhấn nhạt |
| `--primary-glow` | `rgba(238,0,51,0.18)` | focus ring |
| `--bg-app` | `#F5F6F8` | nền app xám rất nhạt |
| `--bg-surface` / `--bg-card` | `#FFFFFF` | thẻ trắng |
| `--border-subtle` | `#E5E7EB` | viền tinh tế |
| `--text-main` | `#1A1A1A` | tiêu đề |
| `--text-secondary` | `#4B5563` | chữ thường |
| `--text-muted` / `--text-dim` | `#6B7280` / `#9CA3AF` | phụ |

- Màu ngữ nghĩa: giữ success (xanh lá), warning (hổ phách), info (xanh dương). `danger` chọn sắc/nền phân biệt rõ với `--primary` để badge lỗi không lẫn nút primary.
- Font: giữ **Inter** (gần Viettel Sans) + JetBrains Mono cho cột số.
- **7 mảng nghiệp vụ:** GIỮ hệ 7 màu badge (VT/ML/CĐBR/CNTT/TĐ/CĐ/HT), chỉ hạ độ bão hòa nhẹ; chữ trắng, tương phản đạt WCAG AA.

### 3.2 Sidebar
- Nền trắng, viền phải mảnh; mục = icon + nhãn; hover đổi nền; active = nền `--primary-soft` + thanh đỏ trái + chữ đỏ/đậm.
- Thu gọn được (chỉ icon) để nhường chỗ bảng rộng; trạng thái thu gọn lưu tùy chọn (localStorage, thuần trình bày).

### 3.3 Header
- Dải trắng, viền dưới mảnh, bóng nhẹ khi cuộn (sticky). Bỏ glassmorphism tím.
- Chỉ còn nhóm nút hành động chính, chuẩn hóa: 1 nút primary đỏ đặc, còn lại `outline`/`ghost` xám; kích thước/padding/icon đồng đều. **Bỏ mọi `style=""` inline gán màu.**

### 3.4 Trang chủ (Dashboard)
- Hàng thẻ thống kê (KPI cards): Số dòng gộp · Số lỗi kiểm tra · Số nhóm · Tiến độ 7 mảng (đã nạp/tổng). Số lấy từ các badge/đếm sẵn có (`badgeRowCount`, `badgeValidationCount`, `badgeGroupCount`, …), không tính toán mới.
- Bên dưới: bảng "Xem trước Masterlist Tổng hợp" (`tabPreview`) với toolbar sẵn có.

### 3.5 Trang Công cụ
Các thẻ hành động (card: icon + tên + mô tả ngắn + nút chạy), gom các chức năng sẵn có (giữ nguyên id/handler):
- Chuyển đổi form cũ (VTB) — `btnOpenVTBConverter`
- Fix group dòng — handler Fix Group hiện có
- Đánh lại chỉ mục — `btnReindexTT` / `reindexAllItems()`
- Tách Ví → DV riêng — `btnSplitWalletDV` / `applyWalletSplit('dv')`
- Tách Ví → Mảng riêng — `btnSplitWalletMang` / `applyWalletSplit('mang')`
- **Tách IP → Mảng riêng — thẻ "Sắp có" (disabled).** Chức năng MỚI, logic sẽ làm ở đợt sau; đợt này chỉ dựng chỗ giao diện.

### 3.6 Tab, thẻ & khu nạp file
- Thanh sub-tab (Tổng hợp) kiểu **underline** (gạch chân đỏ), badge bo tròn màu ngữ nghĩa, cỡ đồng đều; giữ `data-subtab` + id badge.
- Ô upload 7 mảng: bo góc đều, viền đứt khi trống; có file thì viền đặc + nền nhạt theo màu mảng; hover/drag rõ.
- Chuẩn hóa header mọi card: icon + tiêu đề + mô tả phụ, spacing đồng đều.

### 3.7 Bảng dữ liệu
- Header nền xám nhạt, chữ đậm, **sticky**; zebra nhẹ, hover đổi nền; viền mảnh; cột số căn phải + font mono; hàng subtotal/tổng nhấn đậm.

### 3.8 Nút, badge, modal & báo cáo
- Bộ nút thống nhất: `primary` (đỏ) / `outline` / `ghost`, cỡ `sm`/`md`, focus ring `--primary-glow`.
- Modal: bo góc lớn, bóng sâu, overlay tối nhẹ, header/footer nhất quán.
- `report.css`: đồng bộ token màu/spacing/radius/shadow; bảng báo cáo theo §3.7.

## 4. Nguyên tắc xuyên suốt
- Lưới spacing bội số 4px; tăng khoảng trắng, giảm viền/màu rối; phân cấp typography rõ.
- Tương phản đạt WCAG AA (đặc biệt badge 7 mảng, nút primary).
- Bo góc & shadow nhất quán qua biến `--radius-*`, `--shadow-*`.

## 5. Ràng buộc & tiêu chí hoàn thành
- **Không đổi hành vi JS:** giữ nguyên mọi `id`, `data-tab`, `data-subtab`, cấu trúc JS truy vấn; chỉ thêm class trình bày, markup trình bày, và listener điều hướng sidebar ánh xạ sang cơ chế tab hiện có.
- App vẫn chạy offline (không thêm phụ thuộc mạng ngoài bắt buộc).
- 6 khu sidebar + sub-tab Tổng hợp + modal + công cụ VTB hiển thị đúng, không vỡ layout ở màn hình phổ biến (≥1280px; xuống ~1024px sidebar thu gọn).
- Mọi nút/hành động cũ vẫn kích hoạt đúng handler cũ sau khi di chuyển.
- Kiểm thử thủ công: nạp file mẫu → duyệt đủ 6 khu → mở modal → chạy 1 công cụ → xuất báo cáo; giao diện đồng bộ tông Viettel, không lỗi bố cục hay JS.

## 6. Ngoài phạm vi (YAGNI)
- Logic thực của "Tách IP → Mảng riêng" (chỉ dựng placeholder "Sắp có").
- Dark mode; đổi font sang Viettel Sans proprietary; responsive điện thoại.
- Refactor logic JS ngoài listener điều hướng và gỡ `style=""` inline trình bày.
- Thêm bất kỳ tính năng nghiệp vụ mới nào khác.
