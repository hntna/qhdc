# Thiết kế: Làm lại giao diện web theo thương hiệu Viettel (chuyên nghiệp hơn)

- **Ngày:** 2026-09-25
- **Ứng dụng:** QHĐC 2027-2028 — Tổng hợp & So khớp Masterlist (SPA nội bộ, chạy offline qua `server.py`)
- **Mục tiêu:** Làm lại toàn bộ giao diện để trông chuyên nghiệp, đồng bộ theo bộ nhận diện Viettel (tông đỏ chủ đạo), giữ nguyên 100% chức năng.

## 1. Bối cảnh & phạm vi

Ứng dụng hiện dùng design system "Modern SaaS Light Mode" tông indigo (chàm). Cấu trúc:
- `index.html` — giao diện chính, 4 tab: Kiểm tra lỗi (`tabValidation`), Xem trước (`tabPreview`), Phân cấp (`tabHierarchy`), Báo cáo Mảng & DV (`tabStrategy` với các sub-tab).
- `styles.css` — hệ thống giao diện chính (~70KB), gồm cả một tập tiện ích thay thế Tailwind.
- `report.css` — giao diện trang/khu báo cáo.
- `app.js`, `report.js`, `vtb_converter.js` — toàn bộ logic (KHÔNG được thay đổi hành vi).

**Phạm vi:** Làm lại **toàn bộ** giao diện.

**Cách tiếp cận: CSS-first.**
- Viết lại hệ thống thiết kế trong `styles.css` và `report.css` (tokens màu, typography, spacing, shadow, radius, và các component: header, upload, tab, bảng, nút, badge, modal, báo cáo).
- Chỉ chỉnh `index.html` ở những chỗ cần thiết: thêm/đổi class, tinh chỉnh cấu trúc header (tiêu đề 2 dòng), bỏ các `style=""` inline gán màu rời rạc trên nút.
- **Không thay đổi** JavaScript logic, id phần tử, data-attribute mà JS đang phụ thuộc (`data-tab`, `data-subtab`, các `id` như `btnOpenResultFile`, `badgeValidationCount`, v.v.). Chỉ được thêm class trình bày.

## 2. Hệ thống màu & thương hiệu (Viettel)

Cập nhật khối `:root` trong `styles.css`.

| Token | Giá trị | Ghi chú |
|---|---|---|
| `--primary` | `#EE0033` | Viettel Red |
| `--primary-hover` | `#C4002A` | đỏ đậm khi hover |
| `--primary-soft` / active bg | `#FFF1F3` | nền nhấn/active nhạt |
| `--primary-glow` | `rgba(238,0,51,0.18)` | focus ring / glow |
| `--bg-app` | `#F5F6F8` | nền app xám rất nhạt |
| `--bg-surface` / `--bg-card` | `#FFFFFF` | thẻ trắng |
| `--border-subtle` | `#E5E7EB` | viền tinh tế |
| `--text-main` | `#1A1A1A` | tiêu đề |
| `--text-secondary` | `#4B5563` | chữ thường |
| `--text-muted` | `#6B7280` / `--text-dim` `#9CA3AF` | phụ |

- **Màu ngữ nghĩa:** giữ success (xanh lá), warning (hổ phách), info (xanh dương). `danger` dùng chung sắc đỏ brand nhưng chọn sắc độ/nền (`--danger` đậm hơn primary hoặc dùng nền `--danger-bg` rõ ràng) để badge lỗi vẫn phân biệt được với nút primary.
- **Font:** giữ **Inter** (đã nạp offline-friendly qua Google Fonts, gần với Viettel Sans). Giữ JetBrains Mono cho cột số.
- **7 mảng nghiệp vụ:** **GIỮ NGUYÊN hệ 7 màu phân biệt** cho badge (VT/ML/CĐBR/CNTT/TĐ/CĐ/HT), chỉ **hạ độ bão hòa nhẹ** để hài hòa với tông đỏ chủ đạo (điều chỉnh các cặp token `--bg-vt`, `--b-vt`, ... cho bớt rực). Chữ trên badge vẫn trắng, đảm bảo tương phản đạt WCAG AA.

## 3. Component

### 3.1 Header
- Dải nền **trắng tinh**, viền dưới mảnh (`--border-subtle`), bóng đổ rất nhẹ khi cuộn (sticky). Bỏ glassmorphism tím.
- Khối logo vuông bo góc, nền đỏ Viettel, icon trắng, bên trái.
- Tiêu đề **2 dòng**: dòng trên "QHĐC 2027–2028" (đậm), dòng dưới "Tổng hợp & So khớp Masterlist" (nhỏ, xám). Thay cho 1 dòng in hoa dài hiện tại.
- **Nút hành động chuẩn hóa:** đúng 1 nút primary đỏ đặc (File tổng hợp / hành động chính), còn lại kiểu `outline`/`ghost` xám nhất quán. **Bỏ toàn bộ `style=""` inline gán màu** trên các nút trong header. Kích thước, padding, icon, khoảng cách đồng đều.

### 3.2 Thanh Tab
- Đổi sang kiểu **underline tab**: gạch chân đỏ dưới tab đang chọn, không dùng nền khối. Tab thường màu xám, hover đổi màu chữ.
- Badge số liệu bo tròn, cỡ đồng đều, dùng màu ngữ nghĩa (đỏ = lỗi, xám/xanh = thông tin). Giữ nguyên các `id` badge.
- Áp dụng cùng nguyên tắc cho `report-subtab-btn`.

### 3.3 Thẻ & khu nạp file (7 mảng)
- Ô upload bo góc đều, viền đứt nhạt khi trống; khi có file: viền đặc + nền nhạt theo màu mảng. Trạng thái hover/drag rõ (đổi viền/nền, con trỏ).
- Chuẩn hóa header của mọi card: icon + tiêu đề + mô tả phụ, khoảng cách đồng đều theo lưới spacing.

### 3.4 Bảng dữ liệu
- Header bảng nền xám nhạt, chữ đậm, **sticky** khi cuộn.
- Zebra rất nhẹ, hover đổi nền dòng. Viền mảnh.
- Cột số **căn phải**, dùng font mono. Hàng subtotal/tổng nhấn nền đậm hơn + chữ đậm.

### 3.5 Nút, badge, modal
- Bộ nút thống nhất: biến thể `primary` (đỏ) / `outline` / `ghost`; kích thước `sm` / `md`. Focus ring dùng `--primary-glow`.
- Modal: bo góc lớn (`--radius-lg`/`xl`), bóng đổ sâu, overlay tối nhẹ, header/footer nhất quán.

### 3.6 Báo cáo (`report.css`)
- Đồng bộ cùng bộ token màu, spacing, radius, shadow với `styles.css`. Bảng báo cáo áp dụng cùng quy tắc ở 3.4.

## 4. Nguyên tắc xuyên suốt
- Lưới spacing bội số **4px**; tăng khoảng trắng, giảm viền/màu rối.
- Phân cấp typography rõ (cỡ, đậm, màu).
- Tương phản màu đạt **WCAG AA** cho chữ/nền, đặc biệt badge 7 mảng và nút primary.
- Bo góc & shadow nhất quán qua biến `--radius-*`, `--shadow-*`.

## 5. Ràng buộc & tiêu chí hoàn thành
- **Không đổi hành vi JS:** giữ nguyên mọi `id`, `data-tab`, `data-subtab`, cấu trúc mà JS truy vấn. Chỉ thêm class trình bày và chỉnh markup trình bày thuần túy.
- Ứng dụng vẫn chạy offline (không thêm phụ thuộc mạng ngoài bắt buộc; font giữ như hiện trạng).
- Tất cả 4 tab + sub-tab báo cáo + modal + công cụ chuyển đổi VTB hiển thị đúng, không vỡ layout.
- Kiểm thử thủ công: nạp file mẫu, chuyển qua đủ 4 tab, mở modal, xuất báo cáo — giao diện đồng bộ tông Viettel, không lỗi bố cục.

## 6. Ngoài phạm vi (YAGNI)
- Không thêm tính năng mới, không dark mode, không đổi luồng nghiệp vụ.
- Không refactor logic JS ngoài việc gỡ các `style=""` inline liên quan trình bày.
- Không đổi font sang font proprietary Viettel Sans (không có sẵn offline).
