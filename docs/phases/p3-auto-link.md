# P3 (phần 1) – Liên kết tự động

Ngày: 02/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md),
mục P3. Kiểm bằng vitest, trên bản đóng gói `apps/shell/release/dvh-p3a`, bằng Excel và Word 365.

## Kết luận

**Liên kết "Tự động" chạy được theo hai đường:**

1. **Workbook đang mở trong DVH Sheets:** mỗi lần sửa (gom 0,5 giây), Sheets phát dữ liệu Smart Data đọc từ
   ô qua main process. Tài liệu đang mở cập nhật Field và bảng ngay, **không cần lưu workbook**.
2. **Workbook đóng, hoặc sửa bằng Microsoft Excel:** tài liệu theo dõi file nguồn. Lưu xong là cập nhật, đọc
   thẳng từ các ô mà tên ẩn `_dvh.*` trỏ tới, vì Excel không cập nhật phần model DVH trong file.

Cập nhật tự động không ghi vào lịch sử Undo của người đang gõ, và không ghi đè bảng đã sửa tay (liên kết
báo "đã sửa tay" cho tới khi Refresh bằng tay).

| Kiểm tra (e2e `scripts/p3/auto-link.mjs`)                                     | Kết quả                                                                                                                            |
| ----------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| Sửa trong Sheets, chưa lưu: số, thêm dòng, màu nền tiêu đề, Field             | ✅ Docs: `120.50 → 130.00`, thêm dòng AB.4, nền tiêu đề `FFF2CC → C6EFCE`, "Dự án A → Dự án B"; workbook vẫn ở trạng thái chưa lưu |
| Excel thật sửa và lưu file: đổi chữ, thêm dòng ngoài vùng tên cũ, màu tiêu đề | ✅ Docs: "Bê tông lót móng", thêm dòng AB.5 (vùng tự giãn như trong Sheets), nền `BDD7EE`                                          |
| Word mở tài liệu đã lưu                                                       | ✅ Field `dvh:f` vẫn gắn dữ liệu; bảng trong control `dvh:t`, 6 dòng, hàng tiêu đề lặp, màu tiêu đề đúng                           |
| Lỗi trang                                                                     | 0                                                                                                                                  |

## Đã làm

- **Bảng theo cột nguồn** (`autoColumns`, bật cho bảng mới):
  - cột thêm/bớt/đổi tên ở nguồn hiện ra khi Refresh;
  - thiết lập riêng của cột cũ vẫn giữ (ID cột khớp theo tiêu đề, hoặc theo vị trí nếu đổi tên tại chỗ);
  - Sheets dịch ô theo chiều ngang trong các hàng của bảng;
  - collection tự giãn theo tiêu đề gõ thêm bên phải.

  Đổi định dạng nguồn (màu, viền, `numFmt`) cũng được tính là thay đổi và truyền đi.

- **Đọc từ ô** (`apps/docs/src/main/dvh-xlsx-cells.ts`):
  - giá trị, chuỗi dùng chung (cả rich text), công thức đã tính;
  - style: đậm/nghiêng, màu theo RGB hoặc theme, cỡ chữ, nền, viền, căn lề;
  - định dạng số (tự định nghĩa và built-in), độ rộng cột.

  Áp cùng quy tắc tự giãn/co như trong Sheets.

- **Main process:**
  - `dvh-link-watch.ts` theo dõi thư mục chứa workbook (nhận được kiểu lưu bằng đổi tên file của Excel và
    DVH), chờ file yên 0,7 giây và so `size+mtime` trước khi báo;
  - kênh `dvh:live-publish` → `docs:dvh-live` kiểm schema trước khi chuyển tiếp (`dvhLivePayloadSchema`
    trong `dvh-model`).
- **Docs:**
  - `dvh-auto.ts` chạy cả khi panel đóng: khi mở tài liệu thì kéo các link "khi mở/tự động" và đánh dấu link
    thủ công đã cũ;
  - nhận thông báo file đổi và kênh trực tiếp, xử lý tuần tự;
  - bỏ qua nếu nguồn không đổi (so `modelHash`).
- **Link Manager** trong panel:
  - chế độ Thủ công / Khi mở / Tự động;
  - trạng thái (cũ / mất nguồn / bảng đã sửa tay);
  - Cập nhật;
  - **Ngắt liên kết**: Field và bảng giữ nguyên như nội dung thường, bỏ control, xoá khỏi model.
- **Sheets:** `dvh-live.ts` phát ảnh chụp chỉ đọc (không ghi lịch sử, không đổi tên), bỏ qua khi workbook còn
  đang nạp dở.

## Kiểm thử

- `apps/docs/tests/dvh-auto.test.ts` (3):
  - kênh trực tiếp cập nhật mà độ sâu Undo không đổi, bỏ qua workbook khác;
  - file đổi → cập nhật, bảng sửa tay bị bỏ qua và báo `edited`;
  - khi mở / link thủ công báo cũ;
  - ngắt liên kết.
- `apps/docs/tests/dvh-xlsx-cells.test.ts` (2): workbook kiểu Excel có vùng đã giãn, màu theme, `numFmt` tự
  định nghĩa, công thức.
- `packages/dvh-model/tests/table.test.ts` (+1): `autoColumns`.
- Bộ test Docs đầy đủ: 3.121/3.122 (ca còn lại vẫn là `protect-dialog` nhạy thời gian). `dvh-model` 21, test DVH
  của Sheets 26 đều đạt.

## Còn lại trong P3

- **Ghi ngược Field đơn từ Docs về nguồn** (Read/Write, kèm phát hiện xung đột theo revision).
- **Tìm nguồn nâng cao:** theo `fileMap` của project và chỉ mục `docId` của shell; nút "Đổi nguồn" / "Mở nguồn".
- Một lần Refresh ở Sheets vẫn cần vài bước Undo (P4).
