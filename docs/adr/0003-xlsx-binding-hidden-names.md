# D3 – Binding xlsx bằng defined name ẩn `_dvh.*`

- **Trạng thái:** Đã chấp nhận (02/10/2026)
- **Bằng chứng:** S1, S3 – docs/spikes/s3-hidden-names.md

## Bối cảnh

Ô gắn dữ liệu phải đi theo khi người dùng chèn/xoá/di chuyển hàng cột, trong cả Excel lẫn DVH Office.

## Quyết định

Field gắn ô bằng defined name ẩn `_dvh.f.<fieldId>`; vùng bảng bằng `_dvh.t.<tableId>`. Trình soạn thảo Sheets nạp các tên này như "tên hệ thống" (không hiện trong Name Manager) để Univer dời chúng khi cắt/dán và chèn/xoá; khi lưu ghi chúng từ mô hình kèm `hidden="1"`.

## Hệ quả

Excel giữ và dời tên ẩn. Hiện trạng trước P1: Sheets không nạp tên ẩn → cắt/dán không mang tên theo, xoá ô đang gắn làm lưu thất bại, sheet có binding không xoá được. P1 phải xử lý xoá ô gắn ngay lúc thao tác (hỏi người dùng / gỡ binding).
