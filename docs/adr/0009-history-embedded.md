# D9 – Lịch sử nhúng trong file

- **Trạng thái:** Đã chấp nhận (02/10/2026)
- **Bằng chứng:** S1 – docs/spikes/s1-office-compat.md

## Bối cảnh

Chủ dự án chọn nhúng lịch sử vào file (01/10/2026).

## Quyết định

Phần `customXml` riêng (`urn:dvh-office:history:1`) chứa ChangeSet dạng JSON mỗi dòng một bản ghi, không tự nén (gói zip đã nén). Có snapshot định kỳ, gộp ChangeSet cũ, ngưỡng kích thước, hash mô hình để phát hiện sửa bên ngoài, và lệnh "Xuất bản sạch". Khi Xuất/Gửi đi luôn hỏi trước.

## Hệ quả

Lịch sử 5 MB JSON chỉ thêm ~350–390 KB vào file và mở dưới 0,2 giây trong Word/Excel. Lịch sử có thể lộ giá trị đã xoá khi gửi file.
