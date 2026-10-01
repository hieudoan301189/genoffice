# D8 – Bộ điều phối liên kết ở main process của shell

- **Trạng thái:** Đã chấp nhận (02/10/2026)
- **Bằng chứng:** Kế hoạch mục 5 P3

## Bối cảnh

Chỉ shell thấy mọi tab đang mở; kênh điều khiển và MCP bridge đã nằm ở đó.

## Quyết định

`apps/shell/src/main/dvh-link-host.ts` biết tài liệu nào mở ở tab nào và chuyển ChangeSet giữa chúng; nguồn đang đóng được đọc ở main qua gateway/sidecar.

## Hệ quả

Code main của app được biên dịch vào shell: sửa xong phải build lại shell.
