# D10 – Smart Data là lớp bọc chung

- **Trạng thái:** Đã chấp nhận (02/10/2026)
- **Bằng chứng:** Kế hoạch mục 12

## Bối cảnh

Chủ dự án xác định Smart Data dùng cho mọi cấu trúc, không cố định cho QLCL (01/10/2026).

## Quyết định

Lõi chỉ biết Field, Record, Collection, Table và schema do người dùng hoặc gói schema định nghĩa. QLCL là một gói schema cộng adapter trao đổi dữ liệu với QLCL-DVH, nằm ngoài lõi.

## Hệ quả

Kiểm thử lõi phải chạy với ít nhất hai bộ schema khác nhau. Ví dụ `Project.Name`, `WorkItems` chỉ là minh hoạ.
