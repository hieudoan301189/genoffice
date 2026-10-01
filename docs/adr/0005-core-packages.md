# D5 – Lõi DVH trong package riêng, móc mỏng vào ứng dụng

- **Trạng thái:** Đã chấp nhận (02/10/2026)
- **Bằng chứng:** Kế hoạch mục 2.3, CLAUDE.md

## Bối cảnh

Repo là fork của genoffice và đồng bộ upstream định kỳ; các file lớn của upstream (6–7 nghìn dòng) đổi thường xuyên.

## Quyết định

Lõi thuần TypeScript không phụ thuộc Electron/Univer/TipTap: `packages/dvh-model`, `packages/dvh-actions` (sau là `dvh-workflow`, `dvh-script`). Điểm móc vào ứng dụng nằm trong file `dvh-*.ts` riêng; file upstream chỉ thêm dòng đăng ký.

## Hệ quả

Package mới phải có trong `dependencies` của app và trong `exclude` của `externalizeDepsPlugin` (Docs, Sheets main). Kiểm thử bằng vitest không cần giao diện.
