# D7 – Mọi hành động ghi phát ChangeSet

- **Trạng thái:** Đã chấp nhận (02/10/2026)
- **Bằng chứng:** Kế hoạch mục 6.4

## Bối cảnh

History, Recorder và Link cần một nguồn sự thật chung về thay đổi.

## Quyết định

Mỗi action ghi phát một `ChangeSet` (`id`, `txId`, `docId`, `at`, `source`, `action`, `changes[{objectId, path, before, after}]`). History, Recorder, Link Manager đọc cùng dòng này.

## Hệ quả

History không phải gắn vào từng đường sửa; action nào không phát ChangeSet thì không có lịch sử và không ghi được workflow.
