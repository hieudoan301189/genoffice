# D6 – Action Core bọc các catalog op sẵn có

- **Trạng thái:** Đã chấp nhận (02/10/2026)
- **Bằng chứng:** Kế hoạch mục 1, mục 5 P4

## Bối cảnh

Docs (`ai/ops.ts`), Sheets (`WorkbookOperation` + `op-executor`) và Slides (`runTxn`) đã có thao tác chuẩn hoá: kiểm tra trước, một transaction, một Undo, chạy thử.

## Quyết định

`dvh-actions` định nghĩa `ActionDescriptor` và registry; các catalog sẵn có được bọc thành action, không viết lại. Action mới chỉ cho Smart Data, Table, Link, File, Workflow.

## Hệ quả

Tận dụng phần Undo/dry-run đã ổn định; transaction xuyên ứng dụng là saga theo checkpoint.
