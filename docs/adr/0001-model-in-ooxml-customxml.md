# D1 – Mô hình DVH lưu trong phần customXml của OOXML

- **Trạng thái:** Đã chấp nhận (02/10/2026)
- **Bằng chứng:** S1 – docs/spikes/s1-office-compat.md

## Bối cảnh

File phải mở và sửa đúng bằng Word/Excel (quyết định của chủ dự án 01/10/2026). Toàn bộ đường lưu của DVH Office là vá OOXML giữ nguyên byte; thêm một định dạng gói riêng sẽ tạo đường lưu thứ ba cho mọi ứng dụng.

## Quyết định

Mô hình DVH (Field, Collection, Table, Link) nằm trong một phần `customXml/itemN.xml` có gốc `<dvh:model xmlns:dvh="urn:dvh-office:model:1">`, kèm `itemProps` mang store id (GUID). Field là phần tử XML (để XPath của data binding trỏ tới); Table/Link là JSON trong CDATA. **Luôn tìm phần theo namespace**, không theo tên `itemN`.

## Hệ quả

Word, Excel và DVH Office giữ nguyên phần này. Excel đánh số lại `itemN` khi lưu. Mô hình logic vẫn độc lập với định dạng; gói `.dvh` chỉ cân nhắc ở P10.
