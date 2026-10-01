# D2 – Smart Field trong docx là content control cấp dòng có w:dataBinding

- **Trạng thái:** Đã chấp nhận (02/10/2026)
- **Bằng chứng:** S1, S2 – docs/spikes/s1-office-compat.md, docs/spikes/s2-docs-smart-field.md

## Bối cảnh

Trường dữ liệu trong tài liệu phải hiện đúng giá trị cả khi mở bằng Word, và sống qua chỉnh sửa trong DVH Office.

## Quyết định

Smart Field là `w:sdt` cấp dòng với `w:tag="dvh:f:<fieldId>"` và `w:dataBinding` (XPath tới phần tử field trong phần model, `w:storeItemID` = GUID của phần đó). Bảng là `w:sdt` cấp khối `dvh:t:<tableId>`. docx-engine giữ `w:sdtPr` gốc trên các run (`sdtFieldXml`); Docs mang nó bằng mark `dvhField`.

## Hệ quả

Word tự cập nhật nội dung control từ store và ghi ngược khi người dùng sửa trong Word. DVH Office phải ghi ngược chữ sửa trong trường vào store (P1), nếu không Word sẽ ghi đè khi mở. Chép/dán trường cần cấp `w:id` mới.
