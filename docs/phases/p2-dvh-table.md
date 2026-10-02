# P2 – DVH.Table MVP

Ngày: 02/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md),
mục P2. Kiểm bằng vitest, trên bản đóng gói `apps/shell/release/dvh-p2d`, và bằng Word/Excel 365.

## Kết luận

**Đạt các tiêu chí nghiệm thu của P2 (mốc M1).** Một vùng có hàng tiêu đề trong Sheets trở thành collection
`WorkItems`. DVH.Table của collection đó hiển thị đúng cả dữ liệu lẫn style ở Sheets và Docs, theo kịp
khi nguồn thêm hoặc bớt dòng, giữ ID qua lưu/mở lại. Word mở ra thấy bảng thường nằm trong content
control, Excel mở workbook bình thường. 1.000 dòng render trong khoảng 0,26–0,33 giây.

| Tiêu chí nghiệm thu                                           | Kết quả                                                                                                                                                               |
| ------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| _Cross-render table_: đúng dữ liệu và style ở hai phía        | ✅ Sheets: tiêu đề vàng đậm, số đỏ `#,##0.00` giữ nguyên (kiểu nguồn). Docs: nền tiêu đề `FFF2CC` (kiểu nguồn) / `D9E2F3` (kiểu đích), số đã định dạng, `w:tblHeader` |
| Thêm/bớt dòng nguồn rồi Refresh, hai phía cập nhật đúng       | ✅ Sheets: 3 → 4 → 2 dòng, nội dung bên dưới bảng dịch theo, vùng nguồn đặt cạnh không bị xô. Docs: 3 → 5 dòng ở cả hai bảng sau "Cập nhật từ nguồn"                  |
| Lưu và mở lại không đổi ID hay link; Word mở bảng bình thường | ✅ tên `_dvh.c.*` / `_dvh.t.*` giữ nguyên; Refresh sau khi mở lại không báo nhầm "đã sửa tay"; Word: 2 control `dvh:t:*`, mỗi cái 1 bảng 6 dòng, hàng tiêu đề lặp lại |
| 1.000 dòng render trong ngưỡng S4 (< 0,5 giây)                | ✅ 1.000 × 10: 255–328 ms đến khi lệnh xong, 307–378 ms đến khi vẽ xong                                                                                               |

## Đã làm

### `dvh-model`

- `DvhCellStyle`: phần định dạng cả Sheets lẫn Docs vẽ được:
  - đậm, nghiêng, màu chữ, nền, căn lề, cỡ chữ;
  - viền mảnh hoặc vừa.

  Màu là dữ liệu tài liệu (`#RRGGBB`), không phải token giao diện.

- Cột của collection mang thêm `style`, `headerStyle` và `width` của nguồn (phục vụ Keep Source Style).
- `DvhTable`:
  - nguồn là `{collectionId}` hoặc dữ liệu nhúng;
  - cột (`columnId`, tiêu đề, `numFmt`, căn lề, độ rộng, dòng tổng `sum/count/avg/min/max`);
  - tiêu đề nhiều tầng (`headerGroups`), dòng tổng;
  - style `source` / `destination` (header/body/total, nền xen kẽ);
  - `layout {repeatHeader, keepRowsTogether, widths}`;
  - `lastRender` (băm nội dung đã vẽ, để phát hiện sửa tay).

  Model có `tables` (đọc được file P1 chưa có trường này).

- `renderTable`: dựng lưới ô có style, dùng chung cho cả hai phía. `formatCellValue` định dạng số theo
  `numFmt` (nhóm nghìn, số lẻ, %, chữ trong ngoặc kép, ngày d/m/y) cho Docs.
- Tên `_dvh.c.<id>` cho vùng nguồn của collection; `tableSdtPrXml` cho control `dvh:t:<id>`; `contentHash`.

### Sheets (`dvh-tables.ts`)

- **Collection từ vùng chọn:** hàng đầu là tiêu đề, kiểu cột suy từ dữ liệu, style tiêu đề/ô đầu, `numFmt` và
  độ rộng cột lấy từ nguồn. Khi lưu hoặc trước khi render, vùng tự giãn theo dòng gõ ngay dưới và co khi
  dòng cuối trống, giống bảng Excel.
- **Render/Refresh:**
  - giá trị và style của mọi ô ghi bằng một `setValues`;
  - ô tiêu đề nhóm được gộp;
  - kiểu nguồn chép độ rộng cột.

  Khi số dòng đổi, chỉ **dịch ô trong các cột của bảng** (`insertCells` / `deleteCells`), giống "Insert Table
  Rows" của Excel. Lần thử đầu chèn cả hàng sheet đã làm vỡ vùng nguồn đặt cạnh bảng; e2e bắt được và đã sửa.

- **Vùng do DVH quản lý:**
  - gõ vào bảng → cảnh báo một lần;
  - Refresh thấy nội dung khác lần vẽ trước → hỏi trước khi ghi đè (action cần `force`).
- Panel Smart Data có thêm phần **Collections** (tạo từ vùng chọn, chèn bảng tại ô chọn) và **Tables** (chọn
  kiểu style, Refresh).

### docx-engine và Docs

- Bảng nằm trong SDT `dvh:t:*` được đọc ra với `tableSdtPr`. Node `docTable` có thuộc tính `dvhTable`; khi
  lưu, bảng được sinh lại sẽ bọc lại đúng control đó (`wrapDvhTable`), còn bảng không đổi thì giữ nguyên
  từng byte.
- `dvh-tables.ts`:
  - dựng `TableModel` từ lưới bảng (ô có nền, viền, căn lề, chữ đậm/màu, hàng tiêu đề lặp, không ngắt
    dòng, rộng theo trang);
  - chèn sau đoạn có con trỏ;
  - Refresh có kiểm sửa tay.
- Liên kết workbook chép cả collection. "Cập nhật từ nguồn" cập nhật collection rồi refresh mọi bảng dùng
  nó; bảng đã sửa tay thì hỏi.
- Số định dạng theo **vùng của hệ điều hành** (`app.getSystemLocale()`), không theo ngôn ngữ giao diện. Trên
  máy phát triển Windows đặt `en-US` nên Docs ghi `2,450.25`, còn Excel đang tự đặt dấu phân cách riêng nên
  hiện `2.450,25`. Đọc thiết lập riêng của Excel để sau.
- Panel có thêm phần Collections (Chèn bảng) và Tables (kiểu style, Refresh). Bảng DVH có viền đánh dấu
  bằng token, gỡ khi in.

### Action

| Sheets                                   | Docs                                     |
| ---------------------------------------- | ---------------------------------------- |
| `Data.CreateCollection`                  | (collection đến qua link)                |
| `Table.Create`, `Table.Render`           | `Document.InsertTable`                   |
| `Table.Refresh` (`force` khi đã sửa tay) | `Table.Refresh` (`force` khi đã sửa tay) |
| `Table.SetStyle`                         | `Table.SetStyle`                         |
|                                          | `Link.Update` (giờ refresh cả bảng)      |

## Kiểm thử

- `packages/dvh-model/tests/table.test.ts` (7 ca, hai schema nghiệm thu / vật tư):
  - model đi vòng qua `customXml`;
  - tên hệ thống;
  - kiểu nguồn, kiểu đích có dải màu;
  - tiêu đề nhóm và dòng tổng;
  - thêm/bớt dòng;
  - định dạng số `vi-VN` / `en-US`.
- `packages/docx-engine/tests/smart-field.test.ts` (+2): đọc `tableSdtPr`; bảng không đổi giữ nguyên byte;
  `wrapDvhTable`.
- `apps/docs/tests/dvh-tables.test.ts` (2):
  - chèn bảng có định dạng, header lặp, lưu bọc control;
  - cập nhật từ nguồn, sửa tay cần `force`, kiểu đích.
- `apps/sheets/tests/dvh-actions.test.ts`: catalog 9 action.
- Bộ test đầy đủ:
  - docx-engine 1.553/1.553 đạt;
  - Docs 3.116/3.117 (ca còn lại là `protect-dialog`, nhạy thời gian, chạy riêng đạt, như P1);
  - `dvh-model` 20, test DVH của Sheets 10, test DVH của Docs 14 đều đạt.

  Typecheck 5 package sạch, eslint không lỗi, `check-theme-colors` không có màu thô mới.

- Hồi quy P1 (`scripts/p1/docs-smart-data.mjs`) chạy lại trên `dvh-p2d`: đạt, 0 lỗi trang.
- E2E trên bản đóng gói:
  - `scripts/p2/sheets-table.mjs`: collection, render, thêm/bớt dòng nguồn, sửa tay, kiểu đích, 1.000 dòng,
    lưu, mở lại.
  - `scripts/p2/docs-table.mjs`: Docs ← Sheets, hai bảng, thêm dòng nguồn, cập nhật, lưu.
  - Đọc lại bằng `scripts/p2/word-tables.ps1` (Word) và `excel-inspect.ps1` (Excel), chỉ dùng tiến trình do
    script tự mở.

## Còn lại / chuyển phase

- Thay đổi cột (thêm/bớt cột nguồn) mới được phản ánh khi tạo lại bảng; Refresh chưa chèn hoặc xoá cột ở
  Sheets.
- Ở Sheets, một lần Refresh có đổi số dòng gồm vài lệnh (dịch ô, ghi, gộp), nên Undo cần vài bước. Gom thành
  một mục Undo để P4 làm cùng saga.
- Chế độ style Mapped và Hybrid để sau P6, đúng kế hoạch.
- Chưa có op `apply_ops` cho bảng (`insertSmartTable`); agent dùng được qua action registry khi P4/P7 nối vào.
- Field đơn (P1) vẫn hiện số thô ở Docs vì Field chưa mang `numFmt`. Làm cùng T1.
- E2E Docs từng treo 2/4 lần, đều ở `app.close()` của driver:
  - một lần Ctrl+S rơi vào ô `select` của panel nên tài liệu chưa lưu, app hỏi lưu khi đóng;
  - một lần app không đóng sau khi lưu, chưa rõ vì sao.

  Kiểm riêng thì lưu xong phiên Sheets "sạch" (nút Lưu tắt) và đóng bình thường, nên đây không phải lỗi
  "lưu rồi vẫn dirty". Driver giờ click vào trang trước Ctrl+S, giới hạn 20 giây cho việc đóng rồi kill, và
  ghi lại nếu bị kẹt. Lần chạy cuối: không lần nào kẹt.
