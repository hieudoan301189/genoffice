# P1 – Smart Data MVP + liên kết tối thiểu

Ngày: 02/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md),
mục P1. Kiểm bằng vitest, trên bản đóng gói `apps/shell/release/dvh-p1b`, và bằng Word/Excel 365.

## Kết luận

**Đạt các tiêu chí nghiệm thu chính.** Sheets gắn ô vào Field, Docs chèn Field đó dưới dạng content control
có `w:dataBinding`, và "Cập nhật từ nguồn" kéo giá trị mới sang sau khi ô gắn đã bị dời chỗ. Lưu, mở lại và
mở bằng Word/Excel đều giữ nguyên ID, binding và giá trị. Ba việc chuyển sang phase sau (xem mục _Còn lại_):

- collection `WorkItems`;
- thành phần bảng Smart Data dùng chung trong `packages/ui`;
- màn xem lịch sử ở Sheets.

| Tiêu chí nghiệm thu                                                                     | Kết quả                                                                                                                   |
| --------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------- |
| Lõi chạy với ít nhất hai bộ schema                                                      | ✅ hồ sơ nghiệm thu (`Project.Name`, `Contract.Value`) và bảng kê vật tư (`Material.Name`, `Material.Qty`) trong cùng e2e |
| _Move binding_: `Project.Name` từ B5 sang D10 (chèn 5 hàng, cắt dán), Docs vẫn cập nhật | ✅ tên ẩn thành `Sheet1!$D$10`; Docs báo nguồn đã đổi, cập nhật ra "Dự án B"                                              |
| Lưu và mở lại cả hai file, ID và binding không đổi                                      | ✅ cùng ID Field qua 4 lần mở/lưu (xlsx ×2, docx ×2)                                                                      |
| Word mở thấy giá trị trong content control                                              | ✅ 5/5 control `XMLMapping.IsMapped`, chữ khớp model                                                                      |
| Excel mở không báo lỗi                                                                  | ✅ mở chỉ đọc qua COM, D10 = "Dự án B", B11 = 1250000                                                                     |
| Đổi giao diện sáng/tối, file xuất giống hệt                                             | ✅ theo cấu trúc: đường lưu đọc JSON của ProseMirror và model, phần tô Field chỉ là CSS dùng token và bị gỡ khi in        |

## Đã làm

### Docs

- `apps/docs/src/renderer/dvh-smart-data.ts`:
  - nạp phần model và lịch sử theo namespace;
  - chèn Smart Field: text có mark `dvhField` mang `w:sdtPr` với `w:id` riêng;
  - **ghi ngược hai chiều**: chữ sửa trong trường thành giá trị model khi lưu (việc S2 để lại);
  - liên kết workbook (`{docId, relPath, path}`), "Cập nhật từ nguồn", phát hiện giá trị cũ bằng `modelHash`;
  - `dvhDocsCustomXmlParts` cấp hai phần cho `saveDocx`. Hàm này không tiêu thụ change set, vì bản sao phục hồi
    cũng dựng byte.
- `components/DvhSmartDataPanel.tsx`, nút **Insert → Smart Data**:
  - danh sách link, cảnh báo "workbook đã đổi" / "không tìm thấy";
  - bảng Field (giá trị, số lần dùng, Chèn), thêm Field riêng của tài liệu;
  - lịch sử 20 change set gần nhất với diff trước → sau.
- `ai/dvh-ops.ts`: op `insertSmartField` và `setSmartField` trong registry `ai/ops.ts`, nên agent, MCP
  (`apply_ops`) và CLI có luôn. Op ghi chữ vào transaction của batch. Model chỉ được ghi trong `commit()`
  (hook mới của `OpDef`, chạy sau khi cả batch áp xong), nên batch bị từ chối không chạm model. Trường tạo ở op
  trước dùng được ngay ở op sau trong cùng batch.
- `src/main/dvh-source.ts` + IPC `docs:dvh-read-source`: main process đọc phần model của xlsx thẳng từ đĩa
  (chỉ `.xlsx/.xlsm`, ≤ 256 MB), không cần mở workbook.
- i18n `strings-dvh.ts`, 20 ngôn ngữ, mỗi ngôn ngữ một shard.

### Sheets

- `setBoundFieldValue`: ghi ô gắn (một lần sửa ô bình thường, có Undo) và giá trị model.
- `record`/`createField` nhận nguồn change set (`ui`/`ai`).

### Action đầu tiên (`@genoffice/dvh-actions`)

| Docs (`dvh-actions.ts`)                  | Sheets (`dvh-actions.ts`)                       |
| ---------------------------------------- | ----------------------------------------------- |
| `Data.ListFields`, `Data.GetField`       | `Data.ListFields`, `Data.GetField`              |
| `Data.SetField` (qua op `setSmartField`) | `Data.SetField` (ghi ô gắn)                     |
| `Document.InsertField`                   | `Spreadsheet.BindField` (`B5` hoặc `Sheet1!B5`) |
| `Link.Update`                            |                                                 |

Phần lịch sử trong file do `dvh-smart-data` ghi. Change set mà `ActionRegistry.run` trả về là biên nhận cho
người gọi. P4 sẽ hợp nhất hai đường này. Với `GENOFFICE_DEBUG_HOOKS=1`, registry lộ ra tại `window.__dvhActions`
(Docs) và `window.__genofficeDebug.dvhActions` (Sheets) cho driver e2e. `dvh-actions` re-export `z`, nên app
không cần phụ thuộc zod riêng.

## Kiểm thử

- `apps/docs/tests/dvh-smart-data.test.ts` (8 ca):
  - nạp phần có sẵn;
  - chèn và ghi hai phần;
  - ghi ngược hai chiều, gọi lại không ghi trùng;
  - link cập nhật mọi chỗ dùng;
  - đường dẫn tương đối;
  - batch op `insertSmartField` + `setSmartField`;
  - batch bị từ chối không đổi model;
  - registry action.
- `apps/sheets/tests/dvh-actions.test.ts` (3 ca): catalog, phân tích tham chiếu ô, từ chối khi chưa mở workbook.
- Toàn bộ test Docs: 3.114/3.115 đạt. Ca còn lại là `protect-dialog` (băm mật khẩu, nhạy thời gian), không liên
  quan DVH: lỗi khi chạy cả bộ, chạy riêng đạt 8/8, giống S2. `dvh-model` 13, `dvh-actions` 6, test DVH của Sheets
  26 đều đạt. Typecheck Docs/Sheets sạch, eslint không lỗi, `check-theme-colors` không có màu thô mới.
- E2E `scripts/p1/docs-smart-data.mjs` trên bản đóng gói (0 lỗi trang ở cả 5 lần mở):
  1. Sheets: gắn 2 Field vào mỗi workbook (2 schema), lưu.
  2. Docs: link cả hai workbook (hộp chọn file được stub ở main process). Chèn 2 Field bằng panel, 2 Field
     bằng action registry, thêm `Report.Author` rồi gõ vào giữa trường, lưu. Kết quả: 5 SDT có binding, model
     có 5 giá trị, lịch sử gồm `Link.Create` ×2, `Document.InsertField` ×5, `Document.EditField`.
  3. Sheets: chèn 5 hàng, cắt B10 dán D10, `Data.SetField` → "Dự án B", lưu.
  4. Docs mở lại: link workbook nghiệm thu báo _stale_, workbook vật tư không. Cập nhật từ nguồn: trường thành
     "Dự án B", `lastSync.revision` = 1, lịch sử thêm `Link.Update`, lưu.
  5. Word: 5/5 control mapped, đoạn văn "Tên dự án: Dự án B". Excel mở workbook bình thường.
- `scripts/p1/sheets-smart-data.mjs` (bước Sheets trước đó) vẫn là bài kiểm thử cho cắt/dán, xoá ô gắn và mở lại.

## Còn lại / chuyển phase

- **Collection `WorkItems`** chuyển sang P2: collection chính là nguồn của DVH.Table, làm cùng bảng động.
- **Bảng Smart Data dùng chung trong `packages/ui`**: hai panel hiện khác nhu cầu (Sheets gắn ô, Docs có link
  và lịch sử). Gom lại khi P2 thêm bảng.
- **Lịch sử ở Sheets** mới được ghi vào file, chưa có màn xem. Panel lịch sử của Docs dùng lại được.
- Liên kết vẫn **một chiều, cập nhật thủ công** (đúng phạm vi P1). Hai chiều cho Field đơn là P3.
- Giá trị số trong Docs hiện ở dạng thô (`1250000`). Định dạng số theo `numFmt` của ô nguồn để sang P2/T1.
- Main process đọc bất kỳ `.xlsx` nào theo đường dẫn renderer gửi, nhưng chỉ trả phần model. Nếu cần chặt hơn, P3
  có thể giới hạn theo các link có trong tài liệu đang mở.
