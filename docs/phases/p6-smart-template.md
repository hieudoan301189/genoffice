# P6 – Smart Template + Batch Generator (mốc M2)

Ngày: 02/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md),
mục P6. Kiểm bằng vitest trên môi trường cloud (Linux, không có Word/Excel hay bản đóng gói).

## Kết luận

- **Package mới `@genoffice/dvh-template`**: engine mẫu chạy thẳng trên OOXML, không cần editor. Docs (main
  process), action, MCP và CLI dùng chung được.
- **Mẫu thông minh** là một docx dùng các content control:

  | Tag                                      | Cấp  | Ý nghĩa                                                           |
  | ---------------------------------------- | ---- | ----------------------------------------------------------------- |
  | `dvh:f:<fieldId>`                        | run  | Smart Field (P1). Giữ control và binding; phần model cùng giá trị |
  | `dvh:c:<columnId>`                       | run  | cột của bản ghi: dòng đang lặp, hoặc bản ghi của lần sinh         |
  | `dvh:if:<id>`                            | khối | chỉ hiện khi điều kiện `<id>` đúng                                |
  | `dvh:repeat:<collectionId>[:<cond>]`     | khối | lặp khối cho từng bản ghi (lọc theo điều kiện)                    |
  | `dvh:repeatrows:<collectionId>[:<cond>]` | khối | bảng: lặp các dòng chứa `dvh:c:`; dòng trên/dưới giữ nguyên       |
  | `dvh:t:<tableId>`                        | khối | DVH.Table (P2), dựng lại từ collection, ra `<w:tbl>`              |
  | `dvh:img:<fieldId>`                      | run  | ảnh: thay ảnh bằng ảnh của Field (đường dẫn hoặc data URL)        |

  Biểu thức điều kiện lưu trong model (`conditions`, tuỳ chọn, nên hash của file cũ không đổi). Field có thêm
  kiểu `image`. Wrapper `if`/`repeat`/`repeatrows`/`c` bị bỏ khỏi file sinh ra. Field, bảng, ảnh giữ control
  nên file sinh ra vẫn là tài liệu Smart; muốn bản sạch thì dùng chính sách phát hành của P5.

- **Biểu thức an toàn**: mở rộng phép tính số của `dvh-evaluate.ts`.
  - Có so sánh, `AND`/`OR`/`NOT`, nối chuỗi `&`, `[Tiêu đề cột]`, tên có chấm (`Project.Name`, `row.qty`),
    `INDEX`, `COUNT`.
  - Hàm: `IF`, `LEN`, `UPPER`, `LOWER`, `TRIM`, `LEFT`, `RIGHT`, `MID`, `CONTAINS`, `ISBLANK`, `ROUND`, `ABS`,
    `MIN`, `MAX`, `SUM`, `CONCAT`, `TEXT`, `PAD`.
  - Phân tích thành cây rồi thông dịch: không `eval`, giới hạn độ dài và số bước, tên lạ là lỗi (không âm
    thầm thành rỗng).
  - Quy tắc đặt tên file dùng `{biểu thức}`, ví dụ `BB-{PAD(INDEX,3)}-{row.code}`.
- **Batch Generator**: Dữ liệu → Mẫu → Lọc → Xem trước → Sinh → Đóng gói.
  - Mỗi bản ghi của collection ra một tài liệu (hoặc một tài liệu nếu không chọn collection).
  - Xem trước cho biết số file, tên file, tên trùng (tự thêm hậu tố), cảnh báo; khi xuất ra thư mục còn hỏi
    trước nếu sắp ghi đè.
  - Định dạng docx, hoặc PDF qua đường xuất không giao diện có sẵn (`--headless-export`).
  - Ra thư mục hoặc zip, kèm mục lục `_MucLuc.csv` (UTF-8 có BOM để Excel đọc đúng tiếng Việt).
  - **Kết quả xác định**: zip có thứ tự mục và mốc thời gian cố định, nên cùng mẫu và dữ liệu cho ra cùng
    từng byte.
- **Template Designer** (mục "Mẫu thông minh" trong panel Smart Data của Docs):
  - tạo/sửa điều kiện (kiểm cú pháp ngay);
  - bọc khối đang chọn trong điều kiện hoặc vùng lặp (chọn đúng một bảng thì lặp dòng);
  - chèn cột tại con trỏ;
  - xem trước một bản ghi (mở ở tab mới);
  - sinh hàng loạt (Kiểm tra rồi Sinh);
  - chuyển mẫu DVH-Tool; nhập/xuất QLCL.

  Docs giữ các control này khi mở–lưu:
  - control cấp khối qua `sdtShell` sẵn có;
  - control cấp run `dvh:c:`/`dvh:img:` nhờ mở rộng `dvhFieldSdtPr` của `docx-engine`.

- **Chuyển mẫu DVH-Tool**:
  - `<<Field>>` thành Smart Field, kể cả khi Word đã tách placeholder ra nhiều run; giữ định dạng của run đầu.
  - Vùng `BD_Bang[_Tên]` … `KT_Bang[_Tên]` thành vùng lặp theo collection `Tên`, các `<<Cột>>` thành control
    cột:
    - trong bảng: bỏ dòng đánh dấu, lặp các dòng ở giữa, giữ nguyên định dạng bảng của mẫu;
    - ngoài bảng: lặp các đoạn ở giữa.
  - Tạo (hoặc bổ sung) phần model kèm itemProps, relationship và content type.
  - Placeholder vắt qua tab, ngắt dòng hay ảnh thì giữ nguyên và báo lại.
- **Trao đổi với QLCL-DVH** (một chiều mỗi lần):
  - JSON `dvh-exchange` v1, JSON Schema xuất từ schema của `dvh-model`;
  - workbook theo bố cục QLCL: `ThongTin` (tên/giá trị → Field, bỏ dòng tiêu đề) và `Data` cùng các sheet
    khác có dòng tiêu đề (→ collection);
  - nhập lại vào cùng model thì giữ ID theo tên, nên tài liệu đã gắn không bị đứt.
- **Action**: `Document.Generate`, `File.ExportPDF`, `File.Package` (mức `external`, luôn hỏi xác nhận theo P4);
  `Template.SetCondition`, `Template.WrapSection`, `Template.InsertColumn` (mức `write`). Nhờ vậy P7 (AI) và P8
  (quy trình) dùng lại được ngay, và MCP `dvh_execute` gọi được.

## Tệp chính

- `packages/dvh-template/src`:
  - `expr.ts` (biểu thức), `ooxml.ts` (tiện ích XML), `table-ooxml.ts` (DVH.Table → `<w:tbl>`);
  - `fill.ts` (điền mẫu, ZIP xác định), `batch.ts` (kế hoạch, sinh, mục lục, đóng gói);
  - `convert.ts` (DVH-Tool), `exchange.ts` (QLCL), `release.ts` (chính sách phát hành của P5, chuyển từ Docs
    sang đây để dùng chung).
- `packages/dvh-model`: kiểu Field `image`, `conditions` (`dvhConditionSchema`).
- `packages/docx-engine/src/smart-field.ts`: giữ control run-level `dvh:c:`/`dvh:img:`.
- Docs:
  - main: `dvh-batch.ts` (chạy batch, PDF headless, nạp ảnh, đóng gói);
  - IPC `docs:dvh-batch-plan`, `-batch-run`, `-template-preview`, `-convert-template`, `-qlcl-import`,
    `-qlcl-export`, `-package`, `-export-pdf`;
  - renderer: `dvh-template-designer.ts`, `components/DvhTemplateSection.tsx`, `dvh-actions.ts` (6 action mới),
    `mcp-bridge.ts` (bytes của mẫu cho action); 32 chuỗi mới cho cả 20 ngôn ngữ.
- Phụ thuộc: `@genoffice/dvh-template` thêm vào `dependencies` và `exclude` của `externalizeDepsPlugin`
  (Docs), phụ thuộc của shell, script `test`/`typecheck` gốc, lockfile.

## Kiểm thử (vitest)

- `packages/dvh-template` (10):
  - biểu thức: giá trị, so sánh, logic, hàm, `IF` tính lười, quy tắc tên; không chạy được mã, báo lỗi rõ;
  - điền mẫu: Field (control và model part cùng giá trị), cột có định dạng số theo locale, điều kiện đúng/sai,
    lặp dòng bảng (dòng tiêu đề và cuối giữ nguyên), lặp khối có lọc, DVH.Table dựng lại có `<w:tblHeader/>`,
    ảnh thay qua relationship; `strip` không còn DVH;
  - **nghiệm thu M2**: 100 biên bản từ WorkItems; chạy lại ra **từng byte giống hệt**; gói zip giống hệt, 101
    mục kèm mục lục đúng dòng tiêu đề và 100 dòng dữ liệu; bản ghi thứ 100 mang đúng dữ liệu của nó;
  - batch: lọc, tên theo quy tắc, tên trùng, tên file hợp lệ trên Windows;
  - DVH-Tool: placeholder bị tách run, vùng lặp trong bảng và ngoài bảng, phần model được tạo, mẫu sau chuyển
    sinh được tài liệu;
  - QLCL: workbook đi–về, giữ ID khi nhập lại, xuất xác định; JSON có phiên bản và kiểm schema.
- `apps/docs/tests/dvh-template-designer.test.ts` (1): bọc vùng lặp, điều kiện, chèn cột trong editor → **đường
  lưu thật** của Docs ghi đủ control → engine điền ra hai bản ghi, điều kiện đúng, không còn wrapper; khối lồng
  bị từ chối.
- `apps/docs/tests/dvh-template-roundtrip.test.ts` (1): mẫu mở trong Docs giữ control cột ở run và vùng
  điều kiện/lặp ở khối (vùng lặp nhiều đoạn chung một vỏ).
- Bộ đầy đủ:
  - Docs 3.143/3.143; `docx-engine` 1.553 (1 bỏ qua); `dvh-template` 10;
  - typecheck dvh-model/dvh-template/docx-engine/docs/sheets/shell sạch; theme-colors, english-comments,
    eslint, prettier đạt.
- **Bất biến giao diện**: engine không nhận tham số giao diện nào; file sinh ra chỉ phụ thuộc mẫu, dữ liệu và
  locale định dạng số. Kiểm "chạy lại giống từng byte" bao luôn yêu cầu này.

## Để lại cho máy Windows (chưa chạy)

1. Mở 5 file trong lô 100 biên bản bằng **Word**: không báo sửa file; Field hiện đúng giá trị; bảng lặp và bảng
   DVH đúng.
2. Batch định dạng **PDF** (gọi `--headless-export` của chính ứng dụng) ra thư mục và ra zip.
3. Một mẫu DVH-Tool thật của QLCL-DVH: chuyển đổi, đối chiếu báo cáo (số Field, vùng bảng, cảnh báo), sinh thử.
4. Workbook QLCL thật: nhập vào tài liệu, sửa, xuất ngược; QLCL-DVH mở được file xuất.
5. Đổi giao diện sáng/tối rồi sinh lại lô: so sánh từng byte với lần trước.

## Giới hạn còn lại

- Chưa hỗ trợ vùng lặp lồng nhau. Điều kiện chỉ ở cấp khối (không có điều kiện giữa dòng chữ).
- Chưa có giao diện gắn ảnh vào Field (`dvh:img`): mẫu soạn trong Word, hoặc gắn qua chuyển đổi/AI.
- Chưa có kéo thả trong Designer: chèn qua nút và action.
- Bảng DVH sinh ra rộng theo trang A4 lề 2 cm.
- Trao đổi QLCL mới một chiều mỗi lần; đồng bộ hai chiều thuộc P10.
