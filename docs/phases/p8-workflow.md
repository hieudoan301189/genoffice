# P8 – Workflow Engine + Record Anything

Ngày: 02/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md),
mục P8. Kiểm bằng vitest trên môi trường cloud (Linux, không có bản đóng gói).

## Kết luận

- **Package mới `@genoffice/dvh-workflow`**: model, engine, recorder, lưu trữ và các thao tác sửa khối. Dùng
  chung Action Core (P4) và bộ đánh giá biểu thức của P6.
- **Workflow Model** (`dvh-workflow` v1, kiểm bằng zod):
  - tham số có kiểu (`text`, `number`, `boolean`, `record`, `list`, `any`) và giá trị mặc định;
  - bước: gọi action (`args`, `doc` tuỳ chọn, `assign` để lưu kết quả), `set`, `if`/`elseif`/`else`, `forEach`
    (có `where`), `try`/`catch` (biến lỗi), `transaction`, `log`;
  - đối số là giá trị JSON cố định hoặc biểu thức. Tên trần (`records`, `record`) trả nguyên bản ghi/danh sách;
    tên có chấm đọc thuộc tính (`record.code`), `.Count` cho độ dài danh sách, `[Tiêu đề]` đọc bản ghi đang lặp,
    `INDEX` là số thứ tự vòng lặp;
  - bước không có id: địa chỉ là **đường dẫn** (`2`, `2.body.0`, `3.then1.0`, `3.else.1`). Nhờ vậy P9 in ra
    script rồi đọc lại vẫn bằng đúng model gốc.
- **Engine** (`runWorkflow`): mỗi action chạy qua saga participant của P4 (kiểm input, quyền, xác nhận theo chính
  sách, change-set mang `txId` của lượt chạy, nguồn `workflow`).
  - Cả lượt chạy là một giao dịch: lỗi không được bắt thì mọi tài liệu quay về checkpoint trước lượt chạy.
  - `Transaction` lồng nhau: lỗi bên trong chỉ hoàn tác phần của nó; `Try` bao ngoài có thể chạy tiếp.
  - Chạy từng bước, breakpoint theo đường dẫn: dừng trước bước, cho xem biến, chọn Tiếp tục / Bước tiếp / Dừng.
  - Nhật ký từng bước (action, số đối tượng, nhánh, vòng lặp, lỗi bị bắt, hoàn tác).
  - Chạy thử (`dryRun`): xem trước mọi action, không đổi gì.
  - Giới hạn số bước và thời gian, huỷ bằng `AbortSignal`. `Try` **không** bắt được các lỗi này (dùng làm
    sandbox cho P9).
  - Biến như VBA: biến mới thuộc cả lượt chạy; biến vòng lặp và biến lỗi chỉ sống trong khối của nó.
- **Dòng sự kiện của Action Core**: `ActionRegistry.subscribe` phát `ActionEvent` sau mỗi action chạy xong (không
  phát khi chạy thử). `announceUi` dùng cho lệnh giao diện chạy theo đường riêng nhưng tương đương một action
  của catalog: input được kiểm theo schema của action (ghi được thì chạy lại được), không chạy gì.
- **Recorder** (`WorkflowRecorder`) chỉ ghi action có caller `ui`. Action do AI, workflow hay script chạy không
  bao giờ bị ghi.
  - Chuẩn hoá:
    - bỏ thao tác chỉ đổi vùng chọn (nhóm `UI` hoặc mức `read`);
    - gộp các `Set…` liên tiếp cùng đích, giữ giá trị cuối;
    - `…Undo` triệt tiêu bước vừa ghi, `…Redo` đưa lại; action mới thì xoá ngăn redo.
  - Suy tham số: đối số cấp đầu bằng một giá trị của bản ghi đang chọn thì thành `record.<khoá>`, và workflow
    có tham số `record` (mặc định là bản ghi lúc ghi). Tuỳ chọn "lặp cho từng bản ghi" bọc các bước trong
    `ForEach record In records`.
  - Lệnh chưa đi qua Action Core thì báo "Không ghi được: …", không âm thầm thiếu bước.
- **Lưu trữ**:
  - trong model Smart Data của mẫu: `workflows` tuỳ chọn trong phần `objects` của customXml. Tài liệu không có
    workflow giữ nguyên hash. Workflow định dạng mới hơn bị bỏ qua kèm thông báo, không làm hỏng model;
  - hoặc file `<tên>.dvhflow.json` trong thư mục dự án (xuất/nhập qua hộp thoại).
- **Docs**:
  - Một registry dùng chung cho giao diện, có Recorder gắn sẵn.
  - Lệnh của panel Smart Data đi qua Action Core:
    - sửa giá trị Field ngay trong bảng: chạy thẳng `Data.SetField`;
    - phát sự kiện ngữ nghĩa: cập nhật từ nguồn (`Link.Update`), chèn/làm mới/đổi kiểu bảng (`Document.InsertTable`,
      `Table.Refresh`, `Table.SetStyle`), điều kiện, bọc khối, chèn cột (`Template.*`), sinh hàng loạt
      (`Document.Generate`);
    - báo "không ghi được": liên kết workbook, ghi ngược, giải xung đột, đổi nguồn, đổi chế độ, bỏ liên kết, chèn
      Field tại con trỏ, thêm Field, chuyển mẫu DVH-Tool, nhập QLCL, khôi phục/hoàn tác lịch sử.
  - Action mới: `Workflow.List`, `Workflow.Run` (tham số, hoặc `collection` để đổ bản ghi vào tham số danh sách),
    `Workflow.Save`, `Workflow.Delete` (mức `destructive`). AI (P7) và MCP (P4) gọi được ngay.
    - Người dùng chạy: workflow có quyền của người dùng, bước nguy hiểm vẫn hỏi xác nhận.
    - AI hoặc script chạy: chỉ có quyền đọc và ghi.
  - Mục **"Quy trình"** trong panel:
    - chọn bản ghi đang chọn; Ghi / Dừng ghi; danh sách bước đã ghi và cảnh báo; lưu thành quy trình;
    - **trình sửa dạng khối**: thụt lề theo cấp; lên/xuống/xoá; sửa bước (JSON có kiểm); bọc trong Transaction/Try;
      thêm bước action từ catalog; breakpoint từng bước;
    - chạy / chạy thử / chạy từng bước: bước đang dừng được tô, có xem biến;
    - tham số dạng JSON, hoặc lấy dữ liệu từ một collection;
    - lưu vào tài liệu, xuất ra thư mục dự án, nhập, xoá.

  36 chuỗi mới cho cả 20 ngôn ngữ.

## Tệp chính

- `packages/dvh-workflow/src`:
  - `model.ts` (schema, đường dẫn, kiểm tra), `engine.ts` (chạy, giao dịch, debug, giới hạn);
  - `recorder.ts` (Recorder, `recordsOf`), `storage.ts` (customXml/file), `edit.ts` (thao tác khối,
    `describeStep`).
- `packages/dvh-actions/src/index.ts`: `ActionEvent`, `subscribe`, `announceUi`.
- `packages/dvh-model`: `dvhStoredWorkflowSchema`, `workflows` tuỳ chọn trong model và customXml.
- Docs:
  - renderer: `dvh-workflow.ts` (registry giao diện, recorder, runner), `dvh-actions.ts` (`Workflow.*`, `confirm`
    của host), `components/DvhWorkflowSection.tsx`;
  - panel, bảng, mẫu, lịch sử: phát sự kiện hoặc báo không ghi được; CSS dùng token;
  - main: IPC `docs:dvh-workflow-export`/`-import`; preload và kiểu IPC.
- Phụ thuộc: `@genoffice/dvh-workflow` trong `dependencies` của Docs và shell, `exclude` của
  `externalizeDepsPlugin` (Docs), script `test`/`typecheck` gốc, lockfile.

## Kiểm thử (vitest)

- `packages/dvh-workflow/tests/workflow.test.ts` (11):
  - engine:
    - ForEach có `where`, If/ElseIf/Else, Set, `assign`, biểu thức, `.Count`; mọi change-set chung một `txId`,
      nguồn `workflow`;
    - chạy thử không đổi gì;
    - Transaction trong Try: chỉ hoàn tác phần trong, chạy tiếp; lỗi không bắt thì hoàn tác cả lượt;
    - từng bước: dừng đúng các đường dẫn; breakpoint cho xem biến; Dừng thì hoàn tác;
    - giới hạn bước (Try không bắt được); action `destructive` bị chặn với quyền mặc định; ép kiểu tham số,
      tham số thiếu;
    - workflow sai bị từ chối, action ngoài catalog được báo;
  - **nghiệm thu Recorder**:
    - chuẩn hoá: gộp setter, bỏ chọn, triệt tiêu undo/redo, không ghi trước khi bắt đầu, sau khi dừng hay
      lệnh do AI chạy;
    - **đổi bố cục giao diện**: cùng thao tác từ hai "bố cục" khác nhau (ribbon và panel, gọi action theo cách
      khác nhau) cho ra workflow giống từng byte;
    - **chạy lại với bộ dữ liệu khác**: workflow "lặp cho từng bản ghi" chạy trên tài liệu mới với hai bản ghi
      khác; không truyền dữ liệu thì chạy lại đúng bản ghi lúc ghi;
  - lưu trữ (customXml đi–về, bỏ qua định dạng lạ, xoá workflow cuối thì XML như cũ, file JSON, tên file hợp
    lệ) và thao tác khối (bọc, chèn, di chuyển, xoá, không sửa bản gốc).
- `apps/docs/tests/dvh-workflow.test.ts` (2):
  - lệnh giao diện qua registry và `announceUi` (sai input thì bỏ) được ghi;
  - lưu vào model, đi–về qua customXml;
  - AI chạy `Workflow.Run` với bản ghi khác: Field trong trang đổi, bước không bị ghi, một `txId`;
    `revertTransaction` trả về như cũ;
  - bước lỗi thì cả workflow hoàn tác; lưu workflow có action lạ bị từ chối.
- Bộ đầy đủ: Docs 3.146/3.146; `dvh-actions` 23; `dvh-model` 31; typecheck Docs, shell, Sheets và
  `dvh-workflow` sạch; eslint, prettier, theme-colors, english-comments đạt.

## Để lại cho máy Windows (chưa chạy)

1. Trong Docs:
   - chọn bản ghi → Ghi → sửa hai Field trong bảng giá trị, làm mới một DVH.Table, sinh hàng loạt → Dừng;
   - "Lưu thành quy trình" (lặp cho từng bản ghi) → Chạy với dữ liệu từ collection;
   - Lịch sử có một mục nguồn workflow; "Hoàn tác lượt chạy" trả về như cũ.
2. Đặt breakpoint, chạy từng bước, xem biến, Dừng giữa chừng: tài liệu quay về như trước.
3. Lưu tài liệu, mở lại: quy trình còn. Xuất `.dvhflow.json`, nhập vào tài liệu khác.
4. Mở file đã lưu bằng Word: không báo sửa file.

## Giới hạn còn lại

- Recorder hiện có ở Docs. Sheets và Slides đã có registry, sự kiện và saga participant, nhưng chưa có mục ghi
  trên giao diện.
- Một số lệnh của panel phát sự kiện tương đương thay vì chạy action. Ví dụ "chèn cột" ghi theo đoạn văn
  (`Template.InsertColumn`) chứ không theo vị trí con trỏ; chèn Field tại con trỏ chưa ghi được.
- Trình sửa khối sửa tham số bằng JSON; chưa có form theo schema của từng action. Chưa kéo thả.
- Workflow chạy trong một tài liệu (tài liệu đang mở). Engine đã nhận nhiều participant (`doc` theo bước); nối
  vào shell cho workflow nhiều tài liệu để sau.
