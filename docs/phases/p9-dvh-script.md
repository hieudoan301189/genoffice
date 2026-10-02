# P9 – DVH-Script + AI sinh workflow/script

Ngày: 02/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md),
mục P9. Kiểm bằng vitest trên môi trường cloud (Linux, không có bản đóng gói).

## Quyết định đầu phase: parser viết tay

Chọn **parser viết tay**, không dùng Lezer:

- repo không có `@lezer/generator` để biên dịch grammar;
- ngôn ngữ đi theo dòng (mỗi lệnh một dòng, khối đóng bằng `End …`), parser tay vừa nhỏ vừa báo lỗi đúng dòng và
  cột;
- là nguồn duy nhất cho round-trip.

Tô màu dùng `StreamLanguage` của CodeMirror, cũng là thư viện soạn thảo mà ứng dụng HTML đã dùng.

## Kết luận

- **Package mới `@genoffice/dvh-script`**: parser, printer, kiểm tra theo catalog, autocomplete, bộ chạy sandbox,
  skill AI và tài liệu ngôn ngữ.
- **Đặc tả v0** (gần VBA; từ khoá không phân biệt hoa thường; `'` mở chú thích; chuỗi trong ngoặc kép, `""` là
  một dấu nháy):

  ```vb
  Workflow "Biên bản nghiệm thu"
  Id "wf_bb"
  Param records As List = Json([]) Description "các hạng mục"

  total = 0
  For Each row In records Where row.qty > 0
      n = Data.AddRow(text:="Hạng mục " & row.code)
      total = total + row.qty
  Next
  If total > 100 Then
      Log "lớn"
  ElseIf total > 0 Then
      Data.SetField(field:="Total", value:=total) On "doc_report"
  Else
      Log "không có"
  End If
  Try
      Transaction
          Data.DeleteRows(count:=2)
      End Transaction
  Catch err
      Log "lỗi: " & err.message
  End Try
  ```

  - Gán biến; gọi action dạng hàm với tham số đặt tên `:=`; `On <tài liệu>` để chạy ở tài liệu khác.
  - `If/ElseIf/Else/End If`, `For Each … In … [Where …] … Next`, `Try … Catch [biến] … End Try`,
    `Transaction … End Transaction`, `Log`.
  - Giá trị: chuỗi, số, `True/False/Null`, và `Json(...)` cho danh sách, bản ghi hay chuỗi có xuống dòng.
  - Biểu thức là của P6.

- **Biên dịch và in ngược**:
  - `compileScript` cho ra Workflow Model (P8) kèm bản đồ đường dẫn bước → dòng (cho báo lỗi và debugger);
  - `printScript` in model thành script;
  - `canonicalWorkflow` chỉ đổi cách viết, không đổi nghĩa: biểu thức cắt khoảng trắng, chuỗi `'…'` thành `"…"`,
    biểu thức chỉ là hằng thì thành giá trị;
  - với model dạng chuẩn: **`parse(print(w)) == w`**, và `print(parse(print(w))) == print(w)` từng ký tự;
  - Workflow Model bắt mỗi biểu thức nằm trên một dòng, để script luôn in được.
- **Kiểm theo catalog** (`checkScript`), báo kèm dòng:
  - action không có trong catalog, tham số lạ, thiếu tham số bắt buộc;
  - hằng sai kiểu so với JSON Schema của action;
  - cảnh báo action `destructive`/`external` sẽ phải chờ người dùng xác nhận.
- **Autocomplete** (`completeAt`, độc lập với editor):
  - tên action (`Group.Verb`), tham số của action đang gọi (kèm bắt buộc/tuỳ chọn, kiểu, mô tả; bỏ tham số đã
    dùng);
  - từ khoá đầu lệnh (từ khoá đầu file chỉ gợi ý trước lệnh đầu tiên);
  - tham số, biến, biến vòng lặp và hàm biểu thức; không gợi ý trong chuỗi hay chú thích.
- **Chạy** (`runScript`): biên dịch, kiểm, rồi thông dịch model bằng engine P8 với caller `script`.
  - Không `eval`, không có gì chạm tới Node, file hay mạng: lối ra duy nhất là action của catalog, có kiểm quyền
    (mặc định chỉ đọc và ghi).
  - Giới hạn 10.000 bước và 30 giây; `Try` không bắt được các giới hạn này.
  - Lỗi trả về đúng dòng (`lineOfPath`). Cả lượt chạy là một giao dịch.
- **AI sinh script và workflow** (`createDvhScriptSkill`):
  - tool `dvh_script_reference`: tài liệu ngôn ngữ và action kèm tham số;
  - tool `propose_script`: nhận `source` hoặc `workflow` JSON, với một trong ba chế độ:
    - `check`: kiểm và trả script đã chuẩn hoá;
    - `save`: lưu thành workflow của tài liệu;
    - `run`: chạy thử, hỏi xác nhận theo chính sách P4 (action nguy hiểm hoặc vượt ngưỡng đối tượng), rồi chạy
      như một giao dịch có `txId`.
  - Lỗi trả về kèm dòng để mô hình sửa. Nội dung tài liệu vẫn chỉ là dữ liệu.
  - Tác tử Docs có skill này cùng `dvh-planner`; tác tử Sheets có `check` và `run`.
- **Docs**:
  - mục "Quy trình" có thêm chế độ **Script** bên cạnh **Khối**, chuyển qua lại không mất gì;
  - trình soạn CodeMirror: số dòng, tô màu bằng token giao diện, autocomplete từ catalog, gạch lỗi/cảnh báo kèm
    thông báo;
  - **debugger**: chạy từng bước và breakpoint dùng engine P8; dòng đang dừng được tô và cuộn tới; lỗi chạy báo
    kèm số dòng;
  - `Workflow.Run` có giới hạn lồng nhau 8 tầng, để workflow tự gọi chính nó không chạy mãi.

  4 chuỗi mới cho cả 20 ngôn ngữ.

## Tệp chính

- `packages/dvh-script/src`:
  - `lexical.ts` (chú thích, quét ngoài chuỗi/ngoặc, hằng), `parse.ts`, `print.ts`;
  - `check.ts`, `complete.ts`, `run.ts`, `skill.ts`, `reference.ts`.
- `packages/dvh-workflow/src/model.ts`: biểu thức một dòng; kiểu optional tương thích
  `exactOptionalPropertyTypes` (Sheets).
- Docs:
  - `components/DvhScriptEditor.tsx` (mới), `components/DvhWorkflowSection.tsx` (chế độ Script, dòng đang dừng);
  - `ai/AiPanel.tsx` (skill script); `dvh-actions.ts` (giới hạn lồng `Workflow.Run`).
- Sheets: `App.tsx` (skill script).
- Phụ thuộc:
  - `@genoffice/dvh-script` trong `dependencies` và `exclude` của Docs và Sheets;
  - CodeMirror (`state`, `view`, `autocomplete`, `commands`, `language`, `lint`) và `@lezer/highlight` vào
    `devDependencies` của Docs, cùng phiên bản với ứng dụng HTML;
  - script `test`/`typecheck` gốc, lockfile.

## Kiểm thử (vitest)

- `packages/dvh-script/tests/roundtrip.test.ts` (4) – **nghiệm thu round-trip**:
  - script mẫu (đủ mọi loại lệnh) biên dịch đúng model và bản đồ dòng; in rồi đọc lại bằng nguyên model;
  - 9 loại lỗi báo đúng dòng, cột đúng vị trí giá trị lỗi;
  - dạng chuẩn hoá;
  - **300 workflow sinh ngẫu nhiên** (có seed):
    - đủ loại bước, lồng tới 3 tầng;
    - biến trùng từ khoá (`Else`, `Next`, `If`, `Log`);
    - chuỗi có nháy, `'`, xuống dòng, tab, `\`, tiếng Việt; số âm, số thập phân, `1e21`;
    - danh sách và bản ghi lồng nhau; tham số, mô tả, `On`, `assign`, `Catch` có và không có biến;
    - `parse(print(w))` bằng `w` (cả từng byte JSON); in lại giống từng ký tự; tổng cộng hơn 100 action không
      mất action hay tham số nào.
- `packages/dvh-script/tests/run.test.ts` (6):
  - kiểm catalog (tham số lạ, thiếu, sai kiểu, action lạ, cảnh báo destructive);
  - autocomplete (action, tham số, từ khoá, tên trong phạm vi, hàm, không gợi ý trong chuỗi);
  - chạy một giao dịch, lỗi về đúng dòng và hoàn tác;
  - **nghiệm thu sandbox**: các cách sau đều bị chặn:
    - `process.env`, `globalThis.fetch`, `require("fs")`, `eval(...)`, `r.constructor`, `r.__proto__`;
    - action đọc file ngoài catalog;
    - action `destructive` không được cấp quyền.

    Chuỗi chứa `${…}` hay backtick chỉ là chữ. Bốn vòng lặp lồng nhau trong `Try` vẫn dừng ở 10.000 bước.

  - skill AI:
    - lỗi kèm dòng; `check`, `save`, `run` (có tham số); nhận workflow JSON;
    - script xoá cần xác nhận: từ chối thì không đổi gì, đồng ý thì chạy.
- `apps/docs/tests/dvh-script.test.ts` (2):
  - tác tử trong Docs: lỗi tham số kèm dòng; lưu workflow, in lại đúng script gốc; chạy script đổi Field trong
    trang (nguồn `script`); `Workflow.Run` theo tên;
  - workflow tự gọi chính nó dừng ở giới hạn 8 tầng và không đổi gì.
- Bộ đầy đủ: Docs 3.148/3.148; test DVH và AI của Sheets 100/100; `dvh-workflow` 11; `dvh-script` 10.
  Typecheck Docs, Sheets, shell, `dvh-workflow`, `dvh-script` sạch; eslint, prettier, theme-colors,
  english-comments đạt.

## Để lại cho máy Windows (chưa chạy)

1. Mục Quy trình → Script:
   - gõ `Data.` thấy danh sách action; trong ngoặc thấy tham số; lỗi bị gạch kèm thông báo;
   - Áp dụng, chuyển sang Khối rồi quay lại: script giữ nguyên.
2. Đặt breakpoint, chạy từng bước: dòng đang dừng được tô trong trình soạn; Dừng thì tài liệu quay về như cũ.
3. Hỏi tác tử "viết quy trình đặt tên dự án cho từng hạng mục rồi lưu lại": có `propose_script` (`save`), quy
   trình hiện trong danh sách. "Chạy nó với …": có hộp thoại xác nhận nếu cần, Lịch sử có một mục nguồn script.
4. Màu tô cú pháp ở giao diện sáng và tối.

## Giới hạn còn lại

- v0 chưa có `Sub`/hàm người dùng, `While`, mảng chỉ số (`a(1)`); một lệnh một dòng (chưa có `_` nối dòng).
- Chú thích không được giữ khi script đi qua Workflow Model; round-trip được định nghĩa từ phía model.
- Debugger chưa cho sửa biến khi đang dừng; xem biến dạng JSON.
- Sheets mới có skill AI (`check`/`run`); chưa có trình soạn script và chưa lưu workflow trong workbook.
