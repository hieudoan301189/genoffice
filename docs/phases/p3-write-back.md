# P3 (phần 2) – Ghi ngược Field và tìm nguồn nâng cao

Ngày: 02/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md),
mục P3. Phần 1: [p3-auto-link.md](p3-auto-link.md).

Kiểm bằng vitest trên môi trường cloud (Linux, không có Windows, Word/Excel hay bản đóng gói Electron). Các bài
kiểm trên bản đóng gói và bằng Word/Excel thật **chưa chạy**; danh sách ở mục cuối, để chạy trên máy Windows.

## Kết luận

- **Ghi ngược Field đơn (Read/Write)** chạy theo hai đường:
  1. Workbook **đang mở trong DVH Sheets**: Docs gửi yêu cầu qua main process tới đúng tab đó. Tab chạy
     `Data.SetField`, tức là một lần sửa ô bình thường: Undo được, lưu cùng workbook, ghi change-set
     `source: link`.
  2. Workbook **đang đóng**: main process vá thẳng file xlsx ở ba chỗ. Ô mà tên ẩn `_dvh.f.<id>` trỏ tới;
     Field trong phần model (text và `rev`, vá tại chỗ bằng `setFieldTextInXml`); một change-set
     `Data.SetField` thêm vào phần lịch sử (tạo phần lịch sử nếu chưa có). Đặt `fullCalcOnLoad` để Excel tính
     lại công thức phụ thuộc. Ghi bằng atomic write.
- **Xung đột** phát hiện theo revision của từng Field, kèm so sánh giá trị:
  - mỗi lần DVH ghi một Field thì `rev` tăng;
  - Excel không biết `rev`, nên giá trị đọc từ ô cũng được so;
  - liên kết giữ "base" (revision và text lần cuối hai bên thống nhất) của từng Field;
  - bên ghi kiểm lại base ngay trước khi ghi.

  Khi xung đột, Link Manager cho chọn **Giữ bên này** (ghi đè nguồn), **Giữ nguồn**, hoặc **Lịch sử** (lọc lịch
  sử theo Field đó).

- **Tìm nguồn**: thử đường dẫn tương đối, rồi đường dẫn cuối, rồi các file mà `fileMap` của project-store
  biết, rồi chỉ mục `docId` của shell. Mọi ứng viên đều được xác nhận bằng cách đọc `docId` trong phần model.
  - Một kết quả: liên kết tự trỏ sang đó và ghi `Link.Relocate`.
  - Nhiều kết quả: báo "mơ hồ" và liệt kê để người dùng chọn.
  - Không có kết quả: báo "mất nguồn".
- Link Manager có thêm: **Ghi về nguồn**, **Mở nguồn**, **Đổi nguồn…**, số Field đã sửa ở đây mà chưa ghi về
  nguồn, trạng thái **xung đột** và **mơ hồ**.

## Cách chạy

### Khi nào ghi ngược

- Liên kết **Tự động**: khoảng 1,2 giây sau khi ngừng gõ, các Field Read/Write đã sửa được ghi về nguồn.
- Liên kết **Thủ công** và **Khi mở**: người dùng bấm "Ghi về nguồn".
- Field `access: read` không bao giờ ghi ngược; khi kéo về, nguồn luôn thắng.

### So sánh ba bên khi kéo về (`threeWay` trong `dvh-model`)

| Ở đây | Ở nguồn | Kết quả                                                             |
| ----- | ------- | ------------------------------------------------------------------- |
| không | đổi     | lấy giá trị nguồn                                                   |
| đổi   | không   | giữ bên này, chờ ghi ngược (không còn bị nguồn ghi đè như trước P3) |
| đổi   | đổi     | **xung đột**: giữ nguyên bên này, liệt kê trong Link Manager        |

Hai bên giống nhau thì base tiến lên, không cần làm gì. Liên kết tạo trước P3 chưa có base, nên nguồn thắng như
cũ. Base được ghi ngay ở lần kéo đầu tiên.

### Giao thức (`packages/dvh-model/src/writeback.ts`)

- `docs:dvh-write-fields` (Docs → main): `{ docId, paths, writes[{ fieldId, value, expected, force }], origin }`,
  kiểm bằng zod.
- Main giữ bảng `docId → tab Sheets` lấy từ kênh `dvh:live-publish`. Sheets phát một lần ngay khi workbook
  nạp xong, nên tab mở mà chưa sửa gì vẫn được biết tới.
- Main gửi `sheets:dvh-set-fields` tới tab đó và chỉ nhận trả lời `dvh:set-fields-result` **từ chính tab
  đó**, chờ tối đa 5 giây. Tab trả `handled: false` (không giữ workbook ấy) hoặc hết giờ: chuyển sang vá file.
- Kết quả từng Field: `written`, `conflict`, `formula` (ô là công thức, không ghi đè), hoặc `missing`.
- File bị Excel khoá trên Windows: atomic write lỗi và báo rõ; không có gì bị ghi.

### Sau khi ghi

- Docs ghi change-set `Link.WriteBack` và cập nhật base.
- Nếu ghi vào file và nguồn chỉ khác lần đồng bộ trước ở những Field vừa ghi, `lastSync.hash` tiến lên. Nhờ
  vậy liên kết thủ công không tự báo "cũ" vì chính lần sửa của mình. Nguồn có thay đổi khác thì vẫn báo cũ.

### Tìm nguồn

- `apps/docs/src/main/dvh-find-source.ts`: `peekDvhDocId` đọc `docId` từ phần model (tìm theo namespace),
  có cache theo kích thước và mtime.
- `docs:dvh-find-source` lấy ứng viên từ `projectFilePaths()` (khoá `fileMap` đi theo các lần đổi tên trong
  shell) và hook `findDvhDocId` của shell.
- Chỉ mục file của shell (`file-index/store.ts`):
  - thêm cột `dvh_doc_id` cùng index, do worker trích xuất điền cho xlsx/xlsm/docx;
  - CSDL cũ được thêm cột tại chỗ, và các file Office của nó được đánh dấu để trích xuất lại.
- **Đổi nguồn** nhận workbook cùng `docId`, hoặc một bản sao chứa đủ mọi Field/collection mà liên kết đọc (theo
  ID). Workbook khác thì từ chối và báo "không phải nguồn".
- **Mở nguồn** mở workbook trong tab Sheets qua bộ định tuyến của shell (chạy độc lập thì dùng `shell.openPath`).

## Tệp chính

- `packages/dvh-model`:
  - `types.ts`: thêm `rev` cho Field và `lastSync.fields` (base từng Field);
  - `custom-xml.ts`: đọc/ghi `rev`; `setFieldTextInXml(…, rev)`;
  - `writeback.ts` (mới).
- `apps/docs/src/main`:
  - `dvh-xlsx-write.ts` (mới): vá xlsx;
  - `dvh-link-watch.ts`: chuyển yêu cầu ghi sang Sheets / ghi file;
  - `dvh-find-source.ts` (mới);
  - `docs-main.ts`: IPC tìm và mở nguồn;
  - `dvh-xlsx-cells.ts`: tách `workbookLayout`.
- `apps/docs/src/renderer`:
  - `dvh-smart-data.ts`: kéo về theo ba bên, base, `relocateLink`, `applyFieldText`;
  - `dvh-auto.ts`: ghi ngược, xung đột, tự ghi sau khi gõ, tìm nguồn;
  - `DvhSmartDataPanel.tsx` và `dvh-smart-data.css`: giao diện, dùng token.
  - i18n: 20 khoá mới trong `i18n/dvh/*.ts`, cả 20 ngôn ngữ.
- `apps/sheets`:
  - `dvh-live.ts`: phát lần đầu, xử lý `handleSetFields`;
  - `dvh-smart-data.ts`: tăng `rev`;
  - preload và `desktop-api.ts`.
- `apps/shell/src/main/file-index/*`: cột `dvh_doc_id`; `index.ts`: hook `findDvhDocId`.
- `apps/docs`: thêm phụ thuộc `@genoffice/xlsx-gateway`, có trong cả `dependencies` lẫn `exclude` của
  `externalizeDepsPlugin`; `docx-engine` export `findCustomXmlItemByNamespace`.

## Kiểm thử (vitest)

- `packages/dvh-model/tests/writeback.test.ts` (5):
  - `rev` đi trọn vòng qua phần model; `setFieldTextInXml` đặt hoặc thay `rev`;
  - các trường hợp của `threeWay`, `checkFieldWrite`, schema.
- `apps/docs/tests/dvh-xlsx-write.test.ts` (4):
  - workbook kiểu Excel (chuỗi dùng chung, công thức, không có phần lịch sử): ghi ô (giữ style, tạo dòng mới đúng
    thứ tự), ghi model (`rev`), tạo phần lịch sử kèm rels và content type, lần ghi sau nối thêm vào lịch sử;
  - xung đột do Excel sửa (cùng `rev`, khác giá trị) và do DVH (khác `rev`); `force`; công thức; Field lạ;
    workbook khác `docId`.
- `apps/docs/tests/dvh-writeback.test.ts` (4) – **nghiệm thu P3**, editor thật với bytes xlsx thật:
  - **round-trip**: sửa ở Docs → "Ghi về nguồn" → ô và `rev` của workbook đổi, workbook có lịch sử
    `Data.SetField`, Docs có `Link.WriteBack`; lần kéo sau báo `current`;
  - **xung đột**: Docs và Excel cùng sửa → giữ chữ ở Docs, báo xung đột; ghi thường bị từ chối; "Giữ nguồn" rồi
    "Giữ bên này" (ghi đè);
  - liên kết **Tự động** tự ghi về sau khi gõ;
  - **đổi tên hoặc chuyển nguồn**: tìm lại theo `docId`, ghi `Link.Relocate`, lần ghi sau đi tới chỗ mới; hai bản
    sao thì báo `ambiguous`, không tự chọn; không còn file nào thì báo `missing`.
- `apps/docs/tests/dvh-find-source.test.ts` (1): xác nhận theo `docId`, bỏ file khác, bỏ trùng.
- `apps/sheets/tests/dvh-live-write.test.ts` (3), Univer giả lập:
  - ghi ô, tăng `rev`, change-set `link`, đánh dấu cần lưu;
  - ô đã sửa mà chưa lưu thì báo xung đột, `force` thì ghi;
  - công thức; workbook khác; yêu cầu sai dạng.
- `apps/shell/tests/file-index-store.test.ts` (+2): cột `docId`; nâng cấp CSDL cũ.
- Bộ test đầy đủ:
  - Docs 3.131/3.131;
  - `dvh-model` 26; `docx-engine` 1.553 (1 bỏ qua);
  - test DVH của Sheets 70; file-index của shell 43.
  - toàn bộ Sheets 2.904/2.938. 34 ca lỗi đều đã lỗi sẵn trên nhánh trước thay đổi, vì môi trường cloud không có
    sidecar Rust (`xlsx-sidecar` chưa build), LibreOffice, hay cơ chế khoá file của Windows. Đã kiểm lại từng
    file trên bản trước thay đổi.

  Typecheck docs/sheets/shell/dvh-model sạch; `check:theme-colors`, `check:english-comments`, eslint và prettier
  đều đạt.

## Để lại cho máy Windows (chưa chạy)

Cần build lại shell, vì main process của Docs được biên dịch vào shell, rồi chạy trên bản đóng gói.

1. **Round-trip qua Sheets đang mở**: mở workbook và báo cáo trong cùng phiên. Sửa Field ở Docs → ô trong Sheets
   đổi, workbook ở trạng thái chưa lưu, Ctrl+Z trong Sheets hoàn tác được. Lưu → lịch sử có `Data.SetField`
   (`source: link`).
2. **Round-trip với workbook đóng**: đóng tab Sheets, sửa Field ở Docs, bấm "Ghi về nguồn". Mở workbook bằng
   **Excel**: ô mới, không có hộp thoại "sửa file", công thức phụ thuộc tính lại.
3. **Excel đang mở file**: ghi về phải báo lỗi rõ ràng, file không hỏng.
4. **Xung đột**: sửa cùng Field ở Excel (lưu) và ở Docs → Link Manager báo xung đột; thử cả ba nút.
5. **Đổi tên / chuyển**:
   - đổi tên workbook trong Home của shell, rồi mở lại báo cáo → tự nối lại;
   - chuyển workbook sang thư mục đã thêm vào cây thư mục, chờ chỉ mục quét xong → tự nối lại;
   - chép thành hai bản → báo mơ hồ, chọn một.
6. **Mở nguồn** mở workbook trong tab Sheets; **Đổi nguồn** từ chối workbook khác.
7. **Word** mở báo cáo đã lưu sau khi ghi ngược: Field `dvh:f` vẫn gắn dữ liệu, hiện giá trị mới.

Có thể viết driver kiểu `scripts/p3/auto-link.mjs` cho các bước 1, 2, 4, 5. Driver này chưa viết vì không chạy
thử được ở đây.

## Giới hạn còn lại

- Chỉ ghi ngược **Field đơn**; bảng vẫn một chiều (quyết định H).
- Tự ghi ngược chỉ áp dụng cho liên kết Tự động; ở chế độ khác, sửa ở Docs nằm chờ cho tới khi bấm "Ghi về nguồn".
- Chỉ mục `docId` chỉ biết các thư mục mà shell đang lập chỉ mục (thư mục lưu, các thư mục đã thêm, gần đây,
  có sao). File chuyển ra ngoài các thư mục đó thì phải dùng "Đổi nguồn".
- Một lần ghi ngược vào tab Sheets là một mục Undo cho mỗi Field; gom thành một mục là việc của P4
  (transaction).
