# P10 – Project Data Model

Ngày: 02/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md),
mục P10. Kiểm bằng vitest trên môi trường cloud (Linux, không có bản đóng gói, không có QLCL-DVH thật).

## Kết luận

- **Package mới `@genoffice/dvh-project`** (logic thuần): gói schema, dữ liệu dự án, kiểm tra quy tắc, nguồn ảo
  cho tài liệu liên kết, ghi ngược Field và đồng bộ hai chiều.
- **Kho cấp dự án đặt trong `project-store`**:
  - `ProjectStore.readProjectData`/`writeProjectData` lưu `projects/<id>/data/project-data.json` (ghi nguyên tử,
    tối đa 32 MB), đi cùng dự án và bị xoá cùng dự án;
  - `project-store` chỉ giữ JSON, `dvh-project` mới kiểm schema; file hỏng thì đọc như chưa có dữ liệu.
- **Lõi không có schema nghiệp vụ (D10)**:
  - kiểu đối tượng, quan hệ (`ref`), giá trị liệt kê, trường bắt buộc, quy tắc kiểm tra (biểu thức an toàn của
    P6, có `when`) và mẫu đi kèm đều nằm trong **gói schema** (`dvh-schema-pack` v1, file JSON);
  - gói **QLCL công trình xây dựng** (`packages/dvh-project/packs/qlcl-xay-dung/pack.json`) có `Project` (một mỗi
    dự án), `Contractor`, `WorkItem`, `Acceptance`, `Material`, `Test` và 3 quy tắc. Đây là dữ liệu, không phải mã;
  - một test quét mã nguồn của lõi để chắc không có từ vựng nghiệp vụ nào;
  - cài nhiều gói được, tên kiểu không được trùng giữa các gói; gỡ gói thì đối tượng vẫn giữ (thành "chưa có
    kiểu").
- **Đối tượng dự án**:
  - mỗi giá trị có revision riêng (như Field từ P3); giá trị nhập dạng chữ được ép theo kiểu của trường;
  - mỗi thay đổi là một change-set (P5), lưu 500 mục gần nhất;
  - kiểm tra: thiếu trường bắt buộc, sai kiểu số, giá trị ngoài danh sách, tham chiếu tới đối tượng không có
    hoặc sai kiểu, kiểu "một mỗi dự án" bị trùng, quy tắc của gói (đọc được `ref.key` qua tham chiếu).
- **Tài liệu liên kết thẳng tới đối tượng dự án**:
  - dự án là một nguồn liên kết ảo `dvh-project://<projectId>` (docId `proj_<id>`);
  - kiểu "một mỗi dự án" thành Field `Type.key`; kiểu danh sách thành collection `Type` có cột `ID` đầu tiên;
  - id ổn định (`pf_<objectId>_<key>`, `pc_<Type>`), nên toàn bộ cơ chế của P3 dùng được nguyên vẹn: liên kết,
    kéo ba chiều, chế độ Khi mở/Tự động, **ghi ngược có kiểm xung đột** theo revision;
  - main process thông báo `docs:dvh-source-changed` cho tài liệu đang theo dõi dự án sau mỗi thay đổi (sửa
    trong panel, ghi ngược, đồng bộ).
- **Trao đổi hai chiều với QLCL-DVH** (`syncProject`):
  - định dạng: workbook QLCL hoặc JSON `dvh-exchange` của P6;
  - khớp theo ID: cột `ID`, hoặc theo kiểu với đối tượng "một mỗi dự án";
  - so ba chiều từng giá trị với **bản gốc của lần đồng bộ trước với đối tác đó**: chỉ một bên đổi thì bên đó
    thắng; cả hai đổi khác nhau là **xung đột**, không ghi vào bên nào cho tới khi người dùng chọn;
  - xoá cũng so ba chiều: xoá một bên mà bên kia không sửa thì xoá cả hai; xoá một bên mà bên kia có sửa là xung
    đột;
  - lần đồng bộ đầu (chưa có bản gốc): giá trị khác nhau là xung đột, trừ khi một bên trống, để không trộn đoán
    hai nguồn độc lập;
  - dòng QLCL thêm mà chưa có ID thì được tạo đối tượng và **nhận ID** trong file gửi lại;
  - file đối tác được ghi lại bằng dữ liệu đã trộn. Ở chỗ xung đột, file vẫn giữ giá trị của đối tác (đúng
    kiểu: số vẫn là số);
  - **giải xung đột** (`resolveSyncConflict`):
    - "giữ của tôi": được đẩy sang đối tác ở lần đồng bộ sau;
    - "lấy của QLCL": áp dụng ngay;
    - xung đột xoá cũng giải theo cùng cách.
- **Docs**:
  - mục **"Dữ liệu dự án"** trong panel Smart Data:
    - tên dự án; gói schema đã cài và cài gói từ file;
    - chọn kiểu đối tượng, bảng đối tượng sửa tại chỗ (danh sách chọn cho trường liệt kê); thêm, xoá đối tượng;
    - danh sách vấn đề do kiểm tra tìm ra;
    - "Liên kết tài liệu với dữ liệu dự án";
    - đồng bộ với QLCL-DVH (chọn file), đồng bộ lại; xung đột kèm "Giữ của tôi"/"Lấy từ nguồn";
  - main:
    - `ProjectDataService`;
    - IPC `docs:dvh-project-info`, `-install-pack`, `-objects`, `-set`, `-delete`, `-sync`, `-resolve`;
    - `docs:dvh-read-source` đọc được `dvh-project://`; ghi ngược đi qua `setProjectFieldWriter`;
    - theo dõi tự động nhận URI dự án.

  18 chuỗi mới cho cả 20 ngôn ngữ.

- **Gói `.dvh`**: cân nhắc rồi **chưa làm**.
  - Dữ liệu dự án nằm trong thư mục dự án; mỗi tài liệu tự mang model, lịch sử và workflow trong customXml.
  - Việc đóng gói hồ sơ đã có `File.Package` (zip có mục lục) của P6.
  - Chỉ nên làm `.dvh` khi cần chuyển cả dự án (dữ liệu, gói schema, tài liệu) sang máy khác trong một file.

## Tệp chính

- `packages/dvh-project/src`:
  - `pack.ts` (gói schema), `data.ts` (dữ liệu, thao tác, kiểm tra);
  - `source.ts` (nguồn ảo, ghi ngược), `sync.ts` (đồng bộ hai chiều, giải xung đột);
  - `packs/qlcl-xay-dung/pack.json`.
- `packages/project-store/src/store.ts`: `readProjectData`, `writeProjectData`, `MAX_PROJECT_DATA_BYTES`.
- Docs:
  - main: `dvh-project-source.ts` (mới), `dvh-link-watch.ts` (URI dự án, `notifyDvhSourceChanged`,
    `setProjectFieldWriter`), `docs-main.ts` (IPC);
  - `shared/ipc.ts` (`DvhProjectInfo`), preload;
  - renderer: `components/DvhProjectSection.tsx`, panel, CSS.
- Phụ thuộc: `@genoffice/dvh-project` trong `dependencies` của Docs và shell, `exclude` của Docs, script gốc,
  lockfile.

## Kiểm thử (vitest)

- `packages/dvh-project/tests/project.test.ts` (9):
  - D10: mã nguồn lõi không có từ vựng nghiệp vụ;
  - gói hỏng bị từ chối: `ref` không có đích, quy tắc cho kiểu lạ, khoá `id` dành riêng, trùng kiểu giữa các gói;
  - revision từng giá trị, change-set, ép kiểu, "một mỗi dự án", dữ liệu hỏng thì bắt đầu rỗng;
  - kiểm tra: liệt kê, tham chiếu thiếu, ba quy tắc của gói QLCL (lỗi và cảnh báo);
  - nguồn ảo: Field và collection có cột ID; model part hợp lệ;
  - ghi ngược: đúng bản gốc thì ghi; bản gốc cũ thì xung đột và không đổi gì; ép ghi; Field lạ thì báo thiếu;
  - **đồng bộ hai chiều qua workbook QLCL thật** (xuất rồi nhập lại như QLCL-DVH):
    - lần đầu kéo về; sau đó hai bên sửa khác trường thì trộn;
    - sửa cùng trường thì xung đột, file gửi đi giữ giá trị của đối tác;
    - dòng mới không ID nhận ID;
    - "lấy của QLCL" thì lần sau yên; "giữ của tôi" được đẩy ở lần sau;
    - xoá ba chiều và hai loại xung đột xoá.
- `packages/project-store/tests/project-data.test.ts` (2): đi–về JSON trong `<project>/data`, file hỏng đọc
  như rỗng, xoá cùng dự án, từ chối dự án lạ và id không an toàn.
- `apps/docs/tests/dvh-project.test.ts` (3):
  - service ở main: cài gói, đối tượng, đọc `dvh-project://`, ghi ngược có xung đột, thông báo thay đổi;
  - đồng bộ với file `.xlsx` trên đĩa: xung đột, file đối tác giữ giá trị (đúng kiểu số), giải rồi đẩy;
  - **tài liệu liên kết thẳng tới đối tượng dự án**: chèn Field `Project.Name`; dự án đổi thì kéo vào trang.
- Bộ đầy đủ: Docs 3.151/3.151; `project-store` 104 (2 bỏ qua); `dvh-project` 9. Typecheck Docs, Sheets,
  shell, `project-store`, `dvh-project` sạch; eslint, prettier, theme-colors, english-comments đạt.

## Để lại cho máy Windows (chưa chạy)

1. Mở một tài liệu trong một dự án → Dữ liệu dự án:
   - cài `pack.json` của QLCL, thêm `Project` và vài `WorkItem`, xem danh sách vấn đề;
   - "Liên kết tài liệu với dữ liệu dự án", chèn Field `Project.Name`;
   - sửa ở bảng đối tượng thì tài liệu (chế độ Tự động) cập nhật;
   - sửa Field trong tài liệu rồi Ghi ngược thì bảng đối tượng đổi theo.
2. Đồng bộ với một workbook QLCL thật:
   - QLCL-DVH mở được file sau đồng bộ, cột ID có giá trị;
   - sửa cùng ô ở hai bên tạo xung đột, giải được;
   - file đang mở trong Excel thì không đổi bên nào và có thông báo lỗi.
3. Hai tài liệu cùng liên kết một dự án: sửa trong tài liệu A, ghi ngược; tài liệu B ở chế độ Tự động nhận giá
   trị mới.

## Giới hạn còn lại

- Bảng đối tượng chưa có lọc, sắp xếp hay phân trang; dự án rất lớn nên sửa qua đồng bộ.
- Xung đột đồng bộ chờ giải chỉ giữ trong phiên làm việc (bản gốc thì đã lưu). Đóng ứng dụng thì đồng bộ lại sẽ
  thấy lại xung đột.
- Chưa có giao diện gỡ gói schema (đã có hàm `removePack`). Mẫu đi kèm gói mới khai báo, chưa có nút mở.
- Sheets chưa liên kết tới dữ liệu dự án; Sheets vẫn là nguồn workbook như P3.
- Chưa thử với QLCL-DVH thật; định dạng trao đổi bám theo P6.
