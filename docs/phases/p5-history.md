# P5 – Lịch sử theo đối tượng

Ngày: 02/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md),
mục P5. Kiểm bằng vitest trên môi trường cloud (Linux, không có Word/Excel hay bản đóng gói).

## Kết luận

- **Lịch sử nhúng trong file** như ADR D9, nay thêm:
  - `base`: model lúc lưu DVH gần nhất, để lần mở sau biết Word/Excel đã đổi gì;
  - snapshot model mỗi 100 change-set (`SNAPSHOT_EVERY`), làm mốc gộp.
- **Không mất khi tắt đột ngột**: change-set của phiên được ghi đệm ra `userData/dvh-history/<docId>.jsonl`
  (sau mỗi lần ghi, gom 1 giây). Lưu xong thì xoá. Lần mở sau còn sót thì panel báo "N thay đổi từ phiên kết
  thúc mà chưa lưu", người dùng chọn Thêm vào lịch sử hoặc Bỏ; không bao giờ tự gộp.
- **Gộp và ngưỡng**: kích thước đo bằng byte của phần lịch sử như khi lưu, so với ngưỡng 5 MB chốt ở P0.
  - Từ 80% trở lên, panel đề xuất "Gộp lịch sử". Lệnh gộp biến các change-set cũ thành một change-set
    `History.Compact` giữ `before` đầu và `after` cuối của từng đối tượng.
  - Gộp tự động theo snapshot: giữ nguyên các change-set sau snapshot cũ hơn trong hai snapshot gần nhất.
- **Phát hiện sửa bên ngoài**:
  - Docs: khi mở, hash model khác hash lúc lưu (Word ghi ngược qua data binding) → một change-set nguồn
    `external` chứa chênh lệch so với `base`.
  - Sheets: khi workbook nạp xong, ô khác model đã lưu (Excel sửa rồi lưu) → change-set `external`; model nhận
    giá trị từ ô.
  - File cũ chưa có `base`: một mục ở mức model (hash cũ → hash mới), dòng lịch sử vẫn liền.
- **Ghi lại bảng đầy đủ**: thay đổi collection ghi `data` (cột + dòng trước/sau), thiết lập bảng ghi
  `definition`, link ghi `source`/`update`. Nhờ vậy phục hồi được đúng trạng thái và panel hiện được diff theo
  ô.
  - Cập nhật tự động liên tiếp được gộp vào một change-set (giữ `before` đầu, `after` cuối) để lịch sử không
    phình theo từng lần gõ ở Sheets.
- **Bảng Lịch sử** trong panel Smart Data của Docs:
  - lọc theo đối tượng (Field, collection, bảng, link) hoặc tất cả;
  - thời gian, action, nguồn (`UI`/`AI`/`Quy trình`/`Script`/`Liên kết`/`Phục hồi`/`Bên ngoài`), người sửa;
  - diff: giá trị Field; với bảng là số dòng trước → sau, số ô đổi kèm vài ô mẫu, cột thêm/bớt, định dạng đổi.
- **Phục hồi đối tượng**: "Phục hồi tới đây" đưa riêng một Field, collection hoặc bảng về trạng thái ngay sau
  change-set đó.
  - Đi qua action `History.RestoreObject`, thay dữ liệu bảng qua `Table.ReplaceData`.
  - Bảng hiển thị được làm mới. Lần phục hồi là một change-set mới (`restore`); lịch sử không bị xoá.
- **Hoàn tác theo `txId`**: "Hoàn tác lượt chạy" (action `History.RevertTransaction`) đưa mọi đối tượng một
  transaction đã đổi về giá trị trước nó. Nếu sau đó có thay đổi khác chạm cùng đối tượng thì hỏi trước;
  đồng ý mới chạy.
- **Chính sách phát hành**: Lưu thường giữ toàn bộ. "Xuất bản sao…" trong panel lưu một bản sao (không đổi
  file đang mở) theo một trong bốn lựa chọn:
  - **Không lịch sử** (mặc định): Smart Data vẫn còn;
  - **Không có dữ liệu DVH**: bỏ phần model và lịch sử, gỡ mọi content control `dvh:f`/`dvh:t` nhưng giữ chữ,
    control khác giữ nguyên; với workbook thì bỏ tên `_dvh.*`;
  - **Lịch sử từ một ngày**;
  - **Toàn bộ lịch sử**.

  Có cảnh báo rằng lịch sử có thể chứa giá trị đã xoá. P6 (đóng gói hồ sơ) dùng lại `applyReleasePolicy`.

- **Sheets**: lưu theo cùng cấu trúc (`historyForSave`), phát hiện sửa bằng Excel, action
  `History.RestoreObject` / `History.RevertTransaction` cho Field.

## Tệp chính

- `packages/dvh-model/src/history.ts`:
  - `HistorySnapshot`, `base`, `snapshots`;
  - `historyForSave`, `compactHistory`, `compactHistoryKeeping`, `historySizeBytes`, `HISTORY_SIZE_LIMIT`;
  - `externalChangeSet`, `diffModels` mở rộng (collection/bảng/link);
  - `cellDiff`, `revertPlan`, `valueAfter`, `objectHistory`, `historyFrom`.
- Docs:
  - main: `dvh-release.ts` (chính sách phát hành, gỡ control DVH), `dvh-history-buffer.ts` (bộ đệm);
    `docs-main.ts` thêm IPC `docs:dvh-export-copy`, `docs:dvh-history-buffer`, `docs:dvh-read-history-buffer`;
  - `dvh-xlsx-write.ts` giữ `base`/snapshot khi ghi ngược vào workbook đóng;
  - renderer: `dvh-history.ts` (phục hồi, hoàn tác), `dvh-smart-data.ts` (lưu, bên ngoài, bộ đệm, gộp, ghi bảng
    đầy đủ, gộp cập nhật tự động), `dvh-actions.ts` (3 action mới), `components/DvhHistorySection.tsx` (mới),
    `App.tsx` (xuất bản sao); 30 chuỗi mới cho cả 20 ngôn ngữ.
- Sheets: `dvh-smart-data.ts` (`historyForSave`), `dvh-live.ts` (`recordExternalEdits`), `dvh-actions.ts`
  (phục hồi / hoàn tác Field).

## Kiểm thử (vitest)

- `packages/dvh-model/tests/history-p5.test.ts` (5):
  - `base` và snapshot qua serialize/parse;
  - snapshot mỗi 100 change-set và gộp sau mốc; kích thước giảm và nằm dưới 5 MB;
  - change-set `external` (có `base`, và file cũ không có `base`);
  - diff collection và diff theo ô; kế hoạch hoàn tác kèm xung đột.
- `apps/docs/tests/dvh-history-p5.test.ts` (6) – **nghiệm thu P5**:
  - phục hồi riêng collection (bảng hiển thị đúng dữ liệu cũ) rồi riêng Field, phần còn lại không đổi, không
    mục nào bị xoá khỏi lịch sử;
  - lưu rồi mở lại: lịch sử còn nguyên (cùng id, có `base`);
  - Word sửa Field trong phần model → mở lại có change-set `external` đúng giá trị trước/sau;
  - hoàn tác transaction báo xung đột, `force` mới chạy; kích thước dưới ngưỡng và gộp theo yêu cầu;
  - **bản sạch**: không lịch sử (rels và content type sạch), lịch sử từ một ngày, không DVH (control DVH gỡ,
    chữ còn, control khác giữ).
- `apps/docs/tests/dvh-history-buffer.test.ts` (2): bộ đệm đọc lại, bỏ dòng ghi dở, từ chối id lạ.
- `apps/sheets/tests/dvh-live-write.test.ts` (+2): Excel sửa ô → change-set `external`; phục hồi Field về một
  mốc và hoàn tác theo `txId` có kiểm xung đột.
- Bộ đầy đủ:
  - Docs 3.141/3.141; test DVH của Sheets 72/72; `dvh-model` 31;
  - typecheck dvh-model/docs/sheets/shell sạch; theme-colors, english-comments, eslint, prettier đạt.

## Để lại cho máy Windows (chưa chạy)

1. Lưu một báo cáo có Field và bảng liên kết, mở bằng **Word**, sửa chữ trong Field, lưu; mở lại trong DVH →
   Lịch sử có mục "Bên ngoài".
2. Như trên với workbook và **Excel**.
3. Tắt ứng dụng bằng Task Manager giữa chừng (đã sửa Field, chưa lưu) → mở lại tài liệu (đã khôi phục bản
   phục hồi, hoặc bản trên đĩa) → panel báo các thay đổi chưa lưu.
4. "Xuất bản sao… → Không có dữ liệu DVH" → mở bằng Word: không có content control DVH, chữ còn nguyên; Word
   không báo sửa file.

## Giới hạn còn lại

- Phục hồi bảng/collection mới có ở Docs. Ở Sheets mới phục hồi được Field; collection ở Sheets là vùng ô do
  người dùng sửa trực tiếp, nên ghi lại cần đi qua `Spreadsheet.SetRange` (để sang P8 khi Recorder chuẩn hoá các
  thao tác ô).
- Bảng Lịch sử dạng panel mới có ở Docs; Sheets có action nhưng chưa có giao diện riêng.
- Một collection liên kết được phục hồi sẽ bị lần cập nhật kế tiếp từ nguồn ghi đè: đó là ngữ nghĩa của liên
  kết. Muốn giữ thì chuyển liên kết sang Thủ công hoặc Ngắt liên kết.
