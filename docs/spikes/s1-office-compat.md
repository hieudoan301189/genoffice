# Spike S1 – Tương thích Office cho mô hình DVH

Ngày: 01/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md), mục P0.
Máy thử: Microsoft 365 Apps (x64) 16.0.20430; DVH Office bản đóng gói `apps/shell/release/dvh-gemini-fix`.

## Kết luận

**Đi tiếp với D1, D2, D3 và D9.** Cả Word, Excel và DVH Office đều giữ được mô hình DVH nằm trong file.
Có hai việc P1 phải làm thêm (mục "Việc P1 phải làm"). Không cần đổi hướng kiến trúc.

| Quyết định                                          | Kết quả                                                                                                                                                                                                                                                                                                                       |
| --------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| D1 – mô hình trong `customXml`                      | **Đạt.** Word, Excel và DVH Office giữ nguyên part; JSON trong CDATA còn nguyên. Excel đánh số lại part (`item1` ↔ `item2`), nên **luôn tìm part theo namespace**, không theo tên file.                                                                                                                                       |
| D2 – Smart Field là SDT cấp dòng có `w:dataBinding` | **Đạt trong Word**: Word thay nội dung cũ bằng giá trị lấy từ store khi mở file, ghi ngược giá trị khi sửa content control (liên kết hai chiều), chạy với cả XPath lọc `@id` lẫn XPath theo vị trí. **DVH Office**: giữ SDT nếu không sửa đoạn chứa nó; **làm mất SDT nếu sửa chính đoạn đó** (khoảng trống đã biết, xem S2). |
| D3 – binding xlsx bằng defined name ẩn `_dvh.*`     | **Đạt ở mức file**: Excel giữ thuộc tính ẩn và dời tên khi chèn hàng; đường lưu của DVH Office cũng dời tên ẩn và Name Manager không ghi đè hay nhân đôi chúng. **Trong trình soạn thảo**: Sheets không nạp tên ẩn vào Univer, nên công thức dùng `_dvh.*` ra `#NAME?` và Univer không dời chúng trong phiên.                 |
| D9 – lịch sử nhúng trong file                       | **Đạt.** Phần lịch sử 5 MB JSON (mỗi dòng một change-set) nén còn ~350–390 KB trong file; Word mở 142 ms, Excel mở 109 ms, lưu dưới 200 ms; Word/Excel giữ nguyên toàn bộ nội dung.                                                                                                                                           |

## Cách thử

Thư mục `scripts/spikes/s1-office-compat/`:

1. `make-fixtures.mts` tạo `dvh-s1.docx` (từ mẫu trắng của docx-engine) và `dvh-s1.xlsx`, mỗi file có
   phần model (`urn:dvh-office:model:1`) và phần lịch sử (`urn:dvh-office:history:1`); bản `-big` có lịch
   sử 5 MB. Docx có hai content control gắn dữ liệu, nội dung cố ý để cũ (`STALE-A`, `STALE-B`), và một
   content control cấp khối bọc bảng. Xlsx có `_dvh.f.*` (B5), `_dvh.t.*` (A8:C11), tên hiện `TenDuAn` và
   công thức tham chiếu cả hai loại tên.
2. `office-roundtrip.ps1 -Only word|excel` mở, đọc, sửa và lưu bằng Word/Excel thật qua COM. Chỉ dùng
   tiến trình do script tạo ra, có tiến trình canh chừng tự tắt sau 2 phút.
3. `verify.mts` kiểm file Word/Excel đã lưu và đưa chúng qua docx-engine (parse + lưu kiểu vá) và
   xlsx-gateway (lưu cấu trúc + lưu Name Manager): **29 đạt, 0 lỗi, 1 khoảng trống đã biết**.
4. `app-check.mjs` mở lại hai file đó trong app DVH Office đóng gói, sửa rồi lưu.

Chạy lại:

```bash
node node_modules/tsx/dist/cli.mjs scripts/spikes/s1-office-compat/make-fixtures.mts scripts/spikes/s1-office-compat/out/run
powershell -File scripts/spikes/s1-office-compat/office-roundtrip.ps1 -OutDir scripts/spikes/s1-office-compat/out/run -Only excel
powershell -File scripts/spikes/s1-office-compat/office-roundtrip.ps1 -OutDir scripts/spikes/s1-office-compat/out/run -Only word
node node_modules/tsx/dist/cli.mjs scripts/spikes/s1-office-compat/verify.mts scripts/spikes/s1-office-compat/out/run
node scripts/spikes/s1-office-compat/app-check.mjs scripts/spikes/s1-office-compat/out/run "apps/shell/release/dvh-gemini-fix/win-unpacked/DVH Office.exe"
```

## Kết quả chi tiết

### Word

- Mở file: hai content control hiện `Dự án Cầu Bến Thủy 3` và `01/10/2026` thay cho `STALE-A/B`.
- Sửa content control "Tên dự án" thành `Sửa trong Word – Gói thầu số 7`: phần `customXml` cập nhật ngay.
- Lưu: giữ `w:tag="dvh:f:…"`, `w:dataBinding` (XPath, `storeItemID`), content control bảng
  `dvh:t:…`, cả hai phần `customXml` kèm `itemProps`; thêm `settings.xml`, `theme1.xml`… như thường lệ.

### Excel

- Tên `_dvh.*` hiện `Visible = False`; công thức `=_dvh.f.…` tính đúng sau `CalculateFull`.
- Chèn hàng 3: `_dvh.f.*` B5 → B6, `_dvh.t.*` A8:C11 → A9:C12, `TenDuAn` B5 → B6.
- Lưu: giữ thuộc tính `hidden="1"`, giữ hai phần `customXml` (đánh số lại).

### DVH Office (engine)

- docx-engine đọc file Word đã lưu, hiện giá trị trường, nhận bảng trong content control là bảng.
- Lưu sau khi sửa một đoạn khác: mọi entry `customXml/*` giống hệt từng byte; SDT trường và SDT bảng còn.
- **Khoảng trống:** lưu sau khi sửa chính đoạn chứa trường làm mất SDT cấp dòng, vì `parse.ts` duyệt
  vào nội dung `w:sdt` và chỉ giữ các run.
- xlsx-gateway: lưu có chèn 2 hàng dời `_dvh.f.*` B6 → B8, giữ `hidden`; `customXml` giống hệt từng byte.
  `applyDefinedNamesState` (lưu Name Manager) giữ nguyên tên ẩn, không nhân đôi.

### DVH Office (app đóng gói)

- Sheets: mở, sửa A1, lưu → giữ tên ẩn, `customXml`, công thức C6 `=_dvh.f.…`.
- Docs: hiện đúng giá trị trường; gõ thêm cuối tài liệu, lưu → giữ SDT trường, `dataBinding`, SDT bảng,
  `customXml`.
- Công thức **gõ mới** `=_dvh.f.…` ra `#NAME?`; `=B6` và `=TenDuAn` gõ mới tính đúng. Nguyên nhân:
  `apps/sheets/native/xlsx-engine/src/workbook.rs` (`read_defined_names`) cố ý không đưa tên ẩn vào
  trình soạn thảo, đường lưu giữ chúng nguyên văn.

## Việc P1 phải làm

1. **Node trường cấp dòng ở Docs (S2).** docx-engine phải đọc SDT `dvh:f:*` thành một run/node riêng mang
   `sdtPr` gốc và sinh lại đúng SDT khi đoạn được sửa. TipTap cần node inline tương ứng.
2. **Tên `_dvh.*` trong Sheets.** Nạp chúng vào Univer như "tên hệ thống": loại khỏi Name Manager và khỏi
   `collectDefinedNamesState`, để công thức dùng được và Univer dời chúng trong phiên; file vẫn ghi theo
   đường giữ nguyên văn + dời khi lưu như hiện nay. Cần thêm kiểm thử cho đường stream một phần.
3. **Tìm part theo namespace** (`urn:dvh-office:model:1`, `urn:dvh-office:history:1`), không theo
   `customXml/itemN.xml`.
4. **Kiểm thử hồi quy** trong `packages/docx-engine` và `packages/xlsx-gateway` cho các điều trên (lấy từ
   `verify.mts`), để lần đồng bộ upstream nào làm vỡ cũng bị phát hiện.

## Phát hiện phụ (ngoài phạm vi S1)

- **Tài liệu mới của DVH Office mở trong Word ở "Compatibility Mode"** (chế độ Word 2007): mẫu trắng của
  docx-engine (`blank.ts`) không có `word/settings.xml` với `w:compatSetting compatibilityMode=15`.
- **Sheets: công thức nạp từ file có tham chiếu tên không tính lại** khi ô nguồn đổi (D6 `=TenDuAn` giữ
  giá trị cũ), trong khi cùng công thức gõ mới trong phiên thì tính lại (G6).
- **Tự động hoá Word trên máy này:** `Document.SaveAs2(...)` gọi từ PowerShell treo vô hạn, không có hộp
  thoại, kể cả với `/a`; `SaveAs([ref]path, [ref]16)` và `Save()` chạy ~0,3 s. Add-in Office Tab vẫn nạp
  khi khởi động bằng `/a`. Chỉ ảnh hưởng script kiểm thử, không ảnh hưởng người dùng.
