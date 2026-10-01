# Spike S5 – Hàm DVH trả về định dạng

Ngày: 01/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md),
Phụ lục A và nhánh T1. Kiểm trên bản đóng gói `apps/shell/release/dvh-s5` (dựng từ code của spike này).

## Kết luận

**Kênh trả định dạng chạy được. Đi tiếp với T1.** Hàm DVH trả về định dạng mà không sửa workbook khi tính
lại: định dạng hiện ngay, cập nhật theo dữ liệu và theo định dạng nguồn, không tạo mục Undo, được nướng
thành style thật khi lưu và Excel thấy đúng.

Phần **giá trị** khi lưu thì chưa đạt, và đó là khoảng trống đã có sẵn ở Sheets, không do kênh định dạng gây
ra. T1 phải xử lý nó trước khi phát hành (mục "Khoảng trống").

| Tiêu chí S5                         | Kết quả                                                                                                                                                      |
| ----------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Định dạng hiện qua interceptor      | **Đạt** – `DVH.Table` tràn kèm định dạng nguồn (nền, đậm, màu chữ, định dạng số `#,##0.000`); `DVH.Font`, `DVH.FillColor`, `DVH.Font.Color` hiện đúng        |
| Cập nhật khi tính lại               | **Đạt** – sửa giá trị nguồn làm `DVH.Table` lọc lại, dòng mới mang định dạng của nó; xoá công thức thì định dạng biến mất                                    |
| Không tạo mục Undo                  | **Đạt** – mở file xong nút Undo vẫn tắt; lịch sử chỉ có thao tác của người dùng, Undo/Redo khôi phục đúng cả định dạng                                       |
| Tự tính lại khi định dạng nguồn đổi | **Đạt sau khi bổ sung** – Univer cố ý bỏ qua việc tính lại cho thay đổi chỉ về định dạng (như Excel); spike thêm cơ chế riêng (bên dưới)                     |
| Nướng thành style khi lưu XLSX      | **Đạt** – cả công thức nạp từ file lẫn công thức gõ mới                                                                                                      |
| Excel mở thấy đúng                  | **Định dạng: đạt** (kiểm bằng Excel 365 qua COM). **Giá trị: không** – xem "Khoảng trống"                                                                    |
| In/PDF                              | **Đạt theo cấu trúc** – `print-html.ts` đọc style qua `getRange().getCellStyleData()`, chính API driver dùng để kiểm định dạng hiển thị (đi qua interceptor) |

## Thiết kế đã làm

- `apps/sheets/src/renderer/dvh-format-channel.ts`
  - `FormulaFormatStore`: định dạng theo ô sở hữu (ô công thức). Ô đơn dùng chỉ mục điểm, vùng dùng danh
    sách. Chồng lấn thì công thức công bố sau thắng.
  - Interceptor `CELL_CONTENT` ưu tiên 11 (trên NUMFMT, để định dạng số do hàm trả về được áp dụng),
    hiệu ứng Style | Value. Chỉ vẽ lại khi định dạng thực sự đổi (so chữ ký), mỗi khung hình một lần.
  - Dọn định dạng khi ô sở hữu không còn gọi hàm đó (sửa/xoá ô, chèn/xoá hàng cột, xoá sheet, huỷ workbook).
  - **Nguồn định dạng:** hàm đọc định dạng (`DVH.Table`) đăng ký vùng nguồn. Khi `SetRangeValuesMutation`
    đến từ `SetStyleCommand`/`SetBorderCommand`/`ClearSelectionFormatCommand` chạm vùng đó, kênh phát
    mutation `dvh.mutation.format-source-changed`. Mutation này đăng ký với `IActiveDirtyManagerService`
    nên đi đúng hàng đợi tính toán của Univer.
  - `withBakedFormulaFormats`: ghi định dạng đang hiện thành sửa-đổi style trong payload lưu (đè lên sửa
    đổi của người dùng ở cùng ô, như trên màn hình). Journal và Undo không thấy.
- `apps/sheets/src/renderer/dvh-format-functions.ts`: `DVH.Font`, `DVH.FillColor`, `DVH.Font.Color`,
  `DVH.Table`; các hàm thuần được kiểm thử riêng.
- Móc vào: `App.tsx` (cài/huỷ, 2 dòng) và `save-actions.ts` (nướng khi lưu).
- Kiểm thử: `apps/sheets/tests/dvh-format-functions.test.ts` (19 ca). Driver app:
  `scripts/spikes/s5-format-functions/app-check.mjs`; kiểm bằng Excel: `excel-inspect.ps1`.

### Khác với add-in DVH-Excel (cố ý)

- `DVH.Font` **đặt** kiểu chữ; add-in **đảo** kiểu mỗi lần Excel tính lại, nên kết quả phụ thuộc số lần
  tính lại.
- `DVH.Table` có dạng mới: để trống `range_result` thì kết quả tràn tại ô công thức. Dạng cũ (có
  `range_result` ở chỗ khác) chưa tính, trả `#N/A`; công thức nạp từ file vẫn hiện giá trị lưu sẵn của
  add-in nhờ cơ chế dự phòng cache.
- Dạng tràn mới add-in chưa hiểu. Nếu cần mở cùng file bằng Excel có add-in, phải bổ sung dạng này vào
  DVH-Excel hoặc dùng dạng cũ kèm "vùng do DVH quản lý".

## Khoảng trống (T1 phải xử lý)

1. **Giá trị lưu sẵn của công thức DVH nạp từ file không được làm mới khi lưu.** Đúng với mọi hàm `DVH.*`,
   không riêng hàm định dạng. Khi lưu, `formulaValues` lấy từ bộ tính IronCalc trong sidecar
   (`recalcWorkbook`); IronCalc không biết `DVH.*` nên giữ giá trị cũ (`OLD` trong ca thử). Cần lấy giá trị
   Univer đã tính cho các công thức này.
2. **Vùng tràn không được lưu.** Cả công thức nạp từ file lẫn gõ mới chỉ lưu giá trị ô đầu (K1 = `Mã`).
   Các ô tràn chỉ có style, không có `<v>`, công thức thiếu metadata mảng động, nên Excel hiện
   `=@DVH.Table(…)`. Cần ghi giá trị vùng tràn và `cm`/metadata mảng động ở gateway.
3. **Excel không có add-in tự tính lại khi mở** nên các ô `DVH.*` thành `#NAME?` (đã biết từ trước). Định
   dạng đã nướng vẫn giữ. Muốn Excel không có add-in vẫn thấy kết quả thì phải có mục 1–2, và Excel không
   được tính lại toàn bộ khi mở.
4. **Định dạng đã nướng trở thành style thật.** Nếu sau đó kết quả co lại (ít dòng hơn), ô cũ vẫn giữ định
   dạng cũ. Cần đánh dấu trong `customXml` các ô có định dạng do hàm sinh, để khi mở lại trả quyền cho hàm
   và gỡ định dạng thừa.

## Giới hạn khác của bản spike

- `DVH.Table` chưa chép ô gộp, chiều cao hàng và hình trong vùng tiêu đề (add-in có làm).
- Nếu vùng tràn bị chặn (`#SPILL!`), định dạng vẫn được công bố cho các ô đó.
- Định dạng của hàm đè định dạng gõ tay trên cùng ô (giống Conditional Formatting).
- Workbook lớn đang stream một phần dùng IronCalc nên không tính được `DVH.*` (giới hạn chung của hàm DVH).

## Phát hiện phụ

- **13 ca test Sheets đang lỗi từ trước** (có từ commit `6715d733` – DVH Home tools): `center-continuous`,
  `shrink-to-fit`, `xlsx-alignment-carry`, `xlsx-borders` (test còn kỳ vọng hành vi cũ), và 2 ca
  `promote-file-atomically`.
- **19 ca test cần `target/release/xlsx-sidecar.exe`**, trong khi bản build bằng toolchain GNU ra
  `target/x86_64-pc-windows-gnu/release`.
- `Build-DVH-Office.ps1` đã được sửa để chạy được từ ổ D (ghim `RUSTUP_TOOLCHAIN`).

## Chạy lại

```bash
node node_modules/vitest/vitest.mjs run --root apps/sheets tests/dvh-format-functions.test.ts
node scripts/spikes/s5-format-functions/app-check.mjs scripts/spikes/s5-format-functions/out/run "apps/shell/release/dvh-s5/win-unpacked/DVH Office.exe"
powershell -File scripts/spikes/s5-format-functions/excel-inspect.ps1 -Path scripts/spikes/s5-format-functions/out/run/s5-probe.xlsx -Cells "F1,G2,F3,A8,A10,A11" -Out scripts/spikes/s5-format-functions/out/run/excel.json
```
