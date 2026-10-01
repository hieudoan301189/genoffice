# Spike S4 – Ghi bảng lớn vào Univer

Ngày: 02/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md),
P0 và P2 (hiển thị DVH.Table ở Sheets). Đo trong bản đóng gói `apps/shell/release/dvh-s5` trên máy phát triển.

## Kết luận

**Đạt cho cỡ bảng hồ sơ thông thường.** Ghi 1.000 dòng × 10 cột cả giá trị lẫn style từng ô thành **một**
mục Undo mất khoảng 0,35 giây. Lưu mất 0,7 giây. Hướng P2 "ghi một lô qua lệnh, một Undo" dùng được.

**Cần tối ưu `DVH.Table` (kênh định dạng S5) trước khi dùng cho bảng vài nghìn dòng:** chi phí tăng nhanh
hơn tuyến tính, 5.000 dòng mất 3,1 giây.

| Đo                                                        | 1.000 × 10                          | 5.000 × 10              |
| --------------------------------------------------------- | ----------------------------------- | ----------------------- |
| Ghi một lô giá trị + style (Facade `setValues`, một lệnh) | 302 ms lệnh, 351 ms đến khi vẽ xong | 1.421 ms / 1.476 ms     |
| Heap JS tăng sau khi ghi                                  | 16 MB                               | 31 MB                   |
| Một lần Undo hoàn tác cả lô                               | ✅                                  | ✅                      |
| Redo                                                      | 124 ms                              | 432 ms                  |
| `DVH.Table` tràn cùng bảng, mỗi ô một định dạng trên kênh | 339 ms có giá trị, 392 ms vẽ xong   | **3.073 ms / 3.133 ms** |
| Thời gian khung hình khi vẽ lại vùng có định dạng         | ~17 ms                              | ~16 ms                  |
| Lưu (ô có style trong journal + định dạng nướng)          | 709 ms, file 88 KB                  | 3.836 ms, file 457 KB   |
| Số `cellXfs` trong file                                   | 15                                  | 16                      |

Style được gộp tốt khi lưu: 100.000 ô chỉ sinh 16 `cellXfs`.

## Cách đo

`scripts/spikes/s4-large-table/app-check.mjs <outDir> <exe> [rows]`:

- Tạo workbook trống, mở bằng `GENOFFICE_DEBUG_HOOKS=1`.
- Ghi bảng bằng một `FRange.setValues` gồm `ICellData` có style: tiêu đề đậm nền xanh, viền mọi ô, nền xen
  kẽ, định dạng số, cứ 50 dòng một dòng chữ đỏ đậm.
- Đo đến khi hai khung hình đã vẽ, rồi kiểm Undo/Redo bằng phím tắt.
- Đặt `=DVH.Table(A1:J…,1,,)` ở L1 và đo đến khi ô tràn cuối có giá trị.
- Bấm Lưu và đo đến khi có thông báo "Saved.".

Giới hạn của phép đo:

- Một máy, một lần chạy cho mỗi cỡ, không có nhóm đối chứng.
- "Thời gian khung hình" đo bằng cách ép vẽ lại (`resize`), chưa phải cuộn thật.
- Ghi bằng Facade, chưa qua `op-executor`; cả hai đều kết thúc ở cùng mutation `set-range-values` và một mục
  Undo.

## Việc cần làm

1. **Tối ưu kênh định dạng** (nhánh T1, trước P2):
   - `FormulaFormatStore.replace` dựng chữ ký `JSON.stringify` cho toàn bộ ô mỗi lần tính lại → thay bằng
     chữ ký rẻ (băm theo id style nguồn và kích thước vùng).
   - `copiedCellStyle` chuyển đổi style cho từng ô → nhớ kết quả theo style nguồn (thường chỉ vài chục kiểu).
   - Mục tiêu: 5.000 dòng dưới 1 giây.
2. **Ngưỡng nghiệm thu cho P2**: 1.000 dòng dưới 0,5 giây (đạt), 5.000 dòng dưới 1,5 giây (ghi lô đạt,
   `DVH.Table` chưa đạt).
3. Với bảng trên 5.000 dòng, đo lại cả đường lưu (3,8 giây cho 100.000 ô có style) và cân nhắc chạy nền.
