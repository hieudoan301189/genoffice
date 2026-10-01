# Spike S3 – Tên ẩn `_dvh.*` trong Sheets

Ngày: 02/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md),
quyết định D3. Kiểm ở mức gateway (vitest) và trong bản đóng gói `apps/shell/release/dvh-s5`.

## Kết luận

**Giữ D3 (binding xlsx bằng defined name ẩn), nhưng P1 bắt buộc phải có lớp binding trong trình soạn
thảo.** Hiện Sheets không đưa tên ẩn vào Univer: đường lưu chỉ giữ nguyên chúng và dời theo chèn/xoá hàng
cột. Cách này đủ cho các thao tác cấu trúc, nhưng **hỏng với cắt/dán** và **chặn lưu khi xoá đúng ô đang
gắn**.

| Thao tác                                         | Kết quả                                                                                                        | Kiểm ở        |
| ------------------------------------------------ | -------------------------------------------------------------------------------------------------------------- | ------------- |
| Chèn hàng/cột trước ô gắn                        | ✅ tên dời đúng, vẫn ẩn                                                                                        | gateway + app |
| Chèn hàng giữa vùng bảng                         | ✅ vùng `_dvh.t.*` giãn ra                                                                                     | gateway       |
| Xoá hàng phía trên / hàng chỉ cắt vào vùng bảng  | ✅ dời / co đúng                                                                                               | gateway       |
| Di chuyển hàng (move-rows)                       | ✅ tên đi theo hàng                                                                                            | gateway       |
| Đổi tên sheet                                    | ✅ tên viết lại thành `'Dữ liệu'!$B$6`                                                                         | gateway + app |
| File 150.000 dòng, chèn hàng 0,5 giây sau khi mở | ✅ lưu được, tên dời đúng                                                                                      | app           |
| Name Manager lưu lại                             | ✅ tên ẩn giữ nguyên văn, không nhân đôi                                                                       | gateway       |
| **Cắt/dán ô đang gắn**                           | ❌ tên ẩn ở lại ô cũ; tên hiện thì đi theo (Univer theo dõi tên hiện)                                          | app           |
| **Xoá hàng/cột chứa ô đang gắn**                 | ❌ **cả lần lưu thất bại**: "A formula references the deleted range ($B$6) — deletion aborted." File không đổi | gateway + app |
| **Xoá sheet chứa ô đang gắn**                    | ❌ không cho xoá: "A workbook defined name references "Data" — deleting it is not allowed."                    | gateway       |

Nguyên nhân của ba lỗi:

- Tên ẩn chỉ tồn tại trong file. `read_defined_names` (Rust) cố ý không đưa chúng lên trình soạn thảo.
- Gateway có thao tác cấu trúc cho chèn/xoá/di chuyển hàng cột nhưng **không có** cho di chuyển vùng (cắt/dán).
- Khi một tham chiếu bị xoá, gateway chọn **chặn** (`StructuralShiftError`) thay vì ghi `#REF!`. Tên ẩn bị
  đối xử như công thức thường.

Người dùng không được báo trước. Thao tác xoá vẫn thực hiện bình thường trên màn hình, chỉ đến khi lưu mới
lỗi.

## Việc P1 phải làm

1. **Đưa tên `_dvh.*` lên trình soạn thảo như "tên hệ thống".**
   - Nạp vào Univer để Univer tự dời chúng khi cắt/dán và chèn/xoá, và để công thức dùng được (S1 cũng
     cần điều này).
   - Ẩn khỏi Name Manager.
   - Khi lưu, ghi chúng từ mô hình (kèm `hidden="1"`) thay cho bản giữ nguyên văn. Cần sửa
     `read_defined_names`, `applyDefinedNames`, `collectDefinedNamesState`, `applyDefinedNamesState`
     (nhận diện tiền tố `_dvh.`) và thứ tự lưu tách pha khi có thao tác cấu trúc.
2. **Xử lý khi ô gắn bị xoá ngay lúc thao tác, không đợi đến lúc lưu.**
   - Hỏi người dùng: huỷ thao tác, hay gỡ binding (Field chuyển trạng thái "chưa gắn", ghi lịch sử).
   - Với tên `_dvh.*`, gateway nên ghi `#REF!` thay vì chặn cả lần lưu.
3. **Xoá sheet**: gỡ các binding trên sheet đó trước (có xác nhận), rồi mới xoá.
4. Giữ bộ test `packages/xlsx-gateway/tests/dvh-hidden-names.test.ts` làm hồi quy. Hai ca "currently
   aborts/blocks" phải được sửa có chủ đích khi làm mục 2 và 3.

## Cách thử

- Gateway: `packages/xlsx-gateway/tests/dvh-hidden-names.test.ts` (10 ca).
- App: `scripts/spikes/s3-hidden-names/app-check.mjs`. Dùng `GENOFFICE_DEBUG_HOOKS=1` để thao tác qua
  Facade (`insertRowBefore`, `setName`, `deleteRow`), cắt/dán bằng Ctrl+X/Ctrl+V, đọc thông báo lỗi lưu từ
  `.app-toast`.

```bash
node node_modules/vitest/vitest.mjs run --root packages/xlsx-gateway tests/dvh-hidden-names.test.ts
node scripts/spikes/s3-hidden-names/app-check.mjs scripts/spikes/s3-hidden-names/out/run "apps/shell/release/dvh-s5/win-unpacked/DVH Office.exe"
```

## Chưa kiểm

- Không xác định được chắc chắn phần nạp nền của file 150.000 dòng còn chạy hay đã xong lúc chèn hàng:
  trạng thái `preloadComplete` không lộ ra trong bản đóng gói.
- Cắt/dán giữa hai sheet, Undo một thao tác xoá đã bị chặn khi lưu, sheet có tên dạng tham chiếu (`'A1'`).
