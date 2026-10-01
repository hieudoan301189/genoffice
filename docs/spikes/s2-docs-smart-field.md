# Spike S2 – Smart Field cấp dòng trong Docs

Ngày: 02/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md),
quyết định D2. Kiểm ở mức engine/editor (vitest), trong bản đóng gói `apps/shell/release/dvh-s2` và bằng
Word 365.

## Kết luận

**Đạt. Giữ D2.** Khoảng trống S1 ghi nhận (DVH Office làm phẳng content control `dvh:f:*` khi sinh lại đoạn
chứa nó) đã được vá theo đúng tiền lệ checkbox content control:

- docx-engine giữ nguyên `w:sdtPr` gốc trên các run của trường và bọc lại đúng một `w:sdt` khi sinh đoạn;
- Docs có mark `dvhField` mang `sdtPr` qua lại giữa run và TipTap.

Còn một việc bắt buộc cho P1: **sửa chữ bên trong trường phải ghi ngược vào `customXml`**. Nếu không, Word sẽ
ghi đè chữ đó bằng giá trị trong store khi mở file.

| Kiểm tra                                                                              | Kết quả                                                                                  |
| ------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------- |
| Đọc: run trong SDT `dvh:f:*` mang `sdtFieldXml` (tag, `w:dataBinding`, id)            | ✅                                                                                       |
| Đoạn không sửa vẫn lưu giống hệt từng byte                                            | ✅                                                                                       |
| Sửa chữ quanh trường → sinh lại đúng một control quanh chữ của trường                 | ✅                                                                                       |
| Sửa chữ trong trường → control bọc chữ mới                                            | ✅                                                                                       |
| Trường gồm nhiều run khác định dạng → một control                                     | ✅                                                                                       |
| Hai trường liền nhau → hai control riêng                                              | ✅                                                                                       |
| SDT cấp dòng không phải DVH → hành vi cũ (làm phẳng), không đổi phạm vi               | ✅                                                                                       |
| Editor: gõ trước trường / ngay sau trường → nằm ngoài control (mark không tự kéo dài) | ✅ (vitest + app)                                                                        |
| Editor: gõ trong trường → nằm trong control                                           | ✅ (vitest + app)                                                                        |
| App đóng gói: lưu xong giữ 3 SDT, tag, `dataBinding`; `customXml` giống hệt từng byte | ✅                                                                                       |
| Word mở file đã sửa: hai trường vẫn gắn dữ liệu; chữ gõ ngoài trường còn nguyên       | ✅                                                                                       |
| Word mở file đã sửa: chữ gõ **trong** trường                                          | ❌ bị thay bằng giá trị trong store (đúng ngữ nghĩa data binding của Word)               |
| Toàn bộ test docx-engine (1.548) và Docs (3.107)                                      | ✅ (1 ca `protect-dialog` quá thời gian khi chạy cùng lúc với build, chạy riêng thì đạt) |

## Thay đổi

- `packages/docx-engine/src/smart-field.ts`:
  - `dvhFieldSdtPr`: nhận diện SDT có `w:tag` bắt đầu bằng `dvh:f:`;
  - `dvhFieldId`.
- `types.ts`: thuộc tính run `sdtFieldXml`.
- `parse.ts`: run bên trong trường mang `sdtFieldXml`.
- `parse-props.ts`: không gộp run trong và ngoài trường.
- `generate.ts`: `runsXml` bọc các run liền nhau có cùng `sdtFieldXml` trong một `w:sdt` (hợp lệ cả trong
  `w:ins`/`w:del`).
- Docs:
  - `editor/marks.ts`: `DvhFieldMark` (không tự kéo dài);
  - `editor/extensions.ts`: đăng ký mark;
  - `editor/convert.ts`: run ↔ mark và khoá so sánh run;
  - `ai/field-ops.ts`: chữ AI chèn không thừa hưởng mark;
  - `styles.css`: tô nổi bằng `var(--accent-soft)`, tắt khi in.
- Kiểm thử: `packages/docx-engine/tests/smart-field.test.ts` (7 ca), `apps/docs/tests/dvh-field.test.ts`
  (4 ca). Driver: `scripts/spikes/s2-docs-smart-field/app-check.mjs`, `word-inspect.ps1`.

## Việc P1 phải làm

1. **Liên kết hai chiều trong DVH Office.** Khi chữ trong trường đổi, cập nhật phần tử tương ứng trong
   `customXml` (tìm theo `dvhFieldId` và namespace model). Nếu chưa làm được thì khoá chữ trong trường và
   chỉ cho sửa qua bảng Smart Data.
2. **Trường bị xoá hết chữ.** Hiện control biến mất theo (không còn run nào mang mark). Word giữ control rỗng
   kèm placeholder. P1 cần chọn: giữ control rỗng (node inline nguyên khối) hay coi là gỡ binding (ghi lịch sử).
3. **Dán.** `parseHTML` của mark không mang `sdtPr`, nên chép/dán trường trong editor sẽ mất control, và dán
   hai lần sẽ trùng `w:id`. Cần quy tắc dán (tạo id mới, hoặc dán thành chữ thường).
4. **Phạm vi.** Mới bọc SDT có tag `dvh:f:*`. Các content control cấp dòng khác của Word vẫn bị làm phẳng khi
   sửa đoạn; có thể mở rộng cùng cơ chế sau.

## Chạy lại

```bash
node node_modules/vitest/vitest.mjs run --root packages/docx-engine tests/smart-field.test.ts
node node_modules/vitest/vitest.mjs run --root apps/docs tests/dvh-field.test.ts
node scripts/spikes/s2-docs-smart-field/app-check.mjs scripts/spikes/s2-docs-smart-field/out/run scripts/spikes/s1-office-compat/out/r4/word-saved-dvh-s1.docx "apps/shell/release/dvh-s2/win-unpacked/DVH Office.exe"
powershell -File scripts/spikes/s2-docs-smart-field/word-inspect.ps1 -Path scripts/spikes/s2-docs-smart-field/out/run/s2-edited.docx -Out scripts/spikes/s2-docs-smart-field/out/run/word.json
```
