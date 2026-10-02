# Ribbon DVH và các hàm còn lại của DVH-Tool

Ngày: 02/10/2026. Nhánh `dvh/main`.

## 1. Ribbon

**Docs** có thêm tab **DVH** gom toàn bộ Smart Data. Mỗi nút mở đúng mục trong bảng Smart Data, có icon riêng. Các nhóm gồm:

- dữ liệu và liên kết;
- Smart Template;
- tự động hoá: ghi macro, workflow, DVH-Script;
- dữ liệu dự án;
- lịch sử.

Nhóm Smart Data cũ ở tab Insert đã bỏ.

**Sheets** thay đổi như sau:

- Nhóm DVH Tool ở tab Home thành lưới icon, ba nút mỗi cột, giống DVH-Tool.
- Thêm tab **DVH** với các nhóm Smart Data, DVH functions, Văn bản, Hiển thị, Định dạng, Công thức và Sao chép.
- Các lệnh Home của DVH-Tool đều có mặt: đổi kiểu chữ, ẩn ô, ẩn số 0 (lưu `showZeros` vào XLSX), tham chiếu tuyệt đối/tương đối, thêm/bỏ ROUND, sao chép/dán ô hiển thị, đọc số.

## 2. Tên ngắn cho hàm DVH

Mỗi hàm gọi được theo hai cách: `DVH.Table(...)` và `Table(...)`.

- Tên ngắn là cùng một hàm, đăng ký thêm qua `dvh-function-aliases.ts`. Các hàm sau này cũng tự có tên ngắn.
- Tên ngắn trùng hàm có sẵn (`SUM`, `ROMAN`...) hoặc trùng tên Excel dành riêng (`LOOKUP`, `FILTER`, `TRANSLATE`...) bị bỏ qua. Vì vậy `DVH.Sum` và `DVH.Roman` chỉ dùng được dạng có tiền tố.
- Hộp Insert Function xếp tên ngắn vào nhóm "DVH Tool".
- Giá trị tính sẵn của công thức dùng tên ngắn vẫn được ghi vào XLSX.
- Excel có add-in DVH-Tool chỉ nhận dạng `DVH.`, nên tệp dùng tên ngắn sẽ hiện `#NAME?` nếu Excel tính lại.

## 3. Hàm mới trong Sheets

**Hàm đọc dữ liệu** (`dvh-workbook-functions.ts`):

| Hàm | Ghi chú |
| --- | --- |
| `SumColor` | Màu nền hoặc màu chữ. Màu đích là mã OLE hoặc màu của ô mẫu. Tính cả màu do `DVH.FillColor`/`DVH.Font.Color` tô. Đổi màu ô sẽ tự tính lại. |
| `CountColor` | Như SumColor. |
| `ProductColor` | Như SumColor. |
| `FilterColor` | Kết quả tràn. |
| `GetTextCom` | Nối nội dung note. |
| `CountNames` | ALL, ERR, OUT, HIDDEN. |
| `Explain` | SUM/PRODUCT/SUMIF và thay tham chiếu ô bằng giá trị. |
| `FilterSheet` | Danh sách sheet tràn, tô màu theo màu tab. Không ghi đè ô như add-in. |

**Hàm-lệnh** (`dvh-formula-commands.ts`):

- Name: `AddName`, `DelName`, `ShowName`.
- Note: `TextCom`, `DelCom`.
- In: `PrintArea`, `PrintTitle`.
- Định dạng: `FitRow`, `JoinFormat` (nối giữ định dạng từng ký tự), `FindFormat` (tô phần chữ khớp).
- Mã vạch: `Barcode` (QR, DataMatrix, Aztec, MaxiCode, Code128, EAN13, UPCA, ITF, Code39), dùng thư viện `bwip-js` (MIT), chỉ tải khi gọi lần đầu.
- Bộ nhớ đệm: `ClearFindCache`, `FindCacheStats`, `FindFormatStatus`, `ClearFormatCache`.

Cách chạy của hàm-lệnh:

- Hàm trả về dòng trạng thái giống add-in. Thay đổi chạy một lần sau khi tính xong, gói trong một bước Ctrl+Z.
- Công thức trong tệp không bị chạy lại khi mở tệp.
- Chi tiết ở `docs/dvh-function-status.md`, mục "Hàm-lệnh".

**Chưa đưa vào:**

- `PicCom`: note của Univer chưa chứa được hình.
- `RepeatTitle` và `PageBreak`: cần chèn dòng hoặc dãn dòng theo phân trang.
- `Translate`: cần dịch vụ mạng.
- `ClearStyle`: không áp dụng vì DVH Office không sinh style rác.

## 4. Kiểm thử

- `apps/sheets/tests/dvh-workbook-functions.test.ts`: 19 ca, gồm tên ngắn, màu OLE, Explain, Name, FilterSheet, thứ tự JoinFormat, nối rich text, FindFormat, hàng đợi hàm-lệnh. Năm tệp test DVH đều qua (71 ca).
- Toàn bộ test Sheets: 2949 qua, 14 lỗi. Cả 14 lỗi đều có sẵn từ trước: center-continuous, shrink-to-fit, promote-file, xlsx-alignment, xlsx-borders và pivot LibreOffice e2e.
- Đã chạy thử trên Chromium (renderer dev) và cho kết quả đúng:
  - `UNMARK`, `DVH.UnMark`, `TABLE`, `FONT.COLOR`;
  - `SUMCOLOR`, `COUNTCOLOR`, `PRODUCTCOLOR`, `FILTERCOLOR`;
  - `ADDNAME`, `DELNAME("gi*")`, `TEXTCOM`, `DELCOM`, `GETTEXTCOM`;
  - `EXPLAIN` (`=5+7*2`), `FILTERSHEET`, `PRINTAREA`, `PRINTTITLE`, `FITROW`;
  - `JOINFORMAT`, `FINDFORMAT` (chữ khớp tô đỏ đậm).

## 5. Cần kiểm trên Windows (bản đóng gói)

1. Mở một tệp XLSX, nhập `=Barcode("DVH-001";E1;0)` và xem hình QR xuất hiện tại E1. Đổi nội dung thì hình cũ được thay. Bản cloud không có tệp để thử phần chèn hình.
2. Nhập `=AddName("DonGia";B2)`, lưu, mở lại trong Excel và xem Name `DonGia` trong Name Manager.
3. Nhập `=PrintArea(A1:H40)` và `=PrintTitle(A1:H2)`, lưu, rồi xem Print Preview trong Excel.
4. Mở một tệp có sẵn công thức `DVH.AddName` và xác nhận không có Name nào bị tạo lại hay thay đổi khi mở.
