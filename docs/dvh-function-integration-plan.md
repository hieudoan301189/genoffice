# Kiểm kê và kế hoạch đưa hàm DVH vào DVH Office

Ngày kiểm tra: 29/09/2026. Phạm vi: **Sheets** của DVH Office so với các hàm `DVH.*` trong add-in **DVH-Excel**. Phần kiểm kê bên dưới là ảnh chụp trước khi triển khai; xem trạng thái cập nhật ngay sau đây.

## Trạng thái triển khai ngày 29/09/2026

**Đã đăng ký 39/93 tên** vào bộ máy công thức Univer và danh mục Insert Function, gồm:

Ma trận trạng thái từng tên: [dvh-function-status.md](dvh-function-status.md).

- Văn bản/số: `ReadNumber`, `ReadTime`, `UnMark`, `FirstChar`, `Roman`, `GetString`, `GetNumber`, `IsInclude`, `CountText`, `Evaluate` (chỉ biểu thức số học), `GetText_XY`, `InsertText`.
- Vùng và mảng: `JoinText`, `JoinTextIF`, `SplitText`, `Array.GetRow`, `Array.GetColumn`, `Array.DeleteRow`, `Array.QSortNum`, `Array.QSortText`, `Array.InsertChild`, `Array.Sub`, `LookUp_ManyRes`, `LookUp_ManyLookUp`, `SumHight`, `SumLow`, `Sum`.
- Ngày: `WeekdayVN`, `DayInMonth`, `WeekInMonth`, `LunarDay`, `SolarDay`.
- Màu/cột: `RGBtoNum`, `RGBtoHex`, `NumToHex`, `NumToCol`, `ColToNum`, `LastRow`, `LastCol`.

Các hàm này tính ngoại tuyến, không gọi Excel-DNA/Interop; mã nguồn đối chiếu nằm trong `DVH-Excel/Function/clsFun_Text.cs`, `clsFun_Color.cs`, `clsFun_Table.cs`, `clsFun_Array.cs`, `clsFun_Evaluate.cs` và `clsFun_Lunar_Solar.cs`. Đã kiểm tra phép tính công thức thực tế, vùng tham chiếu, mảng tràn và lưu XLSX trên bản portable. Kiểm thử hiện có chưa đủ để tuyên bố tương đương hoàn toàn với mọi trường hợp biên của add-in Excel. `DVH.Evaluate` chưa hỗ trợ tham chiếu ô hoặc hàm Excel trong chuỗi. `LastRow`/`LastCol` trả `#N/A` khi workbook còn stream một phần để tránh báo sai dòng/cột cuối. 54 tên còn lại giữ trạng thái **chưa hỗ trợ/chỉ giữ cache cũ**, trừ những chức năng làm thay đổi workbook sẽ được chuyển thành lệnh ở đợt sau. Review → Translate đã dùng Gemini; `DVH.Translate` dạng công thức cần tính bất đồng bộ và cache riêng.

Trong Excel không cài DVH-Excel, các công thức `DVH.*` có thể hiển thị `#NAME?` khi Excel tính lại. Bản lưu XLSX của DVH Office giữ nguyên văn bản công thức và đã xác nhận ghi giá trị cache cho `DVH.Evaluate` do người dùng nhập; chưa chứng minh mọi hàm mới đều lưu cache đúng trong mọi tình huống.

Đợt tiếp theo: `LastRow`/`LastCol` đã đọc đúng cột/hàng của vùng tham chiếu trên workbook nạp đủ, không thay bằng kích thước UsedRange; workbook lớn đang stream một phần cần đường đọc dữ liệu riêng. `CountNames` cần đọc defined names theo loại `ALL`/`ERR`/`OUT`/`HIDDEN` và tự tính lại khi Name Manager đổi. Hàm này chưa được đăng ký để tránh trả kết quả sai.

`LunarDay` và `SolarDay` trả về số ngày serial như giá trị ngày trong Excel; đặt định dạng ô là **Date** để hiển thị `dd/mm/yyyy`. Ngày âm thuộc tháng nhuận được trả về chuỗi `Nhuận dd/mm/yyyy` như add-in gốc.

## Kết quả kiểm tra

- Hộp **Insert Function** của bản chạy đóng gói hiện liệt kê **505 mục** do Univer và ứng dụng mô tả; tìm `DVH.` được **0 mục**. Số mục không chứng minh cả 505 hàm đều cho kết quả giống Microsoft Excel.
- Danh mục `DVHFunctionCatalog.cs` của DVH-Excel có **93 tên `DVH.*`**. Mã đăng ký 91 hàm trong `Function/clsFun_*.cs` và 2 hàm thanh tab trong `ExcelTab/AddInEntry.cs`. Cả hai nơi đều nằm trong `DVH-Excel.csproj`.
- Nhóm **DVH Tool** vừa đưa vào thẻ Home là các **lệnh định dạng** (căn giữa không gộp, căn đều, thu chữ, giãn dòng, lề văn bản, đường viền). Nó không đăng ký hàm tính `DVH.*`.
- Sheets dùng Univer để tính công thức khi workbook được nạp đủ; workbook lớn có thể chỉ hiển thị kết quả được lưu sẵn hoặc tính phần phụ thuộc nhỏ. Nếu thiếu hàm, mã hiện tại ưu tiên giữ kết quả cache có sẵn để tránh `#NAME?`, nhưng cache cũ không phải kết quả tính mới. Vì vậy **chưa thể tuyên bố tương thích đầy đủ với Excel**.

Nguồn kiểm tra: `DVH-Excel/WPF/Window/DVHFunction/DVHFunctionCatalog.cs`, `DVH-Excel/Function/`, `DVH-Excel/ExcelTab/AddInEntry.cs`, `apps/sheets/src/renderer/function-catalog.ts`, `apps/sheets/src/renderer/function-registry-probe.ts`, `apps/sheets/src/renderer/univer-sync.ts`, `apps/sheets/docs/compatibility.md`. Bài chạy thực tế: `scripts/drivers/driver.function-audit.mjs`.

## Kiểm kê theo nhóm

| Nhóm trong DVH-Tool | Số tên | Cách chuyển dự kiến |
| --- | ---: | --- |
| Text | 15 | Hàm thuần trước; `DVH.Evaluate` cần bộ phân tích biểu thức an toàn. |
| Date | 5 | Hàm thuần, kiểm ngày nhuận, lịch âm và số serial Excel. |
| Array | 12 | Hàm mảng, kiểm kích thước tràn, thứ tự và lỗi. |
| Color | 10 | Tách chuyển đổi mã màu khỏi hàm đọc định dạng và lệnh đổi màu. |
| Comment | 4 | Hàm đọc chú thích; thao tác thêm/xóa thành lệnh có Undo. |
| Format | 8 | Tách hàm trả kết quả khỏi lệnh sửa ô, hình và vùng dữ liệu. |
| Name Manager | 4 | Dùng mô hình defined names hiện có; thao tác thêm/xóa là lệnh. |
| Online | 3 | Chỉ triển khai sau khi có cấu hình mạng, cache và cơ chế đồng ý. |
| Print | 4 | Lệnh thiết lập trang in, không chạy ngầm khi tính lại công thức. |
| Sheet | 1 | Lệnh lọc sheet. |
| SQL | 13 | Bộ truy vấn riêng cho vùng bảng tính; nguồn tệp ngoài cần quyền truy cập rõ ràng. |
| Standard | 2 | Tra cứu TCVN qua nguồn dữ liệu/mạng có cache. |
| AI | 1 | Review → Translate dùng Gemini; `DVH.Translate` dạng công thức cần tính bất đồng bộ và cache riêng. |
| Barcode | 1 | Xác định đầu ra là giá trị ô hay hình/mã vạch trước khi triển khai. |
| Tiện ích | 10 | Hàm chuyển số cột trước; phần còn lại theo ngữ cảnh workbook hoặc lệnh quản lý cache/tab. |
| **Tổng** | **93** | Chuyển theo hành vi, không sao chép nguyên Excel-DNA/Interop. |

## Thứ tự triển khai

1. **Lập ma trận chuẩn cho 93 tên.** Ghi tên, cú pháp, tham số mặc định, kiểu trả về, lỗi, ví dụ thực tế, có đọc ô/định dạng hay làm thay đổi workbook, có dùng mạng/tệp/AI hay không. Dùng chính add-in cũ và tệp XLSX mẫu để chốt kết quả tham chiếu. Kiểm tra phiên bản Excel và dấu phân cách theo locale. Đầu ra: một hàng trạng thái và bộ ca thử cho từng tên.
2. **Đăng ký lớp hàm DVH trong Sheets.** Thử trước `=DVH.UnMark("Tiếng Việt")`, gồm gợi ý trong Insert Function, tính lại khi đầu vào đổi, lưu/mở XLSX và hiển thị lỗi. Đăng ký chung vào bộ máy Univer và danh mục mô tả; không tạo bộ tính riêng cạnh tranh với nó. Xác nhận dấu chấm trong tên hàm, tên hàm trong XML và giá trị cache qua Excel/DVH Office.
3. **Đưa nhóm hàm thuần, dùng ngoại tuyến.** Ưu tiên `ReadNumber`, `ReadTime`, `UnMark`, `FirstChar`, `Roman`, `GetString`, `GetNumber`, `IsInclude`, `JoinTextIF`, `CountText`, `JoinText`, `GetText_XY`, `SplitText`, `InsertText`, `DayInMonth`, `WeekdayVN`, `WeekInMonth`, `RGBtoNum`, `RGBtoHex`, `NumToHex`, `NumToCol`, `ColToNum`. Đối chiếu ca tiếng Việt, số âm/thập phân, ngày nhuận, ô trống và lỗi với add-in cũ. Chỉ đưa hàm đã qua kiểm thử vào danh mục.
4. **Thêm hàm cần dữ liệu workbook.** Ưu tiên mảng, tra cứu nhiều kết quả, lịch âm/dương, đọc màu, `LastRow`/`LastCol`, đọc tên/chú thích. Thiết kế đầu vào vùng dữ liệu và định dạng thống nhất cho workbook nạp đủ lẫn workbook stream; giới hạn kích thước và bảo đảm công thức tính lại khi dữ liệu hoặc định dạng liên quan đổi. Kiểm thử spill, vùng chồng lấn, tham chiếu khác sheet, Undo và XLSX round-trip.
5. **Chuyển chức năng có tác dụng phụ thành lệnh rõ ràng.** `FillColor`, `Font.Color`, `TextCom`, `DelCom`, `PicCom`, `AddName`, `DelName`, `PrintArea`, `PrintTitle`, `RepeatTitle`, `PageBreak`, `FilterSheet`, `ClearStyle`, `FitRow`, `FindFormat`, `JoinFormat`, `Table`, `AddPicture`, các lệnh cache/tab và các chức năng tương tự phải qua UI/lệnh có Undo và lưu file. Không tự sửa workbook mỗi khi máy tính lại công thức. Với công thức cũ, giữ nguyên văn bản và cache, đồng thời báo rõ chức năng nào chưa chạy được.
6. **Làm SQL và nguồn ngoài sau cùng.** `DVH.SQL` cùng 12 hàm SQL bọc quanh nó cần quy định bảng nguồn, named range, quyền đọc `.xlsx/.csv/.db/.mdb`, giới hạn tài nguyên, kết quả mảng và lỗi tương ứng. `.mdb`/ACE của bản Excel-DNA không thể mặc nhiên mang sang ứng dụng đa nền tảng. Các hàm thời tiết, tỷ giá, TCVN và dịch thuật cần chính sách dữ liệu, cache, timeout và lựa chọn bật riêng. Công thức `DVH.Translate` cần tính bất đồng bộ để không làm treo Sheets hoặc gọi API lại mỗi lần tính lại.

## Điều kiện nghiệm thu từng đợt

- Công thức cho cùng kết quả với add-in cũ trên bộ ca chuẩn, kể cả lỗi và giá trị rỗng; hàm có tham chiếu cập nhật khi ô nguồn thay đổi.
- Lưu XLSX rồi mở lại trong DVH Office giữ công thức, kiểu dữ liệu và kết quả; mở trong Excel không báo sửa tệp. Excel **không có add-in DVH** vẫn có thể báo `#NAME?` cho hàm `DVH.*` khi tính lại, nên phải kiểm tra và truyền đạt giới hạn này trước khi phát hành.
- Workbook lớn không bị treo vì vùng quá rộng, hàm SQL hay dịch vụ mạng; hàm chưa hỗ trợ không âm thầm trả kết quả cũ như thể vừa được tính.
- Lệnh làm thay đổi tài liệu có Undo, được ghi vào XLSX, không thực thi ngầm do recalculation; nguồn DVH-Excel gốc vẫn độc lập và không trở thành dependency chạy của DVH Office.

Sau mỗi đợt, cập nhật ma trận trạng thái thành **đã tính và kiểm chứng / chỉ giữ cache / chưa hỗ trợ / chuyển thành lệnh**. Chỉ tuyên bố tương thích Excel theo các hàm và phiên bản đã thực sự kiểm tra.
