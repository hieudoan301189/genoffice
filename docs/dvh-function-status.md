# Trạng thái 93 hàm DVH

Nguồn tên và nhóm: `DVH-Excel/WPF/Window/DVHFunction/DVHFunctionCatalog.cs`, kiểm ngày 29/09/2026. “Đã tính” nghĩa là hàm đã được đăng ký trong Sheets và qua ca thử hiện có; chưa đồng nghĩa tương thích mọi trường hợp với Excel-DNA. “Hàm-lệnh” là hàm thay đổi workbook (xem cuối tài liệu). Bản thử đóng gói hiển thị `DVH.Evaluate("2*3*4")` và lưu cả công thức lẫn giá trị tính sẵn `2*3*4=24` vào XLSX. Excel không có add-in DVH vẫn có thể hiển thị `#NAME?` khi tính lại.

| Hàm | Nhóm | Trạng thái |
| --- | --- | --- |
| `DVH.ReadNumber` | Text | Đã tính trong DVH Office |
| `DVH.ReadTime` | Text | Đã tính trong DVH Office |
| `DVH.UnMark` | Text | Đã tính trong DVH Office |
| `DVH.FirstChar` | Text | Đã tính trong DVH Office |
| `DVH.Roman` | Text | Đã tính trong DVH Office |
| `DVH.GetString` | Text | Đã tính trong DVH Office |
| `DVH.GetNumber` | Text | Đã tính trong DVH Office |
| `DVH.IsInclude` | Text | Đã tính trong DVH Office |
| `DVH.JoinTextIF` | Text | Đã tính trong DVH Office |
| `DVH.CountText` | Text | Đã tính trong DVH Office |
| `DVH.Evaluate` | Text | Đã tính biểu thức số học; chưa hỗ trợ tham chiếu ô/hàm Excel trong chuỗi |
| `DVH.WeekdayVN` | Date | Đã tính trong DVH Office |
| `DVH.DayInMonth` | Date | Đã tính trong DVH Office |
| `DVH.LunarDay` | Date | Đã tính trong DVH Office |
| `DVH.SolarDay` | Date | Đã tính trong DVH Office |
| `DVH.WeekInMonth` | Date | Đã tính trong DVH Office |
| `DVH.SumColor` | Color | Đã tính trong DVH Office (màu nền/màu chữ, kể cả màu do DVH.FillColor tô) |
| `DVH.CountColor` | Color | Đã tính trong DVH Office |
| `DVH.ProductColor` | Color | Đã tính trong DVH Office |
| `DVH.FilterColor` | Color | Đã tính trong DVH Office (kết quả tràn) |
| `DVH.TextCom` | Comment | Hàm-lệnh: thêm note khi nhập/đổi tham số, Ctrl+Z hoàn tác |
| `DVH.DelCom` | Comment | Hàm-lệnh: xóa note trong vùng, Ctrl+Z hoàn tác |
| `DVH.GetTextCom` | Comment | Đã tính trong DVH Office |
| `DVH.PicCom` | Comment | Chưa hỗ trợ: note của DVH Office chưa chứa được hình |
| `DVH.Array.DeleteRow` | Array | Đã tính trong DVH Office |
| `DVH.Array.QSortNum` | Array | Đã tính trong DVH Office |
| `DVH.Array.QSortText` | Array | Đã tính trong DVH Office |
| `DVH.Array.InsertChild` | Array | Đã tính trong DVH Office |
| `DVH.Array.Sub` | Array | Đã tính trong DVH Office |
| `DVH.Array.GetRow` | Array | Đã tính trong DVH Office |
| `DVH.Array.GetColumn` | Array | Đã tính trong DVH Office |
| `DVH.JoinFormat` | Format | Hàm-lệnh: ghi chuỗi nối có định dạng từng ký tự vào ô đích; chạy lại khi nguồn đổi |
| `DVH.Explain` | Format | Đã tính trong DVH Office (SUM/PRODUCT/SUMIF và thay tham chiếu ô) |
| `DVH.ClearStyle` | Format | Không áp dụng: DVH Office không sinh style rác như Excel; dùng Home → Clear → Formats |
| `DVH.PrintArea` | Print | Hàm-lệnh: đặt/xóa vùng in (một vùng; nhiều vùng thì lấy vùng đầu) |
| `DVH.PrintTitle` | Print | Hàm-lệnh: đặt/xóa dòng tiêu đề in; cột tiêu đề chưa hỗ trợ |
| `DVH.RepeatTitle` | Print | Chưa hỗ trợ (cần tính phân trang để chèn dòng); dùng DVH.PrintTitle |
| `DVH.PageBreak` | Print | Chưa hỗ trợ (dãn dòng theo phân trang); dùng Page Layout → Breaks |
| `DVH.FilterSheet` | Sheet | Đã tính trong DVH Office: danh sách tràn, tô màu theo màu tab (chưa tạo hyperlink) |
| `DVH.FitRow` | Format | Hàm-lệnh: chỉnh chiều cao dòng (kể cả ô gộp); clear_formula = TRUE đổi công thức thành giá trị |
| `DVH.Font` | Format | Đã tính trong DVH Office (định dạng hiển thị qua kênh định dạng, lưu vào XLSX) |
| `DVH.FindFormat` | Format | Hàm-lệnh: tô định dạng phần chữ khớp (mặc định đỏ đậm); chạy lại khi dữ liệu đổi |
| `DVH.Table` | Format | Đã tính trong DVH Office (kết quả tràn tại ô công thức, giữ định dạng nguồn) |
| `DVH.VNS.INFO` | Standard | Chưa hỗ trợ |
| `DVH.VNS.SEARCH` | Standard | Chưa hỗ trợ |
| `DVH.Translate` | AI | Chưa hỗ trợ dạng công thức (cần dịch vụ mạng) |
| `DVH.Barcode` | Barcode | Hàm-lệnh: sinh QR/DataMatrix/Aztec/MaxiCode/Code128/EAN13/UPCA/ITF/Code39 và chèn hình tại ô đích (cần mở tệp) |
| `DVH.WEATHER.STATION` | Online | Chưa hỗ trợ |
| `DVH.WEATHER` | Online | Chưa hỗ trợ |
| `DVH.AddName` | Name Manager | Hàm-lệnh: tạo/cập nhật Name (đơn lẻ hoặc hàng loạt), Ctrl+Z hoàn tác |
| `DVH.DelName` | Name Manager | Hàm-lệnh: xóa Name theo ALL/ERR/OUT/HIDDEN/tên/mẫu * |
| `DVH.ShowName` | Name Manager | Hàm-lệnh: hiện Name ẩn |
| `DVH.CountNames` | Name Manager | Đã tính trong DVH Office |
| `DVH.JoinText` | Text | Đã tính trong DVH Office |
| `DVH.GetText_XY` | Text | Đã tính trong DVH Office |
| `DVH.SplitText` | Text | Đã tính trong DVH Office |
| `DVH.InsertText` | Text | Đã tính trong DVH Office |
| `DVH.RGBtoNum` | Color | Đã tính trong DVH Office |
| `DVH.RGBtoHex` | Color | Đã tính trong DVH Office |
| `DVH.NumToHex` | Color | Đã tính trong DVH Office |
| `DVH.GetRangeColor` | Color | Chưa hỗ trợ |
| `DVH.FillColor` | Color | Đã tính trong DVH Office (định dạng hiển thị qua kênh định dạng, lưu vào XLSX) |
| `DVH.Font.Color` | Color | Đã tính trong DVH Office (định dạng hiển thị qua kênh định dạng, lưu vào XLSX) |
| `DVH.LookUp_ManyRes` | Array | Đã tính trong DVH Office |
| `DVH.LookUp_ManyLookUp` | Array | Đã tính trong DVH Office |
| `DVH.SumHight` | Array | Đã tính trong DVH Office |
| `DVH.SumLow` | Array | Đã tính trong DVH Office |
| `DVH.Sum` | Array | Đã tính trong DVH Office |
| `DVH.NumToCol` | Tiện ích | Đã tính trong DVH Office |
| `DVH.ColToNum` | Tiện ích | Đã tính trong DVH Office |
| `DVH.LastRow` | Tiện ích | Đã tính trên workbook nạp đủ; đọc cột chứa ô/vùng tham chiếu |
| `DVH.LastCol` | Tiện ích | Đã tính trên workbook nạp đủ; đọc hàng chứa ô/vùng tham chiếu |
| `DVH.AddPicture` | Format | Có lệnh Insert → Picture; cần đối chiếu vị trí/kích thước |
| `DVH.Exchange` | Online | Chưa hỗ trợ |
| `DVH.SQL` | SQL | Chưa hỗ trợ |
| `DVH.LOOKUP` | SQL | Chưa hỗ trợ |
| `DVH.GETLIST` | SQL | Chưa hỗ trợ |
| `DVH.FILTER` | SQL | Chưa hỗ trợ |
| `DVH.DISTINCT` | SQL | Chưa hỗ trợ |
| `DVH.DSUM` | SQL | Chưa hỗ trợ |
| `DVH.DCOUNT` | SQL | Chưa hỗ trợ |
| `DVH.DAVERAGE` | SQL | Chưa hỗ trợ |
| `DVH.DMAX` | SQL | Chưa hỗ trợ |
| `DVH.DMIN` | SQL | Chưa hỗ trợ |
| `DVH.DJOIN` | SQL | Chưa hỗ trợ |
| `DVH.FIELDS` | SQL | Chưa hỗ trợ |
| `DVH.TABLES` | SQL | Chưa hỗ trợ |
| `DVH.ClearFindCache` | Tiện ích | Đã tính trong DVH Office: quên dấu vân tay FindFormat để áp lại |
| `DVH.FindCacheStats` | Tiện ích | Đã tính trong DVH Office |
| `DVH.FindFormatStatus` | Tiện ích | Đã tính trong DVH Office |
| `DVH.ClearFormatCache` | Tiện ích | Đã tính trong DVH Office: quên dấu vân tay JoinFormat để nối lại |
| `DVH.ShowTabBar` | Tiện ích | Chuyển thành lệnh |
| `DVH.TabBarStatus` | Tiện ích | Chưa hỗ trợ |

Tổng: 93; đã tính: 55 (trong đó `DVH.Evaluate` mới hỗ trợ số học, `LastRow`/`LastCol` cần workbook nạp đủ); hàm-lệnh: 11; còn lại: 27.

Các hàm của DVH-Tool (`DVH_Tool/Function/*.cs`) chưa có trong DVH Office: `DVH.PicCom`, `DVH.RepeatTitle`, `DVH.PageBreak`, `DVH.Translate`; `DVH.ClearStyle` không áp dụng. `DVH.RefreshAll` và `DVH.ShowFormulaLibrary` của DVH-Tool là lệnh, không phải hàm: dùng Data → Refresh All và DVH → DVH functions.

## Tên ngắn

Mọi hàm DVH gọi được theo hai cách: `DVH.Table(...)` và `Table(...)` (bỏ tiền tố `DVH.`; `DVH.Font.Color` → `Font.Color`). Tên ngắn trùng hàm Excel hoặc hàm có sẵn của Sheets (`SUM`, `ROMAN`, `LOOKUP`, `FILTER`, `TRANSLATE`...) không được đăng ký, chỉ dùng dạng `DVH.`. Excel có add-in DVH-Tool chỉ hiểu dạng `DVH.`; tệp dùng tên ngắn mở trong Excel sẽ hiện `#NAME?` khi tính lại (giá trị tính sẵn vẫn được lưu).

## Hàm-lệnh

DVH-Tool sửa workbook ngay trong lúc Excel tính lại (hàm macro, `QueueAsMacro`). DVH Office tách hai việc: hàm trả về dòng trạng thái như add-in, còn thay đổi được xếp hàng theo ô công thức và chạy một lần sau khi tính xong, gói trong một bước Undo:

- công thức người dùng vừa nhập/sửa chạy ở lần tính đầu; công thức có sẵn trong tệp chỉ được ghi nhớ ở lần tính đầu của workbook (tệp đã mang kết quả của nó);
- sau đó chỉ chạy lại khi “dấu vân tay” của yêu cầu đổi (tham số; với `JoinFormat`/`FindFormat` gồm cả nội dung nguồn);
- nhập lại công thức vào ô thì chạy lại; `DVH.ClearFindCache`/`DVH.ClearFormatCache` quên dấu vân tay để áp lại ở lần tính sau.

