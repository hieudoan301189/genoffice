# Trạng thái 93 hàm DVH

Nguồn tên và nhóm: `DVH-Excel/WPF/Window/DVHFunction/DVHFunctionCatalog.cs`, kiểm ngày 29/09/2026. “Đã tính” nghĩa là hàm đã được đăng ký trong Sheets và qua ca thử hiện có; chưa đồng nghĩa tương thích mọi trường hợp với Excel-DNA. “Chuyển thành lệnh” là hướng triển khai, chưa được bàn giao như lệnh mới. Bản thử đóng gói hiển thị `DVH.Evaluate("2*3*4")` và lưu cả công thức lẫn giá trị tính sẵn `2*3*4=24` vào XLSX. Excel không có add-in DVH vẫn có thể hiển thị `#NAME?` khi tính lại.

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
| `DVH.SumColor` | Color | Chưa hỗ trợ |
| `DVH.CountColor` | Color | Chưa hỗ trợ |
| `DVH.ProductColor` | Color | Chưa hỗ trợ |
| `DVH.FilterColor` | Color | Chưa hỗ trợ |
| `DVH.TextCom` | Comment | Chuyển thành lệnh |
| `DVH.DelCom` | Comment | Có lệnh gần tương đương: Review → Delete Note; cần đối chiếu chi tiết |
| `DVH.GetTextCom` | Comment | Chưa hỗ trợ |
| `DVH.PicCom` | Comment | Chuyển thành lệnh |
| `DVH.Array.DeleteRow` | Array | Đã tính trong DVH Office |
| `DVH.Array.QSortNum` | Array | Đã tính trong DVH Office |
| `DVH.Array.QSortText` | Array | Đã tính trong DVH Office |
| `DVH.Array.InsertChild` | Array | Đã tính trong DVH Office |
| `DVH.Array.Sub` | Array | Đã tính trong DVH Office |
| `DVH.Array.GetRow` | Array | Đã tính trong DVH Office |
| `DVH.Array.GetColumn` | Array | Đã tính trong DVH Office |
| `DVH.JoinFormat` | Format | Chuyển thành lệnh |
| `DVH.Explain` | Format | Chưa hỗ trợ |
| `DVH.ClearStyle` | Format | Có lệnh Home → Clear → Formats; cần đối chiếu phạm vi |
| `DVH.PrintArea` | Print | Có lệnh Page Layout → Print Area |
| `DVH.PrintTitle` | Print | Có lệnh Page Layout → Print Titles; cần đối chiếu tham số |
| `DVH.RepeatTitle` | Print | Có lệnh Page Layout → Print Titles; cần đối chiếu tham số |
| `DVH.PageBreak` | Print | Có lệnh Page Layout → Breaks |
| `DVH.FilterSheet` | Sheet | Chuyển thành lệnh |
| `DVH.FitRow` | Format | Có lệnh Home → Format → AutoFit Row Height |
| `DVH.Font` | Format | Có lệnh Home → Font; cần đối chiếu tham số |
| `DVH.FindFormat` | Format | Chuyển thành lệnh |
| `DVH.Table` | Format | Có lệnh Home/Insert → Format as Table; cần đối chiếu chi tiết |
| `DVH.VNS.INFO` | Standard | Chưa hỗ trợ |
| `DVH.VNS.SEARCH` | Standard | Chưa hỗ trợ |
| `DVH.Translate` | AI | Có lệnh Review → Translate dùng Gemini; cú pháp công thức `DVH.Translate` chưa hỗ trợ |
| `DVH.Barcode` | Barcode | Chuyển thành lệnh |
| `DVH.WEATHER.STATION` | Online | Chưa hỗ trợ |
| `DVH.WEATHER` | Online | Chưa hỗ trợ |
| `DVH.AddName` | Name Manager | Có lệnh Formulas → Name Manager → Add |
| `DVH.DelName` | Name Manager | Có lệnh Formulas → Name Manager → Delete |
| `DVH.ShowName` | Name Manager | Có lệnh Formulas → Name Manager |
| `DVH.CountNames` | Name Manager | Chưa hỗ trợ |
| `DVH.JoinText` | Text | Đã tính trong DVH Office |
| `DVH.GetText_XY` | Text | Đã tính trong DVH Office |
| `DVH.SplitText` | Text | Đã tính trong DVH Office |
| `DVH.InsertText` | Text | Đã tính trong DVH Office |
| `DVH.RGBtoNum` | Color | Đã tính trong DVH Office |
| `DVH.RGBtoHex` | Color | Đã tính trong DVH Office |
| `DVH.NumToHex` | Color | Đã tính trong DVH Office |
| `DVH.GetRangeColor` | Color | Chưa hỗ trợ |
| `DVH.FillColor` | Color | Có lệnh Home → Fill Color |
| `DVH.Font.Color` | Color | Có lệnh Home → Font Color |
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
| `DVH.ClearFindCache` | Tiện ích | Chuyển thành lệnh |
| `DVH.FindCacheStats` | Tiện ích | Chưa hỗ trợ |
| `DVH.FindFormatStatus` | Tiện ích | Chưa hỗ trợ |
| `DVH.ClearFormatCache` | Tiện ích | Chuyển thành lệnh |
| `DVH.ShowTabBar` | Tiện ích | Chuyển thành lệnh |
| `DVH.TabBarStatus` | Tiện ích | Chưa hỗ trợ |

Tổng: 93; đã tính: 39 (trong đó `DVH.Evaluate` mới hỗ trợ số học, `LastRow`/`LastCol` cần workbook nạp đủ); còn lại: 54.

