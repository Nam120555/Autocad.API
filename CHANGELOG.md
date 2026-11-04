# Nhật ký thay đổi

## [1.4.0] – 04/11/2026
### Thêm
- Hộp lệnh NX (`NX`): mục **🔥 Hay dùng** liệt kê lệnh dùng nhiều nhất; khi tìm kiếm, lệnh hay dùng được xếp lên trước.
### Thay đổi
- Tách Form ra file riêng cho 5 file lớn: `07.Sampleline`, `28.TaoViewTheoTuyen`, `32.PhatSinhCoc-DoiTenCoc`, `39.TinhKhoiLuongTracNgang`, `40.NhapPointTT` (chỉ di chuyển code, không đổi logic).

## [1.3.0] – 04/11/2026
### Thêm
- Lệnh `CT_THONGKELENH`: thống kê lệnh hay dùng (lưu cục bộ, không gửi đi đâu).
- README.md đầy đủ danh mục lệnh; CHANGELOG.md.
### Thay đổi
- Tách `10.TinhKhoiLuongExcel.cs` (5.389 dòng) thành 5 file theo chức năng.
### Xoá
- ~2.250 dòng mã không còn được gọi (`_REMOVED_*`, 4 lệnh đã tắt trong `15.BeMatTruBeMat.cs`).

## [1.2.0] – 04/11/2026
### Thêm
- Ghi log ra `%AppData%\CivilTool\logs` (`NxLog`), lệnh `CT_LOG`.
- Thông tin phiên bản hiển thị khi nạp plugin và trên chữ ký form.
- `build.bat`: build Release + tạo bộ cài, phiên bản tự động.
### Sửa
- Bỏ tham chiếu AutoCAD 2027 (.NET 10) gây ~4.200 cảnh báo xung đột → còn 55.
- 143 khối `catch { }` rỗng nay ghi log thay vì nuốt lỗi im lặng; lỗi khởi động Ribbon được báo ra dòng lệnh.
- Bộ cài dùng bản Release, xoá DLL cũ khi cài đè; không nạp vào AutoCAD 2024 trở về trước.
### Xoá
- 5 lệnh vỏ rỗng (`CTP_FixTDTN`, `CTSV_CCTN`, `CTPI_PipeNX`, `CTP_DienCaoDoMB`, `CTA_NhapNTD`), `ShowForm`, `AdskGreeting`, các file rỗng.

## [1.1.0]
- Bản phát hành trước khi nâng cấp.
