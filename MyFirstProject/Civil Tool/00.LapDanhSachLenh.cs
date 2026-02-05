// ==============================================================================
// BẢNG THỐNG KÊ DANH SÁCH LỆNH CIVIL TOOL
// ==============================================================================
// Ngày tạo: 2026-02-05
// Mô tả: File tổng hợp danh sách tất cả lệnh trong Civil Tool
// ==============================================================================

/*
╔══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
║                                              BẢNG 1: DANH SÁCH LỆNH ĐẦY ĐỦ                                                                               ║
╚══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ NHÓM 1: CORRIDOR (01.Corridor.cs)                                                                                                                        │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ STT │ Tên lệnh                          │ Mô tả                                              │ Cách dùng                                                 │
├─────┼────────────────────────────────────┼────────────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
│ 1   │ CTC_AddAllSection                  │ Thêm tất cả Section vào Corridor                   │ Gõ lệnh → Chọn Corridor → Chọn nhóm cọc                   │
│ 2   │ CTC_TaoCooridor_DuongDoThi_RePhai  │ Tạo Corridor đường đô thị rẽ phải                  │ Gõ lệnh → Nhập số corridor → Chọn các target              │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ NHÓM 2: PARCEL (02.Parcel.cs)                                                                                                                            │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ STT │ Tên lệnh                          │ Mô tả                                              │ Cách dùng                                                 │
├─────┼────────────────────────────────────┼────────────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
│ 1   │ CTPA_TaoParcel_CacLoaiNha          │ Tạo Parcel từ các polyline nhà                     │ Gõ lệnh → Chọn các polyline                               │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ NHÓM 3: PIPE AND STRUCTURES (04.PipeAndStructures.cs)                                                                                                    │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ STT │ Tên lệnh                          │ Mô tả                                              │ Cách dùng                                                 │
├─────┼────────────────────────────────────┼────────────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
│ 1   │ CTPI_ThayDoi_DuongKinhCong         │ Thay đổi đường kính ống cống                       │ Gõ lệnh → Chọn các ống → Chọn kích thước mới              │
│ 2   │ CTPI_ThayDoi_MatPhangRef_Cong      │ Thay đổi mặt phẳng Reference cho ống               │ Gõ lệnh → Chọn Surface → Chọn các ống                     │
│ 3   │ CTPI_ThayDoi_DoanDocCong           │ Thiết lập độ dốc đoạn cống                         │ Gõ lệnh → Chọn ống từ thượng lưu → Nhập cao độ            │
│ 4   │ CTPI_BangCaoDo_TuNhienHoThu        │ Tạo bảng cao độ tự nhiên hố thu                    │ Gõ lệnh → Chọn 2 Surface → Chọn các hố thu                │
│ 5   │ CTPI_XoayHoThu_Theo2diem           │ Xoay hố thu theo 2 điểm                            │ Gõ lệnh → Chọn hố thu → Chọn 2 điểm làm căn cứ            │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ NHÓM 4: POINT (05.Point.cs)                                                                                                                              │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ STT │ Tên lệnh                          │ Mô tả                                              │ Cách dùng                                                 │
├─────┼────────────────────────────────────┼────────────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
│ 1   │ CTPO_TaoCogoPoint_CaoDo_FromSurface│ Tạo Cogo Point lấy cao độ từ Surface               │ Gõ lệnh → Chọn Surface → Click các vị trí                 │
│ 2   │ CTPO_TaoCogoPoint_CaoDo_Elevationspot│ Chuyển Elevation Spot thành Cogo Point           │ Gõ lệnh → Chọn Surface                                    │
│ 3   │ CTPO_UpdateAllPointGroup           │ Cập nhật tất cả Point Group                        │ Gõ lệnh (tự động chạy)                                    │
│ 4   │ CTPO_CreateCogopointFromText       │ Tạo Cogo Point từ Text có cao độ                   │ Gõ lệnh → Chọn các Text                                   │
│ 5   │ CTPO_An_CogoPoint                  │ Ẩn các Cogo Point đã chọn                          │ Gõ lệnh → Chọn các điểm cần ẩn                            │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ NHÓM 5: PROFILE VÀ PROFILE VIEW (06.ProfileAndProfileView.cs)                                                                                            │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ STT │ Tên lệnh                          │ Mô tả                                              │ Cách dùng                                                 │
├─────┼────────────────────────────────────┼────────────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
│ 1   │ CTP_VeTracDoc_TuNhien              │ Vẽ trắc dọc tự nhiên từ Surface                    │ Gõ lệnh → Chọn Surface → Chọn tuyến → Chọn vị trí đặt     │
│ 2   │ CTP_VeTracDoc_TuNhien_TatCaTuyen   │ Vẽ trắc dọc TN cho tất cả tuyến                    │ Gõ lệnh → Chọn Surface → Chọn điểm đặt → Nhập khoảng cách │
│ 3   │ CTP_Fix_DuongTuNhien_TheoCoc       │ Sửa đường TN theo vị trí cọc                       │ Gõ lệnh → Chọn Profile TN                                 │
│ 4   │ CTP_GanNhanNutGiao_LenTracDoc      │ Gắn nhãn nút giao lên trắc dọc                     │ Gõ lệnh → Chọn Cogo Point → Nhập sai số → Chọn trắc dọc   │
│ 5   │ CTP_TaoCogoPointTuPVI              │ Tạo Cogo Point từ điểm PVI trên trắc dọc           │ Gõ lệnh → Chọn Profile                                    │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ NHÓM 6: SAMPLELINE / CỌC (07.Sampleline.cs)                                                                                                              │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ STT │ Tên lệnh                          │ Mô tả                                              │ Cách dùng                                                 │
├─────┼────────────────────────────────────┼────────────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
│ 1   │ CTS_DoiTenCoc                      │ Đổi tên cọc (chọn từng cọc)                        │ Gõ lệnh → Nhập prefix → Nhập số bắt đầu → Chọn cọc        │
│ 2   │ CTS_DoiTenCoc2                     │ Đổi tên cọc theo số thập phân                      │ Gõ lệnh → Nhập prefix → Nhập số bắt đầu → Chọn cọc        │
│ 3   │ CTS_DoiTenCoc3                     │ Đổi tên cọc theo tên lý trình                      │ Gõ lệnh → Nhập prefix → Chọn nhóm cọc                     │
│ 4   │ CTS_DoiTenCoc_fromCogoPoint        │ Đổi tên cọc từ tên Cogo Point                      │ Gõ lệnh → Chọn Cogo Point → Nhập sai số                   │
│ 5   │ CTS_DoiTenCoc_TheoThuTu            │ Đổi tên cọc theo thứ tự lý trình                   │ Gõ lệnh → Nhập prefix → Chọn nhóm cọc                     │
│ 6   │ CTS_DoiTenCoc_H                    │ Đổi tên cọc với hậu tố H (hầm)                     │ Gõ lệnh → Chọn cọc                                        │
│ 7   │ CTS_TaoBang_ToaDoCoc               │ Tạo bảng tọa độ cọc                                │ Gõ lệnh → Chọn tuyến → Nhập start/end station             │
│ 8   │ CTS_TaoBang_ToaDoCoc2              │ Tạo bảng tọa độ cọc (cách 2)                       │ Gõ lệnh → Chọn tuyến → Chọn nhóm cọc                      │
│ 9   │ CTS_TaoBang_ToaDoCoc3              │ Tạo bảng tọa độ cọc (cách 3)                       │ Gõ lệnh → Chọn tuyến                                      │
│ 10  │ AT_UPdate2Table                    │ Cập nhật bảng tọa độ                               │ Gõ lệnh → Chọn bảng                                       │
│ 11  │ CTS_ChenCoc_TrenTracDoc            │ Chèn cọc hiển thị trên trắc dọc                    │ Gõ lệnh → Chọn Profile View → Chọn vị trí                 │
│ 12  │ CTS_CHENCOC_TRENTRACNGANG          │ Chèn cọc hiển thị trên trắc ngang                  │ Gõ lệnh → Chọn Section View → Chọn vị trí                 │
│ 13  │ CTS_PhatSinhCoc                    │ Phát sinh cọc tự động                              │ Gõ lệnh → Chọn tuyến → Nhập khoảng cách → Chọn nhóm cọc   │
│ 14  │ CTS_PhatSinhCoc_ChiTiet            │ Phát sinh cọc chi tiết                             │ Gõ lệnh → Chọn tuyến → Nhập các thông số                  │
│ 15  │ CTS_PhatSinhCoc_theoKhoangDelta    │ Phát sinh cọc theo khoảng Delta                    │ Gõ lệnh → Chọn tuyến → Nhập Delta                         │
│ 16  │ CTS_PhatSinhCoc_TuCogoPoint        │ Phát sinh cọc từ Cogo Point                        │ Gõ lệnh → Chọn Point Group                                │
│ 17  │ CTS_PhatSinhCoc_TheoBang           │ Phát sinh cọc theo bảng Excel                      │ Gõ lệnh → Chọn tuyến → Chọn file Excel                    │
│ 18  │ CTS_DichCoc_TinhTien               │ Dịch cọc tịnh tiến (nhập giá trị)                  │ Gõ lệnh → Chọn nhóm cọc → Nhập khoảng dịch                │
│ 19  │ CTS_DichCoc_TinhTien40             │ Dịch cọc tịnh tiến 40m                             │ Gõ lệnh → Chọn nhóm cọc                                   │
│ 20  │ CTS_DichCoc_TinhTien_20            │ Dịch cọc tịnh tiến 20m                             │ Gõ lệnh → Chọn nhóm cọc                                   │
│ 21  │ CTS_Copy_NhomCoc                   │ Sao chép nhóm cọc                                  │ Gõ lệnh → Chọn nhóm cọc nguồn → Chọn tuyến đích           │
│ 22  │ CTS_DongBo_2_NhomCoc               │ Đồng bộ 2 nhóm cọc                                 │ Gõ lệnh → Chọn nhóm nguồn → Chọn nhóm đích                │
│ 23  │ CTS_DongBo_2_NhomCoc_TheoDoan      │ Đồng bộ 2 nhóm cọc theo đoạn                       │ Gõ lệnh → Chọn nhóm → Nhập station đầu/cuối               │
│ 24  │ CTS_Copy_BeRong_sampleLine         │ Sao chép bề rộng Sample Line                       │ Gõ lệnh → Chọn cọc mẫu → Chọn các cọc đích                │
│ 25  │ CTS_Thaydoi_BeRong_sampleLine      │ Thay đổi bề rộng Sample Line                       │ Gõ lệnh → Chọn cọc → Nhập bề rộng trái/phải               │
│ 26  │ CTS_Offset_BeRong_sampleLine       │ Offset bề rộng Sample Line                         │ Gõ lệnh → Chọn cọc → Nhập giá trị offset                  │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ NHÓM 7: SECTION VIEW / TRẮC NGANG (08.Sectionview.cs)                                                                                                    │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ STT │ Tên lệnh                          │ Mô tả                                              │ Cách dùng                                                 │
├─────┼────────────────────────────────────┼────────────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
│ 1   │ CTSV_VeTracNgang_ThietKe           │ Vẽ trắc ngang thiết kế                             │ Gõ lệnh → Chọn tuyến → Chọn nhóm cọc → Chọn vị trí đặt    │
│ 2   │ CTSV_VeTatCaTracNgang_ThietKe      │ Vẽ tất cả trắc ngang thiết kế                      │ Gõ lệnh → Chọn các tuyến → Chọn vị trí đặt                │
│ 3   │ CTSV_ChuyenDoi_TNTK_TNTN           │ Chuyển đổi style TN thiết kế sang TN tự nhiên      │ Gõ lệnh → Chọn Section View                               │
│ 4   │ CTSV_DanhCap                       │ Đánh cấp trắc ngang                                │ Gõ lệnh → Chọn Section View → Nhập các thông số           │
│ 5   │ CTSV_DanhCap_XoaBo                 │ Xóa bỏ đánh cấp                                    │ Gõ lệnh → Chọn Section View                               │
│ 6   │ CTSV_DanhCap_VeThem                │ Vẽ thêm đánh cấp                                   │ Gõ lệnh → Chọn Section View                               │
│ 7   │ CTSV_DanhCap_VeThem1               │ Vẽ thêm đánh cấp (cách 1)                          │ Gõ lệnh → Chọn Section View                               │
│ 8   │ CTSV_DanhCap_VeThem2               │ Vẽ thêm đánh cấp (cách 2)                          │ Gõ lệnh → Chọn Section View                               │
│ 9   │ CTSV_DanhCap_CapNhat               │ Cập nhật đánh cấp                                  │ Gõ lệnh → Chọn Section View                               │
│ 10  │ CTSV_ThemVatLieuTrenCatNgang       │ Thêm vật liệu trên mặt cắt ngang                   │ Gõ lệnh → Chọn Section View → Chọn vật liệu               │
│ 11  │ CTSV_ThayDoi_MSS_MinMax            │ Thay đổi Min/Max của Multiple Section Style        │ Gõ lệnh → Chọn Section View → Nhập giá trị                │
│ 12  │ CTSV_ThayDoi_GioiHan_TraiPhai      │ Thay đổi giới hạn trái/phải Section View           │ Gõ lệnh → Chọn Section View → Nhập giá trị                │
│ 13  │ CTSV_ThayDoi_KhungIn               │ Thay đổi khung in Section View                     │ Gõ lệnh → Chọn Section View                               │
│ 14  │ CTSV_KhoaCatNgang_AddPoint         │ Khóa mặt cắt ngang và thêm điểm                    │ Gõ lệnh → Chọn Section View → Click các điểm              │
│ 15  │ CTSV_FitKhungIn                    │ Fit khung in tự động                               │ Gõ lệnh → Chọn Section View                               │
│ 16  │ CTSV_FitKhungIn_55Top              │ Fit khung in 5x5 từ đỉnh                           │ Gõ lệnh → Chọn Section View                               │
│ 17  │ CTSV_FitKhungIn_510Top             │ Fit khung in 5x10 từ đỉnh                          │ Gõ lệnh → Chọn Section View                               │
│ 18  │ CTSV_An_DuongDiaCHAT               │ Ẩn đường địa chất                                  │ Gõ lệnh → Chọn Section View                               │
│ 19  │ CTSV_HieuChinh_Section_Static      │ Hiệu chỉnh Section tĩnh                            │ Gõ lệnh → Chọn Section                                    │
│ 20  │ CTSV_HieuChinh_Section_Dynamic     │ Hiệu chỉnh Section động                            │ Gõ lệnh → Chọn Section                                    │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ NHÓM 8: SURFACE (09.Surfaces.cs)                                                                                                                         │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ STT │ Tên lệnh                          │ Mô tả                                              │ Cách dùng                                                 │
├─────┼────────────────────────────────────┼────────────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
│ 1   │ CTS_TaoSpotElevation_OnSurface_TaiTim│ Tạo Spot Elevation trên Surface tại tim tuyến   │ Gõ lệnh → Chọn Surface → Chọn tuyến                       │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ NHÓM 9: PROPERTY SETS (10.Property Sets.cs)                                                                                                              │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ STT │ Tên lệnh                          │ Mô tả                                              │ Cách dùng                                                 │
├─────┼────────────────────────────────────┼────────────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
│ 1   │ AT_Solid_Set_PropertySet           │ Gán Property Set cho Solid                         │ Gõ lệnh → Chọn Solid                                      │
│ 2   │ AT_Solid_Show_Info                 │ Hiển thị thông tin Solid                           │ Gõ lệnh → Chọn Solid                                      │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ NHÓM 10: TÍNH KHỐI LƯỢNG EXCEL (10.TinhKhoiLuongExcel.cs)                                                                                                │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ STT │ Tên lệnh                          │ Mô tả                                              │ Cách dùng                                                 │
├─────┼────────────────────────────────────┼────────────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
│ 1   │ CTSV_CaiDatBang                    │ Cài đặt định dạng bảng khối lượng                  │ Gõ lệnh → Form cài đặt hiện ra                            │
│ 2   │ CTSV_Taskbar                       │ Mở thanh công cụ Taskbar                           │ Gõ lệnh                                                   │
│ 3   │ CTSV_XuatCad                       │ Xuất bảng khối lượng ra CAD                        │ Gõ lệnh → Chọn các tuyến → Chọn vị trí đặt bảng           │
│ 4   │ CTSV_SoSanhSurface                 │ So sánh 2 Surface                                  │ Gõ lệnh → Chọn 2 Surface → Xem kết quả so sánh            │
│ 5   │ CTSV_LayDienTichTuSectionView      │ Lấy diện tích từ Section View                      │ Gõ lệnh → Chọn Section View                               │
│ 6   │ CTSV_LayKhoiLuongTracNgang         │ Lấy khối lượng từ trắc ngang                       │ Gõ lệnh → Chọn tuyến → Chọn nhóm cọc                      │
│ 7   │ CTSV_XuatSectionArea               │ Xuất diện tích Section ra Excel                    │ Gõ lệnh → Chọn tuyến → Chọn nhóm cọc                      │
│ 8   │ CTSV_TinhKLKetHop                  │ Tính khối lượng kết hợp nhiều tuyến                │ Gõ lệnh → Chọn các tuyến                                  │
│ 9   │ CTSV_XuatKhoiLuong                 │ Xuất khối lượng Material ra Excel                  │ Gõ lệnh → Chọn các tuyến → Sắp xếp vật liệu → Xuất file   │
│ 10  │ CTSV_ThongKeMaterialTracNgang      │ Thống kê Material trên trắc ngang                  │ Gõ lệnh → Chọn tuyến                                      │
│ 11  │ CTSV_PolyArea                      │ Menu tính diện tích Polyline                       │ Gõ lệnh → Chọn chức năng từ menu                          │
│ 12  │ CTSV_TinhDienTichPolyExcel         │ Tính diện tích Polyline xuất Excel                 │ Gõ lệnh → Chọn các Polyline                               │
│ 13  │ CTSV_TinhDienTichPoly              │ Tính diện tích Polyline                            │ Gõ lệnh → Chọn các Polyline                               │
│ 14  │ CTSV_TinhKhoiLuongPoly             │ Tính khối lượng từ Polyline                        │ Gõ lệnh → Chọn các Polyline → Nhập khoảng cách            │
│ 15  │ CTSV_GhiDienTichPoly               │ Ghi chú diện tích lên Polyline                     │ Gõ lệnh → Chọn các Polyline                               │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ NHÓM 11: THỐNG KÊ CỌC (11.ThongkeCoc.cs, 13.ThongKeCoc.cs)                                                                                               │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ STT │ Tên lệnh                          │ Mô tả                                              │ Cách dùng                                                 │
├─────┼────────────────────────────────────┼────────────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
│ 1   │ CTSV_ThongKeCoc                    │ Thống kê cọc ra Excel                              │ Gõ lệnh → Chọn tuyến → Chọn nhóm cọc                      │
│ 2   │ CTSV_ThongKeCoc_ToaDo              │ Thống kê cọc kèm tọa độ ra Excel                   │ Gõ lệnh → Chọn tuyến → Chọn nhóm cọc                      │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ NHÓM 12: SAN NỀN (14.SanNen.cs)                                                                                                                          │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ STT │ Tên lệnh                          │ Mô tả                                              │ Cách dùng                                                 │
├─────┼────────────────────────────────────┼────────────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
│ 1   │ CTSN_TaoLuoi                       │ Tạo lưới ô vuông san nền                           │ Gõ lệnh → Chọn vùng → Nhập kích thước ô                   │
│ 2   │ CTSN_TinhKL                        │ Tính khối lượng đào đắp san nền                    │ Gõ lệnh → Chọn lưới                                       │
│ 3   │ CTSN_NhapCaoDo                     │ Nhập cao độ TN/TK cho các góc lưới                 │ Gõ lệnh → Chọn góc → Nhập cao độ                          │
│ 4   │ CTSN_Taskbar                       │ Mở thanh công cụ San Nền                           │ Gõ lệnh                                                   │
│ 5   │ CTSN_Surface                       │ Lấy cao độ từ Surface cho lưới                     │ Gõ lệnh → Chọn Surface → Chọn lưới                        │
│ 6   │ CTSN_XuatBang                      │ Xuất bảng khối lượng san nền ra CAD                │ Gõ lệnh → Chọn lưới → Chọn vị trí đặt bảng                │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ NHÓM 13: BỀ MẶT TRỪ BỀ MẶT (15.BeMatTruBeMat.cs)                                                                                                         │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ STT │ Tên lệnh                          │ Mô tả                                              │ Cách dùng                                                 │
├─────┼────────────────────────────────────┼────────────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
│ 1   │ CTSV_DaoDap                        │ Tính khối lượng đào đắp (Surface-Surface)          │ Gõ lệnh → Chọn các tuyến → Xuất Excel                     │
│ 2   │ CTSV_HienThiMaterialList           │ Hiển thị Material List trong SampleLineGroup       │ Gõ lệnh → Chọn tuyến                                      │
│ 3   │ CTSV_ChiTietMaterialSection        │ Hiển thị chi tiết Material Section                 │ Gõ lệnh → Chọn Material Section                           │
│ 4   │ CTSV_KhoiLuongTracNgang            │ Xuất khối lượng trắc ngang (Average End Area)      │ Gõ lệnh → Chọn tuyến → Xuất Excel                         │
│ 5   │ CTSV_SoSanhKhoiLuong               │ So sánh khối lượng với Civil 3D                    │ Gõ lệnh → Chọn tuyến                                      │
│ 6   │ CTSV_PhanTichArea                  │ Phân tích chi tiết nguồn gốc AREA                  │ Gõ lệnh → Chọn Material Section                           │
│ 7   │ CTSV_KiemTraKhoiLuong              │ Kiểm tra so sánh AREA với Properties Panel         │ Gõ lệnh → Chọn tuyến                                      │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ NHÓM 14: TASKBAR & CÔNG CỤ QUẢN LÝ (15.CivilToolTaskbar.cs, 18.Menu Risbbon.cs, 19.ExternalTools.cs)                                                     │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ STT │ Tên lệnh                          │ Mô tả                                              │ Cách dùng                                                 │
├─────┼────────────────────────────────────┼────────────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
│ 1   │ CT_Taskbar                         │ Mở thanh công cụ Civil Tool                        │ Gõ lệnh                                                   │
│ 2   │ TASKBAR                            │ Mở thanh công cụ (alias)                           │ Gõ lệnh                                                   │
│ 3   │ CT                                 │ Mở thanh công cụ (alias ngắn)                      │ Gõ lệnh                                                   │
│ 4   │ CT_DanhSachLenh ★                  │ Hiển thị TOÀN BỘ danh sách lệnh + Tìm kiếm         │ Gõ lệnh → Tìm kiếm → Double-click chạy                    │
│ 5   │ CT_VTOADOHG                        │ Chạy tool tọa độ hố ga (DLL ngoài)                 │ Gõ lệnh                                                   │
│ 6   │ show_menu                          │ Reload/Tạo lại Ribbon menu                         │ Gõ lệnh (sau khi thay đổi code + NETLOAD)                 │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

╔══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
║                                              BẢNG 2: LỆNH TẮT NHANH (ALIAS)                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝

┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ BẢNG SỬA TÊN LỆNH NHANH - Thêm vào file ACAD.PGP hoặc Custom.pgp                                                                                         │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ Lệnh tắt │ Lệnh đầy đủ                           │ Mô tả                                                                                                 │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│       TASKBAR / CÔNG CỤ CHÍNH                                                                                                                            │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ CT       │ CT_Taskbar                             │ Mở thanh công cụ chính Civil Tool                                                                     │
│ TB       │ CTSV_Taskbar                           │ Mở thanh công cụ khối lượng                                                                           │
│ SN       │ CTSN_Taskbar                           │ Mở thanh công cụ San nền                                                                              │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│       KHỐI LƯỢNG                                                                                                                                         │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ XKL      │ CTSV_XuatKhoiLuong                     │ Xuất khối lượng Material ra Excel                                                                     │
│ XC       │ CTSV_XuatCad                           │ Xuất bảng khối lượng ra CAD                                                                           │
│ DD       │ CTSV_DaoDap                            │ Tính khối lượng đào đắp                                                                               │
│ PA       │ CTSV_PolyArea                          │ Menu tính diện tích Polyline                                                                          │
│ CD       │ CTSV_CaiDatBang                        │ Cài đặt định dạng bảng                                                                                │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│       CỌC / SAMPLELINE                                                                                                                                   │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ DTC      │ CTS_DoiTenCoc                          │ Đổi tên cọc                                                                                           │
│ DTC2     │ CTS_DoiTenCoc2                         │ Đổi tên cọc (cách 2)                                                                                  │
│ DTC3     │ CTS_DoiTenCoc3                         │ Đổi tên cọc theo lý trình                                                                             │
│ PSC      │ CTS_PhatSinhCoc                        │ Phát sinh cọc                                                                                         │
│ DCT      │ CTS_DichCoc_TinhTien                   │ Dịch cọc tịnh tiến                                                                                    │
│ TDC      │ CTS_TaoBang_ToaDoCoc                   │ Tạo bảng tọa độ cọc                                                                                   │
│ DB2      │ CTS_DongBo_2_NhomCoc                   │ Đồng bộ 2 nhóm cọc                                                                                    │
│ CBR      │ CTS_Copy_BeRong_sampleLine             │ Sao chép bề rộng cọc                                                                                  │
│ TBR      │ CTS_Thaydoi_BeRong_sampleLine          │ Thay đổi bề rộng cọc                                                                                  │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│       TRẮC DỌC                                                                                                                                           │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ VTD      │ CTP_VeTracDoc_TuNhien                  │ Vẽ trắc dọc tự nhiên                                                                                  │
│ VTA      │ CTP_VeTracDoc_TuNhien_TatCaTuyen       │ Vẽ trắc dọc tất cả tuyến                                                                              │
│ FIX      │ CTP_Fix_DuongTuNhien_TheoCoc           │ Sửa đường TN theo cọc                                                                                 │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│       TRẮC NGANG                                                                                                                                         │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ VTN      │ CTSV_VeTracNgang_ThietKe               │ Vẽ trắc ngang thiết kế                                                                                │
│ VNA      │ CTSV_VeTatCaTracNgang_ThietKe          │ Vẽ tất cả trắc ngang                                                                                  │
│ DC       │ CTSV_DanhCap                           │ Đánh cấp trắc ngang                                                                                   │
│ FK       │ CTSV_FitKhungIn                        │ Fit khung in                                                                                          │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│       POINT                                                                                                                                              │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ TCP      │ CTPO_TaoCogoPoint_CaoDo_FromSurface    │ Tạo Cogo Point từ Surface                                                                             │
│ CFT      │ CTPO_CreateCogopointFromText           │ Tạo Cogo Point từ Text                                                                                │
│ UPG      │ CTPO_UpdateAllPointGroup               │ Cập nhật Point Group                                                                                  │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│       PIPE                                                                                                                                               │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ TDK      │ CTPI_ThayDoi_DuongKinhCong             │ Thay đổi đường kính cống                                                                              │
│ TDD      │ CTPI_ThayDoi_DoanDocCong               │ Thay đổi độ dốc cống                                                                                  │
│ XHT      │ CTPI_XoayHoThu_Theo2diem               │ Xoay hố thu theo 2 điểm                                                                               │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│       SAN NỀN                                                                                                                                            │
├──────────┼────────────────────────────────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ TL       │ CTSN_TaoLuoi                           │ Tạo lưới san nền                                                                                      │
│ TKL      │ CTSN_TinhKL                            │ Tính khối lượng san nền                                                                               │
│ NCD      │ CTSN_NhapCaoDo                         │ Nhập cao độ cho lưới                                                                                  │
│ XB       │ CTSN_XuatBang                          │ Xuất bảng san nền                                                                                     │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

╔══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
║                                              BẢNG 3: NỘI DUNG FILE ACAD.PGP                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝

; ==== CIVIL TOOL ALIASES ====
; Copy nội dung bên dưới vào file ACAD.PGP hoặc tạo file Custom.pgp

; Taskbar
CT,         *CT_Taskbar
TB,         *CTSV_Taskbar
SN,         *CTSN_Taskbar

; Khối lượng
XKL,        *CTSV_XuatKhoiLuong
XC,         *CTSV_XuatCad
DD,         *CTSV_DaoDap
PA,         *CTSV_PolyArea
CD,         *CTSV_CaiDatBang

; Cọc / Sampleline
DTC,        *CTS_DoiTenCoc
DTC2,       *CTS_DoiTenCoc2
DTC3,       *CTS_DoiTenCoc3
PSC,        *CTS_PhatSinhCoc
DCT,        *CTS_DichCoc_TinhTien
TDC,        *CTS_TaoBang_ToaDoCoc
DB2,        *CTS_DongBo_2_NhomCoc
CBR,        *CTS_Copy_BeRong_sampleLine
TBR,        *CTS_Thaydoi_BeRong_sampleLine

; Trắc dọc
VTD,        *CTP_VeTracDoc_TuNhien
VTA,        *CTP_VeTracDoc_TuNhien_TatCaTuyen
FIX,        *CTP_Fix_DuongTuNhien_TheoCoc

; Trắc ngang
VTN,        *CTSV_VeTracNgang_ThietKe
VNA,        *CTSV_VeTatCaTracNgang_ThietKe
DC,         *CTSV_DanhCap
FK,         *CTSV_FitKhungIn

; Point
TCP,        *CTPO_TaoCogoPoint_CaoDo_FromSurface
CFT,        *CTPO_CreateCogopointFromText
UPG,        *CTPO_UpdateAllPointGroup

; Pipe
TDK,        *CTPI_ThayDoi_DuongKinhCong
TDD,        *CTPI_ThayDoi_DoanDocCong
XHT,        *CTPI_XoayHoThu_Theo2diem

; San nền
TL,         *CTSN_TaoLuoi
TKL,        *CTSN_TinhKL
NCD,        *CTSN_NhapCaoDo
XB,         *CTSN_XuatBang

; ==== END CIVIL TOOL ALIASES ====

╔══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
║                                              BẢNG 4: QUY TẮC ĐẶT TÊN LỆNH                                                                                ║
╚══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝

PREFIX LỆNH:
┌──────────┬──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ Prefix   │ Ý nghĩa                                                                                                                                      │
├──────────┼──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ CT       │ Civil Tool (lệnh chung)                                                                                                                      │
│ CTC      │ Civil Tool - Corridor                                                                                                                        │
│ CTPA     │ Civil Tool - Parcel                                                                                                                          │
│ CTPI     │ Civil Tool - Pipe                                                                                                                            │
│ CTPO     │ Civil Tool - Point                                                                                                                           │
│ CTP      │ Civil Tool - Profile                                                                                                                         │
│ CTS      │ Civil Tool - Sampleline                                                                                                                      │
│ CTSV     │ Civil Tool - Section View / Volume                                                                                                           │
│ CTSN     │ Civil Tool - San Nền                                                                                                                         │
│ AT       │ AutoCAD Tool (lệnh AutoCAD thuần)                                                                                                            │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

*/

/*
╔══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
║                                              BẢNG 5: HƯỚNG DẪN THÊM LỆNH MỚI VÀO TASKBAR & RIBBON                                                       ║
╚══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝

LỆNH QUẢN LÝ DANH SÁCH LỆNH:
============================
│ CT_DanhSachLenh  │ Mở form hiển thị toàn bộ danh sách lệnh, có thể tìm kiếm và chạy lệnh trực tiếp                                                       │
│ show_menu        │ Reload/tạo lại menu Ribbon với cấu trúc mới nhất                                                                                      │

QUY TRÌNH THÊM LỆNH MỚI:
========================

Bước 1: TẠO LỆNH MỚI
--------------------
- Tạo lệnh mới trong file .cs phù hợp (theo nhóm chức năng)
- Sử dụng prefix đúng quy tắc: CTC_, CTSV_, CTS_, CTPO_, CTPI_, CTSN_, CTP_

Bước 2: CẬP NHẬT FILE NÀY (00.LapDanhSachLenh.cs)
-------------------------------------------------
- Thêm lệnh vào BẢNG 1 (nhóm tương ứng)
- Thêm alias vào BẢNG 2 nếu cần
- Cập nhật nội dung ACAD.PGP trong BẢNG 3

Bước 3: THÊM VÀO TASKBAR (15.CivilToolTaskbar.cs)
-------------------------------------------------
Mở file và tìm hàm GetXxxCommands() tương ứng:

| Nhóm           | Hàm cần sửa                    |
|----------------|--------------------------------|
| Bề mặt         | GetSurfaceCommands()           |
| Cọc            | GetSampleLineCommands()        |
| Tuyến          | GetAlignmentCommands()         |
| Trắc dọc       | GetProfileCommands()           |
| Corridor       | GetCorridorCommands()          |
| Trắc ngang     | GetSectionViewCommands()       |
| Nút giao       | GetIntersectionCommands()      |
| San nền        | GetSanNenCommands()            |
| Khung in       | GetPlanCommands()              |
| Công cụ        | GetExternalToolsCommands()     |

Bước 4: THÊM VÀO RIBBON (18.Menu Risbbon.cs)
--------------------------------------------
Tìm array lệnh phù hợp và thêm vào:

| Nhóm           | Array cần sửa                  |
|----------------|--------------------------------|
| Bề mặt + Point | surfacesCommands, pointCommands|
| Cọc            | samplelineCommands             |
| Tuyến          | profileCommands, corridorCommands|
| Trắc ngang     | sectionviewCommands            |
| Công cụ        | utilitiesCommands, gradingCommands, pipeCommands|

CÚ PHÁP THÊM LỆNH:
------------------
// Taskbar (15.CivilToolTaskbar.cs)
return new List<(string, string)>
{
    ("─────────────", ""),                    // Đường kẻ phân cách
    ("🔧 Tên hiển thị", "TEN_LENH_AUTOCAD"),  // Thêm lệnh mới
};

// Ribbon (18.Menu Risbbon.cs)
(string Command, string Label)[] arrayCommands =
[
    ("TEN_LENH", "🔧 Mô tả lệnh"),
];

ĐIỀU CHỈNH ICON (EMOJI):
------------------------
📊 Thống kê/Báo cáo    📐 Đo lường         ✏ Chỉnh sửa         ➕ Thêm mới
📥 Xuất/Import         🔄 Cập nhật         ❌ Xóa              📋 Bảng/Danh sách
⚙ Cài đặt             🔧 Công cụ          📏 Kích thước       🎨 Style
🏷 Label               ↔ Di chuyển         👁 Hiển thị          🔒 Khóa
📍 Vị trí/Tọa độ      🛤 Corridor          🛣 Tuyến            📈 Trắc dọc
📉 Trắc ngang          ▦ San nền           🗺 Bề mặt           🔍 Tìm kiếm

ĐIỀU CHỈNH CẤU TRÚC MENU:
-------------------------
1. XÓA LỆNH: Xóa dòng tương ứng trong array
2. THÊM PHÂN CÁCH: ("─────────────", ""),
3. ĐỔI THỨ TỰ: Di chuyển dòng trong array
4. ĐỔI ICON: Thay emoji trước tên lệnh
5. ĐỔI TÊN HIỂN THỊ: Thay text trong quotes đầu tiên

VÍ DỤ: THÊM LỆNH "CTSV_TinhDienTichMoi"
---------------------------------------
1. Tạo lệnh trong file 10.TinhKhoiLuongExcel.cs:
   [CommandMethod("CTSV_TinhDienTichMoi")]
   public static void CTSVTinhDienTichMoi() { ... }

2. Thêm vào BẢNG 1 (Nhóm 10):
   | 16  | CTSV_TinhDienTichMoi  | Tính diện tích mới  | Gõ lệnh → Chọn đối tượng |

3. Thêm alias vào BẢNG 2:
   | TDM  | CTSV_TinhDienTichMoi  | Tính diện tích mới |

4. Thêm vào Taskbar - GetSectionViewCommands():
   ("📐 Tính diện tích mới", "CTSV_TinhDienTichMoi"),

5. Thêm vào Ribbon - sectionviewCommands:
   ("CTSV_TinhDienTichMoi", "📐 Tính diện tích mới"),

=================================================================================================
Lưu ý: Đây là file documentation, không chứa code thực thi
Để sử dụng các lệnh, cần build project và load DLL vào AutoCAD Civil 3D
Sau khi thay đổi code, chạy lệnh 'show_menu' trong AutoCAD để reload Ribbon
=================================================================================================
*/

