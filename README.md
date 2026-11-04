# Civil Tool – NXSOFT

Bộ công cụ tự động hoá cho **AutoCAD Civil 3D** (2025 trở lên): tuyến, trắc dọc, trắc ngang, khối lượng, cọc, xuất Excel/bảng CAD, tiện ích AutoCAD và module tính toán đường.

Tác giả: **Huỳnh Ngọc Nam** – Traffic Engineer – 0378584154 – Namhuynh1205@gmail.com

## Cài đặt

1. Tắt AutoCAD / Civil 3D.
2. Chạy `Civil_Tool_NXSOFT_Setup_vX.Y.Z.exe` (thư mục `Civil_Tool_Installer`).
3. Mở Civil 3D — plugin tự nạp, dòng lệnh hiện thông báo kèm số phiên bản.

| Lệnh | Tác dụng |
|---|---|
| `CT` | Mở Taskbar |
| `NX` | Hộp lệnh: tìm & chạy nhanh mọi lệnh |
| `show_menu` | Tạo lại Ribbon |
| `CT_DanhSachLenh` | Danh sách lệnh |
| `CT_LOG` | Mở thư mục log (gửi khi cần hỗ trợ) |
| `CT_THONGKELENH` | Lệnh bạn hay dùng |

## Dành cho người phát triển

| Việc | Cách làm |
|---|---|
| Build + tạo bộ cài | bấm `build.bat` (Release, phiên bản lấy từ `<Version>` trong `MyFirstProject/MyFirstProject.csproj`) |
| Build nhanh khi đang code | `MyFirstProject/Build_UIPro.bat` (Debug) + lệnh `NRL` trong AutoCAD |
| Kiểm tra + commit tự động | ghi thông điệp vào `_ci_message.txt` rồi chạy `ci.bat` → kết quả ở `_ci_ketqua.txt` |
| Phát hành bản mới | sửa `<Version>`, ghi `CHANGELOG.md`, chạy `build.bat`, `git tag vX.Y.Z`, `git push --tags` |

Quy ước code:

- Lệnh mới nên chạy qua `SmartCommand.Execute(...)` để có thanh tiến trình + báo lỗi chuẩn.
- Không dùng `catch { }` rỗng — dùng `NxLog.Swallow(ex)` hoặc `NxLog.Error(ex, "ngữ cảnh")`.
- Thông tin tác giả / phiên bản lấy từ `MyFirstProject.Core.AppInfo`, không gõ cứng.
- Thêm lệnh mới → thêm 1 dòng mô tả vào `MyFirstProject/UI Pro/NxCommandCatalog.g.cs`.

## Danh mục lệnh (229 lệnh)

### AutoCAD Tool

| Lệnh | Mô tả |
|---|---|
| `AT_ANNOTIVE_SCALE_CURRENTONLY` | Annot. Scale |
| `AT_BLOCKLAYOUT` | Block Layout |
| `AT_CHINHBEDAYPLINE` | Chỉnh bề dày Polyline |
| `AT_DANHSOTHUTU` | Đánh số TT |
| `AT_DIMLAYOUT` | Dim Layout |
| `AT_DIMLAYOUT2` | Dim Layout 2 |
| `AT_DOIMAUBYLAYER` | Đổi màu đối tượng về ByLayer |
| `AT_GOPTEXT` | Gộp Text khảo sát |
| `AT_LABEL_FROMTEXT` | Label Text |
| `AT_LINEW` | Thiết lập lineweight cho đối tượng |
| `AT_MT2ML` | MultiText → MultiLine |
| `AT_OFFSET_2BEN` | Offset 2 bên |
| `AT_REBARBIM` | Xuất dữ liệu cốt thép sang mô hình BIM 3D |
| `AT_SOLID_SET_PROPERTYSET` | Gán property set cho solid |
| `AT_SOLID_SHOW_INFO` | Hiển thị thông tin solid |
| `AT_TAOMOI_TEXTLAYOUT` | Tạo mới TL |
| `AT_TBD` | Text By Description |
| `AT_TCD` | Tách chữ đặc biệt |
| `AT_TEXTLAYOUT` | Text Layout |
| `AT_TEXTLINK` | Text Link |
| `AT_TONGDIENTICH_FULL` | Tổng DT (Full) |
| `AT_TONGDIENTICH_REPLACE` | Tổng DT (Replace) |
| `AT_TONGDIENTICH_REPLACE2` | Tổng DT (Replace2) |
| `AT_TONGDIENTICH_REPLACE_CONGTHEM` | Tổng DT (Cộng) |
| `AT_TONGDODAI_FULL` | Tổng dài (Full) |
| `AT_TONGDODAI_REPLACE` | Tổng dài (Replace) |
| `AT_TONGDODAI_REPLACE2` | Tổng dài (Replace2) |
| `AT_TONGDODAI_REPLACE_CONGTHEM` | Tổng dài (Cộng) |
| `AT_TONGHOP` | Đo lường tổng hợp |
| `AT_UPDATE2TABLE` | Cập nhật dữ liệu từ bảng |
| `AT_UPDATELAYOUT` | Update Layout |
| `AT_XOADOITUONG_3DSOLID_BODY` | Xóa 3DSolid |
| `AT_XOADOITUONG_CUNGLAYER` | Xóa cùng Layer |
| `AT_XOAYDOITUONG_THEO2DIEM` | Xoay 2 điểm |
| `AT_XOAYDOITUONG_THEOVIEWPORT` | Xoay VP |
| `AT_XUATCAD` | Xuất bản vẽ CAD theo khung |
| `AT_XUATCADQUICK` | Xuất bản vẽ CAD nhanh |
| `EXPORTCAD` | Xuất bản vẽ CAD |
| `XUATBANG_TOADOPOLYLINE` | Xuất bảng tọa độ Polyline |
| `XUATCAD` | Xuất bản vẽ CAD |
| `XUATCADQ` | Xuất bản vẽ CAD nhanh |

### Tiện ích Civil

| Lệnh | Mô tả |
|---|---|
| `A_TAG` | Gắn nhãn (tag) cho đối tượng |
| `CTDEM` | Đếm text trong vùng chọn |
| `CTSU_TAOSPOTELEVATION_ONSURFACE_TAITIM` | Spot Elevation at Center |
| `CTU_ADDPOLYLINEVERTICES` | ⊕ Thêm đỉnh Polyline |
| `CTU_DRAWTALUY` | ╱ Vẽ Taluy |
| `CTU_DUMPOBJECT` | Dump Object Properties |
| `CTU_EXPLODEAEC` | Explode AEC Objects |
| `CTU_EXPORTCAD2007` | Xuất CAD 2007 |
| `CTU_EXPORTTEXTCOORDS` | Xuất tọa độ Text |
| `CTU_FINDINTERSECTIONS` | ◎ Tìm điểm giao |
| `CTU_MAKEPOINTFROMTEXT` | ⊕ Tạo Point từ Text |
| `CTU_STYLEAUTOOFF` | Tắt Auto-Style |
| `CTU_STYLEAUTOON` | Bật Auto-Style |
| `CTU_TEXTTOMTEXT` | ▭ Text → MText |
| `CTU_TOTALLENGTH` | ━ Tổng chiều dài |
| `CVT2PL` | Chuyển đối tượng sang polyline |
| `DICH` | Dịch text trong bản vẽ |

### Outlook & Email

| Lệnh | Mô tả |
|---|---|
| `CHEN_LINK_MAIL_VAO_CAD` | Chèn link email Outlook vào bản vẽ |
| `DOC_OUTLOOK` | Đọc email Outlook |
| `DOC_PST` | Đọc file PST Outlook |
| `IMPORT_EMAIL_TABLE_TO_CAD` | Nhập bảng email vào CAD |
| `LINK_MAIL_CAD` | Liên kết email với bản vẽ CAD |
| `MO_EXCEL_MAIL` | Mở file Excel danh sách email |
| `OPEN_EXCEL_MAIL` | Mở file Excel danh sách email |
| `READ_OUTLOOK_FILE` | Đọc file Outlook |
| `READ_PST_FILE` | Đọc file PST |
| `UPDATE_EXCEL_MAIL` | Cập nhật Excel danh sách email |
| `XUATLINK_OUTLOOK` | Xuất link email Outlook |
| `XUAT_BANG_EMAIL_CAD` | Xuất bảng email vào CAD |

### Hệ thống & Menu

| Lệnh | Mô tả |
|---|---|
| `CIVIL_RIBBON` | Reload menu |
| `CIVIL_TINHTOAN_RIBBON` | Hiển thị Ribbon module Tính toán |
| `CT` | Mở nhanh Ribbon Taskbar (WinForm) |
| `CT_DANHSACHLENH` | Mở form hiển thị toàn bộ danh sách lệnh, có thể tìm kiếm và chạy lệnh trực tiếp |
| `CT_DANHSACHLENH_CONSOLE` | In danh sách lệnh ra dòng lệnh |
| `CT_DOIICON` | Đổi Icon lệnh |
| `CT_LAMNHE` | Làm nhẹ bản vẽ Civil 3D |
| `CT_LAMNHECIVIL` | Làm nhẹ bản vẽ Civil 3D |
| `CT_LOG` | Mở thư mục log (gửi khi cần hỗ trợ) |
| `CT_MENU_SETTINGS` | Cài đặt menu |
| `CT_PALETTE` | Mở Civil Tool Palette (dockable) |
| `CT_PRINTCENTER` | ⭐ Print Center |
| `CT_SHOW` | Hiện Civil Tool Palette |
| `CT_STYLEINFO` | Style Info |
| `CT_STYLESETTINGS` | Style Manager |
| `CT_TASKBAR` | Mở thanh công cụ chính |
| `CT_TASKBAR_SETTINGS` | Cài đặt thanh công cụ Taskbar |
| `CT_THONGKELENH` | Thống kê các lệnh bạn hay dùng |
| `CT_VTOADOHG` | 📍 Tọa độ hố ga (External) |
| `NRL` | Nạp lại DLL (NetReload) |
| `RELOAD` | Nạp lại DLL (NetReload) |
| `TASKBAR` | Mở Ribbon Taskbar (WinForm) |

### Tuyến (Alignment)

| Lệnh | Mô tả |
|---|---|
| `CTA_DOITENTUYENHANGLOAT` | Đổi tên tuyến (alignment) hàng loạt |
| `CTA_TAOTUYENHANGLOAT` | Tạo tuyến hàng loạt từ polyline |
| `CTA_TAOTUYENNHANH` | Tạo tuyến nhanh từ polyline |

### Corridor

| Lệnh | Mô tả |
|---|---|
| `CTC_ADDALLSECTION` | Thêm tất cả section từ SampleLine vào Corridor |
| `CTC_KIEMTRADUONGCONG_13592` | Kiểm tra theo 13592 |
| `CTC_KIEMTRADUONGCONG_4054` | Kiểm tra theo 4054 |
| `CTC_TAOCOORIDOR_DUONGDOTHI_REPHAI` | Tạo corridor rẽ phải cho đường đô thị |
| `CTC_TAOMVIEWTHEOTUYEN` | Tạo MView theo tuyến |
| `CTC_TAOPOLYLINEVUONGGOCTRACDOC` | Tạo polyline vuông góc trắc dọc |
| `CTC_THIETLAPDUONGCONG` | ⊙ Mở Form Đường Cong |
| `CTC_THONGSODUONGCONG_13592` | ▤ Bảng TCVN 13592 |
| `CTC_THONGSODUONGCONG_4054` | ▤ Bảng TCVN 4054 |
| `CTC_TRACUUDUONGCONG` | ◎ Tra cứu thông số |
| `CVC_CREATECURBRETURN_CORRIDORREGION` | Tạo vùng corridor cho bó vỉa nút giao |

### Thiết lập bản vẽ

| Lệnh | Mô tả |
|---|---|
| `CTDS_CONVERTCM2M` | ◈ cm → m |
| `CTDS_CONVERTMM2M` | ◈ mm → m |
| `CTDS_EXPORTPDF` | Xuất PDF |
| `CTDS_PRINTALLLAYOUTS` | ▤ In tất cả |
| `CTDS_PRINTCURRENTLAYOUT` | ▤ In hiện tại |
| `CTDS_SAVECLEAN` | ◎ Lưu sạch |
| `CTDS_THIETLAP` | ⊙ Thiết lập bản vẽ |
| `CTDS_THIETLAPNANGCAO` | Thiết lập bản vẽ nâng cao |
| `CTDS_TURBOMODE` | Bật chế độ Turbo (tăng tốc bản vẽ) |

### Layer

| Lệnh | Mô tả |
|---|---|
| `CTL_HUB` | Layer Hub |
| `CTL_OFFALIGNMENT` | Tắt layer Alignment |
| `CTL_OFFCORRIDOR` | Tắt layer Corridor |
| `CTL_OFFDEFPOINTS` | Tắt layer Defpoints |
| `CTL_OFFHATCHDAODAP` | Tắt hatch đào đắp |
| `CTL_OFFSAMPLELINE` | Tắt layer SampleLine |
| `CTL_ONALIGNMENT` | Bật layer Alignment |
| `CTL_ONCORRIDOR` | Bật layer Corridor |
| `CTL_ONDEFPOINTS` | Bật layer Defpoints |
| `CTL_ONHATCHDAODAP` | Bật hatch đào đắp |
| `CTL_ONSAMPLELINE` | Bật layer SampleLine |
| `CTL_TOBAOBT` | Tô layer bao bê tông |
| `CTL_TOBAOCOTTHEP` | Tô layer bao cốt thép |
| `CTL_TODEFPOINTS` | D Sang Defpoints |
| `CTL_TODIM` | ◎ Sang Dim |
| `CTL_TOHATCH` | ▤ Sang Hatch |
| `CTL_TOKHUAT` | Tô layer nét khuất |
| `CTL_TORANHGIOI` | Tô layer ranh giới |
| `CTL_TOTEXT` | T Sang Text |
| `CTL_TOTHEP` | Tô layer thép |
| `CTL_TOTRUC` | Tô layer trục |

### Parcel

| Lệnh | Mô tả |
|---|---|
| `CTL_OFFPARCEL` | Tắt layer Parcel |
| `CTL_ONPARCEL` | Bật layer Parcel |
| `CTPA_TAOPARCEL_CACLOAINHA` | Tạo Parcel từ các polyline nhà |
| `CTP_ADDPARCELLABELS` | Thêm nhãn Parcel |

### Cống & Hố thu

| Lệnh | Mô tả |
|---|---|
| `CTPI_BANGCAODO_TUNHIENHOTHU` | Tạo bảng cao độ tự nhiên hố thu |
| `CTPI_THAYDOI_DOANDOCCONG` | Thiết lập độ dốc đoạn cống |
| `CTPI_THAYDOI_DUONGKINHCONG` | Thay đổi đường kính ống cống |
| `CTPI_THAYDOI_MATPHANGREF_CONG` | Thay đổi mặt phẳng Reference cho ống |
| `CTPI_XOAYHOTHU_THEO2DIEM` | Xoay hố thu theo 2 điểm |

### CogoPoint

| Lệnh | Mô tả |
|---|---|
| `CTPO_AN_COGOPOINT` | Ẩn các Cogo Point đã chọn |
| `CTPO_CPTT` | Point→Text CĐ |
| `CTPO_CREATECOGOPOINTFROMTEXT` | Tạo Cogo Point từ Text có cao độ |
| `CTPO_DOITENPOINT` | Đổi tên CogoPoint hàng loạt |
| `CTPO_NHAPTOADOTRUCTIEP` | Direct Point Entry |
| `CTPO_REORDERPOINTS` | Reorder Point Numbers |
| `CTPO_TAOCOGOPOINT_CAODO_ELEVATIONSPOT` | Chuyển Elevation Spot thành Cogo Point |
| `CTPO_TAOCOGOPOINT_CAODO_FROMSURFACE` | Tạo Cogo Point lấy cao độ từ Surface |
| `CTPO_TAOPOINTTHEOBANG` | Import Points from Excel |
| `CTPO_UPDATEALLPOINTGROUP` | Cập nhật tất cả Point Group |

### Trắc dọc (Profile)

| Lệnh | Mô tả |
|---|---|
| `CTP_FIX_DUONGTUNHIEN_THEOCOC` | Sửa đường TN theo vị trí cọc |
| `CTP_GANNHANNUTGIAO_LENTRACDOC` | Gắn nhãn nút giao lên trắc dọc |
| `CTP_TAOCOGOPOINTTUPVI` | Tạo Cogo Point từ điểm PVI trên trắc dọc |
| `CTP_VETRACDOC_TUNHIEN` | Vẽ trắc dọc tự nhiên từ Surface |
| `CTP_VETRACDOC_TUNHIEN_TATCATUYEN` | Vẽ trắc dọc TN cho tất cả tuyến |
| `CTP_XUATBANGCIVIL` | Xuất dữ liệu Trắc Dọc |

### Trắc ngang (Section View)

| Lệnh | Mô tả |
|---|---|
| `CTSV_AN_DUONGDIACHAT` | Ẩn đường địa chất |
| `CTSV_CAIDATBANG` | Cài đặt bảng thống kê |
| `CTSV_CHUYENDOI_TNTK_TNTN` | Chuyển đổi TNTK sang TNTN |
| `CTSV_DANHCAP` | Tính toán và vẽ đánh cấp |
| `CTSV_DANHCAP_CAPNHAT` | Cập nhật khối lượng đánh cấp |
| `CTSV_DANHCAP_VETHEM` | Vẽ thêm đánh cấp |
| `CTSV_DANHCAP_XOABO` | Xóa bỏ đánh cấp đã vẽ |
| `CTSV_DAODAP` | Tính khối lượng đào đắp (Surface-Surface) |
| `CTSV_DOCBANGMATERIALTABLE` | Đọc bảng Material Table khối lượng |
| `CTSV_DOCMATERIALAPI` | Đọc khối lượng vật liệu qua API |
| `CTSV_FIT_KHUNGIN` | Fit khung in theo section |
| `CTSV_GHIDIENTICHPOLY` | Ghi chú diện tích lên Polyline |
| `CTSV_HIEUCHINH_SECTION` | Hiệu chỉnh section (static mode) |
| `CTSV_HIEUCHINH_SECTION_DYNAMIC` | Hiệu chỉnh section (dynamic mode) |
| `CTSV_KHOACATNGANG_ADDPOINT` | Khóa cắt ngang và thêm điểm |
| `CTSV_LAPCTTINHTRACNGANG` | Lập công thức tính khối lượng trắc ngang |
| `CTSV_LAYDIENTICHTUSECTIONVIEW` | Lấy diện tích từ Section View |
| `CTSV_POLYAREA` | Menu tính diện tích Polyline |
| `CTSV_SOSANHSURFACE` | So sánh 2 Surface |
| `CTSV_TAOTRANGINTD` | Tạo trang in Trắc Dọc |
| `CTSV_TASKBAR` | Mở thanh công cụ khối lượng |
| `CTSV_THAYDOI_GIOIHAN_TRAIPHAI` | Thay đổi giới hạn trái/phải |
| `CTSV_THAYDOI_KHUNGIN` | Thay đổi khung in |
| `CTSV_THAYDOI_MSS_MIN_MAX` | Thay đổi mức cao tối thiểu/tối đa |
| `CTSV_THEMVATLIEU_TRENCATNGANG` | Thêm text vật liệu |
| `CTSV_THONGKECOC` | Thống kê cọc và xuất Excel |
| `CTSV_THONGKECOC_TOADO` | Thống kê cọc kèm tọa độ ra Excel |
| `CTSV_TINHDIENTICHMOI` | 📐 Tính diện tích mới |
| `CTSV_TINHDIENTICHPOLY` | Tính diện tích Polyline |
| `CTSV_TINHDIENTICHPOLYEXCEL` | Tính diện tích Polyline xuất Excel |
| `CTSV_TINHKHOILUONGPOLY` | Tính khối lượng từ Polyline |
| `CTSV_TINHKHOILUONGTRACNGANG` | ⭐ Tính KL Trắc Ngang |
| `CTSV_VEDUONGBAOMATERIAL` | Vẽ đường bao vật liệu trên trắc ngang |
| `CTSV_VETRACNGANGTHIETKE` | Vẽ trắc ngang thiết kế |
| `CTSV_XUATCAD` | Xuất bảng khối lượng ra CAD |
| `CVSV_VETATCA_TRACNGANGTHIETKE` | Vẽ tất cả trắc ngang thiết kế |

### Cọc (Sample Line)

| Lệnh | Mô tả |
|---|---|
| `CTS_CHENCOC_TRENTRACDOC` | Chèn cọc trên trắc dọc |
| `CTS_CHENCOC_TRENTRACNGANG` | Chèn cọc trên trắc ngang |
| `CTS_COPY_BERONG_SAMPLELINE` | Sao chép bề rộng sample line |
| `CTS_COPY_NHOMCOC` | Sao chép nhóm cọc |
| `CTS_DICHCOC_TINHTIEN` | Dịch cọc tịnh tiến |
| `CTS_DOITENCOC` | Đổi tên cọc |
| `CTS_DOITENCOC2` | Đổi tên cọc theo đoạn với tiền tố |
| `CTS_DOITENCOC3` | Đổi tên cọc thành dạng Km x+xxx |
| `CTS_DOITENCOC_FROMCOGOPOINT` | Đổi tên sample line theo tên CogoPoint |
| `CTS_DOITENCOC_THEOTHUTU` | Đổi tên cọc theo thứ tự |
| `CTS_DONGBO_2_NHOMCOC` | Đồng bộ 2 nhóm cọc |
| `CTS_DONGBO_2_NHOMCOC_THEODOAN` | Đồng bộ 2 nhóm cọc theo đoạn |
| `CTS_OFFSET_BERONG_SAMPLELINE` | Offset bề rộng sample line |
| `CTS_PHANTICHEXCEL` | Phân tích file TT |
| `CTS_PHATSINHCOC` | Phát sinh cọc tự động |
| `CTS_REBUILDSURFACE` | Rebuild Surface |
| `CTS_THAYDOI_BERONG_SAMPLELINE` | Thay đổi bề rộng sample line |

### Tiện ích BR

| Lệnh | Mô tả |
|---|---|
| `DEMCS` | Đếm số lượng đối tượng/cấu kiện |
| `LTDTB` | Lấy tọa độ và tên theo block/text |
| `TAONHAN` | Tạo nhãn Block Attribute |
| `XTB` | ⭐ Xuất tọa độ Block → Excel |
| `XTBX` | Xuất thống kê (mở rộng) |

### Tính toán thiết kế

| Lệnh | Mô tả |
|---|---|
| `TT_AODUONG` | Kết cấu áo đường |
| `TT_LANREPHAI` | Làn rẽ phải |
| `TT_LANRETRAI` | Làn rẽ trái |
| `TT_MORONGXECHAY` | Mở rộng xe chạy |
| `TT_NHAPDONG` | Chiều dài nhập dòng |
| `TT_TAMNHINNUTGIAO` | Tầm nhìn nút giao |
| `TT_TANGGIAMTOC` | Làn tăng/giảm tốc |
| `TT_THUYLUCCONGTRAN` | Thủy lực cống tràn |
| `TT_TINHTOANBANKINHRE` | Bán kính rẽ phải |
| `TT_VONGXUYEN` | Đảo vòng xuyến |
