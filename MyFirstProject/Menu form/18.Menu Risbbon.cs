using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Windows;

[assembly: CommandClass(typeof(MyFirstProject.Autocad))]

namespace MyFirstProject
{
    public class Autocad
    {
        [CommandMethod("ShowForm")]
        public static void ShowForm()
        {
            TestForm frmTest = new();
            frmTest.Show();
        }

        [CommandMethod("AdskGreeting")]
        public static void AdskGreeting()
        {
            Document? acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ?? throw new InvalidOperationException("No active document found.");
            Database? acCurDb = acDoc.Database ?? throw new InvalidOperationException("No database found for the active document.");

            using Transaction acTrans = acCurDb.TransactionManager.StartTransaction();
            BlockTable acBlkTbl = (BlockTable)acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) ?? throw new InvalidOperationException("BlockTable could not be retrieved.");
            BlockTableRecord acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) ?? throw new InvalidOperationException("BlockTableRecord could not be retrieved.");

            using (MText objText = new())
            {
                objText.Location = new Autodesk.AutoCAD.Geometry.Point3d(2, 2, 0);
                objText.Contents = "Greetings, Welcome to AutoCAD .NET";
                objText.TextStyleId = acCurDb.Textstyle;
                acBlkTblRec.AppendEntity(objText);
                acTrans.AddNewlyCreatedDBObject(objText, true);
            }
            acTrans.Commit();
        }

        [CommandMethod("show_menu")]
        public static void ShowMenu()
        {
            try
            {
                var ribbon = ComponentManager.Ribbon;
                if (ribbon == null)
                {
                    var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    doc?.SendStringToExecute("RIBBON ", true, false, false);
                    ribbon = ComponentManager.Ribbon;
                    if (ribbon == null)
                    {
                        doc?.Editor.WriteMessage("\nKhông thể khởi tạo Ribbon. Hãy bật RIBBON rồi chạy lại lệnh.");
                        return;
                    }
                }

                // Remove previous tab if exists
                var existing = ribbon.Tabs.FirstOrDefault(t => t.Id == "MyFirstProject.C3DTab");
                if (existing != null)
                {
                    ribbon.Tabs.Remove(existing);
                }
                var existingAcad = ribbon.Tabs.FirstOrDefault(t => t.Id == "MyFirstProject.AcadTab");
                if (existingAcad != null)
                {
                    ribbon.Tabs.Remove(existingAcad);
                }

                // Create new Civil tool tab
                RibbonTab tab = new()
                {
                    Title = "Civil tool",
                    Id = "MyFirstProject.C3DTab"
                };
                ribbon.Tabs.Add(tab);

                // Create new Acad tool tab
                RibbonTab acadTab = new()
                {
                    Title = "Acad tool",
                    Id = "MyFirstProject.AcadTab"
                };
                ribbon.Tabs.Add(acadTab);

                // Helper to add a dropdown menu for Civil tool commands
                void AddCivilDropdownPanel(RibbonTab targetTab, string panelTitle, (string Command, string Label)[] commands)
                {
                    if (commands.Length == 0) return; // Skip if no commands

                    RibbonPanelSource src = new() { Title = panelTitle };
                    RibbonPanel panel = new() { Source = src };
                    RibbonSplitButton splitButton = new()
                    {
                        Text = panelTitle,
                        ShowText = true,
                        ShowImage = false,
                        Size = RibbonItemSize.Large
                    };
                    foreach (var (command, label) in commands)
                    {
                        RibbonButton btn = new()
                        {
                            Text = label,
                            ShowText = true,
                            ShowImage = false,
                            Orientation = System.Windows.Controls.Orientation.Vertical,
                            Size = RibbonItemSize.Standard,
                            CommandHandler = new SimpleRibbonCommandHandler(),
                            Tag = command,
                            CommandParameter = command
                        };
                        splitButton.Items.Add(btn);
                    }
                    src.Items.Add(splitButton);
                    targetTab.Panels.Add(panel);
                }

                // Helper to add a panel with a large button
                void AddPanel(RibbonTab targetTab, string title)
                {
                    RibbonPanelSource src = new() { Title = title };
                    RibbonPanel panel = new() { Source = src };
                    RibbonButton btn = new()
                    {
                        Text = title,
                        ShowText = true,
                        ShowImage = false,
                        Orientation = System.Windows.Controls.Orientation.Vertical,
                        Size = RibbonItemSize.Large,
                        CommandHandler = new SimpleRibbonCommandHandler(),
                        Tag = title
                    };
                    src.Items.Add(btn);
                    targetTab.Panels.Add(panel);
                }

                // Helper to add a dropdown menu for Acad tool commands
                void AddAcadDropdownPanel(RibbonTab targetTab, string panelTitle, (string Command, string Label)[] commands)
                {
                    RibbonPanelSource src = new() { Title = panelTitle };
                    RibbonPanel panel = new() { Source = src };
                    RibbonSplitButton splitButton = new()
                    {
                        Text = panelTitle,
                        ShowText = true,
                        ShowImage = false,
                        Size = RibbonItemSize.Large
                    };
                    foreach (var (command, label) in commands)
                    {
                        RibbonButton btn = new()
                        {
                            Text = label,
                            ShowText = true,
                            ShowImage = false,
                            Orientation = System.Windows.Controls.Orientation.Vertical,
                            Size = RibbonItemSize.Standard,
                            CommandHandler = new SimpleRibbonCommandHandler(),
                            Tag = command,
                            CommandParameter = command
                        };
                        splitButton.Items.Add(btn);
                    }
                    src.Items.Add(splitButton);
                    targetTab.Panels.Add(panel);
                }

                // 1. Module Surfaces (09.Surfaces.cs)
                (string Command, string Label)[] surfacesCommands =
                [
                    ("CTS_TaoSpotElevation_OnSurface_TaiTim", "📏 Spot Elevation Tại Tim")
                ];
 
                // 2. Module SampleLine (07.Sampleline.cs) - 27 lệnh
                (string Command, string Label)[] samplelineCommands =
                [
                    ("CTS_DoiTenCoc", "✏ Đổi tên cọc"),
                    ("CTS_DoiTenCoc2", "✏ Đổi tên cọc đoạn"),
                    ("CTS_DoiTenCoc3", "✏ Đổi tên cọc Km"),
                    ("CTS_DoiTenCoc_fromCogoPoint", "✏ Đổi tên từ CogoPoint"),
                    ("CTS_DoiTenCoc_TheoThuTu", "✏ Đổi tên thứ tự"),
                    ("CTS_DoiTenCoc_H", "✏ Đổi tên hậu tố A"),
                    ("CTS_TaoBang_ToaDoCoc", "📐 Tọa độ cọc (X,Y)"),
                    ("CTS_TaoBang_ToaDoCoc2", "📐 Tọa độ cọc (Lý trình)"),
                    ("CTS_TaoBang_ToaDoCoc3", "📐 Tọa độ cọc (Cao độ)"),
                    ("AT_UPdate2Table", "🔄 Cập nhật từ bảng"),
                    ("CTS_ChenCoc_TrenTracDoc", "➕ Chèn trên trắc dọc"),
                    ("CTS_CHENCOC_TRENTRACNGANG", "➕ Chèn trên trắc ngang"),
                    ("CTS_PhatSinhCoc", "➕ Phát sinh cọc auto"),
                    ("CTS_PhatSinhCoc_ChiTiet", "➕ Phát sinh chi tiết"),
                    ("CTS_PhatSinhCoc_theoKhoangDelta", "➕ Phát sinh delta"),
                    ("CTS_PhatSinhCoc_TuCogoPoint", "➕ Phát sinh từ CogoPoint"),
                    ("CTS_PhatSinhCoc_TheoBang", "➕ Phát sinh từ bảng"),
                    ("CTS_DichCoc_TinhTien", "↔ Dịch cọc tịnh tiến"),
                    ("CTS_DichCoc_TinhTien40", "↔ Dịch cọc 40m"),
                    ("CTS_DichCoc_TinhTien_20", "↔ Dịch cọc 20m"),
                    ("CTS_Copy_NhomCoc", "📋 Sao chép nhóm cọc"),
                    ("CTS_DongBo_2_NhomCoc", "🔄 Đồng bộ nhóm cọc"),
                    ("CTS_DongBo_2_NhomCoc_TheoDoan", "🔄 Đồng bộ theo đoạn"),
                    ("CTS_Copy_BeRong_sampleLine", "📏 Copy bề rộng SL"),
                    ("CTS_Thaydoi_BeRong_sampleLine", "📏 Thay đổi bề rộng SL"),
                    ("CTS_Offset_BeRong_sampleLine", "📏 Offset bề rộng SL"),
                    ("CTSV_ThongKeCoc", "📊 Thống kê cọc (Excel)"),
                    ("CTSV_ThongKeCoc_TatCa", "📊 Thống kê tất cả cọc")
                ];

                // 3. Module Alignment & Profile
                (string Command, string Label)[] profileCommands =
                [
                    ("CTPV_TaoProfileView", "➕ Tạo trắc dọc"),
                    ("CTPV_SuaProfileView", "✏ Edit profile"),
                    ("CTPV_ThemBang_LyTrinh", "📋 Thêm bảng lý trình"),
                    ("CTPV_ThemLabel_CaoDo", "🏷 Thêm Label cao độ"),
                    ("CTPV_ThayDoiScale", "📏 Thay đổi Scale"),
                    ("CTPV_FitKhung", "📐 Fit khung")
                ];
 
                // 4. Module Corridor & Parcel
                (string Command, string Label)[] corridorCommands =
                [
                    ("CTC_AddAllSection", "➕ Thêm tất cả Section"),
                    ("CTC_TaoCooridor_DuongDoThi_RePhai", "🛤 Corridor rẽ phải"),
                    ("CTP_TaoBangThongKeParcel", "📦 Thống kê Parcel"),
                    ("CTP_TaoBangThongKeParcel_SapXep", "📦 Thống kê Parcel (Sắp xếp)")
                ];

                // 5. Module SectionView (08.Sectionview.cs) - 21 lệnh
                (string Command, string Label)[] sectionviewCommands =
                [
                    ("CTSV_VeTracNgangThietKe", "🎨 Tạo trắc ngang"),
                    ("CVSV_VeTatCa_TracNgangThietKe", "🎨 Vẽ tất cả TN"),
                    ("CTSV_ChuyenDoi_TNTK_TNTN", "🔄 Chuyển TK sang TN"),
                    ("CTSV_DanhCap", "📐 Đánh cấp - VHC"),
                    ("CTSV_DanhCap_XoaBo", "❌ Xóa đánh cấp"),
                    ("CTSV_DanhCap_VeThem", "➕ Vẽ thêm đánh cấp"),
                    ("CTSV_DanhCap_VeThem1", "➕ Vẽ thêm 1m"),
                    ("CTSV_DanhCap_VeThem2", "➕ Vẽ thêm 2m"),
                    ("CTSV_DanhCap_CapNhat", "🔄 Cập nhật KL đánh cấp"),
                    ("CTSV_ThemVatLieu_TrenCatNgang", "📋 Điền KL trắc ngang"),
                    ("CTSV_ThayDoi_MSS_Min_Max", "⚙ Hiệu chỉnh MSS"),
                    ("CTSV_ThayDoi_GioiHan_traiPhai", "↔ Thay giới hạn T/P"),
                    ("CTSV_ThayDoi_KhungIn", "📋 Dàn khung in"),
                    ("CTSV_KhoaCatNgang_AddPoint", "🔒 Khóa TN + Add Point"),
                    ("CTSV_fit_KhungIn", "📐 Fit khung in"),
                    ("CTSV_fit_KhungIn_5_5_top", "📐 Fit khung 5x5"),
                    ("CTSV_fit_KhungIn_5_10_top", "📐 Fit khung 5x10"),
                    ("CTSV_An_DuongDiaChat", "👁 Ẩn đường địa chất"),
                    ("CTSV_HieuChinh_Section", "✏ Hiệu chỉnh (Static)"),
                    ("CTSV_HieuChinh_Section_Dynamic", "✏ Hiệu chỉnh (Dynamic)"),
                    ("CTSV_ThongKeCoc", "📊 Thống kê cọc (Excel)"),
                    ("CTSV_ThongKeCoc_TatCa", "📊 Thống kê toàn bộ cọc"),
                    ("CTSV_ThongKeCoc_ToaDo", "📍 Thống kê tọa độ cọc"),
                    ("CTSV_Taskbar", "📊 Taskbar Khối Lượng"),
                    ("CTSV_XuatKhoiLuong", "📥 Xuất KL Excel"),
                    ("CTSV_XuatCad", "📥 Xuất KL CAD"),
                    ("CTSV_CaiDatBang", "⚙ Cài đặt bảng KL")
                ];

                // 6. Module San Nền (14.SanNen.cs)
                (string Command, string Label)[] gradingCommands =
                [
                    ("CTSN_Taskbar", "📊 Mở Taskbar SN"),
                    ("CTSN_TaoLuoi", "▦ Quản lý lưới"),
                    ("CTSN_NhapCaoDo", "📝 Điền cao độ lưới"),
                    ("CTSN_Surface", "🏔 Lấy CĐ Surface"),
                    ("CTSN_TinhKL", "📋 Tính khối lượng SN"),
                    ("CTSN_XuatBang", "📤 Xuất bảng KL CAD")
                ];

                // 7. Module Point (05.Point.cs)
                (string Command, string Label)[] pointCommands =
                [
                    ("CTPo_TaoPointTheoBang", "➕ Tạo Point từ bảng"),
                    ("CTPo_ChuyenPointThanhBlock", "🔄 Point → Block"),
                    ("CTPo_TaoBangThongKePoint", "📋 Bảng thống kê Point"),
                    ("CTPo_ThayDoiCaoDo", "✏ Thay đổi cao độ"),
                    ("CTPo_DatTen_theoThuTu", "🏷 Đặt tên thứ tự"),
                    ("CTPo_ThayDoiStyle", "🎨 Thay đổi Style"),
                    ("CTPo_LayThongTin", "ℹ Lấy thông tin")
                ];

                // 8. Module Pipe & Structures (04.PipeAndStructures.cs)
                (string Command, string Label)[] pipeCommands =
                [
                    ("CTPS_TaoBangThongKePipe", "🔧 Thống kê Pipe"),
                    ("CTPS_TaoBangThongKeStructure", "🔧 Thống kê Structure"),
                    ("CTPS_ThayDoi_CaoDo_Pipe", "📏 Đổi cao độ Pipe"),
                    ("CTPS_ThayDoi_CaoDo_Structure", "📏 Đổi cao độ Struct"),
                    ("CTPS_XoayPipe_90do", "🔄 Xoay Pipe 90°"),
                    ("CTPS_XoaConTrung", "❌ Xóa con trùng")
                ];

                // 9. Module Utilities & Property Sets
                (string Command, string Label)[] utilitiesCommands =
                [
                    ("AT_Solid_Set_PropertySet", "📦 Gán Property Set"),
                    ("AT_Solid_Show_Info", "ℹ Thông tin Solid"),
                    ("CT_VTOADOHG", "📍 Tọa độ hố ga"),
                    ("show_menu", "🔄 Reload Menu")
                ];

                // 10. Module Account
                (string Command, string Label)[] accountCommands =
                [
                    ("", "🔑 Đăng nhập"),
                    ("", "ℹ Thông tin"),
                    ("", "📖 Hướng dẫn"),
                ];

                // Các lệnh từ 01.CAD.cs cho Acad tool
                (string Command, string Label)[] acadCommands =
                [
                    ("AT_TongDoDai_Full", "Tổng Độ Dài (Full)"),
                    ("AT_TongDoDai_Replace", "Tổng Độ Dài (Replace)"),
                    ("AT_TongDoDai_Replace2", "Tổng Độ Dài (Replace2)"),
                    ("AT_TongDoDai_Replace_CongThem", "Tổng Độ Dài (Cộng Thêm)"),
                    ("AT_TongDienTich_Full", "Tổng Diện Tích (Full)"),
                    ("AT_TongDienTich_Replace", "Tổng Diện Tích (Replace)"),
                    ("AT_TongDienTich_Replace2", "Tổng Diện Tích (Replace2)"),
                    ("AT_TongDienTich_Replace_CongThem", "Tổng Diện Tích (Cộng Thêm)"),
                    ("AT_TextLink", "Text Link"),
                    ("AT_DanhSoThuTu", "Đánh Số Thứ Tự"),
                    ("AT_XoayDoiTuong_TheoViewport", "Xoay Đối Tượng Theo Viewport"),
                    ("AT_XoayDoiTuong_Theo2Diem", "Xoay Đối Tượng Theo 2 Điểm"),
                    ("AT_TextLayout", "Text Layout"),
                    ("AT_TaoMoi_TextLayout", "Tạo Mới Text Layout"),
                    ("AT_DimLayout2", "Dim Layout 2"),
                    ("AT_DimLayout", "Dim Layout"),
                    ("AT_BlockLayout", "Block Layout"),
                    ("AT_Label_FromText", "Label From Text"),
                    ("AT_XoaDoiTuong_CungLayer", "Xóa Đối Tượng Cùng Layer"),
                    ("AT_XoaDoiTuong_3DSolid_Body", "Xóa 3DSolid/Body"),
                    ("AT_UpdateLayout", "Update Layout"),
                    ("AT_Offset_2Ben", "Offset 2 Bên"),
                    ("AT_annotive_scale_currentOnly", "Annotative Scale Current Only")
                ];

                // Add panels to Civil tool tab in correct order
                AddCivilDropdownPanel(tab, "Bề mặt & Điểm", surfacesCommands);
                AddCivilDropdownPanel(tab, "Lưới cọc", samplelineCommands);
                AddCivilDropdownPanel(tab, "Trắc dọc & Tuyến", profileCommands);
                AddCivilDropdownPanel(tab, "Corridor & Thửa", corridorCommands);
                AddCivilDropdownPanel(tab, "Trắc ngang & KL", sectionviewCommands);
                AddCivilDropdownPanel(tab, "San nền", gradingCommands);
                AddCivilDropdownPanel(tab, "Thoát nước", pipeCommands);
                AddCivilDropdownPanel(tab, "Point", pointCommands);
                AddCivilDropdownPanel(tab, "Tiện ích", utilitiesCommands);
                AddCivilDropdownPanel(tab, "Tài khoản", accountCommands);

                // Thêm menu sổ xuống cho các lệnh Acad tool
                AddAcadDropdownPanel(acadTab, "CAD Commands", acadCommands);

                tab.IsActive = true;
                var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor;
                ed?.WriteMessage("\nĐã tạo tab 'Civil tool' và 'Acad tool' với đầy đủ các lệnh từ Civil Tool files trên Ribbon.");
            }
            catch (System.Exception ex)
            {
                var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor;
                ed?.WriteMessage($"\nLỗi tạo menu: {ex.Message}");
            }
        }

        private class SimpleRibbonCommandHandler : System.Windows.Input.ICommand
        {
            public bool CanExecute(object? parameter) => true;

            public event EventHandler? CanExecuteChanged { add { } remove { } }

            public void Execute(object? parameter)
            {
                try
                {
                    string? commandToRun = null;
                    
                    if (parameter is string cmd)
                    {
                        commandToRun = cmd;
                    }
                    else if (parameter is RibbonButton rb)
                    {
                        commandToRun = rb.CommandParameter as string;
                    }

                    if (string.IsNullOrWhiteSpace(commandToRun)) return;

                    var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    if (doc != null)
                    {
                        doc.SendStringToExecute(commandToRun + " ", true, false, true);
                    }
                }
                catch (System.Exception ex)
                {
                    var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor;
                    ed?.WriteMessage($"\nLỗi thực thi lệnh: {ex.Message}");
                }
            }
        }
    }
}