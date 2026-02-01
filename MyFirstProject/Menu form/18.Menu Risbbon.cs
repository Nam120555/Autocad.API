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
                            Tag = command
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
                            Tag = command
                        };
                        splitButton.Items.Add(btn);
                    }
                    src.Items.Add(splitButton);
                    targetTab.Panels.Add(panel);
                }

                // Lệnh từ Civil Tool files - 01.Corridor.cs
                (string Command, string Label)[] corridorCommands =
                [
                    ("CTC_AddAllSection", "Add All Section"),
                    ("CTC_TaoCooridor_DuongDoThi_RePhai", "Tạo Corridor Rẽ Phải")
                ];

                // Lệnh từ 02.Parcel.cs  
                (string Command, string Label)[] parcelCommands =
                [
                    ("CTPA_TaoParcel_CacLoaiNha", "Tạo Parcel Các Loại Nhà")
                ];

                // Lệnh từ 03.MenuFunction.cs
                (string Command, string Label)[] menuCommands =
                [
                    ("testmyRibbon", "Test My Ribbon")
                ];

                // Lệnh từ 04.PipeAndStructures.cs
                (string Command, string Label)[] pipeCommands =
                [
                    ("CTPI_ThayDoi_DuongKinhCong", "Thay Đổi Đường Kính Cống"),
                    ("CTPI_ThayDoi_MatPhangRef_Cong", "Thay Đổi Mặt Phẳng Ref Cống"),
                    ("CTPI_ThayDoi_DoanDocCong", "Thay Đổi Độ Dốc Cống"),
                    ("CTPI_BangCaoDo_TuNhienHoThu", "Bảng Cao Độ Hố Thu"),
                    ("CTPI_XoayHoThu_Theo2diem", "Xoay Hố Thu Theo 2 Điểm")
                ];

                // Lệnh từ 05.Point.cs
                (string Command, string Label)[] pointCommands =
                [
                    ("CTPO_TaoCogoPoint_CaoDo_FromSurface", "Tạo CogoPoint Từ Surface"),
                    ("CTPO_TaoCogoPoint_CaoDo_Elevationspot", "Tạo CogoPoint Từ Elevation Spot"),
                    ("CTPO_UpdateAllPointGroup", "Update All Point Group"),
                    ("CTPO_CreateCogopointFromText", "Tạo CogoPoint Từ Text"),
                    ("CTPO_An_CogoPoint", "Ẩn CogoPoint")
                ];

                // Lệnh từ 06.ProfileAndProfileView.cs
                (string Command, string Label)[] profileCommands =
                [
                    ("CTP_VeTracDoc_TuNhien", "Vẽ Trắc Dọc Tự Nhiên"),
                    ("CTP_VeTracDoc_TuNhien_TatCaTuyen", "Vẽ Trắc Dọc Tất Cả Tuyến"),
                    ("CTP_Fix_DuongTuNhien_TheoCoc", "Sửa Đường Tự Nhiên Theo Cọc"),
                    ("CTP_GanNhanNutGiao_LenTracDoc", "Gắn Nhãn Nút Giao Lên Trắc Dọc"),
                    ("CTP_TaoCogoPointTuPVI", "Tạo CogoPoint Từ PVI")
                ];

                // Lệnh từ 07.Sampleline.cs (đầy đủ 26 lệnh)
                (string Command, string Label)[] samplelineCommands =
                [
                    ("CTS_DoiTenCoc", "Đổi Tên Cọc"),
                    ("CTS_DoiTenCoc3", "Đổi Tên Cọc Km"),
                    ("CTS_DoiTenCoc2", "Đổi Tên Cọc Theo Đoạn"),
                    ("CTS_TaoBang_ToaDoCoc", "Tạo Bảng Tọa Độ Cọc"),
                    ("CTS_TaoBang_ToaDoCoc2", "Tạo Bảng Tọa Độ Cọc (Lý Trình)"),
                    ("CTS_TaoBang_ToaDoCoc3", "Tạo Bảng Tọa Độ Cọc (Cao Độ)"),
                    ("AT_UPdate2Table", "Cập Nhật 2 Table"),
                    ("CTS_ChenCoc_TrenTracDoc", "Chèn Cọc Trên Trắc Dọc"),
                    ("CTS_CHENCOC_TRENTRACNGANG", "Chèn Cọc Trên Trắc Ngang"),
                    ("CTS_DoiTenCoc_fromCogoPoint", "Đổi Tên Cọc Từ CogoPoint"),
                    ("CTS_PhatSinhCoc", "Phát Sinh Cọc"),
                    ("CTS_PhatSinhCoc_ChiTiet", "Phát Sinh Cọc Chi Tiết"),
                    ("CTS_PhatSinhCoc_theoKhoangDelta", "Phát Sinh Cọc Theo Delta"),
                    ("CTS_PhatSinhCoc_TuCogoPoint", "Phát Sinh Cọc Từ CogoPoint"),
                    ("CTS_DoiTenCoc_TheoThuTu", "Đổi Tên Cọc Theo Thứ Tự"),
                    ("CTS_DichCoc_TinhTien", "Dịch Cọc Tịnh Tiến"),
                    ("CTS_Copy_NhomCoc", "Copy Nhóm Cọc"),
                    ("CTS_DongBo_2_NhomCoc", "Đồng Bộ 2 Nhóm Cọc"),
                    ("CTS_DongBo_2_NhomCoc_TheoDoan", "Đồng Bộ 2 Nhóm Cọc Theo Đoạn"),
                    ("CTS_DichCoc_TinhTien40", "Dịch Cọc 40m"),
                    ("CTS_DichCoc_TinhTien_20", "Dịch Cọc 20m"),
                    ("CTS_DoiTenCoc_H", "Đổi Tên Cọc H"),
                    ("CTS_PhatSinhCoc_TheoBang", "Phát Sinh Cọc Theo Bảng"),
                    ("CTS_Copy_BeRong_sampleLine", "Copy Bề Rộng Sample Line"),
                    ("CTS_Thaydoi_BeRong_sampleLine", "Thay Đổi Bề Rộng Sample Line"),
                    ("CTS_Offset_BeRong_sampleLine", "Offset Bề Rộng Sample Line")
                ];

                // Lệnh từ 08.Sectionview.cs (đầy đủ 20 lệnh)
                (string Command, string Label)[] sectionviewCommands =
                [
                    ("CTSV_VeTracNgangThietKe", "Vẽ Trắc Ngang Thiết Kế"),
                    ("CVSV_VeTatCa_TracNgangThietKe", "Vẽ Tất Cả Trắc Ngang"),
                    ("CTSV_ChuyenDoi_TNTK_TNTN", "Chuyển Đổi TN-TK sang TN-TN"),
                    ("CTSV_DanhCap", "Tính Đánh Cấp"),
                    ("CTSV_DanhCap_XoaBo", "Xóa Bỏ Đánh Cấp"),
                    ("CTSV_DanhCap_VeThem", "Vẽ Thêm Đánh Cấp"),
                    ("CTSV_DanhCap_VeThem2", "Vẽ Thêm Đánh Cấp 2m"),
                    ("CTSV_DanhCap_VeThem1", "Vẽ Thêm Đánh Cấp 1m"),
                    ("CTSV_DanhCap_CapNhat", "Cập Nhật Đánh Cấp"),
                    ("CTSV_ThemVatLieu_TrenCatNgang", "Thêm Vật Liệu Trên Cắt Ngang"),
                    ("CTSV_ThayDoi_MSS_Min_Max", "Thay Đổi MSS Min Max"),
                    ("CTSV_ThayDoi_GioiHan_traiPhai", "Thay Đổi Giới Hạn Trái Phải"),
                    ("CTSV_ThayDoi_KhungIn", "Thay Đổi Khung In"),
                    ("CTSV_KhoaCatNgang_AddPoint", "Khóa Cắt Ngang Add Point"),
                    ("CTSV_fit_KhungIn", "Fit Khung In"),
                    ("CTSV_fit_KhungIn_5_5_top", "Fit Khung In 5x5"),
                    ("CTSV_fit_KhungIn_5_10_top", "Fit Khung In 5x10"),
                    ("CTSV_An_DuongDiaChat", "Ẩn Đường Địa Chất"),
                    ("CTSV_HieuChinh_Section", "Hiệu Chỉnh Section Static"),
                    ("CTSV_HieuChinh_Section_Dynamic", "Hiệu Chỉnh Section Dynamic")
                ];

                // Các panel còn lại không có lệnh
                string[] emptyPanels =
                [
                    "Alignment",
                    "Intersections", 
                    "Feature Line",
                    "Assembly",
                    "Grading"
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

                // Thêm các panel có lệnh cho Civil tool
                AddCivilDropdownPanel(tab, "Corridor", corridorCommands);
                AddCivilDropdownPanel(tab, "Parcel", parcelCommands);
                AddCivilDropdownPanel(tab, "Pipe Network", pipeCommands);
                AddCivilDropdownPanel(tab, "Point", pointCommands);
                AddCivilDropdownPanel(tab, "Profile", profileCommands);
                AddCivilDropdownPanel(tab, "Sampleline", samplelineCommands);
                AddCivilDropdownPanel(tab, "Section View", sectionviewCommands);


                // Thêm các panel còn lại (không có lệnh)
                foreach (var panel in emptyPanels)
                {
                    AddPanel(tab, panel);
                }

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
                    
                    if (parameter is RibbonButton btn && btn.Tag is string tag)
                    {
                        commandToRun = tag;
                    }

                    var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    var ed = doc?.Editor;

                    if (!string.IsNullOrWhiteSpace(commandToRun) && doc != null)
                    {
                        // Execute the command
                        doc.SendStringToExecute(commandToRun + " ", true, false, true);
                    }
                    else
                    {
                        ed?.WriteMessage($"\nLệnh không thể thực thi.");
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