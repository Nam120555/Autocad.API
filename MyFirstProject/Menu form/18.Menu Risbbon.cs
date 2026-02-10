using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

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
    /// <summary>
    /// Helper class for creating ribbon button icons
    /// </summary>
    public static class RibbonIconHelper
    {
        // Dictionary để cache icons đã tạo
        private static readonly Dictionary<string, BitmapImage> _iconCache = [];

        /// <summary>
        /// Tạo icon từ text/emoji (32x32 pixels)
        /// </summary>
        public static BitmapImage? CreateTextIcon(string text, int size = 32)
        {
            if (string.IsNullOrEmpty(text)) return null;
            
            // Lấy emoji/ký tự đầu tiên từ text
            string iconText = ExtractIconChar(text);
            
            // Check cache
            string cacheKey = $"{iconText}_{size}";
            if (_iconCache.TryGetValue(cacheKey, out var cached))
                return cached;

            try
            {
                using var bitmap = new Bitmap(size, size);
                using var graphics = Graphics.FromImage(bitmap);
                
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                
                // Background gradient
                using (var brush = new System.Drawing.Drawing2D.LinearGradientBrush(
                    new Rectangle(0, 0, size, size),
                    System.Drawing.Color.FromArgb(50, 120, 180),
                    System.Drawing.Color.FromArgb(30, 80, 140),
                    45f))
                {
                    graphics.FillRectangle(brush, 0, 0, size, size);
                }
                
                // Border
                using (var pen = new System.Drawing.Pen(System.Drawing.Color.FromArgb(70, 150, 200), 1))
                {
                    graphics.DrawRectangle(pen, 0, 0, size - 1, size - 1);
                }
                
                // Text/Icon
                using var font = new System.Drawing.Font("Segoe UI Symbol", size * 0.5f, System.Drawing.FontStyle.Bold);
                using var textBrush = new SolidBrush(System.Drawing.Color.White);
                
                var textSize = graphics.MeasureString(iconText, font);
                float x = (size - textSize.Width) / 2;
                float y = (size - textSize.Height) / 2;
                graphics.DrawString(iconText, font, textBrush, x, y);

                // Convert to BitmapImage
                var bitmapImage = ConvertToBitmapImage(bitmap);
                if (bitmapImage != null)
                {
                    _iconCache[cacheKey] = bitmapImage;
                }
                return bitmapImage;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Lấy ký tự icon từ text (emoji hoặc chữ cái đầu)
        /// </summary>
        private static string ExtractIconChar(string text)
        {
            if (string.IsNullOrEmpty(text)) return "?";
            
            // Unicode emoji characters thường dùng
            var iconChars = new[] { "◎", "▸", "◈", "◊", "═", "▤", "□", "◻", "▢", "▣", 
                                    "⊙", "⊚", "⊕", "⊗", "★", "☆", "●", "○", "◐", "◑",
                                    "▶", "◀", "▲", "▼", "◄", "►", "◆", "◇", "⬇", "⬆",
                                    "━", "─", "│", "║", "╱", "╲", "⚙", "⚡", "▭" };
            
            foreach (var c in iconChars)
            {
                if (text.Contains(c))
                    return c;
            }
            
            // Fallback: lấy chữ cái đầu tiên (bỏ qua khoảng trắng)
            foreach (char c in text)
            {
                if (char.IsLetterOrDigit(c))
                    return c.ToString().ToUpper();
            }
            
            return "?";
        }

        /// <summary>
        /// Convert System.Drawing.Bitmap to WPF BitmapImage
        /// </summary>
        private static BitmapImage? ConvertToBitmapImage(Bitmap bitmap)
        {
            try
            {
                using var memory = new MemoryStream();
                bitmap.Save(memory, ImageFormat.Png);
                memory.Position = 0;
                
                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                bitmapImage.Freeze();
                
                return bitmapImage;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Tạo icon màu từ text với custom color
        /// </summary>
        public static BitmapImage? CreateColorIcon(string text, System.Drawing.Color bgColor, int size = 32)
        {
            if (string.IsNullOrEmpty(text)) return null;
            
            string iconText = ExtractIconChar(text);
            string cacheKey = $"{iconText}_{bgColor.ToArgb()}_{size}";
            
            if (_iconCache.TryGetValue(cacheKey, out var cached))
                return cached;

            try
            {
                using var bitmap = new Bitmap(size, size);
                using var graphics = Graphics.FromImage(bitmap);
                
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                
                // Background với màu tùy chỉnh
                using (var brush = new SolidBrush(bgColor))
                {
                    graphics.FillRectangle(brush, 0, 0, size, size);
                }
                
                // Text
                using var font = new System.Drawing.Font("Segoe UI Symbol", size * 0.5f, System.Drawing.FontStyle.Bold);
                using var textBrush = new SolidBrush(System.Drawing.Color.White);
                
                var textSize = graphics.MeasureString(iconText, font);
                float x = (size - textSize.Width) / 2;
                float y = (size - textSize.Height) / 2;
                graphics.DrawString(iconText, font, textBrush, x, y);

                var bitmapImage = ConvertToBitmapImage(bitmap);
                if (bitmapImage != null)
                {
                    _iconCache[cacheKey] = bitmapImage;
                }
                return bitmapImage;
            }
            catch
            {
                return null;
            }
        }
    }

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

                // Create new Civil tool tab - CHỈ 1 TAB DUY NHẤT
                RibbonTab tab = new()
                {
                    Title = "Civil Tool",
                    Id = "MyFirstProject.C3DTab"
                };
                ribbon.Tabs.Add(tab);

                // Helper to add a dropdown menu for Civil tool commands WITH ICONS
                void AddCivilDropdownPanel(RibbonTab targetTab, string panelTitle, (string Command, string Label)[] commands)
                {
                    if (commands.Length == 0) return; // Skip if no commands

                    RibbonPanelSource src = new() { Title = panelTitle };
                    RibbonPanel panel = new() { Source = src };
                    
                    // Panel icon từ lệnh đầu tiên
                    var panelIcon = RibbonIconHelper.CreateTextIcon(commands[0].Label, 32);
                    var panelIconSmall = RibbonIconHelper.CreateTextIcon(commands[0].Label, 16);
                    
                    RibbonSplitButton splitButton = new()
                    {
                        Text = panelTitle,
                        ShowText = true,
                        ShowImage = panelIcon != null,
                        Image = panelIconSmall,
                        LargeImage = panelIcon,
                        Size = RibbonItemSize.Large
                    };
                    
                    foreach (var (command, label) in commands)
                    {
                        // Tạo icon từ label (sử dụng emoji/ký tự đầu)
                        var icon = RibbonIconHelper.CreateTextIcon(label, 16);
                        var largeIcon = RibbonIconHelper.CreateTextIcon(label, 32);
                        
                        RibbonButton btn = new()
                        {
                            Text = label,
                            ShowText = true,
                            ShowImage = icon != null,
                            Image = icon,
                            LargeImage = largeIcon,
                            Orientation = System.Windows.Controls.Orientation.Horizontal,
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

                // ══════════════════════════════════════════════════════════════════════════
                // ICON PHONG CÁCH KỸ SƯ GIAO THÔNG - ĐEN TRẮNG, CHUYÊN NGHIỆP
                // ══════════════════════════════════════════════════════════════════════════

                // 1. Module Surfaces (09.Surfaces.cs)
                (string Command, string Label)[] surfacesCommands =
                [
                    ("CTS_TaoSpotElevation_OnSurface_TaiTim", "◎ Spot Elevation Tại Tim")
                ];
 
                // 2. Module SampleLine (07.Sampleline.cs) - Quản lý cọc/lý trình
                (string Command, string Label)[] samplelineCommands =
                [
                    // ── ĐỔI TÊN CỌC ──
                    ("CTS_DoiTenCoc", "▸ Đổi tên cọc"),
                    ("CTS_DoiTenCoc2", "▸ Đổi tên cọc đoạn"),
                    ("CTS_DoiTenCoc3", "▸ Đổi tên cọc Km"),
                    ("CTS_DoiTenCoc_fromCogoPoint", "▸ Đổi tên từ CogoPoint"),
                    ("CTS_DoiTenCoc_TheoThuTu", "▸ Đổi tên thứ tự"),
                    ("CTS_DoiTenCoc_H", "▸ Đổi tên hậu tố A"),
                    // ── TỌA ĐỘ CỌC ──
                    ("CTS_TaoBang_ToaDoCoc", "◈ Tọa độ cọc (X,Y)"),
                    ("CTS_TaoBang_ToaDoCoc2", "◈ Tọa độ cọc (Lý trình)"),
                    ("CTS_TaoBang_ToaDoCoc3", "◈ Tọa độ cọc (Cao độ)"),
                    ("AT_UPdate2Table", "⟳ Cập nhật từ bảng"),
                    // ── PHÁT SINH CỌC ──
                    ("CTS_ChenCoc_TrenTracDoc", "⊕ Chèn trên trắc dọc"),
                    ("CTS_CHENCOC_TRENTRACNGANG", "⊕ Chèn trên trắc ngang"),
                    ("CTS_PhatSinhCoc", "⊕ Phát sinh cọc auto"),
                    ("CTS_PhatSinhCoc_ChiTiet", "⊕ Phát sinh chi tiết"),
                    ("CTS_PhatSinhCoc_theoKhoangDelta", "⊕ Phát sinh delta"),
                    ("CTS_PhatSinhCoc_TuCogoPoint", "⊕ Từ CogoPoint"),
                    ("CTS_PhatSinhCoc_TheoBang", "⊕ Từ bảng"),
                    // ── DỊCH CỌC ──
                    ("CTS_DichCoc_TinhTien", "⇄ Dịch cọc tịnh tiến"),
                    ("CTS_DichCoc_TinhTien40", "⇄ Dịch cọc 40m"),
                    ("CTS_DichCoc_TinhTien_20", "⇄ Dịch cọc 20m"),
                    // ── ĐỒNG BỘ ──
                    ("CTS_Copy_NhomCoc", "⧉ Sao chép nhóm cọc"),
                    ("CTS_DongBo_2_NhomCoc", "⟳ Đồng bộ nhóm cọc"),
                    ("CTS_DongBo_2_NhomCoc_TheoDoan", "⟳ Đồng bộ theo đoạn"),
                    // ── BỀ RỘNG ──
                    ("CTS_Copy_BeRong_sampleLine", "⊢ Copy bề rộng SL"),
                    ("CTS_Thaydoi_BeRong_sampleLine", "⊢ Thay đổi bề rộng SL"),
                    ("CTS_Offset_BeRong_sampleLine", "⊢ Offset bề rộng SL"),
                    // ── THỐNG KÊ ──
                    ("CTSV_ThongKeCoc", "▤ Thống kê cọc (Excel)"),
                    ("CTSV_ThongKeCoc_TatCa", "▤ Thống kê tất cả cọc")
                ];

                // 3. Module Alignment & Profile - TRẮC DỌC
                (string Command, string Label)[] profileCommands =
                [
                    ("CTPV_TaoProfileView", "▬ Tạo trắc dọc"),
                    ("CTPV_ThemBang_LyTrinh", "▦ Thêm bảng lý trình"),
                    ("CTPV_ThemLabel_CaoDo", "▭ Thêm Label cao độ"),
                    ("CTPV_SuaProfileView", "◇ Edit profile"),
                    ("CTPV_ThayDoiScale", "◊ Thay đổi Scale"),
                    ("CTPV_FitKhung", "▢ Fit khung")
                ];
 
                // 4. Module Corridor & Parcel
                (string Command, string Label)[] corridorCommands =
                [
                    ("CTC_AddAllSection", "⊞ Thêm tất cả Section"),
                    ("CTC_TaoCooridor_DuongDoThi_RePhai", "⤻ Corridor rẽ phải"),
                    ("CTP_TaoBangThongKeParcel", "▣ Thống kê Parcel"),
                    ("CTP_TaoBangThongKeParcel_SapXep", "▣ Thống kê Parcel (SX)")
                ];

                // 5. Module SectionView - TRẮC NGANG
                (string Command, string Label)[] sectionviewCommands =
                [
                    // ── TẠO TRẮC NGANG ──
                    ("CTSV_VeTracNgangThietKe", "╋ Tạo trắc ngang"),
                    ("CVSV_VeTatCa_TracNgangThietKe", "╋ Vẽ tất cả TN"),
                    ("CTSV_ChuyenDoi_TNTK_TNTN", "⟳ Chuyển TK sang TN"),
                    // ── ĐÁNH CẤP ──
                    ("CTSV_DanhCap", "△ Đánh cấp - VHC"),
                    ("CTSV_DanhCap_XoaBo", "⊘ Xóa đánh cấp"),
                    ("CTSV_DanhCap_VeThem", "⊕ Vẽ thêm đánh cấp"),
                    ("CTSV_DanhCap_VeThem1", "⊕ Vẽ thêm 1m"),
                    ("CTSV_DanhCap_VeThem2", "⊕ Vẽ thêm 2m"),
                    ("CTSV_DanhCap_CapNhat", "⟳ Cập nhật KL đánh cấp"),
                    ("CTSV_ThemVatLieu_TrenCatNgang", "▤ Điền KL trắc ngang"),
                    // ── HIỆU CHỈNH ──
                    ("CTSV_ThayDoi_MSS_Min_Max", "⚙ Hiệu chỉnh MSS"),
                    ("CTSV_ThayDoi_GioiHan_traiPhai", "⇄ Thay giới hạn T/P"),
                    ("CTSV_ThayDoi_KhungIn", "▦ Dàn khung in"),
                    ("CTSV_KhoaCatNgang_AddPoint", "⊠ Khóa TN + Add Point"),
                    ("CTSV_fit_KhungIn", "▢ Fit khung in"),
                    ("CTSV_fit_KhungIn_5_5_top", "▢ Fit khung 5x5"),
                    ("CTSV_fit_KhungIn_5_10_top", "▢ Fit khung 5x10"),
                    ("CTSV_An_DuongDiaChat", "◌ Ẩn đường địa chất"),
                    ("CTSV_HieuChinh_Section", "◇ Hiệu chỉnh (Static)"),
                    ("CTSV_HieuChinh_Section_Dynamic", "◆ Hiệu chỉnh (Dynamic)"),
                    // ── THỐNG KÊ ──
                    ("CTSV_ThongKeCoc", "▤ Thống kê cọc (Excel)"),
                    ("CTSV_ThongKeCoc_TatCa", "▤ Thống kê toàn bộ cọc"),
                    ("CTSV_ThongKeCoc_ToaDo", "◎ Thống kê tọa độ cọc"),
                    // ── KHỐI LƯỢNG ──
                    ("CTSV_Taskbar", "▥ Taskbar Khối Lượng"),
                    ("CTSV_XuatKhoiLuong", "⬇ Xuất KL Excel"),
                    ("CTSV_XuatCad", "⬇ Xuất KL CAD"),
                    ("CTSV_CaiDatBang", "⚙ Cài đặt bảng KL")
                ];

                // 6. Module San Nền (14.SanNen.cs)
                (string Command, string Label)[] gradingCommands =
                [
                    ("CTSN_Taskbar", "▥ Mở Taskbar SN"),
                    ("CTSN_TaoLuoi", "▦ Quản lý lưới"),
                    ("CTSN_NhapCaoDo", "▭ Điền cao độ lưới"),
                    ("CTSN_Surface", "◬ Lấy CĐ Surface"),
                    ("CTSN_TinhKL", "▤ Tính khối lượng SN"),
                    ("CTSN_XuatBang", "⬇ Xuất bảng KL CAD")
                ];

                // 7. Module Point (05.Point.cs)
                (string Command, string Label)[] pointCommands =
                [
                    ("CTPo_TaoPointTheoBang", "⊕ Tạo Point từ bảng"),
                    ("CTPo_ChuyenPointThanhBlock", "⟳ Point → Block"),
                    ("CTPo_TaoBangThongKePoint", "▤ Bảng thống kê Point"),
                    ("CTPo_ThayDoiCaoDo", "◇ Thay đổi cao độ"),
                    ("CTPo_DatTen_theoThuTu", "▸ Đặt tên thứ tự"),
                    ("CTPo_ThayDoiStyle", "◈ Thay đổi Style"),
                    ("CTPo_LayThongTin", "ⓘ Lấy thông tin")
                ];

                // 8. Module Pipe & Structures - THOÁT NƯỚC
                (string Command, string Label)[] pipeCommands =
                [
                    ("CTPS_TaoBangThongKePipe", "⊙ Thống kê Pipe"),
                    ("CTPS_TaoBangThongKeStructure", "⊙ Thống kê Structure"),
                    ("CTPS_ThayDoi_CaoDo_Pipe", "◇ Đổi cao độ Pipe"),
                    ("CTPS_ThayDoi_CaoDo_Structure", "◇ Đổi cao độ Struct"),
                    ("CTPS_XoayPipe_90do", "⤾ Xoay Pipe 90°"),
                    ("CTPS_XoaConTrung", "⊘ Xóa con trùng")
                ];

                // 9. Module Utilities & Property Sets
                (string Command, string Label)[] utilitiesCommands =
                [
                    ("AT_Solid_Set_PropertySet", "⊞ Gán Property Set"),
                    ("AT_Solid_Show_Info", "ⓘ Thông tin Solid"),
                    ("CT_VTOADOHG", "◎ Tọa độ hố ga"),
                    ("CT_DanhSachLenh", "▤ Danh sách lệnh"),
                    ("show_menu", "⟳ Reload Menu")
                ];



                // 10. CAD Commands - CÔNG CỤ CAD
                (string Command, string Label)[] acadCommands =
                [
                    ("AT_TongDoDai_Full", "━ Tổng Độ Dài (Full)"),
                    ("AT_TongDoDai_Replace", "━ Tổng Độ Dài (Replace)"),
                    ("AT_TongDoDai_Replace2", "━ Tổng Độ Dài (Replace2)"),
                    ("AT_TongDoDai_Replace_CongThem", "━ Tổng Độ Dài (Cộng Thêm)"),
                    ("AT_TongDienTich_Full", "▢ Tổng Diện Tích (Full)"),
                    ("AT_TongDienTich_Replace", "▢ Tổng Diện Tích (Replace)"),
                    ("AT_TongDienTich_Replace2", "▢ Tổng Diện Tích (Replace2)"),
                    ("AT_TongDienTich_Replace_CongThem", "▢ Tổng Diện Tích (Cộng Thêm)"),
                    ("AT_TextLink", "⊙ Text Link"),
                    ("AT_DanhSoThuTu", "▸ Đánh Số Thứ Tự"),
                    ("AT_XoayDoiTuong_TheoViewport", "⤾ Xoay Theo Viewport"),
                    ("AT_XoayDoiTuong_Theo2Diem", "⤾ Xoay Theo 2 Điểm"),
                    ("AT_TextLayout", "▭ Text Layout"),
                    ("AT_TaoMoi_TextLayout", "▭ Tạo Mới Text Layout"),
                    ("AT_DimLayout2", "⊢ Dim Layout 2"),
                    ("AT_DimLayout", "⊢ Dim Layout"),
                    ("AT_BlockLayout", "▣ Block Layout"),
                    ("AT_Label_FromText", "▭ Label From Text"),
                    ("AT_XoaDoiTuong_CungLayer", "⊘ Xóa Cùng Layer"),
                    ("AT_XoaDoiTuong_3DSolid_Body", "⊘ Xóa 3DSolid/Body"),
                    ("AT_UpdateLayout", "⟳ Update Layout"),
                    ("AT_Offset_2Ben", "⇄ Offset 2 Bên"),
                    ("AT_annotive_scale_currentOnly", "◈ Annotative Scale")
                ];

                // 11. LAYER CONTROL - Bật/Tắt Layer (từ LISP)
                (string Command, string Label)[] layerCommands =
                [
                    ("CTL_OnCorridor", "◎ BẬT Corridor"),
                    ("CTL_OffCorridor", "⊘ TẮT Corridor"),
                    ("CTL_OnSampleLine", "◎ BẬT SampleLine"),
                    ("CTL_OffSampleLine", "⊘ TẮT SampleLine"),
                    ("CTL_OnAlignment", "◎ BẬT Alignment"),
                    ("CTL_OffAlignment", "⊘ TẮT Alignment"),
                    ("CTL_OnParcel", "◎ BẬT Parcel"),
                    ("CTL_OffParcel", "⊘ TẮT Parcel"),
                    ("CTL_OnHatchDaoDap", "◎ BẬT Hatch Đào Đắp"),
                    ("CTL_OffHatchDaoDap", "⊘ TẮT Hatch Đào Đắp"),
                    ("CTL_OnDefpoints", "◎ BẬT Defpoints"),
                    ("CTL_OffDefpoints", "⊘ TẮT Defpoints")
                ];

                // 12. LISP UTILITY - Tiện ích từ LISP
                (string Command, string Label)[] lispUtilityCommands =
                [
                    ("CTS_RebuildSurface", "⟳ Rebuild Surface"),
                    ("CTPo_ReorderPoints", "▸ Đánh số Point"),
                    ("CTP_AddParcelLabels", "▭ Thêm nhãn Parcel"),
                    ("CTU_ExportCAD2007", "⬇ Export CAD 2007"),
                    ("CTU_ExplodeAEC", "⊘ Explode AEC"),
                    ("CTU_StyleAutoOn", "◎ BẬT Auto Style"),
                    ("CTU_StyleAutoOff", "⊘ TẮT Auto Style"),
                    ("CTU_DumpObject", "ⓘ Dump Object Info")
                ];

                // 13. DRAWING SETUP - Thiết lập bản vẽ (MỚI từ LISP)
                (string Command, string Label)[] drawingSetupCommands =
                [
                    ("CTDS_ThietLap", "⚙ Thiết lập chuẩn"),
                    ("CTDS_SaveClean", "⊘ Save & Purge"),
                    ("CTDS_PrintAllLayouts", "▤ In tất cả Layout"),
                    ("CTDS_PrintCurrentLayout", "▦ In Layout hiện tại"),
                    ("CTDS_ExportPDF", "⬇ Xuất PDF"),
                    ("CTDS_ConvertMM2M", "⇄ Chuyển MM→M"),
                    ("CTDS_ConvertCM2M", "⇄ Chuyển CM→M")
                ];

                // 14. LAYER QUICK - Đổi layer nhanh (MỚI từ LISP)
                (string Command, string Label)[] layerQuickCommands =
                [
                    ("CTL_ToText", "▸ → 0.TEXT"),
                    ("CTL_ToDefpoints", "▸ → Defpoints"),
                    ("CTL_ToDim", "▸ → 1.DIM"),
                    ("CTL_ToBaoBT", "▸ → 2.BAO BT"),
                    ("CTL_ToBaoCotThep", "▸ → 3.BAO COT THEP"),
                    ("CTL_ToThep", "▸ → 4.THEP"),
                    ("CTL_ToTruc", "▸ → 5.TRUC"),
                    ("CTL_ToKhuat", "▸ → 6.KHUAT"),
                    ("CTL_ToHatch", "▸ → 7.HATCH"),
                    ("CTL_ToRanhGioi", "▸ → 8.RANH GIOI")
                ];

                // 15. COMMON UTILITIES - Tiện ích thường dùng (MỚI từ LISP)
                (string Command, string Label)[] commonUtilitiesCommands =
                [
                    ("CTU_MakePointFromText", "⊕ Tạo Point từ Text"),
                    ("CTU_TotalLength", "━ Tổng chiều dài"),
                    ("CTU_ExportTextCoords", "⬇ Xuất tọa độ Text"),
                    ("CTU_TextToMText", "▭ Text → MText"),
                    ("CTU_FindIntersections", "◎ Tìm điểm giao"),
                    ("CTU_AddPolylineVertices", "⊕ Thêm đỉnh Polyline"),
                    ("CTU_DrawTaluy", "╱ Vẽ Taluy")
                ];

                // 16. CURVE DESIGN STANDARDS - Tiêu chuẩn thiết kế đường cong
                (string Command, string Label)[] curveDesignCommands =
                [
                    ("CTC_ThietLapDuongCong", "⊙ Mở Form Đường Cong"),
                    ("CTC_TraCuuDuongCong", "◎ Tra cứu thông số"),
                    ("CTC_ThongSoDuongCong_4054", "▤ Bảng TCVN 4054 (ngoài ĐT)"),
                    ("CTC_ThongSoDuongCong_13592", "▤ Bảng TCVN 13592 (đô thị)"),
                    ("CTC_KiemTraDuongCong_4054", "⚙ Kiểm tra theo 4054"),
                    ("CTC_KiemTraDuongCong_13592", "⚙ Kiểm tra theo 13592")
                ];
                // ══════════════════════════════════════════════════════════════════════════
                // CẤU TRÚC MENU THEO QUY TRÌNH THIẾT KẾ GIAO THÔNG
                // ══════════════════════════════════════════════════════════════════════════
                // 1. Bề mặt   - Surface + Point (dữ liệu nền)
                // 2. Cọc      - SampleLine (lý trình cọc)  
                // 3. Tuyến    - Tạo Profile View
                // 4. Trắc dọc - Edit Profile
                // 5. Corridor - Corridor + Parcel
                // 6. Trắc ngang - SectionView + Đánh cấp
                // 7. Ngoại giao - San nền + Pipe + External
                // 8. Thống kê - Export + Utilities  
                // 9. Hướng dẫn - CAD Commands
                // ══════════════════════════════════════════════════════════════════════════

                // ▶ PANEL 1: BỀ MẶT
                var panel1 = surfacesCommands.Concat(pointCommands).ToArray();
                AddCivilDropdownPanel(tab, "Bề mặt", panel1);
                
                // ▶ PANEL 2: CỌC
                AddCivilDropdownPanel(tab, "Cọc", samplelineCommands);

                // ▶ PANEL 3: TUYẾN (Tạo)
                AddCivilDropdownPanel(tab, "Tuyến", profileCommands);

                // ▶ PANEL 4: TRẮC DỌC (Edit - tách riêng từ profileCommands phía trên)
                // Đã gộp trong profileCommands

                // ▶ PANEL 5: CORRIDOR
                AddCivilDropdownPanel(tab, "Corridor", corridorCommands);

                // ▶ PANEL 6: TRẮC NGANG
                AddCivilDropdownPanel(tab, "Trắc ngang", sectionviewCommands);

                // ▶ PANEL 7: NGOẠI GIAO (San nền + Pipe)
                var panel7 = gradingCommands.Concat(pipeCommands).ToArray();
                AddCivilDropdownPanel(tab, "Ngoại giao", panel7);

                // ▶ PANEL 8: THỐNG KÊ (Tiện ích + Export)
                AddCivilDropdownPanel(tab, "Thống kê", utilitiesCommands);

                // ▶ PANEL 9: LAYER (Bật/Tắt Layer)
                AddCivilDropdownPanel(tab, "Layer", layerCommands);

                // ▶ PANEL 10: LISP UTILITY (Tiện ích từ LISP)
                AddCivilDropdownPanel(tab, "LISP", lispUtilityCommands);

                // ▶ PANEL 11: HƯỚNG DẪN (CAD Commands)
                AddCivilDropdownPanel(tab, "Hướng dẫn", acadCommands);

                // ▶ PANEL 12: THIẾT LẬP BẢN VẼ (Mới từ LISP)
                AddCivilDropdownPanel(tab, "Thiết lập", drawingSetupCommands);

                // ▶ PANEL 13: ĐỔI LAYER NHANH (Mới từ LISP)
                AddCivilDropdownPanel(tab, "Đổi Layer", layerQuickCommands);

                // ▶ PANEL 14: TIỆN ÍCH THƯỜNG DÙNG (Mới từ LISP)
                AddCivilDropdownPanel(tab, "Tiện ích", commonUtilitiesCommands);

                // ▶ PANEL 15: TIÊU CHUẨN ĐƯỜNG CONG (TCVN 4054 & 13592)
                AddCivilDropdownPanel(tab, "Đường cong", curveDesignCommands);

                // ▶ PANEL 16: CÀI ĐẶT (Icons & Danh sách lệnh)
                (string Command, string Label)[] settingsCommands =
                [
                    ("CT_DoiIcon", "🎨 Đổi Icon lệnh"),
                    ("CT_DanhSachLenh", "📋 Danh sách lệnh"),
                    ("show_menu", "▤ Làm mới Ribbon")
                ];
                AddCivilDropdownPanel(tab, "Cài đặt", settingsCommands);

                tab.IsActive = true;
                var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor;
                ed?.WriteMessage("\nĐã tạo tab 'Civil Tool' với đầy đủ các lệnh trên Ribbon.");
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