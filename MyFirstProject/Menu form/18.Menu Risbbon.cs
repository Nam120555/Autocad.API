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
using Helpers = MyFirstProject.Helpers;

[assembly: CommandClass(typeof(MyFirstProject.Autocad))]

namespace MyFirstProject
{
    /// <summary>
    /// Helper class for creating ribbon button icons
    /// </summary>
    public static class RibbonIconHelper
    {
        private static readonly Dictionary<string, BitmapImage> _iconCache = [];

        /// <summary>
        /// Tạo icon từ text/emoji với màu sắc tùy chỉnh (Nền trong suốt kiểu Civil 3D)
        /// </summary>
        public static BitmapImage? CreateTextIcon(string text, int size = 32, System.Drawing.Color? color = null)
        {
            if (string.IsNullOrEmpty(text)) return null;
            
            string iconText = ExtractIconChar(text);
            color ??= System.Drawing.Color.Black; // Mặc định đen nếu không chỉ định

            string cacheKey = $"{iconText}_{size}_{color.Value.ToArgb()}";
            if (_iconCache.TryGetValue(cacheKey, out var cached))
                return cached;

            try
            {
                using var bitmap = new Bitmap(size, size);
                using var graphics = Graphics.FromImage(bitmap);
                
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                graphics.Clear(System.Drawing.Color.Transparent); // Nền trong suốt
                
                // Vẽ ký hiệu
                float fontSize = size * 0.7f;
                using var font = new System.Drawing.Font("Segoe UI Symbol", fontSize, System.Drawing.FontStyle.Bold);
                using var textBrush = new SolidBrush(color.Value);
                
                var textSize = graphics.MeasureString(iconText, font);
                float x = (size - textSize.Width) / 2;
                float y = (size - textSize.Height) / 2;
                
                // Đổ bóng nhẹ cho chuyên nghiệp
                using (var shadowBrush = new SolidBrush(System.Drawing.Color.FromArgb(100, 100, 100)))
                {
                    graphics.DrawString(iconText, font, shadowBrush, x + 1, y + 1);
                }
                
                graphics.DrawString(iconText, font, textBrush, x, y);

                var bitmapImage = ConvertToBitmapImage(bitmap);
                if (bitmapImage != null) _iconCache[cacheKey] = bitmapImage;
                return bitmapImage;
            }
            catch { return null; }
        }

        private static string ExtractIconChar(string text)
        {
            if (string.IsNullOrEmpty(text)) return "?";
            var iconChars = new[] { "◎", "▸", "◈", "◊", "═", "▤", "□", "◻", "▢", "▣", 
                                    "⊙", "⊚", "⊕", "⊗", "★", "☆", "●", "○", "◐", "◑",
                                    "▶", "◀", "▲", "▼", "◄", "►", "◆", "◇", "⬇", "⬆",
                                    "━", "─", "│", "║", "╱", "╲", "⚙", "⚡", "▭", "╋", "△", "⊞", "▥" };
            
            foreach (var c in iconChars) if (text.Contains(c)) return c;
            foreach (char c in text) if (char.IsLetterOrDigit(c)) return c.ToString().ToUpper();
            return "?";
        }

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
            catch { return null; }
        }

        public static BitmapImage? CreateColorIcon(string text, System.Drawing.Color bgColor, int size = 32)
        {
            // Dự phòng cho backward compatibility hoặc dùng cho nút đặc biệt
            return CreateTextIcon(text, size, bgColor);
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
                    var doc0 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    doc0?.SendStringToExecute("RIBBON ", true, false, false);
                    ribbon = ComponentManager.Ribbon;
                    if (ribbon == null) { doc0?.Editor.WriteMessage("\nKhông thể khởi tạo Ribbon."); return; }
                }

                // Xóa tab cũ
                foreach (var oldId in new[] { "MyFirstProject.C3DTab", "MyFirstProject.AcadTab", "MyFirstProject.VITab" })
                {
                    var old = ribbon.Tabs.FirstOrDefault(t => t.Id == oldId);
                    if (old != null) ribbon.Tabs.Remove(old);
                }

                RibbonTab tab = new() { Title = "Civil Tool", Id = "MyFirstProject.C3DTab" };
                ribbon.Tabs.Add(tab);

                // ══════════════════════════════════════════════════════════════
                // PANEL 1: BỀ MẶT (Surface)
                // ══════════════════════════════════════════════════════════════
                var p1 = Helpers.RibbonFactory.CreatePanel("Bề mặt");
                p1.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Địa hình", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTS_TaoSpotElevation_OnSurface_TaiTim", "Spot Elevation", "Tạo cao độ tại tim tuyến trên Surface."),
                    Helpers.RibbonFactory.CreateButton("CTS_RebuildSurface", "Rebuild Surface", "Cập nhật lại bề mặt địa hình."),
                    Helpers.RibbonFactory.CreateButton("CTSV_SoSanhSurface", "So sánh Surface", "So sánh 2 bề mặt để tính khối lượng chênh lệch."),
                    Helpers.RibbonFactory.CreateButton("CTSV_DaoDap", "Đào đắp", "Tính khối lượng đào đắp giữa 2 bề mặt."),
                }));
                p1.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Material", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTSV_HienThiMaterialList", "Hiển thị ML", "Hiển thị danh sách Material List."),
                    Helpers.RibbonFactory.CreateButton("CTSV_ChiTietMaterialSection", "Chi tiết MS", "Xem chi tiết Material Section."),
                    Helpers.RibbonFactory.CreateButton("CTSV_VeDuongBaoMaterial", "Vẽ đường bao", "Vẽ đường bao Material Section trong SectionView."),
                    Helpers.RibbonFactory.CreateButton("CTSV_PhanTichArea", "Phân tích Area", "Phân tích diện tích từ Material Section."),
                }));
                tab.Panels.Add(p1);

                // ══════════════════════════════════════════════════════════════
                // PANEL 2: CỌC (Sample Lines / Staking)
                // ══════════════════════════════════════════════════════════════
                var p2 = Helpers.RibbonFactory.CreatePanel("Cọc");
                p2.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Đổi tên cọc", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTS_DoiTenCoc", "Đổi tên cọc", "Đổi tên cọc theo chuẩn Km+m."),
                    Helpers.RibbonFactory.CreateButton("CTS_DoiTenCoc2", "Theo đoạn", "Đổi tên cọc theo đoạn tuyến."),
                    Helpers.RibbonFactory.CreateButton("CTS_DoiTenCoc3", "Theo Km", "Đổi tên cọc theo Km."),
                    Helpers.RibbonFactory.CreateButton("CTS_DoiTenCoc_TheoThuTu", "Theo thứ tự", "Đổi tên cọc theo thứ tự 1, 2, 3..."),
                    Helpers.RibbonFactory.CreateButton("CTS_DoiTenCoc_H", "Đổi tên H", "Đổi tên cọc kiểu H."),
                    Helpers.RibbonFactory.CreateButton("CTS_DoiTenCoc_fromCogoPoint", "Từ CogoPoint", "Đổi tên cọc từ tên CogoPoint."),
                }));
                p2.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Tọa độ cọc", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTS_TaoBang_ToaDoCoc", "Bảng tọa độ", "Tạo bảng tọa độ cọc."),
                    Helpers.RibbonFactory.CreateButton("CTS_TaoBang_ToaDoCoc2", "Có lý trình", "Bảng tọa độ cọc có lý trình."),
                    Helpers.RibbonFactory.CreateButton("CTS_TaoBang_ToaDoCoc3", "Có cao độ", "Bảng tọa độ cọc có cao độ."),
                    Helpers.RibbonFactory.CreateButton("AT_UPdate2Table", "Update Table", "Cập nhật 2 bảng tọa độ."),
                }));
                p2.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Phát sinh cọc", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTS_PhatSinhCoc", "Phát sinh cọc", "Phát sinh cọc theo khoảng cách."),
                    Helpers.RibbonFactory.CreateButton("CTS_PhatSinhCoc_ChiTiet", "Chi tiết", "Phát sinh cọc chi tiết."),
                    Helpers.RibbonFactory.CreateButton("CTS_PhatSinhCoc_theoKhoangDelta", "Theo Delta", "Phát sinh cọc theo khoảng delta."),
                    Helpers.RibbonFactory.CreateButton("CTS_PhatSinhCoc_TuCogoPoint", "Từ CogoPoint", "Phát sinh cọc từ CogoPoint."),
                    Helpers.RibbonFactory.CreateButton("CTS_PhatSinhCoc_TheoBang", "Theo bảng", "Phát sinh cọc theo bảng Excel."),
                }));
                p2.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Chèn / Dịch cọc", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTS_ChenCoc_TrenTracDoc", "Chèn trắc dọc", "Chèn cọc trên trắc dọc."),
                    Helpers.RibbonFactory.CreateButton("CTS_CHENCOC_TRENTRACNGANG", "Chèn trắc ngang", "Chèn cọc trên trắc ngang."),
                    Helpers.RibbonFactory.CreateButton("CTS_DichCoc_TinhTien", "Dịch cọc", "Dịch cọc tịnh tiến."),
                    Helpers.RibbonFactory.CreateButton("CTS_DichCoc_TinhTien40", "Dịch 40m", "Dịch cọc 40m."),
                    Helpers.RibbonFactory.CreateButton("CTS_DichCoc_TinhTien_20", "Dịch 20m", "Dịch cọc 20m."),
                }));
                p2.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Nhóm cọc / Bề rộng", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTS_Copy_NhomCoc", "Copy nhóm cọc", "Copy nhóm cọc từ tuyến này sang tuyến khác."),
                    Helpers.RibbonFactory.CreateButton("CTS_DongBo_2_NhomCoc", "Đồng bộ 2 nhóm", "Đồng bộ 2 nhóm cọc."),
                    Helpers.RibbonFactory.CreateButton("CTS_DongBo_2_NhomCoc_TheoDoan", "Đồng bộ theo đoạn", "Đồng bộ 2 nhóm cọc theo đoạn."),
                    Helpers.RibbonFactory.CreateButton("CTS_Copy_BeRong_sampleLine", "Copy bề rộng", "Copy bề rộng Sample Line."),
                    Helpers.RibbonFactory.CreateButton("CTS_Thaydoi_BeRong_sampleLine", "Đổi bề rộng", "Thay đổi bề rộng Sample Line."),
                    Helpers.RibbonFactory.CreateButton("CTS_Offset_BeRong_sampleLine", "Offset bề rộng", "Offset bề rộng Sample Line."),
                }));
                tab.Panels.Add(p2);

                // ══════════════════════════════════════════════════════════════
                // PANEL 3: TUYẾN (Alignment & Corridor)
                // ══════════════════════════════════════════════════════════════
                var p3 = Helpers.RibbonFactory.CreatePanel("Tuyến");
                p3.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Đường cong", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTC_ThietLapDuongCong", "Tra bảng TCVN", "Tra cứu thông số đường cong theo TCVN."),
                    Helpers.RibbonFactory.CreateButton("CTC_TraCuuDuongCong", "Tra cứu nhanh", "Tra cứu nhanh thông số đường cong."),
                    Helpers.RibbonFactory.CreateButton("CTC_ThongSoDuongCong_4054", "TCVN 4054", "Bảng tra TCVN 4054:2005."),
                    Helpers.RibbonFactory.CreateButton("CTC_ThongSoDuongCong_13592", "TCVN 13592", "Bảng tra TCVN 13592:2022."),
                    Helpers.RibbonFactory.CreateButton("CTC_KiemTraDuongCong_4054", "Check 4054", "Kiểm tra Alignment theo TCVN 4054."),
                    Helpers.RibbonFactory.CreateButton("CTC_KiemTraDuongCong_13592", "Check 13592", "Kiểm tra Alignment theo TCVN 13592."),
                }));
                p3.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Corridor", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTC_AddAllSection", "Thêm Section", "Thêm tất cả section vào corridor."),
                    Helpers.RibbonFactory.CreateButton("CTC_TaoCooridor_DuongDoThi_RePhai", "Corridor đô thị", "Tạo corridor đường đô thị rẽ phải."),
                    Helpers.RibbonFactory.CreateButton("CVC_CreateCurbReturn_CorridorRegion", "Curb Return", "Tạo vùng Corridor cho nút giao."),
                }));
                p3.Source.Items.Add(Helpers.RibbonFactory.CreateButton("CTPA_TaoParcel_CacLoaiNha", "Tạo Parcel", "Tạo Parcel các loại nhà."));
                tab.Panels.Add(p3);

                // ══════════════════════════════════════════════════════════════
                // PANEL 4: TRẮC DỌC (Profile)
                // ══════════════════════════════════════════════════════════════
                var p4 = Helpers.RibbonFactory.CreatePanel("Trắc dọc");
                p4.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Profile", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTP_VeTracDoc_TuNhien", "Vẽ trắc dọc TN", "Vẽ trắc dọc tự nhiên."),
                    Helpers.RibbonFactory.CreateButton("CTP_VeTracDoc_TuNhien_TatCaTuyen", "TĐ tất cả tuyến", "Vẽ trắc dọc tự nhiên tất cả tuyến."),
                    Helpers.RibbonFactory.CreateButton("CTP_Fix_DuongTuNhien_TheoCoc", "Sửa đường TN", "Sửa đường tự nhiên theo cọc."),
                    Helpers.RibbonFactory.CreateButton("CTP_GanNhanNutGiao_LenTracDoc", "Nhãn nút giao", "Gán nhãn nút giao lên trắc dọc."),
                    Helpers.RibbonFactory.CreateButton("CTP_TaoCogoPointTuPVI", "Point từ PVI", "Tạo CogoPoint từ PVI."),
                }));
                tab.Panels.Add(p4);

                // ══════════════════════════════════════════════════════════════
                // PANEL 5: TRẮC NGANG (SectionView)
                // ══════════════════════════════════════════════════════════════
                var p5 = Helpers.RibbonFactory.CreatePanel("Trắc ngang");
                p5.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Vẽ trắc ngang", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTSv_VeTracNgangThietKe", "Vẽ TN thiết kế", "Vẽ trắc ngang thiết kế."),
                    Helpers.RibbonFactory.CreateButton("CVSV_VeTatCa_TracNgangThietKe", "Vẽ tất cả TNTK", "Vẽ tất cả trắc ngang thiết kế."),
                    Helpers.RibbonFactory.CreateButton("CTSV_ChuyenDoi_TNTK_TNTN", "TK → TN", "Chuyển đổi TN-TK sang TN-TN."),
                }));
                p5.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Đánh cấp", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTSV_DanhCap", "Đánh cấp", "Tính đánh cấp mái taluy."),
                    Helpers.RibbonFactory.CreateButton("CTSV_DanhCap_XoaBo", "Xóa bỏ", "Xóa bỏ đánh cấp."),
                    Helpers.RibbonFactory.CreateButton("CTSV_DanhCap_VeThem", "Vẽ thêm", "Vẽ thêm đánh cấp."),
                    Helpers.RibbonFactory.CreateButton("CTSV_DanhCap_VeThem2", "Vẽ thêm 2m", "Vẽ thêm đánh cấp 2m."),
                    Helpers.RibbonFactory.CreateButton("CTSV_DanhCap_VeThem1", "Vẽ thêm 1m", "Vẽ thêm đánh cấp 1m."),
                    Helpers.RibbonFactory.CreateButton("CTSV_DanhCap_CapNhat", "Cập nhật", "Cập nhật đánh cấp."),
                }));
                p5.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Hiệu chỉnh", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTSV_ThemVatLieu_TrenCatNgang", "Thêm vật liệu", "Thêm vật liệu trên cắt ngang."),
                    Helpers.RibbonFactory.CreateButton("CTSV_ThayDoi_MSS_Min_Max", "MSS Min/Max", "Thay đổi MSS Min Max."),
                    Helpers.RibbonFactory.CreateButton("CTSV_ThayDoi_GioiHan_traiPhai", "Giới hạn T/P", "Thay đổi giới hạn trái phải."),
                    Helpers.RibbonFactory.CreateButton("CTSV_KhoaCatNgang_AddPoint", "Khóa + AddPoint", "Khóa cắt ngang và add point."),
                    Helpers.RibbonFactory.CreateButton("CTSV_HieuChinh_Section", "HC Section (S)", "Hiệu chỉnh section static."),
                    Helpers.RibbonFactory.CreateButton("CTSV_HieuChinh_Section_Dynamic", "HC Section (D)", "Hiệu chỉnh section dynamic."),
                    Helpers.RibbonFactory.CreateButton("CTSV_An_DuongDiaChat", "Ẩn địa chất", "Ẩn đường địa chất."),
                }));
                p5.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Khung in", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTSV_ThayDoi_KhungIn", "Đổi khung in", "Thay đổi khung in."),
                    Helpers.RibbonFactory.CreateButton("CTSV_fit_KhungIn", "Fit khung in", "Fit khung in tự động."),
                    Helpers.RibbonFactory.CreateButton("CTSV_fit_KhungIn_5_5_top", "Fit 5x5", "Fit khung in 5x5."),
                    Helpers.RibbonFactory.CreateButton("CTSV_fit_KhungIn_5_10_top", "Fit 5x10", "Fit khung in 5x10."),
                }));
                tab.Panels.Add(p5);

                // ══════════════════════════════════════════════════════════════
                // PANEL 6: KHỐI LƯỢNG (Volume & Reports)
                // ══════════════════════════════════════════════════════════════
                var p6 = Helpers.RibbonFactory.CreatePanel("Khối lượng");
                p6.Source.Items.Add(Helpers.RibbonFactory.CreateButton("CTSV_Taskbar", "Taskbar KL", "Mở thanh Taskbar quản lý khối lượng."));
                p6.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Xuất khối lượng", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTSV_XuatKhoiLuong", "Xuất Excel", "Xuất bảng khối lượng chi tiết ra Excel."),
                    Helpers.RibbonFactory.CreateButton("CTSV_XuatCad", "Vẽ bảng CAD", "Vẽ trực tiếp bảng khối lượng vào bản vẽ."),
                    Helpers.RibbonFactory.CreateButton("CTSV_CaiDatBang", "Cài đặt bảng", "Cài đặt thông số bảng khối lượng."),
                    Helpers.RibbonFactory.CreateButton("CTSV_TinhKLKetHop", "KL kết hợp", "Tính khối lượng kết hợp nhiều tuyến."),
                }));
                p6.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Phân tích", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTSV_LayDienTichTuSectionView", "Diện tích section", "Lấy diện tích từ SectionView."),
                    Helpers.RibbonFactory.CreateButton("CTSV_LayKhoiLuongTracNgang", "KL trắc ngang", "Lấy khối lượng từ trắc ngang."),
                    Helpers.RibbonFactory.CreateButton("CTSV_XuatSectionArea", "Xuất Section Area", "Xuất Section Area ra Excel."),
                    Helpers.RibbonFactory.CreateButton("CTSV_ThongKeMaterialTracNgang", "Thống kê Material", "Thống kê Material trắc ngang."),
                    Helpers.RibbonFactory.CreateButton("CTSV_KhoiLuongTracNgang", "KL TN chi tiết", "Khối lượng trắc ngang chi tiết."),
                    Helpers.RibbonFactory.CreateButton("CTSV_SoSanhKhoiLuong", "So sánh KL", "So sánh khối lượng."),
                    Helpers.RibbonFactory.CreateButton("CTSV_KiemTraKhoiLuong", "Kiểm tra KL", "Kiểm tra khối lượng."),
                }));
                p6.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Diện tích Poly", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTSV_PolyArea", "Menu PolyArea", "Menu tính diện tích Polyline."),
                    Helpers.RibbonFactory.CreateButton("CTSV_TinhDienTichPoly", "Tính DT Poly", "Tính diện tích Polyline."),
                    Helpers.RibbonFactory.CreateButton("CTSV_TinhDienTichPolyExcel", "DT Poly Excel", "Xuất diện tích Polyline ra Excel."),
                    Helpers.RibbonFactory.CreateButton("CTSV_TinhKhoiLuongPoly", "KL Poly", "Tính khối lượng Polyline."),
                    Helpers.RibbonFactory.CreateButton("CTSV_GhiDienTichPoly", "Ghi DT lên Poly", "Ghi diện tích lên Polyline."),
                    Helpers.RibbonFactory.CreateButton("CTSV_TinhDienTichMoi", "DT mới", "Tính diện tích mới."),
                }));
                p6.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Thống kê cọc", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTSV_ThongKeCoc", "Thống kê cọc", "Thống kê cọc xuất Excel."),
                    Helpers.RibbonFactory.CreateButton("CTSV_ThongKeCoc_ToaDo", "TK cọc tọa độ", "Thống kê cọc có tọa độ."),
                }));
                tab.Panels.Add(p6);

                // ══════════════════════════════════════════════════════════════
                // PANEL 7: CỐNG / HỐ GA (Pipe & Structure)
                // ══════════════════════════════════════════════════════════════
                var p7 = Helpers.RibbonFactory.CreatePanel("Cống hố ga");
                p7.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Pipe", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTPI_ThayDoi_DuongKinhCong", "ĐK cống", "Thay đổi đường kính cống."),
                    Helpers.RibbonFactory.CreateButton("CTPI_ThayDoi_MatPhangRef_Cong", "MP Reference", "Thay đổi mặt phẳng reference cống."),
                    Helpers.RibbonFactory.CreateButton("CTPI_ThayDoi_DoanDocCong", "Độ dốc cống", "Thay đổi độ dốc cống."),
                    Helpers.RibbonFactory.CreateButton("CTPI_BangCaoDo_TuNhienHoThu", "Bảng CĐ hố thu", "Bảng cao độ tự nhiên hố thu."),
                    Helpers.RibbonFactory.CreateButton("CTPI_XoayHoThu_Theo2diem", "Xoay hố thu", "Xoay hố thu theo 2 điểm."),
                }));
                p7.Source.Items.Add(Helpers.RibbonFactory.CreateButton("CT_VTOADOHG", "Tọa độ hố ga", "Trích xuất tọa độ hố ga NXsoft."));
                tab.Panels.Add(p7);

                // ══════════════════════════════════════════════════════════════
                // PANEL 8: POINT / COGO POINT
                // ══════════════════════════════════════════════════════════════
                var p8 = Helpers.RibbonFactory.CreatePanel("Point");
                p8.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("CogoPoint", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTPO_TaoCogoPoint_CaoDo_FromSurface", "CĐ từ Surface", "Tạo CogoPoint cao độ từ Surface."),
                    Helpers.RibbonFactory.CreateButton("CTPO_TaoCogoPoint_CaoDo_Elevationspot", "Từ Elev Spot", "Tạo CogoPoint từ Elevation Spot."),
                    Helpers.RibbonFactory.CreateButton("CTPO_CreateCogopointFromText", "Từ Text", "Tạo CogoPoint từ Text."),
                    Helpers.RibbonFactory.CreateButton("CTPO_UpdateAllPointGroup", "Update Groups", "Cập nhật tất cả Point Group."),
                    Helpers.RibbonFactory.CreateButton("CTPO_An_CogoPoint", "Ẩn Point", "Ẩn CogoPoint."),
                    Helpers.RibbonFactory.CreateButton("CTPO_CPTT", "Point thành Text", "Chuyển Point thành Text cao độ."),
                    Helpers.RibbonFactory.CreateButton("CTPo_ReorderPoints", "Sắp xếp Point", "Sắp xếp lại thứ tự Point."),
                    Helpers.RibbonFactory.CreateButton("CTU_MakePointFromText", "Text thành Point", "Tạo Point từ Text."),
                }));
                tab.Panels.Add(p8);

                // ══════════════════════════════════════════════════════════════
                // PANEL 9: SAN NỀN
                // ══════════════════════════════════════════════════════════════
                var p9 = Helpers.RibbonFactory.CreatePanel("San nền");
                p9.Source.Items.Add(Helpers.RibbonFactory.CreateButton("CTSN_Taskbar", "Taskbar SN", "Mở thanh Taskbar san nền."));
                p9.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Công cụ SN", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTSN_TaoLuoi", "Tạo lưới", "Tạo lưới san nền."),
                    Helpers.RibbonFactory.CreateButton("CTSN_NhapCaoDo", "Nhập cao độ", "Nhập cao độ san nền."),
                    Helpers.RibbonFactory.CreateButton("CTSN_TinhKL", "Tính KL", "Tính khối lượng san nền."),
                    Helpers.RibbonFactory.CreateButton("CTSN_XuatBang", "Xuất bảng", "Xuất bảng san nền."),
                    Helpers.RibbonFactory.CreateButton("CTSN_Surface", "Surface SN", "Tạo Surface san nền."),
                }));
                tab.Panels.Add(p9);

                // ══════════════════════════════════════════════════════════════
                // PANEL 10: TIỆN ÍCH CAD
                // ══════════════════════════════════════════════════════════════
                var p10 = Helpers.RibbonFactory.CreatePanel("Tiện ích CAD");
                p10.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Đo đạc", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("AT_TongDoDai_Full", "Tổng chiều dài", "Tính tổng chiều dài."),
                    Helpers.RibbonFactory.CreateButton("AT_TongDienTich_Full", "Tổng diện tích", "Tính tổng diện tích."),
                    Helpers.RibbonFactory.CreateButton("AT_TongDienTich_Replace_CongThem", "DT cộng thêm", "Cộng thêm DT vào Text."),
                    Helpers.RibbonFactory.CreateButton("CTU_TotalLength", "Tổng dài nhanh", "Tổng chiều dài nhanh."),
                }));
                p10.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Text / MText", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("AT_MT2ML", "MText thành MLeader", "Chuyển MText sang MLeader."),
                    Helpers.RibbonFactory.CreateButton("CTU_TextToMText", "Text thành MText", "Gộp Text thành MText."),
                    Helpers.RibbonFactory.CreateButton("AT_TextLink", "Text Link", "Liên kết Text."),
                    Helpers.RibbonFactory.CreateButton("AT_DanhSoThuTu", "Đánh số TT", "Đánh số thứ tự."),
                    Helpers.RibbonFactory.CreateButton("AT_Label_FromText", "Label từ Text", "Tạo Label từ Text."),
                    Helpers.RibbonFactory.CreateButton("CTU_ExportTextCoords", "Xuất tọa độ Text", "Xuất tọa độ từ Text."),
                }));
                p10.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Layout / Dim", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("AT_TextLayout", "Text Layout", "Text Layout."),
                    Helpers.RibbonFactory.CreateButton("AT_TaoMoi_TextLayout", "Tạo Text Layout", "Tạo mới Text Layout."),
                    Helpers.RibbonFactory.CreateButton("AT_DimLayout", "Dim Layout", "Dim Layout."),
                    Helpers.RibbonFactory.CreateButton("AT_DimLayout2", "Dim Layout 2", "Dim Layout 2."),
                    Helpers.RibbonFactory.CreateButton("AT_BlockLayout", "Block Layout", "Block Layout."),
                    Helpers.RibbonFactory.CreateButton("AT_UpdateLayout", "Update Layout", "Cập nhật Layout."),
                }));
                p10.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Đối tượng", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("AT_XoayDoiTuong_TheoViewport", "Xoay theo VP", "Xoay đối tượng theo Viewport."),
                    Helpers.RibbonFactory.CreateButton("AT_XoayDoiTuong_Theo2Diem", "Xoay theo 2 điểm", "Xoay đối tượng theo 2 điểm."),
                    Helpers.RibbonFactory.CreateButton("AT_Offset_2Ben", "Offset 2 bên", "Offset 2 bên."),
                    Helpers.RibbonFactory.CreateButton("AT_XoaDoiTuong_CungLayer", "Xóa cùng Layer", "Xóa đối tượng cùng Layer."),
                    Helpers.RibbonFactory.CreateButton("AT_annotive_scale_currentOnly", "Scale Annotative", "Annotative scale current only."),
                    Helpers.RibbonFactory.CreateButton("CTU_AddPolylineVertices", "Thêm đỉnh PL", "Thêm đỉnh Polyline."),
                    Helpers.RibbonFactory.CreateButton("CTU_FindIntersections", "Tìm giao điểm", "Tìm giao điểm."),
                    Helpers.RibbonFactory.CreateButton("CTU_DrawTaluy", "Vẽ taluy", "Vẽ taluy."),
                    Helpers.RibbonFactory.CreateButton("AT_TCD", "Trim chuẩn", "Trim dimension chuẩn."),
                    Helpers.RibbonFactory.CreateButton("AT_TBD", "Extend chuẩn", "Extend dimension chuẩn."),
                }));
                p10.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Xuất dữ liệu", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("XUATBANG_ToaDoPolyline", "Tọa độ Polyline", "Xuất bảng tọa độ Polyline."),
                    Helpers.RibbonFactory.CreateButton("CTU_ExportCAD2007", "Export CAD 2007", "Xuất file CAD 2007."),
                }));
                tab.Panels.Add(p10);

                // ══════════════════════════════════════════════════════════════
                // PANEL 11: LAYER
                // ══════════════════════════════════════════════════════════════
                var p11 = Helpers.RibbonFactory.CreatePanel("Layer");
                p11.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Bật/Tắt Layer", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTL_OnCorridor", "Bật Corridor", "Bật layer Corridor."),
                    Helpers.RibbonFactory.CreateButton("CTL_OffCorridor", "Tắt Corridor", "Tắt layer Corridor."),
                    Helpers.RibbonFactory.CreateButton("CTL_OnSampleLine", "Bật SampleLine", "Bật layer SampleLine."),
                    Helpers.RibbonFactory.CreateButton("CTL_OffSampleLine", "Tắt SampleLine", "Tắt layer SampleLine."),
                    Helpers.RibbonFactory.CreateButton("CTL_OnAlignment", "Bật Alignment", "Bật layer Alignment."),
                    Helpers.RibbonFactory.CreateButton("CTL_OffAlignment", "Tắt Alignment", "Tắt layer Alignment."),
                    Helpers.RibbonFactory.CreateButton("CTL_OnParcel", "Bật Parcel", "Bật layer Parcel."),
                    Helpers.RibbonFactory.CreateButton("CTL_OffParcel", "Tắt Parcel", "Tắt layer Parcel."),
                    Helpers.RibbonFactory.CreateButton("CTL_OnHatchDaoDap", "Bật Hatch ĐĐ", "Bật layer Hatch đào đắp."),
                    Helpers.RibbonFactory.CreateButton("CTL_OffHatchDaoDap", "Tắt Hatch ĐĐ", "Tắt layer Hatch đào đắp."),
                    Helpers.RibbonFactory.CreateButton("CTL_OnDefpoints", "Bật Defpoints", "Bật layer Defpoints."),
                    Helpers.RibbonFactory.CreateButton("CTL_OffDefpoints", "Tắt Defpoints", "Tắt layer Defpoints."),
                }));
                p11.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Chuyển Layer", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTL_ToText", "Sang Text", "Chuyển về Layer Text."),
                    Helpers.RibbonFactory.CreateButton("CTL_ToDefpoints", "Sang Defpoints", "Chuyển về Defpoints."),
                    Helpers.RibbonFactory.CreateButton("CTL_ToDim", "Sang Dim", "Chuyển về Layer Dim."),
                    Helpers.RibbonFactory.CreateButton("CTL_ToBaoBT", "Sang Bao BT", "Chuyển về Layer Bao BT."),
                    Helpers.RibbonFactory.CreateButton("CTL_ToBaoCotThep", "Sang Cốt thép", "Chuyển về Layer Cốt thép."),
                    Helpers.RibbonFactory.CreateButton("CTL_ToThep", "Sang Thép", "Chuyển về Layer Thép."),
                    Helpers.RibbonFactory.CreateButton("CTL_ToTruc", "Sang Trục", "Chuyển về Layer Trục."),
                    Helpers.RibbonFactory.CreateButton("CTL_ToKhuat", "Sang Khuất", "Chuyển về Layer Khuất."),
                    Helpers.RibbonFactory.CreateButton("CTL_ToHatch", "Sang Hatch", "Chuyển về Layer Hatch."),
                    Helpers.RibbonFactory.CreateButton("CTL_ToRanhGioi", "Sang Ranh giới", "Chuyển về Layer Ranh giới."),
                }));
                tab.Panels.Add(p11);

                // ══════════════════════════════════════════════════════════════
                // PANEL 12: THIẾT LẬP & HỆ THỐNG
                // ══════════════════════════════════════════════════════════════
                var p12 = Helpers.RibbonFactory.CreatePanel("Thiết lập");
                p12.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Bản vẽ", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTDS_ThietLap", "Thông số chuẩn", "Thiết lập biến hệ thống chuẩn AutoCAD."),
                    Helpers.RibbonFactory.CreateButton("CTDS_SaveClean", "Lưu và Purge", "Dọn dẹp bản vẽ và lưu file."),
                    Helpers.RibbonFactory.CreateButton("CT_DoiIcon", "Cá nhân hóa", "Thay đổi biểu tượng trên Ribbon."),
                    Helpers.RibbonFactory.CreateButton("CTDS_ConvertMM2M", "MM sang Mét", "Scale từ mm sang mét."),
                    Helpers.RibbonFactory.CreateButton("CTDS_ConvertCM2M", "CM sang Mét", "Scale từ cm sang mét."),
                }));
                p12.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("In ấn", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTDS_PrintAllLayouts", "In tất cả", "In toàn bộ Layouts."),
                    Helpers.RibbonFactory.CreateButton("CTDS_PrintCurrentLayout", "In Layout hiện hành", "In layout đang mở."),
                    Helpers.RibbonFactory.CreateButton("CTDS_ExportPDF", "Xuất PDF", "Xuất bản vẽ ra PDF."),
                }));
                p12.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Tiện ích C3D", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CTU_ExplodeAEC", "Explode AEC", "Phá các đối tượng AEC."),
                    Helpers.RibbonFactory.CreateButton("CTU_StyleAutoOff", "Style Auto OFF", "Tắt Style tự động."),
                    Helpers.RibbonFactory.CreateButton("CTU_StyleAutoOn", "Style Auto ON", "Bật Style tự động."),
                    Helpers.RibbonFactory.CreateButton("CTU_DumpObject", "Dump Object", "Xem thông tin chi tiết đối tượng."),
                    Helpers.RibbonFactory.CreateButton("CTP_AddParcelLabels", "Label Parcel", "Thêm label Parcel."),
                }));
                p12.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Property Set", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("AT_Solid_Set_PropertySet", "Set Property", "Thiết lập Property Set cho 3D Solid."),
                    Helpers.RibbonFactory.CreateButton("AT_Solid_Show_Info", "Info Solid", "Hiển thị thông tin 3D Solid."),
                }));
                p12.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("NX Power", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("NXFIXTDTN", "Sửa trắc dọc", "Sửa đường tự nhiên theo cọc."),
                    Helpers.RibbonFactory.CreateButton("NXPIPE", "Mạng lưới ống", "Thiết kế mạng lưới thoát nước."),
                    Helpers.RibbonFactory.CreateButton("NXNTDADD", "Nạp NTD", "Nạp dữ liệu NTD vào Alignment."),
                    Helpers.RibbonFactory.CreateButton("NXDCDCOC", "CĐ trắc dọc MB", "Phun cao độ thiết kế từ Profile lên bình đồ."),
                    Helpers.RibbonFactory.CreateButton("NXrenameSL", "Đổi tên cọc NX", "Đổi tên Sample Lines hàng loạt."),
                    Helpers.RibbonFactory.CreateButton("NXDTCoc", "Điền tên cọc NX", "Ghi chú tên cọc lên mặt bằng."),
                    Helpers.RibbonFactory.CreateButton("NXCCTTD", "Chèn cọc Profile", "Chèn cọc từ khung nhìn trắc dọc."),
                    Helpers.RibbonFactory.CreateButton("NXCCTN", "Chênh cao TN", "Điền chênh cao tim cọc."),
                    Helpers.RibbonFactory.CreateButton("CWPL", "Bề dày Polyline", "Thay đổi bề dày Polylines hàng loạt."),
                    Helpers.RibbonFactory.CreateButton("NXChangeLW", "Độ dày nét Layer", "Thay đổi LineWeight cho Layer."),
                    Helpers.RibbonFactory.CreateButton("NXNoiText", "Gộp văn bản KS", "Gộp Text phần nguyên và thập phân."),
                }));
                p12.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Hệ thống", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("CT_DanhSachLenh", "Danh sách lệnh", "Hiển thị danh sách tất cả lệnh."),
                    Helpers.RibbonFactory.CreateButton("TASKBAR", "Taskbar chính", "Mở thanh Taskbar chính."),
                }));
                tab.Panels.Add(p12);

                // ══════════════════════════════════════════════════════════════
                // TAB 2: VISUALINFRA (Module bổ sung)
                // ══════════════════════════════════════════════════════════════
                RibbonTab tab2 = new() { Title = "VisualINFRA", Id = "MyFirstProject.VITab" };
                var oldVi = ribbon.Tabs.FirstOrDefault(t => t.Id == "MyFirstProject.VITab");
                if (oldVi != null) ribbon.Tabs.Remove(oldVi);
                ribbon.Tabs.Add(tab2);

                var vi1 = Helpers.RibbonFactory.CreatePanel("SampleLine");
                vi1.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("SL Tools", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("VI_SampleLineCoordinate", "Tọa độ SL", "Tọa độ Sample Line."),
                    Helpers.RibbonFactory.CreateButton("VI_ExportSampleLine", "Xuất SL", "Xuất Sample Line."),
                    Helpers.RibbonFactory.CreateButton("VI_ImportSampleLine", "Nhập SL", "Nhập Sample Line."),
                    Helpers.RibbonFactory.CreateButton("VI_RenameSampleLine", "Đổi tên SL", "Đổi tên Sample Line."),
                    Helpers.RibbonFactory.CreateButton("VI_FillingSampleLine", "Filling SL", "Filling Sample Line."),
                }));
                tab2.Panels.Add(vi1);

                var vi2 = Helpers.RibbonFactory.CreatePanel("Profile Corridor");
                vi2.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Profile Tools", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("VI_CreateProfileView", "Tạo ProfileView", "Tạo Profile View."),
                    Helpers.RibbonFactory.CreateButton("VI_CreateMultiSurfaceProfile", "Multi Surface", "Tạo Multi Surface Profile."),
                    Helpers.RibbonFactory.CreateButton("VI_CreateOffsetProfile", "Offset Profile", "Tạo Offset Profile."),
                    Helpers.RibbonFactory.CreateButton("VI_CreateSurfaceFromCorridor", "Surface từ Cor", "Tạo Surface từ Corridor."),
                    Helpers.RibbonFactory.CreateButton("VI_RebuildAllCorridor", "Rebuild Corridor", "Rebuild tất cả Corridor."),
                    Helpers.RibbonFactory.CreateButton("VI_CorridorInfo", "Info Corridor", "Thông tin Corridor."),
                }));
                tab2.Panels.Add(vi2);

                var vi3 = Helpers.RibbonFactory.CreatePanel("Khối lượng VI");
                vi3.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Volume Tools", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("VI_VolumeCivilRoad", "KL đường", "Khối lượng đường."),
                    Helpers.RibbonFactory.CreateButton("VI_QuickVolume", "KL nhanh", "Tính khối lượng nhanh."),
                    Helpers.RibbonFactory.CreateButton("VI_VolumeNetwork", "KL mạng lưới", "Khối lượng mạng lưới."),
                    Helpers.RibbonFactory.CreateButton("VI_ExportVolume", "Xuất KL", "Xuất khối lượng."),
                    Helpers.RibbonFactory.CreateButton("VI_CompareSurfaceVolume", "So sánh Surface", "So sánh khối lượng Surface."),
                }));
                tab2.Panels.Add(vi3);

                var vi4 = Helpers.RibbonFactory.CreatePanel("Pipe Network");
                vi4.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Pipe Tools", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("VI_ManholeCoordinate", "Tọa độ hố ga", "Tọa độ hố ga."),
                    Helpers.RibbonFactory.CreateButton("VI_PipeInfo", "Info Pipe", "Thông tin Pipe."),
                    Helpers.RibbonFactory.CreateButton("VI_ExportPipeNetwork", "Xuất Pipe", "Xuất Pipe Network."),
                    Helpers.RibbonFactory.CreateButton("VI_ChangePipeElevation", "CĐ Pipe", "Thay đổi cao độ Pipe."),
                    Helpers.RibbonFactory.CreateButton("VI_ChangeStructureElevation", "CĐ Structure", "Thay đổi cao độ Structure."),
                    Helpers.RibbonFactory.CreateButton("VI_PipeNetworkSummary", "Tổng hợp Pipe", "Tổng hợp Pipe Network."),
                }));
                tab2.Panels.Add(vi4);

                var vi5 = Helpers.RibbonFactory.CreatePanel("In ấn VI");
                vi5.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Print Tools", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("VI_CreateKhungTuyen", "Khung tuyến", "Tạo khung tuyến."),
                    Helpers.RibbonFactory.CreateButton("VI_CreateViewport", "Tạo Viewport", "Tạo Viewport."),
                    Helpers.RibbonFactory.CreateButton("VI_CreateKhungTracNgang", "Khung TN", "Tạo khung trắc ngang."),
                    Helpers.RibbonFactory.CreateButton("VI_FitKhungIn", "Fit khung in", "Fit khung in."),
                    Helpers.RibbonFactory.CreateButton("VI_ZoomToSection", "Zoom Section", "Zoom đến Section."),
                    Helpers.RibbonFactory.CreateButton("VI_CopyLayout", "Copy Layout", "Copy Layout."),
                    Helpers.RibbonFactory.CreateButton("VI_ListLayouts", "Danh sách Layout", "Danh sách Layout."),
                }));
                tab2.Panels.Add(vi5);

                var vi6 = Helpers.RibbonFactory.CreatePanel("Tiện ích VI");
                vi6.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Zoom", new List<RibbonButton> {
                    Helpers.RibbonFactory.CreateButton("VI_ZO", "Zoom Object", "Zoom vào đối tượng."),
                    Helpers.RibbonFactory.CreateButton("VI_ZOOMXY", "Zoom XY", "Zoom theo tọa độ XY."),
                }));
                tab2.Panels.Add(vi6);

                tab.IsActive = true;
                var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor;
                ed?.WriteMessage("\nĐã tạo tab 'Civil Tool' (12 panel) và 'VisualINFRA' (6 panel) với đầy đủ lệnh trên Ribbon.");
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