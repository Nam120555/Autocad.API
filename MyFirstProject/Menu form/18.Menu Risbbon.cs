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
using Helpers = Civil3DCsharp.Helpers;

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

                // ══════════════════════════════════════════════════════════════════════════
                // CẤU TRÚC MENU CHUYÊN NGHIỆP: QUY TRÌNH KỸ SƯ GIAO THÔNG
                // ══════════════════════════════════════════════════════════════════════════

                // 1. THIẾT LẬP (Setup)
                var setupPanel = Helpers.RibbonFactory.CreatePanel("Thiết lập");
                
                var drawingButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTDS_ThietLap", "Thông số chuẩn", "Thiết lập các biến hệ thống (LTSCALE, Units...) chuẩn AutoCAD."),
                    Helpers.RibbonFactory.CreateButton("CTDS_SaveClean", "Lưu & Purge", "Dọn dẹp bản vẽ và lưu file (Purge All & Save)."),
                    Helpers.RibbonFactory.CreateButton("CT_DoiIcon", "Cá nhân hóa", "Thay đổi màu sắc và biểu tượng cho từng công cụ.")
                };
                setupPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Bản vẽ", drawingButtons));

                var unitButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTDS_ConvertMM2M", "MM → Mét", "Scale toàn bộ bản vẽ từ mm sang mét (0.001)."),
                    Helpers.RibbonFactory.CreateButton("CTDS_ConvertCM2M", "CM → Mét", "Scale toàn bộ bản vẽ từ cm sang mét (0.01).")
                };
                setupPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Đơn vị", unitButtons));

                var printButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTDS_PrintAllLayouts", "In tất cả Layout", "Tự động in toàn bộ Layouts trong bản vẽ."),
                    Helpers.RibbonFactory.CreateButton("CTDS_PrintCurrentLayout", "In Layout hiện hành", "In nhanh layout đang mở."),
                    Helpers.RibbonFactory.CreateButton("CTDS_ExportPDF", "Xuất PDF", "Xuất bản vẽ ra PDF chất lượng cao.")
                };
                setupPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("In ấn", printButtons));
                tab.Panels.Add(setupPanel);

                // 2. DỮ LIỆU NỀN (Base Data)
                var dataPanel = Helpers.RibbonFactory.CreatePanel("Dữ liệu");
                
                var surfaceButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTS_TaoSpotElevation_OnSurface_TaiTim", "Spot Elev", "Tạo cao độ tại tim tuyến dựa trên Surface."),
                    Helpers.RibbonFactory.CreateButton("CTS_RebuildSurface", "Rebuild Surface", "Cập nhật lại bề mặt địa hình đã chọn.")
                };
                dataPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Địa hình", surfaceButtons));

                var pointButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTPo_TaoPointTheoBang", "Import Point", "Nhập Cogo Point từ danh sách bảng tính/Excel."),
                    Helpers.RibbonFactory.CreateButton("CTPo_ChuyenPointThanhBlock", "Point → Block", "Chuyển đổi Cogo Point sang Block để xuất sang CAD thường."),
                    Helpers.RibbonFactory.CreateButton("CTPo_TaoBangThongKePoint", "Thống kê Point", "Tạo bảng thống kê tọa độ, cao độ Point.")
                };
                dataPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Cogo Point", pointButtons));
                tab.Panels.Add(dataPanel);

                // 3. THIẾT KẾ TUYẾN (Plan Design)
                var planPanel = Helpers.RibbonFactory.CreatePanel("Thiết kế");
                
                var curveButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTC_ThietLapDuongCong", "Tra bảng TCVN", "Mở form tra cứu và áp dụng thông số đường cong theo TCVN."),
                    Helpers.RibbonFactory.CreateButton("CTC_TraCuuDuongCong", "Tra cứu nhanh", "Hiển thị thông số đường cong cơ bản theo vận tốc thiết kế."),
                    Helpers.RibbonFactory.CreateButton("CTC_ThongSoDuongCong_4054", "Bảng 4054:2005", "Xem bảng tra cứu TCVN 4054 cho đường ngoài đô thị."),
                    Helpers.RibbonFactory.CreateButton("CTC_ThongSoDuongCong_13592", "Bảng 13592:2022", "Xem bảng tra cứu TCVN 13592 cho đường đô thị.")
                };
                planPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Đường cong", curveButtons));

                var checkButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTC_KiemTraDuongCong_4054", "Check 4054", "Kiểm tra Alignment hiện có theo TCVN 4054:2005."),
                    Helpers.RibbonFactory.CreateButton("CTC_KiemTraDuongCong_13592", "Check 13592", "Kiểm tra Alignment hiện có theo TCVN 13592:2022.")
                };
                planPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Kiểm tra", checkButtons));
                tab.Panels.Add(planPanel);

                // 4. LÝ TRÌNH & CỌC (Staking)
                var stakingPanel = Helpers.RibbonFactory.CreatePanel("Cọc");
                
                var namingButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTS_DoiTenCoc", "Đổi tên cọc", "Tùy chỉnh tên cọc thủ công hoặc theo quy luật."),
                    Helpers.RibbonFactory.CreateButton("CTS_DoiTenCoc_TheoThuTu", "Đánh số thứ tự", "Đánh lại số thứ tự (1, 2, 3...) cho toàn bộ cọc.")
                };
                stakingPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Quản lý", namingButtons));

                var genStakesButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTS_PhatSinhCoc", "Phát sinh auto", "Tự động phát sinh cọc theo khoảng cách định trước."),
                    Helpers.RibbonFactory.CreateButton("CTS_CHENCOC_TRENTRACNGANG", "Chèn cọc TN", "Chèn thêm cọc tại các vị trí đặc biệt trên trắc ngang.")
                };
                stakingPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Phát sinh", genStakesButtons));
                tab.Panels.Add(stakingPanel);

                // 5. TRẮC DỌC & NGANG (Sections)
                var sectionPanel = Helpers.RibbonFactory.CreatePanel("Mặt cắt");
                
                var profileButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTPV_TaoProfileView", "Trắc dọc", "Tạo khung nhìn trắc dọc (Profile View) tự động."),
                    Helpers.RibbonFactory.CreateButton("CTPV_FitKhung", "Fit Khung", "Căn chỉnh trắc dọc vừa vặn với khổ giấy thiết kế.")
                };
                sectionPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Trắc dọc", profileButtons));

                var crossSectionButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTSv_VeTracNgangThietKe", "Vẽ trắc ngang", "Xuất trắc ngang thiết kế hàng loạt."),
                    Helpers.RibbonFactory.CreateButton("CTSV_DanhCap", "Đánh cấp VHC", "Tính toán và vẽ khối lượng đánh cấp (hữu cơ).")
                };
                sectionPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Trắc ngang", crossSectionButtons));
                tab.Panels.Add(sectionPanel);

                // 5.2 SAN NỀN (Grading)
                var sanNenPanel = Helpers.RibbonFactory.CreatePanel("San nền");
                sanNenPanel.Source.Items.Add(Helpers.RibbonFactory.CreateButton("CTSN_Taskbar", "▦ Thanh công cụ", "Mở thanh Taskbar chuyên dụng cho công tác san nền."));

                var gridButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTSN_TaoLuoi", "▦ Tạo lưới", "Tạo lưới ô vuông san nền, thiết lập kích thước và ranh giới."),
                    Helpers.RibbonFactory.CreateButton("CTSN_Surface", "🗺 Lấy cao độ", "Tự động lấy cao độ tự nhiên/thiết kế từ Surface cho lưới."),
                    Helpers.RibbonFactory.CreateButton("CTSN_NhapCaoDo", "✏ Nhập thủ công", "Nhập hoặc hiệu chỉnh cao độ tại các nút lưới thủ công.")
                };
                sanNenPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Lưới ô vuông", gridButtons));

                var calcButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTSN_TinhKL", "📊 Tính khối lượng", "Tính toán tổng hợp khối lượng đào đắp san nền."),
                    Helpers.RibbonFactory.CreateButton("CTSN_XuatBang", "📋 Xuất bảng CAD", "Vẽ bảng tổng hợp khối lượng san nền vào bản vẽ.")
                };
                sanNenPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Tính toán", calcButtons));
                tab.Panels.Add(sanNenPanel);

                // 6. KHỐI LƯỢNG (Reports)
                var reportPanel = Helpers.RibbonFactory.CreatePanel("Báo cáo");
                reportPanel.Source.Items.Add(Helpers.RibbonFactory.CreateButton("CTSV_Taskbar", "📊 Bảng KL Nhanh", "Mở thanh Taskbar quản lý khối lượng tập trung.", "xuat_kl_image"));
                
                var reportButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTSV_XuatKhoiLuong", "📥 Sang Excel", "Xuất bảng khối lượng chi tiết ra file Excel chuyên nghiệp."),
                    Helpers.RibbonFactory.CreateButton("CTSV_XuatCad", "📋 Vẽ bảng CAD", "Vẽ trực tiếp bảng tổng hợp khối lượng vào bản vẽ.")
                };
                reportPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Kết xuất", reportButtons));
                tab.Panels.Add(reportPanel);

                // 7. TIỆN ÍCH CAD (Utilities)
                var utilPanel = Helpers.RibbonFactory.CreatePanel("Tiện ích");
                
                var measureButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("AT_TongDoDai_Full", "📐 Tổng chiều dài", "Tính tổng chiều dài các đối tượng (Line, Poly, Spline...)."),
                    Helpers.RibbonFactory.CreateButton("AT_TongDienTich_Full", "📐 Tổng diện tích", "Tính tổng diện tích các vùng khép kín, Hatch.")
                };
                utilPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Đo đạc", measureButtons));

                var textButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("AT_MT2ML", "✏ MText → MLeader", "Chuyển nội dung MText sang MLeader có mũi tên."),
                    Helpers.RibbonFactory.CreateButton("CTU_TextToMText", "✏ Text → MText", "Gộp nhiều Text đơn lẻ thành MText.")
                };
                utilPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Văn bản", textButtons));

                var layerButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CTL_OnCorridor", "⚡ Layer Corridor", "Quản lý hiển thị các lớp thiết kế Corridor."),
                    Helpers.RibbonFactory.CreateButton("CTL_ToText", "🎨 Về Layer chuẩn", "Chuyển đối tượng về Layer quy chuẩn của công ty.")
                };
                utilPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Layer", layerButtons));
                tab.Panels.Add(utilPanel);

                // 8. NX POWER TOOLS (Bridge & Master Suite)
                var nxPanel = Helpers.RibbonFactory.CreatePanel("NX Power");
                
                var nxBridgeButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CT_VTOADOHG", "⚡ Tọa độ hố ga", "Công cụ trích xuất tọa độ hố ga chuyên sâu (Bridge NXsoft).")
                };
                nxPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Bridge", nxBridgeButtons));

                var nxCivilButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("NXFIXTDTN", "Sửa trắc dọc", "Sửa đường tự nhiên theo cọc hoặc khoảng cách đều."),
                    Helpers.RibbonFactory.CreateButton("NXPIPE", "Mạng lưới ống", "Thiết kế mạng lưới thoát nước từ đối tượng CAD."),
                    Helpers.RibbonFactory.CreateButton("NXNTDADD", "Nạp NTD", "Nạp dữ liệu từ file NTD vào Alignment hiện có."),
                    Helpers.RibbonFactory.CreateButton("NXDCDCOC", "CĐ trắc dọc → Mặt bằng", "Phun cao độ thiết kế từ Profile lên bình đồ.")
                };
                nxPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Thiết kế NX", nxCivilButtons));

                var nxStakingButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("NXrenameSL", "Đổi tên cọc (NX)", "Đổi tên Sample Lines hàng loạt theo quy chuẩn NX."),
                    Helpers.RibbonFactory.CreateButton("NXDTCoc", "Điền tên cọc (NX)", "Ghi chú tên cọc lên mặt bằng theo phong cách NX."),
                    Helpers.RibbonFactory.CreateButton("NXCCTTD", "Chèn cọc Profile", "Chèn cọc trực tiếp từ khung nhìn trắc dọc.")
                };
                nxPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Cọc & Lý trình", nxStakingButtons));

                var nxUtilsButtons = new List<RibbonButton>
                {
                    Helpers.RibbonFactory.CreateButton("CWPL", "Bề dày Polyline", "Thay đổi bề dày (Width) của Polylines hàng loạt."),
                    Helpers.RibbonFactory.CreateButton("NXChangeLW", "Độ dày nét Layer", "Thay đổi LineWeight cho Layer theo đối tượng chọn."),
                    Helpers.RibbonFactory.CreateButton("NXNoiText", "Gộp văn bản KS", "Gộp Text phần nguyên và thập phân (Số liệu khảo sát).")
                };
                nxPanel.Source.Items.Add(Helpers.RibbonFactory.CreateSplitButton("Tiện ích Master", nxUtilsButtons));

                tab.Panels.Add(nxPanel);

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