// 19.CivilRibbonTaskbar.cs - Thanh Ribbon Civil Tool với màu xanh da trời
// Thiết kế theo phong cách Civil 3D chuyên nghiệp

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;

using AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadException = Autodesk.AutoCAD.Runtime.Exception;

[assembly: CommandClass(typeof(MyFirstProject.CivilRibbonTaskbar))]

namespace MyFirstProject
{
    /// <summary>
    /// Civil Ribbon Taskbar với theme xanh da trời và icon Civil 3D style
    /// </summary>
    public class CivilRibbonTaskbar
    {
        // Màu xanh da trời chủ đạo
        private static readonly System.Windows.Media.Color PrimaryBlue = System.Windows.Media.Color.FromRgb(0, 149, 217);
        private static readonly System.Windows.Media.Color LightBlue = System.Windows.Media.Color.FromRgb(135, 206, 250);
        private static readonly System.Windows.Media.Color DarkBlue = System.Windows.Media.Color.FromRgb(0, 102, 153);
        private static readonly System.Windows.Media.Color AccentBlue = System.Windows.Media.Color.FromRgb(100, 181, 246);

        [CommandMethod("CIVIL_RIBBON")]
        public static void ShowCivilRibbon()
        {
            try
            {
                var ribbon = ComponentManager.Ribbon;
                if (ribbon == null)
                {
                    var doc = AcadApplication.DocumentManager.MdiActiveDocument;
                    doc?.SendStringToExecute("RIBBON ", true, false, false);
                    ribbon = ComponentManager.Ribbon;
                    if (ribbon == null)
                    {
                        doc?.Editor.WriteMessage("\n⚠ Không thể khởi tạo Ribbon. Vui lòng bật RIBBON và chạy lại.");
                        return;
                    }
                }

                // Xóa tab cũ nếu có
                RemoveExistingTabs(ribbon);

                // Tạo tab Civil Tool mới
                CreateCivilToolTab(ribbon);

                // Tạo tab Acad Tool
                CreateAcadToolTab(ribbon);

                var ed = AcadApplication.DocumentManager.MdiActiveDocument?.Editor;
                ed?.WriteMessage("\n✅ Đã tạo Civil Ribbon với theme xanh da trời thành công!");
            }
            catch (System.Exception ex)
            {
                var ed = AcadApplication.DocumentManager.MdiActiveDocument?.Editor;
                ed?.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        private static void RemoveExistingTabs(RibbonControl ribbon)
        {
            var tabIds = new[] { "CivilTool.MainTab", "CivilTool.AcadTab", "MyFirstProject.C3DTab", "MyFirstProject.AcadTab" };
            foreach (var id in tabIds)
            {
                var existing = ribbon.Tabs.FirstOrDefault(t => t.Id == id);
                if (existing != null)
                    ribbon.Tabs.Remove(existing);
            }
        }

        private static void CreateCivilToolTab(RibbonControl ribbon)
        {
            RibbonTab civilTab = new RibbonTab
            {
                Title = "🛠 CIVIL TOOL",
                Id = "CivilTool.MainTab"
            };
            ribbon.Tabs.Add(civilTab);

            // Panel 1: Bề mặt & Điểm (Surface & Points)
            AddCivilPanel(civilTab, "🗺 Bề Mặt", "SURFACE", GetSurfaceCommands());

            // Panel 2: Lưới cọc (Sample Lines)
            AddCivilPanel(civilTab, "📍 Lưới Cọc", "STATION", GetSampleLineCommands());

            // Panel 3: Tuyến & Trắc dọc (Alignment & Profile)
            AddCivilPanel(civilTab, "🛣 Tuyến", "ALIGN", GetAlignmentCommands());

            // Panel 4: Corridor
            AddCivilPanel(civilTab, "🛤 Corridor", "CORRIDOR", GetCorridorCommands());

            // Panel 5: Trắc ngang (Section View)
            AddCivilPanel(civilTab, "📐 Trắc Ngang", "SECTION", GetSectionViewCommands());

            // Panel 6: San nền (Grading)  
            AddCivilPanel(civilTab, "▦ San Nền", "GRADING", GetGradingCommands());

            // Panel 7: Cống & Hố ga (Pipe & Structure)
            AddCivilPanel(civilTab, "🔧 Thoát Nước", "PIPE", GetPipeCommands());

            // Panel 8: Point
            AddCivilPanel(civilTab, "📌 Điểm", "POINT", GetPointCommands());

            // Panel 9: Tiện ích
            AddCivilPanel(civilTab, "⚙ Tiện Ích", "UTILITY", GetUtilityCommands());

            // Panel 10: Tài khoản
            AddCivilPanel(civilTab, "👤 Tài Khoản", "ACCOUNT", GetAccountCommands());

            civilTab.IsActive = true;
        }

        private static void CreateAcadToolTab(RibbonControl ribbon)
        {
            RibbonTab acadTab = new RibbonTab
            {
                Title = "📏 ACAD TOOL",
                Id = "CivilTool.AcadTab"
            };
            ribbon.Tabs.Add(acadTab);

            // Panel: CAD Commands
            AddCivilPanel(acadTab, "📏 Đo Lường", "MEASURE", GetMeasureCommands());
            AddCivilPanel(acadTab, "✏ Chỉnh Sửa", "EDIT", GetEditCommands());
            AddCivilPanel(acadTab, "🔄 Layout", "LAYOUT", GetLayoutCommands());
        }

        private static void AddCivilPanel(RibbonTab tab, string title, string iconType, 
            List<(string Command, string Label, string SubIcon)> commands)
        {
            if (commands.Count == 0) return;

            RibbonPanelSource panelSource = new RibbonPanelSource { Title = title };
            RibbonPanel panel = new RibbonPanel { Source = panelSource };

            // Tạo Split Button với dropdown
            RibbonSplitButton splitButton = new RibbonSplitButton
            {
                Text = title,
                ShowText = true,
                ShowImage = true,
                Size = RibbonItemSize.Large,
                Orientation = System.Windows.Controls.Orientation.Vertical,
                Image = CreateCivilIcon(iconType, 16),
                LargeImage = CreateCivilIcon(iconType, 32),
                ListStyle = RibbonSplitButtonListStyle.List
            };

            foreach (var (command, label, subIcon) in commands)
            {
                if (command == "---")
                {
                    splitButton.Items.Add(new RibbonSeparator());
                    continue;
                }

                RibbonButton btn = new RibbonButton
                {
                    Text = label,
                    ShowText = true,
                    ShowImage = true,
                    Image = CreateCivilIcon(subIcon, 16),
                    LargeImage = CreateCivilIcon(subIcon, 32),
                    Size = RibbonItemSize.Standard,
                    CommandHandler = new CivilRibbonCommandHandler(),
                    Tag = command,
                    CommandParameter = command,
                    ToolTip = new RibbonToolTip
                    {
                        Title = label,
                        Content = $"Lệnh: {command}",
                        Command = command
                    }
                };
                splitButton.Items.Add(btn);
            }

            panelSource.Items.Add(splitButton);
            tab.Panels.Add(panel);
        }

        #region Icon Generator - Civil Engineering Theme

        private static ImageSource CreateCivilIcon(string iconType, int size)
        {
            // Tạo icon theo chuẩn ký hiệu kỹ thuật giao thông
            DrawingVisual visual = new DrawingVisual();
            using (DrawingContext dc = visual.RenderOpen())
            {
                var brushBlue = new SolidColorBrush(PrimaryBlue);
                var brushLight = new SolidColorBrush(LightBlue);
                var brushDark = new SolidColorBrush(DarkBlue);
                var brushAccent = new SolidColorBrush(AccentBlue);
                var brushWhite = System.Windows.Media.Brushes.White;
                var brushRed = new SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 53, 69));
                var brushGreen = new SolidColorBrush(System.Windows.Media.Color.FromRgb(40, 167, 69));
                var brushBrown = new SolidColorBrush(System.Windows.Media.Color.FromRgb(139, 90, 43));
                
                var penBlue = new System.Windows.Media.Pen(brushBlue, size / 10.0);
                var penDark = new System.Windows.Media.Pen(brushDark, size / 8.0);
                var penWhite = new System.Windows.Media.Pen(brushWhite, size / 12.0);
                var penRed = new System.Windows.Media.Pen(brushRed, size / 10.0);
                var penGreen = new System.Windows.Media.Pen(brushGreen, size / 10.0);

                double s = size;
                double m = s * 0.08; // margin

                switch (iconType.ToUpper())
                {
                    case "SURFACE":
                        // TIN Surface - Lưới tam giác địa hình
                        dc.DrawLine(penBlue, new System.Windows.Point(m, s * 0.8), new System.Windows.Point(s * 0.35, s * 0.3));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.35, s * 0.3), new System.Windows.Point(s * 0.65, s * 0.5));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.65, s * 0.5), new System.Windows.Point(s - m, s * 0.2));
                        dc.DrawLine(penBlue, new System.Windows.Point(m, s * 0.8), new System.Windows.Point(s * 0.5, s * 0.7));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.5, s * 0.7), new System.Windows.Point(s - m, s * 0.8));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.35, s * 0.3), new System.Windows.Point(s * 0.5, s * 0.7));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.65, s * 0.5), new System.Windows.Point(s * 0.5, s * 0.7));
                        // Điểm cao độ
                        dc.DrawEllipse(brushDark, null, new System.Windows.Point(s * 0.35, s * 0.3), s * 0.05, s * 0.05);
                        dc.DrawEllipse(brushDark, null, new System.Windows.Point(s * 0.65, s * 0.5), s * 0.05, s * 0.05);
                        break;

                    case "STATION":
                        // Lý trình - Cọc Km trên tuyến đường
                        // Tim tuyến
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 6.0), 
                            new System.Windows.Point(m, s / 2), new System.Windows.Point(s - m, s / 2));
                        // Cọc lý trình (ký hiệu gạch vuông góc)
                        for (double x = s * 0.2; x <= s * 0.8; x += s * 0.2)
                        {
                            dc.DrawLine(penDark, new System.Windows.Point(x, s * 0.35), new System.Windows.Point(x, s * 0.65));
                        }
                        // Cọc Km chính (lớn hơn)
                        dc.DrawLine(new System.Windows.Media.Pen(brushDark, size / 6.0), 
                            new System.Windows.Point(s * 0.5, s * 0.25), new System.Windows.Point(s * 0.5, s * 0.75));
                        // Mũi tên hướng tuyến
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.85, s * 0.4), new System.Windows.Point(s - m, s / 2));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.85, s * 0.6), new System.Windows.Point(s - m, s / 2));
                        break;

                    case "ALIGN":
                        // Đường cong nằm - Horizontal Curve với PI, PC, PT
                        var alignCurve = new StreamGeometry();
                        using (var ctx = alignCurve.Open())
                        {
                            ctx.BeginFigure(new System.Windows.Point(m, s * 0.7), false, false);
                            ctx.QuadraticBezierTo(
                                new System.Windows.Point(s * 0.5, s * 0.15), // PI point
                                new System.Windows.Point(s - m, s * 0.7),
                                true, true);
                        }
                        dc.DrawGeometry(null, new System.Windows.Media.Pen(brushBlue, size / 5.0), alignCurve);
                        // PI - Điểm giao (tam giác)
                        dc.DrawEllipse(brushRed, null, new System.Windows.Point(s * 0.5, s * 0.25), s * 0.08, s * 0.08);
                        // PC, PT - Điểm tiếp đầu/cuối (vuông)
                        dc.DrawRectangle(brushGreen, null, new System.Windows.Rect(s * 0.12, s * 0.65, s * 0.1, s * 0.1));
                        dc.DrawRectangle(brushGreen, null, new System.Windows.Rect(s * 0.78, s * 0.65, s * 0.1, s * 0.1));
                        break;

                    case "CORRIDOR":
                        // Mặt cắt ngang đường - Road Cross Section
                        // Nền đường (hình thang)
                        var roadGeom = new StreamGeometry();
                        using (var ctx = roadGeom.Open())
                        {
                            ctx.BeginFigure(new System.Windows.Point(m, s * 0.7), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.2, s * 0.45), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.8, s * 0.45), true, true);
                            ctx.LineTo(new System.Windows.Point(s - m, s * 0.7), true, true);
                        }
                        dc.DrawGeometry(brushLight, penDark, roadGeom);
                        // Tim đường
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 8.0), 
                            new System.Windows.Point(s * 0.5, s * 0.45), new System.Windows.Point(s * 0.5, s * 0.7));
                        // Làn xe (vạch kẻ)
                        dc.DrawLine(penWhite, new System.Windows.Point(s * 0.35, s * 0.5), new System.Windows.Point(s * 0.35, s * 0.55));
                        dc.DrawLine(penWhite, new System.Windows.Point(s * 0.65, s * 0.5), new System.Windows.Point(s * 0.65, s * 0.55));
                        break;

                    case "SECTION":
                        // Trắc ngang - Cross Section với đào/đắp
                        // Địa hình tự nhiên (đường gấp khúc màu nâu)
                        dc.DrawLine(new System.Windows.Media.Pen(brushBrown, size / 10.0), 
                            new System.Windows.Point(m, s * 0.6), new System.Windows.Point(s * 0.3, s * 0.4));
                        dc.DrawLine(new System.Windows.Media.Pen(brushBrown, size / 10.0), 
                            new System.Windows.Point(s * 0.3, s * 0.4), new System.Windows.Point(s * 0.5, s * 0.55));
                        dc.DrawLine(new System.Windows.Media.Pen(brushBrown, size / 10.0), 
                            new System.Windows.Point(s * 0.5, s * 0.55), new System.Windows.Point(s * 0.7, s * 0.35));
                        dc.DrawLine(new System.Windows.Media.Pen(brushBrown, size / 10.0), 
                            new System.Windows.Point(s * 0.7, s * 0.35), new System.Windows.Point(s - m, s * 0.5));
                        // Đường thiết kế (đường thẳng màu xanh)
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 6.0), 
                            new System.Windows.Point(s * 0.15, s * 0.5), new System.Windows.Point(s * 0.85, s * 0.5));
                        // Mái taluy (gạch chéo)
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.15, s * 0.5), new System.Windows.Point(m, s * 0.7));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.85, s * 0.5), new System.Windows.Point(s - m, s * 0.7));
                        break;

                    case "GRADING":
                        // San nền - Lưới ô vuông với cao độ
                        double g = (s - 2 * m) / 4;
                        // Vẽ lưới
                        for (int i = 0; i <= 4; i++)
                        {
                            dc.DrawLine(new System.Windows.Media.Pen(brushBlue, 1), 
                                new System.Windows.Point(m, m + i * g), new System.Windows.Point(s - m, m + i * g));
                            dc.DrawLine(new System.Windows.Media.Pen(brushBlue, 1), 
                                new System.Windows.Point(m + i * g, m), new System.Windows.Point(m + i * g, s - m));
                        }
                        // Điểm cao độ góc lưới
                        dc.DrawEllipse(brushDark, null, new System.Windows.Point(m + g, m + g), s * 0.04, s * 0.04);
                        dc.DrawEllipse(brushDark, null, new System.Windows.Point(m + 2*g, m + 2*g), s * 0.04, s * 0.04);
                        dc.DrawEllipse(brushDark, null, new System.Windows.Point(m + 3*g, m + g), s * 0.04, s * 0.04);
                        break;

                    case "PIPE":
                        // Cống thoát nước - Pipe với hố ga
                        // Ống cống (hình chữ nhật nghiêng)
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 5.0), 
                            new System.Windows.Point(m, s * 0.6), new System.Windows.Point(s * 0.4, s * 0.5));
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 5.0), 
                            new System.Windows.Point(s * 0.6, s * 0.5), new System.Windows.Point(s - m, s * 0.4));
                        // Hố ga (hình vuông)
                        dc.DrawRectangle(brushDark, penBlue, new System.Windows.Rect(s * 0.4, s * 0.35, s * 0.2, s * 0.3));
                        // Mũi tên hướng chảy
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.75, s * 0.35), new System.Windows.Point(s * 0.85, s * 0.4));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.75, s * 0.45), new System.Windows.Point(s * 0.85, s * 0.4));
                        break;

                    case "POINT":
                        // Điểm đo đạc - Survey Benchmark Symbol
                        // Tam giác đo đạc
                        var triangleGeom = new StreamGeometry();
                        using (var ctx = triangleGeom.Open())
                        {
                            ctx.BeginFigure(new System.Windows.Point(s / 2, s * 0.15), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.2, s * 0.75), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.8, s * 0.75), true, true);
                        }
                        dc.DrawGeometry(null, new System.Windows.Media.Pen(brushBlue, size / 8.0), triangleGeom);
                        // Tâm điểm
                        dc.DrawEllipse(brushDark, null, new System.Windows.Point(s / 2, s * 0.55), s * 0.08, s * 0.08);
                        break;

                    case "UTILITY":
                        // Tiện ích - Cài đặt (bánh răng)
                        dc.DrawEllipse(brushBlue, null, new System.Windows.Point(s / 2, s / 2), s * 0.28, s * 0.28);
                        dc.DrawEllipse(brushWhite, null, new System.Windows.Point(s / 2, s / 2), s * 0.12, s * 0.12);
                        for (int i = 0; i < 6; i++)
                        {
                            double angle = i * Math.PI / 3;
                            double x1 = s / 2 + s * 0.28 * Math.Cos(angle);
                            double y1 = s / 2 + s * 0.28 * Math.Sin(angle);
                            double x2 = s / 2 + s * 0.4 * Math.Cos(angle);
                            double y2 = s / 2 + s * 0.4 * Math.Sin(angle);
                            dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 5.0), 
                                new System.Windows.Point(x1, y1), new System.Windows.Point(x2, y2));
                        }
                        break;

                    case "ACCOUNT":
                        // Tài khoản - User icon
                        dc.DrawEllipse(brushBlue, null, new System.Windows.Point(s / 2, s * 0.32), s * 0.2, s * 0.2);
                        var bodyGeom = new StreamGeometry();
                        using (var ctx = bodyGeom.Open())
                        {
                            ctx.BeginFigure(new System.Windows.Point(s * 0.2, s * 0.9), true, true);
                            ctx.QuadraticBezierTo(
                                new System.Windows.Point(s / 2, s * 0.55),
                                new System.Windows.Point(s * 0.8, s * 0.9), true, true);
                        }
                        dc.DrawGeometry(brushBlue, null, bodyGeom);
                        break;

                    case "MEASURE":
                        // Đo lường - Thước đo
                        dc.DrawRectangle(brushBlue, null, new System.Windows.Rect(m, s * 0.35, s - 2 * m, s * 0.3));
                        for (int i = 0; i <= 10; i++)
                        {
                            double x = m + (s - 2 * m) * i / 10;
                            double h = i % 5 == 0 ? s * 0.2 : (i % 2 == 0 ? s * 0.12 : s * 0.08);
                            dc.DrawLine(penWhite, new System.Windows.Point(x, s * 0.35), new System.Windows.Point(x, s * 0.35 + h));
                        }
                        break;

                    case "EDIT":
                        // Chỉnh sửa - Bút vẽ
                        var pencilGeom = new StreamGeometry();
                        using (var ctx = pencilGeom.Open())
                        {
                            ctx.BeginFigure(new System.Windows.Point(s * 0.15, s * 0.85), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.1, s * 0.78), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.72, s * 0.16), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.84, s * 0.22), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.22, s * 0.84), true, true);
                        }
                        dc.DrawGeometry(brushBlue, penDark, pencilGeom);
                        break;

                    case "LAYOUT":
                        // Layout - Bản vẽ khung
                        dc.DrawRectangle(null, penBlue, new System.Windows.Rect(m, m, s - 2 * m, s - 2 * m));
                        dc.DrawRectangle(brushLight, null, new System.Windows.Rect(s * 0.15, s * 0.15, s * 0.5, s * 0.35));
                        dc.DrawRectangle(brushBlue, null, new System.Windows.Rect(s * 0.15, s * 0.55, s * 0.7, s * 0.3));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.7, s * 0.15), new System.Windows.Point(s * 0.7, s * 0.5));
                        break;

                    // ===== SUB ICONS cho menu items =====
                    case "ADD":
                        dc.DrawEllipse(brushBlue, null, new System.Windows.Point(s / 2, s / 2), s * 0.38, s * 0.38);
                        dc.DrawLine(new System.Windows.Media.Pen(brushWhite, size / 6.0), 
                            new System.Windows.Point(s * 0.25, s / 2), new System.Windows.Point(s * 0.75, s / 2));
                        dc.DrawLine(new System.Windows.Media.Pen(brushWhite, size / 6.0), 
                            new System.Windows.Point(s / 2, s * 0.25), new System.Windows.Point(s / 2, s * 0.75));
                        break;

                    case "RENAME":
                        dc.DrawRectangle(brushLight, penBlue, new System.Windows.Rect(m, s * 0.3, s - 2 * m, s * 0.4));
                        dc.DrawLine(new System.Windows.Media.Pen(brushDark, 2), 
                            new System.Windows.Point(s * 0.2, s / 2), new System.Windows.Point(s * 0.6, s / 2));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.7, s * 0.4), new System.Windows.Point(s * 0.7, s * 0.6));
                        break;

                    case "TABLE":
                        dc.DrawRectangle(null, penBlue, new System.Windows.Rect(m, m, s - 2 * m, s - 2 * m));
                        dc.DrawLine(penBlue, new System.Windows.Point(m, s * 0.35), new System.Windows.Point(s - m, s * 0.35));
                        dc.DrawLine(penBlue, new System.Windows.Point(m, s * 0.55), new System.Windows.Point(s - m, s * 0.55));
                        dc.DrawLine(penBlue, new System.Windows.Point(m, s * 0.75), new System.Windows.Point(s - m, s * 0.75));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.4, m), new System.Windows.Point(s * 0.4, s - m));
                        break;

                    case "EXPORT":
                        // Mũi tên xuất file
                        dc.DrawRectangle(brushLight, penBlue, new System.Windows.Rect(s * 0.25, m, s * 0.5, s * 0.4));
                        var arrowGeom = new StreamGeometry();
                        using (var ctx = arrowGeom.Open())
                        {
                            ctx.BeginFigure(new System.Windows.Point(s / 2, s * 0.45), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.3, s * 0.65), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.42, s * 0.65), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.42, s * 0.9), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.58, s * 0.9), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.58, s * 0.65), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.7, s * 0.65), true, true);
                        }
                        dc.DrawGeometry(brushBlue, null, arrowGeom);
                        break;

                    case "SYNC":
                        // Hai mũi tên vòng tròn
                        dc.DrawEllipse(null, new System.Windows.Media.Pen(brushBlue, size / 7.0), 
                            new System.Windows.Point(s / 2, s / 2), s * 0.32, s * 0.32);
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 7.0), 
                            new System.Windows.Point(s * 0.82, s * 0.4), new System.Windows.Point(s * 0.82, s * 0.2));
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 7.0), 
                            new System.Windows.Point(s * 0.82, s * 0.2), new System.Windows.Point(s * 0.68, s * 0.28));
                        break;

                    case "SETTINGS":
                        dc.DrawEllipse(brushBlue, null, new System.Windows.Point(s / 2, s / 2), s * 0.32, s * 0.32);
                        dc.DrawEllipse(brushWhite, null, new System.Windows.Point(s / 2, s / 2), s * 0.15, s * 0.15);
                        break;

                    case "INFO":
                        dc.DrawEllipse(brushBlue, null, new System.Windows.Point(s / 2, s / 2), s * 0.38, s * 0.38);
                        dc.DrawEllipse(brushWhite, null, new System.Windows.Point(s / 2, s * 0.32), s * 0.06, s * 0.06);
                        dc.DrawRectangle(brushWhite, null, new System.Windows.Rect(s * 0.44, s * 0.45, s * 0.12, s * 0.3));
                        break;

                    case "DELETE":
                        dc.DrawEllipse(brushRed, null, new System.Windows.Point(s / 2, s / 2), s * 0.38, s * 0.38);
                        dc.DrawLine(new System.Windows.Media.Pen(brushWhite, size / 6.0), 
                            new System.Windows.Point(s * 0.3, s * 0.3), new System.Windows.Point(s * 0.7, s * 0.7));
                        dc.DrawLine(new System.Windows.Media.Pen(brushWhite, size / 6.0), 
                            new System.Windows.Point(s * 0.7, s * 0.3), new System.Windows.Point(s * 0.3, s * 0.7));
                        break;

                    case "CALCULATE":
                        // Máy tính - Calculator
                        dc.DrawRectangle(brushBlue, null, new System.Windows.Rect(s * 0.15, m, s * 0.7, s - 2 * m));
                        dc.DrawRectangle(brushLight, null, new System.Windows.Rect(s * 0.22, s * 0.15, s * 0.56, s * 0.2));
                        // Các nút
                        for (int row = 0; row < 3; row++)
                            for (int col = 0; col < 3; col++)
                                dc.DrawRectangle(brushWhite, null, 
                                    new System.Windows.Rect(s * 0.22 + col * s * 0.18, s * 0.42 + row * s * 0.16, s * 0.14, s * 0.12));
                        break;

                    case "DRAW":
                        // Đường polyline
                        var drawGeom = new StreamGeometry();
                        using (var ctx = drawGeom.Open())
                        {
                            ctx.BeginFigure(new System.Windows.Point(m, s * 0.75), false, false);
                            ctx.PolyLineTo(new[] {
                                new System.Windows.Point(s * 0.3, s * 0.35),
                                new System.Windows.Point(s * 0.55, s * 0.6),
                                new System.Windows.Point(s * 0.75, s * 0.25),
                                new System.Windows.Point(s - m, s * 0.45)
                            }, true, true);
                        }
                        dc.DrawGeometry(null, new System.Windows.Media.Pen(brushBlue, size / 5.0), drawGeom);
                        break;

                    case "VIEW":
                        // Mắt - View
                        var eyeGeom = new StreamGeometry();
                        using (var ctx = eyeGeom.Open())
                        {
                            ctx.BeginFigure(new System.Windows.Point(m, s / 2), false, false);
                            ctx.QuadraticBezierTo(new System.Windows.Point(s / 2, s * 0.2), new System.Windows.Point(s - m, s / 2), true, true);
                            ctx.QuadraticBezierTo(new System.Windows.Point(s / 2, s * 0.8), new System.Windows.Point(m, s / 2), true, true);
                        }
                        dc.DrawGeometry(brushLight, penBlue, eyeGeom);
                        dc.DrawEllipse(brushBlue, null, new System.Windows.Point(s / 2, s / 2), s * 0.15, s * 0.15);
                        dc.DrawEllipse(brushDark, null, new System.Windows.Point(s / 2, s / 2), s * 0.07, s * 0.07);
                        break;

                    case "COPY":
                        dc.DrawRectangle(brushLight, penBlue, new System.Windows.Rect(m, m, s * 0.55, s * 0.55));
                        dc.DrawRectangle(brushBlue, penDark, new System.Windows.Rect(s * 0.35, s * 0.35, s * 0.55, s * 0.55));
                        break;

                    case "MOVE":
                        // 4 mũi tên
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 6.0), 
                            new System.Windows.Point(s / 2, s * 0.15), new System.Windows.Point(s / 2, s * 0.85));
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 6.0), 
                            new System.Windows.Point(s * 0.15, s / 2), new System.Windows.Point(s * 0.85, s / 2));
                        // Mũi tên
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.4, s * 0.25), new System.Windows.Point(s / 2, s * 0.15));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.6, s * 0.25), new System.Windows.Point(s / 2, s * 0.15));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.75, s * 0.4), new System.Windows.Point(s * 0.85, s / 2));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.75, s * 0.6), new System.Windows.Point(s * 0.85, s / 2));
                        break;

                    case "RELOAD":
                        // Vòng tròn với mũi tên
                        dc.DrawEllipse(null, new System.Windows.Media.Pen(brushBlue, size / 5.0), 
                            new System.Windows.Point(s / 2, s / 2), s * 0.32, s * 0.32);
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 5.0), 
                            new System.Windows.Point(s * 0.82, s * 0.38), new System.Windows.Point(s * 0.82, s * 0.55));
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 5.0), 
                            new System.Windows.Point(s * 0.65, s * 0.38), new System.Windows.Point(s * 0.82, s * 0.38));
                        break;

                    // ===== THÊM ICON CHUYÊN NGÀNH GIAO THÔNG =====
                    case "PROFILE":
                        // Trắc dọc - Profile View (đường cao độ lên xuống)
                        var profileGeom = new StreamGeometry();
                        using (var ctx = profileGeom.Open())
                        {
                            ctx.BeginFigure(new System.Windows.Point(m, s * 0.7), false, false);
                            ctx.PolyLineTo(new[] {
                                new System.Windows.Point(s * 0.2, s * 0.5),
                                new System.Windows.Point(s * 0.35, s * 0.6),
                                new System.Windows.Point(s * 0.5, s * 0.35),
                                new System.Windows.Point(s * 0.7, s * 0.45),
                                new System.Windows.Point(s - m, s * 0.3)
                            }, true, true);
                        }
                        dc.DrawGeometry(null, new System.Windows.Media.Pen(brushBlue, size / 5.0), profileGeom);
                        // Đường cơ sở
                        dc.DrawLine(penDark, new System.Windows.Point(m, s * 0.8), new System.Windows.Point(s - m, s * 0.8));
                        // Các vạch chia lý trình
                        for (double x = s * 0.2; x <= s * 0.8; x += s * 0.2)
                            dc.DrawLine(penDark, new System.Windows.Point(x, s * 0.78), new System.Windows.Point(x, s * 0.82));
                        break;

                    case "SPIRAL":
                        // Đường cong chuyển tiếp - Clothoid/Spiral
                        var spiralGeom = new StreamGeometry();
                        using (var ctx = spiralGeom.Open())
                        {
                            ctx.BeginFigure(new System.Windows.Point(m, s * 0.8), false, false);
                            ctx.BezierTo(
                                new System.Windows.Point(s * 0.3, s * 0.8),
                                new System.Windows.Point(s * 0.4, s * 0.5),
                                new System.Windows.Point(s * 0.5, s * 0.35),
                                true, true);
                            ctx.BezierTo(
                                new System.Windows.Point(s * 0.6, s * 0.2),
                                new System.Windows.Point(s * 0.8, s * 0.2),
                                new System.Windows.Point(s - m, s * 0.2),
                                true, true);
                        }
                        dc.DrawGeometry(null, new System.Windows.Media.Pen(brushBlue, size / 5.0), spiralGeom);
                        // Ký hiệu TS, SC
                        dc.DrawEllipse(brushGreen, null, new System.Windows.Point(s * 0.15, s * 0.8), s * 0.06, s * 0.06);
                        dc.DrawEllipse(brushRed, null, new System.Windows.Point(s * 0.5, s * 0.35), s * 0.06, s * 0.06);
                        break;

                    case "SUPERELEVATION":
                        // Siêu cao - Độ nghiêng ngang mặt đường
                        // Mặt đường nghiêng
                        var superGeom = new StreamGeometry();
                        using (var ctx = superGeom.Open())
                        {
                            ctx.BeginFigure(new System.Windows.Point(m, s * 0.6), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.5, s * 0.4), true, true);
                            ctx.LineTo(new System.Windows.Point(s - m, s * 0.5), true, true);
                            ctx.LineTo(new System.Windows.Point(s - m, s * 0.7), true, true);
                            ctx.LineTo(new System.Windows.Point(m, s * 0.8), true, true);
                        }
                        dc.DrawGeometry(brushLight, penBlue, superGeom);
                        // Mũi tên chỉ hướng nghiêng
                        dc.DrawLine(penDark, new System.Windows.Point(s * 0.3, s * 0.5), new System.Windows.Point(s * 0.45, s * 0.42));
                        dc.DrawLine(penDark, new System.Windows.Point(s * 0.4, s * 0.38), new System.Windows.Point(s * 0.45, s * 0.42));
                        dc.DrawLine(penDark, new System.Windows.Point(s * 0.48, s * 0.48), new System.Windows.Point(s * 0.45, s * 0.42));
                        break;

                    case "CULVERT":
                        // Cống hộp - Box Culvert
                        dc.DrawRectangle(brushLight, new System.Windows.Media.Pen(brushDark, size / 8.0), 
                            new System.Windows.Rect(s * 0.25, s * 0.3, s * 0.5, s * 0.4));
                        // Nước chảy qua cống
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 6.0), 
                            new System.Windows.Point(m, s * 0.55), new System.Windows.Point(s * 0.25, s * 0.55));
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 6.0), 
                            new System.Windows.Point(s * 0.75, s * 0.55), new System.Windows.Point(s - m, s * 0.55));
                        // Mũi tên hướng chảy
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.85, s * 0.48), new System.Windows.Point(s - m, s * 0.55));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.85, s * 0.62), new System.Windows.Point(s - m, s * 0.55));
                        break;

                    case "BRIDGE":
                        // Cầu - Bridge
                        // Mặt cầu
                        dc.DrawRectangle(brushBlue, penDark, new System.Windows.Rect(m, s * 0.35, s - 2 * m, s * 0.15));
                        // Trụ cầu
                        dc.DrawRectangle(brushDark, null, new System.Windows.Rect(s * 0.2, s * 0.5, s * 0.1, s * 0.35));
                        dc.DrawRectangle(brushDark, null, new System.Windows.Rect(s * 0.7, s * 0.5, s * 0.1, s * 0.35));
                        // Nước bên dưới
                        dc.DrawLine(new System.Windows.Media.Pen(brushLight, size / 8.0), 
                            new System.Windows.Point(m, s * 0.75), new System.Windows.Point(s - m, s * 0.75));
                        break;

                    case "VOLUME":
                        // Khối lượng đào đắp - Cut/Fill Volume
                        // Vùng đào (màu đỏ nhạt)
                        var cutGeom = new StreamGeometry();
                        using (var ctx = cutGeom.Open())
                        {
                            ctx.BeginFigure(new System.Windows.Point(m, s * 0.5), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.4, s * 0.5), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.25, s * 0.25), true, true);
                            ctx.LineTo(new System.Windows.Point(m, s * 0.35), true, true);
                        }
                        dc.DrawGeometry(new SolidColorBrush(System.Windows.Media.Color.FromArgb(150, 220, 53, 69)), null, cutGeom);
                        // Vùng đắp (màu xanh nhạt)
                        var fillGeom = new StreamGeometry();
                        using (var ctx = fillGeom.Open())
                        {
                            ctx.BeginFigure(new System.Windows.Point(s * 0.6, s * 0.5), true, true);
                            ctx.LineTo(new System.Windows.Point(s - m, s * 0.5), true, true);
                            ctx.LineTo(new System.Windows.Point(s - m, s * 0.7), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.75, s * 0.65), true, true);
                        }
                        dc.DrawGeometry(new SolidColorBrush(System.Windows.Media.Color.FromArgb(150, 40, 167, 69)), null, fillGeom);
                        // Đường thiết kế
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 6.0), 
                            new System.Windows.Point(m, s * 0.5), new System.Windows.Point(s - m, s * 0.5));
                        break;

                    case "SLOPE":
                        // Taluy - Mái dốc
                        var slopeGeom = new StreamGeometry();
                        using (var ctx = slopeGeom.Open())
                        {
                            ctx.BeginFigure(new System.Windows.Point(m, s * 0.8), true, true);
                            ctx.LineTo(new System.Windows.Point(s * 0.4, s * 0.3), true, true);
                            ctx.LineTo(new System.Windows.Point(s - m, s * 0.3), true, true);
                            ctx.LineTo(new System.Windows.Point(s - m, s * 0.8), true, true);
                        }
                        dc.DrawGeometry(brushLight, penBlue, slopeGeom);
                        // Gạch chéo taluy
                        for (double i = 0; i < 5; i++)
                        {
                            double x1 = m + i * (s * 0.3) / 4;
                            double y1 = s * 0.8 - i * (s * 0.5) / 4;
                            dc.DrawLine(penBlue, new System.Windows.Point(x1 + s * 0.08, y1 - s * 0.1), new System.Windows.Point(x1, y1));
                        }
                        break;

                    case "EXCEL":
                        // Xuất Excel - Spreadsheet
                        dc.DrawRectangle(brushGreen, null, new System.Windows.Rect(m, m, s - 2 * m, s - 2 * m));
                        // Các ô bảng tính
                        for (int row = 0; row < 4; row++)
                            for (int col = 0; col < 3; col++)
                                dc.DrawRectangle(brushWhite, null, 
                                    new System.Windows.Rect(m + s * 0.08 + col * s * 0.25, m + s * 0.08 + row * s * 0.2, s * 0.2, s * 0.15));
                        break;

                    case "CAD":
                        // Xuất CAD - AutoCAD
                        dc.DrawRectangle(brushDark, null, new System.Windows.Rect(m, m, s - 2 * m, s - 2 * m));
                        // Chữ DWG
                        dc.DrawRectangle(brushBlue, null, new System.Windows.Rect(s * 0.15, s * 0.4, s * 0.7, s * 0.25));
                        break;

                    case "WIDENING":
                        // Mở rộng đường - Road Widening
                        dc.DrawLine(new System.Windows.Media.Pen(brushDark, size / 8.0), 
                            new System.Windows.Point(m, s * 0.4), new System.Windows.Point(s - m, s * 0.4));
                        dc.DrawLine(new System.Windows.Media.Pen(brushDark, size / 8.0), 
                            new System.Windows.Point(m, s * 0.6), new System.Windows.Point(s - m, s * 0.6));
                        // Phần mở rộng (đường đứt)
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 8.0) { DashStyle = System.Windows.Media.DashStyles.Dash }, 
                            new System.Windows.Point(s * 0.4, s * 0.25), new System.Windows.Point(s - m, s * 0.25));
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 8.0) { DashStyle = System.Windows.Media.DashStyles.Dash }, 
                            new System.Windows.Point(s * 0.4, s * 0.75), new System.Windows.Point(s - m, s * 0.75));
                        // Vùng mở rộng
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.4, s * 0.25), new System.Windows.Point(s * 0.4, s * 0.4));
                        dc.DrawLine(penBlue, new System.Windows.Point(s * 0.4, s * 0.6), new System.Windows.Point(s * 0.4, s * 0.75));
                        break;

                    case "INTERSECTION":
                        // Nút giao - Intersection
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 5.0), 
                            new System.Windows.Point(s / 2, m), new System.Windows.Point(s / 2, s - m));
                        dc.DrawLine(new System.Windows.Media.Pen(brushBlue, size / 5.0), 
                            new System.Windows.Point(m, s / 2), new System.Windows.Point(s - m, s / 2));
                        // Bo góc
                        dc.DrawEllipse(null, penDark, new System.Windows.Point(s / 2, s / 2), s * 0.22, s * 0.22);
                        break;

                    default:
                        dc.DrawEllipse(brushBlue, null, new System.Windows.Point(s / 2, s / 2), s * 0.35, s * 0.35);
                        break;
                }
            }

            RenderTargetBitmap rtb = new RenderTargetBitmap(size, size, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(visual);
            return rtb;
        }

        #endregion

        #region Command Lists

        private static List<(string Command, string Label, string SubIcon)> GetSurfaceCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTSu_TaoSpotElevation_OnSurface_TaiTim", "Spot Elevation Tại Tim", "ADD"),
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetSampleLineCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTS_DoiTenCoc", "Đổi Tên Cọc", "RENAME"),
                ("CTS_DoiTenCoc2", "Đổi Tên Cọc Đoạn", "RENAME"),
                ("CTS_DoiTenCoc3", "Đổi Tên Cọc Km", "RENAME"),
                ("CTS_DoiTenCoc_fromCogoPoint", "Đổi Tên Từ CogoPoint", "RENAME"),
                ("CTS_DoiTenCoc_TheoThuTu", "Đổi Tên Thứ Tự", "RENAME"),
                ("CTS_DoiTenCoc_H", "Đổi Tên Hậu Tố A", "RENAME"),
                ("---", "", ""),
                ("CTS_TaoBang_ToaDoCoc", "Tọa Độ Cọc (X,Y)", "TABLE"),
                ("CTS_TaoBang_ToaDoCoc2", "Tọa Độ Cọc (Lý Trình)", "TABLE"),
                ("CTS_TaoBang_ToaDoCoc3", "Tọa Độ Cọc (Cao Độ)", "TABLE"),
                ("AT_UPdate2Table", "Cập Nhật Từ Bảng", "SYNC"),
                ("---", "", ""),
                ("CTS_ChenCoc_TrenTracDoc", "Chèn Trên Trắc Dọc", "ADD"),
                ("CTS_CHENCOC_TRENTRACNGANG", "Chèn Trên Trắc Ngang", "ADD"),
                ("CTS_PhatSinhCoc", "Phát Sinh Cọc Auto", "ADD"),
                ("CTS_PhatSinhCoc_ChiTiet", "Phát Sinh Chi Tiết", "ADD"),
                ("CTS_PhatSinhCoc_theoKhoangDelta", "Phát Sinh Delta", "ADD"),
                ("CTS_PhatSinhCoc_TuCogoPoint", "Phát Sinh Từ CogoPoint", "ADD"),
                ("CTS_PhatSinhCoc_TheoBang", "Phát Sinh Từ Bảng", "ADD"),
                ("---", "", ""),
                ("CTS_DichCoc_TinhTien", "Dịch Cọc Tịnh Tiến", "MOVE"),
                ("CTS_DichCoc_TinhTien40", "Dịch Cọc 40m", "MOVE"),
                ("CTS_DichCoc_TinhTien_20", "Dịch Cọc 20m", "MOVE"),
                ("CTS_Copy_NhomCoc", "Sao Chép Nhóm Cọc", "COPY"),
                ("CTS_DongBo_2_NhomCoc", "Đồng Bộ Nhóm Cọc", "SYNC"),
                ("CTS_DongBo_2_NhomCoc_TheoDoan", "Đồng Bộ Theo Đoạn", "SYNC"),
                ("---", "", ""),
                ("CTS_Copy_BeRong_sampleLine", "Copy Bề Rộng SL", "COPY"),
                ("CTS_Thaydoi_BeRong_sampleLine", "Thay Đổi Bề Rộng SL", "SETTINGS"),
                ("CTS_Offset_BeRong_sampleLine", "Offset Bề Rộng SL", "MOVE"),
                ("---", "", ""),
                ("CTSV_ThongKeCoc", "Thống Kê Cọc (Excel)", "EXPORT"),
                ("CTSV_ThongKeCoc_TatCa", "Thống Kê Tất Cả Cọc", "EXPORT")
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetAlignmentCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTPV_TaoProfileView", "Tạo Trắc Dọc", "PROFILE"),
                ("CTPV_SuaProfileView", "Edit Profile", "SETTINGS"),
                ("CTPV_ThemBang_LyTrinh", "Thêm Bảng Lý Trình", "TABLE"),
                ("CTPV_ThemLabel_CaoDo", "Thêm Label Cao Độ", "ADD"),
                ("CTPV_ThayDoiScale", "Thay Đổi Scale", "SETTINGS"),
                ("CTPV_FitKhung", "Fit Khung", "VIEW")
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetCorridorCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTC_AddAllSection", "Thêm Tất Cả Section", "ADD"),
                ("CTC_TaoCooridor_DuongDoThi_RePhai", "Corridor Rẽ Phải", "DRAW")
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetSectionViewCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTSV_VeTracNgangThietKe", "Tạo Trắc Ngang", "DRAW"),
                ("CVSV_VeTatCa_TracNgangThietKe", "Vẽ Tất Cả TN", "DRAW"),
                ("CTSV_ChuyenDoi_TNTK_TNTN", "Chuyển TK Sang TN", "SYNC"),
                ("---", "", ""),
                ("CTSV_DanhCap", "Đánh Cấp - VHC", "SLOPE"),
                ("CTSV_DanhCap_XoaBo", "Xóa Đánh Cấp", "DELETE"),
                ("CTSV_DanhCap_VeThem", "Vẽ Thêm Đánh Cấp", "SLOPE"),
                ("CTSV_DanhCap_VeThem1", "Vẽ Thêm 1m", "ADD"),
                ("CTSV_DanhCap_VeThem2", "Vẽ Thêm 2m", "ADD"),
                ("CTSV_DanhCap_CapNhat", "Cập Nhật KL Đánh Cấp", "SYNC"),
                ("---", "", ""),
                ("CTSV_ThemVatLieu_TrenCatNgang", "Điền KL Trắc Ngang", "VOLUME"),
                ("CTSV_ThayDoi_MSS_Min_Max", "Hiệu Chỉnh MSS", "SETTINGS"),
                ("CTSV_ThayDoi_GioiHan_traiPhai", "Thay Giới Hạn T/P", "SETTINGS"),
                ("CTSV_ThayDoi_KhungIn", "Dàn Khung In", "VIEW"),
                ("CTSV_KhoaCatNgang_AddPoint", "Khóa TN + Add Point", "ADD"),
                ("---", "", ""),
                ("CTSV_fit_KhungIn", "Fit Khung In", "VIEW"),
                ("CTSV_fit_KhungIn_5_5_top", "Fit Khung 5x5", "VIEW"),
                ("CTSV_fit_KhungIn_5_10_top", "Fit Khung 5x10", "VIEW"),
                ("---", "", ""),
                ("CTSV_An_DuongDiaChat", "Ẩn Đường Địa Chất", "VIEW"),
                ("CTSV_HieuChinh_Section", "Hiệu Chỉnh (Static)", "SETTINGS"),
                ("CTSV_HieuChinh_Section_Dynamic", "Hiệu Chỉnh (Dynamic)", "SETTINGS"),
                ("---", "", ""),
                ("CTSV_Taskbar", "Taskbar Khối Lượng", "VOLUME"),
                ("CTSV_XuatKhoiLuong", "Xuất KL Excel", "EXCEL"),
                ("CTSV_XuatCad", "Xuất KL CAD", "CAD"),
                ("CTSV_CaiDatBang", "Cài Đặt Bảng KL", "SETTINGS")
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetGradingCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTSN_Taskbar", "Mở Taskbar SN", "VOLUME"),
                ("---", "", ""),
                ("CTSN_TaoLuoi", "Quản Lý Lưới", "GRADING"),
                ("CTSN_NhapCaoDo", "Điền Cao Độ Lưới", "ADD"),
                ("CTSN_Surface", "Lấy CĐ Surface", "SURFACE"),
                ("CTSN_TinhKL", "Tính Khối Lượng SN", "VOLUME"),
                ("CTSN_XuatBang", "Xuất Bảng KL CAD", "CAD")
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetPipeCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTPS_TaoBangThongKePipe", "Thống Kê Pipe", "TABLE"),
                ("CTPS_TaoBangThongKeStructure", "Thống Kê Structure", "TABLE"),
                ("CTPS_ThayDoi_CaoDo_Pipe", "Đổi Cao Độ Pipe", "SETTINGS"),
                ("CTPS_ThayDoi_CaoDo_Structure", "Đổi Cao Độ Struct", "SETTINGS"),
                ("CTPS_XoayPipe_90do", "Xoay Pipe 90°", "SYNC"),
                ("CTPS_XoaConTrung", "Xóa Con Trùng", "DELETE")
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetPointCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTPo_TaoPointTheoBang", "Tạo Point Từ Bảng", "ADD"),
                ("CTPo_ChuyenPointThanhBlock", "Point → Block", "SYNC"),
                ("CTPo_TaoBangThongKePoint", "Bảng Thống Kê Point", "TABLE"),
                ("CTPo_ThayDoiCaoDo", "Thay Đổi Cao Độ", "SETTINGS"),
                ("CTPo_DatTen_theoThuTu", "Đặt Tên Thứ Tự", "RENAME"),
                ("CTPo_ThayDoiStyle", "Thay Đổi Style", "SETTINGS"),
                ("CTPo_LayThongTin", "Lấy Thông Tin", "INFO")
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetUtilityCommands()
        {
            return new List<(string, string, string)>
            {
                ("AT_Solid_Set_PropertySet", "Gán Property Set", "ADD"),
                ("AT_Solid_Show_Info", "Thông Tin Solid", "INFO"),
                ("CT_VTOADOHG", "Tọa Độ Hố Ga", "POINT"),
                ("---", "", ""),
                ("CIVIL_RIBBON", "Reload Menu", "RELOAD")
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetAccountCommands()
        {
            return new List<(string, string, string)>
            {
                ("", "Đăng Nhập", "ACCOUNT"),
                ("", "Thông Tin", "INFO"),
                ("", "Hướng Dẫn", "INFO")
            };
        }

        // Acad Tool Commands
        private static List<(string Command, string Label, string SubIcon)> GetMeasureCommands()
        {
            return new List<(string, string, string)>
            {
                ("AT_TongDoDai_Full", "Tổng Độ Dài (Full)", "CALCULATE"),
                ("AT_TongDoDai_Replace", "Tổng Độ Dài (Replace)", "CALCULATE"),
                ("AT_TongDoDai_Replace2", "Tổng Độ Dài (Replace2)", "CALCULATE"),
                ("AT_TongDoDai_Replace_CongThem", "Tổng Độ Dài (Cộng Thêm)", "CALCULATE"),
                ("---", "", ""),
                ("AT_TongDienTich_Full", "Tổng Diện Tích (Full)", "CALCULATE"),
                ("AT_TongDienTich_Replace", "Tổng Diện Tích (Replace)", "CALCULATE"),
                ("AT_TongDienTich_Replace2", "Tổng Diện Tích (Replace2)", "CALCULATE"),
                ("AT_TongDienTich_Replace_CongThem", "Tổng Diện Tích (Cộng Thêm)", "CALCULATE")
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetEditCommands()
        {
            return new List<(string, string, string)>
            {
                ("AT_TextLink", "Text Link", "ADD"),
                ("AT_DanhSoThuTu", "Đánh Số Thứ Tự", "ADD"),
                ("AT_XoaDoiTuong_CungLayer", "Xóa Đối Tượng Cùng Layer", "DELETE"),
                ("AT_XoaDoiTuong_3DSolid_Body", "Xóa 3DSolid/Body", "DELETE"),
                ("AT_Offset_2Ben", "Offset 2 Bên", "COPY"),
                ("AT_annotive_scale_currentOnly", "Annotative Scale Current Only", "SETTINGS"),
                ("---", "", ""),
                ("AT_XoayDoiTuong_TheoViewport", "Xoay Theo Viewport", "SYNC"),
                ("AT_XoayDoiTuong_Theo2Diem", "Xoay Theo 2 Điểm", "SYNC")
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetLayoutCommands()
        {
            return new List<(string, string, string)>
            {
                ("AT_TextLayout", "Text Layout", "ADD"),
                ("AT_TaoMoi_TextLayout", "Tạo Mới Text Layout", "ADD"),
                ("AT_DimLayout", "Dim Layout", "ADD"),
                ("AT_DimLayout2", "Dim Layout 2", "ADD"),
                ("AT_BlockLayout", "Block Layout", "ADD"),
                ("AT_Label_FromText", "Label From Text", "ADD"),
                ("---", "", ""),
                ("AT_UpdateLayout", "Update Layout", "SYNC")
            };
        }

        #endregion
    }

    /// <summary>
    /// Command Handler cho Ribbon buttons
    /// </summary>
    public class CivilRibbonCommandHandler : System.Windows.Input.ICommand
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

                var doc = AcadApplication.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.SendStringToExecute(commandToRun + " ", true, false, true);
                }
            }
            catch (System.Exception ex)
            {
                var ed = AcadApplication.DocumentManager.MdiActiveDocument?.Editor;
                ed?.WriteMessage($"\n❌ Lỗi thực thi lệnh: {ex.Message}");
            }
        }
    }
}
