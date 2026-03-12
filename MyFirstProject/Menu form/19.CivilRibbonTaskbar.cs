// 19.CivilRibbonTaskbar.cs - Thanh Ribbon Civil Tool
// Thiết kế hiện đại theo phong cách Civil 3D - Gộp 1 tab duy nhất

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
    /// Civil Ribbon Taskbar - Thiết kế hiện đại theo phong cách Civil 3D
    /// Gộp tất cả lệnh Civil + ACAD vào 1 tab duy nhất
    /// </summary>
    public class CivilRibbonTaskbar
    {
        // ═══ Bảng màu Navy thống nhất ═══
        private static readonly System.Windows.Media.Color PrimaryBlue = System.Windows.Media.Color.FromRgb(27, 42, 74);      // #1B2A4A Navy đậm
        private static readonly System.Windows.Media.Color LightBlue = System.Windows.Media.Color.FromRgb(210, 220, 235);     // Highlight nhạt
        private static readonly System.Windows.Media.Color DarkBlue = System.Windows.Media.Color.FromRgb(16, 28, 52);         // Navy tối
        private static readonly System.Windows.Media.Color AccentBlue = System.Windows.Media.Color.FromRgb(46, 74, 122);      // #2E4A7A Navy vừa
        private static readonly System.Windows.Media.Color AccentTeal = System.Windows.Media.Color.FromRgb(46, 74, 122);      // Dùng Navy thay Teal
        private static readonly System.Windows.Media.Color SoftGray = System.Windows.Media.Color.FromRgb(122, 122, 122);      // Xám
        
        // Số lệnh featured (hiện nút lớn) cho mỗi panel
        private const int FEATURED_COUNT = 1;

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

                // Load config đã lưu (từ MenuSettingsForm)
                var config = MenuSettingsForm.LoadSavedConfig();

                // Gộp tất cả vào 1 tab duy nhất
                CreateSingleCivilTab(ribbon, config);

                var ed = AcadApplication.DocumentManager.MdiActiveDocument?.Editor;
                ed?.WriteMessage("\n✅ Civil Ribbon đã sẵn sàng!");
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

        /// <summary>
        /// Gộp tất cả Civil + ACAD groups vào 1 tab duy nhất
        /// </summary>
        private static void CreateSingleCivilTab(RibbonControl ribbon, MenuConfig config)
        {
            RibbonTab tab = new RibbonTab
            {
                Title = "Civil Tool",
                Id = "CivilTool.MainTab"
            };
            ribbon.Tabs.Add(tab);

            // Thêm tất cả Civil groups
            foreach (var group in config.CivilGroups)
            {
                if (group.Commands.Count == 0) continue;
                var cmds = group.Commands.Select(c => (c.Command, c.Label, c.SubIcon)).ToList();
                AddCivilPanel(tab, group.GroupName, group.IconType, cmds);
            }

            // Thêm tất cả ACAD groups vào cùng tab
            foreach (var group in config.AcadGroups)
            {
                if (group.Commands.Count == 0) continue;
                var cmds = group.Commands.Select(c => (c.Command, c.Label, c.SubIcon)).ToList();
                AddCivilPanel(tab, group.GroupName, group.IconType, cmds);
            }

            tab.IsActive = true;
        }

        /// <summary>
        /// Tạo panel kiểu Civil 3D hiện đại:
        /// - Lệnh đầu tiên: NÚT LỚN (Large, icon 24px + text bên dưới)
        /// - Các lệnh tiếp: HÀNG NÚT NHỎ (3 nút/hàng, icon 12px + text bên cạnh)
        /// - Nếu nhiều quá (&gt;7): gộp vào SplitButton dropdown
        /// </summary>
        public static void AddCivilPanel(RibbonTab tab, string title, string iconType, 
            List<(string Command, string Label, string SubIcon)> commands)
        {
            if (commands.Count == 0) return;

            // Lọc bỏ separator, chỉ lấy lệnh thực
            var realCmds = commands.Where(c => c.Command != "---" && !string.IsNullOrEmpty(c.Command)).ToList();
            if (realCmds.Count == 0) return;

            RibbonPanelSource panelSource = new RibbonPanelSource { Title = title };
            RibbonPanel panel = new RibbonPanel { Source = panelSource };

            // ═══ 1. NÚT LỚN - Lệnh featured đầu tiên ═══
            int featuredCount = Math.Min(FEATURED_COUNT, realCmds.Count);
            for (int i = 0; i < featuredCount; i++)
            {
                var (command, label, subIcon) = realCmds[i];
                RibbonButton largeBtn = new RibbonButton
                {
                    Text = label,
                    ShowText = true,
                    ShowImage = true,
                    Image = CreateCivilIcon(subIcon, 12),
                    LargeImage = CreateCivilIcon(subIcon, 24),
                    Size = RibbonItemSize.Large,
                    Orientation = System.Windows.Controls.Orientation.Vertical,
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
                panelSource.Items.Add(largeBtn);
            }

            // ═══ 2. HÀNG NÚT NHỎ - Các lệnh phụ xếp 3 nút/hàng ═══
            var secondaryCmds = realCmds.Skip(featuredCount).ToList();
            
            if (secondaryCmds.Count > 0 && secondaryCmds.Count <= 9)
            {
                // Xếp thành các hàng, mỗi hàng tối đa 3 nút nhỏ
                RibbonRowPanel rowPanel = new RibbonRowPanel();
                int colCount = 0;

                foreach (var (command, label, subIcon) in secondaryCmds)
                {
                    RibbonButton smallBtn = CreateSmallButton(command, label, subIcon);
                    rowPanel.Items.Add(smallBtn);
                    colCount++;

                    // Xuống hàng sau mỗi 3 nút (trừ hàng cuối)
                    if (colCount % 3 == 0 && colCount < secondaryCmds.Count)
                    {
                        rowPanel.Items.Add(new RibbonRowBreak());
                    }
                }

                panelSource.Items.Add(rowPanel);
            }
            else if (secondaryCmds.Count > 9)
            {
                // Quá nhiều → Hiện 6 nút nhỏ + SplitButton dropdown cho phần còn lại
                var visibleCmds = secondaryCmds.Take(6).ToList();
                var overflowCmds = secondaryCmds.Skip(6).ToList();

                RibbonRowPanel rowPanel = new RibbonRowPanel();
                int colCount = 0;
                foreach (var (command, label, subIcon) in visibleCmds)
                {
                    rowPanel.Items.Add(CreateSmallButton(command, label, subIcon));
                    colCount++;
                    if (colCount % 3 == 0 && colCount < visibleCmds.Count)
                        rowPanel.Items.Add(new RibbonRowBreak());
                }
                panelSource.Items.Add(rowPanel);

                // SplitButton dropdown cho overflow commands
                RibbonSplitButton overflowSplit = new RibbonSplitButton
                {
                    Text = "Thêm...",
                    ShowText = true,
                    ShowImage = true,
                    Size = RibbonItemSize.Large,
                    Orientation = System.Windows.Controls.Orientation.Vertical,
                    Image = CreateCivilIcon(iconType, 12),
                    LargeImage = CreateCivilIcon(iconType, 24),
                    ListStyle = RibbonSplitButtonListStyle.List
                };

                foreach (var (command, label, subIcon) in overflowCmds)
                {
                    overflowSplit.Items.Add(CreateSmallButton(command, label, subIcon));
                }
                panelSource.Items.Add(overflowSplit);
            }

            tab.Panels.Add(panel);
        }

        /// <summary>
        /// Tạo nút nhỏ chuẩn (Standard size, icon 12px + text bên cạnh)
        /// </summary>
        private static RibbonButton CreateSmallButton(string command, string label, string subIcon)
        {
            return new RibbonButton
            {
                Text = label,
                ShowText = true,
                ShowImage = true,
                Image = CreateCivilIcon(subIcon, 12),
                LargeImage = CreateCivilIcon(subIcon, 24),
                Size = RibbonItemSize.Standard,
                Orientation = System.Windows.Controls.Orientation.Horizontal,
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
        }

        #region Icon Generator - Civil Engineering Theme

        public static ImageSource CreateCivilIcon(string iconType, int size)
        {
            // Tạo icon theo chuẩn ký hiệu kỹ thuật giao thông
            DrawingVisual visual = new DrawingVisual();
            using (DrawingContext dc = visual.RenderOpen())
            {
                // Gradient brushes cho icon hiện đại
                var gradBlue = new LinearGradientBrush(PrimaryBlue, DarkBlue, 45);
                var gradAccent = new LinearGradientBrush(AccentBlue, PrimaryBlue, 90);
                var gradTeal = new LinearGradientBrush(AccentTeal, System.Windows.Media.Color.FromRgb(0, 130, 120), 45);
                
                var brushBlue = new SolidColorBrush(PrimaryBlue);
                var brushLight = new SolidColorBrush(LightBlue);
                var brushDark = new SolidColorBrush(DarkBlue);
                var brushAccent = new SolidColorBrush(AccentBlue);
                var brushWhite = System.Windows.Media.Brushes.White;
                var brushRed = new SolidColorBrush(System.Windows.Media.Color.FromRgb(232, 65, 78));
                var brushGreen = new SolidColorBrush(System.Windows.Media.Color.FromRgb(45, 183, 80));
                var brushBrown = new SolidColorBrush(System.Windows.Media.Color.FromRgb(155, 100, 50));
                var brushOrange = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 152, 0));
                
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

        // ═══ CIVIL TOOL Commands - theo bảng thống kê ═══

        private static List<(string Command, string Label, string SubIcon)> GetCorridorCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTC_AddAllSection", "Thêm tất cả Section", "ADD"),
                ("CTC_TaoCooridor_DuongDoThi_RePhai", "Corridor rẽ phải", "DRAW"),
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetParcelCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTPA_TaoParcel_CacLoaiNha", "Parcel từ Polyline", "DRAW"),
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetPipeCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTPI_ThayDoi_DuongKinhCong", "Đổi ĐK cống", "SETTINGS"),
                ("CTPI_ThayDoi_MatPhangRef_Cong", "MP ref cống", "SETTINGS"),
                ("CTPI_ThayDoi_DoanDocCong", "Đổi dốc cống", "SETTINGS"),
                ("CTPI_BangCaoDo_TuNhienHoThu", "Bảng CĐ hố thu", "TABLE"),
                ("CTPI_XoayHoThu_Theo2diem", "Xoay hố thu", "SYNC"),
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetPointCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTPO_TaoCogoPoint_CaoDo_FromSurface", "CĐ từ Surface", "ADD"),
                ("CTPO_TaoCogoPoint_CaoDo_Elevations", "CĐ từ Spot", "ADD"),
                ("CTPO_CreateCogopointFromText", "Từ Text CĐ", "ADD"),
                ("CTPO_An_CogoPoint", "Ẩn CogoPoint", "VIEW"),
                ("CTPO_UpdateAllPointGroup", "Update PtGroup", "SYNC"),
                ("CTPO_CPTT", "Point→Text CĐ", "SYNC"),
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetProfileCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTP_VeTracDoc_TuNhien", "Trắc dọc TN", "PROFILE"),
                ("CTP_VeTracDoc_TuNhien_TatCaTuyen", "TD TN tất cả", "PROFILE"),
                ("CTP_Fix_DuongTuNhien_TheoCoc", "Fix TN theo cọc", "SETTINGS"),
                ("CTP_GanNhanNutGiao_LenTracDoc", "Nhãn nút giao", "ADD"),
                ("CTP_TaoCogoPointTuPVI", "CogoPoint từ PVI", "POINT"),
                ("CTP_TaoBangThongKeParcel", "TK Parcel", "TABLE"),
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetSampleLineCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTS_DoiTenCoc", "Đổi tên cọc", "RENAME"),
                ("CTS_DoiTenCoc2", "Đổi tên đoạn", "RENAME"),
                ("CTS_DoiTenCoc3", "Đổi tên Km", "RENAME"),
                ("---", "", ""),
                ("CTS_TaoBang_ToaDoCoc", "Bảng tọa độ", "TABLE"),
                ("CTS_PhatSinhCoc", "Phát sinh auto", "ADD"),
                ("CTS_PhatSinhCoc_ChiTiet", "Phát sinh CT", "ADD"),
                ("---", "", ""),
                ("CTS_DichCoc_TinhTien", "Dịch tịnh tiến", "MOVE"),
                ("CTS_Copy_NhomCoc", "Copy nhóm cọc", "COPY"),
                ("CTS_DongBo_2_NhomCoc", "Đồng bộ 2NC", "SYNC"),
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetSectionViewCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTSV_VeTracNgangThietKe", "Vẽ TN thiết kế", "DRAW"),
                ("---", "", ""),
                ("CTSV_DanhCap", "Đánh cấp VHC", "SLOPE"),
                ("CTSV_DanhCap_XoaBo", "Xóa đánh cấp", "DELETE"),
                ("CTSV_DanhCap_VeThem", "Vẽ thêm ĐC", "SLOPE"),
                ("---", "", ""),
                ("CTSV_ThemVatLieu_TrenCatNgang", "Điền vật liệu", "VOLUME"),
                ("CTSV_fit_KhungIn", "Fit khung in", "VIEW"),
                ("CTSV_An_DuongDiaChat", "Ẩn địa chất", "VIEW"),
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetVolumeCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTSV_Taskbar", "Taskbar KL", "VOLUME"),
                ("CTSV_XuatKhoiLuong", "Xuất KL Excel", "EXCEL"),
                ("CTSV_XuatCad", "Xuất KL CAD", "CAD"),
                ("CTSV_CaiDatBang", "Cài đặt bảng", "SETTINGS"),
                ("---", "", ""),
                ("CTSV_ThongKeCoc", "TK cọc Excel", "EXPORT"),
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetGradingCommands()
        {
            return new List<(string, string, string)>
            {
                ("CTSN_Taskbar", "Taskbar SN", "VOLUME"),
                ("---", "", ""),
                ("CTSN_TaoLuoi", "Tạo lưới", "GRADING"),
                ("CTSN_NhapCaoDo", "Nhập CĐ lưới", "ADD"),
                ("CTSN_Surface", "Lấy CĐ Surface", "SURFACE"),
                ("CTSN_TinhKL", "Tính KL SN", "CALCULATE"),
                ("CTSN_XuatBang", "Xuất bảng CAD", "CAD"),
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetToolCommands()
        {
            return new List<(string, string, string)>
            {
                ("CT_Taskbar", "Thanh công cụ", "UTILITY"),
                ("CT_VTOADOHG", "Tọa độ hố ga", "POINT"),
                ("CT_MENU_SETTINGS", "Cài đặt menu", "SETTINGS"),
                ("CTC_TaoMViewTheoTuyen", "Tạo MView theo tuyến", "VIEW"),
                ("---", "", ""),
                ("CIVIL_RIBBON", "Reload menu", "RELOAD"),
            };
        }

        // ═══ ACAD TOOL Commands ═══

        private static List<(string Command, string Label, string SubIcon)> GetMeasureCommands()
        {
            return new List<(string, string, string)>
            {
                ("AT_TongDoDai_Full", "Tổng dài (Full)", "CALCULATE"),
                ("AT_TongDoDai_Replace", "Tổng dài (Replace)", "CALCULATE"),
                ("AT_TongDoDai_Replace2", "Tổng dài (Replace2)", "CALCULATE"),
                ("AT_TongDoDai_Replace_CongThem", "Tổng dài (Cộng)", "CALCULATE"),
                ("---", "", ""),
                ("AT_TongDienTich_Full", "Tổng DT (Full)", "CALCULATE"),
                ("AT_TongDienTich_Replace", "Tổng DT (Replace)", "CALCULATE"),
                ("AT_TongDienTich_Replace2", "Tổng DT (Replace2)", "CALCULATE"),
                ("AT_TongDienTich_Replace_CongThem", "Tổng DT (Cộng)", "CALCULATE"),
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetEditCommands()
        {
            return new List<(string, string, string)>
            {
                ("AT_TextLink", "Text Link", "ADD"),
                ("AT_DanhSoThuTu", "Đánh số TT", "ADD"),
                ("AT_XoaDoiTuong_CungLayer", "Xóa cùng Layer", "DELETE"),
                ("AT_XoaDoiTuong_3DSolid_Body", "Xóa 3DSolid", "DELETE"),
                ("AT_Offset_2Ben", "Offset 2 bên", "COPY"),
                ("AT_annotive_scale_currentOnly", "Annot. Scale", "SETTINGS"),
                ("---", "", ""),
                ("AT_XoayDoiTuong_TheoViewport", "Xoay VP", "SYNC"),
                ("AT_XoayDoiTuong_Theo2Diem", "Xoay 2 điểm", "SYNC"),
            };
        }

        private static List<(string Command, string Label, string SubIcon)> GetLayoutCommands()
        {
            return new List<(string, string, string)>
            {
                ("AT_TextLayout", "Text Layout", "ADD"),
                ("AT_TaoMoi_TextLayout", "Tạo mới TL", "ADD"),
                ("AT_DimLayout", "Dim Layout", "ADD"),
                ("AT_DimLayout2", "Dim Layout 2", "ADD"),
                ("AT_BlockLayout", "Block Layout", "ADD"),
                ("AT_Label_FromText", "Label Text", "ADD"),
                ("---", "", ""),
                ("AT_UpdateLayout", "Update Layout", "SYNC"),
            };
        }

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
