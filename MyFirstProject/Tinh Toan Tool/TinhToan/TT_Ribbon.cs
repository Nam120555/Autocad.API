using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;

using AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadException = Autodesk.AutoCAD.Runtime.Exception;

[assembly: CommandClass(typeof(TinhToan.TT_RibbonCommands))]

namespace TinhToan
{
    public class TT_RibbonCommands
    {
        [CommandMethod("CIVIL_TINHTOAN_RIBBON")]
        public static void ShowCalcRibbon()
        {
            var ribbon = ComponentManager.Ribbon;
            if (ribbon == null)
            {
                var doc = AcadApplication.DocumentManager.MdiActiveDocument;
                doc?.SendStringToExecute("RIBBON ", true, false, false);
                ribbon = ComponentManager.Ribbon;
                if (ribbon == null) return;
            }

            // Xoá tab cũ nếu có để tạo lại
            var existing = ribbon.Tabs.FirstOrDefault(t => t.Id == "CivilTool.CalcTab");
            if (existing != null) ribbon.Tabs.Remove(existing);

            RibbonTab tab = new RibbonTab
            {
                Title = "TÍNH TOÁN CIVIL N",
                Id = "CivilTool.CalcTab"
            };
            ribbon.Tabs.Add(tab);

            // 1. PANEL NÚT GIAO & ĐẢO
            CreateCalcPanel(tab, "Thiết Kế Nút Giao", "SURFACE", new List<(string Command, string Label, string SubIcon)> {
                ("TT_VongXuyen", "⭐ Tính Nút Giao Vòng Xuyến", "CALCULATE"),
                ("TT_TamNhinNutGiao", "Kiểm tra Tầm Nhìn Nút Giao (TCVN 4054)", "VIEW"),
                ("---", "---", ""),
                ("TT_TinhToanBanKinhRe", "Bo Vuốt Bán Kính Rẽ Phải", "ALIGN")
            });

            // 2. PANEL LÀN ĐƯỜNG & CHUYỂN TỐC
            CreateCalcPanel(tab, "Phân Làn & Chuyển Tốc", "ALIGN", new List<(string Command, string Label, string SubIcon)> {
                ("TT_TangGiamToc", "⭐ Tính Làn Tăng / Giảm Tốc", "CORRIDOR"),
                ("TT_NhapDong", "Tính Chiều Dài Nhập Dòng", "CORRIDOR"),
                ("---", "---", ""),
                ("TT_LanReTrai", "Thiết Kế Làn Rẽ Trái", "ALIGN"),
                ("TT_LanRePhai", "Thiết Kế Làn Rẽ Phải", "ALIGN"),
                ("TT_MoRongXeChay", "Mở Rộng Bề Rộng Góc Cua", "WIDENING") // Using correct command name
            });


            tab.IsActive = true;
            
            var ed = AcadApplication.DocumentManager.MdiActiveDocument?.Editor;
            ed?.WriteMessage("\n[Civil Tool] Đã tạo thành công Menu Ribbon 'TÍNH TOÁN CIVIL N'.");
        }

        private static void CreateCalcPanel(RibbonTab tab, string title, string iconType, List<(string Command, string Label, string SubIcon)> commands)
        {
            RibbonPanelSource panelSource = new RibbonPanelSource { Title = title };
            RibbonPanel panel = new RibbonPanel { Source = panelSource };

            var realCmds = commands.Where(c => c.Command != "---").ToList();
            if (realCmds.Count == 0) return;

            var featured = realCmds.FirstOrDefault(c => c.Label.Contains("⭐"));
            if (string.IsNullOrEmpty(featured.Command)) featured = realCmds[0];

            var remainingCmds = realCmds.Where(c => c.Command != featured.Command).ToList();

            // NÚT LỚN
            RibbonButton largeBtn = new RibbonButton
            {
                Text = featured.Label.Replace("⭐ ", ""),
                ShowText = true,
                ShowImage = true,
                Image = MyFirstProject.CivilRibbonTaskbar.CreateCivilIcon(featured.SubIcon, 16),
                LargeImage = MyFirstProject.CivilRibbonTaskbar.CreateCivilIcon(featured.SubIcon, 32),
                Size = RibbonItemSize.Large,
                Orientation = System.Windows.Controls.Orientation.Vertical,
                CommandHandler = new MyFirstProject.CivilRibbonCommandHandler(),
                Tag = featured.Command,
                CommandParameter = featured.Command
            };
            panelSource.Items.Add(largeBtn);

            // DROPDOWN (Nếu có các lệnh phụ)
            if (remainingCmds.Count > 0)
            {
                RibbonSplitButton splitBtn = new RibbonSplitButton
                {
                    Text = "Công Cụ Khác",
                    ShowText = true,
                    ShowImage = true,
                    Image = MyFirstProject.CivilRibbonTaskbar.CreateCivilIcon(iconType, 16),
                    LargeImage = MyFirstProject.CivilRibbonTaskbar.CreateCivilIcon(iconType, 32),
                    Size = RibbonItemSize.Large,
                    Orientation = System.Windows.Controls.Orientation.Vertical,
                    ListStyle = RibbonSplitButtonListStyle.List,
                    IsSplit = false
                };

                foreach (var cmd in commands)
                {
                    if (cmd.Command == "---") { splitBtn.Items.Add(new RibbonSeparator()); continue; }
                    if (cmd.Command == featured.Command) continue;

                    RibbonButton item = new RibbonButton
                    {
                        Text = cmd.Label.Replace("⭐ ", ""),
                        ShowText = true,
                        ShowImage = true,
                        Image = MyFirstProject.CivilRibbonTaskbar.CreateCivilIcon(cmd.SubIcon, 16),
                        LargeImage = MyFirstProject.CivilRibbonTaskbar.CreateCivilIcon(cmd.SubIcon, 32),
                        Size = RibbonItemSize.Standard,
                        Orientation = System.Windows.Controls.Orientation.Horizontal,
                        CommandHandler = new MyFirstProject.CivilRibbonCommandHandler(),
                        Tag = cmd.Command,
                        CommandParameter = cmd.Command
                    };
                    splitBtn.Items.Add(item);
                }
                panelSource.Items.Add(splitBtn);
            }

            tab.Panels.Add(panel);
        }
    }
}
