// 15.CivilToolPalette.cs - Civil Tool Palette (Toolspace Style)
// Thiết kế: PaletteSet dockable, TreeView custom-draw giống Civil 3D Toolspace
// Dark theme chuẩn Autodesk: gradient header, custom TreeView, status bar

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Windows;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.CivilToolPaletteCommands))]

namespace MyFirstProject
{
    public class CivilToolPaletteControl : UserControl
    {
        // ═══ Navy / White Theme ═══
        private static readonly Color BgMain = Color.FromArgb(255, 255, 255);      // Trắng
        private static readonly Color BgTree = Color.FromArgb(255, 255, 255);      // Trắng
        private static readonly Color BgSearch = Color.FromArgb(245, 247, 250);    // Xám rất nhạt
        private static readonly Color BgHeader = Color.FromArgb(27, 42, 74);       // #1B2A4A Navy đậm
        private static readonly Color BgHeaderDark = Color.FromArgb(16, 28, 52);   // Navy tối hơn
        private static readonly Color BgSelected = Color.FromArgb(27, 42, 74);     // Navy — selected
        private static readonly Color BgHover = Color.FromArgb(232, 238, 248);     // Navy rất nhạt — hover
        private static readonly Color BgStatus = Color.FromArgb(27, 42, 74);       // Navy — status bar
        private static readonly Color TextMain = Color.FromArgb(61, 61, 61);       // #3D3D3D Xám đậm
        private static readonly Color TextActive = Color.FromArgb(27, 42, 74);     // Navy — tiêu đề
        private static readonly Color TextDim = Color.FromArgb(122, 122, 122);     // #7A7A7A Xám trung
        private static readonly Color AccentBlue = Color.FromArgb(46, 74, 122);    // #2E4A7A Navy vừa
        private static readonly Color BorderDark = Color.FromArgb(208, 213, 221);  // #D0D5DD Xám nhạt

        // ═══ Category Colors (tông Navy thống nhất) ═══
        private static readonly Color ClrSurface = Color.FromArgb(46, 74, 122);    // Navy
        private static readonly Color ClrAlign = Color.FromArgb(27, 42, 74);       // Navy đậm
        private static readonly Color ClrProfile = Color.FromArgb(46, 74, 122);    // Navy
        private static readonly Color ClrSection = Color.FromArgb(46, 74, 122);    // Navy
        private static readonly Color ClrSample = Color.FromArgb(27, 42, 74);      // Navy đậm
        private static readonly Color ClrCorridor = Color.FromArgb(46, 74, 122);   // Navy
        private static readonly Color ClrPipe = Color.FromArgb(46, 74, 122);       // Navy
        private static readonly Color ClrGrading = Color.FromArgb(27, 42, 74);     // Navy đậm
        private static readonly Color ClrUtils = Color.FromArgb(122, 122, 122);    // Xám

        private TreeView treeView = null!;
        private TextBox searchBox = null!;
        private Panel headerPanel = null!;
        private Panel statusPanel = null!;
        private Label lblStatus = null!;
        private Label lblCount = null!;
        private ImageList imageList = null!;
        private ToolTip toolTip = null!;
        private int _hoveredNodeIndex = -1;

        public CivilToolPaletteControl()
        {
            InitializeComponent();
            LoadCommands();
        }

        private void InitializeComponent()
        {
            this.BackColor = BgMain;
            this.ForeColor = TextMain;
            this.Dock = DockStyle.Fill;

            // ═══ 1. HEADER — Gradient Civil 3D ═══
            headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 42,
                Padding = new Padding(0)
            };
            headerPanel.Paint += (s, e) =>
            {
                using var grad = new LinearGradientBrush(
                    headerPanel.ClientRectangle, BgHeader, BgHeaderDark, 0f);
                e.Graphics.FillRectangle(grad, headerPanel.ClientRectangle);

                // Title
                using var titleFont = new Font("Segoe UI Semibold", 11f);
                e.Graphics.DrawString("  CIVIL TOOL", titleFont, Brushes.White, 4, 10);

                // Version tag
                using var verFont = new Font("Segoe UI", 7.5f);
                using var verBrush = new SolidBrush(Color.FromArgb(200, 210, 230));
                e.Graphics.DrawString("v2.0", verFont, verBrush, headerPanel.Width - 35, 15);
            };

            // Nút Expand/Collapse All
            var btnExpand = new Button
            {
                Text = "⊞",
                Size = new Size(28, 28),
                Location = new Point(0, 7),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10f),
                Cursor = Cursors.Hand
            };
            btnExpand.FlatAppearance.BorderSize = 0;
            btnExpand.FlatAppearance.MouseOverBackColor = Color.FromArgb(46, 74, 122);
            btnExpand.Click += (s, e) =>
            {
                if (treeView.Nodes.Count > 0 && treeView.Nodes[0].IsExpanded)
                    treeView.CollapseAll();
                else
                    treeView.ExpandAll();
            };
            headerPanel.Controls.Add(btnExpand);
            headerPanel.Resize += (s, e) => btnExpand.Location = new Point(headerPanel.Width - 32, 7);

            // ═══ 2. SEARCH BAR ═══
            var searchPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 34,
                BackColor = Color.FromArgb(245, 247, 250),
                Padding = new Padding(6, 4, 6, 4)
            };

            var lblSearchIcon = new Label
            {
                Text = "🔍",
                AutoSize = true,
                Location = new Point(8, 7),
                ForeColor = TextDim,
                Font = new Font("Segoe UI", 9f)
            };
            searchPanel.Controls.Add(lblSearchIcon);

            searchBox = new TextBox
            {
                Location = new Point(30, 5),
                Height = 24,
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                BackColor = BgSearch,
                ForeColor = TextActive,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9f)
            };
            searchBox.TextChanged += SearchBox_TextChanged;
            searchBox.GotFocus += (s, e) => { if (searchBox.ForeColor == TextDim) { searchBox.Text = ""; searchBox.ForeColor = TextActive; } };
            searchPanel.Controls.Add(searchBox);
            searchPanel.Resize += (s, e) => searchBox.Width = searchPanel.Width - 42;

            // ═══ 3. TREEVIEW — Custom Draw ═══
            imageList = new ImageList { ImageSize = new Size(14, 14), ColorDepth = ColorDepth.Depth32Bit };
            CreateCategoryIcons();

            treeView = new TreeView
            {
                Dock = DockStyle.Fill,
                BackColor = BgTree,
                ForeColor = TextMain,
                BorderStyle = BorderStyle.None,
                ShowLines = false,
                ShowPlusMinus = true,
                ShowRootLines = false,
                FullRowSelect = true,
                ItemHeight = 22,
                Font = new Font("Segoe UI", 8.75f),
                ImageList = imageList,
                HotTracking = true,
                DrawMode = TreeViewDrawMode.OwnerDrawAll,
                ShowNodeToolTips = true
            };
            treeView.DrawNode += TreeView_DrawNode;
            treeView.NodeMouseDoubleClick += TreeView_NodeMouseDoubleClick;
            treeView.NodeMouseClick += TreeView_NodeMouseClick;
            treeView.MouseMove += TreeView_MouseMove;
            treeView.MouseLeave += (s, e) => { _hoveredNodeIndex = -1; treeView.Invalidate(); };

            // ═══ 4. STATUS BAR ═══
            statusPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 24,
                BackColor = BgStatus
            };
            statusPanel.Paint += (s, e) =>
            {
                using var grad = new LinearGradientBrush(
                    statusPanel.ClientRectangle, BgStatus, Color.FromArgb(16, 28, 52), 0f);
                e.Graphics.FillRectangle(grad, statusPanel.ClientRectangle);
            };

            lblStatus = new Label
            {
                Text = "Ready",
                AutoSize = true,
                Location = new Point(6, 4),
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 8f)
            };
            statusPanel.Controls.Add(lblStatus);

            lblCount = new Label
            {
                Text = "",
                AutoSize = true,
                ForeColor = Color.FromArgb(200, 210, 230),
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 8f),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            statusPanel.Controls.Add(lblCount);
            statusPanel.Resize += (s, e) => lblCount.Location = new Point(statusPanel.Width - lblCount.Width - 8, 4);

            // ═══ 5. Context Menu ═══
            var ctxMenu = new ContextMenuStrip();
            ctxMenu.BackColor = Color.FromArgb(255, 255, 255);
            ctxMenu.ForeColor = TextMain;
            ctxMenu.Renderer = new DarkToolStripRenderer();

            var mnuRun = new ToolStripMenuItem("▶  Chạy lệnh") { Font = new Font("Segoe UI", 9f) };
            mnuRun.Click += (s, e) => ExecuteSelectedCommand();

            var mnuCopy = new ToolStripMenuItem("📋  Sao chép tên lệnh") { Font = new Font("Segoe UI", 9f) };
            mnuCopy.Click += (s, e) =>
            {
                if (treeView.SelectedNode?.Tag is string cmd)
                    Clipboard.SetText(cmd);
            };

            var mnuExpandAll = new ToolStripMenuItem("⊞  Mở rộng tất cả") { Font = new Font("Segoe UI", 9f) };
            mnuExpandAll.Click += (s, e) => treeView.ExpandAll();

            var mnuCollapseAll = new ToolStripMenuItem("⊟  Thu gọn tất cả") { Font = new Font("Segoe UI", 9f) };
            mnuCollapseAll.Click += (s, e) => treeView.CollapseAll();

            ctxMenu.Items.AddRange(new ToolStripItem[] { mnuRun, mnuCopy, new ToolStripSeparator(), mnuExpandAll, mnuCollapseAll });
            treeView.ContextMenuStrip = ctxMenu;

            // ═══ Assemble ═══
            this.Controls.Add(treeView);
            this.Controls.Add(searchPanel);
            this.Controls.Add(headerPanel);
            this.Controls.Add(statusPanel);

            toolTip = new ToolTip { BackColor = Color.FromArgb(245, 247, 250), ForeColor = TextMain };
        }

        // ═══ Custom TreeView Drawing ═══
        private void TreeView_DrawNode(object? sender, DrawTreeNodeEventArgs e)
        {
            if (e.Node == null || e.Bounds.IsEmpty) return;

            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // Full row bounds
            var rowBounds = new Rectangle(0, e.Bounds.Y, treeView.Width, e.Bounds.Height);
            bool isSelected = e.Node == treeView.SelectedNode;
            bool isHovered = e.Node.Bounds.Y == _hoveredNodeIndex;
            bool isCategory = e.Node.Parent == null;

            // Background
            Color bgColor = isSelected ? BgSelected : (isHovered ? BgHover : BgTree);
            using (var bgBrush = new SolidBrush(bgColor))
                g.FillRectangle(bgBrush, rowBounds);

            // Selected indicator (left bar)
            if (isSelected)
            {
                using var accentBrush = new SolidBrush(AccentBlue);
                g.FillRectangle(accentBrush, new Rectangle(0, e.Bounds.Y, 3, e.Bounds.Height));
            }

            // Indent
            int indent = e.Node.Level * 16 + 18;

            // Expand/Collapse indicator for categories
            if (isCategory && e.Node.Nodes.Count > 0)
            {
                string arrow = e.Node.IsExpanded ? "▾" : "▸";
                using var arrowFont = new Font("Segoe UI", 8f);
                using var arrowBrush = new SolidBrush(TextDim);
                g.DrawString(arrow, arrowFont, arrowBrush, indent - 14, e.Bounds.Y + 4);
            }

            // Icon
            if (e.Node.ImageKey != null && imageList.Images.ContainsKey(e.Node.ImageKey))
            {
                var icon = imageList.Images[e.Node.ImageKey];
                g.DrawImage(icon, indent, e.Bounds.Y + 3, 14, 14);
            }

            // Text
            Color textColor = isCategory ? TextActive : (isSelected ? Color.White : TextMain);
            var font = isCategory ? new Font("Segoe UI Semibold", 8.75f) : new Font("Segoe UI", 8.5f);
            using (font)
            using (var textBrush = new SolidBrush(textColor))
            {
                g.DrawString(e.Node.Text, font, textBrush, indent + 18, e.Bounds.Y + 3);
            }

            // Command count for categories
            if (isCategory)
            {
                string countText = $"({e.Node.Nodes.Count})";
                using var countFont = new Font("Segoe UI", 7f);
                using var countBrush = new SolidBrush(TextDim);
                var textSize = g.MeasureString(e.Node.Text, new Font("Segoe UI Semibold", 8.75f));
                g.DrawString(countText, countFont, countBrush, indent + 20 + textSize.Width, e.Bounds.Y + 5);
            }

            // Bottom separator for categories
            if (isCategory)
            {
                using var sepPen = new Pen(BorderDark);
                g.DrawLine(sepPen, 0, e.Bounds.Bottom - 1, treeView.Width, e.Bounds.Bottom - 1);
            }
        }

        private void TreeView_MouseMove(object? sender, MouseEventArgs e)
        {
            var node = treeView.GetNodeAt(e.Location);
            int newHover = node?.Bounds.Y ?? -1;
            if (newHover != _hoveredNodeIndex)
            {
                _hoveredNodeIndex = newHover;
                treeView.Invalidate();
            }
        }

        // ═══ Category Icons — Color-Coded ═══
        private void CreateCategoryIcons()
        {
            // Surface — TIN mesh triangle
            imageList.Images.Add("surface", DrawIcon(ClrSurface, (g, p) =>
            {
                g.DrawLine(p, 2, 12, 7, 3);
                g.DrawLine(p, 7, 3, 13, 8);
                g.DrawLine(p, 13, 8, 14, 12);
                g.DrawLine(p, 2, 12, 7, 10);
                g.DrawLine(p, 7, 10, 14, 12);
                g.DrawLine(p, 7, 3, 7, 10);
                g.DrawLine(p, 13, 8, 7, 10);
            }));

            // Alignment — Horizontal curve (road centerline)
            imageList.Images.Add("align", DrawIcon(ClrAlign, (g, p) =>
            {
                var pen2 = new Pen(ClrAlign, 2f);
                g.DrawCurve(pen2, new PointF[] {
                    new(1, 12), new(5, 4), new(11, 8), new(15, 3)
                });
                pen2.Dispose();
                // PI marker
                using var redBrush = new SolidBrush(Color.FromArgb(211, 47, 47));
                g.FillEllipse(redBrush, 6, 4, 4, 4);
            }));

            // Profile — Elevation graph
            imageList.Images.Add("profile", DrawIcon(ClrProfile, (g, p) =>
            {
                // Baseline
                g.DrawLine(new Pen(ClrProfile, 1f), 1, 13, 15, 13);
                // Profile line
                var pen2 = new Pen(ClrProfile, 1.8f);
                g.DrawLines(pen2, new PointF[] {
                    new(1, 10), new(4, 6), new(7, 9), new(10, 4), new(14, 7)
                });
                pen2.Dispose();
                // Grid ticks
                var dimPen = new Pen(Color.FromArgb(120, ClrProfile), 0.8f);
                for (int x = 4; x <= 12; x += 4)
                    g.DrawLine(dimPen, x, 12, x, 14);
                dimPen.Dispose();
            }));

            // Section — Cross-section (cut/fill shape)
            imageList.Images.Add("section", DrawIcon(ClrSection, (g, p) =>
            {
                // Natural ground  
                var brownPen = new Pen(Color.FromArgb(155, 100, 50), 1.2f);
                g.DrawLines(brownPen, new PointF[] {
                    new(1, 9), new(4, 6), new(8, 8), new(12, 5), new(15, 7)
                });
                brownPen.Dispose();
                // Design line
                var pen2 = new Pen(ClrSection, 1.8f);
                g.DrawLine(pen2, 3, 8, 13, 8);
                pen2.Dispose();
                // Slopes
                g.DrawLine(p, 3, 8, 1, 12);
                g.DrawLine(p, 13, 8, 15, 12);
            }));

            // Sample Line — Station markers on alignment
            imageList.Images.Add("sample", DrawIcon(ClrSample, (g, p) =>
            {
                // Centerline
                var pen2 = new Pen(ClrSample, 1.5f);
                g.DrawLine(pen2, 1, 8, 15, 8);
                pen2.Dispose();
                // Station ticks
                for (int x = 3; x <= 13; x += 3)
                    g.DrawLine(p, x, 5, x, 11);
                // Main station (thicker)
                g.DrawLine(new Pen(ClrSample, 2f), 8, 3, 8, 13);
            }));

            // Corridor — Road cross-section (trapezoidal)
            imageList.Images.Add("corridor", DrawIcon(ClrCorridor, (g, p) =>
            {
                // Road surface
                g.FillPolygon(new SolidBrush(Color.FromArgb(80, ClrCorridor)),
                    new PointF[] { new(1, 12), new(4, 6), new(12, 6), new(15, 12) });
                g.DrawPolygon(p, new PointF[] { new(1, 12), new(4, 6), new(12, 6), new(15, 12) });
                // Centerline
                var dashPen = new Pen(ClrAlign, 1f) { DashStyle = DashStyle.Dash };
                g.DrawLine(dashPen, 8, 5, 8, 13);
                dashPen.Dispose();
            }));

            // Pipe — Pipe with manhole
            imageList.Images.Add("pipe", DrawIcon(ClrPipe, (g, p) =>
            {
                var pen2 = new Pen(ClrPipe, 2f);
                g.DrawLine(pen2, 1, 9, 5, 8);
                g.DrawLine(pen2, 11, 8, 15, 7);
                pen2.Dispose();
                // Manhole (rectangle)
                g.DrawRectangle(p, 5, 4, 6, 8);
                g.FillRectangle(new SolidBrush(Color.FromArgb(80, ClrPipe)), 5, 4, 6, 8);
            }));

            // Grading — Grid with elevation numbers
            imageList.Images.Add("grading", DrawIcon(ClrGrading, (g, p) =>
            {
                var gridPen = new Pen(ClrGrading, 0.8f);
                for (int i = 0; i <= 3; i++)
                {
                    int pos = 2 + i * 4;
                    g.DrawLine(gridPen, 2, pos, 14, pos);
                    g.DrawLine(gridPen, pos, 2, pos, 14);
                }
                gridPen.Dispose();
                // Elevation dots at intersections
                using var dotBrush = new SolidBrush(Color.FromArgb(60, 60, 60));
                g.FillEllipse(dotBrush, 5, 5, 3, 3);
                g.FillEllipse(dotBrush, 9, 9, 3, 3);
            }));

            // Utils — Gear/cog
            imageList.Images.Add("utils", DrawIcon(ClrUtils, (g, p) =>
            {
                g.DrawEllipse(p, 4, 4, 8, 8);
                g.FillEllipse(new SolidBrush(BgTree), 6, 6, 4, 4);
                for (int i = 0; i < 6; i++)
                {
                    double angle = i * Math.PI / 3;
                    float x1 = 8 + 4 * (float)Math.Cos(angle);
                    float y1 = 8 + 4 * (float)Math.Sin(angle);
                    float x2 = 8 + 6.5f * (float)Math.Cos(angle);
                    float y2 = 8 + 6.5f * (float)Math.Sin(angle);
                    g.DrawLine(new Pen(ClrUtils, 2f), x1, y1, x2, y2);
                }
            }));

            // Command — Small arrow indicator
            imageList.Images.Add("cmd", DrawIcon(AccentBlue, (g, p) =>
            {
                g.FillPolygon(new SolidBrush(AccentBlue),
                    new PointF[] { new(5, 3), new(12, 8), new(5, 13) });
            }));

            // Point — Survey benchmark triangle
            imageList.Images.Add("point", DrawIcon(ClrAlign, (g, p) =>
            {
                g.DrawPolygon(p, new PointF[] { new(8, 2), new(2, 13), new(14, 13) });
                g.FillEllipse(new SolidBrush(ClrAlign), 6, 7, 4, 4);
            }));

            // Volume — Cut/Fill diagram
            var clrVolume = Color.FromArgb(239, 83, 80);
            imageList.Images.Add("volume", DrawIcon(clrVolume, (g, p) =>
            {
                // Cut (red area above line)
                g.FillPolygon(new SolidBrush(Color.FromArgb(100, 220, 53, 69)),
                    new PointF[] { new(1, 8), new(5, 4), new(8, 8) });
                // Fill (green area below line)
                g.FillPolygon(new SolidBrush(Color.FromArgb(100, 40, 167, 69)),
                    new PointF[] { new(8, 8), new(13, 12), new(15, 8) });
                // Design line
                g.DrawLine(new Pen(clrVolume, 1.5f), 1, 8, 15, 8);
            }));
        }

        private Bitmap DrawIcon(Color color, Action<Graphics, Pen> draw)
        {
            var bmp = new Bitmap(16, 16);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            using var pen = new Pen(color, 1.2f);
            draw(g, pen);
            return bmp;
        }

        // ═══ Load Commands ═══
        private void LoadCommands(string filter = "")
        {
            treeView.BeginUpdate();
            treeView.Nodes.Clear();

            var config = MenuSettingsForm.LoadSavedConfig();
            int totalCmds = 0;

            foreach (var group in config.CivilGroups)
            {
                if (group.Commands.Count == 0) continue;

                // Determine icon key from group IconType
                string iconKey = group.IconType?.ToUpper() switch
                {
                    "SURFACE" => "surface",
                    "ALIGN" => "align",
                    "PROFILE" => "profile",
                    "SECTION" => "section",
                    "STATION" => "sample",
                    "CORRIDOR" => "corridor",
                    "PIPE" => "pipe",
                    "GRADING" => "grading",
                    "POINT" => "point",
                    "VOLUME" => "volume",
                    _ => "utils"
                };

                // Clean group name (remove emojis)
                string cleanName = System.Text.RegularExpressions.Regex.Replace(
                    group.GroupName,
                    @"[\u2600-\u27BF\u2B50-\u2B55\u25A0-\u25FF\uFE00-\uFEFF]+|\uD83C[\uDF00-\uDFFF]|\uD83D[\uDC00-\uDFFF]|\uD83E[\uDD00-\uDDFF]",
                    "").Trim();

                var rootNode = new TreeNode(cleanName)
                {
                    ImageKey = iconKey,
                    SelectedImageKey = iconKey
                };

                foreach (var cmd in group.Commands)
                {
                    if (string.IsNullOrEmpty(cmd.Command) || cmd.Command == "---") continue;
                    if (!cmd.IsEnabled) continue;

                    // Filter
                    if (!string.IsNullOrEmpty(filter) &&
                        cmd.Label.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0 &&
                        cmd.Command.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0)
                        continue;

                    var node = new TreeNode(cmd.Label)
                    {
                        Tag = cmd.Command,
                        ImageKey = "cmd",
                        SelectedImageKey = "cmd",
                        ToolTipText = $"Lệnh: {cmd.Command}"
                    };
                    rootNode.Nodes.Add(node);
                    totalCmds++;
                }

                if (rootNode.Nodes.Count > 0)
                    treeView.Nodes.Add(rootNode);
            }

            if (string.IsNullOrEmpty(filter))
                treeView.CollapseAll();
            else
                treeView.ExpandAll();

            treeView.EndUpdate();

            // Update status
            lblCount.Text = $"{totalCmds} lệnh";
            lblCount.Location = new Point(statusPanel.Width - lblCount.Width - 8, 4);
            lblStatus.Text = string.IsNullOrEmpty(filter) ? "Ready" : $"Tìm: \"{filter}\"";
        }

        // ═══ Events ═══
        private void SearchBox_TextChanged(object? sender, EventArgs e)
        {
            LoadCommands(searchBox.Text);
        }

        private void TreeView_NodeMouseDoubleClick(object? sender, TreeNodeMouseClickEventArgs e)
        {
            ExecuteSelectedCommand();
        }

        private void TreeView_NodeMouseClick(object? sender, TreeNodeMouseClickEventArgs e)
        {
            // Toggle expand/collapse on category click
            if (e.Node?.Parent == null && e.Button == MouseButtons.Left)
            {
                if (e.Node.IsExpanded)
                    e.Node.Collapse();
                else
                    e.Node.Expand();
            }
        }

        private void ExecuteSelectedCommand()
        {
            if (treeView.SelectedNode?.Tag is string cmd && !string.IsNullOrEmpty(cmd))
            {
                try
                {
                    lblStatus.Text = $"▶ {cmd}";
                    var doc = AcadApp.DocumentManager.MdiActiveDocument;
                    if (doc != null)
                    {
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        doc.SendStringToExecute($"{cmd} ", true, false, false);
                    }
                }
                catch (System.Exception ex)
                {
                    lblStatus.Text = $"⚠ Lỗi: {ex.Message}";
                }
            }
        }
    }

    // ═══ Dark ToolStrip Renderer ═══
    internal class DarkToolStripRenderer : ToolStripProfessionalRenderer
    {
        public DarkToolStripRenderer() : base(new DarkColorTable()) { }

        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            if (e.Item.Selected)
            {
                using var brush = new SolidBrush(Color.FromArgb(232, 238, 248));
                e.Graphics.FillRectangle(brush, new Rectangle(Point.Empty, e.Item.Size));
                e.Item.ForeColor = Color.FromArgb(27, 42, 74);
            }
            else
            {
                using var brush = new SolidBrush(Color.FromArgb(255, 255, 255));
                e.Graphics.FillRectangle(brush, new Rectangle(Point.Empty, e.Item.Size));
                e.Item.ForeColor = Color.FromArgb(61, 61, 61);
            }
        }

        protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
        {
            using var pen = new Pen(Color.FromArgb(208, 213, 221));
            e.Graphics.DrawLine(pen, 30, e.Item.Height / 2, e.Item.Width - 4, e.Item.Height / 2);
        }
    }

    internal class DarkColorTable : ProfessionalColorTable
    {
        public override Color MenuBorder => Color.FromArgb(208, 213, 221);
        public override Color MenuItemBorder => Color.FromArgb(232, 238, 248);
        public override Color MenuItemSelected => Color.FromArgb(232, 238, 248);
        public override Color MenuStripGradientBegin => Color.FromArgb(255, 255, 255);
        public override Color MenuStripGradientEnd => Color.FromArgb(255, 255, 255);
        public override Color ToolStripDropDownBackground => Color.FromArgb(255, 255, 255);
        public override Color ImageMarginGradientBegin => Color.FromArgb(255, 255, 255);
        public override Color ImageMarginGradientMiddle => Color.FromArgb(255, 255, 255);
        public override Color ImageMarginGradientEnd => Color.FromArgb(255, 255, 255);
        public override Color SeparatorDark => Color.FromArgb(208, 213, 221);
        public override Color SeparatorLight => Color.FromArgb(232, 238, 248);
    }

    // ═══ Palette Commands ═══
    public class CivilToolPaletteCommands
    {
        private static PaletteSet? _ps = null;
        private static CivilToolPaletteControl? _control = null;

        [CommandMethod("CT_Palette")]
        public static void ShowPalette()
        {
            if (_ps == null)
            {
                _ps = new PaletteSet("Civil Tool", new Guid("AA9A60C8-5807-4C04-B473-45B13D5D7260"));
                _control = new CivilToolPaletteControl();
                _ps.Add("Civil Commands", _control);

                _ps.Style = PaletteSetStyles.ShowAutoHideButton |
                           PaletteSetStyles.ShowCloseButton |
                           PaletteSetStyles.ShowPropertiesMenu;
                _ps.MinimumSize = new Size(260, 400);
                _ps.Dock = DockSides.Left;
            }

            _ps.Visible = true;
        }

        [CommandMethod("CT_Show")]
        public static void ShowPaletteAlias() => ShowPalette();
    }
}
