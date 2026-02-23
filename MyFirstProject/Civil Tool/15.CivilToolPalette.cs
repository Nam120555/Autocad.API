
// 15.CivilToolPalette.cs - Palette (Sidebar) giao diện chính cho Civil Tool
// Thiết kế: PaletteSet dockable, chứa TreeView phân nhóm lệnh
// Mục tiêu: "Khoa học, Nhỏ gọn", giống Civil 3D Toolspace

using System;
using System.Drawing;
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
        private TreeView treeView;
        private TextBox searchBox;
        private Label lblSearch;
        private ImageList imageList;

        public CivilToolPaletteControl()
        {
            InitializeComponent();
            LoadCommands();
        }

        private void InitializeComponent()
        {
            this.BackColor = Color.FromArgb(45, 45, 48);
            this.ForeColor = Color.White;
            this.Font = new Font("Segoe UI", 9F);

            // 1. Search Box Panel
            var pnlTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 35,
                Padding = new Padding(5),
                BackColor = Color.FromArgb(30, 30, 30)
            };

            lblSearch = new Label
            {
                Text = "🔍",
                AutoSize = true,
                Location = new Point(5, 8),
                ForeColor = Color.LightGray
            };

            searchBox = new TextBox
            {
                Location = new Point(30, 5),
                Width = 200, // Will anchor
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.FromArgb(60, 60, 65),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            searchBox.TextChanged += SearchBox_TextChanged;

            pnlTop.Controls.Add(lblSearch);
            pnlTop.Controls.Add(searchBox);

            // 2. TreeView
            treeView = new TreeView
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(37, 37, 38),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.None,
                ShowLines = true,
                ShowPlusMinus = true,
                ShowRootLines = true,
                FullRowSelect = true,
                ItemHeight = 22
            };
            
            // ImageList cho icons
            imageList = new ImageList { ImageSize = new Size(16, 16), ColorDepth = ColorDepth.Depth32Bit };
            // Placeholder icons (trong thực tế nên load từ Resources)
            imageList.Images.Add("folder", CreateBitmap(Color.Goldenrod)); // 0
            imageList.Images.Add("cmd", CreateBitmap(Color.SteelBlue));    // 1
            treeView.ImageList = imageList;

            treeView.NodeMouseDoubleClick += TreeView_NodeMouseDoubleClick;

            this.Controls.Add(treeView);
            this.Controls.Add(pnlTop);
        }

        private Bitmap CreateBitmap(Color color)
        {
            Bitmap bmp = new Bitmap(16, 16);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                g.FillEllipse(new SolidBrush(color), 2, 2, 12, 12);
            }
            return bmp;
        }

        private void LoadCommands(string filter = "")
        {
            treeView.BeginUpdate();
            treeView.Nodes.Clear();

            // Root Nodes (Categories)
            var rootSurface = CreateNode("Bề mặt (Surfaces)", "folder");
            var rootAlign = CreateNode("Tuyến (Alignments)", "folder");
            var rootProfile = CreateNode("Trắc dọc (Profiles)", "folder");
            var rootSection = CreateNode("Trắc ngang (Sections)", "folder");
            var rootSample = CreateNode("Cọc (sample Lines)", "folder");
            var rootCorridor = CreateNode("Corridor", "folder");
            var rootPipe = CreateNode("Thoát nước (Pipes)", "folder");
            var rootGrading = CreateNode("San nền (Grading)", "folder");
            var rootUtils = CreateNode("Tiện ích (Utilities)", "folder");

            // Add Commands to Categories (Filter applied here)
            // Surfaces
            AddCmd(rootSurface, "Spot Elevation tại tim", "CTSu_TaoSpotElevation_OnSurface_TaiTim", filter);
            
            // Sample Lines
            AddCmd(rootSample, "Đổi tên Cọc", "CTS_DoiTenCoc", filter);
            AddCmd(rootSample, "Đổi tên Cọc (Đoạn)", "CTS_DoiTenCoc2", filter);
            AddCmd(rootSample, "Phát sinh Cọc Auto", "CTS_PhatSinhCoc", filter);
            AddCmd(rootSample, "Thống kê Cọc (Excel)", "CTSV_ThongKeCoc", filter);

            // Alignments
            AddCmd(rootAlign, "Tạo Trắc dọc", "CTPV_TaoProfileView", filter);
            AddCmd(rootAlign, "Thêm Bảng Lý trình", "CTPV_ThemBang_LyTrinh", filter);

            // Profiles
            AddCmd(rootProfile, "Vẽ Trắc dọc Tự nhiên", "CTP_VeTracDoc_TuNhien", filter);
            AddCmd(rootProfile, "Thống kê Parcel", "CTP_TaoBangThongKeParcel", filter);

            // Sections
            AddCmd(rootSection, "Vẽ Trắc ngang TK", "CTSV_VeTracNgangThietKe", filter);
            AddCmd(rootSection, "Đánh cấp Taluy", "CTSV_DanhCap", filter);
            AddCmd(rootSection, "Xuất Khối lượng Excel", "CTSV_XuatKhoiLuong", filter);

            // Corridors
            AddCmd(rootCorridor, "Thêm tất cả Section", "CTC_AddAllSection", filter);
            AddCmd(rootCorridor, "Tạo Corridor Đô thị", "CTC_TaoCooridor_DuongDoThi_RePhai", filter);

            // Pipes
            AddCmd(rootPipe, "Đổi đường kính Cống", "CTPi_ThayDoi_DuongKinhCong", filter);
            AddCmd(rootPipe, "Thống kê Hố ga", "CTPi_TaoBang_CaoDoDayHoGa", filter); // Assuming rename
            
            // Grading
            AddCmd(rootGrading, "Tạo Lưới San nền", "CTSN_TaoLuoi", filter);
            AddCmd(rootGrading, "Tính Khối lượng SN", "CTSN_TinhKL", filter);

            // Utils
            AddCmd(rootUtils, "Xuất Tọa độ Polyline", "CT_XuatToaDo", filter);
            AddCmd(rootUtils, "Tọa độ Hố ga (DLL)", "CT_VTOADOHG", filter);


            // Add roots to View if they have children
            if (rootSurface.Nodes.Count > 0) treeView.Nodes.Add(rootSurface);
            if (rootAlign.Nodes.Count > 0) treeView.Nodes.Add(rootAlign);
            if (rootProfile.Nodes.Count > 0) treeView.Nodes.Add(rootProfile);
            if (rootSection.Nodes.Count > 0) treeView.Nodes.Add(rootSection);
            if (rootSample.Nodes.Count > 0) treeView.Nodes.Add(rootSample);
            if (rootCorridor.Nodes.Count > 0) treeView.Nodes.Add(rootCorridor);
            if (rootPipe.Nodes.Count > 0) treeView.Nodes.Add(rootPipe);
            if (rootGrading.Nodes.Count > 0) treeView.Nodes.Add(rootGrading);
            if (rootUtils.Nodes.Count > 0) treeView.Nodes.Add(rootUtils);

            treeView.ExpandAll();
            if (string.IsNullOrEmpty(filter)) treeView.CollapseAll(); // Default collapse if no search
            else treeView.ExpandAll(); // Expand if searching

            treeView.EndUpdate();
        }

        private TreeNode CreateNode(string text, string imageKey)
        {
            return new TreeNode(text) { ImageKey = imageKey, SelectedImageKey = imageKey };
        }

        private void AddCmd(TreeNode parent, string text, string command, string filter)
        {
            // Simple filter logic
            if (string.IsNullOrEmpty(filter) || 
                text.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0 ||
                command.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                var node = new TreeNode(text) 
                { 
                    Tag = command,
                    ImageKey = "cmd",
                    SelectedImageKey = "cmd",
                    ToolTipText = command
                };
                parent.Nodes.Add(node);
            }
        }

        private void SearchBox_TextChanged(object sender, EventArgs e)
        {
            LoadCommands(searchBox.Text);
        }

        private void TreeView_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Tag != null)
            {
                string cmd = e.Node.Tag.ToString();
                ExecuteCommand(cmd);
            }
        }

        private void ExecuteCommand(string cmd)
        {
            try
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    // Focus back to AutoCAD before sending command
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    doc.SendStringToExecute($"{cmd} ", true, false, false);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi thực thi lệnh: {ex.Message}");
            }
        }
    }

    public class CivilToolPaletteCommands
    {
        private static PaletteSet? _ps = null;
        private static CivilToolPaletteControl? _control = null;

        [CommandMethod("CT_Palette")]
        public static void ShowPalette()
        {
            if (_ps == null)
            {
                _ps = new PaletteSet("Civil Tool Palette", new Guid("AA9A60C8-5807-4C04-B473-45B13D5D7260"));
                _control = new CivilToolPaletteControl();
                _ps.Add("Civil Commands", _control);
                
                // Docking defaults
                _ps.Style = PaletteSetStyles.ShowAutoHideButton | 
                           PaletteSetStyles.ShowCloseButton |
                           PaletteSetStyles.ShowPropertiesMenu;
                _ps.MinimumSize = new Size(250, 300);
                _ps.Dock = DockSides.Left; // Default dock left like Toolspace
            }
            
            _ps.Visible = true;
        }

        [CommandMethod("CT_Show")]
        public static void ShowPaletteAlias() => ShowPalette();
    }
}
