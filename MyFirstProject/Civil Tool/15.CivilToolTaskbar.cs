// CivilToolTaskbar.cs - Thanh công cụ tổng hợp Civil Tool
// Tổ chức theo nhóm lệnh với dropdown menu

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

// Đăng ký class chứa các lệnh với AutoCAD
[assembly: CommandClass(typeof(MyFirstProject.CivilToolTaskbarCommands))]

namespace MyFirstProject
{
    /// <summary>
    /// Thanh công cụ tổng hợp tất cả các lệnh Civil Tool
    /// </summary>
    public class CivilToolTaskbar : Form
    {
        private static CivilToolTaskbar? instance;
        private Point _dragStart;
        private bool _dragging = false;

        public CivilToolTaskbar()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Civil Tool";
            this.Size = new Size(900, 48);
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.Manual;
            this.TopMost = true;
            this.BackColor = Color.FromArgb(255, 255, 255);
            this.ForeColor = Color.FromArgb(61, 61, 61);
            this.DoubleBuffered = true;

            // Rounded corners (Windows 11)
            this.Region = CreateRoundedRegion(this.Width, this.Height, 8);

            // Drag support
            this.MouseDown += (s, e) => { _dragging = true; _dragStart = e.Location; };
            this.MouseMove += (s, e) => { if (_dragging) { this.Left += e.X - _dragStart.X; this.Top += e.Y - _dragStart.Y; } };
            this.MouseUp += (s, e) => _dragging = false;

            // Custom paint for gradient background
            this.Paint += (s, e) =>
            {
                using var grad = new System.Drawing.Drawing2D.LinearGradientBrush(
                    this.ClientRectangle,
                    Color.FromArgb(27, 42, 74),
                    Color.FromArgb(16, 28, 52), 90f);
                e.Graphics.FillRectangle(grad, this.ClientRectangle);

                // Bottom accent line
                using var accentPen = new Pen(Color.FromArgb(46, 74, 122), 2);
                e.Graphics.DrawLine(accentPen, 0, this.Height - 1, this.Width, this.Height - 1);
            };

            // Đặt ở giữa phía trên màn hình
            var screenBounds = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
            this.Location = new Point((screenBounds.Width - this.Width) / 2, 5);

            // Load config đã lưu
            var config = MenuSettingsForm.LoadSavedConfig();

            int btnWidth = 68;
            int btnHeight = 28;
            int margin = 2;
            int x = 6;
            int y = 10;

            // Tạo nút từ Civil groups
            foreach (var group in config.CivilGroups)
            {
                if (group.Commands.Count == 0) continue;
                var cmds = group.Commands
                    .Where(c => !string.IsNullOrEmpty(c.Command) && c.Command != "---")
                    .Select(c => (c.Label, c.Command))
                    .ToList();
                if (cmds.Count == 0) continue;

                string cleanName = System.Text.RegularExpressions.Regex.Replace(
                    group.GroupName, @"[\u2600-\u27BF\u2B50-\u2B55\u25A0-\u25FF\uFE00-\uFEFF]+|\uD83C[\uDF00-\uDFFF]|\uD83D[\uDC00-\uDFFF]|\uD83E[\uDD00-\uDDFF]", "").Trim();

                this.Controls.Add(CreateDropdownButton(cleanName, x, y, btnWidth, btnHeight, cmds));
                x += btnWidth + margin;
            }

            // Nút Cài đặt
            var btnSettings = CreateButton("⚙", x, y, 30, btnHeight, Color.FromArgb(27, 42, 74));
            btnSettings.ForeColor = Color.White;
            btnSettings.Click += (s, e) => { this.Hide(); RunCommand("CT_MENU_SETTINGS"); this.Show(); };
            this.Controls.Add(btnSettings);
            x += 30 + margin;

            // Close Button
            var btnClose = CreateButton("✕", x, y, 26, btnHeight, Color.FromArgb(46, 74, 122));
            btnClose.ForeColor = Color.White;
            btnClose.FlatAppearance.MouseOverBackColor = Color.FromArgb(211, 47, 47);
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);

            // Cập nhật kích thước form
            this.Size = new Size(x + 30, 48);
            this.Region = CreateRoundedRegion(this.Width, this.Height, 8);
            this.Location = new Point((screenBounds.Width - this.Width) / 2, 5);
        }

        private static System.Drawing.Region CreateRoundedRegion(int w, int h, int r)
        {
            var path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddArc(0, 0, r * 2, r * 2, 180, 90);
            path.AddArc(w - r * 2, 0, r * 2, r * 2, 270, 90);
            path.AddArc(w - r * 2, h - r * 2, r * 2, r * 2, 0, 90);
            path.AddArc(0, h - r * 2, r * 2, r * 2, 90, 90);
            path.CloseFigure();
            return new System.Drawing.Region(path);
        }

        private Button CreateButton(string text, int x, int y, int width, int height, Color bgColor)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(width, height),
                FlatStyle = FlatStyle.Flat,
                BackColor = bgColor,
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 8f)
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(46, 74, 122);
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(46, 74, 122);
            btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(16, 28, 52);
            return btn;
        }

        private Button CreateDropdownButton(string text, int x, int y, int width, int height,
            List<(string Name, string Command)> commands)
        {
            var btn = CreateButton(text + " ▾", x, y, width, height, Color.FromArgb(27, 42, 74));
            btn.Click += (s, e) =>
            {
                var menu = new ContextMenuStrip();
                menu.BackColor = Color.FromArgb(255, 255, 255);
                menu.ForeColor = Color.FromArgb(61, 61, 61);
                menu.Font = new Font("Segoe UI", 9);
                menu.Renderer = new DarkToolStripRenderer();
                menu.ShowImageMargin = false;

                foreach (var cmd in commands)
                {
                    if (string.IsNullOrEmpty(cmd.Command))
                    {
                        menu.Items.Add(new ToolStripSeparator());
                        continue;
                    }

                    var item = new ToolStripMenuItem($"  {cmd.Name}");
                    item.ForeColor = Color.FromArgb(61, 61, 61);
                    string cmdName = cmd.Command;
                    item.Click += (sender, args) => { this.Hide(); RunCommand(cmdName); this.Show(); };
                    menu.Items.Add(item);
                }

                menu.Show(btn, new Point(0, btn.Height));
            };
            return btn;
        }

        private static void RunCommand(string commandName)
        {
            try
            {
                Document doc = AcadApp.DocumentManager.MdiActiveDocument;
                doc.SendStringToExecute($"{commandName}\n", true, false, false);
            }
            catch { }
        }

        #region Static Methods

        public static void ShowTaskbar()
        {
            if (instance == null || instance.IsDisposed)
            {
                instance = new CivilToolTaskbar();
            }
            instance.Show();
            instance.BringToFront();
        }

        public static void CloseTaskbar()
        {
            instance?.Close();
            instance = null;
        }

        #endregion
    }

    /// <summary>
    /// Form hiển thị danh sách tất cả các lệnh Civil Tool
    /// </summary>
    public class CommandListForm : Form
    {
        private static CommandListForm? formInstance;
        private ListView listView = null!;
        private TextBox searchBox = null!;
        private List<(string Group, string Command, string Description)> allCommands = null!;

        public CommandListForm()
        {
            InitializeComponent();
            LoadCommands();
        }

        private void InitializeComponent()
        {
            this.Text = "📋 Danh Sách Lệnh Civil Tool";
            this.Size = new Size(750, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(255, 255, 255);
            this.ForeColor = Color.FromArgb(61, 61, 61);

            // Search box
            var lblSearch = new Label
            {
                Text = "🔍 Tìm kiếm:",
                Location = new Point(10, 12),
                Size = new Size(80, 25),
                ForeColor = Color.FromArgb(27, 42, 74)
            };
            this.Controls.Add(lblSearch);

            searchBox = new TextBox
            {
                Location = new Point(95, 10),
                Size = new Size(300, 25),
                BackColor = Color.FromArgb(245, 247, 250),
                ForeColor = Color.FromArgb(61, 61, 61),
                BorderStyle = BorderStyle.FixedSingle
            };
            searchBox.TextChanged += SearchBox_TextChanged;
            this.Controls.Add(searchBox);

            // Button chạy lệnh
            var btnRun = new Button
            {
                Text = "▶ Chạy lệnh",
                Location = new Point(550, 8),
                Size = new Size(90, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(27, 42, 74),
                ForeColor = Color.White
            };
            btnRun.Click += BtnRun_Click;
            this.Controls.Add(btnRun);

            // Button đóng
            var btnClose = new Button
            {
                Text = "✕ Đóng",
                Location = new Point(645, 8),
                Size = new Size(80, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(245, 247, 250),
                ForeColor = Color.FromArgb(61, 61, 61)
            };
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);

            // ListView
            listView = new ListView
            {
                Location = new Point(10, 45),
                Size = new Size(715, 500),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                BackColor = Color.FromArgb(255, 255, 255),
                ForeColor = Color.FromArgb(61, 61, 61)
            };
            listView.Columns.Add("STT", 45);
            listView.Columns.Add("Nhóm", 100);
            listView.Columns.Add("Lệnh", 220);
            listView.Columns.Add("Mô tả", 330);
            listView.DoubleClick += ListView_DoubleClick;
            this.Controls.Add(listView);

            this.Resize += (s, e) =>
            {
                listView.Size = new Size(this.ClientSize.Width - 20, this.ClientSize.Height - 55);
            };
        }

        private void LoadCommands()
        {
            allCommands = new List<(string, string, string)>
            {
                // Corridor
                ("Corridor", "CTC_AddAllSection", "Thêm tất cả section vào corridor"),
                ("Corridor", "CTC_TaoCooridor_DuongDoThi_RePhai", "Tạo corridor đường đô thị rẽ phải"),
                
                // Parcel
                ("Parcel", "CTPA_TaoParcel_CacLoaiNha", "Tạo parcel từ polyline các loại nhà"),
                
                // Pipe & Structures
                ("Pipe", "CTPI_ThayDoi_DuongKinhCong", "Thay đổi đường kính ống cống"),
                ("Pipe", "CTPI_ThayDoi_MatPhangRef_Cong", "Đặt mặt phẳng reference cho cống"),
                ("Pipe", "CTPI_ThayDoi_DoanDocCong", "Thay đổi độ dốc ống cống"),
                ("Pipe", "CTPI_BangCaoDo_TuNhienHoThu", "Bảng cao độ tự nhiên hố thu"),
                ("Pipe", "CTPI_XoayHoThu_Theo2diem", "Xoay hố thu theo 2 điểm"),
                
                // Point
                ("Point", "CTPO_TaoCogoPoint_CaoDo_FromSurface", "Tạo CogoPoint lấy cao độ từ Surface"),
                ("Point", "CTPO_TaoCogoPoint_CaoDo_Elevationspot", "Tạo CogoPoint từ Spot Elevation"),
                ("Point", "CTPO_CreateCogopointFromText", "Tạo CogoPoint từ Text cao độ"),
                ("Point", "CTPO_An_CogoPoint", "Ẩn các điểm CogoPoint"),
                ("Point", "CTPO_UpdateAllPointGroup", "Cập nhật tất cả Point Group"),
                ("Point", "CTPO_CPTT", "Chuyển Point thành Text cao độ"),
        ("Point", "CTPO_TaoPointTuExcel", "Import tọa độ từ Excel (Block + Name)"),
                
                // Profile
                ("Profile", "CTP_VeTracDoc_TuNhien", "Vẽ trắc dọc địa hình tự nhiên"),
                ("Profile", "CTP_VeTracDoc_TuNhien_TatCaTuyen", "Vẽ trắc dọc TN tất cả tuyến"),
                ("Profile", "CTP_Fix_DuongTuNhien_TheoCoc", "Fix đường TN theo cọc"),
                ("Profile", "CTP_GanNhanNutGiao_LenTracDoc", "Gán nhãn nút giao lên trắc dọc"),
                ("Profile", "CTP_TaoCogoPointTuPVI", "Tạo CogoPoint từ PVI"),
                ("Profile", "CTP_TaoBangThongKeParcel", "Tạo bảng thống kê parcel"),
                
                // Sample Line / Cọc  
                ("Cọc", "CTS_DoiTenCoc", "Đổi tên cọc"),
                ("Cọc", "CTS_DoiTenCoc2", "Đổi tên cọc theo đoạn"),
                ("Cọc", "CTS_DoiTenCoc3", "Đổi tên cọc theo Km"),
                ("Cọc", "CTS_TaoBang_ToaDoCoc", "Tạo bảng tọa độ cọc (X,Y)"),
                ("Cọc", "CTS_PhatSinhCoc", "Phát sinh cọc tự động"),
                ("Cọc", "CTS_PhatSinhCoc_ChiTiet", "Phát sinh cọc chi tiết"),
                ("Cọc", "CTS_DichCoc_TinhTien", "Dịch cọc tịnh tiến"),
                ("Cọc", "CTS_Copy_NhomCoc", "Sao chép nhóm cọc"),
                ("Cọc", "CTS_DongBo_2_NhomCoc", "Đồng bộ 2 nhóm cọc"),
                
                // Section View
                ("Trắc ngang", "CTSV_VeTracNgangThietKe", "Vẽ trắc ngang thiết kế"),
                ("Trắc ngang", "CTSV_DanhCap", "Đánh cấp VHC"),
                ("Trắc ngang", "CTSV_DanhCap_XoaBo", "Xóa bỏ đánh cấp"),
                ("Trắc ngang", "CTSV_DanhCap_VeThem", "Vẽ thêm đánh cấp"),
                ("Trắc ngang", "CTSV_ThemVatLieu_TrenCatNgang", "Điền vật liệu trên cắt ngang"),
                ("Trắc ngang", "CTSV_fit_KhungIn", "Fit khung in trắc ngang"),
                ("Trắc ngang", "CTSV_An_DuongDiaChat", "Ẩn đường địa chất"),
                
                // Khối lượng
                ("Khối lượng", "CTSV_Taskbar", "Mở taskbar khối lượng"),
                ("Khối lượng", "CTSV_XuatKhoiLuong", "Xuất khối lượng ra Excel"),
                ("Khối lượng", "CTSV_XuatCad", "Xuất khối lượng ra CAD"),
                ("Khối lượng", "CTSV_CaiDatBang", "Cài đặt bảng khối lượng"),
                ("Khối lượng", "CTSV_ThongKeCoc", "Thống kê cọc ra Excel"),
                ("Khối lượng", "CTSV_ThongKeCoc_TatCa", "Thống kê tất cả cọc"),
                
                // San nền
                ("San nền", "CTSN_Taskbar", "Mở taskbar san nền"),
                ("San nền", "CTSN_TaoLuoi", "Tạo lưới san nền"),
                ("San nền", "CTSN_NhapCaoDo", "Nhập cao độ lưới"),
                ("San nền", "CTSN_Surface", "Lấy cao độ từ Surface"),
                ("San nền", "CTSN_TinhKL", "Tính khối lượng san nền"),
                ("San nền", "CTSN_XuatBang", "Xuất bảng khối lượng CAD"),
                
                // Taskbar
                ("Tool", "CT_Taskbar", "Mở thanh công cụ Civil Tool"),
                ("Tool", "CT", "Mở thanh công cụ (alias)"),
                ("Tool", "CT_VTOADOHG", "Tọa độ hố ga (DLL ngoài)"),
            };

            RefreshListView();
        }

        private void RefreshListView(string filter = "")
        {
            listView.Items.Clear();
            int stt = 0;
            foreach (var cmd in allCommands)
            {
                if (string.IsNullOrEmpty(filter) ||
                    cmd.Command.ToLower().Contains(filter.ToLower()) ||
                    cmd.Description.ToLower().Contains(filter.ToLower()) ||
                    cmd.Group.ToLower().Contains(filter.ToLower()))
                {
                    stt++;
                    var item = new ListViewItem(stt.ToString());
                    item.SubItems.Add(cmd.Group);
                    item.SubItems.Add(cmd.Command);
                    item.SubItems.Add(cmd.Description);
                    item.Tag = cmd.Command;
                    listView.Items.Add(item);
                }
            }
        }

        private void SearchBox_TextChanged(object? sender, EventArgs e)
        {
            RefreshListView(searchBox.Text);
        }

        private void BtnRun_Click(object? sender, EventArgs e)
        {
            RunSelectedCommand();
        }

        private void ListView_DoubleClick(object? sender, EventArgs e)
        {
            RunSelectedCommand();
        }

        private void RunSelectedCommand()
        {
            if (listView.SelectedItems.Count > 0)
            {
                var cmdName = listView.SelectedItems[0].Tag?.ToString();
                if (!string.IsNullOrEmpty(cmdName))
                {
                    this.Hide();
                    try
                    {
                        Document doc = AcadApp.DocumentManager.MdiActiveDocument;
                        doc.SendStringToExecute($"{cmdName}\n", true, false, false);
                    }
                    catch { }
                    this.Show();
                }
            }
        }

        public static void ShowForm()
        {
            if (formInstance == null || formInstance.IsDisposed)
            {
                formInstance = new CommandListForm();
            }
            formInstance.Show();
            formInstance.BringToFront();
        }
    }

    /// <summary>
    /// Lệnh mở thanh công cụ Civil Tool
    /// </summary>
    public class CivilToolTaskbarCommands
    {
        [CommandMethod("CT_Taskbar")]
        public static void CTTaskbar()
        {
            CivilToolTaskbar.ShowTaskbar();
        }

        [CommandMethod("TASKBAR")]
        public static void CTMenu()
        {
            CivilToolTaskbar.ShowTaskbar();
        }

        [CommandMethod("CT")]
        public static void CTCmd()
        {
            CivilToolTaskbar.ShowTaskbar();
        }

        [CommandMethod("CT_DanhSachLenh")]
        public static void CTDanhSachLenh()
        {
            CommandListForm.ShowForm();
        }
    }
}
