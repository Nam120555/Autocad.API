// CivilToolTaskbar.cs - Thanh công cụ tổng hợp Civil Tool
// Tổ chức theo nhóm lệnh với dropdown menu

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

// Đăng ký class chứa các lệnh với AutoCAD
[assembly: CommandClass(typeof(Civil3DCsharp.CivilToolTaskbarCommands))]

namespace Civil3DCsharp
{
    /// <summary>
    /// Thanh công cụ tổng hợp tất cả các lệnh Civil Tool
    /// </summary>
    public class CivilToolTaskbar : Form
    {
        private static CivilToolTaskbar? instance;

        public CivilToolTaskbar()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "🛠 Civil Tool - Thanh Công Cụ";
            this.Size = new Size(900, 85);
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.StartPosition = FormStartPosition.Manual;
            this.TopMost = true;
            this.BackColor = Color.FromArgb(45, 45, 48);
            this.ForeColor = Color.White;

            // Đặt ở giữa phía trên màn hình
            var screenBounds = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
            this.Location = new Point((screenBounds.Width - this.Width) / 2, 10);

            int btnWidth = 85;
            int btnHeight = 50;
            int margin = 3;
            int x = margin;
            int y = 8;

            // 1. Bề mặt (Surface)
            this.Controls.Add(CreateDropdownButton("🗺️", "Bề mặt", x, y, btnWidth, btnHeight, 
                Color.FromArgb(0, 120, 215), GetSurfaceCommands()));
            x += btnWidth + margin;

            // 2. Cọc (SampleLine)
            this.Controls.Add(CreateDropdownButton("📍", "Cọc", x, y, btnWidth, btnHeight,
                Color.FromArgb(16, 124, 16), GetSampleLineCommands()));
            x += btnWidth + margin;

            // 3. Tuyến (Alignment)
            this.Controls.Add(CreateDropdownButton("🛣️", "Tuyến", x, y, btnWidth, btnHeight,
                Color.FromArgb(202, 80, 16), GetAlignmentCommands()));
            x += btnWidth + margin;

            // 4. Trắc dọc (Profile)
            this.Controls.Add(CreateDropdownButton("📈", "Trắc dọc", x, y, btnWidth, btnHeight,
                Color.FromArgb(0, 153, 188), GetProfileCommands()));
            x += btnWidth + margin;

            // 5. Corridor
            this.Controls.Add(CreateDropdownButton("🛤️", "Corridor", x, y, btnWidth, btnHeight,
                Color.FromArgb(107, 105, 103), GetCorridorCommands()));
            x += btnWidth + margin;

            // 6. Trắc ngang (Section)
            this.Controls.Add(CreateDropdownButton("📉", "Trắc ngang", x, y, btnWidth, btnHeight,
                Color.FromArgb(100, 150, 60), GetSectionViewCommands()));
            x += btnWidth + margin;

            // 7. Nút giao (Intersection)
            this.Controls.Add(CreateDropdownButton("➕", "Nút giao", x, y, btnWidth, btnHeight,
                Color.FromArgb(114, 50, 150), GetIntersectionCommands()));
            x += btnWidth + margin;

            // 8. San nền (Grading)
            this.Controls.Add(CreateDropdownButton("▦", "San nền", x, y, btnWidth, btnHeight,
                Color.FromArgb(0, 100, 100), GetSanNenCommands()));
            x += btnWidth + margin;

            // 9. Khung in (Plan)
            this.Controls.Add(CreateDropdownButton("📋", "Khung in", x, y, btnWidth, btnHeight,
                Color.FromArgb(128, 128, 0), GetPlanCommands()));
            x += btnWidth + margin;

            // 10. Danh sách lệnh - Mở form xem tất cả lệnh
            var btnCommandList = CreateButton("📋", "Lệnh", x, y, 55, btnHeight, Color.FromArgb(100, 80, 150));
            btnCommandList.Click += (s, e) => ShowCommandListForm();
            this.Controls.Add(btnCommandList);
            x += 55 + margin;

            // 11. Công cụ ngoài (External Tools)
            this.Controls.Add(CreateDropdownButton("🔧", "Công cụ", x, y, btnWidth, btnHeight,
                Color.FromArgb(80, 80, 85), GetExternalToolsCommands()));
            x += btnWidth + margin;

            // Close Button
            var btnClose = CreateButton("✕", "", x, y, 35, btnHeight, Color.FromArgb(150, 50, 50));
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);

            // Cập nhật kích thước form dựa trên số lượng nút
            this.Size = new Size(x + 45, 85);
            this.Location = new Point((screenBounds.Width - this.Width) / 2, 10);
        }

        private Button CreateButton(string icon, string text, int x, int y, int width, int height, Color bgColor)
        {
            var btn = new Button
            {
                Text = string.IsNullOrEmpty(text) ? icon : $"{icon}\n{text}",
                Location = new Point(x, y),
                Size = new Size(width, height),
                FlatStyle = FlatStyle.Flat,
                BackColor = bgColor,
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 85);
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(
                Math.Min(bgColor.R + 30, 255),
                Math.Min(bgColor.G + 30, 255),
                Math.Min(bgColor.B + 30, 255));
            return btn;
        }

        private Button CreateDropdownButton(string icon, string text, int x, int y, int width, int height, 
            Color bgColor, List<(string Name, string Command)> commands)
        {
            var btn = CreateButton(icon, text, x, y, width, height, bgColor);
            btn.Click += (s, e) =>
            {
                var menu = new ContextMenuStrip();
                menu.BackColor = Color.FromArgb(45, 45, 48);
                menu.ForeColor = Color.White;
                menu.Font = new Font("Segoe UI", 9);

                foreach (var cmd in commands)
                {
                    var item = new ToolStripMenuItem(cmd.Name);
                    item.BackColor = Color.FromArgb(55, 55, 58);
                    item.ForeColor = Color.White;
                    string cmdName = cmd.Command;
                    item.Click += (sender, args) =>
                    {
                        this.Hide();
                        RunCommand(cmdName);
                        this.Show();
                    };
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

        #region Command Lists

        private List<(string Name, string Command)> GetSurfaceCommands()
        {
            return new List<(string, string)>
            {
                ("📏 Spot Elevation tại tim", "CTS_TaoSpotElevation_OnSurface_TaiTim"),
                ("─────────────", ""),
                ("➕ Tạo Point từ bảng", "CTPo_TaoPointTheoBang"),
                ("🔄 Point → Block", "CTPo_ChuyenPointThanhBlock"),
                ("📋 Bảng thống kê Point", "CTPo_TaoBangThongKePoint"),
                ("✏ Thay đổi cao độ", "CTPo_ThayDoiCaoDo"),
                ("🏷 Đặt tên thứ tự", "CTPo_DatTen_theoThuTu"),
                ("🎨 Thay đổi Style", "CTPo_ThayDoiStyle"),
                ("ℹ Lấy thông tin", "CTPo_LayThongTin"),
            };
        }

        private List<(string Name, string Command)> GetSampleLineCommands()
        {
            return new List<(string, string)>
            {
                ("✏ Đổi tên cọc", "CTS_DoiTenCoc"),
                ("✏ Đổi tên cọc đoạn", "CTS_DoiTenCoc2"),
                ("✏ Đổi tên cọc Km", "CTS_DoiTenCoc3"),
                ("✏ Đổi tên từ CogoPoint", "CTS_DoiTenCoc_fromCogoPoint"),
                ("✏ Đổi tên thứ tự", "CTS_DoiTenCoc_TheoThuTu"),
                ("✏ Đổi tên hậu tố A", "CTS_DoiTenCoc_H"),
                ("─────────────", ""),
                ("📐 Tọa độ cọc (X,Y)", "CTS_TaoBang_ToaDoCoc"),
                ("📐 Tọa độ cọc (Lý trình)", "CTS_TaoBang_ToaDoCoc2"),
                ("📐 Tọa độ cọc (Cao độ)", "CTS_TaoBang_ToaDoCoc3"),
                ("🔄 Cập nhật từ bảng", "AT_UPdate2Table"),
                ("─────────────", ""),
                ("➕ Chèn trên trắc dọc", "CTS_ChenCoc_TrenTracDoc"),
                ("➕ Chèn trên trắc ngang", "CTS_CHENCOC_TRENTRACNGANG"),
                ("➕ Phát sinh cọc auto", "CTS_PhatSinhCoc"),
                ("➕ Phát sinh chi tiết", "CTS_PhatSinhCoc_ChiTiet"),
                ("➕ Phát sinh delta", "CTS_PhatSinhCoc_theoKhoangDelta"),
                ("➕ Phát sinh từ CogoPoint", "CTS_PhatSinhCoc_TuCogoPoint"),
                ("➕ Phát sinh từ bảng", "CTS_PhatSinhCoc_TheoBang"),
                ("─────────────", ""),
                ("↔ Dịch cọc tịnh tiến", "CTS_DichCoc_TinhTien"),
                ("↔ Dịch cọc 40m", "CTS_DichCoc_TinhTien40"),
                ("↔ Dịch cọc 20m", "CTS_DichCoc_TinhTien_20"),
                ("📋 Sao chép nhóm cọc", "CTS_Copy_NhomCoc"),
                ("🔄 Đồng bộ nhóm cọc", "CTS_DongBo_2_NhomCoc"),
                ("🔄 Đồng bộ theo đoạn", "CTS_DongBo_2_NhomCoc_TheoDoan"),
                ("─────────────", ""),
                ("📏 Copy bề rộng SL", "CTS_Copy_BeRong_sampleLine"),
                ("📏 Thay đổi bề rộng SL", "CTS_Thaydoi_BeRong_sampleLine"),
                ("📏 Offset bề rộng SL", "CTS_Offset_BeRong_sampleLine"),
                ("─────────────", ""),
                ("📊 Thống kê cọc (Excel)", "CTSV_ThongKeCoc"),
                ("📊 Thống kê tất cả cọc", "CTSV_ThongKeCoc_TatCa")
            };
        }

        private List<(string Name, string Command)> GetAlignmentCommands()
        {
            return new List<(string, string)>
            {
                ("➕ Tạo trắc dọc", "CTPV_TaoProfileView"),
                ("✏ Edit profile", "CTPV_SuaProfileView"),
                ("📋 Thêm bảng lý trình", "CTPV_ThemBang_LyTrinh"),
                ("🏷 Thêm Label cao độ", "CTPV_ThemLabel_CaoDo"),
                ("📏 Thay đổi Scale", "CTPV_ThayDoiScale"),
                ("📐 Fit khung", "CTPV_FitKhung"),
            };
        }

        private List<(string Name, string Command)> GetSectionViewCommands()
        {
            return new List<(string, string)>
            {
                ("🎨 Tạo trắc ngang", "CTSV_VeTracNgangThietKe"),
                ("🎨 Vẽ tất cả TN", "CVSV_VeTatCa_TracNgangThietKe"),
                ("🔄 Chuyển TK sang TN", "CTSV_ChuyenDoi_TNTK_TNTN"),
                ("─────────────", ""),
                ("📐 Đánh cấp - VHC", "CTSV_DanhCap"),
                ("❌ Xóa đánh cấp", "CTSV_DanhCap_XoaBo"),
                ("➕ Vẽ thêm đánh cấp", "CTSV_DanhCap_VeThem"),
                ("➕ Vẽ thêm 1m", "CTSV_DanhCap_VeThem1"),
                ("➕ Vẽ thêm 2m", "CTSV_DanhCap_VeThem2"),
                ("🔄 Cập nhật KL đánh cấp", "CTSV_DanhCap_CapNhat"),
                ("─────────────", ""),
                ("📋 Điền KL trắc ngang", "CTSV_ThemVatLieu_TrenCatNgang"),
                ("⚙ Hiệu chỉnh MSS", "CTSV_ThayDoi_MSS_Min_Max"),
                ("↔ Thay giới hạn T/P", "CTSV_ThayDoi_GioiHan_traiPhai"),
                ("📋 Dàn khung in", "CTSV_ThayDoi_KhungIn"),
                ("🔒 Khóa TN + Add Point", "CTSV_KhoaCatNgang_AddPoint"),
                ("─────────────", ""),
                ("📐 Fit khung in", "CTSV_fit_KhungIn"),
                ("📐 Fit khung 5x5", "CTSV_fit_KhungIn_5_5_top"),
                ("📐 Fit khung 5x10", "CTSV_fit_KhungIn_5_10_top"),
                ("─────────────", ""),
                ("👁 Ẩn đường địa chất", "CTSV_An_DuongDiaChat"),
                ("✏ Hiệu chỉnh (Static)", "CTSV_HieuChinh_Section"),
                ("✏ Hiệu chỉnh (Dynamic)", "CTSV_HieuChinh_Section_Dynamic"),
                ("─────────────", ""),
                ("📊 Thống kê cọc (Excel)", "CTSV_ThongKeCoc"),
                ("📊 Thống kê toàn bộ cọc", "CTSV_ThongKeCoc_TatCa"),
                ("📍 Thống kê tọa độ cọc", "CTSV_ThongKeCoc_ToaDo"),
                ("─────────────", ""),
                ("📊 Taskbar Khối Lượng", "CTSV_Taskbar"),
                ("📥 Xuất KL Excel", "CTSV_XuatKhoiLuong"),
                ("📥 Xuất KL CAD", "CTSV_XuatCad"),
                ("⚙ Cài đặt bảng KL", "CTSV_CaiDatBang")
            };
        }

        private List<(string Name, string Command)> GetProfileCommands()
        {
            return new List<(string, string)>
            {
                ("📦 Thông kê Parcel", "CTP_TaoBangThongKeParcel"),
                ("📦 Thống kê Parcel (Sắp xếp)", "CTP_TaoBangThongKeParcel_SapXep"),
                ("─────────────", ""),
                ("📦 Gán Property Set", "AT_Solid_Set_PropertySet"),
                ("ℹ Thông tin Solid", "AT_Solid_Show_Info"),
            };
        }

        private List<(string Name, string Command)> GetCorridorCommands()
        {
            return new List<(string, string)>
            {
                ("➕ Thêm tất cả Section", "CTC_AddAllSection"),
                ("🛤 Corridor rẽ phải", "CTC_TaoCooridor_DuongDoThi_RePhai"),
                ("─────────────", ""),
                ("🔧 Thống kê Pipe", "CTPS_TaoBangThongKePipe"),
                ("🔧 Thống kê Structure", "CTPS_TaoBangThongKeStructure"),
                ("📏 Đổi cao độ Pipe", "CTPS_ThayDoi_CaoDo_Pipe"),
                ("📏 Đổi cao độ Struct", "CTPS_ThayDoi_CaoDo_Structure"),
                ("🔄 Xoay Pipe 90°", "CTPS_XoayPipe_90do"),
                ("❌ Xóa con trùng", "CTPS_XoaConTrung"),
            };
        }

        private List<(string Name, string Command)> GetIntersectionCommands()
        {
            return new List<(string, string)>
            {
                ("🏷 Đánh tên nút giao", ""),
                ("⚙ Thiết lập thông số", ""),
            };
        }

        private List<(string Name, string Command)> GetSanNenCommands()
        {
            return new List<(string, string)>
            {
                ("📊 Mở Taskbar SN", "CTSN_Taskbar"),
                ("─────────────", ""),
                ("▦ Quản lý lưới", "CTSN_TaoLuoi"),
                ("📝 Điền cao độ lưới", "CTSN_NhapCaoDo"),
                ("🏔 Lấy CĐ Surface", "CTSN_Surface"),
                ("📋 Tính khối lượng SN", "CTSN_TinhKL"),
                ("📤 Xuất bảng KL CAD", "CTSN_XuatBang"),
            };
        }

        private List<(string Name, string Command)> GetPlanCommands()
        {
            return new List<(string, string)>
            {
                ("📐 Thiết lập khung in", ""),
                ("📋 Dàn khung in", "CTSV_ThayDoi_KhungIn"),
                ("─────────────", ""),
                ("📐 Fit khung in", "CTSV_fit_KhungIn"),
            };
        }

        private List<(string Name, string Command)> GetExternalToolsCommands()
        {
            return new List<(string, string)>
            {
                ("📍 Tọa độ hố ga", "CT_VTOADOHG"),
                ("─────────────", ""),
                ("📋 Danh sách lệnh", "CT_DanhSachLenh"),
            };
        }

        #endregion

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

        private void ShowCommandListForm()
        {
            CommandListForm.ShowForm();
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
            this.BackColor = Color.FromArgb(45, 45, 48);
            this.ForeColor = Color.White;

            // Search box
            var lblSearch = new Label
            {
                Text = "🔍 Tìm kiếm:",
                Location = new Point(10, 12),
                Size = new Size(80, 25),
                ForeColor = Color.White
            };
            this.Controls.Add(lblSearch);

            searchBox = new TextBox
            {
                Location = new Point(95, 10),
                Size = new Size(300, 25),
                BackColor = Color.FromArgb(60, 60, 65),
                ForeColor = Color.White,
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
                BackColor = Color.FromArgb(0, 120, 215),
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
                BackColor = Color.FromArgb(150, 50, 50),
                ForeColor = Color.White
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
                BackColor = Color.FromArgb(35, 35, 38),
                ForeColor = Color.White
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
                ("Pipe", "CTPI_ThayDoi_CaoDoDayCong", "Thay đổi cao độ đáy cống"),
                ("Pipe", "CTPI_TaoBang_CaoDoDayHoGa", "Tạo bảng cao độ đáy hố ga"),
                
                // Point
                ("Point", "CTPO_TaoCogoPoint_CaoDo_FromSurface", "Tạo CogoPoint lấy cao độ từ Surface"),
                ("Point", "CTPO_TaoCogoPoint_CaoDoSpotElevation", "Tạo CogoPoint từ Spot Elevation"),
                ("Point", "CTPO_TaoCogoPoint_CaoDo_FromText", "Tạo CogoPoint từ Text cao độ"),
                ("Point", "CTPO_AnCacPoint", "Ẩn các điểm CogoPoint"),
                ("Point", "CTPO_TaoSurfaceFromPoints", "Tạo Surface từ Points"),
                
                // Profile
                ("Profile", "CTP_VeTracDoc_TuNhien", "Vẽ trắc dọc địa hình tự nhiên"),
                ("Profile", "CTP_SuaProfile_TheoSampleline", "Sửa profile theo sampleline"),
                ("Profile", "CTP_TaoBangThongKeParcel", "Tạo bảng thống kê parcel"),
                ("Profile", "CTP_ThemLabel_NutGiao", "Thêm label nút giao"),
                ("Profile", "CTP_VeTracDoc_ThietKe", "Vẽ trắc dọc thiết kế"),
                
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
