// MenuSettingsForm.cs - Form cài đặt & tùy chỉnh menu lệnh Civil Tool
// Cho phép: gộp/tách lệnh, tạo nhóm mới, đổi tên, drag-drop, lưu/load config

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.MenuSettingsCommand))]

namespace MyFirstProject
{
    #region Data Model

    /// <summary>
    /// Cấu trúc 1 lệnh trong menu
    /// </summary>
    public class MenuCommandItem
    {
        public string Command { get; set; } = "";
        public string Label { get; set; } = "";
        public string SubIcon { get; set; } = "ADD";

        public override string ToString() => $"{Label}  [{Command}]";
    }

    /// <summary>
    /// Cấu trúc 1 nhóm lệnh (Panel)
    /// </summary>
    public class MenuGroupItem
    {
        public string GroupName { get; set; } = "";
        public string IconType { get; set; } = "UTILITY";
        public List<MenuCommandItem> Commands { get; set; } = new();

        public override string ToString() => $"{GroupName} ({Commands.Count} lệnh)";
    }

    /// <summary>
    /// Toàn bộ cấu hình menu
    /// </summary>
    public class MenuConfig
    {
        public List<MenuGroupItem> CivilGroups { get; set; } = new();
        public List<MenuGroupItem> AcadGroups { get; set; } = new();
        public DateTime LastModified { get; set; } = DateTime.Now;
    }

    #endregion

    #region Settings Form

    public class MenuSettingsForm : Form
    {
        // ─── Data ───
        private MenuConfig _config = new();
        private static readonly string ConfigFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "CivilTool");
        private static readonly string ConfigPath = Path.Combine(ConfigFolder, "menu_config.json");

        // ─── Controls ───
        private TabControl tabMain = null!;

        // Civil tab
        private ListBox lstGroups = null!;
        private ListBox lstCommands = null!;
        private ListBox lstAllCommands = null!;

        // Buttons
        private Button btnAddGroup = null!, btnRemoveGroup = null!, btnRenameGroup = null!;
        private Button btnMoveUp = null!, btnMoveDown = null!;
        private Button btnAddCmd = null!, btnRemoveCmd = null!;
        private Button btnCmdUp = null!, btnCmdDown = null!;
        private Button btnSave = null!, btnCancel = null!, btnReset = null!;
        private Button btnExport = null!, btnImport = null!;

        // ─── Constructor ───
        public MenuSettingsForm()
        {
            InitUI();
            LoadConfig();
            RefreshGroupList();
            RefreshAllCommandsList();
        }

        // ─── Kết quả ───
        public MenuConfig GetConfig() => _config;

        #region UI Setup

        private void InitUI()
        {
            // Form
            this.Text = "⚙ Civil Tool — Cài Đặt Menu";
            this.Size = new Size(1050, 680);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9.5f);
            this.BackColor = Color.FromArgb(30, 30, 34);
            this.ForeColor = Color.White;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // ─── HEADER ───
            var header = new Panel { Dock = DockStyle.Top, Height = 50, BackColor = Color.FromArgb(0, 122, 204) };
            var lblTitle = new Label
            {
                Text = "⚙  CÀI ĐẶT MENU LỆNH",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                Padding = new Padding(15, 0, 0, 0)
            };
            header.Controls.Add(lblTitle);
            this.Controls.Add(header);

            // ─── MAIN CONTENT ───
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12)
            };
            this.Controls.Add(mainPanel);
            mainPanel.BringToFront();

            // 3-column layout
            // Col1: Nhóm lệnh (Groups)
            var pnlGroups = CreateSection("📁 NHÓM LỆNH", 10, 10, 250, 480, mainPanel);
            lstGroups = CreateListBox(pnlGroups, 5, 30, 240, 340);
            lstGroups.SelectedIndexChanged += (s, e) => RefreshCommandList();

            btnAddGroup = CreateButton("➕ Thêm nhóm", 5, 380, 115, 32, pnlGroups, BtnAddGroup_Click);
            btnRemoveGroup = CreateButton("🗑 Xóa nhóm", 125, 380, 115, 32, pnlGroups, BtnRemoveGroup_Click);
            btnRenameGroup = CreateButton("✏ Đổi tên", 5, 418, 115, 32, pnlGroups, BtnRenameGroup_Click);
            btnMoveUp = CreateButton("▲ Lên", 125, 418, 55, 32, pnlGroups, BtnGroupUp_Click);
            btnMoveDown = CreateButton("▼ Xuống", 185, 418, 55, 32, pnlGroups, BtnGroupDown_Click);

            // Col2: Lệnh trong nhóm
            var pnlCmds = CreateSection("📋 LỆNH TRONG NHÓM", 270, 10, 270, 480, mainPanel);
            lstCommands = CreateListBox(pnlCmds, 5, 30, 260, 340);

            btnCmdUp = CreateButton("▲", 5, 380, 40, 32, pnlCmds, BtnCmdUp_Click);
            btnCmdDown = CreateButton("▼", 50, 380, 40, 32, pnlCmds, BtnCmdDown_Click);
            btnRemoveCmd = CreateButton("◀ Xóa khỏi nhóm", 95, 380, 165, 32, pnlCmds, BtnRemoveCmd_Click);

            // Col3: Tất cả lệnh (để thêm vào nhóm)
            var pnlAll = CreateSection("📦 TẤT CẢ LỆNH CÓ SẴN", 550, 10, 270, 480, mainPanel);
            lstAllCommands = CreateListBox(pnlAll, 5, 30, 260, 340);

            btnAddCmd = CreateButton("▶ Thêm vào nhóm", 5, 380, 165, 32, pnlAll, BtnAddCmd_Click);

            // Search box
            var txtSearch = new TextBox
            {
                Location = new Point(5, 418),
                Size = new Size(260, 26),
                BackColor = Color.FromArgb(60, 60, 67),
                ForeColor = Color.White,
                PlaceholderText = "🔍 Tìm lệnh...",
                BorderStyle = BorderStyle.FixedSingle
            };
            txtSearch.TextChanged += (s, e) => FilterAllCommands(txtSearch.Text);
            pnlAll.Controls.Add(txtSearch);

            // Col4: Separator + Add separator button
            var pnlSep = CreateSection("➖ NGĂN CÁCH", 830, 10, 190, 200, mainPanel);
            var btnAddSep = CreateButton("Thêm vạch ngăn cách", 5, 30, 180, 32, pnlSep, BtnAddSeparator_Click);
            var lblHint = new Label
            {
                Text = "Vạch ngăn cách (---) giúp\nphân chia lệnh trong 1 nhóm\nthành các phần nhỏ hơn.",
                Location = new Point(5, 70),
                Size = new Size(180, 60),
                ForeColor = Color.FromArgb(150, 150, 160)
            };
            pnlSep.Controls.Add(lblHint);

            // Icon combo for group
            var pnlIcon = CreateSection("🎨 ICON NHÓM", 830, 220, 190, 100, mainPanel);
            var cmbIcon = new ComboBox
            {
                Location = new Point(5, 30),
                Size = new Size(180, 26),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(60, 60, 67),
                ForeColor = Color.White
            };
            var iconTypes = new[] { "SURFACE", "STATION", "ALIGN", "CORRIDOR", "SECTION",
                "GRADING", "PIPE", "POINT", "UTILITY", "ACCOUNT", "PROFILE",
                "VOLUME", "SLOPE", "EXCEL", "CALCULATE", "DRAW", "SETTINGS" };
            cmbIcon.Items.AddRange(iconTypes);
            cmbIcon.SelectedIndexChanged += (s, e) =>
            {
                if (lstGroups.SelectedIndex >= 0 && cmbIcon.SelectedItem is string icon)
                {
                    var groups = GetCurrentGroupList();
                    if (lstGroups.SelectedIndex < groups.Count)
                        groups[lstGroups.SelectedIndex].IconType = icon;
                }
            };
            pnlIcon.Controls.Add(cmbIcon);

            // ─── BOTTOM buttons ───
            var pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 55,
                BackColor = Color.FromArgb(38, 38, 42),
                Padding = new Padding(12, 8, 12, 8)
            };
            this.Controls.Add(pnlBottom);

            btnImport = CreateButton("📂 Import", 5, 8, 100, 36, pnlBottom, BtnImport_Click);
            btnExport = CreateButton("💾 Export", 110, 8, 100, 36, pnlBottom, BtnExport_Click);
            btnReset = CreateButton("🔄 Reset mặc định", 220, 8, 150, 36, pnlBottom, BtnReset_Click);

            btnCancel = CreateButton("✖ Hủy", 750, 8, 100, 36, pnlBottom, (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); });
            btnSave = CreateButton("✅ Lưu & Áp dụng", 860, 8, 150, 36, pnlBottom, BtnSave_Click);
            btnSave.BackColor = Color.FromArgb(0, 122, 204);
            btnSave.FlatAppearance.BorderColor = Color.FromArgb(0, 102, 180);

            // Tab for Civil/Acad
            tabMain = new TabControl
            {
                Location = new Point(12, 5),
                Size = new Size(200, 28),
                Appearance = TabAppearance.Buttons,
                ItemSize = new Size(95, 24)
            };
            tabMain.TabPages.Add("civil", "🛠 CIVIL");
            tabMain.TabPages.Add("acad", "📏 ACAD");
            tabMain.SelectedIndexChanged += (s, e) => { RefreshGroupList(); RefreshAllCommandsList(); };
            mainPanel.Controls.Add(tabMain);
            tabMain.BringToFront();
        }

        private Panel CreateSection(string title, int x, int y, int w, int h, Control parent)
        {
            var panel = new Panel
            {
                Location = new Point(x, y + 35), // offset for tab
                Size = new Size(w, h),
                BackColor = Color.FromArgb(40, 40, 45),
                BorderStyle = BorderStyle.FixedSingle
            };
            var lbl = new Label
            {
                Text = title,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 149, 217),
                Location = new Point(5, 5),
                AutoSize = true
            };
            panel.Controls.Add(lbl);
            parent.Controls.Add(panel);
            return panel;
        }

        private ListBox CreateListBox(Control parent, int x, int y, int w, int h)
        {
            var lb = new ListBox
            {
                Location = new Point(x, y),
                Size = new Size(w, h),
                BackColor = Color.FromArgb(50, 50, 55),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9),
                IntegralHeight = false,
                SelectionMode = SelectionMode.One
            };
            parent.Controls.Add(lb);
            return lb;
        }

        private Button CreateButton(string text, int x, int y, int w, int h, Control parent, EventHandler onClick)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(55, 55, 62),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 8.5f),
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 90);
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 122, 204);
            btn.Click += onClick;
            parent.Controls.Add(btn);
            return btn;
        }

        #endregion

        #region Config Load/Save

        private List<MenuGroupItem> GetCurrentGroupList()
        {
            return tabMain.SelectedIndex == 0 ? _config.CivilGroups : _config.AcadGroups;
        }

        public void LoadConfig()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    string json = File.ReadAllText(ConfigPath);
                    _config = JsonSerializer.Deserialize<MenuConfig>(json) ?? GetDefaultConfig();
                }
                else
                {
                    _config = GetDefaultConfig();
                }
            }
            catch
            {
                _config = GetDefaultConfig();
            }
        }

        private void SaveConfig()
        {
            try
            {
                if (!Directory.Exists(ConfigFolder))
                    Directory.CreateDirectory(ConfigFolder);

                _config.LastModified = DateTime.Now;
                string json = JsonSerializer.Serialize(_config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigPath, json);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi lưu config: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static MenuConfig LoadSavedConfig()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    string json = File.ReadAllText(ConfigPath);
                    return JsonSerializer.Deserialize<MenuConfig>(json) ?? GetDefaultConfig();
                }
            }
            catch { }
            return GetDefaultConfig();
        }

        public static MenuConfig GetDefaultConfig()
        {
            var config = new MenuConfig();

            // ── CIVIL TOOL panels ──
            config.CivilGroups = new List<MenuGroupItem>
            {
                new() { GroupName = "🗺 Bề Mặt", IconType = "SURFACE", Commands = ConvertCmds(new[]
                {
                    ("CTSu_TaoSpotElevation_OnSurface_TaiTim", "Spot Elevation Tại Tim", "ADD"),
                }) },

                new() { GroupName = "📍 Lưới Cọc", IconType = "STATION", Commands = ConvertCmds(new[]
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
                    ("CTSV_ThongKeCoc_TatCa", "Thống Kê Tất Cả Cọc", "EXPORT"),
                }) },

                new() { GroupName = "🛣 Tuyến", IconType = "ALIGN", Commands = ConvertCmds(new[]
                {
                    ("CTPV_TaoProfileView", "Tạo Trắc Dọc", "PROFILE"),
                    ("CTPV_SuaProfileView", "Edit Profile", "SETTINGS"),
                    ("CTPV_ThemBang_LyTrinh", "Thêm Bảng Lý Trình", "TABLE"),
                    ("CTPV_ThemLabel_CaoDo", "Thêm Label Cao Độ", "ADD"),
                    ("CTPV_ThayDoiScale", "Thay Đổi Scale", "SETTINGS"),
                    ("CTPV_FitKhung", "Fit Khung", "VIEW"),
                }) },

                new() { GroupName = "🛤 Corridor", IconType = "CORRIDOR", Commands = ConvertCmds(new[]
                {
                    ("CTC_AddAllSection", "Thêm Tất Cả Section", "ADD"),
                    ("CTC_TaoCooridor_DuongDoThi_RePhai", "Corridor Rẽ Phải", "DRAW"),
                }) },

                new() { GroupName = "📐 Trắc Ngang", IconType = "SECTION", Commands = ConvertCmds(new[]
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
                    ("CTSV_CaiDatBang", "Cài Đặt Bảng KL", "SETTINGS"),
                }) },

                new() { GroupName = "▦ San Nền", IconType = "GRADING", Commands = ConvertCmds(new[]
                {
                    ("CTSN_Taskbar", "Mở Taskbar SN", "VOLUME"),
                    ("---", "", ""),
                    ("CTSN_TaoLuoi", "Quản Lý Lưới", "GRADING"),
                    ("CTSN_NhapCaoDo", "Điền Cao Độ Lưới", "ADD"),
                    ("CTSN_Surface", "Lấy CĐ Surface", "SURFACE"),
                    ("CTSN_TinhKL", "Tính Khối Lượng SN", "VOLUME"),
                    ("CTSN_XuatBang", "Xuất Bảng KL CAD", "CAD"),
                }) },

                new() { GroupName = "🔧 Thoát Nước", IconType = "PIPE", Commands = ConvertCmds(new[]
                {
                    ("CTPS_TaoBangThongKePipe", "Thống Kê Pipe", "TABLE"),
                    ("CTPS_TaoBangThongKeStructure", "Thống Kê Structure", "TABLE"),
                    ("CTPS_ThayDoi_CaoDo_Pipe", "Đổi Cao Độ Pipe", "SETTINGS"),
                    ("CTPS_ThayDoi_CaoDo_Structure", "Đổi Cao Độ Struct", "SETTINGS"),
                    ("CTPS_XoayPipe_90do", "Xoay Pipe 90°", "SYNC"),
                    ("CTPS_XoaConTrung", "Xóa Con Trùng", "DELETE"),
                }) },

                new() { GroupName = "📌 Điểm", IconType = "POINT", Commands = ConvertCmds(new[]
                {
                    ("CTPo_TaoPointTheoBang", "Tạo Point Từ Bảng", "ADD"),
                    ("CTPo_ChuyenPointThanhBlock", "Point → Block", "SYNC"),
                    ("CTPo_TaoBangThongKePoint", "Bảng Thống Kê Point", "TABLE"),
                    ("CTPo_ThayDoiCaoDo", "Thay Đổi Cao Độ", "SETTINGS"),
                    ("CTPo_DatTen_theoThuTu", "Đặt Tên Thứ Tự", "RENAME"),
                    ("CTPo_ThayDoiStyle", "Thay Đổi Style", "SETTINGS"),
                    ("CTPo_LayThongTin", "Lấy Thông Tin", "INFO"),
                }) },

                new() { GroupName = "⚙ Tiện Ích", IconType = "UTILITY", Commands = ConvertCmds(new[]
                {
                    ("AT_Solid_Set_PropertySet", "Gán Property Set", "ADD"),
                    ("AT_Solid_Show_Info", "Thông Tin Solid", "INFO"),
                    ("CT_VTOADOHG", "Tọa Độ Hố Ga", "POINT"),
                    ("---", "", ""),
                    ("CIVIL_RIBBON", "Reload Menu", "RELOAD"),
                    ("CT_MENU_SETTINGS", "Cài Đặt Menu", "SETTINGS"),
                }) },

                new() { GroupName = "👤 Tài Khoản", IconType = "ACCOUNT", Commands = ConvertCmds(new[]
                {
                    ("", "Đăng Nhập", "ACCOUNT"),
                    ("", "Thông Tin", "INFO"),
                    ("", "Hướng Dẫn", "INFO"),
                }) },
            };

            // ── ACAD TOOL panels ──
            config.AcadGroups = new List<MenuGroupItem>
            {
                new() { GroupName = "📏 Đo Lường", IconType = "MEASURE", Commands = ConvertCmds(new[]
                {
                    ("AT_TongDoDai_Full", "Tổng Độ Dài (Full)", "CALCULATE"),
                    ("AT_TongDoDai_Replace", "Tổng Độ Dài (Replace)", "CALCULATE"),
                    ("AT_TongDoDai_Replace2", "Tổng Độ Dài (Replace2)", "CALCULATE"),
                    ("AT_TongDoDai_Replace_CongThem", "Tổng Độ Dài (Cộng Thêm)", "CALCULATE"),
                    ("---", "", ""),
                    ("AT_TongDienTich_Full", "Tổng Diện Tích (Full)", "CALCULATE"),
                    ("AT_TongDienTich_Replace", "Tổng Diện Tích (Replace)", "CALCULATE"),
                    ("AT_TongDienTich_Replace2", "Tổng Diện Tích (Replace2)", "CALCULATE"),
                    ("AT_TongDienTich_Replace_CongThem", "Tổng Diện Tích (Cộng Thêm)", "CALCULATE"),
                }) },

                new() { GroupName = "✏ Chỉnh Sửa", IconType = "EDIT", Commands = ConvertCmds(new[]
                {
                    ("AT_TextLink", "Text Link", "ADD"),
                    ("AT_DanhSoThuTu", "Đánh Số Thứ Tự", "ADD"),
                    ("AT_XoaDoiTuong_CungLayer", "Xóa Đối Tượng Cùng Layer", "DELETE"),
                    ("AT_XoaDoiTuong_3DSolid_Body", "Xóa 3DSolid/Body", "DELETE"),
                    ("AT_Offset_2Ben", "Offset 2 Bên", "COPY"),
                    ("AT_annotive_scale_currentOnly", "Annotative Scale Current Only", "SETTINGS"),
                    ("---", "", ""),
                    ("AT_XoayDoiTuong_TheoViewport", "Xoay Theo Viewport", "SYNC"),
                    ("AT_XoayDoiTuong_Theo2Diem", "Xoay Theo 2 Điểm", "SYNC"),
                }) },

                new() { GroupName = "🔄 Layout", IconType = "LAYOUT", Commands = ConvertCmds(new[]
                {
                    ("AT_TextLayout", "Text Layout", "ADD"),
                    ("AT_TaoMoi_TextLayout", "Tạo Mới Text Layout", "ADD"),
                    ("AT_DimLayout", "Dim Layout", "ADD"),
                    ("AT_DimLayout2", "Dim Layout 2", "ADD"),
                    ("AT_BlockLayout", "Block Layout", "ADD"),
                    ("AT_Label_FromText", "Label From Text", "ADD"),
                    ("---", "", ""),
                    ("AT_UpdateLayout", "Update Layout", "SYNC"),
                }) },
            };

            return config;
        }

        private static List<MenuCommandItem> ConvertCmds((string Cmd, string Label, string Icon)[] items)
        {
            return items.Select(i => new MenuCommandItem
            {
                Command = i.Cmd,
                Label = i.Label,
                SubIcon = i.Icon
            }).ToList();
        }

        #endregion

        #region Refresh UI

        private void RefreshGroupList()
        {
            lstGroups.Items.Clear();
            var groups = GetCurrentGroupList();
            foreach (var g in groups)
                lstGroups.Items.Add(g);

            if (groups.Count > 0)
                lstGroups.SelectedIndex = 0;
        }

        private void RefreshCommandList()
        {
            lstCommands.Items.Clear();
            if (lstGroups.SelectedIndex < 0) return;

            var groups = GetCurrentGroupList();
            if (lstGroups.SelectedIndex >= groups.Count) return;

            var group = groups[lstGroups.SelectedIndex];
            foreach (var cmd in group.Commands)
            {
                if (cmd.Command == "---")
                    lstCommands.Items.Add("─── ngăn cách ───");
                else
                    lstCommands.Items.Add(cmd);
            }
        }

        private List<MenuCommandItem> _allCommandsCache = new();

        private void RefreshAllCommandsList()
        {
            _allCommandsCache = GetAllAvailableCommands();
            lstAllCommands.Items.Clear();
            foreach (var cmd in _allCommandsCache)
                lstAllCommands.Items.Add(cmd);
        }

        private void FilterAllCommands(string filter)
        {
            lstAllCommands.Items.Clear();
            var filtered = string.IsNullOrWhiteSpace(filter)
                ? _allCommandsCache
                : _allCommandsCache.Where(c =>
                    c.Label.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                    c.Command.Contains(filter, StringComparison.OrdinalIgnoreCase)).ToList();

            foreach (var cmd in filtered)
                lstAllCommands.Items.Add(cmd);
        }

        /// <summary>
        /// Trả về tất cả lệnh có sẵn (gộp từ default config)
        /// </summary>
        private List<MenuCommandItem> GetAllAvailableCommands()
        {
            var defaultConfig = GetDefaultConfig();
            var allGroups = tabMain.SelectedIndex == 0 ? defaultConfig.CivilGroups : defaultConfig.AcadGroups;
            return allGroups
                .SelectMany(g => g.Commands)
                .Where(c => !string.IsNullOrEmpty(c.Command) && c.Command != "---")
                .GroupBy(c => c.Command)
                .Select(g => g.First())
                .OrderBy(c => c.Label)
                .ToList();
        }

        #endregion

        #region Group Operations

        private void BtnAddGroup_Click(object? sender, EventArgs e)
        {
            string name = PromptInput("Tên nhóm mới:", "Thêm Nhóm Lệnh");
            if (string.IsNullOrWhiteSpace(name)) return;

            var groups = GetCurrentGroupList();
            groups.Add(new MenuGroupItem { GroupName = name, IconType = "UTILITY" });
            RefreshGroupList();
            lstGroups.SelectedIndex = groups.Count - 1;
        }

        private void BtnRemoveGroup_Click(object? sender, EventArgs e)
        {
            if (lstGroups.SelectedIndex < 0) return;
            var groups = GetCurrentGroupList();
            var group = groups[lstGroups.SelectedIndex];

            if (MessageBox.Show($"Xóa nhóm \"{group.GroupName}\" và {group.Commands.Count} lệnh bên trong?",
                "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;

            groups.RemoveAt(lstGroups.SelectedIndex);
            RefreshGroupList();
        }

        private void BtnRenameGroup_Click(object? sender, EventArgs e)
        {
            if (lstGroups.SelectedIndex < 0) return;
            var groups = GetCurrentGroupList();
            var group = groups[lstGroups.SelectedIndex];

            string name = PromptInput("Tên mới:", "Đổi Tên Nhóm", group.GroupName);
            if (string.IsNullOrWhiteSpace(name)) return;

            group.GroupName = name;
            RefreshGroupList();
        }

        private void BtnGroupUp_Click(object? sender, EventArgs e)
        {
            MoveItem(lstGroups, GetCurrentGroupList(), -1);
            RefreshGroupList();
        }

        private void BtnGroupDown_Click(object? sender, EventArgs e)
        {
            MoveItem(lstGroups, GetCurrentGroupList(), +1);
            RefreshGroupList();
        }

        #endregion

        #region Command Operations

        private void BtnAddCmd_Click(object? sender, EventArgs e)
        {
            if (lstGroups.SelectedIndex < 0)
            {
                MessageBox.Show("Vui lòng chọn nhóm trước.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (lstAllCommands.SelectedItem is not MenuCommandItem cmd) return;

            var groups = GetCurrentGroupList();
            var group = groups[lstGroups.SelectedIndex];
            group.Commands.Add(new MenuCommandItem
            {
                Command = cmd.Command,
                Label = cmd.Label,
                SubIcon = cmd.SubIcon
            });
            RefreshCommandList();
        }

        private void BtnRemoveCmd_Click(object? sender, EventArgs e)
        {
            if (lstGroups.SelectedIndex < 0 || lstCommands.SelectedIndex < 0) return;

            var groups = GetCurrentGroupList();
            var group = groups[lstGroups.SelectedIndex];
            int idx = lstCommands.SelectedIndex;

            if (idx < group.Commands.Count)
            {
                group.Commands.RemoveAt(idx);
                RefreshCommandList();
            }
        }

        private void BtnCmdUp_Click(object? sender, EventArgs e)
        {
            if (lstGroups.SelectedIndex < 0) return;
            var group = GetCurrentGroupList()[lstGroups.SelectedIndex];
            int idx = lstCommands.SelectedIndex;
            if (idx <= 0 || idx >= group.Commands.Count) return;

            (group.Commands[idx], group.Commands[idx - 1]) = (group.Commands[idx - 1], group.Commands[idx]);
            RefreshCommandList();
            lstCommands.SelectedIndex = idx - 1;
        }

        private void BtnCmdDown_Click(object? sender, EventArgs e)
        {
            if (lstGroups.SelectedIndex < 0) return;
            var group = GetCurrentGroupList()[lstGroups.SelectedIndex];
            int idx = lstCommands.SelectedIndex;
            if (idx < 0 || idx >= group.Commands.Count - 1) return;

            (group.Commands[idx], group.Commands[idx + 1]) = (group.Commands[idx + 1], group.Commands[idx]);
            RefreshCommandList();
            lstCommands.SelectedIndex = idx + 1;
        }

        private void BtnAddSeparator_Click(object? sender, EventArgs e)
        {
            if (lstGroups.SelectedIndex < 0) return;
            var group = GetCurrentGroupList()[lstGroups.SelectedIndex];

            int insertAt = lstCommands.SelectedIndex >= 0 ? lstCommands.SelectedIndex + 1 : group.Commands.Count;
            group.Commands.Insert(insertAt, new MenuCommandItem { Command = "---", Label = "", SubIcon = "" });
            RefreshCommandList();
        }

        #endregion

        #region Bottom Buttons

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            SaveConfig();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnReset_Click(object? sender, EventArgs e)
        {
            if (MessageBox.Show("Reset toàn bộ menu về mặc định?",
                "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            _config = GetDefaultConfig();
            RefreshGroupList();
            RefreshAllCommandsList();
        }

        private void BtnImport_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Title = "Import cấu hình menu",
                Filter = "JSON Files|*.json",
                DefaultExt = "json"
            };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            try
            {
                string json = File.ReadAllText(ofd.FileName);
                _config = JsonSerializer.Deserialize<MenuConfig>(json) ?? GetDefaultConfig();
                RefreshGroupList();
                RefreshAllCommandsList();
                MessageBox.Show("Import thành công!", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi import: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExport_Click(object? sender, EventArgs e)
        {
            using var sfd = new SaveFileDialog
            {
                Title = "Export cấu hình menu",
                Filter = "JSON Files|*.json",
                DefaultExt = "json",
                FileName = "CivilTool_MenuConfig.json"
            };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                string json = JsonSerializer.Serialize(_config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(sfd.FileName, json);
                MessageBox.Show("Export thành công!", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi export: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Helpers

        private static void MoveItem<T>(ListBox lb, List<T> list, int direction)
        {
            int idx = lb.SelectedIndex;
            int newIdx = idx + direction;
            if (idx < 0 || newIdx < 0 || newIdx >= list.Count) return;
            (list[idx], list[newIdx]) = (list[newIdx], list[idx]);
            lb.SelectedIndex = newIdx;
        }

        private string PromptInput(string prompt, string title, string defaultValue = "")
        {
            using var form = new Form
            {
                Text = title,
                Size = new Size(400, 160),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = Color.White,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var lbl = new Label
            {
                Text = prompt,
                Location = new Point(15, 15),
                AutoSize = true
            };
            var txt = new TextBox
            {
                Text = defaultValue,
                Location = new Point(15, 40),
                Size = new Size(350, 26),
                BackColor = Color.FromArgb(60, 60, 67),
                ForeColor = Color.White
            };
            var btnOk = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new Point(200, 75),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White
            };
            var btnCn = new Button
            {
                Text = "Hủy",
                DialogResult = DialogResult.Cancel,
                Location = new Point(285, 75),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(55, 55, 62),
                ForeColor = Color.White
            };

            form.Controls.AddRange(new Control[] { lbl, txt, btnOk, btnCn });
            form.AcceptButton = btnOk;
            form.CancelButton = btnCn;

            return form.ShowDialog() == DialogResult.OK ? txt.Text.Trim() : "";
        }

        #endregion
    }

    #endregion

    #region AutoCAD Command

    public class MenuSettingsCommand
    {
        [CommandMethod("CT_MENU_SETTINGS")]
        public static void OpenMenuSettings()
        {
            try
            {
                var form = new MenuSettingsForm();
                var result = AcadApp.ShowModalDialog(form);

                if (result == DialogResult.OK)
                {
                    var config = form.GetConfig();

                    // Áp dụng lại Ribbon với config mới
                    ApplyConfigToRibbon(config);

                    var ed = AcadApp.DocumentManager.MdiActiveDocument?.Editor;
                    ed?.WriteMessage("\n✅ Đã cập nhật menu theo cài đặt mới!");
                }
            }
            catch (System.Exception ex)
            {
                var ed = AcadApp.DocumentManager.MdiActiveDocument?.Editor;
                ed?.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Áp dụng config vào Ribbon hiện tại  
        /// </summary>
        private static void ApplyConfigToRibbon(MenuConfig config)
        {
            var ribbon = Autodesk.Windows.ComponentManager.Ribbon;
            if (ribbon == null) return;

            // Xóa tab cũ
            var tabIds = new[] { "CivilTool.MainTab", "CivilTool.AcadTab" };
            foreach (var id in tabIds)
            {
                var existing = ribbon.Tabs.FirstOrDefault(t => t.Id == id);
                if (existing != null)
                    ribbon.Tabs.Remove(existing);
            }

            // Tạo lại Civil Tab
            var civilTab = new Autodesk.Windows.RibbonTab
            {
                Title = "🛠 CIVIL TOOL",
                Id = "CivilTool.MainTab"
            };
            ribbon.Tabs.Add(civilTab);

            foreach (var group in config.CivilGroups)
            {
                if (group.Commands.Count == 0) continue;
                var cmds = group.Commands.Select(c => (c.Command, c.Label, c.SubIcon)).ToList();
                CivilRibbonTaskbar.AddCivilPanel(civilTab, group.GroupName, group.IconType, cmds);
            }
            civilTab.IsActive = true;

            // Tạo lại Acad Tab
            var acadTab = new Autodesk.Windows.RibbonTab
            {
                Title = "📏 ACAD TOOL",
                Id = "CivilTool.AcadTab"
            };
            ribbon.Tabs.Add(acadTab);

            foreach (var group in config.AcadGroups)
            {
                if (group.Commands.Count == 0) continue;
                var cmds = group.Commands.Select(c => (c.Command, c.Label, c.SubIcon)).ToList();
                CivilRibbonTaskbar.AddCivilPanel(acadTab, group.GroupName, group.IconType, cmds);
            }
        }
    }

    #endregion
}
