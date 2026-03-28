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
        public bool IsEnabled { get; set; } = true;

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
        public Dictionary<string, string> CommandAliases { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> CommandDescriptions { get; set; } = new(StringComparer.OrdinalIgnoreCase);
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

        // ─── Controls ───        // ==========================================
        //         CÁC BIẾN UI & DỮ LIỆU
        // ==========================================
        private Panel pnlGroups = null!, pnlCmds = null!, pnlAll = null!, pnlSep = null!, pnlIcon = null!, pnlAliasTab = null!;
        private TabControl tabMain = null!;

        // Civil tab
        private ListBox lstGroups = null!;
        private CheckedListBox lstCommands = null!;
        private ListBox lstAllCommands = null!;

        // Buttons
        private Button btnAddGroup = null!, btnRemoveGroup = null!, btnRenameGroup = null!;
        private Button btnMoveUp = null!, btnMoveDown = null!;
        private Button btnAddCmd = null!, btnRemoveCmd = null!;
        private Button btnCmdUp = null!, btnCmdDown = null!;
        private Button btnSave = null!, btnCancel = null!, btnReset = null!;
        private Button btnExport = null!, btnImport = null!;
        private ComboBox cmbCmdIcon = null!;
        private PictureBox picCmdIconPreview = null!;
        private PictureBox picGroupIconPreview = null!;
        private DataGridView _dgvAliases = null!;
        private bool _suppressItemCheck = false;

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
            this.BackColor = Color.FromArgb(30, 30, 40);
            this.ForeColor = Color.White;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // ─── HEADER ───
            var header = new Panel { Dock = DockStyle.Top, Height = 50, BackColor = Color.FromArgb(27, 42, 74) };
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
            pnlGroups = CreateSection("📁 NHÓM LỆNH", 10, 10, 250, 480, mainPanel);
            lstGroups = CreateListBox(pnlGroups, 5, 30, 240, 340);
            lstGroups.SelectedIndexChanged += (s, e) => RefreshCommandList();

            btnAddGroup = CreateButton("➕ Thêm nhóm", 5, 380, 115, 32, pnlGroups, BtnAddGroup_Click);
            btnRemoveGroup = CreateButton("🗑 Xóa nhóm", 125, 380, 115, 32, pnlGroups, BtnRemoveGroup_Click);
            btnRenameGroup = CreateButton("✏ Đổi tên", 5, 418, 115, 32, pnlGroups, BtnRenameGroup_Click);
            btnMoveUp = CreateButton("▲ Lên", 125, 418, 55, 32, pnlGroups, BtnGroupUp_Click);
            btnMoveDown = CreateButton("▼ Xuống", 185, 418, 55, 32, pnlGroups, BtnGroupDown_Click);

            // Col2: Lệnh trong nhóm
            pnlCmds = CreateSection("📋 LỆNH TRONG NHÓM", 270, 10, 270, 480, mainPanel);
            lstCommands = new CheckedListBox
            {
                Location = new Point(5, 30),
                Size = new Size(260, 320),
                BackColor = Color.FromArgb(40, 40, 55),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9),
                IntegralHeight = false,
                CheckOnClick = true
            };
            lstCommands.SelectedIndexChanged += (s, e) => UpdateCmdIconCombo();
            lstCommands.ItemCheck += LstCommands_ItemCheck;
            pnlCmds.Controls.Add(lstCommands);

            // Nút Chọn tất cả / Bỏ chọn tất cả
            var btnCheckAll = CreateButton("☑ Tất cả", 5, 355, 128, 28, pnlCmds, BtnCheckAll_Click);
            var btnUncheckAll = CreateButton("☐ Bỏ tất cả", 137, 355, 128, 28, pnlCmds, BtnUncheckAll_Click);

            btnCmdUp = CreateButton("▲", 5, 390, 40, 32, pnlCmds, BtnCmdUp_Click);
            btnCmdDown = CreateButton("▼", 50, 390, 40, 32, pnlCmds, BtnCmdDown_Click);
            btnRemoveCmd = CreateButton("◀ Xóa khỏi nhóm", 95, 390, 165, 32, pnlCmds, BtnRemoveCmd_Click);

            // ───── ICON LỆNH (với preview) ─────
            var lblCmdIcon = new Label
            {
                Text = "🎨 Icon lệnh:",
                Location = new Point(5, 432),
                AutoSize = true,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold)
            };
            pnlCmds.Controls.Add(lblCmdIcon);

            cmbCmdIcon = new ComboBox
            {
                Location = new Point(85, 430),
                Size = new Size(140, 26),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(40, 40, 55),
                ForeColor = Color.White
            };
            var cmdIconTypes = new[] { "SURFACE", "STATION", "ALIGN", "CORRIDOR", "SECTION",
                "GRADING", "PIPE", "POINT", "UTILITY", "ACCOUNT", "PROFILE",
                "VOLUME", "SLOPE", "EXCEL", "CALCULATE", "DRAW", "SETTINGS",
                "ADD", "RENAME", "TABLE", "EXPORT", "SYNC", "INFO", "DELETE",
                "COPY", "MOVE", "RELOAD", "VIEW", "EDIT", "LAYOUT", "MEASURE",
                "SPIRAL", "SUPERELEVATION", "CULVERT", "BRIDGE", "WIDENING", "INTERSECTION", "CAD" };
            cmbCmdIcon.Items.AddRange(cmdIconTypes);
            cmbCmdIcon.SelectedIndexChanged += (s, e) =>
            {
                if (lstGroups.SelectedIndex >= 0 && lstCommands.SelectedIndex >= 0
                    && cmbCmdIcon.SelectedItem is string icon)
                {
                    var groups = GetCurrentGroupList();
                    if (lstGroups.SelectedIndex < groups.Count)
                    {
                        var group = groups[lstGroups.SelectedIndex];
                        int cmdIdx = lstCommands.SelectedIndex;
                        if (cmdIdx < group.Commands.Count && group.Commands[cmdIdx].Command != "---")
                        {
                            group.Commands[cmdIdx].SubIcon = icon;
                        }
                    }
                    UpdateIconPreview(picCmdIconPreview, icon);
                }
            };
            pnlCmds.Controls.Add(cmbCmdIcon);

            // Preview icon lệnh
            picCmdIconPreview = new PictureBox
            {
                Location = new Point(230, 427),
                Size = new Size(32, 32),
                BackColor = Color.FromArgb(40, 40, 55),
                SizeMode = PictureBoxSizeMode.CenterImage,
                BorderStyle = BorderStyle.FixedSingle
            };
            pnlCmds.Controls.Add(picCmdIconPreview);

            // Col3: Tất cả lệnh (để thêm vào nhóm)
            pnlAll = CreateSection("📦 TẤT CẢ LỆNH CÓ SẴN", 550, 10, 270, 480, mainPanel);
            lstAllCommands = CreateListBox(pnlAll, 5, 30, 260, 340);

            btnAddCmd = CreateButton("▶ Thêm vào nhóm", 5, 380, 165, 32, pnlAll, BtnAddCmd_Click);

            // Search box
            var txtSearch = new TextBox
            {
                Location = new Point(5, 418),
                Size = new Size(260, 26),
                BackColor = Color.FromArgb(40, 40, 55),
                ForeColor = Color.White,
                PlaceholderText = "🔍 Tìm lệnh...",
                BorderStyle = BorderStyle.FixedSingle
            };
            txtSearch.TextChanged += (s, e) => FilterAllCommands(txtSearch.Text);
            pnlAll.Controls.Add(txtSearch);

            // Col4: Separator + Add separator button (Khôi phục vị trí X=830)
            pnlSep = CreateSection("➖ NGĂN CÁCH", 830, 10, 190, 200, mainPanel);
            var btnAddSep = CreateButton("Thêm vạch ngăn cách", 5, 30, 180, 32, pnlSep, BtnAddSeparator_Click);
            var lblHint = new Label
            {
                Text = "Vạch ngăn cách (---) giúp\nphân chia lệnh trong 1 nhóm\nthành các phần nhỏ hơn.",
                Location = new Point(5, 70),
                Size = new Size(180, 60),
                ForeColor = Color.FromArgb(180, 180, 180)
            };
            pnlSep.Controls.Add(lblHint);

            // ───── ICON NHÓM (với preview) ─────
            pnlIcon = CreateSection("🎨 ICON NHÓM", 830, 220, 190, 140, mainPanel);
            var cmbIcon = new ComboBox
            {
                Location = new Point(5, 30),
                Size = new Size(135, 26),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(40, 40, 55),
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
                    UpdateIconPreview(picGroupIconPreview, icon);
                }
            };
            pnlIcon.Controls.Add(cmbIcon);

            // Preview icon nhóm
            picGroupIconPreview = new PictureBox
            {
                Location = new Point(145, 27),
                Size = new Size(32, 32),
                BackColor = Color.FromArgb(40, 40, 55),
                SizeMode = PictureBoxSizeMode.CenterImage,
                BorderStyle = BorderStyle.FixedSingle
            };
            pnlIcon.Controls.Add(picGroupIconPreview);

            // Label gợi ý
            var lblIconHint = new Label
            {
                Text = "Chọn icon type để xem\npreview icon trước khi lưu.",
                Location = new Point(5, 70),
                Size = new Size(180, 40),
                ForeColor = Color.FromArgb(180, 180, 180),
                Font = new Font("Segoe UI", 8f)
            };
            pnlIcon.Controls.Add(lblIconHint);

            // Tab thứ 3: Quản lý lệnh tắt (Alias) full form width
            pnlAliasTab = CreateSection("⌨ LỆNH TẮT TOÀN DỰ ÁN", 10, 10, 1010, 480, mainPanel);
            pnlAliasTab.Visible = false; // Hide on default tab
            
            var dgvAliases = new DataGridView
            {
                Location = new Point(10, 35),
                Size = new Size(990, 430),
                BackgroundColor = Color.FromArgb(40, 40, 55),
                ForeColor = Color.Black,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false
            };
            
            dgvAliases.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColLabel", HeaderText = "Danh sách lệnh", ReadOnly = true, FillWeight = 30 });
            dgvAliases.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColCmd", HeaderText = "Cú pháp triển khai gốc", ReadOnly = true, FillWeight = 25 });
            dgvAliases.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColAlias", HeaderText = "Thay đổi cú pháp triển khai (Lệnh tắt)", ReadOnly = false, FillWeight = 20 });
            dgvAliases.Columns.Add(new DataGridViewTextBoxColumn { Name = "ColDesc", HeaderText = "Giải thích công dụng của lệnh", ReadOnly = false, FillWeight = 40 });
            
            pnlAliasTab.Controls.Add(dgvAliases);
            _dgvAliases = dgvAliases;

            // ─── BOTTOM buttons ───
            var pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 55,
                BackColor = Color.FromArgb(22, 22, 32),
                Padding = new Padding(12, 8, 12, 8)
            };
            this.Controls.Add(pnlBottom);

            btnImport = CreateButton("📂 Import", 5, 8, 100, 36, pnlBottom, BtnImport_Click);
            btnExport = CreateButton("💾 Export", 110, 8, 100, 36, pnlBottom, BtnExport_Click);
            btnReset = CreateButton("🔄 Reset mặc định", 220, 8, 150, 36, pnlBottom, BtnReset_Click);

            btnCancel = CreateButton("✖ Hủy", 750, 8, 100, 36, pnlBottom, (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); });
            btnSave = CreateButton("✅ Lưu & Áp dụng", 860, 8, 150, 36, pnlBottom, BtnSave_Click);
            btnSave.BackColor = Color.FromArgb(27, 42, 74);
            btnSave.FlatAppearance.BorderColor = Color.FromArgb(16, 28, 52);

            // Tab for Civil/Acad/Alias
            tabMain = new TabControl
            {
                Location = new Point(12, 5),
                Size = new Size(300, 28),
                Appearance = TabAppearance.Buttons,
                ItemSize = new Size(95, 24)
            };
            tabMain.TabPages.Add("civil", "Civil Tool");
            tabMain.TabPages.Add("acad", "ACAD Tool");
            tabMain.TabPages.Add("alias", "Lệnh Tắt (Alias)");

            tabMain.SelectedIndexChanged += (s, e) => 
            {
                bool isAliasTab = tabMain.SelectedIndex == 2;
                pnlGroups.Visible = !isAliasTab;
                pnlCmds.Visible = !isAliasTab;
                pnlAll.Visible = !isAliasTab;
                pnlSep.Visible = !isAliasTab;
                pnlIcon.Visible = !isAliasTab;
                pnlAliasTab.Visible = isAliasTab;

                if (!isAliasTab)
                {
                    RefreshGroupList(); 
                    RefreshAllCommandsList();
                }
            };
            mainPanel.Controls.Add(tabMain);
            tabMain.BringToFront();
        }

        private Panel CreateSection(string title, int x, int y, int w, int h, Control parent)
        {
            var panel = new Panel
            {
                Location = new Point(x, y + 35), // offset for tab
                Size = new Size(w, h),
                BackColor = Color.FromArgb(38, 38, 50),
                BorderStyle = BorderStyle.FixedSingle
            };
            var lbl = new Label
            {
                Text = title,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.White,
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
                BackColor = Color.FromArgb(40, 40, 55),
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
                BackColor = Color.FromArgb(50, 50, 65),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 8.5f),
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(70, 70, 90);
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(46, 74, 122);
            btn.Click += onClick;
            parent.Controls.Add(btn);
            return btn;
        }

        /// <summary>
        /// Render icon preview vào PictureBox
        /// Dùng CivilRibbonTaskbar.CreateCivilIcon để tạo icon WPF rồi convert sang WinForms Image
        /// </summary>
        private void UpdateIconPreview(PictureBox pic, string iconType)
        {
            try
            {
                if (pic == null || string.IsNullOrEmpty(iconType)) return;

                // Gọi CreateCivilIcon (public static) để lấy icon WPF
                var wpfIcon = CivilRibbonTaskbar.CreateCivilIcon(iconType, 32);
                if (wpfIcon is System.Windows.Media.Imaging.BitmapSource bitmapSource)
                {
                    // Convert WPF BitmapSource → WinForms Bitmap
                    using var stream = new System.IO.MemoryStream();
                    var encoder = new System.Windows.Media.Imaging.PngBitmapEncoder();
                    encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bitmapSource));
                    encoder.Save(stream);
                    stream.Position = 0;
                    pic.Image?.Dispose();
                    pic.Image = System.Drawing.Image.FromStream(stream);
                }
            }
            catch { /* Bỏ qua lỗi render */ }
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
            CleanupInvalidCommands(_config);
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
            MenuConfig config;
            try
            {
                if (File.Exists(ConfigPath))
                {
                    string json = File.ReadAllText(ConfigPath);
                    config = JsonSerializer.Deserialize<MenuConfig>(json) ?? GetDefaultConfig();
                }
                else
                {
                    config = GetDefaultConfig();
                }
            }
            catch { config = GetDefaultConfig(); }

            // Tự động xóa lệnh không tồn tại khỏi config đã lưu
            CleanupInvalidCommands(config);
            return config;
        }

        /// <summary>
        /// Xóa các lệnh không có [CommandMethod] thực tế khỏi config,
        /// tránh lỗi "Unknown command" khi người dùng click.
        /// </summary>
        private static void CleanupInvalidCommands(MenuConfig config)
        {
            try
            {
                // Dùng reflection lấy tất cả [CommandMethod] đã đăng ký
                var validCommands = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
                {
                    try
                    {
                        foreach (var type in asm.GetTypes())
                        {
                            foreach (var method in type.GetMethods(
                                System.Reflection.BindingFlags.Public |
                                System.Reflection.BindingFlags.NonPublic |
                                System.Reflection.BindingFlags.Static |
                                System.Reflection.BindingFlags.Instance))
                            {
                                var attr = method.GetCustomAttributes(typeof(CommandMethodAttribute), false);
                                foreach (CommandMethodAttribute cmdAttr in attr)
                                {
                                    validCommands.Add(cmdAttr.GlobalName);
                                }
                            }
                        }
                    }
                    catch { /* Bỏ qua assembly không load được */ }
                }

                if (validCommands.Count == 0) return; // Không tìm được → không xóa

                // Xóa lệnh không hợp lệ khỏi mỗi group
                void CleanGroups(List<MenuGroupItem> groups)
                {
                    foreach (var group in groups)
                    {
                        group.Commands.RemoveAll(c =>
                            !string.IsNullOrEmpty(c.Command) &&
                            c.Command != "---" &&
                            !validCommands.Contains(c.Command));
                    }
                    // Xóa group trống
                    groups.RemoveAll(g => g.Commands.Count == 0 ||
                        g.Commands.All(c => string.IsNullOrEmpty(c.Command) || c.Command == "---"));
                }

                CleanGroups(config.CivilGroups);
                CleanGroups(config.AcadGroups);
            }
            catch { /* Bỏ qua lỗi reflection */ }
        }

        public static MenuConfig GetDefaultConfig()
        {
            var config = new MenuConfig();

            // ════════════════════════════════════════════════════════
            // CIVIL TOOL — Sắp xếp theo WORKFLOW Civil 3D chuyên nghiệp
            // Thứ tự: Thiết lập → Bề mặt → Tuyến → Trắc dọc → Corridor
            //       → Cọc → Trắc ngang → Khối lượng → Điểm → Cống
            //       → San nền → In ấn → Layer → Tiện ích
            // ════════════════════════════════════════════════════════
            config.CivilGroups = new List<MenuGroupItem>
            {
                // ─── 1. THIẾT LẬP & BẢN VẼ ───
                new() { GroupName = "⚙ Thiết Lập", IconType = "SETTINGS", Commands = ConvertCmds(new[]
                {
                    ("CTDS_ThietLap", "Thiết Lập Bản Vẽ", "SETTINGS"),
                    ("CTDS_SaveClean", "Lưu & Dọn File", "UTILITY"),
                    ("CT_StyleSettings", "Quản Lý Style", "SETTINGS"),
                    ("CT_StyleInfo", "Thông Tin Style", "INFO"),
                    ("---", "", ""),
                    ("CTDS_ConvertMM2M", "Chuyển MM → M", "CALCULATE"),
                    ("CTDS_ConvertCM2M", "Chuyển CM → M", "CALCULATE"),
                    ("---", "", ""),
                    ("CTL_Hub", "Layer Hub", "VIEW"),
                }) },

                // ─── 2. BỀ MẶT (SURFACE) ───
                new() { GroupName = "🗺 Bề Mặt", IconType = "SURFACE", Commands = ConvertCmds(new[]
                {
                    ("CTSu_TaoSpotElevation_OnSurface_TaiTim", "Spot Elevation Tại Tim", "ADD"),
                    ("CTS_RebuildSurface", "Rebuild Surface", "SYNC"),
                    ("---", "", ""),
                    ("CTSV_DaoDap", "Đào Đắp Surface", "VOLUME"),
                    ("CTSV_HienThiMaterialList", "Hiển Thị Material List", "INFO"),
                    ("CTSV_ChiTietMaterialSection", "Chi Tiết Material Section", "INFO"),
                    ("CTSV_PhanTichArea", "Phân Tích Area", "CALCULATE"),
                    ("CTSV_KiemTraKhoiLuong", "Kiểm Tra Khối Lượng", "CALCULATE"),
                }) },

                // ─── 3. TUYẾN (ALIGNMENT) ───
                new() { GroupName = "📏 Tuyến", IconType = "ALIGN", Commands = ConvertCmds(new[]
                {
                    ("CTC_ThongSoDuongCong_4054", "Đường Cong TCVN 4054", "CALCULATE"),
                    ("CTC_ThongSoDuongCong_13592", "Đường Cong TCVN 13592", "CALCULATE"),
                    ("CTC_KiemTraDuongCong_4054", "Kiểm Tra DC 4054", "INFO"),
                    ("CTC_KiemTraDuongCong_13592", "Kiểm Tra DC 13592", "INFO"),
                    ("CTC_TraCuuDuongCong", "Tra Cứu Đường Cong", "INFO"),
                    ("CTC_ThietLapDuongCong", "Thiết Lập Đường Cong", "SETTINGS"),
                }) },

                // ─── 4. TRẮC DỌC (PROFILE) ───
                new() { GroupName = "📈 Trắc Dọc", IconType = "PROFILE", Commands = ConvertCmds(new[]
                {
                    ("CTP_VeTracDoc_TuNhien", "Vẽ Trắc Dọc TN", "DRAW"),
                    ("CTP_VeTracDoc_TuNhien_TatCaTuyen", "Vẽ TN Tất Cả Tuyến", "DRAW"),
                    ("CTP_Fix_DuongTuNhien_TheoCoc", "Fix Đường TN Theo Cọc", "SYNC"),
                    ("---", "", ""),
                    ("CTP_GanNhanNutGiao_LenTracDoc", "Nhãn Nút Giao → Trắc Dọc", "ADD"),
                    ("CTP_TaoCogoPointTuPVI", "Tạo Point Từ PVI", "POINT"),
                    ("CTP_XuatBangCivil", "Xuất Bảng Civil", "EXPORT"),
                    ("CTP_AddParcelLabels", "Thêm Parcel Labels", "ADD"),
                }) },

                // ─── 5. CORRIDOR ───
                new() { GroupName = "🛤 Corridor", IconType = "CORRIDOR", Commands = ConvertCmds(new[]
                {
                    ("CTC_AddAllSection", "Thêm Tất Cả Section", "ADD"),
                    ("CTC_TaoCooridor_DuongDoThi_RePhai", "Corridor Rẽ Phải", "DRAW"),
                    ("CTC_TaoMViewTheoTuyen", "Tạo MView Theo Tuyến", "LAYOUT"),
                }) },

                // ─── 6. CỌC (SAMPLE LINE) ───
                new() { GroupName = "📍 Lưới Cọc", IconType = "STATION", Commands = ConvertCmds(new[]
                {
                    ("CTS_PhatSinhCoc", "⭐ Phát Sinh Cọc", "ADD"),
                    ("CTS_DoiTenCoc", "Đổi Tên Cọc", "RENAME"),
                    ("---", "", ""),
                    ("CTS_DichCoc_TinhTien", "Dịch Cọc Tịnh Tiến", "MOVE"),
                    ("CTS_DichCoc_TinhTien40", "Dịch Cọc 40m", "MOVE"),
                    ("CTS_DichCoc_TinhTien_20", "Dịch Cọc 20m", "MOVE"),
                    ("---", "", ""),
                    ("CTS_Copy_NhomCoc", "Sao Chép Nhóm Cọc", "COPY"),
                    ("CTS_DongBo_2_NhomCoc", "Đồng Bộ Nhóm Cọc", "SYNC"),
                    ("CTS_DongBo_2_NhomCoc_TheoDoan", "Đồng Bộ Theo Đoạn", "SYNC"),
                    ("---", "", ""),
                    ("CTS_Copy_BeRong_sampleLine", "Copy Bề Rộng SL", "COPY"),
                    ("CTS_Thaydoi_BeRong_sampleLine", "Thay Đổi Bề Rộng SL", "SETTINGS"),
                    ("CTS_Offset_BeRong_sampleLine", "Offset Bề Rộng SL", "MOVE"),
                    ("---", "", ""),
                    ("CTSV_ThongKeCoc", "Thống Kê Cọc (Excel)", "EXPORT"),
                }) },

                // ─── 7. TRẮC NGANG (SECTION VIEW) ───
                new() { GroupName = "📐 Trắc Ngang", IconType = "SECTION", Commands = ConvertCmds(new[]
                {
                    ("CTSV_VeTracNgangThietKe", "Tạo Trắc Ngang TK", "DRAW"),
                    ("CVSV_VeTatCa_TracNgangThietKe", "Vẽ Tất Cả TN", "DRAW"),
                    ("CTSV_ChuyenDoi_TNTK_TNTN", "Chuyển TK → TN", "SYNC"),
                    ("---", "", ""),
                    ("CTSV_DanhCap", "Đánh Cấp VHC", "SLOPE"),
                    ("CTSV_DanhCap_XoaBo", "Xóa Đánh Cấp", "DELETE"),
                    ("CTSV_DanhCap_VeThem", "Vẽ Thêm Đánh Cấp", "SLOPE"),
                    ("CTSV_DanhCap_VeThem1", "Vẽ Thêm 1m", "ADD"),
                    ("CTSV_DanhCap_VeThem2", "Vẽ Thêm 2m", "ADD"),
                    ("CTSV_DanhCap_CapNhat", "Cập Nhật Đánh Cấp", "SYNC"),
                    ("---", "", ""),
                    ("CTSV_ThemVatLieu_TrenCatNgang", "Điền Vật Liệu TN", "VOLUME"),
                    ("CTSV_ThayDoi_MSS_Min_Max", "Hiệu Chỉnh MSS", "SETTINGS"),
                    ("CTSV_ThayDoi_GioiHan_traiPhai", "Giới Hạn T/P", "SETTINGS"),
                    ("CTSV_KhoaCatNgang_AddPoint", "Khóa TN + Add Point", "ADD"),
                    ("---", "", ""),
                    ("CTSV_ThayDoi_KhungIn", "Dàn Khung In", "VIEW"),
                    ("CTSV_fit_KhungIn", "Fit Khung In", "VIEW"),
                    ("CTSV_fit_KhungIn_5_5_top", "Fit 5×5", "VIEW"),
                    ("CTSV_fit_KhungIn_5_10_top", "Fit 5×10", "VIEW"),
                    ("---", "", ""),
                    ("CTSV_An_DuongDiaChat", "Ẩn Đường Địa Chất", "VIEW"),
                    ("CTSV_HieuChinh_Section", "Hiệu Chỉnh (Static)", "SETTINGS"),
                    ("CTSV_HieuChinh_Section_Dynamic", "Hiệu Chỉnh (Dynamic)", "SETTINGS"),
                }) },

                // ─── 8. KHỐI LƯỢNG (VOLUME) ─── ⭐ ĐÃ GỘP
                new() { GroupName = "📊 Khối Lượng", IconType = "VOLUME", Commands = ConvertCmds(new[]
                {
                    ("CTSV_TinhKhoiLuongTracNgang", "⭐ Tính KL Trắc Ngang", "CALCULATE"),
                    ("---", "", ""),
                    ("CTSV_Taskbar", "Taskbar Khối Lượng", "VOLUME"),
                    ("CTSV_CaiDatBang", "Cài Đặt Bảng KL", "SETTINGS"),
                    ("CTSV_XuatCad", "Xuất KL → CAD", "CAD"),
                    ("---", "", ""),
                    ("CTSV_SoSanhSurface", "So Sánh Bề Mặt", "CALCULATE"),
                    ("CTSV_PolyArea", "KL Từ Polyline", "CALCULATE"),
                    ("CTSV_TinhDienTichPoly", "Tính DT Polyline", "CALCULATE"),
                    ("CTSV_TinhKhoiLuongPoly", "Tính KL Polyline", "CALCULATE"),
                    ("CTSV_TinhDienTichPolyExcel", "DT Poly → Excel", "EXCEL"),
                    ("CTSV_GhiDienTichPoly", "Ghi DT Poly Lên Bản Vẽ", "DRAW"),
                }) },

                // ─── 9. ĐIỂM (POINT) ───
                new() { GroupName = "📌 Điểm", IconType = "POINT", Commands = ConvertCmds(new[]
                {
                    ("CTPO_NhapToaDoTrucTiep", "⭐ Nhập Tọa Độ Trực Tiếp", "ADD"),
                    ("CTPO_TaoPointTheoBang", "Tạo Point Từ Bảng Excel", "EXCEL"),
                    ("---", "", ""),
                    ("CTPO_TaoCogoPoint_CaoDo_FromSurface", "Point Từ Surface", "SURFACE"),
                    ("CTPO_TaoCogoPoint_CaoDo_Elevationspot", "Point Từ Spot Elevation", "ADD"),
                    ("CTPO_CreateCogopointFromText", "Point Từ Text", "ADD"),
                    ("CTU_MakePointFromText", "Utility: Point Từ Text", "ADD"),
                    ("---", "", ""),
                    ("CTPO_An_CogoPoint", "Ẩn CogoPoint", "VIEW"),
                    ("CTPO_UpdateAllPointGroup", "Cập Nhật Point Group", "SYNC"),
                    ("CTPO_CPTT", "Chuyển Point → Text", "EXPORT"),
                    ("CTPo_ReorderPoints", "Sắp Xếp Lại Points", "SYNC"),
                }) },

                // ─── 10. CỐNG (PIPE & STRUCTURES) ───
                new() { GroupName = "🔧 Cống", IconType = "PIPE", Commands = ConvertCmds(new[]
                {
                    ("CTPI_ThayDoi_DuongKinhCong", "Thay Đổi ĐK Cống", "SETTINGS"),
                    ("CTPI_ThayDoi_MatPhangRef_Cong", "Mặt Phẳng Ref Cống", "SETTINGS"),
                    ("CTPI_ThayDoi_DoanDocCong", "Thay Đổi Dốc Cống", "SETTINGS"),
                    ("---", "", ""),
                    ("CTPI_BangCaoDo_TuNhienHoThu", "Bảng CĐ Tự Nhiên Hố Thu", "TABLE"),
                    ("CTPI_XoayHoThu_Theo2diem", "Xoay Hố Thu", "SYNC"),
                    ("CT_VTOADOHG", "Tọa Độ Hố Ga (DLL)", "POINT"),
                }) },

                // ─── 11. SAN NỀN (GRADING) ───
                new() { GroupName = "▦ San Nền", IconType = "GRADING", Commands = ConvertCmds(new[]
                {
                    ("CTSN_Taskbar", "Mở Taskbar SN", "VOLUME"),
                    ("---", "", ""),
                    ("CTSN_TaoLuoi", "Quản Lý Lưới", "GRADING"),
                    ("CTSN_NhapCaoDo", "Điền Cao Độ Lưới", "ADD"),
                    ("CTSN_Surface", "Lấy CĐ Surface", "SURFACE"),
                    ("CTSN_TinhKL", "Tính KL San Nền", "VOLUME"),
                    ("CTSN_XuatBang", "Xuất Bảng KL → CAD", "CAD"),
                }) },

                // ─── 12. IN ẤN & LAYOUT ───
                new() { GroupName = "🖨 In Ấn", IconType = "LAYOUT", Commands = ConvertCmds(new[]
                {
                    ("CT_PrintCenter", "In Trung Tâm", "DRAW"),
                    ("CTSV_TaoTrangInTD", "Tạo Trang In Trắc Dọc", "LAYOUT"),
                    ("---", "", ""),
                    ("CTDS_PrintAllLayouts", "In Tất Cả Layout", "EXPORT"),
                    ("CTDS_PrintCurrentLayout", "In Layout Hiện Tại", "EXPORT"),
                    ("CTDS_ExportPDF", "Xuất PDF", "EXPORT"),
                }) },

                // ─── 13. TIỆN ÍCH ───
                new() { GroupName = "🔧 Tiện Ích", IconType = "UTILITY", Commands = ConvertCmds(new[]
                {
                    ("CTU_ExportCAD2007", "Export CAD 2007", "EXPORT"),
                    ("CTU_ExplodeAEC", "Explode AEC Object", "DELETE"),
                    ("CTU_TextToMText", "Text → MText", "SYNC"),
                    ("CTU_TotalLength", "Tổng Chiều Dài", "CALCULATE"),
                    ("CTU_ExportTextCoords", "Export Tọa Độ Text", "EXPORT"),
                    ("CTU_FindIntersections", "Tìm Giao Điểm", "CALCULATE"),
                    ("CTU_AddPolylineVertices", "Thêm Đỉnh Polyline", "ADD"),
                    ("CTU_DrawTaluy", "Vẽ Ta-luy", "DRAW"),
                    ("---", "", ""),
                    ("CTU_StyleAutoOff", "Style Auto Off", "SETTINGS"),
                    ("CTU_StyleAutoOn", "Style Auto On", "SETTINGS"),
                    ("CTU_DumpObject", "Dump Object Info", "INFO"),
                    ("---", "", ""),
                    ("AT_Solid_Set_PropertySet", "Gán Property Set", "ADD"),
                    ("AT_Solid_Show_Info", "Thông Tin Solid", "INFO"),
                    ("---", "", ""),
                    ("CT_DoiIcon", "Đổi Icon Lệnh", "SETTINGS"),
                    ("CIVIL_RIBBON", "Reload Menu", "RELOAD"),
                    ("CT_MENU_SETTINGS", "Cài Đặt Menu", "SETTINGS"),
                }) },
            };

            // ══ ACAD TOOL panels ══
            config.AcadGroups = new List<MenuGroupItem>
            {
                new() { GroupName = "📏 Đo Lường", IconType = "MEASURE", Commands = ConvertCmds(new[]
                {
                    ("AT_TongDoDai_Full", "Tổng Độ Dài (Full)", "CALCULATE"),
                    ("AT_TongDoDai_Replace", "Tổng Độ Dài (Replace)", "CALCULATE"),
                    ("AT_TongDoDai_Replace2", "Tổng Độ Dài (Replace2)", "CALCULATE"),
                    ("AT_TongDoDai_Replace_CongThem", "Tổng Độ Dài (Cộng)", "CALCULATE"),
                    ("---", "", ""),
                    ("AT_TongDienTich_Full", "Tổng Diện Tích (Full)", "CALCULATE"),
                    ("AT_TongDienTich_Replace", "Tổng Diện Tích (Replace)", "CALCULATE"),
                    ("AT_TongDienTich_Replace2", "Tổng Diện Tích (Replace2)", "CALCULATE"),
                    ("AT_TongDienTich_Replace_CongThem", "Tổng Diện Tích (Cộng)", "CALCULATE"),
                }) },

                new() { GroupName = "✏ Chỉnh Sửa", IconType = "EDIT", Commands = ConvertCmds(new[]
                {
                    ("AT_TextLink", "Text Link", "ADD"),
                    ("AT_DanhSoThuTu", "Đánh Số Thứ Tự", "ADD"),
                    ("AT_XoaDoiTuong_CungLayer", "Xóa Cùng Layer", "DELETE"),
                    ("AT_XoaDoiTuong_3DSolid_Body", "Xóa 3DSolid/Body", "DELETE"),
                    ("AT_Offset_2Ben", "Offset 2 Bên", "COPY"),
                    ("AT_annotive_scale_currentOnly", "Annotative Scale", "SETTINGS"),
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
            _suppressItemCheck = true;
            foreach (var cmd in group.Commands)
            {
                if (cmd.Command == "---")
                {
                    lstCommands.Items.Add("─── ngăn cách ───", false);
                }
                else
                {
                    lstCommands.Items.Add(cmd, cmd.IsEnabled);
                }
            }
            _suppressItemCheck = false;

            // Cập nhật combo icon lệnh
            UpdateCmdIconCombo();
        }

        private void UpdateCmdIconCombo()
        {
            if (lstGroups.SelectedIndex < 0 || lstCommands.SelectedIndex < 0)
            {
                cmbCmdIcon.SelectedIndex = -1;
                cmbCmdIcon.Enabled = false;
                return;
            }

            var groups = GetCurrentGroupList();
            if (lstGroups.SelectedIndex >= groups.Count) return;

            var group = groups[lstGroups.SelectedIndex];
            int cmdIdx = lstCommands.SelectedIndex;
            if (cmdIdx >= group.Commands.Count || group.Commands[cmdIdx].Command == "---")
            {
                cmbCmdIcon.SelectedIndex = -1;
                cmbCmdIcon.Enabled = false;
                return;
            }

            cmbCmdIcon.Enabled = true;
            string currentIcon = group.Commands[cmdIdx].SubIcon;
            int idx = cmbCmdIcon.Items.IndexOf(currentIcon);
            cmbCmdIcon.SelectedIndex = idx;
        }

        private List<MenuCommandItem> _allCommandsCache = new();

        /// <summary>
        /// Tạo danh sách chú thích mặc định mô tả công dụng của các lệnh phổ biến.
        /// </summary>
        private Dictionary<string, string> GetDefaultCommandDescriptions()
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "CTS_PhatSinhCoc", "Phát sinh / đổi tên cọc tổng hợp (auto, CogoPoint, bảng)." },
                { "CTS_DichCoc_TinhTien", "Dịch cọc sang trái/phải dọc theo tuyến." },
                { "CTSV_ThayDoi_KhungIn", "Dàn khung in hàng loạt tự động." },
                { "CTSV_DanhCap", "Đánh cấp (Vát mái taluy) - VHC." },
                { "CTSN_NhapCaoDo", "Điền cao độ Text tự động từ lưới điểm." },
                { "CTSV_ThemVatLieu_TrenTracNgang", "Điền khối lượng / vật liệu tự động lên trắc ngang." },
                { "CTC_TaoCooridor_DuongDoi", "Tạo Corridor tự động cho đường rẽ phải/trái." },
                { "CTSV_An_DuongDiaChat", "Ẩn hiện đường địa chất tự động." },
                { "CTSV_CaiDatBang", "Cài đặt hiển thị cho bảng khối lượng." },
                { "CTSV_DanhCap_CapNhatKLRaBang", "Cập nhật lại Text Khối Lượng vào bảng sau khi đánh cấp." },
                { "AT_TongDienTich_Replace_CongThem", "Tính tổng diện tích nhiều đối tượng và thay thế chữ." },
                { "AT_TongDoDai_Replace_CongThem", "Cộng chiều dài nhiều đường Line/Polyline." },
                { "AT_XoaDoiTuong_CungLayer", "Chọn 1 đối tượng, xoá toàn bộ nét cùng Layer đó." },
                { "AT_XoayDoiTuong_TheoViewport", "Xoay mọi đối tượng (Text/Block) thẳng theo Viewport Layout." }
            };
        }

        private void RefreshAllCommandsList()
        {
            _allCommandsCache = GetAllAvailableCommands();
            lstAllCommands.Items.Clear();
            _dgvAliases?.Rows.Clear();

            var defaultDescriptions = GetDefaultCommandDescriptions();
            foreach (var cmd in _allCommandsCache)
            {
                lstAllCommands.Items.Add(cmd);
                
                // Add to DataGridView
                if (_dgvAliases != null)
                {
                    string alias = "";
                    if (_config.CommandAliases != null && _config.CommandAliases.TryGetValue(cmd.Command, out string existingAlias))
                        alias = existingAlias;

                    string desc = "";
                    if (_config.CommandDescriptions != null && _config.CommandDescriptions.TryGetValue(cmd.Command, out string existingDesc))
                        desc = existingDesc;
                    else if (defaultDescriptions.TryGetValue(cmd.Command, out string defDesc))
                        desc = defDesc; // Fallback to predefined descriptions

                    _dgvAliases.Rows.Add(cmd.Label, cmd.Command, alias, desc);
                }
            }
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
        /// Trả về tất cả lệnh có sẵn (gộp từ cấu hình mặc định và tất cả lệnh tìm được bằng Reflection)
        /// </summary>
        private List<MenuCommandItem> GetAllAvailableCommands()
        {
            var defaultConfig = GetDefaultConfig();
            
            // 1. Lấy tất cả lệnh từ cấu hình mặc định để lấy Label thân thiện
            var knownCommands = defaultConfig.CivilGroups.Concat(defaultConfig.AcadGroups)
                .SelectMany(g => g.Commands)
                .Where(c => !string.IsNullOrEmpty(c.Command) && c.Command != "---")
                .GroupBy(c => c.Command)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            var allCommands = new List<MenuCommandItem>();

            // 2. Quét toàn bộ project để tìm các [CommandMethod]
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                try
                {
                    foreach (var type in asm.GetTypes())
                    {
                        foreach (var method in type.GetMethods(
                            System.Reflection.BindingFlags.Public |
                            System.Reflection.BindingFlags.NonPublic |
                            System.Reflection.BindingFlags.Static |
                            System.Reflection.BindingFlags.Instance))
                        {
                            var attr = method.GetCustomAttributes(typeof(CommandMethodAttribute), false);
                            foreach (CommandMethodAttribute cmdAttr in attr)
                            {
                                string globalName = cmdAttr.GlobalName;
                                // Tránh thêm trùng lặp
                                if (!allCommands.Any(c => c.Command.Equals(globalName, StringComparison.OrdinalIgnoreCase)))
                                {
                                    if (knownCommands.TryGetValue(globalName, out var known))
                                    {
                                        allCommands.Add(new MenuCommandItem { Command = known.Command, Label = known.Label, SubIcon = known.SubIcon });
                                    }
                                    else
                                    {
                                        // Nếu lệnh chưa từng có trong menu, lấy tên lệnh làm Label
                                        allCommands.Add(new MenuCommandItem { Command = globalName, Label = globalName, SubIcon = "UTILITY" });
                                    }
                                }
                            }
                        }
                    }
                }
                catch { /* Bỏ qua assembly không load được */ }
            }

            return allCommands.OrderBy(c => c.Label).ToList();
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

        private void LstCommands_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            if (_suppressItemCheck) return;
            if (lstGroups.SelectedIndex < 0) return;
            var groups = GetCurrentGroupList();
            if (lstGroups.SelectedIndex >= groups.Count) return;
            var group = groups[lstGroups.SelectedIndex];
            if (e.Index < group.Commands.Count && group.Commands[e.Index].Command != "---")
            {
                group.Commands[e.Index].IsEnabled = (e.NewValue == CheckState.Checked);
            }
        }

        private void BtnCheckAll_Click(object? sender, EventArgs e)
        {
            if (lstGroups.SelectedIndex < 0) return;
            var groups = GetCurrentGroupList();
            if (lstGroups.SelectedIndex >= groups.Count) return;
            var group = groups[lstGroups.SelectedIndex];
            foreach (var cmd in group.Commands)
                cmd.IsEnabled = true;
            RefreshCommandList();
        }

        private void BtnUncheckAll_Click(object? sender, EventArgs e)
        {
            if (lstGroups.SelectedIndex < 0) return;
            var groups = GetCurrentGroupList();
            if (lstGroups.SelectedIndex >= groups.Count) return;
            var group = groups[lstGroups.SelectedIndex];
            foreach (var cmd in group.Commands)
                cmd.IsEnabled = false;
            RefreshCommandList();
        }

        #endregion

        #region Bottom Buttons

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            // Update aliases and descriptions from DGV
            if (_config.CommandAliases == null) _config.CommandAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (_config.CommandDescriptions == null) _config.CommandDescriptions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            
            _config.CommandAliases.Clear();
            _config.CommandDescriptions.Clear();

            if (_dgvAliases != null)
            {
                foreach (DataGridViewRow row in _dgvAliases.Rows)
                {
                    if (row.Cells["ColCmd"].Value is string cmdName)
                    {
                        if (row.Cells["ColAlias"].Value is string alias)
                        {
                            alias = alias.Trim();
                            if (!string.IsNullOrEmpty(alias))
                                _config.CommandAliases[cmdName] = alias;
                        }

                        if (row.Cells["ColDesc"].Value is string desc)
                        {
                            desc = desc.Trim();
                            if (!string.IsNullOrEmpty(desc))
                                _config.CommandDescriptions[cmdName] = desc;
                        }
                    }
                }
            }

            SaveConfig();
            
            // Reload aliases in AutoCad
            CivilToolAutoLoader.ApplyAliases(_config.CommandAliases);
            
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
                Title = "Civil Tool",
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
                Title = "ACAD Tool",
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
