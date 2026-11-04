using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.TaskbarSettingsCommands))]

namespace MyFirstProject
{
    public class TaskbarCommandConfig
    {
        public string Command { get; set; } = "";
        public string Label { get; set; } = "";
        public bool IsEnabled { get; set; } = true;
        public string SubIcon { get; set; } = "";
    }

    public class TaskbarGroupConfig
    {
        public string GroupName { get; set; } = "";
        public string IconType { get; set; } = "";
        public List<TaskbarCommandConfig> Commands { get; set; } = new();
    }

    public class TaskbarMenuConfig
    {
        public List<TaskbarGroupConfig> CivilGroups { get; set; } = new();
    }

    /// <summary>
    /// Giao diện cài đặt Menu theo dạng TreeView (Cây thư mục)
    /// Cho phép Check/Uncheck các lệnh để hiện lên CivilToolTaskbar
    /// </summary>
    public class TaskbarSettingsForm : CivilFormBase
    {
        private TreeView treeMenu = null!;
        private TextBox txtSearch = null!;
        private Button btnSave = null!;
        private Button btnClose = null!;
        private TaskbarMenuConfig _currentConfig = null!;

        public TaskbarSettingsForm()
        {
            InitializeComponent();
            LoadConfigToTree();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("Cài đặt Menu Lệnh - Civil Tool", 550, 700);

            var mainPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };
            this.Controls.Add(mainPanel);
            mainPanel.BringToFront();

            // Header instructions
            var lblTitle = new Label
            {
                Text = "QUẢN LÝ HIỂN THỊ LỆNH (TREE VIEW)",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = NavyMedium,
                Dock = DockStyle.Top,
                Height = 30
            };

            var lblDesc = new Label
            {
                Text = "Tích chọn các lệnh bạn muốn hiển thị trên thanh Taskbar. Nhóm nào không có lệnh kích hoạt sẽ tự động bị ẩn.",
                Font = new Font("Segoe UI", 9.5f),
                ForeColor = TextDark,
                Dock = DockStyle.Top,
                Height = 40
            };

            // Toolbar (Search + CheckAll)
            var pnlToolbar = new Panel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(0, 5, 0, 5) };
            
            txtSearch = new TextBox
            {
                Width = 220,
                Font = new Font("Segoe UI", 10),
                PlaceholderText = "🔍 Tìm nhóm hoặc lệnh...",
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.Left
            };
            txtSearch.TextChanged += TxtSearch_TextChanged;

            var btnCheckAll = CreateButton("☑ Chọn Tất Cả", 230, 0, 100, 30, pnlToolbar);
            btnCheckAll.Dock = DockStyle.Right;
            btnCheckAll.Click += (s, e) => CheckAllNodes(true);

            var btnUncheckAll = CreateButton("☐ Bỏ Chọn Tất Cả", 335, 0, 115, 30, pnlToolbar);
            btnUncheckAll.Dock = DockStyle.Right;
            btnUncheckAll.Click += (s, e) => CheckAllNodes(false);

            pnlToolbar.Controls.Add(txtSearch);
            pnlToolbar.Controls.Add(btnUncheckAll);
            pnlToolbar.Controls.Add(new Panel { Width = 5, Dock = DockStyle.Right });
            pnlToolbar.Controls.Add(btnCheckAll);

            // TreeView
            treeMenu = new TreeView
            {
                Dock = DockStyle.Fill,
                CheckBoxes = true,
                Font = new Font("Segoe UI", 10.5f),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = BgInput,
                ForeColor = TextDark,
                ShowLines = false,
                ItemHeight = 28,
                FullRowSelect = true,
                Margin = new Padding(0, 10, 0, 10)
            };
            treeMenu.AfterCheck += TreeMenu_AfterCheck;

            // Footer
            var pnlFooter = CreateBottomBar();
            pnlFooter.Padding = new Padding(0, 10, 0, 0);

            btnSave = CreatePrimaryButton("LƯU CẤU HÌNH", 0, 10, 180, 40, pnlFooter);
            btnSave.Click += BtnSave_Click;

            btnClose = CreateButton("ĐÓNG", 190, 10, 100, 40, pnlFooter);
            btnClose.Click += (s, e) => this.Close();

            // Compose
            mainPanel.Controls.Add(treeMenu);
            mainPanel.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 10 });
            mainPanel.Controls.Add(pnlToolbar);
            mainPanel.Controls.Add(lblDesc);
            mainPanel.Controls.Add(lblTitle);
            mainPanel.Controls.Add(pnlFooter);
        }

        private void LoadConfigToTree(string filter = "")
        {
            treeMenu.Nodes.Clear();
            _currentConfig = LoadSavedConfig();
            
            treeMenu.BeginUpdate();
            foreach (var group in _currentConfig.CivilGroups)
            {
                bool groupMatches = string.IsNullOrEmpty(filter) || group.GroupName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;
                
                var rootNode = new TreeNode(group.GroupName);
                rootNode.Tag = group; // Lưu tham chiếu nhóm gốc

                int checkedChildCount = 0;
                int addedChildCount = 0;

                foreach (var cmd in group.Commands)
                {
                    bool cmdMatches = string.IsNullOrEmpty(filter) || 
                                      cmd.Label.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                      cmd.Command.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;
                    
                    if (!groupMatches && !cmdMatches) continue;

                    var childNode = new TreeNode($"{cmd.Label}  [{cmd.Command}]")
                    {
                        Checked = cmd.IsEnabled,
                        Tag = cmd // Lưu lại tham chiếu để khi check lưu thay đổi
                    };

                    // Nếu lệnh chứa emoji thì tô màu chữ cho giống nhóm
                    if (childNode.Text.Contains("◎") || childNode.Text.Contains("⊕") || childNode.Text.Contains("▤"))
                        childNode.ForeColor = Color.DarkSlateBlue;

                    rootNode.Nodes.Add(childNode);
                    addedChildCount++;
                    if (cmd.IsEnabled) checkedChildCount++;
                }

                // Thiết lập trạng thái check của root
                if (addedChildCount > 0)
                {
                    rootNode.Checked = checkedChildCount > 0;
                    treeMenu.Nodes.Add(rootNode);
                }
            }

            treeMenu.ExpandAll();
            treeMenu.EndUpdate();
        }

        private bool _isUpdatingCheck = false;
        private void TreeMenu_AfterCheck(object? sender, TreeViewEventArgs e)
        {
            if (_isUpdatingCheck || e.Action == TreeViewAction.Unknown) return;
            _isUpdatingCheck = true;

            try
            {
                // Nếu check node gốc (Nhóm) -> Check/Uncheck toàn bộ node con
                if (e.Node != null && e.Node.Parent == null)
                {
                    foreach (TreeNode child in e.Node.Nodes)
                    {
                        child.Checked = e.Node.Checked;
                        if (child.Tag is TaskbarCommandConfig c)
                        {
                            c.IsEnabled = e.Node.Checked;
                        }
                    }
                }
                // Nếu check node con (Lệnh)
                else if (e.Node != null)
                {
                    if (e.Node.Tag is TaskbarCommandConfig c)
                    {
                        c.IsEnabled = e.Node.Checked;
                    }

                    // Kiểm tra xem parent có nên được check mờ/check hẳn không
                    if (e.Node.Parent != null)
                    {
                        bool anyChecked = false;
                        foreach (TreeNode sibling in e.Node.Parent.Nodes)
                        {
                            if (sibling.Checked) { anyChecked = true; break; }
                        }
                        e.Node.Parent.Checked = anyChecked;
                    }
                }
            }
            finally
            {
                _isUpdatingCheck = false;
            }
        }

        private void CheckAllNodes(bool check)
        {
            _isUpdatingCheck = true;
            treeMenu.BeginUpdate();
            foreach (TreeNode root in treeMenu.Nodes)
            {
                root.Checked = check;
                foreach (TreeNode child in root.Nodes)
                {
                    child.Checked = check;
                    if (child.Tag is TaskbarCommandConfig c)
                    {
                        c.IsEnabled = check;
                    }
                }
            }
            treeMenu.EndUpdate();
            _isUpdatingCheck = false;
        }

        private void TxtSearch_TextChanged(object? sender, EventArgs e)
        {
            LoadConfigToTree(txtSearch.Text);
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            try
            {
                SaveConfig(_currentConfig);
                MessageBox.Show("Đã lưu cấu hình Menu thành công!\n\nTaskbar sẽ được làm mới.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                // Gửi lệnh đóng/mở lại taskbar nếu nó đang mở
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                doc?.SendStringToExecute("CT_Taskbar ", true, false, false);
                
                this.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi lưu cấu hình: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ======================= DATA ACCESS =======================
        public static string GetConfigFilePath()
        {
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CivilTool");
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            return Path.Combine(folder, "MenuConfig.json");
        }

        public static void SaveConfig(TaskbarMenuConfig config)
        {
            try
            {
                string json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(GetConfigFilePath(), json);
            }
            catch { }
        }

        public static TaskbarMenuConfig LoadSavedConfig()
        {
            string path = GetConfigFilePath();
            if (File.Exists(path))
            {
                try
                {
                    string json = File.ReadAllText(path);
                    var cfg = JsonSerializer.Deserialize<TaskbarMenuConfig>(json);
                    if (cfg != null && cfg.CivilGroups.Count > 0)
                        return cfg;
                }
                catch { }
            }
            return CreateDefaultConfig();
        }

        public static TaskbarMenuConfig CreateDefaultConfig()
        {
            var config = new TaskbarMenuConfig();

            // 1. Settings
            var grpSettings = new TaskbarGroupConfig { GroupName = "⚙ Settings", IconType = "UTILS" };
            grpSettings.Commands.Add(new TaskbarCommandConfig { Command = "CTDS_ThietLap", Label = "⊙ Thiết lập bản vẽ", IsEnabled = true });
            grpSettings.Commands.Add(new TaskbarCommandConfig { Command = "CTDS_SaveClean", Label = "◎ Lưu sạch", IsEnabled = true });
            grpSettings.Commands.Add(new TaskbarCommandConfig { Command = "CTDS_ConvertMM2M", Label = "◈ mm → m", IsEnabled = true });
            grpSettings.Commands.Add(new TaskbarCommandConfig { Command = "CT_TASKBAR_SETTINGS", Label = "▤ Cài đặt Menu Cây", IsEnabled = true });
            config.CivilGroups.Add(grpSettings);

            // 2. Surfaces
            var grpSurface = new TaskbarGroupConfig { GroupName = "🗺 Surfaces", IconType = "SURFACE" };
            grpSurface.Commands.Add(new TaskbarCommandConfig { Command = "CTS_TaoSpotElevation_OnSurface_TaiTim", Label = "◎ Spot Elevation Tại Tim", IsEnabled = true });
            config.CivilGroups.Add(grpSurface);

            // 3. Alignments
            var grpAlign = new TaskbarGroupConfig { GroupName = "📏 Alignments", IconType = "ALIGN" };
            grpAlign.Commands.Add(new TaskbarCommandConfig { Command = "CTC_ThietLapDuongCong", Label = "⊙ Mở Form Đường Cong", IsEnabled = true });
            grpAlign.Commands.Add(new TaskbarCommandConfig { Command = "CTC_TraCuuDuongCong", Label = "◎ Tra cứu thông số", IsEnabled = true });
            grpAlign.Commands.Add(new TaskbarCommandConfig { Command = "CTC_ThongSoDuongCong_4054", Label = "▤ Bảng TCVN 4054", IsEnabled = true });
            // grpAlign.Commands.Add(new TaskbarCommandConfig { Command = "CTC_KiemTraDuongCong_4054", Label = "Kiểm tra theo 4054", IsEnabled = true });
            config.CivilGroups.Add(grpAlign);

            // 4. Profiles
            var grpProfile = new TaskbarGroupConfig { GroupName = "📈 Profiles", IconType = "PROFILE" };
            grpProfile.Commands.Add(new TaskbarCommandConfig { Command = "CTPv_TaoProfileView", Label = "▤ Tạo Profile View", IsEnabled = true });
            grpProfile.Commands.Add(new TaskbarCommandConfig { Command = "CTPv_CaiDatViewProfile", Label = "⊙ Cài đặt View Profile", IsEnabled = true });
            grpProfile.Commands.Add(new TaskbarCommandConfig { Command = "CTPv_TaoLabel_DayDep", Label = "◈ Label đáy đắp", IsEnabled = true });
            config.CivilGroups.Add(grpProfile);

            // 5. Sample Lines
            var grpSample = new TaskbarGroupConfig { GroupName = "📍 Sample Lines", IconType = "STATION" };
            grpSample.Commands.Add(new TaskbarCommandConfig { Command = "CTS_PhatSinhCoc", Label = "▸ Phát sinh / Đổi tên cọc", IsEnabled = true });
            grpSample.Commands.Add(new TaskbarCommandConfig { Command = "CTS_DoiTenCoc_fromCogoPoint", Label = "▸ Đổi tên từ CogoPoint", IsEnabled = true });
            grpSample.Commands.Add(new TaskbarCommandConfig { Command = "CTS_TaoBang_ToaDoCoc", Label = "◈ Tọa độ cọc (X,Y)", IsEnabled = true });
            config.CivilGroups.Add(grpSample);

            // 6. Section Views
            var grpSection = new TaskbarGroupConfig { GroupName = "📐 Section Views", IconType = "SECTION" };
            grpSection.Commands.Add(new TaskbarCommandConfig { Command = "CTSv_BandingSectionView", Label = "▤ Banding Section View", IsEnabled = true });
            grpSection.Commands.Add(new TaskbarCommandConfig { Command = "CTSv_TachSectionView", Label = "◎ Tách Section View", IsEnabled = true });
            config.CivilGroups.Add(grpSection);

            // 7. Quantities
            var grpVolumes = new TaskbarGroupConfig { GroupName = "📊 Quantities", IconType = "VOLUME" };
            grpVolumes.Commands.Add(new TaskbarCommandConfig { Command = "CTSV_XuatKhoiLuong", Label = "◈ Xuất khối lượng", IsEnabled = true });
            grpVolumes.Commands.Add(new TaskbarCommandConfig { Command = "CTSV_CaiDat", Label = "Cài đặt diện tích", IsEnabled = true });
            config.CivilGroups.Add(grpVolumes);

            // 8. Points
            var grpPoint = new TaskbarGroupConfig { GroupName = "📌 Points", IconType = "POINT" };
            grpPoint.Commands.Add(new TaskbarCommandConfig { Command = "CTPO_NhapToaDoTrucTiep", Label = "⭐ Direct Point Entry", IsEnabled = true });
            grpPoint.Commands.Add(new TaskbarCommandConfig { Command = "CTPO_CreateCogopointFromText", Label = "Create Point from Text", IsEnabled = true });
            config.CivilGroups.Add(grpPoint);

            // 9. Utilities & Layer
            var grpUtil = new TaskbarGroupConfig { GroupName = "🔧 Utilities", IconType = "UTILS" };
            grpUtil.Commands.Add(new TaskbarCommandConfig { Command = "CTU_TotalLength", Label = "━ Tổng chiều dài", IsEnabled = true });
            grpUtil.Commands.Add(new TaskbarCommandConfig { Command = "CTL_ToText", Label = "T Sang Text", IsEnabled = true });
            grpUtil.Commands.Add(new TaskbarCommandConfig { Command = "CTL_ToDim", Label = "◎ Sang Dim", IsEnabled = true });
            grpUtil.Commands.Add(new TaskbarCommandConfig { Command = "CT_DoiIcon", Label = "Đổi Icon lệnh", IsEnabled = true });
            config.CivilGroups.Add(grpUtil);

            return config;
        }
    }

    public class TaskbarSettingsCommands
    {
        [CommandMethod("CT_TASKBAR_SETTINGS")]
        public static void OpenTaskbarSettings()
        {
            var form = new TaskbarSettingsForm();
            AcadApp.ShowModalDialog(form);
        }
    }
}
