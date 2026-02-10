// IconCustomizationForm.cs - Form đổi icon cho từng lệnh trong Ribbon
// Cho phép user chọn emoji/ký hiệu để thay đổi icon của từng lệnh

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(Civil3DCsharp.IconCustomizationCommands))]

namespace Civil3DCsharp
{
    /// <summary>
    /// Form cho phép đổi icon của từng lệnh
    /// </summary>
    public class IconCustomizationForm : Form
    {
        private ListView listViewCommands = null!;
        private Panel panelIconPicker = null!;
        private TextBox txtSearch = null!;
        private Button btnApply = null!;
        private Button btnClose = null!;
        private Label lblPreview = null!;
        private Label lblCurrentIcon = null!;

        // Danh sách các emoji có thể dùng làm icon
        private static readonly string[] IconEmojis =
        [
            // Hình học cơ bản
            "◎", "◉", "○", "●", "◐", "◑", "◒", "◓",
            "□", "■", "▢", "▣", "◻", "◼", "◽", "◾",
            "△", "▲", "▽", "▼", "◁", "◀", "▷", "▶",
            "◇", "◆", "◈", "◊", "⬡", "⬢", "⬣", "⬠",
            // Mũi tên
            "←", "→", "↑", "↓", "↔", "↕", "↖", "↗", "↘", "↙",
            "⬅", "➡", "⬆", "⬇", "⬈", "⬉", "⬊", "⬋",
            "▸", "◂", "▴", "▾", "►", "◄", "▲", "▼",
            // Ký hiệu kỹ thuật  
            "⊕", "⊖", "⊗", "⊘", "⊙", "⊚", "⊛", "⊜",
            "⊞", "⊟", "⊠", "⊡", "⧉", "⧈", "⧇", "⧆",
            // Đường kẻ
            "═", "║", "╔", "╗", "╚", "╝", "╠", "╣",
            "━", "┃", "┏", "┓", "┗", "┛", "╱", "╲",
            // Sao và hình đặc biệt
            "★", "☆", "✦", "✧", "✪", "✫", "✬", "✭",
            "✓", "✔", "✕", "✖", "✗", "✘", "✚", "✛",
            // Công cụ và biểu tượng
            "⚙", "⚡", "⚠", "⚑", "⚐", "⛏", "⛨", "⛫",
            "⌂", "⌘", "⌥", "⎔", "⎖", "⏏", "⏚", "⏛",
            // Bảng và danh sách
            "▤", "▥", "▦", "▧", "▨", "▩", "▭", "▬",
            // Số và chữ cái
            "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩",
            "Ⓐ", "Ⓑ", "Ⓒ", "Ⓓ", "Ⓔ", "Ⓕ", "Ⓖ", "Ⓗ", "Ⓘ", "Ⓙ"
        ];

        // Dictionary lưu mapping command -> icon
        private static readonly Dictionary<string, string> _commandIcons = [];

        // Biến lưu command đang chọn
        private string? _selectedCommand;

        public IconCustomizationForm()
        {
            InitializeComponent();
            LoadCommands();
        }

        private void InitializeComponent()
        {
            this.Text = "🎨 Đổi Icon cho Lệnh - Civil Tool";
            this.Size = new Size(900, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(700, 500);

            // Panel bên trái - Danh sách lệnh
            var panelLeft = new Panel
            {
                Dock = DockStyle.Left,
                Width = 450,
                Padding = new Padding(10)
            };

            var lblTitle = new Label
            {
                Text = "📋 DANH SÁCH LỆNH",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                Dock = DockStyle.Top,
                Height = 30
            };

            txtSearch = new TextBox
            {
                Dock = DockStyle.Top,
                Height = 25,
                Font = new Font("Segoe UI", 10),
                PlaceholderText = "🔍 Tìm kiếm lệnh..."
            };
            txtSearch.TextChanged += TxtSearch_TextChanged;

            listViewCommands = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Font = new Font("Segoe UI", 10)
            };
            listViewCommands.Columns.Add("Icon", 50);
            listViewCommands.Columns.Add("Lệnh", 180);
            listViewCommands.Columns.Add("Mô tả", 200);
            listViewCommands.SelectedIndexChanged += ListViewCommands_SelectedIndexChanged;

            panelLeft.Controls.Add(listViewCommands);
            panelLeft.Controls.Add(txtSearch);
            panelLeft.Controls.Add(lblTitle);

            // Panel bên phải - Chọn icon
            var panelRight = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            var lblIconTitle = new Label
            {
                Text = "🎨 CHỌN ICON MỚI",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                Dock = DockStyle.Top,
                Height = 30
            };

            // Preview area
            var panelPreview = new Panel
            {
                Dock = DockStyle.Top,
                Height = 80,
                BackColor = Color.FromArgb(240, 240, 240)
            };

            lblCurrentIcon = new Label
            {
                Text = "Chọn lệnh để thay đổi icon",
                Font = new Font("Segoe UI", 10),
                Location = new Point(10, 10),
                AutoSize = true
            };

            lblPreview = new Label
            {
                Text = "?",
                Font = new Font("Segoe UI Symbol", 32),
                Location = new Point(10, 35),
                AutoSize = true
            };

            panelPreview.Controls.Add(lblCurrentIcon);
            panelPreview.Controls.Add(lblPreview);

            // Icon picker flowlayout
            panelIconPicker = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.White,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(5)
            };

            // Tạo các nút icon
            foreach (var emoji in IconEmojis)
            {
                var btn = new Button
                {
                    Text = emoji,
                    Width = 45,
                    Height = 45,
                    Font = new Font("Segoe UI Symbol", 18),
                    FlatStyle = FlatStyle.Flat,
                    Margin = new Padding(3),
                    Tag = emoji
                };
                btn.FlatAppearance.BorderSize = 1;
                btn.FlatAppearance.BorderColor = Color.LightGray;
                btn.Click += IconButton_Click;
                btn.MouseEnter += (s, e) => btn.BackColor = Color.FromArgb(200, 220, 255);
                btn.MouseLeave += (s, e) => btn.BackColor = SystemColors.Control;
                panelIconPicker.Controls.Add(btn);
            }

            // Buttons
            var panelButtons = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50
            };

            btnApply = new Button
            {
                Text = "✓ Áp dụng & Làm mới Ribbon",
                Width = 200,
                Height = 35,
                Location = new Point(10, 8),
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            btnApply.Click += BtnApply_Click;

            btnClose = new Button
            {
                Text = "Đóng",
                Width = 80,
                Height = 35,
                Location = new Point(220, 8),
                Font = new Font("Segoe UI", 10)
            };
            btnClose.Click += (s, e) => this.Close();

            panelButtons.Controls.Add(btnApply);
            panelButtons.Controls.Add(btnClose);

            panelRight.Controls.Add(panelIconPicker);
            panelRight.Controls.Add(panelPreview);
            panelRight.Controls.Add(lblIconTitle);
            panelRight.Controls.Add(panelButtons);

            this.Controls.Add(panelRight);
            this.Controls.Add(panelLeft);
        }

        private void LoadCommands()
        {
            listViewCommands.Items.Clear();

            // Lấy danh sách tất cả các lệnh từ các module
            var allCommands = GetAllCommands();

            foreach (var cmd in allCommands)
            {
                string icon = GetCommandIcon(cmd.Command);
                var item = new ListViewItem(icon);
                item.SubItems.Add(cmd.Command);
                item.SubItems.Add(cmd.Label);
                item.Tag = cmd.Command;
                listViewCommands.Items.Add(item);
            }
        }

        private static List<(string Command, string Label)> GetAllCommands()
        {
            // Trả về danh sách tất cả các lệnh có trong project
            return
            [
                // Surfaces
                ("CTS_TaoSpotElevation_OnSurface_TaiTim", "◎ Spot Elevation Tại Tim"),
                
                // SampleLine - Đổi tên cọc
                ("CTS_DoiTenCoc", "▸ Đổi tên cọc"),
                ("CTS_DoiTenCoc2", "▸ Đổi tên cọc đoạn"),
                ("CTS_DoiTenCoc3", "▸ Đổi tên cọc Km"),
                ("CTS_DoiTenCoc_fromCogoPoint", "▸ Đổi tên từ CogoPoint"),
                ("CTS_DoiTenCoc_TheoThuTu", "▸ Đổi tên thứ tự"),
                ("CTS_DoiTenCoc_H", "▸ Đổi tên hậu tố A"),
                
                // SampleLine - Tọa độ cọc
                ("CTS_TaoBang_ToaDoCoc", "◈ Tọa độ cọc (X,Y)"),
                ("CTS_XUATTOADOCOCSANG_SURVEYNOTES", "◈ Xuất tọa độ → Survey Notes"),
                ("CTS_TaoBang_ToaDoGoc", "◈ Tọa độ góc"),
                
                // ProfileView
                ("CTPv_TaoProfileView", "▤ Tạo Profile View"),
                ("CTPv_TaoLabel_DayDep", "◈ Label đáy đắp"),
                ("CTPv_CaiDatViewProfile", "⊙ Cài đặt View Profile"),
                ("CTPv_XoaLabel_DayDep", "◊ Xóa label đáy"),
                ("CTPv_CapNhatLabel_DayDep", "⊕ Cập nhật label đáy"),
                
                // SectionView
                ("CTSv_BandingSectionView", "▤ Banding Section View"),
                ("CTSv_TachSectionView", "◎ Tách Section View"),
                ("CTSv_XoaLabel_DayDep", "◊ Xóa label đáy"),
                
                // Drawing Setup
                ("CTDS_ThietLap", "⊙ Thiết lập bản vẽ"),
                ("CTDS_SaveClean", "◎ Lưu sạch"),
                ("CTDS_PrintAllLayouts", "▤ In tất cả"),
                ("CTDS_PrintCurrentLayout", "▤ In hiện tại"),
                ("CTDS_ExportPDF", "⬇ Xuất PDF"),
                ("CTDS_ConvertMM2M", "◈ mm → m"),
                ("CTDS_ConvertCM2M", "◈ cm → m"),
                
                // Layer Quick
                ("CTL_ToText", "T Sang Text"),
                ("CTL_ToDefpoints", "D Sang Defpoints"),
                ("CTL_ToDim", "◎ Sang Dim"),
                ("CTL_ToHatch", "▤ Sang Hatch"),
                ("CTL_ToNC", "━ Sang NC"),
                ("CTL_ToBlock", "□ Sang Block"),
                ("CTL_To0", "○ Sang Layer 0"),
                ("CTL_ToHidden", "┄ Sang Hidden"),
                ("CTL_ToCenter", "═ Sang Center"),
                ("CTL_ToPhantom", "┅ Sang Phantom"),
                
                // Common Utilities
                ("CTU_MakePointFromText", "⊕ Tạo Point từ Text"),
                ("CTU_TotalLength", "━ Tổng chiều dài"),
                ("CTU_ExportTextCoords", "⬇ Xuất tọa độ Text"),
                ("CTU_TextToMText", "▭ Text → MText"),
                ("CTU_FindIntersections", "◎ Tìm điểm giao"),
                ("CTU_AddPolylineVertices", "⊕ Thêm đỉnh Polyline"),
                ("CTU_DrawTaluy", "╱ Vẽ Taluy"),
                
                // Curve Design Standards
                ("CTC_ThietLapDuongCong", "⊙ Mở Form Đường Cong"),
                ("CTC_TraCuuDuongCong", "◎ Tra cứu thông số"),
                ("CTC_ThongSoDuongCong_4054", "▤ Bảng TCVN 4054"),
                ("CTC_ThongSoDuongCong_13592", "▤ Bảng TCVN 13592"),
                ("CTC_KiemTraDuongCong_4054", "⚙ Kiểm tra theo 4054"),
                ("CTC_KiemTraDuongCong_13592", "⚙ Kiểm tra theo 13592"),
                
                // Volume Calculation
                ("CTSV_XuatKhoiLuong", "◈ Xuất khối lượng"),
                ("CTSV_CaiDat", "⚙ Cài đặt"),
                ("CTSV_Taskbar", "▤ Taskbar")
            ];
        }

        private static string GetCommandIcon(string command)
        {
            // Trả về icon đã lưu hoặc icon mặc định
            if (_commandIcons.TryGetValue(command, out var icon))
                return icon;

            // Lấy icon mặc định từ label
            var cmd = GetAllCommands().FirstOrDefault(c => c.Command == command);
            if (!string.IsNullOrEmpty(cmd.Label))
            {
                foreach (var emoji in IconEmojis)
                {
                    if (cmd.Label.Contains(emoji))
                        return emoji;
                }
            }
            return "?";
        }

        public static void SetCommandIcon(string command, string icon)
        {
            _commandIcons[command] = icon;
        }

        private void TxtSearch_TextChanged(object? sender, EventArgs e)
        {
            string search = txtSearch.Text.ToLower();
            foreach (ListViewItem item in listViewCommands.Items)
            {
                bool visible = string.IsNullOrEmpty(search) ||
                               item.SubItems[1].Text.ToLower().Contains(search) ||
                               item.SubItems[2].Text.ToLower().Contains(search);
                item.BackColor = visible ? Color.White : Color.LightGray;
            }
        }

        private void ListViewCommands_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (listViewCommands.SelectedItems.Count > 0)
            {
                var item = listViewCommands.SelectedItems[0];
                _selectedCommand = item.Tag as string;
                lblCurrentIcon.Text = $"Lệnh: {item.SubItems[1].Text}";
                lblPreview.Text = item.SubItems[0].Text;
            }
        }

        private void IconButton_Click(object? sender, EventArgs e)
        {
            if (sender is Button btn && _selectedCommand != null)
            {
                string newIcon = btn.Tag as string ?? "?";
                lblPreview.Text = newIcon;

                // Cập nhật trong list
                foreach (ListViewItem item in listViewCommands.Items)
                {
                    if (item.Tag as string == _selectedCommand)
                    {
                        item.SubItems[0].Text = newIcon;
                        break;
                    }
                }

                // Lưu vào dictionary
                SetCommandIcon(_selectedCommand, newIcon);
            }
        }

        private void BtnApply_Click(object? sender, EventArgs e)
        {
            // Làm mới ribbon
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            doc?.SendStringToExecute("show_menu ", true, false, false);

            MessageBox.Show(
                "Đã lưu thay đổi icon!\n\nRibbon sẽ được làm mới với icons mới.",
                "Thông báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
    }

    /// <summary>
    /// Lệnh mở form đổi icon
    /// </summary>
    public class IconCustomizationCommands
    {
        [CommandMethod("CT_DoiIcon")]
        public static void CT_DoiIcon()
        {
            try
            {
                var form = new IconCustomizationForm();
                // Sử dụng ShowModalDialog thay vì ShowModelessDialog để tránh crash
                AcadApp.ShowModalDialog(form);
            }
            catch (System.Exception ex)
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                doc?.Editor.WriteMessage($"\nLỗi mở form: {ex.Message}");
            }
        }

        [CommandMethod("CT_DanhSachLenh")]
        public static void CT_DanhSachLenh()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════");
            ed.WriteMessage("\n  DANH SÁCH TẤT CẢ CÁC LỆNH CIVIL TOOL");
            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════");

            int count = 0;
            foreach (var cmd in GetAllCommandsList())
            {
                count++;
                ed.WriteMessage($"\n  {count,3}. {cmd.Command,-35} - {cmd.Label}");
            }

            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════");
            ed.WriteMessage($"\n  Tổng cộng: {count} lệnh");
            ed.WriteMessage("\n  Gõ CT_DoiIcon để mở form đổi icon cho từng lệnh");
            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════\n");
        }

        private static List<(string Command, string Label)> GetAllCommandsList()
        {
            return
            [
                ("CTS_TaoSpotElevation_OnSurface_TaiTim", "◎ Spot Elevation Tại Tim"),
                ("CTS_DoiTenCoc", "▸ Đổi tên cọc"),
                ("CTS_DoiTenCoc2", "▸ Đổi tên cọc đoạn"),
                ("CTS_DoiTenCoc3", "▸ Đổi tên cọc Km"),
                ("CTS_DoiTenCoc_fromCogoPoint", "▸ Đổi tên từ CogoPoint"),
                ("CTS_DoiTenCoc_TheoThuTu", "▸ Đổi tên thứ tự"),
                ("CTS_DoiTenCoc_H", "▸ Đổi tên hậu tố A"),
                ("CTS_TaoBang_ToaDoCoc", "◈ Tọa độ cọc (X,Y)"),
                ("CTS_XUATTOADOCOCSANG_SURVEYNOTES", "◈ Xuất tọa độ → Survey Notes"),
                ("CTS_TaoBang_ToaDoGoc", "◈ Tọa độ góc"),
                ("CTPv_TaoProfileView", "▤ Tạo Profile View"),
                ("CTPv_TaoLabel_DayDep", "◈ Label đáy đắp"),
                ("CTPv_CaiDatViewProfile", "⊙ Cài đặt View Profile"),
                ("CTPv_XoaLabel_DayDep", "◊ Xóa label đáy"),
                ("CTPv_CapNhatLabel_DayDep", "⊕ Cập nhật label đáy"),
                ("CTSv_BandingSectionView", "▤ Banding Section View"),
                ("CTSv_TachSectionView", "◎ Tách Section View"),
                ("CTSv_XoaLabel_DayDep", "◊ Xóa label đáy"),
                ("CTDS_ThietLap", "⊙ Thiết lập bản vẽ"),
                ("CTDS_SaveClean", "◎ Lưu sạch"),
                ("CTDS_PrintAllLayouts", "▤ In tất cả"),
                ("CTDS_PrintCurrentLayout", "▤ In hiện tại"),
                ("CTDS_ExportPDF", "⬇ Xuất PDF"),
                ("CTDS_ConvertMM2M", "◈ mm → m"),
                ("CTDS_ConvertCM2M", "◈ cm → m"),
                ("CTL_ToText", "T Sang Text"),
                ("CTL_ToDefpoints", "D Sang Defpoints"),
                ("CTL_ToDim", "◎ Sang Dim"),
                ("CTL_ToHatch", "▤ Sang Hatch"),
                ("CTL_ToNC", "━ Sang NC"),
                ("CTL_ToBlock", "□ Sang Block"),
                ("CTL_To0", "○ Sang Layer 0"),
                ("CTL_ToHidden", "┄ Sang Hidden"),
                ("CTL_ToCenter", "═ Sang Center"),
                ("CTL_ToPhantom", "┅ Sang Phantom"),
                ("CTU_MakePointFromText", "⊕ Tạo Point từ Text"),
                ("CTU_TotalLength", "━ Tổng chiều dài"),
                ("CTU_ExportTextCoords", "⬇ Xuất tọa độ Text"),
                ("CTU_TextToMText", "▭ Text → MText"),
                ("CTU_FindIntersections", "◎ Tìm điểm giao"),
                ("CTU_AddPolylineVertices", "⊕ Thêm đỉnh Polyline"),
                ("CTU_DrawTaluy", "╱ Vẽ Taluy"),
                ("CTC_ThietLapDuongCong", "⊙ Mở Form Đường Cong"),
                ("CTC_TraCuuDuongCong", "◎ Tra cứu thông số"),
                ("CTC_ThongSoDuongCong_4054", "▤ Bảng TCVN 4054"),
                ("CTC_ThongSoDuongCong_13592", "▤ Bảng TCVN 13592"),
                ("CTC_KiemTraDuongCong_4054", "⚙ Kiểm tra theo 4054"),
                ("CTC_KiemTraDuongCong_13592", "⚙ Kiểm tra theo 13592"),
                ("CTSV_XuatKhoiLuong", "◈ Xuất khối lượng"),
                ("CTSV_CaiDat", "⚙ Cài đặt"),
                ("CTSV_Taskbar", "▤ Taskbar"),
                ("CT_DoiIcon", "🎨 Đổi Icon lệnh"),
                ("CT_DanhSachLenh", "📋 Danh sách lệnh"),
                ("show_menu", "▤ Hiện Ribbon Menu")
            ];
        }
    }
}
