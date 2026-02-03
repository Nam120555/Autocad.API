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

            // 10. Tài khoản (Account)
            this.Controls.Add(CreateDropdownButton("👤", "Tài khoản", x, y, btnWidth, btnHeight,
                Color.FromArgb(60, 60, 60), GetAccountCommands()));
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

        private List<(string Name, string Command)> GetAccountCommands()
        {
            return new List<(string, string)>
            {
                ("📍 Tọa độ hố ga", "CT_VTOADOHG"),
                ("─────────────", ""),
                ("🔑 Đăng nhập", ""),
                ("ℹ Thông tin", ""),
                ("📖 Hướng dẫn", ""),
                ("📞 Liên hệ", ""),
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

        #endregion
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
    }
}
