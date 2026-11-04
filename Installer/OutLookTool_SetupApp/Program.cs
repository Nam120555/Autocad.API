using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;

namespace OutLookTool_SetupApp
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new SetupForm());
        }
    }

    public class SetupForm : Form
    {
        private TextBox txtInstallPath;
        private CheckBox chkDesktopShortcut;
        private CheckBox chkRunAfter;
        private Button btnInstall;
        private ProgressBar prgInstall;
        private Label lblStatus;

        public SetupForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Cài Đặt Phần Mềm Lọc & Trích Xuất Email Outlook - NXSOFT";
            this.Size = new Size(620, 480);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);
            this.BackColor = Color.FromArgb(245, 247, 250);

            // Banner Top
            Panel pnlBanner = new Panel
            {
                Dock = DockStyle.Top,
                Height = 85,
                BackColor = Color.FromArgb(24, 43, 73),
                Padding = new Padding(20, 15, 20, 10)
            };

            Label lblTitle = new Label
            {
                Text = "CÀI ĐẶT BỘ CÔNG CỤ LỌC & TRÍCH XUẤT EMAIL OUTLOOK",
                Font = new Font("Segoe UI", 12f, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(20, 15)
            };

            Label lblSub = new Label
            {
                Text = "Phiên bản 1.0.0 (Độc lập 100% - Không cần AutoCAD - An toàn dữ liệu)",
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
                ForeColor = Color.FromArgb(200, 220, 245),
                AutoSize = true,
                Location = new Point(20, 45)
            };

            pnlBanner.Controls.AddRange(new Control[] { lblTitle, lblSub });

            // Main Content Area
            Panel pnlBody = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(25, 20, 25, 10)
            };

            Label lblDesc = new Label
            {
                Text = "Trình cài đặt sẽ thiết lập phần mềm vào máy tính của bạn.\nBạn có thể đọc trực tiếp các file .ost và .pst mà không cần mở Outlook.",
                Location = new Point(25, 15),
                Size = new Size(550, 45),
                ForeColor = Color.FromArgb(50, 50, 50)
            };

            // Group Path
            GroupBox grpPath = new GroupBox
            {
                Text = "Thư mục cài đặt:",
                Location = new Point(25, 65),
                Size = new Size(550, 75),
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(24, 43, 73)
            };

            txtInstallPath = new TextBox
            {
                Text = @"C:\OutLook",
                Location = new Point(15, 30),
                Width = 410,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular)
            };

            Button btnBrowse = new Button
            {
                Text = "Duyệt...",
                Location = new Point(435, 28),
                Width = 95,
                Height = 29,
                BackColor = Color.FromArgb(230, 235, 245),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular),
                Cursor = Cursors.Hand
            };
            btnBrowse.Click += (s, e) =>
            {
                using (FolderBrowserDialog fbd = new FolderBrowserDialog())
                {
                    fbd.SelectedPath = txtInstallPath.Text;
                    if (fbd.ShowDialog() == DialogResult.OK)
                    {
                        txtInstallPath.Text = fbd.SelectedPath;
                    }
                }
            };
            grpPath.Controls.AddRange(new Control[] { txtInstallPath, btnBrowse });

            // Group Options
            GroupBox grpOptions = new GroupBox
            {
                Text = "Tùy chọn bổ sung:",
                Location = new Point(25, 150),
                Size = new Size(550, 85),
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(24, 43, 73)
            };

            chkDesktopShortcut = new CheckBox
            {
                Text = "Tạo biểu tượng lối tắt ngoài Màn hình chính (Desktop Shortcut)",
                Checked = true,
                Location = new Point(15, 25),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular)
            };

            chkRunAfter = new CheckBox
            {
                Text = "Khởi chạy phần mềm ngay sau khi cài đặt hoàn tất",
                Checked = true,
                Location = new Point(15, 52),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular)
            };

            grpOptions.Controls.AddRange(new Control[] { chkDesktopShortcut, chkRunAfter });

            prgInstall = new ProgressBar
            {
                Location = new Point(25, 245),
                Size = new Size(550, 20),
                Visible = false,
                Style = ProgressBarStyle.Marquee
            };

            lblStatus = new Label
            {
                Location = new Point(25, 270),
                Size = new Size(550, 25),
                Text = "",
                ForeColor = Color.FromArgb(0, 112, 192),
                Font = new Font("Segoe UI", 9f, FontStyle.Italic)
            };

            pnlBody.Controls.AddRange(new Control[] { lblDesc, grpPath, grpOptions, prgInstall, lblStatus });

            // Bottom Buttons
            Panel pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                BackColor = Color.FromArgb(235, 238, 242),
                Padding = new Padding(20, 10, 20, 10)
            };

            btnInstall = new Button
            {
                Text = "🚀 Tiến Hành Cài Đặt",
                Width = 180,
                Height = 40,
                Location = new Point(400, 10),
                BackColor = Color.FromArgb(0, 112, 192),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Click += BtnInstall_Click;

            Button btnCancel = new Button
            {
                Text = "Hủy Bỏ",
                Width = 100,
                Height = 40,
                Location = new Point(290, 10),
                BackColor = Color.FromArgb(200, 205, 215),
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
                Cursor = Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            btnCancel.Click += (s, e) => this.Close();

            pnlBottom.Controls.AddRange(new Control[] { btnInstall, btnCancel });

            this.Controls.Add(pnlBody);
            this.Controls.Add(pnlBanner);
            this.Controls.Add(pnlBottom);
        }

        private async void BtnInstall_Click(object? sender, EventArgs e)
        {
            string installPath = txtInstallPath.Text.Trim();
            if (string.IsNullOrEmpty(installPath))
            {
                MessageBox.Show("Vui lòng chọn thư mục cài đặt hợp lệ!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnInstall.Enabled = false;
            prgInstall.Visible = true;
            lblStatus.Text = "Đang khởi tạo thư mục và sao chép dữ liệu...";

            try
            {
                if (!Directory.Exists(installPath))
                    Directory.CreateDirectory(installPath);

                string sourceExe = Path.Combine(@"C:\OutLook", "OutLookTool.exe");
                string targetExe = Path.Combine(installPath, "OutLookTool.exe");

                if (!File.Exists(sourceExe))
                {
                    sourceExe = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutLookTool.exe");
                }

                if (File.Exists(sourceExe) && !string.Equals(sourceExe, targetExe, StringComparison.OrdinalIgnoreCase))
                {
                    File.Copy(sourceExe, targetExe, true);
                }

                // Copy batch file
                string sourceBat = @"d:\Autocad 2026_API\Autocad 2026_API\Chay_Tool_Doc_Email_OutLook.bat";
                string targetBat = Path.Combine(installPath, "Chay_Tool_Doc_Email_OutLook.bat");
                if (File.Exists(sourceBat))
                {
                    File.Copy(sourceBat, targetBat, true);
                }

                // Desktop shortcut
                if (chkDesktopShortcut.Checked)
                {
                    lblStatus.Text = "Đang tạo Shortcut ngoài màn hình Desktop...";
                    CreateShortcut(targetExe, "Phần Mềm Lọc Email Outlook");
                }

                lblStatus.Text = "Cài đặt thành công!";
                prgInstall.Visible = false;

                MessageBox.Show("Cài đặt thành công phần mềm Lọc & Trích Xuất Email Outlook!", "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (chkRunAfter.Checked && File.Exists(targetExe))
                {
                    Process.Start(new ProcessStartInfo(targetExe) { UseShellExecute = true });
                }

                this.Close();
            }
            catch (Exception ex)
            {
                prgInstall.Visible = false;
                btnInstall.Enabled = true;
                lblStatus.Text = "Lỗi cài đặt: " + ex.Message;
                MessageBox.Show("Lỗi trong quá trình cài đặt: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateShortcut(string targetPath, string shortcutName)
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string shortcutLocation = Path.Combine(desktopPath, shortcutName + ".lnk");

                // Use Windows Script Host through COM
                Type? shellType = Type.GetTypeFromProgID("WScript.Shell");
                if (shellType != null)
                {
                    dynamic? shell = Activator.CreateInstance(shellType);
                    if (shell != null)
                    {
                        dynamic shortcut = shell.CreateShortcut(shortcutLocation);
                        shortcut.TargetPath = targetPath;
                        shortcut.WorkingDirectory = Path.GetDirectoryName(targetPath);
                        shortcut.Description = "Phần Mềm Lọc & Trích Xuất Email Outlook - NXSOFT";
                        shortcut.Save();
                    }
                }
            }
            catch { }
        }
    }
}
