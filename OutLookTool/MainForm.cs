using System;
using System.IO;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Text.Json;
using ClosedXML.Excel;
using XstReader;

namespace OutLookTool
{
    public class MainForm : Form
    {
        // Core Data
        private List<EmailRecord> _allEmails = new List<EmailRecord>();
        private List<EmailRecord> _filteredEmails = new List<EmailRecord>();
        private List<string> _loadedFiles = new List<string>();

        // Paths
        private static readonly string OutlookDataDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Outlook");
        private const string ArchivePstPath = @"C:\OutLook\QLTKDS01-ARCHIVE.pst";
        private const string OutputFolder = @"C:\OutLook";
        private const string TempCopyFolder = @"C:\OutLook\_TempRead";
        private const string DefaultCacheFile = @"C:\OutLook\EmailData_Cache.json";

        // Controls
        private Button btnLoadCache;
        private Button btnSaveCache;
        private Button btnScanAll;
        private Button btnOpenSingle;
        private Button btnExportFiltered;
        private Button btnExportAll;
        private Button btnOpenOutputFolder;

        private DateTimePicker dtpFrom;
        private DateTimePicker dtpTo;
        private CheckBox chkAllDates;

        private ComboBox cboUnit;
        private ComboBox cboTopic;
        private ComboBox cboImportance;
        private ComboBox cboAttachment;
        private ComboBox cboSourceFile;

        private TextBox txtSearchKeyword;
        private Button btnResetFilter;

        private Label lblKpiTotal;
        private Label lblKpiFiltered;
        private Label lblKpiSenders;
        private Label lblKpiAttachments;

        private DataGridView dgvEmails;
        private ProgressBar progressBar;
        private Label lblStatus;

        public MainForm()
        {
            InitializeComponent();
            LoadInitialDataAsync();
        }

        private void InitializeComponent()
        {
            this.Text = "PHẦN MỀM LỌC & TRÍCH XUẤT EMAIL OUTLOOK - Tác Giả: Huỳnh Ngọc Nam (NXSOFT)";
            this.Size = new Size(1360, 850);
            this.MinimumSize = new Size(1150, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);
            this.BackColor = Color.FromArgb(245, 247, 250);

            // ==========================================
            // 1. TOP ACTION PANEL (Thanh Công Cụ Chính)
            // ==========================================
            Panel pnlTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 62,
                BackColor = Color.FromArgb(24, 43, 73),
                Padding = new Padding(10, 10, 10, 10)
            };

            FlowLayoutPanel flpTop = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = true
            };

            btnLoadCache = CreateButton("⚡ Mở File Dữ Liệu Đã Lưu (Cache)", Color.FromArgb(40, 167, 69), 235);
            btnLoadCache.Click += async (s, e) => await ChooseAndLoadCacheFileAsync();

            btnSaveCache = CreateButton("💾 Lưu File Dữ Liệu Này (Cache)", Color.FromArgb(23, 162, 184), 220);
            btnSaveCache.Click += (s, e) => SaveCurrentDataToCache();

            btnScanAll = CreateButton("📂 Quét Lại Từ Gốc (.ost/.pst)", Color.FromArgb(0, 112, 192), 215);
            btnScanAll.Click += async (s, e) => await ScanAllFilesAsync();

            btnOpenSingle = CreateButton("📄 Chọn 1 File OST/PST...", Color.FromArgb(41, 86, 129), 180);
            btnOpenSingle.Click += async (s, e) => await OpenSingleFileAsync();

            btnExportFiltered = CreateButton("📊 Xuất Excel (Đang lọc)", Color.FromArgb(255, 136, 0), 180);
            btnExportFiltered.Click += (s, e) => ExportExcel(true);

            btnExportAll = CreateButton("📑 Xuất Excel (Tất cả)", Color.FromArgb(112, 48, 160), 165);
            btnExportAll.Click += (s, e) => ExportExcel(false);

            btnOpenOutputFolder = CreateButton("📁 Thư mục", Color.FromArgb(108, 117, 125), 100);
            btnOpenOutputFolder.Click += (s, e) => {
                if (!Directory.Exists(OutputFolder)) Directory.CreateDirectory(OutputFolder);
                Process.Start(new ProcessStartInfo("explorer.exe", OutputFolder) { UseShellExecute = true });
            };

            flpTop.Controls.AddRange(new Control[] {
                btnLoadCache, btnSaveCache, btnScanAll, btnOpenSingle,
                btnExportFiltered, btnExportAll, btnOpenOutputFolder
            });
            pnlTop.Controls.Add(flpTop);

            // ==========================================
            // 2. FILTER BOX (Bộ Lọc Giao Diện Chuẩn)
            // ==========================================
            GroupBox grpFilter = new GroupBox
            {
                Text = "  🔍 BỘ LỌC & TÌM KIẾM DỮ LIỆU EMAIL CHI TIẾT  ",
                Dock = DockStyle.Top,
                Height = 145,
                Padding = new Padding(12, 10, 12, 10),
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(24, 43, 73),
                BackColor = Color.White
            };

            // Hàng 1
            Panel pnlRow1 = new Panel { Dock = DockStyle.Top, Height = 42 };

            Label lblDate = new Label { Text = "Khoảng ngày:", AutoSize = true, Location = new Point(5, 10), Font = new Font("Segoe UI", 9f, FontStyle.Regular), ForeColor = Color.Black };
            dtpFrom = new DateTimePicker { Format = DateTimePickerFormat.Short, Width = 95, Location = new Point(90, 7), Font = new Font("Segoe UI", 9f, FontStyle.Regular), Value = DateTime.Today.AddMonths(-3) };
            Label lblTo = new Label { Text = "-", AutoSize = true, Location = new Point(190, 10), Font = new Font("Segoe UI", 9f, FontStyle.Regular), ForeColor = Color.Black };
            dtpTo = new DateTimePicker { Format = DateTimePickerFormat.Short, Width = 95, Location = new Point(205, 7), Font = new Font("Segoe UI", 9f, FontStyle.Regular), Value = DateTime.Today };
            chkAllDates = new CheckBox { Text = "Tất cả ngày", Checked = true, Location = new Point(310, 8), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Regular), ForeColor = Color.Black };

            dtpFrom.Enabled = false;
            dtpTo.Enabled = false;
            dtpFrom.ValueChanged += (s, e) => ApplyFilter();
            dtpTo.ValueChanged += (s, e) => ApplyFilter();
            chkAllDates.CheckedChanged += (s, e) => {
                dtpFrom.Enabled = !chkAllDates.Checked;
                dtpTo.Enabled = !chkAllDates.Checked;
                ApplyFilter();
            };

            Label lblUnit = new Label { Text = "Đơn vị / Cơ quan:", AutoSize = true, Location = new Point(420, 10), Font = new Font("Segoe UI", 9f, FontStyle.Regular), ForeColor = Color.Black };
            cboUnit = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(535, 7), Width = 230, Font = new Font("Segoe UI", 9f, FontStyle.Regular) };
            cboUnit.Items.Add("-- Tất cả đơn vị --");
            cboUnit.SelectedIndex = 0;
            cboUnit.SelectedIndexChanged += (s, e) => ApplyFilter();

            Label lblTopic = new Label { Text = "Hạng mục công việc:", AutoSize = true, Location = new Point(785, 10), Font = new Font("Segoe UI", 9f, FontStyle.Regular), ForeColor = Color.Black };
            cboTopic = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(915, 7), Width = 220, Font = new Font("Segoe UI", 9f, FontStyle.Regular) };
            cboTopic.Items.Add("-- Tất cả hạng mục --");
            cboTopic.SelectedIndex = 0;
            cboTopic.SelectedIndexChanged += (s, e) => ApplyFilter();

            Label lblSource = new Label { Text = "Nguồn file:", AutoSize = true, Location = new Point(1150, 10), Font = new Font("Segoe UI", 9f, FontStyle.Regular), ForeColor = Color.Black };
            cboSourceFile = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(1220, 7), Width = 110, Font = new Font("Segoe UI", 9f, FontStyle.Regular) };
            cboSourceFile.Items.Add("-- Tất cả --");
            cboSourceFile.SelectedIndex = 0;
            cboSourceFile.SelectedIndexChanged += (s, e) => ApplyFilter();

            pnlRow1.Controls.AddRange(new Control[] {
                lblDate, dtpFrom, lblTo, dtpTo, chkAllDates,
                lblUnit, cboUnit,
                lblTopic, cboTopic,
                lblSource, cboSourceFile
            });

            // Hàng 2
            Panel pnlRow2 = new Panel { Dock = DockStyle.Top, Height = 42, Padding = new Padding(0, 5, 0, 0) };

            Label lblSearch = new Label { Text = "Tìm từ khóa:", AutoSize = true, Location = new Point(5, 10), Font = new Font("Segoe UI", 9f, FontStyle.Regular), ForeColor = Color.Black };
            txtSearchKeyword = new TextBox { Location = new Point(90, 7), Width = 315, Font = new Font("Segoe UI", 9.5f, FontStyle.Regular), PlaceholderText = "Nhập tiêu đề, người gửi, người nhận, nội dung..." };
            txtSearchKeyword.TextChanged += (s, e) => ApplyFilter();

            Label lblImp = new Label { Text = "Mức độ:", AutoSize = true, Location = new Point(420, 10), Font = new Font("Segoe UI", 9f, FontStyle.Regular), ForeColor = Color.Black };
            cboImportance = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(475, 7), Width = 135, Font = new Font("Segoe UI", 9f, FontStyle.Regular) };
            cboImportance.Items.AddRange(new object[] { "-- Tất cả --", "Cao (High)", "Binh thuong", "Thap" });
            cboImportance.SelectedIndex = 0;
            cboImportance.SelectedIndexChanged += (s, e) => ApplyFilter();

            Label lblAttach = new Label { Text = "Đính kèm:", AutoSize = true, Location = new Point(625, 10), Font = new Font("Segoe UI", 9f, FontStyle.Regular), ForeColor = Color.Black };
            cboAttachment = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(695, 7), Width = 130, Font = new Font("Segoe UI", 9f, FontStyle.Regular) };
            cboAttachment.Items.AddRange(new object[] { "-- Tất cả --", "Có đính kèm", "Không đính kèm" });
            cboAttachment.SelectedIndex = 0;
            cboAttachment.SelectedIndexChanged += (s, e) => ApplyFilter();

            btnResetFilter = new Button
            {
                Text = "🔄 Xóa bộ lọc",
                Location = new Point(840, 6),
                Width = 110,
                Height = 29,
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            btnResetFilter.FlatAppearance.BorderSize = 0;
            btnResetFilter.Click += (s, e) => ResetFilter();

            pnlRow2.Controls.AddRange(new Control[] {
                lblSearch, txtSearchKeyword,
                lblImp, cboImportance,
                lblAttach, cboAttachment,
                btnResetFilter
            });

            grpFilter.Controls.Add(pnlRow2);
            grpFilter.Controls.Add(pnlRow1);

            // ==========================================
            // 3. KPI STATUS BAR
            // ==========================================
            Panel pnlKpi = new Panel
            {
                Dock = DockStyle.Top,
                Height = 38,
                BackColor = Color.FromArgb(235, 240, 248),
                Padding = new Padding(12, 5, 12, 5)
            };

            FlowLayoutPanel flpKpi = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };

            lblKpiTotal = CreateKpiBadge("Tổng email:", "0", Color.FromArgb(24, 43, 73));
            lblKpiFiltered = CreateKpiBadge("Đang hiển thị:", "0", Color.FromArgb(0, 112, 192));
            lblKpiSenders = CreateKpiBadge("Người gửi:", "0", Color.FromArgb(38, 128, 0));
            lblKpiAttachments = CreateKpiBadge("Có đính kèm:", "0", Color.FromArgb(112, 48, 160));

            flpKpi.Controls.AddRange(new Control[] { lblKpiTotal, lblKpiFiltered, lblKpiSenders, lblKpiAttachments });
            pnlKpi.Controls.Add(flpKpi);

            // ==========================================
            // 4. DATA GRID VIEW
            // ==========================================
            dgvEmails = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                EnableHeadersVisualStyles = false
            };
            dgvEmails.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(24, 43, 73);
            dgvEmails.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvEmails.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            dgvEmails.ColumnHeadersHeight = 34;
            dgvEmails.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 248, 252);
            dgvEmails.RowTemplate.Height = 28;
            dgvEmails.CellDoubleClick += DgvEmails_CellDoubleClick;

            SetupGridColumns();

            // ==========================================
            // 5. BOTTOM STATUS BAR
            // ==========================================
            Panel pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 32,
                BackColor = Color.FromArgb(240, 242, 245),
                Padding = new Padding(12, 6, 12, 6)
            };

            lblStatus = new Label
            {
                Text = "Sẵn sàng.",
                Dock = DockStyle.Left,
                AutoSize = true,
                ForeColor = Color.FromArgb(60, 60, 60),
                Font = new Font("Segoe UI", 9f, FontStyle.Regular)
            };

            Label lblAuthor = new Label
            {
                Text = "👨‍💻 Tác giả: Huỳnh Ngọc Nam (NXSOFT)  |  ",
                Dock = DockStyle.Right,
                AutoSize = true,
                ForeColor = Color.FromArgb(24, 43, 73),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };

            progressBar = new ProgressBar
            {
                Dock = DockStyle.Right,
                Width = 260,
                Visible = false,
                Style = ProgressBarStyle.Marquee
            };

            pnlBottom.Controls.Add(lblStatus);
            pnlBottom.Controls.Add(lblAuthor);
            pnlBottom.Controls.Add(progressBar);

            // Add all
            this.Controls.Add(dgvEmails);
            this.Controls.Add(pnlKpi);
            this.Controls.Add(grpFilter);
            this.Controls.Add(pnlTop);
            this.Controls.Add(pnlBottom);
        }

        private Button CreateButton(string text, Color bg, int width)
        {
            Button btn = new Button
            {
                Text = text,
                Width = width,
                Height = 42,
                BackColor = bg,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 8, 0)
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private Label CreateKpiBadge(string title, string val, Color color)
        {
            Label lbl = new Label
            {
                Text = $"{title} {val}  |  ",
                AutoSize = true,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                ForeColor = color,
                Margin = new Padding(0, 2, 15, 0)
            };
            return lbl;
        }

        private void SetupGridColumns()
        {
            dgvEmails.Columns.Clear();
            dgvEmails.Columns.Add(new DataGridViewTextBoxColumn { Name = "STT", HeaderText = "STT", Width = 55 });
            dgvEmails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ngay", HeaderText = "Ngày", Width = 95 });
            dgvEmails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Gio", HeaderText = "Giờ", Width = 75 });
            dgvEmails.Columns.Add(new DataGridViewTextBoxColumn { Name = "NguoiGui", HeaderText = "Người gửi", Width = 180 });
            dgvEmails.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "Đơn vị / Cơ quan", Width = 190 });
            dgvEmails.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChuDe", HeaderText = "Chủ đề email", Width = 350 });
            dgvEmails.Columns.Add(new DataGridViewTextBoxColumn { Name = "HangMuc", HeaderText = "Hạng mục", Width = 180 });
            dgvEmails.Columns.Add(new DataGridViewTextBoxColumn { Name = "MucDo", HeaderText = "Mức độ", Width = 90 });
            dgvEmails.Columns.Add(new DataGridViewTextBoxColumn { Name = "DinhKem", HeaderText = "Đính kèm", Width = 85 });
            dgvEmails.Columns.Add(new DataGridViewTextBoxColumn { Name = "NguonFile", HeaderText = "Nguồn file", Width = 180 });
        }

        // ==========================================
        // CACHE & DATA LOADING LOGIC (MỞ NHANH 0.5s)
        // ==========================================
        private async void LoadInitialDataAsync()
        {
            // Kiểm tra nếu có file cache sẵn thì mở ngay lập tức trong 0.5s
            if (File.Exists(DefaultCacheFile))
            {
                await LoadFromCacheFileAsync(DefaultCacheFile);
            }
            else
            {
                // Chưa có cache -> quét lần đầu
                await ScanAllFilesAsync();
            }
        }

        private async Task ChooseAndLoadCacheFileAsync()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Chọn file dữ liệu đã lưu (Cache .json) để mở tức thì";
                ofd.InitialDirectory = OutputFolder;
                ofd.Filter = "Email Cache Files (*.json)|*.json|All files (*.*)|*.*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    await LoadFromCacheFileAsync(ofd.FileName);
                }
            }
        }

        private async Task LoadFromCacheFileAsync(string jsonPath)
        {
            SetLoading(true, $"Đang nạp dữ liệu từ file cache: {Path.GetFileName(jsonPath)}...");

            try
            {
                List<EmailRecord>? loaded = null;
                await Task.Run(() =>
                {
                    string json = File.ReadAllText(jsonPath);
                    loaded = JsonSerializer.Deserialize<List<EmailRecord>>(json);
                });

                if (loaded != null && loaded.Count > 0)
                {
                    _allEmails = loaded;
                    _loadedFiles = _allEmails.Select(e => e.NguonFile).Distinct().ToList();

                    PopulateFilterDropdowns();
                    ApplyFilter();

                    SetLoading(false, $"⚡ Đã nạp cực nhanh {_allEmails.Count:N0} email từ file cache ({Path.GetFileName(jsonPath)})!");
                    return;
                }
            }
            catch (Exception ex)
            {
                SetLoading(false, "Lỗi đọc file cache: " + ex.Message);
            }
        }

        private void SaveCurrentDataToCache()
        {
            if (_allEmails == null || _allEmails.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu email nào để lưu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Title = "Lưu kho dữ liệu Email thành file (.json) để lần sau mở tức thì";
                sfd.InitialDirectory = OutputFolder;
                sfd.FileName = $"EmailData_Cache_{DateTime.Now:yyyyMMdd}.json";
                sfd.Filter = "Email Cache Files (*.json)|*.json";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string json = JsonSerializer.Serialize(_allEmails, new JsonSerializerOptions { WriteIndented = true });
                        File.WriteAllText(sfd.FileName, json);

                        // Đồng thời ghi đè vào cache mặc định
                        File.WriteAllText(DefaultCacheFile, json);

                        MessageBox.Show($"Đã lưu thành công {_allEmails.Count:N0} email vào file:\n\n{sfd.FileName}\n\nTừ lần sau khi mở app, phần mềm sẽ nạp trực tiếp file này chỉ mất 1 giây!",
                                        "Lưu Cache Thành Công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Lỗi khi lưu file cache: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private async Task ScanAllFilesAsync()
        {
            SetLoading(true, "Đang quét trực tiếp các file Outlook .ost và .pst trong máy...");

            _allEmails.Clear();
            _loadedFiles.Clear();

            List<string> filesToRead = new List<string>();

            if (Directory.Exists(OutlookDataDir))
            {
                foreach (var ost in Directory.GetFiles(OutlookDataDir, "*.ost"))
                {
                    filesToRead.Add(ost);
                }
            }

            if (File.Exists(ArchivePstPath))
            {
                filesToRead.Add(ArchivePstPath);
            }

            if (Directory.Exists(OutputFolder))
            {
                foreach (var pst in Directory.GetFiles(OutputFolder, "*.pst"))
                {
                    if (!filesToRead.Contains(pst, StringComparer.OrdinalIgnoreCase))
                        filesToRead.Add(pst);
                }
            }

            if (filesToRead.Count == 0)
            {
                SetLoading(false, "Không tìm thấy file Outlook nào.");
                MessageBox.Show("Không tìm thấy file dữ liệu Outlook (.ost hoặc .pst) nào trong máy!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            await Task.Run(() =>
            {
                foreach (var filePath in filesToRead)
                {
                    string fileName = Path.GetFileName(filePath);
                    _loadedFiles.Add(fileName);

                    List<EmailRecord> emails;
                    try
                    {
                        emails = ReadOutlookFile(filePath, fileName);
                    }
                    catch
                    {
                        emails = ReadWithTempCopy(filePath, fileName);
                    }
                    _allEmails.AddRange(emails);
                }
            });

            _allEmails = _allEmails.OrderByDescending(e => e.DateObj ?? DateTime.MinValue).ToList();
            for (int i = 0; i < _allEmails.Count; i++) _allEmails[i].STT = i + 1;

            // Tự động lưu Cache để lần 2 mở lên có ngay lập tức!
            try
            {
                if (!Directory.Exists(OutputFolder)) Directory.CreateDirectory(OutputFolder);
                string json = JsonSerializer.Serialize(_allEmails);
                File.WriteAllText(DefaultCacheFile, json);
            }
            catch { }

            PopulateFilterDropdowns();
            ApplyFilter();

            SetLoading(false, $"Đã nạp xong {_allEmails.Count:N0} email từ {_loadedFiles.Count} file dữ liệu (Đã tự lưu Cache).");
        }

        private async Task OpenSingleFileAsync()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Chọn file Outlook (.ost hoặc .pst) để đọc";
                ofd.Filter = "Outlook Data Files (*.ost;*.pst)|*.ost;*.pst|All files (*.*)|*.*";
                if (ofd.ShowDialog() != DialogResult.OK) return;

                string filePath = ofd.FileName;
                string fileName = Path.GetFileName(filePath);

                SetLoading(true, $"Đang đọc dữ liệu từ {fileName}...");

                _allEmails.Clear();
                _loadedFiles.Clear();
                _loadedFiles.Add(fileName);

                await Task.Run(() =>
                {
                    List<EmailRecord> emails;
                    try
                    {
                        emails = ReadOutlookFile(filePath, fileName);
                    }
                    catch
                    {
                        emails = ReadWithTempCopy(filePath, fileName);
                    }
                    _allEmails.AddRange(emails);
                });

                _allEmails = _allEmails.OrderByDescending(e => e.DateObj ?? DateTime.MinValue).ToList();
                for (int i = 0; i < _allEmails.Count; i++) _allEmails[i].STT = i + 1;

                PopulateFilterDropdowns();
                ApplyFilter();

                SetLoading(false, $"Đã đọc xong {_allEmails.Count:N0} email từ {fileName}.");
            }
        }

        private void PopulateFilterDropdowns()
        {
            cboUnit.Items.Clear();
            cboUnit.Items.Add("-- Tất cả đơn vị --");
            var units = _allEmails.Select(e => e.DonVi).Where(u => !string.IsNullOrEmpty(u)).Distinct().OrderBy(u => u);
            foreach (var u in units) cboUnit.Items.Add(u);
            cboUnit.SelectedIndex = 0;

            cboTopic.Items.Clear();
            cboTopic.Items.Add("-- Tất cả hạng mục --");
            var topics = _allEmails.Select(e => e.HangMuc).Where(t => !string.IsNullOrEmpty(t)).Distinct().OrderBy(t => t);
            foreach (var t in topics) cboTopic.Items.Add(t);
            cboTopic.SelectedIndex = 0;

            cboSourceFile.Items.Clear();
            cboSourceFile.Items.Add("-- Tất cả --");
            foreach (var f in _loadedFiles) cboSourceFile.Items.Add(f);
            cboSourceFile.SelectedIndex = 0;
        }

        private void ApplyFilter()
        {
            if (_allEmails == null) return;

            string keyword = txtSearchKeyword.Text.Trim().ToLower();
            string selectedUnit = cboUnit.SelectedIndex > 0 ? cboUnit.SelectedItem?.ToString() ?? "" : "";
            string selectedTopic = cboTopic.SelectedIndex > 0 ? cboTopic.SelectedItem?.ToString() ?? "" : "";
            string selectedImp = cboImportance.SelectedIndex > 0 ? cboImportance.SelectedItem?.ToString() ?? "" : "";
            string selectedAttach = cboAttachment.SelectedIndex > 0 ? cboAttachment.SelectedItem?.ToString() ?? "" : "";
            string selectedSource = cboSourceFile.SelectedIndex > 0 ? cboSourceFile.SelectedItem?.ToString() ?? "" : "";

            DateTime fromDate = dtpFrom.Value.Date;
            DateTime toDate = dtpTo.Value.Date.AddDays(1).AddTicks(-1);
            bool checkDates = !chkAllDates.Checked;

            _filteredEmails = _allEmails.Where(e =>
            {
                if (checkDates && e.DateObj.HasValue)
                {
                    if (e.DateObj.Value < fromDate || e.DateObj.Value > toDate) return false;
                }

                if (!string.IsNullOrEmpty(selectedUnit) && e.DonVi != selectedUnit) return false;
                if (!string.IsNullOrEmpty(selectedTopic) && e.HangMuc != selectedTopic) return false;
                if (!string.IsNullOrEmpty(selectedImp) && !e.MucDo.Contains(selectedImp, StringComparison.OrdinalIgnoreCase)) return false;

                if (selectedAttach == "Có đính kèm" && e.DinhKem != "Co") return false;
                if (selectedAttach == "Không đính kèm" && e.DinhKem == "Co") return false;

                if (!string.IsNullOrEmpty(selectedSource) && e.NguonFile != selectedSource) return false;

                if (!string.IsNullOrEmpty(keyword))
                {
                    bool match = (e.ChuDe?.ToLower().Contains(keyword) == true) ||
                                 (e.NguoiGui?.ToLower().Contains(keyword) == true) ||
                                 (e.EmailNguoiGui?.ToLower().Contains(keyword) == true) ||
                                 (e.NguoiNhan_To?.ToLower().Contains(keyword) == true) ||
                                 (e.DonVi?.ToLower().Contains(keyword) == true) ||
                                 (e.HangMuc?.ToLower().Contains(keyword) == true);
                    if (!match) return false;
                }

                return true;
            }).ToList();

            dgvEmails.Rows.Clear();
            int stt = 1;
            foreach (var em in _filteredEmails)
            {
                int rowIdx = dgvEmails.Rows.Add(
                    stt++,
                    em.Ngay,
                    em.Gio,
                    em.NguoiGui,
                    em.DonVi,
                    em.ChuDe,
                    em.HangMuc,
                    em.MucDo,
                    em.DinhKem == "Co" ? "📎 Có" : "Không",
                    em.NguonFile
                );

                if (em.MucDo.Contains("Cao"))
                {
                    dgvEmails.Rows[rowIdx].Cells["MucDo"].Style.ForeColor = Color.Red;
                    dgvEmails.Rows[rowIdx].Cells["MucDo"].Style.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
                }
            }

            lblKpiTotal.Text = $"Tổng email: {_allEmails.Count:N0}  |  ";
            lblKpiFiltered.Text = $"Đang hiển thị: {_filteredEmails.Count:N0}  |  ";
            int senders = _filteredEmails.Select(e => e.NguoiGui).Distinct().Count();
            lblKpiSenders.Text = $"Người gửi: {senders:N0}  |  ";
            int attach = _filteredEmails.Count(e => e.DinhKem == "Co");
            lblKpiAttachments.Text = $"Có đính kèm: {attach:N0}";
        }

        private void ResetFilter()
        {
            chkAllDates.Checked = true;
            cboUnit.SelectedIndex = 0;
            cboTopic.SelectedIndex = 0;
            cboImportance.SelectedIndex = 0;
            cboAttachment.SelectedIndex = 0;
            cboSourceFile.SelectedIndex = 0;
            txtSearchKeyword.Clear();
            ApplyFilter();
        }

        private void DgvEmails_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _filteredEmails.Count) return;
            var em = _filteredEmails[e.RowIndex];

            string info = $"THÔNG TIN CHI TIẾT EMAIL:\n\n" +
                          $"• Chủ đề: {em.ChuDe}\n" +
                          $"• Ngày nhận: {em.Ngay} {em.Gio} ({em.Tuan})\n" +
                          $"• Người gửi: {em.NguoiGui} <{em.EmailNguoiGui}>\n" +
                          $"• Đơn vị: {em.DonVi}\n" +
                          $"• Người nhận (To): {em.NguoiNhan_To}\n" +
                          $"• Người nhận (CC): {em.NguoiNhan_CC}\n" +
                          $"• Hạng mục: {em.HangMuc}\n" +
                          $"• Mức độ: {em.MucDo}\n" +
                          $"• Đính kèm: {(em.DinhKem == "Co" ? "CÓ FILE ĐÍNH KÈM" : "Không có")}\n" +
                          $"• Thư mục gốc: {em.ThuMuc}\n" +
                          $"• Nguồn file: {em.NguonFile}";

            MessageBox.Show(info, "Chi tiết Email", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ExportExcel(bool filteredOnly)
        {
            var dataToExport = filteredOnly ? _filteredEmails : _allEmails;
            if (dataToExport == null || dataToExport.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu email nào để xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!Directory.Exists(OutputFolder)) Directory.CreateDirectory(OutputFolder);
            string tag = filteredOnly ? "Filtered" : "TatCa";
            string outputPath = Path.Combine(OutputFolder, $"ThongKe_Email_{tag}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            SetLoading(true, "Đang tạo file Excel báo cáo...");

            try
            {
                ExportToExcelFile(dataToExport, outputPath);

                string vbaScript = Path.Combine(OutputFolder, "inject_vba_macro.ps1");
                CreateVbaInjectionScript(vbaScript, outputPath);

                SetLoading(false, $"Xuất Excel thành công: {Path.GetFileName(outputPath)}");

                var res = MessageBox.Show($"Đã xuất thành công {dataToExport.Count:N0} email ra file Excel:\n\n{outputPath}\n\nBạn có muốn MỞ FILE NGAY không?",
                                          "Xuất Excel thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    Process.Start(new ProcessStartInfo(outputPath) { UseShellExecute = true });
                    RunVbaInjection(vbaScript);
                }
            }
            catch (Exception ex)
            {
                SetLoading(false, "Lỗi khi xuất Excel!");
                MessageBox.Show($"Lỗi xuất Excel: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetLoading(bool loading, string statusText)
        {
            lblStatus.Text = statusText;
            progressBar.Visible = loading;
            btnLoadCache.Enabled = !loading;
            btnSaveCache.Enabled = !loading;
            btnScanAll.Enabled = !loading;
            btnOpenSingle.Enabled = !loading;
            btnExportFiltered.Enabled = !loading;
            btnExportAll.Enabled = !loading;
        }

        #region Core File Readers & Excel Export

        private static List<EmailRecord> ReadOutlookFile(string filePath, string sourceLabel)
        {
            var results = new List<EmailRecord>();
            using (var xstFile = new XstFile(filePath))
            {
                ProcessFolder(xstFile.RootFolder, "", sourceLabel, results);
            }
            return results;
        }

        private static List<EmailRecord> ReadWithTempCopy(string filePath, string sourceLabel)
        {
            if (!Directory.Exists(TempCopyFolder)) Directory.CreateDirectory(TempCopyFolder);
            string tempPath = Path.Combine(TempCopyFolder, $"temp_{Guid.NewGuid():N}_{Path.GetFileName(filePath)}");

            using (var source = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var dest = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
            {
                source.CopyTo(dest);
            }

            try
            {
                return ReadOutlookFile(tempPath, sourceLabel);
            }
            finally
            {
                try { File.Delete(tempPath); } catch { }
            }
        }

        private static void ProcessFolder(XstFolder folder, string parentPath, string sourceLabel, List<EmailRecord> results)
        {
            string currentPath = string.IsNullOrEmpty(parentPath)
                ? (folder.DisplayName ?? "")
                : $"{parentPath}\\{folder.DisplayName}";

            try
            {
                foreach (var msg in folder.Messages)
                {
                    try
                    {
                        var record = ExtractEmailRecord(msg, currentPath, sourceLabel);
                        if (record != null) results.Add(record);
                    }
                    catch { }

                    try { msg.ClearContents(); } catch { }
                }
            }
            catch { }

            try
            {
                foreach (var subFolder in folder.Folders)
                {
                    ProcessFolder(subFolder, currentPath, sourceLabel, results);
                }
            }
            catch { }
        }

        private static EmailRecord? ExtractEmailRecord(XstMessage msg, string folderPath, string sourceLabel)
        {
            if (msg == null) return null;

            string senderName = "";
            string senderEmail = "";
            string subject = "";
            DateTime? receivedTime = null;
            DateTime? sentTime = null;
            bool hasAttachment = false;
            string importance = "Binh thuong";

            try
            {
                if (msg.Recipients?.Sender != null)
                {
                    senderName = msg.Recipients.Sender.DisplayName ?? "";
                    senderEmail = msg.Recipients.Sender.Address ?? "";
                }
                if (string.IsNullOrEmpty(senderName))
                    senderName = msg.From ?? "";
            }
            catch { }

            try { subject = msg.Subject ?? ""; } catch { }
            try { receivedTime = msg.ReceivedTime; } catch { }
            try { sentTime = msg.SubmittedTime; } catch { }
            try { hasAttachment = msg.HasAttachments; } catch { }

            try
            {
                var imp = msg.Importance;
                if (imp.HasValue)
                {
                    importance = imp.Value.ToString() switch
                    {
                        "High" => "Cao (High)",
                        "Low" => "Thap",
                        _ => "Binh thuong"
                    };
                }
            }
            catch { }

            string toList = "";
            string ccList = "";
            try
            {
                if (msg.Recipients != null)
                {
                    if (msg.Recipients.To != null)
                        toList = string.Join("; ", msg.Recipients.To.Select(r => r.DisplayName ?? r.Address ?? ""));
                    if (msg.Recipients.Cc != null)
                        ccList = string.Join("; ", msg.Recipients.Cc.Select(r => r.DisplayName ?? r.Address ?? ""));
                }
            }
            catch { }

            if (string.IsNullOrEmpty(toList))
                try { toList = msg.To ?? ""; } catch { }
            if (string.IsNullOrEmpty(ccList))
                try { ccList = msg.Cc ?? ""; } catch { }

            DateTime? dateObj = receivedTime ?? sentTime;
            string dateStr = "";
            string timeStr = "";
            string weekStr = "";

            if (dateObj.HasValue)
            {
                dateStr = dateObj.Value.ToString("yyyy-MM-dd");
                timeStr = dateObj.Value.ToString("HH:mm:ss");
                var culture = CultureInfo.InvariantCulture;
                int weekNum = culture.Calendar.GetWeekOfYear(dateObj.Value,
                    CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                weekStr = $"Tuan {weekNum}";
            }

            return new EmailRecord
            {
                Ngay = dateStr,
                Gio = timeStr,
                Tuan = weekStr,
                DateObj = dateObj,
                NguoiGui = senderName,
                EmailNguoiGui = senderEmail,
                DonVi = GetUnitCategory(senderName, senderEmail),
                NguoiNhan_To = toList,
                NguoiNhan_CC = ccList,
                ChuDe = subject,
                HangMuc = GetTopicCategory(subject),
                MucDo = importance,
                DinhKem = hasAttachment ? "Co" : "Khong",
                DungLuong_KB = 0,
                ThuMuc = folderPath,
                NguonFile = sourceLabel
            };
        }

        private static void ExportToExcelFile(List<EmailRecord> emails, string outputPath)
        {
            using (var wb = new XLWorkbook())
            {
                CreateDashboardSheet(wb, emails);
                CreateDetailSheet(wb, emails);
                CreateDailyStatSheet(wb, emails);
                CreateSenderStatSheet(wb, emails);
                CreateRecipientStatSheet(wb, emails);
                CreateMatrixSheet(wb, emails);

                wb.SaveAs(outputPath);
            }
        }

        private static void CreateDashboardSheet(XLWorkbook wb, List<EmailRecord> emails)
        {
            var ws = wb.Worksheets.Add("Trang Chu - Dashboard");

            ws.Range("A1:H1").Merge();
            ws.Cell("A1").Value = "BẢNG THỐNG KÊ EMAIL DỰ ÁN - ĐỌC TRỰC TIẾP TỪ FILE OUTLOOK";
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Font.FontSize = 16;
            ws.Cell("A1").Style.Font.FontColor = XLColor.White;
            ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.FromArgb(24, 43, 73);
            ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Row(1).Height = 35;

            ws.Range("A2:H2").Merge();
            ws.Cell("A2").Value = $"Cập nhật lúc: {DateTime.Now:dd/MM/yyyy HH:mm:ss} | Tác giả: Huỳnh Ngọc Nam (NXSOFT) - Namhuynh1205@gmail.com";
            ws.Cell("A2").Style.Font.Italic = true;
            ws.Cell("A2").Style.Font.FontSize = 10;
            ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Range("A3:H3").Merge();
            ws.Cell("A3").Value = "Phần mềm phát triển bởi Huỳnh Ngọc Nam | Bấm các nút chức năng để mở phần mềm lọc mail hoặc cập nhật live";
            ws.Cell("A3").Style.Font.Italic = true;
            ws.Cell("A3").Style.Font.FontSize = 10;
            ws.Cell("A3").Style.Font.FontColor = XLColor.FromArgb(24, 43, 73);
            ws.Cell("A3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Row(4).Height = 45;

            string[] kpiHeaders = { "TỔNG EMAIL", "", "SỐ NGƯỜI GỬI", "", "MỨC ĐỘ CAO", "CÓ ĐÍNH KÈM", "SỐ NGÀY", "SỐ NGUỒN FILE" };
            var kpiColors = new[]
            {
                XLColor.FromArgb(0, 112, 192), XLColor.FromArgb(0, 112, 192),
                XLColor.FromArgb(38, 128, 0), XLColor.FromArgb(38, 128, 0),
                XLColor.FromArgb(192, 0, 0), XLColor.FromArgb(112, 48, 160),
                XLColor.FromArgb(41, 86, 129), XLColor.FromArgb(41, 86, 129)
            };

            for (int c = 0; c < kpiHeaders.Length; c++)
            {
                if (string.IsNullOrEmpty(kpiHeaders[c])) continue;
                ws.Cell(5, c + 1).Value = kpiHeaders[c];
                ws.Cell(5, c + 1).Style.Font.Bold = true;
                ws.Cell(5, c + 1).Style.Font.FontSize = 10;
                ws.Cell(5, c + 1).Style.Font.FontColor = XLColor.White;
                ws.Cell(5, c + 1).Style.Fill.BackgroundColor = kpiColors[c];
                ws.Cell(5, c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int uniqueSenders = emails.Select(e => e.NguoiGui).Where(s => !string.IsNullOrEmpty(s)).Distinct().Count();
            int highCount = emails.Count(e => e.MucDo.Contains("Cao"));
            int attachCount = emails.Count(e => e.DinhKem == "Co");
            int dayCount = emails.Select(e => e.Ngay).Where(d => !string.IsNullOrEmpty(d)).Distinct().Count();
            int sourceCount = emails.Select(e => e.NguonFile).Distinct().Count();

            ws.Cell("A6").Value = emails.Count;
            ws.Cell("C6").Value = uniqueSenders;
            ws.Cell("E6").Value = highCount;
            ws.Cell("F6").Value = attachCount;
            ws.Cell("G6").Value = dayCount;
            ws.Cell("H6").Value = sourceCount;

            for (int c = 1; c <= 8; c++)
            {
                ws.Cell(6, c).Style.Font.Bold = true;
                ws.Cell(6, c).Style.Font.FontSize = 18;
                ws.Cell(6, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(6, c).Style.Fill.BackgroundColor = XLColor.FromArgb(240, 244, 248);
            }

            // Unit Stats
            ws.Range("A8:D8").Merge();
            ws.Cell("A8").Value = "THỐNG KÊ THEO ĐƠN VỊ";
            ws.Cell("A8").Style.Font.Bold = true;
            ws.Cell("A8").Style.Font.FontSize = 12;
            ws.Cell("A8").Style.Font.FontColor = XLColor.White;
            ws.Cell("A8").Style.Fill.BackgroundColor = XLColor.FromArgb(0, 112, 192);
            ws.Range("A8:D8").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Cell("A9").Value = "STT"; ws.Cell("B9").Value = "Đơn vị";
            ws.Cell("C9").Value = "Số email"; ws.Cell("D9").Value = "Tỷ lệ (%)";
            for (int c = 1; c <= 4; c++)
            {
                ws.Cell(9, c).Style.Font.Bold = true;
                ws.Cell(9, c).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 112, 192);
                ws.Cell(9, c).Style.Font.FontColor = XLColor.White;
                ws.Cell(9, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int row = 10; int stt = 1;
            foreach (var ug in emails.GroupBy(e => e.DonVi).OrderByDescending(g => g.Count()))
            {
                double pct = Math.Round((double)ug.Count() / emails.Count * 100, 1);
                ws.Cell(row, 1).Value = stt;
                ws.Cell(row, 2).Value = ug.Key;
                ws.Cell(row, 3).Value = ug.Count();
                ws.Cell(row, 4).Value = $"{pct} %";
                if (row % 2 == 0) ws.Range(row, 1, row, 4).Style.Fill.BackgroundColor = XLColor.FromArgb(235, 243, 250);
                row++; stt++;
            }
            ws.Range(9, 1, row - 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(9, 1, row - 1, 4).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Topic Stats
            ws.Range("E8:H8").Merge();
            ws.Cell("E8").Value = "THỐNG KÊ THEO HẠNG MỤC CÔNG VIỆC";
            ws.Cell("E8").Style.Font.Bold = true;
            ws.Cell("E8").Style.Font.FontSize = 12;
            ws.Cell("E8").Style.Font.FontColor = XLColor.White;
            ws.Cell("E8").Style.Fill.BackgroundColor = XLColor.FromArgb(38, 128, 0);
            ws.Range("E8:H8").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Cell("E9").Value = "STT"; ws.Cell("F9").Value = "Hạng mục";
            ws.Cell("G9").Value = "Số email"; ws.Cell("H9").Value = "Tỷ lệ (%)";
            for (int c = 5; c <= 8; c++)
            {
                ws.Cell(9, c).Style.Font.Bold = true;
                ws.Cell(9, c).Style.Fill.BackgroundColor = XLColor.FromArgb(38, 128, 0);
                ws.Cell(9, c).Style.Font.FontColor = XLColor.White;
                ws.Cell(9, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int tRow = 10; stt = 1;
            foreach (var tg in emails.GroupBy(e => e.HangMuc).OrderByDescending(g => g.Count()))
            {
                double pct = Math.Round((double)tg.Count() / emails.Count * 100, 1);
                ws.Cell(tRow, 5).Value = stt;
                ws.Cell(tRow, 6).Value = tg.Key;
                ws.Cell(tRow, 7).Value = tg.Count();
                ws.Cell(tRow, 8).Value = $"{pct} %";
                if (tRow % 2 == 0) ws.Range(tRow, 5, tRow, 8).Style.Fill.BackgroundColor = XLColor.FromArgb(235, 247, 235);
                tRow++; stt++;
            }
            ws.Range(9, 5, tRow - 1, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(9, 5, tRow - 1, 8).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Column(1).Width = 8; ws.Column(2).Width = 42; ws.Column(3).Width = 12;
            ws.Column(4).Width = 12; ws.Column(5).Width = 8; ws.Column(6).Width = 42;
            ws.Column(7).Width = 12; ws.Column(8).Width = 15;

            ws.SheetView.FreezeRows(6);
        }

        private static void CreateDetailSheet(XLWorkbook wb, List<EmailRecord> emails)
        {
            var ws = wb.Worksheets.Add("Chi tiet Email");

            string[] headers = {
                "STT", "Ngày", "Giờ", "Tuần", "Người gửi", "Email người gửi",
                "Đơn vị", "Người nhận (To)", "Người nhận (CC)", "Chủ đề",
                "Hạng mục", "Mức độ", "Đính kèm", "Thư mục", "Nguồn file"
            };

            for (int c = 0; c < headers.Length; c++)
            {
                ws.Cell(1, c + 1).Value = headers[c];
                ws.Cell(1, c + 1).Style.Font.Bold = true;
                ws.Cell(1, c + 1).Style.Font.FontSize = 11;
                ws.Cell(1, c + 1).Style.Font.FontColor = XLColor.White;
                ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(24, 43, 73);
                ws.Cell(1, c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            for (int i = 0; i < emails.Count; i++)
            {
                int row = i + 2;
                var em = emails[i];

                ws.Cell(row, 1).Value = i + 1;
                ws.Cell(row, 2).Value = em.Ngay;
                ws.Cell(row, 3).Value = em.Gio;
                ws.Cell(row, 4).Value = em.Tuan;
                ws.Cell(row, 5).Value = em.NguoiGui;
                ws.Cell(row, 6).Value = em.EmailNguoiGui;
                ws.Cell(row, 7).Value = em.DonVi;
                ws.Cell(row, 8).Value = em.NguoiNhan_To;
                ws.Cell(row, 9).Value = em.NguoiNhan_CC;
                ws.Cell(row, 10).Value = em.ChuDe;
                ws.Cell(row, 11).Value = em.HangMuc;
                ws.Cell(row, 12).Value = em.MucDo;
                ws.Cell(row, 13).Value = em.DinhKem;
                ws.Cell(row, 14).Value = em.ThuMuc;
                ws.Cell(row, 15).Value = em.NguonFile;

                if (row % 2 == 0)
                    ws.Range(row, 1, row, 15).Style.Fill.BackgroundColor = XLColor.FromArgb(235, 243, 250);

                if (em.MucDo.Contains("Cao"))
                {
                    ws.Cell(row, 12).Style.Font.FontColor = XLColor.Red;
                    ws.Cell(row, 12).Style.Font.Bold = true;
                }
            }

            int lastRow = emails.Count + 1;
            ws.Range(1, 1, lastRow, 15).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(1, 1, lastRow, 15).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Column(1).Width = 8;
            ws.Column(2).Width = 12;
            ws.Column(3).Width = 10;
            ws.Column(4).Width = 10;
            ws.Column(5).Width = 25;
            ws.Column(6).Width = 30;
            ws.Column(7).Width = 30;
            ws.Column(8).Width = 30;
            ws.Column(9).Width = 30;
            ws.Column(10).Width = 60;
            ws.Column(11).Width = 30;
            ws.Column(12).Width = 15;
            ws.Column(13).Width = 10;
            ws.Column(14).Width = 40;
            ws.Column(15).Width = 30;

            ws.SheetView.FreezeRows(1);
            ws.SheetView.FreezeColumns(1);
        }

        private static void CreateDailyStatSheet(XLWorkbook wb, List<EmailRecord> emails)
        {
            var ws = wb.Worksheets.Add("Thong ke theo Ngay");

            string[] headers = { "STT", "Ngày", "Thứ", "Tuần", "Số email", "Số người gửi", "Số người nhận", "Mức độ Cao", "Có đính kèm" };

            for (int c = 0; c < headers.Length; c++)
            {
                ws.Cell(1, c + 1).Value = headers[c];
                ws.Cell(1, c + 1).Style.Font.Bold = true;
                ws.Cell(1, c + 1).Style.Font.FontColor = XLColor.White;
                ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 112, 192);
                ws.Cell(1, c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            var dateGroups = emails
                .Where(e => !string.IsNullOrEmpty(e.Ngay))
                .GroupBy(e => e.Ngay)
                .OrderByDescending(g => g.Key);

            int row = 2; int stt = 1;
            foreach (var dg in dateGroups)
            {
                DateTime dd;
                string dow = "";
                if (DateTime.TryParse(dg.Key, out dd))
                {
                    dow = dd.DayOfWeek switch
                    {
                        DayOfWeek.Monday => "Thu 2",
                        DayOfWeek.Tuesday => "Thu 3",
                        DayOfWeek.Wednesday => "Thu 4",
                        DayOfWeek.Thursday => "Thu 5",
                        DayOfWeek.Friday => "Thu 6",
                        DayOfWeek.Saturday => "Thu 7",
                        DayOfWeek.Sunday => "Chu Nhat",
                        _ => ""
                    };
                }

                string weekStr = dg.First().Tuan;
                int uniqueSenders = dg.Select(e => e.NguoiGui).Where(s => !string.IsNullOrEmpty(s)).Distinct().Count();

                var allRecipients = dg
                    .SelectMany(e => (e.NguoiNhan_To ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries))
                    .Select(r => r.Trim())
                    .Where(r => !string.IsNullOrEmpty(r))
                    .Distinct()
                    .Count();

                int highCount = dg.Count(e => e.MucDo.Contains("Cao"));
                int attachCount = dg.Count(e => e.DinhKem == "Co");

                ws.Cell(row, 1).Value = stt;
                ws.Cell(row, 2).Value = dg.Key;
                ws.Cell(row, 3).Value = dow;
                ws.Cell(row, 4).Value = weekStr;
                ws.Cell(row, 5).Value = dg.Count();
                ws.Cell(row, 6).Value = uniqueSenders;
                ws.Cell(row, 7).Value = allRecipients;
                ws.Cell(row, 8).Value = highCount;
                ws.Cell(row, 9).Value = attachCount;

                if (dow == "Thu 7" || dow == "Chu Nhat")
                    ws.Range(row, 1, row, 9).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 220, 220);
                else if (row % 2 == 0)
                    ws.Range(row, 1, row, 9).Style.Fill.BackgroundColor = XLColor.FromArgb(235, 247, 235);

                row++; stt++;
            }

            ws.Cell(row, 2).Value = "TỔNG CỘNG";
            ws.Cell(row, 5).Value = emails.Count;
            ws.Range(row, 1, row, 9).Style.Font.Bold = true;
            ws.Range(row, 1, row, 9).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 242, 204);

            ws.Range(1, 1, row, 9).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(1, 1, row, 9).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Column(1).Width = 6; ws.Column(2).Width = 14; ws.Column(3).Width = 10;
            ws.Column(4).Width = 10; ws.Column(5).Width = 12; ws.Column(6).Width = 14;
            ws.Column(7).Width = 16; ws.Column(8).Width = 14; ws.Column(9).Width = 14;

            ws.SheetView.FreezeRows(1);
        }

        private static void CreateSenderStatSheet(XLWorkbook wb, List<EmailRecord> emails)
        {
            var ws = wb.Worksheets.Add("Thong ke Nguoi gui");

            string[] headers = { "STT", "Người gửi", "Email", "Đơn vị", "Số email", "Ngày đầu", "Ngày cuối", "Số ngày gửi" };

            for (int c = 0; c < headers.Length; c++)
            {
                ws.Cell(1, c + 1).Value = headers[c];
                ws.Cell(1, c + 1).Style.Font.Bold = true;
                ws.Cell(1, c + 1).Style.Font.FontColor = XLColor.White;
                ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(192, 0, 0);
                ws.Cell(1, c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            var senderGroups = emails
                .Where(e => !string.IsNullOrEmpty(e.NguoiGui))
                .GroupBy(e => e.NguoiGui)
                .OrderByDescending(g => g.Count());

            int row = 2; int stt = 1;
            foreach (var sg in senderGroups)
            {
                var withDate = sg.Where(e => e.DateObj.HasValue).ToList();
                string firstDate = withDate.Any() ? withDate.Min(e => e.DateObj!.Value).ToString("yyyy-MM-dd") : "";
                string lastDate = withDate.Any() ? withDate.Max(e => e.DateObj!.Value).ToString("yyyy-MM-dd") : "";
                int dayCount = sg.Select(e => e.Ngay).Where(d => !string.IsNullOrEmpty(d)).Distinct().Count();
                var first = sg.First();

                ws.Cell(row, 1).Value = stt;
                ws.Cell(row, 2).Value = sg.Key;
                ws.Cell(row, 3).Value = first.EmailNguoiGui;
                ws.Cell(row, 4).Value = first.DonVi;
                ws.Cell(row, 5).Value = sg.Count();
                ws.Cell(row, 6).Value = firstDate;
                ws.Cell(row, 7).Value = lastDate;
                ws.Cell(row, 8).Value = dayCount;

                if (row % 2 == 0)
                    ws.Range(row, 1, row, 8).Style.Fill.BackgroundColor = XLColor.FromArgb(253, 237, 237);

                row++; stt++;
            }

            ws.Range(1, 1, row - 1, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(1, 1, row - 1, 8).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Column(1).Width = 6; ws.Column(2).Width = 32; ws.Column(3).Width = 35;
            ws.Column(4).Width = 35; ws.Column(5).Width = 12; ws.Column(6).Width = 14;
            ws.Column(7).Width = 14; ws.Column(8).Width = 14;

            ws.SheetView.FreezeRows(1);
        }

        private static void CreateRecipientStatSheet(XLWorkbook wb, List<EmailRecord> emails)
        {
            var ws = wb.Worksheets.Add("Thong ke Nguoi nhan");

            string[] headers = { "STT", "Người nhận", "Tổng email", "Nhận To", "Nhận CC", "Số ngày" };

            for (int c = 0; c < headers.Length; c++)
            {
                ws.Cell(1, c + 1).Value = headers[c];
                ws.Cell(1, c + 1).Style.Font.Bold = true;
                ws.Cell(1, c + 1).Style.Font.FontColor = XLColor.White;
                ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(112, 48, 160);
                ws.Cell(1, c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            var rMap = new Dictionary<string, (int Total, int To, int Cc, HashSet<string> Days)>();
            foreach (var em in emails)
            {
                var toNames = (em.NguoiNhan_To ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries)
                    .Select(r => r.Trim()).Where(r => !string.IsNullOrEmpty(r));
                var ccNames = (em.NguoiNhan_CC ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries)
                    .Select(r => r.Trim()).Where(r => !string.IsNullOrEmpty(r));

                foreach (var name in toNames)
                {
                    if (!rMap.ContainsKey(name)) rMap[name] = (0, 0, 0, new HashSet<string>());
                    var v = rMap[name];
                    rMap[name] = (v.Total + 1, v.To + 1, v.Cc, v.Days);
                    if (!string.IsNullOrEmpty(em.Ngay)) rMap[name].Days.Add(em.Ngay);
                }
                foreach (var name in ccNames)
                {
                    if (!rMap.ContainsKey(name)) rMap[name] = (0, 0, 0, new HashSet<string>());
                    var v = rMap[name];
                    rMap[name] = (v.Total + 1, v.To, v.Cc + 1, v.Days);
                    if (!string.IsNullOrEmpty(em.Ngay)) rMap[name].Days.Add(em.Ngay);
                }
            }

            int row = 2; int stt = 1;
            foreach (var kv in rMap.OrderByDescending(x => x.Value.Total).Take(500))
            {
                ws.Cell(row, 1).Value = stt;
                ws.Cell(row, 2).Value = kv.Key;
                ws.Cell(row, 3).Value = kv.Value.Total;
                ws.Cell(row, 4).Value = kv.Value.To;
                ws.Cell(row, 5).Value = kv.Value.Cc;
                ws.Cell(row, 6).Value = kv.Value.Days.Count;

                if (row % 2 == 0)
                    ws.Range(row, 1, row, 6).Style.Fill.BackgroundColor = XLColor.FromArgb(245, 238, 248);

                row++; stt++;
            }

            ws.Range(1, 1, row - 1, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(1, 1, row - 1, 6).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Column(1).Width = 6; ws.Column(2).Width = 35; ws.Column(3).Width = 12;
            ws.Column(4).Width = 12; ws.Column(5).Width = 12; ws.Column(6).Width = 12;

            ws.SheetView.FreezeRows(1);
        }

        private static void CreateMatrixSheet(XLWorkbook wb, List<EmailRecord> emails)
        {
            var ws = wb.Worksheets.Add("Ma tran Gui x Ngay");

            var senderGroups = emails
                .Where(e => !string.IsNullOrEmpty(e.NguoiGui))
                .GroupBy(e => e.NguoiGui)
                .OrderByDescending(g => g.Count())
                .Take(30)
                .ToList();

            var uniqueDates = emails
                .Where(e => !string.IsNullOrEmpty(e.Ngay))
                .Select(e => e.Ngay)
                .Distinct()
                .OrderBy(d => d)
                .ToList();

            ws.Cell(1, 1).Value = "STT";
            ws.Cell(1, 2).Value = "Người gửi";
            ws.Cell(1, 3).Value = "Đơn vị";

            for (int d = 0; d < uniqueDates.Count; d++)
                ws.Cell(1, d + 4).Value = uniqueDates[d];

            int totalCol = uniqueDates.Count + 4;
            ws.Cell(1, totalCol).Value = "TỔNG";

            ws.Range(1, 1, 1, totalCol).Style.Font.Bold = true;
            ws.Range(1, 1, 1, totalCol).Style.Font.FontSize = 10;
            ws.Range(1, 1, 1, totalCol).Style.Font.FontColor = XLColor.White;
            ws.Range(1, 1, 1, totalCol).Style.Fill.BackgroundColor = XLColor.FromArgb(41, 86, 129);
            ws.Range(1, 1, 1, totalCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            int row = 2; int stt = 1;
            foreach (var s in senderGroups)
            {
                ws.Cell(row, 1).Value = stt;
                ws.Cell(row, 2).Value = s.Key;
                ws.Cell(row, 3).Value = s.First().DonVi;

                int sumTotal = 0;
                for (int d = 0; d < uniqueDates.Count; d++)
                {
                    int cnt = s.Count(e => e.Ngay == uniqueDates[d]);
                    if (cnt > 0)
                    {
                        ws.Cell(row, d + 4).Value = cnt;
                        sumTotal += cnt;

                        if (cnt >= 5)
                            ws.Cell(row, d + 4).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 120, 120);
                        else if (cnt >= 3)
                            ws.Cell(row, d + 4).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 200, 100);
                        else
                            ws.Cell(row, d + 4).Style.Fill.BackgroundColor = XLColor.FromArgb(160, 225, 120);
                    }
                }

                ws.Cell(row, totalCol).Value = sumTotal;
                row++; stt++;
            }

            ws.Cell(row, 2).Value = "TỔNG / NGÀY";
            ws.Cell(row, 2).Style.Font.Bold = true;
            for (int d = 0; d < uniqueDates.Count; d++)
            {
                ws.Cell(row, d + 4).Value = emails.Count(e => e.Ngay == uniqueDates[d]);
            }
            ws.Cell(row, totalCol).Value = emails.Count;
            ws.Range(row, 1, row, totalCol).Style.Font.Bold = true;
            ws.Range(row, 1, row, totalCol).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 242, 204);

            ws.Range(1, 1, row, totalCol).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(1, 1, row, totalCol).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Column(1).Width = 5; ws.Column(2).Width = 32; ws.Column(3).Width = 30;
            for (int d = 0; d < uniqueDates.Count; d++) ws.Column(d + 4).Width = 11;
            ws.Column(totalCol).Width = 8;

            ws.SheetView.FreezeRows(1);
            ws.SheetView.FreezeColumns(3);
        }

        private static void CreateVbaInjectionScript(string scriptPath, string excelPath)
        {
            string exePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutLookTool.exe");
            string script = $@"# inject_vba_macro.ps1
param(
    [string]$TargetExcel = ""{excelPath}""
)

Start-Sleep -Seconds 3

try {{
    $excel = $null
    for ($attempt = 1; $attempt -le 10; $attempt++) {{
        try {{
            $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject(""Excel.Application"")
            break
        }} catch {{
            Start-Sleep -Seconds 2
        }}
    }}

    if (-not $excel) {{ exit 0 }}

    $wb = $null
    foreach ($wbItem in $excel.Workbooks) {{
        try {{
            if ($wbItem.FullName -eq $TargetExcel -or $wbItem.Name -like ""*ThongKe_Email*"") {{
                $wb = $wbItem
                break
            }}
        }} catch {{}}
    }}

    if (-not $wb) {{ $wb = $excel.Workbooks.Open($TargetExcel) }}

    $wsDash = $null
    foreach ($ws in $wb.Worksheets) {{
        if ($ws.Name -like ""*Dashboard*"" -or $ws.Name -like ""*Trang*"") {{
            $wsDash = $ws
            break
        }}
    }}
    if (-not $wsDash) {{ $wsDash = $wb.Worksheets.Item(1) }}

    foreach ($sName in @(""btnCapNhatNgay"", ""btnTuDong30Phut"", ""btnDungTuDong"", ""btnDocTrucTiep"")) {{
        try {{ $wsDash.Shapes.Item($sName).Delete() }} catch {{}}
    }}

    $btn0 = $wsDash.Shapes.AddShape(5, 5, 62, 210, 35)
    $btn0.Name = ""btnDocTrucTiep""
    $btn0.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(24, 43, 73))
    $btn0.Line.Visible = $false
    $btn0.TextFrame2.TextRange.Text = ""MỞ PHẦN MỀM LỌC EMAIL""
    $btn0.TextFrame2.TextRange.Font.Size = 10
    $btn0.TextFrame2.TextRange.Font.Bold = $true
    $btn0.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)
    $btn0.TextFrame2.HorizontalAnchor = 2
    $btn0.TextFrame2.VerticalAnchor = 3
    $btn0.Shadow.Visible = $true

    $btn1 = $wsDash.Shapes.AddShape(5, 225, 62, 180, 35)
    $btn1.Name = ""btnCapNhatNgay""
    $btn1.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(0, 112, 192))
    $btn1.Line.Visible = $false
    $btn1.TextFrame2.TextRange.Text = ""QUÉT OUTLOOK (LIVE)""
    $btn1.TextFrame2.TextRange.Font.Size = 10
    $btn1.TextFrame2.TextRange.Font.Bold = $true
    $btn1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)
    $btn1.TextFrame2.HorizontalAnchor = 2
    $btn1.TextFrame2.VerticalAnchor = 3
    $btn1.Shadow.Visible = $true

    try {{
        foreach ($comp in $wb.VBProject.VBComponents) {{
            if ($comp.Name -eq ""ModCapNhatEmail"") {{
                $wb.VBProject.VBComponents.Remove($comp)
                break
            }}
        }}
    }} catch {{}}

    $vbaCode = @""
Option Explicit

Public Sub MoPhanMemLocMail()
    Dim cmd As String
    cmd = """"{exePath}""""
    Dim wsh As Object
    Set wsh = CreateObject(""""WScript.Shell"""")
    wsh.Run cmd, 1, False
    Set wsh = Nothing
End Sub

Public Sub CapNhatLiveOutlook()
    Dim scriptPath As String
    scriptPath = """"C:\OutLook\update_data_live.ps1""""
    If Dir(scriptPath) = """""" Then
        MsgBox """"Khong tim thay script: """" & scriptPath, vbExclamation, """"Loi""""
        Exit Sub
    End If
    Dim cmd As String
    cmd = """"powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """""""" & scriptPath & """"""""""""
    Dim wsh As Object
    Set wsh = CreateObject(""""WScript.Shell"""")
    wsh.Run cmd, 0, True
    Set wsh = Nothing
    MsgBox """"Da dong bo xong du lieu Email!"", vbInformation, """"Thanh cong""""
End Sub
""@

    $vbModule = $wb.VBProject.VBComponents.Add(1)
    $vbModule.Name = ""ModCapNhatEmail""
    $vbModule.CodeModule.AddFromString($vbaCode)

    $wsDash.Shapes.Item(""btnDocTrucTiep"").OnAction = ""ModCapNhatEmail.MoPhanMemLocMail""
    $wsDash.Shapes.Item(""btnCapNhatNgay"").OnAction = ""ModCapNhatEmail.CapNhatLiveOutlook""

    $excel.DisplayAlerts = $false
    $currentPath = $wb.FullName
    if ($currentPath -like ""*.xlsx"") {{
        $newPath = $currentPath -replace '\.xlsx$', '.xlsm'
        $wb.SaveAs($newPath, 52)
    }} else {{
        $wb.Save()
    }}
    $excel.DisplayAlerts = $true
}} catch {{}}
";
            File.WriteAllText(scriptPath, script, System.Text.Encoding.UTF8);
        }

        private static void RunVbaInjection(string scriptPath)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-ExecutionPolicy Bypass -File \"{scriptPath}\"",
                    UseShellExecute = true,
                    CreateNoWindow = true,
                    WorkingDirectory = OutputFolder
                };
                Process.Start(psi);
            }
            catch { }
        }

        private static string GetUnitCategory(string senderName, string senderEmail)
        {
            string text = $"{senderName} {senderEmail}";

            if (Regex.IsMatch(text, @"SGC_KKH-KT|SGC_KKH|SGC_PKT|PKT", RegexOptions.IgnoreCase))
                return "SGC - Khoi Ke hoach Ky thuat";
            if (Regex.IsMatch(text, @"SGC_KĐSTĐC|SGC_KDSTDC|SGC_KĐSCT|SGC_DSCT", RegexOptions.IgnoreCase))
                return "SGC - Khoi Duong sat Toc do cao";
            if (Regex.IsMatch(text, @"SGC_KQLDA|SGC_QLDA", RegexOptions.IgnoreCase))
                return "SGC - Khoi Quan ly Du an";
            if (Regex.IsMatch(text, @"SGC_KQLXD|SGC_CTCT|BQLXD|Ban QLXD", RegexOptions.IgnoreCase))
                return "SGC - Ban QLXD & Cong trinh";
            if (Regex.IsMatch(text, @"SGC_KTCH|SGC_KTCXD|SGC_KTCCKN", RegexOptions.IgnoreCase))
                return "SGC - Khoi Than toc / Khoi Ky thuat";
            if (Regex.IsMatch(text, @"SGC_KQLKT|SGC_KHT|SGC_VP|SGC_Phòng", RegexOptions.IgnoreCase))
                return "SGC - Khoi Van phong & Hanh chinh";
            if (Regex.IsMatch(text, @"VH-VTK|VH-KCUDT|VH-KXD|VH-PTDA|Vinhomes|vinhomes\.vn", RegexOptions.IgnoreCase))
                return "Vinhomes - Vien Thiet ke & Phat trien DA";
            if (Regex.IsMatch(text, @"VinCons|vincons\.net", RegexOptions.IgnoreCase))
                return "Nha thau VinCons";
            if (Regex.IsMatch(text, @"SYSTRA|systra\.com", RegexOptions.IgnoreCase))
                return "Tu van SYSTRA";
            if (Regex.IsMatch(text, @"CONINCO|ckjvn\.vn|coninco", RegexOptions.IgnoreCase))
                return "Tu van Giam sat CONINCO";
            if (Regex.IsMatch(text, @"Trung Chính|trungchinhtcc", RegexOptions.IgnoreCase))
                return "Nha thau Trung Chinh";
            if (Regex.IsMatch(text, @"Đại Dũng|daidung\.vn", RegexOptions.IgnoreCase))
                return "Nha thau Dai Dung";
            if (Regex.IsMatch(text, @"Tecco1|tecco1\.vn", RegexOptions.IgnoreCase))
                return "Nha thau Tecco 1";
            if (Regex.IsMatch(text, @"Thiên Quốc An|TQA", RegexOptions.IgnoreCase))
                return "Nha thau Thien Quoc An";
            if (Regex.IsMatch(text, @"620idic", RegexOptions.IgnoreCase))
                return "Nha thau 620 IDIC";
            if (Regex.IsMatch(text, @"SGC", RegexOptions.IgnoreCase))
                return "SGC - Don vi khac";

            return "Doi tac / Don vi ngoai";
        }

        private static string GetTopicCategory(string subject)
        {
            if (string.IsNullOrEmpty(subject)) return "Chua xac dinh";

            if (Regex.IsMatch(subject, @"Bãi đúc|đầm|dầm|đốm|gói dầm", RegexOptions.IgnoreCase))
                return "Bai duc & Dam cau";
            if (Regex.IsMatch(subject, @"cọc|ép cọc|khoan nhồi|Bê tông|be tong|TRIAL MIX|C25|C30", RegexOptions.IgnoreCase))
                return "Thi cong Coc & Be tong";
            if (Regex.IsMatch(subject, @"Thiết kế|TK|Shop|SHOP|VMTK|MEP|Track|Design|bản vẽ|BIM", RegexOptions.IgnoreCase))
                return "Thiet ke & Shopdrawing";
            if (Regex.IsMatch(subject, @"cao độ|đường đỏ|ranh|rừng|mặt bằng|mốc|trắc địa", RegexOptions.IgnoreCase))
                return "Trac dia, Cao do & Mat bang";
            if (Regex.IsMatch(subject, @"Tiến độ|báo cáo|vướng mắc|ảnh hưởng|Đề xuất|de xuat", RegexOptions.IgnoreCase))
                return "Tien do & Vuong mac Cong truong";
            if (Regex.IsMatch(subject, @"Văn bản|Công văn|NNMT|phạt|pháp lý|Ký|phê duyệt|Quyết định", RegexOptions.IgnoreCase))
                return "Phap ly & Cong van Hanh chinh";
            if (Regex.IsMatch(subject, @"máy phát|điện|hàng rào|cầu rửa xe|thiết bị|vật tư|cáp", RegexOptions.IgnoreCase))
                return "Ha tang Bai duc & Thiet bi";

            return "Cong viec / Trao doi khac";
        }

        #endregion
    }

    public class EmailRecord
    {
        public int STT { get; set; }
        public string Ngay { get; set; } = "";
        public string Gio { get; set; } = "";
        public string Tuan { get; set; } = "";
        public DateTime? DateObj { get; set; }
        public string NguoiGui { get; set; } = "";
        public string EmailNguoiGui { get; set; } = "";
        public string DonVi { get; set; } = "";
        public string NguoiNhan_To { get; set; } = "";
        public string NguoiNhan_CC { get; set; } = "";
        public string ChuDe { get; set; } = "";
        public string HangMuc { get; set; } = "";
        public string MucDo { get; set; } = "Binh thuong";
        public string DinhKem { get; set; } = "Khong";
        public double DungLuong_KB { get; set; }
        public string ThuMuc { get; set; } = "";
        public string NguonFile { get; set; } = "";
    }
}
