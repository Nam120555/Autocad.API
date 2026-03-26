// Phát sinh cọc và đổi tên cọc theo quy cách Civil
// Quy cách: DT (đầu), CT (cuối), H1/H2 (100m), KM1/KM2 (1000m), 1~N (còn lại)
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.ApplicationServices;
using MyFirstProject.Extensions;

// Alias giải quyết xung đột namespace
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using ContentAlignment = System.Drawing.ContentAlignment;
using FontStyle = System.Drawing.FontStyle;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.PhatSinhCocDoiTenCoc))]

namespace MyFirstProject
{
    #region Form

    /// <summary>
    /// Form phát sinh cọc & đổi tên cọc — Dark Theme
    /// </summary>
    public class CocTuyenDuongForm : Form
    {
        // ─── Theme Colors ───
        private static readonly Color BgDark = Color.FromArgb(30, 30, 34);
        private static readonly Color BgPanel = Color.FromArgb(40, 40, 45);
        private static readonly Color BgInput = Color.FromArgb(50, 50, 58);
        private static readonly Color AccentBlue = Color.FromArgb(0, 122, 204);
        private static readonly Color HeaderBlue = Color.FromArgb(0, 93, 160);
        private static readonly Color TextWhite = Color.White;
        private static readonly Color TextDim = Color.FromArgb(160, 160, 170);
        private static readonly Color BorderColor = Color.FromArgb(65, 65, 75);
        private static readonly Color GreenOk = Color.FromArgb(46, 160, 67);
        private static readonly Color OrangeWarn = Color.FromArgb(210, 150, 40);

        // ─── Controls ───
        private RadioButton rdoPhatSinh = null!;
        private RadioButton rdoDoiTen = null!;
        private ComboBox cboAlignment = null!;
        private WinLabel lblAlignmentInfo = null!;
        private WinLabel lblGroupInfo = null!;

        // Phát sinh controls
        private Panel pnlPhatSinh = null!;
        private NumericUpDown numKhoangCach = null!; // Straight
        private NumericUpDown numKhoangCachCong = null!; // Curve
        private NumericUpDown numBeRongTrai = null!; // Left width
        private NumericUpDown numBeRongPhai = null!; // Right width

        // Naming options
        private CheckBox chkDT_CT = null!;
        private CheckBox chkH = null!;
        private CheckBox chkKM = null!;
        private CheckBox chkCurve = null!;
        private CheckBox chkPrefix = null!;
        private TextBox txtPrefix = null!;

        // Preview
        private DataGridView dgvPreview = null!;
        private WinLabel lblTotal = null!;
        private Button btnPreview = null!;

        // ─── Data ───
        private List<AlignmentInfo> alignments = new();
        public bool IsConfirmed { get; private set; } = false;
        public bool IsPhatSinh => rdoPhatSinh.Checked;
        public bool UseDT_CT => chkDT_CT.Checked;
        public bool UseH => chkH.Checked;
        public bool UseKM => chkKM.Checked;
        public bool UseCurve => chkCurve == null ? true : chkCurve.Checked;
        public bool UsePrefix => chkPrefix == null ? false : chkPrefix.Checked;
        public string PrefixValue => txtPrefix == null ? "" : txtPrefix.Text.Trim();
        public ObjectId SelectedAlignmentId { get; private set; } = ObjectId.Null;
        public double KhoangCach => (double)numKhoangCach.Value;
        public double KhoangCachCong => (double)numKhoangCachCong.Value;
        public double BeRongTrai => (double)numBeRongTrai.Value;
        public double BeRongPhai => (double)numBeRongPhai.Value;

        private class AlignmentInfo
        {
            public ObjectId Id { get; set; }
            public string Name { get; set; } = "";
            public double Length { get; set; }
            public double StartStation { get; set; }
            public double EndStation { get; set; }
            public int GroupCount { get; set; }
            public int SampleLineCount { get; set; }
        }

        public class CocPreviewItem
        {
            public int STT { get; set; }
            public string TenCoc { get; set; } = "";
            public double Station { get; set; }
            public string LyTrinh { get; set; } = "";
        }

        public List<CocPreviewItem> PreviewItems { get; private set; } = new();

        // ─── Constructor ───
        public CocTuyenDuongForm()
        {
            LoadAlignments();
            InitializeUI();
        }

        private void LoadAlignments()
        {
            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                CivilDocument cdoc = CivilApplication.ActiveDocument;
                ObjectIdCollection alIds = cdoc.GetAlignmentIds();
                foreach (ObjectId id in alIds)
                {
                    Alignment? al = tr.GetObject(id, OpenMode.ForRead) as Alignment;
                    if (al == null) continue;
                    int slCount = 0;
                    int groupCount = al.GetSampleLineGroupIds().Count;
                    if (groupCount > 0)
                    {
                        SampleLineGroup? slg = tr.GetObject(al.GetSampleLineGroupIds()[0], OpenMode.ForRead) as SampleLineGroup;
                        if (slg != null) slCount = slg.GetSampleLineIds().Count;
                    }
                    alignments.Add(new AlignmentInfo
                    {
                        Id = id,
                        Name = al.Name,
                        Length = al.Length,
                        StartStation = al.StartingStation,
                        EndStation = al.EndingStation,
                        GroupCount = groupCount,
                        SampleLineCount = slCount
                    });
                }
                tr.Commit();
            }
            catch { }
        }

        private void InitializeUI()
        {
            // ── Form ──
            this.Text = "CỌC TUYẾN ĐƯỜNG — Civil Tool";
            this.Size = new Size(780, 770);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = BgDark;
            this.ForeColor = TextWhite;
            this.Font = new Font("Segoe UI", 9.5f);

            // ── HEADER ──
            var header = new Panel { Dock = DockStyle.Top, Height = 52, BackColor = HeaderBlue };
            var lblTitle = new WinLabel
            {
                Text = "📌  PHÁT SINH CỌC & ĐỔI TÊN CỌC",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = TextWhite,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                Padding = new Padding(16, 0, 0, 0)
            };
            header.Controls.Add(lblTitle);
            this.Controls.Add(header);

            var subtitle = new WinLabel
            {
                Text = "Chọn chức năng: Phát sinh cọc mới hoặc Đổi tên cọc hiện có theo quy cách DT / 1~N / H / KM / CT",
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = TextDim,
                AutoSize = false,
                Height = 26,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Top,
                Padding = new Padding(18, 0, 0, 0),
                BackColor = Color.FromArgb(35, 35, 40)
            };
            this.Controls.Add(subtitle);
            subtitle.BringToFront();

            // ── MAIN CONTENT ──
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15),
                AutoScroll = true
            };
            this.Controls.Add(mainPanel);
            mainPanel.BringToFront();

            int y = 10;

            // ═══════════════════════════════════════════════════
            // SECTION 1: CHỌN CHỨC NĂNG
            // ═══════════════════════════════════════════════════
            var grpMode = CreateGroupBox("⚙  Chọn chức năng", 0, y, 730, 55);
            mainPanel.Controls.Add(grpMode);

            rdoPhatSinh = new RadioButton
            {
                Text = "🔨 Phát sinh cọc mới",
                Location = new Point(15, 25),
                AutoSize = true,
                ForeColor = GreenOk,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                Checked = true
            };
            rdoPhatSinh.CheckedChanged += (s, e) => OnModeChanged();
            grpMode.Controls.Add(rdoPhatSinh);

            rdoDoiTen = new RadioButton
            {
                Text = "✏️ Đổi tên cọc hiện có",
                Location = new Point(350, 25),
                AutoSize = true,
                ForeColor = OrangeWarn,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold)
            };
            grpMode.Controls.Add(rdoDoiTen);

            y += 62;

            // ═══════════════════════════════════════════════════
            // SECTION 2: CHỌN TUYẾN
            // ═══════════════════════════════════════════════════
            var grpAlign = CreateGroupBox("🛣  Chọn Tuyến (Alignment)", 0, y, 730, 75);
            mainPanel.Controls.Add(grpAlign);

            grpAlign.Controls.Add(CreateLabel("Tuyến:", 12, 28));
            cboAlignment = new ComboBox
            {
                Location = new Point(75, 24),
                Size = new Size(320, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextWhite,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f)
            };
            foreach (var al in alignments)
                cboAlignment.Items.Add(al.Name);
            if (cboAlignment.Items.Count > 0) cboAlignment.SelectedIndex = 0;
            cboAlignment.SelectedIndexChanged += (s, e) => OnAlignmentChanged();
            grpAlign.Controls.Add(cboAlignment);

            lblAlignmentInfo = CreateLabel("", 410, 25);
            lblAlignmentInfo.ForeColor = TextDim;
            lblAlignmentInfo.AutoSize = true;
            grpAlign.Controls.Add(lblAlignmentInfo);

            lblGroupInfo = CreateLabel("", 410, 48);
            lblGroupInfo.ForeColor = TextDim;
            lblGroupInfo.AutoSize = true;
            grpAlign.Controls.Add(lblGroupInfo);

            y += 82;

            // ═══════════════════════════════════════════════════
            // SECTION 3: THÔNG SỐ PHÁT SINH (chỉ hiện khi PhatSinh)
            // ═══════════════════════════════════════════════════
            pnlPhatSinh = CreateGroupBox("📏  Thông số phát sinh cọc", 0, y, 730, 140);
            mainPanel.Controls.Add(pnlPhatSinh);

            // Row 1: Khoảng cách
            pnlPhatSinh.Controls.Add(CreateLabel("K/cách cọc (Đ.Thẳng):", 12, 32));
            numKhoangCach = new NumericUpDown
            {
                Location = new Point(180, 28),
                Size = new Size(90, 28),
                Minimum = 5,
                Maximum = 200,
                Value = 25,
                DecimalPlaces = 1,
                Increment = 5,
                BackColor = BgInput,
                ForeColor = TextWhite
            };
            pnlPhatSinh.Controls.Add(numKhoangCach);

            pnlPhatSinh.Controls.Add(CreateLabel("K/cách cọc (Đ.Cong):", 330, 32));
            numKhoangCachCong = new NumericUpDown
            {
                Location = new Point(480, 28),
                Size = new Size(90, 28),
                Minimum = 1,
                Maximum = 200,
                Value = 10,
                DecimalPlaces = 1,
                Increment = 1,
                BackColor = BgInput,
                ForeColor = TextWhite
            };
            pnlPhatSinh.Controls.Add(numKhoangCachCong);

            // Row 2: Bề rộng
            pnlPhatSinh.Controls.Add(CreateLabel("Bề rộng cọc (Trái):", 12, 70));
            numBeRongTrai = new NumericUpDown
            {
                Location = new Point(160, 66),
                Size = new Size(90, 28),
                Minimum = 1,
                Maximum = 200,
                Value = 30,
                DecimalPlaces = 1,
                Increment = 1,
                BackColor = BgInput,
                ForeColor = TextWhite
            };
            pnlPhatSinh.Controls.Add(numBeRongTrai);

            pnlPhatSinh.Controls.Add(CreateLabel("Bề rộng cọc (Phải):", 330, 70));
            numBeRongPhai = new NumericUpDown
            {
                Location = new Point(480, 66),
                Size = new Size(90, 28),
                Minimum = 1,
                Maximum = 200,
                Value = 30,
                DecimalPlaces = 1,
                Increment = 1,
                BackColor = BgInput,
                ForeColor = TextWhite
            };
            pnlPhatSinh.Controls.Add(numBeRongPhai);
            chkCurve = CreateCheckBox("Phát sinh cọc trong đường cong (TĐ, P, TC)", 15, 105, true);
            chkCurve.CheckedChanged += (s, e) => GeneratePreview();
            pnlPhatSinh.Controls.Add(chkCurve);

            y += 148;

            // ═══════════════════════════════════════════════════
            // SECTION 3B: QUY CÁCH ĐẶT TÊN CỌC
            // ═══════════════════════════════════════════════════
            var grpNaming = CreateGroupBox("🏷  Quy cách đặt tên cọc", 0, y, 730, 95);
            mainPanel.Controls.Add(grpNaming);

            chkDT_CT = CreateCheckBox("DT / CT (Đầu Tuyến / Cuối Tuyến)", 15, 28, true);
            chkDT_CT.CheckedChanged += (s, e) => GeneratePreview();
            grpNaming.Controls.Add(chkDT_CT);

            chkH = CreateCheckBox("H1, H2... (mỗi 100m)", 310, 28, true);
            chkH.CheckedChanged += (s, e) => GeneratePreview();
            grpNaming.Controls.Add(chkH);

            chkKM = CreateCheckBox("KM1, KM2... (mỗi 1000m)", 520, 28, true);
            chkKM.CheckedChanged += (s, e) => GeneratePreview();
            grpNaming.Controls.Add(chkKM);

            chkPrefix = CreateCheckBox("Tiền tố cọc thường:", 15, 58, false);
            chkPrefix.CheckedChanged += (s, e) =>
            {
                txtPrefix.Enabled = chkPrefix.Checked;
                GeneratePreview();
            };
            grpNaming.Controls.Add(chkPrefix);

            txtPrefix = new TextBox
            {
                Location = new Point(195, 56),
                Size = new Size(100, 28),
                BackColor = BgInput,
                ForeColor = TextWhite,
                Font = new Font("Segoe UI", 9.5f),
                Enabled = false,
                Text = "C"
            };
            txtPrefix.TextChanged += (s, e) => GeneratePreview();
            grpNaming.Controls.Add(txtPrefix);

            y += 105;

            // ═══════════════════════════════════════════════════
            // SECTION 4: PREVIEW
            // ═══════════════════════════════════════════════════
            var grpPreview = CreateGroupBox("👁  Preview danh sách cọc", 0, y, 730, 310);
            mainPanel.Controls.Add(grpPreview);

            dgvPreview = new DataGridView
            {
                Location = new Point(10, 28),
                Size = new Size(600, 240),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = BgInput,
                GridColor = BorderColor,
                BorderStyle = BorderStyle.FixedSingle,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = BgInput,
                    ForeColor = TextWhite,
                    SelectionBackColor = AccentBlue,
                    SelectionForeColor = TextWhite,
                    Font = new Font("Segoe UI", 9f)
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(55, 55, 62),
                    ForeColor = TextWhite,
                    SelectionBackColor = AccentBlue,
                    SelectionForeColor = TextWhite
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(45, 48, 55),
                    ForeColor = TextWhite,
                    Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                EnableHeadersVisualStyles = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ColumnHeadersHeight = 30
            };

            dgvPreview.Columns.Add("colSTT", "STT");
            dgvPreview.Columns.Add("colTen", "Tên cọc");
            dgvPreview.Columns.Add("colStation", "Station (m)");
            dgvPreview.Columns.Add("colLyTrinh", "Lý trình");
            dgvPreview.Columns["colSTT"]!.Width = 50;
            dgvPreview.Columns["colTen"]!.Width = 120;
            grpPreview.Controls.Add(dgvPreview);

            btnPreview = new Button
            {
                Text = "🔄 Preview",
                Location = new Point(625, 28),
                Size = new Size(95, 40),
                FlatStyle = FlatStyle.Flat,
                BackColor = AccentBlue,
                ForeColor = TextWhite,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnPreview.FlatAppearance.BorderColor = Color.FromArgb(0, 80, 160);
            btnPreview.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 150, 230);
            btnPreview.Click += (s, e) => GeneratePreview();
            grpPreview.Controls.Add(btnPreview);

            lblTotal = CreateLabel("Tổng: 0 cọc", 10, 278);
            lblTotal.ForeColor = TextDim;
            lblTotal.AutoSize = true;
            grpPreview.Controls.Add(lblTotal);

            // ═══════════════════════════════════════════════════
            // BOTTOM BUTTONS
            // ═══════════════════════════════════════════════════
            var pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 55,
                BackColor = Color.FromArgb(38, 38, 42),
                Padding = new Padding(12, 8, 12, 8)
            };
            this.Controls.Add(pnlBottom);

            var btnCancel = CreateButton("✖ Hủy", 470, 8, 110, 36, pnlBottom);
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            var btnExecute = CreateButton("✅ Thực hiện", 600, 8, 130, 36, pnlBottom);
            btnExecute.BackColor = GreenOk;
            btnExecute.FlatAppearance.BorderColor = Color.FromArgb(35, 134, 54);
            btnExecute.Click += BtnExecute_Click;

            // ── Initialize ──
            if (alignments.Count > 0)
            {
                OnAlignmentChanged();
                GeneratePreview();
            }
        }

        // ═══════════════════════════════════════════════════
        // EVENT HANDLERS
        // ═══════════════════════════════════════════════════

        private void OnModeChanged()
        {
            pnlPhatSinh.Visible = rdoPhatSinh.Checked;
            GeneratePreview();
        }

        private void OnAlignmentChanged()
        {
            if (cboAlignment.SelectedIndex < 0 || cboAlignment.SelectedIndex >= alignments.Count) return;
            var al = alignments[cboAlignment.SelectedIndex];
            SelectedAlignmentId = al.Id;
            lblAlignmentInfo.Text = $"L = {al.Length:F1}m  |  Sta: {al.StartStation:F0} → {al.EndStation:F0}";
            lblGroupInfo.Text = $"Nhóm cọc: {al.GroupCount}  |  Số cọc hiện có: {al.SampleLineCount}";
        }

        private void BtnExecute_Click(object? sender, EventArgs e)
        {
            if (SelectedAlignmentId == ObjectId.Null)
            {
                MessageBox.Show("Vui lòng chọn tuyến!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (rdoDoiTen.Checked)
            {
                // Check có cọc không
                var al = alignments[cboAlignment.SelectedIndex];
                if (al.SampleLineCount == 0)
                {
                    MessageBox.Show("Tuyến này chưa có cọc!\nHãy chọn 'Phát sinh cọc mới' trước.",
                        "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            string mode = rdoPhatSinh.Checked ? "Phát sinh cọc mới" : "Đổi tên cọc hiện có";
            var confirm = MessageBox.Show(
                $"Chức năng: {mode}\n" +
                $"Tuyến: {alignments[cboAlignment.SelectedIndex].Name}\n" +
                $"Tổng: {PreviewItems.Count} cọc\n\n" +
                "Tiếp tục?",
                "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (confirm == DialogResult.Yes)
            {
                IsConfirmed = true;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        // ═══════════════════════════════════════════════════
        // PREVIEW LOGIC
        // ═══════════════════════════════════════════════════

        private void GeneratePreview()
        {
            PreviewItems.Clear();
            dgvPreview.Rows.Clear();

            if (cboAlignment.SelectedIndex < 0 || cboAlignment.SelectedIndex >= alignments.Count) return;
            var alInfo = alignments[cboAlignment.SelectedIndex];

            if (rdoPhatSinh.Checked)
                GeneratePreviewPhatSinh(alInfo);
            else
                GeneratePreviewDoiTen(alInfo);

            // Hiển thị preview
            foreach (var item in PreviewItems)
            {
                int idx = dgvPreview.Rows.Add(item.STT, item.TenCoc, $"{item.Station:F3}", item.LyTrinh);
                var row = dgvPreview.Rows[idx];

                // Tô màu theo loại
                if (item.TenCoc == "DT" || item.TenCoc == "CT")
                    row.DefaultCellStyle.ForeColor = GreenOk;
                else if (item.TenCoc.StartsWith("KM"))
                    row.DefaultCellStyle.ForeColor = Color.FromArgb(255, 180, 50);
                else if (item.TenCoc.StartsWith("H"))
                    row.DefaultCellStyle.ForeColor = Color.FromArgb(100, 200, 255);
            }

            lblTotal.Text = $"Tổng: {PreviewItems.Count} cọc";
            lblTotal.ForeColor = PreviewItems.Count > 0 ? GreenOk : TextDim;
        }

        /// <summary>
        /// Preview cho chức năng phát sinh cọc mới
        /// </summary>
        private void GeneratePreviewPhatSinh(AlignmentInfo alInfo)
        {
            double khoangCachThang = (double)numKhoangCach.Value;
            double khoangCachCong = (double)numKhoangCachCong.Value;
            
            double startSta = alInfo.StartStation;
            double endSta = alInfo.EndStation;

            List<double> stations = new() { startSta };
            var curveSegments = GetCurveSegments(alInfo.Id);

            // Iterate through alignment by steps
            double currentSta = startSta;
            while (currentSta < endSta)
            {
                bool isInCurve = curveSegments.Any(c => currentSta >= c.StartStation && currentSta <= c.EndStation);
                double step = isInCurve ? khoangCachCong : khoangCachThang;

                // Căn chỉnh step theo bội số để tròn số học giống Civil
                double nextSta = Math.Ceiling((currentSta + 0.01) / step) * step;
                if (nextSta > endSta) nextSta = endSta;

                if (!stations.Contains(nextSta))
                    stations.Add(nextSta);

                currentSta = nextSta;
            }

            if (Math.Abs(stations.Last() - endSta) > 0.01)
                stations.Add(endSta);

            var specialCurveStations = GetSpecialCurveStations(alInfo.Id);
            if (UseCurve)
            {
                foreach (var kvp in specialCurveStations)
                {
                    if (kvp.Key >= startSta && kvp.Key <= endSta)
                    {
                        if (!stations.Any(s => Math.Abs(s - kvp.Key) < 0.01))
                            stations.Add(kvp.Key);
                    }
                }
            }

            stations = stations.Distinct().OrderBy(s => s).ToList();

            // Đặt tên
            ApplyNamingConvention(stations, specialCurveStations);
        }

        /// <summary>
        /// Preview cho chức năng đổi tên cọc hiện có
        /// </summary>
        private void GeneratePreviewDoiTen(AlignmentInfo alInfo)
        {
            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                Alignment? alignment = tr.GetObject(alInfo.Id, OpenMode.ForRead) as Alignment;
                if (alignment == null) { tr.Commit(); return; }
                if (alignment.GetSampleLineGroupIds().Count == 0) { tr.Commit(); return; }

                SampleLineGroup? slg = tr.GetObject(alignment.GetSampleLineGroupIds()[0], OpenMode.ForRead) as SampleLineGroup;
                if (slg == null) { tr.Commit(); return; }

                List<double> stations = new();
                foreach (ObjectId slId in slg.GetSampleLineIds())
                {
                    SampleLine? sl = tr.GetObject(slId, OpenMode.ForRead) as SampleLine;
                    if (sl != null) stations.Add(sl.Station);
                }
                stations = stations.Distinct().OrderBy(s => s).ToList();

                var specialCurveStations = GetSpecialCurveStations(alInfo.Id);
                ApplyNamingConvention(stations, specialCurveStations);
                tr.Commit();
            }
            catch { }
        }

        /// <summary>
        /// Lấy danh sách các đoạn cong (StartStation, EndStation)
        /// </summary>
        private List<(double StartStation, double EndStation)> GetCurveSegments(ObjectId alignmentId)
        {
            List<(double, double)> curves = new();
            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                Alignment? al = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                if (al != null)
                {
                    foreach (AlignmentEntity entity in al.Entities)
                    {
                        if (entity is AlignmentArc arc)
                        {
                            curves.Add((arc.StartStation, arc.EndStation));
                        }
                        else if (entity is AlignmentSpiral spiral)
                        {
                            curves.Add((spiral.StartStation, spiral.EndStation));
                        }
                    }
                }
                tr.Commit();
            }
            catch { }
            return curves;
        }

        /// <summary>
        /// Lấy danh sách các Station đặc biệt trong đường cong (TĐ, P, TC)
        /// </summary>
        private Dictionary<double, string> GetSpecialCurveStations(ObjectId alignmentId)
        {
            Dictionary<double, string> dict = new();
            if (!UseCurve) return dict;

            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                Alignment? al = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                if (al != null)
                {
                    foreach (AlignmentEntity entity in al.Entities)
                    {
                        if (entity is AlignmentArc arc)
                        {
                            double td = arc.StartStation;
                            double tc = arc.EndStation;
                            double p = (td + tc) / 2.0;
                            
                            // Giữ lại tên nếu có sai số rất nhỏ
                            dict[td] = "TĐ";
                            dict[p] = "P";
                            dict[tc] = "TC";
                        }
                    }
                }
                tr.Commit();
            }
            catch { }
            return dict;
        }

        /// <summary>
        /// Đặt tên cọc theo quy cách: DT, CT, KM, H, TĐ, P, TC, 1~N
        /// </summary>
        private void ApplyNamingConvention(List<double> stations, Dictionary<double, string> specialStations)
        {
            if (stations.Count == 0) return;

            bool useDtCt = chkDT_CT.Checked;
            bool useH = chkH.Checked;
            bool useKm = chkKM.Checked;

            int normalIndex = 1;
            for (int i = 0; i < stations.Count; i++)
            {
                double sta = stations[i];
                string name;

                if (useDtCt && i == 0)
                    name = "DT";
                else if (useDtCt && i == stations.Count - 1)
                    name = "CT";
                else if (specialStations.Any(kvp => Math.Abs(kvp.Key - sta) < 0.01))
                    name = specialStations.First(kvp => Math.Abs(kvp.Key - sta) < 0.01).Value;
                else if (useKm && IsMultipleOf(sta, 1000))
                    name = $"KM{(int)Math.Round(sta / 1000)}";
                else if (useH && IsMultipleOf(sta, 100))
                {
                    int hVal = (int)Math.Round(sta / 100);
                    name = $"H{(useKm ? hVal % 10 : hVal)}";
                }
                else
                {
                    name = (UsePrefix ? PrefixValue : "") + normalIndex.ToString();
                    normalIndex++;
                }

                PreviewItems.Add(new CocPreviewItem
                {
                    STT = i + 1,
                    TenCoc = name,
                    Station = sta,
                    LyTrinh = FormatStation(sta)
                });
            }
        }

        private static bool IsMultipleOf(double station, double divisor)
        {
            double remainder = station % divisor;
            return remainder < 0.01 || (divisor - remainder) < 0.01;
        }

        private static string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double remainder = station - km * 1000;
            return $"Km{km}+{remainder:F3}";
        }

        // ═══════════════════════════════════════════════════
        // UI HELPERS
        // ═══════════════════════════════════════════════════

        private Panel CreateGroupBox(string title, int x, int y, int w, int h)
        {
            var panel = new Panel
            {
                Location = new Point(x, y),
                Size = new Size(w, h),
                BackColor = BgPanel,
                BorderStyle = BorderStyle.FixedSingle
            };
            var lbl = new WinLabel
            {
                Text = title,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                ForeColor = AccentBlue,
                Location = new Point(10, 5),
                AutoSize = true
            };
            panel.Controls.Add(lbl);
            return panel;
        }

        private WinLabel CreateLabel(string text, int x, int y)
        {
            return new WinLabel
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                ForeColor = TextWhite,
                Font = new Font("Segoe UI", 9.5f)
            };
        }

        private Button CreateButton(string text, int x, int y, int w, int h, Control parent)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(55, 55, 62),
                ForeColor = TextWhite,
                Font = new Font("Segoe UI", 9.5f),
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter
            };
            btn.FlatAppearance.BorderColor = BorderColor;
            btn.FlatAppearance.MouseOverBackColor = AccentBlue;
            parent.Controls.Add(btn);
            return btn;
        }

        private CheckBox CreateCheckBox(string text, int x, int y, bool isChecked)
        {
            return new CheckBox
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                Checked = isChecked,
                ForeColor = GreenOk,
                Font = new Font("Segoe UI", 9.5f),
                FlatStyle = FlatStyle.Standard
            };
        }
    }

    #endregion

    #region Command

    public class PhatSinhCocDoiTenCoc
    {
        /// <summary>
        /// Lệnh gộp: Phát sinh cọc hoặc Đổi tên cọc
        /// </summary>
        [CommandMethod("CTPO_CocTuyenDuong")]
        public static void CTPOCocTuyenDuong()
        {
            using CocTuyenDuongForm form = new();
            var result = AcadApp.ShowModalDialog(form);

            if (result != DialogResult.OK || !form.IsConfirmed) return;
            if (form.SelectedAlignmentId == ObjectId.Null) return;

            if (form.IsPhatSinh)
                ExecutePhatSinh(form);
            else
                ExecuteDoiTen(form);
        }

        [CommandMethod("CTS_DoiTenCoc")]
        public static void CTSDoiTenCocAlias()
        {
            CTPOCocTuyenDuong();
        }

        /// <summary>
        /// Phát sinh cọc mới dọc tuyến
        /// </summary>
        private static void ExecutePhatSinh(CocTuyenDuongForm form)
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                Alignment? alignment = tr.GetObject(form.SelectedAlignmentId, OpenMode.ForWrite) as Alignment;
                if (alignment == null) { tr.Commit(); return; }

                // Tạo hoặc lấy SampleLineGroup
                ObjectId slgId;
                if (alignment.GetSampleLineGroupIds().Count == 0)
                    slgId = SampleLineGroup.Create(alignment.Name, form.SelectedAlignmentId);
                else
                    slgId = alignment.GetSampleLineGroupIds()[0];

                // Tạo SampleLine từ preview
                int created = 0;
                foreach (var item in form.PreviewItems)
                {
                    try
                    {
                        string tempName = "z_" + Guid.NewGuid().ToString("N").Substring(0, 10);
                        ObjectId newId = CreateSamplelineWithWidth(tr, tempName, slgId, alignment, item.Station, form.BeRongTrai, form.BeRongPhai);
                        created++;
                    }
                    catch { /* Bỏ qua cọc trùng vị trí */ }
                }

                // Đổi tên tất cả cọc trong group theo quy cách
                SampleLineGroup? slGroup = tr.GetObject(slgId, OpenMode.ForWrite) as SampleLineGroup;
                if (slGroup != null)
                    RenameSampleLines(tr, alignment, slGroup, form.UseDT_CT, form.UseH, form.UseKM, form.UseCurve, form.UsePrefix, form.PrefixValue);

                A.Ed.WriteMessage($"\n✅ Đã phát sinh {created}/{form.PreviewItems.Count} cọc trên tuyến {alignment.Name}");
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Đổi tên cọc hiện có
        /// </summary>
        private static void ExecuteDoiTen(CocTuyenDuongForm form)
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                Alignment? alignment = tr.GetObject(form.SelectedAlignmentId, OpenMode.ForRead) as Alignment;
                if (alignment == null) { tr.Commit(); return; }

                ObjectIdCollection slgIds = alignment.GetSampleLineGroupIds();
                if (slgIds.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠ Tuyến này chưa có nhóm cọc!");
                    tr.Commit();
                    return;
                }

                SampleLineGroup? slGroup = tr.GetObject(slgIds[0], OpenMode.ForWrite) as SampleLineGroup;
                if (slGroup == null) { tr.Commit(); return; }

                int count = RenameSampleLines(tr, alignment, slGroup, form.UseDT_CT, form.UseH, form.UseKM, form.UseCurve, form.UsePrefix, form.PrefixValue);
                A.Ed.WriteMessage($"\n✅ Đã đổi tên {count} cọc trên tuyến {alignment.Name}");
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        #region Core Logic

        /// <summary>
        /// Đổi tên tất cả cọc trong SampleLineGroup theo quy cách:
        /// DT (đầu), CT (cuối), KM (1000m), H (100m), TĐ, P, TC, 1~N (còn lại)
        /// </summary>
        private static int RenameSampleLines(Transaction tr, Alignment al, SampleLineGroup slGroup,
            bool useDtCt = true, bool useH = true, bool useKm = true, bool useCurve = true, bool usePrefix = false, string prefixValue = "")
        {
            Dictionary<double, string> specialStations = new();
            if (useCurve)
            {
                if (al != null)
                {
                    foreach (AlignmentEntity entity in al.Entities)
                    {
                        if (entity is AlignmentArc arc)
                        {
                            specialStations[arc.StartStation] = "TĐ";
                            specialStations[(arc.StartStation + arc.EndStation) / 2.0] = "P";
                            specialStations[arc.EndStation] = "TC";
                        }
                    }
                }
            }

            // Thu thập và sắp xếp theo station
            List<SampleLine> sorted = new();
            foreach (ObjectId slId in slGroup.GetSampleLineIds())
            {
                SampleLine? sl = tr.GetObject(slId, OpenMode.ForWrite) as SampleLine;
                if (sl != null) sorted.Add(sl);
            }
            sorted = sorted.OrderBy(s => s.Station).ToList();

            if (sorted.Count == 0) return 0;

            // Bước 1: Reset tên tạm (tránh trùng khi rename)
            for (int i = 0; i < sorted.Count; i++)
                sorted[i].Name = "z_" + Guid.NewGuid().ToString("N").Substring(0, 10);

            // Bước 2: Đặt tên theo quy cách
            int normalIndex = 1;
            for (int i = 0; i < sorted.Count; i++)
            {
                double station = sorted[i].Station;
                string newName;

                if (useDtCt && i == 0)
                    newName = "DT";
                else if (useDtCt && i == sorted.Count - 1)
                    newName = "CT";
                else if (specialStations.Any(kvp => Math.Abs(kvp.Key - station) < 0.01))
                    newName = specialStations.First(kvp => Math.Abs(kvp.Key - station) < 0.01).Value;
                else if (useKm && IsMultipleOf(station, 1000))
                    newName = $"KM{(int)Math.Round(station / 1000)}";
                else if (useH && IsMultipleOf(station, 100))
                {
                    int hVal = (int)Math.Round(station / 100);
                    newName = $"H{(useKm ? hVal % 10 : hVal)}";
                }
                else
                {
                    newName = (usePrefix ? prefixValue : "") + normalIndex.ToString();
                    normalIndex++;
                }

                bool nameSet = false;
                string finalName = newName;
                int retryCount = 0;
                while (!nameSet)
                {
                    try
                    {
                        sorted[i].Name = finalName;
                        nameSet = true;
                    }
                    catch
                    {
                        retryCount++;
                        finalName = $"{newName}_{retryCount}";
                        if (retryCount > 100) break; // Fallback an toàn
                    }
                }
            }

            return sorted.Count;
        }

        private static bool IsMultipleOf(double station, double divisor)
        {
            double remainder = station % divisor;
            return remainder < 0.01 || (divisor - remainder) < 0.01;
        }

        /// <summary>
        /// Tạo SampleLine với bề rộng trái/phải tùy chỉnh
        /// </summary>
        private static ObjectId CreateSamplelineWithWidth(Transaction tr, string name, ObjectId slgId, Alignment alignment, double station, double widthLeft, double widthRight)
        {
            Autodesk.AutoCAD.Geometry.Point2dCollection points = new();
            double easting = 0, northing = 0;

            // Bên trái (giá trị âm)
            alignment.PointLocation(station, -widthLeft, ref easting, ref northing);
            points.Add(new Autodesk.AutoCAD.Geometry.Point2d(easting, northing));

            // Bên phải (giá trị dương)
            alignment.PointLocation(station, widthRight, ref easting, ref northing);
            points.Add(new Autodesk.AutoCAD.Geometry.Point2d(easting, northing));

            ObjectId slId = SampleLine.Create(name, slgId, points);
            return slId;
        }

        #endregion
    }

    #endregion
}
