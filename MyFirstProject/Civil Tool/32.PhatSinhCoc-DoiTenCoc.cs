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
using Autodesk.Civil.DatabaseServices.Styles;
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
    public class CocTuyenDuongForm : CivilFormBase
    {

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

        // Style controls
        private ComboBox cboSampleLineStyle = null!;
        private ComboBox cboLabelStyle = null!;

        // Naming options
        private CheckBox chkDT_CT = null!;
        private CheckBox chkH = null!;
        private CheckBox chkKM = null!;
        private CheckBox chkCurve = null!;
        private TextBox txtND = null!;
        private TextBox txtTD = null!;
        private TextBox txtP = null!;
        private TextBox txtTC = null!;
        private TextBox txtNC = null!;
        private CheckBox chkPrefix = null!;
        private TextBox txtPrefix = null!;

        // CogoPoint mode controls
        private RadioButton rdoCogoPoint = null!;
        private Panel pnlCogoPoint = null!;

        // Table mode controls
        private RadioButton rdoTable = null!;
        private Panel pnlTable = null!;
        private NumericUpDown numColTenCoc = null!;
        private NumericUpDown numColLyTrinh = null!;
        private NumericUpDown numRowStart = null!;

        // Preview
        private DataGridView dgvPreview = null!;
        private WinLabel lblTotal = null!;
        private Button btnPreview = null!;

        // ─── Data ───
        private List<AlignmentInfo> alignments = new();
        private List<string> sampleLineStyles = new();
        private List<string> labelStyles = new();
        public bool IsConfirmed { get; private set; } = false;
        public bool IsPhatSinh => rdoPhatSinh.Checked;
        public bool UseDT_CT => chkDT_CT.Checked;
        public bool UseH => chkH.Checked;
        public bool UseKM => chkKM.Checked;
        public bool UseCurve => chkCurve == null ? true : chkCurve.Checked;
        public bool UsePrefix => chkPrefix == null ? false : chkPrefix.Checked;
        public string PrefixValue => txtPrefix == null ? "" : txtPrefix.Text.Trim();
        public string StrND => txtND == null ? "NĐ" : (string.IsNullOrWhiteSpace(txtND.Text) ? "NĐ" : txtND.Text.Trim());
        public string StrTD => txtTD == null ? "TĐ" : (string.IsNullOrWhiteSpace(txtTD.Text) ? "TĐ" : txtTD.Text.Trim());
        public string StrP  => txtP == null  ? "P"  : (string.IsNullOrWhiteSpace(txtP.Text)  ? "P"  : txtP.Text.Trim());
        public string StrTC => txtTC == null ? "TC" : (string.IsNullOrWhiteSpace(txtTC.Text) ? "TC" : txtTC.Text.Trim());
        public string StrNC => txtNC == null ? "NC" : (string.IsNullOrWhiteSpace(txtNC.Text) ? "NC" : txtNC.Text.Trim());
        public bool IsCogoPointMode => rdoCogoPoint != null && rdoCogoPoint.Checked;
        public bool IsTableMode => rdoTable != null && rdoTable.Checked;
        public ObjectId SelectedAlignmentId { get; private set; } = ObjectId.Null;
        public double KhoangCach => (double)numKhoangCach.Value;
        public double KhoangCachCong => (double)numKhoangCachCong.Value;
        public double BeRongTrai => (double)numBeRongTrai.Value;
        public double BeRongPhai => (double)numBeRongPhai.Value;
        public int ColTenCoc => numColTenCoc != null ? (int)numColTenCoc.Value : 1;
        public int ColLyTrinh => numColLyTrinh != null ? (int)numColLyTrinh.Value : 2;
        public int RowStart => numRowStart != null ? (int)numRowStart.Value : 2;
        public string SelectedSampleLineStyle => cboSampleLineStyle != null && cboSampleLineStyle.SelectedIndex > 0
            ? cboSampleLineStyle.SelectedItem?.ToString() ?? "" : "";
        public string SelectedLabelStyle => cboLabelStyle != null && cboLabelStyle.SelectedIndex > 0
            ? cboLabelStyle.SelectedItem?.ToString() ?? "" : "";

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

                // Load SampleLine Styles
                try
                {
                    foreach (ObjectId sId in cdoc.Styles.SampleLineStyles)
                    {
                        var style = (StyleBase)tr.GetObject(sId, OpenMode.ForRead);
                        sampleLineStyles.Add(style.Name);
                    }
                }
                catch { }

                // Load SampleLine Label Styles
                try
                {
                    foreach (ObjectId lId in cdoc.Styles.LabelStyles.SampleLineLabelStyles.LabelStyles)
                    {
                        var style = (StyleBase)tr.GetObject(lId, OpenMode.ForRead);
                        labelStyles.Add(style.Name);
                    }
                }
                catch { }

                tr.Commit();
            }
            catch { }
        }

        private void InitializeUI()
        {
            // ── Form ──
            InitBaseTheme("PHÁT SINH CỌC & ĐỔI TÊN CỌC", 780, 850);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            // ── HEADER ──
            var header = CreateHeader("PHÁT SINH CỌC & ĐỔI TÊN CỌC");
            this.Controls.Add(header);

            var subtitle = new WinLabel
            {
                Text = "Chọn chức năng: Phát sinh cọc mới hoặc Đổi tên cọc hiện có theo quy cách DT / 1~N / H / KM / CT",
                Font = new Font("Segoe UI", 9.5f),
                ForeColor = TextDark,
                AutoSize = false,
                Height = 35,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Top,
                Padding = new Padding(15, 0, 0, 0),
                BackColor = BgPanel
            };
            this.Controls.Add(subtitle);
            subtitle.BringToFront();

            // ── MAIN CONTENT ──
            var mainPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15),
                AutoScroll = true,
                FlowDirection = System.Windows.Forms.FlowDirection.TopDown,
                WrapContents = false
            };
            this.Controls.Add(mainPanel);
            mainPanel.BringToFront();

            int y = 5;

            // ═══════════════════════════════════════════════════
            // SECTION 1: CHỌN CHỨC NĂNG
            // ═══════════════════════════════════════════════════
            var grpMode = CreateSection("Chọn chức năng", 0, 0, 730, 60);
            grpMode.Margin = new Padding(0, 5, 0, 10);
            mainPanel.Controls.Add(grpMode);

            rdoPhatSinh = new RadioButton
            {
                Text = "Phát sinh cọc mới",
                Location = new Point(15, 30),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Checked = true
            };
            rdoPhatSinh.CheckedChanged += (s, e) => OnModeChanged();
            grpMode.Controls.Add(rdoPhatSinh);

            rdoDoiTen = new RadioButton
            {
                Text = "Đổi tên cọc",
                Location = new Point(230, 30),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            grpMode.Controls.Add(rdoDoiTen);

            rdoCogoPoint = new RadioButton
            {
                Text = "Từ CogoPoint",
                Location = new Point(400, 30),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            rdoCogoPoint.CheckedChanged += (s, e) => OnModeChanged();
            grpMode.Controls.Add(rdoCogoPoint);

            rdoTable = new RadioButton
            {
                Text = "Từ Bảng",
                Location = new Point(580, 30),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            rdoTable.CheckedChanged += (s, e) => OnModeChanged();
            grpMode.Controls.Add(rdoTable);

            // ═══════════════════════════════════════════════════
            // SECTION 2: CHỌN TUYẾN
            // ═══════════════════════════════════════════════════
            var grpAlign = CreateSection("Chọn Tuyến (Alignment)", 0, 0, 730, 75);
            grpAlign.Margin = new Padding(0, 0, 0, 10);
            mainPanel.Controls.Add(grpAlign);

            grpAlign.Controls.Add(CreateLabel("Tuyến:", 12, 33));
            cboAlignment = new ComboBox
            {
                Location = new Point(75, 29),
                Size = new Size(320, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextDark,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f)
            };
            foreach (var al in alignments)
                cboAlignment.Items.Add(al.Name);
            if (cboAlignment.Items.Count > 0) cboAlignment.SelectedIndex = 0;
            cboAlignment.SelectedIndexChanged += (s, e) => OnAlignmentChanged();
            grpAlign.Controls.Add(cboAlignment);

            lblAlignmentInfo = CreateLabel("", 410, 30);
            lblAlignmentInfo.AutoSize = true;
            grpAlign.Controls.Add(lblAlignmentInfo);

            lblGroupInfo = CreateLabel("", 410, 50);
            lblGroupInfo.AutoSize = true;
            grpAlign.Controls.Add(lblGroupInfo);

            // ═══════════════════════════════════════════════════
            // SECTION 3: THÔNG SỐ PHÁT SINH (chỉ hiện khi PhatSinh)
            // ═══════════════════════════════════════════════════
            pnlPhatSinh = CreateSection("Thông số phát sinh cọc", 0, 0, 730, 110);
            pnlPhatSinh.Margin = new Padding(0, 0, 0, 10);
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
                ForeColor = TextDark
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
                ForeColor = TextDark
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
                ForeColor = TextDark
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
                ForeColor = TextDark
            };
            pnlPhatSinh.Controls.Add(numBeRongPhai);

            // ═══════════════════════════════════════════════════
            // SECTION 3C: TÙY CHỌN STYLE (áp dụng cho mọi chế độ)
            // ═══════════════════════════════════════════════════
            var grpStyle = CreateSection("Style hiển thị cọc", 0, 0, 730, 95);
            grpStyle.Margin = new Padding(0, 0, 0, 10);
            mainPanel.Controls.Add(grpStyle);

            // Style dao diện cọc
            grpStyle.Controls.Add(CreateLabel("Style cọc:", 12, 32));
            cboSampleLineStyle = new ComboBox
            {
                Location = new Point(110, 28),
                Size = new Size(230, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextDark,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f)
            };
            cboSampleLineStyle.Items.Add("<< Giữ nguyên >>");
            string defaultStyle = StyleSettings.Get().SampleLineStyle;
            int selStyleIdx = 0;
            for (int si = 0; si < sampleLineStyles.Count; si++)
            {
                cboSampleLineStyle.Items.Add(sampleLineStyles[si]);
                if (!string.IsNullOrEmpty(defaultStyle) &&
                    sampleLineStyles[si].Equals(defaultStyle, StringComparison.OrdinalIgnoreCase))
                    selStyleIdx = si + 1;
            }
            cboSampleLineStyle.SelectedIndex = selStyleIdx;
            grpStyle.Controls.Add(cboSampleLineStyle);

            // Label style
            grpStyle.Controls.Add(CreateLabel("Label cọc:", 370, 32));
            cboLabelStyle = new ComboBox
            {
                Location = new Point(460, 28),
                Size = new Size(250, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextDark,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f)
            };
            cboLabelStyle.Items.Add("<< Không tạo Label >>");
            string defaultLabel = StyleSettings.Get().SampleLineLabelStyle;
            int selLabelIdx = 0;
            for (int li = 0; li < labelStyles.Count; li++)
            {
                cboLabelStyle.Items.Add(labelStyles[li]);
                if (!string.IsNullOrEmpty(defaultLabel) &&
                    labelStyles[li].Equals(defaultLabel, StringComparison.OrdinalIgnoreCase))
                    selLabelIdx = li + 1;
            }
            cboLabelStyle.SelectedIndex = selLabelIdx;
            grpStyle.Controls.Add(cboLabelStyle);

            var lblStyleHint = CreateLabel("Style được lưu tại CT_StyleSettings", 15, 65);
            lblStyleHint.Font = new Font("Segoe UI", 8.5f, FontStyle.Italic);
            grpStyle.Controls.Add(lblStyleHint);

            // ═══════════════════════════════════════════════════
            // SECTION 3A: PANEL COGOPOINT
            // ═══════════════════════════════════════════════════
            pnlCogoPoint = CreateSection("Từ CogoPoint — chọn nhóm điểm trên bản vẽ", 0, 0, 730, 60);
            pnlCogoPoint.Margin = new Padding(0, 0, 0, 10);
            pnlCogoPoint.Visible = false;
            pnlCogoPoint.Controls.Add(CreateLabel("Chọn 1 CogoPoint thuộc nhóm cần phát sinh cọc (pick trên bản vẽ khi thực hiện)", 15, 30));
            mainPanel.Controls.Add(pnlCogoPoint);

            // ═══════════════════════════════════════════════════
            // SECTION 3B: PANEL TABLE
            // ═══════════════════════════════════════════════════
            pnlTable = CreateSection("Từ Bảng tọa độ (AutoCAD Table)", 0, 0, 730, 100);
            pnlTable.Margin = new Padding(0, 0, 0, 10);
            pnlTable.Visible = false;
            mainPanel.Controls.Add(pnlTable);

            pnlTable.Controls.Add(CreateLabel("Cột tên cọc:", 15, 32));
            numColTenCoc = new NumericUpDown
            {
                Location = new Point(120, 28), Size = new Size(60, 28),
                Minimum = 1, Maximum = 20, Value = 1,
                BackColor = BgInput, ForeColor = TextDark
            };
            pnlTable.Controls.Add(numColTenCoc);

            pnlTable.Controls.Add(CreateLabel("Cột lý trình:", 210, 32));
            numColLyTrinh = new NumericUpDown
            {
                Location = new Point(310, 28), Size = new Size(60, 28),
                Minimum = 1, Maximum = 20, Value = 2,
                BackColor = BgInput, ForeColor = TextDark
            };
            pnlTable.Controls.Add(numColLyTrinh);

            pnlTable.Controls.Add(CreateLabel("Hàng bắt đầu:", 410, 32));
            numRowStart = new NumericUpDown
            {
                Location = new Point(520, 28), Size = new Size(60, 28),
                Minimum = 1, Maximum = 1000, Value = 2,
                BackColor = BgInput, ForeColor = TextDark
            };
            pnlTable.Controls.Add(numRowStart);

            pnlTable.Controls.Add(CreateLabel("Chọn bảng tọa độ cọc (Table) và Alignment trên bản vẽ khi thực hiện", 15, 70));

            // ═══════════════════════════════════════════════════
            // SECTION 4: QUY CÁCH ĐẶT TÊN CỌC
            // ═══════════════════════════════════════════════════
            var grpNaming = CreateSection("Quy cách đặt tên cọc", 0, 0, 730, 135);
            grpNaming.Margin = new Padding(0, 0, 0, 10);
            mainPanel.Controls.Add(grpNaming);

            chkDT_CT = new CheckBox 
            {
                Text = "DT / CT (Đầu Tuyến / Cuối Tuyến)",
                Location = new Point(15, 30),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            chkDT_CT.CheckedChanged += (s, e) => GeneratePreview();
            grpNaming.Controls.Add(chkDT_CT);

            chkH = new CheckBox
            {
                Text = "H1, H2... (mỗi 100m)",
                Location = new Point(310, 30),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            chkH.CheckedChanged += (s, e) => GeneratePreview();
            grpNaming.Controls.Add(chkH);

            chkKM = new CheckBox
            {
                Text = "KM1, KM2... (mỗi 1000m)",
                Location = new Point(520, 30),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            chkKM.CheckedChanged += (s, e) => GeneratePreview();
            grpNaming.Controls.Add(chkKM);

            chkPrefix = new CheckBox
            {
                Text = "Tiền tố cọc thường:",
                Location = new Point(15, 60),
                AutoSize = true,
                Checked = false,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            chkPrefix.CheckedChanged += (s, e) =>
            {
                txtPrefix.Enabled = chkPrefix.Checked;
                GeneratePreview();
            };
            grpNaming.Controls.Add(chkPrefix);

            txtPrefix = new TextBox
            {
                Location = new Point(195, 58),
                Size = new Size(100, 28),
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f),
                Enabled = false,
                Text = "C"
            };
            txtPrefix.TextChanged += (s, e) => GeneratePreview();
            grpNaming.Controls.Add(txtPrefix);

            // Row 3: Checkbox đường cong
            chkCurve = new CheckBox
            {
                Text = "Cọc đường cong:",
                Location = new Point(15, 95),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            chkCurve.CheckedChanged += (s, e) => {
                bool isC = chkCurve.Checked;
                if(txtND != null) {
                    txtND.Enabled = isC; txtTD.Enabled = isC; txtP.Enabled = isC; txtTC.Enabled = isC; txtNC.Enabled = isC;
                }
                GeneratePreview();
            };
            grpNaming.Controls.Add(chkCurve);

            txtND = CreateTextBox(150, 92, 45, "NĐ");
            txtTD = CreateTextBox(205, 92, 45, "TĐ");
            txtP  = CreateTextBox(260, 92, 45, "P");
            txtTC = CreateTextBox(315, 92, 45, "TC");
            txtNC = CreateTextBox(370, 92, 45, "NC");

            txtND.TextChanged += (s, e) => GeneratePreview();
            txtTD.TextChanged += (s, e) => GeneratePreview();
            txtP.TextChanged += (s, e) => GeneratePreview();
            txtTC.TextChanged += (s, e) => GeneratePreview();
            txtNC.TextChanged += (s, e) => GeneratePreview();

            grpNaming.Controls.Add(txtND);
            grpNaming.Controls.Add(txtTD);
            grpNaming.Controls.Add(txtP);
            grpNaming.Controls.Add(txtTC);
            grpNaming.Controls.Add(txtNC);

            // ═══════════════════════════════════════════════════
            // SECTION 5: PREVIEW
            // ═══════════════════════════════════════════════════
            var grpPreview = CreateSection("Preview danh sách cọc", 0, 0, 730, 310);
            grpPreview.Margin = new Padding(0, 0, 0, 10);
            mainPanel.Controls.Add(grpPreview);

            dgvPreview = new DataGridView
            {
                Location = new Point(10, 30),
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
                    ForeColor = TextDark,
                    SelectionBackColor = NavyMedium,
                    SelectionForeColor = Color.White,
                    Font = new Font("Segoe UI", 9f)
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = NavyDark,
                    ForeColor = Color.White,
                    Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                EnableHeadersVisualStyles = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ColumnHeadersHeight = 32
            };

            dgvPreview.Columns.Add("colSTT", "STT");
            dgvPreview.Columns.Add("colTen", "Tên cọc");
            dgvPreview.Columns.Add("colStation", "Station (m)");
            dgvPreview.Columns.Add("colLyTrinh", "Lý trình");
            dgvPreview.Columns["colSTT"]!.Width = 50;
            dgvPreview.Columns["colTen"]!.Width = 120;
            grpPreview.Controls.Add(dgvPreview);

            btnPreview = CreateButton("Preview", 625, 30, 95, 40, grpPreview);
            btnPreview.Click += (s, e) => GeneratePreview();

            lblTotal = CreateLabel("Tổng: 0 cọc", 10, 280);
            lblTotal.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            grpPreview.Controls.Add(lblTotal);

            // ═══════════════════════════════════════════════════
            // BOTTOM BUTTONS
            // ═══════════════════════════════════════════════════
            var pnlBottom = CreateBottomBar();
            this.Controls.Add(pnlBottom);

            var btnCancel = CreateButton("Hủy", 510, 8, 110, 36, pnlBottom);
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            var btnExecute = CreatePrimaryButton("Thực hiện", 630, 8, 120, 36, pnlBottom);
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
            pnlCogoPoint.Visible = rdoCogoPoint.Checked;
            pnlTable.Visible = rdoTable.Checked;
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
                    row.DefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
                else if (item.TenCoc.StartsWith("KM"))
                    row.DefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Italic);
                else if (item.TenCoc.StartsWith("H"))
                    row.DefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Italic);
            }

            lblTotal.Text = $"Tổng: {PreviewItems.Count} cọc";
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

                // Loại bỏ cọc thông thường nằm quá gần cọc đặc biệt (< 0.5m)
                // Giữ lại cọc đặc biệt, bỏ cọc thường bị trùng
                stations = stations.Where(s =>
                {
                    // Nếu station này chính là cọc đặc biệt → giữ
                    if (specialCurveStations.Keys.Any(k => Math.Abs(k - s) < 0.01))
                        return true;
                    // Nếu là startSta hoặc endSta → giữ (DT/CT)
                    if (Math.Abs(s - startSta) < 0.01 || Math.Abs(s - endSta) < 0.01)
                        return true;
                    // Nếu nằm quá gần cọc đặc biệt khác → loại
                    if (specialCurveStations.Keys.Any(k => Math.Abs(k - s) < 0.5))
                        return false;
                    return true;
                }).ToList();
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
        /// Hỗ trợ: AlignmentArc, AlignmentSpiral, AlignmentSCS (compound)
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
                        if (entity is AlignmentSCS scs)
                        {
                            // SCS là compound entity: SpiralIn + Arc + SpiralOut
                            // Toàn bộ từ đầu SpiralIn đến cuối SpiralOut là đoạn cong
                            double scsStart = scs.SpiralIn.StartStation;
                            double scsEnd = scs.SpiralOut.EndStation;
                            curves.Add((scsStart, scsEnd));
                        }
                        else if (entity is AlignmentArc arc)
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
        /// Lấy danh sách các Station đặc biệt trong đường cong
        /// Hỗ trợ đầy đủ:
        ///   - AlignmentSCS (compound):  NĐ — TĐ — P — TC — NC
        ///   - Cung tròn đơn (Arc):      TĐ — P — TC
        ///   - Spiral đơn:               NĐ — NC
        /// </summary>
        private Dictionary<double, string> GetSpecialCurveStations(ObjectId alignmentId)
        {
            Dictionary<double, string> dict = new();
            if (!UseCurve) return dict;

            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                Alignment? al = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                if (al == null) { tr.Commit(); return dict; }

                var curveList = new List<(double startSta, AlignmentEntity entity)>();
                foreach (AlignmentEntity entity in al.Entities)
                {
                    if (entity is AlignmentSCS scs)
                        curveList.Add((scs.SpiralIn.StartStation, entity));
                    else if (entity is AlignmentArc arc)
                        curveList.Add((arc.StartStation, entity));
                    else if (entity is AlignmentSpiral spiral)
                        curveList.Add((spiral.StartStation, entity));
                }

                curveList = curveList.OrderBy(c => c.startSta).ToList();

                int curveIndex = 1; // Đánh số đường cong bắt đầu từ 1
                foreach (var item in curveList)
                {
                    AlignmentEntity entity = item.entity;
                    if (entity is AlignmentSCS scs)
                    {
                        // === ĐƯỜNG CONG CHUYỂN TIẾP SCS (compound entity) ===
                        string idx = curveIndex.ToString();
                        AddStation(dict, scs.SpiralIn.StartStation, $"{StrND}{idx}");   // Nối Đầu
                        AddStation(dict, scs.SpiralIn.EndStation, $"{StrTD}{idx}");     // Tiếp Đầu
                        double pStation = (scs.Arc.StartStation + scs.Arc.EndStation) / 2.0;
                        AddStation(dict, pStation, $"{StrP}{idx}");
                        AddStation(dict, scs.Arc.EndStation, $"{StrTC}{idx}");          // Tiếp Cuối
                        AddStation(dict, scs.SpiralOut.EndStation, $"{StrNC}{idx}");    // Nối Cuối
                        curveIndex++;
                    }
                    else if (entity is AlignmentArc arc)
                    {
                        // === CUNG TRÒN ĐƠN (không có Spiral) ===
                        string idx = curveIndex.ToString();
                        AddStation(dict, arc.StartStation, $"{StrTD}{idx}");
                        double pStation = (arc.StartStation + arc.EndStation) / 2.0;
                        AddStation(dict, pStation, $"{StrP}{idx}");
                        AddStation(dict, arc.EndStation, $"{StrTC}{idx}");
                        curveIndex++;
                    }
                    else if (entity is AlignmentSpiral spiral)
                    {
                        // === SPIRAL ĐƠN (không nằm trong SCS) ===
                        string idx = curveIndex.ToString();
                        AddStation(dict, spiral.StartStation, $"{StrND}{idx}");
                        AddStation(dict, spiral.EndStation, $"{StrNC}{idx}");
                        curveIndex++;
                    }
                }

                tr.Commit();
            }
            catch { }
            return dict;
        }

        /// <summary>
        /// Thêm station đặc biệt vào dictionary, tránh ghi đè nếu đã tồn tại
        /// </summary>
        private static void AddStation(Dictionary<double, string> dict, double station, string name)
        {
            // Kiểm tra trùng (sai số < 0.01m)
            bool exists = dict.Keys.Any(k => Math.Abs(k - station) < 0.01);
            if (!exists)
                dict[station] = name;
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

        // (Removed UI helpers)
    }

    #endregion

    #region Command

    public class PhatSinhCocDoiTenCoc
    {
        /// <summary>
        /// Lệnh gộp: Phát sinh cọc hoặc Đổi tên cọc
        /// </summary>
        [CommandMethod("CTS_PhatSinhCoc")]
        public static void CTSPhatSinhCoc()
        {
            using CocTuyenDuongForm form = new();
            var result = AcadApp.ShowModalDialog(form);

            if (result != DialogResult.OK || !form.IsConfirmed) return;

            if (form.IsCogoPointMode)
            {
                ExecuteFromCogoPoint(form);
                return;
            }
            if (form.IsTableMode)
            {
                ExecuteFromTable(form);
                return;
            }

            if (form.SelectedAlignmentId == ObjectId.Null) return;

            if (form.IsPhatSinh)
                ExecutePhatSinh(form);
            else
                ExecuteDoiTen(form);
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
                    RenameSampleLines(tr, alignment, slGroup, form.UseDT_CT, form.UseH, form.UseKM, form.UseCurve, form.UsePrefix, form.PrefixValue, form.StrND, form.StrTD, form.StrP, form.StrTC, form.StrNC);

                // Áp dụng Style cho SampleLine
                ApplyStyleAndLabel(tr, slgId, form.SelectedSampleLineStyle, form.SelectedLabelStyle);

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

                int count = RenameSampleLines(tr, alignment, slGroup, form.UseDT_CT, form.UseH, form.UseKM, form.UseCurve, form.UsePrefix, form.PrefixValue, form.StrND, form.StrTD, form.StrP, form.StrTC, form.StrNC);

                // Áp dụng Style cho SampleLine
                ApplyStyleAndLabel(tr, slgIds[0], form.SelectedSampleLineStyle, form.SelectedLabelStyle);

                A.Ed.WriteMessage($"\n✅ Đã đổi tên {count} cọc trên tuyến {alignment.Name}");
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Phát sinh cọc từ nhóm CogoPoint
        /// </summary>
        private static void ExecuteFromCogoPoint(CocTuyenDuongForm form)
        {
            try
            {
                // Pick trên bản vẽ sau khi form đóng
                ObjectId cogoPointId = UserInput.GCogoPointId("\n Chọn 1 CogoPoint thuộc nhóm điểm cần phát sinh cọc: \n");
                ObjectId alignmentId = UserInput.GAlignmentId("\n Chọn tim đường cần phát sinh cọc:");

                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                CogoPoint? cogoPoint = tr.GetObject(cogoPointId, OpenMode.ForRead) as CogoPoint;
                if (cogoPoint == null) { tr.Commit(); return; }

                ObjectId pointGroupId = cogoPoint.PrimaryPointGroupId;
                ObjectIdCollection pointIds = UtilitiesC3D.GPointIdsFromPointGroup(pointGroupId);

                Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForWrite) as Alignment;
                if (alignment == null) { tr.Commit(); return; }

                // Tạo hoặc lấy SampleLineGroup
                ObjectId slgId;
                if (alignment.GetSampleLineGroupIds().Count == 0)
                    slgId = SampleLineGroup.Create(alignment.Name, alignmentId);
                else
                    slgId = alignment.GetSampleLineGroupIds()[0];

                double station = 0, offset = 0;
                int created = 0;
                foreach (ObjectId pointId in pointIds)
                {
                    CogoPoint? pt = tr.GetObject(pointId, OpenMode.ForRead) as CogoPoint;
                    if (pt == null) continue;
                    alignment.StationOffset(pt.Easting, pt.Northing, ref station, ref offset);
                    try
                    {
                        string tempName = "z_" + Guid.NewGuid().ToString("N").Substring(0, 10);
                        CreateSamplelineWithWidth(tr, tempName, slgId, alignment, station, form.BeRongTrai, form.BeRongPhai);
                        created++;
                    }
                    catch { }
                }

                // Rename theo quy cách
                SampleLineGroup? slGroup = tr.GetObject(slgId, OpenMode.ForWrite) as SampleLineGroup;
                if (slGroup != null)
                    RenameSampleLines(tr, alignment, slGroup, form.UseDT_CT, form.UseH, form.UseKM, form.UseCurve, form.UsePrefix, form.PrefixValue, form.StrND, form.StrTD, form.StrP, form.StrTC, form.StrNC);

                // Áp dụng Style
                ApplyStyleAndLabel(tr, slgId, form.SelectedSampleLineStyle, form.SelectedLabelStyle);

                A.Ed.WriteMessage($"\n✅ Đã phát sinh {created} cọc từ CogoPoint trên tuyến {alignment.Name}");
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Phát sinh cọc từ bảng tọa độ (AutoCAD Table)
        /// </summary>
        private static void ExecuteFromTable(CocTuyenDuongForm form)
        {
            try
            {
                ObjectId bangId = UserInput.GTable("Chọn bảng tọa độ cọc có lý trình:");
                ObjectId alignmentId = UserInput.GAlignmentId("Chọn tim tuyến cần bổ sung cọc:");

                int soTenCoc = form.ColTenCoc;
                int soLyTrinh = form.ColLyTrinh;
                int soHangCoc = form.RowStart;

                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForWrite) as Alignment;
                Autodesk.AutoCAD.DatabaseServices.Table? bang = tr.GetObject(bangId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Table;
                if (alignment == null || bang == null) { tr.Commit(); return; }

                // Tạo hoặc lấy SampleLineGroup
                ObjectId slgId;
                if (alignment.GetSampleLineGroupIds().Count == 0)
                    slgId = SampleLineGroup.Create(alignment.Name, alignment.ObjectId);
                else
                    slgId = alignment.GetSampleLineGroupIds()[0];

                List<string> listTenCoc = new();
                List<double> listLyTrinh = new();
                int soCocThem = 1;

                for (int i = soHangCoc - 1; i < bang.Rows.Count; i++)
                {
                    string tenCoc = bang.Cells[i, soTenCoc - 1].GetTextString(Autodesk.AutoCAD.DatabaseServices.FormatOption.IgnoreMtextFormat);
                    listTenCoc.Add(tenCoc);
                    // Check duplicate
                    for (int j = 0; j < listTenCoc.Count - 1; j++)
                    {
                        if (tenCoc == listTenCoc[j])
                        {
                            listTenCoc[j] = tenCoc + soCocThem.ToString();
                            soCocThem++;
                        }
                    }
                    listLyTrinh.Add(Convert.ToDouble(bang.Cells[i, soLyTrinh - 1].GetTextString(Autodesk.AutoCAD.DatabaseServices.FormatOption.IgnoreMtextFormat)));
                }

                if (listLyTrinh.Count > 0)
                    listLyTrinh[listLyTrinh.Count - 1] -= 0.01;

                int created = 0;
                for (int i = 0; i < listTenCoc.Count; i++)
                {
                    try
                    {
                        CreateSamplelineWithWidth(tr, listTenCoc[i], slgId, alignment, listLyTrinh[i], form.BeRongTrai, form.BeRongPhai);
                        created++;
                    }
                    catch { }
                }

                A.Ed.WriteMessage($"\n✅ Đã phát sinh {created}/{listTenCoc.Count} cọc từ bảng trên tuyến {alignment.Name}");
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
            bool useDtCt, bool useH, bool useKm, bool useCurve, bool usePrefix, string prefixValue,
            string strND, string strTD, string strP, string strTC, string strNC)
        {
            Dictionary<double, string> specialStations = new();
            if (useCurve && al != null)
            {
                var curveList = new List<(double startSta, AlignmentEntity entity)>();
                foreach (AlignmentEntity entity in al.Entities)
                {
                    if (entity is AlignmentSCS scs)
                        curveList.Add((scs.SpiralIn.StartStation, entity));
                    else if (entity is AlignmentArc arc)
                        curveList.Add((arc.StartStation, entity));
                    else if (entity is AlignmentSpiral spiral)
                        curveList.Add((spiral.StartStation, entity));
                }

                curveList = curveList.OrderBy(c => c.startSta).ToList();

                int curveIndex = 1; // Đánh số đường cong bắt đầu từ 1
                foreach (var item in curveList)
                {
                    AlignmentEntity entity = item.entity;
                    if (entity is AlignmentSCS scs)
                    {
                        // === ĐƯỜNG CONG CHUYỂN TIẾP SCS (compound: SpiralIn → Arc → SpiralOut) ===
                        string idx = curveIndex.ToString();
                        AddStationStatic(specialStations, scs.SpiralIn.StartStation, $"{strND}{idx}");
                        AddStationStatic(specialStations, scs.SpiralIn.EndStation, $"{strTD}{idx}");
                        double pStation = (scs.Arc.StartStation + scs.Arc.EndStation) / 2.0;
                        AddStationStatic(specialStations, pStation, $"{strP}{idx}");
                        AddStationStatic(specialStations, scs.Arc.EndStation, $"{strTC}{idx}");
                        AddStationStatic(specialStations, scs.SpiralOut.EndStation, $"{strNC}{idx}");
                        curveIndex++;
                    }
                    else if (entity is AlignmentArc arc)
                    {
                        // === CUNG TRÒN ĐƠN (không nằm trong SCS) ===
                        string idx = curveIndex.ToString();
                        AddStationStatic(specialStations, arc.StartStation, $"{strTD}{idx}");
                        AddStationStatic(specialStations, (arc.StartStation + arc.EndStation) / 2.0, $"{strP}{idx}");
                        AddStationStatic(specialStations, arc.EndStation, $"{strTC}{idx}");
                        curveIndex++;
                    }
                    else if (entity is AlignmentSpiral spiral)
                    {
                        // === SPIRAL ĐƠN (không nằm trong SCS) ===
                        string idx = curveIndex.ToString();
                        AddStationStatic(specialStations, spiral.StartStation, $"{strND}{idx}");
                        AddStationStatic(specialStations, spiral.EndStation, $"{strNC}{idx}");
                        curveIndex++;
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
        /// Thêm station đặc biệt vào dictionary (static version cho RenameSampleLines)
        /// </summary>
        private static void AddStationStatic(Dictionary<double, string> dict, double station, string name)
        {
            bool exists = dict.Keys.Any(k => Math.Abs(k - station) < 0.01);
            if (!exists)
                dict[station] = name;
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

        /// <summary>
        /// Áp dụng Style dao diện và Label cho tất cả SampleLine trong group
        /// </summary>
        private static void ApplyStyleAndLabel(Transaction tr, ObjectId slgId, string styleName, string labelStyleName)
        {
            CivilDocument cdoc = CivilApplication.ActiveDocument;

            // 1. Áp dụng SampleLine Style
            if (!string.IsNullOrEmpty(styleName))
            {
                try
                {
                    SampleLineGroup? slGroup = tr.GetObject(slgId, OpenMode.ForRead) as SampleLineGroup;
                    if (slGroup != null)
                    {
                        foreach (ObjectId slId in slGroup.GetSampleLineIds())
                        {
                            SampleLine? sl = tr.GetObject(slId, OpenMode.ForWrite) as SampleLine;
                            if (sl != null)
                            {
                                try { sl.StyleName = styleName; }
                                catch { /* Style không tồn tại — bỏ qua */ }
                            }
                        }
                    }
                }
                catch { }
            }

            // 2. Tạo Label Group nếu có chọn Label Style
            if (!string.IsNullOrEmpty(labelStyleName))
            {
                try
                {
                    ObjectId labelStyleId = ObjectId.Null;
                    foreach (ObjectId lsId in cdoc.Styles.LabelStyles.SampleLineLabelStyles.LabelStyles)
                    {
                        var ls = (StyleBase)tr.GetObject(lsId, OpenMode.ForRead);
                        if (ls.Name.Equals(labelStyleName, StringComparison.OrdinalIgnoreCase))
                        {
                            labelStyleId = lsId;
                            break;
                        }
                    }

                    if (labelStyleId != ObjectId.Null)
                    {
                        SampleLineLabelGroup.Create(slgId, labelStyleId);
                    }
                }
                catch { }
            }
        }

        #endregion
    }

    #endregion
}
