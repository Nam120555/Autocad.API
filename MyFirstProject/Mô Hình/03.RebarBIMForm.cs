// RebarBIMForm.cs — Form nhập thông tin cấu kiện & Bảng thông số cốt thép cho lệnh AT_RebarBIM
// Kế thừa CivilFormBase — Theme: Navy / White

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using ContentAlignment = System.Drawing.ContentAlignment;
using FontStyle = System.Drawing.FontStyle;

namespace MyFirstProject.MoHinh
{
    /// <summary>
    /// Item dữ liệu đại diện cho 1 nhóm thanh thép trong bảng thông số.
    /// </summary>
    public class RebarTableItem
    {
        public string GroupCode { get; set; } = "G1";
        public string Name { get; set; } = "Thép chủ";
        public int Diameter { get; set; } = 28;
        public int Quantity { get; set; } = 36;
        public double Length { get; set; } = 11580; // mm
        public string ShapeType { get; set; } = "Straight"; // Straight, Stirrup, LBend, UBend, Custom
        public string Layer { get; set; } = "THEP_D28";
        public int[] ColorRgb { get; set; } = [30, 144, 255];

        public double TotalWeight => Math.Round(Quantity * Length * 0.006165 * Diameter * Diameter / 1000.0, 1);
    }

    /// <summary>
    /// Form nhập thông số cấu kiện & Bảng cốt thép khoa học cho mô hình BIM 3D.
    /// Thiết kế phân tab: Hình học cấu kiện + Bảng thông số cốt thép (DataGridView) + Cài đặt xuất file.
    /// </summary>
    public class RebarBIMForm : CivilFormBase
    {
        // ═══ KẾT QUẢ TRẢ VỀ ═══
        public bool IsConfirmed { get; private set; } = false;
        public string MemberName { get; private set; } = "T3_XaMu";
        public string MemberType { get; private set; } = "XaMu";
        public double MemberWidth { get; private set; } = 11700;  // Dài tổng (X)
        public double MemberHeight { get; private set; } = 1600;  // Cao tổng (Y)
        public double MemberDepth { get; private set; } = 2500;   // Rộng đáy/bề dày (Z)
        public double CantileverLen { get; private set; } = 2200; // Chiều dài vát vai
        public double CantileverH { get; private set; } = 700;   // Chiều cao vát vai
        public int PedestalCount { get; private set; } = 8;       // Số lượng bệ gối
        public double ConcreteCover { get; private set; } = 60;   // Lớp bảo vệ (mm)
        public bool UseTableDataMode { get; private set; } = true; // true: tạo từ bảng; false: quét CAD
        public bool SelectArea { get; private set; } = false;
        public string OutputFolder { get; private set; } = "";
        public bool AutoOpenHtml { get; private set; } = true;

        public List<RebarTableItem> RebarItems { get; } = new();

        // ═══ CONTROLS CHÍNH ═══
        private TabControl tabControl = null!;
        private TextBox txtName = null!;
        private ComboBox cboType = null!;
        private NumericUpDown numWidth = null!, numHeight = null!, numDepth = null!;
        private NumericUpDown numCantiLen = null!, numCantiH = null!;
        private NumericUpDown numPedestals = null!, numCover = null!;
        private Panel pnlPreview = null!;
        private WinLabel lblPreviewVol = null!;

        // Controls Tab Bảng Cốt Thép
        private DataGridView dgvRebar = null!;
        private WinLabel lblSummaryBars = null!;
        private WinLabel lblSummaryWeight = null!;
        private WinLabel lblSummaryRatio = null!;

        // Controls Tab Cài Đặt Xuất
        private RadioButton rdoSourceTable = null!, rdoSourceCad = null!;
        private RadioButton rdoScopeAll = null!, rdoScopeSelect = null!;
        private TextBox txtOutputFolder = null!;
        private CheckBox chkAutoOpen = null!;

        // Loại cấu kiện
        private static readonly (string Key, string Label, string Icon, string Desc)[] MemberTypes =
        [
            ("XaMu",   "Xà Mũ Trụ Cầu (Cap Beam)", "🏗️", "Cấu kiện dầm mũ liên kết đỉnh trụ"),
            ("Dam",    "Dầm Cầu / Dầm Nhà (Beam)",  "📏", "Cấu kiện chịu uốn Super-T / I / Hộp"),
            ("Cot",    "Cột / Trụ Đỡ (Column)",     "🔲", "Cấu kiện chịu nén đứng"),
            ("BanSan", "Bản Mặt Cầu / Sàn (Slab)",  "📋", "Cấu kiện tấm phẳng"),
            ("Tuong",  "Tường Thân Mố (Abutment)",  "🧱", "Cấu kiện tường chắn chịu lực")
        ];

        public RebarBIMForm(string defaultName, string defaultOutputFolder)
        {
            MemberName = string.IsNullOrWhiteSpace(defaultName) ? "T3_XaMu" : defaultName;
            OutputFolder = defaultOutputFolder;

            InitForm();
            LoadDefaultPreset(); // Nạp sẵn bộ mẫu xà mũ T3 chuẩn
            UpdateSummaryStats();
        }

        private void InitForm()
        {
            InitBaseTheme("🏗️  RebarBIM — Bảng Nhập Thông Số Cấu Kiện & Cốt Thép 3D", 1020, 720);
            this.MinimumSize = new Size(960, 650);
            this.StartPosition = FormStartPosition.CenterScreen;

            // ═══ HEADER ═══
            var header = CreateHeaderPanel();
            this.Controls.Add(header);

            // ═══ TAB CONTROL CHÍNH ═══
            tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10f, FontStyle.Regular),
                Padding = new Point(16, 8)
            };

            // Tab 1: Hình học & Bệ gối
            var tabGeo = new TabPage("  📐  1. Hình Học & Bệ Gối  ") { BackColor = BgWhite };
            BuildGeometryTab(tabGeo);
            tabControl.TabPages.Add(tabGeo);

            // Tab 2: Bảng thông số cốt thép
            var tabRebar = new TabPage("  📊  2. Bảng Cốt Thép Chi Tiết  ") { BackColor = BgWhite };
            BuildRebarTab(tabRebar);
            tabControl.TabPages.Add(tabRebar);

            // Tab 3: Cài đặt & Xuất file
            var tabExport = new TabPage("  📂  3. Nguồn Quét & Xuất File  ") { BackColor = BgWhite };
            BuildExportTab(tabExport);
            tabControl.TabPages.Add(tabExport);

            this.Controls.Add(tabControl);
            tabControl.BringToFront();

            // ═══ BOTTOM BAR ═══
            var bottomBar = CreateBottomBarPanel();
            this.Controls.Add(bottomBar);
            bottomBar.BringToFront();
        }

        // ════════════════════════════════════════════════════════════════
        // TAB 1: THÔNG SỐ HÌNH HỌC & PREVIEW 3D
        // ════════════════════════════════════════════════════════════════
        private void BuildGeometryTab(TabPage page)
        {
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 560,
                IsSplitterFixed = true,
                BackColor = BorderLight,
                Padding = new Padding(0)
            };
            split.Panel1.BackColor = BgWhite;
            split.Panel2.BackColor = Color.FromArgb(248, 250, 252);
            page.Controls.Add(split);

            // ── CỘT TRÁI: CÁC THÔNG SỐ NHẬP ──
            var scrollPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(16, 12, 16, 12)
            };
            split.Panel1.Controls.Add(scrollPanel);

            int y = 8;

            // Nhóm 1: Tên & Loại cấu kiện
            var grpInfo = CreateGroupBoxPanel("📋  THÔNG TIN CẤU KIỆN", scrollPanel, ref y);
            grpInfo.Controls.Add(CreateLabel("Tên cấu kiện:", 16, 28));
            txtName = CreateTextBox(140, 25, 360, MemberName);
            grpInfo.Controls.Add(txtName);

            grpInfo.Controls.Add(CreateLabel("Loại cấu kiện:", 16, 62));
            cboType = new ComboBox
            {
                Location = new Point(140, 59),
                Size = new Size(360, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9.5f)
            };
            foreach (var (k, l, icon, d) in MemberTypes)
                cboType.Items.Add($"{icon}  {l}");
            cboType.SelectedIndex = 0;
            cboType.SelectedIndexChanged += OnMemberTypeChanged;
            grpInfo.Controls.Add(cboType);
            grpInfo.Height = 100;

            // Nhóm 2: Kích thước chính (mm)
            var grpDim = CreateGroupBoxPanel("📐  KÍCH THƯỚC CHÍNH (mm)", scrollPanel, ref y);
            int iy = 26;
            grpDim.Controls.Add(CreateDimLabel("Chiều dài tổng (L):", 16, iy));
            numWidth = CreateNumericInput(210, iy - 2, 140, (decimal)MemberWidth);
            numWidth.ValueChanged += (s, e) => { MemberWidth = (double)numWidth.Value; UpdatePreview(); };
            grpDim.Controls.Add(numWidth);
            grpDim.Controls.Add(CreateUnitLabel("mm (phương dọc X)", 360, iy));
            iy += 34;

            grpDim.Controls.Add(CreateDimLabel("Chiều cao dầm (H):", 16, iy));
            numHeight = CreateNumericInput(210, iy - 2, 140, (decimal)MemberHeight);
            numHeight.ValueChanged += (s, e) => { MemberHeight = (double)numHeight.Value; UpdatePreview(); };
            grpDim.Controls.Add(numHeight);
            grpDim.Controls.Add(CreateUnitLabel("mm (phương đứng Y)", 360, iy));
            iy += 34;

            grpDim.Controls.Add(CreateDimLabel("Bề rộng/sâu đáy (B):", 16, iy));
            numDepth = CreateNumericInput(210, iy - 2, 140, (decimal)MemberDepth);
            numDepth.ValueChanged += (s, e) => { MemberDepth = (double)numDepth.Value; UpdatePreview(); };
            grpDim.Controls.Add(numDepth);
            grpDim.Controls.Add(CreateUnitLabel("mm (phương ngang Z)", 360, iy));
            grpDim.Height = iy + 36;

            // Nhóm 3: Chi tiết vát vai & Bệ gối
            var grpDetail = CreateGroupBoxPanel("🏛️  CHI TIẾT VÁT VAI & ĐÁ KÊ GỐI", scrollPanel, ref y);
            iy = 26;
            grpDetail.Controls.Add(CreateDimLabel("Chiều dài vát vai (L_vat):", 16, iy));
            numCantiLen = CreateNumericInput(210, iy - 2, 140, (decimal)CantileverLen);
            numCantiLen.ValueChanged += (s, e) => { CantileverLen = (double)numCantiLen.Value; UpdatePreview(); };
            grpDetail.Controls.Add(numCantiLen);
            grpDetail.Controls.Add(CreateUnitLabel("mm (độ vươn cantilever)", 360, iy));
            iy += 34;

            grpDetail.Controls.Add(CreateDimLabel("Chiều cao vát (H_vat):", 16, iy));
            numCantiH = CreateNumericInput(210, iy - 2, 140, (decimal)CantileverH);
            numCantiH.ValueChanged += (s, e) => { CantileverH = (double)numCantiH.Value; UpdatePreview(); };
            grpDetail.Controls.Add(numCantiH);
            grpDetail.Controls.Add(CreateUnitLabel("mm (chiều cao giảm ở đầu)", 360, iy));
            iy += 34;

            grpDetail.Controls.Add(CreateDimLabel("Số lượng đá kê gối:", 16, iy));
            numPedestals = CreateNumericInput(210, iy - 2, 140, PedestalCount, 0, 32);
            numPedestals.ValueChanged += (s, e) => { PedestalCount = (int)numPedestals.Value; UpdatePreview(); };
            grpDetail.Controls.Add(numPedestals);
            grpDetail.Controls.Add(CreateUnitLabel("gối (chia 2 hàng G1-G8)", 360, iy));
            iy += 34;

            grpDetail.Controls.Add(CreateDimLabel("Lớp bảo vệ bê tông (c):", 16, iy));
            numCover = CreateNumericInput(210, iy - 2, 140, (decimal)ConcreteCover, 20, 150);
            numCover.ValueChanged += (s, e) => { ConcreteCover = (double)numCover.Value; };
            grpDetail.Controls.Add(numCover);
            grpDetail.Controls.Add(CreateUnitLabel("mm", 360, iy));
            grpDetail.Height = iy + 36;

            // ── CỘT PHẢI: PREVIEW HÌNH HỌC 3D ISOMETRIC ──
            var pnlRight = split.Panel2;

            var lblPreviewTitle = new WinLabel
            {
                Text = "TRỰC QUAN HÌNH HỌC CẤU KIỆN 3D",
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                ForeColor = NavyDark,
                Location = new Point(16, 16),
                AutoSize = true
            };
            pnlRight.Controls.Add(lblPreviewTitle);

            pnlPreview = new Panel
            {
                Location = new Point(16, 46),
                Size = new Size(pnlRight.ClientSize.Width - 32, pnlRight.ClientSize.Height - 110),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            pnlPreview.Paint += DrawIsometricPreview;
            pnlRight.Controls.Add(pnlPreview);

            lblPreviewVol = new WinLabel
            {
                Text = "Thể tích bê tông: V = 0.000 m³",
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                ForeColor = Color.FromArgb(16, 149, 193),
                Location = new Point(16, pnlRight.ClientSize.Height - 50),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                AutoSize = true
            };
            pnlRight.Controls.Add(lblPreviewVol);

            UpdatePreview();
        }

        // ════════════════════════════════════════════════════════════════
        // TAB 2: BẢNG NHẬP THÔNG SỐ CỐT THÉP (DATAGRIDVIEW)
        // ════════════════════════════════════════════════════════════════
        private void BuildRebarTab(TabPage page)
        {
            var mainContainer = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(16, 12, 16, 12),
                BackColor = BgWhite
            };
            page.Controls.Add(mainContainer);

            // ── TOOLBAR PHÍA TRÊN BẢNG ──
            var toolBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 44,
                BackColor = Color.FromArgb(245, 248, 252),
                Padding = new Padding(8, 5, 8, 5),
                BorderStyle = BorderStyle.FixedSingle
            };
            mainContainer.Controls.Add(toolBar);

            int bx = 8;
            var btnAdd = CreateToolButton("➕ Thêm thanh", bx, 5, 110, BtnAddRow_Click);
            toolBar.Controls.Add(btnAdd);
            bx += 116;

            var btnDelete = CreateToolButton("➖ Xóa thanh", bx, 5, 100, BtnDeleteRow_Click);
            toolBar.Controls.Add(btnDelete);
            bx += 116;

            var btnPresetXaMu = CreateToolButton("📋 Mẫu Xà Mũ T3", bx, 5, 130, (s, e) => LoadPresetXaMu());
            toolBar.Controls.Add(btnPresetXaMu);
            bx += 136;

            var btnPresetDam = CreateToolButton("📏 Mẫu Dầm I/Super-T", bx, 5, 150, (s, e) => LoadPresetDam());
            toolBar.Controls.Add(btnPresetDam);
            bx += 156;

            var btnScanCad = CreateToolButton("🔍 Quét nhanh từ CAD", bx, 5, 150, BtnScanCadToTable_Click);
            toolBar.Controls.Add(btnScanCad);

            // ── DATAGRIDVIEW BẢNG CỐT THÉP ──
            dgvRebar = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.Fixed3D,
                GridColor = Color.FromArgb(226, 232, 240),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                EnableHeadersVisualStyles = false,
                Font = new Font("Segoe UI", 9.5f),
                RowTemplate = { Height = 28 }
            };

            // Header Style Navy
            dgvRebar.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = NavyDark,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                WrapMode = DataGridViewTriState.False
            };
            dgvRebar.ColumnHeadersHeight = 36;
            dgvRebar.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            // Columns
            dgvRebar.Columns.Add(new DataGridViewTextBoxColumn { Name = "colGroup", HeaderText = "Nhóm", Width = 70, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) } });
            dgvRebar.Columns.Add(new DataGridViewTextBoxColumn { Name = "colDesc", HeaderText = "Tên Thanh / Chức Năng", Width = 210 });
            dgvRebar.Columns.Add(new DataGridViewTextBoxColumn { Name = "colDia", HeaderText = "Ø (mm)", Width = 75, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvRebar.Columns.Add(new DataGridViewTextBoxColumn { Name = "colQty", HeaderText = "Số Lượng", Width = 85, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvRebar.Columns.Add(new DataGridViewTextBoxColumn { Name = "colLen", HeaderText = "Chiều Dài L (mm)", Width = 130, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvRebar.Columns.Add(new DataGridViewTextBoxColumn { Name = "colShape", HeaderText = "Dạng Uốn", Width = 100, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter } });
            dgvRebar.Columns.Add(new DataGridViewTextBoxColumn { Name = "colLayer", HeaderText = "Layer CAD", Width = 130 });
            dgvRebar.Columns.Add(new DataGridViewTextBoxColumn { Name = "colWeight", HeaderText = "Khối Lượng (kg)", Width = 120, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), ForeColor = Color.FromArgb(30, 64, 175) } });

            dgvRebar.CellValueChanged += DgvRebar_CellValueChanged;

            mainContainer.Controls.Add(dgvRebar);
            dgvRebar.BringToFront();

            // ── SUMMARY FOOTER BAR ──
            var summaryBar = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 38,
                BackColor = Color.FromArgb(241, 245, 249),
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(16, 8, 16, 8)
            };
            mainContainer.Controls.Add(summaryBar);
            summaryBar.BringToFront();

            lblSummaryBars = new WinLabel { Text = "Tổng số thanh: 0", Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), ForeColor = NavyDark, Location = new Point(16, 8), AutoSize = true };
            summaryBar.Controls.Add(lblSummaryBars);

            lblSummaryWeight = new WinLabel { Text = "Tổng KL thép: 0.0 kg", Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), ForeColor = Color.FromArgb(220, 38, 38), Location = new Point(280, 8), AutoSize = true };
            summaryBar.Controls.Add(lblSummaryWeight);

            lblSummaryRatio = new WinLabel { Text = "Hàm lượng: 0.0 kg/m³", Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), ForeColor = Color.FromArgb(16, 149, 193), Location = new Point(560, 8), AutoSize = true };
            summaryBar.Controls.Add(lblSummaryRatio);
        }

        // ════════════════════════════════════════════════════════════════
        // TAB 3: TÙY CHỌN NGUỒN QUÉT & XUẤT FILE
        // ════════════════════════════════════════════════════════════════
        private void BuildExportTab(TabPage page)
        {
            var pnl = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(24, 16, 24, 16),
                BackColor = BgWhite
            };
            page.Controls.Add(pnl);

            int y = 12;

            // Nhóm 1: Nguồn dữ liệu tạo BIM
            var grpSource = CreateGroupBoxPanel("⚙️  CHẾ ĐỘ TẠO MÔ HÌNH BIM", pnl, ref y);
            rdoSourceTable = new RadioButton
            {
                Text = "  Tạo mô hình 3D tự động từ BẢNG THÔNG SỐ CỐT THÉP đã nhập (Khuyên dùng - Chuẩn xác)",
                Location = new Point(16, 28),
                AutoSize = true,
                Checked = true,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                ForeColor = NavyDark
            };
            grpSource.Controls.Add(rdoSourceTable);

            rdoSourceCad = new RadioButton
            {
                Text = "  Quét trực tiếp các đối tượng (Polyline/Line) trên các layer THEP_D* từ bản vẽ CAD",
                Location = new Point(16, 62),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5f),
                ForeColor = TextDark
            };
            grpSource.Controls.Add(rdoSourceCad);
            grpSource.Height = 100;

            // Nhóm 2: Phạm vi quét nếu chọn quét CAD
            var grpScope = CreateGroupBoxPanel("🔍  PHẠM VI QUÉT BẢN VẼ", pnl, ref y);
            rdoScopeAll = new RadioButton
            {
                Text = "  Toàn bộ bản vẽ (Model Space)",
                Location = new Point(16, 28),
                AutoSize = true,
                Checked = true,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpScope.Controls.Add(rdoScopeAll);

            rdoScopeSelect = new RadioButton
            {
                Text = "  Chọn vùng đối tượng thủ công trong AutoCAD",
                Location = new Point(16, 58),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpScope.Controls.Add(rdoScopeSelect);
            grpScope.Height = 92;

            // Nhóm 3: Thư mục xuất file
            var grpFolder = CreateGroupBoxPanel("📂  THƯ MỤC LƯU FILE", pnl, ref y);
            grpFolder.Controls.Add(CreateLabel("Đường dẫn lưu file:", 16, 32));
            txtOutputFolder = CreateTextBox(160, 28, 480, OutputFolder);
            txtOutputFolder.ReadOnly = true;
            grpFolder.Controls.Add(txtOutputFolder);

            var btnBrowse = CreateButton("📁 Chọn...", 650, 26, 90, 30);
            btnBrowse.Click += (s, e) =>
            {
                using var fbd = new FolderBrowserDialog { SelectedPath = txtOutputFolder.Text };
                if (fbd.ShowDialog() == DialogResult.OK)
                    txtOutputFolder.Text = fbd.SelectedPath;
            };
            grpFolder.Controls.Add(btnBrowse);

            chkAutoOpen = CreateCheckBox("  Tự động mở file HTML trong trình duyệt sau khi xuất", 16, 70, true);
            grpFolder.Controls.Add(chkAutoOpen);
            grpFolder.Height = 108;
        }

        // ════════════════════════════════════════════════════════════════
        // HEADER & BOTTOM BAR
        // ════════════════════════════════════════════════════════════════
        private Panel CreateHeaderPanel()
        {
            var header = new Panel { Dock = DockStyle.Top, Height = 60 };
            header.Paint += (s, e) =>
            {
                using var grad = new LinearGradientBrush(header.ClientRectangle, NavyDark, NavyDeep, 45f);
                e.Graphics.FillRectangle(grad, header.ClientRectangle);
                using var linePen = new Pen(Color.FromArgb(59, 130, 246), 2);
                e.Graphics.DrawLine(linePen, 0, header.Height - 1, header.Width, header.Height - 1);
            };

            var lblTitle = new WinLabel
            {
                Text = "🏗️  RebarBIM — Bảng Nhập Thông Số Cấu Kiện & Cốt Thép 3D",
                Font = new Font("Segoe UI", 13, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(20, 8),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            header.Controls.Add(lblTitle);

            var lblSub = new WinLabel
            {
                Text = "Quản lý bảng thông số cốt thép khoa học, xem trước hình học và xuất mô hình BIM tương tác 3D",
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = Color.FromArgb(190, 210, 240),
                Location = new Point(20, 34),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            header.Controls.Add(lblSub);

            return header;
        }

        private Panel CreateBottomBarPanel()
        {
            var bar = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 54,
                BackColor = Color.FromArgb(245, 247, 250),
                BorderStyle = BorderStyle.FixedSingle
            };

            var btnCancel = CreateButton("✕  Hủy", 0, 0, 100, 36);
            btnCancel.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            btnCancel.Click += (s, e) => { IsConfirmed = false; this.Close(); };
            bar.Controls.Add(btnCancel);

            var btnExport = CreatePrimaryButton("🚀  Xuất Mô Hình BIM 3D", 0, 0, 210, 36);
            btnExport.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            btnExport.Click += BtnExport_Click;
            bar.Controls.Add(btnExport);

            bar.Layout += (s, e) =>
            {
                btnExport.Location = new Point(bar.ClientSize.Width - btnExport.Width - 16, 8);
                btnCancel.Location = new Point(btnExport.Left - btnCancel.Width - 10, 8);
            };

            return bar;
        }

        // ════════════════════════════════════════════════════════════════
        // SỰ KIỆN & THAO TÁC TRÊN BẢNG CỐT THÉP
        // ════════════════════════════════════════════════════════════════
        private void BtnAddRow_Click(object? sender, EventArgs e)
        {
            int nextId = dgvRebar.Rows.Count + 1;
            int rowIndex = dgvRebar.Rows.Add(
                $"G{nextId}",
                "Thép tăng cường",
                16,
                20,
                (int)MemberWidth - 120,
                "Straight",
                "THEP_D16",
                "0.0"
            );
            RecalcRowWeight(rowIndex);
            UpdateSummaryStats();
        }

        private void BtnDeleteRow_Click(object? sender, EventArgs e)
        {
            if (dgvRebar.SelectedRows.Count > 0)
            {
                dgvRebar.Rows.RemoveAt(dgvRebar.SelectedRows[0].Index);
                UpdateSummaryStats();
            }
        }

        private void DgvRebar_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && (e.ColumnIndex == 2 || e.ColumnIndex == 3 || e.ColumnIndex == 4))
            {
                RecalcRowWeight(e.RowIndex);
                UpdateSummaryStats();
            }
        }

        private void RecalcRowWeight(int rowIndex)
        {
            try
            {
                var row = dgvRebar.Rows[rowIndex];
                int dia = Convert.ToInt32(row.Cells[2].Value);
                int qty = Convert.ToInt32(row.Cells[3].Value);
                double len = Convert.ToDouble(row.Cells[4].Value);

                double weight = qty * len * (0.006165 * dia * dia) / 1000.0;
                row.Cells[7].Value = weight.ToString("N1");
            }
            catch { }
        }

        private void UpdateSummaryStats()
        {
            int totalBars = 0;
            double totalWeight = 0;

            foreach (DataGridViewRow row in dgvRebar.Rows)
            {
                if (row.IsNewRow) continue;
                if (int.TryParse(row.Cells[3].Value?.ToString(), out int q)) totalBars += q;
                if (double.TryParse(row.Cells[7].Value?.ToString(), out double w)) totalWeight += w;
            }

            double vol = MemberWidth * MemberHeight * MemberDepth / 1e9; // m³
            double ratio = vol > 0 ? totalWeight / vol : 0;

            if (lblSummaryBars != null) lblSummaryBars.Text = $"Tổng số thanh: {totalBars:N0} thanh";
            if (lblSummaryWeight != null) lblSummaryWeight.Text = $"Tổng KL thép: {totalWeight:N1} kg";
            if (lblSummaryRatio != null) lblSummaryRatio.Text = $"Hàm lượng: {ratio:N1} kg/m³";
        }

        // ════════════════════════════════════════════════════════════════
        // PRESETS CỐT THÉP CHUẨN
        // ════════════════════════════════════════════════════════════════
        private void LoadDefaultPreset()
        {
            LoadPresetXaMu();
        }

        /// <summary>
        /// Nạp mẫu cốt thép Xà Mũ Trụ Cầu T3 (~728 thanh, ~16.3 tấn thép) đúng ảnh mẫu
        /// </summary>
        private void LoadPresetXaMu()
        {
            dgvRebar.Rows.Clear();

            AddPresetRow("G1", "Thép chủ dọc dưới (2 lớp)", 28, 36, 11580, "Straight", "THEP_D28");
            AddPresetRow("G2", "Thép chủ dọc trên (2 lớp)", 28, 36, 11580, "Straight", "THEP_D28");
            AddPresetRow("G3", "Thép đai chính ngoài (chu vi Y-Z)", 14, 65, 7840, "Stirrup", "THEP_D14_DAI");
            AddPresetRow("G4", "Thép đai phụ trong (chia khoảng)", 12, 65, 4200, "Stirrup", "THEP_D12_DAI");
            AddPresetRow("G5", "Thép cấu tạo sườn chống nứt", 16, 12, 11580, "Straight", "THEP_D16");
            AddPresetRow("G6", "Thép giằng ngang chống phình", 12, 60, 2530, "Straight", "THEP_D12");
            AddPresetRow("G7", "Thép xiên chống cắt vát vai", 20, 16, 2200, "Straight", "THEP_D20");
            AddPresetRow("G8", "Lưới thép bệ kê gối (8 bệ gối)", 14, 80, 650, "Straight", "THEP_D14");
            AddPresetRow("G9", "Thép chờ cột / cọc neo đáy", 25, 48, 1600, "LBend", "THEP_D25");

            UpdateSummaryStats();
        }

        private void LoadPresetDam()
        {
            dgvRebar.Rows.Clear();
            AddPresetRow("G1", "Thép chủ đáy dầm", 32, 16, 11800, "Straight", "THEP_D32");
            AddPresetRow("G2", "Thép đỉnh bản cánh dầm", 20, 8, 11800, "Straight", "THEP_D20");
            AddPresetRow("G3", "Thép đai dầm chính", 12, 80, 2800, "Stirrup", "THEP_D12_DAI");
            AddPresetRow("G4", "Thép sườn bản bụng", 14, 10, 11800, "Straight", "THEP_D14");
            AddPresetRow("G5", "Thép neo đầu dầm", 16, 24, 1500, "LBend", "THEP_D16");
            UpdateSummaryStats();
        }

        private void AddPresetRow(string grp, string desc, int dia, int qty, double len, string shape, string layer)
        {
            int idx = dgvRebar.Rows.Add(grp, desc, dia, qty, (int)len, shape, layer, "0.0");
            RecalcRowWeight(idx);
        }

        // ════════════════════════════════════════════════════════════════
        // QUÉT TỪ BẢN VẼ CAD VÀO BẢNG
        // ════════════════════════════════════════════════════════════════
        private void BtnScanCadToTable_Click(object? sender, EventArgs e)
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;
                var db = doc.Database;
                var ed = doc.Editor;

                var scannedGroups = new Dictionary<string, (int dia, int count, double totalLen, string layer)>(StringComparer.OrdinalIgnoreCase);

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    var rebarLayers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    foreach (ObjectId lId in lt)
                    {
                        var ltr = (LayerTableRecord)tr.GetObject(lId, OpenMode.ForRead);
                        if (RebarUtils.ExtractDiameter(ltr.Name) > 0)
                            rebarLayers.Add(ltr.Name);
                    }

                    if (rebarLayers.Count == 0)
                    {
                        MessageBox.Show(
                            "Không tìm thấy layer cốt thép nào chứa 'D' + số (ví dụ THEP_D25, D14_DAI) trong bản vẽ!",
                            "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                    foreach (ObjectId entId in ms)
                    {
                        if (tr.GetObject(entId, OpenMode.ForRead) is Entity ent && rebarLayers.Contains(ent.Layer))
                        {
                            int dia = RebarUtils.ExtractDiameter(ent.Layer);
                            double len = 0;

                            if (ent is Polyline pl) len = pl.Length;
                            else if (ent is Line ln) len = ln.Length;

                            if (len <= 0) continue;

                            string key = $"D{dia}_{ent.Layer}";
                            if (!scannedGroups.ContainsKey(key))
                                scannedGroups[key] = (dia, 1, len, ent.Layer);
                            else
                            {
                                var cur = scannedGroups[key];
                                scannedGroups[key] = (dia, cur.count + 1, cur.totalLen + len, ent.Layer);
                            }
                        }
                    }
                    tr.Commit();
                }

                if (scannedGroups.Count > 0)
                {
                    dgvRebar.Rows.Clear();
                    int gIdx = 1;
                    foreach (var kvp in scannedGroups)
                    {
                        double avgLen = kvp.Value.totalLen / kvp.Value.count;
                        AddPresetRow($"G{gIdx}", $"Thép quét từ CAD ({kvp.Value.layer})", kvp.Value.dia, kvp.Value.count, Math.Round(avgLen, 0), "Straight", kvp.Value.layer);
                        gIdx++;
                    }
                    UpdateSummaryStats();
                    MessageBox.Show($"Đã quét thành công {scannedGroups.Count} nhóm thép từ bản vẽ CAD vào bảng!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Không tìm thấy thanh thép nào trên các layer cốt thép trong bản vẽ.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi quét CAD: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ════════════════════════════════════════════════════════════════
        // PREVIEW 3D ISOMETRIC VẼ BẰNG GDI+
        // ════════════════════════════════════════════════════════════════
        private void DrawIsometricPreview(object? sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var rect = pnlPreview.ClientRectangle;
            using var bgBrush = new SolidBrush(Color.FromArgb(250, 252, 255));
            g.FillRectangle(bgBrush, rect);

            double w = MemberWidth;
            double h = MemberHeight;
            double d = MemberDepth;
            double canti = CantileverLen;
            double cantiH = CantileverH;

            if (w <= 0 || h <= 0 || d <= 0) return;

            // Scale to fit viewport
            float margin = 40;
            float availW = rect.Width - margin * 2;
            float availH = rect.Height - margin * 2;

            double isoW = w + d * 0.4;
            double isoH = h + d * 0.4;
            float scale = (float)Math.Min(availW / isoW, availH / isoH);

            float pw = (float)(w * scale);
            float ph = (float)(h * scale);
            float pd = (float)(d * scale * 0.4);
            float pcanti = (float)(canti * scale);
            float pcantiH = (float)(cantiH * scale);

            float cx = rect.Width / 2f - pw / 2f + pd / 2f;
            float cy = rect.Height / 2f + ph / 2f - pd / 2f + 10;

            // Mặt trước có vát vai (Front Face)
            var frontPts = new PointF[]
            {
                new(cx, cy - ph + pcantiH),
                new(cx + pcanti, cy),
                new(cx + pw - pcanti, cy),
                new(cx + pw, cy - ph + pcantiH),
                new(cx + pw, cy - ph),
                new(cx, cy - ph)
            };

            using var frontBrush = new SolidBrush(Color.FromArgb(210, 225, 245));
            g.FillPolygon(frontBrush, frontPts);
            using var edgePen = new Pen(NavyDark, 1.8f);
            g.DrawPolygon(edgePen, frontPts);

            // Mặt trên (Top Face)
            var topPts = new PointF[]
            {
                new(cx, cy - ph),
                new(cx + pw, cy - ph),
                new(cx + pw + pd, cy - ph - pd),
                new(cx + pd, cy - ph - pd)
            };
            using var topBrush = new SolidBrush(Color.FromArgb(235, 243, 255));
            g.FillPolygon(topBrush, topPts);
            g.DrawPolygon(edgePen, topPts);

            // Mặt cạnh bên phải (Right Face)
            var rightPts = new PointF[]
            {
                new(cx + pw, cy - ph),
                new(cx + pw + pd, cy - ph - pd),
                new(cx + pw + pd, cy - ph - pd + pcantiH),
                new(cx + pw, cy - ph + pcantiH)
            };
            using var rightBrush = new SolidBrush(Color.FromArgb(180, 205, 235));
            g.FillPolygon(rightBrush, rightPts);
            g.DrawPolygon(edgePen, rightPts);

            // Vẽ các bệ kê gối trên mặt trên (8 bệ gối)
            if (PedestalCount > 0)
            {
                int cols = PedestalCount / 2;
                if (cols < 1) cols = 1;
                float stepX = (pw - 2 * pcanti * 0.6f) / (cols > 1 ? cols - 1 : 1);
                float pedW = Math.Max(12, pw / 20);
                float pedD = Math.Max(8, pd / 4);
                float pedH = 6;

                using var pedBrush = new SolidBrush(Color.FromArgb(100, 116, 139));
                using var pedPen = new Pen(Color.FromArgb(51, 65, 85), 1f);

                for (int r = 0; r < 2; r++)
                {
                    float zOffset = (r == 0 ? 0.3f : 0.7f) * pd;
                    for (int c = 0; c < cols; c++)
                    {
                        float px = cx + pcanti * 0.6f + c * stepX - pedW / 2 + zOffset;
                        float py = cy - ph - zOffset - pedH;

                        g.FillRectangle(pedBrush, px, py, pedW, pedH);
                        g.DrawRectangle(pedPen, px, py, pedW, pedH);
                    }
                }
            }

            // Kích thước ghi chú
            using var dimFont = new Font("Segoe UI", 8f, FontStyle.Bold);
            using var dimBrush = new SolidBrush(NavyDark);
            g.DrawString($"L = {w:N0} mm", dimFont, dimBrush, cx + pw / 2 - 35, cy + 6);
            g.DrawString($"H = {h:N0}", dimFont, dimBrush, cx - 55, cy - ph / 2);
            g.DrawString($"B = {d:N0}", dimFont, dimBrush, cx + pw + 8, cy - ph / 2);
        }

        private void UpdatePreview()
        {
            pnlPreview?.Invalidate();
            double vol = MemberWidth * MemberHeight * MemberDepth / 1e9;
            if (lblPreviewVol != null)
                lblPreviewVol.Text = $"Thể tích bê tông: V = {vol:N3} m³";
        }

        private void OnMemberTypeChanged(object? sender, EventArgs e)
        {
            switch (cboType.SelectedIndex)
            {
                case 0: // XaMu
                    numWidth.Value = 11700; numHeight.Value = 1600; numDepth.Value = 2500;
                    numCantiLen.Value = 2200; numCantiH.Value = 700; numPedestals.Value = 8;
                    LoadPresetXaMu();
                    break;
                case 1: // Dam
                    numWidth.Value = 12000; numHeight.Value = 1200; numDepth.Value = 400;
                    numCantiLen.Value = 0; numCantiH.Value = 0; numPedestals.Value = 0;
                    LoadPresetDam();
                    break;
                case 2: // Cot
                    numWidth.Value = 800; numHeight.Value = 4000; numDepth.Value = 800;
                    numCantiLen.Value = 0; numCantiH.Value = 0; numPedestals.Value = 0;
                    break;
                case 3: // BanSan
                    numWidth.Value = 6000; numHeight.Value = 200; numDepth.Value = 4000;
                    numCantiLen.Value = 0; numCantiH.Value = 0; numPedestals.Value = 0;
                    break;
                case 4: // Tuong
                    numWidth.Value = 5000; numHeight.Value = 3000; numDepth.Value = 300;
                    numCantiLen.Value = 0; numCantiH.Value = 0; numPedestals.Value = 0;
                    break;
            }
            UpdatePreview();
        }

        // ════════════════════════════════════════════════════════════════
        // NÚT XUẤT MÔ HÌNH BIM & VALIDATE
        // ════════════════════════════════════════════════════════════════
        private void BtnExport_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                MessageBox.Show("Vui lòng nhập tên cấu kiện!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tabControl.SelectedIndex = 0;
                txtName.Focus();
                return;
            }

            if (dgvRebar.Rows.Count == 0 && rdoSourceTable.Checked)
            {
                MessageBox.Show("Bảng cốt thép chưa có dữ liệu! Vui lòng thêm thanh hoặc nạp mẫu preset.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tabControl.SelectedIndex = 1;
                return;
            }

            // Thu thập dữ liệu từ Bảng
            RebarItems.Clear();
            foreach (DataGridViewRow row in dgvRebar.Rows)
            {
                if (row.IsNewRow) continue;
                string grp = row.Cells[0].Value?.ToString() ?? "G1";
                string desc = row.Cells[1].Value?.ToString() ?? "Thép";
                int dia = Convert.ToInt32(row.Cells[2].Value ?? 16);
                int qty = Convert.ToInt32(row.Cells[3].Value ?? 1);
                double len = Convert.ToDouble(row.Cells[4].Value ?? 1000);
                string shape = row.Cells[5].Value?.ToString() ?? "Straight";
                string layer = row.Cells[6].Value?.ToString() ?? $"THEP_D{dia}";

                var color = RebarUtils.GetColorForDiameter(dia);

                RebarItems.Add(new RebarTableItem
                {
                    GroupCode = grp,
                    Name = desc,
                    Diameter = dia,
                    Quantity = qty,
                    Length = len,
                    ShapeType = shape,
                    Layer = layer,
                    ColorRgb = color
                });
            }

            // Thu thập kết quả
            IsConfirmed = true;
            MemberName = txtName.Text.Trim();
            MemberType = MemberTypes[cboType.SelectedIndex].Key;
            MemberWidth = (double)numWidth.Value;
            MemberHeight = (double)numHeight.Value;
            MemberDepth = (double)numDepth.Value;
            CantileverLen = (double)numCantiLen.Value;
            CantileverH = (double)numCantiH.Value;
            PedestalCount = (int)numPedestals.Value;
            ConcreteCover = (double)numCover.Value;
            UseTableDataMode = rdoSourceTable.Checked;
            SelectArea = rdoScopeSelect.Checked;
            OutputFolder = txtOutputFolder.Text;
            AutoOpenHtml = chkAutoOpen.Checked;

            this.Close();
        }

        // ════════════════════════════════════════════════════════════════
        // HELPERS TẠO CONTROLS
        // ════════════════════════════════════════════════════════════════
        private Panel CreateGroupBoxPanel(string title, Panel parent, ref int y)
        {
            var pnl = new Panel
            {
                Location = new Point(16, y),
                Size = new Size(parent.ClientSize.Width - 32, 100),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.FromArgb(252, 253, 255),
                BorderStyle = BorderStyle.FixedSingle
            };

            var lbl = new WinLabel
            {
                Text = title,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = NavyDark,
                Location = new Point(12, 6),
                AutoSize = true
            };
            pnl.Controls.Add(lbl);

            parent.Controls.Add(pnl);
            y += pnl.Height + 12;
            return pnl;
        }

        private WinLabel CreateDimLabel(string text, int x, int y)
        {
            return new WinLabel
            {
                Text = text,
                Location = new Point(x, y + 4),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
        }

        private WinLabel CreateUnitLabel(string text, int x, int y)
        {
            return new WinLabel
            {
                Text = text,
                Location = new Point(x, y + 4),
                AutoSize = true,
                ForeColor = TextHint,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Italic)
            };
        }

        private NumericUpDown CreateNumericInput(int x, int y, int w, decimal val, decimal min = 0, decimal max = 50000)
        {
            return new NumericUpDown
            {
                Location = new Point(x, y),
                Size = new Size(w, 28),
                Minimum = min,
                Maximum = max,
                Value = Math.Max(min, Math.Min(max, val)),
                DecimalPlaces = 0,
                ThousandsSeparator = true,
                TextAlign = HorizontalAlignment.Right,
                Font = new Font("Segoe UI", 9.5f)
            };
        }

        private Button CreateToolButton(string text, int x, int y, int w, EventHandler onClick)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = NavyDark,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btn.Click += onClick;
            return btn;
        }
    }
}
