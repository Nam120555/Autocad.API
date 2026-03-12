// 28.TaoViewTheoTuyen.cs - Tạo MView (Viewport) theo hướng Alignment
// Tạo viewport trong Layout theo hướng tuyến alignment, tỉ lệ 1:1, kích thước tùy chỉnh
// Viewport sẽ được xoay (ViewTwist) theo góc tiếp tuyến của alignment tại trung tâm đoạn

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using MyFirstProject.Extensions;

// Alias để tránh xung đột
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using ContentAlignment = System.Drawing.ContentAlignment;
using FontStyle = System.Drawing.FontStyle;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.TaoViewTheoTuyen))]

namespace MyFirstProject
{
    #region Form

    /// <summary>
    /// Form cấu hình tạo MView theo tuyến Alignment
    /// </summary>
    public class TaoMViewForm : Form
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

        // ─── Controls ───
        private ComboBox cboAlignment = null!;
        private WinLabel lblAlignmentInfo = null!;
        private NumericUpDown numWidth = null!;
        private NumericUpDown numHeight = null!;
        private NumericUpDown numScale = null!;
        private NumericUpDown numOverlap = null!;
        private NumericUpDown numFromStation = null!;
        private NumericUpDown numToStation = null!;
        private NumericUpDown numMaxVpPerLayout = null!;
        private CheckBox chkAutoSplit = null!;
        private DataGridView dgvPreview = null!;
        private WinLabel lblTotal = null!;

        // ─── Data ───
        private List<AlignmentInfo> alignments = new();
        public bool IsConfirmed { get; private set; } = false;
        public ObjectId SelectedAlignmentId { get; private set; } = ObjectId.Null;
        public string SelectedAlignmentName { get; private set; } = "";
        public double ViewWidth => (double)numWidth.Value;    // Chiều rộng viewport (m thực tế)
        public double ViewHeight => (double)numHeight.Value;  // Chiều cao viewport (m thực tế)
        public double ViewScale => (double)numScale.Value;    // Tỷ lệ (1 = 1:1)
        public double OverlapPercent => (double)numOverlap.Value; // % chồng lấn giữa các view
        public double FromStation => (double)numFromStation.Value;
        public double ToStation => (double)numToStation.Value;
        public int MaxVpPerLayout => (int)numMaxVpPerLayout.Value;
        public List<ViewSegment> ViewSegments { get; private set; } = new();

        private class AlignmentInfo
        {
            public ObjectId Id { get; set; }
            public string Name { get; set; } = "";
            public double Length { get; set; }
        }

        public class ViewSegment
        {
            public int STT { get; set; }
            public double StartStation { get; set; }
            public double EndStation { get; set; }
            public double MidStation { get; set; }
            public double CenterX { get; set; }
            public double CenterY { get; set; }
            public double AngleRad { get; set; } // Góc xoay tiếp tuyến (radian)
            public double AngleDeg => AngleRad * 180.0 / Math.PI;
        }

        // ─── Constructor ───
        public TaoMViewForm()
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
                    alignments.Add(new AlignmentInfo
                    {
                        Id = id,
                        Name = al.Name,
                        Length = al.Length
                    });
                }
                tr.Commit();
            }
            catch { }
        }

        private void InitializeUI()
        {
            // ── Form ──
            this.Text = "TẠO MVIEW THEO TUYẾN — Civil Tool";
            this.Size = new Size(720, 650);
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
                Text = "🖼  TẠO MVIEW THEO TUYẾN ALIGNMENT",
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
                Text = "Tạo các viewport trong Model Space dọc theo hướng alignment, xoay theo tiếp tuyến tuyến.",
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
            // SECTION 1: CHỌN ALIGNMENT
            // ═══════════════════════════════════════════════════
            var grpAlign = CreateGroupBox("🛣  Chọn Alignment & Phạm vi lý trình", 0, y, 670, 95);
            mainPanel.Controls.Add(grpAlign);

            grpAlign.Controls.Add(CreateLabel("Tuyến:", 12, 32));
            cboAlignment = new ComboBox
            {
                Location = new Point(75, 28),
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

            lblAlignmentInfo = CreateLabel("", 410, 32);
            lblAlignmentInfo.ForeColor = TextDim;
            lblAlignmentInfo.AutoSize = true;
            grpAlign.Controls.Add(lblAlignmentInfo);

            // Từ lý trình
            grpAlign.Controls.Add(CreateLabel("Từ Sta:", 12, 58));
            numFromStation = new NumericUpDown
            {
                Location = new Point(75, 55),
                Size = new Size(130, 28),
                Minimum = 0, Maximum = 999999,
                Value = 0, DecimalPlaces = 2, Increment = 100,
                BackColor = BgInput, ForeColor = TextWhite
            };
            grpAlign.Controls.Add(numFromStation);

            grpAlign.Controls.Add(CreateLabel("Đến Sta:", 220, 58));
            numToStation = new NumericUpDown
            {
                Location = new Point(290, 55),
                Size = new Size(130, 28),
                Minimum = 0, Maximum = 999999,
                Value = 0, DecimalPlaces = 2, Increment = 100,
                BackColor = BgInput, ForeColor = TextWhite
            };
            grpAlign.Controls.Add(numToStation);

            y += 100;

            // ═══════════════════════════════════════════════════
            // SECTION 2: KÍCH THƯỚC VIEWPORT
            // ═══════════════════════════════════════════════════
            var grpSize = CreateGroupBox("📐  Kích thước Viewport (trên giấy mm)", 0, y, 670, 160);
            mainPanel.Controls.Add(grpSize);

            // Chiều rộng
            grpSize.Controls.Add(CreateLabel("Chiều rộng (mm):", 12, 35));
            numWidth = new NumericUpDown
            {
                Location = new Point(160, 32),
                Size = new Size(100, 28),
                Minimum = 50,
                Maximum = 2000,
                Value = 420,
                Increment = 10,
                BackColor = BgInput,
                ForeColor = TextWhite
            };
            grpSize.Controls.Add(numWidth);
            grpSize.Controls.Add(CreateDimLabel("(A3 ngang = 420, A4 = 297)", 270, 35));

            // Chiều cao
            grpSize.Controls.Add(CreateLabel("Chiều cao (mm):", 12, 70));
            numHeight = new NumericUpDown
            {
                Location = new Point(160, 67),
                Size = new Size(100, 28),
                Minimum = 50,
                Maximum = 2000,
                Value = 297,
                Increment = 10,
                BackColor = BgInput,
                ForeColor = TextWhite
            };
            grpSize.Controls.Add(numHeight);
            grpSize.Controls.Add(CreateDimLabel("(A3 ngang = 297, A4 = 210)", 270, 70));

            // Tỷ lệ
            grpSize.Controls.Add(CreateLabel("Tỷ lệ 1:", 12, 105));
            numScale = new NumericUpDown
            {
                Location = new Point(160, 102),
                Size = new Size(100, 28),
                Minimum = 1,
                Maximum = 10000,
                Value = 1000,
                Increment = 100,
                BackColor = BgInput,
                ForeColor = TextWhite
            };
            grpSize.Controls.Add(numScale);
            grpSize.Controls.Add(CreateDimLabel("(VD: 1000 → tỉ lệ 1:1000)", 270, 105));

            // Chồng lấn
            grpSize.Controls.Add(CreateLabel("Chồng lấn (%):", 380, 35));
            numOverlap = new NumericUpDown
            {
                Location = new Point(510, 32),
                Size = new Size(80, 28),
                Minimum = 0,
                Maximum = 50,
                Value = 10,
                Increment = 5,
                BackColor = BgInput,
                ForeColor = TextWhite
            };
            grpSize.Controls.Add(numOverlap);

            // Auto split
            chkAutoSplit = new CheckBox
            {
                Text = "Tự động chia tuyến thành nhiều view",
                Location = new Point(380, 70),
                Size = new Size(280, 24),
                ForeColor = TextWhite,
                Font = new Font("Segoe UI", 9f),
                Checked = true,
                FlatStyle = FlatStyle.Flat
            };
            grpSize.Controls.Add(chkAutoSplit);

            // Số MView tối đa mỗi layout
            grpSize.Controls.Add(CreateLabel("MView/Layout:", 380, 105));
            numMaxVpPerLayout = new NumericUpDown
            {
                Location = new Point(510, 102),
                Size = new Size(80, 28),
                Minimum = 1, Maximum = 50,
                Value = 10, Increment = 1,
                BackColor = BgInput, ForeColor = TextWhite
            };
            grpSize.Controls.Add(numMaxVpPerLayout);

            y += 170;

            // ═══════════════════════════════════════════════════
            // SECTION 3: PREVIEW
            // ═══════════════════════════════════════════════════
            var grpPreview = CreateGroupBox("📋  Preview danh sách View", 0, y, 670, 200);
            mainPanel.Controls.Add(grpPreview);

            dgvPreview = new DataGridView
            {
                Location = new Point(10, 25),
                Size = new Size(650, 140),
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
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(45, 48, 55),
                    ForeColor = TextWhite,
                    Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                EnableHeadersVisualStyles = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvPreview.Columns.Add("STT", "STT");
            dgvPreview.Columns.Add("Start", "Từ Station");
            dgvPreview.Columns.Add("End", "Đến Station");
            dgvPreview.Columns.Add("Angle", "Góc (°)");
            dgvPreview.Columns.Add("CenterXY", "Tâm XY");
            dgvPreview.Columns["STT"]!.Width = 45;
            dgvPreview.Columns["Angle"]!.Width = 70;
            grpPreview.Controls.Add(dgvPreview);

            lblTotal = CreateLabel("Tổng: 0 view", 10, 172);
            lblTotal.ForeColor = TextDim;
            lblTotal.AutoSize = true;
            grpPreview.Controls.Add(lblTotal);

            y += 210;

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

            var btnPreview = CreateButton("🔍 Preview", 340, 8, 120, 36, pnlBottom);
            btnPreview.Click += BtnPreview_Click;

            var btnCancel = CreateButton("✖ Hủy", 475, 8, 100, 36, pnlBottom);
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            var btnExecute = CreateButton("✅ Tạo MView", 590, 8, 120, 36, pnlBottom);
            btnExecute.BackColor = GreenOk;
            btnExecute.FlatAppearance.BorderColor = Color.FromArgb(35, 134, 54);
            btnExecute.Click += BtnExecute_Click;

            // Initialize
            if (alignments.Count > 0)
                OnAlignmentChanged();
        }

        // ═══════════════════════════════════════════════════
        // EVENT HANDLERS
        // ═══════════════════════════════════════════════════

        private void OnAlignmentChanged()
        {
            if (cboAlignment.SelectedIndex < 0 || cboAlignment.SelectedIndex >= alignments.Count) return;
            var al = alignments[cboAlignment.SelectedIndex];
            SelectedAlignmentId = al.Id;
            SelectedAlignmentName = al.Name;
            lblAlignmentInfo.Text = $"L={al.Length:F1}m";

            // Cập nhật phạm vi lý trình từ alignment
            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                Alignment? alignment = tr.GetObject(al.Id, OpenMode.ForRead) as Alignment;
                if (alignment != null)
                {
                    numFromStation.Minimum = (decimal)alignment.StartingStation;
                    numFromStation.Maximum = (decimal)alignment.EndingStation;
                    numFromStation.Value = (decimal)alignment.StartingStation;
                    numToStation.Minimum = (decimal)alignment.StartingStation;
                    numToStation.Maximum = (decimal)alignment.EndingStation;
                    numToStation.Value = (decimal)alignment.EndingStation;
                }
                tr.Commit();
            }
            catch { }
        }

        private void BtnPreview_Click(object? sender, EventArgs e)
        {
            ViewSegments.Clear();
            dgvPreview.Rows.Clear();

            if (SelectedAlignmentId == ObjectId.Null)
            {
                MessageBox.Show("Vui lòng chọn Alignment!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                double scale = (double)numScale.Value;
                double viewWidthReal = (double)numWidth.Value * scale / 1000.0;   // mm → m thực tế
                double viewHeightReal = (double)numHeight.Value * scale / 1000.0;
                double overlap = (double)numOverlap.Value / 100.0;

                // Bước nhảy dọc tuyến = chiều rộng thực tế trừ chồng lấn
                double stepAlongAlignment = viewWidthReal * (1.0 - overlap);

                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                Alignment? alignment = tr.GetObject(SelectedAlignmentId, OpenMode.ForRead) as Alignment;
                if (alignment == null) { tr.Commit(); return; }

                double totalLength = alignment.Length;
                // Dùng phạm vi lý trình từ form
                double startSta = (double)numFromStation.Value;
                double endSta = (double)numToStation.Value;
                if (startSta >= endSta)
                {
                    startSta = alignment.StartingStation;
                    endSta = alignment.EndingStation;
                }

                if (!chkAutoSplit.Checked)
                {
                    // Chỉ tạo 1 view tại giữa tuyến
                    double midSta = (startSta + endSta) / 2.0;
                    AddViewSegment(alignment, startSta, endSta, midSta, 1);
                }
                else
                {
                    // Chia tuyến thành nhiều đoạn
                    int stt = 1;
                    double currentStart = startSta;

                    while (currentStart < endSta)
                    {
                        double currentEnd = Math.Min(currentStart + stepAlongAlignment, endSta);
                        double midSta = (currentStart + currentEnd) / 2.0;

                        AddViewSegment(alignment, currentStart, currentEnd, midSta, stt);
                        stt++;
                        currentStart += stepAlongAlignment;
                    }
                }

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi khi preview: {ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Update grid
            foreach (var seg in ViewSegments)
            {
                dgvPreview.Rows.Add(
                    seg.STT,
                    $"{seg.StartStation:F2}",
                    $"{seg.EndStation:F2}",
                    $"{seg.AngleDeg:F1}",
                    $"({seg.CenterX:F1}, {seg.CenterY:F1})"
                );
            }
            lblTotal.Text = $"Tổng: {ViewSegments.Count} view";
            lblTotal.ForeColor = ViewSegments.Count > 0 ? GreenOk : TextDim;
        }

        private void AddViewSegment(Alignment alignment, double startSta, double endSta, double midSta, int stt)
        {
            try
            {
                double easting = 0, northing = 0;
                alignment.PointLocation(midSta, 0, ref easting, ref northing);

                // Tính góc tiếp tuyến tại điểm mid bằng 2 điểm gần nhau
                double delta = 0.5; // 0.5m offset
                double e1 = 0, n1 = 0, e2 = 0, n2 = 0;
                double sta1 = Math.Max(startSta, midSta - delta);
                double sta2 = Math.Min(endSta, midSta + delta);
                alignment.PointLocation(sta1, 0, ref e1, ref n1);
                alignment.PointLocation(sta2, 0, ref e2, ref n2);

                double angle = Math.Atan2(n2 - n1, e2 - e1);

                ViewSegments.Add(new ViewSegment
                {
                    STT = stt,
                    StartStation = startSta,
                    EndStation = endSta,
                    MidStation = midSta,
                    CenterX = easting,
                    CenterY = northing,
                    AngleRad = angle
                });
            }
            catch { }
        }

        private void BtnExecute_Click(object? sender, EventArgs e)
        {
            if (ViewSegments.Count == 0)
            {
                BtnPreview_Click(null, EventArgs.Empty);

                if (ViewSegments.Count == 0)
                {
                    MessageBox.Show("Không có view nào để tạo! Vui lòng Preview trước.",
                        "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            var confirm = MessageBox.Show(
                $"Bạn có muốn tạo {ViewSegments.Count} MView theo tuyến không?\n\n" +
                $"Kích thước: {numWidth.Value}x{numHeight.Value} mm\n" +
                $"Tỷ lệ: 1:{numScale.Value}",
                "Xác nhận",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm == DialogResult.Yes)
            {
                IsConfirmed = true;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
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

        private WinLabel CreateDimLabel(string text, int x, int y)
        {
            return new WinLabel
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                ForeColor = TextDim,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Italic)
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
                Font = new Font("Segoe UI", 9f),
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter
            };
            btn.FlatAppearance.BorderColor = BorderColor;
            btn.FlatAppearance.MouseOverBackColor = AccentBlue;
            parent.Controls.Add(btn);
            return btn;
        }
    }

    #endregion

    #region Command

    /// <summary>
    /// Lệnh tạo MView theo tuyến Alignment
    /// Vẽ khung polyline trong Model Space + Tạo Viewport trong Layout
    /// </summary>
    public class TaoViewTheoTuyen
    {
        [CommandMethod("CTC_TaoMViewTheoTuyen")]
        public static void CTC_TaoMViewTheoTuyen()
        {
            // Mở form cấu hình
            TaoMViewForm form = new();
            var dialogResult = AcadApp.ShowModalDialog(form);

            if (dialogResult != DialogResult.OK || !form.IsConfirmed) return;
            if (form.ViewSegments.Count == 0) return;

            var ed = A.Ed;
            var db = A.Db;

            double scale = form.ViewScale;
            double vpWidthPaper = form.ViewWidth;
            double vpHeightPaper = form.ViewHeight;
            double viewWidthReal = vpWidthPaper * scale / 1000.0;
            double viewHeightReal = vpHeightPaper * scale / 1000.0;

            // ══════════════════════════════════════════════════════
            // PHẦN 1: Vẽ khung polyline trong Model Space
            // ══════════════════════════════════════════════════════
            int created = 0;
            string layerName = "CTC_MVIEW_TUYEN";
            CreateLayerIfNotExist(layerName, 3);

            // ── Xóa khung cũ trên layer CTC_MVIEW_TUYEN ──
            using (Transaction trDel = db.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTable? btDel = trDel.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    if (btDel != null)
                    {
                        BlockTableRecord? btrDel = trDel.GetObject(btDel[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        if (btrDel != null)
                        {
                            int deleted = 0;
                            foreach (ObjectId entId in btrDel)
                            {
                                var ent = trDel.GetObject(entId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                                if (ent == null) continue;
                                if (!string.Equals(ent.Layer, layerName, StringComparison.OrdinalIgnoreCase)) continue;

                                // Chỉ xóa Polyline và DBText trên layer này
                                if (ent is Polyline || ent is DBText)
                                {
                                    ent.UpgradeOpen();
                                    ent.Erase();
                                    deleted++;
                                }
                            }
                            if (deleted > 0)
                                ed.WriteMessage($"\n🗑 Đã xóa {deleted} đối tượng khung cũ trên layer '{layerName}'");
                        }
                    }
                    trDel.Commit();
                }
                catch { }
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTable? bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    if (bt == null) { tr.Commit(); return; }
                    BlockTableRecord? btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    if (btr == null) { tr.Commit(); return; }

                    foreach (var seg in form.ViewSegments)
                    {
                        try
                        {
                            double angle = seg.AngleRad, cx = seg.CenterX, cy = seg.CenterY;
                            double halfW = viewWidthReal / 2.0, halfH = viewHeightReal / 2.0;
                            double cosA = Math.Cos(angle), sinA = Math.Sin(angle);

                            Point2d[] corners = new Point2d[4];
                            double[] dxArr = { -halfW, halfW, halfW, -halfW };
                            double[] dyArr = { -halfH, -halfH, halfH, halfH };
                            for (int i = 0; i < 4; i++)
                                corners[i] = new Point2d(dxArr[i] * cosA - dyArr[i] * sinA + cx,
                                                          dxArr[i] * sinA + dyArr[i] * cosA + cy);

                            var pline = new Polyline();
                            for (int i = 0; i < 4; i++)
                                pline.AddVertexAt(i, corners[i], 0, 0, 0);
                            pline.Closed = true;
                            pline.Layer = layerName;
                            pline.ColorIndex = 3;
                            btr!.AppendEntity(pline);
                            tr.AddNewlyCreatedDBObject(pline, true);

                            double textH = viewHeightReal * 0.03;
                            var text = new DBText
                            {
                                TextString = $"VIEW {seg.STT}",
                                Position = new Point3d(cx, cy, 0),
                                Height = textH, Rotation = angle,
                                HorizontalMode = TextHorizontalMode.TextCenter,
                                VerticalMode = TextVerticalMode.TextVerticalMid,
                                AlignmentPoint = new Point3d(cx, cy, 0),
                                Layer = layerName, ColorIndex = 3
                            };
                            btr.AppendEntity(text);
                            tr.AddNewlyCreatedDBObject(text, true);

                            double lyX = cx - halfW * cosA + halfH * sinA;
                            double lyY = cy - halfW * sinA - halfH * cosA;
                            var lyText = new DBText
                            {
                                TextString = $"Sta {seg.StartStation:F0} - {seg.EndStation:F0}",
                                Position = new Point3d(lyX, lyY, 0),
                                Height = textH * 0.7, Rotation = angle,
                                Layer = layerName, ColorIndex = 3
                            };
                            btr.AppendEntity(lyText);
                            tr.AddNewlyCreatedDBObject(lyText, true);

                            created++;
                        }
                        catch { }
                    }

                    tr.Commit();
                    ed.WriteMessage($"\n✅ Đã tạo {created} khung MView trong Model Space (tỉ lệ 1:{scale:F0})");
                }
                catch (Autodesk.AutoCAD.Runtime.Exception e)
                {
                    ed.WriteMessage($"\n❌ Lỗi tạo khung: {e.Message}");
                    return;
                }
            }

            if (created == 0) return;

            // ══════════════════════════════════════════════════════
            // PHẦN 2: Tạo Viewport trong Layout
            // Phương pháp: Tạo MView trống → Set UCS theo hướng tuyến → Tỉ lệ 1:1
            // Kích thước viewport = kích thước khung trong Model Space (đồng bộ)
            // ══════════════════════════════════════════════════════
            string baseLayoutName = form.SelectedAlignmentName;
            int maxVpPerLayout = form.MaxVpPerLayout;
            int totalSegments = form.ViewSegments.Count;
            int totalLayouts = (int)Math.Ceiling((double)totalSegments / maxVpPerLayout);

            db.TileMode = false; // Chuyển Paper Space

            int segIndex = 0;
            int totalVpCreated = 0;

            for (int layoutIdx = 1; layoutIdx <= totalLayouts; layoutIdx++)
            {
                // Tên layout: "TenTuyen_1", "TenTuyen_2", ...
                string layoutName = totalLayouts == 1
                    ? baseLayoutName
                    : $"{baseLayoutName}_{layoutIdx}";

                // Tạo layout mới (nếu trùng tên thì thêm hậu tố)
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        LayoutManager layoutMgr = LayoutManager.Current;
                        DBDictionary? layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                        // Đảm bảo tên không trùng
                        string finalName = layoutName;
                        int suffix = 2;
                        while (layoutDict != null && layoutDict.Contains(finalName))
                        {
                            finalName = $"{layoutName}_{suffix}";
                            suffix++;
                        }
                        layoutName = finalName;

                        layoutMgr.CreateLayout(layoutName);
                        layoutMgr.CurrentLayout = layoutName;
                        tr.Commit();
                    }
                    catch
                    {
                        ed.WriteMessage($"\n⚠️ Không thể tạo Layout '{layoutName}'.");
                        continue;
                    }
                }

                // Tạo viewport trong layout này (tối đa maxVpPerLayout)
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        BlockTableRecord? paperBtr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        if (paperBtr == null) { tr.Commit(); continue; }

                        int vpInLayout = 0;
                        double paperGap = 10;
                        double startX = 20, startY = 20;

                        while (segIndex < totalSegments && vpInLayout < maxVpPerLayout)
                        {
                            var seg = form.ViewSegments[segIndex];

                            // Kích thước viewport = kích thước khung Model Space (đồng bộ)
                            double vpW = viewWidthReal;
                            double vpH = viewHeightReal;

                            double vpX = startX + vpW / 2.0 + vpInLayout * (vpW + paperGap);
                            double vpY = startY + vpH / 2.0;

                            // ── Bước 1: Tạo Viewport trống rồi thêm vào database TRƯỚC ──
                            Viewport vp = new Viewport();
                            paperBtr.AppendEntity(vp);
                            tr.AddNewlyCreatedDBObject(vp, true);

                            // ── Bước 2: Set kích thước và vị trí trên giấy (SAU khi add) ──
                            vp.CenterPoint = new Point3d(vpX, vpY, 0);
                            vp.Width = vpW;
                            vp.Height = vpH;
                            vp.On = true;

                            // ── Bước 3: Thiết lập camera nhìn vào Model Space ──
                            // ViewTarget = gốc tọa độ WCS (điểm camera nhắm vào)
                            vp.ViewTarget = Point3d.Origin;
                            // ViewDirection = nhìn từ trên xuống (plan view)
                            vp.ViewDirection = new Vector3d(0, 0, 1);

                            // ── Bước 4: Đặt tâm view = tâm đoạn tuyến (WCS) ──
                            vp.ViewCenter = new Point2d(seg.CenterX, seg.CenterY);

                            // ── Bước 5: Zoom 1:1 bằng ViewHeight ──
                            // ViewHeight = chiều cao vùng hiển thị trong model space
                            // Khi ViewHeight = viewport height → tỉ lệ 1:1
                            vp.ViewHeight = vpH;

                            // ── Bước 6: Xoay view theo hướng tiếp tuyến alignment ──
                            vp.TwistAngle = -seg.AngleRad;

                            // ── Bước 7: Gắn UCS (tách riêng để không gây lỗi viewport) ──
                            try
                            {
                                Point3d ucsOrigin = new Point3d(seg.CenterX, seg.CenterY, 0);
                                Vector3d ucsXAxis = new Vector3d(Math.Cos(seg.AngleRad), Math.Sin(seg.AngleRad), 0);
                                Vector3d ucsYAxis = new Vector3d(-Math.Sin(seg.AngleRad), Math.Cos(seg.AngleRad), 0);

                                string ucsName = $"CTC_VP_{form.SelectedAlignmentName}_{seg.STT}";
                                ObjectId ucsId = GetOrCreateUcs(tr, db, ucsName, ucsOrigin, ucsXAxis, ucsYAxis);
                                if (ucsId != ObjectId.Null)
                                {
                                    vp.SetUcs(ucsId);
                                    vp.UcsPerViewport = true;
                                }
                            }
                            catch { /* Bỏ qua lỗi UCS, viewport vẫn hoạt động */ }

                            // ── Bước 8: Khóa viewport ──
                            vp.Locked = true;

                            vpInLayout++;
                            segIndex++;
                            totalVpCreated++;
                        }

                        tr.Commit();
                        ed.WriteMessage($"\n✅ Layout '{layoutName}': {vpInLayout} viewport (1:1, UCS theo tuyến)");
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception e)
                    {
                        ed.WriteMessage($"\n❌ Lỗi tạo viewport: {e.Message}");
                    }
                }
            }

            ed.WriteMessage($"\n✅ Tổng cộng: {totalVpCreated} viewport trong {totalLayouts} layout (tỉ lệ 1:1, đồng bộ khung Model)");
        }

        private static void CreateLayerIfNotExist(string name, short color)
        {
            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                LayerTable? lt = tr.GetObject(A.Db.LayerTableId, OpenMode.ForWrite) as LayerTable;
                if (lt != null && !lt.Has(name))
                {
                    LayerTableRecord ltr = new()
                    {
                        Name = name,
                        Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                            Autodesk.AutoCAD.Colors.ColorMethod.ByAci, color)
                    };
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }
                tr.Commit();
            }
            catch { }
        }

        /// <summary>
        /// Tạo hoặc lấy UCS đã có trong UcsTable.
        /// Dùng để gán UCS riêng cho từng viewport.
        /// </summary>
        private static ObjectId GetOrCreateUcs(Transaction tr, Database db,
            string ucsName, Point3d origin, Vector3d xAxis, Vector3d yAxis)
        {
            try
            {
                UcsTable? ut = tr.GetObject(db.UcsTableId, OpenMode.ForWrite) as UcsTable;
                if (ut == null) return ObjectId.Null;

                // Nếu đã tồn tại → cập nhật lại origin/axis
                if (ut.Has(ucsName))
                {
                    UcsTableRecord? existing = tr.GetObject(ut[ucsName], OpenMode.ForWrite) as UcsTableRecord;
                    if (existing != null)
                    {
                        existing.Origin = origin;
                        existing.XAxis = xAxis;
                        existing.YAxis = yAxis;
                    }
                    return ut[ucsName];
                }

                // Tạo mới
                UcsTableRecord utr = new()
                {
                    Name = ucsName,
                    Origin = origin,
                    XAxis = xAxis,
                    YAxis = yAxis
                };
                ut.Add(utr);
                tr.AddNewlyCreatedDBObject(utr, true);
                return utr.ObjectId;
            }
            catch
            {
                return ObjectId.Null;
            }
        }
    }

    #endregion
}
