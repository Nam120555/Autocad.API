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

namespace MyFirstProject
{
    /// <summary>
    /// Form cấu hình tạo MView theo tuyến Alignment
    /// </summary>
    public class TaoMViewForm : CivilFormBase
    {
        // ─── Theme Colors ───
        private new static readonly Color GreenOk = Color.FromArgb(46, 160, 67);

        // ─── Controls ───
        private ComboBox cboAlignment = null!;
        private WinLabel lblAlignmentInfo = null!;
        private NumericUpDown numWidth = null!;
        private NumericUpDown numHeight = null!;
        private NumericUpDown numScale = null!;
        private NumericUpDown numOverlap = null!;
        private DataGridView dgvRanges = null!; // Bảng nhập nhiều phân đoạn lý trình
        private NumericUpDown numMaxVpPerLayout = null!;
        private CheckBox chkAutoSplit = null!;
        private DataGridView dgvPreview = null!;
        private WinLabel lblTotal = null!;

        // ─── Data ───
        private List<AlignmentInfo> alignments = new();
        private double currentAlignStartSta = 0;
        private double currentAlignEndSta = 0;
        public bool IsConfirmed { get; private set; } = false;
        public ObjectId SelectedAlignmentId { get; private set; } = ObjectId.Null;
        public string SelectedAlignmentName { get; private set; } = "";
        public double ViewWidth => (double)numWidth.Value;    // Chiều rộng viewport (m thực tế)
        public double ViewHeight => (double)numHeight.Value;  // Chiều cao viewport (m thực tế)
        public double ViewScale => (double)numScale.Value;    // Tỷ lệ (1 = 1:1)
        public double OverlapPercent => (double)numOverlap.Value; // % chồng lấn giữa các view
        public int MaxVpPerLayout => (int)numMaxVpPerLayout.Value;
        public List<ViewSegment> ViewSegments { get; private set; } = new();

        // ─── Pick Station Mode ───
        // Khi user click nút 📍, form sẽ đóng tạm với DialogResult.Retry
        // Command method sẽ pick point trên tuyến rồi mở lại form
        public bool IsPickMode { get; private set; } = false;
        public int PickRowIndex { get; private set; } = -1;
        public string PickTargetColumn { get; private set; } = ""; // "RangeFrom" hoặc "RangeTo"

        /// <summary>
        /// Lấy danh sách các phân đoạn lý trình từ bảng dgvRanges
        /// </summary>
        public List<(double From, double To)> GetStationRanges()
        {
            var ranges = new List<(double, double)>();
            foreach (DataGridViewRow row in dgvRanges.Rows)
            {
                if (row.IsNewRow) continue;
                if (row.Cells["RangeFrom"].Value == null || row.Cells["RangeTo"].Value == null) continue;
                if (double.TryParse(row.Cells["RangeFrom"].Value.ToString(), out double from) &&
                    double.TryParse(row.Cells["RangeTo"].Value.ToString(), out double to))
                {
                    if (from < to) ranges.Add((from, to));
                }
            }
            return ranges;
        }

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
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
        }

        private void InitializeUI()
        {
            // ── Form ──
            InitBaseTheme("TẠO MVIEW THEO TUYẾN — Civil Tool", 720, 820);
            this.FormBorderStyle = FormBorderStyle.Sizable;

            // ── HEADER ──
            var header = CreateHeader("TẠO MVIEW THEO TUYẾN ALIGNMENT");
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
                BackColor = BgPanel
            };
            this.Controls.Add(subtitle);
            subtitle.BringToFront();

            // ── BOTTOM BAR ──
            var pnlBottom = CreateBottomBar();
            this.Controls.Add(pnlBottom);

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
            var grpAlign = CreateGroupBox("Chọn Alignment", 0, y, 670, 55);
            mainPanel.Controls.Add(grpAlign);

            grpAlign.Controls.Add(CreateLabel("Tuyến:", 12, 28));
            cboAlignment = new ComboBox
            {
                Location = new Point(75, 24),
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

            lblAlignmentInfo = CreateLabel("", 410, 28);
            lblAlignmentInfo.ForeColor = TextDim;
            lblAlignmentInfo.AutoSize = true;
            grpAlign.Controls.Add(lblAlignmentInfo);

            y += 60;

            // ═══════════════════════════════════════════════════
            // SECTION 1B: PHÂN ĐOẠN LÝ TRÌNH (NHIỀU ĐOẠN)
            // ═══════════════════════════════════════════════════
            var grpRanges = CreateGroupBox("Phân đoạn lý trình (có thể thêm nhiều đoạn)", 0, y, 670, 170);
            mainPanel.Controls.Add(grpRanges);

            dgvRanges = new DataGridView
            {
                Location = new Point(10, 28),
                Size = new Size(520, 130),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
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
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            var colSTT = new DataGridViewTextBoxColumn { Name = "RangeSTT", HeaderText = "STT", Width = 45, ReadOnly = true };
            var colFrom = new DataGridViewTextBoxColumn { Name = "RangeFrom", HeaderText = "Từ Station" };
            // Nút pick cho Từ Station
            var colPickFrom = new DataGridViewButtonColumn
            {
                Name = "PickFrom",
                HeaderText = "",
                Width = 30,
                Text = "Pick",
                UseColumnTextForButtonValue = true,
                FlatStyle = FlatStyle.Flat
            };
            colPickFrom.DefaultCellStyle.BackColor = BgPanel;
            colPickFrom.DefaultCellStyle.ForeColor = TextDark;
            colPickFrom.DefaultCellStyle.Padding = new Padding(0);
            colPickFrom.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;

            var colTo = new DataGridViewTextBoxColumn { Name = "RangeTo", HeaderText = "Đến Station" };
            // Nút pick cho Đến Station
            var colPickTo = new DataGridViewButtonColumn
            {
                Name = "PickTo",
                HeaderText = "",
                Width = 30,
                Text = "Pick",
                UseColumnTextForButtonValue = true,
                FlatStyle = FlatStyle.Flat
            };
            colPickTo.DefaultCellStyle.BackColor = BgPanel;
            colPickTo.DefaultCellStyle.ForeColor = TextDark;
            colPickTo.DefaultCellStyle.Padding = new Padding(0);
            colPickTo.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;

            dgvRanges.Columns.AddRange(colSTT, colFrom, colPickFrom, colTo, colPickTo);
            dgvRanges.Columns["RangeSTT"]!.Width = 45;
            dgvRanges.Columns["PickFrom"]!.Width = 30;
            dgvRanges.Columns["PickTo"]!.Width = 30;

            // Xử lý click vào nút Pick trong DataGridView
            dgvRanges.CellContentClick += DgvRanges_CellContentClick;

            grpRanges.Controls.Add(dgvRanges);

            // Nút Thêm đoạn
            var btnAddRange = new Button
            {
                Text = "➕ Thêm",
                Location = new Point(545, 28),
                Size = new Size(110, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = BgPanel,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9f),
                Cursor = Cursors.Hand
            };
            btnAddRange.FlatAppearance.BorderColor = BorderColor;
            btnAddRange.FlatAppearance.MouseOverBackColor = NavyMedium;

            btnAddRange.Click += BtnAddRange_Click;
            grpRanges.Controls.Add(btnAddRange);

            // Nút Xóa đoạn
            var btnRemoveRange = new Button
            {
                Text = "🗑 Xóa",
                Location = new Point(545, 66),
                Size = new Size(110, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = BgPanel,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9f),
                Cursor = Cursors.Hand
            };
            btnRemoveRange.FlatAppearance.BorderColor = BorderColor;
            btnRemoveRange.FlatAppearance.MouseOverBackColor = Color.FromArgb(180, 50, 50);

            btnRemoveRange.Click += BtnRemoveRange_Click;
            grpRanges.Controls.Add(btnRemoveRange);

            // Nút Chia đều
            var btnSplitEqual = new Button
            {
                Text = "✂ Chia đều",
                Location = new Point(545, 104),
                Size = new Size(110, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = BgPanel,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9f),
                Cursor = Cursors.Hand
            };
            btnSplitEqual.FlatAppearance.BorderColor = BorderColor;
            btnSplitEqual.FlatAppearance.MouseOverBackColor = NavyMedium;

            btnSplitEqual.Click += BtnSplitEqual_Click;
            grpRanges.Controls.Add(btnSplitEqual);

            y += 175;

            // ═══════════════════════════════════════════════════
            // SECTION 2: KÍCH THƯỚC VIEWPORT
            // ═══════════════════════════════════════════════════
            var grpSize = CreateGroupBox("Kích thước Viewport (trên giấy mm)", 0, y, 670, 160);
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
                ForeColor = TextDark
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
                ForeColor = TextDark
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
                ForeColor = TextDark
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
                ForeColor = TextDark
            };
            grpSize.Controls.Add(numOverlap);

            // Auto split
            chkAutoSplit = new CheckBox
            {
                Text = "Tự động chia tuyến thành nhiều view",
                Location = new Point(380, 70),
                Size = new Size(280, 24),
                ForeColor = TextDark,
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
                BackColor = BgInput, ForeColor = TextDark
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
            var btnPreview = CreateButton("🔍 Preview", 340, 8, 120, 36, pnlBottom);
            btnPreview.Click += BtnPreview_Click;

            var btnCancel = CreateButton("Hủy", 475, 8, 100, 36, pnlBottom);
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            var btnExecute = CreatePrimaryButton("Tạo MView", 590, 8, 120, 36, pnlBottom);
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
                    currentAlignStartSta = alignment.StartingStation;
                    currentAlignEndSta = alignment.EndingStation;

                    // Thêm 1 đoạn mặc định = toàn bộ tuyến
                    dgvRanges.Rows.Clear();
                    dgvRanges.Rows.Add("1", $"{currentAlignStartSta:F2}", $"{currentAlignEndSta:F2}");
                }
                tr.Commit();
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
        }

        private void BtnAddRange_Click(object? sender, EventArgs e)
        {
            int stt = dgvRanges.Rows.Count + 1;
            // Mặc định: thêm đoạn từ cuối đoạn trước đến cuối tuyến
            double fromSta = currentAlignStartSta;
            if (dgvRanges.Rows.Count > 0)
            {
                var lastRow = dgvRanges.Rows[dgvRanges.Rows.Count - 1];
                if (lastRow.Cells["RangeTo"].Value != null &&
                    double.TryParse(lastRow.Cells["RangeTo"].Value.ToString(), out double lastTo))
                {
                    fromSta = lastTo;
                }
            }
            double toSta = Math.Min(fromSta + 5000, currentAlignEndSta); // Mặc định thêm 5km
            if (fromSta >= currentAlignEndSta)
            {
                MessageBox.Show("Đã đến cuối tuyến, không thể thêm đoạn!",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            dgvRanges.Rows.Add($"{stt}", $"{fromSta:F2}", $"{toSta:F2}");
        }

        private void BtnRemoveRange_Click(object? sender, EventArgs e)
        {
            if (dgvRanges.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow row in dgvRanges.SelectedRows)
                {
                    if (!row.IsNewRow)
                        dgvRanges.Rows.Remove(row);
                }
                // Đánh lại STT
                for (int i = 0; i < dgvRanges.Rows.Count; i++)
                    dgvRanges.Rows[i].Cells["RangeSTT"].Value = $"{i + 1}";
            }
            else
            {
                MessageBox.Show("Vui lòng chọn đoạn cần xóa!",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnSplitEqual_Click(object? sender, EventArgs e)
        {
            // Hỏi người dùng muốn chia thành bao nhiêu đoạn
            string input = Microsoft.VisualBasic.Interaction.InputBox(
                $"Chia tuyến (Sta {currentAlignStartSta:F0} → {currentAlignEndSta:F0}) thành bao nhiêu đoạn?",
                "Chia đều phân đoạn", "5");
            if (string.IsNullOrEmpty(input)) return;
            if (!int.TryParse(input, out int numParts) || numParts < 1 || numParts > 100)
            {
                MessageBox.Show("Số đoạn không hợp lệ (1-100)!", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            double totalLen = currentAlignEndSta - currentAlignStartSta;
            double partLen = totalLen / numParts;

            dgvRanges.Rows.Clear();
            for (int i = 0; i < numParts; i++)
            {
                double from = currentAlignStartSta + i * partLen;
                double to = (i == numParts - 1) ? currentAlignEndSta : from + partLen;
                dgvRanges.Rows.Add($"{i + 1}", $"{from:F2}", $"{to:F2}");
            }
        }

        // ═══════════════════════════════════════════════════
        // PICK STATION TRÊN TUYẾN
        // ═══════════════════════════════════════════════════

        /// <summary>
        /// Xử lý khi user click vào nút 📍 trong DataGridView
        /// Đóng form tạm thời với DialogResult.Retry, command method sẽ pick point
        /// </summary>
        private void DgvRanges_CellContentClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            string colName = dgvRanges.Columns[e.ColumnIndex].Name;
            if (colName != "PickFrom" && colName != "PickTo") return;

            if (SelectedAlignmentId == ObjectId.Null)
            {
                MessageBox.Show("Vui lòng chọn Alignment trước!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Lưu thông tin pick mode
            IsPickMode = true;
            PickRowIndex = e.RowIndex;
            PickTargetColumn = (colName == "PickFrom") ? "RangeFrom" : "RangeTo";

            // Đóng form tạm thời với DialogResult.Retry
            this.DialogResult = DialogResult.Retry;
            this.Close();
        }

        /// <summary>
        /// Gán giá trị station đã pick được vào ô tương ứng
        /// Được gọi từ command method sau khi pick point thành công
        /// </summary>
        public void SetPickedStation(double station)
        {
            if (PickRowIndex >= 0 && PickRowIndex < dgvRanges.Rows.Count
                && !string.IsNullOrEmpty(PickTargetColumn))
            {
                dgvRanges.Rows[PickRowIndex].Cells[PickTargetColumn].Value = $"{station:F2}";
            }
            // Reset pick mode
            IsPickMode = false;
            PickRowIndex = -1;
            PickTargetColumn = "";
        }

        /// <summary>
        /// Reset pick mode khi user hủy pick (nhấn ESC)
        /// </summary>
        public void CancelPickMode()
        {
            IsPickMode = false;
            PickRowIndex = -1;
            PickTargetColumn = "";
        }

        // ── Giới hạn tối đa segment ──
        private const int MAX_SEGMENTS = 200;
        private const int BATCH_SIZE = 10; // Số segment mỗi batch transaction

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

            // Lấy danh sách phân đoạn lý trình từ bảng
            var stationRanges = GetStationRanges();
            if (stationRanges.Count == 0)
            {
                MessageBox.Show("Vui lòng thêm ít nhất 1 phân đoạn lý trình!",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

                int stt = 1;

                // Xử lý từng phân đoạn lý trình
                foreach (var (fromSta, toSta) in stationRanges)
                {
                    double startSta = Math.Max(fromSta, alignment.StartingStation);
                    double endSta = Math.Min(toSta, alignment.EndingStation);
                    if (startSta >= endSta) continue;

                    if (!chkAutoSplit.Checked)
                    {
                        // Chỉ tạo 1 view tại giữa mỗi đoạn
                        double midSta = (startSta + endSta) / 2.0;
                        AddViewSegment(alignment, startSta, endSta, midSta, stt);
                        stt++;
                    }
                    else
                    {
                        // Chia đoạn thành nhiều view
                        double currentStart = startSta;

                        while (currentStart < endSta)
                        {
                            double currentEnd = Math.Min(currentStart + stepAlongAlignment, endSta);
                            double midSta = (currentStart + currentEnd) / 2.0;

                            AddViewSegment(alignment, currentStart, currentEnd, midSta, stt);
                            stt++;
                            currentStart += stepAlongAlignment;

                            // Giới hạn tối đa
                            if (stt > MAX_SEGMENTS)
                            {
                                MessageBox.Show(
                                    $"Đã đạt giới hạn tối đa {MAX_SEGMENTS} segments.\n\n" +
                                    $"Gợi ý:\n" +
                                    $"• Tăng kích thước viewport (Chiều rộng)\n" +
                                    $"• Giảm số phân đoạn lý trình\n" +
                                    $"• Giảm % chồng lấn",
                                    "Cảnh báo — Giới hạn segment",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                break;
                            }
                        }
                    }

                    if (stt > MAX_SEGMENTS) break;
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
            lblTotal.Text = $"Tổng: {ViewSegments.Count} view ({stationRanges.Count} phân đoạn)";
            lblTotal.ForeColor = ViewSegments.Count > 0 ? GreenOk : TextDim;

            // Cảnh báo nếu số lượng lớn
            if (ViewSegments.Count > 50)
            {
                MessageBox.Show(
                    $"Số lượng view khá lớn ({ViewSegments.Count} views).\n\n" +
                    $"Lệnh sẽ xử lý từng phân đoạn (batch {BATCH_SIZE} segments)\n" +
                    $"để tránh crash AutoCAD.\n\n" +
                    $"Quá trình có thể mất vài phút.",
                    "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
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
                $"Tỷ lệ: 1:{numScale.Value}\n" +
                $"Xử lý theo batch {BATCH_SIZE} segments/lần để tránh crash.",
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

        private new Panel CreateGroupBox(string title, int x, int y, int w, int h)
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
                ForeColor = NavyMedium,
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
                ForeColor = TextDark,
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

        private new Button CreateButton(string text, int x, int y, int w, int h, Control parent)
        {
            var btn = base.CreateButton(text, x, y, w, h, parent);
            btn.TextAlign = ContentAlignment.MiddleCenter;
            return btn;
        }
    }
}
