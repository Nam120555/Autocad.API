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
            catch { }
        }

        private void InitializeUI()
        {
            // ── Form ──
            InitBaseTheme("TẠO MVIEW THEO TUYẾN — Civil Tool", 720, 820);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

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
            catch { }
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

    #endregion

    #region Command

    /// <summary>
    /// Lệnh tạo MView theo tuyến Alignment
    /// Vẽ khung polyline trong Model Space + Tạo Viewport trong Layout
    /// Xử lý từng phân đoạn (batch) để tránh crash khi tuyến dài
    /// </summary>
    public class TaoViewTheoTuyen
    {
        private const int BATCH_SIZE = 10; // Số segment mỗi batch

        [CommandMethod("CTC_TaoMViewTheoTuyen")]
        public static void CTC_TaoMViewTheoTuyen()
        {
            var ed = A.Ed;

            // Mở form cấu hình - vòng lặp để xử lý pick station trên tuyến
            using TaoMViewForm form = new();
            while (true)
            {
                var dialogResult = AcadApp.ShowModalDialog(form);

                // Nếu user click nút 📍 pick station → form trả về Retry
                if (dialogResult == DialogResult.Retry && form.IsPickMode)
                {
                    // Pick point trên tuyến
                    string label = (form.PickTargetColumn == "RangeFrom") ? "Từ Station" : "Đến Station";

                    var ppo = new PromptPointOptions(
                        $"\nChọn điểm trên tuyến cho [{label}] (hoặc ESC để hủy): ");
                    ppo.AllowNone = true;
                    var ppr = ed.GetPoint(ppo);

                    if (ppr.Status == PromptStatus.OK)
                    {
                        Point3d pickedPoint = ppr.Value;

                        // Tính station trên alignment
                        try
                        {
                            using Transaction tr = A.Db.TransactionManager.StartTransaction();
                            Alignment? alignment = tr.GetObject(
                                form.SelectedAlignmentId, OpenMode.ForRead) as Alignment;
                            if (alignment != null)
                            {
                                double station = 0, offset = 0;
                                alignment.StationOffset(
                                    pickedPoint.X, pickedPoint.Y, ref station, ref offset);
                                tr.Commit();

                                // Cảnh báo nếu điểm xa tuyến (offset > 500m)
                                bool useValue = true;
                                if (Math.Abs(offset) > 500)
                                {
                                    var msgResult = MessageBox.Show(
                                        $"Điểm bạn chọn cách tuyến {Math.Abs(offset):F1}m.\n\n" +
                                        $"Station tính được: {station:F2}\n" +
                                        $"Offset: {offset:F2}m\n\n" +
                                        $"Bạn có muốn sử dụng giá trị này không?",
                                        "Cảnh báo — Điểm xa tuyến",
                                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                    useValue = (msgResult == DialogResult.Yes);
                                }

                                if (useValue)
                                {
                                    form.SetPickedStation(station);
                                    ed.WriteMessage(
                                        $"\n Đã chọn Station = {station:F2} (offset = {offset:F2}m)");
                                }
                                else
                                {
                                    form.CancelPickMode();
                                }
                            }
                            else
                            {
                                tr.Commit();
                                form.CancelPickMode();
                            }
                        }
                        catch
                        {
                            form.CancelPickMode();
                        }
                    }
                    else
                    {
                        // User hủy pick (ESC)
                        form.CancelPickMode();
                    }

                    // Quay lại hiển form
                    continue;
                }

                // Nếu user nhấn OK (Tạo MView)
                if (dialogResult == DialogResult.OK && form.IsConfirmed)
                    break;

                // Nếu user hủy
                return;
            }

            if (form.ViewSegments.Count == 0) return;

            var db = A.Db;

            double scale = form.ViewScale;
            double vpWidthPaper = form.ViewWidth;
            double vpHeightPaper = form.ViewHeight;
            double viewWidthReal = vpWidthPaper * scale / 1000.0;
            double viewHeightReal = vpHeightPaper * scale / 1000.0;

            int totalSegments = form.ViewSegments.Count;

            // ══════════════════════════════════════════════════════
            // PHẦN 1: Vẽ khung polyline trong Model Space
            // ─ Xử lý theo BATCH để tránh crash khi tuyến dài
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

            // ── Tạo khung polyline theo BATCH ──
            // Chia segments thành các batch nhỏ, mỗi batch 1 transaction riêng
            ed.WriteMessage($"\n Bắt đầu tạo {totalSegments} khung MView (batch {BATCH_SIZE} segments/lần)...");

            // Progress meter
            using var progress = new Autodesk.AutoCAD.Runtime.ProgressMeter();
            progress.Start($"Đang tạo {totalSegments} MView theo tuyến...");
            progress.SetLimit(totalSegments);

            for (int batchStart = 0; batchStart < totalSegments; batchStart += BATCH_SIZE)
            {
                int batchEnd = Math.Min(batchStart + BATCH_SIZE, totalSegments);
                int batchCount = batchEnd - batchStart;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        BlockTable? bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        if (bt == null) { tr.Commit(); continue; }
                        BlockTableRecord? btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        if (btr == null) { tr.Commit(); continue; }

                        for (int i = batchStart; i < batchEnd; i++)
                        {
                            var seg = form.ViewSegments[i];
                            try
                            {
                                double angle = seg.AngleRad, cx = seg.CenterX, cy = seg.CenterY;
                                double halfW = viewWidthReal / 2.0, halfH = viewHeightReal / 2.0;
                                double cosA = Math.Cos(angle), sinA = Math.Sin(angle);

                                Point2d[] corners = new Point2d[4];
                                double[] dxArr = { -halfW, halfW, halfW, -halfW };
                                double[] dyArr = { -halfH, -halfH, halfH, halfH };
                                for (int j = 0; j < 4; j++)
                                    corners[j] = new Point2d(dxArr[j] * cosA - dyArr[j] * sinA + cx,
                                                              dxArr[j] * sinA + dyArr[j] * cosA + cy);

                                var pline = new Polyline();
                                for (int j = 0; j < 4; j++)
                                    pline.AddVertexAt(j, corners[j], 0, 0, 0);
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
                                progress.MeterProgress();
                            }
                            catch { }
                        }

                        tr.Commit();
                        ed.WriteMessage($"\n   Hoàn tất batch khung {batchStart + 1}~{batchEnd}/{totalSegments}");
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        ed.WriteMessage($"\n   Lỗi batch khung {batchStart + 1}~{batchEnd}: {ex.Message}");
                    }
                }

                // Cho AutoCAD xử lý giữa các batch
                System.Windows.Forms.Application.DoEvents();
            }

            progress.Stop();
            ed.WriteMessage($"\n Đã tạo {created}/{totalSegments} khung MView trong Model Space (tỉ lệ 1:{scale:F0})");

            if (created == 0) return;

            // ══════════════════════════════════════════════════════
            // PHẦN 2: Tạo Viewport trong Layout
            // ─ Xử lý từng phân đoạn (batch) để tránh crash
            // ─ Lấy hướng từ khung CAD (polyline frame Model Space)
            // ─ Kích thước viewport = kích thước khung CAD (đồng bộ)
            // ─ Xoay view theo hướng tuyến (TwistAngle)
            // ─ Tỉ lệ 1:1 (ViewHeight = chiều cao thực)
            // ══════════════════════════════════════════════════════
            string baseLayoutName = form.SelectedAlignmentName;
            int maxVpPerLayout = form.MaxVpPerLayout;
            int totalLayouts = (int)Math.Ceiling((double)totalSegments / maxVpPerLayout);

            int segIndex = 0;
            int totalVpCreated = 0;

            ed.WriteMessage($"\n Bắt đầu tạo {totalLayouts} layout, mỗi layout tối đa {maxVpPerLayout} viewport...");

            using var progressVp = new Autodesk.AutoCAD.Runtime.ProgressMeter();
            progressVp.Start($"Đang tạo viewport trong Layout...");
            progressVp.SetLimit(totalSegments);

            for (int layoutIdx = 1; layoutIdx <= totalLayouts; layoutIdx++)
            {
                string layoutName = totalLayouts == 1
                    ? baseLayoutName
                    : $"{baseLayoutName}_{layoutIdx}";

                ObjectId layoutId = ObjectId.Null;

                // ── Tạo Layout mới ──
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        LayoutManager layoutMgr = LayoutManager.Current;
                        DBDictionary? layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                        string finalName = layoutName;
                        int suffix = 2;
                        while (layoutDict != null && layoutDict.Contains(finalName))
                        {
                            finalName = $"{layoutName}_{suffix}";
                            suffix++;
                        }
                        layoutName = finalName;

                        layoutId = layoutMgr.CreateLayout(layoutName);
                        layoutMgr.CurrentLayout = layoutName;
                        tr.Commit();
                    }
                    catch
                    {
                        ed.WriteMessage($"\n Không thể tạo Layout '{layoutName}'.");
                        continue;
                    }
                }

                if (layoutId == ObjectId.Null) continue;

                // ═══ Xử lý viewport theo batch trong mỗi layout ═══
                int vpInThisLayout = 0;
                int vpToCreate = Math.Min(maxVpPerLayout, totalSegments - segIndex);

                // Chia viewport trong layout thành các batch nhỏ
                for (int vpBatchStart = 0; vpBatchStart < vpToCreate; vpBatchStart += BATCH_SIZE)
                {
                    int vpBatchEnd = Math.Min(vpBatchStart + BATCH_SIZE, vpToCreate);
                    int vpBatchCount = vpBatchEnd - vpBatchStart;

                    // Danh sách VP đã tạo để set view ở transaction sau
                    var vpDataList = new List<(ObjectId VpId, double CenterX, double CenterY,
                        double VpH, double AngleRad, int STT, double AngleDeg, double VpW)>();

                    // ═══ TRANSACTION A: Tạo viewport trống + bật On ═══
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            Layout? layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                            if (layout == null) { tr.Commit(); continue; }

                            BlockTableRecord? paperBtr = tr.GetObject(
                                layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                            if (paperBtr == null) { tr.Commit(); continue; }

                            // Xóa viewport mặc định (VP#2+) — chỉ ở batch đầu tiên
                            if (vpBatchStart == 0)
                            {
                                List<ObjectId> toDelete = new();
                                foreach (ObjectId entId in paperBtr)
                                {
                                    var obj = tr.GetObject(entId, OpenMode.ForRead);
                                    if (obj is Viewport existingVp && existingVp.Number > 1)
                                        toDelete.Add(entId);
                                }
                                foreach (var delId in toDelete)
                                {
                                    var vpDel = tr.GetObject(delId, OpenMode.ForWrite) as Viewport;
                                    vpDel?.Erase();
                                }
                            }

                            double paperGap = 10;
                            double startX = 20, startY = 20;

                            for (int b = 0; b < vpBatchCount; b++)
                            {
                                int currentVpIdx = vpBatchStart + b;
                                var seg = form.ViewSegments[segIndex + currentVpIdx];

                                double vpW = viewWidthReal;
                                double vpH = viewHeightReal;
                                double vpX = startX + vpW / 2.0 + (vpInThisLayout + b) * (vpW + paperGap);
                                double vpY = startY + vpH / 2.0;

                                // Tạo viewport trống
                                Viewport vp = new Viewport();
                                paperBtr.AppendEntity(vp);
                                tr.AddNewlyCreatedDBObject(vp, true);

                                vp.CenterPoint = new Point3d(vpX, vpY, 0);
                                vp.Width = vpW;
                                vp.Height = vpH;
                                vp.On = true;

                                vpDataList.Add((vp.ObjectId, seg.CenterX, seg.CenterY,
                                    vpH, seg.AngleRad, seg.STT, seg.AngleDeg, vpW));

                                progressVp.MeterProgress();
                            }

                            tr.Commit();
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            ed.WriteMessage($"\n   Lỗi tạo VP batch: {ex.Message}");
                            continue;
                        }
                    }

                    // Cho AutoCAD xử lý
                    System.Windows.Forms.Application.DoEvents();

                    // ═══ TRANSACTION B: Set view cho từng VP ═══
                    if (vpDataList.Count > 0)
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            try
                            {
                                foreach (var vpData in vpDataList)
                                {
                                    Viewport? vp = tr.GetObject(vpData.VpId, OpenMode.ForWrite) as Viewport;
                                    if (vp == null) continue;

                                    vp.ViewTarget = new Point3d(vpData.CenterX, vpData.CenterY, 0);
                                    vp.ViewCenter = new Point2d(0, 0);
                                    vp.ViewHeight = vpData.VpH;
                                    vp.TwistAngle = -vpData.AngleRad;
                                    vp.Locked = true;

                                    ed.WriteMessage($"\n   VP#{vpData.STT}: Center=({vpData.CenterX:F1},{vpData.CenterY:F1}) " +
                                        $"Size={vpData.VpW:F1}x{vpData.VpH:F1} Angle={vpData.AngleDeg:F1}°");
                                }

                                tr.Commit();
                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                            {
                                ed.WriteMessage($"\n   Lỗi set view: {ex.Message}");
                            }
                        }
                    }

                    vpInThisLayout += vpBatchCount;
                    totalVpCreated += vpBatchCount;

                    // Cho AutoCAD xử lý giữa các batch
                    System.Windows.Forms.Application.DoEvents();
                }

                segIndex += vpInThisLayout;
                ed.WriteMessage($"\n Layout '{layoutName}': {vpInThisLayout} viewport hoàn tất");
            }

            progressVp.Stop();

            // Chuyển về paper space để xem layout
            db.TileMode = false;
            ed.WriteMessage($"\n Tổng: {totalVpCreated} viewport trong {totalLayouts} layout (1:1, khung theo CAD)");
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
