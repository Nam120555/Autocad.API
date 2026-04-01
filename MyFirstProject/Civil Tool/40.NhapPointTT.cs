// 40.NhapPointTT.cs — Lệnh CTPO_NhapToaDoTrucTiep
// Nhập tọa độ X, Y, Z, Tên điểm, Ghi chú, Số thứ tự trực tiếp vào Civil 3D
// KHÔNG cần nhập từ Excel — gõ tay trên bảng → tạo COGO Point ngay lập tức

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using MyFirstProject.Extensions;

using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using ContentAlignment = System.Drawing.ContentAlignment;

#pragma warning disable CS8600, CS8602, CS8604, CS8618

[assembly: CommandClass(typeof(MyFirstProject.NhapPointTrucTiepCommands))]

namespace MyFirstProject
{
    // ══════════════════════════════════════════════════════════════
    // DATA CLASS
    // ══════════════════════════════════════════════════════════════

    public class DirectPointData
    {
        public int STT { get; set; }
        public string PointName { get; set; } = "";
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public string Description { get; set; } = "";

        public bool IsValid => !string.IsNullOrWhiteSpace(PointName)
                            && (X != 0 || Y != 0);
    }

    // ══════════════════════════════════════════════════════════════
    // FORM: NHẬP TỌA ĐỘ TRỰC TIẾP
    // ══════════════════════════════════════════════════════════════

    public class FormNhapPointTrucTiep : CivilFormBase
    {
        // ═══ Controls ═══
        private DataGridView dgvPoints = null!;
        private ComboBox cboPointGroup = null!;
        private ComboBox cboPointStyle = null!;
        private ComboBox cboLabelStyle = null!;
        private TextBox txtLayer = null!;
        private NumericUpDown nudStartSTT = null!;
        private Button btnPaste = null!;
        private Button btnClear = null!;
        private Button btnAddRow = null!;
        private Button btnRemoveRow = null!;
        private WinLabel lblStatus = null!;

        // ═══ Dữ liệu ═══
        public List<DirectPointData> PointDataList { get; private set; } = new();
        public string SelectedPointGroup { get; private set; } = "";
        public string SelectedPointStyle { get; private set; } = "";
        public string SelectedLabelStyle { get; private set; } = "";
        public string LayerName { get; private set; } = "COGO_POINTS";
        public bool IsConfirmed { get; private set; }

        public FormNhapPointTrucTiep()
        {
            InitializeComponent();
            LoadCivilData();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("NHẬP TỌA ĐỘ TRỰC TIẾP VÀO CIVIL 3D", 820, 720);

            // ═══ HEADER ═══
            var header = CreateHeader("NHẬP TỌA ĐỘ TRỰC TIẾP");
            this.Controls.Add(header);

            var subtitle = CreateSubtitle("Gõ tọa độ X, Y, Z trực tiếp trên bảng -> Tạo COGO Point ngay lập tức");
            this.Controls.Add(subtitle);

            // ═══ Main Panel ═══
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(16, 8, 16, 8)
            };
            this.Controls.Add(mainPanel);
            mainPanel.BringToFront();

            int y = 8;

            // ═══ 1. CẤU HÌNH ═══
            var grpConfig = CreateGroupBox("Cấu hình Point", 0, y, 770, 90);
            mainPanel.Controls.Add(grpConfig);

            // Hàng 1: Point Group + Layer
            grpConfig.Controls.Add(CreateLabel("Point Group:", 12, 28));
            cboPointGroup = CreateComboBox(120, 25, 250);
            grpConfig.Controls.Add(cboPointGroup);

            grpConfig.Controls.Add(CreateLabel("Layer:", 390, 28));
            txtLayer = CreateTextBox(440, 25, 160, "COGO_POINTS");
            grpConfig.Controls.Add(txtLayer);

            grpConfig.Controls.Add(CreateLabel("STT bắt đầu:", 620, 28));
            nudStartSTT = new NumericUpDown
            {
                Location = new Point(710, 25),
                Size = new Size(50, 26),
                Minimum = 1,
                Maximum = 99999,
                Value = 1,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpConfig.Controls.Add(nudStartSTT);

            // Hàng 2: Point Style + Label Style
            grpConfig.Controls.Add(CreateLabel("Point Style:", 12, 60));
            cboPointStyle = CreateComboBox(120, 57, 240);
            grpConfig.Controls.Add(cboPointStyle);

            grpConfig.Controls.Add(CreateLabel("Label Style:", 375, 60));
            cboLabelStyle = CreateComboBox(450, 57, 240);
            grpConfig.Controls.Add(cboLabelStyle);

            var btnPickSample = new Button
            {
                Text = "Lấy mẫu điểm có sẵn",
                Location = new Point(700, 57),
                Size = new Size(180, 26),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = NavyDark,
                Font = new Font("Segoe UI", 9f),
                Cursor = Cursors.Hand
            };
            btnPickSample.FlatAppearance.BorderColor = BorderLight;
            btnPickSample.FlatAppearance.MouseOverBackColor = HoverNavy;
            btnPickSample.Click += BtnPickSample_Click;
            grpConfig.Controls.Add(btnPickSample);

            // Tăng chiều rộng group box
            grpConfig.Size = new Size(890, 95);
            
            y += 98;

            // ═══ 2. TOOLBAR ═══
            var grpToolbar = new Panel
            {
                Location = new Point(0, y),
                Size = new Size(770, 40),
                BackColor = BgPanel
            };
            mainPanel.Controls.Add(grpToolbar);

            btnAddRow = CreateToolButton("+ Thêm dòng", 5, 5, 110, 30, grpToolbar);
            btnAddRow.Click += (s, e) => AddEmptyRows(5);

            btnRemoveRow = CreateToolButton("- Xóa dòng", 120, 5, 110, 30, grpToolbar);
            btnRemoveRow.Click += BtnRemoveRow_Click;

            btnPaste = CreateToolButton("Dán từ Clipboard", 240, 5, 130, 30, grpToolbar);
            btnPaste.Click += BtnPaste_Click;

            btnClear = CreateToolButton("Xóa tất cả", 380, 5, 100, 30, grpToolbar);
            btnClear.Click += (s, e) => { dgvPoints.Rows.Clear(); AddEmptyRows(10); UpdateStatus(); };

            var btnImportCSV = CreateToolButton("Nhập CSV", 490, 5, 100, 30, grpToolbar);
            btnImportCSV.Click += BtnImportCSV_Click;

            var btnExportCSV = CreateToolButton("Xuất CSV", 600, 5, 100, 30, grpToolbar);
            btnExportCSV.Click += BtnExportCSV_Click;

            y += 44;

            // ═══ 3. BẢNG NHẬP DỮ LIỆU ═══
            var grpTable = CreateGroupBox("Bảng nhập tọa độ (gõ trực tiếp hoặc dán từ Excel)", 0, y, 770, 360);
            mainPanel.Controls.Add(grpTable);

            dgvPoints = new DataGridView
            {
                Location = new Point(8, 26),
                Size = new Size(750, 324),
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                AllowUserToResizeRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
                MultiSelect = true,
                BackgroundColor = Color.White,
                GridColor = BorderLight,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersWidth = 35,
                Font = new Font("Segoe UI", 9.5f),
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.White,
                    ForeColor = TextDark,
                    SelectionBackColor = HoverNavy,
                    SelectionForeColor = NavyDark,
                    Font = new Font("Segoe UI", 9.5f),
                    Padding = new Padding(2)
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = NavyDark,
                    ForeColor = Color.White,
                    Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter,
                    Padding = new Padding(2)
                },
                ColumnHeadersHeight = 32,
                EnableHeadersVisualStyles = false
            };

            // ── Các cột ──
            var colSTT = new DataGridViewTextBoxColumn
            {
                Name = "colSTT", HeaderText = "STT", Width = 55,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
            };

            var colName = new DataGridViewTextBoxColumn
            {
                Name = "colName", HeaderText = "Tên điểm", Width = 120
            };

            var colX = new DataGridViewTextBoxColumn
            {
                Name = "colX", HeaderText = "X (Easting)", Width = 120,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "N3" }
            };

            var colY = new DataGridViewTextBoxColumn
            {
                Name = "colY", HeaderText = "Y (Northing)", Width = 120,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "N3" }
            };

            var colZ = new DataGridViewTextBoxColumn
            {
                Name = "colZ", HeaderText = "Z (Cao độ)", Width = 100,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "N3" }
            };

            var colDesc = new DataGridViewTextBoxColumn
            {
                Name = "colDesc", HeaderText = "Ghi chú", Width = 200
            };

            dgvPoints.Columns.AddRange(colSTT, colName, colX, colY, colZ, colDesc);

            // Tự gán STT khi gõ
            dgvPoints.CellEndEdit += DgvPoints_CellEndEdit;
            dgvPoints.RowsAdded += (s, e) => UpdateStatus();
            dgvPoints.RowsRemoved += (s, e) => UpdateStatus();

            grpTable.Controls.Add(dgvPoints);

            y += 368;

            // ═══ STATUS BAR ═══
            lblStatus = new WinLabel
            {
                Location = new Point(0, y),
                Size = new Size(770, 24),
                Font = new Font("Segoe UI", 9f),
                ForeColor = TextHint,
                Text = "Sẵn sàng - Nhập tọa độ vào bảng hoặc dán từ clipboard",
                BackColor = BgPanel,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(12, 0, 0, 0)
            };
            mainPanel.Controls.Add(lblStatus);

            // ═══ BOTTOM BAR ═══
            var bottomBar = CreateBottomBar(56);
            this.Controls.Add(bottomBar);

            var btnCancel = CreateButton("Hủy", 10, 10, 180, 36, bottomBar);
            btnCancel.ForeColor = Color.Red;
            btnCancel.FlatAppearance.BorderColor = Color.Red;
            btnCancel.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 235, 235);
            btnCancel.Click += (s, e) => { IsConfirmed = false; this.Close(); };

            var lblInfo = new WinLabel
            {
                Location = new Point(200, 10),
                Size = new Size(320, 36),
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = TextHint,
                Text = "Tip: Copy bảng từ Excel -> Ctrl+V hoặc nút Dán",
                TextAlign = ContentAlignment.MiddleCenter
            };
            bottomBar.Controls.Add(lblInfo);

            var btnRun = CreateSuccessButton("Tạo COGO Points", 640, 10, 240, 36, bottomBar);
            btnRun.Click += BtnRun_Click;

            this.AcceptButton = btnRun;
            this.CancelButton = btnCancel;

            // Form wider to accommodate new button
            this.Size = new Size(930, 720);

            // Pre-populate 10 dòng trống
            AddEmptyRows(10);
        }

        // ═══════════════════════════════════════════════════════════
        // LOGIC
        // ═══════════════════════════════════════════════════════════

        private void LoadCivilData()
        {
            try
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                if (doc == null) return;
                var civilDoc = CivilApplication.ActiveDocument;

                // Lưu lại selection cũ nếu có
                string currentGroup = cboPointGroup.Text;
                string currentPointStyle = cboPointStyle.Text;
                string currentLabelStyle = cboLabelStyle.Text;

                using var tr = doc.Database.TransactionManager.StartOpenCloseTransaction();

                // Load Point Groups
                cboPointGroup.Items.Clear();
                cboPointGroup.Items.Add("(Tạo mới theo Layer)");
                foreach (ObjectId pgId in civilDoc.PointGroups)
                {
                    var pg = tr.GetObject(pgId, OpenMode.ForRead) as PointGroup;
                    if (pg != null)
                        cboPointGroup.Items.Add(pg.Name);
                }
                cboPointGroup.SelectedIndex = 0;
                if (!string.IsNullOrEmpty(currentGroup) && cboPointGroup.Items.Contains(currentGroup))
                    cboPointGroup.SelectedItem = currentGroup;

                // Load Point Styles
                cboPointStyle.Items.Clear();
                cboPointStyle.Items.Add("(Mặc định)");
                foreach (ObjectId psId in civilDoc.Styles.PointStyles)
                {
                    try {
                        var ps = tr.GetObject(psId, OpenMode.ForRead) as PointStyle;
                        if (ps != null) cboPointStyle.Items.Add(ps.Name);
                    } catch { } // Bỏ qua folder hoặc item lỗi
                }
                cboPointStyle.SelectedIndex = 0;
                if (!string.IsNullOrEmpty(currentPointStyle) && cboPointStyle.Items.Contains(currentPointStyle))
                    cboPointStyle.SelectedItem = currentPointStyle;

                // Load Label Styles
                cboLabelStyle.Items.Clear();
                cboLabelStyle.Items.Add("(Mặc định)");
                
                // Tránh lỗi truy cập LabelStyles collection
                try {
                    foreach (ObjectId lsId in civilDoc.Styles.LabelStyles.PointLabelStyles.LabelStyles)
                    {
                        var ls = tr.GetObject(lsId, OpenMode.ForRead) as LabelStyle;
                        if (ls != null) cboLabelStyle.Items.Add(ls.Name);
                    }
                } catch { }

                cboLabelStyle.SelectedIndex = 0;
                if (!string.IsNullOrEmpty(currentLabelStyle) && cboLabelStyle.Items.Contains(currentLabelStyle))
                    cboLabelStyle.SelectedItem = currentLabelStyle;

                tr.Commit();
            }
            catch { }
        }

        private void BtnPickSample_Click(object? sender, EventArgs e)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;

            // Start user interaction hides the modal dialog safely
            using (var interaction = ed.StartUserInteraction(this))
            {
                var peo = new PromptEntityOptions("\nChọn COGO Point mẫu để sao chép Style & Layer: ");
                peo.SetRejectMessage("\nĐối tượng bạn chọn không phải là COGO Point. Vui lòng chọn lại.");
                peo.AddAllowedClass(typeof(CogoPoint), true);
                
                var res = ed.GetEntity(peo);
                if (res.Status == PromptStatus.OK)
                {
                    try
                    {
                        using var tr = doc.Database.TransactionManager.StartOpenCloseTransaction();
                        var point = tr.GetObject(res.ObjectId, OpenMode.ForRead) as CogoPoint;
                        if (point != null)
                        {
                            // 1. Sao chép Layer
                            txtLayer.Text = point.Layer;

                            // Đảm bảo dữ liệu các ComboBox đã load đầy đủ trước khi mapping
                            if (cboPointStyle.Items.Count <= 1) LoadCivilData();

                            // 2. Sao chép Point Style Name
                            if (!point.StyleId.IsNull)
                            {
                                var ps = tr.GetObject(point.StyleId, OpenMode.ForRead) as PointStyle;
                                if (ps != null && cboPointStyle.Items.Contains(ps.Name))
                                {
                                    cboPointStyle.SelectedItem = ps.Name;
                                }
                            }
                            else
                            {
                                cboPointStyle.SelectedIndex = 0; // Về default
                            }

                            // 3. Sao chép Label Style Name
                            if (!point.LabelStyleId.IsNull)
                            {
                                var ls = tr.GetObject(point.LabelStyleId, OpenMode.ForRead) as LabelStyle;
                                if (ls != null && cboLabelStyle.Items.Contains(ls.Name))
                                {
                                    cboLabelStyle.SelectedItem = ls.Name;
                                }
                            }
                            else
                            {
                                cboLabelStyle.SelectedIndex = 0; // Về default
                            }

                            MessageBox.Show($"Đã sao chép thành công mẫu từ điểm {point.PointNumber}!",
                                "Thông minh hóa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        tr.Commit();
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show($"Lỗi khi lấy mẫu: {ex.Message}", "Lỗi API", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }


        private void AddEmptyRows(int count)
        {
            int startSTT = (int)nudStartSTT.Value + dgvPoints.Rows.Count
                - (dgvPoints.AllowUserToAddRows ? 1 : 0);

            for (int i = 0; i < count; i++)
            {
                int idx = dgvPoints.Rows.Add();
                dgvPoints.Rows[idx].Cells["colSTT"].Value = startSTT + i;
            }

            UpdateStatus();
        }

        private void DgvPoints_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
        {
            // Auto-renumber STT
            ReNumberSTT();
            UpdateStatus();
        }

        private void ReNumberSTT()
        {
            int stt = (int)nudStartSTT.Value;
            for (int i = 0; i < dgvPoints.Rows.Count; i++)
            {
                if (dgvPoints.Rows[i].IsNewRow) continue;
                dgvPoints.Rows[i].Cells["colSTT"].Value = stt++;
            }
        }

        private void UpdateStatus()
        {
            int validCount = 0;
            for (int i = 0; i < dgvPoints.Rows.Count; i++)
            {
                if (dgvPoints.Rows[i].IsNewRow) continue;
                var xVal = dgvPoints.Rows[i].Cells["colX"].Value;
                var yVal = dgvPoints.Rows[i].Cells["colY"].Value;
                if (xVal != null && yVal != null
                    && double.TryParse(xVal.ToString(), out _)
                    && double.TryParse(yVal.ToString(), out _))
                {
                    validCount++;
                }
            }
            int totalRows = dgvPoints.Rows.Count - (dgvPoints.AllowUserToAddRows ? 1 : 0);
            lblStatus.Text = $"Tổng: {totalRows} dòng | Dữ liệu hợp lệ: {validCount} điểm";
        }

        private void BtnRemoveRow_Click(object? sender, EventArgs e)
        {
            if (dgvPoints.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn dòng cần xóa!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (DataGridViewRow row in dgvPoints.SelectedRows)
            {
                if (!row.IsNewRow)
                    dgvPoints.Rows.Remove(row);
            }
            ReNumberSTT();
            UpdateStatus();
        }

        private void BtnPaste_Click(object? sender, EventArgs e)
        {
            try
            {
                string clipText = Clipboard.GetText();
                if (string.IsNullOrWhiteSpace(clipText))
                {
                    MessageBox.Show("Clipboard trống!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var lines = clipText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                int added = 0;

                // Xóa dòng trống hiện tại
                var emptyRows = new List<int>();
                for (int i = dgvPoints.Rows.Count - 1; i >= 0; i--)
                {
                    if (dgvPoints.Rows[i].IsNewRow) continue;
                    bool isEmpty = true;
                    foreach (DataGridViewCell cell in dgvPoints.Rows[i].Cells)
                    {
                        if (cell.ColumnIndex == 0) continue; // Skip STT
                        if (cell.Value != null && !string.IsNullOrWhiteSpace(cell.Value.ToString()))
                        { isEmpty = false; break; }
                    }
                    if (isEmpty) emptyRows.Add(i);
                }
                foreach (int idx in emptyRows)
                    dgvPoints.Rows.RemoveAt(idx);

                int startSTT = (int)nudStartSTT.Value + dgvPoints.Rows.Count
                    - (dgvPoints.AllowUserToAddRows ? 1 : 0);

                foreach (var line in lines)
                {
                    // Tách bằng Tab hoặc dấu phẩy hoặc khoảng trắng nhiều
                    var parts = line.Split(new[] { '\t', ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length < 2) continue;

                    int rowIdx = dgvPoints.Rows.Add();
                    var row = dgvPoints.Rows[rowIdx];

                    row.Cells["colSTT"].Value = startSTT + added;

                    // Xác định format: Tên, X, Y, Z, Ghi chú
                    if (parts.Length >= 5)
                    {
                        row.Cells["colName"].Value = parts[0].Trim();
                        row.Cells["colX"].Value = parts[1].Trim();
                        row.Cells["colY"].Value = parts[2].Trim();
                        row.Cells["colZ"].Value = parts[3].Trim();
                        row.Cells["colDesc"].Value = parts[4].Trim();
                    }
                    else if (parts.Length == 4)
                    {
                        row.Cells["colName"].Value = parts[0].Trim();
                        row.Cells["colX"].Value = parts[1].Trim();
                        row.Cells["colY"].Value = parts[2].Trim();
                        row.Cells["colZ"].Value = parts[3].Trim();
                    }
                    else if (parts.Length == 3)
                    {
                        // Có thể Name,X,Y hoặc X,Y,Z
                        if (double.TryParse(parts[0], out _))
                        {
                            row.Cells["colX"].Value = parts[0].Trim();
                            row.Cells["colY"].Value = parts[1].Trim();
                            row.Cells["colZ"].Value = parts[2].Trim();
                            row.Cells["colName"].Value = $"P{startSTT + added}";
                        }
                        else
                        {
                            row.Cells["colName"].Value = parts[0].Trim();
                            row.Cells["colX"].Value = parts[1].Trim();
                            row.Cells["colY"].Value = parts[2].Trim();
                        }
                    }
                    else if (parts.Length == 2)
                    {
                        row.Cells["colX"].Value = parts[0].Trim();
                        row.Cells["colY"].Value = parts[1].Trim();
                        row.Cells["colName"].Value = $"P{startSTT + added}";
                    }

                    added++;
                }

                ReNumberSTT();
                UpdateStatus();
                MessageBox.Show($"Đã dán {added} điểm từ Clipboard!", "Thành công",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi khi dán: {ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnImportCSV_Click(object? sender, EventArgs e)
        {
            try
            {
                using var ofd = new OpenFileDialog
                {
                    Title = "Chọn file CSV chứa tọa độ",
                    Filter = "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
                };
                if (ofd.ShowDialog() != DialogResult.OK) return;

                var lines = File.ReadAllLines(ofd.FileName, Encoding.UTF8);
                int added = 0;

                // Xóa dòng trống
                for (int i = dgvPoints.Rows.Count - 1; i >= 0; i--)
                {
                    if (dgvPoints.Rows[i].IsNewRow) continue;
                    bool isEmpty = true;
                    foreach (DataGridViewCell cell in dgvPoints.Rows[i].Cells)
                    {
                        if (cell.ColumnIndex == 0) continue;
                        if (cell.Value != null && !string.IsNullOrWhiteSpace(cell.Value.ToString()))
                        { isEmpty = false; break; }
                    }
                    if (isEmpty) dgvPoints.Rows.RemoveAt(i);
                }

                int startSTT = (int)nudStartSTT.Value + dgvPoints.Rows.Count
                    - (dgvPoints.AllowUserToAddRows ? 1 : 0);

                // Skip header nếu dòng 1 không phải số
                int startLine = 0;
                if (lines.Length > 0)
                {
                    var firstParts = lines[0].Split(new[] { ',', '\t', ';' });
                    if (firstParts.Length >= 2 && !double.TryParse(firstParts[0], out _)
                        && !double.TryParse(firstParts[1], out _))
                        startLine = 1;
                }

                for (int l = startLine; l < lines.Length; l++)
                {
                    var parts = lines[l].Split(new[] { ',', '\t', ';' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length < 2) continue;

                    int rowIdx = dgvPoints.Rows.Add();
                    var row = dgvPoints.Rows[rowIdx];
                    row.Cells["colSTT"].Value = startSTT + added;

                    if (parts.Length >= 5)
                    {
                        row.Cells["colName"].Value = parts[0].Trim();
                        row.Cells["colX"].Value = parts[1].Trim();
                        row.Cells["colY"].Value = parts[2].Trim();
                        row.Cells["colZ"].Value = parts[3].Trim();
                        row.Cells["colDesc"].Value = parts[4].Trim();
                    }
                    else if (parts.Length == 4)
                    {
                        row.Cells["colName"].Value = parts[0].Trim();
                        row.Cells["colX"].Value = parts[1].Trim();
                        row.Cells["colY"].Value = parts[2].Trim();
                        row.Cells["colZ"].Value = parts[3].Trim();
                    }
                    else if (parts.Length == 3)
                    {
                        if (double.TryParse(parts[0], out _))
                        {
                            row.Cells["colX"].Value = parts[0].Trim();
                            row.Cells["colY"].Value = parts[1].Trim();
                            row.Cells["colZ"].Value = parts[2].Trim();
                            row.Cells["colName"].Value = $"P{startSTT + added}";
                        }
                        else
                        {
                            row.Cells["colName"].Value = parts[0].Trim();
                            row.Cells["colX"].Value = parts[1].Trim();
                            row.Cells["colY"].Value = parts[2].Trim();
                        }
                    }
                    else if (parts.Length == 2)
                    {
                        row.Cells["colX"].Value = parts[0].Trim();
                        row.Cells["colY"].Value = parts[1].Trim();
                        row.Cells["colName"].Value = $"P{startSTT + added}";
                    }
                    added++;
                }

                ReNumberSTT();
                UpdateStatus();
                MessageBox.Show($"Đã nhập {added} điểm từ {Path.GetFileName(ofd.FileName)}!",
                    "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi nhập CSV: {ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExportCSV_Click(object? sender, EventArgs e)
        {
            try
            {
                using var sfd = new SaveFileDialog
                {
                    Title = "Lưu file CSV",
                    Filter = "CSV Files (*.csv)|*.csv",
                    FileName = "ToaDo_Points.csv"
                };
                if (sfd.ShowDialog() != DialogResult.OK) return;

                var sb = new StringBuilder();
                sb.AppendLine("STT,Ten Diem,X,Y,Z,Ghi Chu");

                for (int i = 0; i < dgvPoints.Rows.Count; i++)
                {
                    var row = dgvPoints.Rows[i];
                    if (row.IsNewRow) continue;
                    string name = row.Cells["colName"].Value?.ToString() ?? "";
                    string x = row.Cells["colX"].Value?.ToString() ?? "";
                    string y = row.Cells["colY"].Value?.ToString() ?? "";
                    string z = row.Cells["colZ"].Value?.ToString() ?? "";
                    string desc = row.Cells["colDesc"].Value?.ToString() ?? "";
                    if (string.IsNullOrWhiteSpace(x) && string.IsNullOrWhiteSpace(y)) continue;
                    sb.AppendLine($"{i + 1},{name},{x},{y},{z},{desc}");
                }

                File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                MessageBox.Show($"Đã xuất file: {sfd.FileName}", "Thành công",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi xuất CSV: {ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnRun_Click(object? sender, EventArgs e)
        {
            // Thu thập dữ liệu từ bảng
            PointDataList.Clear();
            int stt = (int)nudStartSTT.Value;

            for (int i = 0; i < dgvPoints.Rows.Count; i++)
            {
                var row = dgvPoints.Rows[i];
                if (row.IsNewRow) continue;

                string xStr = row.Cells["colX"].Value?.ToString()?.Trim() ?? "";
                string yStr = row.Cells["colY"].Value?.ToString()?.Trim() ?? "";

                if (!double.TryParse(xStr, out double x)) continue;
                if (!double.TryParse(yStr, out double y)) continue;

                double z = 0;
                string zStr = row.Cells["colZ"].Value?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrEmpty(zStr)) double.TryParse(zStr, out z);

                string name = row.Cells["colName"].Value?.ToString()?.Trim() ?? $"P{stt}";
                string desc = row.Cells["colDesc"].Value?.ToString()?.Trim() ?? "";

                PointDataList.Add(new DirectPointData
                {
                    STT = stt,
                    PointName = name,
                    X = x,
                    Y = y,
                    Z = z,
                    Description = desc
                });
                stt++;
            }

            if (PointDataList.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu hợp lệ!\nVui lòng nhập tọa độ X, Y.",
                    "Thiếu dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Lưu settings
            SelectedPointGroup = cboPointGroup.SelectedIndex > 0 ? cboPointGroup.Text : "";
            SelectedPointStyle = cboPointStyle.SelectedIndex > 0 ? cboPointStyle.Text : "";
            SelectedLabelStyle = cboLabelStyle.SelectedIndex > 0 ? cboLabelStyle.Text : "";
            LayerName = string.IsNullOrWhiteSpace(txtLayer.Text) ? "COGO_POINTS" : txtLayer.Text.Trim();

            IsConfirmed = true;
            this.Close();
        }

        // ═══ Helper: Tool Button ═══
        private Button CreateToolButton(string text, int x, int y, int w, int h, Control parent)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = NavyDark,
                Font = new Font("Segoe UI", 8.5f),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderColor = BorderLight;
            btn.FlatAppearance.MouseOverBackColor = HoverNavy;
            parent.Controls.Add(btn);
            return btn;
        }
    }

    // ══════════════════════════════════════════════════════════════
    // LỆNH CHÍNH
    // ══════════════════════════════════════════════════════════════

    public class NhapPointTrucTiepCommands
    {
        /// <summary>
        /// Nhập tọa độ X, Y, Z trực tiếp vào bảng → Tạo COGO Point trên Civil 3D
        /// Không cần file Excel — gõ tay hoặc dán từ clipboard
        /// </summary>
        [CommandMethod("CTPO_NhapToaDoTrucTiep")]
        public static void CTPO_NhapToaDoTrucTiep()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;
            var civilDoc = CivilApplication.ActiveDocument;

            // 1. Mở form nhập liệu
            List<DirectPointData> pointList;
            string pointGroupName, pointStyleName, labelStyleName, layerName;

            using (var form = new FormNhapPointTrucTiep())
            {
                AcadApp.ShowModalDialog(form);
                if (!form.IsConfirmed) return;

                pointList = form.PointDataList;
                pointGroupName = form.SelectedPointGroup;
                pointStyleName = form.SelectedPointStyle;
                labelStyleName = form.SelectedLabelStyle;
                layerName = form.LayerName;
            }

            if (pointList.Count == 0)
            {
                ed.WriteMessage("\n⚠ Không có dữ liệu để tạo.");
                return;
            }

            ed.WriteMessage($"\n📍 Bắt đầu tạo {pointList.Count} COGO Points...");

            try
            {
                using var tr = db.TransactionManager.StartTransaction();

                CogoPointCollection cogoPointColl = civilDoc.CogoPoints;

                // Lấy Point Style ID
                ObjectId pointStyleId = ObjectId.Null;
                if (!string.IsNullOrEmpty(pointStyleName))
                {
                    try
                    {
                        pointStyleId = civilDoc.Styles.PointStyles[pointStyleName];
                    }
                    catch { }
                }

                // Lấy Label Style ID
                ObjectId labelStyleId = ObjectId.Null;
                if (!string.IsNullOrEmpty(labelStyleName))
                {
                    try
                    {
                        labelStyleId = civilDoc.Styles.LabelStyles.PointLabelStyles.LabelStyles[labelStyleName];
                    }
                    catch { }
                }

                int created = 0;
                int errors = 0;

                foreach (var pt in pointList)
                {
                    try
                    {
                        Point3d insertPt = new(pt.X, pt.Y, pt.Z);

                        // Tạo COGO Point
                        ObjectId ptId = cogoPointColl.Add(insertPt, pt.Description, true);
                        CogoPoint cogoPoint = tr.GetObject(ptId, OpenMode.ForWrite) as CogoPoint;

                        if (cogoPoint != null)
                        {
                            // Đặt tên (PointName)
                            if (!string.IsNullOrWhiteSpace(pt.PointName))
                                cogoPoint.PointName = pt.PointName;

                            // Đặt RawDescription
                            if (!string.IsNullOrWhiteSpace(pt.Description))
                                cogoPoint.RawDescription = pt.Description;

                            // Đặt Style
                            if (!pointStyleId.IsNull)
                                cogoPoint.StyleId = pointStyleId;

                            // Đặt Label Style
                            if (!labelStyleId.IsNull)
                                cogoPoint.LabelStyleId = labelStyleId;

                            created++;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        errors++;
                        ed.WriteMessage($"\n  ⚠ Lỗi tạo điểm {pt.PointName}: {ex.Message}");
                    }
                }

                // Cập nhật Point Group nếu chọn
                if (!string.IsNullOrEmpty(pointGroupName))
                {
                    try
                    {
                        PointGroupCollection pgColl = civilDoc.PointGroups;
                        ObjectId pgId;

                        if (pgColl.Contains(pointGroupName))
                        {
                            pgId = pgColl[pointGroupName];
                        }
                        else
                        {
                            pgId = pgColl.Add(pointGroupName);
                        }

                        // Set query theo description
                        PointGroup pg = tr.GetObject(pgId, OpenMode.ForWrite) as PointGroup;
                        if (pg != null)
                        {
                            pg.Update();
                        }
                    }
                    catch { }
                }

                // Cập nhật tất cả Point Groups
                try
                {
                    civilDoc.PointGroups.UpdateAllPointGroups();
                }
                catch { }

                tr.Commit();

                // ─── Hiển thị kết quả ───
                ed.WriteMessage("\n");
                ed.WriteMessage($"\n╔════════════════════════════════════════════════════╗");
                ed.WriteMessage($"\n║  📍 KẾT QUẢ NHẬP TỌA ĐỘ TRỰC TIẾP               ║");
                ed.WriteMessage($"\n╠════════════════════════════════════════════════════╣");
                ed.WriteMessage($"\n║  ✅ Đã tạo: {created,5} COGO Points                  ║");
                if (errors > 0)
                    ed.WriteMessage($"\n║  ⚠ Lỗi:    {errors,5} điểm                          ║");
                ed.WriteMessage($"\n╚════════════════════════════════════════════════════╝");

                // Bảng chi tiết
                ed.WriteMessage($"\n\n{"STT",-5} {"Tên điểm",-15} {"X (Easting)",-14} {"Y (Northing)",-14} {"Z",-10} {"Ghi chú",-20}");
                ed.WriteMessage($"\n{new string('─', 80)}");

                foreach (var pt in pointList)
                {
                    ed.WriteMessage($"\n{pt.STT,-5} {pt.PointName,-15} {pt.X,-14:F3} {pt.Y,-14:F3} {pt.Z,-10:F3} {pt.Description,-20}");
                }

                ed.WriteMessage($"\n{new string('═', 80)}");
                ed.WriteMessage($"\n\n✅ HOÀN THÀNH nhập {created} điểm!");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }
    }
}
