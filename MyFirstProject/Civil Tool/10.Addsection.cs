// 10.Addsection.cs - Form và logic thêm Section vào Corridor
// Chế độ 1: Từ SampleLine (logic cũ)
// Chế độ 2: Từ Polyline giao cắt alignment

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
using SolidBrush = System.Drawing.SolidBrush;
using RectangleF = System.Drawing.RectangleF;
using Pen = System.Drawing.Pen;
using FontStyle = System.Drawing.FontStyle;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace MyFirstProject
{
    /// <summary>
    /// Form tùy chọn cho lệnh CTC_AddAllSection
    /// Dark theme giống các form khác trong project
    /// </summary>
    public class AddSectionForm : CivilFormBase
    {

        // ─── Controls ───
        private ComboBox cboCorridor = null!;
        private WinLabel lblCorridorInfo = null!;

        private RadioButton rdoSampleLine = null!;
        private RadioButton rdoPolyline = null!;

        // Panel SampleLine
        private Panel pnlSampleLine = null!;
        private ComboBox cboSampleLineGroup = null!;

        // Panel Polyline
        private Panel pnlPolyline = null!;
        private Button btnSelectPolylines = null!;
        private WinLabel lblPolylineCount = null!;
        private NumericUpDown numTolerance = null!;

        private CheckBox chkClearOld = null!;
        private DataGridView dgvPreview = null!;
        private WinLabel lblTotal = null!;

        private Button btnPreview = null!;
        private Button btnExecute = null!;
        private Button btnCancel = null!;

        // ─── Data ───
        private List<CorridorInfo> corridors = new();
        private List<SampleLineGroupInfo> sampleLineGroups = new();
        private List<ObjectId> selectedPolylineIds = new();
        private List<StationPreviewItem> previewStations = new();

        public bool IsConfirmed { get; private set; } = false;
        public ObjectId SelectedCorridorId { get; private set; } = ObjectId.Null;
        public bool UseSampleLineMode => rdoSampleLine.Checked;
        public bool UsePolylineMode => rdoPolyline.Checked;
        public int SelectedSampleLineGroupIndex { get; private set; } = 0;
        public List<ObjectId> SelectedPolylineIds => selectedPolylineIds;
        public bool ClearOldStations => chkClearOld.Checked;
        public List<StationPreviewItem> PreviewStations => previewStations;

        private class CorridorInfo
        {
            public ObjectId Id { get; set; }
            public string Name { get; set; } = "";
            public int BaselineCount { get; set; }
        }

        private class SampleLineGroupInfo
        {
            public ObjectId Id { get; set; }
            public string Name { get; set; } = "";
            public int SampleLineCount { get; set; }
        }

        public class StationPreviewItem
        {
            public int STT { get; set; }
            public string BaselineName { get; set; } = "";
            public double Station { get; set; }
            public string Nguon { get; set; } = ""; // "SampleLine" hoặc "Polyline"
        }

        // ─── Constructor ───
        public AddSectionForm()
        {
            LoadCorridors();
            InitializeUI();
        }

        private void LoadCorridors()
        {
            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                CivilDocument cdoc = CivilApplication.ActiveDocument;
                var corIds = cdoc.CorridorCollection;
                foreach (ObjectId id in corIds)
                {
                    Corridor? cor = tr.GetObject(id, OpenMode.ForRead) as Corridor;
                    if (cor == null) continue;
                    corridors.Add(new CorridorInfo
                    {
                        Id = id,
                        Name = cor.Name,
                        BaselineCount = cor.Baselines.Count
                    });
                }
                tr.Commit();
            }
            catch { }
        }

        private void InitializeUI()
        {
            // ── Form ──
            InitBaseTheme("THÊM SECTION VÀO CORRIDOR", 750, 700);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            // ── HEADER ──
            var header = CreateHeader("THÊM SECTION VÀO CORRIDOR");
            this.Controls.Add(header);

            // ── SUBTITLE ──
            var subtitle = new WinLabel
            {
                Text = "Thêm section tại vị trí SampleLine hoặc vị trí Polyline giao cắt alignment.",
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
            // SECTION 1: CHỌN CORRIDOR
            // ═══════════════════════════════════════════════════
            var grpCorridor = CreateSection("Chọn Corridor", 0, y, 700, 85);
            mainPanel.Controls.Add(grpCorridor);

            grpCorridor.Controls.Add(CreateLabel("Corridor:", 12, 33));
            cboCorridor = new ComboBox
            {
                Location = new Point(100, 29),
                Size = new Size(350, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextDark,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f)
            };
            foreach (var cor in corridors)
                cboCorridor.Items.Add(cor.Name);
            if (cboCorridor.Items.Count > 0) cboCorridor.SelectedIndex = 0;
            cboCorridor.SelectedIndexChanged += (s, e) => OnCorridorChanged();
            grpCorridor.Controls.Add(cboCorridor);

            lblCorridorInfo = CreateLabel("", 470, 33);
            lblCorridorInfo.AutoSize = true;
            grpCorridor.Controls.Add(lblCorridorInfo);

            y += 100;

            // ═══════════════════════════════════════════════════
            // SECTION 2: CHẾ ĐỘ THÊM SECTION
            // ═══════════════════════════════════════════════════
            var grpMode = CreateSection("Chế độ thêm Section", 0, y, 700, 190);
            mainPanel.Controls.Add(grpMode);

            rdoSampleLine = new RadioButton
            {
                Text = "Từ SampleLine — Lấy station từ nhóm cọc có sẵn",
                Location = new Point(15, 30),
                Size = new Size(500, 24),
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f),
                Checked = true,
                FlatStyle = FlatStyle.Flat
            };
            rdoSampleLine.CheckedChanged += (s, e) => OnModeChanged();
            grpMode.Controls.Add(rdoSampleLine);

            rdoPolyline = new RadioButton
            {
                Text = "Từ Polyline — Chọn polyline giao cắt corridor, thêm section tại giao điểm",
                Location = new Point(15, 60),
                Size = new Size(650, 24),
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f),
                FlatStyle = FlatStyle.Flat
            };
            grpMode.Controls.Add(rdoPolyline);

            // ── Panel SampleLine ──
            pnlSampleLine = new Panel
            {
                Location = new Point(10, 95),
                Size = new Size(680, 45),
                BackColor = BgPanel
            };
            grpMode.Controls.Add(pnlSampleLine);

            pnlSampleLine.Controls.Add(CreateLabel("Nhóm SampleLine:", 5, 12));
            cboSampleLineGroup = new ComboBox
            {
                Location = new Point(160, 8),
                Size = new Size(350, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextDark,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f)
            };
            pnlSampleLine.Controls.Add(cboSampleLineGroup);

            // ── Panel Polyline ──
            pnlPolyline = new Panel
            {
                Location = new Point(10, 95),
                Size = new Size(680, 80),
                BackColor = BgPanel,
                Visible = false
            };
            grpMode.Controls.Add(pnlPolyline);

            btnSelectPolylines = CreateButton("Chọn Polyline trên bản vẽ...", 5, 8, 250, 32, pnlPolyline);
            btnSelectPolylines.Click += BtnSelectPolylines_Click;

            lblPolylineCount = CreateLabel("Chưa chọn polyline nào", 270, 14);
            lblPolylineCount.ForeColor = TextDim;
            lblPolylineCount.AutoSize = true;
            pnlPolyline.Controls.Add(lblPolylineCount);

            pnlPolyline.Controls.Add(CreateLabel("Tolerance (m):", 5, 50));
            numTolerance = new NumericUpDown
            {
                Location = new Point(120, 46),
                Size = new Size(80, 28),
                Minimum = (decimal)0.01,
                Maximum = 10,
                Value = (decimal)0.5,
                DecimalPlaces = 2,
                Increment = (decimal)0.1,
                BackColor = BgInput,
                ForeColor = TextDark
            };
            pnlPolyline.Controls.Add(numTolerance);

            var lblToleranceInfo = CreateLabel("(Khoảng cách tối đa từ polyline đến alignment để coi là giao cắt)", 210, 50);
            lblToleranceInfo.Font = new Font("Segoe UI", 8.5f);
            pnlPolyline.Controls.Add(lblToleranceInfo);

            y += 205;

            // ── Checkbox Clear ──
            chkClearOld = new CheckBox
            {
                Text = "Xóa các station bổ sung cũ trước khi thêm mới (ClearAdditionalStations)",
                Location = new Point(15, y),
                Size = new Size(600, 24),
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f),
                Checked = true,
                FlatStyle = FlatStyle.Flat
            };
            mainPanel.Controls.Add(chkClearOld);
            y += 40;

            // ═══════════════════════════════════════════════════
            // SECTION 3: PREVIEW
            // ═══════════════════════════════════════════════════
            var grpPreview = CreateSection("Preview danh sách Station sẽ thêm", 0, y, 700, 200);
            mainPanel.Controls.Add(grpPreview);

            dgvPreview = new DataGridView
            {
                Location = new Point(10, 30),
                Size = new Size(680, 130),
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
            dgvPreview.Columns.Add("STT", "STT");
            dgvPreview.Columns.Add("Baseline", "Baseline");
            dgvPreview.Columns.Add("Station", "Station (m)");
            dgvPreview.Columns.Add("Nguon", "Nguồn");
            dgvPreview.Columns["STT"]!.Width = 50;
            dgvPreview.Columns["Station"]!.Width = 120;
            dgvPreview.Columns["Nguon"]!.Width = 100;
            grpPreview.Controls.Add(dgvPreview);

            lblTotal = CreateLabel("Tổng: 0 station", 10, 170);
            lblTotal.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            grpPreview.Controls.Add(lblTotal);

            // ═══════════════════════════════════════════════════
            // BOTTOM BUTTONS
            // ═══════════════════════════════════════════════════
            var pnlBottom = CreateBottomBar();
            this.Controls.Add(pnlBottom);

            btnPreview = CreateButton("Preview", 380, 8, 110, 36, pnlBottom);
            btnPreview.Click += BtnPreview_Click;

            btnCancel = CreateButton("Hủy", 500, 8, 100, 36, pnlBottom);
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            btnExecute = CreatePrimaryButton("Thực hiện", 610, 8, 120, 36, pnlBottom);
            btnExecute.Click += BtnExecute_Click;

            // Initialize
            if (corridors.Count > 0)
                OnCorridorChanged();
        }

        // ═══════════════════════════════════════════════════
        // EVENT HANDLERS
        // ═══════════════════════════════════════════════════

        private void OnCorridorChanged()
        {
            if (cboCorridor.SelectedIndex < 0 || cboCorridor.SelectedIndex >= corridors.Count) return;
            var cor = corridors[cboCorridor.SelectedIndex];
            SelectedCorridorId = cor.Id;
            lblCorridorInfo.Text = $"Baselines: {cor.BaselineCount}";

            // Load SampleLine groups
            LoadSampleLineGroups(cor.Id);
        }

        private void LoadSampleLineGroups(ObjectId corridorId)
        {
            sampleLineGroups.Clear();
            cboSampleLineGroup.Items.Clear();

            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                Corridor? corridor = tr.GetObject(corridorId, OpenMode.ForRead) as Corridor;
                if (corridor == null) { tr.Commit(); return; }

                foreach (Baseline baseline in corridor.Baselines)
                {
                    Alignment? alignment = tr.GetObject(baseline.AlignmentId, OpenMode.ForRead) as Alignment;
                    if (alignment == null) continue;

                    ObjectIdCollection groupIds = alignment.GetSampleLineGroupIds();
                    for (int i = 0; i < groupIds.Count; i++)
                    {
                        SampleLineGroup? slg = tr.GetObject(groupIds[i], OpenMode.ForRead) as SampleLineGroup;
                        if (slg == null) continue;

                        int slCount = slg.GetSampleLineIds().Count;
                        sampleLineGroups.Add(new SampleLineGroupInfo
                        {
                            Id = groupIds[i],
                            Name = slg.Name,
                            SampleLineCount = slCount
                        });
                        cboSampleLineGroup.Items.Add($"{slg.Name} ({slCount} cọc) — {alignment.Name}");
                    }
                }

                if (cboSampleLineGroup.Items.Count > 0)
                    cboSampleLineGroup.SelectedIndex = 0;

                tr.Commit();
            }
            catch { }
        }

        private void OnModeChanged()
        {
            pnlSampleLine.Visible = rdoSampleLine.Checked;
            pnlPolyline.Visible = rdoPolyline.Checked;

            // Clear preview
            dgvPreview.Rows.Clear();
            previewStations.Clear();
            lblTotal.Text = "Tổng: 0 station";
        }

        private void BtnSelectPolylines_Click(object? sender, EventArgs e)
        {
            // Ẩn form tạm thời để chọn polyline trong bản vẽ
            this.Hide();

            try
            {
                // Cho phép chọn nhiều polyline
                TypedValue[] filter = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
                };
                SelectionFilter selFilter = new SelectionFilter(filter);

                A.Ed.WriteMessage("\nChọn các polyline giao cắt corridor (nhấn Enter khi xong): ");
                PromptSelectionResult selResult = A.Ed.GetSelection(selFilter);

                selectedPolylineIds.Clear();
                if (selResult.Status == PromptStatus.OK)
                {
                    foreach (ObjectId id in selResult.Value.GetObjectIds())
                    {
                        selectedPolylineIds.Add(id);
                    }
                }
            }
            catch { }

            // Hiện lại form
            this.Show();
            lblPolylineCount.Text = $"Đã chọn {selectedPolylineIds.Count} polyline";
        }

        private void BtnPreview_Click(object? sender, EventArgs e)
        {
            previewStations.Clear();
            dgvPreview.Rows.Clear();

            if (SelectedCorridorId == ObjectId.Null)
            {
                MessageBox.Show("Vui lòng chọn Corridor!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                Corridor? corridor = tr.GetObject(SelectedCorridorId, OpenMode.ForRead) as Corridor;
                if (corridor == null) { tr.Commit(); return; }

                if (rdoSampleLine.Checked)
                {
                    PreviewSampleLineMode(tr, corridor);
                }
                else
                {
                    PreviewPolylineMode(tr, corridor);
                }

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi khi preview: {ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Update grid
            int stt = 1;
            foreach (var item in previewStations.OrderBy(s => s.Station))
            {
                item.STT = stt++;
                dgvPreview.Rows.Add(item.STT, item.BaselineName, Math.Round(item.Station, 3), item.Nguon);
            }
            lblTotal.Text = $"Tổng: {previewStations.Count} station";
        }

        private void PreviewSampleLineMode(Transaction tr, Corridor corridor)
        {
            if (cboSampleLineGroup.SelectedIndex < 0) return;

            foreach (Baseline baseline in corridor.Baselines)
            {
                Alignment? alignment = tr.GetObject(baseline.AlignmentId, OpenMode.ForRead) as Alignment;
                if (alignment == null) continue;

                ObjectIdCollection groupIds = alignment.GetSampleLineGroupIds();
                if (groupIds.Count == 0) continue;

                // Chọn nhóm sampleline theo index
                int groupIdx = Math.Min(cboSampleLineGroup.SelectedIndex, groupIds.Count - 1);
                SampleLineGroup? slGroup = tr.GetObject(groupIds[groupIdx], OpenMode.ForRead) as SampleLineGroup;
                if (slGroup == null) continue;

                foreach (ObjectId slId in slGroup.GetSampleLineIds())
                {
                    SampleLine? sl = tr.GetObject(slId, OpenMode.ForRead) as SampleLine;
                    if (sl == null) continue;

                    // Kiểm tra station nằm trong vùng baseline region
                    foreach (BaselineRegion region in baseline.BaselineRegions)
                    {
                        if (sl.Station > region.StartStation && sl.Station < region.EndStation)
                        {
                            previewStations.Add(new StationPreviewItem
                            {
                                BaselineName = baseline.Name,
                                Station = sl.Station,
                                Nguon = "SampleLine"
                            });
                            break;
                        }
                    }
                }
            }
        }

        private void PreviewPolylineMode(Transaction tr, Corridor corridor)
        {
            if (selectedPolylineIds.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn polyline trước!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            double tolerance = (double)numTolerance.Value;

            foreach (Baseline baseline in corridor.Baselines)
            {
                Alignment? alignment = tr.GetObject(baseline.AlignmentId, OpenMode.ForRead) as Alignment;
                if (alignment == null) continue;

                foreach (ObjectId plId in selectedPolylineIds)
                {
                    Polyline? polyline = tr.GetObject(plId, OpenMode.ForRead) as Polyline;
                    if (polyline == null) continue;

                    // Tìm giao cắt giữa polyline và alignment
                    List<double> intersectStations = FindIntersectionStations(alignment, polyline, tolerance);

                    foreach (double station in intersectStations)
                    {
                        // Kiểm tra trùng lặp
                        bool exists = previewStations.Any(s =>
                            s.BaselineName == baseline.Name &&
                            Math.Abs(s.Station - station) < 0.01);

                        if (!exists)
                        {
                            previewStations.Add(new StationPreviewItem
                            {
                                BaselineName = baseline.Name,
                                Station = station,
                                Nguon = "Polyline"
                            });
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Tìm station trên alignment tại các vị trí polyline giao cắt
        /// Phương pháp: duyệt dense points trên polyline, tính offset đến alignment,
        /// phát hiện đổi dấu offset (polyline đi qua alignment)
        /// </summary>
        public static List<double> FindIntersectionStations(Alignment alignment, Polyline polyline, double tolerance)
        {
            List<double> stations = new();

            try
            {
                // Duyệt từng đoạn nhỏ của polyline
                double totalLength = polyline.Length;
                int numSamples = Math.Max((int)(totalLength / 0.5), 100); // Sample mỗi 0.5m hoặc ít nhất 100 điểm
                double step = totalLength / numSamples;

                double prevStation = 0;
                double prevOffset = double.NaN;

                for (int i = 0; i <= numSamples; i++)
                {
                    double dist = i * step;
                    if (dist > totalLength) dist = totalLength;

                    Point3d pt = polyline.GetPointAtDist(dist);

                    try
                    {
                        double station = 0, offset = 0;
                        alignment.StationOffset(pt.X, pt.Y, ref station, ref offset);

                        // Phát hiện giao cắt: offset đổi dấu
                        if (!double.IsNaN(prevOffset))
                        {
                            if ((prevOffset > 0 && offset < 0) || (prevOffset < 0 && offset > 0))
                            {
                                // Interpolation tìm station chính xác hơn
                                double ratio = Math.Abs(prevOffset) / (Math.Abs(prevOffset) + Math.Abs(offset));
                                double exactStation = prevStation + (station - prevStation) * ratio;
                                stations.Add(exactStation);
                            }
                            else if (Math.Abs(offset) <= tolerance)
                            {
                                // Polyline nằm rất gần alignment
                                if (!stations.Any(s => Math.Abs(s - station) < 1.0))
                                {
                                    stations.Add(station);
                                }
                            }
                        }

                        prevStation = station;
                        prevOffset = offset;
                    }
                    catch
                    {
                        // Điểm ngoài phạm vi alignment - reset
                        prevOffset = double.NaN;
                    }
                }
            }
            catch { }

            return stations.Distinct().OrderBy(s => s).ToList();
        }

        private void BtnExecute_Click(object? sender, EventArgs e)
        {
            if (previewStations.Count == 0)
            {
                // Auto preview trước
                BtnPreview_Click(null, EventArgs.Empty);

                if (previewStations.Count == 0)
                {
                    MessageBox.Show("Không có station nào để thêm! Vui lòng Preview trước.",
                        "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            var confirm = MessageBox.Show(
                $"Bạn có muốn thêm {previewStations.Count} station vào corridor không?",
                "Xác nhận",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm == DialogResult.Yes)
            {
                IsConfirmed = true;
                SelectedSampleLineGroupIndex = cboSampleLineGroup.SelectedIndex >= 0 ? cboSampleLineGroup.SelectedIndex : 0;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        // (Removed UI helpers)
    }
}
