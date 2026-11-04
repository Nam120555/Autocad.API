// 10.Addsection.cs - Form và logic thêm Section vào Corridor
// Chế độ 1: Từ SampleLine (logic cũ)
// Chế độ 2: Từ Polyline giao cắt alignment
// HỖ TRỢ: Chọn NHIỀU corridor bằng checkbox (CheckedListBox)

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
    /// Hỗ trợ chọn NHIỀU corridor bằng checkbox
    /// Dark theme giống các form khác trong project
    /// </summary>
    public class AddSectionForm : CivilFormBase
    {

        // ─── Controls ───
        private CheckedListBox clbCorridors = null!;
        private WinLabel lblCorridorInfo = null!;
        private Button btnSelectAllCorridors = null!;
        private Button btnDeselectAllCorridors = null!;

        // Alignment selection
        private CheckedListBox clbAlignments = null!;
        private WinLabel lblAlignmentInfo = null!;
        private Button btnSelectAllAlignments = null!;
        private Button btnDeselectAllAlignments = null!;
        private Button btnPickAlignmentOnPlan = null!;

        private RadioButton rdoSampleLine = null!;
        private RadioButton rdoPolyline = null!;

        // Panel SampleLine
        private Panel pnlSampleLine = null!;
        private CheckedListBox clbSampleLineGroups = null!;
        private Button btnSelectAllSLG = null!;
        private Button btnDeselectAllSLG = null!;

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
        private List<AlignmentDisplayInfo> alignmentDisplayInfos = new();
        private List<SampleLineGroupInfo> sampleLineGroups = new();
        private List<ObjectId> selectedPolylineIds = new();
        private List<StationPreviewItem> previewStations = new();

        public bool IsConfirmed { get; private set; } = false;

        /// <summary>
        /// Danh sách các CorridorId đã chọn (multi-select)
        /// </summary>
        public List<ObjectId> SelectedCorridorIds { get; private set; } = new();

        // Backward compat: trả corridor đầu tiên
        public ObjectId SelectedCorridorId => SelectedCorridorIds.Count > 0 ? SelectedCorridorIds[0] : ObjectId.Null;

        public bool UseSampleLineMode => rdoSampleLine.Checked;
        public bool UsePolylineMode => rdoPolyline.Checked;

        /// <summary>
        /// Danh sách index SampleLineGroup đã chọn (multi-select)
        /// </summary>
        public List<int> SelectedSampleLineGroupIndices { get; private set; } = new();

        // Backward compat
        public int SelectedSampleLineGroupIndex => SelectedSampleLineGroupIndices.Count > 0 ? SelectedSampleLineGroupIndices[0] : 0;

        /// <summary>
        /// Danh sách AlignmentId đã chọn (multi-select)
        /// </summary>
        public List<ObjectId> SelectedAlignmentIds { get; private set; } = new();

        public List<ObjectId> SelectedPolylineIds => selectedPolylineIds;
        public bool ClearOldStations => chkClearOld.Checked;
        public List<StationPreviewItem> PreviewStations => previewStations;

        private class CorridorInfo
        {
            public ObjectId Id { get; set; }
            public string Name { get; set; } = "";
            public int BaselineCount { get; set; }
        }

        private class AlignmentDisplayInfo
        {
            public ObjectId AlignmentId { get; set; }
            public string AlignmentName { get; set; } = "";
            public string CorridorName { get; set; } = "";
            public string BaselineName { get; set; } = "";
            public ObjectId CorridorId { get; set; }
        }

        private class SampleLineGroupInfo
        {
            public ObjectId Id { get; set; }
            public string Name { get; set; } = "";
            public int SampleLineCount { get; set; }
            public ObjectId CorridorId { get; set; }    // Corridor chứa group này
            public ObjectId AlignmentId { get; set; }   // Alignment chứa group này
        }

        public class StationPreviewItem
        {
            public int STT { get; set; }
            public string CorridorName { get; set; } = "";
            public string BaselineName { get; set; } = "";
            public double Station { get; set; }
            public string Nguon { get; set; } = ""; // "SampleLine" hoặc "Polyline"
            public ObjectId CorridorId { get; set; }
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
            InitBaseTheme("THÊM SECTION VÀO CORRIDOR", 780, 920);
            this.FormBorderStyle = FormBorderStyle.Sizable;

            // ── HEADER ──
            var header = CreateHeader("THÊM SECTION VÀO CORRIDOR");
            this.Controls.Add(header);

            // ── SUBTITLE ──
            var subtitle = new WinLabel
            {
                Text = "Thêm section tại vị trí SampleLine hoặc vị trí Polyline giao cắt alignment. Hỗ trợ chọn nhiều tuyến.",
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
            // SECTION 1: CHỌN CORRIDOR (CheckedListBox — multi-select)
            // ═══════════════════════════════════════════════════
            int corridorListHeight = Math.Max(60, Math.Min(corridors.Count * 22 + 10, 120));
            var grpCorridor = CreateSection("Chọn Corridor (tích chọn nhiều tuyến)", 0, y, 730, corridorListHeight + 65);
            mainPanel.Controls.Add(grpCorridor);

            // Buttons: Chọn tất cả / Bỏ chọn tất cả
            btnSelectAllCorridors = CreateButton("✓ Chọn tất cả", 12, 28, 120, 28, grpCorridor);
            btnSelectAllCorridors.Font = new Font("Segoe UI", 8.5f);
            btnSelectAllCorridors.Click += (s, e) =>
            {
                for (int i = 0; i < clbCorridors.Items.Count; i++)
                    clbCorridors.SetItemChecked(i, true);
                OnCorridorSelectionChanged();
            };

            btnDeselectAllCorridors = CreateButton("✗ Bỏ chọn", 140, 28, 100, 28, grpCorridor);
            btnDeselectAllCorridors.Font = new Font("Segoe UI", 8.5f);
            btnDeselectAllCorridors.Click += (s, e) =>
            {
                for (int i = 0; i < clbCorridors.Items.Count; i++)
                    clbCorridors.SetItemChecked(i, false);
                OnCorridorSelectionChanged();
            };

            lblCorridorInfo = CreateLabel("", 260, 32);
            lblCorridorInfo.AutoSize = true;
            lblCorridorInfo.ForeColor = NavyMedium;
            lblCorridorInfo.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            grpCorridor.Controls.Add(lblCorridorInfo);

            clbCorridors = new CheckedListBox
            {
                Location = new Point(12, 60),
                Size = new Size(706, corridorListHeight),
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f),
                BorderStyle = BorderStyle.FixedSingle,
                CheckOnClick = true,
                IntegralHeight = false
            };
            foreach (var cor in corridors)
                clbCorridors.Items.Add($"{cor.Name}  ({cor.BaselineCount} baselines)");
            // Mặc định chọn tất cả
            for (int i = 0; i < clbCorridors.Items.Count; i++)
                clbCorridors.SetItemChecked(i, true);
            clbCorridors.ItemCheck += (s, e) =>
            {
                // ItemCheck fires before the state changes, use BeginInvoke to get correct state
                this.BeginInvoke(new Action(() => OnCorridorSelectionChanged()));
            };
            grpCorridor.Controls.Add(clbCorridors);

            y += grpCorridor.Height + 10;

            // ═══════════════════════════════════════════════════
            // SECTION 1.5: CHỌN TUYẾN (ALIGNMENT) — BỘ LỌC
            // ═══════════════════════════════════════════════════
            int alignmentListHeight = 90;
            var grpAlignment = CreateSection("Chọn Tuyến — Alignment (tích chọn tuyến cần xử lý)", 0, y, 730, alignmentListHeight + 65);
            mainPanel.Controls.Add(grpAlignment);

            btnSelectAllAlignments = CreateButton("✓ Chọn tất cả", 12, 28, 120, 28, grpAlignment);
            btnSelectAllAlignments.Font = new Font("Segoe UI", 8.5f);
            btnSelectAllAlignments.Click += (s, e) =>
            {
                for (int i = 0; i < clbAlignments.Items.Count; i++)
                    clbAlignments.SetItemChecked(i, true);
                OnAlignmentSelectionChanged();
            };

            btnDeselectAllAlignments = CreateButton("✗ Bỏ chọn", 140, 28, 100, 28, grpAlignment);
            btnDeselectAllAlignments.Font = new Font("Segoe UI", 8.5f);
            btnDeselectAllAlignments.Click += (s, e) =>
            {
                for (int i = 0; i < clbAlignments.Items.Count; i++)
                    clbAlignments.SetItemChecked(i, false);
                OnAlignmentSelectionChanged();
            };

            btnPickAlignmentOnPlan = CreateButton("📌 Chọn trên bình đồ", 248, 28, 170, 28, grpAlignment);
            btnPickAlignmentOnPlan.Font = new Font("Segoe UI", 8.5f);
            btnPickAlignmentOnPlan.Click += BtnPickAlignmentOnPlan_Click;

            lblAlignmentInfo = CreateLabel("", 430, 32);
            lblAlignmentInfo.AutoSize = true;
            lblAlignmentInfo.ForeColor = NavyMedium;
            lblAlignmentInfo.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            grpAlignment.Controls.Add(lblAlignmentInfo);

            clbAlignments = new CheckedListBox
            {
                Location = new Point(12, 60),
                Size = new Size(706, alignmentListHeight),
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f),
                BorderStyle = BorderStyle.FixedSingle,
                CheckOnClick = true,
                IntegralHeight = false
            };
            clbAlignments.ItemCheck += (s, e) =>
            {
                this.BeginInvoke(new Action(() => OnAlignmentSelectionChanged()));
            };
            grpAlignment.Controls.Add(clbAlignments);

            y += grpAlignment.Height + 10;

            // ═══════════════════════════════════════════════════
            // SECTION 2: CHẾ ĐỘ THÊM SECTION
            // ═══════════════════════════════════════════════════
            var grpMode = CreateSection("Chế độ thêm Section", 0, y, 730, 250);
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

            // ── Panel SampleLine (CheckedListBox — multi-select) ──
            pnlSampleLine = new Panel
            {
                Location = new Point(10, 95),
                Size = new Size(710, 145),
                BackColor = BgPanel
            };
            grpMode.Controls.Add(pnlSampleLine);

            pnlSampleLine.Controls.Add(CreateLabel("Nhóm SampleLine (tích chọn nhiều nhóm):", 5, 5));

            btnSelectAllSLG = CreateButton("✓ Tất cả", 400, 0, 90, 26, pnlSampleLine);
            btnSelectAllSLG.Font = new Font("Segoe UI", 8f);
            btnSelectAllSLG.Click += (s, e) =>
            {
                for (int i = 0; i < clbSampleLineGroups.Items.Count; i++)
                    clbSampleLineGroups.SetItemChecked(i, true);
            };

            btnDeselectAllSLG = CreateButton("✗ Bỏ chọn", 496, 0, 90, 26, pnlSampleLine);
            btnDeselectAllSLG.Font = new Font("Segoe UI", 8f);
            btnDeselectAllSLG.Click += (s, e) =>
            {
                for (int i = 0; i < clbSampleLineGroups.Items.Count; i++)
                    clbSampleLineGroups.SetItemChecked(i, false);
            };

            clbSampleLineGroups = new CheckedListBox
            {
                Location = new Point(5, 28),
                Size = new Size(700, 110),
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f),
                BorderStyle = BorderStyle.FixedSingle,
                CheckOnClick = true,
                IntegralHeight = false
            };
            pnlSampleLine.Controls.Add(clbSampleLineGroups);

            // ── Panel Polyline ──
            pnlPolyline = new Panel
            {
                Location = new Point(10, 95),
                Size = new Size(710, 80),
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

            y += grpMode.Height + 10;

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
            y += 35;

            // ═══════════════════════════════════════════════════
            // SECTION 3: PREVIEW
            // ═══════════════════════════════════════════════════
            var grpPreview = CreateSection("Preview danh sách Station sẽ thêm", 0, y, 730, 220);
            mainPanel.Controls.Add(grpPreview);

            dgvPreview = new DataGridView
            {
                Location = new Point(10, 30),
                Size = new Size(710, 150),
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
            dgvPreview.Columns.Add("Corridor", "Corridor");
            dgvPreview.Columns.Add("Baseline", "Baseline");
            dgvPreview.Columns.Add("Station", "Station (m)");
            dgvPreview.Columns.Add("Nguon", "Nguồn");
            dgvPreview.Columns["STT"]!.Width = 45;
            dgvPreview.Columns["Corridor"]!.Width = 140;
            dgvPreview.Columns["Station"]!.Width = 100;
            dgvPreview.Columns["Nguon"]!.Width = 85;
            grpPreview.Controls.Add(dgvPreview);

            lblTotal = CreateLabel("Tổng: 0 station", 10, 190);
            lblTotal.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            grpPreview.Controls.Add(lblTotal);

            // ═══════════════════════════════════════════════════
            // BOTTOM BUTTONS
            // ═══════════════════════════════════════════════════
            var pnlBottom = CreateBottomBar();
            this.Controls.Add(pnlBottom);

            btnPreview = CreateButton("Preview", 380, 8, 110, 36, pnlBottom);
            btnPreview.Click += BtnPreview_Click;

            btnCancel = CreateButton("Hủy", 510, 8, 100, 36, pnlBottom);
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            btnExecute = CreatePrimaryButton("Thực hiện", 620, 8, 140, 36, pnlBottom);
            btnExecute.Click += BtnExecute_Click;

            // Initialize: load SampleLine groups cho tất cả corridor đã chọn
            OnCorridorSelectionChanged();
        }

        // ═══════════════════════════════════════════════════
        // EVENT HANDLERS
        // ═══════════════════════════════════════════════════

        /// <summary>
        /// Khi thay đổi corridor được chọn (checkbox) → cập nhật SampleLine groups
        /// </summary>
        private void OnCorridorSelectionChanged()
        {
            // Cập nhật danh sách corridor đã chọn
            SelectedCorridorIds.Clear();
            int checkedCount = 0;
            for (int i = 0; i < clbCorridors.Items.Count; i++)
            {
                if (clbCorridors.GetItemChecked(i) && i < corridors.Count)
                {
                    SelectedCorridorIds.Add(corridors[i].Id);
                    checkedCount++;
                }
            }

            lblCorridorInfo.Text = $"Đã chọn {checkedCount}/{corridors.Count} corridor";

            // Load Alignment cho tất cả corridor đã chọn
            LoadAlignmentsForSelected();

            // Load SampleLine groups cho tất cả corridor đã chọn
            LoadSampleLineGroupsForSelected();
        }

        /// <summary>
        /// Load tất cả Alignment (tuyến) từ các corridor đã chọn
        /// </summary>
        private void LoadAlignmentsForSelected()
        {
            alignmentDisplayInfos.Clear();
            clbAlignments.Items.Clear();
            SelectedAlignmentIds.Clear();

            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();

                foreach (ObjectId corridorId in SelectedCorridorIds)
                {
                    Corridor? corridor = tr.GetObject(corridorId, OpenMode.ForRead) as Corridor;
                    if (corridor == null) continue;

                    foreach (Baseline baseline in corridor.Baselines)
                    {
                        Alignment? alignment = tr.GetObject(baseline.AlignmentId, OpenMode.ForRead) as Alignment;
                        if (alignment == null) continue;

                        // Tránh trùng lặp alignment (1 alignment có thể thuộc nhiều baseline)
                        bool alreadyAdded = alignmentDisplayInfos.Any(a =>
                            a.AlignmentId == baseline.AlignmentId && a.CorridorId == corridorId);
                        if (alreadyAdded) continue;

                        alignmentDisplayInfos.Add(new AlignmentDisplayInfo
                        {
                            AlignmentId = baseline.AlignmentId,
                            AlignmentName = alignment.Name,
                            CorridorName = corridor.Name,
                            BaselineName = baseline.Name,
                            CorridorId = corridorId
                        });

                        string displayText = $"{alignment.Name}  —  Baseline: {baseline.Name}  [{corridor.Name}]";
                        clbAlignments.Items.Add(displayText);
                    }
                }

                // Mặc định chọn tất cả alignment
                for (int i = 0; i < clbAlignments.Items.Count; i++)
                    clbAlignments.SetItemChecked(i, true);

                tr.Commit();
            }
            catch { }

            OnAlignmentSelectionChanged();
        }

        /// <summary>
        /// Khi thay đổi alignment được chọn (checkbox)
        /// </summary>
        private void OnAlignmentSelectionChanged()
        {
            SelectedAlignmentIds.Clear();
            int checkedCount = 0;
            for (int i = 0; i < clbAlignments.Items.Count; i++)
            {
                if (clbAlignments.GetItemChecked(i) && i < alignmentDisplayInfos.Count)
                {
                    SelectedAlignmentIds.Add(alignmentDisplayInfos[i].AlignmentId);
                    checkedCount++;
                }
            }
            lblAlignmentInfo.Text = $"Đã chọn {checkedCount}/{alignmentDisplayInfos.Count} tuyến";
        }

        /// <summary>
        /// Load tất cả SampleLine groups từ các corridor đã chọn
        /// </summary>
        private void LoadSampleLineGroupsForSelected()
        {
            sampleLineGroups.Clear();
            clbSampleLineGroups.Items.Clear();

            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();

                foreach (ObjectId corridorId in SelectedCorridorIds)
                {
                    Corridor? corridor = tr.GetObject(corridorId, OpenMode.ForRead) as Corridor;
                    if (corridor == null) continue;

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
                                SampleLineCount = slCount,
                                CorridorId = corridorId,
                                AlignmentId = baseline.AlignmentId
                            });

                            string displayText = $"{slg.Name} ({slCount} cọc) — {alignment.Name} [{corridor.Name}]";
                            clbSampleLineGroups.Items.Add(displayText);
                        }
                    }
                }

                // Mặc định chọn tất cả SampleLine groups
                for (int i = 0; i < clbSampleLineGroups.Items.Count; i++)
                    clbSampleLineGroups.SetItemChecked(i, true);

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

            if (SelectedCorridorIds.Count == 0)
            {
                MessageBox.Show("Vui lòng tích chọn ít nhất 1 Corridor!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();

                foreach (ObjectId corridorId in SelectedCorridorIds)
                {
                    Corridor? corridor = tr.GetObject(corridorId, OpenMode.ForRead) as Corridor;
                    if (corridor == null) continue;

                    if (rdoSampleLine.Checked)
                    {
                        PreviewSampleLineMode(tr, corridor, corridorId);
                    }
                    else
                    {
                        PreviewPolylineMode(tr, corridor, corridorId);
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
            int stt = 1;
            foreach (var item in previewStations.OrderBy(s => s.CorridorName).ThenBy(s => s.BaselineName).ThenBy(s => s.Station))
            {
                item.STT = stt++;
                dgvPreview.Rows.Add(item.STT, item.CorridorName, item.BaselineName, Math.Round(item.Station, 3), item.Nguon);
            }
            lblTotal.Text = $"Tổng: {previewStations.Count} station trên {SelectedCorridorIds.Count} corridor";
        }

        private void PreviewSampleLineMode(Transaction tr, Corridor corridor, ObjectId corridorId)
        {
            // Lấy danh sách SampleLineGroup đã chọn (checked) cho corridor này
            List<ObjectId> checkedGroupIds = new();
            for (int i = 0; i < clbSampleLineGroups.Items.Count; i++)
            {
                if (clbSampleLineGroups.GetItemChecked(i) && i < sampleLineGroups.Count)
                {
                    var slgInfo = sampleLineGroups[i];
                    if (slgInfo.CorridorId == corridorId)
                    {
                        checkedGroupIds.Add(slgInfo.Id);
                    }
                }
            }

            if (checkedGroupIds.Count == 0) return;

            foreach (Baseline baseline in corridor.Baselines)
            {
                // Bỏ qua baseline nếu alignment không được chọn
                if (!SelectedAlignmentIds.Contains(baseline.AlignmentId)) continue;

                Alignment? alignment = tr.GetObject(baseline.AlignmentId, OpenMode.ForRead) as Alignment;
                if (alignment == null) continue;

                ObjectIdCollection groupIds = alignment.GetSampleLineGroupIds();
                if (groupIds.Count == 0) continue;

                // Duyệt tất cả groups đã chọn của alignment này
                for (int gi = 0; gi < groupIds.Count; gi++)
                {
                    if (!checkedGroupIds.Contains(groupIds[gi])) continue;

                    SampleLineGroup? slGroup = tr.GetObject(groupIds[gi], OpenMode.ForRead) as SampleLineGroup;
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
                                // Kiểm tra trùng
                                bool exists = previewStations.Any(s =>
                                    s.CorridorId == corridorId &&
                                    s.BaselineName == baseline.Name &&
                                    Math.Abs(s.Station - sl.Station) < 0.01);

                                if (!exists)
                                {
                                    previewStations.Add(new StationPreviewItem
                                    {
                                        CorridorName = corridor.Name,
                                        BaselineName = baseline.Name,
                                        Station = sl.Station,
                                        Nguon = "SampleLine",
                                        CorridorId = corridorId
                                    });
                                }
                                break;
                            }
                        }
                    }
                }
            }
        }

        private void PreviewPolylineMode(Transaction tr, Corridor corridor, ObjectId corridorId)
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
                // Bỏ qua baseline nếu alignment không được chọn
                if (!SelectedAlignmentIds.Contains(baseline.AlignmentId)) continue;

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
                            s.CorridorId == corridorId &&
                            s.BaselineName == baseline.Name &&
                            Math.Abs(s.Station - station) < 0.01);

                        if (!exists)
                        {
                            previewStations.Add(new StationPreviewItem
                            {
                                CorridorName = corridor.Name,
                                BaselineName = baseline.Name,
                                Station = station,
                                Nguon = "Polyline",
                                CorridorId = corridorId
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
                $"Bạn có muốn thêm {previewStations.Count} station vào {SelectedCorridorIds.Count} corridor không?",
                "Xác nhận",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm == DialogResult.Yes)
            {
                IsConfirmed = true;
                // Lưu danh sách SLG indices đã chọn
                SelectedSampleLineGroupIndices.Clear();
                for (int i = 0; i < clbSampleLineGroups.Items.Count; i++)
                {
                    if (clbSampleLineGroups.GetItemChecked(i))
                        SelectedSampleLineGroupIndices.Add(i);
                }
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        /// <summary>
        /// Ẩn form, cho user pick alignment trực tiếp trên bình đồ (bản vẽ),
        /// sau đó tự check/uncheck đúng alignment đã pick trong danh sách.
        /// </summary>
        private void BtnPickAlignmentOnPlan_Click(object? sender, EventArgs e)
        {
            this.Hide();

            try
            {
                // Bỏ chọn tất cả trước
                for (int i = 0; i < clbAlignments.Items.Count; i++)
                    clbAlignments.SetItemChecked(i, false);

                // Cho user pick nhiều alignment trên bình đồ
                bool continueSelecting = true;
                int pickedCount = 0;

                while (continueSelecting)
                {
                    A.Ed.WriteMessage($"\n[{pickedCount} tuyến đã chọn] — Pick tuyến (Alignment) trên bình đồ (Enter/ESC để kết thúc): ");

                    PromptEntityOptions peo = new("\nChọn tuyến (Alignment) trên bình đồ (Enter để kết thúc): ");
                    peo.SetRejectMessage("\n— Đối tượng không phải Alignment! Vui lòng chọn lại.");
                    peo.AddAllowedClass(typeof(Alignment), true);
                    peo.AllowNone = true; // Cho phép nhấn Enter để kết thúc

                    PromptEntityResult per = A.Ed.GetEntity(peo);

                    if (per.Status != PromptStatus.OK)
                    {
                        // User nhấn Enter hoặc ESC → kết thúc
                        continueSelecting = false;
                        continue;
                    }

                    ObjectId pickedAlignmentId = per.ObjectId;

                    // Tìm và check alignment tương ứng trong danh sách
                    bool found = false;
                    for (int i = 0; i < alignmentDisplayInfos.Count; i++)
                    {
                        if (alignmentDisplayInfos[i].AlignmentId == pickedAlignmentId)
                        {
                            if (!clbAlignments.GetItemChecked(i))
                            {
                                clbAlignments.SetItemChecked(i, true);
                                pickedCount++;
                                A.Ed.WriteMessage($"\n  ✅ Đã chọn: {alignmentDisplayInfos[i].AlignmentName}");
                            }
                            else
                            {
                                A.Ed.WriteMessage($"\n  ⚠ Tuyến \"{alignmentDisplayInfos[i].AlignmentName}\" đã được chọn trước đó.");
                            }
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        // Alignment không nằm trong corridor đã chọn
                        A.Ed.WriteMessage("\n  ❌ Tuyến này không thuộc corridor đã chọn. Bỏ qua.");
                    }
                }

                A.Ed.WriteMessage($"\n══ Hoàn tất: đã chọn {pickedCount} tuyến trên bình đồ ══\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi khi chọn tuyến: {ex.Message}");
            }

            // Hiện lại form
            this.Show();
            OnAlignmentSelectionChanged();
        }

        // (Removed UI helpers)
    }
}
