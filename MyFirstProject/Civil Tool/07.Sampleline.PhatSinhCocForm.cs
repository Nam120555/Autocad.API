// (C) Copyright 2015 by  
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.Runtime;
using Acad = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using ATable = Autodesk.AutoCAD.DatabaseServices.Table;

using Civil = Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using Autodesk.Civil.Runtime;
using Autodesk.Civil.Settings;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Civil.ApplicationServices;
using CivSurface = Autodesk.Civil.DatabaseServices.TinSurface;
using Section = Autodesk.Civil.DatabaseServices.Section;
using Autodesk.Civil;
using System.Security.Cryptography.X509Certificates;
using System.Xml.Schema;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using MyFirstProject.Extensions;
//using Autodesk.Aec.DatabaseServices;

namespace MyFirstProject
{
    /// <summary>
    /// Form phát sinh cọc chuyên nghiệp kiểu Civil 3D
    /// Cho phép chọn tuyến, chọn loại cọc, xem preview trước khi tạo
    /// </summary>
    public class PhatSinhCocForm : CivilFormBase
    {
        // Alias giải quyết xung đột namespace
        // Dùng System.Drawing.Font thay vì Font (trùng Autodesk.AutoCAD.DatabaseServices.Font)
        // Dùng System.Drawing.Point thay vì Point (trùng Autodesk.Civil.DatabaseServices.Point)
        // ─── UI Controls ───────────────────────────────────────────
        private ComboBox cboAlignment = null!;
        private System.Windows.Forms.Label lblLength = null!;
        private System.Windows.Forms.Label lblGroupCount = null!;

        private CheckBox chkDacBiet = null!;
        private CheckBox chkHKm = null!;
        private NumericUpDown numKmStart = null!;
        private CheckBox chkChiTiet = null!;
        private NumericUpDown numKhoangCach = null!;
        private CheckBox chkLabel = null!;

        private DataGridView dgvPreview = null!;
        private System.Windows.Forms.Label lblTotal = null!;

        private Button btnPreview = null!;
        private Button btnCreate = null!;
        private Button btnCancel = null!;

        // ─── Data ──────────────────────────────────────────────────
        private List<AlignmentInfo> alignments = new();
        public bool IsConfirmed { get; private set; } = false;

        private class AlignmentInfo
        {
            public ObjectId Id { get; set; }
            public string Name { get; set; } = "";
            public double Length { get; set; }
            public int GroupCount { get; set; }
        }

        public class CocPreviewItem
        {
            public int STT { get; set; }
            public string TenCoc { get; set; } = "";
            public double Station { get; set; }
            public string LyTrinh { get; set; } = "";
            public string Loai { get; set; } = "";
        }

        private List<CocPreviewItem> previewItems = new();

        // ─── Constructor ────────────────────────────────────────────
        public PhatSinhCocForm()
        {
            LoadAlignments();
            InitializeComponent();
        }

        private void LoadAlignments()
        {
            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                CivilDocument cdoc = CivilApplication.ActiveDocument;
                ObjectIdCollection alignIds = cdoc.GetAlignmentIds();
                foreach (ObjectId id in alignIds)
                {
                    Alignment? al = tr.GetObject(id, OpenMode.ForRead) as Alignment;
                    if (al == null) continue;
                    alignments.Add(new AlignmentInfo
                    {
                        Id = id,
                        Name = al.Name,
                        Length = al.Length,
                        GroupCount = al.GetSampleLineGroupIds().Count
                    });
                }
                tr.Commit();
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
        }

        private void InitializeComponent()
        {
            // ─── Form ──────────────────────────────────────────
            InitBaseTheme("PHÁT SINH CỌC", 780, 720);
            this.FormBorderStyle = FormBorderStyle.Sizable;

            var header = CreateHeader("PHÁT SINH CỌC");
            this.Controls.Add(header);

            int pad = 15, y = 60;

            // ═══════════════════════════════════════════════════
            // SECTION 1: CHỌN TUYẾN
            // ═══════════════════════════════════════════════════
            var grpAlign = CreateSection("Chọn tuyến", pad, y, 740, 85);
            this.Controls.Add(grpAlign);

            grpAlign.Controls.Add(CreateLabel("Tuyến:", 12, 33));
            cboAlignment = new ComboBox
            {
                Location = new System.Drawing.Point(62, 30),
                Size = new Size(300, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextDark,
                FlatStyle = FlatStyle.Flat,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            foreach (var al in alignments)
                cboAlignment.Items.Add(al.Name);
            if (cboAlignment.Items.Count > 0) cboAlignment.SelectedIndex = 0;
            cboAlignment.SelectedIndexChanged += (s, e) => UpdateAlignmentInfo();
            grpAlign.Controls.Add(cboAlignment);

            lblLength = CreateLabel("Chiều dài: —", 380, 33);
            grpAlign.Controls.Add(lblLength);

            lblGroupCount = CreateLabel("Nhóm cọc: —", 560, 33);
            grpAlign.Controls.Add(lblGroupCount);

            // Hiển thị thông tin ban đầu
            if (alignments.Count > 0) UpdateAlignmentInfo();

            y += 95;

            // ═══════════════════════════════════════════════════
            // SECTION 2: TÙY CHỌN PHÁT SINH
            // ═══════════════════════════════════════════════════
            var grpOptions = CreateSection("Tùy chọn phát sinh cọc", pad, y, 740, 165);
            this.Controls.Add(grpOptions);

            // --- Row 1: Cọc đặc biệt ---
            chkDacBiet = new CheckBox
            {
                Text = "Cọc đặc biệt (điểm hình học: DT, TD, P, TC, CT...)",
                Location = new System.Drawing.Point(15, 33),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            grpOptions.Controls.Add(chkDacBiet);

            // --- Row 2: Cọc H/Km ---
            chkHKm = new CheckBox
            {
                Text = "Cọc H / Km (mỗi 100m)",
                Location = new System.Drawing.Point(15, 63),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            chkHKm.CheckedChanged += (s, e) => numKmStart.Enabled = chkHKm.Checked;
            grpOptions.Controls.Add(chkHKm);

            grpOptions.Controls.Add(CreateLabel("Km bắt đầu:", 400, 65));
            numKmStart = new NumericUpDown
            {
                Location = new System.Drawing.Point(490, 62),
                Size = new Size(80, 28),
                Minimum = 0,
                Maximum = 999,
                Value = 0,
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            grpOptions.Controls.Add(numKmStart);

            // --- Row 3: Cọc chi tiết ---
            chkChiTiet = new CheckBox
            {
                Text = "Cọc chi tiết (cọc phụ giữa các cọc chính)",
                Location = new System.Drawing.Point(15, 93),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            chkChiTiet.CheckedChanged += (s, e) => numKhoangCach.Enabled = chkChiTiet.Checked;
            grpOptions.Controls.Add(chkChiTiet);

            grpOptions.Controls.Add(CreateLabel("Khoảng cách (m):", 370, 95));
            numKhoangCach = new NumericUpDown
            {
                Location = new System.Drawing.Point(490, 92),
                Size = new Size(80, 28),
                Minimum = 5,
                Maximum = 100,
                Value = 20,
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            grpOptions.Controls.Add(numKhoangCach);

            // --- Row 4: Tạo label ---
            chkLabel = new CheckBox
            {
                Text = "Tạo Label tên cọc trên mặt bằng",
                Location = new System.Drawing.Point(15, 123),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            grpOptions.Controls.Add(chkLabel);

            y += 175;

            // ═══════════════════════════════════════════════════
            // SECTION 3: PREVIEW DANH SÁCH CỌC
            // ═══════════════════════════════════════════════════
            var grpPreview = CreateSection("Preview danh sách cọc", pad, y, 740, 310);
            this.Controls.Add(grpPreview);

            dgvPreview = new DataGridView
            {
                Location = new System.Drawing.Point(10, 30),
                Size = new Size(720, 240),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = BgInput,
                GridColor = BorderColor,
                BorderStyle = BorderStyle.FixedSingle,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = BgInput,
                    ForeColor = TextDark,
                    SelectionBackColor = NavyMedium,
                    SelectionForeColor = System.Drawing.Color.White,
                    Font = new System.Drawing.Font("Segoe UI", 9f)
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = NavyDark,
                    ForeColor = System.Drawing.Color.White,
                    Font = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 32
            };

            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "STT", Width = 50, Name = "colSTT" });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Tên cọc", Width = 140, Name = "colTen" });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Lý trình", Width = 130, Name = "colLyTrinh" });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Station (m)", Width = 110, Name = "colStation" });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Loại cọc", Width = 120, Name = "colLoai" });

            grpPreview.Controls.Add(dgvPreview);

            lblTotal = CreateLabel("Tổng: 0 cọc", 10, 280);
            lblTotal.Font = new System.Drawing.Font("Segoe UI", 10f, FontStyle.Bold);
            grpPreview.Controls.Add(lblTotal);

            // ═══════════════════════════════════════════════════
            // BUTTONS
            // ═══════════════════════════════════════════════════
            var pnlBottom = CreateBottomBar();
            this.Controls.Add(pnlBottom);

            btnPreview = CreateButton("Preview", 420, 8, 100, 36, pnlBottom);
            btnPreview.Click += (s, e) => GeneratePreview();

            btnCancel = CreateButton("Hủy", 530, 8, 100, 36, pnlBottom);
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            btnCreate = CreatePrimaryButton("Tạo cọc", 640, 8, 110, 36, pnlBottom);
            btnCreate.Click += (s, e) => { IsConfirmed = true; this.DialogResult = DialogResult.OK; this.Close(); };

            // Auto-preview khi mở form
            if (alignments.Count > 0) GeneratePreview();
        }

        // ─── Logic ──────────────────────────────────────────────
        private void UpdateAlignmentInfo()
        {
            if (cboAlignment.SelectedIndex < 0 || cboAlignment.SelectedIndex >= alignments.Count) return;
            var al = alignments[cboAlignment.SelectedIndex];
            lblLength.Text = $"Chiều dài: {al.Length:F2} m";
            lblGroupCount.Text = $"Nhóm cọc: {al.GroupCount}";
        }

        private AlignmentInfo? GetSelectedAlignment()
        {
            if (cboAlignment.SelectedIndex < 0 || cboAlignment.SelectedIndex >= alignments.Count) return null;
            return alignments[cboAlignment.SelectedIndex];
        }

        private string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double remainder = station - km * 1000;
            return $"{km}+{remainder:F3}";
        }

        /// <summary>
        /// Tính toán preview danh sách cọc (không tạo thực tế)
        /// </summary>
        private void GeneratePreview()
        {
            previewItems.Clear();
            dgvPreview.Rows.Clear();

            var alInfo = GetSelectedAlignment();
            if (alInfo == null) return;

            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                Alignment? alignment = tr.GetObject(alInfo.Id, OpenMode.ForRead) as Alignment;
                if (alignment == null) { tr.Commit(); return; }

                // Danh sách station + tên
                List<(double station, string name, string loai)> allCoc = new();

                // ── 1. Cọc đặc biệt ──
                if (chkDacBiet.Checked)
                {
                    Station[] stations = alignment.GetStationSet(StationTypes.GeometryPoint);
                    int curveIdx = 1;
                    for (int i = 0; i < stations.Length; i++)
                    {
                        string name = "";
                        var gpt = stations[i].GeometryStationType;

                        if (gpt == AlignmentGeometryPointStationType.BegOfAlign)
                            name = "DT";
                        else if (gpt == AlignmentGeometryPointStationType.EndOfAlign)
                            name = "CT";
                        else if (gpt == AlignmentGeometryPointStationType.TanTan)
                        { name = "D" + curveIdx; curveIdx++; }
                        else if (gpt == AlignmentGeometryPointStationType.TanCurve)
                        {
                            name = "TD" + curveIdx;
                            // Thêm P và TC nếu có
                            if (i + 1 < stations.Length)
                                allCoc.Add((stations[i + 1].RawStation, "P" + curveIdx, "Đặc biệt"));
                            if (i + 2 < stations.Length)
                                allCoc.Add((stations[i + 2].RawStation, "TC" + curveIdx, "Đặc biệt"));
                            curveIdx++;
                        }
                        else if (gpt == AlignmentGeometryPointStationType.TanSpiral)
                        {
                            name = "ND" + curveIdx;
                            if (i + 1 < stations.Length)
                                allCoc.Add((stations[i + 1].RawStation, "TD" + curveIdx, "Đặc biệt"));
                            if (i + 2 < stations.Length)
                                allCoc.Add((stations[i + 2].RawStation, "P" + curveIdx, "Đặc biệt"));
                            if (i + 3 < stations.Length)
                                allCoc.Add((stations[i + 3].RawStation, "TC" + curveIdx, "Đặc biệt"));
                            if (i + 4 < stations.Length)
                                allCoc.Add((stations[i + 4].RawStation, "NC" + curveIdx, "Đặc biệt"));
                            curveIdx++;
                        }

                        if (!string.IsNullOrEmpty(name))
                            allCoc.Add((stations[i].RawStation, name, "Đặc biệt"));
                    }
                }

                // ── 2. Cọc H/Km ──
                if (chkHKm.Checked)
                {
                    int kmStart = (int)numKmStart.Value;
                    int hNum = 1;
                    int currentKm = kmStart;

                    for (double st = 100; st < alignment.Length; st += 100)
                    {
                        if (st % 1000 == 0)
                        {
                            currentKm = (int)(st / 1000) + kmStart;
                            allCoc.Add((st, $"Km{currentKm}", "H/Km"));
                            hNum = 1;
                        }
                        else
                        {
                            allCoc.Add((st, $"H{hNum} (Km{currentKm})", "H/Km"));
                            hNum++;
                        }
                    }
                }

                // ── 3. Cọc chi tiết ──
                if (chkChiTiet.Checked)
                {
                    // Lấy tất cả station đã có → sắp xếp → chèn chi tiết
                    var existingStations = allCoc.Select(c => c.station).Distinct().OrderBy(s => s).ToList();
                    // Thêm đầu và cuối nếu chưa có
                    if (existingStations.Count == 0 || existingStations[0] > 0.01)
                        existingStations.Insert(0, 0);
                    if (existingStations.Count == 0 || existingStations.Last() < alignment.Length - 0.01)
                        existingStations.Add(alignment.Length);

                    int khoangCach = (int)numKhoangCach.Value;
                    int detailNum = 1;
                    for (int i = 0; i < existingStations.Count - 1; i++)
                    {
                        double start = existingStations[i];
                        double end = existingStations[i + 1];
                        int segments = (int)Math.Ceiling((end - start) / khoangCach);
                        for (int j = 1; j < segments; j++)
                        {
                            double st = start + j * khoangCach;
                            if (st < end - 0.5) // tránh trùng điểm cuối
                            {
                                allCoc.Add((st, detailNum.ToString(), "Chi tiết"));
                                detailNum++;
                            }
                        }
                    }
                }

                // ── Sắp xếp và đánh số ──
                allCoc = allCoc.OrderBy(c => c.station).ThenByDescending(c => c.loai == "Đặc biệt" ? 2 : c.loai == "H/Km" ? 1 : 0).ToList();

                // Loại bỏ trùng station (giữ cọc ưu tiên cao hơn)
                var deduped = new List<(double station, string name, string loai)>();
                foreach (var coc in allCoc)
                {
                    if (deduped.Count == 0 || Math.Abs(deduped.Last().station - coc.station) > 0.01)
                        deduped.Add(coc);
                }

                // Build preview items
                int stt = 1;
                foreach (var coc in deduped)
                {
                    previewItems.Add(new CocPreviewItem
                    {
                        STT = stt++,
                        TenCoc = coc.name,
                        Station = coc.station,
                        LyTrinh = FormatStation(coc.station),
                        Loai = coc.loai
                    });
                }

                // ── Cập nhật DataGridView ──
                foreach (var item in previewItems)
                {
                    int idx = dgvPreview.Rows.Add(item.STT, item.TenCoc, item.LyTrinh, $"{item.Station:F3}", item.Loai);

                    // Tô màu theo loại cọc — chỉ bold cọc đặc biệt
                    var row = dgvPreview.Rows[idx];
                    if (item.Loai == "Đặc biệt")
                        row.DefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Bold);
                    else if (item.Loai == "H/Km")
                        row.DefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Italic);
                    // Chi tiết: giữ mặc định
                }

                lblTotal.Text = $"Tổng: {previewItems.Count} cọc  |  " +
                    $"Đặc biệt: {previewItems.Count(c => c.Loai == "Đặc biệt")}  |  " +
                    $"H/Km: {previewItems.Count(c => c.Loai == "H/Km")}  |  " +
                    $"Chi tiết: {previewItems.Count(c => c.Loai == "Chi tiết")}";

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                lblTotal.Text = $"Lỗi: {ex.Message}";
                lblTotal.ForeColor = Color.FromArgb(255, 80, 80);
            }
        }

        // ─── Public API ─────────────────────────────────────────
        public ObjectId GetAlignmentId()
        {
            var al = GetSelectedAlignment();
            return al?.Id ?? ObjectId.Null;
        }

        public bool CreateDacBiet => chkDacBiet.Checked;
        public bool CreateHKm => chkHKm.Checked;
        public int KmStart => (int)numKmStart.Value;
        public bool CreateChiTiet => chkChiTiet.Checked;
        public int KhoangCach => (int)numKhoangCach.Value;
        public bool ShouldCreateLabel => chkLabel.Checked;
        public List<CocPreviewItem> PreviewItems => previewItems;
    }
}
