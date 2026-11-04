// 100.LapCongThucCivil.cs
// Lệnh: CTSV_LapCTTinhTracNgang
// Chức năng: Lập công thức tính khối lượng trắc ngang
//   - Chọn bề mặt trên / bề mặt dưới → tính DT mặt cắt
//   - Tính KL theo Average End Area: V = (S₁ + S₂)/2 × L
//   - Bố trí bảng KL lên SectionView
//   - Xuất Excel

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using ClosedXML.Excel;
using MyFirstProject.Extensions;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadDb = Autodesk.AutoCAD.DatabaseServices;
using CivSection = Autodesk.Civil.DatabaseServices.Section;
using ATable = Autodesk.AutoCAD.DatabaseServices.Table;

using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using ContentAlignment = System.Drawing.ContentAlignment;

[assembly: CommandClass(typeof(MyFirstProject.LapCongThucCivil))]

namespace MyFirstProject
{
    // ══════════════════════════════════════════════════════════════
    // DATA CLASSES
    // ══════════════════════════════════════════════════════════════

    /// <summary>
    /// Một công thức tính KL: Bề mặt trên − Bề mặt dưới
    /// </summary>
    public class CongThucKL
    {
        public string TenVatLieu { get; set; } = "";
        public ObjectId SurfaceAboveId { get; set; }
        public string SurfaceAboveName { get; set; } = "";
        public ObjectId SurfaceBelowId { get; set; }
        public string SurfaceBelowName { get; set; } = "";
    }

    /// <summary>
    /// Thông tin 1 bề mặt có thể chọn (từ SectionSource)
    /// </summary>
    public class SurfaceInfo
    {
        public ObjectId SourceId { get; set; }
        public string Name { get; set; } = "";
        public SectionSourceType SourceType { get; set; }

        public override string ToString() => Name;
    }

    /// <summary>
    /// Kết quả tại 1 trắc ngang
    /// </summary>
    public class KetQuaTracNgang
    {
        public string TenCoc { get; set; } = "";
        public double Station { get; set; }
        public string LyTrinh { get; set; } = "";
        public double KhoangCach { get; set; }

        /// <summary>Diện tích theo từng công thức (m²)</summary>
        public Dictionary<string, double> DienTich { get; set; } = new();

        /// <summary>Khối lượng Average End Area (m³)</summary>
        public Dictionary<string, double> KhoiLuong { get; set; } = new();

        /// <summary>Khối lượng tích lũy (m³)</summary>
        public Dictionary<string, double> KhoiLuongTL { get; set; } = new();
    }

    // ══════════════════════════════════════════════════════════════
    // FORM: LẬP CÔNG THỨC
    // ══════════════════════════════════════════════════════════════

    public class FormLapCongThuc : CivilFormBase
    {
        // ── Controls ──
        private ComboBox cboAlignment = null!;
        private ComboBox cboSLG = null!;
        private DataGridView dgvCongThuc = null!;
        private Button btnThemCT = null!;
        private Button btnXoaCT = null!;
        private CheckBox chkExcel = null!;
        private CheckBox chkBoTriBang = null!;

        // ── Data ──
        private List<(string Name, ObjectId AlignmentId, ObjectId SLGId)> alignmentList = new();
        private List<SurfaceInfo> surfaceList = new();
        private List<CongThucKL> congThucList = new();

        // ── Kết quả ──
        public bool IsConfirmed { get; private set; }
        public ObjectId SelectedAlignmentId { get; private set; }
        public ObjectId SelectedSLGId { get; private set; }
        public List<CongThucKL> CongThucs => congThucList;
        public bool XuatExcel => chkExcel.Checked;
        public bool BoTriBang => chkBoTriBang.Checked;

        public FormLapCongThuc()
        {
            LoadAlignments();
            InitializeUI();
        }

        private void LoadAlignments()
        {
            try
            {
                using var tr = A.Db.TransactionManager.StartOpenCloseTransaction();
                CivilDocument cdoc = CivilApplication.ActiveDocument;

                foreach (ObjectId alId in cdoc.GetAlignmentIds())
                {
                    var al = tr.GetObject(alId, OpenMode.ForRead) as Alignment;
                    if (al == null) continue;

                    foreach (ObjectId slgId in al.GetSampleLineGroupIds())
                    {
                        var slg = tr.GetObject(slgId, OpenMode.ForRead) as SampleLineGroup;
                        if (slg != null)
                            alignmentList.Add(($"{al.Name} — {slg.Name}", alId, slgId));
                    }
                }
                tr.Commit();
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
        }

        private void LoadSurfaces()
        {
            surfaceList.Clear();
            if (cboAlignment.SelectedIndex < 0) return;

            var slgId = alignmentList[cboAlignment.SelectedIndex].SLGId;

            try
            {
                using var tr = A.Db.TransactionManager.StartOpenCloseTransaction();
                var slg = tr.GetObject(slgId, OpenMode.ForRead) as SampleLineGroup;
                if (slg == null) { tr.Commit(); return; }

                SectionSourceCollection sources = slg.GetSectionSources();
                foreach (SectionSource src in sources)
                {
                    if (!src.IsSampled) continue;

                    string name = "";
                    try
                    {
                        var ent = tr.GetObject(src.SourceId, OpenMode.ForRead);
                        if (ent is TinSurface ts) name = ts.Name;
                        else if (ent is Corridor cor) name = cor.Name;
                        else name = ent.GetType().Name;
                    }
                    catch { name = $"Source_{src.SourceId}"; }

                    string typeTag = src.SourceType switch
                    {
                        SectionSourceType.TinSurface => "[Mặt đất]",
                        SectionSourceType.CorridorSurface => "[Corridor Srf]",
                        SectionSourceType.Corridor => "[Corridor]",
                        _ => "[Khác]"
                    };

                    surfaceList.Add(new SurfaceInfo
                    {
                        SourceId = src.SourceId,
                        Name = $"{name} {typeTag}",
                        SourceType = src.SourceType
                    });
                }

                tr.Commit();
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

            // Cập nhật ComboBox columns trong DataGridView
            UpdateSurfaceColumns();
        }

        private void UpdateSurfaceColumns()
        {
            var colAbove = dgvCongThuc.Columns["SurfaceAbove"] as DataGridViewComboBoxColumn;
            var colBelow = dgvCongThuc.Columns["SurfaceBelow"] as DataGridViewComboBoxColumn;
            if (colAbove == null || colBelow == null) return;

            var names = surfaceList.Select(s => s.Name).ToArray();
            colAbove.Items.Clear();
            colBelow.Items.Clear();
            foreach (var n in names) { colAbove.Items.Add(n); colBelow.Items.Add(n); }
        }

        private void InitializeUI()
        {
            InitBaseTheme("LẬP CÔNG THỨC TÍNH KL TRẮC NGANG", 750, 600);
            this.FormBorderStyle = FormBorderStyle.Sizable;

            // ═══ HEADER ═══
            var header = CreateHeader("LẬP CÔNG THỨC TÍNH KHỐI LƯỢNG TRẮC NGANG");
            this.Controls.Add(header);

            var subtitle = CreateSubtitle("Chọn bề mặt TRÊN − bề mặt DƯỚI → Diện tích → Khối lượng (Average End Area)");
            this.Controls.Add(subtitle);

            // ═══ MAIN PANEL ═══
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15),
                AutoScroll = true
            };
            this.Controls.Add(mainPanel);
            mainPanel.BringToFront();

            int y = 5;

            // ═══ CHỌN ALIGNMENT ═══
            var grpAlign = CreateGroupBox("Chọn Alignment + SampleLine Group", 0, y, 700, 70);
            mainPanel.Controls.Add(grpAlign);

            grpAlign.Controls.Add(CreateLabel("Alignment:", 12, 32));
            cboAlignment = CreateComboBox(120, 28, 560);
            foreach (var (name, _, _) in alignmentList)
                cboAlignment.Items.Add(name);
            if (cboAlignment.Items.Count > 0) cboAlignment.SelectedIndex = 0;
            cboAlignment.SelectedIndexChanged += (s, e) => LoadSurfaces();
            grpAlign.Controls.Add(cboAlignment);

            y += 80;

            // ═══ CÔNG THỨC ═══
            var grpCT = CreateGroupBox("Công thức tính KL (Bề mặt trên − Bề mặt dưới = Diện tích)", 0, y, 700, 280);
            mainPanel.Controls.Add(grpCT);

            // DataGridView cho công thức
            dgvCongThuc = new DataGridView
            {
                Location = new Point(10, 30),
                Size = new Size(580, 200),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = BgWhite,
                GridColor = BorderColor,
                BorderStyle = BorderStyle.FixedSingle,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = BgWhite,
                    ForeColor = TextDark,
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
                EditMode = DataGridViewEditMode.EditOnEnter
            };

            // Cột tên vật liệu (editable)
            var colName = new DataGridViewTextBoxColumn
            {
                Name = "TenVL",
                HeaderText = "Tên Vật Liệu",
                Width = 150,
                ReadOnly = false
            };
            dgvCongThuc.Columns.Add(colName);

            // Cột bề mặt trên (ComboBox)
            var colAbove = new DataGridViewComboBoxColumn
            {
                Name = "SurfaceAbove",
                HeaderText = "Bề mặt TRÊN",
                Width = 200,
                FlatStyle = FlatStyle.Flat
            };
            dgvCongThuc.Columns.Add(colAbove);

            // Cột bề mặt dưới (ComboBox)
            var colBelow = new DataGridViewComboBoxColumn
            {
                Name = "SurfaceBelow",
                HeaderText = "Bề mặt DƯỚI",
                Width = 200,
                FlatStyle = FlatStyle.Flat
            };
            dgvCongThuc.Columns.Add(colBelow);

            grpCT.Controls.Add(dgvCongThuc);

            // Buttons thêm/xóa
            btnThemCT = CreateButton("＋ Thêm CT", 600, 30, 90, 34, grpCT);
            btnThemCT.Click += BtnThemCT_Click;

            btnXoaCT = CreateButton("－ Xóa", 600, 70, 90, 34, grpCT);
            btnXoaCT.Click += BtnXoaCT_Click;

            // Hint
            var hint = CreateHintLabel(
                "Ví dụ: Đào nền = mặt TN (trên) − mặt Datum (dưới) → DT phần TN cao hơn Datum",
                10, 240);
            grpCT.Controls.Add(hint);

            var hint2 = CreateHintLabel(
                "       Đắp nền = mặt Datum (trên) − mặt TN (dưới) → DT phần Datum cao hơn TN",
                10, 258);
            grpCT.Controls.Add(hint2);

            y += 290;

            // ═══ TÙY CHỌN ═══
            chkExcel = CreateCheckBox("Xuất bảng tổng hợp ra Excel (.xlsx)", 15, y, true);
            mainPanel.Controls.Add(chkExcel);
            y += 28;

            chkBoTriBang = CreateCheckBox("Bố trí bảng KL lên SectionView", 15, y, false);
            mainPanel.Controls.Add(chkBoTriBang);

            // ═══ BOTTOM BAR ═══
            var bottomBar = CreateBottomBar();
            this.Controls.Add(bottomBar);

            var btnCancel = CreateButton("Hủy", 480, 8, 100, 36, bottomBar);
            btnCancel.Click += (s, e) => { IsConfirmed = false; this.Close(); };

            var btnRun = CreateSuccessButton("Tính Khối Lượng", 590, 8, 140, 36, bottomBar);
            btnRun.Click += BtnRun_Click;

            this.AcceptButton = btnRun;
            this.CancelButton = btnCancel;

            // Load surfaces ban đầu
            if (cboAlignment.Items.Count > 0) LoadSurfaces();
        }

        private void BtnThemCT_Click(object? sender, EventArgs e)
        {
            if (surfaceList.Count < 2)
            {
                MessageBox.Show("Cần ít nhất 2 bề mặt trong SampleLineGroup!",
                    "Thiếu dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int idx = dgvCongThuc.Rows.Add();
            dgvCongThuc.Rows[idx].Cells["TenVL"].Value = $"Vật liệu {idx + 1}";
        }

        private void BtnXoaCT_Click(object? sender, EventArgs e)
        {
            if (dgvCongThuc.SelectedRows.Count > 0)
                dgvCongThuc.Rows.Remove(dgvCongThuc.SelectedRows[0]);
        }

        private void BtnRun_Click(object? sender, EventArgs e)
        {
            // Validate
            if (cboAlignment.SelectedIndex < 0)
            {
                MessageBox.Show("Chọn Alignment!", "Thiếu dữ liệu",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (dgvCongThuc.Rows.Count == 0)
            {
                MessageBox.Show("Thêm ít nhất 1 công thức!", "Thiếu dữ liệu",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Thu thập công thức
            congThucList.Clear();
            for (int i = 0; i < dgvCongThuc.Rows.Count; i++)
            {
                var row = dgvCongThuc.Rows[i];
                string tenVL = row.Cells["TenVL"].Value?.ToString() ?? $"VL_{i + 1}";
                string? aboveName = row.Cells["SurfaceAbove"].Value?.ToString();
                string? belowName = row.Cells["SurfaceBelow"].Value?.ToString();

                if (string.IsNullOrEmpty(aboveName) || string.IsNullOrEmpty(belowName))
                {
                    MessageBox.Show($"Hàng {i + 1}: Chưa chọn bề mặt!",
                        "Thiếu dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var surfAbove = surfaceList.FirstOrDefault(s => s.Name == aboveName);
                var surfBelow = surfaceList.FirstOrDefault(s => s.Name == belowName);

                if (surfAbove == null || surfBelow == null)
                {
                    MessageBox.Show($"Hàng {i + 1}: Bề mặt không hợp lệ!",
                        "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                congThucList.Add(new CongThucKL
                {
                    TenVatLieu = tenVL,
                    SurfaceAboveId = surfAbove.SourceId,
                    SurfaceAboveName = surfAbove.Name,
                    SurfaceBelowId = surfBelow.SourceId,
                    SurfaceBelowName = surfBelow.Name
                });
            }

            SelectedAlignmentId = alignmentList[cboAlignment.SelectedIndex].AlignmentId;
            SelectedSLGId = alignmentList[cboAlignment.SelectedIndex].SLGId;
            IsConfirmed = true;
            this.Close();
        }
    }

    // ══════════════════════════════════════════════════════════════
    // LỆNH CHÍNH + ENGINE TÍNH TOÁN
    // ══════════════════════════════════════════════════════════════

    public class LapCongThucCivil
    {
        /// <summary>
        /// Format lý trình dạng Km0+123.456
        /// </summary>
        private static string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double m = station % 1000;
            return $"Km{km}+{m:F3}";
        }

        // ═══════════════════════════════════════════════════════════
        // LỆNH: CTSV_LapCTTinhTracNgang
        // ═══════════════════════════════════════════════════════════

        [CommandMethod("CTSV_LapCTTinhTracNgang")]
        public static void CTSVLapCTTinhTracNgang()
        {
            var ed = A.Ed;

            ed.WriteMessage("\n╔════════════════════════════════════════════════════════════════╗");
            ed.WriteMessage("\n║  LẬP CÔNG THỨC TÍNH KHỐI LƯỢNG TRẮC NGANG                     ║");
            ed.WriteMessage("\n║  Bề mặt TRÊN − Bề mặt DƯỚI = DT → KL (Average End Area)      ║");
            ed.WriteMessage("\n╚════════════════════════════════════════════════════════════════╝");

            // 1. Mở form lập công thức
            using var form = new FormLapCongThuc();
            AcadApp.ShowModalDialog(form);

            if (!form.IsConfirmed || form.CongThucs.Count == 0) return;

            var congThucs = form.CongThucs;
            var alignmentId = form.SelectedAlignmentId;
            var slgId = form.SelectedSLGId;
            bool xuatExcel = form.XuatExcel;
            bool boTriBang = form.BoTriBang;

            try
            {
                using var tr = A.Db.TransactionManager.StartTransaction();

                var alignment = tr.GetObject(alignmentId, OpenMode.ForWrite) as Alignment;
                var slg = tr.GetObject(slgId, OpenMode.ForWrite) as SampleLineGroup;
                if (alignment == null || slg == null) { tr.Commit(); return; }

                // 2. Lấy SampleLines sắp xếp theo station
                var sampleLines = new List<SampleLine>();
                foreach (ObjectId slId in slg.GetSampleLineIds())
                {
                    var sl = tr.GetObject(slId, OpenMode.ForWrite) as SampleLine;
                    if (sl != null) sampleLines.Add(sl);
                }
                sampleLines = sampleLines.OrderBy(s => s.Station).ToList();

                if (sampleLines.Count == 0)
                {
                    ed.WriteMessage("\n⚠️ Không có SampleLine nào!");
                    tr.Commit();
                    return;
                }

                ed.WriteMessage($"\n\n📐 Alignment: {alignment.Name}");
                ed.WriteMessage($"\n📋 Số trắc ngang: {sampleLines.Count}");
                ed.WriteMessage($"\n📝 Số công thức: {congThucs.Count}");

                // In header công thức
                foreach (var ct in congThucs)
                {
                    ed.WriteMessage($"\n   → {ct.TenVatLieu}: {ct.SurfaceAboveName} − {ct.SurfaceBelowName}");
                }

                // 3. Tính DT tại mỗi trắc ngang
                var results = new List<KetQuaTracNgang>();
                double prevStation = 0;
                bool isFirst = true;

                foreach (var sl in sampleLines)
                {
                    try
                    {
                        var result = new KetQuaTracNgang
                        {
                            TenCoc = sl.Name ?? $"SL_{sl.Station:F0}",
                            Station = sl.Station,
                            LyTrinh = FormatStation(sl.Station),
                            KhoangCach = isFirst ? 0 : sl.Station - prevStation
                        };

                        foreach (var ct in congThucs)
                        {
                            double area = TinhDienTichGiua2BeMat(tr, sl, ct.SurfaceAboveId, ct.SurfaceBelowId);
                            result.DienTich[ct.TenVatLieu] = Math.Round(area, 4);
                        }

                        results.Add(result);
                        prevStation = sl.Station;
                        isFirst = false;
                    }
                    catch (System.Exception exSL)
                    {
                        ed.WriteMessage($"\n⚠️ Bỏ qua SL '{sl.Name}': {exSL.Message}");
                    }
                }

                if (results.Count == 0)
                {
                    ed.WriteMessage("\n⚠️ Không tính được diện tích tại bất kỳ trắc ngang nào!");
                    tr.Commit();
                    return;
                }

                // 4. Tính KL Average End Area
                TinhKhoiLuongAEA(results, congThucs);

                // 5. Hiển thị kết quả
                HienThiKetQua(ed, alignment.Name, results, congThucs);

                // 6. Copy clipboard (wrap try-catch riêng, không để crash cả lệnh)
                try { CopyClipboard(alignment.Name, results, congThucs); }
                catch { ed.WriteMessage("\n⚠️ Không copy được clipboard."); }

                // 7. Xuất Excel (phải commit transaction trước khi mở dialog)
                //    Vì SaveFileDialog cần chạy ngoài transaction
                bool needExcel = xuatExcel;
                string alignNameForExcel = alignment.Name;

                // 8. Bố trí bảng lên SectionView (trong transaction)
                if (boTriBang)
                {
                    try { BoTriBangLenSV(tr, ed, slg, sampleLines, results, congThucs); }
                    catch (System.Exception exBT)
                    {
                        ed.WriteMessage($"\n⚠️ Lỗi bố trí bảng: {exBT.Message}");
                    }
                }

                tr.Commit();

                // Xuất Excel SAU khi commit transaction (tránh crash do dialog)
                if (needExcel)
                {
                    try { XuatExcel(ed, alignNameForExcel, results, congThucs); }
                    catch (System.Exception exXL)
                    {
                        ed.WriteMessage($"\n⚠️ Lỗi xuất Excel: {exXL.Message}");
                    }
                }

                ed.WriteMessage("\n\n✅ HOÀN THÀNH tính khối lượng trắc ngang!");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
                ed.WriteMessage($"\n   {ex.StackTrace}");
            }
        }

        // ═══════════════════════════════════════════════════════════
        // TÍNH DIỆN TÍCH GIỮA 2 BỀ MẶT TẠI 1 TRẮC NGANG
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// Tính diện tích phần bề mặt trên CAO HƠN bề mặt dưới
        /// tại một SampleLine.
        /// Phương pháp: Lấy SectionPoints của 2 section, nội suy
        /// tại cùng vị trí offset, tính tích phân (Trapezoid).
        /// </summary>
        private static double TinhDienTichGiua2BeMat(
            Transaction tr, SampleLine sampleLine,
            ObjectId surfaceAboveId, ObjectId surfaceBelowId)
        {
            try
            {
                // Kiểm tra ObjectId hợp lệ trước khi truy cập
                if (surfaceAboveId.IsNull || !surfaceAboveId.IsValid ||
                    surfaceBelowId.IsNull || !surfaceBelowId.IsValid)
                    return 0;

                // Lấy Section cho bề mặt trên
                ObjectId secAboveId = ObjectId.Null;
                try { secAboveId = sampleLine.GetSectionId(surfaceAboveId); }
                catch { return 0; }
                if (secAboveId.IsNull || !secAboveId.IsValid) return 0;

                CivSection? secAbove = tr.GetObject(secAboveId, OpenMode.ForWrite) as CivSection;
                if (secAbove == null) return 0;

                // Lấy Section cho bề mặt dưới
                ObjectId secBelowId = ObjectId.Null;
                try { secBelowId = sampleLine.GetSectionId(surfaceBelowId); }
                catch { return 0; }
                if (secBelowId.IsNull || !secBelowId.IsValid) return 0;

                CivSection? secBelow = tr.GetObject(secBelowId, OpenMode.ForWrite) as CivSection;
                if (secBelow == null) return 0;

                // Chuyển SectionPoints thành list (offset, elevation)
                var profileAbove = GetProfile(secAbove);
                var profileBelow = GetProfile(secBelow);

                if (profileAbove.Count < 2 || profileBelow.Count < 2) return 0;

                // Tính diện tích giữa 2 profile
                return TinhDienTichGiua2Profile(profileAbove, profileBelow);
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Chuyển SectionPoints → List (offset, elevation)
        /// </summary>
        private static List<(double offset, double elev)> GetProfile(CivSection section)
        {
            var pts = new List<(double, double)>();
            try
            {
                foreach (SectionPoint pt in section.SectionPoints)
                    pts.Add((pt.Location.X, pt.Location.Y));
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
            return pts.OrderBy(p => p.Item1).ToList();
        }

        /// <summary>
        /// Tính diện tích phần profile_above CAO HƠN profile_below.
        /// Nội suy 2 profile tại cùng tập offset, rồi tích phân trapezoid.
        /// Chỉ tính vùng có above > below (phần dương).
        /// </summary>
        private static double TinhDienTichGiua2Profile(
            List<(double offset, double elev)> above,
            List<(double offset, double elev)> below)
        {
            // 1. Gom tất cả offset (X) từ cả 2 profile
            var allX = new SortedSet<double>();
            foreach (var p in above) allX.Add(p.offset);
            foreach (var p in below) allX.Add(p.offset);

            // 2. Tìm phạm vi offset chung
            double xMin = Math.Max(above.Min(p => p.offset), below.Min(p => p.offset));
            double xMax = Math.Min(above.Max(p => p.offset), below.Max(p => p.offset));

            var xList = allX.Where(x => x >= xMin && x <= xMax).ToList();
            if (xList.Count < 2) return 0;

            // 3. Tích phân trapezoid — chỉ lấy phần dương (above > below)
            double totalArea = 0;
            for (int i = 0; i < xList.Count - 1; i++)
            {
                double x1 = xList[i];
                double x2 = xList[i + 1];
                double dx = x2 - x1;
                if (dx <= 0) continue;

                double yA1 = Interpolate(above, x1);
                double yA2 = Interpolate(above, x2);
                double yB1 = Interpolate(below, x1);
                double yB2 = Interpolate(below, x2);

                double diff1 = yA1 - yB1; // dương = above cao hơn
                double diff2 = yA2 - yB2;

                // Xử lý 4 trường hợp
                if (diff1 >= 0 && diff2 >= 0)
                {
                    // Cả 2 đầu đều above > below → trapezoid đầy đủ
                    totalArea += (diff1 + diff2) / 2.0 * dx;
                }
                else if (diff1 >= 0 && diff2 < 0)
                {
                    // Giao cắt trong đoạn → chỉ lấy phần dương
                    double ratio = diff1 / (diff1 - diff2);
                    double dxPos = dx * ratio;
                    totalArea += diff1 / 2.0 * dxPos; // tam giác
                }
                else if (diff1 < 0 && diff2 >= 0)
                {
                    // Giao cắt trong đoạn → chỉ lấy phần dương
                    double ratio = -diff1 / (diff2 - diff1);
                    double dxPos = dx * (1 - ratio);
                    totalArea += diff2 / 2.0 * dxPos; // tam giác
                }
                // else: cả 2 đều âm → bỏ qua (below cao hơn above)
            }

            return totalArea;
        }

        /// <summary>
        /// Nội suy elevation tại offset x trên 1 profile
        /// </summary>
        private static double Interpolate(List<(double offset, double elev)> profile, double x)
        {
            if (profile.Count == 0) return 0;
            if (x <= profile[0].offset) return profile[0].elev;
            if (x >= profile[^1].offset) return profile[^1].elev;

            for (int i = 0; i < profile.Count - 1; i++)
            {
                if (x >= profile[i].offset && x <= profile[i + 1].offset)
                {
                    double dx = profile[i + 1].offset - profile[i].offset;
                    if (dx == 0) return profile[i].elev;
                    double t = (x - profile[i].offset) / dx;
                    return profile[i].elev + t * (profile[i + 1].elev - profile[i].elev);
                }
            }
            return profile[^1].elev;
        }

        // ═══════════════════════════════════════════════════════════
        // TÍNH KHỐI LƯỢNG AVERAGE END AREA
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// V = (S₁ + S₂) / 2 × L
        /// </summary>
        private static void TinhKhoiLuongAEA(
            List<KetQuaTracNgang> results,
            List<CongThucKL> congThucs)
        {
            foreach (var ct in congThucs)
            {
                string key = ct.TenVatLieu;
                double cumulative = 0;

                for (int i = 0; i < results.Count; i++)
                {
                    if (i == 0)
                    {
                        results[i].KhoiLuong[key] = 0;
                    }
                    else
                    {
                        double s1 = results[i - 1].DienTich.GetValueOrDefault(key, 0);
                        double s2 = results[i].DienTich.GetValueOrDefault(key, 0);
                        double L = results[i].KhoangCach;
                        double vol = (s1 + s2) / 2.0 * L;
                        results[i].KhoiLuong[key] = Math.Round(vol, 4);
                    }

                    cumulative += results[i].KhoiLuong.GetValueOrDefault(key, 0);
                    results[i].KhoiLuongTL[key] = Math.Round(cumulative, 4);
                }
            }
        }

        // ═══════════════════════════════════════════════════════════
        // HIỂN THỊ KẾT QUẢ
        // ═══════════════════════════════════════════════════════════

        private static void HienThiKetQua(
            Editor ed, string alignName,
            List<KetQuaTracNgang> results,
            List<CongThucKL> congThucs)
        {
            ed.WriteMessage($"\n\n═══ KẾT QUẢ: {alignName} ═══");
            ed.WriteMessage($"\n{"STT",-5}{"Tên Cọc",-18}{"Lý Trình",-15}{"KC",-8}");
            foreach (var ct in congThucs)
                ed.WriteMessage($"{"DT " + ct.TenVatLieu,-14}{"KL " + ct.TenVatLieu,-14}");
            ed.WriteMessage("\n" + new string('─', 50 + congThucs.Count * 28));

            int stt = 1;
            foreach (var r in results)
            {
                ed.WriteMessage($"\n{stt,-5}{r.TenCoc,-18}{r.LyTrinh,-15}{r.KhoangCach,-8:F1}");
                foreach (var ct in congThucs)
                {
                    double dt = r.DienTich.GetValueOrDefault(ct.TenVatLieu, 0);
                    double kl = r.KhoiLuong.GetValueOrDefault(ct.TenVatLieu, 0);
                    ed.WriteMessage($"{dt,-14:F4}{kl,-14:F4}");
                }
                stt++;
            }

            // Tổng
            ed.WriteMessage("\n" + new string('═', 50 + congThucs.Count * 28));
            ed.WriteMessage($"\n{"TỔNG",-46}");
            foreach (var ct in congThucs)
            {
                double tongKL = results.Sum(r => r.KhoiLuong.GetValueOrDefault(ct.TenVatLieu, 0));
                ed.WriteMessage($"{"",14}{tongKL,-14:F2}");
            }
            ed.WriteMessage("\n");
        }

        // ═══════════════════════════════════════════════════════════
        // COPY CLIPBOARD
        // ═══════════════════════════════════════════════════════════

        private static void CopyClipboard(
            string alignName,
            List<KetQuaTracNgang> results,
            List<CongThucKL> congThucs)
        {
            try
            {
                var sb = new StringBuilder();
                sb.Append("STT\tTên Cọc\tLý Trình\tKhoảng cách (m)");
                foreach (var ct in congThucs)
                    sb.Append($"\tDT {ct.TenVatLieu} (m²)\tKL {ct.TenVatLieu} (m³)\tKL TL {ct.TenVatLieu} (m³)");
                sb.AppendLine();

                int stt = 1;
                foreach (var r in results)
                {
                    sb.Append($"{stt}\t{r.TenCoc}\t{r.LyTrinh}\t{r.KhoangCach:F3}");
                    foreach (var ct in congThucs)
                    {
                        double dt = r.DienTich.GetValueOrDefault(ct.TenVatLieu, 0);
                        double kl = r.KhoiLuong.GetValueOrDefault(ct.TenVatLieu, 0);
                        double klTL = r.KhoiLuongTL.GetValueOrDefault(ct.TenVatLieu, 0);
                        sb.Append($"\t{dt:F4}\t{kl:F4}\t{klTL:F4}");
                    }
                    sb.AppendLine();
                    stt++;
                }

                var thread = new System.Threading.Thread(() =>
                {
                    try { Clipboard.SetText(sb.ToString()); } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                });
                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                thread.Start();
                thread.Join(2000);
                A.Ed.WriteMessage("\n📋 Đã copy vào Clipboard! Dán (Ctrl+V) vào Excel.");
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
        }

        // ═══════════════════════════════════════════════════════════
        // XUẤT EXCEL
        // ═══════════════════════════════════════════════════════════

        private static void XuatExcel(
            Editor ed, string alignName,
            List<KetQuaTracNgang> results,
            List<CongThucKL> congThucs)
        {
            try
            {
                string? savePath = null;
                using (SaveFileDialog sfd = new()
                {
                    Filter = "Excel Files|*.xlsx",
                    FileName = $"KL_TracNgang_{alignName.Replace(" ", "_")}.xlsx",
                    Title = "Lưu file Excel"
                })
                {
                    // Hiển thị dialog trên STA thread để tránh crash AutoCAD
                    if (sfd.ShowDialog() == DialogResult.OK)
                        savePath = sfd.FileName;
                }
                if (string.IsNullOrEmpty(savePath)) return;

                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("KL Trắc Ngang");

                // Title
                ws.Cell(1, 1).Value = $"BẢNG TÍNH KHỐI LƯỢNG TRẮC NGANG - {alignName}";
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Font.FontSize = 14;

                // Công thức đã dùng
                int row = 2;
                foreach (var ct in congThucs)
                {
                    ws.Cell(row, 1).Value = $"Công thức: {ct.TenVatLieu} = {ct.SurfaceAboveName} − {ct.SurfaceBelowName}";
                    ws.Cell(row, 1).Style.Font.Italic = true;
                    row++;
                }
                row++;

                // Header
                int headerRow = row;
                int col = 1;
                ws.Cell(headerRow, col++).Value = "STT";
                ws.Cell(headerRow, col++).Value = "Tên Cọc";
                ws.Cell(headerRow, col++).Value = "Lý Trình";
                ws.Cell(headerRow, col++).Value = "KC (m)";

                foreach (var ct in congThucs)
                {
                    ws.Cell(headerRow, col++).Value = $"DT {ct.TenVatLieu} (m²)";
                    ws.Cell(headerRow, col++).Value = $"KL {ct.TenVatLieu} (m³)";
                    ws.Cell(headerRow, col++).Value = $"KL TL {ct.TenVatLieu} (m³)";
                }

                // Format header
                var headerRange = ws.Range(headerRow, 1, headerRow, col - 1);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(27, 42, 74);
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                // Data
                int stt = 1;
                foreach (var r in results)
                {
                    row++;
                    col = 1;
                    ws.Cell(row, col++).Value = stt++;
                    ws.Cell(row, col++).Value = r.TenCoc;
                    ws.Cell(row, col++).Value = r.LyTrinh;
                    ws.Cell(row, col++).Value = Math.Round(r.KhoangCach, 3);

                    foreach (var ct in congThucs)
                    {
                        ws.Cell(row, col++).Value = r.DienTich.GetValueOrDefault(ct.TenVatLieu, 0);
                        ws.Cell(row, col++).Value = r.KhoiLuong.GetValueOrDefault(ct.TenVatLieu, 0);
                        ws.Cell(row, col++).Value = r.KhoiLuongTL.GetValueOrDefault(ct.TenVatLieu, 0);
                    }
                }

                // Tổng
                row++;
                ws.Cell(row, 1).Value = "TỔNG";
                ws.Cell(row, 1).Style.Font.Bold = true;
                col = 5; // bắt đầu từ cột DT
                foreach (var ct in congThucs)
                {
                    double tongDT = results.Sum(r => r.DienTich.GetValueOrDefault(ct.TenVatLieu, 0));
                    double tongKL = results.Sum(r => r.KhoiLuong.GetValueOrDefault(ct.TenVatLieu, 0));
                    ws.Cell(row, col++).Value = Math.Round(tongDT, 2);
                    ws.Cell(row, col).Value = Math.Round(tongKL, 2);
                    ws.Cell(row, col).Style.Font.Bold = true;
                    ws.Cell(row, col).Style.Font.FontColor = XLColor.Red;
                    col++;
                    col++; // skip TL column
                }

                // Data range border
                var dataRange = ws.Range(headerRow, 1, row, col - 1);
                dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                // Auto-fit
                ws.Columns().AdjustToContents();

                wb.SaveAs(savePath);
                ed.WriteMessage($"\n📊 Đã xuất Excel: {savePath}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n⚠️ Lỗi xuất Excel: {ex.Message}");
            }
        }

        // ═══════════════════════════════════════════════════════════
        // BỐ TRÍ BẢNG LÊN SECTIONVIEW
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// Tạo AutoCAD Table entity đặt cạnh mỗi SectionView
        /// hiển thị DT + KL tại trắc ngang đó.
        /// </summary>
        private static void BoTriBangLenSV(
            Transaction tr, Editor ed,
            SampleLineGroup slg,
            List<SampleLine> sampleLines,
            List<KetQuaTracNgang> results,
            List<CongThucKL> congThucs)
        {
            try
            {
                // Lấy SectionViewGroup
                SectionViewGroupCollection svgCol = slg.SectionViewGroups;
                if (svgCol.Count == 0)
                {
                    ed.WriteMessage("\n⚠️ Không có SectionViewGroup. Hãy vẽ trắc ngang trước.");
                    return;
                }

                // Lấy ModelSpace
                AcadDb.BlockTableRecord modelSpace = tr.GetObject(
                    A.Db.CurrentSpaceId, OpenMode.ForWrite) as AcadDb.BlockTableRecord
                    ?? throw new System.Exception("Không mở được ModelSpace");

                // Tạo layer
                string layerName = "C3D-KLTN-TABLE";
                EnsureLayer(tr, layerName, 7);

                int tableCount = 0;
                SectionViewGroup svGroup = svgCol[0];
                ObjectIdCollection svIds = svGroup.GetSectionViewIds();

                foreach (ObjectId svId in svIds)
                {
                    SectionView? sv = tr.GetObject(svId, OpenMode.ForWrite) as SectionView;
                    if (sv == null) continue;

                    SampleLine? sl = tr.GetObject(sv.SampleLineId, OpenMode.ForWrite) as SampleLine;
                    if (sl == null) continue;

                    // Tìm kết quả tương ứng
                    var result = results.FirstOrDefault(r =>
                        Math.Abs(r.Station - sl.Station) < 0.01);
                    if (result == null) continue;

                    // Tạo Table
                    int numRows = congThucs.Count + 2; // header + data rows + tổng
                    int numCols = 3; // Tên VL | DT (m²) | KL (m³)

                    ATable table = new();
                    table.SetDatabaseDefaults();
                    table.Layer = layerName;
                    // Chỉ set TableStyle nếu hợp lệ, tránh crash
                    try
                    {
                        if (A.Db.Tablestyle != ObjectId.Null && A.Db.Tablestyle.IsValid)
                            table.TableStyle = A.Db.Tablestyle;
                    }
                    catch { /* Giữ style mặc định */ }

                    table.SetSize(numRows, numCols);
                    table.SetRowHeight(5);
                    table.SetColumnWidth(25);
                    table.Columns[0].Width = 30;
                    table.Columns[1].Width = 20;
                    table.Columns[2].Width = 20;

                    // Header
                    table.Cells[0, 0].TextHeight = 2.5;
                    table.Cells[0, 0].TextString = $"KL {sl.Name}";
                    table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

                    table.Cells[1, 0].TextString = "Vật liệu";
                    table.Cells[1, 1].TextString = "DT (m²)";
                    table.Cells[1, 2].TextString = "KL (m³)";
                    for (int c = 0; c < numCols; c++)
                    {
                        table.Cells[1, c].TextHeight = 2;
                        table.Cells[1, c].Alignment = CellAlignment.MiddleCenter;
                    }

                    // Data
                    for (int i = 0; i < congThucs.Count; i++)
                    {
                        int r = i + 2;
                        string key = congThucs[i].TenVatLieu;
                        double dt = result.DienTich.GetValueOrDefault(key, 0);
                        double kl = result.KhoiLuong.GetValueOrDefault(key, 0);

                        table.Cells[r, 0].TextString = key;
                        table.Cells[r, 1].TextString = $"{dt:F2}";
                        table.Cells[r, 2].TextString = $"{kl:F2}";

                        for (int c = 0; c < numCols; c++)
                        {
                            table.Cells[r, c].TextHeight = 2;
                            table.Cells[r, c].Alignment = CellAlignment.MiddleCenter;
                        }
                    }

                    // Vị trí: đặt phía trên trái SectionView
                    double tableX = sv.Location.X;
                    double tableY = sv.Location.Y + 15; // trên SectionView
                    table.Position = new Point3d(tableX, tableY, 0);

                    modelSpace.AppendEntity(table);
                    tr.AddNewlyCreatedDBObject(table, true);
                    tableCount++;
                }

                ed.WriteMessage($"\n📐 Đã bố trí {tableCount} bảng KL lên SectionView.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n⚠️ Lỗi bố trí bảng: {ex.Message}");
            }
        }

        /// <summary>
        /// Đảm bảo layer tồn tại
        /// </summary>
        private static void EnsureLayer(Transaction tr, string layerName, short colorIndex)
        {
            var lt = tr.GetObject(A.Db.LayerTableId, OpenMode.ForRead) as AcadDb.LayerTable;
            if (lt == null) return;

            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                var ltr = new AcadDb.LayerTableRecord
                {
                    Name = layerName,
                    Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                        Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex)
                };
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }
    }
}
