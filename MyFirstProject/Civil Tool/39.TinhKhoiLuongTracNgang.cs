// 39.TinhKhoiLuongTracNgang.cs - Lệnh CTSV_TinhKhoiLuongTracNgang
// Tổng hợp: Tính diện tích Shape (đào/đắp/vật liệu) từ Material Section
//            rồi tính khối lượng theo 2 phương pháp:
//            - Average End Area: V = (S₁ + S₂) / 2 × L
//            - Prismoidal:       V = (S₁ + 4Sm + S₂) / 6 × L
// Nguồn diện tích: section.Area (API) hoặc Shoelace (SectionPoints)
// Xuất ra: Command Line + Clipboard + CSV (tùy chọn)

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using MyFirstProject.Extensions;

using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using ContentAlignment = System.Drawing.ContentAlignment;
using Section = Autodesk.Civil.DatabaseServices.Section;

[assembly: CommandClass(typeof(MyFirstProject.TinhKhoiLuongTracNgangCommands))]

namespace MyFirstProject
{
    // ══════════════════════════════════════════════════════════════
    // ENUM: Phương pháp tính
    // ══════════════════════════════════════════════════════════════

    /// <summary>
    /// Phương pháp tính khối lượng giữa 2 mặt cắt
    /// </summary>
    public enum VolumeMethod
    {
        /// <summary>V = (S₁ + S₂) / 2 × L — Mặc định Civil 3D</summary>
        AverageEndArea,
        /// <summary>V = (S₁ + 4Sm + S₂) / 6 × L — Chính xác hơn</summary>
        Prismoidal
    }

    /// <summary>
    /// Nguồn dữ liệu diện tích mặt cắt
    /// </summary>
    public enum AreaSource
    {
        /// <summary>Dùng section.Area từ Civil 3D API (chính xác nhất)</summary>
        API,
        /// <summary>Tự tính từ SectionPoints bằng Shoelace Formula</summary>
        Shoelace,
        /// <summary>Ưu tiên API, nếu = 0 thì fallback sang Shoelace</summary>
        APIWithFallback
    }

    // ══════════════════════════════════════════════════════════════
    // DATA CLASS: Kết quả tại mỗi trắc ngang
    // ══════════════════════════════════════════════════════════════

    public class SectionVolumeResult
    {
        public string TenCoc { get; set; } = "";
        public double Station { get; set; }
        public string LyTrinh { get; set; } = "";

        // Diện tích tại mặt cắt (m²) — giá trị đã chọn theo AreaSource
        public Dictionary<string, double> MaterialAreas { get; set; } = new();

        // Đối chiếu: DT từ API (section.Area)
        public Dictionary<string, double> MaterialAreas_API { get; set; } = new();

        // Đối chiếu: DT từ Shoelace (tự tính từ SectionPoints)
        public Dictionary<string, double> MaterialAreas_Shoelace { get; set; } = new();

        // Khối lượng Average End Area (m³)
        public Dictionary<string, double> MaterialVolumes_AEA { get; set; } = new();

        // Khối lượng Prismoidal (m³)
        public Dictionary<string, double> MaterialVolumes_Prismoidal { get; set; } = new();

        // Khoảng cách đến trắc ngang trước (m)
        public double DistanceToPrev { get; set; }

        // Khối lượng tích lũy AEA
        public Dictionary<string, double> CumulativeVolumes_AEA { get; set; } = new();

        // Khối lượng tích lũy Prismoidal
        public Dictionary<string, double> CumulativeVolumes_Prismoidal { get; set; } = new();
    }

    // ══════════════════════════════════════════════════════════════
    // FORM: CHỌN CÁC LOẠI MATERIAL ĐỂ TÍNH
    // ══════════════════════════════════════════════════════════════

    public class TinhKhoiLuongForm : CivilFormBase
    {
        // Data
        private List<(string Name, ObjectId AlignmentId, ObjectId SLGId)> alignmentList = new();
        private List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> materialList = new();

        // Controls
        private ComboBox cboAlignment = null!;
        private CheckedListBox chkMaterials = null!;
        private RadioButton rdoArea_API = null!;
        private RadioButton rdoArea_Shoelace = null!;
        private RadioButton rdoArea_Fallback = null!;
        private RadioButton rdoVol_AEA = null!;
        private RadioButton rdoVol_Prismoidal = null!;
        private CheckBox chkCompare = null!;
        private CheckBox chkExportCSV = null!;
        private CheckBox chkExportExcel = null!;
        private CheckBox chkWriteToDrawing = null!;

        // Kết quả
        public bool IsConfirmed { get; private set; }
        public ObjectId SelectedAlignmentId =>
            cboAlignment.SelectedIndex >= 0 ? alignmentList[cboAlignment.SelectedIndex].AlignmentId : ObjectId.Null;
        public ObjectId SelectedSLGId =>
            cboAlignment.SelectedIndex >= 0 ? alignmentList[cboAlignment.SelectedIndex].SLGId : ObjectId.Null;
        public List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> SelectedMaterials
        {
            get
            {
                var result = new List<(string, string, Guid, Guid)>();
                for (int i = 0; i < chkMaterials.Items.Count; i++)
                {
                    if (chkMaterials.GetItemChecked(i))
                        result.Add(materialList[i]);
                }
                return result;
            }
        }
        public AreaSource SelectedAreaSource =>
            rdoArea_API.Checked ? AreaSource.API :
            rdoArea_Shoelace.Checked ? AreaSource.Shoelace :
            AreaSource.APIWithFallback;
        public VolumeMethod SelectedVolumeMethod =>
            rdoVol_Prismoidal.Checked ? VolumeMethod.Prismoidal : VolumeMethod.AverageEndArea;
        public bool CompareMethods => chkCompare.Checked;
        public bool ExportCSV => chkExportCSV.Checked;
        public bool ExportExcel => chkExportExcel.Checked;
        public bool WriteToDrawing => chkWriteToDrawing.Checked;

        public TinhKhoiLuongForm()
        {
            LoadData();
            InitializeComponent();
        }

        private void LoadData()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var civilDoc = CivilApplication.ActiveDocument;

            using var tr = db.TransactionManager.StartOpenCloseTransaction();

            foreach (ObjectId alId in civilDoc.GetAlignmentIds())
            {
                var al = tr.GetObject(alId, OpenMode.ForRead) as Alignment;
                if (al == null) continue;

                var slgIds = al.GetSampleLineGroupIds();
                foreach (ObjectId slgId in slgIds)
                {
                    var slg = tr.GetObject(slgId, OpenMode.ForRead) as SampleLineGroup;
                    if (slg != null)
                        alignmentList.Add(($"{al.Name} — {slg.Name}", alId, slgId));
                }
            }
            tr.Commit();
        }

        private void LoadMaterials()
        {
            materialList.Clear();
            chkMaterials.Items.Clear();

            if (cboAlignment.SelectedIndex < 0) return;

            var slgId = alignmentList[cboAlignment.SelectedIndex].SLGId;
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using var tr = doc.Database.TransactionManager.StartOpenCloseTransaction();
            var slg = tr.GetObject(slgId, OpenMode.ForRead) as SampleLineGroup;
            if (slg == null) { tr.Commit(); return; }

            foreach (QTOMaterialList matList in slg.MaterialLists)
            {
                try
                {
                    foreach (QTOMaterial mat in matList)
                    {
                        string type = mat.QuantityType == MaterialQuantityType.Cut ? "🔻 Đào" :
                                      mat.QuantityType == MaterialQuantityType.Fill ? "🔺 Đắp" :
                                      mat.QuantityType == MaterialQuantityType.CutFill ? "🔄 Đào+Đắp" :
                                      mat.QuantityType == MaterialQuantityType.Earthwork ? "🏔️ Earthwork" :
                                      mat.QuantityType == MaterialQuantityType.Structure ? "🏗️ Kết cấu" : "📦 Khác";
                        materialList.Add((mat.Name, matList.Name ?? "Unnamed", matList.Guid, mat.Guid));
                        int idx = chkMaterials.Items.Add($"{type}: {mat.Name} ({matList.Name})");
                        chkMaterials.SetItemChecked(idx, true);
                    }
                }
                catch { }
            }
            tr.Commit();

            if (materialList.Count == 0)
            {
                chkMaterials.Items.Add("⚠️ Chưa có Material List. Hãy chạy Compute Materials trước.");
            }
        }

        private void InitializeComponent()
        {
            InitBaseTheme("📐  TÍNH KHỐI LƯỢNG TRẮC NGANG", 500, 750);

            int y = 0;

            // ═══ HEADER ═══
            var header = CreateHeader("TÍNH KHỐI LƯỢNG TRẮC NGANG");
            this.Controls.Add(header);

            var subtitle = CreateSubtitle("Lấy diện tích Material Section → Tính KL theo Average End Area / Prismoidal");
            this.Controls.Add(subtitle);

            y = 82;

            // ═══ CHỌN ALIGNMENT ═══
            var grpAlign = CreateGroupBox("📏 Chọn Alignment + SampleLine Group", 16, y, 460, 64);
            this.Controls.Add(grpAlign);

            cboAlignment = CreateComboBox(12, 26, 436);
            foreach (var (name, _, _) in alignmentList)
                cboAlignment.Items.Add(name);
            if (cboAlignment.Items.Count > 0)
                cboAlignment.SelectedIndex = 0;
            cboAlignment.SelectedIndexChanged += (s, e) => LoadMaterials();
            grpAlign.Controls.Add(cboAlignment);

            y += 72;

            // ═══ CHỌN MATERIAL ═══
            var grpMat = CreateGroupBox("📦 Chọn Material cần tính", 16, y, 460, 140);
            this.Controls.Add(grpMat);

            chkMaterials = new CheckedListBox
            {
                Location = new Point(12, 26),
                Size = new Size(436, 100),
                CheckOnClick = true,
                Font = new Font("Segoe UI", 9.5f),
                BackColor = BgWhite,
                ForeColor = TextDark,
                BorderStyle = BorderStyle.FixedSingle
            };
            grpMat.Controls.Add(chkMaterials);

            y += 148;

            // ═══ PHƯƠNG PHÁP TÍNH DIỆN TÍCH ═══
            var grpArea = CreateGroupBox("📐 Nguồn diện tích mặt cắt", 16, y, 460, 90);
            this.Controls.Add(grpArea);

            rdoArea_Fallback = new RadioButton
            {
                Text = "⭐ section.Area (API) — fallback Shoelace",
                Location = new Point(12, 26),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9f)
            };
            grpArea.Controls.Add(rdoArea_Fallback);

            rdoArea_API = new RadioButton
            {
                Text = "section.Area (API) — chỉ dùng API",
                Location = new Point(12, 48),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9f)
            };
            grpArea.Controls.Add(rdoArea_API);

            rdoArea_Shoelace = new RadioButton
            {
                Text = "Shoelace (tự tính từ SectionPoints)",
                Location = new Point(12, 68),
                AutoSize = true,
                ForeColor = TextHint,
                Font = new Font("Segoe UI", 9f)
            };
            grpArea.Controls.Add(rdoArea_Shoelace);

            y += 98;

            // ═══ PHƯƠNG PHÁP TÍNH KHỐI LƯỢNG ═══
            var grpVol = CreateGroupBox("📊 Phương pháp tính khối lượng", 16, y, 460, 108);
            this.Controls.Add(grpVol);

            rdoVol_AEA = new RadioButton
            {
                Text = "⭐ Average End Area:  V = (S₁ + S₂) / 2 × L",
                Location = new Point(12, 26),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpVol.Controls.Add(rdoVol_AEA);

            rdoVol_Prismoidal = new RadioButton
            {
                Text = "Prismoidal:  V = (S₁ + 4Sm + S₂) / 6 × L",
                Location = new Point(12, 50),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpVol.Controls.Add(rdoVol_Prismoidal);

            chkCompare = new CheckBox
            {
                Text = "📊 So sánh cả 2 phương pháp (hiển thị chênh lệch %)",
                Location = new Point(12, 78),
                AutoSize = true,
                Checked = false,
                ForeColor = NavyDark,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            grpVol.Controls.Add(chkCompare);

            y += 116;

            // ═══ TÙY CHỌN XUẤT ═══
            var grpExport = CreateGroupBox("📤 Tùy chọn xuất", 16, y, 460, 85);
            this.Controls.Add(grpExport);

            chkExportCSV = new CheckBox
            {
                Text = "📄 Xuất CSV",
                Location = new Point(12, 26),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpExport.Controls.Add(chkExportCSV);

            chkExportExcel = new CheckBox
            {
                Text = "📊 Xuất Excel (.xlsx)",
                Location = new Point(160, 26),
                AutoSize = true,
                Checked = false,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpExport.Controls.Add(chkExportExcel);

            chkWriteToDrawing = new CheckBox
            {
                Text = "✏️ Ghi text lên trắc ngang",
                Location = new Point(12, 52),
                AutoSize = true,
                Checked = false,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpExport.Controls.Add(chkWriteToDrawing);

            y += 93;

            // ═══ BOTTOM BAR ═══
            var bottomBar = CreateBottomBar(56);
            this.Controls.Add(bottomBar);

            var btnCancel = CreateButton("✕ Hủy", 10, 10, 220, 36, bottomBar);
            btnCancel.Click += (s, e) => { IsConfirmed = false; this.Close(); };

            var btnRun = CreateSuccessButton("📐 Tính Khối Lượng", 240, 10, 220, 36, bottomBar);
            btnRun.Click += (s, e) =>
            {
                if (cboAlignment.SelectedIndex < 0)
                {
                    System.Windows.Forms.MessageBox.Show("Vui lòng chọn Alignment!", "Thiếu dữ liệu",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (SelectedMaterials.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show("Vui lòng chọn ít nhất 1 Material!", "Thiếu dữ liệu",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                IsConfirmed = true;
                this.Close();
            };

            this.AcceptButton = btnRun;
            this.CancelButton = btnCancel;

            // Load materials cho alignment đầu tiên
            if (cboAlignment.Items.Count > 0)
                LoadMaterials();
        }
    }

    // ══════════════════════════════════════════════════════════════
    // LỆNH CHÍNH
    // ══════════════════════════════════════════════════════════════

    public class TinhKhoiLuongTracNgangCommands
    {
        /// <summary>
        /// Tính khối lượng trắc ngang từ Material Section
        /// Hỗ trợ: Average End Area + Prismoidal + So sánh
        /// Nguồn diện tích: section.Area (API) / Shoelace / Fallback
        /// </summary>
        [CommandMethod("CTSV_TinhKhoiLuongTracNgang")]
        public static void CTSV_TinhKhoiLuongTracNgang()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            // 1. Mở form chọn
            using var form = new TinhKhoiLuongForm();
            AcadApp.ShowModalDialog(form);
            if (!form.IsConfirmed) return;

            var alignmentId = form.SelectedAlignmentId;
            var slgId = form.SelectedSLGId;
            var selectedMaterials = form.SelectedMaterials;
            var areaSource = form.SelectedAreaSource;
            var volumeMethod = form.SelectedVolumeMethod;
            bool compareMethods = form.CompareMethods;
            bool exportCSV = form.ExportCSV;
            bool exportExcel = form.ExportExcel;
            bool writeToDrawing = form.WriteToDrawing;

            if (alignmentId.IsNull || slgId.IsNull)
            {
                ed.WriteMessage("\n❌ Không chọn được Alignment.");
                return;
            }

            try
            {
                using var tr = db.TransactionManager.StartTransaction();

                var alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                var slg = tr.GetObject(slgId, OpenMode.ForRead) as SampleLineGroup;
                if (alignment == null || slg == null) { tr.Commit(); return; }

                var slIds = slg.GetSampleLineIds();

                // 2. Thu thập diện tích cho từng sample line
                var results = new List<SectionVolumeResult>();

                foreach (ObjectId slId in slIds)
                {
                    var sl = tr.GetObject(slId, OpenMode.ForRead) as SampleLine;
                    if (sl == null) continue;

                    var result = new SectionVolumeResult
                    {
                        TenCoc = sl.Name,
                        Station = sl.Station,
                        LyTrinh = FormatStation(sl.Station)
                    };

                    // Tính diện tích cho mỗi Material + thu thập dữ liệu đối chiếu
                    foreach (var (matName, listName, listGuid, matGuid) in selectedMaterials)
                    {
                        var (area, apiArea, shoelaceArea) = GetMaterialAreaWithVerify(tr, sl, listGuid, matGuid, areaSource);
                        result.MaterialAreas[matName] = Math.Round(area, 4);
                        result.MaterialAreas_API[matName] = Math.Round(apiArea, 4);
                        result.MaterialAreas_Shoelace[matName] = Math.Round(shoelaceArea, 4);
                    }

                    results.Add(result);
                }

                // Sắp xếp theo station
                results = results.OrderBy(r => r.Station).ToList();

                // 3. Tính khối lượng
                CalculateVolumes(results, selectedMaterials);

                // 4. Hiển thị kết quả
                DisplayResults(ed, alignment.Name, results, selectedMaterials,
                    volumeMethod, compareMethods, areaSource);

                // 5. Copy vào clipboard
                CopyToClipboard(alignment.Name, results, selectedMaterials,
                    volumeMethod, compareMethods);

                // 6. Xuất CSV nếu được chọn
                if (exportCSV)
                    ExportToCSV(ed, alignment.Name, results, selectedMaterials,
                        volumeMethod, compareMethods);

                // 6b. Xuất Excel nếu được chọn
                if (exportExcel)
                    ExportToExcel(ed, alignment.Name, results, selectedMaterials,
                        volumeMethod, compareMethods);

                // 7. Ghi text lên bản vẽ nếu được chọn
                if (writeToDrawing)
                    WriteToDrawing(tr, ed, slg, results, selectedMaterials,
                        alignment.Name, volumeMethod);

                tr.Commit();

                ed.WriteMessage("\n\n✅ HOÀN THÀNH tính khối lượng trắc ngang!");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        // ═══════════════════════════════════════════════════════════
        // LẤY DIỆN TÍCH MATERIAL SECTION
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// Lấy diện tích Material Section theo nguồn đã chọn
        /// </summary>
        private static double GetMaterialArea(Transaction tr, SampleLine sampleLine,
            Guid listGuid, Guid matGuid, AreaSource areaSource)
        {
            var (area, _, _) = GetMaterialAreaWithVerify(tr, sampleLine, listGuid, matGuid, areaSource);
            return area;
        }

        /// <summary>
        /// Lấy diện tích + cả 2 nguồn (API + Shoelace) để đối chiếu
        /// Returns: (chọn theo areaSource, API area, Shoelace area)
        /// </summary>
        private static (double chosen, double apiArea, double shoelaceArea) GetMaterialAreaWithVerify(
            Transaction tr, SampleLine sampleLine,
            Guid listGuid, Guid matGuid, AreaSource areaSource)
        {
            try
            {
                ObjectId materialSectionId = sampleLine.GetMaterialSectionId(listGuid, matGuid);
                if (materialSectionId.IsNull || !materialSectionId.IsValid) return (0, 0, 0);

                var section = tr.GetObject(materialSectionId, OpenMode.ForRead, false, true) as Section;
                if (section == null) return (0, 0, 0);

                double api = section.Area;
                double shoelace = CalculateShoelaceArea(section);

                double chosen = areaSource switch
                {
                    AreaSource.API => api,
                    AreaSource.Shoelace => shoelace,
                    _ => api > 0 ? api : shoelace, // APIWithFallback
                };

                return (chosen, api, shoelace);
            }
            catch
            {
                return (0, 0, 0);
            }
        }

        /// <summary>
        /// Tính diện tích bằng Shoelace Formula từ SectionPoints
        /// </summary>
        private static double CalculateShoelaceArea(Section section)
        {
            try
            {
                SectionPointCollection sectionPoints = section.SectionPoints;
                if (sectionPoints.Count < 3) return 0;

                double area = 0;
                int n = sectionPoints.Count;

                for (int i = 0; i < n; i++)
                {
                    int j = (i + 1) % n;
                    double xi = sectionPoints[i].Location.X;
                    double yi = sectionPoints[i].Location.Y;
                    double xj = sectionPoints[j].Location.X;
                    double yj = sectionPoints[j].Location.Y;

                    area += xi * yj - xj * yi;
                }

                return Math.Abs(area) / 2.0;
            }
            catch
            {
                return 0;
            }
        }

        // ═══════════════════════════════════════════════════════════
        // TÍNH KHỐI LƯỢNG — 2 PHƯƠNG PHÁP
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// Tính khối lượng cho tất cả trắc ngang theo cả 2 phương pháp
        /// </summary>
        private static void CalculateVolumes(List<SectionVolumeResult> results,
            List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> materials)
        {
            // Khởi tạo tích lũy
            var cumAEA = new Dictionary<string, double>();
            var cumPrismoidal = new Dictionary<string, double>();
            foreach (var mat in materials)
            {
                cumAEA[mat.MaterialName] = 0;
                cumPrismoidal[mat.MaterialName] = 0;
            }

            for (int i = 0; i < results.Count; i++)
            {
                if (i == 0)
                {
                    // Trắc ngang đầu tiên: chưa có volume
                    results[i].DistanceToPrev = 0;
                    foreach (var mat in materials)
                    {
                        results[i].MaterialVolumes_AEA[mat.MaterialName] = 0;
                        results[i].MaterialVolumes_Prismoidal[mat.MaterialName] = 0;
                        results[i].CumulativeVolumes_AEA[mat.MaterialName] = 0;
                        results[i].CumulativeVolumes_Prismoidal[mat.MaterialName] = 0;
                    }
                }
                else
                {
                    double dist = results[i].Station - results[i - 1].Station;
                    results[i].DistanceToPrev = dist;

                    foreach (var mat in materials)
                    {
                        double s1 = results[i - 1].MaterialAreas.GetValueOrDefault(mat.MaterialName, 0);
                        double s2 = results[i].MaterialAreas.GetValueOrDefault(mat.MaterialName, 0);

                        // Average End Area: V = (S₁ + S₂) / 2 × L
                        double volAEA = VolumeAverageEndArea(s1, s2, dist);
                        results[i].MaterialVolumes_AEA[mat.MaterialName] = Math.Round(volAEA, 3);

                        // Prismoidal: V = (S₁ + 4Sm + S₂) / 6 × L
                        double volPrismoidal = VolumePrismoidal(s1, s2, dist);
                        results[i].MaterialVolumes_Prismoidal[mat.MaterialName] = Math.Round(volPrismoidal, 3);

                        // Tích lũy
                        cumAEA[mat.MaterialName] += volAEA;
                        cumPrismoidal[mat.MaterialName] += volPrismoidal;

                        results[i].CumulativeVolumes_AEA[mat.MaterialName] = Math.Round(cumAEA[mat.MaterialName], 3);
                        results[i].CumulativeVolumes_Prismoidal[mat.MaterialName] = Math.Round(cumPrismoidal[mat.MaterialName], 3);
                    }
                }
            }
        }

        /// <summary>
        /// Phương pháp Average End Area: V = (S₁ + S₂) / 2 × L
        /// Mặc định Civil 3D — đơn giản, phổ biến nhất
        /// </summary>
        private static double VolumeAverageEndArea(double s1, double s2, double distance)
        {
            return (s1 + s2) / 2.0 * distance;
        }

        /// <summary>
        /// Phương pháp Prismoidal: V = (S₁ + 4Sm + S₂) / 6 × L
        /// Sm = √(S₁ × S₂) — trung bình hình học (mid-section)
        /// Chính xác hơn cho địa hình phức tạp
        /// </summary>
        private static double VolumePrismoidal(double s1, double s2, double distance)
        {
            // Trường hợp đặc biệt: một trong 2 diện tích = 0
            if (s1 <= 0 && s2 <= 0) return 0;
            if (s1 <= 0 || s2 <= 0)
            {
                // Khi 1 mặt cắt = 0 → Prismoidal = V_cone = S/3 × L
                // Nhưng để nhất quán, dùng công thức chung với Sm = 0
                double sm = 0;
                return (s1 + 4 * sm + s2) / 6.0 * distance;
            }

            // Sm = trung bình hình học √(S₁ × S₂)
            double smid = Math.Sqrt(s1 * s2);
            return (s1 + 4 * smid + s2) / 6.0 * distance;
        }

        // ═══════════════════════════════════════════════════════════
        // HIỂN THỊ KẾT QUẢ
        // ═══════════════════════════════════════════════════════════

        private static void DisplayResults(Editor ed, string alignmentName,
            List<SectionVolumeResult> results,
            List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> materials,
            VolumeMethod method, bool compare, AreaSource areaSource)
        {
            ed.WriteMessage("\n");
            ed.WriteMessage($"\n╔══════════════════════════════════════════════════════════════════════════════╗");
            ed.WriteMessage($"\n║  📐 BẢNG KHỐI LƯỢNG TRẮC NGANG — {alignmentName,-40} ║");
            ed.WriteMessage($"\n╠══════════════════════════════════════════════════════════════════════════════╣");
            ed.WriteMessage($"\n║  Nguồn DT: {(areaSource == AreaSource.API ? "section.Area (API)" : areaSource == AreaSource.Shoelace ? "Shoelace (SectionPoints)" : "API + Fallback Shoelace"),-20}");
            ed.WriteMessage($"\n║  Phương pháp: {(compare ? "So sánh cả 2" : method == VolumeMethod.AverageEndArea ? "Average End Area" : "Prismoidal"),-20}");
            ed.WriteMessage($"\n╚══════════════════════════════════════════════════════════════════════════════╝");

            // ─── Header ───
            string headerBase = $"\n{"STT",-5} {"Tên Cọc",-14} {"Lý Trình",-14} {"L(m)",-8}";
            string headerMats = "";

            foreach (var mat in materials)
            {
                string shortName = mat.MaterialName.Length > 8
                    ? mat.MaterialName.Substring(0, 8) : mat.MaterialName;

                headerMats += $" {"S_" + shortName,-10}";

                if (compare)
                {
                    headerMats += $" {"V_AEA",-10} {"V_Pris",-10} {"Δ%",-6}";
                }
                else
                {
                    headerMats += $" {"V(m³)",-10}";
                }
            }
            ed.WriteMessage(headerBase + headerMats);

            int lineLen = 42 + materials.Count * (compare ? 38 : 22);
            ed.WriteMessage($"\n{new string('─', lineLen)}");

            // ─── Data ───
            var totalAEA = new Dictionary<string, double>();
            var totalPris = new Dictionary<string, double>();
            foreach (var mat in materials)
            {
                totalAEA[mat.MaterialName] = 0;
                totalPris[mat.MaterialName] = 0;
            }

            int stt = 1;
            foreach (var r in results)
            {
                string line = $"\n{stt,-5} {r.TenCoc,-14} {r.LyTrinh,-14} {r.DistanceToPrev,-8:F2}";

                foreach (var mat in materials)
                {
                    double area = r.MaterialAreas.GetValueOrDefault(mat.MaterialName, 0);
                    double volAEA = r.MaterialVolumes_AEA.GetValueOrDefault(mat.MaterialName, 0);
                    double volPris = r.MaterialVolumes_Prismoidal.GetValueOrDefault(mat.MaterialName, 0);

                    line += $" {area,-10:F3}";

                    if (compare)
                    {
                        double diff = volAEA > 0 ? ((volAEA - volPris) / volAEA * 100) : 0;
                        line += $" {volAEA,-10:F3} {volPris,-10:F3} {diff,-6:F1}";
                    }
                    else
                    {
                        double vol = method == VolumeMethod.AverageEndArea ? volAEA : volPris;
                        line += $" {vol,-10:F3}";
                    }

                    totalAEA[mat.MaterialName] += volAEA;
                    totalPris[mat.MaterialName] += volPris;
                }

                ed.WriteMessage(line);
                stt++;
            }

            // ─── Tổng ───
            ed.WriteMessage($"\n{new string('═', lineLen)}");
            string totalLine = $"\n{"TỔNG",-5} {"",-14} {"",-14} {"",-8}";
            foreach (var mat in materials)
            {
                totalLine += $" {"",-10}";
                if (compare)
                {
                    double diff = totalAEA[mat.MaterialName] > 0
                        ? ((totalAEA[mat.MaterialName] - totalPris[mat.MaterialName]) / totalAEA[mat.MaterialName] * 100)
                        : 0;
                    totalLine += $" {totalAEA[mat.MaterialName],-10:F3} {totalPris[mat.MaterialName],-10:F3} {diff,-6:F1}";
                }
                else
                {
                    double vol = method == VolumeMethod.AverageEndArea
                        ? totalAEA[mat.MaterialName] : totalPris[mat.MaterialName];
                    totalLine += $" {vol,-10:F3}";
                }
            }
            ed.WriteMessage(totalLine);
            ed.WriteMessage($"\n{new string('═', lineLen)}");

            // ─── Tóm tắt ───
            ed.WriteMessage("\n\n📊 TÓM TẮT:");
            foreach (var mat in materials)
            {
                if (compare)
                {
                    double diff = totalAEA[mat.MaterialName] > 0
                        ? ((totalAEA[mat.MaterialName] - totalPris[mat.MaterialName]) / totalAEA[mat.MaterialName] * 100) : 0;
                    ed.WriteMessage($"\n   {mat.MaterialName}:");
                    ed.WriteMessage($"\n      Average End Area : {totalAEA[mat.MaterialName]:N3} m³");
                    ed.WriteMessage($"\n      Prismoidal       : {totalPris[mat.MaterialName]:N3} m³");
                    ed.WriteMessage($"\n      Chênh lệch       : {diff:F2}% (AEA > Prismoidal)");
                }
                else
                {
                    double vol = method == VolumeMethod.AverageEndArea
                        ? totalAEA[mat.MaterialName] : totalPris[mat.MaterialName];
                    string methodName = method == VolumeMethod.AverageEndArea ? "AEA" : "Prismoidal";
                    ed.WriteMessage($"\n   {mat.MaterialName} ({methodName}): {vol:N3} m³");
                }
            }
        }

        // ═══════════════════════════════════════════════════════════
        // COPY VÀO CLIPBOARD
        // ═══════════════════════════════════════════════════════════

        private static void CopyToClipboard(string alignmentName,
            List<SectionVolumeResult> results,
            List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> materials,
            VolumeMethod method, bool compare)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"KHỐI LƯỢNG TRẮC NGANG — {alignmentName}");

            // Header
            sb.Append("STT\tTên Cọc\tLý Trình\tKhoảng cách (m)");
            foreach (var mat in materials)
            {
                sb.Append($"\tDT_{mat.MaterialName} (m²)");
                if (compare)
                {
                    sb.Append($"\tKL_AEA_{mat.MaterialName} (m³)");
                    sb.Append($"\tKL_Prismoidal_{mat.MaterialName} (m³)");
                    sb.Append($"\tChênh lệch (%)");
                }
                else
                {
                    string methodTag = method == VolumeMethod.AverageEndArea ? "AEA" : "Prismoidal";
                    sb.Append($"\tKL_{methodTag}_{mat.MaterialName} (m³)");
                }
            }
            sb.AppendLine();

            // Data
            int stt = 1;
            foreach (var r in results)
            {
                sb.Append($"{stt}\t{r.TenCoc}\t{r.LyTrinh}\t{r.DistanceToPrev:F2}");
                foreach (var mat in materials)
                {
                    double area = r.MaterialAreas.GetValueOrDefault(mat.MaterialName, 0);
                    double volAEA = r.MaterialVolumes_AEA.GetValueOrDefault(mat.MaterialName, 0);
                    double volPris = r.MaterialVolumes_Prismoidal.GetValueOrDefault(mat.MaterialName, 0);

                    sb.Append($"\t{area:F3}");
                    if (compare)
                    {
                        double diff = volAEA > 0 ? ((volAEA - volPris) / volAEA * 100) : 0;
                        sb.Append($"\t{volAEA:F3}\t{volPris:F3}\t{diff:F2}");
                    }
                    else
                    {
                        double vol = method == VolumeMethod.AverageEndArea ? volAEA : volPris;
                        sb.Append($"\t{vol:F3}");
                    }
                }
                sb.AppendLine();
                stt++;
            }

            try
            {
                System.Windows.Forms.Clipboard.SetText(sb.ToString());
            }
            catch { }
        }

        // ═══════════════════════════════════════════════════════════
        // XUẤT CSV
        // ═══════════════════════════════════════════════════════════

        private static void ExportToCSV(Editor ed, string alignmentName,
            List<SectionVolumeResult> results,
            List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> materials,
            VolumeMethod method, bool compare)
        {
            var sfd = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                Title = "Lưu file khối lượng trắc ngang",
                FileName = $"KL_TracNgang_{alignmentName}.csv"
            };

            if (sfd.ShowDialog() != DialogResult.OK) return;

            var sb = new StringBuilder();

            // Header
            sb.Append("STT,TenCoc,LyTrinh,Station,KhoangCach_m");
            foreach (var mat in materials)
            {
                sb.Append($",DT_{mat.MaterialName}_m2");
                if (compare)
                {
                    sb.Append($",KL_AEA_{mat.MaterialName}_m3");
                    sb.Append($",KL_Pris_{mat.MaterialName}_m3");
                    sb.Append($",TichLuy_AEA_{mat.MaterialName}_m3");
                    sb.Append($",TichLuy_Pris_{mat.MaterialName}_m3");
                    sb.Append($",ChenhLech_Pct");
                }
                else
                {
                    string tag = method == VolumeMethod.AverageEndArea ? "AEA" : "Pris";
                    sb.Append($",KL_{tag}_{mat.MaterialName}_m3");
                    sb.Append($",TichLuy_{tag}_{mat.MaterialName}_m3");
                }
            }
            sb.AppendLine();

            // Data
            int stt = 1;
            foreach (var r in results)
            {
                sb.Append($"{stt},{r.TenCoc},{r.LyTrinh},{r.Station:F3},{r.DistanceToPrev:F3}");
                foreach (var mat in materials)
                {
                    double area = r.MaterialAreas.GetValueOrDefault(mat.MaterialName, 0);
                    double volAEA = r.MaterialVolumes_AEA.GetValueOrDefault(mat.MaterialName, 0);
                    double volPris = r.MaterialVolumes_Prismoidal.GetValueOrDefault(mat.MaterialName, 0);
                    double cumAEA = r.CumulativeVolumes_AEA.GetValueOrDefault(mat.MaterialName, 0);
                    double cumPris = r.CumulativeVolumes_Prismoidal.GetValueOrDefault(mat.MaterialName, 0);

                    sb.Append($",{area:F3}");
                    if (compare)
                    {
                        double diff = volAEA > 0 ? ((volAEA - volPris) / volAEA * 100) : 0;
                        sb.Append($",{volAEA:F3},{volPris:F3},{cumAEA:F3},{cumPris:F3},{diff:F2}");
                    }
                    else
                    {
                        double vol = method == VolumeMethod.AverageEndArea ? volAEA : volPris;
                        double cum = method == VolumeMethod.AverageEndArea ? cumAEA : cumPris;
                        sb.Append($",{vol:F3},{cum:F3}");
                    }
                }
                sb.AppendLine();
                stt++;
            }

            // Dòng tổng
            sb.Append($"TỔNG,,,,");
            foreach (var mat in materials)
            {
                double totalAEA = results.Sum(r => r.MaterialVolumes_AEA.GetValueOrDefault(mat.MaterialName, 0));
                double totalPris = results.Sum(r => r.MaterialVolumes_Prismoidal.GetValueOrDefault(mat.MaterialName, 0));

                sb.Append($",");  // Area trống
                if (compare)
                {
                    double diff = totalAEA > 0 ? ((totalAEA - totalPris) / totalAEA * 100) : 0;
                    sb.Append($",{totalAEA:F3},{totalPris:F3},,,{diff:F2}");
                }
                else
                {
                    double vol = method == VolumeMethod.AverageEndArea ? totalAEA : totalPris;
                    sb.Append($",{vol:F3},");
                }
            }
            sb.AppendLine();

            File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
            ed.WriteMessage($"\n✅ Đã xuất file: {sfd.FileName}");
        }

        // ═══════════════════════════════════════════════════════════
        // XUẤT EXCEL (.xlsx) — ClosedXML
        // Layout: 1 sheet/tuyến (Mục A: Diện tích + Mục B: Khối lượng)
        //         + 1 sheet "Tổng hợp" toàn bộ tuyến
        // ═══════════════════════════════════════════════════════════

        private static void ExportToExcel(Editor ed, string alignmentName,
            List<SectionVolumeResult> results,
            List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> materials,
            VolumeMethod method, bool compare)
        {
            var sfd = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                Title = "Lưu file Excel khối lượng trắc ngang",
                FileName = $"KL_TracNgang_{alignmentName}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                using var wb = new XLWorkbook();
                string methodTag = method == VolumeMethod.AverageEndArea ? "AEA" : "Prismoidal";
                int colsPerMat = compare ? 5 : 3;

                // Ghi sheet cho tuyến này
                WriteAlignmentSheet(wb, alignmentName, results, materials, method, compare, methodTag, colsPerMat);

                // Ghi sheet tổng hợp (hiện tại 1 tuyến, sau này mở rộng nhiều tuyến)
                var allData = new List<(string AlignmentName, List<SectionVolumeResult> Results,
                    List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> Materials)>
                {
                    (alignmentName, results, materials)
                };
                WriteSummarySheet(wb, allData, method, compare, methodTag);

                // Ghi sheet đối chiếu DT (API vs Shoelace)
                WriteVerifySheet(wb, alignmentName, results, materials);

                wb.SaveAs(sfd.FileName);
                ed.WriteMessage($"\n✅ Đã xuất file Excel: {sfd.FileName}");

                if (MessageBox.Show("Đã xuất file Excel!\n\n" +
                    $"• Sheet \"{alignmentName}\": Mục A (Diện tích) + Mục B (Khối lượng)\n" +
                    "• Sheet \"Tổng hợp KL\": Tổng hợp khối lượng\n" +
                    "• Sheet \"Đối chiếu\": So sánh API vs Shoelace\n\n" +
                    "Bạn có muốn mở file?",
                    "Hoàn thành", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = sfd.FileName, UseShellExecute = true });
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi xuất Excel: {ex.Message}");
            }
        }

        // ───────────────────────────────────────────────────────────
        // Ghi 1 sheet cho 1 tuyến: Mục A (Diện tích) rồi Mục B (Khối lượng)
        // ───────────────────────────────────────────────────────────

        private static void WriteAlignmentSheet(XLWorkbook wb, string alignmentName,
            List<SectionVolumeResult> results,
            List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> materials,
            VolumeMethod method, bool compare, string methodTag, int colsPerMat)
        {
            // Tên sheet tối đa 31 ký tự
            string sheetName = alignmentName.Length > 31 ? alignmentName.Substring(0, 31) : alignmentName;
            var ws = wb.Worksheets.Add(sheetName);

            int matCount = materials.Count;
            int areaTotalCols = 4 + matCount;
            int volTotalCols = 4 + matCount * colsPerMat;
            int maxCols = Math.Max(areaTotalCols, volTotalCols);

            // ╔══════════════════════════════════════════╗
            // ║  MỤC A: BẢNG TỔNG HỢP DIỆN TÍCH        ║
            // ╚══════════════════════════════════════════╝

            int row = 1;
            // Title Mục A
            ws.Cell(row, 1).Value = "A. BẢNG TỔNG HỢP DIỆN TÍCH MẶT CẮT NGANG";
            ws.Range(row, 1, row, maxCols).Merge();
            StyleSectionTitle(ws.Cell(row, 1));
            row++;

            ws.Cell(row, 1).Value = $"Tuyến: {alignmentName}";
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 4).Value = "Đơn vị: m²";
            ws.Cell(row, 4).Style.Font.Italic = true;
            row += 2;

            // Header Diện tích
            int areaHeaderRow = row;
            int col = 1;
            foreach (var h in new[] { "STT", "Tên Cọc", "Lý Trình", "Station" })
                StyleHeaderCell(ws.Cell(row, col++), h);
            foreach (var mat in materials)
                StyleHeaderCell(ws.Cell(row, col++), $"S_{mat.MaterialName} (m²)", true);
            row++;

            // Data Diện tích
            int stt = 1;
            foreach (var r in results)
            {
                col = 1;
                ws.Cell(row, col++).Value = stt;
                ws.Cell(row, col++).Value = r.TenCoc;
                ws.Cell(row, col++).Value = r.LyTrinh;
                ws.Cell(row, col++).Value = Math.Round(r.Station, 3);
                foreach (var mat in materials)
                    ws.Cell(row, col++).Value = Math.Round(r.MaterialAreas.GetValueOrDefault(mat.MaterialName, 0), 4);
                if (stt % 2 == 0)
                    ws.Range(row, 1, row, areaTotalCols).Style.Fill.BackgroundColor = XLColor.FromArgb(240, 245, 250);
                row++; stt++;
            }

            // Border & format Mục A
            ws.Range(areaHeaderRow, 1, row - 1, areaTotalCols).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            ws.Range(areaHeaderRow, 1, row - 1, areaTotalCols).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            for (int c = 5; c <= areaTotalCols; c++)
                ws.Column(c).Style.NumberFormat.Format = "#,##0.0000";

            // ╔══════════════════════════════════════════╗
            // ║  MỤC B: BẢNG TÍNH KHỐI LƯỢNG            ║
            // ╚══════════════════════════════════════════╝

            row += 2; // Khoảng cách giữa 2 mục

            // Title Mục B
            ws.Cell(row, 1).Value = "B. BẢNG TÍNH KHỐI LƯỢNG ĐÀO ĐẮP";
            ws.Range(row, 1, row, maxCols).Merge();
            StyleSectionTitle(ws.Cell(row, 1));
            row++;

            string formula = method == VolumeMethod.AverageEndArea
                ? "V = (S₁ + S₂) / 2 × L" : "V = (S₁ + 4√(S₁·S₂) + S₂) / 6 × L";
            if (compare) formula = "So sánh AEA + Prismoidal";
            ws.Cell(row, 1).Value = $"Công thức: {formula}";
            ws.Cell(row, 1).Style.Font.Italic = true;
            ws.Cell(row, 1).Style.Font.FontColor = XLColor.FromArgb(100, 100, 100);
            row += 2;

            // Header Row 1: Nhóm Material (merge)
            int volHeaderRow1 = row;
            col = 1;
            foreach (var h in new[] { "STT", "Tên Cọc", "Lý Trình", "L (m)" })
            {
                StyleHeaderCell(ws.Cell(row, col), h);
                ws.Range(row, col, row + 1, col).Merge(); // Merge 2 hàng
                col++;
            }
            foreach (var mat in materials)
            {
                int startCol = col;
                ws.Range(row, startCol, row, startCol + colsPerMat - 1).Merge();
                StyleHeaderCell(ws.Cell(row, startCol), mat.MaterialName, true);

                // Sub-header row
                string[] subH = compare
                    ? new[] { "DT TB (m²)", "V_AEA (m³)", "V_Pris (m³)", "TL_AEA (m³)", "TL_Pris (m³)" }
                    : new[] { "DT TB (m²)", $"V_{methodTag} (m³)", "Tích lũy (m³)" };
                for (int i = 0; i < subH.Length; i++)
                    StyleSubHeaderCell(ws.Cell(row + 1, startCol + i), subH[i]);
                col += colsPerMat;
            }
            row += 2; // Qua 2 hàng header

            // Data Khối lượng
            int volDataStartRow = row;
            stt = 1;
            foreach (var r in results)
            {
                col = 1;
                ws.Cell(row, col++).Value = stt;
                ws.Cell(row, col++).Value = r.TenCoc;
                ws.Cell(row, col++).Value = r.LyTrinh;
                ws.Cell(row, col++).Value = Math.Round(r.DistanceToPrev, 3);

                int matIdx = 0;
                foreach (var mat in materials)
                {
                    if (stt == 1)
                    {
                        // Mặt cắt đầu: chưa có mặt cắt trước → KL = 0
                        ws.Cell(row, col++).Value = 0;
                        if (compare) { ws.Cell(row, col++).Value = 0; ws.Cell(row, col++).Value = 0; ws.Cell(row, col++).Value = 0; ws.Cell(row, col++).Value = 0; }
                        else { ws.Cell(row, col++).Value = 0; ws.Cell(row, col++).Value = 0; }
                    }
                    else
                    {
                        // DT TB lấy từ Mục A (cùng sheet)
                        int areaSrcCol = 5 + matIdx;
                        string aCol = GetExcelColLetter(areaSrcCol);
                        int aPrevRow = areaHeaderRow + stt - 1; // row trước trong Mục A
                        int aCurrRow = areaHeaderRow + stt;     // row hiện tại trong Mục A

                        // DT Trung bình = (DT_prev + DT_curr) / 2
                        ws.Cell(row, col).FormulaA1 = $"=({aCol}{aPrevRow}+{aCol}{aCurrRow})/2";
                        int dtTbCol = col; col++;

                        string dtLetter = GetExcelColLetter(dtTbCol);
                        string lLetter = GetExcelColLetter(4);

                        if (compare)
                        {
                            // V_AEA = DT_TB × L
                            ws.Cell(row, col).FormulaA1 = $"={dtLetter}{row}*{lLetter}{row}";
                            int vAeaCol = col; col++;
                            // V_Pris
                            ws.Cell(row, col++).Value = Math.Round(r.MaterialVolumes_Prismoidal.GetValueOrDefault(mat.MaterialName, 0), 3);
                            // TL_AEA
                            string vAeaL = GetExcelColLetter(vAeaCol);
                            string tlAeaL = GetExcelColLetter(col);
                            ws.Cell(row, col).FormulaA1 = $"={tlAeaL}{row - 1}+{vAeaL}{row}";
                            col++;
                            // TL_Pris
                            ws.Cell(row, col++).Value = Math.Round(r.CumulativeVolumes_Prismoidal.GetValueOrDefault(mat.MaterialName, 0), 3);
                        }
                        else
                        {
                            // V = DT_TB × L
                            ws.Cell(row, col).FormulaA1 = $"={dtLetter}{row}*{lLetter}{row}";
                            int vCol = col; col++;
                            // TL
                            string vL = GetExcelColLetter(vCol);
                            string tlL = GetExcelColLetter(col);
                            ws.Cell(row, col).FormulaA1 = $"={tlL}{row - 1}+{vL}{row}";
                            col++;
                        }
                    }
                    matIdx++;
                }

                if (stt % 2 == 0)
                    ws.Range(row, 1, row, volTotalCols).Style.Fill.BackgroundColor = XLColor.FromArgb(240, 245, 250);
                row++; stt++;
            }

            // Dòng TỔNG
            ws.Cell(row, 1).Value = "TỔNG";
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Range(row, 1, row, 4).Merge();
            ws.Range(row, 1, row, volTotalCols).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 248, 220);
            ws.Range(row, 1, row, volTotalCols).Style.Font.Bold = true;
            col = 5;
            foreach (var mat in materials)
            {
                col++; // Skip DT TB
                if (compare)
                {
                    string l1 = GetExcelColLetter(col);
                    ws.Cell(row, col).FormulaA1 = $"=SUM({l1}{volDataStartRow}:{l1}{row - 1})"; col++;
                    string l2 = GetExcelColLetter(col);
                    ws.Cell(row, col).FormulaA1 = $"=SUM({l2}{volDataStartRow}:{l2}{row - 1})"; col++;
                    col++; col++; // Skip TL
                }
                else
                {
                    string l1 = GetExcelColLetter(col);
                    ws.Cell(row, col).FormulaA1 = $"=SUM({l1}{volDataStartRow}:{l1}{row - 1})"; col++;
                    col++; // Skip TL
                }
            }

            // Border Mục B
            ws.Range(volHeaderRow1, 1, row, volTotalCols).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            ws.Range(volHeaderRow1, 1, row, volTotalCols).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            for (int c = 4; c <= volTotalCols; c++)
                ws.Column(c).Style.NumberFormat.Format = "#,##0.000";

            // Format chung
            ws.Columns().AdjustToContents();
            ws.Column(1).Width = 6;
            ws.Column(2).Width = 16;
            ws.Column(3).Width = 16;
        }

        // ───────────────────────────────────────────────────────────
        // Sheet Tổng hợp KL toàn bộ các tuyến
        // ───────────────────────────────────────────────────────────

        private static void WriteSummarySheet(XLWorkbook wb,
            List<(string AlignmentName, List<SectionVolumeResult> Results,
                List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> Materials)> allData,
            VolumeMethod method, bool compare, string methodTag)
        {
            var ws = wb.Worksheets.Add("Tổng hợp KL");

            ws.Cell(1, 1).Value = "BẢNG TỔNG HỢP KHỐI LƯỢNG TOÀN BỘ CÁC TUYẾN";
            ws.Range(1, 1, 1, 6).Merge();
            StyleSectionTitle(ws.Cell(1, 1));

            string methodText = compare ? "So sánh AEA + Prismoidal"
                : (method == VolumeMethod.AverageEndArea ? "Average End Area" : "Prismoidal");
            ws.Cell(2, 1).Value = $"Phương pháp: {methodText}";
            ws.Cell(2, 1).Style.Font.Italic = true;
            ws.Cell(2, 4).Value = $"Số tuyến: {allData.Count}";

            // Tìm tất cả Material names (union)
            var allMatNames = allData.SelectMany(d => d.Materials.Select(m => m.MaterialName)).Distinct().ToList();

            // Header
            int row = 4;
            int col = 1;
            StyleHeaderCell(ws.Cell(row, col++), "STT");
            StyleHeaderCell(ws.Cell(row, col++), "Tuyến");
            StyleHeaderCell(ws.Cell(row, col++), "Số mặt cắt");
            foreach (var mn in allMatNames)
            {
                if (compare)
                {
                    StyleHeaderCell(ws.Cell(row, col++), $"{mn}_AEA (m³)", true);
                    StyleHeaderCell(ws.Cell(row, col++), $"{mn}_Pris (m³)", true);
                }
                else
                {
                    StyleHeaderCell(ws.Cell(row, col++), $"{mn}_{methodTag} (m³)", true);
                }
            }
            int totalCols = col - 1;
            row++;

            // Data cho mỗi tuyến
            int dataStartRow = row;
            int idx = 1;
            foreach (var (aName, results, materials) in allData)
            {
                col = 1;
                ws.Cell(row, col++).Value = idx;
                ws.Cell(row, col++).Value = aName;
                ws.Cell(row, col++).Value = results.Count;

                foreach (var mn in allMatNames)
                {
                    double totalAEA = results.Sum(r => r.MaterialVolumes_AEA.GetValueOrDefault(mn, 0));
                    double totalPris = results.Sum(r => r.MaterialVolumes_Prismoidal.GetValueOrDefault(mn, 0));
                    if (compare)
                    {
                        ws.Cell(row, col++).Value = Math.Round(totalAEA, 3);
                        ws.Cell(row, col++).Value = Math.Round(totalPris, 3);
                    }
                    else
                    {
                        double v = method == VolumeMethod.AverageEndArea ? totalAEA : totalPris;
                        ws.Cell(row, col++).Value = Math.Round(v, 3);
                    }
                }
                if (idx % 2 == 0)
                    ws.Range(row, 1, row, totalCols).Style.Fill.BackgroundColor = XLColor.FromArgb(240, 245, 250);
                row++; idx++;
            }

            // Dòng TỔNG
            ws.Cell(row, 1).Value = "TỔNG";
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Range(row, 1, row, 3).Merge();
            ws.Range(row, 1, row, totalCols).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 248, 220);
            ws.Range(row, 1, row, totalCols).Style.Font.Bold = true;
            for (int c = 4; c <= totalCols; c++)
            {
                string cL = GetExcelColLetter(c);
                ws.Cell(row, c).FormulaA1 = $"=SUM({cL}{dataStartRow}:{cL}{row - 1})";
            }

            // Border & format
            ws.Range(4, 1, row, totalCols).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            ws.Range(4, 1, row, totalCols).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            for (int c = 4; c <= totalCols; c++)
                ws.Column(c).Style.NumberFormat.Format = "#,##0.000";

            // Diễn giải
            row += 2;
            ws.Cell(row, 1).Value = "DIỄN GIẢI";
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 1).Style.Font.FontSize = 12;
            row++;
            ws.Cell(row++, 1).Value = "• AEA: V = (S₁ + S₂) / 2 × L  —  Average End Area";
            ws.Cell(row++, 1).Value = "• Prismoidal: V = (S₁ + 4·√(S₁·S₂) + S₂) / 6 × L";
            ws.Cell(row++, 1).Value = "• S₁, S₂ = diện tích 2 mặt cắt liên tiếp (m²)";
            ws.Cell(row++, 1).Value = "• L = khoảng cách giữa 2 mặt cắt (m)";

            ws.Columns().AdjustToContents();
            ws.Column(2).Width = Math.Max(ws.Column(2).Width, 25);
        }

        // ───────────────────────────────────────────────────────────
        // Style helpers
        // ───────────────────────────────────────────────────────────

        private static void StyleSectionTitle(IXLCell cell)
        {
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontSize = 13;
            cell.Style.Font.FontColor = XLColor.FromArgb(27, 42, 74);
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        private static void StyleHeaderCell(IXLCell cell, string text, bool accent = false)
        {
            cell.Value = text;
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = accent ? XLColor.FromArgb(46, 74, 122) : XLColor.FromArgb(27, 42, 74);
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        private static void StyleSubHeaderCell(IXLCell cell, string text)
        {
            cell.Value = text;
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontSize = 9;
            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(70, 100, 150);
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        private static string GetExcelColLetter(int colNumber)
        {
            string result = "";
            while (colNumber > 0)
            {
                colNumber--;
                result = (char)('A' + colNumber % 26) + result;
                colNumber /= 26;
            }
            return result;
        }

        // ═══════════════════════════════════════════════════════════
        // GHI TEXT LÊN BẢN VẼ
        // ═══════════════════════════════════════════════════════════

        private static void WriteToDrawing(Transaction tr, Editor ed,
            SampleLineGroup slg, List<SectionVolumeResult> results,
            List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> materials,
            string alignmentName, VolumeMethod method)
        {
            // Lấy Section View Group
            var svgCollection = slg.SectionViewGroups;
            if (svgCollection.Count == 0)
            {
                ed.WriteMessage("\n⚠️ Chưa có Section View Group để ghi text.");
                return;
            }
            var svGroup = svgCollection[0];
            var sectionViewIds = svGroup.GetSectionViewIds();

            // Tạo layer
            string layerName = $"KL_{alignmentName}";
            UtilitiesCAD.CCreateLayer(layerName);

            string methodTag = method == VolumeMethod.AverageEndArea ? "AEA" : "Pris";

            int written = 0;
            foreach (ObjectId svId in sectionViewIds)
            {
                var sectionView = tr.GetObject(svId, OpenMode.ForWrite) as SectionView;
                if (sectionView == null) continue;

                var sampleLineId = sectionView.SampleLineId;
                var sampleLine = tr.GetObject(sampleLineId, OpenMode.ForRead) as SampleLine;
                if (sampleLine == null) continue;

                // Tìm result tương ứng
                var result = results.FirstOrDefault(r => Math.Abs(r.Station - sampleLine.Station) < 0.01);
                if (result == null) continue;

                double x = sectionView.Location.X;
                double y = sectionView.Location.Y;

                // Ghi text cho mỗi material — đặt phía trên SectionView
                int matIndex = 0;
                foreach (var mat in materials)
                {
                    double area = result.MaterialAreas.GetValueOrDefault(mat.MaterialName, 0);
                    double vol = method == VolumeMethod.AverageEndArea
                        ? result.MaterialVolumes_AEA.GetValueOrDefault(mat.MaterialName, 0)
                        : result.MaterialVolumes_Prismoidal.GetValueOrDefault(mat.MaterialName, 0);

                    double textY = y + 8 + matIndex * 1.2;
                    Point3d labelPos = new Point3d(x - 5, textY, 0);
                    Point3d valuePos = new Point3d(x + 3, textY, 0);

                    UtilitiesCAD.CCreateTextWithOutPut(labelPos, 0.35,
                        $"{mat.MaterialName}:", layerName, "Standard");
                    UtilitiesCAD.CCreateTextWithOutPut(valuePos, 0.35,
                        $"S={area:F2}m² V={vol:F2}m³ ({methodTag})", layerName, "Standard");

                    matIndex++;
                    written++;
                }
            }

            ed.WriteMessage($"\n✅ Đã ghi {written} dòng text lên bản vẽ (layer: {layerName}).");
        }

        // ═══════════════════════════════════════════════════════════
        // HELPER
        // ═══════════════════════════════════════════════════════════

        private static string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double m = station % 1000;
            return $"Km{km}+{m:F2}";
        }
    }
}
