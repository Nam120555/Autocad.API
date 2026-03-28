// 39.TinhKhoiLuongTracNgang.cs - Lệnh CTSV_TinhKhoiLuongTracNgang
// Tổng hợp: Tính diện tích Shape (đào/đắp/vật liệu) từ Material Section
//            rồi tính khối lượng theo phương pháp trung bình 2 mặt cắt
// Xuất ra: Command Line + Clipboard + CSV (tùy chọn)

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
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
    // DATA CLASS: Kết quả tại mỗi trắc ngang
    // ══════════════════════════════════════════════════════════════

    public class SectionVolumeResult
    {
        public string TenCoc { get; set; } = "";
        public double Station { get; set; }
        public string LyTrinh { get; set; } = "";

        // Diện tích tại mặt cắt (m²)
        public Dictionary<string, double> MaterialAreas { get; set; } = new();

        // Khối lượng trung bình giữa 2 mặt cắt (m³) — tính sau
        public Dictionary<string, double> MaterialVolumes { get; set; } = new();

        // Khoảng cách đến trắc ngang tiếp theo (m)
        public double DistanceToNext { get; set; }
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
        private CheckBox chkExportCSV = null!;
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
        public bool ExportCSV => chkExportCSV.Checked;
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
                                      mat.QuantityType == MaterialQuantityType.Fill ? "🔺 Đắp" : "📦 Khác";
                        materialList.Add((mat.Name, matList.Name, matList.Guid, mat.Guid));
                        int idx = chkMaterials.Items.Add($"{type}: {mat.Name} ({matList.Name})");
                        chkMaterials.SetItemChecked(idx, true);
                    }
                }
                catch { }
            }
            tr.Commit();

            if (materialList.Count == 0)
            {
                chkMaterials.Items.Add("⚠️ Chưa có Material List. Hãy chạy CTSV_VeTracNgangThietKe trước.");
            }
        }

        private void InitializeComponent()
        {
            InitBaseTheme("📐  TÍNH KHỐI LƯỢNG TRẮC NGANG", 480, 490);

            // ═══ HEADER ═══
            var header = CreateHeader("TÍNH KHỐI LƯỢNG TRẮC NGANG");
            this.Controls.Add(header);

            var subtitle = CreateSubtitle("Lấy diện tích Shape Material → Tính KL đào đắp (trung bình 2 MC)");
            this.Controls.Add(subtitle);

            // ═══ CHỌN ALIGNMENT ═══
            var grpAlign = CreateGroupBox("📏 Chọn Alignment + SampleLine Group", 16, 82, 440, 64);
            this.Controls.Add(grpAlign);

            cboAlignment = CreateComboBox(12, 26, 416);
            foreach (var (name, _, _) in alignmentList)
                cboAlignment.Items.Add(name);
            if (cboAlignment.Items.Count > 0)
                cboAlignment.SelectedIndex = 0;
            cboAlignment.SelectedIndexChanged += (s, e) => LoadMaterials();
            grpAlign.Controls.Add(cboAlignment);

            // ═══ CHỌN MATERIAL ═══
            var grpMat = CreateGroupBox("📦 Chọn Material cần tính", 16, 154, 440, 170);
            this.Controls.Add(grpMat);

            chkMaterials = new CheckedListBox
            {
                Location = new Point(12, 26),
                Size = new Size(416, 130),
                CheckOnClick = true,
                Font = new Font("Segoe UI", 9.5f),
                BackColor = BgWhite,
                ForeColor = TextDark,
                BorderStyle = BorderStyle.FixedSingle
            };
            grpMat.Controls.Add(chkMaterials);

            // ═══ TÙY CHỌN XUẤT ═══
            var grpExport = CreateGroupBox("📤 Tùy chọn xuất", 16, 332, 440, 60);
            this.Controls.Add(grpExport);

            chkExportCSV = new CheckBox
            {
                Text = "📄 Xuất file CSV",
                Location = new Point(12, 26),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpExport.Controls.Add(chkExportCSV);

            chkWriteToDrawing = new CheckBox
            {
                Text = "✏️ Ghi text lên trắc ngang",
                Location = new Point(220, 26),
                AutoSize = true,
                Checked = false,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpExport.Controls.Add(chkWriteToDrawing);

            // ═══ BOTTOM BAR ═══
            var bottomBar = CreateBottomBar(56);
            this.Controls.Add(bottomBar);

            var btnCancel = CreateButton("✕ Hủy", 10, 10, 210, 36, bottomBar);
            btnCancel.Click += (s, e) => { IsConfirmed = false; this.Close(); };

            var btnRun = CreateSuccessButton("📐 Tính Khối Lượng", 230, 10, 210, 36, bottomBar);
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
        /// Diện tích Shape → Khối lượng (phương pháp trung bình 2 mặt cắt)
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
            bool exportCSV = form.ExportCSV;
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

                    // Tính diện tích cho mỗi Material
                    foreach (var (matName, listName, listGuid, matGuid) in selectedMaterials)
                    {
                        double area = GetMaterialArea(tr, sl, listGuid, matGuid);
                        result.MaterialAreas[matName] = area;
                    }

                    results.Add(result);
                }

                // Sắp xếp theo station
                results = results.OrderBy(r => r.Station).ToList();

                // 3. Tính khối lượng (phương pháp trung bình 2 mặt cắt)
                for (int i = 0; i < results.Count; i++)
                {
                    if (i < results.Count - 1)
                    {
                        double dist = results[i + 1].Station - results[i].Station;
                        results[i].DistanceToNext = dist;

                        foreach (var matName in results[i].MaterialAreas.Keys)
                        {
                            double area1 = results[i].MaterialAreas[matName];
                            double area2 = results[i + 1].MaterialAreas.ContainsKey(matName)
                                ? results[i + 1].MaterialAreas[matName] : 0;
                            // V = (S1 + S2) / 2 × L
                            results[i].MaterialVolumes[matName] = (area1 + area2) / 2.0 * dist;
                        }
                    }
                    else
                    {
                        results[i].DistanceToNext = 0;
                        foreach (var matName in results[i].MaterialAreas.Keys)
                            results[i].MaterialVolumes[matName] = 0;
                    }
                }

                // 4. Hiển thị kết quả
                DisplayResults(ed, alignment.Name, results, selectedMaterials);

                // 5. Copy vào clipboard
                CopyToClipboard(alignment.Name, results, selectedMaterials);

                // 6. Xuất CSV nếu được chọn
                if (exportCSV)
                    ExportToCSV(ed, alignment.Name, results, selectedMaterials);

                // 7. Ghi text lên bản vẽ nếu được chọn
                if (writeToDrawing)
                    WriteToDrawing(tr, ed, slg, results, selectedMaterials, alignment.Name);

                tr.Commit();

                ed.WriteMessage("\n\n✅ HOÀN THÀNH tính khối lượng trắc ngang!");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        // ═══ LẤY DIỆN TÍCH MATERIAL SECTION ═══

        /// <summary>
        /// Lấy diện tích Material Section bằng Shoelace Formula
        /// Đọc SectionPoints từ Material Section (closed polygon) → tính diện tích
        /// </summary>
        private static double GetMaterialArea(Transaction tr, SampleLine sampleLine, Guid listGuid, Guid matGuid)
        {
            try
            {
                ObjectId materialSectionId = sampleLine.GetMaterialSectionId(listGuid, matGuid);
                if (materialSectionId.IsNull || !materialSectionId.IsValid) return 0;

                var section = tr.GetObject(materialSectionId, OpenMode.ForRead) as Section;
                if (section == null) return 0;

                SectionPointCollection sectionPoints = section.SectionPoints;
                if (sectionPoints.Count < 3) return 0;

                // Tính diện tích bằng công thức Shoelace (Gauss)
                // |A| = 0.5 × |Σ(x_i × y_{i+1} - x_{i+1} × y_i)|
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

        // ═══ HIỂN THỊ KẾT QUẢ ═══

        private static void DisplayResults(Editor ed, string alignmentName,
            List<SectionVolumeResult> results,
            List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> materials)
        {
            ed.WriteMessage("\n");
            ed.WriteMessage($"\n╔══════════════════════════════════════════════════════════════════════════╗");
            ed.WriteMessage($"\n║  📐 BẢNG KHỐI LƯỢNG TRẮC NGANG — {alignmentName,-38} ║");
            ed.WriteMessage($"\n╚══════════════════════════════════════════════════════════════════════════╝");

            // Header
            string header = $"\n{"STT",-5} {"Tên Cọc",-15} {"Lý Trình",-14} {"L(m)",-8}";
            foreach (var mat in materials)
            {
                string shortName = mat.MaterialName.Length > 10 ? mat.MaterialName.Substring(0, 10) : mat.MaterialName;
                header += $" {"S_" + shortName,-12} {"V_" + shortName,-12}";
            }
            ed.WriteMessage(header);
            ed.WriteMessage($"\n{new string('─', 50 + materials.Count * 24)}");

            // Data
            var totalVolumes = new Dictionary<string, double>();
            foreach (var mat in materials)
                totalVolumes[mat.MaterialName] = 0;

            int stt = 1;
            foreach (var r in results)
            {
                string line = $"\n{stt,-5} {r.TenCoc,-15} {r.LyTrinh,-14} {r.DistanceToNext,-8:F2}";

                foreach (var mat in materials)
                {
                    double area = r.MaterialAreas.ContainsKey(mat.MaterialName) ? r.MaterialAreas[mat.MaterialName] : 0;
                    double vol = r.MaterialVolumes.ContainsKey(mat.MaterialName) ? r.MaterialVolumes[mat.MaterialName] : 0;
                    line += $" {area,-12:F3} {vol,-12:F3}";
                    totalVolumes[mat.MaterialName] += vol;
                }

                ed.WriteMessage(line);
                stt++;
            }

            // Tổng
            ed.WriteMessage($"\n{new string('═', 50 + materials.Count * 24)}");
            string totalLine = $"\n{"TỔNG",-5} {"",-15} {"",-14} {"",-8}";
            foreach (var mat in materials)
            {
                totalLine += $" {"",-12} {totalVolumes[mat.MaterialName],-12:F3}";
            }
            ed.WriteMessage(totalLine);

            ed.WriteMessage($"\n{new string('═', 50 + materials.Count * 24)}");

            // Tóm tắt
            ed.WriteMessage("\n\n📊 TÓM TẮT:");
            foreach (var mat in materials)
            {
                ed.WriteMessage($"\n   {mat.MaterialName}: {totalVolumes[mat.MaterialName]:N2} m³");
            }
        }

        // ═══ COPY VÀO CLIPBOARD ═══

        private static void CopyToClipboard(string alignmentName,
            List<SectionVolumeResult> results,
            List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> materials)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"KHỐI LƯỢNG TRẮC NGANG — {alignmentName}");

            // Header
            sb.Append("STT\tTên Cọc\tLý Trình\tKhoảng cách (m)");
            foreach (var mat in materials)
                sb.Append($"\tDT_{mat.MaterialName} (m²)\tKL_{mat.MaterialName} (m³)");
            sb.AppendLine();

            // Data
            int stt = 1;
            foreach (var r in results)
            {
                sb.Append($"{stt}\t{r.TenCoc}\t{r.LyTrinh}\t{r.DistanceToNext:F2}");
                foreach (var mat in materials)
                {
                    double area = r.MaterialAreas.ContainsKey(mat.MaterialName) ? r.MaterialAreas[mat.MaterialName] : 0;
                    double vol = r.MaterialVolumes.ContainsKey(mat.MaterialName) ? r.MaterialVolumes[mat.MaterialName] : 0;
                    sb.Append($"\t{area:F3}\t{vol:F3}");
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

        // ═══ XUẤT CSV ═══

        private static void ExportToCSV(Editor ed, string alignmentName,
            List<SectionVolumeResult> results,
            List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> materials)
        {
            var sfd = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                Title = "Lưu file khối lượng trắc ngang",
                FileName = $"KL_TracNgang_{alignmentName}.csv"
            };

            if (sfd.ShowDialog() != DialogResult.OK) return;

            var sb = new StringBuilder();
            sb.Append("STT,TenCoc,LyTrinh,Station,KhoangCach_m");
            foreach (var mat in materials)
                sb.Append($",DT_{mat.MaterialName}_m2,KL_{mat.MaterialName}_m3");
            sb.AppendLine();

            int stt = 1;
            foreach (var r in results)
            {
                sb.Append($"{stt},{r.TenCoc},{r.LyTrinh},{r.Station:F3},{r.DistanceToNext:F3}");
                foreach (var mat in materials)
                {
                    double area = r.MaterialAreas.ContainsKey(mat.MaterialName) ? r.MaterialAreas[mat.MaterialName] : 0;
                    double vol = r.MaterialVolumes.ContainsKey(mat.MaterialName) ? r.MaterialVolumes[mat.MaterialName] : 0;
                    sb.Append($",{area:F3},{vol:F3}");
                }
                sb.AppendLine();
                stt++;
            }

            // Dòng tổng
            sb.Append($"TỔNG,,,,");
            foreach (var mat in materials)
            {
                double totalVol = results.Sum(r =>
                    r.MaterialVolumes.ContainsKey(mat.MaterialName) ? r.MaterialVolumes[mat.MaterialName] : 0);
                sb.Append($",,{totalVol:F3}");
            }
            sb.AppendLine();

            File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
            ed.WriteMessage($"\n✅ Đã xuất file: {sfd.FileName}");
        }

        // ═══ GHI TEXT LÊN BẢN VẼ ═══

        private static void WriteToDrawing(Transaction tr, Editor ed,
            SampleLineGroup slg, List<SectionVolumeResult> results,
            List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> materials,
            string alignmentName)
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
                    double area = result.MaterialAreas.ContainsKey(mat.MaterialName)
                        ? result.MaterialAreas[mat.MaterialName] : 0;
                    double vol = result.MaterialVolumes.ContainsKey(mat.MaterialName)
                        ? result.MaterialVolumes[mat.MaterialName] : 0;

                    double textY = y + 8 + matIndex * 1.2; // Offset trên Section View
                    Point3d labelPos = new Point3d(x - 5, textY, 0);
                    Point3d valuePos = new Point3d(x + 3, textY, 0);

                    UtilitiesCAD.CCreateTextWithOutPut(labelPos, 0.35, $"{mat.MaterialName}:", layerName, "Standard");
                    UtilitiesCAD.CCreateTextWithOutPut(valuePos, 0.35, $"S={area:F2}m² V={vol:F2}m³", layerName, "Standard");

                    matIndex++;
                    written++;
                }
            }

            ed.WriteMessage($"\n✅ Đã ghi {written} dòng text lên bản vẽ (layer: {layerName}).");
        }

        // ═══ HELPER ═══

        private static string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double m = station % 1000;
            return $"Km{km}+{m:F2}";
        }
    }
}
