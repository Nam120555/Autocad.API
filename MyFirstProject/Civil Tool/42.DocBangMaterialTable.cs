// 42.DocBangMaterialTable.cs
// Lệnh đọc dữ liệu từ Material Table (Volume Table) trên SectionView
// Hỗ trợ 2 chế độ:
//   1. Chọn bảng trực tiếp trên bản vẽ (chọn Table entity)
//   2. Quét tất cả Material Table trong SectionView Group
// Xuất ra: Command Line + Clipboard + Excel (tùy chọn)

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadDb = Autodesk.AutoCAD.DatabaseServices;
using ATable = Autodesk.AutoCAD.DatabaseServices.Table;
using MyFirstProject.Extensions;
using CivSection = Autodesk.Civil.DatabaseServices.Section;

using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using ContentAlignment = System.Drawing.ContentAlignment;

[assembly: CommandClass(typeof(MyFirstProject.DocBangMaterialTable))]

namespace MyFirstProject
{
    // ══════════════════════════════════════════════════════════════
    // DATA CLASS: Kết quả đọc từ bảng
    // ══════════════════════════════════════════════════════════════

    /// <summary>
    /// Dữ liệu đọc được từ một Material Table trên SectionView
    /// </summary>
    public class MaterialTableData
    {
        public string SampleLineName { get; set; } = "";
        public double Station { get; set; }
        public string LyTrinh { get; set; } = "";

        /// <summary>
        /// Danh sách header cột (tên Material + Area/Volume...)
        /// </summary>
        public List<string> ColumnHeaders { get; set; } = new();

        /// <summary>
        /// Dữ liệu theo hàng: mỗi hàng là 1 material, mỗi cột là giá trị
        /// Row[i] = { MaterialName, Value1, Value2, ... }
        /// </summary>
        public List<List<string>> Rows { get; set; } = new();

        /// <summary>
        /// Diện tích theo tên material (parse từ bảng)
        /// </summary>
        public Dictionary<string, double> MaterialAreas { get; set; } = new();

        /// <summary>
        /// Khối lượng theo tên material (parse từ bảng)
        /// </summary>
        public Dictionary<string, double> MaterialVolumes { get; set; } = new();
    }

    // ══════════════════════════════════════════════════════════════
    // FORM: Hiển thị kết quả đọc bảng
    // ══════════════════════════════════════════════════════════════

    public class MaterialTableResultForm : CivilFormBase
    {
        private DataGridView dgvResults = null!;
        private WinLabel lblInfo = null!;
        private Button btnCopyClipboard = null!;
        private Button btnExportCSV = null!;
        private Button btnClose = null!;

        private List<MaterialTableData> allData;

        public MaterialTableResultForm(List<MaterialTableData> data)
        {
            allData = data;
            InitializeUI();
            LoadData();
        }

        private void InitializeUI()
        {
            InitBaseTheme("DỮ LIỆU MATERIAL TABLE", 900, 600);
            this.FormBorderStyle = FormBorderStyle.Sizable;

            var header = CreateHeader("DỮ LIỆU MATERIAL TABLE - TRẮC NGANG");
            this.Controls.Add(header);

            lblInfo = new WinLabel
            {
                Text = $"Tổng cộng: {allData.Count} trắc ngang",
                Font = new Font("Segoe UI", 9.5f),
                ForeColor = TextDark,
                AutoSize = false,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Top,
                Padding = new Padding(15, 0, 0, 0),
                BackColor = BgPanel
            };
            this.Controls.Add(lblInfo);
            lblInfo.BringToFront();

            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                BackColor = BgPanel
            };
            this.Controls.Add(mainPanel);
            mainPanel.BringToFront();

            dgvResults = new DataGridView
            {
                Dock = DockStyle.Fill,
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
            mainPanel.Controls.Add(dgvResults);

            // Bottom bar
            var pnlBottom = CreateBottomBar();
            this.Controls.Add(pnlBottom);

            btnCopyClipboard = CreateButton("📋 Copy Clipboard", 10, 8, 160, 36, pnlBottom);
            btnCopyClipboard.Click += BtnCopyClipboard_Click;

            btnExportCSV = CreateButton("📁 Xuất CSV", 180, 8, 120, 36, pnlBottom);
            btnExportCSV.Click += BtnExportCSV_Click;

            btnClose = CreatePrimaryButton("Đóng", 780, 8, 100, 36, pnlBottom);
            btnClose.Click += (s, e) => this.Close();
        }

        private void LoadData()
        {
            dgvResults.Columns.Clear();
            dgvResults.Rows.Clear();

            if (allData.Count == 0) return;

            // Gom tất cả tên material
            var allMaterials = new HashSet<string>();
            foreach (var data in allData)
            {
                foreach (var mat in data.MaterialAreas.Keys)
                    allMaterials.Add(mat);
            }

            // Tạo cột
            dgvResults.Columns.Add("STT", "STT");
            dgvResults.Columns.Add("TenCoc", "Tên Cọc");
            dgvResults.Columns.Add("LyTrinh", "Lý Trình");

            foreach (var mat in allMaterials)
            {
                dgvResults.Columns.Add($"Area_{mat}", $"DT {mat} (m²)");
            }

            dgvResults.Columns["STT"]!.Width = 50;
            dgvResults.Columns["TenCoc"]!.Width = 100;
            dgvResults.Columns["LyTrinh"]!.Width = 120;

            // Thêm dữ liệu
            int stt = 1;
            foreach (var data in allData.OrderBy(d => d.Station))
            {
                var row = new List<object> { stt++, data.SampleLineName, data.LyTrinh };
                foreach (var mat in allMaterials)
                {
                    double area = data.MaterialAreas.GetValueOrDefault(mat, 0);
                    row.Add(Math.Round(area, 4));
                }
                dgvResults.Rows.Add(row.ToArray());
            }

            lblInfo.Text = $"Tổng cộng: {allData.Count} trắc ngang | {allMaterials.Count} loại material";
        }

        private void BtnCopyClipboard_Click(object? sender, EventArgs e)
        {
            try
            {
                var sb = new StringBuilder();

                // Header
                var headers = new List<string>();
                foreach (DataGridViewColumn col in dgvResults.Columns)
                    headers.Add(col.HeaderText);
                sb.AppendLine(string.Join("\t", headers));

                // Data
                foreach (DataGridViewRow row in dgvResults.Rows)
                {
                    var cells = new List<string>();
                    foreach (DataGridViewCell cell in row.Cells)
                        cells.Add(cell.Value?.ToString() ?? "");
                    sb.AppendLine(string.Join("\t", cells));
                }

                Clipboard.SetText(sb.ToString());
                MessageBox.Show("Đã copy vào Clipboard!\nDán (Ctrl+V) vào Excel để xem.",
                    "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExportCSV_Click(object? sender, EventArgs e)
        {
            try
            {
                using SaveFileDialog sfd = new()
                {
                    Filter = "CSV Files|*.csv",
                    FileName = "MaterialTable_Data.csv",
                    Title = "Lưu file CSV"
                };

                if (sfd.ShowDialog() != DialogResult.OK) return;

                var sb = new StringBuilder();

                // Header
                var headers = new List<string>();
                foreach (DataGridViewColumn col in dgvResults.Columns)
                    headers.Add(col.HeaderText);
                sb.AppendLine(string.Join(",", headers));

                // Data
                foreach (DataGridViewRow row in dgvResults.Rows)
                {
                    var cells = new List<string>();
                    foreach (DataGridViewCell cell in row.Cells)
                        cells.Add(cell.Value?.ToString() ?? "");
                    sb.AppendLine(string.Join(",", cells));
                }

                File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                MessageBox.Show($"Đã xuất CSV:\n{sfd.FileName}",
                    "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    // ══════════════════════════════════════════════════════════════
    // LỆNH CHÍNH
    // ══════════════════════════════════════════════════════════════

    public class DocBangMaterialTable
    {
        /// <summary>
        /// Format lý trình theo dạng Km0+123.456
        /// </summary>
        private static string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double meters = station % 1000;
            return $"Km{km}+{meters:F3}";
        }

        // ═══════════════════════════════════════════════════════════
        // LỆNH 1: Đọc Material Table từ SectionView đã chọn
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// Đọc dữ liệu từ bảng Material Table (Volume Table) đặt trên SectionView.
        /// Chọn 1 SectionView → quét tất cả Volume Table trên đó → đọc text.
        /// </summary>
        [CommandMethod("CTSV_DocBangMaterialTable")]
        public static void CTSVDocBangMaterialTable()
        {
            var ed = A.Ed;

            ed.WriteMessage("\n╔════════════════════════════════════════════════════════════════╗");
            ed.WriteMessage("\n║   ĐỌC DỮ LIỆU TỪ MATERIAL TABLE TRÊN TRẮC NGANG              ║");
            ed.WriteMessage("\n╚════════════════════════════════════════════════════════════════╝");

            try
            {
                // Hỏi chế độ
                PromptKeywordOptions pko = new("\nChọn chế độ đọc: ");
                pko.Keywords.Add("ChonBang", "ChonBang", "Chọn bảng trực tiếp (B)");
                pko.Keywords.Add("TatCa", "TatCa", "Quét tất cả SectionView (T)");
                pko.Keywords.Add("API", "API", "Đọc từ QTO Material API (A)");
                pko.Keywords.Default = "API";
                PromptResult pr = ed.GetKeywords(pko);
                if (pr.Status != PromptStatus.OK) return;

                switch (pr.StringResult)
                {
                    case "ChonBang":
                        DocBangTrucTiep();
                        break;
                    case "TatCa":
                        DocTatCaSectionView();
                        break;
                    case "API":
                        DocTuQTOAPI();
                        break;
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        // ═══════════════════════════════════════════════════════════
        // CHẾ ĐỘ 1: CHỌN BẢNG TRỰC TIẾP (TABLE ENTITY)
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// Cho phép chọn Table entity trên bản vẽ → đọc text từng cell
        /// </summary>
        private static void DocBangTrucTiep()
        {
            var ed = A.Ed;

            // Cho chọn nhiều bảng
            ed.WriteMessage("\nChọn các bảng (Table) cần đọc dữ liệu:");

            TypedValue[] filter = new[]
            {
                new TypedValue((int)DxfCode.Start, "ACAD_TABLE")
            };
            SelectionFilter selFilter = new(filter);
            PromptSelectionResult selResult = ed.GetSelection(selFilter);

            if (selResult.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nKhông chọn được bảng nào!");
                return;
            }

            using Transaction tr = A.Db.TransactionManager.StartTransaction();

            int tableCount = 0;
            var allResults = new List<MaterialTableData>();

            foreach (ObjectId tableId in selResult.Value.GetObjectIds())
            {
                ATable? table = tr.GetObject(tableId, OpenMode.ForRead) as ATable;
                if (table == null) continue;

                tableCount++;
                ed.WriteMessage($"\n\n📊 Bảng #{tableCount}:");
                ed.WriteMessage($"\n   Số hàng: {table.Rows.Count}, Số cột: {table.Columns.Count}");

                var tableData = new MaterialTableData();
                tableData.SampleLineName = $"Table_{tableCount}";

                // Đọc header (thường hàng 0 hoặc hàng 1)
                int headerRow = 0;
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    try
                    {
                        string headerText = GetCellText(table, headerRow, c);
                        tableData.ColumnHeaders.Add(headerText);
                        ed.WriteMessage($"\n   Cột {c}: \"{headerText}\"");
                    }
                    catch { tableData.ColumnHeaders.Add($"Col{c}"); }
                }

                // Đọc dữ liệu từ hàng 1 trở đi (hoặc hàng 2 nếu có title)
                int dataStartRow = FindDataStartRow(table);
                ed.WriteMessage($"\n   Dữ liệu bắt đầu từ hàng: {dataStartRow}");

                for (int r = dataStartRow; r < table.Rows.Count; r++)
                {
                    var rowData = new List<string>();
                    var rowDisplay = new StringBuilder();
                    rowDisplay.Append($"\n   Hàng {r}: ");

                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        try
                        {
                            string cellText = GetCellText(table, r, c);
                            rowData.Add(cellText);
                            rowDisplay.Append($"[{cellText}] ");

                            // Parse giá trị diện tích (thường ở cột có header chứa "Area")
                            if (tableData.ColumnHeaders.Count > c)
                            {
                                string header = tableData.ColumnHeaders[c].ToUpperInvariant();
                                if (c == 0 || header.Contains("MATERIAL") || header.Contains("VẬT LIỆU") || header.Contains("NAME"))
                                {
                                    // Cột tên material → ghi nhận
                                }
                                else if (header.Contains("AREA") || header.Contains("DIỆN TÍCH") || header.Contains("DT"))
                                {
                                    // Parse diện tích
                                    if (rowData.Count > 0 && double.TryParse(cellText.Replace(",", "."), 
                                        System.Globalization.NumberStyles.Any,
                                        System.Globalization.CultureInfo.InvariantCulture, out double areaVal))
                                    {
                                        string matName = rowData.Count > 0 ? rowData[0] : $"Mat_{r}";
                                        tableData.MaterialAreas[matName] = areaVal;
                                    }
                                }
                            }
                        }
                        catch { rowData.Add(""); }
                    }

                    tableData.Rows.Add(rowData);
                    ed.WriteMessage(rowDisplay.ToString());
                }

                allResults.Add(tableData);
            }

            tr.Commit();

            // Hiển thị form kết quả
            if (allResults.Count > 0 && allResults.Any(r => r.Rows.Count > 0))
            {
                ed.WriteMessage($"\n\n✅ Đã đọc {tableCount} bảng, tổng {allResults.Sum(r => r.Rows.Count)} hàng dữ liệu");

                // Copy raw data vào clipboard
                CopyRawToClipboard(allResults);
            }
            else
            {
                ed.WriteMessage("\n⚠️ Không đọc được dữ liệu nào từ các bảng đã chọn.");
            }
        }

        // ═══════════════════════════════════════════════════════════
        // CHẾ ĐỘ 2: QUÉT TẤT CẢ SECTIONVIEW TRONG NHÓM
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// Chọn 1 SectionView → tìm nhóm SectionView → quét tất cả Volume Table
        /// </summary>
        private static void DocTatCaSectionView()
        {
            var ed = A.Ed;

            ed.WriteMessage("\nChọn 1 trắc ngang bất kỳ trong nhóm cần đọc:");

            PromptEntityOptions peo = new("\nChọn SectionView: ");
            peo.SetRejectMessage("\nĐây không phải SectionView!");
            peo.AddAllowedClass(typeof(SectionView), true);
            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK) return;

            using Transaction tr = A.Db.TransactionManager.StartTransaction();

            SectionView? sectionView = tr.GetObject(per.ObjectId, OpenMode.ForRead) as SectionView;
            if (sectionView == null) { ed.WriteMessage("\nKhông phải SectionView!"); return; }

            // Lấy SampleLine Group
            ObjectId sampleLineId = sectionView.SampleLineId;
            SampleLine? sampleLine = tr.GetObject(sampleLineId, OpenMode.ForRead) as SampleLine;
            if (sampleLine == null) { tr.Commit(); return; }

            ObjectId slgId = sampleLine.GroupId;
            SampleLineGroup? slg = tr.GetObject(slgId, OpenMode.ForRead) as SampleLineGroup;
            if (slg == null) { tr.Commit(); return; }

            ed.WriteMessage($"\n📐 SampleLineGroup: {slg.Name}");

            // Lấy SectionViewGroup
            SectionViewGroupCollection svgCollection = slg.SectionViewGroups;
            if (svgCollection.Count == 0)
            {
                ed.WriteMessage("\n⚠️ Không có SectionViewGroup nào!");
                tr.Commit();
                return;
            }

            var allResults = new List<MaterialTableData>();
            int totalTables = 0;

            foreach (SectionViewGroup svGroup in svgCollection)
            {
                ObjectIdCollection svIds = svGroup.GetSectionViewIds();
                ed.WriteMessage($"\n\n📁 SectionViewGroup: {svIds.Count} trắc ngang");

                foreach (ObjectId svId in svIds)
                {
                    SectionView? sv = tr.GetObject(svId, OpenMode.ForRead) as SectionView;
                    if (sv == null) continue;

                    SampleLine? sl = tr.GetObject(sv.SampleLineId, OpenMode.ForRead) as SampleLine;
                    if (sl == null) continue;

                    // Lấy Volume Tables từ SectionView
                    SectionViewVolumeTableGroup volumeTableGroup = sv.VolumeTables;
                    ObjectIdCollection tableIds = volumeTableGroup.GetVolumeTableIds();

                    if (tableIds.Count == 0) continue;

                    foreach (ObjectId tableId in tableIds)
                    {
                        try
                        {
                            ATable? table = tr.GetObject(tableId, OpenMode.ForRead) as ATable;
                            if (table == null) continue;

                            totalTables++;
                            var tableData = ReadTableData(table, sl.Name, sl.Station);
                            allResults.Add(tableData);

                            ed.WriteMessage($"\n   📊 {sl.Name} (Km{(int)(sl.Station / 1000)}+{sl.Station % 1000:F3}): " +
                                          $"{table.Rows.Count} hàng × {table.Columns.Count} cột");

                            // In chi tiết material areas
                            foreach (var kvp in tableData.MaterialAreas)
                            {
                                ed.WriteMessage($"\n      → {kvp.Key}: {kvp.Value:F4} m²");
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\n   ⚠️ Lỗi đọc bảng: {ex.Message}");
                        }
                    }
                }
            }

            tr.Commit();

            ed.WriteMessage($"\n\n═══════════════════════════════════════════════════════════════");
            ed.WriteMessage($"\n✅ HOÀN THÀNH: Đã đọc {totalTables} bảng từ {allResults.Count} trắc ngang");
            ed.WriteMessage($"\n═══════════════════════════════════════════════════════════════");

            if (allResults.Count > 0)
            {
                // Copy và hiển thị form
                CopyRawToClipboard(allResults);

                using var form = new MaterialTableResultForm(allResults);
                AcadApp.ShowModalDialog(form);
            }
        }

        // ═══════════════════════════════════════════════════════════
        // CHẾ ĐỘ 3: ĐỌC TỪ QTO MATERIAL API (CHÍNH XÁC NHẤT)
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// Đọc dữ liệu diện tích từ QTOMaterial API (không phải từ Table entity)
        /// Đây là cách chính xác nhất vì truy cập trực tiếp dữ liệu tính toán.
        /// </summary>
        [CommandMethod("CTSV_DocMaterialAPI")]
        public static void DocTuQTOAPI()
        {
            var ed = A.Ed;

            ed.WriteMessage("\n╔════════════════════════════════════════════════════════════════╗");
            ed.WriteMessage("\n║   ĐỌC DỮ LIỆU MATERIAL TỪ QTO API                             ║");
            ed.WriteMessage("\n╚════════════════════════════════════════════════════════════════╝");

            using Transaction tr = A.Db.TransactionManager.StartTransaction();

            try
            {
                // 1. Lấy danh sách Alignment có SampleLineGroup
                var alignmentList = new List<(string Name, ObjectId AlignmentId, ObjectId SLGId)>();
                CivilDocument cdoc = CivilApplication.ActiveDocument;

                foreach (ObjectId alId in cdoc.GetAlignmentIds())
                {
                    Alignment? al = tr.GetObject(alId, OpenMode.ForRead) as Alignment;
                    if (al == null) continue;

                    ObjectIdCollection slgIds = al.GetSampleLineGroupIds();
                    foreach (ObjectId slgId in slgIds)
                    {
                        SampleLineGroup? slg = tr.GetObject(slgId, OpenMode.ForRead) as SampleLineGroup;
                        if (slg != null && slg.MaterialLists.Count > 0)
                        {
                            alignmentList.Add(($"{al.Name} — {slg.Name} ({slg.MaterialLists.Count} Material List)", alId, slgId));
                        }
                    }
                }

                if (alignmentList.Count == 0)
                {
                    ed.WriteMessage("\n⚠️ Không tìm thấy Alignment có Material List!");
                    ed.WriteMessage("\n   → Hãy chạy Compute Materials trước.");
                    tr.Commit();
                    return;
                }

                // 2. Cho chọn Alignment
                ed.WriteMessage("\n\nDanh sách Alignment có Material List:");
                for (int i = 0; i < alignmentList.Count; i++)
                {
                    ed.WriteMessage($"\n   [{i + 1}] {alignmentList[i].Name}");
                }

                PromptIntegerOptions pio = new($"\nChọn Alignment (1-{alignmentList.Count}): ")
                {
                    LowerLimit = 1,
                    UpperLimit = alignmentList.Count,
                    DefaultValue = 1
                };
                PromptIntegerResult pir = ed.GetInteger(pio);
                if (pir.Status != PromptStatus.OK) { tr.Commit(); return; }

                int selectedIdx = pir.Value - 1;
                var (name, alignmentId, slgId2) = alignmentList[selectedIdx];

                // 3. Đọc dữ liệu từ QTO
                SampleLineGroup? selectedSlg = tr.GetObject(slgId2, OpenMode.ForRead) as SampleLineGroup;
                if (selectedSlg == null) { tr.Commit(); return; }

                Alignment? selectedAlignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;

                // Lấy Material Lists
                QTOMaterialListCollection materialLists = selectedSlg.MaterialLists;
                var materials = new List<(string Name, Guid ListGuid, Guid MatGuid, string ListName, string QuantityType)>();

                foreach (QTOMaterialList matList in materialLists)
                {
                    try
                    {
                        foreach (QTOMaterial mat in matList)
                        {
                            string qtype = mat.QuantityType switch
                            {
                                MaterialQuantityType.Cut => "Đào",
                                MaterialQuantityType.Fill => "Đắp",
                                MaterialQuantityType.CutFill => "Đào+Đắp",
                                MaterialQuantityType.Earthwork => "Earthwork",
                                MaterialQuantityType.Structure => "Kết cấu",
                                _ => "Khác"
                            };
                            materials.Add((mat.Name, matList.Guid, mat.Guid, matList.Name ?? "Unnamed", qtype));
                        }
                    }
                    catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                }

                if (materials.Count == 0)
                {
                    ed.WriteMessage("\n⚠️ Không tìm thấy Material nào!");
                    tr.Commit();
                    return;
                }

                ed.WriteMessage($"\n\n📋 Tìm thấy {materials.Count} Material:");
                foreach (var mat in materials)
                {
                    ed.WriteMessage($"\n   → [{mat.QuantityType}] {mat.Name} (List: {mat.ListName})");
                }

                // 4. Duyệt SampleLines và lấy diện tích
                ObjectIdCollection slIds = selectedSlg.GetSampleLineIds();
                var allResults = new List<MaterialTableData>();

                // Sort by station
                var sortedSampleLines = new List<SampleLine>();
                foreach (ObjectId slId in slIds)
                {
                    SampleLine? sl = tr.GetObject(slId, OpenMode.ForRead) as SampleLine;
                    if (sl != null) sortedSampleLines.Add(sl);
                }
                sortedSampleLines = sortedSampleLines.OrderBy(s => s.Station).ToList();

                ed.WriteMessage($"\n\n📐 Đang đọc {sortedSampleLines.Count} trắc ngang...\n");
                ed.WriteMessage($"\n{"STT",-5} {"Tên Cọc",-20} {"Lý Trình",-15}");
                foreach (var mat in materials)
                    ed.WriteMessage($" {mat.Name,-15}");
                ed.WriteMessage("\n" + new string('─', 60 + materials.Count * 15));

                int stt = 1;
                foreach (SampleLine sl in sortedSampleLines)
                {
                    var tableData = new MaterialTableData
                    {
                        SampleLineName = sl.Name,
                        Station = sl.Station,
                        LyTrinh = FormatStation(sl.Station)
                    };

                    ed.WriteMessage($"\n{stt,-5} {sl.Name,-20} {FormatStation(sl.Station),-15}");

                    foreach (var (matName, listGuid, matGuid, listName, qtype) in materials)
                    {
                        try
                        {
                            ObjectId materialSectionId = sl.GetMaterialSectionId(listGuid, matGuid);
                            double area = 0;

                            if (!materialSectionId.IsNull && materialSectionId.IsValid)
                            {
                                CivSection? section = tr.GetObject(materialSectionId, OpenMode.ForRead, false, true) as CivSection;
                                if (section != null)
                                {
                                    // Ưu tiên section.Area (API), fallback Shoelace
                                    area = section.Area;
                                    if (area <= 0)
                                    {
                                        area = CalculateShoelaceArea(section);
                                    }
                                }
                            }

                            tableData.MaterialAreas[matName] = Math.Round(area, 4);
                            ed.WriteMessage($" {area,14:F4}");
                        }
                        catch
                        {
                            tableData.MaterialAreas[matName] = 0;
                            ed.WriteMessage($" {"N/A",14}");
                        }
                    }

                    allResults.Add(tableData);
                    stt++;
                }

                tr.Commit();

                ed.WriteMessage($"\n\n═══════════════════════════════════════════════════════════════");
                ed.WriteMessage($"\n✅ HOÀN THÀNH: Đã đọc {allResults.Count} trắc ngang × {materials.Count} material");
                ed.WriteMessage($"\n═══════════════════════════════════════════════════════════════");

                // Hiển thị form kết quả
                if (allResults.Count > 0)
                {
                    CopyRawToClipboard(allResults);

                    using var form = new MaterialTableResultForm(allResults);
                    AcadApp.ShowModalDialog(form);
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
                ed.WriteMessage($"\n   Stack: {ex.StackTrace}");
            }
        }

        // ═══════════════════════════════════════════════════════════
        // HÀM TIỆN ÍCH
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// Lấy text từ cell, bỏ định dạng MText
        /// </summary>
        private static string GetCellText(ATable table, int row, int col)
        {
            try
            {
                // Thử GetTextString trước
                string? text = table.Cells[row, col].Value?.ToString();
                if (!string.IsNullOrEmpty(text))
                {
                    // Bỏ mã MText formatting
                    text = System.Text.RegularExpressions.Regex.Replace(text, @"\{\\[^;]+;([^}]+)\}", "$1");
                    text = System.Text.RegularExpressions.Regex.Replace(text, @"\\[A-Za-z][^;]*;", "");
                    return text.Trim();
                }

                // Fallback: thử đọc TextString
                return table.Cells[row, col].TextString?.Trim() ?? "";
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Tìm hàng bắt đầu dữ liệu (bỏ qua title/header)
        /// </summary>
        private static int FindDataStartRow(ATable table)
        {
            // Kiểm tra xem hàng 0 có phải title không
            // Title thường merge tất cả cột
            try
            {
                // Hàng 0 thường là title, hàng 1 là header, hàng 2+ là data
                if (table.Rows.Count > 2)
                {
                    string cell00 = GetCellText(table, 0, 0);
                    string cell10 = GetCellText(table, 1, 0);

                    // Nếu hàng 1 có nhiều cột populated → hàng 1 là header → data từ hàng 2
                    int populatedCols = 0;
                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        if (!string.IsNullOrEmpty(GetCellText(table, 1, c)))
                            populatedCols++;
                    }

                    if (populatedCols > 1)
                        return 2; // Title + Header → Data từ hàng 2
                }
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

            return 1; // Mặc định: Header ở hàng 0, Data từ hàng 1
        }

        /// <summary>
        /// Đọc dữ liệu từ 1 Table entity và map vào MaterialTableData
        /// </summary>
        private static MaterialTableData ReadTableData(ATable table, string sampleLineName, double station)
        {
            var data = new MaterialTableData
            {
                SampleLineName = sampleLineName,
                Station = station,
                LyTrinh = FormatStation(station)
            };

            try
            {
                // Đọc headers
                int headerRow = 0;
                // Tìm header row (có nhiều cột populated nhất)
                int maxPopulated = 0;
                for (int r = 0; r < Math.Min(3, table.Rows.Count); r++)
                {
                    int count = 0;
                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        if (!string.IsNullOrEmpty(GetCellText(table, r, c)))
                            count++;
                    }
                    if (count > maxPopulated)
                    {
                        maxPopulated = count;
                        headerRow = r;
                    }
                }

                for (int c = 0; c < table.Columns.Count; c++)
                {
                    data.ColumnHeaders.Add(GetCellText(table, headerRow, c));
                }

                // Đọc data rows
                int dataStart = headerRow + 1;
                for (int r = dataStart; r < table.Rows.Count; r++)
                {
                    var rowData = new List<string>();
                    string materialName = "";

                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        string cellText = GetCellText(table, r, c);
                        rowData.Add(cellText);

                        // Heuristic: cột đầu tiên thường là tên material
                        if (c == 0)
                        {
                            materialName = cellText;
                        }
                        else if (!string.IsNullOrEmpty(materialName) && !string.IsNullOrEmpty(cellText))
                        {
                            // Thử parse số
                            string cleanText = cellText
                                .Replace("m²", "").Replace("m³", "")
                                .Replace("sq.m", "").Replace("cu.m", "")
                                .Trim();

                            if (double.TryParse(cleanText.Replace(",", "."),
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture, out double numVal))
                            {
                                // Nếu header chứa Area → Material Area
                                string header = data.ColumnHeaders.Count > c ? data.ColumnHeaders[c].ToUpperInvariant() : "";
                                if (header.Contains("AREA") || header.Contains("DIỆN TÍCH") || header.Contains("DT") ||
                                    header.Contains("CUT") || header.Contains("FILL") || header.Contains("ĐÀO") || header.Contains("ĐẮP"))
                                {
                                    if (!data.MaterialAreas.ContainsKey(materialName))
                                        data.MaterialAreas[materialName] = 0;
                                    data.MaterialAreas[materialName] += numVal;
                                }
                            }
                        }
                    }

                    data.Rows.Add(rowData);
                }
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

            return data;
        }

        /// <summary>
        /// Tính diện tích bằng Shoelace Formula từ SectionPoints
        /// </summary>
        private static double CalculateShoelaceArea(CivSection section)
        {
            try
            {
                SectionPointCollection pts = section.SectionPoints;
                if (pts.Count < 3) return 0;

                double area = 0;
                int n = pts.Count;
                for (int i = 0; i < n; i++)
                {
                    int j = (i + 1) % n;
                    area += pts[i].Location.X * pts[j].Location.Y;
                    area -= pts[j].Location.X * pts[i].Location.Y;
                }
                return Math.Abs(area) / 2.0;
            }
            catch { return 0; }
        }

        /// <summary>
        /// Copy dữ liệu thô vào clipboard (tab-separated, paste vào Excel)
        /// </summary>
        private static void CopyRawToClipboard(List<MaterialTableData> results)
        {
            try
            {
                var sb = new StringBuilder();

                // Gom tất cả tên material
                var allMats = new HashSet<string>();
                foreach (var r in results)
                    foreach (var m in r.MaterialAreas.Keys)
                        allMats.Add(m);

                // Header
                sb.Append("STT\tTên Cọc\tLý Trình");
                foreach (var mat in allMats)
                    sb.Append($"\tDT {mat} (m²)");
                sb.AppendLine();

                // Data
                int stt = 1;
                foreach (var r in results.OrderBy(r => r.Station))
                {
                    sb.Append($"{stt}\t{r.SampleLineName}\t{r.LyTrinh}");
                    foreach (var mat in allMats)
                    {
                        double area = r.MaterialAreas.GetValueOrDefault(mat, 0);
                        sb.Append($"\t{area:F4}");
                    }
                    sb.AppendLine();
                    stt++;
                }

                Clipboard.SetText(sb.ToString());
                A.Ed.WriteMessage("\n📋 Đã copy dữ liệu vào Clipboard! Dán (Ctrl+V) vào Excel.");
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
        }
    }
}
