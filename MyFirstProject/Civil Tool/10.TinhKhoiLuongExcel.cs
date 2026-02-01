// (C) Copyright 2024
// Tính khối lượng vật liệu từ Material Section và xuất ra Excel
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Acad = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using AcadDb = Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

using Civil = Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using Autodesk.Civil.Runtime;
using Autodesk.Civil.Settings;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Civil.ApplicationServices;
using CivSection = Autodesk.Civil.DatabaseServices.Section;
using FormLabel = System.Windows.Forms.Label;

using ClosedXML.Excel;
using MyFirstProject.Extensions;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(Civil3DCsharp.TinhKhoiLuongExcel))]

namespace Civil3DCsharp
{
    #region Data Classes

    /// <summary>
    /// Thông tin diện tích vật liệu tại một cọc
    /// </summary>
    public class MaterialVolumeInfo
    {
        public string StakeName { get; set; } = "";       // Tên cọc: "Cọc 1"
        public string Station { get; set; } = "";          // Lý trình: "Km0+123.456"
        public double StationValue { get; set; }          // Giá trị lý trình: 123.456
        public string MaterialName { get; set; } = "";    // Tên vật liệu: "Đào đất"
        public double Area { get; set; }                  // Diện tích (m²)
    }

    /// <summary>
    /// Thông tin tổng hợp của một cọc
    /// </summary>
    public class StakeInfo
    {
        public string Station { get; set; } = "";         // Lý trình format: "Km0+123.456"
        public string StakeName { get; set; } = "";       // Tên cọc
        public double StationValue { get; set; }          // Giá trị lý trình số
        public double SpacingPrev { get; set; }           // Khoảng cách đến cọc trước
        public Dictionary<string, double> MaterialAreas { get; set; } = new();  // {MaterialName: Area}
        public Dictionary<string, double> MaterialVolumes { get; set; } = new(); // {MaterialName: Volume}
    }

    /// <summary>
    /// Thông tin Alignment có SampleLineGroup
    /// </summary>
    public class AlignmentInfo
    {
        public ObjectId AlignmentId { get; set; }
        public string Name { get; set; } = "";
        public ObjectId SampleLineGroupId { get; set; }
        public bool IsSelected { get; set; }
    }

    #endregion

    #region Forms

    /// <summary>
    /// Form chọn Alignments để tính khối lượng
    /// </summary>
    public class FormChonAlignment : Form
    {
        private CheckedListBox checkedListBox;
        private Button btnOK;
        private Button btnCancel;
        private Button btnSelectAll;
        private Button btnDeselectAll;
        private List<AlignmentInfo> alignmentInfos;

        public List<AlignmentInfo> SelectedAlignments => alignmentInfos.Where(a => a.IsSelected).ToList();

        public FormChonAlignment(List<AlignmentInfo> alignments)
        {
            alignmentInfos = alignments;
            InitializeComponent();
            LoadAlignments();
        }

        private void InitializeComponent()
        {
            this.Text = "Chọn Alignments để tính khối lượng";
            this.Size = new System.Drawing.Size(450, 400);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            FormLabel lblTitle = new()
            {
                Text = "Chọn các Alignment có SampleLineGroup:",
                Location = new System.Drawing.Point(12, 12),
                Size = new System.Drawing.Size(400, 20)
            };
            this.Controls.Add(lblTitle);

            checkedListBox = new CheckedListBox
            {
                Location = new System.Drawing.Point(12, 35),
                Size = new System.Drawing.Size(410, 250),
                CheckOnClick = true
            };
            checkedListBox.ItemCheck += CheckedListBox_ItemCheck;
            this.Controls.Add(checkedListBox);

            btnSelectAll = new Button
            {
                Text = "Chọn tất cả",
                Location = new System.Drawing.Point(12, 295),
                Size = new System.Drawing.Size(100, 28)
            };
            btnSelectAll.Click += (s, e) => SelectAll(true);
            this.Controls.Add(btnSelectAll);

            btnDeselectAll = new Button
            {
                Text = "Bỏ chọn tất cả",
                Location = new System.Drawing.Point(120, 295),
                Size = new System.Drawing.Size(100, 28)
            };
            btnDeselectAll.Click += (s, e) => SelectAll(false);
            this.Controls.Add(btnDeselectAll);

            btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(240, 330),
                Size = new System.Drawing.Size(85, 28),
                DialogResult = DialogResult.OK
            };
            this.Controls.Add(btnOK);

            btnCancel = new Button
            {
                Text = "Hủy",
                Location = new System.Drawing.Point(335, 330),
                Size = new System.Drawing.Size(85, 28),
                DialogResult = DialogResult.Cancel
            };
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void LoadAlignments()
        {
            foreach (var alignment in alignmentInfos)
            {
                checkedListBox.Items.Add(alignment.Name, alignment.IsSelected);
            }
        }

        private void CheckedListBox_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            alignmentInfos[e.Index].IsSelected = (e.NewValue == CheckState.Checked);
        }

        private void SelectAll(bool select)
        {
            for (int i = 0; i < checkedListBox.Items.Count; i++)
            {
                checkedListBox.SetItemChecked(i, select);
                alignmentInfos[i].IsSelected = select;
            }
        }
    }

    /// <summary>
    /// Form sắp xếp thứ tự vật liệu
    /// </summary>
    public class FormSapXepVatLieu : Form
    {
        private ListView listViewMaterials;
        private Button btnMoveUp;
        private Button btnMoveDown;
        private Button btnOK;
        private Button btnCancel;
        private List<string> materials;

        public List<string> OrderedMaterials => materials;

        public FormSapXepVatLieu(List<string> materialNames)
        {
            materials = new List<string>(materialNames);
            InitializeComponent();
            LoadMaterials();
        }

        private void InitializeComponent()
        {
            this.Text = "Sắp xếp thứ tự vật liệu";
            this.Size = new System.Drawing.Size(500, 450);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            FormLabel lblTitle = new()
            {
                Text = "Sắp xếp thứ tự các vật liệu (cột trong Excel):",
                Location = new System.Drawing.Point(12, 12),
                Size = new System.Drawing.Size(450, 20)
            };
            this.Controls.Add(lblTitle);

            listViewMaterials = new ListView
            {
                Location = new System.Drawing.Point(12, 35),
                Size = new System.Drawing.Size(380, 320),
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                HideSelection = false
            };
            listViewMaterials.Columns.Add("STT", 50);
            listViewMaterials.Columns.Add("Tên vật liệu", 300);
            this.Controls.Add(listViewMaterials);

            btnMoveUp = new Button
            {
                Text = "▲ Lên",
                Location = new System.Drawing.Point(405, 100),
                Size = new System.Drawing.Size(70, 35)
            };
            btnMoveUp.Click += BtnMoveUp_Click;
            this.Controls.Add(btnMoveUp);

            btnMoveDown = new Button
            {
                Text = "▼ Xuống",
                Location = new System.Drawing.Point(405, 145),
                Size = new System.Drawing.Size(70, 35)
            };
            btnMoveDown.Click += BtnMoveDown_Click;
            this.Controls.Add(btnMoveDown);

            btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(290, 370),
                Size = new System.Drawing.Size(85, 28),
                DialogResult = DialogResult.OK
            };
            this.Controls.Add(btnOK);

            btnCancel = new Button
            {
                Text = "Hủy",
                Location = new System.Drawing.Point(385, 370),
                Size = new System.Drawing.Size(85, 28),
                DialogResult = DialogResult.Cancel
            };
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void LoadMaterials()
        {
            listViewMaterials.Items.Clear();
            for (int i = 0; i < materials.Count; i++)
            {
                var item = new ListViewItem((i + 1).ToString());
                item.SubItems.Add(materials[i]);
                listViewMaterials.Items.Add(item);
            }
            if (listViewMaterials.Items.Count > 0)
                listViewMaterials.Items[0].Selected = true;
        }

        private void BtnMoveUp_Click(object? sender, EventArgs e)
        {
            if (listViewMaterials.SelectedIndices.Count == 0) return;
            int index = listViewMaterials.SelectedIndices[0];
            if (index <= 0) return;

            // Swap in materials list
            (materials[index], materials[index - 1]) = (materials[index - 1], materials[index]);
            LoadMaterials();
            listViewMaterials.Items[index - 1].Selected = true;
        }

        private void BtnMoveDown_Click(object? sender, EventArgs e)
        {
            if (listViewMaterials.SelectedIndices.Count == 0) return;
            int index = listViewMaterials.SelectedIndices[0];
            if (index >= materials.Count - 1) return;

            // Swap in materials list
            (materials[index], materials[index + 1]) = (materials[index + 1], materials[index]);
            LoadMaterials();
            listViewMaterials.Items[index + 1].Selected = true;
        }
    }

    #endregion

    public class TinhKhoiLuongExcel
    {
        #region Công thức tính toán

        /// <summary>
        /// Tính diện tích đa giác sử dụng Shoelace Formula
        /// </summary>
        public static double CalculatePolygonArea(List<Point2d> points)
        {
            if (points.Count < 3) return 0;

            double area = 0;
            int n = points.Count;
            for (int i = 0; i < n; i++)
            {
                int j = (i + 1) % n;
                area += points[i].X * points[j].Y;
                area -= points[j].X * points[i].Y;
            }
            return Math.Abs(area) / 2.0;
        }

        /// <summary>
        /// Tính khối lượng bằng phương pháp trung bình cộng
        /// Volume[i] = (Area[i-1] + Area[i]) / 2 × Distance[i]
        /// </summary>
        public static double CalculateVolume(double areaPrev, double areaCurrent, double spacing)
        {
            return (areaPrev + areaCurrent) / 2.0 * spacing;
        }

        /// <summary>
        /// Format lý trình theo dạng Km0+123.456
        /// </summary>
        public static string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double meters = station % 1000;
            return $"Km{km}+{meters:F3}";
        }

        #endregion

        #region Commands

        [CommandMethod("CTSV_XuatKhoiLuong")]
        public static void CTSVXuatKhoiLuong()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                // 1. Lấy danh sách Alignments có SampleLineGroup
                List<AlignmentInfo> alignmentInfos = GetAlignmentsWithSampleLineGroups(tr);

                if (alignmentInfos.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông tìm thấy Alignment nào có SampleLineGroup!");
                    return;
                }

                // 2. Hiển thị form chọn Alignments
                FormChonAlignment formChon = new(alignmentInfos);
                if (formChon.ShowDialog() != DialogResult.OK || formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }

                // 3. Thu thập tất cả vật liệu từ các Alignment đã chọn
                HashSet<string> allMaterials = new();
                Dictionary<ObjectId, List<StakeInfo>> alignmentData = new();

                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    var stakeInfos = ExtractMaterialData(tr, alignInfo.SampleLineGroupId, alignInfo.AlignmentId);
                    alignmentData[alignInfo.AlignmentId] = stakeInfos;

                    foreach (var stake in stakeInfos)
                    {
                        foreach (var mat in stake.MaterialAreas.Keys)
                        {
                            allMaterials.Add(mat);
                        }
                    }
                }

                if (allMaterials.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông tìm thấy vật liệu nào trong các Alignment đã chọn!");
                    A.Ed.WriteMessage("\nĐảm bảo rằng bạn đã tạo Material Sections trong QTO Manager.");
                    return;
                }

                // 4. Hiển thị form sắp xếp thứ tự vật liệu
                FormSapXepVatLieu formSapXep = new(allMaterials.ToList());
                if (formSapXep.ShowDialog() != DialogResult.OK)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }

                List<string> orderedMaterials = formSapXep.OrderedMaterials;

                // 5. Tính khối lượng cho từng Alignment
                foreach (var kvp in alignmentData)
                {
                    CalculateVolumes(kvp.Value, orderedMaterials);
                }

                // 6. Chọn nơi lưu file Excel
                SaveFileDialog saveDialog = new()
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    Title = "Chọn nơi lưu file Excel khối lượng",
                    FileName = "KhoiLuongVatLieu.xlsx"
                };

                if (saveDialog.ShowDialog() != DialogResult.OK)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }

                // 7. Xuất ra Excel
                ExportToExcel(saveDialog.FileName, formChon.SelectedAlignments, alignmentData, orderedMaterials, tr);

                A.Ed.WriteMessage($"\nĐã xuất file Excel thành công: {saveDialog.FileName}");
                
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nLỗi: {ex.Message}");
                A.Ed.WriteMessage($"\nStack: {ex.StackTrace}");
            }
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// Lấy danh sách Alignments có SampleLineGroup
        /// </summary>
        private static List<AlignmentInfo> GetAlignmentsWithSampleLineGroups(Transaction tr)
        {
            List<AlignmentInfo> result = new();

            ObjectIdCollection alignmentIds = A.Cdoc.GetAlignmentIds();
            foreach (ObjectId alignmentId in alignmentIds)
            {
                Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
                if (alignment == null) continue;

                ObjectIdCollection slgIds = alignment.GetSampleLineGroupIds();
                if (slgIds.Count > 0)
                {
                    result.Add(new AlignmentInfo
                    {
                        AlignmentId = alignmentId,
                        Name = alignment.Name,
                        SampleLineGroupId = slgIds[0], // Lấy SampleLineGroup đầu tiên
                        IsSelected = true
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// Trích xuất dữ liệu vật liệu từ SampleLineGroup
        /// </summary>
        private static List<StakeInfo> ExtractMaterialData(Transaction tr, ObjectId sampleLineGroupId, ObjectId alignmentId)
        {
            List<StakeInfo> stakeInfos = new();

            SampleLineGroup? slg = tr.GetObject(sampleLineGroupId, AcadDb.OpenMode.ForRead) as SampleLineGroup;
            if (slg == null) return stakeInfos;

            Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
            if (alignment == null) return stakeInfos;

            ObjectIdCollection sampleLineIds = slg.GetSampleLineIds();
            double prevStation = 0;
            bool isFirst = true;

            foreach (ObjectId slId in sampleLineIds)
            {
                SampleLine? sampleLine = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                if (sampleLine == null) continue;

                StakeInfo stakeInfo = new()
                {
                    StakeName = sampleLine.Name,
                    StationValue = sampleLine.Station,
                    Station = FormatStation(sampleLine.Station),
                    SpacingPrev = isFirst ? 0 : (sampleLine.Station - prevStation)
                };

                // Lấy các Section từ SampleLine
                ObjectIdCollection sectionIds = sampleLine.GetSectionIds();
                foreach (ObjectId sectionId in sectionIds)
                {
                    try
                    {
                        // Kiểm tra xem đây có phải là MaterialSection không
                        AcadDb.DBObject? dbObj = tr.GetObject(sectionId, AcadDb.OpenMode.ForRead);
                        
                        if (dbObj is CivSection section)
                        {
                            // Lấy diện tích từ Section thông thường thông qua SectionPoints
                            string sourceName = GetSectionSourceName(tr, section, slg);
                            if (!string.IsNullOrEmpty(sourceName))
                            {
                                double area = CalculateSectionArea(section);
                                if (area > 0)
                                {
                                    if (!stakeInfo.MaterialAreas.ContainsKey(sourceName))
                                        stakeInfo.MaterialAreas[sourceName] = 0;
                                    stakeInfo.MaterialAreas[sourceName] += area;
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Bỏ qua section không đọc được
                    }
                }

                stakeInfos.Add(stakeInfo);
                prevStation = sampleLine.Station;
                isFirst = false;
            }

            return stakeInfos;
        }

        /// <summary>
        /// Lấy tên nguồn Section từ SectionSource
        /// </summary>
        private static string GetSectionSourceName(Transaction tr, CivSection section, SampleLineGroup slg)
        {
            try
            {
                // Tìm SectionSource tương ứng
                SectionSourceCollection sources = slg.GetSectionSources();
                foreach (SectionSource source in sources)
                {
                    if (source.SourceId == section.SourceId)
                    {
                        AcadDb.DBObject? sourceObj = tr.GetObject(source.SourceId, AcadDb.OpenMode.ForRead);
                        if (sourceObj is Corridor corridor)
                        {
                            return "Corridor - " + corridor.Name;
                        }
                        else if (sourceObj is TinSurface surface)
                        {
                            return surface.Name;
                        }
                        else if (sourceObj is Pipe pipe)
                        {
                            return pipe.Name;
                        }
                        else
                        {
                            return source.SourceName;
                        }
                    }
                }
            }
            catch
            {
                // Ignore
            }
            return "";
        }

        /// <summary>
        /// Tính diện tích Section từ SectionPoints
        /// </summary>
        private static double CalculateSectionArea(CivSection section)
        {
            try
            {
                SectionPointCollection points = section.SectionPoints;
                if (points.Count < 3) return 0;

                List<Point2d> pointList = new();
                foreach (SectionPoint pt in points)
                {
                    pointList.Add(new Point2d(pt.Location.X, pt.Location.Y));
                }

                // Đóng đa giác nếu chưa đóng
                if (pointList.Count >= 2)
                {
                    // Thêm điểm đóng về đường cơ sở
                    Point2d first = pointList[0];
                    Point2d last = pointList[pointList.Count - 1];
                    
                    // Nếu section là hở (đường), cần đóng về baseline
                    if (Math.Abs(first.X - last.X) > 0.001 || Math.Abs(first.Y - last.Y) > 0.001)
                    {
                        // Đóng đa giác dọc theo đường Y=0 hoặc Y min
                        double yBase = Math.Min(first.Y, last.Y);
                        pointList.Add(new Point2d(last.X, yBase));
                        pointList.Add(new Point2d(first.X, yBase));
                    }
                }

                return CalculatePolygonArea(pointList);
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Tính khối lượng cho tất cả các cọc
        /// </summary>
        private static void CalculateVolumes(List<StakeInfo> stakeInfos, List<string> materials)
        {
            for (int i = 0; i < stakeInfos.Count; i++)
            {
                var stake = stakeInfos[i];
                foreach (var material in materials)
                {
                    double areaCurrent = stake.MaterialAreas.GetValueOrDefault(material, 0);
                    double areaPrev = (i > 0) ? stakeInfos[i - 1].MaterialAreas.GetValueOrDefault(material, 0) : 0;
                    double spacing = stake.SpacingPrev;

                    stake.MaterialVolumes[material] = CalculateVolume(areaPrev, areaCurrent, spacing);
                }
            }
        }

        /// <summary>
        /// Xuất dữ liệu ra file Excel
        /// </summary>
        private static void ExportToExcel(string filePath, List<AlignmentInfo> alignments,
            Dictionary<ObjectId, List<StakeInfo>> alignmentData, List<string> materials, Transaction tr)
        {
            using var workbook = new XLWorkbook();

            // Dictionary để lưu tổng khối lượng cho sheet tổng hợp
            Dictionary<string, Dictionary<string, double>> totalVolumes = new();

            // Tạo sheet cho từng Alignment
            foreach (var alignInfo in alignments)
            {
                var stakeInfos = alignmentData[alignInfo.AlignmentId];
                string sheetName = SanitizeSheetName(alignInfo.Name);
                var ws = workbook.Worksheets.Add(sheetName);

                // Tạo header
                CreateSheetHeader(ws, materials, alignInfo.Name);

                // Điền dữ liệu
                int row = 3; // Bắt đầu từ hàng 3 (sau header)
                Dictionary<string, double> alignmentTotalVolumes = new();
                
                foreach (var material in materials)
                {
                    alignmentTotalVolumes[material] = 0;
                }

                foreach (var stake in stakeInfos)
                {
                    ws.Cell(row, 1).Value = stake.StakeName;
                    ws.Cell(row, 2).Value = stake.Station;
                    ws.Cell(row, 3).Value = Math.Round(stake.SpacingPrev, 3);

                    int col = 4;
                    foreach (var material in materials)
                    {
                        double area = stake.MaterialAreas.GetValueOrDefault(material, 0);
                        double volume = stake.MaterialVolumes.GetValueOrDefault(material, 0);

                        ws.Cell(row, col).Value = Math.Round(area, 3);
                        ws.Cell(row, col + 1).Value = Math.Round(volume, 3);

                        alignmentTotalVolumes[material] += volume;
                        col += 2;
                    }

                    row++;
                }

                // Thêm hàng tổng cộng
                int totalRow = row;
                ws.Cell(totalRow, 1).Value = "TỔNG CỘNG";
                ws.Range(totalRow, 1, totalRow, 3).Merge();
                ws.Cell(totalRow, 1).Style.Font.Bold = true;
                ws.Cell(totalRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int colTotal = 4;
                foreach (var material in materials)
                {
                    // Sum diện tích
                    ws.Cell(totalRow, colTotal).FormulaA1 = $"SUM({GetColumnLetter(colTotal)}3:{GetColumnLetter(colTotal)}{row - 1})";
                    // Sum khối lượng
                    ws.Cell(totalRow, colTotal + 1).FormulaA1 = $"SUM({GetColumnLetter(colTotal + 1)}3:{GetColumnLetter(colTotal + 1)}{row - 1})";
                    colTotal += 2;
                }

                // Format bảng
                FormatWorksheet(ws, row, 3 + materials.Count * 2);

                // Lưu tổng khối lượng cho sheet tổng hợp
                totalVolumes[sheetName] = alignmentTotalVolumes;
            }

            // Tạo sheet TỔNG HỢP nếu có nhiều hơn 1 alignment
            if (alignments.Count > 1)
            {
                CreateSummarySheet(workbook, alignments, totalVolumes, materials, tr);
            }

            // Lưu file
            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Tạo header cho sheet
        /// </summary>
        private static void CreateSheetHeader(IXLWorksheet ws, List<string> materials, string alignmentName)
        {
            // Tiêu đề sheet
            ws.Cell(1, 1).Value = $"BẢNG TÍNH KHỐI LƯỢNG - {alignmentName}";
            int lastCol = 3 + materials.Count * 2;
            ws.Range(1, 1, 1, lastCol).Merge();
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontSize = 14;
            ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Header row
            ws.Cell(2, 1).Value = "TÊN CỌC";
            ws.Cell(2, 2).Value = "LÝ TRÌNH";
            ws.Cell(2, 3).Value = "KHOẢNG CÁCH (m)";

            int col = 4;
            foreach (var material in materials)
            {
                // Header cho mỗi vật liệu: 2 cột (Diện tích, Khối lượng)
                ws.Cell(2, col).Value = $"{material} - DIỆN TÍCH (m²)";
                ws.Cell(2, col + 1).Value = $"{material} - KHỐI LƯỢNG (m³)";
                col += 2;
            }

            // Format header
            var headerRange = ws.Range(2, 1, 2, lastCol);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.WrapText = true;
        }

        /// <summary>
        /// Format toàn bộ worksheet
        /// </summary>
        private static void FormatWorksheet(IXLWorksheet ws, int lastRow, int lastCol)
        {
            // Border cho toàn bộ bảng
            var tableRange = ws.Range(2, 1, lastRow, lastCol);
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Điều chỉnh độ rộng cột
            ws.Column(1).Width = 15;  // Tên cọc
            ws.Column(2).Width = 15;  // Lý trình
            ws.Column(3).Width = 15;  // Khoảng cách

            for (int c = 4; c <= lastCol; c++)
            {
                ws.Column(c).Width = 18;
            }

            // Căn giữa các cột số
            for (int r = 3; r <= lastRow; r++)
            {
                for (int c = 3; c <= lastCol; c++)
                {
                    ws.Cell(r, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                }
            }
        }

        /// <summary>
        /// Tạo sheet tổng hợp
        /// </summary>
        private static void CreateSummarySheet(XLWorkbook workbook, List<AlignmentInfo> alignments,
            Dictionary<string, Dictionary<string, double>> totalVolumes, List<string> materials, Transaction tr)
        {
            var ws = workbook.Worksheets.Add("TỔNG HỢP");

            // Tiêu đề
            ws.Cell(1, 1).Value = "BẢNG TỔNG HỢP KHỐI LƯỢNG TẤT CẢ CÁC TUYẾN";
            int lastCol = 1 + materials.Count;
            ws.Range(1, 1, 1, lastCol).Merge();
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontSize = 14;
            ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Header
            ws.Cell(2, 1).Value = "TUYẾN";
            int col = 2;
            foreach (var material in materials)
            {
                ws.Cell(2, col).Value = $"{material} (m³)";
                col++;
            }

            // Header format
            var headerRange = ws.Range(2, 1, 2, lastCol);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGreen;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Dữ liệu
            int row = 3;
            foreach (var alignInfo in alignments)
            {
                string sheetName = SanitizeSheetName(alignInfo.Name);
                ws.Cell(row, 1).Value = alignInfo.Name;

                col = 2;
                foreach (var material in materials)
                {
                    double volume = totalVolumes[sheetName].GetValueOrDefault(material, 0);
                    // Tham chiếu đến sheet chi tiết
                    ws.Cell(row, col).FormulaA1 = $"'{sheetName}'!{GetColumnLetter(3 + (materials.IndexOf(material) * 2) + 1)}{GetLastRowForSheet(totalVolumes, sheetName, materials)}";
                    col++;
                }
                row++;
            }

            // Hàng tổng cộng
            ws.Cell(row, 1).Value = "TỔNG CỘNG";
            ws.Cell(row, 1).Style.Font.Bold = true;
            col = 2;
            foreach (var material in materials)
            {
                ws.Cell(row, col).FormulaA1 = $"SUM({GetColumnLetter(col)}3:{GetColumnLetter(col)}{row - 1})";
                col++;
            }

            // Format
            var tableRange = ws.Range(2, 1, row, lastCol);
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Column(1).Width = 30;
            for (int c = 2; c <= lastCol; c++)
            {
                ws.Column(c).Width = 20;
            }
        }

        /// <summary>
        /// Lấy số hàng cuối cùng cho sheet để tham chiếu
        /// </summary>
        private static int GetLastRowForSheet(Dictionary<string, Dictionary<string, double>> totalVolumes, string sheetName, List<string> materials)
        {
            // Cần đếm số hàng trong sheet đó - giả sử là row của TỔNG CỘNG
            // Điều này cần được tính toán từ số lượng stake thực tế
            return 50; // Placeholder - sẽ được cập nhật khi tạo sheet
        }

        /// <summary>
        /// Lấy ký tự cột Excel từ số cột
        /// </summary>
        private static string GetColumnLetter(int columnNumber)
        {
            string columnLetter = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnLetter = Convert.ToChar('A' + modulo) + columnLetter;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnLetter;
        }

        /// <summary>
        /// Làm sạch tên sheet (Excel không cho phép một số ký tự đặc biệt)
        /// </summary>
        private static string SanitizeSheetName(string name)
        {
            char[] invalidChars = { ':', '\\', '/', '?', '*', '[', ']' };
            string result = name;
            foreach (char c in invalidChars)
            {
                result = result.Replace(c, '_');
            }
            // Giới hạn độ dài tên sheet là 31 ký tự
            if (result.Length > 31)
            {
                result = result.Substring(0, 31);
            }
            return result;
        }

        #endregion
    }
}
