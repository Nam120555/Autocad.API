// (C) Copyright 2024
// Thống kê cọc từ SampleLine và xuất ra Excel
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
using FormLabel = System.Windows.Forms.Label;

using ClosedXML.Excel;
using MyFirstProject.Extensions;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(Civil3DCsharp.ThongKeCoc))]

namespace Civil3DCsharp
{
    #region Data Classes

    /// <summary>
    /// Thông tin chi tiết của một cọc (SampleLine)
    /// </summary>
    public class CocInfo
    {
        public int STT { get; set; }                      // Số thứ tự
        public string TenCoc { get; set; } = "";          // Tên cọc
        public string LyTrinh { get; set; } = "";         // Lý trình format: Km0+123.456
        public double LyTrinhValue { get; set; }          // Giá trị lý trình số
        public double Easting { get; set; }               // Tọa độ X
        public double Northing { get; set; }              // Tọa độ Y
        public double KhoangCachDenCocTruoc { get; set; } // Khoảng cách đến cọc trước
        public double BeRongTrai { get; set; }            // Bề rộng bên trái
        public double BeRongPhai { get; set; }            // Bề rộng bên phải
    }

    #endregion

    public class ThongKeCoc
    {
        #region Helper Methods

        /// <summary>
        /// Format lý trình theo dạng Km0+123.456
        /// </summary>
        public static string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double meters = station % 1000;
            return $"Km{km}+{meters:F3}";
        }

        /// <summary>
        /// Lấy danh sách Alignments có SampleLineGroup
        /// </summary>
        private static List<(ObjectId AlignmentId, string Name, ObjectId SampleLineGroupId)> GetAlignmentsWithSampleLineGroups(Transaction tr)
        {
            var result = new List<(ObjectId, string, ObjectId)>();

            ObjectIdCollection alignmentIds = A.Cdoc.GetAlignmentIds();
            foreach (ObjectId alignmentId in alignmentIds)
            {
                Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
                if (alignment == null) continue;

                ObjectIdCollection slgIds = alignment.GetSampleLineGroupIds();
                if (slgIds.Count > 0)
                {
                    result.Add((alignmentId, alignment.Name, slgIds[0]));
                }
            }

            return result;
        }

        /// <summary>
        /// Trích xuất thông tin cọc từ SampleLineGroup
        /// </summary>
        private static List<CocInfo> ExtractCocInfo(Transaction tr, ObjectId sampleLineGroupId, ObjectId alignmentId)
        {
            List<CocInfo> cocInfos = new();

            SampleLineGroup? slg = tr.GetObject(sampleLineGroupId, AcadDb.OpenMode.ForRead) as SampleLineGroup;
            if (slg == null) return cocInfos;

            Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
            if (alignment == null) return cocInfos;

            // Lấy danh sách SampleLine và sắp xếp theo lý trình
            List<SampleLine> sortedSampleLines = new List<SampleLine>();
            foreach (ObjectId slId in slg.GetSampleLineIds())
            {
                if (tr.GetObject(slId, OpenMode.ForRead) is SampleLine sl)
                    sortedSampleLines.Add(sl);
            }
            sortedSampleLines = sortedSampleLines.OrderBy(s => s.Station).ToList();

            double prevStation = 0;
            int stt = 1;

            foreach (SampleLine sampleLine in sortedSampleLines)
            {
                // Lấy tọa độ tâm cọc
                double easting = 0, northing = 0;
                alignment.PointLocation(sampleLine.Station, 0, ref easting, ref northing);

                // Lấy bề rộng trái/phải từ Vertices
                double beRongTrai = 0, beRongPhai = 0;
                foreach (SampleLineVertex vertex in sampleLine.Vertices)
                {
                    double dist = Math.Sqrt(Math.Pow(vertex.Location.X - easting, 2) + Math.Pow(vertex.Location.Y - northing, 2));
                    if (vertex.Side == SampleLineVertexSideType.Left)
                        beRongTrai = dist;
                    else if (vertex.Side == SampleLineVertexSideType.Right)
                        beRongPhai = dist;
                }

                CocInfo cocInfo = new()
                {
                    STT = stt,
                    TenCoc = sampleLine.Name,
                    LyTrinh = FormatStation(sampleLine.Station),
                    LyTrinhValue = sampleLine.Station,
                    Easting = Math.Round(easting, 3),
                    Northing = Math.Round(northing, 3),
                    KhoangCachDenCocTruoc = stt == 1 ? 0 : Math.Round(sampleLine.Station - prevStation, 3),
                    BeRongTrai = Math.Round(beRongTrai, 3),
                    BeRongPhai = Math.Round(beRongPhai, 3)
                };

                cocInfos.Add(cocInfo);
                prevStation = sampleLine.Station;
                stt++;
            }

            return cocInfos;
        }

        /// <summary>
        /// Xuất danh sách cọc ra Excel
        /// </summary>
        private static void ExportToExcel(string filePath, string alignmentName, List<CocInfo> cocInfos)
        {
            using var workbook = new XLWorkbook();
            string sheetName = alignmentName.Length > 31 ? alignmentName.Substring(0, 31) : alignmentName;
            // Thay thế ký tự không hợp lệ
            char[] invalidChars = { ':', '\\', '/', '?', '*', '[', ']' };
            foreach (char c in invalidChars)
            {
                sheetName = sheetName.Replace(c, '_');
            }

            var ws = workbook.Worksheets.Add(sheetName);

            // Tiêu đề
            ws.Cell(1, 1).Value = $"BẢNG THỐNG KÊ CỌC - {alignmentName}";
            ws.Range(1, 1, 1, 9).Merge();
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontSize = 14;
            ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Header
            ws.Cell(2, 1).Value = "STT";
            ws.Cell(2, 2).Value = "TÊN CỌC";
            ws.Cell(2, 3).Value = "LÝ TRÌNH";
            ws.Cell(2, 4).Value = "EASTING (X)";
            ws.Cell(2, 5).Value = "NORTHING (Y)";
            ws.Cell(2, 6).Value = "KHOẢNG CÁCH (m)";
            ws.Cell(2, 7).Value = "BỀ RỘNG TRÁI (m)";
            ws.Cell(2, 8).Value = "BỀ RỘNG PHẢI (m)";
            ws.Cell(2, 9).Value = "TỔNG BỀ RỘNG (m)";

            // Format header
            var headerRange = ws.Range(2, 1, 2, 9);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.WrapText = true;

            // Dữ liệu
            int row = 3;
            foreach (var coc in cocInfos)
            {
                ws.Cell(row, 1).Value = coc.STT;
                ws.Cell(row, 2).Value = coc.TenCoc;
                ws.Cell(row, 3).Value = coc.LyTrinh;
                ws.Cell(row, 4).Value = coc.Easting;
                ws.Cell(row, 5).Value = coc.Northing;
                ws.Cell(row, 6).Value = coc.KhoangCachDenCocTruoc;
                ws.Cell(row, 7).Value = coc.BeRongTrai;
                ws.Cell(row, 8).Value = coc.BeRongPhai;
                ws.Cell(row, 9).Value = coc.BeRongTrai + coc.BeRongPhai;
                row++;
            }

            // Hàng tổng cộng
            ws.Cell(row, 1).Value = "TỔNG CỘNG";
            ws.Range(row, 1, row, 5).Merge();
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 6).FormulaA1 = $"SUM(F3:F{row - 1})";
            ws.Cell(row, 6).Style.Font.Bold = true;

            // Format bảng
            var tableRange = ws.Range(2, 1, row, 9);
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Điều chỉnh độ rộng cột
            ws.Column(1).Width = 8;   // STT
            ws.Column(2).Width = 20;  // Tên cọc
            ws.Column(3).Width = 15;  // Lý trình
            ws.Column(4).Width = 15;  // Easting
            ws.Column(5).Width = 15;  // Northing
            ws.Column(6).Width = 15;  // Khoảng cách
            ws.Column(7).Width = 15;  // Bề rộng trái
            ws.Column(8).Width = 15;  // Bề rộng phải
            ws.Column(9).Width = 15;  // Tổng bề rộng

            // Căn giữa các cột số
            for (int r = 3; r <= row; r++)
            {
                ws.Cell(r, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                for (int c = 4; c <= 9; c++)
                {
                    ws.Cell(r, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                }
            }

            workbook.SaveAs(filePath);
        }

        #endregion

        #region Commands

        /// <summary>
        /// Lệnh thống kê cọc và xuất Excel
        /// </summary>
        [CommandMethod("CTSV_ThongKeCoc")]
        public static void CTSVThongKeCoc()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                // 1. Lấy danh sách Alignments có SampleLineGroup
                var alignments = GetAlignmentsWithSampleLineGroups(tr);

                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông tìm thấy Alignment nào có SampleLineGroup!");
                    return;
                }

                // 2. Chọn Alignment
                string[] alignmentNames = alignments.Select(a => a.Name).ToArray();
                
                // Hiển thị danh sách để chọn
                A.Ed.WriteMessage("\n=== DANH SÁCH ALIGNMENT ===");
                for (int i = 0; i < alignmentNames.Length; i++)
                {
                    A.Ed.WriteMessage($"\n  {i + 1}. {alignmentNames[i]}");
                }

                PromptIntegerOptions pOpts = new("\nNhập số thứ tự Alignment cần thống kê cọc: ");
                pOpts.LowerLimit = 1;
                pOpts.UpperLimit = alignmentNames.Length;
                pOpts.DefaultValue = 1;

                PromptIntegerResult pRes = A.Ed.GetInteger(pOpts);
                if (pRes.Status != PromptStatus.OK) return;

                int selectedIndex = pRes.Value - 1;
                var selectedAlignment = alignments[selectedIndex];

                // 3. Trích xuất thông tin cọc
                List<CocInfo> cocInfos = ExtractCocInfo(tr, selectedAlignment.SampleLineGroupId, selectedAlignment.AlignmentId);

                if (cocInfos.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông có cọc nào trong SampleLineGroup!");
                    return;
                }

                // 4. Hiển thị thống kê trên command line
                A.Ed.WriteMessage($"\n\n=== THỐNG KÊ CỌC - {selectedAlignment.Name} ===");
                A.Ed.WriteMessage($"\nTổng số cọc: {cocInfos.Count}");
                A.Ed.WriteMessage($"\nCọc đầu: {cocInfos.First().TenCoc} - {cocInfos.First().LyTrinh}");
                A.Ed.WriteMessage($"\nCọc cuối: {cocInfos.Last().TenCoc} - {cocInfos.Last().LyTrinh}");
                A.Ed.WriteMessage($"\nTổng chiều dài: {cocInfos.Sum(c => c.KhoangCachDenCocTruoc):F3} m");

                // 5. Hỏi có muốn xuất Excel không
                PromptKeywordOptions keyOpts = new("\nBạn có muốn xuất ra file Excel? [Yes/No] <Yes>: ", "Yes No");
                keyOpts.Keywords.Default = "Yes";
                PromptResult keyRes = A.Ed.GetKeywords(keyOpts);

                if (keyRes.Status == PromptStatus.OK && keyRes.StringResult == "Yes")
                {
                    // 6. Chọn nơi lưu file Excel
                    SaveFileDialog saveDialog = new()
                    {
                        Filter = "Excel Files (*.xlsx)|*.xlsx",
                        Title = "Chọn nơi lưu file Excel thống kê cọc",
                        FileName = $"ThongKeCoc_{selectedAlignment.Name}.xlsx"
                    };

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        ExportToExcel(saveDialog.FileName, selectedAlignment.Name, cocInfos);
                        A.Ed.WriteMessage($"\nĐã xuất file Excel: {saveDialog.FileName}");
                    }
                }

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nLỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Lệnh xuất thống kê cọc tất cả Alignments ra Excel
        /// </summary>
        [CommandMethod("CTSV_ThongKeCoc_TatCa")]
        public static void CTSVThongKeCocTatCa()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                // 1. Lấy danh sách Alignments có SampleLineGroup
                var alignments = GetAlignmentsWithSampleLineGroups(tr);

                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông tìm thấy Alignment nào có SampleLineGroup!");
                    return;
                }

                // 2. Chọn nơi lưu file Excel
                SaveFileDialog saveDialog = new()
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    Title = "Chọn nơi lưu file Excel thống kê cọc tất cả tuyến",
                    FileName = "ThongKeCoc_TatCa.xlsx"
                };

                if (saveDialog.ShowDialog() != DialogResult.OK) return;

                // 3. Tạo workbook với nhiều sheet
                using var workbook = new XLWorkbook();

                foreach (var alignment in alignments)
                {
                    List<CocInfo> cocInfos = ExtractCocInfo(tr, alignment.SampleLineGroupId, alignment.AlignmentId);
                    if (cocInfos.Count == 0) continue;

                    string sheetName = alignment.Name.Length > 31 ? alignment.Name.Substring(0, 31) : alignment.Name;
                    char[] invalidChars = { ':', '\\', '/', '?', '*', '[', ']' };
                    foreach (char c in invalidChars)
                    {
                        sheetName = sheetName.Replace(c, '_');
                    }

                    var ws = workbook.Worksheets.Add(sheetName);

                    // Tiêu đề
                    ws.Cell(1, 1).Value = $"BẢNG THỐNG KÊ CỌC - {alignment.Name}";
                    ws.Range(1, 1, 1, 9).Merge();
                    ws.Cell(1, 1).Style.Font.Bold = true;
                    ws.Cell(1, 1).Style.Font.FontSize = 14;
                    ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Header
                    ws.Cell(2, 1).Value = "STT";
                    ws.Cell(2, 2).Value = "TÊN CỌC";
                    ws.Cell(2, 3).Value = "LÝ TRÌNH";
                    ws.Cell(2, 4).Value = "EASTING (X)";
                    ws.Cell(2, 5).Value = "NORTHING (Y)";
                    ws.Cell(2, 6).Value = "KHOẢNG CÁCH (m)";
                    ws.Cell(2, 7).Value = "BỀ RỘNG TRÁI (m)";
                    ws.Cell(2, 8).Value = "BỀ RỘNG PHẢI (m)";
                    ws.Cell(2, 9).Value = "TỔNG BỀ RỘNG (m)";

                    var headerRange = ws.Range(2, 1, 2, 9);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Dữ liệu
                    int row = 3;
                    foreach (var coc in cocInfos)
                    {
                        ws.Cell(row, 1).Value = coc.STT;
                        ws.Cell(row, 2).Value = coc.TenCoc;
                        ws.Cell(row, 3).Value = coc.LyTrinh;
                        ws.Cell(row, 4).Value = coc.Easting;
                        ws.Cell(row, 5).Value = coc.Northing;
                        ws.Cell(row, 6).Value = coc.KhoangCachDenCocTruoc;
                        ws.Cell(row, 7).Value = coc.BeRongTrai;
                        ws.Cell(row, 8).Value = coc.BeRongPhai;
                        ws.Cell(row, 9).Value = coc.BeRongTrai + coc.BeRongPhai;
                        row++;
                    }

                    // Format
                    var tableRange = ws.Range(2, 1, row - 1, 9);
                    tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    ws.Columns().AdjustToContents();

                    A.Ed.WriteMessage($"\n  Đã thêm sheet: {sheetName} ({cocInfos.Count} cọc)");
                }

                workbook.SaveAs(saveDialog.FileName);
                A.Ed.WriteMessage($"\n\nĐã xuất file Excel: {saveDialog.FileName}");

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nLỗi: {ex.Message}");
            }
        }

        #endregion
    }
}
