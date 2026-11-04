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


// Tắt cảnh báo nullable cho file lớn
#pragma warning disable CS8600, CS8602, CS8618

// Thống kê vật liệu trắc ngang: trích xuất, xuất Excel, tạo bảng CAD
// (tách từ 10.TinhKhoiLuongExcel.cs — partial class TinhKhoiLuongExcel)
namespace MyFirstProject
{
    public partial class TinhKhoiLuongExcel
    {


        /// <summary>
        /// Thông tin thống kê material tại một trắc ngang
        /// </summary>
        public class MaterialStatInfo
        {
            public string SampleLineName { get; set; } = "";
            public double Station { get; set; }
            public string StationFormatted { get; set; } = "";
            public double SpacingPrev { get; set; }
            
            // Diện tích theo vật liệu
            public Dictionary<string, double> MaterialAreas { get; set; } = new();
            
            // Chi tiết từng material (Left/Right offset, Min/Max elevation)
            public Dictionary<string, MaterialDetailInfo> MaterialDetails { get; set; } = new();
        }

        /// <summary>
        /// Chi tiết về một material tại trắc ngang
        /// </summary>
        public class MaterialDetailInfo
        {
            public double Area { get; set; }
            public double LeftOffset { get; set; }
            public double RightOffset { get; set; }
            public double MinElevation { get; set; }
            public double MaxElevation { get; set; }
            public int PointCount { get; set; }
        }

        /// <summary>
        /// Trích xuất thống kê Material từ SampleLineGroup
        /// </summary>
        private static List<MaterialStatInfo> ExtractMaterialStatistics(
            Transaction tr, ObjectId sampleLineGroupId, ObjectId alignmentId)
        {
            List<MaterialStatInfo> stats = new();

            SampleLineGroup? slg = tr.GetObject(sampleLineGroupId, AcadDb.OpenMode.ForRead, false, true) as SampleLineGroup;
            if (slg == null) return stats;

            Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
            if (alignment == null) return stats;

            // Lấy Material Lists
            QTOMaterialListCollection materialLists = slg.MaterialLists;
            if (materialLists.Count == 0)
            {
                A.Ed.WriteMessage($"\n⚠️ Không có Material List trong '{slg.Name}'");
                return stats;
            }

            // Thu thập Materials với GUID
            List<(string Name, Guid MaterialListGuid, Guid MaterialGuid)> materials = new();
            foreach (QTOMaterialList materialList in materialLists)
            {
                try
                {
                    Guid listGuid = materialList.Guid;
                    foreach (QTOMaterial material in materialList)
                    {
                        materials.Add((material.Name, listGuid, material.Guid));
                    }
                }
                catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
            }

            // Sắp xếp SampleLines theo lý trình
            List<SampleLine> sortedSampleLines = new();
            foreach (ObjectId slId in slg.GetSampleLineIds())
            {
                var sl = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                if (sl != null) sortedSampleLines.Add(sl);
            }
            sortedSampleLines = sortedSampleLines.OrderBy(s => s.Station).ToList();

            // Duyệt qua từng SampleLine
            double prevStation = 0;
            bool isFirst = true;

            foreach (SampleLine sampleLine in sortedSampleLines)
            {
                MaterialStatInfo statInfo = new()
                {
                    SampleLineName = sampleLine.Name,
                    Station = sampleLine.Station,
                    StationFormatted = FormatStation(sampleLine.Station),
                    SpacingPrev = isFirst ? 0 : (sampleLine.Station - prevStation)
                };

                // Lấy MaterialSection cho từng Material
                foreach (var (materialName, materialListGuid, materialGuid) in materials)
                {
                    try
                    {
                        ObjectId materialSectionId = sampleLine.GetMaterialSectionId(materialListGuid, materialGuid);

                        if (!materialSectionId.IsNull && materialSectionId.IsValid)
                        {
                            AcadDb.DBObject? sectionObj = tr.GetObject(materialSectionId, AcadDb.OpenMode.ForRead, false, true);

                            if (sectionObj is CivSection section)
                            {
                                // Lấy diện tích
                                double areaFromAPI = section.Area;
                                double areaCalculated = CalculateSectionArea(section);
                                double area = areaFromAPI > 0 ? areaFromAPI : areaCalculated;

                                if (area > 0)
                                {
                                    statInfo.MaterialAreas[materialName] = area;

                                    // Lấy chi tiết material
                                    MaterialDetailInfo detail = new() { Area = area };

                                    try
                                    {
                                        SectionPointCollection points = section.SectionPoints;
                                        detail.PointCount = points.Count;

                                        if (points.Count > 0)
                                        {
                                            double minX = double.MaxValue, maxX = double.MinValue;
                                            double minY = double.MaxValue, maxY = double.MinValue;

                                            foreach (SectionPoint pt in points)
                                            {
                                                if (pt.Location.X < minX) minX = pt.Location.X;
                                                if (pt.Location.X > maxX) maxX = pt.Location.X;
                                                if (pt.Location.Y < minY) minY = pt.Location.Y;
                                                if (pt.Location.Y > maxY) maxY = pt.Location.Y;
                                            }

                                            detail.LeftOffset = minX;
                                            detail.RightOffset = maxX;
                                            detail.MinElevation = minY;
                                            detail.MaxElevation = maxY;
                                        }
                                    }
                                    catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

                                    statInfo.MaterialDetails[materialName] = detail;
                                }
                            }
                        }
                    }
                    catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                }

                stats.Add(statInfo);
                prevStation = sampleLine.Station;
                isFirst = false;
            }

            return stats;
        }

        /// <summary>
        /// Xuất thống kê Material ra Excel
        /// </summary>
        private static void ExportMaterialStatisticsToExcel(
            string filePath,
            List<AlignmentInfo> alignments,
            Dictionary<ObjectId, List<MaterialStatInfo>> alignmentMaterialData,
            List<string> materials)
        {
            using var workbook = new XLWorkbook();

            // Sheet cho từng Alignment
            foreach (var alignInfo in alignments)
            {
                var stats = alignmentMaterialData[alignInfo.AlignmentId];
                string sheetName = SanitizeSheetName($"TN_{alignInfo.Name}");
                var ws = workbook.Worksheets.Add(sheetName);

                // Tiêu đề
                ws.Cell(1, 1).Value = $"BẢNG THỐNG KÊ MATERIAL THEO TRẮC NGANG - {alignInfo.Name}";
                ws.Range(1, 1, 1, 4 + materials.Count * 4).Merge();
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Font.FontSize = 14;
                ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Header hàng 2 - Nhóm
                ws.Cell(2, 1).Value = "THÔNG TIN TRẮC NGANG";
                ws.Range(2, 1, 2, 4).Merge();
                ws.Cell(2, 1).Style.Font.Bold = true;
                ws.Cell(2, 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                ws.Cell(2, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int col = 5;
                foreach (var mat in materials)
                {
                    ws.Cell(2, col).Value = mat;
                    ws.Range(2, col, 2, col + 3).Merge();
                    ws.Cell(2, col).Style.Font.Bold = true;
                    ws.Cell(2, col).Style.Fill.BackgroundColor = XLColor.LightGreen;
                    ws.Cell(2, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    col += 4;
                }

                // Header hàng 3 - Chi tiết
                ws.Cell(3, 1).Value = "STT";
                ws.Cell(3, 2).Value = "TÊN TRẮC NGANG";
                ws.Cell(3, 3).Value = "LÝ TRÌNH";
                ws.Cell(3, 4).Value = "K.CÁCH (m)";

                col = 5;
                foreach (var mat in materials)
                {
                    ws.Cell(3, col).Value = "Diện tích (m²)";
                    ws.Cell(3, col + 1).Value = "Offset trái";
                    ws.Cell(3, col + 2).Value = "Offset phải";
                    ws.Cell(3, col + 3).Value = "Cao độ";
                    col += 4;
                }

                // Format header hàng 3
                int lastCol = 4 + materials.Count * 4;
                ws.Range(3, 1, 3, lastCol).Style.Font.Bold = true;
                ws.Range(3, 1, 3, lastCol).Style.Fill.BackgroundColor = XLColor.LightBlue;
                ws.Range(3, 1, 3, lastCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Dữ liệu
                int row = 4;
                int stt = 1;
                Dictionary<string, double> totalAreas = materials.ToDictionary(m => m, m => 0.0);

                foreach (var stat in stats)
                {
                    ws.Cell(row, 1).Value = stt++;
                    ws.Cell(row, 2).Value = stat.SampleLineName;
                    ws.Cell(row, 3).Value = stat.StationFormatted;
                    ws.Cell(row, 4).Value = Math.Round(stat.SpacingPrev, 3);

                    col = 5;
                    foreach (var mat in materials)
                    {
                        double area = stat.MaterialAreas.GetValueOrDefault(mat, 0);
                        ws.Cell(row, col).Value = Math.Round(area, 4);
                        totalAreas[mat] += area;

                        if (stat.MaterialDetails.ContainsKey(mat))
                        {
                            var detail = stat.MaterialDetails[mat];
                            ws.Cell(row, col + 1).Value = Math.Round(detail.LeftOffset, 3);
                            ws.Cell(row, col + 2).Value = Math.Round(detail.RightOffset, 3);
                            ws.Cell(row, col + 3).Value = $"{detail.MinElevation:F2} ~ {detail.MaxElevation:F2}";
                        }

                        col += 4;
                    }

                    row++;
                }

                // Hàng tổng cộng
                ws.Cell(row, 1).Value = "TỔNG CỘNG";
                ws.Range(row, 1, row, 4).Merge();
                ws.Cell(row, 1).Style.Font.Bold = true;
                ws.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                col = 5;
                foreach (var mat in materials)
                {
                    ws.Cell(row, col).Value = Math.Round(totalAreas[mat], 4);
                    ws.Cell(row, col).Style.Font.Bold = true;
                    col += 4;
                }

                // Border và format
                ws.Range(2, 1, row, lastCol).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(2, 1, row, lastCol).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                ws.Column(1).Width = 8;
                ws.Column(2).Width = 20;
                ws.Column(3).Width = 15;
                ws.Column(4).Width = 12;

                for (int c = 5; c <= lastCol; c++)
                {
                    ws.Column(c).Width = 15;
                }
            }

            // Sheet tổng hợp
            if (alignments.Count > 1)
            {
                var wsSummary = workbook.Worksheets.Add("TỔNG HỢP");
                
                wsSummary.Cell(1, 1).Value = "TỔNG HỢP DIỆN TÍCH MATERIAL TẤT CẢ CÁC TUYẾN";
                wsSummary.Range(1, 1, 1, 1 + materials.Count).Merge();
                wsSummary.Cell(1, 1).Style.Font.Bold = true;
                wsSummary.Cell(1, 1).Style.Font.FontSize = 14;
                wsSummary.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Header
                wsSummary.Cell(2, 1).Value = "TUYẾN";
                int col = 2;
                foreach (var mat in materials)
                {
                    wsSummary.Cell(2, col).Value = $"{mat} (m²)";
                    col++;
                }

                wsSummary.Range(2, 1, 2, 1 + materials.Count).Style.Font.Bold = true;
                wsSummary.Range(2, 1, 2, 1 + materials.Count).Style.Fill.BackgroundColor = XLColor.LightGreen;

                // Dữ liệu
                int row = 3;
                Dictionary<string, double> grandTotals = materials.ToDictionary(m => m, m => 0.0);

                foreach (var alignInfo in alignments)
                {
                    var stats = alignmentMaterialData[alignInfo.AlignmentId];
                    wsSummary.Cell(row, 1).Value = alignInfo.Name;

                    col = 2;
                    foreach (var mat in materials)
                    {
                        double total = stats.Sum(s => s.MaterialAreas.GetValueOrDefault(mat, 0));
                        wsSummary.Cell(row, col).Value = Math.Round(total, 4);
                        grandTotals[mat] += total;
                        col++;
                    }

                    row++;
                }

                // Tổng cộng
                wsSummary.Cell(row, 1).Value = "TỔNG CỘNG";
                wsSummary.Cell(row, 1).Style.Font.Bold = true;

                col = 2;
                foreach (var mat in materials)
                {
                    wsSummary.Cell(row, col).Value = Math.Round(grandTotals[mat], 4);
                    wsSummary.Cell(row, col).Style.Font.Bold = true;
                    col++;
                }

                // Format
                wsSummary.Range(2, 1, row, 1 + materials.Count).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                wsSummary.Range(2, 1, row, 1 + materials.Count).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                wsSummary.Column(1).Width = 25;
                for (int c = 2; c <= 1 + materials.Count; c++)
                {
                    wsSummary.Column(c).Width = 18;
                }
            }

            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Tạo bảng thống kê Material trong CAD
        /// </summary>
        private static void CreateMaterialStatTable(
            Transaction tr,
            Point3d insertPoint,
            string alignmentName,
            List<MaterialStatInfo> stats,
            List<string> materials)
        {
            AcadDb.Database db = HostApplicationServices.WorkingDatabase;
            AcadDb.BlockTable bt = tr.GetObject(db.BlockTableId, AcadDb.OpenMode.ForRead) as AcadDb.BlockTable
                ?? throw new System.Exception("Không thể mở BlockTable");
            AcadDb.BlockTableRecord btr = tr.GetObject(bt[AcadDb.BlockTableRecord.ModelSpace], AcadDb.OpenMode.ForWrite) as AcadDb.BlockTableRecord
                ?? throw new System.Exception("Không thể mở ModelSpace");

            // Tính số cột và hàng
            int numCols = 4 + materials.Count; // STT, Tên, Lý trình, K.cách + Materials
            int numRows = stats.Count + 4; // 2 header + dữ liệu + 1 tổng cộng

            // Tạo Table
            AcadDb.Table table = new()
            {
                Position = insertPoint,
                TableStyle = db.Tablestyle
            };

            table.SetSize(numRows, numCols);

            // Kích thước hàng/cột
            double rowHeight = 8.0;
            for (int r = 0; r < numRows; r++)
            {
                table.Rows[r].Height = rowHeight;
            }

            table.Columns[0].Width = 10;    // STT
            table.Columns[1].Width = 25;    // Tên
            table.Columns[2].Width = 20;    // Lý trình
            table.Columns[3].Width = 15;    // K.cách

            for (int c = 4; c < numCols; c++)
            {
                table.Columns[c].Width = 20;
            }

            // HÀNG 0: Tiêu đề
            table.MergeCells(CellRange.Create(table, 0, 0, 0, numCols - 1));
            table.Cells[0, 0].TextString = $"THỐNG KÊ MATERIAL - {alignmentName}";
            table.Cells[0, 0].TextHeight = 5.0;
            table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

            // HÀNG 1: Nhóm header
            table.MergeCells(CellRange.Create(table, 1, 0, 1, 3));
            table.Cells[1, 0].TextString = "THÔNG TIN TRẮC NGANG";
            table.Cells[1, 0].Alignment = CellAlignment.MiddleCenter;

            if (materials.Count > 0)
            {
                table.MergeCells(CellRange.Create(table, 1, 4, 1, numCols - 1));
                table.Cells[1, 4].TextString = "DIỆN TÍCH VẬT LIỆU (m²)";
                table.Cells[1, 4].Alignment = CellAlignment.MiddleCenter;
            }

            // HÀNG 2: Header chi tiết
            table.Cells[2, 0].TextString = "STT";
            table.Cells[2, 1].TextString = "TÊN TRẮC NGANG";
            table.Cells[2, 2].TextString = "LÝ TRÌNH";
            table.Cells[2, 3].TextString = "K.CÁCH";

            int col = 4;
            foreach (var mat in materials)
            {
                table.Cells[2, col].TextString = mat;
                table.Cells[2, col].Alignment = CellAlignment.MiddleCenter;
                col++;
            }

            // Format header
            for (int c = 0; c < numCols; c++)
            {
                table.Cells[2, c].TextHeight = 3.5;
                table.Cells[2, c].Alignment = CellAlignment.MiddleCenter;
            }

            // DỮ LIỆU
            int row = 3;
            int stt = 1;
            Dictionary<string, double> totals = materials.ToDictionary(m => m, m => 0.0);

            foreach (var stat in stats)
            {
                table.Cells[row, 0].TextString = stt++.ToString();
                table.Cells[row, 1].TextString = stat.SampleLineName;
                table.Cells[row, 2].TextString = stat.StationFormatted;
                table.Cells[row, 3].TextString = Math.Round(stat.SpacingPrev, 2).ToString();

                col = 4;
                foreach (var mat in materials)
                {
                    double area = stat.MaterialAreas.GetValueOrDefault(mat, 0);
                    table.Cells[row, col].TextString = Math.Round(area, 4).ToString();
                    table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                    totals[mat] += area;
                    col++;
                }

                row++;
            }

            // HÀNG TỔNG CỘNG
            table.MergeCells(CellRange.Create(table, row, 0, row, 3));
            table.Cells[row, 0].TextString = "TỔNG CỘNG";
            table.Cells[row, 0].Alignment = CellAlignment.MiddleCenter;

            col = 4;
            foreach (var mat in materials)
            {
                table.Cells[row, col].TextString = Math.Round(totals[mat], 4).ToString();
                table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                col++;
            }

            // TextHeight cho dữ liệu
            for (int r = 3; r < numRows; r++)
            {
                for (int c = 0; c < numCols; c++)
                {
                    table.Cells[r, c].TextHeight = 3.0;
                }
            }

            // Thêm table vào model space
            btr.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
        }
    }
}
