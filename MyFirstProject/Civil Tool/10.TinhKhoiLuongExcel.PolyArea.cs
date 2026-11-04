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

// Tính diện tích / khối lượng từ Polyline vẽ thủ công (CTSV_PolyArea, CTSV_TinhDienTichPoly...)
// (tách từ 10.TinhKhoiLuongExcel.cs — partial class TinhKhoiLuongExcel)
namespace MyFirstProject
{
    public partial class TinhKhoiLuongExcel
    {


        /// <summary>
        /// Lệnh tích hợp tất cả chức năng tính diện tích/khối lượng từ Polyline
        /// Chạy tuần tự: Chọn polyline → Tính diện tích → Tính khối lượng → Ghi lên bản vẽ → Xuất Excel
        /// </summary>
        [CommandMethod("CTSV_PolyArea")]
        public static void CTSVPolyArea()
        {
            try
            {
                A.Ed.WriteMessage("\n╔══════════════════════════════════════════════════════════════╗");
                A.Ed.WriteMessage("\n║   TÍNH KHỐI LƯỢNG TỪ POLYLINE - WORKFLOW TỰ ĐỘNG             ║");
                A.Ed.WriteMessage("\n║   (Chọn polyline 1 lần → Chạy tuần tự 4 bước)                 ║");
                A.Ed.WriteMessage("\n╚══════════════════════════════════════════════════════════════╝");

                // ===== BƯỚC 1: Nhập thông tin =====
                A.Ed.WriteMessage("\n\n📋 BƯỚC 1: NHẬP THÔNG TIN");
                A.Ed.WriteMessage("\n" + new string('─', 50));

                // Nhập tên vật liệu
                PromptStringOptions psoName = new PromptStringOptions("\nNhập tên vật liệu/loại diện tích:")
                {
                    DefaultValue = "Đào nền",
                    AllowSpaces = true
                };
                PromptResult prName = A.Ed.GetString(psoName);
                if (prName.Status != PromptStatus.OK) return;
                string materialName = prName.StringResult;

                // Nhập khoảng cách giữa các mặt cắt
                PromptDoubleOptions pdoSpacing = new PromptDoubleOptions("\nNhập khoảng cách giữa các mặt cắt (m):")
                {
                    AllowNegative = false,
                    AllowZero = false,
                    DefaultValue = 20.0
                };
                PromptDoubleResult pdrSpacing = A.Ed.GetDouble(pdoSpacing);
                if (pdrSpacing.Status != PromptStatus.OK) return;
                double spacing = pdrSpacing.Value;

                // Hỏi có ghi diện tích lên polyline không
                PromptKeywordOptions pkoLabel = new PromptKeywordOptions("\nGhi nhãn diện tích lên polyline? [Yes/No]:")
                {
                    AllowNone = false
                };
                pkoLabel.Keywords.Add("Yes");
                pkoLabel.Keywords.Add("No");
                pkoLabel.Keywords.Default = "Yes";
                PromptResult prLabel = A.Ed.GetKeywords(pkoLabel);
                bool writeLabels = prLabel.Status == PromptStatus.OK && prLabel.StringResult == "Yes";

                double textHeight = 2.5;
                if (writeLabels)
                {
                    PromptDoubleOptions pdoHeight = new PromptDoubleOptions("\nNhập chiều cao text:")
                    {
                        AllowNegative = false,
                        AllowZero = false,
                        DefaultValue = 2.5
                    };
                    PromptDoubleResult pdrHeight = A.Ed.GetDouble(pdoHeight);
                    if (pdrHeight.Status == PromptStatus.OK)
                        textHeight = pdrHeight.Value;
                }

                // ===== BƯỚC 2: Chọn polylines =====
                A.Ed.WriteMessage("\n\n📋 BƯỚC 2: CHỌN POLYLINE");
                A.Ed.WriteMessage("\n" + new string('─', 50));
                A.Ed.WriteMessage("\n⚠️ Chọn theo THỨ TỰ mặt cắt từ đầu đến cuối!");

                TypedValue[] filterList = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<OR"),
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                    new TypedValue((int)DxfCode.Start, "POLYLINE"),
                    new TypedValue((int)DxfCode.Operator, "OR>")
                };
                SelectionFilter filter = new SelectionFilter(filterList);

                PromptSelectionResult psr = A.Ed.GetSelection(filter);
                if (psr.Status != PromptStatus.OK) return;

                SelectionSet ss = psr.Value;
                if (ss.Count < 2)
                {
                    A.Ed.WriteMessage("\n❌ Cần ít nhất 2 polyline để tính khối lượng!");
                    return;
                }

                A.Ed.WriteMessage($"\n✅ Đã chọn {ss.Count} polyline");

                // ===== BƯỚC 3: Tính toán diện tích và khối lượng =====
                A.Ed.WriteMessage("\n\n📋 BƯỚC 3: TÍNH TOÁN");
                A.Ed.WriteMessage("\n" + new string('─', 50));

                List<PolyAreaData> dataList = new List<PolyAreaData>();
                List<ObjectId> polylineIds = new List<ObjectId>();
                double totalVolume = 0;
                double prevArea = 0;
                double totalArea = 0;

                using (Transaction tr = A.Db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(A.Db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    int index = 1;
                    foreach (ObjectId id in ss.GetObjectIds())
                    {
                        polylineIds.Add(id);
                        Autodesk.AutoCAD.DatabaseServices.Entity ent = tr.GetObject(id, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                        
                        double area = 0;
                        Point3d centroid = Point3d.Origin;

                        if (ent is Polyline pline)
                        {
                            area = pline.Closed ? pline.Area : CalculateOpenPolylineArea(pline);
                            centroid = GetPolylineCentroid(pline);

                            // Ghi nhãn nếu được yêu cầu
                            if (writeLabels)
                            {
                                DBText areaText = new DBText
                                {
                                    Position = centroid,
                                    TextString = $"S={Math.Round(area, AreaDecimalPlaces):F2}",
                                    Height = textHeight,
                                    Layer = pline.Layer
                                };
                                btr.AppendEntity(areaText);
                                tr.AddNewlyCreatedDBObject(areaText, true);
                            }
                        }
                        else if (ent is Curve curve)
                        {
                            area = curve.Area;
                        }

                        area = Math.Round(area, AreaDecimalPlaces);
                        totalArea += area;
                        
                        double segmentVolume = 0;
                        if (index > 1)
                        {
                            segmentVolume = CalculateVolume(prevArea, area, spacing, true);
                            totalVolume += segmentVolume;
                        }

                        dataList.Add(new PolyAreaData
                        {
                            STT = index,
                            SectionName = $"MC{index}",
                            Station = (index - 1) * spacing,
                            Area = area,
                            SegmentVolume = segmentVolume,
                            CumulativeVolume = Math.Round(totalVolume, VolumeDecimalPlaces)
                        });

                        prevArea = area;
                        index++;
                    }

                    tr.Commit();
                }

                // ===== BƯỚC 4: Hiển thị kết quả =====
                A.Ed.WriteMessage("\n\n📋 BƯỚC 4: KẾT QUẢ TÍNH TOÁN");
                A.Ed.WriteMessage("\n" + new string('─', 50));
                A.Ed.WriteMessage($"\n   Vật liệu: {materialName}");
                A.Ed.WriteMessage($"\n   Khoảng cách mặt cắt: {spacing} m");
                A.Ed.WriteMessage($"\n   Công thức: V = (S1 + S2) / 2 × L");
                A.Ed.WriteMessage($"\n\n   {"MC",-6} {"Lý trình",-12} {"S (m²)",-12} {"V đoạn (m³)",-14} {"V cộng dồn",-12}");
                A.Ed.WriteMessage($"\n   {new string('─', 56)}");

                foreach (var data in dataList)
                {
                    A.Ed.WriteMessage($"\n   {data.SectionName,-6} {data.Station,-12:F2} {data.Area,-12:F2} {data.SegmentVolume,-14:F2} {data.CumulativeVolume,-12:F2}");
                }

                A.Ed.WriteMessage($"\n   {new string('─', 56)}");
                A.Ed.WriteMessage($"\n   {"TỔNG",-6} {"",-12} {totalArea,-12:F2} {Math.Round(totalVolume, VolumeDecimalPlaces),-14:F2} {Math.Round(totalVolume, VolumeDecimalPlaces),-12:F2}");

                // ===== BƯỚC 5: Xuất Excel =====
                A.Ed.WriteMessage("\n\n📋 BƯỚC 5: XUẤT FILE EXCEL");
                A.Ed.WriteMessage("\n" + new string('─', 50));

                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string safeMatName = string.Join("_", materialName.Split(System.IO.Path.GetInvalidFileNameChars()));
                string fileName = $"KhoiLuong_{safeMatName}_{timestamp}.xlsx";
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filePath = System.IO.Path.Combine(desktopPath, fileName);

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Khối lượng");

                    // Tiêu đề
                    worksheet.Cell(1, 1).Value = $"BẢNG TÍNH KHỐI LƯỢNG - {materialName.ToUpper()}";
                    worksheet.Range(1, 1, 1, 6).Merge();
                    worksheet.Cell(1, 1).Style.Font.Bold = true;
                    worksheet.Cell(1, 1).Style.Font.FontSize = 14;
                    worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    worksheet.Cell(2, 1).Value = $"Khoảng cách mặt cắt: {spacing} m | Phương pháp: Average End Area";
                    worksheet.Range(2, 1, 2, 6).Merge();

                    worksheet.Cell(3, 1).Value = $"Ngày xuất: {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
                    worksheet.Range(3, 1, 3, 6).Merge();

                    // Header
                    worksheet.Cell(5, 1).Value = "STT";
                    worksheet.Cell(5, 2).Value = "Mặt cắt";
                    worksheet.Cell(5, 3).Value = "Lý trình (m)";
                    worksheet.Cell(5, 4).Value = "Diện tích (m²)";
                    worksheet.Cell(5, 5).Value = "KL đoạn (m³)";
                    worksheet.Cell(5, 6).Value = "KL cộng dồn (m³)";

                    var headerRange = worksheet.Range(5, 1, 5, 6);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    // Data
                    int row = 6;
                    foreach (var data in dataList)
                    {
                        worksheet.Cell(row, 1).Value = data.STT;
                        worksheet.Cell(row, 2).Value = data.SectionName;
                        worksheet.Cell(row, 3).Value = data.Station;
                        worksheet.Cell(row, 4).Value = data.Area;
                        worksheet.Cell(row, 5).Value = data.SegmentVolume;
                        worksheet.Cell(row, 6).Value = data.CumulativeVolume;

                        // Border
                        worksheet.Range(row, 1, row, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        worksheet.Range(row, 1, row, 6).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        row++;
                    }

                    // Hàng tổng
                    worksheet.Cell(row, 1).Value = "TỔNG CỘNG";
                    worksheet.Range(row, 1, row, 3).Merge();
                    worksheet.Cell(row, 1).Style.Font.Bold = true;
                    worksheet.Cell(row, 4).Value = Math.Round(totalArea, AreaDecimalPlaces);
                    worksheet.Cell(row, 4).Style.Font.Bold = true;
                    worksheet.Cell(row, 5).Value = Math.Round(totalVolume, VolumeDecimalPlaces);
                    worksheet.Cell(row, 5).Style.Font.Bold = true;
                    worksheet.Cell(row, 6).Value = Math.Round(totalVolume, VolumeDecimalPlaces);
                    worksheet.Cell(row, 6).Style.Font.Bold = true;
                    worksheet.Range(row, 1, row, 6).Style.Fill.BackgroundColor = XLColor.LightGreen;
                    worksheet.Range(row, 1, row, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;

                    // Định dạng cột
                    worksheet.Columns().AdjustToContents();

                    workbook.SaveAs(filePath);
                }

                A.Ed.WriteMessage($"\n✅ Đã xuất file: {filePath}");

                // ===== TỔNG KẾT =====
                A.Ed.WriteMessage("\n\n╔══════════════════════════════════════════════════════════════╗");
                A.Ed.WriteMessage("\n║                    📊 TỔNG KẾT                               ║");
                A.Ed.WriteMessage("\n╠══════════════════════════════════════════════════════════════╣");
                A.Ed.WriteMessage($"\n║   Vật liệu:         {materialName,-40} ║");
                A.Ed.WriteMessage($"\n║   Số mặt cắt:       {dataList.Count,-40} ║");
                A.Ed.WriteMessage($"\n║   Tổng diện tích:   {Math.Round(totalArea, AreaDecimalPlaces):F2} m²{new string(' ', 32)}║");
                A.Ed.WriteMessage($"\n║   TỔNG KHỐI LƯỢNG:  {Math.Round(totalVolume, VolumeDecimalPlaces):F2} m³{new string(' ', 32)}║");
                A.Ed.WriteMessage("\n╚══════════════════════════════════════════════════════════════╝");

                if (writeLabels)
                    A.Ed.WriteMessage($"\n✅ Đã ghi {dataList.Count} nhãn diện tích lên bản vẽ");

                A.Ed.WriteMessage("\n\n=== WORKFLOW HOÀN TẤT ===\n");

                // Mở file Excel
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Tính diện tích polyline và xuất Excel
        /// </summary>
        [CommandMethod("CTSV_TinhDienTichPolyExcel")]
        public static void CTSVTinhDienTichPolyExcel()
        {
            try
            {
                A.Ed.WriteMessage("\n=== TÍNH DIỆN TÍCH POLYLINE + XUẤT EXCEL ===");

                // Nhập tên vật liệu/loại diện tích
                PromptStringOptions psoName = new PromptStringOptions("\nNhập tên vật liệu/loại diện tích:")
                {
                    DefaultValue = "Vật liệu",
                    AllowSpaces = true
                };
                PromptResult prName = A.Ed.GetString(psoName);
                if (prName.Status != PromptStatus.OK) return;
                string materialName = prName.StringResult;

                // Nhập khoảng cách giữa các mặt cắt
                PromptDoubleOptions pdoSpacing = new PromptDoubleOptions("\nNhập khoảng cách giữa các mặt cắt (m):")
                {
                    AllowNegative = false,
                    AllowZero = false,
                    DefaultValue = 20.0
                };
                PromptDoubleResult pdrSpacing = A.Ed.GetDouble(pdoSpacing);
                if (pdrSpacing.Status != PromptStatus.OK) return;
                double spacing = pdrSpacing.Value;

                // Chọn polylines
                TypedValue[] filterList = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<OR"),
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                    new TypedValue((int)DxfCode.Start, "POLYLINE"),
                    new TypedValue((int)DxfCode.Operator, "OR>")
                };
                SelectionFilter filter = new SelectionFilter(filterList);

                A.Ed.WriteMessage("\n📌 Chọn các polyline theo thứ tự mặt cắt:");
                PromptSelectionResult psr = A.Ed.GetSelection(filter);
                if (psr.Status != PromptStatus.OK) return;

                SelectionSet ss = psr.Value;
                if (ss.Count < 1)
                {
                    A.Ed.WriteMessage("\n❌ Không có polyline nào được chọn!");
                    return;
                }

                A.Ed.WriteMessage($"\n✅ Đã chọn {ss.Count} polyline");

                // Lấy diện tích
                List<PolyAreaData> dataList = new List<PolyAreaData>();
                double totalVolume = 0;
                double prevArea = 0;

                using (Transaction tr = A.Db.TransactionManager.StartTransaction())
                {
                    int index = 1;
                    foreach (ObjectId id in ss.GetObjectIds())
                    {
                        Autodesk.AutoCAD.DatabaseServices.Entity ent = tr.GetObject(id, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                        
                        double area = 0;
                        if (ent is Polyline pline)
                        {
                            area = pline.Closed ? pline.Area : CalculateOpenPolylineArea(pline);
                        }
                        else if (ent is Curve curve)
                        {
                            area = curve.Area;
                        }

                        area = Math.Round(area, AreaDecimalPlaces);
                        
                        double segmentVolume = 0;
                        if (index > 1)
                        {
                            segmentVolume = CalculateVolume(prevArea, area, spacing, true);
                            totalVolume += segmentVolume;
                        }

                        dataList.Add(new PolyAreaData
                        {
                            STT = index,
                            SectionName = $"MC{index}",
                            Station = (index - 1) * spacing,
                            Area = area,
                            SegmentVolume = segmentVolume,
                            CumulativeVolume = Math.Round(totalVolume, VolumeDecimalPlaces)
                        });

                        prevArea = area;
                        index++;
                    }

                    tr.Commit();
                }

                // Xuất Excel
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fileName = $"KhoiLuong_Poly_{materialName}_{timestamp}.xlsx";
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filePath = System.IO.Path.Combine(desktopPath, fileName);

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Khối lượng");

                    // Tiêu đề
                    worksheet.Cell(1, 1).Value = $"BẢNG TÍNH KHỐI LƯỢNG - {materialName.ToUpper()}";
                    worksheet.Range(1, 1, 1, 6).Merge();
                    worksheet.Cell(1, 1).Style.Font.Bold = true;
                    worksheet.Cell(1, 1).Style.Font.FontSize = 14;
                    worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    worksheet.Cell(2, 1).Value = $"Khoảng cách mặt cắt: {spacing} m";
                    worksheet.Range(2, 1, 2, 6).Merge();

                    // Header
                    worksheet.Cell(4, 1).Value = "STT";
                    worksheet.Cell(4, 2).Value = "Mặt cắt";
                    worksheet.Cell(4, 3).Value = "Lý trình (m)";
                    worksheet.Cell(4, 4).Value = "Diện tích (m²)";
                    worksheet.Cell(4, 5).Value = "KL đoạn (m³)";
                    worksheet.Cell(4, 6).Value = "KL cộng dồn (m³)";

                    // Định dạng header
                    var headerRange = worksheet.Range(4, 1, 4, 6);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    // Data
                    int row = 5;
                    foreach (var data in dataList)
                    {
                        worksheet.Cell(row, 1).Value = data.STT;
                        worksheet.Cell(row, 2).Value = data.SectionName;
                        worksheet.Cell(row, 3).Value = data.Station;
                        worksheet.Cell(row, 4).Value = data.Area;
                        worksheet.Cell(row, 5).Value = data.SegmentVolume;
                        worksheet.Cell(row, 6).Value = data.CumulativeVolume;
                        row++;
                    }

                    // Hàng tổng
                    worksheet.Cell(row, 1).Value = "TỔNG";
                    worksheet.Range(row, 1, row, 4).Merge();
                    worksheet.Cell(row, 1).Style.Font.Bold = true;
                    worksheet.Cell(row, 5).Value = Math.Round(totalVolume, VolumeDecimalPlaces);
                    worksheet.Cell(row, 5).Style.Font.Bold = true;
                    worksheet.Cell(row, 6).Value = Math.Round(totalVolume, VolumeDecimalPlaces);
                    worksheet.Cell(row, 6).Style.Font.Bold = true;

                    // Định dạng cột
                    worksheet.Columns().AdjustToContents();

                    workbook.SaveAs(filePath);
                }

                A.Ed.WriteMessage($"\n\n📊 TỔNG KẾT:");
                A.Ed.WriteMessage($"\n   Số mặt cắt: {dataList.Count}");
                A.Ed.WriteMessage($"\n   Tổng khối lượng: {Math.Round(totalVolume, VolumeDecimalPlaces):F2} m³");
                A.Ed.WriteMessage($"\n\n📁 Đã xuất file: {filePath}");
                A.Ed.WriteMessage("\n=== HOÀN TẤT ===\n");

                // Mở file Excel
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Class lưu dữ liệu diện tích polyline
        /// </summary>
        public class PolyAreaData
        {
            public int STT { get; set; }
            public string SectionName { get; set; } = "";
            public double Station { get; set; }
            public double Area { get; set; }
            public double SegmentVolume { get; set; }
            public double CumulativeVolume { get; set; }
        }

        /// <summary>
        /// Lệnh tính diện tích từ các polyline được chọn
        /// Tương tự cách Toolcivil2025 sử dụng ((Curve)polyline).Area
        /// </summary>
        [CommandMethod("CTSV_TinhDienTichPoly")]
        public static void CTSVTinhDienTichPoly()
        {
            try
            {
                A.Ed.WriteMessage("\n=== TÍNH DIỆN TÍCH TỪ POLYLINE VẼ THỦ CÔNG ===");
                A.Ed.WriteMessage("\n📌 Chọn các polyline đã vẽ bao quanh vùng diện tích:");

                // Chọn polylines
                PromptSelectionOptions pso = new PromptSelectionOptions();
                pso.MessageForAdding = "\nChọn các polyline (Closed hoặc Open):";
                pso.AllowDuplicates = false;

                TypedValue[] filterList = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<OR"),
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                    new TypedValue((int)DxfCode.Start, "POLYLINE"),
                    new TypedValue((int)DxfCode.Operator, "OR>")
                };
                SelectionFilter filter = new SelectionFilter(filterList);

                PromptSelectionResult psr = A.Ed.GetSelection(pso, filter);
                if (psr.Status != PromptStatus.OK)
                {
                    A.Ed.WriteMessage("\n❌ Không có polyline nào được chọn.");
                    return;
                }

                SelectionSet ss = psr.Value;
                A.Ed.WriteMessage($"\n✅ Đã chọn {ss.Count} polyline");

                double totalArea = 0;
                int count = 0;
                List<PolylineAreaInfo> areaInfos = new List<PolylineAreaInfo>();

                using (Transaction tr = A.Db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in ss.GetObjectIds())
                    {
                        Autodesk.AutoCAD.DatabaseServices.Entity ent = tr.GetObject(id, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                        
                        if (ent is Curve curve)
                        {
                            double area = 0;
                            bool isClosed = false;

                            if (ent is Polyline pline)
                            {
                                isClosed = pline.Closed;
                                if (isClosed)
                                {
                                    area = pline.Area;
                                }
                                else
                                {
                                    // Nếu polyline không đóng, tính diện tích bằng cách giả lập đóng
                                    area = CalculateOpenPolylineArea(pline);
                                }
                            }
                            else if (ent is Polyline2d pline2d)
                            {
                                isClosed = pline2d.Closed;
                                area = pline2d.Area;
                            }
                            else if (ent is Polyline3d pline3d)
                            {
                                isClosed = pline3d.Closed;
                                area = pline3d.Area;
                            }

                            area = Math.Round(area, AreaDecimalPlaces);
                            totalArea += area;
                            count++;

                            string layer = ent.Layer;
                            areaInfos.Add(new PolylineAreaInfo
                            {
                                Index = count,
                                Layer = layer,
                                IsClosed = isClosed,
                                Area = area
                            });

                            A.Ed.WriteMessage($"\n  {count}. Layer [{layer}] - {(isClosed ? "Đóng" : "Mở")} - Diện tích: {area:F2} m²");
                        }
                    }

                    tr.Commit();
                }

                totalArea = Math.Round(totalArea, AreaDecimalPlaces);
                A.Ed.WriteMessage($"\n\n📊 TỔNG KẾT:");
                A.Ed.WriteMessage($"\n   Số polyline: {count}");
                A.Ed.WriteMessage($"\n   Tổng diện tích: {totalArea:F2} m²");
                A.Ed.WriteMessage("\n=== HOÀN TẤT ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Lệnh tính khối lượng từ các polyline vẽ thủ công (theo từng mặt cắt)
        /// </summary>
        [CommandMethod("CTSV_TinhKhoiLuongPoly")]
        public static void CTSVTinhKhoiLuongPoly()
        {
            try
            {
                A.Ed.WriteMessage("\n=== TÍNH KHỐI LƯỢNG TỪ POLYLINE VẼ THỦ CÔNG ===");
                A.Ed.WriteMessage("\n📌 Tính toán khối lượng bằng phương pháp Average End Area");
                A.Ed.WriteMessage("\n📌 Chọn các polyline theo thứ tự mặt cắt từ đầu đến cuối");

                // Nhập khoảng cách giữa các mặt cắt
                PromptDoubleOptions pdoSpacing = new PromptDoubleOptions("\nNhập khoảng cách giữa các mặt cắt (m):")
                {
                    AllowNegative = false,
                    AllowZero = false,
                    DefaultValue = 20.0
                };
                PromptDoubleResult pdrSpacing = A.Ed.GetDouble(pdoSpacing);
                if (pdrSpacing.Status != PromptStatus.OK)
                {
                    return;
                }
                double spacing = pdrSpacing.Value;

                // Chọn polylines theo thứ tự
                A.Ed.WriteMessage("\n📌 Chọn các polyline theo thứ tự mặt cắt:");

                TypedValue[] filterList = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<OR"),
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                    new TypedValue((int)DxfCode.Start, "POLYLINE"),
                    new TypedValue((int)DxfCode.Operator, "OR>")
                };
                SelectionFilter filter = new SelectionFilter(filterList);

                PromptSelectionResult psr = A.Ed.GetSelection(filter);
                if (psr.Status != PromptStatus.OK)
                {
                    A.Ed.WriteMessage("\n❌ Không có polyline nào được chọn.");
                    return;
                }

                SelectionSet ss = psr.Value;
                A.Ed.WriteMessage($"\n✅ Đã chọn {ss.Count} polyline");

                if (ss.Count < 2)
                {
                    A.Ed.WriteMessage("\n❌ Cần ít nhất 2 polyline để tính khối lượng!");
                    return;
                }

                List<double> areas = new List<double>();

                using (Transaction tr = A.Db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in ss.GetObjectIds())
                    {
                        Autodesk.AutoCAD.DatabaseServices.Entity ent = tr.GetObject(id, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                        
                        if (ent is Polyline pline)
                        {
                            double area = pline.Closed ? pline.Area : CalculateOpenPolylineArea(pline);
                            areas.Add(Math.Round(area, AreaDecimalPlaces));
                        }
                        else if (ent is Curve curve)
                        {
                            areas.Add(Math.Round(curve.Area, AreaDecimalPlaces));
                        }
                    }

                    tr.Commit();
                }

                // Tính khối lượng
                double totalVolume = 0;
                A.Ed.WriteMessage($"\n\n📊 CHI TIẾT TÍNH KHỐI LƯỢNG:");
                A.Ed.WriteMessage($"\n   Khoảng cách giữa các mặt cắt: {spacing} m");
                A.Ed.WriteMessage($"\n   Công thức: V = (S1 + S2) / 2 × L\n");
                A.Ed.WriteMessage($"\n   {"MC",-5} {"S (m²)",-12} {"V (m³)",-12} {"Cộng dồn (m³)",-15}");
                A.Ed.WriteMessage($"\n   {new string('-', 45)}");

                for (int i = 0; i < areas.Count; i++)
                {
                    double segmentVolume = 0;
                    if (i > 0)
                    {
                        segmentVolume = CalculateVolume(areas[i - 1], areas[i], spacing, true);
                        totalVolume += segmentVolume;
                    }

                    A.Ed.WriteMessage($"\n   {i + 1,-5} {areas[i],-12:F2} {segmentVolume,-12:F2} {totalVolume,-15:F2}");
                }

                totalVolume = Math.Round(totalVolume, VolumeDecimalPlaces);
                A.Ed.WriteMessage($"\n\n   📦 TỔNG KHỐI LƯỢNG: {totalVolume:F2} m³");
                A.Ed.WriteMessage("\n=== HOÀN TẤT ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Tính diện tích cho polyline mở (không đóng) bằng cách giả lập đóng
        /// </summary>
        private static double CalculateOpenPolylineArea(Polyline pline)
        {
            if (pline.NumberOfVertices < 3) return 0;

            // Sử dụng Shoelace formula
            List<Point2d> points = new List<Point2d>();
            for (int i = 0; i < pline.NumberOfVertices; i++)
            {
                Point2d pt = pline.GetPoint2dAt(i);
                points.Add(pt);
            }

            return CalculatePolygonArea(points);
        }

        /// <summary>
        /// Lệnh tính và ghi diện tích lên từng polyline
        /// </summary>
        [CommandMethod("CTSV_GhiDienTichPoly")]
        public static void CTSVGhiDienTichPoly()
        {
            try
            {
                A.Ed.WriteMessage("\n=== GHI DIỆN TÍCH LÊN POLYLINE ===");

                // Lấy chiều cao text
                PromptDoubleOptions pdoHeight = new PromptDoubleOptions("\nNhập chiều cao text:")
                {
                    AllowNegative = false,
                    AllowZero = false,
                    DefaultValue = 2.5
                };
                PromptDoubleResult pdrHeight = A.Ed.GetDouble(pdoHeight);
                if (pdrHeight.Status != PromptStatus.OK) return;
                double textHeight = pdrHeight.Value;

                // Chọn polylines
                TypedValue[] filterList = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<OR"),
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                    new TypedValue((int)DxfCode.Start, "POLYLINE"),
                    new TypedValue((int)DxfCode.Operator, "OR>")
                };
                SelectionFilter filter = new SelectionFilter(filterList);

                PromptSelectionResult psr = A.Ed.GetSelection(filter);
                if (psr.Status != PromptStatus.OK)
                {
                    A.Ed.WriteMessage("\n❌ Không có polyline nào được chọn.");
                    return;
                }

                SelectionSet ss = psr.Value;
                int count = 0;

                using (Transaction tr = A.Db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(A.Db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    foreach (ObjectId id in ss.GetObjectIds())
                    {
                        Autodesk.AutoCAD.DatabaseServices.Entity ent = tr.GetObject(id, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                        
                        if (ent is Polyline pline)
                        {
                            double area = pline.Closed ? pline.Area : CalculateOpenPolylineArea(pline);
                            area = Math.Round(area, AreaDecimalPlaces);

                            // Tính tâm polyline
                            Point3d centroid = GetPolylineCentroid(pline);

                            // Tạo text hiển thị diện tích
                            DBText areaText = new DBText
                            {
                                Position = centroid,
                                TextString = $"S={area:F2} m²",
                                Height = textHeight,
                                Layer = ent.Layer
                            };

                            btr.AppendEntity(areaText);
                            tr.AddNewlyCreatedDBObject(areaText, true);
                            count++;
                        }
                    }

                    tr.Commit();
                }

                A.Ed.WriteMessage($"\n✅ Đã ghi {count} diện tích lên các polyline.");
                A.Ed.WriteMessage("\n=== HOÀN TẤT ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Tính tâm (centroid) của polyline
        /// </summary>
        private static Point3d GetPolylineCentroid(Polyline pline)
        {
            double sumX = 0, sumY = 0;
            int count = pline.NumberOfVertices;

            for (int i = 0; i < count; i++)
            {
                Point2d pt = pline.GetPoint2dAt(i);
                sumX += pt.X;
                sumY += pt.Y;
            }

            return new Point3d(sumX / count, sumY / count, 0);
        }

        /// <summary>
        /// Class lưu thông tin diện tích polyline
        /// </summary>
        public class PolylineAreaInfo
        {
            public int Index { get; set; }
            public string Layer { get; set; } = "";
            public bool IsClosed { get; set; }
            public double Area { get; set; }
        }
    }
}
