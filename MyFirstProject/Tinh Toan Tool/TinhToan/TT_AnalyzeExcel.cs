using System;
using System.IO;
using System.Text;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using ClosedXML.Excel;

[assembly: CommandClass(typeof(TinhToan.TT_AnalyzeExcelCommands))]

namespace TinhToan
{
    public class TT_AnalyzeExcelCommands
    {
        [CommandMethod("CTS_PHANTICHEXCEL")]
        public void AnalyzeExcel()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            string excelPath = @"E:\3. CÔNG THỨC TÍNH\1. Ket cau ao duong\BTN\2. KCAD LOAI 1 SAN BAY.xlsx";
            string outputPath = @"C:\tmp\Excel_Analysis.txt";

            try
            {
                if (!Directory.Exists(@"C:\tmp"))
                {
                    Directory.CreateDirectory(@"C:\tmp");
                }

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("=== PHÂN TÍCH FILE EXCEL: " + Path.GetFileName(excelPath) + " ===");

                using (var workbook = new XLWorkbook(excelPath))
                {
                    foreach (var sheet in workbook.Worksheets)
                    {
                        sb.AppendLine($"\n\n--- SHEET: {sheet.Name} ---");
                        // We only want to dump cells that have formulas or important constants
                        // Let's scan a grid: A1 to Z100
                        var usedRange = sheet.RangeUsed();
                        if (usedRange == null) continue;

                        int rowCount = Math.Min(usedRange.RowCount(), 200); // Limit to top 200 rows
                        int colCount = Math.Min(usedRange.ColumnCount(), 20); // Limit to 20 cols

                        for (int r = 1; r <= rowCount; r++)
                        {
                            for (int c = 1; c <= colCount; c++)
                            {
                                var cell = sheet.Cell(r, c);
                                if (!cell.IsEmpty())
                                {
                                    string addr = cell.Address?.ToString() ?? "";
                                    if (cell.HasFormula)
                                    {
                                        sb.AppendLine($"[{addr}] FORMULA: {cell.FormulaA1}  => VAL: {cell.Value}");
                                    }
                                    else
                                    {
                                        // Just raw value, could be a label or parameter
                                        string valStr = cell.Value.ToString() ?? "";
                                        if (valStr.Length > 0 && valStr.Length < 50) 
                                        {
                                            sb.AppendLine($"[{addr}] VAL: {valStr}");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                File.WriteAllText(outputPath, sb.ToString());
                doc.Editor.WriteMessage($"\n[CIVIL TOOL] Phân tích thành công! Đã lưu kết quả tại: {outputPath}");
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\n[CIVIL TOOL ERROR] Lỗi phân tích Excel: {ex.Message}");
            }
        }
    }
}
