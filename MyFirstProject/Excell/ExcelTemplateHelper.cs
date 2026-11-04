// ══════════════════════════════════════════════════════════════════════════════
// ExcelTemplateHelper.cs — Hỗ trợ đọc/ghi file template Excel (.xlsx)
// Dùng ClosedXML để mở template có sẵn → điền dữ liệu → lưu file mới
//
// DESIGNED BY HUYNH NGOC NAM / TRAFFIC ENGINEER - SĐT: 0378584154
// ══════════════════════════════════════════════════════════════════════════════

using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

namespace MyFirstProject.ExcelModule
{
    /// <summary>
    /// Hỗ trợ làm việc với file template Excel (.xlsx).
    /// <para>
    /// Cách dùng: Mở template có sẵn (đã có định dạng, merge cell, font, border...) 
    /// → điền dữ liệu vào cell cụ thể → lưu ra file mới.
    /// </para>
    /// </summary>
    public static class ExcelTemplateHelper
    {
        /// <summary>
        /// Đường dẫn mặc định tới thư mục chứa template.
        /// Có thể thay đổi nếu cần.
        /// </summary>
        public static string DefaultTemplateFolder { get; set; } =
            Path.GetFullPath(Path.Combine(
                Path.GetDirectoryName(typeof(ExcelTemplateHelper).Assembly.Location) ?? "",
                @"..\..\..\.."));

        // ══════════════════════════════════════════════════════════════════════
        // ĐIỀN DỮ LIỆU VÀO CÁC CELL CỤ THỂ
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Điền dữ liệu vào các cell cụ thể trong template.
        /// </summary>
        /// <param name="templatePath">Đường dẫn file template gốc.</param>
        /// <param name="outputPath">Đường dẫn file kết quả.</param>
        /// <param name="cellValues">Dictionary&lt;CellAddress, Value&gt; (ví dụ: "B3" → "Dự án ABC").</param>
        /// <param name="sheetIndex">Index sheet (0-based, mặc định = 0).</param>
        /// <example>
        /// <code>
        /// ExcelTemplateHelper.FillTemplate(
        ///     @"D:\Autocad 2026_API\TemPlate Duan.xlsx",
        ///     @"C:\output\DuAn_ABC.xlsx",
        ///     new Dictionary&lt;string, object&gt;
        ///     {
        ///         ["B3"] = "Dự án Đường Giao Thông ABC",
        ///         ["B4"] = "Quận 1, TP.HCM",
        ///         ["D5"] = 1250.5,
        ///         ["D6"] = DateTime.Now
        ///     });
        /// </code>
        /// </example>
        public static void FillTemplate(string templatePath, string outputPath,
            Dictionary<string, object> cellValues, int sheetIndex = 0)
        {
            ValidatePaths(templatePath, outputPath);

            using (var workbook = new XLWorkbook(templatePath))
            {
                var ws = workbook.Worksheets.ElementAt(sheetIndex);

                foreach (var kvp in cellValues)
                {
                    SetCellValue(ws.Cell(kvp.Key), kvp.Value);
                }

                EnsureDirectoryExists(outputPath);
                workbook.SaveAs(outputPath);
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // ĐIỀN DỮ LIỆU VÀO BẢNG (TABLE) TRONG TEMPLATE
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Điền danh sách dữ liệu vào bảng bắt đầu từ cell chỉ định.
        /// Giữ nguyên định dạng của dòng đầu, copy xuống các dòng dưới.
        /// </summary>
        /// <param name="templatePath">Đường dẫn file template gốc.</param>
        /// <param name="outputPath">Đường dẫn file kết quả.</param>
        /// <param name="startCell">Cell bắt đầu ghi dữ liệu (ví dụ: "A5").</param>
        /// <param name="rows">Danh sách các dòng dữ liệu.</param>
        /// <param name="sheetIndex">Index sheet (0-based).</param>
        /// <example>
        /// <code>
        /// var rows = new List&lt;object[]&gt;
        /// {
        ///     new object[] { 1, "Km0+000", 105.123, 11.456, 3.5 },
        ///     new object[] { 2, "Km0+100", 105.223, 11.556, 3.5 },
        /// };
        /// ExcelTemplateHelper.FillTable(templatePath, outputPath, "A5", rows);
        /// </code>
        /// </example>
        public static void FillTable(string templatePath, string outputPath,
            string startCell, List<object[]> rows, int sheetIndex = 0)
        {
            ValidatePaths(templatePath, outputPath);

            using (var workbook = new XLWorkbook(templatePath))
            {
                var ws = workbook.Worksheets.ElementAt(sheetIndex);
                var origin = ws.Cell(startCell);
                int startRow = origin.Address.RowNumber;
                int startCol = origin.Address.ColumnNumber;

                // Nếu cần insert thêm dòng (giữ nguyên dòng template đầu tiên)
                if (rows.Count > 1)
                {
                    // Copy format từ dòng template
                    int templateRow = startRow;
                    int colCount = rows[0].Length;

                    for (int i = 1; i < rows.Count; i++)
                    {
                        int targetRow = startRow + i;
                        // Insert dòng mới trước dòng tiếp theo (nếu cần)
                        ws.Row(targetRow).InsertRowsBelow(0); // Trigger row shift nếu cần

                        // Copy style từ dòng template sang dòng mới
                        for (int c = 0; c < colCount; c++)
                        {
                            var sourceCell = ws.Cell(templateRow, startCol + c);
                            var targetCell = ws.Cell(targetRow, startCol + c);
                            targetCell.Style = sourceCell.Style;
                        }
                    }
                }

                // Ghi dữ liệu
                for (int r = 0; r < rows.Count; r++)
                {
                    for (int c = 0; c < rows[r].Length; c++)
                    {
                        var cell = ws.Cell(startRow + r, startCol + c);
                        SetCellValue(cell, rows[r][c]);
                    }
                }

                EnsureDirectoryExists(outputPath);
                workbook.SaveAs(outputPath);
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // TÌM FILE TEMPLATE
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Tìm file template theo tên trong thư mục mặc định.
        /// </summary>
        /// <param name="templateFileName">Tên file template (ví dụ: "TemPlate Duan.xlsx").</param>
        /// <returns>Đường dẫn đầy đủ hoặc null nếu không tìm thấy.</returns>
        public static string? FindTemplate(string templateFileName)
        {
            // Tìm trong thư mục mặc định
            string path = Path.Combine(DefaultTemplateFolder, templateFileName);
            if (File.Exists(path)) return path;

            // Tìm trong thư mục DLL hiện tại
            string? asmDir = Path.GetDirectoryName(typeof(ExcelTemplateHelper).Assembly.Location);
            if (asmDir != null)
            {
                path = Path.Combine(asmDir, templateFileName);
                if (File.Exists(path)) return path;
            }

            // Tìm đệ quy lên thư mục cha (tối đa 4 cấp)
            string? currentDir = asmDir;
            for (int i = 0; i < 4 && currentDir != null; i++)
            {
                currentDir = Path.GetDirectoryName(currentDir);
                if (currentDir != null)
                {
                    path = Path.Combine(currentDir, templateFileName);
                    if (File.Exists(path)) return path;
                }
            }

            return null;
        }

        // ══════════════════════════════════════════════════════════════════════
        // HELPER METHODS
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Gán giá trị cho cell, tự nhận diện kiểu dữ liệu.
        /// </summary>
        private static void SetCellValue(IXLCell cell, object? value)
        {
            if (value == null)
            {
                cell.Value = "";
                return;
            }

            switch (value)
            {
                case double d:
                    cell.Value = d;
                    break;
                case float f:
                    cell.Value = (double)f;
                    break;
                case int i:
                    cell.Value = i;
                    break;
                case long l:
                    cell.Value = (double)l;
                    break;
                case decimal dec:
                    cell.Value = (double)dec;
                    break;
                case DateTime dt:
                    cell.Value = dt;
                    break;
                case bool b:
                    cell.Value = b;
                    break;
                default:
                    cell.Value = value.ToString() ?? "";
                    break;
            }
        }

        private static void ValidatePaths(string templatePath, string outputPath)
        {
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Không tìm thấy file template: {templatePath}");
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Đường dẫn file xuất ra không được để trống.", nameof(outputPath));
        }

        private static void EnsureDirectoryExists(string filePath)
        {
            string? dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
        }
    }
}
