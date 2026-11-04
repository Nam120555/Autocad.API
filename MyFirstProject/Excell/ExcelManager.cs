// ══════════════════════════════════════════════════════════════════════════════
// ExcelManager.cs — Lớp trung tâm quản lý mọi thao tác đọc/ghi Excel
// Kết hợp ClosedXML (nhanh, không cần Excel) + COM Interop (cần Excel desktop)
//
// DESIGNED BY HUYNH NGOC NAM / TRAFFIC ENGINEER - SĐT: 0378584154
// ══════════════════════════════════════════════════════════════════════════════

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace MyFirstProject.ExcelModule
{
    /// <summary>
    /// Quản lý tập trung mọi thao tác đọc/ghi Excel trong dự án.
    /// <para>
    /// <b>Quy tắc sử dụng:</b>
    /// <list type="bullet">
    ///   <item>Gọi <c>ExcelDependencyLoader.EnsureLoaded()</c> trước khi dùng bất kỳ method nào (AutoLoader đã gọi sẵn).</item>
    ///   <item>Dùng các method ClosedXML (CreateWorkbook, ReadWorkbook, WriteToTemplate) khi KHÔNG cần Excel desktop.</item>
    ///   <item>Dùng các method Interop (ExportToPdf, RunMacro, OpenInExcel) khi CẦN Excel desktop.</item>
    /// </list>
    /// </para>
    /// </summary>
    public static class ExcelManager
    {
        // ══════════════════════════════════════════════════════════════════════
        // CLOSEDXML ENGINE — Đọc/ghi file .xlsx (không cần cài Excel)
        // ══════════════════════════════════════════════════════════════════════

        #region ClosedXML — Tạo Workbook mới

        /// <summary>
        /// Tạo file Excel (.xlsx) mới từ callback configure.
        /// </summary>
        /// <param name="outputPath">Đường dẫn file xuất ra.</param>
        /// <param name="configure">Callback để cấu hình workbook (thêm sheet, dữ liệu, định dạng...).</param>
        /// <example>
        /// <code>
        /// ExcelManager.CreateWorkbook(@"C:\output\report.xlsx", wb =>
        /// {
        ///     var ws = wb.Worksheets.Add("Data");
        ///     ws.Cell("A1").Value = "Hello Excel!";
        /// });
        /// </code>
        /// </example>
        public static void CreateWorkbook(string outputPath, Action<XLWorkbook> configure)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Đường dẫn file không được để trống.", nameof(outputPath));
            if (configure == null)
                throw new ArgumentNullException(nameof(configure));

            // Đảm bảo thư mục tồn tại
            string? dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using (var workbook = new XLWorkbook())
            {
                configure(workbook);
                workbook.SaveAs(outputPath);
            }
        }

        #endregion

        #region ClosedXML — Đọc dữ liệu từ file .xlsx

        /// <summary>
        /// Đọc toàn bộ dữ liệu từ một sheet trong file Excel.
        /// Dòng đầu tiên được coi là header (tên cột).
        /// </summary>
        /// <param name="filePath">Đường dẫn file .xlsx.</param>
        /// <param name="sheetName">Tên sheet (null = sheet đầu tiên).</param>
        /// <returns>Danh sách các dòng, mỗi dòng là Dictionary&lt;TênCột, GiáTrị&gt;.</returns>
        /// <example>
        /// <code>
        /// var data = ExcelManager.ReadWorkbook(@"C:\data\input.xlsx");
        /// foreach (var row in data)
        /// {
        ///     string name = row["Tên trụ"]?.ToString() ?? "";
        ///     double x = Convert.ToDouble(row["X"]);
        /// }
        /// </code>
        /// </example>
        public static List<Dictionary<string, object>> ReadWorkbook(string filePath, string? sheetName = null)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Không tìm thấy file Excel: {filePath}");

            var result = new List<Dictionary<string, object>>();

            using (var workbook = new XLWorkbook(filePath))
            {
                IXLWorksheet ws = string.IsNullOrEmpty(sheetName)
                    ? workbook.Worksheets.First()
                    : workbook.Worksheet(sheetName);

                var range = ws.RangeUsed();
                if (range == null) return result;

                int firstRow = range.FirstRow().RowNumber();
                int lastRow = range.LastRow().RowNumber();
                int firstCol = range.FirstColumn().ColumnNumber();
                int lastCol = range.LastColumn().ColumnNumber();

                // Đọc header từ dòng đầu
                var headers = new List<string>();
                for (int col = firstCol; col <= lastCol; col++)
                {
                    string header = ws.Cell(firstRow, col).GetString().Trim();
                    headers.Add(string.IsNullOrEmpty(header) ? $"Column{col}" : header);
                }

                // Đọc dữ liệu từ dòng 2 trở đi
                for (int row = firstRow + 1; row <= lastRow; row++)
                {
                    var rowData = new Dictionary<string, object>();
                    bool hasData = false;

                    for (int col = firstCol; col <= lastCol; col++)
                    {
                        var cell = ws.Cell(row, col);
                        string key = headers[col - firstCol];
                        object value;

                        if (cell.DataType == XLDataType.Number)
                            value = cell.GetDouble();
                        else if (cell.DataType == XLDataType.DateTime)
                            value = cell.GetDateTime();
                        else if (cell.DataType == XLDataType.Boolean)
                            value = cell.GetBoolean();
                        else
                            value = cell.GetString();

                        rowData[key] = value;
                        if (value != null && value.ToString() != "") hasData = true;
                    }

                    if (hasData) result.Add(rowData);
                }
            }

            return result;
        }

        /// <summary>
        /// Đọc dữ liệu dạng mảng 2D (không dùng header).
        /// Phù hợp khi file Excel không có dòng tiêu đề chuẩn.
        /// </summary>
        /// <param name="filePath">Đường dẫn file .xlsx.</param>
        /// <param name="sheetName">Tên sheet (null = sheet đầu tiên).</param>
        /// <returns>Mảng 2D object[hàng][cột].</returns>
        public static List<List<object>> ReadWorkbookRaw(string filePath, string? sheetName = null)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Không tìm thấy file Excel: {filePath}");

            var result = new List<List<object>>();

            using (var workbook = new XLWorkbook(filePath))
            {
                IXLWorksheet ws = string.IsNullOrEmpty(sheetName)
                    ? workbook.Worksheets.First()
                    : workbook.Worksheet(sheetName);

                var range = ws.RangeUsed();
                if (range == null) return result;

                int lastRow = range.LastRow().RowNumber();
                int lastCol = range.LastColumn().ColumnNumber();

                for (int row = 1; row <= lastRow; row++)
                {
                    var rowData = new List<object>();
                    for (int col = 1; col <= lastCol; col++)
                    {
                        var cell = ws.Cell(row, col);
                        if (cell.DataType == XLDataType.Number)
                            rowData.Add(cell.GetDouble());
                        else if (cell.DataType == XLDataType.DateTime)
                            rowData.Add(cell.GetDateTime());
                        else
                            rowData.Add(cell.GetString());
                    }
                    result.Add(rowData);
                }
            }

            return result;
        }

        #endregion

        #region ClosedXML — Ghi dữ liệu vào Template

        /// <summary>
        /// Mở file template Excel, điền dữ liệu, lưu ra file mới.
        /// Template giữ nguyên định dạng, merge cell, font, border...
        /// </summary>
        /// <param name="templatePath">Đường dẫn file template gốc (.xlsx).</param>
        /// <param name="outputPath">Đường dẫn file kết quả.</param>
        /// <param name="fill">Callback để điền dữ liệu vào workbook.</param>
        /// <example>
        /// <code>
        /// ExcelManager.WriteToTemplate(
        ///     @"D:\Autocad 2026_API\TemPlate Duan.xlsx",
        ///     @"C:\output\DuAn_ABC.xlsx",
        ///     wb =>
        ///     {
        ///         var ws = wb.Worksheet(1);
        ///         ws.Cell("B3").Value = "Dự án ABC";
        ///         ws.Cell("B4").Value = "Quận 1, TP.HCM";
        ///     });
        /// </code>
        /// </example>
        public static void WriteToTemplate(string templatePath, string outputPath, Action<XLWorkbook> fill)
        {
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Không tìm thấy file template: {templatePath}");
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Đường dẫn file xuất ra không được để trống.", nameof(outputPath));
            if (fill == null)
                throw new ArgumentNullException(nameof(fill));

            string? dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using (var workbook = new XLWorkbook(templatePath))
            {
                fill(workbook);
                workbook.SaveAs(outputPath);
            }
        }

        #endregion

        #region ClosedXML — Ghi nhanh danh sách dữ liệu

        /// <summary>
        /// Xuất nhanh danh sách dữ liệu ra file Excel với header + data.
        /// Tự động định dạng bảng đẹp (header đậm, border, auto-fit).
        /// </summary>
        /// <param name="outputPath">Đường dẫn file xuất.</param>
        /// <param name="sheetName">Tên sheet.</param>
        /// <param name="headers">Danh sách tên cột.</param>
        /// <param name="rows">Danh sách các dòng dữ liệu (mỗi dòng là mảng object).</param>
        /// <param name="title">Tiêu đề (tuỳ chọn, hiển thị dòng 1).</param>
        public static void QuickExport(string outputPath, string sheetName, string[] headers,
            List<object[]> rows, string? title = null)
        {
            CreateWorkbook(outputPath, wb =>
            {
                string safeName = sheetName.Length > 31 ? sheetName.Substring(0, 31) : sheetName;
                var ws = wb.Worksheets.Add(safeName);

                int startRow = 1;

                // Title (nếu có)
                if (!string.IsNullOrEmpty(title))
                {
                    ws.Cell(1, 1).Value = title.ToUpper();
                    ws.Range(1, 1, 1, headers.Length).Merge().Style
                        .Font.SetFontSize(14)
                        .Font.SetBold()
                        .Font.SetFontColor(XLColor.DarkBlue)
                        .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    startRow = 3;
                }

                // Headers
                for (int c = 0; c < headers.Length; c++)
                {
                    var cell = ws.Cell(startRow, c + 1);
                    cell.Value = headers[c];
                    cell.Style
                        .Font.SetBold()
                        .Font.SetFontColor(XLColor.White)
                        .Fill.SetBackgroundColor(XLColor.FromHtml("#1F4E79"))
                        .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                        .Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                        .Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                }

                // Data rows
                int dataRow = startRow + 1;
                foreach (var row in rows)
                {
                    for (int c = 0; c < row.Length && c < headers.Length; c++)
                    {
                        var cell = ws.Cell(dataRow, c + 1);
                        if (row[c] is double d)
                            cell.Value = d;
                        else if (row[c] is int i)
                            cell.Value = i;
                        else if (row[c] is DateTime dt)
                            cell.Value = dt;
                        else
                            cell.Value = row[c]?.ToString() ?? "";

                        cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    }
                    dataRow++;
                }

                // Auto-fit
                ws.Columns().AdjustToContents();
                ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
                ws.PageSetup.FitToPages(1, 0);
            });
        }

        #endregion

        #region Tiện ích — Lấy danh sách tên sheet

        /// <summary>
        /// Lấy danh sách tên tất cả các sheet trong file Excel.
        /// </summary>
        public static List<string> GetSheetNames(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Không tìm thấy file Excel: {filePath}");

            using (var workbook = new XLWorkbook(filePath))
            {
                return workbook.Worksheets.Select(ws => ws.Name).ToList();
            }
        }

        #endregion

        #region Tiện ích — Mở file Excel trên máy

        /// <summary>
        /// Mở file bằng ứng dụng mặc định (thường là Excel desktop).
        /// </summary>
        public static void OpenFile(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Không tìm thấy file: {filePath}");

            Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
        }

        #endregion
    }
}
