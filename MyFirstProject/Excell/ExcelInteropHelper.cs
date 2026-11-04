// ══════════════════════════════════════════════════════════════════════════════
// ExcelInteropHelper.cs — Đóng gói COM Interop để điều khiển Excel desktop
// Dùng khi cần: xuất PDF, chạy Macro VBA, mở file trực tiếp trong Excel
//
// ⚠️ YÊU CẦU: Máy phải cài đặt Microsoft Excel (Office 365, 2019, 2021...)
//
// DESIGNED BY HUYNH NGOC NAM / TRAFFIC ENGINEER - SĐT: 0378584154
// ══════════════════════════════════════════════════════════════════════════════

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace MyFirstProject.ExcelModule
{
    // ══════════════════════════════════════════════════════════════════════
    // P/Invoke — Thay thế Marshal.GetActiveObject (đã bị loại bỏ trong .NET 5+)
    // ══════════════════════════════════════════════════════════════════════
    internal static class ComHelper
    {
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(
            [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
            IntPtr pvReserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID(
            [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
            out Guid pclsid);

        /// <summary>
        /// Lấy COM object đang chạy từ ProgID (thay thế Marshal.GetActiveObject bị xóa trong .NET 8).
        /// </summary>
        public static object? GetActiveComObject(string progId)
        {
            int hr = CLSIDFromProgID(progId, out Guid clsid);
            if (hr < 0) return null;

            try
            {
                GetActiveObject(clsid, IntPtr.Zero, out object obj);
                return obj;
            }
            catch (COMException)
            {
                return null; // Không có instance đang chạy
            }
        }
    }

    /// <summary>
    /// Đóng gói COM Interop để điều khiển Excel Application trên máy.
    /// <para>
    /// <b>Tính năng:</b>
    /// <list type="bullet">
    ///   <item>Xuất PDF từ file Excel</item>
    ///   <item>Chạy Macro VBA trong file .xlsm</item>
    ///   <item>Mở file Excel trên desktop</item>
    ///   <item>Tự động dọn dẹp COM objects (tránh kẹt EXCEL.EXE process)</item>
    /// </list>
    /// </para>
    /// </summary>
    public static class ExcelInteropHelper
    {
        // ══════════════════════════════════════════════════════════════════════
        // KIỂM TRA EXCEL CÓ CÀI ĐẶT KHÔNG
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Kiểm tra máy có cài đặt Microsoft Excel không.
        /// </summary>
        public static bool IsExcelInstalled()
        {
            Type? excelType = Type.GetTypeFromProgID("Excel.Application");
            return excelType != null;
        }

        /// <summary>
        /// Kiểm tra và throw exception nếu không có Excel.
        /// </summary>
        private static Type GetExcelType()
        {
            Type? excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new InvalidOperationException(
                    "Không tìm thấy Microsoft Excel trên máy tính.\n" +
                    "Cần cài đặt Excel (Office 365 / 2019 / 2021) để sử dụng tính năng này.\n" +
                    "Hoặc dùng các method ClosedXML trong ExcelManager (không cần cài Excel).");
            return excelType;
        }

        // ══════════════════════════════════════════════════════════════════════
        // XUẤT PDF TỪ FILE EXCEL
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Xuất file Excel thành PDF bằng cách mở Excel COM → ExportAsFixedFormat.
        /// </summary>
        /// <param name="xlsxPath">Đường dẫn file .xlsx nguồn.</param>
        /// <param name="pdfPath">Đường dẫn file .pdf xuất ra.</param>
        /// <param name="openAfterExport">Mở file PDF sau khi xuất (mặc định = true).</param>
        /// <example>
        /// <code>
        /// ExcelInteropHelper.ExportToPdf(
        ///     @"C:\output\report.xlsx",
        ///     @"C:\output\report.pdf");
        /// </code>
        /// </example>
        public static void ExportToPdf(string xlsxPath, string pdfPath, bool openAfterExport = true)
        {
            if (!File.Exists(xlsxPath))
                throw new FileNotFoundException($"Không tìm thấy file Excel: {xlsxPath}");

            Type excelType = GetExcelType();
            dynamic? excelApp = null;
            dynamic? workbooks = null;
            dynamic? workbook = null;

            try
            {
                excelApp = Activator.CreateInstance(excelType)!;
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                workbooks = excelApp.Workbooks;
                workbook = workbooks.Open(Path.GetFullPath(xlsxPath));

                // xlTypePDF = 0
                workbook.ExportAsFixedFormat(0, Path.GetFullPath(pdfPath));

                workbook.Close(false);
                excelApp.Quit();

                if (openAfterExport && File.Exists(pdfPath))
                {
                    Process.Start(new ProcessStartInfo(pdfPath) { UseShellExecute = true });
                }
            }
            finally
            {
                // Dọn dẹp COM objects — quan trọng để tránh kẹt EXCEL.EXE
                SafeReleaseCom(workbook);
                SafeReleaseCom(workbooks);
                SafeReleaseCom(excelApp);

                // Đợi GC thu dọn
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // CHẠY MACRO VBA
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Mở file Excel (.xlsm) và chạy Macro VBA.
        /// </summary>
        /// <param name="xlsmPath">Đường dẫn file .xlsm chứa macro.</param>
        /// <param name="macroName">Tên macro (ví dụ: "Sheet1.MyMacro" hoặc "Module1.CapNhat").</param>
        /// <param name="args">Tham số truyền vào macro (tối đa 30).</param>
        /// <example>
        /// <code>
        /// ExcelInteropHelper.RunMacro(
        ///     @"C:\OutLook\ThongKe_Email.xlsm",
        ///     "Module1.CapNhatDuLieu",
        ///     "2024", "Q1");
        /// </code>
        /// </example>
        public static void RunMacro(string xlsmPath, string macroName, params object[] args)
        {
            if (!File.Exists(xlsmPath))
                throw new FileNotFoundException($"Không tìm thấy file Excel: {xlsmPath}");

            Type excelType = GetExcelType();
            dynamic? excelApp = null;
            dynamic? workbooks = null;
            dynamic? workbook = null;

            try
            {
                excelApp = Activator.CreateInstance(excelType)!;
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                workbooks = excelApp.Workbooks;
                workbook = workbooks.Open(Path.GetFullPath(xlsmPath));

                // Chạy macro với tham số
                string fullMacroName = $"'{Path.GetFileName(xlsmPath)}'!{macroName}";

                switch (args.Length)
                {
                    case 0:
                        excelApp.Run(fullMacroName);
                        break;
                    case 1:
                        excelApp.Run(fullMacroName, args[0]);
                        break;
                    case 2:
                        excelApp.Run(fullMacroName, args[0], args[1]);
                        break;
                    case 3:
                        excelApp.Run(fullMacroName, args[0], args[1], args[2]);
                        break;
                    case 4:
                        excelApp.Run(fullMacroName, args[0], args[1], args[2], args[3]);
                        break;
                    case 5:
                        excelApp.Run(fullMacroName, args[0], args[1], args[2], args[3], args[4]);
                        break;
                    default:
                        // Reflection fallback cho > 5 tham số
                        var runArgs = new object[args.Length + 1];
                        runArgs[0] = fullMacroName;
                        Array.Copy(args, 0, runArgs, 1, args.Length);
                        excelApp.GetType().InvokeMember("Run",
                            System.Reflection.BindingFlags.InvokeMethod, null, excelApp, runArgs);
                        break;
                }

                workbook.Save();
                workbook.Close(false);
                excelApp.Quit();
            }
            finally
            {
                SafeReleaseCom(workbook);
                SafeReleaseCom(workbooks);
                SafeReleaseCom(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // MỞ FILE EXCEL TRÊN DESKTOP
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Mở file Excel trực tiếp trên desktop (visible).
        /// Không giữ reference COM — người dùng tự đóng.
        /// </summary>
        /// <param name="filePath">Đường dẫn file Excel.</param>
        /// <param name="sheetName">Tên sheet cần activate (null = giữ nguyên).</param>
        public static void OpenInExcel(string filePath, string? sheetName = null)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Không tìm thấy file: {filePath}");

            if (string.IsNullOrEmpty(sheetName))
            {
                // Mở đơn giản bằng ShellExecute (nhanh, sạch)
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                return;
            }

            // Mở và navigate đến sheet cụ thể
            Type excelType = GetExcelType();
            dynamic? excelApp = null;
            dynamic? workbooks = null;
            dynamic? workbook = null;

            try
            {
                excelApp = Activator.CreateInstance(excelType)!;
                excelApp.Visible = true; // Hiển thị cho người dùng

                workbooks = excelApp.Workbooks;
                workbook = workbooks.Open(Path.GetFullPath(filePath));

                // Activate sheet theo tên
                try
                {
                    dynamic sheet = workbook.Sheets[sheetName];
                    sheet.Activate();
                    SafeReleaseCom(sheet);
                }
                catch
                {
                    // Sheet không tồn tại → giữ nguyên sheet đầu tiên
                }

                // KHÔNG đóng workbook/quit Excel — để người dùng làm việc
                // KHÔNG release excelApp — để Excel tiếp tục chạy
            }
            catch
            {
                // Nếu COM lỗi, fallback về ShellExecute
                SafeReleaseCom(workbook);
                SafeReleaseCom(workbooks);
                SafeReleaseCom(excelApp);
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // ĐỌC CELL TỪ EXCEL APPLICATION ĐANG MỞ
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Kết nối với Excel đang mở trên máy (nếu có) và đọc giá trị cell.
        /// Dùng để "live link" dữ liệu giữa AutoCAD và Excel đang mở.
        /// </summary>
        /// <param name="cellAddress">Địa chỉ cell (ví dụ: "B3").</param>
        /// <param name="sheetName">Tên sheet (null = sheet active).</param>
        /// <returns>Giá trị cell hoặc null nếu không kết nối được.</returns>
        /// <example>
        /// <code>
        /// // Đọc giá trị từ Excel đang mở trên máy
        /// object? val = ExcelInteropHelper.ReadFromRunningExcel("B3", "Sheet1");
        /// if (val != null) ed.WriteMessage($"\nGiá trị B3 = {val}");
        /// </code>
        /// </example>
        public static object? ReadFromRunningExcel(string cellAddress, string? sheetName = null)
        {
            dynamic? excelApp = null;

            try
            {
                // Kết nối với Excel instance đang chạy
                excelApp = ComHelper.GetActiveComObject("Excel.Application");
                if (excelApp == null) return null;

                dynamic workbook = excelApp.ActiveWorkbook;
                if (workbook == null) return null;

                dynamic sheet;
                if (!string.IsNullOrEmpty(sheetName))
                    sheet = workbook.Sheets[sheetName];
                else
                    sheet = excelApp.ActiveSheet;

                dynamic cell = sheet.Range[cellAddress];
                object? value = cell.Value;

                SafeReleaseCom(cell);
                SafeReleaseCom(sheet);
                SafeReleaseCom(workbook);
                // KHÔNG release excelApp — nó là instance đang chạy của người dùng

                return value;
            }
            catch
            {
                return null; // Không có Excel đang mở, hoặc lỗi khác
            }
        }

        /// <summary>
        /// Ghi giá trị vào cell của Excel đang mở trên máy (live link).
        /// </summary>
        /// <param name="cellAddress">Địa chỉ cell (ví dụ: "B3").</param>
        /// <param name="value">Giá trị cần ghi.</param>
        /// <param name="sheetName">Tên sheet (null = sheet active).</param>
        /// <returns>True nếu ghi thành công.</returns>
        public static bool WriteToRunningExcel(string cellAddress, object value, string? sheetName = null)
        {
            dynamic? excelApp = null;

            try
            {
                excelApp = ComHelper.GetActiveComObject("Excel.Application");
                if (excelApp == null) return false;

                dynamic workbook = excelApp.ActiveWorkbook;
                if (workbook == null) return false;

                dynamic sheet;
                if (!string.IsNullOrEmpty(sheetName))
                    sheet = workbook.Sheets[sheetName];
                else
                    sheet = excelApp.ActiveSheet;

                dynamic cell = sheet.Range[cellAddress];
                cell.Value = value;

                SafeReleaseCom(cell);
                SafeReleaseCom(sheet);
                SafeReleaseCom(workbook);

                return true;
            }
            catch
            {
                return false;
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // TIỆN ÍCH — KILL EXCEL PROCESSES TREO
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Đóng tất cả Excel process đang treo (không có cửa sổ).
        /// Dùng khi COM Interop bị lỗi để lại process zombie.
        /// </summary>
        /// <returns>Số process đã kill.</returns>
        public static int KillZombieExcelProcesses()
        {
            int killed = 0;
            var processes = Process.GetProcessesByName("EXCEL");

            foreach (var proc in processes)
            {
                try
                {
                    if (proc.MainWindowHandle == IntPtr.Zero)
                    {
                        // Process không có cửa sổ = zombie
                        proc.Kill();
                        proc.WaitForExit(3000);
                        killed++;
                    }
                }
                catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                finally
                {
                    proc.Dispose();
                }
            }

            return killed;
        }

        // ══════════════════════════════════════════════════════════════════════
        // HELPER — RELEASE COM OBJECTS
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// An toàn release COM object, tránh kẹt EXCEL.EXE process.
        /// </summary>
        private static void SafeReleaseCom(object? comObject)
        {
            if (comObject == null) return;
            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    Marshal.ReleaseComObject(comObject);
                }
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
        }
    }
}
