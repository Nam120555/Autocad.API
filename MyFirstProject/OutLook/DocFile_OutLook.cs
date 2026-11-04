using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Globalization;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using ClosedXML.Excel;
using XstReader;

namespace MyFirstProject.OutLook
{
    /// <summary>
    /// Đọc trực tiếp file Outlook .ost/.pst bằng XstReader.Api
    /// KHÔNG cần Outlook chạy - Đọc file trực tiếp, AN TOÀN, không mất dữ liệu gốc.
    /// Xuất ra file Excel .xlsx với VBA macro + nút bấm.
    /// </summary>
    public class DocFile_OutLook
    {
        // === Đường dẫn các file Outlook ===
        private static readonly string OutlookDataDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "Microsoft", "Outlook");

        private static readonly string[] DefaultOstFiles = new[]
        {
            "sgc.qltkds01@sgc.net.vn.ost",
            "namhn@sgc.net.vn.ost"
        };

        private const string ArchivePstPath = @"C:\OutLook\QLTKDS01-ARCHIVE.pst";
        private const string OutputFolder = @"C:\OutLook";
        private const string TempCopyFolder = @"C:\OutLook\_TempRead";

        // === Lớp chứa dữ liệu 1 email ===
        private class EmailRecord
        {
            public int STT { get; set; }
            public string Ngay { get; set; } = "";
            public string Gio { get; set; } = "";
            public string Tuan { get; set; } = "";
            public DateTime? DateObj { get; set; }
            public string NguoiGui { get; set; } = "";
            public string EmailNguoiGui { get; set; } = "";
            public string DonVi { get; set; } = "";
            public string NguoiNhan_To { get; set; } = "";
            public string NguoiNhan_CC { get; set; } = "";
            public string ChuDe { get; set; } = "";
            public string HangMuc { get; set; } = "";
            public string MucDo { get; set; } = "Binh thuong";
            public string DinhKem { get; set; } = "Khong";
            public double DungLuong_KB { get; set; }
            public string ThuMuc { get; set; } = "";
            public string NguonFile { get; set; } = "";
        }

        #region AutoCAD Commands

        /// <summary>
        /// Lệnh: DOC_OUTLOOK / READ_OUTLOOK_FILE
        /// Đọc TẤT CẢ file .ost và .pst → xuất 1 file Excel tổng hợp có VBA macro.
        /// </summary>
        [CommandMethod("DOC_OUTLOOK")]
        [CommandMethod("READ_OUTLOOK_FILE")]
        public static void DocTatCaOutlook()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;

            try
            {
                ed.WriteMessage("\n=======================================================");
                ed.WriteMessage("\n  [CIVIL TOOL] ĐỌC TRỰC TIẾP FILE OUTLOOK (.ost/.pst)");
                ed.WriteMessage("\n  Không cần Outlook chạy - Dữ liệu gốc AN TOÀN 100%");
                ed.WriteMessage("\n=======================================================\n");

                // Thu thập danh sách file cần đọc
                List<string> filesToRead = new List<string>();

                // Thêm các file .ost
                foreach (var ostName in DefaultOstFiles)
                {
                    string fullPath = Path.Combine(OutlookDataDir, ostName);
                    if (File.Exists(fullPath))
                    {
                        filesToRead.Add(fullPath);
                        ed.WriteMessage($"\n  [✓] Tìm thấy: {ostName} ({FormatFileSize(fullPath)})");
                    }
                    else
                    {
                        ed.WriteMessage($"\n  [✗] Không tìm thấy: {fullPath}");
                    }
                }

                // Thêm file archive .pst
                if (File.Exists(ArchivePstPath))
                {
                    filesToRead.Add(ArchivePstPath);
                    ed.WriteMessage($"\n  [✓] Tìm thấy: {Path.GetFileName(ArchivePstPath)} ({FormatFileSize(ArchivePstPath)})");
                }

                // Tìm thêm các file .pst khác trong C:\OutLook
                if (Directory.Exists(OutputFolder))
                {
                    foreach (var pst in Directory.GetFiles(OutputFolder, "*.pst"))
                    {
                        if (!filesToRead.Contains(pst, StringComparer.OrdinalIgnoreCase))
                        {
                            filesToRead.Add(pst);
                            ed.WriteMessage($"\n  [✓] Tìm thêm: {Path.GetFileName(pst)} ({FormatFileSize(pst)})");
                        }
                    }
                }

                if (filesToRead.Count == 0)
                {
                    ed.WriteMessage("\n[LỖI] Không tìm thấy file Outlook nào để đọc!");
                    return;
                }

                ed.WriteMessage($"\n\n  Tổng cộng {filesToRead.Count} file sẽ được đọc...\n");

                // Đọc tất cả email
                List<EmailRecord> allEmails = new List<EmailRecord>();
                int fileIndex = 0;

                foreach (var filePath in filesToRead)
                {
                    fileIndex++;
                    string fileName = Path.GetFileName(filePath);
                    ed.WriteMessage($"\n[{fileIndex}/{filesToRead.Count}] Đang đọc: {fileName}...");

                    try
                    {
                        var emails = ReadOutlookFile(filePath, fileName, ed);
                        allEmails.AddRange(emails);
                        ed.WriteMessage($" → {emails.Count:N0} email");
                    }
                    catch (IOException ioEx) when (ioEx.Message.Contains("being used") || ioEx.HResult == -2147024864)
                    {
                        ed.WriteMessage($"\n  [⚠] File đang bị lock (Outlook đang chạy). Thử copy tạm...");
                        try
                        {
                            var emails = ReadWithTempCopy(filePath, fileName, ed);
                            allEmails.AddRange(emails);
                            ed.WriteMessage($" → {emails.Count:N0} email");
                        }
                        catch (System.Exception copyEx)
                        {
                            ed.WriteMessage($"\n  [LỖI] Không thể đọc {fileName}: {copyEx.Message}");
                            ed.WriteMessage($"\n  → Hãy tắt Outlook rồi thử lại.");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n  [LỖI] Lỗi đọc {fileName}: {ex.Message}");
                        // Thử copy tạm
                        try
                        {
                            ed.WriteMessage($"\n  Thử copy tạm để đọc...");
                            var emails = ReadWithTempCopy(filePath, fileName, ed);
                            allEmails.AddRange(emails);
                            ed.WriteMessage($" → {emails.Count:N0} email");
                        }
                        catch (System.Exception retryEx)
                        {
                            ed.WriteMessage($"\n  [LỖI] Vẫn không đọc được: {retryEx.Message}");
                        }
                    }
                }

                if (allEmails.Count == 0)
                {
                    ed.WriteMessage("\n\n[CẢNH BÁO] Không đọc được email nào. Có thể file đang bị lock.");
                    ed.WriteMessage("\n→ Hãy tắt Outlook Desktop rồi chạy lại lệnh DOC_OUTLOOK.");
                    return;
                }

                // Sắp xếp theo ngày mới nhất
                allEmails = allEmails
                    .OrderByDescending(e => e.DateObj ?? DateTime.MinValue)
                    .ToList();

                for (int i = 0; i < allEmails.Count; i++)
                    allEmails[i].STT = i + 1;

                ed.WriteMessage($"\n\n  ══════════════════════════════════════");
                ed.WriteMessage($"\n  TỔNG CỘNG: {allEmails.Count:N0} email từ {filesToRead.Count} file");
                ed.WriteMessage($"\n  ══════════════════════════════════════\n");

                // Xuất Excel
                string outputPath = Path.Combine(OutputFolder,
                    $"ThongKe_Email_TatCa_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

                ed.WriteMessage("\n  Đang xuất file Excel...");
                ExportToExcel(allEmails, outputPath, ed);

                ed.WriteMessage($"\n\n  [✓] ĐÃ XUẤT THÀNH CÔNG!");
                ed.WriteMessage($"\n  File: {outputPath}");
                ed.WriteMessage($"\n  Tổng email: {allEmails.Count:N0}");

                // Tạo file VBA script riêng để nhúng macro
                string vbaScriptPath = Path.Combine(OutputFolder, "inject_vba_macro.ps1");
                CreateVbaInjectionScript(vbaScriptPath, outputPath);
                ed.WriteMessage($"\n  VBA Script: {vbaScriptPath}");

                // Hỏi mở file
                PromptKeywordOptions pko = new PromptKeywordOptions(
                    "\nBạn có muốn mở file Excel và nhúng VBA macro ngay không?");
                pko.Keywords.Add("Co");
                pko.Keywords.Add("MoExcel");
                pko.Keywords.Add("Khong");
                pko.Keywords.Default = "Co";
                pko.AllowNone = true;

                PromptResult pr = ed.GetKeywords(pko);
                if (pr.Status == PromptStatus.OK)
                {
                    if (pr.StringResult == "Co" || string.IsNullOrEmpty(pr.StringResult))
                    {
                        // Mở Excel + chạy VBA injection
                        Process.Start(new ProcessStartInfo(outputPath) { UseShellExecute = true });
                        ed.WriteMessage("\n  → Đã mở file Excel.");
                        ed.WriteMessage("\n  → Chạy PowerShell để nhúng VBA macro...");

                        // Chạy VBA injection script sau 3 giây (đợi Excel mở)
                        RunVbaInjection(vbaScriptPath, ed);
                    }
                    else if (pr.StringResult == "MoExcel")
                    {
                        Process.Start(new ProcessStartInfo(outputPath) { UseShellExecute = true });
                        ed.WriteMessage("\n  → Đã mở file Excel (không nhúng VBA).");
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[LỖI NGHIÊM TRỌNG] {ex.Message}");
                ed.WriteMessage($"\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Lệnh: DOC_PST / READ_PST_FILE
        /// Đọc riêng 1 file PST/OST do người dùng chỉ định.
        /// </summary>
        [CommandMethod("DOC_PST")]
        [CommandMethod("READ_PST_FILE")]
        public static void DocMotFilePst()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;

            try
            {
                PromptStringOptions pso = new PromptStringOptions(
                    "\nNhập đường dẫn file .ost hoặc .pst cần đọc:");
                pso.AllowSpaces = true;

                // Gợi ý mặc định
                string defaultFile = Path.Combine(OutlookDataDir, DefaultOstFiles[0]);
                if (File.Exists(defaultFile))
                {
                    pso.DefaultValue = defaultFile;
                    pso.UseDefaultValue = true;
                }

                PromptResult pr = ed.GetString(pso);
                if (pr.Status != PromptStatus.OK || string.IsNullOrWhiteSpace(pr.StringResult))
                    return;

                string filePath = pr.StringResult.Trim().Trim('"');

                if (!File.Exists(filePath))
                {
                    ed.WriteMessage($"\n[LỖI] File không tồn tại: {filePath}");
                    return;
                }

                string ext = Path.GetExtension(filePath).ToLower();
                if (ext != ".ost" && ext != ".pst")
                {
                    ed.WriteMessage("\n[LỖI] Chỉ hỗ trợ file .ost hoặc .pst");
                    return;
                }

                string fileName = Path.GetFileName(filePath);
                ed.WriteMessage($"\n  Đang đọc: {fileName} ({FormatFileSize(filePath)})...");

                List<EmailRecord> emails;
                try
                {
                    emails = ReadOutlookFile(filePath, fileName, ed);
                }
                catch
                {
                    ed.WriteMessage("\n  File bị lock, thử copy tạm...");
                    emails = ReadWithTempCopy(filePath, fileName, ed);
                }

                if (emails.Count == 0)
                {
                    ed.WriteMessage("\n[CẢNH BÁO] Không đọc được email nào từ file này.");
                    return;
                }

                emails = emails.OrderByDescending(e => e.DateObj ?? DateTime.MinValue).ToList();
                for (int i = 0; i < emails.Count; i++) emails[i].STT = i + 1;

                string baseName = Path.GetFileNameWithoutExtension(fileName);
                string outputPath = Path.Combine(OutputFolder,
                    $"ThongKe_{baseName}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

                ed.WriteMessage($"\n  Tổng: {emails.Count:N0} email. Đang xuất Excel...");
                ExportToExcel(emails, outputPath, ed);

                ed.WriteMessage($"\n\n  [✓] XUẤT THÀNH CÔNG: {outputPath}");

                // Tạo VBA injection script
                string vbaScriptPath = Path.Combine(OutputFolder, "inject_vba_macro.ps1");
                CreateVbaInjectionScript(vbaScriptPath, outputPath);

                PromptKeywordOptions pko = new PromptKeywordOptions("\nMở file Excel?");
                pko.Keywords.Add("Co");
                pko.Keywords.Add("Khong");
                pko.Keywords.Default = "Co";
                pko.AllowNone = true;

                PromptResult prOpen = ed.GetKeywords(pko);
                if (prOpen.Status == PromptStatus.OK &&
                    (prOpen.StringResult == "Co" || string.IsNullOrEmpty(prOpen.StringResult)))
                {
                    Process.Start(new ProcessStartInfo(outputPath) { UseShellExecute = true });
                    RunVbaInjection(vbaScriptPath, ed);
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[LỖI] {ex.Message}");
            }
        }

        #endregion

        #region Core: Đọc file Outlook trực tiếp

        /// <summary>
        /// Đọc trực tiếp file .ost/.pst bằng XstReader.Api
        /// </summary>
        private static List<EmailRecord> ReadOutlookFile(string filePath, string sourceLabel, Editor ed)
        {
            var results = new List<EmailRecord>();

            using (var xstFile = new XstFile(filePath))
            {
                ProcessFolder(xstFile.RootFolder, "", sourceLabel, results, ed);
            }

            return results;
        }

        /// <summary>
        /// Copy file ra thư mục tạm rồi đọc (khi file bị Outlook lock)
        /// </summary>
        private static List<EmailRecord> ReadWithTempCopy(string filePath, string sourceLabel, Editor ed)
        {
            if (!Directory.Exists(TempCopyFolder))
                Directory.CreateDirectory(TempCopyFolder);

            string tempPath = Path.Combine(TempCopyFolder,
                $"temp_{Path.GetFileName(filePath)}");

            ed.WriteMessage($"\n  Đang copy file tạm ({FormatFileSize(filePath)})...");

            // Copy với FileShare.ReadWrite để tránh lock
            using (var source = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var dest = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
            {
                source.CopyTo(dest);
            }

            ed.WriteMessage(" OK. Đang đọc...");

            try
            {
                var results = ReadOutlookFile(tempPath, sourceLabel, ed);
                return results;
            }
            finally
            {
                // Xóa file tạm
                try { File.Delete(tempPath); } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
            }
        }

        /// <summary>
        /// Duyệt đệ quy tất cả folder trong mailbox
        /// </summary>
        private static void ProcessFolder(XstFolder folder, string parentPath, string sourceLabel,
            List<EmailRecord> results, Editor ed)
        {
            string currentPath = string.IsNullOrEmpty(parentPath)
                ? folder.DisplayName ?? ""
                : $"{parentPath}\\{folder.DisplayName}";

            // Đọc tất cả message trong folder
            try
            {
                foreach (var msg in folder.Messages)
                {
                    try
                    {
                        var record = ExtractEmailRecord(msg, currentPath, sourceLabel);
                        if (record != null)
                            results.Add(record);
                    }
                    catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

                    // Giải phóng bộ nhớ cho file lớn
                    try { msg.ClearContents(); } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                }
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

            // Duyệt subfolder
            try
            {
                foreach (var subFolder in folder.Folders)
                {
                    ProcessFolder(subFolder, currentPath, sourceLabel, results, ed);
                }
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

            // Progress mỗi 5000 email
            if (results.Count > 0 && results.Count % 5000 == 0)
            {
                ed.WriteMessage($" [{results.Count:N0}]");
            }
        }

        /// <summary>
        /// Trích xuất thông tin từ 1 email message
        /// </summary>
        private static EmailRecord ExtractEmailRecord(XstMessage msg, string folderPath, string sourceLabel)
        {
            if (msg == null) return null;

            string senderName = "";
            string senderEmail = "";
            string subject = "";
            DateTime? receivedTime = null;
            DateTime? sentTime = null;
            bool hasAttachment = false;
            string importance = "Binh thuong";

            // Lấy thông tin người gửi từ Recipients.Sender
            try
            {
                if (msg.Recipients?.Sender != null)
                {
                    senderName = msg.Recipients.Sender.DisplayName ?? "";
                    senderEmail = msg.Recipients.Sender.Address ?? "";
                }
                // Fallback: dùng From property
                if (string.IsNullOrEmpty(senderName))
                    senderName = msg.From ?? "";
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

            try { subject = msg.Subject ?? ""; } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
            try { receivedTime = msg.ReceivedTime; } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
            try { sentTime = msg.SubmittedTime; } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
            try { hasAttachment = msg.HasAttachments; } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
            try
            {
                var imp = msg.Importance;
                if (imp.HasValue)
                {
                    importance = imp.Value.ToString() switch
                    {
                        "High" => "Cao (High)",
                        "Low" => "Thap",
                        _ => "Binh thuong"
                    };
                }
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

            // Người nhận To / CC - sử dụng Recipients.To và Recipients.Cc
            string toList = "";
            string ccList = "";
            try
            {
                if (msg.Recipients != null)
                {
                    var toRecips = msg.Recipients.To;
                    var ccRecips = msg.Recipients.Cc;
                    if (toRecips != null)
                        toList = string.Join("; ", toRecips.Select(r => r.DisplayName ?? r.Address ?? ""));
                    if (ccRecips != null)
                        ccList = string.Join("; ", ccRecips.Select(r => r.DisplayName ?? r.Address ?? ""));
                }
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

            // Fallback: dùng msg.To và msg.Cc (string properties)
            if (string.IsNullOrEmpty(toList))
                try { toList = msg.To ?? ""; } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
            if (string.IsNullOrEmpty(ccList))
                try { ccList = msg.Cc ?? ""; } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

            // Xác định ngày/giờ
            DateTime? dateObj = receivedTime ?? sentTime;
            string dateStr = "";
            string timeStr = "";
            string weekStr = "";

            if (dateObj.HasValue)
            {
                dateStr = dateObj.Value.ToString("yyyy-MM-dd");
                timeStr = dateObj.Value.ToString("HH:mm:ss");
                var culture = CultureInfo.InvariantCulture;
                int weekNum = culture.Calendar.GetWeekOfYear(dateObj.Value,
                    CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                weekStr = $"Tuan {weekNum}";
            }

            return new EmailRecord
            {
                Ngay = dateStr,
                Gio = timeStr,
                Tuan = weekStr,
                DateObj = dateObj,
                NguoiGui = senderName,
                EmailNguoiGui = senderEmail,
                DonVi = GetUnitCategory(senderName, senderEmail),
                NguoiNhan_To = toList,
                NguoiNhan_CC = ccList,
                ChuDe = subject,
                HangMuc = GetTopicCategory(subject),
                MucDo = importance,
                DinhKem = hasAttachment ? "Co" : "Khong",
                DungLuong_KB = 0, // XstReader không hỗ trợ đọc kích thước email trực tiếp
                ThuMuc = folderPath,
                NguonFile = sourceLabel
            };
        }

        #endregion

        #region Xuất Excel (ClosedXML)

        private static void ExportToExcel(List<EmailRecord> emails, string outputPath, Editor ed)
        {
            if (!Directory.Exists(Path.GetDirectoryName(outputPath)))
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            using (var wb = new XLWorkbook())
            {
                // === Sheet 1: Dashboard ===
                CreateDashboardSheet(wb, emails);

                // === Sheet 2: Chi tiết Email ===
                CreateDetailSheet(wb, emails);

                // === Sheet 3: Thống kê theo Ngày ===
                CreateDailyStatSheet(wb, emails);

                // === Sheet 4: Thống kê Người gửi ===
                CreateSenderStatSheet(wb, emails);

                // === Sheet 5: Thống kê Người nhận ===
                CreateRecipientStatSheet(wb, emails);

                // === Sheet 6: Ma trận Gửi x Ngày ===
                CreateMatrixSheet(wb, emails);

                wb.SaveAs(outputPath);
            }

            ed.WriteMessage($"\n  Đã tạo file: {outputPath}");
        }

        private static void CreateDashboardSheet(XLWorkbook wb, List<EmailRecord> emails)
        {
            var ws = wb.Worksheets.Add("Trang Chu - Dashboard");

            // Title
            ws.Range("A1:H1").Merge();
            ws.Cell("A1").Value = "BẢNG THỐNG KÊ EMAIL OUTLOOK - ĐỌC TRỰC TIẾP TỪ FILE";
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Font.FontSize = 16;
            ws.Cell("A1").Style.Font.FontColor = XLColor.White;
            ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.FromArgb(24, 43, 73);
            ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Row(1).Height = 35;

            // Sync info
            ws.Range("A2:H2").Merge();
            ws.Cell("A2").Value = $"Đồng bộ lúc: {DateTime.Now:dd/MM/yyyy HH:mm:ss} | Phương thức: Đọc trực tiếp file (XstReader) | Không cần Outlook";
            ws.Cell("A2").Style.Font.Italic = true;
            ws.Cell("A2").Style.Font.FontSize = 10;
            ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Hướng dẫn VBA
            ws.Range("A3:H3").Merge();
            ws.Cell("A3").Value = "Bấm các nút bên dưới để cập nhật dữ liệu email trực tiếp (sau khi nhúng VBA macro)";
            ws.Cell("A3").Style.Font.Italic = true;
            ws.Cell("A3").Style.Font.FontSize = 10;
            ws.Cell("A3").Style.Font.FontColor = XLColor.FromArgb(24, 43, 73);
            ws.Cell("A3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Row(4).Height = 45;  // Chỗ cho nút VBA

            // KPI Headers (Row 5)
            string[] kpiHeaders = { "TỔNG EMAIL", "", "SỐ NGƯỜI GỬI", "", "MỨC ĐỘ CAO", "CÓ ĐÍNH KÈM", "SỐ NGÀY", "SỐ NGUỒN FILE" };
            var kpiColors = new[]
            {
                XLColor.FromArgb(0, 112, 192), XLColor.FromArgb(0, 112, 192),
                XLColor.FromArgb(38, 128, 0), XLColor.FromArgb(38, 128, 0),
                XLColor.FromArgb(192, 0, 0), XLColor.FromArgb(112, 48, 160),
                XLColor.FromArgb(41, 86, 129), XLColor.FromArgb(41, 86, 129)
            };

            for (int c = 0; c < kpiHeaders.Length; c++)
            {
                if (string.IsNullOrEmpty(kpiHeaders[c])) continue;
                ws.Cell(5, c + 1).Value = kpiHeaders[c];
                ws.Cell(5, c + 1).Style.Font.Bold = true;
                ws.Cell(5, c + 1).Style.Font.FontSize = 10;
                ws.Cell(5, c + 1).Style.Font.FontColor = XLColor.White;
                ws.Cell(5, c + 1).Style.Fill.BackgroundColor = kpiColors[c];
                ws.Cell(5, c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            // KPI Values (Row 6)
            int uniqueSenders = emails.Select(e => e.NguoiGui).Where(s => !string.IsNullOrEmpty(s)).Distinct().Count();
            int highCount = emails.Count(e => e.MucDo.Contains("Cao"));
            int attachCount = emails.Count(e => e.DinhKem == "Co");
            int dayCount = emails.Select(e => e.Ngay).Where(d => !string.IsNullOrEmpty(d)).Distinct().Count();
            int sourceCount = emails.Select(e => e.NguonFile).Distinct().Count();

            ws.Cell("A6").Value = emails.Count;
            ws.Cell("C6").Value = uniqueSenders;
            ws.Cell("E6").Value = highCount;
            ws.Cell("F6").Value = attachCount;
            ws.Cell("G6").Value = dayCount;
            ws.Cell("H6").Value = sourceCount;

            for (int c = 1; c <= 8; c++)
            {
                ws.Cell(6, c).Style.Font.Bold = true;
                ws.Cell(6, c).Style.Font.FontSize = 18;
                ws.Cell(6, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(6, c).Style.Fill.BackgroundColor = XLColor.FromArgb(240, 244, 248);
            }

            // --- Thống kê Nguồn file (Row 8-9) ---
            ws.Range("A8:D8").Merge();
            ws.Cell("A8").Value = "THỐNG KÊ THEO NGUỒN FILE";
            ws.Cell("A8").Style.Font.Bold = true;
            ws.Cell("A8").Style.Font.FontSize = 12;
            ws.Cell("A8").Style.Font.FontColor = XLColor.White;
            ws.Cell("A8").Style.Fill.BackgroundColor = XLColor.FromArgb(41, 86, 129);
            ws.Range("A8:D8").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Cell("A9").Value = "STT"; ws.Cell("B9").Value = "Nguồn file";
            ws.Cell("C9").Value = "Số email"; ws.Cell("D9").Value = "Tỷ lệ (%)";
            for (int c = 1; c <= 4; c++)
            {
                ws.Cell(9, c).Style.Font.Bold = true;
                ws.Cell(9, c).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 112, 192);
                ws.Cell(9, c).Style.Font.FontColor = XLColor.White;
                ws.Cell(9, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int row = 10;
            int stt = 1;
            foreach (var sg in emails.GroupBy(e => e.NguonFile).OrderByDescending(g => g.Count()))
            {
                double pct = Math.Round((double)sg.Count() / emails.Count * 100, 1);
                ws.Cell(row, 1).Value = stt;
                ws.Cell(row, 2).Value = sg.Key;
                ws.Cell(row, 3).Value = sg.Count();
                ws.Cell(row, 4).Value = $"{pct} %";
                if (row % 2 == 0) ws.Range(row, 1, row, 4).Style.Fill.BackgroundColor = XLColor.FromArgb(235, 243, 250);
                row++; stt++;
            }
            ws.Range(9, 1, row - 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(9, 1, row - 1, 4).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // --- Thống kê Đơn vị (sau nguồn file, cách 1 dòng) ---
            row += 1;
            int unitHeaderRow = row;
            ws.Range(row, 1, row, 4).Merge();
            ws.Cell(row, 1).Value = "THỐNG KÊ THEO ĐƠN VỊ";
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 1).Style.Font.FontSize = 12;
            ws.Cell(row, 1).Style.Font.FontColor = XLColor.White;
            ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 112, 192);
            ws.Range(row, 1, row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            row++;
            ws.Cell(row, 1).Value = "STT"; ws.Cell(row, 2).Value = "Đơn vị";
            ws.Cell(row, 3).Value = "Số email"; ws.Cell(row, 4).Value = "Tỷ lệ (%)";
            for (int c = 1; c <= 4; c++)
            {
                ws.Cell(row, c).Style.Font.Bold = true;
                ws.Cell(row, c).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 112, 192);
                ws.Cell(row, c).Style.Font.FontColor = XLColor.White;
                ws.Cell(row, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            int unitHdrRow = row;

            row++; stt = 1;
            foreach (var ug in emails.GroupBy(e => e.DonVi).OrderByDescending(g => g.Count()))
            {
                double pct = Math.Round((double)ug.Count() / emails.Count * 100, 1);
                ws.Cell(row, 1).Value = stt;
                ws.Cell(row, 2).Value = ug.Key;
                ws.Cell(row, 3).Value = ug.Count();
                ws.Cell(row, 4).Value = $"{pct} %";
                if (row % 2 == 0) ws.Range(row, 1, row, 4).Style.Fill.BackgroundColor = XLColor.FromArgb(235, 243, 250);
                row++; stt++;
            }
            ws.Range(unitHdrRow, 1, row - 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(unitHdrRow, 1, row - 1, 4).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // --- Thống kê Hạng mục (cột E-H) ---
            ws.Range("E8:H8").Merge();
            ws.Cell("E8").Value = "THỐNG KÊ THEO HẠNG MỤC CÔNG VIỆC";
            ws.Cell("E8").Style.Font.Bold = true;
            ws.Cell("E8").Style.Font.FontSize = 12;
            ws.Cell("E8").Style.Font.FontColor = XLColor.White;
            ws.Cell("E8").Style.Fill.BackgroundColor = XLColor.FromArgb(38, 128, 0);
            ws.Range("E8:H8").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Cell("E9").Value = "STT"; ws.Cell("F9").Value = "Hạng mục";
            ws.Cell("G9").Value = "Số email"; ws.Cell("H9").Value = "Tỷ lệ (%)";
            for (int c = 5; c <= 8; c++)
            {
                ws.Cell(9, c).Style.Font.Bold = true;
                ws.Cell(9, c).Style.Fill.BackgroundColor = XLColor.FromArgb(38, 128, 0);
                ws.Cell(9, c).Style.Font.FontColor = XLColor.White;
                ws.Cell(9, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int tRow = 10; stt = 1;
            foreach (var tg in emails.GroupBy(e => e.HangMuc).OrderByDescending(g => g.Count()))
            {
                double pct = Math.Round((double)tg.Count() / emails.Count * 100, 1);
                ws.Cell(tRow, 5).Value = stt;
                ws.Cell(tRow, 6).Value = tg.Key;
                ws.Cell(tRow, 7).Value = tg.Count();
                ws.Cell(tRow, 8).Value = $"{pct} %";
                if (tRow % 2 == 0) ws.Range(tRow, 5, tRow, 8).Style.Fill.BackgroundColor = XLColor.FromArgb(235, 247, 235);
                tRow++; stt++;
            }
            ws.Range(9, 5, tRow - 1, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(9, 5, tRow - 1, 8).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Column widths
            ws.Column(1).Width = 8; ws.Column(2).Width = 42; ws.Column(3).Width = 12;
            ws.Column(4).Width = 12; ws.Column(5).Width = 8; ws.Column(6).Width = 42;
            ws.Column(7).Width = 12; ws.Column(8).Width = 15;

            ws.SheetView.FreezeRows(6);
        }

        private static void CreateDetailSheet(XLWorkbook wb, List<EmailRecord> emails)
        {
            var ws = wb.Worksheets.Add("Chi tiet Email");

            // Headers
            string[] headers = {
                "STT", "Ngày", "Giờ", "Tuần", "Người gửi", "Email người gửi",
                "Đơn vị", "Người nhận (To)", "Người nhận (CC)", "Chủ đề",
                "Hạng mục", "Mức độ", "Đính kèm", "Dung lượng (KB)",
                "Thư mục", "Nguồn file"
            };

            for (int c = 0; c < headers.Length; c++)
            {
                ws.Cell(1, c + 1).Value = headers[c];
                ws.Cell(1, c + 1).Style.Font.Bold = true;
                ws.Cell(1, c + 1).Style.Font.FontSize = 11;
                ws.Cell(1, c + 1).Style.Font.FontColor = XLColor.White;
                ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(24, 43, 73);
                ws.Cell(1, c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            // Data
            for (int i = 0; i < emails.Count; i++)
            {
                int row = i + 2;
                var em = emails[i];

                ws.Cell(row, 1).Value = em.STT;
                ws.Cell(row, 2).Value = em.Ngay;
                ws.Cell(row, 3).Value = em.Gio;
                ws.Cell(row, 4).Value = em.Tuan;
                ws.Cell(row, 5).Value = em.NguoiGui;
                ws.Cell(row, 6).Value = em.EmailNguoiGui;
                ws.Cell(row, 7).Value = em.DonVi;
                ws.Cell(row, 8).Value = em.NguoiNhan_To;
                ws.Cell(row, 9).Value = em.NguoiNhan_CC;
                ws.Cell(row, 10).Value = em.ChuDe;
                ws.Cell(row, 11).Value = em.HangMuc;
                ws.Cell(row, 12).Value = em.MucDo;
                ws.Cell(row, 13).Value = em.DinhKem;
                ws.Cell(row, 14).Value = em.DungLuong_KB;
                ws.Cell(row, 15).Value = em.ThuMuc;
                ws.Cell(row, 16).Value = em.NguonFile;

                // Zebra striping
                if (row % 2 == 0)
                    ws.Range(row, 1, row, 16).Style.Fill.BackgroundColor = XLColor.FromArgb(235, 243, 250);

                // Highlight mức độ cao
                if (em.MucDo.Contains("Cao"))
                {
                    ws.Cell(row, 12).Style.Font.FontColor = XLColor.Red;
                    ws.Cell(row, 12).Style.Font.Bold = true;
                }
            }

            // Borders
            int lastRow = emails.Count + 1;
            ws.Range(1, 1, lastRow, 16).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(1, 1, lastRow, 16).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Column widths
            ws.Column(1).Width = 8;   // STT
            ws.Column(2).Width = 12;  // Ngay
            ws.Column(3).Width = 10;  // Gio
            ws.Column(4).Width = 10;  // Tuan
            ws.Column(5).Width = 25;  // Nguoi gui
            ws.Column(6).Width = 30;  // Email
            ws.Column(7).Width = 30;  // Don vi
            ws.Column(8).Width = 30;  // To
            ws.Column(9).Width = 30;  // CC
            ws.Column(10).Width = 60; // Chu de
            ws.Column(11).Width = 30; // Hang muc
            ws.Column(12).Width = 15; // Muc do
            ws.Column(13).Width = 10; // Dinh kem
            ws.Column(14).Width = 14; // Dung luong
            ws.Column(15).Width = 40; // Thu muc
            ws.Column(16).Width = 30; // Nguon file

            ws.SheetView.FreezeRows(1);
            ws.SheetView.FreezeColumns(1);
        }

        private static void CreateDailyStatSheet(XLWorkbook wb, List<EmailRecord> emails)
        {
            var ws = wb.Worksheets.Add("Thong ke theo Ngay");

            string[] headers = { "STT", "Ngày", "Thứ", "Tuần", "Số email", "Số người gửi",
                                  "Số người nhận", "Mức độ Cao", "Có đính kèm" };

            for (int c = 0; c < headers.Length; c++)
            {
                ws.Cell(1, c + 1).Value = headers[c];
                ws.Cell(1, c + 1).Style.Font.Bold = true;
                ws.Cell(1, c + 1).Style.Font.FontColor = XLColor.White;
                ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 112, 192);
                ws.Cell(1, c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            var dateGroups = emails
                .Where(e => !string.IsNullOrEmpty(e.Ngay))
                .GroupBy(e => e.Ngay)
                .OrderByDescending(g => g.Key);

            int row = 2; int stt = 1;
            foreach (var dg in dateGroups)
            {
                DateTime dd;
                string dow = "";
                if (DateTime.TryParse(dg.Key, out dd))
                {
                    dow = dd.DayOfWeek switch
                    {
                        DayOfWeek.Monday => "Thu 2",
                        DayOfWeek.Tuesday => "Thu 3",
                        DayOfWeek.Wednesday => "Thu 4",
                        DayOfWeek.Thursday => "Thu 5",
                        DayOfWeek.Friday => "Thu 6",
                        DayOfWeek.Saturday => "Thu 7",
                        DayOfWeek.Sunday => "Chu Nhat",
                        _ => ""
                    };
                }

                string weekStr = dg.First().Tuan;
                int uniqueSenders = dg.Select(e => e.NguoiGui).Where(s => !string.IsNullOrEmpty(s)).Distinct().Count();

                var allRecipients = dg
                    .SelectMany(e => (e.NguoiNhan_To ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries))
                    .Select(r => r.Trim())
                    .Where(r => !string.IsNullOrEmpty(r))
                    .Distinct()
                    .Count();

                int highCount = dg.Count(e => e.MucDo.Contains("Cao"));
                int attachCount = dg.Count(e => e.DinhKem == "Co");

                ws.Cell(row, 1).Value = stt;
                ws.Cell(row, 2).Value = dg.Key;
                ws.Cell(row, 3).Value = dow;
                ws.Cell(row, 4).Value = weekStr;
                ws.Cell(row, 5).Value = dg.Count();
                ws.Cell(row, 6).Value = uniqueSenders;
                ws.Cell(row, 7).Value = allRecipients;
                ws.Cell(row, 8).Value = highCount;
                ws.Cell(row, 9).Value = attachCount;

                // Highlight weekend
                if (dow == "Thu 7" || dow == "Chu Nhat")
                    ws.Range(row, 1, row, 9).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 220, 220);
                else if (row % 2 == 0)
                    ws.Range(row, 1, row, 9).Style.Fill.BackgroundColor = XLColor.FromArgb(235, 247, 235);

                row++; stt++;
            }

            // Total row
            ws.Cell(row, 2).Value = "TỔNG CỘNG";
            ws.Cell(row, 5).Value = emails.Count;
            ws.Range(row, 1, row, 9).Style.Font.Bold = true;
            ws.Range(row, 1, row, 9).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 242, 204);

            ws.Range(1, 1, row, 9).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(1, 1, row, 9).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Column(1).Width = 6; ws.Column(2).Width = 14; ws.Column(3).Width = 10;
            ws.Column(4).Width = 10; ws.Column(5).Width = 12; ws.Column(6).Width = 14;
            ws.Column(7).Width = 16; ws.Column(8).Width = 14; ws.Column(9).Width = 14;

            ws.SheetView.FreezeRows(1);
        }

        private static void CreateSenderStatSheet(XLWorkbook wb, List<EmailRecord> emails)
        {
            var ws = wb.Worksheets.Add("Thong ke Nguoi gui");

            string[] headers = { "STT", "Người gửi", "Email", "Đơn vị", "Số email",
                                  "Ngày đầu", "Ngày cuối", "Số ngày gửi" };

            for (int c = 0; c < headers.Length; c++)
            {
                ws.Cell(1, c + 1).Value = headers[c];
                ws.Cell(1, c + 1).Style.Font.Bold = true;
                ws.Cell(1, c + 1).Style.Font.FontColor = XLColor.White;
                ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(192, 0, 0);
                ws.Cell(1, c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            var senderGroups = emails
                .Where(e => !string.IsNullOrEmpty(e.NguoiGui))
                .GroupBy(e => e.NguoiGui)
                .OrderByDescending(g => g.Count());

            int row = 2; int stt = 1;
            foreach (var sg in senderGroups)
            {
                var withDate = sg.Where(e => e.DateObj.HasValue).ToList();
                string firstDate = withDate.Any() ? withDate.Min(e => e.DateObj!.Value).ToString("yyyy-MM-dd") : "";
                string lastDate = withDate.Any() ? withDate.Max(e => e.DateObj!.Value).ToString("yyyy-MM-dd") : "";
                int dayCount = sg.Select(e => e.Ngay).Where(d => !string.IsNullOrEmpty(d)).Distinct().Count();
                var first = sg.First();

                ws.Cell(row, 1).Value = stt;
                ws.Cell(row, 2).Value = sg.Key;
                ws.Cell(row, 3).Value = first.EmailNguoiGui;
                ws.Cell(row, 4).Value = first.DonVi;
                ws.Cell(row, 5).Value = sg.Count();
                ws.Cell(row, 6).Value = firstDate;
                ws.Cell(row, 7).Value = lastDate;
                ws.Cell(row, 8).Value = dayCount;

                if (row % 2 == 0)
                    ws.Range(row, 1, row, 8).Style.Fill.BackgroundColor = XLColor.FromArgb(253, 237, 237);

                row++; stt++;
            }

            ws.Range(1, 1, row - 1, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(1, 1, row - 1, 8).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Column(1).Width = 6; ws.Column(2).Width = 32; ws.Column(3).Width = 35;
            ws.Column(4).Width = 35; ws.Column(5).Width = 12; ws.Column(6).Width = 14;
            ws.Column(7).Width = 14; ws.Column(8).Width = 14;

            ws.SheetView.FreezeRows(1);
        }

        private static void CreateRecipientStatSheet(XLWorkbook wb, List<EmailRecord> emails)
        {
            var ws = wb.Worksheets.Add("Thong ke Nguoi nhan");

            string[] headers = { "STT", "Người nhận", "Tổng email", "Nhận To", "Nhận CC", "Số ngày" };

            for (int c = 0; c < headers.Length; c++)
            {
                ws.Cell(1, c + 1).Value = headers[c];
                ws.Cell(1, c + 1).Style.Font.Bold = true;
                ws.Cell(1, c + 1).Style.Font.FontColor = XLColor.White;
                ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(112, 48, 160);
                ws.Cell(1, c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            // Build recipient map
            var rMap = new Dictionary<string, (int Total, int To, int Cc, HashSet<string> Days)>();
            foreach (var em in emails)
            {
                var toNames = (em.NguoiNhan_To ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries)
                    .Select(r => r.Trim()).Where(r => !string.IsNullOrEmpty(r));
                var ccNames = (em.NguoiNhan_CC ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries)
                    .Select(r => r.Trim()).Where(r => !string.IsNullOrEmpty(r));

                foreach (var name in toNames)
                {
                    if (!rMap.ContainsKey(name)) rMap[name] = (0, 0, 0, new HashSet<string>());
                    var v = rMap[name];
                    rMap[name] = (v.Total + 1, v.To + 1, v.Cc, v.Days);
                    if (!string.IsNullOrEmpty(em.Ngay)) rMap[name].Days.Add(em.Ngay);
                }
                foreach (var name in ccNames)
                {
                    if (!rMap.ContainsKey(name)) rMap[name] = (0, 0, 0, new HashSet<string>());
                    var v = rMap[name];
                    rMap[name] = (v.Total + 1, v.To, v.Cc + 1, v.Days);
                    if (!string.IsNullOrEmpty(em.Ngay)) rMap[name].Days.Add(em.Ngay);
                }
            }

            int row = 2; int stt = 1;
            foreach (var kv in rMap.OrderByDescending(x => x.Value.Total).Take(500))
            {
                ws.Cell(row, 1).Value = stt;
                ws.Cell(row, 2).Value = kv.Key;
                ws.Cell(row, 3).Value = kv.Value.Total;
                ws.Cell(row, 4).Value = kv.Value.To;
                ws.Cell(row, 5).Value = kv.Value.Cc;
                ws.Cell(row, 6).Value = kv.Value.Days.Count;

                if (row % 2 == 0)
                    ws.Range(row, 1, row, 6).Style.Fill.BackgroundColor = XLColor.FromArgb(245, 238, 248);

                row++; stt++;
            }

            ws.Range(1, 1, row - 1, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(1, 1, row - 1, 6).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Column(1).Width = 6; ws.Column(2).Width = 35; ws.Column(3).Width = 12;
            ws.Column(4).Width = 12; ws.Column(5).Width = 12; ws.Column(6).Width = 12;

            ws.SheetView.FreezeRows(1);
        }

        private static void CreateMatrixSheet(XLWorkbook wb, List<EmailRecord> emails)
        {
            var ws = wb.Worksheets.Add("Ma tran Gui x Ngay");

            var senderGroups = emails
                .Where(e => !string.IsNullOrEmpty(e.NguoiGui))
                .GroupBy(e => e.NguoiGui)
                .OrderByDescending(g => g.Count())
                .Take(30)
                .ToList();

            var uniqueDates = emails
                .Where(e => !string.IsNullOrEmpty(e.Ngay))
                .Select(e => e.Ngay)
                .Distinct()
                .OrderBy(d => d)
                .ToList();

            // Headers
            ws.Cell(1, 1).Value = "STT";
            ws.Cell(1, 2).Value = "Người gửi";
            ws.Cell(1, 3).Value = "Đơn vị";

            for (int d = 0; d < uniqueDates.Count; d++)
                ws.Cell(1, d + 4).Value = uniqueDates[d];

            int totalCol = uniqueDates.Count + 4;
            ws.Cell(1, totalCol).Value = "TỔNG";

            // Header style
            ws.Range(1, 1, 1, totalCol).Style.Font.Bold = true;
            ws.Range(1, 1, 1, totalCol).Style.Font.FontSize = 10;
            ws.Range(1, 1, 1, totalCol).Style.Font.FontColor = XLColor.White;
            ws.Range(1, 1, 1, totalCol).Style.Fill.BackgroundColor = XLColor.FromArgb(41, 86, 129);
            ws.Range(1, 1, 1, totalCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Data
            int row = 2; int stt = 1;
            foreach (var s in senderGroups)
            {
                ws.Cell(row, 1).Value = stt;
                ws.Cell(row, 2).Value = s.Key;
                ws.Cell(row, 3).Value = s.First().DonVi;

                int sumTotal = 0;
                for (int d = 0; d < uniqueDates.Count; d++)
                {
                    int cnt = s.Count(e => e.Ngay == uniqueDates[d]);
                    if (cnt > 0)
                    {
                        ws.Cell(row, d + 4).Value = cnt;
                        sumTotal += cnt;

                        // Heatmap colors
                        if (cnt >= 5)
                            ws.Cell(row, d + 4).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 120, 120);
                        else if (cnt >= 3)
                            ws.Cell(row, d + 4).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 200, 100);
                        else
                            ws.Cell(row, d + 4).Style.Fill.BackgroundColor = XLColor.FromArgb(160, 225, 120);
                    }
                }

                ws.Cell(row, totalCol).Value = sumTotal;
                row++; stt++;
            }

            // Total row
            ws.Cell(row, 2).Value = "TỔNG / NGÀY";
            ws.Cell(row, 2).Style.Font.Bold = true;
            for (int d = 0; d < uniqueDates.Count; d++)
            {
                ws.Cell(row, d + 4).Value = emails.Count(e => e.Ngay == uniqueDates[d]);
            }
            ws.Cell(row, totalCol).Value = emails.Count;
            ws.Range(row, 1, row, totalCol).Style.Font.Bold = true;
            ws.Range(row, 1, row, totalCol).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 242, 204);

            // Borders
            ws.Range(1, 1, row, totalCol).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(1, 1, row, totalCol).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Column(1).Width = 5; ws.Column(2).Width = 32; ws.Column(3).Width = 30;
            for (int d = 0; d < uniqueDates.Count; d++) ws.Column(d + 4).Width = 11;
            ws.Column(totalCol).Width = 8;

            ws.SheetView.FreezeRows(1);
            ws.SheetView.FreezeColumns(3);
        }

        #endregion

        #region VBA Injection Script

        /// <summary>
        /// Tạo file PowerShell để nhúng VBA macro + nút bấm vào file Excel
        /// </summary>
        private static void CreateVbaInjectionScript(string scriptPath, string excelPath)
        {
            string script = $@"# inject_vba_macro.ps1
# Nhung VBA Macro va nut bam vao file Excel
# Tu dong chuyen .xlsx thanh .xlsm

param(
    [string]$TargetExcel = ""{excelPath}""
)

Write-Host ""==============================================""
Write-Host ""  NHUNG VBA MACRO VAO FILE EXCEL""
Write-Host ""==============================================""
Write-Host """"

Start-Sleep -Seconds 3  # Cho Excel mo xong

try {{
    Write-Host ""Ket noi vao Excel dang mo...""
    $excel = $null
    for ($attempt = 1; $attempt -le 10; $attempt++) {{
        try {{
            $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject(""Excel.Application"")
            break
        }} catch {{
            Start-Sleep -Seconds 2
        }}
    }}

    if (-not $excel) {{
        Write-Host ""Khong tim thay Excel dang mo.""
        exit 1
    }}

    # Tim workbook
    $wb = $null
    foreach ($wbItem in $excel.Workbooks) {{
        try {{
            if ($wbItem.FullName -eq $TargetExcel -or $wbItem.Name -like ""*ThongKe_Email*"") {{
                $wb = $wbItem
                break
            }}
        }} catch {{}}
    }}

    if (-not $wb) {{
        Write-Host ""Khong tim thay workbook. Dang mo file...""
        $wb = $excel.Workbooks.Open($TargetExcel)
    }}

    Write-Host ""Da ket noi: $($wb.Name)""

    # Tim sheet Dashboard
    $wsDash = $null
    foreach ($ws in $wb.Worksheets) {{
        if ($ws.Name -like ""*Dashboard*"" -or $ws.Name -like ""*Trang*"") {{
            $wsDash = $ws
            break
        }}
    }}
    if (-not $wsDash) {{ $wsDash = $wb.Worksheets.Item(1) }}

    # === Xoa Shape cu ===
    foreach ($sName in @(""btnCapNhatNgay"", ""btnTuDong30Phut"", ""btnDungTuDong"", ""btnDocTrucTiep"")) {{
        try {{ $wsDash.Shapes.Item($sName).Delete() }} catch {{}}
    }}

    # === Tao 4 nut bam ===
    Write-Host ""Tao cac nut bam...""

    # Nut 1: DOC TRUC TIEP (Xanh duong dam)
    $btn0 = $wsDash.Shapes.AddShape(5, 5, 62, 185, 35)
    $btn0.Name = ""btnDocTrucTiep""
    $btn0.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(24, 43, 73))
    $btn0.Line.Visible = $false
    $btn0.TextFrame2.TextRange.Text = ""DOC TRUC TIEP FILE (.ost/.pst)""
    $btn0.TextFrame2.TextRange.Font.Size = 10
    $btn0.TextFrame2.TextRange.Font.Bold = $true
    $btn0.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)
    $btn0.TextFrame2.HorizontalAnchor = 2
    $btn0.TextFrame2.VerticalAnchor = 3
    $btn0.Shadow.Visible = $true
    $btn0.Shadow.OffsetX = 2; $btn0.Shadow.OffsetY = 2

    # Nut 2: CAP NHAT NGAY (Xanh duong)
    $btn1 = $wsDash.Shapes.AddShape(5, 200, 62, 160, 35)
    $btn1.Name = ""btnCapNhatNgay""
    $btn1.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(0, 112, 192))
    $btn1.Line.Visible = $false
    $btn1.TextFrame2.TextRange.Text = ""CAP NHAT NGAY (Outlook)""
    $btn1.TextFrame2.TextRange.Font.Size = 10
    $btn1.TextFrame2.TextRange.Font.Bold = $true
    $btn1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)
    $btn1.TextFrame2.HorizontalAnchor = 2
    $btn1.TextFrame2.VerticalAnchor = 3
    $btn1.Shadow.Visible = $true
    $btn1.Shadow.OffsetX = 2; $btn1.Shadow.OffsetY = 2

    # Nut 3: TU DONG 30 PHUT (Xanh la)
    $btn2 = $wsDash.Shapes.AddShape(5, 370, 62, 180, 35)
    $btn2.Name = ""btnTuDong30Phut""
    $btn2.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(38, 128, 0))
    $btn2.Line.Visible = $false
    $btn2.TextFrame2.TextRange.Text = ""TU DONG MOI 30 PHUT""
    $btn2.TextFrame2.TextRange.Font.Size = 10
    $btn2.TextFrame2.TextRange.Font.Bold = $true
    $btn2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)
    $btn2.TextFrame2.HorizontalAnchor = 2
    $btn2.TextFrame2.VerticalAnchor = 3
    $btn2.Shadow.Visible = $true
    $btn2.Shadow.OffsetX = 2; $btn2.Shadow.OffsetY = 2

    # Nut 4: DUNG TU DONG (Do)
    $btn3 = $wsDash.Shapes.AddShape(5, 560, 62, 145, 35)
    $btn3.Name = ""btnDungTuDong""
    $btn3.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(192, 0, 0))
    $btn3.Line.Visible = $false
    $btn3.TextFrame2.TextRange.Text = ""DUNG TU DONG""
    $btn3.TextFrame2.TextRange.Font.Size = 10
    $btn3.TextFrame2.TextRange.Font.Bold = $true
    $btn3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)
    $btn3.TextFrame2.HorizontalAnchor = 2
    $btn3.TextFrame2.VerticalAnchor = 3
    $btn3.Shadow.Visible = $true
    $btn3.Shadow.OffsetX = 2; $btn3.Shadow.OffsetY = 2

    Write-Host ""  [OK] Da tao 4 nut bam""

    # === Nhung VBA Module ===
    Write-Host ""Nhung VBA Macro...""

    try {{
        foreach ($comp in $wb.VBProject.VBComponents) {{
            if ($comp.Name -eq ""ModCapNhatEmail"") {{
                $wb.VBProject.VBComponents.Remove($comp)
                break
            }}
        }}
    }} catch {{}}

    $vbaCode = @""
Option Explicit

Private Declare PtrSafe Function SetTimer Lib ""user32"" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib ""user32"" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

Private TimerID As LongPtr
Private Const INTERVAL_MS As Long = 1800000  ' 30 phut = 1800000ms

' === MACRO 0: Doc truc tiep file Outlook (.ost/.pst) ===
Public Sub DocTrucTiepFile()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    On Error Resume Next
    ws.Range(""A2"").Value = ""Dang doc truc tiep file Outlook... Vui long cho...""
    Application.StatusBar = ""Dang doc truc tiep file .ost/.pst...""
    DoEvents
    On Error GoTo 0
    
    Dim scriptPath As String
    scriptPath = ""C:\OutLook\read_direct_outlook.ps1""
    
    If Dir(scriptPath) = """" Then
        MsgBox ""Khong tim thay file script:"" & vbCrLf & scriptPath & vbCrLf & vbCrLf & _
               ""Hay chay lenh DOC_OUTLOOK trong AutoCAD de tao lai."", vbExclamation, ""Loi""
        Exit Sub
    End If
    
    Dim cmd As String
    cmd = ""powershell.exe -ExecutionPolicy Bypass -File """""" & scriptPath & """"""""
    
    Dim wsh As Object
    Set wsh = CreateObject(""WScript.Shell"")
    wsh.Run cmd, 1, True
    Set wsh = Nothing
    
    Application.StatusBar = ""Doc truc tiep hoan tat luc "" & Format(Now, ""dd/MM/yyyy HH:mm:ss"")
    
    MsgBox ""Da doc truc tiep file Outlook hoan tat!"" & vbCrLf & _
           ""Ket qua da luu vao file Excel moi trong C:\OutLook"" & vbCrLf & _
           ""Thoi gian: "" & Format(Now, ""dd/MM/yyyy HH:mm:ss""), vbInformation, ""Thanh cong""
End Sub

' === MACRO 1: Cap nhat qua Outlook COM ===
Public Sub CapNhatEmail()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    On Error Resume Next
    ws.Range(""A2"").Value = ""Dang cap nhat... Vui long cho...""
    Application.StatusBar = ""Dang quet va cap nhat du lieu email tu Outlook...""
    DoEvents
    On Error GoTo 0
    
    Dim scriptPath As String
    scriptPath = ""C:\OutLook\update_data_live.ps1""
    
    If Dir(scriptPath) = """" Then
        MsgBox ""Khong tim thay file script:"" & vbCrLf & scriptPath & vbCrLf & vbCrLf & _
               ""Hay chay lenh XUATLINK_OUTLOOK trong AutoCAD de tao lai."", vbExclamation, ""Loi""
        Exit Sub
    End If
    
    Dim cmd As String
    cmd = ""powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """""" & scriptPath & """"""""
    
    Dim wsh As Object
    Set wsh = CreateObject(""WScript.Shell"")
    wsh.Run cmd, 0, True
    Set wsh = Nothing
    
    Application.StatusBar = ""Cap nhat hoan tat luc "" & Format(Now, ""dd/MM/yyyy HH:mm:ss"")
    
    MsgBox ""Da cap nhat xong du lieu email tu Outlook!"" & vbCrLf & _
           ""Tong so email: Xem Dashboard"" & vbCrLf & _
           ""Thoi gian: "" & Format(Now, ""dd/MM/yyyy HH:mm:ss""), vbInformation, ""Thanh cong""
End Sub

' === MACRO 2: Bat tu dong cap nhat moi 30 phut ===
Public Sub CapNhatTuDong()
    If TimerID <> 0 Then
        KillTimer 0, TimerID
        TimerID = 0
    End If
    
    TimerID = SetTimer(0, 0, INTERVAL_MS, AddressOf TimerCallback)
    
    If TimerID <> 0 Then
        Application.StatusBar = ""Da bat tu dong cap nhat moi 30 phut. Timer ID: "" & TimerID
        MsgBox ""Da bat che do TU DONG CAP NHAT moi 30 phut!"" & vbCrLf & vbCrLf & _
               ""Email se duoc quet va cap nhat lien tuc."" & vbCrLf & _
               ""Bam nut 'DUNG TU DONG' de tat."", vbInformation, ""Tu dong cap nhat""
    Else
        MsgBox ""Khong the bat timer tu dong."", vbExclamation, ""Loi""
    End If
End Sub

' === Timer Callback ===
Private Sub TimerCallback(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal nIDEvent As LongPtr, ByVal dwTime As Long)
    On Error Resume Next
    KillTimer 0, nIDEvent
    Application.StatusBar = ""Dang tu dong cap nhat email... "" & Format(Now, ""HH:mm:ss"")
    CapNhatEmail
    TimerID = SetTimer(0, 0, INTERVAL_MS, AddressOf TimerCallback)
    On Error GoTo 0
End Sub

' === MACRO 3: Dung tu dong cap nhat ===
Public Sub DungCapNhat()
    If TimerID <> 0 Then
        KillTimer 0, TimerID
        TimerID = 0
        Application.StatusBar = ""Da tat che do tu dong cap nhat.""
        MsgBox ""Da TAT che do tu dong cap nhat."", vbInformation, ""Dung tu dong""
    Else
        MsgBox ""Hien khong co timer tu dong nao dang chay."", vbInformation, ""Thong bao""
    End If
End Sub
""@

    $vbModule = $wb.VBProject.VBComponents.Add(1)
    $vbModule.Name = ""ModCapNhatEmail""
    $vbModule.CodeModule.AddFromString($vbaCode)
    Write-Host ""  [OK] Da nhung VBA Module: ModCapNhatEmail""

    # Gan macro vao nut
    $wsDash.Shapes.Item(""btnDocTrucTiep"").OnAction = ""ModCapNhatEmail.DocTrucTiepFile""
    $wsDash.Shapes.Item(""btnCapNhatNgay"").OnAction = ""ModCapNhatEmail.CapNhatEmail""
    $wsDash.Shapes.Item(""btnTuDong30Phut"").OnAction = ""ModCapNhatEmail.CapNhatTuDong""
    $wsDash.Shapes.Item(""btnDungTuDong"").OnAction = ""ModCapNhatEmail.DungCapNhat""
    Write-Host ""  [OK] Da gan macro cho 4 nut bam""

    # Luu file .xlsm
    $excel.DisplayAlerts = $false
    $currentPath = $wb.FullName
    if ($currentPath -like ""*.xlsx"") {{
        $newPath = $currentPath -replace '\.xlsx$', '.xlsm'
        $wb.SaveAs($newPath, 52)  # xlOpenXMLWorkbookMacroEnabled
        Write-Host ""  Da luu thanh .xlsm: $newPath""
    }} else {{
        $wb.Save()
        Write-Host ""  Da luu: $currentPath""
    }}
    $excel.DisplayAlerts = $true

    try {{ $wsDash.Select() | Out-Null }} catch {{}}

    Write-Host """"
    Write-Host ""==============================================""
    Write-Host ""  HOAN TAT! DA NHUNG VBA MACRO THANH CONG""
    Write-Host ""==============================================""
    Write-Host ""  [TIM DAM] DOC TRUC TIEP     - Doc file .ost/.pst khong can Outlook""
    Write-Host ""  [XANH]    CAP NHAT NGAY     - Cap nhat qua Outlook COM""
    Write-Host ""  [XANH LA] TU DONG 30 PHUT   - Tu dong cap nhat""
    Write-Host ""  [DO]      DUNG TU DONG      - Tat che do tu dong""
    Write-Host """"

}} catch {{
    Write-Host """"
    Write-Host ""[LOI] $($_.Exception.Message)""
    if ($_.Exception.Message -like ""*programmatic access*"" -or $_.Exception.Message -like ""*VBProject*"") {{
        Write-Host """"
        Write-Host ""BAN CAN BAT QUYEN TRUY CAP VBA:""
        Write-Host ""  1. Trong Excel: File > Options > Trust Center""
        Write-Host ""  2. Bam 'Trust Center Settings'""
        Write-Host ""  3. Chon 'Macro Settings'""
        Write-Host ""  4. Tick 'Trust access to the VBA project object model'""
        Write-Host ""  5. Bam OK, roi chay lai script nay.""
    }}
    Write-Host $_.ScriptStackTrace
}}
";
            File.WriteAllText(scriptPath, script, System.Text.Encoding.UTF8);
        }

        /// <summary>
        /// Chạy PowerShell script để nhúng VBA vào Excel
        /// </summary>
        private static void RunVbaInjection(string scriptPath, Editor ed)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-ExecutionPolicy Bypass -File \"{scriptPath}\"",
                    UseShellExecute = true,
                    CreateNoWindow = false,
                    WorkingDirectory = OutputFolder
                };

                Process.Start(psi);
                ed.WriteMessage("\n  → Script nhúng VBA đang chạy trong cửa sổ PowerShell...");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n  [CẢNH BÁO] Không thể chạy VBA injection: {ex.Message}");
                ed.WriteMessage($"\n  → Chạy thủ công: powershell -File \"{scriptPath}\"");
            }
        }

        #endregion

        #region Phân loại (chuyển từ PowerShell sang C#)

        /// <summary>
        /// Phân loại đơn vị dựa trên tên/email người gửi
        /// (tương đương Get-UnitCategory trong update_data_live.ps1)
        /// </summary>
        private static string GetUnitCategory(string senderName, string senderEmail)
        {
            string text = $"{senderName} {senderEmail}";

            if (Regex.IsMatch(text, @"SGC_KKH-KT|SGC_KKH|SGC_PKT|PKT", RegexOptions.IgnoreCase))
                return "SGC - Khoi Ke hoach Ky thuat";
            if (Regex.IsMatch(text, @"SGC_KĐSTĐC|SGC_KDSTDC|SGC_KĐSCT|SGC_DSCT", RegexOptions.IgnoreCase))
                return "SGC - Khoi Duong sat Toc do cao";
            if (Regex.IsMatch(text, @"SGC_KQLDA|SGC_QLDA", RegexOptions.IgnoreCase))
                return "SGC - Khoi Quan ly Du an";
            if (Regex.IsMatch(text, @"SGC_KQLXD|SGC_CTCT|BQLXD|Ban QLXD", RegexOptions.IgnoreCase))
                return "SGC - Ban QLXD & Cong trinh";
            if (Regex.IsMatch(text, @"SGC_KTCH|SGC_KTCXD|SGC_KTCCKN", RegexOptions.IgnoreCase))
                return "SGC - Khoi Than toc / Khoi Ky thuat";
            if (Regex.IsMatch(text, @"SGC_KQLKT|SGC_KHT|SGC_VP|SGC_Phòng", RegexOptions.IgnoreCase))
                return "SGC - Khoi Van phong & Hanh chinh";
            if (Regex.IsMatch(text, @"VH-VTK|VH-KCUDT|VH-KXD|VH-PTDA|Vinhomes|vinhomes\.vn", RegexOptions.IgnoreCase))
                return "Vinhomes - Vien Thiet ke & Phat trien DA";
            if (Regex.IsMatch(text, @"VinCons|vincons\.net", RegexOptions.IgnoreCase))
                return "Nha thau VinCons";
            if (Regex.IsMatch(text, @"SYSTRA|systra\.com", RegexOptions.IgnoreCase))
                return "Tu van SYSTRA";
            if (Regex.IsMatch(text, @"CONINCO|ckjvn\.vn|coninco", RegexOptions.IgnoreCase))
                return "Tu van Giam sat CONINCO";
            if (Regex.IsMatch(text, @"Trung Chính|trungchinhtcc", RegexOptions.IgnoreCase))
                return "Nha thau Trung Chinh";
            if (Regex.IsMatch(text, @"Đại Dũng|daidung\.vn", RegexOptions.IgnoreCase))
                return "Nha thau Dai Dung";
            if (Regex.IsMatch(text, @"Tecco1|tecco1\.vn", RegexOptions.IgnoreCase))
                return "Nha thau Tecco 1";
            if (Regex.IsMatch(text, @"Thiên Quốc An|TQA", RegexOptions.IgnoreCase))
                return "Nha thau Thien Quoc An";
            if (Regex.IsMatch(text, @"620idic", RegexOptions.IgnoreCase))
                return "Nha thau 620 IDIC";
            if (Regex.IsMatch(text, @"SGC", RegexOptions.IgnoreCase))
                return "SGC - Don vi khac";

            return "Doi tac / Don vi ngoai";
        }

        /// <summary>
        /// Phân loại hạng mục công việc dựa trên chủ đề email
        /// (tương đương Get-TopicCategory trong update_data_live.ps1)
        /// </summary>
        private static string GetTopicCategory(string subject)
        {
            if (string.IsNullOrEmpty(subject)) return "Chua xac dinh";

            if (Regex.IsMatch(subject, @"Bãi đúc|đầm|dầm|đốm|gói dầm", RegexOptions.IgnoreCase))
                return "Bai duc & Dam cau";
            if (Regex.IsMatch(subject, @"cọc|ép cọc|khoan nhồi|Bê tông|be tong|TRIAL MIX|C25|C30", RegexOptions.IgnoreCase))
                return "Thi cong Coc & Be tong";
            if (Regex.IsMatch(subject, @"Thiết kế|TK|Shop|SHOP|VMTK|MEP|Track|Design|bản vẽ|BIM", RegexOptions.IgnoreCase))
                return "Thiet ke & Shopdrawing";
            if (Regex.IsMatch(subject, @"cao độ|đường đỏ|ranh|rừng|mặt bằng|mốc|trắc địa", RegexOptions.IgnoreCase))
                return "Trac dia, Cao do & Mat bang";
            if (Regex.IsMatch(subject, @"Tiến độ|báo cáo|vướng mắc|ảnh hưởng|Đề xuất|de xuat", RegexOptions.IgnoreCase))
                return "Tien do & Vuong mac Cong truong";
            if (Regex.IsMatch(subject, @"Văn bản|Công văn|NNMT|phạt|pháp lý|Ký|phê duyệt|Quyết định", RegexOptions.IgnoreCase))
                return "Phap ly & Cong van Hanh chinh";
            if (Regex.IsMatch(subject, @"máy phát|điện|hàng rào|cầu rửa xe|thiết bị|vật tư|cáp", RegexOptions.IgnoreCase))
                return "Ha tang Bai duc & Thiet bi";

            return "Cong viec / Trao doi khac";
        }

        #endregion

        #region Helpers

        private static string FormatFileSize(string filePath)
        {
            try
            {
                long bytes = new FileInfo(filePath).Length;
                if (bytes >= 1073741824) return $"{bytes / 1073741824.0:F1} GB";
                if (bytes >= 1048576) return $"{bytes / 1048576.0:F1} MB";
                if (bytes >= 1024) return $"{bytes / 1024.0:F1} KB";
                return $"{bytes} B";
            }
            catch { return "?"; }
        }

        #endregion
    }
}
