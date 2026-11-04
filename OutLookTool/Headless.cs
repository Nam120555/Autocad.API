// ==============================================================================
// Headless.cs — Chế độ chạy ngầm: OutLookTool.exe --export
// Quét toàn bộ file Outlook (.ost/.pst), xuất Excel thống kê, KHÔNG hiện giao diện.
// Dùng cho lịch tự động hằng ngày (Task Scheduler).
//
//   C:\OutLook\ThongKe_Email_MoiNhat.xlsx      ← luôn là bản mới nhất (mở file này để tra cứu)
//   C:\OutLook\LichSu\ThongKe_Email_yyyyMMdd.xlsx  ← lưu lại 30 ngày gần nhất
//   C:\OutLook\TuDong_log.txt                   ← nhật ký mỗi lần chạy
// ==============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;

namespace OutLookTool
{
    internal static class Headless
    {
        public static readonly string LatestFile = Path.Combine(MainForm.OutputFolder, "ThongKe_Email_MoiNhat.xlsx");
        private static readonly string HistoryFolder = Path.Combine(MainForm.OutputFolder, "LichSu");
        private static readonly string LogFile = Path.Combine(MainForm.OutputFolder, "TuDong_log.txt");
        private const int KeepDays = 30;

        public static int Run()
        {
            try
            {
                Directory.CreateDirectory(MainForm.OutputFolder);
                Directory.CreateDirectory(HistoryFolder);
                Log("===== Bắt đầu xuất tự động =====");

                var files = FindOutlookFiles();
                if (files.Count == 0 && !OutlookComReader.IsOutlookRunning()) { Log("Không tìm thấy file .ost/.pst nào."); return 2; }

                var all = new List<EmailRecord>();
                bool viaOutlook = false;
                if (OutlookComReader.IsOutlookRunning())
                {
                    Log("Outlook đang mở → đọc trực tiếp qua Outlook");
                    try { all.AddRange(OutlookComReader.ReadAll(Log)); viaOutlook = true; }
                    catch (Exception ex) { Log("  Đọc qua Outlook lỗi: " + ex.Message + " → thử đọc file"); }
                }
                if (!viaOutlook)
                foreach (var f in files)
                {
                    string name = Path.GetFileName(f);
                    List<EmailRecord> emails;
                    try { emails = MainForm.ReadOutlookFile(f, name); }
                    catch (Exception ex1)
                    {
                        Log($"  {name}: đọc trực tiếp lỗi ({ex1.Message}) → đọc qua bản sao");
                        try { emails = MainForm.ReadWithTempCopy(f, name); }
                        catch (Exception ex2) { Log($"  {name}: BỎ QUA – {ex2.Message}"); continue; }
                    }
                    Log($"  {name}: {emails.Count:N0} email");
                    all.AddRange(emails);
                }

                if (all.Count == 0)
                {
                    Log("Không đọc được email nào → GIỮ NGUYÊN file Excel cũ.");
                    return 3;
                }

                all = all.OrderByDescending(e => e.DateObj ?? DateTime.MinValue).ToList();
                for (int i = 0; i < all.Count; i++) all[i].STT = i + 1;

                // Cache để mở giao diện lần sau nhanh
                try { File.WriteAllText(MainForm.DefaultCacheFile, JsonSerializer.Serialize(all)); }
                catch (Exception ex) { Log("  Không ghi được cache: " + ex.Message); }

                // Xuất ra file lịch sử trước, rồi chép đè bản "Mới nhất"
                string dated = Path.Combine(HistoryFolder, $"ThongKe_Email_{DateTime.Now:yyyyMMdd}.xlsx");
                MainForm.ExportToExcelFile(all, dated);
                try
                {
                    File.Copy(dated, LatestFile, true);
                    Log($"Đã xuất {all.Count:N0} email → {LatestFile}");
                }
                catch (IOException)
                {
                    Log($"File {Path.GetFileName(LatestFile)} đang mở trong Excel → chỉ lưu bản {dated}");
                }

                CleanupHistory();
                Log("===== Hoàn tất =====");
                return 0;
            }
            catch (Exception ex)
            {
                Log("LỖI: " + ex);
                return 1;
            }
        }

        private static List<string> FindOutlookFiles()
        {
            var list = new List<string>();
            if (Directory.Exists(MainForm.OutlookDataDir))
                list.AddRange(Directory.GetFiles(MainForm.OutlookDataDir, "*.ost"));
            if (File.Exists(MainForm.ArchivePstPath)) list.Add(MainForm.ArchivePstPath);
            if (Directory.Exists(MainForm.OutputFolder))
                foreach (var pst in Directory.GetFiles(MainForm.OutputFolder, "*.pst"))
                    if (!list.Contains(pst, StringComparer.OrdinalIgnoreCase)) list.Add(pst);
            return list;
        }

        private static void CleanupHistory()
        {
            try
            {
                foreach (var f in Directory.GetFiles(HistoryFolder, "ThongKe_Email_*.xlsx")
                                           .Where(f => File.GetLastWriteTime(f) < DateTime.Now.AddDays(-KeepDays)))
                    File.Delete(f);
            }
            catch (Exception ex) { Log("  Dọn lịch sử lỗi: " + ex.Message); }
        }

        private static void Log(string msg)
        {
            try { File.AppendAllText(LogFile, $"{DateTime.Now:dd/MM/yyyy HH:mm:ss}  {msg}{Environment.NewLine}", Encoding.UTF8); }
            catch { /* không để log làm hỏng việc xuất */ }
        }
    }
}
