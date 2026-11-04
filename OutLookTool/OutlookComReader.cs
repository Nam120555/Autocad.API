// ==============================================================================
// OutlookComReader.cs — Đọc email TRỰC TIẾP từ Outlook classic đang mở (COM/MAPI)
//
// Khi Outlook đang chạy, file .ost bị khoá → không đọc file được.
// Lớp này hỏi thẳng Outlook qua bảng (Table API – nhanh hơn duyệt từng thư rất nhiều).
// ==============================================================================
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

namespace OutLookTool
{
    internal static class OutlookComReader
    {
        public static bool IsOutlookRunning() => Process.GetProcessesByName("OUTLOOK").Length > 0;

        private static readonly string[] Cols =
        {
            "Subject", "SenderName", "SenderEmailAddress", "ReceivedTime", "SentOn",
            "Importance", "To", "CC", "MessageClass", "urn:schemas:httpmail:hasattachment"
        };

        public static List<EmailRecord> ReadAll(Action<string> log)
        {
            var results = new List<EmailRecord>();
            var type = Type.GetTypeFromProgID("Outlook.Application")
                       ?? throw new InvalidOperationException("Không tìm thấy Outlook (classic) trên máy.");
            dynamic outlook = Activator.CreateInstance(type)!;
            dynamic ns = outlook.GetNamespace("MAPI");

            dynamic stores = ns.Stores;
            int storeCount = stores.Count;
            for (int si = 1; si <= storeCount; si++)
            {
                dynamic store = stores[si];
                string storeName = "";
                try { storeName = store.DisplayName ?? ""; } catch (Exception ex) { log("  Store: " + ex.Message); }
                int before = results.Count;
                try { Walk(store.GetRootFolder(), "", storeName, results, log); }
                catch (Exception ex) { log($"  {storeName}: BỎ QUA – {ex.Message}"); }
                log($"  {storeName} (Outlook trực tiếp): {results.Count - before:N0} email");
            }
            return results;
        }

        private static void Walk(dynamic folder, string parent, string source, List<EmailRecord> results, Action<string> log)
        {
            string name = "";
            try { name = folder.Name ?? ""; } catch { /* thư mục hệ thống không đọc được tên */ }
            string path = string.IsNullOrEmpty(parent) ? name : parent + "\\" + name;

            bool isMail = false;
            try { isMail = (int)folder.DefaultItemType == 0; } catch { /* bỏ qua */ }   // 0 = olMailItem

            if (isMail)
            {
                try { ReadFolderTable(folder, path, source, results); }
                catch (Exception ex) { log($"  {path}: {ex.Message}"); }
            }

            try
            {
                dynamic subs = folder.Folders;
                int n = subs.Count;
                for (int k = 1; k <= n; k++)
                    Walk(subs[k], path, source, results, log);
            }
            catch (Exception ex) { log($"  {path} (thư mục con): {ex.Message}"); }
        }

        private static void ReadFolderTable(dynamic folder, string path, string source, List<EmailRecord> results)
        {
            dynamic table = folder.GetTable("", 0);   // 0 = olUserItems
            table.Columns.RemoveAll();
            foreach (var c in Cols) table.Columns.Add(c);

            while (!(bool)table.EndOfTable)
            {
                dynamic row = table.GetNextRow();
                object[] v = (object[])row.GetValues();
                string msgClass = Str(v[8]);
                if (!msgClass.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase)) continue;

                DateTime? received = v[3] is DateTime r && r.Year > 1900 ? r : null;
                DateTime? sent = v[4] is DateTime s && s.Year > 1900 ? s : null;
                string senderEmail = Str(v[2]);
                if (senderEmail.StartsWith("/O=", StringComparison.OrdinalIgnoreCase)) senderEmail = "";   // địa chỉ nội bộ Exchange

                int imp = v[5] is int i ? i : 1;
                results.Add(Make(Str(v[1]), senderEmail, Str(v[0]), received ?? sent,
                                 v[9] is bool att && att,
                                 imp == 2 ? "Cao (High)" : imp == 0 ? "Thap" : "Binh thuong",
                                 Str(v[6]), Str(v[7]), path, source));
            }
            Marshal.FinalReleaseComObject(table);
        }

        private static string Str(object? o) => o?.ToString() ?? "";

        private static EmailRecord Make(string senderName, string senderEmail, string subject, DateTime? date,
                                        bool hasAttachment, string importance, string to, string cc,
                                        string folderPath, string source)
        {
            string dateStr = "", timeStr = "", weekStr = "";
            if (date.HasValue)
            {
                dateStr = date.Value.ToString("yyyy-MM-dd");
                timeStr = date.Value.ToString("HH:mm:ss");
                int week = CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(date.Value, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                weekStr = $"Tuan {week}";
            }
            return new EmailRecord
            {
                Ngay = dateStr, Gio = timeStr, Tuan = weekStr, DateObj = date,
                NguoiGui = senderName, EmailNguoiGui = senderEmail,
                DonVi = MainForm.GetUnitCategory(senderName, senderEmail),
                NguoiNhan_To = to, NguoiNhan_CC = cc,
                ChuDe = subject, HangMuc = MainForm.GetTopicCategory(subject),
                MucDo = importance, DinhKem = hasAttachment ? "Co" : "Khong",
                ThuMuc = folderPath, NguonFile = source
            };
        }
    }
}
