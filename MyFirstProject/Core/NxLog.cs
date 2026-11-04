// ==============================================================================
// NxLog.cs — Ghi log ra file %AppData%\CivilTool\logs\civiltool-yyyyMMdd.log
//
//   NxLog.Info("...")            : thông tin
//   NxLog.Error(ex, "ngữ cảnh")  : lỗi nghiêm trọng (luôn ghi)
//   NxLog.Swallow(ex)            : lỗi được bỏ qua có chủ đích (thay cho catch { } rỗng).
//                                  Mỗi vị trí trong code chỉ ghi LẦN ĐẦU để log không phình to
//                                  khi lỗi xảy ra trong vòng lặp hàng nghìn đối tượng.
//
// Lệnh CT_LOG mở thư mục log (để gửi khi cần hỗ trợ).
// Không bao giờ ném lỗi ra ngoài — log hỏng cũng không làm hỏng lệnh.
// ==============================================================================
using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;

namespace MyFirstProject.Core
{
    public static class NxLog
    {
        private static readonly object _lock = new();
        private static readonly ConcurrentDictionary<string, int> _swallowSites = new();
        private const int KeepDays = 14;
        private static bool _cleaned;

        public static string CurrentLogFile =>
            Path.Combine(AppInfo.LogFolder, $"civiltool-{DateTime.Now:yyyyMMdd}.log");

        public static void Info(string message) => Write("INFO ", message);

        public static void Warn(string message) => Write("WARN ", message);

        public static void Error(Exception ex, string context = "",
            [CallerMemberName] string member = "", [CallerFilePath] string file = "", [CallerLineNumber] int line = 0)
        {
            Write("ERROR", $"{context} @ {Path.GetFileName(file)}:{line} ({member})\n{ex}");
        }

        /// <summary>
        /// Thay cho <c>catch { }</c>: ghi lại lỗi bị bỏ qua (chỉ lần đầu tại mỗi vị trí), không ném lỗi.
        /// </summary>
        public static void Swallow(Exception ex,
            [CallerMemberName] string member = "", [CallerFilePath] string file = "", [CallerLineNumber] int line = 0)
        {
            try
            {
                string site = $"{Path.GetFileName(file)}:{line}";
                int count = _swallowSites.AddOrUpdate(site, 1, (_, c) => c + 1);
                if (count == 1)
                    Write("SKIP ", $"{site} ({member}) {ex.GetType().Name}: {ex.Message}");
            }
            catch { /* log không được phép làm hỏng lệnh */ }
        }

        private static void Write(string level, string message)
        {
            try
            {
                lock (_lock)
                {
                    CleanupOldLogs();
                    File.AppendAllText(CurrentLogFile,
                        $"{DateTime.Now:HH:mm:ss.fff} [{level}] {message}{Environment.NewLine}", Encoding.UTF8);
                }
            }
            catch { /* log không được phép làm hỏng lệnh */ }
        }

        private static void CleanupOldLogs()
        {
            if (_cleaned) return;
            _cleaned = true;
            try
            {
                var limit = DateTime.Now.AddDays(-KeepDays);
                foreach (var f in Directory.GetFiles(AppInfo.LogFolder, "civiltool-*.log")
                                           .Where(f => File.GetLastWriteTime(f) < limit))
                    File.Delete(f);
            }
            catch { /* bỏ qua */ }
        }
    }
}
