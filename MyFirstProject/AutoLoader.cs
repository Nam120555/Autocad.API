// ==============================================================================
// AUTO LOADER - Tự động load DLL khi AutoCAD khởi động
// ==============================================================================
using System;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.Runtime;
using MyFirstProject.Core;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: ExtensionApplication(typeof(MyFirstProject.CivilToolAutoLoader))]

namespace MyFirstProject
{
    /// <summary>
    /// Lớp này tự động chạy khi DLL được NETLOAD
    /// Sẽ hiển thị Taskbar, Ribbon menu tự động
    /// </summary>
    public class CivilToolAutoLoader : IExtensionApplication
    {

        public void Initialize()
        {
            // Load trước các DLL phụ thuộc ClosedXML (ExcelNumberFormat, ...)
            // Dùng Assembly.Load(byte[]) — cách duy nhất hoạt động trong AutoCAD .NET 8
            NxLog.Info($"===== Khởi động {AppInfo.ProductName} v{AppInfo.Version} =====");
            try { ExcelDependencyLoader.EnsureLoaded(); }
            catch (System.Exception ex) { NxLog.Error(ex, "Nạp thư viện Excel"); }

            // Đăng ký sự kiện khi document/editor sẵn sàng
            AcadApp.Idle += OnApplicationIdle;
        }

        private void OnApplicationIdle(object? sender, EventArgs e)
        {
            // Chỉ chạy 1 lần khi application sẵn sàng
            AcadApp.Idle -= OnApplicationIdle;

            try
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                // ── 1. Hiển thị thông báo ──
                doc.Editor.WriteMessage("\n╔══════════════════════════════════════════════════════════╗");
                doc.Editor.WriteMessage($"\n║  {AppInfo.ProductName} v{AppInfo.Version} đã được load thành công!".PadRight(59) + "║");
                doc.Editor.WriteMessage($"\n║  DESIGNED BY {AppInfo.AuthorName} / {AppInfo.AuthorRole}".PadRight(59) + "║");
                doc.Editor.WriteMessage($"\n║  SĐT: {AppInfo.AuthorPhone}".PadRight(59) + "║");
                doc.Editor.WriteMessage("\n║  Gõ 'CT' để mở Taskbar | 'show_menu' để tạo Ribbon       ║");
                doc.Editor.WriteMessage("\n║  Gõ 'CT_DanhSachLenh' để xem danh sách lệnh              ║");
                doc.Editor.WriteMessage("\n║  Gõ 'NX' để mở HỘP LỆNH (tìm & chạy lệnh nhanh)          ║");
                doc.Editor.WriteMessage("\n║  Gõ 'CT_LOG' để mở thư mục log khi cần hỗ trợ            ║");
                doc.Editor.WriteMessage("\n╚══════════════════════════════════════════════════════════╝\n");

                // ── 3. Tự động tạo Ribbon menu ──
                try
                {
                    MyFirstProject.Autocad.ShowMenu();
                    
                    // ── 4. Load Lệnh tắt (Aliases) ──
                    var config = MenuSettingsForm.LoadSavedConfig();
                    if (config != null && config.CommandAliases != null)
                    {
                        ApplyAliases(config.CommandAliases);
                    }
                }
                catch (System.Exception ex)
                {
                    NxLog.Error(ex, "Tạo Ribbon / nạp lệnh tắt");
                    doc.Editor.WriteMessage($"\n[CIVIL TOOL] Không tạo được Ribbon: {ex.Message}. Gõ CT_LOG để xem chi tiết.\n");
                }
            }
            catch (System.Exception ex) { NxLog.Error(ex, "Khởi động Civil Tool"); }
        }

        /// <summary>
        /// Tạo file Lisp cấu hình lệnh tắt và tải vào AutoCAD
        /// </summary>
        public static void ApplyAliases(System.Collections.Generic.Dictionary<string, string> aliases)
        {
            try
            {
                if (aliases == null || aliases.Count == 0) return;
                
                string folder = AppInfo.DataFolder;
                
                string lspPath = Path.Combine(folder, "aliases.lsp");
                
                using (var writer = new StreamWriter(lspPath, false, System.Text.Encoding.UTF8))
                {
                    foreach (var kvp in aliases)
                    {
                        if (string.IsNullOrWhiteSpace(kvp.Key) || string.IsNullOrWhiteSpace(kvp.Value)) continue;
                        writer.WriteLine($"(defun c:{kvp.Value} () (command \"{kvp.Key}\"))");
                    }
                }

                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    string escapedPath = lspPath.Replace("\\", "/");
                    doc.SendStringToExecute($"(load \"{escapedPath}\")\n", true, false, false);
                }
            }
            catch (System.Exception ex) { NxLog.Error(ex, "Áp dụng lệnh tắt"); }
        }



        public void Terminate()
        {
            // Cleanup khi AutoCAD đóng
        }
    }
}
