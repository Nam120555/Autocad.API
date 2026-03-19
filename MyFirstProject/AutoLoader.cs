// ==============================================================================
// AUTO LOADER - Tự động load DLL khi AutoCAD khởi động
// ==============================================================================
using System;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: ExtensionApplication(typeof(MyFirstProject.CivilToolAutoLoader))]

namespace MyFirstProject
{
    /// <summary>
    /// Lớp này tự động chạy khi DLL được NETLOAD
    /// Sẽ hiển thị Taskbar, Ribbon menu, và load NXsoft tự động
    /// </summary>
    public class CivilToolAutoLoader : IExtensionApplication
    {
        // Đường dẫn NXsoft DLL — thay đổi nếu cần
        private const string NXsoftDllPath = @"C:\NXsoft\NXsoft_DotNet.dll";

        public void Initialize()
        {
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

                // ── 1. Load NXsoft tự động ──
                bool nxLoaded = LoadNXsoft(doc);

                // ── 2. Hiển thị thông báo ──
                doc.Editor.WriteMessage("\n╔══════════════════════════════════════════════════════════╗");
                doc.Editor.WriteMessage("\n║           CIVIL TOOL đã được load thành công!            ║");
                if (nxLoaded)
                    doc.Editor.WriteMessage("\n║           NXsoft đã được load tự động!                   ║");
                else
                    doc.Editor.WriteMessage("\n║           NXsoft: không tìm thấy hoặc đã load sẵn       ║");
                doc.Editor.WriteMessage("\n║  Gõ 'CT' để mở Taskbar | 'show_menu' để tạo Ribbon       ║");
                doc.Editor.WriteMessage("\n║  Gõ 'CT_DanhSachLenh' để xem danh sách lệnh              ║");
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
                catch { }
            }
            catch { }
        }

        /// <summary>
        /// Tạo file Lisp cấu hình lệnh tắt và tải vào AutoCAD
        /// </summary>
        public static void ApplyAliases(System.Collections.Generic.Dictionary<string, string> aliases)
        {
            try
            {
                if (aliases == null || aliases.Count == 0) return;
                
                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CivilTool");
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                
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
            catch { /* Skip errors */ }
        }

        /// <summary>
        /// Tự động load NXsoft_DotNet.dll nếu chưa load
        /// </summary>
        private bool LoadNXsoft(Autodesk.AutoCAD.ApplicationServices.Document doc)
        {
            try
            {
                // Kiểm tra file tồn tại
                if (!File.Exists(NXsoftDllPath))
                {
                    doc.Editor.WriteMessage($"\n[NXsoft] Không tìm thấy: {NXsoftDllPath}");
                    return false;
                }

                // Kiểm tra assembly đã load chưa
                bool alreadyLoaded = AppDomain.CurrentDomain.GetAssemblies()
                    .Any(a => a.FullName != null &&
                         a.FullName.IndexOf("NXsoft", StringComparison.OrdinalIgnoreCase) >= 0);

                if (alreadyLoaded)
                {
                    doc.Editor.WriteMessage("\n[NXsoft] Đã load sẵn.");
                    return false;
                }

                // NETLOAD NXsoft
                doc.Editor.WriteMessage("\n[NXsoft] Đang load NXsoft_DotNet.dll...");
                string escapedPath = NXsoftDllPath.Replace("\\", "/");
                doc.SendStringToExecute($"(command \"NETLOAD\" \"{escapedPath}\") ", true, false, false);

                return true;
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\n[NXsoft] Lỗi load: {ex.Message}");
                return false;
            }
        }

        public void Terminate()
        {
            // Cleanup khi AutoCAD đóng
        }
    }
}
