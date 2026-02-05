// ==============================================================================
// AUTO LOADER - Tự động load DLL khi AutoCAD khởi động
// ==============================================================================
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: ExtensionApplication(typeof(MyFirstProject.CivilToolAutoLoader))]

namespace MyFirstProject
{
    /// <summary>
    /// Lớp này tự động chạy khi DLL được NETLOAD
    /// Sẽ hiển thị Taskbar và Ribbon menu tự động
    /// </summary>
    public class CivilToolAutoLoader : IExtensionApplication
    {
        public void Initialize()
        {
            // Đăng ký sự kiện khi document/editor sẵn sàng
            AcadApp.Idle += OnApplicationIdle;
        }

        private void OnApplicationIdle(object? sender, System.EventArgs e)
        {
            // Chỉ chạy 1 lần khi application sẵn sàng
            AcadApp.Idle -= OnApplicationIdle;

            try
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    // Hiển thị thông báo
                    doc.Editor.WriteMessage("\n╔══════════════════════════════════════════════════════════╗");
                    doc.Editor.WriteMessage("\n║           CIVIL TOOL đã được load thành công!            ║");
                    doc.Editor.WriteMessage("\n║  Gõ 'CT' để mở Taskbar | 'show_menu' để tạo Ribbon       ║");
                    doc.Editor.WriteMessage("\n║  Gõ 'CT_DanhSachLenh' để xem danh sách lệnh              ║");
                    doc.Editor.WriteMessage("\n╚══════════════════════════════════════════════════════════╝\n");

                    // Tự động tạo Ribbon menu
                    try
                    {
                        MyFirstProject.Autocad.ShowMenu();
                    }
                    catch { }
                }
            }
            catch { }
        }

        public void Terminate()
        {
            // Cleanup khi AutoCAD đóng
        }
    }
}
