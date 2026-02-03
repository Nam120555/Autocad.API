using System;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;

[assembly: CommandClass(typeof(MyFirstProject.ExternalTools))]

namespace MyFirstProject
{
    public class ExternalTools
    {
        [CommandMethod("CT_VTOADOHG")]
        public static void CTVtoadohg()
        {
            const string dllPath = @"D:\Template\lip cad hay\V_TOADOHG\V_TOADOHG.dll";
            const string commandName = "V_TOADOHG";

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            if (!File.Exists(dllPath))
            {
                doc.Editor.WriteMessage($"\n[Lỗi] Không tìm thấy file DLL tại: {dllPath}");
                return;
            }

            try
            {
                // Kiểm tra xem assembly đã được load chưa (dựa trên tên ManholeCoor đã check)
                bool isLoaded = AppDomain.CurrentDomain.GetAssemblies()
                                    .Any(a => a.FullName.Contains("ManholeCoor"));

                if (!isLoaded)
                {
                    doc.Editor.WriteMessage("\nĐang nạp công cụ Tọa độ hố ga (V_TOADOHG)...");
                    // Sử dụng NETLOAD để AutoCAD đăng ký lệnh
                    string escapedPath = dllPath.Replace("\\", "/");
                    doc.SendStringToExecute($"(command \"NETLOAD\" \"{escapedPath}\") ", true, false, false);
                }

                // Gọi lệnh. Nếu vừa NETLOAD xong, SendStringToExecute sẽ xếp hàng lệnh này sau NETLOAD.
                doc.SendStringToExecute($"{commandName} ", true, false, false);
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\n[Lỗi] Không thể nạp hoặc chạy công cụ: {ex.Message}");
            }
        }
    }
}
