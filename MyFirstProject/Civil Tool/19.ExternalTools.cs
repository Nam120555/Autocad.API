using System;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.EditorInput;
using Civil3DCsharp.Helpers;

[assembly: CommandClass(typeof(MyFirstProject.ExternalTools))]

namespace MyFirstProject
{
    public class ExternalTools
    {
        /// <summary>
        /// Thực thi lệnh từ DLL ngoài một cách an toàn và chuyên nghiệp
        /// </summary>
        private static void SafeExternalExecute(string commandFriendlyName, string dllPath, string commandName, string assemblyMatchPart)
        {
            SmartCommand.Execute(commandFriendlyName, (pm) =>
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                if (!File.Exists(dllPath))
                {
                    throw new FileNotFoundException($"Không tìm thấy tệp tin kỹ thuật tại: {dllPath}");
                }

                pm.SetLimit(2);
                pm.MeterProgress();

                // Kiểm tra xem assembly đã được load chưa để tránh nạp chồng
                bool isLoaded = AppDomain.CurrentDomain.GetAssemblies()
                                    .Any(a => a.FullName != null && a.FullName.IndexOf(assemblyMatchPart, StringComparison.OrdinalIgnoreCase) >= 0);

                if (!isLoaded)
                {
                    doc.Editor.WriteMessage($"\n[NX POWER] Đang khởi tạo module: {commandFriendlyName}...");
                    string escapedPath = dllPath.Replace("\\", "/");
                    doc.SendStringToExecute($"(command \"NETLOAD\" \"{escapedPath}\") ", true, false, false);
                }

                pm.MeterProgress();
                
                // Gọi lệnh. SendStringToExecute sẽ xếp hàng sau NETLOAD nếu vừa nạp xong.
                doc.SendStringToExecute($"{commandName} ", true, false, false);
            });
        }

        [CommandMethod("CT_VTOADOHG")]
        public static void CTVtoadohg()
        {
            const string dllPath = @"D:\Template\lip cad hay\V_TOADOHG\V_TOADOHG.dll";
            const string commandName = "V_TOADOHG";
            const string assemblyPart = "ManholeCoor"; // Tên định danh trong DLL

            SafeExternalExecute("Tọa độ hố ga", dllPath, commandName, assemblyPart);
        }
    }
}
