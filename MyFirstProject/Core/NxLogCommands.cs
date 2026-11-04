// Lệnh CT_LOG: mở thư mục log để gửi khi cần hỗ trợ
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.Core.NxLogCommands))]

namespace MyFirstProject.Core
{
    public class NxLogCommands
    {
        [CommandMethod("CT_LOG")]
        public static void OpenLogFolder()
        {
            var ed = AcadApp.DocumentManager.MdiActiveDocument?.Editor;
            ed?.WriteMessage($"\n[CIVIL TOOL {AppInfo.Version}] Thư mục log: {AppInfo.LogFolder}");
            try { System.Diagnostics.Process.Start("explorer.exe", $"\"{AppInfo.LogFolder}\""); }
            catch (System.Exception ex) { NxLog.Swallow(ex); }
        }
    }
}
