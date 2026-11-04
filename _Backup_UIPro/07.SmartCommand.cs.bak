using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using MyFirstProject.Extensions;

namespace MyFirstProject.Helpers
{
    /// <summary>
    /// Cung cấp các phương thức để thực thi lệnh một cách "thông minh" với Progress Bar và Xử lý lỗi
    /// </summary>
    public static class SmartCommand
    {
        /// <summary>
        /// Thực thi một hành động với thanh tiến trình và xử lý lỗi chuyên nghiệp
        /// </summary>
        public static void Execute(string commandName, Action<ProgressMeter> action)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var pm = new ProgressMeter();

            try
            {
                // Bắt đầu đo thời gian và log (giả lập)
                ed.WriteMessage($"\n[CIVIL TOOL] Bắt đầu lệnh: {commandName}...");
                
                // Thực thi logic chính
                action(pm);

                ed.WriteMessage($"\n[CIVIL TOOL] Hoàn thành lệnh: {commandName}.");
            }
            catch (System.Exception ex)
            {
                ShowErrorDialog(commandName, ex);
            }
            finally
            {
                pm.Stop();
            }
        }

        /// <summary>
        /// Hiển thị hộp thoại lỗi chuyên nghiệp theo phong cách AutoCAD
        /// </summary>
        private static void ShowErrorDialog(string commandName, System.Exception ex)
        {
            try
            {
                // Sử dụng ShowAlertDialog là ổn định nhất trên mọi phiên bản AutoCAD
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(
                    $"DỰ ÁN CIVIL TOOL\n\n" +
                    $"Lỗi tại lệnh: {commandName}\n" +
                    $"- Message: {ex.Message}\n\n" +
                    $"Vui lòng liên hệ hỗ trợ kỹ thuật nếu lỗi vẫn tiếp diễn."
                );

                // Ghi thêm log chi tiết ra Editor
                if (!string.IsNullOrEmpty(ex.StackTrace))
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage($"\n[ERROR DETAILS] {ex.StackTrace}");
                }
            }
            catch
            {
                // Fallback nếu có lỗi trong quá trình báo lỗi
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage($"\n[CRITICAL ERROR] {commandName}: {ex.Message}");
            }
        }
    }
}
