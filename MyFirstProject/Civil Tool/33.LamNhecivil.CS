// 33.LamNhecivil.CS - Lệnh Làm nhẹ bản vẽ Civil 3D
using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.LamNheCivilCmd))]

namespace MyFirstProject
{
    public class LamNheCivilCmd
    {
        /// <summary>
        /// Lệnh CT_LamNheCivil giúp Purge toàn bộ rác (Block, Layer, Linetype, Style...),
        /// Xóa các Regapps (Application rác), Audit (sửa lỗi) và các thiết lập không cần thiết.
        /// </summary>
        [CommandMethod("CT_LamNhe", CommandFlags.Modal | CommandFlags.Session)]
        [CommandMethod("CT_LAMNHECIVIL")]
        public void LamNheCivil()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;

            try
            {
                ed.WriteMessage("\n╔══════════════════════════════════════════════════╗");
                ed.WriteMessage("\n║         ĐANG CHẠY LÀM NHẸ BẢN VẼ CIVIL 3D        ║");
                ed.WriteMessage("\n╚══════════════════════════════════════════════════╝");
                
                // Khóa Document (LockDocument) nếu chạy CommandFlags.Session để có thể SendStringToExecute (hoặc không cần nếu không tác động Editor)
                // Tuy nhiên, SendStringToExecute sẽ ném lệnh vào queue và thực thi một cách tự nhiên
                // Các lệnh này sẽ chạy theo thứ tự trong Command Line

                // 1. Audit - Sửa các lỗi cơ sở dữ liệu bản vẽ (nếu có lỗi file rác sẽ làm file nặng & chậm)
                doc.SendStringToExecute("._AUDIT\nY\n", true, false, false);

                // 2. Xóa RegApps (Registered Applications) - Đây là nguyên nhân hàng đầu gây đứng/lag và nặng file
                doc.SendStringToExecute("._-PURGE\nR\n*\nN\n", true, false, false);

                // 3. Purge All (Xóa tất cả Block, Layer, Linetype, DimStyle, TextStyle... không dùng)
                doc.SendStringToExecute("._-PURGE\nALL\n*\nN\n", true, false, false);

                // 4. Purge All lần 2 vì một số Block lồng nhau sau khi xóa lần 1 mới biến mất khóa (Lần 1 xóa Block cha, lần 2 mới xóa Block con)
                doc.SendStringToExecute("._-PURGE\nALL\n*\nN\n", true, false, false);
                
                // 5. Purge Zero-Length Geometry và Empty Text objects - Tính năng Purge Orphaned Data
                // AutoCad hỗ trợ việc xoá các đối tượng text trống, polyline độ dài 0
                // Bằng lẹnh -PURGE -> Z (Zero-length geometry)
                doc.SendStringToExecute("._-PURGE\nZ\n", true, false, false);
                doc.SendStringToExecute("._-PURGE\nE\n", true, false, false); // Empty text

                // 6. Reset (Xóa) danh sách tỷ lệ kích cỡ chú thích không dùng tới
                doc.SendStringToExecute("._-SCALELISTEDIT\nR\nY\nE\n", true, false, false);
                
                // 7. AECPURGESTYLES - (Riêng phần mềm Civil 3D/Architecture)
                // Giúp xóa toàn bộ Styles (Alignment style, Profile style, Label style...) chưa dùng.
                // Ở Civil 3D khi khởi tạo từ template trống, có thể chứa hàng trăm AEC styles làm file bị nặng ngầm
                doc.SendStringToExecute("._AECPURGESTYLES\n", true, false, false);

                // Lưu ý
                doc.SendStringToExecute("(print \"Hoàn tất làm nhẹ file. Mời bạn lưu file (Ctrl+S) lại để thấy dung lượng giảm đáng kể!\")(princ)\n", true, false, false);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[Lỗi] Không thể làm nhẹ file: {ex.Message}\n");
            }
        }
    }
}
