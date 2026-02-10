using System;
using System.Linq;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using MyFirstProject.Helpers;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.NXsoftMasterSuite))]

namespace MyFirstProject
{
    /// <summary>
    /// Bộ công cụ Master từ NXsoft được chuyển đổi sang C#
    /// </summary>
    public class NXsoftMasterSuite
    {
        #region Civil Design Tools

        [CommandMethod("NXFIXTDTN")]
        public static void NXFIXTDTN()
        {
            SmartCommand.Execute("Sửa đường tự nhiên", (pm) =>
            {
                // TODO: Logic sửa đường TN theo vị trí cọc hoặc khoảng cách đều
                throw new NotImplementedException("Tính năng đang được phát triển.");
            });
        }

        [CommandMethod("NXCCTN")]
        public static void NXCCTN()
        {
            SmartCommand.Execute("Điền chênh cao tim cọc", (pm) =>
            {
                // TODO: Logic điền chênh cao giữa TN và TK tại tim cọc trên trắc ngang
                throw new NotImplementedException("Tính năng đang được phát triển.");
            });
        }

        [CommandMethod("NXPIPE")]
        public static void NXPIPE()
        {
            SmartCommand.Execute("Thiết kế mạng lưới thoát nước", (pm) =>
            {
                // TODO: Logic thiết kế Pipe Network từ Polylines/Feature Lines
                throw new NotImplementedException("Tính năng đang được phát triển.");
            });
        }

        [CommandMethod("NXNTDADD")]
        public static void NXNTDADD()
        {
            SmartCommand.Execute("Nhập dữ liệu NTD vào tuyến", (pm) =>
            {
                // TODO: Logic nạp dữ liệu NTD vào Alignment hiện có
                throw new NotImplementedException("Tính năng đang được phát triển.");
            });
        }

        [CommandMethod("NXDCDCOC")]
        public static void NXDCDCOC()
        {
            SmartCommand.Execute("Điền cao độ thiết kế lên mặt bằng", (pm) =>
            {
                // TODO: Logic lấy cao độ từ Profile điền lên bình đồ tại vị trí cọc
                throw new NotImplementedException("Tính năng đang được phát triển.");
            });
        }

        #endregion

        #region Staking & Sample Lines

        [CommandMethod("NXrenameSL")]
        public static void NXrenameSL()
        {
            SmartCommand.Execute("Đổi tên cọc (Sample Lines)", (pm) =>
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                var ed = doc.Editor;

                var prompt = ed.GetSelection(new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "AECC_SAMPLE_LINE") }));
                if (prompt.Status != PromptStatus.OK) return;

                var prefixPrompt = ed.GetString("\nNhập Prefix mới cho cọc (ví dụ: Km): ");
                if (prefixPrompt.Status != PromptStatus.OK) return;

                using var tr = db.TransactionManager.StartTransaction();
                pm.SetLimit(prompt.Value.Count);

                int count = 0;
                foreach (SelectedObject obj in prompt.Value)
                {
                    pm.MeterProgress();
                    var sl = (SampleLine)tr.GetObject(obj.ObjectId, OpenMode.ForWrite);
                    
                    // Lấy Alignment để tính lý trình
                    var slg = (SampleLineGroup)tr.GetObject(sl.GroupId, OpenMode.ForRead);
                    var alignment = (Alignment)tr.GetObject(slg.ParentAlignmentId, OpenMode.ForRead);
                    
                    double station = sl.Station;
                    int km = (int)(station / 1000);
                    double m = station % 1000;
                    
                    string newName = $"{prefixPrompt.StringResult}{km}+{m:F2}";
                    sl.Name = newName;
                    count++;
                }
                tr.Commit();
                ed.WriteMessage($"\n✓ Đã đổi tên cho {count} cọc theo chuẩn NX.");
            });
        }

        [CommandMethod("NXDTCoc")]
        public static void NXDTCoc()
        {
            SmartCommand.Execute("Điền tên cọc", (pm) =>
            {
                // TODO: Logic điền nhãn tên cọc lên mặt bằng
                throw new NotImplementedException("Tính năng đang được phát triển.");
            });
        }

        [CommandMethod("NXCCTTD")]
        public static void NXCCTTD()
        {
            SmartCommand.Execute("Chèn cọc từ trắc dọc", (pm) =>
            {
                // TODO: Logic chèn Sample Line tại vị trí click trên Profile View
                throw new NotImplementedException("Tính năng đang được phát triển.");
            });
        }

        #endregion

        #region Core Utilities

        [CommandMethod("CWPL")]
        public static void CWPL()
        {
            SmartCommand.Execute("Chỉnh bề dày Polyline nhanh", (pm) =>
            {
                var ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                var db = HostApplicationServices.WorkingDatabase;

                var prompt = ed.GetDouble("\nNhập bề dày mới cho Polylines: ");
                if (prompt.Status != PromptStatus.OK) return;

                var sel = ed.GetSelection(new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "LWPOLYLINE") }));
                if (sel.Status != PromptStatus.OK) return;

                using var tr = db.TransactionManager.StartTransaction();
                pm.SetLimit(sel.Value.Count);
                foreach (SelectedObject obj in sel.Value)
                {
                    pm.MeterProgress();
                    var pl = (Polyline)tr.GetObject(obj.ObjectId, OpenMode.ForWrite);
                    pl.ConstantWidth = prompt.Value;
                }
                tr.Commit();
                ed.WriteMessage($"\n✓ Đã cập nhật bề dày cho {sel.Value.Count} đối tượng.");
            });
        }

        [CommandMethod("NXChangeLW")]
        public static void NXChangeLW()
        {
            SmartCommand.Execute("Chỉnh LineWeight cho Layer", (pm) =>
            {
                // TODO: Logic thay đổi độ dày nét vẽ của Layer theo đối tượng chọn
                throw new NotImplementedException("Tính năng đang được phát triển.");
            });
        }

        [CommandMethod("NXNoiText")]
        public static void NXNoiText()
        {
            SmartCommand.Execute("Gộp Text cao độ (Khảo sát)", (pm) =>
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                var db = doc.Database;

                var pso = new PromptSelectionOptions { MessageForAdding = "\nChọn vùng chứa các Text cao độ cần gộp (Phần nguyên và phần thập phân): " };
                var sel = ed.GetSelection(pso);
                if (sel.Status != PromptStatus.OK) return;

                using var tr = db.TransactionManager.StartTransaction();
                var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                // Gom nhóm các text theo vị trí (gần nhau)
                var allTexts = sel.Value.GetObjectIds()
                    .Select(id => tr.GetObject(id, OpenMode.ForRead) as DBText)
                    .Where(t => t != null)
                    .Select(t => t!)
                    .ToList();

                pm.SetLimit(allTexts.Count());
                int mergedCount = 0;

                // Cài đặt ngưỡng khoảng cách để coi là một cặp (ví dụ 2m)
                double threshold = 2.0;

                var processed = new HashSet<ObjectId>();

                foreach (var t1 in allTexts)
                {
                    if (processed.Contains(t1.ObjectId)) continue;
                    pm.MeterProgress();

                    // Tìm text gần nhất chưa xử lý
                    var t2 = allTexts
                        .Where(t => !processed.Contains(t.ObjectId) && t.ObjectId != t1.ObjectId)
                        .OrderBy(t => t.Position.DistanceTo(t1.Position))
                        .FirstOrDefault();

                    if (t2 != null && t1.Position.DistanceTo(t2.Position) < threshold)
                    {
                        // Giả định text cao hơn là phần nguyên, text thấp hơn là phần thập phân (theo survey chuẩn)
                        // Hoặc text bên trái là phần nguyên. Ở đây dùng Distance để gộp.
                        string combined = t1.TextString + "." + t2.TextString;
                        
                        t1.UpgradeOpen();
                        t1.TextString = combined;
                        
                        t2.UpgradeOpen();
                        t2.Erase();
                        
                        processed.Add(t1.ObjectId);
                        processed.Add(t2.ObjectId);
                        mergedCount++;
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n✓ Đã gộp thành công {mergedCount} cặp Text khảo sát.");
            });
        }

        #endregion
    }
}
