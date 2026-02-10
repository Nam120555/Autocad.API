// UtilityCommands.cs - Các lệnh tiện ích Civil 3D
// Chuyển đổi từ LISP: X7, X37, dump, UpDateStyle...

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadEntity = Autodesk.AutoCAD.DatabaseServices.Entity;

[assembly: CommandClass(typeof(Civil3DCsharp.UtilityCommands))]

namespace Civil3DCsharp
{
    /// <summary>
    /// Các lệnh tiện ích cho Civil 3D
    /// </summary>
    public class UtilityCommands
    {
        // ══════════════════════════════════════════════════════════════
        // EXPORT CAD 2007 (từ LISP X7)
        // ══════════════════════════════════════════════════════════════
        [CommandMethod("CTU_ExportCAD2007")]
        public static void CTU_ExportCAD2007()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            
            doc.SendStringToExecute("AecExportToAutoCAD2007 ", true, false, false);
        }

        // ══════════════════════════════════════════════════════════════
        // EXPLODE AEC OBJECT (từ LISP X37)
        // ══════════════════════════════════════════════════════════════
        [CommandMethod("CTU_ExplodeAEC")]
        public static void CTU_ExplodeAEC()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            
            doc.SendStringToExecute("-AECOBJEXPLODE Yes Current Yes No Yes No Yes ", true, false, false);
        }

        // ══════════════════════════════════════════════════════════════
        // REBUILD ALL SURFACES (từ LISP REBAS)
        // ══════════════════════════════════════════════════════════════
        [CommandMethod("CTS_RebuildSurface")]
        public static void CTS_RebuildSurface()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            
            int count = 0;
            
            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                foreach (ObjectId surfaceId in civilDoc.GetSurfaceIds())
                {
                    var surface = tr.GetObject(surfaceId, OpenMode.ForWrite) as TinSurface;
                    if (surface != null)
                    {
                        surface.Rebuild();
                        count++;
                    }
                }
                tr.Commit();
            }
            
            ed.WriteMessage($"\n◎ Đã rebuild {count} surface(s).");
        }

        // ══════════════════════════════════════════════════════════════
        // STYLE AUTO UPDATE OFF (từ LISP UpDateStyleOff)
        // ══════════════════════════════════════════════════════════════
        [CommandMethod("CTU_StyleAutoOff")]
        public static void CTU_StyleAutoOff()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            
            doc.SendStringToExecute("-AECCREFTEMPLATEAUTOUPDATE 0 ", true, false, false);
            doc.Editor.WriteMessage("\n⊘ Đã TẮT auto update style từ template.");
        }

        // ══════════════════════════════════════════════════════════════
        // STYLE AUTO UPDATE ON (từ LISP UpDateStlyleOn)
        // ══════════════════════════════════════════════════════════════
        [CommandMethod("CTU_StyleAutoOn")]
        public static void CTU_StyleAutoOn()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            
            doc.SendStringToExecute("-AECCREFTEMPLATEAUTOUPDATE 1 ", true, false, false);
            doc.Editor.WriteMessage("\n◎ Đã BẬT auto update style từ template.");
        }

        // ══════════════════════════════════════════════════════════════
        // DUMP OBJECT INFO (từ LISP dump)
        // ══════════════════════════════════════════════════════════════
        [CommandMethod("CTU_DumpObject")]
        public static void CTU_DumpObject()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            
            var peo = new PromptEntityOptions("\n⊙ Chọn đối tượng để xem thông tin: ");
            var per = ed.GetEntity(peo);
            
            if (per.Status != PromptStatus.OK) return;
            
            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var ent = tr.GetObject(per.ObjectId, OpenMode.ForRead);
                
                ed.WriteMessage("\n═══════════════════════════════════════════════════════════════");
                ed.WriteMessage($"\n▸ Object Type: {ent.GetType().Name}");
                ed.WriteMessage($"\n▸ Object ID: {ent.ObjectId}");
                ed.WriteMessage($"\n▸ Handle: {ent.Handle}");
                ed.WriteMessage($"\n▸ Layer: {(ent as AcadEntity)?.Layer ?? "N/A"}");
                
                // DXF Data
                var dxfData = ent.GetType().Name;
                ed.WriteMessage($"\n▸ DXF Name: {ent.ObjectId.ObjectClass.DxfName}");
                ed.WriteMessage("\n═══════════════════════════════════════════════════════════════");
                
                tr.Commit();
            }
        }

        // ══════════════════════════════════════════════════════════════
        // REORDER POINTS (từ LISP reorderpoints)
        // ══════════════════════════════════════════════════════════════
        [CommandMethod("CTPo_ReorderPoints")]
        public static void CTPo_ReorderPoints()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            
            // Lấy số điểm bắt đầu
            var pio = new PromptIntegerOptions("\n▸ Nhập số bắt đầu: ");
            pio.DefaultValue = 1;
            pio.UseDefaultValue = true;
            var pir = ed.GetInteger(pio);
            
            if (pir.Status != PromptStatus.OK) return;
            
            int nextNumber = pir.Value;
            
            ed.WriteMessage($"\n◎ Chọn các CogoPoint theo thứ tự muốn đánh số (bắt đầu từ {nextNumber})...");
            
            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                while (true)
                {
                    var peo = new PromptEntityOptions($"\n⊙ Chọn point #{nextNumber} (Enter để kết thúc): ");
                    peo.SetRejectMessage("\n⊘ Đây không phải CogoPoint!");
                    peo.AddAllowedClass(typeof(CogoPoint), true);
                    peo.AllowNone = true;
                    
                    var per = ed.GetEntity(peo);
                    
                    if (per.Status == PromptStatus.None) break;
                    if (per.Status != PromptStatus.OK) continue;
                    
                    var point = tr.GetObject(per.ObjectId, OpenMode.ForWrite) as CogoPoint;
                    if (point != null)
                    {
                        point.PointNumber = (uint)nextNumber;
                        ed.WriteMessage($"\n  ✓ Đã đánh số point = {nextNumber}");
                        nextNumber++;
                    }
                }
                
                tr.Commit();
            }
            
            ed.WriteMessage($"\n◎ Hoàn thành đánh số {nextNumber - pir.Value} point(s).");
        }

        // ══════════════════════════════════════════════════════════════
        // ADD PARCEL SEGMENT LABELS (từ LISP UHAddMultipleParcelSegmentlabels)
        // ══════════════════════════════════════════════════════════════
        [CommandMethod("CTP_AddParcelLabels")]
        public static void CTP_AddParcelLabels()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            
            ed.WriteMessage("\n▸ Chọn các Parcel để thêm nhãn...");
            
            var pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\n⊙ Chọn Parcels: ";
            
            var filter = new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Start, "AECC_PARCEL")
            });
            
            var psr = ed.GetSelection(pso, filter);
            
            if (psr.Status != PromptStatus.OK) return;
            
            int count = psr.Value.Count;
            
            // Gọi lệnh ADDPARCELSEGMENTLABELS cho từng parcel
            foreach (SelectedObject so in psr.Value)
            {
                doc.SendStringToExecute($"ADDPARCELSEGMENTLABELS (handent \"{so.ObjectId.Handle}\") (list 0.0 0.0 0.0)   ", true, false, false);
            }
            
            ed.WriteMessage($"\n◎ Đang thêm nhãn cho {count} parcel(s)...");
        }
    }
}

