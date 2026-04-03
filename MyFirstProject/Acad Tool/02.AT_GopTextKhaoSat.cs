using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Acad = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Civil = Autodesk.Civil.ApplicationServices;
using Autodesk.Civil;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using Autodesk.Civil.Runtime;
using Autodesk.Civil.Settings;
using ATable = Autodesk.AutoCAD.DatabaseServices.Table;
using CivSurface = Autodesk.Civil.DatabaseServices.TinSurface;
using Section = Autodesk.Civil.DatabaseServices.Section;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.Acad_Tool.AT_GopTextKhaoSat))]

namespace MyFirstProject.Acad_Tool
{
    public class AT_GopTextKhaoSat
    {
        [CommandMethod("AT_GopText")]
        public static void ATGopTextKhaoSatMethod()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            var pso = new PromptSelectionOptions { MessageForAdding = "\nQuét chọn vùng chứa Text (phần nguyên và thập phân): " };
            var sel = ed.GetSelection(pso, new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "TEXT") }));
            if (sel.Status != PromptStatus.OK) return;

            try
            {
                using Transaction tr = db.TransactionManager.StartTransaction();
                
                var allTexts = sel.Value.GetObjectIds()
                    .Select(id => tr.GetObject(id, OpenMode.ForWrite) as DBText)
                    .Where(t => t != null)
                    .Select(t => t!)
                    .ToList();

                int mergedCount = 0;
                double threshold = 2.0;
                var processed = new HashSet<ObjectId>();

                foreach (var t1 in allTexts)
                {
                    if (processed.Contains(t1.ObjectId)) continue;

                    var t2 = allTexts
                        .Where(t => !processed.Contains(t.ObjectId) && t.ObjectId != t1.ObjectId)
                        .OrderBy(t => t.Position.DistanceTo(t1.Position))
                        .FirstOrDefault();

                    if (t2 != null && t1.Position.DistanceTo(t2.Position) < threshold)
                    {
                        string combined = t1.TextString + "." + t2.TextString;
                        t1.TextString = combined;
                        t2.Erase();
                        
                        processed.Add(t1.ObjectId);
                        processed.Add(t2.ObjectId);
                        mergedCount++;
                    }
                }
                tr.Commit();
                ed.WriteMessage($"\n[AT_GopText] ✓ Đã gộp thành công {mergedCount} cặp Text khảo sát.");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                ed.WriteMessage("\nLõi: " + e.Message);
            }
        }
    }
}
