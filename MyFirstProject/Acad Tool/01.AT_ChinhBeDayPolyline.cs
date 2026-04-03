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

[assembly: CommandClass(typeof(MyFirstProject.Acad_Tool.AT_ChinhBeDayPolyline))]

namespace MyFirstProject.Acad_Tool
{
    public class AT_ChinhBeDayPolyline
    {
        [CommandMethod("AT_ChinhBeDayPLine")]
        public static void ATChinhBeDayPLineMethod()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            PromptDoubleResult prompt = ed.GetDouble("\nNhập bề dày liên tục cho Polyline (Ví dụ: 0.5): ");
            if (prompt.Status != PromptStatus.OK) return;

            PromptSelectionResult sel = ed.GetSelection(new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "LWPOLYLINE") }));
            if (sel.Status != PromptStatus.OK) return;

            try
            {
                using Transaction tr = db.TransactionManager.StartTransaction();
                foreach (SelectedObject obj in sel.Value)
                {
                    var pl = (Polyline)tr.GetObject(obj.ObjectId, OpenMode.ForWrite);
                    pl.ConstantWidth = prompt.Value;
                }
                tr.Commit();
                ed.WriteMessage($"\n[AT_ChinhBeDayPLine] ✓ Đã cập nhật bề dày cho {sel.Value.Count} Polyline.");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                ed.WriteMessage("\nLỗi: " + e.Message);
            }
        }
    }
}
