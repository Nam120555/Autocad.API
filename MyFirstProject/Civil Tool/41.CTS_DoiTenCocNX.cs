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

[assembly: CommandClass(typeof(MyFirstProject.CTS_DoiTenCocNX))]

namespace MyFirstProject
{
    public class CTS_DoiTenCocNX
    {
        [CommandMethod("CTS_DoiTenCoc")]
        public static void CTSDoiTenCocMethod()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            PromptSelectionResult prompt = ed.GetSelection(new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "AECC_SAMPLE_LINE") }));
            if (prompt.Status != PromptStatus.OK) return;

            PromptResult prefixPrompt = ed.GetString("\nNhập Prefix mới cho cọc (ví dụ: Km): ");
            if (prefixPrompt.Status != PromptStatus.OK) return;

            try
            {
                using Transaction tr = db.TransactionManager.StartTransaction();
                int count = 0;
                foreach (SelectedObject obj in prompt.Value)
                {
                    // LUÔN DÙNG ForWrite THEO CHUẨN ĐỂ CHỐNG CRASH
                    var sl = (SampleLine)tr.GetObject(obj.ObjectId, OpenMode.ForWrite);
                    
                    var slg = (SampleLineGroup)tr.GetObject(sl.GroupId, OpenMode.ForWrite);
                    var alignment = (Alignment)tr.GetObject(slg.ParentAlignmentId, OpenMode.ForWrite);
                    
                    double station = sl.Station;
                    int km = (int)(station / 1000);
                    double m = station % 1000;
                    
                    string newName = $"{prefixPrompt.StringResult}{km}+{m:F2}";
                    sl.Name = newName;
                    count++;
                }
                tr.Commit();
                ed.WriteMessage($"\n[CTS_DoiTenCocNX] ✓ Đã đổi tên cho {count} cọc theo chuẩn NXSOFT.");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                ed.WriteMessage("\nLỗi: " + e.Message);
            }
        }
    }
}
