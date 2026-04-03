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

[assembly: CommandClass(typeof(MyFirstProject.CTP_DienCaoDoThietKe))]

namespace MyFirstProject
{
    public class CTP_DienCaoDoThietKe
    {
        [CommandMethod("CTP_DienCaoDoMB")]
        public static void CTPDienCaoDoThietKeMethod()
        {
            try
            {
                using Transaction tr = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction();
                var ed = Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage("\n[CTP_DienCaoDoMB] Xuất cao độ từ TD ra MB (Chờ mã gốc NXSOFT).");

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nLỗi: " + e.Message);
            }
        }
    }
}
