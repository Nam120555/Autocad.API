// (C) Copyright 2015 by  
//
using Autodesk.Aec.PropertyData.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using DBObject = Autodesk.AutoCAD.DatabaseServices.DBObject;
using Entity = Autodesk.AutoCAD.DatabaseServices.Entity;
using MyFirstProject.Extensions;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(MyFirstProject.PropertySets))]

namespace MyFirstProject
{
    public class PropertySets
    {
        [CommandMethod("AT_Solid_Set_PropertySet")]
        public static void SetSolidPropertySet()
        {
            var ui = new UserInput();
            var propertySetUtils = new PropertySetUtils();
            var selectedIds = UserInput.GSelectionSetWithType("\nChọn các đối tượng 3D Solid cần thiết lập Property Set: ", "3DSOLID");

            using var tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                foreach (ObjectId objectId in selectedIds)
                {
                    var solid = tr.GetObject(objectId, OpenMode.ForWrite) as Solid3d;
                    if (solid == null) continue;

                    propertySetUtils.SetupSolidWithCalculatedProperties(tr, solid);
                }

                tr.Commit();
                A.Ed.WriteMessage($"\nĐã xử lý {selectedIds.Count} đối tượng 3D Solid.");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nLỗi: {ex.Message}");
                tr.Abort();
            }
        }

        [CommandMethod("AT_Solid_Show_Info")]
        public static void ShowSolidInfo()
        {
            var propertySetUtils = new PropertySetUtils();
            var selectedId = UserInput.GSelectionAnObject("\nChọn một đối tượng 3D Solid để xem thông tin: ");

            using var tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                var solid = tr.GetObject(selectedId, OpenMode.ForRead) as Solid3d;
                if (solid == null)
                {
                    A.Ed.WriteMessage("\nĐối tượng được chọn không phải là 3D Solid.");
                    return;
                }

                propertySetUtils.ShowSolidPropertySetInfo(tr, solid);
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nLỗi: {ex.Message}");
                tr.Abort();
            }
        }
    }
}


