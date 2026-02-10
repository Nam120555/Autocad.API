// (C) Copyright 2015 by  
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.Runtime;
using Acad = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using ATable = Autodesk.AutoCAD.DatabaseServices.Table;

using Civil = Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using Autodesk.Civil.Runtime;
using Autodesk.Civil.Settings;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Civil.ApplicationServices;
using CivSurface = Autodesk.Civil.DatabaseServices.TinSurface;
using Section = Autodesk.Civil.DatabaseServices.Section;
using Autodesk.Civil;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Documents;
using System.Globalization;
using MyFirstProject.Extensions;
// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(MyFirstProject.Parcel))]

namespace MyFirstProject
{
    public class Parcel
    {

        [CommandMethod("CTPA_TaoParcel_CacLoaiNha")]
        public static void CTPATaoParcelCacLoaiNha()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();

                //start here
                ObjectIdCollection polylineIdColl = UserInput.GSelectionSetWithType("Chọn các polyline cần chuyển: \n", "LWPOLYLINE");

                // Tìm hoặc tạo Site "TestSite" (chỉ 1 lần trước vòng lặp)
                Site? site = null;
                foreach (ObjectId siteId in A.Cdoc.GetSiteIds())
                {
                    Site? siteO = tr.GetObject(siteId, OpenMode.ForRead) as Site;
                    if (siteO != null && siteO.Name == "TestSite")
                    {
                        site = siteO;
                        break;
                    }
                }
                // Nếu chưa có Site "TestSite" → tạo mới
                if (site == null)
                {
                    ObjectId newSiteId = Site.Create(A.Cdoc, "TestSite");
                    site = tr.GetObject(newSiteId, OpenMode.ForRead) as Site;
                }
                if (site == null) { A.Ed.WriteMessage("\nKhông thể tạo Site."); return; }

                dynamic acadsite = site.AcadObject;
                dynamic parcellines = acadsite.ParcelSegments;

                foreach (ObjectId item in polylineIdColl)
                {
                    Polyline? polyline = tr.GetObject(item, OpenMode.ForWrite) as Polyline;
                    if (polyline == null) continue;

                    A.Ed.WriteMessage(polyline.Area.ToString() + "\n");
                    polyline.Closed = true;
                    dynamic segment = parcellines.AddFromEntity(polyline.AcadObject, true);
                }

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }









    }
}