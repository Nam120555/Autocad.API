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
[assembly: CommandClass(typeof(Civil3DCsharp.Parcel))]

namespace Civil3DCsharp
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
                foreach (ObjectId item in polylineIdColl)
                {
                    Polyline? polyline = tr.GetObject(item, OpenMode.ForWrite) as Polyline;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    A.Ed.WriteMessage(polyline.Area.ToString() + "\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    polyline.Closed = true;
                    Site? site = null;
                    foreach (ObjectId siteId in A.Cdoc.GetSiteIds())
                    {
                        Site? siteO = tr.GetObject(siteId, OpenMode.ForWrite) as Site;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        if (siteO.Name == "TestSite")
                        {
                            site = siteO;
                        }
                        else site = (Site)Site.Create(A.Cdoc, "TestSite").GetObject(OpenMode.ForRead);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    }
                    //Site site = (Site)Site.Create(a.cdoc, "TestSite").GetObject(OpenMode.ForRead);
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    dynamic acadsite = site.AcadObject;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    dynamic parcellines = acadsite.ParcelSegments;
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