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
using System.Xml.Schema;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using MyFirstProject.Extensions;
//using Autodesk.Aec.DatabaseServices;
// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(MyFirstProject.Sampleline))]

namespace MyFirstProject
{
    public class Sampleline
    {
        


        [CommandMethod("CTS_DoiTenCoc3")]
        public static void CTSDoiTenCoc3()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput uI = new();

                //start here
                // choose an alignment
                ObjectId alignmentId = UserInput.GAlignmentId("\n Chọn tim đường " + "để đổi tên cọc: \n");

                //get alignment for read
                Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                //get first sampleline group
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectId samplelineGroup = alignment.GetSampleLineGroupIds()[0];
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                SampleLineGroup? sampleLineGroup = tr.GetObject(samplelineGroup, OpenMode.ForRead) as SampleLineGroup;

                //reset name sampleline
                int i = 0;
                int j = 1000;
                int value;
                String lyTrinh;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                foreach (ObjectId sampleLineId in sampleLineGroup.GetSampleLineIds())
                {
                    SampleLine? sampleline = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    if (int.TryParse(sampleline.Name, out value))
                    {
                        if ((sampleline.Station % 1000) < 100)
                        {
                            lyTrinh = "0" + (sampleline.Station % 1000).ToString();
                        }
                        else lyTrinh = (sampleline.Station % 1000).ToString();
                        sampleline.Name = "Km " + Math.Floor(sampleline.Station / 1000) + "+" + lyTrinh;
                        j++;
                    }
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    i++;
                }
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                //rename sampleline

                j = 1;
                foreach (ObjectId sampleLineId in sampleLineGroup.GetSampleLineIds())
                {
                    SampleLine? sampleline = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    if (int.TryParse(sampleline.Name, out value))
                    {
                        sampleline.Name = Convert.ToString(j);
                        j++;
                    }
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    i++;
                }
                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }


        [CommandMethod("CTS_DoiTenCoc2")]
        public static void CTSDoiTenCoc2()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput uI = new();

                //start here
                // choose an alignment
                ObjectId alignmentId = UserInput.GAlignmentId("\n Chọn tim đường " + "để đổi tên cọc: \n");

                //get alignment for read
                Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                //get first sampleline group
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectId samplelineGroup = alignment.GetSampleLineGroupIds()[0];
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                SampleLineGroup? sampleLineGroup = tr.GetObject(samplelineGroup, OpenMode.ForRead) as SampleLineGroup;

                int value;
                //rename sampleline
                ObjectId sampleLineDauId = UserInput.GSampleLineId("Chọn sampleLine điểm đầu đoạn tuyến cần đổi tên:");
                SampleLine? sampleLineDau = tr.GetObject(sampleLineDauId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                int sampleLineNunberDau = sampleLineDau.Number;
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                ObjectId sampleLineCuoiId = UserInput.GSampleLineId("Chọn sampleLine điểm đầu đoạn tuyến cần đổi tên:");
                SampleLine? sampleLineCuoi = tr.GetObject(sampleLineCuoiId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                int sampleLineNunberCuoi = sampleLineCuoi.Number;
#pragma warning restore CS8602 // Dereference of a possibly null reference.

#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                String tienTo = UserInput.GString("Nhập tiền tố cho đoạn cọc muốn đổi tên:");
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
                int soTT = UserInput.GInt("Nhập số thứ tự của cọc đầu của tên đoạn muốn đổi tên cọc:");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectIdCollection samplePlineIds = sampleLineGroup.GetSampleLineIds();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                int j = soTT;

                for (int k = sampleLineNunberDau; k < sampleLineNunberCuoi; k++)
                {
                    SampleLine? sampleline = tr.GetObject(samplePlineIds[k], OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    if (int.TryParse(sampleline.Name, out value))
                    {
                        sampleline.Name = tienTo + Convert.ToString(j);
                        j++;
                    }
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                }



                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("CTS_TaoBang_ToaDoCoc")]
        public static void CTSTaoBangToaDoCoc()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput uI = new();
                UtilitiesCAD uti = new();

                //start here

                // get an alignment
                ObjectId alignmentId = UserInput.GAlignmentId("\n Chọn tim đường " + "for export coordinate table");
                Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                //get the first samplelineGroup
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectId sampleLineGroup = alignment.GetSampleLineGroupIds()[0];
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                SampleLineGroup? samplelineGroup = tr.GetObject(sampleLineGroup, OpenMode.ForRead) as SampleLineGroup;

                //check number of sampleline in samplelinegroup
                int numberSampleline = 1;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                foreach (ObjectId sampleLineId in samplelineGroup.GetSampleLineIds())
                {
                    numberSampleline++;
                }
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                // get coordinate and name of point
                string[] samplelineName = new string[numberSampleline];
                string[] eastings = new string[numberSampleline];
                string[] northings = new string[numberSampleline];
                string[] stationsArr = new string[numberSampleline];
                int orderSampleline = 1;
                foreach (ObjectId sampleLineId in samplelineGroup.GetSampleLineIds())
                {
                    SampleLine? sampleline = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Double station = sampleline.Station;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    double easting = 0;
                    double northing = 0;
                    alignment.PointLocation(station, 0, ref easting, ref northing);
                    string lyTrinhStr = $"Km {Math.Floor(Math.Abs(station) / 1000)}+{(Math.Abs(station) % 1000):000.00}";
                    samplelineName[orderSampleline] = sampleline.Name.ToUpper();
                    stationsArr[orderSampleline] = lyTrinhStr;
                    eastings[orderSampleline] = Convert.ToString(Math.Round(easting, 3));
                    northings[orderSampleline] = Convert.ToString(Math.Round(northing, 3));
                    orderSampleline++;
                }

                //create a coordinate table
                UtilitiesCAD.CreateTableCoordinate4C(numberSampleline, 4, alignment.Name, samplelineName, stationsArr, northings, eastings);



                // draw polyline for check coordinnate
                UtilitiesCAD.CreateOpenPolyline(numberSampleline, eastings, northings);


                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("CTS_TaoBang_ToaDoCoc2")]
        public static void CTSTaoBangToaDoCoc2()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput uI = new();
                UtilitiesCAD uti = new();

                //start here

                // get an alignment
                ObjectId alignmentId = UserInput.GAlignmentId("\n Chọn tim đường " + "for export coordinate table");
                Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                //get the first samplelineGroup
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectId sampleLineGroup = alignment.GetSampleLineGroupIds()[0];
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                SampleLineGroup? samplelineGroup = tr.GetObject(sampleLineGroup, OpenMode.ForRead) as SampleLineGroup;

                //check number of sampleline in samplelinegroup
                int numberSampleline = 1;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                foreach (ObjectId sampleLineId in samplelineGroup.GetSampleLineIds())
                {
                    numberSampleline++;
                }
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                // get coordinate and name of point
                string[] samplelineName = new string[numberSampleline];
                string[] eastings = new string[numberSampleline];
                string[] northings = new string[numberSampleline];
                string[] stations = new string[numberSampleline];
                int orderSampleline = 1;
                foreach (ObjectId sampleLineId in samplelineGroup.GetSampleLineIds())
                {
                    SampleLine? sampleline = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Double station = sampleline.Station;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    double easting = 0;
                    double northing = 0;
                    alignment.PointLocation(station, 0, ref easting, ref northing);
                    samplelineName[orderSampleline] = sampleline.Name.ToUpper();
                    eastings[orderSampleline] = Convert.ToString(Math.Round(easting, 3));
                    northings[orderSampleline] = Convert.ToString(Math.Round(northing, 3));
                    stations[orderSampleline] = Convert.ToString(Math.Round(station, 3));
                    orderSampleline++;
                }

                //create a coordinate table
                UtilitiesCAD.CreateTableCoordinate2(numberSampleline, 4, alignment.Name, samplelineName, eastings, northings, stations);

                // draw polyline for check coordinnate
                UtilitiesCAD.CreateOpenPolyline(numberSampleline, eastings, northings);


                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("CTS_TaoBang_ToaDoCoc3")]
        public static void CTSTaoBangToaDoCoc3()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput uI = new();
                UtilitiesCAD uti = new();

                //start here

                // get an alignment
                ObjectId surfaceId = UserInput.GObjId("Chọn mặt phẳng cần lấy cao độ:");
                CivSurface? civSurface = tr.GetObject(surfaceId, OpenMode.ForRead) as CivSurface;
                ObjectId alignmentId = UserInput.GAlignmentId("\n Chọn tim đường " + "for export coordinate table");
                Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                //get the first samplelineGroup
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectId sampleLineGroup = alignment.GetSampleLineGroupIds()[0];
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                SampleLineGroup? samplelineGroup = tr.GetObject(sampleLineGroup, OpenMode.ForRead) as SampleLineGroup;

                //check number of sampleline in samplelinegroup
                int numberSampleline = 1;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                foreach (ObjectId sampleLineId in samplelineGroup.GetSampleLineIds())
                {
                    numberSampleline++;
                }
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                // get coordinate and name of point
                string[] samplelineName = new string[numberSampleline];
                string[] eastings = new string[numberSampleline];
                string[] northings = new string[numberSampleline];
                string[] elevation = new string[numberSampleline];
                int orderSampleline = 1;
                foreach (ObjectId sampleLineId in samplelineGroup.GetSampleLineIds())
                {
                    SampleLine? sampleline = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Double station = sampleline.Station;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    double easting = 0;
                    double northing = 0;
                    double elevate = 0;
                    alignment.PointLocation(station, 0, ref easting, ref northing);
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    elevate = civSurface.FindElevationAtXY(easting, northing);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    samplelineName[orderSampleline] = sampleline.Name.ToUpper();
                    eastings[orderSampleline] = Convert.ToString(Math.Round(easting, 3));
                    northings[orderSampleline] = Convert.ToString(Math.Round(northing, 3));
                    elevation[orderSampleline] = Convert.ToString(Math.Round(elevate, 3));
                    orderSampleline++;
                }

                //create a coordinate table
                UtilitiesCAD.CreateTableCoordinate2(numberSampleline, 4, alignment.Name, samplelineName, eastings, northings, elevation);

                // draw polyline for check coordinnate
                UtilitiesCAD.CreateOpenPolyline(numberSampleline, eastings, northings);


                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_UPdate2Table")]
        public static void ATUPdate2Table()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                ObjectId tableId1 = UserInput.GTable("Chọn 1 bảng " + "for source table:");
                ATable? table1 = tr.GetObject(tableId1, OpenMode.ForWrite) as ATable;
                ObjectId tableId2 = UserInput.GTable("Chọn 1 bảng " + "for destinate table:");
                ATable? table2 = tr.GetObject(tableId2, OpenMode.ForWrite) as ATable;

                //Update table
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                if (table2.Rows.Count == table1.Rows.Count)
                {
                    for (int i = 2; i < table2.Rows.Count; i++)
                    {
                        table2.Cells[i, 0].TextHeight = 2;
                        table2.Cells[i, 0].TextString = table1.Cells[i, 0].TextString;
                        table2.Cells[i, 0].Alignment = CellAlignment.MiddleCenter;

                        table2.Cells[i, 1].TextHeight = 2;
                        table2.Cells[i, 1].TextString = table1.Cells[i, 1].TextString;
                        table2.Cells[i, 1].Alignment = CellAlignment.MiddleCenter;

                        table2.Cells[i, 2].TextHeight = 2;
                        table2.Cells[i, 2].TextString = table1.Cells[i, 2].TextString;
                        table2.Cells[i, 2].Alignment = CellAlignment.MiddleCenter;
                    }
                    ;
                }
#pragma warning restore CS8602 // Dereference of a possibly null reference.
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                ;

                if (table2.Rows.Count < table1.Rows.Count)
                {
                    table2.InsertRows(3, 5, table1.Rows.Count - table2.Rows.Count);
                    for (int i = 2; i < table2.Rows.Count; i++)
                    {
                        table2.Cells[i, 0].TextHeight = 2;
                        table2.Cells[i, 0].TextString = table1.Cells[i, 0].TextString;
                        table2.Cells[i, 0].Alignment = CellAlignment.MiddleCenter;

                        table2.Cells[i, 1].TextHeight = 2;
                        table2.Cells[i, 1].TextString = table1.Cells[i, 1].TextString;
                        table2.Cells[i, 1].Alignment = CellAlignment.MiddleCenter;

                        table2.Cells[i, 2].TextHeight = 2;
                        table2.Cells[i, 2].TextString = table1.Cells[i, 2].TextString;
                        table2.Cells[i, 2].Alignment = CellAlignment.MiddleCenter;

                    }
                    ;
                }
                ;
                if (table2.Rows.Count > table1.Rows.Count)
                {
                    table2.DeleteRows(3, table2.Rows.Count - table1.Rows.Count);
                    for (int i = 2; i < table2.Rows.Count; i++)
                    {
                        table2.Cells[i, 0].TextHeight = 2;
                        table2.Cells[i, 0].TextString = table1.Cells[i, 0].TextString;
                        table2.Cells[i, 0].Alignment = CellAlignment.MiddleCenter;

                        table2.Cells[i, 1].TextHeight = 2;
                        table2.Cells[i, 1].TextString = table1.Cells[i, 1].TextString;
                        table2.Cells[i, 1].Alignment = CellAlignment.MiddleCenter;

                        table2.Cells[i, 2].TextHeight = 2;
                        table2.Cells[i, 2].TextString = table1.Cells[i, 2].TextString;
                        table2.Cells[i, 2].Alignment = CellAlignment.MiddleCenter;

                    }
                    ;
                }
                ;


                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }
           
        [CommandMethod("CTS_ChenCoc_TrenTracDoc")]
        public static void CTSChenCocTrenTracDoc()
        {
            // start transantion
            _ = new            // start transantion
            UserInput();
            _ = new UtilitiesCAD();
            _ = new UtilitiesC3D();

            ObjectId profileViewId = UserInput.GProfileViewId("\n Chọn 1 bảng trắc dọc " + " để chèn cọc: \n");
            UtilitiesC3D.SetDefaultPointSetting("CDTN", "CDTN");

            String answer = "y"; ;
            while (answer == "y" )
            {
                using (Transaction tr = A.Db.TransactionManager.StartTransaction())
                {
                    try
                    {                        

                        //start here
                        //get profileview

                        ProfileView? profileView = tr.GetObject(profileViewId, OpenMode.ForRead) as ProfileView;

#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        ObjectId alignmentId = profileView.AlignmentId;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;

#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        ObjectId sampleLineGroupId = alignment.GetSampleLineGroupIds()[0];
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        SampleLineGroup? sampleLineGroup = tr.GetObject(sampleLineGroupId, OpenMode.ForWrite) as SampleLineGroup;

                        // get point for add sampleline
                        Point3d point3D = UserInput.GPoint("\n Chọn vị trí điểm" + "để thêm cọc ");
                        double stations = 0, elevations = 0;
                        profileView.FindStationAndElevationAtXY(point3D.X, point3D.Y, ref stations, ref elevations);

                        //get sampleline name
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                        String samplelineName = UserInput.GString("Input sampleline name:");
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.

                        // set station variable
#pragma warning disable CS8604 // Possible null reference argument.
                        ObjectId samplelineId = UtilitiesC3D.CreateSampleline(samplelineName, sampleLineGroupId, alignment, stations);
#pragma warning restore CS8604 // Possible null reference argument.
                        SampleLine? sampleline = tr.GetObject(samplelineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        sampleline.StyleName = "Road Sample Line";
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                        tr.Commit();
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception e)
                    {
                        A.Ed.WriteMessage(e.Message);
                    }
                }
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                String answer1 = UserInput.GString("Do you want to add more sampleline? (y/n)");
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                answer = answer1;
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
            }
            
        }
        
        [CommandMethod("CTS_CHENCOC_TRENTRACNGANG")]
        public static void CTSCHENCOCTRENTRACNGANG()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                ObjectId sectionViewId = UserInput.GSectionView("Chọn 1 bảng cắt ngang " + ": \n");
                SectionView? sectionView = tr.GetObject(sectionViewId, OpenMode.ForRead) as SectionView;
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                string decription = UserInput.GString("Nhập mã mô tả cho điểm sẽ chèn: \n ");
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
                string ans = "Enter";
                while (ans == "Enter")
                {
                    Point3d point3D = UserInput.GPoint("\n Chọn vị trí điểm" + "Chọn điểm cần chèn trên trắc ngang: \n ");

                    //find offset, elevation
                    double offset = 1;
                    double elevation = 2;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    sectionView.FindOffsetAndElevationAtXY(point3D.X, point3D.Y, ref offset, ref elevation);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                    //find alignment
                    ObjectId samplelineId = sectionView.SampleLineId;
                    SampleLine? sampleLine = tr.GetObject(samplelineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    double station = sampleLine.Station;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    ObjectId alignmentId = sampleLine.GetParentAlignmentId();
                    Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForWrite) as Alignment;

                    // create point
                    double x = 1;
                    double y = 1;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    alignment.PointLocation(station, offset, ref x, ref y);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    Point3d point3D1 = new(x, y, elevation);
#pragma warning disable CS8604 // Possible null reference argument.
                    UtilitiesC3D.CreateCogoPointFromPoint3D(point3D1, decription);
#pragma warning restore CS8604 // Possible null reference argument.
                    ans = UserInput.GStopWithESC();
                }


                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }       

        [CommandMethod("CTS_DoiTenCoc_fromCogoPoint")]
        public static void CTSDoiTenCocFromCogoPoint()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                ObjectId cogoPointId = UserInput.GCogoPointId("\n Chọn cogo point " + "để lấy tên cho sampleline:\n");
                CogoPoint? cogoPoint = tr.GetObject(cogoPointId, OpenMode.ForWrite) as CogoPoint;
                ObjectIdCollection sampleLineIds = UserInput.GSelectionSet("Chọn sampleLine cần đổi tên: \n");
                foreach (ObjectId objectId in sampleLineIds)
                {
                    SampleLine? sampleLine = tr.GetObject(objectId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    sampleLine.Name = cogoPoint.PointName;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                }
                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("CTS_DoiTenCoc_TheoThuTu")]
        public static void CTSDoiTenCocTheoThuTu()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here

                ObjectIdCollection sampleLineIds = UserInput.GSelectionSet("\n Chọn sampleLine cần đổi tên: \n");
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                String tienToCoc = UserInput.GString("\n Nhập tiền tố cho tên cọc:");
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
                int thuTuCoc = UserInput.GInt("\n Nhập số thứ tự bắt đầu (1):"); ;
                foreach (ObjectId objectId in sampleLineIds)
                {
                    SampleLine? sampleLine = tr.GetObject(objectId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    sampleLine.Name = tienToCoc + thuTuCoc;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    thuTuCoc++;
                }
                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("CTS_DichCoc_TinhTien")]
        public static void CTSDichCocTinhTien()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here

                ObjectIdCollection sampleLineIds = UserInput.GSelectionSet("\n Chọn sampleLine cần dịch cọc: \n");
                double delta = UserInput.GDouble("\n Nhập khoảng dịch cọc:");
                foreach (ObjectId objectId in sampleLineIds)
                {
                    SampleLine? sampleLine = tr.GetObject(objectId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    sampleLine.Station += delta;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                }


                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("CTS_Copy_NhomCoc")]
        public static void CTSCopyNhomCoc()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here

                ObjectId alignmentId = UserInput.GAlignmentId("\n Chọn tim đường có nhóm cọc cần sao chép:");
                Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForWrite) as Alignment;
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                String tenSamplelineGroup = UserInput.GString("\n Nhập tên nhóm cọc cần tạo:");
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.

                //tạo nhóm cọc mới
                ObjectId sampleLineGroupId = SampleLineGroup.Create(tenSamplelineGroup, alignmentId);

                // lấy thông số từ cọc cũ
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectId sampleLineGroupId_0 = alignment.GetSampleLineGroupIds()[0];
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                SampleLineGroup? sampleLineGroup_0 = tr.GetObject(sampleLineGroupId_0, OpenMode.ForWrite) as SampleLineGroup;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectIdCollection sampleLineIds = sampleLineGroup_0.GetSampleLineIds();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                List<double> stationList = [];
                List<String> sampleLineNameS = [];
                foreach (ObjectId sampleLineId in sampleLineIds)
                {
                    SampleLine? sampleLine = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    stationList.Add(sampleLine.Station);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    sampleLineNameS.Add(sampleLine.Name);
                }

                //sampleLineGroup_0.Erase();
                //tạo cọc vào nhóm cọc mới
                for (int i = 0; i < stationList.Count(); i++)
                {
                    A.Ed.WriteMessage(sampleLineNameS[i].ToString() + "-" + stationList[i].ToString());
                    ObjectId sampleLineId = UtilitiesC3D.CreateSampleline(sampleLineNameS[i] + "C", sampleLineGroupId, alignment, stationList[i]);
                    SampleLine? sampleLine = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    sampleLine.Name = sampleLineNameS[i];
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                }




                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("CTS_DongBo_2_NhomCoc")]
        public static void CTSDongBo2NhomCoc()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here

                // lấy thông số từ cọc cũ
                ObjectId sampleLineId_0 = UserInput.GSampleLineId("\n Chọn nhóm cọc NGUỒN để copy: ");
                SampleLine? sampleLine_0 = tr.GetObject(sampleLineId_0, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectId sampleLineGroupId = sampleLine_0.GroupId;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                SampleLineGroup? sampleLineGroup = tr.GetObject(sampleLineGroupId, OpenMode.ForWrite) as SampleLineGroup;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectIdCollection sampleLineIds = sampleLineGroup.GetSampleLineIds();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                int sampleLIneNumber = sampleLineIds.Count;
                List<double> stationList = [];
                List<String> sampleLineNameS = [];
                foreach (ObjectId sampleLineId in sampleLineIds)
                {
                    SampleLine? sampleLine = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    stationList.Add(sampleLine.Station);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    //a.ok(sampleLine.Station.ToString());
                    sampleLineNameS.Add(sampleLine.Name);
                }

                // lấy thông số từ cọc mới
                ObjectId sampleLineId_1 = UserInput.GSampleLineId("\n Chọn nhóm cọc ĐÍCH để copy: ");
                //double lyTrinhBatDau = UI.G_Double(" NHập lý bắt đầu của tuyến: ");
                SampleLine? sampleLine_1 = tr.GetObject(sampleLineId_1, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectId sampleLineGroupId_1 = sampleLine_1.GroupId;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                SampleLineGroup? sampleLineGroup_1 = tr.GetObject(sampleLineGroupId_1, OpenMode.ForWrite) as SampleLineGroup;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectIdCollection sampleLineIds_1 = sampleLineGroup_1.GetSampleLineIds();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                int sampleLIneNumber_1 = sampleLineIds_1.Count;

                //đồng bộ thông tin
                if (sampleLIneNumber_1 <= sampleLIneNumber)
                {
                    for (int i = 0; i < sampleLIneNumber_1; i++)
                    {
                        SampleLine? sampleLine = tr.GetObject(sampleLineIds[i], OpenMode.ForWrite) as SampleLine;
                        SampleLine? sampleLine_c = tr.GetObject(sampleLineIds_1[i], OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        sampleLine_c.Station = sampleLine.Station;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        sampleLine_c.Name = sampleLine.Name + "c";
                    }
                }
                else A.Ok("Số cọc đích nhiều hơn cọc nguồn. Cần xóa bớt cọc đích");


                //rename sampleline
                char[] charsToTrim = ['z', ' ', '\'', 'c'];
                foreach (ObjectId sampleLineId in sampleLineGroup_1.GetSampleLineIds())
                {
                    SampleLine? sampleline = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    sampleline.Name = sampleline.Name.Trim(charsToTrim);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                }


                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("CTS_DongBo_2_NhomCoc_TheoDoan")]
        public static void CTSDongBo2NhomCocTheoDoan()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here

                // lấy thông số từ cọc cũ
                ObjectId sampleLineId_0 = UserInput.GSampleLineId("\n Chọn cọc ĐẦU TIÊN của đoạn thuộc NGUỒN để copy: ");
                SampleLine? sampleLine_0 = tr.GetObject(sampleLineId_0, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectId sampleLineGroupId = sampleLine_0.GroupId;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                SampleLineGroup? sampleLineGroup = tr.GetObject(sampleLineGroupId, OpenMode.ForWrite) as SampleLineGroup;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectIdCollection sampleLineIds = sampleLineGroup.GetSampleLineIds();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                int sampleLIneNumber = sampleLineIds.Count;
                int sampleLIneNumberBegin = sampleLine_0.Number;

                List<double> stationList = [];
                List<String> sampleLineNameS = [];

                for (int i = sampleLIneNumberBegin; i < sampleLineIds.Count; i++)
                {
                    ObjectId sampleLineId = sampleLineIds[i];
                    SampleLine? sampleLine = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    stationList.Add(sampleLine.Station);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    //a.ok(sampleLine.Station.ToString());
                    sampleLineNameS.Add(sampleLine.Name);
                }
                /*
                foreach (ObjectId sampleLineId in sampleLineIds)
                {
                    SampleLine sampleLine = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
                    stationList.Add(sampleLine.Station);
                    //a.ok(sampleLine.Station.ToString());
                    sampleLineNameS.Add(sampleLine.Name);
                }
                */

                // lấy thông số từ cọc mới
                ObjectId sampleLineId_1 = UserInput.GSampleLineId("\n Chọn nhóm cọc ĐÍCH để copy: ");
                //double lyTrinhBatDau = UI.G_Double(" NHập lý bắt đầu của tuyến: ");
                SampleLine? sampleLine_1 = tr.GetObject(sampleLineId_1, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectId sampleLineGroupId_1 = sampleLine_1.GroupId;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                SampleLineGroup? sampleLineGroup_1 = tr.GetObject(sampleLineGroupId_1, OpenMode.ForWrite) as SampleLineGroup;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectIdCollection sampleLineIds_1 = sampleLineGroup_1.GetSampleLineIds();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                int sampleLIneNumber_1 = sampleLineIds_1.Count;

                //đồng bộ thông tin
                if (sampleLIneNumber_1 <= sampleLIneNumber)
                {
                    for (int i = 0; i < sampleLIneNumber_1; i++)
                    {
                        SampleLine? sampleLine = tr.GetObject(sampleLineIds[i + sampleLIneNumberBegin - 1], OpenMode.ForWrite) as SampleLine;
                        SampleLine? sampleLine_c = tr.GetObject(sampleLineIds_1[i], OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        sampleLine_c.Station = sampleLine.Station;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        sampleLine_c.Name = sampleLine.Name + "c";
                    }
                }
                else A.Ok("Số cọc đích nhiều hơn cọc nguồn. Cần xóa bớt cọc đích");


                //rename sampleline
                char[] charsToTrim = ['z', ' ', '\'', 'c'];
                foreach (ObjectId sampleLineId in sampleLineGroup_1.GetSampleLineIds())
                {
                    SampleLine? sampleline = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    sampleline.Name = sampleline.Name.Trim(charsToTrim);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                }


                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }


        [CommandMethod("CTS_DoiTenCoc_H")]
        public static void CTSDoiTenCocH()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here

                ObjectId sampleLineId_S = UserInput.GSampleLineId("\n Chọn sampleLine cọc mốc: \n");
                SampleLine? sampleLine_S = tr.GetObject(sampleLineId_S, OpenMode.ForWrite) as SampleLine;
                ObjectId sampleLineId_D = UserInput.GSampleLineId("\n Chọn sampleLine cọc cần đổi tên: \n");
                SampleLine? sampleLine_D = tr.GetObject(sampleLineId_D, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                sampleLine_D.Name = sampleLine_S.Name + "A";
#pragma warning restore CS8602 // Dereference of a possibly null reference.
#pragma warning restore CS8602 // Dereference of a possibly null reference.



                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

               
        [CommandMethod("CTS_Copy_BeRong_sampleLine")]
        public static void CTSCopyBeRongSampleLine()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here

                ObjectId sampleLineId = UserInput.GSampleLineId("Chọn sampleLine GỐC cần sao chép bề rộng:");
                double offsetLeft = new();
                double offsetRight = new();
                Point3d point3dLeft = new();
                Point3d point3dRight = new();
                Point3d point3dCenter = new();
                SampleLine? sampleLine = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8604 // Possible null reference argument.
                ObjectId alignmentId = sampleLine.GetParentAlignmentId();
#pragma warning restore CS8604 // Possible null reference argument.
                Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForWrite) as Alignment;
                if (sampleLine.Vertices.Count > 3)
                {
                    A.Ok("Kiểm tra lại sampleline có nhiều hơn 3 điểm");
                }
                foreach (SampleLineVertex vertex in sampleLine.Vertices)
                {

                    if (vertex.Side == SampleLineVertexSideType.Left)
                    {
                        point3dLeft = vertex.Location;
                    }
                    if (vertex.Side == SampleLineVertexSideType.Right)
                    {
                        point3dRight = vertex.Location;
                    }
                    if (vertex.Side == SampleLineVertexSideType.Center)
                    {
                        point3dCenter = vertex.Location;
                    }
                    offsetLeft = point3dCenter.DistanceTo(point3dLeft);
                    offsetRight = point3dCenter.DistanceTo(point3dRight);
                }
                double station = new();
                double offset = new();
                double easting = new();
                double northing = new();
                ObjectIdCollection sampleLineIds = UserInput.GSelectionSetWithType("Chọn các sample cần copy bề rộng: \n", "AECC_SAMPLE_LINE");
                A.OkHere();
                foreach (ObjectId SLineId in sampleLineIds)
                {
                    SampleLine? sampleLine1 = tr.GetObject(SLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    if (sampleLine1.Vertices.Count > 3)
                    {
                        A.Ok("Kiểm tra lại sampleline có nhiều hơn 3 điểm");
                    }
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    foreach (SampleLineVertex vertex in sampleLine1.Vertices)
                    {
                        if (vertex.Side == SampleLineVertexSideType.Center)
                        {
                            point3dCenter = vertex.Location;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                            alignment.StationOffset(point3dCenter.X, point3dCenter.Y, ref station, ref offset);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        }
                    }
                    foreach (SampleLineVertex vertex in sampleLine1.Vertices)
                    {
                        if (vertex.Side == SampleLineVertexSideType.Left)
                        {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                            alignment.PointLocation(station, -offsetLeft, ref easting, ref northing);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                            vertex.Location = new Point3d(easting, northing, 0);
                        }
                    }
                    foreach (SampleLineVertex vertex in sampleLine1.Vertices)
                    {
                        if (vertex.Side == SampleLineVertexSideType.Right)
                        {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                            alignment.PointLocation(station, offsetRight, ref easting, ref northing);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                            vertex.Location = new Point3d(easting, northing, 0);
                        }
                    }
                }






                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("CTS_Thaydoi_BeRong_sampleLine")]
        public static void CTSThaydoiBeRongSampleLine()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here

                ObjectId sampleLineId = UserInput.GSampleLineId("Chọn sampleLine  cần thay đổi bề rộng:");
                Point3d point3dCenter = new();
                SampleLine? sampleLine = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8604 // Possible null reference argument.
                ObjectId alignmentId = sampleLine.GetParentAlignmentId();
#pragma warning restore CS8604 // Possible null reference argument.
                Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForWrite) as Alignment;

                double station = new();
                double offset = new();
                double stationLeft = new();
                double offsetLeft = new();
                double stationRight = new();
                double offsetRight = new();
                double easting = new();
                double northing = new();

                Point3d pointLeft = UserInput.GPoint("Chọn vị trí offset trái của sampleLine đã chọn");
                Point3d pointRight = UserInput.GPoint("Chọn vị trí offset phải của sampleLine đã chọn");

#pragma warning disable CS8602 // Dereference of a possibly null reference.
                alignment.StationOffset(pointLeft.X, pointLeft.Y, ref stationLeft, ref offsetLeft);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                alignment.StationOffset(pointRight.X, pointRight.Y, ref stationRight, ref offsetRight);


                if (sampleLine.Vertices.Count > 3)
                {
                    A.Ok("Kiểm tra lại sampleline có nhiều hơn 3 điểm");
                }
                foreach (SampleLineVertex vertex in sampleLine.Vertices)
                {
                    if (vertex.Side == SampleLineVertexSideType.Center)
                    {
                        point3dCenter = vertex.Location;
                        alignment.StationOffset(point3dCenter.X, point3dCenter.Y, ref station, ref offset);
                    }
                }
                foreach (SampleLineVertex vertex in sampleLine.Vertices)
                {
                    if (vertex.Side == SampleLineVertexSideType.Left)
                    {
                        alignment.PointLocation(station, offsetLeft, ref easting, ref northing);
                        vertex.Location = new Point3d(easting, northing, 0);
                    }
                }
                foreach (SampleLineVertex vertex in sampleLine.Vertices)
                {
                    if (vertex.Side == SampleLineVertexSideType.Right)
                    {
                        alignment.PointLocation(station, offsetRight, ref easting, ref northing);
                        vertex.Location = new Point3d(easting, northing, 0);
                    }
                }







                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }
        [CommandMethod("CTS_Offset_BeRong_sampleLine")]
        public static void CTSOffsetBeRongSampleLine()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            try
            {
                var selOpts = new PromptSelectionOptions { MessageForAdding = "\nChọn các sample lines:" };
                var filter = new SelectionFilter([new TypedValue(0, "AECC_SAMPLE_LINE")]);
                var selRes = ed.GetSelection(selOpts, filter);
                if (selRes.Status != PromptStatus.OK) return;

                var peLeft = new PromptEntityOptions("\nChọn Polyline biên TRÁI: ");
                peLeft.SetRejectMessage("\nPhải là Polyline.");
                peLeft.AddAllowedClass(typeof(Polyline), true);
                var perLeft = ed.GetEntity(peLeft);
                if (perLeft.Status != PromptStatus.OK) return;

                var peRight = new PromptEntityOptions("\nChọn Polyline biên PHẢI: ");
                peRight.SetRejectMessage("\nPhải là Polyline.");
                peRight.AddAllowedClass(typeof(Polyline), true);
                var perRight = ed.GetEntity(peRight);
                if (perRight.Status != PromptStatus.OK) return;

                using var tr = db.TransactionManager.StartTransaction();
                var plLeft = (Polyline)tr.GetObject(perLeft.ObjectId, OpenMode.ForRead);
                var plRight = (Polyline)tr.GetObject(perRight.ObjectId, OpenMode.ForRead);

                int updated = 0, noCenter = 0, noLeft = 0, noRight = 0, missLeft = 0, missRight = 0;

                foreach (ObjectId id in selRes.Value.GetObjectIds())
                {
                    var sl = tr.GetObject(id, OpenMode.ForWrite) as SampleLine;
                    if (sl == null) continue;

                    var centerV = sl.Vertices.Cast<SampleLineVertex>().FirstOrDefault(v => v.Side == SampleLineVertexSideType.Center);
                    if (centerV == null) { noCenter++; continue; }
                    var leftV = sl.Vertices.Cast<SampleLineVertex>().FirstOrDefault(v => v.Side == SampleLineVertexSideType.Left);
                    if (leftV == null) { noLeft++; continue; }
                    var rightV = sl.Vertices.Cast<SampleLineVertex>().FirstOrDefault(v => v.Side == SampleLineVertexSideType.Right);
                    if (rightV == null) { noRight++; continue; }

                    Point3d centerPt = centerV.Location;
                    Vector3d dirLeft = (leftV.Location - centerPt); if (dirLeft.Length < 1e-6) dirLeft = new Vector3d(-1, 0, 0); dirLeft = dirLeft.GetNormal();
                    Vector3d dirRight = (rightV.Location - centerPt); if (dirRight.Length < 1e-6) dirRight = new Vector3d(1, 0, 0); dirRight = dirRight.GetNormal();

                    Point3d? hitLeft = IntersectOneDirection(centerPt, dirLeft, plLeft);
                    if (hitLeft == null) { missLeft++; continue; }
                    Point3d? hitRight = IntersectOneDirection(centerPt, dirRight, plRight);
                    if (hitRight == null) { missRight++; continue; }

                    leftV.Location = new(hitLeft.Value.X, hitLeft.Value.Y, leftV.Location.Z);
                    rightV.Location = new(hitRight.Value.X, hitRight.Value.Y, rightV.Location.Z);
                    updated++;
                }

                tr.Commit();
                ed.WriteMessage($"\nCập nhật xong: {updated}. Không center: {noCenter}, không left: {noLeft}, không right: {noRight}, không giao trái: {missLeft}, không giao phải: {missRight}.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nLỗi: " + ex.Message);
            }
        }

        private static Point3d? IntersectOneDirection(Point3d origin, Vector3d dir, Polyline boundary)
        {
            const double maxLen = 100000.0; // 100 km để đủ dài
            using var ray = new Line(origin, origin + dir * maxLen);
            var pts = new Point3dCollection();
            try { boundary.IntersectWith(ray, Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero); }
            catch { return null; }
            if (pts.Count == 0) return null;
            double best = double.MaxValue; Point3d bestPt = Point3d.Origin;
            for (int i = 0; i < pts.Count; i++)
            {
                double d = (pts[i] - origin).DotProduct(dir);
                if (d > 1e-6 && d < best) { best = d; bestPt = pts[i]; }
            }
            return best == double.MaxValue ? null : bestPt;
        }






























































    }

    /// <summary>
    /// Form phát sinh cọc chuyên nghiệp kiểu Civil 3D
    /// Cho phép chọn tuyến, chọn loại cọc, xem preview trước khi tạo
    /// </summary>
    public class PhatSinhCocForm : CivilFormBase
    {
        // Alias giải quyết xung đột namespace
        // Dùng System.Drawing.Font thay vì Font (trùng Autodesk.AutoCAD.DatabaseServices.Font)
        // Dùng System.Drawing.Point thay vì Point (trùng Autodesk.Civil.DatabaseServices.Point)
        // ─── UI Controls ───────────────────────────────────────────
        private ComboBox cboAlignment = null!;
        private System.Windows.Forms.Label lblLength = null!;
        private System.Windows.Forms.Label lblGroupCount = null!;

        private CheckBox chkDacBiet = null!;
        private CheckBox chkHKm = null!;
        private NumericUpDown numKmStart = null!;
        private CheckBox chkChiTiet = null!;
        private NumericUpDown numKhoangCach = null!;
        private CheckBox chkLabel = null!;

        private DataGridView dgvPreview = null!;
        private System.Windows.Forms.Label lblTotal = null!;

        private Button btnPreview = null!;
        private Button btnCreate = null!;
        private Button btnCancel = null!;

        // ─── Data ──────────────────────────────────────────────────
        private List<AlignmentInfo> alignments = new();
        public bool IsConfirmed { get; private set; } = false;

        private class AlignmentInfo
        {
            public ObjectId Id { get; set; }
            public string Name { get; set; } = "";
            public double Length { get; set; }
            public int GroupCount { get; set; }
        }

        public class CocPreviewItem
        {
            public int STT { get; set; }
            public string TenCoc { get; set; } = "";
            public double Station { get; set; }
            public string LyTrinh { get; set; } = "";
            public string Loai { get; set; } = "";
        }

        private List<CocPreviewItem> previewItems = new();

        // ─── Constructor ────────────────────────────────────────────
        public PhatSinhCocForm()
        {
            LoadAlignments();
            InitializeComponent();
        }

        private void LoadAlignments()
        {
            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                CivilDocument cdoc = CivilApplication.ActiveDocument;
                ObjectIdCollection alignIds = cdoc.GetAlignmentIds();
                foreach (ObjectId id in alignIds)
                {
                    Alignment? al = tr.GetObject(id, OpenMode.ForRead) as Alignment;
                    if (al == null) continue;
                    alignments.Add(new AlignmentInfo
                    {
                        Id = id,
                        Name = al.Name,
                        Length = al.Length,
                        GroupCount = al.GetSampleLineGroupIds().Count
                    });
                }
                tr.Commit();
            }
            catch { }
        }

        private void InitializeComponent()
        {
            // ─── Form ──────────────────────────────────────────
            InitBaseTheme("PHÁT SINH CỌC", 780, 720);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            var header = CreateHeader("PHÁT SINH CỌC");
            this.Controls.Add(header);

            int pad = 15, y = 60;

            // ═══════════════════════════════════════════════════
            // SECTION 1: CHỌN TUYẾN
            // ═══════════════════════════════════════════════════
            var grpAlign = CreateSection("Chọn tuyến", pad, y, 740, 85);
            this.Controls.Add(grpAlign);

            grpAlign.Controls.Add(CreateLabel("Tuyến:", 12, 33));
            cboAlignment = new ComboBox
            {
                Location = new System.Drawing.Point(62, 30),
                Size = new Size(300, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextDark,
                FlatStyle = FlatStyle.Flat,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            foreach (var al in alignments)
                cboAlignment.Items.Add(al.Name);
            if (cboAlignment.Items.Count > 0) cboAlignment.SelectedIndex = 0;
            cboAlignment.SelectedIndexChanged += (s, e) => UpdateAlignmentInfo();
            grpAlign.Controls.Add(cboAlignment);

            lblLength = CreateLabel("Chiều dài: —", 380, 33);
            grpAlign.Controls.Add(lblLength);

            lblGroupCount = CreateLabel("Nhóm cọc: —", 560, 33);
            grpAlign.Controls.Add(lblGroupCount);

            // Hiển thị thông tin ban đầu
            if (alignments.Count > 0) UpdateAlignmentInfo();

            y += 95;

            // ═══════════════════════════════════════════════════
            // SECTION 2: TÙY CHỌN PHÁT SINH
            // ═══════════════════════════════════════════════════
            var grpOptions = CreateSection("Tùy chọn phát sinh cọc", pad, y, 740, 165);
            this.Controls.Add(grpOptions);

            // --- Row 1: Cọc đặc biệt ---
            chkDacBiet = new CheckBox
            {
                Text = "Cọc đặc biệt (điểm hình học: DT, TD, P, TC, CT...)",
                Location = new System.Drawing.Point(15, 33),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            grpOptions.Controls.Add(chkDacBiet);

            // --- Row 2: Cọc H/Km ---
            chkHKm = new CheckBox
            {
                Text = "Cọc H / Km (mỗi 100m)",
                Location = new System.Drawing.Point(15, 63),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            chkHKm.CheckedChanged += (s, e) => numKmStart.Enabled = chkHKm.Checked;
            grpOptions.Controls.Add(chkHKm);

            grpOptions.Controls.Add(CreateLabel("Km bắt đầu:", 400, 65));
            numKmStart = new NumericUpDown
            {
                Location = new System.Drawing.Point(490, 62),
                Size = new Size(80, 28),
                Minimum = 0,
                Maximum = 999,
                Value = 0,
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            grpOptions.Controls.Add(numKmStart);

            // --- Row 3: Cọc chi tiết ---
            chkChiTiet = new CheckBox
            {
                Text = "Cọc chi tiết (cọc phụ giữa các cọc chính)",
                Location = new System.Drawing.Point(15, 93),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            chkChiTiet.CheckedChanged += (s, e) => numKhoangCach.Enabled = chkChiTiet.Checked;
            grpOptions.Controls.Add(chkChiTiet);

            grpOptions.Controls.Add(CreateLabel("Khoảng cách (m):", 370, 95));
            numKhoangCach = new NumericUpDown
            {
                Location = new System.Drawing.Point(490, 92),
                Size = new Size(80, 28),
                Minimum = 5,
                Maximum = 100,
                Value = 20,
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            grpOptions.Controls.Add(numKhoangCach);

            // --- Row 4: Tạo label ---
            chkLabel = new CheckBox
            {
                Text = "Tạo Label tên cọc trên mặt bằng",
                Location = new System.Drawing.Point(15, 123),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            grpOptions.Controls.Add(chkLabel);

            y += 175;

            // ═══════════════════════════════════════════════════
            // SECTION 3: PREVIEW DANH SÁCH CỌC
            // ═══════════════════════════════════════════════════
            var grpPreview = CreateSection("Preview danh sách cọc", pad, y, 740, 310);
            this.Controls.Add(grpPreview);

            dgvPreview = new DataGridView
            {
                Location = new System.Drawing.Point(10, 30),
                Size = new Size(720, 240),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = BgInput,
                GridColor = BorderColor,
                BorderStyle = BorderStyle.FixedSingle,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = BgInput,
                    ForeColor = TextDark,
                    SelectionBackColor = NavyMedium,
                    SelectionForeColor = System.Drawing.Color.White,
                    Font = new System.Drawing.Font("Segoe UI", 9f)
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = NavyDark,
                    ForeColor = System.Drawing.Color.White,
                    Font = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 32
            };

            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "STT", Width = 50, Name = "colSTT" });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Tên cọc", Width = 140, Name = "colTen" });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Lý trình", Width = 130, Name = "colLyTrinh" });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Station (m)", Width = 110, Name = "colStation" });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Loại cọc", Width = 120, Name = "colLoai" });

            grpPreview.Controls.Add(dgvPreview);

            lblTotal = CreateLabel("Tổng: 0 cọc", 10, 280);
            lblTotal.Font = new System.Drawing.Font("Segoe UI", 10f, FontStyle.Bold);
            grpPreview.Controls.Add(lblTotal);

            // ═══════════════════════════════════════════════════
            // BUTTONS
            // ═══════════════════════════════════════════════════
            var pnlBottom = CreateBottomBar();
            this.Controls.Add(pnlBottom);

            btnPreview = CreateButton("Preview", 420, 8, 100, 36, pnlBottom);
            btnPreview.Click += (s, e) => GeneratePreview();

            btnCancel = CreateButton("Hủy", 530, 8, 100, 36, pnlBottom);
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            btnCreate = CreatePrimaryButton("Tạo cọc", 640, 8, 110, 36, pnlBottom);
            btnCreate.Click += (s, e) => { IsConfirmed = true; this.DialogResult = DialogResult.OK; this.Close(); };

            // Auto-preview khi mở form
            if (alignments.Count > 0) GeneratePreview();
        }

        // ─── Logic ──────────────────────────────────────────────
        private void UpdateAlignmentInfo()
        {
            if (cboAlignment.SelectedIndex < 0 || cboAlignment.SelectedIndex >= alignments.Count) return;
            var al = alignments[cboAlignment.SelectedIndex];
            lblLength.Text = $"Chiều dài: {al.Length:F2} m";
            lblGroupCount.Text = $"Nhóm cọc: {al.GroupCount}";
        }

        private AlignmentInfo? GetSelectedAlignment()
        {
            if (cboAlignment.SelectedIndex < 0 || cboAlignment.SelectedIndex >= alignments.Count) return null;
            return alignments[cboAlignment.SelectedIndex];
        }

        private string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double remainder = station - km * 1000;
            return $"{km}+{remainder:F3}";
        }

        /// <summary>
        /// Tính toán preview danh sách cọc (không tạo thực tế)
        /// </summary>
        private void GeneratePreview()
        {
            previewItems.Clear();
            dgvPreview.Rows.Clear();

            var alInfo = GetSelectedAlignment();
            if (alInfo == null) return;

            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                Alignment? alignment = tr.GetObject(alInfo.Id, OpenMode.ForRead) as Alignment;
                if (alignment == null) { tr.Commit(); return; }

                // Danh sách station + tên
                List<(double station, string name, string loai)> allCoc = new();

                // ── 1. Cọc đặc biệt ──
                if (chkDacBiet.Checked)
                {
                    Station[] stations = alignment.GetStationSet(StationTypes.GeometryPoint);
                    int curveIdx = 1;
                    for (int i = 0; i < stations.Length; i++)
                    {
                        string name = "";
                        var gpt = stations[i].GeometryStationType;

                        if (gpt == AlignmentGeometryPointStationType.BegOfAlign)
                            name = "DT";
                        else if (gpt == AlignmentGeometryPointStationType.EndOfAlign)
                            name = "CT";
                        else if (gpt == AlignmentGeometryPointStationType.TanTan)
                        { name = "D" + curveIdx; curveIdx++; }
                        else if (gpt == AlignmentGeometryPointStationType.TanCurve)
                        {
                            name = "TD" + curveIdx;
                            // Thêm P và TC nếu có
                            if (i + 1 < stations.Length)
                                allCoc.Add((stations[i + 1].RawStation, "P" + curveIdx, "Đặc biệt"));
                            if (i + 2 < stations.Length)
                                allCoc.Add((stations[i + 2].RawStation, "TC" + curveIdx, "Đặc biệt"));
                            curveIdx++;
                        }
                        else if (gpt == AlignmentGeometryPointStationType.TanSpiral)
                        {
                            name = "ND" + curveIdx;
                            if (i + 1 < stations.Length)
                                allCoc.Add((stations[i + 1].RawStation, "TD" + curveIdx, "Đặc biệt"));
                            if (i + 2 < stations.Length)
                                allCoc.Add((stations[i + 2].RawStation, "P" + curveIdx, "Đặc biệt"));
                            if (i + 3 < stations.Length)
                                allCoc.Add((stations[i + 3].RawStation, "TC" + curveIdx, "Đặc biệt"));
                            if (i + 4 < stations.Length)
                                allCoc.Add((stations[i + 4].RawStation, "NC" + curveIdx, "Đặc biệt"));
                            curveIdx++;
                        }

                        if (!string.IsNullOrEmpty(name))
                            allCoc.Add((stations[i].RawStation, name, "Đặc biệt"));
                    }
                }

                // ── 2. Cọc H/Km ──
                if (chkHKm.Checked)
                {
                    int kmStart = (int)numKmStart.Value;
                    int hNum = 1;
                    int currentKm = kmStart;

                    for (double st = 100; st < alignment.Length; st += 100)
                    {
                        if (st % 1000 == 0)
                        {
                            currentKm = (int)(st / 1000) + kmStart;
                            allCoc.Add((st, $"Km{currentKm}", "H/Km"));
                            hNum = 1;
                        }
                        else
                        {
                            allCoc.Add((st, $"H{hNum} (Km{currentKm})", "H/Km"));
                            hNum++;
                        }
                    }
                }

                // ── 3. Cọc chi tiết ──
                if (chkChiTiet.Checked)
                {
                    // Lấy tất cả station đã có → sắp xếp → chèn chi tiết
                    var existingStations = allCoc.Select(c => c.station).Distinct().OrderBy(s => s).ToList();
                    // Thêm đầu và cuối nếu chưa có
                    if (existingStations.Count == 0 || existingStations[0] > 0.01)
                        existingStations.Insert(0, 0);
                    if (existingStations.Count == 0 || existingStations.Last() < alignment.Length - 0.01)
                        existingStations.Add(alignment.Length);

                    int khoangCach = (int)numKhoangCach.Value;
                    int detailNum = 1;
                    for (int i = 0; i < existingStations.Count - 1; i++)
                    {
                        double start = existingStations[i];
                        double end = existingStations[i + 1];
                        int segments = (int)Math.Ceiling((end - start) / khoangCach);
                        for (int j = 1; j < segments; j++)
                        {
                            double st = start + j * khoangCach;
                            if (st < end - 0.5) // tránh trùng điểm cuối
                            {
                                allCoc.Add((st, detailNum.ToString(), "Chi tiết"));
                                detailNum++;
                            }
                        }
                    }
                }

                // ── Sắp xếp và đánh số ──
                allCoc = allCoc.OrderBy(c => c.station).ThenByDescending(c => c.loai == "Đặc biệt" ? 2 : c.loai == "H/Km" ? 1 : 0).ToList();

                // Loại bỏ trùng station (giữ cọc ưu tiên cao hơn)
                var deduped = new List<(double station, string name, string loai)>();
                foreach (var coc in allCoc)
                {
                    if (deduped.Count == 0 || Math.Abs(deduped.Last().station - coc.station) > 0.01)
                        deduped.Add(coc);
                }

                // Build preview items
                int stt = 1;
                foreach (var coc in deduped)
                {
                    previewItems.Add(new CocPreviewItem
                    {
                        STT = stt++,
                        TenCoc = coc.name,
                        Station = coc.station,
                        LyTrinh = FormatStation(coc.station),
                        Loai = coc.loai
                    });
                }

                // ── Cập nhật DataGridView ──
                foreach (var item in previewItems)
                {
                    int idx = dgvPreview.Rows.Add(item.STT, item.TenCoc, item.LyTrinh, $"{item.Station:F3}", item.Loai);

                    // Tô màu theo loại cọc — chỉ bold cọc đặc biệt
                    var row = dgvPreview.Rows[idx];
                    if (item.Loai == "Đặc biệt")
                        row.DefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Bold);
                    else if (item.Loai == "H/Km")
                        row.DefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 9f, FontStyle.Italic);
                    // Chi tiết: giữ mặc định
                }

                lblTotal.Text = $"Tổng: {previewItems.Count} cọc  |  " +
                    $"Đặc biệt: {previewItems.Count(c => c.Loai == "Đặc biệt")}  |  " +
                    $"H/Km: {previewItems.Count(c => c.Loai == "H/Km")}  |  " +
                    $"Chi tiết: {previewItems.Count(c => c.Loai == "Chi tiết")}";

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                lblTotal.Text = $"Lỗi: {ex.Message}";
                lblTotal.ForeColor = Color.FromArgb(255, 80, 80);
            }
        }

        // ─── Public API ─────────────────────────────────────────
        public ObjectId GetAlignmentId()
        {
            var al = GetSelectedAlignment();
            return al?.Id ?? ObjectId.Null;
        }

        public bool CreateDacBiet => chkDacBiet.Checked;
        public bool CreateHKm => chkHKm.Checked;
        public int KmStart => (int)numKmStart.Value;
        public bool CreateChiTiet => chkChiTiet.Checked;
        public int KhoangCach => (int)numKhoangCach.Value;
        public bool ShouldCreateLabel => chkLabel.Checked;
        public List<CocPreviewItem> PreviewItems => previewItems;
    }
}

