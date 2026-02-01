// (C) Copyright 2015 by  
//
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.Runtime;
using Acad = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using ATable = Autodesk.AutoCAD.DatabaseServices.Table;
using AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application;

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
using Entity = Autodesk.AutoCAD.DatabaseServices.Entity;
using Autodesk.AutoCAD.Colors;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using MyFirstProject.Extensions;
// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(Civil3DCsharp.CAD))]

namespace Civil3DCsharp
{
    class CAD
    {
        [CommandMethod("AT_TongDoDai_Full")]
        public static void AT_TongDoDai_Full()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                ObjectIdCollection polyLineColl = UserInput.GSelectionSet("Chọn các đối tượng (Polyline/Hatch) cần tính tổng độ dài: \n");
                List<String> listLengthPolyline = [];

                //tạo field data
                string str1 = "%<\\AcObjProp Object(%<\\_ObjId ";
                string str2 = ">%).Length \\f \"%lu2\">%";
                string format = "";
                foreach (ObjectId polyLineId in polyLineColl)
                {
                    string strId = polyLineId.OldIdPtr.ToString();
                    format = str1 + strId + str2;
                    listLengthPolyline.Add(format);
                }

                //tính tổng data file
                string formatTong = "0";
                string str3 = "%<\\AcExpr (";
                string str4 = ") \\f \" %lu2\">%";
                foreach (String item in listLengthPolyline)
                {
                    formatTong = formatTong + "+" + item;
                }
                String formatTong_Lite = formatTong[2..];
                formatTong = formatTong_Lite + "=" + str3 + formatTong + str4 + "m";
                // vẽ text
                Point3d point = UserInput.GPoint("Chọn vị trí đặt text: \n");
                LayerTableRecord layer = UtilitiesCAD.CCreateLayer("TinhTong");
                UtilitiesCAD.CCreateText2(point, 2, formatTong, "TinhTong", "Standard");
                Clipboard.SetText(formatTong);



                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_TongDoDai_Replace")]
        public static void AT_TongDoDai_Replace()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                ObjectIdCollection polyLineColl = UserInput.GSelectionSet("Chọn các đối tượng (Polyline/Hatch) cần tính tổng độ dài: \n");
                List<String> listLengthPolyline = [];

                //tạo field data
                string str1 = "%<\\AcObjProp Object(%<\\_ObjId ";
                string str2 = ">%).Length \\f \"%lu2\">%";
                string format = "";
                foreach (ObjectId polyLineId in polyLineColl)
                {
                    string strId = polyLineId.OldIdPtr.ToString();
                    format = str1 + strId + str2;
                    listLengthPolyline.Add(format);
                }

                //tính tổng data file
                string formatTong = "0";
                string str3 = "%<\\AcExpr (";
                string str4 = ") \\f \" %lu2\">%";
                foreach (String item in listLengthPolyline)
                {
                    formatTong = formatTong + "+" + item;
                }
                formatTong = str3 + formatTong + str4 + "m";
                // vẽ text
                ObjectId textId = UserInput.GDbText("Chọn text cần nhập nội dung: \n");
                DBText? text = tr.GetObject(textId, OpenMode.ForWrite) as DBText;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                text.TextString = formatTong;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                A.Ed.Command("_UPDATEFIELD", textId);

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_TongDoDai_Replace2")]
        public static void AT_TongDoDai_Replace2()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                ObjectIdCollection polyLineColl = UserInput.GSelectionSet("Chọn các đối tượng (Polyline/Hatch) cần tính tổng độ dài: \n");
                List<String> listLengthPolyline = [];

                //tạo field data
                string str1 = "%<\\AcObjProp Object(%<\\_ObjId ";
                string str2 = ">%).Length \\f \"%lu2\">%";
                string format = "";
                foreach (ObjectId polyLineId in polyLineColl)
                {
                    string strId = polyLineId.OldIdPtr.ToString();
                    format = str1 + strId + str2;
                    listLengthPolyline.Add(format);
                }

                //tính tổng data file
                string formatTong = "0";
                string str3 = "%<\\AcExpr (";
                string str4 = ") \\f \" %lu2\">%";
                foreach (String item in listLengthPolyline)
                {
                    formatTong = formatTong + "+" + item;
                }
                formatTong = str3 + formatTong + str4;
                // vẽ text
                ObjectId textId = UserInput.GDbText("Chọn text cần nhập nội dung: \n");
                DBText? text = tr.GetObject(textId, OpenMode.ForWrite) as DBText;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                text.TextString = formatTong;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                A.Ed.Command("_UPDATEFIELD", textId);

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_TongDoDai_Replace_CongThem")]
        public static void AT_TongDoDai_Replace_CongThem()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                ObjectIdCollection polyLineColl = UserInput.GSelectionSet("Chọn các đối tượng (Polyline/Hatch) cần tính tổng độ dài: \n");
                List<String> listLengthPolyline = [];

                // vẽ text
                ObjectId textId = UserInput.GDbText("Chọn text cần nhập nội dung: \n");
                DBText? text = tr.GetObject(textId, OpenMode.ForWrite) as DBText;

                //tạo field data
                string str1 = "%<\\AcObjProp Object(%<\\_ObjId ";
                string str2 = ">%).Length \\f \"%lu2\">%";
                string format = "";
                foreach (ObjectId polyLineId in polyLineColl)
                {
                    string strId = polyLineId.OldIdPtr.ToString();
                    format = str1 + strId + str2;
                    listLengthPolyline.Add(format);
                }

                //tính tổng data file
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                string formatTong = text.TextString;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                string str3 = "%<\\AcExpr (";
                string str4 = ") \\f \" %lu2\">%";
                foreach (String item in listLengthPolyline)
                {
                    formatTong = formatTong + "+" + item;
                }
                formatTong = str3 + formatTong + str4;

                text.TextString = formatTong;
                A.Ed.Command("_UPDATEFIELD", textId);

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_TongDienTich_Full")]
        public static void ET_TongDienTich_Full()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                ObjectIdCollection polyLineColl = UserInput.GSelectionSet("Chọn các đối tượng (Polyline/Hatch) cần tính tổng diện tích: \n");
                List<String> listLengthPolyline = [];

                //tạo field data
                string str1 = "%<\\AcObjProp Object(%<\\_ObjId ";
                string str2 = ">%).Area \\f \"%lu2\">%";
                string format = "";
                foreach (ObjectId polyLineId in polyLineColl)
                {
                    string strId = polyLineId.OldIdPtr.ToString();
                    format = str1 + strId + str2;
                    listLengthPolyline.Add(format);
                }

                //tính tổng data file
                string formatTong = "0";
                string str3 = "%<\\AcExpr (";
                string str4 = ") \\f \" %lu2\">%";
                foreach (String item in listLengthPolyline)
                {
                    formatTong = formatTong + "+" + item;
                }
                String formatTong_Lite = formatTong[2..];
                formatTong = formatTong_Lite + "=" + str3 + formatTong + str4 + "m2";
                // vẽ text
                MText text = new();
                Point3d point = UserInput.GPoint("Chọn vị trí đặt text: \n");
                LayerTableRecord layer = UtilitiesCAD.CCreateLayer("TinhTong");
                UtilitiesCAD.CCreateText(point, 2, formatTong, "TinhTong", "Standard");


                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_TongDienTich_Replace")]
        public static void AT_TongDienTich_Replace()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                ObjectIdCollection polyLineColl = UserInput.GSelectionSet("Chọn các đối tượng (Polyline/Hatch) cần tính tổng diện tích: \n");
                List<String> listLengthPolyline = [];

                //tạo field data
                string str1 = "%<\\AcObjProp Object(%<\\_ObjId ";
                string str2 = ">%).Area \\f \"%lu2\">%";
                string format = "";
                foreach (ObjectId polyLineId in polyLineColl)
                {
                    string strId = polyLineId.OldIdPtr.ToString();
                    format = str1 + strId + str2;
                    listLengthPolyline.Add(format);
                }

                //tính tổng data file
                string formatTong = "0";
                string str3 = "%<\\AcExpr (";
                string str4 = ") \\f \" %lu2\">%";
                foreach (String item in listLengthPolyline)
                {
                    formatTong = formatTong + "+" + item;
                }
                formatTong = str3 + formatTong + str4 + "m2";
                // vẽ text
                ObjectId textId = UserInput.GDbText("Chọn text cần nhập nội dung: \n");
                DBText? text = tr.GetObject(textId, OpenMode.ForWrite) as DBText;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                text.TextString = formatTong;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                A.Ed.Command("_UPDATEFIELD", textId);

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_TongDienTich_Replace2")]
        public static void AT_TongDienTich_Replace2()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                ObjectIdCollection polyLineColl = UserInput.GSelectionSet("Chọn các đối tượng (Polyline/Hatch) cần tính tổng diện tích: \n");
                List<String> listLengthPolyline = [];

                //tạo field data
                string str1 = "%<\\AcObjProp Object(%<\\_ObjId ";
                string str2 = ">%).Area \\f \"%lu2\">%";
                string format = "";
                foreach (ObjectId polyLineId in polyLineColl)
                {
                    string strId = polyLineId.OldIdPtr.ToString();
                    format = str1 + strId + str2;
                    listLengthPolyline.Add(format);
                }

                //tính tổng data file
                string formatTong = "0";
                string str3 = "%<\\AcExpr (";
                string str4 = ") \\f \" %lu2\">%";
                foreach (String item in listLengthPolyline)
                {
                    formatTong = formatTong + "+" + item;
                }
                formatTong = str3 + formatTong + str4;
                // vẽ text
                ObjectId textId = UserInput.GDbText("Chọn text cần nhập nội dung: \n");
                DBText? text = tr.GetObject(textId, OpenMode.ForWrite) as DBText;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                text.TextString = formatTong;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                A.Ed.Command("_UPDATEFIELD", textId);

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_TongDienTich_Replace_CongThem")]
        public static void AT_TongDienTich_Replace_CongThem()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                ObjectIdCollection polyLineColl = UserInput.GSelectionSet("Chọn các đối tượng (Polyline/Hatch) cần tính tổng diện tích: \n");
                List<String> listLengthPolyline = [];
                // vẽ text
                ObjectId textId = UserInput.GDbText("Chọn text cần nhập nội dung: \n");
                DBText? text = tr.GetObject(textId, OpenMode.ForWrite) as DBText;

                //tạo field data
                string str1 = "%<\\AcObjProp Object(%<\\_ObjId ";
                string str2 = ">%).Area \\f \"%lu2\">%";
                string format = "";
                foreach (ObjectId polyLineId in polyLineColl)
                {
                    string strId = polyLineId.OldIdPtr.ToString();
                    format = str1 + strId + str2;
                    listLengthPolyline.Add(format);
                }

                //tính tổng data file
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                string formatTong = text.TextString;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                string str3 = "%<\\AcExpr (";
                string str4 = ") \\f \" %lu2\">%";
                foreach (String item in listLengthPolyline)
                {
                    formatTong = formatTong + "+" + item;
                }
                formatTong = str3 + formatTong + str4;

                text.TextString = formatTong;
                A.Ed.Command("_UPDATEFIELD", textId);

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_TextLink")]
        public static void AT_TextLink()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                // Open the Block table for read
                BlockTable? acBlkTbl;
                acBlkTbl = tr.GetObject(A.Db.BlockTableId, OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord? acBlkTblRec;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                //start here
                ObjectId textId = UserInput.GSelectionAnObject("Chọn 1 đối tượng TEXT để lấy nội dung: \n");

                DBText textContest = new();

                if (textId.ObjectClass.DxfName == "TEXT")
                {
                    DBText? textReal = tr.GetObject(textId, OpenMode.ForWrite) as DBText;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    textReal.ColorIndex = 1;
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                }

                else if (textId.ObjectClass.DxfName == "MTEXT")
                {
                    MText? textReal = tr.GetObject(textId, OpenMode.ForWrite) as MText;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    textReal.ColorIndex = 1;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                }

                //đổi text
                ObjectIdCollection text2IdColl = UserInput.GSelectionSet("Chọn các đối tượng TEXT để link nội dung: \n");
                foreach (ObjectId text2Id in text2IdColl)
                {
                    var text2 = tr.GetObject(text2Id, OpenMode.ForWrite);
                    if (text2 is DBText text)
                    {
                        text.TextString = UtilitiesCAD.ConvertMTextToField(textId);
                    }

                    else if (text2 is MText mText)
                    {
                        mText.Contents = UtilitiesCAD.ConvertMTextToField(textId);
                        A.Ed.WriteMessage(text2.ToString());
                    }


                }

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_DanhSoThuTu")]
        public static void ET_DanhSoThuTu()
        {
            // start transantion
            _ = new            // start transantion
            UserInput();
            _ = new UtilitiesCAD();
            _ = new UtilitiesC3D();
            //start here
            int soThuTu = UserInput.GInt("\n Nhập số thứ tự cần đánh số: ");
            String ans = "Enter";
            while (ans == "Enter")
            {
                using (Transaction tr = A.Db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        //đổi text
                        ObjectId textId = UserInput.GSelectionAnObject("\n Chọn các đối tượng TEXT để đánh số thứ tự:");
                        var text = tr.GetObject(textId, OpenMode.ForWrite);
                        if (text is DBText dbText)
                            dbText.TextString = soThuTu.ToString();
                        else if (text is MText mText)
                            mText.Contents = soThuTu.ToString();
                        else soThuTu--;

                        soThuTu++;

                        tr.Commit();
                    }

                    catch (Autodesk.AutoCAD.Runtime.Exception e)
                    {
                        A.Ed.WriteMessage(e.Message);
                    }

                }
                ans = UserInput.GStopWithESC();
            }
        }
        
        [CommandMethod("AT_XoayDoiTuong_TheoViewport")]
        public static void ET_XoayDoiTuong_TheoViewport()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                // Open the Block table for read
                BlockTable? acBlkTbl;
                acBlkTbl = tr.GetObject(A.Db.BlockTableId, OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord? acBlkTblRec;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                // get current viewport
                ViewTableRecord acView = A.Ed.GetCurrentView();
                Double angleAcView = acView.ViewTwist;

                ObjectIdCollection noteIdColl = UserInput.GSelectionSet("\n Chọn các note label cần xoay theo viewport:");
                foreach (ObjectId noteId in noteIdColl)
                {
                    var noteLabel = tr.GetObject(noteId, OpenMode.ForWrite);
                    if (noteLabel is DBText text)
                        text.Rotation = 3.14159 - angleAcView + 3.14159;
                    else if (noteLabel is MText text1)
                        text1.Rotation = 3.14159 - angleAcView + 3.14159;
                    else if (noteLabel is NoteLabel label)
                        label.RotationAngle = 3.14159 - angleAcView;
                    else if (noteLabel is BlockReference reference)
                        reference.Rotation = 3.14159 - angleAcView + 3.14159;
                }



                // Add each offset object
                //acBlkTblRec.AppendEntity(acEnt);
                //tr.AddNewlyCreatedDBObject(acEnt, true);

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_XoayDoiTuong_Theo2Diem")]
        public static void AT_XoayDoiTuong_Theo2Diem()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                // Open the Block table for read
                BlockTable? acBlkTbl;
                acBlkTbl = tr.GetObject(A.Db.BlockTableId, OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord? acBlkTblRec;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                // get current viewport
                ViewTableRecord acView = A.Ed.GetCurrentView();
                //Double angleAcView = acView.ViewTwist;

                ObjectIdCollection noteIdColl = UserInput.GSelectionSet("\n Chọn các note label cần xoay theo viewport:");
                Point3d point1 = UserInput.GPoint("Chọn vị trí điểm ĐẦU làm căn cứ xoay đối tượng:");
                Point3d point2 = UserInput.GPoint("Chọn vị trí điểm CUỐI làm căn cứ xoay đối tượng:");

                // góc giữa 2 điểm
                Double angleAcView = Math.Atan2(point1.Y - point2.Y, point2.X - point1.X);
                foreach (ObjectId noteId in noteIdColl)
                {
                    var noteLabel = tr.GetObject(noteId, OpenMode.ForWrite);
                    if (noteLabel is DBText text)
                        text.Rotation = 3.14159 - angleAcView + 3.14159;
                    else if (noteLabel is MText text1)
                        text1.Rotation = 3.14159 - angleAcView + 3.14159;
                    else if (noteLabel is NoteLabel label)
                        label.RotationAngle = 3.14159 - angleAcView;
                    else if (noteLabel is BlockReference reference)
                        reference.Rotation = 3.14159 - angleAcView + 3.14159;
                }



                // Add each offset object
                //acBlkTblRec.AppendEntity(acEnt);
                //tr.AddNewlyCreatedDBObject(acEnt, true);

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_TextLayout")]
        public static void AT_TextLayout()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                double textHight = UserInput.GDouble("Nhập chiều cao text muốm in ra (2, 3, 4, 5):");
                ObjectIdCollection textIdColl = UserInput.GSelectionSet("Chọn các đối tượng để điều chỉnh tỉ lệ: \n");
                UtilitiesCAD.TextLayout(textIdColl, textHight);

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_TaoMoi_TextLayout")]
        public static void ET_TaoMoi_TextLayout()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here

                //get current sacle annotation
                AnnotationScale annoScaleCurrent = A.Db.Cannoscale;
                A.Ok(annoScaleCurrent.DrawingUnits.ToString() + "\n");
                A.Ok(annoScaleCurrent.PaperUnits.ToString() + "\n");
                double textHeight = UserInput.GDouble("\n Nhập chiều cao text:");
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                String textContain = UserInput.GString("\n Nhập nội dung text:");
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
                Point3d point3D = UserInput.GPoint("\n Chọn vị trí đặt text:");
                UtilitiesCAD.CCreateLayer("0.TEXT");
                UtilitiesCAD.CCreateLayer("0.TEXT_TO");
                String layerName = "0.TEXT";

                if (textHeight > 2)
                {
                    layerName = "0.TEXT_TO";
                }
#pragma warning disable CS8604 // Possible null reference argument.
                UtilitiesCAD.CCreateText(point3D, textHeight * annoScaleCurrent.DrawingUnits, textContain, layerName, "Standard");
#pragma warning restore CS8604 // Possible null reference argument.

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_DimLayout2")]
        public static void ET_DimLayout2()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here

                //get current sacle annotation
                AnnotationScale annoScaleCurrent = A.Db.Cannoscale;
                A.Ok(annoScaleCurrent.DrawingUnits.ToString() + "\n");
                A.Ok(annoScaleCurrent.PaperUnits.ToString() + "\n");
                UtilitiesCAD.CCreateLayer("1.DIM");
                UtilitiesCAD.CreateDimStyle("1-1000");

                //đổi text
                ObjectIdCollection textIdColl = UserInput.GSelectionSet("Chọn các đối tượng để điều chỉnh tỉ lệ: \n");

                foreach (ObjectId dimensionId in textIdColl)
                {

                    if (dimensionId.ObjectClass.Name == "AcDbAlignedDimension")
                    {
                        AlignedDimension? dimension = tr.GetObject(dimensionId, OpenMode.ForWrite) as AlignedDimension;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        dimension.Dimscale = annoScaleCurrent.DrawingUnits;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        dimension.Layer = "1.DIM";
                        dimension.DimensionStyleName = "1-1000";
                        dimension.Dimscale = annoScaleCurrent.DrawingUnits;
                    }

                    else if (dimensionId.ObjectClass.Name == "AcDb2LineAngularDimension")
                    {
                        LineAngularDimension2? dimension = tr.GetObject(dimensionId, OpenMode.ForWrite) as LineAngularDimension2;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        dimension.Dimscale = annoScaleCurrent.DrawingUnits;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        dimension.Layer = "1.DIM";
                        dimension.DimensionStyleName = "1-1000";
                        dimension.Dimscale = annoScaleCurrent.DrawingUnits;
                    }
                    else if (dimensionId.ObjectClass.Name == "AcDbRotatedDimension")
                    {
                        RotatedDimension? dimension = tr.GetObject(dimensionId, OpenMode.ForWrite) as RotatedDimension;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        dimension.Dimscale = annoScaleCurrent.DrawingUnits;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        dimension.Layer = "1.DIM";
                        dimension.DimensionStyleName = "1-1000";
                        dimension.Dimscale = annoScaleCurrent.DrawingUnits;
                    }
                    else if (dimensionId.ObjectClass.Name == "AcDbArcDimension")
                    {
                        ArcDimension? dimension = tr.GetObject(dimensionId, OpenMode.ForWrite) as ArcDimension;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        dimension.Dimscale = annoScaleCurrent.DrawingUnits;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        dimension.Layer = "1.DIM";
                        dimension.DimensionStyleName = "1-1000";
                        dimension.Dimscale = annoScaleCurrent.DrawingUnits;
                    }
                    else if (dimensionId.ObjectClass.Name == "AcDbRadialDimension")
                    {
                        RadialDimension? dimension = tr.GetObject(dimensionId, OpenMode.ForWrite) as RadialDimension;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        dimension.Dimscale = annoScaleCurrent.DrawingUnits;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        dimension.Layer = "1.DIM";
                        dimension.DimensionStyleName = "1-1000";
                        dimension.Dimscale = annoScaleCurrent.DrawingUnits;
                    }
                    else if (dimensionId.ObjectClass.Name == "AcDbDiametricDimension")
                    {
                        DiametricDimension? dimension = tr.GetObject(dimensionId, OpenMode.ForWrite) as DiametricDimension;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        dimension.Dimscale = annoScaleCurrent.DrawingUnits;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        dimension.Layer = "1.DIM";
                        dimension.DimensionStyleName = "1-1000";
                        dimension.Dimscale = annoScaleCurrent.DrawingUnits;
                    }
                    else if (dimensionId.ObjectClass.Name == "AcDbLeader")
                    {
                        Leader? dimension = tr.GetObject(dimensionId, OpenMode.ForWrite) as Leader;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        dimension.Dimscale = annoScaleCurrent.DrawingUnits;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        dimension.Layer = "1.DIM";
                        dimension.DimensionStyleName = "1-1000";
                        dimension.Dimscale = annoScaleCurrent.DrawingUnits;
                    }

                }


                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_DimLayout")]
        public static void ET_DimLayout()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                ObjectIdCollection ObjectIdColl = UserInput.GSelectionSet("Chọn các đối tượng để UPDATE tỉ lệ: \n");

                //get current sacle annotation 
                UtilitiesCAD.DimLayout(ObjectIdColl);

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_BlockLayout")]
        public static void ET_BlockLayout()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here

                //get current sacle annotation
                AnnotationScale annoScaleCurrent = A.Db.Cannoscale;

                //đổi text
                ObjectIdCollection blockIdColl = UserInput.GSelectionSet("Chọn các đối tượng để điều chỉnh tỉ lệ: \n");
                UtilitiesCAD.BlockLayout(blockIdColl);

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_Label_FromText")]
        public static void AT_Label_FromText()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                ObjectId labelStyleNoteId = A.Cdoc.Styles.LabelStyles.GeneralNoteLabelStyles["CAD - GHI CHU 2MM"];
                ObjectId markerStyleId = A.Cdoc.Styles.MarkerStyles["_No Markers"];
                //start here
                ObjectIdCollection textIdColl = UserInput.GSelectionSetWithType("\n Chọn các đối tượng text cần tạo cogo point: ", "TEXT");
                foreach (ObjectId textId in textIdColl)
                {
                    DBText? text = tr.GetObject(textId, OpenMode.ForWrite) as DBText;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Point3d alignmentPoint = text.AlignmentPoint;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    Point3d position = text.Position;


                    //set position và text left
                    if ((text.Justify == AttachmentPoint.BaseLeft) || (text.Justify == AttachmentPoint.BaseAlign) || (text.Justify == AttachmentPoint.BaseFit))
                    {

                        text.Justify = AttachmentPoint.BaseLeft;
                    }

                    else
                    {
                        text.Position = alignmentPoint;
                        text.Justify = AttachmentPoint.BaseLeft;
                    }

                    //lấy tọa độ cogopoint
                    Point3d cogoText = text.Position;

                    //tạo note label
                    ObjectId noteLabelId = NoteLabel.Create(cogoText, labelStyleNoteId, markerStyleId);
                    NoteLabel? noteLabel = tr.GetObject(noteLabelId, OpenMode.ForWrite) as NoteLabel;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    noteLabel.SetTextComponentOverride(noteLabel.GetTextComponentIds()[0], text.TextString);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                    text.Erase();



                }

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_XoaDoiTuong_CungLayer")]
        public static void AT_XoaDoiTuong_CungLayer()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            UserInput UI = new();
            UtilitiesCAD CAD = new();
            UtilitiesC3D C3D = new();
            try
            {
                ObjectId objectId = UserInput.GObjId("/n Chọn đối tượng (TEXT, MTEXT, POLYLINE, LINE) nằm trên layer cần xóa: \n");
                if (objectId.ObjectClass.DxfName == "TEXT")
                {
                    DBText? objectLayer = tr.GetObject(objectId, OpenMode.ForWrite) as DBText;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    String layerName = objectLayer.Layer;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    UtilitiesCAD.CDelLayerAndObjectOnIt(layerName);
                }
                if (objectId.ObjectClass.DxfName == "MTEXT")
                {
                    MText? objectLayer = tr.GetObject(objectId, OpenMode.ForWrite) as MText;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    String layerName = objectLayer.Layer;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    UtilitiesCAD.CDelLayerAndObjectOnIt(layerName);
                }
                if (objectId.ObjectClass.DxfName == "LWPOLYLINE")
                {
                    Polyline? objectLayer = tr.GetObject(objectId, OpenMode.ForWrite) as Polyline;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    String layerName = objectLayer.Layer;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    UtilitiesCAD.CDelLayerAndObjectOnIt(layerName);
                }
                if (objectId.ObjectClass.DxfName == "LINE")
                {
                    Line? objectLayer = tr.GetObject(objectId, OpenMode.ForWrite) as Line;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    String layerName = objectLayer.Layer;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    UtilitiesCAD.CDelLayerAndObjectOnIt(layerName);
                }


                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_XoaDoiTuong_3DSolid_Body")]
        public static void AT_XoaDoiTuong_3DSolid_Body()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            UserInput UI = new();
            UtilitiesCAD CAD = new();
            UtilitiesC3D C3D = new();
            try
            {
                BlockTable bt = (BlockTable)tr.GetObject(A.Doc.Database.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                foreach (ObjectId objId in btr)
                {
                    Entity ent = (Entity)tr.GetObject(objId, OpenMode.ForRead);
                    if (ent is Solid3d)
                    {
                        ent.UpgradeOpen();
                        ent.Erase();
                    }
                }

                foreach (ObjectId objId in btr)
                {
                    Entity ent = (Entity)tr.GetObject(objId, OpenMode.ForRead);
                    if (ent is Body)
                    {
                        ent.UpgradeOpen();
                        ent.Erase();
                    }
                }



                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_UpdateLayout")]
        public static void AT_UpdateLayout() 
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();
                //start here
                ObjectIdCollection ObjectIdColl = UserInput.GSelectionSet("Chọn các đối tượng để UPDATE tỉ lệ: \n");

                //get current sacle annotation                    
                UtilitiesCAD.TextLayout(ObjectIdColl, 2);
                UtilitiesCAD.BlockLayout(ObjectIdColl);
                UtilitiesCAD.DimLayout(ObjectIdColl);
                foreach (ObjectId objectId in ObjectIdColl)
                {
                    UtilitiesCAD.LinePolyLineHatchArraytArcSplineLayout(objectId);
                }


                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }

        [CommandMethod("AT_Offset_2Ben")]
        public static void AT_Offset_2Ben()
        {
            // Lấy tài liệu và trình chỉnh sửa hiện tại
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // Yêu cầu người dùng chọn nhiều đối tượng
                PromptSelectionOptions selOptions = new()
                {
                    MessageForAdding = "\nChọn các đối tượng để offset: ",
                    AllowDuplicates = false
                };

                PromptSelectionResult selResult = ed.GetSelection(selOptions);
                if (selResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nĐã hủy chọn đối tượng.");
                    return;
                }

                SelectionSet selSet = selResult.Value;

                // Yêu cầu người dùng nhập khoảng cách offset
                PromptDoubleOptions distOptions = new("\nNhập khoảng cách offset: ")
                {
                    AllowNegative = false,
                    AllowZero = false,
                    DefaultValue = 1.0
                };

                PromptDoubleResult distResult = ed.GetDouble(distOptions);
                if (distResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nĐã hủy nhập khoảng cách.");
                    return;
                }

                double offsetDist = distResult.Value;

                Database db = doc.Database;

                // Bắt đầu giao dịch
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    foreach (SelectedObject selObj in selSet)
                    {
                        if (selObj != null)
                        {
                            // Lấy đối tượng và kiểm tra loại của nó
                            Entity? ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                            if (ent != null)
                            {
                                DBObjectCollection? offsetCurves1 = null;
                                DBObjectCollection? offsetCurves2 = null;

                                if (ent is Curve curve)
                                {
                                    // Offset đối với các đối tượng Curve
                                    offsetCurves1 = curve.GetOffsetCurves(offsetDist);
                                    offsetCurves2 = curve.GetOffsetCurves(-offsetDist);
                                }
                                else if (ent is Polyline2d poly2d)
                                {
                                    // Offset đối với Polyline2d
                                    offsetCurves1 = poly2d.GetOffsetCurves(offsetDist);
                                    offsetCurves2 = poly2d.GetOffsetCurves(-offsetDist);
                                }
                                else if (ent is Polyline3d)
                                {
                                    ed.WriteMessage("\nKhông hỗ trợ offset Polyline3d.");
                                    continue;
                                }
                                else
                                {
                                    ed.WriteMessage("\nĐối tượng được chọn không thể offset.");
                                    continue;
                                }

                                // Thêm các đối tượng offset vào bản vẽ
                                if (offsetCurves1 != null)
                                {
                                    foreach (Entity offsetEnt in offsetCurves1)
                                    {
                                        btr.AppendEntity(offsetEnt);
                                        tr.AddNewlyCreatedDBObject(offsetEnt, true);
                                    }
                                }

                                if (offsetCurves2 != null)
                                {
                                    foreach (Entity offsetEnt in offsetCurves2)
                                    {
                                        btr.AppendEntity(offsetEnt);
                                        tr.AddNewlyCreatedDBObject(offsetEnt, true);
                                    }
                                }
                            }
                        }
                    }

                    // Xác nhận giao dịch
                    tr.Commit();
                }

                ed.WriteMessage("\nHoàn thành offset.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nLỗi: " + ex.Message);
            }
        }

        [CommandMethod("AT_annotive_scale_currentOnly")]
        public static void AT_annotive_scale_currentOnly()
        {
            // start transaction
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput uI = new();
                UtilitiesCAD uti = new();
                UtilitiesC3D C3D = new();
                // start here

                //get a selection set with annotative objects
                ObjectIdCollection ss = UserInput.GSelectionSet("\n Chọn Annotative Object cần thêm scale: ");

                // Lấy AnnotationScale của bản hiện hành
                AnnotationScale currentScaleId = A.Db.Cannoscale;

                // Lấy tập ObjectContextCollection của Annotation Scales
                ObjectContextManager ocm = A.Db.ObjectContextManager;
                ObjectContextCollection occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");

                foreach (ObjectId id in ss)
                {
                    Entity ent = (Entity)tr.GetObject(id, OpenMode.ForWrite);
                    if (ent.Annotative == AnnotativeStates.True)
                    {

                        ent.AddContext(currentScaleId);

                        foreach (ObjectContext ctx in occ)
                        {
                            AnnotationScale? annoScale = ctx as AnnotationScale;
                            if (annoScale != currentScaleId)
                            {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                                A.Ed.WriteMessage("\nScale: " + annoScale.Name);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                                ent.AddContext(annoScale);
                                ent.RemoveContext(annoScale);
                            }
                        }

                        ent.AddContext(currentScaleId);

                    }

                }







                // commit transaction
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage("\n" + ex.Message);
            }
        }
















    }
}
