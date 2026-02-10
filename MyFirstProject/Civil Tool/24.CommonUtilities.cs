// CommonUtilities.cs - Các tiện ích CAD thường dùng
// Chuyển đổi từ LISP: MP, TONG, EXPTXT, TLP, T2M, APV, INTLINES

using System;
using System.IO;
using System.Collections.Generic;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.CommonUtilitiesCommands))]

namespace MyFirstProject
{
    /// <summary>
    /// Các lệnh tiện ích CAD thường dùng
    /// </summary>
    public class CommonUtilitiesCommands
    {
        // ══════════════════════════════════════════════════════════════
        // TẠO POINT TỪ TEXT (từ LISP MP)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// Tạo các Point 3D từ Text, với cao độ Z = nội dung text (số).
        /// </summary>
        [CommandMethod("CTU_MakePointFromText")]
        public static void CTU_MakePointFromText()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var db = doc.Database;

            // Chọn các Text
            var filter = new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Start, "TEXT")
            });

            var psr = ed.GetSelection(new PromptSelectionOptions
            {
                MessageForAdding = "\n⊙ Chọn các Text chứa cao độ: "
            }, filter);

            if (psr.Status != PromptStatus.OK) return;

            int count = 0;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                foreach (SelectedObject so in psr.Value)
                {
                    var text = tr.GetObject(so.ObjectId, OpenMode.ForRead) as DBText;
                    if (text == null) continue;

                    // Lấy vị trí X, Y từ text
                    Point3d insertPt = text.Position;
                    
                    // Lấy Z từ nội dung text
                    if (double.TryParse(text.TextString.Trim(), out double z))
                    {
                        var point = new DBPoint(new Point3d(insertPt.X, insertPt.Y, z));
                        btr.AppendEntity(point);
                        tr.AddNewlyCreatedDBObject(point, true);
                        count++;
                    }
                }

                tr.Commit();
            }

            ed.WriteMessage($"\n◎ Đã tạo {count} Point từ Text.");
        }

        // ══════════════════════════════════════════════════════════════
        // TÍNH TỔNG CHIỀU DÀI (từ LISP TONG)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// Tính tổng chiều dài của các đối tượng đường (Line, Arc, Polyline, Circle, Spline, Ellipse).
        /// </summary>
        [CommandMethod("CTU_TotalLength")]
        public static void CTU_TotalLength()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var db = doc.Database;

            // Chọn các đối tượng đường
            var filter = new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Start, "LINE,ARC,CIRCLE,POLYLINE,LWPOLYLINE,ELLIPSE,SPLINE")
            });

            var psr = ed.GetSelection(new PromptSelectionOptions
            {
                MessageForAdding = "\n⊙ Chọn các đối tượng để tính tổng chiều dài: "
            }, filter);

            if (psr.Status != PromptStatus.OK) return;

            double totalLength = 0.0;
            int count = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject so in psr.Value)
                {
                    var curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                    if (curve != null)
                    {
                        totalLength += curve.GetDistanceAtParameter(curve.EndParam);
                        count++;
                    }
                }
                tr.Commit();
            }

            ed.WriteMessage($"\n═══════════════════════════════════════════════════════════════");
            ed.WriteMessage($"\n◎ Số đối tượng: {count}");
            ed.WriteMessage($"\n◎ TỔNG CHIỀU DÀI = {totalLength:F3}");
            ed.WriteMessage($"\n═══════════════════════════════════════════════════════════════");

            // Hiển thị hộp thoại thông báo
            System.Windows.Forms.MessageBox.Show(
                $"Số đối tượng: {count}\nTổng chiều dài: {totalLength:F3}",
                "Tổng chiều dài",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information
            );
        }

        // ══════════════════════════════════════════════════════════════
        // XUẤT TỌA ĐỘ TEXT RA FILE (từ LISP EXPTXT)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// Xuất tọa độ và nội dung Text ra file .txt (X, Y, Nội dung).
        /// </summary>
        [CommandMethod("CTU_ExportTextCoords")]
        public static void CTU_ExportTextCoords()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var db = doc.Database;

            // Chọn các Text
            var filter = new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Start, "TEXT")
            });

            var psr = ed.GetSelection(new PromptSelectionOptions
            {
                MessageForAdding = "\n⊙ Chọn các Text để xuất tọa độ: "
            }, filter);

            if (psr.Status != PromptStatus.OK) return;

            // Chọn file xuất
            var saveDialog = new System.Windows.Forms.SaveFileDialog
            {
                Title = "Lưu file tọa độ",
                Filter = "Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                DefaultExt = "txt",
                FileName = "coordinates"
            };

            if (saveDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            int count = 0;
            using (var writer = new StreamWriter(saveDialog.FileName))
            {
                // Header
                writer.WriteLine("X\tY\tContent");

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject so in psr.Value)
                    {
                        var text = tr.GetObject(so.ObjectId, OpenMode.ForRead) as DBText;
                        if (text == null) continue;

                        Point3d pt = text.Position;
                        writer.WriteLine($"{pt.X:F3}\t{pt.Y:F3}\t{text.TextString}");
                        count++;
                    }
                    tr.Commit();
                }
            }

            ed.WriteMessage($"\n◎ Đã xuất {count} tọa độ Text ra file: {saveDialog.FileName}");
        }

        // ══════════════════════════════════════════════════════════════
        // CONVERT TEXT TO MTEXT (từ LISP T2M)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// Chuyển đổi nhiều Text thành MText.
        /// </summary>
        [CommandMethod("CTU_TextToMText")]
        public static void CTU_TextToMText()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            // Chọn các Text
            var filter = new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Start, "TEXT")
            });

            var psr = ed.GetSelection(new PromptSelectionOptions
            {
                MessageForAdding = "\n⊙ Chọn các Text để chuyển sang MText: "
            }, filter);

            if (psr.Status != PromptStatus.OK) return;

            ed.WriteMessage($"\n▸ Đang chuyển đổi {psr.Value.Count} Text sang MText...");

            // Sử dụng lệnh Express Tools TXT2MTXT
            foreach (SelectedObject so in psr.Value)
            {
                doc.SendStringToExecute($"TXT2MTXT (handent \"{so.ObjectId.Handle}\") ", true, false, false);
            }
        }

        // ══════════════════════════════════════════════════════════════
        // TÌM ĐIỂM GIAO CÁC ĐỐI TƯỢNG (từ LISP INTLINES)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// Tìm và đánh dấu điểm giao của các đối tượng đường.
        /// </summary>
        [CommandMethod("CTU_FindIntersections")]
        public static void CTU_FindIntersections()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var db = doc.Database;

            // Chọn các đối tượng
            var filter = new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Start, "LINE,ARC,CIRCLE,POLYLINE,LWPOLYLINE,ELLIPSE,SPLINE")
            });

            var psr = ed.GetSelection(new PromptSelectionOptions
            {
                MessageForAdding = "\n⊙ Chọn các đối tượng để tìm điểm giao: "
            }, filter);

            if (psr.Status != PromptStatus.OK || psr.Value.Count < 2)
            {
                ed.WriteMessage("\n⊘ Cần chọn ít nhất 2 đối tượng!");
                return;
            }

            var curves = new List<Curve>();
            var intersectionPoints = new List<Point3d>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Thu thập các curve
                foreach (SelectedObject so in psr.Value)
                {
                    var curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                    if (curve != null)
                    {
                        curves.Add(curve);
                    }
                }

                // Tìm điểm giao giữa từng cặp curve
                for (int i = 0; i < curves.Count - 1; i++)
                {
                    for (int j = i + 1; j < curves.Count; j++)
                    {
                        var pts = new Point3dCollection();
                        curves[i].IntersectWith(curves[j], Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero);

                        foreach (Point3d pt in pts)
                        {
                            intersectionPoints.Add(pt);
                        }
                    }
                }

                // Tạo Point tại các điểm giao
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                foreach (Point3d pt in intersectionPoints)
                {
                    var point = new DBPoint(pt);
                    point.ColorIndex = 1; // Red
                    btr.AppendEntity(point);
                    tr.AddNewlyCreatedDBObject(point, true);
                }

                tr.Commit();
            }

            // Đặt PDMODE để hiển thị point rõ hơn
            AcadApp.SetSystemVariable("PDMODE", 35);
            AcadApp.SetSystemVariable("PDSIZE", -2);

            ed.WriteMessage($"\n◎ Đã tìm thấy {intersectionPoints.Count} điểm giao và đánh dấu bằng Point.");
            ed.Regen();
        }

        // ══════════════════════════════════════════════════════════════
        // THÊM ĐỈNH CHO POLYLINE (từ LISP APV)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// Thêm vertex (đỉnh) cho Polyline tại các điểm chọn.
        /// </summary>
        [CommandMethod("CTU_AddPolylineVertices")]
        public static void CTU_AddPolylineVertices()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var db = doc.Database;

            // Chọn Polyline
            var peo = new PromptEntityOptions("\n⊙ Chọn Polyline để thêm đỉnh: ");
            peo.SetRejectMessage("\n⊘ Đây không phải Polyline!");
            peo.AddAllowedClass(typeof(Polyline), true);

            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var pline = tr.GetObject(per.ObjectId, OpenMode.ForWrite) as Polyline;
                if (pline == null) return;

                int addedCount = 0;

                while (true)
                {
                    var ppo = new PromptPointOptions($"\n▸ Chọn điểm để thêm vertex #{addedCount + 1} (Enter để kết thúc): ");
                    ppo.AllowNone = true;

                    var ppr = ed.GetPoint(ppo);
                    if (ppr.Status == PromptStatus.None) break;
                    if (ppr.Status != PromptStatus.OK) continue;

                    Point3d pickPt = ppr.Value;

                    // Tìm điểm gần nhất trên polyline
                    Point3d closestPt = pline.GetClosestPointTo(pickPt, false);
                    double param = pline.GetParameterAtPoint(closestPt);
                    int segmentIndex = (int)Math.Floor(param);

                    // Thêm vertex
                    pline.AddVertexAt(segmentIndex + 1, new Point2d(closestPt.X, closestPt.Y), 0, 0, 0);
                    addedCount++;

                    ed.WriteMessage($"\n  ✓ Đã thêm vertex tại ({closestPt.X:F3}, {closestPt.Y:F3})");
                }

                tr.Commit();
                ed.WriteMessage($"\n◎ Đã thêm {addedCount} vertex vào Polyline.");
            }

            ed.Regen();
        }

        // ══════════════════════════════════════════════════════════════
        // VẼ TALUY (từ LISP TLP)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// Vẽ taluy (ký hiệu mái dốc) dọc theo đường baseline đến đường giao tuyến.
        /// </summary>
        [CommandMethod("CTU_DrawTaluy")]
        public static void CTU_DrawTaluy()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var db = doc.Database;

            // Nhập khoảng cách chia taluy
            var pdo = new PromptDoubleOptions("\n▸ Nhập khoảng cách chia taluy: ");
            pdo.DefaultValue = 1.5;
            pdo.UseDefaultValue = true;
            pdo.AllowNegative = false;
            pdo.AllowZero = false;

            var pdr = ed.GetDouble(pdo);
            if (pdr.Status != PromptStatus.OK) return;

            double spacing = pdr.Value;

            // Chọn đường baseline (đường cần chia taluy)
            var peo = new PromptEntityOptions("\n⊙ Chọn đường baseline (đường cần vẽ taluy): ");
            peo.SetRejectMessage("\n⊘ Đây không phải đối tượng đường!");
            peo.AddAllowedClass(typeof(Curve), false);

            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            // Hỏi hướng taluy
            var pko = new PromptKeywordOptions("\n▸ Chọn hướng taluy [Đào/Đắp]: ", "Dao Dap");
            pko.AllowNone = false;
            var pkr = ed.GetKeywords(pko);
            if (pkr.Status != PromptStatus.OK) return;

            bool isDao = pkr.StringResult == "Dao";

            // Chọn đường giao tuyến (mặt tự nhiên)
            var psr = ed.GetSelection(new PromptSelectionOptions
            {
                MessageForAdding = "\n⊙ Chọn đường giao tuyến (mặt tự nhiên): "
            });
            if (psr.Status != PromptStatus.OK) return;

            ed.WriteMessage($"\n▸ Đang vẽ taluy với khoảng cách {spacing}...");

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var baseline = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Curve;
                if (baseline == null) return;

                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Tạo hoặc lấy layer cho taluy
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                string layerLong = "maidai";
                string layerShort = "maingan";

                if (!lt.Has(layerLong))
                {
                    lt.UpgradeOpen();
                    var layer = new LayerTableRecord { Name = layerLong, Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 5) };
                    lt.Add(layer);
                    tr.AddNewlyCreatedDBObject(layer, true);
                }

                if (!lt.Has(layerShort))
                {
                    lt.UpgradeOpen();
                    var layer = new LayerTableRecord { Name = layerShort, Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 4) };
                    lt.Add(layer);
                    tr.AddNewlyCreatedDBObject(layer, true);
                }

                // Thu thập các curve giao tuyến
                var intersectCurves = new List<Curve>();
                foreach (SelectedObject so in psr.Value)
                {
                    var curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                    if (curve != null)
                    {
                        intersectCurves.Add(curve);
                    }
                }

                // Chia điểm trên baseline
                double totalLength = baseline.GetDistanceAtParameter(baseline.EndParam);
                int numPoints = (int)(totalLength / spacing);

                int lineCount = 0;
                for (int i = 0; i <= numPoints; i++)
                {
                    double dist = i * spacing;
                    if (dist > totalLength) break;

                    Point3d pt = baseline.GetPointAtDist(dist);
                    Vector3d tangent = baseline.GetFirstDerivative(baseline.GetParameterAtPoint(pt)).GetNormal();
                    Vector3d perpendicular = isDao ? tangent.RotateBy(Math.PI / 2, Vector3d.ZAxis) : tangent.RotateBy(-Math.PI / 2, Vector3d.ZAxis);

                    // Tìm điểm giao với các đường giao tuyến
                    Point3d? intersectPt = null;
                    double minDist = double.MaxValue;

                    foreach (var curve in intersectCurves)
                    {
                        // Tạo đường vuông góc dài để tìm giao điểm
                        Point3d farPt = pt + perpendicular * 1000;
                        using (var tempLine = new Line(pt, farPt))
                        {
                            var pts = new Point3dCollection();
                            tempLine.IntersectWith(curve, Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero);

                            foreach (Point3d intPt in pts)
                            {
                                double d = pt.DistanceTo(intPt);
                                if (d < minDist)
                                {
                                    minDist = d;
                                    intersectPt = intPt;
                                }
                            }
                        }
                    }

                    // Vẽ line taluy
                    if (intersectPt.HasValue && minDist < 500)
                    {
                        bool isLong = (i % 2 == 0);
                        Point3d endPt = isLong ? intersectPt.Value : pt + (intersectPt.Value - pt) * 0.5;

                        var line = new Line(pt, endPt);
                        line.Layer = isLong ? layerLong : layerShort;
                        btr.AppendEntity(line);
                        tr.AddNewlyCreatedDBObject(line, true);
                        lineCount++;
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n◎ Đã vẽ {lineCount} đường taluy.");
            }

            ed.Regen();
        }
    }
}
