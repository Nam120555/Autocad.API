// 20.VisualInfraCore.cs - Core Utilities từ VisualINFRA
// Viết lại cho AutoCAD 2026 / Civil 3D 2026

using System;
using System.Collections.Generic;
using System.Linq;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using AcEntity = Autodesk.AutoCAD.DatabaseServices.Entity;

namespace MyFirstProject.VisualINFRA
{
    /// <summary>
    /// Các hàm tiện ích cơ bản - tương đương class Func trong VisualINFRA
    /// </summary>
    public static class VIFunc
    {
        /// <summary>
        /// Chuyển đổi độ sang radian
        /// </summary>
        public static double DegreeToRadian(double degree)
        {
            return degree * Math.PI / 180.0;
        }

        /// <summary>
        /// Chuyển đổi radian sang độ
        /// </summary>
        public static double RadianToDegree(double radian)
        {
            return radian * 180.0 / Math.PI;
        }

        /// <summary>
        /// Lấy điểm từ người dùng
        /// </summary>
        public static Point3d? GetPoint(string prompt = "\nChọn điểm: ")
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            
            var options = new PromptPointOptions(prompt);
            var result = ed.GetPoint(options);
            
            if (result.Status == PromptStatus.OK)
                return result.Value;
            return null;
        }

        /// <summary>
        /// Lấy selection set từ người dùng
        /// </summary>
        public static SelectionSet? GetSSget(string prompt = "\nChọn đối tượng: ")
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            
            ed.WriteMessage(prompt);
            var result = ed.GetSelection();
            
            if (result.Status == PromptStatus.OK)
                return result.Value;
            return null;
        }

        /// <summary>
        /// Tạo layer mới nếu chưa tồn tại
        /// </summary>
        public static ObjectId CreateLayer(string layerName, short colorIndex = 7)
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                
                if (lt != null && !lt.Has(layerName))
                {
                    lt.UpgradeOpen();
                    var layer = new LayerTableRecord
                    {
                        Name = layerName,
                        Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                            Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex)
                    };
                    lt.Add(layer);
                    tr.AddNewlyCreatedDBObject(layer, true);
                    tr.Commit();
                    return layer.ObjectId;
                }
                
                return lt != null ? lt[layerName] : ObjectId.Null;
            }
        }

        /// <summary>
        /// Thêm text vào bản vẽ
        /// </summary>
        public static ObjectId AddText(string content, Point3d position, double height = 2.5, 
            double rotation = 0, string layerName = "0")
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                var btr = tr.GetObject(bt![BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                
                var text = new DBText
                {
                    TextString = content,
                    Position = position,
                    Height = height,
                    Rotation = rotation,
                    Layer = layerName
                };
                
                btr!.AppendEntity(text);
                tr.AddNewlyCreatedDBObject(text, true);
                tr.Commit();
                
                return text.ObjectId;
            }
        }

        /// <summary>
        /// Thêm line vào bản vẽ
        /// </summary>
        public static ObjectId AddLine(Point3d startPoint, Point3d endPoint, string layerName = "0")
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                var btr = tr.GetObject(bt![BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                
                var line = new Line(startPoint, endPoint)
                {
                    Layer = layerName
                };
                
                btr!.AppendEntity(line);
                tr.AddNewlyCreatedDBObject(line, true);
                tr.Commit();
                
                return line.ObjectId;
            }
        }

        /// <summary>
        /// Thêm circle vào bản vẽ
        /// </summary>
        public static ObjectId AddCircle(Point3d center, double radius, string layerName = "0")
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                var btr = tr.GetObject(bt![BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                
                var circle = new Circle(center, Vector3d.ZAxis, radius)
                {
                    Layer = layerName
                };
                
                btr!.AppendEntity(circle);
                tr.AddNewlyCreatedDBObject(circle, true);
                tr.Commit();
                
                return circle.ObjectId;
            }
        }

        /// <summary>
        /// Tính khoảng cách giữa 2 điểm
        /// </summary>
        public static double GetDistance(Point3d p1, Point3d p2)
        {
            return p1.DistanceTo(p2);
        }

        /// <summary>
        /// Tính góc giữa 2 điểm (radian)
        /// </summary>
        public static double GetAngle(Point3d p1, Point3d p2)
        {
            return Math.Atan2(p2.Y - p1.Y, p2.X - p1.X);
        }

        /// <summary>
        /// Làm tròn số đến n chữ số thập phân
        /// </summary>
        public static double RoundTo(double value, int decimals = 3)
        {
            return Math.Round(value, decimals);
        }
    }

    /// <summary>
    /// Utilities - Zoom, Highlight entities
    /// </summary>
    public static class VIUtilities
    {
        /// <summary>
        /// Zoom đến một entity
        /// </summary>
        public static void ZoomToEntity(ObjectId entityId)
        {
            if (entityId.IsNull) return;
            
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var entity = tr.GetObject(entityId, OpenMode.ForRead) as AcEntity;
                if (entity != null)
                {
                    var extents = entity.GeometricExtents;
                    ZoomToExtents(extents.MinPoint, extents.MaxPoint);
                }
                tr.Commit();
            }
        }

        /// <summary>
        /// Zoom đến một điểm với tỷ lệ
        /// </summary>
        public static void ZoomToPoint(Point3d point, double scale = 100)
        {
            var minPt = new Point3d(point.X - scale, point.Y - scale, 0);
            var maxPt = new Point3d(point.X + scale, point.Y + scale, 0);
            ZoomToExtents(minPt, maxPt);
        }

        /// <summary>
        /// Zoom đến extents
        /// </summary>
        public static void ZoomToExtents(Point3d minPoint, Point3d maxPoint)
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            
            // Mở rộng 10%
            var dx = (maxPoint.X - minPoint.X) * 0.1;
            var dy = (maxPoint.Y - minPoint.Y) * 0.1;
            
            var min = new Point3d(minPoint.X - dx, minPoint.Y - dy, 0);
            var max = new Point3d(maxPoint.X + dx, maxPoint.Y + dy, 0);
            
            var viewCtr = new Point3d((min.X + max.X) / 2, (min.Y + max.Y) / 2, 0);
            var viewWidth = max.X - min.X;
            var viewHeight = max.Y - min.Y;
            
            using (var view = ed.GetCurrentView())
            {
                view.CenterPoint = new Point2d(viewCtr.X, viewCtr.Y);
                view.Width = viewWidth;
                view.Height = viewHeight;
                ed.SetCurrentView(view);
            }
        }

        /// <summary>
        /// Highlight entity tạm thời
        /// </summary>
        public static void HighlightEntity(ObjectId entityId)
        {
            if (entityId.IsNull) return;
            
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var entity = tr.GetObject(entityId, OpenMode.ForRead) as AcEntity;
                entity?.Highlight();
                tr.Commit();
            }
        }

        /// <summary>
        /// Unhighlight entity
        /// </summary>
        public static void UnhighlightEntity(ObjectId entityId)
        {
            if (entityId.IsNull) return;
            
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var entity = tr.GetObject(entityId, OpenMode.ForRead) as AcEntity;
                entity?.Unhighlight();
                tr.Commit();
            }
        }

        /// <summary>
        /// Zoom đến Alignment
        /// </summary>
        public static void ZoomToAlignment(ObjectId alignmentId)
        {
            ZoomToEntity(alignmentId);
        }

        /// <summary>
        /// Zoom đến Profile View
        /// </summary>
        public static void ZoomToProfileView(ObjectId profileViewId)
        {
            ZoomToEntity(profileViewId);
        }

        /// <summary>
        /// Zoom đến Section View
        /// </summary>
        public static void ZoomToSectionView(ObjectId sectionViewId)
        {
            ZoomToEntity(sectionViewId);
        }

        /// <summary>
        /// Zoom đến Sample Line
        /// </summary>
        public static void ZoomToSampleLine(ObjectId sampleLineId)
        {
            ZoomToEntity(sampleLineId);
        }
    }

    /// <summary>
    /// CAD Class - Thao tác CAD cơ bản
    /// </summary>
    public static class VICadClass
    {
        /// <summary>
        /// Zoom Objects - Zoom đến các đối tượng được chọn
        /// </summary>
        [CommandMethod("VI_ZO")]
        public static void ZoomObjects()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;
            
            var result = ed.GetSelection();
            if (result.Status != PromptStatus.OK) return;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                Point3d? minPt = null;
                Point3d? maxPt = null;
                
                foreach (SelectedObject so in result.Value)
                {
                    var entity = tr.GetObject(so.ObjectId, OpenMode.ForRead) as AcEntity;
                    if (entity != null)
                    {
                        var ext = entity.GeometricExtents;
                        if (minPt == null)
                        {
                            minPt = ext.MinPoint;
                            maxPt = ext.MaxPoint;
                        }
                        else
                        {
                            minPt = new Point3d(
                                Math.Min(minPt.Value.X, ext.MinPoint.X),
                                Math.Min(minPt.Value.Y, ext.MinPoint.Y),
                                0);
                            maxPt = new Point3d(
                                Math.Max(maxPt!.Value.X, ext.MaxPoint.X),
                                Math.Max(maxPt.Value.Y, ext.MaxPoint.Y),
                                0);
                        }
                    }
                }
                
                if (minPt != null && maxPt != null)
                {
                    VIUtilities.ZoomToExtents(minPt.Value, maxPt.Value);
                }
                
                tr.Commit();
            }
            
            ed.WriteMessage("\n✅ Đã zoom đến đối tượng được chọn.");
        }

        /// <summary>
        /// Zoom XY - Zoom đến tọa độ X, Y
        /// </summary>
        [CommandMethod("VI_ZOOMXY")]
        public static void ZoomXY()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            
            var xResult = ed.GetDouble("\nNhập tọa độ X: ");
            if (xResult.Status != PromptStatus.OK) return;
            
            var yResult = ed.GetDouble("\nNhập tọa độ Y: ");
            if (yResult.Status != PromptStatus.OK) return;
            
            var scaleResult = ed.GetDouble("\nNhập scale (100): ");
            double scale = scaleResult.Status == PromptStatus.OK ? scaleResult.Value : 100;
            
            var point = new Point3d(xResult.Value, yResult.Value, 0);
            VIUtilities.ZoomToPoint(point, scale);
            
            ed.WriteMessage($"\n✅ Đã zoom đến ({xResult.Value:F2}, {yResult.Value:F2})");
        }

        /// <summary>
        /// Copy entity đến vị trí mới
        /// </summary>
        public static ObjectId CopyEntity(ObjectId entityId, Vector3d displacement)
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var entity = tr.GetObject(entityId, OpenMode.ForRead) as AcEntity;
                if (entity == null)
                {
                    tr.Commit();
                    return ObjectId.Null;
                }
                
                var clone = entity.Clone() as AcEntity;
                if (clone == null)
                {
                    tr.Commit();
                    return ObjectId.Null;
                }
                
                var matrix = Matrix3d.Displacement(displacement);
                clone.TransformBy(matrix);
                
                var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                var btr = tr.GetObject(bt![BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                
                btr!.AppendEntity(clone);
                tr.AddNewlyCreatedDBObject(clone, true);
                tr.Commit();
                
                return clone.ObjectId;
            }
        }

        /// <summary>
        /// Di chuyển entity
        /// </summary>
        public static void MoveEntity(ObjectId entityId, Vector3d displacement)
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var entity = tr.GetObject(entityId, OpenMode.ForWrite) as AcEntity;
                if (entity != null)
                {
                    var matrix = Matrix3d.Displacement(displacement);
                    entity.TransformBy(matrix);
                }
                tr.Commit();
            }
        }

        /// <summary>
        /// Xoay entity
        /// </summary>
        public static void RotateEntity(ObjectId entityId, Point3d basePoint, double angle)
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var entity = tr.GetObject(entityId, OpenMode.ForWrite) as AcEntity;
                if (entity != null)
                {
                    var matrix = Matrix3d.Rotation(angle, Vector3d.ZAxis, basePoint);
                    entity.TransformBy(matrix);
                }
                tr.Commit();
            }
        }

        /// <summary>
        /// Scale entity
        /// </summary>
        public static void ScaleEntity(ObjectId entityId, Point3d basePoint, double scaleFactor)
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var entity = tr.GetObject(entityId, OpenMode.ForWrite) as AcEntity;
                if (entity != null)
                {
                    var matrix = Matrix3d.Scaling(scaleFactor, basePoint);
                    entity.TransformBy(matrix);
                }
                tr.Commit();
            }
        }

        /// <summary>
        /// Lấy tất cả entities trong layer
        /// </summary>
        public static List<ObjectId> GetEntitiesInLayer(string layerName)
        {
            var result = new List<ObjectId>();
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                var btr = tr.GetObject(bt![BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                
                foreach (ObjectId id in btr!)
                {
                    var entity = tr.GetObject(id, OpenMode.ForRead) as AcEntity;
                    if (entity != null && entity.Layer.Equals(layerName, StringComparison.OrdinalIgnoreCase))
                    {
                        result.Add(id);
                    }
                }
                
                tr.Commit();
            }
            
            return result;
        }
    }
}
