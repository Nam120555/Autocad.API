// 25.PrintLayoutTools.cs - Công cụ In ấn & Layout từ VisualINFRA
// Viết lại cho AutoCAD 2026 / Civil 3D 2026

using System;
using System.Collections.Generic;
using System.Linq;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using Autodesk.Civil;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.VisualINFRA.PrintLayoutTools))]

namespace MyFirstProject.VisualINFRA
{
    /// <summary>
    /// Công cụ In ấn và Layout - từ VisualINFRA
    /// </summary>
    public class PrintLayoutTools
    {
        #region Tạo khung tuyến

        /// <summary>
        /// Tạo khung tuyến cho bình đồ
        /// </summary>
        [CommandMethod("VI_CreateKhungTuyen")]
        public static void CreateKhungTuyen()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // Nhập thông số khung
                var widthResult = ed.GetDouble("\nChiều rộng khung (mm trên giấy, VD: 297): ");
                if (widthResult.Status != PromptStatus.OK) return;
                double paperWidth = widthResult.Value;

                var heightResult = ed.GetDouble("\nChiều cao khung (mm trên giấy, VD: 210): ");
                if (heightResult.Status != PromptStatus.OK) return;
                double paperHeight = heightResult.Value;

                var scaleResult = ed.GetDouble("\nTỷ lệ bản vẽ (VD: 1000 cho 1:1000): ");
                if (scaleResult.Status != PromptStatus.OK) return;
                double scale = scaleResult.Value;

                // Chọn điểm chèn
                var ptResult = ed.GetPoint("\nChọn điểm góc dưới trái của khung: ");
                if (ptResult.Status != PromptStatus.OK) return;
                var insertPoint = ptResult.Value;

                // Tính kích thước thực tế
                double realWidth = paperWidth * scale / 1000;  // Chuyển mm sang m, nhân với scale
                double realHeight = paperHeight * scale / 1000;

                // Tạo layer
                string layerName = "VI_KHUNG_TUYEN";
                VIFunc.CreateLayer(layerName, 7); // Màu trắng

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    var btr = tr.GetObject(bt![BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // Vẽ khung chính
                    var points = new Point2dCollection
                    {
                        new Point2d(insertPoint.X, insertPoint.Y),
                        new Point2d(insertPoint.X + realWidth, insertPoint.Y),
                        new Point2d(insertPoint.X + realWidth, insertPoint.Y + realHeight),
                        new Point2d(insertPoint.X, insertPoint.Y + realHeight)
                    };

                    var pline = new Polyline();
                    for (int i = 0; i < points.Count; i++)
                    {
                        pline.AddVertexAt(i, points[i], 0, 0, 0);
                    }
                    pline.Closed = true;
                    pline.Layer = layerName;

                    btr!.AppendEntity(pline);
                    tr.AddNewlyCreatedDBObject(pline, true);

                    // Thêm text tỷ lệ
                    double textHeight = realHeight * 0.02;
                    var textPos = new Point3d(insertPoint.X + realWidth / 2, insertPoint.Y - textHeight * 2, 0);
                    var scaleText = new DBText
                    {
                        TextString = $"TỶ LỆ 1:{scale:F0}",
                        Position = textPos,
                        Height = textHeight,
                        HorizontalMode = TextHorizontalMode.TextCenter,
                        AlignmentPoint = textPos,
                        Layer = layerName
                    };

                    btr.AppendEntity(scaleText);
                    tr.AddNewlyCreatedDBObject(scaleText, true);

                    tr.Commit();
                }

                ed.WriteMessage($"\n Đã tạo khung tuyến:");
                ed.WriteMessage($"\n   Kích thước: {realWidth:F2}m x {realHeight:F2}m");
                ed.WriteMessage($"\n   Tỷ lệ: 1:{scale:F0}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Tạo khung view cho Layout

        /// <summary>
        /// Tạo viewport trong Layout
        /// </summary>
        [CommandMethod("VI_CreateViewport")]
        public static void CreateViewport()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // Chuyển sang Layout nếu đang ở Model
                if (doc.Database.TileMode)
                {
                    ed.WriteMessage("\n Đang ở Model Space. Chuyển sang Layout để tạo Viewport.");
                    return;
                }

                // Nhập thông số
                var ptResult = ed.GetPoint("\nChọn góc thứ nhất của Viewport: ");
                if (ptResult.Status != PromptStatus.OK) return;
                var pt1 = ptResult.Value;

                var pt2Result = ed.GetCorner("\nChọn góc đối diện: ", pt1);
                if (pt2Result.Status != PromptStatus.OK) return;
                var pt2 = pt2Result.Value;

                var scaleResult = ed.GetDouble("\nTỷ lệ viewport (VD: 1000 cho 1:1000): ");
                if (scaleResult.Status != PromptStatus.OK) return;
                double scale = scaleResult.Value;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    var layoutId = (ObjectId)AcadApplication.GetSystemVariable("CTAB");
                    var layout = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    // Tạo Viewport
                    var vp = new Viewport
                    {
                        CenterPoint = new Point3d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2, 0),
                        Width = Math.Abs(pt2.X - pt1.X),
                        Height = Math.Abs(pt2.Y - pt1.Y),
                        CustomScale = 1.0 / scale
                    };

                    layout!.AppendEntity(vp);
                    tr.AddNewlyCreatedDBObject(vp, true);
                    vp.On = true;

                    ed.WriteMessage($"\n Đã tạo Viewport với tỷ lệ 1:{scale:F0}");
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Tạo khung view trắc ngang

        /// <summary>
        /// Tạo khung view cho trắc ngang
        /// </summary>
        [CommandMethod("VI_CreateKhungTracNgang")]
        public static void CreateKhungTracNgang()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // Nhập thông số
                var widthResult = ed.GetDouble("\nChiều rộng khung (m): ");
                if (widthResult.Status != PromptStatus.OK) return;
                double width = widthResult.Value;

                var heightResult = ed.GetDouble("\nChiều cao khung (m): ");
                if (heightResult.Status != PromptStatus.OK) return;
                double height = heightResult.Value;

                var countResult = ed.GetInteger("\nSố lượng khung: ");
                if (countResult.Status != PromptStatus.OK) return;
                int count = countResult.Value;

                var colResult = ed.GetInteger("\nSố cột (khung trên 1 hàng): ");
                if (colResult.Status != PromptStatus.OK) return;
                int cols = colResult.Value;

                var gapResult = ed.GetDouble("\nKhoảng cách giữa các khung (m): ");
                if (gapResult.Status != PromptStatus.OK) return;
                double gap = gapResult.Value;

                // Chọn điểm chèn
                var ptResult = ed.GetPoint("\nChọn điểm góc dưới trái của khung đầu tiên: ");
                if (ptResult.Status != PromptStatus.OK) return;
                var startPoint = ptResult.Value;

                // Tạo layer
                string layerName = "VI_KHUNG_TN";
                VIFunc.CreateLayer(layerName, 8);

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    var btr = tr.GetObject(bt![BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    for (int i = 0; i < count; i++)
                    {
                        int row = i / cols;
                        int col = i % cols;

                        double x = startPoint.X + col * (width + gap);
                        double y = startPoint.Y - row * (height + gap);

                        // Vẽ khung
                        var points = new Point2dCollection
                        {
                            new Point2d(x, y),
                            new Point2d(x + width, y),
                            new Point2d(x + width, y + height),
                            new Point2d(x, y + height)
                        };

                        var pline = new Polyline();
                        for (int j = 0; j < points.Count; j++)
                        {
                            pline.AddVertexAt(j, points[j], 0, 0, 0);
                        }
                        pline.Closed = true;
                        pline.Layer = layerName;

                        btr!.AppendEntity(pline);
                        tr.AddNewlyCreatedDBObject(pline, true);
                    }

                    tr.Commit();
                }

                ed.WriteMessage($"\n Đã tạo {count} khung trắc ngang ({(count + cols - 1) / cols} hàng x {cols} cột)");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Fit khung in

        /// <summary>
        /// Fit các Section View vào khung in
        /// </summary>
        [CommandMethod("VI_FitKhungIn")]
        public static void FitKhungIn()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;
            var civilDoc = CivilApplication.ActiveDocument;

            try
            {
                // Chọn Section View
                var svResult = ed.GetSelection(new SelectionFilter(new[] { 
                    new TypedValue((int)DxfCode.Start, "AECC_SECTION_VIEW") 
                }));

                if (svResult.Status != PromptStatus.OK || svResult.Value.Count == 0)
                {
                    ed.WriteMessage("\n Không chọn được Section View.");
                    return;
                }

                // Nhập thông số khung
                var widthResult = ed.GetDouble("\nChiều rộng khung (m): ");
                if (widthResult.Status != PromptStatus.OK) return;
                double frameWidth = widthResult.Value;

                var heightResult = ed.GetDouble("\nChiều cao khung (m): ");
                if (heightResult.Status != PromptStatus.OK) return;
                double frameHeight = heightResult.Value;

                // Chọn điểm bắt đầu
                var ptResult = ed.GetPoint("\nChọn điểm góc dưới trái khung đầu tiên: ");
                if (ptResult.Status != PromptStatus.OK) return;
                var startPoint = ptResult.Value;

                var colResult = ed.GetInteger("\nSố cột: ");
                if (colResult.Status != PromptStatus.OK) return;
                int cols = colResult.Value;

                var gapResult = ed.GetDouble("\nKhoảng cách giữa các khung (m): ");
                if (gapResult.Status != PromptStatus.OK) return;
                double gap = gapResult.Value;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    int count = 0;
                    foreach (SelectedObject so in svResult.Value)
                    {
                        var sv = tr.GetObject(so.ObjectId, OpenMode.ForWrite) as SectionView;
                        if (sv == null) continue;

                        int row = count / cols;
                        int col = count % cols;

                        // Tính vị trí mới
                        double x = startPoint.X + col * (frameWidth + gap) + frameWidth / 2;
                        double y = startPoint.Y - row * (frameHeight + gap) + frameHeight / 2;

                        // Di chuyển Section View
                        var currentPos = sv.Location;
                        var displacement = new Vector3d(x - currentPos.X, y - currentPos.Y, 0);
                        
                        var matrix = Matrix3d.Displacement(displacement);
                        sv.TransformBy(matrix);

                        count++;
                    }

                    tr.Commit();
                    ed.WriteMessage($"\n Đã fit {count} Section View vào khung in.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Zoom đến Section View

        /// <summary>
        /// Zoom đến Section View theo tên cọc
        /// </summary>
        [CommandMethod("VI_ZoomToSection")]
        public static void ZoomToSection()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;
            var civilDoc = CivilApplication.ActiveDocument;

            try
            {
                // Nhập tên cọc
                var nameResult = ed.GetString("\nNhập tên cọc (VD: Km0+100): ");
                if (nameResult.Status != PromptStatus.OK) return;
                string searchName = nameResult.StringResult.ToUpper();

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    var btr = tr.GetObject(bt![BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId id in btr!)
                    {
                        var sv = tr.GetObject(id, OpenMode.ForRead) as SectionView;
                        if (sv != null && sv.Name.ToUpper().Contains(searchName))
                        {
                            VIUtilities.ZoomToEntity(sv.ObjectId);
                            ed.WriteMessage($"\n Đã zoom đến Section View: {sv.Name}");
                            tr.Commit();
                            return;
                        }
                    }

                    ed.WriteMessage($"\n Không tìm thấy Section View chứa '{searchName}'.");
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Copy Layout

        /// <summary>
        /// Copy Layout hiện tại
        /// </summary>
        [CommandMethod("VI_CopyLayout")]
        public static void CopyLayout()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // Lấy tên layout hiện tại
                string currentLayout = (string)AcadApplication.GetSystemVariable("CTAB");

                if (currentLayout == "Model")
                {
                    ed.WriteMessage("\n Không thể copy Model Space.");
                    return;
                }

                // Nhập tên layout mới
                var nameResult = ed.GetString($"\nNhập tên Layout mới (từ '{currentLayout}'): ");
                if (nameResult.Status != PromptStatus.OK) return;
                string newName = nameResult.StringResult;

                if (string.IsNullOrWhiteSpace(newName))
                {
                    ed.WriteMessage("\n Tên Layout không hợp lệ.");
                    return;
                }

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var layoutMgr = LayoutManager.Current;

                    // Kiểm tra tên đã tồn tại
                    var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    if (layoutDict!.Contains(newName))
                    {
                        ed.WriteMessage($"\n Layout '{newName}' đã tồn tại.");
                        tr.Commit();
                        return;
                    }

                    // Copy layout
                    layoutMgr.CopyLayout(currentLayout, newName);

                    ed.WriteMessage($"\n Đã copy Layout '{currentLayout}' thành '{newName}'.");
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region List Layouts

        /// <summary>
        /// Liệt kê tất cả Layout
        /// </summary>
        [CommandMethod("VI_ListLayouts")]
        public static void ListLayouts()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                    ed.WriteMessage($"\n\n{'=',-50}");
                    ed.WriteMessage($"\n DANH SÁCH LAYOUT");
                    ed.WriteMessage($"\n{'=',-50}");

                    int stt = 1;
                    foreach (var entry in layoutDict!)
                    {
                        var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                        if (layout != null)
                        {
                            string isCurrent = layout.LayoutName == (string)AcadApplication.GetSystemVariable("CTAB") ? " ← CURRENT" : "";
                            ed.WriteMessage($"\n  {stt}. {layout.LayoutName}{isCurrent}");
                            stt++;
                        }
                    }

                    ed.WriteMessage($"\n{'=',-50}");
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion
    }
}
