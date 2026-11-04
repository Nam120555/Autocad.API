using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.Runtime;
using Acad = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.Acad_Tool.AT_Linew))]

namespace MyFirstProject.Acad_Tool
{
    public class AT_Linew
    {
        /// <summary>
        /// Lệnh chuyển toàn bộ Lineweight của đối tượng trong bản vẽ về Default,
        /// chuyển Lineweight của tất cả Layer về Default,
        /// và đặt biến hệ thống LWDISPLAY = 0 (tắt hiển thị lineweight toàn cục).
        /// </summary>
        [CommandMethod("AT_Linew")]
        public static void ChuyenLineweightVeDefault()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            int countEntities = 0;
            int countLayers = 0;
            int countLockedLayers = 0;

            try
            {
                using Transaction tr = db.TransactionManager.StartTransaction();

                // ===== Bước 0: Tạm mở khóa tất cả Locked Layer =====
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                var lockedLayerIds = new List<ObjectId>();

                foreach (ObjectId layerId in lt)
                {
                    LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                    if (ltr.IsLocked)
                    {
                        ltr.UpgradeOpen();
                        ltr.IsLocked = false;
                        lockedLayerIds.Add(layerId);
                        countLockedLayers++;
                    }
                }

                // ===== Bước 1: Chuyển Lineweight của tất cả Layer về Default =====
                foreach (ObjectId layerId in lt)
                {
                    LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                    if (ltr.LineWeight != LineWeight.ByLineWeightDefault)
                    {
                        ltr.LineWeight = LineWeight.ByLineWeightDefault;
                        countLayers++;
                    }
                }

                // ===== Bước 2: Duyệt toàn bộ Block Table, chuyển Lineweight về Default =====
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                foreach (ObjectId btrId in bt)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                    foreach (ObjectId entId in btr)
                    {
                        Entity ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                        if (ent == null) continue;

                        bool changed = false;

                        // Chuyển Lineweight về ByLineWeightDefault
                        if (ent.LineWeight != LineWeight.ByLineWeightDefault)
                        {
                            ent.LineWeight = LineWeight.ByLineWeightDefault;
                            changed = true;
                        }

                        // Xử lý BlockReference: chuyển thuộc tính AttributeReference
                        if (ent is BlockReference blkRef)
                        {
                            foreach (ObjectId attId in blkRef.AttributeCollection)
                            {
                                AttributeReference attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                                if (attRef == null) continue;

                                if (attRef.LineWeight != LineWeight.ByLineWeightDefault)
                                {
                                    attRef.LineWeight = LineWeight.ByLineWeightDefault;
                                    changed = true;
                                }
                            }
                        }

                        // Xử lý Dimension: chuyển lineweight override của dim
                        if (ent is Dimension dim)
                        {
                            if (dim.Dimlwd != LineWeight.ByLineWeightDefault)
                            {
                                dim.Dimlwd = LineWeight.ByLineWeightDefault;
                                changed = true;
                            }
                            if (dim.Dimlwe != LineWeight.ByLineWeightDefault)
                            {
                                dim.Dimlwe = LineWeight.ByLineWeightDefault;
                                changed = true;
                            }
                        }

                        if (changed)
                            countEntities++;
                    }
                }

                // ===== Bước 3: Đặt LWDISPLAY = 0 (tắt hiển thị lineweight toàn cục) =====
                Application.SetSystemVariable("LWDISPLAY", 0);

                // ===== Bước cuối: Khóa lại các Layer đã tạm mở =====
                foreach (ObjectId lockedId in lockedLayerIds)
                {
                    LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(lockedId, OpenMode.ForWrite);
                    ltr.IsLocked = true;
                }

                tr.Commit();

                // Regenerate bản vẽ để hiển thị thay đổi
                ed.Regen();

                ed.WriteMessage($"\n[AT_Linew] ✓ Hoàn tất!");
                ed.WriteMessage($"\n  - Đã chuyển Lineweight của {countEntities} đối tượng về Default.");
                ed.WriteMessage($"\n  - Đã chuyển Lineweight của {countLayers} Layer về Default.");
                ed.WriteMessage($"\n  - Đã đặt LWDISPLAY = 0 (tắt hiển thị lineweight toàn cục).");
                if (countLockedLayers > 0)
                    ed.WriteMessage($"\n  - Đã tạm mở khóa {countLockedLayers} layer bị lock và khóa lại sau khi xử lý.");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                ed.WriteMessage("\nLỗi: " + e.Message);
            }
        }
    }
}
