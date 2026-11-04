using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.Runtime;
using Acad = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Colors;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.Acad_Tool.AT_DoiMauByLayer))]

namespace MyFirstProject.Acad_Tool
{
    public class AT_DoiMauByLayer
    {
        /// <summary>
        /// Lệnh chuyển toàn bộ đối tượng trong bản vẽ thành màu ByLayer.
        /// Bao gồm: đối tượng trong Model/Paper Space, đối tượng trong Block Definition,
        /// và cả thuộc tính Linetype, Lineweight cũng chuyển ByLayer.
        /// </summary>
        [CommandMethod("AT_DoiMauByLayer")]
        public static void DoiMauByLayerMethod()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            int countEntities = 0;
            int countBlocks = 0;
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

                // 1. Duyệt toàn bộ Block Table (bao gồm *Model_Space, *Paper_Space và các Block Definition)
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                foreach (ObjectId btrId in bt)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                    foreach (ObjectId entId in btr)
                    {
                        Entity ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                        if (ent == null) continue;

                        bool changed = false;

                        // Chuyển màu về ByLayer
                        if (ent.ColorIndex != 256) // 256 = ByLayer
                        {
                            ent.ColorIndex = 256;
                            changed = true;
                        }

                        // Chuyển Linetype về ByLayer
                        if (ent.Linetype != "ByLayer")
                        {
                            ent.Linetype = "ByLayer";
                            changed = true;
                        }

                        // Chuyển Lineweight về ByLayer
                        if (ent.LineWeight != LineWeight.ByLayer)
                        {
                            ent.LineWeight = LineWeight.ByLayer;
                            changed = true;
                        }

                        // Chuyển Plot Style Name về ByLayer
                        if (ent.PlotStyleName != "ByLayer")
                        {
                            ent.PlotStyleName = "ByLayer";
                            changed = true;
                        }

                        // Chuyển Transparency về ByLayer
                        if (ent.Transparency.IsByLayer == false)
                        {
                            ent.Transparency = new Transparency(TransparencyMethod.ByLayer);
                            changed = true;
                        }

                        // Xử lý riêng cho BlockReference: chuyển thuộc tính AttributeReference
                        if (ent is BlockReference blkRef)
                        {
                            foreach (ObjectId attId in blkRef.AttributeCollection)
                            {
                                AttributeReference attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                                if (attRef == null) continue;

                                if (attRef.ColorIndex != 256)
                                {
                                    attRef.ColorIndex = 256;
                                    changed = true;
                                }
                                if (attRef.Linetype != "ByLayer")
                                {
                                    attRef.Linetype = "ByLayer";
                                    changed = true;
                                }
                                if (attRef.LineWeight != LineWeight.ByLayer)
                                {
                                    attRef.LineWeight = LineWeight.ByLayer;
                                    changed = true;
                                }
                            }
                        }

                        // Xử lý cho Dimension: chuyển màu dimension style override
                        if (ent is Dimension dim)
                        {
                            var byLayerColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 256);
                            if (dim.Dimclrd.ColorIndex != 256)
                            {
                                dim.Dimclrd = byLayerColor;
                                changed = true;
                            }
                            if (dim.Dimclre.ColorIndex != 256)
                            {
                                dim.Dimclre = byLayerColor;
                                changed = true;
                            }
                            if (dim.Dimclrt.ColorIndex != 256)
                            {
                                dim.Dimclrt = byLayerColor;
                                changed = true;
                            }
                        }

                        // Xử lý cho MText/DBText nằm trong bảng (Table)
                        if (ent is MText mtext)
                        {
                            if (mtext.ColorIndex != 256)
                            {
                                mtext.ColorIndex = 256;
                                changed = true;
                            }
                        }

                        if (changed)
                        {
                            countEntities++;
                            if (btr.IsFromExternalReference == false &&
                                btr.Name != "*Model_Space" &&
                                !btr.Name.StartsWith("*Paper_Space"))
                            {
                                countBlocks++;
                            }
                        }
                    }
                }

                // ===== Bước cuối: Khóa lại các Layer đã tạm mở =====
                foreach (ObjectId lockedId in lockedLayerIds)
                {
                    LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(lockedId, OpenMode.ForWrite);
                    ltr.IsLocked = true;
                }

                tr.Commit();

                // Regenerate bản vẽ để hiển thị thay đổi
                ed.Regen();

                ed.WriteMessage($"\n[AT_DoiMauByLayer] ✓ Hoàn tất! Đã chuyển {countEntities} đối tượng thành ByLayer.");
                if (countBlocks > 0)
                    ed.WriteMessage($"\n  (Bao gồm {countBlocks} đối tượng bên trong Block Definitions)");
                if (countLockedLayers > 0)
                    ed.WriteMessage($"\n  (Đã tạm mở khóa {countLockedLayers} layer bị lock và khóa lại sau khi xử lý)");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                ed.WriteMessage("\nLỗi: " + e.Message);
            }
        }
    }
}
