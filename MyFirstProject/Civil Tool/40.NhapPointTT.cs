// 40.NhapPointTT.cs — Lệnh CTPO_NhapToaDoTrucTiep
// Nhập tọa độ X, Y, Z, Tên điểm, Ghi chú, Số thứ tự trực tiếp vào Civil 3D
// KHÔNG cần nhập từ Excel — gõ tay trên bảng → tạo COGO Point ngay lập tức

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using MyFirstProject.Extensions;

using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using ContentAlignment = System.Drawing.ContentAlignment;

#pragma warning disable CS8600, CS8602, CS8604, CS8618

[assembly: CommandClass(typeof(MyFirstProject.NhapPointTrucTiepCommands))]

namespace MyFirstProject
{
    // ══════════════════════════════════════════════════════════════
    // DATA CLASS
    // ══════════════════════════════════════════════════════════════

    public class DirectPointData
    {
        public int STT { get; set; }
        public string PointName { get; set; } = "";
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public string Description { get; set; } = "";

        public bool IsValid => !string.IsNullOrWhiteSpace(PointName)
                            && (X != 0 || Y != 0);
    }

    // ══════════════════════════════════════════════════════════════
    // FORM: NHẬP TỌA ĐỘ TRỰC TIẾP
    // ══════════════════════════════════════════════════════════════


    // ══════════════════════════════════════════════════════════════
    // LỆNH CHÍNH
    // ══════════════════════════════════════════════════════════════

    public class NhapPointTrucTiepCommands
    {
        /// <summary>
        /// Nhập tọa độ X, Y, Z trực tiếp vào bảng → Tạo COGO Point trên Civil 3D
        /// Không cần file Excel — gõ tay hoặc dán từ clipboard
        /// </summary>
        [CommandMethod("CTPO_NhapToaDoTrucTiep")]
        public static void CTPO_NhapToaDoTrucTiep()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            // 1. Mở form nhập liệu
            List<DirectPointData> pointList;
            string pointGroupName, pointStyleName, labelStyleName, layerName;
            bool isLightMode;
            string blockName;
            double blockScale, textHeight;

            using (var form = new FormNhapPointTrucTiep())
            {
                AcadApp.ShowModalDialog(form);
                if (!form.IsConfirmed) return;

                pointList = form.PointDataList;
                pointGroupName = form.SelectedPointGroup;
                pointStyleName = form.SelectedPointStyle;
                labelStyleName = form.SelectedLabelStyle;
                layerName = form.LayerName;
                isLightMode = form.IsLightMode;
                blockName = form.SelectedBlockName;
                blockScale = form.BlockScale;
                textHeight = form.TextHeight;
            }

            if (pointList.Count == 0)
            {
                ed.WriteMessage("\n⚠ Không có dữ liệu để tạo.");
                return;
            }

            // ═══════════════════════════════════════════════════════════
            // CHẾ ĐỘ NHẸ: TẠO BLOCK + TEXT (KHÔNG TẠO COGO POINT)
            // ═══════════════════════════════════════════════════════════
            if (isLightMode)
            {
                ed.WriteMessage($"\n⚡ Chế độ nhẹ — Bắt đầu tạo {pointList.Count} Block + Text...");

                try
                {
                    using var tr = db.TransactionManager.StartTransaction();

                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var modelSpace = (BlockTableRecord)tr.GetObject(
                        bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // Tạo layer nếu chưa có
                    EnsureLayer(db, tr, layerName);

                    bool useBlock = !string.IsNullOrEmpty(blockName) && bt.Has(blockName);
                    int created = 0;
                    int errors = 0;

                    foreach (var pt in pointList)
                    {
                        try
                        {
                            Point3d insertPt = new(pt.X, pt.Y, pt.Z);

                            if (useBlock)
                            {
                                // ── Chèn BlockReference ──
                                ObjectId blockDefId = bt[blockName];
                                BlockReference blkRef = new(insertPt, blockDefId)
                                {
                                    Layer = layerName,
                                    ScaleFactors = new Scale3d(blockScale, blockScale, blockScale)
                                };

                                modelSpace.AppendEntity(blkRef);
                                tr.AddNewlyCreatedDBObject(blkRef, true);

                                // Gán Attribute nếu block có AttributeDefinition
                                var blockDef = (BlockTableRecord)tr.GetObject(blockDefId, OpenMode.ForRead);
                                foreach (ObjectId attDefId in blockDef)
                                {
                                    var obj = tr.GetObject(attDefId, OpenMode.ForRead);
                                    if (obj is AttributeDefinition attDef)
                                    {
                                        AttributeReference attRef = new();
                                        attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform);

                                        string tagUpper = attDef.Tag.ToUpper();
                                        if (tagUpper == "NAME" || tagUpper == "TEN" || tagUpper == "LABEL"
                                            || tagUpper == "TEXT" || tagUpper == "TITLE")
                                        {
                                            attRef.TextString = pt.PointName;
                                        }
                                        else if (tagUpper == "X" || tagUpper == "EASTING")
                                        {
                                            attRef.TextString = pt.X.ToString("F3");
                                        }
                                        else if (tagUpper == "Y" || tagUpper == "NORTHING")
                                        {
                                            attRef.TextString = pt.Y.ToString("F3");
                                        }
                                        else if (tagUpper == "Z" || tagUpper == "ELEVATION" || tagUpper == "CAODO")
                                        {
                                            attRef.TextString = pt.Z.ToString("F3");
                                        }
                                        else if (tagUpper == "DESC" || tagUpper == "DESCRIPTION" || tagUpper == "GHICHU")
                                        {
                                            attRef.TextString = pt.Description;
                                        }
                                        else
                                        {
                                            attRef.TextString = pt.PointName;
                                        }

                                        attRef.Layer = layerName;
                                        blkRef.AttributeCollection.AppendAttribute(attRef);
                                        tr.AddNewlyCreatedDBObject(attRef, true);
                                    }
                                }
                            }
                            else
                            {
                                // ── Không dùng Block: Tạo dấu "+" marker ──
                                double markerSize = textHeight * 0.8;

                                // Vạch ngang
                                var lineH = new Line(
                                    new Point3d(pt.X - markerSize, pt.Y, pt.Z),
                                    new Point3d(pt.X + markerSize, pt.Y, pt.Z))
                                { Layer = layerName };
                                modelSpace.AppendEntity(lineH);
                                tr.AddNewlyCreatedDBObject(lineH, true);

                                // Vạch dọc
                                var lineV = new Line(
                                    new Point3d(pt.X, pt.Y - markerSize, pt.Z),
                                    new Point3d(pt.X, pt.Y + markerSize, pt.Z))
                                { Layer = layerName };
                                modelSpace.AppendEntity(lineV);
                                tr.AddNewlyCreatedDBObject(lineV, true);
                            }

                            // ── Tạo Text hiển thị tên điểm ──
                            if (!string.IsNullOrWhiteSpace(pt.PointName))
                            {
                                DBText nameText = new()
                                {
                                    Position = new Point3d(pt.X + textHeight * 0.5, pt.Y + textHeight * 0.5, pt.Z),
                                    Height = textHeight,
                                    TextString = pt.PointName,
                                    Layer = layerName
                                };
                                nameText.SetDatabaseDefaults();
                                modelSpace.AppendEntity(nameText);
                                tr.AddNewlyCreatedDBObject(nameText, true);
                            }

                            created++;
                        }
                        catch (System.Exception ex)
                        {
                            errors++;
                            ed.WriteMessage($"\n  ⚠ Lỗi tạo điểm {pt.PointName}: {ex.Message}");
                        }
                    }

                    tr.Commit();

                    // ─── Hiển thị kết quả ───
                    ed.WriteMessage("\n");
                    ed.WriteMessage($"\n╔════════════════════════════════════════════════════╗");
                    ed.WriteMessage($"\n║  ⚡ KẾT QUẢ NHẬP TỌA ĐỘ (CHẾ ĐỘ NHẸ)            ║");
                    ed.WriteMessage($"\n╠════════════════════════════════════════════════════╣");
                    ed.WriteMessage($"\n║  ✅ Đã tạo: {created,5} điểm (Block + Text)          ║");
                    if (useBlock)
                        ed.WriteMessage($"\n║  📦 Block:  {blockName,-38} ║");
                    else
                        ed.WriteMessage($"\n║  ✚ Marker:  Dấu + và Text                        ║");
                    if (errors > 0)
                        ed.WriteMessage($"\n║  ⚠ Lỗi:    {errors,5} điểm                          ║");
                    ed.WriteMessage($"\n╚════════════════════════════════════════════════════╝");

                    // Bảng chi tiết
                    ed.WriteMessage($"\n\n{"STT",-5} {"Tên điểm",-15} {"X (Easting)",-14} {"Y (Northing)",-14} {"Z",-10} {"Ghi chú",-20}");
                    ed.WriteMessage($"\n{new string('─', 80)}");

                    foreach (var pt in pointList)
                    {
                        ed.WriteMessage($"\n{pt.STT,-5} {pt.PointName,-15} {pt.X,-14:F3} {pt.Y,-14:F3} {pt.Z,-10:F3} {pt.Description,-20}");
                    }

                    ed.WriteMessage($"\n{new string('═', 80)}");
                    ed.WriteMessage($"\n\n⚡ HOÀN THÀNH — {created} điểm nhẹ (không COGO Point)!");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
                }

                return; // Kết thúc luồng chế độ nhẹ
            }

            // ═══════════════════════════════════════════════════════════
            // CHẾ ĐỘ THƯỜNG: TẠO COGO POINT (GIỮ NGUYÊN LOGIC CŨ)
            // ═══════════════════════════════════════════════════════════
            var civilDoc = CivilApplication.ActiveDocument;

            ed.WriteMessage($"\n📍 Bắt đầu tạo {pointList.Count} COGO Points...");

            try
            {
                using var tr = db.TransactionManager.StartTransaction();

                CogoPointCollection cogoPointColl = civilDoc.CogoPoints;

                // Lấy Point Style ID
                ObjectId pointStyleId = ObjectId.Null;
                if (!string.IsNullOrEmpty(pointStyleName))
                {
                    try
                    {
                        pointStyleId = civilDoc.Styles.PointStyles[pointStyleName];
                    }
                    catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                }

                // Lấy Label Style ID
                ObjectId labelStyleId = ObjectId.Null;
                if (!string.IsNullOrEmpty(labelStyleName))
                {
                    try
                    {
                        labelStyleId = civilDoc.Styles.LabelStyles.PointLabelStyles.LabelStyles[labelStyleName];
                    }
                    catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                }

                int created = 0;
                int errors = 0;

                foreach (var pt in pointList)
                {
                    try
                    {
                        Point3d insertPt = new(pt.X, pt.Y, pt.Z);

                        // Tạo COGO Point
                        ObjectId ptId = cogoPointColl.Add(insertPt, pt.Description, true);
                        CogoPoint cogoPoint = tr.GetObject(ptId, OpenMode.ForWrite) as CogoPoint;

                        if (cogoPoint != null)
                        {
                            // Đặt tên (PointName)
                            if (!string.IsNullOrWhiteSpace(pt.PointName))
                                cogoPoint.PointName = pt.PointName;

                            // Đặt RawDescription
                            if (!string.IsNullOrWhiteSpace(pt.Description))
                                cogoPoint.RawDescription = pt.Description;

                            // Đặt Style
                            if (!pointStyleId.IsNull)
                                cogoPoint.StyleId = pointStyleId;

                            // Đặt Label Style
                            if (!labelStyleId.IsNull)
                                cogoPoint.LabelStyleId = labelStyleId;

                            created++;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        errors++;
                        ed.WriteMessage($"\n  ⚠ Lỗi tạo điểm {pt.PointName}: {ex.Message}");
                    }
                }

                // Cập nhật Point Group nếu chọn
                if (!string.IsNullOrEmpty(pointGroupName))
                {
                    try
                    {
                        PointGroupCollection pgColl = civilDoc.PointGroups;
                        ObjectId pgId;

                        if (pgColl.Contains(pointGroupName))
                        {
                            pgId = pgColl[pointGroupName];
                        }
                        else
                        {
                            pgId = pgColl.Add(pointGroupName);
                        }

                        // Set query theo description
                        PointGroup pg = tr.GetObject(pgId, OpenMode.ForWrite) as PointGroup;
                        if (pg != null)
                        {
                            pg.Update();
                        }
                    }
                    catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                }

                // Cập nhật tất cả Point Groups
                try
                {
                    civilDoc.PointGroups.UpdateAllPointGroups();
                }
                catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

                tr.Commit();

                // ─── Hiển thị kết quả ───
                ed.WriteMessage("\n");
                ed.WriteMessage($"\n╔════════════════════════════════════════════════════╗");
                ed.WriteMessage($"\n║  📍 KẾT QUẢ NHẬP TỌA ĐỘ TRỰC TIẾP               ║");
                ed.WriteMessage($"\n╠════════════════════════════════════════════════════╣");
                ed.WriteMessage($"\n║  ✅ Đã tạo: {created,5} COGO Points                  ║");
                if (errors > 0)
                    ed.WriteMessage($"\n║  ⚠ Lỗi:    {errors,5} điểm                          ║");
                ed.WriteMessage($"\n╚════════════════════════════════════════════════════╝");

                // Bảng chi tiết
                ed.WriteMessage($"\n\n{"STT",-5} {"Tên điểm",-15} {"X (Easting)",-14} {"Y (Northing)",-14} {"Z",-10} {"Ghi chú",-20}");
                ed.WriteMessage($"\n{new string('─', 80)}");

                foreach (var pt in pointList)
                {
                    ed.WriteMessage($"\n{pt.STT,-5} {pt.PointName,-15} {pt.X,-14:F3} {pt.Y,-14:F3} {pt.Z,-10:F3} {pt.Description,-20}");
                }

                ed.WriteMessage($"\n{new string('═', 80)}");
                ed.WriteMessage($"\n\n✅ HOÀN THÀNH nhập {created} điểm!");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Tạo layer nếu chưa tồn tại
        /// </summary>
        private static void EnsureLayer(Database db, Transaction tr, string layerName)
        {
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                var ltr = new LayerTableRecord()
                {
                    Name = layerName,
                    Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                        Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 3) // Green
                };
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }
    }
}
