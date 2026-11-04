// 28.TaoViewTheoTuyen.cs - Tạo MView (Viewport) theo hướng Alignment
// Tạo viewport trong Layout theo hướng tuyến alignment, tỉ lệ 1:1, kích thước tùy chỉnh
// Viewport sẽ được xoay (ViewTwist) theo góc tiếp tuyến của alignment tại trung tâm đoạn

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using MyFirstProject.Extensions;

// Alias để tránh xung đột
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using ContentAlignment = System.Drawing.ContentAlignment;
using FontStyle = System.Drawing.FontStyle;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.TaoViewTheoTuyen))]

namespace MyFirstProject
{
    #region Form


    #endregion

    #region Command

    /// <summary>
    /// Lệnh tạo MView theo tuyến Alignment
    /// Vẽ khung polyline trong Model Space + Tạo Viewport trong Layout
    /// Xử lý từng phân đoạn (batch) để tránh crash khi tuyến dài
    /// </summary>
    public class TaoViewTheoTuyen
    {
        private const int BATCH_SIZE = 10; // Số segment mỗi batch

        [CommandMethod("CTC_TaoMViewTheoTuyen")]
        public static void CTC_TaoMViewTheoTuyen()
        {
            var ed = A.Ed;

            // Mở form cấu hình - vòng lặp để xử lý pick station trên tuyến
            using TaoMViewForm form = new();
            while (true)
            {
                var dialogResult = AcadApp.ShowModalDialog(form);

                // Nếu user click nút 📍 pick station → form trả về Retry
                if (dialogResult == DialogResult.Retry && form.IsPickMode)
                {
                    // Pick point trên tuyến
                    string label = (form.PickTargetColumn == "RangeFrom") ? "Từ Station" : "Đến Station";

                    var ppo = new PromptPointOptions(
                        $"\nChọn điểm trên tuyến cho [{label}] (hoặc ESC để hủy): ");
                    ppo.AllowNone = true;
                    var ppr = ed.GetPoint(ppo);

                    if (ppr.Status == PromptStatus.OK)
                    {
                        Point3d pickedPoint = ppr.Value;

                        // Tính station trên alignment
                        try
                        {
                            using Transaction tr = A.Db.TransactionManager.StartTransaction();
                            Alignment? alignment = tr.GetObject(
                                form.SelectedAlignmentId, OpenMode.ForRead) as Alignment;
                            if (alignment != null)
                            {
                                double station = 0, offset = 0;
                                alignment.StationOffset(
                                    pickedPoint.X, pickedPoint.Y, ref station, ref offset);
                                tr.Commit();

                                // Cảnh báo nếu điểm xa tuyến (offset > 500m)
                                bool useValue = true;
                                if (Math.Abs(offset) > 500)
                                {
                                    var msgResult = MessageBox.Show(
                                        $"Điểm bạn chọn cách tuyến {Math.Abs(offset):F1}m.\n\n" +
                                        $"Station tính được: {station:F2}\n" +
                                        $"Offset: {offset:F2}m\n\n" +
                                        $"Bạn có muốn sử dụng giá trị này không?",
                                        "Cảnh báo — Điểm xa tuyến",
                                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                    useValue = (msgResult == DialogResult.Yes);
                                }

                                if (useValue)
                                {
                                    form.SetPickedStation(station);
                                    ed.WriteMessage(
                                        $"\n Đã chọn Station = {station:F2} (offset = {offset:F2}m)");
                                }
                                else
                                {
                                    form.CancelPickMode();
                                }
                            }
                            else
                            {
                                tr.Commit();
                                form.CancelPickMode();
                            }
                        }
                        catch
                        {
                            form.CancelPickMode();
                        }
                    }
                    else
                    {
                        // User hủy pick (ESC)
                        form.CancelPickMode();
                    }

                    // Quay lại hiển form
                    continue;
                }

                // Nếu user nhấn OK (Tạo MView)
                if (dialogResult == DialogResult.OK && form.IsConfirmed)
                    break;

                // Nếu user hủy
                return;
            }

            if (form.ViewSegments.Count == 0) return;

            var db = A.Db;

            double scale = form.ViewScale;
            double vpWidthPaper = form.ViewWidth;
            double vpHeightPaper = form.ViewHeight;
            double viewWidthReal = vpWidthPaper * scale / 1000.0;
            double viewHeightReal = vpHeightPaper * scale / 1000.0;

            int totalSegments = form.ViewSegments.Count;

            // ══════════════════════════════════════════════════════
            // PHẦN 1: Vẽ khung polyline trong Model Space
            // ─ Xử lý theo BATCH để tránh crash khi tuyến dài
            // ══════════════════════════════════════════════════════
            int created = 0;
            string layerName = "CTC_MVIEW_TUYEN";
            CreateLayerIfNotExist(layerName, 3);

            // ── Xóa khung cũ trên layer CTC_MVIEW_TUYEN ──
            using (Transaction trDel = db.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTable? btDel = trDel.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    if (btDel != null)
                    {
                        BlockTableRecord? btrDel = trDel.GetObject(btDel[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        if (btrDel != null)
                        {
                            int deleted = 0;
                            foreach (ObjectId entId in btrDel)
                            {
                                var ent = trDel.GetObject(entId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                                if (ent == null) continue;
                                if (!string.Equals(ent.Layer, layerName, StringComparison.OrdinalIgnoreCase)) continue;

                                // Chỉ xóa Polyline và DBText trên layer này
                                if (ent is Polyline || ent is DBText)
                                {
                                    ent.UpgradeOpen();
                                    ent.Erase();
                                    deleted++;
                                }
                            }
                            if (deleted > 0)
                                ed.WriteMessage($"\n🗑 Đã xóa {deleted} đối tượng khung cũ trên layer '{layerName}'");
                        }
                    }
                    trDel.Commit();
                }
                catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
            }

            // ── Tạo khung polyline theo BATCH ──
            // Chia segments thành các batch nhỏ, mỗi batch 1 transaction riêng
            ed.WriteMessage($"\n Bắt đầu tạo {totalSegments} khung MView (batch {BATCH_SIZE} segments/lần)...");

            // Progress meter
            using var progress = new Autodesk.AutoCAD.Runtime.ProgressMeter();
            progress.Start($"Đang tạo {totalSegments} MView theo tuyến...");
            progress.SetLimit(totalSegments);

            for (int batchStart = 0; batchStart < totalSegments; batchStart += BATCH_SIZE)
            {
                int batchEnd = Math.Min(batchStart + BATCH_SIZE, totalSegments);
                int batchCount = batchEnd - batchStart;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        BlockTable? bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        if (bt == null) { tr.Commit(); continue; }
                        BlockTableRecord? btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        if (btr == null) { tr.Commit(); continue; }

                        for (int i = batchStart; i < batchEnd; i++)
                        {
                            var seg = form.ViewSegments[i];
                            try
                            {
                                double angle = seg.AngleRad, cx = seg.CenterX, cy = seg.CenterY;
                                double halfW = viewWidthReal / 2.0, halfH = viewHeightReal / 2.0;
                                double cosA = Math.Cos(angle), sinA = Math.Sin(angle);

                                Point2d[] corners = new Point2d[4];
                                double[] dxArr = { -halfW, halfW, halfW, -halfW };
                                double[] dyArr = { -halfH, -halfH, halfH, halfH };
                                for (int j = 0; j < 4; j++)
                                    corners[j] = new Point2d(dxArr[j] * cosA - dyArr[j] * sinA + cx,
                                                              dxArr[j] * sinA + dyArr[j] * cosA + cy);

                                var pline = new Polyline();
                                for (int j = 0; j < 4; j++)
                                    pline.AddVertexAt(j, corners[j], 0, 0, 0);
                                pline.Closed = true;
                                pline.Layer = layerName;
                                pline.ColorIndex = 3;
                                btr!.AppendEntity(pline);
                                tr.AddNewlyCreatedDBObject(pline, true);

                                double textH = viewHeightReal * 0.03;
                                var text = new DBText
                                {
                                    TextString = $"VIEW {seg.STT}",
                                    Position = new Point3d(cx, cy, 0),
                                    Height = textH, Rotation = angle,
                                    HorizontalMode = TextHorizontalMode.TextCenter,
                                    VerticalMode = TextVerticalMode.TextVerticalMid,
                                    AlignmentPoint = new Point3d(cx, cy, 0),
                                    Layer = layerName, ColorIndex = 3
                                };
                                btr.AppendEntity(text);
                                tr.AddNewlyCreatedDBObject(text, true);

                                double lyX = cx - halfW * cosA + halfH * sinA;
                                double lyY = cy - halfW * sinA - halfH * cosA;
                                var lyText = new DBText
                                {
                                    TextString = $"Sta {seg.StartStation:F0} - {seg.EndStation:F0}",
                                    Position = new Point3d(lyX, lyY, 0),
                                    Height = textH * 0.7, Rotation = angle,
                                    Layer = layerName, ColorIndex = 3
                                };
                                btr.AppendEntity(lyText);
                                tr.AddNewlyCreatedDBObject(lyText, true);

                                created++;
                                progress.MeterProgress();
                            }
                            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                        }

                        tr.Commit();
                        ed.WriteMessage($"\n   Hoàn tất batch khung {batchStart + 1}~{batchEnd}/{totalSegments}");
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        ed.WriteMessage($"\n   Lỗi batch khung {batchStart + 1}~{batchEnd}: {ex.Message}");
                    }
                }

                // Cho AutoCAD xử lý giữa các batch
                System.Windows.Forms.Application.DoEvents();
            }

            progress.Stop();
            ed.WriteMessage($"\n Đã tạo {created}/{totalSegments} khung MView trong Model Space (tỉ lệ 1:{scale:F0})");

            if (created == 0) return;

            // ══════════════════════════════════════════════════════
            // PHẦN 2: Tạo Viewport trong Layout
            // ─ Xử lý từng phân đoạn (batch) để tránh crash
            // ─ Lấy hướng từ khung CAD (polyline frame Model Space)
            // ─ Kích thước viewport = kích thước khung CAD (đồng bộ)
            // ─ Xoay view theo hướng tuyến (TwistAngle)
            // ─ Tỉ lệ 1:1 (ViewHeight = chiều cao thực)
            // ══════════════════════════════════════════════════════
            string baseLayoutName = form.SelectedAlignmentName;
            int maxVpPerLayout = form.MaxVpPerLayout;
            int totalLayouts = (int)Math.Ceiling((double)totalSegments / maxVpPerLayout);

            int segIndex = 0;
            int totalVpCreated = 0;

            ed.WriteMessage($"\n Bắt đầu tạo {totalLayouts} layout, mỗi layout tối đa {maxVpPerLayout} viewport...");

            using var progressVp = new Autodesk.AutoCAD.Runtime.ProgressMeter();
            progressVp.Start($"Đang tạo viewport trong Layout...");
            progressVp.SetLimit(totalSegments);

            for (int layoutIdx = 1; layoutIdx <= totalLayouts; layoutIdx++)
            {
                string layoutName = totalLayouts == 1
                    ? baseLayoutName
                    : $"{baseLayoutName}_{layoutIdx}";

                ObjectId layoutId = ObjectId.Null;

                // ── Tạo Layout mới ──
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        LayoutManager layoutMgr = LayoutManager.Current;
                        DBDictionary? layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                        string finalName = layoutName;
                        int suffix = 2;
                        while (layoutDict != null && layoutDict.Contains(finalName))
                        {
                            finalName = $"{layoutName}_{suffix}";
                            suffix++;
                        }
                        layoutName = finalName;

                        layoutId = layoutMgr.CreateLayout(layoutName);
                        layoutMgr.CurrentLayout = layoutName;
                        tr.Commit();
                    }
                    catch
                    {
                        ed.WriteMessage($"\n Không thể tạo Layout '{layoutName}'.");
                        continue;
                    }
                }

                if (layoutId == ObjectId.Null) continue;

                // ═══ Xử lý viewport theo batch trong mỗi layout ═══
                int vpInThisLayout = 0;
                int vpToCreate = Math.Min(maxVpPerLayout, totalSegments - segIndex);

                // Chia viewport trong layout thành các batch nhỏ
                for (int vpBatchStart = 0; vpBatchStart < vpToCreate; vpBatchStart += BATCH_SIZE)
                {
                    int vpBatchEnd = Math.Min(vpBatchStart + BATCH_SIZE, vpToCreate);
                    int vpBatchCount = vpBatchEnd - vpBatchStart;

                    // Danh sách VP đã tạo để set view ở transaction sau
                    var vpDataList = new List<(ObjectId VpId, double CenterX, double CenterY,
                        double VpH, double AngleRad, int STT, double AngleDeg, double VpW)>();

                    // ═══ TRANSACTION A: Tạo viewport trống + bật On ═══
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            Layout? layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                            if (layout == null) { tr.Commit(); continue; }

                            BlockTableRecord? paperBtr = tr.GetObject(
                                layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                            if (paperBtr == null) { tr.Commit(); continue; }

                            // Xóa viewport mặc định (VP#2+) — chỉ ở batch đầu tiên
                            if (vpBatchStart == 0)
                            {
                                List<ObjectId> toDelete = new();
                                foreach (ObjectId entId in paperBtr)
                                {
                                    var obj = tr.GetObject(entId, OpenMode.ForRead);
                                    if (obj is Viewport existingVp && existingVp.Number > 1)
                                        toDelete.Add(entId);
                                }
                                foreach (var delId in toDelete)
                                {
                                    var vpDel = tr.GetObject(delId, OpenMode.ForWrite) as Viewport;
                                    vpDel?.Erase();
                                }
                            }

                            double paperGap = 10;
                            double startX = 20, startY = 20;

                            for (int b = 0; b < vpBatchCount; b++)
                            {
                                int currentVpIdx = vpBatchStart + b;
                                var seg = form.ViewSegments[segIndex + currentVpIdx];

                                double vpW = viewWidthReal;
                                double vpH = viewHeightReal;
                                double vpX = startX + vpW / 2.0 + (vpInThisLayout + b) * (vpW + paperGap);
                                double vpY = startY + vpH / 2.0;

                                // Tạo viewport trống
                                Viewport vp = new Viewport();
                                paperBtr.AppendEntity(vp);
                                tr.AddNewlyCreatedDBObject(vp, true);

                                vp.CenterPoint = new Point3d(vpX, vpY, 0);
                                vp.Width = vpW;
                                vp.Height = vpH;
                                vp.On = true;

                                vpDataList.Add((vp.ObjectId, seg.CenterX, seg.CenterY,
                                    vpH, seg.AngleRad, seg.STT, seg.AngleDeg, vpW));

                                progressVp.MeterProgress();
                            }

                            tr.Commit();
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            ed.WriteMessage($"\n   Lỗi tạo VP batch: {ex.Message}");
                            continue;
                        }
                    }

                    // Cho AutoCAD xử lý
                    System.Windows.Forms.Application.DoEvents();

                    // ═══ TRANSACTION B: Set view cho từng VP ═══
                    if (vpDataList.Count > 0)
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            try
                            {
                                foreach (var vpData in vpDataList)
                                {
                                    Viewport? vp = tr.GetObject(vpData.VpId, OpenMode.ForWrite) as Viewport;
                                    if (vp == null) continue;

                                    vp.ViewTarget = new Point3d(vpData.CenterX, vpData.CenterY, 0);
                                    vp.ViewCenter = new Point2d(0, 0);
                                    vp.ViewHeight = vpData.VpH;
                                    vp.TwistAngle = -vpData.AngleRad;
                                    vp.Locked = true;

                                    ed.WriteMessage($"\n   VP#{vpData.STT}: Center=({vpData.CenterX:F1},{vpData.CenterY:F1}) " +
                                        $"Size={vpData.VpW:F1}x{vpData.VpH:F1} Angle={vpData.AngleDeg:F1}°");
                                }

                                tr.Commit();
                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                            {
                                ed.WriteMessage($"\n   Lỗi set view: {ex.Message}");
                            }
                        }
                    }

                    vpInThisLayout += vpBatchCount;
                    totalVpCreated += vpBatchCount;

                    // Cho AutoCAD xử lý giữa các batch
                    System.Windows.Forms.Application.DoEvents();
                }

                segIndex += vpInThisLayout;
                ed.WriteMessage($"\n Layout '{layoutName}': {vpInThisLayout} viewport hoàn tất");
            }

            progressVp.Stop();

            // Chuyển về paper space để xem layout
            db.TileMode = false;
            ed.WriteMessage($"\n Tổng: {totalVpCreated} viewport trong {totalLayouts} layout (1:1, khung theo CAD)");
        }

        private static void CreateLayerIfNotExist(string name, short color)
        {
            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                LayerTable? lt = tr.GetObject(A.Db.LayerTableId, OpenMode.ForWrite) as LayerTable;
                if (lt != null && !lt.Has(name))
                {
                    LayerTableRecord ltr = new()
                    {
                        Name = name,
                        Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                            Autodesk.AutoCAD.Colors.ColorMethod.ByAci, color)
                    };
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }
                tr.Commit();
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
        }

        /// <summary>
        /// Tạo hoặc lấy UCS đã có trong UcsTable.
        /// Dùng để gán UCS riêng cho từng viewport.
        /// </summary>
        private static ObjectId GetOrCreateUcs(Transaction tr, Database db,
            string ucsName, Point3d origin, Vector3d xAxis, Vector3d yAxis)
        {
            try
            {
                UcsTable? ut = tr.GetObject(db.UcsTableId, OpenMode.ForWrite) as UcsTable;
                if (ut == null) return ObjectId.Null;

                // Nếu đã tồn tại → cập nhật lại origin/axis
                if (ut.Has(ucsName))
                {
                    UcsTableRecord? existing = tr.GetObject(ut[ucsName], OpenMode.ForWrite) as UcsTableRecord;
                    if (existing != null)
                    {
                        existing.Origin = origin;
                        existing.XAxis = xAxis;
                        existing.YAxis = yAxis;
                    }
                    return ut[ucsName];
                }

                // Tạo mới
                UcsTableRecord utr = new()
                {
                    Name = ucsName,
                    Origin = origin,
                    XAxis = xAxis,
                    YAxis = yAxis
                };
                ut.Add(utr);
                tr.AddNewlyCreatedDBObject(utr, true);
                return utr.ObjectId;
            }
            catch
            {
                return ObjectId.Null;
            }
        }
    }

    #endregion
}
