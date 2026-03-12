using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using MyFirstProject.Extensions;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(MyFirstProject.AlignmentHangLoat))]

namespace MyFirstProject
{
    /// <summary>
    /// Lệnh tạo tuyến (Alignment) hàng loạt từ các Polyline có sẵn trong bản vẽ.
    /// - Chọn nhiều Polyline cùng lúc
    /// - Tự động đặt tên theo prefix + số thứ tự
    /// - Cho phép chọn Alignment Style, Label Set Style, Site
    /// </summary>
    class AlignmentHangLoat
    {
        // ═══════════════════════════════════════════════════════════════
        //  LỆNH CHÍNH: Tạo Alignment hàng loạt từ Polyline
        // ═══════════════════════════════════════════════════════════════
        [CommandMethod("CTA_TaoTuyenHangLoat")]
        public static void CTA_TaoTuyenHangLoat()
        {
            Editor ed = A.Ed;
            Database db = A.Db;
            CivilDocument cdoc = A.Cdoc;

            using Transaction tr = db.TransactionManager.StartTransaction();
            try
            {
                // ── 1. Chọn nhiều Polyline ──────────────────────────────
                ed.WriteMessage("\n╔════════════════════════════════════════════╗");
                ed.WriteMessage("\n║  TẠO TUYẾN HÀNG LOẠT TỪ POLYLINE          ║");
                ed.WriteMessage("\n╚════════════════════════════════════════════╝\n");

                var filter = new SelectionFilter(new[]
                {
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
                });

                ed.WriteMessage("\n⊙ Chọn các Polyline để tạo tuyến (Alignment):");
                PromptSelectionResult selResult = ed.GetSelection(filter);
                if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
                {
                    ed.WriteMessage("\n⚠️ Không chọn được Polyline nào. Hủy lệnh.");
                    return;
                }

                int soLuongPL = selResult.Value.Count;
                ed.WriteMessage($"\n✅ Đã chọn {soLuongPL} Polyline.");

                // ── 2. Nhập Prefix tên tuyến ────────────────────────────
                PromptStringOptions pso = new("\n⊙ Nhập tiền tố (prefix) tên tuyến [mặc định: Tuyen]: ")
                {
                    AllowSpaces = true,
                    DefaultValue = "Tuyen"
                };
                PromptResult prPrefix = ed.GetString(pso);
                if (prPrefix.Status != PromptStatus.OK) return;
                string prefix = string.IsNullOrWhiteSpace(prPrefix.StringResult) ? "Tuyen" : prPrefix.StringResult;

                // ── 3. Nhập số bắt đầu ─────────────────────────────────
                PromptIntegerOptions pio = new("\n⊙ Nhập số thứ tự bắt đầu [mặc định: 1]: ")
                {
                    DefaultValue = 1,
                    LowerLimit = 0,
                    UpperLimit = 9999
                };
                PromptIntegerResult prStart = ed.GetInteger(pio);
                if (prStart.Status != PromptStatus.OK) return;
                int startIndex = prStart.Value;

                // ── 4. Lấy danh sách Alignment Style ────────────────────
                ObjectId alignStyleId = ObjectId.Null;
                AlignmentStyleCollection styleCollection = cdoc.Styles.AlignmentStyles;
                if (styleCollection.Count > 0)
                {
                    ed.WriteMessage("\n─── Danh sách Alignment Style ───");
                    List<string> styleNames = new();
                    List<ObjectId> styleIdList = new();
                    int idx = 0;
                    foreach (ObjectId sId in styleCollection)
                    {
                        AlignmentStyle style = (AlignmentStyle)tr.GetObject(sId, OpenMode.ForRead);
                        styleNames.Add(style.Name);
                        styleIdList.Add(sId);
                        ed.WriteMessage($"\n  [{idx}] {style.Name}");
                        idx++;
                    }

                    PromptIntegerOptions pioStyle = new($"\n⊙ Chọn Alignment Style [0-{idx - 1}] [mặc định: 0]: ")
                    {
                        DefaultValue = 0,
                        LowerLimit = 0,
                        UpperLimit = idx - 1
                    };
                    PromptIntegerResult prStyle = ed.GetInteger(pioStyle);
                    if (prStyle.Status == PromptStatus.OK)
                        alignStyleId = styleIdList[prStyle.Value];
                    else
                        alignStyleId = styleIdList[0];
                }

                // ── 5. Lấy danh sách Alignment Label Set Style ──────────
                ObjectId labelSetId = ObjectId.Null;
                AlignmentLabelSetStyleCollection labelSetCollection = cdoc.Styles.LabelSetStyles.AlignmentLabelSetStyles;
                if (labelSetCollection.Count > 0)
                {
                    ed.WriteMessage("\n─── Danh sách Alignment Label Set Style ───");
                    List<ObjectId> labelIdList = new();
                    int idx = 0;
                    foreach (ObjectId lId in labelSetCollection)
                    {
                        StyleBase labelSet = (StyleBase)tr.GetObject(lId, OpenMode.ForRead);
                        labelIdList.Add(lId);
                        ed.WriteMessage($"\n  [{idx}] {labelSet.Name}");
                        idx++;
                    }

                    PromptIntegerOptions pioLabel = new($"\n⊙ Chọn Label Set Style [0-{idx - 1}] [mặc định: 0]: ")
                    {
                        DefaultValue = 0,
                        LowerLimit = 0,
                        UpperLimit = idx - 1
                    };
                    PromptIntegerResult prLabel = ed.GetInteger(pioLabel);
                    if (prLabel.Status == PromptStatus.OK)
                        labelSetId = labelIdList[prLabel.Value];
                    else
                        labelSetId = labelIdList[0];
                }

                // ── 6. Chọn Site (hoặc không) ───────────────────────────
                ObjectId siteId = ObjectId.Null; // Null = "No Site" (Siteless)
                ObjectIdCollection siteIds = cdoc.GetSiteIds();
                if (siteIds.Count > 0)
                {
                    ed.WriteMessage("\n─── Danh sách Site ───");
                    ed.WriteMessage("\n  [0] << Không chọn Site (Siteless) >>");
                    List<ObjectId> siteIdList = new() { ObjectId.Null };
                    int idx = 1;
                    foreach (ObjectId stId in siteIds)
                    {
                        Site site = (Site)tr.GetObject(stId, OpenMode.ForRead);
                        siteIdList.Add(stId);
                        ed.WriteMessage($"\n  [{idx}] {site.Name}");
                        idx++;
                    }

                    PromptIntegerOptions pioSite = new($"\n⊙ Chọn Site [0-{idx - 1}] [mặc định: 0 - Siteless]: ")
                    {
                        DefaultValue = 0,
                        LowerLimit = 0,
                        UpperLimit = idx - 1
                    };
                    PromptIntegerResult prSite = ed.GetInteger(pioSite);
                    if (prSite.Status == PromptStatus.OK)
                        siteId = siteIdList[prSite.Value];
                }

                // ── 7. Nhập Layer cho Alignment ─────────────────────────
                PromptStringOptions psoLayer = new("\n⊙ Nhập tên Layer cho tuyến [mặc định: C-ROAD-ALGN]: ")
                {
                    AllowSpaces = false,
                    DefaultValue = "C-ROAD-ALGN"
                };
                PromptResult prLayer = ed.GetString(psoLayer);
                string layerName = string.IsNullOrWhiteSpace(prLayer.StringResult) ? "C-ROAD-ALGN" : prLayer.StringResult;

                // Tạo layer nếu chưa có
                EnsureLayerExists(db, tr, layerName);

                // ── 8. Xác nhận trước khi tạo ──────────────────────────
                ed.WriteMessage("\n╔════════════════════════════════════════════╗");
                ed.WriteMessage($"\n║  SỐ LƯỢNG POLYLINE : {soLuongPL,-22}║");
                ed.WriteMessage($"\n║  PREFIX TÊN        : {prefix,-22}║");
                ed.WriteMessage($"\n║  SỐ BẮT ĐẦU       : {startIndex,-22}║");
                ed.WriteMessage($"\n║  LAYER             : {layerName,-22}║");
                ed.WriteMessage("\n╚════════════════════════════════════════════╝");

                PromptKeywordOptions pko = new("\n⊙ Xác nhận tạo tuyến? [Co/Khong] <Co>: ", "Co Khong")
                {
                    AllowNone = true
                };
                PromptResult prConfirm = ed.GetKeywords(pko);
                if (prConfirm.Status == PromptStatus.OK && prConfirm.StringResult == "Khong")
                {
                    ed.WriteMessage("\n⚠️ Đã hủy lệnh.");
                    return;
                }

                // ── 9. Tạo Alignment hàng loạt ─────────────────────────
                int successCount = 0;
                int failCount = 0;
                int currentIndex = startIndex;

                // Sắp xếp polyline theo vị trí (từ trái sang phải, trên xuống dưới)
                List<ObjectId> sortedPolyIds = SortPolylinesByPosition(selResult.Value, tr);

                foreach (ObjectId plId in sortedPolyIds)
                {
                    try
                    {
                        Polyline? pline = tr.GetObject(plId, OpenMode.ForRead) as Polyline;
                        if (pline == null) continue;

                        // Tạo tên tuyến
                        string alignName = $"{prefix}-{currentIndex:D3}";

                        // Kiểm tra tên trùng
                        alignName = GetUniqueAlignmentName(cdoc, tr, alignName);

                        // Tạo Alignment từ Polyline
                        // Chữ ký: Alignment.Create(CivilDocument, string name, ObjectId polylineId, ObjectId siteId, ObjectId styleId, ObjectId labelSetStyleId)
                        ObjectId newAlignId = Alignment.Create(
                            cdoc,
                            alignName,
                            plId,
                            siteId,
                            alignStyleId,
                            labelSetId
                        );

                        if (newAlignId != ObjectId.Null)
                        {
                            // Gán layer cho alignment
                            Alignment? newAlign = tr.GetObject(newAlignId, OpenMode.ForWrite) as Alignment;
                            if (newAlign != null)
                            {
                                newAlign.Layer = layerName;
                            }

                            successCount++;
                            ed.WriteMessage($"\n  ✅ [{currentIndex:D3}] Tạo tuyến: {alignName}");
                        }
                        else
                        {
                            failCount++;
                            ed.WriteMessage($"\n  ❌ [{currentIndex:D3}] Không thể tạo tuyến từ Polyline");
                        }

                        currentIndex++;
                    }
                    catch (System.Exception ex)
                    {
                        failCount++;
                        ed.WriteMessage($"\n  ❌ [{currentIndex:D3}] Lỗi: {ex.Message}");
                        currentIndex++;
                    }
                }

                tr.Commit();

                // ── 10. Báo cáo kết quả ────────────────────────────────
                ed.WriteMessage("\n╔════════════════════════════════════════════╗");
                ed.WriteMessage("\n║          KẾT QUẢ TẠO TUYẾN HÀNG LOẠT     ║");
                ed.WriteMessage("\n╠════════════════════════════════════════════╣");
                ed.WriteMessage($"\n║  ✅ Thành công : {successCount,-26}║");
                ed.WriteMessage($"\n║  ❌ Thất bại   : {failCount,-26}║");
                ed.WriteMessage($"\n║  📊 Tổng cộng  : {successCount + failCount,-26}║");
                ed.WriteMessage("\n╚════════════════════════════════════════════╝\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
                ed.WriteMessage($"\n   {ex.StackTrace}");
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  LỆNH PHỤ: Tạo Alignment hàng loạt - Chế độ nhanh
        //  (Dùng style mặc định, không hỏi nhiều)
        // ═══════════════════════════════════════════════════════════════
        [CommandMethod("CTA_TaoTuyenNhanh")]
        public static void CTA_TaoTuyenNhanh()
        {
            Editor ed = A.Ed;
            Database db = A.Db;
            CivilDocument cdoc = A.Cdoc;

            using Transaction tr = db.TransactionManager.StartTransaction();
            try
            {
                ed.WriteMessage("\n⊙ TẠO TUYẾN NHANH - Chọn các Polyline:");

                // Chọn polyline
                var filter = new SelectionFilter(new[]
                {
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
                });
                PromptSelectionResult selResult = ed.GetSelection(filter);
                if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
                {
                    ed.WriteMessage("\n⚠️ Không chọn được Polyline nào.");
                    return;
                }

                // Nhập prefix
                PromptStringOptions pso = new("\n⊙ Nhập tên prefix [mặc định: Tuyen]: ")
                {
                    AllowSpaces = true,
                    DefaultValue = "Tuyen"
                };
                PromptResult prPrefix = ed.GetString(pso);
                string prefix = string.IsNullOrWhiteSpace(prPrefix.StringResult) ? "Tuyen" : prPrefix.StringResult;

                // Lấy style mặc định (đầu tiên trong danh sách)
                ObjectId alignStyleId = GetFirstStyleId(cdoc.Styles.AlignmentStyles);
                ObjectId labelSetId = GetFirstStyleId(cdoc.Styles.LabelSetStyles.AlignmentLabelSetStyles);

                // Tạo layer nếu cần
                string layerName = "C-ROAD-ALGN";
                EnsureLayerExists(db, tr, layerName);

                // Sắp xếp và tạo
                List<ObjectId> sortedPolyIds = SortPolylinesByPosition(selResult.Value, tr);
                int successCount = 0;
                int index = 1;

                foreach (ObjectId plId in sortedPolyIds)
                {
                    try
                    {
                        string alignName = GetUniqueAlignmentName(cdoc, tr, $"{prefix}-{index:D3}");

                        ObjectId newAlignId = Alignment.Create(
                            cdoc,
                            alignName,
                            plId,
                            ObjectId.Null, // Siteless
                            alignStyleId,
                            labelSetId
                        );

                        if (newAlignId != ObjectId.Null)
                        {
                            // Gán layer
                            Alignment? newAlign = tr.GetObject(newAlignId, OpenMode.ForWrite) as Alignment;
                            if (newAlign != null) newAlign.Layer = layerName;

                            successCount++;
                            ed.WriteMessage($"\n  ✅ Tạo tuyến: {alignName}");
                        }
                        index++;
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n  ❌ Lỗi Polyline #{index}: {ex.Message}");
                        index++;
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n\n✅ Hoàn thành! Đã tạo {successCount}/{selResult.Value.Count} tuyến.\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  LỆNH PHỤ: Đổi tên Alignment hàng loạt
        // ═══════════════════════════════════════════════════════════════
        [CommandMethod("CTA_DoiTenTuyenHangLoat")]
        public static void CTA_DoiTenTuyenHangLoat()
        {
            Editor ed = A.Ed;
            CivilDocument cdoc = A.Cdoc;

            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                // Lấy tất cả alignment
                ObjectIdCollection alignIds = cdoc.GetAlignmentIds();
                if (alignIds.Count == 0)
                {
                    ed.WriteMessage("\n⚠️ Không có Alignment nào trong bản vẽ.");
                    return;
                }

                ed.WriteMessage("\n─── Danh sách tuyến hiện có ───");
                int idx = 0;
                List<ObjectId> alignList = new();
                foreach (ObjectId aId in alignIds)
                {
                    Alignment? algn = tr.GetObject(aId, OpenMode.ForRead) as Alignment;
                    if (algn != null)
                    {
                        alignList.Add(aId);
                        ed.WriteMessage($"\n  [{idx}] {algn.Name}");
                        idx++;
                    }
                }

                // Nhập prefix mới
                PromptStringOptions pso = new("\n⊙ Nhập prefix mới cho tên tuyến: ")
                {
                    AllowSpaces = true
                };
                PromptResult prPrefix = ed.GetString(pso);
                if (prPrefix.Status != PromptStatus.OK || string.IsNullOrWhiteSpace(prPrefix.StringResult)) return;
                string newPrefix = prPrefix.StringResult;

                // Đổi tên
                int count = 0;
                foreach (ObjectId aId in alignList)
                {
                    Alignment? algn = tr.GetObject(aId, OpenMode.ForWrite) as Alignment;
                    if (algn != null)
                    {
                        count++;
                        string oldName = algn.Name;
                        algn.Name = $"{newPrefix}-{count:D3}";
                        ed.WriteMessage($"\n  ✅ {oldName} → {algn.Name}");
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n\n✅ Đã đổi tên {count} tuyến.\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  HÀM HỖ TRỢ
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Đảm bảo layer tồn tại, nếu chưa thì tạo mới
        /// </summary>
        private static void EnsureLayerExists(Database db, Transaction tr, string layerName)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                LayerTableRecord ltr = new()
                {
                    Name = layerName
                };
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
                A.Ed.WriteMessage($"\n✅ Đã tạo layer mới: {layerName}");
            }
        }

        /// <summary>
        /// Lấy ObjectId đầu tiên từ một StyleCollection (dùng cho chế độ nhanh)
        /// </summary>
        private static ObjectId GetFirstStyleId(System.Collections.IEnumerable collection)
        {
            foreach (ObjectId id in collection)
            {
                return id;
            }
            return ObjectId.Null;
        }

        /// <summary>
        /// Sắp xếp Polyline theo vị trí: từ trái sang phải, trên xuống dưới
        /// (Ưu tiên theo X, sau đó theo Y giảm dần)
        /// </summary>
        private static List<ObjectId> SortPolylinesByPosition(SelectionSet ss, Transaction tr)
        {
            var polyDataList = new List<(ObjectId Id, double X, double Y)>();

            foreach (SelectedObject so in ss)
            {
                if (so == null) continue;
                Polyline? pline = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Polyline;
                if (pline == null) continue;

                // Dùng điểm đầu của Polyline làm vị trí tham chiếu
                Point3d startPt = pline.StartPoint;
                polyDataList.Add((so.ObjectId, startPt.X, startPt.Y));
            }

            // Sắp xếp: X tăng dần, nếu X bằng thì Y giảm dần
            polyDataList.Sort((a, b) =>
            {
                int cmpX = a.X.CompareTo(b.X);
                if (cmpX != 0) return cmpX;
                return b.Y.CompareTo(a.Y);
            });

            return polyDataList.Select(p => p.Id).ToList();
        }

        /// <summary>
        /// Tạo tên Alignment duy nhất (tránh trùng) 
        /// Nếu tên đã tồn tại, thêm hậu tố _1, _2, ...
        /// </summary>
        private static string GetUniqueAlignmentName(CivilDocument cdoc, Transaction tr, string baseName)
        {
            HashSet<string> existingNames = new(StringComparer.OrdinalIgnoreCase);
            ObjectIdCollection alignIds = cdoc.GetAlignmentIds();
            foreach (ObjectId aId in alignIds)
            {
                Alignment? algn = tr.GetObject(aId, OpenMode.ForRead) as Alignment;
                if (algn != null)
                    existingNames.Add(algn.Name);
            }

            if (!existingNames.Contains(baseName))
                return baseName;

            int suffix = 1;
            while (existingNames.Contains($"{baseName}_{suffix}"))
                suffix++;

            return $"{baseName}_{suffix}";
        }
    }
}
