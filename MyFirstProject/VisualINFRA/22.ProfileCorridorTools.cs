// 22.ProfileCorridorTools.cs - Công cụ Profile & Corridor từ VisualINFRA
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
using Autodesk.Civil.DatabaseServices.Styles;

using AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.VisualINFRA.ProfileCorridorTools))]

namespace MyFirstProject.VisualINFRA
{
    /// <summary>
    /// Công cụ Profile và Corridor - từ VisualINFRA
    /// </summary>
    public class ProfileCorridorTools
    {
        #region Profile Tools

        /// <summary>
        /// Tạo Profile View cho Alignment
        /// </summary>
        [CommandMethod("VI_CreateProfileView")]
        public static void CreateProfileView()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            var db = doc.Database;

            try
            {
                // Chọn Alignment
                var alignmentIds = civilDoc.GetAlignmentIds();
                if (alignmentIds.Count == 0)
                {
                    ed.WriteMessage("\n Không có Alignment.");
                    return;
                }

                ed.WriteMessage("\n\nDanh sách Alignment:");
                var alignments = new List<(int Index, ObjectId Id, string Name)>();

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    int idx = 1;
                    foreach (ObjectId id in alignmentIds)
                    {
                        var al = tr.GetObject(id, OpenMode.ForRead) as Alignment;
                        if (al != null)
                        {
                            ed.WriteMessage($"\n  {idx}. {al.Name}");
                            alignments.Add((idx, id, al.Name));
                            idx++;
                        }
                    }
                    tr.Commit();
                }

                var selResult = ed.GetInteger($"\nChọn Alignment (1-{alignments.Count}): ");
                if (selResult.Status != PromptStatus.OK) return;
                if (selResult.Value < 1 || selResult.Value > alignments.Count) return;

                var alignmentId = alignments[selResult.Value - 1].Id;
                string alignmentName = alignments[selResult.Value - 1].Name;

                // Chọn điểm chèn
                var ptResult = ed.GetPoint("\nChọn điểm chèn Profile View: ");
                if (ptResult.Status != PromptStatus.OK) return;
                var insertPoint = ptResult.Value;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                    if (alignment == null)
                    {
                        tr.Commit();
                        return;
                    }

                    // Lấy Profile View Style đầu tiên
                    var profileViewStyleId = GetFirstProfileViewStyle(civilDoc);
                    var bandSetStyleId = GetFirstBandSetStyle(civilDoc);

                    if (profileViewStyleId.IsNull || bandSetStyleId.IsNull)
                    {
                        ed.WriteMessage("\n Không tìm thấy Profile View Style hoặc Band Set Style.");
                        tr.Commit();
                        return;
                    }

                    // Tạo Profile View sử dụng API đúng
                    try
                    {
                        var profileViewId = ProfileView.Create(
                            alignmentId,
                            insertPoint,
                            $"PV_{alignmentName}",
                            bandSetStyleId,
                            profileViewStyleId
                        );

                        if (!profileViewId.IsNull)
                        {
                            ed.WriteMessage($"\n Đã tạo Profile View: PV_{alignmentName}");
                            VIUtilities.ZoomToEntity(profileViewId);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n Không thể tạo Profile View tự động: {ex.Message}");
                        ed.WriteMessage("\n   → Sử dụng lệnh PROFILEVIEWCREATE trong Civil 3D.");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Tạo Profile từ Surface cho tất cả Alignment
        /// </summary>
        [CommandMethod("VI_CreateMultiSurfaceProfile")]
        public static void CreateMultiSurfaceProfile()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            var db = doc.Database;

            try
            {
                // Lấy danh sách Surface
                var surfaceIds = civilDoc.GetSurfaceIds();
                if (surfaceIds.Count == 0)
                {
                    ed.WriteMessage("\n Không có Surface.");
                    return;
                }

                ed.WriteMessage("\n\nDanh sách Surface:");
                var surfaces = new List<(int Index, ObjectId Id, string Name)>();

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    int idx = 1;
                    foreach (ObjectId id in surfaceIds)
                    {
                        var surf = tr.GetObject(id, OpenMode.ForRead) as TinSurface;
                        if (surf != null)
                        {
                            ed.WriteMessage($"\n  {idx}. {surf.Name}");
                            surfaces.Add((idx, id, surf.Name));
                            idx++;
                        }
                    }
                    tr.Commit();
                }

                var surfResult = ed.GetInteger($"\nChọn Surface (1-{surfaces.Count}): ");
                if (surfResult.Status != PromptStatus.OK) return;
                if (surfResult.Value < 1 || surfResult.Value > surfaces.Count) return;

                var surfaceId = surfaces[surfResult.Value - 1].Id;
                string surfaceName = surfaces[surfResult.Value - 1].Name;

                // Tạo Profile cho tất cả Alignment
                var alignmentIds = civilDoc.GetAlignmentIds();
                int created = 0;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var profileStyleId = GetFirstProfileStyle(civilDoc);
                    var labelSetStyleId = GetFirstProfileLabelSetStyle(civilDoc);

                    foreach (ObjectId alId in alignmentIds)
                    {
                        var alignment = tr.GetObject(alId, OpenMode.ForRead) as Alignment;
                        if (alignment == null) continue;

                        try
                        {
                            string profileName = $"EG_{alignment.Name}_{surfaceName}";
                            
                            // Kiểm tra Profile đã tồn tại chưa
                            var existingProfileIds = alignment.GetProfileIds();
                            bool exists = false;
                            foreach (ObjectId pid in existingProfileIds)
                            {
                                var p = tr.GetObject(pid, OpenMode.ForRead) as Profile;
                                if (p != null && p.Name == profileName)
                                {
                                    exists = true;
                                    break;
                                }
                            }

                            if (!exists)
                            {
                                // Sử dụng API đúng cho Civil 3D 2026 (6 tham số)
                                // Lấy layer ID
                                var layerId = db.Clayer; // Sử dụng layer hiện tại
                                
                                var profileId = Profile.CreateFromSurface(
                                    profileName,
                                    alId,
                                    surfaceId,
                                    layerId,
                                    profileStyleId,
                                    labelSetStyleId
                                );
                                
                                if (!profileId.IsNull)
                                {
                                    created++;
                                    ed.WriteMessage($"\n  Tạo Profile: {profileName}");
                                }
                            }
                        }
                        catch
                        {
                            // Bỏ qua nếu không tạo được
                        }
                    }

                    ed.WriteMessage($"\n\n Đã tạo {created} Profile từ Surface {surfaceName}");
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Tạo Profile offset từ Profile hiện có
        /// </summary>
        [CommandMethod("VI_CreateOffsetProfile")]
        public static void CreateOffsetProfile()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                // Chọn Profile View
                var pvResult = ed.GetEntity("\nChọn Profile View: ");
                if (pvResult.Status != PromptStatus.OK) return;

                // Nhập offset
                var offsetResult = ed.GetDouble("\nNhập giá trị offset cao độ (m): ");
                if (offsetResult.Status != PromptStatus.OK) return;
                double offset = offsetResult.Value;

                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var profileView = tr.GetObject(pvResult.ObjectId, OpenMode.ForRead) as ProfileView;
                    if (profileView == null)
                    {
                        ed.WriteMessage("\n Đối tượng không phải Profile View.");
                        tr.Commit();
                        return;
                    }

                    // Lấy Alignment và Profile
                    var alignmentId = profileView.AlignmentId;
                    var alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                    
                    if (alignment == null)
                    {
                        tr.Commit();
                        return;
                    }

                    ed.WriteMessage($"\n\n Các Profile trong {alignment.Name}:");
                    var profileIds = alignment.GetProfileIds();
                    var profiles = new List<(int Index, ObjectId Id, string Name)>();
                    
                    int idx = 1;
                    foreach (ObjectId pid in profileIds)
                    {
                        var p = tr.GetObject(pid, OpenMode.ForRead) as Profile;
                        if (p != null)
                        {
                            ed.WriteMessage($"\n  {idx}. {p.Name}");
                            profiles.Add((idx, pid, p.Name));
                            idx++;
                        }
                    }

                    var selResult = ed.GetInteger($"\nChọn Profile gốc (1-{profiles.Count}): ");
                    if (selResult.Status != PromptStatus.OK || selResult.Value < 1 || selResult.Value > profiles.Count)
                    {
                        tr.Commit();
                        return;
                    }

                    var sourceProfileId = profiles[selResult.Value - 1].Id;
                    var sourceProfile = tr.GetObject(sourceProfileId, OpenMode.ForRead) as Profile;
                    
                    if (sourceProfile == null)
                    {
                        tr.Commit();
                        return;
                    }

                    // Tạo Profile mới bằng cách offset
                    // Lưu ý: Civil 3D không có hàm offset trực tiếp, cần tạo Layout Profile
                    ed.WriteMessage($"\n\n Tính năng này yêu cầu tạo Layout Profile.");
                    ed.WriteMessage($"\n   Profile gốc: {sourceProfile.Name}");
                    ed.WriteMessage($"\n   Offset: {offset:F3}m");
                    ed.WriteMessage($"\n   → Sử dụng Civil 3D UI để tạo Offset Profile.");

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Corridor Tools

        /// <summary>
        /// Tạo Surface từ Corridor
        /// </summary>
        [CommandMethod("VI_CreateSurfaceFromCorridor")]
        public static void CreateSurfaceFromCorridor()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            var db = doc.Database;

            try
            {
                // Lấy danh sách Corridor
                var corridorIds = civilDoc.CorridorCollection;
                if (corridorIds.Count == 0)
                {
                    ed.WriteMessage("\n Không có Corridor.");
                    return;
                }

                ed.WriteMessage("\n\nDanh sách Corridor:");
                var corridors = new List<(int Index, ObjectId Id, string Name)>();

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    int idx = 1;
                    foreach (ObjectId id in corridorIds)
                    {
                        var cor = tr.GetObject(id, OpenMode.ForRead) as Corridor;
                        if (cor != null)
                        {
                            ed.WriteMessage($"\n  {idx}. {cor.Name}");
                            corridors.Add((idx, id, cor.Name));
                            idx++;
                        }
                    }
                    tr.Commit();
                }

                var selResult = ed.GetInteger($"\nChọn Corridor (1-{corridors.Count}): ");
                if (selResult.Status != PromptStatus.OK) return;
                if (selResult.Value < 1 || selResult.Value > corridors.Count) return;

                var corridorId = corridors[selResult.Value - 1].Id;
                string corridorName = corridors[selResult.Value - 1].Name;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var corridor = tr.GetObject(corridorId, OpenMode.ForWrite) as Corridor;
                    if (corridor == null)
                    {
                        tr.Commit();
                        return;
                    }

                    // Tạo Surface từ corridor
                    string surfaceName = $"SURF_{corridorName}";
                    
                    // Kiểm tra Surface đã tồn tại
                    foreach (ObjectId sid in civilDoc.GetSurfaceIds())
                    {
                        var s = tr.GetObject(sid, OpenMode.ForRead) as TinSurface;
                        if (s != null && s.Name == surfaceName)
                        {
                            ed.WriteMessage($"\n Surface '{surfaceName}' đã tồn tại.");
                            tr.Commit();
                            return;
                        }
                    }

                    // Tạo Corridor Surface
                    var corSurfaces = corridor.CorridorSurfaces;
                    if (corSurfaces.Count == 0)
                    {
                        // Thêm corridor surface mới
                        try
                        {
                            corSurfaces.Add(surfaceName);
                            var newCorSurf = corSurfaces[surfaceName];
                            
                            // Thông báo cho người dùng cấu hình surface qua UI
                            ed.WriteMessage($"\n    Surface '{surfaceName}' đã được tạo.");
                            ed.WriteMessage($"\n   → Sử dụng Properties để thêm Point Codes/Links vào Surface.");
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\n Không thể thêm Surface: {ex.Message}");
                        }
                    }

                    corridor.Rebuild();
                    
                    ed.WriteMessage($"\n Đã tạo Surface từ Corridor: {surfaceName}");
                    ed.WriteMessage("\n   → Cập nhật Corridor Surface trong Properties.");

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Rebuild tất cả Corridor
        /// </summary>
        [CommandMethod("VI_RebuildAllCorridor")]
        public static void RebuildAllCorridor()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            var db = doc.Database;

            try
            {
                var corridorIds = civilDoc.CorridorCollection;
                if (corridorIds.Count == 0)
                {
                    ed.WriteMessage("\n Không có Corridor.");
                    return;
                }

                int rebuilt = 0;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in corridorIds)
                    {
                        var cor = tr.GetObject(id, OpenMode.ForWrite) as Corridor;
                        if (cor != null)
                        {
                            cor.Rebuild();
                            rebuilt++;
                            ed.WriteMessage($"\n   Rebuild: {cor.Name}");
                        }
                    }
                    tr.Commit();
                }

                ed.WriteMessage($"\n\n Đã rebuild {rebuilt} Corridor.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Hiển thị thông tin Corridor
        /// </summary>
        [CommandMethod("VI_CorridorInfo")]
        public static void CorridorInfo()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            var db = doc.Database;

            try
            {
                var corridorIds = civilDoc.CorridorCollection;
                
                ed.WriteMessage($"\n\n{'=',-70}");
                ed.WriteMessage($"\n THÔNG TIN CORRIDOR");
                ed.WriteMessage($"\n{'=',-70}");
                ed.WriteMessage($"\n{"Tên",-25} {"Baselines",-10} {"Regions",-10} {"Surfaces",-15}");
                ed.WriteMessage($"\n{new string('-', 70)}");

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in corridorIds)
                    {
                        var cor = tr.GetObject(id, OpenMode.ForRead) as Corridor;
                        if (cor != null)
                        {
                            int baselines = cor.Baselines.Count;
                            int regions = 0;
                            foreach (var bl in cor.Baselines)
                            {
                                regions += bl.BaselineRegions.Count;
                            }
                            int surfaces = cor.CorridorSurfaces.Count;

                            ed.WriteMessage($"\n{cor.Name,-25} {baselines,-10} {regions,-10} {surfaces,-15}");
                        }
                    }
                    tr.Commit();
                }

                ed.WriteMessage($"\n{'=',-70}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Helper Methods

        private static ObjectId GetFirstProfileViewStyle(CivilDocument civilDoc)
        {
            try
            {
                var styles = civilDoc.Styles.ProfileViewStyles;
                if (styles.Count > 0)
                    return styles[0];
            }
            catch { }
            return ObjectId.Null;
        }

        private static ObjectId GetFirstBandSetStyle(CivilDocument civilDoc)
        {
            try
            {
                var styles = civilDoc.Styles.ProfileViewBandSetStyles;
                if (styles.Count > 0)
                    return styles[0];
            }
            catch { }
            return ObjectId.Null;
        }

        private static ObjectId GetFirstProfileStyle(CivilDocument civilDoc)
        {
            try
            {
                var styles = civilDoc.Styles.ProfileStyles;
                if (styles.Count > 0)
                    return styles[0];
            }
            catch { }
            return ObjectId.Null;
        }

        private static ObjectId GetFirstProfileLabelSetStyle(CivilDocument civilDoc)
        {
            try
            {
                var styles = civilDoc.Styles.LabelSetStyles.ProfileLabelSetStyles;
                if (styles.Count > 0)
                    return styles[0];
            }
            catch { }
            return ObjectId.Null;
        }

        #endregion
    }
}
