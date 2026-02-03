// 23.VolumeCalculation.cs - Công cụ tính khối lượng từ VisualINFRA
// Viết lại cho AutoCAD 2026 / Civil 3D 2026

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using Autodesk.Civil;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.VisualINFRA.VolumeCalculation))]

namespace MyFirstProject.VisualINFRA
{
    /// <summary>
    /// Data class chứa thông tin khối lượng
    /// </summary>
    public class VolumeData
    {
        public string StationName { get; set; } = "";
        public double Station { get; set; }
        public double CutArea { get; set; }
        public double FillArea { get; set; }
        public double CutVolume { get; set; }
        public double FillVolume { get; set; }
        public Dictionary<string, double> MaterialAreas { get; set; } = new Dictionary<string, double>();
        public Dictionary<string, double> MaterialVolumes { get; set; } = new Dictionary<string, double>();
    }

    /// <summary>
    /// Công cụ tính khối lượng - từ VisualINFRA
    /// </summary>
    public class VolumeCalculation
    {
        #region Volume Civil Road - Tính KL đường

        /// <summary>
        /// Tính khối lượng đường theo Sample Line Group
        /// </summary>
        [CommandMethod("VI_VolumeCivilRoad")]
        public static void VolumeCivilRoad()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            var db = doc.Database;

            try
            {
                // Chọn Alignment
                var alignmentId = SelectAlignment(ed, civilDoc, db);
                if (alignmentId.IsNull) return;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                    if (alignment == null)
                    {
                        tr.Commit();
                        return;
                    }

                    // Lấy Sample Line Group
                    var slgIds = alignment.GetSampleLineGroupIds();
                    if (slgIds.Count == 0)
                    {
                        ed.WriteMessage("\n❌ Alignment chưa có Sample Line Group.");
                        tr.Commit();
                        return;
                    }

                    var slg = tr.GetObject(slgIds[0], OpenMode.ForRead) as SampleLineGroup;
                    if (slg == null)
                    {
                        tr.Commit();
                        return;
                    }

                    var slIds = slg.GetSampleLineIds();
                    var volumeDataList = new List<VolumeData>();

                    // Thu thập dữ liệu từ các Sample Line
                    foreach (ObjectId slId in slIds)
                    {
                        var sl = tr.GetObject(slId, OpenMode.ForRead) as SampleLine;
                        if (sl == null) continue;

                        var vd = new VolumeData
                        {
                            StationName = sl.Name,
                            Station = sl.Station,
                            // Khởi tạo diện tích mẫu (cần có Section View để tính chính xác)
                            CutArea = 0,
                            FillArea = 0
                        };

                        volumeDataList.Add(vd);
                    }

                    // Hiển thị thông tin Sample Line
                    ed.WriteMessage($"\n\n📊 SAMPLE LINE - {alignment.Name}");
                    ed.WriteMessage($"\n{'=',-60}");
                    ed.WriteMessage($"\n{"STT",-5} {"Tên Cọc",-20} {"Lý Trình",-15} {"Station",-15}");
                    ed.WriteMessage($"\n{new string('-', 60)}");

                    int stt = 1;
                    foreach (var vd in volumeDataList.OrderBy(x => x.Station))
                    {
                        ed.WriteMessage($"\n{stt,-5} {vd.StationName,-20} {FormatStation(vd.Station),-15} {vd.Station,-15:F3}");
                        stt++;
                    }

                    ed.WriteMessage($"\n{'=',-60}");
                    ed.WriteMessage($"\n\n⚠️ Để tính khối lượng chính xác, sử dụng lệnh VI_VolumeNetwork");
                    ed.WriteMessage($"\n   hoặc xem QS trong Civil 3D Compute Materials.");

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Hiển thị thông tin Section View
        /// </summary>
        [CommandMethod("VI_QuickVolume")]
        public static void QuickVolume()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // Chọn Section View
                var svResult = ed.GetEntity("\nChọn Section View: ");
                if (svResult.Status != PromptStatus.OK) return;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var sectionView = tr.GetObject(svResult.ObjectId, OpenMode.ForRead) as SectionView;
                    if (sectionView == null)
                    {
                        ed.WriteMessage("\n❌ Đối tượng không phải Section View.");
                        tr.Commit();
                        return;
                    }

                    ed.WriteMessage($"\n\n📊 THÔNG TIN SECTION VIEW");
                    ed.WriteMessage($"\n{'=',-50}");
                    ed.WriteMessage($"\n  Tên: {sectionView.Name}");

                    // Lấy Sample Line từ Section View
                    var sampleLineId = sectionView.SampleLineId;
                    if (!sampleLineId.IsNull)
                    {
                        var sampleLine = tr.GetObject(sampleLineId, OpenMode.ForRead) as SampleLine;
                        if (sampleLine != null)
                        {
                            ed.WriteMessage($"\n  Sample Line: {sampleLine.Name}");
                            ed.WriteMessage($"\n  Station: {sampleLine.Station:F3}");
                            ed.WriteMessage($"\n  Lý trình: {FormatStation(sampleLine.Station)}");
                        }
                    }

                    ed.WriteMessage($"\n{'=',-50}");
                    ed.WriteMessage($"\n\n  ⚠️ Để xem diện tích chi tiết, sử dụng Civil 3D Section Properties.");

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Volume Network - Tính KL mạng lưới

        /// <summary>
        /// Tính khối lượng theo mạng lưới (grid) - So sánh 2 Surface
        /// </summary>
        [CommandMethod("VI_VolumeNetwork")]
        public static void VolumeNetwork()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            var db = doc.Database;

            try
            {
                // Lấy 2 Surface để so sánh
                var surfaceIds = civilDoc.GetSurfaceIds();
                if (surfaceIds.Count < 2)
                {
                    ed.WriteMessage("\n❌ Cần ít nhất 2 Surface để tính khối lượng.");
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

                // Chọn Surface tự nhiên
                var baseResult = ed.GetInteger($"\nChọn Surface TỰ NHIÊN (1-{surfaces.Count}): ");
                if (baseResult.Status != PromptStatus.OK) return;
                if (baseResult.Value < 1 || baseResult.Value > surfaces.Count) return;
                var baseSurfaceId = surfaces[baseResult.Value - 1].Id;
                string baseName = surfaces[baseResult.Value - 1].Name;

                // Chọn Surface thiết kế
                var compResult = ed.GetInteger($"\nChọn Surface THIẾT KẾ (1-{surfaces.Count}): ");
                if (compResult.Status != PromptStatus.OK) return;
                if (compResult.Value < 1 || compResult.Value > surfaces.Count) return;
                var compSurfaceId = surfaces[compResult.Value - 1].Id;
                string compName = surfaces[compResult.Value - 1].Name;

                if (baseSurfaceId == compSurfaceId)
                {
                    ed.WriteMessage("\n❌ Phải chọn 2 Surface khác nhau.");
                    return;
                }

                // Tính khối lượng
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var baseSurf = tr.GetObject(baseSurfaceId, OpenMode.ForRead) as TinSurface;
                    var compSurf = tr.GetObject(compSurfaceId, OpenMode.ForRead) as TinSurface;

                    if (baseSurf == null || compSurf == null)
                    {
                        tr.Commit();
                        return;
                    }

                    ed.WriteMessage($"\n\n📊 TÍNH KHỐI LƯỢNG ĐÀO ĐẮP");
                    ed.WriteMessage($"\n{'=',-50}");
                    ed.WriteMessage($"\n  Surface tự nhiên: {baseName}");
                    ed.WriteMessage($"\n  Surface thiết kế: {compName}");
                    ed.WriteMessage($"\n{'=',-50}");

                    // Tạo Volume Surface
                    string volSurfName = $"VOL_{baseName}_{compName}";
                    
                    // Kiểm tra đã có Volume Surface chưa
                    TinVolumeSurface? volSurf = null;
                    foreach (ObjectId sid in surfaceIds)
                    {
                        var s = tr.GetObject(sid, OpenMode.ForRead);
                        if (s is TinVolumeSurface tvs && tvs.Name == volSurfName)
                        {
                            volSurf = tvs;
                            break;
                        }
                    }

                    if (volSurf == null)
                    {
                        // Tạo mới Volume Surface
                        try
                        {
                            var volSurfId = TinVolumeSurface.Create(volSurfName, baseSurfaceId, compSurfaceId);
                            volSurf = tr.GetObject(volSurfId, OpenMode.ForRead) as TinVolumeSurface;
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\n⚠️ Không thể tạo Volume Surface: {ex.Message}");
                        }
                    }

                    if (volSurf != null)
                    {
                        var props = volSurf.GetVolumeProperties();
                        
                        double cutVol = props.UnadjustedCutVolume;
                        double fillVol = props.UnadjustedFillVolume;
                        double netVol = cutVol - fillVol;
                        
                        ed.WriteMessage($"\n\n  📐 KẾT QUẢ:");
                        ed.WriteMessage($"\n  {'-',-45}");
                        ed.WriteMessage($"\n  Khối lượng ĐÀO (Cut):  {cutVol:N2} m³");
                        ed.WriteMessage($"\n  Khối lượng ĐẮP (Fill): {fillVol:N2} m³");
                        ed.WriteMessage($"\n  Khối lượng NET:        {netVol:N2} m³");
                        ed.WriteMessage($"\n  {'=',-50}");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Export Volume to Excel

        /// <summary>
        /// Xuất khối lượng ra file CSV
        /// </summary>
        [CommandMethod("VI_ExportVolume")]
        public static void ExportVolume()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            var db = doc.Database;

            try
            {
                var alignmentId = SelectAlignment(ed, civilDoc, db);
                if (alignmentId.IsNull) return;

                // Chọn file lưu
                var sfd = new System.Windows.Forms.SaveFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                    Title = "Lưu file khối lượng",
                    FileName = "Volume_Export.csv"
                };

                if (sfd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    return;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                    if (alignment == null)
                    {
                        tr.Commit();
                        return;
                    }

                    var slgIds = alignment.GetSampleLineGroupIds();
                    if (slgIds.Count == 0)
                    {
                        ed.WriteMessage("\n❌ Không có Sample Line Group.");
                        tr.Commit();
                        return;
                    }

                    var slg = tr.GetObject(slgIds[0], OpenMode.ForRead) as SampleLineGroup;
                    var slIds = slg!.GetSampleLineIds();

                    var sb = new StringBuilder();
                    sb.AppendLine("STT,TenCoc,Station,LyTrinh,X,Y");

                    int stt = 1;
                    foreach (ObjectId slId in slIds)
                    {
                        var sl = tr.GetObject(slId, OpenMode.ForRead) as SampleLine;
                        if (sl == null) continue;

                        // Lấy tọa độ tại tim tuyến
                        double x = 0, y = 0;
                        try
                        {
                            alignment.PointLocation(sl.Station, 0, ref x, ref y);
                        }
                        catch { continue; }

                        string lyTrinh = FormatStation(sl.Station);
                        sb.AppendLine($"{stt},{sl.Name},{sl.Station:F3},{lyTrinh},{x:F3},{y:F3}");
                        stt++;
                    }

                    File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                    ed.WriteMessage($"\n✅ Đã xuất {stt - 1} cọc ra file: {sfd.FileName}");

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Surface Volume Comparison

        /// <summary>
        /// So sánh khối lượng giữa các Surface
        /// </summary>
        [CommandMethod("VI_CompareSurfaceVolume")]
        public static void CompareSurfaceVolume()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            var db = doc.Database;

            try
            {
                var surfaceIds = civilDoc.GetSurfaceIds();
                
                ed.WriteMessage($"\n\n{'=',-70}");
                ed.WriteMessage($"\n📊 BẢNG SO SÁNH KHỐI LƯỢNG SURFACE");
                ed.WriteMessage($"\n{'=',-70}");

                // Lọc Volume Surface
                var volSurfaces = new List<(ObjectId Id, string Name, TinVolumeSurface Surf)>();

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in surfaceIds)
                    {
                        var surf = tr.GetObject(id, OpenMode.ForRead);
                        if (surf is TinVolumeSurface tvs)
                        {
                            volSurfaces.Add((id, tvs.Name, tvs));
                        }
                    }

                    if (volSurfaces.Count == 0)
                    {
                        ed.WriteMessage("\n❌ Không có Volume Surface. Hãy tạo bằng lệnh VI_VolumeNetwork.");
                        tr.Commit();
                        return;
                    }

                    ed.WriteMessage($"\n{"Tên Volume Surface",-30} {"Cut (m³)",-15} {"Fill (m³)",-15} {"Net (m³)",-15}");
                    ed.WriteMessage($"\n{new string('-', 75)}");

                    double totalCut = 0, totalFill = 0, totalNet = 0;

                    foreach (var vs in volSurfaces)
                    {
                        var props = vs.Surf.GetVolumeProperties();
                        double cut = props.UnadjustedCutVolume;
                        double fill = props.UnadjustedFillVolume;
                        double net = cut - fill;

                        ed.WriteMessage($"\n{vs.Name,-30} {cut,-15:N2} {fill,-15:N2} {net,-15:N2}");

                        totalCut += cut;
                        totalFill += fill;
                        totalNet += net;
                    }

                    ed.WriteMessage($"\n{new string('-', 75)}");
                    ed.WriteMessage($"\n{"TỔNG",-30} {totalCut,-15:N2} {totalFill,-15:N2} {totalNet,-15:N2}");
                    ed.WriteMessage($"\n{'=',-70}");

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Helper Methods

        private static ObjectId SelectAlignment(Editor ed, CivilDocument civilDoc, Database db)
        {
            var alignmentIds = civilDoc.GetAlignmentIds();
            if (alignmentIds.Count == 0)
            {
                ed.WriteMessage("\n❌ Không có Alignment.");
                return ObjectId.Null;
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

            var result = ed.GetInteger($"\nChọn Alignment (1-{alignments.Count}): ");
            if (result.Status != PromptStatus.OK) return ObjectId.Null;
            if (result.Value < 1 || result.Value > alignments.Count) return ObjectId.Null;

            return alignments[result.Value - 1].Id;
        }

        private static string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double m = station % 1000;
            return $"Km{km}+{m:F2}";
        }

        #endregion
    }
}
