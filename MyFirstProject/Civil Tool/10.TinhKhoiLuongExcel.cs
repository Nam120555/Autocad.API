// (C) Copyright 2024
// Tính khối lượng vật liệu từ Material Section và xuất ra Excel
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Acad = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using AcadDb = Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

using Civil = Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using Autodesk.Civil.Runtime;
using Autodesk.Civil.Settings;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Civil.ApplicationServices;
using CivSection = Autodesk.Civil.DatabaseServices.Section;
using FormLabel = System.Windows.Forms.Label;

using ClosedXML.Excel;
using MyFirstProject.Extensions;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(MyFirstProject.TinhKhoiLuongExcel))]

// Tắt cảnh báo nullable cho file lớn
#pragma warning disable CS8600, CS8602, CS8618

namespace MyFirstProject
{

    /// <summary>
    /// Thông tin diện tích vật liệu tại một cọc
    /// </summary>
    public class MaterialVolumeInfo
    {
        public string StakeName { get; set; } = "";       // Tên cọc: "Cọc 1"
        public string Station { get; set; } = "";          // Lý trình: "Km0+123.456"
        public double StationValue { get; set; }          // Giá trị lý trình: 123.456
        public string MaterialName { get; set; } = "";    // Tên vật liệu: "Đào đất"
        public double Area { get; set; }                  // Diện tích (m²)
    }

    /// <summary>
    /// Thông tin tổng hợp của một cọc
    /// </summary>
    public class StakeInfo
    {
        public string Station { get; set; } = "";         // Lý trình format: "Km0+123.456"
        public string StakeName { get; set; } = "";       // Tên cọc
        public double StationValue { get; set; }          // Giá trị lý trình số
        public double SpacingPrev { get; set; }           // Khoảng cách đến cọc trước
        
        // Dữ liệu theo nhóm
        public Dictionary<string, double> MaterialAreas { get; set; } = new();   // Vật liệu từ QTO
        public Dictionary<string, double> CorridorAreas { get; set; } = new();   // Shape từ Corridor
        public Dictionary<string, double> SurfaceAreas { get; set; } = new();    // Mặt địa hình
        public Dictionary<string, double> OtherAreas { get; set; } = new();      // Các nguồn khác
        
        // Khối lượng theo nhóm
        public Dictionary<string, double> MaterialVolumes { get; set; } = new();
        public Dictionary<string, double> CorridorVolumes { get; set; } = new();
        public Dictionary<string, double> SurfaceVolumes { get; set; } = new();
        public Dictionary<string, double> OtherVolumes { get; set; } = new();
        
        /// <summary>
        /// Lấy tất cả diện tích (gộp tất cả nhóm)
        /// </summary>
        public Dictionary<string, double> GetAllAreas()
        {
            var all = new Dictionary<string, double>(MaterialAreas);
            foreach (var kvp in CorridorAreas) all[kvp.Key] = kvp.Value;
            foreach (var kvp in SurfaceAreas) all[kvp.Key] = kvp.Value;
            foreach (var kvp in OtherAreas) all[kvp.Key] = kvp.Value;
            return all;
        }
        
        /// <summary>
        /// Lấy tất cả khối lượng (gộp tất cả nhóm)
        /// </summary>
        public Dictionary<string, double> GetAllVolumes()
        {
            var all = new Dictionary<string, double>(MaterialVolumes);
            foreach (var kvp in CorridorVolumes) all[kvp.Key] = kvp.Value;
            foreach (var kvp in SurfaceVolumes) all[kvp.Key] = kvp.Value;
            foreach (var kvp in OtherVolumes) all[kvp.Key] = kvp.Value;
            return all;
        }
        
        // Chi tiết Material Section Data (Left/Right Length, Min/Max Elevation, etc.)
        public Dictionary<string, MaterialSectionDetail> MaterialSectionDetails { get; set; } = new();
    }
    
    /// <summary>
    /// Chi tiết thông tin của Material Section (như trong Properties Panel)
    /// </summary>
    public class MaterialSectionDetail
    {
        public string MaterialName { get; set; } = "";
        public string SectionSurfaceName { get; set; } = "";
        
        // Phạm vi offset
        public double LeftLength { get; set; }          // Giá trị âm (bên trái tim)
        public double RightLength { get; set; }         // Giá trị dương (bên phải tim)
        public double TotalWidth => Math.Abs(LeftLength) + Math.Abs(RightLength);
        
        // Phạm vi cao độ
        public double MinElevation { get; set; }
        public double MaxElevation { get; set; }
        public double Height => MaxElevation - MinElevation;
        
        // Diện tích và điểm
        public double Area { get; set; }
        public int PointCount { get; set; }
        public List<Point3d> Points { get; set; } = new();
    }

    /// <summary>
    /// Thông tin Alignment có SampleLineGroup
    /// </summary>
    public class AlignmentInfo
    {
        public ObjectId AlignmentId { get; set; }
        public string Name { get; set; } = "";
        public ObjectId SampleLineGroupId { get; set; }
        public bool IsSelected { get; set; }
    }


    /// <summary>
    /// Phương pháp tính diện tích Material Section
    /// </summary>
    public enum AreaCalculationMethod
    {
        /// <summary>Sử dụng section.Area từ Civil 3D API (mặc định)</summary>
        CivilAPI,
        /// <summary>Tính từ SectionPoints bằng công thức Dây giày (Shoelace)</summary>
        SectionPoints,
        /// <summary>Ưu tiên SectionPoints, dự phòng API nếu không có điểm</summary>
        SectionPointsFirst
    }

    public partial class TinhKhoiLuongExcel
    {

        
        /// <summary>
        /// Số chữ số thập phân để làm tròn diện tích (m²)
        /// V3Tools sử dụng roundnb = 2
        /// </summary>
        public static int AreaDecimalPlaces { get; set; } = 2;
        
        /// <summary>
        /// Số chữ số thập phân để làm tròn khối lượng (m³)
        /// V3Tools sử dụng roundnb = 2
        /// </summary>
        public static int VolumeDecimalPlaces { get; set; } = 2;
        
        /// <summary>
        /// Phương pháp tính diện tích Material Section hiện tại
        /// Mặc định: SectionPointsFirst - Ưu tiên tính từ đường bao SectionPoints
        /// </summary>
        public static AreaCalculationMethod CurrentAreaMethod { get; set; } = AreaCalculationMethod.SectionPointsFirst;
        

        /// <summary>
        /// Tính diện tích đa giác sử dụng Công thức Dây giày (Shoelace Formula)
        /// </summary>
        public static double CalculatePolygonArea(List<Point2d> points)
        {
            if (points.Count < 3) return 0;

            double area = 0;
            int n = points.Count;
            for (int i = 0; i < n; i++)
            {
                int j = (i + 1) % n;
                area += points[i].X * points[j].Y;
                area -= points[j].X * points[i].Y;
            }
            return Math.Abs(area) / 2.0;
        }

        /// <summary>
        /// Tính khối lượng bằng phương pháp trung bình cộng (Average End Area)
        /// Volume = (S1 + S2) / 2 × L
        /// Công thức chuẩn Civil 3D - như V3Tools sử dụng
        /// </summary>
        /// <param name="areaPrev">Diện tích trắc ngang trước (m²)</param>
        /// <param name="areaCurrent">Diện tích trắc ngang hiện tại (m²)</param>
        /// <param name="spacing">Khoảng cách giữa 2 trắc ngang (m)</param>
        /// <param name="round">Có làm tròn hay không</param>
        /// <returns>Khối lượng (m³)</returns>
        public static double CalculateVolume(double areaPrev, double areaCurrent, double spacing, bool round = true)
        {
            double volume = (areaPrev + areaCurrent) / 2.0 * spacing;
            return round ? Math.Round(volume, VolumeDecimalPlaces) : volume;
        }
        
        /// <summary>
        /// Làm tròn diện tích theo cài đặt
        /// </summary>
        public static double RoundArea(double area)
        {
            return Math.Round(area, AreaDecimalPlaces);
        }

        /// <summary>
        /// Format lý trình theo dạng Km0+123.456
        /// </summary>
        public static string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double meters = station % 1000;
            return $"Km{km}+{meters:F3}";
        }

        /// <summary>
        /// Khoảng cách giữa các bảng khi xuất ra CAD (theo trục X)
        /// </summary>
        private static double TableSpacingX = 50.0;

        /// <summary>
        /// Chiều cao text mặc định trong bảng CAD
        /// </summary>
        private static double TableTextHeight = 3.0;

        /// <summary>
        /// Phương thức để TableSettingsForm cập nhật settings
        /// </summary>
        public static void SetTableSettings(double textHeight, double tableSpacing)
        {
            TableTextHeight = textHeight;
            TableSpacingX = tableSpacing;
        }

        /// <summary>
        /// Lệnh cài đặt khoảng cách giữa các bảng khi xuất CAD
        /// </summary>
        [CommandMethod("CTSV_CaiDatBang")]
        public static void CTSVCaiDatBang()
        {
            A.Ed.WriteMessage($"\n=== CÀI ĐẶT BẢNG KHỐI LƯỢNG ===");
            A.Ed.WriteMessage($"\nKhoảng cách hiện tại giữa các bảng (theo X): {TableSpacingX}");
            A.Ed.WriteMessage($"\nChiều cao text hiện tại: {TableTextHeight}");

            // Hỏi khoảng cách mới
            PromptDoubleOptions pdo = new("\nNhập khoảng cách giữa các bảng (theo X):")
            {
                AllowNegative = false,
                AllowZero = false,
                DefaultValue = TableSpacingX
            };
            PromptDoubleResult pdr = A.Ed.GetDouble(pdo);
            if (pdr.Status == PromptStatus.OK)
            {
                TableSpacingX = pdr.Value;
                A.Ed.WriteMessage($"\nĐã đặt khoảng cách giữa các bảng (theo X): {TableSpacingX}");
            }

            // Hỏi chiều cao text
            PromptDoubleOptions pdoText = new("\nNhập chiều cao text trong bảng:")
            {
                AllowNegative = false,
                AllowZero = false,
                DefaultValue = TableTextHeight
            };
            PromptDoubleResult pdrText = A.Ed.GetDouble(pdoText);
            if (pdrText.Status == PromptStatus.OK)
            {
                TableTextHeight = pdrText.Value;
                A.Ed.WriteMessage($"\nĐã đặt chiều cao text: {TableTextHeight}");
            }

            A.Ed.WriteMessage($"\n=== CÀI ĐẶT HOÀN TẤT ===");
        }

        /// <summary>
        /// Mở Form Taskbar nhỏ gọn
        /// </summary>
        [CommandMethod("CTSV_Taskbar")]
        public static void CTSVTaskbar()
        {
            VolumeTaskbar.ShowTaskbar();
        }

        /// <summary>
        /// Chỉ xuất bảng ra CAD (không xuất Excel)
        /// </summary>
        [CommandMethod("CTSV_XuatCad")]
        public static void CTSVXuatCadOnly()
        {
            try
            {
                A.Ed.WriteMessage("\n=== XUẤT BẢNG KHỐI LƯỢNG RA CAD ===");

                using Transaction tr = A.Db.TransactionManager.StartTransaction();

                // 1. Lấy danh sách Alignment có SampleLineGroup
                List<AlignmentInfo> alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông tìm thấy Alignment nào có SampleLineGroup!");
                    return;
                }

                // 2. Hiển thị form chọn Alignments
                FormChonAlignment formChon = new(alignments);
                if (formChon.ShowDialog() != DialogResult.OK || formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }

                // 3. Trích xuất dữ liệu
                Dictionary<ObjectId, List<StakeInfo>> alignmentData = new();
                HashSet<string> allMaterials = new();

                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    var stakeInfos = ExtractMaterialData(tr, alignInfo.SampleLineGroupId, alignInfo.AlignmentId);
                    alignmentData[alignInfo.AlignmentId] = stakeInfos;
                    
                    foreach (var stake in stakeInfos)
                    {
                        foreach (var mat in stake.MaterialAreas.Keys)
                            allMaterials.Add(mat);
                    }
                }

                if (allMaterials.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông tìm thấy dữ liệu vật liệu nào!");
                    return;
                }

                // 4. Sắp xếp vật liệu
                FormSapXepVatLieu formVatLieu = new(allMaterials.ToList());
                if (formVatLieu.ShowDialog() != DialogResult.OK)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }
                List<string> orderedMaterials = formVatLieu.OrderedMaterials;

                // 5. Tính khối lượng
                foreach (var kvp in alignmentData)
                {
                    CalculateVolumes(kvp.Value, orderedMaterials);
                }

                // 6. Vẽ bảng trong CAD
                PromptPointOptions ppo = new($"\nChọn điểm chèn bảng (các bảng tiếp theo cách nhau {TableSpacingX} đơn vị theo X):");
                ppo.AllowNone = false;
                PromptPointResult ppr = A.Ed.GetPoint(ppo);
                
                if (ppr.Status == PromptStatus.OK)
                {
                    Point3d currentInsertPoint = ppr.Value;
                    
                    foreach (var alignInfo in formChon.SelectedAlignments)
                    {
                        var stakeInfos = alignmentData[alignInfo.AlignmentId];
                        
                        int numCols = 3 + orderedMaterials.Count * 2;
                        double tableWidth = 25.0 + 25.0 + 15.0 + (numCols - 3) * 18.0;
                        
                        CreateCadTable(tr, currentInsertPoint, alignInfo.Name, stakeInfos, orderedMaterials);
                        A.Ed.WriteMessage($"\nĐã vẽ bảng cho '{alignInfo.Name}' tại ({currentInsertPoint.X:F2}, {currentInsertPoint.Y:F2})");
                        
                        currentInsertPoint = new Point3d(
                            currentInsertPoint.X + tableWidth + TableSpacingX, 
                            currentInsertPoint.Y, 
                            currentInsertPoint.Z);
                    }
                }
                
                tr.Commit();
                A.Ed.WriteMessage("\n=== HOÀN TẤT XUẤT CAD ===");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nLỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Lệnh so sánh 2 Surface để tính khối lượng đào đắp
        /// </summary>
        [CommandMethod("CTSV_SoSanhSurface")]
        public static void CTSVSoSanhSurface()
        {
            try
            {
                A.Ed.WriteMessage("\n\n=== SO SÁNH BỀ MẶT TÍNH KHỐI LƯỢNG ===");

                // Lấy danh sách Surface
                var surfaceIds = A.Cdoc.GetSurfaceIds();
                if (surfaceIds.Count < 2)
                {
                    A.Ed.WriteMessage("\n❌ Cần ít nhất 2 Surface để so sánh.");
                    return;
                }

                using Transaction tr = A.Db.TransactionManager.StartTransaction();

                // Liệt kê các Surface
                var surfaces = new List<(int Index, ObjectId Id, string Name, bool IsTin)>();
                int idx = 1;
                A.Ed.WriteMessage("\n\nDanh sách Surface có sẵn:");

                foreach (ObjectId id in surfaceIds)
                {
                    var surf = tr.GetObject(id, AcadDb.OpenMode.ForRead);
                    if (surf is TinSurface tinSurf)
                    {
                        A.Ed.WriteMessage($"\n  {idx}. {tinSurf.Name} (TIN Surface)");
                        surfaces.Add((idx, id, tinSurf.Name, true));
                        idx++;
                    }
                    else if (surf is TinVolumeSurface volSurf)
                    {
                        A.Ed.WriteMessage($"\n  {idx}. {volSurf.Name} (Volume Surface)");
                        surfaces.Add((idx, id, volSurf.Name, false));
                        idx++;
                    }
                }

                // Chọn Surface tự nhiên (Base)
                var baseResult = A.Ed.GetInteger($"\nChọn Surface TỰ NHIÊN (1-{surfaces.Count}): ");
                if (baseResult.Status != PromptStatus.OK) { tr.Commit(); return; }
                if (baseResult.Value < 1 || baseResult.Value > surfaces.Count) { tr.Commit(); return; }
                var baseSurface = surfaces[baseResult.Value - 1];

                // Chọn Surface thiết kế (Comparison)
                var compResult = A.Ed.GetInteger($"\nChọn Surface THIẾT KẾ (1-{surfaces.Count}): ");
                if (compResult.Status != PromptStatus.OK) { tr.Commit(); return; }
                if (compResult.Value < 1 || compResult.Value > surfaces.Count) { tr.Commit(); return; }
                var compSurface = surfaces[compResult.Value - 1];

                if (baseSurface.Id == compSurface.Id)
                {
                    A.Ed.WriteMessage("\n❌ Phải chọn 2 Surface khác nhau.");
                    tr.Commit();
                    return;
                }

                // Tạo hoặc lấy Volume Surface
                string volSurfName = $"VOL_{baseSurface.Name}_{compSurface.Name}";
                TinVolumeSurface? volumeSurface = null;

                // Kiểm tra xem đã có Volume Surface chưa
                foreach (ObjectId sid in surfaceIds)
                {
                    var s = tr.GetObject(sid, AcadDb.OpenMode.ForRead);
                    if (s is TinVolumeSurface tvs && tvs.Name == volSurfName)
                    {
                        volumeSurface = tvs;
                        A.Ed.WriteMessage($"\n✅ Sử dụng Volume Surface có sẵn: {volSurfName}");
                        break;
                    }
                }

                // Nếu chưa có, tạo mới
                if (volumeSurface == null && baseSurface.IsTin && compSurface.IsTin)
                {
                    try
                    {
                        var volSurfId = TinVolumeSurface.Create(volSurfName, baseSurface.Id, compSurface.Id);
                        volumeSurface = tr.GetObject(volSurfId, AcadDb.OpenMode.ForRead) as TinVolumeSurface;
                        A.Ed.WriteMessage($"\n✅ Đã tạo Volume Surface mới: {volSurfName}");
                    }
                    catch (System.Exception ex)
                    {
                        A.Ed.WriteMessage($"\n⚠️ Không thể tạo Volume Surface: {ex.Message}");
                    }
                }

                // Lấy thông tin khối lượng
                if (volumeSurface != null)
                {
                    var props = volumeSurface.GetVolumeProperties();

                    double cutVol = props.UnadjustedCutVolume;
                    double fillVol = props.UnadjustedFillVolume;
                    double netVol = cutVol - fillVol;

                    A.Ed.WriteMessage($"\n\n{'=',-60}");
                    A.Ed.WriteMessage($"\n📊 KẾT QUẢ SO SÁNH BỀ MẶT");
                    A.Ed.WriteMessage($"\n{'=',-60}");
                    A.Ed.WriteMessage($"\n  Surface tự nhiên: {baseSurface.Name}");
                    A.Ed.WriteMessage($"\n  Surface thiết kế: {compSurface.Name}");
                    A.Ed.WriteMessage($"\n{'-',-60}");
                    A.Ed.WriteMessage($"\n  Khối lượng ĐÀO (Cut):  {cutVol,15:N2} m³");
                    A.Ed.WriteMessage($"\n  Khối lượng ĐẮP (Fill): {fillVol,15:N2} m³");
                    A.Ed.WriteMessage($"\n  Khối lượng RÒNG (Net): {netVol,15:N2} m³");
                    A.Ed.WriteMessage($"\n{'=',-60}");

                    // Hỏi có muốn copy ra clipboard không
                    var copyResult = A.Ed.GetKeywords("\nCopy kết quả ra clipboard? [Yes/No] <Yes>: ", new[] { "Yes", "No" });
                    if (copyResult.Status != PromptStatus.OK || copyResult.StringResult != "No")
                    {
                        string clipboardText = $"SO SÁNH BỀ MẶT\n" +
                            $"Surface tự nhiên: {baseSurface.Name}\n" +
                            $"Surface thiết kế: {compSurface.Name}\n" +
                            $"Khối lượng ĐÀO (Cut): {cutVol:N2} m³\n" +
                            $"Khối lượng ĐẮP (Fill): {fillVol:N2} m³\n" +
                            $"Khối lượng RÒNG (Net): {netVol:N2} m³";
                        try
                        {
                            System.Windows.Forms.Clipboard.SetText(clipboardText);
                            A.Ed.WriteMessage("\n✅ Đã copy kết quả ra clipboard!");
                        }
                        catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                    }
                }

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Lấy diện tích vật liệu từ SectionViewGroup
        /// </summary>
        [CommandMethod("CTSV_LayDienTichTuSectionView")]
        public static void CTSVLayDienTichTuSectionView()
        {
            try
            {
                A.Ed.WriteMessage("\n\n=== LẤY DIỆN TÍCH TỪ SECTION VIEW ===");

                using Transaction tr = A.Db.TransactionManager.StartTransaction();

                // Lấy danh sách Alignment có SampleLineGroup
                var alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n❌ Không tìm thấy Alignment nào có SampleLineGroup.");
                    tr.Commit();
                    return;
                }

                // Hiển thị form chọn Alignment
                FormChonAlignment formChon = new(alignments);
                if (formChon.ShowDialog() != DialogResult.OK || formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    tr.Commit();
                    return;
                }

                var selectedAlign = formChon.SelectedAlignments[0];

                // Mở SampleLineGroup với ForWrite để truy cập SectionViewGroups
                SampleLineGroup? slg = tr.GetObject(selectedAlign.SampleLineGroupId, AcadDb.OpenMode.ForRead) as SampleLineGroup;
                if (slg == null)
                {
                    A.Ed.WriteMessage("\n❌ Không thể mở SampleLineGroup.");
                    tr.Commit();
                    return;
                }

                // Lấy SectionViewGroup(s) từ collection
                SectionViewGroupCollection svgCollection = slg.SectionViewGroups;
                if (svgCollection.Count == 0)
                {
                    A.Ed.WriteMessage("\n❌ Không có SectionViewGroup. Hãy tạo Section Views trước.");
                    tr.Commit();
                    return;
                }

                A.Ed.WriteMessage($"\n\n📊 DIỆN TÍCH TỪ SECTION VIEW - {selectedAlign.Name}");
                A.Ed.WriteMessage($"\n{'=',-80}");

                // Duyệt qua từng SectionViewGroup
                foreach (SectionViewGroup svg in svgCollection)
                {
                    if (svg == null) continue;

                    A.Ed.WriteMessage($"\n\n📁 SectionViewGroup: {svg.Name}");
                    A.Ed.WriteMessage($"\n{'-',-70}");

                    // Lấy các SectionView từ SectionViewGroup
                    ObjectIdCollection svIds = svg.GetSectionViewIds();

                    foreach (ObjectId svId in svIds)
                    {
                        SectionView? sectionView = tr.GetObject(svId, AcadDb.OpenMode.ForRead) as SectionView;
                        if (sectionView == null) continue;

                        // Lấy SampleLine tương ứng
                        ObjectId slId = sectionView.SampleLineId;
                        SampleLine? sampleLine = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                        if (sampleLine == null) continue;

                        A.Ed.WriteMessage($"\n  📍 {sampleLine.Name} ({FormatStation(sampleLine.Station)})");

                        // Lấy các Section trong SectionView
                        try
                        {
                            ObjectIdCollection sectionIds = sampleLine.GetSectionIds();
                            SectionSourceCollection sources = slg.GetSectionSources();

                            foreach (ObjectId sectionId in sectionIds)
                            {
                                CivSection? section = tr.GetObject(sectionId, AcadDb.OpenMode.ForRead, false, true) as CivSection;
                                if (section == null) continue;

                                // Tìm tên source
                                string sourceName = "Unknown";
                                string sourceType = "";
                                foreach (SectionSource source in sources)
                                {
                                    if (source.SourceId == section.SourceId)
                                    {
                                        sourceName = source.SourceName;
                                        sourceType = source.SourceType.ToString();
                                        break;
                                    }
                                }

                                double area = section.Area > 0 ? section.Area : CalculateSectionArea(section);
                                if (area > 0)
                                {
                                    A.Ed.WriteMessage($"\n      [{sourceType}] {sourceName}: {area:F4} m²");
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            A.Ed.WriteMessage($"\n      ⚠️ Lỗi: {ex.Message}");
                        }
                    }
                }

                A.Ed.WriteMessage($"\n\n{'=',-80}");
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }
    }
}
