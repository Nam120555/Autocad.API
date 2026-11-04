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


// Tắt cảnh báo nullable cho file lớn
#pragma warning disable CS8600, CS8602, CS8618

// Xuất khối lượng vật liệu theo Material Section (dùng bởi VolumeTaskbar) + hàm trích xuất / xuất Excel / bảng CAD
// (tách từ 10.TinhKhoiLuongExcel.cs — partial class TinhKhoiLuongExcel)
namespace MyFirstProject
{
    public partial class TinhKhoiLuongExcel
    {


        // [ĐÃ GỘP] CTSV_XuatKhoiLuong → Dùng CTSV_TinhKhoiLuongTracNgang (file 39)
        internal static void _REMOVED_CTSVXuatKhoiLuong()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                // 1. Lấy danh sách Alignments có SampleLineGroup
                List<AlignmentInfo> alignmentInfos = GetAlignmentsWithSampleLineGroups(tr);

                if (alignmentInfos.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông tìm thấy Alignment nào có SampleLineGroup!");
                    return;
                }

                // 2. Hiển thị form chọn Alignments
                FormChonAlignment formChon = new(alignmentInfos);
                if (formChon.ShowDialog() != DialogResult.OK || formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }

                // 3. Thu thập tất cả vật liệu từ các Alignment đã chọn
                HashSet<string> allMaterials = new();
                Dictionary<ObjectId, List<StakeInfo>> alignmentData = new();

                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    var stakeInfos = ExtractMaterialData(tr, alignInfo.SampleLineGroupId, alignInfo.AlignmentId);
                    alignmentData[alignInfo.AlignmentId] = stakeInfos;

                    foreach (var stake in stakeInfos)
                    {
                        foreach (var mat in stake.MaterialAreas.Keys)
                        {
                            allMaterials.Add(mat);
                        }
                    }
                }

                if (allMaterials.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông tìm thấy vật liệu nào trong các Alignment đã chọn!");
                    A.Ed.WriteMessage("\nĐảm bảo rằng bạn đã tạo Material Sections trong QTO Manager.");
                    return;
                }

                // 4. Hiển thị form sắp xếp thứ tự vật liệu
                FormSapXepVatLieu formSapXep = new(allMaterials.ToList());
                if (formSapXep.ShowDialog() != DialogResult.OK)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }

                List<string> orderedMaterials = formSapXep.OrderedMaterials;

                // 5. Hiển thị form cài đặt bảng thống kê
                if (!TableSettingsForm.ShowSettings())
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }

                // 6. Tính khối lượng cho từng Alignment
                foreach (var kvp in alignmentData)
                {
                    CalculateVolumes(kvp.Value, orderedMaterials);
                }

                // 6a. Tự động vẽ đường bao Material trong SectionView
                A.Ed.WriteMessage("\n\n=== VẼ ĐƯỜNG BAO MATERIAL SECTIONS ===");
                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    var stakeInfos = alignmentData[alignInfo.AlignmentId];
                    int totalBoundaries = DrawMaterialBoundaries(tr, alignInfo.SampleLineGroupId, stakeInfos, orderedMaterials);
                    A.Ed.WriteMessage($"\n✅ Đã vẽ {totalBoundaries} đường bao cho '{alignInfo.Name}'");
                }

                // 7. Chọn loại xuất (đường bao đã được vẽ tự động ở trên)
                PromptKeywordOptions pkoExport = new("\nChọn loại xuất [Excel/CAD/TracNgang/TatCa]", "Excel CAD TracNgang TatCa");
                pkoExport.Keywords.Default = "Excel";
                pkoExport.AllowNone = true;
                PromptResult prExport = A.Ed.GetKeywords(pkoExport);

                if (prExport.Status != PromptStatus.OK && prExport.Status != PromptStatus.None)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }

                string exportType = prExport.StringResult ?? "Excel";
                bool doExcel = exportType == "Excel" || exportType == "TatCa";
                bool doCad = exportType == "CAD" || exportType == "TatCa";
                bool doTracNgang = exportType == "TracNgang" || exportType == "TatCa";

                // 8. Xuất ra Excel nếu được chọn
                string excelPath = "";
                if (doExcel)
                {
                    SaveFileDialog saveDialog = new()
                    {
                        Filter = "Excel Files (*.xlsx)|*.xlsx",
                        Title = "Chọn nơi lưu file Excel khối lượng",
                        FileName = "KhoiLuongVatLieu.xlsx"
                    };

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        excelPath = saveDialog.FileName;
                        ExportToExcel(excelPath, formChon.SelectedAlignments, alignmentData, orderedMaterials, tr);
                        A.Ed.WriteMessage($"\n✅ Đã xuất file Excel: {excelPath}");
                    }
                }

                // 9. Xuất bảng thống kê theo trắc ngang nếu được chọn
                if (doTracNgang)
                {
                    // Thu thập dữ liệu thống kê Material
                    Dictionary<ObjectId, List<MaterialStatInfo>> alignmentMaterialStats = new();
                    
                    foreach (var alignInfo in formChon.SelectedAlignments)
                    {
                        var materialStats = ExtractMaterialStatistics(tr, alignInfo.SampleLineGroupId, alignInfo.AlignmentId);
                        alignmentMaterialStats[alignInfo.AlignmentId] = materialStats;
                    }

                    // Hỏi loại xuất: Excel hay CAD
                    PromptKeywordOptions pkTracNgang = new("\nXuất thống kê trắc ngang [Excel/CAD]", "Excel CAD");
                    pkTracNgang.Keywords.Default = "Excel";
                    pkTracNgang.AllowNone = true;
                    PromptResult prTracNgang = A.Ed.GetKeywords(pkTracNgang);

                    bool tracNgangExcel = prTracNgang.StringResult != "CAD";

                    if (tracNgangExcel)
                    {
                        SaveFileDialog saveDialogTN = new()
                        {
                            Filter = "Excel Files (*.xlsx)|*.xlsx",
                            Title = "Lưu file Excel thống kê trắc ngang",
                            FileName = $"ThongKeTracNgang_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                        };

                        if (saveDialogTN.ShowDialog() == DialogResult.OK)
                        {
                            ExportMaterialStatisticsToExcel(saveDialogTN.FileName, formChon.SelectedAlignments, 
                                alignmentMaterialStats, orderedMaterials);
                            A.Ed.WriteMessage($"\n✅ Đã xuất file Excel thống kê trắc ngang: {saveDialogTN.FileName}");
                        }
                    }
                    else
                    {
                        PromptPointResult pprTN = A.Ed.GetPoint("\nChọn điểm chèn bảng thống kê trắc ngang: ");
                        if (pprTN.Status == PromptStatus.OK)
                        {
                            Point3d insertPointTN = pprTN.Value;
                            
                            foreach (var alignInfo in formChon.SelectedAlignments)
                            {
                                var materialStats = alignmentMaterialStats[alignInfo.AlignmentId];
                                
                                CreateMaterialStatTable(tr, insertPointTN, alignInfo.Name, materialStats, orderedMaterials);
                                A.Ed.WriteMessage($"\n✅ Đã vẽ bảng thống kê trắc ngang cho '{alignInfo.Name}'");

                                // Offset cho bảng tiếp theo
                                double tableHeight = (materialStats.Count + 4) * 8.0;
                                insertPointTN = new Point3d(insertPointTN.X, insertPointTN.Y - tableHeight - 50, insertPointTN.Z);
                            }
                        }
                    }
                }

                // 10. Vẽ bảng trong CAD nếu được chọn
                if (doCad)
                {
                    PromptPointOptions ppo = new($"\nChọn điểm chèn bảng khối lượng (các bảng tiếp theo sẽ cách nhau {TableSpacingX} đơn vị theo X):");
                    ppo.AllowNone = false;
                    PromptPointResult ppr = A.Ed.GetPoint(ppo);
                    
                    if (ppr.Status == PromptStatus.OK)
                    {
                        Point3d currentInsertPoint = ppr.Value;
                        int tableIndex = 0;
                        
                        // Vẽ bảng cho từng Alignment
                        foreach (var alignInfo in formChon.SelectedAlignments)
                        {
                            var stakeInfos = alignmentData[alignInfo.AlignmentId];
                            
                            // Tính chiều rộng bảng để offset cho bảng tiếp theo
                            int numCols = 3 + orderedMaterials.Count * 2;
                            double tableWidth = 25.0 + 25.0 + 15.0 + (numCols - 3) * 18.0; // Cột 1,2 = 25, Cột 3 = 15, còn lại = 18
                            
                            CreateCadTable(tr, currentInsertPoint, alignInfo.Name, stakeInfos, orderedMaterials);
                            A.Ed.WriteMessage($"\n✅ Đã vẽ bảng cho '{alignInfo.Name}' tại ({currentInsertPoint.X:F2}, {currentInsertPoint.Y:F2})");
                            
                            // Di chuyển điểm chèn sang phải cho bảng tiếp theo
                            currentInsertPoint = new Point3d(
                                currentInsertPoint.X + tableWidth + TableSpacingX, 
                                currentInsertPoint.Y, 
                                currentInsertPoint.Z);
                            
                            tableIndex++;
                        }
                    }
                }

                // 11. (Đường bao đã được vẽ tự động ở bước 6a)

                // 12. Hỏi mở file Excel nếu có
                if (!string.IsNullOrEmpty(excelPath))
                {
                    if (MessageBox.Show("Bạn có muốn mở file Excel?", "Hoàn thành", 
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = excelPath,
                            UseShellExecute = true
                        });
                    }
                }
                
                tr.Commit();
                A.Ed.WriteMessage("\n\n=== HOÀN THÀNH ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nLỗi: {ex.Message}");
                A.Ed.WriteMessage($"\nStack: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// Lấy danh sách Alignments có SampleLineGroup
        /// </summary>
        private static List<AlignmentInfo> GetAlignmentsWithSampleLineGroups(Transaction tr)
        {
            List<AlignmentInfo> result = new();

            ObjectIdCollection alignmentIds = A.Cdoc.GetAlignmentIds();
            foreach (ObjectId alignmentId in alignmentIds)
            {
                Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
                if (alignment == null) continue;

                ObjectIdCollection slgIds = alignment.GetSampleLineGroupIds();
                if (slgIds.Count > 0)
                {
                    result.Add(new AlignmentInfo
                    {
                        AlignmentId = alignmentId,
                        Name = alignment.Name,
                        SampleLineGroupId = slgIds[0], // Lấy SampleLineGroup đầu tiên
                        IsSelected = true
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// Trích xuất dữ liệu vật liệu từ QTO Material List trong SampleLineGroup
        /// </summary>
        private static List<StakeInfo> ExtractMaterialData(Transaction tr, ObjectId sampleLineGroupId, ObjectId alignmentId)
        {
            List<StakeInfo> stakeInfos = new();

            // Mở SampleLineGroup với ForWrite để có thể truy cập MaterialLists
            SampleLineGroup? slg = tr.GetObject(sampleLineGroupId, AcadDb.OpenMode.ForRead, false, true) as SampleLineGroup;
            if (slg == null) return stakeInfos;

            Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
            if (alignment == null) return stakeInfos;

            // Lấy Material Lists từ SampleLineGroup
            QTOMaterialListCollection materialLists = slg.MaterialLists;
            
            if (materialLists.Count == 0)
            {
                A.Ed.WriteMessage("\nKhông tìm thấy Material List nào trong SampleLineGroup!");
                A.Ed.WriteMessage("\nVui lòng tạo Material List trong QTO Manager trước.");
                return stakeInfos;
            }

            // Tạo danh sách QTOMaterial với tên, GUID và MaterialList GUID
            List<(string Name, Guid MaterialListGuid, Guid MaterialGuid, QTOMaterial Material)> materials = new();
            foreach (QTOMaterialList materialList in materialLists)
            {
                try
                {
                    Guid listGuid = materialList.Guid;
                    foreach (QTOMaterial material in materialList)
                    {
                        materials.Add((material.Name, listGuid, material.Guid, material));
                        // A.Ed.WriteMessage($"\n  Material: '{material.Name}', QuantityType: {material.QuantityType}");
                    }
                }
                catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
            }

            // Lấy danh sách SampleLine và sắp xếp theo lý trình
            List<SampleLine> sortedSampleLines = new List<SampleLine>();
            foreach (ObjectId slId in slg.GetSampleLineIds())
            {
                var sl = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                if (sl != null) sortedSampleLines.Add(sl);
            }
            sortedSampleLines = sortedSampleLines.OrderBy(s => s.Station).ToList();

            // Duyệt qua từng SampleLine đã sắp xếp
            double prevStation = 0;
            bool isFirst = true;

            foreach (SampleLine sampleLine in sortedSampleLines)
            {
                StakeInfo stakeInfo = new()
                {
                    StakeName = sampleLine.Name,
                    StationValue = sampleLine.Station,
                    Station = FormatStation(sampleLine.Station),
                    SpacingPrev = isFirst ? 0 : (sampleLine.Station - prevStation)
                };

                // Lấy MaterialSection cho từng Material
                foreach (var (materialName, materialListGuid, materialGuid, material) in materials)
                {
                    try
                    {
                        // Lấy MaterialSection ID từ SampleLine sử dụng MaterialList GUID và Material GUID
                        ObjectId materialSectionId = sampleLine.GetMaterialSectionId(materialListGuid, materialGuid);
                        
                        if (!materialSectionId.IsNull && materialSectionId.IsValid)
                        {
                            // Thử lấy MaterialSection
                            AcadDb.DBObject? sectionObj = tr.GetObject(materialSectionId, AcadDb.OpenMode.ForRead, false, true);
                            
                            if (sectionObj != null)
                            {
                                // Nếu là Section, lấy diện tích từ Civil 3D API
                                if (sectionObj is CivSection section)
                                {
                                    // ═══════════════════════════════════════════════════════════
                                    // DỮ LIỆU TỪ COMPUTE MATERIAL:
                                    // - Tên Material: QTOMaterial.Name
                                    // - Diện tích: Tính theo phương pháp CurrentAreaMethod
                                    //   + CivilAPI: Từ section.Area (tương đương Properties Panel)
                                    //   + SectionPoints: Tính từ đường bao SectionPoints (Shoelace)
                                    //   + SectionPointsFirst: Ưu tiên SectionPoints, fallback API
                                    // ═══════════════════════════════════════════════════════════
                                    
                                    // Lấy diện tích theo phương pháp đã chọn
                                    double area = GetMaterialSectionArea(section, CurrentAreaMethod);
                                    
                                    // Đếm số điểm trong section
                                    int pointCount = 0;
                                    try { pointCount = section.SectionPoints.Count; } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                                    
                                    // Log chi tiết cho debugging
                                    if (isFirst)
                                    {
                                        double apiArea = section.Area;
                                        double shoelaceArea = pointCount >= 3 ? CalculateSectionArea(section) : 0;
                                        
                                        A.Ed.WriteMessage($"\n  ┌───────────────────────────────────────────────────────────┐");
                                        A.Ed.WriteMessage($"\n  │ 📍 SampleLine: {sampleLine.Name,-40} │");
                                        A.Ed.WriteMessage($"\n  ├───────────────────────────────────────────────────────────┤");
                                        A.Ed.WriteMessage($"\n  │ 📋 Material: {materialName,-44} │");
                                        A.Ed.WriteMessage($"\n  │   → Phương pháp: {CurrentAreaMethod,-38} │");
                                        A.Ed.WriteMessage($"\n  │   → AREA (API):        {apiArea,12:F4} m²              │");
                                        A.Ed.WriteMessage($"\n  │   → AREA (Shoelace):   {shoelaceArea,12:F4} m²              │");
                                        A.Ed.WriteMessage($"\n  │   → AREA (đã chọn):    {area,12:F4} m²              │");
                                        A.Ed.WriteMessage($"\n  │   → Số điểm SectionPoints: {pointCount,8}                   │");
                                        
                                        if (area == 0 && pointCount == 0)
                                        {
                                            A.Ed.WriteMessage($"\n  │ ⚠️ SECTION CHƯA ĐƯỢC COMPUTE!                            │");
                                            A.Ed.WriteMessage($"\n  │   → Hãy: Analyze > Compute Materials                     │");
                                        }
                                        else if (Math.Abs(apiArea - shoelaceArea) > 0.01 && apiArea > 0 && shoelaceArea > 0)
                                        {
                                            double diff = Math.Abs(apiArea - shoelaceArea) / apiArea * 100;
                                            A.Ed.WriteMessage($"\n  │ ⚠️ Sai lệch API vs Shoelace: {diff,5:F2}%                     │");
                                        }
                                        else
                                        {
                                            A.Ed.WriteMessage($"\n  │   ✅ Khớp với Properties Panel / Data Section            │");
                                        }
                                        A.Ed.WriteMessage($"\n  └───────────────────────────────────────────────────────────┘");
                                    }
                                    
                                    // Lưu diện tích (kể cả khi = 0 để tracking)
                                    if (!stakeInfo.MaterialAreas.ContainsKey(materialName))
                                        stakeInfo.MaterialAreas[materialName] = 0;
                                    stakeInfo.MaterialAreas[materialName] += area;
                                    
                                    // === LẤY CHI TIẾT MATERIAL SECTION DATA ===
                                    double minOffset = double.MaxValue;
                                    double maxOffset = double.MinValue;
                                    double minElevation = double.MaxValue;
                                    double maxElevation = double.MinValue;
                                    List<Point3d> points = new();
                                    
                                    try
                                    {
                                        foreach (SectionPoint pt in section.SectionPoints)
                                        {
                                            points.Add(pt.Location);
                                            
                                            // Offset (X): âm = trái, dương = phải
                                            if (pt.Location.X < minOffset) minOffset = pt.Location.X;
                                            if (pt.Location.X > maxOffset) maxOffset = pt.Location.X;
                                            
                                            // Elevation (Y)
                                            if (pt.Location.Y < minElevation) minElevation = pt.Location.Y;
                                            if (pt.Location.Y > maxElevation) maxElevation = pt.Location.Y;
                                        }
                                    }
                                    catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                                    
                                    // Tạo và lưu MaterialSectionDetail
                                    MaterialSectionDetail detail = new()
                                    {
                                        MaterialName = materialName,
                                        SectionSurfaceName = section.Name,
                                        LeftLength = minOffset != double.MaxValue ? minOffset : 0,
                                        RightLength = maxOffset != double.MinValue ? maxOffset : 0,
                                        MinElevation = minElevation != double.MaxValue ? minElevation : 0,
                                        MaxElevation = maxElevation != double.MinValue ? maxElevation : 0,
                                        Area = area,
                                        PointCount = points.Count,
                                        Points = points
                                    };
                                    stakeInfo.MaterialSectionDetails[materialName] = detail;
                                }
                            }
                        }
                        else if (isFirst)
                        {
                            A.Ed.WriteMessage($"\n  SampleLine '{sampleLine.Name}' - Material '{materialName}': Không có MaterialSection");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        if (isFirst)
                        {
                            A.Ed.WriteMessage($"\n  Lỗi khi lấy Material '{materialName}': {ex.Message}");
                        }
                    }
                }

                // Lấy các Section khác từ SampleLine (Corridor Shapes, Surfaces, v.v.)
                try
                {
                    ObjectIdCollection sectionIds = sampleLine.GetSectionIds();
                    SectionSourceCollection sectionSources = slg.GetSectionSources();
                    
                    // Tạo HashSet để tránh trùng lặp
                    HashSet<string> processedSources = new();
                    
                    foreach (ObjectId sectionId in sectionIds)
                    {
                        try
                        {
                            CivSection? section = tr.GetObject(sectionId, AcadDb.OpenMode.ForRead, false, true) as CivSection;
                            if (section == null) continue;

                            // Tìm SectionSource tương ứng
                            foreach (SectionSource source in sectionSources)
                            {
                                if (source.SourceId == section.SourceId)
                                {
                                    // Bỏ qua nếu là Material (đã xử lý ở trên)
                                    if (source.SourceType == SectionSourceType.Material)
                                        break;
                                    
                                    // Tạo tên duy nhất cho source
                                    string sourceName = source.SourceName;
                                    string sourceKey = $"{source.SourceType}_{sourceName}";
                                    
                                    // Bỏ qua nếu đã xử lý
                                    if (processedSources.Contains(sourceKey))
                                        break;
                                    processedSources.Add(sourceKey);
                                    
                                    // Tính diện tích
                                    double area = CalculateSectionArea(section);
                                    
                                    // Đặt tên hiển thị theo SourceType
                                    string displayName = sourceName;
                                    
                                    if (isFirst)
                                    {
                                        A.Ed.WriteMessage($"\n  SampleLine '{sampleLine.Name}' - {source.SourceType} '{sourceName}': Area = {area:F4} m²");
                                    }
                                    
                                    if (area > 0)
                                    {
                                        // Phân loại vào đúng nhóm theo SourceType
                                        if (source.SourceType == SectionSourceType.Corridor)
                                        {
                                            if (!stakeInfo.CorridorAreas.ContainsKey(displayName))
                                                stakeInfo.CorridorAreas[displayName] = 0;
                                            stakeInfo.CorridorAreas[displayName] += area;
                                        }
                                        else if (source.SourceType == SectionSourceType.TinSurface)
                                        {
                                            if (!stakeInfo.SurfaceAreas.ContainsKey(displayName))
                                                stakeInfo.SurfaceAreas[displayName] = 0;
                                            stakeInfo.SurfaceAreas[displayName] += area;
                                        }
                                        else
                                        {
                                            if (!stakeInfo.OtherAreas.ContainsKey(displayName))
                                                stakeInfo.OtherAreas[displayName] = 0;
                                            stakeInfo.OtherAreas[displayName] += area;
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                        catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                    }
                }
                catch (System.Exception ex)
                {
                    if (isFirst)
                    {
                        A.Ed.WriteMessage($"\n  Lỗi khi lấy Section sources: {ex.Message}");
                    }
                }

                stakeInfos.Add(stakeInfo);
                prevStation = sampleLine.Station;
                isFirst = false;
            }

            // Debug: hiển thị tổng kết dữ liệu
            if (stakeInfos.Count > 0)
            {
                var allMaterials = stakeInfos.SelectMany(s => s.MaterialAreas.Keys).Distinct().ToList();
                A.Ed.WriteMessage($"\n\nTổng cộng {stakeInfos.Count} sample lines, {allMaterials.Count} vật liệu có dữ liệu: {string.Join(", ", allMaterials)}");
                
                // Tính tổng diện tích cho mỗi material
                foreach (var mat in allMaterials)
                {
                    double totalArea = stakeInfos.Sum(s => s.MaterialAreas.ContainsKey(mat) ? s.MaterialAreas[mat] : 0);
                    A.Ed.WriteMessage($"\n  - {mat}: Tổng diện tích = {totalArea:F4} m²");
                }
            }

            return stakeInfos;
        }

        /// <summary>
        /// Tính diện tích Section từ SectionPoints (công thức Shoelace)
        /// Không sử dụng section.Area để có thể so sánh độc lập
        /// </summary>
        private static double CalculateSectionArea(CivSection section)
        {
            try
            {
                SectionPointCollection points = section.SectionPoints;
                if (points.Count < 3) return 0;

                List<Point2d> pointList = new();
                foreach (SectionPoint pt in points)
                {
                    pointList.Add(new Point2d(pt.Location.X, pt.Location.Y));
                }

                // Sắp xếp điểm theo thứ tự vòng quanh để Shoelace hoạt động đúng
                pointList = SortPointsByAngle(pointList);

                // Đóng đa giác nếu chưa đóng
                if (pointList.Count >= 2)
                {
                    // Thêm điểm đóng về đường cơ sở
                    Point2d first = pointList[0];
                    Point2d last = pointList[pointList.Count - 1];
                    
                    // Nếu section là hở (đường), cần đóng về baseline
                    if (Math.Abs(first.X - last.X) > 0.001 || Math.Abs(first.Y - last.Y) > 0.001)
                    {
                        // Đóng đa giác dọc theo đường Y=0 hoặc Y min
                        double yBase = Math.Min(first.Y, last.Y);
                        pointList.Add(new Point2d(last.X, yBase));
                        pointList.Add(new Point2d(first.X, yBase));
                    }
                }

                return CalculatePolygonArea(pointList);
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Sắp xếp các điểm theo thứ tự góc từ tâm (angular sort)
        /// Đảm bảo điểm nằm theo chiều kim đồng hồ hoặc ngược kim đồng hồ
        /// </summary>
        private static List<Point2d> SortPointsByAngle(List<Point2d> points)
        {
            if (points.Count < 3) return points;

            // Tính tâm (centroid) của đa giác
            double cx = points.Average(p => p.X);
            double cy = points.Average(p => p.Y);

            // Sắp xếp theo góc từ tâm (ngược chiều kim đồng hồ)
            return points.OrderBy(p => Math.Atan2(p.Y - cy, p.X - cx)).ToList();
        }

        /// <summary>
        /// Lấy diện tích Material Section theo phương pháp đã chọn
        /// Hỗ trợ 3 phương pháp: CivilAPI, SectionPoints, SectionPointsFirst
        /// </summary>
        /// <param name="section">Material Section từ Civil 3D</param>
        /// <param name="method">Phương pháp tính diện tích</param>
        /// <returns>Diện tích (m²)</returns>
        private static double GetMaterialSectionArea(CivSection section, AreaCalculationMethod method)
        {
            // Lấy diện tích từ API
            double apiArea = section.Area;
            double shoelaceArea = 0;
            
            // Luôn tính từ SectionPoints nếu có đủ điểm
            try
            {
                if (section.SectionPoints.Count >= 3)
                {
                    shoelaceArea = CalculateSectionArea(section);
                }
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
            
            // Chọn diện tích theo phương pháp
            return method switch
            {
                // Ưu tiên Civil API, dự phòng dùng Công thức Dây giày nếu API = 0
                AreaCalculationMethod.CivilAPI => apiArea > 0 ? apiArea : shoelaceArea,
                
                // Ưu tiên Công thức Dây giày (từ SectionPoints), dự phòng API nếu = 0
                AreaCalculationMethod.SectionPoints => shoelaceArea > 0 ? shoelaceArea : apiArea,
                
                // Ưu tiên Công thức Dây giày nếu có điểm, dự phòng sang API
                AreaCalculationMethod.SectionPointsFirst => shoelaceArea > 0 ? shoelaceArea : apiArea,
                
                // Mặc định: ưu tiên API
                _ => apiArea > 0 ? apiArea : shoelaceArea
            };
        }

        /// <summary>
        /// Vẽ đường bao Material Section trong SectionView từ SectionPoints
        /// </summary>
        /// <param name="tr">Transaction hiện tại</param>
        /// <param name="sampleLineGroupId">ID của SampleLineGroup</param>
        /// <param name="stakeInfos">Danh sách thông tin cọc với điểm SectionPoints</param>
        /// <param name="materials">Danh sách tên vật liệu</param>
        /// <returns>Số lượng đường bao đã vẽ</returns>
        private static int DrawMaterialBoundaries(Transaction tr, ObjectId sampleLineGroupId, 
            List<StakeInfo> stakeInfos, List<string> materials)
        {
            int totalDrawn = 0;
            
            try
            {
                // Lấy SampleLineGroup
                SampleLineGroup? slg = tr.GetObject(sampleLineGroupId, AcadDb.OpenMode.ForRead) as SampleLineGroup;
                if (slg == null) return 0;
                
                // Lấy ModelSpace để thêm polyline
                AcadDb.BlockTableRecord modelSpace = tr.GetObject(
                    A.Db.CurrentSpaceId, AcadDb.OpenMode.ForWrite) as AcadDb.BlockTableRecord ?? 
                    throw new System.Exception("Không thể mở ModelSpace");
                
                // Lấy SectionViewGroup nếu có
                SectionViewGroupCollection svgCollection = slg.SectionViewGroups;
                SectionViewGroup? svGroup = null;
                if (svgCollection.Count > 0)
                {
                    svGroup = svgCollection[0];
                }
                
                
                QTOMaterialListCollection materialLists = slg.MaterialLists;
                List<(string Name, Guid ListGuid, Guid MaterialGuid)> materialInfo = new();
                
                // Màu sắc cho từng vật liệu (xoay vòng)
                short[] materialColors = { 1, 2, 3, 4, 5, 6, 30, 40, 50, 80, 140, 180, 210, 250 };
                int colorIndex = 0;
                Dictionary<string, string> materialLayerNames = new();
                
                foreach (QTOMaterialList materialList in materialLists)
                {
                    try
                    {
                        Guid listGuid = materialList.Guid;
                        foreach (QTOMaterial material in materialList)
                        {
                            if (materials.Contains(material.Name))
                            {
                                materialInfo.Add((material.Name, listGuid, material.Guid));
                                
                                // Tạo layer riêng cho mỗi loại vật liệu nếu chưa có
                                if (!materialLayerNames.ContainsKey(material.Name))
                                {
                                    // Làm sạch tên layer (loại bỏ ký tự không hợp lệ)
                                    string cleanName = material.Name.Replace("<", "").Replace(">", "")
                                        .Replace("/", "-").Replace("\\", "-").Replace(":", "-")
                                        .Replace("?", "").Replace("*", "").Replace("|", "-").Replace("\"", "");
                                    string layerName = $"C3D-BOUNDARY-{cleanName}";
                                    
                                    // Tạo layer với màu riêng
                                    short color = materialColors[colorIndex % materialColors.Length];
                                    CreateLayerIfNotExists(tr, layerName, color);
                                    
                                    materialLayerNames[material.Name] = layerName;
                                    colorIndex++;
                                }
                            }
                        }
                    }
                    catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                }
                
                // Duyệt qua từng SampleLine
                foreach (ObjectId slId in slg.GetSampleLineIds())
                {
                    SampleLine? sampleLine = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                    if (sampleLine == null) continue;
                    
                    // Tìm SectionView tương ứng để lấy tọa độ Location và ElevationMin
                    double sectionViewX = 0;
                    double sectionViewY = 0;
                    double sectionViewElevMin = 0;
                    bool foundSectionView = false;
                    
                    if (svGroup != null)
                    {
                        try
                        {
                            ObjectIdCollection sectionViewIds = svGroup.GetSectionViewIds();
                            foreach (ObjectId svId in sectionViewIds)
                            {
                                try
                                {
                                    SectionView? sv = tr.GetObject(svId, AcadDb.OpenMode.ForWrite, false, true) as SectionView;
                                    if (sv != null && sv.SampleLineId == slId)
                                    {
                                        // Lấy vị trí SectionView và ElevationMin
                                        sectionViewX = sv.Location.X;
                                        sectionViewY = sv.Location.Y;
                                        
                                        // Lấy ElevationMin một cách an toàn
                                        try
                                        {
                                            bool wasAutomatic = sv.IsElevationRangeAutomatic;
                                            if (wasAutomatic)
                                            {
                                                sv.IsElevationRangeAutomatic = false;
                                            }
                                            sectionViewElevMin = sv.ElevationMin;
                                            if (wasAutomatic)
                                            {
                                                sv.IsElevationRangeAutomatic = true;
                                            }
                                        }
                                        catch
                                        {
                                            // Nếu không thể lấy ElevationMin, dùng giá trị mặc định 0
                                            sectionViewElevMin = 0;
                                        }
                                        
                                        foundSectionView = true;
                                        break;
                                    }
                                }
                                catch { continue; }
                            }
                        }
                        catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                    }
                    
                    // Nếu không tìm thấy SectionView, bỏ qua sample line này
                    if (!foundSectionView) continue;
                    
                    // Vẽ boundary cho mỗi material
                    foreach (var (matName, listGuid, matGuid) in materialInfo)
                    {
                        try
                        {
                            ObjectId materialSectionId = sampleLine.GetMaterialSectionId(listGuid, matGuid);
                            if (materialSectionId.IsNull || !materialSectionId.IsValid) continue;
                            
                            CivSection? section = tr.GetObject(materialSectionId, AcadDb.OpenMode.ForRead) as CivSection;
                            if (section == null) continue;
                            
                            SectionPointCollection sectionPoints = section.SectionPoints;
                            if (sectionPoints.Count < 3) continue;
                            
                            // Thu thập các điểm và chuyển sang tọa độ World
                            List<Point2d> worldPoints = new();
                            foreach (SectionPoint pt in sectionPoints)
                            {
                                double worldX = sectionViewX + pt.Location.X;
                                double worldY = sectionViewY + (pt.Location.Y - sectionViewElevMin);
                                worldPoints.Add(new Point2d(worldX, worldY));
                            }
                            
                            // Loại bỏ các điểm trùng lặp
                            List<Point2d> uniquePoints = new();
                            foreach (var pt in worldPoints)
                            {
                                bool isDuplicate = false;
                                foreach (var existing in uniquePoints)
                                {
                                    if (Math.Abs(pt.X - existing.X) < 0.001 && Math.Abs(pt.Y - existing.Y) < 0.001)
                                    {
                                        isDuplicate = true;
                                        break;
                                    }
                                }
                                if (!isDuplicate) uniquePoints.Add(pt);
                            }
                            
                            if (uniquePoints.Count < 3) continue;
                            
                            // Sắp xếp các điểm theo góc từ tâm để tạo polygon khép kín không giao cắt
                            List<Point2d> sortedPoints = SortPointsByAngle(uniquePoints);
                            
                            // Lấy tên layer cho vật liệu này
                            string materialLayerName = materialLayerNames.GetValueOrDefault(matName, "C3D-BOUNDARY-MATERIAL");
                            
                            // Tạo Polyline từ các điểm đã sắp xếp
                            AcadDb.Polyline pline = new();
                            pline.SetDatabaseDefaults();
                            pline.Layer = materialLayerName;
                            
                            for (int i = 0; i < sortedPoints.Count; i++)
                            {
                                pline.AddVertexAt(i, sortedPoints[i], 0, 0, 0);
                            }
                            
                            // Đóng polyline để tạo vòng khép kín
                            pline.Closed = true;
                            
                            // Thêm polyline vào ModelSpace
                            modelSpace.AppendEntity(pline);
                            tr.AddNewlyCreatedDBObject(pline, true);
                            totalDrawn++;
                        }
                        catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                    }
                }
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n⚠️ Lỗi khi vẽ đường bao: {ex.Message}");
            }
            
            return totalDrawn;
        }
        
        /// <summary>
        /// Tạo layer nếu chưa tồn tại
        /// </summary>
        private static void CreateLayerIfNotExists(Transaction tr, string layerName, short colorIndex)
        {
            AcadDb.LayerTable? lt = tr.GetObject(A.Db.LayerTableId, AcadDb.OpenMode.ForRead) as AcadDb.LayerTable;
            if (lt == null) return;
            
            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                AcadDb.LayerTableRecord ltr = new()
                {
                    Name = layerName,
                    Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex)
                };
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }

        /// <summary>
        /// Tính khối lượng cho tất cả các cọc
        /// </summary>
        private static void CalculateVolumes(List<StakeInfo> stakeInfos, List<string> materials)
        {
            for (int i = 0; i < stakeInfos.Count; i++)
            {
                var stake = stakeInfos[i];
                foreach (var material in materials)
                {
                    double areaCurrent = stake.MaterialAreas.GetValueOrDefault(material, 0);
                    double areaPrev = (i > 0) ? stakeInfos[i - 1].MaterialAreas.GetValueOrDefault(material, 0) : 0;
                    double spacing = stake.SpacingPrev;

                    stake.MaterialVolumes[material] = CalculateVolume(areaPrev, areaCurrent, spacing);
                }
            }
        }

        /// <summary>
        /// Xuất dữ liệu ra file Excel
        /// </summary>
        private static void ExportToExcel(string filePath, List<AlignmentInfo> alignments,
            Dictionary<ObjectId, List<StakeInfo>> alignmentData, List<string> materials, Transaction tr)
        {
            using var workbook = new XLWorkbook();

            // Dictionary để lưu tổng khối lượng cho sheet tổng hợp
            Dictionary<string, Dictionary<string, double>> totalVolumes = new();

            // Tạo sheet cho từng Alignment
            foreach (var alignInfo in alignments)
            {
                var stakeInfos = alignmentData[alignInfo.AlignmentId];
                string sheetName = SanitizeSheetName(alignInfo.Name);
                var ws = workbook.Worksheets.Add(sheetName);

                // Tạo header
                CreateSheetHeader(ws, materials, alignInfo.Name);

                // Điền dữ liệu
                int row = 4; // Bắt đầu từ hàng 4 (sau 2 hàng header)
                Dictionary<string, double> alignmentTotalVolumes = new();
                
                foreach (var material in materials)
                {
                    alignmentTotalVolumes[material] = 0;
                }

                foreach (var stake in stakeInfos)
                {
                    ws.Cell(row, 1).Value = stake.StakeName;
                    ws.Cell(row, 2).Value = stake.Station;
                    ws.Cell(row, 3).Value = Math.Round(stake.SpacingPrev, 3);

                    // Ghi diện tích (Cột 4 đến 4 + materials.Count - 1)
                    int col = 4;
                    foreach (var material in materials)
                    {
                        double area = stake.MaterialAreas.GetValueOrDefault(material, 0);
                        ws.Cell(row, col).Value = Math.Round(area, 3);
                        col++;
                    }

                    // Cột Diện tích trung bình, Khối lượng (từ Civil 3D), Khối lượng cộng dồn
                    // Mỗi vật liệu có 3 cột liên tiếp
                    int currentMaterialCol = col;
                    for (int m = 0; m < materials.Count; m++)
                    {
                        string areaColLetter = GetColumnLetter(4 + m);
                        string spacingColLetter = GetColumnLetter(3); // Cột C
                        
                        // 1. Diện tích trung bình (công thức)
                        if (row == 4)
                        {
                            ws.Cell(row, currentMaterialCol).Value = 0;
                        }
                        else
                        {
                            ws.Cell(row, currentMaterialCol).FormulaA1 = $"=({areaColLetter}{row-1}+{areaColLetter}{row})/2";
                        }
                        
                        // 2. Khối lượng = Diện tích trung bình × Khoảng cách
                        string avgAreaColLetter = GetColumnLetter(currentMaterialCol);
                        ws.Cell(row, currentMaterialCol + 1).FormulaA1 = $"={avgAreaColLetter}{row}*{spacingColLetter}{row}";
                        
                        // 3. Khối lượng cộng dồn
                        string volColLetter = GetColumnLetter(currentMaterialCol + 1);
                        if (row == 4)
                        {
                            ws.Cell(row, currentMaterialCol + 2).FormulaA1 = $"={volColLetter}{row}";
                        }
                        else
                        {
                            string cumVolLetter = GetColumnLetter(currentMaterialCol + 2);
                            ws.Cell(row, currentMaterialCol + 2).FormulaA1 = $"={cumVolLetter}{row-1}+{volColLetter}{row}";
                        }

                        currentMaterialCol += 3;
                    }

                    row++;
                }

                // Thêm hàng tổng cộng
                int totalRow = row;
                ws.Cell(totalRow, 1).Value = "TỔNG CỘNG";
                ws.Range(totalRow, 1, totalRow, 3).Merge();
                ws.Cell(totalRow, 1).Style.Font.Bold = true;
                ws.Cell(totalRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Sum tất cả cột diện tích
                int colTotal = 4;
                for (int m = 0; m < materials.Count; m++)
                {
                    ws.Cell(totalRow, colTotal).FormulaA1 = $"SUM({GetColumnLetter(colTotal)}4:{GetColumnLetter(colTotal)}{row - 1})";
                    colTotal++;
                }
                
                // Sum tất cả các cột con (Average Area, Volume, Cumulative Volume)
                // Thực tế chỉ cần sum Volume là chính xác nhất
                for (int m = 0; m < materials.Count; m++)
                {
                    // Avg Area Sum (optional)
                    colTotal++; 
                    
                    // Volume Sum
                    ws.Cell(totalRow, colTotal).FormulaA1 = $"SUM({GetColumnLetter(colTotal)}4:{GetColumnLetter(colTotal)}{row - 1})";
                    colTotal++;
                    
                    // Cumulative Volume Sum (lấy giá trị cuối cùng thay vì sum)
                    ws.Cell(totalRow, colTotal).FormulaA1 = $"{GetColumnLetter(colTotal)}{row - 1}";
                    colTotal++;
                }

                // Format bảng
                FormatWorksheet(ws, row, 3 + materials.Count + materials.Count * 3);

                // Lưu tổng khối lượng cho sheet tổng hợp
                totalVolumes[sheetName] = alignmentTotalVolumes;
            }

            // Tạo sheet TỔNG HỢP nếu có nhiều hơn 1 alignment
            if (alignments.Count > 1)
            {
                CreateSummarySheet(workbook, alignments, totalVolumes, alignmentData, materials, tr);
            }

            // Lưu file
            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Tạo header cho sheet
        /// </summary>
        private static void CreateSheetHeader(IXLWorksheet ws, List<string> materials, string alignmentName)
        {
            // Tiêu đề sheet
            ws.Cell(1, 1).Value = $"BẢNG TÍNH KHỐI LƯỢNG - {alignmentName}";
            int lastCol = 3 + materials.Count * 2;
            ws.Range(1, 1, 1, lastCol).Merge();
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontSize = 14;
            ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Hàng 2: Nhóm header (Thông tin cọc | Diện tích | Khối lượng)
            ws.Cell(2, 1).Value = "THÔNG TIN CỌC";
            ws.Range(2, 1, 2, 3).Merge();
            ws.Cell(2, 1).Style.Font.Bold = true;
            ws.Cell(2, 1).Style.Fill.BackgroundColor = XLColor.LightGray;
            ws.Cell(2, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Nhóm DIỆN TÍCH
            int areaStartCol = 4;
            int areaEndCol = 3 + materials.Count;
            ws.Cell(2, areaStartCol).Value = "DIỆN TÍCH (m²)";
            ws.Range(2, areaStartCol, 2, areaEndCol).Merge();
            ws.Cell(2, areaStartCol).Style.Font.Bold = true;
            ws.Cell(2, areaStartCol).Style.Fill.BackgroundColor = XLColor.LightGreen;
            ws.Cell(2, areaStartCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Nhóm KHỐI LƯỢNG
            int volTableStartCol = areaEndCol + 1;
            int volTableEndCol = 3 + materials.Count + materials.Count * 3;
            ws.Cell(2, volTableStartCol).Value = "KHỐI LƯỢNG (m³)";
            ws.Range(2, volTableStartCol, 2, volTableEndCol).Merge();
            ws.Cell(2, volTableStartCol).Style.Font.Bold = true;
            ws.Cell(2, volTableStartCol).Style.Fill.BackgroundColor = XLColor.LightYellow;
            ws.Cell(2, volTableStartCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Hàng 3: Header chi tiết
            ws.Cell(3, 1).Value = "TÊN CỌC";
            ws.Cell(3, 2).Value = "LÝ TRÌNH";
            ws.Cell(3, 3).Value = "K.CÁCH (m)";

            // Header tên vật liệu cho nhóm diện tích
            int col = 4;
            foreach (var material in materials)
            {
                ws.Cell(3, col).Value = material;
                col++;
            }

            // Header cho nhóm khối lượng (mỗi vật liệu 3 cột)
            foreach (var material in materials)
            {
                ws.Cell(3, col).Value = "DT TB";      // Diện tích trung bình
                ws.Cell(3, col + 1).Value = material; // Khối lượng
                ws.Cell(3, col + 2).Value = "CỘNG DỒN"; // Khối lượng cộng dồn
                
                // Format riêng cho header vật liệu
                ws.Range(3, col, 3, col + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                col += 3;
            }

            // Format header row 3
            var headerRange = ws.Range(3, 1, 3, volTableEndCol);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.WrapText = true;
        }

        /// <summary>
        /// Format toàn bộ worksheet
        /// </summary>
        private static void FormatWorksheet(IXLWorksheet ws, int lastRow, int lastCol)
        {
            // Border cho toàn bộ bảng
            var tableRange = ws.Range(2, 1, lastRow, lastCol);
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Điều chỉnh độ rộng cột
            ws.Column(1).Width = 15;  // Tên cọc
            ws.Column(2).Width = 15;  // Lý trình
            ws.Column(3).Width = 15;  // Khoảng cách

            for (int c = 4; c <= lastCol; c++)
            {
                ws.Column(c).Width = 18;
            }

            // Căn giữa các cột số
            for (int r = 4; r <= lastRow; r++)
            {
                for (int c = 3; c <= lastCol; c++)
                {
                    ws.Cell(r, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                }
            }
        }

        /// <summary>
        /// Tạo sheet tổng hợp
        /// </summary>
        private static void CreateSummarySheet(XLWorkbook workbook, List<AlignmentInfo> alignments,
            Dictionary<string, Dictionary<string, double>> totalVolumes, 
            Dictionary<ObjectId, List<StakeInfo>> alignmentData,
            List<string> materials, Transaction tr)
        {
            var ws = workbook.Worksheets.Add("TỔNG HỢP");

            // Tiêu đề
            ws.Cell(1, 1).Value = "BẢNG TỔNG HỢP KHỐI LƯỢNG TẤT CẢ CÁC TUYẾN";
            int lastCol = 1 + materials.Count;
            ws.Range(1, 1, 1, lastCol).Merge();
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontSize = 14;
            ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Header
            ws.Cell(2, 1).Value = "TUYẾN";
            int col = 2;
            foreach (var material in materials)
            {
                ws.Cell(2, col).Value = $"{material} (m³)";
                col++;
            }

            // Header format
            var headerRange = ws.Range(2, 1, 2, lastCol);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGreen;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Dữ liệu
            int row = 3;
            foreach (var alignInfo in alignments)
            {
                string sheetName = SanitizeSheetName(alignInfo.Name);
                ws.Cell(row, 1).Value = alignInfo.Name;

                col = 2;
                foreach (var material in materials)
                {
                    int mIndex = materials.IndexOf(material);
                    // Cột Volume trong sheet chi tiết:
                    // 1 (Tên) + 1 (Lý trình) + 1 (K.cách) + materials.Count (Các cột Area) + (mIndex * 3) + 2 (Cột Volume trong nhóm 3 cột)
                    int targetCol = 3 + materials.Count + (mIndex * 3) + 2;
                    int lastRow = alignmentData[alignInfo.AlignmentId].Count + 4;
                    
                    ws.Cell(row, col).FormulaA1 = $"='{sheetName}'!{GetColumnLetter(targetCol)}{lastRow}";
                    col++;
                }
                row++;
            }

            // Hàng tổng cộng
            ws.Cell(row, 1).Value = "TỔNG CỘNG";
            ws.Cell(row, 1).Style.Font.Bold = true;
            col = 2;
            foreach (var material in materials)
            {
                ws.Cell(row, col).FormulaA1 = $"SUM({GetColumnLetter(col)}3:{GetColumnLetter(col)}{row - 1})";
                col++;
            }

            // Format
            var tableRange = ws.Range(2, 1, row, lastCol);
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Column(1).Width = 30;
            for (int c = 2; c <= lastCol; c++)
            {
                ws.Column(c).Width = 20;
            }
        }

        /// <summary>
        /// Lấy ký tự cột Excel từ số cột
        /// </summary>
        private static string GetColumnLetter(int columnNumber)
        {
            string columnLetter = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnLetter = Convert.ToChar('A' + modulo) + columnLetter;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnLetter;
        }

        /// <summary>
        /// Làm sạch tên sheet (Excel không cho phép một số ký tự đặc biệt)
        /// </summary>
        private static string SanitizeSheetName(string name)
        {
            char[] invalidChars = { ':', '\\', '/', '?', '*', '[', ']' };
            string result = name;
            foreach (char c in invalidChars)
            {
                result = result.Replace(c, '_');
            }
            // Giới hạn độ dài tên sheet là 31 ký tự
            if (result.Length > 31)
            {
                result = result.Substring(0, 31);
            }
            return result;
        }

        /// <summary>
        /// Tạo bảng khối lượng trong AutoCAD
        /// </summary>
        private static void CreateCadTable(Transaction tr, Point3d insertPoint, string alignmentName, 
            List<StakeInfo> stakeInfos, List<string> materials)
        {
            AcadDb.Database db = HostApplicationServices.WorkingDatabase;
            AcadDb.BlockTable bt = tr.GetObject(db.BlockTableId, AcadDb.OpenMode.ForRead) as AcadDb.BlockTable 
                ?? throw new System.Exception("Không thể mở BlockTable");
            AcadDb.BlockTableRecord btr = tr.GetObject(bt[AcadDb.BlockTableRecord.ModelSpace], AcadDb.OpenMode.ForWrite) as AcadDb.BlockTableRecord
                ?? throw new System.Exception("Không thể mở ModelSpace");

            // Tính số cột: 3 (Tên cọc, Lý trình, K.Cách) + materials*2 (Diện tích + Khối lượng)
            int numCols = 3 + materials.Count * 2;
            int numRows = stakeInfos.Count + 4; // 2 header + dữ liệu + 1 tổng cộng

            // Tạo Table
            AcadDb.Table table = new()
            {
                Position = insertPoint,
                TableStyle = db.Tablestyle
            };

            table.SetSize(numRows, numCols);

            // Kích thước ô
            double rowHeight = 8.0;    // Chiều cao hàng
            double colWidth = 25.0;    // Chiều rộng cột mặc định
            double dataColWidth = 18.0; // Chiều rộng cột dữ liệu

            for (int r = 0; r < numRows; r++)
            {
                table.Rows[r].Height = rowHeight;
            }

            table.Columns[0].Width = colWidth;  // Tên cọc
            table.Columns[1].Width = colWidth;  // Lý trình
            table.Columns[2].Width = 15.0;      // K.Cách

            for (int c = 3; c < numCols; c++)
            {
                table.Columns[c].Width = dataColWidth;
            }

            // ===== HÀNG 0: Tiêu đề bảng =====
            table.MergeCells(CellRange.Create(table, 0, 0, 0, numCols - 1));
            table.Cells[0, 0].TextString = $"BẢNG TÍNH KHỐI LƯỢNG - {alignmentName}";
            table.Cells[0, 0].TextHeight = 5.0;
            table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

            // ===== HÀNG 1: Header nhóm =====
            // Nhóm THÔNG TIN CỌC
            table.MergeCells(CellRange.Create(table, 1, 0, 1, 2));
            table.Cells[1, 0].TextString = "THÔNG TIN CỌC";
            table.Cells[1, 0].Alignment = CellAlignment.MiddleCenter;

            // Nhóm DIỆN TÍCH
            int areaEndCol = 2 + materials.Count;
            table.MergeCells(CellRange.Create(table, 1, 3, 1, areaEndCol));
            table.Cells[1, 3].TextString = "DIỆN TÍCH (m²)";
            table.Cells[1, 3].Alignment = CellAlignment.MiddleCenter;

            // Nhóm KHỐI LƯỢNG
            if (areaEndCol + 1 < numCols)
            {
                table.MergeCells(CellRange.Create(table, 1, areaEndCol + 1, 1, numCols - 1));
                table.Cells[1, areaEndCol + 1].TextString = "KHỐI LƯỢNG (m³)";
                table.Cells[1, areaEndCol + 1].Alignment = CellAlignment.MiddleCenter;
            }

            // ===== HÀNG 2: Header chi tiết =====
            table.Cells[2, 0].TextString = "TÊN CỌC";
            table.Cells[2, 1].TextString = "LÝ TRÌNH";
            table.Cells[2, 2].TextString = "K.CÁCH";

            int col = 3;
            foreach (var material in materials)
            {
                table.Cells[2, col].TextString = material;
                table.Cells[2, col].Alignment = CellAlignment.MiddleCenter;
                col++;
            }
            foreach (var material in materials)
            {
                table.Cells[2, col].TextString = material;
                table.Cells[2, col].Alignment = CellAlignment.MiddleCenter;
                col++;
            }

            // Format header
            for (int c = 0; c < numCols; c++)
            {
                table.Cells[2, c].TextHeight = 3.5;
                table.Cells[2, c].Alignment = CellAlignment.MiddleCenter;
            }

            // ===== HÀNG 3+: Dữ liệu =====
            int row = 3;
            Dictionary<string, double> totalVolumes = new();
            foreach (var material in materials)
            {
                totalVolumes[material] = 0;
            }

            foreach (var stake in stakeInfos)
            {
                table.Cells[row, 0].TextString = stake.StakeName;
                table.Cells[row, 1].TextString = stake.Station;
                table.Cells[row, 2].TextString = Math.Round(stake.SpacingPrev, 2).ToString();

                // Diện tích
                col = 3;
                foreach (var material in materials)
                {
                    double area = stake.MaterialAreas.GetValueOrDefault(material, 0);
                    table.Cells[row, col].TextString = Math.Round(area, 3).ToString();
                    table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                    col++;
                }

                // Khối lượng
                foreach (var material in materials)
                {
                    double volume = stake.MaterialVolumes.GetValueOrDefault(material, 0);
                    table.Cells[row, col].TextString = Math.Round(volume, 3).ToString();
                    table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                    totalVolumes[material] += volume;
                    col++;
                }

                row++;
            }

            // ===== HÀNG CUỐI: Tổng cộng =====
            table.MergeCells(CellRange.Create(table, row, 0, row, 2));
            table.Cells[row, 0].TextString = "TỔNG CỘNG";
            table.Cells[row, 0].Alignment = CellAlignment.MiddleCenter;

            // Sum diện tích (để trống hoặc tính tổng)
            col = 3;
            foreach (var material in materials)
            {
                double totalArea = stakeInfos.Sum(s => s.MaterialAreas.GetValueOrDefault(material, 0));
                table.Cells[row, col].TextString = Math.Round(totalArea, 3).ToString();
                table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                col++;
            }

            // Sum khối lượng
            foreach (var material in materials)
            {
                table.Cells[row, col].TextString = Math.Round(totalVolumes[material], 3).ToString();
                table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                col++;
            }

            // Set TextHeight cho tất cả ô
            for (int r = 0; r < numRows; r++)
            {
                for (int c = 0; c < numCols; c++)
                {
                    if (r >= 3) // Dữ liệu
                    {
                        table.Cells[r, c].TextHeight = 3.0;
                    }
                }
            }

            // Thêm table vào model space
            btr.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
        }
    }
}
