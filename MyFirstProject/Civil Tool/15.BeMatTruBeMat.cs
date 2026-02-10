// (C) Copyright 2024
// Tính khối lượng trắc ngang theo cú pháp Bề mặt trừ Bề mặt
// Sử dụng Material List trong Section View
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

using ClosedXML.Excel;
using MyFirstProject.Extensions;

[assembly: CommandClass(typeof(MyFirstProject.BeMatTruBeMat))]

namespace MyFirstProject
{
    #region Data Classes

    /// <summary>
    /// Thông tin Material có cú pháp Surface - Surface
    /// </summary>
    public class SurfaceMaterialInfo
    {
        public string Name { get; set; } = "";
        public Guid MaterialListGuid { get; set; }
        public Guid MaterialGuid { get; set; }
        public QTOMaterial? Material { get; set; }
        
        // Thông tin bề mặt (nếu có)
        public string TopSurfaceName { get; set; } = "";      // Bề mặt trên (Datum Surface)
        public string BottomSurfaceName { get; set; } = "";   // Bề mặt dưới (Comparison Surface)
        public MaterialQuantityType QuantityType { get; set; }
        
        // Xác định loại khối lượng
        public bool IsCut { get; set; }   // Đào (Cut)
        public bool IsFill { get; set; }  // Đắp (Fill)
    }

    /// <summary>
    /// Thông tin chi tiết của Material Section (như trong Properties Panel)
    /// </summary>
    public class MaterialSectionData
    {
        // Thông tin cơ bản
        public string MaterialName { get; set; } = "";
        public string SectionSurfaceName { get; set; } = "";  // Material List - (7) - Material - (3)
        public string SectionType { get; set; } = "";
        public string StaticDynamic { get; set; } = "Dynamic";
        
        // Vị trí
        public double SectionStation { get; set; }          // 0+000.00
        public string SectionStationFormatted { get; set; } = "";
        
        // Phạm vi offset
        public double LeftLength { get; set; }              // -4.950m (giá trị âm = bên trái)
        public double RightLength { get; set; }             // 4.950m (giá trị dương = bên phải)
        
        // Phạm vi cao độ
        public double SectionMinElevation { get; set; }     // 9.369m
        public double SectionMaxElevation { get; set; }     // 10.270m
        
        // Diện tích và Criteria
        public double Area { get; set; }                    // 0.88sq.m
        public string Criteria { get; set; } = "";
        
        // Điểm chi tiết
        public List<Point3d> Points { get; set; } = new();
        public int PointCount => Points.Count;
        
        // Thuộc tính phụ trợ
        public bool IsCut { get; set; }
        public bool IsFill { get; set; }
        
        // Tính chiều rộng tổng
        public double TotalWidth => Math.Abs(LeftLength) + Math.Abs(RightLength);
        
        // Tính chiều cao
        public double Height => SectionMaxElevation - SectionMinElevation;
    }

    /// <summary>
    /// Thông tin diện tích và khối lượng tại một trắc ngang
    /// </summary>
    public class CrossSectionVolumeInfo
    {
        public string SampleLineName { get; set; } = "";
        public double Station { get; set; }
        public string StationFormatted { get; set; } = "";
        public double SpacingPrev { get; set; }
        
        // Diện tích đào/đắp tại trắc ngang này
        public Dictionary<string, double> CutAreas { get; set; } = new();    // Đào
        public Dictionary<string, double> FillAreas { get; set; } = new();   // Đắp
        
        // Khối lượng đào/đắp từ trắc ngang trước đến trắc ngang này
        public Dictionary<string, double> CutVolumes { get; set; } = new();
        public Dictionary<string, double> FillVolumes { get; set; } = new();
        
        // Chi tiết Material Section Data (MỚI - chứa tất cả thông tin từ Properties Panel)
        public Dictionary<string, MaterialSectionData> MaterialSections { get; set; } = new();
        
        // Chi tiết các điểm section (giữ lại để tương thích)
        public Dictionary<string, List<Point3d>> SectionPoints { get; set; } = new();
    }

    /// <summary>
    /// Kết quả tổng hợp khối lượng
    /// </summary>
    public class VolumeResult
    {
        public string AlignmentName { get; set; } = "";
        public List<CrossSectionVolumeInfo> CrossSections { get; set; } = new();
        public Dictionary<string, double> TotalCutVolumes { get; set; } = new();
        public Dictionary<string, double> TotalFillVolumes { get; set; } = new();
    }

    #endregion

    /// <summary>
    /// Lớp tính khối lượng theo cú pháp Bề mặt - Bề mặt
    /// </summary>
    public class BeMatTruBeMat
    {
        #region Commands

        /// <summary>
        /// Command: Tính khối lượng đào đắp theo Material List (Surface - Surface)
        /// </summary>
        [CommandMethod("CTSV_DaoDap")]
        public static void CTSVDaoDap()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                A.Ed.WriteMessage("\n=== TÍNH KHỐI LƯỢNG ĐÀO ĐẮP (BỀ MẶT - BỀ MẶT) ===\n");

                // 1. Lấy danh sách Alignments có SampleLineGroup
                List<AlignmentInfo> alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Không tìm thấy Alignment nào có SampleLineGroup!");
                    return;
                }

                // 2. Hiển thị form chọn Alignment
                using FormChonAlignment formChon = new(alignments);
                if (Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(formChon) != DialogResult.OK)
                    return;

                if (formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Chưa chọn Alignment nào!");
                    return;
                }

                // 3. Thu thập dữ liệu từ tất cả Alignments
                List<VolumeResult> allResults = new();
                HashSet<string> allMaterialNames = new();

                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    A.Ed.WriteMessage($"\n📊 Đang xử lý: {alignInfo.Name}...");
                    
                    var result = ExtractSurfaceMaterialVolumes(tr, alignInfo);
                    if (result != null)
                    {
                        allResults.Add(result);
                        
                        // Thu thập tên materials
                        foreach (var cs in result.CrossSections)
                        {
                            foreach (var key in cs.CutAreas.Keys) allMaterialNames.Add(key);
                            foreach (var key in cs.FillAreas.Keys) allMaterialNames.Add(key);
                        }
                    }
                }

                if (allResults.Count == 0 || allMaterialNames.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Không tìm thấy dữ liệu Material nào!");
                    A.Ed.WriteMessage("\nVui lòng kiểm tra Material List trong QTO Manager.");
                    return;
                }

                A.Ed.WriteMessage($"\n✅ Tìm thấy {allMaterialNames.Count} loại material:");
                foreach (var mat in allMaterialNames.OrderBy(m => m))
                    A.Ed.WriteMessage($"\n  - {mat}");

                // 4. Chọn loại xuất
                PromptKeywordOptions pkOpts = new("\nChọn loại xuất [Excel/CAD/TatCa]", "Excel CAD TatCa");
                pkOpts.Keywords.Default = "Excel";
                pkOpts.AllowNone = true;
                PromptResult pkResult = A.Ed.GetKeywords(pkOpts);

                if (pkResult.Status != PromptStatus.OK && pkResult.Status != PromptStatus.None)
                    return;

                string exportType = pkResult.StringResult ?? "Excel";
                bool doExcel = exportType == "Excel" || exportType == "TatCa";
                bool doCad = exportType == "CAD" || exportType == "TatCa";

                List<string> orderedMaterials = allMaterialNames.OrderBy(m => m).ToList();

                // 5. Xuất Excel
                string excelPath = "";
                if (doExcel)
                {
                    SaveFileDialog saveDialog = new()
                    {
                        Title = "Lưu file Excel khối lượng đào đắp",
                        Filter = "Excel Files|*.xlsx",
                        DefaultExt = "xlsx",
                        FileName = $"DaoDap_{DateTime.Now:yyyyMMdd_HHmmss}"
                    };

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        excelPath = saveDialog.FileName;
                        ExportDaoDapToExcel(excelPath, allResults, orderedMaterials);
                        A.Ed.WriteMessage($"\n✅ Đã xuất file Excel: {excelPath}");
                    }
                }

                // 6. Vẽ bảng CAD
                if (doCad)
                {
                    // Hỏi loại bảng cần xuất
                    PromptKeywordOptions tableOpts = new("\nChọn loại bảng CAD [TongHop/ChiTiet/TatCa]", "TongHop ChiTiet TatCa");
                    tableOpts.Keywords.Default = "TatCa";
                    tableOpts.AllowNone = true;
                    PromptResult tableResult = A.Ed.GetKeywords(tableOpts);
                    
                    string tableType = tableResult.StringResult ?? "TatCa";
                    bool doSummaryTable = tableType == "TongHop" || tableType == "TatCa";
                    bool doDetailTable = tableType == "ChiTiet" || tableType == "TatCa";

                    PromptPointResult ppr = A.Ed.GetPoint("\nChọn điểm chèn bảng: ");
                    if (ppr.Status == PromptStatus.OK)
                    {
                        Point3d insertPoint = ppr.Value;
                        
                        foreach (var result in allResults)
                        {
                            if (doSummaryTable)
                            {
                                CreateDaoDapCadTable(tr, insertPoint, result, orderedMaterials);
                                A.Ed.WriteMessage($"\n✅ Đã vẽ bảng TỔNG HỢP cho '{result.AlignmentName}'");
                                
                                // Offset cho bảng tiếp theo
                                double tableHeight = (result.CrossSections.Count + 5) * 8.0;
                                insertPoint = new Point3d(insertPoint.X, insertPoint.Y - tableHeight - 50, insertPoint.Z);
                            }
                            
                            if (doDetailTable)
                            {
                                CreateDetailCadTable(tr, insertPoint, result);
                                A.Ed.WriteMessage($"\n✅ Đã vẽ bảng CHI TIẾT cho '{result.AlignmentName}'");
                                
                                // Tính số dòng chi tiết
                                int detailRows = result.CrossSections.Sum(cs => cs.MaterialSections.Count) + 2;
                                double detailHeight = detailRows * 7.0;
                                insertPoint = new Point3d(insertPoint.X, insertPoint.Y - detailHeight - 50, insertPoint.Z);
                            }
                        }
                    }
                }

                // 7. Hỏi mở file Excel
                if (!string.IsNullOrEmpty(excelPath))
                {
                    if (MessageBox.Show("Đã xuất file Excel thành công!\nBạn có muốn mở file?",
                        "Hoàn thành", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
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
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
                A.Ed.WriteMessage($"\nStack: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// Command: Hiển thị thông tin Material List trong SampleLineGroup
        /// </summary>
        [CommandMethod("CTSV_HienThiMaterialList")]
        public static void CTSVHienThiMaterialList()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                A.Ed.WriteMessage("\n=== THÔNG TIN MATERIAL LIST ===\n");

                // Lấy danh sách Alignments có SampleLineGroup
                ObjectIdCollection alignmentIds = A.Cdoc.GetAlignmentIds();
                
                foreach (ObjectId alignmentId in alignmentIds)
                {
                    Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
                    if (alignment == null) continue;

                    ObjectIdCollection slgIds = alignment.GetSampleLineGroupIds();
                    if (slgIds.Count == 0) continue;

                    A.Ed.WriteMessage($"\n📍 Alignment: {alignment.Name}");

                    foreach (ObjectId slgId in slgIds)
                    {
                        SampleLineGroup? slg = tr.GetObject(slgId, AcadDb.OpenMode.ForWrite, false, true) as SampleLineGroup;
                        if (slg == null) continue;

                        A.Ed.WriteMessage($"\n  📊 SampleLineGroup: {slg.Name}");

                        // Hiển thị Material Lists
                        QTOMaterialListCollection materialLists = slg.MaterialLists;
                        A.Ed.WriteMessage($"\n    Số lượng Material Lists: {materialLists.Count}");

                        foreach (QTOMaterialList materialList in materialLists)
                        {
                            A.Ed.WriteMessage($"\n\n    📋 MaterialList: {materialList.Name ?? "(không tên)"}");
                            A.Ed.WriteMessage($"\n       GUID: {materialList.Guid}");

                            foreach (QTOMaterial material in materialList)
                            {
                                A.Ed.WriteMessage($"\n\n       🧱 Material: {material.Name}");
                                A.Ed.WriteMessage($"\n          GUID: {material.Guid}");
                                A.Ed.WriteMessage($"\n          QuantityType: {material.QuantityType}");
                                
                                // Lấy thông tin bề mặt nếu có
                                try
                                {
                                    // Kiểm tra các thuộc tính của material
                                    A.Ed.WriteMessage($"\n          Type: {material.GetType().Name}");
                                }
                                catch { }
                            }
                        }

                        // Hiển thị Section Sources
                        SectionSourceCollection sources = slg.GetSectionSources();
                        A.Ed.WriteMessage($"\n\n    📌 Section Sources: {sources.Count}");
                        
                        foreach (SectionSource source in sources)
                        {
                            A.Ed.WriteMessage($"\n       - {source.SourceName} ({source.SourceType})");
                        }
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
        /// Command: Hiển thị chi tiết dữ liệu Material Section (giống Properties Panel)
        /// </summary>
        [CommandMethod("CTSV_ChiTietMaterialSection")]
        public static void CTSVChiTietMaterialSection()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                A.Ed.WriteMessage("\n=== CHI TIẾT DỮ LIỆU MATERIAL SECTION ===\n");

                // Lấy danh sách Alignments có SampleLineGroup
                List<AlignmentInfo> alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Không tìm thấy Alignment nào có SampleLineGroup!");
                    return;
                }

                // Hiển thị form chọn Alignment
                using FormChonAlignment formChon = new(alignments);
                if (Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(formChon) != DialogResult.OK)
                    return;

                if (formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Chưa chọn Alignment nào!");
                    return;
                }

                // Xử lý từng Alignment
                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    A.Ed.WriteMessage($"\n\n╔══════════════════════════════════════════════════════════╗");
                    A.Ed.WriteMessage($"\n║ ALIGNMENT: {alignInfo.Name,-46} ║");
                    A.Ed.WriteMessage($"\n╚══════════════════════════════════════════════════════════╝");

                    var result = ExtractSurfaceMaterialVolumes(tr, alignInfo);
                    if (result == null) continue;

                    foreach (var cs in result.CrossSections)
                    {
                        A.Ed.WriteMessage($"\n\n┌───────────────────────────────────────────────────────────┐");
                        A.Ed.WriteMessage($"\n│ TRẮC NGANG: {cs.SampleLineName,-44} │");
                        A.Ed.WriteMessage($"\n│ LÝ TRÌNH: {cs.StationFormatted,-47} │");
                        A.Ed.WriteMessage($"\n└───────────────────────────────────────────────────────────┘");

                        foreach (var kvp in cs.MaterialSections)
                        {
                            var data = kvp.Value;
                            A.Ed.WriteMessage($"\n  ┌─────────────────────────────────────────────────────────┐");
                            A.Ed.WriteMessage($"\n  │ MATERIAL: {data.MaterialName,-46} │");
                            A.Ed.WriteMessage($"\n  ├─────────────────────────────────────────────────────────┤");
                            A.Ed.WriteMessage($"\n  │ Section Surface Name: {data.SectionSurfaceName,-32} │");
                            A.Ed.WriteMessage($"\n  │ Section Type:         {data.SectionType,-32} │");
                            A.Ed.WriteMessage($"\n  │ Static/Dynamic:       {data.StaticDynamic,-32} │");
                            A.Ed.WriteMessage($"\n  │ Section Station:      {data.SectionStationFormatted,-32} │");
                            A.Ed.WriteMessage($"\n  ├─────────────────────────────────────────────────────────┤");
                            A.Ed.WriteMessage($"\n  │ Left Length:          {data.LeftLength,10:F3} m                    │");
                            A.Ed.WriteMessage($"\n  │ Right Length:         {data.RightLength,10:F3} m                    │");
                            A.Ed.WriteMessage($"\n  │ Total Width:          {data.TotalWidth,10:F3} m                    │");
                            A.Ed.WriteMessage($"\n  ├─────────────────────────────────────────────────────────┤");
                            A.Ed.WriteMessage($"\n  │ Section Min Elevation:{data.SectionMinElevation,10:F3} m                    │");
                            A.Ed.WriteMessage($"\n  │ Section Max Elevation:{data.SectionMaxElevation,10:F3} m                    │");
                            A.Ed.WriteMessage($"\n  │ Height:               {data.Height,10:F3} m                    │");
                            A.Ed.WriteMessage($"\n  ├─────────────────────────────────────────────────────────┤");
                            A.Ed.WriteMessage($"\n  │ Area:                 {data.Area,10:F4} m²                   │");
                            A.Ed.WriteMessage($"\n  │ Point Count:          {data.PointCount,10}                       │");
                            A.Ed.WriteMessage($"\n  │ Type:                 {(data.IsCut ? "ĐÀO (CUT)" : "ĐẮP (FILL)"),-32} │");
                            A.Ed.WriteMessage($"\n  └─────────────────────────────────────────────────────────┘");
                        }
                    }

                    // Hiển thị tổng cộng
                    A.Ed.WriteMessage($"\n\n╔══════════════════════════════════════════════════════════╗");
                    A.Ed.WriteMessage($"\n║                    TỔNG KHỐI LƯỢNG                       ║");
                    A.Ed.WriteMessage($"\n╠══════════════════════════════════════════════════════════╣");
                    foreach (var kvp in result.TotalCutVolumes)
                    {
                        A.Ed.WriteMessage($"\n║ {kvp.Key,-20} ĐÀO: {kvp.Value,12:F4} m³              ║");
                    }
                    foreach (var kvp in result.TotalFillVolumes)
                    {
                        A.Ed.WriteMessage($"\n║ {kvp.Key,-20} ĐẮP: {kvp.Value,12:F4} m³              ║");
                    }
                    A.Ed.WriteMessage($"\n╚══════════════════════════════════════════════════════════╝");
                }

                tr.Commit();
                A.Ed.WriteMessage("\n\n=== HOÀN THÀNH ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Command: Xuất khối lượng trắc ngang (lấy AREA từ Material Section, tính khối lượng theo Average End Area)
        /// </summary>
        [CommandMethod("CTSV_KhoiLuongTracNgang")]
        public static void CTSVKhoiLuongTracNgang()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                A.Ed.WriteMessage("\n=== XUẤT KHỐI LƯỢNG TRẮC NGANG ===\n");
                A.Ed.WriteMessage("\n📊 Lấy AREA từ Material Section -> Tính khối lượng theo công thức:");
                A.Ed.WriteMessage("\n   V = (S1 + S2) / 2 × L");
                A.Ed.WriteMessage("\n   Trong đó: S1, S2 = Diện tích 2 trắc ngang liên tiếp");
                A.Ed.WriteMessage("\n             L = Khoảng cách giữa 2 trắc ngang\n");

                // 1. Lấy danh sách Alignments
                List<AlignmentInfo> alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Không tìm thấy Alignment nào có SampleLineGroup!");
                    return;
                }

                // 2. Hiển thị form chọn Alignment
                using FormChonAlignment formChon = new(alignments);
                if (Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(formChon) != DialogResult.OK)
                    return;

                if (formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Chưa chọn Alignment nào!");
                    return;
                }

                // 3. Thu thập dữ liệu
                List<VolumeResult> allResults = new();
                HashSet<string> allMaterialNames = new();

                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    A.Ed.WriteMessage($"\n📍 Đang trích xuất dữ liệu: {alignInfo.Name}...");
                    
                    var result = ExtractSurfaceMaterialVolumes(tr, alignInfo);
                    if (result != null)
                    {
                        allResults.Add(result);
                        
                        // Thu thập tên materials từ MaterialSections
                        foreach (var cs in result.CrossSections)
                        {
                            foreach (var key in cs.MaterialSections.Keys) 
                                allMaterialNames.Add(key);
                        }
                    }
                }

                if (allResults.Count == 0 || allMaterialNames.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Không tìm thấy dữ liệu Material Section nào!");
                    return;
                }

                // Sắp xếp materials
                List<string> orderedMaterials = allMaterialNames.OrderBy(m => m).ToList();
                A.Ed.WriteMessage($"\n\n✅ Tìm thấy {orderedMaterials.Count} loại material:");
                foreach (var mat in orderedMaterials)
                    A.Ed.WriteMessage($"\n  - {mat}");

                // 4. Hỏi xuất loại gì
                PromptKeywordOptions pkOpts = new("\nChọn loại xuất [Excel/CAD/TatCa]", "Excel CAD TatCa");
                pkOpts.Keywords.Default = "Excel";
                PromptResult pkResult = A.Ed.GetKeywords(pkOpts);

                bool doExcel = pkResult.StringResult == "Excel" || pkResult.StringResult == "TatCa";
                bool doCad = pkResult.StringResult == "CAD" || pkResult.StringResult == "TatCa";

                // 5. Xuất Excel
                if (doExcel)
                {
                    SaveFileDialog saveDialog = new()
                    {
                        Title = "Lưu file Excel khối lượng trắc ngang",
                        Filter = "Excel Files|*.xlsx",
                        FileName = $"KhoiLuong_TracNgang_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                    };

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        ExportTracNgangToExcel(saveDialog.FileName, allResults, orderedMaterials);
                        A.Ed.WriteMessage($"\n✅ Đã xuất file Excel: {saveDialog.FileName}");
                        
                        // Mở file Excel
                        if (MessageBox.Show("Đã xuất file Excel!\nBạn có muốn mở file?",
                            "Hoàn thành", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = saveDialog.FileName,
                                UseShellExecute = true
                            });
                        }
                    }
                }

                // 6. Xuất CAD
                if (doCad)
                {
                    PromptPointResult ppr = A.Ed.GetPoint("\nChọn điểm chèn bảng: ");
                    if (ppr.Status == PromptStatus.OK)
                    {
                        Point3d insertPoint = ppr.Value;
                        
                        foreach (var result in allResults)
                        {
                            CreateTracNgangCadTable(tr, insertPoint, result, orderedMaterials);
                            A.Ed.WriteMessage($"\n✅ Đã vẽ bảng cho '{result.AlignmentName}'");
                            
                            // Offset cho bảng tiếp theo
                            double tableHeight = (result.CrossSections.Count + 5) * 7.0;
                            insertPoint = new Point3d(insertPoint.X, insertPoint.Y - tableHeight - 50, insertPoint.Z);
                        }
                    }
                }

                tr.Commit();
                A.Ed.WriteMessage("\n\n=== HOÀN THÀNH ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Command: So sánh khối lượng - Kiểm tra AREA từ Material Section với Civil 3D
        /// </summary>
        [CommandMethod("CTSV_SoSanhKhoiLuong")]
        public static void CTSVSoSanhKhoiLuong()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                A.Ed.WriteMessage("\n=== SO SÁNH DIỆN TÍCH & KHỐI LƯỢNG ===\n");
                A.Ed.WriteMessage("\nSo sánh AREA từ section.Area (API) với AREA tính từ SectionPoints (Shoelace)\n");

                // 1. Lấy danh sách Alignments
                List<AlignmentInfo> alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Không tìm thấy Alignment nào có SampleLineGroup!");
                    return;
                }

                // 2. Hiển thị form chọn Alignment
                using FormChonAlignment formChon = new(alignments);
                if (Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(formChon) != DialogResult.OK)
                    return;

                if (formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Chưa chọn Alignment nào!");
                    return;
                }

                // 3. So sánh từng Alignment
                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    A.Ed.WriteMessage($"\n\n╔════════════════════════════════════════════════════════════════════════╗");
                    A.Ed.WriteMessage($"\n║ ALIGNMENT: {alignInfo.Name,-65} ║");
                    A.Ed.WriteMessage($"\n╚════════════════════════════════════════════════════════════════════════╝");

                    // Lấy SampleLineGroup
                    SampleLineGroup? slg = tr.GetObject(alignInfo.SampleLineGroupId, AcadDb.OpenMode.ForWrite, false, true) as SampleLineGroup;
                    if (slg == null) continue;

                    QTOMaterialListCollection materialLists = slg.MaterialLists;
                    if (materialLists.Count == 0)
                    {
                        A.Ed.WriteMessage("\n⚠️ Không có Material List!");
                        continue;
                    }

                    // Thu thập Materials
                    List<(string Name, Guid ListGuid, Guid MatGuid, QTOMaterial Mat)> materials = new();
                    foreach (QTOMaterialList matList in materialLists)
                    {
                        foreach (QTOMaterial mat in matList)
                        {
                            materials.Add((mat.Name, matList.Guid, mat.Guid, mat));
                        }
                    }

                    // Lấy SampleLines và so sánh
                    var slIds = slg.GetSampleLineIds();
                    var sampleLines = new List<SampleLine>();
                    foreach (ObjectId slId in slIds)
                    {
                        var sl = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                        if (sl != null) sampleLines.Add(sl);
                    }
                    sampleLines = sampleLines.OrderBy(s => s.Station).ToList();

                    A.Ed.WriteMessage($"\n\n📊 Tìm thấy {sampleLines.Count} SampleLine, {materials.Count} Material\n");

                    // Bảng so sánh
                    A.Ed.WriteMessage($"\n┌───────────────────────────────────────────────────────────────────────────┐");
                    A.Ed.WriteMessage($"\n│ {"SAMPLELINE",-15} │ {"MATERIAL",-25} │ {"API AREA",-12} │ {"CALC AREA",-12} │ {"DIFF%",-6} │");
                    A.Ed.WriteMessage($"\n├───────────────────────────────────────────────────────────────────────────┤");

                    int matchCount = 0;
                    int totalCount = 0;
                    double totalApiArea = 0;
                    double totalCalcArea = 0;

                    foreach (var sl in sampleLines.Take(10)) // Chỉ lấy 10 đầu tiên để demo
                    {
                        foreach (var (name, listGuid, matGuid, mat) in materials)
                        {
                            try
                            {
                                ObjectId sectionId = sl.GetMaterialSectionId(listGuid, matGuid);
                                if (sectionId.IsNull || !sectionId.IsValid) continue;

                                AcadDb.DBObject? sectionObj = tr.GetObject(sectionId, AcadDb.OpenMode.ForRead, false, true);
                                if (sectionObj is CivSection section)
                                {
                                    double apiArea = section.Area;  // AREA từ API Civil 3D
                                    
                                    // Tính AREA bằng Shoelace formula
                                    double calcArea = CalculateAreaFromPoints(section);

                                    double diff = apiArea > 0 ? Math.Abs(apiArea - calcArea) / apiArea * 100 : 0;

                                    totalApiArea += apiArea;
                                    totalCalcArea += calcArea;
                                    totalCount++;

                                    if (diff < 1) matchCount++;

                                    string status = diff < 1 ? "✅" : (diff < 5 ? "⚠️" : "❌");
                                    string shortName = name.Length > 23 ? name.Substring(0, 20) + "..." : name;

                                    A.Ed.WriteMessage($"\n│ {sl.Name,-15} │ {shortName,-25} │ {apiArea,12:F4} │ {calcArea,12:F4} │ {status,-6} │");
                                }
                            }
                            catch { }
                        }
                    }

                    A.Ed.WriteMessage($"\n├───────────────────────────────────────────────────────────────────────────┤");
                    double totalDiff = totalApiArea > 0 ? Math.Abs(totalApiArea - totalCalcArea) / totalApiArea * 100 : 0;
                    A.Ed.WriteMessage($"\n│ TỔNG                                   │ {totalApiArea,12:F4} │ {totalCalcArea,12:F4} │ {totalDiff,5:F1}% │");
                    A.Ed.WriteMessage($"\n└───────────────────────────────────────────────────────────────────────────┘");

                    // Kết luận
                    double matchPercent = totalCount > 0 ? (double)matchCount / totalCount * 100 : 0;
                    A.Ed.WriteMessage($"\n\n📈 Kết quả: {matchCount}/{totalCount} ({matchPercent:F0}%) khớp (chênh lệch < 1%)");

                    if (matchPercent >= 95)
                        A.Ed.WriteMessage($"\n✅ TUYỆT VỜI! Dữ liệu KHỚP với Civil 3D");
                    else if (matchPercent >= 80)
                        A.Ed.WriteMessage($"\n⚠️ GẦN KHỚP. Một số section có chênh lệch nhỏ");
                    else
                        A.Ed.WriteMessage($"\n❌ CẦN KIỂM TRA LẠI dữ liệu Material Section");
                }

                tr.Commit();
                A.Ed.WriteMessage("\n\n=== HOÀN THÀNH ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Command: Phân tích chi tiết nguồn gốc AREA trong Data Section (VD: 0.89 SQ.M)
        /// </summary>
        [CommandMethod("CTSV_PhanTichArea")]
        public static void CTSVPhanTichArea()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                A.Ed.WriteMessage("\n╔═══════════════════════════════════════════════════════════════════╗");
                A.Ed.WriteMessage("\n║        PHÂN TÍCH CHI TIẾT NGUỒN GỐC AREA TRONG DATA SECTION       ║");
                A.Ed.WriteMessage("\n╚═══════════════════════════════════════════════════════════════════╝");

                A.Ed.WriteMessage("\n\n📚 GIẢI THÍCH AREA (VD: 0.89 SQ.M):");
                A.Ed.WriteMessage("\n   ➤ AREA = Diện tích mặt cắt vật liệu tại 1 trắc ngang");
                A.Ed.WriteMessage("\n   ➤ Nguồn: section.Area (thuộc tính CivSection.Area trong API)");
                A.Ed.WriteMessage("\n   ➤ Tính từ: Đa giác SectionPoints (polygon closed area)");
                A.Ed.WriteMessage("\n   ➤ Công thức: Shoelace = ½|Σ(xᵢyᵢ₊₁ - xᵢ₊₁yᵢ)|");

                // Lấy Alignments
                List<AlignmentInfo> alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0) { A.Ed.WriteMessage("\n\n⚠️ Không có Alignment!"); return; }

                using FormChonAlignment formChon = new(alignments);
                if (Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(formChon) != DialogResult.OK) return;
                if (formChon.SelectedAlignments.Count == 0) return;

                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    SampleLineGroup? slg = tr.GetObject(alignInfo.SampleLineGroupId, AcadDb.OpenMode.ForWrite, false, true) as SampleLineGroup;
                    if (slg == null) continue;

                    QTOMaterialListCollection materialLists = slg.MaterialLists;
                    if (materialLists.Count == 0) { A.Ed.WriteMessage("\n⚠️ Không có Material List!"); continue; }

                    QTOMaterial? firstMat = null;
                    Guid listGuid = Guid.Empty, matGuid = Guid.Empty;
                    foreach (QTOMaterialList matList in materialLists)
                    {
                        foreach (QTOMaterial mat in matList)
                        {
                            firstMat = mat; listGuid = matList.Guid; matGuid = mat.Guid; break;
                        }
                        if (firstMat != null) break;
                    }
                    if (firstMat == null) continue;

                    var slIds = slg.GetSampleLineIds();
                    foreach (ObjectId slId in slIds)
                    {
                        var sampleLine = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                        if (sampleLine == null) continue;

                        try
                        {
                            ObjectId sectionId = sampleLine.GetMaterialSectionId(listGuid, matGuid);
                            if (sectionId.IsNull || !sectionId.IsValid) continue;

                            var sectionObj = tr.GetObject(sectionId, AcadDb.OpenMode.ForRead, false, true);
                            if (sectionObj is CivSection section)
                            {
                                A.Ed.WriteMessage($"\n\n╔════════════════════════════════════════════════════════════════════════╗");
                                A.Ed.WriteMessage($"\n║ ALIGNMENT: {alignInfo.Name,-65} ║");
                                A.Ed.WriteMessage($"\n║ SAMPLELINE: {sampleLine.Name,-64} ║");
                                A.Ed.WriteMessage($"\n║ MATERIAL: {firstMat.Name,-66} ║");
                                A.Ed.WriteMessage($"\n╠════════════════════════════════════════════════════════════════════════╣");

                                // 1. AREA từ API
                                double apiArea = section.Area;
                                A.Ed.WriteMessage($"\n║ 1. AREA TỪ API (section.Area): {apiArea,25:F6} m² ║");

                                // 2. SectionPoints
                                List<Point3d> points = new();
                                try { foreach (SectionPoint pt in section.SectionPoints) points.Add(pt.Location); } catch { }

                                A.Ed.WriteMessage($"\n║ 2. SỐ ĐIỂM (SectionPoints):    {points.Count,25} điểm ║");

                                // 3. Tọa độ đầu tiên
                                A.Ed.WriteMessage($"\n╠════════════════════════════════════════════════════════════════════════╣");
                                A.Ed.WriteMessage($"\n║ 3. TỌA ĐỘ CÁC ĐIỂM (X = Offset, Y = Elevation)                         ║");
                                
                                double minX = double.MaxValue, maxX = double.MinValue;
                                double minY = double.MaxValue, maxY = double.MinValue;

                                for (int i = 0; i < Math.Min(points.Count, 12); i++)
                                {
                                    var pt = points[i];
                                    A.Ed.WriteMessage($"\n║   P{i + 1,2}: X = {pt.X,10:F4}m,  Y = {pt.Y,10:F4}m                              ║");
                                    if (pt.X < minX) minX = pt.X; if (pt.X > maxX) maxX = pt.X;
                                    if (pt.Y < minY) minY = pt.Y; if (pt.Y > maxY) maxY = pt.Y;
                                }
                                if (points.Count > 12) A.Ed.WriteMessage($"\n║   ... và {points.Count - 12} điểm khác                                                   ║");

                                // 4. Shoelace calculation
                                A.Ed.WriteMessage($"\n╠════════════════════════════════════════════════════════════════════════╣");
                                A.Ed.WriteMessage($"\n║ 4. TÍNH BẰNG SHOELACE FORMULA                                          ║");
                                
                                double calcArea = 0;
                                if (points.Count >= 3)
                                {
                                    for (int i = 0; i < points.Count; i++)
                                    {
                                        int j = (i + 1) % points.Count;
                                        calcArea += points[i].X * points[j].Y - points[j].X * points[i].Y;
                                    }
                                    calcArea = Math.Abs(calcArea) / 2.0;
                                }
                                A.Ed.WriteMessage($"\n║   SHOELACE AREA = {calcArea,25:F6} m²                          ║");

                                // 5. So sánh
                                A.Ed.WriteMessage($"\n╠════════════════════════════════════════════════════════════════════════╣");
                                A.Ed.WriteMessage($"\n║ 5. KẾT LUẬN                                                            ║");
                                double diff = Math.Abs(apiArea - calcArea);
                                double diffPct = apiArea > 0 ? (diff / apiArea) * 100 : 0;
                                string status = diffPct < 1 ? "✅ KHỚP" : (diffPct < 5 ? "⚠️ GẦN KHỚP" : "❌ LỆCH");
                                A.Ed.WriteMessage($"\n║   section.Area = {apiArea:F6} m² | Shoelace = {calcArea:F6} m² | {status,-9} ║");
                                A.Ed.WriteMessage($"\n║   Chiều rộng: {Math.Abs(maxX - minX):F3}m | Chiều cao: {Math.Abs(maxY - minY):F3}m                         ║");
                                A.Ed.WriteMessage($"\n╚════════════════════════════════════════════════════════════════════════╝");

                                break;
                            }
                        }
                        catch { }
                        break;
                    }
                }

                tr.Commit();
                A.Ed.WriteMessage("\n\n=== HOÀN THÀNH ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Command: Kiểm tra so sánh AREA với Civil 3D Properties Panel  
        /// </summary>
        [CommandMethod("CTSV_KiemTraKhoiLuong")]
        public static void CTSVKiemTraKhoiLuong()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                A.Ed.WriteMessage("\n╔═══════════════════════════════════════════════════════════════════════════╗");
                A.Ed.WriteMessage("\n║               KIỂM TRA SO SÁNH KHỐI LƯỢNG VỚI CIVIL 3D                    ║");
                A.Ed.WriteMessage("\n╠═══════════════════════════════════════════════════════════════════════════╣");
                A.Ed.WriteMessage("\n║ So sánh: AREA từ Material Section API vs Properties Panel                ║");
                A.Ed.WriteMessage("\n╚═══════════════════════════════════════════════════════════════════════════╝\n");

                List<AlignmentInfo> alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0) { A.Ed.WriteMessage("\n⚠️ Không có Alignment!"); return; }

                using FormChonAlignment formChon = new(alignments);
                if (Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(formChon) != DialogResult.OK) return;
                if (formChon.SelectedAlignments.Count == 0) return;

                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    A.Ed.WriteMessage($"\n\n═══════════════════════════════════════════════════════════════════════════");
                    A.Ed.WriteMessage($"\n  ALIGNMENT: {alignInfo.Name}");
                    A.Ed.WriteMessage($"\n═══════════════════════════════════════════════════════════════════════════");

                    SampleLineGroup? slg = tr.GetObject(alignInfo.SampleLineGroupId, AcadDb.OpenMode.ForWrite, false, true) as SampleLineGroup;
                    if (slg == null) continue;

                    QTOMaterialListCollection materialLists = slg.MaterialLists;
                    if (materialLists.Count == 0) { A.Ed.WriteMessage("\n  ⚠️ Không có Material List!"); continue; }

                    // Thu thập Materials
                    List<(string Name, Guid ListGuid, Guid MatGuid)> materials = new();
                    foreach (QTOMaterialList matList in materialLists)
                    {
                        foreach (QTOMaterial mat in matList)
                        {
                            materials.Add((mat.Name, matList.Guid, mat.Guid));
                        }
                    }

                    // Lấy SampleLines
                    var slIds = slg.GetSampleLineIds();
                    var sampleLines = new List<SampleLine>();
                    foreach (ObjectId slId in slIds)
                    {
                        var sl = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                        if (sl != null) sampleLines.Add(sl);
                    }
                    sampleLines = sampleLines.OrderBy(s => s.Station).ToList();

                    A.Ed.WriteMessage($"\n  Tìm thấy {sampleLines.Count} SampleLine, {materials.Count} Material\n");

                    // Header
                    A.Ed.WriteMessage($"\n  ┌────────────────┬──────────────────────────────┬──────────────┬────────────┐");
                    A.Ed.WriteMessage($"\n  │ SAMPLELINE     │ MATERIAL                     │ section.Area │ SUM AREA   │");
                    A.Ed.WriteMessage($"\n  ├────────────────┼──────────────────────────────┼──────────────┼────────────┤");

                    // Tổng theo Material
                    Dictionary<string, double> totalAreas = new();
                    Dictionary<string, double> totalVolumes = new();
                    double prevStation = 0;

                    foreach (var sl in sampleLines.Take(5)) // Hiển thị 5 đầu tiên
                    {
                        double spacing = sl.Station - prevStation;

                        foreach (var (name, listGuid, matGuid) in materials)
                        {
                            try
                            {
                                ObjectId sectionId = sl.GetMaterialSectionId(listGuid, matGuid);
                                if (sectionId.IsNull || !sectionId.IsValid) continue;

                                var sectionObj = tr.GetObject(sectionId, AcadDb.OpenMode.ForRead, false, true);
                                if (sectionObj is CivSection section)
                                {
                                    double area = section.Area;
                                    
                                    if (!totalAreas.ContainsKey(name)) totalAreas[name] = 0;
                                    totalAreas[name] += area;

                                    string shortName = name.Length > 28 ? name.Substring(0, 25) + "..." : name;
                                    string slName = sl.Name.Length > 12 ? sl.Name.Substring(0, 10) + ".." : sl.Name;
                                    A.Ed.WriteMessage($"\n  │ {slName,-14} │ {shortName,-28} │ {area,12:F4} │ {totalAreas[name],10:F4} │");
                                }
                            }
                            catch { }
                        }
                        prevStation = sl.Station;
                    }

                    A.Ed.WriteMessage($"\n  └────────────────┴──────────────────────────────┴──────────────┴────────────┘");

                    // Tổng kết
                    A.Ed.WriteMessage($"\n\n  📊 TỔNG DIỆN TÍCH (5 SampleLines đầu):");
                    foreach (var kvp in totalAreas)
                    {
                        string shortName = kvp.Key.Length > 35 ? kvp.Key.Substring(0, 32) + "..." : kvp.Key;
                        A.Ed.WriteMessage($"\n      • {shortName,-40}: {kvp.Value,12:F4} m²");
                    }

                    A.Ed.WriteMessage($"\n\n  💡 SO SÁNH VỚI CIVIL 3D:");
                    A.Ed.WriteMessage($"\n      Vui lòng kiểm tra Properties Panel của từng Material Section");
                    A.Ed.WriteMessage($"\n      (Click vào section trong Section View để xem AREA)");
                    A.Ed.WriteMessage($"\n\n      Nếu LỆCH, có thể do:");
                    A.Ed.WriteMessage($"\n      1. MaterialList chưa được Compute lại sau khi chỉnh sửa");
                    A.Ed.WriteMessage($"\n      2. Có nhiều Material Section trùng tên");
                    A.Ed.WriteMessage($"\n      3. QuantityType không phải Volume (Area, Count)");
                }

                tr.Commit();
                A.Ed.WriteMessage("\n\n=== HOÀN THÀNH ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Tính diện tích từ SectionPoints sử dụng Shoelace formula
        /// </summary>
        private static double CalculateAreaFromPoints(CivSection section)
        {
            List<Point3d> points = new();
            try
            {
                foreach (SectionPoint pt in section.SectionPoints)
                    points.Add(pt.Location);
            }
            catch { return 0; }

            if (points.Count < 3) return 0;

            // Shoelace formula
            double area = 0;
            for (int i = 0; i < points.Count; i++)
            {
                int j = (i + 1) % points.Count;
                area += points[i].X * points[j].Y;
                area -= points[j].X * points[i].Y;
            }
            return Math.Abs(area) / 2.0;
        }

        #endregion

        #region Helper Methods

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
                        SampleLineGroupId = slgIds[0],
                        IsSelected = true
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// Trích xuất khối lượng từ Material List với cú pháp Surface - Surface
        /// </summary>
        private static VolumeResult? ExtractSurfaceMaterialVolumes(Transaction tr, AlignmentInfo alignInfo)
        {
            VolumeResult result = new() { AlignmentName = alignInfo.Name };

            SampleLineGroup? slg = tr.GetObject(alignInfo.SampleLineGroupId, AcadDb.OpenMode.ForWrite, false, true) as SampleLineGroup;
            if (slg == null) return null;

            // Lấy Material Lists
            QTOMaterialListCollection materialLists = slg.MaterialLists;
            if (materialLists.Count == 0)
            {
                A.Ed.WriteMessage($"\n⚠️ Không có Material List trong '{slg.Name}'");
                return null;
            }

            // Thu thập Materials với thông tin chi tiết
            List<SurfaceMaterialInfo> materials = new();
            foreach (QTOMaterialList materialList in materialLists)
            {
                try
                {
                    Guid listGuid = materialList.Guid;
                    foreach (QTOMaterial material in materialList)
                    {
                        SurfaceMaterialInfo matInfo = new()
                        {
                            Name = material.Name,
                            MaterialListGuid = listGuid,
                            MaterialGuid = material.Guid,
                            Material = material,
                            QuantityType = material.QuantityType
                        };

                        // Phân tích tên material để xác định Cut/Fill
                        string nameLower = material.Name.ToLower();
                        matInfo.IsCut = nameLower.Contains("cut") || nameLower.Contains("đào") || 
                                       nameLower.Contains("dao") || nameLower.Contains("excavation");
                        matInfo.IsFill = nameLower.Contains("fill") || nameLower.Contains("đắp") || 
                                        nameLower.Contains("dap") || nameLower.Contains("embankment");

                        // Nếu không xác định được, mặc định là Cut
                        if (!matInfo.IsCut && !matInfo.IsFill)
                        {
                            matInfo.IsCut = true;
                        }

                        materials.Add(matInfo);
                    }
                }
                catch { }
            }

            // Lấy và sắp xếp SampleLines theo lý trình
            List<SampleLine> sortedSampleLines = new();
            foreach (ObjectId slId in slg.GetSampleLineIds())
            {
                var sl = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                if (sl != null) sortedSampleLines.Add(sl);
            }
            sortedSampleLines = sortedSampleLines.OrderBy(s => s.Station).ToList();

            // Duyệt qua từng SampleLine
            double prevStation = 0;
            bool isFirst = true;
            CrossSectionVolumeInfo? prevCrossSection = null;

            foreach (SampleLine sampleLine in sortedSampleLines)
            {
                CrossSectionVolumeInfo csInfo = new()
                {
                    SampleLineName = sampleLine.Name,
                    Station = sampleLine.Station,
                    StationFormatted = FormatStation(sampleLine.Station),
                    SpacingPrev = isFirst ? 0 : (sampleLine.Station - prevStation)
                };

                // Lấy MaterialSection cho từng Material
                foreach (var matInfo in materials)
                {
                    try
                    {
                        ObjectId materialSectionId = sampleLine.GetMaterialSectionId(
                            matInfo.MaterialListGuid, 
                            matInfo.MaterialGuid
                        );

                        if (!materialSectionId.IsNull && materialSectionId.IsValid)
                        {
                            AcadDb.DBObject? sectionObj = tr.GetObject(materialSectionId, AcadDb.OpenMode.ForRead, false, true);

                            if (sectionObj is CivSection section)
                            {
                                // === LẤY TẤT CẢ THÔNG TIN TỪ MATERIAL SECTION ===
                                
                                // Lấy diện tích
                                double areaFromAPI = section.Area;
                                double areaCalculated = CalculateSectionArea(section);
                                double area = areaFromAPI > 0 ? areaFromAPI : areaCalculated;
                                
                                // Lấy các điểm section và tính toán các thuộc tính
                                List<Point3d> points = new();
                                double minOffset = double.MaxValue;  // Left Length
                                double maxOffset = double.MinValue;  // Right Length  
                                double minElevation = double.MaxValue;
                                double maxElevation = double.MinValue;
                                
                                try
                                {
                                    foreach (SectionPoint pt in section.SectionPoints)
                                    {
                                        points.Add(pt.Location);
                                        
                                        // Offset (X) - âm = bên trái, dương = bên phải
                                        if (pt.Location.X < minOffset) minOffset = pt.Location.X;
                                        if (pt.Location.X > maxOffset) maxOffset = pt.Location.X;
                                        
                                        // Elevation (Y)
                                        if (pt.Location.Y < minElevation) minElevation = pt.Location.Y;
                                        if (pt.Location.Y > maxElevation) maxElevation = pt.Location.Y;
                                    }
                                }
                                catch { }
                                
                                // Tạo MaterialSectionData với đầy đủ thông tin
                                MaterialSectionData sectionData = new()
                                {
                                    MaterialName = matInfo.Name,
                                    SectionSurfaceName = section.Name,
                                    SectionType = "MaterialSection",
                                    StaticDynamic = "Dynamic",
                                    SectionStation = csInfo.Station,
                                    SectionStationFormatted = csInfo.StationFormatted,
                                    LeftLength = minOffset != double.MaxValue ? minOffset : 0,
                                    RightLength = maxOffset != double.MinValue ? maxOffset : 0,
                                    SectionMinElevation = minElevation != double.MaxValue ? minElevation : 0,
                                    SectionMaxElevation = maxElevation != double.MinValue ? maxElevation : 0,
                                    Area = area,
                                    Points = points,
                                    IsCut = matInfo.IsCut,
                                    IsFill = matInfo.IsFill
                                };
                                
                                // Lưu vào dictionary
                                csInfo.MaterialSections[matInfo.Name] = sectionData;
                                csInfo.SectionPoints[matInfo.Name] = points;

                                if (area > 0)
                                {
                                    // Phân loại vào Cut hoặc Fill
                                    if (matInfo.IsFill)
                                    {
                                        csInfo.FillAreas[matInfo.Name] = area;
                                    }
                                    else
                                    {
                                        csInfo.CutAreas[matInfo.Name] = area;
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                }

                // Tính khối lượng từ trắc ngang trước đến trắc ngang này
                if (prevCrossSection != null)
                {
                    double spacing = csInfo.SpacingPrev;

                    // Tính khối lượng Cut
                    foreach (var mat in csInfo.CutAreas.Keys.Union(prevCrossSection.CutAreas.Keys))
                    {
                        double areaCurrent = csInfo.CutAreas.GetValueOrDefault(mat, 0);
                        double areaPrev = prevCrossSection.CutAreas.GetValueOrDefault(mat, 0);
                        double volume = CalculateVolume(areaPrev, areaCurrent, spacing);
                        csInfo.CutVolumes[mat] = volume;
                    }

                    // Tính khối lượng Fill
                    foreach (var mat in csInfo.FillAreas.Keys.Union(prevCrossSection.FillAreas.Keys))
                    {
                        double areaCurrent = csInfo.FillAreas.GetValueOrDefault(mat, 0);
                        double areaPrev = prevCrossSection.FillAreas.GetValueOrDefault(mat, 0);
                        double volume = CalculateVolume(areaPrev, areaCurrent, spacing);
                        csInfo.FillVolumes[mat] = volume;
                    }
                }

                result.CrossSections.Add(csInfo);
                prevCrossSection = csInfo;
                prevStation = sampleLine.Station;
                isFirst = false;
            }

            // Tính tổng khối lượng
            foreach (var cs in result.CrossSections)
            {
                foreach (var kvp in cs.CutVolumes)
                {
                    if (!result.TotalCutVolumes.ContainsKey(kvp.Key))
                        result.TotalCutVolumes[kvp.Key] = 0;
                    result.TotalCutVolumes[kvp.Key] += kvp.Value;
                }

                foreach (var kvp in cs.FillVolumes)
                {
                    if (!result.TotalFillVolumes.ContainsKey(kvp.Key))
                        result.TotalFillVolumes[kvp.Key] = 0;
                    result.TotalFillVolumes[kvp.Key] += kvp.Value;
                }
            }

            return result;
        }

        /// <summary>
        /// Tính khối lượng theo phương pháp Average End Area
        /// </summary>
        private static double CalculateVolume(double area1, double area2, double distance)
        {
            return ((area1 + area2) / 2.0) * distance;
        }

        /// <summary>
        /// Tính diện tích Section từ SectionPoints (công thức Shoelace)
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

                // Đóng đa giác nếu chưa đóng
                if (pointList.Count >= 2)
                {
                    Point2d first = pointList[0];
                    Point2d last = pointList[pointList.Count - 1];

                    if (Math.Abs(first.X - last.X) > 0.001 || Math.Abs(first.Y - last.Y) > 0.001)
                    {
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
        /// Tính diện tích đa giác bằng công thức Shoelace
        /// </summary>
        private static double CalculatePolygonArea(List<Point2d> points)
        {
            if (points.Count < 3) return 0;

            double area = 0;
            for (int i = 0; i < points.Count; i++)
            {
                int j = (i + 1) % points.Count;
                area += points[i].X * points[j].Y;
                area -= points[j].X * points[i].Y;
            }

            return Math.Abs(area / 2.0);
        }

        /// <summary>
        /// Format lý trình theo chuẩn Việt Nam
        /// </summary>
        private static string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double m = station % 1000;
            return $"{km}+{m:000.00}";
        }

        #endregion

        #region Export Methods

        /// <summary>
        /// Xuất kết quả đào đắp ra Excel
        /// </summary>
        private static void ExportDaoDapToExcel(string filePath, List<VolumeResult> results, List<string> materials)
        {
            using var workbook = new XLWorkbook();

            // Sheet cho từng Alignment
            foreach (var result in results)
            {
                string sheetName = SanitizeSheetName($"DD_{result.AlignmentName}");
                var ws = workbook.Worksheets.Add(sheetName);

                // Tiêu đề
                ws.Cell(1, 1).Value = $"BẢNG TÍNH KHỐI LƯỢNG ĐÀO ĐẮP - {result.AlignmentName}";
                int lastCol = 4 + materials.Count * 4; // STT, Tên, LT, KC + (CutArea, CutVol, FillArea, FillVol) * materials
                ws.Range(1, 1, 1, lastCol).Merge();
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Font.FontSize = 14;
                ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Header hàng 2 - Nhóm
                ws.Cell(2, 1).Value = "THÔNG TIN TRẮC NGANG";
                ws.Range(2, 1, 2, 4).Merge();
                ws.Cell(2, 1).Style.Font.Bold = true;
                ws.Cell(2, 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                ws.Cell(2, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int col = 5;
                foreach (var mat in materials)
                {
                    // Mỗi material có 4 cột
                    ws.Cell(2, col).Value = mat;
                    ws.Range(2, col, 2, col + 3).Merge();
                    ws.Cell(2, col).Style.Font.Bold = true;
                    ws.Cell(2, col).Style.Fill.BackgroundColor = XLColor.LightGreen;
                    ws.Cell(2, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    col += 4;
                }

                // Header hàng 3 - Chi tiết
                ws.Cell(3, 1).Value = "STT";
                ws.Cell(3, 2).Value = "TÊN TRẮC NGANG";
                ws.Cell(3, 3).Value = "LÝ TRÌNH";
                ws.Cell(3, 4).Value = "K.CÁCH (m)";

                col = 5;
                foreach (var mat in materials)
                {
                    ws.Cell(3, col).Value = "DT Đào (m²)";
                    ws.Cell(3, col + 1).Value = "KL Đào (m³)";
                    ws.Cell(3, col + 2).Value = "DT Đắp (m²)";
                    ws.Cell(3, col + 3).Value = "KL Đắp (m³)";
                    col += 4;
                }

                // Format header hàng 3
                ws.Range(3, 1, 3, lastCol).Style.Font.Bold = true;
                ws.Range(3, 1, 3, lastCol).Style.Fill.BackgroundColor = XLColor.LightBlue;
                ws.Range(3, 1, 3, lastCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Dữ liệu
                int row = 4;
                int stt = 1;
                Dictionary<string, double> totalCutAreas = materials.ToDictionary(m => m, m => 0.0);
                Dictionary<string, double> totalCutVolumes = materials.ToDictionary(m => m, m => 0.0);
                Dictionary<string, double> totalFillAreas = materials.ToDictionary(m => m, m => 0.0);
                Dictionary<string, double> totalFillVolumes = materials.ToDictionary(m => m, m => 0.0);

                foreach (var cs in result.CrossSections)
                {
                    ws.Cell(row, 1).Value = stt++;
                    ws.Cell(row, 2).Value = cs.SampleLineName;
                    ws.Cell(row, 3).Value = cs.StationFormatted;
                    ws.Cell(row, 4).Value = Math.Round(cs.SpacingPrev, 3);

                    col = 5;
                    foreach (var mat in materials)
                    {
                        double cutArea = cs.CutAreas.GetValueOrDefault(mat, 0);
                        double cutVol = cs.CutVolumes.GetValueOrDefault(mat, 0);
                        double fillArea = cs.FillAreas.GetValueOrDefault(mat, 0);
                        double fillVol = cs.FillVolumes.GetValueOrDefault(mat, 0);

                        ws.Cell(row, col).Value = Math.Round(cutArea, 4);
                        ws.Cell(row, col + 1).Value = Math.Round(cutVol, 4);
                        ws.Cell(row, col + 2).Value = Math.Round(fillArea, 4);
                        ws.Cell(row, col + 3).Value = Math.Round(fillVol, 4);

                        totalCutAreas[mat] += cutArea;
                        totalCutVolumes[mat] += cutVol;
                        totalFillAreas[mat] += fillArea;
                        totalFillVolumes[mat] += fillVol;

                        col += 4;
                    }

                    row++;
                }

                // Hàng tổng cộng
                ws.Cell(row, 1).Value = "TỔNG CỘNG";
                ws.Range(row, 1, row, 4).Merge();
                ws.Cell(row, 1).Style.Font.Bold = true;
                ws.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.LightYellow;

                col = 5;
                foreach (var mat in materials)
                {
                    ws.Cell(row, col).Value = Math.Round(totalCutAreas[mat], 4);
                    ws.Cell(row, col + 1).Value = Math.Round(totalCutVolumes[mat], 4);
                    ws.Cell(row, col + 2).Value = Math.Round(totalFillAreas[mat], 4);
                    ws.Cell(row, col + 3).Value = Math.Round(totalFillVolumes[mat], 4);

                    ws.Cell(row, col).Style.Font.Bold = true;
                    ws.Cell(row, col + 1).Style.Font.Bold = true;
                    ws.Cell(row, col + 2).Style.Font.Bold = true;
                    ws.Cell(row, col + 3).Style.Font.Bold = true;

                    col += 4;
                }

                // Border và format
                ws.Range(2, 1, row, lastCol).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(2, 1, row, lastCol).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                ws.Column(1).Width = 8;
                ws.Column(2).Width = 20;
                ws.Column(3).Width = 15;
                ws.Column(4).Width = 12;

                for (int c = 5; c <= lastCol; c++)
                {
                    ws.Column(c).Width = 14;
                }

                // Tạo sheet chi tiết Material Section Data
                CreateMaterialSectionDetailSheet(workbook, result);
            }

            // Sheet tổng hợp
            if (results.Count > 1)
            {
                CreateSummarySheetDaoDap(workbook, results, materials);
            }

            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Tạo sheet chi tiết thông tin Material Section (Left/Right Length, Elevation, etc.)
        /// </summary>
        private static void CreateMaterialSectionDetailSheet(XLWorkbook workbook, VolumeResult result)
        {
            string sheetName = SanitizeSheetName($"CT_{result.AlignmentName}");
            var ws = workbook.Worksheets.Add(sheetName);

            // Tiêu đề
            ws.Cell(1, 1).Value = $"CHI TIẾT MATERIAL SECTION - {result.AlignmentName}";
            ws.Range(1, 1, 1, 12).Merge();
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontSize = 14;
            ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Header
            ws.Cell(2, 1).Value = "STT";
            ws.Cell(2, 2).Value = "TÊN TRẮC NGANG";
            ws.Cell(2, 3).Value = "LÝ TRÌNH";
            ws.Cell(2, 4).Value = "MATERIAL";
            ws.Cell(2, 5).Value = "LOẠI";
            ws.Cell(2, 6).Value = "LEFT LENGTH (m)";
            ws.Cell(2, 7).Value = "RIGHT LENGTH (m)";
            ws.Cell(2, 8).Value = "TỔNG RỘNG (m)";
            ws.Cell(2, 9).Value = "MIN ELEV (m)";
            ws.Cell(2, 10).Value = "MAX ELEV (m)";
            ws.Cell(2, 11).Value = "C.CAO (m)";
            ws.Cell(2, 12).Value = "DIỆN TÍCH (m²)";

            ws.Range(2, 1, 2, 12).Style.Font.Bold = true;
            ws.Range(2, 1, 2, 12).Style.Fill.BackgroundColor = XLColor.LightBlue;
            ws.Range(2, 1, 2, 12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Dữ liệu
            int row = 3;
            int stt = 1;

            foreach (var cs in result.CrossSections)
            {
                bool isFirst = true;
                foreach (var kvp in cs.MaterialSections)
                {
                    var data = kvp.Value;

                    ws.Cell(row, 1).Value = stt;
                    ws.Cell(row, 2).Value = isFirst ? cs.SampleLineName : "";
                    ws.Cell(row, 3).Value = isFirst ? cs.StationFormatted : "";
                    ws.Cell(row, 4).Value = data.MaterialName;
                    ws.Cell(row, 5).Value = data.IsCut ? "ĐÀO" : "ĐẮP";
                    ws.Cell(row, 6).Value = Math.Round(data.LeftLength, 3);
                    ws.Cell(row, 7).Value = Math.Round(data.RightLength, 3);
                    ws.Cell(row, 8).Value = Math.Round(data.TotalWidth, 3);
                    ws.Cell(row, 9).Value = Math.Round(data.SectionMinElevation, 3);
                    ws.Cell(row, 10).Value = Math.Round(data.SectionMaxElevation, 3);
                    ws.Cell(row, 11).Value = Math.Round(data.Height, 3);
                    ws.Cell(row, 12).Value = Math.Round(data.Area, 4);

                    // Style cho loại
                    if (data.IsCut)
                        ws.Cell(row, 5).Style.Fill.BackgroundColor = XLColor.LightSalmon;
                    else
                        ws.Cell(row, 5).Style.Fill.BackgroundColor = XLColor.LightGreen;

                    row++;
                    isFirst = false;
                }
                stt++;
            }

            // Border
            ws.Range(2, 1, row - 1, 12).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(2, 1, row - 1, 12).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Column widths
            ws.Column(1).Width = 8;
            ws.Column(2).Width = 20;
            ws.Column(3).Width = 15;
            ws.Column(4).Width = 20;
            ws.Column(5).Width = 10;
            for (int c = 6; c <= 12; c++)
                ws.Column(c).Width = 15;
        }

        /// <summary>
        /// Tạo sheet tổng hợp đào đắp
        /// </summary>
        private static void CreateSummarySheetDaoDap(XLWorkbook workbook, List<VolumeResult> results, List<string> materials)
        {
            var ws = workbook.Worksheets.Add("TỔNG HỢP");

            ws.Cell(1, 1).Value = "TỔNG HỢP KHỐI LƯỢNG ĐÀO ĐẮP TẤT CẢ CÁC TUYẾN";
            int lastCol = 1 + materials.Count * 2;
            ws.Range(1, 1, 1, lastCol).Merge();
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontSize = 14;
            ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Header
            ws.Cell(2, 1).Value = "TUYẾN";
            int col = 2;
            foreach (var mat in materials)
            {
                ws.Cell(2, col).Value = $"{mat} - ĐÀO (m³)";
                ws.Cell(2, col + 1).Value = $"{mat} - ĐẮP (m³)";
                col += 2;
            }

            ws.Range(2, 1, 2, lastCol).Style.Font.Bold = true;
            ws.Range(2, 1, 2, lastCol).Style.Fill.BackgroundColor = XLColor.LightGreen;

            // Dữ liệu
            int row = 3;
            Dictionary<string, double> grandCutTotals = materials.ToDictionary(m => m, m => 0.0);
            Dictionary<string, double> grandFillTotals = materials.ToDictionary(m => m, m => 0.0);

            foreach (var result in results)
            {
                ws.Cell(row, 1).Value = result.AlignmentName;

                col = 2;
                foreach (var mat in materials)
                {
                    double cutTotal = result.TotalCutVolumes.GetValueOrDefault(mat, 0);
                    double fillTotal = result.TotalFillVolumes.GetValueOrDefault(mat, 0);

                    ws.Cell(row, col).Value = Math.Round(cutTotal, 4);
                    ws.Cell(row, col + 1).Value = Math.Round(fillTotal, 4);

                    grandCutTotals[mat] += cutTotal;
                    grandFillTotals[mat] += fillTotal;

                    col += 2;
                }

                row++;
            }

            // Tổng cộng
            ws.Cell(row, 1).Value = "TỔNG CỘNG";
            ws.Cell(row, 1).Style.Font.Bold = true;

            col = 2;
            foreach (var mat in materials)
            {
                ws.Cell(row, col).Value = Math.Round(grandCutTotals[mat], 4);
                ws.Cell(row, col + 1).Value = Math.Round(grandFillTotals[mat], 4);
                ws.Cell(row, col).Style.Font.Bold = true;
                ws.Cell(row, col + 1).Style.Font.Bold = true;
                col += 2;
            }

            // Border
            ws.Range(2, 1, row, lastCol).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(2, 1, row, lastCol).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Column(1).Width = 25;
            for (int c = 2; c <= lastCol; c++)
            {
                ws.Column(c).Width = 18;
            }
        }

        /// <summary>
        /// Tạo bảng CAD cho đào đắp
        /// </summary>
        private static void CreateDaoDapCadTable(Transaction tr, Point3d insertPoint, VolumeResult result, List<string> materials)
        {
            AcadDb.Database db = HostApplicationServices.WorkingDatabase;
            AcadDb.BlockTable bt = tr.GetObject(db.BlockTableId, AcadDb.OpenMode.ForRead) as AcadDb.BlockTable
                ?? throw new System.Exception("Không thể mở BlockTable");
            AcadDb.BlockTableRecord btr = tr.GetObject(bt[AcadDb.BlockTableRecord.ModelSpace], AcadDb.OpenMode.ForWrite) as AcadDb.BlockTableRecord
                ?? throw new System.Exception("Không thể mở ModelSpace");

            // Tính số cột và hàng
            int numCols = 4 + materials.Count * 2; // STT, Tên, LT, KC + (Đào, Đắp) * materials
            int numRows = result.CrossSections.Count + 4; // 2 header + dữ liệu + 1 tổng cộng

            // Tạo Table
            AcadDb.Table table = new()
            {
                Position = insertPoint,
                TableStyle = db.Tablestyle
            };

            table.SetSize(numRows, numCols);

            // Kích thước
            for (int r = 0; r < numRows; r++)
                table.Rows[r].Height = 8.0;

            table.Columns[0].Width = 10;
            table.Columns[1].Width = 25;
            table.Columns[2].Width = 18;
            table.Columns[3].Width = 12;

            for (int c = 4; c < numCols; c++)
                table.Columns[c].Width = 18;

            // HÀNG 0: Tiêu đề
            table.MergeCells(CellRange.Create(table, 0, 0, 0, numCols - 1));
            table.Cells[0, 0].TextString = $"KHỐI LƯỢNG ĐÀO ĐẮP - {result.AlignmentName}";
            table.Cells[0, 0].TextHeight = 5.0;
            table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

            // HÀNG 1: Nhóm header
            table.MergeCells(CellRange.Create(table, 1, 0, 1, 3));
            table.Cells[1, 0].TextString = "THÔNG TIN TRẮC NGANG";
            table.Cells[1, 0].Alignment = CellAlignment.MiddleCenter;

            int col = 4;
            foreach (var mat in materials)
            {
                table.MergeCells(CellRange.Create(table, 1, col, 1, col + 1));
                table.Cells[1, col].TextString = mat;
                table.Cells[1, col].Alignment = CellAlignment.MiddleCenter;
                col += 2;
            }

            // HÀNG 2: Header chi tiết
            table.Cells[2, 0].TextString = "STT";
            table.Cells[2, 1].TextString = "TÊN";
            table.Cells[2, 2].TextString = "LÝ TRÌNH";
            table.Cells[2, 3].TextString = "K.CÁCH";

            col = 4;
            foreach (var mat in materials)
            {
                table.Cells[2, col].TextString = "ĐÀO (m³)";
                table.Cells[2, col + 1].TextString = "ĐẮP (m³)";
                table.Cells[2, col].Alignment = CellAlignment.MiddleCenter;
                table.Cells[2, col + 1].Alignment = CellAlignment.MiddleCenter;
                col += 2;
            }

            for (int c = 0; c < numCols; c++)
            {
                table.Cells[2, c].TextHeight = 3.5;
                table.Cells[2, c].Alignment = CellAlignment.MiddleCenter;
            }

            // DỮ LIỆU
            int row = 3;
            int stt = 1;
            Dictionary<string, double> totalCut = materials.ToDictionary(m => m, m => 0.0);
            Dictionary<string, double> totalFill = materials.ToDictionary(m => m, m => 0.0);

            foreach (var cs in result.CrossSections)
            {
                table.Cells[row, 0].TextString = stt++.ToString();
                table.Cells[row, 1].TextString = cs.SampleLineName;
                table.Cells[row, 2].TextString = cs.StationFormatted;
                table.Cells[row, 3].TextString = Math.Round(cs.SpacingPrev, 2).ToString();

                col = 4;
                foreach (var mat in materials)
                {
                    double cutVol = cs.CutVolumes.GetValueOrDefault(mat, 0);
                    double fillVol = cs.FillVolumes.GetValueOrDefault(mat, 0);

                    table.Cells[row, col].TextString = Math.Round(cutVol, 3).ToString();
                    table.Cells[row, col + 1].TextString = Math.Round(fillVol, 3).ToString();
                    table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                    table.Cells[row, col + 1].Alignment = CellAlignment.MiddleRight;

                    totalCut[mat] += cutVol;
                    totalFill[mat] += fillVol;

                    col += 2;
                }

                row++;
            }

            // HÀNG TỔNG CỘNG
            table.MergeCells(CellRange.Create(table, row, 0, row, 3));
            table.Cells[row, 0].TextString = "TỔNG CỘNG";
            table.Cells[row, 0].Alignment = CellAlignment.MiddleCenter;

            col = 4;
            foreach (var mat in materials)
            {
                table.Cells[row, col].TextString = Math.Round(totalCut[mat], 3).ToString();
                table.Cells[row, col + 1].TextString = Math.Round(totalFill[mat], 3).ToString();
                table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                table.Cells[row, col + 1].Alignment = CellAlignment.MiddleRight;
                col += 2;
            }

            // TextHeight cho dữ liệu
            for (int r = 3; r < numRows; r++)
            {
                for (int c = 0; c < numCols; c++)
                {
                    table.Cells[r, c].TextHeight = 3.0;
                }
            }

            btr.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
        }

        /// <summary>
        /// Tạo bảng CAD chi tiết Material Section (Left/Right Length, Elevation, Area)
        /// </summary>
        private static void CreateDetailCadTable(Transaction tr, Point3d insertPoint, VolumeResult result)
        {
            AcadDb.Database db = HostApplicationServices.WorkingDatabase;
            AcadDb.BlockTable bt = tr.GetObject(db.BlockTableId, AcadDb.OpenMode.ForRead) as AcadDb.BlockTable
                ?? throw new System.Exception("Không thể mở BlockTable");
            AcadDb.BlockTableRecord btr = tr.GetObject(bt[AcadDb.BlockTableRecord.ModelSpace], AcadDb.OpenMode.ForWrite) as AcadDb.BlockTableRecord
                ?? throw new System.Exception("Không thể mở ModelSpace");

            // Đếm số dòng cần thiết
            int totalRows = 2; // Header
            foreach (var cs in result.CrossSections)
            {
                totalRows += cs.MaterialSections.Count;
            }

            int numCols = 10;
            
            // Tạo Table
            AcadDb.Table table = new()
            {
                Position = insertPoint,
                TableStyle = db.Tablestyle
            };

            table.SetSize(totalRows, numCols);

            // Kích thước
            for (int r = 0; r < totalRows; r++)
                table.Rows[r].Height = 7.0;

            table.Columns[0].Width = 8;   // STT
            table.Columns[1].Width = 18;  // Tên TN
            table.Columns[2].Width = 14;  // Lý trình
            table.Columns[3].Width = 18;  // Material
            table.Columns[4].Width = 10;  // Loại
            table.Columns[5].Width = 12;  // Left
            table.Columns[6].Width = 12;  // Right
            table.Columns[7].Width = 12;  // Min Elev
            table.Columns[8].Width = 12;  // Max Elev
            table.Columns[9].Width = 14;  // Area

            // HÀNG 0: Tiêu đề
            table.MergeCells(CellRange.Create(table, 0, 0, 0, numCols - 1));
            table.Cells[0, 0].TextString = $"CHI TIẾT MATERIAL SECTION - {result.AlignmentName}";
            table.Cells[0, 0].TextHeight = 4.5;
            table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

            // HÀNG 1: Header
            table.Cells[1, 0].TextString = "STT";
            table.Cells[1, 1].TextString = "TÊN T.NGANG";
            table.Cells[1, 2].TextString = "LÝ TRÌNH";
            table.Cells[1, 3].TextString = "MATERIAL";
            table.Cells[1, 4].TextString = "LOẠI";
            table.Cells[1, 5].TextString = "LEFT (m)";
            table.Cells[1, 6].TextString = "RIGHT (m)";
            table.Cells[1, 7].TextString = "MIN ELEV";
            table.Cells[1, 8].TextString = "MAX ELEV";
            table.Cells[1, 9].TextString = "AREA (m²)";

            for (int c = 0; c < numCols; c++)
            {
                table.Cells[1, c].TextHeight = 3.0;
                table.Cells[1, c].Alignment = CellAlignment.MiddleCenter;
            }

            // DỮ LIỆU
            int row = 2;
            int stt = 1;

            foreach (var cs in result.CrossSections)
            {
                bool isFirst = true;
                foreach (var kvp in cs.MaterialSections)
                {
                    var data = kvp.Value;

                    table.Cells[row, 0].TextString = stt.ToString();
                    table.Cells[row, 1].TextString = isFirst ? cs.SampleLineName : "";
                    table.Cells[row, 2].TextString = isFirst ? cs.StationFormatted : "";
                    table.Cells[row, 3].TextString = data.MaterialName;
                    table.Cells[row, 4].TextString = data.IsCut ? "ĐÀO" : "ĐẮP";
                    table.Cells[row, 5].TextString = Math.Round(data.LeftLength, 3).ToString("F3");
                    table.Cells[row, 6].TextString = Math.Round(data.RightLength, 3).ToString("F3");
                    table.Cells[row, 7].TextString = Math.Round(data.SectionMinElevation, 3).ToString("F3");
                    table.Cells[row, 8].TextString = Math.Round(data.SectionMaxElevation, 3).ToString("F3");
                    table.Cells[row, 9].TextString = Math.Round(data.Area, 4).ToString("F4");

                    // Text alignment
                    table.Cells[row, 0].Alignment = CellAlignment.MiddleCenter;
                    table.Cells[row, 1].Alignment = CellAlignment.MiddleLeft;
                    table.Cells[row, 2].Alignment = CellAlignment.MiddleCenter;
                    table.Cells[row, 3].Alignment = CellAlignment.MiddleLeft;
                    table.Cells[row, 4].Alignment = CellAlignment.MiddleCenter;
                    for (int c = 5; c < numCols; c++)
                        table.Cells[row, c].Alignment = CellAlignment.MiddleRight;

                    // Text height
                    for (int c = 0; c < numCols; c++)
                        table.Cells[row, c].TextHeight = 2.5;

                    row++;
                    isFirst = false;
                }
                stt++;
            }

            btr.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
        }

        /// <summary>
        /// Xuất khối lượng trắc ngang ra Excel (lấy AREA từ Material Section)
        /// </summary>
        private static void ExportTracNgangToExcel(string filePath, List<VolumeResult> results, List<string> materials)
        {
            using var workbook = new XLWorkbook();

            foreach (var result in results)
            {
                string sheetName = SanitizeSheetName(result.AlignmentName);
                var ws = workbook.Worksheets.Add(sheetName);

                // Tính số cột: STT, Tên, Lý trình, K.Cách, + cho mỗi Material (Area, DT TB, K.Lượng)
                int numMaterialCols = materials.Count * 3;
                int totalCols = 4 + numMaterialCols;

                // Tiêu đề chính
                ws.Cell(1, 1).Value = $"BẢNG TÍNH KHỐI LƯỢNG TRẮC NGANG - {result.AlignmentName}";
                ws.Range(1, 1, 1, totalCols).Merge();
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Font.FontSize = 14;
                ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Header hàng 2 - Nhóm
                ws.Cell(2, 1).Value = "THÔNG TIN TRẮC NGANG";
                ws.Range(2, 1, 2, 4).Merge();
                ws.Cell(2, 1).Style.Font.Bold = true;
                ws.Cell(2, 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                ws.Cell(2, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int col = 5;
                foreach (var mat in materials)
                {
                    ws.Cell(2, col).Value = mat;
                    ws.Range(2, col, 2, col + 2).Merge();
                    ws.Cell(2, col).Style.Font.Bold = true;
                    ws.Cell(2, col).Style.Fill.BackgroundColor = XLColor.LightGreen;
                    ws.Cell(2, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    col += 3;
                }

                // Header hàng 3 - Chi tiết
                ws.Cell(3, 1).Value = "STT";
                ws.Cell(3, 2).Value = "TÊN TRẮC NGANG";
                ws.Cell(3, 3).Value = "LÝ TRÌNH";
                ws.Cell(3, 4).Value = "K.CÁCH (m)";

                col = 5;
                foreach (var mat in materials)
                {
                    ws.Cell(3, col).Value = "AREA (m²)";
                    ws.Cell(3, col + 1).Value = "DT TB (m²)";
                    ws.Cell(3, col + 2).Value = "K.LƯỢNG (m³)";
                    col += 3;
                }

                ws.Range(3, 1, 3, totalCols).Style.Font.Bold = true;
                ws.Range(3, 1, 3, totalCols).Style.Fill.BackgroundColor = XLColor.LightBlue;
                ws.Range(3, 1, 3, totalCols).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Dữ liệu
                int row = 4;
                int stt = 1;
                Dictionary<string, double> totalVolumes = materials.ToDictionary(m => m, m => 0.0);

                foreach (var cs in result.CrossSections)
                {
                    ws.Cell(row, 1).Value = stt++;
                    ws.Cell(row, 2).Value = cs.SampleLineName;
                    ws.Cell(row, 3).Value = cs.StationFormatted;
                    ws.Cell(row, 4).Value = Math.Round(cs.SpacingPrev, 3);

                    col = 5;
                    foreach (var mat in materials)
                    {
                        // Lấy Area từ MaterialSections
                        double area = 0;
                        if (cs.MaterialSections.ContainsKey(mat))
                            area = cs.MaterialSections[mat].Area;

                        ws.Cell(row, col).Value = Math.Round(area, 4);

                        // Tính DT Trung bình và Khối lượng bằng công thức Excel
                        string areaColLetter = GetColumnLetter(col);
                        
                        if (row == 4)
                        {
                            ws.Cell(row, col + 1).Value = 0;
                            ws.Cell(row, col + 2).Value = 0;
                        }
                        else
                        {
                            // DT TB = (Area_trước + Area_hiện) / 2
                            ws.Cell(row, col + 1).FormulaA1 = $"=({areaColLetter}{row-1}+{areaColLetter}{row})/2";
                            
                            // K.Lượng = DT TB × Khoảng cách
                            string avgColLetter = GetColumnLetter(col + 1);
                            ws.Cell(row, col + 2).FormulaA1 = $"={avgColLetter}{row}*D{row}";
                        }

                        col += 3;
                    }
                    row++;
                }

                // Hàng tổng cộng
                ws.Cell(row, 1).Value = "TỔNG CỘNG";
                ws.Range(row, 1, row, 4).Merge();
                ws.Cell(row, 1).Style.Font.Bold = true;
                ws.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.LightYellow;

                col = 5;
                foreach (var mat in materials)
                {
                    // Tổng Area (bỏ trống)
                    ws.Cell(row, col).Value = "";
                    // Tổng DT TB (bỏ trống)
                    ws.Cell(row, col + 1).Value = "";
                    // Tổng K.Lượng = SUM
                    string volColLetter = GetColumnLetter(col + 2);
                    ws.Cell(row, col + 2).FormulaA1 = $"=SUM({volColLetter}4:{volColLetter}{row-1})";
                    ws.Cell(row, col + 2).Style.Font.Bold = true;
                    ws.Cell(row, col + 2).Style.Fill.BackgroundColor = XLColor.LightYellow;

                    col += 3;
                }

                // Format
                ws.Range(2, 1, row, totalCols).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(2, 1, row, totalCols).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                ws.Column(1).Width = 8;
                ws.Column(2).Width = 20;
                ws.Column(3).Width = 15;
                ws.Column(4).Width = 12;
                for (int c = 5; c <= totalCols; c++)
                    ws.Column(c).Width = 14;
            }

            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Tạo bảng CAD khối lượng trắc ngang
        /// </summary>
        private static void CreateTracNgangCadTable(Transaction tr, Point3d insertPoint, VolumeResult result, List<string> materials)
        {
            AcadDb.Database db = HostApplicationServices.WorkingDatabase;
            AcadDb.BlockTable bt = tr.GetObject(db.BlockTableId, AcadDb.OpenMode.ForRead) as AcadDb.BlockTable
                ?? throw new System.Exception("Không thể mở BlockTable");
            AcadDb.BlockTableRecord btr = tr.GetObject(bt[AcadDb.BlockTableRecord.ModelSpace], AcadDb.OpenMode.ForWrite) as AcadDb.BlockTableRecord
                ?? throw new System.Exception("Không thể mở ModelSpace");

            int numCols = 4 + materials.Count * 2; // STT, Tên, LT, KC + (Area, Vol) * materials
            int numRows = result.CrossSections.Count + 4; // 2 header + dữ liệu + tổng cộng

            AcadDb.Table table = new()
            {
                Position = insertPoint,
                TableStyle = db.Tablestyle
            };

            table.SetSize(numRows, numCols);

            // Kích thước
            for (int r = 0; r < numRows; r++)
                table.Rows[r].Height = 7.0;

            table.Columns[0].Width = 8;
            table.Columns[1].Width = 18;
            table.Columns[2].Width = 14;
            table.Columns[3].Width = 10;
            for (int c = 4; c < numCols; c++)
                table.Columns[c].Width = 14;

            // Hàng 0: Tiêu đề
            table.MergeCells(CellRange.Create(table, 0, 0, 0, numCols - 1));
            table.Cells[0, 0].TextString = $"KHỐI LƯỢNG TRẮC NGANG - {result.AlignmentName}";
            table.Cells[0, 0].TextHeight = 4.5;
            table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

            // Hàng 1: Nhóm
            table.MergeCells(CellRange.Create(table, 1, 0, 1, 3));
            table.Cells[1, 0].TextString = "THÔNG TIN";
            table.Cells[1, 0].Alignment = CellAlignment.MiddleCenter;

            int col = 4;
            foreach (var mat in materials)
            {
                table.MergeCells(CellRange.Create(table, 1, col, 1, col + 1));
                // Rút gọn tên material
                string shortName = mat.Length > 20 ? mat.Substring(0, 18) + "..." : mat;
                table.Cells[1, col].TextString = shortName;
                table.Cells[1, col].Alignment = CellAlignment.MiddleCenter;
                table.Cells[1, col].TextHeight = 2.5;
                col += 2;
            }

            // Hàng 2: Chi tiết
            table.Cells[2, 0].TextString = "STT";
            table.Cells[2, 1].TextString = "TÊN";
            table.Cells[2, 2].TextString = "LÝ TRÌNH";
            table.Cells[2, 3].TextString = "K.CÁCH";

            col = 4;
            foreach (var mat in materials)
            {
                table.Cells[2, col].TextString = "AREA (m²)";
                table.Cells[2, col + 1].TextString = "KL (m³)";
                table.Cells[2, col].Alignment = CellAlignment.MiddleCenter;
                table.Cells[2, col + 1].Alignment = CellAlignment.MiddleCenter;
                col += 2;
            }

            for (int c = 0; c < numCols; c++)
            {
                table.Cells[2, c].TextHeight = 2.5;
                table.Cells[2, c].Alignment = CellAlignment.MiddleCenter;
            }

            // Dữ liệu
            int row = 3;
            int stt = 1;
            Dictionary<string, double> totalVols = materials.ToDictionary(m => m, m => 0.0);
            Dictionary<string, double> prevAreas = materials.ToDictionary(m => m, m => 0.0);

            foreach (var cs in result.CrossSections)
            {
                table.Cells[row, 0].TextString = stt++.ToString();
                table.Cells[row, 1].TextString = cs.SampleLineName;
                table.Cells[row, 2].TextString = cs.StationFormatted;
                table.Cells[row, 3].TextString = Math.Round(cs.SpacingPrev, 2).ToString();

                col = 4;
                foreach (var mat in materials)
                {
                    double area = 0;
                    if (cs.MaterialSections.ContainsKey(mat))
                        area = cs.MaterialSections[mat].Area;

                    // Tính khối lượng
                    double avgArea = (prevAreas[mat] + area) / 2;
                    double volume = avgArea * cs.SpacingPrev;
                    if (row == 3) volume = 0; // Trắc ngang đầu tiên

                    table.Cells[row, col].TextString = Math.Round(area, 3).ToString("F3");
                    table.Cells[row, col + 1].TextString = Math.Round(volume, 3).ToString("F3");

                    totalVols[mat] += volume;
                    prevAreas[mat] = area;

                    table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                    table.Cells[row, col + 1].Alignment = CellAlignment.MiddleRight;
                    table.Cells[row, col].TextHeight = 2.5;
                    table.Cells[row, col + 1].TextHeight = 2.5;

                    col += 2;
                }

                row++;
            }

            // Hàng tổng cộng
            table.MergeCells(CellRange.Create(table, row, 0, row, 3));
            table.Cells[row, 0].TextString = "TỔNG CỘNG";
            table.Cells[row, 0].Alignment = CellAlignment.MiddleCenter;
            table.Cells[row, 0].TextHeight = 3.0;

            col = 4;
            foreach (var mat in materials)
            {
                table.Cells[row, col].TextString = "";
                table.Cells[row, col + 1].TextString = Math.Round(totalVols[mat], 3).ToString("F3");
                table.Cells[row, col + 1].Alignment = CellAlignment.MiddleRight;
                table.Cells[row, col + 1].TextHeight = 3.0;
                col += 2;
            }

            btr.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
        }

        /// <summary>
        /// Lấy chữ cái cột Excel từ số cột
        /// </summary>
        private static string GetColumnLetter(int columnNumber)
        {
            string result = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                result = (char)('A' + modulo) + result;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return result;
        }

        /// <summary>
        /// Làm sạch tên sheet Excel
        /// </summary>
        private static string SanitizeSheetName(string name)
        {
            char[] invalidChars = { ':', '\\', '/', '?', '*', '[', ']' };
            string result = name;
            foreach (char c in invalidChars)
            {
                result = result.Replace(c, '_');
            }
            if (result.Length > 31)
            {
                result = result.Substring(0, 31);
            }
            return result;
        }

        #endregion
    }
}
