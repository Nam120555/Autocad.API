// CurveDesignStandards.cs - Thiết lập thông số đường cong theo tiêu chuẩn VN
// TCVN 4054:2005 - Đường ô tô (ngoài đô thị)
// TCVN 13592:2022 - Đường đô thị

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(Civil3DCsharp.CurveDesignStandardsCommands))]

namespace Civil3DCsharp
{
    #region Data Classes

    /// <summary>
    /// Thông số đường cong theo tiêu chuẩn thiết kế
    /// </summary>
    public class CurveDesignParameters
    {
        public int DesignSpeed { get; set; }           // Vận tốc thiết kế (km/h)
        public double MinRadiusLimit { get; set; }      // Bán kính tối thiểu giới hạn (m)
        public double MinRadiusNormal { get; set; }     // Bán kính tối thiểu thông thường (m)
        public double MinRadiusNoSuper { get; set; }    // Bán kính tối thiểu không siêu cao (m)
        public double MaxSuperelevation { get; set; }   // Siêu cao tối đa (%)
        public double MinTransitionLength { get; set; } // Chiều dài đường cong chuyển tiếp tối thiểu (m)
        public double MinCurveLength { get; set; }      // Chiều dài đường cong tối thiểu (m)
    }

    /// <summary>
    /// Cấp đường theo tiêu chuẩn
    /// </summary>
    public class RoadClass
    {
        public string Code { get; set; } = "";
        public string Name { get; set; } = "";
        public int[] DesignSpeeds { get; set; } = Array.Empty<int>();
        public string Description { get; set; } = "";
    }

    #endregion

    #region Standards Data

    /// <summary>
    /// Dữ liệu tiêu chuẩn TCVN 4054:2005 - Đường ô tô (ngoài đô thị)
    /// </summary>
    public static class TCVN4054_2005
    {
        public static string Name => "TCVN 4054:2005";
        public static string Title => "Đường ô tô - Yêu cầu thiết kế";

        public static readonly RoadClass[] RoadClasses =
        [
            new RoadClass { Code = "I", Name = "Cấp I", DesignSpeeds = [120, 100], Description = "Đường cao tốc, đường quốc lộ chính" },
            new RoadClass { Code = "II", Name = "Cấp II", DesignSpeeds = [80, 60], Description = "Đường quốc lộ, tỉnh lộ chính" },
            new RoadClass { Code = "III", Name = "Cấp III", DesignSpeeds = [60, 40], Description = "Đường tỉnh lộ, huyện lộ" },
            new RoadClass { Code = "IV", Name = "Cấp IV", DesignSpeeds = [40, 30], Description = "Đường huyện lộ, đường xã" },
            new RoadClass { Code = "V", Name = "Cấp V", DesignSpeeds = [30, 20], Description = "Đường xã, đường thôn" },
            new RoadClass { Code = "VI", Name = "Cấp VI", DesignSpeeds = [20], Description = "Đường đặc biệt khó khăn" }
        ];

        /// <summary>
        /// Bảng 11 - Bán kính đường cong nằm tối thiểu (TCVN 4054:2005)
        /// </summary>
        public static readonly Dictionary<int, CurveDesignParameters> CurveParameters = new()
        {
            // Vận tốc 120 km/h
            [120] = new CurveDesignParameters
            {
                DesignSpeed = 120,
                MinRadiusLimit = 650,
                MinRadiusNormal = 1000,
                MinRadiusNoSuper = 5500,
                MaxSuperelevation = 8,
                MinTransitionLength = 100,
                MinCurveLength = 170
            },
            // Vận tốc 100 km/h
            [100] = new CurveDesignParameters
            {
                DesignSpeed = 100,
                MinRadiusLimit = 400,
                MinRadiusNormal = 700,
                MinRadiusNoSuper = 4000,
                MaxSuperelevation = 8,
                MinTransitionLength = 85,
                MinCurveLength = 140
            },
            // Vận tốc 80 km/h
            [80] = new CurveDesignParameters
            {
                DesignSpeed = 80,
                MinRadiusLimit = 250,
                MinRadiusNormal = 400,
                MinRadiusNoSuper = 2500,
                MaxSuperelevation = 8,
                MinTransitionLength = 70,
                MinCurveLength = 110
            },
            // Vận tốc 60 km/h
            [60] = new CurveDesignParameters
            {
                DesignSpeed = 60,
                MinRadiusLimit = 125,
                MinRadiusNormal = 250,
                MinRadiusNoSuper = 1500,
                MaxSuperelevation = 8,
                MinTransitionLength = 50,
                MinCurveLength = 85
            },
            // Vận tốc 40 km/h
            [40] = new CurveDesignParameters
            {
                DesignSpeed = 40,
                MinRadiusLimit = 60,
                MinRadiusNormal = 125,
                MinRadiusNoSuper = 600,
                MaxSuperelevation = 8,
                MinTransitionLength = 35,
                MinCurveLength = 55
            },
            // Vận tốc 30 km/h
            [30] = new CurveDesignParameters
            {
                DesignSpeed = 30,
                MinRadiusLimit = 30,
                MinRadiusNormal = 60,
                MinRadiusNoSuper = 350,
                MaxSuperelevation = 8,
                MinTransitionLength = 25,
                MinCurveLength = 40
            },
            // Vận tốc 20 km/h
            [20] = new CurveDesignParameters
            {
                DesignSpeed = 20,
                MinRadiusLimit = 15,
                MinRadiusNormal = 50,
                MinRadiusNoSuper = 250,
                MaxSuperelevation = 6,
                MinTransitionLength = 20,
                MinCurveLength = 30
            }
        };

        /// <summary>
        /// Lấy thông số đường cong theo vận tốc thiết kế
        /// </summary>
        public static CurveDesignParameters? GetParameters(int designSpeed)
        {
            return CurveParameters.TryGetValue(designSpeed, out var param) ? param : null;
        }
    }

    /// <summary>
    /// Dữ liệu tiêu chuẩn TCVN 13592:2022 - Đường đô thị
    /// </summary>
    public static class TCVN13592_2022
    {
        public static string Name => "TCVN 13592:2022";
        public static string Title => "Đường đô thị - Yêu cầu thiết kế";

        public static readonly RoadClass[] RoadClasses =
        [
            new RoadClass { Code = "CT", Name = "Cao tốc đô thị", DesignSpeeds = [100, 80, 60], Description = "Đường cao tốc trong đô thị" },
            new RoadClass { Code = "TL", Name = "Trục chính đô thị", DesignSpeeds = [80, 60], Description = "Đường trục chính" },
            new RoadClass { Code = "CL", Name = "Chính đô thị", DesignSpeeds = [60, 50], Description = "Đường chính đô thị" },
            new RoadClass { Code = "LC", Name = "Liên khu vực", DesignSpeeds = [50, 40], Description = "Đường liên khu vực" },
            new RoadClass { Code = "KV", Name = "Khu vực", DesignSpeeds = [40, 30], Description = "Đường khu vực" },
            new RoadClass { Code = "NB", Name = "Nội bộ", DesignSpeeds = [30, 20], Description = "Đường nội bộ, ngõ hẻm" },
            new RoadClass { Code = "DB", Name = "Đi bộ", DesignSpeeds = [15], Description = "Đường đi bộ" }
        ];

        /// <summary>
        /// Bảng 18 - Chỉ tiêu kỹ thuật đường cong nằm (TCVN 13592:2022)
        /// </summary>
        public static readonly Dictionary<int, CurveDesignParameters> CurveParameters = new()
        {
            // Vận tốc 100 km/h
            [100] = new CurveDesignParameters
            {
                DesignSpeed = 100,
                MinRadiusLimit = 400,
                MinRadiusNormal = 600,
                MinRadiusNoSuper = 3000,
                MaxSuperelevation = 6,
                MinTransitionLength = 80,
                MinCurveLength = 120
            },
            // Vận tốc 80 km/h
            [80] = new CurveDesignParameters
            {
                DesignSpeed = 80,
                MinRadiusLimit = 250,
                MinRadiusNormal = 400,
                MinRadiusNoSuper = 2000,
                MaxSuperelevation = 6,
                MinTransitionLength = 65,
                MinCurveLength = 100
            },
            // Vận tốc 60 km/h
            [60] = new CurveDesignParameters
            {
                DesignSpeed = 60,
                MinRadiusLimit = 130,
                MinRadiusNormal = 200,
                MinRadiusNoSuper = 1200,
                MaxSuperelevation = 6,
                MinTransitionLength = 50,
                MinCurveLength = 70
            },
            // Vận tốc 50 km/h
            [50] = new CurveDesignParameters
            {
                DesignSpeed = 50,
                MinRadiusLimit = 80,
                MinRadiusNormal = 130,
                MinRadiusNoSuper = 800,
                MaxSuperelevation = 6,
                MinTransitionLength = 40,
                MinCurveLength = 60
            },
            // Vận tốc 40 km/h
            [40] = new CurveDesignParameters
            {
                DesignSpeed = 40,
                MinRadiusLimit = 50,
                MinRadiusNormal = 80,
                MinRadiusNoSuper = 500,
                MaxSuperelevation = 6,
                MinTransitionLength = 30,
                MinCurveLength = 50
            },
            // Vận tốc 30 km/h
            [30] = new CurveDesignParameters
            {
                DesignSpeed = 30,
                MinRadiusLimit = 25,
                MinRadiusNormal = 50,
                MinRadiusNoSuper = 300,
                MaxSuperelevation = 4,
                MinTransitionLength = 25,
                MinCurveLength = 35
            },
            // Vận tốc 20 km/h
            [20] = new CurveDesignParameters
            {
                DesignSpeed = 20,
                MinRadiusLimit = 15,
                MinRadiusNormal = 30,
                MinRadiusNoSuper = 200,
                MaxSuperelevation = 4,
                MinTransitionLength = 20,
                MinCurveLength = 25
            },
            // Vận tốc 15 km/h (đường đi bộ)
            [15] = new CurveDesignParameters
            {
                DesignSpeed = 15,
                MinRadiusLimit = 3,
                MinRadiusNormal = 15,
                MinRadiusNoSuper = 100,
                MaxSuperelevation = 2,
                MinTransitionLength = 0,
                MinCurveLength = 15
            }
        };

        /// <summary>
        /// Lấy thông số đường cong theo vận tốc thiết kế
        /// </summary>
        public static CurveDesignParameters? GetParameters(int designSpeed)
        {
            return CurveParameters.TryGetValue(designSpeed, out var param) ? param : null;
        }
    }

    #endregion

    #region Commands

    /// <summary>
    /// Các lệnh thiết lập thông số đường cong
    /// </summary>
    public class CurveDesignStandardsCommands
    {
        // ══════════════════════════════════════════════════════════════
        // HIỂN THỊ THÔNG SỐ ĐƯỜNG CONG THEO TIÊU CHUẨN
        // ══════════════════════════════════════════════════════════════

        /// <summary>
        /// Hiển thị bảng thông số đường cong theo tiêu chuẩn TCVN 4054:2005 (ngoài đô thị)
        /// </summary>
        [CommandMethod("CTC_ThongSoDuongCong_4054")]
        public static void CTC_ThongSoDuongCong_4054()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════");
            ed.WriteMessage($"\n  {TCVN4054_2005.Name} - {TCVN4054_2005.Title}");
            ed.WriteMessage("\n  BẢNG 11 - BÁN KÍNH ĐƯỜNG CONG NẰM TỐI THIỂU");
            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════");
            ed.WriteMessage("\n  Vtk     R_giới hạn   R_thường    R_không SC   isc_max   Lct_min   Lđc_min");
            ed.WriteMessage("\n  (km/h)     (m)          (m)          (m)        (%)       (m)       (m)");
            ed.WriteMessage("\n───────────────────────────────────────────────────────────────────────────────");

            foreach (var param in TCVN4054_2005.CurveParameters.Values.OrderByDescending(p => p.DesignSpeed))
            {
                ed.WriteMessage($"\n  {param.DesignSpeed,4}     {param.MinRadiusLimit,6:F0}       {param.MinRadiusNormal,6:F0}       {param.MinRadiusNoSuper,6:F0}       {param.MaxSuperelevation,3:F0}      {param.MinTransitionLength,4:F0}      {param.MinCurveLength,4:F0}");
            }

            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════");
            ed.WriteMessage("\n  Ghi chú:");
            ed.WriteMessage("\n  - R_giới hạn: Bán kính tối thiểu giới hạn (chỉ dùng khi khó khăn)");
            ed.WriteMessage("\n  - R_thường: Bán kính tối thiểu thông thường (khuyến khích)");
            ed.WriteMessage("\n  - R_không SC: Bán kính tối thiểu không bố trí siêu cao");
            ed.WriteMessage("\n  - isc_max: Siêu cao tối đa");
            ed.WriteMessage("\n  - Lct_min: Chiều dài đường cong chuyển tiếp tối thiểu");
            ed.WriteMessage("\n  - Lđc_min: Chiều dài đường cong tối thiểu");
            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════\n");
        }

        /// <summary>
        /// Hiển thị bảng thông số đường cong theo tiêu chuẩn TCVN 13592:2022 (đô thị)
        /// </summary>
        [CommandMethod("CTC_ThongSoDuongCong_13592")]
        public static void CTC_ThongSoDuongCong_13592()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════");
            ed.WriteMessage($"\n  {TCVN13592_2022.Name} - {TCVN13592_2022.Title}");
            ed.WriteMessage("\n  BẢNG 18 - CHỈ TIÊU KỸ THUẬT ĐƯỜNG CONG NẰM");
            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════");
            ed.WriteMessage("\n  Vtk     R_giới hạn   R_thường    R_không SC   isc_max   Lct_min   Lđc_min");
            ed.WriteMessage("\n  (km/h)     (m)          (m)          (m)        (%)       (m)       (m)");
            ed.WriteMessage("\n───────────────────────────────────────────────────────────────────────────────");

            foreach (var param in TCVN13592_2022.CurveParameters.Values.OrderByDescending(p => p.DesignSpeed))
            {
                ed.WriteMessage($"\n  {param.DesignSpeed,4}     {param.MinRadiusLimit,6:F0}       {param.MinRadiusNormal,6:F0}       {param.MinRadiusNoSuper,6:F0}       {param.MaxSuperelevation,3:F0}      {param.MinTransitionLength,4:F0}      {param.MinCurveLength,4:F0}");
            }

            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════");
            ed.WriteMessage("\n  Cấp đường đô thị:");
            foreach (var rc in TCVN13592_2022.RoadClasses)
            {
                ed.WriteMessage($"\n  - {rc.Code}: {rc.Name} ({string.Join(", ", rc.DesignSpeeds)} km/h) - {rc.Description}");
            }
            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════\n");
        }

        // ══════════════════════════════════════════════════════════════
        // KIỂM TRA ĐƯỜNG CONG HIỆN TẠI THEO TIÊU CHUẨN
        // ══════════════════════════════════════════════════════════════

        /// <summary>
        /// Kiểm tra alignment theo tiêu chuẩn TCVN 4054:2005
        /// </summary>
        [CommandMethod("CTC_KiemTraDuongCong_4054")]
        public static void CTC_KiemTraDuongCong_4054()
        {
            KiemTraDuongCongTheoTieuChuan(true);
        }

        /// <summary>
        /// Kiểm tra alignment theo tiêu chuẩn TCVN 13592:2022
        /// </summary>
        [CommandMethod("CTC_KiemTraDuongCong_13592")]
        public static void CTC_KiemTraDuongCong_13592()
        {
            KiemTraDuongCongTheoTieuChuan(false);
        }

        private static void KiemTraDuongCongTheoTieuChuan(bool isTCVN4054)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;

            // Chọn vận tốc thiết kế
            var speeds = isTCVN4054 
                ? TCVN4054_2005.CurveParameters.Keys.OrderByDescending(x => x).ToArray()
                : TCVN13592_2022.CurveParameters.Keys.OrderByDescending(x => x).ToArray();

            ed.WriteMessage($"\n▸ Vận tốc thiết kế khả dụng: {string.Join(", ", speeds)} km/h");

            var pio = new PromptIntegerOptions("\n▸ Nhập vận tốc thiết kế (km/h): ");
            pio.DefaultValue = 60;
            pio.UseDefaultValue = true;
            var pir = ed.GetInteger(pio);
            if (pir.Status != PromptStatus.OK) return;

            int designSpeed = pir.Value;

            var parameters = isTCVN4054
                ? TCVN4054_2005.GetParameters(designSpeed)
                : TCVN13592_2022.GetParameters(designSpeed);

            if (parameters == null)
            {
                ed.WriteMessage($"\n⊘ Không tìm thấy thông số cho vận tốc {designSpeed} km/h trong tiêu chuẩn này.");
                return;
            }

            // Chọn Alignment
            var peo = new PromptEntityOptions("\n⊙ Chọn Alignment để kiểm tra: ");
            peo.SetRejectMessage("\n⊘ Đây không phải Alignment!");
            peo.AddAllowedClass(typeof(Alignment), true);

            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            var standardName = isTCVN4054 ? TCVN4054_2005.Name : TCVN13592_2022.Name;

            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var alignment = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Alignment;
                if (alignment == null) return;

                ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════");
                ed.WriteMessage($"\n  KIỂM TRA ĐƯỜNG CONG - {standardName}");
                ed.WriteMessage($"\n  Alignment: {alignment.Name}");
                ed.WriteMessage($"\n  Vận tốc thiết kế: {designSpeed} km/h");
                ed.WriteMessage($"\n  Bán kính tối thiểu thông thường: {parameters.MinRadiusNormal} m");
                ed.WriteMessage($"\n  Bán kính tối thiểu giới hạn: {parameters.MinRadiusLimit} m");
                ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════");

                int curveCount = 0;
                int errorCount = 0;
                int warningCount = 0;

                foreach (AlignmentEntity entity in alignment.Entities)
                {
                    if (entity is AlignmentArc arc)
                    {
                        curveCount++;
                        double radius = Math.Abs(arc.Radius);
                        double length = arc.Length;
                        string status = "✓ ĐẠT";

                        if (radius < parameters.MinRadiusLimit)
                        {
                            status = "✗ KHÔNG ĐẠT (< R giới hạn)";
                            errorCount++;
                        }
                        else if (radius < parameters.MinRadiusNormal)
                        {
                            status = "⚠ CẢNH BÁO (< R thường)";
                            warningCount++;
                        }

                        if (length < parameters.MinCurveLength)
                        {
                            status += $" [Lđc < {parameters.MinCurveLength}m]";
                            if (!status.Contains("KHÔNG ĐẠT"))
                            {
                                warningCount++;
                            }
                        }

                        ed.WriteMessage($"\n  Đường cong #{curveCount}: R = {radius:F2}m, L = {length:F2}m → {status}");
                    }
                    else if (entity is AlignmentSCS scs) // Spiral-Curve-Spiral
                    {
                        curveCount++;
                        double radius = Math.Abs(scs.Arc.Radius);
                        double spiralIn = scs.SpiralIn.Length;
                        double spiralOut = scs.SpiralOut.Length;
                        string status = "✓ ĐẠT";

                        if (radius < parameters.MinRadiusLimit)
                        {
                            status = "✗ KHÔNG ĐẠT (< R giới hạn)";
                            errorCount++;
                        }
                        else if (radius < parameters.MinRadiusNormal)
                        {
                            status = "⚠ CẢNH BÁO (< R thường)";
                            warningCount++;
                        }

                        if (spiralIn < parameters.MinTransitionLength || spiralOut < parameters.MinTransitionLength)
                        {
                            status += $" [Lct < {parameters.MinTransitionLength}m]";
                            if (!status.Contains("KHÔNG ĐẠT"))
                            {
                                warningCount++;
                            }
                        }

                        ed.WriteMessage($"\n  SCS #{curveCount}: R = {radius:F2}m, Lct_in = {spiralIn:F2}m, Lct_out = {spiralOut:F2}m → {status}");
                    }
                }

                ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════");
                ed.WriteMessage($"\n  KẾT QUẢ: {curveCount} đường cong");
                ed.WriteMessage($"\n  - Không đạt: {errorCount}");
                ed.WriteMessage($"\n  - Cảnh báo: {warningCount}");
                ed.WriteMessage($"\n  - Đạt: {curveCount - errorCount - warningCount}");
                ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════\n");

                tr.Commit();
            }
        }

        // ══════════════════════════════════════════════════════════════
        // TRA CỨU THÔNG SỐ ĐƯỜNG CONG
        // ══════════════════════════════════════════════════════════════

        /// <summary>
        /// Tra cứu nhanh thông số đường cong theo vận tốc
        /// </summary>
        [CommandMethod("CTC_TraCuuDuongCong")]
        public static void CTC_TraCuuDuongCong()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            // Chọn tiêu chuẩn
            var pko = new PromptKeywordOptions("\n▸ Chọn tiêu chuẩn [4054(ngoàiĐôThị)/13592(đôThị)]: ", "4054 13592");
            pko.AllowNone = false;
            var pkr = ed.GetKeywords(pko);
            if (pkr.Status != PromptStatus.OK) return;

            bool isTCVN4054 = pkr.StringResult == "4054";

            // Nhập vận tốc
            var pio = new PromptIntegerOptions("\n▸ Nhập vận tốc thiết kế (km/h): ");
            pio.DefaultValue = 60;
            pio.UseDefaultValue = true;
            var pir = ed.GetInteger(pio);
            if (pir.Status != PromptStatus.OK) return;

            int designSpeed = pir.Value;

            var parameters = isTCVN4054
                ? TCVN4054_2005.GetParameters(designSpeed)
                : TCVN13592_2022.GetParameters(designSpeed);

            if (parameters == null)
            {
                ed.WriteMessage($"\n⊘ Không tìm thấy thông số cho vận tốc {designSpeed} km/h.");
                ed.WriteMessage($"\n▸ Các vận tốc khả dụng: ");
                var speeds = isTCVN4054
                    ? TCVN4054_2005.CurveParameters.Keys
                    : TCVN13592_2022.CurveParameters.Keys;
                ed.WriteMessage(string.Join(", ", speeds.OrderByDescending(x => x)) + " km/h");
                return;
            }

            string standardName = isTCVN4054 ? TCVN4054_2005.Name : TCVN13592_2022.Name;

            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════");
            ed.WriteMessage($"\n  THÔNG SỐ ĐƯỜNG CONG - {standardName}");
            ed.WriteMessage($"\n  Vận tốc thiết kế: {designSpeed} km/h");
            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════");
            ed.WriteMessage($"\n  ▸ Bán kính tối thiểu giới hạn:     {parameters.MinRadiusLimit,6:F0} m");
            ed.WriteMessage($"\n  ▸ Bán kính tối thiểu thông thường: {parameters.MinRadiusNormal,6:F0} m");
            ed.WriteMessage($"\n  ▸ Bán kính không siêu cao:         {parameters.MinRadiusNoSuper,6:F0} m");
            ed.WriteMessage($"\n  ▸ Siêu cao tối đa:                 {parameters.MaxSuperelevation,6:F0} %");
            ed.WriteMessage($"\n  ▸ Chiều dài chuyển tiếp tối thiểu: {parameters.MinTransitionLength,6:F0} m");
            ed.WriteMessage($"\n  ▸ Chiều dài đường cong tối thiểu:  {parameters.MinCurveLength,6:F0} m");
            ed.WriteMessage("\n═══════════════════════════════════════════════════════════════════════════════\n");

            // Hiển thị MessageBox
            string msg = $"THÔNG SỐ ĐƯỜNG CONG - {standardName}\n" +
                        $"Vận tốc thiết kế: {designSpeed} km/h\n\n" +
                        $"• R tối thiểu giới hạn: {parameters.MinRadiusLimit} m\n" +
                        $"• R tối thiểu thường: {parameters.MinRadiusNormal} m\n" +
                        $"• R không siêu cao: {parameters.MinRadiusNoSuper} m\n" +
                        $"• Siêu cao max: {parameters.MaxSuperelevation} %\n" +
                        $"• Lct min: {parameters.MinTransitionLength} m\n" +
                        $"• Lđc min: {parameters.MinCurveLength} m";

            MessageBox.Show(msg, $"Tra cứu đường cong - V={designSpeed}km/h", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ══════════════════════════════════════════════════════════════
        // MỞ FORM THIẾT LẬP ĐƯỜNG CONG
        // ══════════════════════════════════════════════════════════════

        /// <summary>
        /// Mở form thiết lập thông số đường cong
        /// </summary>
        [CommandMethod("CTC_ThietLapDuongCong")]
        public static void CTC_ThietLapDuongCong()
        {
            var form = new CurveDesignForm();
            AcadApp.ShowModelessDialog(form);
        }
    }

    #endregion

    #region Form

    /// <summary>
    /// Form thiết lập thông số đường cong
    /// </summary>
    public class CurveDesignForm : Form
    {
        private ComboBox cboStandard = null!;
        private ComboBox cboSpeed = null!;
        private DataGridView dgvParameters = null!;
        private System.Windows.Forms.Label lblInfo = null!;
        private Button btnApply = null!;
        private Button btnCheck = null!;
        private Button btnClose = null!;

        public CurveDesignForm()
        {
            InitializeComponent();
            LoadData();
        }

        private void InitializeComponent()
        {
            this.Text = "Thiết lập thông số đường cong - Tiêu chuẩn VN";
            this.Size = new System.Drawing.Size(650, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // Tiêu chuẩn
            var lblStandard = new System.Windows.Forms.Label { Text = "Tiêu chuẩn:", Location = new System.Drawing.Point(15, 20), AutoSize = true };
            cboStandard = new ComboBox
            {
                Location = new System.Drawing.Point(120, 17),
                Width = 250,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboStandard.Items.Add($"{TCVN4054_2005.Name} - Đường ngoài đô thị");
            cboStandard.Items.Add($"{TCVN13592_2022.Name} - Đường đô thị");
            cboStandard.SelectedIndex = 0;
            cboStandard.SelectedIndexChanged += (s, e) => LoadData();

            // Vận tốc
            var lblSpeed = new System.Windows.Forms.Label { Text = "Vận tốc TK:", Location = new System.Drawing.Point(15, 55), AutoSize = true };
            cboSpeed = new ComboBox
            {
                Location = new System.Drawing.Point(120, 52),
                Width = 120,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboSpeed.SelectedIndexChanged += (s, e) => UpdateDisplay();

            var lblKmh = new System.Windows.Forms.Label { Text = "km/h", Location = new System.Drawing.Point(245, 55), AutoSize = true };

            // Info label
            lblInfo = new System.Windows.Forms.Label
            {
                Location = new System.Drawing.Point(15, 90),
                Size = new System.Drawing.Size(600, 60),
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold)
            };

            // DataGridView
            dgvParameters = new DataGridView
            {
                Location = new System.Drawing.Point(15, 155),
                Size = new System.Drawing.Size(600, 250),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };

            dgvParameters.Columns.Add("Param", "Thông số");
            dgvParameters.Columns.Add("Value", "Giá trị");
            dgvParameters.Columns.Add("Unit", "Đơn vị");
            dgvParameters.Columns.Add("Note", "Ghi chú");

            // Buttons
            btnApply = new Button { Text = "Áp dụng", Location = new System.Drawing.Point(370, 420), Width = 80 };
            btnCheck = new Button { Text = "Kiểm tra", Location = new System.Drawing.Point(455, 420), Width = 80 };
            btnClose = new Button { Text = "Đóng", Location = new System.Drawing.Point(540, 420), Width = 80 };

            btnApply.Click += BtnApply_Click;
            btnCheck.Click += BtnCheck_Click;
            btnClose.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[] {
                lblStandard, cboStandard, lblSpeed, cboSpeed, lblKmh,
                lblInfo, dgvParameters, btnApply, btnCheck, btnClose
            });
        }

        private void LoadData()
        {
            cboSpeed.Items.Clear();

            if (cboStandard.SelectedIndex == 0)
            {
                foreach (var speed in TCVN4054_2005.CurveParameters.Keys.OrderByDescending(x => x))
                {
                    cboSpeed.Items.Add(speed);
                }
            }
            else
            {
                foreach (var speed in TCVN13592_2022.CurveParameters.Keys.OrderByDescending(x => x))
                {
                    cboSpeed.Items.Add(speed);
                }
            }

            if (cboSpeed.Items.Count > 0)
            {
                cboSpeed.SelectedIndex = 2; // Mặc định 60 hoặc 80 km/h
            }
        }

        private void UpdateDisplay()
        {
            if (cboSpeed.SelectedItem == null) return;

            int speed = (int)cboSpeed.SelectedItem;
            CurveDesignParameters? parameters = null;
            string standardName = "";

            if (cboStandard.SelectedIndex == 0)
            {
                parameters = TCVN4054_2005.GetParameters(speed);
                standardName = TCVN4054_2005.Name;
            }
            else
            {
                parameters = TCVN13592_2022.GetParameters(speed);
                standardName = TCVN13592_2022.Name;
            }

            if (parameters == null) return;

            lblInfo.Text = $"Tiêu chuẩn: {standardName}\n" +
                          $"Vận tốc thiết kế: {speed} km/h\n" +
                          $"Bán kính tối thiểu thông thường: {parameters.MinRadiusNormal} m";

            dgvParameters.Rows.Clear();
            dgvParameters.Rows.Add("Bán kính tối thiểu giới hạn", parameters.MinRadiusLimit, "m", "Chỉ dùng khi khó khăn");
            dgvParameters.Rows.Add("Bán kính tối thiểu thông thường", parameters.MinRadiusNormal, "m", "Khuyến khích sử dụng");
            dgvParameters.Rows.Add("Bán kính không siêu cao", parameters.MinRadiusNoSuper, "m", "Không cần bố trí siêu cao");
            dgvParameters.Rows.Add("Siêu cao tối đa", parameters.MaxSuperelevation, "%", "");
            dgvParameters.Rows.Add("Chiều dài chuyển tiếp tối thiểu", parameters.MinTransitionLength, "m", "Đường cong Clothoid");
            dgvParameters.Rows.Add("Chiều dài đường cong tối thiểu", parameters.MinCurveLength, "m", "");
        }

        private void BtnApply_Click(object? sender, EventArgs e)
        {
            MessageBox.Show(
                "Thông số đường cong đã được hiển thị.\n\n" +
                "Để áp dụng vào thiết kế:\n" +
                "1. Ghi nhớ các giá trị R min, Lct min\n" +
                "2. Sử dụng trong Alignment Design\n" +
                "3. Thiết lập Superelevation View",
                "Thông báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }

        private void BtnCheck_Click(object? sender, EventArgs e)
        {
            this.Hide();
            
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (cboStandard.SelectedIndex == 0)
            {
                doc?.SendStringToExecute("CTC_KiemTraDuongCong_4054 ", true, false, false);
            }
            else
            {
                doc?.SendStringToExecute("CTC_KiemTraDuongCong_13592 ", true, false, false);
            }
            
            this.Show();
        }
    }

    #endregion
}
