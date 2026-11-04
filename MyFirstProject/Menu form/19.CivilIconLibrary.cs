// 19.CivilIconLibrary.cs - Bộ icon vector chuyên ngành cầu đường cho Ribbon / Taskbar
// Được sinh từ bản thiết kế lưới 32x32. Mỗi icon gồm nhiều lớp (path WPF + màu theo vai trò).
// Quy ước màu theo bản vẽ đường – cầu:
//   - Đường đỏ (design)   : đường thiết kế, mái taluy thiết kế
//   - Nâu (ground)         : đường mặt đất tự nhiên, đồng mức
//   - Xanh dương (align)   : tim tuyến, khung mặt cắt
//   - Đỏ / xanh lá nhạt    : diện tích đào / đắp
//   - Cam (accent)         : đỉnh P, PVI, cọc khống chế
// Tự đổi bảng màu theo theme AutoCAD (biến COLORTHEME: 0 = tối, 1 = sáng).

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using WColor = System.Windows.Media.Color;

namespace MyFirstProject
{
    public static class CivilIconLibrary
    {
        private sealed class Layer
        {
            public string D = "";
            public string? Fill;
            public string? Stroke;
            public double Sw;
            public double[]? Dash;
            public double Fo = 1;
            public double So = 1;
        }

        private static Layer Ly(string d, string? fill, string? stroke, double sw, double[]? dash = null, double fo = 1, double so = 1)
            => new Layer { D = d, Fill = fill, Stroke = stroke, Sw = sw, Dash = dash, Fo = fo, So = so };

        /// <summary>Danh sách tên icon hợp lệ (dùng cho combo chọn icon)</summary>
        public static IReadOnlyList<string> Names => _order;

        private static readonly string[] _order =
        {
            "SURFACE", "ALIGN", "STATION", "PROFILE", "SECTION", "CORRIDOR", "GRADING", "PIPE",
            "CULVERT", "BRIDGE", "POINT", "INTERSECTION", "SUPERELEVATION", "SPIRAL", "WIDENING", "VOLUME",
            "SLOPE", "BENCH", "PAVEMENT", "ADD", "DELETE", "INFO", "RENAME", "TABLE",
            "EXCEL", "CAD", "EXPORT", "SYNC", "RELOAD", "SETTINGS", "UTILITY", "ACCOUNT",
            "CALCULATE", "DRAW", "VIEW", "COPY", "MOVE", "EDIT", "MEASURE", "LAYOUT",
            "PRINT",
        };

        private static readonly Dictionary<string, Layer[]> Icons = new(StringComparer.OrdinalIgnoreCase)
        {
            ["SURFACE"] = new[]
            {
                Ly("M3 26 L9 11 L17 18 Z", "elev_lowmid", null, 0),
                Ly("M9 11 L24 6 L17 18 Z", "elev_high", null, 0),
                Ly("M17 18 L24 6 L29 16 Z", "elev_highmid", null, 0),
                Ly("M17 18 L29 16 L27 28 Z", "elev_mid", null, 0),
                Ly("M3 26 L17 18 L14 29 Z", "elev_low", null, 0),
                Ly("M14 29 L17 18 L27 28 Z", "elev_low", null, 0),
                Ly("M3 26 L9 11 L17 18 Z", null, "ink", 1.3),
                Ly("M9 11 L24 6 L17 18 Z", null, "ink", 1.3),
                Ly("M17 18 L24 6 L29 16 Z", null, "ink", 1.3),
                Ly("M17 18 L29 16 L27 28 Z", null, "ink", 1.3),
                Ly("M3 26 L17 18 L14 29 Z", null, "ink", 1.3),
                Ly("M14 29 L17 18 L27 28 Z", null, "ink", 1.3),
                Ly("M1.6 26 A1.4 1.4 0 1 0 4.4 26 A1.4 1.4 0 1 0 1.6 26 Z", "ink", null, 0),
                Ly("M7.6 11 A1.4 1.4 0 1 0 10.4 11 A1.4 1.4 0 1 0 7.6 11 Z", "ink", null, 0),
                Ly("M15.6 18 A1.4 1.4 0 1 0 18.4 18 A1.4 1.4 0 1 0 15.6 18 Z", "ink", null, 0),
                Ly("M22.6 6 A1.4 1.4 0 1 0 25.4 6 A1.4 1.4 0 1 0 22.6 6 Z", "ink", null, 0),
                Ly("M27.6 16 A1.4 1.4 0 1 0 30.4 16 A1.4 1.4 0 1 0 27.6 16 Z", "ink", null, 0),
                Ly("M25.6 28 A1.4 1.4 0 1 0 28.4 28 A1.4 1.4 0 1 0 25.6 28 Z", "ink", null, 0),
                Ly("M12.6 29 A1.4 1.4 0 1 0 15.4 29 A1.4 1.4 0 1 0 12.6 29 Z", "ink", null, 0),
            },
            ["ALIGN"] = new[]
            {
                Ly("M10.8 15.6 L16 8 L21.2 15.6", null, "ink2", 1.2, new double[] { 2, 2 }),
                Ly("M3 27 L10.8 15.6 Q16 8 21.2 15.6 L29 27", null, "align", 2.6),
                Ly("M8.4 13.9 L13.2 17.3", null, "ink", 1.6),
                Ly("M18.8 17.3 L23.6 13.9", null, "ink", 1.6),
                Ly("M16 4.5 L19.2 9.8 L12.8 9.8 Z", "accent", "accent", 0.8),
            },
            ["STATION"] = new[]
            {
                Ly("M2 21 L30 12", null, "align", 2.6),
                Ly("M3.88 11.47 L9.08 27.65", null, "ink", 1.5),
                Ly("M10.04 9.49 L15.24 25.67", null, "accent", 2),
                Ly("M16.76 7.33 L21.96 23.51", null, "ink", 1.5),
                Ly("M22.92 5.35 L28.12 21.53", null, "ink", 1.5),
                Ly("M10.74 17.58 A1.9 1.9 0 1 0 14.54 17.58 A1.9 1.9 0 1 0 10.74 17.58 Z", "accent", null, 0),
            },
            ["PROFILE"] = new[]
            {
                Ly("M3 4 H29 V28 H3 Z", "panel", "ink2", 1),
                Ly("M3 22 H29", null, "ink2", 1),
                Ly("M3 25 H29", null, "ink2", 0.8, null, 1, 0.7),
                Ly("M9.5 22 V28 M16 22 V28 M22.5 22 V28", null, "ink2", 0.8, null, 1, 0.7),
                Ly("M3 15 L6.5 13.5 L9.5 16.5 L13 12.5 L17 14.5 L21 10.5 L25 14 L29 12.5", null, "ground", 1.6),
                Ly("M3 17.5 L12 12.6 Q16.5 10 21 12.4 L29 16", null, "design", 2.3),
                Ly("M12 12.6 L16.5 10 L21 12.4", null, "design", 0.9, new double[] { 1.5, 1.5 }, 1, 0.8),
                Ly("M14.9 10 A1.6 1.6 0 1 0 18.1 10 A1.6 1.6 0 1 0 14.9 10 Z", null, "accent", 1.4),
            },
            ["SECTION"] = new[]
            {
                Ly("M6 12.5 L16 11 L23.5 12.2 L21 16 L16 21 L11 21.5 L6 20.5 L1.5 22 Z", "fill", null, 0, null, 0.38),
                Ly("M23.5 12.2 L26 12.5 L28 17 L30.5 9 L25 11.5 Z", "cut", null, 0, null, 0.45),
                Ly("M1.5 22 L6 20.5 L11 21.5 L16 21 L21 16 L25 11.5 L30.5 9", null, "ground", 1.6),
                Ly("M6 12.5 L16 11 L26 12.5 L26 15 L16 13.5 L6 15 Z", "asphalt", "asphalt_edge", 0.7),
                Ly("M6 13.8 L1.5 22", null, "design", 2),
                Ly("M26 13.8 L28 17 L30.5 9", null, "design", 2),
                Ly("M16 4 V28", null, "ink2", 1, new double[] { 3, 1.5, 0.6, 1.5 }),
                Ly("M13 6.5 H19", null, "ink2", 1),
            },
            ["CORRIDOR"] = new[]
            {
                Ly("M2 29 L12.5 5 L19.5 5 L30 29 Z", "shoulder", null, 0),
                Ly("M5.5 29 L13.4 5 L18.6 5 L26.5 29 Z", "asphalt", null, 0),
                Ly("M2 29 L12.5 5 M30 29 L19.5 5", null, "ink", 1.4),
                Ly("M5.5 29 L13.4 5 M26.5 29 L18.6 5", null, "road_mark", 1),
                Ly("M16 29 V5", null, "road_mark_y", 1.3, new double[] { 2.5, 2 }),
                Ly("M3.9 24.6 H28.1 M6.7 18.2 H25.3 M9.4 12 H22.6", null, "align", 1.2, null, 1, 0.9),
            },
            ["GRADING"] = new[]
            {
                Ly("M4 12 H12 V20 H4 Z", "fill", null, 0, null, 0.45),
                Ly("M4 20 H12 V28 H4 Z", "fill", null, 0, null, 0.45),
                Ly("M12 4 H20 V12 H12 Z", "cut", null, 0, null, 0.45),
                Ly("M12 20 H20 V28 H12 Z", "fill", null, 0, null, 0.45),
                Ly("M20 4 H28 V12 H20 Z", "cut", null, 0, null, 0.45),
                Ly("M20 12 H28 V20 H20 Z", "cut", null, 0, null, 0.45),
                Ly("M4 4 H28 V28 H4 Z M12 4 V28 M20 4 V28 M4 12 H28 M4 20 H28", null, "ink2", 1.1),
                Ly("M4 18 C10 17 14 11 22 4", null, "ground", 1.6),
                Ly("M4 25 C12 24 19 19 28 9", null, "ground", 1.6),
                Ly("M2.7 4 A1.3 1.3 0 1 0 5.3 4 A1.3 1.3 0 1 0 2.7 4 Z", "ink", null, 0),
                Ly("M2.7 12 A1.3 1.3 0 1 0 5.3 12 A1.3 1.3 0 1 0 2.7 12 Z", "ink", null, 0),
                Ly("M2.7 20 A1.3 1.3 0 1 0 5.3 20 A1.3 1.3 0 1 0 2.7 20 Z", "ink", null, 0),
                Ly("M2.7 28 A1.3 1.3 0 1 0 5.3 28 A1.3 1.3 0 1 0 2.7 28 Z", "ink", null, 0),
                Ly("M10.7 4 A1.3 1.3 0 1 0 13.3 4 A1.3 1.3 0 1 0 10.7 4 Z", "ink", null, 0),
                Ly("M10.7 12 A1.3 1.3 0 1 0 13.3 12 A1.3 1.3 0 1 0 10.7 12 Z", "ink", null, 0),
                Ly("M10.7 20 A1.3 1.3 0 1 0 13.3 20 A1.3 1.3 0 1 0 10.7 20 Z", "ink", null, 0),
                Ly("M10.7 28 A1.3 1.3 0 1 0 13.3 28 A1.3 1.3 0 1 0 10.7 28 Z", "ink", null, 0),
                Ly("M18.7 4 A1.3 1.3 0 1 0 21.3 4 A1.3 1.3 0 1 0 18.7 4 Z", "ink", null, 0),
                Ly("M18.7 12 A1.3 1.3 0 1 0 21.3 12 A1.3 1.3 0 1 0 18.7 12 Z", "ink", null, 0),
                Ly("M18.7 20 A1.3 1.3 0 1 0 21.3 20 A1.3 1.3 0 1 0 18.7 20 Z", "ink", null, 0),
                Ly("M18.7 28 A1.3 1.3 0 1 0 21.3 28 A1.3 1.3 0 1 0 18.7 28 Z", "ink", null, 0),
                Ly("M26.7 4 A1.3 1.3 0 1 0 29.3 4 A1.3 1.3 0 1 0 26.7 4 Z", "ink", null, 0),
                Ly("M26.7 12 A1.3 1.3 0 1 0 29.3 12 A1.3 1.3 0 1 0 26.7 12 Z", "ink", null, 0),
                Ly("M26.7 20 A1.3 1.3 0 1 0 29.3 20 A1.3 1.3 0 1 0 26.7 20 Z", "ink", null, 0),
                Ly("M26.7 28 A1.3 1.3 0 1 0 29.3 28 A1.3 1.3 0 1 0 26.7 28 Z", "ink", null, 0),
            },
            ["PIPE"] = new[]
            {
                Ly("M2 7 H30", null, "ground", 1.6),
                Ly("M9 19.5 L23 22.5 L23 25.5 L9 22.5 Z", "water", null, 0, null, 0.9),
                Ly("M9 19.5 L23 22.5 M9 22.5 L23 25.5", null, "ink", 1.3),
                Ly("M4 7 H9 V25 H4 Z", "concrete", "ink", 1.2),
                Ly("M23 7 H28 V28 H23 Z", "concrete", "ink", 1.2),
                Ly("M13.5 17.6 L18.5 18.7 M16.6 16.9 L18.5 18.7 L16 19.7", null, "water_dark", 1.2),
            },
            ["CULVERT"] = new[]
            {
                Ly("M1.5 27 L10 9 L22 9 L30.5 27 Z", "ground", null, 0, null, 0.35),
                Ly("M10 9 H22", null, "asphalt_top", 2.2),
                Ly("M1.5 27 L10 9 M30.5 27 L22 9", null, "ground", 1.4),
                Ly("M7.5 13.5 H24.5 V27.5 H7.5 Z M10.3 16.3 H21.7 V24.8 H10.3 Z", "concrete", "ink", 1.1),
                Ly("M10.3 22 Q12.2 21 14.1 22 T18 22 T21.7 22 V24.8 H10.3 Z", "water", null, 0),
                Ly("M1 27.5 H31", null, "ink2", 1.2),
            },
            ["BRIDGE"] = new[]
            {
                Ly("M1.5 25.5 Q16 33 30.5 25.5 L30.5 30 L1.5 30 Z", "ground", null, 0, null, 0.45),
                Ly("M5 21 Q8 20 11 21 T17 21 T23 21 T29 21", null, "water", 1.4),
                Ly("M1.5 9 H30.5", null, "ink2", 1),
                Ly("M1.5 10.5 H30.5 V14.5 H1.5 Z", "concrete", "ink", 1.1),
                Ly("M1.5 14.5 L5 14.5 L3.5 27 L1.5 27 Z M30.5 14.5 L27 14.5 L28.5 27 L30.5 27 Z", "concrete", "ink", 1),
                Ly("M10.2 14.5 H12.8 V26.5 H10.2 Z", "concrete", "ink", 1),
                Ly("M19.2 14.5 H21.8 V26.5 H19.2 Z", "concrete", "ink", 1),
                Ly("M8.7 26.5 H14.3 V28.5 H8.7 Z", "concrete", "ink", 0.9),
                Ly("M17.7 26.5 H23.3 V28.5 H17.7 Z", "concrete", "ink", 0.9),
            },
            ["POINT"] = new[]
            {
                Ly("M5 19 A6 6 0 1 0 17 19 A6 6 0 1 0 5 19 Z", null, "ink", 1.5),
                Ly("M3 19 H19 M11 11 V27", null, "ink", 1.2),
                Ly("M9 19 A2 2 0 1 0 13 19 A2 2 0 1 0 9 19 Z", "accent", null, 0),
                Ly("M15.5 14.5 L19 9 H29", null, "ink2", 1),
                Ly("M20 5.5 H29 M20 12.5 H27", null, "align", 1.8),
            },
            ["INTERSECTION"] = new[]
            {
                Ly("M2 11 H7.5 A3.5 3.5 0 0 0 11 7.5 V2 H21 V7.5 A3.5 3.5 0 0 0 24.5 11 H30 V21 H24.5 A3.5 3.5 0 0 0 21 24.5 V30 H11 V24.5 A3.5 3.5 0 0 0 7.5 21 H2 Z", "asphalt", "ink", 1.3),
                Ly("M2 16 H9 M23 16 H30 M16 2 V9 M16 23 V30", null, "road_mark_y", 1.2, new double[] { 2, 1.6 }),
                Ly("M11 11 H21 V21 H11 Z", null, "road_mark", 0.8, new double[] { 1, 1.2 }),
            },
            ["SUPERELEVATION"] = new[]
            {
                Ly("M3 16 H29", null, "ink2", 1, new double[] { 2, 1.5 }),
                Ly("M8 4 V28 M13 4 V28 M18 4 V28 M22.5 4 V28", null, "ink2", 0.7, null, 1, 0.6),
                Ly("M3 19 L8 19 L13 16 L18 13 L22.5 9.5 L29 9.5", null, "align", 2.2),
                Ly("M3 19 L18 19 L22.5 22.5 L29 22.5", null, "design", 2.2),
            },
            ["SPIRAL"] = new[]
            {
                Ly("M3 27 H13", null, "ink2", 1, new double[] { 2, 1.5 }),
                Ly("M21.5 18.5 C24 14 25 10 24.5 5", null, "ink2", 1, new double[] { 2, 1.5 }),
                Ly("M3 27 C12 27 17.5 24.5 21.5 18.5 C24 14 25 10 24.5 5", null, "align", 2.6),
                Ly("M1.4 25.4 H4.6 V28.6 H1.4 Z", "green", null, 0),
                Ly("M19.6 18.5 A1.9 1.9 0 1 0 23.4 18.5 A1.9 1.9 0 1 0 19.6 18.5 Z", "accent", null, 0),
                Ly("M29 27 A12 12 0 0 0 21.5 18.5", null, "ink2", 0.9, null, 1, 0.8),
            },
            ["WIDENING"] = new[]
            {
                Ly("M2 8 H30 V25 H22 L13 19.5 H2 Z", "asphalt", null, 0),
                Ly("M2 8 H30", null, "ink", 1.4),
                Ly("M2 19.5 H13 L22 25 H30", null, "ink", 1.4),
                Ly("M2 13.7 H30", null, "road_mark_y", 1.2, new double[] { 2.2, 1.8 }),
                Ly("M13 19.5 H30", null, "road_mark", 0.9, new double[] { 1.2, 1.2 }),
                Ly("M27 19.8 V24.6 M25.4 21.3 L27 19.6 L28.6 21.3 M25.4 23.2 L27 24.9 L28.6 23.2", null, "accent", 1.2),
            },
            ["VOLUME"] = new[]
            {
                Ly("M2 16 L16 16 L9 20.5 L2 22.5 Z", "fill", null, 0, null, 0.55),
                Ly("M16 16 L22 8.5 L30 12 L30 16 Z", "cut", null, 0, null, 0.55),
                Ly("M4 18 L6 16 M4 21.8 L9.8 16 M9.6 19.5 L13.1 16", null, "fill", 0.9),
                Ly("M19 16 L22.8 12.2 M23 16 L27 12 M27 16 L30 13", null, "cut", 0.9),
                Ly("M2 22.5 L9 20.5 L16 16 L22 8.5 L30 12", null, "ground", 1.6),
                Ly("M2 16 H30", null, "design", 2.2),
                Ly("M3 27 H29 M3 25.5 V28.5 M29 25.5 V28.5", null, "ink2", 1),
            },
            ["SLOPE"] = new[]
            {
                Ly("M2 8 H30", null, "ink", 2),
                Ly("M4 8 V24 M7 8 V15.5 M10 8 V24 M13 8 V15.5 M16 8 V24 M19 8 V15.5 M22 8 V24 M25 8 V15.5 M28 8 V24", null, "ink", 1.2),
                Ly("M2 24.5 Q9 26.5 16 24.5 T30 24.5", null, "ground", 1.5),
            },
            ["BENCH"] = new[]
            {
                Ly("M2 27 L2 22 L8 22 L8 17 L14 17 L14 12 L20 12 L20 7 L30 7 L30 27 Z", "fill", null, 0, null, 0.4),
                Ly("M2 22 L8 22 L8 17 L14 17 L14 12 L20 12 L20 7 L30 7", null, "design", 1.8),
                Ly("M2 27 L30 5", null, "ground", 1.5),
                Ly("M2 4 H30", null, "ink2", 0.9, new double[] { 2, 1.5 }),
            },
            ["PAVEMENT"] = new[]
            {
                Ly("M3 6 H29 V11 H3 Z", "asphalt", "ink", 1),
                Ly("M3 11 H29 V17 H3 Z", "base", "ink", 1),
                Ly("M3 17 H29 V23 H3 Z", "subbase", "ink", 1),
                Ly("M3 23 H29 V27 H3 Z", "ground", "ink", 1, null, 0.5),
                Ly("M5 13 h1.2 M7 15 h1.2 M9 13 h1.2 M11 15 h1.2 M13 13 h1.2 M15 15 h1.2 M17 13 h1.2 M19 15 h1.2 M21 13 h1.2 M23 15 h1.2 M25 13 h1.2 M27 15 h1.2", null, "ink", 1, null, 1, 0.75),
                Ly("M5 20 h0.01 M8 20 h0.01 M11 20 h0.01 M14 20 h0.01 M17 20 h0.01 M20 20 h0.01 M23 20 h0.01 M26 20 h0.01", null, "ink", 1.4, null, 1, 0.6),
            },
            ["ADD"] = new[]
            {
                Ly("M4 16 A12 12 0 1 0 28 16 A12 12 0 1 0 4 16 Z", "green", null, 0),
                Ly("M16 9.5 V22.5 M9.5 16 H22.5", null, "white", 3),
            },
            ["DELETE"] = new[]
            {
                Ly("M4 16 A12 12 0 1 0 28 16 A12 12 0 1 0 4 16 Z", "red", null, 0),
                Ly("M11.3 11.3 L20.7 20.7 M20.7 11.3 L11.3 20.7", null, "white", 3),
            },
            ["INFO"] = new[]
            {
                Ly("M4 16 A12 12 0 1 0 28 16 A12 12 0 1 0 4 16 Z", "blue", null, 0),
                Ly("M14.1 10.2 A1.9 1.9 0 1 0 17.9 10.2 A1.9 1.9 0 1 0 14.1 10.2 Z", "white", null, 0),
                Ly("M16 14.5 V23", null, "white", 3.2),
            },
            ["RENAME"] = new[]
            {
                Ly("M4.5 9 H27.5 A2 2 0 0 1 29.5 11 V21 A2 2 0 0 1 27.5 23 H4.5 A2 2 0 0 1 2.5 21 V11 A2 2 0 0 1 4.5 9 Z", "doc", "ink2", 1.2),
                Ly("M6.5 16 H17", null, "doc_ink", 2.2),
                Ly("M21 11.5 V20.5 M19 11.5 H23 M19 20.5 H23", null, "blue", 1.5),
            },
            ["TABLE"] = new[]
            {
                Ly("M3 5 H29 V27 H3 Z", "doc", "ink2", 1.2),
                Ly("M3 5 H29 V10 H3 Z", "blue", null, 0),
                Ly("M3 15.5 H29 M3 21 H29 M12 10 V27 M20.5 10 V27", null, "doc_line", 1.1),
            },
            ["EXCEL"] = new[]
            {
                Ly("M6 3 H21 L27 9 V29 H6 Z", "doc", "ink2", 1.1),
                Ly("M21 3 V9 H27", null, "ink2", 1.1),
                Ly("M9.5 13 H23.5 V25.5 H9.5 Z M9.5 17.2 H23.5 M9.5 21.4 H23.5 M14.5 13 V25.5", null, "green", 1.2),
                Ly("M3 15 H11 A1.5 1.5 0 0 1 12.5 16.5 V24.5 A1.5 1.5 0 0 1 11 26 H3 A1.5 1.5 0 0 1 1.5 24.5 V16.5 A1.5 1.5 0 0 1 3 15 Z", "green", null, 0),
                Ly("M4.3 17.8 L9.7 23.2 M9.7 17.8 L4.3 23.2", null, "white", 1.8),
            },
            ["CAD"] = new[]
            {
                Ly("M3 5 H29 V27 H3 Z", "doc", "ink2", 1.1),
                Ly("M18 21 H29 M18 21 V27 M18 24 H29", null, "doc_line", 1),
                Ly("M6 18 L10 9 L15 13 L21 7 L26 12", null, "design", 1.6),
                Ly("M5 17 H7 V19 H5 Z M9 8 H11 V10 H9 Z M14 12 H16 V14 H14 Z M20 6 H22 V8 H20 Z M25 11 H27 V13 H25 Z", "blue", null, 0),
            },
            ["EXPORT"] = new[]
            {
                Ly("M4 3 H17 L22 8 V29 H4 Z", "doc", "ink2", 1.1),
                Ly("M17 3 V8 H22", null, "ink2", 1.1),
                Ly("M7.5 13 H16 M7.5 17 H14 M7.5 21 H16", null, "doc_line", 1.2),
                Ly("M14 18 H28 M23.5 13.5 L28 18 L23.5 22.5", null, "blue", 2.4),
            },
            ["SYNC"] = new[]
            {
                Ly("M25.5 12.5 A10 10 0 0 0 7.2 11", null, "blue", 2.4),
                Ly("M6.5 19.5 A10 10 0 0 0 24.8 21", null, "blue", 2.4),
                Ly("M4.5 6.5 L7 11.6 L12.2 9.5", null, "blue", 2.4),
                Ly("M27.5 25.5 L25 20.4 L19.8 22.5", null, "blue", 2.4),
            },
            ["RELOAD"] = new[]
            {
                Ly("M26 16 A10 10 0 1 1 22.6 8.5", null, "green", 2.6),
                Ly("M23.5 3.5 L23 9.2 L17.3 9", null, "green", 2.6),
            },
            ["SETTINGS"] = new[]
            {
                Ly("M16 6.4 L18.44 3.74 L20.78 4.45 L21.33 8.02 L22.79 9.21 L26.39 9.06 L27.55 11.22 L25.42 14.13 L25.6 16 L28.26 18.44 L27.55 20.78 L23.98 21.33 L22.79 22.79 L22.94 26.39 L20.78 27.55 L17.87 25.42 L16 25.6 L13.56 28.26 L11.22 27.55 L10.67 23.98 L9.21 22.79 L5.61 22.94 L4.45 20.78 L6.58 17.87 L6.4 16 L3.74 13.56 L4.45 11.22 L8.02 10.67 L9.21 9.21 L9.06 5.61 L11.22 4.45 L14.13 6.58 Z M11.8 16 A4.2 4.2 0 1 0 20.2 16 A4.2 4.2 0 1 0 11.8 16 Z", "ink2", "ink2", 0.8),
            },
            ["UTILITY"] = new[]
            {
                Ly("M5 11 H27 A2 2 0 0 1 29 13 V25 A2 2 0 0 1 27 27 H5 A2 2 0 0 1 3 25 V13 A2 2 0 0 1 5 11 Z", "accent", "accent_dark", 1.1),
                Ly("M12 11 V7.5 H20 V11", null, "ink2", 2),
                Ly("M3 17.5 H29", null, "accent_dark", 1.2),
                Ly("M14 15.5 H18 V19.5 H14 Z", "ink", null, 0),
            },
            ["ACCOUNT"] = new[]
            {
                Ly("M10.5 11 A5.5 5.5 0 1 0 21.5 11 A5.5 5.5 0 1 0 10.5 11 Z", "ink2", null, 0),
                Ly("M5.5 28 C6 20.5 10.5 18 16 18 C21.5 18 26 20.5 26.5 28 Z", "ink2", null, 0),
            },
            ["CALCULATE"] = new[]
            {
                Ly("M8.5 2.5 H23.5 A2.5 2.5 0 0 1 26 5 V27 A2.5 2.5 0 0 1 23.5 29.5 H8.5 A2.5 2.5 0 0 1 6 27 V5 A2.5 2.5 0 0 1 8.5 2.5 Z", "ink2", null, 0),
                Ly("M9 5.5 H23 V11.5 H9 Z", "lcd", null, 0),
                Ly("M9 14 H12.6 V17.2 H9 Z", "key", null, 0),
                Ly("M14 14 H17.6 V17.2 H14 Z", "key", null, 0),
                Ly("M19 14 H22.6 V17.2 H19 Z", "key", null, 0),
                Ly("M9 18.6 H12.6 V21.8 H9 Z", "key", null, 0),
                Ly("M14 18.6 H17.6 V21.8 H14 Z", "key", null, 0),
                Ly("M19 18.6 H22.6 V21.8 H19 Z", "key", null, 0),
                Ly("M9 23.2 H12.6 V26.4 H9 Z", "key", null, 0),
                Ly("M14 23.2 H17.6 V26.4 H14 Z", "key", null, 0),
                Ly("M19 23.2 H22.6 V26.4 H19 Z", "accent", null, 0),
            },
            ["DRAW"] = new[]
            {
                Ly("M4 25 L11 9 L19 19 L28 6", null, "ink", 1.8),
                Ly("M2.4 23.4 H5.6 V26.6 H2.4 Z M9.4 7.4 H12.6 V10.6 H9.4 Z M17.4 17.4 H20.6 V20.6 H17.4 Z M26.4 4.4 H29.6 V7.6 H26.4 Z", "blue", "blue_dark", 0.7),
            },
            ["VIEW"] = new[]
            {
                Ly("M2.5 16 C7 8.5 25 8.5 29.5 16 C25 23.5 7 23.5 2.5 16 Z", "doc", "ink2", 1.3),
                Ly("M10.8 16 A5.2 5.2 0 1 0 21.2 16 A5.2 5.2 0 1 0 10.8 16 Z", "blue", null, 0),
                Ly("M13.8 16 A2.2 2.2 0 1 0 18.2 16 A2.2 2.2 0 1 0 13.8 16 Z", "ink_dark", null, 0),
            },
            ["COPY"] = new[]
            {
                Ly("M4.5 3 H18.5 A1.5 1.5 0 0 1 20 4.5 V21.5 A1.5 1.5 0 0 1 18.5 23 H4.5 A1.5 1.5 0 0 1 3 21.5 V4.5 A1.5 1.5 0 0 1 4.5 3 Z", "doc", "ink2", 1.2),
                Ly("M13.5 9 H27.5 A1.5 1.5 0 0 1 29 10.5 V27.5 A1.5 1.5 0 0 1 27.5 29 H13.5 A1.5 1.5 0 0 1 12 27.5 V10.5 A1.5 1.5 0 0 1 13.5 9 Z", "doc", "blue", 1.6),
                Ly("M15.5 15 H25.5 M15.5 19 H25.5 M15.5 23 H22", null, "doc_line", 1.2),
            },
            ["MOVE"] = new[]
            {
                Ly("M16 3 V29 M3 16 H29", null, "ink", 2),
                Ly("M12 7 L16 3 L20 7 M12 25 L16 29 L20 25 M7 12 L3 16 L7 20 M25 12 L29 16 L25 20", null, "ink", 2),
                Ly("M13.7 16 A2.3 2.3 0 1 0 18.3 16 A2.3 2.3 0 1 0 13.7 16 Z", "blue", null, 0),
            },
            ["EDIT"] = new[]
            {
                Ly("M6 26 L5 27 L4.5 23.5 L21.5 6.5 L25.5 10.5 L8.5 27.5 Z", "pencil", "accent_dark", 1),
                Ly("M21.5 6.5 L23.5 4.5 Q24.5 3.5 25.5 4.5 L27.5 6.5 Q28.5 7.5 27.5 8.5 L25.5 10.5 Z", "red", "red_dark", 0.8),
                Ly("M4.5 23.5 L8.5 27.5 L3.5 28.5 Z", "ink", null, 0),
            },
            ["MEASURE"] = new[]
            {
                Ly("M5 5 V23 M27 5 V23", null, "ink2", 1.2),
                Ly("M5 17 H27", null, "ink", 1.5),
                Ly("M5 17 L9.5 15 L9.5 19 Z M27 17 L22.5 15 L22.5 19 Z", "ink", null, 0),
                Ly("M11 11.5 H21", null, "blue", 2.2),
                Ly("M3 27 L29 27", null, "design", 1.8),
            },
            ["LAYOUT"] = new[]
            {
                Ly("M2.5 5 H29.5 V27 H2.5 Z", "doc", "ink2", 1.2),
                Ly("M5.5 8 H20.5 V20 H5.5 Z", null, "blue", 1.4),
                Ly("M7.5 17 L11 12 L14 15 L18.5 10.5", null, "design", 1.3),
                Ly("M22.5 8 V24 M22.5 20 H29.5 M22.5 22 H29.5", null, "doc_line", 1),
                Ly("M5.5 23 H20.5", null, "doc_line", 1),
            },
            ["PRINT"] = new[]
            {
                Ly("M9 3 H23 V11 H9 Z", "doc", "ink2", 1.1),
                Ly("M5 10 H27 A2 2 0 0 1 29 12 V20 A2 2 0 0 1 27 22 H5 A2 2 0 0 1 3 20 V12 A2 2 0 0 1 5 10 Z", "ink2", null, 0),
                Ly("M9 18 H23 V29 H9 Z", "doc", "ink2", 1.1),
                Ly("M11.5 22 H20.5 M11.5 25 H18", null, "doc_line", 1.1),
                Ly("M23.7 14 A1.3 1.3 0 1 0 26.3 14 A1.3 1.3 0 1 0 23.7 14 Z", "green", null, 0),
            },
        };

        private static readonly Dictionary<string, string> PaletteDark = new()
        {
            ["ink"] = "#E4E8EE",
            ["ink2"] = "#A8B1BC",
            ["ink_dark"] = "#1E2228",
            ["panel"] = "#2B3139",
            ["ground"] = "#C98A4B",
            ["design"] = "#FF5A4A",
            ["align"] = "#4DA6FF",
            ["accent"] = "#FFB52E",
            ["accent_dark"] = "#B97800",
            ["water"] = "#4FC3F7",
            ["water_dark"] = "#0B6EA6",
            ["cut"] = "#F0503F",
            ["fill"] = "#43C46A",
            ["asphalt"] = "#5E656F",
            ["asphalt_edge"] = "#8A939E",
            ["asphalt_top"] = "#8A939E",
            ["shoulder"] = "#8B7B62",
            ["road_mark"] = "#F2F4F7",
            ["road_mark_y"] = "#FFD54F",
            ["concrete"] = "#B9C0C8",
            ["base"] = "#C9A240",
            ["subbase"] = "#9C8358",
            ["elev_low"] = "#3E9E5C",
            ["elev_lowmid"] = "#8BAF3F",
            ["elev_mid"] = "#C9A93D",
            ["elev_highmid"] = "#C88A3E",
            ["elev_high"] = "#A8683A",
            ["green"] = "#2EA44F",
            ["red"] = "#E5484D",
            ["red_dark"] = "#9E2A2E",
            ["blue"] = "#3B8EEA",
            ["blue_dark"] = "#1C5FB0",
            ["white"] = "#FFFFFF",
            ["doc"] = "#F3F5F8",
            ["doc_ink"] = "#3A414B",
            ["doc_line"] = "#7D8794",
            ["lcd"] = "#BFE6C8",
            ["key"] = "#E4E8EE",
            ["pencil"] = "#FFC23D",
        };

        private static readonly Dictionary<string, string> PaletteLight = new()
        {
            ["ink"] = "#2F353D",
            ["ink2"] = "#5F6873",
            ["ink_dark"] = "#1E2228",
            ["panel"] = "#FFFFFF",
            ["ground"] = "#9A5F2A",
            ["design"] = "#D93425",
            ["align"] = "#1C6ED6",
            ["accent"] = "#E8930C",
            ["accent_dark"] = "#9A5E00",
            ["water"] = "#2A9FD8",
            ["water_dark"] = "#0B5E8E",
            ["cut"] = "#D93A2B",
            ["fill"] = "#2E9E4F",
            ["asphalt"] = "#5A6069",
            ["asphalt_edge"] = "#3C4148",
            ["asphalt_top"] = "#3C4148",
            ["shoulder"] = "#B5A487",
            ["road_mark"] = "#FFFFFF",
            ["road_mark_y"] = "#F5C518",
            ["concrete"] = "#C3C9D0",
            ["base"] = "#D6B04C",
            ["subbase"] = "#AE9468",
            ["elev_low"] = "#4BAE6A",
            ["elev_lowmid"] = "#9CBF4C",
            ["elev_mid"] = "#D9B848",
            ["elev_highmid"] = "#D29546",
            ["elev_high"] = "#B07040",
            ["green"] = "#2E9E4F",
            ["red"] = "#D93A3F",
            ["red_dark"] = "#8E2226",
            ["blue"] = "#1F78D1",
            ["blue_dark"] = "#13518F",
            ["white"] = "#FFFFFF",
            ["doc"] = "#FFFFFF",
            ["doc_ink"] = "#2F353D",
            ["doc_line"] = "#7D8794",
            ["lcd"] = "#BFE6C8",
            ["key"] = "#FFFFFF",
            ["pencil"] = "#FFC23D",
        };

        private static readonly Dictionary<string, ImageSource> Cache = new();
        private static readonly Dictionary<string, WColor> ColorCache = new();

        /// <summary>AutoCAD đang dùng theme tối? (COLORTHEME = 0)</summary>
        public static bool IsDarkTheme()
        {
            try { return Convert.ToInt32(AcadApplication.GetSystemVariable("COLORTHEME")) == 0; }
            catch { return true; }
        }

        /// <summary>Xóa cache (gọi khi đổi theme rồi dựng lại ribbon)</summary>
        public static void ClearCache() => Cache.Clear();

        public static bool Has(string iconType) => !string.IsNullOrEmpty(iconType) && Icons.ContainsKey(iconType);

        /// <summary>
        /// Vẽ icon kích thước <paramref name="size"/> (16 / 32...). Render 2x để sắc nét trên màn hình HiDPI.
        /// Tên không có trong thư viện sẽ dùng icon UTILITY.
        /// </summary>
        public static ImageSource Render(string iconType, int size)
            => Render(iconType, size, IsDarkTheme());

        /// <summary>Vẽ icon với bảng màu chỉ định (dark = nền tối, false = nền sáng như cây Toolspace)</summary>
        public static ImageSource Render(string iconType, int size, bool dark)
        {
            string name = Has(iconType) ? iconType.ToUpperInvariant() : "UTILITY";
            string key = $"{name}|{size}|{(dark ? "D" : "L")}";
            if (Cache.TryGetValue(key, out var cached)) return cached;

            var palette = dark ? PaletteDark : PaletteLight;
            var visual = new DrawingVisual();
            using (DrawingContext dc = visual.RenderOpen())
            {
                double k = size / 32.0;
                dc.PushTransform(new ScaleTransform(k, k));
                foreach (var layer in Icons[name])
                {
                    Geometry geometry;
                    try { geometry = Geometry.Parse(layer.D); }
                    catch { continue; }
                    geometry.Freeze();

                    System.Windows.Media.Brush? fill = MakeBrush(palette, layer.Fill, layer.Fo);
                    System.Windows.Media.Pen? pen = null;
                    var strokeBrush = MakeBrush(palette, layer.Stroke, layer.So);
                    if (strokeBrush != null && layer.Sw > 0)
                    {
                        pen = new System.Windows.Media.Pen(strokeBrush, layer.Sw)
                        {
                            StartLineCap = PenLineCap.Round,
                            EndLineCap = PenLineCap.Round,
                            DashCap = PenLineCap.Round,
                            LineJoin = PenLineJoin.Round,
                        };
                        if (layer.Dash != null && layer.Dash.Length > 0)
                            pen.DashStyle = new DashStyle(layer.Dash.Select(x => x / layer.Sw), 0);
                        pen.Freeze();
                    }
                    if (fill == null && pen == null) continue;
                    dc.DrawGeometry(fill, pen, geometry);
                }
                dc.Pop();
            }

            const int scale = 2;
            var bmp = new RenderTargetBitmap(size * scale, size * scale, 96 * scale, 96 * scale, PixelFormats.Pbgra32);
            bmp.Render(visual);
            bmp.Freeze();
            Cache[key] = bmp;
            return bmp;
        }

        private static System.Windows.Media.Brush? MakeBrush(Dictionary<string, string> palette, string? role, double opacity)
        {
            if (string.IsNullOrEmpty(role) || !palette.TryGetValue(role, out var hex)) return null;
            if (!ColorCache.TryGetValue(hex, out var c))
            {
                c = (WColor)System.Windows.Media.ColorConverter.ConvertFromString(hex);
                ColorCache[hex] = c;
            }
            if (opacity < 1) c = WColor.FromArgb((byte)Math.Round(255 * opacity), c.R, c.G, c.B);
            var brush = new SolidColorBrush(c);
            brush.Freeze();
            return brush;
        }
    }
}
