using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.MoHinh.RebarDataExporter))]

namespace MyFirstProject.MoHinh
{
    #region === Data Models ===

    /// <summary>
    /// Thông tin một thanh thép đơn lẻ.
    /// </summary>
    public class RebarBar
    {
        [JsonPropertyName("id")]
        public string Id { get; set; } = "";

        [JsonPropertyName("groupName")]
        public string GroupName { get; set; } = "";

        [JsonPropertyName("diameter")]
        public double Diameter { get; set; }

        [JsonPropertyName("length")]
        public double Length { get; set; }

        [JsonPropertyName("weight")]
        public double Weight { get; set; }

        [JsonPropertyName("shapeType")]
        public string ShapeType { get; set; } = "Straight"; // Straight, LBend, UBend, Stirrup, Custom

        [JsonPropertyName("layer")]
        public string Layer { get; set; } = "";

        [JsonPropertyName("color")]
        public int[] Color { get; set; } = [255, 0, 0];

        [JsonPropertyName("vertices")]
        public double[][] Vertices { get; set; } = [];

        [JsonPropertyName("direction")]
        public string Direction { get; set; } = "Longitudinal"; // Longitudinal, Transverse, Vertical
    }

    /// <summary>
    /// Nhóm thép (G1, G2, G3...).
    /// </summary>
    public class RebarGroup
    {
        [JsonPropertyName("name")]
        public string Name { get; set; } = "";

        [JsonPropertyName("diameter")]
        public double Diameter { get; set; }

        [JsonPropertyName("count")]
        public int Count { get; set; }

        [JsonPropertyName("totalLength")]
        public double TotalLength { get; set; }

        [JsonPropertyName("totalWeight")]
        public double TotalWeight { get; set; }

        [JsonPropertyName("color")]
        public int[] Color { get; set; } = [255, 0, 0];

        [JsonPropertyName("bars")]
        public List<RebarBar> Bars { get; set; } = [];
    }

    /// <summary>
    /// Thông tin cấu kiện (Xà mũ, Dầm, Cột...).
    /// </summary>
    public class StructuralMember
    {
        [JsonPropertyName("name")]
        public string Name { get; set; } = "";

        [JsonPropertyName("type")]
        public string Type { get; set; } = "CapBeam"; // CapBeam, Beam, Column, Slab, Wall

        [JsonPropertyName("width")]
        public double Width { get; set; }

        [JsonPropertyName("height")]
        public double Height { get; set; }

        [JsonPropertyName("depth")]
        public double Depth { get; set; }

        [JsonPropertyName("concreteVolume")]
        public double ConcreteVolume { get; set; }

        [JsonPropertyName("totalBars")]
        public int TotalBars { get; set; }

        [JsonPropertyName("totalWeight")]
        public double TotalWeight { get; set; }

        [JsonPropertyName("steelRatio")]
        public double SteelRatio { get; set; } // kg/m³

        [JsonPropertyName("outlineVertices")]
        public double[][] OutlineVertices { get; set; } = [];

        [JsonPropertyName("groups")]
        public List<RebarGroup> Groups { get; set; } = [];
    }

    /// <summary>
    /// Root model chứa toàn bộ dữ liệu xuất ra.
    /// </summary>
    public class RebarExportData
    {
        [JsonPropertyName("projectName")]
        public string ProjectName { get; set; } = "";

        [JsonPropertyName("fileName")]
        public string FileName { get; set; } = "";

        [JsonPropertyName("exportDate")]
        public string ExportDate { get; set; } = "";

        [JsonPropertyName("unit")]
        public string Unit { get; set; } = "mm";

        [JsonPropertyName("members")]
        public List<StructuralMember> Members { get; set; } = [];
    }

    #endregion

    #region === Utility: Tính trọng lượng thép ===

    public static class RebarUtils
    {
        /// <summary>
        /// Bảng trọng lượng đơn vị (kg/m) theo đường kính (mm).
        /// </summary>
        private static readonly Dictionary<int, double> WeightPerMeter = new()
        {
            { 6, 0.222 }, { 8, 0.395 }, { 10, 0.617 },
            { 12, 0.888 }, { 14, 1.208 }, { 16, 1.578 },
            { 18, 1.998 }, { 20, 2.466 }, { 22, 2.984 },
            { 25, 3.853 }, { 28, 4.834 }, { 32, 6.313 },
            { 36, 7.990 }, { 40, 9.864 }
        };

        /// <summary>
        /// Tính trọng lượng thanh thép (kg).
        /// </summary>
        public static double CalcWeight(double diameterMm, double lengthMm)
        {
            int d = (int)Math.Round(diameterMm);
            double wpm = WeightPerMeter.ContainsKey(d) ? WeightPerMeter[d] : (d * d * Math.PI / 4.0) * 7850.0 / 1e6;
            return wpm * lengthMm / 1000.0;
        }

        /// <summary>
        /// Trích xuất đường kính từ tên layer (VD: "THEP_D16" → 16, "D20_DAI" → 20).
        /// </summary>
        public static int ExtractDiameter(string layerName)
        {
            // Pattern: D + số
            var match = System.Text.RegularExpressions.Regex.Match(layerName, @"D(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (match.Success && int.TryParse(match.Groups[1].Value, out int d))
                return d;
            return 0;
        }

        /// <summary>
        /// Phân loại hình dạng thanh thép dựa trên vertices.
        /// </summary>
        public static string ClassifyShape(Point3d[] pts)
        {
            if (pts.Length <= 2) return "Straight";
            if (pts.Length == 3) return "LBend";
            if (pts.Length == 4)
            {
                // Kiểm tra nếu là hình chữ U
                var v1 = pts[1] - pts[0];
                var v2 = pts[2] - pts[1];
                var v3 = pts[3] - pts[2];
                if (Math.Abs(v1.DotProduct(v2)) < 1e-3 && Math.Abs(v2.DotProduct(v3)) < 1e-3)
                    return "UBend";
            }
            if (pts.Length >= 5)
            {
                // Kiểm tra nếu khép kín → Stirrup (đai)
                if (pts[0].DistanceTo(pts[pts.Length - 1]) < 10.0)
                    return "Stirrup";
            }
            return "Custom";
        }

        /// <summary>
        /// Xác định hướng thép dựa trên vector chính.
        /// </summary>
        public static string ClassifyDirection(Point3d[] pts)
        {
            if (pts.Length < 2) return "Longitudinal";

            var mainVector = pts[pts.Length - 1] - pts[0];
            double absX = Math.Abs(mainVector.X);
            double absY = Math.Abs(mainVector.Y);
            double absZ = Math.Abs(mainVector.Z);

            if (absX >= absY && absX >= absZ) return "Longitudinal";
            if (absY >= absX && absY >= absZ) return "Transverse";
            return "Vertical";
        }

        /// <summary>
        /// Tạo màu đại diện cho đường kính thép.
        /// </summary>
        public static int[] GetColorForDiameter(int diameter)
        {
            return diameter switch
            {
                6 => [128, 128, 128],    // Xám
                8 => [0, 191, 255],      // Xanh dương nhạt
                10 => [0, 128, 255],     // Xanh dương
                12 => [0, 255, 0],       // Xanh lá
                14 => [255, 255, 0],     // Vàng
                16 => [255, 165, 0],     // Cam
                18 => [255, 100, 100],   // Hồng
                20 => [255, 0, 0],       // Đỏ
                22 => [200, 0, 200],     // Tím
                25 => [128, 0, 255],     // Tím đậm
                28 => [0, 100, 0],       // Xanh lá đậm
                32 => [139, 69, 19],     // Nâu
                36 => [75, 0, 130],      // Indigo
                40 => [0, 0, 0],         // Đen
                _ => [100, 100, 100]
            };
        }

        /// <summary>
        /// Lấy màu thực tế của entity trong CAD (RGB).
        /// </summary>
        public static int[] GetEntityColor(Entity ent, Transaction tr)
        {
            Autodesk.AutoCAD.Colors.Color color = ent.Color;

            // Nếu ByLayer, lấy màu từ Layer
            if (color.IsByLayer)
            {
                LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(ent.LayerId, OpenMode.ForRead);
                color = ltr.Color;
            }

            if (color.ColorMethod == ColorMethod.ByAci)
            {
                // Map ACI color to approximate RGB
                byte r = color.ColorValue.R;
                byte g = color.ColorValue.G;
                byte b = color.ColorValue.B;
                return [r, g, b];
            }

            return [color.Red, color.Green, color.Blue];
        }
    }

    #endregion

    /// <summary>
    /// Lệnh AutoCAD: AT_RebarBIM
    /// Đọc dữ liệu cốt thép từ bản vẽ CAD và xuất ra:
    /// 1. File JSON chứa dữ liệu cốt thép
    /// 2. File HTML tự chứa (self-contained) dùng Three.js để trực quan hóa 3D
    /// </summary>
    public class RebarDataExporter
    {
        [CommandMethod("AT_RebarBIM")]
        public static void ExportRebarBIM()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // ===== Bước 1: Mở Hộp thoại Nhập thông tin cấu kiện (UI WinForms chuyên nghiệp) =====
                string defaultName = Path.GetFileNameWithoutExtension(doc.Name);
                if (string.IsNullOrWhiteSpace(defaultName) || defaultName.StartsWith("Drawing", StringComparison.OrdinalIgnoreCase))
                    defaultName = "T3_XaMu";

                string defaultFolder = Path.GetDirectoryName(doc.Name) ?? "";
                if (string.IsNullOrWhiteSpace(defaultFolder))
                    defaultFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                using var form = new RebarBIMForm(defaultName, defaultFolder);
                Application.ShowModalDialog(form);

                if (!form.IsConfirmed)
                {
                    ed.WriteMessage("\n[AT_RebarBIM] Đã hủy thao tác.");
                    return;
                }

                string memberName = form.MemberName;
                string memberType = form.MemberType;
                double memberWidth = form.MemberWidth;
                double memberHeight = form.MemberHeight;
                double memberDepth = form.MemberDepth;
                bool selectArea = form.SelectArea;
                string outputFolder = string.IsNullOrWhiteSpace(form.OutputFolder) ? defaultFolder : form.OutputFolder;
                bool autoOpen = form.AutoOpenHtml;
                bool useTableMode = form.UseTableDataMode;
                var tableItems = form.RebarItems;

                // ===== Bước 2: Thu thập đối tượng thép =====
                var allBars = new List<RebarBar>();
                int barIndex = 0;

                if (useTableMode && tableItems.Count > 0)
                {
                    ed.WriteMessage($"\n\n[AT_RebarBIM] Tạo mô hình BIM trực tiếp từ Bảng thông số ({tableItems.Count} nhóm thép)...");
                    allBars = GenerateBarsFromTable(memberWidth, memberHeight, memberDepth, tableItems);
                }
                else
                {
                    ed.WriteMessage("\n\nĐang quét dữ liệu cốt thép từ bản vẽ CAD...");
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                    // Lấy tất cả layer có chứa "D" + số (pattern thép)
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    var rebarLayers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    foreach (ObjectId layerId in lt)
                    {
                        LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                        if (RebarUtils.ExtractDiameter(ltr.Name) > 0)
                        {
                            rebarLayers.Add(ltr.Name);
                        }
                    }

                    ed.WriteMessage($"\n  Tìm thấy {rebarLayers.Count} layer thép: {string.Join(", ", rebarLayers)}");

                    // Lấy các đối tượng cần xử lý
                    ObjectIdCollection idsToProcess;

                    if (selectArea)
                    {
                        // Cho người dùng chọn vùng
                        var selOpt = new PromptSelectionOptions();
                        selOpt.MessageForAdding = "\nChọn các đối tượng thép: ";
                        var selRes = ed.GetSelection(selOpt);
                        if (selRes.Status != PromptStatus.OK) return;

                        idsToProcess = new ObjectIdCollection();
                        foreach (ObjectId id in selRes.Value.GetObjectIds())
                            idsToProcess.Add(id);
                    }
                    else
                    {
                        // Lấy toàn bộ từ Model Space
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                        idsToProcess = new ObjectIdCollection();
                        foreach (ObjectId entId in modelSpace)
                            idsToProcess.Add(entId);
                    }

                    // Duyệt các đối tượng
                    foreach (ObjectId entId in idsToProcess)
                    {
                        Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;

                        // Chỉ xử lý đối tượng trên layer thép
                        if (!rebarLayers.Contains(ent.Layer)) continue;

                        int diameter = RebarUtils.ExtractDiameter(ent.Layer);
                        if (diameter <= 0) continue;

                        Point3d[] pts = null;

                        // Xử lý Polyline
                        if (ent is Polyline pl)
                        {
                            pts = new Point3d[pl.NumberOfVertices];
                            for (int i = 0; i < pl.NumberOfVertices; i++)
                                pts[i] = pl.GetPoint3dAt(i);
                        }
                        // Xử lý Polyline2d
                        else if (ent is Polyline2d pl2d)
                        {
                            var ptsList = new List<Point3d>();
                            foreach (ObjectId vId in pl2d)
                            {
                                Vertex2d v = (Vertex2d)tr.GetObject(vId, OpenMode.ForRead);
                                ptsList.Add(v.Position);
                            }
                            pts = ptsList.ToArray();
                        }
                        // Xử lý Polyline3d
                        else if (ent is Polyline3d pl3d)
                        {
                            var ptsList = new List<Point3d>();
                            foreach (ObjectId vId in pl3d)
                            {
                                PolylineVertex3d v = (PolylineVertex3d)tr.GetObject(vId, OpenMode.ForRead);
                                ptsList.Add(v.Position);
                            }
                            pts = ptsList.ToArray();
                        }
                        // Xử lý Line
                        else if (ent is Line line)
                        {
                            pts = [line.StartPoint, line.EndPoint];
                        }

                        if (pts == null || pts.Length < 2) continue;

                        // Tính chiều dài tổng
                        double totalLen = 0;
                        for (int i = 0; i < pts.Length - 1; i++)
                            totalLen += pts[i].DistanceTo(pts[i + 1]);

                        // Tạo RebarBar
                        barIndex++;
                        var bar = new RebarBar
                        {
                            Id = $"B{barIndex:D4}",
                            Diameter = diameter,
                            Length = Math.Round(totalLen, 1),
                            Weight = Math.Round(RebarUtils.CalcWeight(diameter, totalLen), 1),
                            ShapeType = RebarUtils.ClassifyShape(pts),
                            Layer = ent.Layer,
                            Color = RebarUtils.GetEntityColor(ent, tr),
                            Direction = RebarUtils.ClassifyDirection(pts),
                            Vertices = pts.Select(p => new double[] { Math.Round(p.X, 1), Math.Round(p.Y, 1), Math.Round(p.Z, 1) }).ToArray()
                        };

                        allBars.Add(bar);
                    }

                    // ===== Bước 4: Tìm nhóm thép từ Text/MText/Block gần các thanh thép =====
                    AssignGroupNames(allBars, idsToProcess, tr);

                    tr.Commit();
                    } // end using transaction

                    ed.WriteMessage($"\n  Thu thập được {allBars.Count} thanh thép từ CAD.");

                    if (allBars.Count == 0)
                    {
                        ed.WriteMessage("\n\n⚠ Không tìm thấy thanh thép nào trên các layer D* trong CAD.");
                        ed.WriteMessage("\n  Sẽ tạo mô hình từ Bảng thông số cốt thép...");

                        allBars = tableItems.Count > 0
                            ? GenerateBarsFromTable(memberWidth, memberHeight, memberDepth, tableItems)
                            : GenerateSampleData(memberWidth, memberHeight, memberDepth);
                    }
                } // end else (quét CAD)

                // ===== Bước 5: Nhóm và thống kê =====
                var groups = allBars
                    .GroupBy(b => string.IsNullOrEmpty(b.GroupName) ? $"D{b.Diameter}" : b.GroupName)
                    .Select(g =>
                    {
                        var first = g.First();
                        return new RebarGroup
                        {
                            Name = g.Key,
                            Diameter = first.Diameter,
                            Count = g.Count(),
                            TotalLength = Math.Round(g.Sum(b => b.Length), 1),
                            TotalWeight = Math.Round(g.Sum(b => b.Weight), 1),
                            Color = first.Color,
                            Bars = g.ToList()
                        };
                    })
                    .OrderBy(g => g.Name)
                    .ToList();

                double concreteVol = memberWidth * memberHeight * memberDepth / 1e9; // m³
                double totalWeight = groups.Sum(g => g.TotalWeight);

                var member = new StructuralMember
                {
                    Name = memberName,
                    Type = memberType,
                    Width = memberWidth,
                    Height = memberHeight,
                    Depth = memberDepth,
                    ConcreteVolume = Math.Round(concreteVol, 3),
                    TotalBars = allBars.Count,
                    TotalWeight = Math.Round(totalWeight, 1),
                    SteelRatio = concreteVol > 0 ? Math.Round(totalWeight / concreteVol, 1) : 0,
                    OutlineVertices = GenerateOutlineVertices(memberWidth, memberHeight, memberDepth),
                    Groups = groups
                };

                var exportData = new RebarExportData
                {
                    ProjectName = memberName,
                    FileName = doc.Name,
                    ExportDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    Members = [member]
                };

                // ===== Bước 6: Xuất file =====
                Directory.CreateDirectory(outputFolder);
                string jsonPath = Path.Combine(outputFolder, $"{memberName}_CotThep.json");
                string htmlPath = Path.Combine(outputFolder, $"{memberName}_CotThep_3D.html");

                // Xuất JSON
                var jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };
                string json = JsonSerializer.Serialize(exportData, jsonOptions);
                File.WriteAllText(jsonPath, json, Encoding.UTF8);

                // Xuất HTML (tự chứa - self-contained)
                string htmlContent = HtmlViewerGenerator.Generate(json, memberName);
                File.WriteAllText(htmlPath, htmlContent, Encoding.UTF8);

                ed.WriteMessage($"\n\n[AT_RebarBIM] ✓ Hoàn tất xuất mô hình BIM cốt thép!");
                ed.WriteMessage($"\n  ● Số thanh thép: {allBars.Count}");
                ed.WriteMessage($"\n  ● Số nhóm: {groups.Count}");
                ed.WriteMessage($"\n  ● Tổng khối lượng: {totalWeight:N1} kg");
                ed.WriteMessage($"\n  ● Hàm lượng thép: {member.SteelRatio:N1} kg/m³");
                ed.WriteMessage($"\n  ● File JSON: {jsonPath}");
                ed.WriteMessage($"\n  ● File HTML 3D: {htmlPath}");
                ed.WriteMessage($"\n\n  → Đã lưu vào: {outputFolder}");

                // Tự động mở HTML trong trình duyệt nếu được chọn
                if (autoOpen)
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = htmlPath,
                            UseShellExecute = true
                        });
                    }
                    catch { /* Bỏ qua nếu không mở được */ }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n\nLỗi: {ex.Message}\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Gán tên nhóm cho các thanh thép dựa trên Text/MText gần đó.
        /// </summary>
        private static void AssignGroupNames(List<RebarBar> bars, ObjectIdCollection idsToProcess, Transaction tr)
        {
            // Thu thập tất cả Text/MText có dạng "G1", "G2", "G3"...
            var groupLabels = new List<(string Name, Point3d Position)>();

            foreach (ObjectId entId in idsToProcess)
            {
                Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                string text = null;
                Point3d pos = Point3d.Origin;

                if (ent is DBText dbText)
                {
                    text = dbText.TextString;
                    pos = dbText.Position;
                }
                else if (ent is MText mText)
                {
                    text = mText.Contents;
                    pos = mText.Location;
                }

                if (text == null) continue;

                // Kiểm tra pattern nhóm thép: G1, G2, G10, v.v.
                var match = System.Text.RegularExpressions.Regex.Match(text.Trim(), @"^G(\d+)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    groupLabels.Add((text.Trim().ToUpper(), pos));
                }
            }

            // Gán nhóm cho thanh thép dựa trên khoảng cách gần nhất
            foreach (var bar in bars)
            {
                if (bar.Vertices.Length == 0 || groupLabels.Count == 0) continue;

                var barCenter = new Point3d(
                    bar.Vertices.Average(v => v[0]),
                    bar.Vertices.Average(v => v[1]),
                    bar.Vertices.Average(v => v[2])
                );

                double minDist = double.MaxValue;
                string closestGroup = "";

                foreach (var (name, pos) in groupLabels)
                {
                    double dist = barCenter.DistanceTo(pos);
                    if (dist < minDist)
                    {
                        minDist = dist;
                        closestGroup = name;
                    }
                }

                // Chỉ gán nếu label đủ gần (trong phạm vi hợp lý)
                if (minDist < 5000) // 5m
                    bar.GroupName = closestGroup;
            }
        }

        /// <summary>
        /// Tạo outline vertices cho hình hộp cấu kiện.
        /// </summary>
        private static double[][] GenerateOutlineVertices(double w, double h, double d)
        {
            return
            [
                [0, 0, 0], [w, 0, 0], [w, h, 0], [0, h, 0],       // Mặt trước
                [0, 0, d], [w, 0, d], [w, h, d], [0, h, d]          // Mặt sau
            ];
        }

        /// <summary>
        /// Tạo các đối tượng thanh cốt thép 3D trực tiếp từ Bảng thông số người dùng đã nhập.
        /// Ánh xạ chính xác từng dòng trong bảng thành các thanh thép 3D phân bố theo quy luật kết cấu.
        /// </summary>
        private static List<RebarBar> GenerateBarsFromTable(double width, double height, double depth, List<RebarTableItem> items)
        {
            var bars = new List<RebarBar>();
            int idx = 0;
            double cover = 60; // mm

            foreach (var item in items)
            {
                int dia = item.Diameter;
                int qty = item.Quantity;
                double len = item.Length;
                string grp = item.GroupCode;
                string layer = item.Layer;
                int[] color = item.ColorRgb;
                string shape = item.ShapeType;

                if (qty <= 0) continue;

                if (shape.Equals("Stirrup", StringComparison.OrdinalIgnoreCase))
                {
                    // Đai phân bố dọc theo chiều dài X
                    double stepX = (width - 2 * cover) / Math.Max(1, qty - 1);
                    double yMin = cover, yMax = height - cover;
                    double zMin = cover, zMax = depth - cover;

                    if (grp.Equals("G4", StringComparison.OrdinalIgnoreCase))
                    {
                        zMin = cover + (depth - 2 * cover) * 0.35;
                        zMax = cover + (depth - 2 * cover) * 0.65;
                    }

                    for (int i = 0; i < qty; i++)
                    {
                        idx++;
                        double x = cover + i * stepX;
                        bars.Add(new RebarBar
                        {
                            Id = $"B{idx:D4}",
                            GroupName = grp,
                            Diameter = dia,
                            Length = len,
                            Weight = Math.Round(RebarUtils.CalcWeight(dia, len), 1),
                            ShapeType = "Stirrup",
                            Layer = layer,
                            Color = color,
                            Direction = "Transverse",
                            Vertices =
                            [
                                [x, yMin, zMin],
                                [x, yMax, zMin],
                                [x, yMax, zMax],
                                [x, yMin, zMax],
                                [x, yMin, zMin]
                            ]
                        });
                    }
                }
                else if (shape.Equals("LBend", StringComparison.OrdinalIgnoreCase))
                {
                    // Thép neo / chờ cọc
                    int piles = Math.Max(2, (int)Math.Round(Math.Sqrt(qty)));
                    int barsPerPile = Math.Max(1, qty / piles);
                    double stepX = (width - 2 * cover) / (piles + 1);

                    for (int p = 1; p <= piles; p++)
                    {
                        double px = cover + p * stepX;
                        double pz = depth / 2.0;

                        for (int b = 0; b < barsPerPile; b++)
                        {
                            idx++;
                            double angle = b * (2 * Math.PI / barsPerPile);
                            double r = 400;
                            double bx = px + r * Math.Cos(angle);
                            double bz = pz + r * Math.Sin(angle);
                            double hookLen = 350;
                            double embed = Math.Max(500, len - hookLen);

                            bars.Add(new RebarBar
                            {
                                Id = $"B{idx:D4}",
                                GroupName = grp,
                                Diameter = dia,
                                Length = len,
                                Weight = Math.Round(RebarUtils.CalcWeight(dia, len), 1),
                                ShapeType = "LBend",
                                Layer = layer,
                                Color = color,
                                Direction = "Vertical",
                                Vertices =
                                [
                                    [bx + hookLen * Math.Cos(angle), cover, bz + hookLen * Math.Sin(angle)],
                                    [bx, cover, bz],
                                    [bx, cover + embed, bz]
                                ]
                            });
                        }
                    }
                }
                else // Straight
                {
                    if (grp.Equals("G1", StringComparison.OrdinalIgnoreCase) || item.Name.Contains("dưới") || item.Name.Contains("đáy"))
                    {
                        // Thép chủ dọc dưới
                        int layers = Math.Max(1, (int)Math.Ceiling(qty / 18.0));
                        int perLayer = (int)Math.Ceiling((double)qty / layers);
                        double stepZ = (depth - 2 * cover) / Math.Max(1, perLayer - 1);

                        int created = 0;
                        for (int l = 0; l < layers && created < qty; l++)
                        {
                            double y = cover + l * 65;
                            for (int i = 0; i < perLayer && created < qty; i++)
                            {
                                idx++;
                                created++;
                                double z = cover + i * stepZ;
                                bars.Add(new RebarBar
                                {
                                    Id = $"B{idx:D4}",
                                    GroupName = grp,
                                    Diameter = dia,
                                    Length = len,
                                    Weight = Math.Round(RebarUtils.CalcWeight(dia, len), 1),
                                    ShapeType = "Straight",
                                    Layer = layer,
                                    Color = color,
                                    Direction = "Longitudinal",
                                    Vertices = [[cover, y, z], [cover + len, y, z]]
                                });
                            }
                        }
                    }
                    else if (grp.Equals("G2", StringComparison.OrdinalIgnoreCase) || item.Name.Contains("trên") || item.Name.Contains("đỉnh"))
                    {
                        // Thép chủ dọc trên
                        int layers = Math.Max(1, (int)Math.Ceiling(qty / 18.0));
                        int perLayer = (int)Math.Ceiling((double)qty / layers);
                        double stepZ = (depth - 2 * cover) / Math.Max(1, perLayer - 1);

                        int created = 0;
                        for (int l = 0; l < layers && created < qty; l++)
                        {
                            double y = height - cover - l * 65;
                            for (int i = 0; i < perLayer && created < qty; i++)
                            {
                                idx++;
                                created++;
                                double z = cover + i * stepZ;
                                bars.Add(new RebarBar
                                {
                                    Id = $"B{idx:D4}",
                                    GroupName = grp,
                                    Diameter = dia,
                                    Length = len,
                                    Weight = Math.Round(RebarUtils.CalcWeight(dia, len), 1),
                                    ShapeType = "Straight",
                                    Layer = layer,
                                    Color = color,
                                    Direction = "Longitudinal",
                                    Vertices = [[cover, y, z], [cover + len, y, z]]
                                });
                            }
                        }
                    }
                    else if (grp.Equals("G5", StringComparison.OrdinalIgnoreCase) || item.Name.Contains("sườn"))
                    {
                        // Thép sườn 2 mặt bên
                        int perSide = Math.Max(1, qty / 2);
                        double stepY = (height - 2 * cover) / (perSide + 1);

                        for (int s = 1; s <= perSide; s++)
                        {
                            double y = cover + s * stepY;
                            idx++;
                            bars.Add(new RebarBar
                            {
                                Id = $"B{idx:D4}",
                                GroupName = grp,
                                Diameter = dia,
                                Length = len,
                                Weight = Math.Round(RebarUtils.CalcWeight(dia, len), 1),
                                ShapeType = "Straight",
                                Layer = layer,
                                Color = color,
                                Direction = "Longitudinal",
                                Vertices = [[cover, y, cover + 15], [cover + len, y, cover + 15]]
                            });
                            idx++;
                            bars.Add(new RebarBar
                            {
                                Id = $"B{idx:D4}",
                                GroupName = grp,
                                Diameter = dia,
                                Length = len,
                                Weight = Math.Round(RebarUtils.CalcWeight(dia, len), 1),
                                ShapeType = "Straight",
                                Layer = layer,
                                Color = color,
                                Direction = "Longitudinal",
                                Vertices = [[cover, y, depth - cover - 15], [cover + len, y, depth - cover - 15]]
                            });
                        }
                    }
                    else if (grp.Equals("G6", StringComparison.OrdinalIgnoreCase) || item.Name.Contains("giằng"))
                    {
                        // Giằng ngang Z
                        double stepX = (width - 2 * cover) / Math.Max(1, qty - 1);
                        double yMid = height / 2.0;
                        for (int i = 0; i < qty; i++)
                        {
                            idx++;
                            double x = cover + i * stepX;
                            bars.Add(new RebarBar
                            {
                                Id = $"B{idx:D4}",
                                GroupName = grp,
                                Diameter = dia,
                                Length = len,
                                Weight = Math.Round(RebarUtils.CalcWeight(dia, len), 1),
                                ShapeType = "Straight",
                                Layer = layer,
                                Color = color,
                                Direction = "Transverse",
                                Vertices = [[x, yMid, cover], [x, yMid, depth - cover]]
                            });
                        }
                    }
                    else if (grp.Equals("G7", StringComparison.OrdinalIgnoreCase) || item.Name.Contains("xiên"))
                    {
                        // Thép xiên 2 vai xà mũ
                        int half = Math.Max(1, qty / 2);
                        for (int i = 0; i < half; i++)
                        {
                            double z = cover + (i + 0.5) * ((depth - 2 * cover) / half);
                            idx++;
                            bars.Add(new RebarBar
                            {
                                Id = $"B{idx:D4}",
                                GroupName = grp,
                                Diameter = dia,
                                Length = len,
                                Weight = Math.Round(RebarUtils.CalcWeight(dia, len), 1),
                                ShapeType = "Straight",
                                Layer = layer,
                                Color = color,
                                Direction = "Vertical",
                                Vertices = [[cover, height - cover - 200, z], [cover + 1800, cover + 200, z]]
                            });
                            idx++;
                            bars.Add(new RebarBar
                            {
                                Id = $"B{idx:D4}",
                                GroupName = grp,
                                Diameter = dia,
                                Length = len,
                                Weight = Math.Round(RebarUtils.CalcWeight(dia, len), 1),
                                ShapeType = "Straight",
                                Layer = layer,
                                Color = color,
                                Direction = "Vertical",
                                Vertices = [[width - cover, height - cover - 200, z], [width - cover - 1800, cover + 200, z]]
                            });
                        }
                    }
                    else if (grp.Equals("G8", StringComparison.OrdinalIgnoreCase) || item.Name.Contains("gối"))
                    {
                        // Lưới gia cường 8 bệ kê gối
                        int perPed = Math.Max(1, qty / 8);
                        double marginX = width * 0.12;
                        double stepPedX = (width - 2 * marginX) / 3;
                        double[] zPeds = [depth * 0.28, depth * 0.72];
                        double meshY = height - cover - 25;

                        int created = 0;
                        foreach (double zp in zPeds)
                        {
                            for (int c = 0; c < 4 && created < qty; c++)
                            {
                                double xp = marginX + c * stepPedX;
                                for (int k = 0; k < perPed && created < qty; k++)
                                {
                                    idx++;
                                    created++;
                                    double offset = (k - perPed / 2.0) * 60;
                                    bars.Add(new RebarBar
                                    {
                                        Id = $"B{idx:D4}",
                                        GroupName = grp,
                                        Diameter = dia,
                                        Length = len,
                                        Weight = Math.Round(RebarUtils.CalcWeight(dia, len), 1),
                                        ShapeType = "Straight",
                                        Layer = layer,
                                        Color = color,
                                        Direction = "Longitudinal",
                                        Vertices = [[xp - len / 2, meshY, zp + offset], [xp + len / 2, meshY, zp + offset]]
                                    });
                                }
                            }
                        }
                    }
                    else
                    {
                        // Mặc định phân bố đều
                        double stepX = (width - 2 * cover) / Math.Max(1, qty - 1);
                        for (int i = 0; i < qty; i++)
                        {
                            idx++;
                            double x = cover + i * stepX;
                            bars.Add(new RebarBar
                            {
                                Id = $"B{idx:D4}",
                                GroupName = grp,
                                Diameter = dia,
                                Length = len,
                                Weight = Math.Round(RebarUtils.CalcWeight(dia, len), 1),
                                ShapeType = "Straight",
                                Layer = layer,
                                Color = color,
                                Direction = "Longitudinal",
                                Vertices = [[x, height / 2.0, cover], [x, height / 2.0, depth - cover]]
                            });
                        }
                    }
                }
            }

            return bars;
        }

        /// <summary>
        /// Tạo dữ liệu mẫu nếu không tìm thấy thép trong CAD.
        /// Mô phỏng đầy đủ cốt thép xà mũ trụ cầu chuẩn kỹ thuật:
        /// Thép chủ dọc trên/dưới nhiều lớp, đai chính, đai phụ, thép sườn, thép xiên chống cắt, thép bệ gối.
        /// </summary>
        private static List<RebarBar> GenerateSampleData(double width, double height, double depth)
        {
            var bars = new List<RebarBar>();
            int idx = 0;
            double cover = 60; // Lớp bê tông bảo vệ 60mm

            // Chiều dài chủ đạo là trục X (VD: 11,700 mm)
            double mainLen = width - 2 * cover;

            // ===== 1. Thép dọc chủ dưới (G1) - D28 (2 lớp) =====
            int barsPerRowBot = Math.Max(6, (int)((depth - 2 * cover) / 130));
            double spacingBotZ = (depth - 2 * cover) / (barsPerRowBot - 1);
            for (int layer = 0; layer < 2; layer++)
            {
                double y = cover + layer * 70;
                for (int i = 0; i < barsPerRowBot; i++)
                {
                    idx++;
                    double z = cover + i * spacingBotZ;
                    bars.Add(new RebarBar
                    {
                        Id = $"B{idx:D4}",
                        GroupName = "G1",
                        Diameter = 28,
                        Length = Math.Round(mainLen, 1),
                        Weight = Math.Round(RebarUtils.CalcWeight(28, mainLen), 1),
                        ShapeType = "Straight",
                        Layer = "THEP_D28",
                        Color = [30, 144, 255], // Dodger Blue
                        Direction = "Longitudinal",
                        Vertices = [[cover, y, z], [width - cover, y, z]]
                    });
                }
            }

            // ===== 2. Thép dọc chủ trên (G2) - D28 (2 lớp) =====
            int barsPerRowTop = Math.Max(6, (int)((depth - 2 * cover) / 130));
            double spacingTopZ = (depth - 2 * cover) / (barsPerRowTop - 1);
            for (int layer = 0; layer < 2; layer++)
            {
                double y = height - cover - layer * 70;
                for (int i = 0; i < barsPerRowTop; i++)
                {
                    idx++;
                    double z = cover + i * spacingTopZ;
                    bars.Add(new RebarBar
                    {
                        Id = $"B{idx:D4}",
                        GroupName = "G2",
                        Diameter = 28,
                        Length = Math.Round(mainLen, 1),
                        Weight = Math.Round(RebarUtils.CalcWeight(28, mainLen), 1),
                        ShapeType = "Straight",
                        Layer = "THEP_D28",
                        Color = [0, 191, 255], // Deep Sky Blue
                        Direction = "Longitudinal",
                        Vertices = [[cover, y, z], [width - cover, y, z]]
                    });
                }
            }

            // ===== 3. Thép đai chính ngoài (G3) - D14 (Bao toàn tiết diện Y-Z, phân bố dọc X) =====
            int numStirrups = Math.Max(20, (int)(width / 180));
            double spacingStirrup = (width - 2 * cover) / (numStirrups - 1);
            double yMin = cover, yMax = height - cover;
            double zMin = cover, zMax = depth - cover;
            double stirrupLen = 2 * (yMax - yMin) + 2 * (zMax - zMin) + 240;

            for (int i = 0; i < numStirrups; i++)
            {
                idx++;
                double x = cover + i * spacingStirrup;
                bars.Add(new RebarBar
                {
                    Id = $"B{idx:D4}",
                    GroupName = "G3",
                    Diameter = 14,
                    Length = Math.Round(stirrupLen, 1),
                    Weight = Math.Round(RebarUtils.CalcWeight(14, stirrupLen), 1),
                    ShapeType = "Stirrup",
                    Layer = "THEP_D14_DAI",
                    Color = [34, 197, 94], // Emerald Green
                    Direction = "Transverse",
                    Vertices =
                    [
                        [x, yMin, zMin],
                        [x, yMax, zMin],
                        [x, yMax, zMax],
                        [x, yMin, zMax],
                        [x, yMin, zMin]
                    ]
                });
            }

            // ===== 4. Thép đai phụ trong (G4) - D12 (Đai đứng chia khoảng) =====
            for (int i = 0; i < numStirrups; i++)
            {
                idx++;
                double x = cover + i * spacingStirrup;
                double zMid1 = cover + (depth - 2 * cover) * 0.35;
                double zMid2 = cover + (depth - 2 * cover) * 0.65;
                double legLen = 2 * (yMax - yMin) + 2 * (zMid2 - zMid1) + 200;

                bars.Add(new RebarBar
                {
                    Id = $"B{idx:D4}",
                    GroupName = "G4",
                    Diameter = 12,
                    Length = Math.Round(legLen, 1),
                    Weight = Math.Round(RebarUtils.CalcWeight(12, legLen), 1),
                    ShapeType = "Stirrup",
                    Layer = "THEP_D12_DAI",
                    Color = [234, 179, 8], // Amber/Yellow
                    Direction = "Transverse",
                    Vertices =
                    [
                        [x, yMin, zMid1],
                        [x, yMax, zMid1],
                        [x, yMax, zMid2],
                        [x, yMin, zMid2],
                        [x, yMin, zMid1]
                    ]
                });
            }

            // ===== 5. Thép sườn chống nứt (G5) - D16 (Dọc theo hai mặt bên) =====
            int numSideBars = Math.Max(3, (int)((height - 2 * cover) / 220));
            double spacingSideY = (height - 2 * cover) / (numSideBars + 1);
            for (int s = 1; s <= numSideBars; s++)
            {
                double y = cover + s * spacingSideY;

                // Mặt trước (zMin)
                idx++;
                bars.Add(new RebarBar
                {
                    Id = $"B{idx:D4}",
                    GroupName = "G5",
                    Diameter = 16,
                    Length = Math.Round(mainLen, 1),
                    Weight = Math.Round(RebarUtils.CalcWeight(16, mainLen), 1),
                    ShapeType = "Straight",
                    Layer = "THEP_D16",
                    Color = [249, 115, 22], // Orange
                    Direction = "Longitudinal",
                    Vertices = [[cover, y, zMin + 15], [width - cover, y, zMin + 15]]
                });

                // Mặt sau (zMax)
                idx++;
                bars.Add(new RebarBar
                {
                    Id = $"B{idx:D4}",
                    GroupName = "G5",
                    Diameter = 16,
                    Length = Math.Round(mainLen, 1),
                    Weight = Math.Round(RebarUtils.CalcWeight(16, mainLen), 1),
                    ShapeType = "Straight",
                    Layer = "THEP_D16",
                    Color = [249, 115, 22], // Orange
                    Direction = "Longitudinal",
                    Vertices = [[cover, y, zMax - 15], [width - cover, y, zMax - 15]]
                });
            }

            // ===== 6. Thép đai giằng ngang (G6) - D12 (Móc ngang chống phình) =====
            for (int i = 0; i < numStirrups; i += 2)
            {
                double x = cover + i * spacingStirrup;
                for (int s = 1; s <= numSideBars; s++)
                {
                    idx++;
                    double y = cover + s * spacingSideY;
                    double tieLen = depth - 2 * cover + 150;
                    bars.Add(new RebarBar
                    {
                        Id = $"B{idx:D4}",
                        GroupName = "G6",
                        Diameter = 12,
                        Length = Math.Round(tieLen, 1),
                        Weight = Math.Round(RebarUtils.CalcWeight(12, tieLen), 1),
                        ShapeType = "Straight",
                        Layer = "THEP_D12",
                        Color = [168, 85, 247], // Purple
                        Direction = "Transverse",
                        Vertices = [[x, y, zMin], [x, y, zMax]]
                    });
                }
            }

            // ===== 7. Thép xiên tăng cường chống cắt tại vát vai xà mũ (G7) - D20 =====
            int numDiags = 8;
            for (int i = 0; i < numDiags; i++)
            {
                double z = cover + (i + 0.5) * ((depth - 2 * cover) / numDiags);

                // Vai trái
                idx++;
                bars.Add(new RebarBar
                {
                    Id = $"B{idx:D4}",
                    GroupName = "G7",
                    Diameter = 20,
                    Length = 2200,
                    Weight = Math.Round(RebarUtils.CalcWeight(20, 2200), 1),
                    ShapeType = "Straight",
                    Layer = "THEP_D20",
                    Color = [239, 68, 68], // Red
                    Direction = "Vertical",
                    Vertices = [[cover, height - cover - 200, z], [cover + 1800, cover + 200, z]]
                });

                // Vai phải
                idx++;
                bars.Add(new RebarBar
                {
                    Id = $"B{idx:D4}",
                    GroupName = "G7",
                    Diameter = 20,
                    Length = 2200,
                    Weight = Math.Round(RebarUtils.CalcWeight(20, 2200), 1),
                    ShapeType = "Straight",
                    Layer = "THEP_D20",
                    Color = [239, 68, 68], // Red
                    Direction = "Vertical",
                    Vertices = [[width - cover, height - cover - 200, z], [width - cover - 1800, cover + 200, z]]
                });
            }

            // ===== 8. Lưới thép gia cường dưới 8 đá kê gối (G8) - D14 =====
            int numPedestalsX = 4;
            double marginX = width * 0.12;
            double stepPedX = (width - 2 * marginX) / (numPedestalsX - 1);
            double[] zPedRows = [depth * 0.28, depth * 0.72];

            foreach (double zCenter in zPedRows)
            {
                for (int px = 0; px < numPedestalsX; px++)
                {
                    double xCenter = marginX + px * stepPedX;
                    double meshSize = 650;
                    double meshY = height - cover - 30;

                    // Thanh theo X
                    for (int m = -2; m <= 2; m++)
                    {
                        idx++;
                        double z = zCenter + m * 70;
                        bars.Add(new RebarBar
                        {
                            Id = $"B{idx:D4}",
                            GroupName = "G8",
                            Diameter = 14,
                            Length = meshSize,
                            Weight = Math.Round(RebarUtils.CalcWeight(14, meshSize), 1),
                            ShapeType = "Straight",
                            Layer = "THEP_D14",
                            Color = [244, 63, 94], // Rose / Pink
                            Direction = "Longitudinal",
                            Vertices = [[xCenter - meshSize / 2, meshY, z], [xCenter + meshSize / 2, meshY, z]]
                        });
                    }

                    // Thanh theo Z
                    for (int m = -2; m <= 2; m++)
                    {
                        idx++;
                        double x = xCenter + m * 70;
                        bars.Add(new RebarBar
                        {
                            Id = $"B{idx:D4}",
                            GroupName = "G8",
                            Diameter = 14,
                            Length = meshSize,
                            Weight = Math.Round(RebarUtils.CalcWeight(14, meshSize), 1),
                            ShapeType = "Straight",
                            Layer = "THEP_D14",
                            Color = [244, 63, 94],
                            Direction = "Transverse",
                            Vertices = [[x, meshY - 15, zCenter - meshSize / 2], [x, meshY - 15, zCenter + meshSize / 2]]
                        });
                    }
                }
            }

            // ===== 9. Thép chờ cột / cọc liên kết vào đáy xà mũ (G9) - D25 (Uốn neo L) =====
            int numPiles = 4;
            double pileStepX = (width - 2 * marginX) / (numPiles - 1);
            for (int p = 0; p < numPiles; p++)
            {
                double px = marginX + p * pileStepX;
                double pz = depth / 2.0;
                int barsPerPile = 12;

                for (int b = 0; b < barsPerPile; b++)
                {
                    idx++;
                    double angle = b * (2 * Math.PI / barsPerPile);
                    double radius = 450;
                    double bx = px + radius * Math.Cos(angle);
                    double bz = pz + radius * Math.Sin(angle);
                    double hookLen = 400;
                    double embedLen = height * 0.75;

                    bars.Add(new RebarBar
                    {
                        Id = $"B{idx:D4}",
                        GroupName = "G9",
                        Diameter = 25,
                        Length = Math.Round(embedLen + hookLen, 1),
                        Weight = Math.Round(RebarUtils.CalcWeight(25, embedLen + hookLen), 1),
                        ShapeType = "LBend",
                        Layer = "THEP_D25",
                        Color = [20, 184, 166], // Teal
                        Direction = "Vertical",
                        Vertices =
                        [
                            [bx + hookLen * Math.Cos(angle), cover, bz + hookLen * Math.Sin(angle)],
                            [bx, cover, bz],
                            [bx, cover + embedLen, bz]
                        ]
                    });
                }
            }

            return bars;
        }
    }
}
