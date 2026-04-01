// 24.PipeNetworkTools.cs - Công cụ Pipe Network từ VisualINFRA
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

[assembly: CommandClass(typeof(MyFirstProject.VisualINFRA.PipeNetworkTools))]

namespace MyFirstProject.VisualINFRA
{
    /// <summary>
    /// Data class chứa thông tin Pipe
    /// </summary>
    public class PipeInfo
    {
        public string Name { get; set; } = "";
        public string PartFamily { get; set; } = "";
        public double InnerDiameter { get; set; }
        public double OuterDiameter { get; set; }
        public double Length { get; set; }
        public double Slope { get; set; }
        public double StartStation { get; set; }
        public double EndStation { get; set; }
        public double StartInvert { get; set; }
        public double EndInvert { get; set; }
    }

    /// <summary>
    /// Data class chứa thông tin Structure (Hố ga)
    /// </summary>
    public class StructureInfo
    {
        public string Name { get; set; } = "";
        public string PartFamily { get; set; } = "";
        public double X { get; set; }
        public double Y { get; set; }
        public double RimElevation { get; set; }
        public double SumpElevation { get; set; }
        public double Depth { get; set; }
        public int ConnectedPipes { get; set; }
    }

    /// <summary>
    /// Công cụ Pipe Network - từ VisualINFRA
    /// </summary>
    public class PipeNetworkTools
    {
        #region Manhole Coordinate - Tọa độ hố ga

        /// <summary>
        /// Xuất tọa độ hố ga
        /// </summary>
        [CommandMethod("VI_ManholeCoordinate")]
        public static void ManholeCoordinate()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            var db = doc.Database;

            try
            {
                // Lấy danh sách Pipe Network
                var networkId = SelectPipeNetwork(ed, civilDoc, db);
                if (networkId.IsNull) return;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var network = tr.GetObject(networkId, OpenMode.ForRead) as Network;
                    if (network == null)
                    {
                        tr.Commit();
                        return;
                    }

                    var structureIds = network.GetStructureIds();
                    var structures = new List<StructureInfo>();

                    foreach (ObjectId strId in structureIds)
                    {
                        var structure = tr.GetObject(strId, OpenMode.ForRead) as Structure;
                        if (structure == null) continue;

                        var info = new StructureInfo
                        {
                            Name = structure.Name,
                            X = structure.Position.X,
                            Y = structure.Position.Y,
                            RimElevation = structure.RimElevation,
                            SumpElevation = structure.SumpElevation,
                            Depth = structure.RimElevation - structure.SumpElevation,
                            ConnectedPipes = structure.ConnectedPipesCount
                        };

                        try
                        {
                            info.PartFamily = structure.PartFamilyName;
                        }
                        catch { }

                        structures.Add(info);
                    }

                    // Hiển thị kết quả
                    ed.WriteMessage($"\n\n{'=',-90}");
                    ed.WriteMessage($"\n TỌA ĐỘ HỐ GA - {network.Name}");
                    ed.WriteMessage($"\n{'=',-90}");
                    ed.WriteMessage($"\n{"STT",-5} {"Tên",-15} {"X",-15} {"Y",-15} {"Rim",-12} {"Sump",-12} {"Depth",-10}");
                    ed.WriteMessage($"\n{new string('-', 90)}");

                    var sb = new StringBuilder();
                    sb.AppendLine("STT\tTen\tX\tY\tRimElevation\tSumpElevation\tDepth\tConnectedPipes");

                    int stt = 1;
                    foreach (var s in structures.OrderBy(x => x.Name))
                    {
                        ed.WriteMessage($"\n{stt,-5} {s.Name,-15} {s.X,-15:F3} {s.Y,-15:F3} {s.RimElevation,-12:F3} {s.SumpElevation,-12:F3} {s.Depth,-10:F3}");
                        sb.AppendLine($"{stt}\t{s.Name}\t{s.X:F3}\t{s.Y:F3}\t{s.RimElevation:F3}\t{s.SumpElevation:F3}\t{s.Depth:F3}\t{s.ConnectedPipes}");
                        stt++;
                    }

                    ed.WriteMessage($"\n{'=',-90}");
                    ed.WriteMessage($"\n Tổng: {structures.Count} hố ga. Đã copy vào clipboard!");

                    try
                    {
                        System.Windows.Forms.Clipboard.SetText(sb.ToString());
                    }
                    catch { }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Pipe Info - Thông tin Pipe

        /// <summary>
        /// Hiển thị thông tin Pipe Network
        /// </summary>
        [CommandMethod("VI_PipeInfo")]
        public static void PipeInfo()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            var db = doc.Database;

            try
            {
                var networkId = SelectPipeNetwork(ed, civilDoc, db);
                if (networkId.IsNull) return;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var network = tr.GetObject(networkId, OpenMode.ForRead) as Network;
                    if (network == null)
                    {
                        tr.Commit();
                        return;
                    }

                    var pipeIds = network.GetPipeIds();
                    var pipes = new List<PipeInfo>();

                    foreach (ObjectId pipeId in pipeIds)
                    {
                        var pipe = tr.GetObject(pipeId, OpenMode.ForRead) as Pipe;
                        if (pipe == null) continue;

                        var info = new PipeInfo
                        {
                            Name = pipe.Name,
                            InnerDiameter = pipe.InnerDiameterOrWidth,
                            OuterDiameter = pipe.OuterDiameterOrWidth,
                            Length = pipe.Length2D,
                            Slope = pipe.Slope,
                            StartStation = pipe.StartStation,
                            EndStation = pipe.EndStation,
                            StartInvert = pipe.StartPoint.Z,
                            EndInvert = pipe.EndPoint.Z
                        };

                        try
                        {
                            info.PartFamily = pipe.PartFamilyName;
                        }
                        catch { }

                        pipes.Add(info);
                    }

                    // Hiển thị kết quả
                    ed.WriteMessage($"\n\n{'=',-100}");
                    ed.WriteMessage($"\n THÔNG TIN PIPE - {network.Name}");
                    ed.WriteMessage($"\n{'=',-100}");
                    ed.WriteMessage($"\n{"STT",-5} {"Tên",-15} {"Đường kính",-12} {"Chiều dài",-12} {"Độ dốc (%)",-12} {"Z đầu",-10} {"Z cuối",-10}");
                    ed.WriteMessage($"\n{new string('-', 100)}");

                    int stt = 1;
                    double totalLength = 0;

                    foreach (var p in pipes.OrderBy(x => x.Name))
                    {
                        ed.WriteMessage($"\n{stt,-5} {p.Name,-15} {p.InnerDiameter * 1000,-12:F0}mm {p.Length,-12:F2}m {p.Slope * 100,-12:F2} {p.StartInvert,-10:F3} {p.EndInvert,-10:F3}");
                        totalLength += p.Length;
                        stt++;
                    }

                    ed.WriteMessage($"\n{new string('-', 100)}");
                    ed.WriteMessage($"\n{"TỔNG",-20} {"",-12} {totalLength,-12:F2}m");
                    ed.WriteMessage($"\n{'=',-100}");

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Export Pipe Network to CSV

        /// <summary>
        /// Xuất Pipe Network ra file CSV
        /// </summary>
        [CommandMethod("VI_ExportPipeNetwork")]
        public static void ExportPipeNetwork()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            var db = doc.Database;

            try
            {
                var networkId = SelectPipeNetwork(ed, civilDoc, db);
                if (networkId.IsNull) return;

                // Chọn thư mục lưu
                var fbd = new System.Windows.Forms.FolderBrowserDialog
                {
                    Description = "Chọn thư mục lưu file"
                };

                if (fbd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    return;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var network = tr.GetObject(networkId, OpenMode.ForRead) as Network;
                    if (network == null)
                    {
                        tr.Commit();
                        return;
                    }

                    // Xuất Pipes
                    var sbPipes = new StringBuilder();
                    sbPipes.AppendLine("STT,Name,PartFamily,InnerDiameter,OuterDiameter,Length,Slope,StartInvert,EndInvert");

                    var pipeIds = network.GetPipeIds();
                    int stt = 1;
                    foreach (ObjectId pipeId in pipeIds)
                    {
                        var pipe = tr.GetObject(pipeId, OpenMode.ForRead) as Pipe;
                        if (pipe == null) continue;

                        string partFamily = "";
                        try { partFamily = pipe.PartFamilyName; } catch { }

                        sbPipes.AppendLine($"{stt},{pipe.Name},{partFamily},{pipe.InnerDiameterOrWidth:F3},{pipe.OuterDiameterOrWidth:F3},{pipe.Length2D:F3},{pipe.Slope:F4},{pipe.StartPoint.Z:F3},{pipe.EndPoint.Z:F3}");
                        stt++;
                    }

                    string pipeFile = Path.Combine(fbd.SelectedPath, $"{network.Name}_Pipes.csv");
                    File.WriteAllText(pipeFile, sbPipes.ToString(), Encoding.UTF8);

                    // Xuất Structures
                    var sbStructures = new StringBuilder();
                    sbStructures.AppendLine("STT,Name,PartFamily,X,Y,RimElevation,SumpElevation,Depth,ConnectedPipes");

                    var structureIds = network.GetStructureIds();
                    stt = 1;
                    foreach (ObjectId strId in structureIds)
                    {
                        var structure = tr.GetObject(strId, OpenMode.ForRead) as Structure;
                        if (structure == null) continue;

                        string partFamily = "";
                        try { partFamily = structure.PartFamilyName; } catch { }

                        sbStructures.AppendLine($"{stt},{structure.Name},{partFamily},{structure.Position.X:F3},{structure.Position.Y:F3},{structure.RimElevation:F3},{structure.SumpElevation:F3},{structure.RimElevation - structure.SumpElevation:F3},{structure.ConnectedPipesCount}");
                        stt++;
                    }

                    string structureFile = Path.Combine(fbd.SelectedPath, $"{network.Name}_Structures.csv");
                    File.WriteAllText(structureFile, sbStructures.ToString(), Encoding.UTF8);

                    ed.WriteMessage($"\n Đã xuất Pipe Network ra:");
                    ed.WriteMessage($"\n   - {pipeFile}");
                    ed.WriteMessage($"\n   - {structureFile}");

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Change Pipe Elevation

        /// <summary>
        /// Thay đổi cao độ Pipe
        /// </summary>
        [CommandMethod("VI_ChangePipeElevation")]
        public static void ChangePipeElevation()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // Chọn Pipe
                var pipeResult = ed.GetEntity("\nChọn Pipe cần thay đổi cao độ: ");
                if (pipeResult.Status != PromptStatus.OK) return;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var pipe = tr.GetObject(pipeResult.ObjectId, OpenMode.ForWrite) as Pipe;
                    if (pipe == null)
                    {
                        ed.WriteMessage("\n Đối tượng không phải Pipe.");
                        tr.Commit();
                        return;
                    }

                    ed.WriteMessage($"\n Thông tin Pipe: {pipe.Name}");
                    ed.WriteMessage($"\n   Cao độ đầu: {pipe.StartPoint.Z:F3}");
                    ed.WriteMessage($"\n   Cao độ cuối: {pipe.EndPoint.Z:F3}");

                    // Nhập giá trị dịch chuyển
                    var deltaResult = ed.GetDouble("\nNhập giá trị dịch chuyển cao độ (m, + lên, - xuống): ");
                    if (deltaResult.Status != PromptStatus.OK)
                    {
                        tr.Commit();
                        return;
                    }

                    double delta = deltaResult.Value;

                    // Thay đổi cao độ
                    var startPt = pipe.StartPoint;
                    var endPt = pipe.EndPoint;

                    pipe.StartPoint = new Point3d(startPt.X, startPt.Y, startPt.Z + delta);
                    pipe.EndPoint = new Point3d(endPt.X, endPt.Y, endPt.Z + delta);

                    ed.WriteMessage($"\n Đã dịch chuyển Pipe {delta:F3}m");
                    ed.WriteMessage($"\n   Cao độ đầu mới: {pipe.StartPoint.Z:F3}");
                    ed.WriteMessage($"\n   Cao độ cuối mới: {pipe.EndPoint.Z:F3}");

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Change Structure Elevation

        /// <summary>
        /// Thay đổi cao độ Structure (Hố ga)
        /// </summary>
        [CommandMethod("VI_ChangeStructureElevation")]
        public static void ChangeStructureElevation()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // Chọn Structure
                var strResult = ed.GetEntity("\nChọn Structure cần thay đổi cao độ: ");
                if (strResult.Status != PromptStatus.OK) return;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var structure = tr.GetObject(strResult.ObjectId, OpenMode.ForWrite) as Structure;
                    if (structure == null)
                    {
                        ed.WriteMessage("\n Đối tượng không phải Structure.");
                        tr.Commit();
                        return;
                    }

                    ed.WriteMessage($"\n Thông tin Structure: {structure.Name}");
                    ed.WriteMessage($"\n   Rim Elevation: {structure.RimElevation:F3}");
                    ed.WriteMessage($"\n   Sump Elevation: {structure.SumpElevation:F3}");

                    // Chọn loại thay đổi
                    var optResult = ed.GetKeywords("\nThay đổi [Rim/Sump/Both] <Rim>: ", new[] { "Rim", "Sump", "Both" });
                    string option = optResult.Status == PromptStatus.OK ? optResult.StringResult : "Rim";

                    var deltaResult = ed.GetDouble("\nNhập giá trị dịch chuyển (m): ");
                    if (deltaResult.Status != PromptStatus.OK)
                    {
                        tr.Commit();
                        return;
                    }

                    double delta = deltaResult.Value;

                    switch (option.ToUpper())
                    {
                        case "RIM":
                            structure.RimElevation += delta;
                            break;
                        case "SUMP":
                            structure.SumpElevation += delta;
                            break;
                        case "BOTH":
                            structure.RimElevation += delta;
                            structure.SumpElevation += delta;
                            break;
                    }

                    ed.WriteMessage($"\n Đã thay đổi cao độ Structure.");
                    ed.WriteMessage($"\n   Rim Elevation mới: {structure.RimElevation:F3}");
                    ed.WriteMessage($"\n   Sump Elevation mới: {structure.SumpElevation:F3}");

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Pipe Network Summary

        /// <summary>
        /// Tổng hợp thống kê Pipe Network
        /// </summary>
        [CommandMethod("VI_PipeNetworkSummary")]
        public static void PipeNetworkSummary()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            var db = doc.Database;

            try
            {
                var networkIds = civilDoc.GetPipeNetworkIds();
                
                ed.WriteMessage($"\n\n{'=',-80}");
                ed.WriteMessage($"\n TỔNG HỢP PIPE NETWORK");
                ed.WriteMessage($"\n{'=',-80}");
                ed.WriteMessage($"\n{"Tên Network",-30} {"Số Pipe",-12} {"Tổng dài (m)",-15} {"Số Structure",-12}");
                ed.WriteMessage($"\n{new string('-', 80)}");

                int totalPipes = 0, totalStructures = 0;
                double totalLength = 0;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId netId in networkIds)
                    {
                        var network = tr.GetObject(netId, OpenMode.ForRead) as Network;
                        if (network == null) continue;

                        var pipeIds = network.GetPipeIds();
                        var structureIds = network.GetStructureIds();

                        double netLength = 0;
                        foreach (ObjectId pipeId in pipeIds)
                        {
                            var pipe = tr.GetObject(pipeId, OpenMode.ForRead) as Pipe;
                            if (pipe != null)
                                netLength += pipe.Length2D;
                        }

                        ed.WriteMessage($"\n{network.Name,-30} {pipeIds.Count,-12} {netLength,-15:F2} {structureIds.Count,-12}");

                        totalPipes += pipeIds.Count;
                        totalStructures += structureIds.Count;
                        totalLength += netLength;
                    }

                    ed.WriteMessage($"\n{new string('-', 80)}");
                    ed.WriteMessage($"\n{"TỔNG CỘNG",-30} {totalPipes,-12} {totalLength,-15:F2} {totalStructures,-12}");
                    ed.WriteMessage($"\n{'=',-80}");

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Helper Methods

        private static ObjectId SelectPipeNetwork(Editor ed, CivilDocument civilDoc, Database db)
        {
            var networkIds = civilDoc.GetPipeNetworkIds();
            if (networkIds.Count == 0)
            {
                ed.WriteMessage("\n Không có Pipe Network.");
                return ObjectId.Null;
            }

            ed.WriteMessage("\n\nDanh sách Pipe Network:");
            var networks = new List<(int Index, ObjectId Id, string Name)>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                int idx = 1;
                foreach (ObjectId id in networkIds)
                {
                    var network = tr.GetObject(id, OpenMode.ForRead) as Network;
                    if (network != null)
                    {
                        ed.WriteMessage($"\n  {idx}. {network.Name}");
                        networks.Add((idx, id, network.Name));
                        idx++;
                    }
                }
                tr.Commit();
            }

            var result = ed.GetInteger($"\nChọn Pipe Network (1-{networks.Count}): ");
            if (result.Status != PromptStatus.OK) return ObjectId.Null;
            if (result.Value < 1 || result.Value > networks.Count) return ObjectId.Null;

            return networks[result.Value - 1].Id;
        }

        #endregion
    }
}
