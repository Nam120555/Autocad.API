// 21.SampleLineTools.cs - Công cụ Sample Line từ VisualINFRA
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

[assembly: CommandClass(typeof(MyFirstProject.VisualINFRA.SampleLineTools))]

namespace MyFirstProject.VisualINFRA
{
    /// <summary>
    /// Công cụ làm việc với Sample Line - từ VisualINFRA
    /// </summary>
    public class SampleLineTools
    {
        #region Sample Line Coordinate - Xuất tọa độ cọc

        /// <summary>
        /// Xuất tọa độ cọc ra Command Line và Clipboard
        /// </summary>
        [CommandMethod("VI_SampleLineCoordinate")]
        public static void SampleLineCoordinate()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            
            try
            {
                // Chọn Alignment
                var alignmentId = SelectAlignment(ed, civilDoc);
                if (alignmentId.IsNull)
                {
                    ed.WriteMessage("\n❌ Không chọn được Alignment.");
                    return;
                }

                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                    if (alignment == null)
                    {
                        ed.WriteMessage("\n❌ Không đọc được Alignment.");
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
                    
                    ed.WriteMessage($"\n\n{'=',-60}");
                    ed.WriteMessage($"\n📍 TỌA ĐỘ CỌC - {alignment.Name}");
                    ed.WriteMessage($"\n{'=',-60}");
                    ed.WriteMessage($"\n{"STT",-5} {"Tên Cọc",-15} {"Lý Trình",-15} {"X",-15} {"Y",-15}");
                    ed.WriteMessage($"\n{new string('-', 65)}");

                    var sb = new StringBuilder();
                    sb.AppendLine("STT\tTên Cọc\tLý Trình\tX\tY");

                    int stt = 1;
                    foreach (ObjectId slId in slIds)
                    {
                        var sl = tr.GetObject(slId, OpenMode.ForRead) as SampleLine;
                        if (sl == null) continue;

                        // Lấy tọa độ tại tim tuyến
                        double station = sl.Station;
                        double x = 0, y = 0;
                        
                        try
                        {
                            alignment.PointLocation(station, 0, ref x, ref y);
                        }
                        catch
                        {
                            continue;
                        }

                        string lyTrinh = FormatStation(station);
                        
                        ed.WriteMessage($"\n{stt,-5} {sl.Name,-15} {lyTrinh,-15} {x,-15:F3} {y,-15:F3}");
                        sb.AppendLine($"{stt}\t{sl.Name}\t{lyTrinh}\t{x:F3}\t{y:F3}");
                        
                        stt++;
                    }

                    ed.WriteMessage($"\n{new string('=', 65)}");
                    ed.WriteMessage($"\n✅ Tổng: {stt - 1} cọc. Đã copy vào clipboard!");

                    // Copy to clipboard
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
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Export Sample Line to CSV

        /// <summary>
        /// Xuất Sample Line ra file CSV
        /// </summary>
        [CommandMethod("VI_ExportSampleLine")]
        public static void ExportSampleLine()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            
            try
            {
                var alignmentId = SelectAlignment(ed, civilDoc);
                if (alignmentId.IsNull) return;

                // Chọn đường dẫn lưu file
                var sfd = new System.Windows.Forms.SaveFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                    Title = "Lưu file Sample Line",
                    FileName = "SampleLine_Export.csv"
                };

                if (sfd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    return;

                using (var tr = doc.Database.TransactionManager.StartTransaction())
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

                    var sb = new StringBuilder();
                    sb.AppendLine("STT,TenCoc,LyTrinh,Station,X,Y,OffsetTrai,OffsetPhai");

                    var slg = tr.GetObject(slgIds[0], OpenMode.ForRead) as SampleLineGroup;
                    var slIds = slg!.GetSampleLineIds();

                    int stt = 1;
                    foreach (ObjectId slId in slIds)
                    {
                        var sl = tr.GetObject(slId, OpenMode.ForRead) as SampleLine;
                        if (sl == null) continue;

                        double x = 0, y = 0;
                        try
                        {
                            alignment.PointLocation(sl.Station, 0, ref x, ref y);
                        }
                        catch { continue; }

                        // Lấy offset trái phải - sử dụng giá trị mặc định
                        // (API SampleLine trong Civil 3D 2026 không có SwathWidth properties)
                        double leftOffset = -20, rightOffset = 20;

                        sb.AppendLine($"{stt},{sl.Name},{FormatStation(sl.Station)},{sl.Station:F3},{x:F3},{y:F3},{leftOffset:F2},{rightOffset:F2}");
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

        #region Import Sample Line from CSV

        /// <summary>
        /// Nhập Sample Line từ file CSV
        /// </summary>
        [CommandMethod("VI_ImportSampleLine")]
        public static void ImportSampleLine()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            
            try
            {
                var alignmentId = SelectAlignment(ed, civilDoc);
                if (alignmentId.IsNull) return;

                // Chọn file CSV
                var ofd = new System.Windows.Forms.OpenFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|All files (*.*)|*.*",
                    Title = "Chọn file Sample Line để nhập"
                };

                if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    return;

                var lines = File.ReadAllLines(ofd.FileName, Encoding.UTF8);
                if (lines.Length < 2)
                {
                    ed.WriteMessage("\n❌ File rỗng hoặc không đúng định dạng.");
                    return;
                }

                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                    if (alignment == null)
                    {
                        tr.Commit();
                        return;
                    }

                    // Lấy hoặc tạo Sample Line Group
                    var slgIds = alignment.GetSampleLineGroupIds();
                    SampleLineGroup? slg = null;
                    
                    if (slgIds.Count > 0)
                    {
                        slg = tr.GetObject(slgIds[0], OpenMode.ForWrite) as SampleLineGroup;
                    }
                    else
                    {
                        // Tạo mới Sample Line Group
                        var slgId = SampleLineGroup.Create("SLG_Import", alignmentId);
                        slg = tr.GetObject(slgId, OpenMode.ForWrite) as SampleLineGroup;
                    }

                    if (slg == null)
                    {
                        ed.WriteMessage("\n❌ Không thể tạo Sample Line Group.");
                        tr.Commit();
                        return;
                    }

                    int created = 0;
                    for (int i = 1; i < lines.Length; i++) // Bỏ qua header
                    {
                        var parts = lines[i].Split(',');
                        if (parts.Length < 4) continue;

                        // Parse: STT, TenCoc, LyTrinh, Station, ...
                        string tenCoc = parts[1].Trim();
                        if (!double.TryParse(parts[3], out double station))
                            continue;

                        double leftOffset = -20, rightOffset = 20;
                        if (parts.Length >= 7)
                        {
                            double.TryParse(parts[6], out leftOffset);
                            double.TryParse(parts[7], out rightOffset);
                        }

                        try
                        {
                            var slId = SampleLine.Create(tenCoc, slg.ObjectId, station);
                            var sl = tr.GetObject(slId, OpenMode.ForWrite) as SampleLine;
                            
                            // Cập nhật offset nếu cần
                            // (SampleLine API hạn chế, cần thêm code custom)
                            
                            created++;
                        }
                        catch
                        {
                            // Sample line tại vị trí này có thể đã tồn tại
                        }
                    }

                    ed.WriteMessage($"\n✅ Đã tạo {created} Sample Line từ file.");
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Rename Sample Line

        /// <summary>
        /// Đổi tên Sample Line theo pattern
        /// </summary>
        [CommandMethod("VI_RenameSampleLine")]
        public static void RenameSampleLine()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            
            try
            {
                var alignmentId = SelectAlignment(ed, civilDoc);
                if (alignmentId.IsNull) return;

                // Nhập prefix
                var prefixResult = ed.GetString("\nNhập prefix (VD: Km0+): ");
                if (prefixResult.Status != PromptStatus.OK) return;
                string prefix = prefixResult.StringResult;

                using (var tr = doc.Database.TransactionManager.StartTransaction())
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

                    // Sắp xếp theo station
                    var sortedList = new List<(ObjectId Id, double Station)>();
                    foreach (ObjectId slId in slIds)
                    {
                        var sl = tr.GetObject(slId, OpenMode.ForRead) as SampleLine;
                        if (sl != null)
                            sortedList.Add((slId, sl.Station));
                    }
                    sortedList = sortedList.OrderBy(x => x.Station).ToList();

                    int renamed = 0;
                    foreach (var item in sortedList)
                    {
                        var sl = tr.GetObject(item.Id, OpenMode.ForWrite) as SampleLine;
                        if (sl != null)
                        {
                            string newName = $"{prefix}{FormatStation(item.Station)}";
                            sl.Name = newName;
                            renamed++;
                        }
                    }

                    ed.WriteMessage($"\n✅ Đã đổi tên {renamed} Sample Line.");
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        #endregion

        #region Filling Sample Line - Điền thông tin

        /// <summary>
        /// Điền text thông tin lên Sample Line
        /// </summary>
        [CommandMethod("VI_FillingSampleLine")]
        public static void FillingSampleLine()
        {
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;
            
            try
            {
                var alignmentId = SelectAlignment(ed, civilDoc);
                if (alignmentId.IsNull) return;

                // Chọn layer để vẽ text
                var layerResult = ed.GetString("\nNhập tên Layer cho text (Enter = VI_SL_TEXT): ");
                string layerName = string.IsNullOrEmpty(layerResult.StringResult) ? "VI_SL_TEXT" : layerResult.StringResult;

                // Tạo layer
                VIFunc.CreateLayer(layerName, 3); // Màu xanh lá

                // Nhập chiều cao text
                var heightResult = ed.GetDouble("\nChiều cao text (2.5): ");
                double textHeight = heightResult.Status == PromptStatus.OK ? heightResult.Value : 2.5;

                using (var tr = doc.Database.TransactionManager.StartTransaction())
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

                    int added = 0;
                    foreach (ObjectId slId in slIds)
                    {
                        var sl = tr.GetObject(slId, OpenMode.ForRead) as SampleLine;
                        if (sl == null) continue;

                        double x = 0, y = 0;
                        try
                        {
                            alignment.PointLocation(sl.Station, 0, ref x, ref y);
                        }
                        catch { continue; }

                        // Tạo text
                        var pos = new Point3d(x, y + textHeight * 2, 0);
                        VIFunc.AddText(sl.Name, pos, textHeight, 0, layerName);
                        added++;
                    }

                    ed.WriteMessage($"\n✅ Đã điền text cho {added} cọc.");
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

        /// <summary>
        /// Chọn Alignment
        /// </summary>
        private static ObjectId SelectAlignment(Editor ed, CivilDocument civilDoc)
        {
            var alignmentIds = civilDoc.GetAlignmentIds();
            if (alignmentIds.Count == 0)
            {
                ed.WriteMessage("\n❌ Không có Alignment trong bản vẽ.");
                return ObjectId.Null;
            }

            ed.WriteMessage("\n\nDanh sách Alignment:");
            var doc = AcadApplication.DocumentManager.MdiActiveDocument;
            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                int index = 1;
                var alignments = new List<(int Index, ObjectId Id, string Name)>();
                
                foreach (ObjectId id in alignmentIds)
                {
                    var alignment = tr.GetObject(id, OpenMode.ForRead) as Alignment;
                    if (alignment != null)
                    {
                        ed.WriteMessage($"\n  {index}. {alignment.Name}");
                        alignments.Add((index, id, alignment.Name));
                        index++;
                    }
                }

                tr.Commit();

                var result = ed.GetInteger($"\nChọn Alignment (1-{alignments.Count}): ");
                if (result.Status != PromptStatus.OK) return ObjectId.Null;

                int selected = result.Value;
                if (selected < 1 || selected > alignments.Count) return ObjectId.Null;

                return alignments[selected - 1].Id;
            }
        }

        /// <summary>
        /// Format station thành dạng Km+m
        /// </summary>
        private static string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double m = station % 1000;
            return $"Km{km}+{m:F2}";
        }

        #endregion
    }
}
