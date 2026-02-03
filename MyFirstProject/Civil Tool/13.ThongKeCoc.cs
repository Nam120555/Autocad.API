using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using AcadDb = Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

using Civil = Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using ClosedXML.Excel;
using MyFirstProject.Extensions;

[assembly: CommandClass(typeof(Civil3DCsharp.ThongKeCocV2))]

namespace Civil3DCsharp
{
    public class StakeLocationInfo
    {
        public int STT { get; set; }
        public string Name { get; set; } = "";
        public string Station { get; set; } = "";
        public double StationValue { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Distance { get; set; }
    }

    public class ThongKeCocV2
    {
        [CommandMethod("CTSV_ThongKeCoc_ToaDo")]
        public static void CTSVThongKeCocToaDo()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                // 1. Lấy danh sách Alignment
                var alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông tìm thấy Alignment nào có SampleLineGroup!");
                    return;
                }

                // 2. Chọn Alignment
                string[] alignmentNames = alignments.Select(a => a.Name).ToArray();
                A.Ed.WriteMessage("\n=== CHỌN TUYẾN THỐNG KÊ TOA ĐỘ ===");
                for (int i = 0; i < alignmentNames.Length; i++)
                    A.Ed.WriteMessage($"\n  {i + 1}. {alignmentNames[i]}");

                PromptIntegerOptions pOpts = new("\nNhập số thứ tự tuyến: ")
                {
                    LowerLimit = 1,
                    UpperLimit = alignmentNames.Length,
                    DefaultValue = 1
                };
                PromptIntegerResult pRes = A.Ed.GetInteger(pOpts);
                if (pRes.Status != PromptStatus.OK) return;

                var selectedAlign = alignments[pRes.Value - 1];

                // 3. Trích xuất dữ liệu
                var stakeData = ExtractStakeData(tr, selectedAlign.AlignmentId, selectedAlign.SampleLineGroupId);
                if (stakeData.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông có dữ liệu cọc!");
                    return;
                }

                // 4. Hỏi lựa chọn xuất
                PromptKeywordOptions pko = new("\nChọn kiểu xuất: [Excel/CAD/Both] <Both>:", "Excel CAD Both");
                pko.Keywords.Default = "Both";
                PromptResult pr = A.Ed.GetKeywords(pko);
                if (pr.Status != PromptStatus.OK) return;

                string choice = pr.StringResult;

                if (choice == "Excel" || choice == "Both")
                {
                    ExportToExcel(selectedAlign.Name, stakeData);
                }

                if (choice == "CAD" || choice == "Both")
                {
                    ExportToCad(tr, selectedAlign.Name, stakeData);
                }

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nLỗi: {ex.Message}");
            }
        }

        private static List<(ObjectId AlignmentId, string Name, ObjectId SampleLineGroupId)> GetAlignmentsWithSampleLineGroups(Transaction tr)
        {
            var result = new List<(ObjectId, string, ObjectId)>();
            ObjectIdCollection ids = A.Cdoc.GetAlignmentIds();
            foreach (ObjectId id in ids)
            {
                if (tr.GetObject(id, OpenMode.ForRead) is Alignment alignment)
                {
                    ObjectIdCollection slgIds = alignment.GetSampleLineGroupIds();
                    if (slgIds.Count > 0)
                        result.Add((id, alignment.Name, slgIds[0]));
                }
            }
            return result;
        }

        private static List<StakeLocationInfo> ExtractStakeData(Transaction tr, ObjectId alignmentId, ObjectId slgId)
        {
            List<StakeLocationInfo> result = new();
            Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
            SampleLineGroup? slg = tr.GetObject(slgId, OpenMode.ForRead) as SampleLineGroup;
            if (alignment == null || slg == null) return result;

            // Lấy danh sách SampleLine và sắp xếp theo lý trình
            List<SampleLine> sortedSampleLines = new List<SampleLine>();
            foreach (ObjectId slId in slg.GetSampleLineIds())
            {
                if (tr.GetObject(slId, OpenMode.ForRead) is SampleLine sl)
                    sortedSampleLines.Add(sl);
            }
            sortedSampleLines = sortedSampleLines.OrderBy(s => s.Station).ToList();

            double prevStation = 0;
            int stt = 1;

            foreach (SampleLine sl in sortedSampleLines)
            {
                double x = 0, y = 0;
                alignment.PointLocation(sl.Station, 0, ref x, ref y);

                result.Add(new StakeLocationInfo
                {
                    STT = stt++,
                    Name = sl.Name,
                    Station = FormatStation(sl.Station),
                    StationValue = sl.Station,
                    X = Math.Round(x, 3),
                    Y = Math.Round(y, 3),
                    Distance = stt == 2 ? 0 : Math.Round(sl.Station - prevStation, 3)
                });
                prevStation = sl.Station;
            }
            return result;
        }

        private static string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double m = station % 1000;
            return $"Km{km}+{m:F3}";
        }

        private static void ExportToExcel(string alignmentName, List<StakeLocationInfo> data)
        {
            SaveFileDialog sfd = new()
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = $"ToaDoCoc_{alignmentName}.xlsx",
                Title = "Lưu file Excel thống kê tọa độ"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                using var workbook = new XLWorkbook();
                var ws = workbook.Worksheets.Add("Toa Do Coc");

                ws.Cell(1, 1).Value = "BẢNG THỐNG KÊ TỌA ĐỘ CỌC TRÊN TUYẾN";
                ws.Range(1, 1, 1, 6).Merge().Style.Font.SetBold().Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                ws.Cell(2, 1).Value = "Tuyến: " + alignmentName;

                string[] headers = { "STT", "TÊN CỌC", "LÝ TRÌNH", "X", "Y", "KHOẢNG CÁCH" };
                for (int i = 0; i < headers.Length; i++)
                {
                    var cell = ws.Cell(4, i + 1);
                    cell.Value = headers[i];
                    cell.Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.LightGray).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                }

                int row = 5;
                foreach (var item in data)
                {
                    ws.Cell(row, 1).Value = item.STT;
                    ws.Cell(row, 2).Value = item.Name;
                    ws.Cell(row, 3).Value = item.Station;
                    ws.Cell(row, 4).Value = item.X;
                    ws.Cell(row, 5).Value = item.Y;
                    ws.Cell(row, 6).Value = item.Distance;
                    row++;
                }

                ws.Range(4, 1, row - 1, 6).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin).Border.SetInsideBorder(XLBorderStyleValues.Thin);
                ws.Columns().AdjustToContents();
                workbook.SaveAs(sfd.FileName);
                A.Ed.WriteMessage($"\nĐã xuất Excel: {sfd.FileName}");
            }
        }

        private static void ExportToCad(Transaction tr, string alignmentName, List<StakeLocationInfo> data)
        {
            PromptPointOptions ppo = new("\nChọn điểm chèn bảng trong CAD:");
            PromptPointResult ppr = A.Ed.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK) return;

            Point3d insertPt = ppr.Value;
            
            BlockTable bt = (BlockTable)tr.GetObject(A.Db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            AcadDb.Table tb = new();
            tb.TableStyle = A.Db.Tablestyle;
            tb.SetSize(data.Count + 3, 6); // Title + Header + Data
            tb.Position = insertPt;

            // Title
            tb.Cells[0, 0].TextString = "BẢNG TỌA ĐỘ CỌC - " + alignmentName;
            
            // Header
            string[] headers = { "STT", "TÊN CỌC", "LÝ TRÌNH", "X", "Y", "KC" };
            for (int i = 0; i < headers.Length; i++)
            {
                tb.Cells[1, i].TextString = headers[i];
                tb.Cells[1, i].Alignment = CellAlignment.MiddleCenter;
            }

            // Data
            for (int i = 0; i < data.Count; i++)
            {
                int r = i + 2;
                tb.Cells[r, 0].TextString = data[i].STT.ToString();
                tb.Cells[r, 1].TextString = data[i].Name;
                tb.Cells[r, 2].TextString = data[i].Station;
                tb.Cells[r, 3].TextString = data[i].X.ToString("F3");
                tb.Cells[r, 4].TextString = data[i].Y.ToString("F3");
                tb.Cells[r, 5].TextString = data[i].Distance.ToString("F3");

                for (int c = 0; c < 6; c++)
                    tb.Cells[r, c].Alignment = CellAlignment.MiddleCenter;
            }

            tb.GenerateLayout();
            btr.AppendEntity(tb);
            tr.AddNewlyCreatedDBObject(tb, true);
            A.Ed.WriteMessage("\nĐã vẽ bảng Table trong CAD.");
        }
    }
}
