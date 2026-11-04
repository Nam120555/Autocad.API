using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

// ╔══════════════════════════════════════════════════════════════════════════╗
// ║              BR TOOL  —  CÔNG CỤ HỖ TRỢ THIẾT KẾ CẦU                  ║
// ║              Tác giả: HUỲNH NGỌC NAM  |  Kỹ sư Giao thông              ║
// ╠════════════╦═════════════════════════════════════════════════════════════╣
// ║  LỆNH      ║  MÔ TẢ & CÁCH SỬ DỤNG                                     ║
// ╠════════════╬═════════════════════════════════════════════════════════════╣
// ║  XTB       ║  Xuất tọa độ Block + Lý trình → File Excel (.xlsx)         ║
// ║            ║  1. Chọn đường hướng tuyến (Line / Polyline / Spline)      ║
// ║            ║  2. Quét chọn các Block cần xuất                           ║
// ║            ║  3. Nhập số thứ tự bắt đầu (Enter = không đánh số)         ║
// ║            ║  → Xuất: STT | Tên trụ | X | Y | Z | Góc | LT | Block     ║
// ╠════════════╬═════════════════════════════════════════════════════════════╣
// ║  TAONHAN   ║  Tạo Block Attribute nhãn 2 dòng tại vị trí từng Block     ║
// ║            ║  dọc theo hướng tuyến (sắp xếp theo lý trình tăng dần)     ║
// ║            ║  ┌──────────┐                                               ║
// ║            ║  │   001    │  ← Dòng TRÊN: Số thứ tự (001, 002...)        ║
// ║            ║  │ [Block]  │  ← Vị trí chèn                               ║
// ║            ║  │ P KM0+500│  ← Dòng DƯỚI: Lý trình theo tuyến           ║
// ║            ║  └──────────┘                                               ║
// ║            ║  Block Definition tự tạo: "BR_NHAN_STT"                    ║
// ║            ║  Tag: STT (trên) | TENTRU (dưới)                           ║
// ╚════════════╩═════════════════════════════════════════════════════════════╝


namespace MyFirstProject.BrTool
{
    public class BlockToStationExporter
    {
        private class BlockExportData
        {
            public string STT { get; set; } = string.Empty;
            public Point3d Position { get; set; }
            public double RotationDeg { get; set; }
            public double Station { get; set; }
            public string BlockName { get; set; } = string.Empty;
            public string TenTru { get; set; } = string.Empty;
            public string Layer { get; set; } = "0";
        }

        [CommandMethod("XTB")]
        public void ExportCoordinatesAndStation()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 1. Chọn đường hướng tuyến
            PromptEntityOptions peo = new PromptEntityOptions("\nChọn đường hướng tuyến (Line, Polyline, Spline): ");
            peo.SetRejectMessage("\nChỉ chọn đối tượng đường (Line, Polyline, Spline).");
            peo.AddAllowedClass(typeof(Curve), false); // Curve bao phủ tất cả đối tượng tuyến tính

            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;
            ObjectId alignId = per.ObjectId;

            // 2. Quét chọn Block
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nQuét chọn các Block để xuất tọa độ và lý trình: ";
            TypedValue[] tvs = new TypedValue[] { new TypedValue((int)DxfCode.Start, "INSERT") };
            SelectionFilter filter = new SelectionFilter(tvs);

            PromptSelectionResult psr = ed.GetSelection(pso, filter);
            if (psr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nKhông có Block nào được chọn!");
                return;
            }

            // 3. Xin số thứ tự để đánh số trên CAD
            PromptIntegerOptions pio = new PromptIntegerOptions("\nNhập số thứ tự bắt đầu <Nhấn Enter để không đánh số>: ");
            pio.AllowNone = true;
            PromptIntegerResult pir = ed.GetInteger(pio);
            
            bool isNum = pir.Status == PromptStatus.OK;
            int startNpt = isNum ? pir.Value : 0;

            List<BlockExportData> dataList = new List<BlockExportData>();
            double txtHeight = db.Textsize;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Curve alignment = tr.GetObject(alignId, OpenMode.ForRead) as Curve;
                if (alignment == null)
                {
                    ed.WriteMessage("\nĐối tượng hướng tuyến không hợp lệ.");
                    return;
                }

                // Thu thập dữ liệu thô
                foreach (SelectedObject so in psr.Value)
                {
                    BlockReference blk = tr.GetObject(so.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (blk == null) continue;

                    Point3d pos = blk.Position;
                    double rotDeg = blk.Rotation * 180.0 / Math.PI;

                    // Lấy Tên Block (Hỗ trợ giải quyết tên Block Động "*U...")
                    string blockName = blk.Name;
                    if (blk.IsDynamicBlock)
                    {
                        BlockTableRecord dbr = (BlockTableRecord)tr.GetObject(blk.DynamicBlockTableRecord, OpenMode.ForRead);
                        blockName = dbr.Name;
                    }

                    // Tính lý trình bằng cách gióng vuông góc
                    Point3d closestPt = alignment.GetClosestPointTo(pos, false);
                    double station = alignment.GetDistAtPoint(closestPt);

                    dataList.Add(new BlockExportData
                    {
                        Position = pos,
                        RotationDeg = rotDeg,
                        Station = station,
                        BlockName = blockName,
                        Layer = blk.Layer,
                        STT = "" // Mặc định trống
                    });
                }

                // Sắp xếp các Block theo chiều dài Lý trình tăng dần
                dataList.Sort((a, b) => a.Station.CompareTo(b.Station));

                // 4. Xử lý Logic đánh Tên Trụ
                int currentKm = -1;
                int sttForName = 1;

                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                foreach (var item in dataList)
                {
                    // Lấy phần nguyên để ra mốc Km
                    int kmVal = (int)(item.Station / 1000.0);
                    if (kmVal == currentKm)
                    {
                        sttForName++;
                    }
                    else
                    {
                        currentKm = kmVal;
                        sttForName = 1; // Nhảy mốc Km mới thì reset STT tên trụ về 1
                    }

                    string sttStr = sttForName.ToString("D3"); // Format ra chuỗi 3 chữ số: 001, 002...
                    item.TenTru = $"P{kmVal}-{sttStr}";

                    // Nếu có yêu cầu đánh số trực tiếp trên AutoCAD
                    if (isNum)
                    {
                        item.STT = startNpt.ToString();

                        // Tính toán tọa độ vị trí chèn Text (với góc chếch 45 độ - hay pi/4 - như Lisp gốc)
                        double offsetX = txtHeight * Math.Cos(Math.PI / 4.0);
                        double offsetY = txtHeight * Math.Sin(Math.PI / 4.0);
                        Point3d textPos = new Point3d(item.Position.X + offsetX, item.Position.Y + offsetY, item.Position.Z);

                        // Tạo Text
                        DBText txt = new DBText();
                        txt.Position = textPos;
                        txt.TextString = startNpt.ToString();
                        txt.Height = txtHeight;
                        txt.Rotation = 0.0;
                        txt.Layer = item.Layer; // Optionally put it on the same layer as the block

                        ms.AppendEntity(txt);
                        tr.AddNewlyCreatedDBObject(txt, true);

                        startNpt++;
                    }
                }

                tr.Commit();
            }

            // 5. Xuất kết quả ra File Excel
            if (dataList.Count > 0)
            {
                ExportToExcel(doc.Name, dataList);
            }
        }

        private void ExportToExcel(string dwgName, List<BlockExportData> dataList)
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Excel Files|*.xlsx";
                    sfd.Title = "Lưu File Tọa Độ Block (.xlsx)";
                    
                    string defaultName = Path.GetFileNameWithoutExtension(dwgName);
                    if (string.IsNullOrEmpty(defaultName)) defaultName = "ToaDoBlock";
                    sfd.FileName = $"{defaultName}_XTB.xlsx";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                        ed.WriteMessage("\nĐang xuất dữ liệu ra Excel...");

                        // Tích hợp thư viện hiện đại ClosedXML thay vì Interop (chậm và hay kẹt process)
                        using (XLWorkbook workbook = new XLWorkbook())
                        {
                            string sheetName = defaultName;
                            if (sheetName.Length > 31) sheetName = sheetName.Substring(0, 31);
                            var ws = workbook.Worksheets.Add(sheetName);

                            // Header
                            string[] headers = { "STT", "Tên trụ", "X", "Y", "Z", "Góc xoay", "Lý trình (m)", "Tên Block" };
                            for (int c = 0; c < headers.Length; c++)
                            {
                                ws.Cell(1, c + 1).Value = headers[c];
                                ws.Cell(1, c + 1).Style
                                    .Font.SetBold()
                                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                                    .Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                                    .Fill.SetBackgroundColor(XLColor.LightGray);
                                ws.Cell(1, c + 1).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                            }

                            // Data
                            int row = 2;
                            foreach (var item in dataList)
                            {
                                ws.Cell(row, 1).Value = item.STT;
                                ws.Cell(row, 2).Value = item.TenTru;
                                ws.Cell(row, 3).Value = item.Position.X;
                                ws.Cell(row, 4).Value = item.Position.Y;
                                ws.Cell(row, 5).Value = item.Position.Z;
                                ws.Cell(row, 6).Value = item.RotationDeg;
                                ws.Cell(row, 7).Value = item.Station;
                                ws.Cell(row, 8).Value = item.BlockName;
                                
                                ws.Cell(row, 3).Style.NumberFormat.Format = "0.000";
                                ws.Cell(row, 4).Style.NumberFormat.Format = "0.000";
                                ws.Cell(row, 5).Style.NumberFormat.Format = "0.000";
                                ws.Cell(row, 6).Style.NumberFormat.Format = "0.0000";
                                ws.Cell(row, 7).Style.NumberFormat.Format = "0.000";

                                for (int c = 1; c <= 8; c++)
                                {
                                    ws.Cell(row, c).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                                }

                                row++;
                            }

                            ws.Columns().AdjustToContents();
                            workbook.SaveAs(sfd.FileName);
                        }

                        ed.WriteMessage("\nĐã xuất thành công!");

                        // Mở file sau khi lưu thành công
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
                    }
                }
            }
            catch (System.Exception ex)
            {
                Application.ShowAlertDialog($"Lỗi khi xuất File Excel: {ex.Message}");
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // LỆNH: XTBX
        // Xuất tọa độ X và Tên Block ra Excel (không cần chọn hướng tuyến)
        // ══════════════════════════════════════════════════════════════════════

        [CommandMethod("XTBX")]
        public void XuatToaDoXVaTen()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 1. Quét chọn Block
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nChọn các Block cần xuất tọa độ X và Tên: ";
            SelectionFilter filter = new SelectionFilter(new TypedValue[]
            {
                new TypedValue((int)DxfCode.Start, "INSERT")
            });

            PromptSelectionResult psr = ed.GetSelection(pso, filter);
            if (psr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nKhông có Block nào được chọn!");
                return;
            }

            // 2. Thu thập dữ liệu
            var dataList = new List<(string TenBlock, string TenTru, double X, double Y)>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject so in psr.Value)
                {
                    BlockReference blk = tr.GetObject(so.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (blk == null) continue;

                    // Lấy tên block (hỗ trợ Dynamic Block)
                    string blockName = blk.Name;
                    if (blk.IsDynamicBlock)
                    {
                        BlockTableRecord dbr = (BlockTableRecord)tr.GetObject(blk.DynamicBlockTableRecord, OpenMode.ForRead);
                        blockName = dbr.Name;
                    }

                    // Lấy giá trị Attribute TENTRU nếu có
                    string tenTru = "";
                    foreach (ObjectId attId in blk.AttributeCollection)
                    {
                        AttributeReference att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                        if (att != null && att.Tag.ToUpper() == "TENTRU")
                        {
                            tenTru = att.TextString;
                            break;
                        }
                    }

                    dataList.Add((blockName, tenTru, blk.Position.X, blk.Position.Y));
                }

                tr.Commit();
            }

            // Sắp xếp theo X tăng dần
            dataList.Sort((a, b) => a.X.CompareTo(b.X));

            // 3. Xuất Excel
            if (dataList.Count == 0)
            {
                ed.WriteMessage("\nKhông có dữ liệu để xuất!");
                return;
            }

            XuatExcelXVaTen(doc.Name, dataList);
        }

        private void XuatExcelXVaTen(string dwgName, List<(string TenBlock, string TenTru, double X, double Y)> dataList)
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Excel Files|*.xlsx";
                    sfd.Title = "Lưu File Tọa Độ X và Tên Block (.xlsx)";
                    string defaultName = Path.GetFileNameWithoutExtension(dwgName);
                    if (string.IsNullOrEmpty(defaultName)) defaultName = "ToaDoX";
                    sfd.FileName = $"{defaultName}_XTBX.xlsx";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                        ed.WriteMessage("\nĐang xuất dữ liệu ra Excel...");

                        using (XLWorkbook workbook = new XLWorkbook())
                        {
                            string sheetName = defaultName.Length > 31 ? defaultName.Substring(0, 31) : defaultName;
                            var ws = workbook.Worksheets.Add(sheetName);

                            // Header
                            string[] headers = { "STT", "Tên Block", "Tên trụ (TENTRU)", "X", "Y" };
                            for (int c = 0; c < headers.Length; c++)
                            {
                                var cell = ws.Cell(1, c + 1);
                                cell.Value = headers[c];
                                cell.Style.Font.SetBold();
                                cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                cell.Style.Fill.SetBackgroundColor(XLColor.LightBlue);
                                cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                            }

                            // Data
                            for (int i = 0; i < dataList.Count; i++)
                            {
                                int row = i + 2;
                                var item = dataList[i];
                                ws.Cell(row, 1).Value = i + 1;
                                ws.Cell(row, 2).Value = item.TenBlock;
                                ws.Cell(row, 3).Value = item.TenTru;
                                ws.Cell(row, 4).Value = item.X;
                                ws.Cell(row, 5).Value = item.Y;

                                ws.Cell(row, 4).Style.NumberFormat.Format = "0.000";
                                ws.Cell(row, 5).Style.NumberFormat.Format = "0.000";

                                for (int c = 1; c <= 5; c++)
                                    ws.Cell(row, c).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                            }

                            ws.Columns().AdjustToContents();
                            workbook.SaveAs(sfd.FileName);
                        }

                        ed.WriteMessage($"\n✔ Đã xuất {dataList.Count} Block ra Excel thành công!");
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
                    }
                }
            }
            catch (System.Exception ex)
            {
                Application.ShowAlertDialog($"Lỗi khi xuất File Excel: {ex.Message}");
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // LỆNH: TAONHAN
        // Tạo Block Attribute 2 dòng tại vị trí các Block dọc theo hướng tuyến:
        //   - Dòng TRÊN: Số thứ tự (001, 002, ...)
        //   - Dòng DƯỚI: Tên trụ theo lý trình (P KM0+500)
        // ══════════════════════════════════════════════════════════════════════

        private const string NHAN_BLOCK_NAME = "BR_NHAN_STT";

        [CommandMethod("TAONHAN")]
        public void TaoNhanBlockAttribute()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 1. Chọn đường hướng tuyến
            PromptEntityOptions peo = new PromptEntityOptions("\nChọn đường hướng tuyến (Line / Polyline / Spline): ");
            peo.SetRejectMessage("\nChỉ chọn đối tượng đường!");
            peo.AddAllowedClass(typeof(Curve), false);
            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;
            ObjectId alignId = per.ObjectId;

            // 2. Quét chọn các Block
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nChọn các Block cần đánh nhãn: ";
            SelectionFilter filter = new SelectionFilter(new TypedValue[]
            {
                new TypedValue((int)DxfCode.Start, "INSERT")
            });
            PromptSelectionResult psr = ed.GetSelection(pso, filter);
            if (psr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nKhông có Block nào được chọn.");
                return;
            }

            // 3. Nhập số thứ tự bắt đầu
            PromptIntegerOptions pio = new PromptIntegerOptions("\nNhập số thứ tự bắt đầu <1>: ");
            pio.AllowNone = true;
            pio.DefaultValue = 1;
            pio.LowerLimit = 1;
            PromptIntegerResult pir = ed.GetInteger(pio);
            int startStt = (pir.Status == PromptStatus.OK || pir.Status == PromptStatus.None)
                           ? pir.Value : 1;

            // Lấy chiều cao chữ hiện tại của bản vẽ
            double txtH = db.Textsize > 0 ? db.Textsize : 2.5;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Curve alignment = tr.GetObject(alignId, OpenMode.ForRead) as Curve;
                if (alignment == null)
                {
                    ed.WriteMessage("\nĐường hướng tuyến không hợp lệ.");
                    return;
                }

                // Thu thập dữ liệu + tính lý trình
                var dataList = new List<BlockExportData>();
                foreach (SelectedObject so in psr.Value)
                {
                    BlockReference blk = tr.GetObject(so.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (blk == null) continue;

                    Point3d pos = blk.Position;
                    Point3d closestPt = alignment.GetClosestPointTo(pos, false);
                    double station = alignment.GetDistAtPoint(closestPt);

                    dataList.Add(new BlockExportData
                    {
                        Position = pos,
                        Station  = station,
                        Layer    = blk.Layer
                    });
                }

                // Sắp xếp theo lý trình tăng dần
                dataList.Sort((a, b) => a.Station.CompareTo(b.Station));

                // Đảm bảo Block Definition đã tồn tại trong bản vẽ
                ObjectId nhanBlkDefId = EnsureNhanBlockDef(db, tr, txtH);

                // ModelSpace để gắn các BlockReference mới
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                int stt = startStt;
                foreach (var item in dataList)
                {
                    string sttStr  = stt.ToString("D3");          // "001", "002", ...
                    string tenTru  = FormatStation(item.Station); // "P KM0+500"

                    // Tạo BlockReference nhãn tại vị trí block gốc
                    BlockReference nhanRef = new BlockReference(item.Position, nhanBlkDefId);
                    nhanRef.Layer = item.Layer;
                    ms.AppendEntity(nhanRef);
                    tr.AddNewlyCreatedDBObject(nhanRef, true);

                    // Gắn AttributeReference vào BlockReference
                    BlockTableRecord nhanDef = (BlockTableRecord)tr.GetObject(nhanBlkDefId, OpenMode.ForRead);
                    foreach (ObjectId entId in nhanDef)
                    {
                        if (tr.GetObject(entId, OpenMode.ForRead) is AttributeDefinition attDef && !attDef.Constant)
                        {
                            AttributeReference attRef = new AttributeReference();
                            attRef.SetAttributeFromBlock(attDef, nhanRef.BlockTransform);
                            attRef.TextString = attDef.Tag == "STT" ? sttStr : tenTru;
                            nhanRef.AttributeCollection.AppendAttribute(attRef);
                            tr.AddNewlyCreatedDBObject(attRef, true);
                        }
                    }

                    stt++;
                }

                tr.Commit();
                ed.WriteMessage($"\n✔ Đã tạo {dataList.Count} nhãn Block Attribute thành công!");
            }
        }

        /// <summary>
        /// Tạo Block Definition "BR_NHAN_STT" có 2 AttributeDefinition:
        ///   STT   (trên) — số thứ tự
        ///   TENTRU (dưới) — tên trụ theo lý trình
        /// Nếu block đã tồn tại thì trả về ObjectId sẵn có.
        /// </summary>
        private ObjectId EnsureNhanBlockDef(Database db, Transaction tr, double txtH)
        {
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            if (bt.Has(NHAN_BLOCK_NAME))
                return bt[NHAN_BLOCK_NAME];

            bt.UpgradeOpen();
            BlockTableRecord btr = new BlockTableRecord { Name = NHAN_BLOCK_NAME };
            ObjectId btrId = bt.Add(btr);
            tr.AddNewlyCreatedDBObject(btr, true);

            double gap = txtH * 1.8; // khoảng cách từ điểm chèn lên/xuống

            // ── Attribute STT: phía TRÊN điểm chèn ──
            AttributeDefinition attSTT = new AttributeDefinition
            {
                Tag             = "STT",
                Prompt          = "Số thứ tự:",
                TextString      = "001",
                Height          = txtH,
                HorizontalMode  = TextHorizontalMode.TextCenter,
                VerticalMode    = TextVerticalMode.TextBase,
                AlignmentPoint  = new Point3d(0, gap, 0),
                Invisible       = false
            };
            btr.AppendEntity(attSTT);
            tr.AddNewlyCreatedDBObject(attSTT, true);

            // ── Attribute TENTRU: phía DƯỚI điểm chèn ──
            AttributeDefinition attTen = new AttributeDefinition
            {
                Tag             = "TENTRU",
                Prompt          = "Tên trụ (KM):",
                TextString      = "P KM0+000",
                Height          = txtH,
                HorizontalMode  = TextHorizontalMode.TextCenter,
                VerticalMode    = TextVerticalMode.TextTop,
                AlignmentPoint  = new Point3d(0, -gap, 0),
                Invisible       = false
            };
            btr.AppendEntity(attTen);
            tr.AddNewlyCreatedDBObject(attTen, true);

            return btrId;
        }

        /// <summary>
        /// Format lý trình (m) thành chuỗi "P KM{km}+{m:000}".
        /// Ví dụ: 1500.0 → "P KM1+500"  |  12340.0 → "P KM12+340"
        /// </summary>
        private static string FormatStation(double station)
        {
            int km = (int)(station / 1000.0);
            int m  = (int)Math.Round(station % 1000.0);
            if (m >= 1000) { km++; m = 0; } // xử lý case làm tròn vượt quá 999
            return $"P KM{km}+{m:D3}";
        }
    }
}

