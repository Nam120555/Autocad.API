using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║         LẤY TỌA ĐỘ X, Y THEO BLOCK & TÊN ĐIỂM THEO TEXT GẦN NHẤT         ║
// ║         Tác giả: HUỲNH NGỌC NAM  |  Kỹ sư Giao thông                      ║
// ╠══════════════╦═════════════════════════════════════════════════════════════╣
// ║  LỆNH        ║  MÔ TẢ                                                     ║
// ╠══════════════╬═════════════════════════════════════════════════════════════╣
// ║  LTDTB       ║  Lấy tọa độ X, Y của Block; tên điểm lấy từ Text / MText  ║
// ║              ║  gần Block nhất trong bản vẽ.                               ║
// ║              ║  1. Chọn các Block cần lấy tọa độ                           ║
// ║              ║  2. (Tùy chọn) Giới hạn khoảng cách tìm Text gần nhất      ║
// ║              ║  3. Xuất: STT | Tên điểm | X | Y | Khoảng cách text-block  ║
// ╚══════════════╩═════════════════════════════════════════════════════════════╝

namespace MyFirstProject.BrTool
{
    public class LayToaDoTenTheoBlockVaText
    {
        // ── Dữ liệu một điểm đã xử lý ──────────────────────────────────────────
        private class PointData
        {
            public int    STT         { get; set; }
            public string TenDiem     { get; set; } = string.Empty;
            public double X           { get; set; }
            public double Y           { get; set; }
            public double DistToText  { get; set; }   // khoảng cách tới text gần nhất (m)
            public string BlockName   { get; set; } = string.Empty;
        }

        // ════════════════════════════════════════════════════════════════════════
        // LỆNH: LTDTB
        // Lấy tọa độ X, Y theo Block và tên điểm theo Text gần Block nhất
        // ════════════════════════════════════════════════════════════════════════
        [CommandMethod("LTDTB")]
        public void LayToaDoVaTen()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db  = doc.Database;
            Editor   ed  = doc.Editor;

            // ── BƯỚC 1: Chọn các Block cần lấy tọa độ ─────────────────────────
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nChọn các Block cần lấy tọa độ (Enter để kết thúc): ";
            SelectionFilter blkFilter = new SelectionFilter(new TypedValue[]
            {
                new TypedValue((int)DxfCode.Start, "INSERT")
            });

            PromptSelectionResult psr = ed.GetSelection(pso, blkFilter);
            if (psr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nKhông có Block nào được chọn!");
                return;
            }

            // ── BƯỚC 2: Nhập khoảng cách tìm kiếm Text (0 = không giới hạn) ──
            PromptDoubleOptions pdo = new PromptDoubleOptions(
                "\nNhập khoảng cách tối đa tìm Text gần nhất (0 = không giới hạn) <0>: ");
            pdo.AllowNone        = true;
            pdo.AllowNegative    = false;
            pdo.DefaultValue     = 0.0;
            PromptDoubleResult pdr = ed.GetDouble(pdo);

            double maxDist = (pdr.Status == PromptStatus.OK) ? pdr.Value : 0.0;
            bool   limitDist = maxDist > 0.0;

            // ── BƯỚC 3: Thu thập dữ liệu Block + Text trong Transaction ────────
            var resultList = new List<PointData>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // ── 3a. Thu thập tất cả DBText và MText trong bản vẽ ────────────
                var textObjects = new List<(string Content, Point3d Position)>();

                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                foreach (ObjectId entId in ms)
                {
                    try
                    {
                        DBObject ent = tr.GetObject(entId, OpenMode.ForRead);

                        if (ent is DBText dbt && !string.IsNullOrWhiteSpace(dbt.TextString))
                        {
                            textObjects.Add((dbt.TextString.Trim(), dbt.Position));
                        }
                        else if (ent is MText mt && !string.IsNullOrWhiteSpace(mt.Contents))
                        {
                            // Lấy văn bản thuần (bỏ mã định dạng MTEXT)
                            string plain = mt.Text;
                            if (!string.IsNullOrWhiteSpace(plain))
                                textObjects.Add((plain.Trim(), mt.Location));
                        }
                    }
                    catch { /* bỏ qua entity lỗi */ }
                }

                if (textObjects.Count == 0)
                {
                    ed.WriteMessage("\n⚠ Bản vẽ không có Text / MText nào để tìm tên điểm!");
                    tr.Commit();
                    return;
                }

                // ── 3b. Với mỗi Block → tìm Text gần nhất ──────────────────────
                int stt = 1;
                foreach (SelectedObject so in psr.Value)
                {
                    BlockReference blk = tr.GetObject(so.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (blk == null) continue;

                    Point3d blkPos = blk.Position;

                    // Lấy tên Block (hỗ trợ Dynamic Block)
                    string blockName = blk.Name;
                    if (blk.IsDynamicBlock)
                    {
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            blk.DynamicBlockTableRecord, OpenMode.ForRead);
                        blockName = btr.Name;
                    }

                    // Tìm Text gần nhất (dùng khoảng cách phẳng XY)
                    string  nearestText = string.Empty;
                    double  nearestDist = double.MaxValue;

                    foreach (var (content, pos) in textObjects)
                    {
                        double dx   = pos.X - blkPos.X;
                        double dy   = pos.Y - blkPos.Y;
                        double dist = Math.Sqrt(dx * dx + dy * dy);

                        if (limitDist && dist > maxDist) continue;

                        if (dist < nearestDist)
                        {
                            nearestDist = dist;
                            nearestText = content;
                        }
                    }

                    // Nếu không tìm được Text nào (do giới hạn khoảng cách) → để trống
                    if (nearestText == string.Empty && limitDist)
                    {
                        ed.WriteMessage($"\n⚠ Block STT {stt} ({blockName}) không có Text nào trong bán kính {maxDist:F3} m.");
                    }

                    resultList.Add(new PointData
                    {
                        STT        = stt++,
                        TenDiem    = nearestText,
                        X          = blkPos.X,
                        Y          = blkPos.Y,
                        DistToText = nearestDist == double.MaxValue ? 0.0 : nearestDist,
                        BlockName  = blockName
                    });
                }

                tr.Commit();
            }

            // ── BƯỚC 4: Xuất kết quả ────────────────────────────────────────────
            if (resultList.Count == 0)
            {
                ed.WriteMessage("\nKhông có dữ liệu để xuất!");
                return;
            }

            ed.WriteMessage($"\n✔ Đã xử lý {resultList.Count} Block.");

            // ── BƯỚC 5: Hỏi định dạng xuất ─────────────────────────────────────
            string choice = ShowExportFormatDialog();
            if (choice == "EXCEL" || choice == "BOTH")
                ExportToExcel(doc.Name, resultList, ed);
            if (choice == "LISP"  || choice == "BOTH")
                ExportToLisp(doc.Name, resultList, ed);
        }

        // ════════════════════════════════════════════════════════════════════════
        // Xuất danh sách điểm ra File Excel
        // ════════════════════════════════════════════════════════════════════════
        private void ExportToExcel(string dwgName, List<PointData> dataList, Editor ed)
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter   = "Excel Files|*.xlsx";
                    sfd.Title    = "Lưu File Tọa Độ & Tên Điểm (.xlsx)";
                    string defaultName = Path.GetFileNameWithoutExtension(dwgName);
                    if (string.IsNullOrEmpty(defaultName)) defaultName = "ToaDo_TenDiem";
                    sfd.FileName = $"{defaultName}_LTDTB.xlsx";

                    if (sfd.ShowDialog() != DialogResult.OK) return;

                    ed.WriteMessage("\nĐang xuất dữ liệu ra Excel...");

                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        string sheet = defaultName.Length > 31
                            ? defaultName.Substring(0, 31) : defaultName;
                        var ws = wb.Worksheets.Add(sheet);

                        // ── Header ────────────────────────────────────────────
                        string[] headers =
                        {
                            "STT", "Tên điểm", "X (m)", "Y (m)",
                            "Khoảng cách Text-Block (m)", "Tên Block"
                        };

                        for (int c = 0; c < headers.Length; c++)
                        {
                            var cell = ws.Cell(1, c + 1);
                            cell.Value = headers[c];
                            cell.Style
                                .Font.SetBold(true)
                                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                                .Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                                .Fill.SetBackgroundColor(XLColor.FromHtml("#1F5C99"));
                            cell.Style.Font.SetFontColor(XLColor.White);
                            cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                        }

                        // ── Dữ liệu ───────────────────────────────────────────
                        for (int i = 0; i < dataList.Count; i++)
                        {
                            int row  = i + 2;
                            var item = dataList[i];

                            ws.Cell(row, 1).Value = item.STT;
                            ws.Cell(row, 2).Value = item.TenDiem;
                            ws.Cell(row, 3).Value = item.X;
                            ws.Cell(row, 4).Value = item.Y;
                            ws.Cell(row, 5).Value = item.DistToText;
                            ws.Cell(row, 6).Value = item.BlockName;

                            ws.Cell(row, 3).Style.NumberFormat.Format = "0.000";
                            ws.Cell(row, 4).Style.NumberFormat.Format = "0.000";
                            ws.Cell(row, 5).Style.NumberFormat.Format = "0.000";

                            // Tô màu xen kẽ hàng
                            XLColor rowColor = (i % 2 == 0)
                                ? XLColor.FromHtml("#EAF2FB")
                                : XLColor.White;

                            for (int c = 1; c <= 6; c++)
                            {
                                ws.Cell(row, c).Style
                                    .Fill.SetBackgroundColor(rowColor)
                                    .Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                            }

                            // Căn giữa cột STT và khoảng cách
                            ws.Cell(row, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            ws.Cell(row, 5).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        }

                        // ── Freeze hàng header và tự co cột ─────────────────
                        ws.SheetView.FreezeRows(1);
                        ws.Columns().AdjustToContents();

                        wb.SaveAs(sfd.FileName);
                    }

                    ed.WriteMessage($"\n✔ Đã xuất {dataList.Count} điểm ra Excel thành công!\n   → {sfd.FileName}");

                    // Mở file sau khi lưu
                    System.Diagnostics.Process.Start(
                        new System.Diagnostics.ProcessStartInfo(sfd.FileName)
                        { UseShellExecute = true });
                }
            }
            catch (System.Exception ex)
            {
                Application.ShowAlertDialog($"Lỗi khi xuất File Excel:\n{ex.Message}");
            }
        }

        // ════════════════════════════════════════════════════════════════════════
        // Hộp thoại chọn định dạng xuất: Excel / LISP / Cả hai
        // ════════════════════════════════════════════════════════════════════════
        private string ShowExportFormatDialog()
        {
            using (Form frm = new Form())
            {
                frm.Text            = "Chọn định dạng xuất";
                frm.Size            = new System.Drawing.Size(340, 210);
                frm.StartPosition   = FormStartPosition.CenterScreen;
                frm.FormBorderStyle = FormBorderStyle.FixedDialog;
                frm.MaximizeBox     = false;
                frm.MinimizeBox     = false;

                var lbl = new Label
                {
                    Text     = "Xuất dữ liệu tọa độ sang định dạng nào?",
                    Location = new System.Drawing.Point(14, 14),
                    Size     = new System.Drawing.Size(300, 32),
                    Font     = new System.Drawing.Font("Segoe UI", 9.5f)
                };

                var btnExcel = new Button { Text = "📊  Excel (.xlsx)",  Location = new System.Drawing.Point(14,  58), Size = new System.Drawing.Size(140, 38), Tag = "EXCEL" };
                var btnLisp  = new Button { Text = "📜  LISP (.lsp)",    Location = new System.Drawing.Point(162, 58), Size = new System.Drawing.Size(140, 38), Tag = "LISP"  };
                var btnBoth  = new Button { Text = "✅  Cả hai",         Location = new System.Drawing.Point(14, 108), Size = new System.Drawing.Size(288, 38), Tag = "BOTH"  };

                string result = "CANCEL";
                EventHandler click = (s, e) =>
                {
                    result = ((Button)s).Tag.ToString();
                    frm.Close();
                };
                btnExcel.Click += click;
                btnLisp.Click  += click;
                btnBoth.Click  += click;

                frm.Controls.AddRange(new Control[] { lbl, btnExcel, btnLisp, btnBoth });
                frm.ShowDialog();
                return result;
            }
        }

        // ════════════════════════════════════════════════════════════════════════
        // Xuất danh sách điểm ra File LISP (.lsp)
        // Khi load vào AutoCAD: gõ TAODIEM → tự tạo POINT + TEXT tên điểm
        // ════════════════════════════════════════════════════════════════════════
        private void ExportToLisp(string dwgName, List<PointData> dataList, Editor ed)
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter   = "AutoCAD LISP Files|*.lsp";
                    sfd.Title    = "Lưu File LISP Tọa Độ & Tên Điểm (.lsp)";
                    string defaultName = Path.GetFileNameWithoutExtension(dwgName);
                    if (string.IsNullOrEmpty(defaultName)) defaultName = "ToaDo_TenDiem";
                    sfd.FileName = $"{defaultName}_LTDTB.lsp";

                    if (sfd.ShowDialog() != DialogResult.OK) return;

                    ed.WriteMessage("\nĐang xuất dữ liệu ra file LISP...");

                    using (StreamWriter sw = new StreamWriter(sfd.FileName, false,
                        System.Text.Encoding.UTF8))
                    {
                        // ── Phần đầu file ──────────────────────────────────────
                        sw.WriteLine(";; ============================================================");
                        sw.WriteLine(";; File LISP xuất bởi lệnh LTDTB - AutoCAD BR Tool");
                        sw.WriteLine($";; Bản vẽ gốc : {Path.GetFileName(dwgName)}");
                        sw.WriteLine($";; Thời gian  : {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
                        sw.WriteLine($";; Tổng điểm  : {dataList.Count}");
                        sw.WriteLine(";;");
                        sw.WriteLine(";; CÁCH DÙNG:");
                        sw.WriteLine(";;   1. Trong AutoCAD: Tools → Load Application → chọn file này");
                        sw.WriteLine(";;      hoặc gõ: (load \"duong_dan_file.lsp\")");
                        sw.WriteLine(";;   2. Gõ lệnh: TAODIEM  → Enter");
                        sw.WriteLine(";;      AutoCAD sẽ vẽ POINT + TEXT tên điểm tại đúng tọa độ.");
                        sw.WriteLine(";;   3. Gõ lệnh: XEMDS    → xem danh sách điểm ở Command Line.");
                        sw.WriteLine(";; ============================================================");
                        sw.WriteLine();

                        // ── Danh sách tọa độ dạng Lisp list ────────────────────
                        sw.WriteLine(";; Danh sách điểm: (\"TenDiem\" X Y)");
                        sw.WriteLine("(setq *LTDTB-DIEM-LIST* (quote (");
                        foreach (var p in dataList)
                        {
                            // Escape dấu nháy kép trong tên điểm
                            string safeName = p.TenDiem
                                .Replace("\\", "\\\\")
                                .Replace("\"", "\\\"")
                                .Replace("\r", "")
                                .Replace("\n", " ");
                            sw.WriteLine($"    (\"{safeName}\" {p.X:F6} {p.Y:F6})");
                        }
                        sw.WriteLine(")))");
                        sw.WriteLine();

                        // ── Hàm TAODIEM ──────────────────────────────────────────
                        sw.WriteLine(";; ── Hàm chính: tạo POINT và TEXT tên điểm ──────────────────");
                        sw.WriteLine("(defun C:TAODIEM ( / txtH pt ten txtPt)");
                        sw.WriteLine("  ;; Lấy chiều cao chữ hiện tại");
                        sw.WriteLine("  (setq txtH (getvar \"TEXTSIZE\"))");
                        sw.WriteLine("  (if (or (null txtH) (= txtH 0.0)) (setq txtH 2.5))");
                        sw.WriteLine();
                        sw.WriteLine("  ;; Kiểu hiển thị điểm đẹp: dấu X trong vòng tròn (PDMODE=35)");
                        sw.WriteLine("  (command \"._PDMODE\" 35)");
                        sw.WriteLine("  (command \"._PDSIZE\" (* txtH 1.5))");
                        sw.WriteLine();
                        sw.WriteLine("  (foreach diem *LTDTB-DIEM-LIST*");
                        sw.WriteLine("    (setq ten    (car   diem))");
                        sw.WriteLine("    (setq pt     (list  (cadr diem) (caddr diem) 0.0))");
                        sw.WriteLine();
                        sw.WriteLine("    ;; Vẽ điểm POINT");
                        sw.WriteLine("    (command \"._POINT\" pt)");
                        sw.WriteLine();
                        sw.WriteLine("    ;; Vẽ TEXT tên điểm (lệch phải + lên 0.5*txtH so với điểm)");
                        sw.WriteLine("    (setq txtPt (list (+ (cadr diem) (* txtH 0.7))");
                        sw.WriteLine("                     (+ (caddr diem) (* txtH 0.5))");
                        sw.WriteLine("                     0.0))");
                        sw.WriteLine("    (command \"._TEXT\" txtPt txtH 0 ten)");
                        sw.WriteLine("  )");
                        sw.WriteLine();
                        sw.WriteLine($"  (princ (strcat \"\\n✔ Đã tạo {dataList.Count} điểm thành công!\"))");
                        sw.WriteLine("  (princ)");
                        sw.WriteLine(")");
                        sw.WriteLine();

                        // ── Hàm phụ: xem danh sách ──────────────────────────────
                        sw.WriteLine(";; ── Hàm xem danh sách điểm (không vẽ) ────────────────────────");
                        sw.WriteLine("(defun C:XEMDS ( / )");
                        sw.WriteLine($"  (princ \"\\n=== Danh sách {dataList.Count} điểm ===========\")");
                        sw.WriteLine("  (foreach diem *LTDTB-DIEM-LIST*");
                        sw.WriteLine("    (princ (strcat \"\\n  \" (car diem)");
                        sw.WriteLine("                   \"  X=\" (rtos (cadr diem)  2 3)");
                        sw.WriteLine("                   \"  Y=\" (rtos (caddr diem) 2 3)))");
                        sw.WriteLine("  )");
                        sw.WriteLine("  (princ)");
                        sw.WriteLine(")");
                        sw.WriteLine();

                        // ── Thông báo khi load ───────────────────────────────────
                        sw.WriteLine(";; Thông báo khi load thành công");
                        sw.WriteLine($"(princ \"\\n[LTDTB] Load {dataList.Count} điểm. Gõ TAODIEM để vẽ | XEMDS để xem danh sách.\")");
                        sw.WriteLine("(princ)");
                    }

                    ed.WriteMessage($"\n✔ Đã xuất {dataList.Count} điểm ra file LISP thành công!\n   → {sfd.FileName}");
                    ed.WriteMessage("\n   Cách dùng: Load file vào AutoCAD → gõ lệnh TAODIEM");

                    // Mở file sau khi lưu
                    System.Diagnostics.Process.Start(
                        new System.Diagnostics.ProcessStartInfo(sfd.FileName)
                        { UseShellExecute = true });
                }
            }
            catch (System.Exception ex)
            {
                Application.ShowAlertDialog($"Lỗi khi xuất File LISP:\n{ex.Message}");
            }
        }
    }
}
