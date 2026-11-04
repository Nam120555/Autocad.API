using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

// ╔══════════════════════════════════════════════════════════════════════════╗
// ║         BR TOOL  —  THỐNG KÊ TEXT TRÙNG LẶP TRÊN BẢN VẼ CAD           ║
// ║         Tác giả: HUỲNH NGỌC NAM  |  Kỹ sư Giao thông                  ║
// ╠════════════╦═════════════════════════════════════════════════════════════╣
// ║  LỆNH      ║  MÔ TẢ & CÁCH SỬ DỤNG                                     ║
// ╠════════════╬═════════════════════════════════════════════════════════════╣
// ║  DEMCS     ║  Đếm TEXT có nội dung trùng lặp và vẽ bảng thống kê       ║
// ║            ║  1. Chọn tập hợp TEXT / MTEXT cần thống kê                 ║
// ║            ║     (hoặc Enter để quét toàn bộ bản vẽ)                    ║
// ║            ║  2. Chọn điểm góc trái-trên để đặt bảng kết quả            ║
// ║            ║  → Bảng: STT | Nội dung TEXT | Số lần xuất hiện            ║
// ║            ║  → Chỉ liệt kê những chuỗi xuất hiện ≥ 2 lần              ║
// ╚════════════╩═════════════════════════════════════════════════════════════╝

namespace MyFirstProject.BrTool
{
    public class TextDuplicateCounter
    {
        // ══════════════════════════════════════════════════════════════════════
        // LỆNH: DEMCS
        // Đếm số lượng TEXT/MTEXT trùng lặp và vẽ bảng thống kê trên CAD
        // ══════════════════════════════════════════════════════════════════════

        [CommandMethod("DEMCS")]
        public void DemChuSo()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db  = doc.Database;
            Editor   ed  = doc.Editor;

            ed.WriteMessage("\n╔══════════════════════════════════════════════╗");
            ed.WriteMessage("\n║   DEMCS — THỐNG KÊ TEXT TRÙNG LẶP TRÊN CAD  ║");
            ed.WriteMessage("\n╚══════════════════════════════════════════════╝");

            // ── Bước 1: Chọn vùng TEXT cần thống kê ──────────────────────────
            SelectionFilter filter = new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Operator,   "<OR"),
                new TypedValue((int)DxfCode.Start,      "TEXT"),
                new TypedValue((int)DxfCode.Start,      "MTEXT"),
                new TypedValue((int)DxfCode.Operator,   "OR>")
            });

            PromptSelectionOptions pso = new PromptSelectionOptions
            {
                MessageForAdding  = "\nChọn TEXT/MTEXT cần thống kê (Enter = quét toàn bộ bản vẽ): ",
                AllowDuplicates   = false,
                RejectObjectsOnLockedLayers = false
            };

            PromptSelectionResult psr = ed.GetSelection(pso, filter);
            ObjectId[] ids;

            if (psr.Status == PromptStatus.OK)
            {
                ids = psr.Value.GetObjectIds();
            }
            else if (psr.Status == PromptStatus.Error || psr.Status == PromptStatus.None)
            {
                // Người dùng nhấn Enter → quét toàn bộ bản vẽ
                PromptSelectionResult allResult = ed.SelectAll(filter);
                if (allResult.Status != PromptStatus.OK || allResult.Value.Count == 0)
                {
                    ed.WriteMessage("\n⚠ Không tìm thấy TEXT nào trong bản vẽ!");
                    return;
                }
                ids = allResult.Value.GetObjectIds();
            }
            else
            {
                ed.WriteMessage("\n⚠ Đã hủy lệnh.");
                return;
            }

            ed.WriteMessage($"\n→ Đã tìm thấy {ids.Length} đối tượng TEXT/MTEXT.");

            // ── Bước 2: Thu thập nội dung TEXT ───────────────────────────────
            var textCount = new Dictionary<string, int>(StringComparer.Ordinal);
            double txtHeight = db.Textsize > 0 ? db.Textsize : 2.5;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in ids)
                {
                    string content = GetTextContent(tr, id);
                    if (string.IsNullOrWhiteSpace(content)) continue;

                    string key = content.Trim();
                    if (textCount.ContainsKey(key))
                        textCount[key]++;
                    else
                        textCount[key] = 1;
                }
                tr.Commit();
            }

            // ── Bước 3: Lọc chỉ giữ những chuỗi xuất hiện ≥ 2 lần ───────────
            var duplicates = textCount
                .Where(kv => kv.Value >= 2)
                .OrderByDescending(kv => kv.Value)
                .ThenBy(kv => kv.Key)
                .ToList();

            if (duplicates.Count == 0)
            {
                ed.WriteMessage("\n✔ Không có TEXT nào bị trùng lặp trong vùng đã chọn.");

                // Vẫn vẽ bảng thống kê đầy đủ (tất cả TEXT xuất hiện 1 lần)
                var allItems = textCount
                    .OrderByDescending(kv => kv.Value)
                    .ThenBy(kv => kv.Key)
                    .ToList();

                ed.WriteMessage("\n→ Sẽ vẽ bảng thống kê toàn bộ TEXT (kể cả không trùng).");
                DrawStatisticsTable(doc, db, ed, allItems, txtHeight, showAllNotice: true);
                return;
            }

            ed.WriteMessage($"\n✔ Có {duplicates.Count} chuỗi TEXT bị trùng lặp.");
            DrawStatisticsTable(doc, db, ed, duplicates, txtHeight, showAllNotice: false);
        }

        // ── Lấy nội dung chuỗi TEXT hoặc MTEXT ──────────────────────────────
        private static string GetTextContent(Transaction tr, ObjectId id)
        {
            DBObject obj = tr.GetObject(id, OpenMode.ForRead);

            if (obj is DBText dbt)
                return dbt.TextString;

            if (obj is MText mt)
            {
                // Loại bỏ các ký tự định dạng MTEXT thô (\\P, \\pxi..., {...})
                string raw = mt.Contents;
                raw = System.Text.RegularExpressions.Regex.Replace(raw, @"\{[^}]*\}", "");
                raw = raw.Replace("\\P", " ").Replace("\\p", " ");
                raw = System.Text.RegularExpressions.Regex.Replace(raw, @"\\[a-zA-Z][^;]*;", "");
                return raw.Trim();
            }

            return string.Empty;
        }

        // ── Vẽ bảng thống kê trực tiếp trên CAD ─────────────────────────────
        private static void DrawStatisticsTable(
            Document doc,
            Database db,
            Editor   ed,
            List<KeyValuePair<string, int>> items,
            double   txtH,
            bool     showAllNotice)
        {
            // Hỏi người dùng vị trí đặt bảng
            PromptPointOptions ppo = new PromptPointOptions("\nChọn điểm GÓC TRÁI-TRÊN để đặt bảng thống kê: ");
            PromptPointResult  ppr = ed.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n⚠ Đã hủy. Không vẽ bảng.");
                return;
            }

            Point3d origin = ppr.Value;

            // ── Thông số bảng ─────────────────────────────────────────────────
            double rowH     = txtH * 2.2;           // chiều cao mỗi hàng
            double colW1    = txtH * 4.0;            // cột STT
            double colW2    = txtH * 30.0;           // cột Nội dung TEXT
            double colW3    = txtH * 7.0;            // cột Số lần
            double totalW   = colW1 + colW2 + colW3;
            double padLeft  = txtH * 0.5;            // padding nội tâm
            double padTop   = txtH * 0.55;

            int totalRows   = items.Count + 2;       // 1 tiêu đề lớn + 1 header + n dữ liệu
            double totalH   = rowH * totalRows;

            string headerTitle = showAllNotice
                ? "THỐNG KÊ TẤT CẢ TEXT TRÊN BẢN VẼ"
                : "THỐNG KÊ TEXT TRÙNG LẶP TRÊN BẢN VẼ";

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                // ── Hàm tiện ích ──────────────────────────────────────────────
                void AddLine(Point3d p1, Point3d p2, string layer = "0")
                {
                    Line line = new Line(p1, p2) { Layer = layer };
                    ms.AppendEntity(line);
                    tr.AddNewlyCreatedDBObject(line, true);
                }

                void AddText(string txt, Point3d pos, double h,
                             TextHorizontalMode hMode = TextHorizontalMode.TextLeft,
                             bool bold = false)
                {
                    DBText t = new DBText
                    {
                        TextString    = txt,
                        Height        = h,
                        HorizontalMode = hMode,
                        VerticalMode  = TextVerticalMode.TextBase,
                    };

                    if (hMode == TextHorizontalMode.TextCenter ||
                        hMode == TextHorizontalMode.TextRight)
                        t.AlignmentPoint = pos;
                    else
                        t.Position = pos;

                    ms.AppendEntity(t);
                    tr.AddNewlyCreatedDBObject(t, true);
                }

                // ── Khung ngoài bảng ─────────────────────────────────────────
                double x0 = origin.X;
                double y0 = origin.Y;

                // Viền ngoài (4 cạnh)
                AddLine(new Point3d(x0,          y0,          0),
                        new Point3d(x0 + totalW, y0,          0));
                AddLine(new Point3d(x0,          y0 - totalH, 0),
                        new Point3d(x0 + totalW, y0 - totalH, 0));
                AddLine(new Point3d(x0,          y0,          0),
                        new Point3d(x0,          y0 - totalH, 0));
                AddLine(new Point3d(x0 + totalW, y0,          0),
                        new Point3d(x0 + totalW, y0 - totalH, 0));

                // ── Hàng 0: Tiêu đề lớn ──────────────────────────────────────
                double yRow0Top = y0;
                double yRow0Bot = y0 - rowH;

                // Đường kẻ ngang dưới hàng tiêu đề
                AddLine(new Point3d(x0, yRow0Bot, 0), new Point3d(x0 + totalW, yRow0Bot, 0));

                // Text tiêu đề — căn giữa toàn bảng
                double titleX = x0 + totalW / 2.0;
                double titleY = yRow0Bot + padTop;
                AddText(headerTitle, new Point3d(titleX, titleY, 0),
                        txtH * 1.2, TextHorizontalMode.TextCenter);

                // ── Hàng 1: Header cột ───────────────────────────────────────
                double yRow1Top = yRow0Bot;
                double yRow1Bot = yRow1Top - rowH;
                AddLine(new Point3d(x0, yRow1Bot, 0), new Point3d(x0 + totalW, yRow1Bot, 0));

                // Đường dọc phân cột
                double xCol1 = x0 + colW1;
                double xCol2 = xCol1 + colW2;

                AddLine(new Point3d(xCol1, y0, 0), new Point3d(xCol1, y0 - totalH, 0));
                AddLine(new Point3d(xCol2, y0, 0), new Point3d(xCol2, y0 - totalH, 0));

                // Text header cột
                double hdrY = yRow1Bot + padTop;
                AddText("STT",              new Point3d(x0    + colW1/2, hdrY, 0), txtH, TextHorizontalMode.TextCenter);
                AddText("NỘI DUNG TEXT",    new Point3d(xCol1 + colW2/2, hdrY, 0), txtH, TextHorizontalMode.TextCenter);
                AddText("SỐ LẦN",          new Point3d(xCol2 + colW3/2, hdrY, 0), txtH, TextHorizontalMode.TextCenter);

                // ── Hàng dữ liệu ─────────────────────────────────────────────
                for (int i = 0; i < items.Count; i++)
                {
                    double yTop = yRow1Bot - i * rowH;
                    double yBot = yTop - rowH;

                    // Đường kẻ ngang (trừ dòng cuối đã có viền ngoài)
                    if (i < items.Count - 1)
                        AddLine(new Point3d(x0, yBot, 0), new Point3d(x0 + totalW, yBot, 0));

                    double dataY = yBot + padTop;

                    // STT — căn giữa
                    AddText((i + 1).ToString(),
                            new Point3d(x0 + colW1/2, dataY, 0),
                            txtH, TextHorizontalMode.TextCenter);

                    // Nội dung TEXT — căn trái, có padding
                    string content = items[i].Key;
                    // Cắt bớt nếu quá dài (tránh tràn cột)
                    int maxChars = (int)(colW2 / (txtH * 0.65));
                    if (content.Length > maxChars)
                        content = content.Substring(0, maxChars - 3) + "...";

                    AddText(content,
                            new Point3d(xCol1 + padLeft, dataY, 0),
                            txtH, TextHorizontalMode.TextLeft);

                    // Số lần — căn giữa
                    AddText(items[i].Value.ToString(),
                            new Point3d(xCol2 + colW3/2, dataY, 0),
                            txtH, TextHorizontalMode.TextCenter);
                }

                tr.Commit();
            }

            ed.WriteMessage($"\n✔ Đã vẽ bảng thống kê {items.Count} mục thành công!");
            ed.WriteMessage($"\n  Tổng số chuỗi TEXT đã thống kê: {items.Sum(x => x.Value)} lần xuất hiện.");
        }
    }
}
