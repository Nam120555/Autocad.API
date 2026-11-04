// Demtext.cs — Thống kê & Phân loại toàn bộ TEXT trên bản vẽ AutoCAD
// Tác giả: HUỲNH NGỌC NAM | Kỹ sư Giao thông
//
// ╔══════════════════════════════════════════════════════════════════════════╗
// ║         CIVIL TOOL  —  THỐNG KÊ & PHÂN LOẠI TEXT TRÊN BẢN VẼ           ║
// ╠════════════════╦═════════════════════════════════════════════════════════╣
// ║  LỆNH          ║  MÔ TẢ                                                  ║
// ╠════════════════╬═════════════════════════════════════════════════════════╣
// ║  CTDEM         ║  Đếm & phân loại TOÀN BỘ TEXT/MTEXT/Attribute          ║
// ║                ║  • Chọn vùng hoặc Enter = toàn bộ bản vẽ               ║
// ║                ║  • Phân loại theo: Loại đối tượng / Layer / Chiều cao   ║
// ║                ║  • Bảng nội dung trùng lặp (≥ 2 lần)                   ║
// ║                ║  • Vẽ 4 bảng thống kê trực tiếp lên CAD                ║
// ╚════════════════╩═════════════════════════════════════════════════════════╝

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.TextStatisticsCommand))]

namespace MyFirstProject
{
    // ══════════════════════════════════════════════════════════════════════════
    // Model lưu thông tin từng đối tượng TEXT thu thập được
    // ══════════════════════════════════════════════════════════════════════════
    internal class TextItemInfo
    {
        public string ObjectType { get; set; }   // "TEXT" | "MTEXT" | "ATTRIB"
        public string Content    { get; set; }
        public string Layer      { get; set; }
        public double Height     { get; set; }
    }

    // ══════════════════════════════════════════════════════════════════════════
    // LỆNH CTDEM — Thống kê & Phân loại TEXT
    // ══════════════════════════════════════════════════════════════════════════
    public class TextStatisticsCommand
    {
        [CommandMethod("CTDEM")]
        public void CountAndClassifyText()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db  = doc.Database;
            Editor   ed  = doc.Editor;

            ed.WriteMessage("\n╔════════════════════════════════════════════════════╗");
            ed.WriteMessage("\n║   CTDEM — THỐNG KÊ & PHÂN LOẠI TEXT TRÊN BẢN VẼ   ║");
            ed.WriteMessage("\n╚════════════════════════════════════════════════════╝");

            // ── BƯỚC 1: Chọn đối tượng ──────────────────────────────────────
            SelectionFilter filter = BuildTextFilter();

            PromptSelectionOptions pso = new PromptSelectionOptions
            {
                MessageForAdding = "\nChọn TEXT/MTEXT cần thống kê (Enter = toàn bộ bản vẽ): ",
                RejectObjectsOnLockedLayers = false
            };

            ObjectId[] ids;
            PromptSelectionResult psr = ed.GetSelection(pso, filter);

            if (psr.Status == PromptStatus.OK)
            {
                ids = psr.Value.GetObjectIds();
                ed.WriteMessage($"\n→ Đã chọn {ids.Length} đối tượng TEXT.");
            }
            else
            {
                PromptSelectionResult allRes = ed.SelectAll(filter);
                if (allRes.Status != PromptStatus.OK || allRes.Value.Count == 0)
                {
                    ed.WriteMessage("\n⚠ Không tìm thấy TEXT nào trong bản vẽ!");
                    return;
                }
                ids = allRes.Value.GetObjectIds();
                ed.WriteMessage($"\n→ Quét toàn bộ bản vẽ: {ids.Length} đối tượng TEXT.");
            }

            // ── BƯỚC 2: Thu thập dữ liệu ────────────────────────────────────
            var allItems = new List<TextItemInfo>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in ids)
                {
                    CollectTextItems(tr, id, allItems);
                }

                // Thu thập thêm AttributeReference từ BlockReference trong vùng chọn
                CollectAttributesFromBlocks(tr, db, ids, allItems);

                tr.Commit();
            }

            if (allItems.Count == 0)
            {
                ed.WriteMessage("\n⚠ Không thu thập được dữ liệu TEXT hợp lệ!");
                return;
            }

            ed.WriteMessage($"\n✔ Tổng cộng thu thập được {allItems.Count} chuỗi TEXT.");

            // ── BƯỚC 3: Phân loại ────────────────────────────────────────────
            // 3a. Theo loại đối tượng
            var byType = allItems
                .GroupBy(x => x.ObjectType)
                .Select(g => new KeyValuePair<string, int>(g.Key, g.Count()))
                .OrderByDescending(kv => kv.Value)
                .ToList();

            // 3b. Theo Layer
            var byLayer = allItems
                .GroupBy(x => x.Layer)
                .Select(g => new KeyValuePair<string, int>(g.Key, g.Count()))
                .OrderByDescending(kv => kv.Value)
                .ToList();

            // 3c. Theo chiều cao chữ (làm tròn 2 chữ số thập phân)
            var byHeight = allItems
                .GroupBy(x => Math.Round(x.Height, 2))
                .Select(g => new KeyValuePair<string, int>(g.Key.ToString("0.00"), g.Count()))
                .OrderByDescending(kv => kv.Value)
                .ToList();

            // 3d. Nội dung TEXT trùng lặp (≥ 2 lần)
            var duplicates = allItems
                .GroupBy(x => x.Content, StringComparer.Ordinal)
                .Where(g => g.Count() >= 2)
                .Select(g => new KeyValuePair<string, int>(g.Key, g.Count()))
                .OrderByDescending(kv => kv.Value)
                .ThenBy(kv => kv.Key)
                .ToList();

            // ── BƯỚC 4: Báo cáo lên Command Line ────────────────────────────
            PrintReport(ed, allItems, byType, byLayer, byHeight, duplicates);

            // ── BƯỚC 5: Vẽ bảng trên CAD ────────────────────────────────────
            double txtH = db.Textsize > 0 ? db.Textsize : 2.5;

            PromptPointOptions ppo = new PromptPointOptions(
                "\nChọn điểm GÓC TRÁI-TRÊN để đặt bảng thống kê: ");
            PromptPointResult ppr = ed.GetPoint(ppo);

            if (ppr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n⚠ Đã hủy. Không vẽ bảng.");
                return;
            }

            Point3d origin = ppr.Value;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                TablePainter tp = new TablePainter(ms, tr, txtH);

                double curX = origin.X;
                double curY = origin.Y;
                double gap  = txtH * 4; // khoảng cách ngang giữa các bảng

                // Bảng 1 — Theo loại đối tượng
                double w1 = tp.DrawTable(curX, curY,
                    "PHÂN LOẠI THEO LOẠI ĐỐI TƯỢNG",
                    new[] { "Loại đối tượng", "Số lượng" },
                    new[] { txtH * 18, txtH * 8 },
                    byType.Select(kv => new[] { kv.Key, kv.Value.ToString() }).ToList());
                curX += w1 + gap;

                // Bảng 2 — Theo Layer
                double w2 = tp.DrawTable(curX, curY,
                    "PHÂN LOẠI THEO LAYER",
                    new[] { "Layer", "Số lượng" },
                    new[] { txtH * 22, txtH * 8 },
                    byLayer.Select(kv => new[] { kv.Key, kv.Value.ToString() }).ToList());
                curX += w2 + gap;

                // Bảng 3 — Theo chiều cao chữ
                double w3 = tp.DrawTable(curX, curY,
                    "PHÂN LOẠI THEO CHIỀU CAO CHỮ",
                    new[] { "Cao chữ (đv bản vẽ)", "Số lượng" },
                    new[] { txtH * 20, txtH * 8 },
                    byHeight.Select(kv => new[] { kv.Key, kv.Value.ToString() }).ToList());

                // Bảng 4 — Nội dung trùng lặp (xuống dưới)
                if (duplicates.Count > 0)
                {
                    double tableH = (byType.Count + 3) * txtH * 2.2; // ước lượng cao bảng 1
                    double dupY   = origin.Y - tableH - txtH * 6;

                    tp.DrawTable(origin.X, dupY,
                        $"NỘI DUNG TEXT TRÙNG LẶP (≥ 2 LẦN)  —  Tổng: {duplicates.Count} chuỗi",
                        new[] { "STT", "Nội dung TEXT", "Số lần" },
                        new[] { txtH * 5, txtH * 40, txtH * 8 },
                        duplicates.Select((kv, i) =>
                        {
                            string c = kv.Key.Length > 60 ? kv.Key.Substring(0, 57) + "..." : kv.Key;
                            return new[] { (i + 1).ToString(), c, kv.Value.ToString() };
                        }).ToList());
                }
                else
                {
                    // Thông báo không có trùng
                    tp.DrawTable(origin.X, origin.Y - (byType.Count + 3) * txtH * 2.2 - txtH * 6,
                        "NỘI DUNG TEXT TRÙNG LẶP",
                        new[] { "Kết quả" },
                        new[] { txtH * 35 },
                        new List<string[]> { new[] { "✔ Không có chuỗi TEXT nào bị trùng lặp" } });
                }

                tr.Commit();
            }

            ed.WriteMessage($"\n✔ Đã vẽ bảng thống kê hoàn chỉnh lên bản vẽ!");
        }

        // ── Xây dựng SelectionFilter cho TEXT + MTEXT ────────────────────────
        private static SelectionFilter BuildTextFilter()
        {
            return new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Operator, "<OR"),
                new TypedValue((int)DxfCode.Start,    "TEXT"),
                new TypedValue((int)DxfCode.Start,    "MTEXT"),
                new TypedValue((int)DxfCode.Operator, "OR>")
            });
        }

        // ── Thu thập TEXT từ từng ObjectId ───────────────────────────────────
        private static void CollectTextItems(Transaction tr, ObjectId id, List<TextItemInfo> list)
        {
            DBObject obj = tr.GetObject(id, OpenMode.ForRead);

            if (obj is DBText dbt && !string.IsNullOrWhiteSpace(dbt.TextString))
            {
                list.Add(new TextItemInfo
                {
                    ObjectType = "TEXT",
                    Content    = dbt.TextString.Trim(),
                    Layer      = dbt.Layer,
                    Height     = dbt.Height
                });
            }
            else if (obj is MText mt)
            {
                string raw = StripMTextFormatting(mt.Contents);
                if (!string.IsNullOrWhiteSpace(raw))
                {
                    list.Add(new TextItemInfo
                    {
                        ObjectType = "MTEXT",
                        Content    = raw.Trim(),
                        Layer      = mt.Layer,
                        Height     = mt.TextHeight
                    });
                }
            }
        }

        // ── Thu thập AttributeReference từ Block ─────────────────────────────
        private static void CollectAttributesFromBlocks(
            Transaction tr, Database db, ObjectId[] selectedIds, List<TextItemInfo> list)
        {
            // Lấy tất cả BlockReference trong ModelSpace
            BlockTableRecord ms = (BlockTableRecord)tr.GetObject(
                SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

            foreach (ObjectId entId in ms)
            {
                BlockReference blkRef = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                if (blkRef == null) continue;

                foreach (ObjectId attId in blkRef.AttributeCollection)
                {
                    AttributeReference att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                    if (att == null || string.IsNullOrWhiteSpace(att.TextString)) continue;

                    list.Add(new TextItemInfo
                    {
                        ObjectType = "ATTRIB",
                        Content    = att.TextString.Trim(),
                        Layer      = att.Layer,
                        Height     = att.Height
                    });
                }
            }
        }

        // ── Loại bỏ định dạng MTEXT thô ─────────────────────────────────────
        private static string StripMTextFormatting(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return string.Empty;
            // Bỏ các lệnh định dạng như \f...; \H...; \W...; \Q...; \C...; \T...
            raw = Regex.Replace(raw, @"\\[fFHhWwQqCcTtAaLlOoKkSsPpI]([^;]*;)?", "");
            // Bỏ nhóm {...}
            raw = Regex.Replace(raw, @"\{([^{}]*)\}", "$1");
            // Bỏ { và }
            raw = raw.Replace("{", "").Replace("}", "");
            // Thay \P bằng dấu cách
            raw = raw.Replace("\\P", " ").Replace("\\p", " ");
            // Bỏ các ký tự điều khiển còn lại
            raw = Regex.Replace(raw, @"\\.", "");
            return raw.Trim();
        }

        // ── In báo cáo ra Command Line ────────────────────────────────────────
        private static void PrintReport(
            Editor ed,
            List<TextItemInfo> allItems,
            List<KeyValuePair<string, int>> byType,
            List<KeyValuePair<string, int>> byLayer,
            List<KeyValuePair<string, int>> byHeight,
            List<KeyValuePair<string, int>> duplicates)
        {
            ed.WriteMessage("\n");
            ed.WriteMessage("\n══════════════════════════════════════════════════════════");
            ed.WriteMessage($"\n  TỔNG SỐ CHUỖI TEXT THU THẬP ĐƯỢC: {allItems.Count}");
            ed.WriteMessage("\n──────────────────────────────────────────────────────────");

            ed.WriteMessage("\n  [1] PHÂN LOẠI THEO LOẠI ĐỐI TƯỢNG:");
            foreach (var kv in byType)
                ed.WriteMessage($"\n      • {kv.Key,-10} : {kv.Value,5} chuỗi");

            ed.WriteMessage("\n──────────────────────────────────────────────────────────");
            ed.WriteMessage("\n  [2] PHÂN LOẠI THEO LAYER (Top 10):");
            foreach (var kv in byLayer.Take(10))
                ed.WriteMessage($"\n      • {kv.Key,-30} : {kv.Value,5} chuỗi");
            if (byLayer.Count > 10)
                ed.WriteMessage($"\n      ... và {byLayer.Count - 10} layer khác.");

            ed.WriteMessage("\n──────────────────────────────────────────────────────────");
            ed.WriteMessage("\n  [3] PHÂN LOẠI THEO CHIỀU CAO CHỮ:");
            foreach (var kv in byHeight)
                ed.WriteMessage($"\n      • Cao {kv.Key,-8} : {kv.Value,5} chuỗi");

            ed.WriteMessage("\n──────────────────────────────────────────────────────────");
            ed.WriteMessage($"\n  [4] CHUỖI TRÙNG LẶP (≥ 2 lần): {duplicates.Count} chuỗi");
            foreach (var kv in duplicates.Take(20))
            {
                string disp = kv.Key.Length > 45 ? kv.Key.Substring(0, 42) + "..." : kv.Key;
                ed.WriteMessage($"\n      • [{kv.Value,3}x]  {disp}");
            }
            if (duplicates.Count > 20)
                ed.WriteMessage($"\n      ... (xem thêm {duplicates.Count - 20} chuỗi trong bảng trên CAD)");

            ed.WriteMessage("\n══════════════════════════════════════════════════════════\n");
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    // Helper: Vẽ bảng thống kê trên CAD (dùng Line + DBText)
    // ══════════════════════════════════════════════════════════════════════════
    internal class TablePainter
    {
        private readonly BlockTableRecord _ms;
        private readonly Transaction      _tr;
        private readonly double           _txtH;
        private readonly double           _rowH;
        private readonly double           _padX;
        private readonly double           _padY;

        public TablePainter(BlockTableRecord ms, Transaction tr, double txtH)
        {
            _ms   = ms;
            _tr   = tr;
            _txtH = txtH;
            _rowH = txtH * 2.2;
            _padX = txtH * 0.5;
            _padY = txtH * 0.55;
        }

        /// <summary>
        /// Vẽ 1 bảng hoàn chỉnh. Trả về tổng chiều rộng bảng.
        /// </summary>
        public double DrawTable(
            double x0, double y0,
            string title,
            string[] colHeaders,
            double[] colWidths,
            List<string[]> rows)
        {
            double totalW = colWidths.Sum();
            int    nCols  = colHeaders.Length;

            // Tổng số hàng: 1 tiêu đề lớn + 1 header cột + n hàng dữ liệu
            int    totalRows = rows.Count + 2;
            double totalH    = _rowH * totalRows;

            // ── Viền ngoài ──────────────────────────────────────────────────
            AddRect(x0, y0, totalW, totalH);

            // ── Hàng 0: Tiêu đề lớn ─────────────────────────────────────────
            double y1 = y0 - _rowH; // dưới hàng tiêu đề
            AddHLine(x0, y1, totalW);
            AddCenteredText(title, x0 + totalW / 2, y1 + _padY, _txtH * 1.1);

            // ── Hàng 1: Header cột ───────────────────────────────────────────
            double y2 = y1 - _rowH; // dưới header
            AddHLine(x0, y2, totalW);

            double colX = x0;
            for (int c = 0; c < nCols; c++)
            {
                AddCenteredText(colHeaders[c], colX + colWidths[c] / 2, y2 + _padY, _txtH);
                if (c < nCols - 1)
                    AddVLine(colX + colWidths[c], y0, totalH); // đường dọc phân cột
                colX += colWidths[c];
            }

            // ── Hàng dữ liệu ────────────────────────────────────────────────
            for (int r = 0; r < rows.Count; r++)
            {
                double yTop = y2 - r * _rowH;
                double yBot = yTop - _rowH;

                if (r < rows.Count - 1)
                    AddHLine(x0, yBot, totalW);

                colX = x0;
                for (int c = 0; c < nCols; c++)
                {
                    string cell  = (rows[r].Length > c) ? rows[r][c] : "";
                    int maxChars = (int)(colWidths[c] / (_txtH * 0.62));
                    if (cell.Length > maxChars && maxChars > 4)
                        cell = cell.Substring(0, maxChars - 3) + "...";

                    // Cột đầu (STT) hoặc chỉ có 1 cột → căn giữa; còn lại → căn trái
                    bool center = (nCols == 1) || (c == 0 && nCols > 2) || (c == nCols - 1);
                    if (center)
                        AddCenteredText(cell, colX + colWidths[c] / 2, yBot + _padY, _txtH);
                    else
                        AddLeftText(cell, colX + _padX, yBot + _padY, _txtH);

                    colX += colWidths[c];
                }
            }

            return totalW;
        }

        // ── Tiện ích vẽ ─────────────────────────────────────────────────────

        private void AddRect(double x, double y, double w, double h)
        {
            // Top
            AddHLine(x, y, w);
            // Bottom
            AddHLine(x, y - h, w);
            // Left
            AddVLine(x, y, h);
            // Right
            AddVLine(x + w, y, h);
        }

        private void AddHLine(double x, double y, double w)
        {
            var ln = new Line(new Point3d(x, y, 0), new Point3d(x + w, y, 0));
            _ms.AppendEntity(ln);
            _tr.AddNewlyCreatedDBObject(ln, true);
        }

        private void AddVLine(double x, double yTop, double h)
        {
            var ln = new Line(new Point3d(x, yTop, 0), new Point3d(x, yTop - h, 0));
            _ms.AppendEntity(ln);
            _tr.AddNewlyCreatedDBObject(ln, true);
        }

        private void AddCenteredText(string content, double cx, double y, double h)
        {
            var t = new DBText
            {
                TextString     = content,
                Height         = h,
                HorizontalMode = TextHorizontalMode.TextCenter,
                VerticalMode   = TextVerticalMode.TextBase,
                AlignmentPoint = new Point3d(cx, y, 0)
            };
            _ms.AppendEntity(t);
            _tr.AddNewlyCreatedDBObject(t, true);
        }

        private void AddLeftText(string content, double x, double y, double h)
        {
            var t = new DBText
            {
                TextString = content,
                Height     = h,
                Position   = new Point3d(x, y, 0)
            };
            _ms.AppendEntity(t);
            _tr.AddNewlyCreatedDBObject(t, true);
        }
    }
}
