using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Acad = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Civil = Autodesk.Civil.ApplicationServices;
using Autodesk.Civil;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using Autodesk.Civil.Runtime;
using Autodesk.Civil.Settings;
using ATable = Autodesk.AutoCAD.DatabaseServices.Table;
using CivSurface = Autodesk.Civil.DatabaseServices.TinSurface;
using Section = Autodesk.Civil.DatabaseServices.Section;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.CTS_DoiTenCocNX))]

namespace MyFirstProject
{
    public class CTS_DoiTenCocNX
    {
        private static readonly string[] Formats =
        {
            "Tiền tố + Km + m   (Km1+234.50)",
            "Km + m             (1+234.50)",
            "Tiền tố + lý trình (Km1234.50)"
        };

        /// <summary>Tạo tên cọc theo cấu hình trong hộp thoại.</summary>
        private static string BuildName(double station, string prefix, string suffix, int formatIndex, int decimals, bool pad)
        {
            var inv = System.Globalization.CultureInfo.InvariantCulture;
            station = Math.Round(station, decimals);
            string fmt = "F" + decimals;
            if (formatIndex == 2) return prefix + station.ToString(fmt, inv) + suffix;

            int km = (int)Math.Floor(station / 1000.0);
            double m = station - km * 1000.0;
            string mText = m.ToString(fmt, inv);
            if (pad)
            {
                int intLen = mText.IndexOf('.') < 0 ? mText.Length : mText.IndexOf('.');
                if (intLen < 3) mText = new string('0', 3 - intLen) + mText;
            }
            return (formatIndex == 0 ? prefix : "") + km + "+" + mText + suffix;
        }

        [CommandMethod("CTS_DoiTenCoc")]
        public static void CTSDoiTenCocMethod()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            // ── 1. Chọn cọc ──
            var pso = new PromptSelectionOptions { MessageForAdding = "\nChọn các cọc (Sample Line) cần đổi tên: " };
            PromptSelectionResult prompt = ed.GetSelection(pso,
                new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "AECC_SAMPLE_LINE") }));
            if (prompt.Status != PromptStatus.OK) return;

            // Lấy lý trình cọc đầu tiên để xem trước
            double sampleStation = 1234.5;
            try
            {
                using var trRead = db.TransactionManager.StartOpenCloseTransaction();
                var first = trRead.GetObject(prompt.Value[0].ObjectId, OpenMode.ForRead) as SampleLine;
                if (first != null) sampleStation = first.Station;
                trRead.Commit();
            }
            catch { /* dùng giá trị mẫu */ }

            // ── 2. Hộp thoại cấu hình ──
            var dlg = new MyFirstProject.UI.NxCommandDialog("CTS_DoiTenCoc", "Đổi tên cọc",
                $"Đã chọn {prompt.Value.Count} cọc — đặt tên theo lý trình", "⛳");
            dlg.AddSection("Định dạng tên cọc");
            dlg.AddChoice("format", "Kiểu tên", Formats, 0);
            dlg.AddText("prefix", "Tiền tố", "Km", hint: "Ví dụ: Km, C, TC");
            dlg.AddText("suffix", "Hậu tố", "", hint: "Để trống nếu không dùng (vd: _T, -P)");
            dlg.AddNumber("dec", "Số chữ số thập phân", 2, 0, 4, 0);
            dlg.AddCheck("pad", "Bù số 0 phần mét (1+5.00 → 1+005.00)", true);
            dlg.SetPreview(d =>
            {
                int fi = Math.Max(0, Array.IndexOf(Formats, d.GetString("format")));
                string a = BuildName(sampleStation, d.GetString("prefix"), d.GetString("suffix"), fi, d.GetInt("dec"), d.GetBool("pad"));
                string b = BuildName(1005.0, d.GetString("prefix"), d.GetString("suffix"), fi, d.GetInt("dec"), d.GetBool("pad"));
                return $"Cọc đầu tiên ({sampleStation:F3})  →  {a}\nVí dụ lý trình 1005.000  →  {b}";
            });
            dlg.SetValidator(d =>
            {
                string name = d.GetString("prefix") + d.GetString("suffix");
                return name.IndexOfAny(new[] { '<', '>', '/', '\\', '"', ':', ';', '?', '*', '|', ',', '=', '`' }) >= 0
                    ? "Tiền tố / hậu tố chứa ký tự không hợp lệ cho tên cọc." : null;
            });
            dlg.SetHelp("Đổi tên hàng loạt Sample Line theo lý trình.\n\n" +
                        "• Chọn kiểu tên, tiền tố, số lẻ thập phân.\n" +
                        "• Khung \"Xem trước\" hiển thị tên của cọc đầu tiên.\n" +
                        "• Cọc nào trùng tên trong nhóm sẽ được bỏ qua và báo lại.");
            dlg.SetRunText("▶  Đổi tên");
            if (!dlg.ShowAcad()) return;

            string prefix = dlg.GetString("prefix");
            string suffix = dlg.GetString("suffix");
            int formatIndex = Math.Max(0, Array.IndexOf(Formats, dlg.GetString("format")));
            int decimals = dlg.GetInt("dec");
            bool pad = dlg.GetBool("pad");

            // ── 3. Thực hiện ──
            try
            {
                using Transaction tr = db.TransactionManager.StartTransaction();
                int count = 0;
                var failed = new List<string>();
                foreach (SelectedObject obj in prompt.Value)
                {
                    // LUÔN DÙNG ForWrite THEO CHUẨN ĐỂ CHỐNG CRASH
                    var sl = (SampleLine)tr.GetObject(obj.ObjectId, OpenMode.ForWrite);
                    string newName = BuildName(sl.Station, prefix, suffix, formatIndex, decimals, pad);
                    try
                    {
                        sl.Name = newName;
                        count++;
                    }
                    catch (System.Exception)
                    {
                        failed.Add(newName);   // thường do trùng tên trong cùng nhóm cọc
                    }
                }
                tr.Commit();
                ed.WriteMessage($"\n[CTS_DoiTenCoc] ✓ Đã đổi tên {count}/{prompt.Value.Count} cọc.");
                if (failed.Count > 0)
                {
                    MyFirstProject.UI.NxMessageBox.Show(MyFirstProject.UI.NxMessageKind.Warning, "Một số cọc chưa đổi được tên",
                        $"Đã đổi {count} cọc. {failed.Count} cọc bị bỏ qua (có thể trùng tên trong nhóm cọc).",
                        "Tên không đặt được:\n" + string.Join("\n", failed));
                }
            }
            catch (System.Exception e)
            {
                MyFirstProject.UI.NxMessageBox.Error("Lỗi đổi tên cọc", e.Message, e.ToString());
            }
        }
    }
}
