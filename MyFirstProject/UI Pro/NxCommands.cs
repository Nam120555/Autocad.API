// ==============================================================================
// NxCommands.cs — Lệnh của bộ giao diện UI Pro
//   NX / CT_HopLenh     : mở Hộp lệnh (tìm & chạy mọi lệnh Civil Tool)
//   CT_DemoHopThoai     : hộp thoại mẫu (chạy được trên AutoCAD thường)
// ==============================================================================
using System;
using System.Collections.Generic;
using System.Linq;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.UI.NxCommands))]

namespace MyFirstProject.UI
{
    public class NxCommands
    {
        [CommandMethod("NX")]
        public static void OpenCommandCenter() => NxCommandCenter.ShowCenter();

        [CommandMethod("CT_HopLenh")]
        public static void OpenCommandCenter2() => NxCommandCenter.ShowCenter();

        /// <summary>
        /// Hộp thoại mẫu: vẽ vòng tròn + ghi chú tại điểm chọn.
        /// Dùng để xem trước phong cách "box lệnh" và làm khuôn cho lệnh mới.
        /// </summary>
        [CommandMethod("CT_DemoHopThoai")]
        public static void DemoDialog()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;

            // Lấy danh sách layer của bản vẽ
            var layers = new List<string>();
            using (var tr = db.TransactionManager.StartOpenCloseTransaction())
            {
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId id in lt)
                    layers.Add(((LayerTableRecord)tr.GetObject(id, OpenMode.ForRead)).Name);
                tr.Commit();
            }
            layers.Sort(StringComparer.OrdinalIgnoreCase);
            string curLayer;
            using (var tr = db.TransactionManager.StartOpenCloseTransaction())
            {
                curLayer = ((LayerTableRecord)tr.GetObject(db.Clayer, OpenMode.ForRead)).Name;
                tr.Commit();
            }

            var colors = new[] { "ByLayer", "Đỏ (1)", "Vàng (2)", "Xanh lá (3)", "Cyan (4)", "Xanh dương (5)", "Tím (6)", "Trắng (7)" };

            var dlg = new NxCommandDialog("CT_DemoHopThoai", "Vẽ vòng tròn ghi chú",
                "Hộp thoại mẫu UI Pro — chọn tâm, bán kính, layer rồi bấm Thực hiện", "◎");

            dlg.AddSection("Vị trí & kích thước", "Bấm nút ⌖ để lấy trực tiếp trên bản vẽ");
            dlg.AddPickPoint("center", "Tâm vòng tròn", "Chọn tâm vòng tròn");
            dlg.AddPickDistance("radius", "Bán kính", 5.0, "Chỉ định bán kính", min: 0.001, decimals: 3, unit: "m");

            dlg.AddSection("Hiển thị");
            dlg.AddChoice("layer", "Layer", layers, Math.Max(0, layers.IndexOf(curLayer)));
            dlg.AddChoice("color", "Màu", colors, 0);
            dlg.AddCheck("label", "Ghi chú bán kính bên cạnh", true);
            dlg.AddNumber("textH", "Chiều cao chữ", 2.5, 0.1, 100, 2, "m");
            dlg.AddText("prefix", "Tiền tố ghi chú", "R = ", hint: "Ví dụ: \"R = \", \"Bán kính \"");

            dlg.SetPreview(d =>
            {
                double r = d.GetDouble("radius");
                string txt = d.GetBool("label") ? $"{d.GetString("prefix")}{r:F3}" : "(không ghi chú)";
                return $"Vòng tròn R={r:F3}  ·  Chu vi={2 * Math.PI * r:F3}  ·  Diện tích={Math.PI * r * r:F3}\nGhi chú: {txt}";
            });
            dlg.SetHelp("Lệnh mẫu của bộ giao diện UI Pro.\n\n" +
                        "1. Bấm ⌖ cạnh \"Tâm vòng tròn\" để chọn điểm.\n" +
                        "2. Nhập bán kính hoặc bấm ⌖ để đo trên bản vẽ.\n" +
                        "3. Chọn layer, màu, ghi chú → Thực hiện.\n\n" +
                        "Mọi giá trị được nhớ lại cho lần sau.");

            if (!dlg.ShowAcad()) return;

            var center = dlg.GetPoint("center")!.Value;
            double radius = dlg.GetDouble("radius");
            short colorIndex = (short)Math.Max(0, Array.IndexOf(colors, dlg.GetString("color")));
            string layer = dlg.GetString("layer");

            try
            {
                using var lockDoc = doc.LockDocument();
                using var tr = db.TransactionManager.StartTransaction();
                var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                var circle = new Circle(center, Vector3d.ZAxis, radius) { Layer = layer };
                circle.ColorIndex = colorIndex == 0 ? 256 : colorIndex;
                ms.AppendEntity(circle);
                tr.AddNewlyCreatedDBObject(circle, true);

                if (dlg.GetBool("label"))
                {
                    var text = new DBText
                    {
                        Position = new Point3d(center.X + radius * 0.75, center.Y + radius * 0.75, center.Z),
                        Height = dlg.GetDouble("textH"),
                        TextString = $"{dlg.GetString("prefix")}{radius:F3}",
                        Layer = layer,
                        TextStyleId = db.Textstyle
                    };
                    text.ColorIndex = circle.ColorIndex;
                    ms.AppendEntity(text);
                    tr.AddNewlyCreatedDBObject(text, true);
                }
                tr.Commit();
                doc.Editor.WriteMessage($"\n[CT_DemoHopThoai] ✓ Đã vẽ vòng tròn R={radius:F3} trên layer \"{layer}\".");
            }
            catch (System.Exception ex)
            {
                NxMessageBox.Error("Không vẽ được vòng tròn", ex.Message, ex.ToString());
            }
        }
    }
}
