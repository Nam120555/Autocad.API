// LayerControl.cs - Điều khiển Layer ON/OFF cho Civil 3D Objects
// Chuyển đổi từ LISP: 3. ON OFF CIVIL 3D OBJECT.lsp

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.LayerControlCommands))]

namespace MyFirstProject
{
    /// <summary>
    /// Các lệnh điều khiển Layer ON/OFF cho Civil 3D objects
    /// </summary>
    public class LayerControlCommands
    {
        private static Editor GetEditor()
        {
            return AcadApp.DocumentManager.MdiActiveDocument?.Editor!;
        }

        private static void SetLayerState(string layerName, bool turnOn)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var db = doc.Database;
            var ed = doc.Editor;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                
                if (lt.Has(layerName))
                {
                    var layer = (LayerTableRecord)tr.GetObject(lt[layerName], OpenMode.ForWrite);
                    layer.IsOff = !turnOn;
                    ed.WriteMessage($"\n◎ Layer '{layerName}' đã {(turnOn ? "BẬT" : "TẮT")}.");
                }
                else
                {
                    ed.WriteMessage($"\n⊘ Layer '{layerName}' không tồn tại.");
                }
                
                tr.Commit();
            }
            ed.Regen();
        }

        // ══════════════════════════════════════════════════════════════
        // CORRIDOR SECTION
        // ══════════════════════════════════════════════════════════════
        [CommandMethod("CTL_OnCorridor")]
        public static void CTL_OnCorridor()
        {
            SetLayerState("C-ROAD-CORR-LINK", true);
        }

        [CommandMethod("CTL_OffCorridor")]
        public static void CTL_OffCorridor()
        {
            SetLayerState("C-ROAD-CORR-LINK", false);
        }

        // ══════════════════════════════════════════════════════════════
        // SAMPLE LINE
        // ══════════════════════════════════════════════════════════════
        [CommandMethod("CTL_OnSampleLine")]
        public static void CTL_OnSampleLine()
        {
            SetLayerState("C-ROAD-SAMP", true);
        }

        [CommandMethod("CTL_OffSampleLine")]
        public static void CTL_OffSampleLine()
        {
            SetLayerState("C-ROAD-SAMP", false);
        }

        // ══════════════════════════════════════════════════════════════
        // ALIGNMENT
        // ══════════════════════════════════════════════════════════════
        [CommandMethod("CTL_OnAlignment")]
        public static void CTL_OnAlignment()
        {
            SetLayerState("C-ROAD-CNTR-off", true);
        }

        [CommandMethod("CTL_OffAlignment")]
        public static void CTL_OffAlignment()
        {
            SetLayerState("C-ROAD-CNTR-off", false);
        }

        // ══════════════════════════════════════════════════════════════
        // PARCEL
        // ══════════════════════════════════════════════════════════════
        [CommandMethod("CTL_OnParcel")]
        public static void CTL_OnParcel()
        {
            SetLayerState("C-PROP-BNDY", true);
        }

        [CommandMethod("CTL_OffParcel")]
        public static void CTL_OffParcel()
        {
            SetLayerState("C-PROP-BNDY", false);
        }

        // ══════════════════════════════════════════════════════════════
        // HATCH ĐÀO ĐẮP
        // ══════════════════════════════════════════════════════════════
        [CommandMethod("CTL_OnHatchDaoDap")]
        public static void CTL_OnHatchDaoDap()
        {
            SetLayerState("$$$HATCH DAO DAP", true);
        }

        [CommandMethod("CTL_OffHatchDaoDap")]
        public static void CTL_OffHatchDaoDap()
        {
            SetLayerState("$$$HATCH DAO DAP", false);
        }

        // ══════════════════════════════════════════════════════════════
        // DEFPOINTS
        // ══════════════════════════════════════════════════════════════
        [CommandMethod("CTL_OnDefpoints")]
        public static void CTL_OnDefpoints()
        {
            SetLayerState("Defpoints", true);
        }

        [CommandMethod("CTL_OffDefpoints")]
        public static void CTL_OffDefpoints()
        {
            SetLayerState("Defpoints", false);
        }

        // ══════════════════════════════════════════════════════════════
        // HELPER: CHUYỂN ĐỐI TƯỢNG SANG LAYER KHÁC
        // ══════════════════════════════════════════════════════════════
        private static void MoveToLayer(string layerName)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var db = doc.Database;

            // Chọn đối tượng
            var psr = ed.GetSelection(new PromptSelectionOptions
            {
                MessageForAdding = $"\n⊙ Chọn đối tượng để chuyển sang layer '{layerName}': "
            });

            if (psr.Status != PromptStatus.OK) return;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                // Tạo layer nếu chưa có
                if (!lt.Has(layerName))
                {
                    lt.UpgradeOpen();
                    var newLayer = new LayerTableRecord { Name = layerName };
                    lt.Add(newLayer);
                    tr.AddNewlyCreatedDBObject(newLayer, true);
                    ed.WriteMessage($"\n  ✓ Đã tạo layer mới: {layerName}");
                }

                int count = 0;
                foreach (SelectedObject so in psr.Value)
                {
                    var ent = tr.GetObject(so.ObjectId, OpenMode.ForWrite) as Entity;
                    if (ent != null)
                    {
                        ent.Layer = layerName;
                        count++;
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n◎ Đã chuyển {count} đối tượng sang layer '{layerName}'.");
            }
        }

        // ══════════════════════════════════════════════════════════════
        // NHANH: CHUYỂN LAYER CHO BẢN VẼ KỸ THUẬT (từ LISP 0, 00, 11...)
        // ══════════════════════════════════════════════════════════════

        /// <summary>Chuyển sang layer 0.TEXT (Text thông thường)</summary>
        [CommandMethod("CTL_ToText")]
        public static void CTL_ToText() => MoveToLayer("0.TEXT");

        /// <summary>Chuyển sang layer Defpoints (không in)</summary>
        [CommandMethod("CTL_ToDefpoints")]
        public static void CTL_ToDefpoints() => MoveToLayer("Defpoints");

        /// <summary>Chuyển sang layer 1.DIM (Dim kích thước)</summary>
        [CommandMethod("CTL_ToDim")]
        public static void CTL_ToDim() => MoveToLayer("1.DIM");

        /// <summary>Chuyển sang layer 2.BAO BT (Bao bê tông)</summary>
        [CommandMethod("CTL_ToBaoBT")]
        public static void CTL_ToBaoBT() => MoveToLayer("2.BAO BT");

        /// <summary>Chuyển sang layer 3.BAO COT THEP (Bao cốt thép)</summary>
        [CommandMethod("CTL_ToBaoCotThep")]
        public static void CTL_ToBaoCotThep() => MoveToLayer("3.BAO COT THEP");

        /// <summary>Chuyển sang layer 4.THEP (Cốt thép)</summary>
        [CommandMethod("CTL_ToThep")]
        public static void CTL_ToThep() => MoveToLayer("4.THEP");

        /// <summary>Chuyển sang layer 5.TRUC (Trục)</summary>
        [CommandMethod("CTL_ToTruc")]
        public static void CTL_ToTruc() => MoveToLayer("5.TRUC");

        /// <summary>Chuyển sang layer 6.KHUAT (Nét khuất)</summary>
        [CommandMethod("CTL_ToKhuat")]
        public static void CTL_ToKhuat() => MoveToLayer("6.KHUAT");

        /// <summary>Chuyển sang layer 7.HATCH (Hatch)</summary>
        [CommandMethod("CTL_ToHatch")]
        public static void CTL_ToHatch() => MoveToLayer("7.HATCH");

        /// <summary>Chuyển sang layer 8.RANH GIOI (Ranh giới)</summary>
        [CommandMethod("CTL_ToRanhGioi")]
        public static void CTL_ToRanhGioi() => MoveToLayer("8.RANH GIOI");
    }
}
