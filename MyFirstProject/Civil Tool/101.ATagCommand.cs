using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using MyFirstProject.Extensions;

// Alias to avoid ambiguity
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using ContentAlignment = System.Drawing.ContentAlignment;
using FontStyle = System.Drawing.FontStyle;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.ATagCommand))]

namespace MyFirstProject
{
    #region Form

    public class ATagForm : CivilFormBase
    {
        private WinLabel lblBlockCount = null!;
        private WinLabel lblPathInfo = null!;
        private ComboBox cmbSide = null!;
        private NumericUpDown numOffset = null!;
        private NumericUpDown numTextSize = null!;
        
        public List<ObjectId> SelectedBlockIds { get; private set; } = new();
        public ObjectId SelectedPathId { get; private set; } = ObjectId.Null;
        public string Side { get; private set; } = "Trái";
        public double Offset => (double)numOffset.Value;
        public double TextSize => (double)numTextSize.Value;
        public bool IsConfirmed { get; private set; } = false;

        public ATagForm()
        {
            InitializeUI();
        }

        private void InitializeUI()
        {
            InitBaseTheme("A_TAG - Đóng Tuyến & Gắn Tag Trụ", 500, 480);

            var header = CreateHeader("CÔNG CỤ GẮN TAG TRỤ (A_TAG)");
            this.Controls.Add(header);

            int y = 70;

            // --- Chọn dữ liệu ---
            var grpSelection = CreateSection("Chọn đối tượng", 15, y, 460, 100);
            this.Controls.Add(grpSelection);

            var btnPickBlocks = CreateButton("📍 Chọn Block Trụ", 15, 30, 140, 32, grpSelection);
            btnPickBlocks.Click += (s, e) => PickBlocks();
            lblBlockCount = CreateLabel("Chưa chọn block nào", 165, 35);
            grpSelection.Controls.Add(lblBlockCount);

            var btnPickPath = CreateButton("🛣 Chọn Tim Tuyến", 15, 65, 140, 32, grpSelection);
            btnPickPath.Click += (s, e) => PickPath();
            lblPathInfo = CreateLabel("Chưa chọn tim tuyến", 165, 70);
            grpSelection.Controls.Add(lblPathInfo);

            y += 110;

            // --- Cấu hình ---
            var grpConfig = CreateSection("Cấu hình Tag", 15, y, 460, 150);
            this.Controls.Add(grpConfig);

            grpConfig.Controls.Add(CreateLabel("Hướng Offset:", 15, 35));
            cmbSide = new ComboBox
            {
                Location = new Point(140, 32),
                Width = 120,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat
            };
            cmbSide.Items.AddRange(new object[] { "Trái", "Phải" });
            cmbSide.SelectedIndex = 0;
            grpConfig.Controls.Add(cmbSide);

            grpConfig.Controls.Add(CreateLabel("Khoảng cách Offset:", 15, 75));
            numOffset = CreateNumeric(140, 72, 120, 0, 1000, 5, 1);
            grpConfig.Controls.Add(numOffset);

            grpConfig.Controls.Add(CreateLabel("Kích thước Attribute:", 15, 115));
            numTextSize = CreateNumeric(140, 112, 120, 0, 100, 2.5m, 1);
            grpConfig.Controls.Add(numTextSize);

            y += 160;

            // --- Nút bấm ---
            var pnlBottom = CreateBottomBar();
            this.Controls.Add(pnlBottom);

            var btnCancel = CreateButton("Hủy", 230, 8, 100, 34, pnlBottom);
            btnCancel.Click += (s, e) => this.Close();

            var btnExecute = CreatePrimaryButton("Thực hiện", 340, 8, 140, 34, pnlBottom);
            btnExecute.Click += (s, e) =>
            {
                if (SelectedBlockIds.Count == 0) { MessageBox.Show("Vui lòng chọn block trụ!"); return; }
                if (SelectedPathId == ObjectId.Null) { MessageBox.Show("Vui lòng chọn tim tuyến!"); return; }
                Side = cmbSide.SelectedItem?.ToString() ?? "Trái";
                IsConfirmed = true;
                this.DialogResult = DialogResult.OK;
                this.Close();
            };
        }

        private void PickBlocks()
        {
            this.Hide();
            var ed = A.Ed;
            var sel = ed.GetSelection(new PromptSelectionOptions { MessageForAdding = "\nQuét chọn các Block trụ: " }, 
                new SelectionFilter(new TypedValue[] { new TypedValue(0, "INSERT") }));
            
            if (sel.Status == PromptStatus.OK)
            {
                SelectedBlockIds = sel.Value.GetObjectIds().ToList();
                lblBlockCount.Text = $"Đã chọn {SelectedBlockIds.Count} block";
                lblBlockCount.ForeColor = GreenOk;
            }
            this.Show();
        }

        private void PickPath()
        {
            this.Hide();
            var ed = A.Ed;
            var opt = new PromptEntityOptions("\nChọn tim tuyến (Alignment/Polyline): ");
            opt.SetRejectMessage("\nChỉ chọn Alignment hoặc Polyline!");
            opt.AddAllowedClass(typeof(Alignment), true);
            opt.AddAllowedClass(typeof(Polyline), true);
            opt.AddAllowedClass(typeof(Polyline2d), true);
            opt.AddAllowedClass(typeof(Polyline3d), true);

            var res = ed.GetEntity(opt);
            if (res.Status == PromptStatus.OK)
            {
                SelectedPathId = res.ObjectId;
                using var tr = A.Db.TransactionManager.StartTransaction();
                var ent = tr.GetObject(SelectedPathId, OpenMode.ForRead);
                lblPathInfo.Text = (ent is Alignment) ? $"Tuyến: {((Alignment)ent).Name}" : "Tuyến: Polyline";
                lblPathInfo.ForeColor = GreenOk;
                tr.Commit();
            }
            this.Show();
        }
    }

    #endregion

    #region Command

    public class ATagCommand
    {
        [CommandMethod("A_TAG")]
        public static void CTS_ATag()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using var form = new ATagForm();
            if (AcadApp.ShowModalDialog(form) != DialogResult.OK) return;

            using var tr = db.TransactionManager.StartTransaction();
            try
            {
                // 1. Kiểm tra Block TAG_TRU
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (!bt.Has("TAG_TRU"))
                {
                    MessageBox.Show("Lỗi: Không tìm thấy Block 'TAG_TRU' trong bản vẽ!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                ObjectId tagBlockId = bt["TAG_TRU"];

                // 2. Lấy Curve (Tim tuyến)
                Curve? path = tr.GetObject(form.SelectedPathId, OpenMode.ForRead) as Curve;
                if (path == null) return;

                // 3. Chuẩn bị Layer dóng tuyến
                CreateLayerIfNotExist(db, tr, "NNT_dongtuyen", 3);

                // 4. Thu thập dữ liệu và tính lý trình
                var dataList = new List<PierData>();
                foreach (var id in form.SelectedBlockIds)
                {
                    BlockReference? blk = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    if (blk == null) continue;

                    Point3d insPt = blk.Position;
                    Point3d perpPt = path.GetClosestPointTo(insPt, false);
                    double dist = path.GetDistAtPoint(perpPt);
                    
                    dataList.Add(new PierData { BlockId = id, Position = insPt, PerpPoint = perpPt, Distance = dist });
                }

                // 5. Sắp xếp theo lý trình
                dataList = dataList.OrderBy(x => x.Distance).ToList();

                // 6. Xử lý gắn Tag
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                int currentKm = -1;
                int seq = 1;

                foreach (var item in dataList)
                {
                    // Vẽ đường dóng
                    Line line = new Line(item.Position, item.PerpPoint);
                    line.Layer = "NNT_dongtuyen";
                    btr.AppendEntity(line);
                    tr.AddNewlyCreatedDBObject(line, true);

                    // Tính KM và STT
                    int kmVal = (int)Math.Floor(item.Distance / 1000.0);
                    if (kmVal == currentKm) seq++;
                    else { currentKm = kmVal; seq = 1; }

                    string kmStr = "P" + kmVal;
                    string sttStr = seq.ToString("D3");

                    // Tính vị trí chèn Tag
                    Vector3d deriv = path.GetFirstDerivative(item.PerpPoint);
                    double tangentAngle = deriv.AngleOnPlane(new Plane(Point3d.Origin, Vector3d.ZAxis));
                    
                    double offsetAngle = (form.Side == "Trái") ? tangentAngle + Math.PI / 2.0 : tangentAngle - Math.PI / 2.0;
                    Point3d tagInsPt = item.Position.Add(new Vector3d(Math.Cos(offsetAngle), Math.Sin(offsetAngle), 0) * form.Offset);

                    // Tính góc xoay (luôn đọc được từ dưới lên/phải qua)
                    double degAngle = tangentAngle * 180.0 / Math.PI;
                    if (degAngle > 90 && degAngle <= 270) degAngle -= 180;

                    // Chèn Block Tag
                    BlockReference tagRef = new BlockReference(tagInsPt, tagBlockId);
                    tagRef.Rotation = degAngle * Math.PI / 180.0;
                    btr.AppendEntity(tagRef);
                    tr.AddNewlyCreatedDBObject(tagRef, true);

                    // Gán Attributes
                    BlockTableRecord tagDef = (BlockTableRecord)tr.GetObject(tagBlockId, OpenMode.ForRead);
                    foreach (ObjectId attId in tagDef)
                    {
                        AttributeDefinition? attDef = tr.GetObject(attId, OpenMode.ForRead) as AttributeDefinition;
                        if (attDef != null && !attDef.Constant)
                        {
                            using AttributeReference attRef = new AttributeReference();
                            attRef.SetAttributeFromBlock(attDef, tagRef.BlockTransform);
                            
                            if (attRef.Tag.Equals("KM", StringComparison.OrdinalIgnoreCase)) attRef.TextString = kmStr;
                            else if (attRef.Tag.Equals("STT", StringComparison.OrdinalIgnoreCase)) attRef.TextString = sttStr;
                            
                            // Kích thước Attribute
                            attRef.Height = form.TextSize;

                            tagRef.AttributeCollection.AppendAttribute(attRef);
                            tr.AddNewlyCreatedDBObject(attRef, true);
                        }
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n[Thành công] Đã gắn Tag cho {dataList.Count} trụ.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[Lỗi] {ex.Message}");
            }
        }

        private static void CreateLayerIfNotExist(Database db, Transaction tr, string name, short color)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
            if (!lt.Has(name))
            {
                LayerTableRecord ltr = new LayerTableRecord { Name = name, Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, color) };
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }

        private class PierData
        {
            public ObjectId BlockId { get; set; }
            public Point3d Position { get; set; }
            public Point3d PerpPoint { get; set; }
            public double Distance { get; set; }
        }
    }

    #endregion
}
