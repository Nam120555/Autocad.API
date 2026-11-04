using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace MyFirstProject.CivilTool
{
    public class Convert2PolylineTool
    {
        [CommandMethod("CVT2PL")]
        public void ShowConvertForm()
        {
            var form = new Convert2PolylineForm();
            // Hiển thị Form dưới dạng Modal Dialog kết nối với giao diện AutoCAD
            Application.ShowModalDialog(Application.MainWindow.Handle, form, false);
        }
    }

    public class Convert2PolylineForm : Form
    {
        private ComboBox cbbLayer = null!;
        private NumericUpDown numColor = null!;
        private NumericUpDown numWidth = null!;
        private CheckBox chkFlatten = null!;
        private Button btnConvert = null!;
        private Button btnCancel = null!;

        public Convert2PolylineForm()
        {
            InitializeComponents();
            LoadLayers();

            // Nếu người dùng yêu cầu đóng / watermark thì có thể thêm ở đây
            // FormSignatureHelper.ApplySignature(this);
        }

        private void InitializeComponents()
        {
            this.Text = "Convert To Polyline Tool";
            this.Size = new Size(350, 260);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;

            Label lblLayer = new Label { Text = "Layer Mục Tiêu:", Location = new Point(20, 20), AutoSize = true };
            cbbLayer = new ComboBox { Location = new Point(140, 15), Width = 170, DropDownStyle = ComboBoxStyle.DropDownList };

            Label lblColor = new Label { Text = "Màu Sắc (0-256):", Location = new Point(20, 60), AutoSize = true };
            numColor = new NumericUpDown { Location = new Point(140, 55), Width = 170, Minimum = 0, Maximum = 256, Value = 256 };
            Label lblColorInfo = new Label { Text = "(256 = ByLayer, 0 = ByBlock)", Location = new Point(135, 80), AutoSize = true, ForeColor = Color.Gray, Font = new System.Drawing.Font(this.Font.FontFamily, 7.5f) };

            Label lblWidth = new Label { Text = "Bề rộng Nét:", Location = new Point(20, 105), AutoSize = true };
            numWidth = new NumericUpDown { Location = new Point(140, 100), Width = 170, DecimalPlaces = 2, Minimum = 0, Maximum = 1000, Value = 0 };

            chkFlatten = new CheckBox { Text = "Ép Cao độ (Elevation) về 0", Location = new Point(140, 130), AutoSize = true, Checked = true };

            btnConvert = new Button { Text = "Chuyển Đổi", Location = new Point(140, 170), Width = 80, BackColor = Color.LightBlue };
            btnConvert.Click += BtnConvert_Click;

            btnCancel = new Button { Text = "Huỷ", Location = new Point(230, 170), Width = 80 };
            btnCancel.Click += (s, e) => { this.Close(); };

            this.Controls.Add(lblLayer);
            this.Controls.Add(cbbLayer);
            this.Controls.Add(lblColor);
            this.Controls.Add(numColor);
            this.Controls.Add(lblColorInfo);
            this.Controls.Add(lblWidth);
            this.Controls.Add(numWidth);
            this.Controls.Add(chkFlatten);
            this.Controls.Add(btnConvert);
            this.Controls.Add(btnCancel);
        }

        private void LoadLayers()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId lId in lt)
                {
                    LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(lId, OpenMode.ForRead);
                    cbbLayer.Items.Add(ltr.Name);
                    
                    if (ltr.Name == "0") cbbLayer.SelectedItem = "0";
                }
                tr.Commit();
            }

            if (cbbLayer.Items.Count > 0 && cbbLayer.SelectedIndex == -1)
                cbbLayer.SelectedIndex = 0;
        }

        private void BtnConvert_Click(object? sender, EventArgs e)
        {
            this.Hide(); // Ẩn form để lấy Selection
            string layer = cbbLayer.SelectedItem as string ?? "0";
            ExecuteConversion(layer, (int)numColor.Value, (double)numWidth.Value, chkFlatten.Checked);
            this.Close();
        }

        private void ExecuteConversion(string targetLayer, int colorIndex, double width, bool flattenZ)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Bộ lọc: LINE, POLYLINE (tương đương 2d và 3d), CIRCLE
            TypedValue[] filterItems = new TypedValue[] 
            {
                new TypedValue((int)DxfCode.Start, "LINE,POLYLINE,CIRCLE")
            };
            SelectionFilter filter = new SelectionFilter(filterItems);

            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nChọn các Line, 3D Polyline hoặc Circle để chuyển đổi: ";
            
            PromptSelectionResult selRes = ed.GetSelection(pso, filter);
            if (selRes.Status != PromptStatus.OK) return;

            int convertedCount = 0;

            using (DocumentLock dl = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                    foreach (SelectedObject so in selRes.Value)
                    {
                        Entity? ent = tr.GetObject(so.ObjectId, OpenMode.ForWrite) as Entity;
                        if (ent == null) continue;

                        Polyline newLwPoly = new Polyline();
                        bool isConverted = false;

                        // Case 1: LINE
                        if (ent is Line line)
                        {
                            newLwPoly.AddVertexAt(0, new Point2d(line.StartPoint.X, line.StartPoint.Y), 0, 0, 0);
                            newLwPoly.AddVertexAt(1, new Point2d(line.EndPoint.X, line.EndPoint.Y), 0, 0, 0);
                            isConverted = true;
                        }
                        // Case 2: 3D POLYLINE
                        else if (ent is Polyline3d poly3d)
                        {
                            int index = 0;
                            foreach (ObjectId vertexId in poly3d)
                            {
                                PolylineVertex3d v3d = (PolylineVertex3d)tr.GetObject(vertexId, OpenMode.ForRead);
                                newLwPoly.AddVertexAt(index, new Point2d(v3d.Position.X, v3d.Position.Y), 0, 0, 0);
                                index++;
                            }
                            newLwPoly.Closed = poly3d.Closed;
                            isConverted = true;
                        }
                        // Case 3: 2D POLYLINE
                        else if (ent is Polyline2d poly2d)
                        {
                            int index = 0;
                            foreach (ObjectId vertexId in poly2d)
                            {
                                Vertex2d v2d = (Vertex2d)tr.GetObject(vertexId, OpenMode.ForRead);
                                newLwPoly.AddVertexAt(index, new Point2d(v2d.Position.X, v2d.Position.Y), v2d.Bulge, 0, 0); 
                                index++;
                            }
                            newLwPoly.Closed = poly2d.Closed;
                            isConverted = true;
                        }
                        // Case 4: CIRCLE -> Thành Polyline có 2 nửa đường tròn (Arc) nối nhau
                        else if (ent is Circle cir)
                        {
                            Point2d center = new Point2d(cir.Center.X, cir.Center.Y);
                            double radius = cir.Radius;
                            
                            // Tạo 2 điểm ở 0 độ và 180 độ
                            Point2d pt1 = center + new Vector2d(radius, 0);
                            Point2d pt2 = center + new Vector2d(-radius, 0);

                            // Bulge = 1 tương đương với Cung tròn quét góc 180 (Bán nguyệt)
                            newLwPoly.AddVertexAt(0, pt1, 1.0, 0, 0); 
                            newLwPoly.AddVertexAt(1, pt2, 1.0, 0, 0);
                            newLwPoly.Closed = true;
                            
                            isConverted = true;
                        }

                        if (isConverted)
                        {
                            // Áp dụng các thông số đã Setup
                            if (flattenZ) newLwPoly.Elevation = 0.0;
                            newLwPoly.ConstantWidth = width;
                            newLwPoly.ColorIndex = colorIndex;
                            
                            try { newLwPoly.Layer = targetLayer; } 
                            catch { /* Bỏ qua nếu Layer có vấn đề (Bị khóa...) */ }

                            // Append to Drawing
                            modelSpace.AppendEntity(newLwPoly);
                            tr.AddNewlyCreatedDBObject(newLwPoly, true);

                            // Erase The Old Entity
                            ent.Erase(true);
                            convertedCount++;
                        }
                    }
                    tr.Commit();
                }
            }

            ed.WriteMessage($"\n--- [CIVIL TOOL INFO] ---");
            ed.WriteMessage($"\nĐã chuyển đổi thành công {convertedCount} đối tượng.");
        }
    }
}
