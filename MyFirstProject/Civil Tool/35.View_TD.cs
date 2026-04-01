using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Acad = Autodesk.AutoCAD.ApplicationServices;
using WinLabel = System.Windows.Forms.Label;

[assembly: CommandClass(typeof(MyFirstProject.View_TD))]

namespace MyFirstProject
{
    public class View_TD
    {
        // Biến lưu trữ cấu hình tĩnh
        public static string SelectedBlockName = "";
        public static double MViewWidth = 297;
        public static double MViewHeight = 210;
        public static double MViewSpacingX = 10;
        public static double MViewSpacingY = 10;
        public static double StartX = 10;
        public static double StartY = 200; // Tọa độ Y góc dưới trái của hàng trên cùng
        public static int Columns = 1;
        public static int Rows = 2; // Ví dụ 1 cột 2 hàng
        public static double ScaleVal = 1000;

        [CommandMethod("CTSV_TaoTrangInTD")]
        public void CTSVTaoTrangInTD()
        {
            var doc = Acad.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // Xoay vòng lặp nếu người dùng nhấn "Vẽ Khung" và muốn quay lại Form
            bool keepLooping = true;
            while (keepLooping)
            {
                List<string> blockNames = new List<string>();
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        // Chỉ lấy block thường, không phải layout, không anonymous
                        if (!btr.IsLayout && !btr.IsAnonymous)
                        {
                            blockNames.Add(btr.Name);
                        }
                    }
                    tr.Commit();
                }
                blockNames.Sort();

                using (var form = new MainForm_ViewTD(blockNames))
                {
                    var res = Acad.Application.ShowModalDialog(form);
                    if (res == DialogResult.Retry) // Chọn Vẽ khung Model
                    {
                        DrawFrames(db, ed);
                        // Vòng lặp sẽ tiếp tục hiển thị lại form
                    }
                    else if (res == DialogResult.OK) // Chọn Tạo Layout
                    {
                        CreateLayouts(db, ed);
                        keepLooping = false; // Xong việc thì thoát
                    }
                    else // Đóng form (Cancel)
                    {
                        keepLooping = false;
                    }
                }
            }
        }

        private void DrawFrames(Database db, Editor ed)
        {
            // Kích thước khung trên Model = MView x Scale
            double frameW = MViewWidth * (ScaleVal / 1000.0);
            double frameH = MViewHeight * (ScaleVal / 1000.0);

            ed.WriteMessage($"\n--- VẼ KHỐI KHUNG MVIEW TRÊN MODEL (Rộng: {frameW}, Cao: {frameH}) ---");

            while (true)
            {
                var ppo = new PromptPointOptions("\nChọn điểm chèn góc dưới bên trái của Khung (hoặc ESC để quay lại): ");
                var res = ed.GetPoint(ppo);
                if (res.Status != PromptStatus.OK) break;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // Tạo layer nếu chưa có
                    string layerName = "Khung_TD";
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(layerName))
                    {
                        lt.UpgradeOpen();
                        var ltr = new LayerTableRecord();
                        ltr.Name = layerName;
                        ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1); // Đỏ
                        lt.Add(ltr);
                        tr.AddNewlyCreatedDBObject(ltr, true);
                    }

                    var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    var pline = new Polyline();
                    Point3d p = res.Value;
                    pline.AddVertexAt(0, new Point2d(p.X, p.Y), 0, 0, 0);
                    pline.AddVertexAt(1, new Point2d(p.X + frameW, p.Y), 0, 0, 0);
                    pline.AddVertexAt(2, new Point2d(p.X + frameW, p.Y + frameH), 0, 0, 0);
                    pline.AddVertexAt(3, new Point2d(p.X, p.Y + frameH), 0, 0, 0);
                    pline.Closed = true;
                    pline.Layer = layerName;

                    btr.AppendEntity(pline);
                    tr.AddNewlyCreatedDBObject(pline, true);
                    tr.Commit();
                }
            }
        }

        private void CreateLayouts(Database db, Editor ed)
        {
            var pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nChọn các khung Polyline trên Model để tạo Layout (Nhấn Enter khi xong): ";
            var selRes = ed.GetSelection(pso);
            if (selRes.Status != PromptStatus.OK) return;

            List<Polyline> frames = new List<Polyline>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject so in selRes.Value)
                {
                    var ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Polyline;
                    // Phải là polyline có ít nhất 4 đỉnh và đóng kín (hoặc vẽ theo hàm DrawFrames)
                    if (ent != null && ent.NumberOfVertices >= 4)
                    {
                        frames.Add(ent);
                    }
                }

                if (frames.Count == 0)
                {
                    ed.WriteMessage("\nKhông có khung nào hợp lệ được chọn.");
                    tr.Commit();
                    return;
                }

                // Sắp xếp các khung theo thứ tự X (từ trái qua phải)
                frames.Sort((a, b) => a.GetPoint3dAt(0).X.CompareTo(b.GetPoint3dAt(0).X));

                int framesPerLayout = Columns * Rows;
                int numLayouts = (int)Math.Ceiling(frames.Count / (double)framesPerLayout);

                var layoutMgr = LayoutManager.Current;

                int frameIdx = 0;
                for (int i = 0; i < numLayouts; i++)
                {
                    string layoutName = "Layout_TD_" + (i + 1);
                    
                    // Tạo layout mới hoặc lấy layout cũ
                    ObjectId layoutId = layoutMgr.GetLayoutId(layoutName);
                    if (layoutId.IsNull)
                    {
                        layoutId = layoutMgr.CreateLayout(layoutName);
                    }
                    
                    Layout layout = (Layout)tr.GetObject(layoutId, OpenMode.ForWrite);
                    layoutMgr.CurrentLayout = layoutName;
                    
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

                    // Xóa tất cả các viewport hiện có trong layout này (trừ viewport mặc định số 1 thường là khung giấy)
                    // Ở đây, xóa mọi viewports cũ để thiết lập lại
                    foreach (ObjectId id in btr)
                    {
                        var ent = tr.GetObject(id, OpenMode.ForWrite);
                        if (ent is Viewport vp && vp.Number != 1)
                        {
                            ent.Erase();
                        }
                    }

                    // Chèn khung tên (Nút gốc 0,0)
                    if (!string.IsNullOrEmpty(SelectedBlockName) && SelectedBlockName != "<Không dùng>")
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        if (bt.Has(SelectedBlockName))
                        {
                            var bref = new BlockReference(Point3d.Origin, bt[SelectedBlockName]);
                            btr.AppendEntity(bref);
                            tr.AddNewlyCreatedDBObject(bref, true);
                        }
                    }

                    // Tạo các MView cho Layout này
                    int col = 0;
                    int row = 0;
                    // StartY là cạnh Dưới của Viewport tại Hàng 0. Viewport tâm Y sẽ là StartY + MViewHeight/2.
                    // Nhưng tọa độ Y từ trên xuống dưới trên khổ giấy, Rows sẽ đi xuống (trừ Y).
                    for (int j = 0; j < framesPerLayout && frameIdx < frames.Count; j++, frameIdx++)
                    {
                        var frame = frames[frameIdx];

                        // Xác định tâm của khung trên Model
                        // Do khung bên model có thể được vẽ bằng tool (đỉnh 0 là góc dưới trái, 2 là trên phải)
                        // Ta tính Min/Max để an toàn hơn
                        double minX = double.MaxValue, minY = double.MaxValue;
                        double maxX = double.MinValue, maxY = double.MinValue;
                        for (int v = 0; v < frame.NumberOfVertices; v++)
                        {
                            Point3d pt = frame.GetPoint3dAt(v);
                            if (pt.X < minX) minX = pt.X;
                            if (pt.Y < minY) minY = pt.Y;
                            if (pt.X > maxX) maxX = pt.X;
                            if (pt.Y > maxY) maxY = pt.Y;
                        }
                        double centerModelX = (minX + maxX) / 2.0;
                        double centerModelY = (minY + maxY) / 2.0;

                        // Tọa độ tâm MView trên Layout
                        // X tăng dần sang phải
                        double centerX = StartX + col * (MViewWidth + MViewSpacingX) + (MViewWidth / 2.0);
                        // Y giảm dần xuống dưới (Giả sử StartY là mép dưới của khung đầu tiên)
                        // StartY là điểm cơ sở (bottom-left) của Layout Grid (Top-Left view).
                        // Vậy tâm Y của hàng thứ 'row' sẽ là: StartY CÚA HÀNG ĐÓ + MViewHeight/2.
                        double centerY = StartY - row * (MViewHeight + MViewSpacingY) + (MViewHeight / 2.0);

                        var vp = new Viewport();
                        vp.CenterPoint = new Point3d(centerX, centerY, 0);
                        vp.Width = MViewWidth;
                        vp.Height = MViewHeight;

                        btr.AppendEntity(vp);
                        tr.AddNewlyCreatedDBObject(vp, true);

                        // Bật MView
                        vp.On = true;

                        // Gán góc nhìn Model
                        vp.ViewCenter = new Point2d(centerModelX, centerModelY);
                        vp.CustomScale = 1000.0 / ScaleVal; // Ratio 1:ScaleVal

                        vp.Locked = true; // Khoá lại tránh zoom trượt tay

                        col++;
                        if (col >= Columns)
                        {
                            col = 0;
                            row++;
                        }
                    }
                }
                tr.Commit();
                ed.WriteMessage($"\n✅ Đã tạo {numLayouts} Layout(s) chứa {frames.Count} Viewport(s).");
                
                // Trở về model space
                layoutMgr.CurrentLayout = "Model";
            }
        }
    }

    public class MainForm_ViewTD : CivilFormBase
    {
        private ComboBox cboBlocks = null!;
        private NumericUpDown nudWidth = null!, nudHeight = null!, nudSpacingX = null!, nudSpacingY = null!;
        private NumericUpDown nudStartX = null!, nudStartY = null!, nudColumns = null!, nudRows = null!, nudScale = null!;
        private Button btnDrawFrames = null!, btnCreateLayouts = null!, btnClose = null!;

        public MainForm_ViewTD(List<string> blockNames)
        {
            InitComponents(blockNames);
            LoadSettings();
        }

        private void InitComponents(List<string> blockNames)
        {
            InitBaseTheme("TẠO TRANG IN TRẮC DỌC / BÌNH ĐỒ", 550, 640);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            var header = CreateHeader("TẠO TRANG IN TRẮC DỌC / BÌNH ĐỒ");
            this.Controls.Add(header);

            // ── SUBTITLE ──
            var subtitle = new WinLabel
            {
                Text = "Cấu hình khung Viewport và bố trí tự động Layout từ các khung Model Space.",
                Font = new Font("Segoe UI", 9.5f),
                ForeColor = TextDark,
                AutoSize = false,
                Height = 35,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Top,
                Padding = new Padding(15, 0, 0, 0),
                BackColor = BgPanel
            };
            this.Controls.Add(subtitle);
            subtitle.BringToFront();

            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15),
                AutoScroll = true
            };
            this.Controls.Add(mainPanel);
            mainPanel.BringToFront();

            int y = 5;

            // SECTION 1: CẤU HÌNH VIEW
            var grpConfig = CreateSection("Cấu hình Khung View (MView)", 0, y, 500, 185);
            mainPanel.Controls.Add(grpConfig);

            grpConfig.Controls.Add(CreateLabel("Block Khung Tên:", 15, 33));
            cboBlocks = new ComboBox
            {
                Location = new Point(200, 29),
                Size = new Size(270, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextDark,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f)
            };
            cboBlocks.Items.Add("<Không dùng>");
            foreach (var b in blockNames) cboBlocks.Items.Add(b);
            if (cboBlocks.Items.Count > 0) cboBlocks.SelectedIndex = 0;
            grpConfig.Controls.Add(cboBlocks);

            grpConfig.Controls.Add(CreateLabel("Tỷ lệ hiển thị 1/:", 15, 68));
            nudScale = CreateNud(200, 64, 270, 1, 10000, 1000, 0, grpConfig);

            grpConfig.Controls.Add(CreateLabel("Rộng MView (Layout):", 15, 103));
            nudWidth = CreateNud(200, 99, 270, 10, 10000, 297, 1, grpConfig);

            grpConfig.Controls.Add(CreateLabel("Cao MView (Layout):", 15, 138));
            nudHeight = CreateNud(200, 134, 270, 10, 10000, 210, 1, grpConfig);

            y += 200;

            // SECTION 2: BỐ TRÍ
            var grpLayout = CreateSection("Tọa độ và Bố trí Layout", 0, y, 500, 255);
            mainPanel.Controls.Add(grpLayout);

            grpLayout.Controls.Add(CreateLabel("Tọa độ MView đầu X:", 15, 33));
            nudStartX = CreateNud(200, 29, 270, -10000, 10000, 10, 1, grpLayout);

            grpLayout.Controls.Add(CreateLabel("Tọa độ MView đầu Y:", 15, 68));
            nudStartY = CreateNud(200, 64, 270, -10000, 10000, 200, 1, grpLayout);

            grpLayout.Controls.Add(CreateLabel("Khoảng cách X:", 15, 103));
            nudSpacingX = CreateNud(200, 99, 270, 0, 10000, 10, 1, grpLayout);

            grpLayout.Controls.Add(CreateLabel("Khoảng cách Y:", 15, 138));
            nudSpacingY = CreateNud(200, 134, 270, 0, 10000, 10, 1, grpLayout);

            grpLayout.Controls.Add(CreateLabel("Số cột MView:", 15, 173));
            nudColumns = CreateNud(200, 169, 270, 1, 100, 1, 0, grpLayout);

            grpLayout.Controls.Add(CreateLabel("Số hàng MView:", 15, 208));
            nudRows = CreateNud(200, 204, 270, 1, 100, 1, 0, grpLayout);

            // SECTION 3: HÀNH ĐỘNG
            var pnlBottom = CreateBottomBar();
            this.Controls.Add(pnlBottom);

            btnDrawFrames = CreateButton("Vẽ Khung Model", 12, 8, 150, 36, pnlBottom);
            btnDrawFrames.Click += BtnDrawFrames_Click;
            btnDrawFrames.FlatAppearance.BorderColor = NavyMedium;
            btnDrawFrames.ForeColor = NavyMedium;

            btnClose = CreateButton("Đóng", 250, 8, 80, 36, pnlBottom);
            btnClose.Click += (s, e) => this.Close();

            btnCreateLayouts = CreatePrimaryButton("Tạo Layouts & MViews", 340, 8, 180, 36, pnlBottom);
            btnCreateLayouts.Click += BtnCreateLayouts_Click;
        }

        private NumericUpDown CreateNud(int x, int y, int w, decimal min, decimal max, decimal val, int dec, Control parent)
        {
            var nud = new NumericUpDown
            {
                Location = new Point(x, y),
                Size = new Size(w, 28),
                Minimum = min,
                Maximum = max,
                DecimalPlaces = dec,
                Value = val,
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            parent.Controls.Add(nud);
            return nud;
        }

        private void LoadSettings()
        {
            if (!string.IsNullOrEmpty(View_TD.SelectedBlockName) && cboBlocks.Items.Contains(View_TD.SelectedBlockName))
                cboBlocks.SelectedItem = View_TD.SelectedBlockName;
            
            nudWidth.Value = (decimal)View_TD.MViewWidth;
            nudHeight.Value = (decimal)View_TD.MViewHeight;
            nudSpacingX.Value = (decimal)View_TD.MViewSpacingX;
            nudSpacingY.Value = (decimal)View_TD.MViewSpacingY;
            nudStartX.Value = (decimal)View_TD.StartX;
            nudStartY.Value = (decimal)View_TD.StartY;
            nudColumns.Value = View_TD.Columns;
            nudRows.Value = View_TD.Rows;
            nudScale.Value = (decimal)View_TD.ScaleVal;
        }

        private void SaveSettings()
        {
            View_TD.SelectedBlockName = cboBlocks.SelectedItem?.ToString() ?? "";
            View_TD.MViewWidth = (double)nudWidth.Value;
            View_TD.MViewHeight = (double)nudHeight.Value;
            View_TD.MViewSpacingX = (double)nudSpacingX.Value;
            View_TD.MViewSpacingY = (double)nudSpacingY.Value;
            View_TD.StartX = (double)nudStartX.Value;
            View_TD.StartY = (double)nudStartY.Value;
            View_TD.Columns = (int)nudColumns.Value;
            View_TD.Rows = (int)nudRows.Value;
            View_TD.ScaleVal = (double)nudScale.Value;
        }

        private void BtnDrawFrames_Click(object? sender, EventArgs e)
        {
            SaveSettings();
            this.DialogResult = DialogResult.Retry; // Trapping loop in parent
            this.Close();
        }

        private void BtnCreateLayouts_Click(object? sender, EventArgs e)
        {
            SaveSettings();
            this.DialogResult = DialogResult.OK; // Layout generation
            this.Close();
        }
    }
}
