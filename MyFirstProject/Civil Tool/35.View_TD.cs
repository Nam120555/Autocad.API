using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.DatabaseServices;
using Acad = Autodesk.AutoCAD.ApplicationServices;
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using WinPoint = System.Drawing.Point;

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
        public static double StartY = 200; 
        public static int Columns = 3;
        public static double ScaleVal = 1000;
        
        // Điều hướng tương tác pick điểm
        public static int CurrentPickAction = 0; 
        // 1: Pick XREF, 2: Kích thước, 3: Gốc, 4: Hở X, 5: Hở Y, 6: Cắt Tuyến

        [CommandMethod("CTSV_TaoTrangInTD")]
        public void CTSVTaoTrangInTD()
        {
            var doc = Acad.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            bool keepLooping = true;
            while (keepLooping)
            {
                List<string> blockNames = new List<string>();
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite); 
                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForWrite);
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
                    
                    if (res == DialogResult.Retry) // Chọn Vẽ khung Model thủ công
                    {
                        DrawFrames(db, ed);
                    }
                    else if (res == DialogResult.Ignore) // Các thao tác Auto Pick / Auto Align
                    {
                        switch(CurrentPickAction)
                        {
                            case 1: SelectXrefFromScreen(db, ed); break;
                            case 2: PickPointsForSize(ed); break;
                            case 3: PickStartPoint(ed); break;
                            case 4: PickDistance("X", ed); break;
                            case 5: PickDistance("Y", ed); break;
                            case 6: DrawFramesAlongAlignment(db, ed); break;
                        }
                    }
                    else if (res == DialogResult.OK) // Chọn Tạo Layout
                    {
                        CreateLayouts(db, ed);
                        keepLooping = false;
                    }
                    else 
                    {
                        keepLooping = false;
                    }
                }
            }
        }

        private void DrawFramesAlongAlignment(Database db, Editor ed)
        {
            var ppo = new PromptEntityOptions("\n[Rải Dọc Mùa] Chọn Tuyến (Alignment) để tự động dàn khung bình đồ: ");
            ppo.SetRejectMessage("\nLỗi: Chọn sai. Yêu cầu click vào tim đường Alignment!");
            ppo.AddAllowedClass(typeof(Alignment), true);
            var res = ed.GetEntity(ppo);
            if (res.Status != PromptStatus.OK) return;

            // Optional overlap = 5% chiều rộng khung để các bản vẽ có tính liên kết nối tiếp
            double frameW = MViewWidth * (ScaleVal / 1000.0);
            double frameH = MViewHeight * (ScaleVal / 1000.0);
            double overlap = frameW * 0.05; 
            double step = frameW - overlap;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var align = (Alignment)tr.GetObject(res.ObjectId, OpenMode.ForRead);
                    double startSta = align.StartingStation;
                    double endSta = align.EndingStation;

                    string layerName = "KhungView_BinhDoAuto";
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
                    if (!lt.Has(layerName))
                    {
                        var ltr = new LayerTableRecord();
                        ltr.Name = layerName;
                        ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 6); // Magenta
                        lt.Add(ltr);
                        tr.AddNewlyCreatedDBObject(ltr, true);
                    }

                    var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    int count = 0;
                    for (double sta = startSta + frameW / 2.0; sta < endSta + frameW / 2.0; sta += step) // Đảo tiến lý trình
                    {
                        double x = 0, y = 0;
                        align.PointLocation(sta, 0, ref x, ref y);
                        Point3d pCen = new Point3d(x, y, 0);

                        // Lấy Vector tiếp tuyến để xoay vuông góc với cung đường
                        double x1 = 0, y1 = 0, x2 = 0, y2 = 0;
                        double ds = 0.5; // Mượn khoảng siêu nhỏ đo góc
                        if (sta - ds >= startSta && sta + ds <= endSta) {
                            align.PointLocation(sta - ds, 0, ref x1, ref y1);
                            align.PointLocation(sta + ds, 0, ref x2, ref y2);
                        } else if (sta + ds <= endSta) {
                            align.PointLocation(sta, 0, ref x1, ref y1);
                            align.PointLocation(sta + ds, 0, ref x2, ref y2);
                        } else {
                            align.PointLocation(sta - ds, 0, ref x1, ref y1);
                            align.PointLocation(sta, 0, ref x2, ref y2);
                        }
                        double angle = Math.Atan2(y2 - y1, x2 - x1);

                        var pline = new Polyline();
                        Point2d[] pts = new Point2d[4] {
                            new Point2d(-frameW/2, -frameH/2),
                            new Point2d(frameW/2, -frameH/2),
                            new Point2d(frameW/2, frameH/2),
                            new Point2d(-frameW/2, frameH/2)
                        };

                        for (int i=0; i<4; i++) {
                            // Xoay ma trận Point
                            double rotX = pts[i].X * Math.Cos(angle) - pts[i].Y * Math.Sin(angle);
                            double rotY = pts[i].X * Math.Sin(angle) + pts[i].Y * Math.Cos(angle);
                            pline.AddVertexAt(i, new Point2d(pCen.X + rotX, pCen.Y + rotY), 0, 0, 0);
                        }
                        pline.Closed = true;
                        pline.Layer = layerName;

                        btr.AppendEntity(pline);
                        tr.AddNewlyCreatedDBObject(pline, true);
                        count++;
                    }

                    tr.Commit();
                    ed.WriteMessage($"\n>> Đã tự động sinh {count} Khung MView xoay uốn lượn bay bổng dọc theo tuyến {align.Name}.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nLỗi rải theo tuyến: " + ex.Message);
            }
        }

        private void PickPointsForSize(Editor ed)
        {
            var pt1Opts = new PromptPointOptions("\n[Đo MView] Chọn góc DƯỚI BÊN TRÁI của khung: ");
            var pt1Res = ed.GetPoint(pt1Opts);
            if (pt1Res.Status != PromptStatus.OK) return;

            var pt2Opts = new PromptCornerOptions("\nChọn góc TRÊN BÊN PHẢI của khung: ", pt1Res.Value);
            var pt2Res = ed.GetCorner(pt2Opts);
            if (pt2Res.Status != PromptStatus.OK) return;

            try
            {
                double w = Math.Abs(pt2Res.Value.X - pt1Res.Value.X);
                double h = Math.Abs(pt2Res.Value.Y - pt1Res.Value.Y);

                if (w > 0 && h > 0)
                {
                    MViewWidth = Math.Round(w, 2);
                    MViewHeight = Math.Round(h, 2);
                    ed.WriteMessage($"\n>> Cập nhật Kích Thước (Rộng x Cao): {MViewWidth} x {MViewHeight}");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nLỗi đo kích thước: " + ex.Message);
            }
        }

        private void PickStartPoint(Editor ed)
        {
            var ppo = new PromptPointOptions("\n[Đo Gốc Layout] Click 1 điểm làm toạ độ X,Y bắt đầu rải mặt bằng: ");
            var res = ed.GetPoint(ppo);
            if (res.Status == PromptStatus.OK)
            {
                StartX = Math.Round(res.Value.X, 2);
                StartY = Math.Round(res.Value.Y, 2);
                ed.WriteMessage($"\n>> Đã Lưu Tọa Độ Chèn Gốc: X={StartX}, Y={StartY}");
            }
        }

        private void PickDistance(string axis, Editor ed)
        {
            var pdo = new PromptDistanceOptions($"\n[Đo Khoảng Cách] Click chọn 2 điểm trên mặt bằng để đo khoảng hở {axis}: ");
            var res = ed.GetDistance(pdo);
            if (res.Status == PromptStatus.OK)
            {
                if (axis == "X") 
                {
                    MViewSpacingX = Math.Round(res.Value, 2);
                    ed.WriteMessage($"\n>> Cập nhật Hở TRỤC X: {MViewSpacingX}");
                }
                else 
                {
                    MViewSpacingY = Math.Round(res.Value, 2);
                    ed.WriteMessage($"\n>> Cập nhật Hở TRỤC Y: {MViewSpacingY}");
                }
            }
        }

        private void SelectXrefFromScreen(Database db, Editor ed)
        {
            var peo = new PromptEntityOptions("\n[Chỉ Bắt Điểm] Nhấp vào XREF / Block Khung Tên trên mô hình (Zoom to nếu cần): ");
            peo.SetRejectMessage("\nLỗi: Đối tượng bạn chọn không phải là XREF hay Block!");
            peo.AddAllowedClass(typeof(BlockReference), false);

            var res = ed.GetEntity(peo);
            if (res.Status != PromptStatus.OK) return;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bref = (BlockReference)tr.GetObject(res.ObjectId, OpenMode.ForWrite);
                    string bName = bref.IsDynamicBlock ? 
                        ((BlockTableRecord)tr.GetObject(bref.DynamicBlockTableRecord, OpenMode.ForWrite)).Name : 
                        bref.Name;

                    Extents3d ext = bref.GeometricExtents;
                    double w = ext.MaxPoint.X - ext.MinPoint.X;
                    double h = ext.MaxPoint.Y - ext.MinPoint.Y;

                    if (w < 1) w = 297;
                    if (h < 1) h = 210;

                    SelectedBlockName = bName;
                    MViewWidth = Math.Round(w, 2);
                    MViewHeight = Math.Round(h, 2);

                    tr.Commit();
                    ed.WriteMessage($"\n>> Đã nhận XREF: [{bName}] ({MViewWidth} x {MViewHeight})");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nLỗi đo XREF: " + ex.Message);
            }
        }

        private void DrawFrames(Database db, Editor ed)
        {
            double frameW = MViewWidth * (ScaleVal / 1000.0);
            double frameH = MViewHeight * (ScaleVal / 1000.0);

            ed.WriteMessage($"\n--- LỆNH VẼ KHUNG MVIEW LÊN MODEL ({frameW}x{frameH}) ---");

            while (true)
            {
                var ppo = new PromptPointOptions("\nChọn gốc chèn (Góc trái dưới) cho khung nhìn tiếp theo (HOẶC NHẤP ESC ĐỂ TRỞ VỀ FORM): ");
                var res = ed.GetPoint(ppo);
                if (res.Status != PromptStatus.OK) break;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    string layerName = "KhungView_TD";
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
                    if (!lt.Has(layerName))
                    {
                        var ltr = new LayerTableRecord();
                        ltr.Name = layerName;
                        ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1);
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
            pso.MessageForAdding = "\nVui lòng quét các viền Polyline trên Model mà bạn muốn kéo vào Layout (Nhấn ENTER khi quét xong): ";
            var selRes = ed.GetSelection(pso);
            if (selRes.Status != PromptStatus.OK) return;

            List<Polyline> frames = new List<Polyline>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject so in selRes.Value)
                {
                    var ent = tr.GetObject(so.ObjectId, OpenMode.ForWrite) as Polyline;
                    if (ent != null && ent.NumberOfVertices >= 4)
                    {
                        frames.Add(ent);
                    }
                }

                if (frames.Count == 0)
                {
                    ed.WriteMessage("\nKhông tìm thấy khung nào.");
                    tr.Commit();
                    return;
                }

                // Sắp xếp thứ tự cắt khung không dựa vào X tuyệt đối, mà dựa vào Chiều Dài Curve Trung Bình. (Để Layout chạy đúng số tờ)
                frames.Sort((a, b) => a.GetPoint3dAt(0).X.CompareTo(b.GetPoint3dAt(0).X));

                var layoutMgr = LayoutManager.Current;
                
                string layoutName = "Layout_TD_TongHop";
                ObjectId layoutId = layoutMgr.GetLayoutId(layoutName);
                if (layoutId.IsNull)
                {
                    layoutId = layoutMgr.CreateLayout(layoutName);
                }
                
                Layout layout = (Layout)tr.GetObject(layoutId, OpenMode.ForWrite);
                layoutMgr.CurrentLayout = layoutName;
                
                var btrLayout = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

                foreach (ObjectId id in btrLayout)
                {
                    var ent = tr.GetObject(id, OpenMode.ForWrite);
                    if (ent is Viewport vp && vp.Number != 1)
                    {
                        ent.Erase();
                    }
                    else if (ent is BlockReference)
                    {
                        ent.Erase(); 
                    }
                }

                int col = 0;
                int row = 0;
                for (int i = 0; i < frames.Count; i++)
                {
                    var frame = frames[i];

                    // TRUNG TÂM MODEL CHUẨN TỪ ĐA GIÁC BẤT KỲ (XOAY HOẶC KHÔNG XOAY)
                    double sumX = 0, sumY = 0;
                    for (int v = 0; v < frame.NumberOfVertices; v++)
                    {
                        sumX += frame.GetPoint3dAt(v).X;
                        sumY += frame.GetPoint3dAt(v).Y;
                    }
                    double centerModelX = sumX / frame.NumberOfVertices;
                    double centerModelY = sumY / frame.NumberOfVertices;

                    // TÍNH GÓC XOAY TWIST ANGLE
                    Point3d p0 = frame.GetPoint3dAt(0);
                    Point3d p1 = frame.GetPoint3dAt(1);
                    double frameAng = Math.Atan2(p1.Y - p0.Y, p1.X - p0.X);

                    double blX = StartX + col * (MViewWidth + MViewSpacingX);
                    double blY = StartY - row * (MViewHeight + MViewSpacingY);
                    
                    double centerX = blX + MViewWidth / 2.0;
                    double centerY = blY + MViewHeight / 2.0;

                    if (!string.IsNullOrEmpty(SelectedBlockName) && SelectedBlockName != "<Không dùng>")
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                        if (bt.Has(SelectedBlockName))
                        {
                            var bref = new BlockReference(new Point3d(blX, blY, 0), bt[SelectedBlockName]);
                            btrLayout.AppendEntity(bref);
                            tr.AddNewlyCreatedDBObject(bref, true);
                        }
                    }

                    var vp = new Viewport();
                    vp.CenterPoint = new Point3d(centerX, centerY, 0);
                    vp.Width = MViewWidth;
                    vp.Height = MViewHeight;

                    btrLayout.AppendEntity(vp);
                    tr.AddNewlyCreatedDBObject(vp, true);
                    
                    vp.On = true;
                    vp.ViewCenter = new Point2d(centerModelX, centerModelY);
                    vp.CustomScale = 1000.0 / ScaleVal; 
                    
                    // Xoay chống bẻ nghiêng tuyến
                    vp.TwistAngle = -frameAng; // Twist Angle trừ lại góc cong của Model
                    
                    vp.Locked = true;  

                    col++;
                    if (col >= Columns)
                    {
                        col = 0;
                        row++;
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n--- HOÀN TẤT --- Đã rải xong {frames.Count} cụm vào {layoutName}.");
            }
        }
    }

    public class MainForm_ViewTD : CivilFormBase
    {
        private ComboBox cboBlocks = null!;
        private NumericUpDown nudWidth = null!, nudHeight = null!, nudSpacingX = null!, nudSpacingY = null!;
        private NumericUpDown nudStartX = null!, nudStartY = null!, nudColumns = null!, nudScale = null!;
        private Button btnDrawFrames = null!, btnDrawAlign = null!, btnCreateLayouts = null!, btnClose = null!, btnSelectXref = null!;

        public MainForm_ViewTD(List<string> blockNames)
        {
            InitComponents(blockNames);
            LoadSettings();
        }

        private Button CreatePickButton(Control parent, int x, int y, string text, int actionId)
        {
            var btn = new Button
            {
                Text = text,
                Location = new WinPoint(x, y),
                Size = new Size(35, 28),
                BackColor = Color.FromArgb(41, 128, 185), // Navy sáng
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.Click += (s, e) => {
                SaveSettings();
                View_TD.CurrentPickAction = actionId;
                this.DialogResult = DialogResult.Ignore;
                this.Close();
            };
            parent.Controls.Add(btn);
            return btn;
        }

        private void InitComponents(List<string> blockNames)
        {
            InitBaseTheme("SẮP XẾP KHUNG VIEW BÌNH ĐỒ HÀNG LOẠT", 560, 680);
            this.FormBorderStyle = FormBorderStyle.Sizable;

            var header = CreateHeader("QUẢN LÝ VIEWPORT & XREF LAYOUTS");
            this.Controls.Add(header);

            var subtitle = new WinLabel
            {
                Text = "Tạo mảng khung Bình đồ theo Alignment tự động và đồng bộ xoay ngang giấy ảo giác Layout.",
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

            // SECTION 1: CẤU HÌNH VIEW VÀ CHỌN XREF
            var grpConfig = CreateSection("1. Bốc tách và Khai báo XREF Khung Tên Gốc", 0, y, 510, 205);
            mainPanel.Controls.Add(grpConfig);

            btnSelectXref = CreatePrimaryButton("🔗 Quét Nhận Diện Khung Tên / XREF...", 15, 25, 470, 32, grpConfig);
            btnSelectXref.BackColor = Color.FromArgb(231, 76, 60); 
            btnSelectXref.Click += (s, e) => {
                SaveSettings();
                View_TD.CurrentPickAction = 1;
                this.DialogResult = DialogResult.Ignore;
                this.Close();
            };

            grpConfig.Controls.Add(CreateLabel("Khung tên đang dùng:", 15, 73));
            cboBlocks = new ComboBox
            {
                Location = new WinPoint(180, 69),
                Size = new Size(260, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextDark,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f)
            };
            cboBlocks.Items.Add("<Không dùng>");
            foreach (var b in blockNames) cboBlocks.Items.Add(b);
            if (!string.IsNullOrEmpty(View_TD.SelectedBlockName) && !cboBlocks.Items.Contains(View_TD.SelectedBlockName))
            {
                cboBlocks.Items.Add(View_TD.SelectedBlockName);
            }
            if (cboBlocks.Items.Count > 0) cboBlocks.SelectedIndex = 0;
            grpConfig.Controls.Add(cboBlocks);

            grpConfig.Controls.Add(CreateLabel("Tỉ lệ Model hiển thị 1/", 15, 108));
            nudScale = CreateNud(180, 104, 260, 1, 100000, 1000, 0, grpConfig);

            grpConfig.Controls.Add(CreateLabel("Kích thước X (Rộng):", 15, 143));
            nudWidth = CreateNud(180, 139, 260, 0.1m, 100000, 297, 2, grpConfig);

            grpConfig.Controls.Add(CreateLabel("Kích thước Y (Cao):", 15, 178));
            nudHeight = CreateNud(180, 174, 260, 0.1m, 100000, 210, 2, grpConfig);

            var btnPickSize = CreatePickButton(grpConfig, 445, 139, "🖱️", 2);
            btnPickSize.Height = 63; 
            btnPickSize.Text = "📐";

            y += 220;

            // SECTION 2: BỐ TRÍ TRÊN 1 LAYOUT
            var grpLayout = CreateSection("2. Cấu trúc Cuộn Trang / Lưới Dàn Layouts", 0, y, 510, 215);
            mainPanel.Controls.Add(grpLayout);

            grpLayout.Controls.Add(CreateLabel("Toạ độ Chèn Cụm Khung X:", 15, 33));
            nudStartX = CreateNud(180, 29, 260, -100000, 100000, 10, 1, grpLayout);

            grpLayout.Controls.Add(CreateLabel("Toạ độ Chèn Cụm Khung Y:", 15, 68));
            nudStartY = CreateNud(180, 64, 260, -100000, 100000, 200, 1, grpLayout);
            
            var btnPickStart = CreatePickButton(grpLayout, 445, 29, "📍", 3);
            btnPickStart.Height = 63;

            grpLayout.Controls.Add(CreateLabel("Khoảng Đệm Trục X (Hở):", 15, 103));
            nudSpacingX = CreateNud(180, 99, 260, 0, 10000, 30, 1, grpLayout);
            CreatePickButton(grpLayout, 445, 99, "↔️", 4);

            grpLayout.Controls.Add(CreateLabel("Khoảng Đệm Trục Y (Hở):", 15, 138));
            nudSpacingY = CreateNud(180, 134, 260, 0, 10000, 30, 1, grpLayout);
            CreatePickButton(grpLayout, 445, 134, "↕️", 5);

            grpLayout.Controls.Add(CreateLabel("Dàn hàng bao nhiêu Cột?:", 15, 173));
            nudColumns = CreateNud(180, 169, 260, 1, 1000, 3, 0, grpLayout);

            // SECTION 3: HÀNH ĐỘNG THÔNG MINH
            var pnlBottom = CreateBottomBar();
            this.Controls.Add(pnlBottom);

            btnDrawFrames = CreateButton("Vẽ Tay", 12, 8, 110, 36, pnlBottom);
            btnDrawFrames.Click += BtnDrawFrames_Click;
            btnDrawFrames.FlatAppearance.BorderColor = NavyMedium;
            btnDrawFrames.ForeColor = NavyMedium;

            btnDrawAlign = CreatePrimaryButton("🛤️ Auto Cắt Tuyến BĐ", 127, 8, 175, 36, pnlBottom);
            btnDrawAlign.BackColor = Color.FromArgb(142, 68, 173); // MÀU TÍM BÌNH ĐỒ
            btnDrawAlign.Click += (s, e) => {
                SaveSettings();
                View_TD.CurrentPickAction = 6;
                this.DialogResult = DialogResult.Ignore;
                this.Close();
            };

            btnClose = CreateButton("Đóng", 307, 8, 60, 36, pnlBottom);
            btnClose.Click += (s, e) => this.Close();

            btnCreateLayouts = CreatePrimaryButton("Rải Vào Layout", 373, 8, 150, 36, pnlBottom);
            btnCreateLayouts.Click += BtnCreateLayouts_Click;
        }

        private NumericUpDown CreateNud(int x, int y, int w, decimal min, decimal max, decimal val, int dec, Control parent)
        {
            var nud = new NumericUpDown
            {
                Location = new WinPoint(x, y),
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
            View_TD.ScaleVal = (double)nudScale.Value;
        }

        private void BtnDrawFrames_Click(object? sender, EventArgs e)
        {
            SaveSettings();
            this.DialogResult = DialogResult.Retry; 
            this.Close();
        }

        private void BtnCreateLayouts_Click(object? sender, EventArgs e)
        {
            SaveSettings();
            this.DialogResult = DialogResult.OK; 
            this.Close();
        }
    }
}
