// (C) Copyright 2024
// Tool San Nền - Tính khối lượng đào đắp theo phương pháp lưới ô vuông
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Colors;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using ClosedXML.Excel;
using FormLabel = System.Windows.Forms.Label;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadColor = Autodesk.AutoCAD.Colors.Color;
using MyFirstProject.Helpers;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(MyFirstProject.SanNen))]

namespace MyFirstProject
{
    #region Data Classes

    /// <summary>
    /// Thông tin một góc của ô lưới
    /// </summary>
    public class GridCorner
    {
        public Point3d Position { get; set; }           // Vị trí XY
        public double ElevationTN { get; set; }         // Cao độ tự nhiên
        public double ElevationTK { get; set; }         // Cao độ thiết kế
        public double DeltaH => ElevationTK - ElevationTN;  // Độ chênh (+ đắp, - đào)
        public ObjectId TextId { get; set; }            // ID của text hiển thị cao độ
    }

    /// <summary>
    /// Thông tin một ô lưới
    /// </summary>
    public class GridCell
    {
        public int Row { get; set; }                    // Hàng
        public int Col { get; set; }                    // Cột
        public string Name => $"{Row}-{Col}";           // Tên ô: "1-1", "1-2", ...
        public GridCorner[] Corners { get; set; } = new GridCorner[4];  // 4 góc: TL, TR, BR, BL
        public double Area { get; set; }                // Diện tích ô (m²)
        public ObjectId PolylineId { get; set; }        // ID của polyline vẽ ô
        public ObjectId BlockId { get; set; }           // ID của block hiển thị KL

        /// <summary>
        /// Độ chênh trung bình của ô
        /// </summary>
        public double AverageDeltaH
        {
            get
            {
                if (Corners == null || Corners.Length < 4) return 0;
                return (Corners[0].DeltaH + Corners[1].DeltaH + Corners[2].DeltaH + Corners[3].DeltaH) / 4.0;
            }
        }

        /// <summary>
        /// Khối lượng đào (m³) - giá trị dương khi đào
        /// </summary>
        public double CutVolume
        {
            get
            {
                double delta = AverageDeltaH;
                return delta < 0 ? Math.Abs(delta) * Area : 0;
            }
        }

        /// <summary>
        /// Khối lượng đắp (m³) - giá trị dương khi đắp
        /// </summary>
        public double FillVolume
        {
            get
            {
                double delta = AverageDeltaH;
                return delta > 0 ? delta * Area : 0;
            }
        }
    }

    /// <summary>
    /// Thông tin lưới san nền
    /// </summary>
    public class GradingGrid
    {
        public string Name { get; set; } = "SanNen";
        public Point3d Origin { get; set; }             // Điểm gốc (góc dưới trái)
        public double CellSize { get; set; } = 20.0;    // Kích thước ô (m)
        public int Rows { get; set; }                   // Số hàng
        public int Cols { get; set; }                   // Số cột
        public List<GridCell> Cells { get; set; } = new();
        public Dictionary<string, GridCorner> Corners { get; set; } = new();  // Key: "row-col"

        /// <summary>
        /// Tổng khối lượng đào
        /// </summary>
        public double TotalCutVolume => Cells.Sum(c => c.CutVolume);

        /// <summary>
        /// Tổng khối lượng đắp
        /// </summary>
        public double TotalFillVolume => Cells.Sum(c => c.FillVolume);

        /// <summary>
        /// Tổng diện tích
        /// </summary>
        public double TotalArea => Cells.Sum(c => c.Area);
    }

    #endregion

    #region Settings Form

    /// <summary>
    /// Form cài đặt san nền
    /// </summary>
    public class SanNenSettingsForm : Form
    {
        private NumericUpDown nudCellSize = null!;
        private NumericUpDown nudTextHeight = null!;
        private ComboBox cboSurfaceTN = null!;
        private ComboBox cboSurfaceTK = null!;
        private Button btnOK = null!;
        private Button btnCancel = null!;

        public double CellSize { get; private set; } = 20.0;
        public double TextHeight { get; private set; } = 2.5;
        public ObjectId SurfaceTNId { get; private set; }
        public ObjectId SurfaceTKId { get; private set; }

        private List<(string Name, ObjectId Id)> surfaces = new();

        public SanNenSettingsForm(Transaction tr)
        {
            LoadSurfaces(tr);
            InitializeComponent();
        }

        private void LoadSurfaces(Transaction tr)
        {
            try
            {
                CivilDocument civilDoc = CivilApplication.ActiveDocument;
                ObjectIdCollection surfaceIds = civilDoc.GetSurfaceIds();

                foreach (ObjectId id in surfaceIds)
                {
                    if (tr.GetObject(id, OpenMode.ForRead) is TinSurface surface)
                    {
                        surfaces.Add((surface.Name, id));
                    }
                }
            }
            catch { }
        }

        private void InitializeComponent()
        {
            this.Text = "Cài đặt San Nền";
            this.Size = new System.Drawing.Size(400, 280);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = System.Drawing.Color.FromArgb(45, 45, 48);
            this.ForeColor = System.Drawing.Color.White;

            int labelWidth = 150;
            int inputWidth = 180;
            int rowHeight = 35;
            int margin = 20;
            int y = 20;

            // Title
            var lblTitle = new FormLabel
            {
                Text = "⚙ CÀI ĐẶT SAN NỀN",
                Location = new System.Drawing.Point(margin, y),
                Size = new System.Drawing.Size(350, 25),
                Font = new System.Drawing.Font("Segoe UI", 11, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(0, 122, 204)
            };
            this.Controls.Add(lblTitle);
            y += 35;

            // Kích thước ô
            var lblCellSize = new FormLabel
            {
                Text = "Kích thước ô lưới (m):",
                Location = new System.Drawing.Point(margin, y + 3),
                Size = new System.Drawing.Size(labelWidth, 20),
                ForeColor = System.Drawing.Color.White
            };
            this.Controls.Add(lblCellSize);

            nudCellSize = new NumericUpDown
            {
                Location = new System.Drawing.Point(margin + labelWidth + 10, y),
                Size = new System.Drawing.Size(inputWidth, 25),
                DecimalPlaces = 1,
                Minimum = 1M,
                Maximum = 100M,
                Increment = 5M,
                Value = 20M,
                BackColor = System.Drawing.Color.FromArgb(60, 60, 65),
                ForeColor = System.Drawing.Color.White
            };
            this.Controls.Add(nudCellSize);
            y += rowHeight;

            // Chiều cao text
            var lblTextHeight = new FormLabel
            {
                Text = "Chiều cao chữ (mm):",
                Location = new System.Drawing.Point(margin, y + 3),
                Size = new System.Drawing.Size(labelWidth, 20),
                ForeColor = System.Drawing.Color.White
            };
            this.Controls.Add(lblTextHeight);

            nudTextHeight = new NumericUpDown
            {
                Location = new System.Drawing.Point(margin + labelWidth + 10, y),
                Size = new System.Drawing.Size(inputWidth, 25),
                DecimalPlaces = 1,
                Minimum = 0.5M,
                Maximum = 20M,
                Increment = 0.5M,
                Value = 2.5M,
                BackColor = System.Drawing.Color.FromArgb(60, 60, 65),
                ForeColor = System.Drawing.Color.White
            };
            this.Controls.Add(nudTextHeight);
            y += rowHeight;

            // Surface TN
            var lblSurfaceTN = new FormLabel
            {
                Text = "Surface Tự Nhiên:",
                Location = new System.Drawing.Point(margin, y + 3),
                Size = new System.Drawing.Size(labelWidth, 20),
                ForeColor = System.Drawing.Color.White
            };
            this.Controls.Add(lblSurfaceTN);

            cboSurfaceTN = new ComboBox
            {
                Location = new System.Drawing.Point(margin + labelWidth + 10, y),
                Size = new System.Drawing.Size(inputWidth, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = System.Drawing.Color.FromArgb(60, 60, 65),
                ForeColor = System.Drawing.Color.White
            };
            cboSurfaceTN.Items.Add("(Nhập thủ công)");
            foreach (var surf in surfaces)
                cboSurfaceTN.Items.Add(surf.Name);
            cboSurfaceTN.SelectedIndex = 0;
            this.Controls.Add(cboSurfaceTN);
            y += rowHeight;

            // Surface TK
            var lblSurfaceTK = new FormLabel
            {
                Text = "Surface Thiết Kế:",
                Location = new System.Drawing.Point(margin, y + 3),
                Size = new System.Drawing.Size(labelWidth, 20),
                ForeColor = System.Drawing.Color.White
            };
            this.Controls.Add(lblSurfaceTK);

            cboSurfaceTK = new ComboBox
            {
                Location = new System.Drawing.Point(margin + labelWidth + 10, y),
                Size = new System.Drawing.Size(inputWidth, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = System.Drawing.Color.FromArgb(60, 60, 65),
                ForeColor = System.Drawing.Color.White
            };
            cboSurfaceTK.Items.Add("(Nhập thủ công)");
            foreach (var surf in surfaces)
                cboSurfaceTK.Items.Add(surf.Name);
            cboSurfaceTK.SelectedIndex = 0;
            this.Controls.Add(cboSurfaceTK);
            y += rowHeight + 10;

            // Buttons
            int btnWidth = 100;
            int btnHeight = 30;

            btnOK = new Button
            {
                Text = "✓ Tiếp tục",
                Location = new System.Drawing.Point(margin + 50, y),
                Size = new System.Drawing.Size(btnWidth, btnHeight),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                ForeColor = System.Drawing.Color.White,
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.OK
            };
            btnOK.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(0, 100, 180);
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            btnCancel = new Button
            {
                Text = "✕ Hủy",
                Location = new System.Drawing.Point(margin + 50 + btnWidth + 20, y),
                Size = new System.Drawing.Size(btnWidth, btnHeight),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(100, 100, 105),
                ForeColor = System.Drawing.Color.White,
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(80, 80, 85);
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void BtnOK_Click(object? sender, EventArgs e)
        {
            CellSize = (double)nudCellSize.Value;
            TextHeight = (double)nudTextHeight.Value;

            if (cboSurfaceTN.SelectedIndex > 0)
                SurfaceTNId = surfaces[cboSurfaceTN.SelectedIndex - 1].Id;

            if (cboSurfaceTK.SelectedIndex > 0)
                SurfaceTKId = surfaces[cboSurfaceTK.SelectedIndex - 1].Id;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }

    #endregion

    #region Taskbar

    /// <summary>
    /// Thanh công cụ San Nền
    /// </summary>
    public class SanNenTaskbar : Form
    {
        private static SanNenTaskbar? instance;

        public SanNenTaskbar()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "🔧 San Nền - Tính Khối Lượng";
            this.Size = new System.Drawing.Size(680, 90);
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.StartPosition = FormStartPosition.Manual;
            this.TopMost = true;
            this.BackColor = System.Drawing.Color.FromArgb(45, 45, 48);
            this.ForeColor = System.Drawing.Color.White;

            // Đặt vị trí ở góc dưới màn hình
            var screenBounds = Screen.PrimaryScreen?.WorkingArea ?? new System.Drawing.Rectangle(0, 0, 1920, 1080);
            this.Location = new System.Drawing.Point(
                (screenBounds.Width - this.Width) / 2,
                screenBounds.Height - this.Height - 50
            );

            int btnWidth = 95;
            int btnHeight = 45;
            int margin = 8;
            int x = margin;
            int y = 10;

            // Button 1: Cài đặt
            var btnSettings = CreateButton("⚙", "Cài đặt", x, y, btnWidth, btnHeight, 
                System.Drawing.Color.FromArgb(100, 100, 105));
            btnSettings.Click += (s, e) => { this.Hide(); SanNen.CTSNCaiDat(); this.Show(); };
            this.Controls.Add(btnSettings);
            x += btnWidth + margin;

            // Button 2: Tạo lưới
            var btnTaoLuoi = CreateButton("▦", "Tạo lưới", x, y, btnWidth, btnHeight,
                System.Drawing.Color.FromArgb(0, 122, 204));
            btnTaoLuoi.Click += (s, e) => { this.Hide(); SanNen.CTSNTaoLuoi(); this.Show(); };
            this.Controls.Add(btnTaoLuoi);
            x += btnWidth + margin;

            // Button 3: Nhập cao độ
            var btnNhapCD = CreateButton("📝", "Nhập CĐ", x, y, btnWidth, btnHeight,
                System.Drawing.Color.FromArgb(60, 150, 60));
            btnNhapCD.Click += (s, e) => { this.Hide(); SanNen.CTSNNhapCaoDo(); this.Show(); };
            this.Controls.Add(btnNhapCD);
            x += btnWidth + margin;

            // Button 4: Lấy CĐ từ Surface
            var btnSurface = CreateButton("🏔", "Surface", x, y, btnWidth, btnHeight,
                System.Drawing.Color.FromArgb(150, 100, 50));
            btnSurface.Click += (s, e) => { this.Hide(); SanNen.CTSNSurface(); this.Show(); };
            this.Controls.Add(btnSurface);
            x += btnWidth + margin;

            // Button 5: Tính khối lượng
            var btnTinhKL = CreateButton("📊", "Tính KL", x, y, btnWidth, btnHeight,
                System.Drawing.Color.FromArgb(180, 60, 60));
            btnTinhKL.Click += (s, e) => { this.Hide(); SanNen.CTSNTinhKL(); this.Show(); };
            this.Controls.Add(btnTinhKL);
            x += btnWidth + margin;

            // Button 6: Xuất bảng
            var btnXuatBang = CreateButton("📋", "Xuất bảng", x, y, btnWidth, btnHeight,
                System.Drawing.Color.FromArgb(60, 60, 150));
            btnXuatBang.Click += (s, e) => { this.Hide(); SanNen.CTSNXuatBang(); this.Show(); };
            this.Controls.Add(btnXuatBang);
            x += btnWidth + margin;

            // Button Close
            var btnClose = CreateButton("✕", "", x, y, 40, btnHeight,
                System.Drawing.Color.FromArgb(150, 50, 50));
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);
        }

        private Button CreateButton(string icon, string text, int x, int y, int width, int height, System.Drawing.Color bgColor)
        {
            var btn = new Button
            {
                Text = string.IsNullOrEmpty(text) ? icon : $"{icon}\n{text}",
                Location = new System.Drawing.Point(x, y),
                Size = new System.Drawing.Size(width, height),
                FlatStyle = FlatStyle.Flat,
                BackColor = bgColor,
                ForeColor = System.Drawing.Color.White,
                Cursor = Cursors.Hand,
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold)
            };
            btn.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(80, 80, 85);
            btn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(
                Math.Min(bgColor.R + 30, 255),
                Math.Min(bgColor.G + 30, 255),
                Math.Min(bgColor.B + 30, 255));
            return btn;
        }

        public static void ShowTaskbar()
        {
            if (instance == null || instance.IsDisposed)
            {
                instance = new SanNenTaskbar();
            }
            instance.Show();
            instance.BringToFront();
        }

        public static void CloseTaskbar()
        {
            instance?.Close();
            instance = null;
        }
    }

    #endregion

    public class SanNen
    {
        // Lưu trữ lưới hiện tại
        private static GradingGrid? currentGrid;
        private static double defaultCellSize = 20.0;
        private static double defaultTextHeight = 2.5;

        #region Commands

        /// <summary>
        /// Tạo lưới ô vuông trên mặt bằng
        /// </summary>
        [CommandMethod("CTSN_TaoLuoi")]
        public static void CTSNTaoLuoi()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n=== TẠO LƯỚI SAN NỀN ===");

            using Transaction tr = db.TransactionManager.StartTransaction();
            try
            {
            SmartCommand.Execute("Thiết lập lưới ô vuông", (pm) =>
            {
                // 1. Hiển thị form cài đặt
                SanNenSettingsForm settingsForm = new(tr);
                if (settingsForm.ShowDialog() != DialogResult.OK)
                {
                    throw new OperationCanceledException("Người dùng đã hủy thiết lập.");
                }

                double cellSize = settingsForm.CellSize;
                double textHeight = settingsForm.TextHeight;

                // 2. Chọn ranh giới (Polyline hoặc vẽ hình chữ nhật)
                PromptEntityOptions peo = new("\nChọn Polyline ranh giới hoặc [Rectangle]: ")
                {
                    AllowNone = true
                };
                peo.SetRejectMessage("\nChỉ chọn Polyline!");
                peo.AddAllowedClass(typeof(Polyline), true);
                peo.Keywords.Add("Rectangle");

                PromptEntityResult per = ed.GetEntity(peo);
                Point3d minPt, maxPt;

                if (per.Status == PromptStatus.Keyword && per.StringResult == "Rectangle")
                {
                    // Vẽ hình chữ nhật
                    PromptPointResult ppr1 = ed.GetPoint("\nChọn góc thứ nhất:");
                    if (ppr1.Status != PromptStatus.OK) return;

                    PromptCornerOptions pco = new("\nChọn góc đối diện:", ppr1.Value);
                    PromptPointResult ppr2 = ed.GetCorner(pco);
                    if (ppr2.Status != PromptStatus.OK) return;

                    minPt = new Point3d(
                        Math.Min(ppr1.Value.X, ppr2.Value.X),
                        Math.Min(ppr1.Value.Y, ppr2.Value.Y),
                        0);
                    maxPt = new Point3d(
                        Math.Max(ppr1.Value.X, ppr2.Value.X),
                        Math.Max(ppr1.Value.Y, ppr2.Value.Y),
                        0);
                }
                else if (per.Status == PromptStatus.OK)
                {
                    // Lấy bounding box của Polyline
                    Polyline pl = (Polyline)tr.GetObject(per.ObjectId, OpenMode.ForRead);
                    Extents3d ext = pl.GeometricExtents;
                    minPt = ext.MinPoint;
                    maxPt = ext.MaxPoint;
                }
                else
                {
                    throw new OperationCanceledException("Người dùng đã hủy chọn ranh giới.");
                }

                // 3. Tính số hàng và cột
                double width = maxPt.X - minPt.X;
                double height = maxPt.Y - minPt.Y;
                int cols = (int)Math.Ceiling(width / cellSize);
                int rows = (int)Math.Ceiling(height / cellSize);

                ed.WriteMessage($"\nKích thước vùng: {width:F2} x {height:F2} m");
                ed.WriteMessage($"\nSố ô lưới: {rows} hàng x {cols} cột = {rows * cols} ô");

                // 4. Tạo lưới
                currentGrid = new GradingGrid
                {
                    Origin = minPt,
                    CellSize = cellSize,
                    Rows = rows,
                    Cols = cols
                };

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Tạo layer cho lưới
                string layerName = "CTSN_LUOI";
                EnsureLayer(db, tr, layerName, 3); // Màu xanh lá

                // 5. Vẽ các ô lưới
                pm.SetLimit(rows * cols);
                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        pm.MeterProgress();
                        double x = minPt.X + col * cellSize;
                        double y = minPt.Y + row * cellSize;

                        // Tạo polyline cho ô
                        Polyline cellPl = new();
                        cellPl.AddVertexAt(0, new Point2d(x, y), 0, 0, 0);
                        cellPl.AddVertexAt(1, new Point2d(x + cellSize, y), 0, 0, 0);
                        cellPl.AddVertexAt(2, new Point2d(x + cellSize, y + cellSize), 0, 0, 0);
                        cellPl.AddVertexAt(3, new Point2d(x, y + cellSize), 0, 0, 0);
                        cellPl.Closed = true;
                        cellPl.Layer = layerName;

                        btr.AppendEntity(cellPl);
                        tr.AddNewlyCreatedDBObject(cellPl, true);

                        // Tạo GridCell
                        GridCell cell = new()
                        {
                            Row = row + 1,
                            Col = col + 1,
                            Area = cellSize * cellSize,
                            PolylineId = cellPl.ObjectId,
                            Corners = new GridCorner[]
                            {
                                new() { Position = new Point3d(x, y + cellSize, 0) },           // TL
                                new() { Position = new Point3d(x + cellSize, y + cellSize, 0) }, // TR
                                new() { Position = new Point3d(x + cellSize, y, 0) },           // BR
                                new() { Position = new Point3d(x, y, 0) }                        // BL
                            }
                        };

                        // Lấy cao độ từ Surface nếu có
                        if (!settingsForm.SurfaceTNId.IsNull)
                        {
                            TinSurface surfTN = (TinSurface)tr.GetObject(settingsForm.SurfaceTNId, OpenMode.ForRead);
                            foreach (var corner in cell.Corners)
                            {
                                try
                                {
                                    corner.ElevationTN = surfTN.FindElevationAtXY(corner.Position.X, corner.Position.Y);
                                }
                                catch { corner.ElevationTN = 0; }
                            }
                        }

                        if (!settingsForm.SurfaceTKId.IsNull)
                        {
                            TinSurface surfTK = (TinSurface)tr.GetObject(settingsForm.SurfaceTKId, OpenMode.ForRead);
                            foreach (var corner in cell.Corners)
                            {
                                try
                                {
                                    corner.ElevationTK = surfTK.FindElevationAtXY(corner.Position.X, corner.Position.Y);
                                }
                                catch { corner.ElevationTK = 0; }
                            }
                        }

                        // Thêm text tên ô ở tâm
                        Point3d centerPt = new(x + cellSize / 2, y + cellSize / 2, 0);
                        DBText cellName = new()
                        {
                            Position = centerPt,
                            Height = textHeight,
                            TextString = cell.Name,
                            Layer = layerName,
                            HorizontalMode = TextHorizontalMode.TextCenter,
                            VerticalMode = TextVerticalMode.TextVerticalMid,
                            AlignmentPoint = centerPt
                        };
                        btr.AppendEntity(cellName);
                        tr.AddNewlyCreatedDBObject(cellName, true);

                        currentGrid.Cells.Add(cell);
                    }
                }

                tr.Commit();

                ed.WriteMessage($"\n✓ Đã tạo lưới {rows}x{cols} = {rows * cols} ô");
                ed.WriteMessage("\nSử dụng lệnh CTSN_NhapCaoDo để nhập cao độ TN/TK");
            });
            }
            catch (System.OperationCanceledException)
            {
                // Silent catch for user cancellation
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[ERROR] {ex.Message}");
            }
        }

        /// <summary>
        /// Tính khối lượng đào đắp
        /// </summary>
        [CommandMethod("CTSN_TinhKL")]
        public static void CTSNTinhKL()
        {
            SmartCommand.Execute("Tính khối lượng san nền", (pm) =>
            {
                if (currentGrid == null || currentGrid.Cells.Count == 0)
                {
                    throw new System.Exception("Chưa có lưới! Sử dụng lệnh CTSN_TaoLuoi trước.");
                }

                // Tính toán có vẻ nhanh nhưng vẫn báo cáo tiến độ nếu cần
                pm.SetLimit(1);
                pm.MeterProgress();

                Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage("\n=== TÍNH KHỐI LƯỢNG SAN NỀN ===");
                ed.WriteMessage($"\nTổng số ô: {currentGrid.Cells.Count}");
                ed.WriteMessage($"\nTổng diện tích: {currentGrid.TotalArea:F2} m²");
                ed.WriteMessage($"\n\n--- KẾT QUẢ ---");
                ed.WriteMessage($"\n  Khối lượng ĐÀO: {currentGrid.TotalCutVolume:F2} m³");
                ed.WriteMessage($"\n  Khối lượng ĐẮP: {currentGrid.TotalFillVolume:F2} m³");
                ed.WriteMessage($"\n  Chênh lệch: {currentGrid.TotalFillVolume - currentGrid.TotalCutVolume:F2} m³");
                ed.WriteMessage("\n===============================");
            });
        }

        /// <summary>
        /// Nhập cao độ TN/TK cho các góc
        /// </summary>
        [CommandMethod("CTSN_NhapCaoDo")]
        public static void CTSNNhapCaoDo()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            if (currentGrid == null || currentGrid.Cells.Count == 0)
            {
                ed.WriteMessage("\nChưa có lưới! Sử dụng lệnh CTSN_TaoLuoi trước.");
                return;
            }

            ed.WriteMessage("\n=== NHẬP CAO ĐỘ ===");
            ed.WriteMessage("\nChọn góc ô lưới để nhập cao độ (ESC để kết thúc)");

            while (true)
            {
                PromptPointOptions ppo = new("\nChọn điểm góc [Xong/Undo]: ")
                {
                    AllowNone = true
                };
                ppo.Keywords.Add("Xong");
                ppo.Keywords.Add("Undo");

                PromptPointResult ppr = ed.GetPoint(ppo);

                if (ppr.Status == PromptStatus.Keyword)
                {
                    if (ppr.StringResult == "Xong") break;
                    continue;
                }

                if (ppr.Status != PromptStatus.OK) break;

                Point3d clickPt = ppr.Value;

                // Tìm góc gần nhất
                GridCorner? nearestCorner = null;
                GridCell? parentCell = null;
                double minDist = double.MaxValue;

                foreach (var cell in currentGrid.Cells)
                {
                    foreach (var corner in cell.Corners)
                    {
                        double dist = clickPt.DistanceTo(corner.Position);
                        if (dist < minDist && dist < currentGrid.CellSize / 2)
                        {
                            minDist = dist;
                            nearestCorner = corner;
                            parentCell = cell;
                        }
                    }
                }

                if (nearestCorner == null)
                {
                    ed.WriteMessage("\nKhông tìm thấy góc gần đó!");
                    continue;
                }

                // Nhập cao độ TN
                PromptDoubleOptions pdoTN = new($"\nCao độ TN tại ({nearestCorner.Position.X:F2}, {nearestCorner.Position.Y:F2}):")
                {
                    AllowNegative = true,
                    DefaultValue = nearestCorner.ElevationTN
                };
                PromptDoubleResult pdrTN = ed.GetDouble(pdoTN);
                if (pdrTN.Status == PromptStatus.OK)
                    nearestCorner.ElevationTN = pdrTN.Value;

                // Nhập cao độ TK
                PromptDoubleOptions pdoTK = new($"\nCao độ TK:")
                {
                    AllowNegative = true,
                    DefaultValue = nearestCorner.ElevationTK
                };
                PromptDoubleResult pdrTK = ed.GetDouble(pdoTK);
                if (pdrTK.Status == PromptStatus.OK)
                    nearestCorner.ElevationTK = pdrTK.Value;

                ed.WriteMessage($"\n  TN={nearestCorner.ElevationTN:F3}, TK={nearestCorner.ElevationTK:F3}, ΔH={nearestCorner.DeltaH:F3}");
            }

            ed.WriteMessage("\n✓ Hoàn tất nhập cao độ");
            ed.WriteMessage("\nSử dụng lệnh CTSN_TinhKL để tính khối lượng");
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// Đảm bảo layer tồn tại
        /// </summary>
        private static void EnsureLayer(Database db, Transaction tr, string layerName, short colorIndex)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                LayerTableRecord ltr = new()
                {
                    Name = layerName,
                    Color = AcadColor.FromColorIndex(ColorMethod.ByAci, colorIndex)
                };
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }

        #endregion

        #region Taskbar Command

        /// <summary>
        /// Mở thanh công cụ San Nền
        /// </summary>
        [CommandMethod("CTSN_Taskbar")]
        public static void CTSNTaskbar()
        {
            SanNenTaskbar.ShowTaskbar();
        }

        /// <summary>
        /// Cài đặt san nền
        /// </summary>
        public static void CTSNCaiDat()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n=== CÀI ĐẶT SAN NỀN ===");
            ed.WriteMessage($"\n  Kích thước ô hiện tại: {defaultCellSize} m");
            ed.WriteMessage($"\n  Chiều cao chữ: {defaultTextHeight} mm");

            PromptDoubleOptions pdo = new("\nNhập kích thước ô lưới (m):")
            {
                DefaultValue = defaultCellSize,
                AllowNegative = false
            };
            PromptDoubleResult pdr = ed.GetDouble(pdo);
            if (pdr.Status == PromptStatus.OK)
                defaultCellSize = pdr.Value;

            PromptDoubleOptions pdo2 = new("\nNhập chiều cao chữ (mm):")
            {
                DefaultValue = defaultTextHeight,
                AllowNegative = false
            };
            PromptDoubleResult pdr2 = ed.GetDouble(pdo2);
            if (pdr2.Status == PromptStatus.OK)
                defaultTextHeight = pdr2.Value;

            ed.WriteMessage($"\n✓ Đã cập nhật: Ô = {defaultCellSize}m, Chữ = {defaultTextHeight}mm");
        }

        /// <summary>
        /// Lấy cao độ từ Surface
        /// </summary>
        [CommandMethod("CTSN_Surface")]
        public static void CTSNSurface()
        {
            SmartCommand.Execute("Lấy cao độ từ Surface", (pm) =>
            {
                Document doc = AcadApp.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                if (currentGrid == null || currentGrid.Cells.Count == 0)
                {
                    throw new System.Exception("Chưa có lưới! Sử dụng lệnh CTSN_TaoLuoi trước.");
                }

                using Transaction tr = db.TransactionManager.StartTransaction();
                CivilDocument civilDoc = CivilApplication.ActiveDocument;
                ObjectIdCollection surfaceIds = civilDoc.GetSurfaceIds();

                if (surfaceIds.Count == 0)
                {
                    throw new System.Exception("Không tìm thấy Surface nào trong bản vẽ!");
                }

                ed.WriteMessage("\n=== CHỌN SURFACE ===");
                List<(string Name, ObjectId Id)> surfaces = new();
                int idx = 1;
                foreach (ObjectId id in surfaceIds)
                {
                    if (tr.GetObject(id, OpenMode.ForRead) is TinSurface surf)
                    {
                        surfaces.Add((surf.Name, id));
                        ed.WriteMessage($"\n  [{idx}] {surf.Name}");
                        idx++;
                    }
                }

                PromptIntegerOptions pioTN = new("\nChọn số Surface TN (0 = bỏ qua):") { DefaultValue = 0, AllowNegative = false };
                PromptIntegerResult pirTN = ed.GetInteger(pioTN);
                ObjectId surfTNId = ObjectId.Null;
                if (pirTN.Status == PromptStatus.OK && pirTN.Value > 0 && pirTN.Value <= surfaces.Count)
                    surfTNId = surfaces[pirTN.Value - 1].Id;

                PromptIntegerOptions pioTK = new("\nChọn số Surface TK (0 = bỏ qua):") { DefaultValue = 0, AllowNegative = false };
                PromptIntegerResult pirTK = ed.GetInteger(pioTK);
                ObjectId surfTKId = ObjectId.Null;
                if (pirTK.Status == PromptStatus.OK && pirTK.Value > 0 && pirTK.Value <= surfaces.Count)
                    surfTKId = surfaces[pirTK.Value - 1].Id;

                int cellCount = currentGrid.Cells.Count;
                pm.SetLimit(cellCount);
                
                TinSurface? surfTN = surfTNId.IsNull ? null : (TinSurface)tr.GetObject(surfTNId, OpenMode.ForRead);
                TinSurface? surfTK = surfTKId.IsNull ? null : (TinSurface)tr.GetObject(surfTKId, OpenMode.ForRead);

                foreach (var cell in currentGrid.Cells)
                {
                    pm.MeterProgress();
                    foreach (var corner in cell.Corners)
                    {
                        if (surfTN != null)
                        {
                            try { corner.ElevationTN = surfTN.FindElevationAtXY(corner.Position.X, corner.Position.Y); }
                            catch { }
                        }
                        if (surfTK != null)
                        {
                            try { corner.ElevationTK = surfTK.FindElevationAtXY(corner.Position.X, corner.Position.Y); }
                            catch { }
                        }
                    }
                }
                tr.Commit();
                ed.WriteMessage($"\n✓ Đã cập nhật cao độ từ Surface cho {cellCount} ô lưới.");
            });
        }

        /// <summary>
        /// Xuất bảng khối lượng ra CAD
        /// </summary>
        [CommandMethod("CTSN_XuatBang")]
        public static void CTSNXuatBang()
        {
            SmartCommand.Execute("Xuất bảng san nền", (pm) =>
            {
                Document doc = AcadApp.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                if (currentGrid == null || currentGrid.Cells.Count == 0)
                {
                    throw new System.Exception("Chưa có lưới! Sử dụng lệnh CTSN_TaoLuoi trước.");
                }

                PromptPointResult ppr = ed.GetPoint("\nChọn điểm đặt bảng:");
                if (ppr.Status != PromptStatus.OK) throw new OperationCanceledException();

                using Transaction tr = db.TransactionManager.StartTransaction();
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                string layerName = "CTSN_BANG";
                EnsureLayer(db, tr, layerName, 7); // Trắng

                int numCols = 7;
                int numRows = currentGrid.Cells.Count + 2; // Header + Data + Summary
                pm.SetLimit(numRows);

                Autodesk.AutoCAD.DatabaseServices.Table table = new();
                table.SetSize(numRows, numCols);
                table.Position = ppr.Value;
                table.Layer = layerName;

                // Cài đặt bảng
                table.SetRowHeight(8);
                double[] colWidths = { 10, 20, 25, 25, 25, 30, 30 };
                for (int i = 0; i < numCols; i++)
                    table.Columns[i].Width = colWidths[i];

                table.Rows[0].Height = 15;
                for (int i = 1; i < numRows; i++)
                    table.Rows[i].Height = 10;

                // Header
                string[] headers = { "STT", "Ô", "DT (m²)", "CĐ TN", "CĐ TK", "Đào (m³)", "Đắp (m³)" };
                for (int i = 0; i < numCols; i++)
                {
                    table.Cells[0, i].TextString = headers[i];
                    pm.MeterProgress();
                }

                // Data
                int row = 1;
                foreach (var cell in currentGrid.Cells.OrderBy(c => c.Name))
                {
                    pm.MeterProgress();
                    double avgTN = cell.Corners.Average(c => c.ElevationTN);
                    double avgTK = cell.Corners.Average(c => c.ElevationTK);

                    table.Cells[row, 0].TextString = row.ToString();
                    table.Cells[row, 1].TextString = cell.Name;
                    table.Cells[row, 2].TextString = cell.Area.ToString("F2");
                    table.Cells[row, 3].TextString = avgTN.ToString("F3");
                    table.Cells[row, 4].TextString = avgTK.ToString("F3");
                    table.Cells[row, 5].TextString = cell.CutVolume.ToString("F2");
                    table.Cells[row, 6].TextString = cell.FillVolume.ToString("F2");
                    row++;
                }

                // Summary
                pm.MeterProgress();
                table.Cells[row, 0].TextString = "";
                table.Cells[row, 1].TextString = "TỔNG";
                table.Cells[row, 2].TextString = currentGrid.TotalArea.ToString("F2");
                table.Cells[row, 3].TextString = "";
                table.Cells[row, 4].TextString = "";
                table.Cells[row, 5].TextString = currentGrid.TotalCutVolume.ToString("F2");
                table.Cells[row, 6].TextString = currentGrid.TotalFillVolume.ToString("F2");

                btr.AppendEntity(table);
                tr.AddNewlyCreatedDBObject(table, true);

                tr.Commit();
                ed.WriteMessage($"\n✓ Đã xuất bảng tại ({ppr.Value.X:F2}, {ppr.Value.Y:F2})");
            });
        }

        #endregion
    }
}
