using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.TinhToan.TT_ThuyLucCongTranCommands))]

namespace MyFirstProject.TinhToan
{
    public class TT_ThuyLucCongTranCommands
    {
        [CommandMethod("TT_ThuyLucCongTran")]
        public static void ShowThuyLucCongTran()
        {
            if (TT_ThuyLucCongTranForm.Instance == null || TT_ThuyLucCongTranForm.Instance.IsDisposed)
            {
                TT_ThuyLucCongTranForm.Instance = new TT_ThuyLucCongTranForm();
            }
            AcadApp.ShowModelessDialog(TT_ThuyLucCongTranForm.Instance);
        }
    }

    public class TT_ThuyLucCongTranForm : Form
    {
        public static TT_ThuyLucCongTranForm? Instance;

        // Colors from the image
        private Color bgMain = Color.FromArgb(24, 24, 24);
        private Color bgSidebar = Color.FromArgb(30, 30, 30);
        private Color textPrimary = Color.FromArgb(220, 220, 220);
        private Color accentYellow = Color.FromArgb(255, 215, 0);
        private Color accentRed = Color.FromArgb(255, 60, 90);
        private Color accentBlue = Color.FromArgb(52, 152, 219);
        private Color gridColor = Color.FromArgb(50, 50, 50);
        private Color waterColor = Color.FromArgb(120, 52, 152, 219); // Translucent blue
        private Color groundColor = Color.FromArgb(100, 100, 100);

        private PictureBox picCanvas = null!;
        
        // Inputs
        private TextBox txtQ = null!, txtZ_tran = null!, txtZ_them = null!, txtZ_day = null!, txtZ_haluu = null!, txtHc = null!;
        private TextBox txtB1 = null!, txtN = null!, txtL_doc = null!, txtB_tieu = null!, txtL_tieu = null!, txtL_xa = null!;
        private TextBox txtB_xa = null!, txtI_xa = null!, txtManning = null!;
        
        // Outputs
        private Label lblV1 = null!, lblV2 = null!, lblVn = null!, lblHk = null!, lblHht = null!, lblHcong = null!, lblHdeu = null!, lblL_req = null!, lblLn = null!;

        public TT_ThuyLucCongTranForm()
        {
            InitializeComponent();
            MyFirstProject.FormSignatureHelper.AddSignature(this);
        }

        private void InitializeComponent()
        {
            this.Text = "LÊ TRỌNG NHÂN-T27 - ĐỊNH TUYẾN THỦY LỰC ĐA CHẾ ĐỘ";
            this.Size = new Size(1600, 900);
            this.BackColor = bgMain;
            this.ForeColor = textPrimary;
            this.StartPosition = FormStartPosition.CenterScreen;

            // Header
            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 40, BackColor = Color.FromArgb(15, 15, 15) };
            Label lblTitle = new Label { 
                Text = "SƠ ĐỒ TÍNH TOÁN THỦY LỰC CỐNG / TRÀN", 
                ForeColor = accentYellow, 
                Font = new Font("Segoe UI", 12, FontStyle.Bold), 
                Dock = DockStyle.Fill, 
                TextAlign = ContentAlignment.MiddleCenter 
            };
            pnlHeader.Controls.Add(lblTitle);
            this.Controls.Add(pnlHeader);

            // Left Sidebar
            Panel pnlLeft = new Panel { Dock = DockStyle.Left, Width = 300, BackColor = bgSidebar, AutoScroll = true, Padding = new Padding(10) };
            this.Controls.Add(pnlLeft);

            int yPos = 10;
            Label lblGroupTitle = new Label { Text = "[TH2.1] Cống dốc lớn, lưu lượng lớn", ForeColor = Color.Black, BackColor = accentYellow, Font = new Font("Segoe UI", 9, FontStyle.Bold), Location = new Point(10, yPos), Width = 260, Height = 25, TextAlign = ContentAlignment.MiddleLeft };
            pnlLeft.Controls.Add(lblGroupTitle);
            yPos += 30;

            txtQ = CreateInput(pnlLeft, "01. Lưu lượng Q (m³/s):", "15", ref yPos);
            txtZ_tran = CreateInput(pnlLeft, "02. Cao trình đỉnh tràn/mũi (m):", "14", ref yPos);
            txtZ_them = CreateInput(pnlLeft, "03. Cao trình thềm sau dốc (m):", "2", ref yPos);
            txtZ_day = CreateInput(pnlLeft, "04. Cao trình đáy Bể (m):", "0", ref yPos);
            txtZ_haluu = CreateInput(pnlLeft, "05. Cao trình kênh hạ lưu (m):", "2", ref yPos);
            txtHc = CreateInput(pnlLeft, "06. Chiều cao tĩnh Hc (m):", "4", ref yPos);
            txtB1 = CreateInput(pnlLeft, "07. Bề rộng 1 khoang b (m):", "3", ref yPos);
            txtN = CreateInput(pnlLeft, "08. Số lượng khoang N:", "3", ref yPos);
            txtL_doc = CreateInput(pnlLeft, "09. Chiều dài tuyến dốc (m):", "25", ref yPos);
            txtB_tieu = CreateInput(pnlLeft, "10. Bề rộng Bể tiêu năng (m):", "12", ref yPos);
            txtL_tieu = CreateInput(pnlLeft, "11. Chiều dài bể tiêu năng (m):", "12", ref yPos);
            txtL_xa = CreateInput(pnlLeft, "12. Chiều dài kênh xả (m):", "20", ref yPos);
            txtB_xa = CreateInput(pnlLeft, "13. Bề rộng kênh hạ lưu Bk (m):", "15", ref yPos);
            txtI_xa = CreateInput(pnlLeft, "14. Độ dốc kênh xả (j):", "0.001", ref yPos);
            txtManning = CreateInput(pnlLeft, "15. Độ nhám n (Manning):", "0.025", ref yPos);

            Button btnLoadExcel = new Button { Text = "LOAD TỪ EXCEL", ForeColor = Color.White, BackColor = Color.FromArgb(50, 100, 50), Font = new Font("Segoe UI", 9, FontStyle.Bold), Location = new Point(10, yPos + 10), Width = 125, Height = 35, FlatStyle = FlatStyle.Flat };
            Button btnExportExcel = new Button { Text = "XUẤT RA EXCEL", ForeColor = Color.White, BackColor = Color.FromArgb(100, 50, 50), Font = new Font("Segoe UI", 9, FontStyle.Bold), Location = new Point(145, yPos + 10), Width = 125, Height = 35, FlatStyle = FlatStyle.Flat };
            btnLoadExcel.FlatAppearance.BorderSize = 0;
            btnExportExcel.FlatAppearance.BorderSize = 0;
            pnlLeft.Controls.Add(btnLoadExcel);
            pnlLeft.Controls.Add(btnExportExcel);

            // Right Sidebar
            Panel pnlRight = new Panel { Dock = DockStyle.Right, Width = 350, BackColor = bgSidebar, AutoScroll = true, Padding = new Padding(10) };
            this.Controls.Add(pnlRight);
            
            // Schematic Placeholder Image
            PictureBox picSchematic = new PictureBox { Location = new Point(10, 10), Width = 310, Height = 200, BackColor = Color.Black, BorderStyle = BorderStyle.FixedSingle };
            pnlRight.Controls.Add(picSchematic);

            Label lblResHeader = new Label { Text = "=== THIẾT KẾ KẾT QUẢ ===", ForeColor = accentYellow, Font = new Font("Segoe UI", 12, FontStyle.Bold), Location = new Point(10, 230), AutoSize = true };
            pnlRight.Controls.Add(lblResHeader);

            int ry = 270;
            lblV1 = CreateResultRow(pnlRight, "VẬN TỐC TRƯỚC NHẢY (V1)", "16.54 m/s", Color.LightPink, ref ry);
            lblV2 = CreateResultRow(pnlRight, "VẬN TỐC SAU NHẢY (V2)", "0.62 m/s", Color.Cyan, ref ry);
            lblVn = CreateResultRow(pnlRight, "VẬN TỐC HẠ LƯU (V_N)", "1.10 m/s", Color.LightGreen, ref ry);
            lblHk = CreateResultRow(pnlRight, "CHIỀU DÀI HẠ LƯU (L_K)", "0.909 m", Color.MediumPurple, ref ry);
            lblHht = CreateResultRow(pnlRight, "CỘT NƯỚC HK KÊU, LƯU THU NG", "0.467 m", Color.Pink, ref ry);
            lblHcong = CreateResultRow(pnlRight, "CỘT NƯỚC H_K CỐNG (Hk_cong)", "0.657 m", Color.Pink, ref ry);
            lblHdeu = CreateResultRow(pnlRight, "CỘT NƯỚC CHẢY ĐỀU (H0)", "0.909 m", Color.LimeGreen, ref ry);
            lblL_req = CreateResultRow(pnlRight, "CHIỀU CAO ĐỂ YÊU CẦU (P)", "1.207 m", Color.Khaki, ref ry);
            lblLn = CreateResultRow(pnlRight, "CHIỀU DÀI RẼ (Ln)", "8.73 m", Color.Wheat, ref ry);



            // Center Panel Canvas
            Panel pnlCenter = new Panel { Dock = DockStyle.Fill };
            this.Controls.Add(pnlCenter);

            picCanvas = new PictureBox { Dock = DockStyle.Fill };
            picCanvas.Paint += PicCanvas_Paint;
            picCanvas.Resize += (s, e) => picCanvas.Invalidate();
            pnlCenter.Controls.Add(picCanvas);
        }

        private TextBox CreateInput(Panel pnl, string labelText, string defVal, ref int yPos)
        {
            Label lbl = new Label { Text = labelText, ForeColor = Color.LightGray, Font = new Font("Segoe UI", 8), Location = new Point(10, yPos), AutoSize = true };
            pnl.Controls.Add(lbl);
            yPos += 18;

            TextBox txt = new TextBox { Text = defVal, BackColor = Color.FromArgb(40, 40, 40), ForeColor = Color.White, Font = new Font("Consolas", 10, FontStyle.Bold), Location = new Point(10, yPos), Width = 260, BorderStyle = BorderStyle.FixedSingle };
            txt.TextChanged += (s, e) => picCanvas.Invalidate();
            pnl.Controls.Add(txt);
            yPos += 30;
            return txt;
        }

        private Label CreateResultRow(Panel pnl, string title, string val, Color valColor, ref int yPos)
        {
            Label lblGroup = new Label { Text = title, ForeColor = Color.Gray, Font = new Font("Segoe UI", 7, FontStyle.Bold), Location = new Point(10, yPos), AutoSize = true };
            pnl.Controls.Add(lblGroup);
            yPos += 15;

            Panel bgPanel = new Panel { Location = new Point(10, yPos), Width = 310, Height = 25, BackColor = Color.FromArgb(40, 40, 40), BorderStyle = BorderStyle.FixedSingle };
            Label lblVal = new Label { Text = val, ForeColor = valColor, Font = new Font("Consolas", 10, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter };
            bgPanel.Controls.Add(lblVal);
            pnl.Controls.Add(bgPanel);
            
            yPos += 35;
            return lblVal;
        }

        private void PicCanvas_Paint(object? sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            int w = picCanvas.Width;
            int h = picCanvas.Height;

            // Draw Background Grids
            Pen gridPen = new Pen(gridColor, 1);
            int step = 40;
            for (int x = 0; x < w; x += step) g.DrawLine(gridPen, x, 0, x, h);
            for (int y = 0; y < h; y += step) g.DrawLine(gridPen, 0, y, w, y);

            // Origin and scale
            int ox = 50;
            int oy = h - 200; // Y config: Base elevation 0
            float scaleY = 30f;
            float scaleX = 20f;

            // Reading parameters
            float z_tran = TryParse(txtZ_tran.Text, 14f);
            float z_them = TryParse(txtZ_them.Text, 2f);
            float z_day = TryParse(txtZ_day.Text, 0f);
            float z_haluu = TryParse(txtZ_haluu.Text, 2f);

            float l_doc = TryParse(txtL_doc.Text, 25f);
            float l_tieu = TryParse(txtL_tieu.Text, 12f);
            float l_xa = TryParse(txtL_xa.Text, 20f);

            // X coordinates
            float x0 = ox;
            float x1 = x0 + 10 * scaleX; // Reservoir width
            float x2 = x1 + l_doc * scaleX; // Spillway slope
            float x3 = x2 + l_tieu * scaleX; // Stilling basin
            float x4 = x3 + 5 * scaleX; // Step upward
            float x5 = x4 + l_xa * scaleX; // Downstream channel

            // Path for Ground
            GraphicsPath ground = new GraphicsPath();
            ground.AddLine(x0, oy - z_tran * scaleY, x1, oy - z_tran * scaleY);
            ground.AddLine(x1, oy - z_tran * scaleY, x2, oy - z_day * scaleY);
            ground.AddLine(x2, oy - z_day * scaleY, x3, oy - z_day * scaleY);
            ground.AddLine(x3, oy - z_day * scaleY, x4, oy - z_haluu * scaleY);
            ground.AddLine(x4, oy - z_haluu * scaleY, x5, oy - z_haluu * scaleY);
            ground.AddLine(x5, h, x0, h);
            g.FillPath(new SolidBrush(groundColor), ground);
            
            Pen borderPen = new Pen(Color.FromArgb(180, 180, 180), 2);
            g.DrawPath(borderPen, ground);

            // Path for Water
            float w_tran = z_tran + 2f; // Water head over spillway
            float w_day = z_day + 0.5f; // Supercritical flow
            float w_jump = z_day + 3f;  // Subcritical flow inside basin
            float w_haluu = z_haluu + 1f;

            GraphicsPath water = new GraphicsPath();
            water.AddLine(x0, oy - w_tran * scaleY, x1, oy - w_tran * scaleY); // Reservoir level
            
            // Draw profile curve (simplified cubic bezier)
            water.AddBezier(
                new PointF(x1, oy - w_tran * scaleY), 
                new PointF(x1 + 100, oy - w_tran * scaleY), 
                new PointF(x2 - 100, oy - w_day * scaleY), 
                new PointF(x2, oy - w_day * scaleY)
            );
            
            // Hydraulic jump section
            water.AddLine(x2, oy - w_day * scaleY, x2 + 30, oy - w_day * scaleY); // Supercritical length
            water.AddBezier(
                new PointF(x2 + 30, oy - w_day * scaleY),
                new PointF(x2 + 50, oy - w_jump * scaleY),
                new PointF(x2 + 70, oy - w_jump * scaleY),
                new PointF(x3, oy - w_jump * scaleY)
            );

            // Downstream step
            water.AddLine(x3, oy - w_jump * scaleY, x4, oy - w_haluu * scaleY);
            water.AddLine(x4, oy - w_haluu * scaleY, x5, oy - w_haluu * scaleY);
            
            // Close the water region down to ground level
            water.AddLine(x5, oy - z_haluu * scaleY, x4, oy - z_haluu * scaleY);
            water.AddLine(x4, oy - z_haluu * scaleY, x3, oy - z_day * scaleY);
            water.AddLine(x3, oy - z_day * scaleY, x2, oy - z_day * scaleY);
            water.AddLine(x2, oy - z_day * scaleY, x1, oy - z_tran * scaleY);
            water.AddLine(x1, oy - z_tran * scaleY, x0, oy - z_tran * scaleY);

            g.FillPath(new SolidBrush(waterColor), water);
            
            Pen waterSurfacePen = new Pen(Color.FromArgb(100, 200, 255), 2);
            // Redraw top surface line only
            g.DrawCurve(waterSurfacePen, new PointF[] {
                new PointF(x0, oy - w_tran * scaleY), new PointF(x1, oy - w_tran * scaleY)
            });
            
            // Labeling
            Font f = new Font("Consolas", 10, FontStyle.Bold);
            g.DrawString("H_o = 0.91", f, Brushes.Cyan, x3 + 20, oy - w_jump * scaleY - 20);
            g.DrawString("H_k = 0.64", f, Brushes.Magenta, x2 + 10, oy - w_day * scaleY - 20);
            
            // Y-Axis Labels
            Font axisFont = new Font("Segoe UI", 8);
            for (int i = -10; i <= 15; i++)
            {
                float yy = oy - i * scaleY;
                g.DrawString(i.ToString(), axisFont, Brushes.White, 20, yy - 6);
                g.DrawLine(Pens.Gray, 45, yy, 50, yy);
            }
        }

        private float TryParse(string val, float def)
        {
            if (float.TryParse(val, out float res)) return res;
            return def;
        }
    }
}
