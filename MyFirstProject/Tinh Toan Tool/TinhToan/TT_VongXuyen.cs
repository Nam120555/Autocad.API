using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(TinhToan.TT_VongXuyenCommands))]

namespace TinhToan
{
    public class TT_VongXuyenCommands
    {
        [CommandMethod("TT_VongXuyen")]
        public static void ShowVongXuyen()
        {
            if (TT_VongXuyenForm.Instance == null || TT_VongXuyenForm.Instance.IsDisposed)
            {
                TT_VongXuyenForm.Instance = new TT_VongXuyenForm();
            }
            AcadApp.ShowModelessDialog(TT_VongXuyenForm.Instance);
        }
    }

    public class TT_VongXuyenForm : Form
    {
        public static TT_VongXuyenForm? Instance;

        private Color bgMain = Color.FromArgb(30, 30, 30);
        private Color bgSidebar = Color.FromArgb(38, 38, 42);
        private Color bgInput = Color.FromArgb(51, 51, 53);
        private Color textPrimary = Color.FromArgb(240, 240, 240);
        private Color accentGreen = Color.FromArgb(10, 240, 120);
        private Color accentBlue = Color.FromArgb(0, 191, 255);
        private Color accentRed = Color.FromArgb(255, 60, 90);
        private Color accentYellow = Color.FromArgb(255, 190, 0);
        private Color accentPurple = Color.FromArgb(180, 80, 255);

        private TextBox txtVtk = null!, txtTL = null!, txtMu = null!, txtInut = null!;
        private TextBox txtL = null!, txtB1 = null!, txtD1 = null!;

        private Label lblVnut = null!, lblRxe = null!, lblE1 = null!, lblDdt = null!;
        private DoubleBufferedPanel pnlCanvas = null!;

        public static double CurrentRxe = 13.5;
        public static double CurrentE1 = 2.5;
        public static double CurrentDdt = 20.0;
        private double currentVnut = 20;

        public TT_VongXuyenForm()
        {
            InitializeComponent();
            CalculateVx();
            MyFirstProject.FormSignatureHelper.AddSignature(this);
        }

        private void InitializeComponent()
        {
            this.Text = "TÍNH TOÁN NÚT GIAO VÒNG XUYẾN - CIVIL TOOL";
            this.Size = new Size(1100, 750);
            this.BackColor = bgMain;
            this.ForeColor = textPrimary;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.StartPosition = FormStartPosition.CenterScreen;

            // SIDEBAR LEFT
            Panel pnlLeft = new Panel { Dock = DockStyle.Left, Width = 310, BackColor = bgSidebar, AutoScroll = true };
            this.Controls.Add(pnlLeft);

            pnlLeft.Controls.Add(new Label { Text = "■ THÔNG SỐ VÒNG XUYẾN", ForeColor = accentBlue, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, 15), AutoSize = true });

            int yPos = 45;
            txtVtk = CreateInput(pnlLeft, "Vận tốc TK tuyến (V_tk) [km/h]", "50", ref yPos);
            txtTL = CreateInput(pnlLeft, "Tỉ lệ giảm tốc vào nút", "0.40", ref yPos);
            txtMu = CreateInput(pnlLeft, "Hệ số lực ngang (μ)", "0.25", ref yPos);
            txtInut = CreateInput(pnlLeft, "Dốc ngang mặt đường nút (i)", "0.005", ref yPos);

            yPos += 10;
            pnlLeft.Controls.Add(new Label { Text = "■ THÔNG SỐ XE THIẾT KẾ", ForeColor = accentGreen, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, yPos), AutoSize = true });
            yPos += 30;
            txtL = CreateInput(pnlLeft, "Trục xe (L) [m]", "7.77", ref yPos);
            txtB1 = CreateInput(pnlLeft, "Bề rộng làn 1 (b_1) [m]", "3.5", ref yPos);
            txtD1 = CreateInput(pnlLeft, "Dải an toàn giáp đảo (d_1) [m]", "0.5", ref yPos);

            // SIDEBAR RIGHT
            Panel pnlRight = new Panel { Dock = DockStyle.Right, Width = 310, BackColor = bgSidebar };
            this.Controls.Add(pnlRight);

            pnlRight.Controls.Add(new Label { Text = "■ KẾT QUẢ TÍNH TOÁN", ForeColor = accentRed, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, 15), AutoSize = true });

            int ry = 45;
            lblVnut = CreateResult(pnlRight, "Vận tốc xe trong nút (V_nut)", "20.00 km/h", accentGreen, ref ry);
            lblRxe = CreateResult(pnlRight, "Bán kính tối thiểu (R_xe)", "12.35 m", accentYellow, ref ry);
            lblE1 = CreateResult(pnlRight, "Độ mở rộng (e_1)", "2.25 m", accentPurple, ref ry);

            ry += 10;
            pnlRight.Controls.Add(new Label { Text = "ĐƯỜNG KÍNH ĐẢO GIỮA YÊU CẦU", ForeColor = textPrimary, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, ry), AutoSize = true });
            lblDdt = new Label { Text = "20.00 m", ForeColor = accentRed, Font = new System.Drawing.Font("Consolas", 24, FontStyle.Bold), Location = new Point(15, ry + 25), AutoSize = true };
            pnlRight.Controls.Add(lblDdt);

            // CENTER CANVAS 2D
            Panel pnlCenter = new Panel { Dock = DockStyle.Fill };
            this.Controls.Add(pnlCenter);

            Label lblTitle = new Label
            {
                Text = "MẶT BẰNG NÚT GIAO VÒNG XUYẾN (2D)",
                ForeColor = Color.White,
                Font = new System.Drawing.Font("Segoe UI", 12, FontStyle.Bold),
                Dock = DockStyle.Top, TextAlign = ContentAlignment.MiddleCenter, Height = 40, AutoSize = false
            };
            pnlCenter.Controls.Add(lblTitle);

            pnlCanvas = new DoubleBufferedPanel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(25, 28, 35) };
            pnlCanvas.Paint += PnlCanvas_Paint;
            pnlCenter.Controls.Add(pnlCanvas);
            lblTitle.BringToFront();
        }

        private TextBox CreateInput(Panel pnl, string label, string def, ref int y)
        {
            pnl.Controls.Add(new Label { Text = label, ForeColor = textPrimary, Font = new System.Drawing.Font("Segoe UI", 9), Location = new Point(15, y), AutoSize = true });
            var txt = new TextBox { Text = def, BackColor = bgInput, ForeColor = accentBlue, Font = new System.Drawing.Font("Consolas", 10, FontStyle.Bold), Location = new Point(15, y + 20), Width = 260, BorderStyle = BorderStyle.FixedSingle };
            txt.TextChanged += (s, e) => CalculateVx();
            pnl.Controls.Add(txt);
            y += 50;
            return txt;
        }

        private Label CreateResult(Panel pnl, string label, string def, Color clr, ref int y)
        {
            pnl.Controls.Add(new Label { Text = label, ForeColor = textPrimary, Font = new System.Drawing.Font("Segoe UI", 9), Location = new Point(15, y), AutoSize = true });
            var lbl = new Label { Text = def, ForeColor = clr, Font = new System.Drawing.Font("Consolas", 14, FontStyle.Bold), Location = new Point(15, y + 20), AutoSize = true };
            pnl.Controls.Add(lbl);
            y += 60;
            return lbl;
        }

        private void CalculateVx()
        {
            try
            {
                double v_tk = double.Parse(txtVtk.Text);
                double tl = double.Parse(txtTL.Text);
                double mu = double.Parse(txtMu.Text);
                double i_nut = double.Parse(txtInut.Text);
                double l = double.Parse(txtL.Text);
                double b1 = double.Parse(txtB1.Text);
                double d1 = double.Parse(txtD1.Text);

                double v_nut = v_tk * tl;
                double r_xe = Math.Pow(v_nut, 2) / (127.0 * (mu + i_nut));
                double e1 = 0;
                if (r_xe > 0)
                    e1 = Math.Pow(l, 2) / (2 * r_xe + b1) + (0.05 * v_nut) / Math.Sqrt(r_xe);
                double d_dt = 2.0 * (r_xe - b1 / 2.0 - e1 / 2.0 - d1);

                CurrentRxe = r_xe; CurrentE1 = e1; CurrentDdt = d_dt; currentVnut = v_nut;
                lblVnut.Text = $"{v_nut:F2} km/h";
                lblRxe.Text = $"{r_xe:F2} m";
                lblE1.Text = $"{e1:F2} m";
                lblDdt.Text = $"{d_dt:F2} m";

                pnlCanvas?.Invalidate();
            }
            catch { }
        }

        private void PnlCanvas_Paint(object? sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            float cx = pnlCanvas.Width / 2f;
            float cy = pnlCanvas.Height / 2f;

            double r_xe = CurrentRxe;
            double d_dt = CurrentDdt;
            double e1 = CurrentE1;
            double rIsland = d_dt / 2.0;
            double rOuter = r_xe + e1;
            double laneW = 0;
            try { laneW = double.Parse(txtB1.Text); } catch { laneW = 3.5; }

            // Scale để fit vào canvas
            double maxR = rOuter + 30;
            float scale = (float)(Math.Min(cx, cy) * 0.85 / maxR);

            g.TranslateTransform(cx, cy);

            // --- Mặt đường vòng xuyến (phần xe chạy = vành khuyên) ---
            float rOutPx = (float)(rOuter * scale);
            float rInPx = (float)(rIsland * scale);

            using (var bRoad = new SolidBrush(Color.FromArgb(50, 55, 60)))
            using (var bIsland = new SolidBrush(Color.FromArgb(50, 100, 50)))
            using (var bGrass = new SolidBrush(Color.FromArgb(35, 50, 35)))
            using (var pOuter = new Pen(Color.Cyan, 2f))
            using (var pInner = new Pen(Color.FromArgb(100, 200, 100), 2f))
            using (var pDim = new Pen(Color.FromArgb(200, 200, 200), 1f) { DashPattern = new float[] { 4, 3 } })
            using (var pCenter = new Pen(Color.Yellow, 1.5f) { DashPattern = new float[] { 8, 4, 2, 4 } })
            using (var pArrow = new Pen(Color.White, 1.5f))
            {
                // Nền cỏ
                g.FillRectangle(bGrass, -cx, -cy, cx * 2, cy * 2);

                // 4 nhánh đường thẳng
                float legW = (float)(laneW * 2 * scale);
                float legLen = (float)(maxR * scale + 50);
                g.FillRectangle(bRoad, -legW / 2, -legLen, legW, legLen);   // Bắc
                g.FillRectangle(bRoad, -legW / 2, 0, legW, legLen);         // Nam
                g.FillRectangle(bRoad, -legLen, -legW / 2, legLen, legW);   // Tây
                g.FillRectangle(bRoad, 0, -legW / 2, legLen, legW);         // Đông

                // Phần xe chạy vòng xuyến (vành khuyên)
                if (rOutPx > 5)
                {
                    g.FillEllipse(bRoad, -rOutPx, -rOutPx, rOutPx * 2, rOutPx * 2);
                }

                // Đảo tròn trung tâm
                if (rInPx > 2)
                {
                    g.FillEllipse(bIsland, -rInPx, -rInPx, rInPx * 2, rInPx * 2);
                    g.DrawEllipse(pInner, -rInPx, -rInPx, rInPx * 2, rInPx * 2);
                }

                // Viền ngoài phần xe chạy
                if (rOutPx > 5)
                    g.DrawEllipse(pOuter, -rOutPx, -rOutPx, rOutPx * 2, rOutPx * 2);

                // Tim đường vòng xuyến (đường trung bình)
                float rCenterPx = (rInPx + rOutPx) / 2f;
                if (rCenterPx > 5)
                    g.DrawEllipse(pCenter, -rCenterPx, -rCenterPx, rCenterPx * 2, rCenterPx * 2);

                // Viền mép đường thẳng
                g.DrawLine(pOuter, -legW / 2, -rOutPx, -legW / 2, -legLen);
                g.DrawLine(pOuter, legW / 2, -rOutPx, legW / 2, -legLen);
                g.DrawLine(pOuter, -legW / 2, rOutPx, -legW / 2, legLen);
                g.DrawLine(pOuter, legW / 2, rOutPx, legW / 2, legLen);
                g.DrawLine(pOuter, -rOutPx, -legW / 2, -legLen, -legW / 2);
                g.DrawLine(pOuter, -rOutPx, legW / 2, -legLen, legW / 2);
                g.DrawLine(pOuter, rOutPx, -legW / 2, legLen, -legW / 2);
                g.DrawLine(pOuter, rOutPx, legW / 2, legLen, legW / 2);

                // ── Kích thước (Dimension) ──
                // D_đảo
                g.DrawLine(pDim, -rInPx, 0, rInPx, 0);
                DrawDimText(g, 0, -12, $"D_đảo = {d_dt:F1}m", Color.FromArgb(100, 200, 100));

                // R_xe (bán kính)
                float rXePx = (float)(r_xe * scale);
                g.DrawLine(pDim, 0, 0, 0, -rXePx);
                DrawDimText(g, 5, -rXePx / 2, $"R_xe = {r_xe:F1}m", Color.Yellow);

                // Mũi tên chỉ hướng xe chạy (ngược chiều kim đồng hồ)
                DrawCircleArrow(g, pArrow, rCenterPx, 30);
                DrawCircleArrow(g, pArrow, rCenterPx, 120);
                DrawCircleArrow(g, pArrow, rCenterPx, 210);
                DrawCircleArrow(g, pArrow, rCenterPx, 300);
            }

            g.ResetTransform();

            // Tiêu đề thông số
            using (var fTitle = new System.Drawing.Font("Consolas", 11, FontStyle.Bold))
            using (var fInfo = new System.Drawing.Font("Segoe UI", 9))
            {
                g.DrawString($"R_xe = {r_xe:F2} m  |  D_đảo = {d_dt:F2} m  |  e₁ = {e1:F2} m", fTitle, Brushes.Cyan, 15, pnlCanvas.Height - 30);
            }
        }

        private void DrawDimText(Graphics g, float x, float y, string text, Color color)
        {
            using (var f = new System.Drawing.Font("Consolas", 9, FontStyle.Bold))
            using (var b = new SolidBrush(color))
            {
                g.DrawString(text, f, b, x, y);
            }
        }

        private void DrawCircleArrow(Graphics g, Pen pen, float radius, float angleDeg)
        {
            float a1 = angleDeg - 15;
            float sweep = 30;
            g.DrawArc(pen, -radius, -radius, radius * 2, radius * 2, a1, sweep);

            // Mũi tên ở cuối cung
            double endRad = (angleDeg + 15) * Math.PI / 180.0;
            float ex = radius * (float)Math.Cos(endRad);
            float ey = radius * (float)Math.Sin(endRad);

            // Hướng tiếp tuyến (ngược chiều kim đồng hồ)
            double tangent = endRad + Math.PI / 2;
            float arrowLen = 8f;
            float ax1 = ex + arrowLen * (float)Math.Cos(tangent + 0.4);
            float ay1 = ey + arrowLen * (float)Math.Sin(tangent + 0.4);
            float ax2 = ex + arrowLen * (float)Math.Cos(tangent - 0.4);
            float ay2 = ey + arrowLen * (float)Math.Sin(tangent - 0.4);

            g.DrawLine(pen, ex, ey, ax1, ay1);
            g.DrawLine(pen, ex, ey, ax2, ay2);
        }
    }
}
