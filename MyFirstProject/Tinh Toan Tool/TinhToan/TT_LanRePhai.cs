using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(TinhToan.TT_LanRePhaiCommands))]

namespace TinhToan
{
    public class TT_LanRePhaiCommands
    {
        [CommandMethod("TT_LanRePhai")]
        public static void ShowLanRePhai()
        {
            if (TT_LanRePhaiForm.Instance == null || TT_LanRePhaiForm.Instance.IsDisposed)
                TT_LanRePhaiForm.Instance = new TT_LanRePhaiForm();
            AcadApp.ShowModelessDialog(TT_LanRePhaiForm.Instance);
        }
    }

    public class TT_LanRePhaiForm : Form
    {
        public static TT_LanRePhaiForm? Instance;
        private Color bgMain = Color.FromArgb(30, 30, 30);
        private Color bgSidebar = Color.FromArgb(38, 38, 42);
        private Color bgInput = Color.FromArgb(51, 51, 53);
        private Color textPrimary = Color.FromArgb(240, 240, 240);
        private Color accentGreen = Color.FromArgb(10, 240, 120);
        private Color accentBlue = Color.FromArgb(0, 191, 255);
        private Color accentYellow = Color.FromArgb(255, 190, 0);
        private Color accentRed = Color.FromArgb(255, 60, 90);

        private TextBox txtVa = null!, txtVb = null!, txtD = null!, txtA = null!, txtLx = null!, txtLc = null!;
        private Label lblLn = null!, lblLg = null!, lblLv = null!, lblL = null!;
        private DoubleBufferedPanel pnlCanvas = null!;

        public static double CurrentLn = 0, CurrentLv = 0, CurrentLx = 0, CurrentL = 0;
        private double currentLg = 0;

        public TT_LanRePhaiForm()
        {
            this.Text = "THIẾT KẾ LÀN RẼ PHẢI - CIVIL TOOL";
            this.Size = new Size(1100, 750);
            this.BackColor = bgMain; this.ForeColor = textPrimary;
            this.StartPosition = FormStartPosition.CenterScreen;

            Panel pnlLeft = new Panel { Dock = DockStyle.Left, Width = 310, BackColor = bgSidebar, AutoScroll = true };
            this.Controls.Add(pnlLeft);
            pnlLeft.Controls.Add(new Label { Text = "■ THÔNG SỐ (TCVN 13592:2022)", ForeColor = accentBlue, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, 15), AutoSize = true });

            int yPos = 45;
            txtVa = CreateInput(pnlLeft, "Vận tốc xe chạy (V_a) [km/h]", "50", ref yPos);
            txtVb = CreateInput(pnlLeft, "Vận tốc điểm dừng (V_b) [km/h]", "0", ref yPos);
            txtD = CreateInput(pnlLeft, "Bề rộng làn xe (d) [m]", "3.5", ref yPos);
            txtA = CreateInput(pnlLeft, "Gia tốc hãm xe (a) [m/s²]", "2.0", ref yPos);
            txtLx = CreateInput(pnlLeft, "Chiều dài hàng đỗ chờ (L_x) [m]", "30.0", ref yPos);
            txtLc = CreateInput(pnlLeft, "Đoạn chuyển tối thiểu (L_c) [m]", "40.0", ref yPos);

            Panel pnlRight = new Panel { Dock = DockStyle.Right, Width = 310, BackColor = bgSidebar };
            this.Controls.Add(pnlRight);
            pnlRight.Controls.Add(new Label { Text = "■ KẾT QUẢ CHIỀU DÀI", ForeColor = accentRed, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, 15), AutoSize = true });

            int ry = 45;
            lblLn = CreateResult(pnlRight, "Chiều dài hình nêm (L_n)", "29.17 m", accentGreen, ref ry);
            lblLg = CreateResult(pnlRight, "Chiều dài giảm tốc (L_g)", "48.08 m", accentGreen, ref ry);
            lblLv = CreateResult(pnlRight, "Chiều dài chuyển làn (L_v)", "48.08 m", accentYellow, ref ry);

            ry += 20;
            pnlRight.Controls.Add(new Label { Text = "TỔNG CHIỀU DÀI LÀN RẼ PHẢI (L)", ForeColor = Color.White, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, ry), AutoSize = true });
            lblL = new Label { Text = "107.25 m", ForeColor = accentRed, Font = new System.Drawing.Font("Consolas", 24, FontStyle.Bold), Location = new Point(15, ry + 25), AutoSize = true };
            pnlRight.Controls.Add(lblL);

            // CENTER CANVAS 2D
            Panel pnlCenter = new Panel { Dock = DockStyle.Fill };
            this.Controls.Add(pnlCenter);
            var lblTitle = new Label { Text = "MẶT BẰNG LÀN RẼ PHẢI (2D)", ForeColor = Color.White, Font = new System.Drawing.Font("Segoe UI", 12, FontStyle.Bold), Dock = DockStyle.Top, TextAlign = ContentAlignment.MiddleCenter, Height = 40, AutoSize = false };
            pnlCenter.Controls.Add(lblTitle);

            pnlCanvas = new DoubleBufferedPanel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(25, 28, 35) };
            pnlCanvas.Paint += PnlCanvas_Paint;
            pnlCenter.Controls.Add(pnlCanvas);
            lblTitle.BringToFront();

            Calculate();
            MyFirstProject.FormSignatureHelper.AddSignature(this);
        }

        private TextBox CreateInput(Panel pnl, string label, string def, ref int y)
        {
            pnl.Controls.Add(new Label { Text = label, ForeColor = textPrimary, Font = new System.Drawing.Font("Segoe UI", 9), Location = new Point(15, y), AutoSize = true });
            var txt = new TextBox { Text = def, BackColor = bgInput, ForeColor = accentBlue, Font = new System.Drawing.Font("Consolas", 10), Location = new Point(15, y + 20), Width = 260 };
            txt.TextChanged += (s, e) => Calculate();
            pnl.Controls.Add(txt);
            y += 55;
            return txt;
        }

        private Label CreateResult(Panel pnl, string label, string def, Color clr, ref int y)
        {
            pnl.Controls.Add(new Label { Text = label, ForeColor = textPrimary, Font = new System.Drawing.Font("Segoe UI", 9), Location = new Point(15, y), AutoSize = true });
            var lbl = new Label { Text = def, ForeColor = clr, Font = new System.Drawing.Font("Consolas", 14, FontStyle.Bold), Location = new Point(15, y + 20), AutoSize = true };
            pnl.Controls.Add(lbl);
            y += 65;
            return lbl;
        }

        private void Calculate()
        {
            try
            {
                double va = double.Parse(txtVa.Text);
                double vb = double.Parse(txtVb.Text);
                double d = double.Parse(txtD.Text);
                double a = double.Parse(txtA.Text);
                double lx = double.Parse(txtLx.Text);
                double lc = double.Parse(txtLc.Text);

                double ln = va * d / 6.0;
                double lg = a != 0 ? (Math.Pow(va, 2) - Math.Pow(vb, 2)) / (26.0 * a) : 0;
                double lv = Math.Max(lc, Math.Max(0, lg));
                double l = ln + lv + lx;

                CurrentLn = ln; CurrentLv = lv; CurrentLx = lx; CurrentL = l; currentLg = lg;
                lblLn.Text = $"{ln:F2} m"; lblLg.Text = $"{lg:F2} m"; lblLv.Text = $"{lv:F2} m"; lblL.Text = $"{l:F2} m";
                pnlCanvas?.Invalidate();
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
        }

        private void PnlCanvas_Paint(object? sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            float w = pnlCanvas.Width, h = pnlCanvas.Height;
            double ln = CurrentLn, lv = CurrentLv, lx = CurrentLx, totalL = CurrentL;
            if (totalL < 1) return;

            // Scale
            float margin = 80;
            float drawW = w - margin * 2;
            float scale = (float)(drawW / (totalL + 40));
            float mainY = h * 0.4f;
            float laneH = 35;
            float sideH = 30;
            float startX = margin;

            using (var bRoad = new SolidBrush(Color.FromArgb(50, 55, 60)))
            using (var bSide = new SolidBrush(Color.FromArgb(60, 75, 85)))
            using (var pEdge = new Pen(Color.Cyan, 2f))
            using (var pTim = new Pen(Color.Yellow, 1.5f) { DashPattern = new float[] { 10, 5, 2, 5 } })
            using (var pDim = new Pen(Color.FromArgb(200, 200, 200), 1f))
            using (var pNem = new Pen(Color.FromArgb(255, 150, 50), 2f))
            {
                float totalPx = (float)(totalL * scale);

                // Đường chính
                g.FillRectangle(bRoad, startX - 20, mainY, totalPx + 40, laneH);
                g.DrawLine(pEdge, startX - 20, mainY, startX + totalPx + 20, mainY);
                g.DrawLine(pEdge, startX - 20, mainY + laneH, startX + totalPx + 20, mainY + laneH);
                g.DrawLine(pTim, startX - 20, mainY + laneH / 2, startX + totalPx + 20, mainY + laneH / 2);

                // Hình nêm (tam giác vuốt)
                float lnPx = (float)(ln * scale);
                float lvPx = (float)(lv * scale);
                float lxPx = (float)(lx * scale);

                PointF[] nemPts = new PointF[] {
                    new PointF(startX, mainY + laneH),
                    new PointF(startX + lnPx, mainY + laneH),
                    new PointF(startX + lnPx, mainY + laneH + sideH),
                };
                g.FillPolygon(bSide, nemPts);
                g.DrawPolygon(pNem, nemPts);

                // Làn giảm tốc (song song)
                float lvStart = startX + lnPx;
                g.FillRectangle(bSide, lvStart, mainY + laneH, lvPx, sideH);
                g.DrawRectangle(pEdge, lvStart, mainY + laneH, lvPx, sideH);

                // Làn đỗ chờ
                float lxStart = lvStart + lvPx;
                g.FillRectangle(new SolidBrush(Color.FromArgb(70, 50, 50)), lxStart, mainY + laneH, lxPx, sideH);
                g.DrawRectangle(new Pen(Color.FromArgb(255, 80, 80), 2f), lxStart, mainY + laneH, lxPx, sideH);

                // ── Dimension lines ──
                float dimY = mainY + laneH + sideH + 25;
                DrawDimLine(g, pDim, startX, startX + lnPx, dimY, $"L_n = {ln:F1}m", Color.FromArgb(255, 150, 50));
                DrawDimLine(g, pDim, lvStart, lvStart + lvPx, dimY + 22, $"L_v = {lv:F1}m", accentYellow);
                DrawDimLine(g, pDim, lxStart, lxStart + lxPx, dimY, $"L_x = {lx:F1}m", Color.FromArgb(255, 80, 80));
                DrawDimLine(g, pDim, startX, startX + totalPx, dimY + 44, $"L = {totalL:F1}m", accentRed);

                // Nhãn
                DrawLabel(g, startX + lnPx / 2, mainY + laneH + sideH / 2, "NÊM", Color.FromArgb(255, 150, 50));
                DrawLabel(g, lvStart + lvPx / 2, mainY + laneH + sideH / 2, "GIẢM TỐC", accentBlue);
                DrawLabel(g, lxStart + lxPx / 2, mainY + laneH + sideH / 2, "ĐỖ CHỜ", Color.FromArgb(255, 80, 80));

                // Mũi tên hướng xe
                DrawArrow(g, startX + totalPx, mainY + laneH / 2, startX + totalPx - 30, mainY + laneH / 2, Color.White);
                DrawArrow(g, lxStart + lxPx, mainY + laneH + sideH / 2, lxStart + lxPx + 25, mainY + laneH + sideH + 15, Color.FromArgb(255, 80, 80));
            }
        }

        private void DrawDimLine(Graphics g, Pen p, float x1, float x2, float y, string text, Color color)
        {
            g.DrawLine(p, x1, y, x2, y);
            g.DrawLine(p, x1, y - 4, x1, y + 4);
            g.DrawLine(p, x2, y - 4, x2, y + 4);
            using (var f = new System.Drawing.Font("Consolas", 8, FontStyle.Bold))
            using (var b = new SolidBrush(color))
            {
                var sz = g.MeasureString(text, f);
                g.DrawString(text, f, b, (x1 + x2) / 2 - sz.Width / 2, y - sz.Height - 2);
            }
        }

        private void DrawLabel(Graphics g, float x, float y, string text, Color color)
        {
            using (var f = new System.Drawing.Font("Segoe UI", 8, FontStyle.Bold))
            using (var b = new SolidBrush(color))
            {
                var sz = g.MeasureString(text, f);
                g.DrawString(text, f, b, x - sz.Width / 2, y - sz.Height / 2);
            }
        }

        private void DrawArrow(Graphics g, float x1, float y1, float x2, float y2, Color color)
        {
            using (var p = new Pen(color, 2f))
            {
                p.EndCap = LineCap.ArrowAnchor;
                g.DrawLine(p, x1, y1, x2, y2);
            }
        }
    }
}
