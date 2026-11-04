using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(TinhToan.TT_NhapDongCommands))]

namespace TinhToan
{
    public class TT_NhapDongCommands
    {
        [CommandMethod("TT_NhapDong")]
        public static void ShowNhapDong()
        {
            if (TT_NhapDongForm.Instance == null || TT_NhapDongForm.Instance.IsDisposed)
                TT_NhapDongForm.Instance = new TT_NhapDongForm();
            AcadApp.ShowModelessDialog(TT_NhapDongForm.Instance);
        }
    }

    public class TT_NhapDongForm : Form
    {
        public static TT_NhapDongForm? Instance;
        private Color bgMain = Color.FromArgb(30, 30, 30);
        private Color bgSidebar = Color.FromArgb(38, 38, 42);
        private Color bgInput = Color.FromArgb(51, 51, 53);
        private Color textPrimary = Color.FromArgb(240, 240, 240);
        private Color accentGreen = Color.FromArgb(10, 240, 120);
        private Color accentBlue = Color.FromArgb(0, 191, 255);
        private Color accentYellow = Color.FromArgb(255, 190, 0);
        private Color accentRed = Color.FromArgb(255, 60, 90);

        private TextBox txtVa = null!, txtVb = null!, txtD = null!, txtA = null!, txtLc = null!;
        private Label lblLn = null!, lblLg_calc = null!, lblLg_select = null!, lblL = null!;
        private DoubleBufferedPanel pnlCanvas = null!;

        public static double CurrentLn = 0, CurrentLg = 0, CurrentL = 0;

        public TT_NhapDongForm()
        {
            this.Text = "THIẾT KẾ ĐOẠN NHẬP DÒNG (TĂNG TỐC) - CIVIL TOOL";
            this.Size = new Size(1100, 750);
            this.BackColor = bgMain; this.ForeColor = textPrimary;
            this.StartPosition = FormStartPosition.CenterScreen;

            Panel pnlLeft = new Panel { Dock = DockStyle.Left, Width = 310, BackColor = bgSidebar, AutoScroll = true };
            this.Controls.Add(pnlLeft);
            pnlLeft.Controls.Add(new Label { Text = "■ THÔNG SỐ NHẬP DÒNG", ForeColor = accentBlue, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, 15), AutoSize = true });

            int yPos = 45;
            txtVa = CreateInput(pnlLeft, "Vận tốc tuyến chính (V_a) [km/h]", "50", ref yPos);
            txtVb = CreateInput(pnlLeft, "Vận tốc đường nhánh (V_b) [km/h]", "40", ref yPos);
            txtD = CreateInput(pnlLeft, "Bề rộng làn (d) [m]", "3.5", ref yPos);
            txtA = CreateInput(pnlLeft, "Gia tốc tăng tốc (a) [m/s²]", "1.0", ref yPos);
            txtLc = CreateInput(pnlLeft, "Chiều dài tăng tốc tối thiểu (L_c) [m]", "120.0", ref yPos);

            Panel pnlRight = new Panel { Dock = DockStyle.Right, Width = 310, BackColor = bgSidebar };
            this.Controls.Add(pnlRight);
            pnlRight.Controls.Add(new Label { Text = "■ KẾT QUẢ TÍNH TOÁN", ForeColor = accentRed, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, 15), AutoSize = true });

            int ry = 45;
            lblLn = CreateResult(pnlRight, "Chiều dài hình nêm (L_n)", "29.17 m", accentGreen, ref ry);
            lblLg_calc = CreateResult(pnlRight, "L.tăng tốc lý thuyết (L_g')", "34.62 m", accentGreen, ref ry);
            lblLg_select = CreateResult(pnlRight, "L.tăng tốc chọn (L_v)", "120.0 m", accentYellow, ref ry);

            ry += 20;
            pnlRight.Controls.Add(new Label { Text = "TỔNG CHIỀU DÀI NHẬP DÒNG (L)", ForeColor = Color.White, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, ry), AutoSize = true });
            lblL = new Label { Text = "149.17 m", ForeColor = accentRed, Font = new System.Drawing.Font("Consolas", 24, FontStyle.Bold), Location = new Point(15, ry + 25), AutoSize = true };
            pnlRight.Controls.Add(lblL);

            // CENTER CANVAS 2D
            Panel pnlCenter = new Panel { Dock = DockStyle.Fill };
            this.Controls.Add(pnlCenter);
            var lblTitle = new Label { Text = "MẶT BẰNG ĐOẠN NHẬP DÒNG (2D)", ForeColor = Color.White, Font = new System.Drawing.Font("Segoe UI", 12, FontStyle.Bold), Dock = DockStyle.Top, TextAlign = ContentAlignment.MiddleCenter, Height = 40, AutoSize = false };
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
            pnl.Controls.Add(txt); y += 55; return txt;
        }

        private Label CreateResult(Panel pnl, string label, string def, Color clr, ref int y)
        {
            pnl.Controls.Add(new Label { Text = label, ForeColor = textPrimary, Font = new System.Drawing.Font("Segoe UI", 9), Location = new Point(15, y), AutoSize = true });
            var lbl = new Label { Text = def, ForeColor = clr, Font = new System.Drawing.Font("Consolas", 14, FontStyle.Bold), Location = new Point(15, y + 20), AutoSize = true };
            pnl.Controls.Add(lbl); y += 65; return lbl;
        }

        private void Calculate()
        {
            try
            {
                double va = double.Parse(txtVa.Text), vb = double.Parse(txtVb.Text);
                double d = double.Parse(txtD.Text), a = double.Parse(txtA.Text);
                double lc = double.Parse(txtLc.Text);

                double ln = va * d / 6.0;
                double lg_calc = a != 0 ? Math.Abs(Math.Pow(va, 2) - Math.Pow(vb, 2)) / (26.0 * a) : 0;
                double lv = Math.Max(lc, lg_calc);
                double l = ln + lv;

                CurrentLn = ln; CurrentLg = lv; CurrentL = l;
                lblLn.Text = $"{ln:F2} m"; lblLg_calc.Text = $"{lg_calc:F2} m"; lblLg_select.Text = $"{lv:F2} m"; lblL.Text = $"{l:F2} m";
                pnlCanvas?.Invalidate();
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
        }

        private void PnlCanvas_Paint(object? sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            float w = pnlCanvas.Width, h = pnlCanvas.Height;
            double ln = CurrentLn, lv = CurrentLg, totalL = CurrentL;
            if (totalL < 1) return;

            float margin = 80;
            float drawW = w - margin * 2;
            float scale = (float)(drawW / (totalL + 40));
            float mainY = h * 0.4f;
            float laneH = 35;
            float sideH = 28;
            float startX = margin;

            using (var bRoad = new SolidBrush(Color.FromArgb(50, 55, 60)))
            using (var bSide = new SolidBrush(Color.FromArgb(55, 70, 80)))
            using (var pEdge = new Pen(Color.Cyan, 2f))
            using (var pTim = new Pen(Color.Yellow, 1.5f) { DashPattern = new float[] { 10, 5, 2, 5 } })
            using (var pDim = new Pen(Color.FromArgb(200, 200, 200), 1f))
            using (var pNem = new Pen(Color.FromArgb(255, 150, 50), 2f))
            {
                float totalPx = (float)(totalL * scale);
                float lnPx = (float)(ln * scale);
                float lvPx = (float)(lv * scale);

                // Đường chính
                g.FillRectangle(bRoad, startX - 20, mainY, totalPx + 40, laneH);
                g.DrawLine(pEdge, startX - 20, mainY, startX + totalPx + 20, mainY);
                g.DrawLine(pEdge, startX - 20, mainY + laneH, startX + totalPx + 20, mainY + laneH);
                g.DrawLine(pTim, startX - 20, mainY + laneH / 2, startX + totalPx + 20, mainY + laneH / 2);

                // Nhập dòng = tăng tốc: đoạn song song trước → nêm sau
                g.FillRectangle(bSide, startX, mainY + laneH, lvPx, sideH);
                g.DrawRectangle(pEdge, startX, mainY + laneH, lvPx, sideH);

                PointF[] nem = { new PointF(startX + lvPx, mainY + laneH + sideH), new PointF(startX + lvPx, mainY + laneH), new PointF(startX + lvPx + lnPx, mainY + laneH) };
                g.FillPolygon(bSide, nem);
                g.DrawPolygon(pNem, nem);

                // Đường nhánh nhập vào (cong nhẹ)
                using (var pBranch = new Pen(Color.FromArgb(180, 180, 180), 1.5f) { DashPattern = new float[] { 5, 3 } })
                {
                    g.DrawLine(pBranch, startX - 30, mainY + laneH + sideH + 40, startX, mainY + laneH + sideH);
                }

                DrawLabel(g, startX + lvPx / 2, mainY + laneH + sideH / 2, "TĂNG TỐC", accentGreen);
                DrawLabel(g, startX + lvPx + lnPx / 2, mainY + laneH + sideH / 3, "NÊM", Color.FromArgb(255, 150, 50));

                // Mũi tên nhập
                using (var pArr = new Pen(Color.White, 2f) { EndCap = LineCap.ArrowAnchor })
                    g.DrawLine(pArr, startX - 15, mainY + laneH + sideH / 2, startX + 20, mainY + laneH + sideH / 2);

                // Dimensions
                float dimY = mainY + laneH + sideH + 25;
                DrawDimLine(g, pDim, startX, startX + lvPx, dimY, $"L_v = {lv:F1}m", accentYellow);
                DrawDimLine(g, pDim, startX + lvPx, startX + totalPx, dimY, $"L_n = {ln:F1}m", Color.FromArgb(255, 150, 50));
                DrawDimLine(g, pDim, startX, startX + totalPx, dimY + 22, $"L = {totalL:F1}m", accentRed);
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
    }
}
