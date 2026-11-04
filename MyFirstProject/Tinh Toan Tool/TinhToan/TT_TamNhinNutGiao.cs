using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(TinhToan.TT_TamNhinNutGiaoCommands))]

namespace TinhToan
{
    public class TT_TamNhinNutGiaoCommands
    {
        [CommandMethod("TT_TamNhinNutGiao")]
        public static void ShowTamNhinNutGiao()
        {
            if (TT_TamNhinNutGiaoForm.Instance == null || TT_TamNhinNutGiaoForm.Instance.IsDisposed)
                TT_TamNhinNutGiaoForm.Instance = new TT_TamNhinNutGiaoForm();
            AcadApp.ShowModelessDialog(TT_TamNhinNutGiaoForm.Instance);
        }
    }

    public class TT_TamNhinNutGiaoForm : Form
    {
        public static TT_TamNhinNutGiaoForm? Instance;
        private Color bgMain = Color.FromArgb(30, 30, 30);
        private Color bgSidebar = Color.FromArgb(38, 38, 42);
        private Color bgInput = Color.FromArgb(51, 51, 53);
        private Color textPrimary = Color.FromArgb(240, 240, 240);
        private Color accentGreen = Color.FromArgb(10, 240, 120);
        private Color accentBlue = Color.FromArgb(0, 191, 255);
        private Color accentYellow = Color.FromArgb(255, 190, 0);
        private Color accentRed = Color.FromArgb(255, 60, 90);

        private TextBox txtVa = null!, txtVb = null!, txtTpsd = null!, txtFriction = null!;
        private Label lblSa = null!, lblSb = null!, lblResult = null!;
        private DoubleBufferedPanel pnlCanvas = null!;

        public static double CurrentSa = 50, CurrentSb = 40;

        public TT_TamNhinNutGiaoForm()
        {
            this.Text = "KIỂM TRA TẦM NHÌN NÚT GIAO - CIVIL TOOL";
            this.Size = new Size(1100, 750);
            this.BackColor = bgMain; this.ForeColor = textPrimary;
            this.StartPosition = FormStartPosition.CenterScreen;

            Panel pnlLeft = new Panel { Dock = DockStyle.Left, Width = 310, BackColor = bgSidebar, AutoScroll = true };
            this.Controls.Add(pnlLeft);
            pnlLeft.Controls.Add(new Label { Text = "■ THÔNG SỐ TẦM NHÌN (TCVN 4054)", ForeColor = accentBlue, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, 15), AutoSize = true });

            int yPos = 45;
            txtVa = CreateInput(pnlLeft, "Vận tốc tuyến chính (V_a) [km/h]", "50", ref yPos);
            txtVb = CreateInput(pnlLeft, "Vận tốc tuyến phụ (V_b) [km/h]", "40", ref yPos);
            txtTpsd = CreateInput(pnlLeft, "Thời gian PƯ tài xế (t) [s]", "2.5", ref yPos);
            txtFriction = CreateInput(pnlLeft, "Hệ số bám dọc (φ)", "0.5", ref yPos);

            Panel pnlRight = new Panel { Dock = DockStyle.Right, Width = 310, BackColor = bgSidebar };
            this.Controls.Add(pnlRight);
            pnlRight.Controls.Add(new Label { Text = "■ KẾT QUẢ TẦM NHÌN", ForeColor = accentRed, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, 15), AutoSize = true });

            int ry = 45;
            lblSa = CreateResult(pnlRight, "Tầm nhìn tuyến chính (S_a)", "50.00 m", accentGreen, ref ry);
            lblSb = CreateResult(pnlRight, "Tầm nhìn tuyến phụ (S_b)", "40.00 m", accentYellow, ref ry);

            ry += 10;
            pnlRight.Controls.Add(new Label { Text = "TAM GIÁC TẦM NHÌN", ForeColor = Color.White, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, ry), AutoSize = true });
            lblResult = new Label { Text = "S_a × S_b", ForeColor = accentRed, Font = new System.Drawing.Font("Consolas", 18, FontStyle.Bold), Location = new Point(15, ry + 25), AutoSize = true };
            pnlRight.Controls.Add(lblResult);

            // CENTER CANVAS 2D
            Panel pnlCenter = new Panel { Dock = DockStyle.Fill };
            this.Controls.Add(pnlCenter);
            var lblTitle = new Label { Text = "TAM GIÁC TẦM NHÌN NÚT GIAO (2D)", ForeColor = Color.White, Font = new System.Drawing.Font("Segoe UI", 12, FontStyle.Bold), Dock = DockStyle.Top, TextAlign = ContentAlignment.MiddleCenter, Height = 40, AutoSize = false };
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
                double t = double.Parse(txtTpsd.Text), phi = double.Parse(txtFriction.Text);

                // Tầm nhìn hãm xe: S = V*t/3.6 + V²/(254*φ)
                double sa = (va * t / 3.6) + Math.Pow(va, 2) / (254.0 * phi);
                double sb = (vb * t / 3.6) + Math.Pow(vb, 2) / (254.0 * phi);

                CurrentSa = sa; CurrentSb = sb;
                lblSa.Text = $"{sa:F2} m"; lblSb.Text = $"{sb:F2} m";
                lblResult.Text = $"{sa:F1} × {sb:F1} m";
                pnlCanvas?.Invalidate();
            }
            catch { }
        }

        private void PnlCanvas_Paint(object? sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            float w = pnlCanvas.Width, h = pnlCanvas.Height;
            double sa = CurrentSa, sb = CurrentSb;
            if (sa < 1 || sb < 1) return;

            float cx = w / 2f, cy = h / 2f;
            double maxS = Math.Max(sa, sb) + 20;
            float scale = (float)(Math.Min(cx, cy) * 0.7 / maxS);
            float roadW = 30;

            g.TranslateTransform(cx, cy);

            using (var bRoad = new SolidBrush(Color.FromArgb(50, 55, 60)))
            using (var bGrass = new SolidBrush(Color.FromArgb(35, 50, 35)))
            using (var bTriangle = new SolidBrush(Color.FromArgb(50, 255, 60, 60)))
            using (var pEdge = new Pen(Color.Cyan, 2f))
            using (var pTim = new Pen(Color.Yellow, 1.5f) { DashPattern = new float[] { 10, 5, 2, 5 } })
            using (var pTriangle = new Pen(Color.FromArgb(255, 60, 60), 2.5f))
            using (var pDim = new Pen(Color.FromArgb(200, 200, 200), 1f) { DashPattern = new float[] { 4, 3 } })
            {
                float ext = (float)(maxS * scale + 50);

                // Nền cỏ
                g.FillRectangle(bGrass, -ext, -ext, ext * 2, ext * 2);

                // Đường chính (ngang) - Tuyến A
                g.FillRectangle(bRoad, -ext, -roadW / 2, ext * 2, roadW);
                g.DrawLine(pEdge, -ext, -roadW / 2, ext, -roadW / 2);
                g.DrawLine(pEdge, -ext, roadW / 2, ext, roadW / 2);
                g.DrawLine(pTim, -ext, 0, ext, 0);

                // Đường phụ (dọc) - Tuyến B
                g.FillRectangle(bRoad, -roadW / 2, -ext, roadW, ext * 2);
                g.DrawLine(pEdge, -roadW / 2, -ext, -roadW / 2, ext);
                g.DrawLine(pEdge, roadW / 2, -ext, roadW / 2, ext);
                g.DrawLine(pTim, 0, -ext, 0, ext);

                // Tam giác tầm nhìn (góc phần tư phía dưới-phải)
                float saPx = (float)(sa * scale);
                float sbPx = (float)(sb * scale);

                // Điểm giao (gốc), điểm A (trên tuyến chính), điểm B (trên tuyến phụ)
                PointF ptOrigin = new PointF(roadW / 4, roadW / 4);
                PointF ptA = new PointF(saPx, roadW / 4);
                PointF ptB = new PointF(roadW / 4, sbPx);

                PointF[] triangle = { ptOrigin, ptA, ptB };
                g.FillPolygon(bTriangle, triangle);
                g.DrawPolygon(pTriangle, triangle);

                // Dimension S_a (ngang)
                float dimYa = sbPx + 25;
                g.DrawLine(pDim, ptOrigin.X, dimYa, ptA.X, dimYa);
                g.DrawLine(pDim, ptOrigin.X, dimYa - 5, ptOrigin.X, dimYa + 5);
                g.DrawLine(pDim, ptA.X, dimYa - 5, ptA.X, dimYa + 5);
                DrawDimText(g, (ptOrigin.X + ptA.X) / 2, dimYa + 5, $"S_a = {sa:F1}m", accentGreen);

                // Dimension S_b (dọc)
                float dimXb = saPx + 25;
                g.DrawLine(pDim, dimXb, ptOrigin.Y, dimXb, ptB.Y);
                g.DrawLine(pDim, dimXb - 5, ptOrigin.Y, dimXb + 5, ptOrigin.Y);
                g.DrawLine(pDim, dimXb - 5, ptB.Y, dimXb + 5, ptB.Y);
                DrawDimText(g, dimXb + 5, (ptOrigin.Y + ptB.Y) / 2, $"S_b = {sb:F1}m", accentYellow);

                // Nhãn xe
                DrawLabel(g, saPx - 20, -5, "XE A →", Color.White);
                DrawLabel(g, 15, sbPx - 15, "XE B ↓", Color.White);

                // Nhãn tam giác
                DrawLabel(g, saPx / 3, sbPx / 3, "VÙNG\nTẦM NHÌN", Color.FromArgb(255, 100, 100));
            }

            g.ResetTransform();
        }

        private void DrawDimText(Graphics g, float x, float y, string text, Color color)
        {
            using (var f = new System.Drawing.Font("Consolas", 9, FontStyle.Bold))
            using (var b = new SolidBrush(color))
                g.DrawString(text, f, b, x, y);
        }

        private void DrawLabel(Graphics g, float x, float y, string text, Color color)
        {
            using (var f = new System.Drawing.Font("Segoe UI", 9, FontStyle.Bold))
            using (var b = new SolidBrush(color))
                g.DrawString(text, f, b, x, y);
        }
    }
}
