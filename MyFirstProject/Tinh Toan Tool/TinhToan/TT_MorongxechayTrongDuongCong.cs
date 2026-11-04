using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(TinhToan.TT_MoRongXeChayCommands))]

namespace TinhToan
{
    public class TT_MoRongXeChayCommands
    {
        [CommandMethod("TT_MoRongXeChay")]
        public static void ShowMoRongXeChay()
        {
            if (TT_MoRongXeChayForm.Instance == null || TT_MoRongXeChayForm.Instance.IsDisposed)
                TT_MoRongXeChayForm.Instance = new TT_MoRongXeChayForm();
            AcadApp.ShowModelessDialog(TT_MoRongXeChayForm.Instance);
        }
    }

    public class TT_MoRongXeChayForm : Form
    {
        public static TT_MoRongXeChayForm? Instance;
        private Color bgMain = Color.FromArgb(30, 30, 30);
        private Color bgSidebar = Color.FromArgb(38, 38, 42);
        private Color bgInput = Color.FromArgb(51, 51, 53);
        private Color textPrimary = Color.FromArgb(240, 240, 240);
        private Color accentGreen = Color.FromArgb(10, 240, 120);
        private Color accentBlue = Color.FromArgb(0, 191, 255);
        private Color accentRed = Color.FromArgb(255, 60, 90);
        private Color accentYellow = Color.FromArgb(255, 190, 0);

        private TextBox txtVtk = null!, txtHeso = null!, txtMu = null!, txtInut = null!;
        private TextBox txtRxe = null!, txtL = null!, txtB1 = null!;
        private Label lblVnutVal = null!, lblRminVal = null!, lblEVal = null!;
        private DoubleBufferedPanel pnlCanvas = null!;

        public static double CurrentWidening = 4.64;
        public static double CurrentB1 = 2.75;
        public static double CurrentRxe = 17.25;
        private double currentVnut = 20, currentRmin = 15;

        public TT_MoRongXeChayForm()
        {
            InitializeComponent();
            CalculateMathCAD();
            MyFirstProject.FormSignatureHelper.AddSignature(this);
        }

        private void InitializeComponent()
        {
            this.Text = "TÍNH TOÁN MỞ RỘNG XE CHẠY TRONG ĐƯỜNG CONG - CIVIL TOOL";
            this.Size = new Size(1000, 750);
            this.BackColor = bgMain; this.ForeColor = textPrimary;
            this.StartPosition = FormStartPosition.CenterScreen;

            Panel pnlLeft = new Panel { Dock = DockStyle.Left, Width = 280, BackColor = bgSidebar, AutoScroll = true };
            this.Controls.Add(pnlLeft);
            pnlLeft.Controls.Add(new Label { Text = "■ DỮ LIỆU TÍNH TOÁN", ForeColor = accentBlue, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, 15), AutoSize = true });

            int yPos = 45;
            txtVtk = CreateInput(pnlLeft, "Tốc độ TK (V_tk) [km/h]", "50", ref yPos);
            txtHeso = CreateInput(pnlLeft, "Hệ số giảm tốc trong nút", "0.40", ref yPos);
            txtMu = CreateInput(pnlLeft, "Hệ số lực ngang (μ)", "0.20", ref yPos);
            txtInut = CreateInput(pnlLeft, "Dốc ngang nút (i_nut)", "0.01", ref yPos);
            txtRxe = CreateInput(pnlLeft, "Bán kính TK (R_xe) [m]", "17.25", ref yPos);
            txtL = CreateInput(pnlLeft, "Khoảng cách 2 trục xe (L) [m]", "12.8", ref yPos);
            txtB1 = CreateInput(pnlLeft, "Bề rộng 1 làn (b_1) [m]", "2.75", ref yPos);

            Panel pnlRight = new Panel { Dock = DockStyle.Right, Width = 280, BackColor = bgSidebar };
            this.Controls.Add(pnlRight);
            pnlRight.Controls.Add(new Label { Text = "■ KẾT QUẢ THIẾT KẾ", ForeColor = accentRed, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, 15), AutoSize = true });

            pnlRight.Controls.Add(new Label { Text = "e = L²/(2R+b₁) + 0.05V/√R", ForeColor = accentBlue, Font = new System.Drawing.Font("Consolas", 8), Location = new Point(15, 45), AutoSize = true });

            int ry = 80;
            lblVnutVal = CreateResult(pnlRight, "Vận tốc trong nút (V_nut)", "20.00 km/h", accentGreen, ref ry);
            lblRminVal = CreateResult(pnlRight, "Bán kính xe tối thiểu (R_min)", "15.00 m", accentRed, ref ry);

            ry += 10;
            pnlRight.Controls.Add(new Label { Text = "ĐỘ MỞ RỘNG (e_req)", ForeColor = textPrimary, Font = new System.Drawing.Font("Segoe UI", 9), Location = new Point(15, ry), AutoSize = true });
            lblEVal = new Label { Text = "4.64 m", ForeColor = accentGreen, Font = new System.Drawing.Font("Consolas", 26, FontStyle.Bold), Location = new Point(15, ry + 20), AutoSize = true };
            pnlRight.Controls.Add(lblEVal);

            // CENTER CANVAS 2D
            Panel pnlCenter = new Panel { Dock = DockStyle.Fill };
            this.Controls.Add(pnlCenter);
            var lblTitle = new Label { Text = "MẶT CẮT NGANG MỞ RỘNG ĐƯỜNG CONG (2D)", ForeColor = Color.White, Font = new System.Drawing.Font("Segoe UI", 12, FontStyle.Bold), Dock = DockStyle.Top, TextAlign = ContentAlignment.MiddleCenter, Height = 40, AutoSize = false };
            pnlCenter.Controls.Add(lblTitle);

            pnlCanvas = new DoubleBufferedPanel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(25, 28, 35) };
            pnlCanvas.Paint += PnlCanvas_Paint;
            pnlCenter.Controls.Add(pnlCanvas);
            lblTitle.BringToFront();
        }

        private TextBox CreateInput(Panel parent, string label, string defaultValue, ref int yPos)
        {
            parent.Controls.Add(new Label { Text = label, ForeColor = textPrimary, Font = new System.Drawing.Font("Segoe UI", 9), Location = new Point(15, yPos), AutoSize = true });
            var txt = new TextBox { Text = defaultValue, BackColor = bgInput, ForeColor = accentBlue, Font = new System.Drawing.Font("Consolas", 10, FontStyle.Bold), Location = new Point(15, yPos + 20), Width = 230, BorderStyle = BorderStyle.FixedSingle };
            txt.TextChanged += (s, e) => CalculateMathCAD();
            parent.Controls.Add(txt);
            yPos += 55;
            return txt;
        }

        private Label CreateResult(Panel parent, string label, string defaultVal, Color valColor, ref int yPos)
        {
            parent.Controls.Add(new Label { Text = label, ForeColor = textPrimary, Font = new System.Drawing.Font("Segoe UI", 9), Location = new Point(15, yPos), AutoSize = true });
            var val = new Label { Text = defaultVal, ForeColor = valColor, Font = new System.Drawing.Font("Consolas", 12, FontStyle.Bold), Location = new Point(15, yPos + 20), AutoSize = true };
            parent.Controls.Add(val);
            yPos += 60;
            return val;
        }

        private void CalculateMathCAD()
        {
            try
            {
                double v_tk = double.Parse(txtVtk.Text), hs = double.Parse(txtHeso.Text);
                double mu = double.Parse(txtMu.Text), i_nut = double.Parse(txtInut.Text);
                double r_xe = double.Parse(txtRxe.Text), L_truc = double.Parse(txtL.Text);
                double b_1 = double.Parse(txtB1.Text);

                double v_nut = hs * v_tk;
                double r_min = (mu + i_nut) != 0 ? Math.Pow(v_nut, 2) / (127 * (mu + i_nut)) : 0;
                double e_req = r_xe > 0 ? (Math.Pow(L_truc, 2) / (2 * r_xe + b_1)) + ((0.05 * v_nut) / Math.Sqrt(r_xe)) : 0;

                CurrentWidening = Math.Round(e_req, 2); CurrentB1 = b_1; CurrentRxe = r_xe;
                currentVnut = v_nut; currentRmin = r_min;
                lblVnutVal.Text = $"{v_nut:F2} km/h"; lblRminVal.Text = $"{r_min:F2} m"; lblEVal.Text = $"{e_req:F2} m";
                pnlCanvas?.Invalidate();
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
        }

        private void PnlCanvas_Paint(object? sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            float w = pnlCanvas.Width, h = pnlCanvas.Height;
            double r_xe = CurrentRxe, b1 = CurrentB1, eReq = CurrentWidening;
            if (r_xe < 1) return;

            float cx = w / 2f, cy = h / 2f;

            // Vẽ mặt bằng đường cong với mở rộng
            double rIn = r_xe;
            double rOutNormal = r_xe + b1 * 2;     // 2 làn bình thường
            double rOutWidened = r_xe + b1 * 2 + eReq; // 2 làn + mở rộng

            double maxR = rOutWidened + 15;
            float scale = (float)(Math.Min(cx, cy) * 0.8 / maxR);

            // Đặt tâm đường cong ở phía dưới bên phải canvas
            float ocx = w * 0.7f;
            float ocy = h * 0.85f;

            g.TranslateTransform(ocx, ocy);

            using (var bRoad = new SolidBrush(Color.FromArgb(50, 55, 60)))
            using (var bWidened = new SolidBrush(Color.FromArgb(60, 80, 70)))
            using (var pIn = new Pen(Color.Cyan, 2f))
            using (var pOut = new Pen(Color.Cyan, 2f))
            using (var pWidened = new Pen(accentGreen, 2.5f))
            using (var pTim = new Pen(Color.Yellow, 1.5f) { DashPattern = new float[] { 8, 4, 2, 4 } })
            using (var pDim = new Pen(Color.FromArgb(200, 200, 200), 1f) { DashPattern = new float[] { 4, 3 } })
            {
                float rInPx = (float)(rIn * scale);
                float rOutNPx = (float)(rOutNormal * scale);
                float rOutWPx = (float)(rOutWidened * scale);
                float rTimPx = (rInPx + rOutNPx) / 2f;

                // Vẽ phần đường cong (góc phần tư 2 - 180° đến 270°)
                float startAngle = 180;
                float sweepAngle = 90;

                // Đường cong mở rộng (fill)
                using (var path = new GraphicsPath())
                {
                    path.AddArc(-rOutWPx, -rOutWPx, rOutWPx * 2, rOutWPx * 2, startAngle, sweepAngle);
                    path.AddArc(-rInPx, -rInPx, rInPx * 2, rInPx * 2, startAngle + sweepAngle, -sweepAngle);
                    path.CloseFigure();
                    g.FillPath(bRoad, path);
                }

                // Mép trong
                g.DrawArc(pIn, -rInPx, -rInPx, rInPx * 2, rInPx * 2, startAngle, sweepAngle);
                // Mép ngoài bình thường
                g.DrawArc(pOut, -rOutNPx, -rOutNPx, rOutNPx * 2, rOutNPx * 2, startAngle, sweepAngle);
                // Mép ngoài mở rộng
                g.DrawArc(pWidened, -rOutWPx, -rOutWPx, rOutWPx * 2, rOutWPx * 2, startAngle, sweepAngle);
                // Tim đường
                g.DrawArc(pTim, -rTimPx, -rTimPx, rTimPx * 2, rTimPx * 2, startAngle, sweepAngle);

                // Đường thẳng nối (213° - giữa cung)
                double midAngle = 225 * Math.PI / 180.0;
                float mInX = rInPx * (float)Math.Cos(midAngle);
                float mInY = rInPx * (float)Math.Sin(midAngle);
                float mOutNX = rOutNPx * (float)Math.Cos(midAngle);
                float mOutNY = rOutNPx * (float)Math.Sin(midAngle);
                float mOutWX = rOutWPx * (float)Math.Cos(midAngle);
                float mOutWY = rOutWPx * (float)Math.Sin(midAngle);

                // Dim: bề rộng bình thường
                g.DrawLine(pDim, mInX, mInY, mOutNX, mOutNY);
                DrawDimTextRotated(g, (mInX + mOutNX) / 2 + 5, (mInY + mOutNY) / 2, $"b = {b1 * 2:F1}m", accentBlue);

                // Dim: phần mở rộng
                g.DrawLine(pDim, mOutNX, mOutNY, mOutWX, mOutWY);
                DrawDimTextRotated(g, (mOutNX + mOutWX) / 2 + 5, (mOutNY + mOutWY) / 2, $"e = {eReq:F2}m", accentGreen);

                // Bán kính R_xe
                float rXeEndX = rInPx * (float)Math.Cos(200 * Math.PI / 180);
                float rXeEndY = rInPx * (float)Math.Sin(200 * Math.PI / 180);
                g.DrawLine(pDim, 0, 0, rXeEndX, rXeEndY);
                DrawDimTextRotated(g, rXeEndX / 2 + 10, rXeEndY / 2 - 10, $"R = {r_xe:F1}m", accentYellow);

                // Tâm O
                g.FillEllipse(Brushes.White, -3, -3, 6, 6);
                g.DrawString("O", new System.Drawing.Font("Consolas", 10, FontStyle.Bold), Brushes.White, 5, 2);
            }

            g.ResetTransform();

            // Info
            using (var f = new System.Drawing.Font("Consolas", 10, FontStyle.Bold))
            {
                g.DrawString($"R = {r_xe:F1}m  |  b₁ = {b1:F2}m  |  e = {eReq:F2}m", f, Brushes.Cyan, 15, h - 30);
            }
        }

        private void DrawDimTextRotated(Graphics g, float x, float y, string text, Color color)
        {
            using (var f = new System.Drawing.Font("Consolas", 9, FontStyle.Bold))
            using (var b = new SolidBrush(color))
                g.DrawString(text, f, b, x, y);
        }
    }
}
