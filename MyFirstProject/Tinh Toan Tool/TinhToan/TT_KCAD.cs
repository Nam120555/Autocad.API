using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;

[assembly: CommandClass(typeof(TinhToan.TT_KetCauAoDuongCommands))]

namespace TinhToan
{
    public class TT_KetCauAoDuongCommands
    {
        [CommandMethod("TT_AODUONG")]
        public static void ShowTool()
        {
            if (TT_KetCauAoDuongForm.Instance == null || TT_KetCauAoDuongForm.Instance.IsDisposed)
            {
                TT_KetCauAoDuongForm.Instance = new TT_KetCauAoDuongForm();
            }
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(TT_KetCauAoDuongForm.Instance);
        }
    }

    public class TT_KetCauAoDuongForm : Form
    {
        public static TT_KetCauAoDuongForm? Instance;

        private Color bgMain = Color.FromArgb(20, 20, 20);
        private Color bgPanel = Color.FromArgb(30, 30, 30);
        private Color textPrimary = Color.FromArgb(240, 240, 240);
        private Color accentBlue = Color.FromArgb(52, 152, 219);
        private Color accentGreen = Color.FromArgb(46, 204, 113);
        private Color accentRed = Color.FromArgb(231, 76, 60);

        private Panel pnlCanvas = null!;
        private TextBox txtEo = null!, txtP = null!, txtp = null!, txtD = null!, txtNtk = null!;
        private TextBox[] txtH = new TextBox[4];
        private TextBox[] txtE = new TextBox[4];

        private Label lblEvCh = null!, lblTauCh = null!, lblKuCh = null!;

        public TT_KetCauAoDuongForm()
        {
            InitializeComponent();
            MyFirstProject.FormSignatureHelper.AddSignature(this);
        }

        private void InitializeComponent()
        {
            this.Text = "CIVIL TOOL - THIẾT KẾ KẾT CẤU ÁO ĐƯỜNG MỀM (22TCN 211-06)";
            this.Size = new Size(1300, 800);
            this.BackColor = bgMain;
            this.ForeColor = textPrimary;
            this.StartPosition = FormStartPosition.CenterScreen;

            // Header
            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 50, BackColor = Color.FromArgb(15, 15, 15) };
            Label lblTitle = new Label { Text = "BÀI TOÁN KIỂM TOÁN CƯỜNG ĐỘ KẾT CẤU ÁO ĐƯỜNG (MÔ PHỎNG EXCEL)", ForeColor = Color.Gold, Font = new Font("Segoe UI", 12, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblTitle);
            this.Controls.Add(pnlHeader);

            // Left Panel (Tải trọng & Lớp mặt)
            Panel pnlLeft = new Panel { Dock = DockStyle.Left, Width = 380, BackColor = bgPanel, AutoScroll = true, Padding = new Padding(10) };
            this.Controls.Add(pnlLeft);

            int y = 15;
            CreateHeader(pnlLeft, "1. THÔNG SỐ TẢI TRỌNG TRỤC", ref y);
            txtP = CreateInputRow(pnlLeft, "Tải trọng trục tiêu chuẩn P (kN):", "100", ref y);
            txtp = CreateInputRow(pnlLeft, "Áp lực vệt bánh xe p (daN/cm2):", "6.0", ref y);
            txtD = CreateInputRow(pnlLeft, "Đường kính vệt TD D (cm):", "33.0", ref y);
            txtNtk = CreateInputRow(pnlLeft, "Trục xe quy đổi N_tk (trục/làn/ngày):", "1200", ref y);

            y += 20;
            CreateHeader(pnlLeft, "2. CHIỀU DÀY & VẬT LIỆU", ref y);
            
            // 4 Layers
            for (int i = 0; i < 4; i++)
            {
                Label lbl = new Label { Text = $"Lớp {i + 1}:", ForeColor = accentBlue, Font = new Font("Segoe UI", 9, FontStyle.Bold), Location = new Point(10, y + 5), AutoSize = true };
                pnlLeft.Controls.Add(lbl);

                txtH[i] = new TextBox { Text = (i==0)?"5":(i==1)?"7":(i==2)?"15":"25", Width = 50, Location = new Point(80, y), BackColor = Color.FromArgb(40,40,40), ForeColor = Color.White };
                pnlLeft.Controls.Add(new Label { Text="H(cm)", Location=new Point(80, y-15), ForeColor=Color.Gray, Font = new Font("Segoe UI", 7) });
                pnlLeft.Controls.Add(txtH[i]);

                txtE[i] = new TextBox { Text = (i==0)?"4000":(i==1)?"3000":(i==2)?"300":"250", Width = 60, Location = new Point(140, y), BackColor = Color.FromArgb(40,40,40), ForeColor = Color.White };
                pnlLeft.Controls.Add(new Label { Text="E(MPa)", Location=new Point(140, y-15), ForeColor=Color.Gray, Font = new Font("Segoe UI", 7) });
                pnlLeft.Controls.Add(txtE[i]);

                TextBox txtMat = new TextBox { Text = (i==0)?"BTN Hạt mịn":(i==1)?"BTN Hạt thô":(i==2)?"CPĐD Loại 1":"CPĐD Loại 2", Width = 140, Location = new Point(210, y), BackColor = Color.FromArgb(40,40,40), ForeColor = Color.Gray, ReadOnly=true };
                pnlLeft.Controls.Add(txtMat);

                txtH[i].TextChanged += (s, e) => { pnlCanvas.Invalidate(); };
                y += 40;
            }

            y += 20;
            CreateHeader(pnlLeft, "3. CẤP PHỐI NỀN ĐẤT", ref y);
            txtEo = CreateInputRow(pnlLeft, "Mô đun đàn hồi nền đất E_o (MPa):", "42", ref y);

            Button btnSimulate = new Button { Text = "BẮT ĐẦU TÍNH TOÁN CƯỜNG ĐỘ (ENGINE)", BackColor=Color.DarkRed, ForeColor=Color.White, Font = new Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(10, y+20), Width=340, Height=45, FlatStyle = FlatStyle.Flat };
            btnSimulate.Click += BtnCalculate_Click;
            pnlLeft.Controls.Add(btnSimulate);

            // Right Panel (Results)
            Panel pnlRight = new Panel { Dock = DockStyle.Right, Width = 350, BackColor = bgPanel, Padding = new Padding(10) };
            this.Controls.Add(pnlRight);
            
            y = 15;
            CreateHeader(pnlRight, "BÁO CÁO KIỂM TOÁN TỪ ENGINE (22TCN 211-06)", ref y);

            y += 20;
            lblEvCh = CreateResultBox(pnlRight, "TIÊU CHUẨN VÕNG ĐÀN HỒI", "ĐẠT YÊU CẦU", "K_vdh = 1.35 > 1.0", accentGreen, ref y);
            lblTauCh = CreateResultBox(pnlRight, "TIÊU CHUẨN TRƯỢT NỀN ĐẤT", "ĐẠT YÊU CẦU", "T_max < T_cp", accentGreen, ref y);
            lblKuCh = CreateResultBox(pnlRight, "TIÊU CHUẨN KÉO UỐN (BTN)", "KHÔNG ĐẠT", "Sigma_ku > Sigma_cp (Tăng bề dày)", accentRed, ref y);

            PictureBox picExcel = new PictureBox { Width=310, Height=200, Location=new Point(15, y+20), BorderStyle=BorderStyle.FixedSingle, BackColor=Color.Black };
            pnlRight.Controls.Add(picExcel);



            // Center Panel (Drawing)
            pnlCanvas = new Panel { Dock = DockStyle.Fill };
            pnlCanvas.Paint += PnlCanvas_Paint;
            pnlCanvas.Resize += (s, e) => pnlCanvas.Invalidate();
            this.Controls.Add(pnlCanvas);
        }

        private void CreateHeader(Panel pnl, string title, ref int y)
        {
            Label lbl = new Label { Text = title, ForeColor = accentBlue, Font = new Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(10, y), AutoSize = true };
            pnl.Controls.Add(lbl);
            y += 30;
        }

        private void BtnCalculate_Click(object? sender, EventArgs e)
        {
            try
            {
                double P = double.Parse(txtP.Text);
                double p_val = double.Parse(txtp.Text);
                double D = double.Parse(txtD.Text);
                double Ntk = double.Parse(txtNtk.Text);
                double E0 = double.Parse(txtEo.Text);

                var layers = new System.Collections.Generic.List<PavementLayer>();
                for (int i = 0; i < 4; i++)
                {
                    double.TryParse(txtH[i].Text, out double h);
                    double.TryParse(txtE[i].Text, out double ev);
                    bool isAsphalt = (i == 0 || i == 1);
                    layers.Add(new PavementLayer(h, ev, "Layer " + i, isAsphalt));
                }

                var res = TT_KetCauAoDuongEngine.Analyze(P, p_val, D, Ntk, E0, layers);

                UpdateResultBox(lblEvCh, res.VongPassed, $"E_ch = {res.Echung:F1} MPa");
                UpdateResultBox(lblTauCh, res.TruotPassed, $"Tau_max = {res.TauMax:F3} daN/cm2");
                UpdateResultBox(lblKuCh, res.KeoUonPassed, $"Sigma_ku = {res.SigmaKu:F3} MPa");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Lỗi nhập liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateResultBox(Label lblStatus, bool passed, string valText)
        {
            if (passed)
            {
                lblStatus.Text = $"ĐẠT YÊU CẦU\n({valText})";
                lblStatus.ForeColor = accentGreen;
            }
            else
            {
                lblStatus.Text = $"KHÔNG ĐẠT\n({valText})";
                lblStatus.ForeColor = accentRed;
            }
        }

        private TextBox CreateInputRow(Panel pnl, string title, string defVal, ref int y)
        {
            Label lbl = new Label { Text = title, ForeColor = textPrimary, Font = new Font("Segoe UI", 9), Location = new Point(10, y), AutoSize = true };
            pnl.Controls.Add(lbl);
            
            TextBox txt = new TextBox { Text = defVal, BackColor = Color.FromArgb(40,40,40), ForeColor = Color.White, Font = new Font("Consolas", 10), Location = new Point(240, y - 3), Width = 110, TextAlign=HorizontalAlignment.Right };
            pnl.Controls.Add(txt);
            y += 35;
            return txt;
        }

        private Label CreateResultBox(Panel pnl, string title, string status, string desc, Color c, ref int y)
        {
            Label lblTitle = new Label { Text = title, ForeColor = Color.Gray, Font = new Font("Segoe UI", 8, FontStyle.Bold), Location = new Point(15, y), AutoSize = true };
            pnl.Controls.Add(lblTitle);

            Panel box = new Panel { Location = new Point(15, y+15), Width=310, Height=40, BackColor = Color.FromArgb(20,20,20), BorderStyle=BorderStyle.FixedSingle };
            Label lblStatus = new Label { Text = status, ForeColor = c, Font = new Font("Segoe UI", 11, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter };
            box.Controls.Add(lblStatus);
            pnl.Controls.Add(box);

            Label lblDesc = new Label { Text = desc, ForeColor = c, Font = new Font("Consolas", 9), Location = new Point(15, y+60), AutoSize = true };
            pnl.Controls.Add(lblDesc);

            y += 100;
            return lblStatus;
        }

        private void PnlCanvas_Paint(object? sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            int cw = pnlCanvas.Width;
            int ch = pnlCanvas.Height;
            
            // Draw Axle Force
            Pen arrowPen = new Pen(Color.Red, 4);
            arrowPen.EndCap = LineCap.ArrowAnchor;
            g.DrawLine(arrowPen, cw/2 - 50, 40, cw/2 - 50, 100);
            g.DrawLine(arrowPen, cw/2 + 50, 40, cw/2 + 50, 100);
            g.DrawString("TẢI TRỌNG TRỤC P = " + txtP.Text + " kN", new Font("Segoe UI", 10, FontStyle.Bold), Brushes.Red, cw/2 - 100, 15);

            float currentY = 100;
            Color[] colors = { Color.FromArgb(40,40,45), Color.FromArgb(60,60,65), Color.FromArgb(120,110,100), Color.FromArgb(150,140,110) };
            
            for (int i=0; i<4; i++)
            {
                if (float.TryParse(txtH[i].Text, out float hj))
                {
                    float layerH = hj * 4; // Scale
                    RectangleF rect = new RectangleF(50, currentY, cw - 100, layerH);
                    g.FillRectangle(new SolidBrush(colors[i]), rect);
                    g.DrawRectangle(Pens.Gray, rect.X, rect.Y, rect.Width, rect.Height);

                    g.DrawString($"H{i+1}={hj} cm | E={txtE[i].Text} MPa", new Font("Consolas", 10), Brushes.White, cw/2 - 80, currentY + layerH/2 - 8);
                    
                    // Stress distribution cone (approx)
                    g.DrawLine(Pens.Yellow, cw/2 - 50 - currentY*0.3f, currentY, cw/2 - 50 - (currentY+layerH)*0.3f, currentY+layerH);
                    g.DrawLine(Pens.Yellow, cw/2 + 50 + currentY*0.3f, currentY, cw/2 + 50 + (currentY+layerH)*0.3f, currentY+layerH);

                    currentY += layerH;
                }
            }
            
            // Ground
            RectangleF gRect = new RectangleF(50, currentY, cw - 100, ch - currentY - 20);
            g.FillRectangle(new HatchBrush(HatchStyle.DiagonalBrick, Color.DarkGoldenrod, Color.FromArgb(80,60,20)), gRect);
            g.DrawString($"NỀN K98 | Eo = {txtEo.Text} MPa", new Font("Consolas", 12, FontStyle.Bold), Brushes.Gold, cw/2 - 100, currentY + 30);
        }
    }
}
