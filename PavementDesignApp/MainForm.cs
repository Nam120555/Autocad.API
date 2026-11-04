using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.Linq;
using System.Collections.Generic;

namespace PavementDesignApp
{
    public class MainForm : Form
    {
        private Color bgMain = Color.FromArgb(20, 20, 20);
        private Color bgPanel = Color.FromArgb(30, 30, 30);
        private Color textPrimary = Color.FromArgb(240, 240, 240);
        private Color accentBlue = Color.FromArgb(52, 152, 219);
        private Color accentGreen = Color.FromArgb(46, 204, 113);
        private Color accentRed = Color.FromArgb(231, 76, 60);

        private TextBox txtP, txtp, txtD, txtNtk, txtEo, txtCo, txtPhio;
        private DataGridView dgvLayers;
        private Label lblEvCh, lblTauCh, lblKuCh;
        private RichTextBox rtbDetails;
        private Panel pnlCanvas;

        public MainForm()
        {
            InitializeComponent();
            LoadDefaultData();
            
            this.Shown += (s, e) => {
                var lbl = new Label { Text = "DESIGNED BY HUYNH NGOC NAM / TRAFFIC ENGINEER - SĐT: 0378584154", ForeColor = Color.IndianRed, Font = new Font("Segoe UI", 9f, FontStyle.Italic | FontStyle.Bold), AutoSize = true, BackColor = Color.Transparent, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
                Control host = this;
                foreach (Control c in Controls) { if (c.Dock == DockStyle.Fill || c.Dock == DockStyle.Right || c.Dock == DockStyle.Bottom) { if (c is Panel || c is TabControl) host = c; } }
                host.Controls.Add(lbl);
                lbl.Location = new Point(host.ClientSize.Width - lbl.PreferredWidth - 10, host.ClientSize.Height - lbl.PreferredHeight - 8);
                lbl.BringToFront();
                host.Resize += (sender, args) => { lbl.Location = new Point(host.ClientSize.Width - lbl.PreferredWidth - 10, host.ClientSize.Height - lbl.PreferredHeight - 8); lbl.BringToFront(); };
            };
        }

        private void InitializeComponent()
        {
            this.Text = "PREMIUM CIVIL TOOL - THIẾT KẾ KẾT CẤU ÁO ĐƯỜNG MỀM (22TCN 211-06) V2.0";
            this.Size = new Size(1600, 900);
            this.BackColor = bgMain;
            this.ForeColor = textPrimary;
            this.StartPosition = FormStartPosition.CenterScreen;

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 40, BackColor = Color.FromArgb(15, 15, 15) };
            Label lblTitle = new Label { Text = "HỆ THỐNG KIỂM TOÁN CƯỜNG ĐỘ ÁO ĐƯỜNG MỀM VÀ BẢNG TRA ĐIỆN TỬ", ForeColor = Color.Gold, Font = new Font("Segoe UI", 12, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblTitle);
            this.Controls.Add(pnlHeader);

            TabControl tabMain = new TabControl { Dock = DockStyle.Fill, ItemSize = new Size(300, 30), Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            this.Controls.Add(tabMain);

            TabPage tabCalc = new TabPage { Text = "🖥️ CHƯƠNG TRÌNH TÍNH TOÁN ÁO ĐƯỜNG", BackColor = bgMain };
            TabPage tabDict = new TabPage { Text = "📖 TỪ ĐIỂN BẢNG TRA (22TCN 211-06)", BackColor = bgMain };
            tabMain.TabPages.Add(tabCalc);
            tabMain.TabPages.Add(tabDict);

            // ================== TAB 2: DICTIONARY ==================
            DataGridView dgvDict = new DataGridView
            {
                Dock = DockStyle.Fill, BackgroundColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.Black, ReadOnly = true, AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false, Font = new Font("Segoe UI", 10)
            };
            dgvDict.Columns.Add("Cat", "Phân Loại");
            dgvDict.Columns.Add("Name", "Tên Vật Liệu Theo Tiêu Chuẩn 22TCN 211-06");
            dgvDict.Columns.Add("E", "Mô đun ĐH E (MPa)");
            dgvDict.Columns.Add("C", "Lực dính C (MPa)");
            dgvDict.Columns.Add("Phi", "Góc ma sát Phi (độ)");
            dgvDict.Columns.Add("Rku", "Kháng kéo uốn (MPa)");

            foreach (var m in MaterialCatalog.GetAll())
            {
                dgvDict.Rows.Add(m.Category, m.Name, m.E, m.C, m.Phi, m.Rku);
            }
            tabDict.Controls.Add(dgvDict);
            
            Panel pnlDictTop = new Panel { Dock = DockStyle.Top, Height=50, BackColor = Color.DarkSlateBlue };
            pnlDictTop.Controls.Add(new Label { Text = "DỮ LIỆU ĐƯỢC TÍCH HỢP NGUYÊN BẢN TỪ PHỤ LỤC B & C - TCVN 211-06", ForeColor = Color.White, Font = new Font("Segoe UI", 11, FontStyle.Bold), Dock=DockStyle.Fill, TextAlign=ContentAlignment.MiddleCenter});
            tabDict.Controls.Add(pnlDictTop);


            // ================== TAB 1: CALCULATION WORKSPACE ==================
            Panel pnlLeft = new Panel { Dock = DockStyle.Left, Width = 500, BackColor = bgPanel, Padding = new Padding(10) };
            tabCalc.Controls.Add(pnlLeft);

            int y = 10;
            Label lblHeader1 = new Label { Text = "1. TẢI TRỌNG TRỤC XE (P)", ForeColor = accentBlue, Font = new Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(10, y), AutoSize = true };
            pnlLeft.Controls.Add(lblHeader1);
            y += 25;

            txtP = CreateInputBox(pnlLeft, "Tải trọng trục P (kN):", "100", ref y);
            txtp = CreateInputBox(pnlLeft, "Áp lực vệt lốp p (daN/cm2):", "6.0", ref y);
            txtD = CreateInputBox(pnlLeft, "Đường kính tương đương D (cm):", "33.0", ref y);
            txtNtk = CreateInputBox(pnlLeft, "Trục xe quy đổi N_tk (trục/ngày):", "1500", ref y);

            y += 10;
            Label lblHeader2 = new Label { Text = "2. CHIỀU DÀY & BẢNG TRA VẬT LIỆU", ForeColor = accentBlue, Font = new Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(10, y), AutoSize = true };
            pnlLeft.Controls.Add(lblHeader2);
            y += 25;

            // DataGridView cho lớp vật liệu
            dgvLayers = new DataGridView
            {
                Location = new Point(10, y), Width = 475, Height = 250,
                BackgroundColor = Color.FromArgb(40,40,40), ForeColor = Color.Black,
                AllowUserToAddRows = true, AllowUserToDeleteRows = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            
            DataGridViewComboBoxColumn colMat = new DataGridViewComboBoxColumn { Name = "Material", HeaderText = "Tên Vật Liệu (Bảng tra)" };
            foreach (var m in MaterialCatalog.GetAll()) colMat.Items.Add(m.Name);
            
            DataGridViewTextBoxColumn colH = new DataGridViewTextBoxColumn { Name = "H", HeaderText = "H (cm)" };
            DataGridViewTextBoxColumn colE = new DataGridViewTextBoxColumn { Name = "E", HeaderText = "E (MPa)", ReadOnly = true };
            DataGridViewTextBoxColumn colC = new DataGridViewTextBoxColumn { Name = "C", HeaderText = "C (MPa)", ReadOnly = true };
            DataGridViewTextBoxColumn colPhi = new DataGridViewTextBoxColumn { Name = "Phi", HeaderText = "Phi (độ)", ReadOnly = true };
            DataGridViewTextBoxColumn colRku = new DataGridViewTextBoxColumn { Name = "Rku", HeaderText = "Rku", ReadOnly = true };

            dgvLayers.Columns.Add(colMat);
            dgvLayers.Columns.Add(colH);
            dgvLayers.Columns.Add(colE);
            dgvLayers.Columns.Add(colC);
            dgvLayers.Columns.Add(colPhi);
            dgvLayers.Columns.Add(colRku);

            dgvLayers.CellValueChanged += DgvLayers_CellValueChanged;
            dgvLayers.CurrentCellDirtyStateChanged += DgvLayers_CurrentCellDirtyStateChanged;
            dgvLayers.DataError += DgvLayers_DataError;
            pnlLeft.Controls.Add(dgvLayers);
            y += 260;

            Label lblHeader3 = new Label { Text = "3. NỀN ĐẤT TỰ NHIÊN (K95/K98)", ForeColor = accentBlue, Font = new Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(10, y), AutoSize = true };
            pnlLeft.Controls.Add(lblHeader3);
            y += 25;

            txtEo = CreateInputBox(pnlLeft, "Đàn hồi nền E (MPa):", "42", ref y);
            txtCo = CreateInputBox(pnlLeft, "Lực dính C (MPa):", "0.012", ref y);
            txtPhio = CreateInputBox(pnlLeft, "Góc nội ma sát Phi (độ):", "22", ref y);

            Button btnRun = new Button { Text = "BẮT ĐẦU TÍNH TOÁN CƯỜNG ĐỘ", BackColor = accentRed, ForeColor = Color.White, Font = new Font("Segoe UI", 12, FontStyle.Bold), Location = new Point(10, y + 10), Width = 475, Height = 50, FlatStyle = FlatStyle.Flat };
            btnRun.Click += BtnCalculate_Click;
            pnlLeft.Controls.Add(btnRun);

            // Right Side - Detailed Logs
            Panel pnlRight = new Panel { Dock = DockStyle.Right, Width = 450, BackColor = bgPanel, Padding = new Padding(10) };
            tabCalc.Controls.Add(pnlRight);
            
            int ry = 10;
            Label lblHeaderR = new Label { Text = "KẾT QUẢ TỔNG HỢP", ForeColor = accentBlue, Font = new Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(10, ry), AutoSize = true };
            pnlRight.Controls.Add(lblHeaderR);
            ry += 25;

            lblEvCh = CreateResBox(pnlRight, "1. VÕNG ĐÀN HỒI (ECH >= EYC)", ry); ry+=50;
            lblTauCh = CreateResBox(pnlRight, "2. KIỂM TOÁN TRƯỢT (TAU <= CP)", ry); ry+=50;
            lblKuCh = CreateResBox(pnlRight, "3. KÉO UỐN BÊ TÔNG NHỰA", ry); ry+=50;

            ry += 20;
            Label lblHeaderLogs = new Label { Text = "BẢN GHI GIẢI TRÌNH KỸ THUẬT (DETAILED LOGS):", ForeColor = Color.Gold, Font = new Font("Segoe UI", 9, FontStyle.Bold), Location = new Point(10, ry), AutoSize = true };
            pnlRight.Controls.Add(lblHeaderLogs);
            ry += 20;

            rtbDetails = new RichTextBox
            {
                Location = new Point(10, ry), Width = 425, Height = 480,
                BackColor = Color.FromArgb(10,10,10), ForeColor = Color.Lime,
                Font = new Font("Consolas", 10), ReadOnly = true
            };
            pnlRight.Controls.Add(rtbDetails);

            Label lblSign = new Label { Text = "DESIGNED BY HUYNH NGOC NAM\nTRAFFIC ENGINEER - SĐT: 0378584154", ForeColor = accentRed, Font = new Font("Segoe UI", 9, FontStyle.Italic|FontStyle.Bold), Location = new Point(10, ry+ 490), AutoSize = true };
            pnlRight.Controls.Add(lblSign);

            // Center Panel - Canvas
            pnlCanvas = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(25,25,25) };
            pnlCanvas.Paint += PnlCanvas_Paint;
            pnlCanvas.Resize += (s,e)=> pnlCanvas.Invalidate();
            tabCalc.Controls.Add(pnlCanvas);
        }

        private TextBox CreateInputBox(Panel pnl, string text, string def, ref int y)
        {
            pnl.Controls.Add(new Label { Text = text, ForeColor = textPrimary, Font = new Font("Segoe UI", 9), Location = new Point(10, y), AutoSize = true });
            TextBox tb = new TextBox { Text = def, BackColor = Color.FromArgb(50,50,50), ForeColor = Color.White, Location = new Point(250, y-3), Width = 235 };
            pnl.Controls.Add(tb);
            y += 30;
            return tb;
        }

        private Label CreateResBox(Panel pnl, string t, int y)
        {
            pnl.Controls.Add(new Label { Text = t, ForeColor = Color.Silver, Font = new Font("Segoe UI", 8), Location = new Point(10, y), AutoSize = true });
            Label lbl = new Label { Text = "CHƯA KIỂM TRA", ForeColor = Color.Gray, Font = new Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(200, y-2), AutoSize = true };
            pnl.Controls.Add(lbl);
            return lbl;
        }

        private void LoadDefaultData()
        {
            dgvLayers.Rows.Add("BTN chặt hạt nhỏ (15°C)", "5", "3000", "0.35", "40", "4.5");
            dgvLayers.Rows.Add("BTN chặt hạt thô (15°C)", "7", "2200", "0.30", "40", "3.5");
            dgvLayers.Rows.Add("Cấp phối đá dăm (CPĐD) Loại 1", "15", "300", "0.025", "40", "0");
            dgvLayers.Rows.Add("Cấp phối đá dăm (CPĐD) Loại 2", "25", "250", "0.020", "38", "0");
        }

        private void DgvLayers_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            // Bỏ qua lỗi Default Error Dialog nếu có
            e.ThrowException = false;
        }

        private void DgvLayers_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvLayers.IsCurrentCellDirty) dgvLayers.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void DgvLayers_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == 0) // Changed Material
            {
                var val = dgvLayers.Rows[e.RowIndex].Cells[0].Value?.ToString();
                var mat = MaterialCatalog.GetAll().FirstOrDefault(x => x.Name == val);
                if (mat != null)
                {
                    dgvLayers.Rows[e.RowIndex].Cells[2].Value = mat.E;
                    dgvLayers.Rows[e.RowIndex].Cells[3].Value = mat.C;
                    dgvLayers.Rows[e.RowIndex].Cells[4].Value = mat.Phi;
                    dgvLayers.Rows[e.RowIndex].Cells[5].Value = mat.Rku;
                }
            }
            pnlCanvas.Invalidate();
        }

        private void BtnCalculate_Click(object sender, EventArgs e)
        {
            try
            {
                double P = double.Parse(txtP.Text);
                double p_val = double.Parse(txtp.Text);
                double D = double.Parse(txtD.Text);
                double Ntk = double.Parse(txtNtk.Text);
                double E0 = double.Parse(txtEo.Text);
                double C0 = double.Parse(txtCo.Text);
                double Phi0 = double.Parse(txtPhio.Text);

                var layers = new List<PavementLayer>();
                for (int i = 0; i < dgvLayers.Rows.Count; i++)
                {
                    if (dgvLayers.Rows[i].IsNewRow) continue;
                    
                    var row = dgvLayers.Rows[i];
                    string matName = row.Cells[0].Value?.ToString() ?? "";
                    if (!double.TryParse(row.Cells[1].Value?.ToString(), out double h)) continue;
                    double ev = Convert.ToDouble(row.Cells[2].Value);
                    double c = Convert.ToDouble(row.Cells[3].Value);
                    double phi = Convert.ToDouble(row.Cells[4].Value);
                    double rku = Convert.ToDouble(row.Cells[5].Value);
                    
                    var mDef = MaterialCatalog.GetAll().FirstOrDefault(x => x.Name == matName);
                    bool isAsphalt = mDef != null && mDef.IsAsphalt;
                    
                    layers.Add(new PavementLayer(h, ev, c, phi, rku, matName, isAsphalt));
                }

                var res = PavementEngine.Analyze(P, p_val, D, Ntk, E0, C0, Phi0, layers);

                UpdateResultLabel(lblEvCh, res.VongPassed, $"[ {res.Echung:F1} / {res.Eyc:F1} ]");
                UpdateResultLabel(lblTauCh, res.TruotPassed, $"[ Tmx: {res.TauMax:F2} <= Tcp: {res.TauCp:F2} ]");
                UpdateResultLabel(lblKuCh, res.KeoUonPassed, $"[ Sig: {res.SigmaKu:F2} <= Scp: {res.SigmaCp:F2} ]");

                rtbDetails.Text = res.DetailedLog;
                pnlCanvas.Invalidate();
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void UpdateResultLabel(Label lbl, bool passed, string detail)
        {
            if (passed)
            {
                lbl.Text = $"ĐẠT. {detail}";
                lbl.ForeColor = accentGreen;
            }
            else
            {
                lbl.Text = $"KHÔNG ĐẠT. {detail}";
                lbl.ForeColor = accentRed;
            }
        }

        private void PnlCanvas_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            int cw = pnlCanvas.Width;
            int ch = pnlCanvas.Height;
            
            Pen arrowPen = new Pen(Color.Red, 4);
            arrowPen.EndCap = LineCap.ArrowAnchor;
            g.DrawLine(arrowPen, cw/2 - 50, 40, cw/2 - 50, 100);
            g.DrawLine(arrowPen, cw/2 + 50, 40, cw/2 + 50, 100);
            g.DrawString("WHEEL LOAD " + txtP.Text + " kN", new Font("Segoe UI", 10, FontStyle.Bold), Brushes.Red, cw/2 - 70, 15);

            float currentY = 100;
            Color[] colors = { Color.FromArgb(40,40,45), Color.FromArgb(60,60,65), Color.FromArgb(80,80,85), Color.FromArgb(120,110,100), Color.FromArgb(150,140,110), Color.Tan, Color.Gray };
            
            for (int i=0; i<dgvLayers.Rows.Count; i++)
            {
                if (dgvLayers.Rows[i].IsNewRow) continue;
                if (float.TryParse(dgvLayers.Rows[i].Cells[1].Value?.ToString(), out float hj))
                {
                    float layerH = hj * 4; 
                    RectangleF rect = new RectangleF(50, currentY, cw - 100, layerH);
                    g.FillRectangle(new SolidBrush(colors[i%colors.Length]), rect);
                    g.DrawRectangle(Pens.Gray, rect.X, rect.Y, rect.Width, rect.Height);

                    string mat = dgvLayers.Rows[i].Cells[0].Value?.ToString() ?? "Unknown";
                    string ev = dgvLayers.Rows[i].Cells[2].Value?.ToString() ?? "0";
                    g.DrawString($"{mat} | H={hj}cm | E={ev}", new Font("Consolas", 10), Brushes.White, cw/2 - 100, currentY + layerH/2 - 8);
                    
                    g.DrawLine(Pens.Yellow, cw/2 - 50 - currentY*0.3f, currentY, cw/2 - 50 - (currentY+layerH)*0.3f, currentY+layerH);
                    g.DrawLine(Pens.Yellow, cw/2 + 50 + currentY*0.3f, currentY, cw/2 + 50 + (currentY+layerH)*0.3f, currentY+layerH);

                    currentY += layerH;
                }
            }
            
            RectangleF gRect = new RectangleF(50, currentY, cw - 100, ch - currentY - 20);
            g.FillRectangle(new HatchBrush(HatchStyle.DiagonalBrick, Color.DarkGoldenrod, Color.FromArgb(80,60,20)), gRect);
            g.DrawString($"NỀN K95/K98 | Eo = {txtEo.Text} MPa", new Font("Consolas", 12, FontStyle.Bold), Brushes.Gold, cw/2 - 120, currentY + 30);
        }
    }
}
