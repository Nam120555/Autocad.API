using System;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using MyFirstProject.XuatPhanMemTinhToan;

[assembly: CommandClass(typeof(TinhToan.TT_TinhToanBanKinhReCommands))]

namespace TinhToan
{
    public class TT_TinhToanBanKinhReCommands
    {
        [CommandMethod("TT_TinhToanBanKinhRe")]
        public static void ShowBanKinhRePhai()
        {
            if (TT_TinhToanBanKinhReForm.Instance == null || TT_TinhToanBanKinhReForm.Instance.IsDisposed)
            {
                TT_TinhToanBanKinhReForm.Instance = new TT_TinhToanBanKinhReForm();
            }
            AcadApp.ShowModelessDialog(TT_TinhToanBanKinhReForm.Instance);
        }


    }

    public class TT_TinhToanBanKinhReForm : Form
    {
        public static TT_TinhToanBanKinhReForm? Instance;

        private Color bgMain = Color.FromArgb(30, 30, 30);       
        private Color bgSidebar = Color.FromArgb(38, 38, 42);    
        private Color bgInput = Color.FromArgb(51, 51, 53);      
        
        private Color textPrimary = Color.FromArgb(240, 240, 240);
        private Color accentGreen = Color.FromArgb(10, 240, 120);
        private Color accentBlue = Color.FromArgb(0, 191, 255);
        private Color accentRed = Color.FromArgb(255, 60, 90);
        private Color accentYellow = Color.FromArgb(255, 190, 0);

        private TextBox txtVtk = null!;
        private TextBox txtHeso = null!;
        private TextBox txtMu = null!;
        private TextBox txtInut = null!;

        private Label lblVnutVal = null!;
        private Label lblRminVal = null!;
        private Label lblRchonVal = null!;

        private TextBox txtBeRongMat = null!;
        private TextBox txtBeRongViaHe = null!;
        private TextBox txtBeRongDPC = null!;

        private TextBox txtProjectName = null!;
        private TextBox txtLocation = null!;
        private TextBox txtItem = null!;

        private DoubleBufferedPanel pnlCanvas2D = null!;
        
        public static double CurrentRadius = 50.0;
        
        public static Color PassedSysColor = Color.Lime;   
        public static Color FailedSysColor = Color.Red;   
        public static string ResultLayerName = "TT_KetQua_Bankinh";
        public static double CircleDrawRadius = 5.0; 
        public static double CircleGlobalWidth = 0.5;
        
        public static string TextFont = "Arial";
        public static double TextHeight = 2.5;

        public TT_TinhToanBanKinhReForm()
        {
            InitializeComponent();
            CalculateMathCAD();
            this.MouseWheel += Form_MouseWheel;
            MyFirstProject.FormSignatureHelper.AddSignature(this);
        }

        private void Form_MouseWheel(object? sender, MouseEventArgs e) { }

        private void InitializeComponent()
        {
            this.Text = "TÍNH TOÁN THIẾT KẾ ĐƯỜNG CONG RẼ PHẢI - CIVIL TOOL";
            this.Size = new Size(950, 750); 
            this.BackColor = bgMain;
            this.ForeColor = textPrimary;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = true;


            Panel pnlLeft = new Panel { Dock = DockStyle.Left, Width = 280, BackColor = bgSidebar, AutoScroll = true };
            this.Controls.Add(pnlLeft);

            Label lblHeaderProj = new Label { Text = "■ THÔNG TIN DỰ ÁN", ForeColor = accentYellow, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, 15), AutoSize = true };
            pnlLeft.Controls.Add(lblHeaderProj);

            int yPos = 45;
            txtProjectName = CreateInputGroup(pnlLeft, "Tên dự án", "", ref yPos);
            txtLocation = CreateInputGroup(pnlLeft, "Địa chỉ", "", ref yPos);
            txtItem = CreateInputGroup(pnlLeft, "Hạng mục", "Tính toán thiết kế đường cong rẽ phải", ref yPos);

            yPos += 15;
            Label lblHeaderL = new Label { Text = "■ DỮ LIỆU TÍNH TOÁN", ForeColor = accentBlue, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, yPos), AutoSize = true };
            pnlLeft.Controls.Add(lblHeaderL);
            yPos += 30;

            txtVtk = CreateInputGroup(pnlLeft, "Vận tốc thiết kế gốc (V_tk) [km/h]", "80", ref yPos);
            txtHeso = CreateInputGroup(pnlLeft, "Hệ số giảm tốc trong nút", "0.45", ref yPos);
            txtMu = CreateInputGroup(pnlLeft, "Hệ số lực ngang (μ)", "0.2", ref yPos);
            txtInut = CreateInputGroup(pnlLeft, "Dốc ngang mặt đường trong nút (i_nut)", "0.015", ref yPos);

            yPos += 15;
            Label lblHeaderScale = new Label { Text = "■ THÔNG SỐ ĐỂ VẼ MẶT BẰNG 2D", ForeColor = accentGreen, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, yPos), AutoSize = true };
            pnlLeft.Controls.Add(lblHeaderScale);
            yPos += 30;

            txtBeRongMat = CreateInputGroup(pnlLeft, "Bề rộng mặt đường 1 bên (m)", "7.0", ref yPos);
            txtBeRongViaHe = CreateInputGroup(pnlLeft, "Bề rộng vỉa hè (m)", "5.0", ref yPos);
            txtBeRongDPC = CreateInputGroup(pnlLeft, "Bề rộng nửa dải phân cách (m)", "1.5", ref yPos);

            Panel pnlRight = new Panel { Dock = DockStyle.Right, Width = 280, BackColor = bgSidebar };
            this.Controls.Add(pnlRight);

            Label lblHeaderR = new Label { Text = "■ KẾT QUẢ THIẾT KẾ", ForeColor = accentRed, Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(15, 15), AutoSize = true };
            pnlRight.Controls.Add(lblHeaderR);

            Label lblFormula = new Label { Text = "Công thức Tính Tốc Độ Vằn:\nV_nut = Hệ số × V_tk\nR_min = V_nut² / [127(μ + i_nut)]", ForeColor = accentBlue, Font = new System.Drawing.Font("Consolas", 8), Location = new Point(15, 35), Size = new Size(250, 50) };
            pnlRight.Controls.Add(lblFormula);

            int rightYPos = 90;
            lblVnutVal = CreateResultGroup(pnlRight, "Vận tốc xe chạy rẽ phải (V_nut)", "36.00", "km/h", accentGreen, ref rightYPos);
            lblRminVal = CreateResultGroup(pnlRight, "Bán kính cầu tính R_min", "47.464", "m", accentYellow, ref rightYPos);
            lblRchonVal = CreateResultGroup(pnlRight, "BÁN KÍNH LỰA CHỌN (R_chon)", "50.00", "m", accentRed, ref rightYPos);
            rightYPos += 30;

            Button btnExportExcel = new Button {
                Text = "XUẤT EXCEL THUYẾT MINH",
                ForeColor = Color.White,
                BackColor = accentBlue,
                Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold),
                Location = new Point(15, rightYPos),
                Size = new Size(250, 45),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnExportExcel.FlatAppearance.BorderSize = 0;
            btnExportExcel.Click += (s, e) => ExportReport(false);
            pnlRight.Controls.Add(btnExportExcel);

            Button btnExportPDF = new Button {
                Text = "XUẤT BÁO CÁO (PDF)",
                ForeColor = Color.White,
                BackColor = accentRed,
                Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold),
                Location = new Point(15, rightYPos + 55),
                Size = new Size(250, 45),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnExportPDF.FlatAppearance.BorderSize = 0;
            btnExportPDF.Click += (s, e) => ExportReport(true);
            pnlRight.Controls.Add(btnExportPDF);


            Panel pnlCenter = new Panel { Dock = DockStyle.Fill };
            this.Controls.Add(pnlCenter);

            Label lblTitle = new Label { 
                Text = "THIẾT KẾ BÁN KÍNH XE RẼ PHẢI (3D)", 
                ForeColor = Color.White, 
                Font = new System.Drawing.Font("Segoe UI", 14, FontStyle.Bold), 
                Dock = DockStyle.Top,
                TextAlign = ContentAlignment.MiddleCenter,
                Height = 60,
                AutoSize = false
            };
            pnlCenter.Controls.Add(lblTitle);

            pnlCanvas2D = new DoubleBufferedPanel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(20, 25, 30) };
            pnlCanvas2D.Paint += PnlCanvas2D_Paint;
            pnlCanvas2D.Resize += (s, e) => pnlCanvas2D.Invalidate();
            pnlCenter.Controls.Add(pnlCanvas2D);
            
            lblTitle.BringToFront(); 

            CalculateMathCAD(); // Initialize 3D scene on load
        }

        private TextBox CreateInputGroup(Panel parent, string label, string defaultValue, ref int yPos)
        {
            Label lbl = new Label { Text = label, ForeColor = textPrimary, Font = new System.Drawing.Font("Segoe UI", 9), Location = new Point(15, yPos), AutoSize = true };
            TextBox txt = new TextBox { Text = defaultValue, BackColor = bgInput, ForeColor = accentBlue, Font = new System.Drawing.Font("Consolas", 10, FontStyle.Bold), Location = new Point(15, yPos + 20), Width = 230, BorderStyle = BorderStyle.FixedSingle };
            txt.TextChanged += (s, e) => CalculateMathCAD();
            parent.Controls.Add(lbl);
            parent.Controls.Add(txt);
            yPos += 55; 
            return txt;
        }

        private Button CreateColorGroup(Panel parent, string label, Color defaultColor, ref int yPos)
        {
            Label lbl = new Label { Text = label, ForeColor = textPrimary, Font = new System.Drawing.Font("Segoe UI", 9), Location = new Point(15, yPos), AutoSize = true };
            Button btn = new Button { BackgroundImage = null, BackColor = defaultColor, Location = new Point(15, yPos + 20), Width = 230, Height = 25, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.BorderColor = Color.White;
            btn.Click += (s, e) => {
                using (ColorDialog cd = new ColorDialog()) {
                    cd.Color = btn.BackColor;
                    if (cd.ShowDialog() == DialogResult.OK) {
                        btn.BackColor = cd.Color;
                        CalculateMathCAD(); 
                    }
                }
            };
            parent.Controls.Add(lbl);
            parent.Controls.Add(btn);
            yPos += 55;
            return btn;
        }

        private Label CreateResultGroup(Panel parent, string label, string defaultVal, string unit, Color valColor, ref int yPos)
        {
            Label lbl = new Label { Text = label, ForeColor = textPrimary, Font = new System.Drawing.Font("Segoe UI", 9), Location = new Point(15, yPos), AutoSize = true };
            Label val = new Label { Text = defaultVal + " " + unit, ForeColor = valColor, Font = new System.Drawing.Font("Segoe UI", 12, FontStyle.Bold), Location = new Point(15, yPos + 20), AutoSize = true };
            parent.Controls.Add(lbl);
            parent.Controls.Add(val);
            yPos += 60;
            return val;
        }


        private void CalculateMathCAD()
        {
            try
            {
                var culture = System.Globalization.CultureInfo.InvariantCulture;
                double v_tk = double.Parse(txtVtk.Text.Replace(",", "."), culture);
                double hs = double.Parse(txtHeso.Text.Replace(",", "."), culture);
                double mu = double.Parse(txtMu.Text.Replace(",", "."), culture);
                double i_nut = double.Parse(txtInut.Text.Replace(",", "."), culture);

                double v_nut = hs * v_tk;
                double r_min = 0;
                
                if (mu + i_nut != 0)
                {
                    r_min = Math.Pow(v_nut, 2) / (127 * (mu + i_nut));
                }

                double r_chon = Math.Ceiling(r_min / 5.0) * 5.0;
                if (r_chon == 0) r_chon = 50;
                
                CurrentRadius = r_chon;

                lblVnutVal.Text = $"{v_nut:F2} km/h";
                lblRminVal.Text = $"{r_min:F3} m";
                lblRchonVal.Text = $"{r_chon:F2} m";

                if (pnlCanvas2D != null)
                {
                    pnlCanvas2D.Invalidate();
                }
            }
            catch { }
        }

        private void ExportReport(bool isPdf)
        {
            try
            {
                var culture = System.Globalization.CultureInfo.InvariantCulture;
                double v_tk = double.Parse(txtVtk.Text.Replace(",", "."), culture);
                double heso = double.Parse(txtHeso.Text.Replace(",", "."), culture);
                double v_nut = v_tk * heso;
                double mu = double.Parse(txtMu.Text.Replace(",", "."), culture);
                double i_nut = double.Parse(txtInut.Text.Replace(",", "."), culture);
                double r_min = Math.Pow(v_nut, 2) / (127.0 * (mu + i_nut));
                double r_chon = Math.Ceiling(r_min / 5.0) * 5.0;
                
                var listData = new System.Collections.Generic.List<MathCADReportGroup>();

                listData.Add(new MathCADReportGroup { GroupName = "1. SỐ LIỆU TÍNH TOÁN:" });

                var group11 = new MathCADReportGroup { GroupName = "1.1. Tốc độ xe chạy rẽ phải:" };
                group11.Parameters.Add(new MathCADReportModel($"với vận tốc thiết kế: V_tk = {v_tk} (km/h)", "V_nut", $"{heso} × V_tk = {heso} × {v_tk}", $"{v_nut:F0}", "(km/h)"));
                listData.Add(group11);

                var group12 = new MathCADReportGroup { GroupName = "1.2. Bán kính tối thiểu của dòng xe chạy trong nút:" };
                group12.Parameters.Add(new MathCADReportModel("", "R_xe", $"V_nut² / [127 × (μ + i_nut)] = {v_nut:F0}² / [127 × ({mu} + {i_nut})]", $"{r_min:F3}", "(m)"));
                group12.Parameters.Add(new MathCADReportModel("=> Chọn bán kính mép đường xe chạy", "R_xe", "-", $"{r_chon:F0}", "(m)"));
                group12.Parameters.Add(new MathCADReportModel("trong đó: Hệ số lực ngang", "μ", "-", $"{mu}", ""));
                group12.Parameters.Add(new MathCADReportModel("          Dốc ngang mặt đường trong nút", "i_nut", "-", $"{i_nut}", ""));
                listData.Add(group12);

                string projName = txtProjectName.Text.Trim();
                string loc = txtLocation.Text.Trim();
                string itm = txtItem.Text.Trim();

                ExcelExportEngine.ExportMathCADReport("THIẾT KẾ BÁN KÍNH XE RẼ PHẢI", listData, isPdf, projName, loc, itm);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Có lỗi khi tạo báo cáo: " + ex.Message, "Lỗi");
            }
        }

        private void PnlCanvas2D_Paint(object? sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            float cx = pnlCanvas2D.Width / 2f;
            float cy = pnlCanvas2D.Height / 2f;

            double hsRate = 8.0; 
            try { hsRate = Math.Max(5, Math.Min(cx, cy) / 60.0); } catch { }

            float r_chon = (float)(CurrentRadius * hsRate);
            float wRoad = 0; float wDpc = 0; float wVh = 0;

            try
            {
                var culture = System.Globalization.CultureInfo.InvariantCulture;
                wRoad = (float)(double.Parse(txtBeRongMat.Text.Replace(",", "."), culture) * hsRate);
                wVh = (float)(double.Parse(txtBeRongViaHe.Text.Replace(",", "."), culture) * hsRate);
                wDpc = (float)(double.Parse(txtBeRongDPC.Text.Replace(",", "."), culture) * hsRate);
            }
            catch { wRoad = 35f; wVh = 25f; wDpc = 10f; }

            g.TranslateTransform(cx, cy);

            using (SolidBrush bAsphalt = new SolidBrush(Color.FromArgb(45, 50, 55)))
            using (SolidBrush bWalk = new SolidBrush(Color.FromArgb(60, 65, 70)))
            using (SolidBrush bDpc = new SolidBrush(Color.FromArgb(80, 100, 80)))
            using (SolidBrush bDpcFill = new SolidBrush(Color.FromArgb(60, 120, 60)))
            using (SolidBrush bIsland = new SolidBrush(Color.FromArgb(70, 110, 70)))
            using (Pen pTim = new Pen(Color.Yellow, 1.5f) { DashPattern = new float[] { 10, 5, 2, 5 } })
            using (Pen pKerb = new Pen(Color.Cyan, 2f))
            using (Pen pDpc = new Pen(Color.Lime, 2f))
            using (Pen pIsland = new Pen(Color.FromArgb(100, 200, 100), 1.5f))
            {
                // Draw 4 quadrants
                for (int i = 0; i < 4; i++)
                {
                    g.RotateTransform(90);

                    // Toa do goc mep duong
                    float cornerX = wDpc + wRoad;
                    float cornerY = wDpc + wRoad;

                    // Ve mat duong (Asphalt)
                    g.FillRectangle(bAsphalt, wDpc, wDpc, cornerX + r_chon + 500, wRoad);
                    g.FillRectangle(bAsphalt, wDpc, wDpc, wRoad, cornerY + r_chon + 500);

                    // Bo cong nga tu
                    g.FillPie(bAsphalt, cornerX, cornerY, r_chon * 2, r_chon * 2, 180, 90);

                    // Me duong thang
                    g.DrawLine(pKerb, cornerX, cornerY + r_chon, cornerX, cy + cx);
                    g.DrawLine(pKerb, cornerX + r_chon, cornerY, cx + cy, cornerY);

                    // Bo via r_chon
                    g.DrawArc(pKerb, cornerX, cornerY, r_chon * 2, r_chon * 2, 180, 90);

                    // Ve via he bo cong song song voi bo via
                    float innerSize = (r_chon - wVh) * 2;
                    if (innerSize > 0)
                    {
                        using(Pen pVh = new Pen(Color.Silver, 1f))
                        {
                            g.DrawArc(pVh, cornerX + wVh, cornerY + wVh, innerSize, innerSize, 180, 90);
                            g.DrawLine(pVh, cornerX + wVh, cornerY + r_chon, cornerX + wVh, cy + cx);
                            g.DrawLine(pVh, cornerX + r_chon, cornerY + wVh, cx + cy, cornerY + wVh);
                        }
                    }

                    // ══════════════════════════════════════════════════
                    // DẢI PHÂN CÁCH (DPC) - MŨI TRÒN
                    // ══════════════════════════════════════════════════
                    float dpcTipY = cornerY + r_chon + 50; // Vị trí mũi DPC dọc
                    float dpcTipX = cornerX + r_chon + 50; // Vị trí mũi DPC ngang

                    // Đường viền DPC dọc (2 cạnh song song)
                    g.DrawLine(pDpc, 0, dpcTipY, 0, cy + cx);
                    g.DrawLine(pDpc, wDpc * 2, dpcTipY, wDpc * 2, cy + cx);
                    // Fill thân DPC dọc
                    g.FillRectangle(bDpcFill, 0, dpcTipY, wDpc * 2, cy + cx - dpcTipY);

                    // Mũi tròn DPC dọc (bán nguyệt hướng xuống)
                    g.FillEllipse(bDpcFill, 0, dpcTipY - wDpc, wDpc * 2, wDpc * 2);
                    g.DrawArc(pDpc, 0, dpcTipY - wDpc, wDpc * 2, wDpc * 2, 0, 180);

                    // Đường viền DPC ngang (2 cạnh song song)
                    g.DrawLine(pDpc, dpcTipX, 0, cx + cy, 0);
                    g.DrawLine(pDpc, dpcTipX, wDpc * 2, cx + cy, wDpc * 2);
                    // Fill thân DPC ngang
                    g.FillRectangle(bDpcFill, dpcTipX, 0, cx + cy - dpcTipX, wDpc * 2);

                    // Mũi tròn DPC ngang (bán nguyệt hướng trái)
                    g.FillEllipse(bDpcFill, dpcTipX - wDpc, 0, wDpc * 2, wDpc * 2);
                    g.DrawArc(pDpc, dpcTipX - wDpc, 0, wDpc * 2, wDpc * 2, 270, 180);

                    // ══════════════════════════════════════════════════
                    // ĐẢO GIỌT NƯỚC (Channelizing Teardrop Island)
                    // Nằm bên trong góc, giữa đường thẳng và đường cong rẽ
                    // ══════════════════════════════════════════════════
                    float islandOffset = wVh * 0.3f;  // Khoảng cách từ mép vỉa hè
                    float islandRcurve = r_chon - wVh - islandOffset; // Bán kính cong đảo (theo bờ vỉa trong)

                    if (islandRcurve > 10)
                    {
                        // Đảo giọt nước gồm 3 phần:
                        // 1. Cung tròn song song với bo cong (mép trong)
                        // 2. Đoạn thẳng dọc (song song mép đường dọc)
                        // 3. Đoạn thẳng ngang (song song mép đường ngang)
                        // 4. Mũi tròn ở 2 đầu

                        float islandInner = cornerX + wVh + islandOffset;
                        float islandArcSize = islandRcurve * 2;
                        
                        // Chiều dài cánh đảo
                        float wingLen = Math.Min(r_chon * 0.6f, 80f);
                        
                        // Điểm bắt đầu cánh dọc (trên cung)
                        float arcEndY = islandInner + islandRcurve;
                        float arcEndX = islandInner + islandRcurve;
                        
                        // Vẽ đường Path đảo giọt nước
                        using (var islandPath = new System.Drawing.Drawing2D.GraphicsPath())
                        {
                            // Cạnh dọc bên trái (từ mũi trên đi xuống tới cung)
                            float leftX = islandInner;
                            float topMuiY = islandInner; // Mũi trên
                            float bottomArcY = arcEndY;  // Nơi cung bắt đầu

                            // Cạnh ngang bên trên (từ cung đi sang phải tới mũi)
                            float topY = islandInner;
                            float rightMuiX = arcEndX + wingLen;

                            // Xây dựng path đảo giọt nước
                            // Điểm 1: Mũi dọc (trên) - bắt đầu bằng nửa hình tròn nhỏ
                            float muiR = Math.Min(wVh * 0.4f, 8f);

                            // Cánh dọc: đi từ trên xuống
                            islandPath.AddLine(leftX, topMuiY + wingLen, leftX, bottomArcY);

                            // Cung tròn: từ dọc sang ngang (bán kính islandRcurve)
                            islandPath.AddArc(islandInner, islandInner, islandArcSize, islandArcSize, 180, 90);

                            // Cánh ngang: đi sang phải 
                            islandPath.AddLine(arcEndX, topY, arcEndX + wingLen, topY);

                            // Mũi ngang (đầu nhọn giọt nước) - bo tròn
                            islandPath.AddArc(arcEndX + wingLen - muiR, topY - muiR, muiR * 2, muiR * 2, 270, 180);

                            // Quay về: cánh ngang (mép dưới)
                            float islandWidth = Math.Max(wVh * 0.35f, 6f);
                            islandPath.AddLine(arcEndX + wingLen, topY + islandWidth, arcEndX, topY + islandWidth);

                            // Cung tròn mép trong (nhỏ hơn)
                            float innerArcSize = (islandRcurve - islandWidth) * 2;
                            if (innerArcSize > 2)
                            {
                                islandPath.AddArc(islandInner + islandWidth, islandInner + islandWidth,
                                    innerArcSize, innerArcSize, 270, -90);
                            }

                            // Cánh dọc mép phải (đi lên)
                            islandPath.AddLine(leftX + islandWidth, bottomArcY, leftX + islandWidth, topMuiY + wingLen);

                            // Mũi dọc (trên) - bo tròn
                            islandPath.AddArc(leftX - muiR, topMuiY + wingLen - muiR, muiR * 2 + islandWidth, muiR * 2, 0, -180);

                            islandPath.CloseFigure();

                            // Fill đảo
                            g.FillPath(bIsland, islandPath);
                            // Viền đảo
                            g.DrawPath(pIsland, islandPath);
                        }
                    }
                }

                // Ve tim duong 
                g.DrawLine(pTim, -cx - cy, 0, cx + cy, 0);
                g.DrawLine(pTim, 0, -cx - cy, 0, cx + cy);
            }
            g.ResetTransform();

            // Text hien thi
            g.DrawString($"R_chon = {CurrentRadius} m", new System.Drawing.Font("Consolas", 14, FontStyle.Bold), Brushes.Cyan, 20, 20);
            g.DrawString("MẶT BẰNG ĐƯỜNG CONG RẼ PHẢI NÚT GIAO", new System.Drawing.Font("Segoe UI", 12, FontStyle.Bold), Brushes.LightGray, 20, 50);
        }
    }

    public class DoubleBufferedPanel : Panel
    {
        public DoubleBufferedPanel() 
        { 
            this.DoubleBuffered = true; 
            this.ResizeRedraw = true;
        }
    }
}
