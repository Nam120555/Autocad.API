// 35.InHangLoat.cs - In hàng loạt bản vẽ PDF từ Layout/Model
// Chuyển đổi từ LISP: INLAYOUT, INMODEL
// Nâng cấp: Form chọn chuyên nghiệp, kế thừa CivilFormBase

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using ContentAlignment = System.Drawing.ContentAlignment;

[assembly: CommandClass(typeof(MyFirstProject.InHangLoatCommands))]

namespace MyFirstProject
{
    // ══════════════════════════════════════════════════════════════
    // FORM CHỌN CẤU HÌNH IN HÀNG LOẠT
    // ══════════════════════════════════════════════════════════════

    /// <summary>
    /// Form cấu hình in hàng loạt — kế thừa CivilFormBase
    /// Cho chọn: chế độ (Layout/Model), máy in, khổ giấy, plot style,
    /// hướng giấy, hướng dịch chuyển, số lượng
    /// </summary>
    public class InHangLoatForm : CivilFormBase
    {
        // ═══ Controls ═══
        private ComboBox cboMode = null!;
        private ComboBox cboPrinter = null!;
        private ComboBox cboPaperSize = null!;
        private ComboBox cboPlotStyle = null!;
        private ComboBox cboOrientation = null!;
        private ComboBox cboDirection = null!;
        private NumericUpDown numCount = null!;
        private NumericUpDown numSpacing = null!;

        private Panel previewPanel = null!;

        // ═══ Kết quả ═══
        public bool IsConfirmed { get; private set; }
        public bool IsLayoutMode => cboMode.SelectedIndex == 0;
        public string PrinterName => GetPrinterName();
        public string PaperSize => cboPaperSize.Text;
        public string PlotStyleName => GetPlotStyleName();
        public bool IsLandscape => cboOrientation.SelectedIndex == 0;
        public bool IsHorizontal => cboDirection.SelectedIndex == 0;
        public int PrintCount => (int)numCount.Value;
        public double Spacing => (double)numSpacing.Value;

        public InHangLoatForm(bool startAsLayout)
        {
            InitializeComponent(startAsLayout);
        }

        private void InitializeComponent(bool startAsLayout)
        {
            InitBaseTheme("🖨️  IN HÀNG LOẠT PDF", 480, 580);

            // ═══ HEADER ═══
            var header = CreateHeader("IN HÀNG LOẠT PDF");
            this.Controls.Add(header);

            var subtitle = CreateSubtitle("Chọn 2 điểm khung in → dịch chuyển tự động → in hàng loạt ra PDF");
            this.Controls.Add(subtitle);

            int left = 20;

            int inputX = 155;
            int inputW = 290;
            int startY = 90;
            int rowH = 38;

            // ═══ CHẾ ĐỘ IN ═══
            var grpMode = CreateGroupBox("⚙ Cấu Hình In", left, startY, 440, rowH * 7 + 30);
            this.Controls.Add(grpMode);

            int gy = 28; // y bên trong group box

            // — Chế độ —
            grpMode.Controls.Add(CreateLabel("Chế độ in:", 12, gy + 3));
            cboMode = CreateComboBox(inputX - left + 5, gy, inputW);
            cboMode.Items.AddRange(new object[] {
                "📐 Layout (tỷ lệ 1:1)",
                "📏 Model (tỷ lệ Fit)"
            });
            cboMode.SelectedIndex = startAsLayout ? 0 : 1;
            grpMode.Controls.Add(cboMode);
            gy += rowH;

            // — Máy in —
            grpMode.Controls.Add(CreateLabel("Máy in:", 12, gy + 3));
            cboPrinter = CreateComboBox(inputX - left + 5, gy, inputW);
            cboPrinter.Items.AddRange(new object[] {
                "AutoCAD PDF (General Documentation).pc3",
                "DWG To PDF.pc3",
                "PublishToWeb JPG.pc3",
                "PublishToWeb PNG.pc3"
            });
            cboPrinter.SelectedIndex = 0;
            cboPrinter.DropDownStyle = ComboBoxStyle.DropDown; // cho phép nhập tùy chỉnh
            grpMode.Controls.Add(cboPrinter);
            gy += rowH;

            // — Khổ giấy —
            grpMode.Controls.Add(CreateLabel("Khổ giấy:", 12, gy + 3));
            cboPaperSize = CreateComboBox(inputX - left + 5, gy, inputW);
            cboPaperSize.Items.AddRange(new object[] {
                "ISO_full_bleed_A4_(210.00_x_297.00_MM)",
                "ISO_full_bleed_A3_(297.00_x_420.00_MM)",
                "ISO_full_bleed_A2_(420.00_x_594.00_MM)",
                "ISO_full_bleed_A1_(594.00_x_841.00_MM)",
                "ISO_full_bleed_A0_(841.00_x_1189.00_MM)"
            });
            cboPaperSize.SelectedIndex = 1; // A3 mặc định
            cboPaperSize.DropDownStyle = ComboBoxStyle.DropDown;
            grpMode.Controls.Add(cboPaperSize);
            gy += rowH;

            // — Plot Style —
            grpMode.Controls.Add(CreateLabel("Plot Style:", 12, gy + 3));
            cboPlotStyle = CreateComboBox(inputX - left + 5, gy, inputW);
            cboPlotStyle.Items.AddRange(new object[] {
                "monochrome.ctb (đen trắng)",
                "acad.ctb (theo màu)",
                "Screening 25%.ctb",
                "Screening 50%.ctb",
                "Không dùng (None)"
            });
            cboPlotStyle.SelectedIndex = 0;
            grpMode.Controls.Add(cboPlotStyle);
            gy += rowH;

            // — Hướng giấy —
            grpMode.Controls.Add(CreateLabel("Hướng giấy:", 12, gy + 3));
            cboOrientation = CreateComboBox(inputX - left + 5, gy, inputW);
            cboOrientation.Items.AddRange(new object[] {
                "⬌ Ngang (Landscape)",
                "⬍ Dọc (Portrait)"
            });
            cboOrientation.SelectedIndex = 0;
            grpMode.Controls.Add(cboOrientation);
            gy += rowH;

            // — Hướng dịch chuyển —
            grpMode.Controls.Add(CreateLabel("Hướng dịch:", 12, gy + 3));
            cboDirection = CreateComboBox(inputX - left + 5, gy, inputW);
            cboDirection.Items.AddRange(new object[] {
                "→ Sang phải (Ngang)",
                "↓ Xuống dưới (Dọc)"
            });
            cboDirection.SelectedIndex = 0;
            grpMode.Controls.Add(cboDirection);
            gy += rowH;

            // — Số lượng —
            grpMode.Controls.Add(CreateLabel("Số lượng BV:", 12, gy + 3));
            numCount = CreateNumeric(inputX - left + 5, gy, 100, 1, 200, 1, 0);
            grpMode.Controls.Add(numCount);

            // — Khoảng cách —
            var lblSpacing = CreateLabel("Khoảng cách:", inputX - left + 115, gy + 3);
            grpMode.Controls.Add(lblSpacing);
            numSpacing = CreateNumeric(inputX - left + 215, gy, 80, 0, 99999, 420, 1);
            grpMode.Controls.Add(numSpacing);

            // ═══ PREVIEW PANEL ═══
            int previewY = startY + rowH * 7 + 45;
            previewPanel = new Panel
            {
                Location = new Point(left, previewY),
                Size = new Size(440, 80),
                BackColor = BgPanel,
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(previewPanel);

            previewPanel.Paint += PreviewPanel_Paint;

            // Cập nhật preview khi thay đổi
            cboDirection.SelectedIndexChanged += (s, e) => previewPanel.Invalidate();
            numCount.ValueChanged += (s, e) => previewPanel.Invalidate();

            // ═══ HINT ═══
            var hint = CreateHintLabel(
                "💡 Sau khi bấm In, chọn 2 điểm góc khung in trong bản vẽ.",
                left, previewY + 90);
            this.Controls.Add(hint);

            // ═══ BOTTOM BAR ═══
            var bottomBar = CreateBottomBar(56);
            this.Controls.Add(bottomBar);

            var btnCancel = CreateButton("✕ Hủy", 10, 10, 210, 36, bottomBar);
            btnCancel.Click += (s, e) => { IsConfirmed = false; this.Close(); };

            var btnPrint = CreateSuccessButton("🖨️ Chọn Khung In", 230, 10, 210, 36, bottomBar);
            btnPrint.Click += (s, e) => { IsConfirmed = true; this.Close(); };

            // Phím tắt Enter = In, Esc = Hủy
            this.AcceptButton = btnPrint;
            this.CancelButton = btnCancel;
        }

        // ═══ VẼ PREVIEW ═══
        private void PreviewPanel_Paint(object? sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            int count = Math.Min((int)numCount.Value, 8); // hiển thị tối đa 8
            bool isHoriz = cboDirection.SelectedIndex == 0;

            // Tiêu đề
            using var titleFont = new Font("Segoe UI", 8f, FontStyle.Bold);
            g.DrawString($"Xem trước: {(int)numCount.Value} bản vẽ, hướng {(isHoriz ? "Ngang →" : "Dọc ↓")}",
                titleFont, new SolidBrush(NavyDark), 8, 4);

            // Vẽ các khung
            int boxW = isHoriz ? 40 : 30;
            int boxH = isHoriz ? 50 : 40;
            int gap = 6;
            int startX = 10;
            int startYbox = 24;

            using var fillBrush = new SolidBrush(Color.FromArgb(200, 220, 245));
            using var borderPen = new Pen(NavyDark, 1.5f);
            using var numFont = new Font("Segoe UI", 7f, FontStyle.Bold);
            using var numBrush = new SolidBrush(NavyDark);
            using var arrowPen = new Pen(Color.FromArgb(46, 160, 67), 2f);
            arrowPen.EndCap = System.Drawing.Drawing2D.LineCap.ArrowAnchor;

            for (int i = 0; i < count; i++)
            {
                int bx = isHoriz ? startX + i * (boxW + gap) : startX;
                int by = isHoriz ? startYbox : startYbox + i * (boxH / 2 + gap / 2);

                // Khung bản vẽ
                g.FillRectangle(fillBrush, bx, by, boxW, boxH);
                g.DrawRectangle(borderPen, bx, by, boxW, boxH);

                // Số thứ tự
                string num = (i + 1).ToString();
                var numSize = g.MeasureString(num, numFont);
                g.DrawString(num, numFont, numBrush,
                    bx + (boxW - numSize.Width) / 2,
                    by + (boxH - numSize.Height) / 2);

                // Mũi tên giữa các khung
                if (i < count - 1)
                {
                    if (isHoriz)
                    {
                        g.DrawLine(arrowPen,
                            bx + boxW + 2, by + boxH / 2,
                            bx + boxW + gap - 2, by + boxH / 2);
                    }
                    else
                    {
                        g.DrawLine(arrowPen,
                            bx + boxW / 2, by + boxH + 2,
                            bx + boxW / 2, by + boxH + gap / 2 - 1);
                    }
                }
            }

            // Nếu còn nhiều hơn
            if ((int)numCount.Value > 8)
            {
                int dotX = isHoriz ? startX + count * (boxW + gap) : startX + boxW + 10;
                int dotY = isHoriz ? startYbox + boxH / 2 - 4 : startYbox + count * (boxH / 2 + gap / 2);
                g.DrawString($"... +{(int)numCount.Value - 8}", numFont, numBrush, dotX, dotY);
            }
        }

        private string GetPrinterName()
        {
            return cboPrinter.Text;
        }

        private string GetPlotStyleName()
        {
            string sel = cboPlotStyle.Text;
            if (sel.Contains("monochrome")) return "monochrome.ctb";
            if (sel.Contains("acad")) return "acad.ctb";
            if (sel.Contains("25%")) return "Screening 25%.ctb";
            if (sel.Contains("50%")) return "Screening 50%.ctb";
            return "None";
        }
    }

    // ══════════════════════════════════════════════════════════════
    // LỆNH INL & INM
    // ══════════════════════════════════════════════════════════════

    /// <summary>
    /// Lệnh in hàng loạt bản vẽ PDF — chuyển từ LISP INLAYOUT/INMODEL
    /// INL = In Layout (1:1), INM = In Model (Fit)
    /// </summary>
    public class InHangLoatCommands
    {
        /// <summary>
        /// In hàng loạt bản vẽ PDF từ Layout space (scale 1:1)
        /// </summary>
        [CommandMethod("INL")]
        public static void INL()
        {
            InHangLoat(startAsLayout: true);
        }

        /// <summary>
        /// In hàng loạt bản vẽ PDF từ Model space (scale Fit)
        /// </summary>
        [CommandMethod("INM")]
        public static void INM()
        {
            InHangLoat(startAsLayout: false);
        }

        // ═══ LOGIC CHÍNH ═══
        private static void InHangLoat(bool startAsLayout)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            try
            {
                // ═══ 1. Hiển thị Form chọn cấu hình ═══
                using var form = new InHangLoatForm(startAsLayout);
                var result = AcadApp.ShowModalDialog(form);

                if (!form.IsConfirmed) return;

                // Lấy cấu hình từ form
                bool isLayout = form.IsLayoutMode;
                string printerName = form.PrinterName;
                string paperSize = form.PaperSize;
                string plotStyle = form.PlotStyleName;
                string orientation = form.IsLandscape ? "Landscape" : "Portrait";
                bool isHorizontal = form.IsHorizontal;
                int count = form.PrintCount;
                double spacing = form.Spacing;

                string mode = isLayout ? "LAYOUT (1:1)" : "MODEL (Fit)";

                // ═══ 2. Chọn 2 điểm khung in trong bản vẽ ═══
                ed.WriteMessage($"\n▸ Chế độ: {mode} — {count} bản vẽ");

                var pt1Result = ed.GetPoint("\n▸ Chọn điểm GÓC THỨ NHẤT của khung in: ");
                if (pt1Result.Status != PromptStatus.OK) return;
                Point3d p1 = pt1Result.Value;

                var pt2Result = ed.GetCorner("\n▸ Chọn điểm GÓC THỨ HAI (đối diện): ", p1);
                if (pt2Result.Status != PromptStatus.OK) return;
                Point3d p2 = pt2Result.Value;

                // ═══ 3. Thực hiện in ═══
                ed.WriteMessage($"\n\n┌─────────────────────────────────────────┐");
                ed.WriteMessage($"\n│ Bắt đầu in {count} bản vẽ...            │");
                ed.WriteMessage($"\n│ Máy in:    {printerName,-28}│");
                ed.WriteMessage($"\n│ Khổ giấy:  {paperSize,-28}│");
                ed.WriteMessage($"\n│ Hướng:     {orientation,-28}│");
                ed.WriteMessage($"\n└─────────────────────────────────────────┘\n");

                for (int i = 0; i < count; i++)
                {
                    // Tính vị trí khung in mới (dịch chuyển theo khoảng cách)
                    double offsetX = isHorizontal ? spacing * i : 0;
                    double offsetY = isHorizontal ? 0 : -spacing * i;

                    double x1 = p1.X + offsetX;
                    double y1 = p1.Y + offsetY;
                    double x2 = p2.X + offsetX;
                    double y2 = p2.Y + offsetY;

                    string point1 = $"{x1},{y1}";
                    string point2 = $"{x2},{y2}";

                    ed.WriteMessage($"\n  [{i + 1}/{count}] Đang in... ({point1} → {point2})");

                    if (isLayout)
                    {
                        // ═══ LAYOUT MODE — Scale 1:1 ═══
                        doc.SendStringToExecute(
                            $"-PLOT Y " +
                            $"\"\" " +
                            $"\"{printerName}\" " +
                            $"\"{paperSize}\" " +
                            $"Millimeters " +
                            $"{orientation} " +
                            $"No " +
                            $"Window " +
                            $"{point1} " +
                            $"{point2} " +
                            $"1 " +
                            $"Center " +
                            $"Yes " +
                            $"\"{plotStyle}\" " +
                            $"Yes " +
                            $"Yes " +
                            $"Yes " +
                            $"No " +
                            $"No " +
                            $"No " +
                            $"Yes ",
                            true, false, false);
                    }
                    else
                    {
                        // ═══ MODEL MODE — Scale Fit ═══
                        doc.SendStringToExecute(
                            $"-PLOT Y " +
                            $"\"\" " +
                            $"\"{printerName}\" " +
                            $"\"{paperSize}\" " +
                            $"Millimeters " +
                            $"{orientation} " +
                            $"No " +
                            $"Window " +
                            $"{point1} " +
                            $"{point2} " +
                            $"Fit " +
                            $"Center " +
                            $"Yes " +
                            $"\"{plotStyle}\" " +
                            $"Yes " +
                            $"\"As displayed\" " +
                            $"No " +
                            $"Yes " +
                            $"Yes ",
                            true, false, false);
                    }
                }

                ed.WriteMessage($"\n\n✅ Đã gửi lệnh in cho {count} bản vẽ!");
                ed.WriteMessage("\n   ⚠ Các bản vẽ sẽ lần lượt được in ra.");
                ed.WriteMessage("\n   ⚠ File PDF sẽ lưu tại vị trí do printer quy định.\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n⊘ Lỗi: {ex.Message}");
            }
        }
    }
}
