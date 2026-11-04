// 36.PrintCenter.cs - Trung tâm In ấn — gộp tất cả lệnh in/xuất PDF
// Gộp: CTDS_PrintAllLayouts, CTDS_PrintCurrentLayout, CTDS_ExportPDF
// Form kế thừa CivilFormBase

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using ContentAlignment = System.Drawing.ContentAlignment;

[assembly: CommandClass(typeof(MyFirstProject.PrintCenterCommands))]

namespace MyFirstProject
{
    // ══════════════════════════════════════════════════════════════
    // FORM: PRINT CENTER
    // ══════════════════════════════════════════════════════════════

    public class PrintCenterForm : CivilFormBase
    {
        // ═══ Controls ═══
        private RadioButton rdoPrintAll = null!;
        private RadioButton rdoPrintCurrent = null!;
        private RadioButton rdoExportPDF = null!;
        private RadioButton rdoInLayout = null!;
        private RadioButton rdoInModel = null!;

        // In hàng loạt options
        private Panel pnlBatchOptions = null!;
        private ComboBox cboPrinter = null!;
        private ComboBox cboPaperSize = null!;
        private ComboBox cboPlotStyle = null!;
        private ComboBox cboOrientation = null!;
        private ComboBox cboDirection = null!;
        private NumericUpDown numCount = null!;
        private NumericUpDown numSpacing = null!;

        // ═══ Kết quả ═══
        public bool IsConfirmed { get; private set; }
        public string SelectedMode
        {
            get
            {
                if (rdoPrintAll.Checked) return "PrintAll";
                if (rdoPrintCurrent.Checked) return "PrintCurrent";
                if (rdoExportPDF.Checked) return "ExportPDF";
                if (rdoInLayout.Checked) return "InLayout";
                return "InModel";
            }
        }

        // Thuộc tính cho in hàng loạt
        public string PrinterName => cboPrinter.Text;
        public string PaperSize => cboPaperSize.Text;
        public string PlotStyleName => GetPlotStyleName();
        public bool IsLandscape => cboOrientation.SelectedIndex == 0;
        public bool IsHorizontal => cboDirection.SelectedIndex == 0;
        public int PrintCount => (int)numCount.Value;
        public double Spacing => (double)numSpacing.Value;

        public PrintCenterForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("PRINT CENTER", 480, 560);

            // ═══ HEADER ═══
            var header = CreateHeader("PRINT CENTER — Trung Tâm In Ấn");
            this.Controls.Add(header);

            var subtitle = CreateSubtitle("Gộp tất cả chức năng in ấn & xuất PDF");
            this.Controls.Add(subtitle);

            // ═══ NHÓM 1: CHỌN CHẾ ĐỘ IN ═══
            var grpMode = CreateGroupBox("📋 Chọn chế độ in", 16, 82, 440, 145);
            this.Controls.Add(grpMode);

            int ry = 26;
            rdoPrintAll = CreateRadio("🖨️ In TẤT CẢ Layouts (theo Page Setup)", 12, ry, grpMode, true);
            ry += 24;
            rdoPrintCurrent = CreateRadio("📄 In Layout HIỆN TẠI (theo Page Setup)", 12, ry, grpMode);
            ry += 24;
            rdoExportPDF = CreateRadio("📑 Xuất PDF (ExportPDF)", 12, ry, grpMode);
            ry += 24;
            rdoInLayout = CreateRadio("📐 In hàng loạt LAYOUT (chọn khung, scale 1:1)", 12, ry, grpMode);
            ry += 24;
            rdoInModel = CreateRadio("📏 In hàng loạt MODEL (chọn khung, scale Fit)", 12, ry, grpMode);

            // ═══ NHÓM 2: CÀI ĐẶT IN HÀNG LOẠT ═══
            pnlBatchOptions = CreateGroupBox("⚙ Cài đặt in hàng loạt", 16, 234, 440, 230);
            this.Controls.Add(pnlBatchOptions);
            pnlBatchOptions.Enabled = false;

            int by = 26;
            int lx = 12, ix = 140, iw = 280;

            // Máy in
            pnlBatchOptions.Controls.Add(CreateLabel("Máy in:", lx, by + 3));
            cboPrinter = CreateComboBox(ix, by, iw);
            cboPrinter.Items.AddRange(new object[] {
                "AutoCAD PDF (General Documentation).pc3",
                "DWG To PDF.pc3",
                "PublishToWeb JPG.pc3",
                "PublishToWeb PNG.pc3"
            });
            cboPrinter.SelectedIndex = 0;
            cboPrinter.DropDownStyle = ComboBoxStyle.DropDown;
            pnlBatchOptions.Controls.Add(cboPrinter);
            by += 32;

            // Khổ giấy
            pnlBatchOptions.Controls.Add(CreateLabel("Khổ giấy:", lx, by + 3));
            cboPaperSize = CreateComboBox(ix, by, iw);
            cboPaperSize.Items.AddRange(new object[] {
                "ISO_full_bleed_A4_(210.00_x_297.00_MM)",
                "ISO_full_bleed_A3_(297.00_x_420.00_MM)",
                "ISO_full_bleed_A2_(420.00_x_594.00_MM)",
                "ISO_full_bleed_A1_(594.00_x_841.00_MM)",
                "ISO_full_bleed_A0_(841.00_x_1189.00_MM)"
            });
            cboPaperSize.SelectedIndex = 1;
            cboPaperSize.DropDownStyle = ComboBoxStyle.DropDown;
            pnlBatchOptions.Controls.Add(cboPaperSize);
            by += 32;

            // Plot Style
            pnlBatchOptions.Controls.Add(CreateLabel("Plot Style:", lx, by + 3));
            cboPlotStyle = CreateComboBox(ix, by, iw);
            cboPlotStyle.Items.AddRange(new object[] {
                "monochrome.ctb (đen trắng)",
                "acad.ctb (theo màu)",
                "Screening 25%.ctb",
                "Screening 50%.ctb",
                "Không dùng (None)"
            });
            cboPlotStyle.SelectedIndex = 0;
            pnlBatchOptions.Controls.Add(cboPlotStyle);
            by += 32;

            // Hướng giấy + Hướng dịch
            pnlBatchOptions.Controls.Add(CreateLabel("Hướng giấy:", lx, by + 3));
            cboOrientation = CreateComboBox(ix, by, 130);
            cboOrientation.Items.AddRange(new object[] { "⬌ Landscape", "⬍ Portrait" });
            cboOrientation.SelectedIndex = 0;
            pnlBatchOptions.Controls.Add(cboOrientation);

            cboDirection = CreateComboBox(ix + 140, by, 140);
            cboDirection.Items.AddRange(new object[] { "→ Sang phải", "↓ Xuống dưới" });
            cboDirection.SelectedIndex = 0;
            pnlBatchOptions.Controls.Add(cboDirection);
            by += 32;

            // Số lượng + Khoảng cách
            pnlBatchOptions.Controls.Add(CreateLabel("Số lượng:", lx, by + 3));
            numCount = CreateNumeric(ix, by, 80, 1, 200, 1, 0);
            pnlBatchOptions.Controls.Add(numCount);

            pnlBatchOptions.Controls.Add(CreateLabel("K/cách:", ix + 90, by + 3));
            numSpacing = CreateNumeric(ix + 150, by, 80, 0, 99999, 420, 1);
            pnlBatchOptions.Controls.Add(numSpacing);

            // ═══ Sự kiện bật/tắt batch panel ═══
            void UpdateBatchPanel(object? s, EventArgs e)
            {
                bool isBatch = rdoInLayout.Checked || rdoInModel.Checked;
                pnlBatchOptions.Enabled = isBatch;
                pnlBatchOptions.BackColor = isBatch ? BgPanel : Color.FromArgb(235, 235, 235);
            }
            rdoPrintAll.CheckedChanged += UpdateBatchPanel;
            rdoPrintCurrent.CheckedChanged += UpdateBatchPanel;
            rdoExportPDF.CheckedChanged += UpdateBatchPanel;
            rdoInLayout.CheckedChanged += UpdateBatchPanel;
            rdoInModel.CheckedChanged += UpdateBatchPanel;

            // ═══ BOTTOM BAR ═══
            var bottomBar = CreateBottomBar(56);
            this.Controls.Add(bottomBar);

            var btnCancel = CreateButton("✕ Hủy", 10, 10, 210, 36, bottomBar);
            btnCancel.Click += (s, e) => { IsConfirmed = false; this.Close(); };

            var btnRun = CreateSuccessButton("🖨️ Thực Hiện", 230, 10, 210, 36, bottomBar);
            btnRun.Click += (s, e) => { IsConfirmed = true; this.Close(); };

            this.AcceptButton = btnRun;
            this.CancelButton = btnCancel;
        }

        private RadioButton CreateRadio(string text, int x, int y, Control parent, bool isChecked = false)
        {
            var rdo = new RadioButton
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                Checked = isChecked,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            parent.Controls.Add(rdo);
            return rdo;
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
    // LỆNH PRINT CENTER
    // ══════════════════════════════════════════════════════════════

    public class PrintCenterCommands
    {
        /// <summary>
        /// Mở Print Center — gộp tất cả lệnh in ấn
        /// </summary>
        [CommandMethod("CT_PrintCenter")]
        public static void CT_PrintCenter()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            using var form = new PrintCenterForm();
            AcadApp.ShowModalDialog(form);
            if (!form.IsConfirmed) return;

            try
            {
                switch (form.SelectedMode)
                {
                    case "PrintAll":
                        PrintAllLayouts(doc, ed);
                        break;

                    case "PrintCurrent":
                        ed.WriteMessage($"\n▸ Đang in layout hiện tại...");
                        doc.SendStringToExecute("-PLOT N \"\" \"\" \"\" \"\" \"\" \"\" ", true, false, false);
                        break;

                    case "ExportPDF":
                        ed.WriteMessage("\n▸ Đang xuất PDF...");
                        doc.SendStringToExecute("EXPORTPDF ", true, false, false);
                        break;

                    case "InLayout":
                        InHangLoat(doc, ed, form, isLayout: true);
                        break;

                    case "InModel":
                        InHangLoat(doc, ed, form, isLayout: false);
                        break;
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n⊘ Lỗi: {ex.Message}");
            }
        }

        private static void PrintAllLayouts(Autodesk.AutoCAD.ApplicationServices.Document doc, Editor ed)
        {
            var layoutMgr = LayoutManager.Current;
            ed.WriteMessage("\n▸ Đang in tất cả layouts...");

            var layouts = new List<string>();
            var dbDict = doc.Database.LayoutDictionaryId;

            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(dbDict, OpenMode.ForRead);
                foreach (var entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    if (!layout.ModelType)
                        layouts.Add(layout.LayoutName);
                }
                tr.Commit();
            }

            foreach (string layoutName in layouts)
            {
                layoutMgr.CurrentLayout = layoutName;
                doc.SendStringToExecute("-PLOT N \"\" \"\" \"\" \"\" \"\" \"\" ", true, false, false);
            }

            ed.WriteMessage($"\n✅ Đã gửi lệnh in cho {layouts.Count} layout(s).");
        }

        private static void InHangLoat(Autodesk.AutoCAD.ApplicationServices.Document doc, Editor ed,
            PrintCenterForm form, bool isLayout)
        {
            string printerName = form.PrinterName;
            string paperSize = form.PaperSize;
            string plotStyle = form.PlotStyleName;
            string orientation = form.IsLandscape ? "Landscape" : "Portrait";
            bool isHorizontal = form.IsHorizontal;
            int count = form.PrintCount;
            double spacing = form.Spacing;

            // Chọn 2 điểm khung in
            var pt1Result = ed.GetPoint("\n▸ Chọn điểm GÓC THỨ NHẤT của khung in: ");
            if (pt1Result.Status != PromptStatus.OK) return;
            Point3d p1 = pt1Result.Value;

            var pt2Result = ed.GetCorner("\n▸ Chọn điểm GÓC THỨ HAI (đối diện): ", p1);
            if (pt2Result.Status != PromptStatus.OK) return;
            Point3d p2 = pt2Result.Value;

            ed.WriteMessage($"\n\n▸ Bắt đầu in {count} bản vẽ...\n");

            for (int i = 0; i < count; i++)
            {
                double offsetX = isHorizontal ? spacing * i : 0;
                double offsetY = isHorizontal ? 0 : -spacing * i;

                string point1 = $"{p1.X + offsetX},{p1.Y + offsetY}";
                string point2 = $"{p2.X + offsetX},{p2.Y + offsetY}";

                ed.WriteMessage($"\n  [{i + 1}/{count}] Đang in...");

                if (isLayout)
                {
                    doc.SendStringToExecute(
                        $"-PLOT Y \"\" \"{printerName}\" \"{paperSize}\" Millimeters {orientation} No Window {point1} {point2} 1 Center Yes \"{plotStyle}\" Yes Yes Yes No No No Yes ",
                        true, false, false);
                }
                else
                {
                    doc.SendStringToExecute(
                        $"-PLOT Y \"\" \"{printerName}\" \"{paperSize}\" Millimeters {orientation} No Window {point1} {point2} Fit Center Yes \"{plotStyle}\" Yes \"As displayed\" No Yes Yes ",
                        true, false, false);
                }
            }

            ed.WriteMessage($"\n\n✅ Đã gửi lệnh in cho {count} bản vẽ!");
        }
    }
}
