// 37.SampleLineHub.cs - SampleLine Hub — gộp 5 lệnh VI_SampleLine*
// Gộp: VI_SampleLineCoordinate, VI_ExportSampleLine, VI_ImportSampleLine,
//       VI_RenameSampleLine, VI_FillingSampleLine
// Form kế thừa CivilFormBase

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using ContentAlignment = System.Drawing.ContentAlignment;

[assembly: CommandClass(typeof(MyFirstProject.SampleLineHubCommands))]

namespace MyFirstProject
{
    // ══════════════════════════════════════════════════════════════
    // FORM: SAMPLE LINE HUB
    // ══════════════════════════════════════════════════════════════

    public class SampleLineHubForm : CivilFormBase
    {
        // ═══ Controls ═══
        private ComboBox cboAlignment = null!;
        private RadioButton rdoCoordinate = null!;
        private RadioButton rdoExport = null!;
        private RadioButton rdoImport = null!;
        private RadioButton rdoRename = null!;
        private RadioButton rdoFilling = null!;

        // Options panel
        private Panel pnlOptions = null!;
        private TextBox txtPrefix = null!;  // cho Rename
        private NumericUpDown numTextHeight = null!;  // cho Filling
        private TextBox txtLayer = null!;  // cho Filling

        // ═══ Data ═══
        private List<(string Name, ObjectId Id)> alignmentList = new();

        // ═══ Kết quả ═══
        public bool IsConfirmed { get; private set; }
        public ObjectId SelectedAlignmentId =>
            cboAlignment.SelectedIndex >= 0 ? alignmentList[cboAlignment.SelectedIndex].Id : ObjectId.Null;
        public string SelectedMode
        {
            get
            {
                if (rdoCoordinate.Checked) return "Coordinate";
                if (rdoExport.Checked) return "Export";
                if (rdoImport.Checked) return "Import";
                if (rdoRename.Checked) return "Rename";
                return "Filling";
            }
        }
        public string Prefix => txtPrefix.Text;
        public double TextHeight => (double)numTextHeight.Value;
        public string LayerName => string.IsNullOrWhiteSpace(txtLayer.Text) ? "VI_SL_TEXT" : txtLayer.Text;

        public SampleLineHubForm()
        {
            LoadAlignments();
            InitializeComponent();
        }

        private void LoadAlignments()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var civilDoc = CivilApplication.ActiveDocument;

            var alignmentIds = civilDoc.GetAlignmentIds();
            using var tr = doc.Database.TransactionManager.StartOpenCloseTransaction();
            foreach (ObjectId id in alignmentIds)
            {
                var alignment = tr.GetObject(id, OpenMode.ForRead) as Alignment;
                if (alignment != null)
                    alignmentList.Add((alignment.Name, id));
            }
            tr.Commit();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("📐  SAMPLE LINE HUB", 440, 480);

            // ═══ HEADER ═══
            var header = CreateHeader("SAMPLE LINE HUB");
            this.Controls.Add(header);

            var subtitle = CreateSubtitle("Gộp 5 công cụ Sample Line — chọn tuyến + chức năng");
            this.Controls.Add(subtitle);

            // ═══ CHỌN ALIGNMENT ═══
            var grpAlign = CreateGroupBox("📏 Chọn Alignment", 16, 82, 400, 60);
            this.Controls.Add(grpAlign);

            cboAlignment = CreateComboBox(12, 26, 376);
            foreach (var (name, _) in alignmentList)
                cboAlignment.Items.Add(name);
            if (cboAlignment.Items.Count > 0)
                cboAlignment.SelectedIndex = 0;
            grpAlign.Controls.Add(cboAlignment);

            // ═══ CHỌN CHỨC NĂNG ═══
            var grpFunc = CreateGroupBox("⚡ Chức năng", 16, 150, 400, 155);
            this.Controls.Add(grpFunc);

            int ry = 26;
            rdoCoordinate = CreateRadioBtn("📍 Xuất tọa độ cọc (Clipboard + Command Line)", 12, ry, grpFunc, true);
            ry += 25;
            rdoExport = CreateRadioBtn("📤 Xuất Sample Line ra CSV", 12, ry, grpFunc);
            ry += 25;
            rdoImport = CreateRadioBtn("📥 Nhập Sample Line từ CSV", 12, ry, grpFunc);
            ry += 25;
            rdoRename = CreateRadioBtn("✏️ Đổi tên Sample Line (theo lý trình)", 12, ry, grpFunc);
            ry += 25;
            rdoFilling = CreateRadioBtn("📝 Điền text lên vị trí cọc", 12, ry, grpFunc);

            // ═══ TÙY CHỌN THÊM ═══
            pnlOptions = CreateGroupBox("🔧 Tùy chọn", 16, 313, 400, 72);
            pnlOptions.Enabled = false;
            this.Controls.Add(pnlOptions);

            // Prefix (cho Rename)
            pnlOptions.Controls.Add(CreateLabel("Prefix:", 12, 30));
            txtPrefix = CreateTextBox(70, 27, 100, "Km0+");
            pnlOptions.Controls.Add(txtPrefix);

            // Text Height + Layer (cho Filling)
            pnlOptions.Controls.Add(CreateLabel("Cao chữ:", 180, 30));
            numTextHeight = CreateNumeric(240, 27, 60, 0.5m, 50, 2.5m, 1);
            pnlOptions.Controls.Add(numTextHeight);

            pnlOptions.Controls.Add(CreateLabel("Layer:", 310, 30));
            txtLayer = CreateTextBox(355, 27, 35, "");
            pnlOptions.Controls.Add(txtLayer);

            // Sự kiện bật/tắt options
            void UpdateOptions(object? s, EventArgs e)
            {
                pnlOptions.Enabled = rdoRename.Checked || rdoFilling.Checked;
                txtPrefix.Enabled = rdoRename.Checked;
                numTextHeight.Enabled = rdoFilling.Checked;
                txtLayer.Enabled = rdoFilling.Checked;
            }
            rdoCoordinate.CheckedChanged += UpdateOptions;
            rdoExport.CheckedChanged += UpdateOptions;
            rdoImport.CheckedChanged += UpdateOptions;
            rdoRename.CheckedChanged += UpdateOptions;
            rdoFilling.CheckedChanged += UpdateOptions;

            // ═══ BOTTOM BAR ═══
            var bottomBar = CreateBottomBar(56);
            this.Controls.Add(bottomBar);

            var btnCancel = CreateButton("✕ Hủy", 10, 10, 190, 36, bottomBar);
            btnCancel.Click += (s, e) => { IsConfirmed = false; this.Close(); };

            var btnRun = CreateSuccessButton("▶ Thực Hiện", 210, 10, 190, 36, bottomBar);
            btnRun.Click += (s, e) =>
            {
                if (cboAlignment.SelectedIndex < 0)
                {
                    System.Windows.Forms.MessageBox.Show("Vui lòng chọn Alignment!", "Thiếu dữ liệu",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                IsConfirmed = true;
                this.Close();
            };

            this.AcceptButton = btnRun;
            this.CancelButton = btnCancel;
        }

        private RadioButton CreateRadioBtn(string text, int x, int y, Control parent, bool isChecked = false)
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
    }

    // ══════════════════════════════════════════════════════════════
    // LỆNH SAMPLE LINE HUB
    // ══════════════════════════════════════════════════════════════

    public class SampleLineHubCommands
    {
        /// <summary>
        /// Mở SampleLine Hub — gộp 5 lệnh VI_SampleLine*
        /// </summary>
        [CommandMethod("VI_SampleLineHub")]
        public static void VI_SampleLineHub()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            using var form = new SampleLineHubForm();
            AcadApp.ShowModalDialog(form);
            if (!form.IsConfirmed) return;

            var alignmentId = form.SelectedAlignmentId;
            if (alignmentId.IsNull)
            {
                ed.WriteMessage("\n❌ Không chọn được Alignment.");
                return;
            }

            // Gọi lệnh tương ứng từ VisualINFRA.SampleLineTools
            // Các lệnh gốc vẫn giữ nguyên (backward compatible)
            switch (form.SelectedMode)
            {
                case "Coordinate":
                    VisualINFRA.SampleLineTools.SampleLineCoordinate();
                    break;
                case "Export":
                    VisualINFRA.SampleLineTools.ExportSampleLine();
                    break;
                case "Import":
                    VisualINFRA.SampleLineTools.ImportSampleLine();
                    break;
                case "Rename":
                    VisualINFRA.SampleLineTools.RenameSampleLine();
                    break;
                case "Filling":
                    VisualINFRA.SampleLineTools.FillingSampleLine();
                    break;
            }
        }
    }
}
