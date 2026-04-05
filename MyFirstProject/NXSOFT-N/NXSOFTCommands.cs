// ==============================================================================
// NXSOFT-N Commands - Tích hợp vào MyFirstProject
// Namespace: MyFirstProject (build 1 DLL duy nhất)
// ==============================================================================

using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.DatabaseServices;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

// ── Alias để tránh xung đột namespace Autodesk ──
using SWF = System.Windows.Forms;
using SD  = System.Drawing;

[assembly: CommandClass(typeof(MyFirstProject.NXSOFTCommands))]

namespace MyFirstProject
{
    public class NXSOFTCommands
    {
        [CommandMethod("NXFIXTDTN")]
        public void FixTDTN()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n[NXSOFT-N] Sửa đường tự nhiên (NXFIXTDTN) - đang phát triển.");
        }

        [CommandMethod("NXCCTN")]
        public void NxCCTN()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n[NXSOFT-N] Điền chênh cao tim cọc (NXCCTN) - đang phát triển.");
        }

        [CommandMethod("NXPIPE")]
        public void NxPipe()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n[NXSOFT-N] Thiết kế mạng lưới thoát nước (NXPIPE) - đang phát triển.");
        }

        [CommandMethod("NXNTDADD")]
        public void NxNtdAdd()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n[NXSOFT-N] Nhập dữ liệu NTD vào tuyến (NXNTDADD) - đang phát triển.");
        }

        [CommandMethod("NXdcdcoc")]
        public void NxDcdCoc()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n[NXSOFT-N] Điền cao độ thiết kế lên mặt bằng (NXdcdcoc) - đang phát triển.");
        }

        [CommandMethod("NXNoiText")]
        public void NxNoiText()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n[NXSOFT-N] Gộp text khảo sát (NXNoiText) - đang phát triển.");
        }

        [CommandMethod("NXdcdtd")]
        public void NxDcdTd()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n[NXSOFT-N] Điền chênh cao 2 trắc dọc (NXdcdtd) - đang phát triển.");
        }

        // ── ĐỔI TÊN CỌC HÀNG LOẠT (có UI) ──
        [CommandMethod("NXrenameSL")]
        public void NxRenameSL()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db  = doc.Database;
            var ed  = doc.Editor;

            var selRes = ed.GetSelection(new SelectionFilter(new TypedValue[] {
                new TypedValue((int)DxfCode.Start, "AECC_SAMPLE_LINE")
            }));
            if (selRes.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nĐã hủy.");
                return;
            }

            using var form = new NXfrmRenameSL();
            var result = Application.ShowModalDialog(form);
            if (result != SWF.DialogResult.OK) return;

            string prefix = form.SelectedPrefix;

            using var tr = db.TransactionManager.StartTransaction();
            int count = 0;
            foreach (SelectedObject obj in selRes.Value)
            {
                var sl = tr.GetObject(obj.ObjectId, OpenMode.ForWrite) as SampleLine;
                if (sl == null) continue;
                double sta = sl.Station;
                int km = (int)(sta / 1000);
                double m = sta % 1000;
                sl.Name = $"{prefix}{km}+{m:0.00}";
                count++;
            }
            tr.Commit();
            ed.WriteMessage($"\n[NXSOFT-N] ✓ Đã đổi tên {count} cọc.");
        }

        // ── CHỈNH BỀ DÀY POLYLINE ──
        [CommandMethod("CWPL")]
        public void CwPl()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db  = doc.Database;
            var ed  = doc.Editor;

            var prDouble = ed.GetDouble("\nNhập bề dày Polyline (VD: 0.5): ");
            if (prDouble.Status != PromptStatus.OK) return;

            var selRes = ed.GetSelection(new SelectionFilter(new TypedValue[] {
                new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
            }));
            if (selRes.Status != PromptStatus.OK) return;

            using var tr = db.TransactionManager.StartTransaction();
            int count = 0;
            foreach (SelectedObject obj in selRes.Value)
            {
                var poly = tr.GetObject(obj.ObjectId, OpenMode.ForWrite) as Polyline;
                if (poly == null) continue;
                poly.ConstantWidth = prDouble.Value;
                count++;
            }
            tr.Commit();
            ed.WriteMessage($"\n[NXSOFT-N] ✓ Đã cập nhật bề dày {prDouble.Value} cho {count} polyline.");
        }
    }

    // ── FORM: ĐỔI TÊN CỌC ──
    public class NXfrmRenameSL : SWF.Form
    {
        public string SelectedPrefix { get; private set; } = "Km";

        private readonly SWF.TextBox txtPrefix;
        private readonly SWF.Button btnOk;
        private readonly SWF.Button btnCancel;

        public NXfrmRenameSL()
        {
            this.Text = "NXSOFT-N - Đổi Tên Cọc Hàng Loạt";
            this.Size = new SD.Size(400, 210);
            this.StartPosition = SWF.FormStartPosition.CenterParent;
            this.FormBorderStyle = SWF.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = SD.Color.FromArgb(30, 40, 60);
            this.ForeColor = SD.Color.White;

            // Tiêu đề
            var lblTitle = new SWF.Label
            {
                Text = "NXSOFT-N | Đổi Tên Cọc Hàng Loạt",
                Font = new SD.Font("Segoe UI", 10, SD.FontStyle.Bold),
                ForeColor = SD.Color.FromArgb(150, 200, 255),
                Location = new SD.Point(15, 12),
                AutoSize = true
            };

            // Mô tả
            var lblInfo = new SWF.Label
            {
                Text = "Nhập tiền tố → chương trình tự ghép lý trình.\nVí dụ: Km → Km0+000.00, Km0+020.00...",
                ForeColor = SD.Color.LightGray,
                Font = new SD.Font("Segoe UI", 8.5f),
                Location = new SD.Point(15, 38),
                Size = new SD.Size(365, 36)
            };

            // Label
            var lblInput = new SWF.Label
            {
                Text = "Tiền tố:",
                ForeColor = SD.Color.White,
                Font = new SD.Font("Segoe UI", 9),
                Location = new SD.Point(15, 86),
                AutoSize = true
            };

            // TextBox
            txtPrefix = new SWF.TextBox
            {
                Location = new SD.Point(80, 83),
                Size = new SD.Size(290, 28),
                Text = "Km",
                BackColor = SD.Color.FromArgb(50, 65, 90),
                ForeColor = SD.Color.White,
                BorderStyle = SWF.BorderStyle.FixedSingle,
                Font = new SD.Font("Segoe UI", 9)
            };

            // OK
            btnOk = new SWF.Button
            {
                Text = "✓ Áp Dụng",
                Location = new SD.Point(155, 125),
                Size = new SD.Size(105, 32),
                BackColor = SD.Color.FromArgb(27, 100, 180),
                ForeColor = SD.Color.White,
                FlatStyle = SWF.FlatStyle.Flat,
                Font = new SD.Font("Segoe UI", 9, SD.FontStyle.Bold),
                Cursor = SWF.Cursors.Hand
            };
            btnOk.FlatAppearance.BorderColor = SD.Color.FromArgb(46, 74, 122);
            btnOk.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtPrefix.Text))
                {
                    SWF.MessageBox.Show("Vui lòng nhập tiền tố!", "Thiếu thông tin",
                        SWF.MessageBoxButtons.OK, SWF.MessageBoxIcon.Warning);
                    return;
                }
                SelectedPrefix = txtPrefix.Text.Trim();
                DialogResult = SWF.DialogResult.OK;
                Close();
            };

            // Cancel
            btnCancel = new SWF.Button
            {
                Text = "✕ Hủy",
                Location = new SD.Point(270, 125),
                Size = new SD.Size(80, 32),
                BackColor = SD.Color.FromArgb(70, 70, 85),
                ForeColor = SD.Color.White,
                FlatStyle = SWF.FlatStyle.Flat,
                Font = new SD.Font("Segoe UI", 9),
                Cursor = SWF.Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderColor = SD.Color.FromArgb(100, 100, 120);
            btnCancel.Click += (s, e) => { DialogResult = SWF.DialogResult.Cancel; Close(); };

            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
            this.Controls.AddRange(new SWF.Control[] { lblTitle, lblInfo, lblInput, txtPrefix, btnOk, btnCancel });
        }
    }
}
