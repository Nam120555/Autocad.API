using System;
using System.Drawing;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.DatabaseServices;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(NXSOFT_N.MainCommands))]

namespace NXSOFT_N
{
    public class MainCommands
    {
        [CommandMethod("NXFIXTDTN")]
        public void FixTDTN()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n[NXSOFT-N] Đang khởi động Box Lệnh: Sửa đường tự nhiên (NXFIXTDTN) ...");
        }

        [CommandMethod("NXCCTN")]
        public void NxCCTN()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n[NXSOFT-N] Đang khởi động Box Lệnh: Điền chênh cao tim cọc (NXCCTN) ...");
        }

        [CommandMethod("NXPIPE")]
        public void NxPipe()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n[NXSOFT-N] Đang khởi động Box Lệnh: Thiết kế mạng lưới thoát nước (NXPIPE) ...");
        }

        [CommandMethod("NXNTDADD")]
        public void NxNtdAdd()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n[NXSOFT-N] Đang khởi động Box Lệnh: Nhập dữ liệu NTD vào tuyến (NXNTDADD) ...");
        }

        [CommandMethod("NXDCDCOC")]
        public void NxDcdCoc()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n[NXSOFT-N] Đang khởi động Box Lệnh: Điền cao độ thiết kế lên mặt bằng (NXDCDCOC) ...");
        }

        // ==========================================
        // LỆNH ĐỔI TÊN CỌC (UX Winform đầy đủ)
        // ==========================================
        [CommandMethod("NXrenameSL")]
        public void NxRenameSL()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // 1. Quét chọn đối tượng Sample Line (Cọc)
            PromptSelectionResult selRes = ed.GetSelection(new SelectionFilter(new TypedValue[] {
                new TypedValue((int)DxfCode.Start, "AECC_SAMPLE_LINE")
            }));

            if (selRes.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nĐã hủy chọn cọc.");
                return;
            }

            // 2. Mở Hộp thoại (Box lệnh) UI chuẩn
            using (var form = new frmRenameSL())
            {
                var dialogResult = Application.ShowModalDialog(form);
                if (dialogResult != System.Windows.Forms.DialogResult.OK) return;

                string prefix = form.Prefix;

                // 3. Thực thi thuật toán
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    int count = 0;
                    foreach (SelectedObject obj in selRes.Value)
                    {
                        var sl = tr.GetObject(obj.ObjectId, OpenMode.ForWrite) as SampleLine;
                        if (sl != null)
                        {
                            double sta = sl.Station;
                            int km = (int)(sta / 1000);
                            double m = sta % 1000;
                            // VD: Km 1 + 250.50
                            string slName = $"{prefix}{km}+{m.ToString("0.00")}";
                            sl.Name = slName;
                            count++;
                        }
                    }
                    tr.Commit();
                    ed.WriteMessage($"\n[NXSOFT-N] Đã thay đổi thành công tên cho {count} cọc.");
                }
            }
        }

        // ==========================================
        // LỆNH CHỈNH CHIỀU DÀY PLINE NHANH
        // ==========================================
        [CommandMethod("CWPL")]
        public void CwPl()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            PromptDoubleResult prDouble = ed.GetDouble("\nNhập bề dày liên tục cho Polyline (Ví dụ: 0.5): ");
            if (prDouble.Status != PromptStatus.OK) return;

            PromptSelectionResult selRes = ed.GetSelection(new SelectionFilter(new TypedValue[] {
                new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
            }));

            if (selRes.Status != PromptStatus.OK) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                int count = 0;
                foreach (SelectedObject obj in selRes.Value)
                {
                    var poly = tr.GetObject(obj.ObjectId, OpenMode.ForWrite) as Polyline;
                    if (poly != null)
                    {
                        poly.ConstantWidth = prDouble.Value;
                        count++;
                    }
                }
                tr.Commit();
                ed.WriteMessage($"\n[NXSOFT-N] Đã cập nhật bề dày cho {count} polyline.");
            }
        }

        // ==========================================
        // CÁC LỆNH CÒN LẠI (Sẽ triển khai tiếp)
        // ==========================================
        [CommandMethod("NXNoiText")]
        public void NxNoiText()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n[NXSOFT-N] Đang chạy lệnh Gộp text khảo sát (NXNoiText) ...");
        }

        [CommandMethod("NXdcdtd")]
        public void NxDcdTd()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n[NXSOFT-N] Đang chạy lệnh Điền chênh cao 2 trắc dọc (NXdcdtd) ...");
        }
    }

    // ========================================================
    // THIẾT KẾ UI BOX TỰ ĐỘNG - KHÔNG CẦN KÉO THẢ DESIGNER
    // ========================================================
    public class frmRenameSL : Form
    {
        public string Prefix { get; private set; }
        private TextBox txtPrefix;
        private Button btnOk;
        private Button btnCancel;

        public frmRenameSL()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "NXSOFT - Thay Đổi Tên Cọc Hàng Loạt";
            this.Size = new Size(380, 200);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.FromArgb(240, 240, 240); // Theo tông màu sáng/xám Civil 3D

            Label lblInfo = new Label();
            lblInfo.Text = "Nhập tiền tố muốn thêm vào tên cọc đã chọn (VD: Km, Cọc, ...).\nChương trình sẽ tự động đánh số trạm lý trình phía sau.";
            lblInfo.Location = new Point(20, 20);
            lblInfo.Size = new Size(330, 40);

            Label lblInput = new Label();
            lblInput.Text = "Tiền tố (Prefix):";
            lblInput.Location = new Point(20, 75);
            lblInput.AutoSize = true;

            txtPrefix = new TextBox();
            txtPrefix.Location = new Point(120, 72);
            txtPrefix.Size = new Size(220, 30);
            txtPrefix.Text = "Km";

            btnOk = new Button();
            btnOk.Text = "Áp Dụng";
            btnOk.Location = new Point(160, 115);
            btnOk.Size = new Size(90, 30);
            btnOk.BackColor = Color.Navy;
            btnOk.ForeColor = Color.White;
            btnOk.FlatStyle = FlatStyle.Flat;
            btnOk.Click += BtnOk_Click;

            btnCancel = new Button();
            btnCancel.Text = "Hủy";
            btnCancel.Location = new Point(260, 115);
            btnCancel.Size = new Size(80, 30);
            btnCancel.BackColor = Color.LightGray;
            btnCancel.FlatStyle = FlatStyle.Flat;
            btnCancel.Click += BtnCancel_Click;

            this.Controls.Add(lblInfo);
            this.Controls.Add(lblInput);
            this.Controls.Add(txtPrefix);
            this.Controls.Add(btnOk);
            this.Controls.Add(btnCancel);
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            this.Prefix = txtPrefix.Text.Trim();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
