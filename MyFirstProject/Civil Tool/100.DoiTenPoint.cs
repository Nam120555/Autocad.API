using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using MyFirstProject.Extensions;

using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using ContentAlignment = System.Drawing.ContentAlignment;

[assembly: CommandClass(typeof(MyFirstProject.DoiTenPoint))]

namespace MyFirstProject
{
    public class DoiTenPointForm : CivilFormBase
    {
        // ═══ UI Controls ═══
        private RadioButton rdoPointGroup;
        private RadioButton rdoSelectManual;
        private ComboBox cboPointGroup;
        private CheckBox chkRenameGroup;
        
        private TextBox txtPrefix;
        private NumericUpDown numStartIdx;
        private NumericUpDown numStep;
        
        private DataGridView dgvPreview;
        private Button btnPreview;
        private WinLabel lblTotal;
        
        // ═══ Data ═══
        public bool IsConfirmed { get; private set; }
        public bool IsByGroup => rdoPointGroup.Checked;
        public string SelectedGroupName => cboPointGroup.SelectedItem?.ToString() ?? "";
        public bool RenameGroup => chkRenameGroup.Checked && IsByGroup;
        public string Prefix => txtPrefix.Text.Trim();
        public int StartIndex => (int)numStartIdx.Value;
        public int Step => (int)numStep.Value;
        
        // Selected point IDs (if manual)
        public List<ObjectId> SelectedPointIds { get; set; } = new();

        public class PreviewItem
        {
            public ObjectId Id { get; set; }
            public uint PointNumber { get; set; }
            public string OldName { get; set; }
            public string NewName { get; set; }
        }
        
        public List<PreviewItem> PreviewItems { get; private set; } = new();

        public DoiTenPointForm()
        {
            InitializeComponent();
            LoadPointGroups();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("ĐỔI TÊN COGO POINT", 750, 680);

            var header = CreateHeader("ĐỔI TÊN COGO POINT THEO QUY TẮC");
            this.Controls.Add(header);

            var subtitle = CreateSubtitle("Quy tắc: Hằng số (Tiền tố) + Số thứ tự (VD: P1, P2, P3...)");
            this.Controls.Add(subtitle);

            var mainPanel = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(15) };
            this.Controls.Add(mainPanel);
            mainPanel.BringToFront();

            int y = 10;

            // ═══ 1. NGUỒN DỮ LIỆU ═══
            var grpSource = CreateGroupBox("Nguồn dữ liệu", 0, y, 700, 115);
            mainPanel.Controls.Add(grpSource);

            rdoPointGroup = new RadioButton { Text = "Theo Point Group", Location = new Point(15, 30), AutoSize = true, Checked = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            rdoSelectManual = new RadioButton { Text = "Chọn điểm thủ công trên bản vẽ", Location = new Point(200, 30), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            grpSource.Controls.Add(rdoPointGroup);
            grpSource.Controls.Add(rdoSelectManual);

            rdoPointGroup.CheckedChanged += (s, e) => {
                cboPointGroup.Enabled = rdoPointGroup.Checked;
                chkRenameGroup.Enabled = rdoPointGroup.Checked;
            };

            grpSource.Controls.Add(CreateLabel("Point Group:", 15, 58));
            cboPointGroup = CreateComboBox(110, 55, 250);
            grpSource.Controls.Add(cboPointGroup);

            chkRenameGroup = CreateCheckBox("Đổi tên Point Group thành Hằng số", 15, 85);
            grpSource.Controls.Add(chkRenameGroup);

            y += 125;

            // ═══ 2. QUY TẮC ĐẶT TÊN ═══
            var grpRule = CreateGroupBox("Quy tắc đặt tên (Hằng số + STT)", 0, y, 700, 75);
            mainPanel.Controls.Add(grpRule);

            grpRule.Controls.Add(CreateLabel("Hằng số (Tiền tố):", 15, 32));
            txtPrefix = CreateTextBox(140, 30, 100, "P");
            grpRule.Controls.Add(txtPrefix);

            grpRule.Controls.Add(CreateLabel("STT bắt đầu:", 260, 32));
            numStartIdx = CreateNumeric(350, 30, 70, 1, 999999, 1);
            grpRule.Controls.Add(numStartIdx);

            grpRule.Controls.Add(CreateLabel("Bước nhảy:", 440, 32));
            numStep = CreateNumeric(520, 30, 70, 1, 1000, 1);
            grpRule.Controls.Add(numStep);

            y += 85;

            // ═══ 3. PREVIEW ═══
            var grpPreview = CreateGroupBox("Preview thay đổi", 0, y, 700, 300);
            mainPanel.Controls.Add(grpPreview);

            btnPreview = CreateButton("Tạo Preview", 15, 25, 120, 30, grpPreview);
            btnPreview.Click += BtnPreview_Click;

            dgvPreview = CreateDataGrid(15, 65, 670, 195);
            dgvPreview.Columns.Add("colNum", "Point Number");
            dgvPreview.Columns.Add("colOld", "Tên hiện tại");
            dgvPreview.Columns.Add("colNew", "Tên mới");
            grpPreview.Controls.Add(dgvPreview);

            lblTotal = CreateLabel("Tổng: 0 điểm", 15, 270);
            lblTotal.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            grpPreview.Controls.Add(lblTotal);

            // ═══ BOTTOM BAR ═══
            var pnlBottom = CreateBottomBar();
            this.Controls.Add(pnlBottom);

            var btnCancel = CreateButton("Hủy", 430, 8, 110, 36, pnlBottom);
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            var btnExecute = CreatePrimaryButton("Thực hiện", 550, 8, 140, 36, pnlBottom);
            btnExecute.Click += BtnExecute_Click;
        }

        private void LoadPointGroups()
        {
            try
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                if (doc == null) return;
                var civilDoc = CivilApplication.ActiveDocument;

                using var tr = doc.Database.TransactionManager.StartOpenCloseTransaction();
                cboPointGroup.Items.Clear();
                foreach (ObjectId pgId in civilDoc.PointGroups)
                {
                    var pg = tr.GetObject(pgId, OpenMode.ForRead) as PointGroup;
                    if (pg != null)
                    {
                        cboPointGroup.Items.Add(pg.Name);
                    }
                }
                if (cboPointGroup.Items.Count > 0)
                    cboPointGroup.SelectedIndex = 0;
                tr.Commit();
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
        }

        private void BtnPreview_Click(object? sender, EventArgs e)
        {
            if (IsByGroup && string.IsNullOrEmpty(SelectedGroupName))
            {
                MessageBox.Show("Vui lòng chọn Point Group!");
                return;
            }

            if (!IsByGroup)
            {
                this.Hide();
                var pso = new PromptSelectionOptions { MessageForAdding = "\nChọn các CogoPoint cần đổi tên: " };
                var filter = new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "AECC_COGO_POINT") });
                var selRes = AcadApp.DocumentManager.MdiActiveDocument.Editor.GetSelection(pso, filter);
                
                if (selRes.Status == PromptStatus.OK)
                {
                    SelectedPointIds = selRes.Value.GetObjectIds().ToList();
                }
                else
                {
                    SelectedPointIds.Clear();
                }
                this.Show();
                
                if (SelectedPointIds.Count == 0) return;
            }

            GeneratePreview();
        }

        private void GeneratePreview()
        {
            PreviewItems.Clear();
            dgvPreview.Rows.Clear();

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var civilDoc = CivilApplication.ActiveDocument;

            using var tr = doc.Database.TransactionManager.StartTransaction();
            
            List<ObjectId> targetIds = new();

            if (IsByGroup)
            {
                PointGroupCollection pgColl = civilDoc.PointGroups;
                if (pgColl.Contains(SelectedGroupName))
                {
                    ObjectId pgId = pgColl[SelectedGroupName];
                    // Dùng utility đã có để lấy danh sách ObjectId trong PointGroup
                    var coll = UtilitiesC3D.GPointIdsFromPointGroup(pgId);
                    foreach (ObjectId id in coll) targetIds.Add(id);
                }
            }
            else
            {
                targetIds = SelectedPointIds.ToList();
            }

            // Đọc điểm và sắp xếp theo PointNumber
            var points = new List<CogoPoint>();
            foreach (var id in targetIds)
            {
                var pt = tr.GetObject(id, OpenMode.ForRead) as CogoPoint;
                if (pt != null) points.Add(pt);
            }
            
            points = points.OrderBy(p => p.PointNumber).ToList();

            int currentNum = StartIndex;
            string prefix = Prefix;
            int step = Step;

            foreach (var pt in points)
            {
                string newName = $"{prefix}{currentNum}";
                PreviewItems.Add(new PreviewItem
                {
                    Id = pt.ObjectId,
                    PointNumber = pt.PointNumber,
                    OldName = pt.PointName ?? "",
                    NewName = newName
                });
                
                dgvPreview.Rows.Add(pt.PointNumber, pt.PointName ?? "", newName);
                currentNum += step;
            }

            lblTotal.Text = $"Tổng: {PreviewItems.Count} điểm";
            tr.Commit();
        }

        private void BtnExecute_Click(object? sender, EventArgs e)
        {
            if (PreviewItems.Count == 0)
            {
                MessageBox.Show("Vui lòng tạo Preview trước khi thực hiện!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(Prefix))
            {
                 var confirmPrefix = MessageBox.Show("Tiền tố (Hằng số) đang để trống. Tên điểm sẽ chỉ có số thứ tự. Bạn có chắc chắn?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                 if (confirmPrefix == DialogResult.No) return;
            }

            var confirm = MessageBox.Show($"Xác nhận đổi tên {PreviewItems.Count} điểm?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm == DialogResult.Yes)
            {
                IsConfirmed = true;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }
    }

    public class DoiTenPoint
    {
        [CommandMethod("CTPO_DoiTenPoint")]
        public static void CTPODoiTenPoint()
        {
            using var form = new DoiTenPointForm();
            var res = AcadApp.ShowModalDialog(form);

            if (res != DialogResult.OK || !form.IsConfirmed) return;

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            var civilDoc = CivilApplication.ActiveDocument;

            using var tr = db.TransactionManager.StartTransaction();
            try
            {
                int count = 0;
                foreach (var item in form.PreviewItems)
                {
                    var pt = tr.GetObject(item.Id, OpenMode.ForWrite) as CogoPoint;
                    if (pt != null)
                    {
                        pt.PointName = item.NewName;
                        count++;
                    }
                }

                // Đổi tên Point Group nếu được chọn
                if (form.RenameGroup && !string.IsNullOrWhiteSpace(form.Prefix))
                {
                    PointGroupCollection pgColl = civilDoc.PointGroups;
                    if (pgColl.Contains(form.SelectedGroupName))
                    {
                        ObjectId pgId = pgColl[form.SelectedGroupName];
                        var pg = tr.GetObject(pgId, OpenMode.ForWrite) as PointGroup;
                        if (pg != null)
                        {
                            // Tránh lỗi nếu tên Group mới đã tồn tại
                            string newGroupName = form.Prefix;
                            int suffix = 1;
                            while (pgColl.Contains(newGroupName) && newGroupName != form.SelectedGroupName)
                            {
                                newGroupName = $"{form.Prefix}_{suffix}";
                                suffix++;
                            }
                            pg.Name = newGroupName;
                            ed.WriteMessage($"\n✅ Đã đổi tên Point Group thành: {newGroupName}");
                        }
                    }
                }

                // Cập nhật lại Point Groups
                try
                {
                    civilDoc.PointGroups.UpdateAllPointGroups();
                }
                catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

                tr.Commit();
                ed.WriteMessage($"\n✅ Đã đổi tên thành công {count} CogoPoints.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi khi đổi tên: {ex.Message}");
            }
        }
    }
}
