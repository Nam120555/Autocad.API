// (C) Copyright 2024
// Tính khối lượng vật liệu từ Material Section và xuất ra Excel
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Acad = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using AcadDb = Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

using Civil = Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using Autodesk.Civil.Runtime;
using Autodesk.Civil.Settings;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Civil.ApplicationServices;
using CivSection = Autodesk.Civil.DatabaseServices.Section;
using FormLabel = System.Windows.Forms.Label;

using ClosedXML.Excel;
using MyFirstProject.Extensions;


// Tắt cảnh báo nullable cho file lớn
#pragma warning disable CS8600, CS8602, CS8618

// Các Form giao diện của module Tính khối lượng (tách từ 10.TinhKhoiLuongExcel.cs)
namespace MyFirstProject
{

    /// <summary>
    /// Form chọn Alignments để tính khối lượng
    /// </summary>
    public class FormChonAlignment : CivilFormBase
    {
        private CheckedListBox checkedListBox;
        private Button btnOK;
        private Button btnCancel;
        private Button btnSelectAll;
        private Button btnDeselectAll;
        private List<AlignmentInfo> alignmentInfos;

        public List<AlignmentInfo> SelectedAlignments => alignmentInfos.Where(a => a.IsSelected).ToList();

        public FormChonAlignment(List<AlignmentInfo> alignments)
        {
            alignmentInfos = alignments;
            InitializeComponent();
            LoadAlignments();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("Chọn Alignments để tính khối lượng", 450, 400);

            FormLabel lblTitle = CreateLabel("Chọn các Alignment có SampleLineGroup:", 12, 12);
            this.Controls.Add(lblTitle);

            checkedListBox = new CheckedListBox
            {
                Location = new System.Drawing.Point(12, 35),
                Size = new System.Drawing.Size(410, 250),
                BackColor = BgInput,
                ForeColor = TextDark,
                BorderStyle = BorderStyle.FixedSingle,
                CheckOnClick = true
            };
            checkedListBox.ItemCheck += CheckedListBox_ItemCheck;
            this.Controls.Add(checkedListBox);

            btnSelectAll = CreateButton("Chọn tất cả", 12, 295, 100, 28, this);
            btnSelectAll.Click += (s, e) => SelectAll(true);

            btnDeselectAll = CreateButton("Bỏ chọn tất cả", 120, 295, 120, 28, this);
            btnDeselectAll.Click += (s, e) => SelectAll(false);

            btnOK = CreatePrimaryButton("OK", 240, 330, 85, 28, this);
            btnOK.DialogResult = DialogResult.OK;

            btnCancel = CreateButton("Hủy", 335, 330, 85, 28, this);
            btnCancel.DialogResult = DialogResult.Cancel;

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void LoadAlignments()
        {
            foreach (var alignment in alignmentInfos)
            {
                checkedListBox.Items.Add(alignment.Name, alignment.IsSelected);
            }
        }

        private void CheckedListBox_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            alignmentInfos[e.Index].IsSelected = (e.NewValue == CheckState.Checked);
        }

        private void SelectAll(bool select)
        {
            for (int i = 0; i < checkedListBox.Items.Count; i++)
            {
                checkedListBox.SetItemChecked(i, select);
                alignmentInfos[i].IsSelected = select;
            }
        }
    }

    /// <summary>
    /// Form sắp xếp thứ tự vật liệu
    /// </summary>
    public class FormSapXepVatLieu : CivilFormBase
    {
        private ListView listViewMaterials;
        private Button btnMoveUp;
        private Button btnMoveDown;
        private Button btnOK;
        private Button btnCancel;
        private List<string> materials;

        public List<string> OrderedMaterials => materials;

        public FormSapXepVatLieu(List<string> materialNames)
        {
            materials = new List<string>(materialNames);
            InitializeComponent();
            LoadMaterials();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("Sắp xếp thứ tự vật liệu", 500, 450);

            FormLabel lblTitle = CreateLabel("Sắp xếp thứ tự các vật liệu (cột trong Excel):", 12, 12);
            this.Controls.Add(lblTitle);

            listViewMaterials = CreateListView(12, 35, 380, 320);
            listViewMaterials.View = View.Details;
            listViewMaterials.FullRowSelect = true;
            listViewMaterials.MultiSelect = false;
            listViewMaterials.HideSelection = false;

            listViewMaterials.Columns.Add("STT", 50);
            listViewMaterials.Columns.Add("Tên vật liệu", 300);
            this.Controls.Add(listViewMaterials);

            btnMoveUp = CreateButton("▲ Lên", 405, 100, 70, 35, this);
            btnMoveUp.Click += BtnMoveUp_Click;

            btnMoveDown = CreateButton("▼ Xuống", 405, 145, 70, 35, this);
            btnMoveDown.Click += BtnMoveDown_Click;

            btnOK = CreatePrimaryButton("OK", 290, 370, 85, 28, this);
            btnOK.DialogResult = DialogResult.OK;

            btnCancel = CreateButton("Hủy", 385, 370, 85, 28, this);
            btnCancel.DialogResult = DialogResult.Cancel;

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void LoadMaterials()
        {
            listViewMaterials.Items.Clear();
            for (int i = 0; i < materials.Count; i++)
            {
                var item = new ListViewItem((i + 1).ToString());
                item.SubItems.Add(materials[i]);
                listViewMaterials.Items.Add(item);
            }
            if (listViewMaterials.Items.Count > 0)
                listViewMaterials.Items[0].Selected = true;
        }

        private void BtnMoveUp_Click(object? sender, EventArgs e)
        {
            if (listViewMaterials.SelectedIndices.Count == 0) return;
            int index = listViewMaterials.SelectedIndices[0];
            if (index <= 0) return;

            // Swap in materials list
            (materials[index], materials[index - 1]) = (materials[index - 1], materials[index]);
            LoadMaterials();
            listViewMaterials.Items[index - 1].Selected = true;
        }

        private void BtnMoveDown_Click(object? sender, EventArgs e)
        {
            if (listViewMaterials.SelectedIndices.Count == 0) return;
            int index = listViewMaterials.SelectedIndices[0];
            if (index >= materials.Count - 1) return;

            // Swap in materials list
            (materials[index], materials[index + 1]) = (materials[index + 1], materials[index]);
            LoadMaterials();
            listViewMaterials.Items[index + 1].Selected = true;
        }
    }



    /// <summary>
    /// Form cài đặt bảng khối lượng với giao diện đồ họa
    /// </summary>
    public class TableSettingsForm : CivilFormBase
    {
        private NumericUpDown nudTextHeight = null!;
        private NumericUpDown nudRowHeight = null!;
        private NumericUpDown nudColWidth = null!;
        private NumericUpDown nudTableSpacingX = null!;
        private NumericUpDown nudDecimalPlaces = null!;
        private Button btnOK = null!;
        private Button btnCancel = null!;
        private Button btnReset = null!;

        // Kết quả được lưu ở đây (static để persist giữa các lần mở)
        public static double TextHeight { get; set; } = 3.0;
        public static double RowHeight { get; set; } = 5.0;
        public static double ColumnWidth { get; set; } = 20.0;
        public static double TableSpacingX { get; set; } = 50.0;
        public static int DecimalPlaces { get; set; } = 3;

        public TableSettingsForm()
        {
            InitializeComponent();
            LoadCurrentSettings();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("Cài đặt Bảng Khối Lượng", 370, 360);

            var header = CreateHeader("CÀI ĐẶT BẢNG KHỐI LƯỢNG");
            this.Controls.Add(header);

            int labelWidth = 150;
            int inputWidth = 100;
            int rowHeight = 35;
            int margin = 20;
            int y = 55;

            // 1. Chiều cao text
            var lblTextHeight = CreateLabel("Chiều cao chữ (mm):", margin, y + 3);
            this.Controls.Add(lblTextHeight);

            nudTextHeight = new NumericUpDown
            {
                Location = new System.Drawing.Point(margin + labelWidth + 10, y),
                Size = new System.Drawing.Size(inputWidth, 25),
                DecimalPlaces = 1,
                Minimum = 0.5M,
                Maximum = 50M,
                Increment = 0.5M,
                Value = 3.0M,
                BackColor = BgInput,
                ForeColor = TextDark
            };
            this.Controls.Add(nudTextHeight);
            y += rowHeight;

            // 2. Chiều cao hàng
            var lblRowHeight = CreateLabel("Chiều cao hàng (mm):", margin, y + 3);
            this.Controls.Add(lblRowHeight);

            nudRowHeight = new NumericUpDown
            {
                Location = new System.Drawing.Point(margin + labelWidth + 10, y),
                Size = new System.Drawing.Size(inputWidth, 25),
                DecimalPlaces = 1,
                Minimum = 1M,
                Maximum = 100M,
                Increment = 0.5M,
                Value = 5.0M,
                BackColor = BgInput,
                ForeColor = TextDark
            };
            this.Controls.Add(nudRowHeight);
            y += rowHeight;

            // 3. Chiều rộng cột
            var lblColWidth = CreateLabel("Chiều rộng cột (mm):", margin, y + 3);
            this.Controls.Add(lblColWidth);

            nudColWidth = new NumericUpDown
            {
                Location = new System.Drawing.Point(margin + labelWidth + 10, y),
                Size = new System.Drawing.Size(inputWidth, 25),
                DecimalPlaces = 1,
                Minimum = 5M,
                Maximum = 200M,
                Increment = 5M,
                Value = 20.0M,
                BackColor = BgInput,
                ForeColor = TextDark
            };
            this.Controls.Add(nudColWidth);
            y += rowHeight;

            // 4. Khoảng cách giữa các bảng
            var lblSpacing = CreateLabel("Khoảng cách bảng (mm):", margin, y + 3);
            this.Controls.Add(lblSpacing);

            nudTableSpacingX = new NumericUpDown
            {
                Location = new System.Drawing.Point(margin + labelWidth + 10, y),
                Size = new System.Drawing.Size(inputWidth, 25),
                DecimalPlaces = 0,
                Minimum = 0M,
                Maximum = 500M,
                Increment = 10M,
                Value = 50M,
                BackColor = BgInput,
                ForeColor = TextDark
            };
            this.Controls.Add(nudTableSpacingX);
            y += rowHeight;

            // 5. Số chữ số thập phân
            var lblDecimal = CreateLabel("Số chữ số thập phân:", margin, y + 3);
            this.Controls.Add(lblDecimal);

            nudDecimalPlaces = new NumericUpDown
            {
                Location = new System.Drawing.Point(margin + labelWidth + 10, y),
                Size = new System.Drawing.Size(inputWidth, 25),
                DecimalPlaces = 0,
                Minimum = 0M,
                Maximum = 6M,
                Increment = 1M,
                Value = 3M,
                BackColor = BgInput,
                ForeColor = TextDark
            };
            this.Controls.Add(nudDecimalPlaces);
            y += rowHeight + 10;

            // Buttons
            var pnlBottom = CreateBottomBar();
            this.Controls.Add(pnlBottom);

            btnReset = CreateButton("↺ Reset", 10, 8, 90, 36, pnlBottom);
            btnReset.Click += BtnReset_Click;

            btnCancel = CreateButton("Hủy", this.Width - 210, 8, 85, 36, pnlBottom);
            btnCancel.DialogResult = DialogResult.Cancel;

            btnOK = CreatePrimaryButton("Lưu", this.Width - 110, 8, 85, 36, pnlBottom);
            btnOK.DialogResult = DialogResult.OK;
            btnOK.Click += BtnOK_Click;

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void LoadCurrentSettings()
        {
            nudTextHeight.Value = (decimal)TextHeight;
            nudRowHeight.Value = (decimal)RowHeight;
            nudColWidth.Value = (decimal)ColumnWidth;
            nudTableSpacingX.Value = (decimal)TableSpacingX;
            nudDecimalPlaces.Value = DecimalPlaces;
        }

        private void BtnOK_Click(object? sender, EventArgs e)
        {
            // Lưu cài đặt
            TextHeight = (double)nudTextHeight.Value;
            RowHeight = (double)nudRowHeight.Value;
            ColumnWidth = (double)nudColWidth.Value;
            TableSpacingX = (double)nudTableSpacingX.Value;
            DecimalPlaces = (int)nudDecimalPlaces.Value;

            // Cập nhật biến static trong TinhKhoiLuongExcel
            TinhKhoiLuongExcel.SetTableSettings(TextHeight, TableSpacingX);

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnReset_Click(object? sender, EventArgs e)
        {
            // Reset về mặc định
            nudTextHeight.Value = 3.0M;
            nudRowHeight.Value = 5.0M;
            nudColWidth.Value = 20.0M;
            nudTableSpacingX.Value = 50M;
            nudDecimalPlaces.Value = 3M;
        }

        /// <summary>
        /// Mở form và trả về true nếu người dùng nhấn OK
        /// </summary>
        public static bool ShowSettings()
        {
            using var form = new TableSettingsForm();
            return form.ShowDialog() == DialogResult.OK;
        }
    }

    /// <summary>
    /// Form Taskbar nhỏ gọn cho các chức năng tính khối lượng
    /// </summary>
    public class VolumeTaskbar : CivilFormBase
    {
        private Button btnSettings = null!;
        private Button btnExportExcel = null!;
        private Button btnExportCad = null!;
        private Button btnClose = null!;

        // Biến static để lưu instance duy nhất
        private static VolumeTaskbar? _instance;

        public static void ShowTaskbar()
        {
            if (_instance == null || _instance.IsDisposed)
            {
                _instance = new VolumeTaskbar();
            }
            _instance.Show();
            _instance.BringToFront();
        }

        public VolumeTaskbar()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("Khối Lượng", 180, 180);
            this.TopMost = true;
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new System.Drawing.Point(100, 100);
            this.ShowInTaskbar = false;

            var header = CreateHeader("KHỐI LƯỢNG", 30);
            this.Controls.Add(header);

            int btnWidth = 150;
            int btnHeight = 32;
            int margin = 15;
            int y = 40;

            // Button Cài đặt
            btnSettings = CreateButton("⚙ Cài đặt", margin, y, btnWidth, btnHeight, this);
            btnSettings.Click += (s, e) => { TableSettingsForm.ShowSettings(); };
            y += btnHeight + 5;

            // Button Xuất Excel
            btnExportExcel = CreatePrimaryButton("📊 Xuất Excel + CAD", margin, y, btnWidth, btnHeight, this);
            btnExportExcel.Click += (s, e) => { this.Hide(); TinhKhoiLuongExcel._REMOVED_CTSVXuatKhoiLuong(); this.Show(); };
            y += btnHeight + 5;

            // Button Xuất CAD riêng
            btnExportCad = CreateButton("📐 Chỉ xuất CAD", margin, y, btnWidth, btnHeight, this);
            btnExportCad.Click += (s, e) => { this.Hide(); TinhKhoiLuongExcel.CTSVXuatCadOnly(); this.Show(); };
            y += btnHeight + 5;

            // Button Đóng
            btnClose = CreateButton("✕ Đóng", margin, y, btnWidth, btnHeight, this);
            btnClose.Click += (s, e) => this.Close();
        }
    }
}
