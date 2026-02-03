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

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(Civil3DCsharp.TinhKhoiLuongExcel))]

namespace Civil3DCsharp
{
    #region Data Classes

    /// <summary>
    /// Thông tin diện tích vật liệu tại một cọc
    /// </summary>
    public class MaterialVolumeInfo
    {
        public string StakeName { get; set; } = "";       // Tên cọc: "Cọc 1"
        public string Station { get; set; } = "";          // Lý trình: "Km0+123.456"
        public double StationValue { get; set; }          // Giá trị lý trình: 123.456
        public string MaterialName { get; set; } = "";    // Tên vật liệu: "Đào đất"
        public double Area { get; set; }                  // Diện tích (m²)
    }

    /// <summary>
    /// Thông tin tổng hợp của một cọc
    /// </summary>
    public class StakeInfo
    {
        public string Station { get; set; } = "";         // Lý trình format: "Km0+123.456"
        public string StakeName { get; set; } = "";       // Tên cọc
        public double StationValue { get; set; }          // Giá trị lý trình số
        public double SpacingPrev { get; set; }           // Khoảng cách đến cọc trước
        
        // Dữ liệu theo nhóm
        public Dictionary<string, double> MaterialAreas { get; set; } = new();   // Vật liệu từ QTO
        public Dictionary<string, double> CorridorAreas { get; set; } = new();   // Shape từ Corridor
        public Dictionary<string, double> SurfaceAreas { get; set; } = new();    // Mặt địa hình
        public Dictionary<string, double> OtherAreas { get; set; } = new();      // Các nguồn khác
        
        // Khối lượng theo nhóm
        public Dictionary<string, double> MaterialVolumes { get; set; } = new();
        public Dictionary<string, double> CorridorVolumes { get; set; } = new();
        public Dictionary<string, double> SurfaceVolumes { get; set; } = new();
        public Dictionary<string, double> OtherVolumes { get; set; } = new();
        
        /// <summary>
        /// Lấy tất cả diện tích (gộp tất cả nhóm)
        /// </summary>
        public Dictionary<string, double> GetAllAreas()
        {
            var all = new Dictionary<string, double>(MaterialAreas);
            foreach (var kvp in CorridorAreas) all[kvp.Key] = kvp.Value;
            foreach (var kvp in SurfaceAreas) all[kvp.Key] = kvp.Value;
            foreach (var kvp in OtherAreas) all[kvp.Key] = kvp.Value;
            return all;
        }
        
        /// <summary>
        /// Lấy tất cả khối lượng (gộp tất cả nhóm)
        /// </summary>
        public Dictionary<string, double> GetAllVolumes()
        {
            var all = new Dictionary<string, double>(MaterialVolumes);
            foreach (var kvp in CorridorVolumes) all[kvp.Key] = kvp.Value;
            foreach (var kvp in SurfaceVolumes) all[kvp.Key] = kvp.Value;
            foreach (var kvp in OtherVolumes) all[kvp.Key] = kvp.Value;
            return all;
        }
        
        // Chi tiết Material Section Data (Left/Right Length, Min/Max Elevation, etc.)
        public Dictionary<string, MaterialSectionDetail> MaterialSectionDetails { get; set; } = new();
    }
    
    /// <summary>
    /// Chi tiết thông tin của Material Section (như trong Properties Panel)
    /// </summary>
    public class MaterialSectionDetail
    {
        public string MaterialName { get; set; } = "";
        public string SectionSurfaceName { get; set; } = "";
        
        // Phạm vi offset
        public double LeftLength { get; set; }          // Giá trị âm (bên trái tim)
        public double RightLength { get; set; }         // Giá trị dương (bên phải tim)
        public double TotalWidth => Math.Abs(LeftLength) + Math.Abs(RightLength);
        
        // Phạm vi cao độ
        public double MinElevation { get; set; }
        public double MaxElevation { get; set; }
        public double Height => MaxElevation - MinElevation;
        
        // Diện tích và điểm
        public double Area { get; set; }
        public int PointCount { get; set; }
        public List<Point3d> Points { get; set; } = new();
    }

    /// <summary>
    /// Thông tin Alignment có SampleLineGroup
    /// </summary>
    public class AlignmentInfo
    {
        public ObjectId AlignmentId { get; set; }
        public string Name { get; set; } = "";
        public ObjectId SampleLineGroupId { get; set; }
        public bool IsSelected { get; set; }
    }

    #endregion

    #region Forms

    /// <summary>
    /// Form chọn Alignments để tính khối lượng
    /// </summary>
    public class FormChonAlignment : Form
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
            this.Text = "Chọn Alignments để tính khối lượng";
            this.Size = new System.Drawing.Size(450, 400);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            FormLabel lblTitle = new()
            {
                Text = "Chọn các Alignment có SampleLineGroup:",
                Location = new System.Drawing.Point(12, 12),
                Size = new System.Drawing.Size(400, 20)
            };
            this.Controls.Add(lblTitle);

            checkedListBox = new CheckedListBox
            {
                Location = new System.Drawing.Point(12, 35),
                Size = new System.Drawing.Size(410, 250),
                CheckOnClick = true
            };
            checkedListBox.ItemCheck += CheckedListBox_ItemCheck;
            this.Controls.Add(checkedListBox);

            btnSelectAll = new Button
            {
                Text = "Chọn tất cả",
                Location = new System.Drawing.Point(12, 295),
                Size = new System.Drawing.Size(100, 28)
            };
            btnSelectAll.Click += (s, e) => SelectAll(true);
            this.Controls.Add(btnSelectAll);

            btnDeselectAll = new Button
            {
                Text = "Bỏ chọn tất cả",
                Location = new System.Drawing.Point(120, 295),
                Size = new System.Drawing.Size(100, 28)
            };
            btnDeselectAll.Click += (s, e) => SelectAll(false);
            this.Controls.Add(btnDeselectAll);

            btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(240, 330),
                Size = new System.Drawing.Size(85, 28),
                DialogResult = DialogResult.OK
            };
            this.Controls.Add(btnOK);

            btnCancel = new Button
            {
                Text = "Hủy",
                Location = new System.Drawing.Point(335, 330),
                Size = new System.Drawing.Size(85, 28),
                DialogResult = DialogResult.Cancel
            };
            this.Controls.Add(btnCancel);

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
    public class FormSapXepVatLieu : Form
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
            this.Text = "Sắp xếp thứ tự vật liệu";
            this.Size = new System.Drawing.Size(500, 450);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            FormLabel lblTitle = new()
            {
                Text = "Sắp xếp thứ tự các vật liệu (cột trong Excel):",
                Location = new System.Drawing.Point(12, 12),
                Size = new System.Drawing.Size(450, 20)
            };
            this.Controls.Add(lblTitle);

            listViewMaterials = new ListView
            {
                Location = new System.Drawing.Point(12, 35),
                Size = new System.Drawing.Size(380, 320),
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                HideSelection = false
            };
            listViewMaterials.Columns.Add("STT", 50);
            listViewMaterials.Columns.Add("Tên vật liệu", 300);
            this.Controls.Add(listViewMaterials);

            btnMoveUp = new Button
            {
                Text = "▲ Lên",
                Location = new System.Drawing.Point(405, 100),
                Size = new System.Drawing.Size(70, 35)
            };
            btnMoveUp.Click += BtnMoveUp_Click;
            this.Controls.Add(btnMoveUp);

            btnMoveDown = new Button
            {
                Text = "▼ Xuống",
                Location = new System.Drawing.Point(405, 145),
                Size = new System.Drawing.Size(70, 35)
            };
            btnMoveDown.Click += BtnMoveDown_Click;
            this.Controls.Add(btnMoveDown);

            btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(290, 370),
                Size = new System.Drawing.Size(85, 28),
                DialogResult = DialogResult.OK
            };
            this.Controls.Add(btnOK);

            btnCancel = new Button
            {
                Text = "Hủy",
                Location = new System.Drawing.Point(385, 370),
                Size = new System.Drawing.Size(85, 28),
                DialogResult = DialogResult.Cancel
            };
            this.Controls.Add(btnCancel);

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

    #endregion

    #region Forms

    /// <summary>
    /// Form cài đặt bảng khối lượng với giao diện đồ họa
    /// </summary>
    public class TableSettingsForm : Form
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
            this.Text = "Cài đặt Bảng Khối Lượng";
            this.Size = new System.Drawing.Size(350, 320);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = System.Drawing.Color.FromArgb(45, 45, 48);
            this.ForeColor = System.Drawing.Color.White;

            int labelWidth = 150;
            int inputWidth = 100;
            int rowHeight = 35;
            int margin = 20;
            int y = 20;

            // Title
            var lblTitle = new FormLabel
            {
                Text = "⚙ CÀI ĐẶT BẢNG KHỐI LƯỢNG",
                Location = new System.Drawing.Point(margin, y),
                Size = new System.Drawing.Size(300, 25),
                Font = new System.Drawing.Font("Segoe UI", 11, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(0, 122, 204)
            };
            this.Controls.Add(lblTitle);
            y += 35;

            // 1. Chiều cao text
            var lblTextHeight = new FormLabel
            {
                Text = "Chiều cao chữ (mm):",
                Location = new System.Drawing.Point(margin, y + 3),
                Size = new System.Drawing.Size(labelWidth, 20),
                ForeColor = System.Drawing.Color.White
            };
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
                BackColor = System.Drawing.Color.FromArgb(60, 60, 65),
                ForeColor = System.Drawing.Color.White
            };
            this.Controls.Add(nudTextHeight);
            y += rowHeight;

            // 2. Chiều cao hàng
            var lblRowHeight = new FormLabel
            {
                Text = "Chiều cao hàng (mm):",
                Location = new System.Drawing.Point(margin, y + 3),
                Size = new System.Drawing.Size(labelWidth, 20),
                ForeColor = System.Drawing.Color.White
            };
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
                BackColor = System.Drawing.Color.FromArgb(60, 60, 65),
                ForeColor = System.Drawing.Color.White
            };
            this.Controls.Add(nudRowHeight);
            y += rowHeight;

            // 3. Chiều rộng cột
            var lblColWidth = new FormLabel
            {
                Text = "Chiều rộng cột (mm):",
                Location = new System.Drawing.Point(margin, y + 3),
                Size = new System.Drawing.Size(labelWidth, 20),
                ForeColor = System.Drawing.Color.White
            };
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
                BackColor = System.Drawing.Color.FromArgb(60, 60, 65),
                ForeColor = System.Drawing.Color.White
            };
            this.Controls.Add(nudColWidth);
            y += rowHeight;

            // 4. Khoảng cách giữa các bảng
            var lblSpacing = new FormLabel
            {
                Text = "Khoảng cách bảng (mm):",
                Location = new System.Drawing.Point(margin, y + 3),
                Size = new System.Drawing.Size(labelWidth, 20),
                ForeColor = System.Drawing.Color.White
            };
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
                BackColor = System.Drawing.Color.FromArgb(60, 60, 65),
                ForeColor = System.Drawing.Color.White
            };
            this.Controls.Add(nudTableSpacingX);
            y += rowHeight;

            // 5. Số chữ số thập phân
            var lblDecimal = new FormLabel
            {
                Text = "Số chữ số thập phân:",
                Location = new System.Drawing.Point(margin, y + 3),
                Size = new System.Drawing.Size(labelWidth, 20),
                ForeColor = System.Drawing.Color.White
            };
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
                BackColor = System.Drawing.Color.FromArgb(60, 60, 65),
                ForeColor = System.Drawing.Color.White
            };
            this.Controls.Add(nudDecimalPlaces);
            y += rowHeight + 10;

            // Buttons
            int btnWidth = 80;
            int btnHeight = 30;
            int btnSpacing = 10;
            int totalBtnWidth = btnWidth * 3 + btnSpacing * 2;
            int btnStartX = (this.ClientSize.Width - totalBtnWidth) / 2;

            btnOK = new Button
            {
                Text = "✓ Lưu",
                Location = new System.Drawing.Point(btnStartX, y),
                Size = new System.Drawing.Size(btnWidth, btnHeight),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                ForeColor = System.Drawing.Color.White,
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.OK
            };
            btnOK.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(0, 100, 180);
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            btnReset = new Button
            {
                Text = "↺ Reset",
                Location = new System.Drawing.Point(btnStartX + btnWidth + btnSpacing, y),
                Size = new System.Drawing.Size(btnWidth, btnHeight),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(200, 150, 50),
                ForeColor = System.Drawing.Color.White,
                Cursor = Cursors.Hand
            };
            btnReset.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(180, 130, 40);
            btnReset.Click += BtnReset_Click;
            this.Controls.Add(btnReset);

            btnCancel = new Button
            {
                Text = "✕ Hủy",
                Location = new System.Drawing.Point(btnStartX + (btnWidth + btnSpacing) * 2, y),
                Size = new System.Drawing.Size(btnWidth, btnHeight),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(100, 100, 105),
                ForeColor = System.Drawing.Color.White,
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(80, 80, 85);
            this.Controls.Add(btnCancel);

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
    public class VolumeTaskbar : Form
    {
        private Button btnSettings = null!;
        private Button btnExportExcel = null!;
        private Button btnExportCad = null!;
        private Button btnClose = null!;
        private FormLabel lblTitle = null!;

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
            // Form settings
            this.Text = "Khối Lượng";
            this.Size = new System.Drawing.Size(180, 200);
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.TopMost = true;
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new System.Drawing.Point(100, 100);
            this.ShowInTaskbar = false;
            this.BackColor = System.Drawing.Color.FromArgb(45, 45, 48);

            int btnWidth = 150;
            int btnHeight = 32;
            int margin = 10;
            int y = 10;

            // Title
            lblTitle = new FormLabel
            {
                Text = "⚡ KHỐI LƯỢNG",
                Location = new System.Drawing.Point(margin, y),
                Size = new System.Drawing.Size(btnWidth, 20),
                ForeColor = System.Drawing.Color.White,
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblTitle);
            y += 25;

            // Button Cài đặt
            btnSettings = new Button
            {
                Text = "⚙ Cài đặt",
                Location = new System.Drawing.Point(margin, y),
                Size = new System.Drawing.Size(btnWidth, btnHeight),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(60, 60, 65),
                ForeColor = System.Drawing.Color.White,
                Cursor = Cursors.Hand
            };
            btnSettings.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(80, 80, 85);
            btnSettings.Click += (s, e) => { TableSettingsForm.ShowSettings(); };
            this.Controls.Add(btnSettings);
            y += btnHeight + 5;

            // Button Xuất Excel
            btnExportExcel = new Button
            {
                Text = "📊 Xuất Excel + CAD",
                Location = new System.Drawing.Point(margin, y),
                Size = new System.Drawing.Size(btnWidth, btnHeight),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                ForeColor = System.Drawing.Color.White,
                Cursor = Cursors.Hand
            };
            btnExportExcel.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(0, 100, 180);
            btnExportExcel.Click += (s, e) => { this.Hide(); TinhKhoiLuongExcel.CTSVXuatKhoiLuong(); this.Show(); };
            this.Controls.Add(btnExportExcel);
            y += btnHeight + 5;

            // Button Xuất CAD riêng
            btnExportCad = new Button
            {
                Text = "📐 Chỉ xuất CAD",
                Location = new System.Drawing.Point(margin, y),
                Size = new System.Drawing.Size(btnWidth, btnHeight),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(40, 167, 69),
                ForeColor = System.Drawing.Color.White,
                Cursor = Cursors.Hand
            };
            btnExportCad.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(30, 140, 55);
            btnExportCad.Click += (s, e) => { this.Hide(); TinhKhoiLuongExcel.CTSVXuatCadOnly(); this.Show(); };
            this.Controls.Add(btnExportCad);
            y += btnHeight + 5;

            // Button Đóng
            btnClose = new Button
            {
                Text = "✕ Đóng",
                Location = new System.Drawing.Point(margin, y),
                Size = new System.Drawing.Size(btnWidth, btnHeight),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(100, 100, 105),
                ForeColor = System.Drawing.Color.White,
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(80, 80, 85);
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);
        }
    }

    #endregion

    public class TinhKhoiLuongExcel
    {
        #region Hằng số làm tròn (theo V3Tools: roundnb = 2)
        
        /// <summary>
        /// Số chữ số thập phân để làm tròn diện tích (m²)
        /// V3Tools sử dụng roundnb = 2
        /// </summary>
        public static int AreaDecimalPlaces { get; set; } = 2;
        
        /// <summary>
        /// Số chữ số thập phân để làm tròn khối lượng (m³)
        /// V3Tools sử dụng roundnb = 2
        /// </summary>
        public static int VolumeDecimalPlaces { get; set; } = 2;
        
        #endregion

        #region Công thức tính toán

        /// <summary>
        /// Tính diện tích đa giác sử dụng Shoelace Formula
        /// </summary>
        public static double CalculatePolygonArea(List<Point2d> points)
        {
            if (points.Count < 3) return 0;

            double area = 0;
            int n = points.Count;
            for (int i = 0; i < n; i++)
            {
                int j = (i + 1) % n;
                area += points[i].X * points[j].Y;
                area -= points[j].X * points[i].Y;
            }
            return Math.Abs(area) / 2.0;
        }

        /// <summary>
        /// Tính khối lượng bằng phương pháp trung bình cộng (Average End Area)
        /// Volume = (S1 + S2) / 2 × L
        /// Công thức chuẩn Civil 3D - như V3Tools sử dụng
        /// </summary>
        /// <param name="areaPrev">Diện tích trắc ngang trước (m²)</param>
        /// <param name="areaCurrent">Diện tích trắc ngang hiện tại (m²)</param>
        /// <param name="spacing">Khoảng cách giữa 2 trắc ngang (m)</param>
        /// <param name="round">Có làm tròn hay không</param>
        /// <returns>Khối lượng (m³)</returns>
        public static double CalculateVolume(double areaPrev, double areaCurrent, double spacing, bool round = true)
        {
            double volume = (areaPrev + areaCurrent) / 2.0 * spacing;
            return round ? Math.Round(volume, VolumeDecimalPlaces) : volume;
        }
        
        /// <summary>
        /// Làm tròn diện tích theo cài đặt
        /// </summary>
        public static double RoundArea(double area)
        {
            return Math.Round(area, AreaDecimalPlaces);
        }

        /// <summary>
        /// Format lý trình theo dạng Km0+123.456
        /// </summary>
        public static string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double meters = station % 1000;
            return $"Km{km}+{meters:F3}";
        }

        #endregion

        #region Settings

        /// <summary>
        /// Khoảng cách giữa các bảng khi xuất ra CAD (theo trục X)
        /// </summary>
        private static double TableSpacingX = 50.0;

        /// <summary>
        /// Chiều cao text mặc định trong bảng CAD
        /// </summary>
        private static double TableTextHeight = 3.0;

        /// <summary>
        /// Phương thức để TableSettingsForm cập nhật settings
        /// </summary>
        public static void SetTableSettings(double textHeight, double tableSpacing)
        {
            TableTextHeight = textHeight;
            TableSpacingX = tableSpacing;
        }

        /// <summary>
        /// Lệnh cài đặt khoảng cách giữa các bảng khi xuất CAD
        /// </summary>
        [CommandMethod("CTSV_CaiDatBang")]
        public static void CTSVCaiDatBang()
        {
            A.Ed.WriteMessage($"\n=== CÀI ĐẶT BẢNG KHỐI LƯỢNG ===");
            A.Ed.WriteMessage($"\nKhoảng cách hiện tại giữa các bảng (theo X): {TableSpacingX}");
            A.Ed.WriteMessage($"\nChiều cao text hiện tại: {TableTextHeight}");

            // Hỏi khoảng cách mới
            PromptDoubleOptions pdo = new("\nNhập khoảng cách giữa các bảng (theo X):")
            {
                AllowNegative = false,
                AllowZero = false,
                DefaultValue = TableSpacingX
            };
            PromptDoubleResult pdr = A.Ed.GetDouble(pdo);
            if (pdr.Status == PromptStatus.OK)
            {
                TableSpacingX = pdr.Value;
                A.Ed.WriteMessage($"\nĐã đặt khoảng cách giữa các bảng (theo X): {TableSpacingX}");
            }

            // Hỏi chiều cao text
            PromptDoubleOptions pdoText = new("\nNhập chiều cao text trong bảng:")
            {
                AllowNegative = false,
                AllowZero = false,
                DefaultValue = TableTextHeight
            };
            PromptDoubleResult pdrText = A.Ed.GetDouble(pdoText);
            if (pdrText.Status == PromptStatus.OK)
            {
                TableTextHeight = pdrText.Value;
                A.Ed.WriteMessage($"\nĐã đặt chiều cao text: {TableTextHeight}");
            }

            A.Ed.WriteMessage($"\n=== CÀI ĐẶT HOÀN TẤT ===");
        }

        /// <summary>
        /// Mở Form Taskbar nhỏ gọn
        /// </summary>
        [CommandMethod("CTSV_Taskbar")]
        public static void CTSVTaskbar()
        {
            VolumeTaskbar.ShowTaskbar();
        }

        /// <summary>
        /// Chỉ xuất bảng ra CAD (không xuất Excel)
        /// </summary>
        [CommandMethod("CTSV_XuatCad")]
        public static void CTSVXuatCadOnly()
        {
            try
            {
                A.Ed.WriteMessage("\n=== XUẤT BẢNG KHỐI LƯỢNG RA CAD ===");

                using Transaction tr = A.Db.TransactionManager.StartTransaction();

                // 1. Lấy danh sách Alignment có SampleLineGroup
                List<AlignmentInfo> alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông tìm thấy Alignment nào có SampleLineGroup!");
                    return;
                }

                // 2. Hiển thị form chọn Alignments
                FormChonAlignment formChon = new(alignments);
                if (formChon.ShowDialog() != DialogResult.OK || formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }

                // 3. Trích xuất dữ liệu
                Dictionary<ObjectId, List<StakeInfo>> alignmentData = new();
                HashSet<string> allMaterials = new();

                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    var stakeInfos = ExtractMaterialData(tr, alignInfo.SampleLineGroupId, alignInfo.AlignmentId);
                    alignmentData[alignInfo.AlignmentId] = stakeInfos;
                    
                    foreach (var stake in stakeInfos)
                    {
                        foreach (var mat in stake.MaterialAreas.Keys)
                            allMaterials.Add(mat);
                    }
                }

                if (allMaterials.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông tìm thấy dữ liệu vật liệu nào!");
                    return;
                }

                // 4. Sắp xếp vật liệu
                FormSapXepVatLieu formVatLieu = new(allMaterials.ToList());
                if (formVatLieu.ShowDialog() != DialogResult.OK)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }
                List<string> orderedMaterials = formVatLieu.OrderedMaterials;

                // 5. Tính khối lượng
                foreach (var kvp in alignmentData)
                {
                    CalculateVolumes(kvp.Value, orderedMaterials);
                }

                // 6. Vẽ bảng trong CAD
                PromptPointOptions ppo = new($"\nChọn điểm chèn bảng (các bảng tiếp theo cách nhau {TableSpacingX} đơn vị theo X):");
                ppo.AllowNone = false;
                PromptPointResult ppr = A.Ed.GetPoint(ppo);
                
                if (ppr.Status == PromptStatus.OK)
                {
                    Point3d currentInsertPoint = ppr.Value;
                    
                    foreach (var alignInfo in formChon.SelectedAlignments)
                    {
                        var stakeInfos = alignmentData[alignInfo.AlignmentId];
                        
                        int numCols = 3 + orderedMaterials.Count * 2;
                        double tableWidth = 25.0 + 25.0 + 15.0 + (numCols - 3) * 18.0;
                        
                        CreateCadTable(tr, currentInsertPoint, alignInfo.Name, stakeInfos, orderedMaterials);
                        A.Ed.WriteMessage($"\nĐã vẽ bảng cho '{alignInfo.Name}' tại ({currentInsertPoint.X:F2}, {currentInsertPoint.Y:F2})");
                        
                        currentInsertPoint = new Point3d(
                            currentInsertPoint.X + tableWidth + TableSpacingX, 
                            currentInsertPoint.Y, 
                            currentInsertPoint.Z);
                    }
                }
                
                tr.Commit();
                A.Ed.WriteMessage("\n=== HOÀN TẤT XUẤT CAD ===");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nLỗi: {ex.Message}");
            }
        }

        #endregion

        #region Volume Surface Comparison - So sánh bề mặt

        /// <summary>
        /// Thông tin khối lượng từ Volume Surface
        /// </summary>
        public class VolumeSurfaceData
        {
            public string SurfaceName { get; set; } = "";
            public string BaseSurfaceName { get; set; } = "";
            public string ComparisonSurfaceName { get; set; } = "";
            public double CutVolume { get; set; }
            public double FillVolume { get; set; }
            public double NetVolume { get; set; }
            public double CutArea { get; set; }
            public double FillArea { get; set; }
        }

        /// <summary>
        /// Thông tin diện tích từ SectionView
        /// </summary>
        public class SectionViewAreaData
        {
            public string StakeName { get; set; } = "";
            public double Station { get; set; }
            public Dictionary<string, double> MaterialAreas { get; set; } = new();
        }

        /// <summary>
        /// Thông tin dữ liệu từ Volume Table trong Section View
        /// </summary>
        public class VolumeTableDataInfo
        {
            public string MaterialName { get; set; } = "";      // Tên material (ví dụ: "Đào nền", "Đắp nền")
            public double Area { get; set; }                     // Diện tích tại trắc ngang (m²)
            public double SegmentVolume { get; set; }            // Khối lượng đoạn (từ trắc ngang trước đến hiện tại) (m³)
            public double CumulativeVolume { get; set; }         // Khối lượng tích lũy (m³)
            public string VolumeType { get; set; } = "";         // Cut/Fill/Material
        }

        /// <summary>
        /// Thông tin trắc ngang với dữ liệu từ Volume Table
        /// </summary>
        public class CrossSectionVolumeData
        {
            public string SampleLineName { get; set; } = "";
            public double Station { get; set; }
            public string StationFormatted { get; set; } = "";
            public double SpacingPrev { get; set; }              // Khoảng cách đến trắc ngang trước
            public Dictionary<string, VolumeTableDataInfo> Materials { get; set; } = new();
        }

        /// <summary>
        /// Lệnh so sánh 2 Surface để tính khối lượng đào đắp
        /// </summary>
        [CommandMethod("CTSV_SoSanhSurface")]
        public static void CTSVSoSanhSurface()
        {
            try
            {
                A.Ed.WriteMessage("\n\n=== SO SÁNH BỀ MẶT TÍNH KHỐI LƯỢNG ===");

                // Lấy danh sách Surface
                var surfaceIds = A.Cdoc.GetSurfaceIds();
                if (surfaceIds.Count < 2)
                {
                    A.Ed.WriteMessage("\n❌ Cần ít nhất 2 Surface để so sánh.");
                    return;
                }

                using Transaction tr = A.Db.TransactionManager.StartTransaction();

                // Liệt kê các Surface
                var surfaces = new List<(int Index, ObjectId Id, string Name, bool IsTin)>();
                int idx = 1;
                A.Ed.WriteMessage("\n\nDanh sách Surface có sẵn:");

                foreach (ObjectId id in surfaceIds)
                {
                    var surf = tr.GetObject(id, AcadDb.OpenMode.ForRead);
                    if (surf is TinSurface tinSurf)
                    {
                        A.Ed.WriteMessage($"\n  {idx}. {tinSurf.Name} (TIN Surface)");
                        surfaces.Add((idx, id, tinSurf.Name, true));
                        idx++;
                    }
                    else if (surf is TinVolumeSurface volSurf)
                    {
                        A.Ed.WriteMessage($"\n  {idx}. {volSurf.Name} (Volume Surface)");
                        surfaces.Add((idx, id, volSurf.Name, false));
                        idx++;
                    }
                }

                // Chọn Surface tự nhiên (Base)
                var baseResult = A.Ed.GetInteger($"\nChọn Surface TỰ NHIÊN (1-{surfaces.Count}): ");
                if (baseResult.Status != PromptStatus.OK) { tr.Commit(); return; }
                if (baseResult.Value < 1 || baseResult.Value > surfaces.Count) { tr.Commit(); return; }
                var baseSurface = surfaces[baseResult.Value - 1];

                // Chọn Surface thiết kế (Comparison)
                var compResult = A.Ed.GetInteger($"\nChọn Surface THIẾT KẾ (1-{surfaces.Count}): ");
                if (compResult.Status != PromptStatus.OK) { tr.Commit(); return; }
                if (compResult.Value < 1 || compResult.Value > surfaces.Count) { tr.Commit(); return; }
                var compSurface = surfaces[compResult.Value - 1];

                if (baseSurface.Id == compSurface.Id)
                {
                    A.Ed.WriteMessage("\n❌ Phải chọn 2 Surface khác nhau.");
                    tr.Commit();
                    return;
                }

                // Tạo hoặc lấy Volume Surface
                string volSurfName = $"VOL_{baseSurface.Name}_{compSurface.Name}";
                TinVolumeSurface? volumeSurface = null;

                // Kiểm tra xem đã có Volume Surface chưa
                foreach (ObjectId sid in surfaceIds)
                {
                    var s = tr.GetObject(sid, AcadDb.OpenMode.ForRead);
                    if (s is TinVolumeSurface tvs && tvs.Name == volSurfName)
                    {
                        volumeSurface = tvs;
                        A.Ed.WriteMessage($"\n✅ Sử dụng Volume Surface có sẵn: {volSurfName}");
                        break;
                    }
                }

                // Nếu chưa có, tạo mới
                if (volumeSurface == null && baseSurface.IsTin && compSurface.IsTin)
                {
                    try
                    {
                        var volSurfId = TinVolumeSurface.Create(volSurfName, baseSurface.Id, compSurface.Id);
                        volumeSurface = tr.GetObject(volSurfId, AcadDb.OpenMode.ForRead) as TinVolumeSurface;
                        A.Ed.WriteMessage($"\n✅ Đã tạo Volume Surface mới: {volSurfName}");
                    }
                    catch (System.Exception ex)
                    {
                        A.Ed.WriteMessage($"\n⚠️ Không thể tạo Volume Surface: {ex.Message}");
                    }
                }

                // Lấy thông tin khối lượng
                if (volumeSurface != null)
                {
                    var props = volumeSurface.GetVolumeProperties();

                    double cutVol = props.UnadjustedCutVolume;
                    double fillVol = props.UnadjustedFillVolume;
                    double netVol = cutVol - fillVol;

                    A.Ed.WriteMessage($"\n\n{'=',-60}");
                    A.Ed.WriteMessage($"\n📊 KẾT QUẢ SO SÁNH BỀ MẶT");
                    A.Ed.WriteMessage($"\n{'=',-60}");
                    A.Ed.WriteMessage($"\n  Surface tự nhiên: {baseSurface.Name}");
                    A.Ed.WriteMessage($"\n  Surface thiết kế: {compSurface.Name}");
                    A.Ed.WriteMessage($"\n{'-',-60}");
                    A.Ed.WriteMessage($"\n  Khối lượng ĐÀO (Cut):  {cutVol,15:N2} m³");
                    A.Ed.WriteMessage($"\n  Khối lượng ĐẮP (Fill): {fillVol,15:N2} m³");
                    A.Ed.WriteMessage($"\n  Khối lượng RÒNG (Net): {netVol,15:N2} m³");
                    A.Ed.WriteMessage($"\n{'=',-60}");

                    // Hỏi có muốn copy ra clipboard không
                    var copyResult = A.Ed.GetKeywords("\nCopy kết quả ra clipboard? [Yes/No] <Yes>: ", new[] { "Yes", "No" });
                    if (copyResult.Status != PromptStatus.OK || copyResult.StringResult != "No")
                    {
                        string clipboardText = $"SO SÁNH BỀ MẶT\n" +
                            $"Surface tự nhiên: {baseSurface.Name}\n" +
                            $"Surface thiết kế: {compSurface.Name}\n" +
                            $"Khối lượng ĐÀO (Cut): {cutVol:N2} m³\n" +
                            $"Khối lượng ĐẮP (Fill): {fillVol:N2} m³\n" +
                            $"Khối lượng RÒNG (Net): {netVol:N2} m³";
                        try
                        {
                            System.Windows.Forms.Clipboard.SetText(clipboardText);
                            A.Ed.WriteMessage("\n✅ Đã copy kết quả ra clipboard!");
                        }
                        catch { }
                    }
                }

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Lấy diện tích vật liệu từ SectionViewGroup
        /// </summary>
        [CommandMethod("CTSV_LayDienTichTuSectionView")]
        public static void CTSVLayDienTichTuSectionView()
        {
            try
            {
                A.Ed.WriteMessage("\n\n=== LẤY DIỆN TÍCH TỪ SECTION VIEW ===");

                using Transaction tr = A.Db.TransactionManager.StartTransaction();

                // Lấy danh sách Alignment có SampleLineGroup
                var alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n❌ Không tìm thấy Alignment nào có SampleLineGroup.");
                    tr.Commit();
                    return;
                }

                // Hiển thị form chọn Alignment
                FormChonAlignment formChon = new(alignments);
                if (formChon.ShowDialog() != DialogResult.OK || formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    tr.Commit();
                    return;
                }

                var selectedAlign = formChon.SelectedAlignments[0];

                // Mở SampleLineGroup với ForWrite để truy cập SectionViewGroups
                SampleLineGroup? slg = tr.GetObject(selectedAlign.SampleLineGroupId, AcadDb.OpenMode.ForWrite) as SampleLineGroup;
                if (slg == null)
                {
                    A.Ed.WriteMessage("\n❌ Không thể mở SampleLineGroup.");
                    tr.Commit();
                    return;
                }

                // Lấy SectionViewGroup(s) từ collection
                SectionViewGroupCollection svgCollection = slg.SectionViewGroups;
                if (svgCollection.Count == 0)
                {
                    A.Ed.WriteMessage("\n❌ Không có SectionViewGroup. Hãy tạo Section Views trước.");
                    tr.Commit();
                    return;
                }

                A.Ed.WriteMessage($"\n\n📊 DIỆN TÍCH TỪ SECTION VIEW - {selectedAlign.Name}");
                A.Ed.WriteMessage($"\n{'=',-80}");

                // Duyệt qua từng SectionViewGroup
                foreach (SectionViewGroup svg in svgCollection)
                {
                    if (svg == null) continue;

                    A.Ed.WriteMessage($"\n\n📁 SectionViewGroup: {svg.Name}");
                    A.Ed.WriteMessage($"\n{'-',-70}");

                    // Lấy các SectionView từ SectionViewGroup
                    ObjectIdCollection svIds = svg.GetSectionViewIds();

                    foreach (ObjectId svId in svIds)
                    {
                        SectionView? sectionView = tr.GetObject(svId, AcadDb.OpenMode.ForRead) as SectionView;
                        if (sectionView == null) continue;

                        // Lấy SampleLine tương ứng
                        ObjectId slId = sectionView.SampleLineId;
                        SampleLine? sampleLine = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                        if (sampleLine == null) continue;

                        A.Ed.WriteMessage($"\n  📍 {sampleLine.Name} ({FormatStation(sampleLine.Station)})");

                        // Lấy các Section trong SectionView
                        try
                        {
                            ObjectIdCollection sectionIds = sampleLine.GetSectionIds();
                            SectionSourceCollection sources = slg.GetSectionSources();

                            foreach (ObjectId sectionId in sectionIds)
                            {
                                CivSection? section = tr.GetObject(sectionId, AcadDb.OpenMode.ForRead, false, true) as CivSection;
                                if (section == null) continue;

                                // Tìm tên source
                                string sourceName = "Unknown";
                                string sourceType = "";
                                foreach (SectionSource source in sources)
                                {
                                    if (source.SourceId == section.SourceId)
                                    {
                                        sourceName = source.SourceName;
                                        sourceType = source.SourceType.ToString();
                                        break;
                                    }
                                }

                                double area = section.Area > 0 ? section.Area : CalculateSectionArea(section);
                                if (area > 0)
                                {
                                    A.Ed.WriteMessage($"\n      [{sourceType}] {sourceName}: {area:F4} m²");
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            A.Ed.WriteMessage($"\n      ⚠️ Lỗi: {ex.Message}");
                        }
                    }
                }

                A.Ed.WriteMessage($"\n\n{'=',-80}");
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Lệnh lấy dữ liệu từ Volume Tables trong Section Views
        /// </summary>
        [CommandMethod("CTSV_LayKhoiLuongTracNgang")]
        public static void CTSVLayKhoiLuongTracNgang()
        {
            try
            {
                A.Ed.WriteMessage("\n\n=== LẤY KHỐI LƯỢNG TỪ TRẮC NGANG (VOLUME TABLES) ===");
                A.Ed.WriteMessage("\n📋 Trích xuất dữ liệu Area và Volume từ bảng khối lượng trong Section View");

                using Transaction tr = A.Db.TransactionManager.StartTransaction();

                // 1. Lấy danh sách Alignment có SampleLineGroup
                var alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n❌ Không tìm thấy Alignment nào có SampleLineGroup.");
                    tr.Commit();
                    return;
                }

                // 2. Chọn Alignment
                FormChonAlignment formChon = new(alignments);
                if (formChon.ShowDialog() != DialogResult.OK || formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    tr.Commit();
                    return;
                }

                // 3. Trích xuất dữ liệu từ Volume Tables
                Dictionary<string, List<CrossSectionVolumeData>> allData = new();

                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    A.Ed.WriteMessage($"\n\n📊 Đang xử lý: {alignInfo.Name}...");
                    var volumeData = ExtractVolumeTableData(tr, alignInfo.SampleLineGroupId, alignInfo.AlignmentId);
                    if (volumeData.Count > 0)
                    {
                        allData[alignInfo.Name] = volumeData;
                        A.Ed.WriteMessage($" ✅ ({volumeData.Count} trắc ngang)");
                    }
                    else
                    {
                        A.Ed.WriteMessage($" ⚠️ Không có dữ liệu Volume Table");
                    }
                }

                if (allData.Count == 0)
                {
                    A.Ed.WriteMessage("\n❌ Không có dữ liệu Volume Table nào.");
                    A.Ed.WriteMessage("\n💡 Hãy đảm bảo đã tạo Volume Tables trong Section Views.");
                    tr.Commit();
                    return;
                }

                // 4. Hiển thị kết quả và xuất Excel
                foreach (var kvp in allData)
                {
                    A.Ed.WriteMessage($"\n\n{'=',-70}");
                    A.Ed.WriteMessage($"\n📍 ALIGNMENT: {kvp.Key}");
                    A.Ed.WriteMessage($"\n{'=',-70}");

                    // Thu thập tất cả materials
                    HashSet<string> allMaterials = new();
                    foreach (var cs in kvp.Value)
                    {
                        foreach (var mat in cs.Materials.Keys)
                            allMaterials.Add(mat);
                    }

                    // Header
                    A.Ed.WriteMessage($"\n{"Lý trình",-15} | {"Khoảng cách",-12}");
                    foreach (var mat in allMaterials.OrderBy(m => m))
                    {
                        string shortMat = mat.Length > 12 ? mat.Substring(0, 10) + ".." : mat;
                        A.Ed.WriteMessage($" | {shortMat + " (m²)",-14} | {shortMat + " (m³)",-14}");
                    }

                    // Data rows
                    foreach (var cs in kvp.Value)
                    {
                        A.Ed.WriteMessage($"\n{cs.StationFormatted,-15} | {cs.SpacingPrev,12:F2}");
                        foreach (var mat in allMaterials.OrderBy(m => m))
                        {
                            if (cs.Materials.TryGetValue(mat, out var data))
                            {
                                A.Ed.WriteMessage($" | {data.Area,14:F4} | {data.CumulativeVolume,14:F4}");
                            }
                            else
                            {
                                A.Ed.WriteMessage($" | {"-",14} | {"-",14}");
                            }
                        }
                    }
                }

                // 5. Hỏi xuất Excel
                var exportResult = A.Ed.GetKeywords("\n\nXuất ra Excel? [Yes/No] <Yes>: ", new[] { "Yes", "No" });
                if (exportResult.Status != PromptStatus.OK || exportResult.StringResult != "No")
                {
                    var saveFileDialog = new System.Windows.Forms.SaveFileDialog
                    {
                        Filter = "Excel Files|*.xlsx",
                        Title = "Lưu file khối lượng trắc ngang",
                        FileName = $"KhoiLuong_TracNgang_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                    };

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        ExportVolumeTableToExcel(allData, saveFileDialog.FileName);
                        A.Ed.WriteMessage($"\n\n✅ Đã xuất file: {saveFileDialog.FileName}");

                        try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(saveFileDialog.FileName) { UseShellExecute = true }); }
                        catch { }
                    }
                }

                tr.Commit();
                A.Ed.WriteMessage("\n\n=== HOÀN THÀNH ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
                A.Ed.WriteMessage($"\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Trích xuất dữ liệu từ Volume Tables trong Section Views
        /// </summary>
        private static List<CrossSectionVolumeData> ExtractVolumeTableData(Transaction tr, ObjectId sampleLineGroupId, ObjectId alignmentId)
        {
            List<CrossSectionVolumeData> results = new();

            SampleLineGroup? slg = tr.GetObject(sampleLineGroupId, AcadDb.OpenMode.ForWrite, false, true) as SampleLineGroup;
            if (slg == null) return results;

            Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
            if (alignment == null) return results;

            // Lấy Material Lists để có thông tin materials
            QTOMaterialListCollection materialLists = slg.MaterialLists;
            Dictionary<Guid, (string ListName, Guid ListGuid, List<(string MatName, Guid MatGuid, MaterialQuantityType Type)> Materials)> matListInfo = new();

            foreach (QTOMaterialList matList in materialLists)
            {
                var materials = new List<(string, Guid, MaterialQuantityType)>();
                foreach (QTOMaterial mat in matList)
                {
                    materials.Add((mat.Name, mat.Guid, mat.QuantityType));
                }
                matListInfo[matList.Guid] = (matList.Name ?? "Unnamed", matList.Guid, materials);
            }

            // Lấy danh sách SampleLine và sắp xếp theo lý trình
            List<SampleLine> sortedSampleLines = new();
            foreach (ObjectId slId in slg.GetSampleLineIds())
            {
                var sl = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                if (sl != null) sortedSampleLines.Add(sl);
            }
            sortedSampleLines = sortedSampleLines.OrderBy(s => s.Station).ToList();

            // Duyệt qua từng SampleLine
            double prevStation = 0;
            bool isFirst = true;

            foreach (SampleLine sampleLine in sortedSampleLines)
            {
                CrossSectionVolumeData csData = new()
                {
                    SampleLineName = sampleLine.Name,
                    Station = sampleLine.Station,
                    StationFormatted = FormatStation(sampleLine.Station),
                    SpacingPrev = isFirst ? 0 : (sampleLine.Station - prevStation)
                };

                // Lấy dữ liệu từ Material Sections (vì Volume Tables chỉ hiển thị, không có API trực tiếp để đọc)
                foreach (var (listName, listGuid, materials) in matListInfo.Values)
                {
                    foreach (var (matName, matGuid, matType) in materials)
                    {
                        try
                        {
                            ObjectId materialSectionId = sampleLine.GetMaterialSectionId(listGuid, matGuid);
                            if (!materialSectionId.IsNull && materialSectionId.IsValid)
                            {
                                AcadDb.DBObject? sectionObj = tr.GetObject(materialSectionId, AcadDb.OpenMode.ForRead, false, true);
                                if (sectionObj is CivSection section)
                                {
                                    // Sử dụng section.Area từ API Civil 3D (giá trị chính xác - giống V3Tools)
                                    // Làm tròn theo cài đặt
                                    double area = RoundArea(section.Area);
                                    
                                    // Debug: Hiển thị thông tin cho trắc ngang đầu tiên
                                    if (isFirst && results.Count == 0)
                                    {
                                        A.Ed.WriteMessage($"\n  📍 {matName}: section.Area = {area:F4} m²");
                                    }

                                    string volumeType = matType == MaterialQuantityType.Cut ? "Cut" :
                                                       matType == MaterialQuantityType.Fill ? "Fill" : "Material";

                                    // Chỉ thêm nếu area > 0
                                    if (area > 0)
                                    {
                                        csData.Materials[matName] = new VolumeTableDataInfo
                                        {
                                            MaterialName = matName,
                                            Area = area,
                                            CumulativeVolume = 0, // Sẽ tính sau
                                            VolumeType = volumeType
                                        };
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                }

                results.Add(csData);
                prevStation = sampleLine.Station;
                isFirst = false;
            }

            // Tính khối lượng theo Average End Area: V = (S1 + S2) / 2 × L
            for (int i = 1; i < results.Count; i++)
            {
                double spacing = results[i].SpacingPrev;
                foreach (var matName in results[i].Materials.Keys)
                {
                    double areaCurrent = results[i].Materials[matName].Area;
                    double areaPrev = 0;
                    if (results[i - 1].Materials.TryGetValue(matName, out var prevMat))
                    {
                        areaPrev = prevMat.Area;
                    }

                    // Khối lượng đoạn giữa 2 trắc ngang
                    double segmentVolume = CalculateVolume(areaPrev, areaCurrent, spacing);
                    results[i].Materials[matName].SegmentVolume = segmentVolume;
                    
                    // Khối lượng tích lũy từ đầu
                    double prevCumulative = 0;
                    if (results[i - 1].Materials.TryGetValue(matName, out var prevMatData))
                    {
                        prevCumulative = prevMatData.CumulativeVolume;
                    }
                    
                    results[i].Materials[matName].CumulativeVolume = prevCumulative + segmentVolume;
                }
            }

            return results;
        }

        /// <summary>
        /// Xuất dữ liệu Volume Table ra Excel
        /// </summary>
        private static void ExportVolumeTableToExcel(Dictionary<string, List<CrossSectionVolumeData>> allData, string filePath)
        {
            using var workbook = new XLWorkbook();

            foreach (var kvp in allData)
            {
                string sheetName = kvp.Key.Length > 31 ? kvp.Key.Substring(0, 28) + "..." : kvp.Key;
                var worksheet = workbook.Worksheets.Add(sheetName);

                // Thu thập tất cả materials
                HashSet<string> allMaterials = new();
                foreach (var cs in kvp.Value)
                {
                    foreach (var mat in cs.Materials.Keys)
                        allMaterials.Add(mat);
                }
                var sortedMaterials = allMaterials.OrderBy(m => m).ToList();

                // Header Row 1
                worksheet.Cell(1, 1).Value = "STT";
                worksheet.Cell(1, 2).Value = "Lý trình";
                worksheet.Cell(1, 3).Value = "Khoảng cách (m)";

                int col = 4;
                foreach (var mat in sortedMaterials)
                {
                    worksheet.Cell(1, col).Value = mat;
                    worksheet.Range(1, col, 1, col + 2).Merge();  // 3 cột cho mỗi material
                    col += 3;
                }

                // Header Row 2: Chi tiết các cột
                worksheet.Cell(2, 1).Value = "";
                worksheet.Cell(2, 2).Value = "";
                worksheet.Cell(2, 3).Value = "";

                col = 4;
                foreach (var mat in sortedMaterials)
                {
                    worksheet.Cell(2, col).Value = "Diện tích (m²)";
                    worksheet.Cell(2, col + 1).Value = "KL đoạn (m³)";     // Khối lượng từng đoạn
                    worksheet.Cell(2, col + 2).Value = "KL tích lũy (m³)"; // Khối lượng tích lũy
                    col += 3;
                }

                // Data rows
                int row = 3;
                foreach (var cs in kvp.Value)
                {
                    worksheet.Cell(row, 1).Value = row - 2;
                    worksheet.Cell(row, 2).Value = cs.StationFormatted;
                    worksheet.Cell(row, 3).Value = cs.SpacingPrev;

                    col = 4;
                    foreach (var mat in sortedMaterials)
                    {
                        if (cs.Materials.TryGetValue(mat, out var data))
                        {
                            worksheet.Cell(row, col).Value = data.Area;
                            worksheet.Cell(row, col + 1).Value = data.SegmentVolume;
                            worksheet.Cell(row, col + 2).Value = data.CumulativeVolume;
                        }
                        col += 3;
                    }
                    row++;
                }

                // Summary row
                row++;
                worksheet.Cell(row, 1).Value = "TỔNG";
                worksheet.Cell(row, 1).Style.Font.Bold = true;
                worksheet.Range(row, 1, row, 3).Merge();

                col = 4;
                foreach (var mat in sortedMaterials)
                {
                    // Sum diện tích
                    worksheet.Cell(row, col).FormulaA1 = $"SUM({worksheet.Cell(3, col).Address}:{worksheet.Cell(row - 2, col).Address})";
                    // Sum khối lượng đoạn
                    worksheet.Cell(row, col + 1).FormulaA1 = $"SUM({worksheet.Cell(3, col + 1).Address}:{worksheet.Cell(row - 2, col + 1).Address})";
                    worksheet.Cell(row, col + 1).Style.Font.Bold = true;
                    // Max khối lượng tích lũy (= tổng cuối cùng)
                    worksheet.Cell(row, col + 2).FormulaA1 = $"MAX({worksheet.Cell(3, col + 2).Address}:{worksheet.Cell(row - 2, col + 2).Address})";
                    worksheet.Cell(row, col + 2).Style.Font.Bold = true;
                    col += 3;
                }

                // Formatting
                int totalCols = 3 + sortedMaterials.Count * 3;
                worksheet.Range(1, 1, 2, totalCols).Style.Font.Bold = true;
                worksheet.Range(1, 1, 2, totalCols).Style.Fill.BackgroundColor = XLColor.LightBlue;
                worksheet.Range(1, 1, row, totalCols).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range(1, 1, row, totalCols).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                // Set fixed column widths
                worksheet.Column(1).Width = 8;   // STT
                worksheet.Column(2).Width = 15;  // Lý trình
                worksheet.Column(3).Width = 15;  // Khoảng cách
                for (int c = 4; c <= totalCols; c++)
                {
                    worksheet.Column(c).Width = 14;
                }
            }

            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Chi tiết diện tích một Section
        /// </summary>
        public class SectionDetailData
        {
            public string SourceName { get; set; } = "";
            public string SourceType { get; set; } = "";
            public double Area { get; set; }           // Diện tích từ section.Area
            public double LeftLength { get; set; }     // Chiều dài bên trái
            public double RightLength { get; set; }    // Chiều dài bên phải
            public double MinElevation { get; set; }   // Cao độ min
            public double MaxElevation { get; set; }   // Cao độ max
        }

        /// <summary>
        /// Thông tin diện tích Section cho xuất Excel
        /// </summary>
        public class SectionAreaInfo
        {
            public string StakeName { get; set; } = "";
            public double StationValue { get; set; }
            public string Station { get; set; } = "";
            public double SpacingPrev { get; set; }
            public Dictionary<string, double> SectionAreas { get; set; } = new();
            public List<SectionDetailData> Details { get; set; } = new();
        }

        /// <summary>
        /// Lệnh xuất toàn bộ Section Area ra Excel
        /// </summary>
        [CommandMethod("CTSV_XuatSectionArea")]
        public static void CTSVXuatSectionArea()
        {
            try
            {
                A.Ed.WriteMessage("\n\n=== XUẤT SECTION AREA RA EXCEL ===");
                A.Ed.WriteMessage("\n(Lấy toàn bộ diện tích từ TIN Surface, Corridor, Material sections)");

                using Transaction tr = A.Db.TransactionManager.StartTransaction();

                // 1. Lấy danh sách Alignment có SampleLineGroup
                var alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n❌ Không tìm thấy Alignment nào có SampleLineGroup.");
                    tr.Commit();
                    return;
                }

                // 2. Chọn Alignment
                FormChonAlignment formChon = new(alignments);
                if (formChon.ShowDialog() != DialogResult.OK || formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    tr.Commit();
                    return;
                }

                // 3. Trích xuất Section Area cho từng Alignment đã chọn
                Dictionary<string, List<SectionAreaInfo>> allData = new();

                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    A.Ed.WriteMessage($"\n\n📊 Đang xử lý: {alignInfo.Name}...");
                    var sectionData = ExtractAllSectionAreas(tr, alignInfo.SampleLineGroupId, alignInfo.AlignmentId);
                    if (sectionData.Count > 0)
                    {
                        allData[alignInfo.Name] = sectionData;
                        A.Ed.WriteMessage($" ✅ ({sectionData.Count} cọc)");
                    }
                    else
                    {
                        A.Ed.WriteMessage($" ⚠️ Không có dữ liệu");
                    }
                }

                if (allData.Count == 0)
                {
                    A.Ed.WriteMessage("\n❌ Không có dữ liệu Section Area nào.");
                    tr.Commit();
                    return;
                }

                // 4. Xuất ra Excel
                var saveFileDialog = new System.Windows.Forms.SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Lưu file Section Area Excel",
                    FileName = $"SectionArea_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ExportSectionAreaToExcel(allData, saveFileDialog.FileName);
                    A.Ed.WriteMessage($"\n\n✅ Đã xuất file: {saveFileDialog.FileName}");
                    
                    // Mở file
                    try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(saveFileDialog.FileName) { UseShellExecute = true }); }
                    catch { }
                }

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Trích xuất toàn bộ Section Area từ SampleLineGroup
        /// </summary>
        private static List<SectionAreaInfo> ExtractAllSectionAreas(Transaction tr, ObjectId sampleLineGroupId, ObjectId alignmentId)
        {
            List<SectionAreaInfo> sectionInfos = new();

            SampleLineGroup? slg = tr.GetObject(sampleLineGroupId, AcadDb.OpenMode.ForWrite, false, true) as SampleLineGroup;
            if (slg == null) return sectionInfos;

            Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
            if (alignment == null) return sectionInfos;

            // Lấy Section Sources
            SectionSourceCollection sources = slg.GetSectionSources();
            A.Ed.WriteMessage($"\n  Section Sources: {sources.Count}");

            // Lấy danh sách SampleLine và sắp xếp theo lý trình
            List<SampleLine> sortedSampleLines = new List<SampleLine>();
            foreach (ObjectId slId in slg.GetSampleLineIds())
            {
                var sl = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                if (sl != null) sortedSampleLines.Add(sl);
            }
            sortedSampleLines = sortedSampleLines.OrderBy(s => s.Station).ToList();

            // Duyệt qua từng SampleLine đã sắp xếp
            double prevStation = 0;
            bool isFirst = true;

            foreach (SampleLine sampleLine in sortedSampleLines)
            {
                SectionAreaInfo info = new()
                {
                    StakeName = sampleLine.Name,
                    StationValue = sampleLine.Station,
                    Station = FormatStation(sampleLine.Station),
                    SpacingPrev = isFirst ? 0 : (sampleLine.Station - prevStation)
                };

                // Lấy tất cả Section từ SampleLine
                ObjectIdCollection sectionIds = sampleLine.GetSectionIds();

                foreach (ObjectId sectionId in sectionIds)
                {
                    try
                    {
                        CivSection? section = tr.GetObject(sectionId, AcadDb.OpenMode.ForRead, false, true) as CivSection;
                        if (section == null) continue;

                        // Tìm tên source
                        string sourceName = "Unknown";
                        string sourceType = "";
                        foreach (SectionSource source in sources)
                        {
                            if (source.SourceId == section.SourceId)
                            {
                                sourceName = source.SourceName;
                                sourceType = source.SourceType.ToString();
                                break;
                            }
                        }

                        // Lấy diện tích từ API
                        double area = section.Area;
                        if (area <= 0) area = CalculateSectionArea(section);

                        // Lấy thông tin chi tiết từ SectionPoints
                        double leftLength = 0, rightLength = 0, minElev = 0, maxElev = 0;
                        try
                        {
                            SectionPointCollection points = section.SectionPoints;
                            if (points.Count > 0)
                            {
                                double minX = double.MaxValue, maxX = double.MinValue;
                                minElev = double.MaxValue;
                                maxElev = double.MinValue;

                                foreach (SectionPoint pt in points)
                                {
                                    double x = pt.Location.X;
                                    double y = pt.Location.Y;
                                    
                                    if (x < minX) minX = x;
                                    if (x > maxX) maxX = x;
                                    if (y < minElev) minElev = y;
                                    if (y > maxElev) maxElev = y;
                                }

                                leftLength = Math.Abs(Math.Min(0, minX));  // Từ tim ra trái
                                rightLength = Math.Max(0, maxX);           // Từ tim ra phải
                            }
                        }
                        catch { }

                        // Thêm chi tiết
                        info.Details.Add(new SectionDetailData
                        {
                            SourceName = sourceName,
                            SourceType = sourceType,
                            Area = area,
                            LeftLength = leftLength,
                            RightLength = rightLength,
                            MinElevation = minElev,
                            MaxElevation = maxElev
                        });

                        if (area > 0)
                        {
                            string key = $"[{sourceType}] {sourceName}";
                            if (!info.SectionAreas.ContainsKey(key))
                                info.SectionAreas[key] = 0;
                            info.SectionAreas[key] += area;

                            if (isFirst)
                            {
                                A.Ed.WriteMessage($"\n    {key}: Area={area:F4}m², Left={leftLength:F2}m, Right={rightLength:F2}m");
                            }
                        }
                    }
                    catch { }
                }

                // Thêm Material Sections
                QTOMaterialListCollection materialLists = slg.MaterialLists;
                foreach (QTOMaterialList materialList in materialLists)
                {
                    try
                    {
                        Guid listGuid = materialList.Guid;
                        foreach (QTOMaterial material in materialList)
                        {
                            try
                            {
                                ObjectId materialSectionId = sampleLine.GetMaterialSectionId(listGuid, material.Guid);
                                if (!materialSectionId.IsNull && materialSectionId.IsValid)
                                {
                                    AcadDb.DBObject? sectionObj = tr.GetObject(materialSectionId, AcadDb.OpenMode.ForRead, false, true);
                                    if (sectionObj is CivSection section)
                                    {
                                        double area = section.Area;
                                        if (area <= 0) area = CalculateSectionArea(section);

                                        // Lấy thông tin chi tiết từ SectionPoints
                                        double leftLength = 0, rightLength = 0, minElev = 0, maxElev = 0;
                                        try
                                        {
                                            SectionPointCollection points = section.SectionPoints;
                                            if (points.Count > 0)
                                            {
                                                double minX = double.MaxValue, maxX = double.MinValue;
                                                minElev = double.MaxValue;
                                                maxElev = double.MinValue;

                                                foreach (SectionPoint pt in points)
                                                {
                                                    double x = pt.Location.X;
                                                    double y = pt.Location.Y;
                                                    
                                                    if (x < minX) minX = x;
                                                    if (x > maxX) maxX = x;
                                                    if (y < minElev) minElev = y;
                                                    if (y > maxElev) maxElev = y;
                                                }

                                                leftLength = Math.Abs(Math.Min(0, minX));
                                                rightLength = Math.Max(0, maxX);
                                            }
                                        }
                                        catch { }

                                        // Thêm chi tiết Material Section
                                        info.Details.Add(new SectionDetailData
                                        {
                                            SourceName = material.Name,
                                            SourceType = "Material",
                                            Area = area,
                                            LeftLength = leftLength,
                                            RightLength = rightLength,
                                            MinElevation = minElev,
                                            MaxElevation = maxElev
                                        });

                                        if (area > 0)
                                        {
                                            string key = $"[Material] {material.Name}";
                                            if (!info.SectionAreas.ContainsKey(key))
                                                info.SectionAreas[key] = 0;
                                            info.SectionAreas[key] += area;

                                            if (isFirst)
                                            {
                                                A.Ed.WriteMessage($"\n    {key}: Area={area:F4}m², Left={leftLength:F2}m, Right={rightLength:F2}m");
                                            }
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }

                sectionInfos.Add(info);
                prevStation = sampleLine.Station;
                isFirst = false;
            }

            return sectionInfos;
        }

        /// <summary>
        /// Xuất Section Area ra Excel
        /// </summary>
        private static void ExportSectionAreaToExcel(Dictionary<string, List<SectionAreaInfo>> data, string filePath)
        {
            using var workbook = new XLWorkbook();

            foreach (var kvp in data)
            {
                string alignName = kvp.Key;
                var sectionInfos = kvp.Value;
                if (sectionInfos.Count == 0) continue;

                // Lấy tất cả các loại Section (columns)
                var allSectionTypes = sectionInfos
                    .SelectMany(s => s.SectionAreas.Keys)
                    .Distinct()
                    .OrderBy(s => s)
                    .ToList();

                string sheetName = SanitizeSheetName(alignName);
                var ws = workbook.Worksheets.Add(sheetName);

                // Header
                ws.Cell(1, 1).Value = $"BẢNG DIỆN TÍCH SECTION - {alignName}";
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Font.FontSize = 14;
                ws.Range(1, 1, 1, 3 + allSectionTypes.Count).Merge();

                // Row 3: Column headers
                ws.Cell(3, 1).Value = "TT";
                ws.Cell(3, 2).Value = "Lý trình";
                ws.Cell(3, 3).Value = "Khoảng cách (m)";

                int col = 4;
                foreach (var sectionType in allSectionTypes)
                {
                    ws.Cell(3, col).Value = $"DT {sectionType} (m²)";
                    col++;
                }

                // Dữ liệu
                int row = 4;
                int stt = 1;
                foreach (var info in sectionInfos)
                {
                    ws.Cell(row, 1).Value = stt++;
                    ws.Cell(row, 2).Value = info.Station;
                    ws.Cell(row, 3).Value = Math.Round(info.SpacingPrev, 2);

                    col = 4;
                    foreach (var sectionType in allSectionTypes)
                    {
                        double area = info.SectionAreas.ContainsKey(sectionType) ? info.SectionAreas[sectionType] : 0;
                        ws.Cell(row, col).Value = Math.Round(area, 4);
                        col++;
                    }

                    row++;
                }

                // Dòng tổng hợp
                ws.Cell(row, 1).Value = "";
                ws.Cell(row, 2).Value = "TỔNG";
                ws.Cell(row, 2).Style.Font.Bold = true;
                ws.Cell(row, 3).FormulaA1 = $"=SUM(C4:C{row - 1})";

                col = 4;
                foreach (var sectionType in allSectionTypes)
                {
                    string colLetter = GetColumnLetter(col);
                    ws.Cell(row, col).FormulaA1 = $"=SUM({colLetter}4:{colLetter}{row - 1})";
                    ws.Cell(row, col).Style.Font.Bold = true;
                    col++;
                }

                // Format
                FormatWorksheet(ws, row, col - 1);
            }

            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Tính khối lượng kết hợp: So sánh Surface + Diện tích từ Section
        /// </summary>
        [CommandMethod("CTSV_TinhKLKetHop")]
        public static void CTSVTinhKLKetHop()
        {
            try
            {
                A.Ed.WriteMessage("\n\n=== TÍNH KHỐI LƯỢNG KẾT HỢP ===");
                A.Ed.WriteMessage("\n(So sánh Surface + Diện tích từ SectionView)");

                using Transaction tr = A.Db.TransactionManager.StartTransaction();

                // 1. Lấy danh sách Alignment có SampleLineGroup
                var alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n❌ Không tìm thấy Alignment nào có SampleLineGroup.");
                    tr.Commit();
                    return;
                }

                // 2. Chọn Alignment
                FormChonAlignment formChon = new(alignments);
                if (formChon.ShowDialog() != DialogResult.OK || formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    tr.Commit();
                    return;
                }

                // 3. Trích xuất dữ liệu từ Material List + Section
                HashSet<string> allMaterials = new();
                Dictionary<ObjectId, List<StakeInfo>> alignmentData = new();

                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    // Lấy dữ liệu từ Material List
                    var stakeInfos = ExtractMaterialData(tr, alignInfo.SampleLineGroupId, alignInfo.AlignmentId);
                    
                    // Bổ sung dữ liệu từ SectionViewGroup (nếu có)
                    EnrichWithSectionViewData(tr, alignInfo.SampleLineGroupId, stakeInfos);
                    
                    alignmentData[alignInfo.AlignmentId] = stakeInfos;

                    foreach (var stake in stakeInfos)
                    {
                        foreach (var mat in stake.MaterialAreas.Keys)
                            allMaterials.Add(mat);
                        foreach (var mat in stake.CorridorAreas.Keys)
                            allMaterials.Add($"[Corridor] {mat}");
                        foreach (var mat in stake.SurfaceAreas.Keys)
                            allMaterials.Add($"[Surface] {mat}");
                    }
                }

                if (allMaterials.Count == 0)
                {
                    A.Ed.WriteMessage("\n❌ Không tìm thấy dữ liệu vật liệu nào!");
                    tr.Commit();
                    return;
                }

                // 4. So sánh Surface (tùy chọn)
                VolumeSurfaceData? volumeData = null;
                var addVolSurf = A.Ed.GetKeywords("\nBạn có muốn thêm so sánh Volume Surface? [Yes/No] <No>: ", new[] { "Yes", "No" });
                if (addVolSurf.Status == PromptStatus.OK && addVolSurf.StringResult == "Yes")
                {
                    volumeData = GetVolumeSurfaceComparison(tr);
                }

                // 5. Sắp xếp vật liệu
                FormSapXepVatLieu formSapXep = new(allMaterials.ToList());
                if (formSapXep.ShowDialog() != DialogResult.OK)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    tr.Commit();
                    return;
                }
                List<string> orderedMaterials = formSapXep.OrderedMaterials;

                // 6. Tính khối lượng
                foreach (var kvp in alignmentData)
                {
                    CalculateVolumesExtended(kvp.Value, orderedMaterials);
                }

                // 7. Xuất ra Excel
                SaveFileDialog saveDialog = new()
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    Title = "Chọn nơi lưu file Excel khối lượng",
                    FileName = "KhoiLuongKetHop.xlsx"
                };

                if (saveDialog.ShowDialog() != DialogResult.OK)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    tr.Commit();
                    return;
                }

                ExportToExcelExtended(saveDialog.FileName, formChon.SelectedAlignments, alignmentData, orderedMaterials, volumeData, tr);
                A.Ed.WriteMessage($"\n✅ Đã xuất file Excel: {saveDialog.FileName}");

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Lấy thông tin Volume Surface từ người dùng
        /// </summary>
        private static VolumeSurfaceData? GetVolumeSurfaceComparison(Transaction tr)
        {
            var surfaceIds = A.Cdoc.GetSurfaceIds();
            if (surfaceIds.Count < 2) return null;

            var surfaces = new List<(int Index, ObjectId Id, string Name)>();
            int idx = 1;
            A.Ed.WriteMessage("\n\nDanh sách Surface:");

            foreach (ObjectId id in surfaceIds)
            {
                var surf = tr.GetObject(id, AcadDb.OpenMode.ForRead);
                if (surf is TinSurface tinSurf)
                {
                    A.Ed.WriteMessage($"\n  {idx}. {tinSurf.Name}");
                    surfaces.Add((idx, id, tinSurf.Name));
                    idx++;
                }
            }

            if (surfaces.Count < 2) return null;

            var baseResult = A.Ed.GetInteger($"\nChọn Surface TỰ NHIÊN (1-{surfaces.Count}): ");
            if (baseResult.Status != PromptStatus.OK || baseResult.Value < 1 || baseResult.Value > surfaces.Count)
                return null;

            var compResult = A.Ed.GetInteger($"\nChọn Surface THIẾT KẾ (1-{surfaces.Count}): ");
            if (compResult.Status != PromptStatus.OK || compResult.Value < 1 || compResult.Value > surfaces.Count)
                return null;

            var baseSurface = surfaces[baseResult.Value - 1];
            var compSurface = surfaces[compResult.Value - 1];

            if (baseSurface.Id == compSurface.Id) return null;

            // Tạo Volume Surface
            string volSurfName = $"VOL_{baseSurface.Name}_{compSurface.Name}";
            TinVolumeSurface? volumeSurface = null;

            // Kiểm tra có sẵn chưa
            foreach (ObjectId sid in surfaceIds)
            {
                var s = tr.GetObject(sid, AcadDb.OpenMode.ForRead);
                if (s is TinVolumeSurface tvs && tvs.Name == volSurfName)
                {
                    volumeSurface = tvs;
                    break;
                }
            }

            // Tạo mới nếu chưa có
            if (volumeSurface == null)
            {
                try
                {
                    var volSurfId = TinVolumeSurface.Create(volSurfName, baseSurface.Id, compSurface.Id);
                    volumeSurface = tr.GetObject(volSurfId, AcadDb.OpenMode.ForRead) as TinVolumeSurface;
                }
                catch { return null; }
            }

            if (volumeSurface != null)
            {
                var props = volumeSurface.GetVolumeProperties();
                return new VolumeSurfaceData
                {
                    SurfaceName = volSurfName,
                    BaseSurfaceName = baseSurface.Name,
                    ComparisonSurfaceName = compSurface.Name,
                    CutVolume = props.UnadjustedCutVolume,
                    FillVolume = props.UnadjustedFillVolume,
                    NetVolume = props.UnadjustedCutVolume - props.UnadjustedFillVolume
                };
            }

            return null;
        }

        /// <summary>
        /// Bổ sung dữ liệu từ SectionViewGroup vào StakeInfo
        /// </summary>
        private static void EnrichWithSectionViewData(Transaction tr, ObjectId sampleLineGroupId, List<StakeInfo> stakeInfos)
        {
            try
            {
                // Phải mở ForWrite để truy cập SectionViewGroups
                SampleLineGroup? slg = tr.GetObject(sampleLineGroupId, AcadDb.OpenMode.ForWrite) as SampleLineGroup;
                if (slg == null) return;

                SectionViewGroupCollection svgCollection = slg.SectionViewGroups;
                if (svgCollection.Count == 0) return;

                // Tạo dictionary để tra cứu nhanh StakeInfo theo Station
                var stakeByStation = stakeInfos.ToDictionary(s => Math.Round(s.StationValue, 3));

                foreach (SectionViewGroup svg in svgCollection)
                {
                    if (svg == null) continue;

                    ObjectIdCollection svIds = svg.GetSectionViewIds();
                    foreach (ObjectId svId in svIds)
                    {
                        SectionView? sectionView = tr.GetObject(svId, AcadDb.OpenMode.ForRead) as SectionView;
                        if (sectionView == null) continue;

                        SampleLine? sampleLine = tr.GetObject(sectionView.SampleLineId, AcadDb.OpenMode.ForRead) as SampleLine;
                        if (sampleLine == null) continue;

                        double station = Math.Round(sampleLine.Station, 3);
                        if (!stakeByStation.ContainsKey(station)) continue;

                        var stakeInfo = stakeByStation[station];

                        // Lấy thêm thông tin từ các Section trong SectionView
                        try
                        {
                            ObjectIdCollection sectionIds = sampleLine.GetSectionIds();
                            SectionSourceCollection sources = slg.GetSectionSources();

                            foreach (ObjectId sectionId in sectionIds)
                            {
                                CivSection? section = tr.GetObject(sectionId, AcadDb.OpenMode.ForRead, false, true) as CivSection;
                                if (section == null) continue;

                                foreach (SectionSource source in sources)
                                {
                                    if (source.SourceId == section.SourceId)
                                    {
                                        double area = section.Area > 0 ? section.Area : CalculateSectionArea(section);
                                        if (area <= 0) break;

                                        string name = source.SourceName;

                                        // Cập nhật vào đúng nhóm nếu chưa có
                                        switch (source.SourceType)
                                        {
                                            case SectionSourceType.Corridor:
                                                if (!stakeInfo.CorridorAreas.ContainsKey(name))
                                                    stakeInfo.CorridorAreas[name] = area;
                                                break;
                                            case SectionSourceType.TinSurface:
                                                if (!stakeInfo.SurfaceAreas.ContainsKey(name))
                                                    stakeInfo.SurfaceAreas[name] = area;
                                                break;
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Tính khối lượng mở rộng (bao gồm Corridor và Surface)
        /// </summary>
        private static void CalculateVolumesExtended(List<StakeInfo> stakeInfos, List<string> materials)
        {
            for (int i = 0; i < stakeInfos.Count; i++)
            {
                var stake = stakeInfos[i];
                double spacing = stake.SpacingPrev;

                // Tính khối lượng từ Material Areas
                foreach (var material in materials)
                {
                    // Material
                    if (stake.MaterialAreas.ContainsKey(material))
                    {
                        double areaCurrent = stake.MaterialAreas[material];
                        double areaPrev = (i > 0 && stakeInfos[i - 1].MaterialAreas.ContainsKey(material)) 
                            ? stakeInfos[i - 1].MaterialAreas[material] : 0;
                        stake.MaterialVolumes[material] = CalculateVolume(areaPrev, areaCurrent, spacing);
                    }

                    // Corridor (với prefix)
                    string corridorKey = material.Replace("[Corridor] ", "");
                    if (stake.CorridorAreas.ContainsKey(corridorKey))
                    {
                        double areaCurrent = stake.CorridorAreas[corridorKey];
                        double areaPrev = (i > 0 && stakeInfos[i - 1].CorridorAreas.ContainsKey(corridorKey)) 
                            ? stakeInfos[i - 1].CorridorAreas[corridorKey] : 0;
                        stake.CorridorVolumes[corridorKey] = CalculateVolume(areaPrev, areaCurrent, spacing);
                    }

                    // Surface (với prefix)
                    string surfaceKey = material.Replace("[Surface] ", "");
                    if (stake.SurfaceAreas.ContainsKey(surfaceKey))
                    {
                        double areaCurrent = stake.SurfaceAreas[surfaceKey];
                        double areaPrev = (i > 0 && stakeInfos[i - 1].SurfaceAreas.ContainsKey(surfaceKey)) 
                            ? stakeInfos[i - 1].SurfaceAreas[surfaceKey] : 0;
                        stake.SurfaceVolumes[surfaceKey] = CalculateVolume(areaPrev, areaCurrent, spacing);
                    }
                }
            }
        }

        /// <summary>
        /// Xuất Excel mở rộng với thông tin Volume Surface
        /// </summary>
        private static void ExportToExcelExtended(string filePath, List<AlignmentInfo> alignments,
            Dictionary<ObjectId, List<StakeInfo>> alignmentData, List<string> materials, 
            VolumeSurfaceData? volumeData, Transaction tr)
        {
            using var workbook = new XLWorkbook();

            // Sheet cho từng Alignment
            foreach (var alignInfo in alignments)
            {
                var stakeInfos = alignmentData[alignInfo.AlignmentId];
                string sheetName = SanitizeSheetName(alignInfo.Name);
                var ws = workbook.Worksheets.Add(sheetName);

                CreateSheetHeader(ws, materials, alignInfo.Name);

                int row = 4;
                foreach (var stake in stakeInfos)
                {
                    ws.Cell(row, 1).Value = stake.StakeName;
                    ws.Cell(row, 2).Value = stake.Station;
                    ws.Cell(row, 3).Value = Math.Round(stake.SpacingPrev, 3);

                    int col = 4;
                    foreach (var material in materials)
                    {
                        // Lấy diện tích từ đúng nhóm
                        double area = 0;
                        if (stake.MaterialAreas.ContainsKey(material))
                            area = stake.MaterialAreas[material];
                        else if (material.StartsWith("[Corridor] ") && stake.CorridorAreas.ContainsKey(material.Replace("[Corridor] ", "")))
                            area = stake.CorridorAreas[material.Replace("[Corridor] ", "")];
                        else if (material.StartsWith("[Surface] ") && stake.SurfaceAreas.ContainsKey(material.Replace("[Surface] ", "")))
                            area = stake.SurfaceAreas[material.Replace("[Surface] ", "")];

                        ws.Cell(row, col).Value = Math.Round(area, 4);
                        col++;
                    }

                    // Thêm cột khối lượng
                    foreach (var material in materials)
                    {
                        string areaColLetter = GetColumnLetter(4 + materials.IndexOf(material));
                        
                        // DT Trung bình
                        if (row == 4)
                            ws.Cell(row, col).Value = 0;
                        else
                            ws.Cell(row, col).FormulaA1 = $"=({areaColLetter}{row-1}+{areaColLetter}{row})/2";
                        col++;

                        // Khối lượng
                        string avgColLetter = GetColumnLetter(col - 1);
                        ws.Cell(row, col).FormulaA1 = $"={avgColLetter}{row}*C{row}";
                        col++;
                    }

                    row++;
                }

                // Format - tính lại cột cuối cùng
                int totalCols = 3 + materials.Count + materials.Count * 2; // 3 cột cố định + diện tích + (DT trung bình + Khối lượng)
                FormatWorksheet(ws, row - 1, totalCols);
            }

            // Sheet tổng hợp Volume Surface (nếu có)
            if (volumeData != null)
            {
                var wsVol = workbook.Worksheets.Add("Volume Surface");
                wsVol.Cell(1, 1).Value = "SO SÁNH BỀ MẶT";
                wsVol.Cell(1, 1).Style.Font.Bold = true;
                wsVol.Cell(1, 1).Style.Font.FontSize = 14;

                wsVol.Cell(3, 1).Value = "Surface tự nhiên:";
                wsVol.Cell(3, 2).Value = volumeData.BaseSurfaceName;

                wsVol.Cell(4, 1).Value = "Surface thiết kế:";
                wsVol.Cell(4, 2).Value = volumeData.ComparisonSurfaceName;

                wsVol.Cell(6, 1).Value = "Khối lượng ĐÀO (Cut):";
                wsVol.Cell(6, 2).Value = volumeData.CutVolume;
                wsVol.Cell(6, 3).Value = "m³";

                wsVol.Cell(7, 1).Value = "Khối lượng ĐẮP (Fill):";
                wsVol.Cell(7, 2).Value = volumeData.FillVolume;
                wsVol.Cell(7, 3).Value = "m³";

                wsVol.Cell(8, 1).Value = "Khối lượng RÒNG (Net):";
                wsVol.Cell(8, 2).Value = volumeData.NetVolume;
                wsVol.Cell(8, 3).Value = "m³";

                wsVol.Column(1).Width = 25;
                wsVol.Column(2).Width = 20;
                wsVol.Cell(6, 2).Style.NumberFormat.Format = "#,##0.00";
                wsVol.Cell(7, 2).Style.NumberFormat.Format = "#,##0.00";
                wsVol.Cell(8, 2).Style.NumberFormat.Format = "#,##0.00";
            }

            // Sheet chi tiết Material Section cho từng Alignment
            foreach (var alignInfo in alignments)
            {
                var stakeInfos = alignmentData[alignInfo.AlignmentId];
                
                // Kiểm tra xem có MaterialSectionDetails không
                bool hasDetails = stakeInfos.Any(s => s.MaterialSectionDetails.Count > 0);
                if (!hasDetails) continue;
                
                string detailSheetName = SanitizeSheetName($"CT_{alignInfo.Name}");
                var wsDetail = workbook.Worksheets.Add(detailSheetName);
                
                // Tiêu đề
                wsDetail.Cell(1, 1).Value = $"CHI TIẾT MATERIAL SECTION - {alignInfo.Name}";
                wsDetail.Range(1, 1, 1, 10).Merge();
                wsDetail.Cell(1, 1).Style.Font.Bold = true;
                wsDetail.Cell(1, 1).Style.Font.FontSize = 14;
                wsDetail.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                
                // Header
                wsDetail.Cell(2, 1).Value = "STT";
                wsDetail.Cell(2, 2).Value = "TÊN CỌC";
                wsDetail.Cell(2, 3).Value = "LÝ TRÌNH";
                wsDetail.Cell(2, 4).Value = "MATERIAL";
                wsDetail.Cell(2, 5).Value = "LEFT (m)";
                wsDetail.Cell(2, 6).Value = "RIGHT (m)";
                wsDetail.Cell(2, 7).Value = "TỔNG RỘNG (m)";
                wsDetail.Cell(2, 8).Value = "MIN ELEV (m)";
                wsDetail.Cell(2, 9).Value = "MAX ELEV (m)";
                wsDetail.Cell(2, 10).Value = "AREA (m²)";
                
                wsDetail.Range(2, 1, 2, 10).Style.Font.Bold = true;
                wsDetail.Range(2, 1, 2, 10).Style.Fill.BackgroundColor = XLColor.LightBlue;
                wsDetail.Range(2, 1, 2, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                
                // Dữ liệu
                int detailRow = 3;
                int stt = 1;
                
                foreach (var stake in stakeInfos)
                {
                    bool isFirst = true;
                    foreach (var kvp in stake.MaterialSectionDetails)
                    {
                        var detail = kvp.Value;
                        
                        wsDetail.Cell(detailRow, 1).Value = stt;
                        wsDetail.Cell(detailRow, 2).Value = isFirst ? stake.StakeName : "";
                        wsDetail.Cell(detailRow, 3).Value = isFirst ? stake.Station : "";
                        wsDetail.Cell(detailRow, 4).Value = detail.MaterialName;
                        wsDetail.Cell(detailRow, 5).Value = Math.Round(detail.LeftLength, 3);
                        wsDetail.Cell(detailRow, 6).Value = Math.Round(detail.RightLength, 3);
                        wsDetail.Cell(detailRow, 7).Value = Math.Round(detail.TotalWidth, 3);
                        wsDetail.Cell(detailRow, 8).Value = Math.Round(detail.MinElevation, 3);
                        wsDetail.Cell(detailRow, 9).Value = Math.Round(detail.MaxElevation, 3);
                        wsDetail.Cell(detailRow, 10).Value = Math.Round(detail.Area, 4);
                        
                        detailRow++;
                        isFirst = false;
                    }
                    stt++;
                }
                
                // Format
                wsDetail.Range(2, 1, detailRow - 1, 10).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                wsDetail.Range(2, 1, detailRow - 1, 10).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                
                wsDetail.Column(1).Width = 8;
                wsDetail.Column(2).Width = 18;
                wsDetail.Column(3).Width = 15;
                wsDetail.Column(4).Width = 30;
                for (int c = 5; c <= 10; c++)
                    wsDetail.Column(c).Width = 14;
            }

            workbook.SaveAs(filePath);
        }

        #endregion

        #region Commands

        [CommandMethod("CTSV_XuatKhoiLuong")]
        public static void CTSVXuatKhoiLuong()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                // 1. Lấy danh sách Alignments có SampleLineGroup
                List<AlignmentInfo> alignmentInfos = GetAlignmentsWithSampleLineGroups(tr);

                if (alignmentInfos.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông tìm thấy Alignment nào có SampleLineGroup!");
                    return;
                }

                // 2. Hiển thị form chọn Alignments
                FormChonAlignment formChon = new(alignmentInfos);
                if (formChon.ShowDialog() != DialogResult.OK || formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }

                // 3. Thu thập tất cả vật liệu từ các Alignment đã chọn
                HashSet<string> allMaterials = new();
                Dictionary<ObjectId, List<StakeInfo>> alignmentData = new();

                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    var stakeInfos = ExtractMaterialData(tr, alignInfo.SampleLineGroupId, alignInfo.AlignmentId);
                    alignmentData[alignInfo.AlignmentId] = stakeInfos;

                    foreach (var stake in stakeInfos)
                    {
                        foreach (var mat in stake.MaterialAreas.Keys)
                        {
                            allMaterials.Add(mat);
                        }
                    }
                }

                if (allMaterials.Count == 0)
                {
                    A.Ed.WriteMessage("\nKhông tìm thấy vật liệu nào trong các Alignment đã chọn!");
                    A.Ed.WriteMessage("\nĐảm bảo rằng bạn đã tạo Material Sections trong QTO Manager.");
                    return;
                }

                // 4. Hiển thị form sắp xếp thứ tự vật liệu
                FormSapXepVatLieu formSapXep = new(allMaterials.ToList());
                if (formSapXep.ShowDialog() != DialogResult.OK)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }

                List<string> orderedMaterials = formSapXep.OrderedMaterials;

                // 5. Hiển thị form cài đặt bảng thống kê
                if (!TableSettingsForm.ShowSettings())
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }

                // 6. Tính khối lượng cho từng Alignment
                foreach (var kvp in alignmentData)
                {
                    CalculateVolumes(kvp.Value, orderedMaterials);
                }

                // 7. Chọn loại xuất
                PromptKeywordOptions pkoExport = new("\nChọn loại xuất [Excel/CAD/TracNgang/TatCa]", "Excel CAD TracNgang TatCa");
                pkoExport.Keywords.Default = "Excel";
                pkoExport.AllowNone = true;
                PromptResult prExport = A.Ed.GetKeywords(pkoExport);

                if (prExport.Status != PromptStatus.OK && prExport.Status != PromptStatus.None)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }

                string exportType = prExport.StringResult ?? "Excel";
                bool doExcel = exportType == "Excel" || exportType == "TatCa";
                bool doCad = exportType == "CAD" || exportType == "TatCa";
                bool doTracNgang = exportType == "TracNgang" || exportType == "TatCa";

                // 8. Xuất ra Excel nếu được chọn
                string excelPath = "";
                if (doExcel)
                {
                    SaveFileDialog saveDialog = new()
                    {
                        Filter = "Excel Files (*.xlsx)|*.xlsx",
                        Title = "Chọn nơi lưu file Excel khối lượng",
                        FileName = "KhoiLuongVatLieu.xlsx"
                    };

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        excelPath = saveDialog.FileName;
                        ExportToExcel(excelPath, formChon.SelectedAlignments, alignmentData, orderedMaterials, tr);
                        A.Ed.WriteMessage($"\n✅ Đã xuất file Excel: {excelPath}");
                    }
                }

                // 9. Xuất bảng thống kê theo trắc ngang nếu được chọn
                if (doTracNgang)
                {
                    // Thu thập dữ liệu thống kê Material
                    Dictionary<ObjectId, List<MaterialStatInfo>> alignmentMaterialStats = new();
                    
                    foreach (var alignInfo in formChon.SelectedAlignments)
                    {
                        var materialStats = ExtractMaterialStatistics(tr, alignInfo.SampleLineGroupId, alignInfo.AlignmentId);
                        alignmentMaterialStats[alignInfo.AlignmentId] = materialStats;
                    }

                    // Hỏi loại xuất: Excel hay CAD
                    PromptKeywordOptions pkTracNgang = new("\nXuất thống kê trắc ngang [Excel/CAD]", "Excel CAD");
                    pkTracNgang.Keywords.Default = "Excel";
                    pkTracNgang.AllowNone = true;
                    PromptResult prTracNgang = A.Ed.GetKeywords(pkTracNgang);

                    bool tracNgangExcel = prTracNgang.StringResult != "CAD";

                    if (tracNgangExcel)
                    {
                        SaveFileDialog saveDialogTN = new()
                        {
                            Filter = "Excel Files (*.xlsx)|*.xlsx",
                            Title = "Lưu file Excel thống kê trắc ngang",
                            FileName = $"ThongKeTracNgang_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                        };

                        if (saveDialogTN.ShowDialog() == DialogResult.OK)
                        {
                            ExportMaterialStatisticsToExcel(saveDialogTN.FileName, formChon.SelectedAlignments, 
                                alignmentMaterialStats, orderedMaterials);
                            A.Ed.WriteMessage($"\n✅ Đã xuất file Excel thống kê trắc ngang: {saveDialogTN.FileName}");
                        }
                    }
                    else
                    {
                        PromptPointResult pprTN = A.Ed.GetPoint("\nChọn điểm chèn bảng thống kê trắc ngang: ");
                        if (pprTN.Status == PromptStatus.OK)
                        {
                            Point3d insertPointTN = pprTN.Value;
                            
                            foreach (var alignInfo in formChon.SelectedAlignments)
                            {
                                var materialStats = alignmentMaterialStats[alignInfo.AlignmentId];
                                
                                CreateMaterialStatTable(tr, insertPointTN, alignInfo.Name, materialStats, orderedMaterials);
                                A.Ed.WriteMessage($"\n✅ Đã vẽ bảng thống kê trắc ngang cho '{alignInfo.Name}'");

                                // Offset cho bảng tiếp theo
                                double tableHeight = (materialStats.Count + 4) * 8.0;
                                insertPointTN = new Point3d(insertPointTN.X, insertPointTN.Y - tableHeight - 50, insertPointTN.Z);
                            }
                        }
                    }
                }

                // 10. Vẽ bảng trong CAD nếu được chọn
                if (doCad)
                {
                    PromptPointOptions ppo = new($"\nChọn điểm chèn bảng khối lượng (các bảng tiếp theo sẽ cách nhau {TableSpacingX} đơn vị theo X):");
                    ppo.AllowNone = false;
                    PromptPointResult ppr = A.Ed.GetPoint(ppo);
                    
                    if (ppr.Status == PromptStatus.OK)
                    {
                        Point3d currentInsertPoint = ppr.Value;
                        int tableIndex = 0;
                        
                        // Vẽ bảng cho từng Alignment
                        foreach (var alignInfo in formChon.SelectedAlignments)
                        {
                            var stakeInfos = alignmentData[alignInfo.AlignmentId];
                            
                            // Tính chiều rộng bảng để offset cho bảng tiếp theo
                            int numCols = 3 + orderedMaterials.Count * 2;
                            double tableWidth = 25.0 + 25.0 + 15.0 + (numCols - 3) * 18.0; // Cột 1,2 = 25, Cột 3 = 15, còn lại = 18
                            
                            CreateCadTable(tr, currentInsertPoint, alignInfo.Name, stakeInfos, orderedMaterials);
                            A.Ed.WriteMessage($"\n✅ Đã vẽ bảng cho '{alignInfo.Name}' tại ({currentInsertPoint.X:F2}, {currentInsertPoint.Y:F2})");
                            
                            // Di chuyển điểm chèn sang phải cho bảng tiếp theo
                            currentInsertPoint = new Point3d(
                                currentInsertPoint.X + tableWidth + TableSpacingX, 
                                currentInsertPoint.Y, 
                                currentInsertPoint.Z);
                            
                            tableIndex++;
                        }
                    }
                }

                // 11. Hỏi mở file Excel nếu có
                if (!string.IsNullOrEmpty(excelPath))
                {
                    if (MessageBox.Show("Bạn có muốn mở file Excel?", "Hoàn thành", 
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = excelPath,
                            UseShellExecute = true
                        });
                    }
                }
                
                tr.Commit();
                A.Ed.WriteMessage("\n\n=== HOÀN THÀNH ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nLỗi: {ex.Message}");
                A.Ed.WriteMessage($"\nStack: {ex.StackTrace}");
            }
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// Lấy danh sách Alignments có SampleLineGroup
        /// </summary>
        private static List<AlignmentInfo> GetAlignmentsWithSampleLineGroups(Transaction tr)
        {
            List<AlignmentInfo> result = new();

            ObjectIdCollection alignmentIds = A.Cdoc.GetAlignmentIds();
            foreach (ObjectId alignmentId in alignmentIds)
            {
                Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
                if (alignment == null) continue;

                ObjectIdCollection slgIds = alignment.GetSampleLineGroupIds();
                if (slgIds.Count > 0)
                {
                    result.Add(new AlignmentInfo
                    {
                        AlignmentId = alignmentId,
                        Name = alignment.Name,
                        SampleLineGroupId = slgIds[0], // Lấy SampleLineGroup đầu tiên
                        IsSelected = true
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// Trích xuất dữ liệu vật liệu từ QTO Material List trong SampleLineGroup
        /// </summary>
        private static List<StakeInfo> ExtractMaterialData(Transaction tr, ObjectId sampleLineGroupId, ObjectId alignmentId)
        {
            List<StakeInfo> stakeInfos = new();

            // Mở SampleLineGroup với ForWrite để có thể truy cập MaterialLists
            SampleLineGroup? slg = tr.GetObject(sampleLineGroupId, AcadDb.OpenMode.ForWrite, false, true) as SampleLineGroup;
            if (slg == null) return stakeInfos;

            Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
            if (alignment == null) return stakeInfos;

            // Lấy Material Lists từ SampleLineGroup
            QTOMaterialListCollection materialLists = slg.MaterialLists;
            
            if (materialLists.Count == 0)
            {
                A.Ed.WriteMessage("\nKhông tìm thấy Material List nào trong SampleLineGroup!");
                A.Ed.WriteMessage("\nVui lòng tạo Material List trong QTO Manager trước.");
                return stakeInfos;
            }

            // Tạo danh sách QTOMaterial với tên, GUID và MaterialList GUID
            List<(string Name, Guid MaterialListGuid, Guid MaterialGuid, QTOMaterial Material)> materials = new();
            foreach (QTOMaterialList materialList in materialLists)
            {
                try
                {
                    Guid listGuid = materialList.Guid;
                    foreach (QTOMaterial material in materialList)
                    {
                        materials.Add((material.Name, listGuid, material.Guid, material));
                        // A.Ed.WriteMessage($"\n  Material: '{material.Name}', QuantityType: {material.QuantityType}");
                    }
                }
                catch { }
            }

            // Lấy danh sách SampleLine và sắp xếp theo lý trình
            List<SampleLine> sortedSampleLines = new List<SampleLine>();
            foreach (ObjectId slId in slg.GetSampleLineIds())
            {
                var sl = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                if (sl != null) sortedSampleLines.Add(sl);
            }
            sortedSampleLines = sortedSampleLines.OrderBy(s => s.Station).ToList();

            // Duyệt qua từng SampleLine đã sắp xếp
            double prevStation = 0;
            bool isFirst = true;

            foreach (SampleLine sampleLine in sortedSampleLines)
            {
                StakeInfo stakeInfo = new()
                {
                    StakeName = sampleLine.Name,
                    StationValue = sampleLine.Station,
                    Station = FormatStation(sampleLine.Station),
                    SpacingPrev = isFirst ? 0 : (sampleLine.Station - prevStation)
                };

                // Lấy MaterialSection cho từng Material
                foreach (var (materialName, materialListGuid, materialGuid, material) in materials)
                {
                    try
                    {
                        // Lấy MaterialSection ID từ SampleLine sử dụng MaterialList GUID và Material GUID
                        ObjectId materialSectionId = sampleLine.GetMaterialSectionId(materialListGuid, materialGuid);
                        
                        if (!materialSectionId.IsNull && materialSectionId.IsValid)
                        {
                            // Thử lấy MaterialSection
                            AcadDb.DBObject? sectionObj = tr.GetObject(materialSectionId, AcadDb.OpenMode.ForRead, false, true);
                            
                            if (sectionObj != null)
                            {
                                // Nếu là Section, lấy diện tích từ Civil 3D API
                                if (sectionObj is CivSection section)
                                {
                                    // ═══════════════════════════════════════════════════════════
                                    // DỮ LIỆU TỪ COMPUTE MATERIAL:
                                    // - Tên Material: QTOMaterial.Name
                                    // - Diện tích: CivSection.Area (tương đương Properties Panel)
                                    // ═══════════════════════════════════════════════════════════
                                    double area = section.Area;
                                    
                                    // Đếm số điểm trong section
                                    int pointCount = 0;
                                    try { pointCount = section.SectionPoints.Count; } catch { }
                                    
                                    // Log chi tiết cho debugging
                                    if (isFirst)
                                    {
                                        A.Ed.WriteMessage($"\n  ┌───────────────────────────────────────────────────────────┐");
                                        A.Ed.WriteMessage($"\n  │ 📍 SampleLine: {sampleLine.Name,-40} │");
                                        A.Ed.WriteMessage($"\n  ├───────────────────────────────────────────────────────────┤");
                                        A.Ed.WriteMessage($"\n  │ 📋 Material: {materialName,-44} │");
                                        A.Ed.WriteMessage($"\n  │   → AREA (CivSection.Area): {area,12:F4} m²              │");
                                        A.Ed.WriteMessage($"\n  │   → Số điểm SectionPoints: {pointCount,8}                   │");
                                        
                                        if (area == 0 && pointCount == 0)
                                        {
                                            A.Ed.WriteMessage($"\n  │ ⚠️ SECTION CHƯA ĐƯỢC COMPUTE!                            │");
                                            A.Ed.WriteMessage($"\n  │   → Hãy: Analyze > Compute Materials                     │");
                                        }
                                        else if (area == 0 && pointCount > 0)
                                        {
                                            A.Ed.WriteMessage($"\n  │ ⚠️ Có điểm nhưng AREA = 0 (polygon hở?)                  │");
                                        }
                                        else
                                        {
                                            A.Ed.WriteMessage($"\n  │   ✅ Khớp với Properties Panel / Data Section            │");
                                        }
                                        A.Ed.WriteMessage($"\n  └───────────────────────────────────────────────────────────┘");
                                    }
                                    
                                    // Nếu Area = 0 nhưng có SectionPoints, thử tính bằng Shoelace
                                    if (area == 0 && pointCount >= 3)
                                    {
                                        try
                                        {
                                            double calcArea = CalculateSectionArea(section);
                                            if (calcArea > 0) area = calcArea;
                                        }
                                        catch { }
                                    }
                                    
                                    // Lưu diện tích (kể cả khi = 0 để tracking)
                                    if (!stakeInfo.MaterialAreas.ContainsKey(materialName))
                                        stakeInfo.MaterialAreas[materialName] = 0;
                                    stakeInfo.MaterialAreas[materialName] += area;
                                    
                                    // === LẤY CHI TIẾT MATERIAL SECTION DATA ===
                                    double minOffset = double.MaxValue;
                                    double maxOffset = double.MinValue;
                                    double minElevation = double.MaxValue;
                                    double maxElevation = double.MinValue;
                                    List<Point3d> points = new();
                                    
                                    try
                                    {
                                        foreach (SectionPoint pt in section.SectionPoints)
                                        {
                                            points.Add(pt.Location);
                                            
                                            // Offset (X): âm = trái, dương = phải
                                            if (pt.Location.X < minOffset) minOffset = pt.Location.X;
                                            if (pt.Location.X > maxOffset) maxOffset = pt.Location.X;
                                            
                                            // Elevation (Y)
                                            if (pt.Location.Y < minElevation) minElevation = pt.Location.Y;
                                            if (pt.Location.Y > maxElevation) maxElevation = pt.Location.Y;
                                        }
                                    }
                                    catch { }
                                    
                                    // Tạo và lưu MaterialSectionDetail
                                    MaterialSectionDetail detail = new()
                                    {
                                        MaterialName = materialName,
                                        SectionSurfaceName = section.Name,
                                        LeftLength = minOffset != double.MaxValue ? minOffset : 0,
                                        RightLength = maxOffset != double.MinValue ? maxOffset : 0,
                                        MinElevation = minElevation != double.MaxValue ? minElevation : 0,
                                        MaxElevation = maxElevation != double.MinValue ? maxElevation : 0,
                                        Area = area,
                                        PointCount = points.Count,
                                        Points = points
                                    };
                                    stakeInfo.MaterialSectionDetails[materialName] = detail;
                                }
                            }
                        }
                        else if (isFirst)
                        {
                            A.Ed.WriteMessage($"\n  SampleLine '{sampleLine.Name}' - Material '{materialName}': Không có MaterialSection");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        if (isFirst)
                        {
                            A.Ed.WriteMessage($"\n  Lỗi khi lấy Material '{materialName}': {ex.Message}");
                        }
                    }
                }

                // Lấy các Section khác từ SampleLine (Corridor Shapes, Surfaces, v.v.)
                try
                {
                    ObjectIdCollection sectionIds = sampleLine.GetSectionIds();
                    SectionSourceCollection sectionSources = slg.GetSectionSources();
                    
                    // Tạo HashSet để tránh trùng lặp
                    HashSet<string> processedSources = new();
                    
                    foreach (ObjectId sectionId in sectionIds)
                    {
                        try
                        {
                            CivSection? section = tr.GetObject(sectionId, AcadDb.OpenMode.ForRead, false, true) as CivSection;
                            if (section == null) continue;

                            // Tìm SectionSource tương ứng
                            foreach (SectionSource source in sectionSources)
                            {
                                if (source.SourceId == section.SourceId)
                                {
                                    // Bỏ qua nếu là Material (đã xử lý ở trên)
                                    if (source.SourceType == SectionSourceType.Material)
                                        break;
                                    
                                    // Tạo tên duy nhất cho source
                                    string sourceName = source.SourceName;
                                    string sourceKey = $"{source.SourceType}_{sourceName}";
                                    
                                    // Bỏ qua nếu đã xử lý
                                    if (processedSources.Contains(sourceKey))
                                        break;
                                    processedSources.Add(sourceKey);
                                    
                                    // Tính diện tích
                                    double area = CalculateSectionArea(section);
                                    
                                    // Đặt tên hiển thị theo SourceType
                                    string displayName = sourceName;
                                    
                                    if (isFirst)
                                    {
                                        A.Ed.WriteMessage($"\n  SampleLine '{sampleLine.Name}' - {source.SourceType} '{sourceName}': Area = {area:F4} m²");
                                    }
                                    
                                    if (area > 0)
                                    {
                                        // Phân loại vào đúng nhóm theo SourceType
                                        if (source.SourceType == SectionSourceType.Corridor)
                                        {
                                            if (!stakeInfo.CorridorAreas.ContainsKey(displayName))
                                                stakeInfo.CorridorAreas[displayName] = 0;
                                            stakeInfo.CorridorAreas[displayName] += area;
                                        }
                                        else if (source.SourceType == SectionSourceType.TinSurface)
                                        {
                                            if (!stakeInfo.SurfaceAreas.ContainsKey(displayName))
                                                stakeInfo.SurfaceAreas[displayName] = 0;
                                            stakeInfo.SurfaceAreas[displayName] += area;
                                        }
                                        else
                                        {
                                            if (!stakeInfo.OtherAreas.ContainsKey(displayName))
                                                stakeInfo.OtherAreas[displayName] = 0;
                                            stakeInfo.OtherAreas[displayName] += area;
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                        catch { }
                    }
                }
                catch (System.Exception ex)
                {
                    if (isFirst)
                    {
                        A.Ed.WriteMessage($"\n  Lỗi khi lấy Section sources: {ex.Message}");
                    }
                }

                stakeInfos.Add(stakeInfo);
                prevStation = sampleLine.Station;
                isFirst = false;
            }

            // Debug: hiển thị tổng kết dữ liệu
            if (stakeInfos.Count > 0)
            {
                var allMaterials = stakeInfos.SelectMany(s => s.MaterialAreas.Keys).Distinct().ToList();
                A.Ed.WriteMessage($"\n\nTổng cộng {stakeInfos.Count} sample lines, {allMaterials.Count} vật liệu có dữ liệu: {string.Join(", ", allMaterials)}");
                
                // Tính tổng diện tích cho mỗi material
                foreach (var mat in allMaterials)
                {
                    double totalArea = stakeInfos.Sum(s => s.MaterialAreas.ContainsKey(mat) ? s.MaterialAreas[mat] : 0);
                    A.Ed.WriteMessage($"\n  - {mat}: Tổng diện tích = {totalArea:F4} m²");
                }
            }

            return stakeInfos;
        }

        /// <summary>
        /// Lấy tên nguồn Section từ SectionSource
        /// </summary>
        private static string GetSectionSourceName(Transaction tr, CivSection section, SampleLineGroup slg)
        {
            try
            {
                // Tìm SectionSource tương ứng
                SectionSourceCollection sources = slg.GetSectionSources();
                foreach (SectionSource source in sources)
                {
                    if (source.SourceId == section.SourceId)
                    {
                        try
                        {
                            AcadDb.DBObject? sourceObj = tr.GetObject(source.SourceId, AcadDb.OpenMode.ForRead, false, true);
                            if (sourceObj is Corridor corridor)
                            {
                                return "Corridor - " + corridor.Name;
                            }
                            else if (sourceObj is TinSurface surface)
                            {
                                return surface.Name;
                            }
                            else if (sourceObj is Pipe pipe)
                            {
                                return pipe.Name;
                            }
                            else
                            {
                                return source.SourceName;
                            }
                        }
                        catch
                        {
                            return source.SourceName;
                        }
                    }
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception)
            {
                // Ignore AutoCAD exceptions
            }
            catch (System.Exception)
            {
                // Ignore other exceptions
            }
            return "";
        }

        /// <summary>
        /// Tính diện tích Section từ SectionPoints (công thức Shoelace)
        /// Không sử dụng section.Area để có thể so sánh độc lập
        /// </summary>
        private static double CalculateSectionArea(CivSection section)
        {
            try
            {
                SectionPointCollection points = section.SectionPoints;
                if (points.Count < 3) return 0;

                List<Point2d> pointList = new();
                foreach (SectionPoint pt in points)
                {
                    pointList.Add(new Point2d(pt.Location.X, pt.Location.Y));
                }

                // Đóng đa giác nếu chưa đóng
                if (pointList.Count >= 2)
                {
                    // Thêm điểm đóng về đường cơ sở
                    Point2d first = pointList[0];
                    Point2d last = pointList[pointList.Count - 1];
                    
                    // Nếu section là hở (đường), cần đóng về baseline
                    if (Math.Abs(first.X - last.X) > 0.001 || Math.Abs(first.Y - last.Y) > 0.001)
                    {
                        // Đóng đa giác dọc theo đường Y=0 hoặc Y min
                        double yBase = Math.Min(first.Y, last.Y);
                        pointList.Add(new Point2d(last.X, yBase));
                        pointList.Add(new Point2d(first.X, yBase));
                    }
                }

                return CalculatePolygonArea(pointList);
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Tính khối lượng cho tất cả các cọc
        /// </summary>
        private static void CalculateVolumes(List<StakeInfo> stakeInfos, List<string> materials)
        {
            for (int i = 0; i < stakeInfos.Count; i++)
            {
                var stake = stakeInfos[i];
                foreach (var material in materials)
                {
                    double areaCurrent = stake.MaterialAreas.GetValueOrDefault(material, 0);
                    double areaPrev = (i > 0) ? stakeInfos[i - 1].MaterialAreas.GetValueOrDefault(material, 0) : 0;
                    double spacing = stake.SpacingPrev;

                    stake.MaterialVolumes[material] = CalculateVolume(areaPrev, areaCurrent, spacing);
                }
            }
        }

        /// <summary>
        /// Xuất dữ liệu ra file Excel
        /// </summary>
        private static void ExportToExcel(string filePath, List<AlignmentInfo> alignments,
            Dictionary<ObjectId, List<StakeInfo>> alignmentData, List<string> materials, Transaction tr)
        {
            using var workbook = new XLWorkbook();

            // Dictionary để lưu tổng khối lượng cho sheet tổng hợp
            Dictionary<string, Dictionary<string, double>> totalVolumes = new();

            // Tạo sheet cho từng Alignment
            foreach (var alignInfo in alignments)
            {
                var stakeInfos = alignmentData[alignInfo.AlignmentId];
                string sheetName = SanitizeSheetName(alignInfo.Name);
                var ws = workbook.Worksheets.Add(sheetName);

                // Tạo header
                CreateSheetHeader(ws, materials, alignInfo.Name);

                // Điền dữ liệu
                int row = 4; // Bắt đầu từ hàng 4 (sau 2 hàng header)
                Dictionary<string, double> alignmentTotalVolumes = new();
                
                foreach (var material in materials)
                {
                    alignmentTotalVolumes[material] = 0;
                }


                foreach (var stake in stakeInfos)
                {
                    ws.Cell(row, 1).Value = stake.StakeName;
                    ws.Cell(row, 2).Value = stake.Station;
                    ws.Cell(row, 3).Value = Math.Round(stake.SpacingPrev, 3);

                    // Ghi diện tích (Cột 4 đến 4 + materials.Count - 1)
                    int col = 4;
                    foreach (var material in materials)
                    {
                        double area = stake.MaterialAreas.GetValueOrDefault(material, 0);
                        ws.Cell(row, col).Value = Math.Round(area, 3);
                        col++;
                    }

                    // Cột Diện tích trung bình, Khối lượng (từ Civil 3D), Khối lượng cộng dồn
                    // Mỗi vật liệu có 3 cột liên tiếp
                    int currentMaterialCol = col;
                    for (int m = 0; m < materials.Count; m++)
                    {
                        string areaColLetter = GetColumnLetter(4 + m);
                        string spacingColLetter = GetColumnLetter(3); // Cột C
                        
                        // 1. Diện tích trung bình (công thức)
                        if (row == 4)
                        {
                            ws.Cell(row, currentMaterialCol).Value = 0;
                        }
                        else
                        {
                            ws.Cell(row, currentMaterialCol).FormulaA1 = $"=({areaColLetter}{row-1}+{areaColLetter}{row})/2";
                        }
                        
                        // 2. Khối lượng = Diện tích trung bình × Khoảng cách
                        string avgAreaColLetter = GetColumnLetter(currentMaterialCol);
                        ws.Cell(row, currentMaterialCol + 1).FormulaA1 = $"={avgAreaColLetter}{row}*{spacingColLetter}{row}";
                        
                        // 3. Khối lượng cộng dồn
                        string volColLetter = GetColumnLetter(currentMaterialCol + 1);
                        if (row == 4)
                        {
                            ws.Cell(row, currentMaterialCol + 2).FormulaA1 = $"={volColLetter}{row}";
                        }
                        else
                        {
                            string cumVolLetter = GetColumnLetter(currentMaterialCol + 2);
                            ws.Cell(row, currentMaterialCol + 2).FormulaA1 = $"={cumVolLetter}{row-1}+{volColLetter}{row}";
                        }

                        currentMaterialCol += 3;
                    }

                    row++;
                }

                // Thêm hàng tổng cộng
                int totalRow = row;
                ws.Cell(totalRow, 1).Value = "TỔNG CỘNG";
                ws.Range(totalRow, 1, totalRow, 3).Merge();
                ws.Cell(totalRow, 1).Style.Font.Bold = true;
                ws.Cell(totalRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Sum tất cả cột diện tích
                int colTotal = 4;
                for (int m = 0; m < materials.Count; m++)
                {
                    ws.Cell(totalRow, colTotal).FormulaA1 = $"SUM({GetColumnLetter(colTotal)}4:{GetColumnLetter(colTotal)}{row - 1})";
                    colTotal++;
                }
                
                // Sum tất cả các cột con (Average Area, Volume, Cumulative Volume)
                // Thực tế chỉ cần sum Volume là chính xác nhất
                for (int m = 0; m < materials.Count; m++)
                {
                    // Avg Area Sum (optional)
                    colTotal++; 
                    
                    // Volume Sum
                    ws.Cell(totalRow, colTotal).FormulaA1 = $"SUM({GetColumnLetter(colTotal)}4:{GetColumnLetter(colTotal)}{row - 1})";
                    colTotal++;
                    
                    // Cumulative Volume Sum (lấy giá trị cuối cùng thay vì sum)
                    ws.Cell(totalRow, colTotal).FormulaA1 = $"{GetColumnLetter(colTotal)}{row - 1}";
                    colTotal++;
                }

                // Format bảng
                FormatWorksheet(ws, row, 3 + materials.Count + materials.Count * 3);

                // Lưu tổng khối lượng cho sheet tổng hợp
                totalVolumes[sheetName] = alignmentTotalVolumes;
            }

            // Tạo sheet TỔNG HỢP nếu có nhiều hơn 1 alignment
            if (alignments.Count > 1)
            {
                CreateSummarySheet(workbook, alignments, totalVolumes, alignmentData, materials, tr);
            }

            // Lưu file
            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Tạo header cho sheet
        /// </summary>
        private static void CreateSheetHeader(IXLWorksheet ws, List<string> materials, string alignmentName)
        {
            // Tiêu đề sheet
            ws.Cell(1, 1).Value = $"BẢNG TÍNH KHỐI LƯỢNG - {alignmentName}";
            int lastCol = 3 + materials.Count * 2;
            ws.Range(1, 1, 1, lastCol).Merge();
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontSize = 14;
            ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Hàng 2: Nhóm header (Thông tin cọc | Diện tích | Khối lượng)
            ws.Cell(2, 1).Value = "THÔNG TIN CỌC";
            ws.Range(2, 1, 2, 3).Merge();
            ws.Cell(2, 1).Style.Font.Bold = true;
            ws.Cell(2, 1).Style.Fill.BackgroundColor = XLColor.LightGray;
            ws.Cell(2, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Nhóm DIỆN TÍCH
            int areaStartCol = 4;
            int areaEndCol = 3 + materials.Count;
            ws.Cell(2, areaStartCol).Value = "DIỆN TÍCH (m²)";
            ws.Range(2, areaStartCol, 2, areaEndCol).Merge();
            ws.Cell(2, areaStartCol).Style.Font.Bold = true;
            ws.Cell(2, areaStartCol).Style.Fill.BackgroundColor = XLColor.LightGreen;
            ws.Cell(2, areaStartCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Nhóm KHỐI LƯỢNG
            int volTableStartCol = areaEndCol + 1;
            int volTableEndCol = 3 + materials.Count + materials.Count * 3;
            ws.Cell(2, volTableStartCol).Value = "KHỐI LƯỢNG (m³)";
            ws.Range(2, volTableStartCol, 2, volTableEndCol).Merge();
            ws.Cell(2, volTableStartCol).Style.Font.Bold = true;
            ws.Cell(2, volTableStartCol).Style.Fill.BackgroundColor = XLColor.LightYellow;
            ws.Cell(2, volTableStartCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Hàng 3: Header chi tiết
            ws.Cell(3, 1).Value = "TÊN CỌC";
            ws.Cell(3, 2).Value = "LÝ TRÌNH";
            ws.Cell(3, 3).Value = "K.CÁCH (m)";

            // Header tên vật liệu cho nhóm diện tích
            int col = 4;
            foreach (var material in materials)
            {
                ws.Cell(3, col).Value = material;
                col++;
            }

            // Header cho nhóm khối lượng (mỗi vật liệu 3 cột)
            foreach (var material in materials)
            {
                ws.Cell(3, col).Value = "DT TB";      // Diện tích trung bình
                ws.Cell(3, col + 1).Value = material; // Khối lượng
                ws.Cell(3, col + 2).Value = "CỘNG DỒN"; // Khối lượng cộng dồn
                
                // Format riêng cho header vật liệu
                ws.Range(3, col, 3, col + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                col += 3;
            }

            // Format header row 3
            var headerRange = ws.Range(3, 1, 3, volTableEndCol);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.WrapText = true;
        }

        /// <summary>
        /// Format toàn bộ worksheet
        /// </summary>
        private static void FormatWorksheet(IXLWorksheet ws, int lastRow, int lastCol)
        {
            // Border cho toàn bộ bảng
            var tableRange = ws.Range(2, 1, lastRow, lastCol);
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Điều chỉnh độ rộng cột
            ws.Column(1).Width = 15;  // Tên cọc
            ws.Column(2).Width = 15;  // Lý trình
            ws.Column(3).Width = 15;  // Khoảng cách

            for (int c = 4; c <= lastCol; c++)
            {
                ws.Column(c).Width = 18;
            }

            // Căn giữa các cột số
            for (int r = 4; r <= lastRow; r++)
            {
                for (int c = 3; c <= lastCol; c++)
                {
                    ws.Cell(r, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                }
            }
        }

        /// <summary>
        /// Tạo sheet tổng hợp
        /// </summary>
        private static void CreateSummarySheet(XLWorkbook workbook, List<AlignmentInfo> alignments,
            Dictionary<string, Dictionary<string, double>> totalVolumes, 
            Dictionary<ObjectId, List<StakeInfo>> alignmentData,
            List<string> materials, Transaction tr)
        {
            var ws = workbook.Worksheets.Add("TỔNG HỢP");

            // Tiêu đề
            ws.Cell(1, 1).Value = "BẢNG TỔNG HỢP KHỐI LƯỢNG TẤT CẢ CÁC TUYẾN";
            int lastCol = 1 + materials.Count;
            ws.Range(1, 1, 1, lastCol).Merge();
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontSize = 14;
            ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Header
            ws.Cell(2, 1).Value = "TUYẾN";
            int col = 2;
            foreach (var material in materials)
            {
                ws.Cell(2, col).Value = $"{material} (m³)";
                col++;
            }

            // Header format
            var headerRange = ws.Range(2, 1, 2, lastCol);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGreen;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Dữ liệu
            int row = 3;
            foreach (var alignInfo in alignments)
            {
                string sheetName = SanitizeSheetName(alignInfo.Name);
                ws.Cell(row, 1).Value = alignInfo.Name;

                col = 2;
                foreach (var material in materials)
                {
                    int mIndex = materials.IndexOf(material);
                    // Cột Volume trong sheet chi tiết:
                    // 1 (Tên) + 1 (Lý trình) + 1 (K.cách) + materials.Count (Các cột Area) + (mIndex * 3) + 2 (Cột Volume trong nhóm 3 cột)
                    int targetCol = 3 + materials.Count + (mIndex * 3) + 2;
                    int lastRow = alignmentData[alignInfo.AlignmentId].Count + 4;
                    
                    ws.Cell(row, col).FormulaA1 = $"='{sheetName}'!{GetColumnLetter(targetCol)}{lastRow}";
                    col++;
                }
                row++;
            }

            // Hàng tổng cộng
            ws.Cell(row, 1).Value = "TỔNG CỘNG";
            ws.Cell(row, 1).Style.Font.Bold = true;
            col = 2;
            foreach (var material in materials)
            {
                ws.Cell(row, col).FormulaA1 = $"SUM({GetColumnLetter(col)}3:{GetColumnLetter(col)}{row - 1})";
                col++;
            }

            // Format
            var tableRange = ws.Range(2, 1, row, lastCol);
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Column(1).Width = 30;
            for (int c = 2; c <= lastCol; c++)
            {
                ws.Column(c).Width = 20;
            }
        }

        /// <summary>
        /// Lấy số hàng cuối cùng cho sheet để tham chiếu
        /// </summary>
        private static int GetLastRowForSheet(Dictionary<string, Dictionary<string, double>> totalVolumes, string sheetName, List<string> materials)
        {
            // Cần đếm số hàng trong sheet đó - giả sử là row của TỔNG CỘNG
            // Điều này cần được tính toán từ số lượng stake thực tế
            return 50; // Placeholder - sẽ được cập nhật khi tạo sheet
        }

        /// <summary>
        /// Lấy ký tự cột Excel từ số cột
        /// </summary>
        private static string GetColumnLetter(int columnNumber)
        {
            string columnLetter = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnLetter = Convert.ToChar('A' + modulo) + columnLetter;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnLetter;
        }

        /// <summary>
        /// Làm sạch tên sheet (Excel không cho phép một số ký tự đặc biệt)
        /// </summary>
        private static string SanitizeSheetName(string name)
        {
            char[] invalidChars = { ':', '\\', '/', '?', '*', '[', ']' };
            string result = name;
            foreach (char c in invalidChars)
            {
                result = result.Replace(c, '_');
            }
            // Giới hạn độ dài tên sheet là 31 ký tự
            if (result.Length > 31)
            {
                result = result.Substring(0, 31);
            }
            return result;
        }

        /// <summary>
        /// Tạo bảng khối lượng trong AutoCAD
        /// </summary>
        private static void CreateCadTable(Transaction tr, Point3d insertPoint, string alignmentName, 
            List<StakeInfo> stakeInfos, List<string> materials)
        {
            AcadDb.Database db = HostApplicationServices.WorkingDatabase;
            AcadDb.BlockTable bt = tr.GetObject(db.BlockTableId, AcadDb.OpenMode.ForRead) as AcadDb.BlockTable 
                ?? throw new System.Exception("Không thể mở BlockTable");
            AcadDb.BlockTableRecord btr = tr.GetObject(bt[AcadDb.BlockTableRecord.ModelSpace], AcadDb.OpenMode.ForWrite) as AcadDb.BlockTableRecord
                ?? throw new System.Exception("Không thể mở ModelSpace");

            // Tính số cột: 3 (Tên cọc, Lý trình, K.Cách) + materials*2 (Diện tích + Khối lượng)
            int numCols = 3 + materials.Count * 2;
            int numRows = stakeInfos.Count + 4; // 2 header + dữ liệu + 1 tổng cộng

            // Tạo Table
            AcadDb.Table table = new()
            {
                Position = insertPoint,
                TableStyle = db.Tablestyle
            };

            table.SetSize(numRows, numCols);

            // Kích thước ô
            double rowHeight = 8.0;    // Chiều cao hàng
            double colWidth = 25.0;    // Chiều rộng cột mặc định
            double dataColWidth = 18.0; // Chiều rộng cột dữ liệu

            for (int r = 0; r < numRows; r++)
            {
                table.Rows[r].Height = rowHeight;
            }

            table.Columns[0].Width = colWidth;  // Tên cọc
            table.Columns[1].Width = colWidth;  // Lý trình
            table.Columns[2].Width = 15.0;      // K.Cách

            for (int c = 3; c < numCols; c++)
            {
                table.Columns[c].Width = dataColWidth;
            }

            // ===== HÀNG 0: Tiêu đề bảng =====
            table.MergeCells(CellRange.Create(table, 0, 0, 0, numCols - 1));
            table.Cells[0, 0].TextString = $"BẢNG TÍNH KHỐI LƯỢNG - {alignmentName}";
            table.Cells[0, 0].TextHeight = 5.0;
            table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

            // ===== HÀNG 1: Header nhóm =====
            // Nhóm THÔNG TIN CỌC
            table.MergeCells(CellRange.Create(table, 1, 0, 1, 2));
            table.Cells[1, 0].TextString = "THÔNG TIN CỌC";
            table.Cells[1, 0].Alignment = CellAlignment.MiddleCenter;

            // Nhóm DIỆN TÍCH
            int areaEndCol = 2 + materials.Count;
            table.MergeCells(CellRange.Create(table, 1, 3, 1, areaEndCol));
            table.Cells[1, 3].TextString = "DIỆN TÍCH (m²)";
            table.Cells[1, 3].Alignment = CellAlignment.MiddleCenter;

            // Nhóm KHỐI LƯỢNG
            if (areaEndCol + 1 < numCols)
            {
                table.MergeCells(CellRange.Create(table, 1, areaEndCol + 1, 1, numCols - 1));
                table.Cells[1, areaEndCol + 1].TextString = "KHỐI LƯỢNG (m³)";
                table.Cells[1, areaEndCol + 1].Alignment = CellAlignment.MiddleCenter;
            }

            // ===== HÀNG 2: Header chi tiết =====
            table.Cells[2, 0].TextString = "TÊN CỌC";
            table.Cells[2, 1].TextString = "LÝ TRÌNH";
            table.Cells[2, 2].TextString = "K.CÁCH";

            int col = 3;
            foreach (var material in materials)
            {
                table.Cells[2, col].TextString = material;
                table.Cells[2, col].Alignment = CellAlignment.MiddleCenter;
                col++;
            }
            foreach (var material in materials)
            {
                table.Cells[2, col].TextString = material;
                table.Cells[2, col].Alignment = CellAlignment.MiddleCenter;
                col++;
            }

            // Format header
            for (int c = 0; c < numCols; c++)
            {
                table.Cells[2, c].TextHeight = 3.5;
                table.Cells[2, c].Alignment = CellAlignment.MiddleCenter;
            }

            // ===== HÀNG 3+: Dữ liệu =====
            int row = 3;
            Dictionary<string, double> totalVolumes = new();
            foreach (var material in materials)
            {
                totalVolumes[material] = 0;
            }

            foreach (var stake in stakeInfos)
            {
                table.Cells[row, 0].TextString = stake.StakeName;
                table.Cells[row, 1].TextString = stake.Station;
                table.Cells[row, 2].TextString = Math.Round(stake.SpacingPrev, 2).ToString();

                // Diện tích
                col = 3;
                foreach (var material in materials)
                {
                    double area = stake.MaterialAreas.GetValueOrDefault(material, 0);
                    table.Cells[row, col].TextString = Math.Round(area, 3).ToString();
                    table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                    col++;
                }

                // Khối lượng
                foreach (var material in materials)
                {
                    double volume = stake.MaterialVolumes.GetValueOrDefault(material, 0);
                    table.Cells[row, col].TextString = Math.Round(volume, 3).ToString();
                    table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                    totalVolumes[material] += volume;
                    col++;
                }

                row++;
            }

            // ===== HÀNG CUỐI: Tổng cộng =====
            table.MergeCells(CellRange.Create(table, row, 0, row, 2));
            table.Cells[row, 0].TextString = "TỔNG CỘNG";
            table.Cells[row, 0].Alignment = CellAlignment.MiddleCenter;

            // Sum diện tích (để trống hoặc tính tổng)
            col = 3;
            foreach (var material in materials)
            {
                double totalArea = stakeInfos.Sum(s => s.MaterialAreas.GetValueOrDefault(material, 0));
                table.Cells[row, col].TextString = Math.Round(totalArea, 3).ToString();
                table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                col++;
            }

            // Sum khối lượng
            foreach (var material in materials)
            {
                table.Cells[row, col].TextString = Math.Round(totalVolumes[material], 3).ToString();
                table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                col++;
            }

            // Set TextHeight cho tất cả ô
            for (int r = 0; r < numRows; r++)
            {
                for (int c = 0; c < numCols; c++)
                {
                    if (r >= 3) // Dữ liệu
                    {
                        table.Cells[r, c].TextHeight = 3.0;
                    }
                }
            }

            // Thêm table vào model space
            btr.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
        }

        /// <summary>
        /// Command: Thống kê Material theo trắc ngang - Xuất bảng chi tiết
        /// </summary>
        [CommandMethod("CTSV_ThongKeMaterialTracNgang")]
        public static void CTSVThongKeMaterialTracNgang()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                A.Ed.WriteMessage("\n=== THỐNG KÊ MATERIAL THEO TRẮC NGANG ===\n");

                // 1. Lấy danh sách Alignments có SampleLineGroup
                List<AlignmentInfo> alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Không tìm thấy Alignment nào có SampleLineGroup!");
                    A.Ed.WriteMessage("\nVui lòng tạo SampleLineGroup trước.");
                    return;
                }

                // 2. Hiển thị form chọn Alignment
                using FormChonAlignment formChon = new(alignments);
                if (Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(formChon) != DialogResult.OK)
                    return;

                if (formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Chưa chọn Alignment nào!");
                    return;
                }

                // 3. Hỏi loại xuất: CAD Table hay Excel
                PromptKeywordOptions pkOpts = new("\nChọn loại xuất [CAD/Excel]", "CAD Excel");
                pkOpts.Keywords.Default = "Excel";
                pkOpts.AllowNone = true;
                PromptResult pkResult = A.Ed.GetKeywords(pkOpts);

                if (pkResult.Status != PromptStatus.OK && pkResult.Status != PromptStatus.None)
                    return;

                bool exportToExcel = pkResult.StringResult == "Excel" || pkResult.Status == PromptStatus.None;

                // 4. Thu thập dữ liệu Material từ tất cả Alignments
                Dictionary<ObjectId, List<MaterialStatInfo>> alignmentMaterialData = new();
                HashSet<string> allMaterialNames = new();

                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    A.Ed.WriteMessage($"\n📊 Đang xử lý: {alignInfo.Name}...");
                    
                    var materialStats = ExtractMaterialStatistics(tr, alignInfo.SampleLineGroupId, alignInfo.AlignmentId);
                    alignmentMaterialData[alignInfo.AlignmentId] = materialStats;

                    // Thu thập tên materials
                    foreach (var stat in materialStats)
                    {
                        foreach (var mat in stat.MaterialAreas.Keys)
                            allMaterialNames.Add(mat);
                    }
                }

                // 5. Sắp xếp materials
                List<string> orderedMaterials = allMaterialNames.OrderBy(m => m).ToList();

                if (orderedMaterials.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Không tìm thấy Material nào!");
                    A.Ed.WriteMessage("\nVui lòng kiểm tra Material List trong QTO Manager.");
                    return;
                }

                A.Ed.WriteMessage($"\n✅ Tìm thấy {orderedMaterials.Count} loại material:");
                foreach (var mat in orderedMaterials)
                    A.Ed.WriteMessage($"\n  - {mat}");

                // 6. Xuất dữ liệu
                if (exportToExcel)
                {
                    // Xuất Excel
                    SaveFileDialog saveDialog = new()
                    {
                        Title = "Lưu file Excel thống kê Material",
                        Filter = "Excel Files|*.xlsx|All Files|*.*",
                        DefaultExt = "xlsx",
                        FileName = $"ThongKeMaterial_{DateTime.Now:yyyyMMdd_HHmmss}"
                    };

                    if (saveDialog.ShowDialog() != DialogResult.OK)
                        return;

                    ExportMaterialStatisticsToExcel(saveDialog.FileName, formChon.SelectedAlignments, 
                        alignmentMaterialData, orderedMaterials);

                    A.Ed.WriteMessage($"\n\n✅ Đã xuất file Excel: {saveDialog.FileName}");

                    // Hỏi mở file
                    if (MessageBox.Show("Đã xuất file Excel thành công!\nBạn có muốn mở file?", 
                        "Hoàn thành", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = saveDialog.FileName,
                            UseShellExecute = true
                        });
                    }
                }
                else
                {
                    // Xuất CAD Table
                    PromptPointResult ppr = A.Ed.GetPoint("\nChọn điểm chèn bảng: ");
                    if (ppr.Status != PromptStatus.OK)
                        return;

                    Point3d insertPoint = ppr.Value;

                    foreach (var alignInfo in formChon.SelectedAlignments)
                    {
                        var materialStats = alignmentMaterialData[alignInfo.AlignmentId];
                        
                        CreateMaterialStatTable(tr, insertPoint, alignInfo.Name, materialStats, orderedMaterials);
                        
                        A.Ed.WriteMessage($"\n✅ Đã vẽ bảng cho '{alignInfo.Name}'");

                        // Offset cho bảng tiếp theo
                        double tableHeight = (materialStats.Count + 4) * 8.0;
                        insertPoint = new Point3d(insertPoint.X, insertPoint.Y - tableHeight - 50, insertPoint.Z);
                    }
                }

                tr.Commit();
                A.Ed.WriteMessage("\n\n=== HOÀN THÀNH ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
                A.Ed.WriteMessage($"\nStack: {ex.StackTrace}");
            }
        }

        #region Material Statistics Classes and Methods

        /// <summary>
        /// Thông tin thống kê material tại một trắc ngang
        /// </summary>
        public class MaterialStatInfo
        {
            public string SampleLineName { get; set; } = "";
            public double Station { get; set; }
            public string StationFormatted { get; set; } = "";
            public double SpacingPrev { get; set; }
            
            // Diện tích theo vật liệu
            public Dictionary<string, double> MaterialAreas { get; set; } = new();
            
            // Chi tiết từng material (Left/Right offset, Min/Max elevation)
            public Dictionary<string, MaterialDetailInfo> MaterialDetails { get; set; } = new();
        }

        /// <summary>
        /// Chi tiết về một material tại trắc ngang
        /// </summary>
        public class MaterialDetailInfo
        {
            public double Area { get; set; }
            public double LeftOffset { get; set; }
            public double RightOffset { get; set; }
            public double MinElevation { get; set; }
            public double MaxElevation { get; set; }
            public int PointCount { get; set; }
        }

        /// <summary>
        /// Trích xuất thống kê Material từ SampleLineGroup
        /// </summary>
        private static List<MaterialStatInfo> ExtractMaterialStatistics(
            Transaction tr, ObjectId sampleLineGroupId, ObjectId alignmentId)
        {
            List<MaterialStatInfo> stats = new();

            SampleLineGroup? slg = tr.GetObject(sampleLineGroupId, AcadDb.OpenMode.ForWrite, false, true) as SampleLineGroup;
            if (slg == null) return stats;

            Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
            if (alignment == null) return stats;

            // Lấy Material Lists
            QTOMaterialListCollection materialLists = slg.MaterialLists;
            if (materialLists.Count == 0)
            {
                A.Ed.WriteMessage($"\n⚠️ Không có Material List trong '{slg.Name}'");
                return stats;
            }

            // Thu thập Materials với GUID
            List<(string Name, Guid MaterialListGuid, Guid MaterialGuid)> materials = new();
            foreach (QTOMaterialList materialList in materialLists)
            {
                try
                {
                    Guid listGuid = materialList.Guid;
                    foreach (QTOMaterial material in materialList)
                    {
                        materials.Add((material.Name, listGuid, material.Guid));
                    }
                }
                catch { }
            }

            // Sắp xếp SampleLines theo lý trình
            List<SampleLine> sortedSampleLines = new();
            foreach (ObjectId slId in slg.GetSampleLineIds())
            {
                var sl = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                if (sl != null) sortedSampleLines.Add(sl);
            }
            sortedSampleLines = sortedSampleLines.OrderBy(s => s.Station).ToList();

            // Duyệt qua từng SampleLine
            double prevStation = 0;
            bool isFirst = true;

            foreach (SampleLine sampleLine in sortedSampleLines)
            {
                MaterialStatInfo statInfo = new()
                {
                    SampleLineName = sampleLine.Name,
                    Station = sampleLine.Station,
                    StationFormatted = FormatStation(sampleLine.Station),
                    SpacingPrev = isFirst ? 0 : (sampleLine.Station - prevStation)
                };

                // Lấy MaterialSection cho từng Material
                foreach (var (materialName, materialListGuid, materialGuid) in materials)
                {
                    try
                    {
                        ObjectId materialSectionId = sampleLine.GetMaterialSectionId(materialListGuid, materialGuid);

                        if (!materialSectionId.IsNull && materialSectionId.IsValid)
                        {
                            AcadDb.DBObject? sectionObj = tr.GetObject(materialSectionId, AcadDb.OpenMode.ForRead, false, true);

                            if (sectionObj is CivSection section)
                            {
                                // Lấy diện tích
                                double areaFromAPI = section.Area;
                                double areaCalculated = CalculateSectionArea(section);
                                double area = areaFromAPI > 0 ? areaFromAPI : areaCalculated;

                                if (area > 0)
                                {
                                    statInfo.MaterialAreas[materialName] = area;

                                    // Lấy chi tiết material
                                    MaterialDetailInfo detail = new() { Area = area };

                                    try
                                    {
                                        SectionPointCollection points = section.SectionPoints;
                                        detail.PointCount = points.Count;

                                        if (points.Count > 0)
                                        {
                                            double minX = double.MaxValue, maxX = double.MinValue;
                                            double minY = double.MaxValue, maxY = double.MinValue;

                                            foreach (SectionPoint pt in points)
                                            {
                                                if (pt.Location.X < minX) minX = pt.Location.X;
                                                if (pt.Location.X > maxX) maxX = pt.Location.X;
                                                if (pt.Location.Y < minY) minY = pt.Location.Y;
                                                if (pt.Location.Y > maxY) maxY = pt.Location.Y;
                                            }

                                            detail.LeftOffset = minX;
                                            detail.RightOffset = maxX;
                                            detail.MinElevation = minY;
                                            detail.MaxElevation = maxY;
                                        }
                                    }
                                    catch { }

                                    statInfo.MaterialDetails[materialName] = detail;
                                }
                            }
                        }
                    }
                    catch { }
                }

                stats.Add(statInfo);
                prevStation = sampleLine.Station;
                isFirst = false;
            }

            return stats;
        }

        /// <summary>
        /// Xuất thống kê Material ra Excel
        /// </summary>
        private static void ExportMaterialStatisticsToExcel(
            string filePath,
            List<AlignmentInfo> alignments,
            Dictionary<ObjectId, List<MaterialStatInfo>> alignmentMaterialData,
            List<string> materials)
        {
            using var workbook = new XLWorkbook();

            // Sheet cho từng Alignment
            foreach (var alignInfo in alignments)
            {
                var stats = alignmentMaterialData[alignInfo.AlignmentId];
                string sheetName = SanitizeSheetName($"TN_{alignInfo.Name}");
                var ws = workbook.Worksheets.Add(sheetName);

                // Tiêu đề
                ws.Cell(1, 1).Value = $"BẢNG THỐNG KÊ MATERIAL THEO TRẮC NGANG - {alignInfo.Name}";
                ws.Range(1, 1, 1, 4 + materials.Count * 4).Merge();
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Font.FontSize = 14;
                ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Header hàng 2 - Nhóm
                ws.Cell(2, 1).Value = "THÔNG TIN TRẮC NGANG";
                ws.Range(2, 1, 2, 4).Merge();
                ws.Cell(2, 1).Style.Font.Bold = true;
                ws.Cell(2, 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                ws.Cell(2, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int col = 5;
                foreach (var mat in materials)
                {
                    ws.Cell(2, col).Value = mat;
                    ws.Range(2, col, 2, col + 3).Merge();
                    ws.Cell(2, col).Style.Font.Bold = true;
                    ws.Cell(2, col).Style.Fill.BackgroundColor = XLColor.LightGreen;
                    ws.Cell(2, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    col += 4;
                }

                // Header hàng 3 - Chi tiết
                ws.Cell(3, 1).Value = "STT";
                ws.Cell(3, 2).Value = "TÊN TRẮC NGANG";
                ws.Cell(3, 3).Value = "LÝ TRÌNH";
                ws.Cell(3, 4).Value = "K.CÁCH (m)";

                col = 5;
                foreach (var mat in materials)
                {
                    ws.Cell(3, col).Value = "Diện tích (m²)";
                    ws.Cell(3, col + 1).Value = "Offset trái";
                    ws.Cell(3, col + 2).Value = "Offset phải";
                    ws.Cell(3, col + 3).Value = "Cao độ";
                    col += 4;
                }

                // Format header hàng 3
                int lastCol = 4 + materials.Count * 4;
                ws.Range(3, 1, 3, lastCol).Style.Font.Bold = true;
                ws.Range(3, 1, 3, lastCol).Style.Fill.BackgroundColor = XLColor.LightBlue;
                ws.Range(3, 1, 3, lastCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Dữ liệu
                int row = 4;
                int stt = 1;
                Dictionary<string, double> totalAreas = materials.ToDictionary(m => m, m => 0.0);

                foreach (var stat in stats)
                {
                    ws.Cell(row, 1).Value = stt++;
                    ws.Cell(row, 2).Value = stat.SampleLineName;
                    ws.Cell(row, 3).Value = stat.StationFormatted;
                    ws.Cell(row, 4).Value = Math.Round(stat.SpacingPrev, 3);

                    col = 5;
                    foreach (var mat in materials)
                    {
                        double area = stat.MaterialAreas.GetValueOrDefault(mat, 0);
                        ws.Cell(row, col).Value = Math.Round(area, 4);
                        totalAreas[mat] += area;

                        if (stat.MaterialDetails.ContainsKey(mat))
                        {
                            var detail = stat.MaterialDetails[mat];
                            ws.Cell(row, col + 1).Value = Math.Round(detail.LeftOffset, 3);
                            ws.Cell(row, col + 2).Value = Math.Round(detail.RightOffset, 3);
                            ws.Cell(row, col + 3).Value = $"{detail.MinElevation:F2} ~ {detail.MaxElevation:F2}";
                        }

                        col += 4;
                    }

                    row++;
                }

                // Hàng tổng cộng
                ws.Cell(row, 1).Value = "TỔNG CỘNG";
                ws.Range(row, 1, row, 4).Merge();
                ws.Cell(row, 1).Style.Font.Bold = true;
                ws.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                col = 5;
                foreach (var mat in materials)
                {
                    ws.Cell(row, col).Value = Math.Round(totalAreas[mat], 4);
                    ws.Cell(row, col).Style.Font.Bold = true;
                    col += 4;
                }

                // Border và format
                ws.Range(2, 1, row, lastCol).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(2, 1, row, lastCol).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                ws.Column(1).Width = 8;
                ws.Column(2).Width = 20;
                ws.Column(3).Width = 15;
                ws.Column(4).Width = 12;

                for (int c = 5; c <= lastCol; c++)
                {
                    ws.Column(c).Width = 15;
                }
            }

            // Sheet tổng hợp
            if (alignments.Count > 1)
            {
                var wsSummary = workbook.Worksheets.Add("TỔNG HỢP");
                
                wsSummary.Cell(1, 1).Value = "TỔNG HỢP DIỆN TÍCH MATERIAL TẤT CẢ CÁC TUYẾN";
                wsSummary.Range(1, 1, 1, 1 + materials.Count).Merge();
                wsSummary.Cell(1, 1).Style.Font.Bold = true;
                wsSummary.Cell(1, 1).Style.Font.FontSize = 14;
                wsSummary.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Header
                wsSummary.Cell(2, 1).Value = "TUYẾN";
                int col = 2;
                foreach (var mat in materials)
                {
                    wsSummary.Cell(2, col).Value = $"{mat} (m²)";
                    col++;
                }

                wsSummary.Range(2, 1, 2, 1 + materials.Count).Style.Font.Bold = true;
                wsSummary.Range(2, 1, 2, 1 + materials.Count).Style.Fill.BackgroundColor = XLColor.LightGreen;

                // Dữ liệu
                int row = 3;
                Dictionary<string, double> grandTotals = materials.ToDictionary(m => m, m => 0.0);

                foreach (var alignInfo in alignments)
                {
                    var stats = alignmentMaterialData[alignInfo.AlignmentId];
                    wsSummary.Cell(row, 1).Value = alignInfo.Name;

                    col = 2;
                    foreach (var mat in materials)
                    {
                        double total = stats.Sum(s => s.MaterialAreas.GetValueOrDefault(mat, 0));
                        wsSummary.Cell(row, col).Value = Math.Round(total, 4);
                        grandTotals[mat] += total;
                        col++;
                    }

                    row++;
                }

                // Tổng cộng
                wsSummary.Cell(row, 1).Value = "TỔNG CỘNG";
                wsSummary.Cell(row, 1).Style.Font.Bold = true;

                col = 2;
                foreach (var mat in materials)
                {
                    wsSummary.Cell(row, col).Value = Math.Round(grandTotals[mat], 4);
                    wsSummary.Cell(row, col).Style.Font.Bold = true;
                    col++;
                }

                // Format
                wsSummary.Range(2, 1, row, 1 + materials.Count).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                wsSummary.Range(2, 1, row, 1 + materials.Count).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                wsSummary.Column(1).Width = 25;
                for (int c = 2; c <= 1 + materials.Count; c++)
                {
                    wsSummary.Column(c).Width = 18;
                }
            }

            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Tạo bảng thống kê Material trong CAD
        /// </summary>
        private static void CreateMaterialStatTable(
            Transaction tr,
            Point3d insertPoint,
            string alignmentName,
            List<MaterialStatInfo> stats,
            List<string> materials)
        {
            AcadDb.Database db = HostApplicationServices.WorkingDatabase;
            AcadDb.BlockTable bt = tr.GetObject(db.BlockTableId, AcadDb.OpenMode.ForRead) as AcadDb.BlockTable
                ?? throw new System.Exception("Không thể mở BlockTable");
            AcadDb.BlockTableRecord btr = tr.GetObject(bt[AcadDb.BlockTableRecord.ModelSpace], AcadDb.OpenMode.ForWrite) as AcadDb.BlockTableRecord
                ?? throw new System.Exception("Không thể mở ModelSpace");

            // Tính số cột và hàng
            int numCols = 4 + materials.Count; // STT, Tên, Lý trình, K.cách + Materials
            int numRows = stats.Count + 4; // 2 header + dữ liệu + 1 tổng cộng

            // Tạo Table
            AcadDb.Table table = new()
            {
                Position = insertPoint,
                TableStyle = db.Tablestyle
            };

            table.SetSize(numRows, numCols);

            // Kích thước hàng/cột
            double rowHeight = 8.0;
            for (int r = 0; r < numRows; r++)
            {
                table.Rows[r].Height = rowHeight;
            }

            table.Columns[0].Width = 10;    // STT
            table.Columns[1].Width = 25;    // Tên
            table.Columns[2].Width = 20;    // Lý trình
            table.Columns[3].Width = 15;    // K.cách

            for (int c = 4; c < numCols; c++)
            {
                table.Columns[c].Width = 20;
            }

            // HÀNG 0: Tiêu đề
            table.MergeCells(CellRange.Create(table, 0, 0, 0, numCols - 1));
            table.Cells[0, 0].TextString = $"THỐNG KÊ MATERIAL - {alignmentName}";
            table.Cells[0, 0].TextHeight = 5.0;
            table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

            // HÀNG 1: Nhóm header
            table.MergeCells(CellRange.Create(table, 1, 0, 1, 3));
            table.Cells[1, 0].TextString = "THÔNG TIN TRẮC NGANG";
            table.Cells[1, 0].Alignment = CellAlignment.MiddleCenter;

            if (materials.Count > 0)
            {
                table.MergeCells(CellRange.Create(table, 1, 4, 1, numCols - 1));
                table.Cells[1, 4].TextString = "DIỆN TÍCH VẬT LIỆU (m²)";
                table.Cells[1, 4].Alignment = CellAlignment.MiddleCenter;
            }

            // HÀNG 2: Header chi tiết
            table.Cells[2, 0].TextString = "STT";
            table.Cells[2, 1].TextString = "TÊN TRẮC NGANG";
            table.Cells[2, 2].TextString = "LÝ TRÌNH";
            table.Cells[2, 3].TextString = "K.CÁCH";

            int col = 4;
            foreach (var mat in materials)
            {
                table.Cells[2, col].TextString = mat;
                table.Cells[2, col].Alignment = CellAlignment.MiddleCenter;
                col++;
            }

            // Format header
            for (int c = 0; c < numCols; c++)
            {
                table.Cells[2, c].TextHeight = 3.5;
                table.Cells[2, c].Alignment = CellAlignment.MiddleCenter;
            }

            // DỮ LIỆU
            int row = 3;
            int stt = 1;
            Dictionary<string, double> totals = materials.ToDictionary(m => m, m => 0.0);

            foreach (var stat in stats)
            {
                table.Cells[row, 0].TextString = stt++.ToString();
                table.Cells[row, 1].TextString = stat.SampleLineName;
                table.Cells[row, 2].TextString = stat.StationFormatted;
                table.Cells[row, 3].TextString = Math.Round(stat.SpacingPrev, 2).ToString();

                col = 4;
                foreach (var mat in materials)
                {
                    double area = stat.MaterialAreas.GetValueOrDefault(mat, 0);
                    table.Cells[row, col].TextString = Math.Round(area, 4).ToString();
                    table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                    totals[mat] += area;
                    col++;
                }

                row++;
            }

            // HÀNG TỔNG CỘNG
            table.MergeCells(CellRange.Create(table, row, 0, row, 3));
            table.Cells[row, 0].TextString = "TỔNG CỘNG";
            table.Cells[row, 0].Alignment = CellAlignment.MiddleCenter;

            col = 4;
            foreach (var mat in materials)
            {
                table.Cells[row, col].TextString = Math.Round(totals[mat], 4).ToString();
                table.Cells[row, col].Alignment = CellAlignment.MiddleRight;
                col++;
            }

            // TextHeight cho dữ liệu
            for (int r = 3; r < numRows; r++)
            {
                for (int c = 0; c < numCols; c++)
                {
                    table.Cells[r, c].TextHeight = 3.0;
                }
            }

            // Thêm table vào model space
            btr.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
        }

        #endregion

        #region Tính diện tích từ Polyline vẽ thủ công

        /// <summary>
        /// Lệnh tích hợp tất cả chức năng tính diện tích/khối lượng từ Polyline
        /// Chạy tuần tự: Chọn polyline → Tính diện tích → Tính khối lượng → Ghi lên bản vẽ → Xuất Excel
        /// </summary>
        [CommandMethod("CTSV_PolyArea")]
        public static void CTSVPolyArea()
        {
            try
            {
                A.Ed.WriteMessage("\n╔══════════════════════════════════════════════════════════════╗");
                A.Ed.WriteMessage("\n║   TÍNH KHỐI LƯỢNG TỪ POLYLINE - WORKFLOW TỰ ĐỘNG             ║");
                A.Ed.WriteMessage("\n║   (Chọn polyline 1 lần → Chạy tuần tự 4 bước)                 ║");
                A.Ed.WriteMessage("\n╚══════════════════════════════════════════════════════════════╝");

                // ===== BƯỚC 1: Nhập thông tin =====
                A.Ed.WriteMessage("\n\n📋 BƯỚC 1: NHẬP THÔNG TIN");
                A.Ed.WriteMessage("\n" + new string('─', 50));

                // Nhập tên vật liệu
                PromptStringOptions psoName = new PromptStringOptions("\nNhập tên vật liệu/loại diện tích:")
                {
                    DefaultValue = "Đào nền",
                    AllowSpaces = true
                };
                PromptResult prName = A.Ed.GetString(psoName);
                if (prName.Status != PromptStatus.OK) return;
                string materialName = prName.StringResult;

                // Nhập khoảng cách giữa các mặt cắt
                PromptDoubleOptions pdoSpacing = new PromptDoubleOptions("\nNhập khoảng cách giữa các mặt cắt (m):")
                {
                    AllowNegative = false,
                    AllowZero = false,
                    DefaultValue = 20.0
                };
                PromptDoubleResult pdrSpacing = A.Ed.GetDouble(pdoSpacing);
                if (pdrSpacing.Status != PromptStatus.OK) return;
                double spacing = pdrSpacing.Value;

                // Hỏi có ghi diện tích lên polyline không
                PromptKeywordOptions pkoLabel = new PromptKeywordOptions("\nGhi nhãn diện tích lên polyline? [Yes/No]:")
                {
                    AllowNone = false
                };
                pkoLabel.Keywords.Add("Yes");
                pkoLabel.Keywords.Add("No");
                pkoLabel.Keywords.Default = "Yes";
                PromptResult prLabel = A.Ed.GetKeywords(pkoLabel);
                bool writeLabels = prLabel.Status == PromptStatus.OK && prLabel.StringResult == "Yes";

                double textHeight = 2.5;
                if (writeLabels)
                {
                    PromptDoubleOptions pdoHeight = new PromptDoubleOptions("\nNhập chiều cao text:")
                    {
                        AllowNegative = false,
                        AllowZero = false,
                        DefaultValue = 2.5
                    };
                    PromptDoubleResult pdrHeight = A.Ed.GetDouble(pdoHeight);
                    if (pdrHeight.Status == PromptStatus.OK)
                        textHeight = pdrHeight.Value;
                }

                // ===== BƯỚC 2: Chọn polylines =====
                A.Ed.WriteMessage("\n\n📋 BƯỚC 2: CHỌN POLYLINE");
                A.Ed.WriteMessage("\n" + new string('─', 50));
                A.Ed.WriteMessage("\n⚠️ Chọn theo THỨ TỰ mặt cắt từ đầu đến cuối!");

                TypedValue[] filterList = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<OR"),
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                    new TypedValue((int)DxfCode.Start, "POLYLINE"),
                    new TypedValue((int)DxfCode.Operator, "OR>")
                };
                SelectionFilter filter = new SelectionFilter(filterList);

                PromptSelectionResult psr = A.Ed.GetSelection(filter);
                if (psr.Status != PromptStatus.OK) return;

                SelectionSet ss = psr.Value;
                if (ss.Count < 2)
                {
                    A.Ed.WriteMessage("\n❌ Cần ít nhất 2 polyline để tính khối lượng!");
                    return;
                }

                A.Ed.WriteMessage($"\n✅ Đã chọn {ss.Count} polyline");

                // ===== BƯỚC 3: Tính toán diện tích và khối lượng =====
                A.Ed.WriteMessage("\n\n📋 BƯỚC 3: TÍNH TOÁN");
                A.Ed.WriteMessage("\n" + new string('─', 50));

                List<PolyAreaData> dataList = new List<PolyAreaData>();
                List<ObjectId> polylineIds = new List<ObjectId>();
                double totalVolume = 0;
                double prevArea = 0;
                double totalArea = 0;

                using (Transaction tr = A.Db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(A.Db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    int index = 1;
                    foreach (ObjectId id in ss.GetObjectIds())
                    {
                        polylineIds.Add(id);
                        Autodesk.AutoCAD.DatabaseServices.Entity ent = tr.GetObject(id, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                        
                        double area = 0;
                        Point3d centroid = Point3d.Origin;

                        if (ent is Polyline pline)
                        {
                            area = pline.Closed ? pline.Area : CalculateOpenPolylineArea(pline);
                            centroid = GetPolylineCentroid(pline);

                            // Ghi nhãn nếu được yêu cầu
                            if (writeLabels)
                            {
                                DBText areaText = new DBText
                                {
                                    Position = centroid,
                                    TextString = $"S={Math.Round(area, AreaDecimalPlaces):F2}",
                                    Height = textHeight,
                                    Layer = pline.Layer
                                };
                                btr.AppendEntity(areaText);
                                tr.AddNewlyCreatedDBObject(areaText, true);
                            }
                        }
                        else if (ent is Curve curve)
                        {
                            area = curve.Area;
                        }

                        area = Math.Round(area, AreaDecimalPlaces);
                        totalArea += area;
                        
                        double segmentVolume = 0;
                        if (index > 1)
                        {
                            segmentVolume = CalculateVolume(prevArea, area, spacing, true);
                            totalVolume += segmentVolume;
                        }

                        dataList.Add(new PolyAreaData
                        {
                            STT = index,
                            SectionName = $"MC{index}",
                            Station = (index - 1) * spacing,
                            Area = area,
                            SegmentVolume = segmentVolume,
                            CumulativeVolume = Math.Round(totalVolume, VolumeDecimalPlaces)
                        });

                        prevArea = area;
                        index++;
                    }

                    tr.Commit();
                }

                // ===== BƯỚC 4: Hiển thị kết quả =====
                A.Ed.WriteMessage("\n\n📋 BƯỚC 4: KẾT QUẢ TÍNH TOÁN");
                A.Ed.WriteMessage("\n" + new string('─', 50));
                A.Ed.WriteMessage($"\n   Vật liệu: {materialName}");
                A.Ed.WriteMessage($"\n   Khoảng cách mặt cắt: {spacing} m");
                A.Ed.WriteMessage($"\n   Công thức: V = (S1 + S2) / 2 × L");
                A.Ed.WriteMessage($"\n\n   {"MC",-6} {"Lý trình",-12} {"S (m²)",-12} {"V đoạn (m³)",-14} {"V cộng dồn",-12}");
                A.Ed.WriteMessage($"\n   {new string('─', 56)}");

                foreach (var data in dataList)
                {
                    A.Ed.WriteMessage($"\n   {data.SectionName,-6} {data.Station,-12:F2} {data.Area,-12:F2} {data.SegmentVolume,-14:F2} {data.CumulativeVolume,-12:F2}");
                }

                A.Ed.WriteMessage($"\n   {new string('─', 56)}");
                A.Ed.WriteMessage($"\n   {"TỔNG",-6} {"",-12} {totalArea,-12:F2} {Math.Round(totalVolume, VolumeDecimalPlaces),-14:F2} {Math.Round(totalVolume, VolumeDecimalPlaces),-12:F2}");

                // ===== BƯỚC 5: Xuất Excel =====
                A.Ed.WriteMessage("\n\n📋 BƯỚC 5: XUẤT FILE EXCEL");
                A.Ed.WriteMessage("\n" + new string('─', 50));

                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string safeMatName = string.Join("_", materialName.Split(Path.GetInvalidFileNameChars()));
                string fileName = $"KhoiLuong_{safeMatName}_{timestamp}.xlsx";
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filePath = Path.Combine(desktopPath, fileName);

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Khối lượng");

                    // Tiêu đề
                    worksheet.Cell(1, 1).Value = $"BẢNG TÍNH KHỐI LƯỢNG - {materialName.ToUpper()}";
                    worksheet.Range(1, 1, 1, 6).Merge();
                    worksheet.Cell(1, 1).Style.Font.Bold = true;
                    worksheet.Cell(1, 1).Style.Font.FontSize = 14;
                    worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    worksheet.Cell(2, 1).Value = $"Khoảng cách mặt cắt: {spacing} m | Phương pháp: Average End Area";
                    worksheet.Range(2, 1, 2, 6).Merge();

                    worksheet.Cell(3, 1).Value = $"Ngày xuất: {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
                    worksheet.Range(3, 1, 3, 6).Merge();

                    // Header
                    worksheet.Cell(5, 1).Value = "STT";
                    worksheet.Cell(5, 2).Value = "Mặt cắt";
                    worksheet.Cell(5, 3).Value = "Lý trình (m)";
                    worksheet.Cell(5, 4).Value = "Diện tích (m²)";
                    worksheet.Cell(5, 5).Value = "KL đoạn (m³)";
                    worksheet.Cell(5, 6).Value = "KL cộng dồn (m³)";

                    var headerRange = worksheet.Range(5, 1, 5, 6);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    // Data
                    int row = 6;
                    foreach (var data in dataList)
                    {
                        worksheet.Cell(row, 1).Value = data.STT;
                        worksheet.Cell(row, 2).Value = data.SectionName;
                        worksheet.Cell(row, 3).Value = data.Station;
                        worksheet.Cell(row, 4).Value = data.Area;
                        worksheet.Cell(row, 5).Value = data.SegmentVolume;
                        worksheet.Cell(row, 6).Value = data.CumulativeVolume;

                        // Border
                        worksheet.Range(row, 1, row, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        worksheet.Range(row, 1, row, 6).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        row++;
                    }

                    // Hàng tổng
                    worksheet.Cell(row, 1).Value = "TỔNG CỘNG";
                    worksheet.Range(row, 1, row, 3).Merge();
                    worksheet.Cell(row, 1).Style.Font.Bold = true;
                    worksheet.Cell(row, 4).Value = Math.Round(totalArea, AreaDecimalPlaces);
                    worksheet.Cell(row, 4).Style.Font.Bold = true;
                    worksheet.Cell(row, 5).Value = Math.Round(totalVolume, VolumeDecimalPlaces);
                    worksheet.Cell(row, 5).Style.Font.Bold = true;
                    worksheet.Cell(row, 6).Value = Math.Round(totalVolume, VolumeDecimalPlaces);
                    worksheet.Cell(row, 6).Style.Font.Bold = true;
                    worksheet.Range(row, 1, row, 6).Style.Fill.BackgroundColor = XLColor.LightGreen;
                    worksheet.Range(row, 1, row, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;

                    // Định dạng cột
                    worksheet.Columns().AdjustToContents();

                    workbook.SaveAs(filePath);
                }

                A.Ed.WriteMessage($"\n✅ Đã xuất file: {filePath}");

                // ===== TỔNG KẾT =====
                A.Ed.WriteMessage("\n\n╔══════════════════════════════════════════════════════════════╗");
                A.Ed.WriteMessage("\n║                    📊 TỔNG KẾT                               ║");
                A.Ed.WriteMessage("\n╠══════════════════════════════════════════════════════════════╣");
                A.Ed.WriteMessage($"\n║   Vật liệu:         {materialName,-40} ║");
                A.Ed.WriteMessage($"\n║   Số mặt cắt:       {dataList.Count,-40} ║");
                A.Ed.WriteMessage($"\n║   Tổng diện tích:   {Math.Round(totalArea, AreaDecimalPlaces):F2} m²{new string(' ', 32)}║");
                A.Ed.WriteMessage($"\n║   TỔNG KHỐI LƯỢNG:  {Math.Round(totalVolume, VolumeDecimalPlaces):F2} m³{new string(' ', 32)}║");
                A.Ed.WriteMessage("\n╚══════════════════════════════════════════════════════════════╝");

                if (writeLabels)
                    A.Ed.WriteMessage($"\n✅ Đã ghi {dataList.Count} nhãn diện tích lên bản vẽ");

                A.Ed.WriteMessage("\n\n=== WORKFLOW HOÀN TẤT ===\n");

                // Mở file Excel
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Tính diện tích polyline và xuất Excel
        /// </summary>
        [CommandMethod("CTSV_TinhDienTichPolyExcel")]
        public static void CTSVTinhDienTichPolyExcel()
        {
            try
            {
                A.Ed.WriteMessage("\n=== TÍNH DIỆN TÍCH POLYLINE + XUẤT EXCEL ===");

                // Nhập tên vật liệu/loại diện tích
                PromptStringOptions psoName = new PromptStringOptions("\nNhập tên vật liệu/loại diện tích:")
                {
                    DefaultValue = "Vật liệu",
                    AllowSpaces = true
                };
                PromptResult prName = A.Ed.GetString(psoName);
                if (prName.Status != PromptStatus.OK) return;
                string materialName = prName.StringResult;

                // Nhập khoảng cách giữa các mặt cắt
                PromptDoubleOptions pdoSpacing = new PromptDoubleOptions("\nNhập khoảng cách giữa các mặt cắt (m):")
                {
                    AllowNegative = false,
                    AllowZero = false,
                    DefaultValue = 20.0
                };
                PromptDoubleResult pdrSpacing = A.Ed.GetDouble(pdoSpacing);
                if (pdrSpacing.Status != PromptStatus.OK) return;
                double spacing = pdrSpacing.Value;

                // Chọn polylines
                TypedValue[] filterList = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<OR"),
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                    new TypedValue((int)DxfCode.Start, "POLYLINE"),
                    new TypedValue((int)DxfCode.Operator, "OR>")
                };
                SelectionFilter filter = new SelectionFilter(filterList);

                A.Ed.WriteMessage("\n📌 Chọn các polyline theo thứ tự mặt cắt:");
                PromptSelectionResult psr = A.Ed.GetSelection(filter);
                if (psr.Status != PromptStatus.OK) return;

                SelectionSet ss = psr.Value;
                if (ss.Count < 1)
                {
                    A.Ed.WriteMessage("\n❌ Không có polyline nào được chọn!");
                    return;
                }

                A.Ed.WriteMessage($"\n✅ Đã chọn {ss.Count} polyline");

                // Lấy diện tích
                List<PolyAreaData> dataList = new List<PolyAreaData>();
                double totalVolume = 0;
                double prevArea = 0;

                using (Transaction tr = A.Db.TransactionManager.StartTransaction())
                {
                    int index = 1;
                    foreach (ObjectId id in ss.GetObjectIds())
                    {
                        Autodesk.AutoCAD.DatabaseServices.Entity ent = tr.GetObject(id, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                        
                        double area = 0;
                        if (ent is Polyline pline)
                        {
                            area = pline.Closed ? pline.Area : CalculateOpenPolylineArea(pline);
                        }
                        else if (ent is Curve curve)
                        {
                            area = curve.Area;
                        }

                        area = Math.Round(area, AreaDecimalPlaces);
                        
                        double segmentVolume = 0;
                        if (index > 1)
                        {
                            segmentVolume = CalculateVolume(prevArea, area, spacing, true);
                            totalVolume += segmentVolume;
                        }

                        dataList.Add(new PolyAreaData
                        {
                            STT = index,
                            SectionName = $"MC{index}",
                            Station = (index - 1) * spacing,
                            Area = area,
                            SegmentVolume = segmentVolume,
                            CumulativeVolume = Math.Round(totalVolume, VolumeDecimalPlaces)
                        });

                        prevArea = area;
                        index++;
                    }

                    tr.Commit();
                }

                // Xuất Excel
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fileName = $"KhoiLuong_Poly_{materialName}_{timestamp}.xlsx";
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filePath = Path.Combine(desktopPath, fileName);

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Khối lượng");

                    // Tiêu đề
                    worksheet.Cell(1, 1).Value = $"BẢNG TÍNH KHỐI LƯỢNG - {materialName.ToUpper()}";
                    worksheet.Range(1, 1, 1, 6).Merge();
                    worksheet.Cell(1, 1).Style.Font.Bold = true;
                    worksheet.Cell(1, 1).Style.Font.FontSize = 14;
                    worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    worksheet.Cell(2, 1).Value = $"Khoảng cách mặt cắt: {spacing} m";
                    worksheet.Range(2, 1, 2, 6).Merge();

                    // Header
                    worksheet.Cell(4, 1).Value = "STT";
                    worksheet.Cell(4, 2).Value = "Mặt cắt";
                    worksheet.Cell(4, 3).Value = "Lý trình (m)";
                    worksheet.Cell(4, 4).Value = "Diện tích (m²)";
                    worksheet.Cell(4, 5).Value = "KL đoạn (m³)";
                    worksheet.Cell(4, 6).Value = "KL cộng dồn (m³)";

                    // Định dạng header
                    var headerRange = worksheet.Range(4, 1, 4, 6);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    // Data
                    int row = 5;
                    foreach (var data in dataList)
                    {
                        worksheet.Cell(row, 1).Value = data.STT;
                        worksheet.Cell(row, 2).Value = data.SectionName;
                        worksheet.Cell(row, 3).Value = data.Station;
                        worksheet.Cell(row, 4).Value = data.Area;
                        worksheet.Cell(row, 5).Value = data.SegmentVolume;
                        worksheet.Cell(row, 6).Value = data.CumulativeVolume;
                        row++;
                    }

                    // Hàng tổng
                    worksheet.Cell(row, 1).Value = "TỔNG";
                    worksheet.Range(row, 1, row, 4).Merge();
                    worksheet.Cell(row, 1).Style.Font.Bold = true;
                    worksheet.Cell(row, 5).Value = Math.Round(totalVolume, VolumeDecimalPlaces);
                    worksheet.Cell(row, 5).Style.Font.Bold = true;
                    worksheet.Cell(row, 6).Value = Math.Round(totalVolume, VolumeDecimalPlaces);
                    worksheet.Cell(row, 6).Style.Font.Bold = true;

                    // Định dạng cột
                    worksheet.Columns().AdjustToContents();

                    workbook.SaveAs(filePath);
                }

                A.Ed.WriteMessage($"\n\n📊 TỔNG KẾT:");
                A.Ed.WriteMessage($"\n   Số mặt cắt: {dataList.Count}");
                A.Ed.WriteMessage($"\n   Tổng khối lượng: {Math.Round(totalVolume, VolumeDecimalPlaces):F2} m³");
                A.Ed.WriteMessage($"\n\n📁 Đã xuất file: {filePath}");
                A.Ed.WriteMessage("\n=== HOÀN TẤT ===\n");

                // Mở file Excel
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Class lưu dữ liệu diện tích polyline
        /// </summary>
        public class PolyAreaData
        {
            public int STT { get; set; }
            public string SectionName { get; set; } = "";
            public double Station { get; set; }
            public double Area { get; set; }
            public double SegmentVolume { get; set; }
            public double CumulativeVolume { get; set; }
        }

        /// <summary>
        /// Lệnh tính diện tích từ các polyline được chọn
        /// Tương tự cách Toolcivil2025 sử dụng ((Curve)polyline).Area
        /// </summary>
        [CommandMethod("CTSV_TinhDienTichPoly")]
        public static void CTSVTinhDienTichPoly()
        {
            try
            {
                A.Ed.WriteMessage("\n=== TÍNH DIỆN TÍCH TỪ POLYLINE VẼ THỦ CÔNG ===");
                A.Ed.WriteMessage("\n📌 Chọn các polyline đã vẽ bao quanh vùng diện tích:");

                // Chọn polylines
                PromptSelectionOptions pso = new PromptSelectionOptions();
                pso.MessageForAdding = "\nChọn các polyline (Closed hoặc Open):";
                pso.AllowDuplicates = false;

                TypedValue[] filterList = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<OR"),
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                    new TypedValue((int)DxfCode.Start, "POLYLINE"),
                    new TypedValue((int)DxfCode.Operator, "OR>")
                };
                SelectionFilter filter = new SelectionFilter(filterList);

                PromptSelectionResult psr = A.Ed.GetSelection(pso, filter);
                if (psr.Status != PromptStatus.OK)
                {
                    A.Ed.WriteMessage("\n❌ Không có polyline nào được chọn.");
                    return;
                }

                SelectionSet ss = psr.Value;
                A.Ed.WriteMessage($"\n✅ Đã chọn {ss.Count} polyline");

                double totalArea = 0;
                int count = 0;
                List<PolylineAreaInfo> areaInfos = new List<PolylineAreaInfo>();

                using (Transaction tr = A.Db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in ss.GetObjectIds())
                    {
                        Autodesk.AutoCAD.DatabaseServices.Entity ent = tr.GetObject(id, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                        
                        if (ent is Curve curve)
                        {
                            double area = 0;
                            bool isClosed = false;

                            if (ent is Polyline pline)
                            {
                                isClosed = pline.Closed;
                                if (isClosed)
                                {
                                    area = pline.Area;
                                }
                                else
                                {
                                    // Nếu polyline không đóng, tính diện tích bằng cách giả lập đóng
                                    area = CalculateOpenPolylineArea(pline);
                                }
                            }
                            else if (ent is Polyline2d pline2d)
                            {
                                isClosed = pline2d.Closed;
                                area = pline2d.Area;
                            }
                            else if (ent is Polyline3d pline3d)
                            {
                                isClosed = pline3d.Closed;
                                area = pline3d.Area;
                            }

                            area = Math.Round(area, AreaDecimalPlaces);
                            totalArea += area;
                            count++;

                            string layer = ent.Layer;
                            areaInfos.Add(new PolylineAreaInfo
                            {
                                Index = count,
                                Layer = layer,
                                IsClosed = isClosed,
                                Area = area
                            });

                            A.Ed.WriteMessage($"\n  {count}. Layer [{layer}] - {(isClosed ? "Đóng" : "Mở")} - Diện tích: {area:F2} m²");
                        }
                    }

                    tr.Commit();
                }

                totalArea = Math.Round(totalArea, AreaDecimalPlaces);
                A.Ed.WriteMessage($"\n\n📊 TỔNG KẾT:");
                A.Ed.WriteMessage($"\n   Số polyline: {count}");
                A.Ed.WriteMessage($"\n   Tổng diện tích: {totalArea:F2} m²");
                A.Ed.WriteMessage("\n=== HOÀN TẤT ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Lệnh tính khối lượng từ các polyline vẽ thủ công (theo từng mặt cắt)
        /// </summary>
        [CommandMethod("CTSV_TinhKhoiLuongPoly")]
        public static void CTSVTinhKhoiLuongPoly()
        {
            try
            {
                A.Ed.WriteMessage("\n=== TÍNH KHỐI LƯỢNG TỪ POLYLINE VẼ THỦ CÔNG ===");
                A.Ed.WriteMessage("\n📌 Tính toán khối lượng bằng phương pháp Average End Area");
                A.Ed.WriteMessage("\n📌 Chọn các polyline theo thứ tự mặt cắt từ đầu đến cuối");

                // Nhập khoảng cách giữa các mặt cắt
                PromptDoubleOptions pdoSpacing = new PromptDoubleOptions("\nNhập khoảng cách giữa các mặt cắt (m):")
                {
                    AllowNegative = false,
                    AllowZero = false,
                    DefaultValue = 20.0
                };
                PromptDoubleResult pdrSpacing = A.Ed.GetDouble(pdoSpacing);
                if (pdrSpacing.Status != PromptStatus.OK)
                {
                    return;
                }
                double spacing = pdrSpacing.Value;

                // Chọn polylines theo thứ tự
                A.Ed.WriteMessage("\n📌 Chọn các polyline theo thứ tự mặt cắt:");

                TypedValue[] filterList = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<OR"),
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                    new TypedValue((int)DxfCode.Start, "POLYLINE"),
                    new TypedValue((int)DxfCode.Operator, "OR>")
                };
                SelectionFilter filter = new SelectionFilter(filterList);

                PromptSelectionResult psr = A.Ed.GetSelection(filter);
                if (psr.Status != PromptStatus.OK)
                {
                    A.Ed.WriteMessage("\n❌ Không có polyline nào được chọn.");
                    return;
                }

                SelectionSet ss = psr.Value;
                A.Ed.WriteMessage($"\n✅ Đã chọn {ss.Count} polyline");

                if (ss.Count < 2)
                {
                    A.Ed.WriteMessage("\n❌ Cần ít nhất 2 polyline để tính khối lượng!");
                    return;
                }

                List<double> areas = new List<double>();

                using (Transaction tr = A.Db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in ss.GetObjectIds())
                    {
                        Autodesk.AutoCAD.DatabaseServices.Entity ent = tr.GetObject(id, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                        
                        if (ent is Polyline pline)
                        {
                            double area = pline.Closed ? pline.Area : CalculateOpenPolylineArea(pline);
                            areas.Add(Math.Round(area, AreaDecimalPlaces));
                        }
                        else if (ent is Curve curve)
                        {
                            areas.Add(Math.Round(curve.Area, AreaDecimalPlaces));
                        }
                    }

                    tr.Commit();
                }

                // Tính khối lượng
                double totalVolume = 0;
                A.Ed.WriteMessage($"\n\n📊 CHI TIẾT TÍNH KHỐI LƯỢNG:");
                A.Ed.WriteMessage($"\n   Khoảng cách giữa các mặt cắt: {spacing} m");
                A.Ed.WriteMessage($"\n   Công thức: V = (S1 + S2) / 2 × L\n");
                A.Ed.WriteMessage($"\n   {"MC",-5} {"S (m²)",-12} {"V (m³)",-12} {"Cộng dồn (m³)",-15}");
                A.Ed.WriteMessage($"\n   {new string('-', 45)}");

                for (int i = 0; i < areas.Count; i++)
                {
                    double segmentVolume = 0;
                    if (i > 0)
                    {
                        segmentVolume = CalculateVolume(areas[i - 1], areas[i], spacing, true);
                        totalVolume += segmentVolume;
                    }

                    A.Ed.WriteMessage($"\n   {i + 1,-5} {areas[i],-12:F2} {segmentVolume,-12:F2} {totalVolume,-15:F2}");
                }

                totalVolume = Math.Round(totalVolume, VolumeDecimalPlaces);
                A.Ed.WriteMessage($"\n\n   📦 TỔNG KHỐI LƯỢNG: {totalVolume:F2} m³");
                A.Ed.WriteMessage("\n=== HOÀN TẤT ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Tính diện tích cho polyline mở (không đóng) bằng cách giả lập đóng
        /// </summary>
        private static double CalculateOpenPolylineArea(Polyline pline)
        {
            if (pline.NumberOfVertices < 3) return 0;

            // Sử dụng Shoelace formula
            List<Point2d> points = new List<Point2d>();
            for (int i = 0; i < pline.NumberOfVertices; i++)
            {
                Point2d pt = pline.GetPoint2dAt(i);
                points.Add(pt);
            }

            return CalculatePolygonArea(points);
        }

        /// <summary>
        /// Lệnh tính và ghi diện tích lên từng polyline
        /// </summary>
        [CommandMethod("CTSV_GhiDienTichPoly")]
        public static void CTSVGhiDienTichPoly()
        {
            try
            {
                A.Ed.WriteMessage("\n=== GHI DIỆN TÍCH LÊN POLYLINE ===");

                // Lấy chiều cao text
                PromptDoubleOptions pdoHeight = new PromptDoubleOptions("\nNhập chiều cao text:")
                {
                    AllowNegative = false,
                    AllowZero = false,
                    DefaultValue = 2.5
                };
                PromptDoubleResult pdrHeight = A.Ed.GetDouble(pdoHeight);
                if (pdrHeight.Status != PromptStatus.OK) return;
                double textHeight = pdrHeight.Value;

                // Chọn polylines
                TypedValue[] filterList = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<OR"),
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                    new TypedValue((int)DxfCode.Start, "POLYLINE"),
                    new TypedValue((int)DxfCode.Operator, "OR>")
                };
                SelectionFilter filter = new SelectionFilter(filterList);

                PromptSelectionResult psr = A.Ed.GetSelection(filter);
                if (psr.Status != PromptStatus.OK)
                {
                    A.Ed.WriteMessage("\n❌ Không có polyline nào được chọn.");
                    return;
                }

                SelectionSet ss = psr.Value;
                int count = 0;

                using (Transaction tr = A.Db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(A.Db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    foreach (ObjectId id in ss.GetObjectIds())
                    {
                        Autodesk.AutoCAD.DatabaseServices.Entity ent = tr.GetObject(id, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                        
                        if (ent is Polyline pline)
                        {
                            double area = pline.Closed ? pline.Area : CalculateOpenPolylineArea(pline);
                            area = Math.Round(area, AreaDecimalPlaces);

                            // Tính tâm polyline
                            Point3d centroid = GetPolylineCentroid(pline);

                            // Tạo text hiển thị diện tích
                            DBText areaText = new DBText
                            {
                                Position = centroid,
                                TextString = $"S={area:F2} m²",
                                Height = textHeight,
                                Layer = ent.Layer
                            };

                            btr.AppendEntity(areaText);
                            tr.AddNewlyCreatedDBObject(areaText, true);
                            count++;
                        }
                    }

                    tr.Commit();
                }

                A.Ed.WriteMessage($"\n✅ Đã ghi {count} diện tích lên các polyline.");
                A.Ed.WriteMessage("\n=== HOÀN TẤT ===\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Tính tâm (centroid) của polyline
        /// </summary>
        private static Point3d GetPolylineCentroid(Polyline pline)
        {
            double sumX = 0, sumY = 0;
            int count = pline.NumberOfVertices;

            for (int i = 0; i < count; i++)
            {
                Point2d pt = pline.GetPoint2dAt(i);
                sumX += pt.X;
                sumY += pt.Y;
            }

            return new Point3d(sumX / count, sumY / count, 0);
        }

        /// <summary>
        /// Class lưu thông tin diện tích polyline
        /// </summary>
        public class PolylineAreaInfo
        {
            public int Index { get; set; }
            public string Layer { get; set; } = "";
            public bool IsClosed { get; set; }
            public double Area { get; set; }
        }

        #endregion

        #endregion
    }
}
