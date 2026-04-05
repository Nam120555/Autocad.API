// (C) Copyright 2024
// Thống kê cọc từ SampleLine và xuất ra Excel
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
using FormLabel = System.Windows.Forms.Label;

using ClosedXML.Excel;
using MyFirstProject.Extensions;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(MyFirstProject.ThongKeCoc))]

namespace MyFirstProject
{
    #region Data Classes

    /// <summary>
    /// Thông tin chi tiết của một cọc (SampleLine)
    /// </summary>
    public class CocInfo
    {
        public int STT { get; set; }                      // Số thứ tự
        public string TenCoc { get; set; } = "";          // Tên cọc
        public string LyTrinh { get; set; } = "";         // Lý trình format: Km0+123.456
        public double LyTrinhValue { get; set; }          // Giá trị lý trình số
        public double Easting { get; set; }               // Tọa độ X
        public double Northing { get; set; }              // Tọa độ Y
        public double KhoangCachDenCocTruoc { get; set; } // Khoảng cách đến cọc trước
        public double BeRongTrai { get; set; }            // Bề rộng bên trái
        public double BeRongPhai { get; set; }            // Bề rộng bên phải
        public double CaoDoTuNhien { get; set; }          // Cao độ tự nhiên (EG)
        public double CaoDoThietKe { get; set; }          // Cao độ thiết kế (FG)
    }

    /// <summary>
    /// Thông tin cấu hình một tuyến để xuất
    /// </summary>
    public class AlignmentExportConfig
    {
        public bool IsSelected { get; set; }
        public ObjectId AlignmentId { get; set; }
        public string AlignmentName { get; set; } = "";
        public ObjectId SampleLineGroupId { get; set; }
        public string SampleLineGroupName { get; set; } = "";
        public ObjectId ProfileTNId { get; set; }         // Profile Tự nhiên
        public string ProfileTNName { get; set; } = "";
        public ObjectId ProfileTKId { get; set; }         // Profile Thiết kế
        public string ProfileTKName { get; set; } = "";
        
        // Các collections cho dropdown
        public List<(ObjectId Id, string Name)> AvailableSLGs { get; set; } = new();
        public List<(ObjectId Id, string Name)> AvailableProfiles { get; set; } = new();
    }

    #endregion

    #region Selection Form

    /// <summary>
    /// Form chọn tuyến để thống kê cọc (tương tự V3Tools)
    /// </summary>
    public class AlignmentSelectionForm : CivilFormBase
    {
        private DataGridView dgvAlignments = null!;
        private NumericUpDown numDecimal = null!;
        private CheckBox chkCaoDoTuNhien = null!;
        private CheckBox chkCaoDoThietKe = null!;
        private Button btnSelectAll = null!;
        private Button btnExit = null!;
        private Button btnExecute = null!;
        private NumericUpDown numTableSpacing = null!;
        private CheckBox chkExportToAcad = null!;

        public List<AlignmentExportConfig> Configs { get; private set; } = new();
        public int DecimalPlaces { get; private set; } = 2;
        public bool IncludeCaoDoTuNhien { get; private set; } = true;
        public bool IncludeCaoDoThietKe { get; private set; } = true;
        public bool ExportToAcad { get; private set; } = false;
        public double TableSpacing { get; private set; } = 50.0;
        public bool IsConfirmed { get; private set; } = false;

        public AlignmentSelectionForm(List<AlignmentExportConfig> configs)
        {
            Configs = configs;
            InitializeComponent();
            LoadData();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("Tạo bảng thống kê tọa độ cọc", 700, 620);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            var header = CreateHeader("THỐNG KÊ TỌA ĐỘ CỌC");
            this.Controls.Add(header);

            // GroupBox danh sách tuyến
            var pnlList = CreateSection("Danh sách tuyến", 10, 60, 665, 340);
            this.Controls.Add(pnlList);

            // DataGridView
            dgvAlignments = new DataGridView
            {
                Location = new System.Drawing.Point(10, 30),
                Size = new System.Drawing.Size(645, 300),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = BgInput,
                GridColor = BorderColor,
                BorderStyle = BorderStyle.FixedSingle,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = BgInput,
                    ForeColor = TextDark,
                    SelectionBackColor = NavyMedium,
                    SelectionForeColor = System.Drawing.Color.White,
                    Font = new System.Drawing.Font("Segoe UI", 9f)
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = NavyDark,
                    ForeColor = System.Drawing.Color.White,
                    Font = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                EnableHeadersVisualStyles = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            pnlList.Controls.Add(dgvAlignments);

            // Columns
            DataGridViewCheckBoxColumn colCheck = new DataGridViewCheckBoxColumn
            {
                HeaderText = "",
                Width = 30,
                Name = "colCheck"
            };
            dgvAlignments.Columns.Add(colCheck);

            DataGridViewTextBoxColumn colName = new DataGridViewTextBoxColumn
            {
                HeaderText = "Tên tuyến",
                Width = 120,
                Name = "colName",
                ReadOnly = true
            };
            dgvAlignments.Columns.Add(colName);

            DataGridViewComboBoxColumn colSLG = new DataGridViewComboBoxColumn
            {
                HeaderText = "Nhóm cọc",
                Width = 120,
                Name = "colSLG"
            };
            dgvAlignments.Columns.Add(colSLG);

            DataGridViewComboBoxColumn colProfileTN = new DataGridViewComboBoxColumn
            {
                HeaderText = "Profile TN",
                Width = 120,
                Name = "colProfileTN"
            };
            dgvAlignments.Columns.Add(colProfileTN);

            DataGridViewComboBoxColumn colProfileTK = new DataGridViewComboBoxColumn
            {
                HeaderText = "Profile TK",
                Width = 120,
                Name = "colProfileTK"
            };
            dgvAlignments.Columns.Add(colProfileTK);

            // GroupBox Tùy chọn
            var grpOptions = CreateSection("Tùy chọn", 10, 410, 665, 120);
            this.Controls.Add(grpOptions);

            // Độ chính xác
            FormLabel lblDecimal = new FormLabel
            {
                Text = "Độ chính xác:",
                Location = new System.Drawing.Point(15, 35),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            grpOptions.Controls.Add(lblDecimal);

            numDecimal = new NumericUpDown
            {
                Location = new System.Drawing.Point(110, 33),
                Width = 60,
                Minimum = 0,
                Maximum = 6,
                Value = 2,
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9f)
            };
            grpOptions.Controls.Add(numDecimal);

            // Checkboxes
            chkCaoDoTuNhien = new CheckBox
            {
                Text = "Cao độ tự nhiên",
                Location = new System.Drawing.Point(200, 33),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            grpOptions.Controls.Add(chkCaoDoTuNhien);

            chkCaoDoThietKe = new CheckBox
            {
                Text = "Cao độ thiết kế",
                Location = new System.Drawing.Point(340, 33),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            grpOptions.Controls.Add(chkCaoDoThietKe);
            
            // Button chọn tuyến
            btnSelectAll = new Button
            {
                Text = "☑ Chọn tất cả",
                Location = new System.Drawing.Point(500, 30),
                Size = new System.Drawing.Size(145, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9f)
            };
            btnSelectAll.FlatAppearance.BorderColor = BorderColor;
            btnSelectAll.Click += BtnSelectAll_Click;
            grpOptions.Controls.Add(btnSelectAll);

            // AutoCAD Export Options
            var lblSpacing = new FormLabel
            {
                Text = "Khoảng cách bảng:",
                Location = new System.Drawing.Point(15, 75),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f)
            };
            grpOptions.Controls.Add(lblSpacing);

            numTableSpacing = new NumericUpDown
            {
                Location = new System.Drawing.Point(145, 73),
                Width = 80,
                Minimum = 0,
                Maximum = 1000,
                Value = 50,
                DecimalPlaces = 1,
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new System.Drawing.Font("Segoe UI", 9f),
                Enabled = false
            };
            grpOptions.Controls.Add(numTableSpacing);

            chkExportToAcad = new CheckBox
            {
                Text = "Xuất ra bản vẽ AutoCAD",
                Location = new System.Drawing.Point(250, 73),
                AutoSize = true,
                Checked = false,
                ForeColor = NavyDark,
                Font = new System.Drawing.Font("Segoe UI", 9.5f, System.Drawing.FontStyle.Bold)
            };
            chkExportToAcad.CheckedChanged += (s, e) => numTableSpacing.Enabled = chkExportToAcad.Checked;
            grpOptions.Controls.Add(chkExportToAcad);

            // Buttons bottom
            var pnlBottom = CreateBottomBar();
            this.Controls.Add(pnlBottom);

            btnExit = CreateButton("Hủy", 435, 8, 110, 36, pnlBottom);
            btnExit.Click += (s, e) => { IsConfirmed = false; this.Close(); };

            btnExecute = CreatePrimaryButton("Thực hiện", 555, 8, 120, 36, pnlBottom);
            btnExecute.Click += BtnExecute_Click;
        }

        private void LoadData()
        {
            dgvAlignments.Rows.Clear();

            foreach (var config in Configs)
            {
                int rowIndex = dgvAlignments.Rows.Add();
                var row = dgvAlignments.Rows[rowIndex];
                row.Tag = config;

                row.Cells["colCheck"].Value = config.IsSelected;
                row.Cells["colName"].Value = config.AlignmentName;

                // Populate SLG dropdown
                var slgCell = (DataGridViewComboBoxCell)row.Cells["colSLG"];
                slgCell.Items.Clear();
                foreach (var slg in config.AvailableSLGs)
                {
                    slgCell.Items.Add(slg.Name);
                }
                if (config.AvailableSLGs.Count > 0)
                    slgCell.Value = config.AvailableSLGs[0].Name;

                // Populate Profile TN dropdown
                var tnCell = (DataGridViewComboBoxCell)row.Cells["colProfileTN"];
                tnCell.Items.Clear();
                tnCell.Items.Add("(Không)");
                foreach (var p in config.AvailableProfiles)
                {
                    tnCell.Items.Add(p.Name);
                }
                // Auto-select EG profile
                var egProfile = config.AvailableProfiles.FirstOrDefault(p => 
                    p.Name.Contains("EG") || p.Name.Contains("TN") || p.Name.Contains("Natural"));
                tnCell.Value = egProfile.Name ?? "(Không)";

                // Populate Profile TK dropdown
                var tkCell = (DataGridViewComboBoxCell)row.Cells["colProfileTK"];
                tkCell.Items.Clear();
                tkCell.Items.Add("(Không)");
                foreach (var p in config.AvailableProfiles)
                {
                    tkCell.Items.Add(p.Name);
                }
                // Auto-select FG/TK profile
                var fgProfile = config.AvailableProfiles.FirstOrDefault(p => 
                    p.Name.Contains("FG") || p.Name.Contains("TK") || p.Name.Contains("Design"));
                tkCell.Value = fgProfile.Name ?? "(Không)";
            }
        }

        private void BtnSelectAll_Click(object? sender, EventArgs e)
        {
            bool allChecked = dgvAlignments.Rows.Cast<DataGridViewRow>()
                .All(r => (bool)(r.Cells["colCheck"].Value ?? false));

            foreach (DataGridViewRow row in dgvAlignments.Rows)
            {
                row.Cells["colCheck"].Value = !allChecked;
            }

            btnSelectAll.Text = allChecked ? "☑ Chọn tất cả" : "☐ Bỏ chọn tất cả";
        }

        private void BtnExecute_Click(object? sender, EventArgs e)
        {
            // Update configs from grid
            foreach (DataGridViewRow row in dgvAlignments.Rows)
            {
                var config = (AlignmentExportConfig)row.Tag!;
                config.IsSelected = (bool)(row.Cells["colCheck"].Value ?? false);

                string slgName = row.Cells["colSLG"].Value?.ToString() ?? "";
                var slg = config.AvailableSLGs.FirstOrDefault(s => s.Name == slgName);
                config.SampleLineGroupId = slg.Id;
                config.SampleLineGroupName = slg.Name;

                string tnName = row.Cells["colProfileTN"].Value?.ToString() ?? "";
                if (tnName != "(Không)")
                {
                    var tn = config.AvailableProfiles.FirstOrDefault(p => p.Name == tnName);
                    config.ProfileTNId = tn.Id;
                    config.ProfileTNName = tn.Name;
                }

                string tkName = row.Cells["colProfileTK"].Value?.ToString() ?? "";
                if (tkName != "(Không)")
                {
                    var tk = config.AvailableProfiles.FirstOrDefault(p => p.Name == tkName);
                    config.ProfileTKId = tk.Id;
                    config.ProfileTKName = tk.Name;
                }
            }

            DecimalPlaces = (int)numDecimal.Value;
            IncludeCaoDoTuNhien = chkCaoDoTuNhien.Checked;
            IncludeCaoDoThietKe = chkCaoDoThietKe.Checked;
            ExportToAcad = chkExportToAcad.Checked;
            TableSpacing = (double)numTableSpacing.Value;

            // Validate
            int selectedCount = Configs.Count(c => c.IsSelected);
            if (selectedCount == 0)
            {
                MessageBox.Show("Vui lòng chọn ít nhất 1 tuyến!", "Thông báo", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IsConfirmed = true;
            this.Close();
        }
    }

    #endregion

    public class ThongKeCoc
    {
        #region Helper Methods

        /// <summary>
        /// Format lý trình theo dạng Km0+123.456
        /// </summary>
        public static string FormatStation(double station)
        {
            int km = (int)(station / 1000);
            double meters = station % 1000;
            return $"Km{km}+{meters:F3}";
        }

        /// <summary>
        /// Lấy danh sách Alignments có SampleLineGroup
        /// </summary>
        private static List<(ObjectId AlignmentId, string Name, ObjectId SampleLineGroupId)> GetAlignmentsWithSampleLineGroups(Transaction tr)
        {
            var result = new List<(ObjectId, string, ObjectId)>();

            ObjectIdCollection alignmentIds = A.Cdoc.GetAlignmentIds();
            foreach (ObjectId alignmentId in alignmentIds)
            {
                Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
                if (alignment == null) continue;

                ObjectIdCollection slgIds = alignment.GetSampleLineGroupIds();
                if (slgIds.Count > 0)
                {
                    result.Add((alignmentId, alignment.Name, slgIds[0]));
                }
            }

            return result;
        }

        /// <summary>
        /// Trích xuất thông tin cọc từ SampleLineGroup
        /// </summary>
        private static List<CocInfo> ExtractCocInfo(Transaction tr, ObjectId sampleLineGroupId, ObjectId alignmentId)
        {
            List<CocInfo> cocInfos = new();

            SampleLineGroup? slg = tr.GetObject(sampleLineGroupId, AcadDb.OpenMode.ForRead) as SampleLineGroup;
            if (slg == null) return cocInfos;

            Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
            if (alignment == null) return cocInfos;

            // Lấy danh sách SampleLine và sắp xếp theo lý trình
            List<SampleLine> sortedSampleLines = new List<SampleLine>();
            foreach (ObjectId slId in slg.GetSampleLineIds())
            {
                if (tr.GetObject(slId, OpenMode.ForRead) is SampleLine sl)
                    sortedSampleLines.Add(sl);
            }
            sortedSampleLines = sortedSampleLines.OrderBy(s => s.Station).ToList();

            double prevStation = 0;
            int stt = 1;

            foreach (SampleLine sampleLine in sortedSampleLines)
            {
                // Lấy tọa độ tâm cọc
                double easting = 0, northing = 0;
                alignment.PointLocation(sampleLine.Station, 0, ref easting, ref northing);

                // Lấy bề rộng trái/phải từ Vertices
                double beRongTrai = 0, beRongPhai = 0;
                foreach (SampleLineVertex vertex in sampleLine.Vertices)
                {
                    double dist = Math.Sqrt(Math.Pow(vertex.Location.X - easting, 2) + Math.Pow(vertex.Location.Y - northing, 2));
                    if (vertex.Side == SampleLineVertexSideType.Left)
                        beRongTrai = dist;
                    else if (vertex.Side == SampleLineVertexSideType.Right)
                        beRongPhai = dist;
                }

                CocInfo cocInfo = new()
                {
                    STT = stt,
                    TenCoc = sampleLine.Name,
                    LyTrinh = FormatStation(sampleLine.Station),
                    LyTrinhValue = sampleLine.Station,
                    Easting = Math.Round(easting, 3),
                    Northing = Math.Round(northing, 3),
                    KhoangCachDenCocTruoc = stt == 1 ? 0 : Math.Round(sampleLine.Station - prevStation, 3),
                    BeRongTrai = Math.Round(beRongTrai, 3),
                    BeRongPhai = Math.Round(beRongPhai, 3)
                };

                cocInfos.Add(cocInfo);
                prevStation = sampleLine.Station;
                stt++;
            }

            return cocInfos;
        }

        /// <summary>
        /// Xuất danh sách cọc ra Excel
        /// </summary>
        private static void ExportToExcel(string filePath, string alignmentName, List<CocInfo> cocInfos)
        {
            using var workbook = new XLWorkbook();
            string sheetName = alignmentName.Length > 31 ? alignmentName.Substring(0, 31) : alignmentName;
            // Thay thế ký tự không hợp lệ
            char[] invalidChars = { ':', '\\', '/', '?', '*', '[', ']' };
            foreach (char c in invalidChars)
            {
                sheetName = sheetName.Replace(c, '_');
            }

            var ws = workbook.Worksheets.Add(sheetName);

            // Tiêu đề
            ws.Cell(1, 1).Value = $"BẢNG THỐNG KÊ CỌC - {alignmentName}";
            ws.Range(1, 1, 1, 9).Merge();
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontSize = 14;
            ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Header
            ws.Cell(2, 1).Value = "STT";
            ws.Cell(2, 2).Value = "TÊN CỌC";
            ws.Cell(2, 3).Value = "LÝ TRÌNH";
            ws.Cell(2, 4).Value = "EASTING (X)";
            ws.Cell(2, 5).Value = "NORTHING (Y)";
            ws.Cell(2, 6).Value = "KHOẢNG CÁCH (m)";
            ws.Cell(2, 7).Value = "BỀ RỘNG TRÁI (m)";
            ws.Cell(2, 8).Value = "BỀ RỘNG PHẢI (m)";
            ws.Cell(2, 9).Value = "TỔNG BỀ RỘNG (m)";

            // Format header
            var headerRange = ws.Range(2, 1, 2, 9);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.WrapText = true;

            // Dữ liệu
            int row = 3;
            foreach (var coc in cocInfos)
            {
                ws.Cell(row, 1).Value = coc.STT;
                ws.Cell(row, 2).Value = coc.TenCoc;
                ws.Cell(row, 3).Value = coc.LyTrinh;
                ws.Cell(row, 4).Value = coc.Easting;
                ws.Cell(row, 5).Value = coc.Northing;
                ws.Cell(row, 6).Value = coc.KhoangCachDenCocTruoc;
                ws.Cell(row, 7).Value = coc.BeRongTrai;
                ws.Cell(row, 8).Value = coc.BeRongPhai;
                ws.Cell(row, 9).Value = coc.BeRongTrai + coc.BeRongPhai;
                row++;
            }

            // Hàng tổng cộng
            ws.Cell(row, 1).Value = "TỔNG CỘNG";
            ws.Range(row, 1, row, 5).Merge();
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 6).FormulaA1 = $"SUM(F3:F{row - 1})";
            ws.Cell(row, 6).Style.Font.Bold = true;

            // Format bảng
            var tableRange = ws.Range(2, 1, row, 9);
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Điều chỉnh độ rộng cột
            ws.Column(1).Width = 8;   // STT
            ws.Column(2).Width = 20;  // Tên cọc
            ws.Column(3).Width = 15;  // Lý trình
            ws.Column(4).Width = 15;  // Easting
            ws.Column(5).Width = 15;  // Northing
            ws.Column(6).Width = 15;  // Khoảng cách
            ws.Column(7).Width = 15;  // Bề rộng trái
            ws.Column(8).Width = 15;  // Bề rộng phải
            ws.Column(9).Width = 15;  // Tổng bề rộng

            // Căn giữa các cột số
            for (int r = 3; r <= row; r++)
            {
                ws.Cell(r, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                for (int c = 4; c <= 9; c++)
                {
                    ws.Cell(r, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                }
            }

            workbook.SaveAs(filePath);
        }

        #endregion

        #region Commands

        /// <summary>
        /// Lấy danh sách configs cho tất cả Alignments có SampleLineGroup
        /// </summary>
        private static List<AlignmentExportConfig> GetAlignmentConfigs(Transaction tr)
        {
            var configs = new List<AlignmentExportConfig>();

            ObjectIdCollection alignmentIds = A.Cdoc.GetAlignmentIds();
            foreach (ObjectId alignmentId in alignmentIds)
            {
                Alignment? alignment = tr.GetObject(alignmentId, AcadDb.OpenMode.ForRead) as Alignment;
                if (alignment == null) continue;

                ObjectIdCollection slgIds = alignment.GetSampleLineGroupIds();
                if (slgIds.Count == 0) continue;

                var config = new AlignmentExportConfig
                {
                    IsSelected = true,
                    AlignmentId = alignmentId,
                    AlignmentName = alignment.Name
                };

                // Lấy danh sách SampleLineGroups
                foreach (ObjectId slgId in slgIds)
                {
                    if (tr.GetObject(slgId, OpenMode.ForRead) is SampleLineGroup slg)
                    {
                        config.AvailableSLGs.Add((slgId, slg.Name));
                    }
                }

                // Lấy danh sách Profiles
                ObjectIdCollection profileIds = alignment.GetProfileIds();
                foreach (ObjectId profileId in profileIds)
                {
                    if (tr.GetObject(profileId, OpenMode.ForRead) is Profile profile)
                    {
                        config.AvailableProfiles.Add((profileId, profile.Name));
                    }
                }

                if (config.AvailableSLGs.Count > 0)
                {
                    config.SampleLineGroupId = config.AvailableSLGs[0].Id;
                    config.SampleLineGroupName = config.AvailableSLGs[0].Name;
                }

                configs.Add(config);
            }

            return configs;
        }

        /// <summary>
        /// Trích xuất thông tin cọc với cao độ từ profiles
        /// </summary>
        private static List<CocInfo> ExtractCocInfoWithElevation(Transaction tr, AlignmentExportConfig config, int decimalPlaces)
        {
            List<CocInfo> cocInfos = new();

            SampleLineGroup? slg = tr.GetObject(config.SampleLineGroupId, AcadDb.OpenMode.ForRead) as SampleLineGroup;
            if (slg == null) return cocInfos;

            Alignment? alignment = tr.GetObject(config.AlignmentId, AcadDb.OpenMode.ForRead) as Alignment;
            if (alignment == null) return cocInfos;

            // Lấy profiles
            Profile? profileTN = null;
            Profile? profileTK = null;
            
            if (!config.ProfileTNId.IsNull)
                profileTN = tr.GetObject(config.ProfileTNId, OpenMode.ForRead) as Profile;
            if (!config.ProfileTKId.IsNull)
                profileTK = tr.GetObject(config.ProfileTKId, OpenMode.ForRead) as Profile;

            // Lấy danh sách SampleLine và sắp xếp theo lý trình
            List<SampleLine> sortedSampleLines = new List<SampleLine>();
            foreach (ObjectId slId in slg.GetSampleLineIds())
            {
                if (tr.GetObject(slId, OpenMode.ForRead) is SampleLine sl)
                    sortedSampleLines.Add(sl);
            }
            sortedSampleLines = sortedSampleLines.OrderBy(s => s.Station).ToList();

            double prevStation = 0;
            int stt = 1;

            foreach (SampleLine sampleLine in sortedSampleLines)
            {
                // Lấy tọa độ tâm cọc
                double easting = 0, northing = 0;
                alignment.PointLocation(sampleLine.Station, 0, ref easting, ref northing);

                // Lấy bề rộng trái/phải từ Vertices
                double beRongTrai = 0, beRongPhai = 0;
                foreach (SampleLineVertex vertex in sampleLine.Vertices)
                {
                    double dist = Math.Sqrt(Math.Pow(vertex.Location.X - easting, 2) + Math.Pow(vertex.Location.Y - northing, 2));
                    if (vertex.Side == SampleLineVertexSideType.Left)
                        beRongTrai = dist;
                    else if (vertex.Side == SampleLineVertexSideType.Right)
                        beRongPhai = dist;
                }

                // Lấy cao độ từ profiles
                double caoDoTN = 0, caoDoTK = 0;
                try
                {
                    if (profileTN != null)
                        caoDoTN = profileTN.ElevationAt(sampleLine.Station);
                }
                catch { }

                try
                {
                    if (profileTK != null)
                        caoDoTK = profileTK.ElevationAt(sampleLine.Station);
                }
                catch { }

                CocInfo cocInfo = new()
                {
                    STT = stt,
                    TenCoc = sampleLine.Name,
                    LyTrinh = FormatStation(sampleLine.Station),
                    LyTrinhValue = sampleLine.Station,
                    Easting = Math.Round(easting, decimalPlaces),
                    Northing = Math.Round(northing, decimalPlaces),
                    KhoangCachDenCocTruoc = stt == 1 ? 0 : Math.Round(sampleLine.Station - prevStation, decimalPlaces),
                    BeRongTrai = Math.Round(beRongTrai, decimalPlaces),
                    BeRongPhai = Math.Round(beRongPhai, decimalPlaces),
                    CaoDoTuNhien = Math.Round(caoDoTN, decimalPlaces),
                    CaoDoThietKe = Math.Round(caoDoTK, decimalPlaces)
                };

                cocInfos.Add(cocInfo);
                prevStation = sampleLine.Station;
                stt++;
            }

            return cocInfos;
        }

        /// <summary>
        /// Xuất Excel với tùy chọn cao độ
        /// </summary>
        private static void ExportToExcelWithOptions(string filePath, List<(string Name, List<CocInfo> Data)> alignmentData, 
            bool includeTN, bool includeTK, int decimalPlaces)
        {
            using var workbook = new XLWorkbook();

            foreach (var (name, cocInfos) in alignmentData)
            {
                if (cocInfos.Count == 0) continue;

                string sheetName = name.Length > 31 ? name.Substring(0, 31) : name;
                char[] invalidChars = { ':', '\\', '/', '?', '*', '[', ']' };
                foreach (char c in invalidChars)
                {
                    sheetName = sheetName.Replace(c, '_');
                }

                var ws = workbook.Worksheets.Add(sheetName);

                // Xác định số cột
                int colCount = 6; // STT, Tên cọc, Lý trình, X, Y, Khoảng cách
                if (includeTN) colCount++;
                if (includeTK) colCount++;
                colCount += 3; // Bề rộng trái, phải, tổng

                // Tiêu đề
                ws.Cell(1, 1).Value = $"BẢNG THỐNG KÊ TỌA ĐỘ CỌC - {name}";
                ws.Range(1, 1, 1, colCount).Merge();
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Font.FontSize = 14;
                ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Header
                int col = 1;
                ws.Cell(2, col++).Value = "STT";
                ws.Cell(2, col++).Value = "TÊN CỌC";
                ws.Cell(2, col++).Value = "LÝ TRÌNH";
                ws.Cell(2, col++).Value = "EASTING (X)";
                ws.Cell(2, col++).Value = "NORTHING (Y)";
                if (includeTN) ws.Cell(2, col++).Value = "CAO ĐỘ TN (m)";
                if (includeTK) ws.Cell(2, col++).Value = "CAO ĐỘ TK (m)";
                ws.Cell(2, col++).Value = "KHOẢNG CÁCH (m)";
                ws.Cell(2, col++).Value = "BỀ RỘNG TRÁI (m)";
                ws.Cell(2, col++).Value = "BỀ RỘNG PHẢI (m)";
                ws.Cell(2, col++).Value = "TỔNG BỀ RỘNG (m)";

                var headerRange = ws.Range(2, 1, 2, colCount);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Alignment.WrapText = true;

                // Dữ liệu
                int row = 3;
                foreach (var coc in cocInfos)
                {
                    col = 1;
                    ws.Cell(row, col++).Value = coc.STT;
                    ws.Cell(row, col++).Value = coc.TenCoc;
                    ws.Cell(row, col++).Value = coc.LyTrinh;
                    ws.Cell(row, col++).Value = coc.Easting;
                    ws.Cell(row, col++).Value = coc.Northing;
                    if (includeTN) ws.Cell(row, col++).Value = coc.CaoDoTuNhien;
                    if (includeTK) ws.Cell(row, col++).Value = coc.CaoDoThietKe;
                    ws.Cell(row, col++).Value = coc.KhoangCachDenCocTruoc;
                    ws.Cell(row, col++).Value = coc.BeRongTrai;
                    ws.Cell(row, col++).Value = coc.BeRongPhai;
                    ws.Cell(row, col++).Value = coc.BeRongTrai + coc.BeRongPhai;
                    row++;
                }

                // Hàng tổng
                ws.Cell(row, 1).Value = "TỔNG CỘNG";
                int mergeEnd = 5;
                if (includeTN) mergeEnd++;
                if (includeTK) mergeEnd++;
                ws.Range(row, 1, row, mergeEnd).Merge();
                ws.Cell(row, 1).Style.Font.Bold = true;
                
                int kcCol = mergeEnd + 1;
                ws.Cell(row, kcCol).FormulaA1 = $"SUM({GetColumnLetter(kcCol)}3:{GetColumnLetter(kcCol)}{row - 1})";
                ws.Cell(row, kcCol).Style.Font.Bold = true;

                // Format
                var tableRange = ws.Range(2, 1, row, colCount);
                tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                ws.Columns().AdjustToContents();
            }

            workbook.SaveAs(filePath);
        }

        private static string GetColumnLetter(int columnNumber)
        {
            string result = "";
            while (columnNumber > 0)
            {
                columnNumber--;
                result = (char)('A' + columnNumber % 26) + result;
                columnNumber /= 26;
            }
            return result;
        }

        /// <summary>
        /// Vẽ bảng AutoCAD từ danh sách cọc
        /// </summary>
        private static double DrawTableInAutoCAD(Transaction tr, string alignmentName, List<CocInfo> cocInfos, 
            Point3d startPoint, bool includeTN, bool includeTK)
        {
            Autodesk.AutoCAD.DatabaseServices.Table table = new Autodesk.AutoCAD.DatabaseServices.Table();
            table.TableStyle = A.Db.Tablestyle; // Use current AutoCAD table style
            table.SetSize(cocInfos.Count + 3, 6 + (includeTN ? 1 : 0) + (includeTK ? 1 : 0) + 3);
            table.Position = startPoint;

            // Header
            table.Cells[0, 0].TextString = $"BẢNG THỐNG KÊ TỌA ĐỘ CỌC - {alignmentName.ToUpper()}";
            table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;
            table.Cells[0, 0].TextHeight = 2.5;

            int col = 0;
            table.Cells[1, col++].TextString = "STT";
            table.Cells[1, col++].TextString = "TÊN CỌC";
            table.Cells[1, col++].TextString = "LÝ TRÌNH";
            table.Cells[1, col++].TextString = "X (EAST)";
            table.Cells[1, col++].TextString = "Y (NORTH)";
            if (includeTN) table.Cells[1, col++].TextString = "CAO ĐỘ TN";
            if (includeTK) table.Cells[1, col++].TextString = "CAO ĐỘ TK";
            table.Cells[1, col++].TextString = "K.CÁCH (m)";
            table.Cells[1, col++].TextString = "B.RỘNG TRÁI";
            table.Cells[1, col++].TextString = "B.RỘNG PHẢI";
            table.Cells[1, col++].TextString = "TỔNG RỘNG";

            for (int i = 0; i < col; i++)
            {
                table.Cells[1, i].Alignment = CellAlignment.MiddleCenter;
                table.Cells[1, i].TextHeight = 2.0;
            }

            // Data
            for (int i = 0; i < cocInfos.Count; i++)
            {
                int row = i + 2;
                var coc = cocInfos[i];
                col = 0;
                table.Cells[row, col++].TextString = coc.STT.ToString();
                table.Cells[row, col++].TextString = coc.TenCoc;
                table.Cells[row, col++].TextString = coc.LyTrinh;
                table.Cells[row, col++].TextString = coc.Easting.ToString("F3");
                table.Cells[row, col++].TextString = coc.Northing.ToString("F3");
                if (includeTN) table.Cells[row, col++].TextString = coc.CaoDoTuNhien.ToString("F3");
                if (includeTK) table.Cells[row, col++].TextString = coc.CaoDoThietKe.ToString("F3");
                table.Cells[row, col++].TextString = coc.KhoangCachDenCocTruoc.ToString("F3");
                table.Cells[row, col++].TextString = coc.BeRongTrai.ToString("F3");
                table.Cells[row, col++].TextString = coc.BeRongPhai.ToString("F3");
                table.Cells[row, col++].TextString = (coc.BeRongTrai + coc.BeRongPhai).ToString("F3");

                for (int j = 0; j < col; j++)
                {
                    table.Cells[row, j].Alignment = CellAlignment.MiddleCenter;
                    table.Cells[row, j].TextHeight = 1.8;
                }
            }

            // Footer
            int footerRow = cocInfos.Count + 2;
            table.Cells[footerRow, 0].TextString = "TỔNG CỘNG";
            int mergeCols = 5 + (includeTN ? 1 : 0) + (includeTK ? 1 : 0);
            table.MergeCells(CellRange.Create(table, footerRow, 0, footerRow, mergeCols - 1));
            
            double totalDist = cocInfos.Sum(c => c.KhoangCachDenCocTruoc);
            table.Cells[footerRow, mergeCols].TextString = totalDist.ToString("F3");
            table.Cells[footerRow, mergeCols].Alignment = CellAlignment.MiddleCenter;

            table.GenerateLayout();
            
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(A.Db.CurrentSpaceId, OpenMode.ForWrite);
            btr.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);

            // Re-calculate height after layout generation to accurately space tables
            double height = table.Rows.Cast<Row>().Sum(r => r.Height);
            return height;
        }

        /// <summary>
        /// Lệnh thống kê cọc và xuất Excel hoặc AutoCAD Table
        /// </summary>
        [CommandMethod("CTSV_ThongKeCoc")]
        public static void CTSVThongKeCoc()
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                // 1. Lấy danh sách configs
                var configs = GetAlignmentConfigs(tr);

                if (configs.Count == 0)
                {
                    A.Ed.WriteMessage("\n❌ Không tìm thấy Alignment nào có SampleLineGroup!");
                    return;
                }

                // 2. Hiển thị form chọn
                List<AlignmentExportConfig> selectedConfigs;
                int decimalPlaces;
                bool includeCaoDoTuNhien;
                bool includeCaoDoThietKe;
                bool exportToAcad;
                double tableSpacing;

                using (AlignmentSelectionForm form = new AlignmentSelectionForm(configs))
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(form);

                    if (!form.IsConfirmed)
                    {
                        A.Ed.WriteMessage("\n⚠️ Đã hủy thao tác.");
                        return;
                    }

                    selectedConfigs = form.Configs.Where(c => c.IsSelected).ToList();
                    decimalPlaces = form.DecimalPlaces;
                    includeCaoDoTuNhien = form.IncludeCaoDoTuNhien;
                    includeCaoDoThietKe = form.IncludeCaoDoThietKe;
                    exportToAcad = form.ExportToAcad;
                    tableSpacing = form.TableSpacing;
                }

                if (selectedConfigs.Count == 0)
                {
                    A.Ed.WriteMessage("\n❌ Không có tuyến nào được chọn!");
                    return;
                }

                if (exportToAcad)
                {
                    // AutoCAD Table Branch
                    Point3d insertPoint = UserInput.GPoint("\nChọn điểm đặt cho bảng đầu tiên:");
                    A.Ed.WriteMessage("\n\n✍️ ĐANG VẼ BẢNG VÀO BẢN VẼ...");
                    
                    Point3d currentPoint = insertPoint;
                    foreach (var config in selectedConfigs)
                    {
                        var cocInfos = ExtractCocInfoWithElevation(tr, config, decimalPlaces);
                        if (cocInfos.Count > 0)
                        {
                            double tableHeight = DrawTableInAutoCAD(tr, config.AlignmentName, cocInfos, 
                                currentPoint, includeCaoDoTuNhien, includeCaoDoThietKe);
                            
                            // Calculate next point (move down vertically)
                            currentPoint = new Point3d(currentPoint.X, currentPoint.Y - tableHeight - tableSpacing, 0);
                            A.Ed.WriteMessage($"\n   ✓ Đã vẽ: {config.AlignmentName}");
                        }
                    }
                    A.Ed.WriteMessage("\n✅ Đã hoàn tất vẽ bảng.");
                }
                else
                {
                    // Excel Export Branch (EXISTING)
                    string defaultFileName = selectedConfigs.Count == 1 
                        ? $"ThongKeCoc_{selectedConfigs[0].AlignmentName}.xlsx" 
                        : "ThongKeCoc_TatCa.xlsx";
                    
                    SaveFileDialog saveDialog = new()
                    {
                        Filter = "Excel Files (*.xlsx)|*.xlsx",
                        Title = "Chọn nơi lưu file Excel thống kê cọc",
                        FileName = defaultFileName
                    };

                    if (saveDialog.ShowDialog() != DialogResult.OK) return;

                    A.Ed.WriteMessage("\n\n📊 ĐANG XỬ LÝ EXCEL...");
                    var alignmentData = new List<(string Name, List<CocInfo> Data)>();
                    int totalCoc = 0;

                    foreach (var config in selectedConfigs)
                    {
                        var cocInfos = ExtractCocInfoWithElevation(tr, config, decimalPlaces);
                        if (cocInfos.Count > 0)
                        {
                            alignmentData.Add((config.AlignmentName, cocInfos));
                            totalCoc += cocInfos.Count;
                            A.Ed.WriteMessage($"\n   ✓ {config.AlignmentName}: {cocInfos.Count} cọc");
                        }
                    }

                    ExportToExcelWithOptions(saveDialog.FileName, alignmentData, 
                        includeCaoDoTuNhien, includeCaoDoThietKe, decimalPlaces);

                    A.Ed.WriteMessage($"\n\n📊 Tổng: {selectedConfigs.Count} tuyến, {totalCoc} cọc");
                    A.Ed.WriteMessage($"\n✅ Đã xuất file: {saveDialog.FileName}");
                    
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = saveDialog.FileName,
                        UseShellExecute = true
                    });
                }

                A.Ed.WriteMessage("\n=== HOÀN TẤT ===\n");
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        #endregion
    }
}
