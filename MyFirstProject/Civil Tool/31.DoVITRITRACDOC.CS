// 31.DoVITRITRACDOC.cs - Đo vị trí trắc dọc
// Cho chọn tim tuyến, chọn vị trí trên trắc dọc (profile view),
// tạo polyline vuông góc với tuyến tại lý trình đã chọn

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using MyFirstProject.Extensions;

// Alias để tránh xung đột
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using ContentAlignment = System.Drawing.ContentAlignment;
using FontStyle = System.Drawing.FontStyle;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.DoViTriTracDoc))]

namespace MyFirstProject
{
    #region Form

    /// <summary>
    /// Form cấu hình tạo polyline vuông góc tại lý trình trên trắc dọc
    /// </summary>
    public class DoViTriTracDocForm : Form
    {
        // ─── Theme Colors ───
        private static readonly Color BgDark = Color.FromArgb(30, 30, 34);
        private static readonly Color BgPanel = Color.FromArgb(40, 40, 45);
        private static readonly Color BgInput = Color.FromArgb(50, 50, 58);
        private static readonly Color AccentBlue = Color.FromArgb(0, 122, 204);
        private static readonly Color HeaderBlue = Color.FromArgb(0, 93, 160);
        private static readonly Color TextWhite = Color.White;
        private static readonly Color TextDim = Color.FromArgb(160, 160, 170);
        private static readonly Color BorderColor = Color.FromArgb(65, 65, 75);
        private static readonly Color GreenOk = Color.FromArgb(46, 160, 67);

        // ─── Controls ───
        private ComboBox cboAlignment = null!;
        private WinLabel lblAlignmentInfo = null!;
        private DataGridView dgvStations = null!; // Bảng danh sách lý trình đã chọn
        private NumericUpDown numLeftWidth = null!;
        private NumericUpDown numRightWidth = null!;
        private WinLabel lblTotal = null!;
        private WinLabel lblStationInfo = null!;

        // ─── Data ───
        private List<AlignmentInfo> alignments = new();
        public bool IsConfirmed { get; private set; } = false;
        public ObjectId SelectedAlignmentId { get; private set; } = ObjectId.Null;
        public string SelectedAlignmentName { get; private set; } = "";
        public double LeftWidth => (double)numLeftWidth.Value;
        public double RightWidth => (double)numRightWidth.Value;
        public List<StationItem> Stations { get; private set; } = new();

        // ─── Pick Mode ───
        public bool IsPickMode { get; private set; } = false;

        public class StationItem
        {
            public int STT { get; set; }
            public double Station { get; set; }
            public double Elevation { get; set; }
            public string Source { get; set; } = ""; // "ProfileView" hoặc "Manual"
        }

        private class AlignmentInfo
        {
            public ObjectId Id { get; set; }
            public string Name { get; set; } = "";
            public double Length { get; set; }
            public double StartStation { get; set; }
            public double EndStation { get; set; }
        }

        // ─── Constructor ───
        public DoViTriTracDocForm()
        {
            LoadAlignments();
            InitializeUI();
        }

        private void LoadAlignments()
        {
            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                CivilDocument cdoc = CivilApplication.ActiveDocument;
                ObjectIdCollection alIds = cdoc.GetAlignmentIds();
                foreach (ObjectId id in alIds)
                {
                    Alignment? al = tr.GetObject(id, OpenMode.ForRead) as Alignment;
                    if (al == null) continue;
                    alignments.Add(new AlignmentInfo
                    {
                        Id = id,
                        Name = al.Name,
                        Length = al.Length,
                        StartStation = al.StartingStation,
                        EndStation = al.EndingStation
                    });
                }
                tr.Commit();
            }
            catch { }
        }

        private void InitializeUI()
        {
            // ── Form ──
            this.Text = "TẠO POLYLINE VUÔNG GÓC TẠI LÝ TRÌNH — Civil Tool";
            this.Size = new Size(750, 680);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = BgDark;
            this.ForeColor = TextWhite;
            this.Font = new Font("Segoe UI", 9.5f);

            // ── HEADER ──
            var header = new Panel { Dock = DockStyle.Top, Height = 52, BackColor = HeaderBlue };
            var lblTitle = new WinLabel
            {
                Text = "📐  TẠO POLYLINE VUÔNG GÓC TẠI LÝ TRÌNH",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = TextWhite,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                Padding = new Padding(16, 0, 0, 0)
            };
            header.Controls.Add(lblTitle);
            this.Controls.Add(header);

            var subtitle = new WinLabel
            {
                Text = "Chọn tim tuyến → Chọn vị trí trên trắc dọc → Tạo polyline vuông góc với tuyến tại lý trình đó.",
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = TextDim,
                AutoSize = false,
                Height = 26,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Top,
                Padding = new Padding(18, 0, 0, 0),
                BackColor = Color.FromArgb(35, 35, 40)
            };
            this.Controls.Add(subtitle);
            subtitle.BringToFront();

            // ── MAIN CONTENT ──
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15),
                AutoScroll = true
            };
            this.Controls.Add(mainPanel);
            mainPanel.BringToFront();

            int y = 10;

            // ═══════════════════════════════════════════════════
            // SECTION 1: CHỌN ALIGNMENT
            // ═══════════════════════════════════════════════════
            var grpAlign = CreateGroupBox("🛣  Chọn Tim Tuyến (Alignment)", 0, y, 700, 55);
            mainPanel.Controls.Add(grpAlign);

            grpAlign.Controls.Add(CreateLabel("Tuyến:", 12, 28));
            cboAlignment = new ComboBox
            {
                Location = new Point(75, 24),
                Size = new Size(350, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextWhite,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f)
            };
            foreach (var al in alignments)
                cboAlignment.Items.Add(al.Name);
            if (cboAlignment.Items.Count > 0) cboAlignment.SelectedIndex = 0;
            cboAlignment.SelectedIndexChanged += (s, e) => OnAlignmentChanged();
            grpAlign.Controls.Add(cboAlignment);

            lblAlignmentInfo = CreateLabel("", 440, 28);
            lblAlignmentInfo.ForeColor = TextDim;
            lblAlignmentInfo.AutoSize = true;
            grpAlign.Controls.Add(lblAlignmentInfo);

            y += 60;

            // ═══════════════════════════════════════════════════
            // SECTION 2: CHỌN LÝ TRÌNH TRÊN TRẮC DỌC
            // ═══════════════════════════════════════════════════
            var grpStations = CreateGroupBox("📍  Danh sách lý trình (chọn trên trắc dọc)", 0, y, 700, 250);
            mainPanel.Controls.Add(grpStations);

            dgvStations = new DataGridView
            {
                Location = new Point(10, 28),
                Size = new Size(530, 175),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = BgInput,
                GridColor = BorderColor,
                BorderStyle = BorderStyle.FixedSingle,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = BgInput,
                    ForeColor = TextWhite,
                    SelectionBackColor = AccentBlue,
                    SelectionForeColor = TextWhite,
                    Font = new Font("Segoe UI", 9f)
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(45, 48, 55),
                    ForeColor = TextWhite,
                    Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                EnableHeadersVisualStyles = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            dgvStations.Columns.Add("STT", "STT");
            dgvStations.Columns.Add("StationVal", "Lý trình (Station)");
            dgvStations.Columns.Add("ElevationVal", "Cao độ (Elev)");
            dgvStations.Columns.Add("SourceVal", "Nguồn");
            dgvStations.Columns["STT"]!.Width = 45;
            dgvStations.Columns["SourceVal"]!.Width = 100;
            grpStations.Controls.Add(dgvStations);

            // Nút Chọn trên trắc dọc (Pick trên Profile View)
            var btnPickOnProfile = new Button
            {
                Text = "📍 Chọn trên\ntrắc dọc",
                Location = new Point(555, 28),
                Size = new Size(130, 45),
                FlatStyle = FlatStyle.Flat,
                BackColor = AccentBlue,
                ForeColor = TextWhite,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnPickOnProfile.FlatAppearance.BorderColor = Color.FromArgb(0, 80, 160);
            btnPickOnProfile.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 150, 230);
            btnPickOnProfile.Click += BtnPickOnProfile_Click;
            grpStations.Controls.Add(btnPickOnProfile);

            // Nút Xóa lý trình
            var btnRemoveStation = new Button
            {
                Text = "🗑 Xóa",
                Location = new Point(555, 80),
                Size = new Size(130, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(55, 55, 62),
                ForeColor = TextWhite,
                Font = new Font("Segoe UI", 9f),
                Cursor = Cursors.Hand
            };
            btnRemoveStation.FlatAppearance.BorderColor = BorderColor;
            btnRemoveStation.FlatAppearance.MouseOverBackColor = Color.FromArgb(180, 50, 50);
            btnRemoveStation.Click += BtnRemoveStation_Click;
            grpStations.Controls.Add(btnRemoveStation);

            // Nút Xóa tất cả
            var btnClearAll = new Button
            {
                Text = "🗑 Xóa tất cả",
                Location = new Point(555, 118),
                Size = new Size(130, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(55, 55, 62),
                ForeColor = TextWhite,
                Font = new Font("Segoe UI", 9f),
                Cursor = Cursors.Hand
            };
            btnClearAll.FlatAppearance.BorderColor = BorderColor;
            btnClearAll.FlatAppearance.MouseOverBackColor = Color.FromArgb(180, 50, 50);
            btnClearAll.Click += (s, e) =>
            {
                Stations.Clear();
                dgvStations.Rows.Clear();
                UpdateStationGrid();
            };
            grpStations.Controls.Add(btnClearAll);

            // Nút Nhập tay
            var btnManualStation = new Button
            {
                Text = "✏️ Nhập tay",
                Location = new Point(555, 160),
                Size = new Size(130, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(55, 55, 62),
                ForeColor = TextWhite,
                Font = new Font("Segoe UI", 9f),
                Cursor = Cursors.Hand
            };
            btnManualStation.FlatAppearance.BorderColor = BorderColor;
            btnManualStation.FlatAppearance.MouseOverBackColor = AccentBlue;
            btnManualStation.Click += BtnManualStation_Click;
            grpStations.Controls.Add(btnManualStation);

            lblStationInfo = CreateLabel("", 10, 210);
            lblStationInfo.ForeColor = TextDim;
            lblStationInfo.AutoSize = true;
            grpStations.Controls.Add(lblStationInfo);

            lblTotal = CreateLabel("Tổng: 0 lý trình", 10, 228);
            lblTotal.ForeColor = TextDim;
            lblTotal.AutoSize = true;
            grpStations.Controls.Add(lblTotal);

            y += 258;

            // ═══════════════════════════════════════════════════
            // SECTION 3: KÍCH THƯỚC POLYLINE
            // ═══════════════════════════════════════════════════
            var grpSize = CreateGroupBox("📏  Kích thước Polyline vuông góc", 0, y, 700, 75);
            mainPanel.Controls.Add(grpSize);

            grpSize.Controls.Add(CreateLabel("Chiều rộng trái (m):", 12, 35));
            numLeftWidth = new NumericUpDown
            {
                Location = new Point(170, 32),
                Size = new Size(100, 28),
                Minimum = 1,
                Maximum = 500,
                Value = 20,
                DecimalPlaces = 1,
                Increment = 5,
                BackColor = BgInput,
                ForeColor = TextWhite
            };
            grpSize.Controls.Add(numLeftWidth);

            grpSize.Controls.Add(CreateLabel("Chiều rộng phải (m):", 320, 35));
            numRightWidth = new NumericUpDown
            {
                Location = new Point(480, 32),
                Size = new Size(100, 28),
                Minimum = 1,
                Maximum = 500,
                Value = 20,
                DecimalPlaces = 1,
                Increment = 5,
                BackColor = BgInput,
                ForeColor = TextWhite
            };
            grpSize.Controls.Add(numRightWidth);

            grpSize.Controls.Add(CreateDimLabel("(Từ tim tuyến sang 2 bên)", 590, 35));

            y += 85;

            // ═══════════════════════════════════════════════════
            // BOTTOM BUTTONS
            // ═══════════════════════════════════════════════════
            var pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 55,
                BackColor = Color.FromArgb(38, 38, 42),
                Padding = new Padding(12, 8, 12, 8)
            };
            this.Controls.Add(pnlBottom);

            var btnCancel = CreateButton("✖ Hủy", 490, 8, 100, 36, pnlBottom);
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            var btnExecute = CreateButton("✅ Tạo Polyline", 605, 8, 130, 36, pnlBottom);
            btnExecute.BackColor = GreenOk;
            btnExecute.FlatAppearance.BorderColor = Color.FromArgb(35, 134, 54);
            btnExecute.Click += BtnExecute_Click;

            // Initialize
            if (alignments.Count > 0)
                OnAlignmentChanged();
        }

        // ═══════════════════════════════════════════════════
        // EVENT HANDLERS
        // ═══════════════════════════════════════════════════

        private void OnAlignmentChanged()
        {
            if (cboAlignment.SelectedIndex < 0 || cboAlignment.SelectedIndex >= alignments.Count) return;
            var al = alignments[cboAlignment.SelectedIndex];
            SelectedAlignmentId = al.Id;
            SelectedAlignmentName = al.Name;
            lblAlignmentInfo.Text = $"L={al.Length:F1}m  Sta: {al.StartStation:F0} → {al.EndStation:F0}";

            // Xóa danh sách cũ khi đổi tuyến
            Stations.Clear();
            UpdateStationGrid();
        }

        /// <summary>
        /// Click nút "📍 Chọn trên trắc dọc" → đóng form tạm, pick trên profile view
        /// </summary>
        private void BtnPickOnProfile_Click(object? sender, EventArgs e)
        {
            if (SelectedAlignmentId == ObjectId.Null)
            {
                MessageBox.Show("Vui lòng chọn Alignment trước!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Đóng form tạm với DialogResult.Retry để command xử lý pick
            IsPickMode = true;
            this.DialogResult = DialogResult.Retry;
            this.Close();
        }

        /// <summary>
        /// Nhập tay station
        /// </summary>
        private void BtnManualStation_Click(object? sender, EventArgs e)
        {
            if (SelectedAlignmentId == ObjectId.Null)
            {
                MessageBox.Show("Vui lòng chọn Alignment trước!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string input = Microsoft.VisualBasic.Interaction.InputBox(
                "Nhập lý trình (Station):", "Nhập lý trình", "0.00");
            if (string.IsNullOrEmpty(input)) return;

            if (!double.TryParse(input, out double station))
            {
                MessageBox.Show("Giá trị lý trình không hợp lệ!", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            AddStation(station, 0, "Manual");
        }

        /// <summary>
        /// Thêm station vào danh sách
        /// </summary>
        public void AddStation(double station, double elevation, string source)
        {
            // Kiểm tra trùng lặp
            if (Stations.Any(s => Math.Abs(s.Station - station) < 0.01))
            {
                MessageBox.Show($"Lý trình {station:F2} đã tồn tại trong danh sách!",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Stations.Add(new StationItem
            {
                STT = Stations.Count + 1,
                Station = station,
                Elevation = elevation,
                Source = source
            });

            UpdateStationGrid();
        }

        /// <summary>
        /// Reset pick mode
        /// </summary>
        public void CancelPickMode()
        {
            IsPickMode = false;
        }

        private void BtnRemoveStation_Click(object? sender, EventArgs e)
        {
            if (dgvStations.SelectedRows.Count > 0)
            {
                var selectedIndices = new List<int>();
                foreach (DataGridViewRow row in dgvStations.SelectedRows)
                {
                    selectedIndices.Add(row.Index);
                }
                selectedIndices.Sort();
                selectedIndices.Reverse();

                foreach (int idx in selectedIndices)
                {
                    if (idx >= 0 && idx < Stations.Count)
                        Stations.RemoveAt(idx);
                }

                UpdateStationGrid();
            }
            else
            {
                MessageBox.Show("Vui lòng chọn lý trình cần xóa!",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void UpdateStationGrid()
        {
            dgvStations.Rows.Clear();

            // Sắp xếp theo station
            var sorted = Stations.OrderBy(s => s.Station).ToList();
            for (int i = 0; i < sorted.Count; i++)
            {
                sorted[i].STT = i + 1;
                dgvStations.Rows.Add(sorted[i].STT, $"{sorted[i].Station:F2}",
                    $"{sorted[i].Elevation:F3}", sorted[i].Source);
            }
            Stations = sorted;

            lblTotal.Text = $"Tổng: {Stations.Count} lý trình";
            lblTotal.ForeColor = Stations.Count > 0 ? GreenOk : TextDim;
        }

        private void BtnExecute_Click(object? sender, EventArgs e)
        {
            if (Stations.Count == 0)
            {
                MessageBox.Show("Chưa có lý trình nào! Hãy chọn trên trắc dọc hoặc nhập tay.",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var confirm = MessageBox.Show(
                $"Tạo {Stations.Count} polyline vuông góc với tuyến '{SelectedAlignmentName}'?\n\n" +
                $"Chiều rộng: {numLeftWidth.Value}m (trái) + {numRightWidth.Value}m (phải)",
                "Xác nhận",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm == DialogResult.Yes)
            {
                IsConfirmed = true;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        // ═══════════════════════════════════════════════════
        // UI HELPERS
        // ═══════════════════════════════════════════════════

        private Panel CreateGroupBox(string title, int x, int y, int w, int h)
        {
            var panel = new Panel
            {
                Location = new Point(x, y),
                Size = new Size(w, h),
                BackColor = BgPanel,
                BorderStyle = BorderStyle.FixedSingle
            };
            var lbl = new WinLabel
            {
                Text = title,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                ForeColor = AccentBlue,
                Location = new Point(10, 5),
                AutoSize = true
            };
            panel.Controls.Add(lbl);
            return panel;
        }

        private WinLabel CreateLabel(string text, int x, int y)
        {
            return new WinLabel
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                ForeColor = TextWhite,
                Font = new Font("Segoe UI", 9.5f)
            };
        }

        private WinLabel CreateDimLabel(string text, int x, int y)
        {
            return new WinLabel
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                ForeColor = TextDim,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Italic)
            };
        }

        private Button CreateButton(string text, int x, int y, int w, int h, Control parent)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(55, 55, 62),
                ForeColor = TextWhite,
                Font = new Font("Segoe UI", 9f),
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter
            };
            btn.FlatAppearance.BorderColor = BorderColor;
            btn.FlatAppearance.MouseOverBackColor = AccentBlue;
            parent.Controls.Add(btn);
            return btn;
        }
    }

    #endregion

    #region Command

    /// <summary>
    /// Lệnh tạo polyline vuông góc với tuyến tại lý trình chọn trên trắc dọc
    /// Flow: Chọn tuyến → Chọn vị trí trên Profile View → Tạo polyline vuông góc
    /// </summary>
    public class DoViTriTracDoc
    {
        [CommandMethod("CTC_TaoPolylineVuongGocTracDoc")]
        public static void CTC_TaoPolylineVuongGocTracDoc()
        {
            var ed = A.Ed;
            var db = A.Db;

            // Mở form - vòng lặp để xử lý pick trên trắc dọc
            DoViTriTracDocForm form = new();
            while (true)
            {
                var dialogResult = AcadApp.ShowModalDialog(form);

                // Nếu user click nút 📍 chọn trên trắc dọc → form trả về Retry
                if (dialogResult == DialogResult.Retry && form.IsPickMode)
                {
                    // Cho phép pick nhiều điểm liên tục cho đến khi nhấn ESC
                    ed.WriteMessage("\n📍 Chọn các vị trí trên trắc dọc (Profile View). Nhấn ESC khi xong.");

                    while (true)
                    {
                        var ppo = new PromptPointOptions(
                            "\n📍 Chọn vị trí trên trắc dọc (hoặc ESC để quay lại form): ");
                        ppo.AllowNone = true;
                        var ppr = ed.GetPoint(ppo);

                        if (ppr.Status != PromptStatus.OK)
                            break; // User nhấn ESC → quay lại form

                        Point3d pickedPoint = ppr.Value;

                        // Tìm Profile View chứa điểm đã pick
                        try
                        {
                            double station = 0;
                            double elevation = 0;
                            bool found = false;

                            using Transaction tr = db.TransactionManager.StartTransaction();

                            // Tìm tất cả ProfileView trong bản vẽ
                            Alignment? alignment = tr.GetObject(
                                form.SelectedAlignmentId, OpenMode.ForRead) as Alignment;

                            if (alignment != null)
                            {
                                ObjectIdCollection pvIds = alignment.GetProfileViewIds();

                                foreach (ObjectId pvId in pvIds)
                                {
                                    ProfileView? pv = tr.GetObject(pvId, OpenMode.ForRead) as ProfileView;
                                    if (pv == null) continue;

                                    try
                                    {
                                        // Chuyển tọa độ point trên bản vẽ → station/elevation trên profile view
                                        double sta = 0, elev = 0;
                                        pv.FindStationAndElevationAtXY(
                                            pickedPoint.X, pickedPoint.Y, ref sta, ref elev);

                                        // Kiểm tra station nằm trong phạm vi alignment
                                        if (sta >= alignment.StartingStation &&
                                            sta <= alignment.EndingStation)
                                        {
                                            station = sta;
                                            elevation = elev;
                                            found = true;
                                            break;
                                        }
                                    }
                                    catch
                                    {
                                        // Điểm không nằm trong profile view này → thử cái tiếp theo
                                        continue;
                                    }
                                }
                            }

                            tr.Commit();

                            if (found)
                            {
                                form.AddStation(station, elevation, "ProfileView");
                                ed.WriteMessage(
                                    $"\n✅ Station = {station:F2}, Elevation = {elevation:F3}");
                            }
                            else
                            {
                                // Thử tính station bằng cách chiếu điểm lên alignment
                                // (trường hợp user click gần tuyến nhưng không trên profile view)
                                try
                                {
                                    using Transaction tr2 = db.TransactionManager.StartTransaction();
                                    Alignment? al2 = tr2.GetObject(
                                        form.SelectedAlignmentId, OpenMode.ForRead) as Alignment;
                                    if (al2 != null)
                                    {
                                        double sta2 = 0, off2 = 0;
                                        al2.StationOffset(pickedPoint.X, pickedPoint.Y,
                                            ref sta2, ref off2);

                                        if (sta2 >= al2.StartingStation && sta2 <= al2.EndingStation)
                                        {
                                            form.AddStation(sta2, 0, "Alignment");
                                            ed.WriteMessage(
                                                $"\n✅ Station = {sta2:F2} (từ alignment, offset = {off2:F2}m)");
                                        }
                                        else
                                        {
                                            ed.WriteMessage(
                                                "\n⚠️ Không tìm thấy Profile View tại vị trí này. Thử chọn lại.");
                                        }
                                    }
                                    tr2.Commit();
                                }
                                catch
                                {
                                    ed.WriteMessage(
                                        "\n⚠️ Không thể xác định lý trình. Vui lòng chọn trên trắc dọc.");
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
                        }
                    }

                    // Reset pick mode và quay lại form
                    form.CancelPickMode();
                    continue;
                }

                // Nếu user nhấn OK (Tạo Polyline)
                if (dialogResult == DialogResult.OK && form.IsConfirmed)
                    break;

                // Nếu user hủy
                return;
            }

            if (form.Stations.Count == 0) return;

            // ══════════════════════════════════════════════════════
            // TẠO POLYLINE VUÔNG GÓC VỚI TUYẾN TẠI CÁC LÝ TRÌNH
            // ══════════════════════════════════════════════════════
            string layerName = "CTC_VUONGGOC_TRACDOC";
            CreateLayerIfNotExist(layerName, 1); // Màu đỏ

            int created = 0;
            double leftW = form.LeftWidth;
            double rightW = form.RightWidth;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    Alignment? alignment = tr.GetObject(
                        form.SelectedAlignmentId, OpenMode.ForRead) as Alignment;
                    if (alignment == null) { tr.Commit(); return; }

                    BlockTable? bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    if (bt == null) { tr.Commit(); return; }
                    BlockTableRecord? btr = tr.GetObject(
                        bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    if (btr == null) { tr.Commit(); return; }

                    foreach (var stItem in form.Stations)
                    {
                        try
                        {
                            double sta = stItem.Station;

                            // Lấy tọa độ tâm trên tuyến
                            double cx = 0, cy = 0;
                            alignment.PointLocation(sta, 0, ref cx, ref cy);

                            // Tính góc tiếp tuyến tại vị trí này
                            double delta = 0.5;
                            double e1 = 0, n1 = 0, e2 = 0, n2 = 0;
                            double sta1 = Math.Max(alignment.StartingStation, sta - delta);
                            double sta2 = Math.Min(alignment.EndingStation, sta + delta);
                            alignment.PointLocation(sta1, 0, ref e1, ref n1);
                            alignment.PointLocation(sta2, 0, ref e2, ref n2);

                            double tangentAngle = Math.Atan2(n2 - n1, e2 - e1);

                            // Góc vuông góc = góc tiếp tuyến + 90°
                            double perpAngle = tangentAngle + Math.PI / 2.0;
                            double cosP = Math.Cos(perpAngle);
                            double sinP = Math.Sin(perpAngle);

                            // Điểm đầu (bên trái) và điểm cuối (bên phải)
                            Point2d ptLeft = new Point2d(
                                cx + leftW * cosP,
                                cy + leftW * sinP);
                            Point2d ptRight = new Point2d(
                                cx - rightW * cosP,
                                cy - rightW * sinP);

                            // Tạo polyline
                            var pline = new Polyline();
                            pline.AddVertexAt(0, ptLeft, 0, 0, 0);
                            pline.AddVertexAt(1, new Point2d(cx, cy), 0, 0, 0);
                            pline.AddVertexAt(2, ptRight, 0, 0, 0);
                            pline.Layer = layerName;
                            pline.ColorIndex = 1; // Đỏ

                            btr.AppendEntity(pline);
                            tr.AddNewlyCreatedDBObject(pline, true);

                            // Thêm text ghi lý trình
                            double textH = Math.Max(leftW, rightW) * 0.06;
                            var text = new DBText
                            {
                                TextString = $"Sta {sta:F2}",
                                Position = new Point3d(cx, cy, 0),
                                Height = textH,
                                Rotation = tangentAngle,
                                HorizontalMode = TextHorizontalMode.TextCenter,
                                VerticalMode = TextVerticalMode.TextBottom,
                                AlignmentPoint = new Point3d(cx, cy, 0),
                                Layer = layerName,
                                ColorIndex = 1
                            };
                            btr.AppendEntity(text);
                            tr.AddNewlyCreatedDBObject(text, true);

                            created++;
                            ed.WriteMessage(
                                $"\n   ✅ Polyline vuông góc tại Sta {sta:F2} " +
                                $"(góc {tangentAngle * 180.0 / Math.PI:F1}°)");
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage(
                                $"\n   ❌ Lỗi tại Sta {stItem.Station:F2}: {ex.Message}");
                        }
                    }

                    tr.Commit();
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
                }
            }

            ed.WriteMessage(
                $"\n✅ Đã tạo {created}/{form.Stations.Count} polyline vuông góc trên layer '{layerName}'");
        }

        private static void CreateLayerIfNotExist(string name, short color)
        {
            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                LayerTable? lt = tr.GetObject(A.Db.LayerTableId, OpenMode.ForWrite) as LayerTable;
                if (lt != null && !lt.Has(name))
                {
                    LayerTableRecord ltr = new()
                    {
                        Name = name,
                        Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                            Autodesk.AutoCAD.Colors.ColorMethod.ByAci, color)
                    };
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }
                tr.Commit();
            }
            catch { }
        }
    }

    #endregion
}
