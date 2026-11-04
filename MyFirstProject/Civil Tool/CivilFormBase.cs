// CivilFormBase.cs — Base class chung cho tất cả Dialog Form trong dự án Civil Tool
// Theme: Autodesk Civil 3D Studio Dark (#222222 / #2C2C2C / Accent Civil Blue #0696D7)
// Thiết kế đồng bộ hoàn hảo với không gian làm việc hiện đại của AutoCAD & Civil 3D

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using ContentAlignment = System.Drawing.ContentAlignment;
using FontStyle = System.Drawing.FontStyle;

namespace MyFirstProject
{
    /// <summary>
    /// Base Form class cho toàn dự án Civil Tool
    /// Autodesk Civil 3D Dark Studio theme — đồng bộ màu sắc, typography và visual controls
    /// </summary>
    public class CivilFormBase : Form
    {
        // ═══ Bảng màu Autodesk Civil 3D Studio Dark ═══
        public static readonly Color CivilDark = Color.FromArgb(34, 34, 34);          // #222222 — Nền form chuẩn Civil 3D
        public static readonly Color CivilCard = Color.FromArgb(44, 44, 44);          // #2C2C2C — Nền panel / container
        public static readonly Color CivilInput = Color.FromArgb(24, 24, 24);         // #181818 — Ô nhập liệu lõm
        public static readonly Color CivilHeader = Color.FromArgb(26, 26, 26);        // #1A1A1A — Thanh tiêu đề
        
        public static readonly Color CivilBlue = Color.FromArgb(6, 150, 215);         // #0696D7 — Accent đặc trưng Civil 3D
        public static readonly Color CivilBlueHover = Color.FromArgb(26, 170, 238);    // #1AAAFE — Hover xanh sáng
        public static readonly Color CivilBlueDeep = Color.FromArgb(4, 110, 160);      // #046EA0 — Nhấn đậm

        // Biến tương thích ngược với code cũ (tự động mapping sang Dark Palette)
        protected static readonly Color NavyDark = CivilBlue;
        protected static readonly Color NavyMedium = CivilBlueHover;
        protected static readonly Color NavyDeep = CivilBlueDeep;

        protected static readonly Color BgWhite = CivilDark;                          // Chuyển nền trắng cũ sang nền Dark
        protected static readonly Color BgPanel = CivilCard;                          // Chuyển panel cũ sang Card Dark
        protected static readonly Color BgInput = CivilInput;                         // Ô nhập liệu dark

        protected static readonly Color TextDark = Color.FromArgb(235, 235, 235);     // Chữ chính: Trắng sáng
        protected static readonly Color TextNavy = CivilBlue;                         // Tiêu đề: Civil Blue
        protected static readonly Color TextHint = Color.FromArgb(155, 155, 155);     // Chữ phụ: Xám bạc
        protected static readonly Color TextRed = Color.FromArgb(245, 85, 85);        // Cảnh báo: Đỏ san hô mềm
        protected static readonly Color TextDim = Color.FromArgb(135, 135, 135);
        protected static readonly Color TextWhite = Color.White;

        protected static readonly Color BorderLight = Color.FromArgb(62, 62, 62);     // Viền tối tinh tế
        protected static readonly Color BorderColor = Color.FromArgb(62, 62, 62);
        protected static readonly Color GreenOk = Color.FromArgb(40, 167, 69);        // Xanh lá chuẩn
        protected static readonly Color HoverNavy = Color.FromArgb(52, 52, 52);       // Hover nền xám tối

        /// <summary>
        /// Khởi tạo form với theme Autodesk Civil 3D Dark Studio chuẩn
        /// </summary>
        protected void InitBaseTheme(string title, int width = 720, int height = 560)
        {
            this.Text = title;
            this.Size = new Size(width, height);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            this.BackColor = CivilDark;
            this.ForeColor = TextDark;
            this.Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);
            this.DoubleBuffered = true;
        }

        /// <summary>
        /// Tự động thêm chữ ký tác giả góc dưới và đồng bộ dark theme cho các control con
        /// </summary>
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            
            // Tự động đồng bộ dark theme cho toàn bộ control được add vào form
            ApplyCivilDarkTheme(this);

            // Chữ ký tác giả phong cách Autodesk Dark Footer
            MyFirstProject.FormSignatureHelper.AddSignature(this);
        }

        /// <summary>
        /// Tạo Header bar phong cách Civil 3D với dải nhấn Cyan trên cùng
        /// </summary>
        protected Panel CreateHeader(string title, int height = 48)
        {
            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = height,
                BackColor = CivilHeader,
                Padding = new Padding(0)
            };

            header.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                // 1. Dải nhấn xanh Civil 3D Cyan trên cùng (3px)
                using (var penCyan = new Pen(CivilBlue, 3f))
                {
                    g.DrawLine(penCyan, 0, 1, header.Width, 1);
                }

                // 2. Viền dưới mỏng phân cách nội dung
                using (var penBorder = new Pen(BorderLight, 1f))
                {
                    g.DrawLine(penBorder, 0, header.Height - 1, header.Width, header.Height - 1);
                }

                // 3. Biểu tượng hình thoi Civil 3D nhỏ góc trái
                using (var bCyan = new SolidBrush(CivilBlue))
                {
                    g.FillPolygon(bCyan, new PointF[]
                    {
                        new PointF(22, 24),
                        new PointF(29, 17),
                        new PointF(36, 24),
                        new PointF(29, 31)
                    });
                }
            };

            var lbl = new WinLabel
            {
                Text = title,
                Font = new Font("Segoe UI", 12f, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                Padding = new Padding(46, 0, 0, 0),
                BackColor = Color.Transparent
            };
            header.Controls.Add(lbl);
            return header;
        }

        /// <summary>
        /// Tạo subtitle bar
        /// </summary>
        protected WinLabel CreateSubtitle(string text)
        {
            return new WinLabel
            {
                Text = text,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
                ForeColor = TextHint,
                AutoSize = false,
                Height = 24,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Top,
                Padding = new Padding(18, 0, 0, 0),
                BackColor = Color.FromArgb(28, 28, 28)
            };
        }

        /// <summary>
        /// Tạo Section / Card (panel có khung và title phong cách Autodesk)
        /// </summary>
        protected Panel CreateSection(string title, int x, int y, int w, int h) => CreateGroupBox(title, x, y, w, h);

        /// <summary>
        /// Tạo GroupBox (panel có viền bo nhẹ và header Civil Blue)
        /// </summary>
        protected Panel CreateGroupBox(string title, int x, int y, int w, int h)
        {
            var panel = new Panel
            {
                Location = new Point(x, y),
                Size = new Size(w, h),
                BackColor = CivilCard,
                Padding = new Padding(10, 28, 10, 10)
            };

            // Vẽ viền Card mỏng sang trọng
            panel.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                using var pen = new Pen(BorderLight, 1f);
                var rect = new Rectangle(0, 0, panel.Width - 1, panel.Height - 1);
                g.DrawRectangle(pen, rect);

                // Đường kẻ mờ phân tách header của Group
                using var pDiv = new Pen(Color.FromArgb(52, 52, 52), 1f);
                g.DrawLine(pDiv, 8, 26, panel.Width - 9, 26);
            };

            var lbl = new WinLabel
            {
                Text = "●  " + title.ToUpper(),
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = CivilBlue,
                Location = new Point(10, 6),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            panel.Controls.Add(lbl);

            return panel;
        }

        /// <summary>
        /// Tạo Label nội dung
        /// </summary>
        protected WinLabel CreateLabel(string text, int x, int y, bool autoSize = true)
        {
            return new WinLabel
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = autoSize,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
                BackColor = Color.Transparent
            };
        }

        /// <summary>
        /// Tạo Label tiêu đề Civil Blue
        /// </summary>
        protected WinLabel CreateTitleLabel(string text, int x, int y)
        {
            return new WinLabel
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                ForeColor = CivilBlue,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                BackColor = Color.Transparent
            };
        }

        /// <summary>
        /// Tạo Label gợi ý (xám bạc)
        /// </summary>
        protected WinLabel CreateHintLabel(string text, int x, int y)
        {
            return new WinLabel
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                ForeColor = TextHint,
                Font = new Font("Segoe UI", 8.25f, FontStyle.Regular),
                BackColor = Color.Transparent
            };
        }

        /// <summary>
        /// Tạo Label ghi chú cảnh báo (đỏ san hô)
        /// </summary>
        protected WinLabel CreateRedNote(string text, int x, int y)
        {
            return new WinLabel
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                ForeColor = TextRed,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Italic),
                BackColor = Color.Transparent
            };
        }

        /// <summary>
        /// Tạo TextBox Dark Input hiện đại
        /// </summary>
        protected TextBox CreateTextBox(int x, int y, int w = 250, string defaultValue = "")
        {
            var txt = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(w, 26),
                BackColor = CivilInput,
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
                Text = defaultValue
            };
            return txt;
        }

        /// <summary>
        /// Tạo ComboBox Dark Flat
        /// </summary>
        protected ComboBox CreateComboBox(int x, int y, int w = 300)
        {
            return new ComboBox
            {
                Location = new Point(x, y),
                Size = new Size(w, 26),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(38, 38, 38),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular)
            };
        }

        /// <summary>
        /// Tạo NumericUpDown Dark
        /// </summary>
        protected NumericUpDown CreateNumeric(int x, int y, int w = 100,
            decimal min = 0, decimal max = 9999, decimal value = 1, int decimals = 0)
        {
            return new NumericUpDown
            {
                Location = new Point(x, y),
                Size = new Size(w, 26),
                Minimum = min,
                Maximum = max,
                Value = value,
                DecimalPlaces = decimals,
                BackColor = CivilInput,
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular)
            };
        }

        /// <summary>
        /// Tạo CheckBox Dark
        /// </summary>
        protected CheckBox CreateCheckBox(string text, int x, int y, bool isChecked = false)
        {
            return new CheckBox
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular),
                Checked = isChecked,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent
            };
        }

        /// <summary>
        /// Tạo Button chuẩn (Secondary)
        /// </summary>
        protected Button CreateButton(string text, int x, int y, int w = 120, int h = 32, Control? parent = null)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(52, 52, 52),
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular),
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter
            };
            btn.FlatAppearance.BorderColor = BorderLight;
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(68, 68, 68);
            btn.MouseEnter += (s, e) => btn.ForeColor = Color.White;
            btn.MouseLeave += (s, e) => btn.ForeColor = TextDark;

            parent?.Controls.Add(btn);
            return btn;
        }

        /// <summary>
        /// Tạo Button Civil Blue chính (OK, Thực hiện, Lưu)
        /// </summary>
        protected Button CreatePrimaryButton(string text, int x, int y, int w = 140, int h = 34, Control? parent = null)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                FlatStyle = FlatStyle.Flat,
                BackColor = CivilBlue,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = CivilBlueHover;
            btn.FlatAppearance.MouseDownBackColor = CivilBlueDeep;

            parent?.Controls.Add(btn);
            return btn;
        }

        /// <summary>
        /// Tạo Button thành công (Xanh lá)
        /// </summary>
        protected Button CreateSuccessButton(string text, int x, int y, int w = 140, int h = 34, Control? parent = null)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                FlatStyle = FlatStyle.Flat,
                BackColor = GreenOk,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(46, 190, 80);

            parent?.Controls.Add(btn);
            return btn;
        }

        /// <summary>
        /// Tạo Bottom bar chuẩn Dark
        /// </summary>
        protected Panel CreateBottomBar(int height = 50)
        {
            var pnl = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = height,
                BackColor = Color.FromArgb(28, 28, 28),
                Padding = new Padding(14, 8, 14, 8)
            };
            pnl.Paint += (s, e) =>
            {
                using var p = new Pen(BorderLight, 1);
                e.Graphics.DrawLine(p, 0, 0, pnl.Width, 0);
            };
            return pnl;
        }

        /// <summary>
        /// Tạo DataGridView Dark Studio chuẩn Civil 3D
        /// </summary>
        protected DataGridView CreateDataGrid(int x, int y, int w, int h)
        {
            return new DataGridView
            {
                Location = new Point(x, y),
                Size = new Size(w, h),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BackgroundColor = CivilInput,
                GridColor = Color.FromArgb(48, 48, 48),
                BorderStyle = BorderStyle.FixedSingle,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(36, 36, 36),
                    ForeColor = TextDark,
                    SelectionBackColor = Color.FromArgb(9, 71, 113), // #094771 — Highlight xanh Navy Civil
                    SelectionForeColor = Color.White,
                    Font = new Font("Segoe UI", 9f)
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(30, 30, 30),
                    ForeColor = TextDark,
                    SelectionBackColor = Color.FromArgb(9, 71, 113),
                    SelectionForeColor = Color.White
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(24, 24, 24),
                    ForeColor = CivilBlue,
                    Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                EnableHeadersVisualStyles = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
        }

        /// <summary>
        /// Tạo ListView chuẩn Dark
        /// </summary>
        protected ListView CreateListView(int x, int y, int w, int h)
        {
            return new ListView
            {
                Location = new Point(x, y),
                Size = new Size(w, h),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = CivilInput,
                ForeColor = TextDark,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9f)
            };
        }

        /// <summary>
        /// Tự động duyệt và áp dụng giao diện Dark Studio cho toàn bộ Control con trong Form
        /// </summary>
        public static void ApplyCivilDarkTheme(Control root)
        {
            foreach (Control c in root.Controls)
            {
                if (c is Panel pnl)
                {
                    // Nếu là header hoặc footer thì đã có style riêng
                    if (pnl.Dock != DockStyle.Top && pnl.Dock != DockStyle.Bottom)
                    {
                        if (pnl.BackColor == Color.White || pnl.BackColor == Color.FromArgb(245, 247, 250))
                            pnl.BackColor = CivilCard;
                    }
                }
                else if (c is GroupBox gb)
                {
                    gb.ForeColor = CivilBlue;
                    gb.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
                }
                else if (c is WinLabel lbl)
                {
                    if (lbl.ForeColor == Color.FromArgb(61, 61, 61) || lbl.ForeColor == Color.Black)
                        lbl.ForeColor = TextDark;
                }
                else if (c is CheckBox chk)
                {
                    chk.ForeColor = TextDark;
                    chk.FlatStyle = FlatStyle.Flat;
                }
                else if (c is RadioButton rdo)
                {
                    rdo.ForeColor = TextDark;
                    rdo.FlatStyle = FlatStyle.Flat;
                }
                else if (c is TextBox txt)
                {
                    if (txt.BackColor == Color.White)
                    {
                        txt.BackColor = CivilInput;
                        txt.ForeColor = Color.White;
                        txt.BorderStyle = BorderStyle.FixedSingle;
                    }
                }
                else if (c is ComboBox cmb)
                {
                    cmb.BackColor = Color.FromArgb(38, 38, 38);
                    cmb.ForeColor = Color.White;
                    cmb.FlatStyle = FlatStyle.Flat;
                }
                else if (c is ListBox lst)
                {
                    lst.BackColor = CivilInput;
                    lst.ForeColor = TextDark;
                    lst.BorderStyle = BorderStyle.FixedSingle;
                }
                else if (c is TabControl tab)
                {
                    foreach (TabPage page in tab.TabPages)
                    {
                        page.BackColor = CivilDark;
                        page.ForeColor = TextDark;
                        ApplyCivilDarkTheme(page);
                    }
                }

                // Đệ quy cho các container con
                if (c.HasChildren)
                {
                    ApplyCivilDarkTheme(c);
                }
            }
        }
    }
}
