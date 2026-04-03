// CivilFormBase.cs — Base class chung cho tất cả Dialog Form trong dự án
// Theme: Navy (#1B2A4A) / Trắng / Xám đậm (#3D3D3D) / Ghi chú đỏ (#D32F2F)

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
    /// Navy / White theme — dùng chung bảng màu, helper methods
    /// </summary>
    public class CivilFormBase : Form
    {
        // ═══ Bảng màu Navy / White ═══
        protected static readonly Color NavyDark = Color.FromArgb(27, 42, 74);       // #1B2A4A — Header, accent chính
        protected static readonly Color NavyMedium = Color.FromArgb(46, 74, 122);    // #2E4A7A — Hover, accent phụ
        protected static readonly Color NavyDeep = Color.FromArgb(16, 28, 52);       // #101C34 — Gradient đậm

        protected static readonly Color BgWhite = Color.FromArgb(255, 255, 255);     // Nền form
        protected static readonly Color BgPanel = Color.FromArgb(245, 247, 250);     // #F5F7FA — Nền panel/section
        protected static readonly Color BgInput = Color.FromArgb(255, 255, 255);     // Nền input

        protected static readonly Color TextDark = Color.FromArgb(61, 61, 61);       // #3D3D3D — Text chính
        protected static readonly Color TextNavy = Color.FromArgb(27, 42, 74);       // Tiêu đề
        protected static readonly Color TextHint = Color.FromArgb(122, 122, 122);    // #7A7A7A — Hint/phụ
        protected static readonly Color TextRed = Color.FromArgb(211, 47, 47);       // #D32F2F — Ghi chú quan trọng
        protected static readonly Color TextDim = Color.FromArgb(122, 122, 122);
        protected static readonly Color TextWhite = Color.White;

        protected static readonly Color BorderLight = Color.FromArgb(208, 213, 221); // #D0D5DD
        protected static readonly Color BorderColor = Color.FromArgb(208, 213, 221);
        protected static readonly Color GreenOk = Color.FromArgb(46, 160, 67);       // Thành công
        protected static readonly Color HoverNavy = Color.FromArgb(232, 238, 248);   // Hover nhạt

        /// <summary>
        /// Khởi tạo form với theme Navy/White chuẩn
        /// </summary>
        protected void InitBaseTheme(string title, int width = 700, int height = 550)
        {
            this.Text = title;
            this.Size = new Size(width, height);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = BgWhite;
            this.ForeColor = TextDark;
            this.Font = new Font("Segoe UI", 9.5f);
        }

        /// <summary>
        /// Tạo Header bar Navy gradient
        /// </summary>
        protected Panel CreateHeader(string title, int height = 50)
        {
            var header = new Panel { Dock = DockStyle.Top, Height = height };
            header.Paint += (s, e) =>
            {
                using var grad = new LinearGradientBrush(
                    header.ClientRectangle, NavyDark, NavyDeep, 0f);
                e.Graphics.FillRectangle(grad, header.ClientRectangle);
            };

            var lbl = new WinLabel
            {
                Text = title,
                Font = new Font("Segoe UI", 13, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                Padding = new Padding(16, 0, 0, 0),
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
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = TextHint,
                AutoSize = false,
                Height = 24,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Top,
                Padding = new Padding(18, 0, 0, 0),
                BackColor = BgPanel
            };
        }

        /// <summary>
        /// Tạo Section (panel có title)
        /// </summary>
        protected Panel CreateSection(string title, int x, int y, int w, int h) => CreateGroupBox(title, x, y, w, h);

        /// <summary>
        /// Tạo GroupBox (panel có title)
        /// </summary>
        protected Panel CreateGroupBox(string title, int x, int y, int w, int h)
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
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = NavyDark,
                Location = new Point(8, 4),
                AutoSize = true
            };
            panel.Controls.Add(lbl);

            return panel;
        }

        /// <summary>
        /// Tạo Label
        /// </summary>
        protected WinLabel CreateLabel(string text, int x, int y, bool autoSize = true)
        {
            return new WinLabel
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = autoSize,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
        }

        /// <summary>
        /// Tạo Label tiêu đề Navy
        /// </summary>
        protected WinLabel CreateTitleLabel(string text, int x, int y)
        {
            return new WinLabel
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                ForeColor = NavyDark,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
        }

        /// <summary>
        /// Tạo Label hint (xám nhạt)
        /// </summary>
        protected WinLabel CreateHintLabel(string text, int x, int y)
        {
            return new WinLabel
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                ForeColor = TextHint,
                Font = new Font("Segoe UI", 8f)
            };
        }

        /// <summary>
        /// Tạo Label ghi chú đỏ
        /// </summary>
        protected WinLabel CreateRedNote(string text, int x, int y)
        {
            return new WinLabel
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                ForeColor = TextRed,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Italic)
            };
        }

        /// <summary>
        /// Tạo TextBox
        /// </summary>
        protected TextBox CreateTextBox(int x, int y, int w = 250, string defaultValue = "")
        {
            return new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(w, 28),
                BackColor = BgInput,
                ForeColor = TextDark,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9.5f),
                Text = defaultValue
            };
        }

        /// <summary>
        /// Tạo ComboBox
        /// </summary>
        protected ComboBox CreateComboBox(int x, int y, int w = 300)
        {
            return new ComboBox
            {
                Location = new Point(x, y),
                Size = new Size(w, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgPanel,
                ForeColor = TextDark,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f)
            };
        }

        /// <summary>
        /// Tạo NumericUpDown
        /// </summary>
        protected NumericUpDown CreateNumeric(int x, int y, int w = 100,
            decimal min = 0, decimal max = 9999, decimal value = 1, int decimals = 0)
        {
            return new NumericUpDown
            {
                Location = new Point(x, y),
                Size = new Size(w, 28),
                Minimum = min,
                Maximum = max,
                Value = value,
                DecimalPlaces = decimals,
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
        }

        /// <summary>
        /// Tạo CheckBox
        /// </summary>
        protected CheckBox CreateCheckBox(string text, int x, int y, bool isChecked = false)
        {
            return new CheckBox
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9f),
                Checked = isChecked,
                FlatStyle = FlatStyle.Flat
            };
        }

        /// <summary>
        /// Tạo Button chuẩn
        /// </summary>
        protected Button CreateButton(string text, int x, int y, int w = 120, int h = 34, Control? parent = null)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                FlatStyle = FlatStyle.Flat,
                BackColor = BgPanel,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9f),
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter
            };
            btn.FlatAppearance.BorderColor = BorderLight;
            btn.FlatAppearance.MouseOverBackColor = NavyMedium;
            btn.MouseEnter += (s, e) => btn.ForeColor = Color.White;
            btn.MouseLeave += (s, e) => btn.ForeColor = TextDark;
            parent?.Controls.Add(btn);
            return btn;
        }

        /// <summary>
        /// Tạo Button Navy chính (Save, Execute)
        /// </summary>
        protected Button CreatePrimaryButton(string text, int x, int y, int w = 140, int h = 34, Control? parent = null)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                FlatStyle = FlatStyle.Flat,
                BackColor = NavyDark,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter
            };
            btn.FlatAppearance.BorderColor = NavyDeep;
            btn.FlatAppearance.MouseOverBackColor = NavyMedium;
            parent?.Controls.Add(btn);
            return btn;
        }

        /// <summary>
        /// Tạo Button xanh lá (Thực hiện / OK)
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
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(35, 134, 54);
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(40, 180, 75);
            parent?.Controls.Add(btn);
            return btn;
        }

        /// <summary>
        /// Tạo Bottom bar chuẩn
        /// </summary>
        protected Panel CreateBottomBar(int height = 50)
        {
            return new Panel
            {
                Dock = DockStyle.Bottom,
                Height = height,
                BackColor = BgPanel,
                Padding = new Padding(12, 8, 12, 8)
            };
        }

        /// <summary>
        /// Tạo DataGridView chuẩn
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
                BackgroundColor = BgWhite,
                GridColor = BorderLight,
                BorderStyle = BorderStyle.FixedSingle,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = BgWhite,
                    ForeColor = TextDark,
                    SelectionBackColor = NavyDark,
                    SelectionForeColor = Color.White,
                    Font = new Font("Segoe UI", 9f)
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = BgPanel,
                    ForeColor = NavyDark,
                    Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                EnableHeadersVisualStyles = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
        }

        /// <summary>
        /// Tạo ListView chuẩn
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
                BackColor = BgWhite,
                ForeColor = TextDark,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9f)
            };
        }
    }
}
