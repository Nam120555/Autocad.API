// ==============================================================================
// NxTheme.cs — Bộ nhận diện giao diện "UI Pro" cho toàn dự án Civil Tool
// Dùng chung cho: NxCommandDialog, NxMessageBox, NxCommandCenter
// Tông màu kế thừa CivilFormBase: Navy #1B2A4A / Trắng / Xám #3D3D3D / Đỏ #D32F2F
// ==============================================================================
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

using Color = System.Drawing.Color;
using Font = System.Drawing.Font;
using FontStyle = System.Drawing.FontStyle;
using Rectangle = System.Drawing.Rectangle;

namespace MyFirstProject.UI
{
    /// <summary>Bảng màu, font và các hàm vẽ dùng chung cho bộ giao diện UI Pro.</summary>
    public static class NxTheme
    {
        // ═══ Màu chủ đạo ═══
        public static readonly Color Navy = Color.FromArgb(27, 42, 74);          // #1B2A4A
        public static readonly Color NavyDeep = Color.FromArgb(16, 28, 52);      // #101C34
        public static readonly Color NavyMid = Color.FromArgb(46, 74, 122);      // #2E4A7A
        public static readonly Color NavySoft = Color.FromArgb(232, 238, 248);   // nền hover / chip
        public static readonly Color NavyLine = Color.FromArgb(198, 210, 230);

        public static readonly Color Surface = Color.White;
        public static readonly Color Canvas = Color.FromArgb(244, 246, 250);     // nền thân hộp thoại
        public static readonly Color Border = Color.FromArgb(214, 219, 228);
        public static readonly Color BorderFocus = Color.FromArgb(46, 74, 122);

        public static readonly Color Text = Color.FromArgb(45, 52, 64);
        public static readonly Color TextMuted = Color.FromArgb(118, 126, 140);
        public static readonly Color TextOnNavy = Color.White;
        public static readonly Color TextOnNavyMuted = Color.FromArgb(190, 203, 226);

        public static readonly Color Success = Color.FromArgb(46, 160, 67);
        public static readonly Color Warning = Color.FromArgb(214, 137, 16);
        public static readonly Color Danger = Color.FromArgb(211, 47, 47);
        public static readonly Color DangerSoft = Color.FromArgb(253, 236, 236);
        public static readonly Color Gold = Color.FromArgb(240, 180, 41);

        // ═══ Font ═══
        public static readonly Font Body = new("Segoe UI", 9.5f);
        public static readonly Font BodyBold = new("Segoe UI", 9.5f, FontStyle.Bold);
        public static readonly Font Small = new("Segoe UI", 8.25f);
        public static readonly Font SmallBold = new("Segoe UI", 8.25f, FontStyle.Bold);
        public static readonly Font Title = new("Segoe UI Semibold", 13f);
        public static readonly Font Section = new("Segoe UI Semibold", 10f);
        public static readonly Font Mono = new("Consolas", 9.5f);
        public static readonly Font MonoBold = new("Consolas", 9.5f, FontStyle.Bold);
        public static readonly Font Glyph = new("Segoe UI Symbol", 15f);
        public static readonly Font Search = new("Segoe UI", 12f);

        // ═══ Hàm vẽ ═══

        /// <summary>Tạo GraphicsPath hình chữ nhật bo góc.</summary>
        public static GraphicsPath RoundRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            if (radius <= 0) { path.AddRectangle(r); return path; }
            int d = radius * 2;
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        /// <summary>Vẽ huy hiệu tròn chứa ký hiệu (icon) của lệnh.</summary>
        public static void DrawBadge(Graphics g, Rectangle r, string glyph, Color back, Color fore)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            using (var b = new SolidBrush(back)) g.FillEllipse(b, r);
            using var f = new Font("Segoe UI Symbol", r.Height * 0.42f, FontStyle.Regular, GraphicsUnit.Pixel);
            TextRenderer.DrawText(g, glyph, f, r, fore,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        }

        /// <summary>Vẽ "chip" nhỏ (nhãn bo tròn) — trả về chiều rộng đã vẽ.</summary>
        public static int DrawChip(Graphics g, string text, Font font, int x, int yCenter, Color back, Color fore)
        {
            var size = TextRenderer.MeasureText(g, text, font, Size.Empty, TextFormatFlags.NoPadding);
            var r = new Rectangle(x, yCenter - (size.Height + 6) / 2, size.Width + 14, size.Height + 6);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = RoundRect(r, r.Height / 2))
            using (var b = new SolidBrush(back))
                g.FillPath(b, path);
            TextRenderer.DrawText(g, text, font, r, fore,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
            return r.Width;
        }

        /// <summary>Bỏ dấu tiếng Việt + viết thường — dùng cho tìm kiếm "doi ten coc" = "Đổi tên cọc".</summary>
        public static string Fold(string? s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            var norm = s.Replace('đ', 'd').Replace('Đ', 'D').Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(norm.Length);
            foreach (var ch in norm)
                if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                    sb.Append(char.ToLowerInvariant(ch));
            return sb.ToString();
        }

        // ═══ Kéo form không viền bằng chuột ═══
        [DllImport("user32.dll")] private static extern bool ReleaseCapture();
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp);

        /// <summary>Cho phép kéo form bằng cách giữ chuột trên control (header).</summary>
        public static void MakeDraggable(Control handle, Form form)
        {
            handle.MouseDown += (s, e) =>
            {
                if (e.Button != MouseButtons.Left) return;
                ReleaseCapture();
                SendMessage(form.Handle, 0xA1 /*WM_NCLBUTTONDOWN*/, (IntPtr)2 /*HTCAPTION*/, IntPtr.Zero);
            };
        }

        /// <summary>Nút phẳng theo theme (primary = nền Navy).</summary>
        public static Button CreateButton(string text, bool primary, int width = 120)
        {
            var b = new Button
            {
                Text = text,
                Width = width,
                Height = 34,
                FlatStyle = FlatStyle.Flat,
                Font = primary ? BodyBold : Body,
                BackColor = primary ? Navy : Surface,
                ForeColor = primary ? TextOnNavy : Text,
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false,
                Margin = new Padding(8, 0, 0, 0)
            };
            b.FlatAppearance.BorderColor = primary ? NavyDeep : Border;
            b.FlatAppearance.MouseOverBackColor = primary ? NavyMid : NavySoft;
            b.FlatAppearance.MouseDownBackColor = primary ? NavyDeep : NavyLine;
            return b;
        }
    }

    /// <summary>
    /// Form nền không viền có bóng đổ + viền 1px — dùng cho các hộp thoại UI Pro.
    /// Kế thừa CivilFormBase nên luôn có chữ ký tác giả (PROJECT_RULES).
    /// </summary>
    public class NxBorderlessForm : CivilFormBase
    {
        public NxBorderlessForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            BackColor = NxTheme.Surface;
            Font = NxTheme.Body;
            ForeColor = NxTheme.Text;
            Padding = new Padding(1);   // chừa 1px để vẽ viền
            KeyPreview = true;
            DoubleBuffered = true;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
        }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ClassStyle |= 0x00020000; // CS_DROPSHADOW
                return cp;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            using var pen = new Pen(NxTheme.NavyMid);
            e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            Invalidate();
        }
    }
}
