// ==============================================================================
// NxMessageBox.cs — Hộp thông báo chuyên nghiệp thay cho MessageBox mặc định
//   NxMessageBox.Info("Tiêu đề", "Nội dung");
//   NxMessageBox.Success("Hoàn thành", "Đã đổi tên 25 cọc.");
//   NxMessageBox.Warning("Chú ý", "Chưa chọn đối tượng.");
//   NxMessageBox.Error("Lỗi", ex.Message, ex.ToString());   // có nút "Chi tiết"
//   if (NxMessageBox.Confirm("Xác nhận", "Xóa 10 đối tượng?")) { ... }
// ==============================================================================
using System;
using System.Drawing;
using System.Windows.Forms;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

using Color = System.Drawing.Color;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Rectangle = System.Drawing.Rectangle;
using WinLabel = System.Windows.Forms.Label;

namespace MyFirstProject.UI
{
    public enum NxMessageKind { Info, Success, Warning, Error, Question }

    public class NxMessageBox : NxBorderlessForm
    {
        private const int BaseWidth = 460;
        private readonly TextBox? _details;
        private readonly int _collapsedHeight;

        private NxMessageBox(NxMessageKind kind, string title, string message, string? details, bool askYesNo)
        {
            var (accent, glyph) = kind switch
            {
                NxMessageKind.Success => (NxTheme.Success, "✔"),
                NxMessageKind.Warning => (NxTheme.Warning, "!"),
                NxMessageKind.Error => (NxTheme.Danger, "✖"),
                NxMessageKind.Question => (NxTheme.NavyMid, "?"),
                _ => (NxTheme.Navy, "i")
            };

            Text = title;
            Width = BaseWidth;

            // ── Đo chiều cao nội dung ──
            int textW = BaseWidth - 96;
            var msgSize = TextRenderer.MeasureText(message, NxTheme.Body, new Size(textW, 0), TextFormatFlags.WordBreak);
            int contentH = Math.Max(56, 34 + msgSize.Height + 14);

            var content = new Panel { Dock = DockStyle.Top, Height = contentH + 20, BackColor = NxTheme.Surface };
            content.Paint += (s, e) =>
            {
                using (var b = new SolidBrush(accent)) e.Graphics.FillRectangle(b, 0, 0, 4, content.Height);
                NxTheme.DrawBadge(e.Graphics, new Rectangle(22, 22, 38, 38), glyph, accent, Color.White);
                TextRenderer.DrawText(e.Graphics, title, NxTheme.Section, new Point(76, 22), NxTheme.Navy, TextFormatFlags.NoPadding);
            };
            NxTheme.MakeDraggable(content, this);

            var lblMsg = new WinLabel
            {
                Text = message,
                Font = NxTheme.Body,
                ForeColor = NxTheme.Text,
                Location = new Point(76, 48),
                Size = new Size(textW, msgSize.Height + 4),
                BackColor = Color.Transparent
            };
            content.Controls.Add(lblMsg);

            var footer = new Panel { Dock = DockStyle.Top, Height = 56, BackColor = NxTheme.Canvas };

            var btnOk = NxTheme.CreateButton(askYesNo ? "Đồng ý" : "OK", true, 100);
            btnOk.DialogResult = askYesNo ? DialogResult.Yes : DialogResult.OK;
            btnOk.Location = new Point(BaseWidth - 2 - 14 - btnOk.Width, 11);
            footer.Controls.Add(btnOk);
            AcceptButton = btnOk;

            if (askYesNo)
            {
                var btnNo = NxTheme.CreateButton("Không", false, 90);
                btnNo.DialogResult = DialogResult.No;
                btnNo.Location = new Point(btnOk.Left - 8 - btnNo.Width, 11);
                footer.Controls.Add(btnNo);
                CancelButton = btnNo;
            }
            else
            {
                CancelButton = btnOk;
            }

            if (!string.IsNullOrWhiteSpace(details))
            {
                var lnk = new LinkLabel
                {
                    Text = "▸ Chi tiết",
                    AutoSize = true,
                    Font = NxTheme.Small,
                    LinkColor = NxTheme.NavyMid,
                    LinkBehavior = LinkBehavior.HoverUnderline,
                    Location = new Point(18, 20)
                };
                var lnkCopy = new LinkLabel
                {
                    Text = "⧉ Sao chép",
                    AutoSize = true,
                    Font = NxTheme.Small,
                    LinkColor = NxTheme.NavyMid,
                    LinkBehavior = LinkBehavior.HoverUnderline,
                    Location = new Point(96, 20)
                };
                lnkCopy.LinkClicked += (s, e) => { try { Clipboard.SetText(title + "\r\n" + message + "\r\n\r\n" + details); } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); } };
                footer.Controls.Add(lnk);
                footer.Controls.Add(lnkCopy);

                _details = new TextBox
                {
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Both,
                    WordWrap = false,
                    Font = new Font("Consolas", 8.5f),
                    BackColor = Color.FromArgb(250, 250, 252),
                    ForeColor = NxTheme.Text,
                    BorderStyle = BorderStyle.None,
                    Dock = DockStyle.Fill,
                    Text = details!.Replace("\r\n", "\n").Replace("\n", "\r\n"),
                    Visible = false
                };
                lnk.LinkClicked += (s, e) =>
                {
                    _details.Visible = !_details.Visible;
                    lnk.Text = _details.Visible ? "▾ Thu gọn" : "▸ Chi tiết";
                    Height += _details.Visible ? 180 : -180;
                };
            }

            if (_details != null) Controls.Add(_details);
            Controls.Add(footer);
            Controls.Add(content);
            _details?.BringToFront();

            _collapsedHeight = content.Height + footer.Height + 2;
            Height = _collapsedHeight;
            TopMost = true;
        }

        // ══════════════ API tĩnh ══════════════

        public static void Info(string title, string message) => Show(NxMessageKind.Info, title, message);
        public static void Success(string title, string message) => Show(NxMessageKind.Success, title, message);
        public static void Warning(string title, string message) => Show(NxMessageKind.Warning, title, message);
        public static void Error(string title, string message, string? details = null) => Show(NxMessageKind.Error, title, message, details);

        public static bool Confirm(string title, string message)
        {
            using var f = new NxMessageBox(NxMessageKind.Question, title, message, null, true);
            return Display(f) == DialogResult.Yes;
        }

        public static DialogResult Show(NxMessageKind kind, string title, string message, string? details = null)
        {
            using var f = new NxMessageBox(kind, title, message, details, false);
            return Display(f);
        }

        private static DialogResult Display(Form f)
        {
            try { return AcadApp.ShowModalDialog(f); }
            catch (System.Exception) { return f.ShowDialog(); }
        }
    }
}
