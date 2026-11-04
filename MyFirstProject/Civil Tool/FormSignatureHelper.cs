using System;
using System.Drawing;
using System.Windows.Forms;

namespace MyFirstProject
{
    public static class FormSignatureHelper
    {
        private static string SignatureText => MyFirstProject.Core.AppInfo.Signature + "   |   v" + MyFirstProject.Core.AppInfo.Version;

        public static void AddSignature(Form form)
        {
            form.Shown += (s, e) =>
            {
                // 1. Tăng chiều cao Form thêm 30px để có không gian chứa chữ ký
                form.Height += 30;
                if (form.MinimumSize.Height > 0)
                {
                    form.MinimumSize = new Size(form.MinimumSize.Width, form.MinimumSize.Height + 30);
                }

                // 2. Tạo Strip chứa chữ ký phong cách Autodesk Dark
                Panel strip = new Panel
                {
                    Dock = DockStyle.Bottom,
                    Height = 30,
                    BackColor = Color.FromArgb(24, 24, 24), // #181818 - Dark Studio Footer
                    Padding = new Padding(0)
                };

                // Vẽ viền trên siêu mỏng
                strip.Paint += (ps, pe) =>
                {
                    using Pen p = new Pen(Color.FromArgb(48, 48, 48), 1);
                    pe.Graphics.DrawLine(p, 0, 0, strip.Width, 0);
                };

                // 3. Tạo Label chữ ký
                Label lbl = new Label
                {
                    Text = SignatureText,
                    ForeColor = Color.FromArgb(150, 150, 150), // Muted silver
                    Font = new Font("Segoe UI", 8.25f, FontStyle.Regular),
                    AutoSize = true,
                    BackColor = Color.Transparent,
                    Cursor = Cursors.Default
                };

                strip.Controls.Add(lbl);
                form.Controls.Add(strip);

                // 4. SendToBack() đảm bảo panel này nằm sát mép viền nhất
                strip.SendToBack();

                // 5. Căn giữa Label lúc hiển thị và khi resize
                Action centerLabel = () =>
                {
                    int x = (strip.ClientSize.Width - lbl.PreferredWidth) / 2;
                    int y = (strip.ClientSize.Height - lbl.PreferredHeight) / 2 + 1;
                    if (x < 5) x = 5;
                    lbl.Location = new Point(x, y);
                };

                centerLabel();
                strip.Resize += (rs, re) => centerLabel();
            };
        }
    }
}
