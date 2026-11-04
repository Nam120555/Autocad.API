// ==============================================================================
// NxCommandDialog.cs — "Box lệnh" chuyên nghiệp dùng chung cho mọi lệnh
// ------------------------------------------------------------------------------
// Tính năng:
//   • Header Navy có icon, tên lệnh, mô tả; kéo thả được; bóng đổ
//   • Nhóm thông số theo Section (thẻ bo góc), nhãn + ô nhập + đơn vị
//   • Ô nhập: Text, Số (kiểm tra min/max), Danh sách, Checkbox, File,
//     Lấy khoảng cách / Lấy điểm trực tiếp trên bản vẽ (ẩn hộp thoại khi pick)
//   • Xem trước kết quả (Preview) cập nhật theo thời gian thực
//   • Kiểm tra dữ liệu (Validation) — báo lỗi ngay trên hộp thoại
//   • Tự nhớ giá trị lần trước cho từng lệnh (%AppData%\CivilTool\Dialogs)
//   • Phím tắt: Enter = Thực hiện, Esc = Hủy, F1 = Hướng dẫn
//   • Chữ ký tác giả tự động (kế thừa CivilFormBase)
//
// Cách dùng (ví dụ):
//   var dlg = new NxCommandDialog("CTS_DoiTenCoc", "Đổi tên cọc", "Đổi tên hàng loạt cọc", "⛳");
//   dlg.AddSection("Định dạng");
//   dlg.AddText("prefix", "Tiền tố", "Km");
//   dlg.AddNumber("dec", "Số lẻ thập phân", 2, 0, 4, 0);
//   dlg.SetPreview(d => "Km1+234.56");
//   if (!dlg.ShowAcad()) return;
//   string prefix = dlg.GetString("prefix");
// ==============================================================================
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;

using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

using Color = System.Drawing.Color;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Rectangle = System.Drawing.Rectangle;
using ContentAlignment = System.Drawing.ContentAlignment;
using WinLabel = System.Windows.Forms.Label;
using FontStyle = System.Drawing.FontStyle;

namespace MyFirstProject.UI
{
    /// <summary>Hộp thoại lệnh chuẩn UI Pro (xây dựng bằng code, không cần Designer).</summary>
    public class NxCommandDialog : NxBorderlessForm
    {
        // ─── Mô tả một ô nhập ───
        private sealed class Field
        {
            public string Key = "";
            public string Label = "";
            public string DefaultValue = "";
            public Func<string> Read = () => "";
            public Action<string> Write = _ => { };
            public Func<string?>? Check;          // trả về thông báo lỗi riêng của ô (nếu có)
            public NxInputBox? Host;              // để tô đỏ khi lỗi
        }

        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        private readonly string _commandName;
        private readonly List<Field> _fields = new();
        private readonly Dictionary<string, Field> _byKey = new(StringComparer.OrdinalIgnoreCase);
        private readonly ToolTip _tip = new() { InitialDelay = 300, ReshowDelay = 100, AutoPopDelay = 12000 };

        private readonly Panel _header;
        private readonly Panel _body;
        private readonly TableLayoutPanel _stack;
        private readonly Panel _footer;
        private readonly WinLabel _lblError;
        private readonly Button _btnRun;
        private readonly Button _btnCancel;

        private TableLayoutPanel? _grid;          // lưới của section hiện tại
        private WinLabel? _lblPreview;
        private Func<NxCommandDialog, string>? _preview;
        private Func<NxCommandDialog, string?>? _validator;
        private string? _helpText;
        private bool _loading;

        /// <summary>Tự động lưu / nạp lại giá trị lần trước (mặc định: bật).</summary>
        public bool RememberValues { get; set; } = true;

        /// <param name="commandName">Tên lệnh AutoCAD (dùng làm khóa lưu cấu hình).</param>
        /// <param name="title">Tiêu đề hiển thị.</param>
        /// <param name="subtitle">Mô tả ngắn chức năng.</param>
        /// <param name="glyph">Ký hiệu icon (emoji / ký tự Segoe UI Symbol).</param>
        /// <param name="width">Chiều rộng hộp thoại.</param>
        public NxCommandDialog(string commandName, string title, string subtitle = "", string glyph = "⚙", int width = 540)
        {
            _commandName = commandName;
            Text = title;
            Width = width;
            MinimumSize = new Size(420, 240);

            // ═══ HEADER ═══
            _header = new Panel { Dock = DockStyle.Top, Height = 72, BackColor = NxTheme.Navy };
            _header.Paint += (s, e) =>
            {
                var g = e.Graphics;
                using (var grad = new LinearGradientBrush(_header.ClientRectangle, NxTheme.Navy, NxTheme.NavyDeep, 0f))
                    g.FillRectangle(grad, _header.ClientRectangle);

                NxTheme.DrawBadge(g, new Rectangle(18, 16, 40, 40), glyph, NxTheme.NavyMid, Color.White);

                TextRenderer.DrawText(g, title, NxTheme.Title, new Point(70, 13), NxTheme.TextOnNavy, TextFormatFlags.NoPadding);
                if (!string.IsNullOrWhiteSpace(subtitle))
                {
                    var r = new Rectangle(71, 40, _header.Width - 71 - 52, 22);
                    TextRenderer.DrawText(g, subtitle, NxTheme.Small, r, NxTheme.TextOnNavyMuted,
                        TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);
                }
                // Chip tên lệnh ở góc dưới phải
                var chipFont = NxTheme.SmallBold;
                var chipW = TextRenderer.MeasureText(commandName, chipFont).Width + 4;
                NxTheme.DrawChip(g, commandName.ToUpperInvariant(), chipFont, _header.Width - chipW - 30, 58,
                    Color.FromArgb(60, 255, 255, 255), Color.White);
            };
            _header.Resize += (s, e) => _header.Invalidate();
            NxTheme.MakeDraggable(_header, this);

            var btnClose = new WinLabel
            {
                Text = "✕",
                Font = new Font("Segoe UI", 11f),
                ForeColor = NxTheme.TextOnNavyMuted,
                BackColor = Color.Transparent,
                Size = new Size(36, 30),
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnClose.MouseEnter += (s, e) => { btnClose.BackColor = NxTheme.Danger; btnClose.ForeColor = Color.White; };
            btnClose.MouseLeave += (s, e) => { btnClose.BackColor = Color.Transparent; btnClose.ForeColor = NxTheme.TextOnNavyMuted; };
            btnClose.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            _header.Controls.Add(btnClose);
            _header.Resize += (s, e) => btnClose.Location = new Point(_header.Width - btnClose.Width - 2, 2);

            // ═══ BODY ═══
            _body = new Panel { Dock = DockStyle.Fill, BackColor = NxTheme.Canvas, AutoScroll = true, Padding = new Padding(16, 14, 16, 6) };
            _stack = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                BackColor = Color.Transparent,
                Padding = new Padding(0)
            };
            _stack.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _body.Controls.Add(_stack);

            // ═══ FOOTER ═══
            _footer = new Panel { Dock = DockStyle.Bottom, Height = 58, BackColor = NxTheme.Surface };
            _footer.Paint += (s, e) =>
            {
                using var p = new Pen(NxTheme.Border);
                e.Graphics.DrawLine(p, 0, 0, _footer.Width, 0);
            };

            _lblError = new WinLabel
            {
                AutoSize = false,
                ForeColor = NxTheme.Danger,
                Font = NxTheme.SmallBold,
                TextAlign = ContentAlignment.MiddleLeft,
                Location = new Point(14, 6),
                Size = new Size(10, 18),
                Visible = false
            };

            var lnkReset = new LinkLabel
            {
                Text = "↺ Mặc định",
                AutoSize = true,
                Font = NxTheme.Small,
                LinkColor = NxTheme.NavyMid,
                ActiveLinkColor = NxTheme.Navy,
                LinkBehavior = LinkBehavior.HoverUnderline,
                Location = new Point(14, 30)
            };
            lnkReset.LinkClicked += (s, e) => ResetDefaults();
            _tip.SetToolTip(lnkReset, "Khôi phục toàn bộ giá trị mặc định");

            var lblKeys = new WinLabel
            {
                Text = "Enter: thực hiện · Esc: hủy · F1: hướng dẫn",
                AutoSize = true,
                Font = NxTheme.Small,
                ForeColor = NxTheme.TextMuted,
                Location = new Point(100, 30)
            };

            _btnCancel = NxTheme.CreateButton("Hủy", false, 92);
            _btnCancel.DialogResult = DialogResult.Cancel;
            _btnRun = NxTheme.CreateButton("▶  Thực hiện", true, 132);
            _btnRun.Click += (s, e) => TryAccept();

            _footer.Controls.AddRange(new Control[] { _lblError, lnkReset, lblKeys, _btnCancel, _btnRun });
            _footer.Resize += (s, e) =>
            {
                _btnRun.Location = new Point(_footer.Width - _btnRun.Width - 14, 14);
                _btnCancel.Location = new Point(_btnRun.Left - _btnCancel.Width - 8, 14);
                _lblError.Width = Math.Max(10, _btnCancel.Left - 24);
                lblKeys.Visible = lblKeys.Right < _btnCancel.Left - 8;
            };

            Controls.Add(_header);
            Controls.Add(_footer);
            Controls.Add(_body);
            _body.BringToFront();

            AcceptButton = _btnRun;
            CancelButton = _btnCancel;
        }

        // ══════════════════════════════════════════════════════════════════
        //  API XÂY DỰNG HỘP THOẠI
        // ══════════════════════════════════════════════════════════════════

        /// <summary>Bắt đầu một nhóm thông số mới (thẻ có tiêu đề).</summary>
        public NxCommandDialog AddSection(string title, string? description = null)
        {
            var card = new NxCard(title, description) { Margin = new Padding(0, 0, 0, 12) };
            _grid = card.Grid;
            _stack.Controls.Add(card);
            return this;
        }

        /// <summary>Ô nhập chữ.</summary>
        public NxCommandDialog AddText(string key, string label, string defaultValue = "", string? hint = null,
            bool required = false)
        {
            var host = new NxInputBox();
            host.Box.Text = defaultValue;
            var f = NewField(key, label, defaultValue);
            f.Read = () => host.Box.Text;
            f.Write = v => host.Box.Text = v;
            f.Host = host;
            if (required) f.Check = () => string.IsNullOrWhiteSpace(host.Box.Text) ? $"\"{label}\" không được để trống." : null;
            host.Box.TextChanged += (s, e) => OnValueChanged();
            AddRow(label, host, null, hint);
            return this;
        }

        /// <summary>Ô nhập số có kiểm tra giới hạn. decimals = số chữ số lẻ hiển thị.</summary>
        public NxCommandDialog AddNumber(string key, string label, double defaultValue, double min = double.MinValue,
            double max = double.MaxValue, int decimals = 2, string? unit = null, string? hint = null)
        {
            var host = new NxInputBox();
            host.Box.TextAlign = HorizontalAlignment.Right;
            string fmt = "F" + Math.Max(0, decimals);
            string def = defaultValue.ToString(fmt, Inv);
            host.Box.Text = def;
            var f = NewField(key, label, def);
            f.Read = () => host.Box.Text.Trim().Replace(',', '.');
            f.Write = v => host.Box.Text = v;
            f.Host = host;
            f.Check = () =>
            {
                if (!double.TryParse(f.Read(), NumberStyles.Float, Inv, out var d)) return $"\"{label}\" phải là số.";
                if (decimals == 0 && Math.Abs(d - Math.Round(d)) > 1e-9) return $"\"{label}\" phải là số nguyên.";
                if (d < min) return $"\"{label}\" phải ≥ {min.ToString("G", Inv)}.";
                if (d > max) return $"\"{label}\" phải ≤ {max.ToString("G", Inv)}.";
                return null;
            };
            host.Box.TextChanged += (s, e) => OnValueChanged();
            host.Box.KeyPress += (s, e) =>
            {
                if (char.IsControl(e.KeyChar) || char.IsDigit(e.KeyChar) || "-.,eE+".IndexOf(e.KeyChar) >= 0) return;
                e.Handled = true;
            };
            AddRow(label, host, unit != null ? MakeUnit(unit) : null, hint);
            return this;
        }

        /// <summary>Danh sách lựa chọn (lưu theo chuỗi được chọn).</summary>
        public NxCommandDialog AddChoice(string key, string label, IEnumerable<string> items, int defaultIndex = 0,
            string? hint = null)
        {
            var list = items.ToList();
            var cb = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Font = NxTheme.Body,
                BackColor = NxTheme.Surface,
                ForeColor = NxTheme.Text,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 3, 0, 3),
                IntegralHeight = false,
                MaxDropDownItems = 14
            };
            cb.Items.AddRange(list.Cast<object>().ToArray());
            if (list.Count > 0) cb.SelectedIndex = Math.Clamp(defaultIndex, 0, list.Count - 1);
            var f = NewField(key, label, list.Count > 0 ? list[Math.Clamp(defaultIndex, 0, list.Count - 1)] : "");
            f.Read = () => cb.SelectedItem?.ToString() ?? "";
            f.Write = v => { int i = list.IndexOf(v); if (i >= 0) cb.SelectedIndex = i; };
            cb.SelectedIndexChanged += (s, e) => OnValueChanged();
            AddRow(label, cb, null, hint);
            return this;
        }

        /// <summary>Hộp kiểm (Có / Không).</summary>
        public NxCommandDialog AddCheck(string key, string text, bool defaultValue = false, string? hint = null)
        {
            var chk = new CheckBox
            {
                Text = text,
                Checked = defaultValue,
                AutoSize = true,
                Font = NxTheme.Body,
                ForeColor = NxTheme.Text,
                Margin = new Padding(0, 6, 0, 4),
                Cursor = Cursors.Hand
            };
            var f = NewField(key, text, defaultValue ? "1" : "0");
            f.Read = () => chk.Checked ? "1" : "0";
            f.Write = v => chk.Checked = v == "1" || v.Equals("true", StringComparison.OrdinalIgnoreCase);
            chk.CheckedChanged += (s, e) => OnValueChanged();
            AddRow(null, chk, null, hint);
            return this;
        }

        /// <summary>Ô số + nút lấy khoảng cách trực tiếp trên bản vẽ.</summary>
        public NxCommandDialog AddPickDistance(string key, string label, double defaultValue, string prompt,
            double min = 0, double max = double.MaxValue, int decimals = 3, string? unit = null, string? hint = null)
        {
            AddNumber(key, label, defaultValue, min, max, decimals, null, hint);
            var f = _byKey[key];
            var btn = MakePickButton("Lấy khoảng cách trên bản vẽ");
            btn.Click += (s, e) =>
            {
                var ed = AcadApp.DocumentManager.MdiActiveDocument?.Editor;
                if (ed == null) return;
                using (var ui = ed.StartUserInteraction(this))
                {
                    var res = ed.GetDistance(new PromptDistanceOptions("\n" + prompt) { AllowNegative = min < 0, AllowZero = min <= 0 });
                    ui.End();
                    if (res.Status == PromptStatus.OK) f.Write(res.Value.ToString("F" + decimals, Inv));
                }
                Activate();
            };
            ReplaceAction(f, unit, btn);
            return this;
        }

        /// <summary>Ô tọa độ "X, Y, Z" + nút chọn điểm trên bản vẽ.</summary>
        public NxCommandDialog AddPickPoint(string key, string label, string prompt, Point3d? defaultValue = null,
            bool required = true, string? hint = null)
        {
            var host = new NxInputBox();
            string def = defaultValue.HasValue ? FormatPoint(defaultValue.Value) : "";
            host.Box.Text = def;
            host.Box.ReadOnly = false;
            var f = NewField(key, label, def);
            f.Read = () => host.Box.Text.Trim();
            f.Write = v => host.Box.Text = v;
            f.Host = host;
            f.Check = () =>
            {
                var t = f.Read();
                if (t.Length == 0) return required ? $"Chưa chọn \"{label}\" (bấm nút ⌖ để chọn trên bản vẽ)." : null;
                return TryParsePoint(t, out _) ? null : $"\"{label}\" không đúng dạng X, Y, Z.";
            };
            host.Box.TextChanged += (s, e) => OnValueChanged();

            var btn = MakePickButton("Chọn điểm trên bản vẽ");
            btn.Click += (s, e) =>
            {
                var ed = AcadApp.DocumentManager.MdiActiveDocument?.Editor;
                if (ed == null) return;
                using (var ui = ed.StartUserInteraction(this))
                {
                    var res = ed.GetPoint(new PromptPointOptions("\n" + prompt));
                    ui.End();
                    if (res.Status == PromptStatus.OK) f.Write(FormatPoint(res.Value));
                }
                Activate();
            };
            AddRow(label, host, btn, hint);
            return this;
        }

        /// <summary>Ô chọn file (Mở / Lưu).</summary>
        public NxCommandDialog AddFile(string key, string label, string filter = "Tất cả|*.*", bool save = false,
            string defaultValue = "", string? hint = null)
        {
            var host = new NxInputBox();
            host.Box.Text = defaultValue;
            var f = NewField(key, label, defaultValue);
            f.Read = () => host.Box.Text.Trim();
            f.Write = v => host.Box.Text = v;
            f.Host = host;
            f.Check = () =>
            {
                var p = f.Read();
                if (p.Length == 0) return $"Chưa chọn \"{label}\".";
                if (!save && !File.Exists(p)) return $"Không tìm thấy file: {Path.GetFileName(p)}";
                return null;
            };
            host.Box.TextChanged += (s, e) => OnValueChanged();

            var btn = MakeSquareButton("…", save ? "Chọn nơi lưu file" : "Chọn file");
            btn.Click += (s, e) =>
            {
                using FileDialog dlg = save ? new SaveFileDialog() : new OpenFileDialog();
                dlg.Filter = filter;
                try
                {
                    var cur = f.Read();
                    if (cur.Length > 0)
                    {
                        dlg.InitialDirectory = Path.GetDirectoryName(cur) ?? "";
                        dlg.FileName = Path.GetFileName(cur);
                    }
                }
                catch { /* đường dẫn cũ không hợp lệ */ }
                if (dlg.ShowDialog(this) == DialogResult.OK) f.Write(dlg.FileName);
            };
            AddRow(label, host, btn, hint);
            return this;
        }

        /// <summary>Dòng ghi chú (thông tin / cảnh báo đỏ).</summary>
        public NxCommandDialog AddNote(string text, bool important = false)
        {
            EnsureGrid();
            var lbl = new WinLabel
            {
                Text = (important ? "⚠  " : "ⓘ  ") + text,
                AutoSize = true,
                MaximumSize = new Size(Width - 90, 0),
                Font = important ? new Font("Segoe UI", 8.5f, FontStyle.Italic) : NxTheme.Small,
                ForeColor = important ? NxTheme.Danger : NxTheme.TextMuted,
                Margin = new Padding(0, 4, 0, 4)
            };
            int row = _grid!.RowCount++;
            _grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _grid.Controls.Add(lbl, 0, row);
            _grid.SetColumnSpan(lbl, 3);
            return this;
        }

        /// <summary>Khung "Xem trước" — hàm được gọi lại mỗi khi người dùng thay đổi giá trị.</summary>
        public NxCommandDialog SetPreview(Func<NxCommandDialog, string> preview, string title = "Xem trước")
        {
            _preview = preview;
            var card = new NxCard(title, null, accent: NxTheme.Success) { Margin = new Padding(0, 0, 0, 12) };
            _lblPreview = new WinLabel
            {
                AutoSize = true,
                MaximumSize = new Size(Width - 90, 0),
                Font = NxTheme.MonoBold,
                ForeColor = NxTheme.Navy,
                Margin = new Padding(0, 4, 0, 4)
            };
            card.Grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            card.Grid.Controls.Add(_lblPreview, 0, card.Grid.RowCount++);
            card.Grid.SetColumnSpan(_lblPreview, 3);
            _stack.Controls.Add(card);
            return this;
        }

        /// <summary>Hàm kiểm tra tổng thể (trả về chuỗi lỗi, hoặc null nếu hợp lệ).</summary>
        public NxCommandDialog SetValidator(Func<NxCommandDialog, string?> validator)
        {
            _validator = validator;
            return this;
        }

        /// <summary>Nội dung hướng dẫn hiển thị khi bấm F1.</summary>
        public NxCommandDialog SetHelp(string helpText)
        {
            _helpText = helpText;
            return this;
        }

        /// <summary>Đổi chữ trên nút thực hiện.</summary>
        public NxCommandDialog SetRunText(string text)
        {
            _btnRun.Text = text;
            return this;
        }

        // ══════════════════════════════════════════════════════════════════
        //  ĐỌC GIÁ TRỊ
        // ══════════════════════════════════════════════════════════════════

        public string GetString(string key) => _byKey.TryGetValue(key, out var f) ? f.Read() : "";

        public double GetDouble(string key, double fallback = 0)
            => double.TryParse(GetString(key), NumberStyles.Float, Inv, out var d) ? d : fallback;

        public int GetInt(string key, int fallback = 0) => (int)Math.Round(GetDouble(key, fallback));

        public bool GetBool(string key) => GetString(key) == "1";

        public Point3d? GetPoint(string key) => TryParsePoint(GetString(key), out var p) ? p : null;

        /// <summary>Chỉ số của lựa chọn trong AddChoice (so khớp chuỗi).</summary>
        public int GetChoiceIndex(string key, IList<string> items) => items.IndexOf(GetString(key));

        // ══════════════════════════════════════════════════════════════════
        //  HIỂN THỊ
        // ══════════════════════════════════════════════════════════════════

        /// <summary>Hiển thị hộp thoại modal trong AutoCAD. Trả về true nếu người dùng bấm "Thực hiện".</summary>
        public bool ShowAcad()
        {
            try
            {
                return AcadApp.ShowModalDialog(this) == DialogResult.OK;
            }
            catch (System.Exception)
            {
                return ShowDialog() == DialogResult.OK;
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            _loading = true;
            if (RememberValues) LoadSaved();
            _loading = false;

            // Tính chiều cao vừa nội dung (tối đa 85% màn hình)
            _stack.PerformLayout();
            int content = _stack.GetPreferredSize(new Size(_body.ClientSize.Width - _body.Padding.Horizontal, 0)).Height;
            int wanted = _header.Height + _footer.Height + content + _body.Padding.Vertical + 4;
            int maxH = (int)(Screen.FromPoint(Cursor.Position).WorkingArea.Height * 0.85) - 30;
            Height = Math.Max(MinimumSize.Height, Math.Min(wanted, maxH));

            base.OnLoad(e);   // CivilFormBase → gắn chữ ký tác giả
            RefreshState();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            // Focus ô nhập đầu tiên
            var first = _fields.Select(f => f.Host?.Box as Control).FirstOrDefault(c => c != null);
            first?.Focus();
            (first as TextBox)?.SelectAll();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.F1)
            {
                NxMessageBox.Info(Text, _helpText ?? "Nhập các thông số rồi bấm \"Thực hiện\" (Enter).\n" +
                    "Nút ⌖ cho phép lấy giá trị trực tiếp trên bản vẽ.\n" +
                    "Giá trị sẽ được ghi nhớ cho lần dùng sau.");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // ══════════════════════════════════════════════════════════════════
        //  NỘI BỘ
        // ══════════════════════════════════════════════════════════════════

        private Field NewField(string key, string label, string def)
        {
            var f = new Field { Key = key, Label = label, DefaultValue = def };
            _fields.Add(f);
            _byKey[key] = f;
            return f;
        }

        private void EnsureGrid()
        {
            if (_grid == null) AddSection("Thông số");
        }

        private void AddRow(string? label, Control editor, Control? action, string? hint)
        {
            EnsureGrid();
            var g = _grid!;
            int row = g.RowCount++;
            g.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            if (label != null)
            {
                var lbl = new WinLabel
                {
                    Text = label,
                    AutoSize = true,
                    Font = NxTheme.Body,
                    ForeColor = NxTheme.Text,
                    Anchor = AnchorStyles.Left,
                    Margin = new Padding(0, 6, 10, 6)
                };
                g.Controls.Add(lbl, 0, row);
                if (hint != null) _tip.SetToolTip(lbl, hint);
                editor.Dock = DockStyle.Fill;
                g.Controls.Add(editor, 1, row);
            }
            else
            {
                g.Controls.Add(editor, 0, row);
                g.SetColumnSpan(editor, 2);
            }

            if (action != null)
            {
                action.Anchor = AnchorStyles.Left;
                g.Controls.Add(action, 2, row);
            }

            if (hint != null)
            {
                _tip.SetToolTip(editor, hint);
                if (editor is NxInputBox ib) _tip.SetToolTip(ib.Box, hint);
                var h = new WinLabel
                {
                    Text = hint,
                    AutoSize = true,
                    MaximumSize = new Size(Width - 260, 0),
                    Font = NxTheme.Small,
                    ForeColor = NxTheme.TextMuted,
                    Margin = new Padding(label != null ? 1 : 20, 0, 0, 4)
                };
                int hr = g.RowCount++;
                g.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                g.Controls.Add(h, label != null ? 1 : 0, hr);
                g.SetColumnSpan(h, label != null ? 2 : 3);
            }
        }

        private void ReplaceAction(Field f, string? unit, Button pick)
        {
            if (f.Host == null || _grid == null) return;
            var cell = _grid.GetPositionFromControl(f.Host);
            var flow = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Margin = new Padding(0),
                Anchor = AnchorStyles.Left
            };
            flow.Controls.Add(pick);
            if (unit != null) flow.Controls.Add(MakeUnit(unit));
            _grid.Controls.Add(flow, 2, cell.Row);
        }

        private static WinLabel MakeUnit(string unit) => new()
        {
            Text = unit,
            AutoSize = true,
            Font = NxTheme.Small,
            ForeColor = NxTheme.TextMuted,
            Margin = new Padding(6, 9, 0, 0)
        };

        private Button MakePickButton(string tip) => MakeSquareButton("⌖", tip);

        private Button MakeSquareButton(string text, string tip)
        {
            var b = new Button
            {
                Text = text,
                Size = new Size(32, 30),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Symbol", 11f),
                ForeColor = NxTheme.Navy,
                BackColor = NxTheme.NavySoft,
                Cursor = Cursors.Hand,
                Margin = new Padding(6, 3, 0, 3),
                UseVisualStyleBackColor = false,
                TabStop = false
            };
            b.FlatAppearance.BorderColor = NxTheme.NavyLine;
            b.FlatAppearance.MouseOverBackColor = NxTheme.NavyLine;
            _tip.SetToolTip(b, tip);
            return b;
        }

        private void OnValueChanged()
        {
            if (_loading) return;
            RefreshState();
        }

        /// <summary>Cập nhật xem trước + trạng thái lỗi.</summary>
        private void RefreshState()
        {
            string? error = null;
            foreach (var f in _fields)
            {
                var e = f.Check?.Invoke();
                if (f.Host != null) f.Host.HasError = e != null;
                error ??= e;
            }
            if (error == null && _validator != null)
            {
                try { error = _validator(this); }
                catch (System.Exception ex) { error = ex.Message; }
            }

            _lblError.Text = error != null ? "✖  " + error : "";
            _lblError.Visible = error != null;
            _btnRun.Enabled = error == null;
            _btnRun.BackColor = error == null ? NxTheme.Navy : NxTheme.NavyLine;

            if (_preview != null && _lblPreview != null)
            {
                try { _lblPreview.Text = error == null ? _preview(this) : "—"; }
                catch (System.Exception ex) { _lblPreview.Text = "Lỗi xem trước: " + ex.Message; }
            }
        }

        private void TryAccept()
        {
            RefreshState();
            if (!_btnRun.Enabled) { System.Media.SystemSounds.Beep.Play(); return; }
            if (RememberValues) SaveValues();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void ResetDefaults()
        {
            _loading = true;
            foreach (var f in _fields) f.Write(f.DefaultValue);
            _loading = false;
            RefreshState();
        }

        // ─── Lưu / nạp cấu hình ───
        private string SettingsPath
        {
            get
            {
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CivilTool", "Dialogs");
                var safe = string.Concat(_commandName.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
                return Path.Combine(dir, safe + ".json");
            }
        }

        private void SaveValues()
        {
            try
            {
                var data = _fields.ToDictionary(f => f.Key, f => f.Read());
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
                File.WriteAllText(SettingsPath, JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch { /* không chặn lệnh nếu không ghi được cấu hình */ }
        }

        private void LoadSaved()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return;
                var data = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(SettingsPath));
                if (data == null) return;
                foreach (var f in _fields)
                    if (data.TryGetValue(f.Key, out var v) && v != null) f.Write(v);
            }
            catch { /* file cấu hình hỏng → dùng mặc định */ }
        }

        private static string FormatPoint(Point3d p) =>
            $"{p.X.ToString("F3", Inv)}, {p.Y.ToString("F3", Inv)}, {p.Z.ToString("F3", Inv)}";

        private static bool TryParsePoint(string s, out Point3d p)
        {
            p = Point3d.Origin;
            var parts = s.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
            if (parts.Length < 2 || parts.Length > 3) return false;
            var v = new double[3];
            for (int i = 0; i < parts.Length; i++)
                if (!double.TryParse(parts[i], NumberStyles.Float, Inv, out v[i])) return false;
            p = new Point3d(v[0], v[1], v[2]);
            return true;
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  CONTROL PHỤ TRỢ
    // ══════════════════════════════════════════════════════════════════════

    /// <summary>Thẻ (card) bo góc có tiêu đề + vạch nhấn màu, chứa lưới 3 cột: Nhãn | Ô nhập | Nút/Đơn vị.</summary>
    public class NxCard : Panel
    {
        public TableLayoutPanel Grid { get; }
        private readonly Color _accent;

        public NxCard(string title, string? description, Color? accent = null)
        {
            _accent = accent ?? NxTheme.Navy;
            Dock = DockStyle.Top;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            BackColor = NxTheme.Surface;
            Padding = new Padding(16, 38 + (description != null ? 18 : 0), 16, 12);
            DoubleBuffered = true;

            Grid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 3,
                BackColor = NxTheme.Surface,
                Margin = new Padding(0)
            };
            Grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            Grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60));
            Grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            Controls.Add(Grid);

            Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                var r = new Rectangle(0, 0, Width - 1, Height - 1);
                using (var pen = new Pen(NxTheme.Border))
                using (var path = NxTheme.RoundRect(r, 6))
                    g.DrawPath(pen, path);
                using (var b = new SolidBrush(_accent))
                    g.FillRectangle(b, 16, 14, 3, 16);
                TextRenderer.DrawText(g, title.ToUpperInvariant(), NxTheme.Section, new Point(26, 12), NxTheme.Navy, TextFormatFlags.NoPadding);
                if (description != null)
                    TextRenderer.DrawText(g, description, NxTheme.Small, new Rectangle(26, 34, Width - 40, 18), NxTheme.TextMuted,
                        TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis);
            };
            Resize += (s, e) => Invalidate();
        }
    }

    /// <summary>Ô nhập văn bản có viền tự vẽ (đổi màu khi focus / lỗi).</summary>
    public class NxInputBox : Panel
    {
        public TextBox Box { get; }
        private bool _focused;
        private bool _error;

        public bool HasError
        {
            get => _error;
            set { if (_error != value) { _error = value; BackColor = value ? NxTheme.DangerSoft : NxTheme.Surface; Box.BackColor = BackColor; Invalidate(); } }
        }

        public NxInputBox()
        {
            Height = 30;
            MinimumSize = new Size(60, 30);
            Margin = new Padding(0, 3, 0, 3);
            Padding = new Padding(8, 6, 8, 4);
            BackColor = NxTheme.Surface;
            Cursor = Cursors.IBeam;
            DoubleBuffered = true;

            Box = new TextBox
            {
                BorderStyle = BorderStyle.None,
                Font = NxTheme.Body,
                ForeColor = NxTheme.Text,
                BackColor = NxTheme.Surface,
                Dock = DockStyle.Fill
            };
            Box.GotFocus += (s, e) => { _focused = true; Invalidate(); };
            Box.LostFocus += (s, e) => { _focused = false; Invalidate(); };
            Controls.Add(Box);
            Click += (s, e) => Box.Focus();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var c = _error ? NxTheme.Danger : _focused ? NxTheme.BorderFocus : NxTheme.Border;
            using var pen = new Pen(c, _focused || _error ? 2f : 1f);
            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            if (_focused || _error) r = new Rectangle(1, 1, Width - 2, Height - 2);
            e.Graphics.DrawRectangle(pen, r);
        }
    }
}
