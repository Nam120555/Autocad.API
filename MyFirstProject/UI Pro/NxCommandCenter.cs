// ==============================================================================
// NxCommandCenter.cs — HỘP LỆNH (Command Center) kiểu "Command Palette"
// ------------------------------------------------------------------------------
//   Lệnh:  NX  hoặc  CT_HopLenh
//   • Tự động quét TẤT CẢ [CommandMethod] trong DLL → không bao giờ thiếu lệnh
//   • Tìm kiếm không dấu: gõ "doi ten coc" ra "Đổi tên cọc" (CTS_DoiTenCoc)
//   • Lọc theo nhóm, ★ Yêu thích, 🕘 Gần đây (tự lưu)
//   • Bàn phím: ↑/↓ chọn · Enter chạy · Ctrl+D yêu thích · Esc ẩn
// ==============================================================================
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
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
    /// <summary>Một lệnh trong Hộp Lệnh.</summary>
    public sealed class NxCommandItem
    {
        public string Name { get; init; } = "";
        public string Group { get; init; } = "";
        public string Description { get; init; } = "";
        internal string FoldName { get; init; } = "";
        internal string FoldAll { get; init; } = "";
    }

    public class NxCommandCenter : NxBorderlessForm
    {
        private const string NavFav = "★  Yêu thích";
        private const string NavRecent = "🕘  Gần đây";
        private const string NavTop = "🔥  Hay dùng";
        private const string NavAll = "☰  Tất cả lệnh";
        private const int MaxRecent = 15;

        private static NxCommandCenter? _instance;

        private readonly List<NxCommandItem> _all;
        private readonly CenterSettings _settings;
        private readonly TextBox _search;
        private readonly WinLabel _lblCount;
        private readonly ListBox _nav;
        private readonly ListBox _list;
        private readonly Panel _detail;
        private readonly Button _btnFav;
        private readonly List<(string Key, int Count)> _navItems = new();
        private int _hoverIndex = -1;

        // ══════════════ Mở hộp lệnh ══════════════
        public static void ShowCenter()
        {
            if (_instance == null || _instance.IsDisposed) _instance = new NxCommandCenter();
            _instance._search.Text = "";
            _instance.SelectNav(NavAll);
            if (!_instance.Visible)
            {
                try { AcadApp.ShowModelessDialog(_instance); }
                catch (System.Exception) { _instance.Show(); }
            }
            _instance.Activate();
            _instance._search.Focus();
        }

        private NxCommandCenter()
        {
            Text = "Hộp lệnh Civil Tool";
            Size = new Size(880, 580);
            MinimumSize = new Size(640, 420);
            TopMost = false;

            _all = DiscoverCommands();
            _settings = CenterSettings.Load();

            // ═══ HEADER ═══
            var header = new Panel { Dock = DockStyle.Top, Height = 60 };
            header.Paint += (s, e) =>
            {
                var g = e.Graphics;
                using (var grad = new LinearGradientBrush(header.ClientRectangle, NxTheme.Navy, NxTheme.NavyDeep, 0f))
                    g.FillRectangle(grad, header.ClientRectangle);
                NxTheme.DrawBadge(g, new Rectangle(16, 12, 36, 36), "⌘", NxTheme.NavyMid, Color.White);
                TextRenderer.DrawText(g, "HỘP LỆNH CIVIL TOOL", NxTheme.Title, new Point(64, 9), Color.White, TextFormatFlags.NoPadding);
                TextRenderer.DrawText(g, $"{_all.Count} lệnh  ·  Gõ để tìm, không cần dấu  ·  Enter để chạy",
                    NxTheme.Small, new Point(65, 36), NxTheme.TextOnNavyMuted, TextFormatFlags.NoPadding);
            };
            header.Resize += (s, e) => header.Invalidate();
            NxTheme.MakeDraggable(header, this);
            var btnClose = new WinLabel
            {
                Text = "✕",
                Font = new Font("Segoe UI", 11f),
                ForeColor = NxTheme.TextOnNavyMuted,
                BackColor = Color.Transparent,
                Size = new Size(36, 30),
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand
            };
            btnClose.MouseEnter += (s, e) => { btnClose.BackColor = NxTheme.Danger; btnClose.ForeColor = Color.White; };
            btnClose.MouseLeave += (s, e) => { btnClose.BackColor = Color.Transparent; btnClose.ForeColor = NxTheme.TextOnNavyMuted; };
            btnClose.Click += (s, e) => Hide();
            header.Controls.Add(btnClose);
            header.Resize += (s, e) => btnClose.Location = new Point(header.Width - btnClose.Width - 2, 2);

            // ═══ THANH TÌM KIẾM ═══
            var searchBar = new Panel { Dock = DockStyle.Top, Height = 62, BackColor = NxTheme.Canvas, Padding = new Padding(16, 12, 16, 10) };
            var searchHost = new Panel { Dock = DockStyle.Fill, BackColor = NxTheme.Surface, Padding = new Padding(40, 8, 110, 6) };
            searchHost.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using var pen = new Pen(NxTheme.BorderFocus, 1.5f);
                using var path = NxTheme.RoundRect(new Rectangle(0, 0, searchHost.Width - 1, searchHost.Height - 1), 6);
                e.Graphics.DrawPath(pen, path);
                TextRenderer.DrawText(e.Graphics, "🔍", new Font("Segoe UI Symbol", 12f), new Point(12, 9), NxTheme.TextMuted, TextFormatFlags.NoPadding);
            };
            _search = new TextBox { BorderStyle = BorderStyle.None, Font = NxTheme.Search, Dock = DockStyle.Fill, ForeColor = NxTheme.Text };
            _search.HandleCreated += (s, e) => SetCue(_search, "Tìm lệnh: tên lệnh hoặc chức năng (vd: doi ten coc, khoi luong, layer…)");
            _search.TextChanged += (s, e) => { if (_search.Text.Length > 0 && CurrentNav != NavAll) SelectNav(NavAll, false); ApplyFilter(); };
            _search.KeyDown += Search_KeyDown;
            _lblCount = new WinLabel
            {
                AutoSize = false,
                Width = 100,
                Dock = DockStyle.Right,
                TextAlign = ContentAlignment.MiddleRight,
                Font = NxTheme.Small,
                ForeColor = NxTheme.TextMuted
            };
            searchHost.Controls.Add(_search);
            searchHost.Resize += (s, e) => _lblCount.Location = new Point(searchHost.Width - 106, 10);
            _lblCount.Dock = DockStyle.None;
            _lblCount.Height = 20;
            searchHost.Controls.Add(_lblCount);
            searchBar.Controls.Add(searchHost);

            // ═══ CỘT NHÓM (trái) ═══
            _nav = new ListBox
            {
                Dock = DockStyle.Left,
                Width = 220,
                BorderStyle = BorderStyle.None,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 32,
                BackColor = NxTheme.Canvas,
                IntegralHeight = false
            };
            _nav.DrawItem += Nav_DrawItem;
            _nav.SelectedIndexChanged += (s, e) => { if (_nav.SelectedIndex >= 0 && !_suppressNav) ApplyFilter(); };
            var navSplit = new Panel { Dock = DockStyle.Left, Width = 1, BackColor = NxTheme.Border };

            // ═══ DANH SÁCH LỆNH ═══
            _list = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 50,
                BackColor = NxTheme.Surface,
                IntegralHeight = false
            };
            _list.DrawItem += List_DrawItem;
            _list.SelectedIndexChanged += (s, e) => { UpdateDetail(); _list.Invalidate(); };
            _list.DoubleClick += (s, e) => RunSelected();
            _list.MouseMove += (s, e) =>
            {
                int i = _list.IndexFromPoint(e.Location);
                if (i != _hoverIndex) { _hoverIndex = i; _list.Invalidate(); }
            };
            _list.MouseLeave += (s, e) => { _hoverIndex = -1; _list.Invalidate(); };
            _list.MouseDown += (s, e) =>
            {
                int i = _list.IndexFromPoint(e.Location);
                if (i >= 0 && e.X < 34 && _list.Items[i] is NxCommandItem it) { ToggleFavorite(it); }
            };
            _list.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { RunSelected(); e.Handled = true; } };

            // ═══ THANH CHI TIẾT (dưới) ═══
            _detail = new Panel { Dock = DockStyle.Bottom, Height = 66, BackColor = NxTheme.Canvas };
            _detail.Paint += Detail_Paint;
            _btnFav = NxTheme.CreateButton("☆ Yêu thích", false, 112);
            _btnFav.Click += (s, e) => { if (Selected is NxCommandItem it) ToggleFavorite(it); };
            var btnCopy = NxTheme.CreateButton("⧉ Sao chép", false, 104);
            btnCopy.Click += (s, e) => { if (Selected is NxCommandItem it) { try { Clipboard.SetText(it.Name); } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); } } };
            var btnRun = NxTheme.CreateButton("▶  Chạy lệnh", true, 124);
            btnRun.Click += (s, e) => RunSelected();
            _detail.Controls.AddRange(new Control[] { _btnFav, btnCopy, btnRun });
            _detail.Resize += (s, e) =>
            {
                btnRun.Location = new Point(_detail.Width - btnRun.Width - 16, 16);
                btnCopy.Location = new Point(btnRun.Left - btnCopy.Width - 8, 16);
                _btnFav.Location = new Point(btnCopy.Left - _btnFav.Width - 8, 16);
                _detail.Invalidate();
            };

            Controls.Add(header);
            Controls.Add(searchBar);
            Controls.Add(_detail);
            Controls.Add(_nav);
            Controls.Add(navSplit);
            Controls.Add(_list);
            // thứ tự dock: header → search → detail → nav → vạch → list (fill)
            _list.BringToFront();
            navSplit.SendToBack(); _nav.SendToBack(); _detail.SendToBack(); searchBar.SendToBack(); header.SendToBack();

            BuildNav();
            SelectNav(NavAll);
        }

        // ══════════════ QUÉT LỆNH ══════════════
        private static List<NxCommandItem> DiscoverCommands()
        {
            var names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Type?[] types;
            var asm = Assembly.GetExecutingAssembly();
            try { types = asm.GetTypes(); }
            catch (ReflectionTypeLoadException ex) { types = ex.Types; }

            foreach (var t in types)
            {
                if (t == null) continue;
                MethodInfo[] methods;
                try { methods = t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance); }
                catch { continue; }
                foreach (var m in methods)
                {
                    try
                    {
                        foreach (var a in m.GetCustomAttributes(typeof(CommandMethodAttribute), false).OfType<CommandMethodAttribute>())
                            if (!string.IsNullOrWhiteSpace(a.GlobalName)) names[a.GlobalName] = a.GlobalName;
                    }
                    catch { /* bỏ qua method không đọc được attribute */ }
                }
            }

            var list = new List<NxCommandItem>();
            foreach (var name in names.Values)
            {
                if (MenuSettingsForm.IsRetiredCommand(name)) continue;   // ẩn lệnh vỏ rỗng NXsoft
                string group, desc;
                if (NxCommandCatalog.Entries.TryGetValue(name.ToUpperInvariant(), out var info))
                {
                    group = info.Group;
                    desc = info.Description;
                }
                else
                {
                    group = GuessGroup(name);
                    desc = Humanize(name);
                }
                if (string.IsNullOrWhiteSpace(desc)) desc = Humanize(name);
                list.Add(new NxCommandItem
                {
                    Name = name,
                    Group = group,
                    Description = desc,
                    FoldName = NxTheme.Fold(name),
                    FoldAll = NxTheme.Fold(name.Replace('_', ' ') + " " + Humanize(name) + " " + desc + " " + group)
                });
            }
            return list.OrderBy(i => i.Group).ThenBy(i => i.Name, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static string GuessGroup(string name)
        {
            string u = name.ToUpperInvariant();
            (string P, string G)[] rules =
            {
                ("CTSV_", "Trắc ngang (Section View)"), ("CTS_", "Cọc (Sample Line)"), ("CTPO_", "CogoPoint"),
                ("CTPV_", "Trắc dọc View"), ("CTPS_", "Cống & Hố thu"), ("CTPI_", "Cống & Hố thu"),
                ("CTPA_", "Parcel"), ("CTP_", "Trắc dọc (Profile)"), ("CTC_", "Corridor"), ("CTA_", "Tuyến (Alignment)"),
                ("CTL_", "Layer"), ("CTDS_", "Thiết lập bản vẽ"), ("CTU_", "Tiện ích Civil"), ("TT_", "Tính toán thiết kế"),
                ("AT_", "AutoCAD Tool"), ("CT_", "Hệ thống & Menu")
            };
            foreach (var (p, g) in rules) if (u.StartsWith(p)) return g;
            return "Khác";
        }

        /// <summary>"CTS_DoiTenCoc2" → "Doi Ten Coc 2" (dùng khi chưa có mô tả).</summary>
        private static string Humanize(string name)
        {
            var core = Regex.Replace(name, @"^[A-Za-z]{1,5}_", "");
            core = Regex.Replace(core, @"(?<=[a-z0-9])(?=[A-Z])|_", " ");
            core = Regex.Replace(core, @"(?<=[A-Za-z])(?=\d)", " ");
            return Regex.Replace(core, @"\s+", " ").Trim();
        }

        // ══════════════ LỌC & XẾP HẠNG ══════════════
        private string CurrentNav => _nav.SelectedIndex >= 0 && _nav.SelectedIndex < _navItems.Count ? _navItems[_nav.SelectedIndex].Key : NavAll;

        private void BuildNav()
        {
            _navItems.Clear();
            _navItems.Add((NavFav, _settings.Favorites.Count));
            _navItems.Add((NavRecent, _settings.Recent.Count));
            _navItems.Add((NavTop, MyFirstProject.Core.NxUsage.UsedCount));
            _navItems.Add((NavAll, _all.Count));
            foreach (var g in _all.GroupBy(i => i.Group).OrderBy(g => g.Key))
                _navItems.Add((g.Key, g.Count()));
            int sel = _nav.SelectedIndex;
            _nav.BeginUpdate();
            _nav.Items.Clear();
            foreach (var n in _navItems) _nav.Items.Add(n.Key);
            _nav.EndUpdate();
            if (sel >= 0 && sel < _nav.Items.Count) { _suppressNav = true; _nav.SelectedIndex = sel; _suppressNav = false; }
        }

        private bool _suppressNav;

        private void SelectNav(string key, bool apply = true)
        {
            int idx = _navItems.FindIndex(n => n.Key == key);
            if (idx < 0) return;
            _suppressNav = true;
            _nav.SelectedIndex = idx;
            _suppressNav = false;
            if (apply) ApplyFilter();
        }

        private void ApplyFilter()
        {
            string nav = CurrentNav;
            var tokens = NxTheme.Fold(_search.Text).Split(new[] { ' ', '_' }, StringSplitOptions.RemoveEmptyEntries);
            var fav = new HashSet<string>(_settings.Favorites, StringComparer.OrdinalIgnoreCase);

            IEnumerable<NxCommandItem> src = nav switch
            {
                NavFav => _all.Where(i => fav.Contains(i.Name)),
                NavRecent => _settings.Recent.Select(r => _all.FirstOrDefault(i => i.Name.Equals(r, StringComparison.OrdinalIgnoreCase)))
                                             .Where(i => i != null)!.Cast<NxCommandItem>(),
                NavTop => MyFirstProject.Core.NxUsage.Top(30)
                              .Select(u => _all.FirstOrDefault(i => i.Name.Equals(u.Command, StringComparison.OrdinalIgnoreCase)))
                              .Where(i => i != null)!.Cast<NxCommandItem>(),
                NavAll => _all,
                _ => _all.Where(i => i.Group == nav)
            };

            List<NxCommandItem> result;
            if (tokens.Length == 0)
            {
                result = src.ToList();
            }
            else
            {
                result = src
                    .Select(i => (Item: i, Score: Score(i, tokens, fav)))
                    .Where(x => x.Score > 0)
                    .OrderByDescending(x => x.Score)
                    .ThenByDescending(x => MyFirstProject.Core.NxUsage.CountOf(x.Item.Name))   // lệnh hay dùng lên trước
                    .ThenBy(x => x.Item.Name.Length)
                    .Select(x => x.Item)
                    .ToList();
            }

            _list.BeginUpdate();
            _list.Items.Clear();
            foreach (var r in result) _list.Items.Add(r);
            _list.EndUpdate();
            if (_list.Items.Count > 0) _list.SelectedIndex = 0;
            _lblCount.Text = $"{_list.Items.Count} kết quả";
            UpdateDetail();
            _nav.Invalidate();
        }

        private int Score(NxCommandItem i, string[] tokens, HashSet<string> fav)
        {
            int score = 0;
            foreach (var t in tokens)
            {
                if (!i.FoldAll.Contains(t)) return 0;          // tất cả từ khóa phải khớp
                if (i.FoldName.StartsWith(t) || i.FoldName.Contains("_" + t)) score += 6;
                else if (i.FoldName.Contains(t)) score += 4;
                else score += 2;
            }
            string full = string.Join(" ", tokens);
            if (i.FoldName == full.Replace(" ", "_") || i.FoldName == full) score += 20;   // gõ đúng tên lệnh
            if (fav.Contains(i.Name)) score += 3;
            int r = _settings.Recent.FindIndex(x => x.Equals(i.Name, StringComparison.OrdinalIgnoreCase));
            if (r >= 0) score += Math.Max(1, 3 - r / 5);
            return score;
        }

        private NxCommandItem? Selected => _list.SelectedItem as NxCommandItem;

        // ══════════════ HÀNH ĐỘNG ══════════════
        private void RunSelected()
        {
            var it = Selected;
            if (it == null) return;
            _settings.Recent.RemoveAll(x => x.Equals(it.Name, StringComparison.OrdinalIgnoreCase));
            _settings.Recent.Insert(0, it.Name);
            if (_settings.Recent.Count > MaxRecent) _settings.Recent.RemoveRange(MaxRecent, _settings.Recent.Count - MaxRecent);
            _settings.Save();
            BuildNav();

            Hide();
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            doc?.SendStringToExecute(it.Name + " ", true, false, true);
        }

        private void ToggleFavorite(NxCommandItem it)
        {
            int idx = _settings.Favorites.FindIndex(x => x.Equals(it.Name, StringComparison.OrdinalIgnoreCase));
            if (idx >= 0) _settings.Favorites.RemoveAt(idx); else _settings.Favorites.Add(it.Name);
            _settings.Save();
            var keep = it;
            BuildNav();
            if (CurrentNav == NavFav) ApplyFilter();
            else { _list.Invalidate(); UpdateDetail(); }
            if (_list.Items.Contains(keep)) _list.SelectedItem = keep;
        }

        private void Search_KeyDown(object? sender, KeyEventArgs e)
        {
            int n = _list.Items.Count;
            switch (e.KeyCode)
            {
                case Keys.Down when n > 0: _list.SelectedIndex = Math.Min(n - 1, _list.SelectedIndex + 1); e.Handled = true; break;
                case Keys.Up when n > 0: _list.SelectedIndex = Math.Max(0, _list.SelectedIndex - 1); e.Handled = true; break;
                case Keys.PageDown when n > 0: _list.SelectedIndex = Math.Min(n - 1, _list.SelectedIndex + 8); e.Handled = true; break;
                case Keys.PageUp when n > 0: _list.SelectedIndex = Math.Max(0, _list.SelectedIndex - 8); e.Handled = true; break;
                case Keys.Enter: RunSelected(); e.Handled = e.SuppressKeyPress = true; break;
                case Keys.Escape:
                    if (_search.Text.Length > 0) _search.Clear(); else Hide();
                    e.Handled = e.SuppressKeyPress = true; break;
                case Keys.D when e.Control:
                    if (Selected is NxCommandItem it) ToggleFavorite(it);
                    e.Handled = e.SuppressKeyPress = true; break;
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape && !_search.Focused) { Hide(); return true; }
            if (keyData == (Keys.Control | Keys.F)) { _search.Focus(); _search.SelectAll(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); return; }
            base.OnFormClosing(e);
        }

        // ══════════════ VẼ GIAO DIỆN ══════════════
        private void Nav_DrawItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _navItems.Count) return;
            var (key, count) = _navItems[e.Index];
            bool sel = e.Index == _nav.SelectedIndex;
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var bg = new SolidBrush(NxTheme.Canvas)) g.FillRectangle(bg, e.Bounds);

            var r = new Rectangle(e.Bounds.X + 8, e.Bounds.Y + 2, e.Bounds.Width - 16, e.Bounds.Height - 4);
            if (sel)
                using (var path = NxTheme.RoundRect(r, 5))
                using (var b = new SolidBrush(NxTheme.Navy))
                    g.FillPath(b, path);

            bool special = e.Index < 3;
            var font = special ? NxTheme.BodyBold : NxTheme.Body;
            var fore = sel ? Color.White : special ? NxTheme.Navy : NxTheme.Text;
            TextRenderer.DrawText(g, key, font, new Rectangle(r.X + 10, r.Y, r.Width - 50, r.Height), fore,
                TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding);
            TextRenderer.DrawText(g, count.ToString(), NxTheme.Small, new Rectangle(r.Right - 44, r.Y, 36, r.Height),
                sel ? NxTheme.TextOnNavyMuted : NxTheme.TextMuted, TextFormatFlags.VerticalCenter | TextFormatFlags.Right);

            if (e.Index == 2)
                using (var p = new Pen(NxTheme.Border))
                    g.DrawLine(p, e.Bounds.X + 14, e.Bounds.Bottom - 1, e.Bounds.Right - 14, e.Bounds.Bottom - 1);
        }

        private void List_DrawItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _list.Items.Count || _list.Items[e.Index] is not NxCommandItem it) return;
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            bool sel = (e.State & DrawItemState.Selected) != 0;
            bool hover = e.Index == _hoverIndex;
            var b = e.Bounds;

            using (var bg = new SolidBrush(sel ? NxTheme.NavySoft : hover ? Color.FromArgb(248, 250, 253) : NxTheme.Surface))
                g.FillRectangle(bg, b);
            if (sel) using (var acc = new SolidBrush(NxTheme.Navy)) g.FillRectangle(acc, b.X, b.Y, 4, b.Height);
            using (var line = new Pen(Color.FromArgb(238, 241, 246))) g.DrawLine(line, b.X + 34, b.Bottom - 1, b.Right, b.Bottom - 1);

            bool fav = _settings.Favorites.Any(x => x.Equals(it.Name, StringComparison.OrdinalIgnoreCase));
            TextRenderer.DrawText(g, fav ? "★" : "☆", new Font("Segoe UI Symbol", 12f), new Rectangle(b.X + 8, b.Y, 24, b.Height),
                fav ? NxTheme.Gold : Color.FromArgb(200, 205, 214), TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter);

            // Nhóm (chip) bên phải
            var chipSize = TextRenderer.MeasureText(it.Group, NxTheme.Small);
            int chipX = b.Right - chipSize.Width - 30;
            NxTheme.DrawChip(g, it.Group, NxTheme.Small, chipX, b.Y + b.Height / 2,
                sel ? Color.White : NxTheme.NavySoft, NxTheme.NavyMid);

            int textW = chipX - (b.X + 40) - 10;
            TextRenderer.DrawText(g, it.Name, NxTheme.MonoBold, new Rectangle(b.X + 40, b.Y + 7, textW, 18), NxTheme.Navy,
                TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis);
            TextRenderer.DrawText(g, it.Description, NxTheme.Body, new Rectangle(b.X + 40, b.Y + 26, textW, 18), NxTheme.TextMuted,
                TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis);
        }

        private void Detail_Paint(object? sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            using (var p = new Pen(NxTheme.Border)) g.DrawLine(p, 0, 0, _detail.Width, 0);
            var it = Selected;
            int w = _btnFav.Left - 32;
            if (it == null)
            {
                TextRenderer.DrawText(g, "Không tìm thấy lệnh phù hợp — thử từ khóa khác.", NxTheme.Body,
                    new Rectangle(18, 0, w, _detail.Height), NxTheme.TextMuted, TextFormatFlags.VerticalCenter);
                return;
            }
            TextRenderer.DrawText(g, it.Name, new Font("Consolas", 11f, FontStyle.Bold), new Rectangle(18, 12, w, 22), NxTheme.Navy,
                TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis);
            TextRenderer.DrawText(g, it.Group + "  ·  " + it.Description, NxTheme.Small, new Rectangle(18, 38, w, 18), NxTheme.TextMuted,
                TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis);
        }

        private void UpdateDetail()
        {
            var it = Selected;
            bool fav = it != null && _settings.Favorites.Any(x => x.Equals(it.Name, StringComparison.OrdinalIgnoreCase));
            _btnFav.Text = fav ? "★ Bỏ thích" : "☆ Yêu thích";
            _btnFav.Enabled = it != null;
            _detail.Invalidate();
        }

        // ══════════════ TIỆN ÍCH ══════════════
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

        private static void SetCue(TextBox tb, string cue)
        {
            try { SendMessage(tb.Handle, 0x1501 /*EM_SETCUEBANNER*/, (IntPtr)1, cue); } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
        }

        /// <summary>Lưu Yêu thích / Gần đây tại %AppData%\CivilTool\CommandCenter.json</summary>
        private sealed class CenterSettings
        {
            public List<string> Favorites { get; set; } = new();
            public List<string> Recent { get; set; } = new();

            private static string FilePath => Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CivilTool", "CommandCenter.json");

            public static CenterSettings Load()
            {
                try
                {
                    if (File.Exists(FilePath))
                        return JsonSerializer.Deserialize<CenterSettings>(File.ReadAllText(FilePath)) ?? new CenterSettings();
                }
                catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                return new CenterSettings();
            }

            public void Save()
            {
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
                    File.WriteAllText(FilePath, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
                }
                catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
            }
        }
    }
}
