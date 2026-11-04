using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Font = System.Drawing.Font;
using FlowDirection = System.Windows.Forms.FlowDirection;

// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║           DICH TEXT - DỊCH TIẾNG ANH TRONG BẢN VẼ AUTOCAD                  ║
// ║           Tác giả: HUỲNH NGỌC NAM  |  Kỹ sư Giao thông                     ║
// ╠══════════════╦══════════════════════════════════════════════════════════════╣
// ║  LỆNH        ║  MÔ TẢ                                                      ║
// ╠══════════════╬══════════════════════════════════════════════════════════════╣
// ║  DICH        ║  Chọn Text / MText → dịch sang ngôn ngữ tùy chọn           ║
// ║              ║  Sử dụng MyMemory API (miễn phí, không cần API Key)         ║
// ║              ║  Hỗ trợ: EN↔VI, EN↔ZH, EN↔FR, EN↔JA, ...                 ║
// ╚══════════════╩══════════════════════════════════════════════════════════════╝

namespace MyFirstProject.BrTool
{
    public class DichText
    {
        // HttpClient tĩnh — tái dùng cho toàn bộ vòng đời plugin
        private static readonly HttpClient _http = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(15)
        };

        // ════════════════════════════════════════════════════════════════════════
        // LỆNH: DICH
        // Dịch nội dung Text / MText đã chọn sang ngôn ngữ đích
        // ════════════════════════════════════════════════════════════════════════
        [CommandMethod("DICH")]
        public void DichTextCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db  = doc.Database;
            Editor   ed  = doc.Editor;

            // ── BƯỚC 1: Hiển thị dialog chọn cặp ngôn ngữ ─────────────────────
            string langPair = ShowLanguageDialog();
            if (langPair == null)
            {
                ed.WriteMessage("\nĐã hủy lệnh dịch.");
                return;
            }

            ed.WriteMessage($"\n🌐 Cặp ngôn ngữ: {langPair}");
            ed.WriteMessage("\nChọn các Text / MText cần dịch (Enter để kết thúc): ");

            // ── BƯỚC 2: Chọn Text / MText ───────────────────────────────────────
            SelectionFilter filter = new SelectionFilter(new[]
            {
                new Autodesk.AutoCAD.DatabaseServices.TypedValue(
                    (int)DxfCode.Start, "TEXT,MTEXT")
            });

            PromptSelectionResult psr = ed.GetSelection(filter);
            if (psr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nKhông có Text nào được chọn.");
                return;
            }

            // ── BƯỚC 3: Thu thập nội dung cần dịch ─────────────────────────────
            var items = new List<(ObjectId Id, string OriginalText, bool IsMText)>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject so in psr.Value)
                {
                    DBObject ent = tr.GetObject(so.ObjectId, OpenMode.ForRead);
                    if (ent is DBText dbt && !string.IsNullOrWhiteSpace(dbt.TextString))
                        items.Add((so.ObjectId, dbt.TextString.Trim(), false));
                    else if (ent is MText mt && !string.IsNullOrWhiteSpace(mt.Text))
                        items.Add((so.ObjectId, mt.Text.Trim(), true));
                }
                tr.Commit();
            }

            if (items.Count == 0)
            {
                ed.WriteMessage("\nKhông có nội dung text hợp lệ để dịch.");
                return;
            }

            ed.WriteMessage($"\n⏳ Đang dịch {items.Count} đoạn text...");

            // ── BƯỚC 4: Gọi API dịch (chạy đồng bộ, tránh deadlock trong CAD) ──
            var translated = new List<string>();
            int successCount = 0;
            int failCount    = 0;

            // Hiển thị Progress Dialog
            using (var progress = new ProgressForm(items.Count))
            {
                progress.Show();
                System.Windows.Forms.Application.DoEvents();

                for (int i = 0; i < items.Count; i++)
                {
                    progress.SetProgress(i + 1, items[i].OriginalText);
                    System.Windows.Forms.Application.DoEvents();

                    try
                    {
                        string result = TranslateSync(items[i].OriginalText, langPair);
                        translated.Add(result ?? items[i].OriginalText);
                        if (result != null) successCount++;
                        else               failCount++;
                    }
                    catch
                    {
                        translated.Add(items[i].OriginalText); // giữ nguyên nếu lỗi
                        failCount++;
                    }
                }

                progress.Close();
            }

            // ── BƯỚC 5: Ghi kết quả dịch vào bản vẽ ─────────────────────────────
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                for (int i = 0; i < items.Count; i++)
                {
                    DBObject ent = tr.GetObject(items[i].Id, OpenMode.ForWrite);
                    if (ent is DBText dbt)
                        dbt.TextString = translated[i];
                    else if (ent is MText mt)
                        mt.Contents = translated[i];
                }
                tr.Commit();
            }

            ed.WriteMessage($"\n✔ Dịch xong! Thành công: {successCount} | Lỗi: {failCount}");
            if (failCount > 0)
                ed.WriteMessage("\n⚠ Các text bị lỗi sẽ được giữ nguyên nội dung gốc.");
        }

        // ════════════════════════════════════════════════════════════════════════
        // Gọi MyMemory API đồng bộ (blocking)
        // URL: https://api.mymemory.translated.net/get?q=TEXT&langpair=en|vi
        // Miễn phí: 5000 ký tự/ngày không cần API key
        // ════════════════════════════════════════════════════════════════════════
        private static string? TranslateSync(string text, string langPair)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;

            // Giới hạn 500 ký tự/request để tránh lỗi
            if (text.Length > 500) text = text.Substring(0, 500);

            string encoded = Uri.EscapeDataString(text);
            string url     = $"https://api.mymemory.translated.net/get?q={encoded}&langpair={langPair}";

            try
            {
                HttpResponseMessage response = _http.GetAsync(url).GetAwaiter().GetResult();
                if (!response.IsSuccessStatusCode) return null;

                string json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();

                using JsonDocument doc = JsonDocument.Parse(json);
                JsonElement root = doc.RootElement;

                // Kiểm tra responseStatus == 200
                if (root.TryGetProperty("responseStatus", out JsonElement status) &&
                    status.GetInt32() == 200 &&
                    root.TryGetProperty("responseData", out JsonElement data) &&
                    data.TryGetProperty("translatedText", out JsonElement translated))
                {
                    return translated.GetString();
                }
            }
            catch { /* trả null để caller xử lý */ }

            return null;
        }

        // ════════════════════════════════════════════════════════════════════════
        // Dialog chọn cặp ngôn ngữ
        // ════════════════════════════════════════════════════════════════════════
        private string? ShowLanguageDialog()
        {
            using Form frm = new Form();
            frm.Text            = "Cài đặt dịch thuật";
            frm.Size            = new Size(400, 320);
            frm.StartPosition   = FormStartPosition.CenterScreen;
            frm.FormBorderStyle = FormBorderStyle.FixedDialog;
            frm.MaximizeBox     = false;
            frm.MinimizeBox     = false;
            frm.Font            = new Font("Segoe UI", 9.5f);
            frm.BackColor       = Color.FromArgb(245, 247, 250);

            // ── Tiêu đề ───────────────────────────────────────────────────────
            var lblTitle = new Label
            {
                Text      = "🌐  Chọn cặp ngôn ngữ dịch",
                Location  = new Point(14, 14),
                Size      = new Size(360, 28),
                Font      = new Font("Segoe UI Semibold", 11f),
                ForeColor = Color.FromArgb(30, 80, 160)
            };

            // ── Cặp ngôn ngữ phổ biến ─────────────────────────────────────────
            var lblFrom = new Label { Text = "Ngôn ngữ nguồn:", Location = new Point(14, 58), Size = new Size(140, 24) };
            var lblTo   = new Label { Text = "Ngôn ngữ đích:",  Location = new Point(14, 92), Size = new Size(140, 24) };

            var cboFrom = new ComboBox
            {
                Location      = new Point(160, 55),
                Size          = new Size(210, 28),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            var cboTo = new ComboBox
            {
                Location      = new Point(160, 89),
                Size          = new Size(210, 28),
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            // Danh sách ngôn ngữ hỗ trợ
            var langs = new (string Code, string Name)[]
            {
                ("en",    "🇺🇸 Tiếng Anh"),
                ("vi",    "🇻🇳 Tiếng Việt"),
                ("zh-CN", "🇨🇳 Tiếng Trung (Giản)"),
                ("zh-TW", "🇹🇼 Tiếng Trung (Phồn)"),
                ("ja",    "🇯🇵 Tiếng Nhật"),
                ("ko",    "🇰🇷 Tiếng Hàn"),
                ("fr",    "🇫🇷 Tiếng Pháp"),
                ("de",    "🇩🇪 Tiếng Đức"),
                ("ru",    "🇷🇺 Tiếng Nga"),
                ("th",    "🇹🇭 Tiếng Thái"),
            };

            foreach (var (code, name) in langs)
            {
                cboFrom.Items.Add(new LangItem(code, name));
                cboTo.Items.Add(new LangItem(code, name));
            }

            // Mặc định: EN → VI
            cboFrom.SelectedIndex = 0; // en
            cboTo.SelectedIndex   = 1; // vi

            // ── Preset nhanh ──────────────────────────────────────────────────
            var lblPreset = new Label { Text = "Chọn nhanh:", Location = new Point(14, 132), Size = new Size(140, 24) };
            var flPreset  = new FlowLayoutPanel
            {
                Location    = new Point(14, 156),
                Size        = new Size(358, 80),
                AutoScroll  = false,
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
            };

            var presets = new (string Label, int From, int To)[]
            {
                ("EN → VI", 0, 1),
                ("VI → EN", 1, 0),
                ("EN → ZH", 0, 2),
                ("EN → JA", 0, 4),
                ("EN → KO", 0, 5),
                ("EN → FR", 0, 6),
            };

            foreach (var (label, from, to) in presets)
            {
                int f = from, t = to;
                var btn = new Button
                {
                    Text   = label,
                    Size   = new Size(80, 30),
                    Margin = new Padding(0, 0, 6, 6),
                    Tag    = (f, t),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(220, 235, 255),
                    ForeColor = Color.FromArgb(30, 80, 160),
                    Font  = new Font("Segoe UI", 8.5f, FontStyle.Bold)
                };
                btn.Click += (s, e) =>
                {
                    var (fi, ti) = ((int, int))((Button)s).Tag;
                    cboFrom.SelectedIndex = fi;
                    cboTo.SelectedIndex   = ti;
                };
                flPreset.Controls.Add(btn);
            }

            // ── Nút xác nhận / hủy ──────────────────────────────────────────
            var btnOk = new Button
            {
                Text      = "✔  Bắt đầu dịch",
                Location  = new Point(14, 248),
                Size      = new Size(170, 38),
                BackColor = Color.FromArgb(30, 120, 200),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI Semibold", 9.5f),
                DialogResult = DialogResult.OK
            };
            var btnCancel = new Button
            {
                Text         = "✖  Hủy",
                Location     = new Point(198, 248),
                Size         = new Size(170, 38),
                BackColor    = Color.FromArgb(220, 80, 70),
                ForeColor    = Color.White,
                FlatStyle    = FlatStyle.Flat,
                Font         = new Font("Segoe UI Semibold", 9.5f),
                DialogResult = DialogResult.Cancel
            };

            frm.AcceptButton = btnOk;
            frm.CancelButton = btnCancel;
            frm.Controls.AddRange(new Control[]
            {
                lblTitle, lblFrom, lblTo, cboFrom, cboTo,
                lblPreset, flPreset, btnOk, btnCancel
            });

            if (frm.ShowDialog() != DialogResult.OK) return null;

            var fromLang = (LangItem)cboFrom.SelectedItem!;
            var toLang   = (LangItem)cboTo.SelectedItem!;

            if (fromLang.Code == toLang.Code)
            {
                MessageBox.Show("Ngôn ngữ nguồn và đích không được giống nhau!",
                    "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            return $"{fromLang.Code}|{toLang.Code}";
        }

        // Helper lớp item cho ComboBox
        private class LangItem
        {
            public string Code { get; }
            public string Name { get; }
            public LangItem(string code, string name) { Code = code; Name = name; }
            public override string ToString() => Name;
        }
    }

    // ════════════════════════════════════════════════════════════════════════════
    // Form hiển thị tiến trình dịch
    // ════════════════════════════════════════════════════════════════════════════
    internal class ProgressForm : Form
    {
        private readonly Label  _lblStatus;
        private readonly Label  _lblCurrent;
        private readonly ProgressBar _bar;
        private readonly int    _total;

        public ProgressForm(int total)
        {
            _total = total;

            Text            = "Đang dịch...";
            Size            = new Size(420, 160);
            StartPosition   = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            MinimizeBox     = false;
            ControlBox      = false;
            BackColor       = Color.FromArgb(245, 247, 250);
            Font            = new Font("Segoe UI", 9.5f);

            _lblStatus = new Label
            {
                Text     = $"🌐 Đang gọi API dịch... (0/{total})",
                Location = new Point(14, 14),
                Size     = new Size(380, 24),
                Font     = new Font("Segoe UI Semibold", 9.5f),
                ForeColor = Color.FromArgb(30, 80, 160)
            };

            _bar = new ProgressBar
            {
                Location = new Point(14, 46),
                Size     = new Size(380, 22),
                Minimum  = 0,
                Maximum  = total,
                Value    = 0,
                Style    = ProgressBarStyle.Continuous
            };

            _lblCurrent = new Label
            {
                Text      = "",
                Location  = new Point(14, 76),
                Size      = new Size(380, 40),
                ForeColor = Color.Gray,
                Font      = new Font("Segoe UI", 8.5f)
            };

            Controls.AddRange(new Control[] { _lblStatus, _bar, _lblCurrent });
        }

        public void SetProgress(int current, string currentText)
        {
            _lblStatus.Text  = $"🌐 Đang dịch... ({current}/{_total})";
            _bar.Value       = Math.Min(current, _total);
            string preview   = currentText.Length > 80 ? currentText.Substring(0, 80) + "…" : currentText;
            _lblCurrent.Text = $"▶ {preview}";
        }
    }
}
