// 1000.Translate.cs — Dịch Text / MText trong AutoCAD qua API miễn phí
// Tác giả: HUỲNH NGỌC NAM | Kỹ sư Giao thông
//
// ╔══════════════════════════════════════════════════════════════════════════╗
// ║         CIVIL TOOL  —  DỊCH TEXT TRONG BẢN VẼ AUTOCAD                  ║
// ╠════════════════╦═════════════════════════════════════════════════════════╣
// ║  LỆNH          ║  MÔ TẢ                                                  ║
// ╠════════════════╬═════════════════════════════════════════════════════════╣
// ║  DICH          ║  Chọn Text / MText → dịch sang ngôn ngữ tùy chọn      ║
// ║                ║  • Mặc định: Tiếng Anh → Tiếng Việt                    ║
// ║                ║  • Hỗ trợ: EN↔VI, EN↔ZH, EN↔JA, EN↔FR, EN↔KO...     ║
// ║                ║  • API: MyMemory (miễn phí, không cần đăng ký)         ║
// ║                ║  • Giữ nguyên text gốc nếu API lỗi / mất mạng         ║
// ╚════════════════╩═════════════════════════════════════════════════════════╝

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Net.Http;
using System.Text.Json;
using System.Windows.Forms;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Font = System.Drawing.Font;
using FlowDirection = System.Windows.Forms.FlowDirection;

namespace MyFirstProject
{
    public class TranslateTextCommand
    {
        // HttpClient tĩnh — tái dùng cho toàn bộ vòng đời plugin (best practice)
        private static readonly HttpClient _http = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(20)
        };

        // ════════════════════════════════════════════════════════════════════════
        // LỆNH: DICH
        // Dịch Text / MText được chọn trong bản vẽ sang ngôn ngữ đích
        // ════════════════════════════════════════════════════════════════════════
        [CommandMethod("DICH")]
        public void DichTextCommand()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db  = doc.Database;
            Editor   ed  = doc.Editor;

            ed.WriteMessage("\n╔══════════════════════════════════════════════╗");
            ed.WriteMessage("\n║   DICH — DỊCH TEXT TRONG BẢN VẼ AUTOCAD     ║");
            ed.WriteMessage("\n╚══════════════════════════════════════════════╝");

            // ── BƯỚC 1: Chọn ngôn ngữ ─────────────────────────────────────────
            var (fromCode, fromName, toCode, toName) = ShowLanguageDialog();
            if (fromCode == null)
            {
                ed.WriteMessage("\nĐã hủy lệnh DICH.");
                return;
            }
            string langPair = $"{fromCode}|{toCode}";
            ed.WriteMessage($"\n🌐 Dịch: {fromName} → {toName}");

            // ── BƯỚC 2: Chọn Text / MText ─────────────────────────────────────
            SelectionFilter filter = new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Operator, "<OR"),
                new TypedValue((int)DxfCode.Start,    "TEXT"),
                new TypedValue((int)DxfCode.Start,    "MTEXT"),
                new TypedValue((int)DxfCode.Operator, "OR>")
            });

            PromptSelectionOptions pso = new PromptSelectionOptions
            {
                MessageForAdding = "\nChọn Text / MText cần dịch (Enter = toàn bộ bản vẽ): ",
                RejectObjectsOnLockedLayers = false
            };

            ObjectId[] ids;
            PromptSelectionResult psr = ed.GetSelection(pso, filter);

            if (psr.Status == PromptStatus.OK)
            {
                ids = psr.Value.GetObjectIds();
                ed.WriteMessage($"\n→ Đã chọn {ids.Length} đối tượng.");
            }
            else
            {
                // Enter không chọn → lấy toàn bộ bản vẽ
                PromptSelectionResult allPsr = ed.SelectAll(filter);
                if (allPsr.Status != PromptStatus.OK || allPsr.Value.Count == 0)
                {
                    ed.WriteMessage("\n⚠ Không có Text / MText nào trong bản vẽ!");
                    return;
                }
                ids = allPsr.Value.GetObjectIds();
                ed.WriteMessage($"\n→ Quét toàn bộ bản vẽ: {ids.Length} đối tượng.");
            }

            // ── BƯỚC 3: Thu thập nội dung text ───────────────────────────────
            var items = new List<(ObjectId Id, string Text, bool IsMText)>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in ids)
                {
                    try
                    {
                        DBObject ent = tr.GetObject(id, OpenMode.ForRead);
                        if (ent is DBText dbt && !string.IsNullOrWhiteSpace(dbt.TextString))
                            items.Add((id, dbt.TextString.Trim(), false));
                        else if (ent is MText mt && !string.IsNullOrWhiteSpace(mt.Text))
                            items.Add((id, mt.Text.Trim(), true));
                    }
                    catch { /* bỏ qua entity bị lỗi */ }
                }
                tr.Commit();
            }

            if (items.Count == 0)
            {
                ed.WriteMessage("\n⚠ Không có nội dung text hợp lệ để dịch.");
                return;
            }

            ed.WriteMessage($"\n⏳ Đang dịch {items.Count} đoạn text...");

            // ── BƯỚC 4: Gọi API dịch (đồng bộ để tránh deadlock trong CAD) ───
            var translated   = new List<string>(items.Count);
            int successCount = 0;
            int failCount    = 0;

            using (var progressDlg = new TranslateProgressForm(items.Count))
            {
                progressDlg.Show();
                System.Windows.Forms.Application.DoEvents();

                for (int i = 0; i < items.Count; i++)
                {
                    progressDlg.SetProgress(i + 1, items[i].Text);
                    System.Windows.Forms.Application.DoEvents();

                    string? result = CallTranslateApi(items[i].Text, langPair);

                    if (result != null)
                    {
                        translated.Add(result);
                        successCount++;
                    }
                    else
                    {
                        translated.Add(items[i].Text); // giữ nguyên khi lỗi
                        failCount++;
                    }
                }

                progressDlg.Close();
            }

            // ── BƯỚC 5: Ghi kết quả dịch vào bản vẽ ─────────────────────────
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                for (int i = 0; i < items.Count; i++)
                {
                    try
                    {
                        DBObject ent = tr.GetObject(items[i].Id, OpenMode.ForWrite);
                        if (ent is DBText dbt)
                            dbt.TextString = translated[i];
                        else if (ent is MText mt)
                            mt.Contents = translated[i];
                    }
                    catch { /* bỏ qua nếu entity đã bị xóa */ }
                }
                tr.Commit();
            }

            ed.WriteMessage($"\n✔ Dịch xong! Thành công: {successCount} / {items.Count}");
            if (failCount > 0)
                ed.WriteMessage($"\n⚠ {failCount} đoạn text bị lỗi → giữ nguyên nội dung gốc.");
            ed.WriteMessage($"\n   API: MyMemory | Cặp ngôn ngữ: {fromName} → {toName}");
        }

        // ════════════════════════════════════════════════════════════════════════
        // Gọi MyMemory Translation API (miễn phí, không cần API Key)
        // Giới hạn: ~5000 ký tự / ngày (không đăng ký)
        //           ~50000 ký tự / ngày (có email)
        // Endpoint: https://api.mymemory.translated.net/get?q=TEXT&langpair=en|vi
        // ════════════════════════════════════════════════════════════════════════
        private static string? CallTranslateApi(string text, string langPair)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;

            // Cắt nếu quá dài (giới hạn MyMemory ~500 ký tự/request)
            if (text.Length > 500) text = text.Substring(0, 500);

            try
            {
                string encoded = Uri.EscapeDataString(text);
                string url     = $"https://api.mymemory.translated.net/get?q={encoded}&langpair={langPair}";

                HttpResponseMessage resp = _http.GetAsync(url).GetAwaiter().GetResult();
                if (!resp.IsSuccessStatusCode) return null;

                string json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();

                using JsonDocument jDoc = JsonDocument.Parse(json);
                JsonElement root = jDoc.RootElement;

                // Kiểm tra responseStatus == 200
                if (root.TryGetProperty("responseStatus", out JsonElement status) &&
                    status.GetInt32() == 200 &&
                    root.TryGetProperty("responseData", out JsonElement data) &&
                    data.TryGetProperty("translatedText", out JsonElement translatedEl))
                {
                    string? result = translatedEl.GetString();
                    // Kiểm tra nếu API trả về cảnh báo QUERY LENGTH (nghĩa là quá hạn ngày)
                    if (result != null &&
                        result.StartsWith("PLEASE SELECT", StringComparison.OrdinalIgnoreCase))
                        return null;
                    return result;
                }
            }
            catch
            {
                // Mất mạng hoặc timeout → trả null
            }

            return null;
        }

        // ════════════════════════════════════════════════════════════════════════
        // Dialog chọn cặp ngôn ngữ
        // Trả về (fromCode, fromName, toCode, toName) hoặc (null,...) nếu hủy
        // ════════════════════════════════════════════════════════════════════════
        private (string? From, string FromName, string? To, string ToName)
            ShowLanguageDialog()
        {
            using Form frm = new Form();
            frm.Text            = "Cài đặt dịch thuật — DICH";
            frm.Size            = new Size(420, 340);
            frm.StartPosition   = FormStartPosition.CenterScreen;
            frm.FormBorderStyle = FormBorderStyle.FixedDialog;
            frm.MaximizeBox     = false;
            frm.MinimizeBox     = false;
            frm.Font            = new Font("Segoe UI", 9.5f);
            frm.BackColor       = Color.FromArgb(245, 247, 250);

            // ── Tiêu đề ───────────────────────────────────────────────────────
            var pnlTop = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 48,
                BackColor = Color.FromArgb(30, 100, 190)
            };
            var lblTitle = new Label
            {
                Text      = "🌐  Dịch Text trong AutoCAD",
                ForeColor = Color.White,
                Font      = new Font("Segoe UI Semibold", 12f),
                Location  = new Point(12, 10),
                Size      = new Size(380, 30)
            };
            pnlTop.Controls.Add(lblTitle);

            // ── Danh sách ngôn ngữ ─────────────────────────────────────────────
            var langs = new (string Code, string Display)[]
            {
                ("en",    "🇺🇸 Tiếng Anh  (EN)"),
                ("vi",    "🇻🇳 Tiếng Việt (VI)"),
                ("zh-CN", "🇨🇳 Tiếng Trung Giản (ZH-CN)"),
                ("zh-TW", "🇹🇼 Tiếng Trung Phồn (ZH-TW)"),
                ("ja",    "🇯🇵 Tiếng Nhật  (JA)"),
                ("ko",    "🇰🇷 Tiếng Hàn   (KO)"),
                ("fr",    "🇫🇷 Tiếng Pháp  (FR)"),
                ("de",    "🇩🇪 Tiếng Đức   (DE)"),
                ("ru",    "🇷🇺 Tiếng Nga   (RU)"),
                ("th",    "🇹🇭 Tiếng Thái  (TH)"),
            };

            // ComboBox Nguồn
            var lblFrom = new Label { Text = "Ngôn ngữ nguồn:", Location = new Point(16, 68), Size = new Size(140, 24) };
            var cboFrom = new ComboBox
            {
                Location      = new Point(164, 65),
                Size          = new Size(230, 28),
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            // ComboBox Đích
            var lblTo = new Label { Text = "Ngôn ngữ đích:", Location = new Point(16, 104), Size = new Size(140, 24) };
            var cboTo = new ComboBox
            {
                Location      = new Point(164, 101),
                Size          = new Size(230, 28),
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            foreach (var (code, display) in langs)
            {
                cboFrom.Items.Add(new LangEntry(code, display));
                cboTo.Items.Add(new LangEntry(code, display));
            }
            cboFrom.SelectedIndex = 0; // mặc định EN
            cboTo.SelectedIndex   = 1; // mặc định VI

            // ── Nút preset nhanh ──────────────────────────────────────────────
            var lblPreset = new Label
            {
                Text     = "Chọn nhanh:",
                Location = new Point(16, 144),
                Size     = new Size(140, 24)
            };
            var pnlPreset = new FlowLayoutPanel
            {
                Location      = new Point(16, 168),
                Size          = new Size(378, 76),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents  = true,
            };

            (string Label, int F, int T)[] presets =
            {
                ("EN → VI", 0, 1), ("VI → EN", 1, 0),
                ("EN → ZH", 0, 2), ("EN → JA", 0, 4),
                ("EN → KO", 0, 5), ("EN → FR", 0, 6),
                ("EN → DE", 0, 7), ("EN → RU", 0, 8),
            };

            foreach (var (label, f, t) in presets)
            {
                int fi = f, ti = t;
                var btn = new Button
                {
                    Text      = label,
                    Size      = new Size(84, 30),
                    Margin    = new Padding(0, 0, 6, 6),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(210, 230, 255),
                    ForeColor = Color.FromArgb(20, 70, 150),
                    Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold)
                };
                btn.Click += (s, _) => { cboFrom.SelectedIndex = fi; cboTo.SelectedIndex = ti; };
                pnlPreset.Controls.Add(btn);
            }

            // ── Ghi chú API ───────────────────────────────────────────────────
            var lblNote = new Label
            {
                Text      = "ℹ API: MyMemory (miễn phí) — giới hạn ~5000 ký tự/ngày",
                Location  = new Point(16, 252),
                Size      = new Size(380, 22),
                ForeColor = Color.FromArgb(100, 100, 100),
                Font      = new Font("Segoe UI", 8f)
            };

            // ── Nút OK / Hủy ──────────────────────────────────────────────────
            var btnOk = new Button
            {
                Text         = "✔  Bắt đầu dịch",
                Location     = new Point(16, 278),
                Size         = new Size(180, 38),
                BackColor    = Color.FromArgb(30, 120, 200),
                ForeColor    = Color.White,
                FlatStyle    = FlatStyle.Flat,
                Font         = new Font("Segoe UI Semibold", 9.5f),
                DialogResult = DialogResult.OK
            };
            var btnCancel = new Button
            {
                Text         = "✖  Hủy",
                Location     = new Point(210, 278),
                Size         = new Size(180, 38),
                BackColor    = Color.FromArgb(210, 70, 60),
                ForeColor    = Color.White,
                FlatStyle    = FlatStyle.Flat,
                Font         = new Font("Segoe UI Semibold", 9.5f),
                DialogResult = DialogResult.Cancel
            };

            frm.AcceptButton = btnOk;
            frm.CancelButton = btnCancel;
            frm.Controls.AddRange(new Control[]
            {
                pnlTop,
                lblFrom, cboFrom,
                lblTo,   cboTo,
                lblPreset, pnlPreset,
                lblNote,
                btnOk, btnCancel
            });

            if (frm.ShowDialog() != DialogResult.OK)
                return (null, "", null, "");

            var from = (LangEntry)cboFrom.SelectedItem!;
            var to   = (LangEntry)cboTo.SelectedItem!;

            if (from.Code == to.Code)
            {
                MessageBox.Show(
                    "Ngôn ngữ nguồn và đích không được giống nhau!",
                    "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return (null, "", null, "");
            }

            return (from.Code, from.ShortName, to.Code, to.ShortName);
        }

        // ── Helper: mục ComboBox ─────────────────────────────────────────────
        private class LangEntry
        {
            public string Code      { get; }
            public string ShortName { get; }
            private readonly string _display;
            public LangEntry(string code, string display)
            {
                Code      = code;
                _display  = display;
                // Tách tên ngắn từ display: "🇺🇸 Tiếng Anh  (EN)" → "Tiếng Anh"
                ShortName = display.Length > 4 ? display.Substring(3).Trim().Split('(')[0].Trim() : display;
            }
            public override string ToString() => _display;
        }
    }

    // ════════════════════════════════════════════════════════════════════════════
    // Form hiển thị tiến trình dịch
    // ════════════════════════════════════════════════════════════════════════════
    internal class TranslateProgressForm : Form
    {
        private readonly Label       _lblStatus;
        private readonly Label       _lblPreview;
        private readonly ProgressBar _bar;
        private readonly int         _total;

        public TranslateProgressForm(int total)
        {
            _total = total;

            Text            = "Đang dịch...";
            Size            = new Size(440, 160);
            StartPosition   = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            MinimizeBox     = false;
            ControlBox      = false;
            BackColor       = Color.FromArgb(245, 247, 250);
            Font            = new Font("Segoe UI", 9.5f);

            _lblStatus = new Label
            {
                Text      = $"🌐 Đang gọi API dịch... (0 / {total})",
                Location  = new Point(14, 14),
                Size      = new Size(405, 24),
                Font      = new Font("Segoe UI Semibold", 9.5f),
                ForeColor = Color.FromArgb(30, 80, 160)
            };

            _bar = new ProgressBar
            {
                Location = new Point(14, 46),
                Size     = new Size(405, 22),
                Minimum  = 0,
                Maximum  = total,
                Value    = 0,
                Style    = ProgressBarStyle.Continuous
            };

            _lblPreview = new Label
            {
                Text      = "",
                Location  = new Point(14, 76),
                Size      = new Size(405, 40),
                ForeColor = Color.Gray,
                Font      = new Font("Segoe UI", 8.5f)
            };

            Controls.AddRange(new Control[] { _lblStatus, _bar, _lblPreview });
        }

        public void SetProgress(int current, string currentText)
        {
            _lblStatus.Text = $"🌐 Đang dịch... ({current} / {_total})";
            _bar.Value      = Math.Min(current, _total);
            string preview  = currentText.Length > 85
                ? currentText.Substring(0, 82) + "…"
                : currentText;
            _lblPreview.Text = $"▶  {preview}";
        }
    }
}
