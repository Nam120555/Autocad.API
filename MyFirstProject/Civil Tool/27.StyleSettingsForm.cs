// 27.StyleSettingsForm.cs - Form quản lý Style Civil 3D toàn bộ dự án
// Cho phép: chọn style mặc định cho Alignment, Profile, SampleLine, SectionView, Surface, Pipe, Point
// Lưu config vào JSON, cung cấp Static API cho các lệnh khác sử dụng

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;

// Aliases to resolve ambiguous references with AutoCAD/Civil types
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using ContentAlignment = System.Drawing.ContentAlignment;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using MyFirstProject.Extensions;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.StyleSettingsCommands))]

namespace MyFirstProject
{
    #region Data Model

    /// <summary>
    /// Cấu hình style đã lưu cho từng loại đối tượng
    /// </summary>
    public class StyleConfig
    {
        // Alignment
        public string AlignmentStyle { get; set; } = "";
        public string AlignmentLabelSet { get; set; } = "";

        // Profile
        public string ProfileStyle { get; set; } = "";
        public string ProfileLabelSet { get; set; } = "";

        // ProfileView
        public string ProfileViewStyle { get; set; } = "";
        public string ProfileViewBandSet { get; set; } = "";

        // SampleLine
        public string SampleLineStyle { get; set; } = "";
        public string SampleLineLabelStyle { get; set; } = "";

        // SectionView
        public string SectionViewStyle { get; set; } = "";
        public string SectionViewBandSet { get; set; } = "";

        // Surface
        public string SurfaceStyle { get; set; } = "";
        public string SurfaceLabel { get; set; } = "";

        // Pipe & Structure
        public string PipeStyle { get; set; } = "";
        public string StructureStyle { get; set; } = "";

        // Point
        public string PointStyle { get; set; } = "";
        public string PointLabelStyle { get; set; } = "";

        // Chung
        public string DefaultSite { get; set; } = "";
        public string DefaultLayer { get; set; } = "C-ROAD-ALGN";
        public DateTime LastModified { get; set; } = DateTime.Now;
    }

    #endregion

    #region Static Helper - StyleSettings

    /// <summary>
    /// Helper tĩnh để các lệnh khác lấy style đã cấu hình
    /// Sử dụng: var styleId = StyleSettings.GetAlignmentStyleId(cdoc, tr);
    /// </summary>
    public static class StyleSettings
    {
        private static readonly string ConfigFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CivilTool");
        private static readonly string ConfigPath = Path.Combine(ConfigFolder, "style_config.json");

        private static StyleConfig? _cache;

        public static StyleConfig Get()
        {
            if (_cache != null) return _cache;
            try
            {
                if (File.Exists(ConfigPath))
                {
                    string json = File.ReadAllText(ConfigPath);
                    _cache = JsonSerializer.Deserialize<StyleConfig>(json) ?? new StyleConfig();
                }
                else
                {
                    _cache = new StyleConfig();
                }
            }
            catch
            {
                _cache = new StyleConfig();
            }
            return _cache;
        }

        public static void Save(StyleConfig config)
        {
            try
            {
                if (!Directory.Exists(ConfigFolder))
                    Directory.CreateDirectory(ConfigFolder);

                config.LastModified = DateTime.Now;
                string json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigPath, json);
                _cache = config;
            }
            catch { }
        }

        public static void ClearCache() => _cache = null;

        // ═══ Helpers lấy ObjectId từ tên style đã lưu ═══

        public static ObjectId GetAlignmentStyleId(CivilDocument cdoc, Transaction tr)
        {
            return FindStyleByName(cdoc.Styles.AlignmentStyles, tr, Get().AlignmentStyle);
        }

        public static ObjectId GetAlignmentLabelSetId(CivilDocument cdoc, Transaction tr)
        {
            return FindStyleByName(cdoc.Styles.LabelSetStyles.AlignmentLabelSetStyles, tr, Get().AlignmentLabelSet);
        }

        public static ObjectId GetProfileStyleId(CivilDocument cdoc, Transaction tr)
        {
            return FindStyleByName(cdoc.Styles.ProfileStyles, tr, Get().ProfileStyle);
        }

        public static ObjectId GetProfileViewStyleId(CivilDocument cdoc, Transaction tr)
        {
            return FindStyleByName(cdoc.Styles.ProfileViewStyles, tr, Get().ProfileViewStyle);
        }

        public static ObjectId GetSurfaceStyleId(CivilDocument cdoc, Transaction tr)
        {
            return FindStyleByName(cdoc.Styles.SurfaceStyles, tr, Get().SurfaceStyle);
        }

        public static ObjectId GetPointStyleId(CivilDocument cdoc, Transaction tr)
        {
            return FindStyleByName(cdoc.Styles.PointStyles, tr, Get().PointStyle);
        }

        /// <summary>
        /// Tìm ObjectId của style theo tên trong collection
        /// Nếu không tìm thấy, trả về style đầu tiên hoặc ObjectId.Null
        /// </summary>
        private static ObjectId FindStyleByName(System.Collections.IEnumerable collection, Transaction tr, string styleName)
        {
            ObjectId firstId = ObjectId.Null;
            foreach (ObjectId id in collection)
            {
                if (firstId == ObjectId.Null) firstId = id;

                if (!string.IsNullOrEmpty(styleName))
                {
                    try
                    {
                        var style = (StyleBase)tr.GetObject(id, OpenMode.ForRead);
                        if (style.Name.Equals(styleName, StringComparison.OrdinalIgnoreCase))
                            return id;
                    }
                    catch { }
                }
            }
            return firstId;
        }
    }

    #endregion

    #region Style Settings Form

    /// <summary>
    /// Form quản lý style cho toàn bộ dự án
    /// Dark theme giống MenuSettingsForm
    /// </summary>
    public class StyleSettingsForm : CivilFormBase

        // ─── Data ───
        private StyleConfig _config;
        private readonly CivilDocument _cdoc;
        private readonly Transaction _tr;

        // ─── Controls ───
        private ListBox lstCategories = null!;
        private Panel pnlDetail = null!;
        private WinLabel lblCategoryTitle = null!;
        private readonly Dictionary<string, List<StyleRow>> _styleRows = new();

        // Danh sách categories
        private readonly string[] _categories = new[]
        {
            "Alignment (Tuyến)",
            "Profile (Trắc dọc)",
            "ProfileView (Khung TD)",
            "SampleLine (Cọc)",
            "SectionView (Trắc ngang)",
            "Surface (Bề mặt)",
            "Pipe & Structure (Ống)",
            "Point (Điểm)",
            "Chung (General)"
        };

        private class StyleRow
        {
            public string Label { get; set; } = "";
            public string ConfigKey { get; set; } = "";
            public ComboBox Combo { get; set; } = null!;
            public WinLabel LblInfo { get; set; } = null!;
        }

        // ─── Constructor ───
        public StyleSettingsForm(CivilDocument cdoc, Transaction tr)
        {
            _cdoc = cdoc;
            _tr = tr;
            _config = StyleSettings.Get();
            InitializeUI();
            PopulateCategories();
            lstCategories.SelectedIndex = 0;
        }

        public StyleConfig GetConfig() => _config;

        #region UI Initialization

        private void InitializeUI()
        {
            InitBaseTheme("Civil Tool — Cài Đặt Style", 920, 620);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            var header = CreateHeader("CÀI ĐẶT STYLE — CIVIL 3D");
            this.Controls.Add(header);

            // ── SUBTITLE ──
            var subtitle = new WinLabel
            {
                Text = "Chọn style mặc định cho từng loại đối tượng. Cài đặt sẽ áp dụng cho tất cả lệnh Civil Tool.",
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = TextDim,
                AutoSize = false,
                Height = 28,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Top,
                Padding = new Padding(18, 0, 0, 0),
                BackColor = BgPanel
            };
            this.Controls.Add(subtitle);
            subtitle.BringToFront();

            // ── BOTTOM BAR ──
            var pnlBottom = CreateBottomBar();
            this.Controls.Add(pnlBottom);

            var btnReset = CreateButton("Reset mặc định", 12, 8, 155, 36, pnlBottom);
            btnReset.Click += (s, e) =>
            {
                if (MessageBox.Show("Xóa tất cả cài đặt style và quay về mặc định?",
                    "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    _config = new StyleConfig();
                    StyleSettings.ClearCache();
                    if (lstCategories.SelectedIndex >= 0)
                        ShowCategoryDetail(lstCategories.SelectedIndex);
                    MessageBox.Show("Đã reset về mặc định!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };

            var btnCancel = CreateButton("Hủy", 640, 8, 100, 36, pnlBottom);
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            var btnSave = CreatePrimaryButton("Lưu & Áp dụng", 752, 8, 150, 36, pnlBottom);
            btnSave.Click += BtnSave_Click;

            // ── BACKGROUND PANEL for layout stability ──
            var mainContainer = CreateSection("", 10, 85, 885, 460);
            this.Controls.Add(mainContainer);
            mainContainer.BringToFront();

            // ── LEFT: Category List ──
            var pnlLeft = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(240, 440),
                BackColor = BgPanel,
                BorderStyle = BorderStyle.None
            };
            mainContainer.Controls.Add(pnlLeft);

            var lblCatHeader = new WinLabel
            {
                Text = "LOẠI ĐỐI TƯỢNG",
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = NavyMedium,
                Location = new Point(8, 8),
                AutoSize = true
            };
            pnlLeft.Controls.Add(lblCatHeader);

            lstCategories = new ListBox
            {
                Location = new Point(5, 32),
                Size = new Size(228, 400),
                BackColor = BgPanel,
                ForeColor = TextDark,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 10f),
                IntegralHeight = false,
                ItemHeight = 32,
                DrawMode = DrawMode.OwnerDrawFixed
            };
            lstCategories.DrawItem += LstCategories_DrawItem;
            lstCategories.SelectedIndexChanged += (s, e) =>
            {
                if (lstCategories.SelectedIndex >= 0)
                    ShowCategoryDetail(lstCategories.SelectedIndex);
            };
            pnlLeft.Controls.Add(lstCategories);

            // Vertical separator
            var vLine = new Panel
            {
                Location = new Point(255, 10),
                Size = new Size(1, 440),
                BackColor = BorderColor
            };
            mainContainer.Controls.Add(vLine);

            // ── RIGHT: Detail Panel ──
            var pnlRight = new Panel
            {
                Location = new Point(265, 10),
                Size = new Size(610, 440),
                BackColor = BgPanel,
                BorderStyle = BorderStyle.None
            };
            mainContainer.Controls.Add(pnlRight);

            lblCategoryTitle = new WinLabel
            {
                Text = "",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = NavyMedium,
                Location = new Point(10, 8),
                AutoSize = true
            };
            pnlRight.Controls.Add(lblCategoryTitle);

            // Separator line
            var separator = new Panel
            {
                Location = new Point(10, 38),
                Size = new Size(590, 1),
                BackColor = BorderColor
            };
            pnlRight.Controls.Add(separator);

            pnlDetail = new Panel
            {
                Location = new Point(10, 45),
                Size = new Size(590, 390),
                BackColor = BgPanel,
                AutoScroll = true
            };
            pnlRight.Controls.Add(pnlDetail);
        }

        private void LstCategories_DrawItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            e.DrawBackground();

            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Color bgColor = isSelected ? NavyMedium : BgPanel;
            Color fgColor = isSelected ? Color.White : TextDark;

            using var bgBrush = new SolidBrush(bgColor);
            e.Graphics.FillRectangle(bgBrush, e.Bounds);

            using var fgBrush = new SolidBrush(fgColor);
            var textBounds = new RectangleF(e.Bounds.X + 10, e.Bounds.Y + 6, e.Bounds.Width - 12, e.Bounds.Height);
            e.Graphics.DrawString(lstCategories.Items[e.Index]?.ToString() ?? "", e.Font!, fgBrush, textBounds);

            // Bottom border
            using var borderPen = new Pen(Color.FromArgb(208, 213, 221));
            e.Graphics.DrawLine(borderPen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
        }

        #endregion

        #region Category Detail

        private void PopulateCategories()
        {
            lstCategories.Items.Clear();
            foreach (var cat in _categories)
                lstCategories.Items.Add(cat);
        }

        private void ShowCategoryDetail(int categoryIndex)
        {
            pnlDetail.Controls.Clear();
            lblCategoryTitle.Text = _categories[categoryIndex];

            switch (categoryIndex)
            {
                case 0: BuildAlignmentPanel(); break;
                case 1: BuildProfilePanel(); break;
                case 2: BuildProfileViewPanel(); break;
                case 3: BuildSampleLinePanel(); break;
                case 4: BuildSectionViewPanel(); break;
                case 5: BuildSurfacePanel(); break;
                case 6: BuildPipePanel(); break;
                case 7: BuildPointPanel(); break;
                case 8: BuildGeneralPanel(); break;
            }
        }

        // ── ALIGNMENT ──
        private void BuildAlignmentPanel()
        {
            int y = 10;
            AddStyleCombo("Alignment Style:", "AlignmentStyle", _cdoc.Styles.AlignmentStyles, _config.AlignmentStyle, ref y);
            AddStyleCombo("Label Set Style:", "AlignmentLabelSet", _cdoc.Styles.LabelSetStyles.AlignmentLabelSetStyles, _config.AlignmentLabelSet, ref y);
            AddInfoLabel("Style mặc định sẽ áp dụng cho: CTA_TaoTuyenHangLoat, CTA_TaoTuyenNhanh", ref y);
        }

        // ── PROFILE ──
        private void BuildProfilePanel()
        {
            int y = 10;
            AddStyleCombo("Profile Style:", "ProfileStyle", _cdoc.Styles.ProfileStyles, _config.ProfileStyle, ref y);
            AddStyleCombo("Label Set Style:", "ProfileLabelSet", _cdoc.Styles.LabelSetStyles.ProfileLabelSetStyles, _config.ProfileLabelSet, ref y);
            AddInfoLabel("Áp dụng cho: CTP_VeTracDoc_TuNhien, CTP_VeTracDoc_TuNhien_TatCaTuyen", ref y);
        }

        // ── PROFILE VIEW ──
        private void BuildProfileViewPanel()
        {
            int y = 10;
            AddStyleCombo("ProfileView Style:", "ProfileViewStyle", _cdoc.Styles.ProfileViewStyles, _config.ProfileViewStyle, ref y);
            AddStyleCombo("Band Set Style:", "ProfileViewBandSet", _cdoc.Styles.ProfileViewBandSetStyles, _config.ProfileViewBandSet, ref y);
            AddInfoLabel("Áp dụng cho: CTP_VeTracDoc_TuNhien, CTPV_TaoProfileView", ref y);
        }

        // ── SAMPLELINE ──
        private void BuildSampleLinePanel()
        {
            int y = 10;
            AddStyleCombo("SampleLine Style:", "SampleLineStyle", _cdoc.Styles.SampleLineStyles, _config.SampleLineStyle, ref y);
            AddStyleCombo("Label Style:", "SampleLineLabelStyle", _cdoc.Styles.LabelStyles.SampleLineLabelStyles.LabelStyles, _config.SampleLineLabelStyle, ref y);
            AddInfoLabel("Áp dụng cho: CTS_PhatSinhCoc, CTS_DoiTenCoc", ref y);
        }

        // ── SECTION VIEW ──
        private void BuildSectionViewPanel()
        {
            int y = 10;
            AddStyleCombo("SectionView Style:", "SectionViewStyle", _cdoc.Styles.SectionViewStyles, _config.SectionViewStyle, ref y);
            AddStyleCombo("Band Set Style:", "SectionViewBandSet", _cdoc.Styles.SectionViewBandSetStyles, _config.SectionViewBandSet, ref y);
            AddInfoLabel("Áp dụng cho: CTSV_VeTracNgangThietKe, CVSV_VeTatCa_TracNgangThietKe", ref y);
        }

        // ── SURFACE ──
        private void BuildSurfacePanel()
        {
            int y = 10;
            AddStyleCombo("Surface Style:", "SurfaceStyle", _cdoc.Styles.SurfaceStyles, _config.SurfaceStyle, ref y);
            AddInfoLabel("Áp dụng cho: CTSu_TaoSpotElevation_OnSurface_TaiTim", ref y);
        }

        // ── PIPE & STRUCTURE ──
        private void BuildPipePanel()
        {
            int y = 10;
            AddStyleCombo("Pipe Style:", "PipeStyle", _cdoc.Styles.PipeStyles, _config.PipeStyle, ref y);
            AddStyleCombo("Structure Style:", "StructureStyle", _cdoc.Styles.StructureStyles, _config.StructureStyle, ref y);
            AddInfoLabel("Áp dụng cho: CTPS_TaoBangThongKePipe, CTPS_TaoBangThongKeStructure", ref y);
        }

        // ── POINT ──
        private void BuildPointPanel()
        {
            int y = 10;
            AddStyleCombo("Point Style:", "PointStyle", _cdoc.Styles.PointStyles, _config.PointStyle, ref y);
            AddStyleCombo("Point Label Style:", "PointLabelStyle", _cdoc.Styles.LabelStyles.PointLabelStyles.LabelStyles, _config.PointLabelStyle, ref y);
            AddInfoLabel("Áp dụng cho: CTPo_TaoPointTheoBang, CTPo_ThayDoiStyle", ref y);
        }

        // ── GENERAL ──
        private void BuildGeneralPanel()
        {
            int y = 10;

            // Site selection
            AddLabel("Site mặc định:", ref y);
            var cmbSite = new ComboBox
            {
                Location = new Point(180, y - 22),
                Size = new Size(380, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextWhite,
                Font = new Font("Segoe UI", 9.5f),
                FlatStyle = FlatStyle.Flat
            };
            cmbSite.Items.Add("<< Không chọn Site (Siteless) >>");
            int selectedSiteIdx = 0;
            try
            {
                ObjectIdCollection siteIds = _cdoc.GetSiteIds();
                int idx = 1;
                foreach (ObjectId sId in siteIds)
                {
                    var site = (Site)_tr.GetObject(sId, OpenMode.ForRead);
                    cmbSite.Items.Add(site.Name);
                    if (site.Name.Equals(_config.DefaultSite, StringComparison.OrdinalIgnoreCase))
                        selectedSiteIdx = idx;
                    idx++;
                }
            }
            catch { }
            cmbSite.SelectedIndex = selectedSiteIdx;
            cmbSite.SelectedIndexChanged += (s, e) =>
            {
                _config.DefaultSite = cmbSite.SelectedIndex <= 0 ? "" : cmbSite.SelectedItem?.ToString() ?? "";
            };
            pnlDetail.Controls.Add(cmbSite);
            y += 15;

            // Default Layer
            AddLabel("Layer mặc định:", ref y);
            var txtLayer = new TextBox
            {
                Location = new Point(180, y - 22),
                Size = new Size(380, 28),
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f),
                BorderStyle = BorderStyle.FixedSingle,
                Text = _config.DefaultLayer
            };
            txtLayer.TextChanged += (s, e) => _config.DefaultLayer = txtLayer.Text;
            pnlDetail.Controls.Add(txtLayer);
            y += 15;

            AddInfoLabel("Site và Layer sẽ áp dụng khi tạo Alignment, Profile mới.", ref y);

            // Thông tin config hiện tại
            y += 10;
            var lblConfigInfo = new WinLabel
            {
                Text = $"Config: {Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CivilTool", "style_config.json")}\n" +
                       $"Cập nhật: {_config.LastModified:dd/MM/yyyy HH:mm}",
                Location = new Point(10, y),
                Size = new Size(570, 45),
                ForeColor = TextDim,
                Font = new Font("Segoe UI", 8.5f)
            };
            pnlDetail.Controls.Add(lblConfigInfo);
        }

        #endregion

        #region UI Helpers

        private void AddStyleCombo(string label, string configKey, System.Collections.IEnumerable styleCollection, string currentValue, ref int y)
        {
            var lbl = new WinLabel
            {
                Text = label,
                Location = new Point(10, y + 3),
                Size = new Size(160, 22),
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            pnlDetail.Controls.Add(lbl);

            // ComboBox
            var cmb = new ComboBox
            {
                Location = new Point(180, y),
                Size = new Size(380, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f),
                FlatStyle = FlatStyle.Flat
            };

            // Populate
            cmb.Items.Add("<< Mặc định (style đầu tiên) >>");
            int selectedIdx = 0;
            int idx = 1;
            try
            {
                foreach (ObjectId id in styleCollection)
                {
                    try
                    {
                        var style = (StyleBase)_tr.GetObject(id, OpenMode.ForRead);
                        cmb.Items.Add(style.Name);
                        if (!string.IsNullOrEmpty(currentValue) &&
                            style.Name.Equals(currentValue, StringComparison.OrdinalIgnoreCase))
                        {
                            selectedIdx = idx;
                        }
                        idx++;
                    }
                    catch { }
                }
            }
            catch { }

            cmb.SelectedIndex = selectedIdx;

            // Info label (hiển thị số lượng style)
            var lblCount = new WinLabel
            {
                Text = $"({idx - 1} style)",
                Location = new Point(568, y + 3),
                AutoSize = true,
                ForeColor = TextDim,
                Font = new Font("Segoe UI", 8.5f)
            };
            pnlDetail.Controls.Add(lblCount);

            // Event: cập nhật config khi chọn
            string key = configKey;
            cmb.SelectedIndexChanged += (s, e) =>
            {
                string value = cmb.SelectedIndex <= 0 ? "" : cmb.SelectedItem?.ToString() ?? "";
                SetConfigValue(key, value);
            };

            pnlDetail.Controls.Add(cmb);
            y += 42;
        }

        private void AddLabel(string text, ref int y)
        {
            var lbl = new WinLabel
            {
                Text = text,
                Location = new Point(10, y + 3),
                Size = new Size(160, 22),
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            pnlDetail.Controls.Add(lbl);
            y += 30;
        }

        private void AddInfoLabel(string text, ref int y)
        {
            y += 8;
            var lbl = new WinLabel
            {
                Text = text,
                Location = new Point(10, y),
                Size = new Size(580, 30),
                ForeColor = Color.FromArgb(211, 47, 47),
                Font = new Font("Segoe UI", 8.5f, FontStyle.Italic)
            };
            pnlDetail.Controls.Add(lbl);
            y += 35;
        }

        private void SetConfigValue(string key, string value)
        {
            var prop = typeof(StyleConfig).GetProperty(key);
            if (prop != null && prop.PropertyType == typeof(string))
            {
                prop.SetValue(_config, value);
            }
        }

        private Button CreateButton(string text, int x, int y, int w, int h, Control parent)
        {
            var btn = base.CreateButton(text, x, y, w, h, parent);
            btn.TextAlign = ContentAlignment.MiddleCenter;
            return btn;
        }

        #endregion

        #region Save

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            StyleSettings.Save(_config);
            MessageBox.Show(
                "Đã lưu cài đặt Style!\n\nCác lệnh Civil Tool sẽ sử dụng style đã chọn.",
                "Thành công",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        #endregion
    }

    #endregion

    #region Commands

    /// <summary>
    /// Lệnh mở Form cài đặt Style
    /// </summary>
    public class StyleSettingsCommands
    {
        [CommandMethod("CT_StyleSettings")]
        public static void CT_StyleSettings()
        {
            try
            {
                Database db = A.Db;
                CivilDocument cdoc = A.Cdoc;

                using Transaction tr = db.TransactionManager.StartTransaction();
                try
                {
                    using var form = new StyleSettingsForm(cdoc, tr);
                    var result = AcadApp.ShowModalDialog(form);

                    if (result == DialogResult.OK)
                    {
                        A.Ed.WriteMessage("\n Đã cập nhật cài đặt Style thành công!");
                    }
                    else
                    {
                        A.Ed.WriteMessage("\n Đã hủy cài đặt Style.");
                    }

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
                }
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi khởi tạo: {ex.Message}");
            }
        }

        /// <summary>
        /// Lệnh hiển thị style hiện tại đang dùng (console)
        /// </summary>
        [CommandMethod("CT_StyleInfo")]
        public static void CT_StyleInfo()
        {
            var config = StyleSettings.Get();
            var ed = A.Ed;

            ed.WriteMessage("\n╔════════════════════════════════════════════╗");
            ed.WriteMessage("\n║     STYLE CÀI ĐẶT HIỆN TẠI               ║");
            ed.WriteMessage("\n╠════════════════════════════════════════════╣");
            PrintStyleRow(ed, "Alignment Style", config.AlignmentStyle);
            PrintStyleRow(ed, "Alignment Label", config.AlignmentLabelSet);
            PrintStyleRow(ed, "Profile Style", config.ProfileStyle);
            PrintStyleRow(ed, "Profile Label", config.ProfileLabelSet);
            PrintStyleRow(ed, "ProfileView Style", config.ProfileViewStyle);
            PrintStyleRow(ed, "ProfileView Band", config.ProfileViewBandSet);
            PrintStyleRow(ed, "SampleLine Style", config.SampleLineStyle);
            PrintStyleRow(ed, "SectionView Style", config.SectionViewStyle);
            PrintStyleRow(ed, "SectionView Band", config.SectionViewBandSet);
            PrintStyleRow(ed, "Surface Style", config.SurfaceStyle);
            PrintStyleRow(ed, "Pipe Style", config.PipeStyle);
            PrintStyleRow(ed, "Structure Style", config.StructureStyle);
            PrintStyleRow(ed, "Point Style", config.PointStyle);
            PrintStyleRow(ed, "Point Label", config.PointLabelStyle);
            PrintStyleRow(ed, "Default Site", config.DefaultSite);
            PrintStyleRow(ed, "Default Layer", config.DefaultLayer);
            ed.WriteMessage("\n╚════════════════════════════════════════════╝\n");
        }

        private static void PrintStyleRow(Autodesk.AutoCAD.EditorInput.Editor ed, string label, string value)
        {
            string displayVal = string.IsNullOrEmpty(value) ? "(mặc định)" : value;
            ed.WriteMessage($"\n║  {label,-20}: {displayVal,-23}║");
        }
    }

    #endregion
}
