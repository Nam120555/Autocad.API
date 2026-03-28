// 21.LayerControl.cs - Điều khiển Layer ON/OFF cho Civil 3D Objects
// Bao gồm Form LayerControlHub gộp tất cả lệnh CTL_*
// Kế thừa CivilFormBase cho theme Navy/White

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using ContentAlignment = System.Drawing.ContentAlignment;

[assembly: CommandClass(typeof(MyFirstProject.LayerControlCommands))]

namespace MyFirstProject
{
    // ══════════════════════════════════════════════════════════════
    // DATA CLASS — Thông tin 1 nhóm layer
    // ══════════════════════════════════════════════════════════════
    
    public class LayerGroupInfo
    {
        public string DisplayName { get; set; } = "";
        public string LayerName { get; set; } = "";
        public string Icon { get; set; } = "";
        public Color TagColor { get; set; } = Color.Gray;
        public bool IsOn { get; set; } = true;
    }

    // ══════════════════════════════════════════════════════════════
    // FORM: LAYER CONTROL HUB
    // ══════════════════════════════════════════════════════════════

    /// <summary>
    /// Form gộp tất cả lệnh CTL_On/Off — checkbox để bật/tắt từng nhóm layer
    /// </summary>
    public class LayerControlHubForm : CivilFormBase
    {
        private List<LayerGroupInfo> layerGroups = new();
        private Panel layerListPanel = null!;
        private List<CheckBox> checkBoxes = new();

        public LayerControlHubForm()
        {
            BuildLayerGroups();
            InitializeComponent();
        }

        private void BuildLayerGroups()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;

            // Định nghĩa các nhóm layer Civil 3D
            var definitions = new List<(string Display, string Layer, string Icon, Color TagColor)>
            {
                ("Corridor (Hành lang)", "C-ROAD-CORR-LINK", "🛣️", Color.FromArgb(76, 175, 80)),
                ("SampleLine (Mắt cắt)", "C-ROAD-SAMP", "📐", Color.FromArgb(33, 150, 243)),
                ("Alignment (Tim tuyến)", "C-ROAD-CNTR-off", "📏", Color.FromArgb(255, 152, 0)),
                ("Parcel (Thửa đất)", "C-PROP-BNDY", "🔲", Color.FromArgb(156, 39, 176)),
                ("Hatch Đào Đắp", "$$$HATCH DAO DAP", "🎨", Color.FromArgb(244, 67, 54)),
                ("Defpoints", "Defpoints", "📌", Color.FromArgb(158, 158, 158)),
            };

            using var tr = db.TransactionManager.StartOpenCloseTransaction();
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            foreach (var (display, layer, icon, color) in definitions)
            {
                bool exists = lt.Has(layer);
                bool isOn = true;

                if (exists)
                {
                    var ltr = (LayerTableRecord)tr.GetObject(lt[layer], OpenMode.ForRead);
                    isOn = !ltr.IsOff;
                }

                layerGroups.Add(new LayerGroupInfo
                {
                    DisplayName = display,
                    LayerName = layer,
                    Icon = icon,
                    TagColor = color,
                    IsOn = exists ? isOn : false
                });
            }
            tr.Commit();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("⚡  LAYER CONTROL HUB", 420, 480);

            // ═══ HEADER ═══
            var header = CreateHeader("LAYER CONTROL HUB");
            this.Controls.Add(header);

            var subtitle = CreateSubtitle("Bật / Tắt nhanh các nhóm Layer Civil 3D");
            this.Controls.Add(subtitle);

            // ═══ NÚT NHANH ═══
            int btnY = 82;
            var btnAllOn = CreatePrimaryButton("☑ BẬT TẤT CẢ", 20, btnY, 180, 32);
            btnAllOn.Click += (s, e) => SetAll(true);
            this.Controls.Add(btnAllOn);

            var btnAllOff = CreateButton("☐ TẮT TẤT CẢ", 210, btnY, 180, 32);
            btnAllOff.Click += (s, e) => SetAll(false);
            this.Controls.Add(btnAllOff);

            // ═══ DANH SÁCH LAYER ═══
            layerListPanel = new Panel
            {
                Location = new Point(20, 124),
                Size = new Size(370, layerGroups.Count * 50 + 10),
                AutoScroll = true,
                BackColor = BgPanel,
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(layerListPanel);

            int y = 8;
            for (int i = 0; i < layerGroups.Count; i++)
            {
                var group = layerGroups[i];
                AddLayerRow(group, y, i);
                y += 50;
            }

            // ═══ BOTTOM BAR ═══
            var bottomBar = CreateBottomBar(56);
            this.Controls.Add(bottomBar);

            var btnClose = CreateButton("✕ Đóng", 10, 10, 170, 36, bottomBar);
            btnClose.Click += (s, e) => this.Close();

            var btnApply = CreateSuccessButton("✓ Áp Dụng & Đóng", 190, 10, 175, 36, bottomBar);
            btnApply.Click += BtnApply_Click;

            this.AcceptButton = btnApply;
            this.CancelButton = btnClose;
        }

        private void AddLayerRow(LayerGroupInfo group, int y, int index)
        {
            // Panel cho mỗi row
            var rowPanel = new Panel
            {
                Location = new Point(5, y),
                Size = new Size(346, 44),
                BackColor = BgWhite
            };

            // Viền màu bên trái
            rowPanel.Paint += (s, e) =>
            {
                using var brush = new SolidBrush(group.TagColor);
                e.Graphics.FillRectangle(brush, 0, 0, 5, rowPanel.Height);
            };

            // CheckBox
            var chk = new CheckBox
            {
                Location = new Point(14, 12),
                AutoSize = true,
                Checked = group.IsOn,
                Text = "",
                FlatStyle = FlatStyle.Flat,
                Tag = index
            };
            chk.CheckedChanged += (s, e) =>
            {
                layerGroups[index].IsOn = chk.Checked;
                UpdateRowVisual(rowPanel, chk.Checked, group.TagColor);
            };
            checkBoxes.Add(chk);
            rowPanel.Controls.Add(chk);

            // Icon + Tên
            var lbl = new WinLabel
            {
                Text = $"{group.Icon}  {group.DisplayName}",
                Location = new Point(38, 4),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            lbl.Click += (s, e) => { chk.Checked = !chk.Checked; };
            rowPanel.Controls.Add(lbl);

            // Tên layer gốc
            var lblLayer = new WinLabel
            {
                Text = $"Layer: {group.LayerName}",
                Location = new Point(38, 24),
                AutoSize = true,
                ForeColor = TextHint,
                Font = new Font("Segoe UI", 7.5f),
                Cursor = Cursors.Hand
            };
            lblLayer.Click += (s, e) => { chk.Checked = !chk.Checked; };
            rowPanel.Controls.Add(lblLayer);

            // Trạng thái
            var lblStatus = new WinLabel
            {
                Text = group.IsOn ? "● ON" : "○ OFF",
                Location = new Point(290, 12),
                AutoSize = true,
                ForeColor = group.IsOn ? GreenOk : TextHint,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            chk.CheckedChanged += (s, e) =>
            {
                lblStatus.Text = chk.Checked ? "● ON" : "○ OFF";
                lblStatus.ForeColor = chk.Checked ? GreenOk : TextHint;
            };
            rowPanel.Controls.Add(lblStatus);

            UpdateRowVisual(rowPanel, group.IsOn, group.TagColor);
            layerListPanel.Controls.Add(rowPanel);
        }

        private void UpdateRowVisual(Panel panel, bool isOn, Color tagColor)
        {
            panel.BackColor = isOn ? BgWhite : Color.FromArgb(245, 245, 245);
        }

        private void SetAll(bool on)
        {
            foreach (var chk in checkBoxes)
                chk.Checked = on;
        }

        private void BtnApply_Click(object? sender, EventArgs e)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            using var tr = db.TransactionManager.StartTransaction();
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            int changed = 0;
            foreach (var group in layerGroups)
            {
                if (lt.Has(group.LayerName))
                {
                    var layer = (LayerTableRecord)tr.GetObject(lt[group.LayerName], OpenMode.ForWrite);
                    bool currentlyOff = layer.IsOff;
                    bool shouldBeOff = !group.IsOn;

                    if (currentlyOff != shouldBeOff)
                    {
                        layer.IsOff = shouldBeOff;
                        changed++;
                        ed.WriteMessage($"\n  {group.Icon} {group.DisplayName}: {(group.IsOn ? "BẬT" : "TẮT")}");
                    }
                }
            }

            tr.Commit();
            ed.Regen();

            if (changed > 0)
                ed.WriteMessage($"\n✅ Đã thay đổi {changed} layer.\n");
            else
                ed.WriteMessage($"\n◎ Không có thay đổi.\n");

            this.Close();
        }
    }

    // ══════════════════════════════════════════════════════════════
    // LỆNH CHÍNH
    // ══════════════════════════════════════════════════════════════

    /// <summary>
    /// Các lệnh điều khiển Layer ON/OFF cho Civil 3D objects
    /// </summary>
    public class LayerControlCommands
    {
        // ══════════════════════════════════════════════════════════════
        // LỆNH MỚI: LAYER CONTROL HUB (mở form gộp)
        // ══════════════════════════════════════════════════════════════

        /// <summary>
        /// Mở form Layer Control Hub — gộp tất cả On/Off vào 1 form
        /// </summary>
        [CommandMethod("CTL_Hub")]
        public static void CTL_Hub()
        {
            using var form = new LayerControlHubForm();
            AcadApp.ShowModalDialog(form);
        }

        // ══════════════════════════════════════════════════════════════
        // CÁC LỆNH GỐC (giữ lại cho backward compatibility)
        // ══════════════════════════════════════════════════════════════

        private static Editor GetEditor()
        {
            return AcadApp.DocumentManager.MdiActiveDocument?.Editor!;
        }

        private static void SetLayerState(string layerName, bool turnOn)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var db = doc.Database;
            var ed = doc.Editor;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                
                if (lt.Has(layerName))
                {
                    var layer = (LayerTableRecord)tr.GetObject(lt[layerName], OpenMode.ForWrite);
                    layer.IsOff = !turnOn;
                    ed.WriteMessage($"\n◎ Layer '{layerName}' đã {(turnOn ? "BẬT" : "TẮT")}.");
                }
                else
                {
                    ed.WriteMessage($"\n⊘ Layer '{layerName}' không tồn tại.");
                }
                
                tr.Commit();
            }
            ed.Regen();
        }

        // ─── CORRIDOR ───
        [CommandMethod("CTL_OnCorridor")]
        public static void CTL_OnCorridor() => SetLayerState("C-ROAD-CORR-LINK", true);

        [CommandMethod("CTL_OffCorridor")]
        public static void CTL_OffCorridor() => SetLayerState("C-ROAD-CORR-LINK", false);

        // ─── SAMPLE LINE ───
        [CommandMethod("CTL_OnSampleLine")]
        public static void CTL_OnSampleLine() => SetLayerState("C-ROAD-SAMP", true);

        [CommandMethod("CTL_OffSampleLine")]
        public static void CTL_OffSampleLine() => SetLayerState("C-ROAD-SAMP", false);

        // ─── ALIGNMENT ───
        [CommandMethod("CTL_OnAlignment")]
        public static void CTL_OnAlignment() => SetLayerState("C-ROAD-CNTR-off", true);

        [CommandMethod("CTL_OffAlignment")]
        public static void CTL_OffAlignment() => SetLayerState("C-ROAD-CNTR-off", false);

        // ─── PARCEL ───
        [CommandMethod("CTL_OnParcel")]
        public static void CTL_OnParcel() => SetLayerState("C-PROP-BNDY", true);

        [CommandMethod("CTL_OffParcel")]
        public static void CTL_OffParcel() => SetLayerState("C-PROP-BNDY", false);

        // ─── HATCH ĐÀO ĐẮP ───
        [CommandMethod("CTL_OnHatchDaoDap")]
        public static void CTL_OnHatchDaoDap() => SetLayerState("$$$HATCH DAO DAP", true);

        [CommandMethod("CTL_OffHatchDaoDap")]
        public static void CTL_OffHatchDaoDap() => SetLayerState("$$$HATCH DAO DAP", false);

        // ─── DEFPOINTS ───
        [CommandMethod("CTL_OnDefpoints")]
        public static void CTL_OnDefpoints() => SetLayerState("Defpoints", true);

        [CommandMethod("CTL_OffDefpoints")]
        public static void CTL_OffDefpoints() => SetLayerState("Defpoints", false);

        // ══════════════════════════════════════════════════════════════
        // CHUYỂN ĐỐI TƯỢNG SANG LAYER KHÁC
        // ══════════════════════════════════════════════════════════════
        
        private static void MoveToLayer(string layerName)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var db = doc.Database;

            var psr = ed.GetSelection(new PromptSelectionOptions
            {
                MessageForAdding = $"\n⊙ Chọn đối tượng để chuyển sang layer '{layerName}': "
            });

            if (psr.Status != PromptStatus.OK) return;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                if (!lt.Has(layerName))
                {
                    lt.UpgradeOpen();
                    var newLayer = new LayerTableRecord { Name = layerName };
                    lt.Add(newLayer);
                    tr.AddNewlyCreatedDBObject(newLayer, true);
                    ed.WriteMessage($"\n  ✓ Đã tạo layer mới: {layerName}");
                }

                int count = 0;
                foreach (SelectedObject so in psr.Value)
                {
                    var ent = tr.GetObject(so.ObjectId, OpenMode.ForWrite) as Entity;
                    if (ent != null)
                    {
                        ent.Layer = layerName;
                        count++;
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n◎ Đã chuyển {count} đối tượng sang layer '{layerName}'.");
            }
        }

        [CommandMethod("CTL_ToText")]
        public static void CTL_ToText() => MoveToLayer("0.TEXT");

        [CommandMethod("CTL_ToDefpoints")]
        public static void CTL_ToDefpoints() => MoveToLayer("Defpoints");

        [CommandMethod("CTL_ToDim")]
        public static void CTL_ToDim() => MoveToLayer("1.DIM");

        [CommandMethod("CTL_ToBaoBT")]
        public static void CTL_ToBaoBT() => MoveToLayer("2.BAO BT");

        [CommandMethod("CTL_ToBaoCotThep")]
        public static void CTL_ToBaoCotThep() => MoveToLayer("3.BAO COT THEP");

        [CommandMethod("CTL_ToThep")]
        public static void CTL_ToThep() => MoveToLayer("4.THEP");

        [CommandMethod("CTL_ToTruc")]
        public static void CTL_ToTruc() => MoveToLayer("5.TRUC");

        [CommandMethod("CTL_ToKhuat")]
        public static void CTL_ToKhuat() => MoveToLayer("6.KHUAT");

        [CommandMethod("CTL_ToHatch")]
        public static void CTL_ToHatch() => MoveToLayer("7.HATCH");

        [CommandMethod("CTL_ToRanhGioi")]
        public static void CTL_ToRanhGioi() => MoveToLayer("8.RANH GIOI");
    }
}
