using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using MyFirstProject.Extensions;

using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.AlignmentHangLoat))]

namespace MyFirstProject
{
    // ═══════════════════════════════════════════════════════════════
    //  FORM: Tạo Alignment hàng loạt
    // ═══════════════════════════════════════════════════════════════
    public class AlignmentBatchForm : CivilFormBase
    {
        private TextBox txtPrefix = null!;
        private NumericUpDown numStart = null!;
        private ComboBox cboStyle = null!;
        private ComboBox cboLabelSet = null!;
        private ComboBox cboSite = null!;
        private TextBox txtLayer = null!;
        private WinLabel lblPolyCount = null!;

        public bool IsConfirmed { get; private set; }
        public string Prefix => txtPrefix.Text;
        public int StartIndex => (int)numStart.Value;
        public string LayerName => txtLayer.Text;

        private List<(ObjectId Id, string Name)> styles = new();
        private List<(ObjectId Id, string Name)> labelSets = new();
        private List<(ObjectId Id, string Name)> sites = new();

        public AlignmentBatchForm(int polylineCount)
        {
            LoadCivilData();
            InitUI(polylineCount);
        }

        private void LoadCivilData()
        {
            try
            {
                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                CivilDocument cdoc = CivilApplication.ActiveDocument;

                foreach (ObjectId sId in cdoc.Styles.AlignmentStyles)
                {
                    AlignmentStyle st = (AlignmentStyle)tr.GetObject(sId, OpenMode.ForRead);
                    styles.Add((sId, st.Name));
                }

                foreach (ObjectId lId in cdoc.Styles.LabelSetStyles.AlignmentLabelSetStyles)
                {
                    StyleBase lb = (StyleBase)tr.GetObject(lId, OpenMode.ForRead);
                    labelSets.Add((lId, lb.Name));
                }

                foreach (ObjectId stId in cdoc.GetSiteIds())
                {
                    Site site = (Site)tr.GetObject(stId, OpenMode.ForRead);
                    sites.Add((stId, site.Name));
                }

                tr.Commit();
            }
            catch { }
        }

        private void InitUI(int polyCount)
        {
            InitBaseTheme("TẠO TUYẾN HÀNG LOẠT — Civil Tool", 600, 520);

            this.Controls.Add(CreateHeader("🛣  TẠO TUYẾN HÀNG LOẠT TỪ POLYLINE"));

            var sub = CreateSubtitle($"Đã chọn {polyCount} Polyline — Thiết lập thông số bên dưới");
            this.Controls.Add(sub);
            sub.BringToFront();

            var main = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20, 12, 20, 0), AutoScroll = true };
            this.Controls.Add(main);
            main.BringToFront();

            int y = 5;

            lblPolyCount = CreateTitleLabel($"📐 Số Polyline đã chọn: {polyCount}", 0, y);
            main.Controls.Add(lblPolyCount);
            y += 35;

            // ═══ GROUP 1: TÊN TUYẾN ═══
            var grp1 = CreateGroupBox("📝  Đặt tên tuyến", 0, y, 545, 95);
            main.Controls.Add(grp1);

            grp1.Controls.Add(CreateLabel("Prefix:", 12, 30));
            txtPrefix = CreateTextBox(100, 27, 200, "Tuyen");
            grp1.Controls.Add(txtPrefix);

            grp1.Controls.Add(CreateLabel("Số bắt đầu:", 320, 30));
            numStart = CreateNumeric(420, 27, 100, 0, 9999, 1, 0);
            grp1.Controls.Add(numStart);

            grp1.Controls.Add(CreateHintLabel("Tên sẽ là: Tuyen-001, Tuyen-002, ...", 12, 65));
            y += 105;

            // ═══ GROUP 2: STYLE ═══
            var grp2 = CreateGroupBox("🎨  Style & Label", 0, y, 545, 110);
            main.Controls.Add(grp2);

            grp2.Controls.Add(CreateLabel("Alignment Style:", 12, 30));
            cboStyle = CreateComboBox(160, 27, 370);
            foreach (var s in styles) cboStyle.Items.Add(s.Name);
            if (cboStyle.Items.Count > 0) cboStyle.SelectedIndex = 0;
            grp2.Controls.Add(cboStyle);

            grp2.Controls.Add(CreateLabel("Label Set Style:", 12, 65));
            cboLabelSet = CreateComboBox(160, 62, 370);
            foreach (var l in labelSets) cboLabelSet.Items.Add(l.Name);
            if (cboLabelSet.Items.Count > 0) cboLabelSet.SelectedIndex = 0;
            grp2.Controls.Add(cboLabelSet);

            y += 120;

            // ═══ GROUP 3: SITE & LAYER ═══
            var grp3 = CreateGroupBox("📂  Site & Layer", 0, y, 545, 110);
            main.Controls.Add(grp3);

            grp3.Controls.Add(CreateLabel("Site:", 12, 30));
            cboSite = CreateComboBox(100, 27, 430);
            cboSite.Items.Add("<< Không chọn Site (Siteless) >>");
            foreach (var s in sites) cboSite.Items.Add(s.Name);
            cboSite.SelectedIndex = 0;
            grp3.Controls.Add(cboSite);

            grp3.Controls.Add(CreateLabel("Layer:", 12, 65));
            txtLayer = CreateTextBox(100, 62, 430, "C-ROAD-ALGN");
            grp3.Controls.Add(txtLayer);

            y += 120;

            var note = CreateRedNote("⚠ Nếu tên trùng, hệ thống sẽ tự thêm hậu tố _1, _2...", 0, y);
            main.Controls.Add(note);

            var bottom = CreateBottomBar();
            this.Controls.Add(bottom);

            CreateButton("✖ Hủy", 340, 8, 100, 34, bottom).Click += (s, e) =>
            { this.DialogResult = DialogResult.Cancel; this.Close(); };

            CreateSuccessButton("✅ Tạo tuyến", 450, 8, 120, 34, bottom).Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtPrefix.Text))
                {
                    MessageBox.Show("Vui lòng nhập prefix tên tuyến!", "Thiếu dữ liệu",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                IsConfirmed = true;
                this.DialogResult = DialogResult.OK;
                this.Close();
            };
        }

        public ObjectId GetSelectedStyleId() =>
            cboStyle.SelectedIndex >= 0 && cboStyle.SelectedIndex < styles.Count
                ? styles[cboStyle.SelectedIndex].Id : ObjectId.Null;

        public ObjectId GetSelectedLabelSetId() =>
            cboLabelSet.SelectedIndex >= 0 && cboLabelSet.SelectedIndex < labelSets.Count
                ? labelSets[cboLabelSet.SelectedIndex].Id : ObjectId.Null;

        public ObjectId GetSelectedSiteId() =>
            cboSite.SelectedIndex > 0 && (cboSite.SelectedIndex - 1) < sites.Count
                ? sites[cboSite.SelectedIndex - 1].Id : ObjectId.Null;
    }

    // ═══════════════════════════════════════════════════════════════
    //  FORM: Đổi tên Alignment hàng loạt
    // ═══════════════════════════════════════════════════════════════
    public class RenameBatchForm : CivilFormBase
    {
        private TextBox txtPrefix = null!;
        private ListView lvAlignments = null!;

        public bool IsConfirmed { get; private set; }
        public string NewPrefix => txtPrefix.Text;

        private List<(ObjectId Id, string Name)> alignments;

        public RenameBatchForm(List<(ObjectId Id, string Name)> aligns)
        {
            alignments = aligns;
            InitUI();
        }

        private void InitUI()
        {
            InitBaseTheme("ĐỔI TÊN TUYẾN HÀNG LOẠT — Civil Tool", 550, 450);

            this.Controls.Add(CreateHeader("✏️  ĐỔI TÊN TUYẾN HÀNG LOẠT"));

            var sub = CreateSubtitle($"Có {alignments.Count} tuyến trong bản vẽ");
            this.Controls.Add(sub);
            sub.BringToFront();

            var main = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20, 12, 20, 0) };
            this.Controls.Add(main);
            main.BringToFront();

            int y = 5;

            main.Controls.Add(CreateLabel("Prefix mới:", 0, y));
            txtPrefix = CreateTextBox(100, y - 3, 390);
            main.Controls.Add(txtPrefix);
            y += 35;

            main.Controls.Add(CreateHintLabel("Tên mới sẽ là: [Prefix]-001, [Prefix]-002, ...", 0, y));
            y += 25;

            main.Controls.Add(CreateTitleLabel("Danh sách tuyến hiện có:", 0, y));
            y += 25;

            lvAlignments = CreateListView(0, y, 495, 230);
            lvAlignments.Columns.Add("STT", 50);
            lvAlignments.Columns.Add("Tên hiện tại", 220);
            lvAlignments.Columns.Add("Tên mới (preview)", 220);
            main.Controls.Add(lvAlignments);

            int stt = 1;
            foreach (var a in alignments)
            {
                var item = new ListViewItem(stt.ToString());
                item.SubItems.Add(a.Name);
                item.SubItems.Add($"???-{stt:D3}");
                lvAlignments.Items.Add(item);
                stt++;
            }

            txtPrefix.TextChanged += (s, e) =>
            {
                for (int i = 0; i < lvAlignments.Items.Count; i++)
                {
                    string p = string.IsNullOrWhiteSpace(txtPrefix.Text) ? "???" : txtPrefix.Text;
                    lvAlignments.Items[i].SubItems[2].Text = $"{p}-{(i + 1):D3}";
                }
            };

            var bottom = CreateBottomBar();
            this.Controls.Add(bottom);

            CreateButton("✖ Hủy", 300, 8, 100, 34, bottom).Click += (s, e) =>
            { this.DialogResult = DialogResult.Cancel; this.Close(); };

            CreatePrimaryButton("✅ Đổi tên", 410, 8, 110, 34, bottom).Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtPrefix.Text))
                {
                    MessageBox.Show("Vui lòng nhập prefix mới!", "Thiếu dữ liệu",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                IsConfirmed = true;
                this.DialogResult = DialogResult.OK;
                this.Close();
            };
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  LỆNH
    // ═══════════════════════════════════════════════════════════════
    class AlignmentHangLoat
    {
        [CommandMethod("CTA_TaoTuyenHangLoat")]
        public static void CTA_TaoTuyenHangLoat()
        {
            Editor ed = A.Ed;
            Database db = A.Db;
            CivilDocument cdoc = A.Cdoc;

            var filter = new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
            });

            ed.WriteMessage("\n⊙ Chọn các Polyline để tạo tuyến (Alignment):");
            PromptSelectionResult selResult = ed.GetSelection(filter);
            if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
            {
                ed.WriteMessage("\n⚠️ Không chọn được Polyline nào. Hủy lệnh.");
                return;
            }

            string prefix;
            int startIndex;
            ObjectId alignStyleId, labelSetId, siteId;
            string layerName;

            using (var form = new AlignmentBatchForm(selResult.Value.Count))
            {
                var result = AcadApp.ShowModalDialog(form);
                if (result != DialogResult.OK || !form.IsConfirmed)
                {
                    ed.WriteMessage("\n⚠️ Đã hủy lệnh.");
                    return;
                }
                prefix = form.Prefix;
                startIndex = form.StartIndex;
                alignStyleId = form.GetSelectedStyleId();
                labelSetId = form.GetSelectedLabelSetId();
                siteId = form.GetSelectedSiteId();
                layerName = form.LayerName;
            }

            using Transaction tr = db.TransactionManager.StartTransaction();
            try
            {

                EnsureLayerExists(db, tr, layerName);

                List<ObjectId> sortedPolyIds = SortPolylinesByPosition(selResult.Value, tr);
                int successCount = 0, failCount = 0;
                int currentIndex = startIndex;

                foreach (ObjectId plId in sortedPolyIds)
                {
                    try
                    {
                        string alignName = GetUniqueAlignmentName(cdoc, tr, $"{prefix}-{currentIndex:D3}");
                        ObjectId newAlignId = Alignment.Create(cdoc, alignName, plId, siteId, alignStyleId, labelSetId);

                        if (newAlignId != ObjectId.Null)
                        {
                            Alignment? newAlign = tr.GetObject(newAlignId, OpenMode.ForWrite) as Alignment;
                            if (newAlign != null) newAlign.Layer = layerName;
                            successCount++;
                            ed.WriteMessage($"\n  ✅ [{currentIndex:D3}] Tạo tuyến: {alignName}");
                        }
                        else
                        {
                            failCount++;
                            ed.WriteMessage($"\n  ❌ [{currentIndex:D3}] Không thể tạo tuyến");
                        }
                        currentIndex++;
                    }
                    catch (System.Exception ex)
                    {
                        failCount++;
                        ed.WriteMessage($"\n  ❌ [{currentIndex:D3}] Lỗi: {ex.Message}");
                        currentIndex++;
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n\n✅ Hoàn thành! Thành công: {successCount}, Thất bại: {failCount}\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        [CommandMethod("CTA_TaoTuyenNhanh")]
        public static void CTA_TaoTuyenNhanh()
        {
            Editor ed = A.Ed;
            Database db = A.Db;
            CivilDocument cdoc = A.Cdoc;

            var filter = new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
            });

            ed.WriteMessage("\n⊙ TẠO TUYẾN NHANH - Chọn các Polyline:");
            PromptSelectionResult selResult = ed.GetSelection(filter);
            if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
            {
                ed.WriteMessage("\n⚠️ Không chọn được Polyline nào.");
                return;
            }

            string batchPrefix;
            int batchStartIndex;
            ObjectId batchStyleId, batchLabelSetId;
            string batchLayerName;

            using (var form = new AlignmentBatchForm(selResult.Value.Count))
            {
                var result = AcadApp.ShowModalDialog(form);
                if (result != DialogResult.OK || !form.IsConfirmed) return;
                batchPrefix = form.Prefix;
                batchStartIndex = form.StartIndex;
                batchStyleId = form.GetSelectedStyleId();
                batchLabelSetId = form.GetSelectedLabelSetId();
                batchLayerName = form.LayerName;
            }

            using Transaction tr = db.TransactionManager.StartTransaction();
            try
            {
                EnsureLayerExists(db, tr, batchLayerName);

                List<ObjectId> sortedPolyIds = SortPolylinesByPosition(selResult.Value, tr);
                int successCount = 0, index = batchStartIndex;

                foreach (ObjectId plId in sortedPolyIds)
                {
                    try
                    {
                        string alignName = GetUniqueAlignmentName(cdoc, tr, $"{batchPrefix}-{index:D3}");
                        ObjectId newAlignId = Alignment.Create(
                            cdoc, alignName, plId, ObjectId.Null,
                            batchStyleId, batchLabelSetId);

                        if (newAlignId != ObjectId.Null)
                        {
                            Alignment? newAlign = tr.GetObject(newAlignId, OpenMode.ForWrite) as Alignment;
                            if (newAlign != null) newAlign.Layer = batchLayerName;
                            successCount++;
                            ed.WriteMessage($"\n  ✅ Tạo tuyến: {alignName}");
                        }
                        index++;
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n  ❌ Lỗi Polyline #{index}: {ex.Message}");
                        index++;
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n\n✅ Hoàn thành! Đã tạo {successCount}/{selResult.Value.Count} tuyến.\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        [CommandMethod("CTA_DoiTenTuyenHangLoat")]
        public static void CTA_DoiTenTuyenHangLoat()
        {
            Editor ed = A.Ed;
            CivilDocument cdoc = A.Cdoc;

            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                ObjectIdCollection alignIds = cdoc.GetAlignmentIds();
                if (alignIds.Count == 0)
                {
                    ed.WriteMessage("\n⚠️ Không có Alignment nào trong bản vẽ.");
                    return;
                }

                var alignList = new List<(ObjectId Id, string Name)>();
                foreach (ObjectId aId in alignIds)
                {
                    Alignment? algn = tr.GetObject(aId, OpenMode.ForRead) as Alignment;
                    if (algn != null) alignList.Add((aId, algn.Name));
                }

                string newPrefix;
                using (var form = new RenameBatchForm(alignList))
                {
                    var result = AcadApp.ShowModalDialog(form);
                    if (result != DialogResult.OK || !form.IsConfirmed) return;
                    newPrefix = form.NewPrefix;
                }

                int count = 0;
                foreach (var (id, _) in alignList)
                {
                    Alignment? algn = tr.GetObject(id, OpenMode.ForWrite) as Alignment;
                    if (algn != null)
                    {
                        count++;
                        string oldName = algn.Name;
                        algn.Name = $"{newPrefix}-{count:D3}";
                        ed.WriteMessage($"\n  ✅ {oldName} → {algn.Name}");
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n\n✅ Đã đổi tên {count} tuyến.\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        // ═══ HÀM HỖ TRỢ ═══

        private static void EnsureLayerExists(Database db, Transaction tr, string layerName)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                LayerTableRecord ltr = new() { Name = layerName };
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }

        private static List<ObjectId> SortPolylinesByPosition(SelectionSet ss, Transaction tr)
        {
            var polyDataList = new List<(ObjectId Id, double X, double Y)>();
            foreach (SelectedObject so in ss)
            {
                if (so == null) continue;
                Polyline? pline = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Polyline;
                if (pline == null) continue;
                Point3d startPt = pline.StartPoint;
                polyDataList.Add((so.ObjectId, startPt.X, startPt.Y));
            }
            polyDataList.Sort((a, b) =>
            {
                int cmpX = a.X.CompareTo(b.X);
                return cmpX != 0 ? cmpX : b.Y.CompareTo(a.Y);
            });
            return polyDataList.Select(p => p.Id).ToList();
        }

        private static string GetUniqueAlignmentName(CivilDocument cdoc, Transaction tr, string baseName)
        {
            HashSet<string> existing = new(StringComparer.OrdinalIgnoreCase);
            foreach (ObjectId aId in cdoc.GetAlignmentIds())
            {
                Alignment? algn = tr.GetObject(aId, OpenMode.ForRead) as Alignment;
                if (algn != null) existing.Add(algn.Name);
            }
            if (!existing.Contains(baseName)) return baseName;
            int suffix = 1;
            while (existing.Contains($"{baseName}_{suffix}")) suffix++;
            return $"{baseName}_{suffix}";
        }
    }
}
