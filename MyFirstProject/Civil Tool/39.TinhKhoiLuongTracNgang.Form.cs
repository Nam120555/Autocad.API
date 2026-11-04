// 39.TinhKhoiLuongTracNgang.cs - Lệnh CTSV_TinhKhoiLuongTracNgang
// Tổng hợp: Tính diện tích Shape (đào/đắp/vật liệu) từ Material Section
//            rồi tính khối lượng theo 2 phương pháp:
//            - Average End Area: V = (S₁ + S₂) / 2 × L
//            - Prismoidal:       V = (S₁ + 4Sm + S₂) / 6 × L
// Nguồn diện tích: section.Area (API) hoặc Shoelace (SectionPoints)
// Xuất ra: Command Line + Clipboard + CSV (tùy chọn)

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using MyFirstProject.Extensions;

using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using ContentAlignment = System.Drawing.ContentAlignment;
using Section = Autodesk.Civil.DatabaseServices.Section;

namespace MyFirstProject
{
    public class TinhKhoiLuongForm : CivilFormBase
    {
        // Data
        private List<(string Name, ObjectId AlignmentId, ObjectId SLGId)> alignmentList = new();
        private List<(string MaterialName, string ListName)> materialList = new();

        // Controls
        private CheckedListBox chkAlignments = null!;
        private CheckedListBox chkMaterials = null!;
        private RadioButton rdoArea_API = null!;
        private RadioButton rdoArea_Shoelace = null!;
        private RadioButton rdoArea_Fallback = null!;
        private RadioButton rdoVol_AEA = null!;
        private RadioButton rdoVol_Prismoidal = null!;
        private CheckBox chkCompare = null!;
        private CheckBox chkExportCSV = null!;
        private CheckBox chkExportExcel = null!;
        private CheckBox chkWriteToDrawing = null!;
        private RadioButton rdoSource_ComputeMat = null!;
        private RadioButton rdoSource_HatchExp = null!;

        // Kết quả
        public bool IsConfirmed { get; private set; }
        
        public List<(string Name, ObjectId AlignmentId, ObjectId SLGId)> SelectedAlignments
        {
            get
            {
                var result = new List<(string Name, ObjectId AlignmentId, ObjectId SLGId)>();
                for (int i = 0; i < chkAlignments.Items.Count; i++)
                {
                    if (chkAlignments.GetItemChecked(i))
                        result.Add(alignmentList[i]);
                }
                return result;
            }
        }

        public List<(string MaterialName, string ListName)> SelectedMaterials
        {
            get
            {
                var result = new List<(string, string)>();
                for (int i = 0; i < chkMaterials.Items.Count; i++)
                {
                    if (chkMaterials.GetItemChecked(i))
                        result.Add(materialList[i]);
                }
                return result;
            }
        }
        public DataSourceMode SelectedDataSource =>
            rdoSource_HatchExp.Checked ? DataSourceMode.HatchExplosion : DataSourceMode.ComputeMaterials;
        public AreaSource SelectedAreaSource =>
            rdoArea_API.Checked ? AreaSource.API :
            rdoArea_Shoelace.Checked ? AreaSource.Shoelace :
            AreaSource.APIWithFallback;
        public VolumeMethod SelectedVolumeMethod =>
            rdoVol_Prismoidal.Checked ? VolumeMethod.Prismoidal : VolumeMethod.AverageEndArea;
        public bool CompareMethods => chkCompare.Checked;
        public bool ExportCSV => chkExportCSV.Checked;
        public bool ExportExcel => chkExportExcel.Checked;
        public bool WriteToDrawing => chkWriteToDrawing.Checked;

        public TinhKhoiLuongForm()
        {
            LoadData();
            InitializeComponent();
        }

        private void LoadData()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var civilDoc = CivilApplication.ActiveDocument;

            using var tr = db.TransactionManager.StartOpenCloseTransaction();

            foreach (ObjectId alId in civilDoc.GetAlignmentIds())
            {
                var al = tr.GetObject(alId, OpenMode.ForRead) as Alignment;
                if (al == null) continue;

                var slgIds = al.GetSampleLineGroupIds();
                foreach (ObjectId slgId in slgIds)
                {
                    var slg = tr.GetObject(slgId, OpenMode.ForRead) as SampleLineGroup;
                    if (slg != null)
                        alignmentList.Add(($"{al.Name} — {slg.Name}", alId, slgId));
                }
            }
            tr.Commit();
        }

        private void LoadMaterials()
        {
            materialList.Clear();
            chkMaterials.Items.Clear();

            var checkedIndices = chkAlignments.CheckedIndices;
            if (checkedIndices.Count == 0) return;

            var uniqueMaterials = new HashSet<string>();

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using var tr = doc.Database.TransactionManager.StartOpenCloseTransaction();

            foreach (int idx in checkedIndices)
            {
                var slgId = alignmentList[idx].SLGId;
                var slg = tr.GetObject(slgId, OpenMode.ForRead) as SampleLineGroup;
                if (slg == null) continue;

                if (rdoSource_HatchExp.Checked)
                {
                    // Móc trực tiếp Hatch Layer từ Trắc ngang đồ họa
                    var slIds = slg.GetSampleLineIds();
                    var uniqueLayers = new HashSet<string>();

                    int scanCount = Math.Min(5, slIds.Count);
                    for (int i = 0; i < scanCount; i++)
                    {
                        var sl = tr.GetObject(slIds[i], OpenMode.ForRead) as SampleLine;
                        if (sl == null) continue;

                        var secIds = sl.GetSectionIds();
                        foreach (ObjectId secId in secIds)
                        {
                            var sec = tr.GetObject(secId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Section;
                            if (sec == null) continue;

                            try
                            {
                                var objs = new Autodesk.AutoCAD.DatabaseServices.DBObjectCollection();
                                sec.Explode(objs);

                                foreach (Autodesk.AutoCAD.DatabaseServices.DBObject obj in objs)
                                {
                                    if (obj is Hatch h)
                                    {
                                        uniqueLayers.Add(h.Layer);
                                    }
                                    else if (obj is BlockReference br)
                                    {
                                        var subObjs = new Autodesk.AutoCAD.DatabaseServices.DBObjectCollection();
                                        try { br.Explode(subObjs); } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                                        foreach (Autodesk.AutoCAD.DatabaseServices.DBObject subObj in subObjs)
                                        {
                                            if (subObj is Hatch subH) uniqueLayers.Add(subH.Layer);
                                            subObj.Dispose();
                                        }
                                    }
                                    obj.Dispose();
                                }
                            }
                            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                        }
                    }

                    foreach (var layer in uniqueLayers)
                    {
                        string key = layer;
                        if (!uniqueMaterials.Contains(key))
                        {
                            uniqueMaterials.Add(key);
                            materialList.Add((layer, "Hatch_Layer"));
                            int chkIdx = chkMaterials.Items.Add($"Layer: {layer}");
                            chkMaterials.SetItemChecked(chkIdx, true);
                        }
                    }
                }
                else
                {
                    // Bắt nguồn từ Compute Materials
                    foreach (QTOMaterialList matList in slg.MaterialLists)
                    {
                        try
                        {
                            foreach (QTOMaterial mat in matList)
                            {
                                string key = $"{mat.Name}|{matList.Name}";
                                if (!uniqueMaterials.Contains(key))
                                {
                                    uniqueMaterials.Add(key);
                                    string type = mat.QuantityType == MaterialQuantityType.Cut ? "Đào" :
                                                  mat.QuantityType == MaterialQuantityType.Fill ? "Đắp" :
                                                  mat.QuantityType == MaterialQuantityType.CutFill ? "Đào+Đắp" :
                                                  mat.QuantityType == MaterialQuantityType.Earthwork ? "Earthwork" :
                                                  mat.QuantityType == MaterialQuantityType.Structure ? "Kết cấu" : "Khác";
                                    materialList.Add((mat.Name, matList.Name ?? "Unnamed"));
                                    int chkIdx = chkMaterials.Items.Add($"{type}: {mat.Name} ({matList.Name})");
                                    chkMaterials.SetItemChecked(chkIdx, true);
                                }
                            }
                        }
                        catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                    }
                }
            }

            if (materialList.Count == 0 && !rdoSource_HatchExp.Checked)
            {
                chkMaterials.Items.Add("Chưa có Material List trên các tuyến đã chọn.");
            }
            tr.Commit();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("TÍNH KHỐI LƯỢNG TRẮC NGANG", 500, 890);

            int y = 0;

            // ═══ HEADER ═══
            var header = CreateHeader("TÍNH KHỐI LƯỢNG TRẮC NGANG");
            this.Controls.Add(header);

            var subtitle = CreateSubtitle("Lấy diện tích Material Section → Tính KL theo Average End Area / Prismoidal");
            this.Controls.Add(subtitle);

            y = 82;

            // ═══ CHỌN ALIGNMENT ═══
            var grpAlign = CreateGroupBox("Chọn các Alignment + SampleLine Group", 16, y, 460, 120);
            this.Controls.Add(grpAlign);

            chkAlignments = new CheckedListBox
            {
                Location = new Point(12, 26),
                Size = new Size(436, 80),
                CheckOnClick = true,
                Font = new Font("Segoe UI", 9.5f),
                BackColor = BgWhite,
                ForeColor = TextDark,
                BorderStyle = BorderStyle.FixedSingle
            };
            foreach (var (name, _, _) in alignmentList)
                chkAlignments.Items.Add(name);
            if (chkAlignments.Items.Count > 0)
                chkAlignments.SetItemChecked(0, true);
            chkAlignments.ItemCheck += (s, e) => this.BeginInvoke(new Action(() => LoadMaterials()));
            grpAlign.Controls.Add(chkAlignments);

            y += 128;

            // ═══ NGUỒN DỮ LIỆU ĐẦU VÀO ═══
            var grpDataSource = CreateGroupBox("Nguồn Dữ Liệu Khối Lượng", 16, y, 460, 75);
            this.Controls.Add(grpDataSource);

            rdoSource_ComputeMat = new RadioButton
            {
                Text = "Từ Compute Materials (QTOMaterialList)",
                Location = new Point(12, 24),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            rdoSource_ComputeMat.CheckedChanged += (s, e) => { if (rdoSource_ComputeMat.Checked) LoadMaterials(); };
            grpDataSource.Controls.Add(rdoSource_ComputeMat);

            rdoSource_HatchExp = new RadioButton
            {
                Text = "Từ Hatch Đồ Họa (Rã ảo Trắc ngang để đo diện tích)",
                Location = new Point(12, 48),
                AutoSize = true,
                ForeColor = Color.DarkRed,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            rdoSource_HatchExp.CheckedChanged += (s, e) => { if (rdoSource_HatchExp.Checked) LoadMaterials(); };
            grpDataSource.Controls.Add(rdoSource_HatchExp);

            y += 85;

            // ═══ CHỌN MATERIAL ═══
            var grpMat = CreateGroupBox("Chọn Material / Category cần tính", 16, y, 460, 140);
            this.Controls.Add(grpMat);

            chkMaterials = new CheckedListBox
            {
                Location = new Point(12, 26),
                Size = new Size(436, 100),
                CheckOnClick = true,
                Font = new Font("Segoe UI", 9.5f),
                BackColor = BgWhite,
                ForeColor = TextDark,
                BorderStyle = BorderStyle.FixedSingle
            };
            grpMat.Controls.Add(chkMaterials);

            y += 148;

            // ═══ PHƯƠNG PHÁP TÍNH DIỆN TÍCH ═══
            var grpArea = CreateGroupBox("Nguồn diện tích mặt cắt", 16, y, 460, 90);
            this.Controls.Add(grpArea);

            rdoArea_Fallback = new RadioButton
            {
                Text = "section.Area (API) — fallback Shoelace",
                Location = new Point(12, 26),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9f)
            };
            grpArea.Controls.Add(rdoArea_Fallback);

            rdoArea_API = new RadioButton
            {
                Text = "section.Area (API) — chỉ dùng API",
                Location = new Point(12, 48),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9f)
            };
            grpArea.Controls.Add(rdoArea_API);

            rdoArea_Shoelace = new RadioButton
            {
                Text = "Shoelace (tự tính từ SectionPoints)",
                Location = new Point(12, 68),
                AutoSize = true,
                ForeColor = TextHint,
                Font = new Font("Segoe UI", 9f)
            };
            grpArea.Controls.Add(rdoArea_Shoelace);

            y += 98;

            // ═══ PHƯƠNG PHÁP TÍNH KHỐI LƯỢNG ═══
            var grpVol = CreateGroupBox("Phương pháp tính khối lượng", 16, y, 460, 108);
            this.Controls.Add(grpVol);

            rdoVol_AEA = new RadioButton
            {
                Text = "Average End Area:  V = (S₁ + S₂) / 2 × L",
                Location = new Point(12, 26),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpVol.Controls.Add(rdoVol_AEA);

            rdoVol_Prismoidal = new RadioButton
            {
                Text = "Prismoidal:  V = (S₁ + 4Sm + S₂) / 6 × L",
                Location = new Point(12, 50),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpVol.Controls.Add(rdoVol_Prismoidal);

            chkCompare = new CheckBox
            {
                Text = "So sánh cả 2 phương pháp (hiển thị chênh lệch %)",
                Location = new Point(12, 78),
                AutoSize = true,
                Checked = false,
                ForeColor = NavyDark,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            grpVol.Controls.Add(chkCompare);

            y += 116;

            // ═══ TÙY CHỌN XUẤT ═══
            var grpExport = CreateGroupBox("Tùy chọn xuất", 16, y, 460, 85);
            this.Controls.Add(grpExport);

            chkExportCSV = new CheckBox
            {
                Text = "Xuất CSV",
                Location = new Point(12, 26),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpExport.Controls.Add(chkExportCSV);

            chkExportExcel = new CheckBox
            {
                Text = "Xuất Excel (.xlsx)",
                Location = new Point(160, 26),
                AutoSize = true,
                Checked = false,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpExport.Controls.Add(chkExportExcel);

            chkWriteToDrawing = new CheckBox
            {
                Text = "Ghi text lên trắc ngang",
                Location = new Point(12, 52),
                AutoSize = true,
                Checked = false,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpExport.Controls.Add(chkWriteToDrawing);

            y += 93;

            // ═══ BOTTOM BAR ═══
            var bottomBar = CreateBottomBar(56);
            this.Controls.Add(bottomBar);

            var btnCancel = CreateButton("Hủy", 10, 10, 220, 36, bottomBar);
            btnCancel.Click += (s, e) => { IsConfirmed = false; this.Close(); };

            var btnRun = CreateSuccessButton("Tính Khối Lượng", 240, 10, 220, 36, bottomBar);
            btnRun.Click += (s, e) =>
            {
                if (chkAlignments.CheckedItems.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show("Vui lòng chọn ít nhất 1 Alignment!", "Thiếu dữ liệu",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (SelectedMaterials.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show("Vui lòng chọn ít nhất 1 Material!", "Thiếu dữ liệu",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                IsConfirmed = true;
                this.Close();
            };

            this.AcceptButton = btnRun;
            this.CancelButton = btnCancel;

            // Load materials cho alignment đầu tiên
            if (chkAlignments.Items.Count > 0)
                LoadMaterials();
        }
    }
}
