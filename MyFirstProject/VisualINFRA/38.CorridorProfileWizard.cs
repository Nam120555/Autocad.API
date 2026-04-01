// 38.CorridorProfileWizard.cs - Corridor & Profile Wizard
// Gộp: VI_CreateProfileView, VI_CreateMultiSurfaceProfile, VI_CreateOffsetProfile,
//       VI_CreateSurfaceFromCorridor, VI_RebuildAllCorridor, VI_CorridorInfo
// Form kế thừa CivilFormBase

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using ContentAlignment = System.Drawing.ContentAlignment;

[assembly: CommandClass(typeof(MyFirstProject.CorridorProfileWizardCommands))]

namespace MyFirstProject
{
    // ══════════════════════════════════════════════════════════════
    // FORM: CORRIDOR & PROFILE WIZARD
    // ══════════════════════════════════════════════════════════════

    public class CorridorProfileWizardForm : CivilFormBase
    {
        // ═══ Tab control ═══
        private TabControl tabControl = null!;
        private TabPage tabProfile = null!;
        private TabPage tabCorridor = null!;

        // Profile tab controls
        private RadioButton rdoCreatePV = null!;
        private RadioButton rdoCreateSurfProf = null!;
        private RadioButton rdoOffsetProf = null!;

        // Corridor tab controls
        private RadioButton rdoCreateSurfCorr = null!;
        private RadioButton rdoRebuildAll = null!;
        private RadioButton rdoCorridorInfo = null!;

        // Shared
        private ComboBox cboAlignment = null!;
        private ComboBox cboSurface = null!;
        private ComboBox cboCorridor = null!;

        // Data
        private List<(string Name, ObjectId Id)> alignmentList = new();
        private List<(string Name, ObjectId Id)> surfaceList = new();
        private List<(string Name, ObjectId Id)> corridorList = new();

        public bool IsConfirmed { get; private set; }
        public string SelectedMode
        {
            get
            {
                if (rdoCreatePV.Checked) return "CreateProfileView";
                if (rdoCreateSurfProf.Checked) return "CreateMultiSurfaceProfile";
                if (rdoOffsetProf.Checked) return "OffsetProfile";
                if (rdoCreateSurfCorr.Checked) return "CreateSurfaceFromCorridor";
                if (rdoRebuildAll.Checked) return "RebuildAllCorridor";
                return "CorridorInfo";
            }
        }

        public CorridorProfileWizardForm()
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

            // Alignments
            foreach (ObjectId id in civilDoc.GetAlignmentIds())
            {
                var al = tr.GetObject(id, OpenMode.ForRead) as Alignment;
                if (al != null) alignmentList.Add((al.Name, id));
            }

            // Surfaces
            foreach (ObjectId id in civilDoc.GetSurfaceIds())
            {
                var s = tr.GetObject(id, OpenMode.ForRead) as TinSurface;
                if (s != null) surfaceList.Add((s.Name, id));
            }

            // Corridors
            foreach (ObjectId id in civilDoc.CorridorCollection)
            {
                var c = tr.GetObject(id, OpenMode.ForRead) as Corridor;
                if (c != null) corridorList.Add((c.Name, id));
            }

            tr.Commit();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("CORRIDOR & PROFILE WIZARD", 470, 520);

            // ═══ HEADER ═══
            var header = CreateHeader("CORRIDOR & PROFILE WIZARD");
            this.Controls.Add(header);

            var subtitle = CreateSubtitle("Gộp Profile View, Surface Profile, Corridor Tools");
            this.Controls.Add(subtitle);

            // ═══ TAB CONTROL ═══
            tabControl = new TabControl
            {
                Location = new Point(16, 78),
                Size = new Size(428, 340),
                Font = new Font("Segoe UI", 10f)
            };
            this.Controls.Add(tabControl);

            // ─── TAB PROFILE ───
            tabProfile = new TabPage("Profile");
            tabProfile.BackColor = BgPanel;
            tabProfile.Padding = new Padding(10);
            tabControl.TabPages.Add(tabProfile);

            int py = 12;
            tabProfile.Controls.Add(CreateLabel("Alignment:", 10, py + 3));
            cboAlignment = CreateComboBox(110, py, 290);
            foreach (var (name, _) in alignmentList)
                cboAlignment.Items.Add(name);
            if (cboAlignment.Items.Count > 0) cboAlignment.SelectedIndex = 0;
            tabProfile.Controls.Add(cboAlignment);
            py += 38;

            tabProfile.Controls.Add(CreateLabel("Surface:", 10, py + 3));
            cboSurface = CreateComboBox(110, py, 290);
            foreach (var (name, _) in surfaceList)
                cboSurface.Items.Add(name);
            if (cboSurface.Items.Count > 0) cboSurface.SelectedIndex = 0;
            tabProfile.Controls.Add(cboSurface);
            py += 50;

            // Chức năng Profile
            var grpProfile = CreateGroupBox("Chức năng Profile", 8, py, 390, 120);
            tabProfile.Controls.Add(grpProfile);

            int ry = 26;
            rdoCreatePV = new RadioButton
            {
                Text = "Tạo Profile View (chọn điểm chèn)",
                Location = new Point(12, ry), AutoSize = true, Checked = true,
                ForeColor = TextDark, Font = new Font("Segoe UI", 9.5f)
            };
            grpProfile.Controls.Add(rdoCreatePV);
            ry += 28;

            rdoCreateSurfProf = new RadioButton
            {
                Text = "Tạo Surface Profile cho TẤT CẢ Alignment",
                Location = new Point(12, ry), AutoSize = true,
                ForeColor = TextDark, Font = new Font("Segoe UI", 9.5f)
            };
            grpProfile.Controls.Add(rdoCreateSurfProf);
            ry += 28;

            rdoOffsetProf = new RadioButton
            {
                Text = "Tạo Offset Profile (chọn Profile View)",
                Location = new Point(12, ry), AutoSize = true,
                ForeColor = TextDark, Font = new Font("Segoe UI", 9.5f)
            };
            grpProfile.Controls.Add(rdoOffsetProf);

            // ─── TAB CORRIDOR ───
            tabCorridor = new TabPage("Corridor");
            tabCorridor.BackColor = BgPanel;
            tabCorridor.Padding = new Padding(10);
            tabControl.TabPages.Add(tabCorridor);

            int cy = 12;
            tabCorridor.Controls.Add(CreateLabel("Corridor:", 10, cy + 3));
            cboCorridor = CreateComboBox(110, cy, 290);
            foreach (var (name, _) in corridorList)
                cboCorridor.Items.Add(name);
            if (cboCorridor.Items.Count > 0) cboCorridor.SelectedIndex = 0;
            tabCorridor.Controls.Add(cboCorridor);
            cy += 50;

            var grpCorridor = CreateGroupBox("Chức năng Corridor", 8, cy, 390, 120);
            tabCorridor.Controls.Add(grpCorridor);

            ry = 26;
            rdoCreateSurfCorr = new RadioButton
            {
                Text = "Tạo Surface từ Corridor",
                Location = new Point(12, ry), AutoSize = true, Checked = true,
                ForeColor = TextDark, Font = new Font("Segoe UI", 9.5f)
            };
            grpCorridor.Controls.Add(rdoCreateSurfCorr);
            ry += 28;

            rdoRebuildAll = new RadioButton
            {
                Text = "Rebuild TẤT CẢ Corridor",
                Location = new Point(12, ry), AutoSize = true,
                ForeColor = TextDark, Font = new Font("Segoe UI", 9.5f)
            };
            grpCorridor.Controls.Add(rdoRebuildAll);
            ry += 28;

            rdoCorridorInfo = new RadioButton
            {
                Text = "Xem thông tin Corridor",
                Location = new Point(12, ry), AutoSize = true,
                ForeColor = TextDark, Font = new Font("Segoe UI", 9.5f)
            };
            grpCorridor.Controls.Add(rdoCorridorInfo);

            // Thống kê
            var statsPanel = new Panel
            {
                Location = new Point(8, cy + 130),
                Size = new Size(390, 60),
                BackColor = BgWhite,
                BorderStyle = BorderStyle.FixedSingle
            };
            tabCorridor.Controls.Add(statsPanel);

            var lblStats = new WinLabel
            {
                Text = $"Alignments: {alignmentList.Count}  |  Surfaces: {surfaceList.Count}  |  Corridors: {corridorList.Count}",
                Location = new Point(10, 5),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            statsPanel.Controls.Add(lblStats);

            var lblHint = new WinLabel
            {
                Text = "Chọn tab → chức năng → bấm Thực Hiện",
                Location = new Point(10, 30),
                AutoSize = true,
                ForeColor = TextHint,
                Font = new Font("Segoe UI", 8.5f)
            };
            statsPanel.Controls.Add(lblHint);

            // ═══ BOTTOM BAR ═══
            var bottomBar = CreateBottomBar(56);
            this.Controls.Add(bottomBar);

            var btnCancel = CreateButton("Hủy", 10, 10, 200, 36, bottomBar);
            btnCancel.Click += (s, e) => { IsConfirmed = false; this.Close(); };

            var btnRun = CreateSuccessButton("Thực Hiện", 220, 10, 210, 36, bottomBar);
            btnRun.Click += (s, e) => { IsConfirmed = true; this.Close(); };

            this.AcceptButton = btnRun;
            this.CancelButton = btnCancel;
        }
    }

    // ══════════════════════════════════════════════════════════════
    // LỆNH CORRIDOR PROFILE WIZARD
    // ══════════════════════════════════════════════════════════════

    public class CorridorProfileWizardCommands
    {
        /// <summary>
        /// Mở Corridor & Profile Wizard — gộp 6 lệnh VI_Create*/VI_Corridor*
        /// </summary>
        [CommandMethod("VI_CorridorWizard")]
        public static void VI_CorridorWizard()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            using var form = new CorridorProfileWizardForm();
            AcadApp.ShowModalDialog(form);
            if (!form.IsConfirmed) return;

            switch (form.SelectedMode)
            {
                case "CreateProfileView":
                    VisualINFRA.ProfileCorridorTools.CreateProfileView();
                    break;
                case "CreateMultiSurfaceProfile":
                    VisualINFRA.ProfileCorridorTools.CreateMultiSurfaceProfile();
                    break;
                case "OffsetProfile":
                    VisualINFRA.ProfileCorridorTools.CreateOffsetProfile();
                    break;
                case "CreateSurfaceFromCorridor":
                    VisualINFRA.ProfileCorridorTools.CreateSurfaceFromCorridor();
                    break;
                case "RebuildAllCorridor":
                    VisualINFRA.ProfileCorridorTools.RebuildAllCorridor();
                    break;
                case "CorridorInfo":
                    VisualINFRA.ProfileCorridorTools.CorridorInfo();
                    break;
            }
        }
    }
}
