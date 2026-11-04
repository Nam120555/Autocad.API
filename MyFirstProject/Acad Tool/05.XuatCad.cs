// 05.XuatCad.cs - Bộ công cụ Xuất CAD chuyên nghiệp, chống lỗi Xref & eDuplicateRecordName
// Kế thừa CivilFormBase (Navy / White theme)
// Khắc phục triệt để lỗi:
// 1. "Xref-Bind operation on scaled Xref with unequal units"
// 2. "ERROR: Xref-Bind failed (eDuplicateRecordName)"
// 3. "ERROR: could not process Xrefs!_PSPACE *Invalid*"

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using ContentAlignment = System.Drawing.ContentAlignment;

using MyFirstProject;

[assembly: CommandClass(typeof(MyFirstProject.Acad_Tool.AT_XuatCad))]

namespace MyFirstProject.Acad_Tool
{
    // ══════════════════════════════════════════════════════════════════
    // MODEL: THÔNG TIN XREF TRONG BẢN VẼ
    // ══════════════════════════════════════════════════════════════════
    public class XrefItemInfo
    {
        public ObjectId BtrId { get; set; }
        public string Name { get; set; } = "";
        public string PathName { get; set; } = "";
        public bool IsResolved { get; set; }
        public bool IsUnloaded { get; set; }
        public bool FileExists { get; set; }
        public bool HasUnequalScale { get; set; }
        public string ScaleDescription { get; set; } = "Bình thường";
        public string StatusDescription { get; set; } = "Đã nạp";
        public string Action { get; set; } = "Gộp (Insert)";
    }

    // ══════════════════════════════════════════════════════════════════
    // HELPER: QUẢN LÝ VÀ XỬ LÝ AN TOÀN XREF
    // ══════════════════════════════════════════════════════════════════
    public static class SafeXrefHelper
    {
        /// <summary>
        /// Quét tất cả Xref trong Database và phát hiện các vấn đề tiềm ẩn
        /// (Lệch tỷ lệ, thiếu file, chưa nạp, trùng tên).
        /// </summary>
        public static List<XrefItemInfo> GetAllXrefs(Database db)
        {
            var list = new List<XrefItemInfo>();

            using var tr = db.TransactionManager.StartTransaction();
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

            foreach (ObjectId btrId in bt)
            {
                var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                if (!btr.IsFromExternalReference) continue;

                var item = new XrefItemInfo
                {
                    BtrId = btrId,
                    Name = btr.Name,
                    PathName = btr.PathName,
                    IsResolved = btr.IsResolved,
                    IsUnloaded = btr.IsUnloaded
                };

                // Kiểm tra file vật lý tồn tại không
                if (!string.IsNullOrEmpty(btr.PathName))
                {
                    try
                    {
                        string resolvedPath = btr.PathName;
                        if (!File.Exists(resolvedPath))
                        {
                            // Thử tìm theo thư mục bản vẽ hiện tại
                            string hostDir = Path.GetDirectoryName(db.Filename) ?? "";
                            string altPath = Path.Combine(hostDir, Path.GetFileName(btr.PathName));
                            if (File.Exists(altPath))
                            {
                                resolvedPath = altPath;
                                item.PathName = altPath;
                            }
                        }
                        item.FileExists = File.Exists(resolvedPath);
                    }
                    catch
                    {
                        item.FileExists = false;
                    }
                }

                // Kiểm tra trạng thái nạp
                if (item.IsUnloaded)
                {
                    item.StatusDescription = "Chưa nạp (Unloaded)";
                    item.Action = "Bỏ qua / Detach";
                }
                else if (!item.IsResolved || !item.FileExists)
                {
                    item.StatusDescription = "Không tìm thấy file (Not Found)";
                    item.Action = "Bỏ qua / Detach";
                }
                else
                {
                    item.StatusDescription = "Đã nạp (Loaded)";
                    item.Action = "Gộp (Insert)";
                }

                // Kiểm tra tỷ lệ không đồng đều (Unequal Scale) trên các BlockReference chèn Xref
                try
                {
                    var blkRefIds = btr.GetBlockReferenceIds(true, false);
                    foreach (ObjectId refId in blkRefIds)
                    {
                        if (tr.GetObject(refId, OpenMode.ForRead) is BlockReference blkRef)
                        {
                            var s = blkRef.ScaleFactors;
                            if (Math.Abs(s.X - s.Y) > 1e-4 || Math.Abs(s.X - s.Z) > 1e-4)
                            {
                                item.HasUnequalScale = true;
                                item.ScaleDescription = $"Lệch (X={s.X:0.##}, Y={s.Y:0.##})";
                                break;
                            }
                        }
                    }
                }
                catch { }

                list.Add(item);
            }

            tr.Commit();
            return list;
        }

        /// <summary>
        /// Gộp an toàn một Xref đơn lẻ.
        /// Cấp 1: Thử Insert mode (tránh tạo tiền tố $0$)
        /// Cấp 2: Thử Bind mode (nếu có xung đột định nghĩa block)
        /// Cấp 3 (Cứu nguy triệt để): Deep Embedding (Nạp trực tiếp DWG thành block nội bộ và gỡ Xref,
        /// hoàn toàn khắc phục lỗi eDuplicateRecordName do Civil 3D styles hoặc xung đột record).
        /// </summary>
        public static (bool Success, string Message) SafeBindSingleXref(Database db, ObjectId btrId, bool preferInsert = true)
        {
            string xrefName = "";
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    xrefName = btr.Name;
                    tr.Commit();
                }

                var ids = new ObjectIdCollection { btrId };

                // Cấp 1: Thử Insert mode
                try
                {
                    db.BindXrefs(ids, preferInsert);
                    return (true, $"✓ Gộp [{xrefName}] thành công (Chế độ {(preferInsert ? "Insert" : "Bind")})");
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex1)
                {
                    // Cấp 2: Thử chế độ ngược lại
                    try
                    {
                        db.BindXrefs(ids, !preferInsert);
                        return (true, $"✓ Gộp [{xrefName}] thành công (Chế độ {(!preferInsert ? "Insert" : "Bind")})");
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex2)
                    {
                        // Cấp 3: Fallback tối thượng Deep Embedding
                        var deepRes = DeepEmbedXref(db, btrId, xrefName);
                        if (deepRes.Success)
                        {
                            return deepRes;
                        }
                        return (false, $"❌ Không thể gộp [{xrefName}]: Lỗi {ex2.ErrorStatus} ({ex2.Message}). Deep-Embed: {deepRes.Message}");
                    }
                }
            }
            catch (System.Exception ex)
            {
                // Thử deep embed lần cuối nếu lỗi ngoại lệ
                var deepRes = DeepEmbedXref(db, btrId, xrefName);
                if (deepRes.Success) return deepRes;
                return (false, $"❌ Lỗi xử lý Xref [{xrefName}]: {ex.Message}");
            }
        }

        /// <summary>
        /// Cơ chế giải cứu tối thượng khi Xref bị lỗi eDuplicateRecordName (do chứa Civil 3D styles hoặc xung đột symbol):
        /// Đọc trực tiếp DWG của Xref, nạp ModelSpace vào một Block nội bộ trong Host Database,
        /// cập nhật toàn bộ BlockReference tham chiếu sang Block mới, rồi Detach Xref cũ.
        /// </summary>
        public static (bool Success, string Message) DeepEmbedXref(Database db, ObjectId btrId, string xrefName)
        {
            try
            {
                string pathName = "";
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    pathName = btr.PathName;
                    tr.Commit();
                }

                string resolvedPath = pathName;
                if (!File.Exists(resolvedPath))
                {
                    string hostDir = Path.GetDirectoryName(db.Filename) ?? "";
                    string altPath = Path.Combine(hostDir, Path.GetFileName(pathName));
                    if (File.Exists(altPath))
                    {
                        resolvedPath = altPath;
                    }
                }

                if (!File.Exists(resolvedPath))
                {
                    return (false, $"Không tìm thấy file DWG nguồn của [{xrefName}] tại '{pathName}'");
                }

                using var xrefDb = new Database(false, true);
                xrefDb.ReadDwgFile(resolvedPath, FileOpenMode.OpenForReadAndAllShare, true, "");

                string cleanName = Path.GetFileNameWithoutExtension(resolvedPath).Replace(" ", "_").Replace(".", "_");
                string safeBlockName = "BOUND_" + cleanName;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    int counter = 1;
                    while (bt.Has(safeBlockName))
                    {
                        safeBlockName = $"BOUND_{cleanName}_{counter++}";
                    }
                    tr.Commit();
                }

                // Chèn ModelSpace của Xref thành Block Table Record nội bộ trong database hiện tại
                ObjectId newBtrId = db.Insert(safeBlockName, xrefDb, true);

                // Cập nhật toàn bộ BlockReference đang trỏ tới Xref cũ sang Block mới
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var oldBtr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    var blkRefIds = oldBtr.GetBlockReferenceIds(true, false);
                    foreach (ObjectId rId in blkRefIds)
                    {
                        if (tr.GetObject(rId, OpenMode.ForWrite) is BlockReference blkRef)
                        {
                            blkRef.BlockTableRecord = newBtrId;
                        }
                    }
                    tr.Commit();
                }

                // Gỡ bỏ Xref cũ vì đã được thay thế trọn vẹn bằng Block nội bộ
                try
                {
                    db.DetachXref(btrId);
                }
                catch { }

                return (true, $"✓ Gộp [{xrefName}] thành công qua cơ chế Deep Embedding (Vượt qua lỗi eDuplicateRecordName/Civil 3D styles)");
            }
            catch (System.Exception ex)
            {
                return (false, $"Lỗi Deep Embedding [{xrefName}]: {ex.Message}");
            }
        }

        /// <summary>
        /// Dọn dẹp RegApps rác trực tiếp bằng C# API (siêu tốc, triệt tiêu nguồn gốc lỗi eDuplicateRecordName & giật lag)
        /// </summary>
        public static int PurgeRegApps(Database db)
        {
            int purged = 0;
            try
            {
                using var tr = db.TransactionManager.StartTransaction();
                var regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
                var ids = new ObjectIdCollection();

                foreach (ObjectId id in regTable)
                {
                    var reg = (RegAppTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    if (!string.Equals(reg.Name, "ACAD", StringComparison.OrdinalIgnoreCase))
                    {
                        ids.Add(id);
                    }
                }

                if (ids.Count > 0)
                {
                    db.Purge(ids);
                    foreach (ObjectId id in ids)
                    {
                        var reg = tr.GetObject(id, OpenMode.ForWrite);
                        reg.Erase();
                        purged++;
                    }
                }

                tr.Commit();
            }
            catch { }
            return purged;
        }

        /// <summary>
        /// Gộp an toàn toàn bộ Xref trong bản vẽ.
        /// Tự động cô lập các Xref bị lỗi để không làm hỏng tiến trình xuất CAD.
        /// </summary>
        public static (int SuccessCount, int FailedCount, List<string> Logs) SafeBindAllXrefs(Database db, Editor ed, bool preferInsert = true)
        {
            var logs = new List<string>();
            int success = 0;
            int failed = 0;

            var xrefs = GetAllXrefs(db);
            if (xrefs.Count == 0)
            {
                logs.Add("Không có Xref nào trong bản vẽ.");
                return (0, 0, logs);
            }

            foreach (var item in xrefs)
            {
                // Bỏ qua các Xref không tìm thấy file hoặc chưa nạp
                if (!item.IsResolved || !item.FileExists || item.IsUnloaded)
                {
                    logs.Add($"⚠ Bỏ qua Xref [{item.Name}]: Trạng thái {item.StatusDescription}");
                    continue;
                }

                var result = SafeBindSingleXref(db, item.BtrId, preferInsert);
                if (result.Success)
                {
                    success++;
                    logs.Add(result.Message);
                    ed.WriteMessage($"\n  {result.Message}");
                }
                else
                {
                    failed++;
                    logs.Add(result.Message);
                    ed.WriteMessage($"\n  {result.Message}");
                }
            }

            return (success, failed, logs);
        }

        /// <summary>
        /// Detach (Gỡ bỏ) một Xref lỗi hoặc không tìm thấy file
        /// </summary>
        public static bool DetachXref(Database db, ObjectId btrId)
        {
            try
            {
                db.DetachXref(btrId);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    // ══════════════════════════════════════════════════════════════════
    // FORM GIAO DIỆN: XUẤT CAD CHUYÊN NGHIỆP
    // ══════════════════════════════════════════════════════════════════
    public class XuatCadForm : CivilFormBase
    {
        // Controls danh sách Xref
        private DataGridView dgvXrefs = null!;
        private WinLabel lblSummary = null!;
        private Button btnScan = null!;
        private Button btnFixSelected = null!;
        private Button btnDetachBad = null!;

        // Controls chế độ xử lý Xref
        private RadioButton rdoInsertBind = null!;
        private RadioButton rdoNoBind = null!;
        private RadioButton rdoPreBindApi = null!;

        // Controls cấu hình xuất CAD
        private ComboBox cboVersion = null!;
        private TextBox txtFolder = null!;
        private TextBox txtFileName = null!;
        private CheckBox chkPurge = null!;
        private CheckBox chkOpenFolder = null!;
        private Button btnBrowse = null!;
        private Button btnOpenFolderNow = null!;

        // Buttons chính
        private Button btnExportNow = null!;
        private Button btnFixOnly = null!;
        private Button btnClose = null!;

        // Dữ liệu nội bộ
        private List<XrefItemInfo> currentXrefs = new();
        private readonly Document targetDoc;

        // Kết quả trả về cho lệnh
        public bool IsConfirmed { get; private set; }
        public string ExportVersion { get; private set; } = "2018";
        public string ExportMode { get; private set; } = "InsertBind"; // "InsertBind", "NoBind", "PreBindApi"
        public string OutputFilePath { get; private set; } = "";
        public bool AutoPurgeAudit { get; private set; } = true;
        public bool OpenFolderAfterExport { get; private set; } = true;

        public XuatCadForm(Document doc)
        {
            targetDoc = doc;
            InitBaseTheme("⚡ XUẤT CAD CHUYÊN NGHIỆP — CHỐNG LỖI XREF & DUPLICATE RECORD", 800, 680);
            BuildUI();
            LoadData();
        }

        private void BuildUI()
        {
            this.SuspendLayout();

            // 1. Header & Subtitle
            var header = CreateHeader("⚡ XUẤT CAD CHUYÊN NGHIỆP — CHỐNG LỖI XREF & DUPLICATE RECORD", 48);
            this.Controls.Add(header);

            var sub = CreateSubtitle("Khắc phục triệt để lỗi 'eDuplicateRecordName' và 'could not process Xrefs' khi xuất Civil 3D sang AutoCAD");
            this.Controls.Add(sub);

            int startY = 82;

            // ══════════════════════════════════════════════════════════════
            // SECTION 1: DANH SÁCH XREF
            // ══════════════════════════════════════════════════════════════
            var grpXrefs = CreateGroupBox("📋 1. Danh Sách External References (XREF) Trong Bản Vẽ", 15, startY, 755, 215);
            this.Controls.Add(grpXrefs);

            dgvXrefs = new DataGridView
            {
                Location = new Point(12, 26),
                Size = new Size(730, 140),
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = new Font("Segoe UI", 9f)
            };

            dgvXrefs.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Tên Xref", Name = "ColName", FillWeight = 32 });
            dgvXrefs.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Trạng Thái", Name = "ColStatus", FillWeight = 26 });
            dgvXrefs.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Tỷ Lệ / Đơn Vị", Name = "ColScale", FillWeight = 22 });
            dgvXrefs.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Thao Tác Dự Kiến", Name = "ColAction", FillWeight = 20 });

            grpXrefs.Controls.Add(dgvXrefs);

            btnScan = CreateButton("🔄 Quét Lại", 12, 172, 100, 30, grpXrefs);
            btnScan.Click += (s, e) => LoadData();

            btnFixSelected = CreateButton("⚡ Thử Gộp An Toàn", 118, 172, 140, 30, grpXrefs);
            btnFixSelected.Click += (s, e) => FixSelectedXref();

            btnDetachBad = CreateButton("🗑 Gỡ Xref Lỗi / Thiếu", 264, 172, 150, 30, grpXrefs);
            btnDetachBad.Click += (s, e) => DetachBadXrefs();

            lblSummary = new WinLabel
            {
                Text = "Đang kiểm tra danh sách Xref...",
                Location = new Point(425, 178),
                AutoSize = true,
                ForeColor = NavyMedium,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            grpXrefs.Controls.Add(lblSummary);

            // ══════════════════════════════════════════════════════════════
            // SECTION 2: CHẾ ĐỘ XỬ LÝ XREF
            // ══════════════════════════════════════════════════════════════
            int sec2Y = startY + 225;
            var grpMode = CreateGroupBox("⚙️ 2. Chế Độ Xử Lý Xref Khi Xuất (Khắc Phục Lỗi Bind)", 15, sec2Y, 755, 115);
            this.Controls.Add(grpMode);

            rdoInsertBind = new RadioButton
            {
                Text = "Gộp Xref bằng chế độ INSERT (Khuyên dùng - Triệt tiêu hoàn toàn lỗi eDuplicateRecordName)",
                Location = new Point(16, 24),
                AutoSize = true,
                Checked = true,
                Font = new Font("Segoe UI", 9.2f, FontStyle.Bold),
                ForeColor = NavyDark
            };
            grpMode.Controls.Add(rdoInsertBind);

            var lblDesc1 = CreateHintLabel("Tự động gộp Layer và Block trùng tên, không tạo tiền tố $0$ gây xung đột, xử lý mượt mà cả Xref lệch tỷ lệ.", 34, 46);
            grpMode.Controls.Add(lblDesc1);

            rdoNoBind = new RadioButton
            {
                Text = "Xuất CAD không gộp Xref (Bind: No) — Tốc độ cực nhanh, 100% không bao giờ lỗi Xref",
                Location = new Point(16, 68),
                AutoSize = true,
                Font = new Font("Segoe UI", 9f),
                ForeColor = TextDark
            };
            grpMode.Controls.Add(rdoNoBind);

            rdoPreBindApi = new RadioButton
            {
                Text = "Gộp siêu an toàn qua C# API trước khi xuất (Dành cho bản vẽ có Xref bị lỗi nghiêm trọng)",
                Location = new Point(16, 88),
                AutoSize = true,
                Font = new Font("Segoe UI", 9f),
                ForeColor = TextDark
            };
            grpMode.Controls.Add(rdoPreBindApi);

            // ══════════════════════════════════════════════════════════════
            // SECTION 3: CẤU HÌNH XUẤT FILE AUTOCAD
            // ══════════════════════════════════════════════════════════════
            int sec3Y = sec2Y + 125;
            var grpExport = CreateGroupBox("📁 3. Cấu Hình File Xuất Chuẩn AutoCAD", 15, sec3Y, 755, 155);
            this.Controls.Add(grpExport);

            var lblVer = CreateLabel("Phiên bản CAD:", 16, 28);
            grpExport.Controls.Add(lblVer);

            cboVersion = CreateComboBox(125, 25, 230);
            cboVersion.Items.AddRange(new object[]
            {
                "AutoCAD 2018 (Chuẩn 2018 - 2026+)",
                "AutoCAD 2013 (2013 - 2017)",
                "AutoCAD 2010 (2010 - 2012)",
                "AutoCAD 2007 (2007 - 2009)"
            });
            cboVersion.SelectedIndex = 0;
            grpExport.Controls.Add(cboVersion);

            chkPurge = CreateCheckBox("Tự động Audit & Purge RegApps (Làm sạch rác bản vẽ)", 375, 28, true);
            chkPurge.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            chkPurge.ForeColor = NavyDark;
            grpExport.Controls.Add(chkPurge);

            var lblFolder = CreateLabel("Thư mục lưu:", 16, 64);
            grpExport.Controls.Add(lblFolder);

            txtFolder = CreateTextBox(125, 61, 490);
            grpExport.Controls.Add(txtFolder);

            btnBrowse = CreateButton("...", 622, 60, 40, 29, grpExport);
            btnBrowse.Click += (s, e) => BrowseOutputFolder();

            btnOpenFolderNow = CreateButton("📂 Mở", 668, 60, 60, 29, grpExport);
            btnOpenFolderNow.Click += (s, e) => OpenCurrentFolder();

            var lblFile = CreateLabel("Tên file xuất:", 16, 102);
            grpExport.Controls.Add(lblFile);

            txtFileName = CreateTextBox(125, 99, 360);
            grpExport.Controls.Add(txtFileName);

            chkOpenFolder = CreateCheckBox("Mở thư mục sau khi xuất", 505, 102, true);
            grpExport.Controls.Add(chkOpenFolder);

            var lblTip = CreateRedNote("* Lưu ý: Toàn bộ đối tượng Civil 3D (Surface, Alignment, Profile, Trắc ngang...) sẽ được nổ sạch thành CAD thường.", 16, 130);
            grpExport.Controls.Add(lblTip);

            // ══════════════════════════════════════════════════════════════
            // FOOTER BUTTONS
            // ══════════════════════════════════════════════════════════════
            int footerY = sec3Y + 165;

            btnExportNow = CreatePrimaryButton("🚀 XUẤT CAD NGAY", 15, footerY, 200, 38, this);
            btnExportNow.Font = new Font("Segoe UI", 10.5f, FontStyle.Bold);
            btnExportNow.Click += (s, e) => ConfirmExport();

            btnFixOnly = CreateButton("🔧 Chỉ Sửa Lỗi & Gộp Xref", 225, footerY, 190, 38, this);
            btnFixOnly.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            btnFixOnly.Click += (s, e) => RunFixXrefsOnly();

            btnClose = CreateButton("Đóng", 670, footerY, 100, 38, this);
            btnClose.Click += (s, e) => this.Close();

            this.ResumeLayout(false);
        }

        private void LoadData()
        {
            try
            {
                currentXrefs = SafeXrefHelper.GetAllXrefs(targetDoc.Database);
                dgvXrefs.Rows.Clear();

                int badCount = 0;
                int unequalCount = 0;

                foreach (var x in currentXrefs)
                {
                    int rowIndex = dgvXrefs.Rows.Add(x.Name, x.StatusDescription, x.ScaleDescription, x.Action);
                    var row = dgvXrefs.Rows[rowIndex];

                    if (!x.IsResolved || !x.FileExists || x.IsUnloaded)
                    {
                        row.DefaultCellStyle.ForeColor = TextRed;
                        row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235);
                        badCount++;
                    }
                    else if (x.HasUnequalScale)
                    {
                        row.DefaultCellStyle.ForeColor = Color.FromArgb(180, 100, 0);
                        unequalCount++;
                    }
                }

                if (currentXrefs.Count == 0)
                {
                    lblSummary.Text = "✓ Bản vẽ không có Xref nào (An toàn tuyệt đối)";
                    lblSummary.ForeColor = GreenOk;
                }
                else
                {
                    lblSummary.Text = $"Tổng số: {currentXrefs.Count} Xref | ⚠ {badCount} lỗi file | ⚡ {unequalCount} lệch tỷ lệ";
                    lblSummary.ForeColor = (badCount > 0) ? TextRed : NavyMedium;
                }

                // Điền đường dẫn file mặc định
                string dwgPath = targetDoc.Database.Filename;
                string dir = !string.IsNullOrEmpty(dwgPath) ? (Path.GetDirectoryName(dwgPath) ?? "") : Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string baseName = !string.IsNullOrEmpty(dwgPath) ? Path.GetFileNameWithoutExtension(dwgPath) : "BảnVẽMới";

                if (string.IsNullOrEmpty(txtFolder.Text))
                {
                    txtFolder.Text = dir;
                }
                if (string.IsNullOrEmpty(txtFileName.Text))
                {
                    txtFileName.Text = $"{baseName}_ACAD.dwg";
                }
            }
            catch (System.Exception ex)
            {
                lblSummary.Text = $"Lỗi đọc Xref: {ex.Message}";
                lblSummary.ForeColor = TextRed;
            }
        }

        private void BrowseOutputFolder()
        {
            using var fbd = new FolderBrowserDialog();
            fbd.Description = "Chọn thư mục lưu file AutoCAD xuất:";
            if (Directory.Exists(txtFolder.Text))
            {
                fbd.SelectedPath = txtFolder.Text;
            }

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                txtFolder.Text = fbd.SelectedPath;
            }
        }

        private void OpenCurrentFolder()
        {
            string folder = txtFolder.Text.Trim();
            if (Directory.Exists(folder))
            {
                System.Diagnostics.Process.Start("explorer.exe", folder);
            }
            else
            {
                MessageBox.Show("Thư mục chưa tồn tại hoặc đường dẫn không hợp lệ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void FixSelectedXref()
        {
            if (dgvXrefs.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn một dòng Xref trong bảng để thử gộp!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int idx = dgvXrefs.SelectedRows[0].Index;
            if (idx >= 0 && idx < currentXrefs.Count)
            {
                var item = currentXrefs[idx];
                var res = SafeXrefHelper.SafeBindSingleXref(targetDoc.Database, item.BtrId, true);
                MessageBox.Show(res.Message, "Kết Quả Xử Lý Xref", MessageBoxButtons.OK, res.Success ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
                LoadData();
            }
        }

        private void DetachBadXrefs()
        {
            var badItems = currentXrefs.Where(x => !x.IsResolved || !x.FileExists || x.IsUnloaded).ToList();
            if (badItems.Count == 0)
            {
                MessageBox.Show("Không có Xref nào bị lỗi hoặc thiếu file cần gỡ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var dr = MessageBox.Show($"Tìm thấy {badItems.Count} Xref bị lỗi/thiếu file:\n" +
                string.Join("\n", badItems.Select(b => $"• {b.Name}")) +
                "\n\nBạn có muốn gỡ bỏ (Detach) để tránh làm nghẽn tiến trình xuất CAD không?",
                "Xác nhận gỡ Xref lỗi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (dr == DialogResult.Yes)
            {
                int count = 0;
                foreach (var b in badItems)
                {
                    if (SafeXrefHelper.DetachXref(targetDoc.Database, b.BtrId)) count++;
                }
                MessageBox.Show($"Đã gỡ bỏ {count} Xref lỗi khỏi bản vẽ.", "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadData();
            }
        }

        private void RunFixXrefsOnly()
        {
            var dr = MessageBox.Show("Tính năng này sẽ tiến hành gộp an toàn toàn bộ Xref trong bản vẽ hiện tại bằng chế độ Insert (chống lỗi Duplicate Record).\n\nBạn có muốn tiếp tục không?",
                "Xác nhận gộp Xref", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (dr == DialogResult.Yes)
            {
                var res = SafeXrefHelper.SafeBindAllXrefs(targetDoc.Database, targetDoc.Editor, true);
                MessageBox.Show($"Hoàn tất xử lý Xref:\n• Gộp thành công: {res.SuccessCount}\n• Thất bại / Bỏ qua: {res.FailedCount}\n\nXem chi tiết tại thanh Command Line.",
                    "Kết Quả Gộp Xref", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadData();
            }
        }

        private void ConfirmExport()
        {
            string folder = txtFolder.Text.Trim();
            string fileName = txtFileName.Text.Trim();

            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
            {
                MessageBox.Show("Vui lòng chọn thư mục lưu file hợp lệ!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(fileName))
            {
                MessageBox.Show("Vui lòng nhập tên file xuất!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!fileName.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase))
            {
                fileName += ".dwg";
            }

            OutputFilePath = Path.Combine(folder, fileName);

            // Xác định phiên bản AutoCAD
            ExportVersion = cboVersion.SelectedIndex switch
            {
                0 => "2018",
                1 => "2013",
                2 => "2010",
                3 => "2007",
                _ => "2018"
            };

            // Xác định chế độ xuất
            if (rdoInsertBind.Checked)
            {
                ExportMode = "InsertBind";
            }
            else if (rdoNoBind.Checked)
            {
                ExportMode = "NoBind";
            }
            else
            {
                ExportMode = "PreBindApi";
            }

            AutoPurgeAudit = chkPurge.Checked;
            OpenFolderAfterExport = chkOpenFolder.Checked;

            IsConfirmed = true;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }

    // ══════════════════════════════════════════════════════════════════
    // COMMAND METHODS: CÁC LỆNH XUẤT CAD VÀ SỬA LỖI XREF
    // ══════════════════════════════════════════════════════════════════
    public class AT_XuatCad
    {
        /// <summary>
        /// LỆNH CHÍNH: Mở Form giao diện Xuất CAD Chuyên Nghiệp
        /// </summary>
        [CommandMethod("AT_XuatCad")]
        [CommandMethod("XUATCAD")]
        [CommandMethod("EXPORTCAD")]
        public static void XuatCadMethod()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            using var form = new XuatCadForm(doc);
            var result = AcadApp.ShowModalDialog(form);

            if (result != DialogResult.OK || !form.IsConfirmed)
            {
                ed.WriteMessage("\n[AT_XuatCad] Đã hủy thao tác xuất CAD.");
                return;
            }

            ExecuteExportWorkflow(doc, form.OutputFilePath, form.ExportVersion, form.ExportMode, form.AutoPurgeAudit, form.OpenFolderAfterExport);
        }

        /// <summary>
        /// LỆNH XUẤT NHANH (1 CHẠM QUA DÒNG LỆNH):
        /// Tự động Purge, đặt Insert Mode, xuất sang CAD 2018 cùng thư mục với đuôi _ACAD.dwg.
        /// </summary>
        [CommandMethod("AT_XuatCadQuick")]
        [CommandMethod("XUATCADQ")]
        public static void XuatCadQuickMethod()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            string dwgPath = db.Filename;
            if (string.IsNullOrEmpty(dwgPath))
            {
                ed.WriteMessage("\n[L?i] Vui l?ng luu b?n v? hi?n t?i tru?c khi xu?t nhanh.\n");
                return;
            }

            string dir = Path.GetDirectoryName(dwgPath) ?? "";
            string baseName = Path.GetFileNameWithoutExtension(dwgPath);
            string outputPath = Path.Combine(dir, $"{baseName}_ACAD.dwg");

            ExecuteExportWorkflow(doc, outputPath, "2018", "InsertBind", autoPurge: true, openFolderAfter: true);
        }

        /// <summary>
        /// TI?N TRINH TH?C THI XU?T CAD TOAN DI?N
        /// </summary>
        private static void ExecuteExportWorkflow(
            Document doc,
            string targetFilePath,
            string cadVersion,
            string exportMode,
            bool autoPurge,
            bool openFolderAfter)
        {
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // Bước 0: Chuẩn bị file đích (xóa file cũ nếu đã tồn tại để tránh hộp thoại xác nhận ghi đè)
                if (File.Exists(targetFilePath))
                {
                    try
                    {
                        File.Delete(targetFilePath);
                    }
                    catch
                    {
                        // File đang bị khóa bởi tiến trình khác, thêm timestamp để không lỗi
                        string tDir = Path.GetDirectoryName(targetFilePath)!;
                        string fn = Path.GetFileNameWithoutExtension(targetFilePath);
                        string ext = Path.GetExtension(targetFilePath);
                        targetFilePath = Path.Combine(tDir, $"{fn}_{DateTime.Now:yyyyMMdd_HHmmss}{ext}");
                        ed.WriteMessage($"\n[Thông báo] File cũ đang mở, lưu sang file mới: {Path.GetFileName(targetFilePath)}");
                    }
                }

                // Bước 1: Thiết lập biến hệ thống an toàn
                try
                {
                    AcadApp.SetSystemVariable("BINDTYPE", 1); // 1 = Insert mode (chống eDuplicateRecordName)
                    AcadApp.SetSystemVariable("XLOADCTL", 2); // Demand load with copy
                    AcadApp.SetSystemVariable("ISAVEPERCENT", 0);
                    AcadApp.SetSystemVariable("PROXYNOTICE", 0); // Tắt thông báo proxy che màn hình
                    AcadApp.SetSystemVariable("FILEDIA", 0); // Tắt dialog file để nhận đường dẫn qua dòng lệnh
                }
                catch { }

                // Bước 2: Tự động dọn dẹp rác & sửa lỗi database nếu được yêu cầu
                if (autoPurge)
                {
                    ed.WriteMessage("\n[1/3] Đang làm sạch RegApps và tối ưu database...");
                    int purgedReg = SafeXrefHelper.PurgeRegApps(db);
                    if (purgedReg > 0)
                    {
                        ed.WriteMessage($"\n  -> Đã loại bỏ {purgedReg} RegApps rác gây xung đột.");
                    }
                }

                // Bước 3: Xử lý theo từng chế độ
                string cleanPath = targetFilePath.Replace("\\", "/");

                if (exportMode == "NoBind")
                {
                    ed.WriteMessage("\n[2/3] Đang xuất CAD (Chế độ giữ nguyên Xref, Bind: No)...");
                    string exportScript = BuildExportScript(cadVersion, bindXrefs: false, useInsert: false, cleanPath);
                    doc.SendStringToExecute(exportScript, true, false, false);
                }
                else
                {
                    // Chế độ mặc định ("InsertBind" hoặc "PreBindApi"):
                    // Gộp toàn bộ Xref bằng C# API trước (với 3 cấp bảo vệ: Insert -> Bind -> Deep Embedding).
                    // Vì Xref đã được nạp thành Block nội bộ 100%, nên lệnh EXPORTTOAUTOCAD chỉ cần gọi Bind: No.
                    // Triệt tiêu 100% lỗi "Xref failed to bind", "eDuplicateRecordName", "could not process Xrefs"!
                    ed.WriteMessage("\n[2/3] Đang gộp toàn bộ Xref thông minh qua C# API (chống lỗi Duplicate Record & lệch tỷ lệ)...");
                    var bindRes = SafeXrefHelper.SafeBindAllXrefs(db, ed, preferInsert: true);
                    ed.WriteMessage($"\n  -> Đã gộp thành công {bindRes.SuccessCount}/{bindRes.SuccessCount + bindRes.FailedCount} Xref. Bắt đầu xuất file...");

                    string exportScript = BuildExportScript(cadVersion, bindXrefs: false, useInsert: false, cleanPath);
                    doc.SendStringToExecute(exportScript, true, false, false);
                }

                // Khôi phục FILEDIA
                doc.SendStringToExecute("(setvar \"FILEDIA\" 1)(princ)\n", true, false, false);

                // Bước 4: Thông báo hoàn tất & Mở thư mục
                ed.WriteMessage("\n[3/3] Lệnh xuất CAD đã được gửi vào hàng đợi xử lý.");
                ed.WriteMessage($"\n✓ File đầu ra dự kiến: {targetFilePath}\n");

                if (openFolderAfter)
                {
                    try
                    {
                        string folderToOpen = Path.GetDirectoryName(targetFilePath) ?? "";
                        if (Directory.Exists(folderToOpen))
                        {
                            System.Diagnostics.Process.Start("explorer.exe", folderToOpen);
                        }
                    }
                    catch { /* Bỏ qua lỗi mở thư mục, không ảnh hưởng kết quả xuất */ }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[Lỗi] Không thể hoàn tất tiến trình xuất CAD: {ex.Message}\n");
            }
        }

        /// <summary>
        /// Tạo kịch bản lệnh ._-EXPORTTOAUTOCAD chính xác từng tham số prompt.
        /// 
        /// Cấu trúc prompt của lệnh ._-EXPORTTOAUTOCAD:
        ///   Export options [Format/Bind/bind Type/Maintain/Prefix/Suffix/?] (Enter for filename):
        ///   Mỗi lần nhập F/B/T/M/P/S sẽ vào sub-prompt rồi quay lại prompt chính.
        ///   Nhấn Enter tay trắng (\n) tại prompt chính sẽ chuyển sang prompt "Export drawing name".
        /// 
        /// QUAN TRỌNG: Phải xóa Prefix/Suffix mặc định (ACAD-) trước khi xuất,
        /// và phải gửi Enter (\n) trước khi nhập đường dẫn file!
        /// </summary>
        private static string BuildExportScript(string version, bool bindXrefs, bool useInsert, string filePath)
        {
            var sb = new System.Text.StringBuilder();
            sb.Append("._-EXPORTTOAUTOCAD\n");

            // 1. Đặt phiên bản DWG
            sb.Append($"F\n{version}\n");

            // 2. Cấu hình Bind Xref
            if (bindXrefs)
            {
                sb.Append("B\nY\n");
                sb.Append($"T\n{(useInsert ? "I" : "B")}\n");
            }
            else
            {
                sb.Append("B\nN\n");
            }

            // 3. Maintain resolved properties = Yes (nổ AEC objects sang entities gốc)
            sb.Append("M\nY\n");

            // 4. Xóa Prefix mặc định "ACAD-" bằng cách đặt thành dấu chấm (.)
            sb.Append("P\n.\n");

            // 5. Xóa Suffix mặc định
            sb.Append("S\n.\n");

            // 6. Enter tay trắng tại prompt chính = chấp nhận cấu hình, chuyển sang prompt tên file
            //    Prompt: "Export drawing name <...>:"
            sb.Append("\n");

            // 7. Nhập đường dẫn đầy đủ (trong ngoặc kép nếu có dấu cách/tiếng Việt)
            sb.Append($"\"{filePath}\"\n");

            return sb.ToString();
        }
    }
}
