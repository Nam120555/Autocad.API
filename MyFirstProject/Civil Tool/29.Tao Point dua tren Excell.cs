// (C) Copyright 2026
// Import tọa độ từ Excel vào Civil 3D — Chèn Block + gán Attribute Name
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using ClosedXML.Excel;
using MyFirstProject.Extensions;

using WinLabel = System.Windows.Forms.Label;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

#pragma warning disable CS8600, CS8602, CS8604, CS8618

[assembly: CommandClass(typeof(MyFirstProject.ImportExcelPoints))]

namespace MyFirstProject
{
    #region Data Classes

    /// <summary>
    /// Dữ liệu 1 điểm đọc từ Excel
    /// </summary>
    public class ExcelPointData
    {
        public string Name { get; set; } = "";
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public string Description { get; set; } = "";

        public override string ToString() => $"{Name} ({X:F3}, {Y:F3}, {Z:F3})";
    }

    #endregion

    #region Form

    /// <summary>
    /// Form chọn file Excel, mapping cột, chọn Block, preview dữ liệu
    /// Theme Navy/White kế thừa CivilFormBase
    /// </summary>
    public class FormImportExcelPoint : CivilFormBase
    {
        // ═══ Controls ═══
        private TextBox txtFilePath;
        private Button btnBrowse;
        private ComboBox cboSheet;
        private ComboBox cboColName, cboColX, cboColY, cboColZ;
        private NumericUpDown nudStartRow;
        private DataGridView dgvPreview;
        private ComboBox cboBlock;
        private TextBox txtLayer;
        private NumericUpDown nudScale;
        private CheckBox chkCreateCogoPoint;
        private CheckBox chkHasHeader;
        private Button btnRefresh;

        // ═══ Dữ liệu ═══
        private XLWorkbook? workbook;
        public List<ExcelPointData> PointDataList { get; private set; } = new();
        public string SelectedBlockName { get; private set; } = "";
        public string LayerName { get; private set; } = "EXCEL_POINTS";
        public double BlockScale { get; private set; } = 1.0;
        public bool CreateCogoPoint { get; private set; } = false;

        public FormImportExcelPoint()
        {
            InitializeComponent();
            LoadBlockList();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("Import tọa độ từ Excel", 780, 680);

            // ═══ Header ═══
            var header = CreateHeader("📊 Import tọa độ từ Excel → Civil 3D");
            this.Controls.Add(header);

            var subtitle = CreateSubtitle("Đọc tọa độ X, Y từ file Excel và chèn Block / Point vào bản vẽ");
            this.Controls.Add(subtitle);

            // ═══ Panel cuộn chính ═══
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(20, 10, 20, 10)
            };
            this.Controls.Add(mainPanel);
            mainPanel.BringToFront();

            int y = 10;

            // ── 1. Chọn file Excel ──
            var grpFile = CreateGroupBox("📁 File Excel", 0, y, 720, 75);
            mainPanel.Controls.Add(grpFile);

            grpFile.Controls.Add(CreateLabel("File:", 12, 30));

            txtFilePath = CreateTextBox(55, 27, 520);
            txtFilePath.ReadOnly = true;
            txtFilePath.BackColor = Color.FromArgb(245, 247, 250);
            grpFile.Controls.Add(txtFilePath);

            btnBrowse = CreatePrimaryButton("Chọn file...", 585, 25, 120, 30, grpFile);
            btnBrowse.Click += BtnBrowse_Click;

            y += 85;

            // ── 2. Cấu hình Sheet & cột ──
            var grpMapping = CreateGroupBox("⚙ Cấu hình cột dữ liệu", 0, y, 720, 180);
            mainPanel.Controls.Add(grpMapping);

            // Sheet
            grpMapping.Controls.Add(CreateLabel("Sheet:", 12, 32));
            cboSheet = CreateComboBox(120, 29, 240);
            cboSheet.SelectedIndexChanged += CboSheet_Changed;
            grpMapping.Controls.Add(cboSheet);

            chkHasHeader = CreateCheckBox("Dòng đầu là tiêu đề", 400, 32, true);
            grpMapping.Controls.Add(chkHasHeader);

            // Dòng bắt đầu
            grpMapping.Controls.Add(CreateLabel("Dòng bắt đầu:", 12, 67));
            nudStartRow = CreateNumeric(120, 64, 80, 1, 9999, 2, 0);
            grpMapping.Controls.Add(nudStartRow);

            // Cột Name
            grpMapping.Controls.Add(CreateLabel("Cột Name:", 12, 102));
            cboColName = CreateComboBox(120, 99, 120);
            grpMapping.Controls.Add(cboColName);

            // Cột X
            grpMapping.Controls.Add(CreateLabel("Cột X:", 260, 102));
            cboColX = CreateComboBox(320, 99, 120);
            grpMapping.Controls.Add(cboColX);

            // Cột Y
            grpMapping.Controls.Add(CreateLabel("Cột Y:", 460, 102));
            cboColY = CreateComboBox(520, 99, 120);
            grpMapping.Controls.Add(cboColY);

            // Cột Z (tùy chọn)
            grpMapping.Controls.Add(CreateLabel("Cột Z:", 12, 140));
            cboColZ = CreateComboBox(120, 137, 120);
            grpMapping.Controls.Add(cboColZ);

            grpMapping.Controls.Add(CreateHintLabel("(tùy chọn — để trống nếu không có cao độ)", 250, 142));

            // Nút Refresh Preview
            btnRefresh = CreateButton("🔄 Xem trước", 520, 135, 120, 30, grpMapping);
            btnRefresh.Click += BtnRefresh_Click;

            y += 190;

            // ── 3. Preview ──
            var grpPreview = CreateGroupBox("👁 Xem trước dữ liệu (10 dòng đầu)", 0, y, 720, 160);
            mainPanel.Controls.Add(grpPreview);

            dgvPreview = CreateDataGrid(8, 26, 700, 125);
            dgvPreview.Columns.Add("colSTT", "STT");
            dgvPreview.Columns.Add("colName", "Name");
            dgvPreview.Columns.Add("colX", "X");
            dgvPreview.Columns.Add("colY", "Y");
            dgvPreview.Columns.Add("colZ", "Z");
            dgvPreview.Columns["colSTT"].Width = 50;
            grpPreview.Controls.Add(dgvPreview);

            y += 170;

            // ── 4. Cấu hình Block ──
            var grpBlock = CreateGroupBox("🔲 Block & Layer", 0, y, 720, 110);
            mainPanel.Controls.Add(grpBlock);

            grpBlock.Controls.Add(CreateLabel("Block:", 12, 30));
            cboBlock = CreateComboBox(120, 27, 250);
            grpBlock.Controls.Add(cboBlock);

            grpBlock.Controls.Add(CreateHintLabel("(để trống → chèn DBPoint + Text)", 380, 32));

            grpBlock.Controls.Add(CreateLabel("Layer:", 12, 65));
            txtLayer = CreateTextBox(120, 62, 200, "EXCEL_POINTS");
            grpBlock.Controls.Add(txtLayer);

            grpBlock.Controls.Add(CreateLabel("Tỷ lệ:", 350, 65));
            nudScale = CreateNumeric(410, 62, 80, 0.01m, 1000, 1, 2);
            grpBlock.Controls.Add(nudScale);

            chkCreateCogoPoint = CreateCheckBox("Tạo COGO Point", 520, 65, false);
            grpBlock.Controls.Add(chkCreateCogoPoint);

            y += 120;

            // ═══ Bottom Bar ═══
            var bottomBar = CreateBottomBar(55);
            this.Controls.Add(bottomBar);

            var btnCancel = CreateButton("Hủy", 620, 10, 100, 34, bottomBar);
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            var btnExecute = CreateSuccessButton("✅ Thực hiện", 470, 10, 140, 34, bottomBar);
            btnExecute.Click += BtnExecute_Click;

            var lblCount = CreateHintLabel("", 20, 18);
            lblCount.Name = "lblCount";
            bottomBar.Controls.Add(lblCount);
        }

        // ═══ Event Handlers ═══

        private void BtnBrowse_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Title = "Chọn file Excel chứa tọa độ",
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                FilterIndex = 1
            };

            if (ofd.ShowDialog() != DialogResult.OK) return;

            txtFilePath.Text = ofd.FileName;

            try
            {
                workbook?.Dispose();
                workbook = new XLWorkbook(ofd.FileName);

                // Load sheets
                cboSheet.Items.Clear();
                foreach (var ws in workbook.Worksheets)
                {
                    cboSheet.Items.Add(ws.Name);
                }
                if (cboSheet.Items.Count > 0)
                    cboSheet.SelectedIndex = 0;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Không thể mở file Excel:\n{ex.Message}",
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CboSheet_Changed(object? sender, EventArgs e)
        {
            if (workbook == null || cboSheet.SelectedIndex < 0) return;

            try
            {
                var ws = workbook.Worksheet(cboSheet.Text);
                int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 1;

                // Load columns into comboboxes
                var colItems = new List<string> { "(không chọn)" };
                for (int c = 1; c <= lastCol; c++)
                {
                    string colLetter = GetColumnLetter(c);
                    // Nếu có header, hiển thị tên header
                    string headerVal = "";
                    if (chkHasHeader.Checked)
                    {
                        var cell = ws.Cell(1, c);
                        if (!cell.IsEmpty())
                            headerVal = $" — {cell.GetString()}";
                    }
                    colItems.Add($"{colLetter}{headerVal}");
                }

                UpdateCombo(cboColName, colItems, lastCol >= 1 ? 1 : 0);
                UpdateCombo(cboColX, colItems, lastCol >= 2 ? 2 : 0);
                UpdateCombo(cboColY, colItems, lastCol >= 3 ? 3 : 0);
                UpdateCombo(cboColZ, colItems, lastCol >= 4 ? 4 : 0);
            }
            catch { }
        }

        private void UpdateCombo(ComboBox cbo, List<string> items, int selectedIndex)
        {
            cbo.Items.Clear();
            foreach (var item in items) cbo.Items.Add(item);
            if (selectedIndex < cbo.Items.Count) cbo.SelectedIndex = selectedIndex;
        }

        private void BtnRefresh_Click(object? sender, EventArgs e)
        {
            LoadPreview();
        }

        private void LoadPreview()
        {
            dgvPreview.Rows.Clear();
            if (workbook == null || cboSheet.SelectedIndex < 0) return;

            try
            {
                var ws = workbook.Worksheet(cboSheet.Text);
                int startRow = (int)nudStartRow.Value;
                int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
                int colName = cboColName.SelectedIndex;
                int colX = cboColX.SelectedIndex;
                int colY = cboColY.SelectedIndex;
                int colZ = cboColZ.SelectedIndex;

                int count = 0;
                for (int r = startRow; r <= lastRow && count < 10; r++)
                {
                    string name = colName > 0 ? ws.Cell(r, colName).GetString() : "";
                    string x = colX > 0 ? ws.Cell(r, colX).GetString() : "";
                    string y = colY > 0 ? ws.Cell(r, colY).GetString() : "";
                    string z = colZ > 0 ? ws.Cell(r, colZ).GetString() : "";

                    if (string.IsNullOrWhiteSpace(x) && string.IsNullOrWhiteSpace(y)) continue;

                    dgvPreview.Rows.Add(count + 1, name, x, y, z);
                    count++;
                }

                // Cập nhật số lượng
                int totalRows = lastRow - startRow + 1;
                var lblCount = this.Controls.Find("lblCount", true).FirstOrDefault() as WinLabel;
                if (lblCount != null)
                    lblCount.Text = $"Tổng số dòng dữ liệu: {totalRows}";
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi đọc dữ liệu:\n{ex.Message}",
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExecute_Click(object? sender, EventArgs e)
        {
            if (workbook == null)
            {
                MessageBox.Show("Vui lòng chọn file Excel!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (cboColX.SelectedIndex <= 0 || cboColY.SelectedIndex <= 0)
            {
                MessageBox.Show("Vui lòng chọn cột X và cột Y!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                ReadAllData();

                if (PointDataList.Count == 0)
                {
                    MessageBox.Show("Không có dữ liệu hợp lệ trong file Excel!",
                        "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Lưu cấu hình
                SelectedBlockName = cboBlock.SelectedIndex > 0 ? cboBlock.Text : "";
                LayerName = string.IsNullOrWhiteSpace(txtLayer.Text) ? "EXCEL_POINTS" : txtLayer.Text.Trim();
                BlockScale = (double)nudScale.Value;
                CreateCogoPoint = chkCreateCogoPoint.Checked;

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Lỗi:\n{ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ReadAllData()
        {
            PointDataList.Clear();
            if (workbook == null || cboSheet.SelectedIndex < 0) return;

            var ws = workbook.Worksheet(cboSheet.Text);
            int startRow = (int)nudStartRow.Value;
            int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            int colName = cboColName.SelectedIndex;
            int colX = cboColX.SelectedIndex;
            int colY = cboColY.SelectedIndex;
            int colZ = cboColZ.SelectedIndex;

            for (int r = startRow; r <= lastRow; r++)
            {
                string xStr = colX > 0 ? ws.Cell(r, colX).GetString().Trim() : "";
                string yStr = colY > 0 ? ws.Cell(r, colY).GetString().Trim() : "";

                if (!double.TryParse(xStr, out double x)) continue;
                if (!double.TryParse(yStr, out double y)) continue;

                double z = 0;
                if (colZ > 0)
                {
                    string zStr = ws.Cell(r, colZ).GetString().Trim();
                    double.TryParse(zStr, out z);
                }

                string name = colName > 0 ? ws.Cell(r, colName).GetString().Trim() : $"P{r}";

                PointDataList.Add(new ExcelPointData
                {
                    Name = name,
                    X = x,
                    Y = y,
                    Z = z,
                    Description = name
                });
            }
        }

        private void LoadBlockList()
        {
            try
            {
                cboBlock.Items.Clear();
                cboBlock.Items.Add("(Không dùng Block — chèn Point + Text)");

                Document doc = AcadApp.DocumentManager.MdiActiveDocument;
                if (doc == null) { cboBlock.SelectedIndex = 0; return; }

                Database db = doc.Database;
                using Transaction tr = db.TransactionManager.StartTransaction();
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                foreach (ObjectId btrId in bt)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    // Chỉ hiện block do user tạo (không phải ModelSpace, PaperSpace, *anonymous...)
                    if (!btr.IsAnonymous && !btr.IsLayout && !btr.Name.StartsWith("*"))
                    {
                        cboBlock.Items.Add(btr.Name);
                    }
                }
                tr.Commit();

                cboBlock.SelectedIndex = 0;
            }
            catch { cboBlock.SelectedIndex = 0; }
        }

        private static string GetColumnLetter(int colNumber)
        {
            string result = "";
            while (colNumber > 0)
            {
                colNumber--;
                result = (char)('A' + colNumber % 26) + result;
                colNumber /= 26;
            }
            return result;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                workbook?.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    #endregion

    #region Command

    /// <summary>
    /// Lệnh Import tọa độ từ Excel vào Civil 3D
    /// </summary>
    class ImportExcelPoints
    {
        [CommandMethod("CTPO_TaoPointTuExcel")]
        public static void CTPO_TaoPointTuExcel()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Hiển thị Form
            FormImportExcelPoint form = new();
            if (AcadApp.ShowModalDialog(form) != DialogResult.OK)
            {
                ed.WriteMessage("\n❌ Đã hủy lệnh Import Excel.");
                return;
            }

            var pointList = form.PointDataList;
            string blockName = form.SelectedBlockName;
            string layerName = form.LayerName;
            double scale = form.BlockScale;
            bool createCogo = form.CreateCogoPoint;

            if (pointList.Count == 0)
            {
                ed.WriteMessage("\n⚠ Không có dữ liệu để import.");
                return;
            }

            ed.WriteMessage($"\n📊 Bắt đầu import {pointList.Count} điểm từ Excel...");

            using Transaction tr = db.TransactionManager.StartTransaction();
            try
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(
                    bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Tạo layer nếu chưa có
                EnsureLayer(db, tr, layerName);

                int successCount = 0;
                int errorCount = 0;
                bool useBlock = !string.IsNullOrEmpty(blockName) && bt.Has(blockName);

                foreach (var pt in pointList)
                {
                    try
                    {
                        Point3d insertPt = new(pt.X, pt.Y, pt.Z);

                        if (useBlock)
                        {
                            // ── Chèn BlockReference ──
                            ObjectId blockDefId = bt[blockName];
                            BlockReference blkRef = new(insertPt, blockDefId)
                            {
                                Layer = layerName,
                                ScaleFactors = new Scale3d(scale, scale, scale)
                            };

                            modelSpace.AppendEntity(blkRef);
                            tr.AddNewlyCreatedDBObject(blkRef, true);

                            // Gán Attribute nếu block có AttributeDefinition
                            BlockTableRecord blockDef = (BlockTableRecord)tr.GetObject(blockDefId, OpenMode.ForRead);
                            foreach (ObjectId attDefId in blockDef)
                            {
                                Autodesk.AutoCAD.DatabaseServices.DBObject obj = tr.GetObject(attDefId, OpenMode.ForRead);
                                if (obj is AttributeDefinition attDef)
                                {
                                    AttributeReference attRef = new();
                                    attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform);

                                    // Gán giá trị Name cho attribute
                                    string tagUpper = attDef.Tag.ToUpper();
                                    if (tagUpper == "NAME" || tagUpper == "TEN" || tagUpper == "LABEL"
                                        || tagUpper == "TEXT" || tagUpper == "TITLE")
                                    {
                                        attRef.TextString = pt.Name;
                                    }
                                    else if (tagUpper == "X" || tagUpper == "EASTING")
                                    {
                                        attRef.TextString = pt.X.ToString("F3");
                                    }
                                    else if (tagUpper == "Y" || tagUpper == "NORTHING")
                                    {
                                        attRef.TextString = pt.Y.ToString("F3");
                                    }
                                    else if (tagUpper == "Z" || tagUpper == "ELEVATION" || tagUpper == "CAODO")
                                    {
                                        attRef.TextString = pt.Z.ToString("F3");
                                    }
                                    else
                                    {
                                        attRef.TextString = pt.Name;
                                    }

                                    attRef.Layer = layerName;
                                    blkRef.AttributeCollection.AppendAttribute(attRef);
                                    tr.AddNewlyCreatedDBObject(attRef, true);
                                }
                            }
                        }
                        else
                        {
                            // ── Fallback: Chèn DBPoint + Text ──
                            DBPoint dbPoint = new(insertPt)
                            {
                                Layer = layerName
                            };
                            modelSpace.AppendEntity(dbPoint);
                            tr.AddNewlyCreatedDBObject(dbPoint, true);

                            // Tạo text hiển thị Name
                            if (!string.IsNullOrWhiteSpace(pt.Name))
                            {
                                DBText nameText = new()
                                {
                                    Position = new Point3d(pt.X + 0.5, pt.Y + 0.5, pt.Z),
                                    Height = 1.0,
                                    TextString = pt.Name,
                                    Layer = layerName
                                };
                                nameText.SetDatabaseDefaults();
                                modelSpace.AppendEntity(nameText);
                                tr.AddNewlyCreatedDBObject(nameText, true);
                            }
                        }

                        // Tạo COGO Point nếu có yêu cầu
                        if (createCogo)
                        {
                            UtilitiesC3D.CreateCogoPointFromPoint3D(insertPt, pt.Description);
                        }

                        successCount++;
                    }
                    catch
                    {
                        errorCount++;
                    }
                }

                tr.Commit();

                ed.WriteMessage($"\n✅ Import hoàn tất! Thành công: {successCount} | Lỗi: {errorCount}");
                if (useBlock)
                    ed.WriteMessage($"\n   Block: {blockName} | Layer: {layerName} | Tỷ lệ: {scale}");
                else
                    ed.WriteMessage($"\n   Chèn DBPoint + Text | Layer: {layerName}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Tạo layer nếu chưa tồn tại
        /// </summary>
        private static void EnsureLayer(Database db, Transaction tr, string layerName)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                LayerTableRecord ltr = new()
                {
                    Name = layerName,
                    Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                        Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 3) // Green
                };
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }
    }

    #endregion
}
