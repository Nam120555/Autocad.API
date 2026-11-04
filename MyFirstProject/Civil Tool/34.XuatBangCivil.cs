// Xuất bảng Civil 3D → CSV / Excel / AutoCAD Table thuần
// Nâng cấp từ CivilTableExport (Iteris-Soft) — đọc trực tiếp Table.Cells API
//
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.ApplicationServices;
using MyFirstProject.Extensions;
using ClosedXML.Excel;

// Aliases
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.XuatBangCivil))]

namespace MyFirstProject
{
    #region Data

    /// <summary>
    /// Dữ liệu bảng được đọc từ Civil 3D Table
    /// </summary>
    public class CivilTableData
    {
        public string TableName { get; set; } = "";
        public int RowCount { get; set; }
        public int ColCount { get; set; }
        public List<List<string>> Cells { get; set; } = new();

        public string GetCell(int row, int col)
        {
            if (row < 0 || row >= Cells.Count) return "";
            if (col < 0 || col >= Cells[row].Count) return "";
            return Cells[row][col] ?? "";
        }
    }

    #endregion

    #region Form

    /// <summary>
    /// Form preview + xuất bảng Civil 3D
    /// </summary>
    public class XuatBangCivilForm : CivilFormBase
    {
        private DataGridView dgvPreview = null!;
        private WinLabel lblInfo = null!;
        private ComboBox cboDelimiter = null!;
        private ComboBox cboEncoding = null!;
        private NumericUpDown numTextHeight = null!;
        private NumericUpDown numRowHeight = null!;

        public CivilTableData? TableData { get; set; }
        public enum ExportMode { None, CSV, Excel, AutoCADTable }
        public ExportMode SelectedMode { get; private set; } = ExportMode.None;
        public string CsvDelimiter => cboDelimiter.SelectedIndex switch
        {
            0 => ",",
            1 => ";",
            2 => "\t",
            _ => ","
        };
        public Encoding SelectedEncoding => cboEncoding.SelectedIndex switch
        {
            0 => Encoding.UTF8,
            1 => Encoding.GetEncoding(1252),
            _ => Encoding.UTF8
        };
        public double TextHeight => (double)numTextHeight.Value;
        public double RowHeight => (double)numRowHeight.Value;

        public XuatBangCivilForm(CivilTableData data)
        {
            TableData = data;
            InitializeUI();
            LoadPreview();
        }

        private void InitializeUI()
        {
            InitBaseTheme("XUẤT BẢNG CIVIL 3D — Civil Tool", 850, 680);

            // Header
            var header = CreateHeader("📊  XUẤT BẢNG CIVIL 3D", 50);
            this.Controls.Add(header);

            var subtitle = CreateSubtitle($"  Bảng: {TableData?.TableName}  |  {TableData?.RowCount} dòng × {TableData?.ColCount} cột");
            this.Controls.Add(subtitle);

            // Main content panel
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12, 8, 12, 0),
                AutoScroll = true,
                BackColor = BgWhite
            };
            this.Controls.Add(mainPanel);

            // Info label
            lblInfo = CreateLabel("", 0, 0);
            lblInfo.Dock = DockStyle.Top;
            lblInfo.Height = 28;
            lblInfo.ForeColor = TextHint;
            mainPanel.Controls.Add(lblInfo);

            // DataGridView preview
            dgvPreview = CreateDataGrid(0, 28, 810, 380);
            dgvPreview.Dock = DockStyle.Top;
            dgvPreview.Height = 380;
            mainPanel.Controls.Add(dgvPreview);

            // Options panel
            var pnlOptions = CreateGroupBox("⚙️  Tùy chọn xuất", 0, 420, 810, 90);
            pnlOptions.Dock = DockStyle.Top;
            mainPanel.Controls.Add(pnlOptions);

            // CSV delimiter
            pnlOptions.Controls.Add(CreateLabel("Delimiter CSV:", 12, 32));
            cboDelimiter = new ComboBox
            {
                Location = new Point(120, 28),
                Size = new Size(90, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextDark,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f)
            };
            cboDelimiter.Items.AddRange(new[] { ", (comma)", "; (semicolon)", "Tab" });
            cboDelimiter.SelectedIndex = 0;
            pnlOptions.Controls.Add(cboDelimiter);

            // Encoding
            pnlOptions.Controls.Add(CreateLabel("Encoding:", 230, 32));
            cboEncoding = new ComboBox
            {
                Location = new Point(310, 28),
                Size = new Size(130, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BgInput,
                ForeColor = TextDark,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f)
            };
            cboEncoding.Items.AddRange(new[] { "UTF-8", "Windows-1252" });
            cboEncoding.SelectedIndex = 0;
            pnlOptions.Controls.Add(cboEncoding);

            // AutoCAD Table options
            pnlOptions.Controls.Add(CreateLabel("Cao chữ:", 480, 32));
            numTextHeight = new NumericUpDown
            {
                Location = new Point(545, 28),
                Size = new Size(70, 28),
                Minimum = 1,
                Maximum = 50,
                Value = 2,
                DecimalPlaces = 1,
                Increment = 0.5m,
                BackColor = BgInput,
                ForeColor = TextDark
            };
            pnlOptions.Controls.Add(numTextHeight);

            pnlOptions.Controls.Add(CreateLabel("Cao dòng:", 630, 32));
            numRowHeight = new NumericUpDown
            {
                Location = new Point(705, 28),
                Size = new Size(70, 28),
                Minimum = 2,
                Maximum = 100,
                Value = 5,
                DecimalPlaces = 1,
                Increment = 0.5m,
                BackColor = BgInput,
                ForeColor = TextDark
            };
            pnlOptions.Controls.Add(numRowHeight);

            pnlOptions.Controls.Add(CreateHintLabel("💡 Cao chữ / Cao dòng áp dụng cho xuất AutoCAD Table", 12, 62));

            // Reorder controls (DockStyle.Top adds in reverse)
            mainPanel.Controls.SetChildIndex(lblInfo, 3);
            mainPanel.Controls.SetChildIndex(dgvPreview, 2);
            mainPanel.Controls.SetChildIndex(pnlOptions, 1);

            // Bottom bar with export buttons
            var bottomBar = CreateBottomBar(60);
            this.Controls.Add(bottomBar);

            var btnClose = CreateButton("✕  Đóng", 12, 10, 130, 40, bottomBar);
            btnClose.Click += (s, e) => { SelectedMode = ExportMode.None; this.Close(); };

            var btnCSV = CreatePrimaryButton("📄  Xuất CSV", 200, 10, 170, 40, bottomBar);
            btnCSV.Click += (s, e) => { SelectedMode = ExportMode.CSV; this.DialogResult = DialogResult.OK; this.Close(); };

            var btnExcel = CreateSuccessButton("📊  Xuất Excel", 390, 10, 170, 40, bottomBar);
            btnExcel.Click += (s, e) => { SelectedMode = ExportMode.Excel; this.DialogResult = DialogResult.OK; this.Close(); };

            var btnCAD = CreatePrimaryButton("📐  Vẽ lên CAD", 580, 10, 170, 40, bottomBar);
            btnCAD.BackColor = Color.FromArgb(180, 100, 20);
            btnCAD.FlatAppearance.BorderColor = Color.FromArgb(140, 80, 15);
            btnCAD.FlatAppearance.MouseOverBackColor = Color.FromArgb(210, 130, 40);
            btnCAD.Click += (s, e) => { SelectedMode = ExportMode.AutoCADTable; this.DialogResult = DialogResult.OK; this.Close(); };
        }

        private void LoadPreview()
        {
            if (TableData == null || TableData.Cells.Count == 0) return;

            dgvPreview.Columns.Clear();
            dgvPreview.Rows.Clear();

            // Create columns
            int colCount = TableData.ColCount;
            for (int c = 0; c < colCount; c++)
            {
                dgvPreview.Columns.Add($"col{c}", $"Cột {c + 1}");
            }

            // Fill rows
            foreach (var rowData in TableData.Cells)
            {
                object[] values = new object[colCount];
                for (int c = 0; c < colCount; c++)
                {
                    values[c] = c < rowData.Count ? rowData[c] : "";
                }
                dgvPreview.Rows.Add(values);
            }

            // Style first row as header
            if (dgvPreview.Rows.Count > 0)
            {
                dgvPreview.Rows[0].DefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
                dgvPreview.Rows[0].DefaultCellStyle.BackColor = Color.FromArgb(232, 238, 248);
            }

            lblInfo.Text = $"  📋  {TableData.RowCount} dòng × {TableData.ColCount} cột  |  Bảng: {TableData.TableName}";
            lblInfo.ForeColor = GreenOk;
        }
    }

    #endregion

    #region Command

    public class XuatBangCivil
    {
        [CommandMethod("CTP_XuatBangCivil")]
        public void CTP_XuatBangCivilCmd()
        {
            try
            {
                // 1. Chọn bảng Civil 3D
                PromptEntityOptions peo = new PromptEntityOptions("\nChọn bảng Civil 3D cần xuất: ");
                peo.SetRejectMessage("\nĐối tượng không phải bảng Civil!");
                peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Table), false);
                PromptEntityResult per = A.Ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;

                // 2. Đọc dữ liệu bảng
                CivilTableData tableData = ReadCivilTable(per.ObjectId);
                if (tableData.Cells.Count == 0)
                {
                    A.Ed.WriteMessage("\n❌ Bảng rỗng hoặc không đọc được dữ liệu!");
                    return;
                }

                // 3. Hiện form preview
                XuatBangCivilForm form = new XuatBangCivilForm(tableData);
                AcadApp.ShowModalDialog(form);

                if (form.SelectedMode == XuatBangCivilForm.ExportMode.None) return;

                // 4. Xuất theo mode đã chọn
                switch (form.SelectedMode)
                {
                    case XuatBangCivilForm.ExportMode.CSV:
                        ExportToCSV(tableData, form.CsvDelimiter, form.SelectedEncoding);
                        break;
                    case XuatBangCivilForm.ExportMode.Excel:
                        ExportToExcel(tableData);
                        break;
                    case XuatBangCivilForm.ExportMode.AutoCADTable:
                        DrawAutoCADTable(tableData, form.TextHeight, form.RowHeight);
                        break;
                }
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Đọc dữ liệu từ Civil 3D Table sử dụng Cells API
        /// </summary>
        private CivilTableData ReadCivilTable(ObjectId tableId)
        {
            CivilTableData data = new();

            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                var table = tr.GetObject(tableId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Table;
                if (table == null) { tr.Commit(); return data; }

                data.TableName = "CivilTable_" + table.Handle.ToString();
                data.RowCount = table.Rows.Count;
                data.ColCount = table.Columns.Count;

                for (int r = 0; r < table.Rows.Count; r++)
                {
                    List<string> row = new();
                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        try
                        {
                            string cellText = table.Cells[r, c].GetTextString(FormatOption.IgnoreMtextFormat);
                            // Xử lý text trống
                            row.Add(string.IsNullOrWhiteSpace(cellText) ? "" : cellText.Trim());
                        }
                        catch
                        {
                            row.Add("");
                        }
                    }
                    data.Cells.Add(row);
                }

                tr.Commit();
            }
            catch
            {
                // Fallback: Explode method (tương tự plugin gốc)
                data = ReadCivilTableByExplode(tableId);
            }

            return data;
        }

        /// <summary>
        /// Fallback: Đọc bảng bằng cách Explode (tương thích với các bảng đặc biệt)
        /// </summary>
        private CivilTableData ReadCivilTableByExplode(ObjectId tableId)
        {
            CivilTableData data = new();
            var textItems = new List<(Point3d Position, string Text)>();

            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                var table = tr.GetObject(tableId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Table;
                if (table == null) { tr.Commit(); return data; }

                data.TableName = "CivilTable_" + table.Handle.ToString();

                using DBObjectCollection exploded = new();
                table.Explode(exploded);
                foreach (Autodesk.AutoCAD.DatabaseServices.DBObject obj in exploded)
                {
                    if (obj is BlockReference blkRef)
                    {
                        using DBObjectCollection innerExploded = new();
                        blkRef.Explode(innerExploded);
                        foreach (Autodesk.AutoCAD.DatabaseServices.DBObject innerObj in innerExploded)
                        {
                            if (innerObj is MText mtext)
                            {
                                textItems.Add((mtext.Location, mtext.Text?.Trim() ?? ""));
                            }
                            innerObj.Dispose();
                        }
                    }
                    else if (obj is MText directMText)
                    {
                        textItems.Add((directMText.Location, directMText.Text?.Trim() ?? ""));
                    }
                    obj.Dispose();
                }

                tr.Commit();
            }
            catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

            if (textItems.Count == 0) return data;

            // Sắp xếp theo Y giảm dần → X tăng dần (tương tự plugin gốc)
            var sortedRows = new SortedDictionary<double, SortedDictionary<double, string>>(
                Comparer<double>.Create((a, b) => b.CompareTo(a)));

            foreach (var (pos, text) in textItems)
            {
                double y = Math.Round(pos.Y, 1);
                double x = Math.Round(pos.X, 1);
                if (!sortedRows.ContainsKey(y))
                    sortedRows[y] = new SortedDictionary<double, string>();
                sortedRows[y][x] = text;
            }

            int maxCols = sortedRows.Max(r => r.Value.Count);

            foreach (var rowPair in sortedRows)
            {
                List<string> row = new();
                foreach (var cellPair in rowPair.Value)
                {
                    row.Add(cellPair.Value);
                }
                // Pad to max columns
                while (row.Count < maxCols) row.Add("");
                data.Cells.Add(row);
            }

            data.RowCount = data.Cells.Count;
            data.ColCount = maxCols;

            return data;
        }

        /// <summary>
        /// Xuất CSV
        /// </summary>
        private void ExportToCSV(CivilTableData data, string delimiter, Encoding encoding)
        {
            SaveFileDialog sfd = new()
            {
                Title = "Lưu file CSV",
                Filter = "CSV Files (*.csv)|*.csv",
                DefaultExt = "csv",
                FileName = $"{data.TableName}.csv",
                OverwritePrompt = true
            };

            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                StringBuilder sb = new();
                foreach (var row in data.Cells)
                {
                    // Wrap text trong dấu ngoặc kép nếu chứa delimiter
                    var cells = row.Select(cell =>
                    {
                        if (cell.Contains(delimiter) || cell.Contains("\"") || cell.Contains("\n"))
                            return $"\"{cell.Replace("\"", "\"\"")}\"";
                        return cell;
                    });
                    sb.AppendLine(string.Join(delimiter, cells));
                }

                File.WriteAllText(sfd.FileName, sb.ToString(), encoding);
                A.Ed.WriteMessage($"\n✅ Đã xuất CSV: {sfd.FileName}  ({data.RowCount} dòng × {data.ColCount} cột)");

                // Hỏi mở file
                if (MessageBox.Show("Mở file CSV vừa xuất?", "Thành công",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = sfd.FileName,
                        UseShellExecute = true
                    });
                }
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi xuất CSV: {ex.Message}");
            }
        }

        /// <summary>
        /// Xuất Excel (ClosedXML)
        /// </summary>
        private void ExportToExcel(CivilTableData data)
        {
            SaveFileDialog sfd = new()
            {
                Title = "Lưu file Excel",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                FileName = $"{data.TableName}.xlsx",
                OverwritePrompt = true
            };

            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                using var workbook = new XLWorkbook();
                string sheetName = data.TableName.Length > 31 ? data.TableName.Substring(0, 31) : data.TableName;
                char[] invalidChars = { ':', '\\', '/', '?', '*', '[', ']' };
                foreach (char c in invalidChars) sheetName = sheetName.Replace(c, '_');
                if (string.IsNullOrWhiteSpace(sheetName)) sheetName = "Sheet1";

                var ws = workbook.Worksheets.Add(sheetName);

                // Ghi dữ liệu
                for (int r = 0; r < data.Cells.Count; r++)
                {
                    for (int c = 0; c < data.Cells[r].Count; c++)
                    {
                        string cellVal = data.Cells[r][c];
                        // Thử parse số
                        if (double.TryParse(cellVal, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out double numVal))
                        {
                            ws.Cell(r + 1, c + 1).Value = numVal;
                        }
                        else
                        {
                            ws.Cell(r + 1, c + 1).Value = cellVal;
                        }
                    }
                }

                // Format header row
                if (data.Cells.Count > 0)
                {
                    var headerRange = ws.Range(1, 1, 1, data.ColCount);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightSteelBlue;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                // Format all cells
                var allRange = ws.Range(1, 1, data.RowCount, data.ColCount);
                allRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                allRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                allRange.Style.Alignment.WrapText = true;

                // Auto-fit columns
                ws.Columns().AdjustToContents();

                workbook.SaveAs(sfd.FileName);
                A.Ed.WriteMessage($"\n✅ Đã xuất Excel: {sfd.FileName}  ({data.RowCount} dòng × {data.ColCount} cột)");

                // Hỏi mở file
                if (MessageBox.Show("Mở file Excel vừa xuất?", "Thành công",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = sfd.FileName,
                        UseShellExecute = true
                    });
                }
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi xuất Excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Vẽ AutoCAD Table thuần (Lines + MText) lên bản vẽ
        /// Không phụ thuộc Civil 3D — có thể mở ở AutoCAD thuần
        /// </summary>
        private void DrawAutoCADTable(CivilTableData data, double textHeight, double rowHeight)
        {
            if (data.Cells.Count == 0) return;

            // Chọn điểm chèn
            PromptPointResult ppr = A.Ed.GetPoint("\nChọn điểm chèn (góc trên trái): ");
            if (ppr.Status != PromptStatus.OK) return;

            Point3d insertPt = ppr.Value;

            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                BlockTable bt = (BlockTable)tr.GetObject(A.Db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                int rows = data.RowCount;
                int cols = data.ColCount;

                // Tính toán chiều rộng cột dựa trên nội dung
                double[] colWidths = new double[cols];
                double minColWidth = textHeight * 5;
                double maxColWidth = textHeight * 30;

                for (int c = 0; c < cols; c++)
                {
                    double maxLen = 0;
                    for (int r = 0; r < rows; r++)
                    {
                        string cell = data.GetCell(r, c);
                        double len = cell.Length * textHeight * 0.7;
                        if (len > maxLen) maxLen = len;
                    }
                    colWidths[c] = Math.Max(minColWidth, Math.Min(maxColWidth, maxLen + textHeight * 2));
                }

                double totalWidth = colWidths.Sum();
                double totalHeight = rows * rowHeight;

                // Vẽ đường ngang (horizontal lines)
                for (int r = 0; r <= rows; r++)
                {
                    double y = insertPt.Y - r * rowHeight;
                    Point3d pt1 = new Point3d(insertPt.X, y, 0);
                    Point3d pt2 = new Point3d(insertPt.X + totalWidth, y, 0);

                    Line line = new Line(pt1, pt2);
                    // Đường viền trên/dưới đậm hơn
                    if (r == 0 || r == 1 || r == rows)
                        line.LineWeight = LineWeight.LineWeight030;

                    btr.AppendEntity(line);
                    tr.AddNewlyCreatedDBObject(line, true);
                }

                // Vẽ đường dọc (vertical lines)
                double xOffset = 0;
                for (int c = 0; c <= cols; c++)
                {
                    double x = insertPt.X + xOffset;
                    Point3d pt1 = new Point3d(x, insertPt.Y, 0);
                    Point3d pt2 = new Point3d(x, insertPt.Y - totalHeight, 0);

                    Line line = new Line(pt1, pt2);
                    if (c == 0 || c == cols)
                        line.LineWeight = LineWeight.LineWeight030;

                    btr.AppendEntity(line);
                    tr.AddNewlyCreatedDBObject(line, true);

                    if (c < cols) xOffset += colWidths[c];
                }

                // Vẽ text (MText)
                xOffset = 0;
                for (int c = 0; c < cols; c++)
                {
                    for (int r = 0; r < rows; r++)
                    {
                        string cellText = data.GetCell(r, c);
                        if (string.IsNullOrWhiteSpace(cellText)) continue;

                        // Vị trí text: giữa ô
                        double x = insertPt.X + xOffset + colWidths[c] / 2;
                        double y = insertPt.Y - r * rowHeight - rowHeight / 2;

                        MText mtext = new MText();
                        mtext.Location = new Point3d(x, y, 0);
                        mtext.Contents = cellText;
                        mtext.TextHeight = textHeight;
                        mtext.Width = colWidths[c] - textHeight;
                        mtext.Attachment = AttachmentPoint.MiddleCenter;

                        // Header row bold
                        if (r == 0)
                        {
                            mtext.Contents = "{\\fSegoe UI|b1;" + cellText + "}";
                        }

                        btr.AppendEntity(mtext);
                        tr.AddNewlyCreatedDBObject(mtext, true);
                    }
                    xOffset += colWidths[c];
                }

                tr.Commit();
                A.Ed.WriteMessage($"\n✅ Đã vẽ bảng AutoCAD: {rows} dòng × {cols} cột tại ({insertPt.X:F1}, {insertPt.Y:F1})");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi vẽ bảng: {ex.Message}");
            }
        }
    }

    #endregion
}
