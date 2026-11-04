using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using ClosedXML.Excel;
using System.Windows.Forms;

namespace MyFirstProject.XuatPhanMemTinhToan
{
    public static class ExcelExportEngine
    {
        public static void ExportMathCADReport(string title, List<MathCADReportGroup> groups, bool exportPdf = false, string projectName = "", string location = "", string item = "")
        {
            try
            {
                using (var sfd = new SaveFileDialog())
                {
                    if (exportPdf)
                    {
                        sfd.Filter = "PDF Files|*.pdf";
                        sfd.Title = "Lưu File Báo Cáo Kỹ Thuật (PDF)";
                        sfd.FileName = $"BaoCao_{title.Replace(" ", "_")}.pdf";
                    }
                    else
                    {
                        sfd.Filter = "Excel Files|*.xlsx";
                        sfd.Title = "Lưu File Báo Cáo Kỹ Thuật";
                        sfd.FileName = $"BaoCao_{title.Replace(" ", "_")}.xlsx";
                    }

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        // Luôn tạo file Excel tạm nếu cần xuất PDF
                        string excelPath = exportPdf ? Path.ChangeExtension(sfd.FileName, ".xlsx") : sfd.FileName;

                        using (var workbook = new XLWorkbook())
                        {
                            var ws = workbook.Worksheets.Add(title.Length > 31 ? title.Substring(0, 31) : title);

                            // Title
                            ws.Cell("A1").Value = title.ToUpper();
                            ws.Range("A1:E1").Merge().Style
                                .Font.SetFontSize(16)
                                .Font.SetBold()
                                .Font.SetFontColor(XLColor.DarkBlue)
                                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                            int r = 3;

                            if (!string.IsNullOrEmpty(projectName) || !string.IsNullOrEmpty(location) || !string.IsNullOrEmpty(item))
                            {
                                if (!string.IsNullOrEmpty(projectName))
                                {
                                    ws.Cell(r, 1).Value = "Dự án: " + projectName;
                                    ws.Range(r, 1, r, 5).Merge().Style.Font.SetBold().Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                                    r++;
                                }
                                if (!string.IsNullOrEmpty(location))
                                {
                                    ws.Cell(r, 1).Value = "Địa điểm: " + location;
                                    ws.Range(r, 1, r, 5).Merge().Style.Font.SetItalic().Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                                    r++;
                                }
                                if (!string.IsNullOrEmpty(item))
                                {
                                    ws.Cell(r, 1).Value = "Hạng mục: " + item;
                                    ws.Range(r, 1, r, 5).Merge().Style.Font.SetBold().Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                                    r++;
                                }
                                r++; // Add an empty row before headers
                            }

                            // Header
                            string[] headers = { "Tên thông số", "Ký hiệu", "Công thức", "Giá trị", "Đơn vị" };
                            for (int c = 0; c < headers.Length; c++)
                            {
                                var cell = ws.Cell(r, c + 1);
                                cell.Value = headers[c];
                                cell.Style.Font.SetBold()
                                          .Font.SetFontColor(XLColor.White)
                                          .Fill.SetBackgroundColor(XLColor.FromHtml("#1F4E79")) // Dark Blue Premium
                                          .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                                          .Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                                          .Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                            }

                            r++;

                            // Content
                            foreach (var group in groups)
                            {
                                if (!string.IsNullOrEmpty(group.GroupName))
                                {
                                    ws.Cell(r, 1).Value = group.GroupName;
                                    ws.Range(r, 1, r, 5).Merge().Style
                                        .Font.SetBold()
                                        .Font.SetItalic()
                                        .Fill.SetBackgroundColor(XLColor.FromHtml("#DDEBF7"))
                                        .Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                                    r++;
                                }

                                foreach (var param in group.Parameters)
                                {
                                    ws.Cell(r, 1).Value = param.TenThongSo;
                                    ws.Cell(r, 2).Value = param.KyHieu;
                                    ws.Cell(r, 3).Value = param.CongThuc;
                                    ws.Cell(r, 4).Value = param.GiaTri;
                                    ws.Cell(r, 5).Value = param.DonVi;

                                    // Styling Row
                                    ws.Range(r, 1, r, 5).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                                    ws.Range(r, 2, r, 5).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                    
                                    // Math Font for formulas
                                    ws.Cell(r, 2).Style.Font.SetFontName("Cambria Math").Font.SetItalic();
                                    ws.Cell(r, 3).Style.Font.SetFontName("Cambria Math");

                                    // Number alignment
                                    ws.Cell(r, 4).Style.Font.SetBold().Font.SetFontColor(XLColor.FromHtml("#C00000")); // Accent Red for focus
                                    
                                    r++;
                                }
                            }

                            // Auto fit
                            ws.Columns().AdjustToContents();
                            ws.Column(1).Width = Math.Max(ws.Column(1).Width, 35); // Ensure name column is wide enough
                            ws.Column(3).Width = Math.Max(ws.Column(3).Width, 30); // Formula column

                            // Page margins and landscape for PDF
                            ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
                            ws.PageSetup.FitToPages(1, 0);

                            // Save
                            workbook.SaveAs(excelPath);
                        }

                        if (exportPdf)
                        {
                            try
                            {
                                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                                if (excelType == null) throw new Exception("Không tìm thấy Microsoft Excel trên máy tính. Cần cài đặt Excel để sử dụng tính năng xuất thẳng sang PDF.");
                                
                                dynamic excelObj = Activator.CreateInstance(excelType)!;
                                excelObj.Visible = false;
                                dynamic workbooks = excelObj.Workbooks;
                                dynamic workbookObj = workbooks.Open(excelPath);
                                // xlTypePDF = 0
                                workbookObj.ExportAsFixedFormat(0, sfd.FileName);
                                workbookObj.Close(false);
                                excelObj.Quit();
                                
                                // Clean up temp file
                                if (File.Exists(excelPath)) File.Delete(excelPath);
                                
                                MessageBox.Show("Xuất file PDF thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                Process.Start(new ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Trần Excel Sang PDF lỗi: {ex.Message}", "Lỗi Excel Interop", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Xuất file Excel thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            Process.Start(new ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi xuất file: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
