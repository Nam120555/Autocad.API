using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Colors;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace MyFirstProject.OutLook
{
    public class XuatLink_OutLook
    {
        private const string OutLookFolder = @"C:\OutLook";
        private const string BatFilePath = @"C:\OutLook\CapNhat_Email_TuDong.bat";
        private const string Ps1FilePath = @"C:\OutLook\sync_emails.ps1";
        private const string PrimaryExcelPath = @"C:\OutLook\ThongKe_Email_QLTKDS01_V3_Linked.xlsm";
        private const string SecondaryExcelPath = @"C:\OutLook\ThongKe_Email_QLTKDS01_ALL.xlsx";
        private const string MsgArchiveFolder = @"C:\OutLook\SavedEmails_Archive";
        private const string LiveUpdatePs1Path = @"C:\OutLook\update_data_live.ps1";
        public const string OstFilePath = @"C:\Users\NAM\AppData\Local\Microsoft\Outlook\sgc.qltkds01@sgc.net.vn.ost";

        /// <summary>
        /// AutoCAD Command: XUATLINK_OUTLOOK / UPDATE_EXCEL_MAIL
        /// quet toan bo hom thu Outlook va cap nhat sang file Excel Dashboard.
        /// </summary>
        [CommandMethod("XUATLINK_OUTLOOK")]
        [CommandMethod("UPDATE_EXCEL_MAIL")]
        public static void RunSyncOutlookMail()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;

            try
            {
                ed.WriteMessage("\n=======================================================");
                ed.WriteMessage("\n  [CIVIL TOOL] ĐANG CẬP NHẬT DỮ LIỆU EMAIL TỪ HÒM THƯ OUTLOOK...");
                ed.WriteMessage("\n=======================================================\n");

                EnsureScriptFilesExist(ed);

                if (!File.Exists(BatFilePath))
                {
                    ed.WriteMessage("\n[LỖI] Không thể tạo hoặc tìm thấy file script tại: " + BatFilePath);
                    return;
                }

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = "/c \"" + BatFilePath + "\"",
                    UseShellExecute = true,
                    CreateNoWindow = false,
                    WorkingDirectory = OutLookFolder
                };

                ed.WriteMessage("\n Tiến trình cập nhật email tự động đang chạy trong cửa sổ cmd...");
                Process proc = Process.Start(psi);

                if (proc != null)
                {
                    proc.WaitForExit();
                    ed.WriteMessage("\n Tiến trình quét & đồng bộ email hoàn tất.");
                }

                string targetExcel = GetLatestExcelPath();
                if (File.Exists(targetExcel))
                {
                    ed.WriteMessage("\n File Excel kết quả: " + targetExcel);

                    PromptKeywordOptions pko = new PromptKeywordOptions("\nBạn có muốn mở file Excel báo cáo ngay bây giờ không?");
                    pko.Keywords.Add("Co");
                    pko.Keywords.Add("Khong");
                    pko.Keywords.Default = "Co";
                    pko.AllowNone = true;

                    PromptResult pr = ed.GetKeywords(pko);
                    if (pr.Status == PromptStatus.OK && (pr.StringResult == "Co" || string.IsNullOrEmpty(pr.StringResult)))
                    {
                        Process.Start(new ProcessStartInfo(targetExcel) { UseShellExecute = true });
                    }
                }
                else
                {
                    ed.WriteMessage("\n[CẢNH BÁO] Không tìm thấy file Excel kết quả sau khi chạy script.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\n[LỖI] Không thể chạy cập nhật: " + ex.Message);
            }
        }

        /// <summary>
        /// AutoCAD Command: MO_EXCEL_MAIL / OPEN_EXCEL_MAIL
        /// Mở trực tiếp file Excel báo cáo thống kê email.
        /// </summary>
        [CommandMethod("MO_EXCEL_MAIL")]
        [CommandMethod("OPEN_EXCEL_MAIL")]
        public static void OpenExcelMailReport()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;

            try
            {
                string targetExcel = GetLatestExcelPath();
                if (File.Exists(targetExcel))
                {
                    ed.WriteMessage("\n[CIVIL TOOL] Đang mở file Excel báo cáo: " + targetExcel);
                    Process.Start(new ProcessStartInfo(targetExcel) { UseShellExecute = true });
                }
                else
                {
                    ed.WriteMessage("\n[LỖI] Chưa tìm thấy file Excel báo cáo email tại " + OutLookFolder);
                    ed.WriteMessage("\nHãy gõ lệnh 'XUATLINK_OUTLOOK' hoặc 'UPDATE_EXCEL_MAIL' để khởi tạo dữ liệu.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\n[LỖI] Không thể mở file Excel: " + ex.Message);
            }
        }

        /// <summary>
        /// AutoCAD Command: CHEN_LINK_MAIL_VAO_CAD / LINK_MAIL_CAD
        /// Gán Hyperlink email Outlook (.msg hoặc outlook:ID) vào đối tượng được chọn trong bản vẽ.
        /// </summary>
        [CommandMethod("CHEN_LINK_MAIL_VAO_CAD")]
        [CommandMethod("LINK_MAIL_CAD")]
        public static void LinkEmailToCadEntity()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                PromptEntityOptions peo = new PromptEntityOptions("\nChọn đối tượng CAD (Text, MText, Block, Polyline) để đính kèm link Email:");
                PromptEntityResult per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;

                string defaultPath = "";
                if (Directory.Exists(MsgArchiveFolder))
                {
                    var files = Directory.GetFiles(MsgArchiveFolder, "*.msg");
                    if (files.Length > 0)
                    {
                        Array.Sort(files, (a, b) => File.GetLastWriteTime(b).CompareTo(File.GetLastWriteTime(a)));
                        defaultPath = files[0];
                    }
                }

                PromptStringOptions pso = new PromptStringOptions("\nNhập đường dẫn file Mail (.msg) hoặc mã EntryID (outlook:EntryID):");
                pso.AllowSpaces = true;
                if (!string.IsNullOrEmpty(defaultPath))
                {
                    pso.DefaultValue = defaultPath;
                    pso.UseDefaultValue = true;
                }

                PromptResult pr = ed.GetString(pso);
                if (pr.Status != PromptStatus.OK || string.IsNullOrWhiteSpace(pr.StringResult)) return;

                string linkTarget = pr.StringResult.Trim();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Entity ent = tr.GetObject(per.ObjectId, OpenMode.ForWrite) as Entity;
                    if (ent != null)
                    {
                        Autodesk.AutoCAD.DatabaseServices.HyperLink hl = new Autodesk.AutoCAD.DatabaseServices.HyperLink();
                        hl.Name = linkTarget;
                        hl.Description = "Mở Email Outlook: " + Path.GetFileName(linkTarget);

                        ent.Hyperlinks.Add(hl);
                        ed.WriteMessage("\n[THÀNH CÔNG] Đã đính kèm link Mail thành công vào đối tượng " + ent.GetType().Name);
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\n[LỖI] Gán link email thất bại: " + ex.Message);
            }
        }

        /// <summary>
        /// AutoCAD Command: XUAT_BANG_EMAIL_CAD / IMPORT_EMAIL_TABLE_TO_CAD
        /// Vẽ bảng thống kê email mới nhất trực tiếp trên không gian bản vẽ AutoCAD.
        /// </summary>
        [CommandMethod("XUAT_BANG_EMAIL_CAD")]
        [CommandMethod("IMPORT_EMAIL_TABLE_TO_CAD")]
        public static void ImportEmailTableToCad()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                PromptPointOptions ppo = new PromptPointOptions("\nChọn điểm chèn bảng thống kê Email:");
                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status != PromptStatus.OK) return;

                Point3d insertPt = ppr.Value;

                var emailList = GetRecentSavedEmails(10);
                if (emailList.Count == 0)
                {
                    ed.WriteMessage("\n[THÔNG BÁO] Chưa có dữ liệu file .msg tại " + MsgArchiveFolder);
                    ed.WriteMessage("\nHãy gõ lệnh 'XUATLINK_OUTLOOK' để tự động quét & tạo lưu trữ email.");
                    return;
                }

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    Table tb = new Table();
                    tb.TableStyle = db.Tablestyle;
                    tb.SetSize(emailList.Count + 2, 5);
                    tb.Position = insertPt;

                    tb.Columns[0].Width = 10;
                    tb.Columns[1].Width = 25;
                    tb.Columns[2].Width = 35;
                    tb.Columns[3].Width = 70;
                    tb.Columns[4].Width = 25;

                    // Title Row
                    tb.Cells[0, 0].TextString = "BẢNG THỐNG KÊ EMAIL OUTLOOK DỰ ÁN MỚI NHẤT";
                    tb.Cells[0, 0].TextHeight = 3.5;
                    tb.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;
                    tb.Cells[0, 0].BackgroundColor = Autodesk.AutoCAD.Colors.Color.FromRgb(24, 43, 73);
                    tb.Cells[0, 0].ContentColor = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);

                    // Header Row
                    string[] headers = new string[] { "STT", "Ngày sửa", "Tên File Email", "Đường dẫn chi tiết", "Định dạng" };
                    for (int c = 0; c < headers.Length; c++)
                    {
                        tb.Cells[1, c].TextString = headers[c];
                        tb.Cells[1, c].TextHeight = 2.5;
                        tb.Cells[1, c].Alignment = CellAlignment.MiddleCenter;
                        tb.Cells[1, c].BackgroundColor = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 112, 192);
                        tb.Cells[1, c].ContentColor = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);
                    }

                    // Data Rows
                    for (int i = 0; i < emailList.Count; i++)
                    {
                        int row = i + 2;
                        var fileInfo = emailList[i];

                        tb.Cells[row, 0].TextString = (i + 1).ToString();
                        tb.Cells[row, 1].TextString = fileInfo.LastWriteTime.ToString("dd/MM/yyyy HH:mm");
                        tb.Cells[row, 2].TextString = fileInfo.Name;
                        tb.Cells[row, 3].TextString = fileInfo.FullName;
                        tb.Cells[row, 4].TextString = fileInfo.Extension.ToUpper();

                        for (int c = 0; c < 5; c++)
                        {
                            tb.Cells[row, c].TextHeight = 2.0;
                            tb.Cells[row, c].Alignment = (c == 0 || c == 1 || c == 4) ? CellAlignment.MiddleCenter : CellAlignment.MiddleLeft;
                        }
                    }

                    tb.GenerateLayout();
                    btr.AppendEntity(tb);
                    tr.AddNewlyCreatedDBObject(tb, true);
                    tr.Commit();

                    ed.WriteMessage("\n[THÀNH CÔNG] Đã tạo bảng thống kê " + emailList.Count + " email mới nhất trong bản vẽ!");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\n[LỖI] Tạo bảng email thất bại: " + ex.Message);
            }
        }

        #region Helper Methods

        private static string GetLatestExcelPath()
        {
            if (File.Exists(PrimaryExcelPath)) return PrimaryExcelPath;
            if (File.Exists(SecondaryExcelPath)) return SecondaryExcelPath;
            return PrimaryExcelPath;
        }

        private static List<FileInfo> GetRecentSavedEmails(int topCount)
        {
            List<FileInfo> list = new List<FileInfo>();
            if (Directory.Exists(MsgArchiveFolder))
            {
                DirectoryInfo di = new DirectoryInfo(MsgArchiveFolder);
                list = di.GetFiles("*.msg")
                          .OrderByDescending(f => f.LastWriteTime)
                          .Take(topCount)
                          .ToList();
            }
            return list;
        }

        private static void EnsureScriptFilesExist(Editor ed)
        {
            try
            {
                if (!Directory.Exists(OutLookFolder))
                {
                    Directory.CreateDirectory(OutLookFolder);
                    ed.WriteMessage("\n[CIVIL TOOL] Đã khởi tạo thư mục: " + OutLookFolder);
                }

                if (!File.Exists(BatFilePath))
                {
                    string batContent = "@echo off\r\n" +
                                        "chcp 65001 > NUL\r\n" +
                                        "title Cập Nhật Email Tự Động Từ Outlook\r\n" +
                                        "echo ========================================================\r\n" +
                                        "echo   ĐANG QUÉT VÀ CẬP NHẬT EMAIL MỚI NHẤT TỪ OUTLOOK...\r\n" +
                                        "echo ========================================================\r\n" +
                                        "echo.\r\n" +
                                        "powershell -ExecutionPolicy Bypass -File \"C:\\OutLook\\sync_emails.ps1\"\r\n" +
                                        "echo.\r\n" +
                                        "echo Hoàn thành! Nhấn phím bất kỳ để đóng cửa sổ này.\r\n" +
                                        "pause > NUL\r\n";
                    File.WriteAllText(BatFilePath, batContent, System.Text.Encoding.UTF8);
                    ed.WriteMessage("\n[CIVIL TOOL] Đã tự tạo file bat: " + BatFilePath);
                }

                // Tạo file update_data_live.ps1 nếu chưa có
                if (!File.Exists(LiveUpdatePs1Path))
                {
                    // Copy từ thư mục plugin nếu có, hoặc tạo stub
                    string pluginSource = System.Reflection.Assembly.GetExecutingAssembly().Location;
                    string pluginDir = Path.GetDirectoryName(pluginSource);
                    string bundledLivePs1 = Path.Combine(pluginDir, "OutLook", "update_data_live.ps1");

                    if (File.Exists(bundledLivePs1))
                    {
                        File.Copy(bundledLivePs1, LiveUpdatePs1Path, true);
                        ed.WriteMessage("\n[CIVIL TOOL] Đã tạo file live update: " + LiveUpdatePs1Path);
                    }
                    else
                    {
                        ed.WriteMessage("\n[CẢNH BÁO] Chưa tìm thấy file update_data_live.ps1 đi kèm.");
                        ed.WriteMessage("\n  Hãy copy file update_data_live.ps1 vào thư mục C:\\OutLook\\");
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\n[CẢNH BÁO] Lỗi kiểm tra script file: " + ex.Message);
            }
        }

        #endregion
    }
}
