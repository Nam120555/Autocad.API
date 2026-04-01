using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Windows;
using Helpers = MyFirstProject.Helpers;

[assembly: CommandClass(typeof(MyFirstProject.Autocad))]

namespace MyFirstProject
{
    /// <summary>
    /// Helper class for creating ribbon button icons
    /// </summary>
    public static class RibbonIconHelper
    {
        private static readonly Dictionary<string, BitmapImage> _iconCache = [];

        /// <summary>
        /// Tạo icon từ text/emoji với màu sắc tùy chỉnh (Nền trong suốt kiểu Civil 3D)
        /// </summary>
        public static BitmapImage? CreateTextIcon(string text, int size = 32, System.Drawing.Color? color = null)
        {
            if (string.IsNullOrEmpty(text)) return null;
            
            string iconText = ExtractIconChar(text);
            color ??= System.Drawing.Color.Black;

            string cacheKey = $"{iconText}_{size}_{color.Value.ToArgb()}";
            if (_iconCache.TryGetValue(cacheKey, out var cached))
                return cached;

            try
            {
                using var bitmap = new Bitmap(size, size);
                using var graphics = Graphics.FromImage(bitmap);
                
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                graphics.Clear(System.Drawing.Color.Transparent);
                
                float fontSize = size * 0.7f;
                using var font = new System.Drawing.Font("Segoe UI Symbol", fontSize, System.Drawing.FontStyle.Bold);
                using var textBrush = new SolidBrush(color.Value);
                
                var textSize = graphics.MeasureString(iconText, font);
                float x = (size - textSize.Width) / 2;
                float y = (size - textSize.Height) / 2;
                
                using (var shadowBrush = new SolidBrush(System.Drawing.Color.FromArgb(100, 100, 100)))
                {
                    graphics.DrawString(iconText, font, shadowBrush, x + 1, y + 1);
                }
                
                graphics.DrawString(iconText, font, textBrush, x, y);

                var bitmapImage = ConvertToBitmapImage(bitmap);
                if (bitmapImage != null) _iconCache[cacheKey] = bitmapImage;
                return bitmapImage;
            }
            catch { return null; }
        }

        private static string ExtractIconChar(string text)
        {
            if (string.IsNullOrEmpty(text)) return "?";
            var iconChars = new[] { "◎", "▸", "◈", "◊", "═", "▤", "□", "◻", "▢", "▣", 
                                    "⊙", "⊚", "⊕", "⊗", "★", "☆", "●", "○", "◐", "◑",
                                    "▶", "◀", "▲", "▼", "◄", "►", "◆", "◇", "⬇", "⬆",
                                    "━", "─", "│", "║", "╱", "╲", "▭", "╋", "△", "⊞", "▥" };
            
            foreach (var c in iconChars) if (text.Contains(c)) return c;
            foreach (char c in text) if (char.IsLetterOrDigit(c)) return c.ToString().ToUpper();
            return "?";
        }

        private static BitmapImage? ConvertToBitmapImage(Bitmap bitmap)
        {
            try
            {
                using var memory = new MemoryStream();
                bitmap.Save(memory, ImageFormat.Png);
                memory.Position = 0;
                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                bitmapImage.Freeze();
                return bitmapImage;
            }
            catch { return null; }
        }

        public static BitmapImage? CreateColorIcon(string text, System.Drawing.Color bgColor, int size = 32)
        {
            return CreateTextIcon(text, size, bgColor);
        }
    }

    public class Autocad
    {
        [CommandMethod("ShowForm")]
        public static void ShowForm()
        {
            TestForm frmTest = new();
            frmTest.Show();
        }

        [CommandMethod("AdskGreeting")]
        public static void AdskGreeting()
        {
            Document? acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ?? throw new InvalidOperationException("No active document found.");
            Database? acCurDb = acDoc.Database ?? throw new InvalidOperationException("No database found for the active document.");

            using Transaction acTrans = acCurDb.TransactionManager.StartTransaction();
            BlockTable acBlkTbl = (BlockTable)acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) ?? throw new InvalidOperationException("BlockTable could not be retrieved.");
            BlockTableRecord acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) ?? throw new InvalidOperationException("BlockTableRecord could not be retrieved.");

            using (MText objText = new())
            {
                objText.Location = new Autodesk.AutoCAD.Geometry.Point3d(2, 2, 0);
                objText.Contents = "Greetings, Welcome to AutoCAD .NET";
                objText.TextStyleId = acCurDb.Textstyle;
                acBlkTblRec.AppendEntity(objText);
                acTrans.AddNewlyCreatedDBObject(objText, true);
            }
            acTrans.Commit();
        }

        [CommandMethod("show_menu")]
        public static void ShowMenu()
        {
            // Redirect sang CIVIL_RIBBON (19.CivilRibbonTaskbar.cs)
            // Đã gộp tất cả Civil + ACAD + VisualINFRA vào 1 tab duy nhất
            CivilRibbonTaskbar.ShowCivilRibbon();
        }

        private class SimpleRibbonCommandHandler : System.Windows.Input.ICommand
        {
            public bool CanExecute(object? parameter) => true;

            public event EventHandler? CanExecuteChanged { add { } remove { } }

            public void Execute(object? parameter)
            {
                try
                {
                    string? commandToRun = null;
                    
                    if (parameter is string cmd)
                    {
                        commandToRun = cmd;
                    }
                    else if (parameter is RibbonButton rb)
                    {
                        commandToRun = rb.CommandParameter as string;
                    }

                    if (string.IsNullOrWhiteSpace(commandToRun)) return;

                    var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    if (doc != null)
                    {
                        doc.SendStringToExecute(commandToRun + " ", true, false, true);
                    }
                }
                catch (System.Exception ex)
                {
                    var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor;
                    ed?.WriteMessage($"\nLỗi thực thi lệnh: {ex.Message}");
                }
            }
        }
    }
}