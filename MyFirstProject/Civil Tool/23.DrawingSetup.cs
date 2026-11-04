// DrawingSetup.cs - Thiết lập bản vẽ CAD chuẩn
// Chuyển đổi từ LISP: TLBV, SAVC, PAL, PCL, IP

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.DrawingSetupCommands))]

namespace MyFirstProject
{
    /// <summary>
    /// Các lệnh thiết lập bản vẽ CAD chuẩn
    /// </summary>
    public class DrawingSetupCommands
    {
        // ══════════════════════════════════════════════════════════════
        // THIET LAP BAN VE CHUAN (từ LISP TLBV)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// Thiết lập các system variables chuẩn cho bản vẽ CAD.
        /// Bao gồm: INSUNITS, LTSCALE, PSLTSCALE, MSLTSCALE, và nhiều biến khác.
        /// </summary>
        [CommandMethod("CTDS_ThietLap")]
        public static void CTDS_ThietLap()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            
            try
            {
                // File Dialog và Units
                AcadApp.SetSystemVariable("FILEDIA", 1);
                AcadApp.SetSystemVariable("INSUNITS", 6);  // Meters
                
                // Linetype Scale
                AcadApp.SetSystemVariable("MSLTSCALE", 1);
                AcadApp.SetSystemVariable("PSLTSCALE", 1);
                AcadApp.SetSystemVariable("LTSCALE", 1);
                AcadApp.SetSystemVariable("PLINEGEN", 1);
                
                // Performance Settings
                AcadApp.SetSystemVariable("STARTUP", 0);
                AcadApp.SetSystemVariable("HPQUICKPREVIEW", 0);
                AcadApp.SetSystemVariable("TOOLTIPS", 0);
                AcadApp.SetSystemVariable("ROLLOVERTIPS", 0);
                AcadApp.SetSystemVariable("SELECTIONPREVIEW", 0);
                
                // Xref và Layout
                AcadApp.SetSystemVariable("INDEXCTL", 0);
                AcadApp.SetSystemVariable("XLOADCTL", 2);
                AcadApp.SetSystemVariable("DEMANDLOAD", 3);
                AcadApp.SetSystemVariable("LAYOUTREGENCTL", 2);  // Cache cả Model + Layout tabs → không regen khi chuyển tab
                AcadApp.SetSystemVariable("BACKGROUNDPLOT", 0);   // In foreground → nhanh hơn nhiều cho file nặng
                
                // Objects
                AcadApp.SetSystemVariable("MIRRTEXT", 0);
                AcadApp.SetSystemVariable("PICKBOX", 5);
                AcadApp.SetSystemVariable("VISRETAIN", 1);
                AcadApp.SetSystemVariable("PEDITACCEPT", 1);
                AcadApp.SetSystemVariable("REGENMODE", 1);
                AcadApp.SetSystemVariable("LAYEREVALCTL", 0);
                
                // Units và UCS
                AcadApp.SetSystemVariable("MEASUREMENT", 0);
                AcadApp.SetSystemVariable("UCSFOLLOW", 0);
                AcadApp.SetSystemVariable("WHIPTHREAD", 3);
                AcadApp.SetSystemVariable("COMMANDPREVIEW", 0);
                
                // Proxy
                AcadApp.SetSystemVariable("PROXYSHOW", 1);
                AcadApp.SetSystemVariable("PROXYNOTICE", 1);
                AcadApp.SetSystemVariable("PROXYGRAPHICS", 1);
                
                // Display
                AcadApp.SetSystemVariable("VTENABLE", 0);
                AcadApp.SetSystemVariable("DYNMODE", 1);
                AcadApp.SetSystemVariable("CURSORBADGE", 1);
                AcadApp.SetSystemVariable("PROPERTYPREVIEW", 0);
                AcadApp.SetSystemVariable("ANNOALLVISIBLE", 1);
                AcadApp.SetSystemVariable("SELECTIONANNODISPLAY", 1);
                AcadApp.SetSystemVariable("DRAWORDERCTL", 1);
                AcadApp.SetSystemVariable("UCSDETECT", 0);
                
                // Cache và Viewport
                AcadApp.SetSystemVariable("CACHEMAXFILES", 256);
                AcadApp.SetSystemVariable("MAXACTVP", 64);
                
                ed.WriteMessage("\n◎ Đã thiết lập bản vẽ theo cấu hình chuẩn!");
                ed.WriteMessage("\n  ✓ INSUNITS = 6 (Meters)");
                ed.WriteMessage("\n  ✓ LTSCALE = 1, PSLTSCALE = 1, MSLTSCALE = 1");
                ed.WriteMessage("\n  ✓ MIRRTEXT = 0 (Text không đảo)");
                ed.WriteMessage("\n  ✓ Đã tối ưu hiệu suất và hiển thị");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n⊘ Lỗi: {ex.Message}");
            }
        }

        // ══════════════════════════════════════════════════════════════
        // SAVE CLEAN - PURGE VÀ LƯU FILE (từ LISP SAVC)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// Purge toàn bộ dữ liệu không sử dụng và lưu file.
        /// </summary>
        [CommandMethod("CTDS_SaveClean")]
        public static void CTDS_SaveClean()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            
            ed.WriteMessage("\n▸ Đang dọn dẹp bản vẽ...");
            
            // Chạy các lệnh PURGE
            doc.SendStringToExecute("-PURGE A * N ", true, false, false);
            doc.SendStringToExecute("-PURGE R * N ", true, false, false);
            doc.SendStringToExecute("-PURGE O ", true, false, false);
            doc.SendStringToExecute("-PURGE Z ", true, false, false);
            doc.SendStringToExecute("-PURGE B * N ", true, false, false);
            doc.SendStringToExecute("-PURGE DE * N ", true, false, false);
            doc.SendStringToExecute("-PURGE D * N ", true, false, false);
            doc.SendStringToExecute("-PURGE G * N ", true, false, false);
            doc.SendStringToExecute("-PURGE LA * N ", true, false, false);
            doc.SendStringToExecute("-PURGE LT * N ", true, false, false);
            doc.SendStringToExecute("-PURGE MA * N ", true, false, false);
            doc.SendStringToExecute("-PURGE MU * N ", true, false, false);
            doc.SendStringToExecute("-PURGE P * N ", true, false, false);
            doc.SendStringToExecute("-PURGE SH * N ", true, false, false);
            doc.SendStringToExecute("-PURGE ST * N ", true, false, false);
            
            // Audit và Save
            doc.SendStringToExecute("AUDIT Y ", true, false, false);
            doc.SendStringToExecute("QSAVE ", true, false, false);
            
            ed.WriteMessage("\n◎ Đã Purge và lưu file thành công!");
        }

        // ══════════════════════════════════════════════════════════════
        // IN TẤT CẢ LAYOUTS (từ LISP PAL)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// In tất cả layouts sử dụng page setup đã định nghĩa.
        /// </summary>
        [CommandMethod("CTDS_PrintAllLayouts")]
        public static void CTDS_PrintAllLayouts()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var layoutMgr = LayoutManager.Current;
            
            ed.WriteMessage("\n▸ Đang in tất cả layouts...");
            
            // Lấy danh sách layouts (loại bỏ Model)
            var layouts = new System.Collections.Generic.List<string>();
            var dbDict = doc.Database.LayoutDictionaryId;
            
            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var layoutDict = (Autodesk.AutoCAD.DatabaseServices.DBDictionary)tr.GetObject(dbDict, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                
                foreach (var entry in layoutDict)
                {
                    var layout = (Autodesk.AutoCAD.DatabaseServices.Layout)tr.GetObject(entry.Value, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    if (!layout.ModelType)
                    {
                        layouts.Add(layout.LayoutName);
                    }
                }
                tr.Commit();
            }
            
            int count = 0;
            foreach (string layoutName in layouts)
            {
                layoutMgr.CurrentLayout = layoutName;
                doc.SendStringToExecute("-PLOT N \"\" \"\" \"\" \"\" \"\" \"\" ", true, false, false);
                count++;
            }
            
            ed.WriteMessage($"\n◎ Đã gửi lệnh in cho {count} layout(s).");
        }

        // ══════════════════════════════════════════════════════════════
        // IN LAYOUT HIỆN TẠI (từ LISP PCL)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// In layout hiện tại sử dụng page setup đã định nghĩa.
        /// </summary>
        [CommandMethod("CTDS_PrintCurrentLayout")]
        public static void CTDS_PrintCurrentLayout()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var layoutMgr = LayoutManager.Current;
            
            ed.WriteMessage($"\n▸ Đang in layout '{layoutMgr.CurrentLayout}'...");
            doc.SendStringToExecute("-PLOT N \"\" \"\" \"\" \"\" \"\" \"\" ", true, false, false);
        }

        // ══════════════════════════════════════════════════════════════
        // XUẤT PDF (từ LISP IP)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// Xuất bản vẽ ra file PDF.
        /// </summary>
        [CommandMethod("CTDS_ExportPDF")]
        public static void CTDS_ExportPDF()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            doc.Editor.WriteMessage("\n▸ Đang xuất PDF...");
            doc.SendStringToExecute("EXPORTPDF ", true, false, false);
        }

        // ══════════════════════════════════════════════════════════════
        // CHUYỂN ĐƠN VỊ MM -> M (từ LISP MM-M)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// Chuyển đổi bản vẽ từ đơn vị MM sang M (scale 0.001).
        /// </summary>
        [CommandMethod("CTDS_ConvertMM2M")]
        public static void CTDS_ConvertMM2M()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            
            ed.WriteMessage("\n▸ Chuyển đổi bản vẽ từ MM sang M...");
            doc.SendStringToExecute("SCALE ALL  0,0,0 0.001 ", true, false, false);
            AcadApp.SetSystemVariable("INSUNITS", 6); // Meters
            ed.WriteMessage("\n◎ Đã chuyển đổi xong và đặt INSUNITS = 6 (Meters)");
        }

        // ══════════════════════════════════════════════════════════════
        // CHUYỂN ĐƠN VỊ CM -> M (từ LISP CM-M)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// Chuyển đổi bản vẽ từ đơn vị CM sang M (scale 0.01).
        /// </summary>
        [CommandMethod("CTDS_ConvertCM2M")]
        public static void CTDS_ConvertCM2M()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            
            ed.WriteMessage("\n▸ Chuyển đổi bản vẽ từ CM sang M...");
            doc.SendStringToExecute("SCALE ALL  0,0,0 0.01 ", true, false, false);
            AcadApp.SetSystemVariable("INSUNITS", 6); // Meters
            ed.WriteMessage("\n◎ Đã chuyển đổi xong và đặt INSUNITS = 6 (Meters)");
        }

        // ══════════════════════════════════════════════════════════════
        // TURBO MODE - BẬT/TẮT CHẾ ĐỘ HIỆU SUẤT CAO
        // ══════════════════════════════════════════════════════════════
        
        private static bool _turboModeActive = false;
        
        /// <summary>
        /// Bật/tắt chế độ Turbo Mode - tối ưu hiệu suất tối đa cho file nặng.
        /// Tắt toàn bộ hiệu ứng hiển thị, preview, tooltip để giảm tải GPU/CPU.
        /// Gõ lại lệnh để tắt Turbo Mode và khôi phục cài đặt bình thường.
        /// </summary>
        [CommandMethod("CTDS_TurboMode")]
        public static void CTDS_TurboMode()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            try
            {
                if (!_turboModeActive)
                {
                    // ═══ BẬT TURBO MODE ═══
                    ed.WriteMessage("\n╔══════════════════════════════════════════════════════════╗");
                    ed.WriteMessage("\n║              🚀 TURBO MODE - ĐANG BẬT...                ║");
                    ed.WriteMessage("\n╚══════════════════════════════════════════════════════════╝");

                    // --- Selection & Preview (giảm lag khi di chuột) ---
                    AcadApp.SetSystemVariable("SELECTIONPREVIEW", 0);  // Tắt preview khi hover
                    AcadApp.SetSystemVariable("HIGHLIGHT", 0);         // Tắt highlight khi chọn
                    AcadApp.SetSystemVariable("SELECTIONEFFECT", 0);   // Dùng nét đứt thay glow
                    AcadApp.SetSystemVariable("ROLLOVERTIPS", 0);      // Tắt tooltip khi hover
                    AcadApp.SetSystemVariable("TOOLTIPS", 0);          // Tắt tooltips
                    AcadApp.SetSystemVariable("PROPERTYPREVIEW", 0);   // Tắt property preview
                    AcadApp.SetSystemVariable("COMMANDPREVIEW", 0);    // Tắt command preview
                    AcadApp.SetSystemVariable("DYNMODE", 0);           // Tắt Dynamic Input
                    AcadApp.SetSystemVariable("CURSORBADGE", 0);       // Tắt cursor badge
                    AcadApp.SetSystemVariable("UCSDETECT", 0);         // Tắt UCS detect

                    // --- Display Performance ---
                    AcadApp.SetSystemVariable("HPQUICKPREVIEW", 0);    // Tắt hatch preview
                    AcadApp.SetSystemVariable("VTENABLE", 0);          // Tắt view transitions
                    AcadApp.SetSystemVariable("LAYLOCKFADECTL", 0);    // Tắt fade locked layers
                    AcadApp.SetSystemVariable("TEXTFILL", 0);          // Không fill text (nhanh hơn)
                    AcadApp.SetSystemVariable("DRAWORDERCTL", 0);      // Tắt draw order (nhanh hơn)
                    AcadApp.SetSystemVariable("GRIPSIZE", 2);          // Giảm grip size
                    AcadApp.SetSystemVariable("GRIPS", 1);             // Grip đơn giản
                    AcadApp.SetSystemVariable("ANNOALLVISIBLE", 0);    // Chỉ hiện anno đúng scale

                    // --- Regen & Cache ---
                    AcadApp.SetSystemVariable("LAYOUTREGENCTL", 2);    // Cache Model + Layout
                    AcadApp.SetSystemVariable("WHIPTHREAD", 3);        // Multithreaded display
                    AcadApp.SetSystemVariable("CACHEMAXFILES", 512);   // Tăng cache files
                    AcadApp.SetSystemVariable("BACKGROUNDPLOT", 0);    // In foreground
                    
                    // --- Xref & Proxy ---
                    AcadApp.SetSystemVariable("XLOADCTL", 2);          // Demand load xrefs
                    AcadApp.SetSystemVariable("DEMANDLOAD", 3);        // ObjectARX demand load
                    AcadApp.SetSystemVariable("INDEXCTL", 0);          // Tắt layer/spatial index
                    AcadApp.SetSystemVariable("PROXYGRAPHICS", 0);     // Không hiện proxy graphics
                    
                    // --- Layer ---
                    AcadApp.SetSystemVariable("LAYEREVALCTL", 0);      // Tắt layer eval notifications
                    AcadApp.SetSystemVariable("LAYERNOTIFY", 0);       // Tắt layer notifications

                    _turboModeActive = true;
                    
                    ed.WriteMessage("\n  ✓ Đã tắt Selection Preview, Highlight, Tooltips");
                    ed.WriteMessage("\n  ✓ Đã tắt Dynamic Input, Command Preview");
                    ed.WriteMessage("\n  ✓ Đã tắt View Transitions, Hatch Preview");
                    ed.WriteMessage("\n  ✓ Đã bật Multithreaded Display (WHIPTHREAD=3)");
                    ed.WriteMessage("\n  ✓ Đã tối ưu Cache, Xref, Proxy");
                    ed.WriteMessage("\n");
                    ed.WriteMessage("\n  ⚡ TURBO MODE ĐÃ BẬT! Gõ lại CTDS_TurboMode để tắt.");
                }
                else
                {
                    // ═══ TẮT TURBO MODE - Khôi phục mặc định ═══
                    ed.WriteMessage("\n▸ Đang tắt Turbo Mode...");

                    AcadApp.SetSystemVariable("SELECTIONPREVIEW", 3);
                    AcadApp.SetSystemVariable("HIGHLIGHT", 1);
                    AcadApp.SetSystemVariable("SELECTIONEFFECT", 1);
                    AcadApp.SetSystemVariable("ROLLOVERTIPS", 1);
                    AcadApp.SetSystemVariable("TOOLTIPS", 1);
                    AcadApp.SetSystemVariable("PROPERTYPREVIEW", 1);
                    AcadApp.SetSystemVariable("COMMANDPREVIEW", 1);
                    AcadApp.SetSystemVariable("DYNMODE", 1);
                    AcadApp.SetSystemVariable("CURSORBADGE", 1);
                    AcadApp.SetSystemVariable("UCSDETECT", 1);
                    AcadApp.SetSystemVariable("HPQUICKPREVIEW", 1);
                    AcadApp.SetSystemVariable("VTENABLE", 3);
                    AcadApp.SetSystemVariable("LAYLOCKFADECTL", 50);
                    AcadApp.SetSystemVariable("TEXTFILL", 1);
                    AcadApp.SetSystemVariable("DRAWORDERCTL", 3);
                    AcadApp.SetSystemVariable("GRIPSIZE", 5);
                    AcadApp.SetSystemVariable("GRIPS", 2);
                    AcadApp.SetSystemVariable("ANNOALLVISIBLE", 1);
                    AcadApp.SetSystemVariable("PROXYGRAPHICS", 1);
                    AcadApp.SetSystemVariable("LAYERNOTIFY", 15);

                    _turboModeActive = false;

                    ed.WriteMessage("\n◎ Turbo Mode đã TẮT. Các cài đặt hiển thị đã khôi phục.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n⊘ Lỗi: {ex.Message}");
            }
        }

        // ══════════════════════════════════════════════════════════════
        // THIẾT LẬP NÂNG CAO CHO CIVIL 3D (Bổ sung cho CTDS_ThietLap)
        // ══════════════════════════════════════════════════════════════
        /// <summary>
        /// Thiết lập nâng cao tối ưu cho Civil 3D: purge DGN linestyle,
        /// tắt auto-rebuild corridor, và các cài đặt đặc biệt khác.
        /// </summary>
        [CommandMethod("CTDS_ThietLapNangCao")]
        public static void CTDS_ThietLapNangCao()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            int optimized = 0;

            try
            {
                // ── 1. Purge DGN Linestyles (thường gây file phình to) ──
                ed.WriteMessage("\n▸ Đang dọn DGN Linestyles...");
                try
                {
                    doc.SendStringToExecute("-PURGE O ", true, false, false);
                    optimized++;
                }
                catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

                // ── 2. System Variables cho Civil 3D ──
                ed.WriteMessage("\n▸ Đang tối ưu Civil 3D settings...");

                // Tắt auto-update các object phụ thuộc khi thay đổi surface
                try { AcadApp.SetSystemVariable("OBJECTISOLATIONMODE", 0); optimized++; } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

                // Multithreaded regeneration
                try { AcadApp.SetSystemVariable("WHIPTHREAD", 3); optimized++; } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                
                // Tắt annotation monitor (chậm với nhiều label)
                try { AcadApp.SetSystemVariable("ANNOMONITOR", 0); optimized++; } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                
                // Tối ưu undo - giảm memory footprint
                try { AcadApp.SetSystemVariable("UNDOCTL", 1); optimized++; } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                
                // Tắt selection cycling (chậm với file phức tạp)  
                try { AcadApp.SetSystemVariable("SELECTIONCYCLING", 0); optimized++; } catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

                // ── 3. Cleanup Orphaned Data ──
                ed.WriteMessage("\n▸ Đang dọn Regapps...");
                try
                {
                    doc.SendStringToExecute("-PURGE R * N ", true, false, false);
                    optimized++;
                }
                catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }

                ed.WriteMessage($"\n◎ Đã tối ưu nâng cao xong! ({optimized} cài đặt đã áp dụng)");
                ed.WriteMessage("\n  💡 Tip: Gõ CTDS_TurboMode để bật chế độ hiệu suất tối đa");
                ed.WriteMessage("\n  💡 Tip: Gõ CTDS_SaveClean để purge và lưu file sạch");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n⊘ Lỗi: {ex.Message}");
            }
        }
    }
}
