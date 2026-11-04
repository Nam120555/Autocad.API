// ==============================================================================
// NxUsage.cs — Thống kê số lần dùng từng lệnh của Civil Tool (chỉ lưu trên máy)
//
// Lưu tại %AppData%\CivilTool\usage.json — không gửi đi đâu.
// Mục đích: biết lệnh nào dùng nhiều nhất để ưu tiên hoàn thiện / đưa lên Ribbon.
// Lệnh CT_THONGKELENH: in bảng xếp hạng lệnh hay dùng.
// ==============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.Core.NxUsage))]

namespace MyFirstProject.Core
{
    public class NxUsage
    {
        public class Entry
        {
            public int Count { get; set; }
            public DateTime LastUsed { get; set; }
        }

        private static readonly object _lock = new();
        private static HashSet<string> _ourCommands = new(StringComparer.OrdinalIgnoreCase);
        private static Dictionary<string, Entry> _data = new(StringComparer.OrdinalIgnoreCase);
        private static bool _started;

        private static string FilePath => Path.Combine(AppInfo.DataFolder, "usage.json");

        /// <summary>Gọi 1 lần khi plugin khởi động (AutoLoader.Initialize).</summary>
        public static void Start()
        {
            if (_started) return;
            _started = true;
            try
            {
                _ourCommands = CollectOwnCommands();
                Load();
                var dm = AcadApp.DocumentManager;
                foreach (Document d in dm) Attach(d);
                dm.DocumentCreated += (s, e) => { if (e.Document != null) Attach(e.Document); };
                NxLog.Info($"Thống kê lệnh: theo dõi {_ourCommands.Count} lệnh của Civil Tool");
            }
            catch (System.Exception ex) { NxLog.Error(ex, "Khởi động thống kê lệnh"); }
        }

        private static void Attach(Document doc)
        {
            try { doc.CommandWillStart += OnCommandWillStart; }
            catch (System.Exception ex) { NxLog.Swallow(ex); }
        }

        private static void OnCommandWillStart(object? sender, CommandEventArgs e)
        {
            try
            {
                string name = e.GlobalCommandName ?? "";
                if (name.Length == 0 || !_ourCommands.Contains(name)) return;
                lock (_lock)
                {
                    if (!_data.TryGetValue(name, out var entry)) _data[name] = entry = new Entry();
                    entry.Count++;
                    entry.LastUsed = DateTime.Now;
                    Save();
                }
            }
            catch (System.Exception ex) { NxLog.Swallow(ex); }
        }

        /// <summary>Lấy tên toàn bộ [CommandMethod] trong DLL này.</summary>
        private static HashSet<string> CollectOwnCommands()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Type?[] types;
            try { types = typeof(NxUsage).Assembly.GetTypes(); }
            catch (ReflectionTypeLoadException rtle) { types = rtle.Types; }
            foreach (var t in types)
            {
                if (t == null) continue;
                MethodInfo[] methods;
                try { methods = t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance); }
                catch (System.Exception ex) { NxLog.Swallow(ex); continue; }
                foreach (var m in methods)
                {
                    try
                    {
                        foreach (var a in m.GetCustomAttributes<CommandMethodAttribute>(false))
                            if (!string.IsNullOrEmpty(a.GlobalName)) set.Add(a.GlobalName);
                    }
                    catch (System.Exception ex) { NxLog.Swallow(ex); }
                }
            }
            return set;
        }

        private static void Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return;
                var d = JsonSerializer.Deserialize<Dictionary<string, Entry>>(File.ReadAllText(FilePath));
                if (d != null) _data = new Dictionary<string, Entry>(d, StringComparer.OrdinalIgnoreCase);
            }
            catch (System.Exception ex) { NxLog.Error(ex, "Đọc usage.json"); }
        }

        private static void Save()
        {
            try
            {
                File.WriteAllText(FilePath, JsonSerializer.Serialize(_data, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch (System.Exception ex) { NxLog.Swallow(ex); }
        }

        /// <summary>Số lần đã dùng 1 lệnh (0 nếu chưa dùng).</summary>
        public static int CountOf(string command)
        {
            lock (_lock) return _data.TryGetValue(command ?? "", out var e) ? e.Count : 0;
        }

        /// <summary>Số lệnh khác nhau đã từng dùng.</summary>
        public static int UsedCount { get { lock (_lock) return _data.Count; } }

        /// <summary>Danh sách lệnh hay dùng (giảm dần) — dùng cho Hộp lệnh / Ribbon sau này.</summary>
        public static List<(string Command, int Count, DateTime LastUsed)> Top(int n = 20)
        {
            lock (_lock)
                return _data.OrderByDescending(kv => kv.Value.Count)
                            .Take(n)
                            .Select(kv => (kv.Key, kv.Value.Count, kv.Value.LastUsed))
                            .ToList();
        }

        [CommandMethod("CT_THONGKELENH")]
        public static void ShowStats()
        {
            var ed = AcadApp.DocumentManager.MdiActiveDocument?.Editor;
            if (ed == null) return;
            var top = Top(30);
            ed.WriteMessage($"\n═══ {AppInfo.ProductName} v{AppInfo.Version} — LỆNH HAY DÙNG ═══");
            if (top.Count == 0)
            {
                ed.WriteMessage("\n  Chưa có dữ liệu. Thống kê bắt đầu ghi từ phiên bản này.\n");
                return;
            }
            int i = 0;
            foreach (var (cmd, count, last) in top)
            {
                string desc = NxCatalogLookup(cmd);
                ed.WriteMessage($"\n  {++i,2}. {cmd,-32} {count,5} lần   (gần nhất {last:dd/MM/yyyy})  {desc}");
            }
            ed.WriteMessage($"\n  File dữ liệu: {FilePath}\n");
        }

        private static string NxCatalogLookup(string cmd)
        {
            try
            {
                return MyFirstProject.UI.NxCommandCatalog.Entries.TryGetValue(cmd.ToUpperInvariant(), out var v) ? v.Description : "";
            }
            catch (System.Exception ex) { NxLog.Swallow(ex); return ""; }
        }
    }
}
