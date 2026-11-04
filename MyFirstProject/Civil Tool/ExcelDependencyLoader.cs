// ExcelDependencyLoader.cs
// Giải quyết lỗi "Could not load ExcelNumberFormat" trong AutoCAD
// AutoCAD .NET 8 không tự tìm DLL phụ thuộc của ClosedXML
//
using System;
using System.IO;
using System.Reflection;

namespace MyFirstProject
{
    /// <summary>
    /// Đảm bảo tất cả DLL phụ thuộc của ClosedXML được load trước khi sử dụng.
    /// Gọi EnsureLoaded() TRƯỚC bất kỳ dòng code nào dùng ClosedXML.
    /// </summary>
    public static class ExcelDependencyLoader
    {
        private static bool _loaded = false;
        private static readonly object _lock = new();

        /// <summary>
        /// Load tất cả DLL phụ thuộc ClosedXML. Gọi trước khi dùng XLWorkbook.
        /// </summary>
        public static void EnsureLoaded()
        {
            if (_loaded) return;
            lock (_lock)
            {
                if (_loaded) return;

                string pluginDir = Path.GetDirectoryName(typeof(ExcelDependencyLoader).Assembly.Location) ?? "";
                if (string.IsNullOrEmpty(pluginDir)) return;

                // Build lookup set once — O(n) instead of O(n*m) nested foreach
                var loadedNames = new System.Collections.Generic.HashSet<string>();
                foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
                {
                    var name = asm.GetName().Name;
                    if (name != null) loadedNames.Add(name);
                }

                // Danh sách DLL mà ClosedXML cần
                string[] deps = new[]
                {
                    "ExcelNumberFormat",
                    "DocumentFormat.OpenXml",
                    "DocumentFormat.OpenXml.Framework",
                    "ClosedXML.Parser",
                    "SixLabors.Fonts",
                    "RBush",
                    "System.IO.Packaging"
                };

                foreach (string name in deps)
                {
                    if (loadedNames.Contains(name)) continue;

                    string dllPath = Path.Combine(pluginDir, name + ".dll");
                    if (File.Exists(dllPath))
                    {
                        try
                        {
                            // Đọc bytes rồi load — cách này bypass mọi context/probing issue
                            byte[] bytes = File.ReadAllBytes(dllPath);
                            Assembly.Load(bytes);
                        }
                        catch { }
                    }
                }

                _loaded = true;
            }
        }
    }
}
