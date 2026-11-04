// ==============================================================================
// AppInfo.cs — Thông tin ứng dụng dùng chung (tên, phiên bản, tác giả, thư mục)
// Muốn đổi SĐT / tên hiển thị: CHỈ sửa ở file này.
// Phiên bản lấy tự động từ <Version> trong MyFirstProject.csproj.
// ==============================================================================
using System;
using System.IO;
using System.Reflection;

namespace MyFirstProject.Core
{
    public static class AppInfo
    {
        public const string ProductName = "CIVIL TOOL";
        public const string AuthorName  = "HUYNH NGOC NAM";
        public const string AuthorRole  = "TRAFFIC ENGINEER";
        public const string AuthorPhone = "0378584154";
        public const string AuthorEmail = "Namhuynh1205@gmail.com";

        /// <summary>Dòng chữ ký chuẩn hiển thị trên các form.</summary>
        public static string Signature =>
            $"●  DESIGNED BY {AuthorName} / {AuthorRole}   |   SĐT: {AuthorPhone}";

        /// <summary>Phiên bản hiển thị, ví dụ "1.2.0".</summary>
        public static string Version
        {
            get
            {
                var asm = typeof(AppInfo).Assembly;
                var info = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
                if (!string.IsNullOrEmpty(info))
                {
                    int plus = info.IndexOf('+');          // bỏ phần "+commitHash" nếu có
                    return plus > 0 ? info.Substring(0, plus) : info;
                }
                return asm.GetName().Version?.ToString(3) ?? "0.0.0";
            }
        }

        /// <summary>%AppData%\CivilTool — nơi lưu cấu hình, alias, log.</summary>
        public static string DataFolder
        {
            get
            {
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CivilTool");
                try { Directory.CreateDirectory(dir); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine(ex); }
                return dir;
            }
        }

        /// <summary>%AppData%\CivilTool\logs</summary>
        public static string LogFolder
        {
            get
            {
                string dir = Path.Combine(DataFolder, "logs");
                try { Directory.CreateDirectory(dir); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine(ex); }
                return dir;
            }
        }
    }
}
