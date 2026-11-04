---
trigger: glob
globs: MyFirstProject/**/*.cs
---

# Mẫu code AutoCAD / Civil 3D trong dự án

Lệnh chuẩn:

```csharp
[CommandMethod("CTSV_TenLenh")]
public static void CTSVTenLenh()
{
    SmartCommand.Execute("CTSV_TenLenh", pm =>
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        using var tr = doc.Database.TransactionManager.StartTransaction();
        // ... đọc bằng OpenMode.ForRead, chỉ ForWrite khi cần sửa
        tr.Commit();
    });
}
```

- Mỗi file lệnh có `[assembly: CommandClass(typeof(...))]` để AutoCAD nạp nhanh.
- Lấy đối tượng Civil: `CivilApplication.ActiveDocument`, `tr.GetObject(id, OpenMode.ForRead) as Alignment`.
- Chọn đối tượng: dùng `PromptEntityOptions.SetRejectMessage` + `AddAllowedClass` để lọc đúng kiểu.
- Người dùng bấm Esc (`PromptStatus.Cancel`) → `return` im lặng, không báo lỗi.
- Thông báo dòng lệnh bằng tiếng Việt, có tiền tố `[TEN_LENH]`.
- Xuất Excel: ClosedXML (`ExcelDependencyLoader.EnsureLoaded()` trước). Không dùng Excel Interop cho tính năng mới.
- Không chặn luồng UI của AutoCAD quá lâu: vòng lặp hàng nghìn đối tượng → dùng `ProgressMeter` (SmartCommand cung cấp `pm`).
