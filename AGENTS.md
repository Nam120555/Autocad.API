# Civil Tool – NXSOFT · Hướng dẫn cho AI agent

Plugin C# cho **AutoCAD Civil 3D** (tác giả Huỳnh Ngọc Nam). Trả lời và viết giao diện bằng **tiếng Việt**.
Người dùng là kỹ sư giao thông, không phải lập trình viên chuyên nghiệp → giải thích ngắn gọn, làm từng bước có kiểm tra.

## Cấu trúc
- `MyFirstProject/` – plugin chính (DLL `MyFirstProject.dll`, ~230 lệnh). **Không đổi `AssemblyName`/tên bundle** (máy đã cài sẽ bị nạp trùng 2 plugin).
  - `Core/` – `AppInfo` (tên, SĐT, phiên bản, thư mục dữ liệu), `NxLog` (ghi log), `NxUsage` (thống kê lệnh).
  - `Civil Tool/` – lệnh Civil 3D. `Acad Tool/` – lệnh AutoCAD thuần. `Br Tool/` – tiện ích BR.
  - `Extensions/` – tiện ích dùng chung, `07.SmartCommand.cs` (chạy lệnh có tiến trình + báo lỗi).
  - `UI Pro/` – Hộp lệnh NX, theme, `NxCommandCatalog.g.cs` (mô tả lệnh). `Menu form/` – Ribbon, cài đặt menu.
  - `Tinh Toan Tool/TinhToan/` – module tính toán thiết kế đường (TT_*).
  - `NetReload/` – project riêng để hot-reload (lệnh `NRL`). Không sửa nếu không được yêu cầu.
- `Installer/` – Inno Setup (`setup_civiltool.iss`) + `PackageContents.xml`.
- `OutLookTool/`, `PavementDesignApp/` – ứng dụng độc lập, tách khỏi plugin.

## Build & kiểm tra (Windows, .NET SDK 8)
- Build nhanh: `dotnet build MyFirstProject/MyFirstProject.csproj -c Debug`
- Build phát hành + bộ cài: `build.bat` (Release; phiên bản lấy từ `<Version>` trong csproj).
- Kiểm tra + commit: ghi thông điệp vào `_ci_message.txt` rồi chạy `ci.bat` → đọc `_ci_ketqua.txt`.
- Nếu AutoCAD đang mở, DLL bản Debug có thể bị khoá → build `-c Release` để kiểm tra biên dịch.
- Mốc hiện tại: **0 lỗi, ~55 cảnh báo**. Thay đổi nào làm tăng cảnh báo đáng kể phải giải thích.
- Không chạy được test tự động trong AutoCAD → sau khi sửa lệnh, ghi rõ cho người dùng **lệnh nào cần thử lại bằng tay**.

## Quy tắc bắt buộc
1. **Target `net8.0-windows`, API qua NuGet `Chuongmep.Civil3D.Api.* 2025.0.0`.** Không thêm `HintPath` tới `C:\Program Files\Autodesk\AutoCAD 2027\` (DLL .NET 10 → hàng nghìn cảnh báo xung đột).
2. **Không viết `catch { }` rỗng.** Dùng `catch (System.Exception ex) { NxLog.Swallow(ex); }` (lỗi bỏ qua có chủ đích) hoặc `NxLog.Error(ex, "ngữ cảnh")`.
3. **Không gõ cứng tên/SĐT/phiên bản** – dùng `MyFirstProject.Core.AppInfo`.
4. **Mọi form tính toán phải có chữ ký tác giả**: gọi `FormSignatureHelper.AddSignature(form)` hoặc kế thừa `CivilFormBase`.
5. **Lệnh mới**: chạy logic qua `SmartCommand.Execute("TEN_LENH", pm => { ... })`; thêm 1 dòng vào `NxCommandCatalog.g.cs`; cập nhật `README.md` nếu là lệnh người dùng.
6. **Tiền tố tên lệnh** theo nhóm: `AT_` AutoCAD · `CTA_` tuyến · `CTP_` trắc dọc · `CTSV_` trắc ngang · `CTS_` cọc/sample line · `CTC_` corridor · `CTPO_` CogoPoint · `CTPI_` cống · `CTL_` layer · `CTU_` tiện ích · `CTDS_` thiết lập bản vẽ · `TT_` tính toán · `CT_` hệ thống.
7. **Không đổi tên lệnh đã có** (người dùng có alias, ribbon, thói quen). Muốn đổi → thêm lệnh mới, giữ lệnh cũ.
8. **Không tạo lệnh "vỏ rỗng"** chỉ in thông báo. Chưa làm xong thì không đăng ký `[CommandMethod]`.
9. **Không sao lưu bằng file `.bak`** – dùng git. Không commit `bin/`, `obj/`, `.vs/`, log.
10. **Tính toán khối lượng/diện tích là số liệu thật cho hồ sơ thiết kế** → không đổi công thức, cách làm tròn (`AreaDecimalPlaces`, `VolumeDecimalPlaces`) khi chưa có bản vẽ mẫu + kết quả đúng để so sánh.
11. Phát hành: sửa `<Version>` → ghi `CHANGELOG.md` → `build.bat` → `git tag vX.Y.Z`.

## Lỗi hay gặp
- `Exception` bị nhập nhằng giữa `System` và `Autodesk.AutoCAD.Runtime` → luôn viết `System.Exception`.
- ClosedXML trong AutoCAD: gọi `ExcelDependencyLoader.EnsureLoaded()` trước khi dùng `XLWorkbook`.
- Form không modal sửa bản vẽ phải `using (doc.LockDocument())`.
- File lớn (`08.Sectionview.cs`, `07.Sampleline.cs`, `39.TinhKhoiLuongTracNgang.cs`...) – khi sửa, ưu tiên tách `partial class` sang file mới thay vì nhồi thêm.
