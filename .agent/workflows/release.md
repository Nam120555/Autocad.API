---
description: Phát hành phiên bản mới của Civil Tool (tăng version, changelog, bộ cài, tag)
---

1. Hỏi người dùng số phiên bản mới (gợi ý: sửa lỗi → tăng số cuối, tính năng mới → tăng số giữa).
2. Sửa `<Version>` trong `MyFirstProject/MyFirstProject.csproj`.
3. Thêm mục mới đầu `CHANGELOG.md` (Thêm / Thay đổi / Sửa / Xoá) dựa trên `git log` từ tag trước.
4. Chạy `build.bat` → kiểm tra file `Civil_Tool_NXSOFT_Setup_vX.Y.Z.exe` được tạo trong `D:\Autocad 2026_API\Civil_Tool_Installer`.
5. `git add -A`, `git commit -m "release: vX.Y.Z"`, `git tag vX.Y.Z`.
6. Hỏi người dùng trước khi `git push origin master --tags`.
