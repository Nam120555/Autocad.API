---
description: Build kiểm tra plugin Civil Tool và báo số lỗi / cảnh báo
---

// turbo
1. Chạy `dotnet build MyFirstProject/MyFirstProject.csproj -c Release -nologo -v minimal` tại thư mục gốc repo.
2. Nếu có lỗi: liệt kê từng lỗi (file:dòng + giải thích tiếng Việt) và sửa. Build lại đến khi 0 lỗi.
3. So số cảnh báo với mốc ~55. Nếu tăng nhiều, chỉ ra cảnh báo mới do thay đổi vừa làm.
4. Báo kết quả ngắn gọn cho người dùng, kèm danh sách lệnh cần thử lại trong Civil 3D.
