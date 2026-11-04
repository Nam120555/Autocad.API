---
description: Tạo một lệnh mới cho Civil Tool theo đúng quy ước dự án
---

1. Hỏi người dùng (nếu chưa rõ): tên lệnh, nhóm (tiền tố theo AGENTS.md), đầu vào cần chọn, kết quả mong muốn.
2. Tạo file trong thư mục đúng nhóm (`Civil Tool/`, `Acad Tool/`...), đặt tên `NN.TenChucNang.cs`.
3. Viết lệnh theo mẫu trong `.agents/rules/autocad-csharp.md` (SmartCommand, transaction, tiếng Việt).
4. Nếu có form: kế thừa `CivilFormBase` hoặc gọi `FormSignatureHelper.AddSignature(form)`.
5. Thêm dòng mô tả vào `MyFirstProject/UI Pro/NxCommandCatalog.g.cs` và bảng lệnh trong `README.md`.
6. Chạy workflow `/build`.
7. Hướng dẫn người dùng cách thử lệnh trong Civil 3D (gõ `NRL` để nạp lại nếu đang dùng NetReload).
