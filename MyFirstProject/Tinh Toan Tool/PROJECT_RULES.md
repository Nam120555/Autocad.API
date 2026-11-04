# Civil Tool Suite - Project Rules

Đây là tài liệu lưu trữ các nguyên tắc xây dựng giao diện của dự án MyFirstProject.

## Nguyên tắc 1: Chữ ký tác giả trên giao diện (Author Signature)

Bất kỳ Form giao diện (UI) tính toán nào cũng **BẮT BUỘC** phải có khối hiển thị thông tin tác giả như sau:

- **Thiết kế bởi**: HUYNH NGOC NAM
- **Nghề nghiệp**: TRAFFIC ENGINEER
- **Số điện thoại**: 0378584154

**Cách lập trình chuẩn (C# WinForms):**
Nên thêm khối `Label` vào góc dưới của form hoặc ở thanh SideBar bên phải/trái:

```csharp
Label lblSignatureRole = new Label
{
    Text = "DESIGNED BY HUYNH NGOC NAM\nTRAFFIC ENGINEER - SĐT: 0378584154",
    ForeColor = Color.FromArgb(231, 76, 60), // accentRed
    Font = new Font("Segoe UI", 9, FontStyle.Italic | FontStyle.Bold),
    Location = new Point(15, ry + 120), // Tùy chỉnh tham số Y cho phù hợp với panel
    AutoSize = true
};
panel.Controls.Add(lblSignatureRole);
```

Nguyên tắc này áp dụng cho TOÀN BỘ các file của dự án MyFirstProject, đảm bảo giao diện nào khi bật lên cũng hiển thị đầy đủ *tên, nghề và SĐT* của tác giả gốc.
