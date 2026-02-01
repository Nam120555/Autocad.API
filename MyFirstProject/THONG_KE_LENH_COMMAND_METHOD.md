# TH?NG KÊ L?NH COMMAND METHOD TRONG D? ÁN

## ?? T?ng Quan
- **T?ng s? file .cs ?ã ??c:** 12 file
- **T?ng s? l?nh CommandMethod:** 98 l?nh
- **Ngày c?p nh?t báo cáo:** 2024-12-19

---

## ?? **1. Menu form\18.Menu Risbbon.cs**
**S? l?nh:** 3
- `ShowForm` - Hi?n th? form test
- `AdskGreeting` - T?o text chào m?ng
- `show_menu` - Hi?n th? menu ribbon

---

## ?? **2. Civil Tool\01.Corridor.cs**
**S? l?nh:** 2
- `CTC_AddAllSection` - Thêm t?t c? section vào corridor
- `CTC_TaoCooridor_DuongDoThi_RePhai` - T?o corridor ???ng ?ô th? r? ph?i

---

## ?? **3. Civil Tool\02.Parcel.cs**
**S? l?nh:** 1
- `CTPA_TaoParcel_CacLoaiNha` - T?o parcel các lo?i nhà

---

## ?? **4. Civil Tool\04.PipeAndStructures.cs**
**S? l?nh:** 5
- `CTPI_ThayDoi_DuongKinhCong` - Thay ??i ???ng kính c?ng
- `CTPI_ThayDoi_MatPhangRef_Cong` - Thay ??i m?t ph?ng reference c?ng
- `CTPI_ThayDoi_DoanDocCong` - Thay ??i ?? d?c c?ng
- `CTPI_BangCaoDo_TuNhienHoThu` - B?ng cao ?? t? nhiên h? thu
- `CTPI_XoayHoThu_Theo2diem` - Xoay h? thu theo 2 ?i?m

---

## ?? **5. Civil Tool\05.Point.cs**
**S? l?nh:** 5
- `CTPO_TaoCogoPoint_CaoDo_FromSurface` - T?o CogoPoint cao ?? t? Surface
- `CTPO_TaoCogoPoint_CaoDo_Elevationspot` - T?o CogoPoint t? Elevation spot
- `CTPO_UpdateAllPointGroup` - Update t?t c? Point Group
- `CTPO_CreateCogopointFromText` - T?o CogoPoint t? Text
- `CTPO_An_CogoPoint` - ?n CogoPoint

---

## ?? **6. Civil Tool\06.ProfileAndProfileView.cs**
**S? l?nh:** 5
- `CTP_VeTracDoc_TuNhien` - V? tr?c d?c t? nhiên
- `CTP_VeTracDoc_TuNhien_TatCaTuyen` - V? tr?c d?c t? nhiên t?t c? tuy?n
- `CTP_Fix_DuongTuNhien_TheoCoc` - S?a ???ng t? nhiên theo c?c
- `CTP_GanNhanNutGiao_LenTracDoc` - G?n nhãn nút giao lên tr?c d?c
- `CTP_TaoCogoPointTuPVI` - T?o CogoPoint t? PVI

---

## ?? **7. Civil Tool\07.Sampleline.cs**
**S? l?nh:** 26
- `CTS_DoiTenCoc` - ??i tên c?c
- `CTS_DoiTenCoc3` - ??i tên c?c theo Km
- `CTS_DoiTenCoc2` - ??i tên c?c theo ?o?n
- `CTS_TaoBang_ToaDoCoc` - T?o b?ng t?a ?? c?c
- `CTS_TaoBang_ToaDoCoc2` - T?o b?ng t?a ?? c?c (có lý trình)
- `CTS_TaoBang_ToaDoCoc3` - T?o b?ng t?a ?? c?c (có cao ??)
- `AT_UPdate2Table` - C?p nh?t 2 table
- `CTS_ChenCoc_TrenTracDoc` - Chèn c?c trên tr?c d?c
- `CTS_CHENCOC_TRENTRACNGANG` - Chèn c?c trên tr?c ngang
- `CTS_DoiTenCoc_fromCogoPoint` - ??i tên c?c t? CogoPoint
- `CTS_PhatSinhCoc` - Phát sinh c?c
- `CTS_PhatSinhCoc_ChiTiet` - Phát sinh c?c chi ti?t
- `CTS_PhatSinhCoc_theoKhoangDelta` - Phát sinh c?c theo kho?ng delta
- `CTS_PhatSinhCoc_TuCogoPoint` - Phát sinh c?c t? CogoPoint
- `CTS_DoiTenCoc_TheoThuTu` - ??i tên c?c theo th? t?
- `CTS_DichCoc_TinhTien` - D?ch c?c t?nh ti?n
- `CTS_Copy_NhomCoc` - Copy nhóm c?c
- `CTS_DongBo_2_NhomCoc` - ??ng b? 2 nhóm c?c
- `CTS_DongBo_2_NhomCoc_TheoDoan` - ??ng b? 2 nhóm c?c theo ?o?n
- `CTS_DichCoc_TinhTien40` - D?ch c?c 40m
- `CTS_DichCoc_TinhTien_20` - D?ch c?c 20m
- `CTS_DoiTenCoc_H` - ??i tên c?c H
- `CTS_PhatSinhCoc_TheoBang` - Phát sinh c?c theo b?ng
- `CTS_Copy_BeRong_sampleLine` - Copy b? r?ng sample line
- `CTS_Thaydoi_BeRong_sampleLine` - Thay ??i b? r?ng sample line
- `CTS_Offset_BeRong_sampleLine` - Offset b? r?ng sample line

---

## ?? **8. Civil Tool\08.Sectionview.cs**
**S? l?nh:** 20
- `CTSV_VeTracNgangThietKe` - V? tr?c ngang thi?t k?
- `CVSV_VeTatCa_TracNgangThietKe` - V? t?t c? tr?c ngang thi?t k?
- `CTSV_ChuyenDoi_TNTK_TNTN` - Chuy?n ??i TN-TK sang TN-TN
- `CTSV_DanhCap` - Tính ?ánh c?p
- `CTSV_DanhCap_XoaBo` - Xóa b? ?ánh c?p
- `CTSV_DanhCap_VeThem` - V? thêm ?ánh c?p
- `CTSV_DanhCap_VeThem2` - V? thêm ?ánh c?p 2m
- `CTSV_DanhCap_VeThem1` - V? thêm ?ánh c?p 1m
- `CTSV_DanhCap_CapNhat` - C?p nh?t ?ánh c?p
- `CTSV_ThemVatLieu_TrenCatNgang` - Thêm v?t li?u trên c?t ngang
- `CTSV_ThayDoi_MSS_Min_Max` - Thay ??i MSS Min Max
- `CTSV_ThayDoi_GioiHan_traiPhai` - Thay ??i gi?i h?n trái ph?i
- `CTSV_ThayDoi_KhungIn` - Thay ??i khung in
- `CTSV_KhoaCatNgang_AddPoint` - Khóa c?t ngang add point
- `CTSV_fit_KhungIn` - Fit khung in
- `CTSV_fit_KhungIn_5_5_top` - Fit khung in 5x5
- `CTSV_fit_KhungIn_5_10_top` - Fit khung in 5x10
- `CTSV_An_DuongDiaChat` - ?n ???ng ??a ch?t
- `CTSV_HieuChinh_Section` - Hi?u ch?nh section static
- `CTSV_HieuChinh_Section_Dynamic` - Hi?u ch?nh section dynamic

---

## ?? **9. Civil Tool\09.Surfaces.cs**
**S? l?nh:** 1
- `CTS_TaoSpotElevation_OnSurface_TaiTim` - T?o spot elevation trên surface t?i tim

---

## ?? **10. Civil Tool\Property Sets.cs** ? **M?I**
**S? l?nh:** 2
- `AT_Solid_Set_PropertySet` - Thi?t l?p Property Set cho 3D Solid
- `AT_Solid_Show_Info` - Hi?n th? thông tin 3D Solid

---

## ?? **11. Acad Tool\01.CAD.cs**
**S? l?nh:** 23
- `AT_TongDoDai_Full` - T?ng ?? dài (Full)
- `AT_TongDoDai_Replace` - T?ng ?? dài (Replace)
- `AT_TongDoDai_Replace2` - T?ng ?? dài (Replace2)
- `AT_TongDoDai_Replace_CongThem` - T?ng ?? dài (C?ng thêm)
- `AT_TongDienTich_Full` - T?ng di?n tích (Full)
- `AT_TongDienTich_Replace` - T?ng di?n tích (Replace)
- `AT_TongDienTich_Replace2` - T?ng di?n tích (Replace2)
- `AT_TongDienTich_Replace_CongThem` - T?ng di?n tích (C?ng thêm)
- `AT_TextLink` - Text Link
- `AT_DanhSoThuTu` - ?ánh s? th? t?
- `AT_XoayDoiTuong_TheoViewport` - Xoay ??i t??ng theo viewport
- `AT_XoayDoiTuong_Theo2Diem` - Xoay ??i t??ng theo 2 ?i?m
- `AT_TextLayout` - Text Layout
- `AT_TaoMoi_TextLayout` - T?o m?i Text Layout
- `AT_DimLayout2` - Dim Layout 2
- `AT_DimLayout` - Dim Layout
- `AT_BlockLayout` - Block Layout
- `AT_Label_FromText` - Label From Text
- `AT_XoaDoiTuong_CungLayer` - Xóa ??i t??ng cùng layer
- `AT_XoaDoiTuong_3DSolid_Body` - Xóa 3DSolid/Body
- `AT_UpdateLayout` - Update Layout
- `AT_Offset_2Ben` - Offset 2 bên
- `AT_annotive_scale_currentOnly` - Annotative scale current only

---

## ?? **12. Acad Tool\Command_XUATBANG_ToaDoPolyline.cs**
**S? l?nh:** 1
- `XUATBANG_ToaDoPolyline` - Xu?t b?ng t?a ?? polyline

---

## ?? **TH?NG KÊ THEO DANH M?C**

### **Civil Tool (Civil 3D Commands)**
- **Corridor:** 2 l?nh
- **Parcel:** 1 l?nh  
- **Pipe & Structures:** 5 l?nh
- **Point:** 5 l?nh
- **Profile & ProfileView:** 5 l?nh
- **Sampleline:** 26 l?nh
- **Sectionview:** 20 l?nh
- **Surfaces:** 1 l?nh
- **Property Sets:** 2 l?nh ? **M?I**
- **T?ng Civil Tool:** 67 l?nh

### **Acad Tool (AutoCAD Commands)**
- **CAD General:** 23 l?nh
- **Export Tools:** 1 l?nh
- **T?ng Acad Tool:** 24 l?nh

### **Menu & UI**
- **Menu Ribbon:** 3 l?nh
- **T?ng Menu:** 3 l?nh

### **Extensions & Utilities**
- **PropertySetUtils:** Class h? tr? (không có CommandMethod tr?c ti?p)

---

## ?? **TOP 5 FILE CÓ NHI?U L?NH NH?T**

1. **07.Sampleline.cs** - 26 l?nh (27%)
2. **01.CAD.cs** - 23 l?nh (23%)
3. **08.Sectionview.cs** - 20 l?nh (20%)
4. **05.Point.cs** - 5 l?nh (5%)
5. **04.PipeAndStructures.cs** - 5 l?nh (5%)

---

## ?? **C?P NH?T M?I**

### **Ngày 19/12/2024:**
- ? **Thêm file m?i:** `Civil Tool\Property Sets.cs` v?i 2 l?nh
- ? **Thêm class h? tr?:** `Extensions\07.PropertySetUtils.cs`
- ?? **Ch?c n?ng m?i:** X? lý 3D Solid và Property Sets
- ?? **T?ng c?ng:** 98 l?nh (+2 l?nh)

### **Các l?nh m?i:**
- `AT_Solid_Set_PropertySet` - Thi?t l?p Property Set cho 3D Solid t? layer name
- `AT_Solid_Show_Info` - Hi?n th? thông tin chi ti?t v? 3D Solid

---

## ?? **GHI CHÚ**

### **Quy ??c ??t Tên L?nh:**
- **CT** = Civil Tool
- **AT** = Acad Tool (bao g?m c? Property Sets)
- **PA** = Parcel
- **PI** = Pipe
- **PO** = Point
- **P** = Profile
- **S** = Sampleline
- **SV** = SectionView
- **S** = Surface

### **Các File Không Có CommandMethod:**
- Extensions\01.ActiveDocument.cs
- Extensions\02.UtilitiesCAD.cs
- Extensions\03.UtilitiesC3D.cs
- Extensions\04.UserInput.cs
- Extensions\05.CsharpUtilities.cs
- Extensions\06.Civil3DExtensions.cs
- Extensions\07.PropertySetUtils.cs ? **M?I**
- Menu form\TestForm.cs
- Menu form\TestForm.Designer.cs
- Civil Tool\01.Corridor.Designer.cs
- Civil Tool\05.Point.Designer.cs

### **??c ?i?m N?i B?t:**
- D? án t?p trung ch? y?u vào Civil 3D (68% l?nh)
- Nhi?u l?nh x? lý Sampleline và Sectionview
- Có h? th?ng l?nh AutoCAD general ??y ??
- Giao di?n Ribbon menu hoàn ch?nh
- **M?I:** H? tr? x? lý 3D Solid và Property Sets ?

---

**?? Báo cáo ???c c?p nh?t t? ??ng b?i GitHub Copilot**