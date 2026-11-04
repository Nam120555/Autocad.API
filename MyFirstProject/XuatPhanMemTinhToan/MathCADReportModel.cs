using System;
using System.Collections.Generic;

namespace MyFirstProject.XuatPhanMemTinhToan
{
    public class MathCADReportModel
    {
        public string TenThongSo { get; set; } = string.Empty;
        public string KyHieu { get; set; } = string.Empty;
        public string CongThuc { get; set; } = string.Empty;
        public string GiaTri { get; set; } = string.Empty;
        public string DonVi { get; set; } = string.Empty;

        public MathCADReportModel(string tenThongSo, string kyHieu, string congThuc, string giaTri, string donVi)
        {
            TenThongSo = tenThongSo;
            KyHieu = kyHieu;
            CongThuc = congThuc;
            GiaTri = giaTri;
            DonVi = donVi;
        }
    }

    public class MathCADReportGroup
    {
        public string GroupName { get; set; } = string.Empty;
        public List<MathCADReportModel> Parameters { get; set; } = new List<MathCADReportModel>();
    }
}
