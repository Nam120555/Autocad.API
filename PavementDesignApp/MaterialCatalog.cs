using System.Collections.Generic;

namespace PavementDesignApp
{
    public class MaterialDef
    {
        public string Category { get; set; } // Phân loại (Nhóm BTN, Nhóm Móng đá, Nhóm Đất nền...)
        public string Name { get; set; }
        public double E { get; set; } // Elastic Modulus (MPa)
        public double C { get; set; } // Cohesion (MPa)
        public double Phi { get; set; } // Friction angle (Degrees)
        public double Rku { get; set; } // Bending resistance for Monolithic (MPa)
        public bool IsAsphalt { get; set; }

        public MaterialDef(string category, string name, double e, double c, double phi, double rku, bool isAsphalt)
        {
            Category = category;
            Name = name;
            E = e;
            C = c;
            Phi = phi;
            Rku = rku;
            IsAsphalt = isAsphalt;
        }

        public override string ToString() => Name;
    }

    public static class MaterialCatalog
    {
        public static List<MaterialDef> GetAll()
        {
            return new List<MaterialDef>
            {
                // ==========================
                // BẢNG B.1 - BÊ TÔNG NHỰA NÓNG (15*C) - TT VÕNG / KÉO UỐN KHI LẠNH
                // ==========================
                new MaterialDef("Bê tông nhựa 15°C", "BTN polyme chặt hạt nhỏ (15°C)", 4200, 0.45, 42, 6.0, true),
                new MaterialDef("Bê tông nhựa 15°C", "BTN polyme chặt hạt trung (15°C)", 3500, 0.40, 42, 5.0, true),
                new MaterialDef("Bê tông nhựa 15°C", "BTN chặt hạt nhỏ (15°C)", 3000, 0.35, 40, 4.5, true),
                new MaterialDef("Bê tông nhựa 15°C", "BTN chặt hạt trung (15°C)", 2500, 0.32, 40, 4.0, true),
                new MaterialDef("Bê tông nhựa 15°C", "BTN chặt hạt thô (15°C)", 2200, 0.30, 40, 3.5, true),
                new MaterialDef("Bê tông nhựa 15°C", "BTN rỗng hạt trung (15°C)", 1800, 0.25, 38, 2.5, true),
                new MaterialDef("Bê tông nhựa 15°C", "BTN rỗng hạt thô (15°C)", 1500, 0.22, 38, 2.0, true),

                // ==========================
                // BẢNG B.2 - BÊ TÔNG NHỰA NÓNG (30*C) - TIẾT DIỆN TÍNH TRƯỢT 
                // ==========================
                new MaterialDef("Bê tông nhựa 30°C", "BTN polyme chặt hạt nhỏ (30°C)", 1800, 0.20, 35, 2.0, true),
                new MaterialDef("Bê tông nhựa 30°C", "BTN polyme chặt hạt trung (30°C)", 1600, 0.18, 35, 1.8, true),
                new MaterialDef("Bê tông nhựa 30°C", "BTN chặt hạt nhỏ (30°C)", 1400, 0.16, 33, 1.5, true),
                new MaterialDef("Bê tông nhựa 30°C", "BTN chặt hạt thô (30°C)", 1200, 0.14, 34, 1.4, true),

                // ==========================
                // BẢNG C.1 - LỚP MÓNG GIA CỐ (XI MĂNG, VÔI)
                // ==========================
                new MaterialDef("Lớp móng gia cố", "Cấp phối đá dăm gia cố XM 5%", 1000, 0.08, 45, 0.8, false),
                new MaterialDef("Lớp móng gia cố", "Đất đồi gia cố XM 6-8%", 600, 0.06, 42, 0.4, false),
                new MaterialDef("Lớp móng gia cố", "Đất cát gia cố XM", 400, 0.05, 38, 0.3, false),

                // ==========================
                // BẢNG C.2 & C.3 - VẬT LIỆU TỰ NHIÊN DẠNG HẠT
                // ==========================
                new MaterialDef("Cấp phối đá", "Cấp phối đá dăm (CPĐD) Loại 1", 300, 0.025, 40, 0.0, false),
                new MaterialDef("Cấp phối đá", "Cấp phối đá dăm (CPĐD) Loại 2", 250, 0.020, 38, 0.0, false),
                new MaterialDef("Cấp phối tự nhiên", "Cấp phối đồi - sỏi cuội Tốt", 220, 0.018, 36, 0.0, false),
                new MaterialDef("Cấp phối tự nhiên", "Cấp phối sỏi suối", 180, 0.015, 35, 0.0, false),
                new MaterialDef("Cấp phối tự nhiên", "Đá dăm nước (Macadam)", 250, 0.020, 38, 0.0, false),

                // ==========================
                // BẢNG B.4 & ĐẤT NỀN - NỀN ĐẤT (Trạng thái trung bình, Ẩm)
                // ==========================
                new MaterialDef("Đất nền K98", "Nền đất Á cát (K98 - Ướt)", 42, 0.012, 22, 0.0, false),
                new MaterialDef("Đất nền K98", "Nền đất Á sét (K98 - Ướt)", 38, 0.016, 18, 0.0, false),
                new MaterialDef("Đất nền K98", "Nền đất Sét (K98 - Ướt)", 32, 0.025, 12, 0.0, false),
                new MaterialDef("Đất nền K98", "Nền đất Cát hạt trung (K98)", 45, 0.005, 32, 0.0, false),

                new MaterialDef("Đất nền K95", "Nền đất Á cát (K95 - Ướt)", 35, 0.010, 20, 0.0, false),
                new MaterialDef("Đất nền K95", "Nền đất Á sét (K95 - Ướt)", 30, 0.014, 16, 0.0, false),
                new MaterialDef("Đất nền K95", "Nền đất Sét (K95 - Ướt)", 26, 0.020, 10, 0.0, false)
            };
        }
    }
}
