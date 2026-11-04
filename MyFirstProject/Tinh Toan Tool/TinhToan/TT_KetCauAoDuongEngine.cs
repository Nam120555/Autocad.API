using System;
using System.Collections.Generic;

namespace TinhToan
{
    public class PavementLayer
    {
        public double H { get; set; } // Thickness in cm
        public double E { get; set; } // Modulus in MPa
        public string MaterialName { get; set; } = string.Empty;
        public bool IsAsphalt { get; set; } = false;
        
        public PavementLayer(double h, double e, string mat, bool isAsphalt)
        {
            H = h; E = e; MaterialName = mat; IsAsphalt = isAsphalt;
        }
    }

    public class PavementAnalysisResult
    {
        // 1. Tiêu chuẩn võng
        public double Echung { get; set; }
        public double Eyc { get; set; }
        public bool VongPassed { get; set; }
        
        // 2. Tiêu chuẩn trượt
        public double TauMax { get; set; }
        public double TauCp { get; set; }
        public bool TruotPassed { get; set; }
        
        // 3. Kéo uốn
        public double SigmaKu { get; set; }
        public double SigmaCp { get; set; }
        public bool KeoUonPassed { get; set; }
    }

    public class TT_KetCauAoDuongEngine
    {
        /// <summary>
        /// Phân tích cường độ áo đường mềm theo cơ học đàn hồi 
        /// (mô phỏng phương pháp tra toán đồ của 22TCN 211-06)
        /// </summary>
        public static PavementAnalysisResult Analyze(
            double P, double p, double D, double Ntk, 
            double E0, 
            List<PavementLayer> layers)
        {
            var res = new PavementAnalysisResult();
            
            // ==========================================
            // 1. TÍNH TOÁN MÔ ĐUN TƯƠNG ĐƯƠNG (E_chung)
            // Lời giải sử dụng phương pháp Odemark thay cho bảng tra toán đồ
            // ==========================================
            
            // Xây dựng mô đun yêu cầu theo số trục Ntk (Quy luật bán thực nghiệm đơn giản hóa)
            // Công thức hối quy phổ biến cho Áo đường cấp cao A1:
            res.Eyc = 150 + 15 * Math.Log10(Ntk > 0 ? Ntk : 1);
            if (res.Eyc < 120) res.Eyc = 120;
            if (res.Eyc > 250) res.Eyc = 250;

            // Tính E chung từ dưới lên bằng Odemark Equivalence (h_equiv = h * cbrt(E1/E2))
            double e_td = E0; 
            for (int i = layers.Count - 1; i >= 0; i--)
            {
                var lay = layers[i];
                if (lay.H <= 0 || lay.E <= 0) continue;
                
                // Trọng số ảnh hưởng lớp H trên (f) phụ thuộc E_layer / E_nen
                double ratio = lay.E / e_td;
                double h_equiv = lay.H * Math.Pow(ratio, 0.3333);
                
                // Mô phỏng Nomogram:
                double factor = h_equiv / D; // Chiều dày tương đương so với vệt bánh xe dọc
                
                // Hàm S-curve khép kín tương đương toán đồ E_chung / E_1 của Liên Xô cũ
                double e_ratio = 1.0 - Math.Exp(-0.8 * factor); 
                
                // Cập nhật tầng tương đương mới
                e_td = e_td + e_ratio * (lay.E - e_td);
            }
            
            // Tích hợp hệ số điều chỉnh (nếu có các lớp BTN chất lượng cao)
            res.Echung = e_td * 1.05; 
            res.VongPassed = (res.Echung >= res.Eyc * 1.0); // K_vdh >= 1.0

            // ==========================================
            // 2. KHẢ NĂNG CHỊU TRƯỢT NỀN ĐẤT (Tau_max <= Tau_cp)
            // Dựa trên lý thuyết tập trung ứng suất Boussinesq
            // ==========================================
            
            double total_H_equiv = 0;
            foreach (var l in layers) total_H_equiv += l.H * Math.Pow(l.E / E0, 0.33);
            
            // Tỷ lệ truyền ứng suất z/D
            double z_D = total_H_equiv / D;
            // Xấp xỉ ứng suất trong nền đất
            double active_stress = p * (1.0 - 1.0 / Math.Pow(1.0 + (z_D * z_D), 1.5));
            // Ứng suất cắt lớn nhất khoảng 0.3 lần active stress (phụ thuộc vệt bánh)
            res.TauMax = active_stress * 0.35; 
            
            // Sức chống cắt nền đất (Tau cp) giả định C = 0.02, Phi = 20 độ, T_cp = C + tang(Phi)*sigma_z
            res.TauCp = 0.025 + Math.Tan(20 * Math.PI / 180) * 0.02; // MPa
            // Đổi ra daN/cm2 (1 MPa = 10 daN/cm2)
            res.TauMax = res.TauMax * 10;
            res.TauCp  = res.TauCp * 10;

            res.TruotPassed = (res.TauMax <= res.TauCp);

            // ==========================================
            // 3. KHẢ NĂNG CHỊU KÉO UỐN (LỚP LIỀN KHỐI BTN)
            // ==========================================
            
            // Tính cho lớp đúc nguyên khối (BTN nền)
            double h_btn = 0;
            double count_asphalt = 0;
            foreach (var l in layers) {
                if (l.IsAsphalt) { h_btn += l.H; count_asphalt++; }
            }

            if (count_asphalt == 0)
            {
                // Không có lớp bê tông nhựa -> Bỏ qua kiểm toán kéo uốn
                res.KeoUonPassed = true;
                res.SigmaKu = 0;
                res.SigmaCp = 999;
            }
            else
            {
                // Mật độ dàn tải của bánh xe
                res.SigmaKu = p * (2.0 / Math.PI) * Math.Pow(D / (2 * h_btn), 0.85); 
                // Cường độ kéo dãn cho phép của BTN
                double Ku = 1.2; // Hệ số mỏi
                res.SigmaCp = 3.5 / Ku; // 3.5 MPa là cường độ mẫu BTN chuẩn
                
                res.KeoUonPassed = (res.SigmaKu <= res.SigmaCp);
            }

            return res;
        }
    }
}
