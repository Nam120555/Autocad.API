using System;
using System.Collections.Generic;
using System.Text;

namespace PavementDesignApp
{
    public class PavementLayer
    {
        public double H { get; set; } // Thickness in cm
        public double E { get; set; } // Modulus in MPa
        public double C { get; set; } // Cohesion in MPa
        public double Phi { get; set; } // Internal friction angle (Degrees)
        public double Rku { get; set; } // Tensile strength
        public string MaterialName { get; set; } = string.Empty;
        public bool IsAsphalt { get; set; } = false;
        
        public PavementLayer(double h, double e, double c, double phi, double rku, string mat, bool isAsphalt)
        {
            H = h; E = e; C = c; Phi = phi; Rku = rku; MaterialName = mat; IsAsphalt = isAsphalt;
        }
    }

    public class PavementAnalysisResult
    {
        public double Echung { get; set; }
        public double Eyc { get; set; }
        public bool VongPassed { get; set; }
        
        public double TauMax { get; set; }
        public double TauCp { get; set; }
        public bool TruotPassed { get; set; }
        
        public double SigmaKu { get; set; }
        public double SigmaCp { get; set; }
        public bool KeoUonPassed { get; set; }

        public string DetailedLog { get; set; } = string.Empty;
    }

    public class PavementEngine
    {
        public static PavementAnalysisResult Analyze(
            double P, double p_val, double D, double Ntk, 
            double E0, double C0, double Phi0, 
            List<PavementLayer> layers)
        {
            var res = new PavementAnalysisResult();
            StringBuilder log = new StringBuilder();

            log.AppendLine("📌 BẮT ĐẦU KIỂM TOÁN CƯỜNG ĐỘ ÁO ĐƯỜNG MỀM 22TCN 211-06");
            log.AppendLine(new string('-', 50));
            log.AppendLine($"Tải trọng trục: P = {P} kN | Áp lực p = {p_val} daN/cm2 | D = {D} cm");
            log.AppendLine($"Trục xê tiêu chuẩn: Ntk = {Ntk} trục/ngày.đêm");
            log.AppendLine($"Nền đất: Eo = {E0} MPa | C = {C0} MPa | Phi = {Phi0} độ");
            log.AppendLine();

            // ----------------------------------------------------
            // 1. CHỈ TIÊU VÕNG ĐÀN HỒI (Odemark's Equivalent)
            // ----------------------------------------------------
            res.Eyc = 150 + 15 * Math.Log10(Ntk > 0 ? Ntk : 1);
            if (res.Eyc < 120) res.Eyc = 120;
            if (res.Eyc > 250) res.Eyc = 250;

            log.AppendLine($"[1] TÍNH TOÁN THEO TIÊU CHUẨN VÕNG ĐÀN HỒI");
            log.AppendLine($"-> E_yeu_cau (Eyc) = {res.Eyc:F1} MPa");

            double e_td = E0; 
            for (int i = layers.Count - 1; i >= 0; i--)
            {
                var lay = layers[i];
                if (lay.H <= 0 || lay.E <= 0) continue;
                
                double ratio = lay.E / e_td;
                double h_equiv = lay.H * Math.Pow(ratio, 0.3333);
                double factor = h_equiv / D; 
                double e_ratio = 1.0 - Math.Exp(-0.8 * factor); 
                e_td = e_td + e_ratio * (lay.E - e_td);

                log.AppendLine($"  + Lớp {i+1} ({lay.MaterialName}): H = {lay.H}cm, E = {lay.E}MPa. E_td gom xuống = {e_td:F1} MPa");
            }
            
            res.Echung = e_td * 1.05; 
            res.VongPassed = (res.Echung >= res.Eyc * 1.0);
            log.AppendLine($">> KẾT LUẬN: E_chung = {res.Echung:F1} MPa so với E_yc = {res.Eyc:F1} MPa -> {(res.VongPassed?"ĐẠT":"KHÔNG ĐẠT")}");
            log.AppendLine();

            // ----------------------------------------------------
            // 2. KHẢ NĂNG CHỊU TRƯỢT NỀN ĐẤT
            // ----------------------------------------------------
            log.AppendLine($"[2] TÍNH TOÁN THEO TIÊU CHUẨN TRƯỢT NỀN ĐẤT");
            double total_H_equiv = 0;
            foreach (var l in layers) total_H_equiv += l.H * Math.Pow(l.E / E0, 0.33);
            
            double z_D = total_H_equiv / D;
            double active_stress = p_val * (1.0 - 1.0 / Math.Pow(1.0 + (z_D * z_D), 1.5));
            res.TauMax = active_stress * 0.35; 
            log.AppendLine($"  + Tỷ lệ tương đương z/D = {z_D:F2}. Nội suy ứng suất tập trung Boussinesq T_max.");
            
            // Lấy đặc trưng cắt của Nền đất / CPĐD dưới cùng
            res.TauCp = C0 + Math.Tan(Phi0 * Math.PI / 180) * 0.02; // MPa
            res.TauMax = res.TauMax * 10;
            res.TauCp  = res.TauCp * 10;
            res.TruotPassed = (res.TauMax <= res.TauCp);

            log.AppendLine($"  + Sức kháng cắt Tau_cp = {res.TauCp:F3} daN/cm2 | Tau_max = {res.TauMax:F3} daN/cm2");
            log.AppendLine($">> KẾT LUẬN: {(res.TruotPassed?"ĐẠT":"KHÔNG ĐẠT")}");
            log.AppendLine();


            // ----------------------------------------------------
            // 3. KHẢ NĂNG CHỊU KÉO UỐN (LỚP ASPhalt)
            // ----------------------------------------------------
            log.AppendLine($"[3] TÍNH TOÁN THEO TIÊU CHUẨN KÉO UỐN BÊ TÔNG NHỰA");
            double h_btn = 0;
            double rku_avg = 0;
            int count_asphalt = 0;
            foreach (var l in layers) {
                if (l.IsAsphalt) { h_btn += l.H; rku_avg += l.Rku; count_asphalt++; }
            }

            if (count_asphalt == 0)
            {
                log.AppendLine("  + Khuôn đường không có Bê tông nhựa -> KHÔNG XÉT Kéo Uốn.");
                res.KeoUonPassed = true;
                res.SigmaKu = 0;
                res.SigmaCp = 999;
            }
            else
            {
                rku_avg = rku_avg / count_asphalt;
                res.SigmaKu = p_val * (2.0 / Math.PI) * Math.Pow(D / (2 * h_btn), 0.85); 
                double Ku = 1.2; 
                res.SigmaCp = rku_avg / Ku; 

                res.KeoUonPassed = (res.SigmaKu <= res.SigmaCp);
                log.AppendLine($"  + Chiều dày dãn nứt khối BTN H_btn = {h_btn} cm.");
                log.AppendLine($"  + Sigma kéo uốn phát sinh = {res.SigmaKu:F3} MPa.");
                log.AppendLine($"  + Sigma kéo cho phép (Rku / Ku) = {res.SigmaCp:F3} MPa.");
                log.AppendLine($">> KẾT LUẬN: {(res.KeoUonPassed?"ĐẠT":"KHÔNG ĐẠT")}");
            }

            log.AppendLine(new string('-', 50));
            log.AppendLine("HỆ THỐNG ĐÃ HOÀN TẤT PHÂN TÍCH.");
            
            res.DetailedLog = log.ToString();
            return res;
        }
    }
}
