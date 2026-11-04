// Phát sinh cọc và đổi tên cọc theo quy cách Civil
// Quy cách: DT (đầu), CT (cuối), H1/H2 (100m), KM1/KM2 (1000m), 1~N (còn lại)
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using Autodesk.Civil.ApplicationServices;
using MyFirstProject.Extensions;

// Alias giải quyết xung đột namespace
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using ContentAlignment = System.Drawing.ContentAlignment;
using FontStyle = System.Drawing.FontStyle;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(MyFirstProject.PhatSinhCocDoiTenCoc))]

namespace MyFirstProject
{
    #region Form


    #endregion

    #region Command

    public class PhatSinhCocDoiTenCoc
    {
        /// <summary>
        /// Lệnh gộp: Phát sinh cọc hoặc Đổi tên cọc
        /// </summary>
        [CommandMethod("CTS_PhatSinhCoc")]
        public static void CTSPhatSinhCoc()
        {
            using CocTuyenDuongForm form = new();
            var result = AcadApp.ShowModalDialog(form);

            if (result != DialogResult.OK || !form.IsConfirmed) return;

            if (form.IsCogoPointMode)
            {
                ExecuteFromCogoPoint(form);
                return;
            }
            if (form.IsTableMode)
            {
                ExecuteFromTable(form);
                return;
            }

            if (form.SelectedAlignmentIds.Count == 0) return;

            if (form.IsPhatSinh)
            {
                foreach(var alId in form.SelectedAlignmentIds)
                    ExecutePhatSinh(form, alId);
            }
            else
            {
                foreach(var alId in form.SelectedAlignmentIds)
                    ExecuteDoiTen(form, alId);
            }
        }

        /// <summary>
        /// Phát sinh cọc mới dọc tuyến
        /// </summary>
        private static void ExecutePhatSinh(CocTuyenDuongForm form, ObjectId alId)
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                Alignment? alignment = tr.GetObject(alId, OpenMode.ForWrite) as Alignment;
                if (alignment == null) { tr.Commit(); return; }

                // Tạo hoặc lấy SampleLineGroup
                ObjectId slgId;
                if (alignment.GetSampleLineGroupIds().Count == 0)
                    slgId = SampleLineGroup.Create(alignment.Name, alId);
                else
                    slgId = alignment.GetSampleLineGroupIds()[0];

                var itemsForThisAl = form.PreviewItems.Where(x => x.AlignmentId == alId).ToList();

                // Tạo SampleLine từ preview
                int created = 0;
                foreach (var item in itemsForThisAl)
                {
                    try
                    {
                        string tempName = "z_" + Guid.NewGuid().ToString("N").Substring(0, 10);
                        ObjectId newId = CreateSamplelineWithWidth(tr, tempName, slgId, alignment, item.Station, form.BeRongTrai, form.BeRongPhai);
                        created++;
                    }
                    catch { /* Bỏ qua cọc trùng vị trí */ }
                }

                // Đổi tên tất cả cọc trong group theo quy cách
                SampleLineGroup? slGroup = tr.GetObject(slgId, OpenMode.ForWrite) as SampleLineGroup;
                if (slGroup != null)
                    RenameSampleLines(tr, alignment, slGroup, form.UseDT_CT, form.UseH, form.UseKM, form.UseCurve, form.UsePrefix, form.PrefixValue, form.StrND, form.StrTD, form.StrP, form.StrTC, form.StrNC);

                // Áp dụng Style cho SampleLine
                ApplyStyleAndLabel(tr, slgId, form.SelectedSampleLineStyle, form.SelectedLabelStyle);

                A.Ed.WriteMessage($"\n✅ Đã phát sinh {created}/{itemsForThisAl.Count} cọc trên tuyến {alignment.Name}");
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi tuyến {alId}: {ex.Message}");
            }
        }

        /// <summary>
        /// Đổi tên cọc hiện có
        /// </summary>
        private static void ExecuteDoiTen(CocTuyenDuongForm form, ObjectId alId)
        {
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                Alignment? alignment = tr.GetObject(alId, OpenMode.ForRead) as Alignment;
                if (alignment == null) { tr.Commit(); return; }

                ObjectIdCollection slgIds = alignment.GetSampleLineGroupIds();
                if (slgIds.Count == 0)
                {
                    A.Ed.WriteMessage($"\n⚠ Tuyến {alignment.Name} chưa có nhóm cọc!");
                    tr.Commit();
                    return;
                }

                SampleLineGroup? slGroup = tr.GetObject(slgIds[0], OpenMode.ForWrite) as SampleLineGroup;
                if (slGroup == null) { tr.Commit(); return; }

                int count = RenameSampleLines(tr, alignment, slGroup, form.UseDT_CT, form.UseH, form.UseKM, form.UseCurve, form.UsePrefix, form.PrefixValue, form.StrND, form.StrTD, form.StrP, form.StrTC, form.StrNC);

                // Áp dụng Style cho SampleLine
                ApplyStyleAndLabel(tr, slgIds[0], form.SelectedSampleLineStyle, form.SelectedLabelStyle);

                A.Ed.WriteMessage($"\n✅ Đã đổi tên {count} cọc trên tuyến {alignment.Name}");
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi tuyến {alId}: {ex.Message}");
            }
        }

        /// <summary>
        /// Phát sinh cọc từ nhóm CogoPoint
        /// </summary>
        private static void ExecuteFromCogoPoint(CocTuyenDuongForm form)
        {
            try
            {
                // Pick trên bản vẽ sau khi form đóng
                ObjectId cogoPointId = UserInput.GCogoPointId("\n Chọn 1 CogoPoint thuộc nhóm điểm cần phát sinh cọc: \n");
                ObjectId alignmentId = UserInput.GAlignmentId("\n Chọn tim đường cần phát sinh cọc:");

                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                CogoPoint? cogoPoint = tr.GetObject(cogoPointId, OpenMode.ForRead) as CogoPoint;
                if (cogoPoint == null) { tr.Commit(); return; }

                ObjectId pointGroupId = cogoPoint.PrimaryPointGroupId;
                ObjectIdCollection pointIds = UtilitiesC3D.GPointIdsFromPointGroup(pointGroupId);

                Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForWrite) as Alignment;
                if (alignment == null) { tr.Commit(); return; }

                // Tạo hoặc lấy SampleLineGroup
                ObjectId slgId;
                if (alignment.GetSampleLineGroupIds().Count == 0)
                    slgId = SampleLineGroup.Create(alignment.Name, alignmentId);
                else
                    slgId = alignment.GetSampleLineGroupIds()[0];

                double station = 0, offset = 0;
                int created = 0;
                foreach (ObjectId pointId in pointIds)
                {
                    CogoPoint? pt = tr.GetObject(pointId, OpenMode.ForRead) as CogoPoint;
                    if (pt == null) continue;
                    alignment.StationOffset(pt.Easting, pt.Northing, ref station, ref offset);
                    try
                    {
                        string tempName = "z_" + Guid.NewGuid().ToString("N").Substring(0, 10);
                        CreateSamplelineWithWidth(tr, tempName, slgId, alignment, station, form.BeRongTrai, form.BeRongPhai);
                        created++;
                    }
                    catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                }

                // Rename theo quy cách
                SampleLineGroup? slGroup = tr.GetObject(slgId, OpenMode.ForWrite) as SampleLineGroup;
                if (slGroup != null)
                    RenameSampleLines(tr, alignment, slGroup, form.UseDT_CT, form.UseH, form.UseKM, form.UseCurve, form.UsePrefix, form.PrefixValue, form.StrND, form.StrTD, form.StrP, form.StrTC, form.StrNC);

                // Áp dụng Style
                ApplyStyleAndLabel(tr, slgId, form.SelectedSampleLineStyle, form.SelectedLabelStyle);

                A.Ed.WriteMessage($"\n✅ Đã phát sinh {created} cọc từ CogoPoint trên tuyến {alignment.Name}");
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Phát sinh cọc từ bảng tọa độ (AutoCAD Table)
        /// </summary>
        private static void ExecuteFromTable(CocTuyenDuongForm form)
        {
            try
            {
                ObjectId bangId = UserInput.GTable("Chọn bảng tọa độ cọc có lý trình:");
                ObjectId alignmentId = UserInput.GAlignmentId("Chọn tim tuyến cần bổ sung cọc:");

                int soTenCoc = form.ColTenCoc;
                int soLyTrinh = form.ColLyTrinh;
                int soHangCoc = form.RowStart;

                using Transaction tr = A.Db.TransactionManager.StartTransaction();
                Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForWrite) as Alignment;
                Autodesk.AutoCAD.DatabaseServices.Table? bang = tr.GetObject(bangId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Table;
                if (alignment == null || bang == null) { tr.Commit(); return; }

                // Tạo hoặc lấy SampleLineGroup
                ObjectId slgId;
                if (alignment.GetSampleLineGroupIds().Count == 0)
                    slgId = SampleLineGroup.Create(alignment.Name, alignment.ObjectId);
                else
                    slgId = alignment.GetSampleLineGroupIds()[0];

                List<string> listTenCoc = new();
                List<double> listLyTrinh = new();
                int soCocThem = 1;

                for (int i = soHangCoc - 1; i < bang.Rows.Count; i++)
                {
                    string tenCoc = bang.Cells[i, soTenCoc - 1].GetTextString(Autodesk.AutoCAD.DatabaseServices.FormatOption.IgnoreMtextFormat);
                    listTenCoc.Add(tenCoc);
                    // Check duplicate
                    for (int j = 0; j < listTenCoc.Count - 1; j++)
                    {
                        if (tenCoc == listTenCoc[j])
                        {
                            listTenCoc[j] = tenCoc + soCocThem.ToString();
                            soCocThem++;
                        }
                    }
                    listLyTrinh.Add(Convert.ToDouble(bang.Cells[i, soLyTrinh - 1].GetTextString(Autodesk.AutoCAD.DatabaseServices.FormatOption.IgnoreMtextFormat)));
                }

                if (listLyTrinh.Count > 0)
                    listLyTrinh[listLyTrinh.Count - 1] -= 0.01;

                int created = 0;
                for (int i = 0; i < listTenCoc.Count; i++)
                {
                    try
                    {
                        CreateSamplelineWithWidth(tr, listTenCoc[i], slgId, alignment, listLyTrinh[i], form.BeRongTrai, form.BeRongPhai);
                        created++;
                    }
                    catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
                }

                A.Ed.WriteMessage($"\n✅ Đã phát sinh {created}/{listTenCoc.Count} cọc từ bảng trên tuyến {alignment.Name}");
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
            }
        }

        #region Core Logic

        /// <summary>
        /// Đổi tên tất cả cọc trong SampleLineGroup theo quy cách:
        /// DT (đầu), CT (cuối), KM (1000m), H (100m), TĐ, P, TC, 1~N (còn lại)
        /// </summary>
        private static int RenameSampleLines(Transaction tr, Alignment al, SampleLineGroup slGroup,
            bool useDtCt, bool useH, bool useKm, bool useCurve, bool usePrefix, string prefixValue,
            string strND, string strTD, string strP, string strTC, string strNC)
        {
            Dictionary<double, string> specialStations = new();
            if (useCurve && al != null)
            {
                var curveList = new List<(double startSta, AlignmentEntity entity)>();
                foreach (AlignmentEntity entity in al.Entities)
                {
                    if (entity is AlignmentSCS scs)
                        curveList.Add((scs.SpiralIn.StartStation, entity));
                    else if (entity is AlignmentArc arc)
                        curveList.Add((arc.StartStation, entity));
                    else if (entity is AlignmentSpiral spiral)
                        curveList.Add((spiral.StartStation, entity));
                }

                curveList = curveList.OrderBy(c => c.startSta).ToList();

                int curveIndex = 1; // Đánh số đường cong bắt đầu từ 1
                foreach (var item in curveList)
                {
                    AlignmentEntity entity = item.entity;
                    if (entity is AlignmentSCS scs)
                    {
                        // === ĐƯỜNG CONG CHUYỂN TIẾP SCS (compound: SpiralIn → Arc → SpiralOut) ===
                        string idx = curveIndex.ToString();
                        AddStationStatic(specialStations, scs.SpiralIn.StartStation, $"{strND}{idx}");
                        AddStationStatic(specialStations, scs.SpiralIn.EndStation, $"{strTD}{idx}");
                        double pStation = (scs.Arc.StartStation + scs.Arc.EndStation) / 2.0;
                        AddStationStatic(specialStations, pStation, $"{strP}{idx}");
                        AddStationStatic(specialStations, scs.Arc.EndStation, $"{strTC}{idx}");
                        AddStationStatic(specialStations, scs.SpiralOut.EndStation, $"{strNC}{idx}");
                        curveIndex++;
                    }
                    else if (entity is AlignmentArc arc)
                    {
                        // === CUNG TRÒN ĐƠN (không nằm trong SCS) ===
                        string idx = curveIndex.ToString();
                        AddStationStatic(specialStations, arc.StartStation, $"{strTD}{idx}");
                        AddStationStatic(specialStations, (arc.StartStation + arc.EndStation) / 2.0, $"{strP}{idx}");
                        AddStationStatic(specialStations, arc.EndStation, $"{strTC}{idx}");
                        curveIndex++;
                    }
                    else if (entity is AlignmentSpiral spiral)
                    {
                        // === SPIRAL ĐƠN (không nằm trong SCS) ===
                        string idx = curveIndex.ToString();
                        AddStationStatic(specialStations, spiral.StartStation, $"{strND}{idx}");
                        AddStationStatic(specialStations, spiral.EndStation, $"{strNC}{idx}");
                        curveIndex++;
                    }
                }
            }

            // Thu thập và sắp xếp theo station
            List<SampleLine> sorted = new();
            foreach (ObjectId slId in slGroup.GetSampleLineIds())
            {
                SampleLine? sl = tr.GetObject(slId, OpenMode.ForWrite) as SampleLine;
                if (sl != null) sorted.Add(sl);
            }
            sorted = sorted.OrderBy(s => s.Station).ToList();

            if (sorted.Count == 0) return 0;

            // Bước 1: Reset tên tạm (tránh trùng khi rename)
            for (int i = 0; i < sorted.Count; i++)
                sorted[i].Name = "z_" + Guid.NewGuid().ToString("N").Substring(0, 10);

            // Bước 2: Đặt tên theo quy cách
            int normalIndex = 1;
            for (int i = 0; i < sorted.Count; i++)
            {
                double station = sorted[i].Station;
                string newName;

                if (useDtCt && i == 0)
                    newName = "DT";
                else if (useDtCt && i == sorted.Count - 1)
                    newName = "CT";
                else if (specialStations.Any(kvp => Math.Abs(kvp.Key - station) < 0.01))
                    newName = specialStations.First(kvp => Math.Abs(kvp.Key - station) < 0.01).Value;
                else if (useKm && IsMultipleOf(station, 1000))
                    newName = $"KM{(int)Math.Round(station / 1000)}";
                else if (useH && IsMultipleOf(station, 100))
                {
                    int hVal = (int)Math.Round(station / 100);
                    newName = $"H{(useKm ? hVal % 10 : hVal)}";
                }
                else
                {
                    newName = (usePrefix ? prefixValue : "") + normalIndex.ToString();
                    normalIndex++;
                }

                bool nameSet = false;
                string finalName = newName;
                int retryCount = 0;
                while (!nameSet)
                {
                    try
                    {
                        sorted[i].Name = finalName;
                        nameSet = true;
                    }
                    catch
                    {
                        retryCount++;
                        finalName = $"{newName}_{retryCount}";
                        if (retryCount > 100) break; // Fallback an toàn
                    }
                }
            }

            return sorted.Count;
        }

        private static bool IsMultipleOf(double station, double divisor)
        {
            double remainder = station % divisor;
            return remainder < 0.01 || (divisor - remainder) < 0.01;
        }

        /// <summary>
        /// Thêm station đặc biệt vào dictionary (static version cho RenameSampleLines)
        /// </summary>
        private static void AddStationStatic(Dictionary<double, string> dict, double station, string name)
        {
            bool exists = dict.Keys.Any(k => Math.Abs(k - station) < 0.01);
            if (!exists)
                dict[station] = name;
        }

        /// <summary>
        /// Tạo SampleLine với bề rộng trái/phải tùy chỉnh
        /// </summary>
        private static ObjectId CreateSamplelineWithWidth(Transaction tr, string name, ObjectId slgId, Alignment alignment, double station, double widthLeft, double widthRight)
        {
            Autodesk.AutoCAD.Geometry.Point2dCollection points = new();
            double easting = 0, northing = 0;

            // Bên trái (giá trị âm)
            alignment.PointLocation(station, -widthLeft, ref easting, ref northing);
            points.Add(new Autodesk.AutoCAD.Geometry.Point2d(easting, northing));

            // Bên phải (giá trị dương)
            alignment.PointLocation(station, widthRight, ref easting, ref northing);
            points.Add(new Autodesk.AutoCAD.Geometry.Point2d(easting, northing));

            ObjectId slId = SampleLine.Create(name, slgId, points);
            return slId;
        }

        /// <summary>
        /// Áp dụng Style dao diện và Label cho tất cả SampleLine trong group
        /// </summary>
        private static void ApplyStyleAndLabel(Transaction tr, ObjectId slgId, string styleName, string labelStyleName)
        {
            CivilDocument cdoc = CivilApplication.ActiveDocument;

            // 1. Áp dụng SampleLine Style
            if (!string.IsNullOrEmpty(styleName))
            {
                try
                {
                    SampleLineGroup? slGroup = tr.GetObject(slgId, OpenMode.ForRead) as SampleLineGroup;
                    if (slGroup != null)
                    {
                        foreach (ObjectId slId in slGroup.GetSampleLineIds())
                        {
                            SampleLine? sl = tr.GetObject(slId, OpenMode.ForWrite) as SampleLine;
                            if (sl != null)
                            {
                                try { sl.StyleName = styleName; }
                                catch { /* Style không tồn tại — bỏ qua */ }
                            }
                        }
                    }
                }
                catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
            }

            // 2. Tạo Label Group nếu có chọn Label Style
            if (!string.IsNullOrEmpty(labelStyleName))
            {
                try
                {
                    ObjectId labelStyleId = ObjectId.Null;
                    foreach (ObjectId lsId in cdoc.Styles.LabelStyles.SampleLineLabelStyles.LabelStyles)
                    {
                        var ls = (StyleBase)tr.GetObject(lsId, OpenMode.ForRead);
                        if (ls.Name.Equals(labelStyleName, StringComparison.OrdinalIgnoreCase))
                        {
                            labelStyleId = lsId;
                            break;
                        }
                    }

                    if (labelStyleId != ObjectId.Null)
                    {
                        SampleLineLabelGroup.Create(slgId, labelStyleId);
                    }
                }
                catch (global::System.Exception __nxEx) { global::MyFirstProject.Core.NxLog.Swallow(__nxEx); }
            }
        }

        #endregion
    }

    #endregion
}
