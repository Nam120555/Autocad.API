using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using MyFirstProject.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadColor = Autodesk.AutoCAD.Colors.Color;
using AcadDb = Autodesk.AutoCAD.DatabaseServices;
using ColorMethod = Autodesk.AutoCAD.Colors.ColorMethod;

namespace MyFirstProject
{
    /// <summary>
    /// Lệnh vẽ đường bao Material Section trong SectionView
    /// Layer được đặt tên theo Material List
    /// </summary>
    public class LayDuLieuCTSV
    {
        #region Bảng màu cho các loại Material

        /// <summary>
        /// Bảng màu cho các loại Material thường gặp
        /// </summary>
        private static readonly Dictionary<string, short> MaterialColors = new()
        {
            // Đào - màu nâu/đỏ
            { "ĐÀO", 30 },      // Cam
            { "DAO", 30 },
            { "CUT", 30 },
            
            // Đắp - màu xanh dương
            { "ĐẮP", 5 },       // Xanh dương
            { "DAP", 5 },
            { "FILL", 5 },
            
            // Mặt đường - màu xám
            { "MẶT ĐƯỜNG", 8 },  // Xám
            { "MAT DUONG", 8 },
            { "PAVEMENT", 8 },
            
            // Móng đường - màu nâu
            { "MÓNG", 40 },
            { "MONG", 40 },
            { "BASE", 40 },
            
            // Bê tông - màu trắng
            { "BÊ TÔNG", 7 },
            { "BE TONG", 7 },
            { "CONCRETE", 7 },
            
            // Đá - màu xám đậm
            { "ĐÁ", 251 },
            { "DA", 251 },
            { "ROCK", 251 },
        };

        /// <summary>
        /// Màu mặc định cho Material không có trong bảng
        /// </summary>
        private static short DefaultMaterialColor = 3; // Xanh lá

        #endregion

        #region Lệnh chính

        /// <summary>
        /// Lệnh vẽ đường bao Material Section trong tất cả SectionView
        /// </summary>
        [CommandMethod("CTSV_VeDuongBaoMaterial")]
        public static void CTSVVeDuongBaoMaterial()
        {
            using AcadDb.Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                A.Ed.WriteMessage("\n╔════════════════════════════════════════════════════════════════╗");
                A.Ed.WriteMessage("\n║      VẼ ĐƯỜNG BAO MATERIAL SECTION TRONG SECTIONVIEW           ║");
                A.Ed.WriteMessage("\n╚════════════════════════════════════════════════════════════════╝");

                // 1. Lấy danh sách Alignment có SampleLineGroup
                List<AlignmentInfoLocal> alignments = GetAlignmentsWithSampleLineGroups(tr);
                if (alignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Không tìm thấy Alignment có SampleLineGroup!");
                    return;
                }

                // 2. Hiển thị form chọn Alignment
                using FormChonAlignmentLocal formChon = new(alignments);
                if (AcadApp.ShowModalDialog(formChon) != DialogResult.OK)
                {
                    A.Ed.WriteMessage("\nĐã hủy lệnh.");
                    return;
                }

                if (formChon.SelectedAlignments.Count == 0)
                {
                    A.Ed.WriteMessage("\n⚠️ Chưa chọn Alignment nào!");
                    return;
                }

                int totalBoundaries = 0;

                // 3. Vẽ đường bao cho từng Alignment được chọn
                foreach (var alignInfo in formChon.SelectedAlignments)
                {
                    A.Ed.WriteMessage($"\n\n📐 Đang xử lý: {alignInfo.Name}");
                    
                    int count = DrawMaterialBoundariesForAlignment(tr, alignInfo.SampleLineGroupId);
                    totalBoundaries += count;
                    
                    A.Ed.WriteMessage($"\n   ✅ Đã vẽ {count} đường bao");
                }

                tr.Commit();
                
                A.Ed.WriteMessage($"\n\n═══════════════════════════════════════════════════════════════");
                A.Ed.WriteMessage($"\n✅ HOÀN THÀNH: Đã vẽ tổng cộng {totalBoundaries} đường bao Material");
                A.Ed.WriteMessage($"\n═══════════════════════════════════════════════════════════════\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n❌ Lỗi: {ex.Message}");
                A.Ed.WriteMessage($"\n   Stack: {ex.StackTrace}");
            }
        }

        #endregion

        #region Hàm vẽ đường bao

        /// <summary>
        /// Vẽ đường bao Material cho một SampleLineGroup
        /// </summary>
        /// <param name="tr">Transaction hiện tại</param>
        /// <param name="sampleLineGroupId">ID của SampleLineGroup</param>
        /// <returns>Số lượng đường bao đã vẽ</returns>
        private static int DrawMaterialBoundariesForAlignment(AcadDb.Transaction tr, AcadDb.ObjectId sampleLineGroupId)
        {
            int totalDrawn = 0;

            try
            {
                // Lấy SampleLineGroup
                SampleLineGroup? slg = tr.GetObject(sampleLineGroupId, AcadDb.OpenMode.ForRead) as SampleLineGroup;
                if (slg == null) return 0;

                // Lấy ModelSpace để thêm polyline
                AcadDb.BlockTableRecord modelSpace = tr.GetObject(
                    A.Db.CurrentSpaceId, AcadDb.OpenMode.ForWrite) as AcadDb.BlockTableRecord ??
                    throw new System.Exception("Không thể mở ModelSpace");

                // Lấy SectionViewGroup
                SectionViewGroupCollection svgCollection = slg.SectionViewGroups;
                if (svgCollection.Count == 0)
                {
                    A.Ed.WriteMessage("\n   ⚠️ Chưa có SectionViewGroup. Hãy tạo Section Views trước.");
                    return 0;
                }

                SectionViewGroup svGroup = svgCollection[0];

                // Lấy danh sách Material từ Material Lists
                QTOMaterialListCollection materialLists = slg.MaterialLists;
                List<(string MaterialName, string ListName, Guid ListGuid, Guid MaterialGuid)> materialInfo = new();

                foreach (QTOMaterialList materialList in materialLists)
                {
                    try
                    {
                        string listName = materialList.Name;
                        Guid listGuid = materialList.Guid;

                        foreach (QTOMaterial material in materialList)
                        {
                            materialInfo.Add((material.Name, listName, listGuid, material.Guid));
                        }
                    }
                    catch { }
                }

                if (materialInfo.Count == 0)
                {
                    A.Ed.WriteMessage("\n   ⚠️ Không tìm thấy Material List nào.");
                    return 0;
                }

                // Tạo các layer cho Material (layer = tên Material)
                Dictionary<string, string> materialLayers = new();
                foreach (var (matName, listName, _, _) in materialInfo)
                {
                    if (!materialLayers.ContainsKey(matName))
                    {
                        // Layer = "C3D-MAT-" + tên Material
                        string layerName = $"C3D-MAT-{SanitizeLayerName(matName)}";
                        short colorIndex = GetMaterialColor(matName);
                        CreateLayerIfNotExists(tr, layerName, colorIndex);
                        materialLayers[matName] = layerName;
                    }
                }

                // Lấy các SectionView IDs
                AcadDb.ObjectIdCollection sectionViewIds = svGroup.GetSectionViewIds();

                // Duyệt qua từng SectionView
                foreach (AcadDb.ObjectId svId in sectionViewIds)
                {
                    SectionView? sectionView = tr.GetObject(svId, AcadDb.OpenMode.ForRead) as SectionView;
                    if (sectionView == null) continue;

                    // Lấy SampleLine tương ứng
                    AcadDb.ObjectId slId = sectionView.SampleLineId;
                    SampleLine? sampleLine = tr.GetObject(slId, AcadDb.OpenMode.ForRead) as SampleLine;
                    if (sampleLine == null) continue;

                    // Lấy vị trí SectionView
                    double sectionViewX = sectionView.Location.X;
                    double sectionViewY = sectionView.Location.Y;

                    // Lấy ElevationMin để tính offset Y
                    sectionView.UpgradeOpen();
                    sectionView.IsElevationRangeAutomatic = false;
                    double sectionViewElevMin = sectionView.ElevationMin;
                    sectionView.IsElevationRangeAutomatic = true;
                    sectionView.DowngradeOpen();

                    // Vẽ boundary cho mỗi Material
                    foreach (var (matName, listName, listGuid, matGuid) in materialInfo)
                    {
                        try
                        {
                            AcadDb.ObjectId materialSectionId = sampleLine.GetMaterialSectionId(listGuid, matGuid);
                            if (materialSectionId.IsNull || !materialSectionId.IsValid) continue;

                            Section? section = tr.GetObject(materialSectionId, AcadDb.OpenMode.ForRead) as Section;
                            if (section == null) continue;

                            SectionPointCollection sectionPoints = section.SectionPoints;
                            if (sectionPoints.Count < 3) continue;

                            // Tạo Polyline từ SectionPoints
                            AcadDb.Polyline pline = new();
                            pline.SetDatabaseDefaults();
                            pline.Layer = materialLayers[matName];

                            int vertexIndex = 0;
                            foreach (SectionPoint pt in sectionPoints)
                            {
                                // Chuyển đổi tọa độ Section sang World
                                // Công thức:
                                //   WorldX = SectionView.Location.X + SectionPoint.Offset
                                //   WorldY = SectionView.Location.Y + (SectionPoint.Elevation - ElevationMin)
                                double worldX = sectionViewX + pt.Location.X;
                                double worldY = sectionViewY + (pt.Location.Y - sectionViewElevMin);

                                pline.AddVertexAt(vertexIndex, new Point2d(worldX, worldY), 0, 0, 0);
                                vertexIndex++;
                            }

                            // Đóng polyline nếu điểm đầu và cuối không trùng
                            if (sectionPoints.Count >= 3)
                            {
                                Point3d first = sectionPoints[0].Location;
                                Point3d last = sectionPoints[sectionPoints.Count - 1].Location;

                                if (Math.Abs(first.X - last.X) > 0.001 || Math.Abs(first.Y - last.Y) > 0.001)
                                {
                                    pline.Closed = true;
                                }
                            }

                            // Thêm polyline vào ModelSpace
                            modelSpace.AppendEntity(pline);
                            tr.AddNewlyCreatedDBObject(pline, true);
                            totalDrawn++;
                        }
                        catch { }
                    }
                }
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\n   ⚠️ Lỗi: {ex.Message}");
            }

            return totalDrawn;
        }

        #endregion

        #region Hàm tiện ích

        /// <summary>
        /// Lấy màu cho Material dựa trên tên
        /// </summary>
        private static short GetMaterialColor(string materialName)
        {
            string upperName = materialName.ToUpperInvariant();

            foreach (var kvp in MaterialColors)
            {
                if (upperName.Contains(kvp.Key))
                {
                    return kvp.Value;
                }
            }

            return DefaultMaterialColor;
        }

        /// <summary>
        /// Chuẩn hóa tên layer (loại bỏ ký tự không hợp lệ)
        /// </summary>
        private static string SanitizeLayerName(string name)
        {
            // Các ký tự không hợp lệ trong tên layer AutoCAD
            char[] invalidChars = { '<', '>', '/', '\\', '"', ':', ';', '?', '*', '|', ',', '=', '`' };

            string result = name;
            foreach (char c in invalidChars)
            {
                result = result.Replace(c, '_');
            }

            // Giới hạn độ dài tên layer
            if (result.Length > 200)
            {
                result = result.Substring(0, 200);
            }

            return result;
        }

        /// <summary>
        /// Tạo layer nếu chưa tồn tại
        /// </summary>
        private static void CreateLayerIfNotExists(AcadDb.Transaction tr, string layerName, short colorIndex)
        {
            AcadDb.LayerTable? lt = tr.GetObject(A.Db.LayerTableId, AcadDb.OpenMode.ForRead) as AcadDb.LayerTable;
            if (lt == null) return;

            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                AcadDb.LayerTableRecord ltr = new()
                {
                    Name = layerName,
                    Color = AcadColor.FromColorIndex(ColorMethod.ByAci, colorIndex)
                };
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }

        /// <summary>
        /// Lấy danh sách Alignment có SampleLineGroup
        /// </summary>
        private static List<AlignmentInfoLocal> GetAlignmentsWithSampleLineGroups(AcadDb.Transaction tr)
        {
            List<AlignmentInfoLocal> result = new();

            try
            {
                AcadDb.ObjectIdCollection alignmentIds = A.Cdoc.GetAlignmentIds();

                foreach (AcadDb.ObjectId alignId in alignmentIds)
                {
                    try
                    {
                        Alignment? alignment = tr.GetObject(alignId, AcadDb.OpenMode.ForRead) as Alignment;
                        if (alignment == null) continue;

                        AcadDb.ObjectIdCollection slgIds = alignment.GetSampleLineGroupIds();
                        if (slgIds.Count == 0) continue;

                        foreach (AcadDb.ObjectId slgId in slgIds)
                        {
                            SampleLineGroup? slg = tr.GetObject(slgId, AcadDb.OpenMode.ForRead) as SampleLineGroup;
                            if (slg == null) continue;

                            result.Add(new AlignmentInfoLocal
                            {
                                Name = $"{alignment.Name} - {slg.Name}",
                                AlignmentId = alignId,
                                SampleLineGroupId = slgId
                            });
                        }
                    }
                    catch { }
                }
            }
            catch { }

            return result;
        }

        #endregion

        #region Data Classes

        /// <summary>
        /// Thông tin Alignment (local class để tránh phụ thuộc)
        /// </summary>
        private class AlignmentInfoLocal
        {
            public AcadDb.ObjectId AlignmentId { get; set; }
            public string Name { get; set; } = "";
            public AcadDb.ObjectId SampleLineGroupId { get; set; }
            public bool IsSelected { get; set; }
        }

        /// <summary>
        /// Form chọn Alignment đơn giản
        /// </summary>
        private class FormChonAlignmentLocal : Form
        {
            private CheckedListBox checkedListBox;
            private Button btnOK;
            private Button btnCancel;
            private List<AlignmentInfoLocal> alignmentInfos;

            public List<AlignmentInfoLocal> SelectedAlignments => 
                alignmentInfos.Where(a => a.IsSelected).ToList();

            public FormChonAlignmentLocal(List<AlignmentInfoLocal> alignments)
            {
                alignmentInfos = alignments;
                InitializeComponent();
                LoadData();
            }

            private void InitializeComponent()
            {
                Text = "Chọn Alignment để vẽ đường bao";
                Width = 450;
                Height = 350;
                StartPosition = FormStartPosition.CenterScreen;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;

                checkedListBox = new CheckedListBox
                {
                    Location = new System.Drawing.Point(12, 12),
                    Width = 410,
                    Height = 250,
                    CheckOnClick = true
                };
                checkedListBox.ItemCheck += (s, e) =>
                {
                    int index = e.Index;
                    alignmentInfos[index].IsSelected = e.NewValue == CheckState.Checked;
                };

                btnOK = new Button
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Location = new System.Drawing.Point(266, 270),
                    Width = 75
                };

                btnCancel = new Button
                {
                    Text = "Hủy",
                    DialogResult = DialogResult.Cancel,
                    Location = new System.Drawing.Point(347, 270),
                    Width = 75
                };

                Controls.Add(checkedListBox);
                Controls.Add(btnOK);
                Controls.Add(btnCancel);

                AcceptButton = btnOK;
                CancelButton = btnCancel;
            }

            private void LoadData()
            {
                checkedListBox.Items.Clear();
                foreach (var info in alignmentInfos)
                {
                    int idx = checkedListBox.Items.Add(info.Name);
                    checkedListBox.SetItemChecked(idx, true);
                    info.IsSelected = true;
                }
            }
        }

        #endregion
    }
}

