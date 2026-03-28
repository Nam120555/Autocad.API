// 03.TongHopDoLuong.cs - Form gộp 8 lệnh AT_TongDoDai/AT_TongDienTich
// Thay vì 8 lệnh riêng → 1 form chọn chế độ
// Kế thừa CivilFormBase

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using MyFirstProject.Extensions;
using WinLabel = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using FontStyle = System.Drawing.FontStyle;
using ContentAlignment = System.Drawing.ContentAlignment;

[assembly: CommandClass(typeof(MyFirstProject.TongHopDoLuongCommands))]

namespace MyFirstProject
{
    // ══════════════════════════════════════════════════════════════
    // FORM: TỔNG HỢP ĐO LƯỜNG
    // ══════════════════════════════════════════════════════════════

    public class TongHopDoLuongForm : CivilFormBase
    {
        // ═══ Controls ═══
        private RadioButton rdoLength = null!;
        private RadioButton rdoArea = null!;
        private RadioButton rdoModeFull = null!;
        private RadioButton rdoModeReplace = null!;
        private RadioButton rdoModeReplace2 = null!;
        private RadioButton rdoModeCongThem = null!;

        // ═══ Kết quả ═══
        public bool IsConfirmed { get; private set; }
        public bool IsLength => rdoLength.Checked;
        public string Mode
        {
            get
            {
                if (rdoModeFull.Checked) return "Full";
                if (rdoModeReplace.Checked) return "Replace";
                if (rdoModeReplace2.Checked) return "Replace2";
                return "CongThem";
            }
        }

        public TongHopDoLuongForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            InitBaseTheme("📐  TỔNG HỢP ĐO LƯỜNG", 400, 400);

            // ═══ HEADER ═══
            var header = CreateHeader("TỔNG HỢP ĐO LƯỜNG");
            this.Controls.Add(header);

            var subtitle = CreateSubtitle("Gộp 8 lệnh tính tổng độ dài / diện tích");
            this.Controls.Add(subtitle);

            // ═══ NHÓM 1: LOẠI ĐO ═══
            var grpType = CreateGroupBox("📏 Loại đo lường", 20, 85, 350, 70);
            this.Controls.Add(grpType);

            rdoLength = new RadioButton
            {
                Text = "📏 Tổng Độ Dài (Length)",
                Location = new Point(15, 30),
                AutoSize = true,
                Checked = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 10f)
            };
            grpType.Controls.Add(rdoLength);

            rdoArea = new RadioButton
            {
                Text = "⬜ Tổng Diện Tích (Area)",
                Location = new Point(200, 30),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 10f)
            };
            grpType.Controls.Add(rdoArea);

            // ═══ NHÓM 2: CHẾ ĐỘ XUẤT ═══
            var grpMode = CreateGroupBox("📝 Chế độ xuất kết quả", 20, 165, 350, 140);
            this.Controls.Add(grpMode);

            int ry = 28;
            rdoModeFull = new RadioButton
            {
                Text = "🆕 Tạo text MỚI tại vị trí chọn (kèm đơn vị)",
                Location = new Point(15, ry),
                AutoSize = true, Checked = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpMode.Controls.Add(rdoModeFull);
            ry += 28;

            rdoModeReplace = new RadioButton
            {
                Text = "✏️ Ghi vào TEXT có sẵn (kèm đơn vị)",
                Location = new Point(15, ry),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpMode.Controls.Add(rdoModeReplace);
            ry += 28;

            rdoModeReplace2 = new RadioButton
            {
                Text = "✏️ Ghi vào TEXT có sẵn (KHÔNG đơn vị)",
                Location = new Point(15, ry),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpMode.Controls.Add(rdoModeReplace2);
            ry += 28;

            rdoModeCongThem = new RadioButton
            {
                Text = "➕ CỘNG THÊM vào text đã có giá trị",
                Location = new Point(15, ry),
                AutoSize = true,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 9.5f)
            };
            grpMode.Controls.Add(rdoModeCongThem);

            // ═══ HINT ═══
            var hint = CreateHintLabel(
                "💡 Chọn loại đo + chế độ xuất rồi bấm Thực Hiện",
                20, 315);
            this.Controls.Add(hint);

            // ═══ BOTTOM BAR ═══
            var bottomBar = CreateBottomBar(56);
            this.Controls.Add(bottomBar);

            var btnCancel = CreateButton("✕ Hủy", 10, 10, 170, 36, bottomBar);
            btnCancel.Click += (s, e) => { IsConfirmed = false; this.Close(); };

            var btnRun = CreateSuccessButton("▶ Thực Hiện", 190, 10, 170, 36, bottomBar);
            btnRun.Click += (s, e) => { IsConfirmed = true; this.Close(); };

            this.AcceptButton = btnRun;
            this.CancelButton = btnCancel;
        }
    }

    // ══════════════════════════════════════════════════════════════
    // LỆNH CHÍNH
    // ══════════════════════════════════════════════════════════════

    public class TongHopDoLuongCommands
    {
        /// <summary>
        /// Lệnh gộp: Mở form chọn loại đo + chế độ xuất
        /// Thay thế 8 lệnh AT_TongDoDai_*/AT_TongDienTich_*
        /// </summary>
        [CommandMethod("AT_TongHop")]
        public static void AT_TongHop()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            // Hiện form
            using var form = new TongHopDoLuongForm();
            AcadApp.ShowModalDialog(form);
            if (!form.IsConfirmed) return;

            bool isLength = form.IsLength;
            string mode = form.Mode;
            string propName = isLength ? "Length" : "Area";
            string unit = isLength ? "m" : "m2";
            string prompt = isLength
                ? "Chọn các đối tượng (Polyline) cần tính tổng độ dài: \n"
                : "Chọn các đối tượng (Polyline/Hatch) cần tính tổng diện tích: \n";

            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                // 1. Chọn đối tượng
                ObjectIdCollection polyLineColl = UserInput.GSelectionSet(prompt);
                List<string> listFields = new();

                // 2. Tạo field data
                string str1 = "%<\\AcObjProp Object(%<\\_ObjId ";
                string str2 = $">%).{propName} \\f \"%lu2\">%";

                foreach (ObjectId id in polyLineColl)
                {
                    string strId = id.OldIdPtr.ToString();
                    listFields.Add(str1 + strId + str2);
                }

                // 3. Tính tổng
                string str3 = "%<\\AcExpr (";
                string str4 = ") \\f \" %lu2\">%";

                switch (mode)
                {
                    case "Full":
                    {
                        // Tạo text MỚI (kèm đơn vị)
                        string formatTong = "0";
                        foreach (string item in listFields)
                            formatTong += "+" + item;

                        string formatLite = formatTong[2..]; // bỏ "0+"
                        formatTong = formatLite + "=" + str3 + formatTong + str4 + unit;

                        Point3d point = UserInput.GPoint("Chọn vị trí đặt text: \n");
                        UtilitiesCAD.CCreateLayer("TinhTong");
                        if (isLength)
                            UtilitiesCAD.CCreateText2(point, 2, formatTong, "TinhTong", "Standard");
                        else
                            UtilitiesCAD.CCreateText(point, 2, formatTong, "TinhTong", "Standard");

                        System.Windows.Forms.Clipboard.SetText(formatTong);
                        ed.WriteMessage($"\n✅ Đã tạo text tổng {(isLength ? "độ dài" : "diện tích")}.");
                        break;
                    }

                    case "Replace":
                    {
                        // Ghi vào text sẵn (kèm đơn vị)
                        string formatTong = "0";
                        foreach (string item in listFields)
                            formatTong += "+" + item;
                        formatTong = str3 + formatTong + str4 + unit;

                        ObjectId textId = UserInput.GDbText("Chọn text cần nhập nội dung: \n");
                        DBText? text = tr.GetObject(textId, OpenMode.ForWrite) as DBText;
                        if (text != null)
                        {
                            text.TextString = formatTong;
                            A.Ed.Command("_UPDATEFIELD", textId);
                        }
                        ed.WriteMessage($"\n✅ Đã ghi tổng vào text (kèm đơn vị {unit}).");
                        break;
                    }

                    case "Replace2":
                    {
                        // Ghi vào text sẵn (KHÔNG đơn vị)
                        string formatTong = "0";
                        foreach (string item in listFields)
                            formatTong += "+" + item;
                        formatTong = str3 + formatTong + str4;

                        ObjectId textId = UserInput.GDbText("Chọn text cần nhập nội dung: \n");
                        DBText? text = tr.GetObject(textId, OpenMode.ForWrite) as DBText;
                        if (text != null)
                        {
                            text.TextString = formatTong;
                            A.Ed.Command("_UPDATEFIELD", textId);
                        }
                        ed.WriteMessage("\n✅ Đã ghi tổng vào text (không đơn vị).");
                        break;
                    }

                    case "CongThem":
                    {
                        // Cộng thêm vào text cũ
                        ObjectId textId = UserInput.GDbText("Chọn text cần cộng thêm: \n");
                        DBText? text = tr.GetObject(textId, OpenMode.ForWrite) as DBText;
                        if (text != null)
                        {
                            string formatTong = text.TextString;
                            foreach (string item in listFields)
                                formatTong += "+" + item;
                            formatTong = str3 + formatTong + str4;

                            text.TextString = formatTong;
                            A.Ed.Command("_UPDATEFIELD", textId);
                        }
                        ed.WriteMessage("\n✅ Đã cộng thêm vào text.");
                        break;
                    }
                }

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n⊘ Lỗi: {ex.Message}");
            }
        }
    }
}
