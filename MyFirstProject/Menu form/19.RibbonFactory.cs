using System;
using System.Collections.Generic;
using System.Windows.Media.Imaging;
using Autodesk.Windows;

namespace Civil3DCsharp.Helpers
{
    public static class RibbonFactory
    {
        /// <summary>
        /// Tạo một Ribbon Panel với các nút bấm chuyên nghiệp
        /// </summary>
        public static RibbonPanel CreatePanel(string title)
        {
            var panelSource = new RibbonPanelSource { Title = title };
            return new RibbonPanel { Source = panelSource };
        }

        /// <summary>
        /// Tạo một nút bấm chuẩn với Tooltip chi tiết
        /// </summary>
        public static RibbonButton CreateButton(string command, string label, string description, string? tooltipImage = null)
        {
            var button = new RibbonButton
            {
                Text = label,
                ShowText = true,
                ShowImage = true,
                CommandParameter = command + " ",
                CommandHandler = new RibbonCommandHandler(),
                Orientation = System.Windows.Controls.Orientation.Vertical,
                Size = RibbonItemSize.Large
            };

            // Thiết lập Rich Tooltip
            button.ToolTip = new RibbonToolTip
            {
                Title = label,
                Content = description,
                Command = command
            };

            return button;
        }

        /// <summary>
        /// Tạo một SplitButton (nút gộp) để nhóm các lệnh liên quan
        /// </summary>
        public static RibbonSplitButton CreateSplitButton(string label, List<RibbonButton> subButtons)
        {
            var splitButton = new RibbonSplitButton
            {
                Text = label,
                ShowText = true,
                ShowImage = true,
                Orientation = System.Windows.Controls.Orientation.Vertical,
                Size = RibbonItemSize.Large,
                ListStyle = RibbonSplitButtonListStyle.List,
                IsSplit = true
            };

            foreach (var btn in subButtons)
            {
                splitButton.Items.Add(btn);
            }

            return splitButton;
        }

        /// <summary>
        /// Tạo icon emoji đơn giản cho Ribbon (Dùng làm placeholder nếu không có file ảnh)
        /// </summary>
        public static BitmapImage? CreateEmojiIcon(string emoji, int size = 32)
        {
            // Trong thực tế, AutoCAD ribbon cần BitmapSource. 
            // Nếu không có file png/ico, ta có thể dùng hệ thống icon chuẩn hoặc emoji font.
            // Placeholder: Trả về null để AutoCAD dùng icon mặc định
            return null;
        }
    }

    /// <summary>
    /// Xử lý sự kiện click chuột cho các nút trên Ribbon
    /// </summary>
    public class RibbonCommandHandler : System.Windows.Input.ICommand
    {
        public bool CanExecute(object? parameter) => true;
#pragma warning disable CS0067 // Sự kiện bắt buộc bởi ICommand nhưng không cần raise thủ công
        public event EventHandler? CanExecuteChanged;
#pragma warning restore CS0067

        public void Execute(object? parameter)
        {
            if (parameter is RibbonButton button)
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.SendStringToExecute(button.CommandParameter.ToString(), true, false, true);
                }
            }
        }
    }
}
