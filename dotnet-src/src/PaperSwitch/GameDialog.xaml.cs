using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace PaperSwitch
{
    /// <summary>
    /// 文具工坊風格的訊息對話框，取代系統 MessageBox。
    /// Show() 的參數與 MessageBox.Show 相同，可直接替換。
    /// </summary>
    public partial class GameDialog : Window
    {
        private MessageBoxResult _result = MessageBoxResult.None;
        private MessageBoxResult _cancelResult = MessageBoxResult.None;

        private GameDialog(string message, string title, MessageBoxImage image)
        {
            InitializeComponent();
            TitleText.Text = title;
            MessageText.Text = message;
            IconText.Text = image switch
            {
                MessageBoxImage.Error => "⚠️",
                MessageBoxImage.Warning => "⚠️",
                MessageBoxImage.Question => "❓",
                _ => "💬"
            };
            Window? owner = null;
            try { owner = Application.Current?.MainWindow; }
            catch (InvalidOperationException) { /* 非 UI 執行緒（例如測試）：改為置中螢幕 */ }
            if (owner != null && owner.IsVisible && !ReferenceEquals(owner, this)) Owner = owner;
            else WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        private void AddButton(string text, MessageBoxResult result, bool primary, bool isDefault, bool isCancel)
        {
            var button = new Button
            {
                Content = text,
                MinWidth = 92,
                Height = 34,
                Margin = new Thickness(8, 0, 0, 0),
                IsDefault = isDefault,
                IsCancel = isCancel,
                Style = (Style)FindResource(primary ? "BtnPrimaryCraft" : "BtnOutlineCraft")
            };
            button.Click += (_, _) => { _result = result; DialogResult = true; };
            if (isCancel) _cancelResult = result;
            ButtonPanel.Children.Add(button);
        }

        public static MessageBoxResult Show(string message, string title = "PaperSwitch",
            MessageBoxButton buttons = MessageBoxButton.OK, MessageBoxImage image = MessageBoxImage.Information)
        {
            var dlg = new GameDialog(message, title, image);
            switch (buttons)
            {
                case MessageBoxButton.YesNo:
                    dlg.AddButton("否", MessageBoxResult.No, false, false, true);
                    dlg.AddButton("是", MessageBoxResult.Yes, true, true, false);
                    break;
                case MessageBoxButton.OKCancel:
                    dlg.AddButton("取消", MessageBoxResult.Cancel, false, false, true);
                    dlg.AddButton("確定", MessageBoxResult.OK, true, true, false);
                    break;
                default:
                    dlg.AddButton("確定", MessageBoxResult.OK, true, true, true);
                    break;
            }
            dlg.ShowDialog();
            return dlg._result == MessageBoxResult.None ? dlg._cancelResult : dlg._result;
        }

        /// <summary>匯出完成：顯示成果摘要與輸出位置，可直接開啟資料夾。</summary>
        public static void ShowExportDone(string summary, string folder, string? note = null)
        {
            string message = summary + "\n\n存放位置：\n" + folder + (string.IsNullOrWhiteSpace(note) ? "" : "\n\n" + note);
            var dlg = new GameDialog(message, "匯出完成", MessageBoxImage.Information);
            dlg.IconText.Text = "✅";
            dlg.AddButton("完成", MessageBoxResult.OK, false, false, true);
            dlg.AddButton("開啟輸出資料夾", MessageBoxResult.Yes, true, true, false);
            dlg.ShowDialog();
            if (dlg._result == MessageBoxResult.Yes && Directory.Exists(folder))
            {
                Process.Start(new ProcessStartInfo { FileName = folder, UseShellExecute = true });
            }
        }
    }
}
