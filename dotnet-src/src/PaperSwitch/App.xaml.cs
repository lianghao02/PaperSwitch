using System;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Threading;
using PaperSwitch.Services;

namespace PaperSwitch
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // 全域例外捕獲與記錄
            AppDomain.CurrentDomain.UnhandledException += (s, ev) =>
            {
                LogUnhandledException(ev.ExceptionObject as Exception, "AppDomain Unhandled");
            };

            DispatcherUnhandledException += (s, ev) =>
            {
                LogUnhandledException(ev.Exception, "Dispatcher Unhandled");
                ev.Handled = true;
                MessageBox.Show($"發生未預期的錯誤: {ev.Exception.Message}", "系統提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
            };
        }

        private static void LogUnhandledException(Exception? ex, string source)
        {
            if (ex == null) return;
            try
            {
                string content = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{source}]\n{ex}\n\n";
                File.AppendAllText(AppPaths.CrashLogPath, content, Encoding.UTF8);
            }
            catch { }
        }
    }
}
