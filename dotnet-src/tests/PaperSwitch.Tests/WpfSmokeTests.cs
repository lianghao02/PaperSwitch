using System;
using System.Threading;
using System.Windows;
using Xunit;

namespace PaperSwitch.Tests
{
    public class WpfSmokeTests
    {
        [Fact]
        public void MainWindow_Xaml_ShouldLoadOnStaThread()
        {
            Exception? startupException = null;
            string? windowTitle = null;

            var thread = new Thread(() =>
            {
                App? application = null;
                MainWindow? window = null;
                try
                {
                    application = new App();
                    application.InitializeComponent();
                    window = new MainWindow();
                    windowTitle = window.Title;
                }
                catch (Exception ex)
                {
                    startupException = ex;
                }
                finally
                {
                    window?.Close();
                    application?.Shutdown();
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();

            Assert.True(thread.Join(TimeSpan.FromSeconds(15)), "WPF 主視窗初始化逾時。");
            Assert.Null(startupException);
            Assert.Contains("PaperSwitch", windowTitle);
        }
    }
}
