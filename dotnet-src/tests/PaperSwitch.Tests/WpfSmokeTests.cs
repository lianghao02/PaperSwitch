using System;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Xunit;

namespace PaperSwitch.Tests
{
    public class WpfSmokeTests
    {
        private static Application EnsureApplication()
        {
            if (Application.Current == null)
            {
                var app = new App();
                app.InitializeComponent();
                return app;
            }
            return Application.Current;
        }

        [Fact]
        public void MainWindow_Xaml_ShouldLoadOnStaThread()
        {
            Exception? startupException = null;
            string? windowTitle = null;

            var thread = new Thread(() =>
            {
                MainWindow? window = null;
                try
                {
                    EnsureApplication();
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
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();

            Assert.True(thread.Join(TimeSpan.FromSeconds(15)), "WPF 主視窗初始化逾時。");
            Assert.Null(startupException);
            Assert.Contains("PaperSwitch", windowTitle);
        }

        private static void RenderWindowToPng(Window window, string filePath, int width = 1200, int height = 750)
        {
            window.Width = width;
            window.Height = height;
            var content = (UIElement)window.Content;
            content.Measure(new Size(width, height));
            content.Arrange(new Rect(0, 0, width, height));
            content.UpdateLayout();

            var dv = new DrawingVisual();
            using (var dc = dv.RenderOpen())
            {
                dc.DrawRectangle(window.Background ?? Brushes.White, null, new Rect(0, 0, width, height));
                var vb = new VisualBrush(content) { Stretch = Stretch.None };
                dc.DrawRectangle(vb, null, new Rect(0, 0, width, height));
            }

            var rtb = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(dv);

            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(rtb));
            using var fs = File.Create(filePath);
            encoder.Save(fs);
        }

        [Fact]
        public void GameDialog_RenderSnapshot_ShouldSaveImage()
        {
            Exception? failure = null;
            var thread = new Thread(() =>
            {
                try
                {
                    EnsureApplication();
                    var ctor = typeof(GameDialog).GetConstructor(
                        System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                        null, new[] { typeof(string), typeof(string), typeof(MessageBoxImage) }, null)!;
                    var dlg = (GameDialog)ctor.Invoke(new object[]
                    {
                        "已合併成 1 份 PDF：\n開會通知.pdf\n\n存放位置：\nC:\\Users\\測試\\PaperSwitch\\Converted",
                        "匯出完成", MessageBoxImage.Information
                    });
                    var add = typeof(GameDialog).GetMethod("AddButton", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
                    add.Invoke(dlg, new object[] { "完成", MessageBoxResult.OK, false, false, true });
                    add.Invoke(dlg, new object[] { "開啟輸出資料夾", MessageBoxResult.Yes, true, true, false });
                    string outDir = Environment.GetEnvironmentVariable("PAPERSWITCH_SNAPSHOT_DIR") is { Length: > 0 } customDir
                        ? customDir
                        : Path.Combine(Path.GetTempPath(), "PaperSwitch-ui-snapshots");
                    Directory.CreateDirectory(outDir);
                    RenderWindowToPng(dlg, Path.Combine(outDir, "rendered_dialog.png"), 460, 300);
                }
                catch (Exception ex) { failure = ex; }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (failure != null) throw failure;
        }

        [Fact]
        public void MainWindow_RenderSnapshot_ShouldSaveImages()
        {
            var thread = new Thread(() =>
            {
                MainWindow? window = null;
                try
                {
                    EnsureApplication();
                    window = new MainWindow();

                    var vm = window.DataContext as PaperSwitch.ViewModels.MainViewModel;
                    // 快照輸出至暫存資料夾，可用 PAPERSWITCH_SNAPSHOT_DIR 指定；不寫死個人電腦路徑。
                    string outDir = Environment.GetEnvironmentVariable("PAPERSWITCH_SNAPSHOT_DIR") is { Length: > 0 } customDir
                        ? customDir
                        : Path.Combine(Path.GetTempPath(), "PaperSwitch-ui-snapshots");
                    Directory.CreateDirectory(outDir);

                    // 1. 空狀態渲染
                    RenderWindowToPng(window, Path.Combine(outDir, "rendered_empty.png"));

                    // 2. 有資料狀態渲染
                    if (vm != null)
                    {
                        var item1 = new PaperSwitch.Models.PaperItem
                        {
                            SourceFileName = "開會通知.pdf",
                            DisplayPageNumber = 1,
                            TotalPagesInSource = 3,
                            IsLoadingThumbnail = false,
                            IsSelected = true
                        };
                        var item2 = new PaperSwitch.Models.PaperItem
                        {
                            SourceFileName = "開會通知.pdf",
                            DisplayPageNumber = 2,
                            TotalPagesInSource = 3,
                            IsLoadingThumbnail = false,
                            Rotation = 90
                        };
                        var item3 = new PaperSwitch.Models.PaperItem
                        {
                            SourceFileName = "簡報投影片.pdf",
                            DisplayPageNumber = 1,
                            TotalPagesInSource = 1,
                            IsLoadingThumbnail = false
                        };
                        vm.Pages.Add(item1);
                        vm.Pages.Add(item2);
                        vm.Pages.Add(item3);

                        RenderWindowToPng(window, Path.Combine(outDir, "rendered_populated.png"));

                        // 3. 轉檔任務：被公司加密程序鎖住的失敗卡片
                        var locked = new PaperSwitch.Models.ConversionTaskItem(@"C:\測試\551.docx");
                        locked.MarkFailed("轉出的 PDF 被公司加密程序鎖住，需要先解密。原始檔案沒有被修改。", null, @"C:\測試\551.pdf", needsDecrypt: true);
                        vm.ConversionTasks.Add(locked);
                        vm.IsConversionQueueExpanded = true;
                        RenderWindowToPng(window, Path.Combine(outDir, "rendered_locked_task.png"));
                    }
                }
                finally
                {
                    window?.Close();
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
            Assert.True(thread.Join(TimeSpan.FromSeconds(15)), "WPF 渲染快照逾時。");
        }
    }
}
