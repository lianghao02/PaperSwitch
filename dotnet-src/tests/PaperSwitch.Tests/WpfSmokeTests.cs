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
            }

            var rtb = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(dv);
            // 直接渲染配置後的內容，避免 VisualBrush 置中裁切最小尺寸快照。
            rtb.Render(content);

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
        public void RenameDialog_RenderSnapshot_ShouldNotClipControls()
        {
            Exception? failure = null;
            var thread = new Thread(() =>
            {
                try
                {
                    EnsureApplication();
                    var dlg = new RenameDialog("1-5,8", 1, "依畫布頁序選取", "輸入畫布位置，例如 1-5,8,12-20；不是原始文件頁碼。", "選取頁面");
                    var content = (UIElement)dlg.Content;
                    content.Measure(new Size(440, 480));
                    int measuredHeight = (int)Math.Ceiling(content.DesiredSize.Height);
                    content.Arrange(new Rect(0, 0, 440, measuredHeight));
                    content.UpdateLayout();

                    var nameBox = (FrameworkElement)dlg.FindName("NameTextBox");
                    var applyButton = (FrameworkElement)dlg.FindName("ApplyButton");
                    Assert.True(nameBox.ActualHeight >= 36, $"輸入框高度不足 ({nameBox.ActualHeight})");
                    Assert.True(applyButton.ActualHeight >= 40, $"按鈕高度不足 ({applyButton.ActualHeight})");
                    double buttonBottom = applyButton.TranslatePoint(new Point(0, applyButton.ActualHeight), content).Y;
                    Assert.True(buttonBottom <= measuredHeight + 1, $"按鈕超出視窗底部 ({buttonBottom} > {measuredHeight})");

                    string outDir = Environment.GetEnvironmentVariable("PAPERSWITCH_SNAPSHOT_DIR") is { Length: > 0 } customDir
                        ? customDir
                        : Path.Combine(Path.GetTempPath(), "PaperSwitch-ui-snapshots");
                    Directory.CreateDirectory(outDir);
                    var dv = new DrawingVisual();
                    using (var dc = dv.RenderOpen())
                        dc.DrawRectangle(dlg.Background ?? Brushes.White, null, new Rect(0, 0, 440, measuredHeight));
                    var rtb = new RenderTargetBitmap(440, measuredHeight, 96, 96, PixelFormats.Pbgra32);
                    rtb.Render(dv);
                    rtb.Render(content);
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(rtb));
                    using var fs = File.Create(Path.Combine(outDir, "rendered_range_dialog_fixed.png"));
                    encoder.Save(fs);
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

                        vm.Options.AddPageNumbers = true;
                        RenderWindowToPng(window, Path.Combine(outDir, "rendered_page_numbers.png"), 1264, 860);
                        RenderWindowToPng(window, Path.Combine(outDir, "rendered_page_numbers_min.png"), 960, 720);

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

        [Fact]
        public void Arranger_DragScrollAndSingleRowExport_ShouldWorkAtSmallAndLargeSizes()
        {
            Exception? failure = null;
            var thread = new Thread(() =>
            {
                MainWindow? window = null;
                try
                {
                    EnsureApplication();
                    window = new MainWindow();
                    var vm = (PaperSwitch.ViewModels.MainViewModel)window.DataContext;
                    for (int n = 1; n <= 154; n++)
                        vm.Pages.Add(new PaperSwitch.Models.PaperItem
                        {
                            SourceFileName = "拖曳驗收.pdf", DisplayPageNumber = n,
                            TotalPagesInSource = 154, IsLoadingThumbnail = false
                        });

                    string outDir = Environment.GetEnvironmentVariable("PAPERSWITCH_SNAPSHOT_DIR")
                        ?? Path.Combine(Path.GetTempPath(), "PaperSwitch-ui-snapshots");
                    Directory.CreateDirectory(outDir);
                    foreach (var size in new[] { (960, 720), (1280, 900), (1920, 1080) })
                    {
                        RenderWindowToPng(window, Path.Combine(outDir, $"export_single_row_{size.Item1}.png"), size.Item1, size.Item2);
                        var names = new[] { "OpenOutputButton", "ExportAllButton", "ExportSplitButton", "ExportSelectedButton", "ExportImagesButton" };
                        var buttons = Array.ConvertAll(names, name => (System.Windows.Controls.Button)window.FindName(name));
                        var content = (UIElement)window.Content;
                        double rowTop = buttons[0].TranslatePoint(new Point(), content).Y;
                        double previousRight = 0;
                        foreach (var button in buttons)
                        {
                            Point origin = button.TranslatePoint(new Point(), content);
                            Assert.Equal(rowTop, origin.Y, 1);
                            Assert.True(origin.X >= previousRight);
                            Assert.True(origin.X + button.ActualWidth <= size.Item1);
                            Assert.True(button.ActualHeight >= 40);
                            previousRight = origin.X + button.ActualWidth;
                        }

                        var sidebarScroll = (System.Windows.Controls.ScrollViewer)window.FindName("SidebarScrollViewer");
                        var hintCard = (FrameworkElement)window.FindName("SidebarHintCard");
                        double hintBottom = hintCard.TranslatePoint(new Point(0, hintCard.ActualHeight), sidebarScroll).Y;
                        Assert.True(hintBottom <= sidebarScroll.ActualHeight + 2, $"視窗 {size.Item1}x{size.Item2} 下左側工坊小提示不可被截斷 ({hintBottom} > {sidebarScroll.ActualHeight})");
                    }

                    var scroll = (System.Windows.Controls.ScrollViewer)window.FindName("MainScrollViewer");
                    Assert.True(scroll.ScrollableHeight > scroll.ViewportHeight);
                    var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance;
                    void Invoke(string name, params object[] args) => typeof(MainWindow).GetMethod(name, flags)!.Invoke(window, args);
                    var timer = (System.Windows.Threading.DispatcherTimer)typeof(MainWindow).GetField("_dragScrollTimer", flags)!.GetValue(window)!;
                    vm.DraggedPageCount = 1;
                    Invoke("UpdateDragAutoScroll", new Point(scroll.ActualWidth / 2, scroll.ActualHeight - 10));
                    Assert.True(timer.IsEnabled);
                    // 指標停在邊緣時仍持續捲動，不依賴新的 MouseMove。
                    for (int n = 0; n < 4; n++) Invoke("DragScrollTimer_Tick", timer, EventArgs.Empty);
                    double downOffset = scroll.VerticalOffset;
                    Assert.True(downOffset > 0);
                    Invoke("UpdateDragAutoScroll", new Point(scroll.ActualWidth / 2, 10));
                    Invoke("DragScrollTimer_Tick", timer, EventArgs.Empty);
                    Assert.True(scroll.VerticalOffset < downOffset);

                    Invoke("UpdateDragAutoScroll", new Point(scroll.ActualWidth / 2, scroll.ActualHeight / 2));
                    Assert.False(timer.IsEnabled);
                    Invoke("UpdateDragAutoScroll", new Point(-1, scroll.ActualHeight - 10));
                    Assert.False(timer.IsEnabled);
                    Invoke("UpdateDragAutoScroll", new Point(scroll.ActualWidth / 2, scroll.ActualHeight - 10));
                    vm.DraggedPageCount = 0;
                    double stoppedOffset = scroll.VerticalOffset;
                    Invoke("DragScrollTimer_Tick", timer, EventArgs.Empty);
                    Assert.False(timer.IsEnabled);
                    Assert.Equal(stoppedOffset, scroll.VerticalOffset);
                    scroll.ScrollToBottom();
                    scroll.UpdateLayout();
                    Invoke("UpdateDragAutoScroll", new Point(scroll.ActualWidth / 2, scroll.ActualHeight - 10));
                    Assert.False(timer.IsEnabled);
                }
                catch (Exception ex) { failure = ex; }
                finally { window?.Close(); }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
            Assert.True(thread.Join(TimeSpan.FromSeconds(30)), "拖曳及單列排版驗證逾時。");
            if (failure != null) throw failure;
        }

        [Fact]
        public void LightboxWindow_ContextMenu_ExportsCurrentPageOnlyOnSta()
        {
            var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            Exception? failure = null;
            try
            {
                var pdf = Path.Combine(dir, "lightbox-test.pdf");
                Assert.True(PaperSwitch.Services.PdfService.Instance.CreateBlankA4Pdf(pdf));

                var thread = new Thread(() =>
                {
                    LightboxWindow? window = null;
                    try
                    {
                        EnsureApplication();
                        var hostVm = new PaperSwitch.ViewModels.MainViewModel(new PaperSwitch.Services.ExportPreferencesStore(Path.Combine(dir, "prefs.json")));
                        var page = PaperSwitch.Services.PdfImportService.ReadPages(pdf, "來源")[0];
                        hostVm.Pages.Add(page);

                        var callbacks = new PaperSwitch.ViewModels.LightboxHostCallbacks(
                            hostVm.RotatePage,
                            hostVm.ExportCustomPagesAsync);
                        window = new LightboxWindow(page, hostVm.Pages, callbacks);

                        string outDir = Environment.GetEnvironmentVariable("PAPERSWITCH_SNAPSHOT_DIR")
                            ?? Path.Combine(Path.GetTempPath(), "PaperSwitch-ui-snapshots");
                        Directory.CreateDirectory(outDir);
                        RenderWindowToPng(window, Path.Combine(outDir, "lightbox_context_menu.png"), 1040, 780);

                        var scroll = (System.Windows.Controls.ScrollViewer)window.FindName("PreviewScrollViewer");
                        Assert.NotNull(scroll.ContextMenu);
                        Assert.Equal(2, scroll.ContextMenu.Items.Count);

                        var exportCurrentPdf = (System.Windows.Controls.MenuItem)window.FindName("MenuExportCurrentPdf");
                        var exportCurrentPng = (System.Windows.Controls.MenuItem)window.FindName("MenuExportCurrentPng");
                        Assert.True(exportCurrentPdf.IsEnabled);
                        Assert.True(exportCurrentPng.IsEnabled);
                    }
                    catch (Exception ex) { failure = ex; }
                    finally { window?.Close(); }
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.IsBackground = true;
                thread.Start();
                Assert.True(thread.Join(TimeSpan.FromSeconds(15)), "燈箱右鍵選單驗證逾時。");
                if (failure != null) throw failure;
            }
            finally { Directory.Delete(dir, true); }
        }

        [Fact]
        public void MainWindow_WorkspaceContextMenu_AdaptsToCardAndMultiSelectionOnSta()
        {
            Exception? failure = null;
            var thread = new Thread(() =>
            {
                MainWindow? window = null;
                try
                {
                    EnsureApplication();
                    window = new MainWindow();
                    var vm = (PaperSwitch.ViewModels.MainViewModel)window.DataContext;
                    var p1 = new PaperSwitch.Models.PaperItem { SourceFileName = "文件A.pdf", DisplayPageNumber = 1, TotalPagesInSource = 2, IsLoadingThumbnail = false };
                    var p2 = new PaperSwitch.Models.PaperItem { SourceFileName = "文件A.pdf", DisplayPageNumber = 2, TotalPagesInSource = 2, IsLoadingThumbnail = false };
                    vm.Pages.Add(p1);
                    vm.Pages.Add(p2);

                    string outDir = Environment.GetEnvironmentVariable("PAPERSWITCH_SNAPSHOT_DIR")
                        ?? Path.Combine(Path.GetTempPath(), "PaperSwitch-ui-snapshots");
                    Directory.CreateDirectory(outDir);
                    RenderWindowToPng(window, Path.Combine(outDir, "workspace_context_menu.png"), 1120, 760);
                    window.Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Background);

                    var itemsControl = (System.Windows.Controls.ItemsControl)window.FindName("CardsItemsControl");
                    var container1 = (FrameworkElement)itemsControl.ItemContainerGenerator.ContainerFromIndex(0);
                    var container2 = (FrameworkElement)itemsControl.ItemContainerGenerator.ContainerFromIndex(1);
                    Assert.NotNull(container1);
                    Assert.NotNull(container2);

                    var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance;
                    static System.Windows.Controls.Border FindCardBorder(DependencyObject root)
                    {
                        for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(root); i++)
                        {
                            var child = System.Windows.Media.VisualTreeHelper.GetChild(root, i);
                            if (child is System.Windows.Controls.Border b && b.Name == "CardBorder") return b;
                            var nested = FindCardBorder(child);
                            if (nested != null) return nested;
                        }
                        return null!;
                    }

                    var card1 = FindCardBorder(container1);
                    var card2 = FindCardBorder(container2);
                    Assert.NotNull(card1.ContextMenu);
                    Assert.NotNull(card2.ContextMenu);

                    var canvasPdfItem = (System.Windows.Controls.MenuItem)window.FindName("CanvasMenuExportSelectedPdf");
                    var canvasSplitItem = (System.Windows.Controls.MenuItem)window.FindName("CanvasMenuExportSplitPdf");

                    p1.IsSelected = true;
                    vm.SelectedPage = p1;
                    vm.NotifySelectionChanged();
                    window.Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Background);
                    Assert.Equal("匯出這頁 (PDF)", canvasPdfItem.Header?.ToString());
                    Assert.Equal(Visibility.Collapsed, canvasSplitItem.Visibility);

                    p2.IsSelected = true;
                    vm.NotifySelectionChanged();
                    window.Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Background);
                    Assert.Equal("合併匯出已選 2 頁 (PDF)", canvasPdfItem.Header?.ToString());
                    Assert.Equal(Visibility.Visible, canvasSplitItem.Visibility);

                    var rightDownMethod = typeof(MainWindow).GetMethod("Card_PreviewMouseRightButtonDown", flags)!;
                    rightDownMethod.Invoke(window, new object?[] { card1, null });
                    Assert.True(p1.IsSelected);
                    Assert.True(p2.IsSelected);

                    p1.IsSelected = false;
                    p2.IsSelected = true;
                    vm.NotifySelectionChanged();
                    rightDownMethod.Invoke(window, new object?[] { card1, null });
                    Assert.True(p1.IsSelected);
                    Assert.False(p2.IsSelected);
                }
                catch (Exception ex) { failure = ex; }
                finally { window?.Close(); }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
            Assert.True(thread.Join(TimeSpan.FromSeconds(15)), "工作桌右鍵選單驗證逾時。");
            if (failure != null) throw failure;
        }
    }
}
