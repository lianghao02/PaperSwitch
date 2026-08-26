using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using PaperSwitch.Models;
using PaperSwitch.Services;
using PaperSwitch.ViewModels;

namespace PaperSwitch
{
    public partial class MainWindow : Window
    {
        private Point _dragStartPoint;
        private PaperItem? _draggedItem;
        private int _lastClickedIndex = -1;
        private bool _isBoxSelecting;
        private Point _boxSelectionStart;
        private readonly HashSet<PaperItem> _initialBoxSelection = new();
        private PaperItem? _dropIndicatorItem;

        public MainWindow()
        {
            InitializeComponent();
        }

        #region 鍵盤快捷鍵極速導航

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (ViewModel.IsBusy) return;

            // Ctrl + A: 全選 / 取消全選切換
            if (e.Key == Key.A && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (ViewModel.SelectedPageCount == ViewModel.TotalPageCount)
                    ViewModel.DeselectAll();
                else
                    ViewModel.SelectAll();
                e.Handled = true;
                return;
            }

            // Ctrl + 方向鍵：左右移動一格，上下移動一個視覺列距。
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (e.Key == Key.Left)
                {
                    ViewModel.MoveSelectedLeft();
                    e.Handled = true;
                    return;
                }
                if (e.Key == Key.Right)
                {
                    ViewModel.MoveSelectedRight();
                    e.Handled = true;
                    return;
                }
                if (e.Key == Key.Up)
                {
                    ViewModel.MoveSelectedBy(-GetVisualColumnCount());
                    e.Handled = true;
                    return;
                }
                if (e.Key == Key.Down)
                {
                    ViewModel.MoveSelectedBy(GetVisualColumnCount());
                    e.Handled = true;
                    return;
                }
            }

            // Home / End: 移至開頭 / 移至結尾
            if (e.Key == Key.Home)
            {
                ViewModel.MoveSelectedToStart();
                e.Handled = true;
                return;
            }
            if (e.Key == Key.End)
            {
                ViewModel.MoveSelectedToEnd();
                e.Handled = true;
                return;
            }

            // R / Shift + R: 旋轉
            if (e.Key == Key.R)
            {
                if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
                    ViewModel.RotateSelected(-90);
                else
                    ViewModel.RotateSelected(90);
                e.Handled = true;
                return;
            }

            // Delete: 刪除選取紙張
            if (e.Key == Key.Delete || e.Key == Key.Back)
            {
                ViewModel.DeleteSelected();
                e.Handled = true;
                return;
            }
        }

        private int GetVisualColumnCount()
        {
            const double cardHorizontalMargin = 20.0;
            double availableWidth = MainScrollViewer.ViewportWidth;
            if (availableWidth <= 0) availableWidth = MainScrollViewer.ActualWidth;

            return Math.Max(1, (int)(availableWidth / (ViewModel.ZoomLevel + cardHorizontalMargin)));
        }

        #endregion

        #region 滑鼠滾輪縮放 (Ctrl + Wheel)

        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                e.Handled = true;
                double delta = (e.Delta > 0) ? 15.0 : -15.0;
                double newZoom = Math.Clamp(ViewModel.ZoomLevel + delta, 90.0, 480.0);
                ViewModel.ZoomLevel = newZoom;
            }
        }

        #endregion

        #region 檔案拖放 (Drag & Drop)

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private async void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[]?)e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    await ViewModel.AddFilesAsync(files);
                }
            }
        }

        #endregion

        #region 卡片選取、多選與拖曳換位

        private void Card_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is PaperItem item)
            {
                if (e.ClickCount == 2)
                {
                    // 雙擊開啟大圖燈箱 (Lightbox)
                    var lightbox = new LightboxWindow(item, ViewModel.Pages)
                    {
                        Owner = this
                    };
                    lightbox.ShowDialog();
                    e.Handled = true;
                    return;
                }

                _dragStartPoint = e.GetPosition(this);
                _draggedItem = item;

                int currentIndex = ViewModel.Pages.IndexOf(item);

                if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    // Ctrl + Click: 切換單一選取
                    item.IsSelected = !item.IsSelected;
                    _lastClickedIndex = currentIndex;
                }
                else if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift && _lastClickedIndex >= 0)
                {
                    // Shift + Click: 連續區間選取
                    int start = Math.Min(_lastClickedIndex, currentIndex);
                    int end = Math.Max(_lastClickedIndex, currentIndex);

                    for (int i = 0; i < ViewModel.Pages.Count; i++)
                    {
                        ViewModel.Pages[i].IsSelected = (i >= start && i <= end);
                    }
                }
                else
                {
                    // 普通點擊：若尚未選取則設為唯一選取；若已選取則保留群組供拖曳
                    if (!item.IsSelected)
                    {
                        foreach (var p in ViewModel.Pages)
                        {
                            p.IsSelected = (p == item);
                        }
                    }
                    _lastClickedIndex = currentIndex;
                }

                ViewModel.SelectedPage = item;
                ViewModel.NotifySelectionChanged();
            }
        }

        private void Card_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && _draggedItem != null)
            {
                Point currentPoint = e.GetPosition(this);
                Vector diff = _dragStartPoint - currentPoint;

                if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    // 啟動拖曳
                    var selected = ViewModel.Pages.Where(p => p.IsSelected).ToList();
                    if (!selected.Contains(_draggedItem))
                    {
                        _draggedItem.IsSelected = true;
                        ViewModel.NotifySelectionChanged();
                    }

                    var dragData = new DataObject("PaperItemChunk", _draggedItem);
                    var draggedPages = ViewModel.Pages.Where(page => page.IsSelected).ToList();
                    foreach (var page in draggedPages)
                    {
                        page.IsBeingDragged = true;
                    }

                    ViewModel.DraggedPageCount = draggedPages.Count;
                    try
                    {
                        DragDrop.DoDragDrop((DependencyObject)sender, dragData, DragDropEffects.Move);
                    }
                    finally
                    {
                        foreach (var page in draggedPages)
                        {
                            page.IsBeingDragged = false;
                        }
                        ViewModel.DraggedPageCount = 0;
                        _draggedItem = null;
                        ClearDropIndicator();
                    }
                }
            }
        }

        private void Card_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("PaperItemChunk") && sender is Border card)
            {
                bool insertAfter = e.GetPosition(card).X > card.ActualWidth / 2;
                ShowDropIndicator(card, insertAfter);
                e.Effects = DragDropEffects.Move;
                e.Handled = true;
            }
        }

        private void Card_DragLeave(object sender, DragEventArgs e)
        {
            ClearDropIndicator();
        }

        private void Card_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent("PaperItemChunk") || sender is not FrameworkElement element || element.DataContext is not PaperItem target)
            {
                return;
            }

            int targetIndex = ViewModel.Pages.IndexOf(target);
            if (targetIndex < 0) return;

            bool insertAfter = e.GetPosition(element).X > element.ActualWidth / 2;
            if (insertAfter)
            {
                targetIndex++;
            }

            ViewModel.RelocateSelectedChunk(targetIndex);
            ClearDropIndicator();
            e.Handled = true;
        }

        private void ShowDropIndicator(Border target, bool insertAfter)
        {
            if (target.DataContext is not PaperItem item)
            {
                return;
            }

            // 插入提示使用畫布覆蓋層，不改動卡片 Margin，避免目標因重排而來回抖動。
            _dropIndicatorItem = item;
            ShowFloatingInsertionSlot(target, insertAfter);
        }

        private void ClearDropIndicator()
        {
            if (_dropIndicatorItem != null)
            {
                _dropIndicatorItem.IsDropTarget = false;
                _dropIndicatorItem.IsDropAfter = false;
            }

            _dropIndicatorItem = null;
            DropInsertionSlot.Visibility = Visibility.Collapsed;
        }

        private void ShowFloatingInsertionSlot(Border target, bool insertAfter)
        {
            Rect targetBounds = target.TransformToVisual(SelectionCanvas)
                .TransformBounds(new Rect(0, 0, target.ActualWidth, target.ActualHeight));

            const double slotWidth = 36.0;
            double left = insertAfter
                ? targetBounds.Right - (slotWidth / 2)
                : targetBounds.Left - (slotWidth / 2);

            Canvas.SetLeft(DropInsertionSlot, left);
            Canvas.SetTop(DropInsertionSlot, targetBounds.Top);
            DropInsertionSlot.Height = Math.Max(64, targetBounds.Height);
            DropInsertionSlot.Visibility = Visibility.Visible;
        }

        private void MainScrollViewer_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (IsInsidePaperCard(e.OriginalSource as DependencyObject)) return;

            _isBoxSelecting = true;
            _boxSelectionStart = e.GetPosition(SelectionCanvas);
            _initialBoxSelection.Clear();

            foreach (var page in ViewModel.Pages)
            {
                if (page.IsSelected) _initialBoxSelection.Add(page);
                if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control)
                {
                    page.IsSelected = false;
                }
            }

            SelectionRectangle.Visibility = Visibility.Visible;
            Canvas.SetLeft(SelectionRectangle, _boxSelectionStart.X);
            Canvas.SetTop(SelectionRectangle, _boxSelectionStart.Y);
            SelectionRectangle.Width = 0;
            SelectionRectangle.Height = 0;
            MainScrollViewer.CaptureMouse();
            e.Handled = true;
        }

        private void MainScrollViewer_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (!_isBoxSelecting || e.LeftButton != MouseButtonState.Pressed) return;

            Point current = e.GetPosition(SelectionCanvas);
            var selectionBounds = new Rect(_boxSelectionStart, current);
            Canvas.SetLeft(SelectionRectangle, selectionBounds.Left);
            Canvas.SetTop(SelectionRectangle, selectionBounds.Top);
            SelectionRectangle.Width = selectionBounds.Width;
            SelectionRectangle.Height = selectionBounds.Height;

            foreach (var page in ViewModel.Pages)
            {
                var container = CardsItemsControl.ItemContainerGenerator.ContainerFromItem(page) as FrameworkElement;
                if (container == null) continue;

                Rect itemBounds = container.TransformToVisual(SelectionCanvas)
                    .TransformBounds(new Rect(0, 0, container.ActualWidth, container.ActualHeight));
                page.IsSelected = _initialBoxSelection.Contains(page) || selectionBounds.IntersectsWith(itemBounds);
            }

            ViewModel.NotifySelectionChanged();
        }

        private void MainScrollViewer_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isBoxSelecting) return;

            _isBoxSelecting = false;
            SelectionRectangle.Visibility = Visibility.Collapsed;
            MainScrollViewer.ReleaseMouseCapture();
            ViewModel.SelectedPage = ViewModel.Pages.FirstOrDefault(page => page.IsSelected);
            ViewModel.NotifySelectionChanged();
            e.Handled = true;
        }

        private static bool IsInsidePaperCard(DependencyObject? source)
        {
            while (source != null)
            {
                if (source is FrameworkElement element && element.DataContext is PaperItem)
                {
                    return true;
                }

                source = VisualTreeHelper.GetParent(source);
            }

            return false;
        }

        #endregion

        #region 按鈕動作事件

        private async void SelectFiles_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Multiselect = true,
                Title = "選取要加入排版工坊的文件或圖片",
                Filter = "所有支援格式 (*.pdf;*.docx;*.doc;*.xlsx;*.xls;*.pptx;*.ppt;*.png;*.jpg;*.webp)|*.pdf;*.docx;*.doc;*.xlsx;*.xls;*.pptx;*.ppt;*.png;*.jpg;*.jpeg;*.bmp;*.webp;*.gif|PDF 文件 (*.pdf)|*.pdf|Word 文件 (*.docx;*.doc)|*.docx;*.doc|Excel 試算表 (*.xlsx;*.xls)|*.xlsx;*.xls|PowerPoint 簡報 (*.pptx;*.ppt)|*.pptx;*.ppt|圖片檔案 (*.png;*.jpg;*.webp;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp;*.webp;*.gif|所有檔案 (*.*)|*.*"
            };

            if (dialog.ShowDialog() == true && dialog.FileNames.Length > 0)
            {
                await ViewModel.AddFilesAsync(dialog.FileNames);
            }
        }

        private void CardRotate_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is PaperItem item)
            {
                item.RotateClockwise();
            }
        }

        private void CardDelete_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is PaperItem item)
            {
                ViewModel.Pages.Remove(item);
                ViewModel.NotifySelectionChanged();
            }
        }

        private void BatchRotateCW_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.RotateSelected(90);
        }

        private void BatchRotateCCW_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.RotateSelected(-90);
        }

        private void RenameSelected_Click(object sender, RoutedEventArgs e)
        {
            var selectedPages = ViewModel.Pages.Where(page => page.IsSelected).ToList();
            if (selectedPages.Count == 0 && ViewModel.SelectedPage != null)
            {
                selectedPages.Add(ViewModel.SelectedPage);
            }

            if (selectedPages.Count == 0)
            {
                MessageBox.Show("請先選取要重新命名的紙張。", "重新命名", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string suggestedName = Path.GetFileNameWithoutExtension(selectedPages[0].SourceFileName);
            var dialog = new RenameDialog(suggestedName, selectedPages.Count) { Owner = this };
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            if (!ViewModel.TryRenameSelectedPages(dialog.ProposedName, out string message))
            {
                MessageBox.Show(message, "重新命名", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ClearProcessingFolders_Click(object sender, RoutedEventArgs e)
        {
            string converted = AppPaths.ConvertedDirectory;
            string temporary = AppPaths.TemporaryConversionDirectory;
            var confirmation = MessageBox.Show(
                "此動作會永久刪除 PaperSwitch 已產生的成品與轉檔暫存檔。\n\n" +
                $"成品：{converted}\n" +
                $"暫存：{temporary}\n\n" +
                "不會刪除原始文件，也不會清空目前畫布。是否繼續？",
                "清理成品與暫存資料夾",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirmation != MessageBoxResult.Yes)
            {
                return;
            }

            try
            {
                var result = StorageMaintenanceService.ClearProcessingDirectories();
                if (result.FailedPaths.Count == 0)
                {
                    ViewModel.StatusMessage = $"已清理 {result.DeletedFileCount} 個成品與暫存檔";
                    MessageBox.Show($"已清理 {result.DeletedFileCount} 個檔案。\n\n畫布中的紙張不會受影響；若其來源是已清理的暫存檔，請重新匯入後再導出。",
                        "清理完成", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    ViewModel.StatusMessage = $"已清理 {result.DeletedFileCount} 個檔案，另有 {result.FailedPaths.Count} 個項目未能刪除";
                    MessageBox.Show($"已清理 {result.DeletedFileCount} 個檔案，但有 {result.FailedPaths.Count} 個項目正在被使用或沒有權限刪除。\n\n請關閉使用中的 PDF 或 Office 後再試一次。",
                        "部分清理完成", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"清理資料夾時發生錯誤：{ex.Message}", "清理失敗", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SplitMode_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.Options.MergeIntoSinglePdf = false;
        }

        #endregion
    }
}
