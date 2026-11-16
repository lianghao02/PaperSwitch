using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
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
        private bool _dropIndicatorAfter;
        private bool _dragOccurred;
        private const double DragScrollEdge = 56;
        private const double DragScrollStep = 24;
        private readonly DispatcherTimer _dragScrollTimer = new(DispatcherPriority.Input) { Interval = TimeSpan.FromMilliseconds(40) };
        private Point _dragScrollPosition;
        private readonly DispatcherTimer _viewportSyncTimer = new(DispatcherPriority.Background) { Interval = TimeSpan.FromMilliseconds(90) };
        private CancellationTokenSource? _viewportThumbnailCts;

        public MainWindow()
        {
            InitializeComponent();
            CanvasContextMenu.DataContext = ViewModel;
            _dragScrollTimer.Tick += DragScrollTimer_Tick;
            _viewportSyncTimer.Tick += ViewportSyncTimer_Tick;
            Closed += (_, _) =>
            {
                StopDragAutoScroll();
                _viewportSyncTimer.Stop();
                _viewportThumbnailCts?.Cancel();
                _viewportThumbnailCts?.Dispose();
            };
            Closing += (_, e) => { if (ViewModel.IsBusy) { e.Cancel = true; ViewModel.CancelOperation(); } };
        }

        #region 鍵盤快捷鍵極速導航

        private void Window_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var source = e.OriginalSource as DependencyObject;
            if (ViewModel.IsBusy) return;
            bool isPlainLeftClick = (Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control;

            if (!isPlainLeftClick
                || !IsInsideElement(source, ArrangerCanvas)
                || IsInsidePaperCard(source)
                || IsInsideScrollBar(source))
            {
                return;
            }

            // 視窗預覽階段優先處理，保證中央畫布任何空白處都能取消選取。
            if (Keyboard.FocusedElement is TextBoxBase or PasswordBox)
            {
                Keyboard.ClearFocus();
            }

            ViewModel.DeselectAll();
            ClearDropIndicator();
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (ViewModel.IsBusy) return;

            // 當焦點位於文字輸入框或其子元素時，完全放行文字編輯操作，絕不攔截畫布快捷鍵。
            if (Keyboard.FocusedElement is TextBoxBase or PasswordBox || IsInsideTextBox(e.OriginalSource as DependencyObject))
            {
                return;
            }

            // Ctrl + Z / Ctrl + Y：復原 / 重做編排動作
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (e.Key == Key.Z)
                {
                    ViewModel.Undo();
                    e.Handled = true;
                    return;
                }
                if (e.Key == Key.Y)
                {
                    ViewModel.Redo();
                    e.Handled = true;
                    return;
                }
            }

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
            if (e.Key is Key.Left or Key.Right or Key.Up or Key.Down)
            {
                int columns = GetVisualColumnCount();
                int delta = e.Key switch
                {
                    Key.Left => -1,
                    Key.Right => 1,
                    Key.Up => -columns,
                    _ => columns
                };
                NavigateSelection(delta, (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift);
                e.Handled = true;
                return;
            }

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

            // Delete: 刪除選取紙張 (僅 Key.Delete)
            if (e.Key == Key.Delete)
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

        private void NavigateSelection(int delta, bool extendSelection)
        {
            if (ViewModel.Pages.Count == 0) return;

            int currentIndex = ViewModel.SelectedPage is null
                ? 0
                : ViewModel.Pages.IndexOf(ViewModel.SelectedPage);
            if (currentIndex < 0) currentIndex = 0;
            int targetIndex = Math.Clamp(currentIndex + delta, 0, ViewModel.Pages.Count - 1);

            if (extendSelection)
            {
                if (_lastClickedIndex < 0) _lastClickedIndex = currentIndex;
                int start = Math.Min(_lastClickedIndex, targetIndex);
                int end = Math.Max(_lastClickedIndex, targetIndex);
                for (int index = 0; index < ViewModel.Pages.Count; index++)
                {
                    ViewModel.Pages[index].IsSelected = index >= start && index <= end;
                }
            }
            else
            {
                foreach (var page in ViewModel.Pages) page.IsSelected = false;
                ViewModel.Pages[targetIndex].IsSelected = true;
                _lastClickedIndex = targetIndex;
            }

            ViewModel.SelectedPage = ViewModel.Pages[targetIndex];
            ViewModel.NotifySelectionChanged();
            CardsItemsControl.UpdateLayout();
            if (CardsItemsControl.ItemContainerGenerator.ContainerFromItem(ViewModel.SelectedPage) is FrameworkElement card)
            {
                card.BringIntoView();
            }
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
            if (ViewModel.IsBusy) return;
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
                if (Keyboard.FocusedElement is TextBoxBase or PasswordBox)
                {
                    Keyboard.ClearFocus();
                }

                if (e.ClickCount == 2)
                {
                    // 雙擊開啟大圖燈箱 (Lightbox)
                    OpenLightboxForPage(item);
                    e.Handled = true;
                    return;
                }

                _dragOccurred = false;
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

                    for (int j = 0; j < ViewModel.Pages.Count; j++)
                    {
                        ViewModel.Pages[j].IsSelected = (j >= start && j <= end);
                    }
                }
                else
                {
                    // 普通點擊：若尚未選取則設為唯一單選
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

        private void Card_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is PaperItem item)
            {
                bool hasModifier = (Keyboard.Modifiers & (ModifierKeys.Control | ModifierKeys.Shift)) != ModifierKeys.None;
                if (!_dragOccurred && !hasModifier)
                {
                    // 普通單點放開（無拖曳、無修飾鍵）：退出多選狀態，變為僅單選此卡片
                    foreach (var p in ViewModel.Pages)
                    {
                        p.IsSelected = (p == item);
                    }
                    ViewModel.SelectedPage = item;
                    _lastClickedIndex = ViewModel.Pages.IndexOf(item);
                    ViewModel.NotifySelectionChanged();
                }
            }
            _draggedItem = null;
            _dragOccurred = false;
        }

        private void OpenLightboxForPage(PaperItem item)
        {
            var callbacks = new LightboxHostCallbacks(
                ViewModel.RotatePage,
                ViewModel.ExportCustomPagesAsync);
            var lightbox = new LightboxWindow(item, ViewModel.Pages, callbacks)
            {
                Owner = this
            };
            lightbox.ShowDialog();
        }

        private void Card_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is PaperItem item)
            {
                if (Keyboard.FocusedElement is TextBoxBase or PasswordBox)
                {
                    Keyboard.ClearFocus();
                }

                _draggedItem = null;
                _dragOccurred = false;

                if (!item.IsSelected)
                {
                    foreach (var p in ViewModel.Pages)
                    {
                        p.IsSelected = (p == item);
                    }
                    _lastClickedIndex = ViewModel.Pages.IndexOf(item);
                    ViewModel.SelectedPage = item;
                    ViewModel.NotifySelectionChanged();
                }
                else
                {
                    ViewModel.SelectedPage = item;
                }
            }
        }

        private void Card_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            if (sender is FrameworkElement element && element.ContextMenu != null)
            {
                element.ContextMenu.DataContext = ViewModel;
            }
        }

        private void ContextOpenLightbox_Click(object sender, RoutedEventArgs e)
        {
            PaperItem? target = null;
            if (sender is MenuItem menuItem &&
                ItemsControl.ItemsControlFromItemContainer(menuItem) is ContextMenu contextMenu &&
                contextMenu.PlacementTarget is FrameworkElement targetElement &&
                targetElement.DataContext is PaperItem clickedItem)
            {
                target = clickedItem;
            }

            target ??= ViewModel.SelectedPage ?? ViewModel.Pages.FirstOrDefault(p => p.IsSelected);
            if (target != null)
            {
                OpenLightboxForPage(target);
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
                    _dragOccurred = true;
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
                        StopDragAutoScroll();
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
                if (ViewModel.IsBusy || card.DataContext is PaperItem { IsSelected: true })
                {
                    ClearDropIndicator();
                    e.Effects = DragDropEffects.None;
                    e.Handled = true;
                    return;
                }
                bool insertAfter = e.GetPosition(card).X > card.ActualWidth / 2;
                ShowDropIndicator(card, insertAfter);
                e.Effects = DragDropEffects.Move;
                e.Handled = true;
            }
        }

        private void Card_DragLeave(object sender, DragEventArgs e)
        {
            // 不立即清除指示，避免游標跨過兩張卡片之間的縫隙時插入提示線閃爍消失；
            // 由 MainScrollViewer_PreviewDragOver 與 PreviewDragLeave 統一接手判斷。
        }

        private void Card_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent("PaperItemChunk") || sender is not FrameworkElement element || element.DataContext is not PaperItem target)
            {
                return;
            }

            StopDragAutoScroll();
            ClearDropIndicator();
            e.Handled = true;
            // 放回正在搬移的頁面，代表不更動順序（也涵蓋不連續多選）。
            if (ViewModel.IsBusy || target.IsSelected)
            {
                e.Effects = DragDropEffects.None;
                return;
            }

            bool insertAfter = e.GetPosition(element).X > element.ActualWidth / 2;
            CommitDropAt(target, insertAfter);
        }

        private void CommitDropAt(PaperItem target, bool insertAfter)
        {
            int targetIndex = ViewModel.Pages.IndexOf(target);
            if (targetIndex < 0) return;

            if (insertAfter)
            {
                targetIndex++;
            }

            ViewModel.RelocateSelectedChunk(targetIndex);
        }

        private void MainScrollViewer_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (ViewModel.IsBusy || !e.Data.GetDataPresent("PaperItemChunk"))
            {
                StopDragAutoScroll();
                return;
            }

            Point position = e.GetPosition(MainScrollViewer);
            UpdateDragAutoScroll(position);

            // 若游標剛好位在兩張卡片中間的縫隙，直接由捲動容器尋找相鄰卡片並維持插入提示，避免閃滅或無法放下。
            if (!IsInsidePaperCard(e.OriginalSource as DependencyObject))
            {
                if (TryResolveNearestGapDropTarget(position, out var targetCard, out bool insertAfter) && targetCard != null)
                {
                    ShowDropIndicator(targetCard, insertAfter);
                    e.Effects = DragDropEffects.Move;
                }
                else
                {
                    ClearDropIndicator();
                    e.Effects = DragDropEffects.None;
                }
                e.Handled = true;
            }
        }

        private void UpdateDragAutoScroll(Point position)
        {
            _dragScrollPosition = position;
            if (GetDragScrollDelta(position) == 0) StopDragAutoScroll();
            else if (!_dragScrollTimer.IsEnabled)
            {
                _dragScrollTimer.Start();
                // 一進入邊緣就回應，後續由計時器支援指標停留時持續捲動。
                DragScrollTimer_Tick(this, EventArgs.Empty);
            }
        }

        private double GetDragScrollDelta(Point position)
        {
            double height = MainScrollViewer.ActualHeight;
            double width = MainScrollViewer.ActualWidth;
            if (position.X < 0 || position.X >= width || position.Y < 0 || position.Y >= height || height <= 0)
                return 0;

            double edge = Math.Min(DragScrollEdge, height / 3);
            if (position.Y < edge && MainScrollViewer.VerticalOffset > 0)
                return -DragScrollStep;
            if (position.Y > height - edge && MainScrollViewer.VerticalOffset < MainScrollViewer.ScrollableHeight)
                return DragScrollStep;
            return 0;
        }

        private void DragScrollTimer_Tick(object? sender, EventArgs e)
        {
            double delta = GetDragScrollDelta(_dragScrollPosition);
            if (ViewModel.IsBusy || !ViewModel.IsDraggingPages || delta == 0)
            {
                StopDragAutoScroll();
                return;
            }

            MainScrollViewer.ScrollToVerticalOffset(MainScrollViewer.VerticalOffset + delta);
            MainScrollViewer.UpdateLayout();
            // 捲動後重新命中卡片或縫隙，避免橘綠插入提示停在原本的畫面位置。
            ClearDropIndicator();
            if (TryResolveDropTargetAtScrollPoint(_dragScrollPosition, out var card, out bool insertAfter) && card != null)
            {
                ShowDropIndicator(card, insertAfter);
            }
        }

        private bool TryResolveDropTargetAtScrollPoint(Point scrollPosition, out Border? targetCard, out bool insertAfter)
        {
            targetCard = null;
            insertAfter = false;

            var hit = MainScrollViewer.InputHitTest(scrollPosition) as DependencyObject;
            while (hit != null && hit != MainScrollViewer)
            {
                if (hit is Border { Name: "CardBorder", DataContext: PaperItem item } card)
                {
                    if (item.IsSelected) return false;
                    Point cardPosition = MainScrollViewer.TranslatePoint(scrollPosition, card);
                    targetCard = card;
                    insertAfter = cardPosition.X > card.ActualWidth / 2;
                    return true;
                }
                hit = VisualTreeHelper.GetParent(hit);
            }

            return TryResolveNearestGapDropTarget(scrollPosition, out targetCard, out insertAfter);
        }

        private bool TryResolveNearestGapDropTarget(Point scrollPosition, out Border? targetCard, out bool insertAfter)
        {
            targetCard = null;
            insertAfter = false;

            const double gapThresholdX = 28.0;
            const double verticalTolerance = 10.0;
            double bestDistance = double.MaxValue;

            for (int index = 0; index < ViewModel.Pages.Count; index++)
            {
                var item = ViewModel.Pages[index];
                if (item.IsSelected) continue;

                if (CardsItemsControl.ItemContainerGenerator.ContainerFromIndex(index) is not FrameworkElement container)
                {
                    continue;
                }

                Border? card = container as Border ?? FindVisualChildBorder(container, "CardBorder");
                if (card == null || card.ActualWidth <= 0 || card.ActualHeight <= 0) continue;

                Rect bounds = card.TransformToVisual(MainScrollViewer)
                    .TransformBounds(new Rect(0, 0, card.ActualWidth, card.ActualHeight));

                if (scrollPosition.Y < bounds.Top - verticalTolerance || scrollPosition.Y > bounds.Bottom + verticalTolerance)
                {
                    continue;
                }

                if (scrollPosition.X >= bounds.Left - gapThresholdX && scrollPosition.X <= bounds.Right + gapThresholdX)
                {
                    double edgeDistance = scrollPosition.X < bounds.Left
                        ? bounds.Left - scrollPosition.X
                        : scrollPosition.X > bounds.Right
                            ? scrollPosition.X - bounds.Right
                            : Math.Min(scrollPosition.X - bounds.Left, bounds.Right - scrollPosition.X);

                    if (edgeDistance < bestDistance)
                    {
                        bestDistance = edgeDistance;
                        targetCard = card;
                        insertAfter = scrollPosition.X > (bounds.Left + bounds.Right) / 2;
                    }
                }
            }

            return targetCard != null;
        }

        private static Border? FindVisualChildBorder(DependencyObject parent, string name)
        {
            int count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is Border border && border.Name == name) return border;
                var nested = FindVisualChildBorder(child, name);
                if (nested != null) return nested;
            }
            return null;
        }

        private void MainScrollViewer_PreviewDragLeave(object sender, DragEventArgs e)
        {
            // 子卡片之間移動也會冒出 DragLeave，只在真正離開捲動區時停止。
            Point position = e.GetPosition(MainScrollViewer);
            if (!new Rect(MainScrollViewer.RenderSize).Contains(position))
            {
                StopDragAutoScroll();
                ClearDropIndicator();
            }
        }

        private void MainScrollViewer_PreviewDrop(object sender, DragEventArgs e)
        {
            StopDragAutoScroll();
            if (!e.Handled && !ViewModel.IsBusy && e.Data.GetDataPresent("PaperItemChunk"))
            {
                if (_dropIndicatorItem is { IsSelected: false } indicatorTarget)
                {
                    CommitDropAt(indicatorTarget, _dropIndicatorAfter);
                    e.Handled = true;
                }
                else if (TryResolveDropTargetAtScrollPoint(e.GetPosition(MainScrollViewer), out var card, out bool insertAfter)
                         && card?.DataContext is PaperItem { IsSelected: false } target)
                {
                    CommitDropAt(target, insertAfter);
                    e.Handled = true;
                }
                ClearDropIndicator();
            }
        }

        private void StopDragAutoScroll() => _dragScrollTimer.Stop();

        private void ShowDropIndicator(Border target, bool insertAfter)
        {
            if (target.DataContext is not PaperItem item)
            {
                return;
            }

            // 插入提示使用畫布覆蓋層，不改動卡片 Margin，避免目標因重排而來回抖動。
            _dropIndicatorItem = item;
            _dropIndicatorAfter = insertAfter;
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
            _dropIndicatorAfter = false;
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

        private void MainScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (ViewModel.Pages.Count <= 120) return;
            _viewportSyncTimer.Stop();
            _viewportSyncTimer.Start();
        }

        private void MainScrollViewer_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (ViewModel.Pages.Count <= 120) return;
            _viewportSyncTimer.Stop();
            _viewportSyncTimer.Start();
        }

        private void ViewportSyncTimer_Tick(object? sender, EventArgs e)
        {
            _viewportSyncTimer.Stop();
            if (ViewModel.Pages.Count <= 120 || MainScrollViewer.ActualHeight <= 0) return;

            double viewportH = MainScrollViewer.ViewportHeight > 0 ? MainScrollViewer.ViewportHeight : MainScrollViewer.ActualHeight;
            double loadTop = -viewportH * 0.8;
            double loadBottom = viewportH * 1.8;
            double retainTop = -viewportH * 2.5;
            double retainBottom = viewportH * 3.5;

            var visibleAndBuffer = new List<PaperItem>();
            var retainSet = new HashSet<PaperItem>();

            for (int index = 0; index < ViewModel.Pages.Count; index++)
            {
                var item = ViewModel.Pages[index];
                if (CardsItemsControl.ItemContainerGenerator.ContainerFromIndex(index) is not FrameworkElement container || container.ActualHeight <= 0)
                {
                    continue;
                }

                Rect bounds = container.TransformToVisual(MainScrollViewer)
                    .TransformBounds(new Rect(0, 0, container.ActualWidth, container.ActualHeight));

                if (bounds.Bottom >= retainTop && bounds.Top <= retainBottom)
                {
                    retainSet.Add(item);
                    if (bounds.Bottom >= loadTop && bounds.Top <= loadBottom)
                    {
                        visibleAndBuffer.Add(item);
                    }
                }
            }

            _viewportThumbnailCts?.Cancel();
            _viewportThumbnailCts?.Dispose();
            _viewportThumbnailCts = new CancellationTokenSource();
            _ = ViewModel.UpdateVisibleThumbnailWindowAsync(visibleAndBuffer, retainSet, _viewportThumbnailCts.Token);
        }

        private void MainScrollViewer_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (IsInsidePaperCard(e.OriginalSource as DependencyObject)) return;

            // 點到捲軸（滑塊、軌道、箭頭）時交給捲軸處理，不能當成空白畫布開始框選，否則會搶走滑鼠而無法拖曳捲軸。
            if (IsInsideScrollBar(e.OriginalSource as DependencyObject)) return;

            // 空白畫布是明確的取消選取動作，也同時移除可能殘留的拖曳插入橘框。
            ClearDropIndicator();
            _isBoxSelecting = true;
            _boxSelectionStart = e.GetPosition(SelectionCanvas);
            _initialBoxSelection.Clear();

            foreach (var page in ViewModel.Pages)
            {
                if (page.IsSelected) _initialBoxSelection.Add(page);
            }

            // Ctrl + 空白處保留既有選取，讓後續框選能追加；一般點擊則取消所有選取。
            if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control)
            {
                ViewModel.DeselectAll();
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

            bool selectionChanged = false;
            foreach (var page in ViewModel.Pages)
            {
                var container = CardsItemsControl.ItemContainerGenerator.ContainerFromItem(page) as FrameworkElement;
                if (container == null) continue;

                Rect itemBounds = container.TransformToVisual(SelectionCanvas)
                    .TransformBounds(new Rect(0, 0, container.ActualWidth, container.ActualHeight));
                bool nextSelected = _initialBoxSelection.Contains(page) || selectionBounds.IntersectsWith(itemBounds);
                if (page.IsSelected != nextSelected)
                {
                    page.IsSelected = nextSelected;
                    selectionChanged = true;
                }
            }

            if (selectionChanged)
            {
                ViewModel.NotifySelectionChanged();
            }
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

        private static bool IsInsideTextBox(DependencyObject? source)
        {
            while (source != null)
            {
                if (source is TextBoxBase or PasswordBox)
                {
                    return true;
                }

                source = VisualTreeHelper.GetParent(source);
            }

            return false;
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

        internal static bool IsInsideScrollBar(DependencyObject? source)
        {
            while (source != null)
            {
                if (source is ScrollBar)
                {
                    return true;
                }

                source = VisualTreeHelper.GetParent(source);
            }

            return false;
        }

        private static bool IsInsideElement(DependencyObject? source, DependencyObject container)
        {
            while (source != null)
            {
                if (ReferenceEquals(source, container))
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

        private void CardDelete_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement { DataContext: PaperItem item })
            {
                if (!item.IsSelected)
                {
                    foreach (var p in ViewModel.Pages) p.IsSelected = (p == item);
                    item.IsSelected = true;
                    ViewModel.SelectedPage = item;
                    ViewModel.NotifySelectionChanged();
                }
                ViewModel.DeleteSelected();
            }
            e.Handled = true;
        }

        private void BatchRotateCW_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.RotateSelected(90);
        }

        private void BatchRotateCCW_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.RotateSelected(-90);
        }

        private void ClearProcessingFolders_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel.IsBusy) return;
            var result = StorageMaintenanceService.ClearTemporaryExcept(ViewModel.GetActiveSourcePaths());
            ViewModel.StatusMessage = $"已清理 {result.DeletedFileCount} 個未使用暫存；保留畫布、復原及任務來源。失敗 {result.FailedPaths.Count} 項。";
        }

        private void ChooseOutput_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel.IsBusy) return;
            var dialog = new OpenFolderDialog { Title = "選擇成品輸出資料夾" };
            if (dialog.ShowDialog(this) == true) ViewModel.Options.OutputDirectory = dialog.FolderName;
        }

        private void SelectRange_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new RenameDialog("", 1, "依畫布頁序選取", "輸入畫布位置，例如 1-5,8,12-20；不是原始文件頁碼。", "選取頁面") { Owner = this };
            if (dialog.ShowDialog() != true) return;
            if (!ViewModel.TrySelectRange(dialog.ProposedName, out var error))
                GameDialog.Show(error, "頁碼範圍", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ShowFailed_Click(object sender, RoutedEventArgs e)
        {
            SidebarToggle.IsChecked = false;
            ViewModel.IsConversionQueueExpanded = true;
            var task = ViewModel.ConversionTasks.FirstOrDefault(t => t.State == ConversionTaskState.Failed);
            if (task != null) Dispatcher.BeginInvoke(() =>
            {
                TaskItemsControl.UpdateLayout();
                if (TaskItemsControl.ItemContainerGenerator.ContainerFromItem(task) is FrameworkElement element)
                    element.BringIntoView();
            }, System.Windows.Threading.DispatcherPriority.Loaded);
        }

        private void ExportFileNameTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (ViewModel.ExportAllPdfCommand.CanExecute(null))
                {
                    ViewModel.ExportAllPdfCommand.Execute(null);
                }
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                Keyboard.ClearFocus();
                e.Handled = true;
            }
        }

        #endregion
    }
}
