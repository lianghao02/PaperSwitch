using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using System.Windows.Threading;
using PaperSwitch.Models;
using PaperSwitch.ViewModels;

namespace PaperSwitch
{
    public partial class LightboxWindow : Window
    {
        public LightboxViewModel ViewModel { get; }

        private Point? _panStart;
        private Point _panOffset;
        private readonly Cursor _grabCursor = Controls.PreviewCursors.CreateGrab();

        private void PreviewPanStart(object sender, MouseButtonEventArgs e)
        {
            var viewer = PreviewScrollViewer;
            // 捲軸與翻頁按鈕維持原本操作，只在文件區抓取。
            if (ViewModel.IsLoading || (viewer.ScrollableWidth <= 0 && viewer.ScrollableHeight <= 0)) return;
            for (DependencyObject? node = e.OriginalSource as DependencyObject; node != null && node != viewer;
                 node = System.Windows.Media.VisualTreeHelper.GetParent(node))
                if (node is System.Windows.Controls.Primitives.ScrollBar) return;
            _panStart = e.GetPosition(viewer);
            _panOffset = new Point(viewer.HorizontalOffset, viewer.VerticalOffset);
            viewer.CaptureMouse();
            viewer.Cursor = _grabCursor;
            e.Handled = true;
        }

        private void PreviewPanMove(object sender, MouseEventArgs e)
        {
            if (_panStart is not Point start) return;
            if (e.LeftButton != MouseButtonState.Pressed) { EndPan(); return; }
            var position = e.GetPosition(PreviewScrollViewer);
            PreviewScrollViewer.ScrollToHorizontalOffset(_panOffset.X + start.X - position.X);
            PreviewScrollViewer.ScrollToVerticalOffset(_panOffset.Y + start.Y - position.Y);
            e.Handled = true;
        }

        private void PreviewPanEnd(object sender, MouseButtonEventArgs e)
        {
            if (_panStart == null) return;
            EndPan();
            e.Handled = true;
        }

        private void PreviewPanCaptureLost(object sender, MouseEventArgs e) => EndPan();
        private void PreviewScrollChanged(object sender, ScrollChangedEventArgs e) => UpdatePanCursor();

        private void EndPan()
        {
            _panStart = null;
            if (PreviewScrollViewer.IsMouseCaptured) PreviewScrollViewer.ReleaseMouseCapture();
            UpdatePanCursor();
        }

        private void UpdatePanCursor()
        {
            if (_panStart != null) return;
            PreviewScrollViewer.Cursor = PreviewScrollViewer.ScrollableWidth > 0 || PreviewScrollViewer.ScrollableHeight > 0
                ? Cursors.Hand : Cursors.Arrow;
        }

        public LightboxWindow(PaperItem initialPage, ObservableCollection<PaperItem> allPages)
        {
            InitializeComponent();
            ViewModel = new LightboxViewModel(initialPage, allPages);
            DataContext = ViewModel;
            Closed += (_, _) => { EndPan(); _grabCursor.Dispose(); };
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
                e.Handled = true;
                return;
            }
            if (e.Key == Key.Left)
            {
                ViewModel.PreviousPageCommand.Execute(null);
                e.Handled = true;
                return;
            }
            if (e.Key == Key.Right)
            {
                ViewModel.NextPageCommand.Execute(null);
                e.Handled = true;
                return;
            }
            if (e.Key == Key.R)
            {
                if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
                    ViewModel.RotateCounterClockwiseCommand.Execute(null);
                else
                    ViewModel.RotateClockwiseCommand.Execute(null);
                e.Handled = true;
                return;
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void PreviewScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (sender is not ScrollViewer scrollViewer || ViewModel.IsLoading)
            {
                return;
            }

            // 以滑鼠所在位置為縮放錨點，讓閱讀細節時不會跳回左上角。
            var cursor = e.GetPosition(scrollViewer);
            var previousScale = ViewModel.ZoomScale;
            var nextScale = e.Delta > 0
                ? Math.Min(4.0, previousScale + 0.15)
                : Math.Max(0.35, previousScale - 0.15);

            if (Math.Abs(nextScale - previousScale) < double.Epsilon)
            {
                e.Handled = true;
                return;
            }

            var contentX = scrollViewer.HorizontalOffset + cursor.X;
            var contentY = scrollViewer.VerticalOffset + cursor.Y;
            var ratio = nextScale / previousScale;
            ViewModel.ZoomScale = nextScale;

            Dispatcher.BeginInvoke(() =>
            {
                scrollViewer.ScrollToHorizontalOffset(Math.Max(0, contentX * ratio - cursor.X));
                scrollViewer.ScrollToVerticalOffset(Math.Max(0, contentY * ratio - cursor.Y));
            }, DispatcherPriority.Loaded);

            e.Handled = true;
        }
    }
}
