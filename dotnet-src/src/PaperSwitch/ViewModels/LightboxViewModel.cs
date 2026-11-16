using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PaperSwitch.Models;
using PaperSwitch.Services;

namespace PaperSwitch.ViewModels
{
    /// <summary>
    /// 燈箱對宿主視窗的最小回呼合約，避免燈箱 ViewModel 直接依賴 MainViewModel。
    /// </summary>
    public sealed record LightboxHostCallbacks(
        Action<PaperItem, int>? RotatePage = null,
        Func<ExportKind, IReadOnlyList<PaperItem>, Task>? ExportPagesAsync = null);

    /// <summary>
    /// 燈箱大圖檢視 ViewModel
    /// </summary>
    public partial class LightboxViewModel : ObservableObject
    {
        [ObservableProperty]
        private PaperItem? _currentPage;

        public ObservableCollection<PaperItem> AllPages { get; }

        [ObservableProperty]
        private int _currentIndex;

        [ObservableProperty]
        private ImageSource? _highResImage;

        [ObservableProperty]
        private bool _isLoading;

        [ObservableProperty]
        private bool _isExporting;

        [ObservableProperty]
        private double _zoomScale = 1.0;

        private readonly LightboxHostCallbacks _callbacks;
        private readonly ThumbnailCacheService _thumbnailService = ThumbnailCacheService.Instance;

        public string PageIndicator => $"{CurrentIndex + 1} / {AllPages.Count}";
        public bool HasPrevious => CurrentIndex > 0;
        public bool HasNext => CurrentIndex < AllPages.Count - 1;
        public bool CanExportCurrent => CurrentPage != null && !IsExporting && _callbacks.ExportPagesAsync != null;

        public LightboxViewModel(PaperItem initialPage, ObservableCollection<PaperItem> allPages, Action<PaperItem, int>? rotate = null)
            : this(initialPage, allPages, new LightboxHostCallbacks(RotatePage: rotate))
        {
        }

        public LightboxViewModel(PaperItem initialPage, ObservableCollection<PaperItem> allPages, LightboxHostCallbacks callbacks)
        {
            _callbacks = callbacks ?? new LightboxHostCallbacks();
            _currentPage = initialPage;
            AllPages = allPages;
            _currentIndex = allPages.IndexOf(initialPage);
            if (_currentIndex < 0) _currentIndex = 0;

            _ = LoadCurrentHighResAsync();
        }

        partial void OnCurrentPageChanged(PaperItem? value)
        {
            NotifyExportCanExecuteChanged();
        }

        partial void OnIsExportingChanged(bool value)
        {
            NotifyExportCanExecuteChanged();
        }

        private void NotifyExportCanExecuteChanged()
        {
            OnPropertyChanged(nameof(CanExportCurrent));
            ExportCurrentPdfCommand.NotifyCanExecuteChanged();
            ExportCurrentPngCommand.NotifyCanExecuteChanged();
        }

        private async Task LoadCurrentHighResAsync()
        {
            if (CurrentPage == null) return;

            IsLoading = true;
            OnPropertyChanged(nameof(PageIndicator));
            OnPropertyChanged(nameof(HasPrevious));
            OnPropertyChanged(nameof(HasNext));

            try
            {
                var image = await _thumbnailService.GetHighResPreviewAsync(
                    CurrentPage.SourceFilePath,
                    CurrentPage.SourcePageIndex,
                    3840
                );
                HighResImage = image ?? CurrentPage.Thumbnail;
                ResetZoom();
            }
            finally
            {
                IsLoading = false;
            }
        }

        [RelayCommand]
        public async Task NextPage()
        {
            if (HasNext)
            {
                CurrentIndex++;
                CurrentPage = AllPages[CurrentIndex];
                await LoadCurrentHighResAsync();
            }
        }

        [RelayCommand]
        public async Task PreviousPage()
        {
            if (HasPrevious)
            {
                CurrentIndex--;
                CurrentPage = AllPages[CurrentIndex];
                await LoadCurrentHighResAsync();
            }
        }

        [RelayCommand]
        public void RotateClockwise()
        {
            if (CurrentPage == null) return;
            if (_callbacks.RotatePage != null) _callbacks.RotatePage(CurrentPage, 90);
            else CurrentPage.RotateClockwise();
        }

        [RelayCommand]
        public void RotateCounterClockwise()
        {
            if (CurrentPage == null) return;
            if (_callbacks.RotatePage != null) _callbacks.RotatePage(CurrentPage, -90);
            else CurrentPage.RotateCounterClockwise();
        }

        [RelayCommand(CanExecute = nameof(CanExportCurrent))]
        public Task ExportCurrentPdfAsync() =>
            RunExportAsync(ExportKind.SplitPdf, CurrentPage != null ? new[] { CurrentPage } : Array.Empty<PaperItem>());

        [RelayCommand(CanExecute = nameof(CanExportCurrent))]
        public Task ExportCurrentPngAsync() =>
            RunExportAsync(ExportKind.Images, CurrentPage != null ? new[] { CurrentPage } : Array.Empty<PaperItem>());

        private async Task RunExportAsync(ExportKind kind, IReadOnlyList<PaperItem> pages)
        {
            if (IsExporting || pages.Count == 0 || _callbacks.ExportPagesAsync == null) return;
            IsExporting = true;
            try
            {
                await _callbacks.ExportPagesAsync(kind, pages);
            }
            finally
            {
                IsExporting = false;
            }
        }

        [RelayCommand]
        public void ZoomIn()
        {
            if (ZoomScale < 4.0) ZoomScale = Math.Min(4.0, ZoomScale + 0.25);
        }

        [RelayCommand]
        public void ZoomOut()
        {
            if (ZoomScale > 0.35) ZoomScale = Math.Max(0.35, ZoomScale - 0.25);
        }

        [RelayCommand]
        public void ResetZoom()
        {
            ZoomScale = 1.0;
        }
    }
}
