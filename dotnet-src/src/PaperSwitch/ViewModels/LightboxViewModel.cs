using System;
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
    /// 燈箱大圖檢視 ViewModel
    /// </summary>
    public partial class LightboxViewModel : ObservableObject
    {
        [ObservableProperty]
        private PaperItem _currentPage;

        public ObservableCollection<PaperItem> AllPages { get; }

        [ObservableProperty]
        private int _currentIndex;

        [ObservableProperty]
        private ImageSource? _highResImage;

        [ObservableProperty]
        private bool _isLoading;

        [ObservableProperty]
        private double _zoomScale = 1.0;

        private readonly ThumbnailCacheService _thumbnailService = ThumbnailCacheService.Instance;

        public string PageIndicator => $"{CurrentIndex + 1} / {AllPages.Count}";
        public bool HasPrevious => CurrentIndex > 0;
        public bool HasNext => CurrentIndex < AllPages.Count - 1;

        public LightboxViewModel(PaperItem initialPage, ObservableCollection<PaperItem> allPages)
        {
            _currentPage = initialPage;
            AllPages = allPages;
            _currentIndex = allPages.IndexOf(initialPage);
            if (_currentIndex < 0) _currentIndex = 0;

            _ = LoadCurrentHighResAsync();
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
                    1600
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
            CurrentPage?.RotateClockwise();
        }

        [RelayCommand]
        public void RotateCounterClockwise()
        {
            CurrentPage?.RotateCounterClockwise();
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
