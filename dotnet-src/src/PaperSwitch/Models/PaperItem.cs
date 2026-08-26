using System;
using System.ComponentModel;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;

namespace PaperSwitch.Models
{
    /// <summary>
    /// 代表排版工坊畫布中的單頁紙張資料模型
    /// </summary>
    public partial class PaperItem : ObservableObject
    {
        public string Id { get; } = Guid.NewGuid().ToString("N");

        [ObservableProperty]
        private string _sourceFilePath = string.Empty;

        [ObservableProperty]
        private string _sourceFileName = string.Empty;

        [ObservableProperty]
        private int _sourcePageIndex;

        [ObservableProperty]
        private int _displayPageNumber = 1;

        [ObservableProperty]
        private int _totalPagesInSource = 1;

        [ObservableProperty]
        private int _rotation; // 0, 90, 180, 270

        [ObservableProperty]
        private ImageSource? _thumbnail;

        [ObservableProperty]
        private bool _isSelected;

        [ObservableProperty]
        private bool _isDropTarget;

        [ObservableProperty]
        private bool _isDropAfter;

        [ObservableProperty]
        private bool _isBeingDragged;

        [ObservableProperty]
        private bool _isLoadingThumbnail = true;

        [ObservableProperty]
        private double _originalWidth = 595.0; // 預設 A4 寬度 (pt)

        [ObservableProperty]
        private double _originalHeight = 842.0; // 預設 A4 高度 (pt)

        [ObservableProperty]
        private bool _hasError;

        [ObservableProperty]
        private string? _errorMessage;

        public double AspectRatio => (OriginalHeight > 0) ? (OriginalWidth / OriginalHeight) : 0.707;

        public string PageDisplayInfo => $"{DisplayPageNumber} / {TotalPagesInSource}";

        public string RotationDisplay => Rotation > 0 ? $"{Rotation}°" : string.Empty;

        public void RotateClockwise()
        {
            Rotation = (Rotation + 90) % 360;
            OnPropertyChanged(nameof(RotationDisplay));
        }

        public void RotateCounterClockwise()
        {
            Rotation = (Rotation + 270) % 360;
            OnPropertyChanged(nameof(RotationDisplay));
        }
    }
}
