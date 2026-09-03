using System;
using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;

namespace PaperSwitch.Models
{
    /// <summary>
    /// 單一匯入檔案在批次轉檔流程中的可見狀態。
    /// 僅追蹤工作階段，不會修改原始檔案。
    /// </summary>
    public enum ConversionTaskState
    {
        Waiting,
        Processing,
        Completed,
        Failed,
        Skipped
    }

    public partial class ConversionTaskItem : ObservableObject
    {
        public ConversionTaskItem(string sourcePath)
        {
            SourcePath = sourcePath;
            FileName = Path.GetFileName(sourcePath);
        }

        public string SourcePath { get; }
        public string FileName { get; }

        [ObservableProperty]
        private ConversionTaskState _state = ConversionTaskState.Waiting;

        [ObservableProperty]
        private string _statusMessage = "等待處理";

        [ObservableProperty]
        private string _technicalDetails = string.Empty;

        [ObservableProperty]
        private string? _actionPath;

        [ObservableProperty]
        private int _completedPageCount;

        [ObservableProperty]
        private int _pendingPdfCount;

        [ObservableProperty]
        private bool _hasLoadFailure;

        public string StateIcon => State switch
        {
            ConversionTaskState.Waiting => "○",
            ConversionTaskState.Processing => "▶",
            ConversionTaskState.Completed => "✓",
            ConversionTaskState.Failed => "⚠",
            ConversionTaskState.Skipped => "—",
            _ => "○"
        };

        public string StateLabel => State switch
        {
            ConversionTaskState.Waiting => "等待",
            ConversionTaskState.Processing => "處理中",
            ConversionTaskState.Completed => "完成",
            ConversionTaskState.Failed => "失敗",
            ConversionTaskState.Skipped => "跳過",
            _ => "等待"
        };

        public bool CanRetry => State == ConversionTaskState.Failed;
        public bool CanSkip => State is ConversionTaskState.Waiting or ConversionTaskState.Failed;
        public bool HasTechnicalDetails => !string.IsNullOrWhiteSpace(TechnicalDetails);
        public bool HasActionPath => !string.IsNullOrWhiteSpace(ActionPath);

        partial void OnStateChanged(ConversionTaskState value)
        {
            OnPropertyChanged(nameof(StateIcon));
            OnPropertyChanged(nameof(StateLabel));
            OnPropertyChanged(nameof(CanRetry));
            OnPropertyChanged(nameof(CanSkip));
        }

        partial void OnTechnicalDetailsChanged(string value) => OnPropertyChanged(nameof(HasTechnicalDetails));
        partial void OnActionPathChanged(string? value) => OnPropertyChanged(nameof(HasActionPath));

        public void MarkWaiting()
        {
            State = ConversionTaskState.Waiting;
            StatusMessage = "等待處理";
            TechnicalDetails = string.Empty;
            ActionPath = SourcePath;
            CompletedPageCount = 0;
            PendingPdfCount = 0;
            HasLoadFailure = false;
        }

        public void MarkProcessing(string message)
        {
            State = ConversionTaskState.Processing;
            StatusMessage = message;
            TechnicalDetails = string.Empty;
            ActionPath = SourcePath;
        }

        public void MarkCompleted(int pageCount)
        {
            State = ConversionTaskState.Completed;
            CompletedPageCount = pageCount;
            StatusMessage = pageCount > 0 ? $"完成・已載入 {pageCount} 頁" : "完成";
            TechnicalDetails = string.Empty;
        }

        public void MarkFailed(string message, string? technicalDetails = null, string? actionPath = null)
        {
            State = ConversionTaskState.Failed;
            StatusMessage = message;
            TechnicalDetails = technicalDetails ?? string.Empty;
            ActionPath = actionPath ?? SourcePath;
        }

        public void MarkSkipped(string message = "已跳過，不影響其他檔案")
        {
            State = ConversionTaskState.Skipped;
            StatusMessage = message;
            TechnicalDetails = string.Empty;
            ActionPath = SourcePath;
        }
    }
}
