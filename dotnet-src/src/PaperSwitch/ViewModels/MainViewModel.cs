using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PaperSwitch.Models;
using PaperSwitch.Services;

namespace PaperSwitch.ViewModels
{
    /// <summary>
    /// 主介面 ViewModel：統籌檔案匯入、非同步轉換、紙張排版畫布、多選換位與導出
    /// </summary>
    public partial class MainViewModel : ObservableObject
    {
        public ObservableCollection<PaperItem> Pages { get; } = new();
        public ObservableCollection<ConversionTaskItem> ConversionTasks { get; } = new();

        [ObservableProperty]
        private PaperItem? _selectedPage;

        [ObservableProperty]
        private double _zoomLevel = 160.0; // 預設卡片寬度 (90px ~ 400px)

        [ObservableProperty]
        private bool _isBusy;

        [ObservableProperty]
        private int _draggedPageCount;

        [ObservableProperty]
        private string _statusMessage = "準備就緒，請將文件或圖片拖放至此";

        [ObservableProperty]
        private double _progressValue;

        [ObservableProperty]
        private bool _isProgressIndeterminate = true;

        [ObservableProperty]
        private ExportOptions _options = new();

        [ObservableProperty]
        private bool _isCheckingUpdate;

        [ObservableProperty]
        private string _checkUpdateButtonText = "🔄 檢查更新";

        public string AppVersionBadge => UpdateService.Instance.GetCurrentVersion();

        private readonly PdfService _pdfService = PdfService.Instance;
        private readonly OfficeConverterService _officeService = OfficeConverterService.Instance;
        private readonly ImageConverterService _imageService = ImageConverterService.Instance;
        private readonly ThumbnailCacheService _thumbnailService = ThumbnailCacheService.Instance;
        private readonly UpdateService _updateService = UpdateService.Instance;
        private readonly ArrangementHistory _history = new();
        private readonly SemaphoreSlim _conversionBatchGate = new(1, 1);
        private const int MaximumConcurrentThumbnailRenders = 4;
        private const int LargeCanvasThumbnailWindowThreshold = 120;
        private readonly SemaphoreSlim _thumbnailRenderGate = new(MaximumConcurrentThumbnailRenders, MaximumConcurrentThumbnailRenders);
        private bool _suppressPagesCollectionNotifications;

        public int TotalPageCount => Pages.Count;
        public int SelectedPageCount => Pages.Count(p => p.IsSelected);
        public bool HasPages => Pages.Count > 0;
        public bool HasSelectedPages => SelectedPageCount > 0;
        public bool HasMultipleSelectedPages => SelectedPageCount > 1;
        public bool CanExportContextSelection => !IsBusy && (HasSelectedPages || SelectedPage != null);
        public string ContextExportSelectedPdfText => SelectedPageCount > 1
            ? $"合併匯出已選 {SelectedPageCount} 頁 (PDF)"
            : "匯出這頁 (PDF)";
        public string ContextExportSplitSelectedPdfText => SelectedPageCount > 1
            ? $"拆分成 {SelectedPageCount} 份獨立 PDF"
            : "獨立匯出這頁 (PDF)";
        public string ContextExportSelectedPngText => SelectedPageCount > 1
            ? $"匯出已選 {SelectedPageCount} 頁 (PNG)"
            : "匯出這頁 (PNG)";
        public bool IsDraggingPages => DraggedPageCount > 0;
        public bool CanUndo => !IsBusy && _history.Undo.Count > 0;
        public bool CanRedo => !IsBusy && _history.Redo.Count > 0;
        public bool HasConversionTasks => ConversionTasks.Count > 0;
        public int FailedConversionTaskCount => ConversionTasks.Count(task => task.State == ConversionTaskState.Failed);
        public int ActiveConversionTaskCount => ConversionTasks.Count(task => task.State is ConversionTaskState.Waiting or ConversionTaskState.Processing);
        public string ConversionQueueSummary => ActiveConversionTaskCount > 0
            ? $"{ActiveConversionTaskCount} 個檔案待處理"
            : FailedConversionTaskCount > 0
                ? $"{FailedConversionTaskCount} 個檔案需要處理"
                : "本輪轉檔已完成";

        [ObservableProperty]
        private bool _isConversionQueueExpanded;

        public bool CanExportSplitPdf => CanExport && !HasSelectedPages;
        public bool CanExportSelectedPdf => CanExport && HasSelectedPages;
        public string ImageExportScopeText => HasSelectedPages
            ? $"將選取的 {SelectedPageCount} 頁轉成 PNG 圖片"
            : $"將畫布全部 {TotalPageCount} 頁轉成 PNG 圖片";
        public string ExportAllPdfButtonText => TotalPageCount > 0
            ? $"✨ 匯出全部 PDF ({TotalPageCount} 頁)"
            : "✨ 匯出全部 PDF";

        public string ExportSelectedPdfButtonText => SelectedPageCount > 0
            ? $"📑 另存選取 ({SelectedPageCount} 頁)"
            : (TotalPageCount > 0 ? $"📑 另存選取 (全部 {TotalPageCount} 頁)" : "📑 另存選取");

        public string ExportSplitPdfButtonText => SelectedPageCount > 0
            ? $"📄 拆分選取 ({SelectedPageCount} 頁)"
            : (TotalPageCount > 0 ? $"📄 拆分選取 (全部 {TotalPageCount} 頁)" : "📄 拆分選取");

        public string ExportImagesButtonText => SelectedPageCount > 0
            ? $"🖼️ 匯出圖片 ({SelectedPageCount} 頁)"
            : (TotalPageCount > 0 ? $"🖼️ 匯出圖片 (全部 {TotalPageCount} 頁)" : "🖼️ 匯出圖片");

        public string SummaryText => $"已裝載 {TotalPageCount} 頁紙張" + (SelectedPageCount > 0 ? $" (已選取 {SelectedPageCount} 頁)" : string.Empty);

        private readonly ExportPreferencesStore _exportPreferences;

        public string ExportFileNamePreview
        {
            get
            {
                var first = GetEffectiveExportPages().FirstOrDefault();
                string source = first == null ? string.Empty : ExportFileNaming.SourceName(first);
                int page = first?.DisplayPageNumber ?? 1;
                bool description = Options.IncludeFileNameDescription;
                return $"全部：{ExportFileNaming.Merged(Options.CustomFileName, false, description, ExportFileNaming.FirstSourceName(Pages))}\n" +
                    $"另存：{ExportFileNaming.Merged(Options.CustomFileName, HasSelectedPages, description)}\n" +
                    $"拆分：{ExportFileNaming.Page(source, page, 1, Options.CustomFileName, ".pdf", description)}\n" +
                    $"圖片：{ExportFileNaming.Page(source, page, 1, Options.CustomFileName, ".png", description)}";
            }
        }

        private void ExportOptionsChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName is nameof(ExportOptions.CustomFileName) or nameof(ExportOptions.IncludeFileNameDescription))
                OnPropertyChanged(nameof(ExportFileNamePreview));
            if (e.PropertyName is nameof(ExportOptions.IncludeFileNameDescription) or nameof(ExportOptions.OutputDirectory) && !_exportPreferences.Save(Options.IncludeFileNameDescription, Options.OutputDirectory))
                StatusMessage = "命名設定無法儲存，下次啟動將使用原設定；本次匯出仍可使用目前選項。";
        }

        partial void OnOptionsChanged(ExportOptions? oldValue, ExportOptions newValue)
        {
            if (oldValue != null) oldValue.PropertyChanged -= ExportOptionsChanged;
            newValue.PropertyChanged += ExportOptionsChanged;
            OnPropertyChanged(nameof(ExportFileNamePreview));
        }

        public MainViewModel() : this(new ExportPreferencesStore()) { }

        public MainViewModel(ExportPreferencesStore exportPreferences)
        {
            _exportPreferences = exportPreferences;
            Options.IncludeFileNameDescription = _exportPreferences.Load();
            Options.PropertyChanged += ExportOptionsChanged;

            // 使用者成品與發行檔分離，建置清理 dist 時不會遺失轉檔資料。
            Options.OutputDirectory = _exportPreferences.LoadOutputDirectory() ?? AppPaths.ConvertedDirectory;

            Pages.CollectionChanged += (s, e) =>
            {
                if (_suppressPagesCollectionNotifications) return;
                NotifyPagesStructureChanged();
            };

            ConversionTasks.CollectionChanged += (s, e) => NotifyConversionQueueChanged();
        }

        private void NotifyPagesStructureChanged()
        {
            OnPropertyChanged(nameof(TotalPageCount));
            OnPropertyChanged(nameof(SelectedPageCount));
            OnPropertyChanged(nameof(HasPages));
            OnPropertyChanged(nameof(HasSelectedPages));
            OnPropertyChanged(nameof(HasMultipleSelectedPages));
            OnPropertyChanged(nameof(SummaryText));
            NotifyDynamicExportButtonTexts();
            NotifyContextActionStates();
            ExportAllPdfCommand.NotifyCanExecuteChanged();
            ExportSelectedPdfCommand.NotifyCanExecuteChanged();
            ExportSplitPdfCommand.NotifyCanExecuteChanged();
            ExportImagesCommand.NotifyCanExecuteChanged();
        }

        public void NotifySelectionChanged()
        {
            OnPropertyChanged(nameof(SelectedPageCount));
            OnPropertyChanged(nameof(HasSelectedPages));
            OnPropertyChanged(nameof(HasMultipleSelectedPages));
            OnPropertyChanged(nameof(SummaryText));
            NotifyDynamicExportButtonTexts();
            NotifyContextActionStates();
            ExportSelectedPdfCommand.NotifyCanExecuteChanged();
            ExportSplitPdfCommand.NotifyCanExecuteChanged();
            ExportImagesCommand.NotifyCanExecuteChanged();
        }

        private void NotifyContextActionStates()
        {
            OnPropertyChanged(nameof(CanExportContextSelection));
            OnPropertyChanged(nameof(ContextExportSelectedPdfText));
            OnPropertyChanged(nameof(ContextExportSplitSelectedPdfText));
            OnPropertyChanged(nameof(ContextExportSelectedPngText));
            ExportContextSelectedPdfCommand.NotifyCanExecuteChanged();
            ExportContextSplitSelectedPdfCommand.NotifyCanExecuteChanged();
            ExportContextSelectedPngCommand.NotifyCanExecuteChanged();
            RotateSelectedClockwiseCommand.NotifyCanExecuteChanged();
            RotateSelectedCounterClockwiseCommand.NotifyCanExecuteChanged();
            SelectSourceCommand.NotifyCanExecuteChanged();
            DeleteSelectedCommand.NotifyCanExecuteChanged();
        }

        partial void OnSelectedPageChanged(PaperItem? value)
        {
            NotifyContextActionStates();
        }

        private void NotifyDynamicExportButtonTexts()
        {
            OnPropertyChanged(nameof(ExportScopeSummary));
            OnPropertyChanged(nameof(ExportFileNamePreview));
            OnPropertyChanged(nameof(ImageExportScopeText));
            OnPropertyChanged(nameof(ExportAllPdfButtonText));
            OnPropertyChanged(nameof(ExportSelectedPdfButtonText));
            OnPropertyChanged(nameof(ExportSplitPdfButtonText));
            OnPropertyChanged(nameof(ExportImagesButtonText));
        }

        private void NotifyConversionQueueChanged()
        {
            OnPropertyChanged(nameof(HasConversionTasks));
            OnPropertyChanged(nameof(FailedConversionTaskCount));
            OnPropertyChanged(nameof(ActiveConversionTaskCount));
            OnPropertyChanged(nameof(ConversionQueueSummary));
            if (ActiveConversionTaskCount > 0 || FailedConversionTaskCount > 0)
            {
                IsConversionQueueExpanded = true;
            }
        }

        partial void OnDraggedPageCountChanged(int value)
        {
            OnPropertyChanged(nameof(IsDraggingPages));
        }

        #region PDF 導出與裝訂

        public bool CanInteract => !IsBusy;
        public string ExportScopeSummary => $"全部 {TotalPageCount} 頁／選取 {SelectedPageCount} 頁；PNG：" + (HasSelectedPages ? "選取頁面" : "全部頁面") + "；檔名留空自動命名";
        public bool CanExport => HasPages && !IsBusy;
        private CancellationTokenSource? _operationCancellation;
        private readonly ExportCoordinator _exportCoordinator = new();

        partial void OnIsBusyChanged(bool value)
        {
            OnPropertyChanged(nameof(CanInteract));
            OnPropertyChanged(nameof(CanExport));
            NotifyHistoryChanged();
            NotifyContextActionStates();
            ExportAllPdfCommand.NotifyCanExecuteChanged();
            ExportSelectedPdfCommand.NotifyCanExecuteChanged();
            ExportSplitPdfCommand.NotifyCanExecuteChanged();
            ExportImagesCommand.NotifyCanExecuteChanged();
        }

        public List<PaperItem> GetContextTargetPages()
        {
            var selected = Pages.Where(p => p.IsSelected).ToList();
            if (selected.Count > 0) return selected;
            return SelectedPage != null && Pages.Contains(SelectedPage)
                ? new List<PaperItem> { SelectedPage }
                : new List<PaperItem>();
        }

        [RelayCommand(CanExecute = nameof(CanExportContextSelection))]
        public Task ExportContextSelectedPdfAsync()
        {
            var targets = GetContextTargetPages();
            if (targets.Count == 0) return Task.CompletedTask;
            var kind = targets.Count > 1 ? ExportKind.SelectedPdf : ExportKind.SplitPdf;
            return ExportCustomPagesAsync(kind, targets, showCompletionDialog: true);
        }

        [RelayCommand(CanExecute = nameof(CanExportContextSelection))]
        public Task ExportContextSplitSelectedPdfAsync()
        {
            var targets = GetContextTargetPages();
            if (targets.Count == 0) return Task.CompletedTask;
            return ExportCustomPagesAsync(ExportKind.SplitPdf, targets, showCompletionDialog: true);
        }

        [RelayCommand(CanExecute = nameof(CanExportContextSelection))]
        public Task ExportContextSelectedPngAsync()
        {
            var targets = GetContextTargetPages();
            if (targets.Count == 0) return Task.CompletedTask;
            return ExportCustomPagesAsync(ExportKind.Images, targets, showCompletionDialog: true);
        }

        [RelayCommand(CanExecute = nameof(CanInteract))]
        public void RotateSelectedClockwise() => RotateSelected(90);

        [RelayCommand(CanExecute = nameof(CanInteract))]
        public void RotateSelectedCounterClockwise() => RotateSelected(-90);

        [RelayCommand]
        public void CancelOperation()
        {
            _operationCancellation?.Cancel();
            StatusMessage = "已要求停止；目前步驟完成後停止，不會強制關閉 Office。";
        }

        public IEnumerable<string> GetActiveSourcePaths() => Pages.Select(p => p.SourceFilePath)
            .Concat(_history.SourcePaths).Concat(ConversionTasks.Select(t => t.ActionPath ?? t.SourcePath));

        public List<PaperItem> GetEffectiveExportPages()
        {
            var selected = Pages.Where(p => p.IsSelected).ToList();
            return selected.Count > 0 ? selected : Pages.ToList();
        }

        [RelayCommand(CanExecute = nameof(CanExport))]
        public Task ExportAllPdfAsync() => RunExportAsync(ExportKind.AllPdf);
        [RelayCommand(CanExecute = nameof(CanExportSplitPdf))]
        public Task ExportSplitPdfAsync() => CanExportSplitPdf ? RunExportAsync(ExportKind.SplitPdf) : Task.CompletedTask;
        [RelayCommand(CanExecute = nameof(CanExportSelectedPdf))]
        public Task ExportSelectedPdfAsync() => CanExportSelectedPdf ? RunExportAsync(ExportKind.SelectedPdf) : Task.CompletedTask;
        [RelayCommand(CanExecute = nameof(CanExport))]
        public Task ExportImagesAsync() => RunExportAsync(ExportKind.Images);
        public Task ExportPdfAsync() => RunExportAsync(Options.MergeIntoSinglePdf ? ExportKind.AllPdf : ExportKind.SplitPdf);

        private Task RunExportAsync(ExportKind kind)
        {
            if (!CanExport) return Task.CompletedTask;
            IEnumerable<PaperItem> targetPages = kind is ExportKind.AllPdf or ExportKind.SplitPdf ? Pages : GetEffectiveExportPages();
            return ExportCustomPagesAsync(kind, targetPages, showCompletionDialog: true);
        }

        /// <summary>
        /// 共用匯出管線，供主視窗與燈箱預覽使用，統一套用目前輸出設定、防重複執行與完成提示。
        /// </summary>
        public Task ExportCustomPagesAsync(ExportKind kind, IEnumerable<PaperItem> targetPages) =>
            ExportCustomPagesAsync(kind, targetPages, showCompletionDialog: true);

        public async Task ExportCustomPagesAsync(ExportKind kind, IEnumerable<PaperItem> targetPages, bool showCompletionDialog)
        {
            if (IsBusy) return;
            var pages = ExportCoordinator.Snapshot(targetPages ?? Array.Empty<PaperItem>());
            if (pages.Count == 0) return;
            var options = ExportCoordinator.Snapshot(Options);
            IsBusy = true;
            _operationCancellation = new();
            IsProgressIndeterminate = kind != ExportKind.Images;
            StatusMessage = $"正在檢查及輸出 {pages.Count} 頁…";
            try
            {
                var progress = new Progress<(int Current, int Total, string FileName)>(p =>
                {
                    ProgressValue = p.Current * 100.0 / Math.Max(1, p.Total);
                    StatusMessage = $"輸出 {p.Current} / {p.Total}：{p.FileName}";
                });
                var result = await _exportCoordinator.ExportAsync(kind, pages, options, _operationCancellation.Token, progress);
                StatusMessage = result.Summary;
                if (showCompletionDialog)
                {
                    if (result.Errors.Count > 0)
                        GameDialog.Show(result.Summary + "\n\n" + string.Join("\n", result.Errors.Take(8)), "輸出需要處理", MessageBoxButton.OK, MessageBoxImage.Warning);
                    else if (result.Files.Count > 0)
                        GameDialog.ShowExportDone(result.Summary, options.OutputDirectory);
                }
            }
            finally
            {
                _operationCancellation.Dispose();
                _operationCancellation = null;
                IsBusy = false;
            }
        }

        [RelayCommand]
        public void OpenConvertedFolder()
        {
            string outputDir = string.IsNullOrWhiteSpace(Options.OutputDirectory)
                ? AppPaths.ConvertedDirectory
                : Options.OutputDirectory;

            if (Directory.Exists(outputDir))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = outputDir,
                    UseShellExecute = true
                });
            }
        }

        [RelayCommand]
        public async Task CheckUpdateAsync()
        {
            if (IsCheckingUpdate) return;

            IsCheckingUpdate = true;
            CheckUpdateButtonText = "⏳ 檢查中...";

            try
            {
                var result = await _updateService.CheckForUpdatesAsync();

                if (!result.Success)
                {
                    GameDialog.Show(
                        $"檢查更新失敗：\n{result.ErrorMessage}\n\n請確認電腦網路連線是否通暢，或稍後再試。",
                        "檢查更新",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                if (!result.HasUpdate)
                {
                    GameDialog.Show(
                        $"🎉 目前已是最新版本 ({result.CurrentVersion})，無需更新！",
                        "檢查更新",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                else
                {
                    var message = $"發現新版本：{result.LatestVersion} (目前版本: {result.CurrentVersion})\n\n" +
                                  (string.IsNullOrWhiteSpace(result.ReleaseNotes) ? string.Empty : $"【更新摘要】：\n{result.ReleaseNotes}\n\n") +
                                  "選擇「是」會下載新版、驗證 SHA-256，並在本程式關閉後自動替換與重新啟動。\n" +
                                  "若驗證資料或寫入權限不足，則會保留目前版本並提供手動下載備援。\n\n是否立即更新？";

                    var dialogResult = GameDialog.Show(
                        message,
                        $"🎉 發現新版本 {result.LatestVersion}",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (dialogResult == MessageBoxResult.Yes)
                    {
                        CheckUpdateButtonText = "⏳ 正在下載並驗證...";
                        var applyResult = await _updateService.DownloadAndApplyAsync(result);
                        if (applyResult.ShouldExitApplication)
                        {
                            GameDialog.Show(
                                applyResult.Message,
                                "正在套用更新",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information);
                            Application.Current.Shutdown();
                            return;
                        }

                        var fallback = GameDialog.Show(
                            $"{applyResult.Message}\n\n是否改為開啟 Release 頁面手動下載？",
                            "無法自動更新",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Warning);
                        if (fallback == MessageBoxResult.Yes && !string.IsNullOrWhiteSpace(result.HtmlUrl))
                        {
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = result.HtmlUrl,
                                UseShellExecute = true
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                GameDialog.Show(
                    $"檢查更新過程發生異常：\n{ex.Message}",
                    "檢查更新",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                IsCheckingUpdate = false;
                CheckUpdateButtonText = "🔄 檢查更新";
            }
        }

        #endregion

        public static string SanitizeFileName(string? fileName, string defaultName = "PaperSwitch_裝訂成品.pdf") =>
            ExportFileNaming.SanitizeFileName(fileName, defaultName);
    }
}
