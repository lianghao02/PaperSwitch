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
        private readonly Stack<ArrangementSnapshot> _undoHistory = new();
        private readonly Stack<ArrangementSnapshot> _redoHistory = new();
        private readonly SemaphoreSlim _conversionBatchGate = new(1, 1);
        private const int MaximumConcurrentThumbnailRenders = 4;

        public int TotalPageCount => Pages.Count;
        public int SelectedPageCount => Pages.Count(p => p.IsSelected);
        public bool HasPages => Pages.Count > 0;
        public bool HasSelectedPages => SelectedPageCount > 0;
        public bool IsDraggingPages => DraggedPageCount > 0;
        public bool CanUndo => _undoHistory.Count > 0;
        public bool CanRedo => _redoHistory.Count > 0;
        public bool HasConversionTasks => ConversionTasks.Count > 0;
        public int FailedConversionTaskCount => ConversionTasks.Count(task => task.State == ConversionTaskState.Failed);
        public int ActiveConversionTaskCount => ConversionTasks.Count(task => task.State is ConversionTaskState.Waiting or ConversionTaskState.Processing);
        public string ConversionQueueSummary => ActiveConversionTaskCount > 0
            ? $"{ActiveConversionTaskCount} 個檔案待處理"
            : FailedConversionTaskCount > 0
                ? $"{FailedConversionTaskCount} 個檔案需要處理"
                : "本輪轉檔已完成";

        public string SummaryText => $"已裝載 {TotalPageCount} 頁紙張" + (SelectedPageCount > 0 ? $" (已選取 {SelectedPageCount} 頁)" : string.Empty);

        public MainViewModel()
        {
            // 使用者成品與發行檔分離，建置清理 dist 時不會遺失轉檔資料。
            Options.OutputDirectory = AppPaths.ConvertedDirectory;

            Pages.CollectionChanged += (s, e) =>
            {
                OnPropertyChanged(nameof(TotalPageCount));
                OnPropertyChanged(nameof(SelectedPageCount));
                OnPropertyChanged(nameof(HasPages));
                OnPropertyChanged(nameof(HasSelectedPages));
                OnPropertyChanged(nameof(SummaryText));
                ExportAllPdfCommand.NotifyCanExecuteChanged();
                ExportSelectedPdfCommand.NotifyCanExecuteChanged();
                ExportSplitPdfCommand.NotifyCanExecuteChanged();
            };

            ConversionTasks.CollectionChanged += (s, e) => NotifyConversionQueueChanged();
        }

        public void NotifySelectionChanged()
        {
            OnPropertyChanged(nameof(SelectedPageCount));
            OnPropertyChanged(nameof(HasSelectedPages));
            OnPropertyChanged(nameof(SummaryText));
            ExportSelectedPdfCommand.NotifyCanExecuteChanged();
            ExportSplitPdfCommand.NotifyCanExecuteChanged();
        }

        private void NotifyConversionQueueChanged()
        {
            OnPropertyChanged(nameof(HasConversionTasks));
            OnPropertyChanged(nameof(FailedConversionTaskCount));
            OnPropertyChanged(nameof(ActiveConversionTaskCount));
            OnPropertyChanged(nameof(ConversionQueueSummary));
        }

        partial void OnDraggedPageCountChanged(int value)
        {
            OnPropertyChanged(nameof(IsDraggingPages));
        }

        private bool ApplyPageEdit(Action editAction)
        {
            var before = CaptureArrangementSnapshot();
            editAction();

            if (before.Matches(Pages, SelectedPage))
            {
                return false;
            }

            _undoHistory.Push(before);
            _redoHistory.Clear();
            NotifyHistoryChanged();
            return true;
        }

        private ArrangementSnapshot CaptureArrangementSnapshot()
        {
            return new ArrangementSnapshot(Pages, SelectedPage);
        }

        private void RestoreArrangementSnapshot(ArrangementSnapshot snapshot)
        {
            Pages.Clear();

            foreach (var state in snapshot.PageStates)
            {
                state.Restore();
                Pages.Add(state.Page);
            }

            SelectedPage = snapshot.SelectedPage is not null && Pages.Contains(snapshot.SelectedPage)
                ? snapshot.SelectedPage
                : Pages.FirstOrDefault(page => page.IsSelected);
            NotifySelectionChanged();
        }

        private void NotifyHistoryChanged()
        {
            OnPropertyChanged(nameof(CanUndo));
            OnPropertyChanged(nameof(CanRedo));
        }

        private void ClearHistory()
        {
            _undoHistory.Clear();
            _redoHistory.Clear();
            NotifyHistoryChanged();
        }

        [RelayCommand]
        public void Undo()
        {
            if (!CanUndo) return;

            _redoHistory.Push(CaptureArrangementSnapshot());
            RestoreArrangementSnapshot(_undoHistory.Pop());
            NotifyHistoryChanged();
            StatusMessage = "已復原上一個編排動作";
        }

        [RelayCommand]
        public void Redo()
        {
            if (!CanRedo) return;

            _undoHistory.Push(CaptureArrangementSnapshot());
            RestoreArrangementSnapshot(_redoHistory.Pop());
            NotifyHistoryChanged();
            StatusMessage = "已重做下一個編排動作";
        }

        [RelayCommand]
        public void SelectAll()
        {
            foreach (var page in Pages)
            {
                page.IsSelected = true;
            }
            NotifySelectionChanged();
        }

        [RelayCommand]
        public void DeselectAll()
        {
            foreach (var page in Pages)
            {
                page.IsSelected = false;
            }
            SelectedPage = null;
            NotifySelectionChanged();
        }

        [RelayCommand]
        public void InvertSelection()
        {
            foreach (var page in Pages)
            {
                page.IsSelected = !page.IsSelected;
            }
            NotifySelectionChanged();
        }

        [RelayCommand]
        public void RotateSelected(int deltaDegrees)
        {
            var targetPages = Pages.Where(p => p.IsSelected).ToList();
            if (targetPages.Count == 0 && SelectedPage != null)
            {
                targetPages.Add(SelectedPage);
            }

            if (targetPages.Count == 0) return;

            if (ApplyPageEdit(() =>
            {
                foreach (var p in targetPages)
                {
                    if (deltaDegrees > 0)
                        p.RotateClockwise();
                    else
                        p.RotateCounterClockwise();
                }
            }))
            {
                StatusMessage = $"已將 {targetPages.Count} 頁紙張旋轉 {deltaDegrees}°";
            }
        }

        [RelayCommand]
        public void DeleteSelected()
        {
            var targetPages = Pages.Where(p => p.IsSelected).ToList();
            if (targetPages.Count == 0 && SelectedPage != null)
            {
                targetPages.Add(SelectedPage);
            }

            if (targetPages.Count == 0) return;

            int nextSelectionIndex = targetPages.Count > 0
                ? Pages.IndexOf(targetPages[0])
                : 0;

            if (!ApplyPageEdit(() =>
            {
                foreach (var p in targetPages)
                {
                    Pages.Remove(p);
                }

                foreach (var page in Pages)
                {
                    page.IsSelected = false;
                }

                if (Pages.Count > 0)
                {
                    SelectedPage = Pages[Math.Min(nextSelectionIndex, Pages.Count - 1)];
                    SelectedPage.IsSelected = true;
                }
                else
                {
                    SelectedPage = null;
                }
            })) return;

            NotifySelectionChanged();
            StatusMessage = $"已移除 {targetPages.Count} 頁紙張";
        }

        [RelayCommand]
        public void ClearAll()
        {
            if (Pages.Count == 0) return;

            if (!ApplyPageEdit(() =>
            {
                Pages.Clear();
                SelectedPage = null;
            })) return;

            _thumbnailService.ClearCache();
            NotifySelectionChanged();
            StatusMessage = "已清空排版畫布";
        }

        [RelayCommand]
        public void InsertBlankPage()
        {
            string blankPath = Path.Combine(AppPaths.TemporaryConversionDirectory, $"blank_{Guid.NewGuid():N}.pdf");
            if (!_pdfService.CreateBlankA4Pdf(blankPath))
            {
                MessageBox.Show("建立空白頁失敗，請確認暫存資料夾是否可寫入。", "插入空白頁", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var item = new PaperItem
            {
                SourceFilePath = blankPath,
                SourceFileName = "空白頁.pdf",
                SourcePageIndex = 0,
                DisplayPageNumber = 1,
                TotalPagesInSource = 1,
                OriginalWidth = 595,
                OriginalHeight = 842,
                IsSelected = true,
                IsLoadingThumbnail = true
            };

            ApplyPageEdit(() =>
            {
                foreach (var page in Pages) page.IsSelected = false;
                int insertIndex = SelectedPage is null ? Pages.Count : Pages.IndexOf(SelectedPage) + 1;
                Pages.Insert(Math.Clamp(insertIndex, 0, Pages.Count), item);
                SelectedPage = item;
            });

            _ = LoadThumbnailAsync(item);
            NotifySelectionChanged();
            StatusMessage = "已插入一張 A4 空白頁";
        }

        /// <summary>
        /// 設定選取紙張在「拆分獨立存檔」模式下的輸出檔名。
        /// 不會重新命名來源檔、Office 暫存檔或使用者原始文件。
        /// </summary>
        public bool TryRenameSelectedPages(string proposedName, out string message)
        {
            var targetPages = Pages.Where(page => page.IsSelected).ToList();
            if (targetPages.Count == 0 && SelectedPage != null)
            {
                targetPages.Add(SelectedPage);
            }

            if (targetPages.Count == 0)
            {
                message = "請先選取要重新命名的紙張。";
                return false;
            }

            string nameStem = Path.GetFileNameWithoutExtension(proposedName.Trim());
            if (string.IsNullOrWhiteSpace(nameStem) || nameStem.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                message = "檔名不可空白，也不可包含 \\ / : * ? \" < > | 等字元。";
                return false;
            }

            if (!ApplyPageEdit(() =>
            {
                if (targetPages.Count == 1)
                {
                    targetPages[0].SourceFileName = $"{nameStem}.pdf";
                }
                else
                {
                    int digits = Math.Max(3, targetPages.Count.ToString().Length);
                    for (int index = 0; index < targetPages.Count; index++)
                    {
                        targetPages[index].SourceFileName = $"{nameStem}_{(index + 1).ToString($"D{digits}")}.pdf";
                    }
                }
            }))
            {
                message = "選取紙張的名稱未變更。";
                return false;
            }

            message = targetPages.Count == 1
                ? $"已將選取紙張命名為「{nameStem}.pdf」"
                : $"已將 {targetPages.Count} 張選取紙張命名為「{nameStem}_001.pdf」起的連號名稱";
            StatusMessage = message;
            return true;
        }

        #region 鍵盤快速導航與多選位移演算法

        [RelayCommand]
        public void MoveSelectedLeft()
        {
            ApplyPageEdit(MoveSelectedLeftCore);
        }

        private void MoveSelectedLeftCore()
        {
            // 向左移動 (←)：由左至右 (最小索引開始) 依序交換，防範索引覆寫衝突
            var selectedIndices = Pages
                .Select((p, idx) => (Page: p, Index: idx))
                .Where(x => x.Page.IsSelected)
                .OrderBy(x => x.Index)
                .ToList();

            if (selectedIndices.Count == 0 || selectedIndices.First().Index == 0) return;

            foreach (var item in selectedIndices)
            {
                int curIdx = Pages.IndexOf(item.Page);
                if (curIdx > 0 && !Pages[curIdx - 1].IsSelected)
                {
                    Pages.Move(curIdx, curIdx - 1);
                }
            }
        }

        [RelayCommand]
        public void MoveSelectedRight()
        {
            ApplyPageEdit(MoveSelectedRightCore);
        }

        private void MoveSelectedRightCore()
        {
            // 向右移動 (→)：由右至左 (最大索引開始) 依序交換
            var selectedIndices = Pages
                .Select((p, idx) => (Page: p, Index: idx))
                .Where(x => x.Page.IsSelected)
                .OrderByDescending(x => x.Index)
                .ToList();

            if (selectedIndices.Count == 0 || selectedIndices.First().Index == Pages.Count - 1) return;

            foreach (var item in selectedIndices)
            {
                int curIdx = Pages.IndexOf(item.Page);
                if (curIdx < Pages.Count - 1 && !Pages[curIdx + 1].IsSelected)
                {
                    Pages.Move(curIdx, curIdx + 1);
                }
            }
        }

        /// <summary>
        /// 依畫布方向移動選取群組；負值往前／上，正值往後／下。
        /// </summary>
        public void MoveSelectedBy(int positions)
        {
            if (positions == 0) return;

            ApplyPageEdit(() =>
            {
                int steps = Math.Abs(positions);
                for (int i = 0; i < steps; i++)
                {
                    if (positions < 0)
                        MoveSelectedLeftCore();
                    else
                        MoveSelectedRightCore();
                }
            });
        }

        [RelayCommand]
        public void MoveSelectedToStart()
        {
            var selectedItems = Pages.Where(p => p.IsSelected).ToList();
            if (selectedItems.Count == 0) return;

            ApplyPageEdit(() =>
            {
                for (int i = 0; i < selectedItems.Count; i++)
                {
                    var item = selectedItems[i];
                    int curIdx = Pages.IndexOf(item);
                    Pages.Move(curIdx, i);
                }
            });
        }

        [RelayCommand]
        public void MoveSelectedToEnd()
        {
            var selectedItems = Pages.Where(p => p.IsSelected).ToList();
            if (selectedItems.Count == 0) return;

            ApplyPageEdit(() =>
            {
                int targetIndex = Pages.Count - 1;
                for (int i = selectedItems.Count - 1; i >= 0; i--)
                {
                    var item = selectedItems[i];
                    int curIdx = Pages.IndexOf(item);
                    Pages.Move(curIdx, targetIndex);
                }
            });
        }

        /// <summary>
        /// 多選打包拖曳抽取與整包重新插入演算法
        /// </summary>
        public void RelocateSelectedChunk(int targetInsertionIndex)
        {
            var selectedItems = Pages.Where(p => p.IsSelected).ToList();
            if (selectedItems.Count == 0) return;

            if (targetInsertionIndex < 0) targetInsertionIndex = 0;
            if (targetInsertionIndex > Pages.Count) targetInsertionIndex = Pages.Count;

            ApplyPageEdit(() =>
            {
                // 取得目標插入錨點項目
                PaperItem? anchorItem = targetInsertionIndex < Pages.Count ? Pages[targetInsertionIndex] : null;

                // 移除所有選取項目
                foreach (var item in selectedItems)
                {
                    Pages.Remove(item);
                }

                // 計算新的插入點
                int newInsertPos = anchorItem != null ? Pages.IndexOf(anchorItem) : Pages.Count;
                if (newInsertPos < 0) newInsertPos = Pages.Count;

                for (int i = 0; i < selectedItems.Count; i++)
                {
                    Pages.Insert(newInsertPos + i, selectedItems[i]);
                }
            });

            NotifySelectionChanged();
        }

        #endregion

        #region 檔案載入與轉檔排程

        public async Task AddFilesAsync(string[] filePaths)
        {
            if (filePaths == null || filePaths.Length == 0) return;

            var existingFiles = filePaths.Where(File.Exists).ToArray();
            if (existingFiles.Length == 0) return;

            var tasks = existingFiles.Select(CreateConversionTask).ToList();
            await ProcessConversionTasksAsync(tasks);
        }

        private ConversionTaskItem CreateConversionTask(string sourcePath)
        {
            var task = new ConversionTaskItem(sourcePath);
            task.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName is nameof(ConversionTaskItem.State) or nameof(ConversionTaskItem.StatusMessage))
                {
                    NotifyConversionQueueChanged();
                }
            };
            ConversionTasks.Add(task);
            return task;
        }

        private async Task ProcessConversionTasksAsync(IReadOnlyList<ConversionTaskItem> tasks)
        {
            if (tasks.Count == 0) return;

            await _conversionBatchGate.WaitAsync();
            try
            {
                IsBusy = true;
                IsProgressIndeterminate = false;
                ProgressValue = 0;
                StatusMessage = $"準備處理 {tasks.Count} 個檔案...";

                string workTempDir = AppPaths.TemporaryConversionDirectory;
                var pdfsToLoad = new List<(string PdfPath, ConversionTaskItem Task)>();

                for (int fileIndex = 0; fileIndex < tasks.Count; fileIndex++)
                {
                    var task = tasks[fileIndex];
                    if (task.State == ConversionTaskState.Skipped)
                    {
                        continue;
                    }

                    int displayIndex = fileIndex + 1;
                    StatusMessage = $"正在處理第 {displayIndex} / {tasks.Count} 個檔案：\n{task.FileName}";
                    await ProcessSingleConversionTaskAsync(task, workTempDir, pdfsToLoad);
                    ProgressValue = displayIndex * 70.0 / tasks.Count;
                }

                ProgressValue = 75;
                StatusMessage = pdfsToLoad.Count > 0
                    ? $"正在讀取 {pdfsToLoad.Count} 份 PDF 的頁面與縮圖..."
                    : "本輪沒有可載入的 PDF 文件。";
                var newItems = new List<PaperItem>();

                for (int pdfIndex = 0; pdfIndex < pdfsToLoad.Count; pdfIndex++)
                {
                    var (pdf, task) = pdfsToLoad[pdfIndex];
                    StatusMessage = $"正在讀取第 {pdfIndex + 1} / {pdfsToLoad.Count} 份 PDF：\n{Path.GetFileName(pdf)}";
                    await LoadPdfPagesForTaskAsync(pdf, task, newItems);
                    ProgressValue = 75 + ((pdfIndex + 1) * 20.0 / Math.Max(pdfsToLoad.Count, 1));
                }

                foreach (var task in tasks.Where(task => task.State == ConversionTaskState.Processing && task.PendingPdfCount == 0))
                {
                    task.MarkCompleted(task.CompletedPageCount);
                }

                _ = LoadThumbnailsAsync(newItems);

                if (newItems.Count > 0)
                {
                    ClearHistory();
                }

                ProgressValue = 100;
                int failedCount = tasks.Count(task => task.State == ConversionTaskState.Failed);
                StatusMessage = failedCount == 0
                    ? $"已順利裝載 {newItems.Count} 頁新紙張至工坊畫布"
                    : $"已裝載 {newItems.Count} 頁紙張；另有 {failedCount} 個檔案需要處理";
            }
            catch (Exception ex)
            {
                foreach (var task in tasks.Where(task => task.State == ConversionTaskState.Processing))
                {
                    task.MarkFailed("此檔案未能完成處理。原始檔案沒有被修改，請確認後重新嘗試。", ex.ToString());
                }

                StatusMessage = "本輪處理發生未預期錯誤，請查看任務清單中的失敗原因。";
            }
            finally
            {
                IsBusy = false;
                IsProgressIndeterminate = true;
                NotifySelectionChanged();
                NotifyConversionQueueChanged();
                _conversionBatchGate.Release();
            }
        }

        private async Task ProcessSingleConversionTaskAsync(
            ConversionTaskItem task,
            string workTempDir,
            ICollection<(string PdfPath, ConversionTaskItem Task)> pdfsToLoad)
        {
            if (!File.Exists(task.SourcePath))
            {
                task.MarkFailed("找不到原始檔案，請確認檔案未被移動或刪除。", actionPath: task.SourcePath);
                return;
            }

            string extension = Path.GetExtension(task.SourcePath).ToLowerInvariant();
            task.MarkProcessing("正在轉換...");

            try
            {
                if (extension == ".pdf")
                {
                    RegisterPdfForLoading(task, task.SourcePath, pdfsToLoad);
                    return;
                }

                if (_officeService.IsWordFile(task.SourcePath))
                {
                    string outputPath = Path.Combine(workTempDir, $"{Path.GetFileNameWithoutExtension(task.SourcePath)}_{Guid.NewGuid():N}.pdf");
                    var result = await _officeService.ConvertWordAsync(task.SourcePath, outputPath);
                    await RegisterOfficeResultAsync(task, result, pdfsToLoad);
                    return;
                }

                if (_officeService.IsExcelFile(task.SourcePath))
                {
                    var result = await _officeService.ConvertExcelAsync(task.SourcePath, workTempDir, Options.ExcelFitToPage);
                    await RegisterOfficeResultAsync(task, result, pdfsToLoad);
                    return;
                }

                if (_officeService.IsPowerPointFile(task.SourcePath))
                {
                    string outputPath = Path.Combine(workTempDir, $"{Path.GetFileNameWithoutExtension(task.SourcePath)}_{Guid.NewGuid():N}.pdf");
                    var result = await _officeService.ConvertPowerPointAsync(task.SourcePath, outputPath);
                    await RegisterOfficeResultAsync(task, result, pdfsToLoad);
                    return;
                }

                if (_imageService.IsImageFile(task.SourcePath))
                {
                    string outputPath = Path.Combine(workTempDir, $"{Path.GetFileNameWithoutExtension(task.SourcePath)}_{Guid.NewGuid():N}.pdf");
                    if (_imageService.ConvertImagesToPdf(new[] { task.SourcePath }, outputPath))
                    {
                        RegisterPdfForLoading(task, outputPath, pdfsToLoad);
                    }
                    else
                    {
                        task.MarkFailed("圖片無法轉為 PDF。原始圖片沒有被修改，請確認檔案未損毀後重新嘗試。");
                    }
                    return;
                }

                task.MarkSkipped("此檔案格式目前不支援，已跳過且不影響其他檔案。");
            }
            catch (Exception ex)
            {
                task.MarkFailed("轉檔時發生未預期錯誤。原始檔案沒有被修改，請確認後重新嘗試。", ex.ToString());
            }
        }

        private async Task RegisterOfficeResultAsync(
            ConversionTaskItem task,
            ConversionResult result,
            ICollection<(string PdfPath, ConversionTaskItem Task)> pdfsToLoad)
        {
            if (!result.Success)
            {
                MarkFriendlyConversionFailure(task, result);
                return;
            }

            foreach (var pdfPath in result.GeneratedPdfPaths)
            {
                if (await _officeService.WaitForPdfReadyAsync(pdfPath))
                {
                    RegisterPdfForLoading(task, pdfPath, pdfsToLoad);
                    continue;
                }

                task.MarkFailed(
                    "Office 已產生暫存 PDF，但目前仍受加密或保護程序處理，請依規定解密後重新加入。原始檔案沒有被修改。",
                    "Office 暫存 PDF 未在可等待期限內成為標準 %PDF 檔案。",
                    pdfPath);
            }
        }

        private static void RegisterPdfForLoading(
            ConversionTaskItem task,
            string pdfPath,
            ICollection<(string PdfPath, ConversionTaskItem Task)> pdfsToLoad)
        {
            task.PendingPdfCount++;
            task.StatusMessage = "轉換完成，正在讀取 PDF 頁面...";
            pdfsToLoad.Add((pdfPath, task));
        }

        private async Task LoadPdfPagesForTaskAsync(string pdfPath, ConversionTaskItem task, ICollection<PaperItem> newItems)
        {
            try
            {
                int totalPages = _pdfService.GetPageCount(pdfPath);
                if (totalPages <= 0)
                {
                    task.MarkFailed("PDF 無法讀取或內容已損毀。原始檔案沒有被修改，請確認檔案後重新嘗試。", actionPath: pdfPath);
                    return;
                }

                string fileName = Path.GetFileName(pdfPath);
                for (int pageIndex = 0; pageIndex < totalPages; pageIndex++)
                {
                    var (width, height, _) = _pdfService.GetPageDimensions(pdfPath, pageIndex);
                    var item = new PaperItem
                    {
                        SourceFilePath = pdfPath,
                        SourceFileName = fileName,
                        SourcePageIndex = pageIndex,
                        DisplayPageNumber = pageIndex + 1,
                        TotalPagesInSource = totalPages,
                        Rotation = 0,
                        OriginalWidth = width,
                        OriginalHeight = height,
                        IsLoadingThumbnail = true
                    };
                    newItems.Add(item);
                    Pages.Add(item);
                }

                task.CompletedPageCount += totalPages;
            }
            catch (Exception ex)
            {
                task.MarkFailed("PDF 頁面載入失敗。原始檔案沒有被修改，請確認檔案後重新嘗試。", ex.ToString(), pdfPath);
            }
            finally
            {
                task.PendingPdfCount = Math.Max(0, task.PendingPdfCount - 1);
                if (task.State == ConversionTaskState.Processing && task.PendingPdfCount == 0)
                {
                    task.MarkCompleted(task.CompletedPageCount);
                }
            }

            await Task.CompletedTask;
        }

        private static void MarkFriendlyConversionFailure(ConversionTaskItem task, ConversionResult result)
        {
            string technical = string.IsNullOrWhiteSpace(result.DiagnosticInfo) ? result.ErrorMessage ?? string.Empty : result.DiagnosticInfo;
            string message = result.ErrorMessage ?? "轉檔失敗。";
            string lower = technical.ToLowerInvariant();

            if (message.Contains("未安裝", StringComparison.OrdinalIgnoreCase))
            {
                task.MarkFailed("找不到相容的 Microsoft Office。請確認對應的 Word、Excel 或 PowerPoint 已安裝。原始檔案沒有被修改。", technical);
            }
            else if (lower.Contains("used") || lower.Contains("lock") || lower.Contains("使用中") || lower.Contains("拒絕存取") || lower.Contains("access denied"))
            {
                task.MarkFailed("檔案可能正被 Office 或其他程式開啟。請關閉檔案後重新嘗試；原始檔案沒有被修改。", technical);
            }
            else if (lower.Contains("0x800") || lower.Contains("comexception") || lower.Contains("無法啟動"))
            {
                task.MarkFailed("Office 無法啟動或被系統政策限制。請關閉 Office 後重試，並確認授權與資訊安全設定。原始檔案沒有被修改。", technical);
            }
            else
            {
                task.MarkFailed("無法轉換此檔案。請確認檔案未損毀、未受保護且未被其他程式使用；原始檔案沒有被修改。", technical);
            }
        }

        [RelayCommand]
        public async Task RetryConversionTaskAsync(ConversionTaskItem? task)
        {
            if (task is null || !task.CanRetry || IsBusy) return;

            task.MarkWaiting();
            await ProcessConversionTasksAsync(new[] { task });
        }

        [RelayCommand]
        public void SkipConversionTask(ConversionTaskItem? task)
        {
            if (task is null || !task.CanSkip) return;
            task.MarkSkipped();
            NotifyConversionQueueChanged();
        }

        [RelayCommand]
        public void OpenConversionTaskLocation(ConversionTaskItem? task)
        {
            if (task is null) return;
            string path = task.ActionPath ?? task.SourcePath;
            if (File.Exists(path))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"/select,\"{path}\"",
                    UseShellExecute = true
                });
            }
        }

        private async Task LoadThumbnailsAsync(List<PaperItem> items)
        {
            using var renderSlots = new System.Threading.SemaphoreSlim(MaximumConcurrentThumbnailRenders);
            var tasks = items.Select(async item =>
            {
                await renderSlots.WaitAsync();
                try
                {
                    await LoadThumbnailAsync(item);
                }
                finally
                {
                    renderSlots.Release();
                }
            });
            await Task.WhenAll(tasks);
        }

        private async Task LoadThumbnailAsync(PaperItem item)
        {
            try
            {
                var thumb = await _thumbnailService.GetThumbnailAsync(item.SourceFilePath, item.SourcePageIndex, 360);
                item.Thumbnail = thumb;
                item.IsLoadingThumbnail = false;
            }
            catch
            {
                item.IsLoadingThumbnail = false;
                item.HasError = true;
            }
        }

        #endregion

        #region PDF 導出與裝訂

        [RelayCommand]
        public async Task ExportPdfAsync()
        {
            if (Pages.Count == 0)
            {
                MessageBox.Show("目前畫布無任何紙張可供導出！", "提醒", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            IsBusy = true;
            IsProgressIndeterminate = true;
            StatusMessage = "正在進行向量無損裝訂合成...";

            try
            {
                string outputDir = string.IsNullOrWhiteSpace(Options.OutputDirectory)
                    ? AppPaths.ConvertedDirectory
                    : Options.OutputDirectory;

                Directory.CreateDirectory(outputDir);

                if (Options.MergeIntoSinglePdf)
                {
                    string safeBaseName = SanitizeFileName(Options.CustomFileName, "PaperSwitch_裝訂成品.pdf");
                    string finalName = safeBaseName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ? safeBaseName : safeBaseName + ".pdf";

                    string outputPath = Path.Combine(outputDir, finalName);
                    int counter = 1;
                    while (File.Exists(outputPath))
                    {
                        outputPath = Path.Combine(outputDir, $"{Path.GetFileNameWithoutExtension(finalName)}_{counter}.pdf");
                        counter++;
                    }

                    bool ok = await Task.Run(() => _pdfService.ExportArrangedPdf(Pages, outputPath));
                    if (ok)
                    {
                        StatusMessage = $"裝訂完成！已匯出至: {Path.GetFileName(outputPath)}";
                        OpenConvertedFolder();
                    }
                    else
                    {
                        MessageBox.Show("導出 PDF 過程發生錯誤，請檢查來源檔案。", "導出失敗", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    string safePrefix = string.IsNullOrWhiteSpace(Options.CustomFileName) ? string.Empty : SanitizeFileName(Options.CustomFileName, string.Empty);
                    var files = await Task.Run(() => _pdfService.ExportIndividualPdfs(Pages, outputDir, safePrefix));
                    StatusMessage = $"裝訂完成！共匯出 {files.Count} 個獨立單頁 PDF";
                    OpenConvertedFolder();
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"導出失敗: {ex.Message}";
                MessageBox.Show($"導出過程發生錯誤: {ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        /// <summary>
        /// 取得有效匯出頁面清單：若有勾選/選取頁面則只處理選取頁面；若無選取則處理畫布全部頁面。
        /// </summary>
        public List<PaperItem> GetEffectiveExportPages()
        {
            var selected = Pages.Where(p => p.IsSelected).ToList();
            return selected.Count > 0 ? selected : Pages.ToList();
        }

        [RelayCommand(CanExecute = nameof(HasPages))]
        public async Task ExportAllPdfAsync()
        {
            Options.MergeIntoSinglePdf = true;
            await ExportPdfAsync();
        }

        [RelayCommand(CanExecute = nameof(HasPages))]
        public async Task ExportSplitPdfAsync()
        {
            var targetPages = GetEffectiveExportPages();
            if (targetPages.Count == 0)
            {
                MessageBox.Show("目前畫布無任何紙張可供導出！", "提醒", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            IsBusy = true;
            IsProgressIndeterminate = true;
            StatusMessage = "正在進行單頁獨立切分導出...";

            try
            {
                string outputDir = string.IsNullOrWhiteSpace(Options.OutputDirectory)
                    ? AppPaths.ConvertedDirectory
                    : Options.OutputDirectory;

                Directory.CreateDirectory(outputDir);

                string safePrefix = string.IsNullOrWhiteSpace(Options.CustomFileName) ? string.Empty : SanitizeFileName(Options.CustomFileName, string.Empty);
                var files = await Task.Run(() => _pdfService.ExportIndividualPdfs(targetPages, outputDir, safePrefix));
                StatusMessage = $"裝訂完成！共匯出 {files.Count} 個獨立單頁 PDF";
                OpenConvertedFolder();
            }
            catch (Exception ex)
            {
                StatusMessage = $"導出失敗: {ex.Message}";
                MessageBox.Show($"導出過程發生錯誤: {ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        [RelayCommand(CanExecute = nameof(HasPages))]
        public async Task ExportSelectedPdfAsync()
        {
            var targetPages = GetEffectiveExportPages();
            if (targetPages.Count == 0)
            {
                MessageBox.Show("目前畫布無任何紙張可供導出！", "提醒", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            bool hasSelection = HasSelectedPages;
            IsBusy = true;
            IsProgressIndeterminate = true;
            StatusMessage = hasSelection
                ? "正在進行選取頁面裝訂合成..."
                : "正在進行畫布全部頁面裝訂合成...";

            try
            {
                string outputDir = string.IsNullOrWhiteSpace(Options.OutputDirectory)
                    ? AppPaths.ConvertedDirectory
                    : Options.OutputDirectory;

                Directory.CreateDirectory(outputDir);

                string defaultName = hasSelection ? "PaperSwitch_選取頁面.pdf" : "PaperSwitch_裝訂成品.pdf";
                string safeBaseName;
                if (!string.IsNullOrWhiteSpace(Options.CustomFileName))
                {
                    string stem = Path.GetFileNameWithoutExtension(Options.CustomFileName.Trim());
                    safeBaseName = hasSelection
                        ? SanitizeFileName($"{stem}_選取頁面.pdf", defaultName)
                        : SanitizeFileName($"{stem}.pdf", defaultName);
                }
                else
                {
                    safeBaseName = defaultName;
                }

                string finalName = safeBaseName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ? safeBaseName : safeBaseName + ".pdf";

                string outputPath = Path.Combine(outputDir, finalName);
                int counter = 1;
                while (File.Exists(outputPath))
                {
                    outputPath = Path.Combine(outputDir, $"{Path.GetFileNameWithoutExtension(finalName)}_{counter}.pdf");
                    counter++;
                }

                bool ok = await Task.Run(() => _pdfService.ExportArrangedPdf(targetPages, outputPath));
                if (ok)
                {
                    StatusMessage = hasSelection
                        ? $"裝訂完成！已匯出選取頁面至: {Path.GetFileName(outputPath)}"
                        : $"裝訂完成！已匯出至: {Path.GetFileName(outputPath)}";
                    OpenConvertedFolder();
                }
                else
                {
                    MessageBox.Show("導出 PDF 過程發生錯誤，請檢查來源檔案。", "導出失敗", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"導出失敗: {ex.Message}";
                MessageBox.Show($"導出過程發生錯誤: {ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
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
                    MessageBox.Show(
                        $"檢查更新失敗：\n{result.ErrorMessage}\n\n請確認電腦網路連線是否通暢，或稍後再試。",
                        "檢查更新",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                if (!result.HasUpdate)
                {
                    MessageBox.Show(
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

                    var dialogResult = MessageBox.Show(
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
                            MessageBox.Show(
                                applyResult.Message,
                                "正在套用更新",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information);
                            Application.Current.Shutdown();
                            return;
                        }

                        var fallback = MessageBox.Show(
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
                MessageBox.Show(
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

        private sealed class ArrangementSnapshot
        {
            public IReadOnlyList<PageState> PageStates { get; }
            public PaperItem? SelectedPage { get; }

            public ArrangementSnapshot(IEnumerable<PaperItem> pages, PaperItem? selectedPage)
            {
                PageStates = pages.Select(page => new PageState(page)).ToList();
                SelectedPage = selectedPage;
            }

            public bool Matches(IEnumerable<PaperItem> pages, PaperItem? selectedPage)
            {
                var currentPages = pages.ToList();
                return SelectedPage == selectedPage
                    && PageStates.Count == currentPages.Count
                    && PageStates.Zip(currentPages, (state, page) => state.Matches(page)).All(matches => matches);
            }
        }

        private sealed class PageState
        {
            public PaperItem Page { get; }
            private int Rotation { get; }
            private bool IsSelected { get; }
            private string SourceFileName { get; }

            public PageState(PaperItem page)
            {
                Page = page;
                Rotation = page.Rotation;
                IsSelected = page.IsSelected;
                SourceFileName = page.SourceFileName;
            }

            public bool Matches(PaperItem page)
            {
                return ReferenceEquals(Page, page)
                    && Rotation == page.Rotation
                    && IsSelected == page.IsSelected
                    && SourceFileName == page.SourceFileName;
            }

            public void Restore()
            {
                Page.Rotation = Rotation;
                Page.IsSelected = IsSelected;
                Page.SourceFileName = SourceFileName;
            }
        }

        public static string SanitizeFileName(string? fileName, string defaultName = "PaperSwitch_裝訂成品.pdf")
        {
            if (string.IsNullOrWhiteSpace(fileName)) return defaultName;
            string clean = fileName.Trim();
            foreach (char invalid in Path.GetInvalidFileNameChars())
            {
                clean = clean.Replace(invalid, '_');
            }
            clean = clean.Trim('.', ' ', '_');
            return string.IsNullOrWhiteSpace(clean) ? defaultName : clean;
        }
    }
}
