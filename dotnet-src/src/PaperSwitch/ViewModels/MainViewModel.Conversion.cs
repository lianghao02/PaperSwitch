using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using CommunityToolkit.Mvvm.Input;
using PaperSwitch.Models;
using PaperSwitch.Services;

namespace PaperSwitch.ViewModels
{
    /// <summary>
    /// 主介面 ViewModel：檔案匯入、Office 轉檔佇列、加密續轉引導與大頁數縮圖視窗排程。
    /// </summary>
    public partial class MainViewModel
    {
        #region 檔案載入與轉檔排程

        public async Task AddFilesAsync(string[] filePaths)
        {
            if (IsBusy) return;
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
                _operationCancellation = new();
                IsProgressIndeterminate = false;
                ProgressValue = 0;
                StatusMessage = $"準備處理 {tasks.Count} 個檔案...";
                var beforeImport = CaptureArrangementSnapshot();

                string workTempDir = AppPaths.TemporaryConversionDirectory;
                var pdfsToLoad = new List<(string PdfPath, ConversionTaskItem Task)>();

                for (int fileIndex = 0; fileIndex < tasks.Count; fileIndex++)
                {
                    var task = tasks[fileIndex];
                    if (_operationCancellation?.IsCancellationRequested == true)
                    {
                        task.MarkSkipped("已停止，未開始轉檔。");
                        continue;
                    }
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

                if (newItems.Count > 0)
                {
                    RecordUndoSnapshot(beforeImport);
                    var initialBatch = Pages.Count > LargeCanvasThumbnailWindowThreshold
                        ? newItems.Take(60).ToList()
                        : newItems;
                    _ = LoadThumbnailsAsync(initialBatch);
                }

                ProgressValue = 100;
                int failedCount = tasks.Count(task => task.State == ConversionTaskState.Failed);
                StatusMessage = _operationCancellation?.IsCancellationRequested == true ? $"已停止；本輪載入 {newItems.Count} 頁，未開始的任務已跳過。" : failedCount == 0
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
                _operationCancellation?.Dispose();
                _operationCancellation = null;
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
                    "轉出的 PDF 被公司加密程序鎖住，需要先解密。原始檔案沒有被修改。",
                    "Office 暫存 PDF 未在可等待期限內成為標準 %PDF 檔案。",
                    pdfPath,
                    needsDecrypt: true);
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
                var pageData = await Task.Run(() => PdfImportService.ReadPages(pdfPath, task.FileName));
                int totalPages = pageData.Count;
                if (totalPages <= 0)
                {
                    task.MarkFailed("PDF 無法讀取或內容已損毀。原始檔案沒有被修改，請確認檔案後重新嘗試。", actionPath: pdfPath);
                    return;
                }

                _suppressPagesCollectionNotifications = true;
                try
                {
                    for (int index = 0; index < pageData.Count; index++)
                    {
                        newItems.Add(pageData[index]);
                        Pages.Add(pageData[index]);
                        if (index % 64 == 63) await System.Windows.Threading.Dispatcher.Yield();
                    }
                }
                finally
                {
                    _suppressPagesCollectionNotifications = false;
                }
                NotifyPagesStructureChanged();

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

        /// <summary>
        /// 暫存 PDF 被公司加密程序鎖住時的處理說明：可反覆開啟資料夾，解密後按「繼續」。
        /// </summary>
        [RelayCommand]
        public async Task ShowDecryptGuideAsync(ConversionTaskItem? task)
        {
            if (task is null || !task.ShowDecryptGuide || IsBusy) return;

            while (true)
            {
                string fileName = Path.GetFileName(task.ActionPath ?? task.SourcePath);
                var choice = GameDialog.ShowChoice(
                    $"「{task.FileName}」已轉成 PDF，但被公司的加密程序鎖住，目前無法讀取。\n\n" +
                    "請依序處理：\n" +
                    "1. 按「開啟位置」，找到被選取的 PDF（" + fileName + "）。\n" +
                    "2. 用公司規定的方式把它解密，檔名與位置請維持不變。\n" +
                    "3. 回到這裡按「我已解密，繼續」。\n\n" +
                    "原始 Office 檔案不會被修改。",
                    "需要先解密",
                    primaryText: "我已解密，繼續",
                    secondaryText: "開啟位置",
                    cancelText: "稍後再說");

                if (choice == MessageBoxResult.No)
                {
                    OpenConversionTaskLocation(task);
                    continue;
                }

                if (choice == MessageBoxResult.Yes)
                {
                    await ContinueAfterDecryptAsync(task);
                    if (task.ShowDecryptGuide) continue;
                }

                return;
            }
        }

        private async Task ContinueAfterDecryptAsync(ConversionTaskItem task)
        {
            string path = task.ActionPath ?? task.SourcePath;
            if (!File.Exists(path))
            {
                GameDialog.Show("找不到這份 PDF，可能已被移動或清除。請按「重新嘗試」重新轉檔。", "找不到檔案", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_officeService.CheckPdfHeaderAndSize(path).State != "pdf")
            {
                GameDialog.Show("這份 PDF 還是加密狀態，無法讀取。\n請確認已解密，且檔名與位置沒有改變。", "還不能讀取", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            await _conversionBatchGate.WaitAsync();
            try
            {
                IsBusy = true;
                IsProgressIndeterminate = true;
                task.MarkProcessing("讀取已解密的 PDF...");
                StatusMessage = $"正在讀取已解密的 PDF：\n{task.FileName}";
                var beforeImport = CaptureArrangementSnapshot();

                var pdfsToLoad = new List<(string PdfPath, ConversionTaskItem Task)>();
                RegisterPdfForLoading(task, path, pdfsToLoad);

                var newItems = new List<PaperItem>();
                foreach (var (pdf, owner) in pdfsToLoad)
                {
                    await LoadPdfPagesForTaskAsync(pdf, owner, newItems);
                }

                if (task.State == ConversionTaskState.Processing && task.PendingPdfCount == 0)
                {
                    task.MarkCompleted(task.CompletedPageCount);
                }

                if (newItems.Count > 0)
                {
                    RecordUndoSnapshot(beforeImport);
                    var initialBatch = Pages.Count > LargeCanvasThumbnailWindowThreshold
                        ? newItems.Take(60).ToList()
                        : newItems;
                    _ = LoadThumbnailsAsync(initialBatch);
                }
                StatusMessage = $"已載入 {newItems.Count} 頁新紙張";
            }
            catch (Exception ex)
            {
                task.MarkFailed("讀取解密後的 PDF 時發生錯誤。原始檔案沒有被修改，請確認後重新嘗試。", ex.ToString());
            }
            finally
            {
                IsBusy = false;
                NotifySelectionChanged();
                NotifyConversionQueueChanged();
                _conversionBatchGate.Release();
            }
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

        private async Task LoadThumbnailsAsync(IReadOnlyList<PaperItem> items, CancellationToken cancellationToken = default)
        {
            var pending = items
                .Where(item => item.Thumbnail == null && !item.HasError && !string.IsNullOrEmpty(item.SourceFilePath))
                .ToList();
            if (pending.Count == 0) return;

            var groups = pending.GroupBy(item => item.SourceFilePath, StringComparer.OrdinalIgnoreCase);
            var tasks = groups.Select(async group =>
            {
                if (cancellationToken.IsCancellationRequested) return;
                var groupItems = group.ToList();
                var indices = groupItems.Select(item => item.SourcePageIndex).Distinct().ToList();

                await _thumbnailRenderGate.WaitAsync(cancellationToken);
                try
                {
                    if (cancellationToken.IsCancellationRequested) return;
                    var rendered = await _thumbnailService.GetThumbnailsForPdfAsync(group.Key, indices, 360, cancellationToken);
                    foreach (var item in groupItems)
                    {
                        if (cancellationToken.IsCancellationRequested) break;
                        if (rendered.TryGetValue(item.SourcePageIndex, out var thumb))
                        {
                            item.Thumbnail = thumb;
                        }
                        item.IsLoadingThumbnail = false;
                    }
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch
                {
                    foreach (var item in groupItems)
                    {
                        item.IsLoadingThumbnail = false;
                        item.HasError = true;
                    }
                }
                finally
                {
                    _thumbnailRenderGate.Release();
                }
            });

            try { await Task.WhenAll(tasks); }
            catch (OperationCanceledException) { }
        }

        public async Task UpdateVisibleThumbnailWindowAsync(IReadOnlyList<PaperItem> visibleAndBufferPages, IReadOnlySet<PaperItem> retainSet, CancellationToken cancellationToken = default)
        {
            if (Pages.Count > LargeCanvasThumbnailWindowThreshold)
            {
                foreach (var page in Pages)
                {
                    if (page.Thumbnail != null && !retainSet.Contains(page) && !string.IsNullOrEmpty(page.SourceFilePath))
                    {
                        page.Thumbnail = null;
                        page.IsLoadingThumbnail = true;
                    }
                }
            }

            await LoadThumbnailsAsync(visibleAndBufferPages, cancellationToken);
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
    }
}
