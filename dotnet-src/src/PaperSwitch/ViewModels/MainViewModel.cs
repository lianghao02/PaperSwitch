using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
        private const int MaximumConcurrentThumbnailRenders = 4;

        public int TotalPageCount => Pages.Count;
        public int SelectedPageCount => Pages.Count(p => p.IsSelected);
        public bool HasPages => Pages.Count > 0;
        public bool HasSelectedPages => SelectedPageCount > 0;
        public bool IsDraggingPages => DraggedPageCount > 0;
        public bool CanUndo => _undoHistory.Count > 0;
        public bool CanRedo => _redoHistory.Count > 0;

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
            };
        }

        public void NotifySelectionChanged()
        {
            OnPropertyChanged(nameof(SelectedPageCount));
            OnPropertyChanged(nameof(HasSelectedPages));
            OnPropertyChanged(nameof(SummaryText));
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

            IsBusy = true;
            IsProgressIndeterminate = false;
            ProgressValue = 0;
            StatusMessage = $"準備匯入 {existingFiles.Length} 個檔案...";

            string workTempDir = AppPaths.TemporaryConversionDirectory;

            var pdfsToLoad = new List<string>();

            try
            {
                for (int fileIndex = 0; fileIndex < existingFiles.Length; fileIndex++)
                {
                    string path = existingFiles[fileIndex];
                    int displayIndex = fileIndex + 1;

                    string ext = Path.GetExtension(path).ToLowerInvariant();
                    StatusMessage = $"正在轉換第 {displayIndex} / {existingFiles.Length} 個檔案：\n{Path.GetFileName(path)}";

                    if (ext == ".pdf")
                    {
                        pdfsToLoad.Add(path);
                    }
                    else if (_officeService.IsWordFile(path))
                    {
                        string outPdf = Path.Combine(workTempDir, $"{Path.GetFileNameWithoutExtension(path)}_{Guid.NewGuid():N}.pdf");
                        var result = await _officeService.ConvertWordAsync(path, outPdf);
                        if (result.Success)
                        {
                            await AddReadyOfficePdfAsync(outPdf, pdfsToLoad);
                        }
                        else
                        {
                            MessageBox.Show(result.ErrorMessage, "Word 轉檔提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                    else if (_officeService.IsExcelFile(path))
                    {
                        var result = await _officeService.ConvertExcelAsync(path, workTempDir, Options.ExcelFitToPage);
                        if (result.Success)
                        {
                            foreach (var p in result.GeneratedPdfPaths)
                            {
                                await AddReadyOfficePdfAsync(p, pdfsToLoad);
                            }
                        }
                        else
                        {
                            MessageBox.Show(result.ErrorMessage, "Excel 轉檔提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                    else if (_officeService.IsPowerPointFile(path))
                    {
                        string outPdf = Path.Combine(workTempDir, $"{Path.GetFileNameWithoutExtension(path)}_{Guid.NewGuid():N}.pdf");
                        var result = await _officeService.ConvertPowerPointAsync(path, outPdf);
                        if (result.Success)
                        {
                            await AddReadyOfficePdfAsync(outPdf, pdfsToLoad);
                        }
                        else
                        {
                            MessageBox.Show(result.ErrorMessage, "PowerPoint 轉檔提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                    else if (_imageService.IsImageFile(path))
                    {
                        string outPdf = Path.Combine(workTempDir, $"{Path.GetFileNameWithoutExtension(path)}_{Guid.NewGuid():N}.pdf");
                        if (_imageService.ConvertImagesToPdf(new[] { path }, outPdf))
                        {
                            pdfsToLoad.Add(outPdf);
                        }
                    }

                    ProgressValue = displayIndex * 70.0 / existingFiles.Length;
                }

                // 拆解 PDF 頁面並加入畫布
                ProgressValue = 75;
                StatusMessage = $"檔案轉換完成，正在載入 {pdfsToLoad.Count} 份 PDF 的頁面與縮圖...";
                var newItems = new List<PaperItem>();

                for (int pdfIndex = 0; pdfIndex < pdfsToLoad.Count; pdfIndex++)
                {
                    string pdf = pdfsToLoad[pdfIndex];
                    StatusMessage = $"正在讀取第 {pdfIndex + 1} / {pdfsToLoad.Count} 份 PDF：\n{Path.GetFileName(pdf)}";
                    int totalPages = _pdfService.GetPageCount(pdf);
                    string fileName = Path.GetFileName(pdf);

                    for (int pageIdx = 0; pageIdx < totalPages; pageIdx++)
                    {
                        var (w, h, rot) = _pdfService.GetPageDimensions(pdf, pageIdx);
                        var item = new PaperItem
                        {
                            SourceFilePath = pdf,
                            SourceFileName = fileName,
                            SourcePageIndex = pageIdx,
                            DisplayPageNumber = pageIdx + 1,
                            TotalPagesInSource = totalPages,
                            Rotation = rot,
                            OriginalWidth = w,
                            OriginalHeight = h,
                            IsLoadingThumbnail = true
                        };
                        newItems.Add(item);
                        Pages.Add(item);
                    }

                    ProgressValue = 75 + ((pdfIndex + 1) * 20.0 / Math.Max(pdfsToLoad.Count, 1));
                }

                // 背景非同步載入縮圖
                _ = LoadThumbnailsAsync(newItems);

                // 匯入會加入新的外部來源，作為新的手動編排歷史起點，避免復原誤移除新匯入紙張。
                if (newItems.Count > 0)
                {
                    ClearHistory();
                }

                ProgressValue = 100;
                StatusMessage = $"已順利裝載 {newItems.Count} 頁新紙張至工坊畫布";
            }
            catch (Exception ex)
            {
                StatusMessage = $"處理檔案時發生例外: {ex.Message}";
                MessageBox.Show($"載入檔案失敗: {ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
                IsProgressIndeterminate = true;
                NotifySelectionChanged();
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

        private async Task AddReadyOfficePdfAsync(string pdfPath, ICollection<string> pdfsToLoad)
        {
            if (await _officeService.WaitForPdfReadyAsync(pdfPath))
            {
                pdfsToLoad.Add(pdfPath);
                return;
            }

            StatusMessage = "暫存 PDF 尚未完成，已保留供解密後重新匯入";
            var choice = MessageBox.Show(
                $"已產生暫存 PDF，但尚非可處理的標準 PDF。\n\n" +
                "可能正受公務端檔案加密程序處理；請依規定完成解密後，再將該 PDF 重新加入工坊。\n\n" +
                $"保留位置：{pdfPath}\n\n是否現在開啟所在資料夾？",
                "Office 文件轉檔提醒",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (choice == MessageBoxResult.Yes)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"/select,\"{pdfPath}\"",
                    UseShellExecute = true
                });
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
                    string finalName = string.IsNullOrWhiteSpace(Options.CustomFileName) ? "PaperSwitch_裝訂成品.pdf" : Options.CustomFileName;
                    if (!finalName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) finalName += ".pdf";

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
                    var files = await Task.Run(() => _pdfService.ExportIndividualPdfs(Pages, outputDir, Options.CustomFileName));
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
                                  "是否立即開啟瀏覽器前往下載更新？";

                    var dialogResult = MessageBox.Show(
                        message,
                        $"🎉 發現新版本 {result.LatestVersion}",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (dialogResult == MessageBoxResult.Yes && !string.IsNullOrWhiteSpace(result.HtmlUrl))
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = result.HtmlUrl,
                            UseShellExecute = true
                        });
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
    }
}
