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

        private readonly PdfService _pdfService = PdfService.Instance;
        private readonly OfficeConverterService _officeService = OfficeConverterService.Instance;
        private readonly ImageConverterService _imageService = ImageConverterService.Instance;
        private readonly ThumbnailCacheService _thumbnailService = ThumbnailCacheService.Instance;

        public int TotalPageCount => Pages.Count;
        public int SelectedPageCount => Pages.Count(p => p.IsSelected);
        public bool HasPages => Pages.Count > 0;
        public bool HasSelectedPages => SelectedPageCount > 0;
        public bool IsDraggingPages => DraggedPageCount > 0;

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

            foreach (var p in targetPages)
            {
                if (deltaDegrees > 0)
                    p.RotateClockwise();
                else
                    p.RotateCounterClockwise();
            }
            StatusMessage = $"已將 {targetPages.Count} 頁紙張旋轉 {deltaDegrees}°";
        }

        [RelayCommand]
        public void DeleteSelected()
        {
            var targetPages = Pages.Where(p => p.IsSelected).ToList();
            if (targetPages.Count == 0 && SelectedPage != null)
            {
                targetPages.Add(SelectedPage);
            }

            int nextSelectionIndex = targetPages.Count > 0
                ? Pages.IndexOf(targetPages[0])
                : 0;

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

            NotifySelectionChanged();
            StatusMessage = $"已移除 {targetPages.Count} 頁紙張";
        }

        [RelayCommand]
        public void ClearAll()
        {
            Pages.Clear();
            SelectedPage = null;
            _thumbnailService.ClearCache();
            NotifySelectionChanged();
            StatusMessage = "已清空排版畫布";
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

            int steps = Math.Abs(positions);
            for (int i = 0; i < steps; i++)
            {
                if (positions < 0)
                    MoveSelectedLeft();
                else
                    MoveSelectedRight();
            }
        }

        [RelayCommand]
        public void MoveSelectedToStart()
        {
            var selectedItems = Pages.Where(p => p.IsSelected).ToList();
            if (selectedItems.Count == 0) return;

            for (int i = 0; i < selectedItems.Count; i++)
            {
                var item = selectedItems[i];
                int curIdx = Pages.IndexOf(item);
                Pages.Move(curIdx, i);
            }
        }

        [RelayCommand]
        public void MoveSelectedToEnd()
        {
            var selectedItems = Pages.Where(p => p.IsSelected).ToList();
            if (selectedItems.Count == 0) return;

            int targetIndex = Pages.Count - 1;
            for (int i = selectedItems.Count - 1; i >= 0; i--)
            {
                var item = selectedItems[i];
                int curIdx = Pages.IndexOf(item);
                Pages.Move(curIdx, targetIndex);
            }
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

            NotifySelectionChanged();
        }

        #endregion

        #region 檔案載入與轉檔排程

        public async Task AddFilesAsync(string[] filePaths)
        {
            if (filePaths == null || filePaths.Length == 0) return;

            IsBusy = true;
            IsProgressIndeterminate = true;
            StatusMessage = "正在處理匯入之檔案...";

            string workTempDir = AppPaths.TemporaryConversionDirectory;

            var pdfsToLoad = new List<string>();

            try
            {
                foreach (var path in filePaths)
                {
                    if (!File.Exists(path)) continue;

                    string ext = Path.GetExtension(path).ToLowerInvariant();
                    StatusMessage = $"正在轉換: {Path.GetFileName(path)}...";

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
                }

                // 拆解 PDF 頁面並加入畫布
                StatusMessage = "正在載入頁面與生成縮圖...";
                var newItems = new List<PaperItem>();

                foreach (var pdf in pdfsToLoad)
                {
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
                }

                // 背景非同步載入縮圖
                _ = LoadThumbnailsAsync(newItems);

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
                NotifySelectionChanged();
            }
        }

        private async Task LoadThumbnailsAsync(List<PaperItem> items)
        {
            await Task.WhenAll(items.Select(LoadThumbnailAsync));
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

        #endregion
    }
}
