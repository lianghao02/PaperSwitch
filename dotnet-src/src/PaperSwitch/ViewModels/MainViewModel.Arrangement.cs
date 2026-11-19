using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using CommunityToolkit.Mvvm.Input;
using PaperSwitch.Models;
using PaperSwitch.Services;

namespace PaperSwitch.ViewModels
{
    /// <summary>
    /// 主介面 ViewModel：紙張編排、選取、旋轉、插入空白頁與復原／重做歷史管理。
    /// </summary>
    public partial class MainViewModel
    {
        private bool ApplyPageEdit(Action editAction)
        {
            if (IsBusy) return false;
            var before = CaptureArrangementSnapshot();
            editAction();

            if (before.Matches(Pages, SelectedPage))
            {
                return false;
            }

            _history.Undo.Push(before);
            _history.Redo.Clear();
            NotifyHistoryChanged();
            return true;
        }

        private ArrangementSnapshot CaptureArrangementSnapshot()
        {
            return new ArrangementSnapshot(Pages, SelectedPage);
        }

        private void RestoreArrangementSnapshot(ArrangementSnapshot snapshot)
        {
            _suppressPagesCollectionNotifications = true;
            try
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
            }
            finally
            {
                _suppressPagesCollectionNotifications = false;
            }

            NotifyPagesStructureChanged();
        }

        private void NotifyHistoryChanged()
        {
            OnPropertyChanged(nameof(ExportFileNamePreview));
            OnPropertyChanged(nameof(CanUndo));
            OnPropertyChanged(nameof(CanRedo));
        }

        private void RecordUndoSnapshot(ArrangementSnapshot beforeSnapshot)
        {
            if (beforeSnapshot.Matches(Pages, SelectedPage)) return;
            _history.Undo.Push(beforeSnapshot);
            _history.Redo.Clear();
            NotifyHistoryChanged();
        }

        [RelayCommand]
        public void Undo()
        {
            if (!CanUndo) return;

            _history.Redo.Push(CaptureArrangementSnapshot());
            RestoreArrangementSnapshot(_history.Undo.Pop());
            NotifyHistoryChanged();
            StatusMessage = "已復原上一個編排動作";
        }

        [RelayCommand]
        public void Redo()
        {
            if (!CanRedo) return;

            _history.Undo.Push(CaptureArrangementSnapshot());
            RestoreArrangementSnapshot(_history.Redo.Pop());
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

        public bool TrySelectRange(string text, out string error)
        {
            error = "";
            if (IsBusy) { error = "請等待目前工作完成。"; return false; }
            try
            {
                var indexes = PageRangeSelection.Parse(text, Pages.Count);
                for (int i = 0; i < Pages.Count; i++) Pages[i].IsSelected = indexes.Contains(i);
                SelectedPage = Pages.FirstOrDefault(p => p.IsSelected);
                NotifySelectionChanged();
                return true;
            }
            catch (FormatException ex) { error = ex.Message; return false; }
        }

        [RelayCommand]
        public void SelectSource()
        {
            if (IsBusy || SelectedPage == null) return;
            var source = SelectedPage.SourceFilePath;
            foreach (var page in Pages) page.IsSelected = string.Equals(page.SourceFilePath, source, StringComparison.OrdinalIgnoreCase);
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

        public void RotatePage(PaperItem page, int delta)
        {
            if (!Pages.Contains(page)) return;
            ApplyPageEdit(() => page.Rotation = (page.Rotation + delta + 360) % 360);
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
                GameDialog.Show("建立空白頁失敗，請確認暫存資料夾是否可寫入。", "插入空白頁", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var item = new PaperItem
            {
                SourceFilePath = blankPath,
                SourceFileName = "空白頁.pdf",
                IsGeneratedBlank = true,
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
            if (IsBusy) return;
            var selectedItems = Pages.Where(p => p.IsSelected).ToList();
            if (selectedItems.Count == 0) return;

            if (targetInsertionIndex < 0) targetInsertionIndex = 0;
            if (targetInsertionIndex > Pages.Count) targetInsertionIndex = Pages.Count;

            // 先以未抽出前的位置扣除前方選取頁，避免錨點被移除後誤落到最後。
            int newInsertPos = Pages.Take(targetInsertionIndex).Count(p => !p.IsSelected);
            var reordered = Pages.Where(p => !p.IsSelected).ToList();
            reordered.InsertRange(newInsertPos, selectedItems);
            if (Pages.SequenceEqual(reordered)) return;

            ApplyPageEdit(() =>
            {
                // 移除所有選取項目
                foreach (var item in selectedItems)
                {
                    Pages.Remove(item);
                }

                for (int i = 0; i < selectedItems.Count; i++)
                {
                    Pages.Insert(newInsertPos + i, selectedItems[i]);
                }
            });

            NotifySelectionChanged();
        }

        #endregion
    }
}
