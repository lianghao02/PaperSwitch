using System.Linq;
using PaperSwitch.Models;
using PaperSwitch.ViewModels;
using Xunit;

namespace PaperSwitch.Tests
{
    public class ArrangerTests
    {
        [Fact]
        public void PaperItem_Rotation_ShouldCycleCorrectly()
        {
            var item = new PaperItem { Rotation = 0 };
            Assert.False(item.HasManualRotation);
            Assert.Equal(string.Empty, item.RotationDisplay);

            item.RotateClockwise();
            Assert.Equal(90, item.Rotation);
            Assert.True(item.HasManualRotation);
            Assert.Equal("90°", item.RotationDisplay);

            item.RotateClockwise();
            Assert.Equal(180, item.Rotation);
            Assert.True(item.HasManualRotation);

            item.RotateClockwise();
            Assert.Equal(270, item.Rotation);
            Assert.True(item.HasManualRotation);

            item.RotateClockwise();
            Assert.Equal(0, item.Rotation);
            Assert.False(item.HasManualRotation);
            Assert.Equal(string.Empty, item.RotationDisplay);

            item.RotateCounterClockwise();
            Assert.Equal(270, item.Rotation);
            Assert.True(item.HasManualRotation);
        }

        [Fact]
        public void MainViewModel_MoveSelectedLeft_ShouldSwapCorrectly()
        {
            var vm = new MainViewModel();
            var item0 = new PaperItem { DisplayPageNumber = 1 };
            var item1 = new PaperItem { DisplayPageNumber = 2, IsSelected = true };
            var item2 = new PaperItem { DisplayPageNumber = 3 };

            vm.Pages.Add(item0);
            vm.Pages.Add(item1);
            vm.Pages.Add(item2);

            vm.MoveSelectedLeft();

            Assert.Equal(item1, vm.Pages[0]);
            Assert.Equal(item0, vm.Pages[1]);
            Assert.Equal(item2, vm.Pages[2]);
        }

        [Fact]
        public void MainViewModel_MoveSelectedRight_ShouldSwapCorrectly()
        {
            var vm = new MainViewModel();
            var item0 = new PaperItem { DisplayPageNumber = 1, IsSelected = true };
            var item1 = new PaperItem { DisplayPageNumber = 2 };
            var item2 = new PaperItem { DisplayPageNumber = 3 };

            vm.Pages.Add(item0);
            vm.Pages.Add(item1);
            vm.Pages.Add(item2);

            vm.MoveSelectedRight();

            Assert.Equal(item1, vm.Pages[0]);
            Assert.Equal(item0, vm.Pages[1]);
            Assert.Equal(item2, vm.Pages[2]);
        }

        [Fact]
        public void MainViewModel_MoveSelectedBy_ShouldKeepMultiSelectionTogether()
        {
            var vm = new MainViewModel();
            var items = Enumerable.Range(1, 6).Select(number => new PaperItem
            {
                DisplayPageNumber = number,
                IsSelected = number is 3 or 4
            }).ToArray();

            foreach (var item in items)
            {
                vm.Pages.Add(item);
            }

            vm.MoveSelectedBy(-2);

            Assert.Equal(new[] { 3, 4, 1, 2, 5, 6 }, vm.Pages.Select(item => item.DisplayPageNumber));
            Assert.All(vm.Pages.Take(2), item => Assert.True(item.IsSelected));
        }

        [Fact]
        public void MainViewModel_RelocateSelectedChunk_ShouldInsertAtTargetPosition()
        {
            var vm = new MainViewModel();
            var item0 = new PaperItem { DisplayPageNumber = 1 };
            var item1 = new PaperItem { DisplayPageNumber = 2, IsSelected = true };
            var item2 = new PaperItem { DisplayPageNumber = 3, IsSelected = true };
            var item3 = new PaperItem { DisplayPageNumber = 4 };
            var item4 = new PaperItem { DisplayPageNumber = 5 };

            vm.Pages.Add(item0);
            vm.Pages.Add(item1);
            vm.Pages.Add(item2);
            vm.Pages.Add(item3);
            vm.Pages.Add(item4);

            // 將 item1, item2 整包移至最後 (索引 5)
            vm.RelocateSelectedChunk(5);

            Assert.Equal(item0, vm.Pages[0]);
            Assert.Equal(item3, vm.Pages[1]);
            Assert.Equal(item4, vm.Pages[2]);
            Assert.Equal(item1, vm.Pages[3]);
            Assert.Equal(item2, vm.Pages[4]);
        }

        [Fact]
        public void MainViewModel_SanitizeFileName_ShouldCleanInvalidChars()
        {
            Assert.Equal("PaperSwitch_裝訂成品.pdf", MainViewModel.SanitizeFileName(null));
            Assert.Equal("PaperSwitch_裝訂成品.pdf", MainViewModel.SanitizeFileName("   "));
            Assert.Equal("報告_2026", MainViewModel.SanitizeFileName("報告:2026"));
            Assert.Equal("公文_案號_123", MainViewModel.SanitizeFileName("公文/案號\\123"));
            Assert.Equal("合約書", MainViewModel.SanitizeFileName("合約書?*|<>"));
        }

        [Fact]
        public void MainViewModel_TryRenameSelectedPages_ShouldApplySequentialNames()
        {
            var vm = new MainViewModel();
            var first = new PaperItem { IsSelected = true, SourceFileName = "舊檔案.pdf" };
            var second = new PaperItem { IsSelected = true, SourceFileName = "另一個舊檔案.pdf" };
            vm.Pages.Add(first);
            vm.Pages.Add(second);

            bool renamed = vm.TryRenameSelectedPages("報告總表", out _);

            Assert.True(renamed);
            Assert.Equal("報告總表_001.pdf", first.SourceFileName);
            Assert.Equal("報告總表_002.pdf", second.SourceFileName);
        }

        [Fact]
        public void MainViewModel_TryRenameSelectedPages_ShouldRejectInvalidFileName()
        {
            var vm = new MainViewModel();
            vm.Pages.Add(new PaperItem { IsSelected = true });

            bool renamed = vm.TryRenameSelectedPages("不可使用:冒號", out string message);

            Assert.False(renamed);
            Assert.Contains("不可", message);
        }

        [Fact]
        public void MainViewModel_UndoRedo_ShouldRestoreReorderedPages()
        {
            var vm = new MainViewModel();
            var first = new PaperItem { DisplayPageNumber = 1 };
            var second = new PaperItem { DisplayPageNumber = 2, IsSelected = true };
            var third = new PaperItem { DisplayPageNumber = 3 };
            vm.Pages.Add(first);
            vm.Pages.Add(second);
            vm.Pages.Add(third);

            vm.MoveSelectedLeft();

            Assert.Equal(new[] { 2, 1, 3 }, vm.Pages.Select(page => page.DisplayPageNumber));
            Assert.True(vm.CanUndo);

            vm.Undo();

            Assert.Equal(new[] { 1, 2, 3 }, vm.Pages.Select(page => page.DisplayPageNumber));
            Assert.True(second.IsSelected);
            Assert.True(vm.CanRedo);

            vm.Redo();

            Assert.Equal(new[] { 2, 1, 3 }, vm.Pages.Select(page => page.DisplayPageNumber));
        }

        [Fact]
        public void MainViewModel_Undo_ShouldRestoreDeletedPageRotationAndExportName()
        {
            var vm = new MainViewModel();
            var page = new PaperItem { IsSelected = true, SourceFileName = "原始檔案.pdf" };
            vm.Pages.Add(page);

            vm.RotateSelected(90);
            vm.TryRenameSelectedPages("整理後檔案", out _);
            vm.DeleteSelected();

            Assert.Empty(vm.Pages);

            vm.Undo();

            Assert.Single(vm.Pages);
            Assert.Equal("整理後檔案.pdf", page.SourceFileName);
            Assert.Equal(90, page.Rotation);
            Assert.True(page.IsSelected);
        }

        [Fact]
        public void MainViewModel_DeselectAll_ShouldClearSelectedPageAndItemStates()
        {
            var vm = new MainViewModel();
            var first = new PaperItem { IsSelected = true };
            var second = new PaperItem { IsSelected = true };
            vm.Pages.Add(first);
            vm.Pages.Add(second);
            vm.SelectedPage = first;

            vm.DeselectAll();

            Assert.Null(vm.SelectedPage);
            Assert.All(vm.Pages, page => Assert.False(page.IsSelected));
        }

        [Fact]
        public void MainViewModel_ExportCommands_CanExecute_ReflectsHasPages()
        {
            var vm = new MainViewModel();

            // 空畫布時，全部匯出、另存選取與拆分選取皆不可執行
            Assert.False(vm.HasPages);
            Assert.False(vm.ExportAllPdfCommand.CanExecute(null));
            Assert.False(vm.ExportSelectedPdfCommand.CanExecute(null));
            Assert.False(vm.ExportSplitPdfCommand.CanExecute(null));

            var page1 = new PaperItem { DisplayPageNumber = 1, IsSelected = false };
            var page2 = new PaperItem { DisplayPageNumber = 2, IsSelected = false };
            vm.Pages.Add(page1);
            vm.Pages.Add(page2);

            // 有頁面但無選取時：三顆按鈕皆可執行（預設處理全部頁面）
            Assert.True(vm.HasPages);
            Assert.False(vm.HasSelectedPages);
            Assert.True(vm.ExportAllPdfCommand.CanExecute(null));
            Assert.True(vm.ExportSelectedPdfCommand.CanExecute(null));
            Assert.True(vm.ExportSplitPdfCommand.CanExecute(null));

            // 有選取時：三顆按鈕亦皆可執行
            page2.IsSelected = true;
            vm.NotifySelectionChanged();
            Assert.True(vm.HasSelectedPages);
            Assert.True(vm.ExportAllPdfCommand.CanExecute(null));
            Assert.True(vm.ExportSelectedPdfCommand.CanExecute(null));
            Assert.True(vm.ExportSplitPdfCommand.CanExecute(null));

            // 清空所有選取：三顆按鈕仍維持可執行（只要 HasPages 為 True）
            vm.DeselectAll();
            Assert.False(vm.HasSelectedPages);
            Assert.True(vm.ExportSelectedPdfCommand.CanExecute(null));
            Assert.True(vm.ExportSplitPdfCommand.CanExecute(null));

            // 清空畫布：按鈕全部停用
            vm.Pages.Clear();
            Assert.False(vm.HasPages);
            Assert.False(vm.ExportAllPdfCommand.CanExecute(null));
            Assert.False(vm.ExportSelectedPdfCommand.CanExecute(null));
            Assert.False(vm.ExportSplitPdfCommand.CanExecute(null));
        }

        [Fact]
        public void MainViewModel_GetEffectiveExportPages_NoSelection_ReturnsAllPagesInCanvasOrder()
        {
            var vm = new MainViewModel();
            for (int i = 1; i <= 5; i++)
            {
                vm.Pages.Add(new PaperItem { DisplayPageNumber = i, IsSelected = false });
            }

            var effective = vm.GetEffectiveExportPages();
            Assert.Equal(5, effective.Count);
            Assert.Equal(new[] { 1, 2, 3, 4, 5 }, effective.Select(p => p.DisplayPageNumber));
        }

        [Fact]
        public void MainViewModel_GetEffectiveExportPages_WithSelection_ReturnsOnlySelectedInCanvasOrder()
        {
            var vm = new MainViewModel();
            for (int i = 1; i <= 5; i++)
            {
                vm.Pages.Add(new PaperItem { DisplayPageNumber = i, IsSelected = false });
            }

            // 選取第 2、5 頁
            vm.Pages[1].IsSelected = true;
            vm.Pages[4].IsSelected = true;
            vm.NotifySelectionChanged();

            var effective = vm.GetEffectiveExportPages();
            Assert.Equal(2, effective.Count);
            Assert.Equal(new[] { 2, 5 }, effective.Select(p => p.DisplayPageNumber));
        }

        [Fact]
        public void MainViewModel_GetEffectiveExportPages_ReorderedAndRotated_PreservesCanvasState()
        {
            var vm = new MainViewModel();
            var pages = Enumerable.Range(1, 7).Select(num => new PaperItem
            {
                DisplayPageNumber = num,
                Rotation = 0,
                IsSelected = false
            }).ToList();

            foreach (var p in pages)
            {
                vm.Pages.Add(p);
            }

            // 選取第 2, 5, 7 頁，並為第 5 頁旋轉 90 度
            vm.Pages[1].IsSelected = true;
            vm.Pages[4].IsSelected = true;
            vm.Pages[6].IsSelected = true;
            vm.Pages[4].Rotation = 90;
            vm.NotifySelectionChanged();

            var effective = vm.GetEffectiveExportPages();
            Assert.Equal(3, effective.Count);
            Assert.Equal(new[] { 2, 5, 7 }, effective.Select(p => p.DisplayPageNumber));
            Assert.Equal(90, effective[1].Rotation);

            // 重新排列：將第 1 頁移至最後（索引 0 移到最後，原本選取的 2, 5, 7 順序仍維持）
            var first = vm.Pages[0];
            vm.Pages.RemoveAt(0);
            vm.Pages.Add(first);

            var reorderedEffective = vm.GetEffectiveExportPages();
            Assert.Equal(3, reorderedEffective.Count);
            Assert.Equal(new[] { 2, 5, 7 }, reorderedEffective.Select(p => p.DisplayPageNumber));
            Assert.Equal(90, reorderedEffective[1].Rotation);
        }
    }
}
