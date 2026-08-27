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

            item.RotateClockwise();
            Assert.Equal(90, item.Rotation);

            item.RotateClockwise();
            Assert.Equal(180, item.Rotation);

            item.RotateClockwise();
            Assert.Equal(270, item.Rotation);

            item.RotateClockwise();
            Assert.Equal(0, item.Rotation);

            item.RotateCounterClockwise();
            Assert.Equal(270, item.Rotation);
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
    }
}
