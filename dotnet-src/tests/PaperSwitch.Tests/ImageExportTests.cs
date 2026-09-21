using System.IO;
using System.Linq;
using PaperSwitch.Models;
using PaperSwitch.Services;
using PaperSwitch.ViewModels;
using Xunit;

namespace PaperSwitch.Tests
{
    public class ImageExportTests
    {
        [Fact]
        public void BuildPngFileName_WithoutCustomPrefix_GeneratesSourceAndCanvasOrder()
        {
            string fileName = ImageExportService.BuildPngFileName("Document.pdf", 3, 1, null);
            Assert.Equal("Document_頁3_001.png", fileName);
        }

        [Fact]
        public void BuildPngFileName_WithCustomPrefix_UsesPrefixAndCanvasOrder()
        {
            string fileName = ImageExportService.BuildPngFileName("Document.pdf", 3, 2, "會議簡報");
            Assert.Equal("會議簡報_002.png", fileName);
        }

        [Fact]
        public void BuildPngFileName_WithInvalidChars_SanitizesCorrectly()
        {
            string fileNameWithPrefix = ImageExportService.BuildPngFileName("DocTest.pdf", 1, 5, "前綴/?:*A");
            Assert.Equal("前綴____A_005.png", fileNameWithPrefix);

            string fileNameWithoutPrefix = ImageExportService.BuildPngFileName("Doc/?:*Test.pdf", 2, 3, null);
            Assert.Equal("Doc____Test_頁2_003.png", fileNameWithoutPrefix);
        }

        [Fact]
        public void GetUniqueFilePath_WhenFileDoesNotExist_ReturnsOriginalPath()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "PaperSwitch_Test_" + Path.GetRandomFileName());
            Directory.CreateDirectory(tempDir);
            try
            {
                string target = ImageExportService.GetUniqueFilePath(tempDir, "sample_001.png");
                Assert.Equal(Path.Combine(tempDir, "sample_001.png"), target);
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void GetUniqueFilePath_WhenFileExists_AddsCounterSuffix()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "PaperSwitch_Test_" + Path.GetRandomFileName());
            Directory.CreateDirectory(tempDir);
            try
            {
                string originalPath = Path.Combine(tempDir, "sample_001.png");
                File.WriteAllText(originalPath, "dummy");

                string target = ImageExportService.GetUniqueFilePath(tempDir, "sample_001.png");
                Assert.Equal(Path.Combine(tempDir, "sample_001_1.png"), target);

                File.WriteAllText(target, "dummy2");
                string target2 = ImageExportService.GetUniqueFilePath(tempDir, "sample_001.png");
                Assert.Equal(Path.Combine(tempDir, "sample_001_2.png"), target2);
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void ImageExport_SelectionBehavior_MatchesGetEffectiveExportPages()
        {
            var vm = new MainViewModel();
            for (int i = 1; i <= 6; i++)
            {
                vm.Pages.Add(new PaperItem
                {
                    DisplayPageNumber = i,
                    SourceFileName = $"Doc{i}.pdf",
                    IsSelected = false
                });
            }

            // 無選取：應處理全數 6 頁
            var allPages = vm.GetEffectiveExportPages();
            Assert.Equal(6, allPages.Count);
            Assert.Equal(new[] { 1, 2, 3, 4, 5, 6 }, allPages.Select(p => p.DisplayPageNumber));

            // 單選第 3 頁
            vm.Pages[2].IsSelected = true;
            vm.NotifySelectionChanged();
            var singleSelected = vm.GetEffectiveExportPages();
            Assert.Single(singleSelected);
            Assert.Equal(3, singleSelected[0].DisplayPageNumber);

            // 多選第 3、5 頁，並模擬重排與旋轉
            vm.Pages[4].IsSelected = true;
            vm.Pages[4].Rotation = 90;
            vm.NotifySelectionChanged();

            // 調整順序：第 5 頁（索引 4）移到最前面
            var page5 = vm.Pages[4];
            vm.Pages.RemoveAt(4);
            vm.Pages.Insert(0, page5);

            var multiSelected = vm.GetEffectiveExportPages();
            Assert.Equal(2, multiSelected.Count);
            // 畫布順序應先第 5 頁，後第 3 頁
            Assert.Equal(5, multiSelected[0].DisplayPageNumber);
            Assert.Equal(90, multiSelected[0].Rotation);
            Assert.Equal(3, multiSelected[1].DisplayPageNumber);

            // 移除已選取之第 5 頁，畫布僅剩第 3 頁被選取
            vm.Pages.Remove(page5);
            var afterRemoved = vm.GetEffectiveExportPages();
            Assert.Single(afterRemoved);
            Assert.Equal(3, afterRemoved[0].DisplayPageNumber);
        }
    }
}
