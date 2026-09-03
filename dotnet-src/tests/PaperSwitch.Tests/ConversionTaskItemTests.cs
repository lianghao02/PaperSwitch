using System;
using System.IO;
using System.Threading.Tasks;
using PaperSwitch.Models;
using PaperSwitch.ViewModels;
using Xunit;

namespace PaperSwitch.Tests
{
    public class ConversionTaskItemTests
    {
        [Fact]
        public void ConversionTaskItem_ShouldExposeClearStateLabels()
        {
            var task = new ConversionTaskItem(@"C:\測試\報告.docx");
            Assert.Equal(ConversionTaskState.Waiting, task.State);
            Assert.Equal("等待", task.StateLabel);
            Assert.Equal("○", task.StateIcon);

            task.MarkProcessing("正在轉換...");
            Assert.Equal("處理中", task.StateLabel);
            Assert.Equal("▶", task.StateIcon);

            task.MarkCompleted(3);
            Assert.Equal(ConversionTaskState.Completed, task.State);
            Assert.Contains("3 頁", task.StatusMessage);

            task.MarkFailed("請關閉 Office 後重新嘗試。", "COMException 0x800A");
            Assert.Equal("失敗", task.StateLabel);
            Assert.True(task.CanRetry);
            Assert.True(task.HasTechnicalDetails);

            task.MarkSkipped();
            Assert.Equal("跳過", task.StateLabel);
            Assert.False(task.CanRetry);
        }

        [Fact]
        public async Task AddFilesAsync_UnsupportedFile_ShouldBeSkippedWithoutChangingSource()
        {
            string path = Path.Combine(Path.GetTempPath(), $"paperswitch-unsupported-{Guid.NewGuid():N}.txt");
            const string content = "原始檔案不可被 PaperSwitch 修改";
            File.WriteAllText(path, content);

            try
            {
                var viewModel = new MainViewModel();
                await viewModel.AddFilesAsync(new[] { path });

                var task = Assert.Single(viewModel.ConversionTasks);
                Assert.Equal(ConversionTaskState.Skipped, task.State);
                Assert.Equal(content, File.ReadAllText(path));
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task AddFilesAsync_FailedPdf_ShouldNotBlockFollowingTask()
        {
            string directory = Path.Combine(Path.GetTempPath(), $"paperswitch-queue-{Guid.NewGuid():N}");
            Directory.CreateDirectory(directory);
            string brokenPdf = Path.Combine(directory, "損毀.pdf");
            string unsupportedFile = Path.Combine(directory, "保留.txt");
            File.WriteAllText(brokenPdf, "這不是標準 PDF");
            File.WriteAllText(unsupportedFile, "後續任務仍應被處理");

            try
            {
                var viewModel = new MainViewModel();
                await viewModel.AddFilesAsync(new[] { brokenPdf, unsupportedFile });

                Assert.Equal(2, viewModel.ConversionTasks.Count);
                Assert.Equal(ConversionTaskState.Failed, viewModel.ConversionTasks[0].State);
                Assert.Equal(ConversionTaskState.Skipped, viewModel.ConversionTasks[1].State);
                Assert.Equal("後續任務仍應被處理", File.ReadAllText(unsupportedFile));
            }
            finally
            {
                if (Directory.Exists(directory)) Directory.Delete(directory, true);
            }
        }
    }
}
