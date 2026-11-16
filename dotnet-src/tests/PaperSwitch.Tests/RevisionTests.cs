using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text.Json;
using PaperSwitch.Models;
using PaperSwitch.Services;
using PaperSwitch.ViewModels;
using PdfSharp.Pdf;
using Xunit;

namespace PaperSwitch.Tests;

public sealed class RevisionTests
{
    private sealed class ImmediateProgress(Action callback) : IProgress<(int Current, int Total, string FileName)>
    {
        public void Report((int Current, int Total, string FileName) value) => callback();
    }

    [Fact]
    public async Task SplitCancellationPreservesCompletedFilesAndReportsRemainingPages()
    {
        var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var source = Path.Combine(dir, "source.pdf");
            Assert.True(PdfService.Instance.CreateBlankA4Pdf(source));
            var page = PdfImportService.ReadPages(source, "來源")[0];
            using var cancellation = new CancellationTokenSource();
            var result = await new ExportCoordinator().ExportAsync(ExportKind.SplitPdf,
                new[] { page, page, page }, new ExportOptions { OutputDirectory = dir }, cancellation.Token,
                new ImmediateProgress(() => cancellation.Cancel()));
            Assert.True(result.Cancelled);
            Assert.Single(result.Files);
            Assert.Equal(1, result.CompletedPages);
            Assert.Equal(1, PdfService.Instance.GetPageCount(result.Files[0]));
            Assert.Empty(Directory.GetFiles(dir, "*.partial"));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public async Task MissingPageMustRejectWholeMergeAndPreserveExistingOutput()
    {
        var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var source = Path.Combine(dir, "source.pdf");
            Assert.True(PdfService.Instance.CreateBlankA4Pdf(source));
            var pages = PdfImportService.ReadPages(source, "source.pdf");
            pages.Add(new PaperItem { SourceFilePath = Path.Combine(dir, "missing.pdf") });
            var output = Path.Combine(dir, "output.pdf");
            Assert.False(PdfService.Instance.ExportArrangedPdf(pages, output));
            Assert.False(File.Exists(output));
            var result = await new ExportCoordinator().ExportAsync(ExportKind.AllPdf, pages,
                new ExportOptions { OutputDirectory = dir }, default);
            Assert.NotEmpty(result.Errors);
            Assert.Empty(result.Files);
            File.WriteAllText(output, "既有成品");
            Assert.False(PdfService.Instance.ExportArrangedPdf(pages.Take(1), output));
            Assert.Equal("既有成品", File.ReadAllText(output));
            Assert.Empty(Directory.GetFiles(dir, "*.partial"));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void PreviewRotationHistoryAndCleanupPreserveRestorableSource()
    {
        var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var path = Path.Combine(dir, "source.pdf");
            Assert.True(PdfService.Instance.CreateBlankA4Pdf(path));
            var vm = new MainViewModel(new ExportPreferencesStore(Path.Combine(dir, "preferences.json")));
            var page = PdfImportService.ReadPages(path, "來源")[0];
            page.IsSelected = true;
            vm.Pages.Add(page);
            vm.SelectedPage = page;
            var preview = new LightboxViewModel(page, vm.Pages, vm.RotatePage);
            preview.RotateClockwise();
            Assert.Equal(90, page.Rotation);
            vm.Undo(); Assert.Equal(0, page.Rotation);
            vm.Redo(); Assert.Equal(90, page.Rotation);
            vm.DeleteSelected();
            StorageMaintenanceService.ClearDirectoryExcept(dir, vm.GetActiveSourcePaths());
            Assert.True(File.Exists(path));
            vm.Undo();
            Assert.True(PdfService.Instance.ExportArrangedPdf(vm.Pages, Path.Combine(dir, "restored.pdf")));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Theory]
    [InlineData("1-3,3,5", true)]
    [InlineData("1–3，5", true)]
    [InlineData("1－3、5", true)]
    [InlineData(" 1~3 , 5 ", true)]
    [InlineData("0,1", false)]
    [InlineData("2-8", false)]
    [InlineData("4-2", false)]
    [InlineData("1,", false)]
    public void RangeSelectionRejectsInvalidInputWithoutChangingSelection(string input, bool valid)
    {
        var vm = new MainViewModel();
        for (int i = 0; i < 5; i++) vm.Pages.Add(new PaperItem { IsSelected = i == 4 });
        Assert.Equal(valid, vm.TrySelectRange(input, out _));
        Assert.Equal(valid ? 4 : 1, vm.SelectedPageCount);
    }

    [Fact]
    public void ExportArrangedPdf_ErrorMessageIsPropagatedInsteadOfSwallowed()
    {
        var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var missing = new PaperItem { SourceFilePath = Path.Combine(dir, "missing.pdf") };
            bool ok = PdfService.Instance.ExportArrangedPdf(new[] { missing }, Path.Combine(dir, "out.pdf"), out var errorMessage);
            Assert.False(ok);
            Assert.False(string.IsNullOrWhiteSpace(errorMessage));
            Assert.Contains("missing.pdf", errorMessage);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public async Task AppendImport_PreservesExistingUndoHistory_AndIsItselfUndoable()
    {
        var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var pdf1 = Path.Combine(dir, "first.pdf");
            var pdf2 = Path.Combine(dir, "second.pdf");
            Assert.True(PdfService.Instance.CreateBlankA4Pdf(pdf1));
            Assert.True(PdfService.Instance.CreateBlankA4Pdf(pdf2));

            var vm = new MainViewModel(new ExportPreferencesStore(Path.Combine(dir, "preferences.json")));
            await vm.AddFilesAsync(new[] { pdf1 });
            Assert.Single(vm.Pages);

            vm.RotatePage(vm.Pages[0], 90);
            Assert.Equal(90, vm.Pages[0].Rotation);
            Assert.True(vm.CanUndo);

            await vm.AddFilesAsync(new[] { pdf2 });
            Assert.Equal(2, vm.Pages.Count);
            Assert.True(vm.CanUndo);

            // 第一次復原應撤銷第二次匯入，回到僅有第一頁且維持 90° 的狀態
            vm.Undo();
            Assert.Single(vm.Pages);
            Assert.Equal(90, vm.Pages[0].Rotation);
            Assert.True(vm.CanUndo);

            // 第二次復原應回到旋轉前的 0° 狀態
            vm.Undo();
            Assert.Single(vm.Pages);
            Assert.Equal(0, vm.Pages[0].Rotation);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public async Task LargeCanvas_ThumbnailWindow_EvictsOnlyDistantItemsAboveThreshold()
    {
        var vm = new MainViewModel();
        var dummyImage = new System.Windows.Media.Imaging.WriteableBitmap(2, 2, 96, 96, System.Windows.Media.PixelFormats.Pbgra32, null);
        dummyImage.Freeze();

        // 120 頁以內不該主動清除既有縮圖
        for (int i = 0; i < 100; i++)
        {
            vm.Pages.Add(new PaperItem
            {
                SourceFilePath = $"dummy_{i}.pdf",
                Thumbnail = dummyImage,
                IsLoadingThumbnail = false
            });
        }

        await vm.UpdateVisibleThumbnailWindowAsync(Array.Empty<PaperItem>(), new HashSet<PaperItem>());
        Assert.All(vm.Pages, p => Assert.NotNull(p.Thumbnail));

        // 超過 120 頁後，應僅保留 retainSet 中的縮圖並釋放遠端縮圖
        for (int i = 100; i < 130; i++)
        {
            vm.Pages.Add(new PaperItem
            {
                SourceFilePath = $"dummy_{i}.pdf",
                Thumbnail = dummyImage,
                IsLoadingThumbnail = false
            });
        }

        var retain = new HashSet<PaperItem>(vm.Pages.Take(20));
        await vm.UpdateVisibleThumbnailWindowAsync(Array.Empty<PaperItem>(), retain);

        Assert.All(vm.Pages.Take(20), p => Assert.NotNull(p.Thumbnail));
        Assert.All(vm.Pages.Skip(20), p =>
        {
            Assert.Null(p.Thumbnail);
            Assert.True(p.IsLoadingThumbnail);
        });
    }

    [Fact]
    public async Task BusyStateBlocksDifferentExportCommandsAndDirectCalls()
    {
        var vm = new MainViewModel();
        vm.Pages.Add(new PaperItem { IsSelected = true });
        vm.IsBusy = true;
        Assert.False(vm.ExportAllPdfCommand.CanExecute(null));
        Assert.False(vm.ExportSelectedPdfCommand.CanExecute(null));
        Assert.False(vm.ExportImagesCommand.CanExecute(null));
        await vm.ExportAllPdfAsync();
        Assert.True(vm.IsBusy);
        Assert.False(vm.TrySelectRange("1", out _));
        vm.IsBusy = false;
        Assert.True(vm.ExportSelectedPdfCommand.CanExecute(null));
    }

    [Fact]
    public void PreferencesRetainOutputDirectoryAndSnapshotIsIndependent()
    {
        var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var store = new ExportPreferencesStore(Path.Combine(dir, "preferences.json"));
            Assert.True(store.Save(true, dir));
            Assert.Equal(dir, store.LoadOutputDirectory());
            var page = new PaperItem { Rotation = 90 };
            var snapshot = ExportCoordinator.Snapshot(new[] { page });
            page.Rotation = 180;
            Assert.Equal(90, snapshot[0].Rotation);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public async Task ScanPdfReadOnlyAcceptanceAndMetadataBenchmark()
    {
        var sourceDir = Environment.GetEnvironmentVariable("PAPERSWITCH_TEST_SOURCE");
        var outputRoot = Environment.GetEnvironmentVariable("PAPERSWITCH_TEST_OUTPUT");
        if (string.IsNullOrEmpty(sourceDir) || string.IsNullOrEmpty(outputRoot)) return;
        var output = Path.Combine(outputRoot, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(output);
        var files = Directory.GetFiles(sourceDir, "*.pdf");
        Assert.NotEmpty(files);
        var hashes = files.ToDictionary(p => p, p => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(p))));
        int pageCount = 0;
        var readable = new List<PaperItem>();
        var rejected = new List<string>();
        foreach (var file in files)
        {
            try
            {
                var pages = PdfImportService.ReadPages(file, Path.GetFileName(file));
                pageCount += pages.Count;
                if (pages.Count > 0) readable.Add(pages[0]);
            }
            catch (Exception ex) { rejected.Add(Path.GetFileName(file) + ": " + ex.GetType().Name); }
        }
        Assert.NotEmpty(readable);
        var sample = readable.Take(3).Reverse().ToList();
        sample[0].Rotation = 90;
        var options = new ExportOptions { OutputDirectory = output, AddPageNumbers = true };
        var coordinator = new ExportCoordinator();
        var merged = await coordinator.ExportAsync(ExportKind.AllPdf, sample, options, default);
        Assert.Empty(merged.Errors); Assert.Single(merged.Files);
        Assert.Equal(sample.Count, PdfService.Instance.GetPageCount(merged.Files[0]));
        Assert.Equal((PdfService.Instance.GetPageDimensions(sample[0].SourceFilePath, 0).Rotate + 90) % 360,
            PdfService.Instance.GetPageDimensions(merged.Files[0], 0).Rotate);
        var selected = await coordinator.ExportAsync(ExportKind.SelectedPdf, sample.Take(2).ToList(), options, default);
        Assert.Empty(selected.Errors); Assert.Single(selected.Files);
        Assert.Equal(2, PdfService.Instance.GetPageCount(selected.Files[0]));
        var split = await coordinator.ExportAsync(ExportKind.SplitPdf, sample, options, default);
        Assert.Empty(split.Errors); Assert.Equal(sample.Count, split.Files.Count);
        var png = await coordinator.ExportAsync(ExportKind.Images, sample.Take(1).ToList(), options, default);
        Assert.Empty(png.Errors); Assert.Single(png.Files);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        var stopped = await coordinator.ExportAsync(ExportKind.AllPdf, sample, options, cancelled.Token);
        Assert.True(stopped.Cancelled); Assert.Empty(stopped.Files);
        var benchmarks = new List<object>();
        foreach (int count in new[] { 20, 100, 300 })
        {
            var path = Path.Combine(output, $"benchmark-{count}.pdf");
            var repeated = Enumerable.Range(0, count).Select(i => sample[i % sample.Count]).ToList();
            Assert.True(PdfService.Instance.ExportArrangedPdf(repeated, path));
            var watch = Stopwatch.StartNew();
            var baseline = Enumerable.Range(0, count).Select(i => PdfService.Instance.GetPageDimensions(path, i)).ToList();
            watch.Stop(); var oldMs = watch.Elapsed.TotalMilliseconds;
            long before = GC.GetTotalMemory(false);
            watch.Restart();
            var current = PdfImportService.ReadPages(path, "基準文件");
            watch.Stop();
            Assert.Equal(count, current.Count);
            for (int i = 0; i < count; i++) Assert.Equal(baseline[i].Width, current[i].OriginalWidth, 2);
            benchmarks.Add(new { Pages = count, RepeatedOpenMs = oldMs, SingleOpenMs = watch.Elapsed.TotalMilliseconds,
                ManagedBytesDelta = GC.GetTotalMemory(false) - before });
        }
        Assert.Equal(files.Length, Directory.GetFiles(sourceDir, "*.pdf").Length);
        foreach (var pair in hashes) Assert.Equal(pair.Value, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(pair.Key))));
        Assert.Empty(Directory.GetFiles(output, "*.partial"));
        File.WriteAllText(Path.Combine(outputRoot, "scan-result.json"), JsonSerializer.Serialize(new
        {
            Files = files.Length, Pages = pageCount, ReadableFiles = readable.Count, Rejected = rejected,
            SourceHashesUnchanged = true, Output = output, Benchmarks = benchmarks,
            Environment = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription
        }, new JsonSerializerOptions { WriteIndented = true }));
    }

    [Fact]
    public async Task LightboxViewModel_CurrentPageContextExport_DispatchesOnlyCurrentPage()
    {
        var pages = new System.Collections.ObjectModel.ObservableCollection<PaperItem>
        {
            new() { SourceFilePath = "a.pdf", SourcePageIndex = 0, SourceFileName = "a.pdf", DisplayPageNumber = 1 },
            new() { SourceFilePath = "a.pdf", SourcePageIndex = 1, SourceFileName = "a.pdf", DisplayPageNumber = 2 }
        };

        var exports = new List<(ExportKind Kind, IReadOnlyList<PaperItem> Targets)>();
        bool observedExportingDuringCallback = false;
        LightboxViewModel? vm = null;

        var callbacks = new LightboxHostCallbacks(
            RotatePage: (p, delta) => p.Rotation = (p.Rotation + delta + 360) % 360,
            ExportPagesAsync: async (kind, targets) =>
            {
                observedExportingDuringCallback = vm?.IsExporting == true && vm.CanExportCurrent == false;
                exports.Add((kind, targets.ToList()));
                await Task.Yield();
            });

        vm = new LightboxViewModel(pages[0], pages, callbacks);
        Assert.True(vm.CanExportCurrent);
        Assert.True(vm.ExportCurrentPdfCommand.CanExecute(null));
        Assert.True(vm.ExportCurrentPngCommand.CanExecute(null));

        await vm.NextPage();
        Assert.Same(pages[1], vm.CurrentPage);

        await vm.ExportCurrentPdfAsync();
        Assert.True(observedExportingDuringCallback);
        Assert.False(vm.IsExporting);
        await vm.ExportCurrentPngAsync();

        Assert.Equal(2, exports.Count);
        Assert.Equal(ExportKind.SplitPdf, exports[0].Kind);
        Assert.Equal(new[] { pages[1] }, exports[0].Targets);
        Assert.Equal(ExportKind.Images, exports[1].Kind);
        Assert.Equal(new[] { pages[1] }, exports[1].Targets);
    }

    [Fact]
    public void LightboxViewModel_WithoutExportCallback_DisablesExportCommandsSafely()
    {
        var page = new PaperItem { SourceFilePath = "a.pdf", SourceFileName = "a.pdf", IsSelected = true };
        var pages = new System.Collections.ObjectModel.ObservableCollection<PaperItem> { page };

        var vm = new LightboxViewModel(page, pages);
        Assert.False(vm.CanExportCurrent);
        Assert.False(vm.ExportCurrentPdfCommand.CanExecute(null));
        Assert.False(vm.ExportCurrentPngCommand.CanExecute(null));
    }

    [Fact]
    public void MainViewModel_WorkspaceContextExportState_AdaptsToSingleAndMultiSelection()
    {
        var vm = new MainViewModel();
        var p1 = new PaperItem { SourceFilePath = "a.pdf", SourcePageIndex = 0, DisplayPageNumber = 1 };
        var p2 = new PaperItem { SourceFilePath = "b.pdf", SourcePageIndex = 0, DisplayPageNumber = 1 };
        var p3 = new PaperItem { SourceFilePath = "c.pdf", SourcePageIndex = 0, DisplayPageNumber = 1 };
        vm.Pages.Add(p1);
        vm.Pages.Add(p2);
        vm.Pages.Add(p3);

        Assert.False(vm.CanExportContextSelection);
        Assert.False(vm.HasMultipleSelectedPages);

        p2.IsSelected = true;
        vm.SelectedPage = p2;
        vm.NotifySelectionChanged();
        Assert.True(vm.CanExportContextSelection);
        Assert.False(vm.HasMultipleSelectedPages);
        Assert.Equal("匯出這頁 (PDF)", vm.ContextExportSelectedPdfText);
        Assert.Equal("匯出這頁 (PNG)", vm.ContextExportSelectedPngText);
        Assert.Equal(new[] { p2 }, vm.GetContextTargetPages());

        p1.IsSelected = true;
        vm.NotifySelectionChanged();
        Assert.True(vm.HasMultipleSelectedPages);
        Assert.Equal("合併匯出已選 2 頁 (PDF)", vm.ContextExportSelectedPdfText);
        Assert.Equal("拆分成 2 份獨立 PDF", vm.ContextExportSplitSelectedPdfText);
        Assert.Equal("匯出已選 2 頁 (PNG)", vm.ContextExportSelectedPngText);
        Assert.Equal(new[] { p1, p2 }, vm.GetContextTargetPages());
    }

    [Fact]
    public async Task MainViewModel_ExportCustomPagesAsync_ExportsOnlyRequestedPagesWithoutDialog()
    {
        var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        var outDir = Path.Combine(dir, "out");
        Directory.CreateDirectory(dir);
        Directory.CreateDirectory(outDir);
        try
        {
            var pdf1 = Path.Combine(dir, "alpha.pdf");
            var pdf2 = Path.Combine(dir, "beta.pdf");
            Assert.True(PdfService.Instance.CreateBlankA4Pdf(pdf1));
            Assert.True(PdfService.Instance.CreateBlankA4Pdf(pdf2));

            var vm = new MainViewModel(new ExportPreferencesStore(Path.Combine(dir, "preferences.json")));
            vm.Options.OutputDirectory = outDir;
            await vm.AddFilesAsync(new[] { pdf1, pdf2 });
            Assert.Equal(2, vm.Pages.Count);

            await vm.ExportCustomPagesAsync(ExportKind.SplitPdf, new[] { vm.Pages[1] }, showCompletionDialog: false);
            var splitFiles = Directory.GetFiles(outDir, "*.pdf");
            Assert.Single(splitFiles);
            Assert.Contains("beta", Path.GetFileNameWithoutExtension(splitFiles[0]), StringComparison.OrdinalIgnoreCase);

            await vm.ExportCustomPagesAsync(ExportKind.SelectedPdf, new[] { vm.Pages[0], vm.Pages[1] }, showCompletionDialog: false);
            var allPdfFiles = Directory.GetFiles(outDir, "*.pdf");
            Assert.Equal(2, allPdfFiles.Length);
            var mergedFile = allPdfFiles.Except(splitFiles).Single();
            Assert.Equal(2, PdfService.Instance.GetPageCount(mergedFile));

            await vm.ExportCustomPagesAsync(ExportKind.Images, new[] { vm.Pages[0] }, showCompletionDialog: false);
            Assert.Single(Directory.GetFiles(outDir, "*.png"));
        }
        finally { Directory.Delete(dir, true); }
    }
}
