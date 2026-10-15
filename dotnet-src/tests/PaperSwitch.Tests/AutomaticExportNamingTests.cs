using System.IO;
using PaperSwitch.Models;
using PaperSwitch.Services;
using PaperSwitch.ViewModels;
using PdfSharp.Pdf;
using Xunit;

namespace PaperSwitch.Tests;

public sealed class AutomaticExportNamingTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "PaperSwitch_AutoNames_" + Guid.NewGuid().ToString("N"));
    private static readonly DateTime ExportTime = new(2026, 10, 15, 9, 30, 0);
    public AutomaticExportNamingTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, true);
    private string Pdf(string name)
    {
        string path = Path.Combine(_directory, name);
        using var document = new PdfDocument();
        document.AddPage();
        document.Save(path);
        return path;
    }

    [Theory]
    [InlineData("案件資料.pdf", "案件資料.pdf")]
    [InlineData("照片.v2.jpg", "照片.v2.pdf")]
    public void Merge_EmptyInputUsesOriginalSourceStem(string source, string expected)
        => Assert.Equal(expected, ExportFileNaming.Merged("  ", false, false, source, ExportTime));

    [Fact]
    public void Merge_RespectsReorderedFirstSourceAndIgnoresUnnamedBlankPages()
    {
        string first = Pdf("甲.pdf"), second = Pdf("乙.pdf");
        var pages = new List<PaperItem> {
            new() { SourceFilePath = first, SourceFileName = "甲.pdf" },
            new() { SourceFilePath = second, SourceFileName = "乙.pdf" }
        };
        pages.Reverse();
        string output = ExportFileNaming.MergedOutputPath(_directory, "", false, false, ExportFileNaming.FirstSourceName(pages), ExportTime);
        Assert.Equal("乙_1.pdf", Path.GetFileName(output)); // 來源乙.pdf 已存在，必須保留。
        byte[] original = File.ReadAllBytes(second);
        Assert.True(PdfService.Instance.ExportArrangedPdf(pages, output));
        Assert.Equal(2, PdfService.Instance.GetPageCount(output));
        Assert.Equal(original, File.ReadAllBytes(second));
        pages.Insert(0, new PaperItem { IsGeneratedBlank = true, SourceFileName = "空白頁.pdf" });
        Assert.Equal("乙.pdf", ExportFileNaming.FirstSourceName(pages));
    }

    [Theory]
    [InlineData(true, false, "案件資料.pdf", "20261015_0930_001.pdf")]
    [InlineData(true, true, "案件資料.pdf", "20261015_0930_選取頁面_001.pdf")]
    [InlineData(false, false, null, "20261015_0930_001.pdf")]
    public void Merge_SelectedOrMissingSourceUsesTimestamp(bool selected, bool description, string? source, string expected)
        => Assert.Equal(expected, ExportFileNaming.Merged(null, selected, description, source, ExportTime));

    [Theory]
    [InlineData(false, "手動名稱.pdf")]
    [InlineData(true, "手動名稱_選取頁面.pdf")]
    public void ExplicitInputTakesPriorityOverTimestampAndSource(bool description, string expected)
        => Assert.Equal(expected, ExportFileNaming.Merged("手動名稱.pdf", true, description, "來源.pdf", ExportTime));

    [Fact]
    public void TimestampMerge_UsesNextSequenceAndPreservesExistingFiles()
    {
        string first = ExportFileNaming.MergedOutputPath(_directory, "", true, false, null, ExportTime);
        File.WriteAllText(first, "原始成品");
        string second = ExportFileNaming.MergedOutputPath(_directory, "", true, false, null, ExportTime);
        File.WriteAllText(second, "第二份成品");
        string third = ExportFileNaming.MergedOutputPath(_directory, "", true, false, null, ExportTime);
        Assert.Equal("20261015_0930_001.pdf", Path.GetFileName(first));
        Assert.Equal("20261015_0930_002.pdf", Path.GetFileName(second));
        Assert.Equal("20261015_0930_003.pdf", Path.GetFileName(third));
        Assert.Equal("原始成品", File.ReadAllText(first));
    }

    [Fact]
    public void Split_MixedSourcesKeepOwnNamesInCanvasOrder()
    {
        string source = Pdf("source.pdf");
        var pages = new[] {
            new PaperItem { SourceFilePath = source, SourceFileName = "乙.pdf", DisplayPageNumber = 2 },
            new PaperItem { SourceFilePath = source, SourceFileName = "甲.pdf", DisplayPageNumber = 1 }
        };
        var files = PdfService.Instance.ExportIndividualPdfs(pages, Path.Combine(_directory, "out"));
        Assert.Equal(new[] { "乙_001.pdf", "甲_002.pdf" }, files.Select(Path.GetFileName));
        Assert.All(files, file => Assert.Equal(1, PdfService.Instance.GetPageCount(file)));
    }

    [Fact]
    public void BlankPages_UseOneTimestampForEntireSplitBatch()
    {
        string source = Pdf("blank.pdf");
        var page = new PaperItem { SourceFilePath = source, SourceFileName = "空白頁.pdf", IsGeneratedBlank = true };
        var files = PdfService.Instance.ExportIndividualPdfs(new[] { page, page }, Path.Combine(_directory, "blank-out"));
        string first = Path.GetFileName(files[0]), second = Path.GetFileName(files[1]);
        Assert.Matches(@"^\d{8}_\d{4}_001\.pdf$", first);
        Assert.Equal(first.Replace("_001.pdf", "_002.pdf"), second);
        Assert.Null(ExportFileNaming.FirstSourceName(new[] { page }));
        Assert.Equal("空白頁.pdf", ExportFileNaming.SourceName(new PaperItem { SourceFileName = "空白頁.pdf" }));
    }

    [Theory]
    [InlineData(".pdf")]
    [InlineData(".png")]
    public void MissingPageName_UsesProvidedBatchTimestamp(string extension)
        => Assert.Equal("20261015_0930_002" + extension, ExportFileNaming.Page("", 1, 2, null, extension, false, ExportTime));

    [Fact]
    public async Task ImportedImage_KeepsOriginalNameForPngExportInsteadOfTemporaryGuid()
    {
        string sourceStem = "來源圖片_" + Guid.NewGuid().ToString("N");
        string source = Path.Combine(_directory, sourceStem + ".png");
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScL1xQAAAABJRU5ErkJggg==");
        File.WriteAllBytes(source, png);
        var vm = new MainViewModel();
        string? generated = null;
        try
        {
            await vm.AddFilesAsync(new[] { source });
            var page = Assert.Single(vm.Pages);
            generated = page.SourceFilePath;
            Assert.Equal(Path.GetFileName(source), page.SourceFileName);
            var result = await ImageExportService.Instance.ExportPagesToPngAsync(vm.Pages, Path.Combine(_directory, "png-out"));
            Assert.Equal(0, result.ErrorCount);
            Assert.Equal(sourceStem + "_001.png", Path.GetFileName(Assert.Single(result.ExportedFiles)));
            Assert.Equal(png, File.ReadAllBytes(source));
        }
        finally
        {
            if (generated != null && Path.GetFileName(generated).StartsWith(sourceStem + "_", StringComparison.Ordinal)
                && Path.GetExtension(generated) == ".pdf") File.Delete(generated);
        }
    }
}
