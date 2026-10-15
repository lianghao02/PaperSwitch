using System.IO;
using PaperSwitch.Models;
using PaperSwitch.Services;
using PaperSwitch.ViewModels;
using PdfSharp.Pdf;
using Xunit;

namespace PaperSwitch.Tests;

public sealed class ExportNamingTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "PaperSwitch_Naming_" + Guid.NewGuid().ToString("N"));
    private string SettingsPath => Path.Combine(_directory, "preferences.json");
    public ExportNamingTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, true);

    [Theory]
    [InlineData(false, false, "報告.pdf")]
    [InlineData(true, false, "報告.pdf")]
    [InlineData(true, true, "報告_選取頁面.pdf")]
    [InlineData(false, true, "報告.pdf")]
    public void Merged_DescriptionAffectsOnlySelectedPages(bool selected, bool description, string expected)
        => Assert.Equal(expected, ExportFileNaming.Merged("報告.pdf", selected, description));

    [Theory]
    [InlineData(false, "報告_001.png")]
    [InlineData(true, "報告_頁3_001.png")]
    public void Png_NamingRetainsSequenceAndOptionalSourcePage(bool description, string expected)
        => Assert.Equal(expected, ImageExportService.BuildPngFileName("來源.pdf", 3, 1, "報告.png", description));

    [Theory]
    [InlineData(false, "報告_001.pdf", "報告_001_1.pdf")]
    [InlineData(true, "報告_頁3_001.pdf", "報告_頁3_001_1.pdf")]
    public void Pdf_ActualExportPreservesExistingFile(bool description, string name, string duplicate)
    {
        string source = Path.Combine(_directory, "source.pdf");
        using (var pdf = new PdfDocument()) { pdf.AddPage(); pdf.Save(source); }
        string output = Path.Combine(_directory, "out");
        var pages = new[] { new PaperItem { SourceFilePath = source, SourceFileName = "來源.pdf", SourcePageIndex = 0, DisplayPageNumber = 3 } };
        string first = Assert.Single(PdfService.Instance.ExportIndividualPdfs(pages, output, "報告", description));
        byte[] original = File.ReadAllBytes(first);
        string second = Assert.Single(PdfService.Instance.ExportIndividualPdfs(pages, output, "報告", description));
        Assert.Equal(name, Path.GetFileName(first));
        Assert.Equal(duplicate, Path.GetFileName(second));
        Assert.Equal(original, File.ReadAllBytes(first));
        Assert.Equal(1, PdfService.Instance.GetPageCount(second));
    }

    [Fact]
    public void Preference_MissingCorruptAndReloadedValuesAreHandled()
    {
        var store = new ExportPreferencesStore(SettingsPath);
        Assert.False(store.Load());
        File.WriteAllText(SettingsPath, "invalid json");
        Assert.False(store.Load());
        Assert.True(store.Save(true));
        Assert.True(new ExportPreferencesStore(SettingsPath).Load());
        Assert.True(store.Save(false));
        Assert.False(new ExportPreferencesStore(SettingsPath).Load());
    }

    [Fact]
    public void Preview_TracksNameSelectionAndPersistedSwitch()
    {
        var vm = new MainViewModel(new ExportPreferencesStore(SettingsPath));
        vm.Options.CustomFileName = "報告.v2.pdf";
        vm.Pages.Add(new PaperItem { SourceFileName = "來源.pdf", DisplayPageNumber = 3, IsSelected = true });
        vm.NotifySelectionChanged();
        var notifications = new List<string?>();
        vm.PropertyChanged += (_, e) => notifications.Add(e.PropertyName);
        vm.Options.IncludeFileNameDescription = true;
        Assert.Contains(nameof(vm.ExportFileNamePreview), notifications);
        Assert.Contains("另存：報告.v2_選取頁面.pdf", vm.ExportFileNamePreview);
        Assert.Contains("拆分：報告.v2_頁3_001.pdf", vm.ExportFileNamePreview);
        Assert.True(new MainViewModel(new ExportPreferencesStore(SettingsPath)).Options.IncludeFileNameDescription);
        vm.DeselectAll();
        Assert.Contains("另存：報告.v2.pdf", vm.ExportFileNamePreview);
        vm.Options.IncludeFileNameDescription = false;
        Assert.DoesNotContain("_頁3", vm.ExportFileNamePreview);
        Assert.Contains("拆分：報告.v2_001.pdf", vm.ExportFileNamePreview);
        vm.Options.CustomFileName = "";
        vm.Pages[0].IsSelected = true;
        vm.NotifySelectionChanged();
        notifications.Clear();
        Assert.True(vm.TryRenameSelectedPages("新名稱", out _));
        Assert.Contains(nameof(vm.ExportFileNamePreview), notifications);
        Assert.Contains("拆分：新名稱_001.pdf", vm.ExportFileNamePreview);
    }

    [Fact]
    public void SourceExtensionAndInvalidCustomNamesMatchPreviewRules()
    {
        Assert.Equal("照片.v2_001.png", ExportFileNaming.Page("照片.v2.jpg", 1, 1, null, ".png", false));
        Assert.Equal("PaperSwitch_001.pdf", ExportFileNaming.Page("來源.pdf", 1, 1, "///", ".pdf", false));
    }

    [Fact]
    public void Preference_UnwritablePathDoesNotThrow()
    {
        string blocker = Path.Combine(_directory, "not-a-directory");
        File.WriteAllText(blocker, "keep");
        var store = new ExportPreferencesStore(Path.Combine(blocker, "preferences.json"));
        Assert.False(store.Save(true));
        Assert.Equal("keep", File.ReadAllText(blocker));
    }
}
