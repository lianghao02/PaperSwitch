using System.Globalization;
using System.IO;
using PaperSwitch.Models;

namespace PaperSwitch.Services;

/// <summary>匯出共用命名規則；手動名稱優先，流水號與用途說明分開處理。</summary>
public static class ExportFileNaming
{
    public static string GetStem(string? value, string fallback = "PaperSwitch", bool sourceFile = false)
    {
        string stem = string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
        foreach (char invalid in Path.GetInvalidFileNameChars()) stem = stem.Replace(invalid, '_');
        if (sourceFile) stem = Path.GetFileNameWithoutExtension(stem);
        else if (stem.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) || stem.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
            stem = stem[..^4];
        stem = stem.Trim('.', ' ', '_');
        return string.IsNullOrWhiteSpace(stem) ? fallback : stem;
    }

    public static string SourceName(PaperItem page) => page.IsGeneratedBlank && page.SourceFileName == "空白頁.pdf"
        ? string.Empty : page.SourceFileName;

    public static string? FirstSourceName(IEnumerable<PaperItem> pages) => pages.Select(SourceName)
        .FirstOrDefault(name => !string.IsNullOrWhiteSpace(GetStem(name, "", sourceFile: true)));

    private static string Timestamp(DateTime time) => time.ToString("yyyyMMdd_HHmm", CultureInfo.InvariantCulture);

    public static string Merged(string? customName, bool selected, bool includeDescription,
        string? sourceName = null, DateTime? exportTime = null, int sequence = 1)
    {
        string sourceStem = GetStem(sourceName, "", sourceFile: true);
        bool timestamp = string.IsNullOrWhiteSpace(customName) && (selected || string.IsNullOrWhiteSpace(sourceStem));
        string stem = !string.IsNullOrWhiteSpace(customName) ? GetStem(customName)
            : timestamp ? Timestamp(exportTime ?? DateTime.Now) : sourceStem;
        string description = includeDescription && selected ? "_選取頁面" : "";
        return $"{stem}{description}{(timestamp ? $"_{sequence:D3}" : "")}.pdf";
    }

    public static string MergedOutputPath(string outputDirectory, string? customName, bool selected,
        bool includeDescription, string? sourceName, DateTime exportTime)
    {
        bool timestamp = string.IsNullOrWhiteSpace(customName) &&
            (selected || string.IsNullOrWhiteSpace(GetStem(sourceName, "", sourceFile: true)));
        string name = Merged(customName, selected, includeDescription, sourceName, exportTime);
        if (!timestamp) return ImageExportService.GetUniqueFilePath(outputDirectory, name);
        int sequence = 1;
        string path = Path.Combine(outputDirectory, name);
        while (File.Exists(path))
            path = Path.Combine(outputDirectory, Merged(customName, selected, includeDescription, sourceName, exportTime, ++sequence));
        return path;
    }

    public static string Page(string sourceName, int sourcePage, int sequence, string? customName, string extension,
        bool includeDescription, DateTime? exportTime = null)
    {
        string sourceStem = GetStem(sourceName, "", sourceFile: true);
        string stem = !string.IsNullOrWhiteSpace(customName) ? GetStem(customName)
            : string.IsNullOrWhiteSpace(sourceStem) ? Timestamp(exportTime ?? DateTime.Now) : sourceStem;
        return $"{stem}{(includeDescription ? $"_頁{sourcePage}" : "")}_{sequence:D3}{extension}";
    }
}
