using System.IO;

namespace PaperSwitch.Services;

/// <summary>匯出與預覽共用命名規則；用途說明可選，流水號與防覆寫分開處理。</summary>
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

    public static string Merged(string? customName, bool selected, bool includeDescription)
    {
        string stem = GetStem(customName);
        return stem + (includeDescription && selected ? "_選取頁面" : "") + ".pdf";
    }

    public static string Page(string sourceName, int sourcePage, int sequence, string? customName, string extension, bool includeDescription)
    {
        string stem = string.IsNullOrWhiteSpace(customName) ? GetStem(sourceName, sourceFile: true) : GetStem(customName);
        return $"{stem}{(includeDescription ? $"_頁{sourcePage}" : "")}_{sequence:D3}{extension}";
    }
}
