using PaperSwitch.Models;
using PdfSharp.Pdf.IO;

namespace PaperSwitch.Services;

/// <summary>每份 PDF 只開啟一次；傳回資料後由 UI 分批加入集合。</summary>
public static class PdfImportService
{
    public static List<PaperItem> ReadPages(string path, string name, CancellationToken token = default)
    {
        using var document = PdfReader.Open(path, PdfDocumentOpenMode.Import);
        var result = new List<PaperItem>(document.PageCount);
        for (int index = 0; index < document.PageCount; index++)
        {
            token.ThrowIfCancellationRequested();
            var page = document.Pages[index];
            result.Add(new PaperItem
            {
                SourceFilePath = path, SourceFileName = name, SourcePageIndex = index,
                DisplayPageNumber = index + 1, TotalPagesInSource = document.PageCount,
                OriginalWidth = page.Width.Point, OriginalHeight = page.Height.Point,
                IsLoadingThumbnail = true
            });
        }
        return result;
    }
}
