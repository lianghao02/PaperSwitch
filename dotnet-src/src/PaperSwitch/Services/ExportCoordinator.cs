using System.IO;
using PaperSwitch.Models;

namespace PaperSwitch.Services;

public enum ExportKind { AllPdf, SelectedPdf, SplitPdf, Images }
public sealed record ExportResult(int Requested, IReadOnlyList<string> Files, IReadOnlyList<string> Errors, bool Cancelled)
{
    public int CompletedPages { get; init; }
    public string Summary => $"要求 {Requested} 頁；完成 {CompletedPages} 頁／{Files.Count} 個檔案；未輸出 {Requested - CompletedPages} 頁；錯誤 {Errors.Count} 項" + (Cancelled ? "；已停止後續工作" : "");
}

/// <summary>輸出只使用本次快照，不讀取處理中的畫面設定。</summary>
public sealed class ExportCoordinator
{
    public static List<PaperItem> Snapshot(IEnumerable<PaperItem> pages) => pages.Select(p => new PaperItem
    {
        SourceFilePath = p.SourceFilePath, SourceFileName = p.SourceFileName,
        SourcePageIndex = p.SourcePageIndex, DisplayPageNumber = p.DisplayPageNumber,
        Rotation = p.Rotation, IsGeneratedBlank = p.IsGeneratedBlank
    }).ToList();

    public static ExportOptions Snapshot(ExportOptions o) => new()
    {
        OutputDirectory = o.OutputDirectory, CustomFileName = o.CustomFileName,
        IncludeFileNameDescription = o.IncludeFileNameDescription, AddPageNumbers = o.AddPageNumbers,
        PageNumberPosition = o.PageNumberPosition, PageNumberFontSize = o.PageNumberFontSize
    };

    public async Task<ExportResult> ExportAsync(ExportKind kind, IReadOnlyList<PaperItem> pages,
        ExportOptions options, CancellationToken token, IProgress<(int Current, int Total, string FileName)>? progress = null)
    {
        var files = new List<string>();
        var errors = new List<string>();
        try
        {
            token.ThrowIfCancellationRequested();
            if (kind == ExportKind.Images)
            {
                var result = await Task.Run(() => ImageExportService.Instance.ExportPagesToPngAsync(pages, options.OutputDirectory,
                    options.CustomFileName, progress, token, options.IncludeFileNameDescription));
                files.AddRange(result.ExportedFiles);
                errors.AddRange(result.Errors.Select(e => $"{e.Page.SourceFileName} 第 {e.Page.DisplayPageNumber} 頁：{e.ErrorMessage}"));
                return new(pages.Count, files, errors, token.IsCancellationRequested) { CompletedPages = files.Count };
            }
            await Task.Run(() =>
            {
                // 合併與拆分先完整檢查來源，拒絕靜默漏頁。
                var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var p in pages)
                {
                    token.ThrowIfCancellationRequested();
                    if (!counts.TryGetValue(p.SourceFilePath, out var count))
                        counts[p.SourceFilePath] = count = PdfService.Instance.GetPageCount(p.SourceFilePath);
                    if (p.SourcePageIndex < 0 || p.SourcePageIndex >= count)
                        errors.Add($"{p.SourceFileName} 第 {p.DisplayPageNumber} 頁無法讀取，請重新匯入。 ");
                }
                if (errors.Count > 0) return;
                if (kind == ExportKind.SplitPdf)
                    PdfService.Instance.ExportIndividualPdfs(pages, options.OutputDirectory,
                        options.CustomFileName, options.IncludeFileNameDescription, options, token, path =>
                        {
                            files.Add(path);
                            progress?.Report((files.Count, pages.Count, Path.GetFileName(path)));
                        });
                else
                {
                    var path = ExportFileNaming.MergedOutputPath(options.OutputDirectory, options.CustomFileName,
                        kind == ExportKind.SelectedPdf, options.IncludeFileNameDescription,
                        kind == ExportKind.AllPdf ? ExportFileNaming.FirstSourceName(pages) : null, DateTime.Now);
                    if (PdfService.Instance.ExportArrangedPdf(pages, path, options, token)) files.Add(path);
                    else if (!token.IsCancellationRequested) errors.Add("PDF 未完成；請檢查來源與輸出資料夾是否可讀寫。");
                }
            });
        }
        catch (OperationCanceledException) { }
        catch (Exception ex) { errors.Add(ex.Message); }
        return new(pages.Count, files, errors, token.IsCancellationRequested)
        {
            CompletedPages = kind is ExportKind.AllPdf or ExportKind.SelectedPdf ? (files.Count > 0 ? pages.Count : 0) : files.Count
        };
    }
}
