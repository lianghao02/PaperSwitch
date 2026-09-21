using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;
using PaperSwitch.Models;

namespace PaperSwitch.Services
{
    /// <summary>
    /// 基於 Windows.Data.Pdf (WinRT) 與 WPF PngBitmapEncoder 的 PDF 頁面圖片匯出服務
    /// </summary>
    public class ImageExportService
    {
        public static ImageExportService Instance { get; } = new();

        /// <summary>
        /// 預設匯出圖片目標寬度 (提供清晰閱讀與列印品質)
        /// </summary>
        public const int DefaultTargetWidth = 2400;

        /// <summary>
        /// 產生檔案名稱（純邏輯函式，利於單元測試）
        /// </summary>
        public static string BuildPngFileName(
            string sourceFileName,
            int sourcePageNumber,
            int canvasIndex,
            string? customPrefix = null)
        {
            if (!string.IsNullOrWhiteSpace(customPrefix))
            {
                string cleanPrefix = SanitizeFileName(customPrefix.Trim());
                return $"{cleanPrefix}_{canvasIndex:D3}.png";
            }

            string cleanSource = SanitizeFileName(sourceFileName);
            string stem = Path.GetFileNameWithoutExtension(cleanSource);
            string cleanStem = SanitizeFileName(stem);
            return $"{cleanStem}_頁{sourcePageNumber}_{canvasIndex:D3}.png";
        }

        /// <summary>
        /// 取得防覆寫的唯一輸出完整路徑
        /// </summary>
        public static string GetUniqueFilePath(string outputDirectory, string desiredFileName)
        {
            string targetPath = Path.Combine(outputDirectory, desiredFileName);
            if (!File.Exists(targetPath))
            {
                return targetPath;
            }

            string stem = Path.GetFileNameWithoutExtension(desiredFileName);
            string ext = Path.GetExtension(desiredFileName);
            int counter = 1;

            while (File.Exists(targetPath))
            {
                targetPath = Path.Combine(outputDirectory, $"{stem}_{counter}{ext}");
                counter++;
            }

            return targetPath;
        }

        /// <summary>
        /// 清理檔名中的非法字元
        /// </summary>
        public static string SanitizeFileName(string fileName, string fallback = "PaperSwitch_Page")
        {
            if (string.IsNullOrWhiteSpace(fileName)) return fallback;
            string clean = fileName.Trim();
            foreach (char invalid in Path.GetInvalidFileNameChars())
            {
                clean = clean.Replace(invalid, '_');
            }
            clean = clean.Trim('.', ' ', '_');
            return string.IsNullOrWhiteSpace(clean) ? fallback : clean;
        }

        /// <summary>
        /// 渲染單一 PDF 頁面為 BitmapSource，若有旋轉角度則套用旋轉
        /// </summary>
        public async Task<BitmapSource?> RenderPageToBitmapAsync(
            string pdfPath,
            int pageIndex,
            int rotation = 0,
            int targetWidth = DefaultTargetWidth)
        {
            if (!File.Exists(pdfPath) || pageIndex < 0) return null;

            try
            {
                var storageFile = await StorageFile.GetFileFromPathAsync(Path.GetFullPath(pdfPath));
                var pdfDoc = await PdfDocument.LoadFromFileAsync(storageFile);
                if (pageIndex >= (int)pdfDoc.PageCount) return null;

                using var page = pdfDoc.GetPage((uint)pageIndex);
                using var memStream = new InMemoryRandomAccessStream();
                var options = new PdfPageRenderOptions { DestinationWidth = (uint)Math.Max(1, targetWidth) };
                await page.RenderToStreamAsync(memStream, options);
                memStream.Seek(0);

                using var netStream = memStream.AsStreamForRead();
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.StreamSource = netStream;
                bitmap.EndInit();
                bitmap.Freeze();

                if (rotation % 360 != 0)
                {
                    var rotated = new TransformedBitmap(bitmap, new RotateTransform(rotation));
                    rotated.Freeze();
                    return rotated;
                }

                return bitmap;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ImageExportService] 渲染頁面失敗 {pdfPath} P.{pageIndex}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 批次匯出指定的頁面集合為 PNG 圖片
        /// </summary>
        public async Task<ImageExportResult> ExportPagesToPngAsync(
            IEnumerable<PaperItem> pages,
            string outputDirectory,
            string? customPrefix = null,
            IProgress<(int Current, int Total, string FileName)>? progress = null,
            CancellationToken cancellationToken = default)
        {
            var pageList = pages.ToList();
            var exportedFiles = new List<string>();
            var errors = new List<(PaperItem Page, string ErrorMessage)>();

            if (pageList.Count == 0)
            {
                return new ImageExportResult(exportedFiles, errors);
            }

            Directory.CreateDirectory(outputDirectory);

            int total = pageList.Count;
            for (int i = 0; i < total; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var item = pageList[i];
                int canvasIndex = i + 1;
                string desiredName = BuildPngFileName(item.SourceFileName, item.DisplayPageNumber, canvasIndex, customPrefix);
                progress?.Report((canvasIndex, total, desiredName));

                try
                {
                    var bitmap = await RenderPageToBitmapAsync(item.SourceFilePath, item.SourcePageIndex, item.Rotation);
                    if (bitmap == null)
                    {
                        errors.Add((item, $"無法渲染來源檔案 {item.SourceFileName} 的第 {item.DisplayPageNumber} 頁。"));
                        continue;
                    }

                    string finalPath = GetUniqueFilePath(outputDirectory, desiredName);

                    using (var fileStream = new FileStream(finalPath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        var encoder = new PngBitmapEncoder();
                        encoder.Frames.Add(BitmapFrame.Create(bitmap));
                        encoder.Save(fileStream);
                    }

                    exportedFiles.Add(finalPath);
                }
                catch (Exception ex)
                {
                    errors.Add((item, ex.Message));
                    System.Diagnostics.Debug.WriteLine($"[ImageExportService] 匯出圖片失敗 {item.SourceFileName}: {ex.Message}");
                }
            }

            return new ImageExportResult(exportedFiles, errors);
        }
    }

    /// <summary>
    /// 圖片匯出結果模型
    /// </summary>
    public sealed record ImageExportResult(
        IReadOnlyList<string> ExportedFiles,
        IReadOnlyList<(PaperItem Page, string ErrorMessage)> Errors)
    {
        public bool Success => Errors.Count == 0 && ExportedFiles.Count > 0;
        public int ExportedCount => ExportedFiles.Count;
        public int ErrorCount => Errors.Count;
    }
}
