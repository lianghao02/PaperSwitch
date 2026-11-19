using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;

namespace PaperSwitch.Services
{
    /// <summary>
    /// 基於 Windows.Data.Pdf (WinRT) 的高畫質非同步 PDF 縮圖渲染與快取服務
    /// </summary>
    public class ThumbnailCacheService
    {
        public static ThumbnailCacheService Instance { get; } = new();

        private const int MaximumThumbnailEntries = 360;
        private const int MaximumPreviewEntries = 8;
        private const int PreviewWidthThreshold = 720;

        private readonly CachePool _thumbnailPool = new(MaximumThumbnailEntries);
        private readonly CachePool _previewPool = new(MaximumPreviewEntries);
        private readonly ConcurrentDictionary<string, Lazy<Task<ImageSource?>>> _inflight = new();

        private sealed class CacheEntry(ImageSource image)
        {
            public ImageSource Image { get; } = image;
            public LinkedListNode<string>? Node { get; set; }
        }

        private sealed class CachePool(int capacity)
        {
            private readonly object _lock = new();
            private readonly Dictionary<string, CacheEntry> _map = new();
            private readonly LinkedList<string> _lru = new();

            public bool TryGet(string key, out ImageSource? image)
            {
                lock (_lock)
                {
                    if (_map.TryGetValue(key, out var entry))
                    {
                        _lru.Remove(entry.Node!);
                        _lru.AddFirst(entry.Node!);
                        image = entry.Image;
                        return true;
                    }
                }
                image = null;
                return false;
            }

            public void Put(string key, ImageSource image)
            {
                lock (_lock)
                {
                    if (_map.TryGetValue(key, out var existing))
                    {
                        _lru.Remove(existing.Node!);
                        _map.Remove(key);
                    }
                    var node = _lru.AddFirst(key);
                    _map[key] = new CacheEntry(image) { Node = node };
                    while (_map.Count > capacity)
                    {
                        var oldest = _lru.Last!;
                        _map.Remove(oldest.Value);
                        _lru.RemoveLast();
                    }
                }
            }

            public void Clear()
            {
                lock (_lock)
                {
                    _map.Clear();
                    _lru.Clear();
                }
            }
        }

        private CachePool SelectPool(int targetWidth) =>
            targetWidth > PreviewWidthThreshold ? _previewPool : _thumbnailPool;

        public void ClearCache()
        {
            _thumbnailPool.Clear();
            _previewPool.Clear();
        }

        public async Task<ImageSource?> GetThumbnailAsync(string pdfPath, int pageIndex, int targetWidth = 320)
        {
            if (!File.Exists(pdfPath) || pageIndex < 0) return null;

            string key = $"{pdfPath}|{pageIndex}|{targetWidth}";
            var pool = SelectPool(targetWidth);
            if (pool.TryGet(key, out var cached)) return cached;

            var lazyRender = _inflight.GetOrAdd(key, _ => new Lazy<Task<ImageSource?>>(
                () => RenderAndCacheAsync(pdfPath, pageIndex, targetWidth, key, pool)));

            try
            {
                return await lazyRender.Value;
            }
            finally
            {
                _inflight.TryRemove(key, out _);
            }
        }

        /// <summary>
        /// 針對同一份 PDF 開檔一次並渲染多頁縮圖，渲染完隨即釋放文件句柄（不鎖檔）。
        /// </summary>
        public async Task<Dictionary<int, ImageSource?>> GetThumbnailsForPdfAsync(
            string pdfPath,
            IReadOnlyList<int> pageIndices,
            int targetWidth = 360,
            CancellationToken cancellationToken = default)
        {
            var results = new Dictionary<int, ImageSource?>(pageIndices.Count);
            if (!File.Exists(pdfPath) || pageIndices.Count == 0) return results;

            var pool = SelectPool(targetWidth);
            var missing = new List<int>();
            foreach (int pageIndex in pageIndices)
            {
                if (pageIndex < 0 || results.ContainsKey(pageIndex)) continue;
                string key = $"{pdfPath}|{pageIndex}|{targetWidth}";
                if (pool.TryGet(key, out var cached))
                {
                    results[pageIndex] = cached;
                }
                else
                {
                    missing.Add(pageIndex);
                }
            }

            if (missing.Count == 0 || cancellationToken.IsCancellationRequested) return results;

            try
            {
                var storageFile = await StorageFile.GetFileFromPathAsync(Path.GetFullPath(pdfPath));
                var pdfDoc = await PdfDocument.LoadFromFileAsync(storageFile);

                foreach (int pageIndex in missing)
                {
                    if (cancellationToken.IsCancellationRequested) break;
                    if (pageIndex >= (int)pdfDoc.PageCount)
                    {
                        results[pageIndex] = null;
                        continue;
                    }

                    var image = await RenderPageCoreAsync(pdfDoc, pageIndex, targetWidth);
                    string key = $"{pdfPath}|{pageIndex}|{targetWidth}";
                    pool.Put(key, image);
                    results[pageIndex] = image;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ThumbnailCacheService] 批次縮圖渲染失敗 {pdfPath}: {ex.Message}");
                foreach (int pageIndex in missing)
                {
                    if (results.ContainsKey(pageIndex)) continue;
                    var fallback = CreateFallbackThumbnail(pageIndex + 1);
                    pool.Put($"{pdfPath}|{pageIndex}|{targetWidth}", fallback);
                    results[pageIndex] = fallback;
                }
            }

            return results;
        }

        private async Task<ImageSource?> RenderAndCacheAsync(string pdfPath, int pageIndex, int targetWidth, string key, CachePool pool)
        {
            ImageSource image;
            try
            {
                var storageFile = await StorageFile.GetFileFromPathAsync(Path.GetFullPath(pdfPath));
                var pdfDoc = await PdfDocument.LoadFromFileAsync(storageFile);
                if (pageIndex >= (int)pdfDoc.PageCount) return null;
                image = await RenderPageCoreAsync(pdfDoc, pageIndex, targetWidth);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ThumbnailCacheService] 縮圖渲染失敗 {pdfPath} 頁 {pageIndex}: {ex.Message}");
                image = CreateFallbackThumbnail(pageIndex + 1);
            }

            pool.Put(key, image);
            return image;
        }

        private static async Task<ImageSource> RenderPageCoreAsync(PdfDocument pdfDoc, int pageIndex, int targetWidth)
        {
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
            return bitmap;
        }

        /// <summary>
        /// 渲染高解析大圖以供燈箱 (Lightbox) 檢視
        /// </summary>
        public async Task<ImageSource?> GetHighResPreviewAsync(string pdfPath, int pageIndex, int targetWidth = 3840)
        {
            return await GetThumbnailAsync(pdfPath, pageIndex, targetWidth);
        }

        private ImageSource CreateFallbackThumbnail(int pageNumber)
        {
            var visual = new DrawingVisual();
            using (var dc = visual.RenderOpen())
            {
                var bgBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F7F4ED"));
                var borderPen = new Pen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E3DDD2")), 2);
                var textBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#736B63"));

                dc.DrawRectangle(bgBrush, borderPen, new Rect(0, 0, 210, 297));

                var typeface = new Typeface(new FontFamily("Microsoft JhengHei, Segoe UI"), FontStyles.Normal, FontWeights.SemiBold, FontStretches.Normal);
                var text = new FormattedText(
                    $"第 {pageNumber} 頁\n(預覽載入中)",
                    System.Globalization.CultureInfo.CurrentCulture,
                    FlowDirection.LeftToRight,
                    typeface,
                    16,
                    textBrush,
                    1.0
                );
                text.TextAlignment = TextAlignment.Center;
                dc.DrawText(text, new Point(105, 130));
            }

            var rtb = new RenderTargetBitmap(210, 297, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(visual);
            rtb.Freeze();
            return rtb;
        }
    }
}
