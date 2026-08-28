using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
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

        private const int MaximumCacheEntries = 180;
        private readonly object _cacheLock = new();
        private readonly Dictionary<string, CacheEntry> _cache = new();
        private readonly ConcurrentDictionary<string, Lazy<Task<ImageSource?>>> _inflight = new();

        private sealed class CacheEntry(ImageSource image)
        {
            public ImageSource Image { get; } = image;
            public LinkedListNode<string>? Node { get; set; }
        }

        private readonly LinkedList<string> _leastRecentlyUsed = new();

        public void ClearCache()
        {
            lock (_cacheLock)
            {
                _cache.Clear();
                _leastRecentlyUsed.Clear();
            }
        }

        public async Task<ImageSource?> GetThumbnailAsync(string pdfPath, int pageIndex, int targetWidth = 320)
        {
            if (!File.Exists(pdfPath) || pageIndex < 0) return null;

            string key = $"{pdfPath}|{pageIndex}|{targetWidth}";
            if (TryGetCached(key, out var cached)) return cached;

            var lazyRender = _inflight.GetOrAdd(key, _ => new Lazy<Task<ImageSource?>>(
                () => RenderAndCacheAsync(pdfPath, pageIndex, targetWidth, key)));

            try
            {
                return await lazyRender.Value;
            }
            finally
            {
                _inflight.TryRemove(key, out _);
            }
        }

        private bool TryGetCached(string key, out ImageSource? image)
        {
            lock (_cacheLock)
            {
                if (_cache.TryGetValue(key, out var entry))
                {
                    _leastRecentlyUsed.Remove(entry.Node!);
                    _leastRecentlyUsed.AddFirst(entry.Node!);
                    image = entry.Image;
                    return true;
                }
            }
            image = null;
            return false;
        }

        private async Task<ImageSource?> RenderAndCacheAsync(string pdfPath, int pageIndex, int targetWidth, string key)
        {
            ImageSource image;
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
                bitmap.BeginInit(); bitmap.CacheOption = BitmapCacheOption.OnLoad; bitmap.StreamSource = netStream; bitmap.EndInit(); bitmap.Freeze();
                image = bitmap;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ThumbnailCacheService] 縮圖渲染失敗 {pdfPath} 頁 {pageIndex}: {ex.Message}");
                image = CreateFallbackThumbnail(pageIndex + 1);
            }

            lock (_cacheLock)
            {
                var node = _leastRecentlyUsed.AddFirst(key);
                _cache[key] = new CacheEntry(image) { Node = node };
                while (_cache.Count > MaximumCacheEntries)
                {
                    var oldest = _leastRecentlyUsed.Last!;
                    _cache.Remove(oldest.Value);
                    _leastRecentlyUsed.RemoveLast();
                }
            }
            return image;
        }

        /// <summary>
        /// 渲染高解析大圖以供燈箱 (Lightbox) 檢視
        /// </summary>
        public async Task<ImageSource?> GetHighResPreviewAsync(string pdfPath, int pageIndex, int targetWidth = 1600)
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
