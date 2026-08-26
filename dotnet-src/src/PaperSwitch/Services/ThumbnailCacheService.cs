using System;
using System.Collections.Concurrent;
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

        private readonly ConcurrentDictionary<string, ImageSource> _cache = new();

        public void ClearCache()
        {
            _cache.Clear();
        }

        public async Task<ImageSource?> GetThumbnailAsync(string pdfPath, int pageIndex, int targetWidth = 320)
        {
            if (!File.Exists(pdfPath) || pageIndex < 0) return null;

            string key = $"{pdfPath}|{pageIndex}|{targetWidth}";
            if (_cache.TryGetValue(key, out var cached))
            {
                return cached;
            }

            try
            {
                var storageFile = await StorageFile.GetFileFromPathAsync(Path.GetFullPath(pdfPath));
                var pdfDoc = await PdfDocument.LoadFromFileAsync(storageFile);

                if (pageIndex >= (int)pdfDoc.PageCount) return null;

                using var page = pdfDoc.GetPage((uint)pageIndex);
                using var memStream = new InMemoryRandomAccessStream();

                var options = new PdfPageRenderOptions();
                if (targetWidth > 0)
                {
                    options.DestinationWidth = (uint)targetWidth;
                }

                await page.RenderToStreamAsync(memStream, options);

                memStream.Seek(0);
                using var netStream = memStream.AsStreamForRead();

                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.StreamSource = netStream;
                bitmap.EndInit();
                bitmap.Freeze(); // 凍結物件以支援跨 UI 執行緒綁定

                _cache[key] = bitmap;
                return bitmap;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ThumbnailCacheService] 縮圖渲染失敗 {pdfPath} 頁 {pageIndex}: {ex.Message}");
                // 產生優雅的後備預設手帳卡片
                var fallback = CreateFallbackThumbnail(pageIndex + 1);
                _cache[key] = fallback;
                return fallback;
            }
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
