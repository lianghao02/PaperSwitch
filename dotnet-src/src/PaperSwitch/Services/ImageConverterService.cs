using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Media.Imaging;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace PaperSwitch.Services
{
    /// <summary>
    /// 圖片無損封裝與轉 PDF 服務
    /// </summary>
    public class ImageConverterService
    {
        public static ImageConverterService Instance { get; } = new();

        private static readonly HashSet<string> SupportedImageExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".png", ".jpg", ".jpeg", ".bmp", ".webp", ".gif", ".tif", ".tiff", ".ico"
        };

        public bool IsImageFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            var ext = Path.GetExtension(path);
            return SupportedImageExtensions.Contains(ext);
        }

        /// <summary>
        /// 將單一或多個圖片檔案轉換/封裝為單一 PDF
        /// </summary>
        public bool ConvertImagesToPdf(IEnumerable<string> imagePaths, string outputPdfPath)
        {
            var validImages = imagePaths.Where(p => File.Exists(p) && IsImageFile(p)).ToList();
            if (validImages.Count == 0) return false;

            var targetDir = Path.GetDirectoryName(outputPdfPath);
            if (!string.IsNullOrEmpty(targetDir) && !Directory.Exists(targetDir))
            {
                Directory.CreateDirectory(targetDir);
            }

            try
            {
                using var document = new PdfDocument();
                var imageResources = new List<(MemoryStream Stream, XImage Image)>();

                try
                {
                    foreach (var imgPath in validImages)
                    {
                        try
                        {
                            // 使用 MemoryStream 封裝圖片，避免檔案鎖死
                            byte[] bytes = File.ReadAllBytes(imgPath);
                            using var ms = new MemoryStream(bytes);
                            
                            // 透過 WPF Decoder 取得正確的寬高與方向
                            var decoder = BitmapDecoder.Create(ms, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                            var frame = decoder.Frames[0];

                        double pixelWidth = frame.PixelWidth;
                        double pixelHeight = frame.PixelHeight;

                        // 依據 96 DPI 換算為 PDF 點數 (Point: 72 points per inch)
                        double dpiX = frame.DpiX > 0 ? frame.DpiX : 96.0;
                        double dpiY = frame.DpiY > 0 ? frame.DpiY : 96.0;

                        double ptWidth = (pixelWidth / dpiX) * 72.0;
                        double ptHeight = (pixelHeight / dpiY) * 72.0;

                        var page = document.AddPage();
                        page.Width = XUnit.FromPoint(ptWidth);
                        page.Height = XUnit.FromPoint(ptHeight);

                        using var gfx = XGraphics.FromPdfPage(page);
                        
                        // 將所有 WPF 可解碼格式先正規化為 PNG，避免 PdfSharp 直接讀取來源格式
                        // （尤其 WebP、GIF、部分螢幕擷圖）時產生空白 PDF 頁面。
                        var pngEncoder = new PngBitmapEncoder();
                        pngEncoder.Frames.Add(frame);
                        using var normalizedPng = new MemoryStream();
                        pngEncoder.Save(normalizedPng);
                        byte[] pngBytes = normalizedPng.ToArray();

                            // PdfSharp 會在 document.Save 時才讀取影像資料，必須保留來源串流與 XImage 至存檔完成。
                            var normalizedPngRead = new MemoryStream(pngBytes.Length);
                            normalizedPngRead.Write(pngBytes, 0, pngBytes.Length);
                            normalizedPngRead.Position = 0;
                            var xImage = XImage.FromStream(normalizedPngRead);
                            imageResources.Add((normalizedPngRead, xImage));
                            gfx.DrawImage(xImage, 0, 0, ptWidth, ptHeight);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ImageConverterService] 處理圖片失敗 {imgPath}: {ex.Message}");
                        }
                    }

                    if (document.PageCount == 0) return false;

                    document.Save(outputPdfPath);
                    return true;
                }
                finally
                {
                    foreach (var resource in imageResources)
                    {
                        resource.Image.Dispose();
                        resource.Stream.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ImageConverterService] 匯出圖片 PDF 失敗: {ex.Message}");
                return false;
            }
        }
    }
}
