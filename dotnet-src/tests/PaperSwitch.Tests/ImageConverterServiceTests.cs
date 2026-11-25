using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using PaperSwitch.Services;
using Xunit;

namespace PaperSwitch.Tests
{
    public class ImageConverterServiceTests : IDisposable
    {
        private readonly string _testDirectory = Path.Combine(Path.GetTempPath(), $"PaperSwitch_ImageTests_{Guid.NewGuid():N}");

        public ImageConverterServiceTests()
        {
            Directory.CreateDirectory(_testDirectory);
        }

        public void Dispose()
        {
            if (Directory.Exists(_testDirectory))
            {
                Directory.Delete(_testDirectory, true);
            }
        }

        [Fact]
        public async Task ConvertImagesToPdf_PngInput_CreatesPreviewablePdf()
        {
            string imagePath = Path.Combine(_testDirectory, "sample.png");
            string pdfPath = Path.Combine(_testDirectory, "sample.pdf");
            byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScL1xQAAAABJRU5ErkJggg==");
            File.WriteAllBytes(imagePath, png);

            bool converted = ImageConverterService.Instance.ConvertImagesToPdf(new[] { imagePath }, pdfPath);
            var thumbnail = await ThumbnailCacheService.Instance.GetThumbnailAsync(pdfPath, 0, 120);

            Assert.True(converted);
            Assert.True(File.Exists(pdfPath));
            Assert.Equal(1, PdfService.Instance.GetPageCount(pdfPath));
            Assert.NotNull(thumbnail);
            var bitmapSource = Assert.IsAssignableFrom<BitmapSource>(thumbnail);
            Assert.True(bitmapSource.PixelWidth >= 120);
        }

        [Fact]
        public void ConvertImagesToPdf_MultiPageTiff_CreatesAllPages()
        {
            string tiffPath = Path.Combine(_testDirectory, "scan_multi.tiff");
            string pdfPath = Path.Combine(_testDirectory, "scan_multi.pdf");

            var encoder = new TiffBitmapEncoder();
            for (int i = 0; i < 3; i++)
            {
                var wb = new WriteableBitmap(64, 96, 96, 96, System.Windows.Media.PixelFormats.Bgra32, null);
                encoder.Frames.Add(BitmapFrame.Create(wb));
            }
            using (var fs = new FileStream(tiffPath, FileMode.Create, FileAccess.Write))
            {
                encoder.Save(fs);
            }

            bool converted = ImageConverterService.Instance.ConvertImagesToPdf(new[] { tiffPath }, pdfPath);

            Assert.True(converted);
            Assert.Equal(3, PdfService.Instance.GetPageCount(pdfPath));
        }
    }
}
