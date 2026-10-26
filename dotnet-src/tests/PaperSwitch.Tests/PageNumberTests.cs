using System.IO;
using PaperSwitch.Models;
using PaperSwitch.Services;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Xunit;

namespace PaperSwitch.Tests;

public sealed class PageNumberTests
{
    [Theory]
    [InlineData(0, 8)]
    [InlineData(1, 12)]
    [InlineData(2, 24)]
    public async Task NumberedExport_RotationsAndSizes_RenderAtReadingBottom(int position, double fontSize)
    {
        var directory = Path.Combine(Path.GetTempPath(), "PaperSwitch-number-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var source = Path.Combine(directory, "source.pdf");
            using (var document = new PdfDocument())
            {
                foreach (var (width, height) in new[] { (420, 595), (842, 595) })
                    for (var rotation = 0; rotation < 360; rotation += 90)
                    {
                        var page = document.AddPage();
                        page.Width = XUnit.FromPoint(width);
                        page.Height = XUnit.FromPoint(height);
                        page.CropBox = new PdfRectangle(new XPoint(24, 36), new XPoint(width - 30, height - 42));
                        page.Rotate = rotation;
                    }
                document.Save(source);
            }
            var original = File.ReadAllBytes(source);
            var items = Enumerable.Range(0, 8).Select(index => new PaperItem
            {
                SourceFilePath = source, SourceFileName = "source.pdf", SourcePageIndex = index,
                Rotation = 90, DisplayPageNumber = index + 1
            }).ToArray();
            var options = new ExportOptions { AddPageNumbers = true, PageNumberPosition = position, PageNumberFontSize = fontSize };
            var output = Path.Combine(directory, "numbered.pdf");
            Assert.True(PdfService.Instance.ExportArrangedPdf(items, output, options));
            using (var document = PdfReader.Open(output, PdfDocumentOpenMode.Import))
            {
                Assert.Equal(8, document.PageCount);
                for (var index = 0; index < 8; index++) Assert.Equal(((index % 4 + 1) * 90) % 360, document.Pages[index].Rotate);
            }
            // 真正經 WinRT 渲染；白紙上的非白像素只能來自新增頁碼。
            for (var index = 0; index < 8; index++)
            {
                var bitmap = await ImageExportService.Instance.RenderPageToBitmapAsync(output, index, 0, 842);
                Assert.NotNull(bitmap);
                var converted = new System.Windows.Media.Imaging.FormatConvertedBitmap(bitmap, System.Windows.Media.PixelFormats.Bgra32, null, 0);
                var width = converted.PixelWidth;
                var height = converted.PixelHeight;
                var pixels = new byte[width * height * 4];
                converted.CopyPixels(pixels, width * 4, 0);
                var points = new List<(int X, int Y)>();
                for (var y = 0; y < height; y++)
                    for (var x = 0; x < width; x++)
                        if (pixels[(y * width + x) * 4] < 150) points.Add((x, y));
                Assert.NotEmpty(points);
                Assert.True(points.Min(point => point.Y) > height * .85, $"第 {index + 1} 頁頁碼未在底部");
                var center = points.Average(point => point.X) / width;
                Assert.InRange(center, position switch { 0 => 0, 2 => .85, _ => .45 }, position switch { 0 => .15, 2 => 1, _ => .55 });
            }
            var split = PdfService.Instance.ExportIndividualPdfs(items.Take(3), Path.Combine(directory, "split"), options: options);
            Assert.Equal(3, split.Count);
            for (var index = 0; index < split.Count; index++)
            {
                var combined = await ImageExportService.Instance.RenderPageToBitmapAsync(output, index, 0, 842);
                var individual = await ImageExportService.Instance.RenderPageToBitmapAsync(split[index], 0, 0, 842);
                Assert.NotNull(combined);
                Assert.NotNull(individual);
                Assert.Equal(combined.PixelWidth, individual.PixelWidth);
                Assert.Equal(combined.PixelHeight, individual.PixelHeight);
                var first = new byte[combined.PixelWidth * combined.PixelHeight * 4];
                var second = new byte[first.Length];
                combined.CopyPixels(first, combined.PixelWidth * 4, 0);
                individual.CopyPixels(second, individual.PixelWidth * 4, 0);
                Assert.Equal(first, second);
            }
            Assert.Equal(original, File.ReadAllBytes(source));
            // 保留可選 QA 證據，方便檢視實際字級與連續編號。
            if (Environment.GetEnvironmentVariable("PAPERSWITCH_SNAPSHOT_DIR") is { Length: > 0 } snapshots)
            {
                Directory.CreateDirectory(snapshots);
                File.Copy(output, Path.Combine(snapshots, $"page-numbers-{position}-{fontSize}.pdf"), true);
                for (var index = 0; index < split.Count; index++)
                    File.Copy(split[index], Path.Combine(snapshots, $"split-{position}-{index + 1}.pdf"), true);
            }
        }
        finally { Directory.Delete(directory, true); }
    }
}
