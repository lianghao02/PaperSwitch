using System;
using System.IO;
using System.Linq;
using PaperSwitch.Models;
using PaperSwitch.Services;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using Xunit;

namespace PaperSwitch.Tests
{
    public class PdfServiceTests : IDisposable
    {
        private readonly string _testDir;
        private readonly PdfService _pdfService = PdfService.Instance;

        public PdfServiceTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "PaperSwitch_Tests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_testDir);
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_testDir))
                {
                    Directory.Delete(_testDir, true);
                }
            }
            catch { }
        }

        private string CreateDummyPdf(string fileName, int pageCount = 3)
        {
            string path = Path.Combine(_testDir, fileName);
            using var doc = new PdfDocument();
            for (int i = 0; i < pageCount; i++)
            {
                var page = doc.AddPage();
                page.Width = PdfSharp.Drawing.XUnit.FromPoint(595);
                page.Height = PdfSharp.Drawing.XUnit.FromPoint(842);
            }
            doc.Save(path);
            return path;
        }

        private string CreateSizedPdf(string fileName, double widthPoints, double heightPoints)
        {
            string path = Path.Combine(_testDir, fileName);
            using var doc = new PdfDocument();
            var page = doc.AddPage();
            page.Width = XUnit.FromPoint(widthPoints);
            page.Height = XUnit.FromPoint(heightPoints);
            doc.Save(path);
            return path;
        }

        [Fact]
        public void CreateBlankA4Pdf_ShouldCreateSinglePortraitA4Page()
        {
            string output = Path.Combine(_testDir, "blank-a4.pdf");

            bool ok = _pdfService.CreateBlankA4Pdf(output);
            var (width, height, rotate) = _pdfService.GetPageDimensions(output, 0);

            Assert.True(ok);
            Assert.Equal(1, _pdfService.GetPageCount(output));
            Assert.InRange(width, 594, 596);
            Assert.InRange(height, 841, 843);
            Assert.Equal(0, rotate);
        }

        [Theory]
        [InlineData(420, 595)]
        [InlineData(595, 420)]
        [InlineData(20000, 20000)]
        public void GetPageDimensions_ShouldPreserveSupportedPageBounds(double width, double height)
        {
            string path = CreateSizedPdf($"size-{width}-{height}.pdf", width, height);

            var dimensions = _pdfService.GetPageDimensions(path, 0);

            Assert.Equal(width, dimensions.Width, 2);
            Assert.Equal(height, dimensions.Height, 2);
        }

        [Fact]
        public void GetPageCount_ShouldReturnAccurateCount()
        {
            string pdfPath = CreateDummyPdf("test_3pages.pdf", 3);
            int count = _pdfService.GetPageCount(pdfPath);
            Assert.Equal(3, count);
        }

        [Fact]
        public void MergePdfs_ShouldMergeMultiplePdfsIntoOne()
        {
            string pdf1 = CreateDummyPdf("pdf1.pdf", 2);
            string pdf2 = CreateDummyPdf("pdf2.pdf", 3);
            string outputMerged = Path.Combine(_testDir, "merged.pdf");

            bool ok = _pdfService.MergePdfs(new[] { pdf1, pdf2 }, outputMerged);

            Assert.True(ok);
            Assert.True(File.Exists(outputMerged));
            Assert.Equal(5, _pdfService.GetPageCount(outputMerged));
        }

        [Fact]
        public void SplitPdf_ShouldSplitIntoSinglePagePdfs()
        {
            string pdf = CreateDummyPdf("multi.pdf", 4);
            string splitDir = Path.Combine(_testDir, "splits");

            var results = _pdfService.SplitPdf(pdf, splitDir);

            Assert.Equal(4, results.Count);
            foreach (var file in results)
            {
                Assert.True(File.Exists(file));
                Assert.Equal(1, _pdfService.GetPageCount(file));
            }
        }

        [Fact]
        public void ExportArrangedPdf_ShouldReorderAndApplyRotation()
        {
            string pdf = CreateDummyPdf("source.pdf", 3);
            string arrangedOutput = Path.Combine(_testDir, "arranged.pdf");

            var items = new[]
            {
                new PaperItem { SourceFilePath = pdf, SourceFileName = "source.pdf", SourcePageIndex = 2, Rotation = 90 },
                new PaperItem { SourceFilePath = pdf, SourceFileName = "source.pdf", SourcePageIndex = 0, Rotation = 180 },
                new PaperItem { SourceFilePath = pdf, SourceFileName = "source.pdf", SourcePageIndex = 1, Rotation = 270 }
            };

            bool ok = _pdfService.ExportArrangedPdf(items, arrangedOutput);

            Assert.True(ok);
            Assert.True(File.Exists(arrangedOutput));
            Assert.Equal(3, _pdfService.GetPageCount(arrangedOutput));

            var (w1, h1, rot1) = _pdfService.GetPageDimensions(arrangedOutput, 0);
            var (w2, h2, rot2) = _pdfService.GetPageDimensions(arrangedOutput, 1);
            var (w3, h3, rot3) = _pdfService.GetPageDimensions(arrangedOutput, 2);

            Assert.Equal(90, rot1);
            Assert.Equal(180, rot2);
            Assert.Equal(270, rot3);
        }

        [Fact]
        public void ExportArrangedPdf_EmptyOrMissingSources_ShouldNotCreateOutput()
        {
            string emptyOutput = Path.Combine(_testDir, "empty.pdf");
            string missingOutput = Path.Combine(_testDir, "missing.pdf");

            Assert.False(_pdfService.ExportArrangedPdf(Array.Empty<PaperItem>(), emptyOutput));
            Assert.False(_pdfService.ExportArrangedPdf(new[]
            {
                new PaperItem { SourceFilePath = Path.Combine(_testDir, "not-found.pdf"), SourcePageIndex = 0 }
            }, missingOutput));
            Assert.False(File.Exists(emptyOutput));
            Assert.False(File.Exists(missingOutput));
        }

        [Fact]
        public void ExportArrangedPdf_UnwritableOutputPath_ShouldReturnFalse()
        {
            string source = CreateDummyPdf("source-for-blocked-output.pdf", 1);
            string blockedDirectory = Path.Combine(_testDir, "blocked");
            File.WriteAllText(blockedDirectory, "此路徑刻意建立為檔案");
            string output = Path.Combine(blockedDirectory, "output.pdf");

            bool ok = _pdfService.ExportArrangedPdf(new[]
            {
                new PaperItem { SourceFilePath = source, SourcePageIndex = 0 }
            }, output);

            Assert.False(ok);
        }

        [Fact]
        public void ExportIndividualPdfs_ShouldPreserveOrderRotationAndPrefix()
        {
            string source = CreateDummyPdf("individual-source.pdf", 2);
            string outputDir = Path.Combine(_testDir, "individual");
            var items = new[]
            {
                new PaperItem { SourceFilePath = source, SourcePageIndex = 1, DisplayPageNumber = 2, Rotation = 90 },
                new PaperItem { SourceFilePath = source, SourcePageIndex = 0, DisplayPageNumber = 1, Rotation = 180 }
            };

            var results = _pdfService.ExportIndividualPdfs(items, outputDir, "測試輸出.pdf");

            Assert.Equal(2, results.Count);
            Assert.EndsWith("測試輸出_001.pdf", results[0]);
            Assert.EndsWith("測試輸出_002.pdf", results[1]);
            Assert.All(results, path => Assert.Equal(1, _pdfService.GetPageCount(path)));
            Assert.Equal(90, _pdfService.GetPageDimensions(results[0], 0).Rotate);
            Assert.Equal(180, _pdfService.GetPageDimensions(results[1], 0).Rotate);
        }

        [Fact]
        public void ExportArrangedPdf_SelectedSubset_ExportsOnlySelectedPages()
        {
            string pdf = CreateDummyPdf("source_subset.pdf", 5);
            string subsetOutput = Path.Combine(_testDir, "subset.pdf");

            var allItems = Enumerable.Range(0, 5).Select(i => new PaperItem
            {
                SourceFilePath = pdf,
                SourceFileName = "source_subset.pdf",
                SourcePageIndex = i,
                Rotation = i == 3 ? 90 : 0,
                IsSelected = i is 1 or 3 // 選取第 2、4 頁 (索引 1, 3)
            }).ToList();

            var selectedItems = allItems.Where(item => item.IsSelected).ToList();
            bool ok = _pdfService.ExportArrangedPdf(selectedItems, subsetOutput);

            Assert.True(ok);
            Assert.True(File.Exists(subsetOutput));
            Assert.Equal(2, _pdfService.GetPageCount(subsetOutput));

            var (_, _, rot0) = _pdfService.GetPageDimensions(subsetOutput, 0);
            var (_, _, rot1) = _pdfService.GetPageDimensions(subsetOutput, 1);
            Assert.Equal(0, rot0);
            Assert.Equal(90, rot1);
        }

        [Fact]
        public void ExportIndividualPdfs_SelectedSubset_SplitsOnlySelectedPages()
        {
            string pdf = CreateDummyPdf("source_split_subset.pdf", 5);
            string splitDir = Path.Combine(_testDir, "split_subset");

            var allItems = Enumerable.Range(0, 5).Select(i => new PaperItem
            {
                SourceFilePath = pdf,
                SourceFileName = "source_split_subset.pdf",
                SourcePageIndex = i,
                DisplayPageNumber = i + 1,
                Rotation = i == 4 ? 180 : 0,
                IsSelected = i is 1 or 4 // 選取第 2、5 頁
            }).ToList();

            var selectedItems = allItems.Where(item => item.IsSelected).ToList();
            var results = _pdfService.ExportIndividualPdfs(selectedItems, splitDir, "選取拆分");

            Assert.Equal(2, results.Count);
            Assert.All(results, path => Assert.True(File.Exists(path)));
            Assert.All(results, path => Assert.Equal(1, _pdfService.GetPageCount(path)));
            Assert.EndsWith("選取拆分_001.pdf", results[0]);
            Assert.EndsWith("選取拆分_002.pdf", results[1]);
            Assert.Equal(0, _pdfService.GetPageDimensions(results[0], 0).Rotate);
            Assert.Equal(180, _pdfService.GetPageDimensions(results[1], 0).Rotate);
        }
    }
}
