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
    }
}
