using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PaperSwitch.Models;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PaperSwitch.Services
{
    /// <summary>
    /// 基於 PdfSharp 的無損向量 PDF 處理核心服務
    /// </summary>
    public class PdfService
    {
        public static PdfService Instance { get; } = new();

        private static void DrawPageNumber(PdfPage page, int number, ExportOptions options)
        {
            const double margin = 18;
            var rotation = ((page.Rotate % 360) + 360) % 360;
            // 先在未旋轉的頁面座標繪製，補償閱讀方向，最後保留原本 /Rotate。
            page.Rotate = 0;
            try
            {
                using var graphics = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
                var media = page.MediaBox;
                var crop = page.CropBox;
                var left = Math.Max(media.X1, crop.X1);
                var bottom = Math.Max(media.Y1, crop.Y1);
                var right = Math.Min(media.X2, crop.X2);
                var top = Math.Min(media.Y2, crop.Y2);
                if (right <= left || top <= bottom) { left = media.X1; bottom = media.Y1; right = media.X2; top = media.Y2; }
                var width = right - left;
                var height = top - bottom;
                graphics.TranslateTransform(left - media.X1, media.Y2 - top);
                var visibleWidth = rotation is 90 or 270 ? height : width;
                var visibleHeight = rotation is 90 or 270 ? width : height;
                var matrix = rotation switch
                {
                    90 => new XMatrix(0, -1, 1, 0, 0, height),
                    180 => new XMatrix(-1, 0, 0, -1, width, height),
                    270 => new XMatrix(0, 1, -1, 0, width, 0),
                    _ => XMatrix.Identity
                };
                graphics.MultiplyTransform(matrix);
                var size = double.IsFinite(options.PageNumberFontSize)
                    ? Math.Clamp(options.PageNumberFontSize, 8, 24) : 12;
                var font = new XFont("Arial", size);
                var format = new XStringFormat
                {
                    Alignment = options.PageNumberPosition switch
                    {
                        0 => XStringAlignment.Near,
                        2 => XStringAlignment.Far,
                        _ => XStringAlignment.Center
                    },
                    LineAlignment = XLineAlignment.Far
                };
                graphics.DrawString(number.ToString(System.Globalization.CultureInfo.InvariantCulture), font, XBrushes.Black,
                    new XRect(margin, margin, Math.Max(1, visibleWidth - margin * 2), Math.Max(size * 1.5, visibleHeight - margin * 2)), format);
            }
            finally { page.Rotate = rotation; }
        }

        /// <summary>
        /// 建立一張可由既有無損導出流程處理的標準 A4 空白 PDF。
        /// </summary>
        public bool CreateBlankA4Pdf(string outputPath)
        {
            try
            {
                var directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                using var document = new PdfDocument();
                var page = document.AddPage();
                page.Width = XUnit.FromMillimeter(210);
                page.Height = XUnit.FromMillimeter(297);
                document.Save(outputPath);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PdfService] 建立空白頁失敗 {outputPath}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 取得 PDF 總頁數
        /// </summary>
        public int GetPageCount(string pdfPath)
        {
            if (!File.Exists(pdfPath)) return 0;
            try
            {
                using var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);
                return document.PageCount;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PdfService] 讀取 PDF 頁數失敗 {pdfPath}: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 取得 PDF 指定頁面尺寸與方向資訊
        /// </summary>
        public (double Width, double Height, int Rotate) GetPageDimensions(string pdfPath, int pageIndex)
        {
            if (!File.Exists(pdfPath)) return (595.0, 842.0, 0);
            try
            {
                using var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);
                if (pageIndex >= 0 && pageIndex < document.PageCount)
                {
                    var page = document.Pages[pageIndex];
                    return (page.Width.Point, page.Height.Point, page.Rotate);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PdfService] 讀取 PDF 頁面尺寸失敗 {pdfPath}: {ex.Message}");
            }
            return (595.0, 842.0, 0);
        }

        /// <summary>
        /// 依據使用者在畫布中編排的頁面清單與旋轉角度，進行 100% 向量無損 PDF 合成
        /// </summary>
        public bool ExportArrangedPdf(IEnumerable<PaperItem> items, string outputPath, ExportOptions? options = null)
        {
            var openedDocs = new Dictionary<string, PdfDocument>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var itemList = items.ToList();
                if (itemList.Count == 0) return false;

                var targetDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(targetDir) && !Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }

                using var outputDocument = new PdfDocument();
                foreach (var item in itemList)
                {
                    if (!File.Exists(item.SourceFilePath)) continue;

                    if (!openedDocs.TryGetValue(item.SourceFilePath, out var inputDoc))
                    {
                        inputDoc = PdfReader.Open(item.SourceFilePath, PdfDocumentOpenMode.Import);
                        openedDocs[item.SourceFilePath] = inputDoc;
                    }

                    if (item.SourcePageIndex >= 0 && item.SourcePageIndex < inputDoc.PageCount)
                    {
                        var srcPage = inputDoc.Pages[item.SourcePageIndex];
                        var newPage = outputDocument.AddPage(srcPage);
                        // 計算累加旋轉角度
                        int combinedRotate = (srcPage.Rotate + item.Rotation) % 360;
                        newPage.Rotate = combinedRotate;
                        if (options?.AddPageNumbers == true)
                            DrawPageNumber(newPage, outputDocument.PageCount, options);
                    }
                }

                if (outputDocument.PageCount == 0) return false;

                outputDocument.Save(outputPath);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PdfService] 匯出編排 PDF 失敗: {ex.Message}");
                return false;
            }
            finally
            {
                foreach (var doc in openedDocs.Values)
                {
                    try { doc.Dispose(); } catch { }
                }
            }
        }

        /// <summary>
        /// 將多頁項目個別獨立匯出為單一頁面 PDF 檔案
        /// </summary>
        public List<string> ExportIndividualPdfs(IEnumerable<PaperItem> items, string outputDir, string? fileNamePrefix = null, bool includeDescription = false, ExportOptions? options = null)
        {
            var results = new List<string>();
            var itemList = items.ToList();
            if (itemList.Count == 0) return results;

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            var openedDocs = new Dictionary<string, PdfDocument>(StringComparer.OrdinalIgnoreCase);
            DateTime exportTime = DateTime.Now;


            try
            {
                int index = 1;
                foreach (var item in itemList)
                {
                    if (!File.Exists(item.SourceFilePath)) continue;

                    if (!openedDocs.TryGetValue(item.SourceFilePath, out var inputDoc))
                    {
                        inputDoc = PdfReader.Open(item.SourceFilePath, PdfDocumentOpenMode.Import);
                        openedDocs[item.SourceFilePath] = inputDoc;
                    }

                    if (item.SourcePageIndex >= 0 && item.SourcePageIndex < inputDoc.PageCount)
                    {
                        var srcPage = inputDoc.Pages[item.SourcePageIndex];
                        using var singleDoc = new PdfDocument();
                        var newPage = singleDoc.AddPage(srcPage);
                        newPage.Rotate = (srcPage.Rotate + item.Rotation) % 360;
                        if (options?.AddPageNumbers == true)
                            DrawPageNumber(newPage, index, options);

                        string fileName = ExportFileNaming.Page(ExportFileNaming.SourceName(item), item.DisplayPageNumber, index,
                            fileNamePrefix, ".pdf", includeDescription, exportTime);
                        string targetPath = ImageExportService.GetUniqueFilePath(outputDir, fileName);
                        singleDoc.Save(targetPath);
                        results.Add(targetPath);
                        index++;
                    }
                }
            }
            finally
            {
                foreach (var doc in openedDocs.Values)
                {
                    try { doc.Dispose(); } catch { }
                }
            }

            return results;
        }

        /// <summary>
        /// 合併多個 PDF 檔案為單一 PDF
        /// </summary>
        public bool MergePdfs(IEnumerable<string> pdfPaths, string outputPath)
        {
            try
            {
                var pathList = pdfPaths.Where(File.Exists).ToList();
                if (pathList.Count == 0) return false;

                var targetDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(targetDir) && !Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }

                using var outputDocument = new PdfDocument();
                foreach (var path in pathList)
                {
                    try
                    {
                        using var inputDoc = PdfReader.Open(path, PdfDocumentOpenMode.Import);
                        for (int i = 0; i < inputDoc.PageCount; i++)
                        {
                            outputDocument.AddPage(inputDoc.Pages[i]);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[PdfService] 合併時讀取 {path} 失敗: {ex.Message}");
                    }
                }

                if (outputDocument.PageCount == 0) return false;
                outputDocument.Save(outputPath);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PdfService] 合併 PDF 失敗: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 將多頁 PDF 拆分為單頁 PDF
        /// </summary>
        public List<string> SplitPdf(string inputPdfPath, string outputDir)
        {
            var results = new List<string>();
            if (!File.Exists(inputPdfPath)) return results;

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            var stem = Path.GetFileNameWithoutExtension(inputPdfPath);

            using var inputDoc = PdfReader.Open(inputPdfPath, PdfDocumentOpenMode.Import);
            for (int i = 0; i < inputDoc.PageCount; i++)
            {
                using var singleDoc = new PdfDocument();
                singleDoc.AddPage(inputDoc.Pages[i]);

                var fileName = inputDoc.PageCount == 1 ? $"{stem}.pdf" : $"{stem}_第{i + 1}頁.pdf";
                var targetPath = Path.Combine(outputDir, fileName);

                int counter = 1;
                while (File.Exists(targetPath))
                {
                    targetPath = Path.Combine(outputDir, $"{stem}_第{i + 1}頁_{counter}.pdf");
                    counter++;
                }

                singleDoc.Save(targetPath);
                results.Add(targetPath);
            }

            return results;
        }
    }
}
