using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using PaperSwitch.Models;

namespace PaperSwitch.Services
{
    /// <summary>
    /// Microsoft Office COM 自動化轉檔服務 (含 STA 隔離、生命週期銷毀與 IGEF 中介偵測)
    /// </summary>
    public class OfficeConverterService
    {
        public static OfficeConverterService Instance { get; } = new();

        public const int DefaultPdfReadyTimeoutSeconds = 90;
        private const int IgefDetectionGracePeriodSeconds = 5;

        private static readonly SemaphoreSlim _comLock = new(1, 1);

        public bool IsWordFile(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return ext is ".docx" or ".doc" or ".dotx" or ".dot" or ".rtf";
        }

        public bool IsExcelFile(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return ext is ".xlsx" or ".xls" or ".xlsm" or ".xlsb" or ".csv";
        }

        public bool IsPowerPointFile(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return ext is ".pptx" or ".ppt" or ".ppsx" or ".pps" or ".potx" or ".pot";
        }

        public bool IsOfficeFile(string path) => IsWordFile(path) || IsExcelFile(path) || IsPowerPointFile(path);

        /// <summary>
        /// 檢查 PDF 檔頭是否為標準 %PDF 或 Office 暫存之 IGEF 格式
        /// </summary>
        public (string State, long Size) CheckPdfHeaderAndSize(string pdfPath)
        {
            if (!File.Exists(pdfPath)) return ("error", 0);
            try
            {
                var fileInfo = new FileInfo(pdfPath);
                long currentSize = fileInfo.Length;
                using var fs = new FileStream(pdfPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                byte[] header = new byte[4];
                int read = fs.Read(header, 0, 4);
                if (read >= 4)
                {
                    if (header[0] == 0x25 && header[1] == 0x50 && header[2] == 0x44 && header[3] == 0x46) // %PDF
                    {
                        return ("pdf", currentSize);
                    }
                    if (header[0] == 0x49 && header[1] == 0x47 && header[2] == 0x45 && header[3] == 0x46) // IGEF
                    {
                        return ("igef", currentSize);
                    }
                }
                return ("unknown", currentSize);
            }
            catch
            {
                return ("error", 0);
            }
        }

        /// <summary>
        /// 等待 Office 背景寫入與加密程序完成，直到 PDF 標頭轉為標準格式
        /// </summary>
        public async Task<bool> WaitForPdfReadyAsync(string pdfPath, int timeoutSeconds = DefaultPdfReadyTimeoutSeconds)
        {
            var deadline = DateTime.UtcNow.AddSeconds(timeoutSeconds);
            long previousSize = -1;
            DateTime? igefDetectedAt = null;

            while (DateTime.UtcNow <= deadline)
            {
                var (state, currentSize) = CheckPdfHeaderAndSize(pdfPath);

                if (state == "pdf" && currentSize > 0 && currentSize == previousSize)
                {
                    try
                    {
                        using var fs = new FileStream(pdfPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                        if (fs.Length > 0) return true;
                    }
                    catch
                    {
                        // 檔案可能仍被短暫鎖定
                    }
                }

                if (state == "igef")
                {
                    igefDetectedAt ??= DateTime.UtcNow;
                    if ((DateTime.UtcNow - igefDetectedAt.Value).TotalSeconds >= IgefDetectionGracePeriodSeconds)
                    {
                        return false;
                    }
                }
                else
                {
                    igefDetectedAt = null;
                }

                if (currentSize > 0)
                {
                    previousSize = currentSize;
                }

                await Task.Delay(250);
            }

            return false;
        }

        private static Task<T> RunInStaThreadAsync<T>(Func<T> action)
        {
            var tcs = new TaskCompletionSource<T>();
            var thread = new Thread(() =>
            {
                try
                {
                    var res = action();
                    tcs.SetResult(res);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
            return tcs.Task;
        }

        /// <summary>
        /// 在專屬 STA 執行緒中執行 Word COM 轉檔
        /// </summary>
        public async Task<ConversionResult> ConvertWordAsync(string inputPath, string outputPath)
        {
            await _comLock.WaitAsync();
            try
            {
                return await RunInStaThreadAsync(() =>
                {
                    dynamic? wordApp = null;
                    dynamic? doc = null;
                    try
                    {
                        var wordType = Type.GetTypeFromProgID("Word.Application");
                        if (wordType == null)
                        {
                            return ConversionResult.Fail(inputPath, "本機未安裝相容的 Microsoft Word，無法進行 Word 轉檔。");
                        }

                        wordApp = Activator.CreateInstance(wordType);
                        if (wordApp is null)
                        {
                            return ConversionResult.Fail(inputPath, "無法啟動 Microsoft Word，請確認 Office 安裝、授權及資訊安全政策。");
                        }

                        wordApp.Visible = false;
                        wordApp.DisplayAlerts = 0; // wdAlertsNone

                        var targetDir = Path.GetDirectoryName(outputPath);
                        if (!string.IsNullOrEmpty(targetDir) && !Directory.Exists(targetDir))
                        {
                            Directory.CreateDirectory(targetDir);
                        }

                        doc = wordApp.Documents.Open(FileName: inputPath, ReadOnly: true, Visible: false);
                        // 17 = wdExportFormatPDF
                        doc.ExportAsFixedFormat(OutputFileName: outputPath, ExportFormat: 17);

                        return ConversionResult.Ok(inputPath, outputPath);
                    }
                    catch (Exception ex)
                    {
                        return ConversionResult.Fail(inputPath, $"Word 轉檔失敗: {ex.Message}", ex.ToString());
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            try { doc.Close(SaveChanges: false); } catch { }
                            try { Marshal.FinalReleaseComObject(doc); } catch { }
                        }
                        if (wordApp != null)
                        {
                            try { wordApp.Quit(SaveChanges: false); } catch { }
                            try { Marshal.FinalReleaseComObject(wordApp); } catch { }
                        }
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                });
            }
            finally
            {
                _comLock.Release();
            }
        }

        /// <summary>
        /// 在專屬 STA 執行緒中執行 Excel COM 轉檔 (支援多分頁獨立拆分與全空白分頁過濾)
        /// </summary>
        public async Task<ConversionResult> ConvertExcelAsync(string inputPath, string outputDir, bool fitToPage = true)
        {
            await _comLock.WaitAsync();
            try
            {
                return await RunInStaThreadAsync(() =>
                {
                    dynamic? excelApp = null;
                    dynamic? workbook = null;
                    var generatedPdfs = new List<string>();

                    try
                    {
                        var excelType = Type.GetTypeFromProgID("Excel.Application");
                        if (excelType == null)
                        {
                            return ConversionResult.Fail(inputPath, "本機未安裝相容的 Microsoft Excel，無法進行 Excel 轉檔。");
                        }

                        excelApp = Activator.CreateInstance(excelType);
                        if (excelApp is null)
                        {
                            return ConversionResult.Fail(inputPath, "無法啟動 Microsoft Excel，請確認 Office 安裝、授權及資訊安全政策。");
                        }

                        excelApp.Visible = false;
                        excelApp.DisplayAlerts = false;

                        if (!Directory.Exists(outputDir))
                        {
                            Directory.CreateDirectory(outputDir);
                        }

                        workbook = excelApp.Workbooks.Open(Filename: inputPath, ReadOnly: true);
                        var stem = Path.GetFileNameWithoutExtension(inputPath);

                        var validSheets = new List<dynamic>();
                        foreach (dynamic ws in workbook.Worksheets)
                        {
                            if (ws.Visible == -1) // -1 = xlSheetVisible
                            {
                                bool hasContent = false;
                                try
                                {
                                    if (ws.Shapes.Count > 0) hasContent = true;
                                    else
                                    {
                                        double count = excelApp.WorksheetFunction.CountA(ws.Cells);
                                        if (count > 0) hasContent = true;
                                    }
                                }
                                catch
                                {
                                    hasContent = true; // 讀取失敗時採保守保留策略
                                }

                                if (hasContent)
                                {
                                    validSheets.Add(ws);
                                }
                            }
                        }

                        if (validSheets.Count == 0)
                        {
                            // 保底：若全部被判定為空，則加入所有可見工作表
                            foreach (dynamic ws in workbook.Worksheets)
                            {
                                if (ws.Visible == -1) validSheets.Add(ws);
                            }
                        }

                        if (validSheets.Count <= 1)
                        {
                            var singlePdfPath = Path.Combine(outputDir, $"{stem}.pdf");
                            if (validSheets.Count == 1)
                            {
                                dynamic ws = validSheets[0];
                                if (fitToPage)
                                {
                                    try
                                    {
                                        ws.PageSetup.Zoom = false;
                                        ws.PageSetup.FitToPagesWide = 1;
                                        ws.PageSetup.FitToPagesTall = false;
                                    }
                                    catch { }
                                }
                                ws.ExportAsFixedFormat(0 /* xlTypePDF */, singlePdfPath);
                            }
                            else
                            {
                                workbook.ExportAsFixedFormat(0 /* xlTypePDF */, singlePdfPath);
                            }
                            generatedPdfs.Add(singlePdfPath);
                        }
                        else
                        {
                            foreach (dynamic ws in validSheets)
                            {
                                string sheetName = ws.Name;
                                string safeSheetName = string.Join("_", sheetName.Split(Path.GetInvalidFileNameChars())).Trim();
                                if (string.IsNullOrEmpty(safeSheetName)) safeSheetName = "工作表";

                                var sheetPdfPath = Path.Combine(outputDir, $"{stem}_{safeSheetName}.pdf");
                                int counter = 1;
                                while (File.Exists(sheetPdfPath))
                                {
                                    sheetPdfPath = Path.Combine(outputDir, $"{stem}_{safeSheetName}_{counter}.pdf");
                                    counter++;
                                }

                                if (fitToPage)
                                {
                                    try
                                    {
                                        ws.PageSetup.Zoom = false;
                                        ws.PageSetup.FitToPagesWide = 1;
                                        ws.PageSetup.FitToPagesTall = false;
                                    }
                                    catch { }
                                }

                                ws.ExportAsFixedFormat(0 /* xlTypePDF */, sheetPdfPath);
                                generatedPdfs.Add(sheetPdfPath);
                            }
                        }

                        return ConversionResult.Ok(inputPath, generatedPdfs.ToArray());
                    }
                    catch (Exception ex)
                    {
                        return ConversionResult.Fail(inputPath, $"Excel 轉檔失敗: {ex.Message}", ex.ToString());
                    }
                    finally
                    {
                        if (workbook != null)
                        {
                            try { workbook.Close(SaveChanges: false); } catch { }
                            try { Marshal.FinalReleaseComObject(workbook); } catch { }
                        }
                        if (excelApp != null)
                        {
                            try { excelApp.Quit(); } catch { }
                            try { Marshal.FinalReleaseComObject(excelApp); } catch { }
                        }
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                });
            }
            finally
            {
                _comLock.Release();
            }
        }

        /// <summary>
        /// 在專屬 STA 執行緒中執行 PowerPoint COM 轉檔
        /// </summary>
        public async Task<ConversionResult> ConvertPowerPointAsync(string inputPath, string outputPath)
        {
            await _comLock.WaitAsync();
            try
            {
                return await RunInStaThreadAsync(() =>
                {
                    dynamic? pptApp = null;
                    dynamic? presentation = null;
                    try
                    {
                        var pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                        if (pptType == null)
                        {
                            return ConversionResult.Fail(inputPath, "本機未安裝相容的 Microsoft PowerPoint，無法進行簡報轉檔。");
                        }

                        pptApp = Activator.CreateInstance(pptType);
                        if (pptApp is null)
                        {
                            return ConversionResult.Fail(inputPath, "無法啟動 Microsoft PowerPoint，請確認 Office 安裝、授權及資訊安全政策。");
                        }

                        var targetDir = Path.GetDirectoryName(outputPath);
                        if (!string.IsNullOrEmpty(targetDir) && !Directory.Exists(targetDir))
                        {
                            Directory.CreateDirectory(targetDir);
                        }

                        // 32 = ppSaveAsPDF, WithWindow: 0 = msoFalse
                        presentation = pptApp.Presentations.Open(inputPath, WithWindow: 0);
                        presentation.SaveAs(outputPath, 32);

                        return ConversionResult.Ok(inputPath, outputPath);
                    }
                    catch (Exception ex)
                    {
                        return ConversionResult.Fail(inputPath, $"PowerPoint 轉檔失敗: {ex.Message}", ex.ToString());
                    }
                    finally
                    {
                        if (presentation != null)
                        {
                            try { presentation.Close(); } catch { }
                            try { Marshal.FinalReleaseComObject(presentation); } catch { }
                        }
                        if (pptApp != null)
                        {
                            try { pptApp.Quit(); } catch { }
                            try { Marshal.FinalReleaseComObject(pptApp); } catch { }
                        }
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                });
            }
            finally
            {
                _comLock.Release();
            }
        }
    }
}
