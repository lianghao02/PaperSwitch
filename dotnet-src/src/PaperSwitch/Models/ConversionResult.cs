using System;
using System.Collections.Generic;

namespace PaperSwitch.Models
{
    /// <summary>
    /// 檔案轉換結果封裝
    /// </summary>
    public class ConversionResult
    {
        public bool Success { get; set; }
        public string SourcePath { get; set; } = string.Empty;
        public List<string> GeneratedPdfPaths { get; set; } = new();
        public string? ErrorMessage { get; set; }
        public string? DiagnosticInfo { get; set; }

        public static ConversionResult Ok(string sourcePath, params string[] generatedPaths)
        {
            return new ConversionResult
            {
                Success = true,
                SourcePath = sourcePath,
                GeneratedPdfPaths = new List<string>(generatedPaths)
            };
        }

        public static ConversionResult Fail(string sourcePath, string error, string? diagnostic = null)
        {
            return new ConversionResult
            {
                Success = false,
                SourcePath = sourcePath,
                ErrorMessage = error,
                DiagnosticInfo = diagnostic
            };
        }
    }
}
