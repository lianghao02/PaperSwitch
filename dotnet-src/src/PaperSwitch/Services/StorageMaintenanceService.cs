using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("PaperSwitch.Tests")]

namespace PaperSwitch.Services
{
    /// <summary>
    /// 負責清理 PaperSwitch 自己建立的使用者資料目錄內容。
    /// 僅清空內容，保留目錄本身，且不會觸及使用者選取的原始檔。
    /// </summary>
    internal static class StorageMaintenanceService
    {
        public static StorageCleanupResult ClearProcessingDirectories()
        {
            var result = new StorageCleanupResult();
            ClearDirectoryContents(AppPaths.ConvertedDirectory, result);
            ClearDirectoryContents(AppPaths.TemporaryConversionDirectory, result);
            return result;
        }

        /// <summary>
        /// 清除轉檔暫存資料夾，但保留仍被畫布或待處理任務使用的檔案。
        /// 匯出完成後呼叫，避免解密或轉出的暫存 PDF 長期留在電腦上。
        /// </summary>
        public static StorageCleanupResult ClearTemporaryExcept(IEnumerable<string?> keepPaths)
            => ClearDirectoryExcept(AppPaths.TemporaryConversionDirectory, keepPaths);

        internal static StorageCleanupResult ClearDirectoryExcept(string directory, IEnumerable<string?> keepPaths)
        {
            var result = new StorageCleanupResult();
            var keep = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var path in keepPaths)
            {
                if (string.IsNullOrWhiteSpace(path)) continue;
                try { keep.Add(Path.GetFullPath(path)); }
                catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException) { }
            }

            foreach (var file in Directory.EnumerateFiles(directory, "*", SearchOption.AllDirectories))
            {
                if (keep.Contains(Path.GetFullPath(file))) continue;
                try
                {
                    File.Delete(file);
                    result.DeletedFileCount++;
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    result.FailedPaths.Add(file);
                }
            }

            return result;
        }

        /// <summary>程式結束時清空轉檔暫存；失敗不影響關閉。</summary>
        public static void ClearTemporaryOnExit()
        {
            try { ClearDirectoryContents(AppPaths.TemporaryConversionDirectory, new StorageCleanupResult()); }
            catch (Exception) { }
        }

        private static void ClearDirectoryContents(string directory, StorageCleanupResult result)
        {
            Directory.CreateDirectory(directory);

            foreach (var file in Directory.EnumerateFiles(directory, "*", SearchOption.AllDirectories))
            {
                try
                {
                    File.Delete(file);
                    result.DeletedFileCount++;
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    result.FailedPaths.Add(file);
                }
            }

            foreach (var subdirectory in Directory.EnumerateDirectories(directory, "*", SearchOption.AllDirectories)
                         .OrderByDescending(path => path.Length))
            {
                try
                {
                    Directory.Delete(subdirectory, false);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    result.FailedPaths.Add(subdirectory);
                }
            }
        }
    }

    internal sealed class StorageCleanupResult
    {
        public int DeletedFileCount { get; set; }

        public List<string> FailedPaths { get; } = new();
    }
}
