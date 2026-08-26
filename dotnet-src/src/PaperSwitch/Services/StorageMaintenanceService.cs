using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
