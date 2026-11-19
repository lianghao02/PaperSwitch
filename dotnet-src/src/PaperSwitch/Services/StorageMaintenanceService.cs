using System;
using System.Collections.Generic;
using System.Diagnostics;
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
            // 成品不屬於暫存，不由一般清理刪除。
            ClearDirectoryContents(AppPaths.TemporaryConversionDirectory, result);
            return result;
        }

        /// <summary>
        /// 清除轉檔暫存資料夾，但保留仍被畫布或待處理任務使用的檔案，
        /// 並一併回收已結束或異常終止之舊工作階段遺留目錄（不影響其他執行中的實例）。
        /// </summary>
        public static StorageCleanupResult ClearTemporaryExcept(IEnumerable<string?> keepPaths)
        {
            var result = ClearDirectoryExcept(AppPaths.TemporaryConversionDirectory, keepPaths);
            ClearStaleSessionDirectories(AppPaths.TemporaryRootDirectory, AppPaths.TemporaryConversionDirectory, result);
            return result;
        }

        internal static void ClearStaleSessionDirectories(
            string temporaryRootDirectory,
            string currentSessionDirectory,
            StorageCleanupResult result,
            Func<int, bool>? isPaperSwitchProcessAlive = null)
        {
            if (!Directory.Exists(temporaryRootDirectory)) return;

            string currentFullPath = Path.GetFullPath(currentSessionDirectory)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            var isAliveCheck = isPaperSwitchProcessAlive ?? IsPaperSwitchProcessAlive;

            foreach (var sessionDir in Directory.EnumerateDirectories(temporaryRootDirectory))
            {
                string fullPath = Path.GetFullPath(sessionDir)
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                if (string.Equals(fullPath, currentFullPath, StringComparison.OrdinalIgnoreCase)) continue;

                string folderName = Path.GetFileName(fullPath);
                int dashIndex = folderName.IndexOf('-');
                if (dashIndex > 0 && int.TryParse(folderName.AsSpan(0, dashIndex), out int pid) && isAliveCheck(pid))
                {
                    continue;
                }

                ClearDirectoryContents(fullPath, result);
                try
                {
                    Directory.Delete(fullPath, false);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    result.FailedPaths.Add(fullPath);
                }
            }
        }

        private static bool IsPaperSwitchProcessAlive(int pid)
        {
            if (pid <= 0) return false;
            try
            {
                using var process = Process.GetProcessById(pid);
                if (process.HasExited) return false;
                return process.ProcessName.Contains("PaperSwitch", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

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
