using System;
using System.IO;

namespace PaperSwitch.Services
{
    /// <summary>
    /// 集中管理使用者資料路徑，避免建置清理發行目錄時刪除成品與診斷紀錄。
    /// </summary>
    internal static class AppPaths
    {
        private const string ApplicationFolderName = "PaperSwitch";

        public static string DataDirectory => EnsureDirectory(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), ApplicationFolderName));

        public static string ConvertedDirectory => EnsureDirectory(Path.Combine(DataDirectory, "converted"));

        public static string TemporaryConversionDirectory => EnsureDirectory(Path.Combine(DataDirectory, "temp_converted"));

        public static string CrashLogPath => Path.Combine(DataDirectory, "crash.log");

        private static string EnsureDirectory(string path)
        {
            Directory.CreateDirectory(path);
            return path;
        }
    }
}
