using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading.Tasks;

namespace PaperSwitch.Services
{
    /// <summary>
    /// 更新檢查結果模型
    /// </summary>
    public class UpdateCheckResult
    {
        public bool Success { get; set; }
        public bool HasUpdate { get; set; }
        public string CurrentVersion { get; set; } = string.Empty;
        public string LatestVersion { get; set; } = string.Empty;
        public string ReleaseNotes { get; set; } = string.Empty;
        public string HtmlUrl { get; set; } = string.Empty;
        public string AssetName { get; set; } = string.Empty;
        public string AssetDownloadUrl { get; set; } = string.Empty;
        public string AssetSha256 { get; set; } = string.Empty;
        public string? ErrorMessage { get; set; }
    }

    /// <summary>
    /// GitHub Release 最新版本檢查服務
    /// </summary>
    public class UpdateService
    {
        private static readonly Lazy<UpdateService> _instance = new(() => new UpdateService());
        public static UpdateService Instance => _instance.Value;

        private const string DefaultRepoOwner = "lianghao02";
        private const string DefaultRepoName = "PaperSwitch";
        private static readonly HttpClient _httpClient = new()
        {
            Timeout = TimeSpan.FromSeconds(8)
        };

        static UpdateService()
        {
            // GitHub API 強制要求合法的 User-Agent header
            if (!_httpClient.DefaultRequestHeaders.Contains("User-Agent"))
            {
                _httpClient.DefaultRequestHeaders.Add("User-Agent", "PaperSwitch-WPF-Client");
            }
        }

        /// <summary>
        /// 讀取本機目前版本字串 (例如 "v4.0.0")
        /// </summary>
        public string GetCurrentVersion()
        {
            try
            {
                // 優先讀取應用程式執行檔同目錄的 version.txt
                var appDir = AppDomain.CurrentDomain.BaseDirectory;
                var versionFilePath = Path.Combine(appDir, "version.txt");
                if (File.Exists(versionFilePath))
                {
                    var text = File.ReadAllText(versionFilePath).Trim();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        return text.StartsWith("v", StringComparison.OrdinalIgnoreCase) ? text : $"v{text}";
                    }
                }

                // 備援讀取專案 Assembly Version
                var asm = Assembly.GetExecutingAssembly();
                var ver = asm.GetName().Version;
                if (ver != null)
                {
                    return $"v{ver.Major}.{ver.Minor}.{ver.Build}";
                }
            }
            catch
            {
                // 忽略讀取例外，採用預設版號
            }

            return "v4.0.0";
        }

        /// <summary>
        /// 比對兩個版本字串 (例如 "v4.0.0" vs "v4.1.0")，若最新版高於目前版則回傳 true
        /// </summary>
        public static bool IsNewerVersion(string currentVerStr, string latestVerStr)
        {
            var cleanCurrent = currentVerStr.TrimStart('v', 'V', ' ');
            var cleanLatest = latestVerStr.TrimStart('v', 'V', ' ');

            if (Version.TryParse(cleanCurrent, out var cur) && Version.TryParse(cleanLatest, out var lat))
            {
                return lat > cur;
            }

            // 備援字串比較
            return !string.Equals(cleanCurrent, cleanLatest, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 非同步向 GitHub Release API 檢查是否有新版本
        /// </summary>
        public async Task<UpdateCheckResult> CheckForUpdatesAsync(string repoOwner = DefaultRepoOwner, string repoName = DefaultRepoName)
        {
            var currentVersion = GetCurrentVersion();
            var apiUrl = $"https://api.github.com/repos/{repoOwner}/{repoName}/releases/latest";

            try
            {
                using var response = await _httpClient.GetAsync(apiUrl);
                if (!response.IsSuccessStatusCode)
                {
                    var statusCode = (int)response.StatusCode;
                    return new UpdateCheckResult
                    {
                        Success = false,
                        CurrentVersion = currentVersion,
                        ErrorMessage = $"GitHub 伺服器回應狀態碼: {statusCode} ({response.ReasonPhrase})"
                    };
                }

                var json = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                var tagName = root.TryGetProperty("tag_name", out var tagProp) ? tagProp.GetString() ?? string.Empty : string.Empty;
                var releaseNotes = root.TryGetProperty("body", out var bodyProp) ? bodyProp.GetString() ?? string.Empty : string.Empty;
                var htmlUrl = root.TryGetProperty("html_url", out var urlProp) ? urlProp.GetString() ?? string.Empty : $"https://github.com/{repoOwner}/{repoName}/releases/latest";

                if (string.IsNullOrWhiteSpace(tagName))
                {
                    return new UpdateCheckResult
                    {
                        Success = false,
                        CurrentVersion = currentVersion,
                        ErrorMessage = "無法從 GitHub 回應中取得版本標籤 (tag_name)。"
                    };
                }

                var hasUpdate = IsNewerVersion(currentVersion, tagName);
                var standaloneAsset = FindStandaloneAsset(root);

                return new UpdateCheckResult
                {
                    Success = true,
                    HasUpdate = hasUpdate,
                    CurrentVersion = currentVersion,
                    LatestVersion = tagName.StartsWith("v", StringComparison.OrdinalIgnoreCase) ? tagName : $"v{tagName}",
                    ReleaseNotes = releaseNotes,
                    HtmlUrl = htmlUrl,
                    AssetName = standaloneAsset.Name,
                    AssetDownloadUrl = standaloneAsset.DownloadUrl,
                    AssetSha256 = standaloneAsset.Sha256
                };
            }
            catch (TaskCanceledException)
            {
                return new UpdateCheckResult
                {
                    Success = false,
                    CurrentVersion = currentVersion,
                    ErrorMessage = "連線逾時，請確認網路連線是否通暢。"
                };
            }
            catch (HttpRequestException ex)
            {
                return new UpdateCheckResult
                {
                    Success = false,
                    CurrentVersion = currentVersion,
                    ErrorMessage = $"網路連線異常: {ex.Message}"
                };
            }
            catch (Exception ex)
            {
                return new UpdateCheckResult
                {
                    Success = false,
                    CurrentVersion = currentVersion,
                    ErrorMessage = $"檢查更新時發生未預期錯誤: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// 下載已發行的 Standalone EXE、驗證 GitHub 提供的 SHA-256，並交由獨立命令稿
        /// 在主程式結束後完成原子替換與重新啟動。執行中的 EXE 無法自行覆寫，故不可省略更新器。
        /// </summary>
        public async Task<UpdateApplyResult> DownloadAndApplyAsync(UpdateCheckResult update)
        {
            if (string.IsNullOrWhiteSpace(update.AssetDownloadUrl) || string.IsNullOrWhiteSpace(update.AssetName))
            {
                return UpdateApplyResult.Failed("此版本未附可自動更新的 Standalone EXE，請改用 Release 頁面手動下載。");
            }

            if (string.IsNullOrWhiteSpace(update.AssetSha256))
            {
                return UpdateApplyResult.Failed("此 Release 未提供 SHA-256 雜湊，為避免下載檔遭竄改，已停止自動替換。請改用 Release 頁面手動下載。");
            }

            string? currentExe = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
            if (string.IsNullOrWhiteSpace(currentExe) || !File.Exists(currentExe))
            {
                return UpdateApplyResult.Failed("無法確認目前執行檔位置，無法安全自動更新。");
            }

            string? currentDirectory = Path.GetDirectoryName(currentExe);
            if (string.IsNullOrWhiteSpace(currentDirectory) || !CanWriteToDirectory(currentDirectory))
            {
                return UpdateApplyResult.Failed("目前程式所在資料夾沒有寫入權限，無法自動替換。請下載新版後手動覆蓋執行檔。");
            }

            try
            {
                string updateDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PaperSwitch", "updates");
                Directory.CreateDirectory(updateDirectory);

                string downloadedExe = Path.Combine(updateDirectory, $"{Path.GetFileNameWithoutExtension(update.AssetName)}_{Guid.NewGuid():N}.exe");
                await using (var source = await _httpClient.GetStreamAsync(update.AssetDownloadUrl))
                await using (var destination = new FileStream(downloadedExe, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1024 * 128, useAsync: true))
                {
                    await source.CopyToAsync(destination);
                }

                string actualHash = await CalculateSha256Async(downloadedExe);
                if (!string.Equals(actualHash, NormalizeSha256(update.AssetSha256), StringComparison.OrdinalIgnoreCase))
                {
                    File.Delete(downloadedExe);
                    return UpdateApplyResult.Failed("新版檔案的 SHA-256 驗證失敗，已刪除下載檔並保留目前版本。");
                }

                string updaterPath = Path.Combine(updateDirectory, $"apply_update_{Guid.NewGuid():N}.cmd");
                File.WriteAllText(updaterPath, BuildUpdaterScript(
                    Environment.ProcessId,
                    currentExe,
                    downloadedExe,
                    update.LatestVersion), new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

                Process.Start(new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c \"{updaterPath}\"",
                    WorkingDirectory = updateDirectory,
                    UseShellExecute = false,
                    CreateNoWindow = true
                });

                return UpdateApplyResult.Started("新版已下載並完成 SHA-256 驗證，程式關閉後會自動替換並重新啟動。");
            }
            catch (HttpRequestException ex)
            {
                return UpdateApplyResult.Failed($"下載更新失敗：{ex.Message}");
            }
            catch (IOException ex)
            {
                return UpdateApplyResult.Failed($"寫入更新檔失敗：{ex.Message}");
            }
            catch (Exception ex)
            {
                return UpdateApplyResult.Failed($"準備更新時發生未預期錯誤：{ex.Message}");
            }
        }

        private static (string Name, string DownloadUrl, string Sha256) FindStandaloneAsset(JsonElement root)
        {
            if (!root.TryGetProperty("assets", out var assets) || assets.ValueKind != JsonValueKind.Array)
            {
                return default;
            }

            foreach (var asset in assets.EnumerateArray())
            {
                string name = asset.TryGetProperty("name", out var nameElement) ? nameElement.GetString() ?? string.Empty : string.Empty;
                if (!name.EndsWith("-Standalone.exe", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string downloadUrl = asset.TryGetProperty("browser_download_url", out var urlElement)
                    ? urlElement.GetString() ?? string.Empty
                    : string.Empty;
                string digest = asset.TryGetProperty("digest", out var digestElement)
                    ? digestElement.GetString() ?? string.Empty
                    : string.Empty;
                return (name, downloadUrl, NormalizeSha256(digest));
            }

            return default;
        }

        private static async Task<string> CalculateSha256Async(string path)
        {
            await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1024 * 128, useAsync: true);
            byte[] hash = await SHA256.HashDataAsync(stream);
            return Convert.ToHexString(hash);
        }

        private static string NormalizeSha256(string value)
        {
            return value.Replace("sha256:", string.Empty, StringComparison.OrdinalIgnoreCase).Trim();
        }

        private static bool CanWriteToDirectory(string directory)
        {
            try
            {
                string probe = Path.Combine(directory, $".paperswitch-write-check-{Guid.NewGuid():N}.tmp");
                using (File.Create(probe)) { }
                File.Delete(probe);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string BuildUpdaterScript(int processId, string currentExe, string downloadedExe, string latestVersion)
        {
            string escapedCurrentExe = currentExe.Replace("\"", "\"\"");
            string escapedDownloadedExe = downloadedExe.Replace("\"", "\"\"");
            string versionFile = Path.Combine(Path.GetDirectoryName(currentExe)!, "version.txt").Replace("\"", "\"\"");
            string version = latestVersion.StartsWith("v", StringComparison.OrdinalIgnoreCase) ? latestVersion : $"v{latestVersion}";

            return $"@echo off\r\n" +
                   "setlocal\r\n" +
                   $":wait\r\ntasklist /FI \"PID eq {processId}\" /NH | find \"{processId}\" >nul\r\n" +
                   "if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto wait)\r\n" +
                   $"move /Y \"{escapedCurrentExe}\" \"{escapedCurrentExe}.bak\" >nul\r\n" +
                   $"move /Y \"{escapedDownloadedExe}\" \"{escapedCurrentExe}\" >nul\r\n" +
                   $"if not exist \"{escapedCurrentExe}\" goto cleanup\r\n" +
                   $"> \"{versionFile}\" echo {version}\r\n" +
                   $"start \"\" \"{escapedCurrentExe}\"\r\n" +
                   $":cleanup\r\ndel \"%~f0\"\r\n";
        }
    }

    public sealed class UpdateApplyResult
    {
        public bool Success { get; private init; }
        public bool ShouldExitApplication { get; private init; }
        public string Message { get; private init; } = string.Empty;

        public static UpdateApplyResult Started(string message) => new() { Success = true, ShouldExitApplication = true, Message = message };
        public static UpdateApplyResult Failed(string message) => new() { Success = false, Message = message };
    }
}
