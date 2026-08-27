using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
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

                return new UpdateCheckResult
                {
                    Success = true,
                    HasUpdate = hasUpdate,
                    CurrentVersion = currentVersion,
                    LatestVersion = tagName.StartsWith("v", StringComparison.OrdinalIgnoreCase) ? tagName : $"v{tagName}",
                    ReleaseNotes = releaseNotes,
                    HtmlUrl = htmlUrl
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
    }
}