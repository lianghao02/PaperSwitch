using System.IO;
using System.Text;
using System.Text.Json;

namespace PaperSwitch.Services;

/// <summary>僅保存命名偏好；設定檔損壞或無法寫入不阻斷文件匯出。</summary>
public sealed class ExportPreferencesStore
{
    private readonly string _path;
    public ExportPreferencesStore(string? path = null) => _path = path ?? Path.Combine(AppPaths.DataDirectory, "export-preferences.json");

    public bool Load()
    {
        try { return JsonSerializer.Deserialize<Preferences>(File.ReadAllText(_path, Encoding.UTF8))?.IncludeDescription ?? false; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException) { return false; }
    }

    public bool Save(bool includeDescription)
    {
        string temporaryPath = _path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(_path))!);
            File.WriteAllText(temporaryPath, JsonSerializer.Serialize(new Preferences { IncludeDescription = includeDescription }), new UTF8Encoding(false));
            File.Move(temporaryPath, _path, true);
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return false; }
        finally
        {
            try { if (File.Exists(temporaryPath)) File.Delete(temporaryPath); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
        }
    }

    public sealed class Preferences { public bool IncludeDescription { get; set; } }
}
