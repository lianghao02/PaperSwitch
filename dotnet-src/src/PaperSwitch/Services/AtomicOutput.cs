using System.IO;

namespace PaperSwitch.Services;

/// <summary>先完成本次暫存檔，再以不覆寫的方式交付成品。</summary>
internal static class AtomicOutput
{
    public static void Write(string destination, Action<string> write)
    {
        var fullPath = Path.GetFullPath(destination);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var temporary = fullPath + "." + Guid.NewGuid().ToString("N") + ".partial";
        try
        {
            write(temporary);
            File.Move(temporary, fullPath, false);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }
}
