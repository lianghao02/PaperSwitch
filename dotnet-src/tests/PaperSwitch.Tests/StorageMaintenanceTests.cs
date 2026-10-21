using System;
using System.IO;
using PaperSwitch.Services;
using Xunit;

namespace PaperSwitch.Tests
{
    public class StorageMaintenanceTests
    {
        [Fact]
        public void ClearDirectoryExcept_ShouldDeleteOnlyFilesNotInKeepList()
        {
            string dir = Path.Combine(Path.GetTempPath(), $"paperswitch-clean-{Guid.NewGuid():N}");
            Directory.CreateDirectory(dir);
            string kept = Path.Combine(dir, "畫布使用中.pdf");
            string unused = Path.Combine(dir, "已匯出不再使用.pdf");
            File.WriteAllText(kept, "keep");
            File.WriteAllText(unused, "remove");

            try
            {
                var result = StorageMaintenanceService.ClearDirectoryExcept(dir, new string?[] { kept.ToUpperInvariant(), null, "" });

                Assert.False(File.Exists(unused));
                Assert.Equal(1, result.DeletedFileCount);
                Assert.Empty(result.FailedPaths);
                Assert.Equal(1, Directory.GetFiles(dir).Length);
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }
    }
}
