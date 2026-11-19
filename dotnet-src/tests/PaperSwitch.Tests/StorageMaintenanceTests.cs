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
                Assert.Single(Directory.GetFiles(dir));
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        [Fact]
        public void ClearStaleSessionDirectories_DeletesDeadSessionsAndPreservesLiveInstance()
        {
            string root = Path.Combine(Path.GetTempPath(), $"paperswitch-sessions-{Guid.NewGuid():N}");
            string currentSession = Path.Combine(root, "100-current");
            string livePeerSession = Path.Combine(root, "200-livepeer");
            string crashedSession = Path.Combine(root, "300-crashed");
            Directory.CreateDirectory(currentSession);
            Directory.CreateDirectory(livePeerSession);
            Directory.CreateDirectory(crashedSession);
            File.WriteAllText(Path.Combine(currentSession, "cur.pdf"), "cur");
            File.WriteAllText(Path.Combine(livePeerSession, "peer.pdf"), "peer");
            File.WriteAllText(Path.Combine(crashedSession, "stale.pdf"), "stale");

            try
            {
                var result = new StorageCleanupResult();
                StorageMaintenanceService.ClearStaleSessionDirectories(
                    root,
                    currentSession,
                    result,
                    pid => pid == 200);

                Assert.True(Directory.Exists(currentSession));
                Assert.True(Directory.Exists(livePeerSession));
                Assert.True(File.Exists(Path.Combine(livePeerSession, "peer.pdf")));
                Assert.False(Directory.Exists(crashedSession));
                Assert.Equal(1, result.DeletedFileCount);
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }
    }
}
