using System;
using System.IO;
using System.Threading.Tasks;
using PaperSwitch.Services;
using Xunit;

namespace PaperSwitch.Tests
{
    public class OfficeConverterServiceTests
    {
        [Fact]
        public void CheckPdfHeaderAndSize_EncryptedIgefFile_ReturnsIgefState()
        {
            string path = Path.Combine(Path.GetTempPath(), $"paperswitch-{Guid.NewGuid():N}.pdf");

            try
            {
                File.WriteAllBytes(path, new byte[] { (byte)'I', (byte)'G', (byte)'E', (byte)'F', 0x02 });

                var result = OfficeConverterService.Instance.CheckPdfHeaderAndSize(path);

                Assert.Equal("igef", result.State);
                Assert.True(result.Size > 0);
            }
            finally
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
        }

        [Fact]
        public async Task WaitForPdfReadyAsync_EncryptedIgefFile_ReturnsFalseWithoutLoadingIt()
        {
            string path = Path.Combine(Path.GetTempPath(), $"paperswitch-{Guid.NewGuid():N}.pdf");

            try
            {
                File.WriteAllBytes(path, new byte[] { (byte)'I', (byte)'G', (byte)'E', (byte)'F', 0x02 });

                var startedAt = DateTime.UtcNow;
                bool isReady = await OfficeConverterService.Instance.WaitForPdfReadyAsync(path, timeoutSeconds: 10);

                Assert.False(isReady);
                Assert.True((DateTime.UtcNow - startedAt).TotalSeconds < 8);
            }
            finally
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
        }
    }
}
