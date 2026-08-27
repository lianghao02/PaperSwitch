using System;
using PaperSwitch.Services;
using PaperSwitch.ViewModels;
using Xunit;

namespace PaperSwitch.Tests
{
    public class UpdateServiceTests
    {
        [Fact]
        public void GetCurrentVersion_ShouldReturnValidVersionString()
        {
            var version = UpdateService.Instance.GetCurrentVersion();
            Assert.False(string.IsNullOrWhiteSpace(version));
            Assert.StartsWith("v", version, StringComparison.OrdinalIgnoreCase);
        }

        [Theory]
        [InlineData("v4.0.0", "v4.1.0", true)]
        [InlineData("v4.0.0", "v4.0.1", true)]
        [InlineData("4.0.0", "v4.0.1", true)]
        [InlineData("v4.0.0", "v5.0.0", true)]
        [InlineData("v4.0.0", "v4.0.0", false)]
        [InlineData("v4.1.0", "v4.0.0", false)]
        [InlineData("v4.0.1", "v4.0.0", false)]
        public void IsNewerVersion_ShouldCompareSemanticVersionsAccurately(string current, string latest, bool expected)
        {
            var result = UpdateService.IsNewerVersion(current, latest);
            Assert.Equal(expected, result);
        }

        [Fact]
        public void MainViewModel_AppVersionBadge_ShouldMatchUpdateServiceVersion()
        {
            var vm = new MainViewModel();
            Assert.Equal(UpdateService.Instance.GetCurrentVersion(), vm.AppVersionBadge);
            Assert.False(vm.IsCheckingUpdate);
            Assert.Equal("🔄 檢查更新", vm.CheckUpdateButtonText);
        }
    }
}