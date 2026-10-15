using PaperSwitch.Models;
using PaperSwitch.ViewModels;
using Xunit;

namespace PaperSwitch.Tests;

public sealed class ExportButtonPolicyTests
{
    [Fact]
    public void SelectionChangeNotifiesCommandsAndImageScope()
    {
        var vm = new MainViewModel();
        var page = new PaperItem();
        vm.Pages.Add(page);
        int splitChanges = 0, selectedChanges = 0;
        vm.ExportSplitPdfCommand.CanExecuteChanged += (_, _) => splitChanges++;
        vm.ExportSelectedPdfCommand.CanExecuteChanged += (_, _) => selectedChanges++;
        page.IsSelected = true;
        vm.NotifySelectionChanged();
        Assert.False(vm.ExportSplitPdfCommand.CanExecute(null));
        Assert.True(vm.ExportSelectedPdfCommand.CanExecute(null));
        Assert.True(splitChanges > 0 && selectedChanges > 0);
        Assert.Equal("將選取的 1 頁轉成 PNG 圖片", vm.ImageExportScopeText);
        vm.DeselectAll();
        Assert.True(vm.ExportSplitPdfCommand.CanExecute(null));
        Assert.False(vm.ExportSelectedPdfCommand.CanExecute(null));
        Assert.Equal("將畫布全部 1 頁轉成 PNG 圖片", vm.ImageExportScopeText);
    }

    [Fact]
    public async Task InvalidSelectionStateAlsoBlocksDirectExportCalls()
    {
        var vm = new MainViewModel();
        vm.Pages.Add(new PaperItem { IsSelected = true });
        string originalStatus = vm.StatusMessage;
        await vm.ExportSplitPdfAsync();
        Assert.False(vm.IsBusy);
        Assert.Equal(originalStatus, vm.StatusMessage);
        vm.DeselectAll();
        await vm.ExportSelectedPdfAsync();
        Assert.False(vm.IsBusy);
        Assert.Equal(originalStatus, vm.StatusMessage);
    }
}
