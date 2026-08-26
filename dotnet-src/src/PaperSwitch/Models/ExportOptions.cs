using CommunityToolkit.Mvvm.ComponentModel;

namespace PaperSwitch.Models
{
    /// <summary>
    /// PDF 導出與排版設定選項
    /// </summary>
    public partial class ExportOptions : ObservableObject
    {
        [ObservableProperty]
        private bool _mergeIntoSinglePdf = true;

        [ObservableProperty]
        private bool _addPageNumbers = false;

        [ObservableProperty]
        private string _outputDirectory = string.Empty;

        [ObservableProperty]
        private string _customFileName = string.Empty;

        [ObservableProperty]
        private bool _excelFitToPage = true;
    }
}
