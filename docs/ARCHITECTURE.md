# 系統架構與資料流向 (Architecture)

## 1. 系統整體架構 (C# 12 / .NET 8 / WPF)

本專案採用 Windows 原生 WPF + MVVM 架構，提供向量無損 PDF 裝訂與安全隔離的 Office COM 轉檔服務。啟動效能須於目標電腦另行量測，不以文件宣告取代實測。

```
[MainWindow.xaml / LightboxWindow.xaml] (WPF XAML 溫暖手作工坊)
                   │
                   ▼ (雙向資料綁定 / RelayCommand)
[MainViewModel / LightboxViewModel] (CommunityToolkit.Mvvm)
                   │
    ┌──────────────┼──────────────┬──────────────┐
    ▼              ▼              ▼              ▼
[OfficeConverter] [PdfService]  [ThumbnailCache] [ImageConverter]
 (Word/Excel/PPT)  (PdfSharp)    (Windows.Data)   (BitmapDecoder)
   STA 隔離執行   向量無損合成   WinRT 非同步縮圖  無損包裝 PDF
    │              │              │              │
    └──────────────┴───────┬──────┴──────────────┘
                           ▼
 %LOCALAPPDATA%\\PaperSwitch\\converted/ (裝訂產出 PDF 檔案)
```

## 2. 核心模組職責

- **`MainWindow` / `LightboxWindow`**：
  - 手作插畫手帳質感介面（水彩紙米白 `#F7F4ED`、陶土橘 `#D97736`、森林苔綠 `#5B8266`）。
  - 支援檔案拖放 (Drag & Drop)、多選打包拖曳、鍵盤快速換位與大圖燈箱預覽。
- **`MainViewModel`**：
  - 統籌狀態管理、紙張清單 (`ObservableCollection<PaperItem>`)、非同步佇列調度與導出參數。
- **`OfficeConverterService`**：
  - 專屬 STA 執行緒隔離與 `SemaphoreSlim` 併發鎖。
  - 動態 COM Automation (`Word.Application`, `Excel.Application`, `PowerPoint.Application`)。
  - `Marshal.FinalReleaseComObject` 與確定性垃圾回收，防止背景進程殘留。
  - Excel 智慧工作表偵測、自動寬度適應與全空白分頁過濾。
  - IGEF 中介狀態自動偵測與最長 90 秒等待；逾時時保留檔案並阻止其進入 PDF 流程。
- **`PdfService`**：
  - 基於 `PdfSharp` 進行 100% 向量無損頁面抽取、旋轉角度矩陣累加、合併與單頁拆分。
- **`ThumbnailCacheService`**：
  - 使用 Windows 10/11 內建 `Windows.Data.Pdf` (WinRT) 進行多執行緒高解析光柵化縮圖，搭配記憶體快取。
- **`ImageConverterService`**：
  - 使用 WPF `BitmapDecoder` 與 `PdfSharp` 將 PNG/JPG/WebP/BMP 影像轉換為標準向量頁面 PDF。

## 3. 資料與資源生命週期

1. **檔案匯入**：使用者拖入之 Office 或圖片檔案，經由對應 Service 轉換為標準中介 PDF，存放於 `%LOCALAPPDATA%\\PaperSwitch\\temp_converted`；IGEF 檔案不載入，保留供依規定解密後重新匯入。
2. **縮圖產生**：`ThumbnailCacheService` 非同步產生 `BitmapImage` 並呼叫 `Freeze()` 供 UI 執行緒直接渲染。
3. **無損導出**：導出時僅讀取來源 PDF 原始頁面物件並套用旋轉矩陣，不重新壓縮點陣圖，輸出至 `%LOCALAPPDATA%\\PaperSwitch\\converted`。
4. **發行隔離**：建置腳本只清理 `dist\\publish`；此目錄僅存程式發行檔，不得用於使用者文件。
5. **進程與記憶體銷毀**：COM 物件操作完成即時觸發 `FinalReleaseComObject` 與垃圾回收，確保離開工坊時系統零負擔。
