# 系統架構與核心演算法 (Architecture)

## 1. 系統整體架構 (C# 12 / .NET 8 / WPF)

本專案採用 Windows 原生 WPF + MVVM 架構，提供向量無損 PDF 裝訂與安全隔離的 Office COM 轉檔服務。

```text
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
 %LOCALAPPDATA%\PaperSwitch\converted/ (裝訂產出 PDF 檔案)
```

---

## 2. 核心模組職責

- **`MainWindow` / `LightboxWindow`**：
  - 手作插畫手帳質感介面（水彩紙米白 `#F7F4ED`、陶土橘 `#D97736`、森林苔綠 `#5B8266`）。
  - 支援檔案拖放 (Drag & Drop)、多選打包拖曳、鍵盤快速換位與大圖燈箱預覽。
- **`MainViewModel`**：
  - 統籌狀態管理、紙張清單 (`ObservableCollection<PaperItem>`)、非同步佇列調度、Undo/Redo 歷史堆疊與導出參數。
  - 以 `ObservableCollection<ConversionTaskItem>` 追蹤每個匯入檔案的等待、處理中、完成、失敗與跳過狀態；單檔失敗不會中斷後續任務。
- **`OfficeConverterService`**：
  - 專屬 STA 執行緒隔離與 `SemaphoreSlim` 併發鎖。
  - 動態 COM Automation (`Word.Application`, `Excel.Application`, `PowerPoint.Application`)。
  - `Marshal.FinalReleaseComObject` 與確定性垃圾回收，防止背景進程殘留。
  - Excel 智慧工作表偵測、自動寬度適應 (`FitToPagesWide = 1`) 與全空白分頁過濾。
  - IGEF 中介狀態自動偵測與最長 90 秒等待；逾時時保留檔案並阻止其進入 PDF 流程。
- **`PdfService`**：
  - 基於 `PdfSharp` 進行 100% 向量無損頁面抽取、旋轉角度矩陣累加、合併與單頁拆分。
- **`ThumbnailCacheService`**：
  - 使用 Windows 10/11 內建 `Windows.Data.Pdf` (WinRT) 進行多執行緒高解析光柵化縮圖，搭配記憶體快取。
- **`ImageConverterService`**：
  - 使用 WPF `BitmapDecoder` 與 `PdfSharp` 將 PNG/JPG/WebP/BMP 影像轉換為標準向量頁面 PDF。

---

## 3. 核心演算法與互動引擎

### 3.1 多選打包拖曳抽取演算法 (Multi-page Chunk Relocation)
1. **多選識別**：使用者可透過點擊、`Ctrl + Click`、`Shift + Click` 或框選複選多頁紙張。
2. **整包抽取**：拖曳任意選取卡片時，系統鎖定所有選取項目清單 `selectedItems`。
3. **相對保序置放**：
   - 先將選取的頁面物件整體自 `ObservableCollection<PaperItem>` 抽出，其餘未選取頁面保持原相對順序。
   - 計算目標插入索引，將 `selectedItems` 整包依序插入目標位置。
   - 自動重建選取狀態與摘要統計，不造成畫面閃爍或卡片重建。

### 3.2 鍵盤雙向換位導航引擎 (Keyboard Reordering Engine)
- **防止索引衝突之方向性遍歷**：
  - `Ctrl + ←`（向前移動）：由**最小索引**開始向左依序交換位置。
  - `Ctrl + →`（向後移動）：由**最大索引**開始向右依序交換位置，避免陣列操作時發生索引重複覆寫衝突。
- **極端位置與旋轉快捷鍵**：
  - `Home` / `End`：整包移至第 1 頁或最後一頁。
  - `R` / `Shift + R`：選取頁面順時針 / 逆時針 90° 旋轉。
  - `Delete`：移除選取頁面。

### 3.3 編排復原與重做引擎 (Arrangement Undo / Redo)
- **快照範圍邊界 (Lightweight Snapshot)**：
  - 每次手動編排動作（換位、旋轉、刪除、清空、插入空白頁、重新命名）僅記錄輕量狀態快照：紙張順序、旋轉角度、選取狀態與匯出檔名。
  - **嚴禁複製縮圖點陣圖或來源實體檔案**，確保記憶體極致輕量。
- **匯入狀態隔離**：批次檔案匯入完成後自動清空 Undo/Redo 歷史，避免復原誤將新匯入項目移除，亦不觸發重新 COM 轉檔。

### 3.4 向量無損合成與旋轉數學公式 (Zero-loss Vector PDF Synthesis)
- **縮圖渲染**：WinRT `Windows.Data.Pdf` 非同步產生點陣縮圖後，呼叫 `BitmapImage.Freeze()` 供 WPF UI 跨執行緒安全綁定。
- **向量導出**：PdfSharp 抽取原始 PDF 頁面物件，累加旋轉角度矩陣：
  $$\text{page.Rotate} = (\text{srcPage.Rotate} + \text{item.Rotation}) \pmod{360}$$
  完全不以點陣縮圖二次重壓，確保文字向量可選取性與最高輸出品質。

---

## 4. Office COM 穩定性與安全隔離協議

1. **專屬 STA 執行緒隔離**：
   - 跨執行緒調用一律包裹於專屬 STA 執行緒（`thread.SetApartmentState(ApartmentState.STA)`）。
   - 全域 `SemaphoreSlim` 序列化 COM 調用，防止多執行緒併發衝突。
2. **確定性資源釋放**：
   - `finally` 區塊明確呼叫 `Close(false)`、`Quit()` 與 `Marshal.FinalReleaseComObject`。
   - 調用 `GC.Collect()` 與 `GC.WaitForPendingFinalizers()` 根除背景殘留進程。
3. **Office IGEF 中介狀態保護機制**：
   - MS Office `ExportAsFixedFormat` 在敏感度標籤 (AIP/MIP) 或背景列印 flush 時，初期可能短暫呈現 `IGEF\x02` 中介標頭。
   - `CheckPdfHeaderAndSize()` 辨識 IGEF 並在 `WaitForPdfReadyAsync()` 進行最長 90 秒輪詢。
   - 若逾時或仍為受保護狀態，嚴禁送入 PdfSharp 導出；保留暫存檔並明確提示使用者解密。

---

## 5. 資料與資源生命週期

1. **檔案匯入**：使用者拖入之 Office 或圖片檔案，經由對應 Service 轉換為標準中介 PDF，存放於 `%LOCALAPPDATA%\PaperSwitch\temp_converted`。
2. **縮圖快取**：`ThumbnailCacheService` 非同步產生並快取於記憶體，支援即時滑桿尺寸調整 (90–480px)。
3. **無損導出**：導出產物輸出至 `%LOCALAPPDATA%\PaperSwitch\converted`，支援合併為單一 PDF 或獨立單頁輸出。
4. **發行隔離**：建置腳本只清理 `dist\publish`；此目錄僅存放程式發行檔，不得用於使用者暫存或導出檔案。

