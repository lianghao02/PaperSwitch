# MEMORY.md — 經驗持久化與 Bug 紀錄

## 📌 持久化經驗與重要架構決策

### 1. C# WPF 多選打包拖曳抽取演算法 (Multi-page Chunk Relocation)
- **演算法機制**：
  1. 使用者可透過點擊、`Ctrl + Click`、`Shift + Click` 或 `Ctrl + A` 複選多頁紙張。
  2. 拖曳選取的任意卡片時，系統識別所有選取項目 `selectedItems`。
  3. 放置時：先將選取的頁面物件整體打包自 `ObservableCollection<PaperItem>` 抽出，其餘剩餘頁面保持原相對位置，再於目標插入位置整批插入，並自動重建選取狀態與摘要統計。

### 2. 鍵盤換位導航引擎 (Keyboard Reordering Engine)
- **實作細節**：
  - `Ctrl + ←` / `Ctrl + →`：支援單選與多選頁面群組移動。向前移動（←）時需由左至右（最小索引開始）依序交換；向後移動（→）時需由右至左（最大索引開始）依序交換，避免在陣列操作時發生索引重複覆寫衝突。
  - `Home` / `End`：整包移至第 1 頁或最後一頁。
  - `R` / `Shift + R`：順時針 / 逆時針 90° 旋轉。

### 3. 編排復原與重做 (Arrangement Undo / Redo)
- **狀態快照範圍**：每筆手動編排動作保存紙張物件順序、旋轉角度、選取狀態與匯出用檔名；不複製縮圖或來源檔案。
- **行為邊界**：移動、拖曳重排、旋轉、刪除、清空與重新命名可透過 `Ctrl+Z`／`Ctrl+Y` 復原或重做。匯入完成後清空歷史，避免復原誤移除新匯入紙張，也不重跑 Office COM 轉檔。

### 4. 向量無損合成與旋轉 (Zero-loss Vector PDF Synthesis)
- **架構設計**：
  - 縮圖渲染：利用 Windows 10/11 內建 `Windows.Data.Pdf` (WinRT) 進行高畫質非同步光柵化並呼叫 `BitmapImage.Freeze()` 供 WPF UI 跨執行緒安全綁定。
  - 導出合成：完全不使用縮圖重壓，而是採用 `PdfSharp` 抽取原始 PDF Page 物件，並累加旋轉矩陣 `page.Rotate = (srcPage.Rotate + item.Rotation) % 360`，保持 100% 原始向量清晰度、文字可選取性與極小檔案體積。

### 5. 溫暖手作插畫風格規範 (Warm Cozy Craft Palette)
- **視覺規範**：
  - 底色：水彩紙米白 `#F7F4ED`
  - 面板卡片：純白 `#FFFFFF` + 米黃柔邊 `#E3DDD2`
  - 主強調色：陶土暖橘 `#D97736`（懸停 `#C46424`）
  - 次強調色：自然苔綠 `#5B8266`
  - 文字色調：深炭焙棕黑 `#2D2825`、鉛筆灰褐 `#736B63`
  - 語彙：採用生活感手作詞彙（紙張排版工坊、裝訂成冊、待整理清單）。

### 6. Office COM 穩定性防禦
- **執行緒隔離與資源釋放**：
  - 跨執行緒調用必須包裹於專屬 STA 執行緒中（`thread.SetApartmentState(ApartmentState.STA)`）。
  - 全域 `SemaphoreSlim` 序列化調用，並在 `finally` 確實 `Close(false)`、`Quit()` 與 `Marshal.FinalReleaseComObject`，最後呼叫 `GC.Collect()` 與 `GC.WaitForPendingFinalizers()`。
  - Excel 寬度適應：`ws.PageSetup.FitToPagesWide = 1`。

### 7. Office IGEF 中介狀態與非同步安全保留機制 (Office Intermediate IGEF State)
- **問題本質**：Windows 下 MS Office 執行 `ExportAsFixedFormat` 時，輸出檔案初期可能短暫呈現 `IGEF\x02` 中介標記（常見於 MIP/AIP 敏感度標籤或背景列印 flush），之後才轉為標準 `%PDF-`。
- **處方規範**：
  1. `CheckPdfHeaderAndSize()` 辨識 `IGEF` 並在 `WaitForPdfReadyAsync()` 期間最長輪詢 90 秒，直到寫入完成。
  2. 逾時或仍為 IGEF 時不得送入 `PdfService`、縮圖或拆頁流程；需保留暫存檔並提示使用者依規定解密後重新匯入。
  3. 使用者匯出資料與暫存檔固定於 `%LOCALAPPDATA%\\PaperSwitch`，不可置於會被建置腳本清理的 `dist`。

---

## 📦 外部依賴追蹤
| 依賴套件 / 框架 | 版本要求 | 職責用途 |
|---|---|---|
| `.NET 8.0 LTS` | `net8.0-windows10.0.19041.0` | Windows 原生桌面執行階段與 WinRT PDF 支援 |
| `CommunityToolkit.Mvvm` | >=8.3.2 | MVVM 雙向綁定、ObservableObject、RelayCommand |
| `PdfSharp` | >=6.1.1 | PDF 頁面無損抽取、矩陣旋轉與向量合併 |
| `Windows.Data.Pdf` | 系統內建 (WinRT) | 超高清 PDF 頁面光柵化縮圖即時渲染 |

---

## 📅 版本演進與里程碑
- **2026-08-25 (v4.0.0)**：全面重構為 C# 12 / .NET 8 LTS / WPF 原生桌面應用程式，具備 STA COM 精準防禦、WinRT 縮圖與向量無損合成；啟動時間待目標電腦實測。
- **2026-08-24 (v3.0.0)**：加入 GitHub Release 檢查與受 SHA-256、備份、語法驗證保護的一鍵更新。
- **2026-08-24 (v2.1.0)**：全面重構為「溫暖手作插畫工坊 (Warm Cozy Craft)」風格。
- **2026-08-23 (v2.0.0)**：發布 10/10 完美里程碑（多選打包拖曳、方向鍵極速換位）。
- **2026-08-04 (v1.0.0)**：專案初始建立。
