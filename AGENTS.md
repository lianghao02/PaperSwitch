# AGENTS.md — Agent 角色規範與行為準則

本檔定義 `09_PaperSwitch` 專案專屬限制與開發準則；全域憲法由全域環境統一注入。

## 1. 專案核心架構邊界
- **原生 C# 12 / .NET 8 WPF 優先**：專案核心代碼維護於 `dotnet-src/` 內，採用 MVVM 架構 (`CommunityToolkit.Mvvm`) 與 XAML 手帳主題樣式。
- **Office COM 生命週期精準管理**：調用 Word/Excel/PowerPoint COM 元件時，必須受專屬 STA 執行緒與 `SemaphoreSlim` 保護，且必須在 `finally` 區塊呼叫 `Marshal.FinalReleaseComObject`、`GC.Collect()` 與 `GC.WaitForPendingFinalizers()`。
- **無損向量合成**：導出合成 PDF 時必須透過 `PdfSharp` 頁面物件無損抽取與旋轉矩陣，嚴禁以縮圖二次重壓為 PDF（確保向量字型清晰度與極小檔案體積）。
- **Windows 原生渲染**：PDF 縮圖與大圖燈箱預覽優先採用 Windows 10/11 內建 `Windows.Data.Pdf` (WinRT) 進行多執行緒非同步快取渲染。

## 2. 視覺風格規範
- 介面必須遵循「溫暖手作插畫手帳風格 (Warm Cozy Craft Palette)」：
  - 底色：`#F7F4ED`（紙白）、卡片：`#FFFFFF`、邊框：`#E3DDD2`。
  - 主強調色：`#D97736`（陶土暖橘）、次強調色：`#5B8266`（森林苔綠）、文字：`#2D2825`（炭焙棕黑）。
  - 文案採用溫暖親切的生活感用語，避免生硬冷冽的工程字眼。

## 3. 專案結構與文件
進入本專案修改核心邏輯前，請依序參閱：
1. [`README.md`](README.md)：整體專案功能與使用邊界
2. [`docs/MEMORY.md`](docs/MEMORY.md)：歷史經驗與關鍵演算法決策
3. [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md)：系統架構與模組資料流
4. [`CHANGELOG.md`](CHANGELOG.md)：版本變更歷程
5. `legacy-python/`：舊版 Python 程式碼封存與參考
