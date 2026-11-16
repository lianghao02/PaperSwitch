# 09_PaperSwitch Agent 開發規範

本專案遵循目前有效之全域開發憲法；本檔僅定義專案專屬規則與例外。

---

## 1. 技術棧與核心架構邊界
- **原生 C# 12 / .NET 8 WPF 優先**：專案核心代碼維護於 `dotnet-src/` 內，採用 MVVM 架構 (`CommunityToolkit.Mvvm`)，維持原生秒開、低資源占用與獨立發布能力。
- **.NET SDK 環境變數**：本機 SDK 位於 `%LOCALAPPDATA%\Microsoft\dotnet`，CLI 建置與測試需確保 `DOTNET_ROOT` 正確配置。
- **Office COM 生命週期精準管理**：
  - 調用 Word/Excel/PowerPoint COM 元件時，必須受專屬 STA 執行緒與 `SemaphoreSlim` 保護。
  - 必須在 `finally` 區塊明確呼叫 `Marshal.FinalReleaseComObject`、`GC.Collect()` 與 `GC.WaitForPendingFinalizers()`，杜絕背景殘留孤兒 Office 進程。
- **無損向量合成**：
  - 導出合成 PDF 時必須透過 `PdfSharp` 進行頁面物件無損抽取與旋轉矩陣計算，**嚴禁以點陣圖縮圖二次重壓為 PDF**（確保向量文字清晰度與最小體積）。
- **Windows 原生渲染**：
  - PDF 縮圖與大圖預覽優先採用 Windows 10/11 內建 WinRT `Windows.Data.Pdf` 進行非同步多執行緒渲染。

---

## 2. 業務領域與轉檔佇列邊界 (UX 簡潔原則)
- **多檔轉換佇列透明度**：
  - 工具核心定位為簡潔直覺之輪播轉檔工具，**嚴禁過度設計將簡單工具改造成複雜 Dashboard**。
  - 轉檔佇列需清晰回饋：目前處理中檔案、等待佇列清單、已完成數、失敗明確原因、手動重新嘗試與跳過按鈕。
- **視覺風格規範**：
  - 介面採「2D Q 版文具工坊」：遊戲風格與商業工具感並重，使用奶油紙色、焦糖描邊、薄荷綠選取與杏桃橘主要操作；使用木框、皮革縫線、金屬包邊與原創手繪文具圖示，靜態畫面即可辨識遊戲風格；美術素材內嵌，文件內容與功能名稱保持清楚，保留既有操作配置。
  - 文案維持親切生活感，避免冷冽生硬字眼。

---

## 3. Agent 協作與工作區規範
- **協作原則**：遵循全域憲法單一主責原則，Agent 角色可互換；接手與移交時依規定維護 `HANDOFF.md`。
- **工作區隔離**：需要平行開發或隔離任務時，以獨立工作區或分支隔離，不讓兩個 Agent 同時操作同一工作目錄。
- **邊界保護**：各 Agent 嚴禁更動對方的專屬配置（`.agents/` vs `.gemini/`），專案共通規則均維護於本檔。

---

## 4. 核心驗證方式
- 修改程式碼、Core、COM 轉換或 ViewModel 後，必須執行標準 QA 腳本進行完整建置與測試：
  ```powershell
  pwsh -File dotnet-src\scripts\qa.ps1
  ```
- 或直接執行單元測試：
  ```powershell
  $env:DOTNET_ROOT = "$env:LOCALAPPDATA\Microsoft\dotnet"
  $env:PATH = "$env:LOCALAPPDATA\Microsoft\dotnet;$env:PATH"
  dotnet test dotnet-src\PaperSwitch.sln --no-restore --nologo
  ```
