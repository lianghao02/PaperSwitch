# HANDOFF

## 目前狀態
🟢 **已正式發布，進入 Stable／Maintenance 維護階段**

## 專案版本與發布資訊
- **正式發布版本**：`v4.1.2` (Tag: `v4.1.2`)
- **Default Branch**：`main`
- **Release 狀態**：GitHub Release 正式發布完成，附帶 `PaperSwitch-v4.1.2-Standalone.exe` 與 `SHA256SUMS.txt`
- **自動化測試**：38 / 38 測試項 100% 通過（0 略過、0 警告、0 錯誤）

## 已完成事項
- 穩定化分支 `refactor/paperswitch-stabilization` 已安全合併至 `main`。
- 補強 PDF 輸出邊界防護（空來源防護、不可寫入路徑錯誤處理、異常尺寸安全保留）。
- 補強 A4、頁面尺寸、巨大頁面、異常輸出、獨立輸出及 WPF XAML 啟動等回歸測試（測試數擴充至 38 項）。
- `RUN.bat` 改為純 ASCII 薄啟動器，實際決策集中於 `dotnet-src/scripts/run.ps1`。
- 新增 `docs/MIGRATION_AUDIT.md`，完成正式 .NET 功能與舊版 Python 之功能稽核。
- 完成發布後文件校準：修正 `README.md` 下載清單對齊 GitHub Release 資產，更新 `CHANGELOG.md` 與專案版本。

## 刻意保留與架構約束（請勿隨意更動）
- **保留 `legacy-python/`**：因舊版 PDF 批次轉圖片與 LibreOffice 後援尚未等價遷移，刻意保留備援，請勿隨意刪除。
- **保留既有 COM 與 PDF 核心**：維持 Office COM STA 執行緒隔離、`SemaphoreSlim`、COM Release 確定性釋放與 PdfSharp 向量輸出架構。
- **維持現有 WPF UI/MVVM**：不進行不必要之 UI 大改或 ViewModel 拆分。

## 驗證結果
- **Release Build**：0 警告、0 錯誤。
- **xUnit Tests**：38 / 38 通過。
- **發行驗證**：`dist/PaperSwitch-v4.1.2-Standalone.exe` 檔案完整，SHA-256 校驗無誤，GitHub Releases 與 API 辨識正常。

## Git 狀態
- **Branch**：`main`
- **Working Tree**：Clean

## 下一步
本專案已正式進入穩定維護階段（Stable / Maintenance）。後續若無明確之 Bug 回報或全新業務需求，不需進行大型重構或非必要之程式碼異動。
