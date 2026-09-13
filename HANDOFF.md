# HANDOFF

## 目前狀態
🟢 **可交付（功能與測試皆通過，等待審查與提交）**

## 專案版本與發布資訊
- **專案版本**：`v4.2.0`
- **Default Branch**：`main`
- **Release 狀態**：維持既有 `v4.2.0`（本輪不建立新 Tag／Release）
- **自動化測試**：44 / 44 測試項 100% 通過（0 略過、0 警告、0 錯誤）

## 已完成事項
- **統一選取頁面匯出與拆分範圍**：
  - 實作 `GetEffectiveExportPages()`：有選取頁面時只處理選取的頁面；無選取頁面時預設處理畫布全部頁面。
  - 「另存選取頁面」：有選取時僅將選取的頁面合併為 1 個 PDF；無選取時將畫布全部頁面合併為 1 個 PDF。
  - 「拆分選取頁面」：文案由「拆分獨立存檔」更名為「📄 拆分選取頁面」，有選取時僅將選取的頁面各自獨立拆分存檔；無選取時將畫布全部頁面各自獨立拆分存檔。
  - 「匯出全部 PDF」：永遠處理全部畫布頁面，行為維持不變。
  - 按鈕啟用邏輯調整：只要畫布有頁面（`HasPages`），三大按鈕皆可點擊使用；不再因無選取而 Disabled。
- **補強單元與回歸測試**：
  - 新增 `MainViewModel_ExportCommands_CanExecute_ReflectsHasPages`。
  - 新增 `MainViewModel_GetEffectiveExportPages_NoSelection_ReturnsAllPagesInCanvasOrder`。
  - 新增 `MainViewModel_GetEffectiveExportPages_WithSelection_ReturnsOnlySelectedInCanvasOrder`。
  - 新增 `MainViewModel_GetEffectiveExportPages_ReorderedAndRotated_PreservesCanvasState`。
  - 新增 `ExportIndividualPdfs_SelectedSubset_SplitsOnlySelectedPages`。
  - 單元測試數由 41 項擴充至 44 項，100% PASS。
- **同步更新相關文件**：
  - `README.md`、`CHANGELOG.md`、`IMPLEMENTATION_PLAN.md`。

## 刻意保留與架構約束（請勿隨意更動）
- **保留三大按鈕佈局**：不新增第四顆按鈕，維持共用檔名／前綴文字框。
- **維持 100% 向量無損合成**：以 PdfSharp 保留原始旋轉、向量字型與文字層。
- **保留既有 COM 與執行緒架構**：維持 Office COM STA 執行緒隔離與 `SemaphoreSlim`。
- **不主動建立 Tag 或 GitHub Release**：維持目前 `v4.2.0` 版本號。

## 驗證結果
- **Release Build**：0 警告、0 錯誤。
- **xUnit Tests**：44 / 44 通過。
- **WPF Smoke Test**：STA Thread XAML 載入通過。

## Git 狀態
- **Branch**：`main`
- **Working Tree**：Modified（待提交）

## 下一步
1. 檢視 `git diff` 確認無多餘改動。
2. 進行 Git commit：`fix: 統一選取頁面匯出與拆分範圍`。
3. 推送至遠端 `origin/main`。
4. 回報工作結果。
