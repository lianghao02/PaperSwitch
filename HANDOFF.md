# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/PaperSwitch
- **Branch**：main
- **Commit SHA**：0aabeb87fea540b5f7e056d77bd202b7c3c1bc2a（基準標籤 v4.7.1；本輪修訂尚未提交）
- **Skill Version**：v1.0.0
- **Task Type**：IMPROVE
- **Local Path Hint**：09_PaperSwitch

---

## 目前狀態
可交付（已完成第一階段＋第二階段全面修訂、共用輸入對話框自適應高度修復、移除 F2／工具盒重新命名按鈕，以及紙張工作桌單選／多選右鍵匯出與大圖預覽精簡單頁匯出；121/121 測試全數通過、標準 AppHost 版發行成品已建置，未提交 Git Commit／未推送）

## 本輪目標
依健檢（AUDIT）結果與使用者回饋，完整執行核心互動穩定、視覺對齊修正、大頁數效能／記憶體控制與工作桌／燈箱右鍵匯出擴充：
1. 修復共用輸入視窗 `RenameDialog.xaml` 固定高度破版（改為 `SizeToContent="Height"`、輸入框套用 `WorkshopTextBox`、移除壓扁按鈕高度），讓「頁序範圍」等對話框完整呈現。
2. 依使用者拍板移除 `F2` 重新命名快捷鍵及工具盒「重新命名」按鈕，「頁面」區改由「＋ 空白頁」整列滿寬呈現。
3. **紙張工作桌右鍵選單**：右鍵卡片自動維持既有／單選脈絡，提供開啟大圖預覽、單選／多選頁面匯出（單頁 PDF/PNG、多頁合併 PDF、拆分成獨立 PDF、PNG）、順／逆旋轉、同來源全選與刪除；右鍵空白處提供全選、取消選取、插入空白頁與已選頁快速匯出。
4. **大圖預覽區（燈箱）右鍵精簡**：依使用者指示只保留「匯出目前這頁 (PDF)」與「匯出目前這頁 (PNG)」，並以 `LightboxHostCallbacks` + `MainViewModel.ExportCustomPagesAsync` 解耦宿主邏輯。
5. 解決卡片跨縫隙拖曳閃爍與縫隙放下無效問題、補齊合併錯誤原因回傳與立即取消、`GameDialog` 自動對齊目前頂層可見視窗（含燈箱開啟狀態）。
6. 追加匯入新檔案納入 `Ctrl+Z` 復原歷史、頁序範圍支援全形符號容錯、視覺與對比修正，以及大頁數（> 120 頁）可視區縮圖按需載入與遠端記憶體釋放。

## 基準與已確認事實 (Baseline & Confirmed Facts)
- 接手前基準：Commit `0aabeb8`（v4.7.1），工作區原有未提交之 `.agents/AGENTS.md`、`.gemini/AGENTS.md` 刪除與 `AGENTS.md` 修改（本輪維持未動，不予覆寫）。
- 前期 `qa.ps1` 為 109/109 通過；本輪新增 12 個回歸與單元驗證案例後為 **121/121 通過**。

## 已完成 (Completed)
- **核心與服務層**：
  - `dotnet-src/src/PaperSwitch/Services/PageRangeSelection.cs`：相容全形減號 `－`、`—`、波浪號 `~`/`～` 與全形空白 `\u3000`，保留禁止空分段契約。
  - `dotnet-src/src/PaperSwitch/Services/PdfService.cs` & `ExportCoordinator.cs`：新增 `out string? errorMessage` 重載傳出真實錯誤（含具體遺失／無效檔名），迴圈外複用 `XFont`／`XStringFormat` 並支援即時取消。
  - `dotnet-src/src/PaperSwitch/ViewModels/MainViewModel.cs`：抽離 `ExportCustomPagesAsync(ExportKind, IEnumerable<PaperItem>)` 作為工作桌與燈箱共用匯出管線，新增 `GetContextTargetPages`、`ExportContextSelectedPdfCommand`、`ExportContextSplitSelectedPdfCommand`、`ExportContextSelectedPngCommand`、`RotateSelectedClockwiseCommand`、`RotateSelectedCounterClockwiseCommand` 與動態單／多選標籤；追加匯入納入 `Ctrl+Z` 復原；加入 `UpdateVisibleThumbnailWindowAsync` 與共用 `_thumbnailRenderGate` 控管超過 120 頁之縮圖記憶體。
  - `dotnet-src/src/PaperSwitch/ViewModels/LightboxViewModel.cs`：新增 `LightboxHostCallbacks(RotatePage, ExportPagesAsync)`（保留既有 3 參數建構子相容性），精簡僅保留 `ExportCurrentPdfCommand` 與 `ExportCurrentPngCommand` 單頁匯出及 `IsExporting` 鎖定。
- **介面與互動層**：
  - `dotnet-src/src/PaperSwitch/GameDialog.xaml.cs` & `RenameDialog.xaml`：對話框自動依附目前頂層可見作用視窗（避免在燈箱內匯出時彈窗鎖在背後）；移除硬寫按鈕高度，並將 `RenameDialog` 改為 `SizeToContent="Height"` 自適應高度。
  - `dotnet-src/src/PaperSwitch/Styles/GameWorkshop.xaml`：新增 2D Q 版工坊右鍵樣式 `WorkshopContextMenu`、`WorkshopMenuItem`、`WorkshopMenuSeparator`；`SidebarToolButton` 微調至 `MinHeight="37"`。
  - `dotnet-src/src/PaperSwitch/LightboxWindow.xaml` & `LightboxWindow.xaml.cs`：保持精簡頂部列與中央大圖預覽區 `Background="Transparent"` 全區感應右鍵，僅保留匯出目前這頁 (PDF / PNG)。
  - `dotnet-src/src/PaperSwitch/MainWindow.xaml` & `MainWindow.xaml.cs`：為紙張卡片 `CardBorder` 與畫布空白區 `MainScrollViewer` 掛載工坊風格右鍵選單（含單／多選匯出、大圖預覽、旋轉、同來源與刪除），`Card_PreviewMouseRightButtonDown` 確保未選取卡片改為單選、已選卡片保留多選狀態；移除左側「重新命名」按鈕與 `Key.F2`，雙擊與右鍵共用 `OpenLightboxForPage` 注入 `LightboxHostCallbacks`；保留拖曳縫隙定位優化與可視區縮圖防抖。
- **回歸測試**：
  - `dotnet-src/tests/PaperSwitch.Tests/RevisionTests.cs` & `WpfSmokeTests.cs`：新增工作桌單／多選右鍵匯出狀態與文辭測試、工作桌卡片／空白畫布 STA 右鍵與右鍵保留多選行為測試、燈箱僅保留當頁匯出與 `ExportCustomPagesAsync` 單頁／多頁實測。

## 異動檔案 (Changed Files)
- dotnet-src/src/PaperSwitch/GameDialog.xaml.cs
- dotnet-src/src/PaperSwitch/LightboxWindow.xaml
- dotnet-src/src/PaperSwitch/LightboxWindow.xaml.cs
- dotnet-src/src/PaperSwitch/RenameDialog.xaml
- dotnet-src/src/PaperSwitch/MainWindow.xaml
- dotnet-src/src/PaperSwitch/MainWindow.xaml.cs
- dotnet-src/src/PaperSwitch/Services/ExportCoordinator.cs
- dotnet-src/src/PaperSwitch/Services/PageRangeSelection.cs
- dotnet-src/src/PaperSwitch/Services/PdfService.cs
- dotnet-src/src/PaperSwitch/Styles/GameWorkshop.xaml
- dotnet-src/src/PaperSwitch/ViewModels/LightboxViewModel.cs
- dotnet-src/src/PaperSwitch/ViewModels/MainViewModel.cs
- dotnet-src/tests/PaperSwitch.Tests/RevisionTests.cs
- dotnet-src/tests/PaperSwitch.Tests/WpfSmokeTests.cs
- HANDOFF.md

## 刻意未修改 (Do Not Do / Deliberately Omitted)
- 保留原工作區既有的 `.agents/AGENTS.md`、`.gemini/AGENTS.md` 與 `AGENTS.md` 未提交狀態未動。
- 自動更新腳本 PID 精確比對、清理 `StorageMaintenanceTests.cs:27` 既有 `xUnit2013` 警告與清理 `dist/` 歷史驗收目錄（屬方向四，留待下輪整理）。
- 未升版、未 Commit、未 Push。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無。
- **P2 (重要/當次)**：若使用者確認結果滿意，可接著決定是否執行方向四收尾（自動更新器細節與目錄／Git 提交整理）。
- **P3 (改善建議/暫緩)**：跨電腦、實體高 DPI 實機與完整 Office 轉檔人工測試。

## 驗證結果 (Validation)
### 已執行測試與結果
- `pwsh -NoProfile -File dotnet-src/scripts/qa.ps1`：121／121 通過（0 警告、0 錯誤；含工作桌右鍵單／多選匯出、保留多選與燈箱單頁右鍵匯出測試）。
- `pwsh -NoProfile -File dotnet-src/scripts/build.ps1`：成功建置標準 AppHost 版 `dist\publish\PaperSwitch.exe`（168 KB）。
- `pwsh -NoProfile -File dotnet-src/scripts/run.ps1 -ValidateOnly`：正確指向 `D:\Development\GitHub\09_PaperSwitch\dist\publish\PaperSwitch.exe`。
- 視覺快照覆核（`lightbox_context_menu_v2.png`、`workspace_context_menu_v1.png`）：已確認燈箱回到簡潔標題列與預覽佈局，主視窗工作桌卡片正常呈現。

## Git 狀態
- Commit：0aabeb8；本輪 15 個修改檔案＋既有的 AGENTS 異動未提交。
- Push：否。
- Branch：main。

## 下一步建議動作 (Next Recommended Action)
請使用者雙擊根目錄 `RUN.bat` 實際體驗修訂後的互動手感與畫面；確認無誤後可告知是否繼續處理方向四收尾與 Git 提交。
