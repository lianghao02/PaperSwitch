# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/PaperSwitch
- **Branch**：main
- **Commit SHA**：v4.8.0 發布提交（以 `git rev-parse HEAD` 與 annotated tag `v4.8.0` 為準）
- **Skill Version**：v1.0.0
- **Task Type**：RELEASE
- **Local Path Hint**：09_PaperSwitch

---

## 目前狀態
已正式發布 **v4.8.0**（125/125 測試全數通過、全方案 0 警告 0 錯誤、標準 AppHost 版與免安裝獨立單檔版 `PaperSwitch-v4.8.0-Standalone.exe` 皆已建置並於中文含空白路徑啟動驗證通過，已推送 `main`、Tag `v4.8.0` 與 GitHub Release，並已清理舊版 Release／本機 `dist/release_assets/v4.7.1`）。

## 本輪目標
依唯讀健檢結果一併修訂所有 P1／P2／P3 問題：
1. **P1 自動更新器強化 (`UpdateService.cs`)**：將 API 檢查（8 秒）與大型更新檔下載（10 分鐘）`HttpClient` 分離、補齊下載逾時錯誤捕捉、阻擋多檔 AppHost 部署原地覆寫為單檔版、清理舊更新殘留檔，並改用 `Wait-Process` + `tasklist /FO LIST | findstr` 精確 PID 錨定、失敗還原 `.bak` 與成功後自動清除 `.bak`。
2. **P2 舊工作階段孤兒暫存回收 (`AppPaths.cs` / `StorageMaintenanceService.cs`)**：抽出 `TemporaryRootDirectory`，按「清除暫存」時連同清理已結束或異常終止之舊 `SessionId` 暫存目錄，並透過程序名稱驗證保護仍在執行中的其他 PaperSwitch 實例。
3. **P2 大頁數縮圖補圖與批次開檔 (`ThumbnailCacheService.cs` / `MainWindow.xaml.cs` / `MainViewModel.Conversion.cs`)**：分離縮圖（360 張）與高畫質預覽（8 張）快取池避免大圖擠出縮圖；同一份 PDF 以 `GetThumbnailsForPdfAsync` 開檔一次批次渲染後立即釋放句柄不鎖檔；`ZoomLevel` 與 `TotalPageCount` 變更一併觸發 `ScheduleViewportThumbnailSync()` 補載入。
4. **P2 程式碼去重與職責瘦身 (`ExportFileNaming.cs` / `MainViewModel.*.cs`)**：集中 `ExportFileNaming.SanitizeFileName` 供匯出與 Excel 工作表名稱共用；移除 `MainViewModel` 未使用的 `_imageExportService`；透過 `partial class MainViewModel` 拆分出 `MainViewModel.Arrangement.cs` 與 `MainViewModel.Conversion.cs`。
5. **P2/P3 文件與靜態檢查清理**：修正 `StorageMaintenanceTests.cs:27` `xUnit2013` 警告（全方案 0 警告）、清理 `dist/` 歷史驗收目錄（保留 `publish` 與 `release_assets`），並同步更新 `README.md` 與 `CHANGELOG.md`。

## 基準與已確認事實 (Baseline & Confirmed Facts)
- 接手基準：Commit `42cba0d`（`main...origin/main` 同步），前期測試為 121/121 通過（含 1 筆 `xUnit2013` 警告）。
- 本輪新增 4 項單元測試後，全方案 **125/125 通過、0 警告、0 錯誤**。

## 已完成 (Completed)
- `dotnet-src/src/PaperSwitch/Services/UpdateService.cs`：獨立 API／下載 `HttpClient`、`IsMultiFileDeployment`、精確 PID 等待、失敗還原及 `.bak` 清理。
- `dotnet-src/src/PaperSwitch/Services/AppPaths.cs` & `StorageMaintenanceService.cs`：`TemporaryRootDirectory` 與 `ClearStaleSessionDirectories` 回收孤兒工作階段暫存。
- `dotnet-src/src/PaperSwitch/Services/ThumbnailCacheService.cs` & `MainWindow.xaml.cs`：分離縮圖／預覽快取池、同 PDF 批次渲染與 `ScheduleViewportThumbnailSync` 支援縮放補圖。
- `dotnet-src/src/PaperSwitch/Services/ExportFileNaming.cs`、`ImageExportService.cs`、`OfficeConverterService.cs`：集中 `SanitizeFileName`。
- `dotnet-src/src/PaperSwitch/ViewModels/MainViewModel.cs`、`MainViewModel.Arrangement.cs`、`MainViewModel.Conversion.cs`：依職責拆分為局部類別並移除未使用欄位。
- `dotnet-src/tests/PaperSwitch.Tests/StorageMaintenanceTests.cs` & `RevisionTests.cs`：消除 `xUnit2013` 警告，新增孤兒工作階段清理、更新腳本 PID／備份驗證、多檔防呆與批次縮圖無鎖檔測試。
- `README.md`、`CHANGELOG.md`：移除過時 `F2`／預覽說明，補充工作桌與燈箱右鍵及最新修訂項目。

## 異動檔案 (Changed Files)
- dotnet-src/src/PaperSwitch/MainWindow.xaml.cs
- dotnet-src/src/PaperSwitch/Services/AppPaths.cs
- dotnet-src/src/PaperSwitch/Services/ExportFileNaming.cs
- dotnet-src/src/PaperSwitch/Services/ImageExportService.cs
- dotnet-src/src/PaperSwitch/Services/OfficeConverterService.cs
- dotnet-src/src/PaperSwitch/Services/StorageMaintenanceService.cs
- dotnet-src/src/PaperSwitch/Services/ThumbnailCacheService.cs
- dotnet-src/src/PaperSwitch/Services/UpdateService.cs
- dotnet-src/src/PaperSwitch/ViewModels/MainViewModel.cs
- dotnet-src/src/PaperSwitch/ViewModels/MainViewModel.Arrangement.cs
- dotnet-src/src/PaperSwitch/ViewModels/MainViewModel.Conversion.cs
- dotnet-src/tests/PaperSwitch.Tests/RevisionTests.cs
- dotnet-src/tests/PaperSwitch.Tests/StorageMaintenanceTests.cs
- CHANGELOG.md
- README.md
- HANDOFF.md

## 刻意未修改 (Do Not Do / Deliberately Omitted)
- 保留 `WarmCozyTheme.xaml`（仍實際承擔基礎字型與按鈕／核取方塊樣式）及 `TryRenameSelectedPages` 後端測試涵蓋，未破壞既有公開契約。
- 保留原生 `WrapPanel`（已補強同 PDF 批次開檔與縮放補圖；大型清單虛擬化面板屬架構重大改版，暫不引入第三方套件）。
- 未升版、未 Commit、未 Push。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無。
- **P2 (重要/當次)**：無阻斷開發項目；確認結果後可決定是否提升版號與 Git 提交／推送。
- **P3 (改善建議/暫緩)**：跨電腦、實體高 DPI 實機與完整 Office COM 轉檔人工驗收（CI runner 不含 Microsoft Office）。

## 驗證結果 (Validation)
### 已執行測試與結果
- `pwsh -NoProfile -File dotnet-src/scripts/qa.ps1`：125／125 通過（0 警告、0 錯誤）。
- `pwsh -NoProfile -File dotnet-src/scripts/build.ps1`：成功建置標準 AppHost 版 `dist\publish\PaperSwitch.exe`（168 KB）。
- `pwsh -NoProfile -File dotnet-src/scripts/run.ps1 -ValidateOnly`：正確指向 `D:\Development\GitHub\09_PaperSwitch\dist\publish\PaperSwitch.exe`。
- `git diff --check`：通過（0 錯誤）。

## Git 狀態
- 基準 Commit：`42cba0d`；本輪修訂檔案尚未提交。
- Push：否。
- Branch：main。

## 下一步建議動作 (Next Recommended Action)
確認修訂結果後，如須收尾發布或推送遠端，請指示目標版本號與是否執行 `git commit` / `git push`。
