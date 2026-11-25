# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/PaperSwitch
- **Branch**：main
- **Commit SHA**：dc2d1e70c34920244b2c9d0398d4b67db23b1593（基準標籤 v4.8.0；本輪細節修訂尚未提交）
- **Skill Version**：v1.0.0
- **Task Type**：FIX
- **Local Path Hint**：09_PaperSwitch

---

## 目前狀態
已提交並推送至 `main`（於 v4.8.0 基準上完成選檔清單補齊、多頁 TIFF 展開、實測列數鍵盤導航修正、收合工具盒完全釋放工作桌寬度、轉檔任務捲軸留白與 README 快捷鍵勘誤；127/127 測試全數通過、全方案 0 警告 0 錯誤）

## 本輪目標
修訂 v4.8.0 唯讀健檢確認的最後 3 項細節小問題：
1. 補齊「＋ 加入檔案」`OpenFileDialog` 篩選清單（納入 `.xlsm`、`.xlsb`、`.csv`、`.rtf`、`.ppsx`、`.pps`、`.tif`、`.tiff`、`.ico` 等後端實際支援格式），並讓 `ImageConverterService` 支援多頁 TIFF 逐頁展開為 PDF 頁面。
2. 修正 `MainWindow.GetVisualColumnCount()`：改為實測工作桌第一排容器個數（含精準 padding／卡片邊距備援），解決寬螢幕小縮圖下 `↑`／`↓` 偏移問題。
3. 收合左側 `☰ 工具盒` 時將整個左側面板 `SidebarPanel` 隱藏（消除約 100px 空白直柱），並在工作桌綠標題帶左側顯示 `☰ 工具盒` 展開按鈕；同步替轉檔任務列表加上右側捲軸留白，並修正 `README.md` 快捷鍵表（移除未綁定的 `Backspace`）。

## 基準與已確認事實 (Baseline & Confirmed Facts)
- 接手基準：Commit `dc2d1e7`（`v4.8.0`），前期測試 125/125 通過、0 警告。
- 本輪新增 2 項驗證測試後，全方案 **127/127 通過、0 警告、0 錯誤**。

## 已完成 (Completed)
- `dotnet-src/src/PaperSwitch/Services/ImageConverterService.cs`：多頁 TIFF（`.tif`/`.tiff`）逐幀轉成多頁 PDF，其餘圖片維持單幀處理避免動圖爆頁。
- `dotnet-src/src/PaperSwitch/MainWindow.xaml`：收合工具盒時隱藏整個 `SidebarPanel` 並在綠標題列顯示 `SidebarRestoreToggle`；`TaskItemsControl` 加入右側 `4px` 留白避免貼近捲軸。
- `dotnet-src/src/PaperSwitch/MainWindow.xaml.cs`：`GetVisualColumnCount()` 改為實測第一列卡片數＋精準備援計算；補齊 `OpenFileDialog.Filter` 支援格式清單。
- `dotnet-src/tests/PaperSwitch.Tests/ImageConverterServiceTests.cs` & `WpfSmokeTests.cs`：新增多頁 TIFF 3 頁轉換測試，以及收合工具盒釋放寬度＋視覺列數實測對齊快照測試。
- `README.md`：修正刪除快捷鍵說明為 `Delete`。

## 異動檔案 (Changed Files)
- dotnet-src/src/PaperSwitch/MainWindow.xaml
- dotnet-src/src/PaperSwitch/MainWindow.xaml.cs
- dotnet-src/src/PaperSwitch/Services/ImageConverterService.cs
- dotnet-src/tests/PaperSwitch.Tests/ImageConverterServiceTests.cs
- dotnet-src/tests/PaperSwitch.Tests/WpfSmokeTests.cs
- README.md
- HANDOFF.md

## 刻意未修改 (Do Not Do / Deliberately Omitted)
- 維持既有 2D Q 版文具工坊配置與極簡工作流程，不額外堆疊複雜功能或改寫虛擬化面板。
- 未升版（仍維持 v4.8.0，僅提交修訂至 `main`）。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無。
- **P2 (重要/當次)**：無；本專案功能與介面均已收斂至穩定交付水準。
- **P3 (改善建議/暫緩)**：跨電腦與實體 Office COM 環境人工驗收。

## 驗證結果 (Validation)
### 已執行測試與結果
- `pwsh -NoProfile -File dotnet-src/scripts/qa.ps1`：127／127 通過（0 警告、0 錯誤）。
- `pwsh -NoProfile -File dotnet-src/scripts/build.ps1`：成功建置標準 AppHost 版 `dist\publish\PaperSwitch.exe`（168 KB）。
- `pwsh -NoProfile -File dotnet-src/scripts/run.ps1 -ValidateOnly`：正確指向 `D:\Development\GitHub\09_PaperSwitch\dist\publish\PaperSwitch.exe`。
- 視覺快照覆核（`sidebar_expanded_v1.png`、`sidebar_collapsed_v1.png`）：確認 1280×800 收合後左側空白直柱完全消失，單列卡片由 9 張增為 11 張，綠標題帶左側 `☰ 工具盒` 清楚對齊。
- `git diff --check`：通過（0 錯誤）。

## Git 狀態
- 基準 Commit：`dc2d1e7`；本輪 7 個修改檔案已提交並推送至 `origin/main`（最新 Commit 以 `git rev-parse HEAD` 為準）。
- Push：是。
- Branch：main。

## 下一步建議動作 (Next Recommended Action)
目前版本已穩定可日常使用；若後續需要對外釋出新執行檔，再依 Release 流程升版打包即可。
