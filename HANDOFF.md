# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/PaperSwitch
- **Branch**：main
- **Commit SHA**：fc7f643dcc04fad5ffd4389c4ac393d625d239a0（發布前基準；本輪發行提交以 v4.4.0 標籤及 Git 實際狀態為準）
- **Skill Version**：v1.0.0
- **Task Type**：RELEASE
- **Local Path Hint**：09_PaperSwitch

## 目前狀態
v4.4.0 本機發布建置與檢核完成，已取得使用者提交、推送及建立 Release 授權；遠端發布及 CI 狀態以發行頁面與 GitHub Actions 為準。

## 本輪目標
完成左側操作面板、全高畫布與可選命名後綴功能的版本、README、CHANGELOG、Git 與 GitHub Release 發布。

## 基準與已確認事實 (Baseline & Confirmed Facts)
承接同一工作階段已確認的介面與命名修改，未覆寫其他工作。發布前 main 與 origin/main 分歧為 0／0；本機基準 v4.3.1，先前公開最新版為 v4.3.0。

## 已完成 (Completed)
- 左側圓角面板集中來源、工具、轉檔、匯出，支援收合與捲動；右側畫布保留完整高度。
- 命名開關預設關閉、記住偏好；PDF／PNG 與預覽共用命名規則，流水號與同名防覆寫保留。
- 新增命名、偏好容錯、實際 PDF 重複匯出與預覽更新測試，共 64 項通過。
- 版本提升 v4.4.0，更新 README、CHANGELOG 與 docs/RELEASE-v4.4.0.md。
- 產生獨立單檔 EXE 與 SHA256SUMS.txt，完成兩個獨立中文空白路徑啟動驗證。

## 異動檔案 (Changed Files)
MainWindow.xaml、ExportOptions.cs、MainViewModel.cs、PdfService.cs、ImageExportService.cs；新增 ExportFileNaming.cs、ExportPreferencesStore.cs、ExportNamingTests.cs；更新 ImageExportTests.cs、專案版本與兩份 version.txt，以及 README、CHANGELOG、HANDOFF 和發布報告。實際清單以 Git diff 為準。

## 刻意未修改 (Do Not Do / Deliberately Omitted)
未改 PDF 合成或 Office COM 核心、原始使用者文件、更新服務與 CI 設定；大圖檢視沿用燈箱。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無已確認本機阻斷；遠端結果以 Release／Actions 為準。
- **P2 (重要/當次)**：無待修正的已重現問題。
- **P3 (改善建議/暫緩)**：不同實體電腦、磁碟、Windows 使用者、Windows 10 與實體高 DPI 完整驗收未執行；單頁／總覽共用工作區不在本輪範圍。

## 驗證結果 (Validation)
### 已執行測試與結果
Release 建置 0 警告、0 錯誤；qa.ps1 64／64 通過；單檔版本 4.4.0.0；啟動入口驗證通過；兩個獨立路徑均取得主視窗；WPF 渲染、側欄收合／捲動與命名預覽檢查通過；敏感資料與差異格式檢查通過。詳見發布報告。
### 尚未驗證項目
不同實體環境與完整高 DPI 人工流程；不以離屏高解析渲染代替 OS DPI 實測。
### 已知風險 (Known Risks)
短視窗或展開清單時需捲動側欄。Office 轉檔仍受既有 IGEF／單位加密環境限制，需依核准流程取得標準 PDF。

## Git 狀態
- Commit：發布提交以 v4.4.0 標籤為準，避免將提交自己的 SHA 寫入同一提交。
- Push：已獲授權；實際結果以 origin/main 與 tag 為準。
- Working Tree：交付前檢查，以 git status 為準。
- Branch：main。

## 下一步建議動作 (Next Recommended Action)
由 GitHub Release 下載 v4.4.0 單檔；檢查 Actions 及標籤提交，不再重做已通過的介面與命名工作。

## 發布狀態 (Release Status)
本機成品可交付，遠端 Release：https://github.com/lianghao02/PaperSwitch/releases/tag/v4.4.0 ；跨實體環境限制已在發布說明揭露。