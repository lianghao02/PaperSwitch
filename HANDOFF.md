# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/PaperSwitch
- **Branch**：main
- **Commit SHA**：4386f08a728cfa4e15b7d43845075dd389b9603b（發布前基準；本輪發行提交以 v4.4.2 標籤與 Git 實際狀態為準）
- **Skill Version**：v1.0.0
- **Task Type**：RELEASE
- **Local Path Hint**：09_PaperSwitch

## 目前狀態
v4.4.2 本機發布建置與檢核完成，已取得使用者提交、推送與 Release 授權；遠端結果以 Git、發行頁面及 Actions 為準。

## 本輪目標
全合併使用畫布第一個有效來源名稱；全拆／圖片各自使用來源名稱加流水號；選取匯出使用日期時間加流水號。手動名稱優先，沒有來源名稱時採日期時間備援。

## 基準與已確認事實 (Baseline & Confirmed Facts)
開始時 main 與 origin/main 同步且工作目錄乾淨。先前空白合併固定命名為 PaperSwitch.pdf；轉檔頁面的 SourceFileName 來自暫存 PDF，可能帶入 GUID，已改為保留匯入任務的原始檔名。

## 已完成 (Completed)
- 全合併從畫布順序挑第一個有效來源；前方未命名空白頁略過，全部無來源時使用 yyyyMMdd_HHmm_001.pdf。
- 選取匯出使用日期時間加流水號；同分鐘同名已存在時使用下一個可用編號，避免疊加 _001_1。
- 全拆／圖片依各頁來源名稱與本次畫布／選取順序編號；空白頁使用每批固定時間備援。
- 插入空白頁增加來源標記，真正名為空白頁.pdf 的匯入文件不會被誤判。頁面重新命名仍維持有效。
- Office／圖片匯入後，顯示與匯出使用原始名稱，不使用轉檔暫存檔名；實際來源 PDF 路徑保留供渲染與合成。
- 手動輸入名稱、後綴設定、匯出按鈕選取防護與既有防覆寫規則保留；未重新加入預覽介面。

- 更新 v4.4.2 版本、README、CHANGELOG 與發布報告，產生單檔與 SHA-256 清單，完成兩個獨立中文空白路徑啟動檢查。

## 異動檔案 (Changed Files)
- dotnet-src/src/PaperSwitch/Services/ExportFileNaming.cs
- dotnet-src/src/PaperSwitch/Services/PdfService.cs
- dotnet-src/src/PaperSwitch/Services/ImageExportService.cs
- dotnet-src/src/PaperSwitch/ViewModels/MainViewModel.cs
- dotnet-src/src/PaperSwitch/Models/PaperItem.cs
- dotnet-src/src/PaperSwitch/MainWindow.xaml（僅檔名輸入提示）
- dotnet-src/tests/PaperSwitch.Tests/AutomaticExportNamingTests.cs（新增）
- README.md
- HANDOFF.md
- CHANGELOG.md、version.txt、dotnet-src/src/PaperSwitch/version.txt、dotnet-src/src/PaperSwitch/PaperSwitch.csproj
- docs/RELEASE-v4.4.2.md（新增）

## 刻意未修改 (Do Not Do / Deliberately Omitted)
未更動 PDF 合成、Office COM 核心、原始使用者文件、版面配置或 CI；舊 Release 附件保留。手動重新命名仍受尊重，不為還原原名而覆蓋使用者命名。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無已確認阻斷。
- **P2 (重要/當次)**：無已重現待修正問題。
- **P3 (改善建議/暫緩)**：實際 Office 原生轉換仍受既有 IGEF 環境限制；完整人工操作與不同 Windows 環境未重新驗收。

## 驗證結果 (Validation)
### 已執行測試與結果
- qa.ps1：Release 0 警告、0 錯誤，80／80 測試通過；最後調整後再次 dotnet test 通過。
- 實際多來源 PDF 合併與拆分：順序與來源名稱正確，同名來源保持不變。
- 實際 PNG 匯入、轉成 PDF 再匯出 PNG：使用原始圖片名稱、沒有暫存 GUID，原始圖片位元組一致。
- 驗證選取時間命名、同分鐘接續編號、手動名稱優先、空白頁備援與每批固定時間。
- v4.4.2 單檔版本 4.4.2.0、專案啟動入口與兩個獨立路徑啟動檢查通過；詳見 docs/RELEASE-v4.4.2.md。
### 尚未驗證項目
完整人工匯入／拖曳／匯出流程、Office COM 原生來源、不同 Windows 及實體高 DPI。
### 已知風險 (Known Risks)
來源名稱因本次修正改回原始匯入名稱；Excel 多工作表轉出的頁面會共用工作簿來源名稱，以序號區分。既有 Office IGEF 限制不變。

## Git 狀態
- Commit：發行提交以 v4.4.2 標籤所指提交為準，避免將提交自己的 SHA 寫入同一提交。
- Push：使用者已授權，結果以 origin/main 與 tag 為準。
- Working Tree：交付前以 git status 核對，僅本輪確認範圍。
- Branch：main。

## 下一步建議動作 (Next Recommended Action)
由 GitHub Release 下載 v4.4.2，核對發行提交對應的 Actions，延續已通過的命名與介面成果。

## 發布狀態 (Release Status)
本機成品可交付；遠端狀態以 https://github.com/lianghao02/PaperSwitch/releases/tag/v4.4.2 為準。不同實體環境與 Office IGEF 限制已揭露。