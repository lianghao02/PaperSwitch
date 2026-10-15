# v4.4.1 版本發布驗收報告

## 發布資訊
- **專案全名**：lianghao02/PaperSwitch
- **發布版本**：v4.4.1
- **發布 Commit SHA**：以 v4.4.1 標籤所指提交為準，可用 git rev-list -n 1 v4.4.1 核對。
- **發布日期**：2026-10-15（Asia/Taipei）
- **Release**：https://github.com/lianghao02/PaperSwitch/releases/tag/v4.4.1

## Release Gates 檢驗結果
- **Git 範圍檢查**：承接同一工作階段六個已確認的介面／測試／文件變更；發布前 main 與 origin/main 分歧 0／0。本輪額外更新版本、CHANGELOG 及發布報告，提交後狀態以 Git 為準。
- **自動化驗證**：qa.ps1 完整通過，Release 建置 0 警告、0 錯誤；66／66 測試通過，包含按鈕範圍切換、通知、圖片提示及直接匯出呼叫的選取防護。
- **版面驗證**：WPF 離屏渲染確認 14 個工具標籤均單行且寬度足夠；空畫布／無選取／有選取三種按鈕狀態正確。小視窗可捲動取得最後的圖片匯出按鈕，側欄可收合／展開。
- **發行建置**：build.ps1 -SelfContained 產生 win-x64 獨立單檔；檔案版本 4.4.1.0。run.ps1 -ValidateOnly 選中最新成品。
- **本機跨路徑啟動**：EXE 單獨複製到兩個含中文與空白的新資料夾，開發工具不在 PATH，DOTNET_ROOT 指向不存在位置，單檔解壓快取亦使用新路徑。兩次均取得 PaperSwitch 主視窗標題與有效控制代碼，未意外結束；僅關閉本輪測試程序。
- **文件與版本**：兩份 version.txt、專案 Version／FileVersion／AssemblyVersion、README、CHANGELOG、HANDOFF 與發布報告同步。
- **敏感資料**：提交前檢查新增差異與測試檔案，確認不含機密或個人絕對路徑；附件僅 EXE 與 SHA-256 清單，不含 PDB、使用者文件或個人設定。
- **GitHub CI**：沿用既有 Windows 建置與測試工作流程，正式公開前確認發行提交 CI 成功。實際紀錄以標籤所指提交的 Actions 為準。
- **跨環境限制**：上述為本機隔離路徑測試，不等同不同實體電腦、磁碟、Windows 使用者或 Windows 10 的完整驗收。實際 OS 125%／150% 縮放及完整人工流程尚未驗證。

## 發布產出物清單
- **發行檔案**：PaperSwitch-v4.4.1-Standalone.exe
- **校驗清單**：SHA256SUMS.txt
- **SHA-256**：6ef744628f76e3f7fef160d81462368b6dc547bd9a0e12ad4ccffca07e719f75

## 核心變更摘要
- 工具雙欄七列直接展開，左欄加寬，文字單行。
- 匯出區只保留輸出檔名、後綴勾選框與五個常駐按鈕。
- 全拆與選取匯出依選取狀態互斥可用，圖片保留全部／選取範圍。
- 減少按鈕與轉檔卡片陰影；選取紙張增加「✓ 已選」標記。
- PDF／Office 核心與後綴命名規則未更動，既有 IGEF 環境限制沿用。