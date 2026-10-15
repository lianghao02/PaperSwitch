# v4.4.0 版本發布驗收報告

## 發布資訊
- **專案全名**：lianghao02/PaperSwitch
- **發布版本**：v4.4.0
- **發布 Commit SHA**：以 v4.4.0 標籤所指提交為準，可用 git rev-list -n 1 v4.4.0 核對。
- **發布日期**：2026-10-15（Asia/Taipei）
- **發行頁面**：https://github.com/lianghao02/PaperSwitch/releases/tag/v4.4.0

## Release Gates 檢驗結果
- **Git 範圍檢查**：main，origin 為 lianghao02/PaperSwitch；開始時遠端分歧 0／0，本輪僅包含已確認的介面、命名、測試及版本文件。提交後工作目錄與遠端結果以 Git 為準。
- **自動化測試**：qa.ps1 通過，Release 0 警告、0 錯誤，64／64 項測試通過，含 WPF 載入及實際 PDF 防覆寫。
- **本機套件驗收**：使用既有 build.ps1 -SelfContained 產生 win-x64 單檔，檔案版本 4.4.0.0；run.ps1 -ValidateOnly 選中最新成品。
- **跨路徑啟動**：EXE 單獨複製至兩個含中文與空白的新資料夾；清除開發工具 PATH、DOTNET_ROOT 指向不存在位置，並使用各自新的單檔解壓快取。兩次均取得 PaperSwitch 主視窗標題及有效視窗控制代碼，未意外結束；測試後僅關閉本輪程序。
- **跨環境限制**：上述為本機隔離路徑驗證，沒有冒稱不同實體電腦、磁碟、Windows 使用者或 Windows 10 的完整 Fresh Environment 驗收；實體 125%／150% 縮放及完整人工流程亦未驗證。
- **文件與版本**：根目錄及應用程式 version.txt、專案 Version／FileVersion／AssemblyVersion、README、CHANGELOG 與 HANDOFF 同步。
- **敏感資料與路徑**：新增差異與新增原始碼檢查通過；發行附件僅包含 EXE 與 SHA-256 清單，不包含 PDB、原始文件、個人設定或開發環境。
- **GitHub CI**：使用既有 windows-latest 工作流程；實際結果以發行標籤所指提交的 Actions 記錄為準。

## 發布產出物清單
- **發行檔案**：PaperSwitch-v4.4.0-Standalone.exe
- **校驗清單**：SHA256SUMS.txt
- **SHA-256**：44114193d1211311b46e9b236708292854b85247a37cd44b07a931a28fb3fc88

## 核心變更摘要
- 左側可收合圓角操作面板與右側全高 PDF 排版畫布。
- 附加用途說明開關、偏好保存與即時檔名預覽。
- 保留流水號與同名防覆寫；預覽隨選取與重新命名更新。
- 含前輪 v4.3.1 的動態頁數、卡片懸停操作與子視窗主題調整。
- Office 仍需安裝 Microsoft Office；IGEF 文件加密限制延續，正常 Office 轉 PDF 未因本輪而宣稱已完成環境驗收。