# v4.4.2 版本發布驗收報告

## 發布資訊
- **專案全名**：lianghao02/PaperSwitch
- **發布版本**：v4.4.2
- **發布 Commit SHA**：以 v4.4.2 標籤所指提交為準，可用 git rev-list -n 1 v4.4.2 核對。
- **發布日期**：2026-10-15（Asia/Taipei）
- **Release**：https://github.com/lianghao02/PaperSwitch/releases/tag/v4.4.2

## Release Gates 檢驗結果
- **Git 範圍**：承接本工作階段已確認的九個命名／測試／文件修訂。發布前 main 與 origin/main 分歧 0／0；本輪另外更新版本、CHANGELOG 及發布報告，提交後狀態以 Git 為準。
- **自動化驗證**：qa.ps1 完整通過，Release 建置 0 警告、0 錯誤；80／80 測試通過。
- **命名與檔案實測**：實際多來源 PDF 合併／拆分名稱正確，重複檔名保留既有內容。實際 PNG 匯入、轉成 PDF 再輸出 PNG 使用原始名稱，原始圖片位元組一致。驗證日期時間流水號、同分鐘重複輸出、手動名稱優先及未命名空白頁備援。
- **發行建置**：既有 build.ps1 -SelfContained 產生 win-x64 單檔，檔案版本 4.4.2.0；run.ps1 -ValidateOnly 選中最新成品。
- **本機跨路徑啟動**：EXE 單獨放入兩個中文空白路徑，移除開發工具 PATH、DOTNET_ROOT 指向不存在路徑，單檔解壓快取採新資料夾；兩次均取得主視窗標題與有效控制代碼，未意外結束。只關閉本輪測試程序。
- **文件與版本**：兩份 version.txt、專案 Version／FileVersion／AssemblyVersion、README、CHANGELOG、HANDOFF 與發布報告同步。
- **敏感資料與附件**：提交前檢查新增差異與新增測試，排除機密與個人絕對路徑；發行附件僅 EXE 及 SHA256SUMS.txt，不含 PDB、使用者文件或個人設定。
- **GitHub CI**：沿用既有 Windows 建置與測試流程，發行提交 CI 成功後才公開；實際結果以發行標籤對應的 Actions 為準。
- **驗證邊界**：上述是本機獨立路徑檢查，不代表不同實體電腦、磁碟、Windows 使用者、Windows 10 或實體高 DPI 的完整驗收。Office COM 原生轉換未重新驗收，既有 IGEF 限制不變。

## 發布產出物清單
- **發行檔案**：PaperSwitch-v4.4.2-Standalone.exe
- **校驗清單**：SHA256SUMS.txt
- **SHA-256**：6b63613cca493a783b17c7e1feba0f4f554d875ab5914e6c6cae068557dc8421

## 核心變更摘要
- 留空全合併採畫布第一個有效來源名稱；全拆／圖片各自保留來源名稱與流水號。
- 留空選取匯出採日期時間加流水號，同分鐘同名時選擇下一個可用編號。
- 未命名空白頁使用每批固定時間備援，真正名為空白頁.pdf 的來源仍保留原名。
- 保留原始匯入任務檔名，避免 Office／圖片轉檔暫存 GUID 帶入輸出。
- 手動檔名、使用者重新命名、用途後綴、既有匯出按鈕規則及 PDF／Office 核心維持。