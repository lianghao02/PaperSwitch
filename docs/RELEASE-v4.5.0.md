# v4.5.0 版本發布驗收報告

## 發布資訊
- **專案全名**：lianghao02/PaperSwitch
- **發布版本**：v4.5.0
- **發布 Commit SHA**：以 v4.5.0 標籤所指提交為準，可用 git rev-list -n 1 v4.5.0 核對。
- **發布日期**：2026-10-21（Asia/Taipei）
- **Release**：https://github.com/lianghao02/PaperSwitch/releases/tag/v4.5.0

## Release Gates 檢驗結果
- **自動化驗證**：qa.ps1 完整通過，Release 建置 0 警告、0 錯誤；84 項測試通過（含對話框快照、失敗任務旗標、暫存清理與捲軸判斷測試）。
- **使用者實測**：發布前已由使用者實際使用一輪後授權發布；實測細節未另行記錄。
- **發行建置**：build.ps1 -SelfContained 產生 win-x64 單檔，檔案版本 4.5.0.0。
- **文件與版本**：兩份 version.txt、專案 Version／FileVersion／AssemblyVersion、README、CHANGELOG、HANDOFF 與發布報告同步。
- **附件**：僅 EXE 與 SHA256SUMS.txt，不含 PDB、使用者文件或個人設定。
- **驗證邊界**：未涵蓋 Windows 10、不同實體電腦與實體高 DPI 的完整驗收。

## 發布產出物清單
- **發行檔案**：PaperSwitch-v4.5.0-Standalone.exe
- **校驗清單**：SHA256SUMS.txt
- **SHA-256**：0d868b7693478085de2f3e045074961acef7233fdcb9e9326788df1d69560f1d

## 核心變更摘要
- 2D Q 版文具工坊介面；匯出按鈕名稱與說明更清楚，匯出完成彈出對話框。
- 工具盒分組、選取更醒目、所有提示改為工坊風格對話框。
- Office 轉出的 PDF 被公司加密程序鎖住時，提供「處理步驟」與「我已解密，繼續」。
- 匯出完成與關閉程式時自動清除轉檔暫存。
- 修正點捲軸被當成框選、按鈕文字顏色失效與圖示黑線。
