# v4.7.1 發布驗收報告

- Repository：lianghao02/PaperSwitch；日期：2026-11-15。
- 原始碼以 annotated tag v4.7.1 指向為準。
- 更新：成果匯出五按鈕單列；修正拖起放回原位誤移最後；拖曳上下邊緣自動捲動。

## 驗證
- 升版後 qa.ps1 完整 109／109 通過，0 建置錯誤。
- 960、1280、1920 三種寬度同列排版；排序、復原及捲動停止條件通過。
- 已用 50 頁畫布及 D:/scan 的 23 頁 PDF 實測原位放下、上下邊緣拖曳與復原，沒有匯出或覆寫來源。
- 中央離線版本檢查 OK；本機 AppHost 更新至 4.7.1。
- 單檔 EXE：PaperSwitch-v4.7.1-Standalone.exe，83,336,498 bytes。
- SHA256：9eb4e63b6f225f709f8625f93d9736544d52260f2f71f8bf4a5d2ee6763a75ea。
- 只複製 EXE 到 C 槽中文空白路徑；移除 DOTNET_ROOT 系列及 Bundle 設定，PATH 僅 Windows System32；主視窗建立並正常關閉，複製雜湊一致。
- QA 與封裝首次同時還原套件發生 NuGet 檔案衝突；改依序執行後完整通過。

## 發布與清理
- 使用者已授權提交、推送、更新發行檔及移除本機舊成品。
- 發布頁：https://github.com/lianghao02/PaperSwitch/releases/tag/v4.7.1。
- 遠端附件核對後才清理本機 dist/release_assets/v4.7.0；GitHub 歷史 Release 保留。
- 本機測試中的預覽 EXE 若仍有開啟畫布，暫不移除，避免影響使用者工作。

## 限制
- 未新增不同 Windows 使用者、Windows 10 實機、實體高 DPI 及 Office 現場流程驗收。
- 滑鼠工具實測快速拖曳進出邊緣；停留連續捲動另以 WPF 計時器回呼驗證。
