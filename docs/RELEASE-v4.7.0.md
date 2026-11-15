# v4.7.0 發布驗收報告

- Repository：lianghao02/PaperSwitch。
- 版本：v4.7.0，日期：2026-11-15；原始碼提交以 annotated tag v4.7.0 指向為準。
- 更新內容：安全暫存清理、完整性匯出、原子成品寫入、統一忙碌狀態、預覽旋轉歷史、頁序範圍／同來源選取、安全停止、輸出位置保存及 PDF metadata 背景讀取。

## 已完成的成品驗證

- 正式升版後完整 QA：99／99 通過，0 建置錯誤。
- 使用者指定 D:/scan：90 份 PDF、798 頁可讀，原始數量及 SHA-256 不變；抽樣輸出通過，範圍見 REVISION_VALIDATION.md。
- 中央 check_release_versions.py 09：OK，0 版本矛盾。
- 本機 AppHost：dist/publish/PaperSwitch.exe，版本 4.7.0.0。
- 下載成品：dist/release_assets/v4.7.0/PaperSwitch-v4.7.0-Standalone.exe，83,333,938 bytes；內含 .NET Runtime，保留既有自動更新附件命名規則。
- SHA256：87ed53416ae48daabbe95164b021affbbd2213fcb0836fcf5d4f0fd894230b02。
- 啟動驗證：只複製單檔 EXE 至 C 槽中文含空白暫存路徑；清除 DOTNET_ROOT 系列及 Bundle 環境設定，PATH 僅保留 Windows System32；取得 PaperSwitch 主視窗，正常關閉，複製前後雜湊一致。
- 差異格式與待提交範圍敏感資料掃描通過；使用者原始 PDF 與測試衍生成品不加入版本庫。

## 發布與清理

- 本次已獲使用者明確授權提交、推送、製作發行檔；清理範圍限本機舊發行檔，GitHub 歷史發行保留。
- GitHub 發布頁：https://github.com/lianghao02/PaperSwitch/releases/tag/v4.7.0。
- 上傳並核對遠端附件後，才清除本機 dist/release_assets/v4.6.0；新版 QA、雜湊及啟動驗證證據保留。
- 最終遠端狀態與清理結果由後續交接及終端驗證補記。

## 限制

- 同一電腦跨磁碟／中文空白路徑與 Runtime 環境隔離啟動通過，不代表不同 Windows 使用者或 Windows 10 實機均已驗證。
- 真人滑鼠與實體高 DPI、Office 正常轉檔及現場取消未新增驗收；Office 仍須本機安裝 Microsoft Office，取消不強制終止 COM。
- 未簽章 EXE 的安全軟體判定由使用環境決定；本輪不更改防毒或組織文件保護設定。
