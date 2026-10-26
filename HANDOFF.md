# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/PaperSwitch
- **Branch**：main
- **Commit SHA**：v4.6.0（以版本標籤指向的提交為準；升版前基準 bb4a99b）
- **Skill Version**：v1.0.0
- **Task Type**：RELEASE
- **Local Path Hint**：09_PaperSwitch

---

## 目前狀態
可交付（已升版 v4.6.0、完整 QA 與獨立 EXE 驗證通過；提交 67f7efb 與版本標籤已推送，GitHub Release 已公開）

## 本輪目標
依使用者授權完成三項功能修訂後，升版 v4.6.0、提交並推送；本機成品與驗證詳見 docs/RELEASE-v4.6.0.md。

## 基準與已確認事實 (Baseline & Confirmed Facts)
- 開始時 main 工作目錄為 Clean，HEAD 為 bb4a99b，版本 v4.5.0。
- 原預覽滾輪縮放且以捲軸移動，無左鍵拖曳；ExportOptions 有 AddPageNumbers 欄位但未接入匯出。
- 主畫面功能按鈕使用 19～20 DIP 點陣插畫，全選／反選與復原／旋轉有共用圖案。

## 已完成 (Completed)
- 預覽文件區可按住左鍵拖曳，僅在有捲動範圍時啟用；保留捲軸、滾輪縮放、翻頁與旋轉。手掌游標與內建抓取游標切換；放開、遺失擷取或關閉視窗結束拖曳，關閉時釋放游標。
- 全部、選取與拆分 PDF 匯出接入頁碼設定；預設關閉，啟用時預設下方置中／12 pt，可選左下、下方置中、右下及 8／10／12／14／16／18／24 pt。
- 依本次匯出順序從 1 編號，拆分亦連續編號；PNG 不套用，頁碼偏好不跨啟動保存。
- PdfSharp 疊加向量文字，保留來源內容，補償旋轉與 CropBox 裁切範圍，頁碼距邊界 18 pt。
- 功能按鈕採 ActionIcon 向量輪廓並繼承按鈕文字色；區分全選／反選、復原／旋轉，匯出按鈕以文字為主。標誌與空白狀態保留原插畫。
- 新增 PageNumberTests，擴充 WPF 頁碼設定與最小尺寸快照，修正舊快照工具的 VisualBrush 置中裁切。
- 更新 README 與中央 IMPROVEMENTS；獨立自包含試用檔產於 dist/revision-20261025/PaperSwitch.exe，保留既有發布檔。

## 異動檔案 (Changed Files)
- README.md、HANDOFF.md、CHANGELOG.md、docs/RELEASE-v4.6.0.md、兩份 version.txt 與 PaperSwitch.csproj。
- dotnet-src/src/PaperSwitch/LightboxWindow.xaml、LightboxWindow.xaml.cs、MainWindow.xaml。
- dotnet-src/src/PaperSwitch/Models/ExportOptions.cs、Services/PdfService.cs、ViewModels/MainViewModel.cs、Styles/GameWorkshop.xaml。
- 新增 Controls/ActionIcon.cs、Controls/PreviewCursors.cs。
- 新增 dotnet-src/tests/PaperSwitch.Tests/PageNumberTests.cs；更新 WpfSmokeTests.cs。
- 中央 00_Dev-Control-Center/IMPROVEMENTS.md（另倉庫；原待修訂紀錄改為已完成）。

## 刻意未修改 (Do Not Do / Deliberately Omitted)
- Office COM、公司加密處理、PNG 匯出內容、既有發布檔未覆寫；本輪另升版 v4.6.0。
- 未變更圖集原始素材，未引入圖示套件或外部字型。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無。
- **P2 (重要/當次)**：無。
- **P3 (改善建議/暫緩)**：實體高 DPI／其他電腦與真實辦公文件人工驗收；跨專案展示站版本同步暫緩：13_Project-Hub 存在其他工作階段的未提交介面與圖片修改，未改動其檔案。

## 驗證結果 (Validation)
### 已執行測試與結果
- pwsh -NoProfile -File dotnet-src/scripts/qa.ps1：87／87 通過，最終建置 0 錯誤。重新編譯測試專案時既有 StorageMaintenanceTests.cs:27 有 xUnit2013 警告，未修改無關測試。
- 新增 3 組頁碼測試：三種位置、8／12／24 pt、兩種紙張尺寸、四種來源旋轉再加編排旋轉、裁切範圍，以 WinRT 真正渲染並驗證底部位置；拆分前三頁與合併結果逐像素一致，來源檔案位元組不變。
- 最後修改快照渲染方式後，WpfSmokeTests 3／3 通過；已檢視 1264×860 與 960×720 快照，窄視窗匯出按鈕自動換行，文字完整。
- Windows 原生介面實測：載入既有 workshop-sample.pdf 六頁，展開頁碼設定、雙擊預覽、滾輪放大、按住左鍵拖曳，垂直捲動位置由 37.05 移至 97.05；Esc 關閉預覽正常。
- 自包含 win-x64 單檔 publish 成功（約 79.5 MiB），已用獨立 EXE 啟動取得主視窗及新版頁碼控制項；前輪已完成啟動實測；舊試用包現已移除。
- git diff --check 通過；異動檔案常見權杖／私鑰模式掃描無警訊。
- 新版 QA 紀錄：dist/release_assets/v4.6.0/qa.log；PDF 與快照：dist/qa-v4.6.0/（Git 忽略）。舊 qa-20261025 已清除。
### 尚未驗證項目
- 實體 125%／150% DPI、其他電腦與真實公司加密／Office 轉檔；核心流程未更動。
### 已知風險 (Known Risks)
- 頁碼疊在既有頁面底部，原稿已有頁尾時可能重疊，README 已說明。
- 預覽顯示來源頁面，不模擬匯出時才新增的頁碼。

## Git 狀態
- Commit：以 v4.6.0 標籤為準
- Push：以 origin/main 與版本標籤的遠端核對結果為準
- Working Tree：以 git status --short --branch 為準
- Branch：main

## 下一步建議動作 (Next Recommended Action)
使用 dist/release_assets/v4.6.0/PaperSwitch-v4.6.0-Standalone.exe；或由 https://github.com/lianghao02/PaperSwitch/releases/tag/v4.6.0 下載。dist 現在僅保留 publish、qa-v4.6.0、release_assets，publish 已更新至相同 SHA256 的 v4.6.0，啟動器 ValidateOnly 通過。

## 發布狀態 (Release Status)
v4.6.0 本機成品、完整測試、版本對齊及跨磁碟環境隔離啟動已通過；不同使用者與實機限制詳見版本驗收報告。GitHub Release 已公開，EXE 與校驗檔上傳完成，遠端 digest 已核對。
