# v4.6.0 版本驗收報告

## 發布資訊

- **專案**：lianghao02/PaperSwitch
- **版本**：v4.6.0
- **版本提交**：以 Git 標籤 `v4.6.0` 指向的提交為準。
- **日期**：2026-10-26
- **交付範圍**：原始碼升版、提交與推送，GitHub Release 已公開，EXE 與 SHA256 附件已上傳。

## 更新內容

- 預覽放大後可按住左鍵抓取拖曳，提供手掌／抓取游標。
- PDF 匯出可加左下、下方置中、右下頁碼與 8～24 pt 字級，依本次匯出順序連續編號；旋轉與裁切頁面亦適用。
- 功能按鈕改用簡潔向量圖示搭配文字，保留文具插畫與配色，匯出按鈕以文字為主。

## 驗收結果

- **完整 QA**：`pwsh -NoProfile -File dotnet-src/scripts/qa.ps1`，87／87 通過，建置 0 錯誤。
- **PDF 渲染**：三種位置、8／12／24 pt、直橫式、四種來源旋轉加編排旋轉及 CropBox；拆分前三頁與合併結果逐像素一致，來源位元組不變。
- **原生操作**：前輪已載入六頁測試 PDF，確認預覽放大及左鍵拖曳；960×720 與 1264×860 版面快照已檢視。
- **獨立 EXE**：win-x64 自包含單檔，檔案版本 4.6.0.0。
- **跨路徑啟動**：由 D 槽複製至 C 槽中文含空白路徑，僅複製 EXE；清除 DOTNET_ROOT、DOTNET_ROOT_X64、DOTNET_BUNDLE_EXTRACT_BASE_DIR，PATH 僅保留 Windows System32，取得 PaperSwitch 主視窗後正常關閉。
- **版本對齊**：中央 `check_release_versions.py 09` 回報 OK。
- **Git 與敏感資料**：提交前執行差異檢查及異動範圍的常見權杖／私鑰模式掃描；最終分支與遠端狀態以 Git 為準。

## 本機成品

- `dist/release_assets/v4.6.0/PaperSwitch-v4.6.0-Standalone.exe`
- `dist/release_assets/v4.6.0/SHA256SUMS.txt`
- **SHA256**：`a8a5597498c05c5f68adb6bd73414d0ead4ac5c92f0e0d7f7cb13754ff1eeb86`
- QA 紀錄：`dist/release_assets/v4.6.0/qa.log`；快照與渲染 PDF：`dist/qa-v4.6.0/`。啟動器使用的 `dist/publish` 已更新為相同 SHA256 的自包含 v4.6.0；舊版交付包、過期試用包、舊 QA／預覽、過期說明及重複建置目錄已清除。新版 22 個驗證／交付檔案清理前後雜湊一致。成品與暫存資料不納入 Git。

## 限制

- 頁碼於匯出加入，預覽不模擬新增頁碼；PNG 不套用，原稿已有頁尾時可能重疊。
- 同一台電腦已驗證跨磁碟／中文空白路徑與環境隔離；不同 Windows 使用者、Windows 10 實機、實體高 DPI 及真實公司 Office／加密流程未新增驗收。
- **GitHub Release**：[v4.6.0 下載頁](https://github.com/lianghao02/PaperSwitch/releases/tag/v4.6.0)。遠端 EXE digest 與本機 SHA256 一致，已設為最新正式版本。
