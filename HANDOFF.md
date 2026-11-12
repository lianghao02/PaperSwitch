# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/PaperSwitch
- **Branch**：main
- **Commit SHA**：本次提交以 `b47725a` 為基準；實際 SHA 請見 `git log -1`。
- **Skill Version**：v1.0.0
- **Task Type**：HANDOFF
- **Local Path Hint**：09_PaperSwitch

---

## 目前狀態
提交前檢查通過；依使用者要求提交並推送既有發布封裝與暫存清理修訂。

## 本輪目標
核對既有 4 個未提交檔案，執行必要 QA，提交並推送至 origin/main。

## 基準與已確認事實 (Baseline & Confirmed Facts)
- 程式版本維持 4.6.0；本輪不建立版本標籤或 GitHub Release。
- origin 為 `https://github.com/lianghao02/PaperSwitch.git`。
- 2026-11-12 重新 fetch 後，本機基準與 origin/main 一致。
- 前輪記錄曾遭 Trend Micro `Unauthorized file encryption` 攔截；本輪未重新驗證防毒判定或確認根因。

## 已完成 (Completed)
- 核對既有修訂：預設使用 Framework-Dependent 多檔 AppHost；自包含單檔封裝須同時指定 `-SelfContained -SingleFile`。
- 核對既有修訂：匯出完成對話框關閉後不再立即清理暫存；保留程式結束及手動清理流程。
- 核對既有修訂：補齊應用程式描述、著作權及語言中繼資料。
- 恢復 build.ps1 的 UTF-8 BOM，維持 Windows PowerShell 5.1 中文相容性。
- 重新執行完整 QA、差異與敏感資料檢查。

## 異動檔案 (Changed Files)
- dotnet-src/scripts/build.ps1
- dotnet-src/src/PaperSwitch/PaperSwitch.csproj
- dotnet-src/src/PaperSwitch/ViewModels/MainViewModel.cs
- HANDOFF.md

## 刻意未修改 (Do Not Do / Deliberately Omitted)
- PDF 向量合成、Office COM 轉檔、圖片匯出及 UI 樣式維持既有實作。
- 不變更防毒設定，不加入掃描例外。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無已知阻斷項目。
- **P2 (重要/當次)**：無。
- **P3 (改善建議/暫緩)**：如仍出現防毒攔截，需依實際事件另行調查；不能以本輪測試推定誤報已排除。

## 驗證結果 (Validation)
### 已執行測試與結果
- 2026-11-12 `pwsh -NoProfile -File dotnet-src/scripts/qa.ps1`：87／87 通過，建置 0 警告、0 錯誤；含 WPF 載入、快照產生及 PDF 渲染測試。
- 初次沙箱內 QA 因 Windows SDK 目錄存取限制失敗；獲准在沙箱外重跑後通過。
- `git diff --check`：通過。
- 待提交程式差異敏感資料掃描：未發現警訊。
### 尚未驗證項目
- 本輪未重新發布成品、人工操作視窗或驗證防毒攔截；未進行跨電腦驗收。
### 已知風險 (Known Risks)
- 預設多檔 AppHost 需要對應 .NET Desktop Runtime 與同目錄組件，不能僅攜帶 EXE。
- Office 正常轉檔仍受本機 Office 與組織文件保護環境限制。

## Git 狀態
- Commit：本次提交，實際 SHA 請見 `git log -1`。
- Push：本輪目標為 origin/main；實際結果以遠端與最終回覆為準。
- Working Tree：本次僅納入上述 4 個檔案，提交後確認。
- Branch：main。

## 下一步建議動作 (Next Recommended Action)
依使用者下一項需求繼續；若要正式發布，另執行發布與環境驗收。

## 發布狀態 (Release Status)
本輪僅提交原始碼，不建立正式發行版本。
