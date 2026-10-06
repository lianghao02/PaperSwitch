# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/PaperSwitch
- **Branch**：main
- **Commit SHA**：a422e176f508d2d48139cca60934f0152d7e4cd2（本輪提交前基準；最新提交以 Git 記錄為準）
- **Skill Version**：v1.0.0
- **Task Type**：HANDOFF
- **Local Path Hint**：09_PaperSwitch

---

## 目前狀態
本機入口修復與標準 PDF 功能維持通過。Office IGEF 依辦公室文件加密環境限制追蹤：GoPatrol 服務及保護驅動正在執行；可讀 Office PDF 仍待管理端允許的流程驗收。未發現須修改產品演算法的缺陷。下方舊交接屬歷史，不代表目前 Git 或測試狀態。

## 本輪目標
釐清 Office IGEF 輸出的環境限制，校正 README 說明並留下允許流程的驗收條件；維持既有功能與文件保護。

## 基準與已確認事實 (Baseline & Confirmed Facts)
修復前 publish 只有 PDB 與 version.txt，沒有 EXE；現有 Release DLL 比真正來源新，但 bin／obj 生成的 .cs 被納入新舊判斷。Windows PowerShell 5.1 讀取未帶 BOM 的中文腳本另發生語法錯誤。修復前 README、HANDOFF、AGENTS 已有未提交修改，均承接保留；HEAD／分支不變。

## 已完成 (Completed)
2026-10-06 GitHub 同步交接：使用者已授權提交與推送前輪成果；本輪只提交已核對範圍。最新 Commit SHA、遠端同步與 CI 結果統一見控制中心 `docs/github-sync/RESULTS.md`，不將提交本身的 SHA 寫入同一份提交。 Office 輸出仍受 IGEF 文件加密環境限制；保留既有拒絕防護，不宣稱正常 Office 轉 PDF 已通過。

2026-10-06 IGEF 診斷：確認正在執行的 GoPatrol 服務／驅動、Office 外掛註冊及有效 DLL 簽章。依辦公室文件加密限制追蹤，精確規則未確認。補充 PowerShell COM 對照逾時，未取得輸出，未列為成功；只回收本輪隱藏 Word，Office 殘留 0。README 更正 IGEF 非微軟一般暫存狀態及 5 秒拒絕／90 秒等待上限。詳見中央 docs/new-build-acceptance/IGEF-DIAGNOSIS.md。

2026-10-06 代表性驗收：使用正式 Release DLL 轉換合成 RTF／CSV／PPTX，三份輸出均為 IGEF；正常 Office→PDF 未通過，不能標示完整驗收。既有拒絕機制、獨立標準 PDF 合併／旋轉／分頁通過，來源 SHA-256 一致，Office 殘留程序為 0；未變更產品 C# 或系統保護。詳見中央 docs/new-build-acceptance/RESULTS.md，依允許的 PDF 產出方式處理環境限制後再驗收。

2026-10-05 本機入口修復：run.ps1 排除 bin／obj、逐一檢查有效成品、以 Release DLL 判斷新舊，支援未指定 RID 的 Release；run/build 腳本使用 UTF-8 BOM 相容 PowerShell 5.1。build.ps1 不再先刪除整個 publish 資料夾；已實際建置發布 EXE。新增隔離入口回歸測試，並校正 README 清理功能的實際行為。完整證據見中央 docs/paperswitch-launch-repair/RESULTS.md。

2026-10-05 README 文件更新：補齊專案概念、開發原因、典型流程、已知 Bug／限制及回報方式，並依實際入口校正必要操作說明。本次沒有修改產品程式、環境或個人資料，未 Commit／Push；前輪成果與既有待辦繼承。文件檢核與逐案索引由控制中心 docs/readme-refresh/RESULTS.md 彙整，不代表本次重新驗收全部功能。

清除清冊中 v4.0.0 至 v4.2.1 舊成品、舊 SHA、Debug 與測試傾印；保留現行 v4.3.0 EXE/SHA、release_notes 及 Release 建置成品，未改 C# 或啟動器。

## 異動檔案 (Changed Files)
本次：README.md、HANDOFF.md，中央診斷報告／結果報告／改善總表／交接；診斷產物在 Git 忽略的 artifacts。前輪 run/build/test-launcher 腳本及本機成品繼承，本次未變更產品程式或成品。

## 刻意未修改 (Do Not Do / Deliberately Omitted)
C#／XAML／專案檔、RUN.bat、全域環境與正式 v4.3.0 獨立發行 EXE／SHA／release_notes 未改；32 個保護檔案及既有 AGENTS 修改雜湊一致，46 個既有使用者資料檔案在 GUI 驗證前後一致。舊交接內容保留。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無已確認的本機入口阻斷。
- **P2 (重要/待外部流程)**：Office IGEF 依辦公室加密環境限制追蹤；GoPatrol 元件狀態已確認，精確管理規則未確認。以單位核准流程取得標準 PDF，之後再驗收三種 Office；目前不將其列為須改演算法的 Bug，也不標成正常轉檔通過。
- **P3 (改善建議/暫緩)**：未因整理擴大重構；正式發布另依 release-gate 驗證。

## 驗證結果 (Validation)
### 已執行測試與結果
本輪 IGEF 診斷：5 個保護檔案 SHA-256 一致、兩個 Repository HEAD／分支／索引不變；本輪文件差異及語言檢查通過。Office 殘留 0，DesktopFramesPlus 2.9.4.0 仍在執行。補充 PowerShell COM 對照逾時未通過，僅回收本輪 Word；另一次唯讀 Git 檢查停滯，回收後完成狀態／索引與文件範圍核對，不推定與保護元件有因果關係。未修改產品 C#，未重跑未受影響的核心測試。

Windows PowerShell 5.1 與 PowerShell 7 隔離入口各 8 項通過（共 16 項，含中文／空白路徑）；實際建置發布成功；既有核心測試 50 通過、0 失敗、0 略過；兩種主機 ValidateOnly 通過；發布 EXE 主視窗成功啟動，驗證後僅關閉本輪建立的程序。資料與來源保護雜湊通過。Git 差異檢查結果見中央修復報告。
### 尚未驗證項目
未重新驗收全部原生功能或其他電腦/Windows 10 發布環境。
### 已知風險 (Known Risks)
入口新舊判斷依檔案時間，不等同內容雜湊或完整發布驗收；保留 Win10、Office 與其他電腦的驗證邊界。正式獨立版未重新打包。

## Git 狀態
- Commit：上述 SHA 為提交前基準；最新 SHA 見 `git log -1` 與中央同步報告。
- Push：實際推送及遠端核對結果見中央 `docs/github-sync/RESULTS.md`。
- Working Tree：最終狀態見中央同步報告；不含被忽略的環境、成品與使用者資料。
- Branch：main。

## 下一步建議動作 (Next Recommended Action)
可用 RUN.bat 啟動本機開發成品；標準 PDF 功能沿用。Office 先採單位允許的 PDF 產出／讀取流程，再驗收可讀 3、阻斷 0、來源不變與 Office 殘留 0。本輪停止擴大診斷與修改；日後提交前另取得授權。

## 發布狀態 (Release Status)
本輪僅建置本機 framework-dependent 開發成品，未建立或上傳正式發布版；現行 v4.3.0 獨立成品保留。

---

## 承接的前輪交接（原文保留，屬歷史）


> 2026-10-05 環境修復交接：本輪僅修正 AGENTS.md 的共用 Skill 正式來源為 configs/skills，程式碼與既有環境不變。Working Tree 為 Modified，未 Commit／Push；下列發布與功能紀錄為承接的前輪成果。

# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/PaperSwitch
- **Branch**：main
- **Commit SHA**：d458feb
- **Skill Version**：v1.0.0
- **Task Type**：RELEASE

> **目前實際 Git 狀態**：Working Tree 為 Clean 狀態；全套 50 項測試與 Release QA 均 100% 通過。

## 目前狀態
🟢 **v4.3.0 正式發布完成（Stable / Maintenance 維護階段，所有功能、測試、發行檔案與文件皆已就緒）**

## 專案版本與發布資訊
- **當前專案版本**：v4.3.0
- **Default Branch**：main
- **發行成品**：
  - `dist/release_assets/PaperSwitch-v4.3.0-Standalone.exe` (75.42 MB)
  - `dist/release_assets/SHA256SUMS.txt`
  - SHA-256: `bf46cbbcbd3bca0da2335c0e7349dc0d91167ab37f22f7bdc3e0bc4d3e1dfec0`
- **自動化測試**：50 / 50 測試項 100% 通過（0 略過、0 警告、0 錯誤）

## 本輪完成事項
- **實作 Windows 原生高品質 PDF 轉圖片 (PNG) 服務**：
  - 新增 ImageExportService.cs，使用 Windows 10/11 內建 WinRT Windows.Data.Pdf.PdfDocument 進行原生向量渲染（預設寬度 2400px，提供高清晰品質）。
  - 保留旋轉狀態：當紙張有旋轉（90° / 180° / 270°）時，自動透過 WPF TransformedBitmap 套用旋轉。
  - 使用 WPF 原生 PngBitmapEncoder 儲存為標準 PNG 檔案，零新增外部第三方 NuGet 相依。
- **安全命名與同名衝突防護 (Safe Data Processing)**：
  - 命名格式支援：
    - 自訂檔名／前綴：{Prefix}_{畫布序號:D3}.png
    - 無自訂檔名：{來源檔名幹}_頁{來源頁碼}_{畫布序號:D3}.png
  - 同名衝突防覆寫：目標檔案存在時自動添加 _1.png、_2.png 等後綴，絕不覆寫或刪除既有檔案與來源 PDF。
- **操作模型完全相容現有規範**：
  - 連接 MainViewModel.ExportImagesAsync()，共用 GetEffectiveExportPages() 決策邏輯：
    - Case 1（無選取）：匯出目前畫布全部頁面。
    - Case 2（單選 1 頁）：只輸出該頁 PNG。
    - Case 3（多選 N 頁）：只輸出被選取的 N 頁。
    - Case 4（已移除頁面）：以當前畫布集合為準，已移除之頁面不輸出。
    - Case 5（重排順序）：依畫布實際排版順序編號並輸出。
- **UI 與互動體驗 (Windows Tool UX) 及非最大化版面防擠壓優化**：
  - 於 MainWindow.xaml 底部操作列新增 [ 🖼️ 匯出圖片 ] 次按鈕（BtnOutlineCraft），與另存、拆分維持一致手帳風格視覺層級。
  - 匯出過程提供進度百分比與檔名狀態回饋，完成後自動開啟輸出資料夾。
  - **解決非最大化視窗檔名文字框擠壓問題**：
    - 左側精簡統計資訊，移除重複的「已選取 0 頁」與常態提示文字，僅保留「已裝載 X 頁 (已選取 Y 頁)」膠囊標籤。
    - 將輔助維護性質之「🧹 清除暫存」按鈕移至頂部工具列【群組 3：清理與維護】，底部僅保留核心輸出設定。
    - 精簡按鈕文案（「另存選取」、「拆分選取」）與 Padding，並設定檔名文字框 Width=120，使底部列在 960px 視窗化寬度下仍保有充裕間距，徹底防止檔名輸入框被推擠遮蔽。
- **單元與整合測試驗證**：
  - 新增 ImageExportTests.cs，包含命名生成、非法字元清理、防覆寫序號生成與選取行為測試。
  - 執行 dotnet-src\scripts\qa.ps1，50 項測試全數通過（含 Release Build 0 警告、0 錯誤）。

## 異動檔案清單
- dotnet-src/src/PaperSwitch/Services/ImageExportService.cs (NEW)
- dotnet-src/src/PaperSwitch/ViewModels/MainViewModel.cs (MODIFIED)
- dotnet-src/src/PaperSwitch/MainWindow.xaml (MODIFIED)
- dotnet-src/tests/PaperSwitch.Tests/ImageExportTests.cs (NEW)
- dotnet-src/scripts/run.ps1 (MODIFIED)
- dotnet-src/tests/PaperSwitch.Tests/ArrangerTests.cs (MODIFIED)

## 驗證結果
- **Release Build**：0 警告、0 錯誤。
- **xUnit Tests**：50 / 50 通過。
- **WPF Smoke Test**：STA Thread XAML 載入通過。
- **QA Script**：pwsh -File dotnet-src\scripts\qa.ps1 綠燈通過。

## 下一步
v4.3.0 已正式完成封裝、校驗與全套文件更新。目前專案回歸 Stable / Maintenance 維護階段。
