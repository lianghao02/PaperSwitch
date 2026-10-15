# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/PaperSwitch
- **Branch**：main
- **Commit SHA**：8f9feddf387500c5983dcb7085cbfe575b389d89（發布前基準；本輪發行提交以 v4.4.1 標籤與 Git 實際狀態為準）
- **Skill Version**：v1.0.0
- **Task Type**：RELEASE
- **Local Path Hint**：09_PaperSwitch

## 目前狀態
v4.4.1 本機發布建置與檢核完成，已取得使用者提交、推送及建立 Release 授權；遠端結果以 Git、Release 與 GitHub Actions 為準。

## 本輪目標
依使用者確認，工具區改為雙欄七列、文字不換行；匯出區只保留輸出檔名、後綴勾選框與五個常駐按鈕，並調整卡片與選取狀態。

## 基準與已確認事實 (Baseline & Confirmed Facts)
接續已發布 v4.4.0，開始時 main 與 origin/main 同步且工作目錄乾淨。沿用先前核對的 Acrobat 文件置中、工具靠邊配置，僅修訂本次已確認範圍。

## 已完成 (Completed)
- 左側面板加寬，工具雙欄共 14 個按鈕全部展開；一般按鈕改為無常態陰影的小圓角樣式，保留滑鼠、按下、鍵盤焦點與停用狀態。
- 移除「更多操作」「其他匯出方式」及檔名預覽的介面；保留後綴偏好與既有命名、防覆寫規則。
- 五個匯出按鈕依指定順序直接顯示。全拆僅無選取時可用，選取匯出僅有選取時可用；空畫布皆停用。
- 同時加入匯出方法內的狀態防護，避免直接呼叫繞過停用條件。圖片依選取／全部處理，提示顯示實際範圍。
- PDF 紙張使用較小圓角，已選取頁面增加「✓ 已選」標記；轉檔卡片維持淡米底並減少陰影。

- 更新 v4.4.1 版本、README、CHANGELOG 與發布報告；產生獨立單檔與 SHA-256 清單，完成兩個中文空白路徑啟動檢查。

## 異動檔案 (Changed Files)
- dotnet-src/src/PaperSwitch/MainWindow.xaml
- dotnet-src/src/PaperSwitch/ViewModels/MainViewModel.cs
- dotnet-src/tests/PaperSwitch.Tests/ArrangerTests.cs
- dotnet-src/tests/PaperSwitch.Tests/ExportButtonPolicyTests.cs（新增）
- README.md
- HANDOFF.md
- CHANGELOG.md、version.txt、dotnet-src/src/PaperSwitch/version.txt、dotnet-src/src/PaperSwitch/PaperSwitch.csproj
- docs/RELEASE-v4.4.1.md（新增）

## 刻意未修改 (Do Not Do / Deliberately Omitted)
未更動 PDF／Office 核心、命名服務、偏好格式或 CI。內部命名預覽計算仍保留，使用者介面已移除；不覆寫舊 Release 附件。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無已確認阻斷。
- **P2 (重要/當次)**：無已重現待修正問題。
- **P3 (改善建議/暫緩)**：實體 Windows 高 DPI、Windows 10 與完整人工匯入／拖曳／匯出流程仍未驗收。

## 驗證結果 (Validation)
### 已執行測試與結果
- qa.ps1：Release 0 警告、0 錯誤，66／66 測試通過；包含新按鈕狀態通知、圖片範圍提示及無效選取狀態的直接匯出防護。
- WPF 離屏排版檢查：14 個工具按鈕皆單行且文字寬度可容納；驗證空白、無選取與有選取的實際按鈕狀態，小視窗可捲動取得最後的圖片匯出按鈕；側欄收合／展開與全高畫布維持。
- 圖片預覽以合成測試頁面產生，不使用原始使用者文件。
- v4.4.1 獨立單檔版本 4.4.1.0、啟動入口與兩個獨立中文空白路徑啟動通過；詳見 docs/RELEASE-v4.4.1.md。
### 尚未驗證項目
實際 OS 125%／150% 縮放與不同 Windows 環境。高解析渲染不等同實體 DPI 驗收。
### 已知風險 (Known Risks)
側欄加寬會減少畫布每排可排列的縮圖數，小視窗需捲動左欄。Office IGEF 限制沿用前輪結論。

## Git 狀態
- Commit：發行提交以 v4.4.1 標籤所指提交為準，避免提交自己的 SHA 寫入同一提交。
- Push：已獲使用者授權，結果以 origin/main 與 tag 為準。
- Working Tree：交付前以 git status 核對，僅本輪確認範圍。
- Branch：main。

## 下一步建議動作 (Next Recommended Action)
由 GitHub Release 下載 v4.4.1，核對發行標籤對應的 Actions；延續已通過的介面與命名成果。

## 發布狀態 (Release Status)
本機成品可交付；遠端發布狀態以 https://github.com/lianghao02/PaperSwitch/releases/tag/v4.4.1 為準。不同實體環境與高 DPI 限制已揭露。