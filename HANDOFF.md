# HANDOFF

## 目前狀態
可交付

## 本輪目標
完成 .NET 正式版穩定化稽核、Python 功能對照、核心回歸測試、啟動器整理與最終驗證。

## 已完成
- 以 `b169051032f81ca633b1657279ca1ee7f6a8d8be` 建立基準並建立 `refactor/paperswitch-stabilization`。
- 新增 `docs/MIGRATION_AUDIT.md`，確認正式 .NET 功能與舊版差異。
- 補強 A4、頁面尺寸、巨大頁面、異常輸出、獨立輸出及 WPF XAML 啟動測試。
- PDF 合併與編排輸出遇到不可寫路徑時改為穩定回傳失敗，不讓例外穿透 UI。
- `RUN.bat` 改為純 ASCII 薄啟動器，實際決策集中於 `dotnet-src/scripts/run.ps1`。

## 刻意未修改
- Office COM STA、`SemaphoreSlim` 與 COM 釋放流程。
- PdfSharp 向量頁面抽取、旋轉及既有排版規則。
- WPF UI、操作流程、預設輸出規則與版本號。
- `legacy-python/`；因 PDF 轉圖片與 LibreOffice 後援尚未等價遷移而保留。

## 尚未完成
- 未以真實 Word、Excel、PowerPoint 檔案執行 Office COM 轉檔；Repository 沒有無敏感測試樣本。
- 未測 200～500 頁大型 PDF 的實際捲動效能。

## 驗證結果
### 已執行
- Baseline：Release Build 0 警告、0 錯誤；既有測試 30/30 通過。
- 最終 QA：clean、Release Build 與測試 38/38 通過，0 略過。
- WPF 主視窗 XAML 在 STA 執行緒初始化成功。
- 圖片轉 PDF、WinRT 縮圖、向量合併、拆分、旋轉、A4 與異常輸出測試通過。
- 啟動器在既有成品、缺少成品自動建置、中文與空白路徑三種情境通過。

### 尚未驗證
- 真實 Office COM 文件轉換。
- Windows 自動化介面未提供原生 App 視窗控制，未執行滑鼠操作巡覽；以 WPF XAML 測試及可回應處理序取代。

### 已知風險
- 公務端點加密仍可能使 Office 暫存 PDF 保持 IGEF；現有行為會保留檔案並提示依規定解密，不嘗試繞過。
- .NET 尚無舊版的 PDF 批次轉圖片與 LibreOffice 後援。

## Git 狀態
- Commit：`2621bbe`（本輪實作；本 HANDOFF 另以後續文件提交保存）
- Push：待本輪推送
- Working Tree：Clean（提交本文件後）
- Branch：`refactor/paperswitch-stabilization`

## 下一步
由使用者在實際公務電腦以非敏感 Office 樣本試用改善分支；確認後再評估合併至 `main`，不自動合併。
