# 📑 PaperSwitch 紙張排版工坊 v4.6.0

[![Release](https://img.shields.io/github/v/release/lianghao02/PaperSwitch?color=orange&label=Release)](https://github.com/lianghao02/PaperSwitch/releases/latest)
[![Platform](https://img.shields.io/badge/platform-Windows%20.NET%208%20LTS-blue.svg)](https://dotnet.microsoft.com/)
[![Theme](https://img.shields.io/badge/style-2D_Workshop-34745F.svg)](README.md)

> **2D 文具工坊風格的 Windows 原生文件轉 PDF、批次合併與視覺化紙張排版工具。**<br>
> 全面重構自 Python 服務架構，升級為 C# 12 / .NET 8.0 LTS / WPF 原生應用程式，帶來秒開速度、確定性 Office COM 生命週期防禦與極致流暢的紙張排版體驗。

---

## v4.6.0 更新重點

- **預覽抓取拖曳**：放大後按住滑鼠左鍵移動畫面，手掌與抓取游標提供操作回饋。
- **PDF 匯出頁碼**：全部、選取與拆分匯出支援左下、下方置中、右下及 8～24 pt 字級，依本次匯出順序從 1 編號。
- **清楚的向量圖示**：功能按鈕改為簡潔輪廓搭配文字，保留文具插畫與既有配色；窄視窗的匯出按鈕自動換行。

完整更新內容見 [CHANGELOG](CHANGELOG.md)，驗證範圍見 [v4.6.0 發布報告](docs/RELEASE-v4.6.0.md)。
## 專案概念與開發原因

PaperSwitch 將 PDF、Office 文件與圖片集中為可預覽、排序及匯出的桌面工作流程。開發動機是混合文件通常需要在多種工具間反覆轉換，容易弄錯頁面順序、旋轉方向或選取範圍。

採 .NET／WPF 與背景工作佇列處理，PDF 頁面處理與 Office 轉換分別使用適合的元件。保留 PDF 的向量內容不代表所有圖片或 Office 來源都能成為向量資料。

**典型流程**：加入文件 → 預覽與調整頁面 → 選擇全部／部分頁面 → 匯出 → 檢查內容與頁數。

## 📥 快速下載與使用指南 (Quick Start)

請前往 **[👉 GitHub Releases 最新發布頁面](https://github.com/lianghao02/PaperSwitch/releases/latest)** 下載最新版本：

| 下載檔案類型 | 檔案名稱 | 適用對象與說明 |
| :--- | :--- | :--- |
| 🌟 **免安裝獨立單檔版<br>（官方推薦）** | **`PaperSwitch-v4.6.0-Standalone.exe`** | 內嵌完整 .NET 8 執行環境，可直接開啟工坊；Office 文件轉換仍需本機安裝 Microsoft Office。 |
| 🛡️ **SHA-256 校驗清單** | **`SHA256SUMS.txt`** | 提供發行成品 SHA-256 雜湊值供安全性核對。 |

### 💡 首次啟動與使用須知
1. **Windows SmartScreen 提示**：
   - 首次執行時若 Windows 出現「Windows 已保護您的電腦」藍色保護畫面，請點選 **「其他資訊」** ➜ 點擊 **「仍要執行」** 即可正常啟動。
2. **Office 轉檔支援**：
   - PDF 與圖片（PNG, JPG, WebP, BMP）排版 **完全零環境依賴**。
   - 若需拖曳轉換 Word (`.docx`/`.doc`)、Excel (`.xlsx`/`.xls`) 或 PPT (`.pptx`/`.ppt`)，本機需安裝有 Microsoft Office 桌面軟體。

---

## 操作介面

目前為 2D Q 版文具工坊介面，遊戲風格與商業工具感並重。

上方放置加入檔案與復原／重做；左側工具盒保留雙欄排版按鈕、轉檔任務與工坊管理，可獨立捲動或收合。中央工作桌保留文件縮圖、選取狀態與縮圖大小滑桿。成果匯出固定於視窗底部，檔名、後綴設定及五個操作按鈕保持可見，按鈕空間不足時自動換行。雙擊紙張仍使用既有大圖檢視。
### 匯出檔名

「加上後綴」預設關閉，並記住上次設定。以「報告」為例，關閉時另存為 `報告.pdf`，拆分為 `報告_001.pdf`，圖片為 `報告_001.png`；開啟時選取合併為 `報告_選取頁面.pdf`，拆分與圖片會加上來源頁碼，例如 `報告_頁3_001.pdf`。匯出全部 PDF 不添加用途文字。

手動填寫檔名時優先使用輸入內容。檔名欄留空時：

| 匯出方式 | 自動命名 |
|---|---|
| 全合併匯出 | 使用畫布第一個有效來源名稱，例如 `案件資料.pdf` |
| 全拆匯出 | 各頁來源名稱＋流水號，例如 `案件資料_001.pdf` |
| 選取匯出 | 日期時間＋流水號，例如 `20261015_0930_001.pdf` |
| 轉成圖片匯出 | 各頁來源名稱＋流水號，例如 `案件資料_001.png` |
| 沒有有效來源名稱 | 使用日期時間備援；未命名空白頁不當作原始文件名稱 |

同一批拆分／圖片匯出共用同一時間。時間命名遇到同名成品時使用下一個可用流水號；來源／手動命名的同名檔案仍加上 `_1`、`_2` 防覆寫。原始檔名取自匯入來源，頁面經使用者重新命名後仍尊重該名稱。

介面不顯示檔名預覽，只保留「加上後綴」勾選框。命名偏好存於使用者本機資料目錄的 `PaperSwitch/export-preferences.json`，不保存來源文件路徑或內容。
## 📖 核心操作流程

1. **匯入檔案**：直接將 PDF、Word、Excel、PPT 或圖片檔案拖曳至工坊視窗，系統會自動在背景排程轉檔並載入每頁縮圖。左側操作面板的「轉檔任務」會逐檔顯示等待、處理中、完成、失敗或跳過狀態。
   - 失敗檔案不會阻斷後續文件；可直接選擇「重新嘗試」、「跳過」或「開啟位置」。
   - 介面顯示可理解的處理建議；需要查核時才展開「技術資訊」。原始檔案不會被修改。
2. **紙張排版與調整**：
   - **調整順序**：使用滑鼠拖曳卡片，或選取後按 `Ctrl + ←` / `Ctrl + →` 左右移動；按 `Home` / `End` 瞬移至首頁或末頁。
   - **復原與重做**：按 `Ctrl + Z` 復原，或按 `Ctrl + Y` 重做最近一次紙張編排動作。
   - **旋轉頁面**：選取卡片後按鍵盤 `R`（順時針 90°）或 `Shift + R`（逆時針 90°）。
   - **大圖檢視 (Lightbox)**：雙擊任意卡片開啟高解析燈箱，支援滾輪縮放（35% ~ 400%）與鍵盤翻頁；放大後按住滑鼠左鍵即可抓取拖曳，游標會由手掌變為抓取狀態。
   - **序號重新命名**：選取單頁或多頁卡片後按 `F2`，輸入基礎名稱即可批次自動流水編號。
3. **裝訂與匯出 PDF／圖片**：底部固定成果匯出區依序顯示輸出檔名、「加上後綴」勾選框及五個按鈕，不再收合：
   - **開啟輸出資料夾**：開啟成品資料夾。
   - **全合併匯出**：將全部頁面依畫布順序合併成一份 PDF，不受選取範圍影響。
   - **全拆匯出**：僅未選取縮圖時可用，將全部頁面各自輸出為獨立 PDF；有選取時停用。
   - **選取匯出**：將選取頁面合併成一份 PDF；沒有選取時停用。
   - **轉成圖片匯出**：有選取時輸出選取頁面，沒有選取時輸出全部頁面，格式為 PNG；滑鼠提示顯示實際範圍。
   - **PDF 頁碼**：勾選「PDF 加上頁碼」後，可選左下、下方置中、右下，以及 8、10、12、14、16、18、24 pt。預設關閉；開啟時預設下方置中、12 pt，依本次匯出順序從 1 編號，拆分 PDF 也連續編號。PNG 不套用此設定。頁碼疊在頁面底部，若原稿已有頁尾，請留意可能重疊。
   - 畫布沒有頁面時，四個匯出按鈕皆停用，仍可開啟輸出資料夾。停用按鈕保留可閱讀文字及原因提示。

---

## 🌟 核心特色與技術亮點

### 1. 2D Q 版文具工坊
- **文具工坊與清楚操作**：標誌與空白狀態保留手繪文具插畫；功能按鈕使用可縮放向量輪廓搭配文字，圖示繼承按鈕色彩，匯出按鈕以文字為主，維持原有操作配置。
- **內嵌素材與縮放**：原創圖集由內建 imagegen 製作，使用九宮格延伸材質並共用快取，不需要網路。素材與生成規格見 [工坊素材](dotnet-src/src/PaperSwitch/Assets/Workshop/README.md)。
- **遊戲按鍵與標題**：紙質／木質／苔綠按鍵與系統楷體標題，保留鍵盤焦點、可讀停用文字與原有功能名稱。
- **清楚的文件工作桌**：奶油紙色、焦糖框線、薄荷綠選取標記與杏桃橘主要操作；文件縮圖維持原始內容。
- **固定匯出與狀態列**：工具盒捲動時仍可操作成果匯出；停用按鈕保留可讀文字與原因提示。
### 2. 🗂️ 沉浸式「紙張排版工坊 (Page Arranger)」
- **混合格式拖入即轉**：支援 PDF、Word (`.docx`/`.doc`)、Excel (`.xlsx`/`.xls`)、PPT (`.pptx`/`.ppt`) 與圖片（`.png`/`.jpg`/`.webp`/`.bmp`）自由拖曳匯入，自動背景排程轉檔。
- **100% 向量無損合成**：採用 `PdfSharp` 向量矩陣旋轉與抽取，完全保留原始向量字型與文字層，零模糊、零檔案肥大。
- **Windows 原生超高清縮圖**：使用 Windows 10/11 內建 `Windows.Data.Pdf` (WinRT) 進行多執行緒非同步縮圖渲染與記憶體快取。
- **多選打包整批拖曳**：支援 `Ctrl+點擊`、`Shift+點擊` 複選，抓取任意選取卡片即可整包換位。
- **雙擊大圖燈箱 (Lightbox)**：
  - 預設「適合視窗」等比縮放，解決高解析頁面局部裁切盲區。
  - 直接滾動滑鼠滾輪即可平滑縮放（**35% ～ 400%**），以滑鼠座標為縮放錨點，閱讀細節不跳位。
  - 支援鍵盤左右翻頁（`←/→`）、旋轉（`R/Shift+R`）與原尺寸切換。
- **批次重新命名 (F2)**：選取單頁或多頁卡片後按下 `F2`，可批次指定基礎檔名並自動附加 3 位序列號。

### 3. ⌨️ 極速鍵盤導航操作
| 快捷鍵 | 功能描述 |
| :--- | :--- |
| **`Ctrl + Z` / `Ctrl + Y`** | 復原 / 重做最近一次手動編排動作（移動、旋轉、刪除、清空或重新命名） |
| **`Ctrl + ←` / `Ctrl + →`** | 將選取的紙張向前 / 向後移動 1 格（防覆寫無衝突交換演算法） |
| **`Home` / `End`** | 一鍵將選取的紙張瞬移至最前端（第一頁）或最末端 |
| **`R` / `Shift + R`** | 順時針 90° / 逆時針 90° 旋轉選取頁面 |
| **`Delete` / `Backspace`** | 移除選取的紙張 |
| **`Ctrl + A`** | 全選 / 取消全選切換 |
| **`F2`** | 重新命名選取頁面（支援多選序列化命名） |
| **`Ctrl + 滑鼠滾輪`** | 主畫布即時平滑縮放卡片尺寸（90px ~ 400px） |
| **滑鼠滾輪 (燈箱內)** | 大圖燈箱內平滑等比縮放（35% ~ 400%） |

### 4. 🛡️ Office COM 生命週期精準管理與環境隔離
- **專屬 STA 執行緒隔離**：Word、Excel、PowerPoint 轉檔均運行於獨立 STA 執行緒與全域 Semaphore 佇列。
- **確定性資源銷毀與 Null 守衛**：建立 COM 時即時進行安全 Null 檢查，退出時透過 `Marshal.FinalReleaseComObject` 與垃圾回收徹底銷毀，無任何背景殘留進程。
- **Excel 自動分頁與全空白過濾**：自動識別多工作表並獨立匯出，智慧過濾無資料與無形狀之全空白分頁，並支援頁面寬度自動縮放。
- **IGEF／非標準 PDF 偵測**：IGEF 不應視為微軟 Office 的一般暫存格式。預設 PDF 就緒等待上限為 90 秒，持續偵測到 IGEF 約 5 秒即拒絕載入；保留暫存輸出，提示使用者依單位規定取得可讀的標準 PDF 後重新匯入。
- **儲存與暫存健康管理**：產出匯出 PDF 存放於 `%LOCALAPPDATA%\PaperSwitch\converted`，Office 暫存存放於 `%LOCALAPPDATA%\PaperSwitch\temp_converted`。左側排版工具的「清除暫存」按鈕經確認後會清除這兩個資料夾的內容；需要保存的匯出成品請先另行保存。啟動時不會自動清除，原始匯入文件不受此按鈕影響。

---

## 📂 專案目錄結構

```text
09_PaperSwitch/
├── dotnet-src/
│   ├── PaperSwitch.sln
│   ├── src/
│   │   └── PaperSwitch/
│   │       ├── App.xaml / App.xaml.cs
│   │       ├── MainWindow.xaml / MainWindow.xaml.cs
│   │       ├── LightboxWindow.xaml / LightboxWindow.xaml.cs
│   │       ├── RenameDialog.xaml / RenameDialog.xaml.cs
│   │       ├── Models/ (PaperItem.cs, ExportOptions.cs, ConversionResult.cs)
│   │       ├── ViewModels/ (MainViewModel.cs, LightboxViewModel.cs)
│   │       ├── Services/
│   │       │   ├── PdfService.cs
│   │       │   ├── OfficeConverterService.cs
│   │       │   ├── ThumbnailCacheService.cs
│   │       │   ├── ImageConverterService.cs
│   │       │   ├── StorageMaintenanceService.cs
│   │       │   └── AppPaths.cs
│   │       ├── Styles/ (GameWorkshop.xaml, WarmCozyTheme.xaml, Colors.xaml)
│   │       └── PaperSwitch.csproj
│   ├── tests/
│   │   └── PaperSwitch.Tests/ (xUnit 單元測試)
│   └── scripts/
│       ├── build.ps1 (一鍵發行腳本)
│       ├── run.ps1 (啟動與成品新舊檢查)
│       ├── test-launcher.ps1 (隔離入口回歸測試)
│       └── qa.ps1 (自動化建置與測試檢核)
├── dist/
│   └── publish/ (編譯發行成品: PaperSwitch.exe)
├── RUN.bat (雙擊啟動 C# 原生應用程式)
├── README.md (專案說明)
└── CHANGELOG.md (版本變更歷程)
```

---

## 🚀 本地開發與建置

### 1. 直接啟動開發版本
雙擊專案根目錄的 **[`RUN.bat`](RUN.bat)**。BAT 僅選擇 PowerShell 主機並呼叫 `dotnet-src/scripts/run.ps1`。啟動器依序尋找有效的 `dist\publish`、`win-x64` Release、未指定 RID 的 Release 成品；若全部缺失或落後於真正的來源檔案，才執行建置。`bin`／`obj` 的生成檔案不會觸發重建，Release 成品以 `PaperSwitch.dll` 判斷新舊。支援 Windows PowerShell 5.1 與 PowerShell 7，路徑可包含空白或中文。

只檢查入口而不啟動／建置：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\dotnet-src\scripts\run.ps1 -ValidateOnly
```

成功時輸出可啟動的 EXE 路徑；需要建置時會回報錯誤。開發成品需要 .NET 8 Desktop Runtime，建置需要 .NET 8 SDK；免安裝獨立版另依上方下載說明。

### 2. 手動建置與發行
```powershell
# 編譯 Framework-Dependent 版 (發行至 dist/publish/PaperSwitch.exe；需 .NET 8 Desktop Runtime)
powershell -ExecutionPolicy Bypass -File .\dotnet-src\scripts\build.ps1

# 編譯 Self-Contained 單一免安裝獨立版 (包含完整 .NET 8 執行階段)
powershell -ExecutionPolicy Bypass -File .\dotnet-src\scripts\build.ps1 -SelfContained
```

### 3. 執行品質檢驗 (QA)
```powershell
powershell -ExecutionPolicy Bypass -File .\dotnet-src\scripts\qa.ps1

# 隔離入口測試：假成品只作時間檢查，不會啟動，也不會修改正式資料。
powershell -NoProfile -ExecutionPolicy Bypass -File .\dotnet-src\scripts\test-launcher.ps1 -PowerShellHost powershell.exe
pwsh -NoProfile -File .\dotnet-src\scripts\test-launcher.ps1 -PowerShellHost pwsh.exe
```

## 已知 Bug、限制與疑難排解

以下區分已確認問題、功能限制及待驗證項目；歷史修正不代表舊發行包已自動更新，也不代表本次文件更新重新完成所有功能測試。

| 狀態 | 情境 | 處理方式 |
|---|---|---|
| 已修復（本機啟動，2026-10-05） | 原先 publish 缺 EXE，且 bin／obj 生成檔案造成錯誤重建判定；Windows PowerShell 5.1 另有中文字元編碼問題。 | 已修正入口判斷及腳本編碼，補齊本機發布 EXE；兩種 PowerShell 入口檢查與主視窗啟動通過。正式獨立發行包維持原版。 |
| 環境限制 | Office 轉換依賴本機 Office 與 COM 元件。 | 獨立版包含 .NET Runtime，不代表隨附 Microsoft Office；PDF／影像與 Office 轉換要分別驗證。 |
| 辦公室加密環境限制（2026-10-06） | 本機 Word／Excel／PowerPoint 合成輸出均為 IGEF；GoPatrol 服務及保護驅動正在執行。 | 依單位核准流程取得可讀的標準 PDF 再匯入；如需直接轉 Office，由管理端確認允許的工具讀取流程。既有拒絕防護與標準 PDF 功能通過，正常 Office 轉檔仍待驗收；精確加密規則尚未確認。見 [診斷與處理方式](../00_Dev-Control-Center/docs/new-build-acceptance/IGEF-DIAGNOSIS.md)。 |
| 操作限制 | 選取頁面與全部頁面的處理範圍不同。 | 匯出前確認選取數、頁面順序及旋轉，勿只依賴縮圖判斷成品。 |

旋轉重複套用、文字輸入時 Delete／Backspace 被攔截等歷史修正見 [CHANGELOG.md](CHANGELOG.md)。本機入口修復與驗證範圍見 [修復報告](../00_Dev-Control-Center/docs/paperswitch-launch-repair/RESULTS.md)及 [HANDOFF.md](HANDOFF.md)；本輪未重新驗收所有 Office 轉換或其他電腦的操作流程。

### 問題回報

請提供使用版本／啟動方式、作業系統與相關環境、重現步驟、預期及實際結果，以及去識別的錯誤訊息或最小樣本。先保留現場與來源資料；不要附真實案件、完整帳號、密碼、Token 或 API Key。版本修正以對應原始碼與發行包為準。
