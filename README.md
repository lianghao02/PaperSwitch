# 📑 PaperSwitch 紙張排版工坊 v4.0.0

[![Release](https://img.shields.io/github/v/release/lianghao02/PaperSwitch?color=orange&label=Release)](https://github.com/lianghao02/PaperSwitch/releases/latest)
[![Platform](https://img.shields.io/badge/platform-Windows%20.NET%208%20LTS-blue.svg)](https://dotnet.microsoft.com/)
[![Theme](https://img.shields.io/badge/style-Warm_Cozy_Craft-E28445.svg)](README.md)
[![Constitution](https://img.shields.io/badge/Constitution-v8.1-purple.svg)](https://github.com/lianghao02/home)

> **手帳質感的 Windows 原生文件轉 PDF、批次合併與視覺化紙張排版工坊。**<br>
> 全面重構自 Python 服務架構，升級為 C# 12 / .NET 8.0 LTS / WPF 原生應用程式，帶來秒開速度、確定性 Office COM 生命週期防禦與極致流暢的紙張排版體驗。

---

## 📥 快速下載與使用指南 (Quick Start)

請前往 **[👉 GitHub Releases 最新發布頁面](https://github.com/lianghao02/PaperSwitch/releases/latest)** 下載最新版本：

| 下載檔案類型 | 檔案名稱 | 適用對象與說明 |
| :--- | :--- | :--- |
| 🌟 **免安裝獨立單檔版<br>（強烈推薦）** | **`PaperSwitch-v4.0.0-Standalone.exe`** | **最簡單方便！** 內嵌完整 .NET 8 執行環境，**免安裝任何軟體**，下載後直接雙擊即可開啟工坊使用。 |
| ⚡ **極致輕量版** | **`PaperSwitch-v4.0.0-FrameworkDependent.zip`** | 體積僅約 26MB；適合本機已安裝 [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0) 的使用者，解壓縮後執行 `PaperSwitch.exe`。 |

### 💡 首次啟動與使用須知
1. **Windows SmartScreen 提示**：
   - 首次執行時若 Windows 出現「Windows 已保護您的電腦」藍色保護畫面，請點選 **「其他資訊」** ➜ 點擊 **「仍要執行」** 即可正常啟動。
2. **Office 轉檔支援**：
   - PDF 與圖片（PNG, JPG, WebP, BMP）排版 **完全零環境依賴**。
   - 若需拖曳轉換 Word (`.docx`/`.doc`)、Excel (`.xlsx`/`.xls`) 或 PPT (`.pptx`/`.ppt`)，本機需安裝有 Microsoft Office 桌面軟體。

---

## 📖 核心操作流程

1. **匯入檔案**：直接將 PDF、Word、Excel、PPT 或圖片檔案拖曳至工坊視窗，系統會自動在背景排程轉檔並載入每頁縮圖。
2. **紙張排版與調整**：
   - **調整順序**：使用滑鼠拖曳卡片，或選取後按 `Ctrl + ←` / `Ctrl + →` 左右移動；按 `Home` / `End` 瞬移至首頁或末頁。
   - **旋轉頁面**：選取卡片後按鍵盤 `R`（順時針 90°）或 `Shift + R`（逆時針 90°）。
   - **大圖檢視 (Lightbox)**：雙擊任意卡片開啟高解析燈箱，支援滾輪平滑縮放（35% ~ 400%）與鍵盤翻頁。
   - **序號重新命名**：選取單頁或多頁卡片後按 `F2`，輸入基礎名稱即可批次自動流水編號。
3. **匯出 PDF**：點擊右上方「匯出為 PDF」按鈕，即可瞬間產出 100% 向量無損合成的 PDF 文件。

---

## 🌟 核心特色與技術亮點

### 1. 🌾 溫暖手作插畫手帳風格 (Warm Cozy Craft Style)
- **水彩紙質感視覺**：厚磅水彩紙米白底（`#F7F4ED`）、陶土暖橘（`#D97736`）、森林苔綠（`#5B8266`）與墨黑（`#2D2825`），長時間辦公閱讀舒適不刺眼。
- **細緻手作陰影與圓角**：柔和立體卡片、手繪風標籤、流暢動畫與溫暖的懸停特效。

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
- **IGEF 中介狀態偵測**：自動偵測微軟 Office 暫存狀態，最長等待 90 秒；若受端點加密保護，保留暫存 PDF 並引導使用者依規定解密後重新匯入，絕不嘗試繞過保護。
- **儲存與暫存健康管理**：產出匯出 PDF 存放於 `%LOCALAPPDATA%\PaperSwitch\converted`，Office 暫存存放於 `%LOCALAPPDATA%\PaperSwitch\temp_converted`，並具備 7 天過期暫存自動清理機制。

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
│   │       ├── Styles/ (WarmCozyTheme.xaml, Colors.xaml)
│   │       └── PaperSwitch.csproj
│   ├── tests/
│   │   └── PaperSwitch.Tests/ (xUnit 單元測試)
│   └── scripts/
│       ├── build.ps1 (一鍵發行腳本)
│       └── qa.ps1 (自動化建置與測試檢核)
├── legacy-python/ (原 Python 舊架構備援封存)
├── dist/
│   └── publish/ (編譯發行成品: PaperSwitch.exe)
├── RUN.bat (雙擊啟動 C# 原生應用程式)
├── README.md (專案說明)
└── CHANGELOG.md (版本變更歷程)
```

---

## 🚀 本地開發與建置

### 1. 直接啟動開發版本
雙擊專案根目錄的 **[`RUN.bat`](RUN.bat)**，若尚未編譯會自動執行建置並直接啟動 `dist\publish\PaperSwitch.exe`。

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
```
