# 📝 變更歷史 (CHANGELOG)

本專案遵循 [Keep a Changelog](https://keepachangelog.com/zh-TW/1.0.0/) 結構與 [Semantic Versioning](https://semver.org/lang/zh-TW/) 規範。

---

## 🚀 v1.7.1 (2026-08-23)

### 🔤 視覺字體清晰度與層級階梯大升級 (Typography Hierarchy)
- **H1 / H2 / H3 視覺階梯明確化**：
  - **H1 主標題 (`header h1`)**：提升至 `2.1rem`，採用高對比純白光影漸層與陰影。
  - **H2 區塊標題 (`.panel-title`)**：放大至 `1.22rem`，純白高對比加上 `2px` 霧天藍底線錨點。
  - **H3 選項/卡片標題 (`.radio-title`, `.drop-text`)**：加粗至 `800 ExtraBold`，提升至 `1.02rem ~ 1.16rem`。
  - **檔案清單與內文**：檔案名稱提升至純白 `700 Bold`，輔助說明文字提升對比度至 `#94a3b8`，告別暗淡模糊。
- **全域字型堆疊與渲染優化**：導入現代清晰的系統字型堆疊，啟用 `-webkit-font-smoothing` 與 `-moz-osx-font-smoothing`，在深色莫蘭迪背景下文字極致清晰銳利。

---

## 🚀 v1.7.0 (2026-08-23)

### 🎨 介面版面全新重構 (80% 寬度與 3-Column 直式流水線)
- **80% 視窗寬度排版**：介面總寬度升級為 `82vw`（約 80% 螢幕寬度），最大寬度 `1600px`，消除寬螢幕過度留黑，視覺大器沉穩。
- **1 ➔ 2 ➔ 3 人性化三欄直式佈局**：
  - **左欄（📥 1. 檔案拖曳區）**：大面積垂直等高拖曳上傳與格式導引。
  - **中欄（📋 2. 檔案佇列與進度）**：整合總體真進度條、檔案卡片佇列與即時動態日誌視窗。
  - **右欄（⚙️ 3. 轉換功能設定）**：四大模式選擇、合併檔名設定、🚀 開始轉換按鈕與維護工具區。
- **操作視線極致優化**：由左至右順暢推進「拖曳 ➔ 檢視 ➔ 轉換」，三欄等高自適應，徹底免除外層滾動條。

---

## 🚀 v1.6.0 (2026-08-19)

### ✂️ 新功能：PDF 拆單頁 (Split PDF to Single-Page PDFs)
- **`DocumentConverter.split_pdf`**：導入 `pypdf` 無損向量抽取演算法，支援將多頁 PDF 文件一鍵獨立拆分導出為 `{主檔名}_第1頁.pdf`、`{主檔名}_第2頁.pdf` 等獨立單頁 PDF 檔案。
- **全格式自動向下相容**：若在拆單頁模式上傳 Word/Excel/PPT 或圖片，系統會自動輔助轉換為臨時 PDF 後再完成單頁拆分。
- **四大模式對稱工整命名**：
  - 📄 **文件轉 PDF** (`single`)
  - 📚 **多檔併 PDF** (`merge`)
  - ✂️ **PDF 拆單頁** (`split`)
  - 🖼️ **PDF 轉圖片** (`pdf_to_images`)
- **流水線動態日誌與進度連動**：前端即時反饋拆單頁步進進度與產出路徑。

---

## 🚀 v1.5.0 (2026-08-19)

### 🖥️ 原生桌面級體驗重大升級 (Desktop Experience & Lifecycle)
- **關閉網頁自動關閉伺服器 (Heartbeat Watchdog)**：實作雙向心跳守護機制，當所有瀏覽器分頁關閉超過 12 秒後自動安全終止後端行程；自帶 F5 重新整理保護，絕不誤殺。
- **網頁端即時動態日誌 (Activity Feed)**：於介面底部整合「🖥️ 系統執行日誌」終端視窗，即時推播後台 Word/Excel 轉檔路徑、分頁拆分與清理訊息，無需查看 CMD。
- **CMD 黑色視窗自動隱藏 (pythonw 靜默運行)**：升級 `RUN.bat` / `setup_and_run.ps1` 啟動流程，自癒檢查並喚起瀏覽器後自動關閉命令列視窗，完全無黑窗殘留。
- **維護工具「🛑 關閉伺服器」按鈕**：增設一鍵停止伺服器 API 與安全退出按鈕。

---

## 🚀 v1.4.0 (2026-08-19)

### 🎨 視覺與互動重大升級 (Morandi Design & Real Progress)
- **莫蘭迪色彩規範 (Morandi Palette)**：全域視覺升級為沉穩溫潤的莫蘭迪低飽和色系（霧天藍 `#6ba4c8`、鼠尾草綠 `#7ea88f`、灰杏黃 `#d1a368`、豆沙灰紅 `#d47a7a`、霧深藍灰背景 `#161c28`），徹底消除高飽和光害。
- **逐檔真實進度流水線 (Live Progress)**：重構單檔/轉圖處理為非同步流水線，進度條精確隨完成檔數步進，卡片狀態標籤即時連動 `⏳ 轉檔中`（自帶呼吸發光動態）與 `✅ 已完成`。
- **Excel 空白分頁智慧過濾**：導入 `WorksheetFunction.CountA` 與 `ws.Shapes.Count` 雙重檢測，自動過濾 100% 無內容與無圖片的空白工作表，避免產出多餘空白 PDF。

---

## 🚀 v1.3.0 (2026-08-19)

### 🖼️ 新功能：PDF 反向轉圖片 (PDF to Images)
- **`DocumentConverter.convert_pdf_to_images`**：整合本機 PyMuPDF (`fitz`) 高清渲染引擎，支援將 PDF 檔案的每一頁獨立拆分導出為高畫質 200 DPI PNG 圖片。
- **Web UI 介面模式整合**：在「⚙️ 3. 轉換功能設定」新增 **🖼️ PDF 轉圖片模式** 選項卡片，支援一鍵切換與雙向轉換（Office/圖片 ⇄ PDF）。
- **轉檔自動向下相容處理**：於「PDF 轉圖片模式」下若上傳 Word/Excel/PPT 文件，系統會自動輔助轉換為臨時 PDF 後再拆分出高畫質 PNG 圖片。

---

## 🚀 v1.2.0 (2026-08-17)

### 新增與文件
- **獨立啟動**：加入 Python 3.13 可攜環境自癒流程，可從 `RUN.bat` 建置並啟動。
- **依賴說明**：區分 Office COM、圖片轉換與 PDF 合併各自需要的軟體。
- **入口修正**：README 改指向實際存在的 `RUN.bat`，並補齊移機及離線建置方式。

## 🚀 v1.1.1 (2026-08-13)

### 🔴 高風險 Bug 修復與安全性提升
- **`_parse_multipart` 二進位盲裁修復**：徹底移除原先對 `data` 盲裁 `data.endswith(b"--")` 的危險邏輯，改為 RFC 7578 二進位安全邊界解析，解決二進位圖片與 PDF 尾端包含 `0x2D 0x2D` 時檔尾遭誤裁損毀的問題。
- **Win32COM 併發互斥鎖 (`COM_LOCK`)**：導入全域互斥鎖，解決 `ThreadingHTTPServer` 多執行緒併發調用 MS Office COM 時引發 STA `BUSY` (0x80010002) 例外與伺服器死鎖 (Deadlock)。
- **COM 資源釋放防禦**：在 `DocumentConverter` 的 `finally` 區塊加入雙層例外防禦，確保 Close / Quit / CoUninitialize 無論是否異常皆安全釋放。
- **PDF Writer 安全關閉**：將 `merge_pdfs` 的 `writer.close()` 置於 `finally` 區塊，確保 I/O 異常時控制代碼不遺留洩漏。

---

## 🚀 v1.1.0 (2026-08-10)

- **核心架構升級**：支援多檔併發上傳與快取最佳化。
- **快取目錄重構**：移除非必要依賴，統一暫存與輸出路徑。

---

## 🚀 v1.0.0 (2026-08-04)

- **PaperSwitch 萬能文件轉 PDF 處理器正式發布**：
  - 支援 Word (.docx, .doc)、Excel (.xlsx, .xls)、PowerPoint (.pptx, .ppt)、圖片 (.png, .jpg, .bmp, .webp) 與 PDF 無損轉換與合併。
  - Excel 智慧分頁與欄位自適應適應單頁寬度 (`FitToPagesWide = 1`)。
  - 支援獨立單檔模式與多檔合併 PDF 模式。
  - 採用 ThreadingHTTPServer 與深色玻璃擬物視窗 UI (Dark Glassmorphic UI)。
