# CHANGELOG.md — PaperSwitch 版本歷程與更新日誌
<!-- 本檔案遵循 Keep a Changelog 結構與 Semantic Versioning 規範 -->

## 🏆 v1.0.0 里程碑：PaperSwitch - 萬能文件轉 PDF 處理器發布

### 重大更新摘要
依使用者需求，**正式廢棄 PDF 拖拽排序與視覺化頁面編輯器功能**，系統完全還原至純粹且極致流暢的 Word/Excel/圖片/PDF 萬能轉換與合併小工具。

### ✨ 重點更新特色
- 🖥️ **三區劃分極致 UI (Dashboard)**：
  - 介面劃分為 **「拖曳區」**、**「佇列與進度區」**、**「功能設定區」** 三大區塊。
  - 佇列區內建動態進度條 (0%~100%) 與各檔案狀態標籤（待轉換、轉檔中、已完成、失敗）。
- ⚡ **ThreadingHTTPServer 多執行緒併發處理**：
  - 後端伺服器升級為多執行緒 `ThreadingHTTPServer`，大幅提升大量檔案排隊處理時的 HTTP 併發響應速度。
- 📊 **PowerPoint 簡報 (.pptx, .ppt) 轉 PDF 支援**：
  - 新增 MS PowerPoint COM / LibreOffice 簡報轉檔引擎，達成 Word + Excel + PPT + 圖片 4 大主流格式全能轉換與合併。
- 🧹 **「清空歷史暫存檔」使用者自選手動按鈕**：
  - 新增 `/api/clear-storage` API 與 `🧹 清空歷史暫存檔` 按鈕，讓使用者手動彈性決定何時掃除 `uploads/` 與 `converted/` 舊暫存檔。
  - 當 Excel 檔案包含多個工作表（例如：`工作表1`、`工作表2`、`工作表3`）時，系統自動將每一個分頁獨立拆分匯出為單獨的 PDF 檔案（例如：`檔名_工作表1.pdf`、`檔名_工作表2.pdf`）。
  - 每個獨立分頁同步套用 **「水平欄寬 1 頁寬，垂直依列數自然延伸」** 的不裁切自動縮放保護。
- 📚 **獨立單檔 vs 合併 PDF 雙模式**：
  - **獨立單檔模式**：佇列中每個檔案分別轉為同名 PDF。
  - **合併 PDF 模式**：將佇列中所有 Word/Excel/圖片/PDF 檔案依序合併產出為單一 PDF 檔。
- 📂 **一鍵開啟輸出資料夾**：
  - 轉檔完成後自動出現 **「📂 開啟 PDF 輸出資料夾」** 按鈕。
  - 點擊後直接呼叫 Windows 原生檔案總管開啟 `converted/` 目錄。
- 🚀 **Word/Excel 完美轉檔**：
  - 整合 MS Office COM Automation，100% 還原字型與頁面配置。
  - Excel 自動開啟 `FitToPagesWide` 自適應，杜絕欄位被剪裁錯位。
- 🎨 **現代化深色玻璃 UI**：
  - 打造高顏值深色玻璃擬物視窗（Dark Glassmorphic UI）。
  - 支援全功能檔案拖拽上傳區與即時狀態回饋。
- 📦 **Config-First 集中管理**：
  - 內建 `.env` 環境變數設定與預設伺服器埠號配置。
