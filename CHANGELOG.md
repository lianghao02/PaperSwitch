# 📝 變更歷史 (CHANGELOG)

本專案遵循 [Keep a Changelog](https://keepachangelog.com/zh-TW/1.0.0/) 結構與 [Semantic Versioning](https://semver.org/lang/zh-TW/) 規範。

---

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
