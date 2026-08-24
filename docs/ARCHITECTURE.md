# 系統架構與資料流向

## 1. 系統整體架構

本專案以單一 `app.py` 提供本機 HTTP 服務、轉檔引擎與前端介面；服務預設僅綁定 `127.0.0.1`，Office COM 操作由全域 `COM_LOCK` 序列化保護。

```
[前端 Web UI（暖色手作工坊）]
         │
         ├─ POST /api/convert          → 四種轉換模式
         ├─ POST /api/arranger/render  → PDF 縮圖渲染
         └─ POST /api/arranger/export  → 向量 PDF 匯出
         ├─ GET /api/update/check      → GitHub Release 版本檢查
         └─ POST /api/update/perform   → 驗證、備份與套用 app.py 更新
                         ▼
     uploads/（請求暫存）與 converted/（產出及編排來源 PDF）
```

## 2. 模組職責
- **`WebAppHandler`**：負責主畫面、API 路由、存取權杖驗證與 multipart 檔案接收。
- **更新器**：僅讀取固定 GitHub Release 的 `app.py` 與 `version.json` 資產，先比對 SHA-256 與語法，再以備份及原子替換套用更新。
- **`DocumentConverter`**：
  - `convert_word`：利用 `win32com.client` 呼叫本機 Word 渲染產出 PDF (wdFormatPDF = 17)。
  - `convert_excel`：呼叫 Excel 引擎匯出 PDF (xlTypePDF = 0)，並自動校正頁面寬度自適應。
  - `convert_image`：利用 `Pillow` 將單張或多張圖片無損合成 PDF。
  - `split_pdf`、`convert_pdf_to_images`：處理 PDF 拆頁與轉圖片。
  - `render_pdf_thumbnails`、`export_arranged_pdf`：提供紙張編排工坊的縮圖與無損匯出。

## 資料生命週期

1. 上傳檔案依請求識別碼存入 `uploads/`。
2. 一般轉檔輸出至 `converted/`；合併用途的暫存 PDF 在合併後清除。
3. 編排工坊需保留其轉換後 PDF，直到匯出或使用者透過維護工具清理，才能重新讀取原始向量頁面。
4. 一鍵更新只替換 `app.py` 與 `version.json`；`uploads/`、`converted/`、內嵌 Python 與使用者設定不在更新範圍內。
