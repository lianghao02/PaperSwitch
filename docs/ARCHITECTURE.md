# 系統架構與資料流向

## 1. 系統整體架構

本專案採 Single-Process Lightweight Web Microservice 架構：

```
[前端 Web UI (深色玻璃擬物)]
         │
         │ HTTP Multipart POST /api/convert
         ▼
[HTTP 請求分發器 (WebAppHandler)]
         │
         ├───────────────┬───────────────┐
         ▼               ▼               ▼
   [Word 轉檔器]   [Excel 轉檔器]   [圖片轉檔器]
  (MS Word COM)   (MS Excel COM)      (Pillow)
         │               │               │
         └───────────────┼───────────────┘
                         ▼
             [產出 PDF 至 converted/]
```

## 2. 模組職責
- **`WebAppHandler`**：負責路由分發（`/` 渲染 HTML、`/api/convert` 處理檔案接收與轉檔呼叫）。
- **`DocumentConverter`**：
  - `convert_word`：利用 `win32com.client` 呼叫本機 Word 渲染產出 PDF (wdFormatPDF = 17)。
  - `convert_excel`：呼叫 Excel 引擎匯出 PDF (xlTypePDF = 0)，並自動校正頁面寬度自適應。
  - `convert_image`：利用 `Pillow` 將單張或多張圖片無損合成 PDF。
