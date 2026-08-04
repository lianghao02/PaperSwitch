# plan.md — 實作路線與技術規劃

## 階段一：核心轉換引擎開發 (已完成)
- 實作 `DocumentConverter` 類別，透過 `win32com.client` 存取 MS Office COM。
- 整合 `Pillow` 進行圖片無損 PDF 封裝。

## 階段二： Web UI 與伺服器整合 (已完成)
- 以 Python 內建 `HTTPServer` 打造零額外重量級依賴的 API。
- 嵌入深色玻璃擬物 UI HTML/CSS/JS 範本。

## 階段三：包裝與部署優化 (進行中)
- 提供 `.env` 設定檔支援。
- 支援以 PyInstaller 打包為單一 `.exe` 免安裝檔。
