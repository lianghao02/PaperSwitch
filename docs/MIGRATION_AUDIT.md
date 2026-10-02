# Python → .NET 功能遷移稽核

稽核基準：`d458feb`
稽核日期：2026-09-21（更新於 v4.3.0 發布後）

本文件只比對使用者可感知功能。正式維護入口是 `dotnet-src/`；`legacy-python/` 已非正式執行入口，主要保留作歷史實作及尚未等價遷移之 LibreOffice fallback 參考。

| 功能 | Python | .NET | 狀態 |
|---|---|---|---|
| 檔案載入 | PDF、Office、常見圖片，Web 拖放／選取 | PDF、Office、常見圖片，WPF 拖放／選取 | 完整 |
| 紙張尺寸 | 保留來源 PDF／圖片尺寸 | 保留來源尺寸，另可插入 A4 空白頁 | 完整 |
| 頁面排列 | 單頁／多頁拖曳重排 | 單頁／多頁打包拖曳、鍵盤重排、復原／重做 | 完整 |
| 旋轉 | 單頁與多選 90° 旋轉 | 單頁與多選 90° 旋轉，匯出保留向量內容 | 完整 |
| 縮放 | 縮圖 90–480 px、預覽縮放 | 縮圖 90–480 px、預覽 35%–400% | 完整 |
| 邊界處理 | 無效 PDF、IGEF、檔名與拖曳邊界防禦 | 無效 PDF、IGEF、檔名、選取及拖曳邊界防禦 | 完整 |
| 留白／空白頁 | 無獨立空白頁功能 | 可插入標準 A4 空白頁 | 完整 |
| PDF 合併輸出 | 支援 | 支援依目前畫布順序向量合成 | 完整 |
| PDF 拆分輸出 | 支援原始 PDF 拆頁 | 支援將畫布紙張逐頁獨立輸出 | 完整 |
| PDF 轉圖片 | 可輸出高畫質 PNG | Windows 內建 Windows.Data.Pdf 原生渲染 PNG（2400px 高品質、保留旋轉、支援無選取全出／有選取單出或多出、不修改來源 PDF） | 完整 |
| 圖片轉 PDF | 支援 Pillow 轉換 | 支援 WPF 解碼後由 PdfSharp 建立 PDF | 完整 |
| Word／Excel／PowerPoint | COM 優先，另有 LibreOffice 後援 | Windows Office COM；Excel 可見有效工作表分頁處理 | 部分 |
| LibreOffice 後援 | 支援 | 未提供；正式規格要求安裝 Microsoft Office 桌面版 | 缺少 |
| 批次處理 | 依序處理並彙整成功／失敗 | 逐檔等待、處理中、完成、失敗、跳過，可重試 | 完整 |
| 錯誤提示 | API 回傳及 Web 訊息 | 可理解原因、補救動作與可展開技術資訊 | 完整 |
| 工作狀態 | 全域進度與日誌 | 全域進度加逐檔任務狀態 | 完整 |
| 更新 | GitHub Release 檢查與 Python 檔案熱更新 | GitHub Release 檢查、SHA-256 驗證、Standalone 替換重啟 | 完整 |
| 本機服務／瀏覽器介面 | Flask HTTP 服務與 Edge App | 原生 WPF，不再開啟本機 HTTP 服務 | 已淘汰 |
| LibreOffice／Docker 部署 | 舊版提供 | Windows 原生工具不採用 | 已淘汰 |
| 作業取消 | 未提供安全的單檔轉換取消 | 未提供；Office COM 工作完成前不可中斷 | 無法確認 |

## 結論

.NET 版已完整覆蓋 PaperSwitch 目前正式定位所需的匯入、視覺排版、旋轉、拆分、合併、批次狀態、Office COM 流程與批次 PDF 轉圖片 (PNG)，且增加復原／重做、空白頁及原生更新。

目前主要剩餘功能差異僅剩 LibreOffice 後援。`legacy-python/` 已非正式執行入口，主要保留作歷史實作及尚未等價遷移之 LibreOffice fallback 參考，不納入主構建亦不急於刪除。是否需要遷移應由未來產品規格明確決定。
