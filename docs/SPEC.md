# spec.md — 需求規格與功能邊界

## 1. 產品範圍

PaperSwitch 是 Windows 原生 C# 12／.NET 8 WPF 工具，讓使用者匯入多種文件後，在視覺畫布中無損整理並導出 PDF。

- 匯入：PDF、Word（`.doc`／`.docx`）、Excel（`.xls`／`.xlsx`）、PowerPoint（`.ppt`／`.pptx`）與圖片（`.png`／`.jpg`／`.webp`／`.bmp`）。
- Office 轉檔：透過 Microsoft Office COM，使用專屬 STA 執行緒與序列化保護；使用者電腦須安裝對應 Office 桌面軟體。
- 編排：多選、拖曳重排、鍵盤移動、旋轉、刪除、重新命名與復原／重做。
- 預覽：Windows `Windows.Data.Pdf` 非同步縮圖與燈箱檢視。
- 導出：以 PdfSharp 抽取原始 PDF 頁面並套用旋轉矩陣，保持向量、字型與文字層。

## 2. 非功能需求

- 支援 Windows 10／11 x64 與 `net8.0-windows10.0.19041.0`。
- 匯入及導出需持續提供使用者可見的處理狀態；大量檔案不可表現為無回應。
- 暫存檔存放於 `%LOCALAPPDATA%\PaperSwitch\temp_converted`，導出檔存放於 `%LOCALAPPDATA%\PaperSwitch\converted`。
- IGEF 或受保護的中介 PDF 不得繞過保護；應保留暫存檔並提示依規定解密後重新匯入。
- 不得以縮圖或點陣重壓取代原始 PDF 頁面，避免犧牲輸出品質。
