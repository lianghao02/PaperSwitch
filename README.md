# 📄 PaperSwitch - 萬能文件轉 PDF 處理器

[![Version](https://img.shields.io/badge/version-v1.0.0-blue.svg)](https://github.com/lianghao02/PaperSwitch)
[![GitHub Pages](https://img.shields.io/badge/GitHub%20Pages-Live-green.svg)](https://lianghao02.github.io/PaperSwitch/)
[![Constitution](https://img.shields.io/badge/Constitution-v7.1-purple.svg)](https://github.com/lianghao02/home)

> **PaperSwitch** 是一款高效、強健、100% 本機資安防護的全能文件轉 PDF 與 PDF 合併處理器。
> 🌐 線上展示頁面：[https://lianghao02.github.io/PaperSwitch/](https://lianghao02.github.io/PaperSwitch/)

---

## 🌟 核心特色 (Key Features)

- 📄 **4 大主流公務/商用格式無損轉換**：
  - **Word** (`.docx`, `.doc`)
  - **Excel** (`.xlsx`, `.xls`)
  - **PowerPoint** (`.pptx`, `.ppt`)
  - **圖片** (`.png`, `.jpg`, `.jpeg`, `.bmp`, `.webp`)
  - **PDF** (直出與多檔強健合併)
- 📊 **Excel 智慧分頁處方**：
  - 水平欄位強制自動縮放適應單頁寬度 (`FitToPagesWide = 1`)，右側表格絕不硬切裁離。
  - 多工作表（多分頁）自動獨立拆分導出為 `檔名_分頁名.pdf`。
- ⚡ **ThreadingHTTPServer 多執行緒併發引擎**：
  - 大量檔案拖拽處理秒速響應，多佇列平行轉檔。
- 🛡️ **100% 本機資安防護**：
  - 採用原生 MS Office COM 與 Pillow 引擎本機渲染，檔案零外洩風險。
- 🧹 **彈性手動清理**：
  - 提供 `🧹 清空歷史暫存檔` 手動控制按鈕，隨心所欲整理本機硬碟。

---

## 🚀 快速開始 (Quick Start)

### 雙擊啟動 (Windows)
雙擊執行 [`start.bat`](file:///C:/Users/chia-hao/Documents/GitHub/PaperSwitch/start.bat) 即可自動辨識 Python 虛擬環境並開啟 Web 控制台！

### 手動啟動
```powershell
python app.py
```
預設伺服器位置：`http://127.0.0.1:8080`
