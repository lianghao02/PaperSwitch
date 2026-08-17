# MEMORY.md — 經驗持久化與 Bug 紀錄

## 📌 持久化經驗與 Bug 坑洞
- **Python 3.13+ `cgi` 模組廢棄移除坑**：Python 3.13 / 3.14 依據 PEP 594 正式移除了標準庫 `cgi` 模組（`ModuleNotFoundError: No module named 'cgi'`）。**處方**：在 `app.py` 中實作原生零依賴的 `_parse_multipart()` 方法，直接解析 `multipart/form-data` Byte Stream，徹底解決版號相容性問題。
- **Windows BAT 檔案 UTF-8 亂碼解析坑**：Windows `cmd.exe` 預設採 ANSI (CP950) 解析批次檔。若 `.bat` 包含多位元 UTF-8 Emoji 特殊字元或在 `if (...)` 區塊中使用 `::` 註解，CMD 解析器會產生斷詞錯誤（例如 `python app.py` 被誤拆成 `pp.py`）。**處方**：批次檔一律採 CP950/ANSI 編碼保存，改用 `REM` 註解並移除 Emoji 特殊符號。
- **Excel 欄位斷頁坑**：已於 `DocumentConverter.convert_excel` 加入 `ws.PageSetup.FitToPagesWide = 1`，強制將 Excel 欄位自動縮放至單頁 PDF 寬度，解決原本多頁錯位問題。
- **Word COM 執行緒鎖定坑**：在多執行緒或 Web Request 中呼叫 Win32COM 時，必須顯式呼叫 `pythoncom.CoInitialize()` 與 `CoUninitialize()`，避免 `CoInitialize has not been called` 崩潰。
- **PNG 透明背景坑**：PNG 圖片帶有 Alpha 透明通道，直接儲存為 PDF 會引發 `OSError: cannot write mode RGBA as PDF`。已加入 `img.convert('RGB')` 自動預處理。

## 📦 外部依賴追蹤
| 依賴套件 / 片段 | 來源 / 版本 | 目的 |
|---|---|---|
| `pywin32` | PyPI / >=306 | 呼叫 Windows MS Office COM Automation 引擎 |
| `Pillow` | PyPI / >=10.0.0 | 圖片讀取與無損 PDF 封裝 |
| `python-dotenv` | PyPI / >=1.0.0 | `.env` 環境變數載入 |

## ⚡️ 上游衝突紀錄
- 目前為獨立初始專案，尚無上游衝突。

## 🔖 GitHub 借鏡清單
- [Stirling-Tools/Stirling-PDF](https://github.com/Stirling-Tools/Stirling-PDF) ⭐40,000+ — 借鏡 Web UI 拖拽卡片排版模式。
- [Aluunt/docx2pdf](https://github.com/Aluunt/docx2pdf) ⭐800+ — 借鏡 Win32COM Automation 呼叫機制。

## 📅 學習歷史
- **2026-08-04**：完成專案初始設計、全功能 `app.py` 整合與《全域開發憲法 v7.1》9 核心文件建置。
