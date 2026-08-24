# 📑 紙飛機 · PaperSwitch 文件與紙張轉換所 v3.0.0

[![Version](https://img.shields.io/badge/version-v3.0.0-orange.svg)](CHANGELOG.md)
[![Theme](https://img.shields.io/badge/style-Warm_Cozy_Craft-E28445.svg)](README.md)
[![Python](https://img.shields.io/badge/Python-3.13-blue.svg)](https://www.python.org/)
[![Constitution](https://img.shields.io/badge/Constitution-v8.0-purple.svg)](https://github.com/lianghao02/home)

> **樸實、溫暖、有親和力的本機文件轉 PDF、批次合併與視覺化頁面排版工坊。**<br>
> 專為公務人員、行政助理與辦公室工作者打造，在木質工作桌上輕鬆完成文件轉換、紙張重排與裝訂成冊。

---

## 🌟 核心特色與亮點

### 1. 🌾 樸實溫潤的「手作插畫手帳風格 (Warm Cozy Craft Style)」
- **告別冷冽科技感**：全域採用厚磅水彩紙米白底（`#F7F4ED`）、陶土暖橘（`#D97736`）與森林苔綠（`#5B8266`），長時間辦公閱讀舒適不刺眼。
- **親切生活感手作語彙**：以「文件投遞區」、「待整理清單」、「裝訂工坊」取代生硬的工程用語。

### 2. 🗂️ 沉浸式「紙張排版工坊 (PDF Page Arranger)」
- **全幅工作大畫布**：支援 PDF / Word / Excel / PPT / 圖片等格式自由混合拖入。
- **無損向量合成**：底層採用 PyMuPDF 即時渲染高畫質縮圖，並以 `pypdf` 頁面矩陣旋轉與抽取，100% 保留原始字型向量與文字搜尋能力。
- **📦 多選打包一起拖曳**：勾選或複選多張紙張，滑鼠抓取任意一張即可「整包打包移動」，並帶有立體懸浮徽章。
- **🚀 邊緣平滑自動滾動**：拖曳靠近畫布頂部或底部時，物理動態平滑加速滾動，上百頁也能一滑到底。
- **🔍 獨立縮圖滑桿與滾輪縮放**：支援滑桿縮放（90px ~ 480px）與工作區 `Ctrl + 滑鼠滾輪` 平滑放大縮小。
- **🖼️ 雙擊全螢幕大圖預覽燈箱**：雙擊紙張即時彈出高畫質大圖預覽，支援鍵盤 `←/→` 翻頁、`R / Shift+R` 旋轉與 `ESC` 關閉。

### 3. ⌨️ 極速「鍵盤換位導航引擎 (Keyboard Reordering)」
- **`Ctrl + ←` / `Ctrl + →`**（或 `Alt + ←/→`）：將選取的紙張向前/向後移動 1 格。
- **`Home` / `End`**：一鍵將選取的紙張瞬移至最前端（P.1）或最末端。
- **`←` / `→`**：在紙張間快速切換選取焦點，畫面自動平滑捲動追蹤。
- **`R` / `Shift + R`**：順時針 90° / 逆時針 90° 旋轉。
- **`Delete`**：移除選取的紙張。
- **`Ctrl + A`**：全選 / 取消全選。

### 4. 🚀 智慧自癒啟動系統 (`setup_and_run.ps1` / `RUN.bat`)
- **零 Python 前置安裝需求**：雙擊 `RUN.bat` 即可自動偵測環境，無 Python 時自動建置專案內可攜式 `python_embed` 環境並配置依賴。
- **進程殘留自癒清理**：啟動前自動檢查並釋放可能佔用 port 8080 的舊進程，避免連線衝突。

---

## 📂 支援檔案格式

| 格式分類 | 支援副檔名 | 轉換機制 |
| :--- | :--- | :--- |
| **Word 文件** | `.docx`、`.doc` | 本機 MS Office Word COM Automation（無損排版） |
| **Excel 試算表** | `.xlsx`、`.xls` | 本機 MS Office Excel COM（支援多分頁獨立拆分與自動寬度適應） |
| **PowerPoint 簡報**| `.pptx`、`.ppt` | 本機 MS Office PowerPoint COM（投影片無損轉 PDF） |
| **圖片圖紙** | `.png`、`.jpg`、`.jpeg`、`.bmp`、`.webp` | Pillow 高清無損封裝為 PDF |
| **PDF 文件** | `.pdf` | 直接加入合併、拆分或排版工坊進行頁面重組 |

---

## 🚀 快速開始

### 推薦方式（一鍵啟動）
1. 下載專案壓縮包並解壓縮。
2. 雙擊執行根目錄的 **[`RUN.bat`](RUN.bat)**。
3. 瀏覽器將自動開啟工坊介面：`http://127.0.0.1:8080`。

### 手動環境啟動
```powershell
# 建立虛擬環境並安裝相依套件
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

# 啟動服務
python app.py
```

## 🔄 GitHub Release 一鍵更新

在「工坊工具箱」按下 **🔄 檢查更新**，PaperSwitch 會比對本機 `version.json` 與 GitHub 最新 Release。若有較新的穩定版，使用者可直接下載、驗證並套用 `app.py`，完成後服務會自動重新啟動。

更新前會建立 `app.py.bak` 與 `version.json.bak`。若下載、SHA-256 或 Python 語法驗證失敗，原本版本不會被覆寫。

### 發布者的 Release 資產契約

每個可直接更新的 GitHub Release 必須附上以下兩個資產：

1. `app.py`
2. `version.json`，內容至少包含版本號與 `app.py` 的 SHA-256：

```json
{
  "version": "2.2.0",
  "app_sha256": "<app.py 的 64 位元 SHA-256 十六進位值>"
}
```

更新來源固定為 `lianghao02/PaperSwitch` 的 GitHub Release，**不會**直接覆寫 `main` 分支內容。私人儲存庫不適合此第一階段機制，因為應用程式不保存 GitHub Token。

---

## 🛡️ 安全防護與使用邊界

1. **本機隔離原則**：服務預設僅綁定本機迴路介面（`127.0.0.1:8080`），檔案均於本機記憶體與暫存目錄處理，不外傳任何雲端伺服器。
2. **Office COM 穩定性防禦**：所有 COM 操作皆由全域執行緒鎖 `COM_LOCK` 保護，並在 `finally` 區塊強制執行 `Close(False)`、`Quit()` 與 `pythoncom.CoUninitialize()`，防禦背景進程殘留。
3. **重要文件備份**：轉換產出均存放於 `converted/` 目錄，絕不覆寫或修改使用者原始檔案。

---

## 📚 專案文件導航

- 變更歷程：[CHANGELOG.md](CHANGELOG.md)
- 架構設計：[`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md)
- 需求規格：[`docs/spec.md`](docs/spec.md)
- 經驗與技術決策：[`docs/MEMORY.md`](docs/MEMORY.md)
- Agent 準則：[`AGENTS.md`](AGENTS.md)
