# MEMORY.md — 經驗持久化與 Bug 紀錄

## 📌 持久化經驗與 Bug 坑洞

### 1. HTML5 DragSession 中斷坑 (DOM Destruction on DragStart)
- **問題現象**：在點擊選取多張卡片後，滑鼠第一次拖曳卡片時拖曳動作立刻中斷，無法順利放置。
- **根本原因**：若在 `dragstart` 事件中呼叫了 `renderArrangerGrid()` 或重建 DOM，瀏覽器會銷毀目前滑鼠正抓取的 DOM 元素，導致原生 DragSession 被強制中止。
- **處方規範**：在拖曳開始（`dragstart`）與拖曳過程（`dragover`）中，**嚴禁呼叫任何重繪 DOM 的函式**。僅能透過直接修改現有元素的 `classList`（如 `.dragging`, `.drag-lead`, `.dragging-stacked`）與更新全域索引陣列；直到 `drop` 或 `dragend` 結束後，才進行單次完整的 `renderArrangerGrid()`。

### 2. 多選打包拖曳抽取演算法 (Multi-page Chunk Relocation)
- **演算法機制**：
  1. 使用者可透過點擊、`Ctrl + A` 複選多頁紙張。
  2. 拖曳選取的任意卡片時，系統收集所有選取索引 `draggedPageIndices = [...]`。
  3. 放置時：先將選取的頁面物件整體打包抽出，其餘剩餘頁面保持原相對位置，再於目標插入位置整批插入，並精確重建新的選取狀態與索引映射。

### 3. 鍵盤換位導航引擎 (Keyboard Reordering Engine)
- **實作細節**：
  - `Ctrl + ←` / `Ctrl + →`：支援單選與多選頁面群組移動。向前移動（←）時需由左至右（最小索引開始）依序交換；向後移動（→）時需由右至左（最大索引開始）依序交換，避免在陣列操作時發生索引重複覆寫衝突。
  - `Home` / `End`：整包移至第 1 頁或最後一頁。
  - 焦點切換（`←/→`）：透過 `card.scrollIntoView({ behavior: 'smooth', block: 'nearest' })` 保持視覺追蹤。

### 4. 邊緣平滑自動滾動 (Smooth Physics Auto-scroll)
- **實作細節**：使用 `requestAnimationFrame` 迴圈驅動，依據滑鼠與畫布頂部/底部的距離計算滾動速度 `velocity = (threshold - dist) * 0.25`，解決拖曳長篇 PDF 頁面時無法向下滾動的體驗痛點。

### 5. 向量無損合成與旋轉 (Zero-loss Vector PDF Synthesis)
- **架構設計**：
  - 縮圖渲染：利用 PyMuPDF (`fitz`) 以 180dpi 快速光柵化產生預覽縮圖。
  - 導出合成：完全不使用縮圖重壓，而是採用 `pypdf` (`PdfReader` / `PdfWriter`) 抽取原始 PDF Page 物件，並套用 `page.rotate(deg)`，保持 100% 原始向量清晰度、文字可選取性與極小檔案體積。

### 6. 溫暖手作插畫風格規範 (Warm Cozy Craft Palette)
- **視覺規範**：
  - 底色：水彩紙米白 `#F7F4ED`（搭配微點陣紙紋）
  - 面板卡片：純白 `#FFFFFF` + 米黃柔邊 `#EFEAE1`
  - 主強調色：陶土暖橘 `#D97736`（懸停 `#C46424`）
  - 次強調色：自然苔綠 `#5B8266`
  - 文字色調：深炭焙棕黑 `#2D2825`、鉛筆灰褐 `#736B63`
  - 語彙：採用生活感手作詞彙（投遞區、待整理清單、裝訂成冊、紙張排版工坊、工坊製作筆記）。

### 7. Office COM 穩定性防禦
- **執行緒隔離與資源釋放**：
  - 跨執行緒調用必須包裹 `pythoncom.CoInitialize()` 與 `CoUninitialize()`。
  - 全域 `COM_LOCK` 序列化調用，並在 `finally` 確實 `Close(False)` 與 `Quit()`。
  - Excel 寬度適應：`ws.PageSetup.FitToPagesWide = 1`。

### 8. Office IGEF 中介狀態與非同步安全保留機制 (Office Intermediate IGEF State)
- **問題本質**：Windows 下 MS Office 執行 `ExportAsFixedFormat` 時，輸出檔案初期可能短暫呈現 `IGEF\x02` 中介標記（常見於 MIP/AIP 敏感度標籤或背景列印 flush），之後才轉為標準 `%PDF-`。
- **處方規範**：
  1. `_check_pdf_header_and_size()` 辨識 `IGEF` 並在 `_wait_for_pdf_ready()` 期間持續輪詢直到寫入完成。
  2. **嚴禁在逾時或拆頁/轉圖失敗時呼叫 `unlink()` 刪除中介 PDF**，必須將檔案完整保留於 `converted/`，供使用者或非同步流程稍後重試。

---

## 📦 外部依賴追蹤
| 依賴套件 | 版本要求 | 職責用途 |
|---|---|---|
| `pywin32` | >=306 | Windows MS Office COM Automation 引擎 |
| `pypdf` | >=3.0.0 | PDF 頁面無損抽取、矩陣旋轉與向量合併 |
| `PyMuPDF` (fitz) | >=1.23.0 | PDF 頁面超高清縮圖即時渲染 |
| `Pillow` | >=10.0.0 | 圖片讀取、格式校正與無損 PDF 封裝 |
| `python-dotenv` | >=1.0.0 | `.env` 本機設定載入 |

---

## 📅 版本演進與里程碑
- **2026-08-24 (v3.0.0)**：加入 GitHub Release 檢查與受 SHA-256、備份、語法驗證保護的一鍵 `app.py` 更新；Office IGEF 暫存檔於無法處理時保留並提供解密再轉檔提示。
- **2026-08-24 (v2.1.0)**：全面重構為「溫暖手作插畫工坊 (Warm Cozy Craft)」風格，告別冷冽暗黑科技感。
- **2026-08-23 (v2.0.0)**：發布 10/10 完美里程碑（多選打包拖曳、方向鍵極速換位、PyMuPDF 縮圖、平滑邊緣自動滾動）。
- **2026-08-04 (v1.0.0)**：專案初始建立，支援 Office COM 與圖片轉 PDF。
