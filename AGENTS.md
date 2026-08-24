# AGENTS.md — Agent 角色規範與行為準則

本檔定義 `09_PaperSwitch` 專案專屬限制與開發準則；全域憲法由全域環境統一注入。

## 1. 專案核心架構邊界
- **單檔與內建模組優先**：核心前後端皆整合於 [`app.py`](app.py) 單檔中，未經明確架構重構需求，不得隨意拆散為多檔案增加部署複雜度。
- **Office COM 資源防禦**：調用 Word/Excel/PowerPoint COM 元件時，必須受全域 `COM_LOCK` 保護，且必須在 `finally` 區塊明確釋放物件並呼叫 `pythoncom.CoUninitialize()`。
- **DOM 拖曳安全禁令**：在前端 Drag & Drop（`dragstart`, `dragover`）過程中，**嚴禁呼叫任何破壞或重建 DOM 的函式**（如 `renderArrangerGrid()`），避免中斷瀏覽器原生 DragSession。
- **無損向量合成**：導出合成 PDF 時必須透過 `pypdf` 頁面矩陣抽取與旋轉，嚴禁以縮圖二次重壓為 PDF。

## 2. 視覺風格規範
- 介面必須遵循「溫暖手作插畫手帳風格 (Warm Cozy Craft Palette)」：
  - 底色：`#F7F4ED`（紙白）、卡片：`#FFFFFF`、邊框：`#E3DDD2`。
  - 主強調色：`#D97736`（陶土暖橘）、次強調色：`#5B8266`（森林苔綠）、文字：`#2D2825`（炭焙棕黑）。
  - 文案採用溫暖親切的生活感用語，避免生硬冷冽的工程字眼。

## 3. 文件閱讀路徑
進入本專案修改核心邏輯前，請依序參閱：
1. [`README.md`](README.md)：整體專案功能與使用邊界
2. [`docs/MEMORY.md`](docs/MEMORY.md)：歷史 Bug 坑洞與關鍵演算法決策
3. [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md)：系統架構與模組資料流
4. [`CHANGELOG.md`](CHANGELOG.md)：版本變更歷程
