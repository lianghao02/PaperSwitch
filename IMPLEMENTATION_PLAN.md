# 實作計畫

## 目標與驗收條件

為本機 PaperSwitch 提供 GitHub Release 的版本檢查與一鍵更新 `app.py`，並建立免安裝可攜套件：使用者可看見目前與最新版本、更新摘要，並在更新前建立備份；下載或驗證失敗時保留可運作舊版。可攜版解壓縮後可透過 VBS 啟動器以獨立 Edge 視窗開啟。

## 不做範圍

- 不修改使用者的 `uploads/`、`converted/` 或 `.env`；封裝時一律建立空白工作資料夾。
- 不建立自製 EXE、不要求管理員權限，也不將端點加密的受保護 PDF 納入處理範圍。
- 不在程式內保存 GitHub Token，僅支援可公開讀取的 Release。

## 現況與限制

- 遠端儲存庫為 `lianghao02/PaperSwitch`；工作目錄已有使用者未提交異動，必須保留。
- 更新來源僅採 GitHub Release 的固定資產，不採用可變動的 `main` 分支。
- 完整性檢查使用 Release 中繼資料提供的 SHA-256；GitHub Release 的發布權限仍是更新信任邊界。

## 已確認決策

- 第一階段只更新 `app.py` 與本機 `version.json`。
- Release 必須附上 `app.py` 與 `version.json` 資產；`version.json` 需包含 `version`、`app_sha256`。
- 更新前建立 `app.py.bak`，成功後以原子替換更新檔案，並由目前 Python 程序重新啟動。
- 可攜套件以現有 `python_embed/` 為唯一執行環境；封裝前必須實際匯入 `PIL`、`pypdf`、`pymupdf`、`win32com.client`、`dotenv`，缺任一項即拒絕產出 ZIP。
- VBS 啟動器使用 `pythonw.exe` 隱藏終端視窗，輪詢 `/api/heartbeat` 後才以 Edge `--app` 模式開啟；`RUN.bat` 保留為可見終端的故障排除備援。

## 工作清單

- [x] 新增本機版本資訊與 Release 讀取、比對、下載、驗證及回復邏輯｜Python 語法與本機模擬 Release 驗證。
- [x] 新增更新檢查／執行 API 與暖色介面互動｜靜態 JavaScript 語法與 API 模擬驗證。
- [x] 新增版本檔與使用說明｜檢查版本格式及更新資產契約。
- [x] 執行既有可行檢查並記錄結果｜`py_compile`、JavaScript 語法、`git diff --check`。
- [x] 新增 VBS 無終端視窗啟動器與安全封裝腳本｜靜態檢查啟動器、封裝前依賴驗證。
- [ ] 上傳已驗證的 ZIP 至既有 v3.0.0 GitHub Release｜檢查 Release 資產名稱、大小與 SHA-256。

## 風險與因應

- GitHub 無法連線或無 Release：回傳可理解訊息，既有功能不中斷。
- Release 資產缺漏或雜湊不符：拒絕更新，保留原版與備份。
- 重啟期間瀏覽器連線中斷：前端收到成功後輪詢服務恢復並重新整理。
- 可攜環境少任一既有依賴：封裝腳本拒絕產出 ZIP，先補齊 `requirements.txt` 中既有套件再繼續。

## 驗證紀錄

- Python 3.13：`-m py_compile app.py` 通過。
- JavaScript：以 Node.js 載入內嵌 `<script>` 通過。
- 本機模擬 Release：驗證較新版本可檢出、`app.py.bak` 建立、原子替換成功；SHA-256 不符時拒絕覆寫。
- 暫時 HTTP 伺服器：`GET /api/update/check` 與 `POST /api/update/perform` 模擬回應通過。
- `git diff --check` 通過。
- 可攜環境檢查：已發現目前 `python_embed/` 缺少 `win32com.client`，尚不可發布可處理 Office 文件的可攜 ZIP。
- 可攜封裝：以 `scripts/build_portable.ps1` 產出 ZIP，並在解壓後直接以 ZIP 內 `python_embed` 通過 `app.py` 編譯與 `PIL`、`pypdf`、`pymupdf`、`win32com.client`、`dotenv` 匯入驗證；最終 ZIP SHA-256 為 `772650b57608b8a8112af31302b333a9bedd161b5f476838eec0d390bf89d44d`。

## 剩餘問題

- 尚未以實際可攜 ZIP 解壓縮後的 VBS 啟動流程進行端對端驗證；它會開啟 Edge 應用程式視窗，建議於目標電腦人工雙擊驗收。
