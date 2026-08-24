# 實作計畫

## 目標與驗收條件

為本機 PaperSwitch 提供 GitHub Release 的版本檢查與一鍵更新 `app.py`：使用者可看見目前與最新版本、更新摘要，並在更新前建立備份；下載或驗證失敗時保留可運作舊版。

## 不做範圍

- 不建立、發布或推送 GitHub Release。
- 不更新 `python_embed`、第三方套件、啟動器或使用者的 `uploads/`、`converted/`。
- 不在程式內保存 GitHub Token，僅支援可公開讀取的 Release。

## 現況與限制

- 遠端儲存庫為 `lianghao02/PaperSwitch`；工作目錄已有使用者未提交異動，必須保留。
- 更新來源僅採 GitHub Release 的固定資產，不採用可變動的 `main` 分支。
- 完整性檢查使用 Release 中繼資料提供的 SHA-256；GitHub Release 的發布權限仍是更新信任邊界。

## 已確認決策

- 第一階段只更新 `app.py` 與本機 `version.json`。
- Release 必須附上 `app.py` 與 `version.json` 資產；`version.json` 需包含 `version`、`app_sha256`。
- 更新前建立 `app.py.bak`，成功後以原子替換更新檔案，並由目前 Python 程序重新啟動。

## 工作清單

- [x] 新增本機版本資訊與 Release 讀取、比對、下載、驗證及回復邏輯｜Python 語法與本機模擬 Release 驗證。
- [x] 新增更新檢查／執行 API 與暖色介面互動｜靜態 JavaScript 語法與 API 模擬驗證。
- [x] 新增版本檔與使用說明｜檢查版本格式及更新資產契約。
- [x] 執行既有可行檢查並記錄結果｜`py_compile`、JavaScript 語法、`git diff --check`。

## 風險與因應

- GitHub 無法連線或無 Release：回傳可理解訊息，既有功能不中斷。
- Release 資產缺漏或雜湊不符：拒絕更新，保留原版與備份。
- 重啟期間瀏覽器連線中斷：前端收到成功後輪詢服務恢復並重新整理。

## 驗證紀錄

- Python 3.13：`-m py_compile app.py` 通過。
- JavaScript：以 Node.js 載入內嵌 `<script>` 通過。
- 本機模擬 Release：驗證較新版本可檢出、`app.py.bak` 建立、原子替換成功；SHA-256 不符時拒絕覆寫。
- 暫時 HTTP 伺服器：`GET /api/update/check` 與 `POST /api/update/perform` 模擬回應通過。
- `git diff --check` 通過。

## 剩餘問題

- 首個可公開讀取的 GitHub Release 尚未建立；功能完成後需由維護者依資產契約發布。
- 尚未以實際 GitHub Release 和實際程序重啟做端對端測試；需在首個 Release 發布後驗證。
