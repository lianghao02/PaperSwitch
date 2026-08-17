# PaperSwitch 文件轉 PDF 處理器 v1.1.1

[![Version](https://img.shields.io/badge/version-v1.1.1-blue.svg)](CHANGELOG.md)
[![Python](https://img.shields.io/badge/Python-3.13-blue.svg)](https://www.python.org/)
[![Constitution](https://img.shields.io/badge/Constitution-v8.0-purple.svg)](https://github.com/lianghao02/home)

PaperSwitch 是 Windows 本機文件轉 PDF 與 PDF 合併工具。後端使用 Python，Office 文件轉換依賴電腦上已安裝的 Microsoft Office COM 元件；圖片處理使用 Pillow，PDF 合併使用 pypdf。

## v1.1.1 更新重點

- 為 Office COM 操作加入序列化鎖定與確實關閉流程，降低多工作同時轉換的衝突。
- 修正 multipart 二進位上傳資料可能被截斷的問題。
- 強化 PDF 寫出及失敗時的資源釋放。

## 支援格式

- Word：`.docx`、`.doc`
- Excel：`.xlsx`、`.xls`
- PowerPoint：`.pptx`、`.ppt`
- 圖片：`.png`、`.jpg`、`.jpeg`、`.bmp`、`.webp`
- PDF：直接加入合併佇列

## 環境與啟動

- Windows 10 或更新版本
- Python 3.13
- 若要轉換 Office 文件，需安裝相容版本的 Microsoft Office

雙擊 [`start.bat`](start.bat)，或手動執行：

```powershell
python -m pip install -r requirements.txt
python app.py
```

預設服務位置為 `http://127.0.0.1:8080`。

## 使用限制與安全

- 服務預設只應綁定本機介面；若自行開放到區域網路，需另外加入存取控制。
- Office COM 轉換結果會受本機 Office 版本、字型、巨集及文件損毀情況影響。
- 轉換完成後應抽查頁數、版面及字型；重要原始檔請保留備份。
- 暫存清理只應針對本工具建立的工作目錄。

架構、設計與開發文件集中於 [`docs/`](docs/)；詳細異動請參閱 [CHANGELOG.md](CHANGELOG.md)。
