# ============================================================
# CONFIG — 所有可變參數集中於此，嚴禁在下方程式碼中散落魔術數字
# ============================================================
import json
import os
import shutil
import subprocess
import sys
import webbrowser
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from PIL import Image
from dotenv import load_dotenv
from pypdf import PdfReader, PdfWriter

# 讀取 .env 設定檔
load_dotenv()

BASE_DIR = Path(__file__).resolve().parent

CONFIG = {
    "host": os.getenv("HOST", "0.0.0.0" if os.getenv("RENDER") or os.getenv("DOCKER") else "127.0.0.1"),
    "port": int(os.getenv("PORT", "8080")),
    "upload_dir": os.getenv("UPLOAD_DIR", str(BASE_DIR / "uploads")),
    "output_dir": os.getenv("OUTPUT_DIR", str(BASE_DIR / "converted")),
    "excel_fit_to_page": os.getenv("EXCEL_FIT_TO_PAGE", "True").lower() in ("true", "1"),
    "log_level": os.getenv("LOG_LEVEL", "INFO"),
}

if sys.platform == "win32":
    import pythoncom
    import win32com.client

# 確保專案內部的暫存與輸出資料夾存在
Path(CONFIG["upload_dir"]).mkdir(parents=True, exist_ok=True)
Path(CONFIG["output_dir"]).mkdir(parents=True, exist_ok=True)


# ============================================================
# 多引擎核心轉換器 (Word, Excel, PPT, 圖片, PDF)
# ============================================================
class DocumentConverter:
    """處理 Word, Excel, PowerPoint, 圖片轉 PDF 及 PDF 合併的核心處理器"""

    @staticmethod
    def convert_word(input_path: str, output_path: str) -> bool:
        """將 Word 檔案 (.doc, .docx) 轉為 PDF"""
        input_abs = str(Path(input_path).resolve())
        output_abs = str(Path(output_path).resolve())

        if sys.platform == "win32":
            word, doc = None, None
            try:
                pythoncom.CoInitialize()
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = False
                doc = word.Documents.Open(input_abs, ReadOnly=True)
                doc.SaveAs(output_abs, FileFormat=17)  # 17 = wdFormatPDF
                print(f"✅ [MS Word COM 轉換成功] -> {output_abs}")
                return True
            except Exception as e:
                print(f"❌ [MS Word COM 轉換失敗] {input_abs}: {e}")
                return False
            finally:
                if doc:
                    doc.Close(False)
                if word:
                    word.Quit()
                pythoncom.CoUninitialize()
        else:
            try:
                cmd = [
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(Path(output_abs).parent),
                    input_abs,
                ]
                subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                gen_pdf = Path(output_abs).parent / f"{Path(input_abs).stem}.pdf"
                if gen_pdf.exists() and str(gen_pdf) != output_abs:
                    gen_pdf.replace(output_abs)
                print(f"✅ [LibreOffice Word 轉換成功] -> {output_abs}")
                return True
            except Exception as e:
                print(f"❌ [LibreOffice Word 轉換失敗]: {e}")
                return False

    @staticmethod
    def convert_excel(input_path: str, output_path: str) -> list[str]:
        """將 Excel 檔案 (.xls, .xlsx) 的每個可見分頁分別匯出為獨立的 PDF 檔案"""
        input_abs = str(Path(input_path).resolve())
        output_base_path = Path(output_path).resolve()
        output_dir = output_base_path.parent
        stem = output_base_path.stem

        generated_pdfs = []

        if sys.platform == "win32":
            excel, wb = None, None
            try:
                pythoncom.CoInitialize()
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(input_abs, ReadOnly=True)

                visible_sheets = [ws for ws in wb.Worksheets if ws.Visible == -1]

                if len(visible_sheets) <= 1:
                    if visible_sheets:
                        ws = visible_sheets[0]
                        if CONFIG["excel_fit_to_page"]:
                            try:
                                ws.PageSetup.Zoom = False
                                ws.PageSetup.FitToPagesWide = 1
                                ws.PageSetup.FitToPagesTall = False
                            except Exception:
                                pass
                        ws.ExportAsFixedFormat(0, str(output_base_path))
                        generated_pdfs.append(str(output_base_path))
                else:
                    for ws in visible_sheets:
                        sheet_name = ws.Name
                        safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (" ", "_", "-")).strip()
                        sheet_pdf_path = str(output_dir / f"{stem}_{safe_sheet_name}.pdf")

                        if CONFIG["excel_fit_to_page"]:
                            try:
                                ws.PageSetup.Zoom = False
                                ws.PageSetup.FitToPagesWide = 1
                                ws.PageSetup.FitToPagesTall = False
                            except Exception:
                                pass

                        ws.ExportAsFixedFormat(0, sheet_pdf_path)
                        generated_pdfs.append(sheet_pdf_path)

                print(f"✅ [MS Excel 分頁獨立拆分成功] 共產出 {len(generated_pdfs)} 個 PDF -> {output_dir}")
                return generated_pdfs
            except Exception as e:
                print(f"❌ [MS Excel 分頁拆分轉檔失敗] {input_abs}: {e}")
                return []
            finally:
                if wb:
                    wb.Close(False)
                if excel:
                    excel.Quit()
                pythoncom.CoUninitialize()
        else:
            try:
                cmd = [
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(output_dir),
                    input_abs,
                ]
                subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                gen_pdf = output_dir / f"{stem}.pdf"
                if gen_pdf.exists():
                    generated_pdfs.append(str(gen_pdf))
                print(f"✅ [LibreOffice Excel 轉換成功] -> {output_dir}")
                return generated_pdfs
            except Exception as e:
                print(f"❌ [LibreOffice Excel 轉換失敗]: {e}")
                return []

    @staticmethod
    def convert_powerpoint(input_path: str, output_path: str) -> bool:
        """將 PowerPoint 檔案 (.ppt, .pptx) 轉為 PDF"""
        input_abs = str(Path(input_path).resolve())
        output_abs = str(Path(output_path).resolve())

        if sys.platform == "win32":
            ppt, presentation = None, None
            try:
                pythoncom.CoInitialize()
                ppt = win32com.client.DispatchEx("PowerPoint.Application")
                # PowerPoint COM 需要搭配 WithWindow=False 打開簡報
                presentation = ppt.Presentations.Open(input_abs, WithWindow=False)
                presentation.SaveAs(output_abs, 32)  # 32 = ppSaveAsPDF
                print(f"✅ [MS PowerPoint COM 轉換成功] -> {output_abs}")
                return True
            except Exception as e:
                print(f"❌ [MS PowerPoint COM 轉換失敗] {input_abs}: {e}")
                return False
            finally:
                if presentation:
                    presentation.Close()
                if ppt:
                    ppt.Quit()
                pythoncom.CoUninitialize()
        else:
            try:
                cmd = [
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(Path(output_abs).parent),
                    input_abs,
                ]
                subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                gen_pdf = Path(output_abs).parent / f"{Path(input_abs).stem}.pdf"
                if gen_pdf.exists() and str(gen_pdf) != output_abs:
                    gen_pdf.replace(output_abs)
                print(f"✅ [LibreOffice PowerPoint 轉換成功] -> {output_abs}")
                return True
            except Exception as e:
                print(f"❌ [LibreOffice PowerPoint 轉換失敗]: {e}")
                return False

    @staticmethod
    def convert_image(input_paths: list[str], output_path: str) -> bool:
        """將單張或多張圖片合成轉換為單一 PDF"""
        output_abs = str(Path(output_path).resolve())
        try:
            opened_images = []
            for p in input_paths:
                img = Image.open(p)
                if img.mode != "RGB":
                    img = img.convert("RGB")
                opened_images.append(img)

            if not opened_images:
                print("⚠️ 無有效圖片可轉換")
                return False

            first_img = opened_images[0]
            rest_imgs = opened_images[1:]
            first_img.save(
                output_abs,
                "PDF",
                save_all=True,
                append_images=rest_imgs,
                resolution=100.0,
            )
            print(f"✅ [圖片轉換成功] ({len(input_paths)} 張) -> {output_abs}")
            return True
        except Exception as e:
            print(f"❌ [圖片轉換失敗]: {e}")
            return False

    @staticmethod
    def merge_pdfs(pdf_paths: list[str], output_path: str) -> bool:
        """將多個 PDF 檔案合併為單一 PDF (防禦型強健實作)"""
        output_abs = str(Path(output_path).resolve())
        try:
            writer = PdfWriter()
            append_count = 0
            for p in pdf_paths:
                p_abs = str(Path(p).resolve())
                if not Path(p_abs).exists():
                    print(f"⚠️ 合併時跳過不存在之 PDF 檔: {p_abs}")
                    continue
                try:
                    reader = PdfReader(p_abs)
                    for page in reader.pages:
                        writer.add_page(page)
                    append_count += 1
                except Exception as page_err:
                    print(f"⚠️ 讀取/加入 PDF 檔頁面失敗 {p_abs}: {page_err}")

            if append_count == 0:
                print("❌ 無任何可用的 PDF 檔進行合併")
                return False

            with open(output_abs, "wb") as f:
                writer.write(f)
            writer.close()
            print(f"✅ [PDF 合併成功] ({append_count}/{len(pdf_paths)} 檔) -> {output_abs}")
            return True
        except Exception as e:
            print(f"❌ [PDF 合併失敗]: {e}")
            return False


# ============================================================
# Web UI 前端 HTML 樣式 (含手動清理暫存檔與 PPT/PPTX 支援)
# ============================================================
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PaperSwitch - 萬能文件轉 PDF 處理器</title>
    <!-- Google Fonts -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Outfit:wght@600;700;800&display=swap" rel="stylesheet">
    
    <style>
        :root {
            --font-heading: 'Outfit', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            --font-body: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            --bg-color: #0b1329;
            --card-bg: rgba(30, 41, 59, 0.8);
            --panel-bg: rgba(15, 23, 42, 0.55);
            --border-color: rgba(255, 255, 255, 0.12);
            --accent-color: #38bdf8;
            --accent-hover: #0284c7;
            --secondary-bg: rgba(255, 255, 255, 0.08);
            --secondary-hover: rgba(255, 255, 255, 0.16);
            --text-h1: #ffffff;
            --text-h2: #f8fafc;
            --text-h3: #e2e8f0;
            --text-sub: #94a3b8;
            --success-color: #4ade80;
            --warning-color: #fbbf24;
            --error-color: #f87171;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        
        html, body {
            height: 100vh;
            overflow: hidden;
            background-color: var(--bg-color);
            color: var(--text-h2);
            font-family: var(--font-body);
            -webkit-font-smoothing: antialiased;
        }
        
        body {
            display: flex;
            flex-direction: column;
            padding: 20px 24px;
            box-sizing: border-box;
        }
        
        header {
            text-align: center;
            margin-bottom: 16px;
            flex-shrink: 0;
        }
        header h1 {
            font-family: var(--font-heading);
            font-size: 1.85rem;
            font-weight: 800;
            letter-spacing: -0.025em;
            background: linear-gradient(135deg, #ffffff 30%, #38bdf8 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        header p {
            font-family: var(--font-body);
            color: var(--text-sub);
            font-size: 0.88rem;
            margin-top: 4px;
        }

        .dashboard {
            display: grid;
            grid-template-columns: 1fr 1.2fr 1fr;
            gap: 20px;
            width: 100%;
            max-width: 1280px;
            margin: 0 auto;
            flex: 1;
            min-height: 0;
        }
        @media (max-width: 900px) {
            html, body { overflow: auto; height: auto; }
            .dashboard { grid-template-columns: 1fr; }
        }

        .panel {
            background: var(--card-bg);
            backdrop-filter: blur(16px);
            border: 1px solid var(--border-color);
            border-radius: 16px;
            padding: 20px;
            display: flex;
            flex-direction: column;
            box-shadow: 0 20px 40px -15px rgba(0, 0, 0, 0.5);
            height: 100%;
            min-height: 0;
        }
        
        .panel-title {
            font-family: var(--font-heading);
            font-size: 1.08rem;
            font-weight: 700;
            color: var(--text-h2);
            margin-bottom: 14px;
            display: flex;
            align-items: center;
            gap: 8px;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 10px;
            flex-shrink: 0;
        }

        .drop-zone {
            border: 2px dashed var(--border-color);
            border-radius: 12px;
            padding: 24px 16px;
            text-align: center;
            cursor: pointer;
            transition: all 0.25s ease;
            background: var(--panel-bg);
            flex: 1;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            min-height: 0;
        }
        .drop-zone:hover, .drop-zone.dragover { border-color: var(--accent-color); background: rgba(56, 189, 248, 0.08); transform: translateY(-2px); }
        .drop-icon { font-size: 52px; margin-bottom: 14px; }
        .drop-text { font-family: var(--font-heading); font-size: 1.05rem; font-weight: 700; color: var(--text-h2); }
        .drop-hint { font-family: var(--font-body); font-size: 0.82rem; color: var(--text-sub); margin-top: 10px; line-height: 1.55; }

        .progress-box { margin-bottom: 14px; flex-shrink: 0; }
        .progress-info { font-family: var(--font-body); display: flex; justify-content: space-between; font-size: 0.85rem; font-weight: 600; margin-bottom: 6px; color: var(--text-sub); }
        .progress-bar-bg { width: 100%; height: 8px; background: rgba(255,255,255,0.1); border-radius: 4px; overflow: hidden; }
        .progress-bar-fill { height: 100%; width: 0%; background: var(--accent-color); transition: width 0.3s ease; }
        
        .queue-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; flex-shrink: 0; }
        .queue-count { font-size: 0.85rem; font-weight: 600; color: var(--text-sub); }
        .clear-btn { font-size: 0.8rem; font-weight: 600; color: var(--error-color); background: none; border: none; cursor: pointer; text-decoration: underline; }
        
        .file-queue {
            flex: 1;
            overflow-y: auto;
            display: flex;
            flex-direction: column;
            gap: 8px;
            padding-right: 4px;
            min-height: 0;
        }
        .file-card { display: flex; justify-content: space-between; align-items: center; padding: 10px 12px; background: var(--panel-bg); border: 1px solid rgba(255,255,255,0.06); border-radius: 8px; font-size: 0.85rem; }
        .file-info { display: flex; flex-direction: column; gap: 2px; overflow: hidden; }
        .file-name { font-weight: 600; color: var(--text-h3); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 170px; }
        .file-meta { font-size: 0.75rem; color: var(--text-sub); }
        .file-status { font-size: 0.75rem; font-weight: 700; padding: 3px 8px; border-radius: 4px; }
        .status-ready { background: rgba(255,255,255,0.1); color: var(--text-sub); }
        .status-processing { background: rgba(251, 191, 36, 0.2); color: var(--warning-color); }
        .status-done { background: rgba(74, 222, 128, 0.2); color: var(--success-color); }
        .status-fail { background: rgba(248, 113, 113, 0.2); color: var(--error-color); }

        .option-group { margin-bottom: 16px; display: flex; flex-direction: column; gap: 10px; flex-shrink: 0; }
        .option-label { font-family: var(--font-heading); font-size: 0.92rem; font-weight: 700; color: var(--text-h3); }
        .radio-card { display: flex; align-items: flex-start; gap: 10px; padding: 14px; background: var(--panel-bg); border: 1px solid var(--border-color); border-radius: 10px; cursor: pointer; transition: all 0.2s ease; }
        .radio-card:hover, .radio-card.active { border-color: var(--accent-color); background: rgba(56, 189, 248, 0.06); }
        .radio-card input { margin-top: 3px; accent-color: var(--accent-color); }
        .radio-text { display: flex; flex-direction: column; gap: 4px; }
        .radio-title { font-family: var(--font-heading); font-size: 0.92rem; font-weight: 700; color: var(--text-h2); }
        .radio-desc { font-size: 0.78rem; color: var(--text-sub); line-height: 1.35; }
        
        .input-group { display: flex; flex-direction: column; gap: 6px; margin-bottom: 16px; flex-shrink: 0; }
        .input-field { width: 100%; padding: 10px 12px; background: rgba(15, 23, 42, 0.6); border: 1px solid var(--border-color); border-radius: 8px; color: var(--text-h2); font-size: 0.85rem; outline: none; }
        .input-field:focus { border-color: var(--accent-color); }

        .btn-group { display: flex; flex-direction: column; gap: 10px; margin-top: auto; flex-shrink: 0; }
        .btn { font-family: var(--font-heading); width: 100%; padding: 14px; background: var(--accent-color); color: #0f172a; font-weight: 800; border: none; border-radius: 10px; cursor: pointer; transition: all 0.2s ease; font-size: 0.98rem; text-align: center; }
        .btn:hover { background: var(--accent-hover); color: #ffffff; }
        .btn-secondary { background: var(--secondary-bg); color: var(--text-h2); border: 1px solid var(--border-color); }
        .btn-secondary:hover { background: var(--secondary-hover); }
        .btn-danger-outline { background: rgba(248, 113, 113, 0.1); color: var(--error-color); border: 1px solid rgba(248, 113, 113, 0.3); }
        .btn-danger-outline:hover { background: rgba(248, 113, 113, 0.2); }

        #globalStatus { margin-top: 12px; text-align: center; font-size: 0.85rem; font-weight: 600; min-height: 20px; flex-shrink: 0; }
    </style>
</head>
<body>
    <header>
        <h1>📄 PaperSwitch</h1>
        <p>萬能 Word / Excel / PowerPoint / 圖片 轉 PDF 與 PDF 合併引擎</p>
    </header>

    <div class="dashboard">
        <!-- 區塊 1: 拖曳區 -->
        <div class="panel">
            <h2 class="panel-title">📥 1. 檔案拖曳區</h2>
            <div class="drop-zone" id="dropZone" onclick="document.getElementById('fileInput').click()">
                <div class="drop-icon">✨</div>
                <div class="drop-text">將檔案拖拽至此處</div>
                <div class="drop-hint">或點擊此區域選擇檔案<br><br>支援格式：<br>.docx, .doc, .xlsx, .xls<br>.pptx, .ppt<br>.jpg, .png, .bmp, .webp, .pdf</div>
                <input type="file" id="fileInput" multiple style="display: none;" onchange="handleFiles(this.files)">
            </div>
        </div>

        <!-- 區塊 2: 佇列與目前進度區 -->
        <div class="panel">
            <h2 class="panel-title">📋 2. 檔案佇列與進度</h2>
            
            <div class="progress-box">
                <div class="progress-info">
                    <span id="progressStatusText">佇列就緒</span>
                    <span id="progressPercentText">0%</span>
                </div>
                <div class="progress-bar-bg">
                    <div class="progress-bar-fill" id="progressBar"></div>
                </div>
            </div>

            <div class="queue-header">
                <span class="queue-count" id="queueCountText">共 0 個檔案</span>
                <button class="clear-btn" onclick="clearQueue()">清空佇列</button>
            </div>

            <div class="file-queue" id="fileQueue">
                <div style="text-align: center; color: var(--text-sub); font-size: 0.85rem; margin: auto 0;">尚無待處理檔案</div>
            </div>
        </div>

        <!-- 區塊 3: 功能設定區 -->
        <div class="panel">
            <h2 class="panel-title">⚙️ 3. 轉換功能設定</h2>
            
            <div class="option-group">
                <h3 class="option-label">輸出模式設定：</h3>
                
                <label class="radio-card active" id="cardSingle" onclick="setMode('single')">
                    <input type="radio" name="convertMode" value="single" checked>
                    <div class="radio-text">
                        <span class="radio-title">📄 獨立單檔模式</span>
                        <span class="radio-desc">每個檔案個別轉換為對應同名 PDF (Excel 多分頁自動拆分獨立產出)</span>
                    </div>
                </label>

                <label class="radio-card" id="cardMerge" onclick="setMode('merge')">
                    <input type="radio" name="convertMode" value="merge">
                    <div class="radio-text">
                        <span class="radio-title">📚 合併 PDF 模式</span>
                        <span class="radio-desc">將佇列中所有檔案按順序合併為單一 PDF</span>
                    </div>
                </label>
            </div>

            <div class="input-group" id="mergedFilenameBox" style="display: none;">
                <h4 class="option-label" style="font-size: 0.8rem;">合併 PDF 檔名：</h4>
                <input type="text" id="mergedFilename" class="input-field" value="combined_output.pdf" placeholder="例如：combined_output.pdf">
            </div>

            <div class="btn-group">
                <button class="btn" onclick="uploadAndConvert()">🚀 開始轉換</button>
                <button class="btn btn-secondary" id="openFolderBtn" onclick="openOutputFolder()" style="display: none;">📂 開啟 PDF 輸出資料夾</button>
                <button class="btn btn-secondary btn-danger-outline" onclick="clearStorage()">🧹 清空歷史暫存檔 (uploads/與converted/)</button>
            </div>

            <div id="globalStatus"></div>
        </div>
    </div>

    <script>
        let selectedFiles = [];
        let currentMode = 'single';

        const dropZone = document.getElementById('dropZone');

        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eName => {
            dropZone.addEventListener(eName, e => { e.preventDefault(); e.stopPropagation(); });
        });
        dropZone.addEventListener('dragover', () => dropZone.classList.add('dragover'));
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
        dropZone.addEventListener('drop', e => {
            dropZone.classList.remove('dragover');
            handleFiles(e.dataTransfer.files);
        });

        function handleFiles(files) {
            Array.from(files).forEach(f => {
                if (!selectedFiles.some(existing => existing.name === f.name && existing.size === f.size)) {
                    selectedFiles.push({ file: f, status: 'ready', message: '' });
                }
            });
            renderQueue();
        }

        function clearQueue() {
            selectedFiles = [];
            updateProgress(0, '佇列就緒');
            renderQueue();
        }

        function removeFile(index) {
            selectedFiles.splice(index, 1);
            renderQueue();
        }

        function setMode(mode) {
            currentMode = mode;
            document.getElementById('cardSingle').classList.toggle('active', mode === 'single');
            document.getElementById('cardMerge').classList.toggle('active', mode === 'merge');
            document.querySelector('input[value="single"]').checked = (mode === 'single');
            document.querySelector('input[value="merge"]').checked = (mode === 'merge');
            document.getElementById('mergedFilenameBox').style.display = (mode === 'merge') ? 'flex' : 'none';
        }

        function renderQueue() {
            const queueEl = document.getElementById('fileQueue');
            const countEl = document.getElementById('queueCountText');
            countEl.innerText = `共 ${selectedFiles.length} 個檔案`;

            if (selectedFiles.length === 0) {
                queueEl.innerHTML = '<div style="text-align: center; color: var(--text-sub); font-size: 0.85rem; margin: auto 0;">尚無待處理檔案</div>';
                return;
            }

            queueEl.innerHTML = selectedFiles.map((item, idx) => {
                let badgeClass = 'status-ready';
                let statusLabel = '待轉換';
                if (item.status === 'processing') { badgeClass = 'status-processing'; statusLabel = '轉檔中'; }
                else if (item.status === 'done') { badgeClass = 'status-done'; statusLabel = '已完成'; }
                else if (item.status === 'fail') { badgeClass = 'status-fail'; statusLabel = '失敗'; }

                return `
                    <div class="file-card">
                        <div class="file-info">
                            <span class="file-name" title="${item.file.name}">📄 ${item.file.name}</span>
                            <span class="file-meta">${(item.file.size / 1024).toFixed(1)} KB</span>
                        </div>
                        <div style="display: flex; align-items: center; gap: 8px;">
                            <span class="file-status ${badgeClass}">${statusLabel}</span>
                            <button onclick="removeFile(${idx})" style="background:none; border:none; color:var(--text-sub); cursor:pointer; font-size:0.9rem;">✕</button>
                        </div>
                    </div>
                `;
            }).join('');
        }

        function updateProgress(percent, statusText) {
            document.getElementById('progressBar').style.width = percent + '%';
            document.getElementById('progressPercentText').innerText = percent + '%';
            document.getElementById('progressStatusText').innerText = statusText;
        }

        async function uploadAndConvert() {
            if (selectedFiles.length === 0) { alert('請先新增檔案至佇列！'); return; }

            const globalStatus = document.getElementById('globalStatus');
            const openFolderBtn = document.getElementById('openFolderBtn');

            globalStatus.className = '';
            globalStatus.innerText = '⏳ 正在上傳與處理中...';
            openFolderBtn.style.display = 'none';

            selectedFiles.forEach(item => item.status = 'processing');
            renderQueue();
            updateProgress(30, '處理中...');

            const formData = new FormData();
            selectedFiles.forEach(item => formData.append('files', item.file));
            formData.append('mode', currentMode);

            let mergedName = document.getElementById('mergedFilename').value.trim();
            if (!mergedName.endsWith('.pdf')) mergedName += '.pdf';
            formData.append('merged_filename', mergedName);

            try {
                updateProgress(60, '轉檔渲染中...');
                const resp = await fetch('/api/convert', { method: 'POST', body: formData });
                const res = await resp.json();

                if (res.success) {
                    selectedFiles.forEach(item => item.status = 'done');
                    updateProgress(100, '轉換成功');
                    globalStatus.className = 'status-success';
                    globalStatus.innerText = (currentMode === 'merge') ? `✅ 已成功合併為 ${mergedName}` : '✅ 佇列檔案全部轉換完成！';
                    openFolderBtn.style.display = 'block';
                } else {
                    selectedFiles.forEach(item => item.status = 'fail');
                    updateProgress(100, '轉檔失敗');
                    globalStatus.className = 'status-error';
                    globalStatus.innerText = '❌ 轉檔失敗：' + (res.message || '未知錯誤');
                }
            } catch (err) {
                selectedFiles.forEach(item => item.status = 'fail');
                updateProgress(100, '網路異常');
                globalStatus.className = 'status-error';
                globalStatus.innerText = '❌ 網路連線或伺服器異常';
            }
            renderQueue();
        }

        async function openOutputFolder() {
            try {
                await fetch('/api/open-folder', { method: 'POST' });
            } catch (err) {
                alert('無法開啟本機資料夾');
            }
        }

        async function clearStorage() {
            if (confirm('確定要清空 uploads/ 與 converted/ 資料夾下的所有歷史暫存檔嗎？')) {
                try {
                    const resp = await fetch('/api/clear-storage', { method: 'POST' });
                    const res = await resp.json();
                    if (res.success) {
                        alert(`✅ 清理完成！已清除 ${res.count} 個歷史暫存檔案。`);
                    } else {
                        alert('❌ 清理失敗: ' + res.message);
                    }
                } catch (err) {
                    alert('❌ 網路請求失敗');
                }
            }
        }
    </script>
</body>
</html>
"""


# ============================================================
# HTTP 請求監聽處理器
# ============================================================
class WebAppHandler(SimpleHTTPRequestHandler):

    def do_GET(self):
        if self.path in ["/", "/index.html"]:
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(HTML_TEMPLATE.encode("utf-8"))
        elif self.path == "/favicon.ico":
            self.send_response(204)  # No Content
            self.end_headers()
        else:
            self.send_error(404, "Page Not Found")

    def do_POST(self):
        try:
            if self.path == "/api/convert":
                saved_paths, form_data = self._parse_multipart()

                if not saved_paths:
                    self._send_json({"success": False, "message": "未接收到有效檔案或表單格式錯誤"})
                    return

                mode = form_data.get("mode", "single")
                merged_filename = form_data.get("merged_filename", "combined_output.pdf")
                if not merged_filename.endswith(".pdf"):
                    merged_filename += ".pdf"

                converter = DocumentConverter()
                generated_pdfs = []
                overall_success = True

                for p in saved_paths:
                    ext = Path(p).suffix.lower()
                    stem = Path(p).stem
                    temp_pdf = Path(CONFIG["output_dir"]) / f"temp_{stem}.pdf"
                    final_pdf = Path(CONFIG["output_dir"]) / f"{stem}.pdf"

                    # 單檔模式下若目標 PDF 已存在，自動加上計數器 avoid 覆蓋
                    if mode == "single":
                        counter = 1
                        while final_pdf.exists():
                            final_pdf = Path(CONFIG["output_dir"]) / f"{stem}_{counter}.pdf"
                            counter += 1

                    target_path = str(temp_pdf) if mode == "merge" else str(final_pdf)

                    if ext in [".doc", ".docx"]:
                        res = converter.convert_word(p, target_path)
                        if res:
                            generated_pdfs.append(target_path)
                        else:
                            overall_success = False
                    elif ext in [".xls", ".xlsx"]:
                        excel_pdfs = converter.convert_excel(p, target_path)
                        if excel_pdfs:
                            generated_pdfs.extend(excel_pdfs)
                        else:
                            overall_success = False
                    elif ext in [".ppt", ".pptx"]:
                        res = converter.convert_powerpoint(p, target_path)
                        if res:
                            generated_pdfs.append(target_path)
                        else:
                            overall_success = False
                    elif ext in [".png", ".jpg", ".jpeg", ".bmp", ".webp"]:
                        res = converter.convert_image([p], target_path)
                        if res:
                            generated_pdfs.append(target_path)
                        else:
                            overall_success = False
                    elif ext == ".pdf":
                        # PDF 檔處理：合併模式時使用上傳暫存檔，單檔模式時複製至 output_dir
                        if mode == "merge":
                            generated_pdfs.append(p)
                        else:
                            try:
                                shutil.copy2(p, target_path)
                                generated_pdfs.append(target_path)
                            except Exception as copy_err:
                                print(f"❌ PDF 複製失敗: {copy_err}")
                                overall_success = False
                    else:
                        overall_success = False

                if mode == "merge" and generated_pdfs:
                    merged_output_path = str(Path(CONFIG["output_dir"]) / merged_filename)
                    merge_res = converter.merge_pdfs(generated_pdfs, merged_output_path)

                    for tmp_pdf in generated_pdfs:
                        if "temp_" in Path(tmp_pdf).name:
                            try:
                                os.remove(tmp_pdf)
                            except Exception:
                                pass

                    if not merge_res:
                        overall_success = False

                self._send_json(
                    {
                        "success": overall_success,
                        "message": "轉換完成" if overall_success else "部分或全部檔案轉檔失敗",
                    }
                )

            elif self.path == "/api/clear-storage":
                cleaned_count = self._clear_storage()
                self._send_json({"success": True, "count": cleaned_count, "message": "歷史暫存檔已清空"})

            elif self.path == "/api/open-folder":
                self._open_converted_folder()
                self._send_json({"success": True, "message": "已開啟輸出資料夾"})
            else:
                self.send_error(404, "API Endpoint Not Found")
        except Exception as e:
            print(f"❌ [do_POST 伺服器異常] {e}")
            try:
                self._send_json({"success": False, "message": f"伺服器處理失敗: {str(e)}"})
            except Exception:
                pass

    def _clear_storage(self) -> int:
        """手動清空 uploads/ 與 converted/ 目錄下的所有檔案"""
        cleaned_count = 0
        for folder in [CONFIG["upload_dir"], CONFIG["output_dir"]]:
            folder_path = Path(folder)
            if folder_path.exists():
                for item in folder_path.glob("*"):
                    try:
                        if item.is_file():
                            item.unlink()
                            cleaned_count += 1
                        elif item.is_dir():
                            shutil.rmtree(item)
                            cleaned_count += 1
                    except Exception as e:
                        print(f"⚠️ 清理舊暫存失敗 {item}: {e}")
        print(f"🧹 [手動清理暫存檔完成] 共刪除 {cleaned_count} 個項目")
        return cleaned_count

    def _open_converted_folder(self):
        """開啟本機的 PDF 輸出資料夾 (converted)"""
        output_abs = str(Path(CONFIG["output_dir"]).resolve())
        try:
            if sys.platform == "win32":
                os.startfile(output_abs)
            elif sys.platform == "darwin":
                subprocess.run(["open", output_abs])
            else:
                subprocess.run(["xdg-open", output_abs])
            print(f"📂 [開啟資料夾] -> {output_abs}")
        except Exception as e:
            print(f"❌ 開啟資料夾失敗: {e}")

    def _parse_multipart(self) -> tuple[list[str], dict[str, str]]:
        """100% 二進位安全 (Binary-Safe) 原生 multipart/form-data 解析器"""
        content_type = self.headers.get("Content-Type", "")
        content_length = int(self.headers.get("Content-Length", 0))

        if content_length <= 0 or "multipart/form-data" not in content_type:
            return [], {}

        boundary = None
        for part in content_type.split(";"):
            part = part.strip()
            if part.startswith("boundary="):
                boundary = part.split("=", 1)[1].strip('"').strip("'")
                break

        if not boundary:
            return [], {}

        body = self.rfile.read(content_length)

        # RFC 7578 標準分割符號為 \r\n--boundary
        delimiter = f"\r\n--{boundary}".encode("utf-8")
        first_boundary = f"--{boundary}\r\n".encode("utf-8")

        # 去除首個 boundary 前導字元
        if body.startswith(first_boundary):
            body = body[len(first_boundary):]
        elif body.startswith(f"--{boundary}".encode("utf-8")):
            body = body[len(f"--{boundary}".encode("utf-8")):]

        parts = body.split(delimiter)

        saved_files = []
        form_fields = {}

        for part in parts:
            if not part or part.startswith(b"--") or part == b"\r\n":
                continue

            if part.startswith(b"\r\n"):
                part = part[2:]

            if b"\r\n\r\n" not in part:
                continue

            headers_blob, data = part.split(b"\r\n\r\n", 1)

            # 清除末端 multipart 結尾符號 (--\r\n)
            if data.endswith(b"--\r\n"):
                data = data[:-4]
            elif data.endswith(b"--"):
                data = data[:-2]

            headers_str = headers_blob.decode("utf-8", errors="ignore")
            filename = None
            field_name = None

            for line in headers_str.split("\r\n"):
                if "Content-Disposition:" in line:
                    for param in line.split(";"):
                        param = param.strip()
                        if param.startswith('name="') or param.startswith("name="):
                            field_name = param.split("=", 1)[1].strip('"').strip("'")
                        if param.startswith('filename="') or param.startswith("filename="):
                            filename = param.split("=", 1)[1].strip('"').strip("'")

            if filename and data:
                safe_filename = Path(filename).name
                save_path = Path(CONFIG["upload_dir"]) / safe_filename

                # 同名檔案自動補上計數器後綴，防範複數同名檔上傳被覆蓋
                counter = 1
                stem = save_path.stem
                suffix = save_path.suffix
                while save_path.exists():
                    save_path = Path(CONFIG["upload_dir"]) / f"{stem}_{counter}{suffix}"
                    counter += 1

                with open(save_path, "wb") as f:
                    f.write(data)
                saved_files.append(str(save_path))
            elif field_name and not filename:
                form_fields[field_name] = data.decode("utf-8", errors="ignore")

        return saved_files, form_fields

    def _send_json(self, data: dict):
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.end_headers()
        self.wfile.write(json.dumps(data).encode("utf-8"))


# ============================================================
# 主入口程式 (升級 ThreadingHTTPServer 多執行緒併發處理)
# ============================================================
def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    server_address = (CONFIG["host"], CONFIG["port"])
    httpd = ThreadingHTTPServer(server_address, WebAppHandler)
    url = f"http://{CONFIG['host']}:{CONFIG['port']}"
    print(f"[PaperSwitch] 多執行緒轉檔伺服器運行中: {url}")

    # 自動開啟預設瀏覽器
    webbrowser.open(url)

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[INFO] 伺服器已安全停止")


if __name__ == "__main__":
    main()
