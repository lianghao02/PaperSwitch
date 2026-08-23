# ============================================================
# CONFIG — 所有可變參數集中於此，嚴禁在下方程式碼中散落魔術數字
# ============================================================
import json
import hmac
import os
import shutil
import subprocess
import sys
import threading
import time
import uuid
import urllib.parse
import webbrowser
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

# 守護 pythonw.exe (無終端模式) 下 sys.stdout 與 sys.stderr 為 None 的例外問題
if sys.stdout is None:
    class _NullWriter:
        def write(self, s): pass
        def flush(self): pass
        def reconfigure(self, *args, **kwargs): pass
    sys.stdout = _NullWriter()

if sys.stderr is None:
    class _NullWriterErr:
        def write(self, s): pass
        def flush(self): pass
        def reconfigure(self, *args, **kwargs): pass
    sys.stderr = _NullWriterErr()

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
import traceback

def _global_exception_handler(exc_type, exc_value, exc_traceback):
    try:
        with open(BASE_DIR / "crash.log", "w", encoding="utf-8") as f:
            traceback.print_exception(exc_type, exc_value, exc_traceback, file=f)
    except Exception:
        pass
sys.excepthook = _global_exception_handler
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
try:
    import pymupdf
    fitz = pymupdf
    HAS_PYMUPDF = True
except ImportError:
    try:
        import fitz
        HAS_PYMUPDF = True
    except ImportError:
        fitz = None
        HAS_PYMUPDF = False
from PIL import Image
from dotenv import load_dotenv
from pypdf import PdfReader, PdfWriter

# 全域 COM Automation 互斥鎖，解決多執行緒併發 STA 死鎖與 BUSY 例外
COM_LOCK = threading.Lock()

# 心跳守護變數：關閉所有瀏覽器標籤頁後自動停止伺服器 (防 Chrome 背景標籤頁節流誤殺)
LAST_HEARTBEAT_TIME = time.time()
HEARTBEAT_TIMEOUT = 300.0  # 寬限 5 分鐘無任何網頁連線才安全停止伺服器

# 讀取 .env 設定檔
load_dotenv()

BASE_DIR = Path(__file__).resolve().parent


def _resolve_config_path(value: str) -> Path:
    """將相對設定路徑固定解析為專案根目錄下的路徑。"""
    path = Path(value)
    return (BASE_DIR / path).resolve() if not path.is_absolute() else path.resolve()


def _is_safe_storage_directory(path_value: str) -> bool:
    """只允許清理專案根目錄內、且非專案根目錄本身的暫存資料夾。"""
    try:
        path = Path(path_value).resolve()
        project_root = BASE_DIR.resolve()
        return path != project_root and path.is_relative_to(project_root)
    except (OSError, ValueError):
        return False


def _unique_output_path(output_dir: Path, filename: str) -> Path:
    """產生輸出資料夾內不覆寫既有檔案的安全路徑。"""
    candidate = (output_dir / filename).resolve()
    if not candidate.is_relative_to(output_dir.resolve()):
        raise ValueError("輸出檔案路徑不在指定資料夾內")

    counter = 1
    while candidate.exists():
        candidate = output_dir / f"{Path(filename).stem}_{counter}{Path(filename).suffix}"
        counter += 1
    return candidate


def _safe_pdf_filename(filename: str) -> str | None:
    """只接受單一 PDF 檔名，避免路徑跳脫或覆寫任意位置。"""
    if not filename or "/" in filename or "\\" in filename:
        return None
    candidate = Path(filename)
    if candidate.is_absolute() or candidate.name in ("", ".", ".."):
        return None
    name = candidate.name
    return name if name.lower().endswith(".pdf") else f"{name}.pdf"


def _is_loopback_host(host: str) -> bool:
    """判斷伺服器是否僅對本機開放。"""
    return host.strip().lower() in {"localhost", "127.0.0.1", "::1"}

CONFIG = {
    "host": os.getenv("HOST", "0.0.0.0" if os.getenv("RENDER") or os.getenv("DOCKER") else "127.0.0.1"),
    "port": int(os.getenv("PORT", "8080")),
    "upload_dir": os.getenv("UPLOAD_DIR", str(BASE_DIR / "uploads")),
    "output_dir": os.getenv("OUTPUT_DIR", str(BASE_DIR / "converted")),
    "access_token": os.getenv("ACCESS_TOKEN", ""),
    "excel_fit_to_page": os.getenv("EXCEL_FIT_TO_PAGE", "True").lower() in ("true", "1"),
    "log_level": os.getenv("LOG_LEVEL", "INFO"),
}

CONFIG["upload_dir"] = str(_resolve_config_path(CONFIG["upload_dir"]))
CONFIG["output_dir"] = str(_resolve_config_path(CONFIG["output_dir"]))

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
            with COM_LOCK:
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
                        try:
                            doc.Close(False)
                        except Exception:
                            pass
                    if word:
                        try:
                            word.Quit()
                        except Exception:
                            pass
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
    def _is_sheet_empty(excel_app, ws) -> bool:
        """判定 Excel 工作表是否完全無內容 (既無儲存格資料，亦無圖形/圖表/圖片)"""
        try:
            if ws.Shapes.Count > 0:
                return False
            non_empty_count = excel_app.WorksheetFunction.CountA(ws.Cells)
            return non_empty_count == 0
        except Exception:
            return False

    @staticmethod
    def convert_excel(input_path: str, output_path: str) -> list[str]:
        """將 Excel 檔案 (.xls, .xlsx) 的每個有效可見分頁分別匯出為獨立的 PDF 檔案 (自動過濾全空白分頁)"""
        input_abs = str(Path(input_path).resolve())
        output_base_path = Path(output_path).resolve()
        output_dir = output_base_path.parent
        stem = output_base_path.stem

        generated_pdfs = []

        if sys.platform == "win32":
            with COM_LOCK:
                excel, wb = None, None
                try:
                    pythoncom.CoInitialize()
                    excel = win32com.client.DispatchEx("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                    wb = excel.Workbooks.Open(input_abs, ReadOnly=True)

                    visible_sheets = [ws for ws in wb.Worksheets if ws.Visible == -1]
                    # 自動過濾無資料與無圖形的完全空白工作表
                    valid_sheets = [ws for ws in visible_sheets if not DocumentConverter._is_sheet_empty(excel, ws)]
                    if not valid_sheets:
                        valid_sheets = visible_sheets  # 保底策略：若整份活頁簿皆為空，保留所有可見分頁

                    if len(valid_sheets) <= 1:
                        if valid_sheets:
                            ws = valid_sheets[0]
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
                        for ws in valid_sheets:
                            sheet_name = ws.Name
                            safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (" ", "_", "-")).strip() or "工作表"
                            sheet_pdf_path = str(_unique_output_path(output_dir, f"{stem}_{safe_sheet_name}.pdf"))

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
                        try:
                            wb.Close(False)
                        except Exception:
                            pass
                    if excel:
                        try:
                            excel.Quit()
                        except Exception:
                            pass
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
            with COM_LOCK:
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
                        try:
                            presentation.Close()
                        except Exception:
                            pass
                    if ppt:
                        try:
                            ppt.Quit()
                        except Exception:
                            pass
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
        writer = PdfWriter()
        try:
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
            print(f"✅ [PDF 合併成功] ({append_count}/{len(pdf_paths)} 檔) -> {output_abs}")
            return True
        except Exception as e:
            print(f"❌ [PDF 合併失敗]: {e}")
            return False
        finally:
            try:
                writer.close()
            except Exception:
                pass

    @staticmethod
    def convert_pdf_to_images(input_pdf_path: str, output_dir: str, dpi: int = 200) -> list[str]:
        """將 PDF 檔案的每一個頁面獨立拆分導出為高畫質 PNG 圖片"""
        if not HAS_PYMUPDF or fitz is None:
            print("❌ [PDF 轉圖片失敗] 尚未安裝 PyMuPDF 套件，請執行 pip install PyMuPDF。")
            return []

        input_abs = str(Path(input_pdf_path).resolve())
        output_dir_path = Path(output_dir).resolve()
        stem = Path(input_pdf_path).stem
        generated_imgs = []

        try:
            doc = fitz.open(input_abs)
            total_pages = len(doc)

            if total_pages == 0:
                print(f"⚠️ PDF 檔案無頁面: {input_abs}")
                doc.close()
                return []

            for page_index in range(total_pages):
                page = doc[page_index]
                pix = page.get_pixmap(dpi=dpi)

                if total_pages == 1:
                    filename = f"{stem}.png"
                else:
                    filename = f"{stem}_page_{page_index + 1}.png"

                output_img_path = _unique_output_path(output_dir_path, filename)
                pix.save(str(output_img_path))
                generated_imgs.append(str(output_img_path))

            doc.close()
            print(f"✅ [PDF 反向轉圖片成功] 共拆分產出 {len(generated_imgs)} 張高畫質 PNG -> {output_dir_path}")
            return generated_imgs
        except Exception as e:
            print(f"❌ [PDF 轉圖片失敗] {input_abs}: {e}")
            return []

    @staticmethod
    def split_pdf(input_pdf_path: str, output_dir: str) -> list[str]:
        """將多頁 PDF 檔案的每個頁面獨立拆分導出為單頁 PDF (採用 pypdf 無損向量抽取)"""
        input_abs = str(Path(input_pdf_path).resolve())
        output_dir_path = Path(output_dir).resolve()
        stem = Path(input_pdf_path).stem
        generated_pdfs = []
        reader = None
        try:
            reader = PdfReader(input_abs)
            total_pages = len(reader.pages)
            if total_pages == 0:
                print(f"⚠️ PDF 檔案無頁面: {input_abs}")
                return []

            if total_pages == 1:
                out_path = _unique_output_path(output_dir_path, f"{stem}.pdf")
                writer = PdfWriter()
                writer.add_page(reader.pages[0])
                with open(out_path, "wb") as f:
                    writer.write(f)
                writer.close()
                generated_pdfs.append(str(out_path))
            else:
                for idx, page in enumerate(reader.pages):
                    out_path = _unique_output_path(output_dir_path, f"{stem}_第{idx + 1}頁.pdf")
                    writer = PdfWriter()
                    writer.add_page(page)
                    with open(out_path, "wb") as f:
                        writer.write(f)
                    writer.close()
                    generated_pdfs.append(str(out_path))

            print(f"✅ [PDF 拆單頁成功] 共拆分產出 {len(generated_pdfs)} 個獨立單頁 PDF -> {output_dir_path}")
            return generated_pdfs
        except Exception as e:
            print(f"❌ [PDF 拆單頁失敗] {input_abs}: {e}")
            return []
        finally:
            if reader and hasattr(reader, "stream") and hasattr(reader.stream, "close"):
                try:
                    reader.stream.close()
                except Exception:
                    pass


# ============================================================
# Web UI 前端 HTML 樣式 (含手動清理暫存檔與 PPT/PPTX 支援)
# ============================================================
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PaperSwitch - 萬能文件轉 PDF 處理器</title>
    <style>
        :root {
            --font-heading: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft JhengHei UI", "Microsoft JhengHei", Arial, sans-serif;
            --font-body: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft JhengHei UI", "Microsoft JhengHei", Arial, sans-serif;
            --bg-color: #161c28;
            --card-bg: rgba(26, 34, 48, 0.88);
            --panel-bg: rgba(18, 24, 35, 0.65);
            --border-color: rgba(255, 255, 255, 0.11);
            --border-hover: rgba(107, 164, 200, 0.55);
            --accent-color: #6ba4c8;
            --accent-hover: #568eb2;
            --secondary-bg: rgba(255, 255, 255, 0.08);
            --secondary-hover: rgba(255, 255, 255, 0.15);
            --text-h1: #ffffff;
            --text-h2: #f1f5f9;
            --text-h3: #e2e8f0;
            --text-body: #cbd5e1;
            --text-sub: #94a3b8;
            --success-color: #7ea88f;
            --warning-color: #d1a368;
            --error-color: #d47a7a;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        
        html, body {
            height: 100vh;
            overflow: hidden;
            background-color: var(--bg-color);
            color: var(--text-body);
            font-family: var(--font-body);
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
        }
        
        body {
            display: flex;
            flex-direction: column;
            padding: 16px 24px 24px;
            box-sizing: border-box;
        }
        
        header {
            text-align: center;
            margin-bottom: 16px;
            flex-shrink: 0;
        }
        header h1 {
            font-family: var(--font-heading);
            font-size: 2.1rem;
            font-weight: 800;
            letter-spacing: -0.03em;
            background: linear-gradient(135deg, #ffffff 40%, #9bc8e4 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            text-shadow: 0 4px 16px rgba(0, 0, 0, 0.4);
        }
        header p {
            font-family: var(--font-body);
            color: var(--text-sub);
            font-size: 0.95rem;
            font-weight: 500;
            margin-top: 4px;
            letter-spacing: 0.02em;
        }

        .dashboard {
            display: grid;
            grid-template-columns: 0.95fr 1.25fr 1.15fr;
            gap: 18px;
            width: 82vw;
            max-width: 1600px;
            min-width: 1040px;
            margin: 0 auto;
            align-items: stretch;
            height: calc(100vh - 120px);
            min-height: 640px;
        }
        @media (max-width: 1080px) {
            html, body { overflow: auto; height: auto; }
            .dashboard { grid-template-columns: 1fr; width: 94vw; min-width: auto; height: auto; }
        }

        .panel {
            background: var(--card-bg);
            backdrop-filter: blur(16px);
            border: 1px solid var(--border-color);
            border-radius: 16px;
            padding: 20px;
            display: flex;
            flex-direction: column;
            box-shadow: 0 20px 40px -15px rgba(0, 0, 0, 0.55);
            min-height: 0;
            height: 100%;
        }

        .panel-upload { grid-column: auto; grid-row: auto; }
        .panel-queue { grid-column: auto; grid-row: auto; }
        .panel-settings { grid-column: auto; grid-row: auto; }
        
        .panel-title {
            font-family: var(--font-heading);
            font-size: 1.22rem;
            font-weight: 800;
            color: var(--text-h1);
            letter-spacing: -0.015em;
            margin-bottom: 16px;
            display: flex;
            align-items: center;
            gap: 8px;
            border-bottom: 2px solid rgba(107, 164, 200, 0.25);
            padding-bottom: 12px;
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
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            flex: 1;
            min-height: 0;
        }
        .drop-zone:hover, .drop-zone.dragover { border-color: var(--accent-color); background: rgba(107, 164, 200, 0.12); transform: translateY(-2px); }
        .drop-icon { font-size: 52px; margin-bottom: 16px; filter: drop-shadow(0 4px 12px rgba(209, 163, 104, 0.3)); }
        .drop-text { font-family: var(--font-heading); font-size: 1.16rem; font-weight: 800; color: var(--text-h1); letter-spacing: -0.01em; }
        .drop-hint { font-family: var(--font-body); font-size: 0.92rem; font-weight: 500; color: var(--text-sub); margin-top: 10px; line-height: 1.7; }
        .format-chips { display: flex; flex-wrap: wrap; justify-content: center; gap: 6px; margin-top: 16px; }
        .format-chip { padding: 4px 10px; color: var(--text-h2); background: rgba(142, 156, 174, 0.18); border: 1px solid rgba(142, 156, 174, 0.32); border-radius: 999px; font-size: 0.8rem; font-weight: 600; line-height: 1.3; }

        .progress-box { margin-bottom: 14px; flex-shrink: 0; }
        .progress-info { font-family: var(--font-body); display: flex; justify-content: space-between; font-size: 0.9rem; font-weight: 700; margin-bottom: 6px; color: var(--text-h2); }
        .progress-bar-bg { width: 100%; height: 9px; background: rgba(255,255,255,0.1); border-radius: 5px; overflow: hidden; }
        .progress-bar-fill { height: 100%; width: 0%; background: var(--accent-color); transition: width 0.3s ease; }
        
        .queue-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; flex-shrink: 0; }
        .queue-count { font-size: 0.9rem; font-weight: 700; color: var(--text-h3); }
        .clear-btn { font-size: 0.84rem; font-weight: 700; color: var(--error-color); background: none; border: none; cursor: pointer; text-decoration: underline; }
        
        .file-queue {
            overflow-y: auto;
            display: flex;
            flex-direction: column;
            gap: 8px;
            padding-right: 4px;
            flex: 1;
            min-height: 180px;
            max-height: 320px;
        }
        .file-card { display: flex; justify-content: space-between; align-items: center; padding: 11px 13px; background: var(--panel-bg); border: 1px solid rgba(255,255,255,0.08); border-radius: 8px; font-size: 0.88rem; }
        .file-info { display: flex; flex-direction: column; gap: 3px; overflow: hidden; }
        .file-name { font-weight: 700; color: var(--text-h1); font-size: 0.92rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 220px; }
        .file-meta { font-size: 0.78rem; font-weight: 500; color: var(--text-sub); }
        .file-status { font-size: 0.78rem; font-weight: 700; padding: 3px 8px; border-radius: 4px; transition: all 0.2s ease; }
        
        @keyframes pulseGlow {
            0% { opacity: 0.75; }
            50% { opacity: 1; transform: scale(1.02); }
            100% { opacity: 0.75; }
        }
        .status-ready { background: rgba(255,255,255,0.1); color: var(--text-body); border: 1px solid rgba(255,255,255,0.15); }
        .status-processing { background: rgba(209, 163, 104, 0.25); color: var(--warning-color); border: 1px solid rgba(209, 163, 104, 0.5); animation: pulseGlow 1.4s infinite ease-in-out; }
        .status-done { background: rgba(126, 168, 143, 0.25); color: var(--success-color); border: 1px solid rgba(126, 168, 143, 0.45); }
        .status-fail { background: rgba(212, 122, 122, 0.25); color: var(--error-color); border: 1px solid rgba(212, 122, 122, 0.45); }

        .option-group { margin-bottom: 12px; display: flex; flex-direction: column; gap: 8px; flex-shrink: 0; }
        .option-label { font-family: var(--font-heading); font-size: 0.96rem; font-weight: 800; color: var(--text-h2); letter-spacing: -0.01em; }
        .radio-card { display: flex; align-items: flex-start; gap: 10px; padding: 12px 14px; background: var(--panel-bg); border: 1px solid var(--border-color); border-radius: 10px; cursor: pointer; transition: all 0.2s ease; }
        .radio-card:hover { border-color: var(--border-hover); background: rgba(107, 164, 200, 0.1); }
        .radio-card:focus-within { outline: 2px solid rgba(107, 164, 200, 0.75); outline-offset: 2px; }
        .radio-card.active { border-color: var(--accent-color); background: linear-gradient(135deg, rgba(107, 164, 200, 0.22), rgba(107, 164, 200, 0.08)); box-shadow: inset 3.5px 0 0 var(--accent-color); }
        .radio-card.active .radio-title { color: #ffffff; }
        .radio-card input { margin-top: 3px; accent-color: var(--accent-color); }
        .radio-text { display: flex; flex-direction: column; gap: 3px; }
        .radio-title { font-family: var(--font-heading); font-size: 1.02rem; font-weight: 800; color: var(--text-h1); letter-spacing: -0.01em; }
        .radio-desc { font-size: 0.82rem; font-weight: 500; color: var(--text-sub); line-height: 1.4; }
        
        .input-group { display: flex; flex-direction: column; gap: 6px; margin-bottom: 12px; flex-shrink: 0; }
        .input-field { width: 100%; padding: 10px 12px; background: rgba(18, 24, 35, 0.7); border: 1px solid var(--border-color); border-radius: 8px; color: var(--text-h1); font-size: 0.9rem; font-weight: 600; outline: none; }
        .input-field:focus { border-color: var(--accent-color); }

        .btn-group { display: flex; flex-direction: column; gap: 8px; margin-top: 8px; flex-shrink: 0; }
        .btn { font-family: var(--font-heading); width: 100%; padding: 13px; background: var(--accent-color); color: #0b121e; font-weight: 800; border: none; border-radius: 10px; cursor: pointer; transition: all 0.2s ease; font-size: 1.02rem; text-align: center; letter-spacing: 0.02em; }
        .btn:hover { background: var(--accent-hover); color: #ffffff; }
        .btn-secondary { background: var(--secondary-bg); color: var(--text-h1); border: 1px solid var(--border-color); font-size: 0.95rem; }
        .btn-secondary:hover { background: var(--secondary-hover); }

        .maintenance-actions { margin-top: auto; padding-top: 14px; border-top: 1px solid var(--border-color); }
        .maintenance-label { display: block; margin-bottom: 8px; color: var(--text-sub); font-size: 0.78rem; font-weight: 800; letter-spacing: 0.05em; text-transform: uppercase; }
        .maintenance-btns { display: flex; flex-wrap: wrap; gap: 6px; }
        .maintenance-btns .btn-tool { width: auto; min-height: auto; padding: 7px 11px; font-size: 0.82rem; font-family: var(--font-heading); font-weight: 700; border-radius: 8px; cursor: pointer; transition: all 0.2s ease; }
        .btn-tool-secondary { background: var(--secondary-bg); color: var(--text-h2); border: 1px solid var(--border-color); }
        .btn-tool-secondary:hover { background: var(--secondary-hover); border-color: rgba(107, 164, 200, 0.4); }
        .btn-tool-danger { background: rgba(212, 122, 122, 0.12); color: var(--error-color); border: 1px solid rgba(212, 122, 122, 0.35); }
        .btn-tool-danger:hover { background: rgba(212, 122, 122, 0.22); }

        .log-console-box { margin-top: 12px; background: rgba(10, 14, 22, 0.85); border: 1px solid var(--border-color); border-radius: 10px; padding: 11px 13px; flex-shrink: 0; }
        .log-console-header { display: flex; justify-content: space-between; align-items: center; font-size: 0.82rem; font-weight: 800; color: var(--text-h3); margin-bottom: 6px; }
        .log-clear-btn { background: none; border: none; color: var(--text-sub); font-size: 0.76rem; font-weight: 600; cursor: pointer; text-decoration: underline; }
        .log-clear-btn:hover { color: var(--text-h1); }
        .log-console { height: 130px; overflow-y: auto; font-family: Consolas, "Courier New", monospace; font-size: 0.78rem; display: flex; flex-direction: column; gap: 4px; }
        .log-line { line-height: 1.45; word-break: break-all; }
        .log-info { color: var(--text-sub); }
        .log-success { color: var(--success-color); font-weight: 600; }
        .log-warn { color: var(--warning-color); }
        .log-error { color: var(--error-color); font-weight: 600; }

        #globalStatus { margin-top: 8px; text-align: center; font-size: 0.9rem; font-weight: 700; min-height: 20px; flex-shrink: 0; }
    </style>
</head>
<body>
    <header>
        <h1>📄 PaperSwitch</h1>
        <p>萬能 Word / Excel / PowerPoint / 圖片 轉 PDF 與 PDF 合併引擎</p>
    </header>

    <div class="dashboard">
        <!-- 區塊 1: 檔案拖曳區 (左欄) -->
        <div class="panel panel-upload">
            <h2 class="panel-title">📥 1. 檔案拖曳區</h2>
            <div class="drop-zone" id="dropZone" onclick="document.getElementById('fileInput').click()">
                <div class="drop-icon">✨</div>
                <div class="drop-text">將檔案拖曳至此處</div>
                <div class="drop-hint">或點擊此區域選擇檔案</div>
                <div class="format-chips" aria-label="支援格式">
                    <span class="format-chip">Word</span>
                    <span class="format-chip">Excel</span>
                    <span class="format-chip">PowerPoint</span>
                    <span class="format-chip">圖片</span>
                    <span class="format-chip">PDF</span>
                </div>
                <input type="file" id="fileInput" multiple style="display: none;" onchange="handleFiles(this.files)">
            </div>
        </div>

        <!-- 區塊 2: 檔案佇列與進度 (中欄) -->
        <div class="panel panel-queue">
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

            <!-- 即時動態日誌終端框 -->
            <div class="log-console-box">
                <div class="log-console-header">
                    <span>🖥️ 系統執行日誌</span>
                    <button class="log-clear-btn" type="button" onclick="clearConsoleLog()">清除</button>
                </div>
                <div class="log-console" id="logConsole">
                    <div class="log-line log-info">[系統連線] PaperSwitch 伺服器運作正常</div>
                </div>
            </div>
        </div>

        <!-- 區塊 3: 轉換功能設定 (右欄) -->
        <div class="panel panel-settings">
            <h2 class="panel-title">⚙️ 3. 轉換功能設定</h2>
            
            <div class="option-group">
                <h3 class="option-label">輸出模式設定：</h3>
                
                <label class="radio-card active" id="cardSingle" onclick="setMode('single')">
                    <input type="radio" name="convertMode" value="single" checked>
                    <div class="radio-text">
                        <span class="radio-title">📄 文件轉 PDF</span>
                        <span class="radio-desc">每個檔案個別轉換為對應同名 PDF (Excel 多分頁自動拆分獨立產出)</span>
                    </div>
                </label>

                <label class="radio-card" id="cardMerge" onclick="setMode('merge')">
                    <input type="radio" name="convertMode" value="merge">
                    <div class="radio-text">
                        <span class="radio-title">📚 多檔併 PDF</span>
                        <span class="radio-desc">將佇列中所有檔案按自訂順序合併為單一完整 PDF 檔案</span>
                    </div>
                </label>

                <label class="radio-card" id="cardSplit" onclick="setMode('split')">
                    <input type="radio" name="convertMode" value="split">
                    <div class="radio-text">
                        <span class="radio-title">✂️ PDF 拆單頁</span>
                        <span class="radio-desc">將多頁 PDF 檔案的每個頁面獨立拆解導出為單頁 PDF</span>
                    </div>
                </label>

                <label class="radio-card" id="cardPdfToImg" onclick="setMode('pdf_to_images')">
                    <input type="radio" name="convertMode" value="pdf_to_images">
                    <div class="radio-text">
                        <span class="radio-title">🖼️ PDF 轉圖片</span>
                        <span class="radio-desc">將 PDF 檔案的每個頁面獨立渲染導出為高清 PNG 圖片</span>
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
            </div>

            <div id="globalStatus" aria-live="polite"></div>

            <div class="maintenance-actions">
                <span class="maintenance-label">維護工具</span>
                <div class="maintenance-btns">
                    <button class="btn-tool btn-tool-secondary" onclick="openOutputFolder()" title="開啟本機 converted/ PDF 輸出資料夾">📂 開啟輸出資料夾</button>
                    <button class="btn-tool btn-tool-danger" onclick="clearStorage()" title="清理 uploads/ 與 converted/ 暫存檔">🧹 清除歷史暫存檔</button>
                    <button class="btn-tool btn-tool-danger" onclick="shutdownServer()" title="停止後端伺服器並關閉進程">🛑 關閉伺服器</button>
                </div>
            </div>
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
            document.getElementById('cardSplit').classList.toggle('active', mode === 'split');
            document.getElementById('cardPdfToImg').classList.toggle('active', mode === 'pdf_to_images');
            document.querySelector('input[value="single"]').checked = (mode === 'single');
            document.querySelector('input[value="merge"]').checked = (mode === 'merge');
            document.querySelector('input[value="split"]').checked = (mode === 'split');
            document.querySelector('input[value="pdf_to_images"]').checked = (mode === 'pdf_to_images');
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
                if (item.status === 'processing') { badgeClass = 'status-processing'; statusLabel = '處理中'; }
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

        function addLog(msg, type = 'info') {
            const consoleEl = document.getElementById('logConsole');
            if (!consoleEl) return;
            const now = new Date();
            const timeStr = now.toTimeString().split(' ')[0];
            const line = document.createElement('div');
            line.className = `log-line log-${type}`;
            line.innerText = `[${timeStr}] ${msg}`;
            consoleEl.appendChild(line);
            consoleEl.scrollTop = consoleEl.scrollHeight;
        }

        function clearConsoleLog() {
            const consoleEl = document.getElementById('logConsole');
            if (consoleEl) consoleEl.innerHTML = '';
        }

        async function shutdownServer() {
            if (confirm('確定要關閉 PaperSwitch 後端伺服器並結束行程嗎？')) {
                addLog('正在通知後端伺服器安全停止...', 'warn');
                try {
                    await postApi('/api/shutdown');
                    addLog('✅ 伺服器已成功停止，您現在可以安全關閉此網頁。', 'success');
                    alert('伺服器已安全停止。');
                } catch (e) {
                    addLog('伺服器已結束連線。', 'info');
                }
            }
        }

        // 前端每 3 秒發送心跳維持連線
        setInterval(() => {
            fetch('/api/heartbeat').catch(() => {});
        }, 3000);

        // 當切換回分頁或點擊頁面時，立即補發心跳喚醒 (防 Chrome 背景標籤頁節流)
        ['visibilitychange', 'focus', 'click'].forEach(evt => {
            window.addEventListener(evt, () => {
                fetch('/api/heartbeat').catch(() => {});
            });
        });

        async function uploadAndConvert() {
            if (selectedFiles.length === 0) { alert('請先新增檔案至佇列！'); return; }

            const globalStatus = document.getElementById('globalStatus');
            const openFolderBtn = document.getElementById('openFolderBtn');

            globalStatus.className = '';
            openFolderBtn.style.display = 'none';

            let mergedName = document.getElementById('mergedFilename').value.trim();
            if (!mergedName.endsWith('.pdf')) mergedName += '.pdf';

            const total = selectedFiles.length;

            if (currentMode === 'merge') {
                selectedFiles.forEach(item => item.status = 'processing');
                renderQueue();
                updateProgress(25, `正在上傳 ${total} 個檔案並準備合併...`);
                globalStatus.innerText = '⏳ 正在合併處理中...';
                addLog(`開始合併 ${total} 個檔案 ➔ ${mergedName}`, 'info');

                const formData = new FormData();
                selectedFiles.forEach(item => formData.append('files', item.file));
                formData.append('mode', 'merge');
                formData.append('merged_filename', mergedName);

                try {
                    updateProgress(65, '正在合併渲染 PDF...');
                    const resp = await postApi('/api/convert', formData);
                    const res = await resp.json();

                    if (res.success) {
                        selectedFiles.forEach(item => item.status = 'done');
                        updateProgress(100, '合併成功');
                        globalStatus.className = 'status-success';
                        globalStatus.innerText = `✅ 已成功合併 ${total} 個檔案為 ${mergedName}`;
                        addLog(`✅ [多檔併 PDF 成功] 共 ${total} 個檔案 ➔ converted/${mergedName}`, 'success');
                        openFolderBtn.style.display = 'block';
                    } else {
                        selectedFiles.forEach(item => item.status = 'fail');
                        updateProgress(100, '合併失敗');
                        globalStatus.className = 'status-error';
                        globalStatus.innerText = '❌ 合併失敗：' + (res.message || '未知錯誤');
                        addLog(`❌ [多檔併 PDF 失敗] ${res.message || '未知錯誤'}`, 'error');
                    }
                } catch (err) {
                    selectedFiles.forEach(item => item.status = 'fail');
                    updateProgress(100, '網路異常');
                    globalStatus.className = 'status-error';
                    globalStatus.innerText = '❌ 網路連線或伺服器異常';
                    addLog('❌ 網路連線或伺服器通訊異常', 'error');
                }
                renderQueue();
            } else {
                let successCount = 0;
                let failCount = 0;

                selectedFiles.forEach(item => item.status = 'ready');
                renderQueue();
                updateProgress(0, `準備依序處理 (共 ${total} 個)...`);
                addLog(`開始依序處理佇列 (共 ${total} 個檔案)...`, 'info');

                for (let i = 0; i < total; i++) {
                    const item = selectedFiles[i];
                    item.status = 'processing';
                    renderQueue();

                    const currentPercent = Math.round((i / total) * 100);
                    let actionName = '轉檔';
                    if (currentMode === 'pdf_to_images') actionName = '轉圖片';
                    else if (currentMode === 'split') actionName = '拆單頁';

                    updateProgress(currentPercent, `正在${actionName} (${i + 1}/${total})：${item.file.name}`);
                    globalStatus.innerText = `⏳ 正在處理第 ${i + 1}/${total} 個檔案：${item.file.name}...`;
                    addLog(`正在${actionName} (${i + 1}/${total})：${item.file.name}`, 'info');

                    const singleFormData = new FormData();
                    singleFormData.append('files', item.file);
                    singleFormData.append('mode', currentMode);

                    try {
                        const resp = await postApi('/api/convert', singleFormData);
                        const res = await resp.json();
                        if (res.success) {
                            item.status = 'done';
                            successCount++;
                            const detailLog = res.log || `處理成功 ➔ converted/`;
                            addLog(`✅ [成功] ${item.file.name} ➔ ${detailLog}`, 'success');
                        } else {
                            item.status = 'fail';
                            failCount++;
                            addLog(`❌ [失敗] ${item.file.name}：${res.message || '處理異常'}`, 'error');
                        }
                    } catch (err) {
                        item.status = 'fail';
                        failCount++;
                        addLog(`❌ [連線異常] ${item.file.name}`, 'error');
                    }

                    renderQueue();
                }

                updateProgress(100, (failCount === 0) ? '全部處理完成' : `完成 ${successCount} 個，失敗 ${failCount} 個`);
                if (failCount === 0) {
                    globalStatus.className = 'status-success';
                    if (currentMode === 'pdf_to_images') {
                        globalStatus.innerText = `✅ 全部 ${total} 個 PDF 已成功轉為圖片！`;
                    } else if (currentMode === 'split') {
                        globalStatus.innerText = `✅ 全部 ${total} 個 PDF 已成功拆分單頁完成！`;
                    } else {
                        globalStatus.innerText = `✅ 佇列 ${total} 個檔案全部轉換完成！`;
                    }
                    addLog(`🎉 全部 ${total} 個項目處理完畢！`, 'success');
                    openFolderBtn.style.display = 'block';
                } else {
                    globalStatus.className = 'status-error';
                    globalStatus.innerText = `⚠️ 處理完畢：${successCount} 個成功，${failCount} 個失敗。`;
                    addLog(`⚠️ 處理完畢：${successCount} 個成功，${failCount} 個失敗。`, 'warn');
                    if (successCount > 0) openFolderBtn.style.display = 'block';
                }
            }
        }

        async function openOutputFolder() {
            try {
                await postApi('/api/open-folder');
                addLog('📂 已開啟本機 converted/ 輸出資料夾', 'info');
            } catch (err) {
                alert('無法開啟本機資料夾');
            }
        }

        async function clearStorage() {
            if (confirm('確定要清空 uploads/ 與 converted/ 資料夾下的所有歷史暫存檔嗎？')) {
                try {
                    const resp = await postApi('/api/clear-storage');
                    const res = await resp.json();
                    if (res.success) {
                        alert(`✅ 清理完成！已清除 ${res.count} 個歷史暫存檔案。`);
                        addLog(`🧹 [暫存清理] 已清除 ${res.count} 個歷史暫存檔案。`, 'info');
                    } else {
                        alert('❌ 清理失敗: ' + res.message);
                    }
                } catch (err) {
                    alert('❌ 網路請求失敗');
                }
            }
        }

        async function postApi(path, body = undefined) {
            const sendRequest = async () => {
                const token = sessionStorage.getItem('paperswitchAccessToken');
                const headers = token ? { 'X-PaperSwitch-Token': token } : {};
                return fetch(path, { method: 'POST', body, headers });
            };

            let response = await sendRequest();
            if (response.status === 401) {
                const token = prompt('此伺服器需要存取權杖，請輸入 ACCESS_TOKEN：');
                if (!token) throw new Error('未提供存取權杖');
                sessionStorage.setItem('paperswitchAccessToken', token);
                response = await sendRequest();
            }
            return response;
        }
    </script>
</body>
</html>
"""


# ============================================================
# HTTP 請求監聽處理器
# ============================================================
class WebAppHandler(SimpleHTTPRequestHandler):

    def log_message(self, format, *args):
        """過濾心跳與圖示輪詢日誌，維持終端視窗清爽聚焦。"""
        if hasattr(self, "path") and self.path in ("/api/heartbeat", "/favicon.ico"):
            return
        super().log_message(format, *args)

    def do_GET(self):
        global LAST_HEARTBEAT_TIME
        parsed = urllib.parse.urlparse(self.path)
        clean_path = parsed.path
        if clean_path in ["/", "/index.html"]:
            LAST_HEARTBEAT_TIME = time.time()
            encoded = HTML_TEMPLATE.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(encoded)))
            self.send_header("Connection", "close")
            self.end_headers()
            self.wfile.write(encoded)
        elif clean_path == "/api/heartbeat":
            LAST_HEARTBEAT_TIME = time.time()
            self._send_json({"status": "alive"})
        elif clean_path == "/favicon.ico":
            self.send_response(204)  # No Content
            self.send_header("Content-Length", "0")
            self.send_header("Connection", "close")
            self.end_headers()
        else:
            self.send_error(404, "Page Not Found")

    def do_POST(self):
        try:
            if not self._is_request_authorized():
                self._send_json({"success": False, "message": "未授權的 API 請求"}, status=401)
                return

            if self.path == "/api/convert":
                saved_paths, form_data = self._parse_multipart()

                if not saved_paths:
                    self._send_json({"success": False, "message": "未接收到有效檔案或表單格式錯誤"})
                    return

                mode = form_data.get("mode", "single")
                if mode not in ("single", "merge", "split", "pdf_to_images"):
                    self._send_json({"success": False, "message": "不支援的輸出模式"}, status=400)
                    return

                merged_filename = _safe_pdf_filename(form_data.get("merged_filename", "combined_output.pdf"))
                if not merged_filename:
                    self._send_json({"success": False, "message": "合併 PDF 檔名不可包含路徑"}, status=400)
                    return

                converter = DocumentConverter()
                generated_pdfs = []
                temporary_pdfs = []
                overall_success = True
                request_id = Path(saved_paths[0]).parent.name

                for p in saved_paths:
                    ext = Path(p).suffix.lower()
                    stem = Path(p).stem
                    temp_pdf = Path(CONFIG["output_dir"]) / f"temp_{request_id}_{stem}.pdf"
                    final_pdf = _unique_output_path(Path(CONFIG["output_dir"]), f"{stem}.pdf")

                    target_path = str(temp_pdf) if mode == "merge" else str(final_pdf)

                    if mode == "split":
                        if ext == ".pdf":
                            split_paths = converter.split_pdf(p, CONFIG["output_dir"])
                            if split_paths:
                                generated_pdfs.extend(split_paths)
                            else:
                                overall_success = False
                        else:
                            temp_pdf = Path(CONFIG["output_dir"]) / f"temp_{request_id}_{stem}.pdf"
                            conv_res = False
                            if ext in [".doc", ".docx"]:
                                conv_res = converter.convert_word(p, str(temp_pdf))
                            elif ext in [".xls", ".xlsx"]:
                                excel_pdfs = converter.convert_excel(p, str(temp_pdf))
                                if excel_pdfs:
                                    temp_pdf = Path(excel_pdfs[0])
                                    conv_res = True
                            elif ext in [".ppt", ".pptx"]:
                                conv_res = converter.convert_powerpoint(p, str(temp_pdf))
                            elif ext in [".png", ".jpg", ".jpeg", ".bmp", ".webp"]:
                                conv_res = converter.convert_image([p], str(temp_pdf))

                            if conv_res and temp_pdf.exists():
                                split_paths = converter.split_pdf(str(temp_pdf), CONFIG["output_dir"])
                                if split_paths:
                                    generated_pdfs.extend(split_paths)
                                else:
                                    overall_success = False
                                try:
                                    temp_pdf.unlink(missing_ok=True)
                                except Exception:
                                    pass
                            else:
                                overall_success = False
                    elif mode == "pdf_to_images":
                        if ext == ".pdf":
                            img_paths = converter.convert_pdf_to_images(p, CONFIG["output_dir"])
                            if img_paths:
                                generated_pdfs.extend(img_paths)
                            else:
                                overall_success = False
                        else:
                            temp_pdf = Path(CONFIG["output_dir"]) / f"temp_{request_id}_{stem}.pdf"
                            conv_res = False
                            if ext in [".doc", ".docx"]:
                                conv_res = converter.convert_word(p, str(temp_pdf))
                            elif ext in [".xls", ".xlsx"]:
                                excel_pdfs = converter.convert_excel(p, str(temp_pdf))
                                if excel_pdfs:
                                    temp_pdf = Path(excel_pdfs[0])
                                    conv_res = True
                            elif ext in [".ppt", ".pptx"]:
                                conv_res = converter.convert_powerpoint(p, str(temp_pdf))
                            elif ext in [".png", ".jpg", ".jpeg", ".bmp", ".webp"]:
                                conv_res = converter.convert_image([p], str(temp_pdf))

                            if conv_res and temp_pdf.exists():
                                img_paths = converter.convert_pdf_to_images(str(temp_pdf), CONFIG["output_dir"])
                                if img_paths:
                                    generated_pdfs.extend(img_paths)
                                else:
                                    overall_success = False
                                try:
                                    temp_pdf.unlink(missing_ok=True)
                                except Exception:
                                    pass
                            else:
                                overall_success = False
                    elif ext in [".doc", ".docx"]:
                        res = converter.convert_word(p, target_path)
                        if res:
                            generated_pdfs.append(target_path)
                            if mode == "merge":
                                temporary_pdfs.append(target_path)
                        else:
                            overall_success = False
                    elif ext in [".xls", ".xlsx"]:
                        excel_pdfs = converter.convert_excel(p, target_path)
                        if excel_pdfs:
                            generated_pdfs.extend(excel_pdfs)
                            if mode == "merge":
                                temporary_pdfs.extend(excel_pdfs)
                        else:
                            overall_success = False
                    elif ext in [".ppt", ".pptx"]:
                        res = converter.convert_powerpoint(p, target_path)
                        if res:
                            generated_pdfs.append(target_path)
                            if mode == "merge":
                                temporary_pdfs.append(target_path)
                        else:
                            overall_success = False
                    elif ext in [".png", ".jpg", ".jpeg", ".bmp", ".webp"]:
                        res = converter.convert_image([p], target_path)
                        if res:
                            generated_pdfs.append(target_path)
                            if mode == "merge":
                                temporary_pdfs.append(target_path)
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
                    merged_output_path = str(_unique_output_path(Path(CONFIG["output_dir"]), merged_filename))
                    merge_res = converter.merge_pdfs(generated_pdfs, merged_output_path)

                    for tmp_pdf in temporary_pdfs:
                        if Path(tmp_pdf).exists():
                            try:
                                os.remove(tmp_pdf)
                            except Exception:
                                pass

                    if not merge_res:
                        overall_success = False

                log_desc = f"產出 {len(generated_pdfs)} 個項目至 converted/" if overall_success else "轉檔處理異常"
                self._send_json(
                    {
                        "success": overall_success,
                        "message": "轉換完成" if overall_success else "部分或全部檔案轉檔失敗",
                        "outputs": [Path(p).name for p in generated_pdfs],
                        "log": log_desc,
                    }
                )

            elif self.path == "/api/clear-storage":
                cleaned_count = self._clear_storage()
                self._send_json({"success": True, "count": cleaned_count, "message": "歷史暫存檔已清空"})

            elif self.path == "/api/open-folder":
                self._open_converted_folder()
                self._send_json({"success": True, "message": "已開啟輸出資料夾"})

            elif self.path == "/api/shutdown":
                self._send_json({"success": True, "message": "伺服器正在安全停止..."})
                threading.Thread(target=self.server.shutdown).start()
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
            if not _is_safe_storage_directory(folder):
                print(f"❌ 拒絕清理專案根目錄外或根目錄本身的資料夾: {folder}")
                continue
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

        request_upload_dir = Path(CONFIG["upload_dir"]) / uuid.uuid4().hex
        request_upload_dir.mkdir(parents=True, exist_ok=False)
        saved_files = []
        form_fields = {}

        for part in parts:
            if not part or part == b"\r\n" or part.startswith(b"--"):
                continue

            if part.startswith(b"\r\n"):
                part = part[2:]

            if b"\r\n\r\n" not in part:
                continue

            headers_blob, data = part.split(b"\r\n\r\n", 1)

            # 根據 RFC 7578，delimiter 為 \r\n--boundary，分割後各 part 之 data 尾端完全獨立且乾淨。
            # 若最後一個 part 後方帶有結尾標籤指示符 `--` 或 `--\r\n`，會在 split 的末尾區段過濾掉，
            # 絕對不得對解出之 data 二進位內容執行盲目截斷 (data.endswith(b"--"))。

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
                save_path = request_upload_dir / safe_filename

                # 同名檔案自動補上計數器後綴，防範複數同名檔上傳被覆蓋
                counter = 1
                stem = save_path.stem
                suffix = save_path.suffix
                while save_path.exists():
                    save_path = request_upload_dir / f"{stem}_{counter}{suffix}"
                    counter += 1

                with open(save_path, "wb") as f:
                    f.write(data)
                saved_files.append(str(save_path))
            elif field_name and not filename:
                form_fields[field_name] = data.decode("utf-8", errors="ignore")

        return saved_files, form_fields

    def _is_request_authorized(self) -> bool:
        """公開綁定時，要求 API 呼叫者提供環境設定的存取權杖。"""
        token = CONFIG["access_token"]
        return not token or hmac.compare_digest(self.headers.get("X-PaperSwitch-Token", ""), token)

    def _send_json(self, data: dict, status: int = 200):
        body = json.dumps(data).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Connection", "close")
        self.end_headers()
        self.wfile.write(body)


def _heartbeat_watchdog(server):
    """心跳守護線程：當超過寬限期未收到任何瀏覽器分頁心跳時，自動安全停止伺服器並結束行程"""
    # 給予啟動初期 45 秒寬限期，等待瀏覽器完成開啟與首次頁面載入
    time.sleep(45)
    while True:
        time.sleep(3)
        if time.time() - LAST_HEARTBEAT_TIME > HEARTBEAT_TIMEOUT:
            print("\n[PaperSwitch] 偵測到所有網頁分頁已關閉（心跳超時），自動安全停止伺服器並退出行程...")
            threading.Thread(target=server.shutdown).start()
            break


# ============================================================
# 主入口程式 (升級 ThreadingHTTPServer 多執行緒併發處理與心跳守護)
# ============================================================
def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    if not _is_loopback_host(CONFIG["host"]) and not CONFIG["access_token"]:
        raise RuntimeError("公開綁定 HOST 時必須設定 ACCESS_TOKEN，避免未授權轉檔與清理")

    server_address = (CONFIG["host"], CONFIG["port"])
    httpd = ThreadingHTTPServer(server_address, WebAppHandler)
    url = f"http://{CONFIG['host']}:{CONFIG['port']}"
    print(f"[PaperSwitch] 多執行緒轉檔伺服器運行中: {url}")

    # 啟動心跳監控守護線程 (寬鬆待命模式)
    # watchdog = threading.Thread(target=_heartbeat_watchdog, args=(httpd,), daemon=True)
    # watchdog.start()

    # 自動開啟預設瀏覽器
    webbrowser.open(url)

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[INFO] 伺服器已安全停止")


if __name__ == "__main__":
    main()
