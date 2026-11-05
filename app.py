import base64
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path
from typing import List, Optional, Set

from fastapi import FastAPI, File, Form, HTTPException, UploadFile, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse

IS_WINDOWS = sys.platform == "win32"
MAX_FILE_SIZE = 25 * 1024 * 1024
ALLOWED_EXTENSIONS = {".pdf", ".png", ".jpg", ".jpeg"}
ALLOWED_MIME_TYPES = {
    "application/pdf",
    "image/png",
    "image/jpeg",
    "image/jpg",
}
DEFAULT_TEMP_DIR = Path(os.environ.get("TEMP_DIR", Path(tempfile.gettempdir()) / "print_jobs"))
DEFAULT_PRINTER_ENV = os.environ.get("PRINTER_NAME", "").strip() or None

app = FastAPI(title="Web Print Server", docs_url=None, redoc_url=None)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def sanitize_printer_name(name: str) -> str:
    cleaned = name.strip()
    if not re.match(r"^[a-zA-Z0-9_\-\s\(\)\.]+$", cleaned):
        raise ValueError("Invalid printer name characters.")
    return cleaned

def get_available_printers() -> List[str]:
    if IS_WINDOWS:
        try:
            import win32print
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            printers = [p[2] for p in win32print.EnumPrinters(flags)]
            return sorted(list(dict.fromkeys(printers)))
        except Exception:
            return []

    try:
        result = subprocess.run(
            ["lpstat", "-e"],
            capture_output=True,
            text=True,
            check=False,
            timeout=5,
        )
        if result.returncode == 0:
            printers = [p.strip() for p in result.stdout.splitlines() if p.strip()]
            if printers:
                return printers
    except Exception:
        pass

    try:
        result = subprocess.run(
            ["lpstat", "-p"],
            capture_output=True,
            text=True,
            check=False,
            timeout=5,
        )
        if result.returncode == 0:
            printers = []
            for line in result.stdout.splitlines():
                parts = line.strip().split()
                if len(parts) >= 2 and parts[0] == "printer":
                    printers.append(parts[1])
            return printers
    except Exception:
        pass

    return []

def get_default_printer() -> Optional[str]:
    if DEFAULT_PRINTER_ENV:
        return DEFAULT_PRINTER_ENV

    if IS_WINDOWS:
        try:
            import win32print
            return win32print.GetDefaultPrinter()
        except Exception:
            pass
        printers = get_available_printers()
        return printers[0] if printers else None

    try:
        result = subprocess.run(
            ["lpstat", "-d"],
            capture_output=True,
            text=True,
            check=False,
            timeout=5,
        )
        if result.returncode == 0:
            output = result.stdout.strip()
            if ":" in output:
                target = output.split(":", 1)[1].strip()
                if target and target.lower() != "no system default destination":
                    return target
    except Exception:
        pass

    printers = get_available_printers()
    if printers:
        return printers[0]

    return None

def print_windows(
    file_path: Path,
    printer_name: str,
    copies: int = 1,
    color_mode: str = "color",
    selected_pages: Optional[Set[int]] = None,
) -> None:
    import win32con
    import win32print
    import win32ui
    from PIL import Image, ImageWin

    ext = file_path.suffix.lower()
    pages_to_print = []

    if ext == ".pdf":
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(file_path)
        for page_idx, page in enumerate(pdf, start=1):
            if selected_pages is not None and page_idx not in selected_pages:
                continue
            rendered = page.render(scale=2.5).to_pil()
            if color_mode == "bw":
                rendered = rendered.convert("L").convert("RGB")
            elif rendered.mode != "RGB":
                rendered = rendered.convert("RGB")
            pages_to_print.append(rendered)
        pdf.close()
    else:
        if selected_pages is not None and 1 not in selected_pages:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="The single image page was unselected.",
            )
        img = Image.open(file_path)
        if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
            bg = Image.new("RGB", img.size, (255, 255, 255))
            mask = img.split()[-1] if len(img.split()) == 4 else None
            bg.paste(img, mask=mask)
            img = bg
        elif img.mode != "RGB":
            img = img.convert("RGB")

        if color_mode == "bw":
            img = img.convert("L").convert("RGB")

        pages_to_print.append(img)

    if not pages_to_print:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="No pages selected for printing.",
        )

    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(printer_name)
    try:
        hDC.StartDoc(file_path.name)
        printable_width = hDC.GetDeviceCaps(win32con.HORZRES)
        printable_height = hDC.GetDeviceCaps(win32con.VERTRES)

        for _ in range(copies):
            for page_img in pages_to_print:
                hDC.StartPage()
                img_w, img_h = page_img.size
                scale = min(printable_width / img_w, printable_height / img_h)
                draw_w = int(img_w * scale)
                draw_h = int(img_h * scale)
                x = (printable_width - draw_w) // 2
                y = (printable_height - draw_h) // 2

                dib = ImageWin.Dib(page_img)
                dib.draw(hDC.GetHandleOutput(), (x, y, x + draw_w, y + draw_h))
                hDC.EndPage()

        hDC.EndDoc()
    finally:
        hDC.DeleteDC()

def print_linux_cups(
    file_path: Path,
    printer_name: str,
    copies: int = 1,
    fit_to_page: bool = True,
    color_mode: str = "color",
    selected_pages: Optional[Set[int]] = None,
) -> None:
    cmd = ["lp", "-d", printer_name, "-n", str(copies)]
    if fit_to_page:
        cmd.extend(["-o", "fit-to-page"])

    if color_mode == "bw":
        cmd.extend(["-o", "print-color-mode=monochrome", "-o", "ColorModel=Gray"])
    else:
        cmd.extend(["-o", "print-color-mode=color", "-o", "ColorModel=RGB"])

    if selected_pages is not None:
        sorted_pages = sorted(list(selected_pages))
        if sorted_pages:
            cmd.extend(["-P", ",".join(str(p) for p in sorted_pages)])

    cmd.append(str(file_path))

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=False,
            timeout=30,
        )
    except subprocess.TimeoutExpired:
        raise HTTPException(
            status_code=status.HTTP_504_GATEWAY_TIMEOUT,
            detail="Print command timed out after 30 seconds.",
        )
    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="CUPS 'lp' executable was not found.",
        )
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Execution error: {str(e)}",
        )

    if result.returncode != 0:
        err_msg = result.stderr.strip() or result.stdout.strip() or "Unknown CUPS error"
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=f"CUPS error: {err_msg}",
        )

def execute_print(
    file_path: Path,
    printer_name: Optional[str] = None,
    copies: int = 1,
    fit_to_page: bool = True,
    color_mode: str = "color",
    selected_pages: Optional[Set[int]] = None,
) -> str:
    target_printer = printer_name or get_default_printer()
    if not target_printer:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="No printer configured or detected on this host.",
        )

    try:
        target_printer = sanitize_printer_name(target_printer)
    except ValueError as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(e),
        )

    if copies < 1 or copies > 99:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Number of copies must be between 1 and 99.",
        )

    if IS_WINDOWS:
        try:
            print_windows(
                file_path=file_path,
                printer_name=target_printer,
                copies=copies,
                color_mode=color_mode,
                selected_pages=selected_pages,
            )
        except HTTPException:
            raise
        except Exception as e:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Windows print error: {str(e)}",
            )
    else:
        print_linux_cups(
            file_path=file_path,
            printer_name=target_printer,
            copies=copies,
            fit_to_page=fit_to_page,
            color_mode=color_mode,
            selected_pages=selected_pages,
        )

    return target_printer

@app.get("/health")
def health_check():
    return {"status": "ok", "platform": "windows" if IS_WINDOWS else "linux"}

@app.get("/manifest.json")
def get_manifest():
    manifest_data = {
        "name": "Web Print",
        "short_name": "Print",
        "start_url": "/",
        "display": "standalone",
        "background_color": "#F2F2F7",
        "theme_color": "#007AFF",
        "icons": [
            {
                "src": "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 512 512'%3E%3Crect width='512' height='512' rx='115' fill='%23007AFF'/%3E%3Cpath fill='%23FFFFFF' d='M384 192V112H128v80H96v176h32v48h256v-48h32V192h-32zm-224-48h192v48H160v-48zm192 240H160V304h192v80zm32-112c-8.8 0-16-7.2-16-16s7.2-16 16-16 16 7.2 16 16-7.2 16-16 16z'/%3E%3C/svg%3E",
                "sizes": "192x192 512x512",
                "type": "image/svg+xml",
            }
        ],
    }
    return JSONResponse(content=manifest_data)

@app.get("/api/printers")
def list_printers():
    printers = get_available_printers()
    default_p = get_default_printer()
    return {
        "printers": printers,
        "default": default_p,
    }

@app.post("/api/preview")
async def generate_preview(file: UploadFile = File(...)):
    if not file.filename:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="No file selected.",
        )

    file_ext = Path(file.filename).suffix.lower()
    if file_ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
            detail=f"Unsupported file extension '{file_ext}'.",
        )

    DEFAULT_TEMP_DIR.mkdir(parents=True, exist_ok=True)
    temp_path = DEFAULT_TEMP_DIR / f"preview_{uuid.uuid4()}{file_ext}"

    try:
        total_bytes = 0
        with open(temp_path, "wb") as buffer:
            while True:
                chunk = await file.read(1024 * 1024)
                if not chunk:
                    break
                total_bytes += len(chunk)
                if total_bytes > MAX_FILE_SIZE:
                    raise HTTPException(
                        status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                        detail="File exceeds maximum allowed size.",
                    )
                buffer.write(chunk)

        thumbnails = []
        if file_ext == ".pdf":
            import pypdfium2 as pdfium
            pdf = pdfium.PdfDocument(temp_path)
            for page in pdf:
                pil_img = page.render(scale=0.45).to_pil()
                if pil_img.mode != "RGB":
                    pil_img = pil_img.convert("RGB")
                buf = io.BytesIO()
                pil_img.save(buf, format="JPEG", quality=75)
                b64 = base64.b64encode(buf.getvalue()).decode("ascii")
                thumbnails.append(f"data:image/jpeg;base64,{b64}")
            pdf.close()
        else:
            from PIL import Image
            img = Image.open(temp_path)
            if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
                bg = Image.new("RGB", img.size, (255, 255, 255))
                mask = img.split()[-1] if len(img.split()) == 4 else None
                bg.paste(img, mask=mask)
                img = bg
            elif img.mode != "RGB":
                img = img.convert("RGB")
            img.thumbnail((400, 400))
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=75)
            b64 = base64.b64encode(buf.getvalue()).decode("ascii")
            thumbnails.append(f"data:image/jpeg;base64,{b64}")

        return {
            "success": True,
            "filename": file.filename,
            "count": len(thumbnails),
            "pages": thumbnails,
        }
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass
        await file.close()

@app.post("/api/print")
async def submit_print_job(
    file: UploadFile = File(...),
    printer: Optional[str] = Form(None),
    copies: int = Form(1),
    color_mode: str = Form("color"),
    pages: Optional[str] = Form(None),
):
    if not file.filename:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="No file selected.",
        )

    file_ext = Path(file.filename).suffix.lower()
    if file_ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
            detail=f"Unsupported file extension '{file_ext}'. Allowed formats: {', '.join(sorted(ALLOWED_EXTENSIONS))}",
        )

    selected_pages_set: Optional[Set[int]] = None
    if pages and pages.strip():
        try:
            parts = [p.strip() for p in pages.split(",") if p.strip()]
            selected_pages_set = {int(p) for p in parts if int(p) > 0}
            if not selected_pages_set:
                raise ValueError()
        except ValueError:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Invalid pages parameter. Must be comma-separated integers.",
            )

    DEFAULT_TEMP_DIR.mkdir(parents=True, exist_ok=True)
    job_id = str(uuid.uuid4())
    temp_file_path = DEFAULT_TEMP_DIR / f"{job_id}{file_ext}"

    try:
        total_bytes = 0
        with open(temp_file_path, "wb") as buffer:
            while True:
                chunk = await file.read(1024 * 1024)
                if not chunk:
                    break
                total_bytes += len(chunk)
                if total_bytes > MAX_FILE_SIZE:
                    raise HTTPException(
                        status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                        detail=f"File exceeds maximum allowed size of {MAX_FILE_SIZE // (1024 * 1024)} MB.",
                    )
                buffer.write(chunk)

        if total_bytes == 0:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Uploaded file is empty.",
            )

        printer_used = execute_print(
            file_path=temp_file_path,
            printer_name=printer.strip() if printer else None,
            copies=copies,
            fit_to_page=True,
            color_mode=color_mode.strip().lower(),
            selected_pages=selected_pages_set,
        )

        return {
            "success": True,
            "job_id": job_id,
            "printer": printer_used,
            "copies": copies,
            "color_mode": color_mode,
            "filename": file.filename,
            "message": f"Sent {file.filename} to {printer_used}",
        }
    finally:
        if temp_file_path.exists():
            try:
                temp_file_path.unlink()
            except OSError:
                pass
        await file.close()

HTML_CONTENT = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover">
    <title>Web Print</title>
    <link rel="manifest" href="/manifest.json">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="default">
    <meta name="apple-mobile-web-app-title" content="Web Print">
    <meta name="theme-color" content="#F2F2F7">
    <link rel="apple-touch-icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 512 512'%3E%3Crect width='512' height='512' rx='115' fill='%23007AFF'/%3E%3Cpath fill='%23FFFFFF' d='M384 192V112H128v80H96v176h32v48h256v-48h32V192h-32zm-224-48h192v48H160v-48zm192 240H160V304h192v80zm32-112c-8.8 0-16-7.2-16-16s7.2-16 16-16 16 7.2 16 16-7.2 16-16 16z'/%3E%3C/svg%3E">
    <style>
        :root {
            --bg-color: #F2F2F7;
            --card-bg: #FFFFFF;
            --text-primary: #000000;
            --text-secondary: #8E8E93;
            --accent: #007AFF;
            --accent-active: #0051A8;
            --destructive: #FF3B30;
            --success: #34C759;
            --border-color: #E5E5EA;
            --separator: #C6C6C8;
            --segmented-bg: #E3E3E8;
            --font-stack: -apple-system, BlinkMacSystemFont, "SF Pro Text", "SF Pro Display", "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        }

        @media (prefers-color-scheme: dark) {
            :root {
                --bg-color: #000000;
                --card-bg: #1C1C1E;
                --text-primary: #FFFFFF;
                --text-secondary: #8E8E93;
                --accent: #0A84FF;
                --accent-active: #0066CC;
                --destructive: #FF453A;
                --success: #32D74B;
                --border-color: #2C2C2E;
                --separator: #38383A;
                --segmented-bg: #2C2C2E;
            }
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            -webkit-tap-highlight-color: transparent;
        }

        body {
            font-family: var(--font-stack);
            background-color: var(--bg-color);
            color: var(--text-primary);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            align-items: center;
            padding: max(20px, env(safe-area-inset-top)) max(16px, env(safe-area-inset-right)) max(30px, env(safe-area-inset-bottom)) max(16px, env(safe-area-inset-left));
        }

        .container {
            width: 100%;
            max-width: 460px;
            display: flex;
            flex-direction: column;
            gap: 16px;
        }

        header {
            padding: 10px 4px 0 4px;
        }

        header h1 {
            font-size: 34px;
            font-weight: 700;
            letter-spacing: -0.5px;
        }

        header p {
            font-size: 15px;
            color: var(--text-secondary);
            margin-top: 4px;
        }

        .banner {
            display: none;
            padding: 14px 16px;
            border-radius: 14px;
            font-size: 15px;
            font-weight: 500;
            line-height: 1.4;
            align-items: center;
            gap: 10px;
            animation: fadeIn 0.25s ease-out;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(-8px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .banner.success {
            display: flex;
            background-color: rgba(52, 199, 89, 0.15);
            color: var(--success);
            border: 1px solid rgba(52, 199, 89, 0.3);
        }

        .banner.error {
            display: flex;
            background-color: rgba(255, 59, 48, 0.15);
            color: var(--destructive);
            border: 1px solid rgba(255, 59, 48, 0.3);
        }

        .banner-icon {
            flex-shrink: 0;
            width: 20px;
            height: 20px;
        }

        .banner-text {
            flex-grow: 1;
            word-break: break-word;
        }

        .card {
            background-color: var(--card-bg);
            border-radius: 14px;
            overflow: hidden;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
        }

        .upload-zone {
            padding: 28px 20px;
            text-align: center;
            cursor: pointer;
            transition: background-color 0.15s ease;
            position: relative;
        }

        .upload-zone:active {
            background-color: rgba(0, 122, 255, 0.06);
        }

        .upload-zone.dragover {
            background-color: rgba(0, 122, 255, 0.1);
        }

        .upload-icon {
            width: 48px;
            height: 48px;
            margin: 0 auto 10px;
            fill: var(--accent);
        }

        .upload-title {
            font-size: 17px;
            font-weight: 600;
            margin-bottom: 4px;
        }

        .upload-subtitle {
            font-size: 13px;
            color: var(--text-secondary);
        }

        .file-preview {
            display: none;
            padding: 14px 16px;
            align-items: center;
            gap: 12px;
        }

        .file-icon {
            width: 34px;
            height: 34px;
            fill: var(--accent);
            flex-shrink: 0;
        }

        .file-details {
            flex-grow: 1;
            min-width: 0;
        }

        .file-name {
            font-size: 16px;
            font-weight: 600;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .file-size {
            font-size: 13px;
            color: var(--text-secondary);
            margin-top: 2px;
        }

        .file-clear-btn {
            background: none;
            border: none;
            color: var(--destructive);
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            padding: 6px 10px;
        }

        .file-clear-btn:active {
            opacity: 0.6;
        }

        .preview-section {
            display: flex;
            flex-direction: column;
            gap: 10px;
        }

        .preview-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 4px;
        }

        .preview-title {
            font-size: 14px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: var(--text-secondary);
        }

        .preview-actions {
            display: flex;
            gap: 8px;
        }

        .pill-btn {
            background: none;
            border: none;
            font-family: inherit;
            font-size: 14px;
            font-weight: 500;
            color: var(--accent);
            cursor: pointer;
            padding: 4px 6px;
        }

        .pill-btn:active {
            opacity: 0.6;
        }

        .preview-loader {
            display: none;
            align-items: center;
            justify-content: center;
            gap: 10px;
            padding: 24px;
            color: var(--text-secondary);
            font-size: 14px;
        }

        .mini-spinner {
            width: 18px;
            height: 18px;
            border: 2px solid rgba(0, 122, 255, 0.2);
            border-top-color: var(--accent);
            border-radius: 50%;
            animation: spin 0.75s linear infinite;
        }

        .pages-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(95px, 1fr));
            gap: 10px;
        }

        .page-card {
            position: relative;
            background: var(--card-bg);
            border-radius: 10px;
            overflow: hidden;
            border: 2px solid transparent;
            cursor: pointer;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
            transition: transform 0.1s ease, border-color 0.15s ease, opacity 0.15s ease;
            user-select: none;
        }

        .page-card:active {
            transform: scale(0.97);
        }

        .page-card.selected {
            border-color: var(--accent);
        }

        .page-card.deselected {
            opacity: 0.4;
            filter: grayscale(80%);
        }

        .page-thumb {
            width: 100%;
            aspect-ratio: 1 / 1.414;
            object-fit: cover;
            display: block;
            background-color: #FAFAFA;
        }

        .page-badge {
            position: absolute;
            bottom: 4px;
            left: 4px;
            background: rgba(0, 0, 0, 0.65);
            color: #FFFFFF;
            font-size: 11px;
            font-weight: 600;
            padding: 2px 6px;
            border-radius: 6px;
            backdrop-filter: blur(4px);
        }

        .check-badge {
            position: absolute;
            top: 6px;
            right: 6px;
            width: 22px;
            height: 22px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.15s ease;
        }

        .page-card.selected .check-badge {
            background-color: var(--accent);
            color: #FFFFFF;
        }

        .page-card.deselected .check-badge {
            background-color: rgba(255, 255, 255, 0.85);
            border: 2px solid var(--text-secondary);
        }

        .check-badge svg {
            width: 13px;
            height: 13px;
            fill: currentColor;
        }

        .page-card.deselected .check-badge svg {
            display: none;
        }

        .settings-list {
            display: flex;
            flex-direction: column;
        }

        .settings-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 13px 16px;
            position: relative;
        }

        .settings-row:not(:last-child)::after {
            content: "";
            position: absolute;
            left: 16px;
            right: 0;
            bottom: 0;
            height: 0.5px;
            background-color: var(--separator);
        }

        .settings-label {
            font-size: 17px;
            font-weight: 400;
        }

        .printer-select-wrapper {
            position: relative;
            display: flex;
            align-items: center;
            max-width: 65%;
        }

        .printer-select {
            appearance: none;
            -webkit-appearance: none;
            background: transparent;
            border: none;
            font-family: inherit;
            font-size: 17px;
            color: var(--accent);
            text-align: right;
            padding-right: 18px;
            width: 100%;
            cursor: pointer;
            outline: none;
            text-overflow: ellipsis;
            white-space: nowrap;
            overflow: hidden;
        }

        .select-chevron {
            position: absolute;
            right: 0;
            pointer-events: none;
            width: 12px;
            height: 12px;
            fill: var(--text-secondary);
        }

        .segmented-control {
            display: flex;
            background-color: var(--segmented-bg);
            border-radius: 9px;
            padding: 2px;
            position: relative;
        }

        .segment-btn {
            background: transparent;
            border: none;
            font-family: inherit;
            font-size: 14px;
            font-weight: 500;
            padding: 6px 14px;
            border-radius: 7px;
            color: var(--text-primary);
            cursor: pointer;
            transition: background-color 0.15s ease, box-shadow 0.15s ease;
        }

        .segment-btn.active {
            background-color: var(--card-bg);
            font-weight: 600;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.12);
        }

        .stepper {
            display: flex;
            align-items: center;
            background-color: var(--bg-color);
            border-radius: 9px;
            padding: 3px;
            gap: 2px;
        }

        .stepper-btn {
            background: var(--card-bg);
            border: none;
            width: 32px;
            height: 30px;
            border-radius: 7px;
            font-size: 18px;
            font-weight: 600;
            color: var(--text-primary);
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            box-shadow: 0 1px 2px rgba(0, 0, 0, 0.08);
            transition: opacity 0.1s ease;
        }

        .stepper-btn:active {
            opacity: 0.5;
        }

        .stepper-value {
            min-width: 36px;
            text-align: center;
            font-size: 16px;
            font-weight: 600;
        }

        .action-container {
            margin-top: 4px;
        }

        .print-btn {
            width: 100%;
            height: 52px;
            border-radius: 14px;
            border: none;
            background-color: var(--accent);
            color: #FFFFFF;
            font-family: inherit;
            font-size: 17px;
            font-weight: 600;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            cursor: pointer;
            transition: background-color 0.15s ease, transform 0.1s ease;
        }

        .print-btn:active:not(:disabled) {
            background-color: var(--accent-active);
            transform: scale(0.98);
        }

        .print-btn:disabled {
            opacity: 0.45;
            cursor: not-allowed;
        }

        .spinner {
            display: none;
            width: 20px;
            height: 20px;
            border: 2.5px solid rgba(255, 255, 255, 0.3);
            border-top-color: #FFFFFF;
            border-radius: 50%;
            animation: spin 0.75s linear infinite;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        input[type="file"] {
            display: none;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>Print</h1>
            <p>Self-hosted home print service</p>
        </header>

        <div id="banner" class="banner">
            <svg id="bannerIcon" class="banner-icon" viewBox="0 0 20 20" fill="currentColor"></svg>
            <span id="bannerText" class="banner-text"></span>
        </div>

        <div class="card">
            <div id="uploadZone" class="upload-zone">
                <svg class="upload-icon" viewBox="0 0 24 24">
                    <path d="M19.35 10.04C18.67 6.59 15.64 4 12 4 9.11 4 6.6 5.64 5.35 8.04 2.34 8.36 0 10.91 0 14c0 3.31 2.69 6 6 6h13c2.76 0 5-2.24 5-5 0-2.64-2.05-4.78-4.65-4.96zM14 13v4h-4v-4H7l5-5 5 5h-3z"/>
                </svg>
                <div class="upload-title">Select Document or Photo</div>
                <div class="upload-subtitle">PDF, PNG, JPG (up to 25 MB)</div>
            </div>

            <div id="filePreview" class="file-preview">
                <svg class="file-icon" viewBox="0 0 24 24">
                    <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                </svg>
                <div class="file-details">
                    <div id="fileName" class="file-name"></div>
                    <div id="fileSize" class="file-size"></div>
                </div>
                <button type="button" id="clearFileBtn" class="file-clear-btn">Remove</button>
            </div>
        </div>

        <input type="file" id="fileInput" accept=".pdf,image/png,image/jpeg,image/jpg">

        <div id="previewSection" class="preview-section" style="display: none;">
            <div class="preview-header">
                <span id="pagesCountLabel" class="preview-title">Pages (0 selected)</span>
                <div class="preview-actions">
                    <button type="button" id="selectAllBtn" class="pill-btn">Select All</button>
                    <button type="button" id="deselectAllBtn" class="pill-btn">Deselect All</button>
                </div>
            </div>

            <div id="previewLoader" class="card preview-loader">
                <div class="mini-spinner"></div>
                <span>Generating page previews...</span>
            </div>

            <div id="pagesGrid" class="pages-grid"></div>
        </div>

        <div class="card settings-list">
            <div class="settings-row">
                <span class="settings-label">Printer</span>
                <div class="printer-select-wrapper">
                    <select id="printerSelect" class="printer-select">
                        <option value="">Detecting...</option>
                    </select>
                    <svg class="select-chevron" viewBox="0 0 24 24">
                        <path d="M8.59 16.59L13.17 12 8.59 7.41 10 6l6 6-6 6-1.41-1.41z"/>
                    </svg>
                </div>
            </div>

            <div class="settings-row">
                <span class="settings-label">Color</span>
                <div class="segmented-control">
                    <button type="button" id="btnColor" class="segment-btn active" data-mode="color">Color</button>
                    <button type="button" id="btnBw" class="segment-btn" data-mode="bw">B&W</button>
                </div>
            </div>

            <div class="settings-row">
                <span class="settings-label">Copies</span>
                <div class="stepper">
                    <button type="button" id="stepDown" class="stepper-btn">&minus;</button>
                    <span id="copiesVal" class="stepper-value">1</span>
                    <button type="button" id="stepUp" class="stepper-btn">&plus;</button>
                </div>
            </div>
        </div>

        <div class="action-container">
            <button type="button" id="printBtn" class="print-btn" disabled>
                <span class="spinner" id="spinner"></span>
                <span id="btnText">Print Document</span>
            </button>
        </div>
    </div>

    <script>
        const fileInput = document.getElementById('fileInput');
        const uploadZone = document.getElementById('uploadZone');
        const filePreview = document.getElementById('filePreview');
        const fileNameEl = document.getElementById('fileName');
        const fileSizeEl = document.getElementById('fileSize');
        const clearFileBtn = document.getElementById('clearFileBtn');
        const previewSection = document.getElementById('previewSection');
        const previewLoader = document.getElementById('previewLoader');
        const pagesGrid = document.getElementById('pagesGrid');
        const pagesCountLabel = document.getElementById('pagesCountLabel');
        const selectAllBtn = document.getElementById('selectAllBtn');
        const deselectAllBtn = document.getElementById('deselectAllBtn');
        const printerSelect = document.getElementById('printerSelect');
        const btnColor = document.getElementById('btnColor');
        const btnBw = document.getElementById('btnBw');
        const stepDown = document.getElementById('stepDown');
        const stepUp = document.getElementById('stepUp');
        const copiesVal = document.getElementById('copiesVal');
        const printBtn = document.getElementById('printBtn');
        const btnText = document.getElementById('btnText');
        const spinner = document.getElementById('spinner');
        const banner = document.getElementById('banner');
        const bannerIcon = document.getElementById('bannerIcon');
        const bannerText = document.getElementById('bannerText');

        let selectedFile = null;
        let copiesCount = 1;
        let colorMode = 'color';
        let bannerTimeout = null;
        let totalPages = 0;
        let selectedPageIndices = new Set();

        const checkSvg = '<path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clip-rule="evenodd"/>';
        const errorSvg = '<path fill-rule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7 4a1 1 0 11-2 0 1 1 0 012 0zm-1-9a1 1 0 00-1 1v4a1 1 0 102 0V6a1 1 0 00-1-1z" clip-rule="evenodd"/>';
        const checkMiniSvg = '<svg viewBox="0 0 20 20"><path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd"/></svg>';

        function showBanner(type, message) {
            if (bannerTimeout) {
                clearTimeout(bannerTimeout);
            }
            banner.className = 'banner ' + type;
            bannerIcon.innerHTML = type === 'success' ? checkSvg : errorSvg;
            bannerText.textContent = message;
            bannerTimeout = setTimeout(() => {
                banner.className = 'banner';
            }, 6000);
        }

        banner.addEventListener('click', () => {
            banner.className = 'banner';
        });

        function formatBytes(bytes) {
            if (bytes === 0) return '0 B';
            const k = 1024;
            const sizes = ['B', 'KB', 'MB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
        }

        function updateSelectionUI() {
            pagesCountLabel.textContent = `Pages (${selectedPageIndices.size} of ${totalPages} selected)`;
            if (!selectedFile || selectedPageIndices.size === 0) {
                printBtn.disabled = true;
                btnText.textContent = selectedPageIndices.size === 0 && selectedFile ? 'Select At Least 1 Page' : 'Print Document';
            } else {
                printBtn.disabled = false;
                btnText.textContent = `Print ${selectedPageIndices.size} Page${selectedPageIndices.size > 1 ? 's' : ''}`;
            }
        }

        function renderPageCards(thumbnails) {
            pagesGrid.innerHTML = '';
            totalPages = thumbnails.length;
            selectedPageIndices = new Set();

            thumbnails.forEach((src, idx) => {
                const pageNumber = idx + 1;
                selectedPageIndices.add(pageNumber);

                const card = document.createElement('div');
                card.className = 'page-card selected';
                card.dataset.page = pageNumber;

                const img = document.createElement('img');
                img.className = 'page-thumb';
                img.src = src;
                img.alt = `Page ${pageNumber}`;

                const badge = document.createElement('div');
                badge.className = 'page-badge';
                badge.textContent = `P. ${pageNumber}`;

                const check = document.createElement('div');
                check.className = 'check-badge';
                check.innerHTML = checkMiniSvg;

                card.appendChild(img);
                card.appendChild(badge);
                card.appendChild(check);

                card.addEventListener('click', () => {
                    if (selectedPageIndices.has(pageNumber)) {
                        selectedPageIndices.delete(pageNumber);
                        card.className = 'page-card deselected';
                    } else {
                        selectedPageIndices.add(pageNumber);
                        card.className = 'page-card selected';
                    }
                    updateSelectionUI();
                });

                pagesGrid.appendChild(card);
            });

            previewLoader.style.display = 'none';
            pagesGrid.style.display = 'grid';
            updateSelectionUI();
        }

        async function loadPreview(file) {
            previewSection.style.display = 'flex';
            previewLoader.style.display = 'flex';
            pagesGrid.style.display = 'none';
            pagesGrid.innerHTML = '';

            const formData = new FormData();
            formData.append('file', file);

            try {
                const res = await fetch('/api/preview', {
                    method: 'POST',
                    body: formData,
                });
                const data = await res.json();
                if (!res.ok) {
                    throw new Error(data.detail || 'Failed to render preview');
                }
                renderPageCards(data.pages || []);
            } catch (err) {
                previewLoader.style.display = 'none';
                showBanner('error', 'Could not generate preview: ' + err.message);
            }
        }

        function updateFileUI() {
            if (selectedFile) {
                uploadZone.style.display = 'none';
                filePreview.style.display = 'flex';
                fileNameEl.textContent = selectedFile.name;
                fileSizeEl.textContent = formatBytes(selectedFile.size);
                loadPreview(selectedFile);
            } else {
                uploadZone.style.display = 'block';
                filePreview.style.display = 'none';
                previewSection.style.display = 'none';
                pagesGrid.innerHTML = '';
                selectedPageIndices.clear();
                totalPages = 0;
                fileInput.value = '';
                printBtn.disabled = true;
                btnText.textContent = 'Print Document';
            }
        }

        selectAllBtn.addEventListener('click', () => {
            const cards = pagesGrid.querySelectorAll('.page-card');
            cards.forEach((card) => {
                const pageNumber = parseInt(card.dataset.page, 10);
                selectedPageIndices.add(pageNumber);
                card.className = 'page-card selected';
            });
            updateSelectionUI();
        });

        deselectAllBtn.addEventListener('click', () => {
            const cards = pagesGrid.querySelectorAll('.page-card');
            cards.forEach((card) => {
                card.className = 'page-card deselected';
            });
            selectedPageIndices.clear();
            updateSelectionUI();
        });

        btnColor.addEventListener('click', () => {
            colorMode = 'color';
            btnColor.className = 'segment-btn active';
            btnBw.className = 'segment-btn';
        });

        btnBw.addEventListener('click', () => {
            colorMode = 'bw';
            btnBw.className = 'segment-btn active';
            btnColor.className = 'segment-btn';
        });

        uploadZone.addEventListener('click', () => fileInput.click());

        fileInput.addEventListener('change', (e) => {
            const files = e.target.files;
            if (files && files.length > 0) {
                selectedFile = files[0];
                updateFileUI();
            }
        });

        clearFileBtn.addEventListener('click', () => {
            selectedFile = null;
            updateFileUI();
        });

        uploadZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadZone.classList.add('dragover');
        });

        uploadZone.addEventListener('dragleave', () => {
            uploadZone.classList.remove('dragover');
        });

        uploadZone.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadZone.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files && files.length > 0) {
                selectedFile = files[0];
                updateFileUI();
            }
        });

        stepDown.addEventListener('click', () => {
            if (copiesCount > 1) {
                copiesCount--;
                copiesVal.textContent = copiesCount;
            }
        });

        stepUp.addEventListener('click', () => {
            if (copiesCount < 99) {
                copiesCount++;
                copiesVal.textContent = copiesCount;
            }
        });

        async function fetchPrinters() {
            try {
                const res = await fetch('/api/printers');
                if (!res.ok) throw new Error('Could not retrieve printer list');
                const data = await res.json();
                printerSelect.innerHTML = '';

                if (data.printers && data.printers.length > 0) {
                    data.printers.forEach((p) => {
                        const opt = document.createElement('option');
                        opt.value = p;
                        opt.textContent = p + (p === data.default ? ' (Default)' : '');
                        if (p === data.default) opt.selected = true;
                        printerSelect.appendChild(opt);
                    });
                } else if (data.default) {
                    const opt = document.createElement('option');
                    opt.value = data.default;
                    opt.textContent = data.default + ' (Default)';
                    opt.selected = true;
                    printerSelect.appendChild(opt);
                } else {
                    const opt = document.createElement('option');
                    opt.value = '';
                    opt.textContent = 'Auto Default';
                    printerSelect.appendChild(opt);
                }
            } catch (err) {
                printerSelect.innerHTML = '<option value="">Auto Default</option>';
            }
        }

        printBtn.addEventListener('click', async () => {
            if (!selectedFile || selectedPageIndices.size === 0) return;

            printBtn.disabled = true;
            spinner.style.display = 'inline-block';
            btnText.textContent = 'Sending to Printer...';

            const formData = new FormData();
            formData.append('file', selectedFile);
            if (printerSelect.value) {
                formData.append('printer', printerSelect.value);
            }
            formData.append('copies', copiesCount);
            formData.append('color_mode', colorMode);

            const sortedPages = Array.from(selectedPageIndices).sort((a, b) => a - b);
            formData.append('pages', sortedPages.join(','));

            try {
                const res = await fetch('/api/print', {
                    method: 'POST',
                    body: formData,
                });
                const result = await res.json();

                if (!res.ok) {
                    throw new Error(result.detail || 'Printing failed');
                }

                showBanner('success', result.message || 'Job submitted successfully.');
                selectedFile = null;
                updateFileUI();
            } catch (err) {
                showBanner('error', err.message || 'Error occurred while printing.');
                printBtn.disabled = false;
            } finally {
                spinner.style.display = 'none';
                if (selectedFile && selectedPageIndices.size > 0) {
                    printBtn.disabled = false;
                    btnText.textContent = `Print ${selectedPageIndices.size} Page${selectedPageIndices.size > 1 ? 's' : ''}`;
                } else {
                    btnText.textContent = 'Print Document';
                }
            }
        });

        fetchPrinters();
    </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def index_page():
    return HTMLResponse(content=HTML_CONTENT)

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 5000))
    uvicorn.run("app:app", host="0.0.0.0", port=port, reload=False)
