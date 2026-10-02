import asyncio
import logging
import os
import re
import shutil
import subprocess
import tempfile
import socket
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Optional

import img2pdf
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from telegram import InputFile, Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    ConversationHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("pdf-converter-bot")

WAIT_FILE, WAIT_NAME = range(2)

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp"}
TEXT_EXTS = {".txt", ".md", ".log", ".csv"}
OFFICE_EXTS = {
    ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx",
    ".odt", ".ods", ".odp", ".rtf"
}

MAX_BOT_DOWNLOAD = 20 * 1024 * 1024  # ~20MB Telegram Bot API standard getFile limit


def sanitize_filename(name: str, max_len: int = 64) -> str:
    name = (name or "").strip()
    if not name:
        return "converted"
    name = re.sub(r"\.pdf$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"[^A-Za-z0-9 _-]+", "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return (name or "converted")[:max_len]


def safe_basename(filename: str) -> str:
    return Path(filename).name or "file"


def find_soffice() -> Optional[str]:
    p = shutil.which("soffice") or shutil.which("libreoffice")
    if p:
        return p

    mac_candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        str(Path.home() / "Applications/LibreOffice.app/Contents/MacOS/soffice"),
    ]
    for c in mac_candidates:
        if Path(c).exists():
            return c
    return None


def convert_text_to_pdf(input_path: Path, pdf_path: Path) -> None:
    text = input_path.read_text(errors="ignore")

    c = canvas.Canvas(str(pdf_path), pagesize=A4)
    width, height = A4
    margin = 50
    y = height - margin
    line_height = 14
    max_chars_per_line = 95

    for raw_line in text.splitlines():
        line = raw_line.rstrip("\n")
        while len(line) > max_chars_per_line:
            if y < margin:
                c.showPage()
                y = height - margin
            c.drawString(margin, y, line[:max_chars_per_line])
            line = line[max_chars_per_line:]
            y -= line_height

        if y < margin:
            c.showPage()
            y = height - margin
        c.drawString(margin, y, line)
        y -= line_height

    c.save()


def convert_image_to_pdf(input_path: Path, pdf_path: Path) -> None:
    with open(input_path, "rb") as f_in:
        img_bytes = f_in.read()
    pdf_bytes = img2pdf.convert(img_bytes)
    pdf_path.write_bytes(pdf_bytes)


def convert_office_to_pdf(input_path: Path, out_dir: Path, timeout_sec: int = 90) -> Path:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice not found. Install it or ensure 'soffice' is in PATH.")

    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--norestore",
        "--convert-to",
        "pdf",
        str(input_path),
        "--outdir",
        str(out_dir),
    ]

    logger.info("Running: %s", " ".join(cmd))
    try:
        res = subprocess.run(
            cmd,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout_sec,
        )
    except subprocess.CalledProcessError as e:
        stderr = (e.stderr or b"").decode(errors="ignore")
        stdout = (e.stdout or b"").decode(errors="ignore")
        raise RuntimeError(
            "LibreOffice conversion failed.\n"
            f"STDOUT:\n{stdout}\n\nSTDERR:\n{stderr}"
        ) from e

    expected = out_dir / (input_path.stem + ".pdf")
    if expected.exists():
        return expected

    pdfs = sorted(out_dir.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not pdfs:
        stdout = (res.stdout or b"").decode(errors="ignore")
        stderr = (res.stderr or b"").decode(errors="ignore")
        raise RuntimeError(
            "LibreOffice finished but no PDF was created.\n"
            f"STDOUT:\n{stdout}\n\nSTDERR:\n{stderr}"
        )
    return pdfs[0]


def convert_to_pdf(input_path: Path, work_dir: Path) -> Path:
    ext = input_path.suffix.lower()
    pdf_path = work_dir / "output.pdf"

    if ext == ".pdf":
        shutil.copyfile(input_path, pdf_path)
        return pdf_path

    if ext in IMAGE_EXTS:
        convert_image_to_pdf(input_path, pdf_path)
        return pdf_path

    if ext in TEXT_EXTS:
        convert_text_to_pdf(input_path, pdf_path)
        return pdf_path

    if ext in OFFICE_EXTS:
        out_pdf = convert_office_to_pdf(input_path, work_dir)
        shutil.copyfile(out_pdf, pdf_path)
        return pdf_path

    raise RuntimeError(f"Unsupported file type: {ext} (send an image, text, PDF, or Office file).")


def cleanup_user_state(context: ContextTypes.DEFAULT_TYPE) -> None:
    work_dir_s = context.user_data.get("work_dir")
    context.user_data.clear()
    if work_dir_s:
        shutil.rmtree(work_dir_s, ignore_errors=True)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    cleanup_user_state(context)
    await update.message.reply_text(
        "Send me a file (document/photo). I’ll convert it to PDF.\n"
        "Then I’ll ask you what filename you want."
    )
    return WAIT_FILE


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    cleanup_user_state(context)
    await update.message.reply_text("Cancelled. Send another file anytime.")
    return WAIT_FILE


async def receive_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    msg = update.message

    file_id = None
    original_name = None
    file_size = None

    if msg.document:
        doc = msg.document
        file_id = doc.file_id
        original_name = doc.file_name or "file"
        file_size = doc.file_size
    elif msg.photo:
        photo = msg.photo[-1]
        file_id = photo.file_id
        original_name = "photo.jpg"
        file_size = photo.file_size
    else:
        await msg.reply_text("Please send a document or a photo.")
        return WAIT_FILE

    if file_size and file_size > MAX_BOT_DOWNLOAD:
        await msg.reply_text(
            "That file looks bigger than 20MB.\n"
            "Bots can only download up to ~20MB with the standard Bot API.\n"
            "Please send a smaller file."
        )
        return WAIT_FILE

    cleanup_user_state(context)

    work_dir = Path(tempfile.mkdtemp(prefix="tg_pdf_"))
    input_path = work_dir / safe_basename(original_name)

    try:
        tg_file = await context.bot.get_file(file_id)
        await tg_file.download_to_drive(custom_path=input_path)
    except Exception as e:
        shutil.rmtree(work_dir, ignore_errors=True)
        await msg.reply_text(f"Failed to download your file: {e}")
        return WAIT_FILE

    context.user_data["work_dir"] = str(work_dir)
    context.user_data["input_path"] = str(input_path)

    suggested = sanitize_filename(Path(input_path.name).stem)
    await msg.reply_text(
        "Got it ✅\n"
        "Keep the original filename or change it?\n"
        f"Reply KEEP to use {suggested}.pdf, or CHANGE to choose a new name."
    )
    return WAIT_NAME


async def receive_name_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    msg = update.message
    choice = (msg.text or "").strip().casefold()

    if choice in {"keep", "same", "original"}:
        return await keep_original_name_and_convert(update, context)
    if choice in {"change", "rename", "new"}:
        await msg.reply_text("Send the PDF filename you want (without .pdf).")
        return WAIT_NAME

    return await receive_name_and_convert(update, context)


async def keep_original_name_and_convert(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    msg = update.message
    work_dir_s = context.user_data.get("work_dir")
    input_path_s = context.user_data.get("input_path")
    if not work_dir_s or not input_path_s:
        await msg.reply_text("I lost the file context. Send the file again.")
        cleanup_user_state(context)
        return WAIT_FILE

    work_dir = Path(work_dir_s)
    input_path = Path(input_path_s)
    if not work_dir.exists() or not input_path.exists():
        await msg.reply_text("I lost the file context. Send the file again.")
        cleanup_user_state(context)
        return WAIT_FILE

    await msg.reply_text("Converting… ⏳")
    try:
        pdf_path = await asyncio.to_thread(convert_to_pdf, input_path, work_dir)
        out_name = f"{sanitize_filename(input_path.stem)}.pdf"
        with open(pdf_path, "rb") as f:
            await msg.reply_document(document=InputFile(f, filename=out_name))
        await msg.reply_text("Done ✅ Send another file anytime.")
    except subprocess.TimeoutExpired:
        await msg.reply_text("Conversion timed out. Try a smaller/simple file.")
    except Exception as e:
        await msg.reply_text(f"Conversion failed: {e}")
    finally:
        cleanup_user_state(context)

    return WAIT_FILE


async def receive_name_and_convert(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    msg = update.message
    desired = sanitize_filename(msg.text or "")

    work_dir_s = context.user_data.get("work_dir")
    input_path_s = context.user_data.get("input_path")
    if not work_dir_s or not input_path_s:
        await msg.reply_text("I lost the file context. Send the file again.")
        cleanup_user_state(context)
        return WAIT_FILE

    work_dir = Path(work_dir_s)
    input_path = Path(input_path_s)

    if not work_dir.exists() or not input_path.exists():
        await msg.reply_text("I lost the file context. Send the file again.")
        cleanup_user_state(context)
        return WAIT_FILE

    await msg.reply_text("Converting… ⏳")

    try:
        pdf_path = await asyncio.to_thread(convert_to_pdf, input_path, work_dir)

        out_name = f"{desired}.pdf"
        with open(pdf_path, "rb") as f:
            await msg.reply_document(document=InputFile(f, filename=out_name))

        await msg.reply_text("Done ✅ Send another file anytime.")
    except subprocess.TimeoutExpired:
        await msg.reply_text("Conversion timed out. Try a smaller/simple file.")
    except Exception as e:
        await msg.reply_text(f"Conversion failed: {e}")
    finally:
        cleanup_user_state(context)

    return WAIT_FILE


def _tcp_port_open(host: str, port: int, timeout: float = 0.6) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        if self.path != "/healthz":
            self.send_error(404)
            return
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"ok")

    def log_message(self, format: str, *args: object) -> None:
        return


def main() -> None:
    token = os.getenv("BOT_TOKEN")
    if not token:
        raise RuntimeError("Set BOT_TOKEN environment variable.")

    local_api = (os.getenv("LOCAL_BOT_API") or "").strip().rstrip("/")

    use_local = False
    if local_api:
        m = re.match(r"^https?://([^:/\s]+):(\d{2,5})$", local_api)
        if m:
            host = m.group(1)
            port = int(m.group(2))
            use_local = _tcp_port_open(host, port)
            if not use_local:
                logger.error("LOCAL_BOT_API not reachable at %s (port closed). Using Telegram API.", local_api)
        else:
            logger.error("LOCAL_BOT_API invalid: %r. Using Telegram API.", local_api)

    builder = ApplicationBuilder().token(token)

    if use_local:
        builder = builder.base_url(f"{local_api}/bot").base_file_url(f"{local_api}/file/bot")
        logger.info("Bot running (LOCAL_BOT_API=%s)...", local_api)
    else:
        logger.info("Bot running (Telegram API)...")

    app = builder.build()

    file_filter = filters.Document.ALL | filters.PHOTO

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(file_filter, receive_file),
        ],
        states={
            WAIT_FILE: [MessageHandler(file_filter, receive_file)],
            WAIT_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_name_choice)],
        },
        fallbacks=[
            CommandHandler("cancel", cancel),
            CommandHandler("start", start),
        ],
        allow_reentry=True,
    )

    app.add_handler(conv)
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(CommandHandler("start", start))

    port = int(os.getenv("PORT", "8080"))
    health_server = ThreadingHTTPServer(("0.0.0.0", port), HealthHandler)
    threading.Thread(target=health_server.serve_forever, daemon=True).start()
    logger.info("Health endpoint listening on port %s", port)

    app.run_polling()


if __name__ == "__main__":
    main()
