import asyncio
import logging
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional

import img2pdf
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

from telegram import Update, InputFile
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ConversationHandler,
    ContextTypes,
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
# "Office-like" formats LibreOffice usually handles well
OFFICE_EXTS = {
    ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx",
    ".odt", ".ods", ".odp", ".rtf"
}

def sanitize_filename(name: str, max_len: int = 64) -> str:
    name = name.strip()
    if not name:
        return "converted"
    # Remove extension if user typed ".pdf"
    name = re.sub(r"\.pdf$", "", name, flags=re.IGNORECASE)
    # Keep letters/numbers/space/_/-
    name = re.sub(r"[^A-Za-z0-9 _-]+", "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    if not name:
        name = "converted"
    return name[:max_len]

def find_soffice() -> Optional[str]:
    """
    Try to find LibreOffice 'soffice' executable.
    Works on Linux (soffice in PATH) and macOS default app install path.
    """
    # 1) PATH
    p = shutil.which("soffice") or shutil.which("libreoffice")
    if p:
        return p

    # 2) Common macOS paths
    mac_candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        str(Path.home() / "Applications/LibreOffice.app/Contents/MacOS/soffice"),
    ]
    for c in mac_candidates:
        if Path(c).exists():
            return c

    return None

def convert_text_to_pdf(input_path: Path, pdf_path: Path) -> None:
    # Simple text -> PDF (no fancy markdown rendering)
    text = input_path.read_text(errors="ignore")

    c = canvas.Canvas(str(pdf_path), pagesize=A4)
    width, height = A4
    margin = 50
    y = height - margin
    line_height = 14

    # Very simple wrapping
    max_chars_per_line = 95

    for raw_line in text.splitlines():
        line = raw_line.rstrip("\n")
        while len(line) > max_chars_per_line:
            c.drawString(margin, y, line[:max_chars_per_line])
            line = line[max_chars_per_line:]
            y -= line_height
            if y < margin:
                c.showPage()
                y = height - margin
        c.drawString(margin, y, line)
        y -= line_height
        if y < margin:
            c.showPage()
            y = height - margin

    c.save()

def convert_image_to_pdf(input_path: Path, pdf_path: Path) -> None:
    # img2pdf expects bytes
    with open(input_path, "rb") as f_in:
        img_bytes = f_in.read()
    pdf_bytes = img2pdf.convert(img_bytes)
    pdf_path.write_bytes(pdf_bytes)

def convert_office_to_pdf(input_path: Path, out_dir: Path, timeout_sec: int = 90) -> Path:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice not found. Install LibreOffice, or make sure 'soffice' is available."
        )

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

    # Run conversion
    logger.info("Running: %s", " ".join(cmd))
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout_sec)

    # LibreOffice usually outputs: <same_basename>.pdf
    expected = out_dir / (input_path.stem + ".pdf")
    if expected.exists():
        return expected

    # Fallback: pick newest PDF in out_dir
    pdfs = sorted(out_dir.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not pdfs:
        raise RuntimeError("LibreOffice conversion finished but no PDF was created.")
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

    # Try LibreOffice for office formats (and other formats it can handle)
    if ext in OFFICE_EXTS or True:
        out_pdf = convert_office_to_pdf(input_path, work_dir)
        # Standardize to output.pdf
        shutil.copyfile(out_pdf, pdf_path)
        return pdf_path

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text(
        "Send me a file (document/photo). I’ll convert it to PDF.\n"
        "Then I’ll ask you what filename you want."
    )
    return WAIT_FILE

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text("Cancelled. Send /start to begin again.")
    return ConversationHandler.END

async def receive_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    msg = update.message

    # Telegram Bot API getFile download limit is 20MB (standard bot API)
    # We'll check known sizes and warn early.
    file_id = None
    original_name = None
    file_size = None

    if msg.document:
        doc = msg.document
        file_id = doc.file_id
        original_name = doc.file_name or "file"
        file_size = doc.file_size
    elif msg.photo:
        # Take largest photo size
        photo = msg.photo[-1]
        file_id = photo.file_id
        original_name = "photo.jpg"
        file_size = photo.file_size
    else:
        await msg.reply_text("Please send a document or a photo.")
        return WAIT_FILE

    if file_size and file_size > 20 * 1024 * 1024:
        await msg.reply_text(
            "That file looks bigger than 20MB.\n"
            "With the standard Telegram Bot API, bots can only download up to 20MB.\n"
            "Please send a smaller file."
        )
        return WAIT_FILE

    # Prepare workspace for this user
    work_dir = Path(tempfile.mkdtemp(prefix="tg_pdf_"))
    input_path = work_dir / original_name

    try:
        tg_file = await context.bot.get_file(file_id)
        # download_to_drive(custom_path=...) is the modern PTB method  [oai_citation:3‡docs.python-telegram-bot.org](https://docs.python-telegram-bot.org/en/v22.1/telegram.file.html)
        await tg_file.download_to_drive(custom_path=input_path)
    except Exception as e:
        shutil.rmtree(work_dir, ignore_errors=True)
        await msg.reply_text(f"Failed to download your file: {e}")
        return WAIT_FILE

    # Store paths for next step
    context.user_data["work_dir"] = str(work_dir)
    context.user_data["input_path"] = str(input_path)
    suggested = sanitize_filename(Path(original_name).stem)

    await msg.reply_text(
        f"Got it ✅\nNow send the PDF filename you want (without .pdf).\n"
        f"Example: {suggested}"
    )
    return WAIT_NAME

async def receive_name_and_convert(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    msg = update.message
    desired = sanitize_filename(msg.text or "")

    work_dir = Path(context.user_data.get("work_dir", ""))
    input_path = Path(context.user_data.get("input_path", ""))

    if not work_dir.exists() or not input_path.exists():
        await msg.reply_text("I lost the file context. Send /start and upload again.")
        context.user_data.clear()
        return ConversationHandler.END

    await msg.reply_text("Converting… ⏳")

    try:
        pdf_path = convert_to_pdf(input_path, work_dir)

        out_name = f"{desired}.pdf"
        with open(pdf_path, "rb") as f:
            await msg.reply_document(document=InputFile(f, filename=out_name))
            # InputFile supports custom filename  [oai_citation:4‡docs.python-telegram-bot.org](https://docs.python-telegram-bot.org/en/v21.7/telegram.inputfile.html?utm_source=chatgpt.com)

        await msg.reply_text("Done ✅ Send another file anytime.")
    except subprocess.TimeoutExpired:
        await msg.reply_text("Conversion timed out. Try a smaller/simple file.")
    except subprocess.CalledProcessError as e:
        await msg.reply_text(
            "LibreOffice conversion failed.\n"
            f"Error: {e}"
        )
    except Exception as e:
        await msg.reply_text(f"Conversion failed: {e}")
    finally:
        # Cleanup
        context.user_data.clear()
        shutil.rmtree(work_dir, ignore_errors=True)

    return ConversationHandler.END

def main() -> None:
    token = os.getenv("BOT_TOKEN")
    if not token:
        raise RuntimeError("Set BOT_TOKEN environment variable.")

    app = ApplicationBuilder().token(token).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            WAIT_FILE: [MessageHandler(filters.Document.ALL | filters.PHOTO, receive_file)],
            WAIT_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_name_and_convert)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )

    app.add_handler(conv)
    app.add_handler(CommandHandler("cancel", cancel))

    logger.info("Bot running...")
    app.run_polling()

if __name__ == "__main__":
    main()