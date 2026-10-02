# PDF Converter Telegram Bot

A clean, modern Telegram bot that converts uploaded files into PDF format, lets users define the output name, and returns the converted document directly in chat.

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.11%2B-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python 3.11+" />
  <img src="https://img.shields.io/badge/Telegram-Bot-26A5E4?style=for-the-badge&logo=telegram&logoColor=white" alt="Telegram Bot" />
  <img src="https://img.shields.io/badge/PDF-Converter-FF6F61?style=for-the-badge" alt="PDF Converter" />
</p>

## Overview

This project is designed for fast and reliable PDF conversion from Telegram uploads. It supports images, text files, and common office documents, and it keeps the user experience simple: upload a file, choose a name, and receive a PDF.

## Features

- Converts images to PDF
- Converts text files such as TXT, MD, CSV, and LOG
- Converts Office documents via LibreOffice
- Accepts existing PDF files unchanged
- Uses a friendly conversation flow in Telegram
- Sanitizes filenames for clean output naming
- Cleans temporary user data automatically after conversion

## Supported formats

| Category | Formats |
| --- | --- |
| Images | .png, .jpg, .jpeg, .webp |
| Text | .txt, .md, .log, .csv |
| Office | .doc, .docx, .ppt, .pptx, .xls, .xlsx, .odt, .ods, .odp, .rtf |
| PDF | .pdf |

## Project structure

```text
pdf_bot/
├── src/
│   ├── __init__.py
│   └── bot.py               # Main Telegram bot implementation
├── bot.py                   # Small startup entrypoint
├── requirements.txt         # Python dependencies
├── README.md                # Project documentation
├── .gitignore               # Ignore env and local files
├── .env.example             # Sample environment variables
├── .venv/                   # Local virtual environment
├── __pycache__/             # Python cache
└── .DS_Store                # macOS metadata
```

## Setup

1. Clone the project:

```bash
git clone <repo-url>
cd pdf_bot
```

2. Create a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Copy the sample environment file:

```bash
cp .env.example .env
```

5. Add your Telegram credentials to `.env`:

```env
BOT_TOKEN=your_telegram_bot_token_here
```

Load those values into your local shell before starting the bot:

```bash
set -a
source .env
set +a
```

`LOCAL_BOT_API` is optional. Leave it unset to use Telegram's public Bot API. Set it only when you run a separate Telegram Bot API server that the bot can reach. On a cloud host, add `BOT_TOKEN` in the host's environment-variable or secrets settings instead of uploading `.env`.

6. Install LibreOffice for Office document conversion.

7. Run the bot:

```bash
python bot.py
```

## How it works

1. Start the bot with `/start`.
2. Upload a document or image.
3. Keep the uploaded filename or choose a new PDF filename.
4. Receive the converted PDF in Telegram.
5. Use `/cancel` to reset the current flow.

## Always-on deployment

The bot is a long-running polling process. It runs only while the machine or cloud service running `python bot.py` is online; `LOCAL_BOT_API` selects the Telegram API endpoint and does not host the bot. For Choreo service deployment, use `python bot.py` as the start command and expose the port specified by `PORT` (defaults to `8080`); the process serves `GET /healthz` for service health checks while polling Telegram. Set `BOT_TOKEN` in Choreo's environment-variable or secret settings. Do not set `LOCAL_BOT_API` unless that host can reach your own Bot API server. Install LibreOffice on the host if you need Office document conversion.

## Environment variables

| Variable | Required | Description |
| --- | --- | --- |
| `BOT_TOKEN` | Yes | Telegram bot token from BotFather |
| `LOCAL_BOT_API` | No | Optional URL of a separate Telegram Bot API server; unset uses Telegram's public API |
| `PORT` | No | HTTP health-check port for service deployments; defaults to `8080` |

## Requirements

- Python 3.11+
- Telegram bot token
- LibreOffice (for office file conversion)

## Notes

This project uses temporary directories for each conversion session and removes old state after processing to keep the bot reliable and clean.

## License

This project is intended for personal and educational use. Add your preferred license if you plan to distribute it publicly.

---

Built for fast, clean PDF conversion from Telegram uploads.
