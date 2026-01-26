# WorkabiliBot

Workabili Bot is a Python-based automation tool designed to streamline
Workability-related workflows through a simple GUI and configurable rules.

## Project Structure

workabili-bot/
├─ app/ # Application source code
├─ config/ # YAML configuration files
├─ pyproject.toml
├─ requirements.txt
└─ README.md

## Setup

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

python -m app.gui

## Notes

This project is under active development.

Installing Playwright the Python package is not enough.
After installing requirements, you must also run:

playwright install

This downloads the browser binaries (Chromium / Firefox / WebKit).
Without this, your app will explode in very confusing ways 😅
