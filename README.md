# WorkabiliBot

Workabili Bot is a Python-based automation tool designed to streamline
Workability-related workflows through a simple GUI and configurable rules.

## Project Structure

```text
workabili-bot/
  app/              Application source code
  config/           YAML configuration files
  assets/           GUI images and sounds
  input/            Input Excel files
  output/           Logs, screenshots, traces, and run ledger
  pyproject.toml
  requirements.txt
  README.md
```

## Setup

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
playwright install
```

Run the GUI:

```bash
python -m app.gui
```

Installing the Playwright Python package is not enough. After installing
requirements, `playwright install` downloads the browser binaries needed by the
automation.

## Non-GUI Search Smoke Test

To test the WAI search flow without opening the GUI or processing an Excel file:

```bash
python -m app.search_smoke --ssid 1304529546
```

By default, this uses `credentials.username` from `config/config.yaml` and prompts
for the password securely. It opens the browser, logs in, opens Student Records,
sets the search filters, searches the SSID, prints the result, then pauses so you
can inspect the page.

Optional examples:

```bash
python -m app.search_smoke --ssid 1304529546 --username your.email@example.org
python -m app.search_smoke --ssid 1304529546 --slow-mo-ms 800
python -m app.search_smoke --ssid 1304529546 --no-pause
```

To dry-run the prior-year Transfer modal without clicking Save:

```bash
python -m app.search_smoke --ssid 1304529546 --exercise-prior-year-transfer --aeries-school "CS - BARBARA PHELPS CS"
```

For non-interactive local testing, set `WAI_PASSWORD` in the shell before running.
Do not commit passwords into source files.

## Notes

This project is under active development.
