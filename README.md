# WorkabiliBot

Workabili-Bot 3000 is a Python-based automation tool designed to streamline
Workability-related workflows through a simple GUI and rules in config.yaml.

> [!WARNING]
> **This tool is built specifically for SBCSS and has only been tested at SBCSS.**
> If you are from a different county, district, or institution, `config/config.yaml`
> must be completely reconfigured before use — school mappings, org names, validation
> rules, and WAI field values are all SBCSS-specific. Running it unconfigured against
> another org's WAI instance will produce incorrect results.
>
> Please reach out before attempting to use this: Derek.Carlson@sbcss.net

This project is under active development and is not even alpha yet, so please be careful.

## Project Structure

```text
workabili-bot/
  app/              Application source code (GUI, runner, smoke tests)
  assets/           GUI images and sounds
  config/           YAML configuration
  input/            Input Excel files (only Input_TEMPLATE.xlsx is tracked; all data files are gitignored)
  output/           Logs, screenshots, traces, run ledger (not tracked in git)
  tests/            Test suite
  Workabili-Bot3000.vbs   Windows launcher script
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
python -m app.search_smoke --ssid 1234567890
```

By default, this uses `credentials.username` from `config/config.yaml` and prompts
for the password securely. It opens the browser, logs in, opens Student Records,
sets the search filters, searches the SSID, prints the result, then pauses so you
can inspect the page.

Optional examples:

```bash
python -m app.search_smoke --ssid 1234567890 --username your.email@example.org
python -m app.search_smoke --ssid 1234567890 --slow-mo-ms 800
python -m app.search_smoke --ssid 1234567890 --no-pause
```

To dry-run the prior-year Transfer modal without clicking Save:

```bash
python -m app.search_smoke --ssid 1234567890 --exercise-prior-year-transfer --aeries-school "CS - BARBARA PHELPS CS"
```

For non-interactive local testing, set `WAI_PASSWORD` in the shell before running.
Do not commit passwords into source files.

## Notes

This project is under active development and is not even alpha yet, so
please be careful.  config.yaml needs to be entirely updated if you are
from a school, district, or institution different from SBCSS.  This 
has only been tested on SBCSS.  
