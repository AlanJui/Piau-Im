# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

This is a Taiwanese Hokkien (河洛話/台語) phonetic annotation system. It processes Chinese characters in Excel workbooks, looks up their Taiwanese pronunciations from SQLite dictionaries, and generates HTML pages using Ruby tags for web publication.

## Environment Setup

```bash
python -m venv .venv
.venv/Scripts/activate          # Windows
pip install -r requirements.txt
```

ChromeDriver is required for Selenium-based scripts. Configure the path in `config.env`:
```
CHROMEDRIVER_PATH=d:\bin\chromedriver-win64\chromedriver.exe
```

## Development Commands

```bash
# Code formatting
black .
ruff format .

# Linting
ruff check .
pylint .          # Only F (fatal) and E (error) checks are enabled via .pylintrc

# Type checking
mypy .

# Run any script directly (most scripts use __main__ entry points)
python a200_查找及填入漢字標音.py

# Regenerate docs/index.html
python a999_自動生成index_html.py
```

Code style: line length 160, double quotes (configured in `pyproject.toml` and `setup.cfg`).

## Architecture

### Core Module Layer (`mod_*.py`)

| Module | Purpose |
|--------|---------|
| `mod_程式.py` | Base framework: `Program` (config) and `ExcelCell` (cell processor) classes |
| `mod_標音.py` | Phonetic conversion: `PiauIm` class, TLPA↔TL↔BP↔MPS2 conversions |
| `mod_ca_ji_tian.py` | SQLite lookups: `HanJiTian` class queries character pronunciations |
| `mod_字庫.py` | Excel-to-dict conversion: `JiKhooDict` manages in-memory character libraries |
| `mod_excel_access.py` | xlwings helpers: cell addressing, worksheet CRUD, named ranges |
| `mod_database.py` | Singleton SQLite connection manager with context manager support |
| `mod_logging.py` | Logging helpers writing to `process_log.txt` and `error_log.txt` |
| `mod_帶調符音標.py` | Tone-diacritic handling utilities |

### Application Scripts (`a###_*.py`)

Scripts follow a numbered naming convention indicating workflow stage:

- **a000–a002**: Reset/clear worksheets
- **a100**: Populate Chinese characters from text file into Excel
- **a200–a260**: Phonetic annotation lookup and fill (core annotation workflow)
- **a300–a320**: Manual correction of annotations
- **a400**: Generate HTML output with Ruby tags
- **a500–a530**: Import/export annotation data
- **a600–a622**: Guangyun (廣韻) and Fifteen-Yin (十五音) dictionary lookups
- **a700–a750**: Typing practice worksheet generators
- **a800–a890**: Character library (漢字庫) maintenance and export
- **a900–a910**: Batch processing
- **a999**: Generate `docs/index.html` for GitHub Pages

### Excel Worksheet Cell Layout

The annotation workbook uses a structured row layout per character column:
- **Row 1**: 人工標音 (manual annotation override)
- **Row 2**: 台語音標 (Taiwan Language Romanization - TL/TLPA)
- **Row 3**: 漢字 (Chinese character)
- **Row 4**: 漢字標音 (final phonetic annotation in target system)

### Databases

- `Ho_Lok_Ue.db` — Primary Taiwanese Hokkien character dictionary
- `Kong_Un.db` — Guangyun (廣韻) Middle Chinese dictionary
- `雅俗通十五音字典.db` — Fifteen-Yin (十五音) dictionary

Database selection is controlled via environment variables loaded from `.env`.

### Supported Phonetic Systems

TLPA, TL (台羅), BP (閩拼), MPS2 (方音符號), POJ (白話字), 十五音 (Fifteen-Yin). Conversions are handled by `PiauIm` class in `mod_標音.py`.

### Output

- `docs/` — HTML files deployed to GitHub Pages (via `.github/workflows/static.yml`)
- `output*/` — Intermediate and final Excel/text outputs per processing run
