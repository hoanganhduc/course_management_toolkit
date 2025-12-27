# Course Management Toolkit

<div align="center">
  <a href="https://www.buymeacoffee.com/hoanganhduc" target="_blank" rel="noopener noreferrer">
    <img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me A Coffee" height="40" style="margin-right: 10px;" />
  </a>
  <a href="https://ko-fi.com/hoanganhduc" target="_blank" rel="noopener noreferrer">
    <img src="https://storage.ko-fi.com/cdn/kofi3.png?v=3" alt="Ko-fi" height="40" />
  </a>
  <a href="https://bmacc.app/tip/hoanganhduc" target="_blank" rel="noopener noreferrer">
		<img src="https://bmacc.app/images/bmacc-logo.png" alt="Buy Me a Crypto Coffee" style="height: 40px;">
	</a>
</div>

![Version](https://img.shields.io/github/v/release/hoanganhduc/course_management_toolkit?label=version) ![Pre-release](https://img.shields.io/github/v/tag/hoanganhduc/course_management_toolkit?label=pre-release&sort=semver) ![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python) ![GitHub](https://img.shields.io/badge/GitHub-Repo-black?logo=github) ![Status](https://img.shields.io/badge/status-work--in--progress-yellow) ![License](https://img.shields.io/github/license/hoanganhduc/course_management_toolkit)


Utilities for managing course rosters, grading, OCR extraction, and Canvas/Google Classroom workflows.

## Table of contents

- [Install (editable)](#install-editable)
- [Install into per-user venv](#install-into-per-user-venv)
- [Run](#run)
- [Common workflows](#common-workflows)
- [Configuration](#configuration)
- [Override grades](#override-grades)
- [Notes](#notes)
- [External tools (optional)](#external-tools-optional)
- [Troubleshooting](#troubleshooting)
- [Troubleshooting OCR](#troubleshooting-ocr)
- [Documentation](#documentation)
- [Samples](#samples)

## Install (editable)

```bash
pip install -e .
```

## Install into per-user venv

Linux/macOS:

```bash
make install
~/.course_venv/bin/course
```

Windows:

```bat
make.bat install
%USERPROFILE%\.course_venv\Scripts\course.exe
```

## Run

```bash
course
```

Interactive menu tips:
- Use arrow keys (or W/S) to move, Enter to select, q to quit.
- You can also type the menu number quickly to jump to an option.

## Common workflows

Update a MAT*.xlsx file with grades from the local database:

```bash
course --update-mat-excel MAT3500-3-Toan-roi-rac-4TC.xlsx
```

Sync Canvas roster into the local database:

```bash
course --sync-canvas
```

Sync Google Classroom roster into the local database:

```bash
course --sync-google-classroom
```

Export a roster to CSV:

```bash
course --export-roster
```

Clear stored configuration or credentials:

```bash
course --clear-config
course --clear-credentials
```

## Configuration

The tool reads settings from `config.json` stored in a course-specific folder determined by `.course_code`.
On first run, you will be prompted for a course code (e.g., MAT3500) and it will be cached in `.course_code`.

Default config locations:
- Windows: `%APPDATA%\\course\\<course_code>\\config.json`
- macOS: `~/Library/Application Support/course/<course_code>/config.json`
- Linux: `~/.config/course/<course_code>/config.json`

Credential and token files live in the same folder by default:
- `credentials.json` (Google service account)
- `token.pickle` (Google OAuth tokens)

You can override paths via `CREDENTIALS_PATH` and `TOKEN_PATH`.
To remove stored settings or tokens, use `--clear-config` and `--clear-credentials`.

## Override grades

Place `override_grades.xlsx` in the working directory (see `sample/override_grades.xlsx` for the format).
Non-empty CC/CK/GK cells override computed grades, and the `Reason` column is appended to the final evaluation output.

## Notes

Some features rely on external system tools (for example, Tesseract OCR and Poppler for PDF processing).
Student databases are resolved from the current working directory (for example, running `course` in a folder will read or write `students.db` there).

## External tools (optional)

PDF extraction and local OCR require system tools. Install them before using `--ocr-service tesseract` or features that convert PDF pages to images.

Official pages (external tools used by this project):
- Tesseract OCR: https://tesseract-ocr.github.io/
- Poppler: https://poppler.freedesktop.org/
- OCR.Space: https://ocr.space/
- PaddleOCR: https://github.com/PaddlePaddle/PaddleOCR

Windows (PowerShell):

```powershell
winget install -e --id UB-Mannheim.TesseractOCR
winget install -e --id oschwartz10612.Poppler
```

If commands are not found, add the install folders to `PATH` (common defaults):
- `C:\Program Files\Tesseract-OCR`
- `C:\Program Files\poppler\Library\bin`

macOS (Homebrew):

```bash
brew install tesseract poppler
```

Linux:

```bash
# Debian/Ubuntu
sudo apt-get update
sudo apt-get install -y tesseract-ocr poppler-utils

# Fedora
sudo dnf install -y tesseract poppler-utils

# Arch
sudo pacman -S tesseract poppler
```

Verify:

```bash
tesseract --version
pdftoppm -h
```

For `--ocr-service ocrspace`, set `OCRSPACE_API_KEY` in your config file.


## Troubleshooting

- If `course` cannot find `students.db`, confirm you are running the command in the intended working directory.
- If OCR commands are missing, recheck your PATH or reinstall Tesseract/Poppler.
- If Canvas/Google Classroom calls fail, verify API keys and course IDs in `config.json`.

## Troubleshooting OCR

Common errors and fixes:

- `tesseract: command not found` (macOS/Linux) or `'tesseract' is not recognized` (Windows): confirm the install and that the bin folder is on `PATH`.
  - Windows: `where tesseract` and `where pdftoppm` should return paths. If not, add `C:\\Program Files\\Tesseract-OCR` and `C:\\Program Files\\poppler\\Library\\bin` to `PATH`, then reopen the terminal.
  - macOS: `brew --prefix tesseract` and `brew --prefix poppler` should point to installed prefixes; ensure Homebrew is on `PATH`.
  - Linux: `which tesseract` and `which pdftoppm` should resolve. If missing, reinstall with your package manager.
- `pdftoppm` missing: Poppler is not installed or not on `PATH`. Reinstall Poppler and re-open your terminal.
- `TesseractNotFoundError` in Python: the OS command is not visible to the Python process; confirm your IDE/terminal inherits the updated `PATH`.

## Documentation

Sphinx documentation lives in `docs/`.

Build HTML docs:

```bash
pip install -r docs/requirements.txt
cd docs
make html
```

Windows:

```bat
pip install -r docs\requirements.txt
cd docs
make.bat html
```

## Samples

See `sample/index.md` for anonymized input examples:
- `sample/MAT-examples.xlsx`
- `sample/canvas_gradebook.csv`
- `sample/override_grades.xlsx`
- `sample/config.sample.json`
- `sample/credentials.sample.json`
