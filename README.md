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

![Version](https://img.shields.io/github/v/release/hoanganhduc/course_management_toolkit?label=version) ![Pre-release](https://img.shields.io/github/v/tag/hoanganhduc/course_management_toolkit?label=pre-release&sort=semver) ![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python) ![GitHub](https://img.shields.io/badge/GitHub-Repo-black?logo=github) ![Docker](https://img.shields.io/badge/Docker-ready-blue?logo=docker) ![GitHub](https://img.shields.io/badge/GitHub-Repo-black?logo=github) ![Status](https://img.shields.io/badge/status-work--in--progress-yellow) ![License](https://img.shields.io/github/license/hoanganhduc/course_management_toolkit)

Utilities for managing course rosters, grading, OCR extraction, and Canvas/Google Classroom workflows. **Work in progress**, mainly designed for **personal use** but open-sourced for others to adapt. Code with the help of [GitHub Copilot](https://github.com/features/copilot/) and [ChatGPT Codex](https://openai.com/).

## Table of contents

- [Course Management Toolkit](#course-management-toolkit)
  - [Table of contents](#table-of-contents)
  - [Install (editable)](#install-editable)
  - [Install into per-user venv](#install-into-per-user-venv)
  - [Run](#run)
- [Common workflows](#common-workflows)
- [Configuration](#configuration)
- [Weekly automation guide](#weekly-automation-guide)
- [Course calendar builder](#course-calendar-builder)
- [Override grades](#override-grades)
  - [Notes](#notes)
  - [External tools (optional)](#external-tools-optional)
  - [Troubleshooting](#troubleshooting)
  - [Troubleshooting OCR](#troubleshooting-ocr)
  - [Documentation](#documentation)
  - [Samples](#samples)
  - [License](#license)

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
course -ume MAT3500-3-Toan-roi-rac-4TC.xlsx
```

Sync Canvas roster into the local database:

```bash
course --sync-canvas
course -sc
```

Notes:
- Canvas sync now stores submission comments and rubric evaluations per assignment in the database.

List auto-generated CLI short aliases:

```bash
course --list-cli-aliases
course -lca
```

Sync Google Classroom roster into the local database:

```bash
course --sync-google-classroom
course -sgc
```

Notes:
- Canvas and Google Classroom score sync normalizes grades to a 10-point scale when max points are available.
- Student ID inference for MAT Excel updates only works with VNU University of Science, Hanoi email format.

Grade resubmissions (lists assignments that need regrading, excludes Roll Call Attendance, and prompts per student unless default is enabled). When keeping old grade, the newer submission is assigned the most recent graded score from the submission history:

```bash
course --grade-resubmission
course --grade-resubmission --keep-old-grade
course -grs
course -grs --keep-old-grade
```

Export a roster to CSV:

Course management toolkit for automating student records, grading workflows, PDF/OCR extraction,
Canvas/Google Classroom operations, AI-assisted checks, and weekly reporting.

```bash
course --export-roster
course -ero
```

Preview an import without writing to the database:

```bash
course --preview-import students.xlsx
course -pi students.xlsx
```

Notes:
- MAT*.xlsx imports ignore score columns (CC, GK, CK, totals); only roster fields are imported.

Export an anonymized roster:

```bash
course --export-anonymized
course -ean
```

Generate a weekly GitHub Actions workflow template:

```bash
course --generate-weekly-workflow
course -gww
```

Run weekly automation for a closed assignment (downloads, checks, grades, reminders):

```bash
course --run-weekly-automation --weekly-assignment-id 123456 --weekly-teacher-canvas-id 987654
course -rwa --weekly-assignment-id 123456 --weekly-teacher-canvas-id 987654
```

If you omit ``--weekly-assignment-id``, the tool scans ``weekly_reports/`` for
previous runs, lists already-processed assignments, then runs on closed Canvas
assignments that are not yet in the weekly reports.

Run weekly automation locally (no GitHub repo needed). Reports are archived under
`weekly_reports/<timestamp>` with a `students.db.bak` copy:

```bash
course --run-weekly-local --weekly-assignment-id 123456 --weekly-local-root "C:\path\to\course-folder"
course -rwl --weekly-assignment-id 123456 --weekly-local-root "C:\path\to\course-folder"
```

Weekly report folders include evidence and outputs such as:
- `run_report.txt`, `data_validation_report.txt`, `grade_diff.csv`
- `weekly_automation_summary.json`
- `final_evaluations/`, `student_submissions/`
- `flagged_submissions_<assignment-name>_<assignment-id>/`
- `students.db.bak`

Clear stored configuration or credentials:

```bash
course --clear-config
course --clear-credentials
course -ccfg
course -ccred
```

Tip: `--google-credentials-path` and `--google-token-path` copy the files into the default config folder with standard filenames, even if you only set them in a separate command before running `--sync-google-classroom`.

Backup or restore the database/config:

```bash
course --backup-db
course --restore-db
course --backup-config
course --restore-config
course -bd
course -rd
course -bc
course -rc
```

Validate student data and export a report:

```bash
course --validate-data
course -vd
```

Preview updates without writing files:

```bash
course --update-mat-excel MAT3500-3-Toan-roi-rac-4TC.xlsx --dry-run --export-grade-diff
course -ume MAT3500-3-Toan-roi-rac-4TC.xlsx -dr --export-grade-diff
```

Student detail sort order (for `--all-details` and `--export-all-details`):

```bash
course --export-all-details students.txt --student-sort-method first_last
course --export-all-details students.txt --student-sort-method last_first
course --export-all-details students.txt --student-sort-method id
course -E students.txt --student-sort-method first_last
course -E students.txt --student-sort-method last_first
course -E students.txt --student-sort-method id
```

You can also set `STUDENT_SORT_METHOD` in `config.json` (first_last, last_first, id).

## Configuration

The tool reads settings from `config.json` stored in a course-specific folder determined by `.course_code`.
On first run, you will be prompted for a course code (e.g., MAT3500) and it will be cached in `.course_code`.
When you load a config file with `--config`, it is copied into the default config folder as `config.json`.

## Weekly automation guide

Local weekly run (no GitHub required):
- Ensure `students.db` is present in the folder you want to use (e.g., a Dropbox folder).
- Ensure your `config.json` is already set (Canvas API URL/key, course ID, OCR keys if used).
- Run `course --run-weekly-local --weekly-assignment-id <ID> --weekly-local-root "<folder>"`.
- Reports and evidence are stored in `weekly_reports/<timestamp>/` under the chosen folder.
  If the assignment ID is omitted, the tool processes closed assignments not yet listed in
  `weekly_reports/`.

GitHub Actions weekly run:
- Store `students.db` in the course repo that hosts the workflow.
- Add required secrets (`CANVAS_API_URL`, `CANVAS_API_KEY`, optional `OCRSPACE_API_KEY`).
- Run `course --generate-weekly-workflow` to create `.github/workflows/weekly-course-tasks.yml`.
- The workflow clones the toolkit repo, runs weekly automation, archives results into
  `weekly_reports/<timestamp>/`, removes the toolkit clone, and commits the updated
  `students.db` plus `weekly_reports/`.

Default config locations:
- Windows: `%APPDATA%\\course\\<course_code>\\config.json`
- macOS: `~/Library/Application Support/course/<course_code>/config.json`
- Linux: `~/.config/course/<course_code>/config.json`

Credential and token files live in the same folder by default:
- `credentials.json` (Google service account)
- `token.pickle` (Google OAuth tokens)

You can override paths via `CREDENTIALS_PATH` and `TOKEN_PATH`.
When you provide Google Classroom credentials/token paths via CLI or the menu, the files are copied into the default folder with the standard names.
To remove stored settings or tokens, use `--clear-config` and `--clear-credentials`.
You can also set `GOOGLE_CLASSROOM_COURSE_ID` in `config.json` to skip the course selection prompt.

Optional settings:
- `LOG_DIR`, `LOG_LEVEL`, `LOG_MAX_BYTES`, `LOG_BACKUP_COUNT` for rotating logs.
- `DB_BACKUP_KEEP`, `CONFIG_BACKUP_KEEP` for backup retention.
- `GRADE_AUDIT_ENABLED`, `GRADE_AUDIT_FIELDS` to control grade audit history stored in the database.
- `COURSE_CODE`, `COURSE_NAME` for calendar titles when not provided via CLI or input files.
- `WEIGHT_CC`, `WEIGHT_GK`, `WEIGHT_CK` for final total score weighting (must sum to 1.0).

## Course calendar builder

Build a course calendar from first-week sessions and export TXT/Markdown/ICS files. If the input file is omitted,
the tool prompts for dates and times interactively.

Example (input file):

```text
course_code: MAT3508
course_name: Discrete Mathematics
weeks: 15
extra_week: yes
holiday: 2026-01-01
holiday: 2026-02-17
session: 2026-01-05 08:00-10:00 | Room 101 | Lecture
session: 2026-01-07 08:00-10:00 | Room 101 | Lecture
```

Notes:
- Course code and course name are required for calendar titles. Provide them in the input file, config (`COURSE_CODE`, `COURSE_NAME`), CLI (`--calendar-course-code`, `--calendar-course-name`), or cache the course code in `.course_code`.
- Fixed Vietnamese holidays (1/1, 4/30, 5/1, 9/2) are automatically excluded.
- Add lunar holidays such as Tet or Hung Vuong via `holiday:` lines or interactive input.
- Add unofficial holidays via `unofficial_holiday:` or `extra_holiday:` lines (comma-separated dates).
- Holidays are auto-fetched via AI using the default provider.
- A make-up week is added only when holidays skip sessions.
- Start time must be earlier than end time.
- Sample input: `sample/calendar/course_calendar_sample_input.txt` (triggers a make-up week).
- Unofficial-holiday sample: `sample/calendar/course_calendar_with_unofficial_holidays.txt`.
- Add unofficial holidays with `--calendar-extra-holidays 2026-03-10,2026-04-05`.

```bash
course --build-course-calendar --calendar-input first_week.txt --calendar-course-code MAT3508 --calendar-course-name "Discrete Mathematics"
course --build-course-calendar
course -bcc --calendar-input first_week.txt --calendar-course-code MAT3508 --calendar-course-name "Discrete Mathematics"
course -bcc
```

Import an existing iCal file into Canvas (requires `CANVAS_LMS_API_URL`, `CANVAS_LMS_API_KEY`, `CANVAS_LMS_COURSE_ID`):

```bash
course --import-canvas-calendar-ics course_calendar.ics --skip-duplicates --dry-run
course -icci course_calendar.ics --skip-duplicates -dr
```
Use `--force` to import without duplicate checks.

Outputs: `course_calendar.txt`, `course_calendar.md`, `course_calendar.ics` in the chosen output directory.
Use `--calendar-output-dir` and `--calendar-output-base` to customize the location and base filename.

## Override grades

Place `override_grades.xlsx` in the working directory (see `sample/overrides/override_grades.xlsx` for the format).
Required columns: `Mã Sinh Viên` or `Họ và Tên`, plus at least one of `CC`/`GK`/`CK` (order does not matter). `STT` and `Lý do` are optional.
Common header aliases are accepted, for example `MSSV`, `Mã SV`, `Họ tên`, `Midterm` (Giữa kỳ), `Final` (Cuối kỳ), `CC` (Chuyên cần), `Reason` (Lý do).
Non-empty CC/GK/CK cells override computed grades, and `Lý do` is appended to the final evaluation output when present.
When using Canvas gradebook CSVs, `Unposted Final Score` is used if `Final Score` is missing or all-zero for CC/GK/CK.
Assignment-group scores are omitted from the final evaluation report when all component scores are 0.
To refine the per-student report with AI, set `REPORT_REFINE_METHOD` to `gemini`, `huggingface`, or `local` in config (requires the corresponding API key for remote providers). The report includes both the default model and the model actually used when AI refinement runs.
Local LLM settings (defaults to Ollama):
- `LOCAL_LLM_COMMAND` (default: `ollama`)
- `LOCAL_LLM_MODEL` (default: `llama3.2:3b`)
- `LOCAL_LLM_ARGS` (optional extra CLI args)
- `LOCAL_LLM_GGUF_DIR` (default: `C:\llm`, scanned recursively for `.gguf` files)
Runtime overrides: `--local-llm-command`, `--local-llm-model`, `--local-llm-args`, `--local-llm-gguf-dir`.
Installing local AI models (examples):
- Ollama: https://ollama.com/ then run `ollama pull llama3.2:3b` and set `LOCAL_LLM_COMMAND=ollama`, `LOCAL_LLM_MODEL=llama3.2:3b`.
- llama.cpp: build `llama-cli` (https://github.com/ggerganov/llama.cpp), set `LOCAL_LLM_COMMAND` to the `llama-cli` path and `LOCAL_LLM_ARGS` to include `-m <path-to-gguf>`.
Use `--refine local` or set `REPORT_REFINE_METHOD=local` to use the local model.
To verify AI connectivity, run `course --test-ai` (or `course -tai`) and check the status for each model. Use `--test-ai-model` (or `-tam`) to test a specific model name, or `--test-ai-gemini-model`/`--test-ai-huggingface-model` when testing `--test-ai all`. For local models, run `course --test-ai local` (or `course -tai local`).
To detect locally installed models (Ollama or llama.cpp compatible), run `course --detect-local-ai` (or `course -dla`).
To list available Gemini models for your API key, run `course --list-ai-models gemini` (or `course -lam gemini`). Hugging Face lists the top public text-generation models (up to 50).
When an AI call is rate-limited, the tool retries and may switch to a different available model with similar capabilities.
Submission quality checks (meaningfulness) can be tuned via config keys: `QUALITY_MIN_CHARS`, `QUALITY_UNIQUE_CHAR_RATIO_MIN`, `QUALITY_REPEAT_CHAR_RATIO_MAX`, `QUALITY_VN_CHAR_RATIO_MIN`, `QUALITY_ALNUM_RATIO_MIN`, `QUALITY_SYMBOL_RATIO_MAX`, `QUALITY_EMPTY_LINE_RATIO_MAX`, `QUALITY_MATH_DENSITY_THRESHOLD`, `QUALITY_LENGTH_RATIO_LOW`, `QUALITY_LENGTH_RATIO_MEDIUM`, `QUALITY_LENGTH_RATIO_HIGH`.
When updating MAT Excel files, use `--export-grade-diff` to save a CSV of old vs new values; database grade changes are tracked in `Grade Audit` when enabled.
A brief per-run summary is appended to `run_report.txt` in the working directory.

## Canvas announcements

Create an announcement from a short message (manual input or a TXT file), optionally refine with AI, preview, and post:

```bash
course --add-canvas-announcement --announcement-title "Week 3" --announcement-message "Please submit by Friday"
course --add-canvas-announcement --announcement-title "Reminder" --announcement-file announcement.txt --refine gemini
course -aa --announcement-title "Week 3" --announcement-message "Please submit by Friday"
course -aa --announcement-title "Reminder" --announcement-file announcement.txt --refine gemini
```

Use `--dry-run` to preview without posting. Omit `--refine` to post the original text without AI.

Sample inputs/outputs:
- `sample/announcements/announcement_input.txt`
- `sample/announcements/announcement_refined_output.txt`
- `sample/announcements/announcement_input_vi.txt`
- `sample/announcements/announcement_refined_output_vi.txt`

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
- Post-OCR AI refinement is disabled; improve scan quality or switch OCR engines if text quality is poor.

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

See `sample/index.md` for anonymized input examples organized by category:
- `sample/mat/MAT-examples.xlsx`
- `sample/overrides/override_grades.xlsx`
- `sample/config/config.sample.json`
- `sample/config/credentials.sample.json`

## License

GPL-3.0-only. See `LICENSE`.
