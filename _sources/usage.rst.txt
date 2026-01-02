Usage
=====

Install (editable)
------------------

.. code-block:: bash

   pip install -e .

Install into per-user venv
--------------------------

Linux/macOS:

.. code-block:: bash

   make install
   ~/.course_venv/bin/course

Windows:

.. code-block:: bat

   make.bat install
   %USERPROFILE%\.course_venv\Scripts\course.exe

Run
---

.. code-block:: bash

   course

Interactive menu tips
---------------------

- Use arrow keys (or W/S) to move, Enter to select, q to quit.
- Type a menu number quickly to jump to that option.

Clear stored settings
---------------------

.. code-block:: bash

   course --clear-config
   course --clear-credentials
   course -ccfg
   course -ccred

Tip: ``--google-credentials-path`` and ``--google-token-path`` copy the files into
the default config folder with standard filenames, even if you only set them in
a separate command before running ``--sync-google-classroom``.

Configuration and file locations
--------------------------------

The tool reads settings from ``config.json`` stored in a course-specific folder
determined by ``.course_code``. On first run, you will be prompted for a course
code (for example, MAT3500) and it will be cached in ``.course_code``.
When you load a config file with ``--config``, it is copied into the default
config folder as ``config.json``.
Sample config: ``sample/config/config.sample.json``.

Default config locations:

- Windows: ``%APPDATA%\course\<course_code>\config.json``
- macOS: ``~/Library/Application Support/course/<course_code>/config.json``
- Linux: ``~/.config/course/<course_code>/config.json``

Credential and token files live in the same folder by default:

- ``credentials.json`` (Google service account)
- ``token.pickle`` (Google OAuth tokens)

You can override paths via ``CREDENTIALS_PATH`` and ``TOKEN_PATH`` in the config
file.
When you provide Google Classroom credentials/token paths via CLI or the menu,
the files are copied into the default folder with the standard names.
You can also set ``GOOGLE_CLASSROOM_COURSE_ID`` in ``config.json`` to skip the
course selection prompt.

Final evaluation weights can be configured via ``WEIGHT_CC``, ``WEIGHT_GK``,
and ``WEIGHT_CK`` in ``config.json``. The weights must sum to 1.0.

OCR dependencies and setup
--------------------------

Local OCR and PDF conversion require system tools. Install them before using
``--ocr-service tesseract`` or features that convert PDF pages to images.

Windows (PowerShell):

.. code-block:: powershell

   winget install -e --id UB-Mannheim.TesseractOCR
   winget install -e --id oschwartz10612.Poppler

macOS (Homebrew):

.. code-block:: bash

   brew install tesseract poppler

Linux:

.. code-block:: bash

   # Debian/Ubuntu
   sudo apt-get update
   sudo apt-get install -y tesseract-ocr poppler-utils

   # Fedora
   sudo dnf install -y tesseract poppler-utils

   # Arch
   sudo pacman -S tesseract poppler

Verify:

.. code-block:: bash

   tesseract --version
   pdftoppm -h

Note: Post-OCR AI refinement is disabled; improve scan quality or switch OCR engines if text quality is poor.

Canvas and Google Classroom setup
---------------------------------

Populate the following keys in ``config.json`` (or load from a JSON file with
``--config``):

- ``CANVAS_LMS_API_URL``
- ``CANVAS_LMS_API_KEY``
- ``CANVAS_LMS_COURSE_ID``
- ``CREDENTIALS_PATH``
- ``TOKEN_PATH``

Canvas operations will use these defaults unless overridden by flags like
``--canvas-course-id``.

Common workflows
----------------

Sync Canvas roster into the local database:

.. code-block:: bash

   course --sync-canvas
   course -sc

Notes:

- Canvas sync now stores submission comments and rubric evaluations per assignment in the database.

List auto-generated CLI short aliases:

.. code-block:: bash

   course --list-cli-aliases
   course -lca

Sync Google Classroom roster into the local database:

.. code-block:: bash

   course --sync-google-classroom
   course -sgc

Notes:

- Canvas and Google Classroom score sync normalizes grades to a 10-point scale when max points are available.
- Student ID inference for MAT Excel updates only works with VNU University of Science, Hanoi email format.

Grade resubmissions (lists assignments that need regrading, excludes Roll Call Attendance, and prompts per student unless default is enabled). When keeping old grade, the newer submission is assigned the most recent graded score from the submission history:

.. code-block:: bash

   course --grade-resubmission
   course --grade-resubmission --keep-old-grade
   course -grs
   course -grs --keep-old-grade

Update a MAT*.xlsx file with grades from the local database:

.. code-block:: bash

   course --update-mat-excel MAT3500-3-Toan-roi-rac-4TC.xlsx
   course -ume MAT3500-3-Toan-roi-rac-4TC.xlsx

Export a roster to CSV:

.. code-block:: bash

   course --export-roster
   course -ero

Preview an import (no write):

.. code-block:: bash

   course --preview-import students.xlsx
   course -pi students.xlsx

Notes:

- MAT*.xlsx imports ignore score columns (CC, GK, CK, totals); only roster fields are imported.

Export an anonymized roster:

.. code-block:: bash

   course --export-anonymized
   course -ean

Generate a weekly workflow template:

.. code-block:: bash

   course --generate-weekly-workflow
   course -gww

Run weekly automation:

.. code-block:: bash

   course --run-weekly-automation --weekly-assignment-id 123456 --weekly-teacher-canvas-id 987654
   course -rwa --weekly-assignment-id 123456 --weekly-teacher-canvas-id 987654

If ``--weekly-assignment-id`` is omitted, the tool scans ``weekly_reports/`` to list
assignments already processed and then runs on closed assignments not yet in the reports.

Run weekly automation locally (no GitHub repo needed):

.. code-block:: bash

   course --run-weekly-local --weekly-assignment-id 123456 --weekly-local-root "C:\\path\\to\\course-folder"
   course -rwl --weekly-assignment-id 123456 --weekly-local-root "C:\\path\\to\\course-folder"

Weekly automation guide
-----------------------

Local weekly run:

- Ensure ``students.db`` exists in the target folder.
- Ensure ``config.json`` is configured (Canvas API URL/key, course ID, OCR keys if needed).
- Run ``course --run-weekly-local`` with ``--weekly-assignment-id`` and an optional
  ``--weekly-local-root`` to choose where reports are stored.
- Reports and evidence are stored in ``weekly_reports/<timestamp>/`` with a
  ``students.db.bak`` backup.
- If the assignment ID is omitted, the tool processes closed assignments not yet listed
  in ``weekly_reports/``.

GitHub Actions weekly run:

- Store ``students.db`` in the repo that hosts the workflow.
- Configure secrets ``CANVAS_API_URL``, ``CANVAS_API_KEY``, and optional OCR keys.
- Generate the workflow via ``course --generate-weekly-workflow``.
- The workflow clones the toolkit repo, runs weekly checks, archives artifacts into
  ``weekly_reports/<timestamp>/``, removes the toolkit clone, and commits updates.


Course calendar builder
----------------------

Build a course calendar from first-week sessions and export TXT/Markdown/ICS files. If the input file is omitted,
the tool prompts for dates and times interactively.

Example input file::

  course_code: MAT3508
  course_name: Discrete Mathematics
  weeks: 15
  extra_week: yes
  holiday: 2026-01-01
  holiday: 2026-02-17
  session: 2026-01-05 08:00-10:00 | Room 101 | Lecture
  session: 2026-01-07 08:00-10:00 | Room 101 | Lecture

Notes:
- Course code and course name are required for calendar titles. Provide them in the input file, config (``COURSE_CODE``, ``COURSE_NAME``), CLI (``--calendar-course-code``, ``--calendar-course-name``), or cache the course code in ``.course_code``.
- Fixed Vietnamese holidays (1/1, 4/30, 5/1, 9/2) are automatically excluded.
- Add lunar holidays such as Tet or Hung Vuong via ``holiday:`` lines or interactive input.
- Add unofficial holidays via ``unofficial_holiday:`` or ``extra_holiday:`` lines (comma-separated dates).
- Holidays are auto-fetched via AI using the default provider.
- A make-up week is added only when holidays skip sessions.
- Start time must be earlier than end time.
- Sample input: ``sample/calendar/course_calendar_sample_input.txt`` (triggers a make-up week).
- Unofficial-holiday sample: ``sample/calendar/course_calendar_with_unofficial_holidays.txt``.
- Add unofficial holidays with ``--calendar-extra-holidays 2026-03-10,2026-04-05``.

Usage::

  course --build-course-calendar --calendar-input first_week.txt --calendar-course-code MAT3508 --calendar-course-name "Discrete Mathematics"
  course --build-course-calendar
  course -bcc --calendar-input first_week.txt --calendar-course-code MAT3508 --calendar-course-name "Discrete Mathematics"
  course -bcc

Import an existing iCal file into Canvas (requires ``CANVAS_LMS_API_URL``, ``CANVAS_LMS_API_KEY``, ``CANVAS_LMS_COURSE_ID``)::

  course --import-canvas-calendar-ics course_calendar.ics --skip-duplicates --dry-run
  course -icci course_calendar.ics --skip-duplicates -dr
  # Use --force to import without duplicate checks.

Outputs: ``course_calendar.txt``, ``course_calendar.md``, ``course_calendar.ics`` in the chosen output directory.
Use ``--calendar-output-dir`` and ``--calendar-output-base`` to customize the location and base filename.
Backup and restore
------------------

.. code-block:: bash

   course --backup-db
   course --restore-db
   course --backup-config
   course --restore-config
   course -bd
   course -rd
   course -bc
   course -rc

Data validation report
----------------------

.. code-block:: bash

   course --validate-data
   course -vd

Dry-run mode
------------

Preview changes without writing files:

.. code-block:: bash

   course --update-mat-excel MAT3500-3-Toan-roi-rac-4TC.xlsx --dry-run --export-grade-diff
   course -ume MAT3500-3-Toan-roi-rac-4TC.xlsx -dr --export-grade-diff

Student detail sort order (for ``--all-details`` and ``--export-all-details``):

.. code-block:: bash

   course --export-all-details students.txt --student-sort-method first_last
   course --export-all-details students.txt --student-sort-method last_first
   course --export-all-details students.txt --student-sort-method id
   course -E students.txt --student-sort-method first_last
   course -E students.txt --student-sort-method last_first
   course -E students.txt --student-sort-method id

You can also set ``STUDENT_SORT_METHOD`` in ``config.json`` (first_last, last_first, id).

Logging
-------

Rotating logs are written to the config folder by default. Configure:

- ``LOG_DIR``
- ``LOG_LEVEL``
- ``LOG_MAX_BYTES``
- ``LOG_BACKUP_COUNT``

Backup retention is controlled by:

- ``DB_BACKUP_KEEP``
- ``CONFIG_BACKUP_KEEP``

A brief per-run summary is appended to ``run_report.txt`` in the working directory.

Canvas announcements
--------------------

Create an announcement from a short message (manual input or a TXT file), optionally refine with AI, preview, and post::

  course --add-canvas-announcement --announcement-title "Week 3" --announcement-message "Please submit by Friday"
  course --add-canvas-announcement --announcement-title "Reminder" --announcement-file announcement.txt --refine gemini
  course -aa --announcement-title "Week 3" --announcement-message "Please submit by Friday"
  course -aa --announcement-title "Reminder" --announcement-file announcement.txt --refine gemini

Use ``--dry-run`` to preview without posting. Omit ``--refine`` to post the original text without AI.

Sample inputs/outputs:
- ``sample/announcements/announcement_input.txt``
- ``sample/announcements/announcement_refined_output.txt``
- ``sample/announcements/announcement_input_vi.txt``
- ``sample/announcements/announcement_refined_output_vi.txt``

Override grades
----------------

Place ``override_grades.xlsx`` in the working directory (see ``sample/overrides/override_grades.xlsx`` for the format).
Required columns: ``Mã Sinh Viên`` or ``Họ và Tên``, plus at least one of ``CC``/``GK``/``CK`` (order does not matter). ``STT`` and ``Lý do`` are optional.
Common header aliases are accepted, for example ``MSSV``, ``Mã SV``, ``Họ tên``, ``Midterm`` (Giữa kỳ), ``Final`` (Cuối kỳ), ``CC`` (Chuyên cần), ``Reason`` (Lý do).
Non-empty CC/GK/CK cells override computed grades, and ``Lý do`` is appended to the final evaluation output when present.
When using Canvas gradebook CSVs, ``Unposted Final Score`` is used if ``Final Score`` is missing or all-zero for CC/GK/CK.
Assignment-group scores are omitted from the final evaluation report when all component scores are 0.
Final evaluation TXT output includes the weighted formula used for the total score.

AI report refinement
-------------------

Set ``REPORT_REFINE_METHOD`` to ``gemini``, ``huggingface``, or ``local`` in ``config.json`` (requires the corresponding API key for remote providers).
When AI refinement runs, the report includes the default model and the model actually used.
Local LLM settings (defaults to Ollama):

- ``LOCAL_LLM_COMMAND`` (default: ``ollama``)
- ``LOCAL_LLM_MODEL`` (default: ``llama3.2:3b``)
- ``LOCAL_LLM_ARGS`` (optional extra CLI args)
- ``LOCAL_LLM_GGUF_DIR`` (default: ``C:\llm``, scanned recursively for ``.gguf`` files)
Runtime overrides: ``--local-llm-command``, ``--local-llm-model``, ``--local-llm-args``, ``--local-llm-gguf-dir``.

Installing local AI models (examples):
- Ollama: https://ollama.com/ then run ``ollama pull llama3.2:3b`` and set ``LOCAL_LLM_COMMAND=ollama``, ``LOCAL_LLM_MODEL=llama3.2:3b``.
- llama.cpp: build ``llama-cli`` (https://github.com/ggerganov/llama.cpp), set ``LOCAL_LLM_COMMAND`` to the ``llama-cli`` path and ``LOCAL_LLM_ARGS`` to include ``-m <path-to-gguf>``.
Use ``--refine local`` or set ``REPORT_REFINE_METHOD=local`` to use the local model.


AI model verification and listing
---------------------------------

Verify credentials and connectivity:

.. code-block:: bash

   course --test-ai
   course -tai

Verify the local model:

.. code-block:: bash

   course --test-ai local
   course -tai local

Detect locally installed models (Ollama or llama.cpp compatible):

.. code-block:: bash

   course --detect-local-ai
   course -dla

Test a specific model name:

.. code-block:: bash

   course --test-ai gemini --test-ai-model gemini-2.5-flash
   course -tai gemini -tam gemini-2.5-flash

List available models:

.. code-block:: bash

   course --list-ai-models gemini
   course -lam gemini

When an AI call is rate-limited, the tool retries and may switch to a different available model with similar capabilities.

Submission quality checks
-------------------------

Meaningfulness checks can be tuned via config keys:

- ``QUALITY_MIN_CHARS``
- ``QUALITY_UNIQUE_CHAR_RATIO_MIN``
- ``QUALITY_REPEAT_CHAR_RATIO_MAX``
- ``QUALITY_VN_CHAR_RATIO_MIN``
- ``QUALITY_ALNUM_RATIO_MIN``
- ``QUALITY_SYMBOL_RATIO_MAX``
- ``QUALITY_EMPTY_LINE_RATIO_MAX``
- ``QUALITY_MATH_DENSITY_THRESHOLD``
- ``QUALITY_LENGTH_RATIO_LOW``
- ``QUALITY_LENGTH_RATIO_MEDIUM``
- ``QUALITY_LENGTH_RATIO_HIGH``

Troubleshooting
---------------

- If ``course`` cannot find ``students.db``, confirm you are running the command
  in the intended working directory.
- If OCR commands are missing, recheck your PATH or reinstall Tesseract/Poppler.
- If Canvas/Google Classroom calls fail, verify API keys and course IDs in
  ``config.json``.
