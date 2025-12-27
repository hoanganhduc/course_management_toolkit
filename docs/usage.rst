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

Configuration and file locations
--------------------------------

The tool reads settings from ``config.json`` stored in a course-specific folder
determined by ``.course_code``. On first run, you will be prompted for a course
code (for example, MAT3500) and it will be cached in ``.course_code``.

Default config locations:

- Windows: ``%APPDATA%\course\<course_code>\config.json``
- macOS: ``~/Library/Application Support/course/<course_code>/config.json``
- Linux: ``~/.config/course/<course_code>/config.json``

Credential and token files live in the same folder by default:

- ``credentials.json`` (Google service account)
- ``token.pickle`` (Google OAuth tokens)

You can override paths via ``CREDENTIALS_PATH`` and ``TOKEN_PATH`` in the config
file.

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

Sync Google Classroom roster into the local database:

.. code-block:: bash

   course --sync-google-classroom

Update a MAT*.xlsx file with grades from the local database:

.. code-block:: bash

   course --update-mat-excel MAT3500-3-Toan-roi-rac-4TC.xlsx

Export a roster to CSV:

.. code-block:: bash

   course --export-roster

Override grades
----------------

Place ``override_grades.xlsx`` in the working directory (see ``sample/override_grades.xlsx`` for the format).
Required columns: ``M? Sinh Vi?n`` or ``H? v? T?n``, plus at least one of ``CC``/``GK``/``CK`` (order does not matter). ``STT`` and ``L? do`` are optional.
Common header aliases are accepted, for example ``MSSV``, ``M? SV``, ``H? t?n``, ``Midterm`` (Gi?a k?), ``Final`` (Cu?i k?), ``CC`` (Chuy?n c?n), ``Reason`` (L? do).
Non-empty CC/GK/CK cells override computed grades, and ``L? do`` is appended to the final evaluation output when present.

AI report refinement
-------------------

Set ``REPORT_REFINE_METHOD`` to ``gemini`` or ``huggingface`` in ``config.json`` (requires the corresponding API key).
When AI refinement runs, the report includes the default model and the model actually used.

AI model verification and listing
---------------------------------

Verify credentials and connectivity:

.. code-block:: bash

   course --test-ai

Test a specific model name:

.. code-block:: bash

   course --test-ai gemini --test-ai-model gemini-2.5-flash

List available models:

.. code-block:: bash

   course --list-ai-models gemini

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
