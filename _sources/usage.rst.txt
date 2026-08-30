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

Full CLI flag reference is available in :doc:`cli_reference`.

Interactive menu tips
---------------------

- Use arrow keys (or W/S) to move, Enter to select, q to quit.
- Type a menu number quickly to jump to that option.
- All menu actions have CLI equivalents; see :doc:`cli_reference` for the full list.

Menu ↔ CLI examples
-------------------

.. list-table::
   :header-rows: 1

   * - Menu item
     - CLI equivalent
   * - List students by email domain
     - ``course --list-email-domain gmail.com``
   * - List students with duplicate names
     - ``course --list-duplicate-names --duplicate-name-field name``
   * - Import students from Google Sheet
     - ``course --add-google-sheet <URL>``
   * - Download Google Classroom submissions
     - ``course --download-google-classroom-submissions --gc-download-coursework-id <ID>``
   * - Run weekly automation
     - ``course --run-weekly-automation --weekly-assignment-id <ID>``
   * - Classroom50 preflight
     - ``course --classroom50-preflight``
   * - Sync Classroom50 roster
     - ``course --sync-classroom50 --classroom50-org ORG --classroom50-classroom SHORT``

Classroom50 (foundation50)
--------------------------

Wraps Classroom50 instructor tooling (``gh teacher``) for roster list/sync/export.
Does not reimplement Classroom50. Requires GitHub CLI and the Classroom50 teacher
extension.

.. code-block:: bash

   course --classroom50-preflight
   course --list-classroom50-classrooms --classroom50-org my-org
   course --list-classroom50-roster --classroom50-org my-org --classroom50-classroom short-name
   course --list-classroom50-assignments --classroom50-org my-org --classroom50-classroom short-name
   course --sync-classroom50 -sc50 --classroom50-org my-org --classroom50-classroom short-name
   course --export-classroom50-roster classroom50_roster.csv

Download of student submissions is available only on the full human ``course`` CLI:

.. code-block:: bash

   course --download-classroom50 --classroom50-org my-org --classroom50-classroom short-name \
     --classroom50-assignment assignment-slug --classroom50-download-dest ./c50_submissions

See the **Classroom50** section in :doc:`cli_reference` for all flags.

Agent-safe entrypoints
----------------------

Dedicated modules force agent mode (``COURSE_AGENT_MODE=1``) and expose a
restricted surface for AI agents. Destructive or high-blast-radius operations
are refused; use the interactive ``course`` CLI as a human when those are needed.

.. list-table::
   :header-rows: 1

   * - Module
     - Allowed (typical)
     - Refused (examples)
   * - ``python -m course_hoanganhduc.c50_agent``
     - preflight, list-*, sync, export
     - download
   * - ``python -m course_hoanganhduc.canvas_agent``
     - preflight, list-assignments/members, search-user, sync
     - unenroll, grade, invite, announce, download, messages, pages
   * - ``python -m course_hoanganhduc.gclass_agent``
     - preflight, list-courses/students, sync
     - unenroll, grade, download
   * - ``python -m course_hoanganhduc.db_agent``
     - search, details, list-*, export-*, count
     - modify, restore-db, import-apply, delete

Examples:

.. code-block:: bash

   python -m course_hoanganhduc.c50_agent preflight
   python -m course_hoanganhduc.c50_agent list-classrooms --org my-org
   python -m course_hoanganhduc.c50_agent sync --org my-org --classroom short-name --db students.db
   python -m course_hoanganhduc.canvas_agent list-members --course-id 12345
   python -m course_hoanganhduc.gclass_agent list-courses
   python -m course_hoanganhduc.db_agent search "Nguyen"

Allowlists (agent mode, fail-closed when a course/org id is required):

- ``CLASSROOM50_ORG_ALLOWLIST`` — Classroom50 org short-names
- ``CANVAS_COURSE_ALLOWLIST`` — Canvas course IDs
- ``GCLASS_COURSE_ALLOWLIST`` — Google Classroom course IDs
- ``GCLASS_ACCOUNT_ALLOWLIST`` — exact primary emails for the narrow Classroom
  administrator smoke-test path

Comma-separated values. Empty allowlist with a required id fails closed.

Skills in `ai-agents-skills` (profile ``course-management``) route through these
modules: ``classroom50``, ``course-canvas``, ``course-google-classroom``,
``course-db``.

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

The organization field in VCF exports can be customized via ``UNIVERSITY_NAME``
in ``config.json``.

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
- ``GOOGLE_CLASSROOM_GRADE_CATEGORY_METHOD``
- ``CREDENTIALS_PATH``
- ``TOKEN_PATH``
- ``GOOGLE_SHEET_URL``
- ``GOOGLE_SHEET_LECTURER_TOPICS_URL``
- ``GOOGLE_SHEET_STUDENT_REGISTRATION_URL``
- ``INTERNSHIP_REGISTRATION_SHEET_URL``

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

Google Classroom workflows
--------------------------

Administrator command names
~~~~~~~~~~~~~~~~~~~~~~~~~~~

Installing the package creates two console commands. ``course`` remains the main,
legacy course-management CLI; Google Classroom assignment creation is deliberately
isolated in ``course-gclass-admin``. The administrator console command and its Python
module invocation are equivalent:

.. code-block:: bash

   course-gclass-admin --help
   python3 -m course_hoanganhduc.gclass_admin_cli --help

When the per-user virtual environment installed by ``make install`` is not activated,
use either executable from that environment explicitly:

.. code-block:: bash

   ~/.course_venv/bin/course-gclass-admin --help
   ~/.course_venv/bin/python -m course_hoanganhduc.gclass_admin_cli --help

Create an assignment (interactive administrator)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Assignment creation uses the dedicated ``course-gclass-admin`` command. It does not
reuse the legacy CLI's broad scopes, ``token.pickle``, or
``GOOGLE_CLASSROOM_CREDENTIALS``/``GOOGLE_CLASSROOM_TOKEN`` variables, and
``gclass_agent`` explicitly refuses this operation.

Start with an offline preview. This reads only the assignment spec; it does not
resolve or inspect credential paths, require an account, open OAuth, write a file, or
call Google:

.. code-block:: bash

   course-gclass-admin create-assignment \
     --course-id d:my-course \
     --spec sample/google_classroom/assignment-minimal.sample.json \
     --dry-run

The supported execution modes have intentionally different safety boundaries:

.. list-table::
   :header-rows: 1

   * - Mode
     - Credentials and network
     - Confirmation and permitted result
   * - ``--dry-run``
     - None; parses only the local assignment spec.
     - No prompt. Produces the canonical operation preview and makes no mutation.
   * - Interactive creation
     - Verifies the OAuth account and active canonical course.
     - One readable ``Create assignment? [y/N]`` summary. Supports the complete
       documented assignment surface.
   * - ``--yes``
     - Requires an existing coursework token, never opens a browser, and refreshes
       the token automatically when needed.
     - No prompt. Limited to ``DRAFT`` assignments without Drive-sharing effects.
   * - Agent-safe draft
     - Requires an existing token, explicit agent mode, and exact account/course
       allowlists; never opens a browser.
     - A separately approved digest permits only the minimal draft smoke-test shape,
       with duplicate detection and strict read-back.

Authorize the exact primary email address once. The initial Google browser consent is
still necessary because it grants the application the coursework-write,
course-lookup, and identity scopes; it is Google's authorization boundary, not a
package confirmation. The CLI no longer asks the user to type ``AUTH <email>``,
``CREATE <digest>``, or
``SHARE <digest>``. Normal creation has only the readable ``y/N`` prompt shown above,
while token replacement is selected explicitly with ``authorize --replace-token``.

Subsequent creation commands reuse the stored token. If it is expired and has a
refresh token, the command refreshes it before contacting Classroom and stores the
updated JSON atomically; no browser is opened. A revoked, non-refreshable, or otherwise
invalid token fails closed and must be replaced explicitly:

.. code-block:: bash

   course-gclass-admin authorize --account teacher@example.edu

   course-gclass-admin create-assignment \
     --account teacher@example.edu \
     --course-id d:my-course \
     --spec sample/google_classroom/assignment-minimal.sample.json

For a headless remote host, add ``--no-open-browser`` and leave that command running.
After browser consent, the browser's ``127.0.0.1`` redirect may not reach the remote
listener. Read the port from the new redirect URL and, in a second terminal on the
remote host, run:

.. code-block:: bash

   course-gclass-admin complete-loopback --port PORT

Paste the complete redirect URL only at its hidden prompt. The helper validates the
exact ``http://127.0.0.1:PORT/`` boundary and the singular OAuth code/state, uses a
direct loopback connection with no proxy or redirect, and does not place the URL in
shell history or process arguments. Never reuse an earlier callback URL.

Interactive creation displays one control-character-safe summary containing the
verified account, canonical course name and ID, assignment title, release mode,
material/rubric counts, and Drive-sharing modes, then asks ``Create assignment?
[y/N]``. It resolves a course alias with ``courses.get`` and sends the canonical
course ID to the create endpoint. The exported assignment state machine performs the
same ``d:``/``p:`` alias resolution before its first write.

For trusted local automation with a pre-existing coursework token, ``--yes`` skips
that prompt only when the frozen plan is a ``DRAFT`` with no Drive-sharing effect:

.. code-block:: bash

   course-gclass-admin create-assignment \
     --account teacher@example.edu \
     --course-id 123456789012 \
     --spec sample/google_classroom/assignment-minimal.sample.json \
     --yes

Published, scheduled, and Drive-sharing assignments remain interactive. Agent mode
is refused in both normal paths.

A separately scoped smoke-test path exists for a cooperative, explicitly authorized
agent. The ordinary ``gclass_agent`` entrypoint continues to refuse creation. The
smoke-test path accepts only a minimal no-material/no-rubric draft and requires an
existing token, explicit agent mode, exact account/course allowlists, a canonical
course ID, ``--yes``, and an approval digest. First prepare the envelope:

.. code-block:: bash

   COURSE_AGENT_MODE=1 \
   GCLASS_ACCOUNT_ALLOWLIST=teacher@example.edu \
   GCLASS_COURSE_ALLOWLIST=123456789012 \
   course-gclass-admin prepare-agent-safe-draft \
     --account teacher@example.edu \
     --course-id 123456789012 \
     --spec sample/google_classroom/assignment-test-draft.sample.json

Preparation may refresh the isolated token and metadata, but it performs no Google
Classroom mutation. It is not authorization. After a user or trusted harness approves
the exact emitted envelope, run ``create-assignment`` with the same inputs plus
``--agent-safe-draft --yes --expect-approval-digest DIGEST``. The command serializes
the all-state duplicate scan through read-back with a per-token lock. It reuses one
identical developer-associated draft, blocks every collision, and never retries an
ambiguous create or failed read-back.

These checks prevent accidental use through the restricted entrypoint but are not an
OS security boundary: code running under the same user can import Python functions,
alter environment variables, and access that user's token files. Put credentials
behind a separate OS identity, keychain, or user-presence broker if hostile same-user
agents are in scope.

Live safety and terminal outcomes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Before either administrator create path writes coursework, it requests only
``id,name,courseState`` for the course and requires an explicit ``ACTIVE`` state. A
missing state fails closed. The interactive path resolves ``d:``/``p:`` aliases to
Google's canonical ID; the agent-safe path forbids aliases and verifies that the
lookup returns the exact allowlisted canonical ID. The agent-safe path additionally
projects only the fields required for its all-state duplicate scan, create receipt,
and final fetch. It accepts a result only when the
stored item is a developer-associated ``DRAFT`` whose title, optional description,
assignment type, all-student targeting, submission-modification mode, empty materials,
and absent schedule, deadline, points, topic, and grading period exactly match the
approved shape. The success receipt is emitted only after that fetch-by-ID check.

Mutation calls use ``execute(num_retries=0)`` to disable
``google-api-python-client`` retries. The dedicated one-shot HTTP transport also
disables response-triggered OAuth replay, connection/read/status retries, and redirect
following. Thus a Classroom mutation is sent at most once by the supported transport;
normal access-token refresh may still occur before that request.

An ambiguous transport failure, HTTP 408/429/5xx response, malformed successful create
response, or missing created ID is reported as ``error: outcome_unknown`` with exit
status 3. This is terminal: inspect Classroom before deciding whether to issue a new
command, because create has no caller idempotency key. A create whose ID is known but
whose response or strict read-back cannot be validated is reported as
``error: partial_create`` with exit status 4 and the recovery coursework ID when
available. Neither case is automatically retried or deleted.

Assignment spec
^^^^^^^^^^^^^^^

The spec is a strict UTF-8 JSON object. Duplicate keys, unknown fields, symlinks,
non-regular files, and files over 1 MiB are rejected. The stable REST v1 assignment
fields map as follows:

.. list-table::
   :header-rows: 1

   * - JSON field
     - Accepted value and behavior
   * - ``title``
     - Required UTF-8 text, 1--3000 characters, not whitespace-only.
   * - ``description``
     - Optional UTF-8 text up to 30,000 characters; omit or use ``null`` for none.
   * - ``state``
     - ``DRAFT`` (safe default) or ``PUBLISHED``. Scheduling uses ``DRAFT`` plus
       ``scheduled_at``.
   * - ``scheduled_at``
     - Optional timezone-aware RFC 3339 timestamp at least 60 seconds in the future.
   * - ``due_at``
     - Optional timezone-aware RFC 3339 timestamp in the future and, if scheduled,
       later than ``scheduled_at``. It is serialized as matching UTC ``dueDate`` and
       ``dueTime`` fields.
   * - ``max_points``
     - Optional non-negative integer. Zero means ungraded.
   * - ``materials``
     - Optional array of at most 20 writable materials. Omitted, ``null``, or ``[]``
       all mean no attachment.
   * - ``assignee``
     - Optional object. ``mode`` is ``ALL_STUDENTS`` (default) or
       ``INDIVIDUAL_STUDENTS``; the latter requires a non-empty, duplicate-free
       ``student_ids`` array.
   * - ``submission_modification_mode``
     - ``MODIFIABLE_UNTIL_TURNED_IN`` (default) or ``MODIFIABLE``.
   * - ``topic_id``
     - Optional existing topic ID.
   * - ``grading_period``
     - Optional object with ``mode`` ``AUTO`` (omit the API field), ``NONE`` (send an
       empty ID), or ``EXPLICIT`` (requires ``id``).
   * - ``rubric``
     - Optional inline scored or unscored rubric; omit or use ``null`` for none.

Writable material forms are deliberately explicit:

.. code-block:: json

   {
     "materials": [
       {"type": "drive_file", "file_id": "DRIVE_ID", "share_mode": "VIEW"},
       {"type": "link", "url": "https://example.edu/guide"},
       {"type": "youtube", "video_id": "VIDEO_ID"}
     ]
   }

Drive ``share_mode`` may be ``VIEW``, ``EDIT``, or ``STUDENT_COPY``. This command
attaches an existing Drive file ID; it does not upload or search Drive. Link inputs
must be absolute HTTPS URLs without embedded credentials. Google Forms, Gemini Gems,
and NotebookLM notebooks are read-only material variants in stable REST v1 and are
therefore rejected on create.

An inline rubric has ``scoring_mode`` ``SCORED`` or ``UNSCORED`` and a ``criteria``
array. A rubric allows 1--50 criteria and 1--10 levels per criterion. In a scored
rubric every level has finite, non-negative, distinct ``points`` within its criterion
and levels are ordered monotonically by points. In an unscored rubric every level
omits ``points`` and has a non-empty title. A spreadsheet-sourced rubric and preview
learning goals are intentionally outside this stable surface.

See ``sample/google_classroom/assignment-full.sample.json`` for every supported
field. Replace its IDs and timestamps before use. The implementation follows the
`CourseWork resource <https://developers.google.com/workspace/classroom/reference/rest/v1/courses.courseWork>`_,
`Material resource <https://developers.google.com/workspace/classroom/reference/rest/v1/Material>`_,
and `rubric limitations <https://developers.google.com/workspace/classroom/rubrics/limitations>`_.

Rubric creation is staged: create the assignment as a draft, create its rubric, then
publish or schedule only if the rubric succeeded. Stable rubric mutation can also
depend on Google Workspace for Education licensing. If a later rubric or release
stage fails, the command reports the surviving assignment ID and does not delete it
automatically. Scheduled-time lead and due-time constraints are checked again
immediately before the staged release; if time elapsed during draft and rubric
creation makes them stale, the assignment remains a draft and its recovery IDs are
reported.

Coursework credentials
^^^^^^^^^^^^^^^^^^^^^^

Create a Desktop-app OAuth client in Google Cloud, enable the Classroom API, and add
this fixed scope set to the consent configuration:

- ``https://www.googleapis.com/auth/classroom.coursework.students``
- ``https://www.googleapis.com/auth/classroom.courses.readonly``
- ``https://www.googleapis.com/auth/userinfo.email``

The identity scope exposes only the authenticated account's primary email. The
broader Classroom ``profile.emails`` scope is deliberately not requested. Google can
additionally report the OpenID Connect ``openid`` and ``email`` identity scopes. The
command accepts only those documented additions while still requiring every
requested scope and rejecting every unrelated extra scope. The token records the
actual reported grant set; tokens without valid grant evidence must be reauthorized.

Credential resolution uses the first applicable source:

1. ``--credentials`` / ``--token`` command flags. Paths must be absolute and outside
   Git worktrees. A custom client path requires a token path; token-only mode is
   allowed for an already-authorized token.
2. ``COURSE_GCLASS_CREDENTIALS`` / ``COURSE_GCLASS_COURSEWORK_TOKEN`` environment
   variables, with the same pairing rule.
3. Platform defaults:

   - Linux: ``$XDG_CONFIG_HOME/course/google-classroom`` or
     ``~/.config/course/google-classroom``.
   - macOS: ``~/Library/Application Support/course/google-classroom``.
   - Windows path resolution uses ``%APPDATA%\\course\\google-classroom``, but native
     Windows authorization and mutation currently fail closed until ACL and reparse
     protections have native verification.

The default OAuth client is ``credentials.json``. Tokens are JSON files under
``tokens/`` named with the first 20 hexadecimal characters of the expected account's
SHA-256 digest. Metadata contains only hashes/fingerprints and timestamps. The code
does not inspect ``GOOGLE_CLASSROOM_CREDENTIALS``, ``GOOGLE_CLASSROOM_TOKEN``, or
``token.pickle`` for this workflow.

On POSIX systems, every client/token/metadata file must be a regular, non-symlink
file owned by the current user with no group or other permissions. Credential path
components must be owned by the current user or root and must not be writable by
other users (root-owned sticky temporary directories are the one exception). Set
permissions before authorization:

.. code-block:: bash

   chmod 700 ~/.config/course ~/.config/course/google-classroom
   chmod 600 ~/.config/course/google-classroom/credentials.json

Tokens are written atomically with mode ``0600`` while a nonblocking per-token lock
is held. Input OAuth and token endpoints are ignored and replaced with Google's
canonical HTTPS endpoints; service-account files are rejected. OAuth uses PKCE and a
random ``127.0.0.1`` loopback port with a bounded wait. Authorization explicitly
requests offline consent. A fresh or replacement authorization is stored only if
Google returns a durable refresh token; otherwise an existing token remains
unchanged. Dependency loggers that can emit the callback request or token response
are suppressed for the duration of the interactive exchange. The authenticated
Google OAuth UserInfo primary email must be verified and exactly match ``--account``
before a token is stored or a course is inspected.

Inspect resolution and file safety without network access, refresh, OAuth, or writes:

.. code-block:: bash

   course-gclass-admin auth-status --account teacher@example.edu
   course-gclass-admin auth-status \
     --account teacher@example.edu \
     --credentials /absolute/path/credentials.json \
     --token /absolute/path/coursework-token.json \
     --show-paths

Without ``--show-paths``, status output identifies files by basename and reports the
selected source (``cli``, ``environment``, or ``default``) without printing absolute
locations. Add ``--show-paths`` only when the full resolved client and token paths are
needed for troubleshooting. Status remains offline and does not prove that Google
currently accepts the token or that the expected account is signed in.

Status distinguishes safe token syntax from local usability. An expired token with
no refresh token is reported unusable and returns a nonzero status; offline status
does not claim that the account is currently authenticated. Readiness is
conservative near expiry and rejects malformed access-only/no-expiry and
refresh-only/future-expiry combinations that the online Google credential loader
cannot use.

Use ``authorize --replace-token`` only when intentionally replacing an existing
token. Without that explicit flag, an existing token is left untouched; replacement
still requires an interactive Google OAuth flow. Existing legacy tokens cannot be
migrated safely and must be authorized afresh.

Command routing and separate credentials
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

``course-gclass-admin`` currently owns assignment preview, coursework authorization
and status, and assignment creation only. The existing ``course`` CLI retains course
and student listing, roster sync, grading, unenrollment, and submission/Drive download.
The restricted ``python -m course_hoanganhduc.gclass_agent`` entrypoint exposes only
preflight, course/student listing, and roster sync; it continues to refuse creation,
grading, unenrollment, and downloads.

The two credential profiles are intentionally separate and are not interchangeable.
The coursework token uses strict JSON, account-bound storage, and the fixed narrow
scope set documented above; it does not grant the Classroom roster/profile,
Drive-download, topic, or Sheets permissions used by the legacy workflows. The legacy
``token.pickle`` is not accepted by ``course-gclass-admin``.

Offline verification and recorded smoke result
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

From the repository root after ``make install``, the assignment builder, credential
resolver and one-shot transport, administrator CLI, restricted-agent regression,
Classroom50 regression, and CLI parser can be tested locally without contacting
Google:

.. code-block:: bash

   ~/.course_venv/bin/python scripts/test_gclass_coursework.py
   ~/.course_venv/bin/python scripts/test_gclass_coursework_auth.py
   ~/.course_venv/bin/python scripts/test_gclass_admin_cli.py
   ~/.course_venv/bin/python scripts/test_course_agents.py
   ~/.course_venv/bin/python scripts/test_classroom50.py
   ~/.course_venv/bin/python scripts/test_cli_flags.py
   ~/.course_venv/bin/python -m compileall -q course_hoanganhduc

The current offline run passes 116 tests across the first five test scripts, all 215
CLI flag parse cases, and the package compilation check.

The recorded opt-in live smoke test used
``sample/google_classroom/assignment-test-draft.sample.json``. It created one minimal
``DRAFT`` with no attachment, rubric, schedule, deadline, points, topic, or targeted
students, then fetched it by coursework ID and passed the strict field comparison.
The documentation intentionally omits the account, token location, course ID, and
coursework ID.

Legacy roster, grading, download, and unenroll workflows
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sync Google Classroom roster into the local database:

.. code-block:: bash

   course --sync-google-classroom
   course -sgc

Notes:

- Canvas and Google Classroom score sync normalizes grades to a 10-point scale when max points are available.
- Student ID inference for MAT Excel updates only works with VNU University of Science, Hanoi email format.
- ``GOOGLE_CLASSROOM_GRADE_CATEGORY_METHOD`` controls topic/category aggregation (average, sum, weighted; default: average).
- ``GOOGLE_CLASSROOM_CC_TOPICS``, ``GOOGLE_CLASSROOM_GK_TOPICS``, ``GOOGLE_CLASSROOM_CK_TOPICS`` map CC/GK/CK to Google Classroom topic names (comma-separated).
- Required Google APIs: Classroom, Drive (submission downloads), Sheets (Google Sheet imports).
- If you enable new APIs or scopes, delete ``token.pickle`` or re-run to re-auth.
- If topic mappings are not set, the tool auto-matches topic names using phrases like "Chuyên cần", "Giữa kỳ/giữa kì", "Cuối kỳ/cuối kì".
- When Canvas and Google Classroom grades conflict, reports include both sources and interactive flows prompt which source to use.

Setup quick steps:

- Create/choose a Google Cloud project.
- Enable APIs: Classroom API, Drive API, Google Sheets API.
- Configure OAuth consent screen (external/internal as needed).
- Create OAuth client credentials (Desktop app) and download ``credentials.json``.
- Place ``credentials.json`` and (after first run) ``token.pickle`` in the config folder or pass paths via CLI.

Grade Google Classroom assignments:

.. code-block:: bash

   course --grade-google-classroom
   course --grade-google-classroom --gc-coursework-id 1234567890 --gc-grade-score 8

Notes:

- ``--gc-include-graded`` includes already graded submissions.
- ``--gc-apply-all`` grades all listed submissions without selection.

Download latest Google Classroom submissions and run checks:

.. code-block:: bash

   course --download-google-classroom-submissions
   course --download-google-classroom-submissions --gc-download-coursework-id 1234567890 --gc-download-dest-dir ./gclassroom_submissions

Unenroll Google Classroom students by email domain:

.. code-block:: bash

   course --unenroll-google-classroom --gc-unenroll-domain gmail.com
   course --unenroll-google-classroom --gc-unenroll-domain gmail.com,outlook.com --gc-unenroll-all
   course --unenroll-google-classroom --gc-unenroll-email student1@gmail.com,student2@outlook.com
   course --unenroll-google-classroom --gc-unenroll-select
   course --unenroll-google-classroom --gc-unenroll-missing-student-id
   course --unenroll-google-classroom --gc-unenroll-domain gmail.com --dry-run

Notes:

- Successful unenroll removes matching students from the local database (by Email or Google_ID).

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

Import students from a Google Sheet URL:

.. code-block:: bash

   course --add-google-sheet "https://docs.google.com/spreadsheets/d/FILE_ID/edit#gid=0"

Notes:

- Google Sheets API must be enabled for the project tied to your credentials.

Notes:

- MAT*.xlsx imports ignore score columns (CC, GK, CK, totals); only roster fields are imported.

Import internship registrations (Skills/Wishlist):

.. code-block:: bash

   course --import-internships "https://docs.google.com/spreadsheets/d/FILE_ID/edit#gid=0"
   course --import-internships sample/internship_data.csv
   course --import-registrations "https://docs.google.com/spreadsheets/d/FILE_ID/edit#gid=0"
   course --import-registrations sample/internship_registrations_data.csv

Manage company contact data in ``companies.db``:

.. code-block:: bash

   course --import-companies "Danh sách công ty liên hệ (không xóa).xlsx"
   course --import-companies "https://docs.google.com/spreadsheets/d/FILE_ID/edit#gid=0"
   course --export-companies companies_export.xlsx

Notes:
- Internship and Company imports support both Vietnamese and English column headers.

List students by email domain:

.. code-block:: bash

   course --list-email-domain gmail.com,outlook.com

List students with duplicate names (or display names):

.. code-block:: bash

   course --list-duplicate-names
   course --list-duplicate-names --duplicate-name-field google
   course --list-duplicate-names --duplicate-name-field canvas --duplicate-name-format csv --duplicate-name-output duplicate_canvas
   course --list-duplicate-names --duplicate-name-field "Some Custom Field" --duplicate-name-format json --duplicate-name-output dupes.json

Notes:

- ``--duplicate-name-field`` accepts ``name``, ``google``, ``canvas``, or a custom field name.
- ``--duplicate-name-format`` supports ``txt``, ``csv``, or ``json`` (default: ``txt``).

List students missing IDs:

.. code-block:: bash

   course --list-missing-ids
   course --list-missing-ids google,canvas --missing-ids-format csv --missing-ids-output missing_ids.csv

List Google Classroom students:

.. code-block:: bash

   course --list-google-students
   course --list-google-students --google-course-id 1234567890

List students by submission status:

.. code-block:: bash

   course --list-submission-status google:turned_in
   course --list-submission-status canvas:submitted
   course --list-submission-status google:NEW@Quiz 1
   course --list-submission-status google:NEW@Quiz 1,Quiz 2

Notes:

- Google Classroom values: ``NEW``, ``CREATED``, ``TURNED_IN``, ``RETURNED``, ``RECLAIMED_BY_STUDENT``.
- Canvas values: ``UNSUBMITTED``, ``SUBMITTED``, ``GRADED``, ``PENDING_REVIEW``, ``COMPLETE``.

Combine listing filters (AND semantics):

.. code-block:: bash

   course --list-email-domain gmail.com --list-missing-ids student
   course --list-email-domain gmail.com --list-submission-status google:turned_in
   course --list-missing-ids student --list-submission-status google:NEW
   course --list-email-domain gmail.com --list-missing-ids student --list-submission-status google:CREATED
   course --list-submission-status canvas:UNSUBMITTED --list-missing-ids student
   course --list-email-domain gmail.com --list-submission-status google:TURNED_IN@Quiz 1
   course --list-email-domain gmail.com --list-submission-status google:TURNED_IN@Quiz 1,Quiz 2

Notes:

- ``--list-duplicate-names`` and Canvas/Google roster listings cannot be combined with other listing flags.
- Export options (like ``--missing-ids-output``) are ignored in combined mode; the merged list prints to the console.

Unenroll Canvas students:

.. code-block:: bash

   course --unenroll-canvas --canvas-unenroll-domain gmail.com
   course --unenroll-canvas --canvas-unenroll-email student1@gmail.com,student2@outlook.com
   course --unenroll-canvas --canvas-unenroll-select
   course --unenroll-canvas --canvas-unenroll-missing-student-id
   course --unenroll-canvas --canvas-unenroll-domain gmail.com --dry-run

Notes:

- Successful unenroll removes matching students from the local database (by Email or Canvas ID).

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
- If Classroom50 preflight fails, confirm ``gh`` auth and that the Classroom50
  teacher extension is installed.
- If an agent entrypoint exits with ``allowlist_required`` / ``not_allowlisted``,
  set the matching ``*_ALLOWLIST`` environment variable.
- Agent entrypoints refusing an operation is intentional; use the full
  ``course`` CLI as a human for downloads, unenroll, grading, and DB mutations.
