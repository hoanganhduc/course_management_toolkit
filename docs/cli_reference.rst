CLI Reference
=============

This reference is generated from the ``course`` argument parser
(``course_hoanganhduc/core.py`` plus Classroom50 flags from
``course_hoanganhduc/c50_flags.py``) and lists CLI flags grouped by section.

Agent-restricted entrypoints (``python -m course_hoanganhduc.c50_agent``,
``canvas_agent``, ``gclass_agent``, ``db_agent``) are documented under
:doc:`usage` and are **not** the full interactive ``course`` surface.

Automation
----------

- ``--generate-weekly-workflow``: Generate a sample GitHub Actions workflow for weekly automation
- ``--run-weekly-automation``: Run weekly automation for a closed assignment
- ``--run-weekly-local``: Run weekly automation locally and archive reports
- ``--weekly-assignment-id``: Canvas assignment ID for weekly automation
- ``--weekly-category``: Assignment group/category filter for missing-submission reminders
- ``--weekly-dest-dir``: Output directory for weekly downloads
- ``--weekly-local-root``: Local folder for weekly report archiving (default: cwd)
- ``--weekly-meaningful-threshold``: Meaningfulness threshold for weekly automation
- ``--weekly-notify-missing``: Send reminders for missing submissions after due date
- ``--weekly-ocr-lang``: OCR language for weekly automation
- ``--weekly-ocr-service``: OCR service for weekly automation
- ``--weekly-refine``: AI refinement method for weekly notices
- ``--weekly-score``: Score to assign to clean submissions (default: 10)
- ``--weekly-similarity-threshold``: Similarity threshold for weekly automation
- ``--weekly-teacher-canvas-id``: Canvas user ID for summary notifications
- ``--workflow-assignment-id``: Assignment ID placeholder for workflow
- ``--workflow-course-code``: Course code placeholder for workflow
- ``--workflow-course-id``: Course ID placeholder for workflow
- ``--workflow-students-branch``: Deprecated alias for --workflow-toolkit-branch
- ``--workflow-students-repo``: Deprecated alias for --workflow-toolkit-repo
- ``--workflow-teacher-canvas-id``: Teacher Canvas ID placeholder for workflow
- ``--workflow-toolkit-branch``: Branch for course toolkit repo
- ``--workflow-toolkit-repo``: Repo URL for course toolkit

Canvas: Admin Tools
-------------------

- ``--canvas-deadline-category`` (``-cdc``): Assignment category (group) to filter when changing deadlines
- ``--canvas-lock-category`` (``-clc``): Assignment category (group) to filter when changing lock dates
- ``--change-canvas-deadlines`` (``-ccd``): Change deadlines for one or more Canvas assignments (provide assignment IDs, or leave blank to select interactively)
- ``--change-canvas-lock-dates`` (``-ccl``): Change lock dates (lock_at) for one or more Canvas assignments (provide assignment IDs, or leave blank to select interactively)
- ``--create-canvas-groups``: Create groups in a Canvas course group set
- ``--delete-empty-canvas-groups`` (``-deg``): Delete all empty groups (groups with no members) from a Canvas course group set
- ``--grade-resubmission`` (``-grs``): List resubmissions and regrade them (optionally provide assignment IDs)
- ``--group-name-pattern``: Pattern for group names, e.g., 'Group {i}' (default: 'Group {i}')
- ``--group-set-id``: Canvas group set ID to create groups in (leave blank to select interactively)
- ``--keep-old-grade``: Keep previous grade for resubmissions without prompting
- ``--new-canvas-due-date`` (``-ncd``): New due date for Canvas assignments (format: YYYY-MM-DD HH:MM)
- ``--new-canvas-lock-date`` (``-ncl``): New lock date for Canvas assignments (format: YYYY-MM-DD HH:MM)
- ``--no-restricted`` (``-nres``): Disable restricted mode for grading Canvas assignments (list all assignments with submissions and all students who submitted)
- ``--num-groups``: Number of groups to create (default: 5)

Canvas: People and Communication
--------------------------------

- ``--add-canvas-announcement`` (``-aa``): Create a new announcement in Canvas course
- ``--announcement-file``: TXT file with announcement body
- ``--announcement-message``: Short message for Canvas announcement
- ``--announcement-title``: Title for Canvas announcement
- ``--canvas-assignment-category`` (``-cac``): Assignment category (group) to filter when listing Canvas assignments
- ``--canvas-course-id`` (``-cc``): Canvas course ID (overrides default)
- ``--canvas-unenroll-all``: Unenroll all matched Canvas students without selection
- ``--canvas-unenroll-domain``: Email domain(s) to match for Canvas unenroll (comma-separated)
- ``--canvas-unenroll-email``: Email(s) to match for Canvas unenroll (comma-separated)
- ``--canvas-unenroll-missing-student-id``: Unenroll Canvas students missing Student ID in local database
- ``--canvas-unenroll-select``: Select Canvas students to unenroll from a list
- ``--comment-canvas-submission`` (``-cs``): Add a comment to a Canvas assignment submission
- ``--download-canvas-assignment`` (``-da``): Download all submission files for a Canvas assignment (optionally provide assignment ID)
- ``--download-dest-dir`` (``-dd``): Destination directory for downloaded Canvas assignment files
- ``--edit-canvas-pages`` (``-ep``): List and edit Canvas course pages
- ``--fetch-canvas-messages`` (``-fm``): Fetch and reply to Canvas inbox messages
- ``--grade-canvas-assignment`` (``-ga``): Grade Canvas assignment submissions interactively
- ``--invite``: Invite students after --add-file import (skips those already enrolled)
- ``--invite-canvas-email`` (``-ie``): Invite a single user to Canvas course by email
- ``--invite-canvas-file`` (``-if``): Invite multiple users to Canvas course from a TXT file or string of pairs/emails
- ``--invite-canvas-name``: Name for Canvas invite (for single user)
- ``--invite-canvas-role`` (``-ir``): Role for Canvas invite (student/teacher/ta, default: student)
- ``--invite-role``: Role for --invite (student/teacher/ta, default: student)
- ``--invite-section``: Default Canvas section name for --invite (optional)
- ``--list-canvas-assignments``: List all assignments on Canvas LMS
- ``--list-canvas-members`` (``-cm``): List all members (teachers, TAs, students) of a Canvas course
- ``--list-multiple-submissions-on-time`` (``-lm``): List students who submitted twice or more to an assignment and the first submission is on time (optionally provide assignment ID)
- ``--notify-incomplete-reviews`` (``-nr``): Find and notify students who have not completed required peer reviews for a Canvas assignment
- ``--review-assignment-id`` (``-rai``): Canvas assignment ID for peer review notification
- ``--search-canvas-user`` (``-cu``): Search for a user in Canvas by name or email
- ``--sync-canvas`` (``-sc``): Sync Canvas course members to local database
- ``--unenroll-canvas``: Unenroll Canvas students by domain/email/select/missing-id

Canvas: Rubrics and Grading
---------------------------

- ``--add-canvas-grading-scheme`` (``-ags``): Add a grading scheme to Canvas course from JSON file
- ``--check-student-submission-similarity`` (``-css``): Check similarities between submissions of the same student for different assignments. Optionally provide a Canvas student ID or a comma-separated list of IDs. If not provided, will prompt for selection interactively.
- ``--export-canvas-grading-scheme`` (``-egs``): List and export Canvas grading schemes (grading standards) to JSON
- ``--export-canvas-rubrics`` (``-er``): Export Canvas rubrics to TXT/CSV file
- ``--final-evals-announce`` (``-fea``): Also create a course announcement after sending final evaluations.
- ``--final-evals-course-id`` (``-fec``): Canvas course ID to use when sending final evaluations (overrides default CANVAS_LMS_COURSE_ID).
- ``--import-canvas-rubrics`` (``-imr``): Import rubrics from TXT/CSV file to Canvas course
- ``--list-canvas-rubrics`` (``-lr``): List all unique rubrics used in Canvas course
- ``--rubric-assignment-id`` (``-rid``): Assignment ID to filter rubrics
- ``--send-final-evaluations`` (``-sfe``): Send final evaluation results to students via Canvas. Optionally provide directory with evaluation files (default: final_evaluations).
- ``--update-canvas-rubric-id`` (``-uri``): Rubric ID to associate with assignments (leave blank to select interactively)
- ``--update-canvas-rubrics`` (``-ur``): Update rubric for one or more Canvas assignments (provide assignment IDs, or leave blank to select interactively)

Classroom50
-----------

Assignment administration (separate command)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Installing the package creates ``course-c50-admin``
(``course_hoanganhduc.c50_admin_cli:main``), the dedicated Classroom50 operator CLI.
The equivalent source-tree invocation is
``python -m course_hoanganhduc.c50_admin_cli``. This mirrors the
``course-gclass-admin`` boundary: the ``course`` flags listed below keep their read,
roster-sync, export, and human-download workflows, while every operation that changes
remote Classroom50 state lives on the separate command. The restricted
``python -m course_hoanganhduc.c50_agent`` entrypoint continues to expose
``preflight``, ``list-*``, allowlisted ``sync``, and ``export``; it explicitly refuses
``assignment-add``, ``assignment-remove``, ``invite``, and ``download``.

Every subcommand refuses in agent mode, requires an interactive terminal unless
``--dry-run`` is given, and builds one fixed ``gh teacher`` command line. There is no
free-form argument passthrough. The four subcommands are:

- ``assignment-add --org ORG --classroom SHORT-NAME --slug SLUG --name NAME``
  registers an assignment. Optional ``--template``, ``--tests``, ``--empty-repo``,
  ``--mode {individual,group}``, ``--max-group-size N`` (required with ``--mode
  group``), ``--available-from``, ``--due``, ``--feedback-pr``, ``--allowed-files``,
  ``--pass-threshold``, ``--student-permission``, ``--description``, ``--locked``.
  ``gh teacher assignment add`` replaces an existing entry in place and the pinned CLI
  cannot set submission mode, so re-running it restores the default every-push mode and
  discards a tagged-commit setting made in the web form. An existing slug is therefore
  refused; overriding needs both ``--allow-overwrite`` and an interactive confirmation.
- ``assignment-remove --org ORG --classroom SHORT-NAME --slug SLUG`` deletes the entry
  from ``assignments.json``. It does not delete student repositories, and re-adding the
  same slug is not a clean reset; the confirmation states both. An absent slug is
  refused rather than silently succeeding.
- ``invite --target ORG[/REPO] --username LOGIN [--username LOGIN ...]`` sends
  invitations. Use ``--admin`` for an organization owner invitation, or
  ``--permission {pull,triage,push,maintain,admin}`` for a repository invitation; the
  two are mutually exclusive and each is valid for only one target kind. Repository
  invitations are idempotent, organization invitations are not, so organization targets
  are preflighted against ``gh teacher member list`` and logins that are already members
  or hold a pending invitation are skipped. That read needs the ``admin:org`` scope.
- ``download --org ORG --classroom SHORT-NAME --assignment SLUG --dest DIR`` collects
  submissions. ``--by-pattern`` skips the team lookup, so it fetches no ``result.json``
  and writes no ``scores.csv``. It is permitted only after the assignment record reports
  empty-repository mode, and refused otherwise.

Exit codes are ``0`` success, ``2`` validation failure or refusal, ``3`` the ``gh`` call
failed and the remote outcome is unverified, and ``4`` a batch invite completed with
failures. Student acceptance and student submission are not provided: both run under a
student's own credentials.

Toolkit flags
~~~~~~~~~~~~~

- ``--classroom50-assignment``: Classroom50 assignment slug (human download)
- ``--classroom50-classroom``: Classroom50 classroom short-name
- ``--classroom50-download-dest``: Destination directory for Classroom50 downloads
- ``--classroom50-org``: Classroom50 GitHub organization
- ``--classroom50-preflight``: Run Classroom50 auth preflight (whoami)
- ``--classroom50-report``: Write Classroom50 JSON report to PATH
- ``--download-classroom50``: Download student submissions (human CLI only)
- ``--export-classroom50-roster``: Export local roster to Classroom50 CSV (default: classroom50_roster.csv)
- ``--list-classroom50-assignments``: List Classroom50 assignments
- ``--list-classroom50-classrooms``: List Classroom50 classrooms
- ``--list-classroom50-roster``: List Classroom50 roster
- ``--sync-classroom50`` (``-sc50``): Sync Classroom50 roster into local student database

Configuration
-------------

- ``--backup-config``: Back up config.json to a timestamped file (optional: backup dir)
- ``--clear-config`` (``-ccfg``): Delete stored config.json from the default location
- ``--clear-credentials`` (``-ccred``): Delete stored credentials.json and token.pickle from the default location
- ``--config`` (``-cfg``): Load config from JSON file and save to default location
- ``--config-backup-keep``: Number of config backups to retain (default from config)
- ``--course-code`` (``-ccode``): Course code for config folder (e.g., MAT3500)
- ``--detect-local-ai``: Detect locally installed AI models (Ollama-compatible)
- ``--list-ai-models`` (``-lam``): List available AI models for a provider ('gemini', 'huggingface', 'local', or 'all')
- ``--local-llm-args``: Extra args for local LLM command
- ``--local-llm-command``: Command to run the local LLM (default: ollama)
- ``--local-llm-gguf-dir``: Directory to scan for .gguf models (llama.cpp)
- ``--local-llm-model``: Local LLM model name (default from config/settings)
- ``--restore-config``: Restore config.json from a backup (default: latest)
- ``--student-sort-method``: Student sort method for detail outputs (first_last, last_first, id)
- ``--test-ai`` (``-tai``): Test AI services ('gemini', 'huggingface', 'local', or 'all')
- ``--test-ai-gemini-model``: Override Gemini model name for --test-ai all
- ``--test-ai-huggingface-model``: Override HuggingFace model name for --test-ai all
- ``--test-ai-model`` (``-tam``): Override model name for --test-ai (provider-specific)

Course Calendar
---------------

- ``--build-course-calendar``: Build course calendar and export to TXT/MD/ICS
- ``--calendar-course-code``: Course code for calendar title
- ``--calendar-course-name``: Course name for calendar summaries
- ``--calendar-extra-holidays``: Comma-separated extra holiday dates (YYYY-MM-DD,YYYY-MM-DD)
- ``--calendar-extra-week``: Allow one make-up week when holidays skip sessions
- ``--calendar-input``: TXT file with first-week schedule and optional holidays
- ``--calendar-output-base``: Output base name for calendar files (default: course_calendar)
- ``--calendar-output-dir``: Output directory for calendar exports (default: cwd)
- ``--calendar-weeks``: Number of official weeks (default: 15)
- ``--force``: Force Canvas calendar import (do not skip duplicates)
- ``--import-canvas-calendar-ics``: Import an iCal (.ics) file and create Canvas calendar events
- ``--skip-duplicates``: Skip Canvas calendar events that match existing entries

Exams (Multichoice)
-------------------

- ``--evaluate-multichoice-answers``: Evaluate student answers for multiple-choice exam (provide exam type: midterm/final, default: global EXAM_TYPE)
- ``--extract-multichoice-answers`` (``-ema``): Extract student answers from scanned multi-choice exam sheet PDF
- ``--extract-multichoice-solutions`` (``-ems``): Extract multiple-choice exam solutions from PDF (each page is one sheet code)
- ``--sync-multichoice-evaluations`` (``-sme``): Sync multichoice exam evaluations to Canvas assignment (provide exam type: midterm/final, default: global EXAM_TYPE)

General
-------

- ``--dry-run``: Preview actions without writing files or databases
- ``--log-backups``: Number of rotated log files to keep
- ``--log-dir``: Directory for log files (default: config folder)
- ``--log-level``: Logging level (default: INFO)
- ``--log-max-bytes``: Max size in bytes for rotating logs
- ``--refine``: Refine generated messages/announcements with AI
- ``--verbose`` (``-v``): Enable verbose output

Google Classroom
----------------

Assignment administration (separate command)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Installing the package creates two different console commands:

- ``course`` runs the general toolkit CLI
  (``course_hoanganhduc.cli:main``). It is not an alias for assignment
  administration.
- ``course-gclass-admin`` runs the dedicated assignment administration CLI
  (``course_hoanganhduc.gclass_admin_cli:main``). The equivalent source-tree or
  module invocation is
  ``python -m course_hoanganhduc.gclass_admin_cli``.

This is a capability boundary, not a wholesale Google Classroom CLI replacement.
``course-gclass-admin`` owns assignment preview, isolated coursework OAuth, and
assignment creation only. The legacy ``course`` flags listed later in this section
continue to own their read, roster-sync, grading, submission-download, and unenroll
workflows. The restricted
``python -m course_hoanganhduc.gclass_agent`` entrypoint continues to expose
``preflight``, ``list-courses``, ``list-students``, and allowlisted ``sync``; it
explicitly refuses create, grade, download, and unenroll operations. Any future
migration is outside the current implementation; use the command ownership above as
the supported routing.

Use either form consistently in the commands below. The dedicated CLI exposes five
subcommands:

- ``auth-status --account EMAIL`` inspects the isolated coursework credentials
  offline. It does not refresh a token, contact Google, or write files. Add
  ``--show-paths`` to disclose resolved paths; otherwise output contains only safe
  names and fingerprints. ``--credentials`` and ``--token`` override path
  resolution. Without those flags, resolution checks
  ``COURSE_GCLASS_CREDENTIALS`` and ``COURSE_GCLASS_COURSEWORK_TOKEN`` before the
  platform-specific default config directory.
- ``authorize --account EMAIL`` runs installed-app loopback OAuth in an interactive
  terminal. There is no local ``AUTH ...`` phrase to type: Google browser consent
  and account verification are the authorization. ``--no-open-browser`` prints the
  consent URL instead of opening it. ``--replace-token`` is required before an
  existing isolated token can be replaced; ``--credentials`` and ``--token`` select
  explicit files.
- ``complete-loopback --port PORT`` supports authorization on a headless remote
  host. It reads the browser redirect URL from a hidden prompt, validates the exact
  loopback boundary, and delivers it without placing its code or state in shell
  history or process arguments.
- ``prepare-agent-safe-draft --account EMAIL --course-id CANONICAL_ID --spec FILE``
  validates a narrowly constrained draft in explicit agent mode. It requires an
  existing token, exact account and course allowlists, and a canonical course ID;
  it never opens a browser or starts fresh OAuth. It may refresh and persist an
  expired token or verification metadata, so ``credentialStateMayChange`` can be
  true, but ``classroomMutation`` is false. Its JSON approval envelope binds the
  account, OAuth client, token location, course, and operation. Preparation is
  neither Classroom mutation nor approval to mutate.
- ``create-assignment --course-id ID_OR_ALIAS --spec FILE`` validates and creates
  one assignment. ``--dry-run`` needs neither ``--account`` nor credentials. A live
  human run requires ``--account`` and normally displays one readable summary
  followed by ``Create assignment? [y/N]:``; no ``CREATE`` or ``SHARE`` digest phrase
  is required. ``--no-open-browser`` controls fresh interactive OAuth when no token
  exists.

``-y`` / ``--yes`` skips the normal confirmation only when all of the following are
true: the assignment is a ``DRAFT``, it has no Drive-sharing effects, and an isolated
coursework token already exists. This mode disables browser authorization. Published,
scheduled, and Drive-sharing plans still require the normal interactive confirmation.
General mutation remains forbidden in agent mode.

The separate ``--agent-safe-draft`` path permits only a minimal all-students draft
with a title and optional description; attachments, due or scheduled times, points,
topics, grading periods, individual targeting, and rubrics are rejected. It requires
``COURSE_AGENT_MODE=1``, exact ``GCLASS_ACCOUNT_ALLOWLIST`` and
``GCLASS_COURSE_ALLOWLIST`` values, a canonical course ID, an existing token,
``--yes``, and the full ``--expect-approval-digest`` emitted by a separately reviewed
preparation. It does not open a browser. An identical developer-associated draft can
be reused after strict read-back; conflicting same-title work is rejected.

Every Classroom mutation request uses the one-shot transport: response-triggered
OAuth replay, HTTP retries, and redirects are disabled. An ambiguous mutation returns
``error: outcome_unknown`` with exit status 3; inspect Classroom before deciding
whether to run anything again. A known partial result returns ``error: partial_create``
with exit status 4 and available receipt fields. A failed mutation is not replayed or
deleted automatically. No automatic rerun is attempted after an ambiguous initial
create; inspect Classroom manually before deciding what to do next.

Examples (all identities and IDs are placeholders):

.. code-block:: bash

   course-gclass-admin auth-status \
     --account ACCOUNT_EMAIL

   python -m course_hoanganhduc.gclass_admin_cli authorize \
     --account ACCOUNT_EMAIL \
     --no-open-browser

   course-gclass-admin create-assignment \
     --course-id COURSE_ID_OR_ALIAS \
     --spec sample/google_classroom/assignment-minimal.sample.json \
     --dry-run

   course-gclass-admin create-assignment \
     --account ACCOUNT_EMAIL \
     --course-id COURSE_ID_OR_ALIAS \
     --spec sample/google_classroom/assignment-minimal.sample.json \
     --yes

For the agent-safe path, first run preparation and separately review its JSON
envelope. Then pass its full ``approvalDigest`` to the live command:

.. code-block:: bash

   COURSE_AGENT_MODE=1 \
   GCLASS_ACCOUNT_ALLOWLIST=ACCOUNT_EMAIL \
   GCLASS_COURSE_ALLOWLIST=CANONICAL_COURSE_ID \
   course-gclass-admin prepare-agent-safe-draft \
     --account ACCOUNT_EMAIL \
     --course-id CANONICAL_COURSE_ID \
     --spec sample/google_classroom/assignment-test-draft.sample.json

   COURSE_AGENT_MODE=1 \
   GCLASS_ACCOUNT_ALLOWLIST=ACCOUNT_EMAIL \
   GCLASS_COURSE_ALLOWLIST=CANONICAL_COURSE_ID \
   course-gclass-admin create-assignment \
     --account ACCOUNT_EMAIL \
     --course-id CANONICAL_COURSE_ID \
     --spec sample/google_classroom/assignment-test-draft.sample.json \
     --agent-safe-draft \
     --yes \
     --expect-approval-digest APPROVAL_DIGEST

The existing flags below still use the legacy Google credential/token flow. They do
not create assignments.

- ``--download-google-classroom-submissions`` (``-dgcs``): Download latest Google Classroom submissions and run checks
- ``--gc-apply-all``: Apply grading to all listed submissions without selection
- ``--gc-coursework-id``: Google Classroom coursework ID(s) to grade (optional)
- ``--gc-download-coursework-id``: Google Classroom coursework ID(s) to download (optional)
- ``--gc-download-dest-dir``: Download folder for Google Classroom submissions (optional)
- ``--gc-grade-score``: Score to assign to selected submissions (optional)
- ``--gc-include-graded``: Include already graded submissions
- ``--gc-meaningful-threshold``: Meaningfulness threshold for Google Classroom checks
- ``--gc-ocr-lang``: OCR language for Google Classroom checks
- ``--gc-ocr-service``: OCR service for Google Classroom checks (ocrspace/tesseract/paddleocr)
- ``--gc-similarity-threshold``: Similarity threshold for Google Classroom checks
- ``--gc-unenroll-all``: Unenroll all matched Google Classroom students without selection
- ``--gc-unenroll-domain``: Email domain(s) to match for Google Classroom unenroll (comma-separated)
- ``--gc-unenroll-email``: Email(s) to match for Google Classroom unenroll (comma-separated)
- ``--gc-unenroll-missing-student-id``: Unenroll Google Classroom students missing Student ID in local database
- ``--gc-unenroll-select``: Select Google Classroom students to unenroll from a list
- ``--google-course-id`` (``-gci``): Google Classroom course ID (prompts if None)
- ``--google-credentials-path`` (``-gcp``): Path to Google Classroom credentials JSON file
- ``--google-token-path`` (``-gtp``): Path to Google Classroom token pickle file
- ``--grade-google-classroom`` (``-ggc``): Grade Google Classroom assignment submissions
- ``--list-google-courses`` (``-lgc``): List Google Classroom courses for the current account
- ``--list-google-students``: List Google Classroom students for a course (uses --google-course-id or prompts)
- ``--sync-google-classroom`` (``-sgc``): Sync students in the local database with active students from Google Classroom course
- ``--unenroll-google-classroom`` (``-ugc``): Unenroll Google Classroom students by email domain

OCR and PDFs
------------

- ``--add-blackboard-counts`` (``-b``): Extract and add blackboard counts from PDF to database
- ``--export-blackboard-counts`` (``-B``): Export blackboard counts by date for all students to TXT/Markdown file (use .txt or .md extension, default: TXT)
- ``--extract-text`` (``-t``): Extract handwriting text from PDF and save to TXT file
- ``--ocr-lang`` (``-L``): OCR language for PDF extraction (default: auto)
- ``--ocr-service`` (``-O``): OCR service to use for PDF extraction (default: 'ocrspace'). The 'ocrspace' service uses the OCR.space API and works better for handwriting text. The other two services work better for printed text and require additional local installation.
- ``--print-blackboard-counts`` (``-p``): Print blackboard counts by date for all students
- ``--simple-text`` (``-T``): Extract simple text (no layout) from PDF OCR

Student Database
----------------

- ``--add-file`` (``-a``): Import students from Excel, CSV, or PDF file into the database
- ``--add-google-sheet`` (``-gsh``): Import students from Google Sheet URL (optional: URL, default from config)
- ``--all-details`` (``-A``): Show details of all students
- ``--backup-db``: Back up students.db to a timestamped file (optional: backup dir)
- ``--db`` (``-db``): Database file name (default: students.db, saved in script folder)
- ``--db-backup-keep``: Number of database backups to retain (default from config)
- ``--details`` (``-d``): Show details of a student by name, student id, or email
- ``--dry-run-rows``: Number of preview rows to show with --dry-run + --add-file (default: 5)
- ``--duplicate-name-field``: Field to detect duplicates (name/google/canvas or custom field name)
- ``--duplicate-name-format``: Output format for duplicate-name report (default: txt)
- ``--duplicate-name-output``: Output path for duplicate-name report (optional; extension inferred from format if missing)
- ``--export-all-details`` (``-E``): Export all student or company details (including submission attachments) to TXT file
- ``--export-anonymized`` (``-ean``): Export anonymized roster to CSV (optional: output path)
- ``--export-companies``: Export company list to Excel file
- ``--export-emails`` (``-e``): Export all student emails to TXT file (avoids duplicates)
- ``--export-emails-and-names`` (``-en``): Export all student emails and names to TXT file (default: emails_and_names.txt)
- ``--export-excel`` (``-x``): Export student list to Excel file
- ``--export-final-grade-distribution``: Export final grade distribution to a TXT file. Optionally provide output path (default: ./final_grade_distribution.txt).
- ``--export-grade-diff``: Export grade updates to CSV when updating MAT files (optional: output path)
- ``--export-roster`` (``-ero``): Export classroom roster to CSV file (default: classroom_roster.csv)
- ``--export-type`` (``-et``): Entity type to export with --export-all-details (student/company, default: student)
- ``--export-vcf`` (``-vcf``): Export student contact info to VCF (iOS compatible). Default: students_contacts.vcf
- ``--filter-file`` (``-ff``): Filter export by identifiers in this file (TXT/CSV/XLSX)
- ``--generate-final-evaluations`` (``-gfe``): Generate per-student final evaluation TXT reports (optional: output dir, default: ./final_evaluations).
- ``--import-companies``: Import company data from Excel/CSV to companies.db
- ``--import-internships``: Import student internship data from Google Sheet URL or local CSV/Excel file (optional: URL/Path, default from config)
- ``--import-mini-projects``: Import mini-project data from Google Sheets (lecturer topics + student registrations)
- ``--import-progress-reports``: Import student progress reports from their linked Google Sheets
- ``--import-registrations``: Import student registration data from Google Sheet URL or local CSV/Excel file (optional: URL/Path, default from config)
- ``--list-duplicate-names``: List students who share the same full name
- ``--list-email-domain`` (``-led``): List students whose email matches domain(s) (comma-separated, e.g., gmail.com,outlook.com)
- ``--list-missing-ids``: List students missing Google/Canvas/Student IDs (optional: google,canvas,student,all or comma-separated)
- ``--list-submission-status``: List students by submission status (shows attachment details; prefix with google: or canvas:, optional @assignment title)
- ``--load`` (``-l``): Load students from database file
- ``--load-override-grades`` (``-log``): Load override_grades.xlsx and persist overrides to the database (default: override_grades.xlsx).
- ``--mini-project-lecturer-sheet``: Google Sheet URL for lecturer mini-project topics (optional; default from config)
- ``--mini-project-registration-sheet``: Google Sheet URL for student mini-project registrations (optional; default from config)
- ``--missing-ids-format``: Output format for missing-ids report (default: txt)
- ``--missing-ids-output``: Output path for missing-ids report (optional; extension inferred from format if missing)
- ``--modify`` (``-m``): Interactively modify the student database
- ``--restore-db``: Restore students.db from a backup (default: latest)
- ``--save`` (``-s``): Save current students to database file
- ``--search`` (``-S``): Search for students by keyword (name, student id, email, etc.)
- ``--sheet-name`` (``-sn``): Specific sheet name to import from Excel file
- ``--sheet-selection`` (``-ss``): Sheet selection mode: first (default), select (interactive), all, merge
- ``--sync-mat-canvas``: Sync CC/GK/CK scores from MAT*.xlsx to Canvas assignments (uses configured assignment IDs)
- ``--sync-mat-types``: Comma-separated list of score types to sync (CC,GK,CK). Default: all available.
- ``--update-mat-excel`` (``-ume``): Update MAT*.xlsx file(s) with grades from database (provide one or more file paths)
- ``--validate-data``: Validate student data and write a report (optional: output path)
