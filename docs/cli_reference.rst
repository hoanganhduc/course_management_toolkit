CLI Reference
=============

This reference is generated from `course_hoanganhduc/core.py` and lists CLI flags grouped by section.

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
- ``--invite-role`` (``--invite-rote``): Role for --invite (student/teacher/ta, default: student)
- ``--invite-section``: Default Canvas section name for --invite (optional)
- ``--list-canvas-assignments``: List all assignments on Canvas LMS
- ``--list-canvas-members`` (``-cm``): List all members (teachers, TAs, students) of a Canvas course
- ``--list-multiple-submissions-on-time`` (``-lm``): List students who submitted twice or more to an assignment and the first submission is on time (optionally provide assignment ID)
- ``--notify-incomplete-reviews`` (``-nr``): Find and notify students who have not completed required peer reviews for a Canvas assignment
- ``--review-assignment-id`` (``-rai``): Canvas assignment ID for peer review notification
- ``--search-canvas-user`` (``-cu``): Search for a user in Canvas by name or email
- ``--sync-canvas`` (``-sc``): Sync Canvas course members to local database
- ``--unenroll-canvas``: Unenroll Canvas students by domain/email/select/missing-id
  Successful unenroll removes matching students from the local database (by Email or Canvas ID).

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

Configuration
-------------

- ``--backup-config``: Back up config.json to a timestamped file (optional: backup dir)
- ``--clear-config`` (``-ccfg``): Delete stored config.json from the default location
- ``--clear-credentials`` (``-ccred``): Delete stored credentials.json and token.pickle from the default location
- ``--config`` (``-cfg``, ``-c``): Load config from JSON file and save to default location
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
- ``--list-cli-aliases``: List auto-generated short aliases for long-only CLI flags
- ``--log-backups``: Number of rotated log files to keep
- ``--log-dir``: Directory for log files (default: config folder)
- ``--log-level``: Logging level (default: INFO)
- ``--log-max-bytes``: Max size in bytes for rotating logs
- ``--refine``: Refine generated messages/announcements with AI
- ``--verbose`` (``-v``): Enable verbose output
- ``--version``: Show package name and version and exit.

Google Classroom
----------------

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
- ``--list-google-students``: List Google Classroom students for a course (uses ``--google-course-id`` or prompts)
- ``--sync-google-classroom`` (``-sgc``): Sync students in the local database with active students from Google Classroom course
- ``--unenroll-google-classroom`` (``-ugc``): Unenroll Google Classroom students by email domain
  Successful unenroll removes matching students from the local database (by Email or Google_ID).

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
- ``--import-mini-projects``: Import mini-project data from Google Sheets (lecturer topics + student registrations)
- ``--mini-project-lecturer-sheet``: Google Sheet URL for lecturer mini-project topics (optional; default from config)
- ``--mini-project-registration-sheet``: Google Sheet URL for student mini-project registrations (optional; default from config)
- ``--all-details`` (``-A``): Show details of all students
- ``--backup-db``: Back up students.db to a timestamped file (optional: backup dir)
- ``--db`` (``-db``, ``-D``): Database file name (default: students.db, saved in script folder)
- ``--db-backup-keep``: Number of database backups to retain (default from config)
- ``--details`` (``-d``): Show details of a student by name, student id, or email
- ``--dry-run-rows``: Number of preview rows to show with --dry-run + --add-file (default: 5)
- ``--duplicate-name-field``: Field to detect duplicates (name/google/canvas or custom field name)
- ``--duplicate-name-format``: Output format for duplicate-name report (default: txt)
- ``--duplicate-name-output``: Output path for duplicate-name report (optional; extension inferred from format if missing)
- ``--export-all-details`` (``-E``): Export all student details to TXT file
- ``--export-anonymized`` (``-ean``): Export anonymized roster to CSV (optional: output path)
- ``--export-emails`` (``-e``): Export all student emails to TXT file (avoids duplicates)
- ``--export-emails-and-names`` (``-en``): Export all student emails and names to TXT file (default: emails_and_names.txt)
- ``--export-excel`` (``-x``): Export student list to Excel file
- ``--export-final-grade-distribution``: Export final grade distribution to a TXT file. Optionally provide output path (default: ./final_grade_distribution.txt).
- ``--export-grade-diff``: Export grade updates to CSV when updating MAT files (optional: output path)
- ``--export-roster`` (``-ero``): Export classroom roster to CSV file (default: classroom_roster.csv)
- ``--generate-final-evaluations`` (``-gfe``): Generate per-student final evaluation TXT reports (optional: output dir, default: ./final_evaluations).
- ``--list-duplicate-names``: List students who share the same full name
- ``--list-email-domain`` (``-led``): List students whose email matches domain(s) (comma-separated, e.g., gmail.com,outlook.com)
- ``--list-missing-ids``: List students missing Google/Canvas/Student IDs (optional: google,canvas,student,all or comma-separated)
- ``--list-submission-status``: List students by submission status (prefix with ``google:`` or ``canvas:``, optional ``@assignment``). Values: Google Classroom = ``NEW``, ``CREATED``, ``TURNED_IN``, ``RETURNED``, ``RECLAIMED_BY_STUDENT``; Canvas = ``UNSUBMITTED``, ``SUBMITTED``, ``GRADED``, ``PENDING_REVIEW``, ``COMPLETE``.
- ``--load`` (``-l``): Load students from database file
- ``--load-override-grades`` (``-log``): Load override_grades.xlsx and persist overrides to the database (default: override_grades.xlsx).
- ``--missing-ids-format``: Output format for missing-ids report (default: txt)
- ``--missing-ids-output``: Output path for missing-ids report (optional; extension inferred from format if missing)
- ``--modify`` (``-m``): Interactively modify the student database
- ``--restore-db``: Restore students.db from a backup (default: latest)
- ``--save`` (``-s``): Save current students to database file
- ``--search`` (``-S``): Search for students by keyword (name, student id, email, etc.)
- ``--sync-mat-canvas``: Sync CC/GK/CK scores from MAT*.xlsx to Canvas assignments (uses configured assignment IDs)
- ``--sync-mat-types``: Comma-separated list of score types to sync (CC,GK,CK). Default: all available.
- ``--update-mat-excel`` (``-ume``): Update MAT*.xlsx file(s) with grades from database (provide one or more file paths)
- ``--validate-data``: Validate student data and write a report (optional: output path)

