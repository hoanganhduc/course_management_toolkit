Changelog
=========

Version 0.1.1 (2025-12-27)
--------------------------

- Added multi-file glob support for import/update CLI commands.
- Added student detail sort methods for exports and reports.
- Normalized Canvas/Google Classroom sync scores to a 10-point scale when possible.
- MAT*.xlsx roster imports now ignore score columns (CC, GK, CK, totals).
- Improved student detail formatting and Classroom total score output.
- Improved MAT Excel updates by inferring missing student IDs from email when needed (VNU University of Science, Hanoi emails only).
- Added resubmission grading workflow with optional keep-old-grade default.
- Resubmission grading now lists only assignments needing regrade and excludes Roll Call Attendance.
- Canvas sync now stores submission comments and rubric evaluations per assignment in the database.
- Removed Canvas gradebook download (Canvas API not available).
- Added auto-generated short aliases for long-only CLI flags.
- Added ``--list-cli-aliases`` to display auto-generated short aliases.
- Updated documentation to cover new sync and import behavior.

Version 0.1.0 (2025-12-27)
--------------------------

- Packaged the original script as a Python CLI with install helpers and standardized flags.
- Added course-scoped config/credential storage with cached course codes.
- Rebuilt the no-args menu with sections, arrow-key navigation, and numeric quick-jump.
- Added ``--clear-config`` and ``--clear-credentials`` helpers plus Windows/Linux compatibility fixes.
- Canvas/Google Classroom sync now resolves duplicates; Canvas grade sync stores final scores only.
- Added override grades (flexible headers/aliases) with clearer reporting in MAT exports.
- Added AI model testing/listing, rate-limit fallback, and model details in reports.
- Expanded submission quality checks with configurable thresholds and richer diagnostics.
- Added backup/restore commands, dry-run mode, validation reports, grade audit history, and grade diff exports.
- Added import previews, anonymized exports, and per-run summaries in ``run_report.txt``.
- Added weekly automation workflow generation and non-interactive Canvas checks/reminders.
- Added local weekly automation with archived reports and flagged-submission evidence.
- Weekly workflow now clones the toolkit, archives reports with DB backups, and tags evidence by assignment.
- Weekly automation can auto-detect closed assignments not yet in weekly reports.
- Added local LLM support (Ollama-compatible) for AI refinement and message generation, with CLI overrides.
- Added local model detection via ``--detect-local-ai`` (Ollama or llama.cpp) and optional ``.gguf`` scan.
- Updated docs and samples; added GPL-3.0-only license.
