Changelog
=========

Version 0.1.1 (2025-12-27)
--------------------------

- Added multi-file glob support for import/update CLI commands.
- Added student detail sort methods and clearer export/report formatting.
- Normalized Canvas/Google Classroom sync scores to a 10-point scale when possible.
- MAT*.xlsx roster imports now ignore score columns (CC, GK, CK, totals).
- MAT Excel updates can infer missing student IDs from VNU University of Science, Hanoi emails.
- Added resubmission grading workflow with optional keep-old-grade default.
- Canvas sync stores submission comments and rubric evaluations.
- Canvas grade parsing now falls back to Unposted Final Score when Final Score is empty/zero (CC/GK/CK).
- Final evaluation reports omit assignment-group scores when all component scores are 0.
- Added course calendar builder (TXT/Markdown/ICS) with holiday exclusions, unofficial holidays, and make-up week logic.
- Course calendar titles include course code/name and canceled sessions are tagged.
- Added auto-generated short aliases for long-only CLI flags.
- Added ``--list-cli-aliases`` to display auto-generated short aliases.
- Documentation and samples updated for new sync, calendar, and local AI tooling.

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
