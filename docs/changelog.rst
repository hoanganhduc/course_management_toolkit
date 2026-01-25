Changelog
=========

Version 0.1.4 (2026-01-25)
--------------------------

- Enhanced ``list --submission-status`` to show attachment details (file count, size, type, upload time).
- Updated Canvas and Google Classroom sync to fetch detailed attachment info (supports Drive files, Links, Forms, YouTube videos).
- Updated ``--export-all-details`` to include full attachment breakdown.

Version 0.1.3 (2026-01-25)
--------------------------

- Improved ``export_companies_to_vcf`` with ``raw_data`` fallback logic to extract contact person and phone numbers.
- VCF exports (students and companies) now skip entries with no phone number information.
- Enhanced company VCF export filename detection to use ``companies_contacts.vcf`` by default.
- Improved company data import column mapping (better Vietnamese keyword support and email priority).
- Fixed ``UnboundLocalError`` in ``course --import-internships`` by adding missing local import.
- Fixed ``UnpicklingError`` in ``course --export-vcf`` (and other commands) when loading SQLite databases with ``companies`` table.
- Added company field mappings for VCF export and student loading.
- VCF export now strictly uses ``UNIVERSITY_NAME`` from config (defaults to empty if not set) instead of hardcoded default.

Version 0.1.2 (2026-01-24)
--------------------------

- Added ``--import-registrations`` command to support student internship registration data (skills, wishlist, notes).
- Added ``INTERNSHIP_REGISTRATION_SHEET_URL`` configuration key.
- Added company contact management with ``--import-companies`` and ``--export-companies``, stored in ``companies.db``.
- Improved export formatting for student details (JSON unwrapping for progress reports, translated labels, hidden empty fields).
- Added support for English headers in internship data and registration imports.
- Updated documentation and sample data to English.

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
- Final evaluation weights are configurable (``WEIGHT_CC``, ``WEIGHT_GK``, ``WEIGHT_CK``) and the TXT output includes the formula.
- Added course calendar builder (TXT/Markdown/ICS) with holiday exclusions, unofficial holidays, and make-up week logic.
- Course calendar titles include course code/name and canceled sessions are tagged.
- Added Canvas calendar import from iCal (.ics) files with dry-run and duplicate skipping.
- Improved Canvas announcement flow with short input, AI refinement, and confirmation before posting.
- Added auto-generated short aliases for long-only CLI flags.
- Added ``--list-cli-aliases`` to display auto-generated short aliases.
- Added duplicate-name reporting (Name/Google Classroom/Canvas display names) with TXT/CSV/JSON exports.
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
