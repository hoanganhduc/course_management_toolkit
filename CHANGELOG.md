# Changelog

## [0.3.0] - 2026-08-29

- Added `course-c50-admin`, a human-only Classroom50 operator CLI with
  `assignment-add`, `assignment-remove`, `invite`, and `download` subcommands. Each
  verb refuses in agent mode, requires an interactive terminal unless `--dry-run` is
  given, builds one fixed `gh teacher` command line, and redacts secrets from captured
  output.
- `assignment-add` refuses an existing slug by default because `gh teacher assignment
  add` replaces the entry in place and the pinned CLI cannot set submission mode, so a
  repeat run restores the default every-push mode and discards a tagged-commit setting
  made in the web form. Overriding requires `--allow-overwrite` and an interactive
  confirmation; there is no `--yes`.
- `assignment-remove` refuses an absent slug and states in its confirmation that
  student repositories survive and that re-adding the same slug is not a clean reset.
- `invite` preflights organization targets against `gh teacher member list` and skips
  logins that are already members or hold a pending invitation, since organization
  invitations are not idempotent. Repository targets need no preflight.
- Narrowed the blanket `--by-pattern` download ban to an evidence-based gate:
  `--by-pattern` is now permitted only when the assignment record reports
  empty-repository mode, which has no `result.json` or automatic score to collect, and
  refused for autograded assignments where it would skip both.
- `python -m course_hoanganhduc.c50_agent` now returns a structured refusal for
  `assignment-add`, `assignment-remove`, and `invite` instead of an argparse usage
  error. No agent-safe Classroom50 mutation path was added.
- Documented that `CLASSROOM50_ORG_ALLOWLIST` must be set in the invocation rather than
  `~/.bashrc`, whose non-interactive early return leaves it unset under `env -i`, cron,
  systemd, and CI.

## [0.2.0] - 2026-08-28

- Simplified Google Classroom assignment creation: removed redundant typed
  `AUTH`/`CREATE`/`SHARE` phrases, added one readable `y/N` summary, and added
  `--yes` automation for drafts without Drive-sharing effects.
- Added a narrowly allowlisted agent smoke-test path for one minimal Classroom draft,
  with an account/course/client/token-bound approval envelope, atomic existing-token
  enforcement, serialized all-state duplicate detection, zero-retry creation, and
  strict read-back. The normal agent entrypoint and general mutation API still refuse
  creation.
- Hardened Classroom mutation transport so connection retries, redirects, and
  response-triggered OAuth replays cannot repeat a create request. Agent-safe live
  calls use narrow response fields and require an explicit active course state.
- Fixed isolated Google Classroom coursework authorization when Google reports the
  additional `openid`/`email` identity scopes alongside `userinfo.email`, while
  continuing to reject unrelated or missing grants.
- Added `course-gclass-admin complete-loopback` for hidden-input remote OAuth
  callbacks and suppressed dependency logging that could expose callback codes or
  bearer tokens during authorization.
- Added Classroom50 (foundation50) integration: wrap `gh teacher` for preflight, list classrooms/roster/assignments, roster sync into local DB, C50 CSV export, and human-only submission download (`--download-classroom50`).
- Added agent-safe entrypoints: `python -m course_hoanganhduc.c50_agent`, `canvas_agent`, `gclass_agent`, `db_agent` (force agent mode; refuse destructive LMS/DB ops; org/course allowlists fail closed).
- Shared helpers in `course_agent_common`; GitHub numeric id kept separate from Username in student identity maps (`data.py`).
- CLI flags under the Classroom50 group (see `docs/cli_reference.rst`).
- Documented agent entrypoints and Classroom50 workflows in README and `docs/usage.rst`.

## [0.1.4] - 2026-01-25
- Enhanced `--list-submission-status` to show attachment details (file count, size, type, upload time).
- Updated Canvas and Google Classroom sync to fetch detailed attachment info (supports Drive files, Links, Forms, YouTube videos).
- Updated `--export-all-details` to include full attachment breakdown.
- Refined default sort order for roster exports: Section > First Name > Last Name > Student ID.
- Implemented Vietnamese-specific collation for correct alphabetical sorting (e.g., handles 'Â', 'Đ' correctly).

## [0.1.3] - 2026-01-25
- Improved `export_companies_to_vcf` with `raw_data` fallback logic to extract contact person and phone numbers.
- VCF exports (students and companies) now skip entries with no phone number information.
- Enhanced company VCF export filename detection to use `companies_contacts.vcf` by default.
- Improved company data import column mapping (better Vietnamese keyword support and email priority).
- Fixed `UnboundLocalError` in `course --import-internships` by adding missing local import.
- Fixed `UnpicklingError` in `course --export-vcf` (and other commands) when loading SQLite databases with `companies` table.
- Added company field mappings for VCF export and student loading.
- VCF export now strictly uses `UNIVERSITY_NAME` from config (defaults to empty if not set) instead of hardcoded default.

## [0.1.2] - 2026-01-24
- Added `--import-registrations` command to support student internship registration data (skills, wishlist, notes).
- Added `INTERNSHIP_REGISTRATION_SHEET_URL` configuration key.
- Added company contact management with `--import-companies` and `--export-companies`, stored in `companies.db`.
- Improved export formatting for student details (JSON unwrapping for progress reports, translated labels, hidden empty fields).
- Added support for English headers in internship data and registration imports.
- Updated documentation and sample data to English.

## [0.1.1] - 2025-12-28
- Added multi-file glob support for import/update CLI commands.
- Added student detail sort methods and clearer export/report formatting.
- Normalized Canvas/Google Classroom sync scores to a 10-point scale when possible.
- MAT*.xlsx roster imports ignore score columns (CC, GK, CK, totals).
- MAT Excel updates can infer missing student IDs from VNU University of Science, Hanoi emails.
- Added resubmission grading workflow with optional keep-old-grade default.
- Canvas sync stores submission comments and rubric evaluations.
- Canvas grade parsing falls back to Unposted Final Score when Final Score is empty/zero (CC/GK/CK).
- Final evaluation reports omit assignment-group scores when all component scores are 0.
- Final evaluation weights are configurable (`WEIGHT_CC`, `WEIGHT_GK`, `WEIGHT_CK`) and the TXT output includes the formula.
- Added course calendar builder (TXT/Markdown/ICS) with holiday exclusions, unofficial holidays, and make-up week logic.
- Course calendar titles include course code/name and canceled sessions are tagged.
- Added Canvas calendar import from iCal (.ics) files with dry-run and duplicate skipping.
- Improved Canvas announcement flow with short input, AI refinement, and confirmation before posting.
- Added auto-generated short aliases for long-only CLI flags and `--list-cli-aliases`.
- Documentation and samples updated for new sync, calendar, and local AI tooling.

## [0.1.0] - 2025-12-27
- Packaged the original script as a Python CLI with install helpers and standardized flags.
- Added course-scoped config/credential storage with cached course codes.
- Rebuilt the no-args menu with sections, arrow-key navigation, and numeric quick-jump.
- Added `--clear-config` and `--clear-credentials` helpers plus Windows/Linux compatibility fixes.
- Canvas/Google Classroom sync now resolves duplicates; Canvas grade sync stores final scores only.
- Added override grades (flexible headers/aliases) with clearer reporting in MAT exports.
- Added AI model testing/listing, rate-limit fallback, and model details in reports.
- Expanded submission quality checks with configurable thresholds and richer diagnostics.
- Added backup/restore commands, dry-run mode, validation reports, grade audit history, and grade diff exports.
- Added import previews, anonymized exports, and per-run summaries in `run_report.txt`.
- Added weekly automation workflow generation and non-interactive Canvas checks/reminders.
- Added local weekly automation with archived reports and flagged-submission evidence.
- Weekly workflow now clones the toolkit, archives reports with DB backups, and tags evidence by assignment.
- Weekly automation can auto-detect closed assignments not yet in weekly reports.
- Added local LLM support (Ollama-compatible) for AI refinement and message generation, with CLI overrides.
- Added local model detection via `--detect-local-ai` (Ollama or llama.cpp) and optional `.gguf` scan.
- Updated docs and samples; added GPL-3.0-only license.
