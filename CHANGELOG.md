# Changelog

## [0.4.0] - 2026-09-06

- Added `course --invite-google-classroom` (`-igc`), the first write path from the local
  database to Google Classroom. Candidates are filtered with `--gc-invite-email`,
  `--gc-invite-domain`, `--gc-invite-section`, and `--gc-invite-class`, selected from a
  numbered list unless `--gc-invite-all` is given, and invited as `--gc-invite-role`
  `STUDENT` or `TEACHER`. Menu entry 67 runs the same flow.
- The invite is idempotent in three layers: the course roster and teacher list are read
  first, pending invitations are matched against the locally stored `Google_ID` before any
  `userProfiles.get` fan-out (capped by `profile_lookup_limit`), and a `409 ALREADY_EXISTS`
  from `invitations().create` is recorded as `skipped_pending` rather than a failure. A
  second run therefore issues no `create` calls.
- A `400 FAILED_PRECONDITION` from `create` is recorded as `skipped_precondition`, not as a
  failure. The Classroom reference documents two causes -- the invited user already holds
  this role or a greater one, or the account is disabled -- and the first is an enrolment
  the bulk read missed, which `409` does not cover. The classification reads `error.status`
  from the response body rather than matching the message text, and a `400` without that
  canonical name is still a failure.
- A `403` stops the loop instead of repeating per student: the reference documents it as
  "the requesting user is not permitted to create invitations for this course", so it is
  course-wide and continuing would produce one identical failure and burn one quota unit per
  remaining candidate. Those are reported as `skipped_aborted`. A `404` names the course as
  well as the user, since the reference defines it as "the course or the user does not exist".
- Successful invitations record `Google_ID`, `Google_Invitation_ID`, and
  `Google_Classroom_Invited_At` in the local database. `Google_ID` is filled only when empty
  and never overwritten, because a Google account identifier does not change and a stored one
  is what makes the second run cost no profile lookups. The other two are rewritten on every
  successful invitation, since they describe the most recent invitation rather than the
  student.
- `invitations().create` retries `429`/`500`/`503` with bounded backoff. This differs from
  the coursework module's `num_retries=0` on purpose: an invitation is deduplicated by its
  own 409, so a repeat is safe, whereas a coursework create is not.
- Added `course-c50-admin roster-import --org ORG --classroom SHORT (--db PATH | --csv PATH)`,
  which writes the Classroom50 roster and, as `gh teacher roster import` does once the commit
  lands, invites every listed student who is not already an organization member. Both effects
  are named in the confirmation. `roster import` is an upsert, so rows absent from the CSV are
  left untouched and rows that exist only on the remote roster are reported under `remoteOnly`
  rather than removed.
- `roster-import` reads the remote roster first, but a matching roster is not on its own proof
  that the work is done: `gh teacher roster import` commits the CSV before it invites, so a run
  that failed in between -- the exit-3 case whose documented recovery is to re-run -- leaves the
  roster matching and nobody invited. When nothing would change, the organization membership is
  therefore read as well, and the command short-circuits with `roster already matches the local
  database and every listed student is already in the organization; nothing to import` only when
  everyone is in; otherwise the missing logins are reported under `notYetMembers` and the import
  runs, which is what sends the invitations. This extra read happens only on the otherwise
  no-op path. A membership read that fails, typically a token without the `admin:org` scope,
  keeps exit 0 rather than turning a pure read into outcome-unknown, sets
  `membershipNotVerified`, and points at the new `--force`, which imports a matching roster
  without the check. Students with no GitHub username are reported under `skippedNoUsername`
  and do not change the exit code, because skipping them is a deliberate non-write the operator
  saw before confirming.
- The CSV generated from `--db` is written into a `tempfile.mkdtemp` directory that is removed
  even when `gh` fails, unless `--keep-csv` is given. `--keep-csv` with `--csv` is refused as an
  invalid flag combination rather than ignored, since that path creates no temporary file and a
  silently inert flag hides a mistaken invocation. `gh` stderr, which carries a
  version-specific partial-onboarding warning, is redacted and passed through verbatim as
  `ghStderr` rather than parsed.
- Reading the database for `--db` prints progress on stdout, so that output is redirected onto
  the error stream: stdout stays parseable JSON on every `course-c50-admin` verb.
- `roster-import --dry-run` makes no `gh` call, matching the other verbs and keeping `--dry-run`
  the one mode `_require_human` allows without a terminal. It therefore reports a `plan`
  (`csvSource`, `rowCount`, `firstUsernames`, `skippedNoUsername`) instead of a roster diff, and
  the `--db` form prints no `argv`, since the CSV path does not exist yet.
- Added `course-c50-admin roster-add` for a single late student, with the same invitation side
  effect, so one arrival does not require a full CSV round trip.
- `python -m course_hoanganhduc.c50_agent` now refuses `roster-import` and `roster-add`, and
  `python -m course_hoanganhduc.gclass_agent` refuses `invite`. No agent-safe write path to
  either platform was added. `gclass_agent` kept its refusal list in two places that could
  drift apart; it is now one `REFUSED_VERBS` constant.
- Added `c50_roster.partition_roster_candidates`, which reports the students
  `export_roster_csv` drops for having no GitHub username. `export_roster_csv` itself is
  unchanged, byte for byte, and pinned by a test.
- Added `course --add-csv FILE`, a CSV source for the local database that needs no Google
  account. Unlike `--add-file` the columns are a fixed contract rather than auto-detected:
  `Student ID` and `Name` are required, as is at least one of `Email` and `GitHub Username`,
  since without one of those the roster feeds neither platform. `GitHub ID`, `Section` and
  `Class` are read when present. Headers are matched case- and accent-insensitively, so the
  Vietnamese spellings work, and a header that matches nothing is refused with the columns that
  were actually read. Rows missing a student id or a name are skipped and counted rather than
  failing the run. The file is decoded as UTF-8 (BOM or not), then Windows-1252, then Latin-1,
  naming the encoding whenever it falls back -- deliberately the opposite of
  `roster-import --csv`, which refuses anything but UTF-8 because a mis-decoded name would be
  pushed to GitHub. Merging goes through the existing `save_database`, whose deduplication
  already collapses repeats. `--add-file` is untouched and still handles Excel, PDF, and the
  `MAT*.xlsx` quirks.
- Added `course --onboard`, one interactive pass from a registration CSV to both platforms:
  pick a Google Classroom course, pick a Classroom50 classroom, pick the CSV, then add whoever
  is missing. Either platform can be skipped, and skipping both is a valid run that only loads
  the CSV into the database and reports what the form contains. `--onboard-csv` skips the file
  prompt and `--onboard-skip-github-check` validates usernames by syntax alone when offline.
  No new selector flags were added: `--google-course-id`, `--classroom50-org`,
  `--classroom50-classroom`, and `--classroom50-report` are the existing ones, so an org
  configured as `CLASSROOM50_ORG` resolves the same way it does everywhere else.
- `--onboard` now takes a configured course as a default rather than asking for it again.
  `CLASSROOM50_ORG` and `CLASSROOM50_CLASSROOM` were absent from the `load_config` whitelist, so
  neither could be recorded in `config.json` at all, and the resolver returned on the first
  config file it found, which meant a file holding any other key silenced the environment
  variable as well. All three selectors now fall through flag, then config file, then
  environment, tier by tier, and carry which tier answered: a flag is used as given, while a
  config or environment value is shown and confirmed first, and declining falls through to the
  same picker as supplying nothing. A replaced org sends the configured classroom to the picker
  too, since a classroom configured for one org is no default for another. Off a terminal, and
  under `--dry-run`, the value is taken without a prompt and the run says which value it took.
- Google authorization now works on a machine with no browser, and finishes in the one terminal
  it started in. `run_local_server` opens the browser before it prints the authorization URL and
  does not guard that call, so on a headless server `webbrowser.get` raised and the URL was never
  shown: a first authorization could not be completed at all. It also blocks in `handle_request()`
  until a browser reaches its loopback port, so the process that prints the URL can never also ask
  for the answer. Where a browser is registered the loopback server still runs. Where none is, no
  listener is started at all: the redirect URI is set to `http://127.0.0.1:8910/`, the URL is
  printed to be opened elsewhere, and the redirect that browser fails to load is pasted back at a
  hidden prompt, which exchanges it for the token. A paste is checked for the loopback boundary and
  for exactly one code and state before it is used, a paste from a different run is refused by the
  state it carries, and a malformed one is asked for again rather than ending the run.
  `--no-open-browser` forces that path on a machine that does have a browser.
- `course-gclass-admin authorize` finishes the same way, so both authorization lanes behave
  alike. It assumed a browser rather than looking for one, and its headless procedure needed
  `complete-loopback` in a second terminal. It now looks first, keeps the loopback server where
  a browser exists, and otherwise prints the URL and takes the failed
  `http://127.0.0.1:8911/` redirect back at its own hidden prompt. The redirect port is its
  own, so a redirect from the other lane is refused rather than exchanged. The paste is held to
  the boundary `complete-loopback` enforces, the granted scopes are still verified before a
  token is written, and `complete-loopback` itself is unchanged for an authorization that is
  already listening.
- The header reader now matches Google Form columns. A form header carries a bilingual
  parenthesis and the asterisk the form adds to a required question, which the alias table did
  not model: of the nine columns a real registration form exports, two matched. Aliases are now
  compared against the header with the parenthesised part dropped, with the asterisk dropped,
  and against the parenthesised part alone, so `Họ và Tên (Full Name)` matches through either
  language. A header carrying both `vnu` and `hus` is read as the school address, which also
  covers the missing letter in the form's own `Emai VNU-HUS`.
- Columns are no longer claimed first-come. Each attribute takes the best-ranked column that
  claims it, so a form carrying a school address and a personal one lands the school address in
  `Email` instead of losing it to `Email Address`. `Personal Email`, `Date of Birth`, and
  `Timestamp` are read while resolving rows and never written to the database, so `--add-csv`
  writes exactly the attributes it wrote before, and a CSV whose only address column is
  `Email Address` still fills `Email`.
- Addresses are validated and chosen per student: the school column wins, an empty or malformed
  one falls back to the personal column, and the report says which column each student came
  from. GitHub usernames are validated against GitHub's own account rule -- not the classroom
  slug rule, which is a different grammar -- and then confirmed with `gh api /users/<name>`,
  which yields the numeric id that `gh teacher roster import` re-resolves anyway. An API that
  cannot be reached marks the name *unverified* rather than invalid, because a flaky network
  must not reject a class; those students are named in the report and their `github_id` is
  written empty, which import treats as "look it up", where a stale one fails the line.
- A bad row is skipped with its own reason and the run continues. A row survives when either
  lane can still use it: with no usable address it still reaches the Classroom50 roster, which
  accepts an empty `email` per row, and the report names those students because Google
  Classroom invites by address and will never see them. Repeat submissions collapse on
  `Student ID`, keeping the last row, since Google Forms appends rather than replaces. Two
  different students claiming one GitHub username drop *both*, named, because `roster import`
  upserts by username and would otherwise overwrite one with the other in silence.
- Everything reads before anything writes. The Classroom50 roster is pulled into the database
  first, so a student already on the roster under a different identifier is not added twice.
  Then `gh teacher roster sync` runs in its reporting mode, which issues no write request:
  exit 2 means changes are pending, not that the run failed, and it stops the onboarding and
  asks the operator to run `--write` first. A roster row still awaiting its email link carries
  no username, is invisible to the importer's own diff, and would otherwise be imported as a
  second row for the same person.
- The comparison table is printed before the single confirmation, listing every student against
  both platforms, rejected rows included with the reason. The Google column comes from the three
  reads the invite path already does; for the shortlist it would invite, `invitations.list` is
  asked once per candidate with `userId` set to the address, so a pending invitation beyond
  `profile_lookup_limit` is not previewed as "will invite" and then silently skipped at write
  time. The Classroom50 column asks the roster and the organization membership separately, for
  the same reason `roster-import` does: the roster commit precedes the invitations.
- Writes then run in order: the database, then the Google Classroom invitations, then
  `roster-import` called unchanged, keeping its own confirmation gate. A confirmed onboarding
  therefore answers twice. `--dry-run` makes no `gh` call, no Google call, and no database
  write, so the comparison table appears only in a real run.
- A Google Classroom course list that cannot be read -- no credentials on the machine, an
  expired token, an API that is down -- is now the same as choosing to skip it, with the reason
  printed. It used to raise, which took the Classroom50 half of the run down with it.
- Pending invitations are still matched by `Google_ID` after the `profile_lookup_limit` cap is
  reached. The reader used to stop at the cap, which also dropped the remaining `userId` values
  from the comparison set even though matching against them costs nothing; those students were
  re-invited and caught only by the `409` backstop.
- A database that cannot be read while recording sent invitations now says so instead of
  failing silently. The invitations really went out, so this stays non-fatal, but the next run
  pays for the profile lookups again and the operator should know why.
- `gclass_invite` labelled a course teacher as `already enrolled` and a student matched only by
  `Google_ID` as a teacher, because the student and teacher sets were merged before the check.
  The counts and the skip decision were always right; only the reported reason was wrong.
- The Classroom50 adapter is coupled to the installed `gh teacher` version, which is now
  recorded where that matters. The extension was upgraded from v1.25.1 to v1.45.0 (commit
  `a5a51293b65d`) and re-pinned, for two upstream fixes that bear on `roster-import`:
  non-ASCII names are no longer passed through as invalid UTF-8, and a student accepting an
  invitation mid-import no longer corrupts the roster. No toolkit code changed for the
  upgrade. `_SLUG_RE` stays at 38 characters although v1.45.0 accepts 99, because the same
  release added a per-repo name budget the regex does not model. README gained a "Version
  coupling with `gh teacher`" table, and `c50_auth.gh_teacher_available` now states that it
  probes for presence only and never reads the version.

## [0.3.1] - 2026-08-30

- `course-c50-admin assignment-add` now covers autograded weekly registration:
  `--tests`, `--feedback-pr` / `--no-feedback-pr`, `--pass-threshold`, `--allowed-files`
  (repeatable, order preserved), and `--student-permission`. Without them the adapter
  could express only the empty-repository final project, so weekly registration had no
  guarded path.
- `--tests` is checked before the call: the path must be a readable JSON file holding a
  bare array of test specs. The pinned CLI's `-` stdin form is not offered because no
  operand may begin with `-`.
- `--empty-repo` is now refused together with any of `--template`, `--tests`,
  `--feedback-pr`, `--allowed-files`, and `--pass-threshold`, matching the pinned CLI,
  where the empty-repository setting is immutable after creation.
- Fixed `docs/cli_reference.rst`, which listed `--tests`, `--feedback-pr`,
  `--allowed-files`, `--pass-threshold`, `--student-permission`, and `--locked` before
  any of them were implemented. `--locked` is not a flag of `gh teacher assignment add`
  at all; `--autograder` and `--runtime` stay with the reviewed raw command.

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
