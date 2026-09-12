Changelog
=========

Version 0.4.0 (2026-09-06)
--------------------------

- Added ``course --invite-google-classroom`` (``-igc``), the first write path from the
  local database to Google Classroom, with ``--gc-invite-email``/``-domain``/``-section``/
  ``-class`` filters, a numbered selection unless ``--gc-invite-all`` is given, and
  ``--gc-invite-role`` ``STUDENT`` or ``TEACHER``.
- The invite is idempotent in three layers: the roster and teacher list are read first,
  pending invitations are matched against the stored ``Google_ID`` before any
  ``userProfiles.get`` fan-out, and a ``409 ALREADY_EXISTS`` is recorded as
  ``skipped_pending``. A ``400 FAILED_PRECONDITION`` is a skip, a ``403`` is course-wide
  and stops the loop, and ``429``/``500``/``503`` are retried with bounded backoff.
- Successful invitations record ``Google_ID`` (filled only when empty), plus
  ``Google_Invitation_ID`` and ``Google_Classroom_Invited_At``, which describe the latest
  invitation and are rewritten each time.
- Added ``course --add-csv FILE``, a fixed-contract CSV source for the database that needs
  no Google account: ``Student ID`` and ``Name`` are required, as is at least one of
  ``Email`` and ``GitHub Username``; headers are matched case- and accent-insensitively;
  and the file is decoded as UTF-8, then Windows-1252, then Latin-1, naming the fallback.
- Added ``course --onboard``, one interactive pass from a registration CSV to both
  platforms: pick a Google Classroom course, pick a Classroom50 classroom, pick the CSV,
  add whoever is missing. Either platform can be skipped. ``--onboard-csv`` skips the file
  prompt, ``--onboard-skip-github-check`` validates usernames by syntax alone, and the
  course, org, classroom, and report selectors are the existing flags.
- ``--onboard`` now takes a configured course, org, and classroom as defaults. The two
  Classroom50 keys were missing from the ``load_config`` whitelist and the resolver
  returned on the first config file it found, silencing the environment tier; all three
  selectors now fall through flag, then config, then environment, and a value that was not
  typed for this run is shown and confirmed before use. Declining falls through to the
  picker; off a terminal and under ``--dry-run`` the default is taken and named.
- Google authorization now works without a browser, and finishes in one terminal.
  ``run_local_server`` opens the browser before it prints the authorization URL and does not
  guard that call, so on a headless server the URL was never shown; it also blocks until a
  browser reaches its loopback port, so the process printing the URL can never also read the
  answer. A machine with a browser still uses the loopback server. A machine without one
  starts no listener: it sets the redirect URI to ``http://127.0.0.1:8910/``, prints the URL
  to open elsewhere, and exchanges the failed redirect pasted back at a hidden prompt. The
  paste is validated, and its state must match the run that printed the URL.
  ``--no-open-browser`` forces that path on a machine that does have a browser.
- ``course-gclass-admin authorize`` finishes the same way, on its own
  ``http://127.0.0.1:8911/`` redirect, and now looks for a browser instead of assuming
  one. The paste is held to the boundary ``complete-loopback`` enforces and granted
  scopes are still verified before a token is written. ``complete-loopback`` is
  unchanged and still serves an authorization that is already listening.
- The header reader now matches Google Form columns, whose bilingual parentheses and
  required-question asterisk previously matched two of nine. Each attribute also takes the
  best-ranked column claiming it, so a form carrying a school address and a personal one
  lands the school address in ``Email`` and the personal one in ``Personal Email``. Every
  column read is written: the mapped ones under their attribute name, the rest under their
  own header, so a question added to the form later needs no alias to be kept and a later
  question about a form column is answered from the database rather than from the sheet.
  The date of birth is written as ``Dob``, the name ``read_students_from_excel_csv`` already
  uses, so both import lanes leave one spelling behind.
- A student who fills the form again to correct a mistake now reaches the database. The
  merge that runs on every save filled only fields that were empty, so a correction to an
  answer already stored was dropped without a word. Between two records that both carry a
  readable ``Timestamp``, the later one now wins and every replaced value is printed. A
  ``Timestamp`` missing on either side, one that does not parse, and two equal ones all keep
  the older fill-what-is-empty rule, so the Google Classroom and Classroom50 records, which
  carry no ``Timestamp``, can neither overwrite nor be overwritten, re-importing the same
  sheet changes nothing, and the outcome no longer depends on the order rows arrive in.
  ``Google_*`` and ``Additional *`` fields are never overwritten. A newer submission that
  changes ``GitHub Username`` clears any stored ``GitHub ID``, because ``roster import``
  re-resolves the id from the username and fails the line when the two disagree. Two
  conflicting sections are still kept side by side as ``Additional Section`` when no
  ``Timestamp`` tells them apart; between two submissions the later section is a correction
  and replaces the earlier one.
- That merge now actually runs on the Google Sheet lane. ``read_students_from_excel_csv``
  and ``read_students_from_pdf`` each carried a private copy of the deduplication loop and
  collapsed every duplicate before ``save_database`` was reached, so the rule above saw no
  pair to decide and a corrected submission was dropped however often the import ran. Both
  copies are gone and both lanes hand their rows to ``save_database``, as ``--add-csv``
  already did. The PDF copy also tested two names it never defined, which would have raised
  ``NameError`` on the branch that compares two records sharing no student number.
- Two records are one student when they share a ``Google_ID``, or a ``Student ID``, or an
  email address while at most one of them carries a student number, or a name while neither
  does. ``Google_ID`` is one Google account and so one person whatever the two rows call
  them; the email key preserves what the deleted sheet-lane copy did, and is refused when
  both sides carry a student number, because a shared address is a typo or a family and
  merging two students loses a whole registration.
- ``Timestamp`` and ``Dấu thời gian`` are recognised column headers. The submission time
  previously survived only through the keep-unknown-columns branch and only for a form
  written in English; a Vietnamese form lost it, and with it every "the later one wins"
  decision.
- No import drops a row in silence any more. A student number that is not 8 digits keeps
  the registration: the value moves to ``Invalid Student ID``, ``Student ID`` is left empty
  so the roster export and the grade join stay clean, and the student is named on stdout.
  A row that names nobody is still dropped, but by name and address. ``refine_database``,
  which runs on every save, likewise prints what it removed instead of only how many. A
  bare ``continue`` here is what made a student who had filled the registration form
  indistinguishable from one who never did.
- ``--add-google-sheet`` accepts ``--dry-run``: it reads and merges exactly as a real
  import would, prints each field a newer submission would replace, and writes nothing.
- ``--sync-google-classroom`` matches on ``Google_ID`` before any string. Matching on
  display name, name and address alone created a second row for every student who joined
  Classroom with a personal address and filled the form with the school one -- the source
  of MAT3508's 70 database rows against 67 Classroom accounts. The display name is now read
  under both spellings it has been stored as and written under one, so a name the sync had
  just stored is found on the next run instead of being rewritten every time.
- The sync no longer guesses when nobody is there to ask. An ambiguous record with no
  terminal is skipped and reported rather than resolved, because creating a row is the one
  outcome that re-running cannot undo; the course picker says which flag to pass instead of
  raising ``EOFError``. After each sync, rows carrying a ``Google_ID`` the course no longer
  returns, rows carrying none at all, and records skipped as ambiguous are each listed with
  both counts. Nothing is deleted or marked: only the teacher can tell a student who left
  from one who filled the form before being enrolled.
- ``--list-duplicate-accounts`` compares people rather than records. It was handed the raw
  roster, so two rows of one student were compared with each other and the toolkit reported
  its own duplicate as a student holding two Google accounts -- two of MAT3508's three
  findings. Records are grouped into people first, each person is named by the row they
  filled in themselves, and a pair found under both the name and the address signal is
  reported once.
- A student number rejected at import is a finding of its own. Its owner is no longer
  listed as never having answered the form, the new ``student_id_invalid`` check says what
  is wrong with the number rather than only that it is -- how many digits it has and how
  many are missing or spare -- and the empty ``Student ID`` left behind is not also
  reported as a formatting fault.
- The audit asks GitHub whether each registered username exists by default, whenever ``gh``
  is installed; ``--no-audit-check-github`` turns it off and an absent ``gh`` skips it. The
  check existed but was off unless requested, so a username that is well-formed and belongs
  to nobody passed every report, and nothing else in the audit can see it.
- ``--list-classroom50-membership`` counts from the live roster instead of the local
  database, and separates four groups: in the organization, invitation pending, on the
  roster and not yet invited, and in the database but never imported to the roster. The
  last two were one pile, and they need different commands. Usernames are compared
  case-insensitively, since the database export lower-cases them and the roster keeps
  GitHub's own capitalisation. When the roster cannot be read, the header says the count
  came from the database.
- Addresses and GitHub usernames are validated per student, the username against GitHub's
  own account rule and then against ``gh api /users/<name>``; an unreachable API marks the
  name unverified rather than invalid. Bad rows are skipped with a reason and the run
  continues; repeat submissions collapse on ``Student ID``; two students claiming one
  username drop both.
- Everything reads before anything writes: the roster is pulled into the database, then
  ``gh teacher roster sync`` reports, whose exit 2 means changes are pending and stops the
  run, then both platforms are read into one comparison table shown before the single
  confirmation. ``--dry-run`` calls neither platform.
- Added ``course-c50-admin roster-import`` and ``roster-add``, which write the Classroom50
  roster and, as ``gh teacher roster import`` does once the commit lands, invite every
  listed non-member. Because the commit precedes the invitations, a matching roster alone
  does not end the run: the organization membership is read too, and ``--force`` skips that
  check.
- ``python -m course_hoanganhduc.c50_agent`` refuses ``roster-import`` and ``roster-add``,
  and ``gclass_agent`` refuses ``invite``; no agent-safe write path to either platform was
  added.
- The ``gh teacher`` extension was upgraded from v1.25.1 to v1.45.0 (commit
  ``a5a51293b65d``) and re-pinned, for the invalid-UTF-8 and concurrent-acceptance fixes
  that bear on ``roster-import``. No toolkit code changed for the upgrade, and the
  version-coupled constants are now listed in the README.
- Added ``course --import-classroom50-scores``, which reads the classroom's collected marks
  from ``<classroom>/scores.json`` in the configuration repository rather than from a
  downloaded submission tree. That is one authenticated read, the same class of operation as
  ``--list-classroom50-roster``, so ``import-scores`` is available on the restricted
  ``c50_agent`` entrypoint where ``download`` stays refused. ``--classroom50-assignment``
  limits it to one slug and ``--classroom50-scores-report`` writes the JSON report.
- Marks land in ``Classroom50 Grades``, ``Classroom50 Submissions``, ``Classroom50
  Submission Details`` and ``Classroom50 Score Overrides``, keyed by assignment slug.
  Entries in ``scores.json`` are keyed by repository owner, so a group entry is spread over
  its ``member_usernames`` -- the set Classroom50 actually credits -- less the teacher, HTA
  and TA teams, which appear on every repository because the graders are organization
  owners.
- ``scores.json`` moves only when a teacher presses *Sync now*, so the file's own
  last-commit date is stored in ``Classroom50 Scores Collected At``, not the clock of the
  machine that read it. A snapshot older than an assignment's deadline records ``unknown
  (stale)`` and never ``not submitted``, because it cannot tell that apart from a submission
  made since the last collection; one taken before the assignment opened records ``not
  collected yet``.
- ``final-project`` has grading off and an empty repository, so it never appears in
  ``scores.json``, not even as an empty bucket. Asking for that slug says so and exits ``0``
  rather than failing. Its marks are entered through the existing ``--load-override-grades``
  lane instead, which gains ``Override Final Project`` and the free-text ``Override Final
  Project Reason``; both are audited.
- Added ``course --import-classroom50-groups``, which reads group membership for the
  ``mode=group`` assignments. That mode is a shared repository through collaborators and not
  a GitHub Team, so ``gh teacher team`` refuses it, and membership is the repository's
  collaborator list intersected with the classroom team: the owner is always present, and
  the filter is team membership and never permission, because a founder kept as repository
  admin to invite teammates is still a student. ``--classroom50-group-assignment`` limits it
  to one slug, ``--classroom50-read-team-json`` also reads ``team.json`` from each group
  repository, and ``--classroom50-groups-report`` writes the JSON report.
- A group whose sources disagree, which holds more than five students, or which shares a
  student with another group keeps whatever it already had. The reason is recorded in
  ``Classroom50 Group Conflicts`` and ``MiniProject_Group_Conflict``, and the rest of the
  class still imports. Findings are ``roster_audit.Issue`` values, so the consequential ones
  reach the existing announcement text without a second reporting path.
- Added ``course --import-project-issues``, which reads the mini-project topic board and
  matches proposals to students by GitHub username alone, which is all the issue form asks
  for. Nothing is written back: no label is set and no comment is posted. Staff's recorded
  choice of the canonical issue outranks the earliest-valid rule, so an old issue repaired
  later cannot displace it. ``--project-issues-repo`` names the board, or
  ``CLASSROOM50_PROJECT_ISSUES_REPO`` does; ``--project-issues-read-proposal`` also reads
  ``proposal/proposal.md`` at each pinned commit, and ``--project-issues-report`` writes the
  JSON report.
- All three imports honour ``--dry-run``, which prints the same report and stops before the
  save, and none of them creates a student: a username the database does not know is
  reported as unmatched.
- ``--export-to-excel`` grows a ``C50: <slug>`` column per collected assignment beside the
  Google Classroom and Canvas ones, and every new scalar reaches both it and
  ``--export-all-details``. The six Classroom50 containers are held back from the plain key
  loop of both detail views through one shared tuple, rather than the two that had already
  drifted apart, and both views render them from the same bilingual formatter.

Version 0.3.1 (2026-08-30)
--------------------------

- ``course-c50-admin assignment-add`` now covers autograded weekly registration:
  ``--tests``, ``--feedback-pr``/``--no-feedback-pr``, ``--pass-threshold``,
  ``--allowed-files`` (repeatable, order preserved), and ``--student-permission``.
- ``--tests`` is validated before the call: a readable JSON file holding a bare array of
  test specs. The pinned CLI's ``-`` stdin form is not offered, because no operand may
  begin with ``-``.
- ``--empty-repo`` is refused together with ``--template``, ``--tests``, ``--feedback-pr``,
  ``--allowed-files``, or ``--pass-threshold``, matching the pinned CLI, where the
  empty-repository setting is immutable after creation.
- Fixed :doc:`cli_reference`, which listed six ``assignment-add`` flags before they were
  implemented; ``--locked`` is not a flag of ``gh teacher assignment add`` at all.

Version 0.3.0 (2026-08-29)
--------------------------

- Added ``course-c50-admin``, a human-only Classroom50 operator CLI with
  ``assignment-add``, ``assignment-remove``, ``invite``, and ``download``. Each verb
  refuses in agent mode, requires an interactive terminal unless ``--dry-run`` is given,
  builds one fixed ``gh teacher`` command line, and redacts secrets from captured output.
- ``assignment-add`` refuses an existing slug, because the pinned CLI replaces the entry in
  place and cannot set submission mode, so a repeat restores every-push mode and discards a
  tagged-commit setting made in the web form. Overriding needs ``--allow-overwrite`` and an
  interactive confirmation; there is no ``--yes``.
- ``assignment-remove`` refuses an absent slug and states that student repositories survive.
- ``invite`` preflights organization targets against ``gh teacher member list``, since
  organization invitations are not idempotent; repository targets need no preflight.
- ``download --by-pattern`` is permitted only for empty-repository assignments, decided by
  reading the assignment record rather than by a blanket ban.

Version 0.2.0 (2026-08-28)
--------------------------

- Assignment creation uses one readable yes/no summary instead of exact typed
  ``AUTH``/``CREATE``/``SHARE`` phrases. ``--yes`` supports noninteractive drafts
  without Drive-sharing effects while publication, scheduling, Drive sharing, and
  general agent-mode creation continue to fail closed.
- Added a separate cooperative agent smoke-test path for one minimal draft. It binds
  an approval envelope to the account, canonical course, OAuth client, token source,
  and frozen operation; atomically requires an existing token; serializes all-state
  duplicate detection through strict read-back; and leaves general agent mutation
  refused.
- Hardened the mutation transport against connection retries, redirects, and
  response-triggered OAuth request replay. The agent-safe path also uses narrow
  response projections and requires an explicit ``ACTIVE`` course state.
- Verified the minimal no-attachment smoke workflow end to end against the exact
  allowlisted pilot course: one ``DRAFT`` was created, fetched by ID, and matched
  against the approved shape. Offline coursework, auth/transport, admin, restricted
  agent, and Classroom50 suites pass alongside all CLI parse cases.
- Added the administrator ``course-gclass-admin`` command for offline assignment
  previews, isolated OAuth authorization/status, and confirmed assignment creation.
- Added complete stable REST v1 assignment request builders for scheduling, due
  dates, grading, assignees, topics, grading periods, Drive/link/YouTube materials,
  and optional inline scored or unscored rubrics.
- Added zero-retry mutation orchestration. Rubric workflows create a draft first,
  attach the rubric, and only then publish or schedule; ambiguous and partial
  outcomes expose recovery identifiers without silently retrying or deleting.
- Course aliases are resolved before mutation, and staged releases revalidate
  scheduling and deadline constraints immediately before publication/scheduling.
- Added account-scoped JSON coursework tokens with a fixed, verified grant set,
  canonical OAuth endpoints, strict POSIX file checks, atomic writes, and per-token
  locks. The legacy pickle token is not read by this surface.
- OAuth replacement now explicitly requests offline consent and preserves existing
  storage unless Google returns a durable refresh token.
- OAuth token parsing accepts Google's additional ``openid``/``email`` identity
  scopes without enabling global relaxed-scope handling; unrelated or missing
  permissions still fail closed.
- Added a hidden-input, direct-loopback callback helper for headless remote OAuth;
  OAuth dependency logging is suppressed during the exchange so callback codes and
  bearer tokens are not emitted by verbose library loggers.
- The Google Classroom agent entrypoint now explicitly refuses
  ``create-assignment``.
- Added Classroom50 (foundation50) integration: wrap ``gh teacher`` for preflight,
  list classrooms/roster/assignments, roster sync into local DB, C50 CSV export,
  and human-only submission download (``--download-classroom50``).
- Added agent-safe entrypoints: ``python -m course_hoanganhduc.c50_agent``,
  ``canvas_agent``, ``gclass_agent``, ``db_agent`` (force agent mode; refuse
  destructive LMS/DB ops; org/course allowlists fail closed).
- Shared helpers in ``course_agent_common``; GitHub numeric id kept separate from
  Username in student identity maps (``data.py``).
- CLI flags under the **Classroom50** group (see :doc:`cli_reference`).

Version 0.1.4 (2026-01-25)
--------------------------

- Enhanced ``--list-submission-status`` to show attachment details (file count, size, type, upload time).
- Updated Canvas and Google Classroom sync to fetch detailed attachment info (supports Drive files, Links, Forms, YouTube videos).
- Updated ``--export-all-details`` to include full attachment breakdown.
- Refined default sort order for roster exports: Section > First Name > Last Name > Student ID.
- Implemented Vietnamese-specific collation for correct alphabetical sorting (e.g., handles 'Â', 'Đ' correctly).

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

Version 0.1.1 (2025-12-28)
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
