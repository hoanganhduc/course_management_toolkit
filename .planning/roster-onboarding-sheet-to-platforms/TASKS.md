# Tasks: Roster Onboarding from Sheet to Platforms

- [x] Read `gh teacher roster` syntax from the pinned CLI's own help, and settle whether
  `roster import` replaces or merges before writing any builder.
- [x] Implement `gclass_invite.invite_students_to_google_classroom` with an injectable
  service and sleep function, three-layer idempotency, and fill-only write-back.
- [x] Re-export it from the `google_classroom` facade and its `__all__`.
- [x] Wire the seven `course` flags, the dispatch block, menu entry 67 and its handler, and
  the Classroom50 next-step hint.
- [x] Hoist `gclass_agent.REFUSED_VERBS` and add `invite`.
- [x] Add `c50_roster.partition_roster_candidates` without changing `export_roster_csv`.
- [x] Add the three roster argv builders and methods to `HumanCLI`.
- [x] Implement `course-c50-admin roster-import`, including the short-circuit and the
  temporary-CSV lifecycle.
- [x] Implement `course-c50-admin roster-add`.
- [x] Refuse `roster-import` and `roster-add` on the agent entrypoint.
- [x] Write the offline tests for all three surfaces.
- [x] Update documentation, changelog, and version metadata, including the deliberate
  dry-run deviation.
- [x] Run the offline verification matrix.

## Verification evidence

- `scripts/test_gclass_invite.py` 23 tests OK; `scripts/test_c50_admin_cli.py` 93 tests OK
  (67 before this work); `scripts/test_classroom50.py` 26 tests OK (23 before);
  `scripts/test_course_agents.py` exit 0; `scripts/test_cli_flags.py` parsed 222 flags, up
  from 215, covering the seven new ones; `scripts/test_gclass_admin_cli.py` 38 OK;
  `scripts/test_gclass_coursework.py` 32 OK; `scripts/test_gclass_coursework_auth.py` 21 OK
  with 5 skipped.
- `python3 -m compileall -q course_hoanganhduc` passes; all 63 package and script files parse
  under `ast` feature version 3.9; `pyproject.toml` parses and reports `0.4.0`;
  `git diff --check` is clean.
- The lazy-import pin exits 0: importing `c50_admin_cli` does not import
  `course_hoanganhduc.data`. `scripts/test_c50_admin_cli.py` re-checks this in a subprocess,
  so the property cannot regress silently.
- The idempotency of the Classroom invite is pinned against a fake whose pending-invitation
  list grows on every create: the second run issues zero `create` calls and reports
  `invited: 0`.
- `roster-import --dry-run --csv` prints the exact argv with the real path and makes no `gh`
  call; the `--db` form prints a `plan` with `csvSource: generated-from-db` and no `argv`.
- Under `COURSE_C50_AGENT_MODE=1`, `roster-import --dry-run` returns
  `{"code": "agent_forbidden"}` with exit 2. `python -m course_hoanganhduc.c50_agent
  roster-import` returns `roster import is not available in agent mode`, and
  `python -m course_hoanganhduc.gclass_agent invite` returns `invite is not available on the
  agent surface`.
- `export_roster_csv` output is pinned byte for byte on a fixed fixture, so extracting
  `partition_roster_candidates` out of it cannot have changed `--export-classroom50-roster`.
- Installed into `~/.course_venv` (`pip install -e . --no-deps`, reporting
  `course-hoanganhduc-0.4.0`) and exercised through the console scripts against a real
  three-student `students.db`. The `--db` form loaded the database through the real
  `load_database`, counted 2 exportable rows, and reported the third under
  `skippedNoUsername`; `course --export-classroom50-roster` wrote the same two rows with the
  canonical header and no BOM, and feeding that file back through `roster-import --csv
  --dry-run` produced the matching argv. Running the real `load_database` exposed one defect
  the stubbed tests could not: it prints progress on stdout, which corrupted the JSON report.
  It is now redirected onto the error stream and pinned by
  `test_database_progress_output_stays_off_stdout`.
- `--invite-google-classroom` reaches `gclass_invite.py` credential loading before failing on
  the absent `credentials.json`, confirming the facade export, the seven flags, and the
  dispatch are wired. The failure is an uncaught `FileNotFoundError`; the shipped
  `--unenroll-google-classroom` fails identically on the same input, so the new flag matches
  the surrounding convention rather than introducing this.

## Conformance against the platform documentation

Checked after implementation, against primary sources rather than recollection.

**Classroom50** -- originally the installed `gh teacher` v1.25.1 help output and the v1.40.0
Go source under `classroom50-web-v1.40.0/cli/gh-teacher/`; re-checked against v1.45.0's own
help after the upgrade below. Every item holds on both versions:

- `roster import <org> <classroom> <path>` and `roster add <org> <classroom> <username>` match
  the argv the builders emit, as do the flag names `--first-name`, `--last-name`, `--email`,
  `--section`.
- v1.25.1's own help states "upsert every row" and "After the commit lands, any student who
  isn't already in the org (and doesn't have a pending invite) is invited", confirming on the
  pinned version what was previously read only from v1.40.0 source.
- The header contract holds: `configrepo.RosterColumns` is
  `username,first_name,last_name,email,section,github_id,role` and `ParseImportCSV` accepts
  that, the 6-column prefix, or the 5-column prefix. `export_roster_csv` emits the 6-column
  form. On v1.25.1 `github_id` is ignored on input; v1.40.0 cross-checks it.
- v1.40.0 calls `TrimUTF8BOM`; v1.25.1 carries no such symbol. Writing the CSV without a BOM
  is therefore required on the pinned version, not merely tidy.
- `roster list --json` emits `role` as a seventh field on v1.25.1; `parse_roster_payload`
  drops it deliberately, and nothing writes it.

### The v1.45.0 upgrade

`foundation50/gh-teacher` releases fast: v1.25.1 (2026-08-04) was installed and pinned;
v1.45.0 (2026-09-04) is current. The v1.40.0 tree read for semantics was a source snapshot,
never a requirement. On 2026-09-05 the extension was upgraded to v1.45.0 and re-pinned there,
on the operator's instruction, because two upstream fixes land after v1.25.1 and bear on this
feature:

- **v1.34.0 (PR #746), non-ASCII names.** Before the fix, `roster import` "passed
  invalid-UTF-8 bytes through verbatim", committing a Windows-1252 Excel export's diacritics
  into `roster.csv` as permanent mojibake. v1.45.0 decodes the fallback and prints "Notice: %s
  is not UTF-8 encoded and was read as Windows-1252", a symbol confirmed present in the binary.
  Correcting an earlier claim on this page: neither toolkit path was ever exposed. `--db`
  writes valid UTF-8, and `_csv_rows_from_file` opens an operator-supplied `--csv` with
  `encoding="utf-8"`, so a Windows-1252 file is refused with `value_error` and exit 2 before
  `gh` is reached. The fix therefore only helps someone typing `gh teacher roster import`
  directly. Refusing is kept deliberately here, because a transliterated Vietnamese name would
  be committed to GitHub; `course --add-csv`, which only writes to the local database, makes
  the opposite choice and reports the encoding it fell back to.
- **v1.35.0 (PR #773), concurrent invite acceptance.** A student accepting mid-pass was
  classified as "accepted, then unenrolled": their invite-metadata team was deleted, an
  identity-only duplicate row appended, and the teacher's original row reaped, losing names
  and emails. A whole-class `roster-import` is precisely what produces many simultaneous
  acceptances, so this toolkit's main path is what triggered it. This was the deciding
  reason to upgrade.

Compatibility was established before upgrading, not assumed. `internal/configrepo/students_csv.go`,
`internal/roster/roster.go` and `internal/roster/list.go` are **byte-identical** between
v1.40.0 and v1.45.0, so the evidence gathered against v1.40.0 applies unchanged to the
installed build. v1.41.0-v1.45.0 touch assignments, autograders and `tests.json`, not the
roster. Re-verified on the installed v1.45.0: the three argv shapes, the flag names, "upsert
every row", the org-invite side effect, the six-column header (accepted as the canonical
seven minus `role`), and `roster list --json` emitting `role` -- which `parse_roster_payload`
still drops by design. The `github_id` cross-check "fails that line" only when the cell names
a *different* account; this toolkit emits it empty. No toolkit code changed for the upgrade,
and the full offline suite is green on it.

Rollback: the v1.25.1 binary and manifest are backed up in the session scratchpad;
`gh extension remove teacher && gh extension install foundation50/gh-teacher --pin v1.25.1`
restores it. `gh student` was left at v1.25.1 -- the toolkit never invokes it.

### Divergence after the v1.45.0 upgrade

One shipped assumption in `c50_cli_human.py` no longer matches the CLI, found by re-probing
after the upgrade:

- **Assignment slug cap.** `_SLUG_RE` is `^[a-z0-9][a-z0-9-]{1,38}$`, matching v1.25.1.
  v1.33.0 (PR #693) raised the CLI's cap to `{1,99}`, so the toolkit now refuses slugs
  v1.45.0 would accept. Left unchanged deliberately: v1.33.0 also added a per-repo name
  budget (student repos are `<classroom>-<slug>-<username>`) that this regex does not model,
  so the stricter bound fails early rather than at the network call. Raising it is an
  operator decision, and would mean modelling that budget rather than swapping a number.
- **`--empty-repo` conflicts.** v1.45.0 rejects `--submission-mode` and `--submission-tag`
  alongside `--empty-repo`, beyond the five v1.25.1 rejected. `assignment_add_argv` cannot
  emit either flag, so the local guard stays complete for everything it can produce. Comment
  updated; no code change.
- `--feedback-pr` is still "default true, pass `=false` to disable" on v1.45.0, so that
  builder branch is unchanged.

**Google Classroom** -- the REST reference. Three defects were found and fixed:

- `invitations.create` documents `400 FAILED_PRECONDITION` for "the user already has this
  role or a role with greater permissions" or a disabled account. The code counted every
  `400` as a failure, so an enrolment the bulk read missed was reported as a failed write.
  `409 ALREADY_EXISTS` covers only a pending invitation, so this was a real hole in the
  third idempotency layer. Now `skipped_precondition`, classified from `error.status`.
- `404` is documented as "the course or the user does not exist"; the detail blamed only the
  student's Google account. It now names the course too.
- The pacing comment cited "roughly 50 writes / 10s / project", which is not a documented
  figure. The published limits are 3,000 queries/minute/project and 1,200/minute/user on a
  60-second moving average. The interval was already well inside both; only the stated
  reason was wrong.

Confirmed correct as written: `invitations` is a top-level resource
(`POST https://classroom.googleapis.com/v1/invitations`), not nested under `courses`; the
body fields `courseId`, `userId`, `role` and the `STUDENT`/`TEACHER` enum values; `409` for
an existing invitation; `403` as a course-wide condition, which is what justifies aborting
the loop; `invitations.list` default `pageSize` 500; and the scopes already present in
`gclass_auth.SCOPES` (`classroom.rosters` for invitations, `classroom.profile.emails` for
`userProfiles.get`), so no scope change and no token deletion.

One documented mismatch is left alone: `courses.students.list` has a default `pageSize` of 30
and no published maximum, and the code requests 200. The reference warns the server may return
fewer than requested; the paginator follows `nextPageToken` and does not assume a page size, so
the request is a hint and cannot lose students.

## Deliberate deviation from the approved plan

The plan put the `roster list` read before the `--dry-run` branch so that a dry run could
report `new` and `remoteOnly`. It ships the other way: `--dry-run` returns before any `gh`
call. Every other verb on this CLI does that, and `--dry-run` is the one mode
`_require_human` allows without a terminal, so a dry run that needed credentials and a
network would be a worse inconsistency than the one it fixed. The dry-run payload therefore
carries `csvSource`, `rowCount`, `firstUsernames`, `skippedNoUsername`, and an explicit
`rosterDiff: "not computed: --dry-run makes no gh calls"` instead of `new` and `remoteOnly`.
The real run still computes and reports both.

## Unchecked

These are operator gates. Nothing below has been executed; all of it needs credentials, a
network, and an authorized human.

- [ ] **Non-destructive gate (blocks release).** Source and binary now agree on v1.45.0,
  but the upsert has still never been run against a real organization. On a scratch org:
  `gh teacher roster add SCRATCH cls sentinel-user`, then `course-c50-admin roster-import`
  with a database that does *not* contain `sentinel-user`, then
  `gh teacher roster list SCRATCH cls --json` — `sentinel-user` must still be there.
- [ ] **Classroom idempotency gate.** Run `--invite-google-classroom` twice against a test
  course; the second run must report `invited: 0` and make no `create` call, checked with
  `--verbose`.
- [ ] **Classroom50 idempotency gate.** Run `roster-import` twice; the second run must print
  `roster already matches the local database; nothing to import` and make exactly one `gh`
  call.
- [ ] **OAuth scope gate.** Confirm `token.pickle` survives the first invite, proving
  `gclass_auth.SCOPES` was not touched — a changed scope set deletes the token.
- Classroom write pacing (3,000 queries/minute/project, 1,200/minute/user, 60-second moving
  average) has not been exercised at class scale. Bounded backoff covers a burst; beyond
  that, candidates land in `failed` and the re-run is safe.
- Whether the token in use holds `admin:org` was not inspected. That affects the existing
  `invite` preflight, not the roster verbs, which do not read pending organization
  invitations.

## Onboarding driver (`course --onboard`)

The write half above needs a course id, an org, a classroom, and a database that already holds
the students. This closes the read half: one interactive pass from the registration form to
both platforms.

- [x] Teach `roster_csv` the Google Form header shape: parenthesised bilingual labels, the
  required-question asterisk, ranked aliases so an attribute takes the best column claiming
  it, and the `vnu`+`hus` rule that also covers the form's `Emai VNU-HUS` typo.
- [x] Add `Personal Email`, `Date of Birth`, and `Timestamp` as process-only columns, read
  while resolving and never written, so `--add-csv` keeps writing what it wrote before.
- [x] Add `onboard.py`: email and GitHub username validation, `gh api /users/<name>`
  existence checks, per-row rejection with reasons, `Student ID` dedupe, and the
  GitHub-username clash rule.
- [x] Read both platforms before any write and render the comparison table: three Google
  reads plus a per-candidate `invitations.list(userId=…)`, and the two independent
  Classroom50 reads plus `roster sync` in reporting mode.
- [x] Merge into the database, then invite through the existing Google path, then call
  `c50_admin_cli.main(["roster-import", …])` unchanged.
- [x] Add the two pickers and reuse `utils.input_with_completion` for the file prompt.
- [x] Wire `--onboard`, `--onboard-csv`, `--onboard-skip-github-check`, reusing the existing
  course, org, classroom, and report flags.
- [x] Write `scripts/test_onboard.py` and the regression pins in `scripts/test_roster_csv.py`
  and `scripts/test_classroom50.py`.
- [x] Update documentation and changelog.

### Verification evidence

- `scripts/test_onboard.py` 53 OK; `scripts/test_roster_csv.py` 13 OK (10 before, +3 pinning
  the pre-existing aliases and `--add-csv`); `scripts/test_classroom50.py` 31 OK (26 before,
  +5 reading `gh teacher roster import --help` on the pinned v1.45.0);
  `scripts/test_c50_admin_cli.py` 97 OK; `scripts/test_gclass_invite.py` 28 OK;
  `scripts/test_course_agents.py` exit 0; `scripts/test_gclass_admin_cli.py` 38 OK;
  `scripts/test_gclass_coursework.py` 32 OK; `scripts/test_gclass_coursework_auth.py` 21 OK
  with 5 skipped; `scripts/test_cli_flags.py` parsed 226 flags, up from 223, covering the
  three new ones.
- `python3 -m compileall -q course_hoanganhduc` passes, and importing `onboard` pulls in
  neither pandas nor the database layer.
- Run by hand on a form-shaped CSV built from a real registration header: `--dry-run` made no
  `gh` call and no write; a real run made only read calls (`gh api /users/<name>`), printed
  the table, and wrote the expected attributes to a scratch database. No invitation was sent
  and no roster was imported.

### Deliberate deviation from the approved plan

- Repeat submissions collapse to the **last row**, not to the newest `Timestamp`. Google
  Forms appends, so a later row is a later submission by the same student, and a form date
  like `9/5/2026` cannot be read without guessing a locale.
- `--dry-run` is fully offline, so the comparison table is a step of the real run, printed
  before its confirmation, rather than something a dry run can preview.
- The `Student ID` → `GitHub Username` → `Email` dedupe key chain is unreachable from the CSV
  path: `parse_student_rows` already drops a row missing `Student ID` or `Name`. The chain
  stays for `resolve_rows`' other callers and is pinned by a test that builds a `ParseResult`
  directly.

### Unchecked

- [ ] **Pending-link gate (blocks the first real import).** Run
  `gh teacher roster sync ORG CLS` in reporting mode. Exit 2 means links are pending: run it
  with `--write` before onboarding, or a roster row awaiting its email link is imported a
  second time for the same person.
- [ ] **Diacritics gate.** Onboard **one** student whose name carries Vietnamese diacritics,
  then `roster list` and read the stored name back. The Python side writes UTF-8 without a
  BOM; the Go side and the git commit were never verified here.
- [ ] **Picker gate.** Confirm both pickers list the real courses and classrooms, and that
  choosing `0` in either leaves that platform untouched.
- [ ] **Idempotency gate.** Onboard the same form twice. The second run must invite nobody
  and add no roster row.
