# Specification: Roster Onboarding from Sheet to Platforms

## Goal

Close the write half of the start-of-term roster path. The toolkit already reads a Google
Sheet into the local student database (`course --add-google-sheet`), but had no way to put
those students onto either teaching platform: Google Classroom had only a pull sync and an
unenroll, and Classroom50 had only `roster list`. Enrollment was therefore done by hand in
two web interfaces. This work adds `Sheet -> database -> Google Classroom` and
`database -> Classroom50 roster + organization`, with the hard requirement that a student
already present on a platform is skipped there, so a repeated run neither duplicates nor
destroys anything.

## Scope

- In scope: `invite_students_to_google_classroom` and its `course` flags, menu entry, and
  next-step hint; `c50_roster.partition_roster_candidates`; `HumanCLI.roster_list`,
  `roster_import`, and `roster_add`; `course-c50-admin roster-import` and `roster-add`;
  agent refusals for all three new verbs; offline tests.
- Out of scope: unenrolling or removing anyone; `gh teacher roster remove` and
  `roster update`; changes to `invite_users` or `export_roster_csv`; reading the Sheet
  directly at write time; any agent-reachable write path; a single orchestrating command.

## Assumptions

- The installed CLI is the pinned `gh teacher` v1.45.0, commit `a5a51293b65d`, upgraded
  from v1.25.1 on 2026-09-05. Subcommand syntax comes from that build's own help output.
  The extension stays pinned rather than tracking latest.
- `gh teacher roster import` is an upsert, not a replacement. This was read from the
  Classroom50 source at `cli/gh-teacher/internal/roster/roster.go:693-730`, whose roster
  files are byte-identical between v1.40.0 and v1.45.0: it loads the existing roster,
  upserts each incoming row, and returns an empty change when nothing was added or updated.
  The installed binary's own help states "upsert every row", so source and binary now agree
  on one version. Gate 1 below still verifies the behaviour directly.
- `roster import` also invites: once the commit lands, a listed student who is not already
  an organization member and holds no pending invitation is invited. No separate invitation
  loop is needed, and the confirmation must name this second effect.
- `roster import` does not preserve trailing columns the way `roster add`/`roster update` do,
  so remote rows are never re-emitted through this toolkit's CSV dialect; doing that would
  erase the `role` column the remote holds.
- Google Classroom authorization keeps the existing token-pickle path. `classroom.rosters`
  and `classroom.profile.emails` are already in `gclass_auth.SCOPES`, and changing that set
  deletes every user's token.
- The local database is the only source of candidates for both platforms.
- Python 3.9 compatibility is required.

## Interfaces

- `course_hoanganhduc.gclass_invite.invite_students_to_google_classroom`, re-exported
  through `google_classroom` because `core` imports that facade with a star import.
- `course_hoanganhduc.c50_roster.partition_roster_candidates`.
- `course_hoanganhduc.c50_cli_human.HumanCLI`: `roster_list_argv`, `roster_import_argv`,
  `roster_add_argv`, `roster_list`, `roster_import`, `roster_add`.
- `course_hoanganhduc.c50_admin_cli`: `roster-import`, `roster-add`.
- `course` flags: `--invite-google-classroom` (`-igc`), `--gc-invite-email`,
  `--gc-invite-domain`, `--gc-invite-section`, `--gc-invite-class`, `--gc-invite-role`,
  `--gc-invite-all`; menu action 67.

## Acceptance Criteria

- A second `--invite-google-classroom` run over an unchanged course issues no
  `invitations().create` call and reports `invited: 0`.
- Idempotency is layered: the paginated student and teacher lists supply enrolled emails and
  user ids; pending invitations are matched against the locally stored `Google_ID` before any
  `userProfiles().get`, and that fan-out is capped by `profile_lookup_limit`; a `409
  ALREADY_EXISTS` from `create` is recorded as `skipped_pending`, not as a failure.
- A `403` stops the invite loop and reports the remaining candidates as `skipped_aborted`,
  because the condition is course-wide rather than per student.
- A `400 FAILED_PRECONDITION` is recorded as `skipped_precondition`, never as a failure: the
  reference gives its causes as the user already holding this role or greater, or a disabled
  account. Classification reads `error.status` from the body, so a plain `400` stays a failure.
- A `404` is a per-student failure whose detail names the missing Google account and the
  course, since the reference defines it as "the course or the user does not exist". It is
  never retried. `429`, `500`, and `503` are retried with bounded backoff.
- An address passed to `--gc-invite-email` that is not in the database is reported as
  `skipped_not_in_db` pointing at `--add-google-sheet`, never invited.
- Database write-back is fill-only: `Google_ID`, `Google_Invitation_ID`, and
  `Google_Classroom_Invited_At` are set only when empty, under
  `audit_source="gclass-invite"`, and never during a dry run.
- `course-c50-admin roster-import` reads the remote roster first and, when nothing would
  change, prints `roster already matches the local database; nothing to import` and returns
  0 having made exactly one `gh` call.
- Rows present on the remote roster but absent from the CSV are reported under `remoteOnly`
  and are neither re-sent nor removed.
- Students with no GitHub username are reported under `skippedNoUsername` and do not change
  the exit code; the operator sees the count and the names in the confirmation first.
- The CSV generated from `--db` is UTF-8 without a BOM, carries the six canonical columns,
  lives in a `tempfile.mkdtemp` directory, and is removed even when `gh` fails, unless
  `--keep-csv` is given.
- `--dry-run` on both new verbs makes no `gh` call at all, matching every other verb on this
  CLI and keeping `--dry-run` the one mode `_require_human` allows without a terminal.
- A failed `gh` roster call exits 3 as `outcome_unknown`, because the roster is committed
  before invitations are sent.
- `python -m course_hoanganhduc.c50_agent` refuses `roster-import` and `roster-add`, and
  `python -m course_hoanganhduc.gclass_agent` refuses `invite`, with structured refusals.
- `export_roster_csv` output is unchanged byte for byte.

## Verification

- New offline `unittest` script `scripts/test_gclass_invite.py` with an injected Classroom
  service whose pending-invitation list grows on every create.
- `scripts/test_c50_admin_cli.py` and `scripts/test_classroom50.py`, extended.
- `scripts/test_course_agents.py` and `scripts/test_cli_flags.py` regressions.
- `python3 -m compileall -q course_hoanganhduc` and Python 3.9 grammar parsing.

## Risks

- The merge conclusion was read from v1.40.0 source while v1.25.1 is installed. Nothing here
  parses `gh` output, and gate 1 in TASKS.md checks the behaviour directly.
- `roster import` sends invitations as a side effect of a command named "import". The
  confirmation names both effects; an operator who skims it still sends mail.
- A Classroom `Invitation` carries a numeric `userId` and never an email, so deduplicating by
  address needs a `userProfiles().get` fan-out. Mitigated by matching `Google_ID` first,
  capping the fan-out, keeping the 409 backstop, and writing `Google_ID` back so the second
  run needs none. This is a property of the API, not of this repository.
- Classroom's published quota is 3,000 queries/minute/project and 1,200/minute/user, checked
  on a 60-second moving average; a large class can reach 429. Bounded backoff covers a
  burst, and the rest is a safe re-run.
- `core` dispatch is a flat sequence of `if` blocks with no early exit, so
  `--invite-google-classroom --unenroll-google-classroom` runs both. This is the existing
  behaviour of every flag pair in that file and was not changed here.
