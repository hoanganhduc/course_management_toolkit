# Task Plan: Roster Onboarding from Sheet to Platforms

## Context

`--add-google-sheet` fills the local database and deduplicates on the way in, but nothing
carried those students onward. A repository-wide search found no call to
`invitations().create`, and the Classroom50 adapter wrapped only `roster list`. The two
halves are independent and were built in parallel; the shape of each was already set by
precedent in this repository — `gclass_unenroll` for the Classroom flow, and
`course-c50-admin` for anything that changes remote Classroom50 state.

## Steps

1. Add `gclass_invite.invite_students_to_google_classroom` with an injectable service and
   sleep function, three-layer idempotency, and fill-only database write-back.
2. Re-export it from the `google_classroom` facade, including `__all__`.
3. Wire seven flags, a dispatch block, menu entry 67 and its handler, and the
   Classroom50 next-step hint into `core`.
4. Hoist `gclass_agent`'s duplicated refusal list into one constant and add `invite`.
5. Add `c50_roster.partition_roster_candidates` by extracting the username resolution
   `export_roster_csv` already used, leaving that function's output unchanged.
6. Add the three roster argv builders and their thin methods to `HumanCLI`.
7. Add the three outcome-unknown codes and the `roster-import` subcommand and branch to
   `c50_admin_cli`.
8. Add `roster-add`.
9. Add both verbs to `c50_agent.REFUSED_VERBS`.
10. Write the offline tests for all three surfaces.
11. Update documentation, changelog, and version metadata.

## Decisions

| Decision | Rationale | Status |
|---|---|---|
| Classroom keeps the existing token-pickle path | The two needed scopes are already granted; changing `SCOPES` deletes every user's token | Locked |
| The database is the only candidate source | One source of truth, filterable and auditable | Locked |
| `--gc-invite-email` filters the database, never bypasses it | A typo is reported as `skipped_not_in_db` instead of silently inviting a stranger | Locked |
| Classroom invite is a `course` flag; the roster write is a `course-c50-admin` subcommand | Each lands on the gate structure its blast radius needs | Locked |
| `roster import`, not a separate organization invite loop | One command upserts the roster and invites, and is idempotent | Locked |
| Send a plain CSV from the database, never re-emit remote rows | `import` does not preserve trailing columns the way `add`/`update` do; re-emitting would erase `role` | Locked |
| `invite_users` keeps its current shape | It changes organization membership only, leaving the roster empty — half the goal | Locked |
| Injectable `service` and `sleep_fn` | The unenroll template builds its own service and is therefore untested; this is what makes the invite testable | Locked |
| Return a report dict, not an `int` | Deleting has one outcome, inviting has six; `canvas_invites` set this precedent | Locked |
| Import `load_database` lazily | `data` pulls pandas and friends; both new surfaces must stay importable without them | Locked |
| Retry `invitations().create` on transient status | An invitation deduplicates itself through its own 409, unlike a coursework create | Locked |
| Stop the loop on 403 | It is a course-wide policy failure; continuing yields identical failures and burns quota | Locked |
| Read the remote roster and short-circuit when nothing changes | Saves a round trip and, more importantly, sends no invitations | Locked |
| Missing GitHub username exits 0 | It is a deliberate non-write the operator confirmed, not a failed write; a non-zero exit here would train operators to ignore exit codes | Locked |
| `--dry-run` makes no `gh` call | Matches every other verb on this CLI and keeps `--dry-run` usable without a terminal; the cost is that the roster diff cannot be shown, so a `plan` is reported instead | Locked |
| Pass `gh` stderr through verbatim, never parse it | The partial-onboarding warning is version-specific | Locked |
| Delete the temporary CSV even on failure | It is reproducible from the database; leaving student emails in a temp directory is worse | Locked |
| No new function in `c50_ops` | That module holds policy; this verb's policy is `_require_human` plus `_confirm`, already in the CLI | Locked |
| Not on the `course` CLI | `dispatch_classroom50` has no human gate and is reachable from `c50_agent` | Locked |
| No single orchestrating command | See below | Locked |

### Why there is no orchestrator

Four reasons, in order of weight. The gates do not reconcile: the Classroom half may open a
browser for OAuth and prompt, while the Classroom50 half refuses unless agent mode is off
*and* stdin is a terminal, so one process means either relaxing `_require_human` — the core
invariant of that CLI — or shelling out, which is a runbook with extra steps. The exit-code
contracts collide: `course` returns nothing meaningful, `course-c50-admin` has a 0/2/3/4
contract pinned by tests. Idempotency already solves re-running, which is the only thing an
orchestrator's checkpointing would buy. And one command writing to two platforms after a
single `y` is a worse artifact than two commands with two confirmations, each naming its own
platform and counts. What ships instead is three commands and a printed, pre-filled next
step after a successful invite.

## Verification Plan

| Check | Command | Expected result |
|---|---|---|
| Classroom invite tests | `python3 scripts/test_gclass_invite.py` | Pass offline |
| Admin CLI tests | `python3 scripts/test_c50_admin_cli.py` | Pass |
| Classroom50 regression | `python3 scripts/test_classroom50.py` | Pass |
| Agent regression | `python3 scripts/test_course_agents.py` | Pass |
| CLI flag regression | `python3 scripts/test_cli_flags.py` | Pass, new flags parsed |
| Lazy-import pin | `python3 -c "import course_hoanganhduc.c50_admin_cli, sys; raise SystemExit('course_hoanganhduc.data' in sys.modules)"` | Exit 0 |
| Syntax/import | `python3 -m compileall -q course_hoanganhduc` | Pass |
| Python 3.9 grammar | Parse package and scripts with `ast` feature version 3.9 | Pass |

## Out of Scope

No unenrollment or roster removal, no `roster remove`/`roster update`, no change to
`invite_users` or `export_roster_csv`, no direct Sheet read at write time, no agent-reachable
write path, no orchestrator. Live execution stays behind the manual gates in TASKS.md.
