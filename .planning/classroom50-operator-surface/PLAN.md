# Task Plan: Classroom50 Operator Surface

## Context

`AgentCLI` exposes four read verbs of the roughly twenty `gh teacher` offers, and
`HumanCLI` exposes only `download` without `--by-pattern`. Every mutating step in the
course runbook therefore falls through to its "approved adapter lacks the operation"
branch and stops. The `gclass` admin CLI already established how this repository adds a
mutating surface: a dedicated console script, human-only, with the agent entrypoint still
refusing.

## Steps

1. Add pure preflight helpers to `c50_ops` and offline tests that pin their behaviour.
2. Extend `HumanCLI` with fixed-argv `assignment_add`, `assignment_remove`, `invite`,
   `member_list`, and `list_assignments`, and narrow the `--by-pattern` rule on `download`.
3. Implement `c50_admin_cli` with argparse subcommands, agent/TTY gates, `--dry-run`,
   interactive confirmation, and the four exit codes.
4. Register structured refusal stubs for the three new mutating verbs on the agent
   entrypoint.
5. Register the console script, bump the version, and document the surface.
6. Run the offline verification matrix.

## Decisions

| Decision | Rationale | Status |
|---|---|---|
| Dedicated admin CLI | Mirrors `course-gclass-admin`; keeps mutation off the agent entrypoint | Locked |
| No agent-safe mutation path | The runbook routes these operations to an authorized human | Locked |
| Fixed argv per verb | Preserves the closed-map property; no free-form command construction | Locked |
| Refuse existing slug by default | `assignment add` is an in-place upsert and resets submission mode under the pinned CLI | Locked |
| `--allow-overwrite` plus confirmation, never `--yes` | Overwrite is the destructive case; it must stay human | Locked |
| Refuse absent slug on remove | The CLI exits 0 silently; an operator needs to know nothing matched | Locked |
| `--by-pattern` gated on `empty_repo`, not banned | Autograded assignments need the scores summary; empty-repo assignments have none | Locked |
| Refuse `--by-pattern` when the slug is absent from the manifest | The gate cannot be evaluated, so it fails closed | Locked |
| Org invite preflight, repo invite not | GitHub rejects org re-invites; repo invites are idempotent | Locked |
| Skip, do not fail, on already-present logins | One known member must not abort a batch | Locked |
| Reject flag-like argument values | Prevents an operand from being reinterpreted as a `gh` flag | Locked |
| `--dry-run` makes no subprocess call | Keeps the runbook's "prints commands by default" steps executable offline | Locked |
| Student acceptance and submission excluded | They require the student's own credentials and are forbidden to an agent | Locked |

## Verification Plan

| Check | Command | Expected result |
|---|---|---|
| Admin CLI tests | `python3 scripts/test_c50_admin_cli.py` | Pass offline |
| Classroom50 regression | `python3 scripts/test_classroom50.py` | Pass |
| Agent regression | `python3 scripts/test_course_agents.py` | Pass |
| CLI flag regression | `python3 scripts/test_cli_flags.py` | Pass |
| Syntax/import | `python3 -m compileall -q course_hoanganhduc` | Pass |
| Python 3.9 grammar | Parse package and scripts with `ast` feature version 3.9 | Pass |

## Out of Scope

No student acceptance or submission, no agent-safe mutation path, no roster write-back, no
live classroom call. Live execution remains gated by the course runbook and belongs in the
disposable pilot classroom.
