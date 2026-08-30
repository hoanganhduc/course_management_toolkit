# Specification: Classroom50 Operator Surface

## Goal

Add a human-only Classroom50 administration surface covering assignment registration,
assignment removal, organization/repository invitation, and submission download, so that
the reviewed teacher-side operations in the HK1 2026-2027 course documents can be executed
through an approved adapter instead of dead-ending at `human/operator required`. The agent
entrypoint stays read-only.

## Scope

- In scope: a dedicated `course-c50-admin` console script; fixed-argv `HumanCLI` methods for
  `assignment add`, `assignment remove`, `invite`, `member list`, and `download`; an
  existing-slug guard on `assignment-add`; an empty-repository gate on `download --by-pattern`;
  an org-membership preflight on `invite`; a no-network `--dry-run` on every verb; structured
  refusal stubs on the agent entrypoint; offline tests.
- Out of scope: student acceptance and student submission, which run under a student's own
  credentials and are forbidden to an implementation agent by the runbook operator boundary;
  any agent-safe mutation path; roster push or write-back; `init`, `teardown`,
  `rotate-service-token`, `staff`, `autograder`, `audit`, `login`, `logout`, `remove`, and
  `classroom add/remove/migrate`.

## Assumptions

- The installed CLI is the pinned `gh teacher` v1.25.1, commit
  `b9fbf2a8526ca7b20cc50c27d5419f6e07eaeea6`. Subcommand syntax, flags, and idempotency
  behaviour are taken from that build's own help output and the pinned
  `schemas/assignments-v1.schema.json`.
- Python 3.9 compatibility is required.
- Existing Classroom50 read, sync, and export behaviour remains unchanged.
- Remote writes remain gated by the runbook: this surface makes an operation *reachable*,
  it does not authorize it.

## Interfaces

- `course_hoanganhduc.c50_cli_human.HumanCLI`: `assignment_add`, `assignment_remove`,
  `invite`, `member_list`, `list_assignments`, `download`.
- `course_hoanganhduc.c50_ops`: `find_assignment`, `assignment_allows_by_pattern`,
  `existing_member_logins`, `agent_refuse`.
- `course_hoanganhduc.c50_admin_cli`: `assignment-add`, `assignment-remove`, `invite`,
  `download`.
- Console script: `course-c50-admin`.

## Acceptance Criteria

- Every mutating verb refuses in agent mode with `agent_forbidden` and refuses without a
  TTY unless `--dry-run`.
- `--dry-run` prints the exact argv and the preconditions that will be checked, performs no
  subprocess call, and exits 0.
- `assignment-add` reads the assignment manifest first and refuses an existing slug with
  `assignment_exists`, naming the submission-mode hazard. Overwrite requires both
  `--allow-overwrite` and an interactive confirmation; there is no `--yes`.
- `assignment-add` rejects a slug outside `^[a-z0-9][a-z0-9-]{1,38}$`, rejects
  `--empty-repo` together with `--template`, and requires `--max-group-size >= 2` with
  `--mode group`.
- `assignment-remove` refuses an absent slug with `assignment_absent` rather than exiting 0,
  and its confirmation states that student repositories are not deleted and that re-adding
  the same slug is not a clean reset.
- `download --by-pattern` is permitted only when the assignment manifest records
  `empty_repo: true`; otherwise it fails with `by_pattern_not_permitted`. Plain `download`
  is unaffected.
- `invite` to an organization lists actual membership first and skips logins that are
  already members or hold a pending invitation, reporting them as skipped instead of
  failing the batch. Repository invites are not preflighted, since they are idempotent.
- No argument that begins with `-` reaches the `gh` argv, and no free-form argv is
  constructible through the adapter.
- Command stderr is redacted on every new path before it is raised or printed.
- The agent entrypoint answers `assignment-add`, `assignment-remove`, `invite`, and
  `download` with a structured refusal, not an argparse usage error.

## Verification

- New offline `unittest` script `scripts/test_c50_admin_cli.py` using injected runners.
- Existing `scripts/test_classroom50.py`, extended for the new refusals and the
  `--by-pattern` gate.
- `scripts/test_course_agents.py` and `scripts/test_cli_flags.py` regressions.
- `python3 -m compileall -q course_hoanganhduc`.

## Risks

- The adapter cannot restore a submission mode that only the web form can set; the
  existing-slug guard is a guard, not a repair.
- `gh teacher assignment remove` followed by `add` bypasses the CLI's own immutability
  guard on `empty_repo`; the removal confirmation warns, but nothing prevents an operator
  from doing it in two steps.
- Reading pending organization invitations requires the `admin:org` scope; without it the
  preflight may under-report and an org invite can still fail.
- Same-UID callers can bypass these application-level guards, as with the existing surfaces.
