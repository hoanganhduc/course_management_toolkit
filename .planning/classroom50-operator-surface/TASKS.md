# Tasks: Classroom50 Operator Surface

- [x] Confirm the pinned CLI's subcommand syntax and idempotency behaviour from its own help.
- [x] Add failing offline tests for the preflight helpers, adapter argv, and CLI gates.
- [x] Implement `c50_ops` preflight helpers and the shared agent refusal.
- [x] Extend `HumanCLI` and narrow the `--by-pattern` rule.
- [x] Implement `c50_admin_cli` and register the console script.
- [x] Add the three agent refusal stubs.
- [x] Update documentation, changelog, and version metadata.
- [x] Run the offline verification matrix.

## Verification evidence

- `scripts/test_c50_admin_cli.py` 52 tests OK, `scripts/test_classroom50.py` 23 tests OK,
  `scripts/test_course_agents.py` 6 tests OK, `scripts/test_cli_flags.py` 215 flags parsed.
- `compileall`, Python 3.9 grammar parsing of the package and scripts, `pyproject.toml` TOML
  parsing, and `git diff --check` all pass.
- `assignment-add --dry-run` reproduces the reviewed command at line 1218 of
  `classroom50-assignment-plan-hk1-2026-2027.md` argument for argument.
- All four verbs on `python -m course_hoanganhduc.c50_agent` return the structured
  `<verb> is not available in agent mode` refusal with exit 1, and the admin CLI under
  `COURSE_C50_AGENT_MODE=1` returns `{"code": "agent_forbidden"}` with exit 2.

## Unchecked

- Every `gh teacher` mutation path is offline-verified only. No live `assignment add`,
  `assignment remove`, `invite`, or `download` was executed; those require an authorized
  operator in `classroom50-pilot-2026` under the runbook remote-write and human gates.
- The pinned CLI's syntax was read from `gh teacher --help` at v1.25.1. A different installed
  version invalidates the fixed argv builders.
- `member list` needs the `admin:org` scope to report pending invitations; the installed
  token's scopes were not inspected, so the organization invite preflight may under-report.
