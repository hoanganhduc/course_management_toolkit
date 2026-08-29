# Task Plan: Google Classroom Assignment Creation

## Context

The package has Google Classroom read/sync/grading helpers but no coursework create
operation. Legacy auth uses broad scopes, inconsistent locations, and pickle tokens;
the new human mutation surface must be isolated from it.

## Steps

1. Add offline tests defining minimal/full bodies, invalid combinations, staged rubric
   behavior, ambiguous outcomes, auth path precedence, secret-file checks, and dry-run.
2. Implement pure coursework/material/rubric builders and mutation orchestration.
3. Implement deterministic credential paths, secure JSON loading/writing, OAuth endpoint
   sanitization, token/account verification, and auth status.
4. Implement the dedicated admin CLI and general agent-mode refusal.
5. Export/document the feature, update version metadata, and run all verification.
6. Replace exact typed confirmation phrases with one readable yes/no confirmation,
   plus noninteractive ``--yes`` limited to no-Drive drafts.
7. Add a narrow agent-safe smoke-test path: explicit agent mode, account/course
   allowlists, a no-Classroom-mutation preparation command that produces an account/
   course/client/token-bound approval digest, minimal draft validation, a per-token operation lock,
   all-state paginated duplicate detection, and read-back verification. Keep the
   general coursework mutation API forbidden in agent mode.
8. Run one user-authorized live create/read-back check against the exact pilot course.

## Decisions

| Decision | Rationale | Status |
|---|---|---|
| Dedicated admin CLI | Avoids legacy CLI startup writes and agent exposure | Locked |
| One interactive yes/no | Preserve target awareness without typed account/digest phrases | Locked |
| ``--yes`` only for no-Drive drafts | Enables safe smoke automation without silent publish/share effects | Locked |
| Separate agent-safe minimal-draft path | Makes requested automation explicit without weakening the general agent refusal | Locked |
| Account/course allowlists plus bound approval digest | Binds an agent run to the reviewed principal, canonical course, OAuth client/token source, and frozen payload | Locked |
| Exact-title preflight and read-back | Avoids sequential duplicate smoke drafts and verifies the actual stored shape | Locked |
| Per-token lock around list/create/read-back | Serializes cooperating local runs across Classroom's non-atomic list/create boundary | Locked |
| Explicit `DRAFT` default | Prevents accidental publication and doc-default ambiguity | Locked |
| Inline rubrics only | Preserves deterministic preview; avoids mutable Sheet TOCTOU | Locked |
| Three dedicated OAuth scopes | Create coursework, resolve course, verify account | Locked |
| Narrow Google identity-scope handling | Require all requested grants; allow only additional `openid`/`email` scopes without global relaxation | Locked |
| Hidden remote loopback handoff | Keep callback code/state out of shell history and argv; direct validated localhost request only | Locked |
| JSON token per account | Avoids pickle execution and cross-account token collision | Locked |
| Zero mutation retries | Classroom create has no caller idempotency key | Locked |
| One-shot Classroom HTTP transport | Disable connection/redirect/auth-response replay below google-api-python-client as well as its public retry loop | Locked |
| Strict live field projections | Minimize API data and make missing course/read-back evidence fail closed | Locked |
| Resolve aliases before writes | Prevent canonical response IDs from redirecting or invalidating staged follow-ups | Locked |
| Require durable refresh token | Prevent replacement with an expiring access-only credential | Locked |
| Recheck time before release | Draft/rubric calls can consume a near-term schedule or deadline | Locked |
| POSIX mutation support first | Native Windows secret ACL behavior is unverified | Locked |

## Verification Plan

| Check | Command | Expected result |
|---|---|---|
| Coursework unit tests | `python3 scripts/test_gclass_coursework.py` | Pass offline |
| Auth unit tests | `python3 scripts/test_gclass_coursework_auth.py` | Pass offline |
| Admin CLI tests | `python3 scripts/test_gclass_admin_cli.py` | Pass offline |
| Agent regression | `python3 scripts/test_course_agents.py` | Pass |
| CLI regression | `python3 scripts/test_cli_flags.py` | Pass |
| Syntax/import | `python3 -m compileall -q course_hoanganhduc` | Pass |
| Python 3.9 grammar | Parse package and scripts with `ast` feature version 3.9 | Pass |
| Current Google libraries | Run auth suite in isolated dependency environment | Pass |
| Clean package | Build/install wheel and run installed admin dry-run | Pass |

## Out of Scope

No automated publication, scheduling, Drive sharing, arbitrary agent creation,
credential migration, legacy auth remediation, live deletion, or rollback is performed.
