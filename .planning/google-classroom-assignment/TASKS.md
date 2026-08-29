# Tasks: Google Classroom Assignment Creation

- [x] Confirm scope, official REST v1 surface, and security assumptions.
- [x] Inspect current Google modules, CLI startup, credential resolver, and tests.
- [x] Add failing offline tests for coursework, auth, CLI, and agent refusal.
- [x] Implement coursework/material/rubric builders and mutation state machine.
- [x] Implement isolated credential resolver, JSON token storage, and account binding.
- [x] Implement the dedicated admin CLI and package entrypoint.
- [x] Update facade, documentation, changelog, ignore rules, and version metadata.
- [x] Run focused and regression verification.
- [x] Reproduce and fix Google's additional OAuth identity-scope response without relaxing
  unrelated grant validation.
- [x] Add a hidden-input, proxy-free remote loopback helper and suppress secret-bearing
  dependency logs during authorization.
- [x] Record skipped live/Windows checks and residual risks.
- [x] Replace exact ``AUTH``/``CREATE``/``SHARE`` phrases with one readable yes/no.
- [x] Add noninteractive ``--yes`` for no-Drive drafts while retaining agent refusal.
- [x] Update documentation and regression tests for the simplified human flow.
- [x] Run the complete offline verification matrix after the UX revision.
- [x] Add and test a separately allowlisted, exact-digest agent-safe minimal-draft path.
- [x] Re-run the complete offline verification matrix after the agent-safe revision.
- [x] Create and read-back verify one draft in the explicitly allowlisted pilot course.

## Verification evidence

- Coursework, auth, admin CLI, agent, and Classroom50 suites pass: 116
  tests plus all 215 CLI parse cases in the current Google dependency environment.
- All four Google-dependency auth cases pass with the installed Google and OAuth
  client libraries, including the synthetic reproduction of Google's additional
  identity scopes and malformed-scope fail-closed checks.
- Clean wheel installation, installed console-script dry run, ``compileall``, Python
  3.9 grammar parsing, TOML parsing, and ``git diff --check`` pass.
- Code, test, and adversarial security reviews are READY with no remaining concrete
  high-, medium-, or low-severity finding.
- The bound live smoke test created one ``DRAFT`` in the explicitly allowlisted
  pilot course and strict read-back verification passed. Live identifiers are
  intentionally omitted.

## Residual and unchecked environments

- Live OAuth consent, remote loopback completion, token storage/refresh, UserInfo
  account verification, active-course lookup, minimal draft creation, tenant write
  policy, and teacher mutation permission succeeded for the configured account/course.
  Rubric licensing, publication/scheduling, and real Drive sharing remain unchecked.
- Native Windows authorization/mutation remains fail-closed pending ACL/reparse-point
  tests; no native Python 3.9 interpreter was available (59 files passed Python 3.9
  grammar parsing only). Sphinx was unavailable, so the rendered documentation build
  was not run.
- Same-UID code can bypass application guards; ambiguous initial create has no
  idempotency key; prompts and previews may contain PII-bearing opaque IDs.
