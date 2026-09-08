# Specification: Google Classroom Assignment Creation

## Goal

Add a Google Classroom assignment administration surface with complete stable REST v1
assignment options, optional inline rubrics, deterministic dry-run output, narrowly
scoped draft automation, and credential handling isolated from the package's legacy
pickle-based flows.

## Scope

- In scope: pure request builders; Drive/link/YouTube materials; draft, published,
  and scheduled coursework; due dates, grading, topics, assignees, grading periods,
  inline rubrics; staged draft-rubric-release mutation; dedicated JSON OAuth tokens;
  account and canonical-course verification; a separate admin CLI; a single readable
  interactive confirmation; noninteractive ``--yes`` for drafts without Drive-sharing
  effects; offline tests; one opt-in live draft smoke test through a separately
  allowlisted agent-safe path.
- Out of scope: local Drive upload/search, Sheet-sourced rubrics, preview APIs,
  add-ons, questions/material-only posts, deletion, automated rollback, unattended
  publication/scheduling/Drive sharing, and general-purpose agent-mode creation. Native
  Windows mutation fails closed until native ACL/reparse tests exist.

## Assumptions

- The target is Google Classroom REST v1 as documented on 2026-08-28.
- Python 3.9 compatibility is required.
- Existing legacy Google Classroom behavior remains compatible and unchanged.
- Real assignment creation always uses an expected primary email. Interactive creation
  uses one normal yes/no prompt; noninteractive creation is limited to safe drafts.

## Interfaces

- `course_hoanganhduc.gclass_coursework`: builders and mutation state machine.
- `course_hoanganhduc.gclass_coursework_auth`: isolated auth paths and JSON tokens.
- `course_hoanganhduc.gclass_admin_cli`: `auth-status`, `authorize`,
  `complete-loopback`, and `create-assignment` commands.
- Console script: `course-gclass-admin`.

## Acceptance Criteria

- A title-only assignment produces a valid explicit draft body with no materials.
- All stable writable assignment fields and writable material variants validate and
  serialize exactly.
- Rubric requests create a draft first and publish/schedule only after rubric success.
- Course aliases are resolved before mutation, and all response/follow-up parent IDs
  must match Google's canonical course and coursework identifiers.
- Time-sensitive fields are revalidated before the first write and again before a
  staged publish/schedule operation.
- Mutation calls use a one-shot Classroom transport: client retries, connection
  retries, redirects, and response-triggered OAuth replays are disabled. Malformed
  successful mutation responses and ambiguous transport/server failures are reported
  as outcome unknown.
- Dry-run performs no credential lookup, filesystem mutation, auth, or network access.
- New auth never reads pickle tokens or trusts OAuth endpoints from input JSON.
- Fresh/replacement OAuth must yield a durable refresh token before token storage is
  changed.
- Remote loopback completion reads callbacks with hidden input, validates the exact
  local boundary, bypasses proxies/redirects, and never places code/state in argv.
- Real creation verifies the authenticated primary email and canonical course ID.
- Course verification requests only ID, name, and state, and requires an explicit
  ``ACTIVE`` state. The agent-safe duplicate/create/read-back calls use strict field
  projections containing only the values needed for collision and shape checks.
- Authorization does not require a redundant local ``AUTH`` phrase; Google browser
  consent and exact account verification remain authoritative.
- Interactive creation renders a control-character-safe summary of the authenticated
  account, canonical course ID/name, title, release mode, materials, rubric, and Drive
  sharing, then accepts one ``y``/``yes`` confirmation.
- ``--yes`` works without a TTY only with a pre-existing coursework token and for
  ``DRAFT`` assignments with no Drive-sharing effect. It never starts fresh OAuth or
  overrides restricted agent mode. Published, scheduled, and Drive-sharing plans
  remain interactive.
- The normal agent-safe entrypoint explicitly refuses creation.
- A separate ``--agent-safe-draft`` test path requires explicit agent mode, account
  and canonical-course allowlists, ``--yes``, a matching approval digest from the
  no-Classroom-mutation ``prepare-agent-safe-draft`` command, an existing token, and
  the minimal no-material/no-rubric draft shape. The approval digest binds the normalized account,
  account fingerprint, canonical course, frozen live operation, OAuth client
  fingerprint, resolved auth source, and token-path fingerprint. Preparation may
  refresh the token and metadata and is not itself authorization. The path cannot
  accept aliases, deadlines, scheduling, points, topics, targeted students,
  publication, or sharing effects.
- Before that path creates anything, it paginates draft coursework with zero retries.
  An exact-title match is read back: an identical draft is reused, while a shape
  mismatch, multiple matches, or a same-title published/deleted item fails closed. A
  per-token operation lock serializes the duplicate scan through read-back. A new
  draft is also read back by ID and compared with the frozen expected shape.

## Verification

- New offline `unittest` scripts for builders/state machine, auth, and admin CLI.
- Existing course-agent, CLI-flag, and Classroom50 regressions.
- `compileall` and a clean-environment Google dependency import/runtime check.
- Opt-in live check creates at most one no-attachment draft in an explicitly named
  and allowlisted course, or reuses an identical existing draft. It retrieves the
  selected coursework by ID and never retries the create call.

## Risks

- An ambiguous initial POST cannot be reconciled without an idempotency key.
- Draft creation with Drive material may change sharing before later rubric failure.
- OAuth tokens remain bearer secrets protected by local filesystem permissions.
- Same-UID callers can bypass application-level agent guards; hostile automation requires
  an OS identity or user-presence boundary.
- A successful live smoke test intentionally leaves the requested draft in the pilot
  course; exact-title duplicate detection prevents normal sequential reruns, but the
  list/create boundary is not atomic and concurrent runs can still race.
- Dry-run output and live prompts can contain student and material identifiers and
  should be treated as PII-bearing terminal output.
- Live permissions, tenant policy, and rubric licensing require opt-in smoke testing.
