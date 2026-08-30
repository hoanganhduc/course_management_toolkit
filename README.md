# Course Management Toolkit

<div align="center">
  <a href="https://www.buymeacoffee.com/hoanganhduc" target="_blank" rel="noopener noreferrer">
    <img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me A Coffee" height="40" style="margin-right: 10px;" />
  </a>
  <a href="https://ko-fi.com/hoanganhduc" target="_blank" rel="noopener noreferrer">
    <img src="https://storage.ko-fi.com/cdn/kofi3.png?v=3" alt="Ko-fi" height="40" />
  </a>
</div>

![Version](https://img.shields.io/github/v/release/hoanganhduc/course_management_toolkit?label=version) ![Pre-release](https://img.shields.io/github/v/tag/hoanganhduc/course_management_toolkit?label=pre-release&sort=semver) ![Python](https://img.shields.io/badge/Python-3.9%2B-blue?logo=python) ![GitHub](https://img.shields.io/badge/GitHub-Repo-black?logo=github) ![Docker](https://img.shields.io/badge/Docker-ready-blue?logo=docker) ![GitHub](https://img.shields.io/badge/GitHub-Repo-black?logo=github) ![Status](https://img.shields.io/badge/status-work--in--progress-yellow) ![License](https://img.shields.io/github/license/hoanganhduc/course_management_toolkit)

Utilities for managing course rosters, grading, OCR extraction, Canvas/Google Classroom workflows, and **Classroom50** (foundation50) roster sync. Includes **agent-safe entrypoints** for AI coding agents (restricted surfaces; destructive LMS ops stay on the interactive `course` CLI). **Work in progress**, mainly designed for **personal use** but open-sourced for others to adapt. Code with the help of [GitHub Copilot](https://github.com/features/copilot/) and [ChatGPT Codex](https://openai.com/).

## Table of contents

- [Course Management Toolkit](#course-management-toolkit)
  - [Table of contents](#table-of-contents)
  - [Install (editable)](#install-editable)
  - [Install into per-user venv](#install-into-per-user-venv)
  - [Run](#run)
- [Common workflows](#common-workflows)
- [Classroom50 (foundation50)](#classroom50-foundation50)
- [Agent-safe entrypoints](#agent-safe-entrypoints)
- [Configuration](#configuration)
- [Weekly automation guide](#weekly-automation-guide)
- [Course calendar builder](#course-calendar-builder)
- [Override grades](#override-grades)
  - [Notes](#notes)
  - [External tools (optional)](#external-tools-optional)
  - [Troubleshooting](#troubleshooting)
  - [Troubleshooting OCR](#troubleshooting-ocr)
  - [Documentation](#documentation)
  - [Samples](#samples)
  - [License](#license)

## Install (editable)

```bash
pip install -e .
```

## Install into per-user venv

Linux/macOS:

```bash
make install
~/.course_venv/bin/course
```

Windows:

```bat
make.bat install
%USERPROFILE%\.course_venv\Scripts\course.exe
```

## Run

```bash
course
```

Interactive menu tips:
- Use arrow keys (or W/S) to move, Enter to select, q to quit.
- You can also type the menu number quickly to jump to an option.
Menu coverage note:
- All menu actions have CLI equivalents; see `docs/cli_reference.rst` for the full flag list.

Menu ↔ CLI examples:

| Menu item | CLI equivalent |
| --- | --- |
| List students by email domain | `course --list-email-domain gmail.com` |
| List students with duplicate names | `course --list-duplicate-names --duplicate-name-field name` |
| List Google Classroom students | `course --list-google-students` |
| List students by submission status | `course --list-submission-status google:turned_in` |
| Import students from Google Sheet | `course --add-google-sheet <URL>` |
| Download Google Classroom submissions | `course --download-google-classroom-submissions --gc-download-coursework-id <ID>` |
| Run weekly automation | `course --run-weekly-automation --weekly-assignment-id <ID>` |
| Classroom50 preflight | `course --classroom50-preflight` |
| Sync Classroom50 roster | `course --sync-classroom50 --classroom50-org ORG --classroom50-classroom SHORT` |

## Common workflows

Update a MAT*.xlsx file with grades from the local database:

```bash
course --update-mat-excel MAT3500-3-Toan-roi-rac-4TC.xlsx
course -ume MAT3500-3-Toan-roi-rac-4TC.xlsx
```

Sync Canvas roster into the local database:

```bash
course --sync-canvas
course -sc
```

Notes:
- Canvas sync now stores submission comments and rubric evaluations per assignment in the database.

## Classroom50 (foundation50)

This toolkit wraps [Classroom50](https://github.com/foundation50/classroom50) instructor workflows (`gh teacher`) for roster list/sync/export. It does **not** reimplement Classroom50; the adapter calls the teacher extension.

Prerequisites:

- GitHub CLI (`gh`) installed and authenticated
- Classroom50 **teacher** extension available (`gh teacher …`)
- Local `students.db` (or path you pass) for sync/export

Human CLI examples:

```bash
# Auth / whoami
course --classroom50-preflight

# List classrooms in an org
course --list-classroom50-classrooms --classroom50-org my-org

# List roster / assignments for a classroom short-name
course --list-classroom50-roster --classroom50-org my-org --classroom50-classroom short-name
course --list-classroom50-assignments --classroom50-org my-org --classroom50-classroom short-name

# Pull roster into local DB (fill-only identity join; GitHub numeric id ≠ Username)
course --sync-classroom50 -sc50 --classroom50-org my-org --classroom50-classroom short-name
course --sync-classroom50 --classroom50-org my-org --classroom50-classroom short-name --classroom50-report report.json

# Export local roster as Classroom50 CSV dialect
course --export-classroom50-roster classroom50_roster.csv

# Download submissions (human CLI only; not available on the agent entrypoint)
course --download-classroom50 --classroom50-org my-org --classroom50-classroom short-name \
  --classroom50-assignment assignment-slug --classroom50-download-dest ./c50_submissions
```

Flags are listed under **Classroom50** in `docs/cli_reference.rst`.

### Operator commands (`course-c50-admin`)

Installing the package also creates `course-c50-admin`, a **human-only** entrypoint for the
teacher-side operations that change remote state. It is separate from `course` for the same
reason `course-gclass-admin` is: mutation never reaches the agent entrypoint. Every verb
refuses in agent mode, requires an interactive terminal unless `--dry-run` is given, and
builds one fixed `gh teacher` command — there is no free-form argv.

```bash
# Print the exact command and the preconditions, without touching the network
course-c50-admin assignment-add --org my-org --classroom short-name \
  --slug final-project --name "Final Examination Mini-Project" \
  --mode group --max-group-size 5 --empty-repo --dry-run

# Register an assignment (interactive confirmation)
course-c50-admin assignment-add --org my-org --classroom short-name \
  --slug week-1 --name "Week 1" --template my-org/week-1-template

# Remove an assignment entry (does not delete student repositories)
course-c50-admin assignment-remove --org my-org --classroom short-name --slug week-1

# Invite users to an organization or a repository
course-c50-admin invite --target my-org --username alice --username bob
course-c50-admin invite --target my-org/week-1-alice --username bob --permission maintain

# Download submissions
course-c50-admin download --org my-org --classroom short-name \
  --assignment week-1 --dest ./c50_submissions
```

Guards worth knowing before you run it:

- **`assignment-add` refuses an existing slug.** `gh teacher assignment add` replaces the
  entry in place, and the pinned CLI cannot set submission mode, so re-running it restores
  the default every-push mode and discards a tagged-commit setting made in the web form.
  Overwriting requires both `--allow-overwrite` and an interactive confirmation; there is
  no `--yes`.
- **`assignment-remove` refuses an absent slug** instead of exiting 0, and its confirmation
  states that student repositories survive and that re-adding the same slug is not a clean
  reset.
- **`download --by-pattern` is permitted only for empty-repository assignments.** The flag
  skips the team lookup, so it collects no `result.json` and writes no `scores.csv`; an
  empty-repository assignment has neither, an autograded one does. The assignment record is
  read first, and the flag is refused if `empty_repo` is not set.
- **`invite` preflights organizations, not repositories.** Repository invitations are
  idempotent; organization invitations are not, so actual membership is read first and
  logins that are already members or hold a pending invitation are skipped rather than
  failing the batch. Reading pending invitations needs the `admin:org` scope.

Exit codes: `0` success, `2` validation or refusal, `3` the `gh` call failed and the remote
outcome is unverified, `4` a batch invite had failures. Student acceptance and student
submission are **not** provided: they run under a student's own credentials.


## Agent-safe entrypoints

For AI agents (and for restricted automation), prefer dedicated modules that force **agent mode** (`COURSE_AGENT_MODE=1`). These surfaces refuse destructive or high-blast-radius ops (unenroll, grade-apply, invite, announcement, submission download, interactive DB modify, etc.). Use the interactive `course` CLI as a human when those are required.

| Module | Role | Allowed (typical) | Refused (examples) |
| --- | --- | --- | --- |
| `python -m course_hoanganhduc.c50_agent` | Classroom50 | preflight, list-*, sync, export | assignment-add, assignment-remove, invite, download |
| `python -m course_hoanganhduc.canvas_agent` | Canvas | preflight, list-assignments/members, search-user, sync | unenroll, grade, invite, announce, download, messages, pages |
| `python -m course_hoanganhduc.gclass_agent` | Google Classroom | preflight, list-courses/students, sync | create-assignment, unenroll, grade, download |
| `python -m course_hoanganhduc.db_agent` | Local students.db | search, details, list-*, export-*, count | modify, restore-db, import-apply, delete |

Examples:

```bash
python -m course_hoanganhduc.c50_agent preflight
python -m course_hoanganhduc.c50_agent list-classrooms --org my-org
python -m course_hoanganhduc.c50_agent sync --org my-org --classroom short-name --db students.db
python -m course_hoanganhduc.canvas_agent list-members --course-id 12345
python -m course_hoanganhduc.gclass_agent list-courses
python -m course_hoanganhduc.db_agent search "Nguyen"
python -m course_hoanganhduc.db_agent export-roster --db students.db
```

**Allowlists (agent mode, fail-closed):** when a course/org id is required, set:

| Environment variable | Used by |
| --- | --- |
| `CLASSROOM50_ORG_ALLOWLIST` | `c50_agent` / Classroom50 ops |
| `CANVAS_COURSE_ALLOWLIST` | `canvas_agent` |
| `GCLASS_COURSE_ALLOWLIST` | `gclass_agent` / narrow admin smoke test |
| `GCLASS_ACCOUNT_ALLOWLIST` | narrow Google Classroom admin smoke test |

Comma-separated ids. Empty allowlist with a required id fails closed.

Set these **in the invocation** (or in a file the spawner reads), not in `~/.bashrc`. A
typical `.bashrc` returns early for non-interactive shells, so a variable exported there is
absent under `env -i`, cron, systemd, CI, or any non-login agent spawn, and the fail-closed
check then rejects every call:

```bash
CLASSROOM50_ORG_ALLOWLIST=my-org python -m course_hoanganhduc.c50_agent list-classrooms --org my-org
```

**ai-agents-skills:** install the `course-management` profile to get skills `classroom50`, `course-canvas`, `course-google-classroom`, and `course-db` that route through these modules. See [ai-agents-skills course management docs](https://github.com/hoanganhduc/ai-agents-skills/blob/main/docs/course-management.md).

Student database utilities:
- Find duplicate names or display names (Name/Google Classroom/Canvas).
- Export duplicate-name reports to TXT/CSV/JSON.

List auto-generated CLI short aliases:

```bash
course --list-cli-aliases
course -lca
```

Google Classroom workflows:

### Create a Google Classroom assignment

Assignment creation has a separate administrator command. The restricted agent
entrypoint refuses it, and it never uses the legacy pickle-token flow. This is an
application guard, not a security boundary against a process running as the same OS
user; use a separate OS identity or user-presence credential broker when hostile
same-user automation is in scope.

`course` remains the toolkit's general installed CLI. Installing the package also
adds `course-gclass-admin`, the dedicated assignment-administration entrypoint;
`python -m course_hoanganhduc.gclass_admin_cli` invokes the same CLI and accepts the
same subcommands and options.

The administrator CLI currently owns **assignment creation only**; it is not a
replacement for every Google Classroom workflow. Course/student listing and roster
sync remain on the existing `course` and restricted `gclass_agent` surfaces, while
grading, unenroll, and submission/Drive download remain human-only `course`
operations. Its isolated account-scoped JSON token, fixed scopes, verified principal,
filesystem checks, and one-shot transport are materially safer for assignment
creation than the legacy pickle-token path, but other capabilities need their own
scopes and behavior. The current split is intentional and supported: keep using
`course-gclass-admin` for assignment creation and the existing commands for their
documented workflows. Their credential stores are deliberately separate and are not
interchangeable.

Preview a minimal assignment without credentials, OAuth, or network access:

```bash
course-gclass-admin create-assignment \
  --course-id d:my-course \
  --spec sample/google_classroom/assignment-minimal.sample.json \
  --dry-run
```

Authorize the exact teacher account once, then create with one readable `y/N`
confirmation. Authorization opens Google's browser OAuth consent: it grants the
external Classroom permissions, rather than confirming a package-local operation.
The human flow no longer asks you to type an `AUTH ...`, `CREATE ...`, `SHARE ...`,
or digest phrase. A revoked token, changed grant, or explicit `--replace-token`
operation can require Google consent again; otherwise later commands reuse the
stored refresh token and refresh an expired access token automatically.

```bash
course-gclass-admin authorize --account teacher@example.edu

course-gclass-admin create-assignment \
  --account teacher@example.edu \
  --course-id d:my-course \
  --spec sample/google_classroom/assignment-minimal.sample.json
```

Trusted local automation with a pre-existing token can skip the prompt for a draft
with no Drive-sharing effects. It still verifies the account, resolves and validates
the active course, and is refused by the restricted agent entrypoint:

```bash
course-gclass-admin create-assignment \
  --account teacher@example.edu \
  --course-id 123456789012 \
  --spec sample/google_classroom/assignment-minimal.sample.json \
  --yes
```

The execution modes have deliberately different boundaries:

| Mode | Credentials and network | Confirmation and allowed result |
| --- | --- | --- |
| `--dry-run` | Neither required; no Google call is made | No prompt; emits only a canonical operation plan and digest |
| Interactive create | Verifies the account and active course; can use the stored token | One readable `y/N` prompt; supports draft, publish, schedule, Drive sharing, and rubrics |
| `--yes` | Requires an existing token, disables browser launch, and still verifies the account and course | No prompt; limited to a draft with no Drive-sharing effects and refused in general agent mode |
| Agent-safe automation | Requires explicit agent mode, exact allowlists, an existing token, a canonical course ID, and a separately prepared bound digest | No terminal prompt; limited to the minimal draft smoke-test shape described below |

Published, scheduled, or Drive-sharing assignments therefore always use the single
interactive confirmation.

For a cooperative, user-authorized agent smoke test, the administrator CLI has a
separate minimal-draft path. It is not a general agent mutation API: the normal
`gclass_agent` command still refuses creation. The path requires explicit agent
mode, exact account/course allowlists, an existing token, a canonical course ID, no
materials/rubric/deadline/points/topic/targeting, and an approval digest from a
separate preparation command. Preparation can refresh credential files but performs
no Classroom mutation:

```bash
COURSE_AGENT_MODE=1 \
GCLASS_ACCOUNT_ALLOWLIST=teacher@example.edu \
GCLASS_COURSE_ALLOWLIST=123456789012 \
course-gclass-admin prepare-agent-safe-draft \
  --account teacher@example.edu \
  --course-id 123456789012 \
  --spec sample/google_classroom/assignment-test-draft.sample.json
```

After the user or trusted harness approves that exact envelope, run
`create-assignment` with the same environment and arguments plus
`--agent-safe-draft --yes --expect-approval-digest DIGEST`. The live
path holds a per-token lock, scans every coursework state with pagination, reuses one
identical developer-associated draft, blocks collisions, creates through the
one-shot transport, and reads the selected draft back. The digest is supplied as an
argument by the approved automation; the CLI does not ask a person to retype it.
Environment allowlists and digests prevent accidental drift but do not protect
against hostile code running as the same OS user.

The dedicated Classroom transport disables response-triggered authentication
replay, HTTP adapter retries, redirects, and `google-api-python-client` retries for
each API call. Token refresh may occur before a Classroom request, but the request
itself is sent at most once. A connection failure after a create POST can still make
the result unknowable; `outcome_unknown` (exit 3) and `partial_create` (exit 4,
including a failed read-back) are terminal receipts. Inspect Classroom before any
manual rerun—never retry either outcome automatically.

On a headless remote host, keep authorization running in the first terminal:

```bash
course-gclass-admin authorize \
  --account teacher@example.edu \
  --no-open-browser
```

After consenting in a browser, its redirect to `127.0.0.1` may not reach the
remote host. Note the new redirect URL's port, open a second terminal on the remote
host, and run:

```bash
course-gclass-admin complete-loopback --port PORT
```

Paste the complete browser redirect URL only at the hidden prompt. The helper
accepts only the matching `http://127.0.0.1:PORT/` callback, connects directly
without a proxy or redirect, and does not put the authorization code in shell
history or process arguments. Do not use the callback from an earlier attempt.

Course aliases (`d:` or `p:`) are resolved with `courses.get` before the first
write. This applies to the public assignment state machine as well as the CLI, so
follow-up rubric and release calls are bound to Google's canonical course ID.

The spec supports draft, published, or scheduled assignments; UTC-normalized due
dates; points; all-student or individual assignment; submission modification mode;
topic and grading-period selection; Drive, HTTPS link, and YouTube materials; and an
optional inline scored or unscored rubric. Omit `materials`, use `null`, or use `[]`
for an assignment without attachments. See `docs/usage.rst` for the complete schema
and credential rules.

The new credentials are deliberately isolated from legacy Google operations:

- The fixed requested scope set is Classroom coursework-students, courses-readonly,
  and the Google OAuth `userinfo.email` scope (primary account email only). Google
  can additionally report the OpenID Connect `openid`/`email` identity scopes
  in the grant response; only those documented additions are accepted. Every
  requested permission must still be present, and any unrelated extra permission
  fails closed. The token file records and validates both requested and actually
  reported grants.
- Linux defaults to
  `$XDG_CONFIG_HOME/course/google-classroom/` (or
  `~/.config/course/google-classroom/`); macOS uses
  `~/Library/Application Support/course/google-classroom/`.
- Tokens are account-scoped JSON files, not `token.pickle`. Custom credential paths
  must be absolute, outside Git worktrees, and a custom OAuth client path must be
  paired with a token path.
- OAuth client/token files must be owned by the current user and mode `0600`; their
  directories must not be writable by another user. Use `chmod 600 FILE` and
  `chmod 700 DIRECTORY` on POSIX systems.
- Native Windows mutation currently fails closed because equivalent ACL and reparse
  point protections have not yet been verified.
- Interactive authorization requests offline consent and stores a new token only
  when Google returns a durable refresh token. A failed replacement leaves the
  existing token unchanged.
- OAuth dependency logging that could contain callback codes or bearer tokens is
  suppressed during the interactive exchange.

Inspect credential presence, safety, scope/client matching, expiry, and refreshability
without refreshing a token, launching OAuth, writing, or calling Google. By default,
the result contains safe names and fingerprints rather than full paths:

```bash
course-gclass-admin auth-status --account teacher@example.edu
```

Add `--show-paths` only when you need to discover the fully resolved credential and
token locations for local troubleshooting; its output can reveal local filesystem
layout:

```bash
course-gclass-admin auth-status \
  --account teacher@example.edu \
  --show-paths
```

Environment overrides are `COURSE_GCLASS_CREDENTIALS` and
`COURSE_GCLASS_COURSEWORK_TOKEN`. The legacy `GOOGLE_CLASSROOM_*` variables and
`token.pickle` are never used by assignment creation.

Verification recorded on 2026-08-29:

- 116 offline regression tests passed across coursework construction/state handling,
  isolated authorization, the administrator CLI, restricted agent surfaces, and
  Classroom50 integration; all 215 general CLI flags also passed parser dry-runs.
- A user-authorized, allowlisted live smoke test created one minimal `DRAFT` with no
  attachment, deadline, points, or rubric, and strict read-back verification passed.
- Live publishing/scheduling, Drive sharing, rubrics, and native Windows credential
  mutation remain outside that smoke-test coverage; native Windows continues to fail
  closed.

Sync Google Classroom roster into the local database:

```bash
course --sync-google-classroom
course -sgc
```

Notes:
- Canvas and Google Classroom score sync normalizes grades to a 10-point scale when max points are available.
- Student ID inference for MAT Excel updates only works with VNU University of Science, Hanoi email format.
- Google APIs required: Classroom, Drive (for submission downloads), Sheets (for Google Sheet imports).
- After enabling new APIs or scopes, delete `token.pickle` or re-run to re-auth.
- If `GOOGLE_CLASSROOM_CC_TOPICS`/`GOOGLE_CLASSROOM_GK_TOPICS`/`GOOGLE_CLASSROOM_CK_TOPICS` are not set, topic names are auto-matched using phrases like "Chuyên cần", "Giữa kỳ/giữa kì", "Cuối kỳ/cuối kì".
- When Canvas and Google Classroom grades conflict, the report includes both sources and notes the conflict; interactive flows prompt which source to use.

Setup quick steps:
- Create/choose a Google Cloud project.
- Enable APIs: Classroom API, Drive API, Google Sheets API.
- Configure OAuth consent screen (external/internal as needed).
- Create OAuth client credentials (Desktop app) and download `credentials.json`.
- Place `credentials.json` and (after first run) `token.pickle` in the config folder or pass paths via CLI.

Grade Google Classroom assignments:

```bash
course --grade-google-classroom
course --grade-google-classroom --gc-coursework-id 1234567890 --gc-grade-score 8
```

Notes:
- Use `--gc-include-graded` to include already graded submissions.
- Use `--gc-apply-all` to grade all listed submissions without selection.

Download latest Google Classroom submissions and run checks:

```bash
course --download-google-classroom-submissions
course --download-google-classroom-submissions --gc-download-coursework-id 1234567890 --gc-download-dest-dir ./gclassroom_submissions
```

Notes:
- Downloads the latest submission attachments per student.
- Similarity/meaningfulness checks run on PDFs; notification drafts are generated on request.

Unenroll Google Classroom students by email domain:

```bash
course --unenroll-google-classroom --gc-unenroll-domain gmail.com
course --unenroll-google-classroom --gc-unenroll-domain gmail.com,outlook.com --gc-unenroll-all
course --unenroll-google-classroom --gc-unenroll-email student1@gmail.com,student2@outlook.com
course --unenroll-google-classroom --gc-unenroll-select
course --unenroll-google-classroom --gc-unenroll-missing-student-id
course --unenroll-google-classroom --gc-unenroll-domain gmail.com --dry-run
```
Note:
- Successful unenroll removes matching students from the local database (by Email or Google_ID).

Grade resubmissions (lists assignments that need regrading, excludes Roll Call Attendance, and prompts per student unless default is enabled). When keeping old grade, the newer submission is assigned the most recent graded score from the submission history:

```bash
course --grade-resubmission
course --grade-resubmission --keep-old-grade
course -grs
course -grs --keep-old-grade
```

Export a roster to CSV:

Course management toolkit for automating student records, grading workflows, PDF/OCR extraction,
Canvas/Google Classroom operations, AI-assisted checks, and weekly reporting.

```bash
course --export-roster
course -ero
```

Note: The default export sort order is **Section > First Name (Vietnamese) > Last Name > Student ID**. Vietnamese characters are handled correctly (e.g., "Â" comes before "B").

Preview an import without writing to the database:

```bash
course --preview-import students.xlsx
course -pi students.xlsx
```

Import students from a Google Sheet URL:

```bash
course --add-google-sheet "https://docs.google.com/spreadsheets/d/FILE_ID/edit#gid=0"
```

Notes:
- Google Sheets API must be enabled for the project tied to your credentials.

Import internship data (Active) from Google Sheet or local file:

```bash
course --import-internships "https://docs.google.com/spreadsheets/d/FILE_ID/edit#gid=0"
course --import-internships sample/internship_data.csv
```

Import internship registration data (Skills/Wishlist) from Google Sheet or local file:

```bash
course --import-registrations "https://docs.google.com/spreadsheets/d/FILE_ID/edit#gid=0"
course --import-registrations sample/internship_registrations_data.csv
```

Import and manage company contacts in `companies.db` from a local file or Google Sheet:

```bash
course --import-companies "Danh sách công ty liên hệ (không xóa).xlsx"
course --import-companies "https://docs.google.com/spreadsheets/d/FILE_ID/edit#gid=0"
course --export-companies companies_export.xlsx
```

Notes:
- `UNIVERSITY_NAME` in `config.json` sets the organization field in VCF exports.
- `MAT*.xlsx` imports ignore score columns (CC, GK, CK, totals); only roster fields are imported.
- Internship and Company imports support both Vietnamese and English column headers.

List students by email domain:

```bash
course --list-email-domain gmail.com,outlook.com
```

List students with duplicate names (or display names):

```bash
course --list-duplicate-names
course --list-duplicate-names --duplicate-name-field google
course --list-duplicate-names --duplicate-name-field canvas --duplicate-name-format csv --duplicate-name-output duplicate_canvas
course --list-duplicate-names --duplicate-name-field "Some Custom Field" --duplicate-name-format json --duplicate-name-output dupes.json
```
Notes:
- `--duplicate-name-field` accepts `name`, `google`, `canvas`, or a custom field name.
- `--duplicate-name-format` supports `txt`, `csv`, or `json` (default: `txt`).

List students missing IDs:

```bash
course --list-missing-ids
course --list-missing-ids google,canvas --missing-ids-format csv --missing-ids-output missing_ids.csv
```

List Google Classroom students:

```bash
course --list-google-students
course --list-google-students --google-course-id 1234567890
```

List students by submission status:

```bash
course --list-submission-status google:turned_in
course --list-submission-status canvas:submitted
course --list-submission-status google:NEW@Quiz 1
course --list-submission-status google:NEW@Quiz 1,Quiz 2
```

Notes:
- Google Classroom values: `NEW`, `CREATED`, `TURNED_IN`, `RETURNED`, `RECLAIMED_BY_STUDENT`.
- Canvas values: `UNSUBMITTED`, `SUBMITTED`, `GRADED`, `PENDING_REVIEW`, `COMPLETE`.
- The output now includes attachment details (file count, upload time, file names, sizes, and types) if data has been synced.

Combine listing filters (AND semantics):

```bash
course --list-email-domain gmail.com --list-missing-ids student
course --list-email-domain gmail.com --list-submission-status google:turned_in
course --list-missing-ids student --list-submission-status google:NEW
course --list-email-domain gmail.com --list-missing-ids student --list-submission-status google:CREATED
course --list-submission-status canvas:UNSUBMITTED --list-missing-ids student
course --list-email-domain gmail.com --list-submission-status google:TURNED_IN@Quiz 1
course --list-email-domain gmail.com --list-submission-status google:TURNED_IN@Quiz 1,Quiz 2
```

Notes:
- `--list-duplicate-names` and Canvas/Google roster listings cannot be combined with other listing flags.
- Export options (like `--missing-ids-output`) are ignored in combined mode; the merged list prints to the console.

Unenroll Canvas students:

```bash
course --unenroll-canvas --canvas-unenroll-domain gmail.com
course --unenroll-canvas --canvas-unenroll-email student1@gmail.com,student2@outlook.com
course --unenroll-canvas --canvas-unenroll-select
course --unenroll-canvas --canvas-unenroll-missing-student-id
course --unenroll-canvas --canvas-unenroll-domain gmail.com --dry-run
```
Note:
- Successful unenroll removes matching students from the local database (by Email or Canvas ID).

Export an anonymized roster:

```bash
course --export-anonymized
course -ean
```

Generate a weekly GitHub Actions workflow template:

```bash
course --generate-weekly-workflow
course -gww
```

Run weekly automation for a closed assignment (downloads, checks, grades, reminders):

```bash
course --run-weekly-automation --weekly-assignment-id 123456 --weekly-teacher-canvas-id 987654
course -rwa --weekly-assignment-id 123456 --weekly-teacher-canvas-id 987654
```

If you omit ``--weekly-assignment-id``, the tool scans ``weekly_reports/`` for
previous runs, lists already-processed assignments, then runs on closed Canvas
assignments that are not yet in the weekly reports.

Run weekly automation locally (no GitHub repo needed). Reports are archived under
`weekly_reports/<timestamp>` with a `students.db.bak` copy:

```bash
course --run-weekly-local --weekly-assignment-id 123456 --weekly-local-root "C:\path\to\course-folder"
course -rwl --weekly-assignment-id 123456 --weekly-local-root "C:\path\to\course-folder"
```

Weekly report folders include evidence and outputs such as:
- `run_report.txt`, `data_validation_report.txt`, `grade_diff.csv`
- `weekly_automation_summary.json`
- `final_evaluations/`, `student_submissions/`
- `flagged_submissions_<assignment-name>_<assignment-id>/`
- `students.db.bak`

Clear stored configuration or credentials:

```bash
course --clear-config
course --clear-credentials
course -ccfg
course -ccred
```

Tip: `--google-credentials-path` and `--google-token-path` copy the files into the default config folder with standard filenames, even if you only set them in a separate command before running `--sync-google-classroom`.

Backup or restore the database/config:

```bash
course --backup-db
course --restore-db
course --backup-config
course --restore-config
course -bd
course -rd
course -bc
course -rc
```

Validate student data and export a report:

```bash
course --validate-data
course -vd
```

Generate per-student final evaluation reports (writes to `final_evaluations/` by default):

```bash
course --generate-final-evaluations
course -gfe
```

Preview updates without writing files:

```bash
course --update-mat-excel MAT3500-3-Toan-roi-rac-4TC.xlsx --dry-run --export-grade-diff
course -ume MAT3500-3-Toan-roi-rac-4TC.xlsx -dr --export-grade-diff
```

Student detail sort order (for `--all-details` and `--export-all-details`):

```bash
course --export-all-details students.txt --student-sort-method first_last
course --export-all-details students.txt --student-sort-method last_first
course --export-all-details students.txt --student-sort-method id
course -E students.txt --student-sort-method first_last
course -E students.txt --student-sort-method last_first
course -E students.txt --student-sort-method id
```

Note: `--export-all-details` now includes detailed attachment information (file names, sizes, types, upload times) from both Canvas and Google Classroom submissions.

You can also set `STUDENT_SORT_METHOD` in `config.json` (first_last, last_first, id).

## Configuration

The tool reads settings from `config.json` stored in a course-specific folder determined by `.course_code`.
On first run, you will be prompted for a course code (e.g., MAT3500) and it will be cached in `.course_code`.
When you load a config file with `--config`, it is copied into the default config folder as `config.json`.

## Weekly automation guide

Local weekly run (no GitHub required):
- Ensure `students.db` is present in the folder you want to use (e.g., a Dropbox folder).
- Ensure your `config.json` is already set (Canvas API URL/key, course ID, OCR keys if used).
- Run `course --run-weekly-local --weekly-assignment-id <ID> --weekly-local-root "<folder>"`.
- Reports and evidence are stored in `weekly_reports/<timestamp>/` under the chosen folder.
  If the assignment ID is omitted, the tool processes closed assignments not yet listed in
  `weekly_reports/`.

GitHub Actions weekly run:
- Store `students.db` in the course repo that hosts the workflow.
- Add required secrets (`CANVAS_API_URL`, `CANVAS_API_KEY`, optional `OCRSPACE_API_KEY`).
- Run `course --generate-weekly-workflow` to create `.github/workflows/weekly-course-tasks.yml`.
- The workflow clones the toolkit repo, runs weekly automation, archives results into
  `weekly_reports/<timestamp>/`, removes the toolkit clone, and commits the updated
  `students.db` plus `weekly_reports/`.

Default config locations:
- Windows: `%APPDATA%\\course\\<course_code>\\config.json`
- macOS: `~/Library/Application Support/course/<course_code>/config.json`
- Linux: `~/.config/course/<course_code>/config.json`

Credential and token files live in the same folder by default:
- `credentials.json` (Google service account)
- `token.pickle` (Google OAuth tokens)

You can override paths via `CREDENTIALS_PATH` and `TOKEN_PATH`.
When you provide Google Classroom credentials/token paths via CLI or the menu, the files are copied into the default folder with the standard names.
To remove stored settings or tokens, use `--clear-config` and `--clear-credentials`.
You can also set `GOOGLE_CLASSROOM_COURSE_ID` in `config.json` to skip the course selection prompt.

Optional settings:
- `LOG_DIR`, `LOG_LEVEL`, `LOG_MAX_BYTES`, `LOG_BACKUP_COUNT` for rotating logs.
- `DB_BACKUP_KEEP`, `CONFIG_BACKUP_KEEP` for backup retention.
- `GRADE_AUDIT_ENABLED`, `GRADE_AUDIT_FIELDS` to control grade audit history stored in the database.
- `COURSE_CODE`, `COURSE_NAME` for calendar titles when not provided via CLI or input files.
- `UNIVERSITY_NAME` (default: empty) for VCF export organization field.
- `WEIGHT_CC`, `WEIGHT_GK`, `WEIGHT_CK` for final total score weighting (must sum to 1.0).
- `GOOGLE_CLASSROOM_GRADE_CATEGORY_METHOD` for topic/category aggregation (average, sum, weighted; default: average).
- `GOOGLE_CLASSROOM_CC_TOPICS`, `GOOGLE_CLASSROOM_GK_TOPICS`, `GOOGLE_CLASSROOM_CK_TOPICS` for CC/GK/CK mapping (comma-separated topic names).
- `GOOGLE_SHEET_URL` for a default Google Sheet import source.
- `GOOGLE_SHEET_LECTURER_TOPICS_URL` for the lecturer mini-project topics sheet.
- `GOOGLE_SHEET_STUDENT_REGISTRATION_URL` for the student mini-project registration sheet.
- `INTERNSHIP_REGISTRATION_SHEET_URL` for the student internship registration sheet.
- `COMPANIES_SHEET_URL` for the company contact data sheet.

## Course calendar builder

Build a course calendar from first-week sessions and export TXT/Markdown/ICS files. If the input file is omitted,
the tool prompts for dates and times interactively.

Example (input file):

```text
course_code: MAT3508
course_name: Discrete Mathematics
weeks: 15
extra_week: yes
holiday: 2026-01-01
holiday: 2026-02-17
session: 2026-01-05 08:00-10:00 | Room 101 | Lecture
session: 2026-01-07 08:00-10:00 | Room 101 | Lecture
```

Notes:
- Course code and course name are required for calendar titles. Provide them in the input file, config (`COURSE_CODE`, `COURSE_NAME`), CLI (`--calendar-course-code`, `--calendar-course-name`), or cache the course code in `.course_code`.
- Fixed Vietnamese holidays (1/1, 4/30, 5/1, 9/2) are automatically excluded.
- Add lunar holidays such as Tet or Hung Vuong via `holiday:` lines or interactive input.
- Add unofficial holidays via `unofficial_holiday:` or `extra_holiday:` lines (comma-separated dates).
- Holidays are auto-fetched via AI using the default provider.
- A make-up week is added only when holidays skip sessions.
- Start time must be earlier than end time.
- Sample input: `sample/calendar/course_calendar_sample_input.txt` (triggers a make-up week).
- Unofficial-holiday sample: `sample/calendar/course_calendar_with_unofficial_holidays.txt`.
- Add unofficial holidays with `--calendar-extra-holidays 2026-03-10,2026-04-05`.

```bash
course --build-course-calendar --calendar-input first_week.txt --calendar-course-code MAT3508 --calendar-course-name "Discrete Mathematics"
course --build-course-calendar
course -bcc --calendar-input first_week.txt --calendar-course-code MAT3508 --calendar-course-name "Discrete Mathematics"
course -bcc
```

Import an existing iCal file into Canvas (requires `CANVAS_LMS_API_URL`, `CANVAS_LMS_API_KEY`, `CANVAS_LMS_COURSE_ID`):

```bash
course --import-canvas-calendar-ics course_calendar.ics --skip-duplicates --dry-run
course -icci course_calendar.ics --skip-duplicates -dr
```
Use `--force` to import without duplicate checks.

Outputs: `course_calendar.txt`, `course_calendar.md`, `course_calendar.ics` in the chosen output directory.
Use `--calendar-output-dir` and `--calendar-output-base` to customize the location and base filename.

## Override grades

Place `override_grades.xlsx` in the working directory (see `sample/overrides/override_grades.xlsx` for the format).
Required columns: `Mã Sinh Viên` or `Họ và Tên`, plus at least one of `CC`/`GK`/`CK` (order does not matter). `STT` and `Lý do` are optional.
Common header aliases are accepted, for example `MSSV`, `Mã SV`, `Họ tên`, `Midterm` (Giữa kỳ), `Final` (Cuối kỳ), `CC` (Chuyên cần), `Reason` (Lý do).
Non-empty CC/GK/CK cells override computed grades, and `Lý do` is appended to the final evaluation output when present.
When using Canvas gradebook CSVs, `Unposted Final Score` is used if `Final Score` is missing or all-zero for CC/GK/CK.
Assignment-group scores are omitted from the final evaluation report when all component scores are 0.
To refine the per-student report with AI, set `REPORT_REFINE_METHOD` to `gemini`, `huggingface`, or `local` in config (requires the corresponding API key for remote providers). The report includes both the default model and the model actually used when AI refinement runs.
Local LLM settings (defaults to Ollama):
- `LOCAL_LLM_COMMAND` (default: `ollama`)
- `LOCAL_LLM_MODEL` (default: `llama3.2:3b`)
- `LOCAL_LLM_ARGS` (optional extra CLI args)
- `LOCAL_LLM_GGUF_DIR` (default: `C:\llm`, scanned recursively for `.gguf` files)
Runtime overrides: `--local-llm-command`, `--local-llm-model`, `--local-llm-args`, `--local-llm-gguf-dir`.
Installing local AI models (examples):
- Ollama: https://ollama.com/ then run `ollama pull llama3.2:3b` and set `LOCAL_LLM_COMMAND=ollama`, `LOCAL_LLM_MODEL=llama3.2:3b`.
- llama.cpp: build `llama-cli` (https://github.com/ggerganov/llama.cpp), set `LOCAL_LLM_COMMAND` to the `llama-cli` path and `LOCAL_LLM_ARGS` to include `-m <path-to-gguf>`.
Use `--refine local` or set `REPORT_REFINE_METHOD=local` to use the local model.
To verify AI connectivity, run `course --test-ai` (or `course -tai`) and check the status for each model. Use `--test-ai-model` (or `-tam`) to test a specific model name, or `--test-ai-gemini-model`/`--test-ai-huggingface-model` when testing `--test-ai all`. For local models, run `course --test-ai local` (or `course -tai local`).
To detect locally installed models (Ollama or llama.cpp compatible), run `course --detect-local-ai` (or `course -dla`).
To list available Gemini models for your API key, run `course --list-ai-models gemini` (or `course -lam gemini`). Hugging Face lists the top public text-generation models (up to 50).
When an AI call is rate-limited, the tool retries and may switch to a different available model with similar capabilities.
Submission quality checks (meaningfulness) can be tuned via config keys: `QUALITY_MIN_CHARS`, `QUALITY_UNIQUE_CHAR_RATIO_MIN`, `QUALITY_REPEAT_CHAR_RATIO_MAX`, `QUALITY_VN_CHAR_RATIO_MIN`, `QUALITY_ALNUM_RATIO_MIN`, `QUALITY_SYMBOL_RATIO_MAX`, `QUALITY_EMPTY_LINE_RATIO_MAX`, `QUALITY_MATH_DENSITY_THRESHOLD`, `QUALITY_LENGTH_RATIO_LOW`, `QUALITY_LENGTH_RATIO_MEDIUM`, `QUALITY_LENGTH_RATIO_HIGH`.
When updating MAT Excel files, use `--export-grade-diff` to save a CSV of old vs new values; database grade changes are tracked in `Grade Audit` when enabled.
A brief per-run summary is appended to `run_report.txt` in the working directory.

## Canvas announcements

Create an announcement from a short message (manual input or a TXT file), optionally refine with AI, preview, and post:

```bash
course --add-canvas-announcement --announcement-title "Week 3" --announcement-message "Please submit by Friday"
course --add-canvas-announcement --announcement-title "Reminder" --announcement-file announcement.txt --refine gemini
course -aa --announcement-title "Week 3" --announcement-message "Please submit by Friday"
course -aa --announcement-title "Reminder" --announcement-file announcement.txt --refine gemini
```

Use `--dry-run` to preview without posting. Omit `--refine` to post the original text without AI.

Sample inputs/outputs:
- `sample/announcements/announcement_input.txt`
- `sample/announcements/announcement_refined_output.txt`
- `sample/announcements/announcement_input_vi.txt`
- `sample/announcements/announcement_refined_output_vi.txt`

## Notes

Some features rely on external system tools (for example, Tesseract OCR and Poppler for PDF processing).
Student databases are resolved from the current working directory (for example, running `course` in a folder will read or write `students.db` there).

## External tools (optional)

PDF extraction and local OCR require system tools. Install them before using `--ocr-service tesseract` or features that convert PDF pages to images.

Official pages (external tools used by this project):
- Tesseract OCR: https://tesseract-ocr.github.io/
- Poppler: https://poppler.freedesktop.org/
- OCR.Space: https://ocr.space/
- PaddleOCR: https://github.com/PaddlePaddle/PaddleOCR

Windows (PowerShell):

```powershell
winget install -e --id UB-Mannheim.TesseractOCR
winget install -e --id oschwartz10612.Poppler
```

If commands are not found, add the install folders to `PATH` (common defaults):
- `C:\Program Files\Tesseract-OCR`
- `C:\Program Files\poppler\Library\bin`

macOS (Homebrew):

```bash
brew install tesseract poppler
```

Linux:

```bash
# Debian/Ubuntu
sudo apt-get update
sudo apt-get install -y tesseract-ocr poppler-utils

# Fedora
sudo dnf install -y tesseract poppler-utils

# Arch
sudo pacman -S tesseract poppler
```

Verify:

```bash
tesseract --version
pdftoppm -h
```

For `--ocr-service ocrspace`, set `OCRSPACE_API_KEY` in your config file.


## Troubleshooting

- If `course` cannot find `students.db`, confirm you are running the command in the intended working directory.
- If OCR commands are missing, recheck your PATH or reinstall Tesseract/Poppler.
- If Canvas/Google Classroom calls fail, verify API keys and course IDs in `config.json`.
- If Classroom50 preflight fails, confirm `gh` auth and that the Classroom50 teacher extension is installed (`gh teacher …`).
- If an agent entrypoint exits with `allowlist_required` / `not_allowlisted`, set the matching `*_ALLOWLIST` env var (see [Agent-safe entrypoints](#agent-safe-entrypoints)).
- Agent entrypoints refusing an operation is intentional; use the full `course` CLI as a human for downloads, unenroll, grading, and DB mutations.

## Troubleshooting OCR

Common errors and fixes:

- `tesseract: command not found` (macOS/Linux) or `'tesseract' is not recognized` (Windows): confirm the install and that the bin folder is on `PATH`.
  - Windows: `where tesseract` and `where pdftoppm` should return paths. If not, add `C:\\Program Files\\Tesseract-OCR` and `C:\\Program Files\\poppler\\Library\\bin` to `PATH`, then reopen the terminal.
  - macOS: `brew --prefix tesseract` and `brew --prefix poppler` should point to installed prefixes; ensure Homebrew is on `PATH`.
  - Linux: `which tesseract` and `which pdftoppm` should resolve. If missing, reinstall with your package manager.
- `pdftoppm` missing: Poppler is not installed or not on `PATH`. Reinstall Poppler and re-open your terminal.
- `TesseractNotFoundError` in Python: the OS command is not visible to the Python process; confirm your IDE/terminal inherits the updated `PATH`.
- Post-OCR AI refinement is disabled; improve scan quality or switch OCR engines if text quality is poor.

## Documentation

Sphinx documentation lives in `docs/`.
- CLI reference (auto-generated from the parser): `docs/cli_reference.rst`.

Build HTML docs:

```bash
pip install -r docs/requirements.txt
cd docs
make html
```

Windows:

```bat
pip install -r docs\requirements.txt
cd docs
make.bat html
```

## Samples

See `sample/index.md` for anonymized input examples organized by category:
- `sample/mat/MAT-examples.xlsx`
- `sample/overrides/override_grades.xlsx`
- `sample/config/config.sample.json`
- `sample/config/credentials.sample.json`

## License

GPL-3.0-only. See `LICENSE`.
