# Specification: Course Initialization

## Goal

Turn the start-of-term setup of a course into one command. Today a new course is created by
hand: someone makes a folder, copies the four wrapper scripts from a sibling course, writes
`.course_code`, creates `%APPDATA%\course\<code>\config.json` in a text editor, copies
`credentials.json` across, and finds the Google Classroom, Canvas, and Classroom50
identifiers by reading URLs out of a browser. `course --init-course` does all of that, and —
this is the point of the feature — it *discovers* the identifiers by listing what the
account can already see, so the operator picks a course by number instead of pasting an
opaque id. When a platform is not set up at all, the wizard asks whether to set it up and
otherwise prints the steps for finding the id by hand. A missing platform never fails the
initialization.

## Scope

- In scope: new module `course_hoanganhduc/init_course.py`; `course` flags `--init-course`,
  `--init-dir`, `--init-name`, `--inherit-from`, `--init-google-id`,
  `--init-canvas-id`, `--init-c50-org`, `--init-c50-classroom`, `--init-sheet-url`,
  `--init-no-scaffold`, `--init-no-db`, `--init-non-interactive`, `--init-force`; the
  `core.main` dispatch that runs before config resolution; a Canvas course listing, which
  did not exist; offline tests in `scripts/test_init_course.py`.
- Out of scope: creating a course on any platform (Google Classroom `courses().create`,
  Canvas, `gh teacher classroom create`). The wizard guides and links; the operator creates.
  Also out of scope: importing students (that is `--add-google-sheet` and `--onboard`),
  editing an existing course's config beyond the keys listed here, any agent-reachable path,
  and the legacy in-folder `config.json` layout used by courses before 2026.

## Assumptions

- Config lives at `get_default_config_path()`'s location, which is
  `%APPDATA%\course\<code>\config.json` on Windows and the XDG/Application Support
  equivalents elsewhere (`config.py:84-123`). The in-folder `config.json` seen in
  `Google Drive/DATA/HUS/<year>/HK<n>_<year>/Courses/*` is the pre-2026 layout and is read
  as a *source* to inherit from, never written as a *target*.
- `.course_code` holds the lower-case code and is what makes a folder a course
  (`config.py:44-56`). `COURSE_CODE` inside `config.json` is the display form.
- Account-wide keys really are account-wide. Checked across the seven config folders under
  `%APPDATA%\course`: `CANVAS_LMS_API_URL`, `CANVAS_LMS_API_KEY`, `OCRSPACE_API_KEY`,
  `OCRSPACE_API_URL`, `HUGGINGFACE_API_KEY`, `GEMINI_DEFAULT_MODEL`, `LOCAL_LLM_MODEL`,
  `LOCAL_LLM_COMMAND`, `LOCAL_LLM_GGUF_DIR` each hold one distinct value wherever they are
  set; `GEMINI_API_KEY` holds two (one rotation). Secrets were compared by length and by
  truncated hash, never printed.
- The three current courses (`mat1206e`, `mat3500`, `mat3508`) carry only three or four keys
  each and no API keys at all, while `mat1204`, `mat3302`, and `mat3397` carry the full set.
  So the inheritance source must be ranked by how many inheritable keys a config actually
  holds, not by how recent it is — the newest config is the poorest one.
- `update_config_values` prints the whole update dict under `--verbose` and under `DRY_RUN`
  (`config.py:217-227`). Inherited values include API keys, so this module writes
  `config.json` itself and routes every value it prints through a redactor.
- `--dry-run` reaches `_apply_config_overrides` only at `core.py:1449`, after the point where
  init must dispatch. Init therefore reads `args.dry_run` directly and performs no write of
  its own when it is set.
- Changing `gclass_auth.SCOPES` invalidates every stored token (`gclass_auth.py:149`). The
  existing scopes already allow `courses().list`, so discovery needs no scope change.
- `GOOGLE_SHEET_URL` is stored as the full URL including the `#gid=` fragment, which selects
  the tab. A pasted URL is kept verbatim; only a bare spreadsheet id is expanded.
- The `/c/<token>` segment of a Google Classroom URL decodes as base64 of the decimal course
  id. This round-tripped on two samples but is not a documented guarantee, so a decoded id is
  offered as a suggestion the operator confirms, never written without confirmation.
- Python 3.9 compatibility is required.

## Interfaces

- `course_hoanganhduc.init_course`: `ACCOUNT_WIDE_KEYS`, `COURSE_SPECIFIC_KEYS`, `GUIDANCE`,
  `SourceCourse`, `PlatformOutcome`, `ScaffoldResult`, `InitResult`, `InitCourseError`,
  `normalize_course_code`, `course_id_from_classroom_url`, `normalize_sheet_url`,
  `read_config_file`, `find_source_courses`, `inherit_values`, `choose_from_list`,
  `Discovery`, `scaffold_course_folder`, `run_init_course`.
- `course` flags as listed under Scope, in the `Configuration` argument group, dispatched in
  `core.main` immediately after the `--clear-config` / `--clear-credentials` block and
  before the config load at `core.py:1396`, terminating with `raise SystemExit(...)` the way
  that block does.

## Acceptance Criteria

- `--init-course` on a fresh folder writes `%APPDATA%\course\<code>\config.json`,
  `.course_code`, an empty `students.db`, and the six wrapper scripts, and reports each file
  it created.
- A second `--init-course` on the same folder refuses to overwrite `config.json` and says to
  pass `--init-force`; files already present in the folder are reported as skipped, not
  rewritten.
- Every platform is offered as a numbered list built from what the account can see. Google
  Classroom courses come from `gclass_auth.list_google_classroom_courses`, Classroom50
  classrooms from `c50_ops.list_classrooms`, Canvas courses from `canvas_auth`'s client.
  Nothing in this module re-implements a listing.
- Every platform menu carries the same four kinds of answer: a number, `m` to enter an id or
  URL by hand, `?` to print the guidance for finding that id, and `0` to skip and set the
  platform up later. `?` prints and re-asks; it is not an answer.
- When a listing fails or returns nothing, the wizard asks whether the platform should be set
  up now. Answering yes prints `GUIDANCE[<platform>]` and then re-offers manual entry;
  answering no records the platform as skipped and initialization continues.
- An id supplied by flag is verified against the same listing when the listing is reachable.
  A verification that cannot run at all is reported as unknown, never as a failure.
- A Google Classroom URL pasted at the `m` prompt is decoded to a numeric id and shown back
  for confirmation before use; a decode that does not yield digits falls through to asking
  for the id.
- Account-wide keys are inherited from the chosen source course. The source picker shows how
  many inheritable keys each candidate holds and offers the richest first.
- `credentials.json` is copied from the source course's config folder when the target has
  none. `token.pickle` is never copied: tokens are per-authorization and stale ones confuse
  the refresh path.
- No secret ever reaches stdout. Every printed value passes through the redactor, which
  matches `secret|token|key|password|api` in the key name.
- `--init-non-interactive` never calls `input`; a field with no flag and no inherited value
  is left unset and reported.
- `--dry-run --init-course` writes nothing — no config, no folder file, no database — and
  prints the same summary it would otherwise produce.

## Verification

- New offline `unittest` script `scripts/test_init_course.py` with an injected `Discovery`,
  a scripted `input_fn`, and a captured `out`, run against `tempfile.mkdtemp` folders. No
  network, no googleapiclient, no canvasapi, no pandas.
- `scripts/test_cli_flags.py` regression for the new flags under `COURSE_PARSE_ONLY=1`.
- `python -m compileall -q course_hoanganhduc` and a Python 3.9 grammar parse.

## Risks

- The base64 URL decode is a hypothesis, not a documented API. It is confirmation-gated, so
  the worst case is one wrong suggestion the operator declines.
- `find_source_courses` reads other courses' configs, which hold API keys. They are only ever
  copied into the new config and counted for the picker; they are never printed.
- Listing Canvas courses is a new API call on a token that also grants writes. It is a plain
  `get_courses()` read, made only when the operator asks for the Canvas step.
