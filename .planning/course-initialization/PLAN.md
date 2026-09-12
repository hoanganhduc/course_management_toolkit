# Plan: Course Initialization

## Shape

One new module plus one dispatch block. `init_course.py` holds everything; `core.py` gains
an argument group's worth of flags and a twelve-line branch. No existing function changes
behaviour.

The module is layered so that the parts worth testing have no I/O in them:

1. **Pure** — `normalize_course_code`, `course_id_from_classroom_url`, `normalize_sheet_url`,
   `inherit_values`, `redact`. String in, string out.
2. **Filesystem** — `read_config_file`, `find_source_courses`, `scaffold_course_folder`,
   `write_course_config`. Take explicit paths; no reliance on the process working directory
   except where the caller passes it in.
3. **Platform** — `Discovery`, whose four methods are the only places that import
   `gclass_auth`, `c50_ops`, or `canvas_auth`. Tests inject a stand-in object.
4. **Wizard** — `choose_from_list` and the four `resolve_*` steps, each taking `input_fn` and
   `out` so a test drives them with a list of answers and reads the transcript back.
5. **Driver** — `run_init_course`, which sequences the above and returns an `InitResult`.

Imports of `googleapiclient`, `canvasapi`, and `pandas` stay inside functions, matching the
rule `onboard.py` states in its own docstring.

## Steps

1. Write `init_course.py` bottom-up: key taxonomy, pure helpers, guidance text, filesystem
   layer, `Discovery`, the picker, the four platform steps, the driver.
2. Wire `core.py`: flags into the `Configuration` group next to `--clear-config`; dispatch
   immediately after the `--clear-config` block, before the config load, ending in
   `raise SystemExit`.
3. Write `scripts/test_init_course.py` against the seams from step 1.
4. Run the tests, the flag regression, `compileall`, and a 3.9 grammar parse.
5. Exercise the real command end to end on a throwaway code in a temp folder, with
   `--dry-run` first.

## Decisions

- **Write `config.json` directly rather than through `update_config_values`.** That helper
  prints the update dict under `--verbose` and under `DRY_RUN`, and the dict contains
  inherited API keys. The canonical *path* still comes from `get_default_config_path`, so
  the platform layout logic is not duplicated.
- **Rank inheritance sources by inheritable-key count, not by mtime.** The three current
  courses hold three or four keys each and no API keys; the rich configs are the older ones.
  Sorting by recency would recommend the emptiest source.
- **One picker for all four platforms.** `choose_from_list` returns a tagged action rather
  than a value, so the number / manual / guidance / skip contract is written once and every
  platform behaves the same.
- **Copy `credentials.json`, never `token.pickle`.** A token is bound to an authorization and
  a scope set; a stale copy makes the refresh path fail in a way that reads as a credentials
  problem. Copying the client secret file and letting the first API call authorize is the
  path the rest of the toolkit already expects.
- **Guidance is data, not `print` calls.** `GUIDANCE` is a dict of line lists, so a test can
  assert the wizard offered guidance without matching against console output.
- **`--init-force` rather than reusing `--force`.** `--force` already exists globally with a
  different meaning; a second meaning on the same flag would be a trap.
- **A non-interactive run does not list Google Classroom before the course is authorized.**
  Tokens are per course (`config.get_default_token_path` is the course config dir plus
  `token.pickle`, and every existing course folder has its own), so a new course has none.
  The first listing therefore has to run the OAuth flow, and that flow blocks either on a
  browser or on a redirect pasted into `getpass` (`gclass_auth.py:176-180`). Under
  `--init-non-interactive` the id given on the command line is recorded unverified with
  `chưa uỷ quyền Google` as the reason, instead of the process hanging.

## Non-goals

Creating anything on a remote platform. The wizard's answer to "this course does not exist
on Google Classroom" is a numbered list of steps and a link, not a `courses().create` call.
That keeps the feature read-only against every platform and keeps it out of reach of the
agent refusal lists, which is where a create verb would have to be argued about.
