# Tasks: Course Initialization

## Build

- [x] T1 `init_course.py`: `ACCOUNT_WIDE_KEYS`, `COURSE_SPECIFIC_KEYS`, `GUIDANCE`, the
      named tuples, `InitCourseError`, and the pure helpers (`normalize_course_code`,
      `course_id_from_classroom_url`, `normalize_sheet_url`, `redact`, `_is_set`).
- [x] T2 Filesystem layer: `read_config_file`, `find_source_courses`, `inherit_values`,
      `write_course_config`, `scaffold_course_folder`, `WRAPPER_SCRIPTS`.
- [x] T3 `Discovery` with `google_courses`, `canvas_courses`, `c50_classrooms`, and the three
      `verify_*` methods; all platform imports inside the methods. Verification folded into
      `_verify_in_list` against the listing instead of separate `verify_*` methods: the
      listing is the evidence, and one call covers both "show me the options" and "is this
      id real".
- [x] T4 `choose_from_list` and the four `resolve_*` steps.
- [x] T5 `run_init_course` and `summarize`.
- [x] T6 `core.py` flags in the `Configuration` group.
- [x] T7 `core.py` dispatch after the `--clear-config` block, before the config load.
- [x] T8 `scripts/test_init_course.py` — 57 tests.

## Gates

Each gate is a claim that must be checked directly, not inferred.

- [x] G1 **The dispatch runs before config resolution.** `course --init-course --course-code
      NEWCODE` in a folder with no `.course_code` must not hit the 60-second course-code
      prompt in `get_default_config_path`. Check by running it in an empty temp folder.
      Checked live in an empty scratch folder for `ZZINIT99`: no prompt, exit 0.
- [x] G2 **No secret is printed.** Run the wizard with `--verbose` against a source course
      that holds `CANVAS_LMS_API_KEY` and `GEMINI_API_KEY`, capture stdout, and assert
      neither value appears in it. Checked live against `mat3397`: 7 secret-looking values
      compared against the captured `--verbose` transcript, none present.
- [x] G3 **`--dry-run` writes nothing.** Snapshot the target folder and the config directory
      before and after; both must be byte-identical, and the config folder must not have been
      created. Checked live; the first attempt failed here (the config folder *was* created)
      and was fixed by adding `config.get_config_base_dir`, which has no side effects.
- [x] G4 **The second run does not clobber.** Run twice; the second must exit non-zero
      without modifying `config.json`, and must name `--init-force`. Checked live: exit 2,
      config sha256 unchanged, message names the flag.
- [x] G5 **Inheritance picks the rich source.** With `mat3508` (4 keys, no secrets) and
      `mat3397` (24 keys) both present, the default offer must be `mat3397`. Checked against
      the real config directory: mat3397 15, mat1204 14, mat3302 14, and mat3500 / mat3508 /
      mat1206e 0 inheritable keys each. Ranking by mtime would have offered mat1206e.
- [x] G6 **A failed listing does not fail the run.** Inject a `Discovery` whose
      `google_courses` raises; the run must reach the summary with Google Classroom recorded
      as unavailable and the other platforms still resolved.
- [x] G7 **`--init-non-interactive` never reads stdin.** Pass an `input_fn` that raises; the
      run must complete. Also checked live with stdin closed and a 60 s timeout, including
      the `--init-google-id` path that would otherwise reach the OAuth flow.
- [x] G8 Flag regression: `COURSE_PARSE_ONLY=1 course --init-course --help` parses, and
      `scripts/test_cli_flags.py` still passes. 262 flags parse; 95 declared and 158
      generated short options, none reassigned.
- [x] G9 `python -m compileall -q course_hoanganhduc` clean, and the new module parses under
      Python 3.9 grammar.

## Fixed alongside (separate from the feature)

- **`CANVAS_FINAL_ASSIGNMENT_ID` never reached the code that reads it.** `mat1204`,
  `mat3302` and `mat3397` all set it, and `load_config`'s own docstring (`config.py:405`)
  listed it as a key the function handles, but it was missing from the `known_keys`
  allowlist. Checked by calling `load_config` on the real `mat3397` config: of 25 keys in
  the file, 23 came back, and the two dropped were this one and `GEMINI_API_KEY2`. Its
  sibling `CANVAS_MIDTERM_ASSIGNMENT_ID` *was* in the list, so this was an omission, not a
  policy. `_apply_config_overrides` only assigns keys present in the returned dict, so
  `data.CANVAS_FINAL_ASSIGNMENT_ID` stayed `""` and three code paths were dead:
  `data.py:2291` (Canvas fallback for a missing CK), `data.py:3100` (ck id for the MAT Excel
  import) and `data.py:9886` (`--sync-multichoice-evaluations final`).

  Fixed by adding the key to `known_keys`, guarded by `scripts/test_config_known_keys.py`
  (6 tests). An audit of the whole allowlist says this was the only case of its kind: of the
  59 uppercase globals in `settings.py`, exactly two were absent from `known_keys` — this
  one and `DEFAULT_COURSE_CODE`, which is the fallback course code and not meant to be
  overridable.

  Blast radius, checked rather than assumed. `data.py:2291` only consults the id when CK is
  already `None`, and reads it as the attribute `Assignment: <id>`; a student record without
  that attribute yields `None` and falls through to the previous path, so the change can add
  a CK but never overwrite one. `data.py:3100` and `data.py:9886` both also need
  `CANVAS_LMS_COURSE_ID`, which is `""` in all three of those configs. All three configs
  carry a midterm id and a final id, and the midterm half has been loading all along, so
  the fix puts the final id on the footing the midterm id already had.

- **`CANVAS_CONFIG_PATH` crashed instead of refusing on Windows.**
  `canvas_agent._configure_canvas` guards the file with POSIX checks — `os.getuid()` for
  ownership and a mode with no group or other bits. `os.getuid` does not exist on Windows,
  so any run with `CANVAS_CONFIG_PATH` set died at that call and printed
  `canvas_agent error: module 'os' has no attribute 'getuid'`. The mode half could not have
  passed either: `os.chmod(path, 0o600)` on Windows leaves the mode reading `0o666`.
  Fixed by refusing that env var explicitly on `os.name == "nt"`, mirroring
  `gclass_coursework_auth._require_posix_security`, and naming the alternative in the
  message. The uid, mode, nlink, size, `O_NOFOLLOW` and TOCTOU re-check are untouched on
  POSIX. Test first: `test_canvas_private_config_refuses_windows_intelligibly` asserted
  exit 1, `Windows` in the transcript and no `getuid`; it failed on the missing `Windows`
  before the change. Checked live afterwards — the refusal prints, and the env-var route it
  points to returns `{"ok": true, ...}`.

  `test_canvas_config_rejects_public_mode_and_symlink` was passing on Windows for the wrong
  reason (the `AttributeError`, not the mode check), so it is now POSIX-only alongside
  `test_canvas_preflight_loads_explicit_private_config`.

- **Nine tests asserted POSIX-only behaviour without saying so.** Marked
  `@unittest.skipIf(os.name == "nt", ...)`, matching the convention already in both files.
  `test_gclass_coursework_auth.py`: three `TestPathResolution` tests feed paths like
  `/secrets/client.json`; `Path.is_absolute()` is host-flavoured, so on Windows those are
  relative and `resolve_coursework_auth_paths` rejects them before the assertion under test.
  `test_gclass_admin_cli.py`: six `--agent-safe-draft` tests reach `TokenFileLock` and the
  secure token store, which `_require_posix_security` refuses on native Windows by design;
  each failure message was checked individually and all six carried that same refusal. No
  product code changed for these, and the control was not weakened.

## Follow-ups (not this change)
- **POSIX credential security has no Windows equivalent.** `_require_posix_security` and the
  Canvas config loader both assume uid ownership and permission bits. Giving them real
  Windows behaviour means ACL-based checks (owner SID, no inherited grants), which is a
  separate design decision, not a test fix. Until then, two narrow paths refuse on native
  Windows: `course-gclass-admin authorize` and `create-assignment`, which write the
  account-scoped coursework token store and take its lock, and the optional
  `CANVAS_CONFIG_PATH`. `auth-status` still reports, and nothing else is affected — roster
  sync and course listing authorize through `gclass_auth.py` and the per-course
  `token.pickle`, which has no POSIX gate.
- **`GEMINI_API_KEY2` is a dead key, not the same gap.** Also dropped by `load_config`, but
  adding it to `known_keys` would change nothing: it has no global in `settings.py`, and
  `_apply_config_overrides` only sets attributes a module already has. A repo-wide search
  (including dynamically built key names) finds no reader at all. It is therefore left out
  of `ACCOUNT_WIDE_KEYS`, so new courses do not receive a copy of a live API key that
  nothing uses.
- The per-course `course.sh` wrappers already in `Courses/*` point at
  `~/.course_venv/bin/course`, which does not exist on Windows. The wrappers this feature
  generates probe both layouts; the existing ones are left alone.
