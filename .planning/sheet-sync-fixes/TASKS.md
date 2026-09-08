# Tasks: Sheet-to-Database Sync Fixes

- [x] Widen `_dedup_students`' duplicate predicate to `Google_ID` -> `Student ID` -> email
  (only when at most one side has a student number) -> name, keeping `_incoming_is_newer`,
  `FieldUpdate` and `_MERGE_PROTECTED_PREFIXES` as they were.
- [x] Delete the private merge loop in `read_students_from_excel_csv` and route the rows
  through `save_database`, matching `roster_csv.py`.
- [x] Delete the private merge loop in `read_students_from_pdf` the same way.
- [x] Add `Timestamp` and `Dấu thời gian` to `normalize_columns`.
- [x] Keep a row whose student number is malformed: park the value in `Invalid Student ID`,
  leave `Student ID` empty, and skip the `<id>@hus.edu.vn` fallback address for it.
- [x] Name every dropped and every flagged row on stdout at import time.
- [x] Name every row `refine_database` removes, with its student number.
- [x] Thread `dry_run` through `import_students_from_file` to both readers and wire
  `--dry-run` into the `--add-google-sheet` dispatch.
- [x] Add `gclass_sync._display_name_of`, reading both spellings, writing one.
- [x] Match Classroom records on `Google_ID` before any string, and keep
  `local_by_google_id` current as rows are matched and created.
- [x] Catch `EOFError` at both prompts: skip and report the ambiguous record, and leave the
  course picker with a message naming `--google-course-id`.
- [x] Add the read-only reconciliation: both counts, rows whose `Google_ID` the course no
  longer returns, rows with no `Google_ID`, and records skipped as ambiguous.
- [x] Group records into people before comparing, in `list_duplicate_accounts`; name each
  person by their form answer; report one pair once even when two signals find it.
- [x] Read `Invalid Student ID` in `has_form_response` and `identity_keys`; add the
  `student_id_invalid` check with `_invalid_id_detail`, and suppress the
  `student_id_format` finding it would otherwise duplicate.
- [x] Default `--audit-check-github` to on when `gh` is on PATH; add
  `--no-audit-check-github`.
- [x] Count `--list-classroom50-membership` from the live roster, with
  `roster_usernames_from_payload` and `usernames_not_on_roster`, and print the fourth group.
- [x] Write `scripts/test_sheet_import_merge.py`, `test_sheet_import_report.py`,
  `test_gclass_match.py` and `test_roster_audit_dupes.py`.
- [x] Fix the `student()` fixture in `scripts/test_roster_audit.py`.
- [x] Run the whole offline suite.
- [x] Documentation and changelog.

## Verification evidence

- New: `scripts/test_sheet_import_merge.py` 9 OK, `test_sheet_import_report.py` 13 OK,
  `test_gclass_match.py` 12 OK, `test_roster_audit_dupes.py` 21 OK -- 55 new tests.
- Regressions: `test_roster_audit.py` 56 OK, `test_dedup_students.py` 32 OK,
  `test_roster_csv.py` 16 OK, `test_onboard.py` 61 OK, `test_c50_admin_cli.py` 97 OK,
  `test_gclass_invite.py` 28 OK, `test_gclass_auth.py` 19 OK,
  `test_gclass_coursework.py` 32 OK; `test_cli_flags.py` parses 239 flags and
  `test_import_internship.py` verifies, both exit 0.
- `test_classroom50.py`, `test_gclass_admin_cli.py`, `test_gclass_coursework_auth.py` and
  `test_course_agents.py` fail. All four were confirmed to fail identically on the
  unmodified tree, under `git stash push -u`, and none of them exercises anything this work
  touches. `test_course_agents.py`'s single failure is a Canvas preflight config test.
- `test_roster_audit.py`'s three failures during development were the fixture, not the
  code: `student()` handed every record one shared hardcoded `Google_ID`, so records the
  tests meant to be two different students were joined the moment `identity_keys` began
  reading that field. Two Classroom accounts never share a `userId` in reality. The
  assertions are unchanged; the fixture now derives the id from the address.
- The import path was exercised on a real MAT1206E form header outside the test suite: the
  nine-column header maps correctly, the rejected row survives with its nine-digit number
  in `Invalid Student ID` and an empty `Student ID`, the repeated-header row is dropped
  *and named* on stdout, and the later of one student's two submissions wins.
- `roster_audit`'s stdlib-only property survives the new roster reader: `models.py` imports
  nothing, `__init__.py` imports only `.version`, and `c50_roster.py` imports only the
  standard library and `.models`, so the lazy `parse_roster_payload` import inside
  `roster_usernames_from_payload` cannot reach pandas or the data layer.
  `test_roster_audit_dupes.py` carries its own `ImportPurity` test.

## Defects found while working, and what was done with each

- **`read_students_from_pdf` could raise `NameError`.** Its copy of the merge loop tested
  `email_s` and `email_u`, neither of which is defined in that function. The branch was
  reachable only when two records shared no student number, which is the ordinary case for
  an OCR'd list. Removed with the copy rather than fixed in place.
- **`refine_database` and the import disagree about a strange name.** The import's
  `is_strange_name` rejects any name containing a digit; `refine_database`'s own copy
  (`data.py:4196`) uses a fixed list of bad names and a few regexes and has no digit rule.
  Both are now loud about what they remove, and the difference is left alone: it is not
  part of this change.
- **`normalize_columns` can lose the Student ID column entirely.** It recognises the column
  only when its values look like 8-digit numbers, so a sheet in which every student number
  is malformed drops the column, and with it everything this work added around
  `Invalid Student ID`. The incident's shape -- one bad number among good ones -- is
  covered and tested. Reported, not fixed.

## Operator gates

This file records the shape of what happened, not who it happened to. Names, addresses,
student numbers and usernames are placeholders throughout; the real per-student lists live
in each course folder's `HUONG-DAN.md`, which is outside this repository.

- [x] **Dry-run gate.** Ran against backed-up copies of both databases.
  `--add-google-sheet --dry-run` named the student whose number had been rejected, reported
  the fields it would update, and wrote nothing. Before this work that count was always
  zero.
- [x] **Reconciliation gate.** After the real import, both databases match the live
  Classroom membership exactly: MAT3508 67 live students / 67 `Google_ID`s in the database
  and MAT1206E 32 / 32, with no id on either side that the other lacks. The 70-vs-67
  discrepancy is gone. The stale username is gone from `--list-classroom50-membership` and
  its replacement sits in the new fourth group, `CHƯA IMPORT LÊN ROSTER`.
  `--list-duplicate-accounts` on MAT3508 went from three groups to two: the one student
  holding two genuine Classroom accounts stayed, the two false pairs are gone, and one
  unlinked form row is now reported (see below). Every one of the 94 GitHub usernames
  across both courses was re-checked with `gh api users/<u>` and all exist.
- [x] **Announcement gate.** Both `thong_bao_ra_soat.txt` files were rewritten from the
  post-import state. MAT1206E is now 32 of 32 filled -- the "chưa điền form" section is
  empty and one student number remains to fix. MAT3508 is 66 students across 67 Classroom
  accounts, 62 filled and 4 not, with the "nghi trùng" section split into the two cases
  below.

## Unchecked

Teacher's commands. Not reachable from here. The account lists they need are in each
course folder's `HUONG-DAN.md`.

- [ ] `roster-import` for both courses: four usernames each, none of them on the roster
  yet. MAT3508's `classroom50_roster.csv` and `classroom50_roster_invite.csv` are both
  stale and must be re-exported first.
- [ ] Posting the two announcements to Google Classroom.
- [ ] Linking one MAT3508 student's two records by hand. Their Classroom account is a
  personal address and their form answer is the school one; the two share no `Google_ID`,
  `Student ID` or email, and the names differ by word order, so neither the import nor the
  Classroom sync will ever join them -- by design, since guessing across folded names is
  what produced the false pairs on 06/09. Copy the `Google_ID` across by hand, or have the
  student refill the form from the address they use for Classroom.
