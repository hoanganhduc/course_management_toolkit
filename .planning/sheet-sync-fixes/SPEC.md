# Specification: Sheet-to-Database Sync Fixes

## Goal

Make `students.db` describe the three live sources it is supposed to mirror -- the Google
Form sheet, Google Classroom, and the Classroom50 roster -- and make every row that an
import drops or skips say so by name. On 2026-09-06, while drafting the roster-check
announcements for MAT3508 and MAT1206E, all three counts were wrong and were nearly sent
to students:

- Two students who had corrected their form the previous day were still listed under "you
  filled this in wrong". `--add-google-sheet` had been run repeatedly and the corrected
  answer never reached the database.
- One MAT1206E student filled the form and vanished from it. Their student number
  carried one digit too many, a bare `continue` discarded the whole submission, and
  nothing was printed.
- MAT3508 reported 70 students where Google Classroom held 67, and produced three
  "suspected duplicate accounts", two of which were one student whose second row the
  Classroom sync had created.

Every one of those had to be found by hand, by reading the sheet and calling the Classroom
API directly. The fix that shipped that morning (`c7dee7d`) had no effect because it was
applied one layer below where the import actually merges.

## Scope

- In scope: `_dedup_students`' duplicate predicate; removing the three private copies of
  the merge loop; the `Timestamp` column alias; per-row reporting in
  `read_students_from_excel_csv` and `refine_database`; `--dry-run` on `--add-google-sheet`;
  `Google_ID` matching, display-name spelling, and non-interactive behaviour in
  `gclass_sync`; the duplicate-account, invalid-id and GitHub-existence checks in
  `roster_audit`; the population `--list-classroom50-membership` counts from.
- Out of scope: deleting anything. No stale row is removed, no student is unenrolled, no
  roster is written. Everything new here reports, and the teacher acts. Also out of scope:
  `roster-import`, `invite`, grade sync, submission download, and the announcement text
  itself.

## Assumptions

- A student number at this school is exactly 8 digits; `is_valid_student_id` already
  encodes that and is unchanged.
- `Google_ID` is Classroom's `userId`. One value is one Google account and therefore one
  person, and it is the only identifier in the record that cannot be two people. A student
  can hold two Classroom accounts, so two distinct ids may still be one human -- that is a
  finding for the teacher, not something to merge.
- Google Forms writes a submission time in the sheet's own language: `Timestamp` in an
  English form, `Dấu thời gian` in a Vietnamese one. Only the English spelling reached the
  database before, and it did so by accident, through the keep-unknown-columns branch.
- `_MERGE_PROTECTED_PREFIXES` (`Google_`, `Additional `) already keeps the form from
  overwriting what Classroom wrote. `Google_Classroom_Display_Name` sits behind that
  prefix, which is correct: Classroom is the source of truth for it.
- `roster_audit` is stdlib-only and must stay that way. `scripts/test_roster_audit.py`
  asserts in-process that neither `pandas` nor `course_hoanganhduc.data` is ever imported.
- The database is a pickle of a `list` of `Student`, an attribute bag, so fields whose
  names contain spaces (`Student ID`, `Invalid Student ID`, `GitHub Username`) are reached
  through `getattr` and `__dict__`.

## Interfaces

- `course_hoanganhduc.data`: `_dedup_students` (predicate widened), `normalize_columns`
  (two aliases), `read_students_from_excel_csv` and `read_students_from_pdf` (both gain
  `dry_run`), `import_students_from_file` (gains `dry_run`), `refine_database` (output).
- New student field `Invalid Student ID`, written by the import and read by the audit.
- `course_hoanganhduc.gclass_sync`: new module-level `_display_name_of`.
- `course_hoanganhduc.roster_audit`: new `_invalid_id_detail`, `_person_ref`,
  `roster_usernames_from_payload`, `usernames_not_on_roster`; `list_duplicate_accounts`
  gains a keyword-only `people`; `format_membership` gains `not_on_roster` and
  `population`; `CONSEQUENTIAL_CODES` gains `student_id_invalid`.
- `course` flags: `--no-audit-check-github`; `--dry-run` now reaches `--add-google-sheet`.

## Acceptance Criteria

- A second form submission carrying a later `Timestamp` replaces the stored answer no
  matter which import lane it arrives through, and no matter what order the rows sit in
  the file.
- A form whose header says `Dấu thời gian` still produces a `Timestamp`, so "the later one
  wins" has something to compare.
- There is exactly one merge rule in the package. `read_students_from_excel_csv` and
  `read_students_from_pdf` hand their rows to `save_database`, as `--add-csv` already did.
- Two records are one student when they share a `Google_ID`; or a `Student ID`; or an
  email address while at most one of them carries a student number; or a name while
  neither does. Two records that carry *different* student numbers are never merged, even
  at one address.
- A row whose student number will not parse is kept. The original value is stored under
  `Invalid Student ID`, `Student ID` is left empty so the roster export and the grade join
  stay clean, and the student is named on stdout.
- A row that names nobody is still dropped, but by name and address. `refine_database`,
  which runs on every save, likewise names what it removes.
- `--add-google-sheet --dry-run` reports the rows it would change and writes nothing.
- The Classroom sync matches `Google_ID` before any string, and reads the display name
  under both spellings while writing only `Google_Classroom_Display_Name`.
- With no terminal, an ambiguous Classroom record is skipped and reported. It is never
  resolved by guessing, because creating a row is the one outcome a re-run cannot undo.
- After a sync, rows holding a `Google_ID` the course no longer returns, rows holding no
  `Google_ID` at all, and records skipped as ambiguous are each listed. None is removed or
  marked.
- `--list-duplicate-accounts` compares people, not records: two rows of one Google account
  are one person and not a finding; two accounts under one folded name still are. One pair
  found under two signals is one finding.
- A student number rejected at import counts as a form answer, so its owner never appears
  under "never filled the form"; it is reported as `student_id_invalid` with the digit
  count that is wrong, and the empty `Student ID` it left behind is not reported a second
  time as a formatting fault.
- The GitHub existence check runs by default when `gh` is installed; `--no-audit-check-github`
  turns it off, and an absent `gh` skips it rather than failing.
- `--list-classroom50-membership` counts from the live roster when it can read one, and
  separates "in the database but never imported to the roster" from "on the roster and not
  yet invited". Usernames are compared case-insensitively, since the database export
  lower-cases and the live roster keeps GitHub's own capitalisation. If the roster cannot
  be read, the header says the count came from the database.

## Verification

- Four new offline `unittest` scripts: `scripts/test_sheet_import_merge.py`,
  `test_sheet_import_report.py`, `test_gclass_match.py`, `test_roster_audit_dupes.py`.
- `scripts/test_roster_audit.py` and `scripts/test_dedup_students.py` as regressions.
- End to end on copies of both course databases, reconciled against the three live sources
  per `HUONG-DAN.md` §0.

## Risks

- Widening the merge predicate can join two students who are not one. The email key is
  therefore refused whenever both sides carry a student number, and the name key still
  requires that neither does. `test_sheet_import_merge.py` pins both refusals.
- `Google_ID` as a merge key acts on data the Classroom sync wrote. If a wrong id was ever
  stored against a student, this merges two people. Nothing observed suggests it has been,
  and the read-only reconciliation now lists every row whose id the course does not
  return, which is where such a value would surface.
- The three private merge loops differed slightly from each other; deleting them is a
  behaviour change wherever they differed. The email key was carried over into
  `_dedup_students` precisely to keep the sheet lane's behaviour, and `read_students_from_pdf`'s
  copy referenced two undefined names (`email_s`, `email_u`) on a branch that would have
  raised `NameError` had it ever been reached.
- Turning the GitHub check on by default adds one network call per registered username to
  a command that was offline. It is skipped when `gh` is absent, and `unverified` is never
  reported as an error.

## Found and deliberately not fixed

`normalize_columns` recognises a Student ID column only if its values look like 8-digit
numbers. A sheet in which *every* student number is malformed therefore loses the column
entirely, and with it the `Invalid Student ID` flagging this work added. The incident
shape -- one bad number among good ones -- is covered; an all-bad file is not. It is
outside this change and is recorded here rather than repaired.
