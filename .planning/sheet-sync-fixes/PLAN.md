# Task Plan: Sheet-to-Database Sync Fixes

## Context

Three counts in the 2026-09-06 roster-check announcements were wrong, and each traced to a
different layer of the import path. The morning's fix (`c7dee7d`) had corrected the merge
that `save_database` runs, and its own commit message recorded the symptom without naming
the cause: *"On the live 68-record database the new branch fires zero times."* It fires
zero times because `read_students_from_excel_csv` carries a private copy of the merge loop
and collapses every duplicate before `save_database` is ever called. There are three such
copies in `data.py`, they disagree with each other about what a duplicate is, and one of
them references undefined names.

`--add-csv` (`roster_csv.py`) already does all of this correctly and says so in its own
docstring: *"no second merge rule is invented here."* The plan is to make the other lanes
match it.

## Steps

1. **One merge path.** Widen `_dedup_students`' predicate to `Google_ID` -> `Student ID`
   -> email -> name, then delete the private copies in `read_students_from_excel_csv` and
   `read_students_from_pdf` and route both through `save_database`. Add the `Timestamp` /
   `Dấu thời gian` aliases.
2. **No silent drops.** Replace the two bare `continue`s: keep a row whose student number
   is malformed, moving the value to `Invalid Student ID`; drop a row that names nobody,
   by name. Name what `refine_database` removes. Thread `dry_run` from `--add-google-sheet`
   down to the import.
3. **Classroom sync.** Match on `Google_ID` first; read the display name under both
   spellings and write one; skip and report an ambiguous record when there is no terminal;
   add a read-only reconciliation that lists stale and unmatched rows without touching them.
4. **Truthful reports.** Group records into people before looking for duplicate accounts;
   read `Invalid Student ID` as both a form answer and a finding; default the GitHub
   existence check to on; count `--list-classroom50-membership` from the live roster and
   split the fourth group out of "not invited".
5. Write the four test scripts and run the whole offline suite.
6. Documentation and changelog.

## Decisions

| Decision | Rationale | Status |
|---|---|---|
| Delete the private merge loops rather than fix them | Three copies disagreeing is the root cause; fixing each repeats it | Locked |
| Keep the email key in `_dedup_students` | The sheet lane's copy matched on email; dropping it would be a regression | Locked |
| Refuse the email key when both sides have a student number | A shared address is a family or a typo, not one student; merging loses a registration | Locked |
| `Google_ID` as the first merge key | One Google account is one person; it is the only identifier that cannot be two | Locked |
| Keep a row with a bad student number, do not drop it | The student did register; the number is one field of the answer | Locked |
| `Student ID` left empty, value parked in `Invalid Student ID` | The roster export and the grade join read that field and must not get rubbish | Locked |
| Report stale rows, never delete them | Only the teacher can tell a student who left from one who registered early | Locked |
| Skip, never guess, when ambiguity meets no terminal | Creating a row is the one outcome a re-run cannot undo | Locked |
| GitHub check on by default, `--no-audit-check-github` to disable | A username belonging to nobody costs the student their repository and nothing else sees it | Locked |
| Four membership groups, not three | "Never imported" and "imported but not invited" need different commands | Locked |
| `roster_audit` stays stdlib-only | Pinned by an in-process import test; the lazy roster import respects it | Locked |

## Verification Plan

| Check | Command | Expected result |
|---|---|---|
| Merge rule | `python scripts/test_sheet_import_merge.py` | Pass offline |
| Import reporting | `python scripts/test_sheet_import_report.py` | Pass offline |
| Classroom matching | `python scripts/test_gclass_match.py` | Pass, fake service, no network |
| Audit reports | `python scripts/test_roster_audit_dupes.py` | Pass, stdlib only |
| Merge regression | `python scripts/test_dedup_students.py` | Pass |
| Audit regression | `python scripts/test_roster_audit.py` | Pass |
| Whole suite | every `scripts/test_*.py` | No new failure |
| End to end | `--add-google-sheet --dry-run` on database copies | Names the student with the rejected number; non-zero update count |

## Out of Scope

No deletion of any kind: no stale row removed, no student unenrolled, no roster written, no
invitation sent. `roster-import`, `invite`, grade sync and submission download are the
teacher's commands and are untouched. The two drafted `thong_bao_ra_soat.txt` files are not
edited here; they describe the database as it was before this work and must be re-read
after the re-import.
