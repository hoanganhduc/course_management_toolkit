#!/usr/bin/env python3
"""Tests for the merge that runs on every database save.

A student who fills the registration form a second time to correct a mistake
used to be ignored: the merge only filled fields that were empty, so the
correction never reached the database.  The rule now is that the later of two
form submissions wins, decided by the Timestamp the form records.

Every case below is built from a shape the real MAT3508 database contains: a
form submission, a Google Classroom record that carries no Timestamp, and the
correction one student sent after mistyping the GitHub username field.

Unlike the other suites in this directory, this one imports the real database
layer, so it needs the course interpreter rather than a bare python3:

    ~/.course_venv/bin/python scripts/test_dedup_students.py -v
"""

from __future__ import annotations

import os
import sys
import unittest
from typing import Any, Dict, List

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc.data import (
    _dedup_students,
    _incoming_is_newer,
    _parse_timestamp,
)
from course_hoanganhduc.models import Student

# Two submissions by the same student, an hour apart, in the format the form
# actually writes.
EARLIER = "9/5/2026 7:45:03"
LATER = "9/5/2026 13:26:09"


def student(**fields: Any) -> Student:
    return Student(**fields)


def merge(*records: Student) -> List[Student]:
    return _dedup_students(list(records)).students


def fields_of(record: Student) -> Dict[str, Any]:
    return dict(record.__dict__)


class TestParseTimestamp(unittest.TestCase):
    """The merge order is only as trustworthy as this parser."""

    def test_the_form_format_parses(self):
        parsed = _parse_timestamp(EARLIER)
        self.assertIsNotNone(parsed)
        self.assertEqual(
            (parsed.year, parsed.month, parsed.day, parsed.hour, parsed.minute, parsed.second),
            (2026, 9, 5, 7, 45, 3),
        )

    def test_the_form_reads_the_month_first(self):
        # The sheet is written in the form's US locale, so 9/5 is September 5th
        # and 5/9 is May 9th.  Pinned because the whole ordering rests on it:
        # were these two read the other way round the comparison would invert.
        september = _parse_timestamp("9/5/2026 7:45:03")
        may = _parse_timestamp("5/9/2026 7:45:03")
        self.assertEqual((september.month, september.day), (9, 5))
        self.assertEqual((may.month, may.day), (5, 9))

    def test_the_clock_is_read_as_24_hour(self):
        # 13:26 is afternoon, not an unparsed 1:26.  A form filled in the
        # morning must sort before one filled after lunch on the same day.
        self.assertLess(_parse_timestamp(EARLIER), _parse_timestamp(LATER))

    def test_an_iso_timestamp_parses_too(self):
        parsed = _parse_timestamp("2026-09-05 07:45:03")
        self.assertEqual((parsed.year, parsed.month, parsed.day), (2026, 9, 5))
        self.assertEqual(parsed, _parse_timestamp(EARLIER))

    def test_ordering_holds_across_months(self):
        self.assertLess(_parse_timestamp("8/30/2026 23:59:59"), _parse_timestamp(EARLIER))

    def test_nothing_readable_gives_none(self):
        for value in ("", "   ", None, "khong phai ngay", "0"):
            with self.subTest(value=value):
                self.assertIsNone(_parse_timestamp(value))


class TestIncomingIsNewer(unittest.TestCase):
    """The full truth table: both sides must carry a readable Timestamp."""

    def test_a_later_submission_is_newer(self):
        self.assertTrue(
            _incoming_is_newer(student(Timestamp=EARLIER), student(Timestamp=LATER))
        )

    def test_an_earlier_submission_is_not(self):
        self.assertFalse(
            _incoming_is_newer(student(Timestamp=LATER), student(Timestamp=EARLIER))
        )

    def test_the_same_submission_is_not(self):
        self.assertFalse(
            _incoming_is_newer(student(Timestamp=EARLIER), student(Timestamp=EARLIER))
        )

    def test_a_stored_record_without_a_timestamp_is_never_replaced(self):
        self.assertFalse(_incoming_is_newer(student(Name="A"), student(Timestamp=LATER)))

    def test_an_incoming_record_without_a_timestamp_never_replaces(self):
        self.assertFalse(_incoming_is_newer(student(Timestamp=EARLIER), student(Name="A")))

    def test_an_unreadable_timestamp_decides_nothing(self):
        self.assertFalse(
            _incoming_is_newer(student(Timestamp="hom qua"), student(Timestamp=LATER))
        )
        self.assertFalse(
            _incoming_is_newer(student(Timestamp=EARLIER), student(Timestamp="hom nay"))
        )

    def test_mixing_a_zoned_and_a_naive_timestamp_decides_nothing(self):
        # Comparing the two raises in Python rather than returning an answer;
        # the merge must not crash a save over it.
        zoned = student(Timestamp="2026-09-05 07:45:03+07:00")
        naive = student(Timestamp=LATER)
        self.assertFalse(_incoming_is_newer(zoned, naive))
        self.assertFalse(_incoming_is_newer(naive, zoned))


class TestNewerSubmissionWins(unittest.TestCase):
    def test_a_correction_replaces_the_stored_answer(self):
        stored = student(**{
            "Student ID": "24001201",
            "Name": "Vũ Văn Quân",
            "GitHub Username": "Quan Vu",
            "Timestamp": EARLIER,
        })
        correction = student(**{
            "Student ID": "24001201",
            "Name": "Vũ Văn Quân",
            "GitHub Username": "24001201-creator",
            "Timestamp": LATER,
        })
        merged = merge(stored, correction)
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0].__dict__["GitHub Username"], "24001201-creator")
        self.assertEqual(merged[0].__dict__["Timestamp"], LATER)

    def test_an_older_submission_arriving_second_is_ignored(self):
        stored = student(**{"Student ID": "1", "Email": "moi@hus.edu.vn", "Timestamp": LATER})
        stale = student(**{"Student ID": "1", "Email": "cu@hus.edu.vn", "Timestamp": EARLIER})
        merged = merge(stored, stale)
        self.assertEqual(merged[0].__dict__["Email"], "moi@hus.edu.vn")
        self.assertEqual(merged[0].__dict__["Timestamp"], LATER)

    def test_the_result_does_not_depend_on_the_order(self):
        old = student(**{"Student ID": "1", "Email": "cu@hus.edu.vn", "Timestamp": EARLIER})
        new = student(**{"Student ID": "1", "Email": "moi@hus.edu.vn", "Timestamp": LATER})
        forwards = merge(old, new)[0].__dict__["Email"]
        old = student(**{"Student ID": "1", "Email": "cu@hus.edu.vn", "Timestamp": EARLIER})
        new = student(**{"Student ID": "1", "Email": "moi@hus.edu.vn", "Timestamp": LATER})
        backwards = merge(new, old)[0].__dict__["Email"]
        self.assertEqual(forwards, "moi@hus.edu.vn")
        self.assertEqual(backwards, "moi@hus.edu.vn")

    def test_reimporting_the_same_sheet_changes_nothing(self):
        row = {"Student ID": "1", "Email": "a@hus.edu.vn", "Timestamp": EARLIER}
        result = _dedup_students([student(**row), student(**row)])
        self.assertEqual(len(result.students), 1)
        self.assertEqual(result.updates, [])
        self.assertEqual(fields_of(result.students[0]), row)

    def test_a_third_submission_beats_both_earlier_ones(self):
        newest = "9/6/2026 9:00:00"
        merged = merge(
            student(**{"Student ID": "1", "Class": "K63A3", "Timestamp": EARLIER}),
            student(**{"Student ID": "1", "Class": "K68A4", "Timestamp": newest}),
            student(**{"Student ID": "1", "Class": "K61A6", "Timestamp": LATER}),
        )
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0].__dict__["Class"], "K68A4")
        self.assertEqual(merged[0].__dict__["Timestamp"], newest)


class TestRecordsWithoutATimestamp(unittest.TestCase):
    """The 21 Google Classroom records carry none, and must stay untouched."""

    def test_a_record_without_a_timestamp_only_fills_what_is_empty(self):
        stored = student(**{"Student ID": "1", "Email": "truong@hus.edu.vn"})
        incoming = student(**{"Student ID": "1", "Email": "khac@gmail.com", "Class": "K68A4"})
        merged = merge(stored, incoming)
        self.assertEqual(merged[0].__dict__["Email"], "truong@hus.edu.vn")
        self.assertEqual(merged[0].__dict__["Class"], "K68A4")

    def test_a_form_row_cannot_overwrite_a_classroom_record(self):
        classroom = student(**{
            "Student ID": "1",
            "Name": "Nguyễn Văn A",
            "Google_Classroom_Email": "a@hus.edu.vn",
        })
        form = student(**{
            "Student ID": "1",
            "Name": "Nguyễn Văn A",
            "Google_Classroom_Email": "khac@hus.edu.vn",
            "Timestamp": LATER,
        })
        merged = merge(classroom, form)
        self.assertEqual(merged[0].__dict__["Google_Classroom_Email"], "a@hus.edu.vn")

    def test_google_fields_survive_even_between_two_timestamped_records(self):
        # Both records have been through a merge already, so both carry a
        # Timestamp and the Google identity written by the Classroom sync.
        stored = student(**{
            "Student ID": "1",
            "Google_ID": "111",
            "Google_Classroom_Display_Name": "Nguyen Van A",
            "Timestamp": EARLIER,
        })
        newer = student(**{
            "Student ID": "1",
            "Google_ID": "222",
            "Google_Classroom_Display_Name": "A Nguyen",
            "Timestamp": LATER,
        })
        merged = merge(stored, newer)
        self.assertEqual(merged[0].__dict__["Google_ID"], "111")
        self.assertEqual(merged[0].__dict__["Google_Classroom_Display_Name"], "Nguyen Van A")

    def test_a_conflicting_section_is_still_kept_beside_the_first(self):
        first = student(**{"Student ID": "1", "Section": "MAT3508 2"})
        second = student(**{"Student ID": "1", "Section": "MAT3508 3"})
        merged = merge(first, second)
        self.assertEqual(merged[0].__dict__["Section"], "MAT3508 2")
        self.assertEqual(merged[0].__dict__["Additional Section"], "MAT3508 3")

    def test_an_additional_field_is_never_overwritten(self):
        stored = student(**{
            "Student ID": "1",
            "Additional Section": "MAT3508 2",
            "Timestamp": EARLIER,
        })
        newer = student(**{
            "Student ID": "1",
            "Additional Section": "MAT3508 9",
            "Timestamp": LATER,
        })
        merged = merge(stored, newer)
        self.assertEqual(merged[0].__dict__["Additional Section"], "MAT3508 2")

    def test_the_better_written_name_still_wins_without_a_timestamp(self):
        merged = merge(
            student(**{"Student ID": "1", "Name": "vu van quan"}),
            student(**{"Student ID": "1", "Name": "Vũ Văn Quân"}),
        )
        self.assertEqual(merged[0].__dict__["Name"], "Vũ Văn Quân")


class TestSectionCorrection(unittest.TestCase):
    def test_a_newer_submission_corrects_the_section_instead_of_doubling_it(self):
        # Two registration lists mean two sections; two form submissions mean
        # the student fixed a typo.  Only the Timestamp tells them apart.
        merged = merge(
            student(**{"Student ID": "1", "Section": "MAT3508 2", "Timestamp": EARLIER}),
            student(**{"Student ID": "1", "Section": "MAT3508 3", "Timestamp": LATER}),
        )
        self.assertEqual(merged[0].__dict__["Section"], "MAT3508 3")
        self.assertNotIn("Additional Section", merged[0].__dict__)


class TestStaleGithubId(unittest.TestCase):
    def test_changing_the_username_clears_the_id_of_the_old_account(self):
        stored = student(**{
            "Student ID": "1",
            "GitHub Username": "cu",
            "GitHub ID": "999",
            "Timestamp": EARLIER,
        })
        newer = student(**{"Student ID": "1", "GitHub Username": "moi", "Timestamp": LATER})
        merged = merge(stored, newer)
        self.assertEqual(merged[0].__dict__["GitHub Username"], "moi")
        self.assertEqual(merged[0].__dict__["GitHub ID"], "")

    def test_an_id_supplied_with_the_new_username_is_kept(self):
        stored = student(**{
            "Student ID": "1",
            "GitHub Username": "cu",
            "GitHub ID": "999",
            "Timestamp": EARLIER,
        })
        newer = student(**{
            "Student ID": "1",
            "GitHub Username": "moi",
            "GitHub ID": "900000042",
            "Timestamp": LATER,
        })
        merged = merge(stored, newer)
        self.assertEqual(merged[0].__dict__["GitHub ID"], "900000042")

    def test_an_unchanged_username_leaves_the_id_alone(self):
        stored = student(**{
            "Student ID": "1",
            "GitHub Username": "same",
            "GitHub ID": "999",
            "Email": "cu@hus.edu.vn",
            "Timestamp": EARLIER,
        })
        newer = student(**{
            "Student ID": "1",
            "GitHub Username": "same",
            "Email": "moi@hus.edu.vn",
            "Timestamp": LATER,
        })
        merged = merge(stored, newer)
        self.assertEqual(merged[0].__dict__["GitHub ID"], "999")
        self.assertEqual(merged[0].__dict__["Email"], "moi@hus.edu.vn")


class TestReporting(unittest.TestCase):
    def test_every_replaced_value_is_reported(self):
        result = _dedup_students([
            student(**{
                "Student ID": "24001201",
                "GitHub Username": "Quan Vu",
                "Email": "cu@hus.edu.vn",
                "Timestamp": EARLIER,
            }),
            student(**{
                "Student ID": "24001201",
                "GitHub Username": "24001201-creator",
                "Email": "moi@hus.edu.vn",
                "Timestamp": LATER,
            }),
        ])
        reported = {(u.field, u.old, u.new) for u in result.updates}
        self.assertEqual(
            reported,
            {
                ("GitHub Username", "Quan Vu", "24001201-creator"),
                ("Email", "cu@hus.edu.vn", "moi@hus.edu.vn"),
                ("Timestamp", EARLIER, LATER),
            },
        )
        self.assertEqual({u.student for u in result.updates}, {"24001201"})

    def test_a_merge_that_only_fills_blanks_reports_nothing(self):
        result = _dedup_students([
            student(**{"Student ID": "1", "Email": "a@hus.edu.vn"}),
            student(**{"Student ID": "1", "Class": "K68A4"}),
        ])
        self.assertEqual(result.updates, [])


class TestUnrelatedRecords(unittest.TestCase):
    def test_two_different_students_are_not_merged(self):
        merged = merge(
            student(**{"Student ID": "1", "Name": "Nguyễn Văn A", "Timestamp": EARLIER}),
            student(**{"Student ID": "2", "Name": "Trần Thị B", "Timestamp": LATER}),
        )
        self.assertEqual(len(merged), 2)

    def test_the_real_correction_end_to_end(self):
        # What the database held for 24001201 before the correction, and the row
        # his second submission produced.
        stored = student(**{
            "Student ID": "24001201",
            "Name": "Vũ Văn Quân",
            "Email": "24001201@hus.edu.vn",
            "GitHub Username": "Quan Vu",
            "GitHub ID": "111111",
            "Section": "MAT3508 3",
            "Class": "K61A6",
            "Timestamp": EARLIER,
        })
        correction = student(**{
            "Student ID": "24001201",
            "Name": "Vũ Văn Quân",
            "Email": "24001201@hus.edu.vn",
            "GitHub Username": "24001201-creator",
            "Section": "MAT3508 3",
            "Class": "K61A6",
            "Timestamp": LATER,
        })
        result = _dedup_students([stored, correction])
        self.assertEqual(len(result.students), 1)
        self.assertEqual(
            fields_of(result.students[0]),
            {
                "Student ID": "24001201",
                "Name": "Vũ Văn Quân",
                "Email": "24001201@hus.edu.vn",
                "GitHub Username": "24001201-creator",
                "GitHub ID": "",
                "Section": "MAT3508 3",
                "Class": "K61A6",
                "Timestamp": LATER,
            },
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
