#!/usr/bin/env python3
"""Tests for the one merge rule the Google Sheet import now goes through.

``--add-google-sheet`` used to collapse duplicates with a private copy of the
merge loop, and that copy only ever filled fields that were empty.  A student who
submitted the form a second time to correct an answer therefore never reached the
database, however often the import ran, and the fix that was written for
``_dedup_students`` never saw a single duplicate pair because the copy had
already merged them away.  The copy is gone: the import hands its rows to
``save_database`` and ``_dedup_students`` decides, exactly as ``--add-csv``
already did.

The header below is the real MAT1206E form header.  Every student in this file
is invented: the incident was real, but a test is checked into the repository
and must not carry anyone's name, number, address, or account.  The two
submissions have the shape the real pair had -- one username, then a different
one the next day.

This suite imports the real database layer, so it needs the course interpreter
rather than a bare python3:

    ~/.course_venv/bin/python scripts/test_sheet_import_merge.py -v
"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from typing import Any, List, Optional

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc.data import (  # noqa: E402
    _dedup_students,
    load_database,
    read_students_from_excel_csv,
)
from course_hoanganhduc.models import Student  # noqa: E402
from course_hoanganhduc.roster_audit import list_invalid_info  # noqa: E402

HEADER = (
    "Timestamp,Email Address,Họ và Tên (Full Name),Emai VNU-HUS,GitHub Username,"
    "Mã Sinh Viên (Student ID),Ngày sinh (Date of Birth),Lớp (Class) *,"
    "Lớp học phần (Course Section)"
)
HEADER_VI = HEADER.replace("Timestamp", "Dấu thời gian")

EARLIER = (
    "9/5/2026 7:45:03,giangdp@gmail.com,Đỗ Phú Giang,24007001@hus.edu.vn,dophugiang-x1,"
    "24007001,1/1/2006,K68A4,MAT1206E 1"
)
LATER = (
    "9/6/2026 15:44:00,giangdp@gmail.com,Đỗ Phú Giang,24007001@hus.edu.vn,"
    "dophugiang-33,24007001,1/1/2006,K68A4,MAT1206E 1"
)
# A second invented student, who answered every question but the GitHub one.
UNANSWERED = (
    "9/5/2026 9:12:40,khiempt@gmail.com,Phạm Tuấn Khiêm,24007003@hus.edu.vn,,"
    "24007003,2/2/2006,K68A4,MAT1206E 1"
)


def sheet(rows: List[str], header: str = HEADER) -> str:
    """One form export on disk, in the encoding the Sheet actually downloads as."""
    path = os.path.join(tempfile.mkdtemp(), "form.csv")
    with open(path, "w", encoding="utf-8") as handle:
        handle.write(header + "\n" + "\n".join(rows) + "\n")
    return path


def db_file() -> str:
    return os.path.join(tempfile.mkdtemp(), "students.db")


def field(record: Any, key: str) -> str:
    return str(getattr(record, key, "") or "")


def only(records: List[Any]) -> Any:
    if len(records) != 1:
        raise AssertionError(f"expected one student, got {len(records)}")
    return records[0]


def by_id(records: List[Any], student_id: str) -> Any:
    return only([r for r in records if field(r, "Student ID") == student_id])


def merge(*records: Student) -> List[Student]:
    return _dedup_students(list(records)).students


def make(name: str, email: str, student_id: Optional[str] = None, **fields: Any) -> Student:
    if student_id is not None:
        fields["Student ID"] = student_id
    return Student(Name=name, Email=email, **fields)


class SheetImport(unittest.TestCase):
    """End to end: a CSV in, a merged roster out."""

    def test_a_later_submission_replaces_the_stored_answer(self) -> None:
        students = read_students_from_excel_csv(sheet([EARLIER, LATER]))
        self.assertEqual(field(only(students), "GitHub Username"), "dophugiang-33")

    def test_an_earlier_submission_does_not_replace_it(self) -> None:
        """Order in the file decides nothing; the Timestamp does."""
        students = read_students_from_excel_csv(sheet([LATER, EARLIER]))
        self.assertEqual(field(only(students), "GitHub Username"), "dophugiang-33")

    def test_a_correction_reaches_a_database_that_holds_the_old_answer(self) -> None:
        """The incident itself: two imports, days apart, into the same file."""
        path = db_file()
        read_students_from_excel_csv(sheet([EARLIER]), db_path=path)
        self.assertEqual(field(only(load_database(path)), "GitHub Username"), "dophugiang-x1")

        read_students_from_excel_csv(sheet([LATER]), db_path=path)
        stored = only(load_database(path))
        self.assertEqual(field(stored, "GitHub Username"), "dophugiang-33")
        self.assertEqual(field(stored, "Student ID"), "24007001")

    def test_a_question_left_blank_is_stored_empty(self) -> None:
        """A cell nobody filled in is empty, not an account named 'nan'.

        pandas hands a blank cell over as NaN, and NaN is truthy, so the import
        read an unanswered question as an answer and stored ``str(nan)``.
        """
        students = read_students_from_excel_csv(sheet([EARLIER, UNANSWERED]))
        self.assertEqual(field(by_id(students, "24007003"), "GitHub Username"), "")

    def test_a_question_left_blank_reaches_the_audit_as_missing(self) -> None:
        """Why it matters: a stored 'nan' is a syntactically valid username.

        The audit therefore never told this student their answer was missing,
        and every student who skipped the question shared one account with
        every other, which is what ``github_shared`` is meant to catch.
        """
        students = read_students_from_excel_csv(sheet([EARLIER, UNANSWERED]))
        issues = list_invalid_info(students)
        missing = [i for i in issues if i.code == "github_missing"]
        self.assertEqual([i.student.student_id for i in missing], ["24007003"])
        self.assertEqual([i for i in issues if i.code == "github_shared"], [])

    def test_the_vietnamese_timestamp_header_still_decides(self) -> None:
        """A form in Vietnamese writes 'Dấu thời gian'; without the alias the
        submission time was lost and 'later wins' had nothing to compare."""
        students = read_students_from_excel_csv(sheet([EARLIER, LATER], header=HEADER_VI))
        stored = only(students)
        self.assertEqual(field(stored, "Timestamp"), "9/6/2026 15:44:00")
        self.assertEqual(field(stored, "GitHub Username"), "dophugiang-33")


class MergeKeys(unittest.TestCase):
    """Which two records are one student, decided in one place for every lane."""

    def test_one_google_account_is_one_person(self) -> None:
        """Ngô Việt Minh's shape: one account, two spellings, two addresses."""
        merged = merge(
            make("MINH NGO VIET", "ngovietminh01012000@gmail.com", Google_ID="118"),
            make("Ngô Việt Minh", "24007005@hus.edu.vn", "24007005", Google_ID="118"),
        )
        self.assertEqual(len(merged), 1)
        self.assertEqual(field(merged[0], "Student ID"), "24007005")

    def test_one_address_is_one_person_when_no_number_contradicts(self) -> None:
        merged = merge(
            make("B SHELL", "b@gmail.com"),
            make("Bê", "b@gmail.com", "24007010"),
        )
        self.assertEqual(len(merged), 1)
        self.assertEqual(field(merged[0], "Student ID"), "24007010")

    def test_two_student_numbers_are_two_people_at_one_address(self) -> None:
        """A shared address is not a shared identity, and merging two students
        into one loses a whole registration."""
        merged = merge(
            make("Nguyễn Văn Anh", "nha@gmail.com", "24001111"),
            make("Nguyễn Thị Em", "nha@gmail.com", "24002222"),
        )
        self.assertEqual(len(merged), 2)

    def test_two_google_accounts_are_two_people(self) -> None:
        merged = merge(
            make("VIET MINH NGO", "24007005@hus.edu.vn", Google_ID="118"),
            make("MINH NGO VIET", "ngovietminh01012000@gmail.com", Google_ID="119"),
        )
        self.assertEqual(len(merged), 2)

    def test_a_name_alone_joins_only_records_with_no_number(self) -> None:
        merged = merge(
            make("Phạm Tuấn Khiêm", "khiempt@gmail.com"),
            make("Phạm Tuấn Khiêm", "24007003@hus.edu.vn"),
        )
        self.assertEqual(len(merged), 1)


# One invented student, and the number they typed with a digit too many.
NAME = "Nguyễn Văn A"
NUMBER = "24001111"
ADDRESS = f"{NUMBER}@hus.edu.vn"
MISTYPED = "240011110"


class ParkedStudentNumber(unittest.TestCase):
    """The number a student mistyped, after they have gone back and fixed it.

    An import that cannot accept a number parks it under ``Invalid Student ID``
    and leaves ``Student ID`` empty.  The correction fills that empty field, but
    the merge loop only ever fills or overwrites -- no branch of it clears -- so
    the parked value outlived the mistake, and the audit went on reading it as
    the current answer.

    The numbers below are invented.  This suite describes a shape the roster
    really has, but it is checked into the repository, so it must not carry a
    student's name, number, address, or account.
    """

    def parked(self, **fields: Any) -> Student:
        return make(NAME, ADDRESS, **{"Invalid Student ID": MISTYPED,
                                      "Timestamp": "9/6/2026 16:18:33", **fields})

    def corrected(self, **fields: Any) -> Student:
        return make(NAME, ADDRESS, NUMBER, Timestamp="9/7/2026 16:18:33", **fields)

    def test_the_correction_retires_the_parked_number(self) -> None:
        merged = only(merge(self.parked(), self.corrected()))
        self.assertEqual(field(merged, "Student ID"), NUMBER)
        self.assertEqual(field(merged, "Invalid Student ID"), "")

    def test_the_retirement_is_reported_as_an_update(self) -> None:
        result = _dedup_students([self.parked(), self.corrected()])
        retired = [u for u in result.updates if u.field == "Invalid Student ID"]
        self.assertEqual([(u.old, u.new) for u in retired], [(MISTYPED, "")])

    def test_a_number_still_unfixed_stays_parked(self) -> None:
        merged = only(merge(
            self.parked(),
            self.parked(Timestamp="9/7/2026 16:18:33",
                        **{"GitHub Username": "nguyenvana"}),
        ))
        self.assertEqual(field(merged, "Student ID"), "")
        self.assertEqual(field(merged, "Invalid Student ID"), MISTYPED)


if __name__ == "__main__":
    unittest.main(verbosity=2)
