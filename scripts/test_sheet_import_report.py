#!/usr/bin/env python3
"""Tests that the Google Sheet import never drops a row in silence.

A student filled the MAT1206E form and vanished: their student number carried
one digit too many, a bare ``continue`` threw the whole submission away, and the
only trace was a summary line that read like ordinary deduplication.  They were
indistinguishable from a student who had never answered at all, and the class
notice nearly told them so.

Every student below is invented.  The incident was real, but this file is
checked into the repository, so it carries nobody's name, number, address, or
account.

A row whose student number will not parse is still a student, so it is kept with
the bad value moved to ``Invalid Student ID`` and the field left empty; a row
that names nobody is still dropped, but by name.  ``--dry-run`` says what an
import would change and writes nothing.

This suite imports the real database layer, so it needs the course interpreter
rather than a bare python3:

    ~/.course_venv/bin/python scripts/test_sheet_import_report.py -v
"""

from __future__ import annotations

import contextlib
import io
import os
import sys
import tempfile
import unittest
from typing import Any, List, Tuple

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc.data import (  # noqa: E402
    import_students_from_file,
    load_database,
    refine_database,
    save_database,
)
from course_hoanganhduc.models import Student  # noqa: E402

HEADER = (
    "Timestamp,Email Address,Họ và Tên (Full Name),Emai VNU-HUS,GitHub Username,"
    "Mã Sinh Viên (Student ID),Ngày sinh (Date of Birth),Lớp (Class) *,"
    "Lớp học phần (Course Section)"
)

# Nine digits where eight belong: the row that disappeared.
NINE_DIGITS = (
    "9/6/2026 16:00:00,hale@gmail.com,Lê Văn Hà,24007002@hus.edu.vn,levanha,"
    "240070020,2/2/2006,K68A4,MAT1206E 1"
)
# The header the Sheet repeats inside the data when the form is edited.
NOT_A_STUDENT = "9/6/2026 16:10:00,x@gmail.com,Họ Và Tên,,,,,,"
GOOD = (
    "9/5/2026 7:45:03,giangdp@gmail.com,Đỗ Phú Giang,24007001@hus.edu.vn,dophugiang-x1,"
    "24007001,1/1/2006,K68A4,MAT1206E 1"
)
CORRECTED = GOOD.replace("9/5/2026 7:45:03", "9/6/2026 15:44:00").replace(
    "dophugiang-x1", "dophugiang-33"
)


def sheet(rows: List[str]) -> str:
    path = os.path.join(tempfile.mkdtemp(), "form.csv")
    with open(path, "w", encoding="utf-8") as handle:
        handle.write(HEADER + "\n" + "\n".join(rows) + "\n")
    return path


def db_file() -> str:
    return os.path.join(tempfile.mkdtemp(), "students.db")


def imported(rows: List[str], **kwargs: Any) -> Tuple[List[Any], str]:
    """Import a sheet and hand back both the roster and what was said about it."""
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer):
        students = import_students_from_file(sheet(rows), **kwargs)
    return list(students or []), buffer.getvalue()


def by_name(students: List[Any], name: str) -> Any:
    for student in students:
        if str(getattr(student, "Name", "") or "").strip() == name:
            return student
    raise AssertionError(f"{name} is not in the roster")


def field(record: Any, key: str) -> str:
    return str(getattr(record, key, "") or "")


class RejectedStudentNumber(unittest.TestCase):
    """The sheet as it really arrives: one bad number among the good ones."""

    def test_the_registration_survives_the_bad_number(self) -> None:
        students, _ = imported([GOOD, NINE_DIGITS])
        self.assertEqual(len(students), 2)

    def test_the_bad_number_is_parked_and_the_field_left_empty(self) -> None:
        """An eight-digit field is what the roster export and the grade join read,
        so the bad value must not sit in it -- and must not be thrown away either."""
        student = by_name(imported([GOOD, NINE_DIGITS])[0], "Lê Văn Hà")
        self.assertEqual(field(student, "Student ID"), "")
        self.assertEqual(field(student, "Invalid Student ID"), "240070020")

    def test_the_rest_of_the_answer_is_kept(self) -> None:
        student = by_name(imported([GOOD, NINE_DIGITS])[0], "Lê Văn Hà")
        self.assertEqual(field(student, "GitHub Username"), "levanha")

    def test_the_student_is_named_on_stdout(self) -> None:
        _, output = imported([GOOD, NINE_DIGITS])
        self.assertIn("Lê Văn Hà", output)
        self.assertIn("240070020", output)

    def test_a_valid_number_is_left_alone(self) -> None:
        student = by_name(imported([GOOD])[0], "Đỗ Phú Giang")
        self.assertEqual(field(student, "Student ID"), "24007001")
        self.assertEqual(field(student, "Invalid Student ID"), "")


class DroppedRows(unittest.TestCase):
    def test_a_row_that_names_nobody_is_dropped(self) -> None:
        students, _ = imported([GOOD, NOT_A_STUDENT])
        self.assertEqual([field(s, "Name") for s in students], ["Đỗ Phú Giang"])

    def test_the_dropped_row_is_named(self) -> None:
        """The count alone reads like deduplication, which is why nobody noticed."""
        _, output = imported([GOOD, NOT_A_STUDENT])
        self.assertIn("Họ Và Tên", output)

    def test_a_missing_address_does_not_print_as_nan(self) -> None:
        _, output = imported([NOT_A_STUDENT.replace("x@gmail.com", "")])
        self.assertNotIn("nan", output.lower())


class DryRun(unittest.TestCase):
    def setUp(self) -> None:
        self.path = db_file()
        with contextlib.redirect_stdout(io.StringIO()):
            save_database([Student(**{
                "Name": "Đỗ Phú Giang",
                "Email": "giangdp@gmail.com",
                "Student ID": "24007001",
                "GitHub Username": "dophugiang-x1",
                "Timestamp": "9/5/2026 7:45:03",
            })], self.path)

    def test_nothing_is_written(self) -> None:
        imported([CORRECTED], db_path=self.path, dry_run=True)
        with contextlib.redirect_stdout(io.StringIO()):
            stored = load_database(self.path)
        self.assertEqual(field(stored[0], "GitHub Username"), "dophugiang-x1")

    def test_the_update_is_reported_before_it_happens(self) -> None:
        _, output = imported([CORRECTED], db_path=self.path, dry_run=True)
        self.assertIn("Dry run", output)
        self.assertIn("dophugiang-33", output)

    def test_the_real_run_then_makes_the_change(self) -> None:
        imported([CORRECTED], db_path=self.path)
        with contextlib.redirect_stdout(io.StringIO()):
            stored = load_database(self.path)
        self.assertEqual(field(stored[0], "GitHub Username"), "dophugiang-33")


class RefineDatabase(unittest.TestCase):
    """The second silent drop: this one runs on every single save."""

    def test_what_it_removes_is_named(self) -> None:
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            kept = refine_database([
                Student(**{"Name": "Đỗ Phú Giang", "Student ID": "24007001"}),
                Student(**{"Name": "Họ và tên", "Student ID": "24009999"}),
            ])
        self.assertEqual([field(s, "Name") for s in kept], ["Đỗ Phú Giang"])
        output = buffer.getvalue()
        self.assertIn("Họ và tên", output)
        self.assertIn("24009999", output)

    def test_a_clean_roster_says_nothing(self) -> None:
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            refine_database([Student(**{"Name": "Đỗ Phú Giang", "Student ID": "24007001"})])
        self.assertNotIn("Removed", buffer.getvalue())


if __name__ == "__main__":
    unittest.main(verbosity=2)
