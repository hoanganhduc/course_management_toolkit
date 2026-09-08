#!/usr/bin/env python3
"""Tests for matching a Google Classroom record to the student it belongs to.

The sync used to match on strings alone -- display name, name, address -- and
never on ``Google_ID``, although every Classroom record carries one and it is the
only key that cannot be two people.  A student who joined Classroom with a
personal Gmail and filled the form with the school address matched on nothing, so
the sync wrote a second row; MAT3508 ended with 70 rows for 67 students and three
"suspected duplicate accounts", two of which were this bug reporting itself.

Nothing here removes anything.  A row the course no longer holds is named and
left alone: only the teacher can tell a student who left from one who filled the
form before being enrolled.

The Classroom API is replaced by a fake service, so no network and no credentials
are used, but the module imports the real database layer:

    ~/.course_venv/bin/python scripts/test_gclass_match.py -v
"""

from __future__ import annotations

import contextlib
import io
import os
import sys
import unittest
from typing import Any, Dict, List, Tuple
from unittest import mock

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc import gclass_sync  # noqa: E402
from course_hoanganhduc.models import Student  # noqa: E402

COURSE_ID = "770000000001"


def gc(user_id: str, full_name: str, email: str = "") -> Dict[str, Any]:
    """One record in the shape ``courses().students().list()`` returns."""
    return {
        "userId": user_id,
        "profile": {"name": {"fullName": full_name}, "emailAddress": email},
    }


class _Request:
    def __init__(self, payload: Dict[str, Any]) -> None:
        self._payload = payload

    def execute(self) -> Dict[str, Any]:
        return self._payload


class _Students:
    def __init__(self, rows: List[Dict[str, Any]]) -> None:
        self._rows = rows

    def list(self, **_kwargs: Any) -> _Request:
        return _Request({"students": self._rows})


class _Courses:
    def __init__(self, rows: List[Dict[str, Any]]) -> None:
        self._rows = rows

    def students(self) -> _Students:
        return _Students(self._rows)


class _Service:
    """Only the one call the matching path makes; grades are not fetched here."""

    def __init__(self, rows: List[Dict[str, Any]]) -> None:
        self._rows = rows

    def courses(self) -> _Courses:
        return _Courses(self._rows)


def sync(
    students: List[Student], remote: List[Dict[str, Any]], *, answer: Any = EOFError
) -> Tuple[int, int, str]:
    """Run one sync against a fake course; returns (added, updated, output).

    ``answer`` stands in for the operator at the ambiguity prompt.  The default is
    the unattended case: nobody is there, and the sync must not guess.
    """
    if answer is EOFError:
        prompt = mock.Mock(side_effect=EOFError)
    else:
        prompt = mock.Mock(return_value=answer)
    buffer = io.StringIO()
    with mock.patch.object(gclass_sync, "build", lambda *a, **k: _Service(remote)), \
            mock.patch.object(
                gclass_sync, "_get_google_classroom_credentials", lambda *a, **k: object()
            ), \
            mock.patch("builtins.input", prompt), \
            contextlib.redirect_stdout(buffer):
        added, updated = gclass_sync.sync_students_with_google_classroom(
            students, course_id=COURSE_ID
        )
    return added, updated, buffer.getvalue()


def field(record: Any, key: str) -> str:
    return str(getattr(record, key, "") or "")


class MatchByGoogleId(unittest.TestCase):
    def setUp(self) -> None:
        # The form reply, already carrying the id an earlier sync wrote into it.
        self.local = Student(**{
            "Name": "Ngô Việt Minh",
            "Email": "24007005@hus.edu.vn",
            "Student ID": "24007005",
            "Google_ID": "118",
        })

    def test_one_account_matches_whatever_the_name_and_address_say(self) -> None:
        students = [self.local]
        added, _, _ = sync(students, [gc("118", "MINH NGO VIET", "ngovietminh@gmail.com")])
        self.assertEqual(added, 0)
        self.assertEqual(len(students), 1)

    def test_the_form_answer_is_not_overwritten_by_the_classroom_name(self) -> None:
        students = [self.local]
        sync(students, [gc("118", "MINH NGO VIET", "ngovietminh@gmail.com")])
        self.assertEqual(field(students[0], "Name"), "Ngô Việt Minh")
        self.assertEqual(field(students[0], "Student ID"), "24007005")

    def test_the_classroom_display_name_is_stored_under_one_spelling(self) -> None:
        """It is written as Google_Classroom_Display_Name and was read back under
        another spelling, so every run re-wrote a name it already held."""
        students = [self.local]
        _, updated, _ = sync(students, [gc("118", "MINH NGO VIET", "ngovm@gmail.com")])
        self.assertEqual(updated, 1)
        self.assertEqual(
            field(students[0], "Google_Classroom_Display_Name"), "MINH NGO VIET"
        )
        self.assertEqual(gclass_sync._display_name_of(students[0]), "MINH NGO VIET")

    def test_a_second_run_changes_nothing(self) -> None:
        students = [self.local]
        remote = [gc("118", "MINH NGO VIET", "ngovm@gmail.com")]
        sync(students, remote)
        added, updated, _ = sync(students, remote)
        self.assertEqual((added, updated), (0, 0))
        self.assertEqual(len(students), 1)

    def test_a_genuinely_new_student_is_still_added(self) -> None:
        students = [self.local]
        added, _, _ = sync(
            students,
            [gc("118", "MINH NGO VIET", "ngovm@gmail.com"),
             gc("119", "MINH NGO BA", "24007009@hus.edu.vn")],
        )
        self.assertEqual(added, 1)
        self.assertEqual(len(students), 2)


class Ambiguity(unittest.TestCase):
    """Two candidates and nobody to ask."""

    def rows(self) -> List[Student]:
        return [
            Student(**{"Name": "Nguyễn Văn A", "Email": "a@hus.edu.vn",
                       "Student ID": "24001111"}),
            Student(**{"Name": "Trần Văn B", "Email": "b@gmail.com",
                       "Student ID": "24002222"}),
        ]

    def test_no_terminal_means_skip_not_invent(self) -> None:
        """Creating a row on a guess is the one outcome re-running cannot undo."""
        students = self.rows()
        added, _, _ = sync(students, [gc("118", "Nguyễn Văn A", "b@gmail.com")])
        self.assertEqual(added, 0)
        self.assertEqual(len(students), 2)

    def test_the_skipped_record_is_named(self) -> None:
        students = self.rows()
        _, _, output = sync(students, [gc("118", "Nguyễn Văn A", "b@gmail.com")])
        self.assertIn("Nguyễn Văn A", output)
        self.assertIn("b@gmail.com", output)

    def test_an_operator_who_picks_a_candidate_is_obeyed(self) -> None:
        students = self.rows()
        added, updated, _ = sync(
            students, [gc("118", "Nguyễn Văn A", "b@gmail.com")], answer="1"
        )
        self.assertEqual((added, updated), (0, 1))
        self.assertEqual(field(students[0], "Google_ID"), "118")


class Reconciliation(unittest.TestCase):
    """Read-only: the counts and the lists, and not one row touched."""

    def test_a_row_the_course_no_longer_holds_is_reported(self) -> None:
        students = [
            Student(**{"Name": "Ngô Việt Minh", "Email": "24007005@hus.edu.vn",
                       "Student ID": "24007005", "Google_ID": "118"}),
            Student(**{"Name": "Đã Rời Lớp", "Email": "gone@hus.edu.vn",
                       "Student ID": "24000000", "Google_ID": "999"}),
        ]
        _, _, output = sync(students, [gc("118", "VIET MINH NGO", "ngovm@gmail.com")])
        self.assertIn("Đã Rời Lớp", output)
        self.assertIn("24000000", output)

    def test_it_is_reported_and_not_removed(self) -> None:
        students = [
            Student(**{"Name": "Ngô Việt Minh", "Email": "24007005@hus.edu.vn",
                       "Student ID": "24007005", "Google_ID": "118"}),
            Student(**{"Name": "Đã Rời Lớp", "Email": "gone@hus.edu.vn",
                       "Student ID": "24000000", "Google_ID": "999"}),
        ]
        sync(students, [gc("118", "VIET MINH NGO", "ngovm@gmail.com")])
        self.assertEqual(len(students), 2)
        self.assertEqual(field(students[1], "Name"), "Đã Rời Lớp")

    def test_a_row_with_no_google_id_is_reported(self) -> None:
        """Someone who filled the form but is not on Classroom under any account."""
        students = [Student(**{"Name": "Lê Văn Hà", "Email": "hale@gmail.com",
                               "Student ID": "24007002"})]
        _, _, output = sync(students, [])
        self.assertIn("no Google ID", output)
        self.assertIn("Lê Văn Hà", output)

    def test_both_counts_are_printed(self) -> None:
        students = [Student(**{"Name": "Ngô Việt Minh", "Google_ID": "118"})]
        _, _, output = sync(students, [gc("118", "VIET MINH NGO", "ngovm@gmail.com")])
        self.assertIn("Google Classroom returned 1 student(s)", output)
        self.assertIn("the database holds 1", output)


if __name__ == "__main__":
    unittest.main(verbosity=2)
