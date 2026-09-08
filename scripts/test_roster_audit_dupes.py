#!/usr/bin/env python3
"""Tests for the three audit reports that were describing the wrong world.

MAT3508 produced three "suspected duplicate accounts" on 06/09/2026 and two of
them were the same student twice: the sync had written a second row, and the
duplicate check compared *records* rather than people, so it reported the
toolkit's own bug as a student holding two accounts.  It now compares people,
which is what :func:`group_by_person` was already computing one line above.

Lê Văn Hà's registration was rejected over one extra digit and the number is
now kept under ``Invalid Student ID``; the audit has to read it, or he reappears
as somebody who never answered the form.

``--list-classroom50-membership`` counted from the database, so a name that had
never been pushed to the roster came out under "CHƯA ĐƯỢC MỜI" beside students
whose invitation really was missing.  Those are two different jobs.

Like the other roster-audit suite this one is stdlib-only and runs anywhere:

    python3 scripts/test_roster_audit_dupes.py -v
"""

from __future__ import annotations

import os
import sys
import unittest
from typing import Any, Dict, List

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc.roster_audit import (  # noqa: E402
    MEMBER_ABSENT,
    MEMBER_ACTIVE,
    MEMBER_INVITED,
    MembershipCheck,
    SEVERITY_ERROR,
    audit_roster,
    format_announcement,
    format_membership,
    group_by_person,
    has_form_response,
    list_duplicate_accounts,
    list_invalid_info,
    list_missing_form,
    roster_usernames_from_payload,
    usernames_not_on_roster,
)


def student(**fields: Any) -> Dict[str, Any]:
    """A roster record, with a Google account of its own unless one is given."""
    base = {
        "Name": "Nguyễn Văn A",
        "Email": "24001111@hus.edu.vn",
        "Student ID": "24001111",
        "GitHub Username": "nguyenvana",
        "Section": "MAT1206E 1",
        "Class": "K68A4",
    }
    base.update(fields)
    if "Google_ID" not in fields:
        base["Google_ID"] = "google-" + str(base.get("Email") or base["Name"])
    return {k: v for k, v in base.items() if v is not None}


def codes(issues: List[Any]) -> List[str]:
    return [issue.code for issue in issues]


# The real MAT3508 shape: one Classroom account, two rows, nothing in common but
# the id the sync wrote.
SHELL = student(
    Name="NGOC TRAN TRONG", Email="trantrongngoc0101@gmail.com", Google_ID="118",
    **{"Student ID": None, "GitHub Username": None},
)
FORM = student(
    Name="Trần Trọng Ngọc", Email="24007006@hus.edu.vn", Google_ID="118",
    **{"Student ID": "24007006", "GitHub Username": "trantrongngoc"},
)


class DuplicatesAreBetweenPeople(unittest.TestCase):
    def test_two_rows_of_one_google_account_are_one_person(self) -> None:
        self.assertEqual(len(group_by_person([SHELL, FORM])), 1)

    def test_and_are_not_reported_as_two_accounts(self) -> None:
        self.assertEqual(list_duplicate_accounts([SHELL, FORM]), [])

    def test_the_pair_is_counted_once_in_the_class_size(self) -> None:
        report = audit_roster([SHELL, FORM])
        self.assertEqual(report.total, 1)
        self.assertEqual(report.with_form, 1)
        self.assertEqual(report.duplicate_accounts, [])

    def test_two_real_accounts_under_one_name_are_still_reported(self) -> None:
        """Ngô Việt Minh: two Classroom accounts, and the one finding that was real."""
        rows = [
            student(Name="VIET MINH NGO", Email="24007005@hus.edu.vn", Google_ID="200",
                    **{"Student ID": None, "GitHub Username": None}),
            student(Name="MINH NGO VIET", Email="ngovietminh01012000@gmail.com",
                    Google_ID="201", **{"Student ID": None, "GitHub Username": None}),
        ]
        groups = list_duplicate_accounts(rows)
        self.assertEqual(len(groups), 1)
        self.assertEqual(len(groups[0].members), 2)

    def test_a_person_is_named_by_their_form_answer(self) -> None:
        """Two people share a name; each is shown as they wrote themselves down."""
        rows = [
            SHELL, FORM,
            student(Name="Trần Trọng Ngọc", Email="trantn@gmail.com", Google_ID="300",
                    **{"Student ID": "24009999", "GitHub Username": "loc2"}),
        ]
        groups = list_duplicate_accounts(rows)
        self.assertEqual(len(groups), 1)
        self.assertEqual(
            sorted(ref.student_id for ref in groups[0].members),
            ["24007006", "24009999"],
        )

    def test_one_pair_is_one_finding_under_two_signals(self) -> None:
        rows = [
            student(Name="Lê Văn C", Email="levanc@hus.edu.vn", Google_ID="400",
                    **{"Student ID": "24003333"}),
            student(Name="Lê Văn C", Email="levanc@gmail.com", Google_ID="401",
                    **{"Student ID": "24004444"}),
        ]
        self.assertEqual(len(list_duplicate_accounts(rows)), 1)


class RejectedStudentNumber(unittest.TestCase):
    DINH = student(
        Name="Lê Văn Hà", Email="24007002@hus.edu.vn", Google_ID="500",
        **{"Student ID": None, "GitHub Username": "levanha",
           "Invalid Student ID": "240070020"},
    )

    def test_a_rejected_number_is_still_a_form_answer(self) -> None:
        self.assertTrue(has_form_response(self.DINH))

    def test_he_is_not_listed_as_never_having_answered(self) -> None:
        """Telling a student who did fill the form to fill it sends them looking
        for the wrong mistake."""
        self.assertEqual(list_missing_form([self.DINH]), [])

    def test_the_number_is_reported_with_what_is_wrong_with_it(self) -> None:
        issues = [i for i in list_invalid_info([self.DINH]) if i.field == "Student ID"]
        self.assertEqual(codes(issues), ["student_id_invalid"])
        self.assertEqual(issues[0].severity, SEVERITY_ERROR)
        self.assertEqual(issues[0].found, "240070020")
        self.assertIn("9 chữ số", issues[0].detail)

    def test_the_blank_it_left_is_not_reported_as_a_second_fault(self) -> None:
        self.assertNotIn("student_id_format", codes(list_invalid_info([self.DINH])))

    def test_a_number_that_is_merely_short_says_so(self) -> None:
        short = student(**{"Student ID": None, "Invalid Student ID": "2400199"})
        issues = [i for i in list_invalid_info([short]) if i.code == "student_id_invalid"]
        self.assertIn("thiếu 1", issues[0].detail)

    def test_it_reaches_the_announcement(self) -> None:
        text = format_announcement(audit_roster([self.DINH]))
        self.assertIn("Lê Văn Hà", text)
        self.assertIn("240070020", text)

    def test_a_valid_number_raises_nothing(self) -> None:
        self.assertEqual(
            [c for c in codes(list_invalid_info([student()])) if c.startswith("student_id")],
            [],
        )


class RosterPopulation(unittest.TestCase):
    def test_the_live_roster_is_read_off_the_payload(self) -> None:
        payload = [
            {"username": "AliceVN", "first_name": "Alice", "email": "a@hus.edu.vn"},
            {"username": "bob", "first_name": "Bob", "email": "b@hus.edu.vn"},
        ]
        self.assertEqual(roster_usernames_from_payload(payload), ["AliceVN", "bob"])

    def test_a_roster_row_with_no_account_is_not_a_username(self) -> None:
        payload = [{"username": None, "email": "c@hus.edu.vn"}]
        self.assertEqual(roster_usernames_from_payload(payload), [])

    def test_the_comparison_ignores_case(self) -> None:
        """The database export lower-cases; the roster keeps GitHub's own casing."""
        self.assertEqual(usernames_not_on_roster(["alicevn"], ["AliceVN"]), [])

    def test_a_name_never_imported_is_reported(self) -> None:
        self.assertEqual(usernames_not_on_roster(["alicevn", "levanha"], ["AliceVN"]),
                         ["levanha"])

    def test_uninvited_and_unimported_are_two_different_groups(self) -> None:
        """Both need work, and the command that does it is not the same one."""
        checks = {
            "alicevn": MembershipCheck("AliceVN", MEMBER_ACTIVE, "đã vào org"),
            "bob": MembershipCheck("bob", MEMBER_INVITED, "đã mời"),
            "carol": MembershipCheck("carol", MEMBER_ABSENT, "chưa có lời mời nào"),
        }
        text = format_membership(checks, not_on_roster=["levanha"])
        self.assertIn("CHƯA ĐƯỢC MỜI", text)
        self.assertIn("CHƯA IMPORT LÊN ROSTER", text)
        self.assertIn("levanha", text.split("CHƯA IMPORT LÊN ROSTER")[1])
        self.assertNotIn("levanha", text.split("CHƯA IMPORT LÊN ROSTER")[0])

    def test_the_header_does_not_claim_a_roster_it_could_not_read(self) -> None:
        checks = {"bob": MembershipCheck("bob", MEMBER_INVITED, "đã mời")}
        self.assertIn("trên roster", format_membership(checks))
        self.assertIn("database", format_membership(checks, population="database"))

    def test_an_empty_roster_with_names_to_import_still_reports(self) -> None:
        text = format_membership({}, not_on_roster=["levanha"])
        self.assertIn("levanha", text)


class ImportPurity(unittest.TestCase):
    def test_module_does_not_pull_in_the_database_layer(self) -> None:
        self.assertNotIn("pandas", sys.modules)
        self.assertNotIn("course_hoanganhduc.data", sys.modules)


if __name__ == "__main__":
    unittest.main(verbosity=2)
