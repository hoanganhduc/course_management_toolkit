#!/usr/bin/env python3
"""Offline tests for the roster audit.

The module under test is stdlib-only and imports the database layer lazily, so
these tests run under a bare interpreter with no pandas installed.  Every case
is built from the shapes the real MAT3508 roster actually contains: a student
who never answered the form, a GitHub username that is a person's name, a
personal Gmail on Classroom, and a pair of records that are one person under two
orderings of the same three words.
"""

from __future__ import annotations

import os
import sys
import unittest
from typing import Any, Dict, List

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc.onboard import (  # noqa: E402
    GITHUB_MISSING,
    GITHUB_UNVERIFIED,
    GithubCheck,
)
from course_hoanganhduc.c50_cli import RunResult  # noqa: E402
from course_hoanganhduc.roster_audit import (  # noqa: E402
    MEMBER_ABSENT,
    MEMBER_ACTIVE,
    MEMBER_INVITED,
    MEMBER_UNVERIFIED,
    SEVERITY_ERROR,
    SEVERITY_WARNING,
    audit_roster,
    fold_name,
    format_announcement,
    format_audit,
    format_issues,
    format_membership,
    has_form_response,
    list_duplicate_accounts,
    list_invalid_info,
    list_missing_form,
    report_to_dict,
    verify_org_membership,
)


def student(**fields: Any) -> Dict[str, Any]:
    """A roster record; dicts and objects are both accepted by the audit.

    ``Google_ID`` is derived from the address instead of defaulting to one shared
    constant.  It names a Google account, and the audit reads it as proof of one
    person, so a fixture that hands the same id to two records meant to be two
    students would join them behind the test's back.  A case that means them to
    share an account says so by passing ``Google_ID`` itself.
    """
    base = {
        "Name": "Nguyễn Văn A",
        "Email": "24001111@hus.edu.vn",
        "Student ID": "24001111",
        "GitHub Username": "nguyenvana",
        "Section": "MAT3508 1",
        "Class": "K68A4",
    }
    base.update(fields)
    if "Google_ID" not in fields:
        base["Google_ID"] = "google-" + str(base.get("Email") or base["Name"])
    return {k: v for k, v in base.items() if v is not None}


def codes(issues: List[Any]) -> List[str]:
    return [issue.code for issue in issues]


class FormCoverage(unittest.TestCase):
    def test_student_id_is_the_form_marker(self) -> None:
        self.assertTrue(has_form_response(student()))
        self.assertFalse(has_form_response(student(**{"Student ID": None})))

    def test_missing_form_lists_only_unregistered(self) -> None:
        rows = [student(), student(Name="B", **{"Student ID": None, "Email": "b@hus.edu.vn"})]
        missing = list_missing_form(rows)
        self.assertEqual([ref.name for ref in missing], ["B"])

    def test_one_person_with_two_records_is_counted_once(self) -> None:
        """The commonest shape on the real roster: a Classroom shell plus a form reply."""
        rows = [
            student(
                Name="KHIEM PHAM TUAN",
                **{"Student ID": None, "GitHub Username": None,
                   "Email": "24007003@hus.edu.vn"},
            ),
            student(
                Name="Phạm Tuấn Khiêm",
                **{"Student ID": "24007003", "GitHub Username": "phamtuankhiem1",
                   "Email": "24007003@hus.edu.vn"},
            ),
        ]
        self.assertEqual(list_missing_form(rows), [])

    def test_a_shell_is_linked_by_the_id_inside_its_address(self) -> None:
        """The form reply carries a personal address, so only the number joins them."""
        rows = [
            student(Name="LINH PHAN XUAN",
                    **{"Student ID": None, "Email": "24007004@hus.edu.vn"}),
            student(Name="Phan Xuân Linh",
                    **{"Student ID": "24007004", "Email": "linhpx@gmail.com"}),
        ]
        self.assertEqual(list_missing_form(rows), [])

    def test_a_shell_is_linked_by_the_classroom_address(self) -> None:
        rows = [
            student(Name="B SHELL", **{"Student ID": None, "Email": "b@gmail.com"}),
            student(
                Name="Bê",
                **{"Student ID": "24007010", "Email": "24007010@hus.edu.vn",
                   "Google_Classroom_Email": "b@gmail.com"},
            ),
        ]
        self.assertEqual(list_missing_form(rows), [])

    def test_two_accounts_sharing_no_address_are_both_listed(self) -> None:
        """One name under two logins: still two roster entries, so still two to chase."""
        rows = [
            student(Name="VIET MINH NGO",
                    **{"Student ID": None, "Email": "24007005@hus.edu.vn"}),
            student(Name="MINH NGO VIET",
                    **{"Student ID": None, "Email": "ngovietminh01012000@gmail.com"}),
        ]
        self.assertEqual(
            sorted(ref.name for ref in list_missing_form(rows)),
            ["MINH NGO VIET", "VIET MINH NGO"],
        )

    def test_a_repeated_shell_is_named_once(self) -> None:
        rows = [
            student(Name="B", **{"Student ID": None, "Email": "b@hus.edu.vn"}),
            student(Name="B AGAIN", **{"Student ID": None, "Email": "b@hus.edu.vn"}),
        ]
        self.assertEqual(len(list_missing_form(rows)), 1)

    def test_counts_are_per_person_not_per_record(self) -> None:
        rows = [
            student(Name="KHIEM PHAM TUAN",
                    **{"Student ID": None, "Email": "24007003@hus.edu.vn"}),
            student(Name="Phạm Tuấn Khiêm",
                    **{"Student ID": "24007003", "Email": "24007003@hus.edu.vn"}),
            student(Name="B", **{"Student ID": None, "Email": "b@hus.edu.vn"}),
        ]
        report = audit_roster(rows)
        self.assertEqual(report.total, 2)
        self.assertEqual(report.with_form, 1)
        self.assertEqual(len(report.missing_form), 1)

    def test_form_fields_are_not_checked_for_unregistered_students(self) -> None:
        row = student(
            Name="B",
            **{"Student ID": None, "GitHub Username": None, "Section": None, "Class": None},
        )
        self.assertNotIn("github_missing", codes(list_invalid_info([row])))
        self.assertNotIn("section_missing", codes(list_invalid_info([row])))


class FieldChecks(unittest.TestCase):
    def test_github_username_with_a_space_is_an_error(self) -> None:
        issues = list_invalid_info([student(**{"GitHub Username": "Quan Vu"})])
        found = [i for i in issues if i.code == "github_syntax"]
        self.assertEqual(len(found), 1)
        self.assertEqual(found[0].severity, SEVERITY_ERROR)
        self.assertEqual(found[0].found, "Quan Vu")
        self.assertIn("GitHub", found[0].fix)

    def test_non_school_email_is_reported_with_the_address(self) -> None:
        issues = list_invalid_info([student(Email="24001110@gmail.com")])
        found = [i for i in issues if i.code == "email_not_school"]
        self.assertEqual(len(found), 1)
        self.assertEqual(found[0].found, "24001110@gmail.com")
        self.assertEqual(found[0].severity, SEVERITY_WARNING)

    def test_school_email_is_accepted(self) -> None:
        self.assertEqual(codes(list_invalid_info([student()])), [])

    def test_name_based_school_address_is_not_an_id_mismatch(self) -> None:
        """A HUS address built from a name must never look like a wrong ID."""
        row = student(
            Name="Huỳnh Anh Phúc",
            Email="huynhanhphuc_t00@hus.edu.vn",
            **{"Student ID": "24007007"},
        )
        self.assertNotIn("email_id_mismatch", codes(list_invalid_info([row])))

    def test_numeric_local_part_that_disagrees_with_the_id_is_reported(self) -> None:
        row = student(Email="24009999@hus.edu.vn", **{"Student ID": "24001111"})
        self.assertIn("email_id_mismatch", codes(list_invalid_info([row])))

    def test_a_parked_number_is_reported_while_the_field_is_empty(self) -> None:
        row = student(**{"Student ID": None, "Invalid Student ID": "240011110"})
        found = [i for i in list_invalid_info([row]) if i.code == "student_id_invalid"]
        self.assertEqual(len(found), 1)
        self.assertEqual(found[0].found, "240011110")

    def test_a_corrected_number_retires_the_parked_one(self) -> None:
        """A student who typed one digit too many, then answered the form again.

        The correction reaches ``Student ID``, but the value the first import
        parked stays behind, and the audit read that instead of the field: it
        told the student to fix a number they had already fixed.
        """
        row = student(**{"Invalid Student ID": "240011110"})
        self.assertNotIn("student_id_invalid", codes(list_invalid_info([row])))

    def test_a_number_that_is_still_wrong_is_reported_without_a_parked_value(self) -> None:
        row = student(**{"Student ID": "240011110"})
        self.assertIn("student_id_format", codes(list_invalid_info([row])))

    def test_personal_gmail_on_classroom_is_reported(self) -> None:
        row = student(Google_Classroom_Email="trantrongngoc0101@gmail.com")
        found = [i for i in list_invalid_info([row]) if i.code == "google_account_personal"]
        self.assertEqual(len(found), 1)
        self.assertIn("rời lớp", found[0].fix)

    def test_shared_github_username_names_the_other_student(self) -> None:
        rows = [
            student(Name="A", **{"Student ID": "24001111", "GitHub Username": "shared"}),
            student(
                Name="B",
                Email="24002222@hus.edu.vn",
                Google_ID="2",
                **{"Student ID": "24002222", "GitHub Username": "shared"},
            ),
        ]
        found = [i for i in list_invalid_info(rows) if i.code == "github_shared"]
        self.assertEqual(len(found), 2)
        self.assertIn("B", found[0].detail)
        self.assertIn("A", found[1].detail)

    def test_shared_student_id_is_an_error(self) -> None:
        rows = [
            student(Name="A"),
            student(Name="B", Email="b@hus.edu.vn", Google_ID="2", **{"GitHub Username": "b"}),
        ]
        found = [i for i in list_invalid_info(rows) if i.code == "student_id_shared"]
        self.assertEqual(len(found), 2)
        self.assertEqual(found[0].severity, SEVERITY_ERROR)

    def test_unknown_section_is_reported_only_when_the_list_is_given(self) -> None:
        row = student(Section="MAT3508 9")
        self.assertNotIn("section_unknown", codes(list_invalid_info([row])))
        self.assertIn(
            "section_unknown",
            codes(list_invalid_info([row], expected_sections=["MAT3508 1"])),
        )

    def test_class_spelling_is_measured_against_the_majority(self) -> None:
        rows = [student(Name=f"S{i}", Email=f"2400{i:04d}@hus.edu.vn", Google_ID=str(i),
                        **{"Student ID": f"2400{i:04d}", "GitHub Username": f"s{i}"})
                for i in range(5)]
        rows.append(
            student(
                Name="Odd",
                Email="24009000@hus.edu.vn",
                Google_ID="9",
                Class="K68A4 - Khoa học dữ liệu",
                **{"Student ID": "24009000", "GitHub Username": "odd"},
            )
        )
        found = [i for i in list_invalid_info(rows) if i.code == "class_format"]
        self.assertEqual([i.student.name for i in found], ["Odd"])
        self.assertIn("K68A4", found[0].fix)

    def test_lowercase_class_code_is_still_recognised(self) -> None:
        rows = [student(), student(Name="B", Email="24002222@hus.edu.vn", Google_ID="2",
                                   Class="k68a4",
                                   **{"Student ID": "24002222", "GitHub Username": "b"})]
        found = [i for i in list_invalid_info(rows) if i.code == "class_format"]
        self.assertEqual([i.student.name for i in found], ["B"])

    def test_github_check_absent_means_syntax_only(self) -> None:
        self.assertEqual(codes(list_invalid_info([student()])), [])

    def test_github_reported_missing_is_an_error(self) -> None:
        checks = {"nguyenvana": GithubCheck("nguyenvana", GITHUB_MISSING, "", "404")}
        found = [i for i in list_invalid_info([student()], checks=checks)]
        self.assertEqual(codes(found), ["github_not_found"])
        self.assertEqual(found[0].severity, SEVERITY_ERROR)

    def test_unverified_github_is_a_warning_not_a_rejection(self) -> None:
        checks = {"nguyenvana": GithubCheck("nguyenvana", GITHUB_UNVERIFIED, "", "network")}
        found = list_invalid_info([student()], checks=checks)
        self.assertEqual(codes(found), ["github_unverified"])
        self.assertEqual(found[0].severity, SEVERITY_WARNING)


class DuplicateAccounts(unittest.TestCase):
    def test_fold_name_ignores_accents_and_word_order(self) -> None:
        self.assertEqual(fold_name("MINH NGO VIET"), fold_name("Ngô Việt Minh"))
        self.assertEqual(fold_name("Trần Trọng Ngọc"), "ngoc tran trong")

    def test_same_person_under_two_orderings_is_grouped(self) -> None:
        rows = [
            student(Name="MINH NGO VIET", Email="ngovietminh@gmail.com", Google_ID="1",
                    **{"Student ID": None, "GitHub Username": None}),
            student(Name="VIET MINH NGO", Email="24007005@hus.edu.vn", Google_ID="2",
                    **{"Student ID": None, "GitHub Username": None}),
        ]
        groups = list_duplicate_accounts(rows)
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0].key, "minh ngo viet")
        self.assertEqual(len(groups[0].members), 2)

    def test_same_local_part_under_two_domains_is_grouped(self) -> None:
        rows = [
            student(Name="A", Email="vanan@hus.edu.vn"),
            student(Name="B", Email="vanan@gmail.com", Google_ID="2",
                    **{"Student ID": "24002222", "GitHub Username": "b"}),
        ]
        groups = list_duplicate_accounts(rows)
        self.assertEqual([g.key for g in groups], ["vanan"])

    def test_distinct_students_are_not_grouped(self) -> None:
        rows = [
            student(Name="A"),
            student(Name="B", Email="24002222@hus.edu.vn", Google_ID="2",
                    **{"Student ID": "24002222", "GitHub Username": "b"}),
        ]
        self.assertEqual(list_duplicate_accounts(rows), [])


class ReportShape(unittest.TestCase):
    def rows(self) -> List[Dict[str, Any]]:
        return [
            student(),
            student(Name="B", Email="b@hus.edu.vn", Google_ID="2",
                    **{"Student ID": None, "GitHub Username": None,
                       "Section": None, "Class": None}),
            student(Name="C", Email="24003333@hus.edu.vn", Google_ID="3",
                    **{"Student ID": "24003333", "GitHub Username": "Quan Vu"}),
        ]

    def test_counts_add_up(self) -> None:
        report = audit_roster(self.rows())
        self.assertEqual(report.total, 3)
        self.assertEqual(report.with_form, 2)
        self.assertEqual(len(report.missing_form), 1)

    def test_report_is_json_ready(self) -> None:
        import json

        payload = report_to_dict(audit_roster(self.rows()))
        json.dumps(payload)
        self.assertEqual(payload["total"], 3)
        self.assertEqual(payload["issues"][0]["student"]["name"], "C")

    def test_text_report_quotes_the_value_and_the_fix(self) -> None:
        text = format_issues(audit_roster(self.rows()).issues)
        self.assertIn("Quan Vu", text)
        self.assertIn("PHẢI SỬA", text)
        self.assertIn("->", text)

    def test_audit_text_covers_every_section(self) -> None:
        text = format_audit(audit_roster(self.rows()))
        self.assertIn("Sĩ số: 3", text)
        self.assertIn("chưa điền form", text)


class Announcement(unittest.TestCase):
    def report(self):
        return audit_roster(
            [
                student(),
                student(Name="B", Email="b@hus.edu.vn", Google_ID="2",
                        **{"Student ID": None, "GitHub Username": None}),
            ]
        )

    def test_counts_are_stated(self) -> None:
        text = format_announcement(self.report())
        self.assertIn("2 bạn trên Google Classroom", text)
        self.assertIn("1 bạn chưa điền", text)

    def test_the_register_is_minh_and_ban_not_thay_and_em(self) -> None:
        text = format_announcement(self.report(), deadline="10/09")
        self.assertIn("Mình đã đối chiếu", text)
        self.assertNotIn("thầy", text)
        self.assertNotIn("các em", text)
        self.assertNotIn("em nào", text)

    def test_students_who_have_not_filled_the_form_are_named(self) -> None:
        text = format_announcement(self.report())
        self.assertIn("CHƯA ĐIỀN FORM (1 bạn)", text)
        self.assertIn("b@hus.edu.vn", text)

    def test_each_wrong_field_is_named_with_its_value_and_its_fix(self) -> None:
        rows = [
            student(Name="C", Email="24003333@hus.edu.vn", Google_ID="3",
                    **{"Student ID": "24003333", "GitHub Username": "Quan Vu"}),
        ]
        text = format_announcement(audit_roster(rows))
        self.assertIn("THÔNG TIN ĐIỀN CHƯA ĐÚNG (1 bạn)", text)
        self.assertIn("C (24003333)", text)
        self.assertIn("GitHub Username:", text)
        self.assertIn('bạn điền "Quan Vu"', text)
        self.assertIn("=>", text)

    def test_the_wrong_value_is_quoted_once_not_twice(self) -> None:
        rows = [student(**{"GitHub Username": "Quan Vu"})]
        text = format_announcement(audit_roster(rows))
        self.assertEqual(text.count("Quan Vu"), 1)

    def test_suspected_duplicate_accounts_are_listed(self) -> None:
        rows = [
            student(Name="MINH NGO VIET", Email="ngovm@gmail.com", Google_ID="1",
                    **{"Student ID": None, "GitHub Username": None}),
            student(Name="VIET MINH NGO", Email="24007005@hus.edu.vn", Google_ID="2",
                    **{"Student ID": None, "GitHub Username": None}),
        ]
        text = format_announcement(audit_roster(rows))
        self.assertIn("NGHI TRÙNG TÀI KHOẢN", text)
        self.assertIn("MINH NGO VIET (ngovm@gmail.com)", text)
        self.assertIn("VIET MINH NGO (24007005@hus.edu.vn)", text)

    def test_a_suspected_duplicate_is_not_told_to_leave_the_course(self) -> None:
        rows = [
            student(Name="MINH NGO VIET", Email="ngovm@gmail.com", Google_ID="1",
                    **{"Student ID": None, "GitHub Username": None}),
            student(Name="VIET MINH NGO", Email="24007005@hus.edu.vn", Google_ID="2",
                    **{"Student ID": None, "GitHub Username": None}),
        ]
        text = format_announcement(audit_roster(rows))
        self.assertIn("Đừng tự rời lớp", text)
        self.assertNotIn("rời lớp bằng tài khoản", text)

    def test_harmless_findings_stay_out_of_the_post(self) -> None:
        """The audit keeps reporting the address; the post no longer repeats it."""
        rows = [
            student(),
            student(Name="Trần Thị B", Email="24002222@hus.edu.vn", Google_ID="2",
                    **{"Student ID": "24002222", "GitHub Username": "tranthib"}),
            student(Name="D", Email="d@gmail.com", Google_ID="4",
                    **{"Student ID": "24004444", "GitHub Username": "dee",
                       "Class": "k68a4"}),
        ]
        report = audit_roster(rows)
        self.assertLessEqual(
            {"email_not_school", "google_account_personal", "class_format"},
            set(codes(report.issues)),
        )
        text = format_announcement(report)
        self.assertNotIn("THÔNG TIN ĐIỀN CHƯA ĐÚNG", text)
        self.assertNotIn("d@gmail.com", text)
        self.assertNotIn("k68a4", text)

    def test_a_wrong_student_number_reaches_the_post(self) -> None:
        rows = [
            student(Name="E", Email="e@hus.edu.vn", Google_ID="5",
                    **{"Student ID": "2400", "GitHub Username": "eee"}),
        ]
        text = format_announcement(audit_roster(rows))
        self.assertIn("THÔNG TIN ĐIỀN CHƯA ĐÚNG (1 bạn)", text)
        self.assertIn('bạn điền "2400"', text)

    def test_two_students_sharing_one_github_account_are_both_named(self) -> None:
        rows = [
            student(),
            student(Name="F", Email="24006666@hus.edu.vn", Google_ID="6",
                    **{"Student ID": "24006666", "GitHub Username": "nguyenvana"}),
        ]
        text = format_announcement(audit_roster(rows))
        self.assertIn("Nguyễn Văn A (24001111)", text)
        self.assertIn("F (24006666)", text)
        self.assertIn("GitHub Username:", text)

    def test_empty_sections_are_dropped_and_numbering_stays_contiguous(self) -> None:
        text = format_announcement(audit_roster([student()]))
        self.assertNotIn("CHƯA ĐIỀN FORM", text)
        self.assertNotIn("THÔNG TIN ĐIỀN CHƯA ĐÚNG", text)
        self.assertIn("1. CÁCH TỰ KIỂM TRA", text)

    def test_form_url_and_deadline_appear_when_given(self) -> None:
        text = format_announcement(
            self.report(), form_url="https://example.org/form", deadline="23:59 ngày 10/09"
        )
        self.assertIn("https://example.org/form", text)
        self.assertIn("23:59 ngày 10/09", text)

    def test_course_name_is_used_as_the_heading(self) -> None:
        self.assertTrue(format_announcement(self.report(), course_name="MAT3508").startswith("[MAT3508]"))


class Gh:
    """A ``gh`` stand-in that answers the two endpoints the check calls.

    ``members`` are the logins the org reports as members; ``invited`` are the
    raw lines the invitation listing prints, which is where an email-only
    invitation shows up as ``null``.
    """

    def __init__(self, *, members=(), invited=(), invitations_fail=False,
                 members_fail=False) -> None:
        self.members = {str(m).lower() for m in members}
        self.invited = list(invited)
        self.invitations_fail = invitations_fail
        self.members_fail = members_fail
        self.calls: List[List[str]] = []

    def __call__(self, argv: List[str]) -> RunResult:
        self.calls.append(list(argv))
        path = next((a for a in argv if a.startswith("/orgs/")), "")
        if path.endswith("/invitations"):
            if self.invitations_fail:
                return RunResult(1, "", "HTTP 403: API rate limit exceeded")
            body = "".join(f"{line}\n" for line in self.invited)
            return RunResult(0, body, "")
        if self.members_fail:
            return RunResult(1, "", "HTTP 403: API rate limit exceeded")
        if path.rsplit("/", 1)[-1].lower() in self.members:
            return RunResult(0, "", "")
        return RunResult(1, "", "gh: Not Found (HTTP 404)")

    def invitation_calls(self) -> List[List[str]]:
        return [c for c in self.calls if any(a.endswith("/invitations") for a in c)]


class OrgMembership(unittest.TestCase):
    def test_a_member_is_reported_as_a_member(self) -> None:
        gh = Gh(members=["alice"])
        checks = verify_org_membership(["alice"], "VNU-HUS", runner=gh)
        self.assertEqual(checks["alice"].state, MEMBER_ACTIVE)

    def test_a_pending_invitation_is_not_absence(self) -> None:
        """The 404 trap: not a member and not missed, merely not accepted yet."""
        gh = Gh(members=[], invited=["alice"])
        checks = verify_org_membership(["alice"], "VNU-HUS", runner=gh)
        self.assertEqual(checks["alice"].state, MEMBER_INVITED)

    def test_absence_needs_both_answers(self) -> None:
        gh = Gh(members=[], invited=[])
        checks = verify_org_membership(["alice"], "VNU-HUS", runner=gh)
        self.assertEqual(checks["alice"].state, MEMBER_ABSENT)

    def test_logins_match_whatever_the_case(self) -> None:
        gh = Gh(members=[], invited=["Alice"])
        checks = verify_org_membership(["ALICE"], "VNU-HUS", runner=gh)
        self.assertEqual(checks["alice"].state, MEMBER_INVITED)

    def test_an_unreadable_invitation_list_leaves_nobody_absent(self) -> None:
        gh = Gh(members=[], invitations_fail=True)
        checks = verify_org_membership(["alice"], "VNU-HUS", runner=gh)
        self.assertEqual(checks["alice"].state, MEMBER_UNVERIFIED)

    def test_a_rate_limited_member_call_decides_nothing(self) -> None:
        gh = Gh(members_fail=True)
        checks = verify_org_membership(["alice"], "VNU-HUS", runner=gh)
        self.assertEqual(checks["alice"].state, MEMBER_UNVERIFIED)

    def test_an_email_only_invitation_carries_no_login(self) -> None:
        gh = Gh(members=[], invited=["null"])
        checks = verify_org_membership(["null"], "VNU-HUS", runner=gh)
        self.assertEqual(checks["null"].state, MEMBER_ABSENT)

    def test_the_invitation_list_is_read_once_for_the_whole_roster(self) -> None:
        gh = Gh(members=["alice"], invited=["bob"])
        verify_org_membership(["alice", "bob", "carol"], "VNU-HUS", runner=gh)
        self.assertEqual(len(gh.invitation_calls()), 1)

    def test_skip_asks_github_nothing(self) -> None:
        gh = Gh(members=["alice"])
        checks = verify_org_membership(["alice"], "VNU-HUS", runner=gh, skip=True)
        self.assertEqual(checks["alice"].state, MEMBER_UNVERIFIED)
        self.assertEqual(gh.calls, [])

    def test_only_the_uninvited_are_written_up_as_work(self) -> None:
        gh = Gh(members=["alice"], invited=["bob"])
        text = format_membership(
            verify_org_membership(["alice", "bob", "carol"], "VNU-HUS", runner=gh)
        )
        self.assertIn("CHƯA ĐƯỢC MỜI", text)
        self.assertIn("carol", text)
        head = text.split("CHƯA ĐƯỢC MỜI")[0]
        self.assertIn("alice", head)
        self.assertIn("bob", head)

    def test_an_empty_roster_asks_github_nothing(self) -> None:
        gh = Gh()
        self.assertEqual(verify_org_membership([], "VNU-HUS", runner=gh), {})
        self.assertEqual(gh.calls, [])


class ImportPurity(unittest.TestCase):
    def test_module_does_not_pull_in_the_database_layer(self) -> None:
        self.assertNotIn("pandas", sys.modules)
        self.assertNotIn("course_hoanganhduc.data", sys.modules)


if __name__ == "__main__":
    unittest.main(verbosity=2)
