#!/usr/bin/env python3
"""Offline tests for the CSV -> database -> Google Classroom + Classroom50 driver.

Every outward call is injected: `gh` goes through a fake ``Runner``, Google
Classroom through a fake service object, and the database layer through a stub
in ``sys.modules`` (``onboard`` imports it lazily, inside the function).  So
nothing here reaches the network, the clock, or a real roster.
"""

from __future__ import annotations

import csv
import io
import os
import sys
import types
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from test_gclass_invite import (  # noqa: E402
    FakeClassroom,
    FakeHttpError,
    _Request,
    _page,
    enrolled,
)

from course_hoanganhduc import gclass_invite, onboard  # noqa: E402
from course_hoanganhduc.c50_cli import RunResult  # noqa: E402
from course_hoanganhduc.c50_flags import (  # noqa: E402
    SOURCE_CONFIG,
    SOURCE_FLAG,
    Resolved,
    _resolve_classroom,
    _resolve_org,
)
from course_hoanganhduc.c50_roster import CANONICAL_COLUMNS  # noqa: E402
from course_hoanganhduc.models import Student  # noqa: E402
from course_hoanganhduc.onboard import (  # noqa: E402
    C50_NO_USERNAME,
    C50_ON_ROSTER_MEMBER,
    C50_ON_ROSTER_NOT_MEMBER,
    C50_ON_ROSTER_UNKNOWN_MEMBER,
    C50_UNLINKED_EMAIL,
    C50_WILL_ADD,
    EMAIL_SOURCE_PERSONAL,
    EMAIL_SOURCE_SCHOOL,
    GC_ENROLLED,
    GC_PENDING,
    GC_WILL_INVITE,
    GITHUB_EXISTS,
    GITHUB_MISSING,
    GITHUB_UNVERIFIED,
    ResolvedRow,
    _render_report,
    check_github_exists,
    classify_rows,
    confirm_pending_by_email,
    read_classroom50_state,
    read_google_state,
    resolve_rows,
    run_onboarding,
    validate_email,
    validate_github_username,
)
from course_hoanganhduc.roster_csv import ParseResult, parse_student_rows  # noqa: E402


class FakeCourse(FakeClassroom):
    """FakeClassroom plus the ``userId`` filter ``invitations.list`` accepts.

    The driver's preview asks ``invitations.list(courseId=..., userId=<email>)``,
    which the API documents and the invite path never uses, so the borrowed fake
    does not model it.
    """

    def _inv_list(self, courseId=None, pageSize=None, pageToken=None, userId=None):
        if userId is not None:
            self.calls.append(("invitations.list.byUser", courseId, userId))
            matches = [
                inv
                for inv in self.pending
                if self.profiles.get(str(inv.get("userId"))) == userId
            ]
            return _Request(lambda: {"invitations": matches} if matches else {})
        return super()._inv_list(courseId, pageSize, pageToken)


# --- CSV fixtures -----------------------------------------------------------

# The real Google Form header, copied verbatim: bilingual parentheses, the typo
# in "Emai VNU-HUS", and the required-field asterisk all included.
FORM_HEADER = (
    "Timestamp,Email Address,Họ và Tên (Full Name),Emai VNU-HUS,GitHub Username,"
    "Mã Sinh Viên (Student ID),Ngày sinh (Date of Birth),Lớp (Class) *,"
    "Lớp học phần (Course Section)"
)


def form_row(
    *,
    timestamp: str = "9/1/2026 8:00:00",
    personal: str = "an.nguyen@gmail.com",
    name: str = "Nguyễn Văn A",
    school: str = "an.nguyen@vnu.edu.vn",
    github: str = "annguyen",
    student_id: str = "22001234",
    dob: str = "1/1/2005",
    class_name: str = "K70A9",
    section: str = "MAT3500 1",
) -> str:
    return ",".join(
        [timestamp, personal, name, school, github, student_id, dob, class_name, section]
    )


def form_csv(*rows: str, header: str = FORM_HEADER) -> str:
    return "\n".join([header, *rows]) + "\n"


def resolved(**kwargs: Any) -> ResolvedRow:
    base: Dict[str, Any] = {
        "student_id": "22001234",
        "name": "Nguyễn Văn A",
        "email": "an.nguyen@vnu.edu.vn",
        "email_source": EMAIL_SOURCE_SCHOOL,
        "github_username": "annguyen",
        "github_id": "9001",
        "section": "MAT3500 1",
        "class_name": "K70A9",
        "timestamp": "",
        "notes": [],
    }
    base.update(kwargs)
    return ResolvedRow(**base)


# --- fakes ------------------------------------------------------------------


class FakeRunner:
    """Answers the `gh` calls the driver makes, and records every argv.

    Keyed on the command shape rather than the exact argv so a test only has to
    state the parts it cares about; anything unclaimed comes back as a failure
    rather than a silent success.
    """

    def __init__(
        self,
        *,
        users: Optional[Dict[str, Any]] = None,
        roster: Optional[List[Dict[str, Any]]] = None,
        members: Optional[List[Dict[str, Any]]] = None,
        classrooms: Optional[List[Dict[str, Any]]] = None,
        sync_exit: int = 0,
        sync_output: str = "",
        member_list_fails: bool = False,
    ) -> None:
        self.users = dict(users or {})
        self.roster = list(roster or [])
        self.members = list(members or [])
        self.classrooms = list(classrooms or [])
        self.sync_exit = sync_exit
        self.sync_output = sync_output
        self.member_list_fails = member_list_fails
        self.calls: List[List[str]] = []
        self.imported_csv: Optional[str] = None

    def __call__(self, argv: List[str]) -> RunResult:
        import json

        self.calls.append(list(argv))
        if argv[:2] == ["gh", "api"]:
            name = argv[2].rsplit("/", 1)[-1]
            answer = self.users.get(name.lower())
            if answer is None:
                return RunResult(1, "", "HTTP 404: Not Found (/users)")
            if isinstance(answer, str) and answer.startswith("!"):
                return RunResult(1, "", answer[1:])
            return RunResult(0, f"{answer}\n", "")
        parts = argv[2:]
        if parts[:2] == ["classroom", "list"]:
            return RunResult(0, json.dumps(self.classrooms), "")
        if parts[:2] == ["roster", "list"]:
            return RunResult(0, json.dumps(self.roster), "")
        if parts[:2] == ["member", "list"]:
            if self.member_list_fails:
                return RunResult(1, "", "HTTP 403: missing admin:org scope")
            return RunResult(0, json.dumps(self.members), "")
        if parts[:2] == ["roster", "sync"]:
            return RunResult(self.sync_exit, self.sync_output, "")
        if parts[:2] == ["roster", "import"]:
            self.imported_csv = Path(argv[-1]).read_text(encoding="utf-8")
            return RunResult(0, "", "")
        return RunResult(1, "", f"unexpected argv: {argv}")

    # -- assertion helpers
    def ran(self, *needle: str) -> List[List[str]]:
        n = list(needle)
        return [c for c in self.calls if any(c[i : i + len(n)] == n for i in range(len(c)))]

    def import_rows(self) -> List[Dict[str, str]]:
        return list(csv.DictReader(io.StringIO(self.imported_csv or "")))


def roster_row(**kwargs: Any) -> Dict[str, Any]:
    row: Dict[str, Any] = {"username": "", "email": "", "github_id": ""}
    row.update(kwargs)
    return row


class _OnboardCase(unittest.TestCase):
    """Installs the database stub and the prompt stubs the driver goes through."""

    def setUp(self) -> None:
        self.saved: List[Dict[str, Any]] = []
        self.db: List[Student] = []
        self._real_data = sys.modules.get("course_hoanganhduc.data")
        self._real_input = gclass_invite.get_input_with_quit
        # The invite path runs its own numbered selection and confirmation; the
        # driver reuses it deliberately, so both answers have to be supplied.
        self._invite_answers = ["a", "y"]
        gclass_invite.get_input_with_quit = lambda *a, **k: (
            self._invite_answers.pop(0) if self._invite_answers else "y"
        )
        self._tmp: List[Any] = []
        # Skipping Google Classroom is what `_select_course_interactively`
        # returning None means; standing in for it keeps the tests off the
        # credentials file.
        self._real_picker = onboard.pick_google_course
        onboard.pick_google_course = lambda *a, **k: None

    def tearDown(self) -> None:
        onboard.pick_google_course = self._real_picker
        gclass_invite.get_input_with_quit = self._real_input
        if self._real_data is None:
            sys.modules.pop("course_hoanganhduc.data", None)
        else:
            sys.modules["course_hoanganhduc.data"] = self._real_data
        for directory in self._tmp:
            directory.cleanup()

    def install_db(self, students: Optional[List[Student]] = None) -> None:
        self.db = list(students or [])
        module = types.ModuleType("course_hoanganhduc.data")
        module.load_database = lambda db_path, verbose=False: self.db
        def save_database(rows, db_path=None, verbose=False, audit_source=None):
            # A real save is visible to the next load, and the Classroom50 import
            # runs its own load; recording without storing would hide the
            # students the driver had just written.
            self.db[:] = list(rows)
            self.saved.append({"rows": list(rows), "audit_source": audit_source})

        module.save_database = save_database
        sys.modules["course_hoanganhduc.data"] = module

    def write_csv(self, text: str) -> str:
        import tempfile

        directory = tempfile.TemporaryDirectory()
        self._tmp.append(directory)
        path = Path(directory.name) / "form.csv"
        path.write_text(text, encoding="utf-8")
        return str(path)

    def run_driver(
        self,
        csv_text: str,
        *,
        answers: Optional[List[str]] = None,
        students: Optional[List[Student]] = None,
        **kwargs: Any,
    ):
        if "course_hoanganhduc.data" not in sys.modules or students is not None:
            self.install_db(students)
        queue = list(answers if answers is not None else ["y"])
        params: Dict[str, Any] = {
            "db_path": "students.db",
            "csv_path": self.write_csv(csv_text),
            "input_fn": lambda prompt: queue.pop(0) if queue else "n",
            "tty_check": lambda: True,
        }
        params.update(kwargs)
        buffer = io.StringIO()
        with redirect_stdout(buffer):
            report = run_onboarding(**params)
        return report, buffer.getvalue()


# --- 1-3: the real form header ---------------------------------------------


class TestHeaderMapping(unittest.TestCase):
    def test_01_real_form_header_maps_every_column(self):
        parsed = parse_student_rows(form_csv(form_row()))
        self.assertEqual(len(parsed.rows), 1)
        row = parsed.rows[0]
        self.assertEqual(row["Student ID"], "22001234")
        self.assertEqual(row["Name"], "Nguyễn Văn A")
        self.assertEqual(row["Email"], "an.nguyen@vnu.edu.vn")
        self.assertEqual(row["Personal Email"], "an.nguyen@gmail.com")
        self.assertEqual(row["GitHub Username"], "annguyen")
        self.assertEqual(row["Section"], "MAT3500 1")
        self.assertEqual(row["Class"], "K70A9")

    def test_02_parentheses_and_asterisk_do_not_break_aliases(self):
        header = (
            "Mã Sinh Viên (Student ID),Họ và Tên (Full Name),Lớp (Class) *,"
            "Lớp học phần (Course Section),GitHub Username"
        )
        parsed = parse_student_rows(header + "\n22001234,Nguyễn Văn A,K70A9,MAT3500 1,annguyen\n")
        row = parsed.rows[0]
        self.assertEqual(row["Student ID"], "22001234")
        self.assertEqual(row["Name"], "Nguyễn Văn A")
        self.assertEqual(row["Class"], "K70A9")
        self.assertEqual(row["Section"], "MAT3500 1")

    def test_03_typo_emai_vnu_hus_still_maps_to_email(self):
        header = "Mã Sinh Viên (Student ID),Họ và Tên (Full Name),Emai VNU-HUS,Email Address"
        parsed = parse_student_rows(
            header + "\n22001234,Nguyễn Văn A,school@vnu.edu.vn,personal@gmail.com\n"
        )
        row = parsed.rows[0]
        self.assertEqual(row["Email"], "school@vnu.edu.vn")
        self.assertEqual(row["Personal Email"], "personal@gmail.com")


# --- 4-7: which address ------------------------------------------------------


class TestEmailChoice(unittest.TestCase):
    def resolve(self, **row: str):
        return resolve_rows(ParseResult([dict(row)], 0, []), {})

    def test_04_school_address_wins_when_both_are_present(self):
        result = self.resolve(
            **{
                "Student ID": "1",
                "Name": "A",
                "Email": "a@vnu.edu.vn",
                "Personal Email": "a@gmail.com",
            }
        )
        self.assertEqual(result.ok[0].email, "a@vnu.edu.vn")
        self.assertEqual(result.ok[0].email_source, EMAIL_SOURCE_SCHOOL)

    def test_05_empty_school_address_falls_back_to_the_personal_one(self):
        result = self.resolve(
            **{"Student ID": "1", "Name": "A", "Email": "", "Personal Email": "a@gmail.com"}
        )
        self.assertEqual(result.ok[0].email, "a@gmail.com")
        self.assertEqual(result.ok[0].email_source, EMAIL_SOURCE_PERSONAL)

    def test_06_malformed_school_address_falls_back_without_rejecting_the_row(self):
        result = self.resolve(
            **{
                "Student ID": "1",
                "Name": "A",
                "Email": "a@@vnu",
                "Personal Email": "a@gmail.com",
            }
        )
        self.assertEqual(result.rejected, [])
        self.assertEqual(result.ok[0].email, "a@gmail.com")
        self.assertEqual(result.ok[0].email_source, EMAIL_SOURCE_PERSONAL)

    def test_07_both_addresses_bad_rejects_the_row_and_names_both_columns(self):
        result = self.resolve(
            **{"Student ID": "1", "Name": "A", "Email": "nope", "Personal Email": "also nope"}
        )
        self.assertEqual(result.ok, [])
        self.assertEqual(len(result.rejected), 1)
        reason = result.rejected[0].reason
        self.assertIn("email trường", reason)
        self.assertIn("email cá nhân", reason)

    def test_07c_a_usable_username_keeps_an_addressless_row_but_names_it(self):
        """Classroom50 takes the row; Google Classroom never can, so say so."""
        result = self.resolve(
            **{
                "Student ID": "1",
                "Name": "Hoàng Văn E",
                "Email": "nope",
                "Personal Email": "also nope",
                "GitHub Username": "hoangvane",
            }
        )
        self.assertEqual(result.rejected, [])
        self.assertEqual(result.ok[0].email, "")
        rendered = _render_report({"read": 1, "csv": "f.csv"}, result)
        self.assertIn("Không có email dùng được (1)", rendered)
        self.assertIn("Hoàng Văn E (1)", rendered)

    def test_07b_lowercases_and_strips_the_address(self):
        self.assertEqual(validate_email("  MailTo:A.B@VNU.Edu.VN "), ("a.b@vnu.edu.vn", None))
        self.assertEqual(validate_email("a b@x.edu")[0], "")


# --- 8-11: the GitHub username ----------------------------------------------


class TestGithubValidation(unittest.TestCase):
    def test_08_profile_url_is_normalised_to_the_bare_name(self):
        self.assertEqual(validate_github_username("https://github.com/abc/"), ("abc", None))
        self.assertEqual(validate_github_username("@abc"), ("abc", None))
        self.assertEqual(validate_github_username("github.com/abc?tab=repos"), ("abc", None))

    def test_09_syntactically_invalid_names_are_rejected(self):
        for bad in ("-abc", "abc-", "a--b", "a" * 40, "ab c"):
            name, reason = validate_github_username(bad)
            self.assertEqual(name, "", bad)
            self.assertTrue(reason, bad)

    def test_10_a_404_from_gh_api_rejects_the_row(self):
        runner = FakeRunner(users={})
        checks = check_github_exists(["ghost"], runner=runner)
        self.assertEqual(checks["ghost"].state, GITHUB_MISSING)
        result = resolve_rows(
            ParseResult(
                [{"Student ID": "1", "Name": "A", "Email": "a@x.edu", "GitHub Username": "ghost"}],
                0,
                [],
            ),
            checks,
        )
        self.assertEqual(result.ok, [])
        self.assertIn("không tồn tại", result.rejected[0].reason)

    def test_11_a_network_error_leaves_the_row_in_place_as_unverified(self):
        runner = FakeRunner(users={"annguyen": "!dial tcp: no route to host"})
        checks = check_github_exists(["annguyen"], runner=runner)
        self.assertEqual(checks["annguyen"].state, GITHUB_UNVERIFIED)
        result = resolve_rows(
            ParseResult(
                [
                    {
                        "Student ID": "1",
                        "Name": "A",
                        "Email": "a@x.edu",
                        "GitHub Username": "annguyen",
                    }
                ],
                0,
                [],
            ),
            checks,
        )
        self.assertEqual(len(result.ok), 1)
        self.assertEqual(result.ok[0].github_id, "")
        self.assertTrue(any("chưa kiểm được" in n for n in result.ok[0].notes))

    def test_11b_an_existing_account_contributes_its_id(self):
        runner = FakeRunner(users={"annguyen": 9001})
        checks = check_github_exists(["annguyen"], runner=runner)
        self.assertEqual(checks["annguyen"].state, GITHUB_EXISTS)
        self.assertEqual(checks["annguyen"].github_id, "9001")


# --- 12, 13, 16, 33, 34: duplicates and rejections ---------------------------


class TestDedupeAndRejection(_OnboardCase):
    def test_12_a_repeated_student_id_keeps_the_later_submission(self):
        text = form_csv(
            form_row(timestamp="9/1/2026 8:00:00", github="oldname", section="MAT3500 1"),
            form_row(timestamp="9/2/2026 9:00:00", github="newname", section="MAT3500 2"),
        )
        parsed = parse_student_rows(text)
        result = resolve_rows(parsed, {})
        self.assertEqual(len(result.ok), 1)
        self.assertEqual(result.ok[0].github_username, "newname")
        self.assertEqual(result.ok[0].section, "MAT3500 2")
        self.assertEqual(len(result.duplicates), 1)

    def test_13_a_row_without_a_student_id_is_dropped_and_the_others_run_on(self):
        # parse_student_rows enforces Student ID and Name per row, so such a row
        # never reaches resolve_rows; it is counted, not silently lost.
        text = form_csv(
            form_row(student_id="", name="Không Mã", github="nomssv"),
            form_row(),
        )
        report, _ = self.run_driver(text, dry_run=True)
        self.assertEqual(report["skippedIncomplete"], 1)
        self.assertEqual(report["read"], 1)
        self.assertEqual(report["resolved"], 1)

    def test_16_every_dropped_row_is_named_with_its_reason_in_the_report(self):
        text = form_csv(
            form_row(student_id="1", name="Sai Email", school="bad", personal="worse", github=""),
            form_row(student_id="2", name="Sai GitHub", github="-nope"),
            form_row(student_id="3", name="Hợp Lệ", github="hople"),
        )
        report, output = self.run_driver(text, dry_run=True)
        reasons = {item["name"]: item["reason"] for item in report["rejected"]}
        self.assertEqual(set(reasons), {"Sai Email", "Sai GitHub"})
        self.assertIn("email trường", reasons["Sai Email"])
        self.assertIn("GitHub Username", reasons["Sai GitHub"])
        for name in ("Sai Email", "Sai GitHub"):
            self.assertIn(name, output)

    def test_33_two_students_claiming_one_username_both_go_to_rejected(self):
        text = form_csv(
            form_row(student_id="1", name="Người Một", github="shared"),
            form_row(student_id="2", name="Người Hai", github="shared"),
        )
        result = resolve_rows(parse_student_rows(text), {})
        self.assertEqual(result.ok, [])
        self.assertEqual(len(result.rejected), 2)
        for item in result.rejected:
            self.assertIn("Người Một", item.reason)
            self.assertIn("Người Hai", item.reason)

    def test_34_without_a_student_id_the_username_becomes_the_dedupe_key(self):
        # Built directly, because parse_student_rows would drop these rows first;
        # the fallback chain guards resolve_rows' other callers.
        rows = [
            {"Student ID": "", "Name": "A", "Email": "a@x.edu", "GitHub Username": "same"},
            {"Student ID": "", "Name": "A", "Email": "a2@x.edu", "GitHub Username": "same"},
        ]
        result = resolve_rows(ParseResult(rows, 0, []), {})
        self.assertEqual(len(result.ok), 1)
        self.assertEqual(result.ok[0].email, "a2@x.edu")
        self.assertEqual(len(result.duplicates), 1)

    def test_34b_a_row_with_no_usable_identifier_at_all_is_rejected(self):
        rows = [{"Student ID": "", "Name": "A", "Email": "", "GitHub Username": ""}]
        result = resolve_rows(ParseResult(rows, 0, []), {})
        self.assertEqual(result.ok, [])
        self.assertEqual(len(result.rejected), 1)


# --- 14, 15, 37: skipping, dry run, and the file picker ----------------------


class TestSkipAndDryRun(_OnboardCase):
    def test_14_skipping_both_platforms_still_loads_the_database(self):
        runner = FakeRunner(users={"annguyen": 9001})
        report, output = self.run_driver(
            form_csv(form_row()),
            runner=runner,
            students=[],
            google_course_id=None,
            c50_org=None,
        )
        self.assertEqual(report["google"], {"status": "skipped"})
        self.assertEqual(report["c50"], {"status": "skipped"})
        self.assertTrue(report["databaseSaved"])
        self.assertEqual(len(self.saved), 1)
        self.assertEqual(runner.ran("teacher"), [])
        self.assertIn("Bỏ qua Google Classroom", output)
        self.assertIn("Bỏ qua Classroom50", output)

    def test_15_dry_run_writes_nothing_and_makes_no_gh_call(self):
        runner = FakeRunner(users={"annguyen": 9001})
        report, output = self.run_driver(
            form_csv(form_row()),
            runner=runner,
            dry_run=True,
            google_course_id="C",
            c50_org="org",
            c50_classroom="lop",
        )
        self.assertEqual(runner.calls, [])
        self.assertEqual(self.saved, [])
        self.assertFalse(report["databaseSaved"])
        self.assertEqual(report["resolved"], 1)
        self.assertIn("Dry run", output)

    def test_15b_dry_run_marks_github_unverified_rather_than_pretending(self):
        report, _ = self.run_driver(
            form_csv(form_row()), runner=FakeRunner(), dry_run=True
        )
        self.assertEqual(report["unverifiedGithub"], ["annguyen"])

    def test_37_the_csv_picker_is_the_shared_input_with_completion(self):
        seen: Dict[str, Any] = {}

        def fake_input_with_completion(prompt, **kwargs):
            seen.update(kwargs)
            return "/tmp/x.csv"

        from course_hoanganhduc import utils

        real = utils.input_with_completion
        utils.input_with_completion = fake_input_with_completion
        try:
            self.assertEqual(onboard.pick_csv_path(), "/tmp/x.csv")
        finally:
            utils.input_with_completion = real
        self.assertTrue(seen["select_file"])
        self.assertTrue(seen["file_filter"]("Roster.CSV"))
        self.assertFalse(seen["file_filter"]("notes.txt"))
        # No second file lister to keep in step with the first one.
        source = Path(onboard.__file__).read_text(encoding="utf-8")
        self.assertNotIn("os.listdir", source)
        self.assertNotIn("iterdir", source)


# --- 17-19, 32: Google Classroom state ---------------------------------------


class TestGoogleState(unittest.TestCase):
    def state(self, fake: FakeCourse, **kwargs: Any):
        buffer = io.StringIO()
        with redirect_stdout(buffer):
            return read_google_state(fake, "C", **kwargs)

    def test_17_an_enrolled_student_is_labelled_enrolled_not_invited(self):
        fake = FakeCourse(students=[enrolled("An.Nguyen@VNU.edu.vn")])
        plans = classify_rows([resolved()], self.state(fake), None)
        self.assertEqual(plans[0].google, GC_ENROLLED)

    def test_18_a_google_id_match_survives_a_changed_email(self):
        fake = FakeCourse(students=[enrolled("old.address@vnu.edu.vn", user_id="g-1")])
        state = self.state(fake)
        self.assertIn("g-1", state.enrolled_user_ids)
        # The row's address is not on the course, so the email tier alone would
        # invite again; the id tier is what stops it.
        self.assertNotIn("an.nguyen@vnu.edu.vn", state.enrolled_emails)

    def test_19_a_pending_invitation_is_labelled_pending(self):
        fake = FakeCourse(
            invitations=[{"id": "i1", "userId": "u1"}],
            profiles={"u1": "an.nguyen@vnu.edu.vn"},
        )
        plans = classify_rows([resolved()], self.state(fake), None)
        self.assertEqual(plans[0].google, GC_PENDING)

    def test_32_a_pending_invitation_past_the_lookup_cap_is_caught_by_userid_list(self):
        fake = FakeCourse(
            invitations=[{"id": "i1", "userId": "u1"}],
            profiles={"u1": "an.nguyen@vnu.edu.vn"},
        )
        state = self.state(fake, profile_lookup_limit=0)
        plans = classify_rows([resolved()], state, None)
        self.assertEqual(plans[0].google, GC_WILL_INVITE)

        confirmed = confirm_pending_by_email(fake, "C", ["an.nguyen@vnu.edu.vn"])
        self.assertEqual(confirmed, {"an.nguyen@vnu.edu.vn"})
        plans = classify_rows([resolved()], state, None, pending_confirmed=confirmed)
        self.assertEqual(plans[0].google, GC_PENDING)

    def test_32b_a_failed_lookup_is_not_read_as_no_invitation(self):
        class Broken:
            def invitations(self):
                def boom(**kwargs):
                    raise FakeHttpError(500)

                return types.SimpleNamespace(list=boom)

        self.assertEqual(confirm_pending_by_email(Broken(), "C", ["a@x.edu"]), set())


# --- 20, 21, 29, 30, 31: Classroom50 state -----------------------------------


class TestClassroom50State(unittest.TestCase):
    def state(self, runner: FakeRunner):
        buffer = io.StringIO()
        with redirect_stdout(buffer):
            return read_classroom50_state("org", "lop", runner=runner)

    def test_20_on_the_roster_but_not_in_the_org_is_said_so(self):
        runner = FakeRunner(roster=[roster_row(username="annguyen")], members=[])
        plans = classify_rows([resolved()], None, self.state(runner))
        self.assertEqual(plans[0].classroom50, C50_ON_ROSTER_NOT_MEMBER)

    def test_20b_on_the_roster_and_in_the_org_is_said_so(self):
        runner = FakeRunner(
            roster=[roster_row(username="annguyen")], members=[{"login": "AnNguyen"}]
        )
        plans = classify_rows([resolved()], None, self.state(runner))
        self.assertEqual(plans[0].classroom50, C50_ON_ROSTER_MEMBER)

    def test_21_a_failed_member_list_reads_as_unknown_not_as_present(self):
        runner = FakeRunner(roster=[roster_row(username="annguyen")], member_list_fails=True)
        state = self.state(runner)
        self.assertFalse(state.members_known)
        plans = classify_rows([resolved()], None, state)
        self.assertEqual(plans[0].classroom50, C50_ON_ROSTER_UNKNOWN_MEMBER)

    def test_29_an_email_only_roster_row_is_matched_instead_of_added_again(self):
        runner = FakeRunner(roster=[roster_row(email="An.Nguyen@vnu.edu.vn")], members=[])
        state = self.state(runner)
        self.assertEqual(state.roster_usernames, set())
        self.assertEqual(state.unlinked_emails, {"an.nguyen@vnu.edu.vn"})
        plans = classify_rows([resolved()], None, state)
        self.assertEqual(plans[0].classroom50, C50_UNLINKED_EMAIL)

    def test_29b_an_unrelated_row_is_still_a_new_addition(self):
        runner = FakeRunner(roster=[roster_row(email="someone.else@vnu.edu.vn")], members=[])
        plans = classify_rows([resolved()], None, self.state(runner))
        self.assertEqual(plans[0].classroom50, C50_WILL_ADD)

    def test_30_sync_exit_2_is_reported_as_pending_not_as_a_failure(self):
        runner = FakeRunner(sync_exit=2, sync_output="1 pending link")
        state = self.state(runner)
        self.assertEqual(state.sync_exit, 2)
        message = onboard._sync_message(state)
        self.assertIn("exit 2", message)
        self.assertIn("không phải lỗi", message)

    def test_31_sync_exit_1_is_reported_as_a_failure(self):
        state = self.state(FakeRunner(sync_exit=1))
        message = onboard._sync_message(state)
        self.assertIn("thất bại", message)
        self.assertNotIn("không phải lỗi", message)

    def test_31b_roster_sync_never_emits_write(self):
        from course_hoanganhduc.c50_cli_human import HumanCLI

        argv = HumanCLI(runner=FakeRunner()).roster_sync_argv("org", "lop")
        self.assertEqual(argv, ["gh", "teacher", "roster", "sync", "org", "lop"])
        self.assertNotIn("--write", argv)


# --- 22, 30, 31 end to end: the confirmation gate ----------------------------


class TestConfirmationGate(_OnboardCase):
    def full_run(self, runner: FakeRunner, fake: FakeCourse, **kwargs: Any):
        return self.run_driver(
            form_csv(form_row()),
            runner=runner,
            service=fake,
            students=[],
            google_course_id="C",
            c50_org="org",
            c50_classroom="lop",
            **kwargs,
        )

    def test_22_the_table_is_printed_before_anything_is_written(self):
        runner = FakeRunner(users={"annguyen": 9001})
        fake = FakeCourse()
        report, output = self.full_run(runner, fake, answers=["n"])
        self.assertIn("Sinh viên", output)
        self.assertIn(GC_WILL_INVITE, output)
        self.assertIn(C50_WILL_ADD, output)
        self.assertTrue(report.get("cancelled"))
        self.assertFalse(report["databaseSaved"])
        self.assertEqual(self.saved, [])
        self.assertEqual(fake.creates(), [])
        self.assertEqual(runner.ran("roster", "import"), [])

    def test_22b_a_confirmed_run_writes_in_the_documented_order(self):
        runner = FakeRunner(users={"annguyen": 9001})
        fake = FakeCourse()
        report, _ = self.full_run(runner, fake, answers=["y", "y"])
        self.assertTrue(report["databaseSaved"])
        self.assertEqual(report["c50"]["status"], "imported")
        self.assertEqual(len(fake.creates()), 1)
        self.assertEqual(len(runner.ran("roster", "import")), 1)

    def test_30b_sync_exit_2_stops_the_import_before_it_writes(self):
        runner = FakeRunner(users={"annguyen": 9001}, sync_exit=2)
        fake = FakeCourse()
        report, output = self.full_run(runner, fake, answers=["y", "y"])
        self.assertEqual(report["c50"]["status"], "sync_pending")
        self.assertEqual(report["c50"]["exitCode"], 2)
        self.assertEqual(runner.ran("roster", "import"), [])
        self.assertIn("--write", output)
        # Google Classroom is a separate lane and is not held back by it.
        self.assertEqual(len(fake.creates()), 1)

    def test_31c_sync_exit_1_stops_the_import_as_a_failure(self):
        runner = FakeRunner(users={"annguyen": 9001}, sync_exit=1)
        report, _ = self.full_run(runner, FakeCourse(), answers=["y", "y"])
        self.assertEqual(report["c50"]["status"], "sync_failed")
        self.assertEqual(runner.ran("roster", "import"), [])

    def test_17b_an_enrolled_student_is_not_invited_again(self):
        runner = FakeRunner(users={"annguyen": 9001})
        fake = FakeCourse(students=[enrolled("an.nguyen@vnu.edu.vn")])
        report, output = self.full_run(runner, fake, answers=["y", "y"])
        self.assertIn(GC_ENROLLED, output)
        self.assertEqual(fake.creates(), [])
        self.assertEqual(report["google"]["skipped_enrolled"], 1)


# --- 23-28: what actually reaches the Classroom50 roster ---------------------


class TestRosterExport(_OnboardCase):
    def import_rows(self, csv_text: str, *, runner: FakeRunner, students=None, **kwargs: Any):
        report, output = self.run_driver(
            csv_text,
            runner=runner,
            students=students if students is not None else [],
            google_course_id=None,
            c50_org="org",
            c50_classroom="lop",
            answers=kwargs.pop("answers", ["y", "y"]),
            **kwargs,
        )
        return runner.import_rows(), report, output

    def test_23_the_exported_csv_is_the_six_canonical_columns_only(self):
        runner = FakeRunner(users={"annguyen": 9001})
        rows, _, _ = self.import_rows(form_csv(form_row()), runner=runner)
        header = (runner.imported_csv or "").splitlines()[0].split(",")
        self.assertEqual(header, CANONICAL_COLUMNS)
        self.assertNotIn("student_id", header)
        # The student id is kept locally, where the two lanes are joined.
        self.assertEqual(self.saved[0]["rows"][0].__dict__["Student ID"], "22001234")

    def test_24_course_section_becomes_section_and_class_stays_local(self):
        runner = FakeRunner(users={"annguyen": 9001})
        rows, _, _ = self.import_rows(
            form_csv(form_row(class_name="K70A9", section="MAT3500 1")), runner=runner
        )
        self.assertEqual(rows[0]["section"], "MAT3500 1")
        self.assertNotIn("K70A9", runner.imported_csv or "")
        student = self.saved[0]["rows"][0]
        self.assertEqual(student.__dict__["Class"], "K70A9")
        self.assertEqual(student.__dict__["Section"], "MAT3500 1")

    def test_25_the_vietnamese_name_split_is_first_all_but_last_word(self):
        runner = FakeRunner(users={"annguyen": 9001})
        rows, _, _ = self.import_rows(
            form_csv(form_row(name="Nguyễn Văn A")), runner=runner
        )
        self.assertEqual(rows[0]["first_name"], "Nguyễn Văn")
        self.assertEqual(rows[0]["last_name"], "A")

    def test_26_a_stale_github_id_in_the_database_is_overwritten(self):
        runner = FakeRunner(users={"annguyen": 9001})
        stale = Student(
            **{
                "Student ID": "22001234",
                "Name": "Nguyễn Văn A",
                "GitHub Username": "annguyen",
                "GitHub ID": "1111",
            }
        )
        rows, _, _ = self.import_rows(
            form_csv(form_row()), runner=runner, students=[stale]
        )
        self.assertEqual(rows[0]["github_id"], "9001")
        self.assertNotIn("1111", runner.imported_csv or "")

    def test_27_an_unverified_account_exports_an_empty_github_id(self):
        runner = FakeRunner(users={"annguyen": "!rate limit exceeded"})
        rows, report, _ = self.import_rows(form_csv(form_row()), runner=runner)
        self.assertEqual(rows[0]["github_id"], "")
        self.assertEqual(rows[0]["username"], "annguyen")
        self.assertEqual(report["unverifiedGithub"], ["annguyen"])

    def test_28_an_empty_email_still_reaches_the_roster(self):
        runner = FakeRunner(users={"annguyen": 9001})
        rows, _, _ = self.import_rows(
            form_csv(form_row(school="", personal="")), runner=runner
        )
        self.assertEqual(rows[0]["username"], "annguyen")
        self.assertEqual(rows[0]["email"], "")

    def test_28b_a_row_without_a_username_is_kept_off_the_roster_and_said_so(self):
        runner = FakeRunner(users={"annguyen": 9001})
        rows, report, output = self.import_rows(
            form_csv(form_row(), form_row(student_id="2", name="Không GitHub", github="")),
            runner=runner,
        )
        self.assertEqual([row["username"] for row in rows], ["annguyen"])
        self.assertIn(C50_NO_USERNAME, output)
        # Kept in the database and still invitable by email: only the roster
        # needs a username.
        self.assertEqual(report["resolved"], 2)


# --- 35, 36: the header rule and the reused flags ----------------------------


# --- 38-44: settings a course was configured with once ----------------------


class TestPresets(_OnboardCase):
    """A configured course should not have to be re-picked, but not go unnoticed."""

    ENVIRONMENT = ("GOOGLE_CLASSROOM_COURSE_ID", "CLASSROOM50_ORG", "CLASSROOM50_CLASSROOM")

    def blank_args(self):
        import argparse

        # The environment is the third tier, so a variable left over from the
        # machine running these tests would answer for the config file.
        saved = {key: os.environ.pop(key, None) for key in self.ENVIRONMENT}
        self.addCleanup(
            lambda: [os.environ.__setitem__(k, v) for k, v in saved.items() if v is not None]
        )
        return argparse.Namespace(
            google_course_id=None, classroom50_org=None, classroom50_classroom=None
        )

    def asker(self, answers: List[str]):
        """An ``input_fn`` that records what it was asked.

        The driver puts its questions in the prompt, which ``input`` writes to
        the terminal and a stub does not, so the questions cannot be read off
        captured stdout the way the printed report can.
        """
        seen: List[str] = []
        queue = list(answers)

        def ask(prompt: str) -> str:
            seen.append(prompt)
            return queue.pop(0) if queue else "n"

        return ask, seen

    def confirmations(self, prompts: List[str]) -> List[str]:
        return [p for p in prompts if "dùng giá trị này?" in p]

    def preset_run(self, runner, service, **kwargs: Any):
        return self.run_driver(
            form_csv(form_row()), runner=runner, service=service, students=[], **kwargs
        )

    def test_38_the_config_file_answers_all_three_and_keeps_the_source(self):
        import argparse

        args = self.blank_args()
        cfg = {
            "GOOGLE_CLASSROOM_COURSE_ID": "cfg-course",
            "CLASSROOM50_ORG": "cfg-org",
            "CLASSROOM50_CLASSROOM": "cfg-room",
        }
        presets = onboard.resolve_presets(args, cfg)
        self.assertEqual(
            [p.value for p in presets], ["cfg-course", "cfg-org", "cfg-room"]
        )
        self.assertEqual([p.source for p in presets], [SOURCE_CONFIG] * 3)

        # A flag beats the file, and carries the authority a default does not.
        flagged = argparse.Namespace(
            google_course_id="flag-course",
            classroom50_org="flag-org",
            classroom50_classroom="flag-room",
        )
        self.assertEqual(
            [p.source for p in onboard.resolve_presets(flagged, cfg)], [SOURCE_FLAG] * 3
        )
        self.assertEqual([p.value for p in onboard.resolve_presets(args, None)], [None] * 3)

    def test_39_a_configured_value_is_shown_and_used_once_confirmed(self):
        runner = FakeRunner(users={"annguyen": 9001})
        # three defaults, the write gate, then roster-import's own gate
        ask, seen = self.asker(["y"] * 5)
        report, _ = self.preset_run(
            runner,
            FakeCourse(),
            google_course_id=Resolved("cfg-course", SOURCE_CONFIG),
            c50_org=Resolved("org", SOURCE_CONFIG),
            c50_classroom=Resolved("lop", SOURCE_CONFIG),
            input_fn=ask,
        )
        confirmed = self.confirmations(seen)
        self.assertEqual(len(confirmed), 3)
        for value in ("cfg-course", "org", "lop"):
            self.assertTrue(any(value in prompt for prompt in confirmed), value)
        self.assertEqual(report["googleCourseId"], "cfg-course")
        self.assertEqual(report["classroom50"], {"org": "org", "classroom": "lop"})
        self.assertTrue(report["databaseSaved"])
        self.assertEqual(report["c50"]["status"], "imported")

    def test_40_declining_a_stale_course_falls_through_to_the_picker(self):
        seen: List[Any] = []

        def picker(credentials_path, token_path, verbose=False, open_browser=None):
            seen.append((credentials_path, token_path, open_browser))
            return "picked-course"

        onboard.pick_google_course = picker  # tearDown puts the real one back
        report, _ = self.preset_run(
            FakeRunner(users={"annguyen": 9001}),
            FakeCourse(),
            google_course_id=Resolved("stale-course", SOURCE_CONFIG),
            c50_org=None,
            answers=["n", "y"],
        )
        self.assertEqual(len(seen), 1)
        # The picker authorizes on its way to listing the courses, so the browser
        # choice has to reach it too, not only the later credential call.
        self.assertIsNone(seen[0][2])
        self.assertEqual(report["googleCourseId"], "picked-course")

    def test_41_a_flag_is_an_instruction_and_is_not_asked_about(self):
        ask, seen = self.asker(["y", "y"])
        report, _ = self.preset_run(
            FakeRunner(users={"annguyen": 9001}),
            FakeCourse(),
            google_course_id="C",
            c50_org="org",
            c50_classroom="lop",
            input_fn=ask,
        )
        self.assertEqual(self.confirmations(seen), [])
        self.assertTrue(report["databaseSaved"])

    def test_42_a_declined_org_asks_for_another_and_reads_empty_as_skip(self):
        runner = FakeRunner(users={"annguyen": 9001})
        ask, seen = self.asker(["n", "", "y"])
        report, output = self.preset_run(
            runner,
            None,
            google_course_id=None,
            c50_org=Resolved("old-org", SOURCE_CONFIG),
            c50_classroom=Resolved("lop", SOURCE_CONFIG),
            input_fn=ask,
        )
        self.assertTrue(any("Org GitHub khác" in prompt for prompt in seen))
        self.assertIn("Bỏ qua Classroom50", output)
        self.assertEqual(report["classroom50"], {"org": "", "classroom": ""})
        self.assertEqual(runner.ran("teacher"), [])

    def test_42b_a_replaced_org_sends_the_configured_classroom_to_the_picker(self):
        asked: List[str] = []
        real = onboard.pick_c50_classroom
        onboard.pick_c50_classroom = lambda org, **kwargs: (asked.append(org), "lop-moi")[1]
        ask, seen = self.asker(["n", "org-moi", "y", "y"])
        try:
            report, _ = self.preset_run(
                FakeRunner(users={"annguyen": 9001}),
                None,
                google_course_id=None,
                c50_org=Resolved("old-org", SOURCE_CONFIG),
                c50_classroom=Resolved("lop-cu", SOURCE_CONFIG),
                input_fn=ask,
            )
        finally:
            onboard.pick_c50_classroom = real
        self.assertEqual(asked, ["org-moi"])
        # The old classroom belongs to the old org, so it is never offered.
        self.assertFalse(any("lop-cu" in prompt for prompt in seen))
        self.assertEqual(report["classroom50"], {"org": "org-moi", "classroom": "lop-moi"})

    def test_43_off_a_terminal_the_default_is_taken_and_said_so(self):
        ask, seen = self.asker(["y"])
        report, output = self.preset_run(
            FakeRunner(users={"annguyen": 9001}),
            FakeCourse(),
            google_course_id=Resolved("cfg-course", SOURCE_CONFIG),
            c50_org=None,
            tty_check=lambda: False,
            input_fn=ask,
        )
        self.assertEqual(self.confirmations(seen), [])
        self.assertIn("theo cấu hình sẵn có", output)
        self.assertEqual(report["googleCourseId"], "cfg-course")

    def test_44_the_onboard_lane_and_the_classroom50_flags_read_the_same_tiers(self):
        """One resolver, so ``--onboard`` and ``--classroom50-*`` cannot drift apart."""
        args = self.blank_args()
        for cfg in ({"CLASSROOM50_ORG": "o", "CLASSROOM50_CLASSROOM": "c"}, {}, None):
            presets = onboard.resolve_presets(args, cfg)
            self.assertEqual(presets.c50_org.value, _resolve_org(args, cfg))
            self.assertEqual(presets.c50_classroom.value, _resolve_classroom(args, cfg))


class TestHeaderRuleAndFlags(unittest.TestCase):
    def test_35_a_header_carrying_vnu_and_hus_is_read_as_the_school_address(self):
        header = "Mã Sinh Viên (Student ID),Họ và Tên (Full Name),Email VNU HUS,Email Address"
        parsed = parse_student_rows(
            header + "\n1,A,school@vnu.edu.vn,personal@gmail.com\n"
        )
        self.assertEqual(parsed.rows[0]["Email"], "school@vnu.edu.vn")
        self.assertEqual(parsed.rows[0]["Personal Email"], "personal@gmail.com")

    def test_35b_the_rule_does_not_swallow_an_unrelated_vnu_hus_question(self):
        from course_hoanganhduc.roster_csv import _map_headers

        mapping = _map_headers(["Mã Sinh Viên", "Họ và Tên", "Bạn học ở VNU-HUS?", "Email Address"])
        # The rule wants both needles; a yes/no question mentioning the school
        # carries neither an address nor the word "email".
        self.assertNotIn(2, mapping)
        self.assertEqual(mapping[3], "Email")

    def test_36_onboard_reuses_the_existing_classroom50_and_google_flags(self):
        import argparse

        from course_hoanganhduc.c50_flags import _resolve_classroom, _resolve_org

        args = argparse.Namespace(
            classroom50_org="from-flag", classroom50_classroom="from-flag-class"
        )
        self.assertEqual(_resolve_org(args, None), "from-flag")
        self.assertEqual(_resolve_classroom(args, None), "from-flag-class")

        env = argparse.Namespace(classroom50_org=None, classroom50_classroom=None)
        old = {k: os.environ.get(k) for k in ("CLASSROOM50_ORG", "CLASSROOM50_CLASSROOM")}
        os.environ["CLASSROOM50_ORG"] = "from-env"
        os.environ["CLASSROOM50_CLASSROOM"] = "from-env-class"
        try:
            self.assertEqual(_resolve_org(env, None), "from-env")
            self.assertEqual(_resolve_classroom(env, None), "from-env-class")
        finally:
            for key, value in old.items():
                if value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = value

    def test_36b_the_handler_passes_the_shared_flags_straight_through(self):
        source = Path(
            REPO_ROOT, "course_hoanganhduc", "core.py"
        ).read_text(encoding="utf-8")
        handler = source.split('if getattr(args, "onboard", False):', 1)[1][:1600]
        for expected in (
            "presets = resolve_presets(args, config)",
            "google_course_id=presets.google_course_id",
            "c50_org=presets.c50_org",
            "c50_classroom=presets.c50_classroom",
            'report_path=getattr(args, "classroom50_report", None)',
            'open_browser=False if getattr(args, "no_open_browser", False) else None',
        ):
            self.assertIn(expected, handler)
        # No parallel --onboard-org / --onboard-course-id to drift out of step.
        self.assertNotIn("--onboard-org", source)
        self.assertNotIn("--onboard-course-id", source)


if __name__ == "__main__":
    unittest.main(verbosity=2)
