#!/usr/bin/env python3
"""Offline tests for Google Classroom invitations.

The module under test builds its own service and sleeps between writes; both are
injectable, so nothing here touches the network or the clock.
"""

from __future__ import annotations

import io
import json
import os
import sys
import types
import unittest
from contextlib import redirect_stdout
from typing import Any, Dict, List, Optional

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)


class FakeHttpError(Exception):
    """Stands in for googleapiclient.errors.HttpError.

    Carries both halves the module reads: ``resp.status`` and the JSON body, whose
    ``error.status`` names the canonical code. Several Classroom conditions share
    HTTP 400, so the body is what tells them apart.
    """

    def __init__(
        self,
        status: int,
        message: str = "secret response body",
        canonical: str = "",
    ) -> None:
        super().__init__(f"HTTP {status}: {message}")
        self.resp = types.SimpleNamespace(status=status)
        body = {"error": {"code": status, "message": message}}
        if canonical:
            body["error"]["status"] = canonical
        self.content = json.dumps(body).encode("utf-8")


def _stub_heavy_deps() -> None:
    """Stub the third-party stack that gclass_auth and utils import at module scope.

    googleapiclient.errors is stubbed with FakeHttpError so the module under test
    catches exactly what the fakes raise. The database layer is imported lazily
    inside the function under test, so it is stubbed per-test instead.
    """

    def ensure(name, attrs=None):
        if name in sys.modules:
            return
        module = types.ModuleType(name)
        for key, value in (attrs or {}).items():
            setattr(module, key, value)
        sys.modules[name] = module

    ensure("googleapiclient")
    ensure("googleapiclient.discovery", {"build": lambda *a, **k: None})
    ensure("googleapiclient.errors", {"HttpError": FakeHttpError})
    ensure("google")
    ensure("google.oauth2")
    ensure("google.oauth2.credentials", {"Credentials": object})
    ensure("google.auth")
    ensure("google.auth.transport")
    ensure("google.auth.transport.requests", {"Request": object})
    ensure("google_auth_oauthlib")
    ensure("google_auth_oauthlib.flow", {"InstalledAppFlow": object})
    ensure("pandas")
    ensure("openpyxl")
    ensure("openpyxl.styles", {"Alignment": object})
    ensure("pytesseract")
    ensure("pdf2image", {"convert_from_path": lambda *a, **k: []})
    ensure("PIL", {"Image": object, "ImageOps": object, "ImageFilter": object})
    ensure("requests")
    ensure("PyPDF2")
    ensure("numpy")
    ensure("cv2")
    ensure("sklearn")
    ensure("sklearn.feature_extraction")
    ensure("sklearn.feature_extraction.text", {"TfidfVectorizer": object})
    ensure("sklearn.metrics")
    ensure("sklearn.metrics.pairwise", {"cosine_similarity": lambda *a, **k: None})
    ensure("tqdm", {"tqdm": lambda x, **k: x})
    ensure("canvasapi", {"Canvas": object})
    ensure("paddleocr", {"PaddleOCR": object})


_stub_heavy_deps()

from course_hoanganhduc import gclass_invite  # noqa: E402
from course_hoanganhduc.gclass_invite import (  # noqa: E402
    invite_students_to_google_classroom,
)
from course_hoanganhduc.models import Student  # noqa: E402


# --- fakes ------------------------------------------------------------------


class _Request:
    def __init__(self, fn):
        self._fn = fn

    def execute(self, **kwargs):
        if kwargs:
            raise AssertionError(f"unexpected execute kwargs: {kwargs}")
        return self._fn()


def _page(items: List[Dict[str, Any]], key: str, page_size: int, token: Optional[str]):
    start = int(token or 0)
    chunk = items[start : start + page_size]
    body: Dict[str, Any] = {key: chunk}
    if start + page_size < len(items):
        body["nextPageToken"] = str(start + page_size)
    return body


class FakeClassroom:
    """A course whose pending-invitation list grows as invitations are created."""

    def __init__(
        self,
        students: Optional[List[Dict[str, Any]]] = None,
        teachers: Optional[List[Dict[str, Any]]] = None,
        invitations: Optional[List[Dict[str, Any]]] = None,
        profiles: Optional[Dict[str, str]] = None,
        create_errors: Optional[Dict[str, List[Exception]]] = None,
    ) -> None:
        self.students = list(students or [])
        self.teachers = list(teachers or [])
        self.pending = list(invitations or [])
        self.profiles = dict(profiles or {})
        self.create_errors = {k: list(v) for k, v in (create_errors or {}).items()}
        self.calls: List[Any] = []
        self._next = 1

    # -- resources
    def courses(self):
        return types.SimpleNamespace(
            students=lambda: types.SimpleNamespace(list=self._students_list),
            teachers=lambda: types.SimpleNamespace(list=self._teachers_list),
        )

    def invitations(self):
        return types.SimpleNamespace(list=self._inv_list, create=self._inv_create)

    def userProfiles(self):
        return types.SimpleNamespace(get=self._profile_get)

    # -- handlers
    def _students_list(self, courseId, pageSize, pageToken=None):
        self.calls.append(("students.list", courseId, pageToken))
        return _Request(lambda: _page(self.students, "students", pageSize, pageToken))

    def _teachers_list(self, courseId, pageSize, pageToken=None):
        self.calls.append(("teachers.list", courseId, pageToken))
        return _Request(lambda: _page(self.teachers, "teachers", pageSize, pageToken))

    def _inv_list(self, courseId, pageSize, pageToken=None):
        self.calls.append(("invitations.list", courseId, pageToken))
        return _Request(
            lambda: _page(self.pending, "invitations", pageSize, pageToken)
        )

    def _profile_get(self, userId):
        self.calls.append(("userProfiles.get", userId))

        def run():
            if userId not in self.profiles:
                raise FakeHttpError(404)
            return {"id": userId, "emailAddress": self.profiles[userId]}

        return _Request(run)

    def _inv_create(self, body):
        self.calls.append(("invitations.create", dict(body)))
        email = body["userId"]

        def run():
            queue = self.create_errors.get(email)
            if queue:
                raise queue.pop(0)
            user_id = f"uid-{self._next}"
            invitation_id = f"inv-{self._next}"
            self._next += 1
            self.pending.append({"id": invitation_id, "userId": user_id})
            self.profiles[user_id] = email
            return {
                "id": invitation_id,
                "userId": user_id,
                "courseId": body["courseId"],
                "role": body["role"],
            }

        return _Request(run)

    # -- assertions helpers
    def creates(self) -> List[Dict[str, Any]]:
        return [c[1] for c in self.calls if c[0] == "invitations.create"]

    def profile_gets(self) -> List[str]:
        return [c[1] for c in self.calls if c[0] == "userProfiles.get"]


def enrolled(email: str, user_id: str = "") -> Dict[str, Any]:
    return {"userId": user_id or f"u-{email}", "profile": {"emailAddress": email}}


def student(name: str, email: str, **extra: Any) -> Student:
    return Student(Name=name, Email=email, **extra)


class _InviteCase(unittest.TestCase):
    """Runs the entry point with the database and the prompts stubbed out."""

    def setUp(self) -> None:
        self.sleeps: List[float] = []
        self.saved: List[Dict[str, Any]] = []
        self._real_data = sys.modules.get("course_hoanganhduc.data")
        self._real_input = gclass_invite.get_input_with_quit
        gclass_invite.get_input_with_quit = lambda *a, **k: "y"

    def tearDown(self) -> None:
        gclass_invite.get_input_with_quit = self._real_input
        if self._real_data is None:
            sys.modules.pop("course_hoanganhduc.data", None)
        else:
            sys.modules["course_hoanganhduc.data"] = self._real_data

    def install_db(self, students: List[Student]) -> None:
        module = types.ModuleType("course_hoanganhduc.data")
        module.load_database = lambda db_path, verbose=False: students
        module.save_database = lambda rows, db_path=None, verbose=False, audit_source=None: (
            self.saved.append(
                {"rows": rows, "db_path": db_path, "audit_source": audit_source}
            )
        )
        sys.modules["course_hoanganhduc.data"] = module

    def run_invite(self, fake: FakeClassroom, students: List[Student], **kwargs: Any):
        self.install_db(students)
        params = {
            "course_id": "C",
            "apply_all": True,
            "db_path": "students.db",
            "service": fake,
            "sleep_fn": self.sleeps.append,
        }
        params.update(kwargs)
        buffer = io.StringIO()
        with redirect_stdout(buffer):
            report = invite_students_to_google_classroom(**params)
        return report, buffer.getvalue()


# --- 1-5: idempotency -------------------------------------------------------


class TestAlreadyPresent(_InviteCase):
    def test_01_enrolled_match_is_case_insensitive(self):
        fake = FakeClassroom(students=[enrolled("Alice@X.EDU")])
        report, _ = self.run_invite(fake, [student("Alice", "alice@x.edu")])
        self.assertEqual(report["skipped_enrolled"], 1)
        self.assertEqual(report["invited"], 0)
        self.assertEqual(fake.creates(), [])

    def test_02_pending_matched_by_google_id_costs_no_profile_lookup(self):
        fake = FakeClassroom(invitations=[{"id": "i1", "userId": "123"}])
        report, _ = self.run_invite(
            fake, [student("Bob", "bob@x.edu", Google_ID="123")]
        )
        self.assertEqual(report["skipped_pending"], 1)
        self.assertEqual(fake.profile_gets(), [])
        self.assertEqual(fake.creates(), [])

    def test_03_pending_resolved_through_user_profile(self):
        fake = FakeClassroom(
            invitations=[{"id": "i1", "userId": "999"}],
            profiles={"999": "carol@x.edu"},
        )
        report, _ = self.run_invite(fake, [student("Carol", "carol@x.edu")])
        self.assertEqual(report["skipped_pending"], 1)
        self.assertEqual(fake.profile_gets(), ["999"])
        self.assertEqual(fake.creates(), [])

    def test_04_only_the_new_student_is_invited(self):
        fake = FakeClassroom(
            students=[enrolled("alice@x.edu")],
            invitations=[{"id": "i1", "userId": "123"}],
            profiles={"123": "bob@x.edu"},
        )
        report, _ = self.run_invite(
            fake,
            [
                student("Alice", "alice@x.edu"),
                student("Bob", "bob@x.edu"),
                student("Carol", "carol@x.edu"),
            ],
        )
        self.assertEqual(report["invited"], 1)
        self.assertEqual(
            fake.creates(),
            [{"courseId": "C", "userId": "carol@x.edu", "role": "STUDENT"}],
        )

    def test_05_second_run_is_a_no_op(self):
        fake = FakeClassroom()
        roster = [student("Carol", "carol@x.edu"), student("Dan", "dan@x.edu")]
        first, _ = self.run_invite(fake, roster)
        self.assertEqual(first["invited"], 2)
        before = len(fake.creates())
        second, _ = self.run_invite(fake, roster)
        self.assertEqual(second["invited"], 0)
        self.assertEqual(second["skipped_pending"], 2)
        self.assertEqual(len(fake.creates()), before)


# --- 6-10: error handling ---------------------------------------------------


class TestCreateErrors(_InviteCase):
    def test_06_conflict_is_pending_not_failure(self):
        fake = FakeClassroom(create_errors={"carol@x.edu": [FakeHttpError(409)]})
        report, _ = self.run_invite(fake, [student("Carol", "carol@x.edu")])
        self.assertEqual(report["failed"], 0)
        self.assertEqual(report["skipped_pending"], 1)

    def test_07_not_found_reports_a_missing_google_account(self):
        fake = FakeClassroom(create_errors={"carol@x.edu": [FakeHttpError(404)]})
        report, _ = self.run_invite(fake, [student("Carol", "carol@x.edu")])
        self.assertEqual(report["failed"], 1)
        detail = report["results"][0]["detail"]
        self.assertIn("no Google account", detail)
        # The API documents 404 as "the course or the user does not exist", so the
        # detail must not blame the student alone.
        self.assertIn("does not exist", detail)

    def test_07b_failed_precondition_is_a_skip_not_a_failure(self):
        # Documented cause: "the user already has this role or a role with greater
        # permissions". It arrives as HTTP 400, which 409 does not cover.
        fake = FakeClassroom(
            create_errors={
                "carol@x.edu": [
                    FakeHttpError(
                        400,
                        "User already has this role.",
                        canonical="FAILED_PRECONDITION",
                    )
                ]
            }
        )
        report, _ = self.run_invite(fake, [student("Carol", "carol@x.edu")])
        self.assertEqual(report["failed"], 0)
        self.assertEqual(report["skipped_precondition"], 1)
        # The API message is passed through, not matched on.
        self.assertEqual(report["results"][0]["detail"], "User already has this role.")

    def test_07c_plain_bad_request_is_still_a_failure(self):
        # A 400 without the canonical FAILED_PRECONDITION name is a real failure;
        # the skip must not swallow every 400.
        fake = FakeClassroom(create_errors={"carol@x.edu": [FakeHttpError(400, "bad")]})
        report, _ = self.run_invite(fake, [student("Carol", "carol@x.edu")])
        self.assertEqual(report["failed"], 1)
        self.assertEqual(report["skipped_precondition"], 0)

    def test_08_forbidden_stops_the_loop(self):
        fake = FakeClassroom(create_errors={"aaa@x.edu": [FakeHttpError(403)]})
        report, _ = self.run_invite(
            fake,
            [
                student("Aaa", "aaa@x.edu"),
                student("Bbb", "bbb@x.edu"),
                student("Ccc", "ccc@x.edu"),
            ],
        )
        self.assertEqual(len(fake.creates()), 1)
        self.assertEqual(report["failed"], 1)
        self.assertEqual(report["skipped_aborted"], 2)

    def test_09_transient_status_is_retried(self):
        fake = FakeClassroom(
            create_errors={"carol@x.edu": [FakeHttpError(429), FakeHttpError(503)]}
        )
        report, _ = self.run_invite(fake, [student("Carol", "carol@x.edu")])
        self.assertEqual(report["invited"], 1)
        self.assertEqual(len(fake.creates()), 3)
        self.assertEqual(self.sleeps, [1.0, 2.0])

    def test_09b_transient_status_gives_up_after_four_attempts(self):
        fake = FakeClassroom(
            create_errors={"carol@x.edu": [FakeHttpError(429) for _ in range(5)]}
        )
        report, _ = self.run_invite(fake, [student("Carol", "carol@x.edu")])
        self.assertEqual(report["failed"], 1)
        self.assertEqual(len(fake.creates()), 4)
        self.assertEqual(self.sleeps, [1.0, 2.0, 4.0])

    def test_10_permanent_status_is_never_retried(self):
        fake = FakeClassroom(
            create_errors={"carol@x.edu": [FakeHttpError(404) for _ in range(4)]}
        )
        self.run_invite(fake, [student("Carol", "carol@x.edu")])
        self.assertEqual(len(fake.creates()), 1)
        self.assertEqual(self.sleeps, [])


# --- 11-16: candidate selection --------------------------------------------


class TestCandidates(_InviteCase):
    def test_11_dry_run_sends_nothing(self):
        fake = FakeClassroom()
        report, _ = self.run_invite(
            fake, [student("Carol", "carol@x.edu")], dry_run=True
        )
        self.assertEqual(fake.creates(), [])
        self.assertEqual(report["invited"], 1)
        self.assertTrue(report["dry_run"])

    def test_12_student_without_email_is_reported(self):
        fake = FakeClassroom()
        report, _ = self.run_invite(
            fake, [student("Nobody", ""), student("Carol", "carol@x.edu")]
        )
        self.assertEqual(report["skipped_missing"], 1)
        names = [r["name"] for r in report["results"] if r["status"] == "skipped_missing"]
        self.assertEqual(names, ["Nobody"])

    def test_13_duplicate_email_is_collapsed(self):
        fake = FakeClassroom()
        report, _ = self.run_invite(
            fake,
            [student("Carol", "carol@x.edu"), student("Carol Again", "CAROL@x.edu")],
        )
        self.assertEqual(report["skipped_duplicate"], 1)
        self.assertEqual(len(fake.creates()), 1)

    def test_14_filters_narrow_the_candidate_set(self):
        roster = [
            student("Carol", "carol@x.edu", Section="s1", Class="K68"),
            student("Dan", "dan@y.edu", Section="s2", Class="K69"),
        ]
        for kwargs, expected in (
            ({}, ["carol@x.edu", "dan@y.edu"]),
            ({"domains": "x.edu"}, ["carol@x.edu"]),
            ({"domains": "@y.edu"}, ["dan@y.edu"]),
            ({"section": "s2"}, ["dan@y.edu"]),
            ({"class_name": "K68"}, ["carol@x.edu"]),
            ({"emails": "dan@y.edu"}, ["dan@y.edu"]),
            ({"domains": "x.edu", "section": "s2"}, []),
        ):
            with self.subTest(**kwargs):
                fake = FakeClassroom()
                self.run_invite(fake, roster, **kwargs)
                self.assertEqual(
                    sorted(c["userId"] for c in fake.creates()), sorted(expected)
                )

    def test_15_requested_email_absent_from_the_database(self):
        fake = FakeClassroom()
        report, _ = self.run_invite(
            fake, [student("Carol", "carol@x.edu")], emails="ghost@x.edu"
        )
        self.assertEqual(report["skipped_not_in_db"], 1)
        self.assertIn("--add-google-sheet", report["results"][0]["detail"])
        self.assertEqual(fake.creates(), [])

    def test_16_course_teacher_counts_as_already_present(self):
        fake = FakeClassroom(teachers=[enrolled("carol@x.edu")])
        report, _ = self.run_invite(fake, [student("Carol", "carol@x.edu")])
        self.assertEqual(report["skipped_enrolled"], 1)
        self.assertEqual(
            report["results"][0]["detail"], "already a teacher on this course"
        )
        self.assertEqual(fake.creates(), [])

    def test_16b_google_id_match_against_a_student_is_not_a_teacher(self):
        # The candidate's address is nowhere on the course; only the userId
        # matches, and it belongs to a student.
        fake = FakeClassroom(students=[enrolled("old-address@x.edu", user_id="u-42")])
        report, _ = self.run_invite(
            fake, [student("Carol", "carol@x.edu", Google_ID="u-42")]
        )
        self.assertEqual(report["skipped_enrolled"], 1)
        self.assertEqual(report["results"][0]["detail"], "already enrolled")
        self.assertEqual(fake.creates(), [])


# --- 17-21: write-back, paging, contract ------------------------------------


class TestWriteBackAndContract(_InviteCase):
    def test_17_database_write_back_only_fills_blanks(self):
        fake = FakeClassroom()
        rows = [
            student("Carol", "carol@x.edu"),
            student("Dan", "dan@x.edu", Google_ID="keep-me"),
        ]
        self.run_invite(fake, rows)
        self.assertEqual(len(self.saved), 1)
        self.assertEqual(self.saved[0]["audit_source"], "gclass-invite")
        by_email = {getattr(s, "Email"): s for s in self.saved[0]["rows"]}
        self.assertEqual(getattr(by_email["dan@x.edu"], "Google_ID"), "keep-me")
        self.assertTrue(getattr(by_email["carol@x.edu"], "Google_ID"))
        self.assertTrue(getattr(by_email["carol@x.edu"], "Google_Invitation_ID"))
        self.assertTrue(getattr(by_email["carol@x.edu"], "Google_Classroom_Invited_At"))

    def test_18_dry_run_does_not_touch_the_database(self):
        fake = FakeClassroom()
        self.run_invite(fake, [student("Carol", "carol@x.edu")], dry_run=True)
        self.assertEqual(self.saved, [])

    def test_18b_update_local_db_false_does_not_touch_the_database(self):
        fake = FakeClassroom()
        self.run_invite(
            fake, [student("Carol", "carol@x.edu")], update_local_db=False
        )
        self.assertEqual(self.saved, [])

    def test_19_every_page_is_read(self):
        fake = FakeClassroom(
            students=[enrolled(f"s{i}@x.edu") for i in range(250)],
            invitations=[{"id": f"i{i}", "userId": str(i)} for i in range(600)],
        )
        report, _ = self.run_invite(fake, [student("Carol", "carol@x.edu")])
        student_tokens = [c[2] for c in fake.calls if c[0] == "students.list"]
        invite_tokens = [c[2] for c in fake.calls if c[0] == "invitations.list"]
        self.assertEqual(student_tokens, [None, "200"])
        self.assertEqual(invite_tokens, [None, "500"])
        self.assertEqual(report["invited"], 1)

    def test_20_report_keys_are_the_documented_contract(self):
        fake = FakeClassroom()
        report, _ = self.run_invite(fake, [student("Carol", "carol@x.edu")])
        self.assertEqual(
            set(report),
            {
                "course_id",
                "results",
                "invited",
                "skipped_enrolled",
                "skipped_pending",
                "skipped_missing",
                "skipped_duplicate",
                "skipped_not_in_db",
                "skipped_aborted",
                "skipped_precondition",
                "failed",
                "dry_run",
            },
        )
        self.assertEqual(
            set(report["results"][0]),
            {"name", "email", "status", "detail", "user_id", "invitation_id"},
        )

    def test_21_profile_lookup_limit_is_respected(self):
        fake = FakeClassroom(
            invitations=[{"id": f"i{i}", "userId": f"u{i}"} for i in range(5)],
            profiles={f"u{i}": f"other{i}@x.edu" for i in range(5)},
        )
        self.run_invite(
            fake, [student("Carol", "carol@x.edu")], profile_lookup_limit=1
        )
        self.assertEqual(len(fake.profile_gets()), 1)

    def test_21b_google_id_still_matches_past_the_lookup_limit(self):
        # Reading the invitation list costs nothing extra once it is fetched, so
        # hitting the lookup cap must not stop the free Google_ID comparison.
        fake = FakeClassroom(
            invitations=[{"id": f"i{i}", "userId": f"u{i}"} for i in range(3)],
            profiles={f"u{i}": f"other{i}@x.edu" for i in range(3)},
        )
        report, _ = self.run_invite(
            fake,
            [student("Carol", "carol@x.edu", Google_ID="u2")],
            profile_lookup_limit=1,
        )
        self.assertEqual(report["skipped_pending"], 1)
        self.assertEqual(len(fake.profile_gets()), 1)
        self.assertEqual(fake.creates(), [])

    def test_21c_unreadable_database_after_the_invites_is_reported(self):
        # The invitations are already out, so this must warn rather than raise --
        # but staying silent would hide that the next run pays for the lookups again.
        fake = FakeClassroom()
        rows = [student("Carol", "carol@x.edu")]
        reads: List[str] = []

        def load_database(db_path, verbose=False):
            reads.append(db_path)
            if len(reads) > 1:
                raise OSError("database is locked")
            return rows

        module = types.ModuleType("course_hoanganhduc.data")
        module.load_database = load_database
        module.save_database = lambda *a, **k: self.saved.append(k)
        sys.modules["course_hoanganhduc.data"] = module

        buffer = io.StringIO()
        with redirect_stdout(buffer):
            report = invite_students_to_google_classroom(
                course_id="C",
                apply_all=True,
                db_path="students.db",
                service=fake,
                sleep_fn=self.sleeps.append,
            )
        self.assertEqual(report["invited"], 1)
        self.assertEqual(self.saved, [])
        output = buffer.getvalue()
        self.assertIn("database is locked", output)
        self.assertIn("Google_ID was not saved", output)


if __name__ == "__main__":
    unittest.main(verbosity=2)
