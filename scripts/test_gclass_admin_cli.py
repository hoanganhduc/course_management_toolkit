#!/usr/bin/env python3
"""Offline tests for the Google Classroom assignment administration CLI."""

from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from contextlib import contextmanager
from datetime import datetime, timezone
from io import StringIO
from pathlib import Path
from unittest import mock

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc.gclass_admin_cli import (  # noqa: E402
    AssignmentSpecError,
    build_assignment_operation_plan,
    load_assignment_spec,
    main,
)
from course_hoanganhduc.gclass_coursework import (  # noqa: E402
    GoogleClassroomOutcomeUnknownError,
    GoogleClassroomPartialCreateError,
)
from course_hoanganhduc.gclass_coursework_auth import (  # noqa: E402
    COURSEWORK_SCOPES,
    CourseworkAuthSession,
    CredentialSecurityError,
    TokenFileLock,
)


NOW = datetime(2026, 8, 28, 8, 0, tzinfo=timezone.utc)


def _has_git_ancestor(path):
    current = Path(path).absolute()
    while True:
        try:
            os.lstat(current / ".git")
            return True
        except FileNotFoundError:
            pass
        parent = current.parent
        if parent == current:
            return False
        current = parent


def _select_test_temp_parent():
    candidates = [Path(tempfile.gettempdir()), Path.home()]
    if os.name != "nt":
        candidates.append(Path("/var/tmp"))
    seen = set()
    for candidate in candidates:
        normalized = os.path.normcase(str(candidate.absolute()))
        if normalized in seen:
            continue
        seen.add(normalized)
        if (
            candidate.is_dir()
            and os.access(candidate, os.W_OK | os.X_OK)
            and not _has_git_ancestor(candidate)
        ):
            return candidate
    raise RuntimeError("No writable test temporary directory exists outside a Git worktree")


TEST_TEMP_PARENT = _select_test_temp_parent()


class FakeRequest:
    def __init__(self, owner, operation, kwargs, result=None, error=None):
        self.owner = owner
        self.operation = operation
        self.kwargs = kwargs
        self.result = result
        self.error = error

    def execute(self, *, num_retries):
        if num_retries != 0:
            raise AssertionError("Every Google request must use num_retries=0")
        self.owner.calls.append(
            (self.operation, self.kwargs, {"num_retries": num_retries})
        )
        if self.error is not None:
            raise self.error
        return self.result


class FakeCourseWork:
    def __init__(self, owner):
        self.owner = owner

    def create(self, **kwargs):
        value = self.owner.take_response("courseWork.create")
        return FakeRequest(self.owner, "courseWork.create", kwargs, result=value)

    def list(self, **kwargs):
        value = self.owner.take_response("courseWork.list")
        return FakeRequest(self.owner, "courseWork.list", kwargs, result=value)

    def get(self, **kwargs):
        value = self.owner.take_response("courseWork.get")
        return FakeRequest(self.owner, "courseWork.get", kwargs, result=value)


class FakeCourses:
    def __init__(self, owner):
        self.owner = owner

    def get(self, **kwargs):
        value = self.owner.take_response("courses.get")
        return FakeRequest(self.owner, "courses.get", kwargs, result=value)

    def courseWork(self):
        return FakeCourseWork(self.owner)


class FakeService:
    def __init__(self, responses):
        self.responses = dict(responses)
        self.calls = []

    def courses(self):
        return FakeCourses(self)

    def take_response(self, operation):
        value = self.responses.get(operation)
        if isinstance(value, tuple):
            if not value:
                raise AssertionError(f"Unexpected request: {operation}")
            result = value[0]
            remaining = value[1:]
            if remaining:
                self.responses[operation] = remaining
            else:
                self.responses.pop(operation, None)
            return result
        if operation not in self.responses:
            raise AssertionError(f"Unexpected request: {operation}")
        return self.responses.pop(operation)


class PromptAnswers:
    def __init__(self, answers):
        self.answers = iter(answers)
        self.prompts = []

    def __call__(self, prompt):
        self.prompts.append(prompt)
        return next(self.answers)


@contextmanager
def environment(**updates):
    previous = {key: os.environ.get(key) for key in updates}
    try:
        for key, value in updates.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value
        yield
    finally:
        for key, value in previous.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value


class SpecFixture(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(dir=str(TEST_TEMP_PARENT))
        self.root = Path(self.temporary.name)

    def tearDown(self):
        self.temporary.cleanup()

    def write_spec(self, value, name="assignment.json"):
        path = self.root / name
        path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
        return path

    def write_token_placeholder(self, name="token.json"):
        path = self.root / name
        path.write_text("{}", encoding="utf-8")
        if os.name != "nt":
            os.chmod(path, 0o600)
        return path


class TestAssignmentSpec(SpecFixture):
    def test_minimal_null_and_empty_materials_are_valid(self):
        for materials in (None, []):
            with self.subTest(materials=materials):
                path = self.write_spec(
                    {
                        "title": "No attachment",
                        "materials": materials,
                        "rubric": None,
                    },
                    name=f"spec-{materials is None}.json",
                )
                assignment, rubric = load_assignment_spec(path, now=NOW)
                self.assertNotIn("materials", assignment)
                self.assertIsNone(rubric)
                self.assertEqual(assignment["state"], "DRAFT")

    def test_full_spec_builds_every_stable_option(self):
        path = self.write_spec(
            {
                "title": "Full assignment",
                "description": "Details",
                "state": "DRAFT",
                "scheduled_at": "2026-09-01T15:00:00+07:00",
                "due_at": "2026-09-02T15:30:00+07:00",
                "max_points": 10,
                "materials": [
                    {"type": "drive_file", "file_id": "drive-1", "share_mode": "STUDENT_COPY"},
                    {"type": "link", "url": "https://example.edu/task"},
                    {"type": "youtube", "video_id": "video-1"},
                ],
                "assignee": {"mode": "INDIVIDUAL_STUDENTS", "student_ids": ["1", "2"]},
                "submission_modification_mode": "MODIFIABLE",
                "topic_id": "topic-1",
                "grading_period": {"mode": "EXPLICIT", "id": "period-1"},
                "rubric": {
                    "scoring_mode": "SCORED",
                    "criteria": [
                        {
                            "title": "Correctness",
                            "levels": [
                                {"title": "No", "points": 0},
                                {"title": "Yes", "points": 10},
                            ],
                        }
                    ],
                },
            }
        )
        assignment, rubric = load_assignment_spec(path, now=NOW)
        self.assertEqual(
            assignment,
            {
                "title": "Full assignment",
                "description": "Details",
                "workType": "ASSIGNMENT",
                "state": "DRAFT",
                "assigneeMode": "INDIVIDUAL_STUDENTS",
                "individualStudentsOptions": {"studentIds": ["1", "2"]},
                "submissionModificationMode": "MODIFIABLE",
                "scheduledTime": "2026-09-01T08:00:00Z",
                "dueDate": {"year": 2026, "month": 9, "day": 2},
                "dueTime": {"hours": 8, "minutes": 30, "seconds": 0},
                "maxPoints": 10,
                "materials": [
                    {
                        "driveFile": {
                            "driveFile": {"id": "drive-1"},
                            "shareMode": "STUDENT_COPY",
                        }
                    },
                    {"link": {"url": "https://example.edu/task"}},
                    {"youtubeVideo": {"id": "video-1"}},
                ],
                "topicId": "topic-1",
                "gradingPeriodId": "period-1",
            },
        )
        self.assertEqual(
            rubric,
            {
                "criteria": [
                    {
                        "title": "Correctness",
                        "levels": [
                            {"title": "No", "points": 0},
                            {"title": "Yes", "points": 10},
                        ],
                    }
                ]
            },
        )

    def test_shipped_assignment_samples_load(self):
        sample_root = Path(REPO_ROOT) / "sample" / "google_classroom"
        minimal, no_rubric = load_assignment_spec(
            sample_root / "assignment-minimal.sample.json", now=NOW
        )
        test_draft, no_test_rubric = load_assignment_spec(
            sample_root / "assignment-test-draft.sample.json", now=NOW
        )
        full, rubric = load_assignment_spec(
            sample_root / "assignment-full.sample.json", now=NOW
        )
        self.assertNotIn("materials", minimal)
        self.assertIsNone(no_rubric)
        self.assertEqual(test_draft["title"], "Classroom API test assignment")
        self.assertEqual(test_draft["state"], "DRAFT")
        self.assertNotIn("materials", test_draft)
        self.assertIsNone(no_test_rubric)
        self.assertEqual(len(full["materials"]), 3)
        self.assertEqual(full["scheduledTime"], "2099-09-01T01:00:00Z")
        self.assertEqual(rubric["criteria"][0]["levels"][-1]["points"], 10)

    def test_unknown_fields_and_bad_nested_shapes_are_rejected(self):
        cases = [
            {"title": "A", "unknown": True},
            {"title": "A", "materials": [{"type": "link", "url": "https://x.example", "extra": 1}]},
            {"title": "A", "materials": [{"type": "form", "url": "https://x.example"}]},
            {"title": "A", "assignee": {"mode": "ALL_STUDENTS", "student_ids": [], "extra": 1}},
            {"title": "A", "grading_period": {"mode": "AUTO", "extra": 1}},
            {"title": "A", "rubric": {"criteria": []}},
        ]
        for index, value in enumerate(cases):
            with self.subTest(value=value):
                with self.assertRaises((AssignmentSpecError, ValueError)):
                    load_assignment_spec(self.write_spec(value, f"bad-{index}.json"), now=NOW)

    def test_oversized_rubric_points_are_reported_as_spec_validation(self):
        path = self.write_spec(
            {
                "title": "A",
                "rubric": {
                    "scoring_mode": "SCORED",
                    "criteria": [
                        {
                            "levels": [
                                {"title": "Small", "points": 0},
                                {"title": "Huge", "points": 10**309},
                            ]
                        }
                    ],
                },
            }
        )
        with self.assertRaisesRegex(AssignmentSpecError, "finite and nonnegative"):
            load_assignment_spec(path, now=NOW)

    @unittest.skipIf(os.name == "nt", "O_NOFOLLOW behavior is POSIX-specific")
    def test_duplicate_json_symlink_and_oversized_spec_are_rejected(self):
        duplicate = self.root / "duplicate.json"
        duplicate.write_text('{"title":"A","title":"B"}', encoding="utf-8")
        with self.assertRaises(AssignmentSpecError):
            load_assignment_spec(duplicate, now=NOW)

        target = self.write_spec({"title": "A"}, "target.json")
        link = self.root / "link.json"
        link.symlink_to(target)
        with self.assertRaises(AssignmentSpecError):
            load_assignment_spec(link, now=NOW)

        oversized = self.root / "oversized.json"
        oversized.write_bytes(b" " * (1024 * 1024 + 1))
        with self.assertRaises(AssignmentSpecError):
            load_assignment_spec(oversized, now=NOW)

        fifo = self.root / "blocking.fifo"
        os.mkfifo(fifo)
        with self.assertRaises(AssignmentSpecError):
            load_assignment_spec(fifo, now=NOW)

    def test_operation_plan_is_deterministic_and_marks_stages(self):
        assignment, rubric = load_assignment_spec(
            self.write_spec(
                {
                    "title": "A",
                    "state": "PUBLISHED",
                    "rubric": {
                        "scoring_mode": "UNSCORED",
                        "criteria": [{"levels": [{"title": "Complete"}]}],
                    },
                }
            ),
            now=NOW,
        )
        first = build_assignment_operation_plan("d:course", assignment, rubric)
        second = build_assignment_operation_plan("d:course", assignment, rubric)
        self.assertEqual(first, second)
        self.assertEqual(
            first["operations"],
            ["create-draft", "create-rubric", "publish-assignment"],
        )
        live = build_assignment_operation_plan(
            "d:course", assignment, rubric, dry_run=False
        )
        self.assertEqual(first["executionMode"], "DRY_RUN")
        self.assertEqual(live["executionMode"], "LIVE")
        self.assertTrue(first["dryRun"])
        self.assertFalse(live["dryRun"])
        self.assertNotEqual(first["operationDigest"], live["operationDigest"])
        self.assertEqual(len(first["operationDigest"]), 64)


class TestAdminCLI(SpecFixture):
    def run_cli(self, argv, **kwargs):
        stdout = StringIO()
        stderr = StringIO()
        result = main(argv, stdout=stdout, stderr=stderr, now=NOW, **kwargs)
        return result, stdout.getvalue(), stderr.getvalue()

    def test_complete_loopback_uses_hidden_input_and_direct_local_request(self):
        callback = (
            "http://127.0.0.1:43821/?state=synthetic-state&"
            "code=synthetic-code&scope=email"
        )
        calls = []

        class Response:
            status = 200

            @staticmethod
            def read(limit):
                calls.append(("read", limit))
                return b"The authentication flow has completed."

        class Connection:
            def __init__(self, host, port, timeout):
                calls.append(("connect", host, port, timeout))

            def request(self, method, target, headers):
                calls.append(("request", method, target, headers))

            @staticmethod
            def getresponse():
                return Response()

            @staticmethod
            def close():
                calls.append(("close",))

        prompts = []

        def hidden_input(prompt):
            prompts.append(prompt)
            return callback

        result, output, errors = self.run_cli(
            ["complete-loopback", "--port", "43821"],
            tty_check=lambda: True,
            secret_input_fn=hidden_input,
            loopback_connection_factory=Connection,
        )
        self.assertEqual(result, 0, errors)
        self.assertEqual(
            json.loads(output),
            {"callbackDelivered": True, "port": 43821},
        )
        self.assertEqual(len(prompts), 1)
        self.assertNotIn(callback, output + errors + "".join(prompts))
        self.assertEqual(calls[0], ("connect", "127.0.0.1", 43821, 10))
        request = calls[1]
        self.assertEqual(request[0:2], ("request", "GET"))
        self.assertIn("state=synthetic-state", request[2])
        self.assertIn("code=synthetic-code", request[2])
        self.assertEqual(request[3]["Connection"], "close")
        self.assertEqual(calls[-1], ("close",))

    def test_complete_loopback_rejects_unsafe_or_malformed_urls_before_connecting(self):
        cases = {
            "https": "https://127.0.0.1:43821/?state=s&code=c",
            "hostname": "http://localhost:43821/?state=s&code=c",
            "port": "http://127.0.0.1:9/?state=s&code=c",
            "userinfo": "http://user@127.0.0.1:43821/?state=s&code=c",
            "path": "http://127.0.0.1:43821/callback?state=s&code=c",
            "fragment": "http://127.0.0.1:43821/?state=s&code=c#fragment",
            "missing-state": "http://127.0.0.1:43821/?code=c",
            "missing-code": "http://127.0.0.1:43821/?state=s",
            "duplicate-state": "http://127.0.0.1:43821/?state=s&state=t&code=c",
            "oauth-error": "http://127.0.0.1:43821/?state=s&error=denied&code=c",
            "newline": "http://127.0.0.1:43821/?state=s&code=c\nignored",
        }
        for label, callback in cases.items():
            with self.subTest(label=label):
                connections = []

                def forbidden_connection(*args, **kwargs):
                    connections.append((args, kwargs))
                    raise AssertionError("unsafe callback attempted a connection")

                result, output, errors = self.run_cli(
                    ["complete-loopback", "--port", "43821"],
                    tty_check=lambda: True,
                    secret_input_fn=lambda _prompt, value=callback: value,
                    loopback_connection_factory=forbidden_connection,
                )
                self.assertEqual(result, 2)
                self.assertEqual(output, "")
                self.assertEqual(connections, [])
                self.assertNotIn(callback, errors)

    def test_complete_loopback_requires_human_tty(self):
        calls = []

        def forbidden_input(prompt):
            calls.append(prompt)
            raise AssertionError("hidden input should not run")

        result, output, errors = self.run_cli(
            ["complete-loopback", "--port", "43821"],
            tty_check=lambda: False,
            secret_input_fn=forbidden_input,
        )
        self.assertEqual(result, 2)
        self.assertEqual(output, "")
        self.assertEqual(calls, [])
        self.assertIn("interactive terminal", errors)

    def test_complete_loopback_connection_failure_is_redacted_and_closed(self):
        callback = "http://127.0.0.1:43821/?state=secret-state&code=secret-code"
        calls = []

        class Connection:
            def __init__(self, host, port, timeout):
                calls.append(("connect", host, port, timeout))

            @staticmethod
            def request(method, target, headers):
                raise OSError("secret-code transport detail")

            @staticmethod
            def close():
                calls.append(("close",))

        result, output, errors = self.run_cli(
            ["complete-loopback", "--port", "43821"],
            tty_check=lambda: True,
            secret_input_fn=lambda _prompt: callback,
            loopback_connection_factory=Connection,
        )
        self.assertEqual(result, 2)
        self.assertEqual(output, "")
        self.assertNotIn("secret-code", errors)
        self.assertEqual(calls[-1], ("close",))

        with environment(COURSE_AGENT_MODE="1"):
            result, output, errors = self.run_cli(
                ["complete-loopback", "--port", "43821"],
                tty_check=lambda: True,
                secret_input_fn=lambda _prompt: callback,
                loopback_connection_factory=Connection,
            )
        self.assertEqual(result, 2)
        self.assertEqual(output, "")
        self.assertIn("agent", errors.lower())

    def test_dry_run_needs_no_account_credentials_tty_or_auth(self):
        path = self.write_spec({"title": "Offline", "materials": None})
        before = {
            item.relative_to(self.root): item.read_bytes()
            for item in self.root.rglob("*")
            if item.is_file()
        }

        def forbidden_auth(*args, **kwargs):
            raise AssertionError("dry-run attempted authentication")

        with environment(
            COURSE_GCLASS_CREDENTIALS="relative-invalid.json",
            COURSE_GCLASS_COURSEWORK_TOKEN="also-relative.json",
            COURSE_AGENT_MODE="1",
        ):
            with mock.patch(
                "socket.create_connection",
                side_effect=AssertionError("dry-run opened a socket"),
            ):
                result, output, errors = self.run_cli(
                    [
                        "create-assignment",
                        "--course-id",
                        "d:offline",
                        "--spec",
                        str(path),
                        "--dry-run",
                    ],
                    tty_check=lambda: False,
                    auth_factory=forbidden_auth,
                )
        self.assertEqual(result, 0, errors)
        rendered = json.loads(output)
        self.assertTrue(rendered["dryRun"])
        self.assertNotIn("materials", rendered["assignment"])
        self.assertEqual(rendered["operations"], ["create-assignment"])
        self.assertEqual(errors, "")
        after = {
            item.relative_to(self.root): item.read_bytes()
            for item in self.root.rglob("*")
            if item.is_file()
        }
        self.assertEqual(after, before)

    @unittest.skipIf(os.name == "nt", "POSIX credential status test")
    def test_auth_status_returns_nonzero_for_expired_access_only_token(self):
        token = self.root / "expired-token.json"
        token.write_text(
            json.dumps(
                {
                    "token": "expired-access",
                    "refresh_token": None,
                    "client_id": "abc.apps.googleusercontent.com",
                    "client_secret": "secret",
                    "scopes": list(COURSEWORK_SCOPES),
                    "granted_scopes": list(COURSEWORK_SCOPES),
                    "expiry": "2020-01-01T00:00:00Z",
                }
            ),
            encoding="utf-8",
        )
        os.chmod(token, 0o600)
        result, output, errors = self.run_cli(
            [
                "auth-status",
                "--account",
                "teacher@example.edu",
                "--token",
                str(token),
            ]
        )
        self.assertEqual(result, 1, errors)
        status = json.loads(output)
        self.assertTrue(status["token_safe"])
        self.assertFalse(status["token_usable"])

    @unittest.skipIf(os.name == "nt", "POSIX credential status test")
    def test_auth_status_fails_closed_for_extreme_expiry(self):
        token = self.root / "extreme-expiry-token.json"
        token.write_text(
            json.dumps(
                {
                    "token": "access",
                    "refresh_token": "refresh",
                    "client_id": "abc.apps.googleusercontent.com",
                    "client_secret": "secret",
                    "scopes": list(COURSEWORK_SCOPES),
                    "granted_scopes": list(COURSEWORK_SCOPES),
                    "expiry": "0001-01-01T00:00:00+14:00",
                }
            ),
            encoding="utf-8",
        )
        os.chmod(token, 0o600)
        result, output, errors = self.run_cli(
            [
                "auth-status",
                "--account",
                "teacher@example.edu",
                "--token",
                str(token),
            ]
        )
        self.assertEqual(result, 1, errors)
        self.assertEqual(errors, "")
        status = json.loads(output)
        self.assertFalse(status["token_safe"])
        self.assertFalse(status["token_usable"])

    def test_real_create_requires_account_and_confirmation_or_yes(self):
        path = self.write_spec({"title": "A"})
        base = ["create-assignment", "--course-id", "1", "--spec", str(path)]
        result, _, errors = self.run_cli(base, tty_check=lambda: True)
        self.assertEqual(result, 2)
        self.assertIn("--account", errors)

        result, _, errors = self.run_cli(
            base + ["--account", "teacher@example.edu"],
            tty_check=lambda: False,
        )
        self.assertEqual(result, 2)
        self.assertIn("interactive terminal", errors)

        with environment(COURSE_AGENT_MODE="1"):
            result, _, errors = self.run_cli(
                base + ["--account", "teacher@example.edu", "--yes"],
                tty_check=lambda: False,
            )
        self.assertEqual(result, 2)
        self.assertIn("agent", errors.lower())

    def test_declined_confirmation_stops_before_mutation(self):
        path = self.write_spec({"title": "A"})
        service = FakeService(
            {"courses.get": {"id": "1", "name": "Course", "courseState": "ACTIVE"}}
        )

        def fake_auth(paths, **kwargs):
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="fp",
            )

        answers = PromptAnswers(["n"])
        result, _, errors = self.run_cli(
            [
                "create-assignment",
                "--course-id",
                "1",
                "--spec",
                str(path),
                "--account",
                "teacher@example.edu",
            ],
            input_fn=answers,
            tty_check=lambda: True,
            auth_factory=fake_auth,
        )
        self.assertEqual(result, 2)
        self.assertEqual(
            service.calls,
            [
                (
                    "courses.get",
                    {"id": "1", "fields": "id,name,courseState"},
                    {"num_retries": 0},
                )
            ],
        )
        self.assertIn("cancelled", errors.lower())

    def test_course_lookup_missing_state_stops_before_mutation(self):
        path = self.write_spec({"title": "A"})
        token = self.write_token_placeholder()
        service = FakeService(
            {"courses.get": {"id": "1", "name": "Course without state"}}
        )

        def fake_auth(paths, **kwargs):
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="fp",
            )

        result, output, errors = self.run_cli(
            [
                "create-assignment",
                "--course-id",
                "1",
                "--spec",
                str(path),
                "--account",
                "teacher@example.edu",
                "--token",
                str(token),
                "--yes",
            ],
            tty_check=lambda: False,
            auth_factory=fake_auth,
        )
        self.assertEqual(result, 2)
        self.assertEqual(output, "")
        self.assertIn("state: (missing)", errors)
        self.assertEqual(
            service.calls,
            [
                (
                    "courses.get",
                    {"id": "1", "fields": "id,name,courseState"},
                    {"num_retries": 0},
                )
            ],
        )

    def test_yes_creates_safe_draft_without_tty_or_prompt(self):
        path = self.write_spec({"title": "Automated draft", "materials": None})
        token = self.write_token_placeholder()
        service = FakeService(
            {
                "courses.get": {"id": "1", "name": "Course", "courseState": "ACTIVE"},
                "courseWork.create": {"id": "cw-1", "courseId": "1", "state": "DRAFT"},
            }
        )
        auth_kwargs = []

        def fake_auth(paths, **kwargs):
            auth_kwargs.append(kwargs)
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="fp",
            )

        def forbidden_input(_prompt):
            raise AssertionError("--yes prompted for input")

        result, output, errors = self.run_cli(
            [
                "create-assignment",
                "--course-id",
                "1",
                "--spec",
                str(path),
                "--account",
                "teacher@example.edu",
                "--token",
                str(token),
                "--yes",
            ],
            input_fn=forbidden_input,
            tty_check=lambda: False,
            auth_factory=fake_auth,
        )
        self.assertEqual(result, 0, errors)
        self.assertTrue(json.loads(output)["created"])
        self.assertEqual(
            [call[0] for call in service.calls],
            ["courses.get", "courseWork.create"],
        )
        self.assertEqual(len(auth_kwargs), 1)
        self.assertFalse(auth_kwargs[0]["open_browser"])
        self.assertTrue(auth_kwargs[0]["require_existing_token"])

    def test_agent_safe_draft_creates_once_after_allowlisted_duplicate_preflight(self):
        path = self.write_spec(
            {
                "title": "Classroom API test assignment",
                "description": "Draft smoke test",
                "materials": None,
                "rubric": None,
            }
        )
        token = self.write_token_placeholder()
        assignment, rubric = load_assignment_spec(path, now=NOW)
        stored = dict(
            assignment,
            id="cw-1",
            courseId="123456789012",
            associatedWithDeveloper=True,
        )
        course = {
            "id": "123456789012",
            "name": "Pilot Course\u202e",
            "courseState": "ACTIVE",
        }
        service = FakeService(
            {
                "courses.get": (course, course),
                "courseWork.list": {"courseWork": []},
                "courseWork.create": stored,
                "courseWork.get": stored,
            }
        )
        auth_kwargs = []

        def fake_auth(paths, **kwargs):
            auth_kwargs.append(kwargs)
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="a" * 20,
            )

        with environment(
            COURSE_AGENT_MODE="1",
            GCLASS_ACCOUNT_ALLOWLIST="teacher@example.edu",
            GCLASS_COURSE_ALLOWLIST="123456789012",
        ):
            prepared_result, prepared_output, prepared_errors = self.run_cli(
                [
                    "prepare-agent-safe-draft",
                    "--course-id",
                    "123456789012",
                    "--spec",
                    str(path),
                    "--account",
                    "teacher@example.edu",
                    "--token",
                    str(token),
                ],
                tty_check=lambda: False,
                auth_factory=fake_auth,
            )
            self.assertEqual(prepared_result, 0, prepared_errors)
            prepared = json.loads(prepared_output)
            self.assertFalse(prepared["classroomMutation"])
            self.assertNotIn("\u202e", prepared_output)
            self.assertIn("\\u202e", prepared["courseName"])
            digest = prepared["approvalDigest"]
            result, output, errors = self.run_cli(
                [
                    "create-assignment",
                    "--course-id",
                    "123456789012",
                    "--spec",
                    str(path),
                    "--account",
                    "teacher@example.edu",
                    "--token",
                    str(token),
                    "--yes",
                    "--agent-safe-draft",
                    "--expect-approval-digest",
                    digest,
                ],
                tty_check=lambda: False,
                auth_factory=fake_auth,
            )
        self.assertEqual(result, 0, errors)
        receipt = json.loads(output)
        self.assertTrue(receipt["created"])
        self.assertTrue(receipt["readBackVerified"])
        self.assertEqual(receipt["courseWorkId"], "cw-1")
        self.assertEqual(len(auth_kwargs), 2)
        self.assertTrue(
            all(
                kwargs.get("open_browser") is False
                and kwargs.get("require_existing_token") is True
                for kwargs in auth_kwargs
            )
        )
        self.assertEqual(
            [call[0] for call in service.calls],
            [
                "courses.get",
                "courses.get",
                "courseWork.list",
                "courseWork.create",
                "courseWork.get",
            ],
        )
        self.assertEqual(
            service.calls[2][1],
            {
                "courseId": "123456789012",
                "courseWorkStates": ["DRAFT", "PUBLISHED", "DELETED"],
                "pageSize": 100,
                "fields": "nextPageToken,courseWork(id,title,state)",
            },
        )
        self.assertEqual(
            service.calls[3][1],
            {
                "courseId": "123456789012",
                "body": assignment,
                "fields": "id,courseId,state,associatedWithDeveloper",
            },
        )
        self.assertEqual(
            service.calls[4][1],
            {
                "courseId": "123456789012",
                "id": "cw-1",
                "fields": (
                    "id,courseId,title,description,workType,state,assigneeMode,"
                    "individualStudentsOptions,submissionModificationMode,"
                    "scheduledTime,dueDate,dueTime,maxPoints,materials,topicId,"
                    "gradingPeriodId,associatedWithDeveloper"
                ),
            },
        )

    def test_agent_safe_draft_reuses_one_identical_match_across_pages(self):
        path = self.write_spec({"title": "Smoke", "description": "Exact"})
        token = self.write_token_placeholder()
        assignment, _rubric = load_assignment_spec(path, now=NOW)
        stored = dict(
            assignment,
            id="cw-existing",
            courseId="1",
            associatedWithDeveloper=True,
        )
        course = {
            "id": "1",
            "name": "Pilot",
            "courseState": "ACTIVE",
        }
        service = FakeService(
            {
                "courses.get": (course, course),
                "courseWork.list": (
                    {
                        "courseWork": [
                            {
                                "id": "other",
                                "courseId": "1",
                                "title": "Other",
                                "state": "DRAFT",
                            }
                        ],
                        "nextPageToken": "page-2",
                    },
                    {"courseWork": [stored]},
                ),
                "courseWork.get": stored,
            }
        )

        def fake_auth(paths, **kwargs):
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="b" * 20,
            )

        with environment(
            COURSE_AGENT_MODE="1",
            GCLASS_ACCOUNT_ALLOWLIST="teacher@example.edu",
            GCLASS_COURSE_ALLOWLIST="1",
        ):
            prepared_result, prepared_output, prepared_errors = self.run_cli(
                [
                    "prepare-agent-safe-draft",
                    "--course-id",
                    "1",
                    "--spec",
                    str(path),
                    "--account",
                    "teacher@example.edu",
                    "--token",
                    str(token),
                ],
                tty_check=lambda: False,
                auth_factory=fake_auth,
            )
            self.assertEqual(prepared_result, 0, prepared_errors)
            digest = json.loads(prepared_output)["approvalDigest"]
            result, output, errors = self.run_cli(
                [
                    "create-assignment",
                    "--course-id",
                    "1",
                    "--spec",
                    str(path),
                    "--account",
                    "teacher@example.edu",
                    "--token",
                    str(token),
                    "--yes",
                    "--agent-safe-draft",
                    "--expect-approval-digest",
                    digest,
                ],
                tty_check=lambda: False,
                auth_factory=fake_auth,
            )
        self.assertEqual(result, 0, errors)
        receipt = json.loads(output)
        self.assertFalse(receipt["created"])
        self.assertTrue(receipt["reusedExisting"])
        self.assertTrue(receipt["readBackVerified"])
        self.assertEqual(receipt["courseWorkId"], "cw-existing")
        self.assertEqual(
            [call[0] for call in service.calls],
            [
                "courses.get",
                "courses.get",
                "courseWork.list",
                "courseWork.list",
                "courseWork.get",
            ],
        )
        self.assertEqual(service.calls[3][1]["pageToken"], "page-2")

    def test_agent_safe_draft_fails_closed_on_collision_and_policy_mismatch(self):
        path = self.write_spec({"title": "Smoke", "description": "Exact"})
        token = self.write_token_placeholder()
        assignment, _rubric = load_assignment_spec(path, now=NOW)

        policy_cases = [
            (
                "account-allowlist",
                "1",
                ["--expect-approval-digest", "0" * 64],
                {
                    "GCLASS_ACCOUNT_ALLOWLIST": "other@example.edu",
                    "GCLASS_COURSE_ALLOWLIST": "1",
                },
                "GCLASS_ACCOUNT_ALLOWLIST",
            ),
            (
                "course-allowlist",
                "1",
                ["--expect-approval-digest", "0" * 64],
                {
                    "GCLASS_ACCOUNT_ALLOWLIST": "teacher@example.edu",
                    "GCLASS_COURSE_ALLOWLIST": "2",
                },
                "GCLASS_COURSE_ALLOWLIST",
            ),
            (
                "alias",
                "d:1",
                ["--expect-approval-digest", "0" * 64],
                {
                    "GCLASS_ACCOUNT_ALLOWLIST": "teacher@example.edu",
                    "GCLASS_COURSE_ALLOWLIST": "d:1",
                },
                "canonical course ID",
            ),
        ]
        for label, course_id, extra, allowlists, expected in policy_cases:
            with self.subTest(label=label):
                auth_calls = []

                def forbidden_auth(*args, **kwargs):
                    auth_calls.append((args, kwargs))
                    raise AssertionError("policy mismatch reached authentication")

                with environment(COURSE_AGENT_MODE="1", **allowlists):
                    result, output, errors = self.run_cli(
                        [
                            "create-assignment",
                            "--course-id",
                            course_id,
                            "--spec",
                            str(path),
                            "--account",
                            "teacher@example.edu",
                            "--token",
                            str(token),
                            "--yes",
                            "--agent-safe-draft",
                        ]
                        + extra,
                        tty_check=lambda: False,
                        auth_factory=forbidden_auth,
                    )
                self.assertEqual(result, 2)
                self.assertEqual(output, "")
                self.assertEqual(auth_calls, [])
                self.assertIn(expected, errors)
                if label == "account-allowlist":
                    self.assertNotIn("teacher@example.edu", errors)

        unsafe_path = self.write_spec(
            {"title": "Smoke", "max_points": 1}, "unsafe-agent.json"
        )
        auth_calls = []
        with environment(
            COURSE_AGENT_MODE="1",
            GCLASS_ACCOUNT_ALLOWLIST="teacher@example.edu",
            GCLASS_COURSE_ALLOWLIST="1",
        ):
            result, output, errors = self.run_cli(
                [
                    "create-assignment",
                    "--course-id",
                    "1",
                    "--spec",
                    str(unsafe_path),
                    "--account",
                    "teacher@example.edu",
                    "--token",
                    str(token),
                    "--yes",
                    "--agent-safe-draft",
                    "--expect-approval-digest",
                    "0" * 64,
                ],
                tty_check=lambda: False,
                auth_factory=lambda *args, **kwargs: auth_calls.append((args, kwargs)),
            )
        self.assertEqual(result, 2)
        self.assertEqual(output, "")
        self.assertEqual(auth_calls, [])
        self.assertIn("maxPoints", errors)

        digest_service = FakeService(
            {
                "courses.get": {
                    "id": "1",
                    "name": "Pilot",
                    "courseState": "ACTIVE",
                }
            }
        )

        def digest_auth(paths, **kwargs):
            return CourseworkAuthSession(
                service=digest_service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="c" * 20,
            )

        with environment(
            COURSE_AGENT_MODE="1",
            GCLASS_ACCOUNT_ALLOWLIST="teacher@example.edu",
            GCLASS_COURSE_ALLOWLIST="1",
        ):
            result, output, errors = self.run_cli(
                [
                    "create-assignment",
                    "--course-id",
                    "1",
                    "--spec",
                    str(path),
                    "--account",
                    "teacher@example.edu",
                    "--token",
                    str(token),
                    "--yes",
                    "--agent-safe-draft",
                    "--expect-approval-digest",
                    "0" * 64,
                ],
                tty_check=lambda: False,
                auth_factory=digest_auth,
            )
        self.assertEqual(result, 2)
        self.assertEqual(output, "")
        self.assertIn("approval digest", errors)
        self.assertEqual(
            [call[0] for call in digest_service.calls], ["courses.get"]
        )

        stored = dict(
            assignment,
            id="cw-existing",
            courseId="1",
            associatedWithDeveloper=False,
            materials=[{"link": {"url": "https://example.edu/unexpected"}}],
        )
        course = {"id": "1", "name": "Pilot", "courseState": "ACTIVE"}
        service = FakeService(
            {
                "courses.get": (course, course),
                "courseWork.list": {"courseWork": [stored]},
                "courseWork.get": stored,
            }
        )

        def fake_auth(paths, **kwargs):
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="d" * 20,
            )

        with environment(
            COURSE_AGENT_MODE="1",
            GCLASS_ACCOUNT_ALLOWLIST="teacher@example.edu",
            GCLASS_COURSE_ALLOWLIST="1",
        ):
            prepared_result, prepared_output, prepared_errors = self.run_cli(
                [
                    "prepare-agent-safe-draft",
                    "--course-id",
                    "1",
                    "--spec",
                    str(path),
                    "--account",
                    "teacher@example.edu",
                    "--token",
                    str(token),
                ],
                tty_check=lambda: False,
                auth_factory=fake_auth,
            )
            self.assertEqual(prepared_result, 0, prepared_errors)
            digest = json.loads(prepared_output)["approvalDigest"]
            result, output, errors = self.run_cli(
                [
                    "create-assignment",
                    "--course-id",
                    "1",
                    "--spec",
                    str(path),
                    "--account",
                    "teacher@example.edu",
                    "--token",
                    str(token),
                    "--yes",
                    "--agent-safe-draft",
                    "--expect-approval-digest",
                    digest,
                ],
                tty_check=lambda: False,
                auth_factory=fake_auth,
            )
        self.assertEqual(result, 2)
        self.assertEqual(output, "")
        self.assertIn("same-title draft does not match", errors.lower())
        self.assertIn("associatedWithDeveloper", errors)
        self.assertIn("materials", errors)
        self.assertNotIn("courseWork.create", [call[0] for call in service.calls])

    def test_agent_safe_draft_blocks_same_title_published_or_deleted_replay(self):
        path = self.write_spec({"title": "One shot", "description": "Exact"})
        token = self.write_token_placeholder()
        course = {"id": "1", "name": "Pilot", "courseState": "ACTIVE"}
        for state in ("PUBLISHED", "DELETED"):
            with self.subTest(state=state):
                service = FakeService(
                    {
                        "courses.get": (course, course),
                        "courseWork.list": {
                            "courseWork": [
                                {
                                    "id": "cw-old",
                                    "courseId": "1",
                                    "title": "One shot",
                                    "state": state,
                                }
                            ]
                        },
                    }
                )

                def fake_auth(paths, **kwargs):
                    return CourseworkAuthSession(
                        service=service,
                        profile={
                            "id": "u",
                            "emailAddress": "teacher@example.edu",
                        },
                        paths=paths,
                        client_id_fingerprint="e" * 20,
                    )

                with environment(
                    COURSE_AGENT_MODE="1",
                    GCLASS_ACCOUNT_ALLOWLIST="teacher@example.edu",
                    GCLASS_COURSE_ALLOWLIST="1",
                ):
                    prepared_result, prepared_output, prepared_errors = self.run_cli(
                        [
                            "prepare-agent-safe-draft",
                            "--course-id",
                            "1",
                            "--spec",
                            str(path),
                            "--account",
                            "teacher@example.edu",
                            "--token",
                            str(token),
                        ],
                        tty_check=lambda: False,
                        auth_factory=fake_auth,
                    )
                    self.assertEqual(prepared_result, 0, prepared_errors)
                    digest = json.loads(prepared_output)["approvalDigest"]
                    result, output, errors = self.run_cli(
                        [
                            "create-assignment",
                            "--course-id",
                            "1",
                            "--spec",
                            str(path),
                            "--account",
                            "teacher@example.edu",
                            "--token",
                            str(token),
                            "--yes",
                            "--agent-safe-draft",
                            "--expect-approval-digest",
                            digest,
                        ],
                        tty_check=lambda: False,
                        auth_factory=fake_auth,
                    )
                self.assertEqual(result, 2)
                self.assertEqual(output, "")
                self.assertIn("outside DRAFT", errors)
                self.assertNotIn(
                    "courseWork.create", [call[0] for call in service.calls]
                )

    def test_agent_safe_draft_readback_failure_is_partial_and_never_retried(self):
        path = self.write_spec({"title": "Read back", "description": "Exact"})
        token = self.write_token_placeholder()
        assignment, _rubric = load_assignment_spec(path, now=NOW)
        course = {"id": "1", "name": "Pilot", "courseState": "ACTIVE"}
        created = dict(
            assignment,
            id="cw-created",
            courseId="1",
            associatedWithDeveloper=True,
        )
        service = FakeService(
            {
                "courses.get": (course, course),
                "courseWork.list": {"courseWork": []},
                "courseWork.create": created,
                "courseWork.get": "invalid-read-back",
            }
        )

        def fake_auth(paths, **kwargs):
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="f" * 20,
            )

        with environment(
            COURSE_AGENT_MODE="1",
            GCLASS_ACCOUNT_ALLOWLIST="teacher@example.edu",
            GCLASS_COURSE_ALLOWLIST="1",
        ):
            prepared_result, prepared_output, prepared_errors = self.run_cli(
                [
                    "prepare-agent-safe-draft",
                    "--course-id",
                    "1",
                    "--spec",
                    str(path),
                    "--account",
                    "teacher@example.edu",
                    "--token",
                    str(token),
                ],
                tty_check=lambda: False,
                auth_factory=fake_auth,
            )
            self.assertEqual(prepared_result, 0, prepared_errors)
            digest = json.loads(prepared_output)["approvalDigest"]
            result, output, errors = self.run_cli(
                [
                    "create-assignment",
                    "--course-id",
                    "1",
                    "--spec",
                    str(path),
                    "--account",
                    "teacher@example.edu",
                    "--token",
                    str(token),
                    "--yes",
                    "--agent-safe-draft",
                    "--expect-approval-digest",
                    digest,
                ],
                tty_check=lambda: False,
                auth_factory=fake_auth,
            )
        self.assertEqual(result, 4)
        self.assertEqual(output, "")
        error = json.loads(errors)
        self.assertEqual(error["error"], "partial_create")
        self.assertEqual(error["stage"], "read-back")
        self.assertEqual(error["courseWorkId"], "cw-created")
        self.assertEqual(
            [call[0] for call in service.calls].count("courseWork.create"), 1
        )
        self.assertEqual(
            [call[0] for call in service.calls].count("courseWork.get"), 1
        )

    def test_agent_safe_draft_approval_detects_spec_drift_before_preflight(self):
        path = self.write_spec({"title": "Frozen", "description": "Version one"})
        token = self.write_token_placeholder()
        course = {"id": "1", "name": "Pilot", "courseState": "ACTIVE"}
        service = FakeService({"courses.get": (course, course)})

        def fake_auth(paths, **kwargs):
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="1" * 20,
            )

        with environment(
            COURSE_AGENT_MODE="1",
            GCLASS_ACCOUNT_ALLOWLIST="teacher@example.edu",
            GCLASS_COURSE_ALLOWLIST="1",
        ):
            prepared_result, prepared_output, prepared_errors = self.run_cli(
                [
                    "prepare-agent-safe-draft",
                    "--course-id",
                    "1",
                    "--spec",
                    str(path),
                    "--account",
                    "teacher@example.edu",
                    "--token",
                    str(token),
                ],
                tty_check=lambda: False,
                auth_factory=fake_auth,
            )
            self.assertEqual(prepared_result, 0, prepared_errors)
            digest = json.loads(prepared_output)["approvalDigest"]
            path.write_text(
                json.dumps(
                    {"title": "Frozen", "description": "Version two"}
                ),
                encoding="utf-8",
            )
            result, output, errors = self.run_cli(
                [
                    "create-assignment",
                    "--course-id",
                    "1",
                    "--spec",
                    str(path),
                    "--account",
                    "teacher@example.edu",
                    "--token",
                    str(token),
                    "--yes",
                    "--agent-safe-draft",
                    "--expect-approval-digest",
                    digest,
                ],
                tty_check=lambda: False,
                auth_factory=fake_auth,
            )
        self.assertEqual(result, 2)
        self.assertEqual(output, "")
        self.assertIn("approval digest", errors)
        self.assertEqual(
            [call[0] for call in service.calls], ["courses.get", "courses.get"]
        )

    def test_agent_safe_draft_operation_lock_blocks_concurrent_local_run(self):
        path = self.write_spec({"title": "Locked", "description": "Exact"})
        token = self.write_token_placeholder()
        course = {"id": "1", "name": "Pilot", "courseState": "ACTIVE"}
        service = FakeService({"courses.get": (course, course)})

        def fake_auth(paths, **kwargs):
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="2" * 20,
            )

        with environment(
            COURSE_AGENT_MODE="1",
            GCLASS_ACCOUNT_ALLOWLIST="teacher@example.edu",
            GCLASS_COURSE_ALLOWLIST="1",
        ):
            prepared_result, prepared_output, prepared_errors = self.run_cli(
                [
                    "prepare-agent-safe-draft",
                    "--course-id",
                    "1",
                    "--spec",
                    str(path),
                    "--account",
                    "teacher@example.edu",
                    "--token",
                    str(token),
                ],
                tty_check=lambda: False,
                auth_factory=fake_auth,
            )
            self.assertEqual(prepared_result, 0, prepared_errors)
            digest = json.loads(prepared_output)["approvalDigest"]
            with TokenFileLock(token):
                result, output, errors = self.run_cli(
                    [
                        "create-assignment",
                        "--course-id",
                        "1",
                        "--spec",
                        str(path),
                        "--account",
                        "teacher@example.edu",
                        "--token",
                        str(token),
                        "--yes",
                        "--agent-safe-draft",
                        "--expect-approval-digest",
                        digest,
                    ],
                    tty_check=lambda: False,
                    auth_factory=fake_auth,
                )
        self.assertEqual(result, 2)
        self.assertEqual(output, "")
        self.assertIn("already in use", errors)
        self.assertEqual(
            [call[0] for call in service.calls], ["courses.get", "courses.get"]
        )

    def test_yes_requires_preexisting_token_before_authentication(self):
        path = self.write_spec({"title": "Automated draft"})
        calls = []

        def forbidden_auth(*args, **kwargs):
            calls.append((args, kwargs))
            raise AssertionError("--yes attempted fresh authorization")

        result, output, errors = self.run_cli(
            [
                "create-assignment",
                "--course-id",
                "1",
                "--spec",
                str(path),
                "--account",
                "teacher@example.edu",
                "--token",
                str(self.root / "missing-token.json"),
                "--yes",
            ],
            tty_check=lambda: False,
            auth_factory=forbidden_auth,
        )
        self.assertEqual(result, 2)
        self.assertEqual(output, "")
        self.assertEqual(calls, [])
        self.assertIn("existing coursework token", errors)

    def test_yes_rejects_release_and_drive_sharing_before_authentication(self):
        cases = [
            ("published", {"title": "A", "state": "PUBLISHED"}, "DRAFT"),
            (
                "scheduled",
                {"title": "A", "scheduled_at": "2026-09-01T08:00:00Z"},
                "DRAFT",
            ),
            (
                "drive",
                {
                    "title": "A",
                    "materials": [
                        {"type": "drive_file", "file_id": "file-1", "share_mode": "VIEW"}
                    ],
                },
                "Drive",
            ),
        ]
        for label, spec, expected_error in cases:
            with self.subTest(label=label):
                path = self.write_spec(spec, f"{label}.json")
                calls = []

                def forbidden_auth(*args, **kwargs):
                    calls.append((args, kwargs))
                    raise AssertionError("unsafe --yes plan attempted authentication")

                result, output, errors = self.run_cli(
                    [
                        "create-assignment",
                        "--course-id",
                        "1",
                        "--spec",
                        str(path),
                        "--account",
                        "teacher@example.edu",
                        "--yes",
                    ],
                    tty_check=lambda: False,
                    auth_factory=forbidden_auth,
                )
                self.assertEqual(result, 2)
                self.assertEqual(output, "")
                self.assertEqual(calls, [])
                self.assertIn(expected_error, errors)

    def test_real_create_verifies_canonical_course_and_uses_zero_retries(self):
        path = self.write_spec({"title": "A"})
        service = FakeService(
            {
                "courses.get": {"id": "999", "name": "Algorithms", "courseState": "ACTIVE"},
                "courseWork.create": {"id": "cw-1", "courseId": "999", "state": "DRAFT"},
            }
        )
        auth_calls = []

        def fake_auth(paths, **kwargs):
            auth_calls.append((paths, kwargs))
            return CourseworkAuthSession(
                service=service,
                profile={"id": "user-1", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="client-fingerprint",
            )

        answers = PromptAnswers(["yes"])
        result, output, errors = self.run_cli(
            [
                "create-assignment",
                "--course-id",
                "d:algorithms",
                "--spec",
                str(path),
                "--account",
                "teacher@example.edu",
                "--token",
                str(self.root / "token.json"),
            ],
            input_fn=answers,
            tty_check=lambda: True,
            auth_factory=fake_auth,
        )
        self.assertEqual(result, 0, errors)
        self.assertEqual(len(auth_calls), 1)
        self.assertEqual(
            service.calls,
            [
                (
                    "courses.get",
                    {"id": "d:algorithms", "fields": "id,name,courseState"},
                    {"num_retries": 0},
                ),
                (
                    "courseWork.create",
                    {
                        "courseId": "999",
                        "body": {
                            "title": "A",
                            "workType": "ASSIGNMENT",
                            "state": "DRAFT",
                            "assigneeMode": "ALL_STUDENTS",
                            "submissionModificationMode": "MODIFIABLE_UNTIL_TURNED_IN",
                        },
                    },
                    {"num_retries": 0},
                ),
            ],
        )
        receipt = json.loads(output)
        self.assertEqual(receipt["courseId"], "999")
        self.assertEqual(receipt["courseWorkId"], "cw-1")
        self.assertNotIn("teacher@example.edu", output)
        self.assertNotIn("title", receipt)
        self.assertEqual(len(answers.prompts), 1)
        prompt = answers.prompts[0]
        self.assertIn("Account: t***@example.edu", prompt)
        self.assertIn("Course: Algorithms (999)", prompt)
        self.assertIn("Title: A", prompt)
        self.assertIn("Release: DRAFT", prompt)
        self.assertIn("Create assignment? [y/N]", prompt)
        self.assertNotIn("operationDigest", prompt)

    def test_confirmation_summary_escapes_terminal_control_characters(self):
        path = self.write_spec({"title": "Normal title"})
        service = FakeService(
            {
                "courses.get": {
                    "id": "1",
                    "name": "Course\x1b[2J\nforged",
                    "courseState": "ACTIVE",
                }
            }
        )

        def fake_auth(paths, **kwargs):
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="fp",
            )

        answers = PromptAnswers(["n"])
        result, _, errors = self.run_cli(
            [
                "create-assignment",
                "--course-id",
                "1",
                "--spec",
                str(path),
                "--account",
                "teacher@example.edu",
            ],
            input_fn=answers,
            tty_check=lambda: True,
            auth_factory=fake_auth,
        )
        self.assertEqual(result, 2)
        self.assertIn("cancelled", errors.lower())
        self.assertEqual(len(answers.prompts), 1)
        self.assertNotIn("\x1b", answers.prompts[0])
        self.assertIn("\\u001b", answers.prompts[0])
        self.assertIn("\\u000a", answers.prompts[0])

    def test_interactive_create_supports_published_and_scheduled_release(self):
        cases = [
            (
                "published",
                {"title": "Publish", "state": "PUBLISHED"},
                {"id": "cw-1", "courseId": "1", "state": "PUBLISHED"},
                "PUBLISHED",
            ),
            (
                "scheduled",
                {"title": "Schedule", "scheduled_at": "2026-09-01T08:00:00Z"},
                {
                    "id": "cw-1",
                    "courseId": "1",
                    "state": "DRAFT",
                    "scheduledTime": "2026-09-01T08:00:00Z",
                },
                "SCHEDULED",
            ),
        ]
        for label, spec, response, expected_release in cases:
            with self.subTest(label=label):
                path = self.write_spec(spec, f"interactive-{label}.json")
                service = FakeService(
                    {
                        "courses.get": {
                            "id": "1",
                            "name": "Course",
                            "courseState": "ACTIVE",
                        },
                        "courseWork.create": response,
                    }
                )

                def fake_auth(paths, **kwargs):
                    return CourseworkAuthSession(
                        service=service,
                        profile={"id": "u", "emailAddress": "teacher@example.edu"},
                        paths=paths,
                        client_id_fingerprint="fp",
                    )

                answers = PromptAnswers(["yes"])
                result, output, errors = self.run_cli(
                    [
                        "create-assignment",
                        "--course-id",
                        "1",
                        "--spec",
                        str(path),
                        "--account",
                        "teacher@example.edu",
                        "--token",
                        str(self.root / "token.json"),
                    ],
                    input_fn=answers,
                    tty_check=lambda: True,
                    auth_factory=fake_auth,
                )
                self.assertEqual(result, 0, errors)
                self.assertEqual(json.loads(output)["releaseStatus"], expected_release)
                self.assertEqual(len(answers.prompts), 1)
                self.assertIn(f"Release: {expected_release}", answers.prompts[0])
                self.assertEqual(
                    [call[0] for call in service.calls],
                    ["courses.get", "courseWork.create"],
                )

    def test_spec_is_frozen_before_confirmation(self):
        path = self.write_spec({"title": "Confirmed body"})
        service = FakeService(
            {
                "courses.get": {"id": "1", "name": "Course", "courseState": "ACTIVE"},
                "courseWork.create": {"id": "cw-1", "courseId": "1", "state": "DRAFT"},
            }
        )

        def fake_auth(paths, **kwargs):
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="fp",
            )

        answers = iter(["yes"])

        def swapping_input(prompt):
            path.write_text(json.dumps({"title": "Changed body"}), encoding="utf-8")
            return next(answers)

        result, _, errors = self.run_cli(
            [
                "create-assignment",
                "--course-id",
                "1",
                "--spec",
                str(path),
                "--account",
                "teacher@example.edu",
                "--token",
                str(self.root / "token.json"),
            ],
            input_fn=swapping_input,
            tty_check=lambda: True,
            auth_factory=fake_auth,
        )
        self.assertEqual(result, 0, errors)
        self.assertEqual(
            service.calls[-1][1]["body"]["title"], "Confirmed body"
        )

    def test_drive_material_uses_one_readable_confirmation(self):
        path = self.write_spec(
            {
                "title": "A",
                "materials": [
                    {"type": "drive_file", "file_id": "file-1", "share_mode": "VIEW"}
                ],
            }
        )
        service = FakeService(
            {
                "courses.get": {"id": "1", "name": "Course", "courseState": "ACTIVE"},
                "courseWork.create": {"id": "cw-1", "courseId": "1", "state": "DRAFT"},
            }
        )

        def fake_auth(paths, **kwargs):
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="fp",
            )

        answers = PromptAnswers(["yes"])
        result, _, errors = self.run_cli(
            [
                "create-assignment",
                "--course-id",
                "1",
                "--spec",
                str(path),
                "--account",
                "teacher@example.edu",
                "--token",
                str(self.root / "token.json"),
            ],
            input_fn=answers,
            tty_check=lambda: True,
            auth_factory=fake_auth,
        )
        self.assertEqual(result, 0, errors)
        self.assertEqual(len(answers.prompts), 1)
        self.assertIn("Drive sharing: VIEW", answers.prompts[0])
        self.assertNotIn("SHARE ", answers.prompts[0])

    def test_authorize_requires_tty_and_explicit_replace_flag(self):
        credentials = self.root / "credentials.json"
        token = self.root / "token.json"
        credentials.write_text("{}", encoding="utf-8")
        token.write_text("{}", encoding="utf-8")
        calls = []

        def forbidden_auth(*args, **kwargs):
            calls.append((args, kwargs))
            raise AssertionError("authorize should have failed before OAuth")

        base = [
            "authorize",
            "--account",
            "teacher@example.edu",
            "--credentials",
            str(credentials),
            "--token",
            str(token),
        ]
        result, output, errors = self.run_cli(
            base,
            tty_check=lambda: False,
            auth_factory=forbidden_auth,
        )
        self.assertEqual(result, 2)
        self.assertEqual(output, "")
        self.assertIn("interactive terminal", errors)

        result, output, errors = self.run_cli(
            base,
            tty_check=lambda: True,
            auth_factory=forbidden_auth,
        )
        self.assertEqual(result, 2)
        self.assertEqual(output, "")
        self.assertIn("--replace-token", errors)
        self.assertEqual(calls, [])

    def test_authorize_replace_uses_explicit_flag_without_text_confirmation(self):
        credentials = self.root / "credentials.json"
        token = self.root / "token.json"
        credentials.write_text("{}", encoding="utf-8")
        token.write_text("{}", encoding="utf-8")
        calls = []

        def fake_auth(received_paths, **kwargs):
            calls.append((received_paths, kwargs))
            return CourseworkAuthSession(
                service=object(),
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=received_paths,
                client_id_fingerprint="client-fp",
            )

        def forbidden_input(_prompt):
            raise AssertionError("authorize prompted for redundant local confirmation")

        result, output, errors = self.run_cli(
            [
                "authorize",
                "--account",
                "teacher@example.edu",
                "--credentials",
                str(credentials),
                "--token",
                str(token),
                "--replace-token",
            ],
            input_fn=forbidden_input,
            tty_check=lambda: True,
            auth_factory=fake_auth,
        )
        self.assertEqual(result, 0, errors)
        self.assertTrue(calls[0][1]["force_authorize"])
        self.assertTrue(calls[0][1]["replace_token"])
        self.assertNotIn("teacher@example.edu", output)

    def test_known_and_unknown_failures_do_not_echo_secret_messages(self):
        path = self.write_spec({"title": "A"})

        for failure in (
            CredentialSecurityError("token refresh failed safely"),
            RuntimeError("access-token-very-secret"),
        ):
            with self.subTest(failure=type(failure).__name__):
                def bad_auth(*args, **kwargs):
                    raise failure

                result, output, errors = self.run_cli(
                    [
                        "create-assignment",
                        "--course-id",
                        "1",
                        "--spec",
                        str(path),
                        "--account",
                        "teacher@example.edu",
                        "--token",
                        str(self.root / "token.json"),
                    ],
                    tty_check=lambda: True,
                    auth_factory=bad_auth,
                )
                self.assertEqual(result, 2)
                self.assertEqual(output, "")
                self.assertNotIn("access-token-very-secret", errors)

    def test_outcome_unknown_is_reported_without_automatic_retry(self):
        path = self.write_spec({"title": "A"})
        token = self.write_token_placeholder()
        service = FakeService(
            {"courses.get": {"id": "1", "name": "Course", "courseState": "ACTIVE"}}
        )

        def fake_auth(paths, **kwargs):
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="fp",
            )

        def unknown_create(*args, **kwargs):
            raise GoogleClassroomOutcomeUnknownError("assignment create", "1")

        result, _, errors = self.run_cli(
            [
                "create-assignment",
                "--course-id",
                "1",
                "--spec",
                str(path),
                "--account",
                "teacher@example.edu",
                "--token",
                str(token),
                "--yes",
            ],
            tty_check=lambda: False,
            auth_factory=fake_auth,
            create_factory=unknown_create,
        )
        self.assertEqual(result, 3)
        error = json.loads(errors)
        self.assertEqual(error["error"], "outcome_unknown")

    def test_partial_create_is_serialized_as_exit_four_receipt(self):
        path = self.write_spec({"title": "A"})
        token = self.write_token_placeholder()
        service = FakeService(
            {"courses.get": {"id": "1", "name": "Course", "courseState": "ACTIVE"}}
        )

        def fake_auth(paths, **kwargs):
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="fp",
            )

        def partial_create(*args, **kwargs):
            raise GoogleClassroomPartialCreateError(
                stage="release",
                course_id="1",
                course_work_id="cw-1",
                alternate_link="https://classroom.google.com/c/example",
                rubric_created=True,
                intended_release="PUBLISHED",
                status=503,
            )

        result, output, errors = self.run_cli(
            [
                "create-assignment",
                "--course-id",
                "1",
                "--spec",
                str(path),
                "--account",
                "teacher@example.edu",
                "--token",
                str(token),
                "--yes",
            ],
            tty_check=lambda: False,
            auth_factory=fake_auth,
            create_factory=partial_create,
        )
        self.assertEqual(result, 4)
        self.assertEqual(output, "")
        receipt = json.loads(errors)
        self.assertEqual(
            receipt,
            {
                "alternateLink": "https://classroom.google.com/c/example",
                "courseId": "1",
                "courseWorkId": "cw-1",
                "error": "partial_create",
                "intendedRelease": "PUBLISHED",
                "rubricCreated": True,
                "stage": "release",
                "status": 503,
            },
        )

    def test_staged_mismatched_course_response_has_no_follow_up_mutation(self):
        path = self.write_spec(
            {
                "title": "A",
                "state": "PUBLISHED",
                "rubric": {
                    "scoring_mode": "UNSCORED",
                    "criteria": [{"levels": [{"title": "Complete"}]}],
                },
            }
        )
        service = FakeService(
            {
                "courses.get": {
                    "id": "1",
                    "name": "Course",
                    "courseState": "ACTIVE",
                },
                "courseWork.create": {
                    "id": "cw-1",
                    "courseId": "OTHER",
                    "state": "DRAFT",
                },
            }
        )

        def fake_auth(paths, **kwargs):
            return CourseworkAuthSession(
                service=service,
                profile={"id": "u", "emailAddress": "teacher@example.edu"},
                paths=paths,
                client_id_fingerprint="fp",
            )

        result, _, errors = self.run_cli(
            [
                "create-assignment",
                "--course-id",
                "1",
                "--spec",
                str(path),
                "--account",
                "teacher@example.edu",
                "--token",
                str(self.root / "token.json"),
            ],
            input_fn=PromptAnswers(["yes"]),
            tty_check=lambda: True,
            auth_factory=fake_auth,
        )
        self.assertEqual(result, 4)
        self.assertEqual(json.loads(errors)["stage"], "create-response")
        self.assertEqual(
            [call[0] for call in service.calls],
            ["courses.get", "courseWork.create"],
        )


if __name__ == "__main__":
    unittest.main()
