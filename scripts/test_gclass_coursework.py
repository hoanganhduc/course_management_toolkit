#!/usr/bin/env python3
"""Offline tests for Google Classroom assignment construction and mutations."""

from __future__ import annotations

import os
import sys
import unittest
from datetime import datetime, timedelta, timezone

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc.course_agent_common import CourseAgentError  # noqa: E402
from course_hoanganhduc.gclass_coursework import (  # noqa: E402
    GoogleClassroomOutcomeUnknownError,
    GoogleClassroomPartialCreateError,
    build_google_classroom_assignment_body,
    build_google_classroom_rubric_body,
    create_google_classroom_assignment_with_rubric,
    google_classroom_drive_material,
    google_classroom_link_material,
    google_classroom_youtube_material,
)


NOW = datetime(2026, 9, 1, tzinfo=timezone.utc)


class FakeRequest:
    def __init__(self, log, name, kwargs, result=None, error=None):
        self.log = log
        self.name = name
        self.kwargs = kwargs
        self.result = result
        self.error = error

    def execute(self, *, num_retries):
        if num_retries != 0:
            raise AssertionError("Every Google request must use num_retries=0")
        self.log.append((self.name, self.kwargs, num_retries))
        if self.error:
            raise self.error
        return self.result


class FakeHttpError(Exception):
    def __init__(self, status):
        super().__init__(f"HTTP {status} secret response body")
        self.resp = type("Response", (), {"status": status})()


class FakeRubrics:
    def __init__(self, owner):
        self.owner = owner

    def create(self, **kwargs):
        return self.owner.request("rubrics.create", kwargs)

    def list(self, **kwargs):
        return self.owner.request("rubrics.list", kwargs)


class FakeCourseWork:
    def __init__(self, owner):
        self.owner = owner

    def create(self, **kwargs):
        return self.owner.request("courseWork.create", kwargs)

    def patch(self, **kwargs):
        return self.owner.request("courseWork.patch", kwargs)

    def get(self, **kwargs):
        return self.owner.request("courseWork.get", kwargs)

    def rubrics(self):
        return FakeRubrics(self.owner)


class FakeCourses:
    def __init__(self, owner):
        self.owner = owner

    def courseWork(self):
        return FakeCourseWork(self.owner)

    def get(self, **kwargs):
        return self.owner.request("courses.get", kwargs)


class FakeService:
    def __init__(self, responses):
        self.responses = {key: list(value) for key, value in responses.items()}
        self.log = []

    def courses(self):
        return FakeCourses(self)

    def request(self, name, kwargs):
        queue = self.responses.get(name, [])
        if not queue:
            raise AssertionError(f"Unexpected request: {name}")
        value = queue.pop(0)
        if isinstance(value, BaseException):
            return FakeRequest(self.log, name, kwargs, error=value)
        return FakeRequest(self.log, name, kwargs, result=value)


class TestAssignmentBuilder(unittest.TestCase):
    def test_minimal_assignment_has_no_materials(self):
        body = build_google_classroom_assignment_body("Minimal", now=NOW)
        self.assertEqual(
            body,
            {
                "title": "Minimal",
                "workType": "ASSIGNMENT",
                "state": "DRAFT",
                "assigneeMode": "ALL_STUDENTS",
                "submissionModificationMode": "MODIFIABLE_UNTIL_TURNED_IN",
            },
        )

    def test_none_and_empty_materials_are_omitted(self):
        self.assertNotIn(
            "materials",
            build_google_classroom_assignment_body("A", materials=None, now=NOW),
        )
        self.assertNotIn(
            "materials",
            build_google_classroom_assignment_body("A", materials=[], now=NOW),
        )
        for invalid in ("", {}, 0):
            with self.subTest(invalid=invalid):
                with self.assertRaises(ValueError):
                    build_google_classroom_assignment_body(
                        "A", materials=invalid, now=NOW
                    )

    def test_full_assignment_and_utc_conversion(self):
        body = build_google_classroom_assignment_body(
            "Bài tập",
            description="Mô tả\nchi tiết",
            scheduled_at="2026-09-02T08:00:00+07:00",
            due_at="2026-09-03T08:30:00.123456+07:00",
            max_points=100,
            materials=[
                google_classroom_drive_material("Drive_ID-1", "STUDENT_COPY"),
                google_classroom_link_material("https://example.edu/a"),
                google_classroom_youtube_material("abc_DEF-123"),
            ],
            assignee_mode="INDIVIDUAL_STUDENTS",
            individual_student_ids=["123", "456"],
            submission_modification_mode="MODIFIABLE",
            topic_id="789",
            grading_period_mode="EXPLICIT",
            grading_period_id="gp-1",
            now=NOW,
        )
        self.assertEqual(
            body,
            {
                "title": "Bài tập",
                "description": "Mô tả\nchi tiết",
                "workType": "ASSIGNMENT",
                "state": "DRAFT",
                "assigneeMode": "INDIVIDUAL_STUDENTS",
                "individualStudentsOptions": {"studentIds": ["123", "456"]},
                "submissionModificationMode": "MODIFIABLE",
                "scheduledTime": "2026-09-02T01:00:00Z",
                "dueDate": {"year": 2026, "month": 9, "day": 3},
                "dueTime": {
                    "hours": 1,
                    "minutes": 30,
                    "seconds": 0,
                    "nanos": 123456000,
                },
                "maxPoints": 100,
                "materials": [
                    {
                        "driveFile": {
                            "driveFile": {"id": "Drive_ID-1"},
                            "shareMode": "STUDENT_COPY",
                        }
                    },
                    {"link": {"url": "https://example.edu/a"}},
                    {"youtubeVideo": {"id": "abc_DEF-123"}},
                ],
                "topicId": "789",
                "gradingPeriodId": "gp-1",
            },
        )

    def test_grading_period_tristate(self):
        auto = build_google_classroom_assignment_body("A", now=NOW)
        none = build_google_classroom_assignment_body(
            "A", grading_period_mode="NONE", now=NOW
        )
        self.assertNotIn("gradingPeriodId", auto)
        self.assertEqual(none["gradingPeriodId"], "")

    def test_text_and_material_boundaries(self):
        build_google_classroom_assignment_body("😀" * 3000, now=NOW)
        with self.assertRaises(ValueError):
            build_google_classroom_assignment_body("😀" * 3001, now=NOW)
        with self.assertRaises(ValueError):
            build_google_classroom_assignment_body("   ", now=NOW)
        with self.assertRaises(ValueError):
            build_google_classroom_assignment_body(
                "A", materials=[google_classroom_link_material("https://x.example")] * 21, now=NOW
            )
        with self.assertRaises(ValueError):
            google_classroom_link_material("http://example.edu")
        with self.assertRaises(ValueError):
            google_classroom_link_material("https://user:pass@example.edu/a")

    def test_invalid_cross_field_combinations(self):
        with self.assertRaises(ValueError):
            build_google_classroom_assignment_body(
                "A", state="PUBLISHED", scheduled_at=NOW + timedelta(hours=1), now=NOW
            )
        with self.assertRaises(ValueError):
            build_google_classroom_assignment_body(
                "A", scheduled_at=NOW + timedelta(seconds=59), now=NOW
            )
        with self.assertRaises(ValueError):
            build_google_classroom_assignment_body(
                "A",
                scheduled_at=NOW + timedelta(hours=2),
                due_at=NOW + timedelta(hours=1),
                now=NOW,
            )
        with self.assertRaises(ValueError):
            build_google_classroom_assignment_body("A", max_points=True, now=NOW)
        with self.assertRaises(ValueError):
            build_google_classroom_assignment_body(
                "A", assignee_mode="INDIVIDUAL_STUDENTS", individual_student_ids=[], now=NOW
            )
        with self.assertRaises(ValueError):
            build_google_classroom_assignment_body(
                "A", individual_student_ids=["123"], now=NOW
            )
        with self.assertRaises(ValueError):
            build_google_classroom_assignment_body(
                "A",
                assignee_mode="INDIVIDUAL_STUDENTS",
                individual_student_ids="123",
                now=NOW,
            )


class TestRubricBuilder(unittest.TestCase):
    def test_scored_and_unscored(self):
        scored = build_google_classroom_rubric_body(
            [
                {
                    "title": "Correctness",
                    "levels": [
                        {"title": "No", "points": 0},
                        {"title": "Yes", "points": 4},
                    ],
                }
            ],
            scoring_mode="SCORED",
        )
        self.assertEqual(scored["criteria"][0]["levels"][1]["points"], 4)
        unscored = build_google_classroom_rubric_body(
            [{"levels": [{"title": "Needs work"}, {"title": "Complete"}]}],
            scoring_mode="UNSCORED",
        )
        self.assertNotIn("points", unscored["criteria"][0]["levels"][0])

    def test_invalid_rubrics(self):
        cases = [
            ([], "SCORED"),
            ([{"levels": [{"title": "A", "points": 0}]}], "SCORED"),
            (
                [{"levels": [{"title": "A", "points": 1}, {"title": "B"}]}],
                "SCORED",
            ),
            (
                [{"levels": [{"title": "A", "points": 1}, {"title": "B", "points": 1}]}],
                "SCORED",
            ),
            ([{"levels": [{"title": "A", "points": 1}]}], "UNSCORED"),
            ([{"id": "bad", "levels": [{"title": "A"}]}], "UNSCORED"),
        ]
        for criteria, mode in cases:
            with self.subTest(criteria=criteria, mode=mode):
                with self.assertRaises(ValueError):
                    build_google_classroom_rubric_body(criteria, scoring_mode=mode)

    def test_oversized_json_integer_points_are_a_validation_error(self):
        with self.assertRaisesRegex(ValueError, "finite and nonnegative"):
            build_google_classroom_rubric_body(
                [
                    {
                        "levels": [
                            {"title": "Small", "points": 0},
                            {"title": "Huge", "points": 10**309},
                        ]
                    }
                ],
                scoring_mode="SCORED",
            )

    def test_multiple_zero_point_criteria_are_not_the_forbidden_single_zero_case(self):
        rubric = build_google_classroom_rubric_body(
            [
                {"levels": [{"title": "A", "points": 0}]},
                {"levels": [{"title": "B", "points": 0}]},
            ],
            scoring_mode="SCORED",
        )
        self.assertEqual(len(rubric["criteria"]), 2)


class TestMutationStateMachine(unittest.TestCase):
    def setUp(self):
        self.body = build_google_classroom_assignment_body("A", now=NOW)
        self.rubric = build_google_classroom_rubric_body(
            [{"levels": [{"title": "No", "points": 0}, {"title": "Yes", "points": 1}]}],
            scoring_mode="SCORED",
        )

    def rubric_response(self, rubric_id="r1", **overrides):
        response = {
            "courseId": "canonical",
            "courseWorkId": "cw1",
            "criteria": self.rubric["criteria"],
        }
        if rubric_id is not None:
            response["id"] = rubric_id
        response.update(overrides)
        return response

    def test_no_rubric_is_one_create(self):
        service = FakeService(
            {"courseWork.create": [[{"unused": True}]]}
        )
        # Correct the nested fixture deliberately so a response is a mapping.
        service.responses["courseWork.create"] = [
            {"courseId": "canonical", "id": "cw1", "state": "DRAFT"}
        ]
        result = create_google_classroom_assignment_with_rubric(
            service, "canonical", self.body, now=NOW
        )
        self.assertEqual(result["releaseStatus"], "DRAFT")
        self.assertEqual([x[0] for x in service.log], ["courseWork.create"])
        self.assertEqual(service.log[0][2], 0)

    def test_course_alias_is_resolved_before_the_first_write(self):
        service = FakeService(
            {
                "courses.get": [{"id": "canonical", "courseState": "ACTIVE"}],
                "courseWork.create": [
                    {"courseId": "canonical", "id": "cw1", "state": "DRAFT"}
                ],
            }
        )
        result = create_google_classroom_assignment_with_rubric(
            service, "d:math_101", self.body, now=NOW
        )
        self.assertEqual(result["courseWork"]["courseId"], "canonical")
        self.assertEqual(
            [entry[0] for entry in service.log],
            ["courses.get", "courseWork.create"],
        )
        self.assertEqual(
            service.log[0][1], {"id": "d:math_101", "fields": "id"}
        )
        self.assertEqual(service.log[1][1]["courseId"], "canonical")

    def test_no_rubric_published_and_scheduled_are_single_create_requests(self):
        scheduled = build_google_classroom_assignment_body(
            "Scheduled",
            scheduled_at=NOW + timedelta(hours=2),
            now=NOW,
        )
        cases = {
            "published": (
                dict(self.body, state="PUBLISHED"),
                {"courseId": "canonical", "id": "cw1", "state": "PUBLISHED"},
                "PUBLISHED",
            ),
            "scheduled": (
                scheduled,
                {
                    "courseId": "canonical",
                    "id": "cw1",
                    "state": "DRAFT",
                    "scheduledTime": "2026-09-01T02:00:00Z",
                },
                "SCHEDULED",
            ),
        }
        for label, (body, response, expected_release) in cases.items():
            with self.subTest(label=label):
                service = FakeService({"courseWork.create": [response]})
                result = create_google_classroom_assignment_with_rubric(
                    service, "canonical", body, now=NOW
                )
                self.assertEqual(result["releaseStatus"], expected_release)
                self.assertEqual(
                    [entry[0] for entry in service.log], ["courseWork.create"]
                )

    def test_rubric_then_publish(self):
        published = dict(self.body, state="PUBLISHED")
        service = FakeService(
            {
                "courseWork.create": [
                    {"courseId": "canonical", "id": "cw1", "state": "DRAFT"}
                ],
                "rubrics.create": [self.rubric_response()],
                "courseWork.patch": [
                    {"courseId": "canonical", "id": "cw1", "state": "PUBLISHED"}
                ],
            }
        )
        result = create_google_classroom_assignment_with_rubric(
            service, "canonical", published, self.rubric, now=NOW
        )
        self.assertEqual(result["releaseStatus"], "PUBLISHED")
        self.assertEqual(
            [x[0] for x in service.log],
            ["courseWork.create", "rubrics.create", "courseWork.patch"],
        )
        create_body = service.log[0][1]["body"]
        self.assertEqual(create_body["state"], "DRAFT")
        self.assertNotIn("scheduledTime", create_body)
        self.assertEqual(service.log[2][1]["body"], {"state": "PUBLISHED"})
        self.assertEqual(service.log[2][1]["updateMask"], "state")

    def test_rubric_failure_leaves_draft(self):
        service = FakeService(
            {
                "courseWork.create": [
                    {"courseId": "canonical", "id": "cw1", "state": "DRAFT"}
                ],
                "rubrics.create": [FakeHttpError(403)],
            }
        )
        with self.assertRaises(GoogleClassroomPartialCreateError) as ctx:
            create_google_classroom_assignment_with_rubric(
                service, "canonical", self.body, self.rubric, now=NOW
            )
        self.assertEqual(ctx.exception.stage, "rubric")
        self.assertEqual(ctx.exception.course_work_id, "cw1")
        self.assertEqual([x[0] for x in service.log], ["courseWork.create", "rubrics.create"])

    def test_ambiguous_initial_create_is_never_retried(self):
        service = FakeService({"courseWork.create": [TimeoutError("secret")]})
        with self.assertRaises(GoogleClassroomOutcomeUnknownError):
            create_google_classroom_assignment_with_rubric(
                service, "canonical", self.body, now=NOW
            )
        self.assertEqual([x[0] for x in service.log], ["courseWork.create"])

    def test_non_mapping_successful_create_is_outcome_unknown(self):
        service = FakeService({"courseWork.create": [["unexpected response"]]})
        with self.assertRaises(GoogleClassroomOutcomeUnknownError):
            create_google_classroom_assignment_with_rubric(
                service, "canonical", self.body, now=NOW
            )
        self.assertEqual([x[0] for x in service.log], ["courseWork.create"])

    def test_missing_created_id_is_reported_as_unknown(self):
        service = FakeService(
            {"courseWork.create": [{"courseId": "canonical", "state": "DRAFT"}]}
        )
        with self.assertRaises(GoogleClassroomOutcomeUnknownError):
            create_google_classroom_assignment_with_rubric(
                service, "canonical", self.body, now=NOW
            )
        self.assertEqual([x[0] for x in service.log], ["courseWork.create"])

    def test_created_resource_must_match_requested_release(self):
        service = FakeService(
            {
                "courseWork.create": [
                    {"courseId": "canonical", "id": "cw1", "state": "PUBLISHED"}
                ]
            }
        )
        with self.assertRaises(GoogleClassroomPartialCreateError) as ctx:
            create_google_classroom_assignment_with_rubric(
                service, "canonical", self.body, now=NOW
            )
        self.assertEqual(ctx.exception.stage, "create-response")

    def test_rubric_response_without_id_leaves_draft(self):
        published = dict(self.body, state="PUBLISHED")
        service = FakeService(
            {
                "courseWork.create": [
                    {"courseId": "canonical", "id": "cw1", "state": "DRAFT"}
                ],
                "rubrics.create": [self.rubric_response(rubric_id=None)],
            }
        )
        with self.assertRaises(GoogleClassroomPartialCreateError) as ctx:
            create_google_classroom_assignment_with_rubric(
                service, "canonical", published, self.rubric, now=NOW
            )
        self.assertEqual(ctx.exception.stage, "rubric")
        self.assertEqual(
            [entry[0] for entry in service.log],
            ["courseWork.create", "rubrics.create"],
        )

    def test_release_response_must_show_intended_state(self):
        published = dict(self.body, state="PUBLISHED")
        service = FakeService(
            {
                "courseWork.create": [
                    {"courseId": "canonical", "id": "cw1", "state": "DRAFT"}
                ],
                "rubrics.create": [self.rubric_response()],
                "courseWork.patch": [
                    {"courseId": "canonical", "id": "cw1", "state": "DRAFT"}
                ],
            }
        )
        with self.assertRaises(GoogleClassroomPartialCreateError) as ctx:
            create_google_classroom_assignment_with_rubric(
                service, "canonical", published, self.rubric, now=NOW
            )
        self.assertEqual(ctx.exception.stage, "release")

    def test_direct_mutation_is_refused_in_agent_mode(self):
        previous = {
            "COURSE_AGENT_MODE": os.environ.get("COURSE_AGENT_MODE"),
            "GCLASS_COURSE_ALLOWLIST": os.environ.get("GCLASS_COURSE_ALLOWLIST"),
        }
        os.environ["COURSE_AGENT_MODE"] = "1"
        os.environ["GCLASS_COURSE_ALLOWLIST"] = "canonical"
        try:
            service = FakeService({})
            with self.assertRaises(CourseAgentError):
                create_google_classroom_assignment_with_rubric(
                    service, "canonical", self.body, now=NOW
                )
            self.assertEqual(service.log, [])
        finally:
            for key, value in previous.items():
                if value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = value

    def test_ambiguous_patch_reconciles_with_one_get(self):
        published = dict(self.body, state="PUBLISHED")
        service = FakeService(
            {
                "courseWork.create": [
                    {"courseId": "canonical", "id": "cw1", "state": "DRAFT"}
                ],
                "rubrics.create": [self.rubric_response()],
                "courseWork.patch": [TimeoutError("secret")],
                "courseWork.get": [
                    {"courseId": "canonical", "id": "cw1", "state": "PUBLISHED"}
                ],
            }
        )
        result = create_google_classroom_assignment_with_rubric(
            service, "canonical", published, self.rubric, now=NOW
        )
        self.assertEqual(result["releaseStatus"], "PUBLISHED")
        self.assertEqual(service.log[-1][0], "courseWork.get")

    def test_scheduled_rubric_stages_then_schedules_with_zero_retries(self):
        scheduled = build_google_classroom_assignment_body(
            "A",
            scheduled_at=NOW + timedelta(hours=2),
            due_at=NOW + timedelta(hours=3),
            now=NOW,
        )
        service = FakeService(
            {
                "courseWork.create": [
                    {"courseId": "canonical", "id": "cw1", "state": "DRAFT"}
                ],
                "rubrics.create": [self.rubric_response()],
                "courseWork.patch": [
                    {
                        "courseId": "canonical",
                        "id": "cw1",
                        "state": "DRAFT",
                        "scheduledTime": "2026-09-01T02:00:00Z",
                    }
                ],
            }
        )
        result = create_google_classroom_assignment_with_rubric(
            service, "canonical", scheduled, self.rubric, now=NOW
        )
        self.assertEqual(result["releaseStatus"], "SCHEDULED")
        self.assertEqual(
            [entry[0] for entry in service.log],
            ["courseWork.create", "rubrics.create", "courseWork.patch"],
        )
        self.assertEqual(
            service.log[-1][1],
            {
                "courseId": "canonical",
                "id": "cw1",
                "updateMask": "scheduledTime",
                "body": {"scheduledTime": "2026-09-01T02:00:00Z"},
            },
        )
        self.assertTrue(all(entry[2] == 0 for entry in service.log))

    def test_mismatched_create_course_stops_before_rubric_or_release(self):
        published = dict(self.body, state="PUBLISHED")
        service = FakeService(
            {
                "courseWork.create": [
                    {"courseId": "OTHER", "id": "cw1", "state": "DRAFT"}
                ]
            }
        )
        with self.assertRaises(GoogleClassroomPartialCreateError) as ctx:
            create_google_classroom_assignment_with_rubric(
                service, "canonical", published, self.rubric, now=NOW
            )
        self.assertEqual(ctx.exception.stage, "create-response")
        self.assertEqual(ctx.exception.course_id, "canonical")
        self.assertEqual([entry[0] for entry in service.log], ["courseWork.create"])

    def test_ambiguous_rubric_reconciles_only_one_exact_resource(self):
        service = FakeService(
            {
                "courseWork.create": [
                    {"courseId": "canonical", "id": "cw1", "state": "DRAFT"}
                ],
                "rubrics.create": [TimeoutError("secret")],
                "rubrics.list": [{"rubrics": [self.rubric_response()]}],
            }
        )
        result = create_google_classroom_assignment_with_rubric(
            service, "canonical", self.body, self.rubric, now=NOW
        )
        self.assertEqual(result["rubric"]["id"], "r1")
        self.assertEqual(
            [entry[0] for entry in service.log],
            ["courseWork.create", "rubrics.create", "rubrics.list"],
        )
        self.assertEqual(service.log[-1][1]["pageSize"], 1)
        self.assertTrue(all(entry[2] == 0 for entry in service.log))

    def test_ambiguous_rubric_nonmatches_remain_unknown_and_draft(self):
        cases = {
            "empty": {"rubrics": []},
            "wrong-parent": {
                "rubrics": [self.rubric_response(courseId="OTHER")]
            },
            "wrong-work-parent": {
                "rubrics": [self.rubric_response(courseWorkId="OTHER")]
            },
            "wrong-criteria": {
                "rubrics": [self.rubric_response(criteria=[{"levels": []}])]
            },
            "multiple": {
                "rubrics": [self.rubric_response(), self.rubric_response("r2")]
            },
            "paginated": {
                "rubrics": [self.rubric_response()],
                "nextPageToken": "unexpected",
            },
        }
        for label, listed in cases.items():
            with self.subTest(label=label):
                service = FakeService(
                    {
                        "courseWork.create": [
                            {
                                "courseId": "canonical",
                                "id": "cw1",
                                "state": "DRAFT",
                            }
                        ],
                        "rubrics.create": [TimeoutError("secret")],
                        "rubrics.list": [listed],
                    }
                )
                with self.assertRaises(GoogleClassroomPartialCreateError) as ctx:
                    create_google_classroom_assignment_with_rubric(
                        service,
                        "canonical",
                        dict(self.body, state="PUBLISHED"),
                        self.rubric,
                        now=NOW,
                    )
                self.assertEqual(ctx.exception.stage, "rubric")
                self.assertIsNone(ctx.exception.rubric_created)
                self.assertNotIn(
                    "courseWork.patch", [entry[0] for entry in service.log]
                )

    def test_ambiguous_rubric_list_failure_remains_unknown(self):
        service = FakeService(
            {
                "courseWork.create": [
                    {"courseId": "canonical", "id": "cw1", "state": "DRAFT"}
                ],
                "rubrics.create": [TimeoutError("secret")],
                "rubrics.list": [FakeHttpError(503)],
            }
        )
        with self.assertRaises(GoogleClassroomPartialCreateError) as ctx:
            create_google_classroom_assignment_with_rubric(
                service, "canonical", self.body, self.rubric, now=NOW
            )
        self.assertEqual(ctx.exception.stage, "rubric")
        self.assertIsNone(ctx.exception.rubric_created)

    def test_definite_release_failure_leaves_published_intent_partial(self):
        service = FakeService(
            {
                "courseWork.create": [
                    {"courseId": "canonical", "id": "cw1", "state": "DRAFT"}
                ],
                "rubrics.create": [self.rubric_response()],
                "courseWork.patch": [FakeHttpError(403)],
            }
        )
        with self.assertRaises(GoogleClassroomPartialCreateError) as ctx:
            create_google_classroom_assignment_with_rubric(
                service,
                "canonical",
                dict(self.body, state="PUBLISHED"),
                self.rubric,
                now=NOW,
            )
        self.assertEqual(ctx.exception.stage, "release")
        self.assertEqual(ctx.exception.status, 403)
        self.assertTrue(ctx.exception.rubric_created)

    def test_ambiguous_release_mismatched_get_stays_partial(self):
        cases = {
            "course": {"courseId": "OTHER", "id": "cw1", "state": "PUBLISHED"},
            "course-work": {
                "courseId": "canonical",
                "id": "OTHER",
                "state": "PUBLISHED",
            },
        }
        for label, observed in cases.items():
            with self.subTest(label=label):
                service = FakeService(
                    {
                        "courseWork.create": [
                            {
                                "courseId": "canonical",
                                "id": "cw1",
                                "state": "DRAFT",
                            }
                        ],
                        "rubrics.create": [self.rubric_response()],
                        "courseWork.patch": [TimeoutError("secret")],
                        "courseWork.get": [observed],
                    }
                )
                with self.assertRaises(GoogleClassroomPartialCreateError) as ctx:
                    create_google_classroom_assignment_with_rubric(
                        service,
                        "canonical",
                        dict(self.body, state="PUBLISHED"),
                        self.rubric,
                        now=NOW,
                    )
                self.assertEqual(ctx.exception.stage, "release")
                self.assertEqual(ctx.exception.observed_course_work, observed)

    def test_time_sensitive_fields_are_revalidated_before_first_write(self):
        cases = [
            build_google_classroom_assignment_body(
                "Scheduled",
                scheduled_at=NOW + timedelta(hours=2),
                now=NOW,
            ),
            build_google_classroom_assignment_body(
                "Due",
                due_at=NOW + timedelta(hours=2),
                now=NOW,
            ),
        ]
        for body in cases:
            with self.subTest(body=body["title"]):
                service = FakeService({})
                with self.assertRaises(ValueError):
                    create_google_classroom_assignment_with_rubric(
                        service,
                        "canonical",
                        body,
                        now=NOW + timedelta(hours=3),
                    )
                self.assertEqual(service.log, [])

    def test_time_sensitive_fields_are_revalidated_before_staged_release(self):
        cases = {
            "schedule-lead-expired": (
                build_google_classroom_assignment_body(
                    "Scheduled",
                    scheduled_at=NOW + timedelta(hours=2),
                    due_at=NOW + timedelta(hours=3),
                    now=NOW,
                ),
                NOW + timedelta(hours=1, minutes=59, seconds=30),
            ),
            "due-expired": (
                build_google_classroom_assignment_body(
                    "Published",
                    state="PUBLISHED",
                    due_at=NOW + timedelta(hours=2),
                    now=NOW,
                ),
                NOW + timedelta(hours=3),
            ),
        }
        for label, (body, release_now) in cases.items():
            with self.subTest(label=label):
                service = FakeService(
                    {
                        "courseWork.create": [
                            {
                                "courseId": "canonical",
                                "id": "cw1",
                                "state": "DRAFT",
                            }
                        ],
                        "rubrics.create": [self.rubric_response()],
                    }
                )
                times = iter([NOW, release_now])
                with self.assertRaises(GoogleClassroomPartialCreateError) as ctx:
                    create_google_classroom_assignment_with_rubric(
                        service,
                        "canonical",
                        body,
                        self.rubric,
                        now=lambda: next(times),
                    )
                self.assertEqual(ctx.exception.stage, "release-timing")
                self.assertEqual(ctx.exception.course_work_id, "cw1")
                self.assertTrue(ctx.exception.rubric_created)
                self.assertEqual(
                    [entry[0] for entry in service.log],
                    ["courseWork.create", "rubrics.create"],
                )


if __name__ == "__main__":
    unittest.main()
