#!/usr/bin/env python3
"""Offline tests for the Classroom50 score-import lane.

The module under test is stdlib-only, never imports the database layer and
never imports pandas, so these tests run under a bare interpreter.  Every
network call goes through an injected runner, so nothing here reaches GitHub.
"""

from __future__ import annotations

import json
import os
import sys
import unittest
from typing import Any, Dict, List, Optional, Sequence

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc.c50_cli import Classroom50Error, RunResult  # noqa: E402
from course_hoanganhduc.c50_scores import (  # noqa: E402
    FIELD_COLLECTED_AT,
    FIELD_DETAILS,
    FIELD_GRADES,
    FIELD_OVERRIDES,
    FIELD_SUBMISSIONS,
    STATE_NOT_COLLECTED,
    STATE_NOT_SUBMITTED,
    STATE_SUBMITTED,
    STATE_UNKNOWN_STALE,
    credit_findings,
    credit_snapshot,
    credited_members,
    expand_entries,
    fetch_scores,
    fetch_scores_collected_at,
    import_scores,
    merge_into_students,
    parse_assignments,
    parse_scores,
    previous_credit_snapshot,
    staff_logins,
    submission_state,
)

COLLECTED_AT = "2026-09-12T09:00:00Z"


def payload(**overrides: Any) -> Dict[str, Any]:
    """One validated ``result.json`` body, minus the ``assignment`` key.

    The required string fields are spelled the way ``collect_scores.py`` spells
    them: ``datetime`` (not ``submitted_at``), and ``release`` holding a release
    name rather than a timestamp.
    """
    body = {
        "schema": "classroom50/result/v1",
        "classroom": "vnu-hus-mat1206e-winter-2026",
        "owner": "alice",
        "assignment_type": "individual",
        "submission": "https://github.com/vnu-hus/x/pull/1",
        "commit": "a" * 40,
        "release": "autograding-2026-09-10",
        "review": "https://github.com/vnu-hus/x/pull/1#issue-1",
        "datetime": "2026-09-10T10:00:00Z",
        "score": 85,
        "max-score": 100,
        "tests": [{"test-name": "t1", "passed": True, "score": 85, "max-score": 100}],
    }
    body.update(overrides)
    return body


def scores_doc(assignments: Dict[str, Any]) -> Dict[str, Any]:
    return {"schema": "classroom50/scores/v1", "assignments": assignments}


ROSTER = [
    {"login": "alice", "kind": "user", "role": "", "github_id": 111},
    {"login": "bob", "kind": "user", "role": "", "github_id": 222},
    {"login": "carol", "kind": "user", "role": "", "github_id": 333},
    {"login": "hoanganhduc", "kind": "user", "role": "teacher", "github_id": 700},
]

MANIFEST = {
    "assignments": [
        {
            "slug": "ch01-introduction",
            "mode": "individual",
            "available_from": "2026-09-04T06:00:00Z",
            "due": "2026-09-30T16:59:00Z",
        },
        {
            "slug": "w00-group-collaboration",
            "mode": "group",
            "max_group_size": 5,
            "available_from": "2026-09-06T23:00:00Z",
            "due": "2026-09-11T05:00:00Z",
        },
        {
            "slug": "final-project",
            "mode": "group",
            "empty_repo": True,
            "grading": {"mode": "off"},
            "available_from": "2026-09-11T06:00:00Z",
            "due": "2026-11-04T16:59:00Z",
        },
    ]
}


class Student:
    """A database record as the rest of the toolkit sees one: an attribute bag."""

    def __init__(self, name: str, student_id: str, login: str, github_id: str = "") -> None:
        self.Name = name
        setattr(self, "Student ID", student_id)
        setattr(self, "GitHub Username", login)
        if github_id:
            setattr(self, "GitHub ID", github_id)
        self.Email = f"{student_id}@hus.edu.vn"


def klass() -> List[Student]:
    return [
        Student("Nguyễn Thị An", "24001111", "alice", "111"),
        Student("Trần Văn Bình", "24002222", "bob", "222"),
        Student("Lê Thị Cúc", "24003333", "carol", "333"),
    ]


class FakeRunner:
    """A ``gh`` stand-in that answers by endpoint and counts its calls."""

    def __init__(
        self,
        *,
        scores: Any = None,
        commits: Any = None,
        roster: Any = None,
        assignments: Any = None,
        failures: Optional[Dict[str, RunResult]] = None,
    ) -> None:
        self.scores = scores if scores is not None else scores_doc({})
        self.commits = (
            commits
            if commits is not None
            else [{"commit": {"committer": {"date": COLLECTED_AT}}}]
        )
        self.roster = roster if roster is not None else ROSTER
        self.assignments = assignments if assignments is not None else MANIFEST
        self.failures = failures or {}
        self.calls: List[List[str]] = []

    def __call__(self, argv: Sequence[str]) -> RunResult:
        self.calls.append(list(argv))
        joined = " ".join(argv)
        for marker, result in self.failures.items():
            if marker in joined:
                return result
        if "scores.json" in joined and "commits?path=" not in joined:
            body: Any = self.scores
        elif "commits?path=" in joined:
            body = self.commits
        elif "roster list" in joined:
            body = self.roster
        elif "assignment list" in joined:
            body = self.assignments
        else:
            return RunResult(returncode=1, stderr=f"unexpected call: {joined}")
        return RunResult(returncode=0, stdout=json.dumps(body))


def written(students: Sequence[Student]) -> int:
    """How many C50 fields the merge actually set across the class."""
    fields = (
        FIELD_GRADES,
        FIELD_SUBMISSIONS,
        FIELD_DETAILS,
        FIELD_OVERRIDES,
        FIELD_COLLECTED_AT,
    )
    return sum(1 for s in students for f in fields if hasattr(s, f))


class TestDocumentShape(unittest.TestCase):
    def test_unknown_schema_is_refused_and_quotes_what_it_read(self):
        with self.assertRaises(Classroom50Error) as ctx:
            parse_scores({"schema": "classroom50/scores/v2", "assignments": {}})
        self.assertIn("classroom50/scores/v2", str(ctx.exception))

    def test_unknown_schema_writes_nothing(self):
        students = klass()
        runner = FakeRunner(scores={"schema": "nope/v9", "assignments": {}})
        with self.assertRaises(Classroom50Error):
            import_scores(students, org="VNU-HUS", classroom="c", runner=runner)
        self.assertEqual(written(students), 0)

    def test_todays_real_file_runs_clean(self):
        # Both classrooms hold exactly this today: no schema key, no buckets.
        students = klass()
        runner = FakeRunner(scores={"assignments": {}})
        report = import_scores(students, org="VNU-HUS", classroom="c", runner=runner)
        self.assertEqual(report.updated, 0)
        self.assertEqual(report.matched, [])
        self.assertEqual(report.unmatched, [])
        self.assertEqual(report.findings, [])

    def test_error_object_instead_of_a_document_is_caught_by_the_type_check(self):
        runner = FakeRunner()
        runner.failures = {}
        runner.scores = {"message": "Not Found", "documentation_url": "...", "status": "404"}
        # It parses as a dict, so only the schema/assignments check can catch it.
        with self.assertRaises(Classroom50Error):
            parse_scores(runner.scores)

    def test_commits_endpoint_must_answer_with_a_list(self):
        runner = FakeRunner(commits={"message": "Not Found", "status": "404"})
        with self.assertRaises(Classroom50Error) as ctx:
            fetch_scores_collected_at("VNU-HUS", "c", runner=runner)
        self.assertIn("expected list", str(ctx.exception))

    def test_nonzero_exit_with_partial_data_is_an_error_not_a_short_list(self):
        partial = RunResult(
            returncode=1,
            stdout=json.dumps(scores_doc({"ch01-introduction": {"type": "individual", "entries": []}})),
            stderr="gh: API rate limit exceeded for user ID 700 (HTTP 403)",
        )
        students = klass()
        runner = FakeRunner(failures={"scores.json": partial})
        with self.assertRaises(Classroom50Error) as ctx:
            fetch_scores("VNU-HUS", "c", runner=runner)
        self.assertIn("rate limit", str(ctx.exception))
        self.assertEqual(written(students), 0)

    def test_private_repo_hint_names_the_404(self):
        runner = FakeRunner(
            failures={"scores.json": RunResult(returncode=1, stderr="gh: Not Found (HTTP 404)")}
        )
        with self.assertRaises(Classroom50Error) as ctx:
            fetch_scores("VNU-HUS", "c", runner=runner)
        self.assertIn("404, not 403", str(ctx.exception))

    def test_entry_without_owner_is_skipped_and_the_rest_still_parse(self):
        doc = parse_scores(
            scores_doc(
                {
                    "ch01-introduction": {
                        "type": "individual",
                        "entries": [
                            {"submissions": [payload()]},
                            {"owner": "bob", "submissions": [payload(owner="bob")]},
                        ],
                    }
                }
            )
        )
        self.assertEqual([e.owner for e in doc.assignments["ch01-introduction"].entries], ["bob"])
        self.assertTrue(any("owner" in w for w in doc.warnings))


class TestIndividualGrades(unittest.TestCase):
    def merge(self, doc: Dict[str, Any], students: Sequence[Student]):
        parsed = parse_scores(doc)
        return merge_into_students(
            students,
            expand_entries(parsed),
            collected_at=COLLECTED_AT,
            roster=ROSTER,
            assignments=parse_assignments(MANIFEST),
            warnings=parsed.warnings,
        )

    def test_integer_score_lands_on_the_record_with_the_tests(self):
        students = klass()
        self.merge(
            scores_doc(
                {
                    "ch01-introduction": {
                        "type": "individual",
                        "entries": [{"owner": "alice", "submissions": [payload()]}],
                    }
                }
            ),
            students,
        )
        grade = getattr(students[0], FIELD_GRADES)["ch01-introduction"]
        self.assertEqual(grade["grade"], 85)
        self.assertEqual(grade["max_points"], 100)
        self.assertEqual(grade["release"], "autograding-2026-09-10")
        self.assertEqual(grade["datetime"], "2026-09-10T10:00:00Z")
        self.assertFalse(grade["group"])
        details = getattr(students[0], FIELD_DETAILS)["ch01-introduction"]
        self.assertEqual(details[0]["tests"][0]["test-name"], "t1")

    def test_newest_submission_first_is_preserved_and_wins(self):
        students = klass()
        newest = payload(score=90, datetime="2026-09-11T10:00:00Z", commit="b" * 40)
        oldest = payload(score=10, datetime="2026-09-01T10:00:00Z")
        self.merge(
            scores_doc(
                {
                    "ch01-introduction": {
                        "type": "individual",
                        "entries": [{"owner": "alice", "submissions": [newest, oldest]}],
                    }
                }
            ),
            students,
        )
        self.assertEqual(getattr(students[0], FIELD_GRADES)["ch01-introduction"]["grade"], 90)
        stored = getattr(students[0], FIELD_DETAILS)["ch01-introduction"]
        self.assertEqual([p["score"] for p in stored], [90, 10])

    def test_late_is_copied_not_recomputed(self):
        students = klass()
        # Submitted well before the deadline, yet upstream marked it late: the
        # stored flag is upstream's, so it must survive verbatim.
        self.merge(
            scores_doc(
                {
                    "ch01-introduction": {
                        "type": "individual",
                        "entries": [
                            {"owner": "alice", "submissions": [payload(late=True)]}
                        ],
                    }
                }
            ),
            students,
        )
        self.assertIs(getattr(students[0], FIELD_GRADES)["ch01-introduction"]["late"], True)

    def test_override_is_recorded_and_nothing_is_written_upstream(self):
        students = klass()
        runner = FakeRunner(
            scores=scores_doc(
                {
                    "ch01-introduction": {
                        "type": "individual",
                        "entries": [
                            {"owner": "alice", "override": True, "submissions": [payload()]}
                        ],
                    }
                }
            )
        )
        report = import_scores(students, org="VNU-HUS", classroom="c", runner=runner)
        self.assertEqual(getattr(students[0], FIELD_OVERRIDES), {"ch01-introduction": True})
        self.assertEqual(report.overrides, ["ch01-introduction/alice"])
        for argv in runner.calls:
            joined = " ".join(argv)
            for verb in ("-X PUT", "-X POST", "-X PATCH", "-X DELETE", "--method"):
                self.assertNotIn(verb, joined)

    def test_unknown_login_is_unmatched_and_creates_no_student(self):
        students = klass()
        report = self.merge(
            scores_doc(
                {
                    "ch01-introduction": {
                        "type": "individual",
                        "entries": [{"owner": "mallory", "submissions": [payload()]}],
                    }
                }
            ),
            students,
        )
        self.assertEqual(report.unmatched, ["mallory"])
        self.assertEqual(len(students), 3)


class TestJoinKeys(unittest.TestCase):
    def test_numeric_github_id_survives_a_username_change(self):
        students = klass()
        # The student renamed the account; the database still holds the old
        # login but the same immutable id the roster reports.
        setattr(students[0], "GitHub Username", "alice-old")
        roster = [dict(row) for row in ROSTER]
        roster[0]["login"] = "alice-new"
        parsed = parse_scores(
            scores_doc(
                {
                    "ch01-introduction": {
                        "type": "individual",
                        "entries": [
                            {"owner": "alice-new", "submissions": [payload(owner="alice-new")]}
                        ],
                    }
                }
            )
        )
        report = merge_into_students(
            students, expand_entries(parsed), collected_at=COLLECTED_AT, roster=roster
        )
        self.assertEqual(report.unmatched, [])
        self.assertIn("ch01-introduction", getattr(students[0], FIELD_GRADES))

    def test_roster_without_github_id_falls_back_to_the_username_with_a_warning(self):
        students = klass()
        roster = [{"login": row["login"], "role": row["role"]} for row in ROSTER]
        parsed = parse_scores(
            scores_doc(
                {
                    "ch01-introduction": {
                        "type": "individual",
                        "entries": [{"owner": "alice", "submissions": [payload()]}],
                    }
                }
            )
        )
        report = merge_into_students(
            students, expand_entries(parsed), collected_at=COLLECTED_AT, roster=roster
        )
        self.assertEqual(report.matched, ["alice"])
        self.assertTrue(any("github_id" in w for w in report.warnings))


class TestGroupCredit(unittest.TestCase):
    def group_doc(self, members: Sequence[str], owner: str = "alice") -> Dict[str, Any]:
        return scores_doc(
            {
                "w00-group-collaboration": {
                    "type": "group",
                    "entries": [
                        {
                            "owner": owner,
                            "member_usernames": list(members),
                            "submissions": [
                                payload(owner=owner, assignment_type="group", score=70)
                            ],
                        }
                    ],
                }
            }
        )

    def test_every_credited_member_gets_the_same_grade(self):
        students = klass()
        parsed = parse_scores(self.group_doc(["alice", "bob", "carol"]))
        merge_into_students(
            students, expand_entries(parsed), collected_at=COLLECTED_AT, roster=ROSTER
        )
        for student in students:
            grade = getattr(student, FIELD_GRADES)["w00-group-collaboration"]
            self.assertEqual(grade["grade"], 70)
            self.assertTrue(grade["group"])
            self.assertEqual(grade["owner"], "alice")

    def test_staff_in_member_usernames_is_subtracted_not_reported_as_a_fault(self):
        # The collector's roster is the union of the student team and every
        # staff team, so a teacher on a group repository is credited upstream by
        # design.  That is upstream working, but it is not a grade.
        staff = staff_logins(ROSTER)
        self.assertEqual(staff, frozenset({"hoanganhduc"}))
        students = klass()
        runner = FakeRunner(scores=self.group_doc(["alice", "bob", "hoanganhduc"]))
        report = import_scores(students, org="VNU-HUS", classroom="c", runner=runner)
        self.assertEqual(report.credited, {"w00-group-collaboration": {"alice": ["alice", "bob"]}})
        self.assertNotIn("hoanganhduc", report.matched)
        self.assertNotIn("hoanganhduc", report.unmatched)
        self.assertEqual(report.findings, [])

    def test_owner_is_kept_when_every_member_is_staff(self):
        entry = parse_scores(self.group_doc(["hoanganhduc"], owner="alice")).assignments[
            "w00-group-collaboration"
        ].entries[0]
        self.assertEqual(
            credited_members(entry, kind="group", staff_logins=frozenset({"hoanganhduc"})),
            ["alice"],
        )

    def test_owner_only_credit_names_the_teammates_who_get_nothing(self):
        students = klass()
        parsed = parse_scores(self.group_doc(["alice"]))
        findings = credit_findings(
            parsed,
            students,
            staff_logins=staff_logins(ROSTER),
            expected_members={"w00-group-collaboration": {"alice": ["alice", "bob", "carol"]}},
        )
        self.assertEqual({f.code for f in findings}, {"group_credit_owner_only"})
        self.assertEqual(sorted(f.found for f in findings), ["bob", "carol"])
        for finding in findings:
            self.assertTrue(finding.fix.strip())
            self.assertEqual(finding.severity, "error")

    def test_dropped_member_is_caught_against_the_previous_snapshot(self):
        students = klass()
        parsed = parse_scores(self.group_doc(["alice", "bob"]))
        previous = {"w00-group-collaboration": {"alice": ["alice", "bob", "carol"]}}
        findings = credit_findings(parsed, students, previous=previous)
        self.assertEqual([f.code for f in findings], ["group_member_dropped"])
        self.assertEqual(findings[0].found, "carol")

    def test_first_run_has_no_snapshot_so_nothing_is_reported_as_dropped(self):
        students = klass()
        parsed = parse_scores(self.group_doc(["alice", "bob"]))
        self.assertEqual(credit_findings(parsed, students), [])

    def test_previous_snapshot_round_trips_through_the_database(self):
        students = klass()
        parsed = parse_scores(self.group_doc(["alice", "bob", "carol"]))
        merge_into_students(
            students, expand_entries(parsed), collected_at=COLLECTED_AT, roster=ROSTER
        )
        self.assertEqual(previous_credit_snapshot(students), credit_snapshot(parsed))


class TestSubmissionState(unittest.TestCase):
    def test_a_snapshot_older_than_the_deadline_never_claims_not_submitted(self):
        state = submission_state(
            submitted=False,
            collected_at="2026-09-08T12:00:00Z",
            available_from="2026-09-04T06:00:00Z",
            due="2026-09-30T16:59:00Z",
        )
        self.assertEqual(state, STATE_UNKNOWN_STALE)
        self.assertNotEqual(state, STATE_NOT_SUBMITTED)

    def test_a_snapshot_older_than_the_opening_says_only_that(self):
        self.assertEqual(
            submission_state(
                submitted=False,
                collected_at="2026-08-10T04:52:10Z",
                available_from="2026-09-04T06:00:00Z",
                due="2026-09-30T16:59:00Z",
            ),
            STATE_NOT_COLLECTED,
        )

    def test_after_the_deadline_the_absence_is_meaningful(self):
        self.assertEqual(
            submission_state(
                submitted=False,
                collected_at="2026-10-01T00:00:00Z",
                available_from="2026-09-04T06:00:00Z",
                due="2026-09-30T16:59:00Z",
            ),
            STATE_NOT_SUBMITTED,
        )

    def test_naive_timestamps_are_read_as_utc_not_as_local_time(self):
        # +07:00 would shift a naive reading across the boundary below.
        self.assertEqual(
            submission_state(
                submitted=False,
                collected_at="2026-09-30T18:00:00",
                due="2026-09-30T16:59:00Z",
            ),
            STATE_NOT_SUBMITTED,
        )

    def test_a_stale_snapshot_writes_unknown_for_an_uncollected_slug(self):
        students = klass()
        # Taken after ch01 opened (2026-09-04T06:00Z) but before the group
        # assignment did (2026-09-06T23:00Z), so the same snapshot is stale for
        # one slug and simply too early for the other.
        runner = FakeRunner(
            commits=[{"commit": {"committer": {"date": "2026-09-05T12:00:00Z"}}}]
        )
        import_scores(students, org="VNU-HUS", classroom="c", runner=runner)
        states = getattr(students[0], FIELD_SUBMISSIONS)
        self.assertEqual(states["ch01-introduction"], STATE_UNKNOWN_STALE)
        self.assertEqual(states["w00-group-collaboration"], STATE_NOT_COLLECTED)
        # final-project is never graded upstream, so it gets no state at all.
        self.assertNotIn("final-project", states)

    def test_a_credited_student_reads_as_submitted(self):
        students = klass()
        runner = FakeRunner(
            scores=scores_doc(
                {
                    "ch01-introduction": {
                        "type": "individual",
                        "entries": [{"owner": "alice", "submissions": [payload()]}],
                    }
                }
            )
        )
        import_scores(students, org="VNU-HUS", classroom="c", runner=runner)
        self.assertEqual(
            getattr(students[0], FIELD_SUBMISSIONS)["ch01-introduction"], STATE_SUBMITTED
        )


class TestReruns(unittest.TestCase):
    def test_an_older_snapshot_does_not_walk_a_grade_back(self):
        students = klass()
        good = scores_doc(
            {
                "ch01-introduction": {
                    "type": "individual",
                    "entries": [{"owner": "alice", "submissions": [payload(score=95)]}],
                }
            }
        )
        import_scores(
            students, org="VNU-HUS", classroom="c", runner=FakeRunner(scores=good)
        )
        stale = FakeRunner(
            scores=scores_doc(
                {
                    "ch01-introduction": {
                        "type": "individual",
                        "entries": [{"owner": "alice", "submissions": [payload(score=1)]}],
                    }
                }
            ),
            commits=[{"commit": {"committer": {"date": "2026-08-10T04:52:10Z"}}}],
        )
        report = import_scores(students, org="VNU-HUS", classroom="c", runner=stale)
        self.assertEqual(getattr(students[0], FIELD_GRADES)["ch01-introduction"]["grade"], 95)
        self.assertTrue(report.stale_skipped)
        self.assertEqual(getattr(students[0], FIELD_COLLECTED_AT), COLLECTED_AT)

    def test_an_older_submission_does_not_replace_a_newer_one(self):
        students = klass()
        import_scores(
            students,
            org="VNU-HUS",
            classroom="c",
            runner=FakeRunner(
                scores=scores_doc(
                    {
                        "ch01-introduction": {
                            "type": "individual",
                            "entries": [
                                {
                                    "owner": "alice",
                                    "submissions": [
                                        payload(score=95, datetime="2026-09-11T10:00:00Z")
                                    ],
                                }
                            ],
                        }
                    }
                )
            ),
        )
        import_scores(
            students,
            org="VNU-HUS",
            classroom="c",
            runner=FakeRunner(
                scores=scores_doc(
                    {
                        "ch01-introduction": {
                            "type": "individual",
                            "entries": [
                                {
                                    "owner": "alice",
                                    "submissions": [
                                        payload(score=1, datetime="2026-09-01T10:00:00Z")
                                    ],
                                }
                            ],
                        }
                    }
                )
            ),
        )
        self.assertEqual(getattr(students[0], FIELD_GRADES)["ch01-introduction"]["grade"], 95)


class TestAssignmentSelection(unittest.TestCase):
    def test_final_project_reports_that_it_has_no_automatic_grade(self):
        students = klass()
        runner = FakeRunner()
        report = import_scores(
            students,
            org="VNU-HUS",
            classroom="c",
            assignment="final-project",
            runner=runner,
        )
        self.assertEqual(report.updated, 0)
        self.assertTrue(
            any("--load-override-grades" in w for w in report.warnings), report.warnings
        )

    def test_an_unknown_slug_names_the_slugs_that_do_exist(self):
        students = klass()
        with self.assertRaises(Classroom50Error) as ctx:
            import_scores(
                students,
                org="VNU-HUS",
                classroom="c",
                assignment="ch99-nope",
                runner=FakeRunner(),
            )
        self.assertEqual(ctx.exception.code, "unknown_assignment")
        self.assertIn("ch01-introduction", str(ctx.exception))
        self.assertEqual(written(students), 0)

    def test_selecting_one_slug_leaves_the_others_untouched(self):
        students = klass()
        runner = FakeRunner(
            scores=scores_doc(
                {
                    "ch01-introduction": {
                        "type": "individual",
                        "entries": [{"owner": "alice", "submissions": [payload()]}],
                    },
                    "w00-group-collaboration": {
                        "type": "group",
                        "entries": [
                            {
                                "owner": "bob",
                                "member_usernames": ["bob", "carol"],
                                "submissions": [payload(owner="bob", assignment_type="group")],
                            }
                        ],
                    },
                }
            )
        )
        import_scores(
            students,
            org="VNU-HUS",
            classroom="c",
            assignment="ch01-introduction",
            runner=runner,
        )
        self.assertEqual(
            list(getattr(students[0], FIELD_GRADES)), ["ch01-introduction"]
        )
        self.assertFalse(hasattr(students[1], FIELD_GRADES))


class TestManifest(unittest.TestCase):
    def test_grading_off_and_empty_repo_both_mean_ungraded(self):
        manifest = parse_assignments(MANIFEST)
        self.assertTrue(manifest["ch01-introduction"].graded)
        self.assertFalse(manifest["final-project"].graded)
        self.assertEqual(manifest["w00-group-collaboration"].mode, "group")

    def test_staff_roles_are_read_from_the_role_column_the_parser_drops(self):
        self.assertEqual(staff_logins([{"login": "x", "role": "HTA"}]), frozenset({"x"}))
        self.assertEqual(staff_logins([{"login": "y", "role": "ta"}]), frozenset({"y"}))
        self.assertEqual(staff_logins([{"login": "z", "role": ""}]), frozenset())
        self.assertEqual(staff_logins({"students": ROSTER}), frozenset({"hoanganhduc"}))


class TestModuleHygiene(unittest.TestCase):
    """What the lane is allowed to send, checked against the calls it makes.

    Grepping the source would flag the comment that names ``gh teacher`` as the
    source of ``STAFF_ROLES``; prose about a command is not a call to it.  These
    read the argv instead.
    """

    def all_calls(self) -> List[List[str]]:
        students = klass()
        runner = FakeRunner(
            scores=scores_doc(
                {
                    "ch01-introduction": {
                        "type": "individual",
                        "entries": [{"owner": "alice", "submissions": [payload()]}],
                    },
                    "w00-group-collaboration": {
                        "type": "group",
                        "entries": [
                            {
                                "owner": "bob",
                                "member_usernames": ["bob", "carol"],
                                "submissions": [payload(owner="bob", assignment_type="group")],
                            }
                        ],
                    },
                }
            )
        )
        import_scores(students, org="VNU-HUS", classroom="c", runner=runner)
        self.assertTrue(runner.calls)
        return runner.calls

    def test_every_call_is_a_read(self):
        for argv in self.all_calls():
            joined = " ".join(argv)
            for forbidden in ("--method", "-X PUT", "-X POST", "-X PATCH", "-X DELETE"):
                self.assertNotIn(forbidden, joined)
            for forbidden in ("dispatches", "workflow", "download", "zipball"):
                self.assertNotIn(forbidden, joined.lower())

    def test_every_call_is_one_of_the_four_known_read_shapes(self):
        # Two `gh api` reads of the config repo, plus the two `gh teacher` list
        # verbs AgentCLI already allows.  Anything else is a new capability and
        # should have to be added here deliberately.
        allowed_teacher = (["gh", "teacher", "roster", "list"], ["gh", "teacher", "assignment", "list"])
        for argv in self.all_calls():
            if argv[:2] == ["gh", "api"]:
                continue
            self.assertIn(argv[:4], allowed_teacher, argv)

    def test_the_config_repo_is_read_by_path_not_by_clone(self):
        api = [" ".join(a) for a in self.all_calls() if a[:2] == ["gh", "api"]]
        self.assertTrue(any("contents/c/scores.json" in c for c in api), api)
        self.assertTrue(any("commits?path=c/scores.json" in c for c in api), api)

    def test_no_paginate_aggregate_in_jq(self):
        # --jq runs once per page under --paginate, so an aggregate there
        # silently answers per page.  The lane avoids the pairing entirely.
        source = open(
            os.path.join(REPO_ROOT, "course_hoanganhduc", "c50_scores.py")
        ).read()
        self.assertNotIn("--paginate", source)


if __name__ == "__main__":
    unittest.main(verbosity=2)
