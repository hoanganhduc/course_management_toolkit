#!/usr/bin/env python3
"""Offline tests for the Classroom50 group lane.

The module under test is stdlib-only, never imports the database layer and never
imports pandas, so these tests run under a bare interpreter.  Every network call
goes through an injected runner, so nothing here reaches GitHub.

The fixtures are shaped like the two real classrooms: two ``mode=group``
assignments, one autograded and one with ``grading.mode=off``, an organization
whose two owners are collaborators on every repository, and usernames that
contain hyphens.
"""

from __future__ import annotations

import ast
import json
import sys
import unittest
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from course_hoanganhduc.c50_cli import Classroom50Error, RunResult  # noqa: E402
from course_hoanganhduc.c50_groups import (  # noqa: E402
    FIELD_GROUP,
    FIELD_GROUP_CONFLICTS,
    FIELD_MP_CONFLICT,
    FIELD_MP_MEMBERS,
    FIELD_MP_REPO,
    FIELD_MP_SOURCE,
    MAX_MEMBERS,
    SOURCE_COLLABORATORS,
    SOURCE_SCORES,
    Collaborator,
    Group,
    all_roster_logins,
    course_for_classroom,
    credit_members,
    detect_group_conflicts,
    format_groups,
    format_groups_txt,
    import_groups,
    list_groups,
    list_repo_collaborators,
    merge_groups_into_students,
    parse_group_repo_name,
    parse_team_json,
    report_to_dict,
    select_group_slugs,
    staff_team_slugs,
    student_members,
    timeout_runner,
)
from course_hoanganhduc.c50_scores import (  # noqa: E402
    FIELD_GRADE_CANDIDATES,
    FIELD_GRADES,
    parse_assignments,
)
from course_hoanganhduc.roster_audit import (  # noqa: E402
    SEVERITY_ERROR,
    SEVERITY_WARNING,
    AuditReport,
    format_announcement,
)

ORG = "VNU-HUS"
CLASSROOM = "vnu-hus-mat1206e-winter-2026"
GROUP_SLUG = "w00-group-collaboration"
FINAL_SLUG = "final-project"

# The two organization owners.  Both are collaborators on every repository and
# both are on the teacher team, which is why they survive crediting and are only
# removed by the staff subtraction.
TEACHERS = ("hoanganhduc", "Ivy")

GITHUB_IDS = {
    "alice": 111,
    "bob": 222,
    "carol": 333,
    "dave": 444,
    "erin": 555,
    "frank": 666,
    "24007008-debug": 888,
    "hoanganhduc": 700,
    "ivy": 777,
}

ROSTER = [
    {"login": "alice", "kind": "user", "role": "", "github_id": 111},
    {"login": "bob", "kind": "user", "role": "", "github_id": 222},
    {"login": "carol", "kind": "user", "role": "", "github_id": 333},
    {"login": "dave", "kind": "user", "role": "", "github_id": 444},
    {"login": "erin", "kind": "user", "role": "", "github_id": 555},
    {"login": "frank", "kind": "user", "role": "", "github_id": 666},
    {"login": "hoanganhduc", "kind": "user", "role": "teacher", "github_id": 700},
    {"login": "Ivy", "kind": "user", "role": "teacher", "github_id": 777},
]

MANIFEST = {
    "assignments": [
        {"slug": "ch01-introduction", "mode": "individual"},
        {
            "slug": GROUP_SLUG,
            "mode": "group",
            "max_group_size": 5,
            "available_from": "2026-09-06T23:00:00Z",
            "due": "2026-09-11T05:00:00Z",
        },
        {
            "slug": FINAL_SLUG,
            "mode": "group",
            "max_group_size": 5,
            "empty_repo": True,
            "grading": {"mode": "off"},
            "available_from": "2026-09-11T06:00:00Z",
            "due": "2026-11-04T16:59:00Z",
        },
    ]
}

# The slugs are recorded, not derived: GitHub renames a team whose name is taken,
# so a rebuilt slug can name a team that does not exist.
CLASSROOM_CONFIG = {
    "org": "vnu-hus",
    "team": {"id": 9000, "slug": "classroom50-vnu-hus-mat1206e-winter-2026"},
    "teams": {
        "teacher": {"id": 9001, "slug": "renamed-teacher-team-2"},
        "hta": {"id": 9003, "slug": "renamed-hta-team-2"},
        "ta": {"id": 9004, "slug": "renamed-ta-team-2"},
    },
}


def repo_name(slug: str, founder: str, classroom: str = CLASSROOM) -> str:
    return f"{classroom}-{slug}-{founder}"


W00_REPO = repo_name(GROUP_SLUG, "alice")
FINAL_REPO = repo_name(FINAL_SLUG, "alice")


def scores_doc(assignments: Dict[str, Any]) -> Dict[str, Any]:
    return {"schema": "classroom50/scores/v1", "assignments": assignments}


def group_bucket(owner: str, members: Sequence[str]) -> Dict[str, Any]:
    """One ``scores.json`` bucket for a group assignment."""
    return {
        "type": "group",
        "entries": [
            {
                "owner": owner,
                "member_usernames": list(members),
                "submissions": [
                    {
                        "schema": "classroom50/result/v1",
                        "classroom": CLASSROOM,
                        "owner": owner,
                        "assignment_type": "group",
                        "submission": "https://github.com/x",
                        "commit": "0" * 40,
                        "release": "v1",
                        "review": "https://github.com/x/pull/1",
                        "datetime": "2026-09-10T02:00:00Z",
                        "score": 90,
                        "max-score": 100,
                    }
                ],
            }
        ],
    }


def team_payload(**overrides: Any) -> Dict[str, Any]:
    payload: Dict[str, Any] = {
        "course": "MAT1206E",
        "group_name": "Nhóm Alpha",
        "founder": "alice",
        "members": [
            {
                "full_name": "Nguyễn Thị An",
                "student_id": "24001111",
                "github_username": "alice",
            },
            {
                "full_name": "Trần Văn Bình",
                "student_id": "24002222",
                "github_username": "bob",
            },
        ],
    }
    payload.update(overrides)
    return payload


def w00_team_payload(**overrides: Any) -> Dict[str, Any]:
    """A ``team.json`` in the shape ``check_submission.py`` demands.

    The w00 template ships its own checker, and that checker wants exactly
    ``team_name`` and ``members`` as bare logins.  It is a different file from
    the final project's, so it gets a different fixture rather than being made
    to stand in for one.
    """
    payload: Dict[str, Any] = {
        "team_name": "Nhóm Alpha",
        "members": ["alice", "bob"],
    }
    payload.update(overrides)
    return payload


class Student:
    """The attribute bag the database layer hands around."""

    def __init__(self, name: str, student_id: str, username: str, github_id: str = "") -> None:
        self.Name = name
        setattr(self, "Student ID", student_id)
        setattr(self, "GitHub Username", username)
        if github_id:
            setattr(self, "GitHub ID", github_id)
        self.Email = f"{student_id}@hus.edu.vn"


def klass() -> List[Student]:
    # Vietnamese names on purpose: ``refine_database`` deletes records whose name
    # looks like a test fixture.
    return [
        Student("Nguyễn Thị An", "24001111", "alice", "111"),
        Student("Trần Văn Bình", "24002222", "bob", "222"),
        Student("Lê Thị Cúc", "24003333", "carol", "333"),
        Student("Phạm Văn Dũng", "24004444", "dave", "444"),
        Student("Vũ Thị Én", "24005555", "erin", "555"),
        Student("Đỗ Văn Phong", "24006666", "frank", "666"),
    ]


def make_group(**overrides: Any) -> Group:
    fields: Dict[str, Any] = {
        "slug": GROUP_SLUG,
        "repo": W00_REPO,
        "founder": "alice",
        "members": ["alice", "bob"],
        "source": SOURCE_COLLABORATORS,
        "credited": ["alice", "bob"],
        "snapshot": ["alice", "bob"],
        "outsiders": [],
        "group_name": "",
    }
    fields.update(overrides)
    return Group(**fields)


class FakeRunner:
    """Answers the exact argv shapes the group lane builds, and nothing else."""

    def __init__(
        self,
        *,
        repos: Optional[Sequence[str]] = None,
        collaborators: Optional[Dict[str, Sequence[str]]] = None,
        team_json: Optional[Dict[str, Any]] = None,
        teams: Optional[Dict[str, Sequence[str]]] = None,
        roster: Optional[Sequence[Dict[str, Any]]] = None,
        scores: Optional[Dict[str, Any]] = None,
        config: Optional[Dict[str, Any]] = None,
    ) -> None:
        self.repos = list(repos) if repos is not None else [W00_REPO, FINAL_REPO]
        self.default_collaborators = ["alice", "bob", "carol", *TEACHERS]
        self.collaborators = dict(collaborators or {})
        self.team_json = dict(team_json or {})
        self.teams = dict(teams or {})
        self.roster = list(roster) if roster is not None else list(ROSTER)
        self.scores = scores if scores is not None else scores_doc({})
        self.config = config if config is not None else CLASSROOM_CONFIG
        self.assignments = MANIFEST
        self.failures: Dict[str, RunResult] = {}
        self.calls: List[List[str]] = []
        self.timeouts: List[Optional[float]] = []
        self.depth = 0
        self.max_depth = 0

    # -- helpers ---------------------------------------------------------
    def paths(self) -> List[str]:
        return [call[-1] for call in self.calls]

    def api_paths(self) -> List[str]:
        return [call[-1] for call in self.calls if call[1:2] == ["api"]]

    @staticmethod
    def _page(path: str, rows: Sequence[Any]) -> List[Any]:
        return list(rows) if "page=1" in path else []

    def _collab_rows(self, repo: str) -> List[Dict[str, Any]]:
        logins = self.collaborators.get(repo, self.default_collaborators)
        rows = []
        for login in logins:
            # ``role_name`` is present on purpose: nothing below may read it.
            role = "admin" if login in TEACHERS or login.endswith("-owner") else "write"
            rows.append(
                {"login": login, "id": GITHUB_IDS.get(login.lower(), 0), "role_name": role}
            )
        return rows

    # -- the runner protocol ---------------------------------------------
    def __call__(self, argv: Sequence[str], *, timeout: Optional[float] = None) -> RunResult:
        self.calls.append(list(argv))
        self.timeouts.append(timeout)
        self.depth += 1
        self.max_depth = max(self.max_depth, self.depth)
        try:
            return self._answer(list(argv))
        finally:
            self.depth -= 1

    def _answer(self, argv: List[str]) -> RunResult:
        joined = " ".join(argv)
        for marker, result in self.failures.items():
            if marker in joined:
                return result
        path = argv[-1]
        body: Any
        if "roster list" in joined:
            body = self.roster
        elif "assignment list" in joined:
            body = self.assignments
        elif "commits?path=" in joined:
            body = [{"commit": {"committer": {"date": "2026-09-10T03:00:00Z"}}}]
        elif "scores.json" in joined:
            body = self.scores
        elif "classroom.json" in joined:
            body = self.config
        elif "/contents/team.json" in joined:
            repo = path.split("/")[2]
            payload = self.team_json.get(repo)
            if payload is None:
                return RunResult(returncode=1, stderr="gh: Not Found (HTTP 404)")
            if isinstance(payload, str):
                return RunResult(returncode=0, stdout=payload)
            return RunResult(returncode=0, stdout=json.dumps(payload))
        elif "/teams/" in joined and "/members" in joined:
            slug = path.split("/")[4]
            rows = [
                {"login": login, "id": GITHUB_IDS.get(login.lower(), 0)}
                for login in self.teams.get(slug, [])
            ]
            body = self._page(path, rows)
        elif "/collaborators" in joined:
            body = self._page(path, self._collab_rows(path.split("/")[2]))
        elif "/repos" in joined:
            rows = [{"name": name, "full_name": f"{ORG}/{name}"} for name in self.repos]
            body = self._page(path, rows)
        else:
            return RunResult(returncode=1, stderr=f"unexpected call: {joined}")
        return RunResult(returncode=0, stdout=json.dumps(body))


class RecordingSleeper:
    def __init__(self) -> None:
        self.slept: List[float] = []

    def __call__(self, seconds: float) -> None:
        self.slept.append(seconds)


def run_import(runner: FakeRunner, students: Sequence[Any], **kwargs: Any) -> Any:
    """``import_groups`` with the clock and the pacing taken out."""
    kwargs.setdefault("sleeper", RecordingSleeper())
    kwargs.setdefault("pace", 0)
    return import_groups(
        students, org=ORG, classroom=CLASSROOM, runner=runner, **kwargs
    )


def group_of(student: Any, slug: str = GROUP_SLUG) -> Optional[Dict[str, Any]]:
    recorded = getattr(student, FIELD_GROUP, None)
    if not isinstance(recorded, dict):
        return None
    return recorded.get(slug)


def codes(findings: Sequence[Any]) -> List[str]:
    return sorted({issue.code for issue in findings})


# ---------------------------------------------------------------------------
# repository names
# ---------------------------------------------------------------------------


class TestRepoNames(unittest.TestCase):
    def test_founder_keeps_the_hyphens_in_its_username(self) -> None:
        classroom = "vnu-hus-mat3508-winter-2026"
        parsed = parse_group_repo_name(
            f"{classroom}-w00-individual-onboarding-24007008-debug",
            classroom,
            ["w00-individual-onboarding"],
        )
        assert parsed is not None
        self.assertEqual(parsed.founder, "24007008-debug")
        self.assertEqual(parsed.slug, "w00-individual-onboarding")

    def test_the_longest_matching_slug_wins(self) -> None:
        classroom = "vnu-hus-mat3508-winter-2026"
        parsed = parse_group_repo_name(
            f"{classroom}-w00-individual-onboarding-24007008-debug",
            classroom,
            ["w00-individual", "w00-individual-onboarding"],
        )
        assert parsed is not None
        self.assertEqual(parsed.slug, "w00-individual-onboarding")
        self.assertEqual(parsed.founder, "24007008-debug")

    def test_another_classrooms_repository_is_not_claimed(self) -> None:
        self.assertIsNone(
            parse_group_repo_name(
                "vnu-hus-mat3508-winter-2026-final-project-alice",
                CLASSROOM,
                [FINAL_SLUG],
            )
        )

    def test_an_unrelated_repository_is_not_claimed(self) -> None:
        self.assertIsNone(parse_group_repo_name("classroom50", CLASSROOM, [FINAL_SLUG]))

    def test_repository_names_are_folded_before_matching(self) -> None:
        parsed = parse_group_repo_name(
            f"{CLASSROOM.upper()}-{FINAL_SLUG}-Alice", CLASSROOM, [FINAL_SLUG]
        )
        assert parsed is not None
        self.assertEqual(parsed.founder, "alice")


# ---------------------------------------------------------------------------
# the two crediting steps
# ---------------------------------------------------------------------------


class TestCreditingSteps(unittest.TestCase):
    """Step 3 must equal upstream; step 4 is ours.  Both are asserted."""

    def setUp(self) -> None:
        self.roster_logins = all_roster_logins(ROSTER)
        self.staff = frozenset({"hoanganhduc", "ivy"})
        self.collaborators = [
            Collaborator(login="alice", github_id="111"),
            Collaborator(login="bob", github_id="222"),
            Collaborator(login="carol", github_id="333"),
            Collaborator(login="hoanganhduc", github_id="700"),
            Collaborator(login="ivy", github_id="777"),
        ]

    def test_step_three_keeps_the_teachers_and_step_four_removes_them(self) -> None:
        credited = credit_members(
            self.collaborators, roster_logins=self.roster_logins, owner="alice"
        )
        # Upstream credits anyone on the classroom teams, staff included.
        self.assertIn("hoanganhduc", credited)
        self.assertIn("ivy", credited)
        self.assertEqual(credited, ["alice", "bob", "carol", "hoanganhduc", "ivy"])

        students = student_members(credited, staff_logins=self.staff)
        self.assertEqual(students, ["alice", "bob", "carol"])

    def test_an_admin_collaborator_is_still_a_student(self) -> None:
        collaborators = list(self.collaborators) + [
            Collaborator(login="dave", github_id="444")
        ]
        credited = credit_members(
            collaborators, roster_logins=self.roster_logins, owner="alice"
        )
        self.assertIn("dave", credited)
        self.assertIn("dave", student_members(credited, staff_logins=self.staff))

    def test_the_owner_is_kept_even_off_the_roster(self) -> None:
        credited = credit_members([], roster_logins=frozenset({"bob"}), owner="Zoe")
        self.assertEqual(credited, ["zoe"])

    def test_someone_outside_the_roster_is_not_credited(self) -> None:
        collaborators = list(self.collaborators) + [
            Collaborator(login="outsider", github_id="999")
        ]
        credited = credit_members(
            collaborators, roster_logins=self.roster_logins, owner="alice"
        )
        self.assertNotIn("outsider", credited)

    def test_five_students_and_two_teachers_are_not_over_size(self) -> None:
        collaborators = [
            Collaborator(login=login, github_id=str(GITHUB_IDS[login]))
            for login in ("alice", "bob", "carol", "dave", "erin")
        ] + [
            Collaborator(login="hoanganhduc", github_id="700"),
            Collaborator(login="ivy", github_id="777"),
        ]
        credited = credit_members(
            collaborators, roster_logins=self.roster_logins, owner="alice"
        )
        self.assertEqual(len(credited), 7)
        members = student_members(credited, staff_logins=self.staff)
        self.assertEqual(len(members), MAX_MEMBERS)

        findings = detect_group_conflicts([make_group(members=members)], index={})
        self.assertNotIn("group_over_size", codes(findings))


# ---------------------------------------------------------------------------
# the shape of the calls
# ---------------------------------------------------------------------------


class TestCallShapes(unittest.TestCase):
    def test_the_collaborator_request_carries_no_membership_filter(self) -> None:
        runner = FakeRunner()
        list_repo_collaborators(ORG, W00_REPO, runner=runner, pace=0)
        path = runner.calls[0][-1]
        self.assertEqual(runner.calls[0][:2], ["gh", "api"])
        self.assertNotIn("affiliation", path)
        self.assertNotIn("permission", path)
        self.assertNotIn("role_name", path)
        self.assertEqual(path, f"repos/{ORG}/{W00_REPO}/collaborators?per_page=100&page=1")

    def test_every_call_is_one_of_four_known_shapes(self) -> None:
        runner = FakeRunner(team_json={W00_REPO: team_payload()})
        run_import(runner, klass(), with_team_json=True)
        for call in runner.calls:
            self.assertEqual(call[0], "gh")
            self.assertIn(call[1], {"api", "teacher"})
            if call[1] == "teacher":
                # The only two teacher verbs this lane may reach.
                self.assertIn(call[2], {"roster", "assignment"})
                self.assertEqual(call[3], "list")
            else:
                path = call[-1]
                self.assertTrue(
                    path.startswith((f"/orgs/{ORG}/", f"repos/{ORG}/")),
                    f"unexpected path: {path}",
                )
                self.assertNotIn("-X", call)
                self.assertNotIn("--method", call)
                self.assertNotIn("--paginate", call)

    def test_pages_are_walked_by_number(self) -> None:
        runner = FakeRunner()
        run_import(runner, klass())
        listing = [path for path in runner.api_paths() if "per_page=" in path]
        self.assertTrue(listing)
        for path in listing:
            self.assertIn("per_page=100&page=", path)
            self.assertNotIn("--paginate", path)

    def test_the_timeout_is_offered_to_every_call_of_the_repository_loop(self) -> None:
        # The loop is the hazard: one hung repository stalls the whole pass.  The
        # single gradebook read is bounded by the runner itself, which
        # ``test_a_hung_command_becomes_a_named_error`` covers.
        runner = FakeRunner(team_json={W00_REPO: team_payload()})
        run_import(runner, klass(), with_team_json=True, timeout=12.5)
        looped = [
            timeout
            for call, timeout in zip(runner.calls, runner.timeouts)
            if call[1:2] == ["api"]
            and any(
                part in call[-1]
                for part in ("/repos?", "/collaborators", "/contents/team.json")
            )
        ]
        # one listing + two collaborator reads + two team.json reads
        self.assertEqual(len(looped), 5)
        self.assertEqual(set(looped), {12.5})

    def test_the_repositories_are_read_one_at_a_time(self) -> None:
        runner = FakeRunner(
            team_json={W00_REPO: team_payload(), FINAL_REPO: team_payload()}
        )
        run_import(runner, klass(), with_team_json=True)
        self.assertEqual(runner.max_depth, 1)
        per_repo = [
            path
            for path in runner.api_paths()
            if "/collaborators" in path or "/contents/team.json" in path
        ]
        self.assertEqual(
            per_repo,
            [
                f"repos/{ORG}/{FINAL_REPO}/collaborators?per_page=100&page=1",
                f"repos/{ORG}/{FINAL_REPO}/contents/team.json",
                f"repos/{ORG}/{W00_REPO}/collaborators?per_page=100&page=1",
                f"repos/{ORG}/{W00_REPO}/contents/team.json",
            ],
        )


# ---------------------------------------------------------------------------
# the two routes
# ---------------------------------------------------------------------------


class TestTwoRoutes(unittest.TestCase):
    def test_the_final_project_alone_never_reads_the_gradebook(self) -> None:
        runner = FakeRunner(repos=[FINAL_REPO])
        report = run_import(runner, klass(), assignment=FINAL_SLUG)
        self.assertNotIn("scores.json", " ".join(runner.api_paths()))
        self.assertEqual(report.groups[FINAL_SLUG][0]["source"], SOURCE_COLLABORATORS)

    def test_the_graded_assignment_uses_the_published_member_list(self) -> None:
        runner = FakeRunner(
            repos=[W00_REPO],
            scores=scores_doc(
                {GROUP_SLUG: group_bucket("alice", ["alice", "bob", "hoanganhduc"])}
            ),
        )
        students = klass()
        report = run_import(runner, students, assignment=GROUP_SLUG)
        self.assertIn("scores.json", " ".join(runner.api_paths()))
        row = report.groups[GROUP_SLUG][0]
        self.assertEqual(row["source"], SOURCE_SCORES)
        # carol is a collaborator today but was not credited; the teacher is
        # dropped by the staff subtraction, not by the route.
        self.assertEqual(row["members"], ["alice", "bob"])
        self.assertIsNone(group_of(students[2]))

    def test_one_run_takes_both_routes(self) -> None:
        runner = FakeRunner(
            scores=scores_doc({GROUP_SLUG: group_bucket("alice", ["alice", "bob"])})
        )
        report = run_import(runner, klass())
        sources = {slug: rows[0]["source"] for slug, rows in report.groups.items()}
        self.assertEqual(
            sources, {GROUP_SLUG: SOURCE_SCORES, FINAL_SLUG: SOURCE_COLLABORATORS}
        )

    def test_different_groups_in_the_two_assignments_are_not_a_conflict(self) -> None:
        final_repo = repo_name(FINAL_SLUG, "carol")
        runner = FakeRunner(
            repos=[W00_REPO, final_repo],
            collaborators={
                W00_REPO: ["alice", "bob", *TEACHERS],
                final_repo: ["carol", "dave", *TEACHERS],
            },
        )
        students = klass()
        report = run_import(runner, students)
        self.assertEqual(report.findings, [])
        self.assertEqual(group_of(students[0])["members"], ["alice", "bob"])
        self.assertEqual(
            group_of(students[2], FINAL_SLUG)["members"], ["carol", "dave"]
        )

    def test_an_unknown_assignment_is_refused_by_name(self) -> None:
        manifest = parse_assignments(MANIFEST)
        with self.assertRaises(Classroom50Error) as caught:
            select_group_slugs(manifest, "ch01-introduction")
        self.assertEqual(caught.exception.code, "unknown_assignment")
        self.assertIn(FINAL_SLUG, str(caught.exception))

    def test_no_assignment_names_both_group_slugs(self) -> None:
        self.assertEqual(
            select_group_slugs(parse_assignments(MANIFEST)),
            sorted([FINAL_SLUG, GROUP_SLUG]),
        )

    def test_a_classroom_without_group_work_lists_no_repository(self) -> None:
        runner = FakeRunner()
        runner.assignments = {"assignments": [{"slug": "ch01-introduction", "mode": "individual"}]}
        report = run_import(runner, klass())
        self.assertEqual(report.slugs, [])
        self.assertIn("this classroom has no group assignment", report.warnings)
        self.assertNotIn("/repos", " ".join(runner.api_paths()))


# ---------------------------------------------------------------------------
# who counts as staff
# ---------------------------------------------------------------------------


class TestStaffResolution(unittest.TestCase):
    def test_the_roster_answers_first_and_no_team_is_read(self) -> None:
        runner = FakeRunner()
        run_import(runner, klass())
        joined = " ".join(runner.api_paths())
        self.assertNotIn("/teams/", joined)
        self.assertNotIn("classroom.json", joined)

    def test_the_recorded_team_slug_is_read_never_rebuilt(self) -> None:
        roster = [dict(row, role="") for row in ROSTER]
        runner = FakeRunner(
            roster=roster,
            teams={
                "renamed-teacher-team-2": ["hoanganhduc", "Ivy"],
                "renamed-hta-team-2": [],
                "renamed-ta-team-2": [],
            },
        )
        students = klass()
        report = run_import(runner, students)
        self.assertEqual(
            staff_team_slugs(CLASSROOM_CONFIG),
            ["renamed-teacher-team-2", "renamed-hta-team-2", "renamed-ta-team-2"],
        )
        paths = " ".join(runner.api_paths())
        self.assertIn(f"/orgs/{ORG}/teams/renamed-teacher-team-2/members", paths)
        self.assertNotIn("classroom50-vnu-hus-mat1206e-winter-2026/members", paths)
        self.assertTrue(any("roster names no teacher" in w for w in report.warnings))
        # The teachers were still removed, so the group is the three students.
        self.assertEqual(group_of(students[0])["members"], ["alice", "bob", "carol"])

    def test_a_configuration_without_teams_warns_instead_of_guessing(self) -> None:
        roster = [dict(row, role="") for row in ROSTER]
        runner = FakeRunner(roster=roster, config={"org": "vnu-hus"})
        report = run_import(runner, klass())
        self.assertEqual(staff_team_slugs({"org": "vnu-hus"}), [])
        self.assertNotIn("/teams/", " ".join(runner.api_paths()))
        self.assertTrue(any("roster names no teacher" in w for w in report.warnings))


# ---------------------------------------------------------------------------
# team.json
# ---------------------------------------------------------------------------


class TestTeamJson(unittest.TestCase):
    def fault_codes(self, payload: Any) -> List[str]:
        team = parse_team_json(payload, repo=W00_REPO)
        return sorted({fault.code for fault in team.faults})

    def test_a_complete_file_has_no_fault(self) -> None:
        team = parse_team_json(team_payload(), repo=W00_REPO)
        self.assertEqual(team.faults, [])
        self.assertEqual(team.course, "MAT1206E")
        self.assertEqual(team.founder, "alice")
        self.assertEqual([m.github_username for m in team.members], ["alice", "bob"])

    def test_a_list_is_not_a_team_file(self) -> None:
        self.assertEqual(self.fault_codes([]), ["team_json_invalid"])

    def test_a_missing_key_and_an_extra_key_are_both_named(self) -> None:
        payload = team_payload()
        del payload["group_name"]
        payload["extra"] = 1
        self.assertIn("team_json_keys", self.fault_codes(payload))

    def test_an_unknown_course_is_a_fault(self) -> None:
        self.assertIn("team_json_course", self.fault_codes(team_payload(course="MAT9999")))

    def test_six_members_are_too_many(self) -> None:
        members = [
            {
                "full_name": f"Sinh viên {n}",
                "student_id": f"2400{n}{n}{n}{n}",
                "github_username": f"user{n}",
            }
            for n in range(1, 7)
        ]
        payload = team_payload(members=members, founder="user1")
        self.assertIn("team_json_size", self.fault_codes(payload))

    def test_a_username_that_github_could_not_hold_is_a_fault(self) -> None:
        payload = team_payload()
        payload["members"][1]["github_username"] = "-bad-name-"
        self.assertIn("team_json_username", self.fault_codes(payload))

    def test_a_founder_outside_the_member_list_is_a_fault(self) -> None:
        self.assertIn(
            "team_json_founder_absent", self.fault_codes(team_payload(founder="carol"))
        )

    def test_an_unfilled_template_is_a_fault(self) -> None:
        payload = team_payload()
        payload["members"][1]["full_name"] = "THAY BANG HO TEN"
        self.assertIn("team_json_placeholder", self.fault_codes(payload))

    def test_a_repeated_username_is_a_fault(self) -> None:
        payload = team_payload()
        payload["members"][1]["github_username"] = "Alice"
        self.assertIn("team_json_duplicate_member", self.fault_codes(payload))

    def test_a_repeated_student_number_is_a_fault(self) -> None:
        payload = team_payload()
        payload["members"][1]["student_id"] = "24001111"
        self.assertIn("team_json_duplicate_student_id", self.fault_codes(payload))

    def test_a_missing_member_key_is_a_fault(self) -> None:
        payload = team_payload()
        del payload["members"][1]["student_id"]
        self.assertIn("team_json_member_keys", self.fault_codes(payload))

    def test_an_absent_file_is_skipped_not_failed(self) -> None:
        runner = FakeRunner(repos=[W00_REPO])
        report = run_import(runner, klass(), with_team_json=True)
        self.assertEqual(report.skipped_team_json, [W00_REPO])
        self.assertEqual(report.findings, [])
        self.assertEqual(report.updated, 3)

    def test_unparseable_content_becomes_a_named_error(self) -> None:
        runner = FakeRunner(repos=[W00_REPO], team_json={W00_REPO: "{not json"})
        students = klass()
        report = run_import(runner, students, with_team_json=True)
        self.assertIn("team_json_invalid", codes(report.findings))
        self.assertEqual(report.updated, 0)
        self.assertIsNone(group_of(students[0]))


# ---------------------------------------------------------------------------
# conflicts
# ---------------------------------------------------------------------------


class TestConflicts(unittest.TestCase):
    def test_a_student_on_two_repositories_holds_both_groups(self) -> None:
        second = repo_name(GROUP_SLUG, "carol")
        runner = FakeRunner(
            repos=[W00_REPO, second],
            collaborators={
                W00_REPO: ["alice", "bob", *TEACHERS],
                second: ["carol", "bob", *TEACHERS],
            },
        )
        students = klass()
        report = run_import(runner, students, assignment=GROUP_SLUG)
        self.assertIn("group_multi_repo", codes(report.findings))
        self.assertEqual(sorted(report.quarantined[GROUP_SLUG]), sorted([W00_REPO, second]))
        for student in students[:3]:
            self.assertIsNone(group_of(student))
        conflicts = getattr(students[1], FIELD_GROUP_CONFLICTS)
        self.assertTrue(conflicts[GROUP_SLUG])

    def test_a_conflict_in_one_group_leaves_the_rest_of_the_class_written(self) -> None:
        held = repo_name(GROUP_SLUG, "carol")
        clean = repo_name(GROUP_SLUG, "erin")
        runner = FakeRunner(
            repos=[W00_REPO, held, clean],
            collaborators={
                W00_REPO: ["alice", "bob", *TEACHERS],
                held: ["carol", "bob", *TEACHERS],
                clean: ["erin", "frank", *TEACHERS],
            },
        )
        students = klass()
        report = run_import(runner, students, assignment=GROUP_SLUG)
        self.assertEqual(group_of(students[4])["members"], ["erin", "frank"])
        self.assertEqual(report.updated, 2)

    def test_six_students_are_held_as_over_size(self) -> None:
        runner = FakeRunner(
            repos=[W00_REPO],
            collaborators={
                W00_REPO: ["alice", "bob", "carol", "dave", "erin", "frank", *TEACHERS]
            },
        )
        students = klass()
        setattr(students[0], FIELD_GROUP, {GROUP_SLUG: {"repo": "old", "members": []}})
        report = run_import(runner, students, assignment=GROUP_SLUG)
        over = [f for f in report.findings if f.code == "group_over_size"]
        self.assertEqual(len(over), 1)
        self.assertEqual(over[0].severity, SEVERITY_ERROR)
        self.assertIn(str(MAX_MEMBERS), over[0].detail)
        self.assertEqual(report.updated, 0)
        # The earlier value is kept, not deleted.
        self.assertEqual(group_of(students[0])["repo"], "old")

    def test_an_outsider_is_dropped_but_the_group_is_written(self) -> None:
        runner = FakeRunner(
            repos=[W00_REPO],
            collaborators={W00_REPO: ["alice", "bob", "stranger", *TEACHERS]},
        )
        students = klass()
        report = run_import(runner, students, assignment=GROUP_SLUG)
        outsider = [f for f in report.findings if f.code == "group_outsider"]
        self.assertEqual(len(outsider), 1)
        self.assertEqual(outsider[0].severity, SEVERITY_WARNING)
        self.assertEqual(outsider[0].found, "stranger")
        self.assertEqual(group_of(students[0])["members"], ["alice", "bob"])
        self.assertEqual(report.quarantined, {})

    def test_a_declared_member_who_is_not_a_collaborator_holds_the_group(self) -> None:
        payload = team_payload()
        payload["members"].append(
            {
                "full_name": "Lê Thị Cúc",
                "student_id": "24003333",
                "github_username": "carol",
            }
        )
        runner = FakeRunner(
            repos=[FINAL_REPO],
            collaborators={FINAL_REPO: ["alice", "bob", *TEACHERS]},
            team_json={FINAL_REPO: payload},
        )
        students = klass()
        report = run_import(runner, students, assignment=FINAL_SLUG, with_team_json=True)
        self.assertIn("team_json_extra_member", codes(report.findings))
        self.assertEqual(report.updated, 0)

    def test_a_collaborator_left_out_of_the_file_holds_the_group(self) -> None:
        payload = team_payload(members=[team_payload()["members"][0]])
        runner = FakeRunner(
            repos=[FINAL_REPO],
            collaborators={FINAL_REPO: ["alice", "bob", *TEACHERS]},
            team_json={FINAL_REPO: payload},
        )
        report = run_import(runner, klass(), assignment=FINAL_SLUG, with_team_json=True)
        self.assertIn("team_json_missing_member", codes(report.findings))
        self.assertEqual(report.updated, 0)

    def test_a_founder_that_disagrees_with_the_repository_holds_the_group(self) -> None:
        runner = FakeRunner(
            repos=[FINAL_REPO],
            collaborators={FINAL_REPO: ["alice", "bob", *TEACHERS]},
            team_json={FINAL_REPO: team_payload(founder="bob")},
        )
        report = run_import(runner, klass(), assignment=FINAL_SLUG, with_team_json=True)
        founder = [f for f in report.findings if f.code == "founder_mismatch"]
        self.assertEqual(len(founder), 1)
        self.assertEqual(founder[0].found, "bob")
        self.assertEqual(report.updated, 0)

    def test_the_wrong_course_holds_the_group(self) -> None:
        runner = FakeRunner(
            repos=[FINAL_REPO],
            collaborators={FINAL_REPO: ["alice", "bob", *TEACHERS]},
            team_json={FINAL_REPO: team_payload(course="MAT3508")},
        )
        report = run_import(runner, klass(), assignment=FINAL_SLUG, with_team_json=True)
        self.assertIn("course_mismatch", codes(report.findings))
        self.assertEqual(course_for_classroom(CLASSROOM), "MAT1206E")
        self.assertEqual(report.updated, 0)

    def test_one_student_number_in_two_files_holds_both_groups(self) -> None:
        second = repo_name(FINAL_SLUG, "carol")
        first_file = team_payload(members=[team_payload()["members"][0]])
        second_file = team_payload(
            founder="carol",
            group_name="Nhóm Beta",
            members=[
                {
                    "full_name": "Lê Thị Cúc",
                    "student_id": "24001111",
                    "github_username": "carol",
                }
            ],
        )
        runner = FakeRunner(
            repos=[FINAL_REPO, second],
            collaborators={
                FINAL_REPO: ["alice", *TEACHERS],
                second: ["carol", *TEACHERS],
            },
            team_json={FINAL_REPO: first_file, second: second_file},
        )
        report = run_import(runner, klass(), assignment=FINAL_SLUG, with_team_json=True)
        self.assertEqual(codes(report.findings), ["student_id_shared_groups"])
        self.assertEqual(
            sorted(report.quarantined[FINAL_SLUG]), sorted([FINAL_REPO, second])
        )

    def test_a_shared_group_name_is_only_a_warning(self) -> None:
        second = repo_name(FINAL_SLUG, "carol")
        first_file = team_payload(members=[team_payload()["members"][0]])
        second_file = team_payload(
            founder="carol",
            members=[
                {
                    "full_name": "Lê Thị Cúc",
                    "student_id": "24003333",
                    "github_username": "carol",
                }
            ],
        )
        runner = FakeRunner(
            repos=[FINAL_REPO, second],
            collaborators={
                FINAL_REPO: ["alice", *TEACHERS],
                second: ["carol", *TEACHERS],
            },
            team_json={FINAL_REPO: first_file, second: second_file},
        )
        students = klass()
        report = run_import(runner, students, assignment=FINAL_SLUG, with_team_json=True)
        self.assertEqual(codes(report.findings), ["group_name_shared"])
        self.assertEqual(
            [f.severity for f in report.findings], [SEVERITY_WARNING]
        )
        self.assertEqual(report.quarantined, {})
        self.assertEqual(report.updated, 2)
        self.assertEqual(group_of(students[0], FINAL_SLUG)["group_name"], "Nhóm Alpha")

    def test_the_sheet_lane_disagreeing_is_a_warning_that_writes_nothing_back(self) -> None:
        runner = FakeRunner(
            repos=[W00_REPO], collaborators={W00_REPO: ["alice", "bob", *TEACHERS]}
        )
        students = klass()
        setattr(students[0], "MiniProject_Group_Members", "An, Cúc")
        report = run_import(
            runner,
            students,
            assignment=GROUP_SLUG,
            sheet_groups={W00_REPO: ["alice", "carol"]},
        )
        mismatch = [f for f in report.findings if f.code == "sheet_group_mismatch"]
        self.assertEqual(len(mismatch), 1)
        self.assertEqual(mismatch[0].severity, SEVERITY_WARNING)
        self.assertEqual(report.quarantined, {})
        self.assertEqual(group_of(students[0])["members"], ["alice", "bob"])
        self.assertEqual(getattr(students[0], "MiniProject_Group_Members"), "An, Cúc")

    def test_a_membership_change_since_the_collection_is_a_warning(self) -> None:
        runner = FakeRunner(
            repos=[W00_REPO],
            collaborators={W00_REPO: ["alice", "bob", "carol", *TEACHERS]},
            scores=scores_doc({GROUP_SLUG: group_bucket("alice", ["alice", "bob"])}),
        )
        students = klass()
        report = run_import(runner, students, assignment=GROUP_SLUG)
        changed = [
            f for f in report.findings if f.code == "membership_changed_since_collect"
        ]
        self.assertEqual(len(changed), 1)
        self.assertEqual(changed[0].severity, SEVERITY_WARNING)
        self.assertIn("carol", changed[0].detail)
        self.assertEqual(report.quarantined, {})
        self.assertEqual(group_of(students[0])["members"], ["alice", "bob"])

    def test_a_gradebook_crediting_only_the_founder_names_the_others(self) -> None:
        # The same upstream state the scores lane reports as
        # ``group_credit_owner_only``: here it surfaces as the membership drift
        # that names the four students the collector did not pay.
        runner = FakeRunner(
            repos=[W00_REPO],
            collaborators={
                W00_REPO: ["alice", "bob", "carol", "dave", "erin", *TEACHERS]
            },
            scores=scores_doc({GROUP_SLUG: group_bucket("alice", ["alice"])}),
        )
        students = klass()
        report = run_import(runner, students, assignment=GROUP_SLUG)
        changed = [
            f for f in report.findings if f.code == "membership_changed_since_collect"
        ]
        self.assertEqual(len(changed), 1)
        for login in ("bob", "carol", "dave", "erin"):
            self.assertIn(login, changed[0].detail)
        self.assertEqual(group_of(students[0])["members"], ["alice"])

    def test_a_member_lost_since_the_last_run_holds_the_group(self) -> None:
        runner = FakeRunner(
            repos=[W00_REPO], collaborators={W00_REPO: ["alice", "bob", *TEACHERS]}
        )
        students = klass()
        report = run_import(
            runner,
            students,
            assignment=GROUP_SLUG,
            previous={GROUP_SLUG: {W00_REPO: ["alice", "bob", "carol"]}},
        )
        dropped = [f for f in report.findings if f.code == "group_member_dropped"]
        self.assertEqual(len(dropped), 1)
        self.assertEqual(dropped[0].severity, SEVERITY_ERROR)
        self.assertEqual(dropped[0].found, "carol")
        self.assertEqual(report.updated, 0)

    def test_the_first_run_reports_nobody_as_dropped(self) -> None:
        runner = FakeRunner(
            repos=[W00_REPO], collaborators={W00_REPO: ["alice", "bob", *TEACHERS]}
        )
        report = run_import(runner, klass(), assignment=GROUP_SLUG, previous=None)
        self.assertNotIn("group_member_dropped", codes(report.findings))

    def test_every_finding_names_something_the_student_can_do(self) -> None:
        second = repo_name(GROUP_SLUG, "carol")
        runner = FakeRunner(
            repos=[W00_REPO, second],
            collaborators={
                W00_REPO: ["alice", "bob", "carol", "dave", "erin", "frank", "stranger", *TEACHERS],
                second: ["carol", *TEACHERS],
            },
            team_json={W00_REPO: team_payload(course="MAT3508", founder="bob")},
        )
        report = run_import(
            runner,
            klass(),
            assignment=GROUP_SLUG,
            with_team_json=True,
            previous={GROUP_SLUG: {second: ["carol", "dave"]}},
        )
        self.assertGreaterEqual(len(report.findings), 4)
        for issue in report.findings:
            self.assertTrue(issue.fix.strip(), issue.code)
            self.assertTrue(issue.detail.strip(), issue.code)
            self.assertIn(issue.severity, {SEVERITY_ERROR, SEVERITY_WARNING})

    def test_a_warning_writes_and_an_error_holds(self) -> None:
        groups = [
            make_group(outsiders=["stranger"]),
            make_group(repo=repo_name(GROUP_SLUG, "erin"), founder="erin",
                       members=["erin", "frank", "dave", "carol", "bob", "alice"]),
        ]
        findings = detect_group_conflicts(groups, index={})
        severities = {f.code: f.severity for f in findings}
        self.assertEqual(severities["group_outsider"], SEVERITY_WARNING)
        self.assertEqual(severities["group_over_size"], SEVERITY_ERROR)


# ---------------------------------------------------------------------------
# writing the records
# ---------------------------------------------------------------------------


class TestWrites(unittest.TestCase):
    def test_a_clean_final_project_group_fills_the_mini_project_fields(self) -> None:
        runner = FakeRunner(
            repos=[FINAL_REPO], collaborators={FINAL_REPO: ["alice", "bob", *TEACHERS]}
        )
        students = klass()
        run_import(runner, students, assignment=FINAL_SLUG)
        alice = students[0]
        self.assertEqual(getattr(alice, FIELD_MP_REPO), FINAL_REPO)
        self.assertEqual(getattr(alice, FIELD_MP_MEMBERS), "alice, bob")
        self.assertEqual(getattr(alice, FIELD_MP_SOURCE), SOURCE_COLLABORATORS)
        self.assertEqual(getattr(alice, FIELD_MP_CONFLICT), "")
        self.assertEqual(group_of(alice, FINAL_SLUG)["founder"], "alice")

    def test_a_held_final_project_group_records_the_reason_and_keeps_the_value(self) -> None:
        runner = FakeRunner(
            repos=[FINAL_REPO],
            collaborators={
                FINAL_REPO: ["alice", "bob", "carol", "dave", "erin", "frank", *TEACHERS]
            },
        )
        students = klass()
        setattr(students[0], FIELD_GROUP, {FINAL_SLUG: {"repo": "kept", "members": []}})
        run_import(runner, students, assignment=FINAL_SLUG)
        alice = students[0]
        self.assertEqual(group_of(alice, FINAL_SLUG)["repo"], "kept")
        self.assertFalse(hasattr(alice, FIELD_MP_REPO))
        self.assertTrue(getattr(alice, FIELD_MP_CONFLICT).startswith("group_over_size"))
        self.assertTrue(getattr(alice, FIELD_GROUP_CONFLICTS)[FINAL_SLUG])

    def test_nothing_is_ever_deleted_by_a_rerun(self) -> None:
        runner = FakeRunner(
            repos=[W00_REPO], collaborators={W00_REPO: ["alice", "bob", *TEACHERS]}
        )
        students = klass()
        run_import(runner, students, assignment=GROUP_SLUG)
        setattr(students[0], "Name", "Nguyễn Thị An")
        run_import(runner, students, assignment=GROUP_SLUG)
        self.assertEqual(group_of(students[0])["members"], ["alice", "bob"])

    def test_the_numeric_id_survives_a_renamed_account(self) -> None:
        roster = [
            dict(row, login="alice-new") if row["login"] == "alice" else row
            for row in ROSTER
        ]
        new_repo = repo_name(GROUP_SLUG, "alice-new")
        runner = FakeRunner(
            repos=[new_repo],
            roster=roster,
            collaborators={new_repo: ["alice-new", "bob", *TEACHERS]},
        )
        runner.default_collaborators = ["alice-new", "bob", *TEACHERS]
        GITHUB_IDS["alice-new"] = 111
        try:
            students = klass()  # the database still says "alice"
            report = run_import(runner, students, assignment=GROUP_SLUG)
        finally:
            GITHUB_IDS.pop("alice-new", None)
        self.assertIn("alice-new", report.matched)
        self.assertEqual(group_of(students[0])["members"], ["alice-new", "bob"])

    def test_a_roster_without_ids_falls_back_to_usernames_with_a_warning(self) -> None:
        roster = [
            {key: value for key, value in row.items() if key != "github_id"}
            for row in ROSTER
        ]
        runner = FakeRunner(
            repos=[W00_REPO],
            roster=roster,
            collaborators={W00_REPO: ["alice", "bob", *TEACHERS]},
        )
        students = klass()
        report = run_import(runner, students, assignment=GROUP_SLUG)
        self.assertTrue(any("matching by username only" in w for w in report.warnings))
        self.assertEqual(group_of(students[0])["members"], ["alice", "bob"])

    def test_a_login_with_no_record_is_unmatched_and_creates_nobody(self) -> None:
        runner = FakeRunner(
            repos=[W00_REPO], collaborators={W00_REPO: ["alice", "erin", *TEACHERS]}
        )
        students = klass()[:3]
        report = run_import(runner, students, assignment=GROUP_SLUG)
        self.assertIn("erin", report.unmatched)
        self.assertIn("alice", report.matched)
        self.assertEqual(len(students), 3)

    def test_listing_groups_writes_to_nobody(self) -> None:
        runner = FakeRunner()
        payload = list_groups(org=ORG, classroom=CLASSROOM, runner=runner, pace=0)
        self.assertEqual(payload["updated"], 0)
        self.assertEqual(sorted(payload["slugs"]), sorted([FINAL_SLUG, GROUP_SLUG]))
        self.assertIn("alice", payload["unmatched"])

    def test_the_report_serialises_its_findings(self) -> None:
        runner = FakeRunner(
            repos=[W00_REPO],
            collaborators={W00_REPO: ["alice", "bob", "stranger", *TEACHERS]},
        )
        report = run_import(runner, klass(), assignment=GROUP_SLUG)
        payload = report_to_dict(report)
        self.assertEqual(payload["findings"][0]["code"], "group_outsider")
        self.assertIn("student_id", payload["findings"][0]["student"])
        json.dumps(payload)
        self.assertIn("group(s)", format_groups(report))


# ---------------------------------------------------------------------------
# the announcement
# ---------------------------------------------------------------------------


class TestAnnouncement(unittest.TestCase):
    def test_errors_reach_the_notice_and_warnings_do_not(self) -> None:
        groups = [
            make_group(
                members=["alice", "bob", "carol", "dave", "erin", "frank"],
                outsiders=["stranger"],
            )
        ]
        findings = detect_group_conflicts(groups, index={})
        self.assertEqual(
            codes(findings), ["group_outsider", "group_over_size"]
        )
        report = AuditReport(
            total=6, with_form=6, missing_form=[], issues=findings, duplicate_accounts=[]
        )
        text = format_announcement(report, course_name="MAT1206E")
        self.assertIn(W00_REPO, text)
        self.assertNotIn("stranger", text)


# ---------------------------------------------------------------------------
# failures
# ---------------------------------------------------------------------------


class TestFailures(unittest.TestCase):
    def test_a_secondary_rate_limit_is_retried_and_then_named(self) -> None:
        runner = FakeRunner(repos=[W00_REPO])
        runner.failures["/collaborators"] = RunResult(
            returncode=1,
            stderr="gh: API rate limit exceeded for user ID 700 (HTTP 403)",
        )
        sleeper = RecordingSleeper()
        with self.assertRaises(Classroom50Error) as caught:
            import_groups(
                klass(),
                org=ORG,
                classroom=CLASSROOM,
                assignment=GROUP_SLUG,
                runner=runner,
                sleeper=sleeper,
                pace=0,
            )
        self.assertEqual(caught.exception.code, "gh_rate_limited")
        attempts = [p for p in runner.api_paths() if "/collaborators" in p]
        self.assertEqual(len(attempts), 4)
        self.assertEqual(sleeper.slept, [1.0, 2.0, 4.0])

    def test_a_retry_after_header_is_obeyed(self) -> None:
        runner = FakeRunner(repos=[W00_REPO])
        runner.failures["/collaborators"] = RunResult(
            returncode=1,
            stderr="gh: You have exceeded a secondary rate limit. retry-after: 7",
        )
        sleeper = RecordingSleeper()
        with self.assertRaises(Classroom50Error):
            import_groups(
                klass(),
                org=ORG,
                classroom=CLASSROOM,
                assignment=GROUP_SLUG,
                runner=runner,
                sleeper=sleeper,
                pace=0,
            )
        self.assertEqual(sleeper.slept, [7.0, 7.0, 7.0])

    def test_a_permission_denial_is_not_retried(self) -> None:
        runner = FakeRunner(repos=[W00_REPO])
        runner.failures["/collaborators"] = RunResult(
            returncode=1,
            stderr="gh: Must have admin rights to Repository. (HTTP 403)",
        )
        sleeper = RecordingSleeper()
        with self.assertRaises(Classroom50Error) as caught:
            import_groups(
                klass(),
                org=ORG,
                classroom=CLASSROOM,
                assignment=GROUP_SLUG,
                runner=runner,
                sleeper=sleeper,
                pace=0,
            )
        self.assertTrue(caught.exception.code.endswith("_failed"))
        attempts = [p for p in runner.api_paths() if "/collaborators" in p]
        self.assertEqual(len(attempts), 1)
        self.assertEqual(sleeper.slept, [])

    def test_a_partial_page_with_a_failing_exit_code_is_an_error(self) -> None:
        runner = FakeRunner()
        runner.failures[f"/orgs/{ORG}/repos"] = RunResult(
            returncode=1,
            stdout=json.dumps([{"name": W00_REPO}])
            + '{"message":"API rate limit exceeded","status":"403"}',
            stderr="gh: API error",
        )
        students = klass()
        with self.assertRaises(Classroom50Error):
            run_import(runner, students, assignment=GROUP_SLUG)
        for student in students:
            self.assertIsNone(group_of(student))

    def test_an_error_object_with_a_clean_exit_is_caught_by_the_type_check(self) -> None:
        runner = FakeRunner()
        runner.failures[f"/orgs/{ORG}/repos"] = RunResult(
            returncode=0,
            stdout=json.dumps(
                {"message": "Not Found", "documentation_url": "https://x", "status": "404"}
            ),
        )
        with self.assertRaises(Classroom50Error) as caught:
            run_import(runner, klass(), assignment=GROUP_SLUG)
        self.assertEqual(caught.exception.code, "bad_json")

    def test_a_hung_command_becomes_a_named_error(self) -> None:
        runner = timeout_runner(0.3)
        with self.assertRaises(Classroom50Error) as caught:
            runner([sys.executable, "-c", "import time; time.sleep(5)"])
        self.assertEqual(caught.exception.code, "gh_timeout")

    def test_a_missing_binary_becomes_a_named_error(self) -> None:
        runner = timeout_runner(1.0)
        with self.assertRaises(Classroom50Error) as caught:
            runner(["gh-teacher-does-not-exist", "api", "/orgs"])
        self.assertEqual(caught.exception.code, "missing_binary")


# ---------------------------------------------------------------------------
# hygiene
# ---------------------------------------------------------------------------


class TestLongFormExport(unittest.TestCase):
    """The file a teacher keeps: one block per group, everything that is known.

    The point of the format is the join.  Who a login is comes from the
    database, what the group says about itself comes from ``team.json``, and
    what it was marked comes from the gradebook; the tests below are mostly
    about each of those three surviving the trip into the file.
    """

    def export(
        self,
        *,
        students: Optional[Sequence[Any]] = None,
        **runner_kwargs: Any,
    ) -> str:
        records = klass() if students is None else students
        options: Dict[str, Any] = {
            "repos": [W00_REPO],
            "scores": scores_doc({GROUP_SLUG: group_bucket("alice", ["alice", "bob"])}),
            "team_json": {W00_REPO: w00_team_payload()},
        }
        options.update(runner_kwargs)
        report = run_import(FakeRunner(**options), records, with_team_json=True)
        return format_groups_txt(
            report, org=ORG, classroom=CLASSROOM, students=records
        )

    def test_the_repository_and_the_address_to_open_it_are_both_there(self) -> None:
        text = self.export()
        self.assertIn(W00_REPO, text)
        self.assertIn(f"https://github.com/{ORG}/{W00_REPO}", text)

    def test_a_member_carries_the_name_and_number_the_database_has(self) -> None:
        text = self.export()
        self.assertIn("Nguyễn Thị An", text)
        self.assertIn("24001111", text)

    def test_the_founder_is_marked_among_the_members(self) -> None:
        text = self.export()
        self.assertIn("@alice", text)
        self.assertIn("[người tạo kho]", text)

    def test_what_the_group_declared_about_itself_is_printed(self) -> None:
        # Parsed and thrown away before this: ``import_groups`` kept the name
        # out of ``team.json`` and dropped the members it was read from.
        text = self.export()
        self.assertIn("Nhóm Alpha", text)
        self.assertIn("theo team.json", text)
        self.assertIn("24002222", text)

    def test_the_marks_and_the_moment_they_were_earned_are_printed(self) -> None:
        text = self.export()
        self.assertIn("90/100", text)
        self.assertIn("2026-09-10T02:00:00Z", text)

    def test_a_login_with_no_database_record_says_so_and_is_not_dropped(self) -> None:
        text = self.export(students=[klass()[0]])
        self.assertIn("@bob", text)
        self.assertIn("không có bản ghi trong database", text)

    def test_an_export_without_a_database_says_what_is_missing(self) -> None:
        text = self.export(students=[])
        self.assertIn("không kèm database", text)
        self.assertIn("không có bản ghi trong database", text)

    def test_the_recorded_grade_names_the_repository_it_was_written_from(self) -> None:
        records = klass()
        setattr(
            records[0],
            FIELD_GRADES,
            {GROUP_SLUG: {"grade": 0, "max_points": 100, "owner": "zoe"}},
        )
        self.assertIn("ghi từ kho zoe", self.export(students=records))

    def test_two_repositories_paying_one_member_are_flagged_beside_them(self) -> None:
        records = klass()
        setattr(
            records[0],
            FIELD_GRADE_CANDIDATES,
            {
                GROUP_SLUG: {
                    "alice": {"grade": 100, "max_points": 100, "owner": "alice"},
                    "zoe": {"grade": 0, "max_points": 100, "owner": "zoe"},
                }
            },
        )
        text = self.export(students=records)
        self.assertIn("2 kho cùng tính điểm", text)
        self.assertIn("zoe: 0/100", text)

    def test_a_collaborator_who_is_not_in_the_class_is_named(self) -> None:
        text = self.export(
            collaborators={W00_REPO: ["alice", "bob", "stranger", *TEACHERS]}
        )
        line = next(
            row for row in text.splitlines() if "Ngoài danh sách lớp" in row
        )
        self.assertIn("stranger", line)

    def test_a_broken_team_json_prints_the_fault_and_how_to_fix_it(self) -> None:
        text = self.export(team_json={W00_REPO: team_payload(course="MAT9999")})
        self.assertIn("Lỗi (", text)
        self.assertIn("cách sửa:", text)

    def test_a_group_reports_what_was_found_not_what_was_written(self) -> None:
        # The only caller never saves, so a line claiming the group was written
        # down would be false wherever it is actually read.
        text = self.export()
        self.assertIn("không có mâu thuẫn", text)
        self.assertNotIn("đã ghi vào database", text)

    def test_a_repository_with_no_team_json_is_not_blamed_on_the_flag(self) -> None:
        # The flag was passed; the group simply never wrote the file.  Telling
        # the teacher to pass it again sends them after a bug that is not there.
        text = self.export(team_json={})
        self.assertIn("kho chưa có file team.json", text)
        self.assertNotIn("--read-team-json", text)

    def test_an_assignment_with_no_marks_is_not_given_a_collection_time(self) -> None:
        records = klass()
        report = run_import(
            FakeRunner(repos=[FINAL_REPO]), records, assignment=FINAL_SLUG
        )
        text = format_groups_txt(
            report,
            org=ORG,
            classroom=CLASSROOM,
            students=records,
            collected_at="2026-09-10T08:59:28Z",
        )
        self.assertNotIn("2026-09-10T08:59:28Z", text)
        self.assertIn("không kho nào của bài này có trong scores.json", text)

    def test_an_ungraded_assignment_does_not_claim_the_collector_credited_it(
        self,
    ) -> None:
        # ``credited`` is a copy of the collaborator list when nothing was
        # published, and reading it under the collector's name would say the
        # marks went somewhere they never went.
        records = klass()
        report = run_import(
            FakeRunner(repos=[FINAL_REPO]), records, assignment=FINAL_SLUG
        )
        text = format_groups_txt(
            report, org=ORG, classroom=CLASSROOM, students=records
        )
        line = next(
            row for row in text.splitlines() if "Classroom50 tính điểm" in row
        )
        self.assertIn("chưa công bố", line)

    def test_a_graded_assignment_still_names_who_the_collector_paid(self) -> None:
        line = next(
            row
            for row in self.export().splitlines()
            if "Classroom50 tính điểm" in row
        )
        self.assertIn("alice", line)
        self.assertIn("bob", line)

    def test_the_header_counts_groups_rather_than_writes(self) -> None:
        text = self.export()
        self.assertIn("nhóm có mâu thuẫn", text)
        self.assertNotIn("bản ghi được cập nhật", text)

    def test_the_report_row_keeps_what_it_used_to_discard(self) -> None:
        report = run_import(
            FakeRunner(
                repos=[W00_REPO],
                scores=scores_doc(
                    {GROUP_SLUG: group_bucket("alice", ["alice", "bob"])}
                ),
                team_json={W00_REPO: w00_team_payload()},
            ),
            klass(),
            with_team_json=True,
        )
        row = report.groups[GROUP_SLUG][0]
        for key in ("credited", "snapshot", "outsiders", "reasons", "team", "result"):
            self.assertIn(key, row)
        self.assertEqual(row["team"]["group_name"], "Nhóm Alpha")
        self.assertEqual(row["result"]["submissions"][0]["score"], 90)
        # ``list-groups`` prints this dict, so the widening has to survive it.
        serialised = report_to_dict(report)["groups"][GROUP_SLUG][0]
        self.assertEqual(serialised["credited"], row["credited"])
        json.dumps(serialised)

    def test_a_run_that_never_read_team_json_says_so_instead_of_guessing(self) -> None:
        records = klass()
        report = run_import(
            FakeRunner(
                repos=[W00_REPO],
                scores=scores_doc(
                    {GROUP_SLUG: group_bucket("alice", ["alice", "bob"])}
                ),
            ),
            records,
        )
        text = format_groups_txt(
            report, org=ORG, classroom=CLASSROOM, students=records
        )
        self.assertIn("--read-team-json", text)


class TestHygiene(unittest.TestCase):
    def setUp(self) -> None:
        self.source = (REPO_ROOT / "course_hoanganhduc" / "c50_groups.py").read_text(
            encoding="utf-8"
        )
        self.constants = [
            node.value
            for node in ast.walk(ast.parse(self.source))
            if isinstance(node, ast.Constant) and isinstance(node.value, str)
        ]

    def test_the_module_builds_no_command_but_gh(self) -> None:
        # Prose about a command is not a call to it, so this reads the strings
        # the module can actually put in an argv.
        started = [text for text in self.constants if text.startswith("gh")]
        self.assertEqual(
            sorted(set(started)), ["gh", "gh_rate_limited", "gh_timeout"]
        )

    def test_no_membership_filter_and_no_bulk_pagination(self) -> None:
        for text in self.constants:
            self.assertNotIn("affiliation", text)
            self.assertNotIn("--paginate", text)
            self.assertNotIn("--method", text)
            self.assertNotIn("gh teacher team", text)

    def test_the_module_stays_free_of_the_database_layer(self) -> None:
        self.assertNotIn("import pandas", self.source)
        self.assertNotIn("from .data import", self.source)
        self.assertNotIn("sys.modules", self.source)


if __name__ == "__main__":
    unittest.main(verbosity=2)
