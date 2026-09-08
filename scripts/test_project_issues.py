#!/usr/bin/env python3
"""Offline tests for the mini-project issue lane.

The module under test is stdlib-only, never imports the database layer and never
imports pandas, so these tests run under a bare interpreter.  Every network call
goes through an injected runner, so nothing here reaches GitHub.

The fixtures are shaped like the two real topic boards: a private repository
whose issue form has fourteen required fields and five required checkboxes, a
staff verification issue that is not a proposal, group repositories named
``<classroom>-final-project-<founder>`` where both the classroom slug and the
usernames contain hyphens, and answers written the way students write them --
bullets, ``@`` prefixes, fenced templates copied whole out of the guide.

Two parsing rules in this lane are deliberately opposite, and both are tested:
the issue-body parser ignores everything inside a code fence, while the comment
parser reads inside fences, because the guide prints its comment templates in
``text`` fences and students copy the fence along with the text.
"""

from __future__ import annotations

import ast
import base64
import json
import re
import sys
import unittest
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from course_hoanganhduc.c50_cli import Classroom50Error, RunResult  # noqa: E402
from course_hoanganhduc.project_issues import (  # noqa: E402
    FIELD_COMMIT, FIELD_FINAL, FIELD_GROUP_NAME, FIELD_HASH, FIELD_MEMBERS,
    FIELD_NUMBER, FIELD_REPO, FIELD_STATUS, FIELD_TITLE, FIELD_TOPIC_SOURCE,
    FORM_PATH, KIND_FINAL, KIND_MEMBERSHIP_CHANGE, KIND_PROPOSAL_UPDATE,
    LABEL_DUPLICATE, LABEL_RECORDED, LABEL_SUBMITTED, MINI_PROJECT_SLUG,
    NO_RESPONSE, PINNED_FIELDS, PROPOSAL_PATH, REDACTION, SCHEMA_PINNED,
    SCHEMA_REPO, SOURCE_ISSUE, SOURCE_PROPOSAL_MD, SOURCE_STAFF_DUPLICATE,
    SOURCE_STAFF_RECORDED, FormField, FormSchema, choose_canonical, content_hash,
    detect_member_conflicts, fetch_comments, fetch_edit_history, fetch_issues,
    format_project_issues, import_project_issues, is_proposal_issue,
    issue_snapshot, list_project_issues, load_form_schema, merge_into_students,
    normalize_commit, normalize_label, normalize_repo_url, normalize_yes_no,
    parse_comments, parse_form_schema, parse_issue_body, parse_members,
    parse_proposal, parse_proposal_markers, proposal_findings, recorded_groups,
    redact_personal, report_to_dict, values_from_markers,
)
from course_hoanganhduc.roster_audit import (  # noqa: E402
    CONSEQUENTIAL_CODES, SEVERITY_ERROR, SEVERITY_WARNING,
)

# ---------------------------------------------------------------------------
# fixtures
# ---------------------------------------------------------------------------

BOARD = "VNU-HUS/mat1206e-2026-project-topics"
ORG = "vnu-hus"
CLASSROOM = "vnu-hus-mat1206e-winter-2026"

# Both halves of a repository name carry hyphens, exactly as upstream names them.
REPO_AN = f"{ORG}/{CLASSROOM}-final-project-an-nguyen"
REPO_BINH = f"{ORG}/{CLASSROOM}-final-project-binh-tran"
KNOWN_REPOS = frozenset({REPO_AN, REPO_BINH})

STAFF = frozenset({"hoanganhduc", "ivy"})

SHA_ONE = "3af80db0206215cfd43314b5feba8060c632e417"
SHA_TWO = "cbb330fa1d2e3f405162738495a6b7c8d9e0f102"
# A blob SHA is the same shape as a commit SHA; only the endpoint tells them
# apart, which is why the lane never reads one out of a contents response.
BLOB_SHA = "0123456789abcdef0123456789abcdef01234567"

CONFIRMATION_OPTIONS = (
    "All listed members agree to this proposal.",
    "The repository URL above is the exact Classroom50 group repository.",
    "team.json in that repository lists exactly the members above.",
    "The proposal commit above belongs to that repository.",
    "No personal data beyond GitHub usernames appears in this issue.",
)


def _form_yaml(labels: Optional[Dict[str, str]] = None) -> str:
    """The board's issue form, at the indentation the real file uses.

    A body item opens at two spaces, its ``id`` sits at four and its label at
    six; a checkbox option label also spells ``label:`` but sits at eight behind
    a dash.  Reproducing those columns exactly is the point of this fixture: it
    is what keeps the five confirmation options from reading as five fields.
    """
    overrides = dict(labels or {})
    lines = [
        "name: Mini-project topic proposal",
        "description: Register one mini-project topic for one Classroom50 group.",
        'title: "[Proposal] "',
        'labels: ["status: submitted"]',
        "body:",
        "  - type: markdown",
        "    attributes:",
        "      value: |",
        "        Read the Submission Guide before filling this in.",
        "        Use GitHub usernames only.",
    ]
    for field_id, label in PINNED_FIELDS:
        if field_id == "confirmations":
            continue
        kind = "textarea" if field_id in {
            "members", "selected_problem", "public_summary", "scope",
            "method_resources", "expected_output", "overlap",
        } else "input"
        if field_id == "topic_source":
            kind = "dropdown"
        lines.extend(
            [
                f"  - type: {kind}",
                f"    id: {field_id}",
                "    attributes:",
                f"      label: {overrides.get(field_id, label)}",
                f"      description: Answer for {field_id}.",
            ]
        )
        if kind == "dropdown":
            lines.append("      options:")
            lines.extend(
                [
                    "        - Based on one suggested idea",
                    "        - Combines or adapts suggested ideas",
                    "        - Self-proposed",
                ]
            )
        lines.extend(["    validations:", "      required: true"])
    lines.extend(
        [
            "  - type: checkboxes",
            "    id: confirmations",
            "    attributes:",
            f"      label: {overrides.get('confirmations', 'Group confirmations')}",
            "      options:",
        ]
    )
    for option in CONFIRMATION_OPTIONS:
        lines.extend([f"        - label: {option}", "          required: true"])
    lines.extend(["    validations:", "      required: true", ""])
    return "\n".join(lines)


FORM_YAML = _form_yaml()

SCHEMA = FormSchema(
    repo=BOARD, fields=parse_form_schema(FORM_YAML, repo=BOARD), source=SCHEMA_REPO
)


def proposal_values(
    *,
    repo: str = REPO_AN,
    founder: str = "an-nguyen",
    members: Sequence[str] = ("an-nguyen", "bao-le", "chi-pham"),
    group_name: str = "Nhóm Alpha",
    project_title: str = "Nhận dạng chữ viết tay",
    commit: str = SHA_ONE,
    topic_source: str = "Self-proposed",
    selected_problem: str = "Phân loại chữ số viết tay trên tập dữ liệu của lớp.",
) -> Dict[str, str]:
    """One filled-in form, keyed by field id."""
    return {
        "group_name": group_name,
        "repository_url": f"https://github.com/{repo}",
        "founder": f"@{founder}",
        "members": "\n".join(f"- @{name}" for name in members),
        "project_title": project_title,
        "topic_source": topic_source,
        "selected_problem": selected_problem,
        "public_summary": "Xây dựng bộ phân loại chữ số và đánh giá độ chính xác.",
        "scope": "Chỉ dùng dữ liệu công khai của lớp; không thu thập dữ liệu mới.",
        "method_resources": "Hồi quy logistic và một mạng nơ-ron nhỏ; thư viện numpy.",
        "expected_output": "Notebook huấn luyện kèm bảng kết quả.",
        "overlap": "Không trùng với chủ đề nào đã được ghi nhận.",
        "proposal_commit": f"https://github.com/{repo}/commit/{commit}",
        "confirmations": "\n".join(f"- [X] {option}" for option in CONFIRMATION_OPTIONS),
    }


def render_body(
    values: Dict[str, str],
    *,
    level: str = "###",
    omit: Sequence[str] = (),
    schema: Sequence[FormField] = SCHEMA.fields,
) -> str:
    """The markdown GitHub renders from a filled-in issue form."""
    parts: List[str] = []
    for field in schema:
        if field.field_id in omit:
            continue
        answer = values.get(field.field_id, NO_RESPONSE)
        parts.append(f"{level} {field.label}\n\n{answer}\n")
    return "\n".join(parts)


def issue_row(
    number: int,
    *,
    body: str,
    title: str = "[Proposal] Nhận dạng chữ viết tay",
    author: str = "an-nguyen",
    association: str = "NONE",
    state: str = "open",
    labels: Sequence[str] = (LABEL_SUBMITTED,),
    created_at: str = "2026-09-21T02:00:00Z",
    updated_at: str = "2026-09-21T02:00:00Z",
) -> Dict[str, Any]:
    """One row of ``GET /repos/{board}/issues``."""
    return {
        "number": number,
        "title": title,
        "body": body,
        "user": {"login": author},
        "author_association": association,
        "state": state,
        "labels": [{"name": name} for name in labels],
        "html_url": f"https://github.com/{BOARD}/issues/{number}",
        "created_at": created_at,
        "updated_at": updated_at,
    }


def comment_row(
    comment_id: int,
    body: str,
    *,
    author: str = "an-nguyen",
    association: str = "NONE",
    created_at: str = "2026-09-22T02:00:00Z",
) -> Dict[str, Any]:
    """One row of ``GET /repos/{board}/issues/{n}/comments``."""
    return {
        "id": comment_id,
        "user": {"login": author},
        "author_association": association,
        "body": body,
        "created_at": created_at,
        "updated_at": created_at,
    }


def edit_node(
    number: int,
    *,
    last_edited_at: Optional[str] = None,
    edit_count: int = 1,
    comments: Sequence[Tuple[int, Optional[str]]] = (),
) -> Dict[str, Any]:
    """One node of the ``userContentEdits`` GraphQL answer."""
    return {
        "number": number,
        "lastEditedAt": last_edited_at,
        "userContentEdits": {"totalCount": edit_count},
        "comments": {
            "nodes": [
                {"databaseId": cid, "lastEditedAt": at} for cid, at in comments
            ]
        },
    }


STAFF_ISSUE = issue_row(
    1,
    title="[Staff verification] Board setup check (not a proposal)",
    body="Board setup check for the topic board. This is not a proposal.\n",
    author="hoanganhduc",
    association="OWNER",
    state="closed",
    labels=(),
    created_at="2026-09-05T01:00:00Z",
    updated_at="2026-09-05T01:00:00Z",
)


class Student:
    """The attribute bag the database hands the lane."""

    def __init__(
        self,
        name: str,
        student_id: str,
        username: str,
        *,
        github_id: str = "",
        email: str = "",
    ) -> None:
        self.Name = name
        setattr(self, "Student ID", student_id)
        setattr(self, "GitHub Username", username)
        if github_id:
            setattr(self, "GitHub ID", github_id)
        self.Email = email


def klass() -> List[Student]:
    """Five students.

    Vietnamese names on purpose: ``refine_database`` deletes records whose name
    looks like a test fixture.
    """
    return [
        Student("Nguyễn Văn An", "24001101", "an-nguyen", github_id="501101"),
        Student("Lê Thị Bảo", "24001102", "bao-le", github_id="501102"),
        Student("Phạm Minh Chi", "24001103", "chi-pham", github_id="501103"),
        Student("Trần Thanh Bình", "24001104", "binh-tran", github_id="501104"),
        Student("Vũ Tiến Dũng", "24001105", "dung-vo", github_id="501105"),
    ]


class FakeRunner:
    """Answers the exact argv shapes this lane builds, and nothing else.

    Anything the lane has no business calling raises, so a test that claims "no
    request was made to that URL" is checking the runner's own refusal as well
    as the recorded call list.
    """

    def __init__(
        self,
        *,
        form: Optional[str] = FORM_YAML,
        issues: Sequence[Dict[str, Any]] = (),
        comments: Optional[Dict[int, Sequence[Dict[str, Any]]]] = None,
        edits: Optional[Sequence[Dict[str, Any]]] = (),
        proposals: Optional[Dict[Tuple[str, str], str]] = None,
    ) -> None:
        self.form = form
        self.issues = list(issues)
        self.comments = dict(comments or {})
        self.edits = edits
        self.proposals = dict(proposals or {})
        self.failures: Dict[str, RunResult] = {}
        self.calls: List[List[str]] = []
        self.timeouts: List[Optional[float]] = []

    def paths(self) -> List[str]:
        return [call[-1] for call in self.calls]

    def api_paths(self) -> List[str]:
        return [
            call[-1]
            for call in self.calls
            if call[1:2] == ["api"] and call[2:3] != ["graphql"]
        ]

    @staticmethod
    def _page(path: str, rows: Sequence[Any]) -> List[Any]:
        return list(rows) if "page=1" in path else []

    def __call__(self, argv: Sequence[str], *, timeout: Optional[float] = None) -> RunResult:
        self.calls.append(list(argv))
        self.timeouts.append(timeout)
        return self._answer(list(argv))

    def _graphql(self) -> RunResult:
        if self.edits is None:
            return RunResult(1, "", "gh: Could not resolve to a Repository (HTTP 404)")
        payload = {
            "data": {
                "repository": {
                    "issues": {
                        "pageInfo": {"hasNextPage": False, "endCursor": None},
                        "nodes": list(self.edits),
                    }
                }
            }
        }
        return RunResult(0, json.dumps(payload), "")

    def _answer(self, argv: List[str]) -> RunResult:
        path = argv[-1]
        for fragment, result in self.failures.items():
            if fragment in path:
                return result
        if argv[1:3] == ["api", "graphql"]:
            return self._graphql()
        if path.endswith(FORM_PATH):
            if self.form is None:
                return RunResult(1, "", "gh: Not Found (HTTP 404)")
            return RunResult(0, self.form, "")
        if f"/contents/{PROPOSAL_PATH}" in path:
            head, _, ref = path.partition("?ref=")
            repo = head.split("/contents/")[0][len("repos/"):]
            text = self.proposals.get((repo, ref))
            if text is None:
                return RunResult(1, "", "gh: Not Found (HTTP 404)")
            return RunResult(0, text, "")
        if "/issues?" in path:
            return RunResult(0, json.dumps(self._page(path, self.issues)), "")
        found = re.search(r"/issues/(\d+)/comments", path)
        if found is not None:
            rows = self.comments.get(int(found.group(1)), [])
            return RunResult(0, json.dumps(self._page(path, rows)), "")
        raise AssertionError(f"the lane made an unexpected call: {argv}")


def one_proposal(
    number: int = 7,
    values: Optional[Dict[str, str]] = None,
    **kwargs: Any,
) -> Any:
    """A parsed proposal built straight from values, with no network at all."""
    body = render_body(values if values is not None else proposal_values())
    return parse_proposal(
        _raw(issue_row(number, body=body)),
        SCHEMA,
        known_repos=KNOWN_REPOS,
        expected_org="VNU-HUS",
        **kwargs,
    )


def _raw(row: Dict[str, Any]) -> Any:
    """One issue payload as the lane's own ``RawIssue``."""
    runner = FakeRunner(issues=[row])
    return fetch_issues(BOARD, runner=runner, sleeper=lambda _s: None, pace=0.0)[0]


def _comments(rows: Sequence[Dict[str, Any]]) -> Any:
    runner = FakeRunner(comments={7: list(rows)})
    return fetch_comments(BOARD, 7, runner=runner, sleeper=lambda _s: None, pace=0.0)


def _index(students: Sequence[Student]) -> Dict[str, Student]:
    from course_hoanganhduc.c50_scores import _student_index

    index, _warnings = _student_index(students, None)
    return index


# ---------------------------------------------------------------------------
# the issue form
# ---------------------------------------------------------------------------


class TestFormSchema(unittest.TestCase):
    def test_the_form_yields_exactly_the_fourteen_fields(self) -> None:
        self.assertEqual(len(SCHEMA.fields), 14)
        self.assertEqual(
            [(field.field_id, field.label) for field in SCHEMA.fields],
            list(PINNED_FIELDS),
        )

    def test_the_markdown_item_carries_no_id_and_is_skipped(self) -> None:
        self.assertNotIn("", [field.field_id for field in SCHEMA.fields])
        self.assertNotIn(
            "Read the Submission Guide before filling this in.",
            [field.label for field in SCHEMA.fields],
        )

    def test_the_five_checkbox_options_are_not_read_as_fields(self) -> None:
        labels = [field.label for field in SCHEMA.fields]
        for option in CONFIRMATION_OPTIONS:
            self.assertNotIn(option, labels)
        self.assertIn("Group confirmations", labels)

    def test_the_schema_is_read_from_the_repository_when_it_can_be(self) -> None:
        runner = FakeRunner()
        schema = load_form_schema(BOARD, runner=runner, sleeper=lambda _s: None, pace=0.0)
        self.assertEqual(schema.source, SCHEMA_REPO)
        self.assertEqual(len(schema.fields), 14)
        self.assertEqual(
            runner.calls[0][2:4], ["-H", "Accept: application/vnd.github.raw"]
        )
        self.assertTrue(runner.calls[0][-1].endswith(FORM_PATH))

    def test_an_unreadable_form_falls_back_to_the_pinned_labels(self) -> None:
        schema = load_form_schema(
            BOARD, runner=FakeRunner(form=None), sleeper=lambda _s: None, pace=0.0
        )
        self.assertEqual(schema.source, SCHEMA_PINNED)
        self.assertEqual(schema.fields, PINNED_FIELDS)

    def test_a_renamed_label_in_the_yml_is_followed(self) -> None:
        renamed = parse_form_schema(
            _form_yaml({"selected_problem": "Chosen problem statement"}), repo=BOARD
        )
        by_id = {field.field_id: field.label for field in renamed}
        self.assertEqual(by_id["selected_problem"], "Chosen problem statement")
        values = proposal_values()
        body = render_body(values, schema=renamed)
        parsed = parse_issue_body(body, SCHEMA._replace(fields=renamed))
        self.assertEqual(parsed["selected_problem"], values["selected_problem"])

    def test_a_label_key_ignores_case_accents_and_trailing_punctuation(self) -> None:
        self.assertEqual(normalize_label("  Group   NAME:  "), "group name")
        self.assertEqual(normalize_label("Tổ chức"), "to chuc")


# ---------------------------------------------------------------------------
# the issue body
# ---------------------------------------------------------------------------


class TestBodyParsing(unittest.TestCase):
    def test_the_rendered_form_parses_all_fourteen_fields(self) -> None:
        values = proposal_values()
        parsed = parse_issue_body(render_body(values), SCHEMA)
        self.assertEqual(len(parsed), 14)
        self.assertEqual(parsed["project_title"], values["project_title"])
        self.assertEqual(parsed["topic_source"], "Self-proposed")

    def test_the_guide_example_heading_level_parses_too(self) -> None:
        values = proposal_values()
        parsed = parse_issue_body(render_body(values, level="##"), SCHEMA)
        self.assertEqual(len(parsed), 14)
        self.assertEqual(parsed["group_name"], values["group_name"])

    def test_a_heading_inside_a_code_fence_is_not_a_heading(self) -> None:
        values = proposal_values()
        values["selected_problem"] = (
            "Ví dụ về một mẫu bị chép nhầm:\n"
            "\n"
            "```text\n"
            "### Project title\n"
            "Chủ đề giả\n"
            "```\n"
            "\n"
            "Bài toán thật là phân loại chữ số viết tay."
        )
        parsed = parse_issue_body(render_body(values), SCHEMA)
        self.assertEqual(parsed["project_title"], "Nhận dạng chữ viết tay")
        self.assertIn("### Project title", parsed["selected_problem"])
        self.assertIn("Bài toán thật", parsed["selected_problem"])

    def test_the_staff_verification_issue_is_not_a_proposal(self) -> None:
        self.assertFalse(is_proposal_issue(_raw(STAFF_ISSUE), SCHEMA))

    def test_a_filled_form_is_a_proposal(self) -> None:
        body = render_body(proposal_values())
        self.assertTrue(is_proposal_issue(_raw(issue_row(7, body=body)), SCHEMA))

    def test_the_title_prefix_alone_does_not_make_a_proposal(self) -> None:
        row = issue_row(9, body="Xin chào, em muốn đăng ký chủ đề ạ.\n")
        self.assertTrue(row["title"].startswith("[Proposal] "))
        self.assertFalse(is_proposal_issue(_raw(row), SCHEMA))

    def test_three_headings_are_below_the_proposal_threshold(self) -> None:
        values = proposal_values()
        body = render_body(
            values,
            omit=[field.field_id for field in SCHEMA.fields[3:]],
        )
        self.assertFalse(is_proposal_issue(_raw(issue_row(9, body=body)), SCHEMA))


# ---------------------------------------------------------------------------
# normalisation
# ---------------------------------------------------------------------------


class TestNormalisation(unittest.TestCase):
    def test_every_way_of_writing_a_repository_url_lands_on_one_key(self) -> None:
        written = [
            f"https://github.com/{REPO_AN}",
            f"https://github.com/{REPO_AN}/",
            f"https://github.com/{REPO_AN}.git",
            f"github.com/{REPO_AN}",
            f"https://www.github.com/{REPO_AN}",
            f"https://github.com/{REPO_AN}".upper().replace("HTTPS://GITHUB.COM", "https://github.com"),
            REPO_AN,
            f"<https://github.com/{REPO_AN}>",
        ]
        for value in written:
            with self.subTest(value=value):
                self.assertEqual(normalize_repo_url(value), REPO_AN)

    def test_a_url_that_is_not_a_repository_normalises_to_nothing(self) -> None:
        self.assertEqual(normalize_repo_url(""), "")
        self.assertEqual(normalize_repo_url("chưa có ạ"), "")

    def test_a_commit_url_and_a_bare_sha_agree(self) -> None:
        self.assertEqual(
            normalize_commit(f"https://github.com/{REPO_AN}/commit/{SHA_ONE}"), SHA_ONE
        )
        self.assertEqual(normalize_commit(SHA_ONE.upper()), SHA_ONE)
        self.assertEqual(normalize_commit(f"  {SHA_ONE}  "), SHA_ONE)

    def test_a_short_sha_is_not_a_commit(self) -> None:
        self.assertEqual(normalize_commit(SHA_ONE[:7]), "")
        self.assertEqual(normalize_commit("commit gần nhất"), "")

    def test_yes_and_no_survive_the_way_students_write_them(self) -> None:
        for value in ("Yes", "YES", "yes.", "có", "Có", "rồi", "true", "1"):
            with self.subTest(value=value):
                self.assertIs(normalize_yes_no(value), True)
        for value in ("no.", "N", "No", "chưa", "Chưa", "không", "0"):
            with self.subTest(value=value):
                self.assertIs(normalize_yes_no(value), False)
        for value in ("", "đang chờ nhóm trả lời"):
            with self.subTest(value=value):
                self.assertIsNone(normalize_yes_no(value))

    def test_members_survive_bullets_at_signs_and_separators(self) -> None:
        logins, rejected = parse_members(
            "- @an-nguyen\n* `bao-le`\n1. @Chi-Pham, dung-vo; @an-nguyen"
        )
        self.assertEqual(logins, ("an-nguyen", "bao-le", "chi-pham", "dung-vo"))
        self.assertEqual(rejected, ())

    def test_anything_that_is_not_a_username_is_handed_back_separately(self) -> None:
        logins, rejected = parse_members("- @an-nguyen\n- Nguyễn Văn An\n- -bad-name")
        self.assertEqual(logins, ("an-nguyen",))
        self.assertEqual(rejected, ("Nguyễn Văn An", "-bad-name"))

    def test_student_numbers_and_addresses_are_removed(self) -> None:
        text, ids, emails = redact_personal(
            "Nhóm của 24001101 và an.nguyen@vnu.edu.vn xin đăng ký."
        )
        self.assertEqual((ids, emails), (1, 1))
        self.assertNotIn("24001101", text)
        self.assertNotIn("vnu.edu.vn", text)
        self.assertEqual(text.count(REDACTION), 2)

    def test_a_four_digit_year_is_not_a_student_number(self) -> None:
        text, ids, emails = redact_personal("Học kỳ 1 năm 2026")
        self.assertEqual((ids, emails), (0, 0))
        self.assertEqual(text, "Học kỳ 1 năm 2026")

    def test_the_content_digest_ignores_key_order_and_sees_a_changed_answer(self) -> None:
        first = content_hash({"a": "một", "b": "hai"})
        self.assertEqual(first, content_hash({"b": "hai", "a": "một"}))
        self.assertNotEqual(first, content_hash({"a": "một", "b": "ba"}))


# ---------------------------------------------------------------------------
# comments
# ---------------------------------------------------------------------------

PROPOSAL_UPDATE_COMMENT = f"""PROPOSAL UPDATE

```text
Previous proposal commit: {SHA_ONE}
New proposal commit: {SHA_TWO}
Selected problem changed: yes
Summary of changes: thu hẹp phạm vi còn chữ số viết tay
Reason for the changes: dữ liệu chữ cái quá lớn cho thời gian còn lại
```
"""

MEMBERSHIP_COMMENT = """MEMBERSHIP CHANGE REQUEST

```text
Current members: @an-nguyen, @bao-le, @chi-pham
Proposed members: @an-nguyen, @bao-le
Reason: một bạn xin rút khỏi nhóm
Commit updating team.json: pending
```
"""

FINAL_COMMENT = f"""FINAL SUBMISSION

```text
Final commit: https://github.com/{REPO_AN}/commit/{SHA_TWO}
Report: report/report.pdf
Slides: slides/slides.pdf
All listed members agree that this is the final submission: Yes
```
"""


class TestComments(unittest.TestCase):
    def test_a_template_copied_with_its_fence_is_still_read(self) -> None:
        facts = parse_comments(
            _comments([comment_row(11, PROPOSAL_UPDATE_COMMENT)]), staff_logins=STAFF
        )
        self.assertIn(KIND_PROPOSAL_UPDATE, facts.kinds)
        self.assertEqual(facts.proposal_commit, SHA_TWO)
        self.assertIs(facts.problem_changed, True)

    def test_the_longer_reason_label_wins_over_the_shorter_one(self) -> None:
        facts = parse_comments(
            _comments([comment_row(11, PROPOSAL_UPDATE_COMMENT)]), staff_logins=STAFF
        )
        # 'reason for the changes' must not be eaten by the bare 'reason' label,
        # or the update would read as a membership request's reason.
        self.assertIn(KIND_PROPOSAL_UPDATE, facts.kinds)
        self.assertFalse(facts.membership_pending)

    def test_a_final_submission_records_its_commit(self) -> None:
        facts = parse_comments(
            _comments([comment_row(12, FINAL_COMMENT)]), staff_logins=STAFF
        )
        self.assertIn(KIND_FINAL, facts.kinds)
        self.assertIs(facts.final_submitted, True)
        self.assertEqual(facts.final_commit, SHA_TWO)

    def test_a_membership_request_from_a_student_stays_pending(self) -> None:
        facts = parse_comments(
            _comments([comment_row(13, MEMBERSHIP_COMMENT)]), staff_logins=STAFF
        )
        self.assertIn(KIND_MEMBERSHIP_CHANGE, facts.kinds)
        self.assertTrue(facts.membership_pending)
        self.assertEqual(facts.membership_proposed, ("an-nguyen", "bao-le"))

    def test_a_staff_answer_closes_the_membership_request(self) -> None:
        facts = parse_comments(
            _comments(
                [
                    comment_row(13, MEMBERSHIP_COMMENT),
                    comment_row(
                        14,
                        "MEMBERSHIP CHANGE REQUEST\n\nĐã cập nhật, các em tiếp tục nhé.",
                        author="hoanganhduc",
                        association="OWNER",
                        created_at="2026-09-23T02:00:00Z",
                    ),
                ]
            ),
            staff_logins=STAFF,
        )
        self.assertFalse(facts.membership_pending)

    def test_only_staff_can_mark_a_duplicate(self) -> None:
        by_staff = parse_comments(
            _comments(
                [
                    comment_row(
                        15,
                        "Duplicate of #3 -- nhóm kia đăng ký trước.",
                        author="hoanganhduc",
                        association="OWNER",
                    )
                ]
            ),
            staff_logins=STAFF,
        )
        self.assertEqual(by_staff.duplicate_of, 3)
        by_student = parse_comments(
            _comments([comment_row(15, "Duplicate of #3 ạ?")]), staff_logins=STAFF
        )
        self.assertIsNone(by_student.duplicate_of)

    def test_the_newer_proposal_commit_beats_the_body(self) -> None:
        proposal = one_proposal(
            comments=_comments([comment_row(11, PROPOSAL_UPDATE_COMMENT)]),
            staff_logins=STAFF,
        )
        self.assertEqual(proposal.commit, SHA_TWO)
        self.assertEqual(proposal.invalid, ())


# ---------------------------------------------------------------------------
# the canonical choice
# ---------------------------------------------------------------------------


class TestCanonical(unittest.TestCase):
    def test_the_earliest_valid_issue_for_a_repository_wins(self) -> None:
        first = one_proposal(3)
        second = one_proposal(9)
        chosen = choose_canonical([second, first])
        self.assertEqual(chosen[REPO_AN].number, 3)
        self.assertEqual(chosen[REPO_AN].superseded, (9,))
        self.assertEqual(chosen[REPO_AN].source, "computed")

    def test_an_invalid_first_issue_yields_to_the_next_valid_one(self) -> None:
        broken = proposal_values()
        broken["proposal_commit"] = "chưa có ạ"
        first = one_proposal(3, broken)
        second = one_proposal(9)
        self.assertTrue(first.invalid)
        chosen = choose_canonical([first, second])
        self.assertEqual(chosen[REPO_AN].number, 9)

    def test_a_staff_recorded_label_beats_the_creation_order(self) -> None:
        body = render_body(proposal_values())
        first = parse_proposal(
            _raw(issue_row(3, body=body)),
            SCHEMA,
            known_repos=KNOWN_REPOS,
            expected_org="VNU-HUS",
        )
        second = parse_proposal(
            _raw(issue_row(9, body=body, labels=(LABEL_SUBMITTED, LABEL_RECORDED))),
            SCHEMA,
            known_repos=KNOWN_REPOS,
            expected_org="VNU-HUS",
        )
        chosen = choose_canonical([first, second])
        self.assertEqual(chosen[REPO_AN].number, 9)
        self.assertEqual(chosen[REPO_AN].source, SOURCE_STAFF_RECORDED)
        self.assertEqual(chosen[REPO_AN].superseded, (3,))

    def test_a_staff_duplicate_pointer_names_the_canonical_issue(self) -> None:
        body = render_body(proposal_values())
        pointer = _comments(
            [
                comment_row(
                    21,
                    "Duplicate of #9",
                    author="hoanganhduc",
                    association="OWNER",
                )
            ]
        )
        first = parse_proposal(
            _raw(issue_row(3, body=body)),
            SCHEMA,
            comments=pointer,
            staff_logins=STAFF,
            known_repos=KNOWN_REPOS,
            expected_org="VNU-HUS",
        )
        second = parse_proposal(
            _raw(issue_row(9, body=body, labels=(LABEL_SUBMITTED, LABEL_DUPLICATE))),
            SCHEMA,
            known_repos=KNOWN_REPOS,
            expected_org="VNU-HUS",
        )
        chosen = choose_canonical([first, second])
        self.assertEqual(chosen[REPO_AN].number, 9)
        self.assertEqual(chosen[REPO_AN].source, SOURCE_STAFF_DUPLICATE)

    def test_repairing_an_older_issue_does_not_unseat_a_staff_decision(self) -> None:
        body = render_body(proposal_values())
        repaired = parse_proposal(
            _raw(issue_row(3, body=body)),
            SCHEMA,
            known_repos=KNOWN_REPOS,
            expected_org="VNU-HUS",
        )
        recorded = parse_proposal(
            _raw(issue_row(9, body=body, labels=(LABEL_RECORDED,))),
            SCHEMA,
            known_repos=KNOWN_REPOS,
            expected_org="VNU-HUS",
        )
        self.assertEqual(repaired.invalid, ())
        chosen = choose_canonical([repaired, recorded])
        self.assertEqual(chosen[REPO_AN].number, 9)
        self.assertEqual(chosen[REPO_AN].source, SOURCE_STAFF_RECORDED)

    def test_a_proposal_with_no_readable_repository_is_not_canonical_anywhere(self) -> None:
        broken = proposal_values()
        broken["repository_url"] = "em quên mất ạ"
        self.assertEqual(choose_canonical([one_proposal(3, broken)]), {})

    def test_two_repositories_each_keep_their_own_canonical_issue(self) -> None:
        an = one_proposal(3)
        binh = one_proposal(
            5,
            proposal_values(
                repo=REPO_BINH,
                founder="binh-tran",
                members=("binh-tran", "dung-vo"),
                group_name="Nhóm Beta",
                project_title="Tìm đường trên bản đồ",
            ),
        )
        chosen = choose_canonical([an, binh])
        self.assertEqual(sorted(chosen), sorted({REPO_AN, REPO_BINH}))
        self.assertEqual(chosen[REPO_BINH].number, 5)


# ---------------------------------------------------------------------------
# cross-issue conflicts
# ---------------------------------------------------------------------------


class TestConflicts(unittest.TestCase):
    def setUp(self) -> None:
        self.students = klass()
        self.index = _index(self.students)

    def test_one_student_named_on_two_repositories_quarantines_both_groups(self) -> None:
        an = one_proposal(3)
        binh = one_proposal(
            5,
            proposal_values(
                repo=REPO_BINH,
                founder="binh-tran",
                members=("binh-tran", "dung-vo", "chi-pham"),
                group_name="Nhóm Beta",
            ),
        )
        proposals = [an, binh]
        findings, quarantine = detect_member_conflicts(
            choose_canonical(proposals), proposals, self.index
        )
        codes = {issue.code for issue in findings}
        self.assertIn("issue_member_shared", codes)
        # the shared student and every teammate of both groups
        self.assertEqual(
            quarantine,
            {"an-nguyen", "bao-le", "chi-pham", "binh-tran", "dung-vo"},
        )

    def test_the_canonical_rule_alone_cannot_see_that_conflict(self) -> None:
        an = one_proposal(3)
        binh = one_proposal(
            5,
            proposal_values(
                repo=REPO_BINH,
                founder="binh-tran",
                members=("binh-tran", "dung-vo", "chi-pham"),
            ),
        )
        chosen = choose_canonical([an, binh])
        # Two repositories, two canonical issues, nothing superseded: keying on
        # the repository URL is exactly what makes the member clash invisible.
        self.assertEqual(len(chosen), 2)
        self.assertEqual(chosen[REPO_AN].superseded, ())
        self.assertEqual(chosen[REPO_BINH].superseded, ())

    def test_a_membership_list_that_disagrees_with_the_recorded_group(self) -> None:
        an = one_proposal(3)
        groups = {REPO_AN: ("an-nguyen", ("an-nguyen", "bao-le", "dung-vo"))}
        findings, quarantine = detect_member_conflicts(
            choose_canonical([an]), [an], self.index, groups=groups
        )
        by_code = {issue.code: issue for issue in findings}
        self.assertIn("issue_group_mismatch", by_code)
        self.assertIn("thừa", by_code["issue_group_mismatch"].detail)
        self.assertIn("thiếu", by_code["issue_group_mismatch"].detail)
        self.assertEqual(
            quarantine, {"an-nguyen", "bao-le", "chi-pham", "dung-vo"}
        )

    def test_a_founder_who_is_not_the_recorded_founder(self) -> None:
        an = one_proposal(3)
        groups = {REPO_AN: ("bao-le", ("an-nguyen", "bao-le", "chi-pham"))}
        findings, quarantine = detect_member_conflicts(
            choose_canonical([an]), [an], self.index, groups=groups
        )
        codes = {issue.code for issue in findings}
        self.assertIn("issue_founder_mismatch", codes)
        self.assertIn("an-nguyen", quarantine)

    def test_a_group_that_agrees_with_the_record_raises_nothing(self) -> None:
        an = one_proposal(3)
        groups = {REPO_AN: ("an-nguyen", ("an-nguyen", "bao-le", "chi-pham"))}
        findings, quarantine = detect_member_conflicts(
            choose_canonical([an]), [an], self.index, groups=groups
        )
        self.assertEqual(findings, [])
        self.assertEqual(quarantine, set())

    def test_without_a_recorded_group_only_the_member_clash_is_checked(self) -> None:
        an = one_proposal(3)
        findings, quarantine = detect_member_conflicts(
            choose_canonical([an]), [an], self.index, groups=None
        )
        self.assertEqual(findings, [])
        self.assertEqual(quarantine, set())

    def test_the_recorded_group_is_read_from_the_database_not_from_github(self) -> None:
        students = klass()
        setattr(
            students[0],
            "Classroom50 Group",
            {
                MINI_PROJECT_SLUG: {
                    "repo": REPO_AN,
                    "founder": "an-nguyen",
                    "members": ["an-nguyen", "bao-le", "chi-pham"],
                }
            },
        )
        groups = recorded_groups(students)
        self.assertEqual(
            groups[REPO_AN], ("an-nguyen", ("an-nguyen", "bao-le", "chi-pham"))
        )


# ---------------------------------------------------------------------------
# per-proposal findings
# ---------------------------------------------------------------------------


class TestFindings(unittest.TestCase):
    def setUp(self) -> None:
        self.students = klass()
        self.index = _index(self.students)

    def test_a_repository_outside_the_organisation_is_reported(self) -> None:
        values = proposal_values()
        values["repository_url"] = "https://github.com/some-other-org/random-repo"
        proposal = one_proposal(3, values)
        findings = proposal_findings(proposal, self.index)
        codes = {issue.code for issue in findings}
        self.assertIn("issue_repo_unknown", codes)
        self.assertTrue(any("ngoài tổ chức" in reason for reason in proposal.invalid))

    def test_a_repository_that_no_group_owns_is_reported(self) -> None:
        values = proposal_values()
        values["repository_url"] = f"https://github.com/{ORG}/{CLASSROOM}-final-project-ghost"
        proposal = one_proposal(3, values)
        self.assertTrue(
            any("không trỏ tới repo nhóm" in reason for reason in proposal.invalid)
        )
        self.assertIn(
            "issue_repo_unknown", {issue.code for issue in proposal_findings(proposal, self.index)}
        )

    def test_a_no_response_answer_is_read_as_a_bypassed_form(self) -> None:
        values = proposal_values()
        values["overlap"] = NO_RESPONSE
        proposal = one_proposal(3, values)
        self.assertIn("overlap", proposal.bypass)
        self.assertIn("overlap", proposal.missing)
        codes = {issue.code for issue in proposal_findings(proposal, self.index)}
        self.assertIn("issue_form_bypassed", codes)

    def test_personal_data_is_removed_before_it_is_stored(self) -> None:
        values = proposal_values()
        values["selected_problem"] = (
            "Nhóm của 24001101 (an.nguyen@vnu.edu.vn) phân loại chữ số viết tay."
        )
        proposal = one_proposal(3, values)
        self.assertNotIn("24001101", proposal.values["selected_problem"])
        self.assertNotIn("vnu.edu.vn", proposal.values["selected_problem"])
        self.assertTrue(proposal.redactions)
        codes = {issue.code for issue in proposal_findings(proposal, self.index)}
        self.assertIn("issue_personal_data", codes)

    def test_a_pending_membership_request_is_reported_not_applied(self) -> None:
        proposal = one_proposal(
            3, comments=_comments([comment_row(13, MEMBERSHIP_COMMENT)]), staff_logins=STAFF
        )
        by_code = {issue.code: issue for issue in proposal_findings(proposal, self.index)}
        self.assertIn("membership_change_pending", by_code)
        self.assertIn("an-nguyen, bao-le", by_code["membership_change_pending"].found)
        # The proposal still names all three: the request is staff's to decide.
        self.assertEqual(proposal.members, ("an-nguyen", "bao-le", "chi-pham"))

    def test_a_rewritten_body_with_no_announcement_is_reported(self) -> None:
        proposal = one_proposal(3)
        findings = proposal_findings(proposal, self.index, previous_hash="0" * 64)
        by_code = {issue.code: issue for issue in findings}
        self.assertIn("issue_body_edited", by_code)
        self.assertEqual(
            by_code["issue_body_edited"].found, proposal.content_hash[:12]
        )

    def test_an_announced_rewrite_is_not_reported(self) -> None:
        proposal = one_proposal(
            3,
            comments=_comments([comment_row(11, PROPOSAL_UPDATE_COMMENT)]),
            staff_logins=STAFF,
        )
        findings = proposal_findings(proposal, self.index, previous_hash="0" * 64)
        self.assertNotIn(
            "issue_body_edited", {issue.code for issue in findings}
        )

    def test_a_bumped_timestamp_with_unchanged_content_is_not_an_edit(self) -> None:
        body = render_body(proposal_values())
        proposal = parse_proposal(
            _raw(
                issue_row(
                    3,
                    body=body,
                    created_at="2026-09-21T02:00:00Z",
                    updated_at="2026-10-02T09:00:00Z",
                )
            ),
            SCHEMA,
            known_repos=KNOWN_REPOS,
            expected_org="VNU-HUS",
        )
        findings = proposal_findings(
            proposal, self.index, previous_hash=proposal.content_hash
        )
        self.assertEqual(findings, [])

    def test_a_last_edited_timestamp_alone_raises_the_finding(self) -> None:
        from course_hoanganhduc.project_issues import EditRecord

        proposal = one_proposal(
            3,
            edit=EditRecord(
                number=3,
                last_edited_at="2026-10-02T09:00:00Z",
                edit_count=3,
                edited_comments=(),
            ),
        )
        codes = {issue.code for issue in proposal_findings(proposal, self.index)}
        self.assertIn("issue_body_edited", codes)

    def test_every_finding_carries_a_fix_the_student_can_act_on(self) -> None:
        values = proposal_values()
        values["overlap"] = NO_RESPONSE
        values["repository_url"] = "https://github.com/some-other-org/random-repo"
        values["selected_problem"] = "Liên hệ 24001101 để biết thêm."
        proposal = one_proposal(
            3, values, comments=_comments([comment_row(13, MEMBERSHIP_COMMENT)]),
            staff_logins=STAFF,
        )
        findings = proposal_findings(proposal, self.index, previous_hash="0" * 64)
        self.assertTrue(findings)
        for issue in findings:
            with self.subTest(code=issue.code):
                self.assertTrue(issue.fix.strip(), issue.code)
                self.assertTrue(issue.detail.strip(), issue.code)
                self.assertIn(issue.severity, (SEVERITY_ERROR, SEVERITY_WARNING))

    def test_the_consequential_codes_carry_the_errors_this_lane_raises(self) -> None:
        for code in (
            "issue_member_shared",
            "issue_group_mismatch",
            "issue_founder_mismatch",
            "issue_repo_unknown",
        ):
            with self.subTest(code=code):
                self.assertIn(code, CONSEQUENTIAL_CODES)
        for code in ("issue_form_bypassed", "issue_personal_data", "issue_body_edited"):
            with self.subTest(code=code):
                self.assertNotIn(code, CONSEQUENTIAL_CODES)


# ---------------------------------------------------------------------------
# proposal.md
# ---------------------------------------------------------------------------


def proposal_md(
    *,
    title: str = "Nhận dạng chữ số viết tay",
    repeated: bool = False,
) -> str:
    """The pinned ``proposal/proposal.md``, in its marker form."""
    blocks = [
        ("PROJECT_TITLE", title),
        ("TOPIC_SOURCE", "Self-proposed"),
        ("PUBLIC_SUMMARY", "Bộ phân loại chữ số viết tay."),
        ("SELECTED_PROBLEM", "Phân loại chữ số viết tay từ ảnh quét."),
        ("SCOPE_AND_FEASIBILITY", "Bốn tuần, dữ liệu có sẵn."),
        ("PLANNED_METHOD", "Hồi quy logistic."),
        ("DATA_TOOLS_AND_SOURCES", "numpy, dữ liệu của lớp."),
        ("EXPECTED_OUTPUT", "Notebook và bảng kết quả."),
        ("PROBLEM_AND_MOTIVATION", "Bài toán nhập liệu thủ công."),
        ("MILESTONES", "Tuần 1 dữ liệu, tuần 2 mô hình."),
        ("REFERENCES", "Giáo trình môn học."),
        ("CONSIDERATIONS", "Không có dữ liệu cá nhân."),
    ]
    text = "# Proposal\n\n"
    for name, value in blocks:
        text += f"<!-- BEGIN:{name} -->\n{value}\n<!-- END:{name} -->\n\n"
    if repeated:
        text += "<!-- BEGIN:PROJECT_TITLE -->\nChủ đề khác\n<!-- END:PROJECT_TITLE -->\n"
    return text


class TestProposalFile(unittest.TestCase):
    def test_the_markers_are_read_out_of_the_pinned_file(self) -> None:
        markers = parse_proposal_markers(proposal_md())
        self.assertEqual(markers["PROJECT_TITLE"], "Nhận dạng chữ số viết tay")
        self.assertEqual(markers["DATA_TOOLS_AND_SOURCES"], "numpy, dữ liệu của lớp.")

    def test_a_marker_that_appears_twice_is_dropped(self) -> None:
        markers = parse_proposal_markers(proposal_md(repeated=True))
        self.assertNotIn("PROJECT_TITLE", markers)
        self.assertIn("SELECTED_PROBLEM", markers)

    def test_one_form_field_is_fed_by_two_markers(self) -> None:
        values = values_from_markers(parse_proposal_markers(proposal_md()))
        self.assertIn("Hồi quy logistic.", values["method_resources"])
        self.assertIn("numpy", values["method_resources"])

    def test_markers_without_a_form_field_are_not_invented_into_one(self) -> None:
        values = values_from_markers(parse_proposal_markers(proposal_md()))
        self.assertNotIn("overlap", values)
        self.assertNotIn("CONSIDERATIONS", values)
        self.assertNotIn("MILESTONES", values)

    def test_the_pinned_file_overrides_the_issue_body(self) -> None:
        markers = parse_proposal_markers(proposal_md())
        proposal = one_proposal(3, markers=markers)
        self.assertEqual(proposal.project_title, "Nhận dạng chữ số viết tay")
        self.assertEqual(proposal.sources["project_title"], SOURCE_PROPOSAL_MD)
        # A field with no marker keeps the issue as its source.
        self.assertEqual(proposal.sources["overlap"], SOURCE_ISSUE)

    def test_the_file_is_read_raw_and_pinned_to_the_commit(self) -> None:
        from course_hoanganhduc.project_issues import read_proposal_markers

        runner = FakeRunner(proposals={(REPO_AN, SHA_ONE): proposal_md()})
        markers = read_proposal_markers(
            REPO_AN, SHA_ONE, runner=runner, sleeper=lambda _s: None, pace=0.0
        )
        self.assertEqual(markers["PROJECT_TITLE"], "Nhận dạng chữ số viết tay")
        argv = runner.calls[0]
        self.assertEqual(argv[2:4], ["-H", "Accept: application/vnd.github.raw"])
        self.assertTrue(argv[-1].endswith(f"?ref={SHA_ONE}"))

    def test_a_blob_sha_in_an_envelope_is_never_taken_for_a_commit(self) -> None:
        """The contents API's ``sha`` names the blob, not the commit holding it.

        The lane asks for the raw file, so the envelope should never appear at
        all.  A proxy or an older ``gh`` can still hand one back, and the blob
        sha is forty hex characters like any commit, so its shape cannot save
        us: read as file text it yields no markers, which is the right answer.
        """
        from course_hoanganhduc.project_issues import (
            normalize_commit,
            read_proposal_markers,
        )

        envelope = json.dumps(
            {
                "name": "proposal.md",
                "path": "proposal/proposal.md",
                "sha": BLOB_SHA,
                "encoding": "base64",
                "content": base64.b64encode(proposal_md().encode("utf-8")).decode(),
            }
        )
        runner = FakeRunner(proposals={(REPO_AN, SHA_ONE): envelope})
        markers = read_proposal_markers(
            REPO_AN, SHA_ONE, runner=runner, sleeper=lambda _s: None, pace=0.0
        )
        self.assertEqual(markers, {})
        self.assertEqual(normalize_commit(BLOB_SHA), BLOB_SHA)  # shape says nothing

    def test_a_missing_file_is_not_an_error(self) -> None:
        from course_hoanganhduc.project_issues import read_proposal_markers

        self.assertEqual(
            read_proposal_markers(
                REPO_AN, SHA_TWO, runner=FakeRunner(), sleeper=lambda _s: None, pace=0.0
            ),
            {},
        )

    def test_a_value_that_is_not_a_commit_sha_is_refused(self) -> None:
        from course_hoanganhduc.project_issues import read_proposal_markers

        with self.assertRaises(Classroom50Error) as caught:
            read_proposal_markers(REPO_AN, SHA_ONE[:7], runner=FakeRunner())
        self.assertEqual(caught.exception.code, "bad_commit")


# ---------------------------------------------------------------------------
# the lane, end to end
# ---------------------------------------------------------------------------


class TestLane(unittest.TestCase):
    def run_lane(
        self,
        runner: FakeRunner,
        students: Optional[List[Student]] = None,
        **kwargs: Any,
    ) -> Any:
        return import_project_issues(
            students if students is not None else klass(),
            board=BOARD,
            known_repos=KNOWN_REPOS,
            staff_logins=STAFF,
            runner=runner,
            sleeper=lambda _s: None,
            pace=0.0,
            **kwargs,
        )

    def test_a_board_that_is_not_owner_slash_repo_is_refused(self) -> None:
        with self.assertRaises(Classroom50Error) as caught:
            import_project_issues([], board="mat1206e-2026-project-topics")
        self.assertEqual(caught.exception.code, "bad_board")

    def test_todays_board_reads_clean_and_writes_nothing(self) -> None:
        students = klass()
        before = [dict(student.__dict__) for student in students]
        runner = FakeRunner(issues=[STAFF_ISSUE], edits=[edit_node(1)])
        report = self.run_lane(runner, students)
        self.assertEqual(report.total, 1)
        self.assertEqual(report.proposals, ())
        self.assertEqual(report.canonical, ())
        self.assertEqual(report.updated, 0)
        self.assertEqual(len(report.skipped), 1)
        self.assertEqual(report.skipped[0].number, 1)
        self.assertIn("không khớp mẫu proposal", report.skipped[0].reason)
        self.assertEqual([dict(s.__dict__) for s in students], before)

    def test_the_listing_asks_for_closed_issues_too(self) -> None:
        runner = FakeRunner(issues=[STAFF_ISSUE], edits=[edit_node(1)])
        self.run_lane(runner)
        listing = [path for path in runner.api_paths() if "/issues?" in path]
        self.assertTrue(listing)
        for path in listing:
            self.assertIn("state=all", path)

    def test_pull_requests_are_not_proposals(self) -> None:
        body = render_body(proposal_values())
        pull = issue_row(6, body=body, title="[Proposal] nhầm PR")
        pull["pull_request"] = {"url": "https://api.github.com/…/pulls/6"}
        runner = FakeRunner(
            issues=[STAFF_ISSUE, pull, issue_row(7, body=body)],
            edits=[edit_node(1), edit_node(7)],
        )
        report = self.run_lane(runner)
        self.assertEqual(report.total, 2)
        self.assertEqual([p.number for p in report.proposals], [7])

    def test_a_full_pass_writes_the_topic_onto_every_member(self) -> None:
        students = klass()
        body = render_body(proposal_values())
        runner = FakeRunner(
            issues=[STAFF_ISSUE, issue_row(7, body=body)],
            comments={7: [comment_row(11, FINAL_COMMENT)]},
            edits=[edit_node(1), edit_node(7)],
        )
        report = self.run_lane(runner, students)
        self.assertEqual(report.matched, ("an-nguyen", "bao-le", "chi-pham"))
        self.assertEqual(report.unmatched, ())
        self.assertEqual(report.updated, 3)
        an = students[0]
        self.assertEqual(getattr(an, FIELD_NUMBER), 7)
        self.assertEqual(getattr(an, FIELD_REPO), REPO_AN)
        self.assertEqual(getattr(an, FIELD_TITLE), "Nhận dạng chữ viết tay")
        self.assertEqual(getattr(an, FIELD_TOPIC_SOURCE), "Self-proposed")
        self.assertEqual(getattr(an, FIELD_GROUP_NAME), "Nhóm Alpha")
        self.assertEqual(getattr(an, FIELD_MEMBERS), "an-nguyen, bao-le, chi-pham")
        self.assertEqual(getattr(an, FIELD_STATUS), LABEL_SUBMITTED)
        self.assertEqual(getattr(an, FIELD_FINAL), "yes")
        # The final commit is a different fact from the proposal commit, and
        # announcing the former must not silently rewrite the latter.
        self.assertEqual(getattr(an, FIELD_COMMIT), SHA_ONE)
        self.assertTrue(getattr(an, FIELD_HASH))
        # nobody outside the proposal is touched
        self.assertFalse(hasattr(students[3], FIELD_NUMBER))

    def test_a_student_is_joined_by_username_and_never_by_number(self) -> None:
        students = klass()
        # Same student number as the proposal's founder, different account.
        setattr(students[0], "GitHub Username", "an-nguyen-2")
        body = render_body(proposal_values())
        runner = FakeRunner(
            issues=[issue_row(7, body=body)], edits=[edit_node(7)]
        )
        report = self.run_lane(runner, students)
        self.assertEqual(report.unmatched, ("an-nguyen",))
        self.assertFalse(hasattr(students[0], FIELD_NUMBER))
        self.assertEqual(len(students), 5)

    def test_a_quarantined_student_keeps_the_topic_and_loses_the_group(self) -> None:
        students = klass()
        body_an = render_body(proposal_values())
        body_binh = render_body(
            proposal_values(
                repo=REPO_BINH,
                founder="binh-tran",
                members=("binh-tran", "dung-vo", "chi-pham"),
                group_name="Nhóm Beta",
                project_title="Tìm đường trên bản đồ",
            )
        )
        runner = FakeRunner(
            issues=[issue_row(7, body=body_an), issue_row(8, body=body_binh)],
            edits=[edit_node(7), edit_node(8)],
        )
        report = self.run_lane(runner, students)
        self.assertIn("issue_member_shared", {issue.code for issue in report.findings})
        self.assertEqual(len(report.quarantined), 5)
        chi = students[2]
        self.assertTrue(getattr(chi, FIELD_TITLE))
        self.assertTrue(getattr(chi, FIELD_NUMBER))
        self.assertFalse(hasattr(chi, FIELD_REPO))
        self.assertFalse(hasattr(chi, FIELD_GROUP_NAME))
        self.assertFalse(hasattr(chi, FIELD_MEMBERS))

    def test_the_pinned_file_is_read_only_for_a_valid_proposal(self) -> None:
        body = render_body(proposal_values())
        runner = FakeRunner(
            issues=[issue_row(7, body=body)],
            edits=[edit_node(7)],
            proposals={(REPO_AN, SHA_ONE): proposal_md()},
        )
        report = self.run_lane(runner, with_proposal_md=True)
        self.assertEqual(report.proposals[0].project_title, "Nhận dạng chữ số viết tay")
        self.assertTrue(
            any(PROPOSAL_PATH in path for path in runner.api_paths())
        )

    def test_a_blob_sha_never_reaches_the_database_or_the_report(self) -> None:
        """The pin travels one way: the issue's commit into ``?ref=``.

        Nothing the contents call answers with is allowed back out as a commit,
        so an envelope's blob sha must appear in neither the stored field nor
        the report even when the file itself is served.
        """
        students = klass()
        body = render_body(proposal_values())
        envelope = json.dumps(
            {
                "path": "proposal/proposal.md",
                "sha": BLOB_SHA,
                "encoding": "base64",
                "content": base64.b64encode(proposal_md().encode("utf-8")).decode(),
            }
        )
        runner = FakeRunner(
            issues=[issue_row(7, body=body)],
            edits=[edit_node(7)],
            proposals={(REPO_AN, SHA_ONE): envelope},
        )
        report = self.run_lane(runner, students, with_proposal_md=True)
        an = next(s for s in students if getattr(s, "GitHub Username", "") == "an-nguyen")
        self.assertEqual(getattr(an, FIELD_COMMIT), SHA_ONE)
        self.assertNotIn(BLOB_SHA, json.dumps(report_to_dict(report), ensure_ascii=False))
        pinned = [path for path in runner.api_paths() if PROPOSAL_PATH in path]
        self.assertEqual(len(pinned), 1)
        self.assertTrue(pinned[0].endswith(f"?ref={SHA_ONE}"))

    def test_no_request_is_made_to_a_repository_the_issue_invented(self) -> None:
        values = proposal_values()
        values["repository_url"] = "https://github.com/some-other-org/random-repo"
        runner = FakeRunner(
            issues=[issue_row(7, body=render_body(values))], edits=[edit_node(7)]
        )
        report = self.run_lane(runner, with_proposal_md=True)
        self.assertIn("issue_repo_unknown", {issue.code for issue in report.findings})
        for path in runner.paths():
            self.assertNotIn("some-other-org", path)
        self.assertFalse(any(PROPOSAL_PATH in path for path in runner.api_paths()))

    @staticmethod
    def _partial_then_error(message: str, stderr: str) -> RunResult:
        """What a broken paginated read actually prints: data, then an error.

        The rows parse, so parsing is not evidence the call succeeded; only the
        exit code is.
        """
        body = render_body(proposal_values())
        return RunResult(
            1,
            json.dumps([issue_row(7, body=body)])
            + json.dumps(
                {
                    "message": message,
                    "documentation_url": "https://docs.github.com/rest",
                    "status": "403",
                }
            ),
            stderr,
        )

    def test_a_partial_page_followed_by_an_error_object_stops_the_pass(self) -> None:
        students = klass()
        body = render_body(proposal_values())
        runner = FakeRunner(issues=[issue_row(7, body=body)], edits=[edit_node(7)])
        runner.failures["/issues?"] = self._partial_then_error(
            "Resource not accessible by personal access token",
            "gh: Resource not accessible by personal access token (HTTP 403)",
        )
        with self.assertRaises(Classroom50Error) as caught:
            self.run_lane(runner, students)
        self.assertIn("fetch_issues", caught.exception.code)
        self.assertEqual(len(runner.api_paths()), 2)  # the form, then one refusal
        self.assertFalse(any(hasattr(s, FIELD_NUMBER) for s in students))

    def test_a_secondary_rate_limit_is_retried_and_then_named(self) -> None:
        students = klass()
        body = render_body(proposal_values())
        runner = FakeRunner(issues=[issue_row(7, body=body)], edits=[edit_node(7)])
        runner.failures["/issues?"] = self._partial_then_error(
            "API rate limit exceeded",
            "gh: API rate limit exceeded for user ID 700 (HTTP 403)",
        )
        with self.assertRaises(Classroom50Error) as caught:
            self.run_lane(runner, students)
        self.assertEqual(caught.exception.code, "gh_rate_limited")
        # The two 403s take different branches: this one was retried.
        self.assertGreater(len(runner.api_paths()), 2)
        self.assertFalse(any(hasattr(s, FIELD_NUMBER) for s in students))

    def test_an_unreadable_edit_history_is_a_warning_not_a_stop(self) -> None:
        body = render_body(proposal_values())
        runner = FakeRunner(issues=[issue_row(7, body=body)], edits=None)
        report = self.run_lane(runner)
        self.assertEqual(len(report.proposals), 1)
        self.assertTrue(
            any("lịch sử sửa issue" in warning for warning in report.warnings)
        )

    def test_a_pinned_schema_says_so_in_the_report(self) -> None:
        body = render_body(proposal_values())
        runner = FakeRunner(
            form=None, issues=[issue_row(7, body=body)], edits=[edit_node(7)]
        )
        report = self.run_lane(runner)
        self.assertEqual(report.schema_source, SCHEMA_PINNED)
        self.assertTrue(any(FORM_PATH in warning for warning in report.warnings))

    def test_the_snapshot_of_one_run_feeds_the_edit_check_of_the_next(self) -> None:
        students = klass()
        body = render_body(proposal_values())
        runner = FakeRunner(
            issues=[issue_row(7, body=body)], edits=[edit_node(7)]
        )
        first = self.run_lane(runner, students)
        self.assertEqual(
            [issue.code for issue in first.findings], []
        )
        snapshot = issue_snapshot(students)
        self.assertEqual(snapshot[7], first.proposals[0].content_hash)

        changed = proposal_values(project_title="Nhận dạng chữ viết tay tiếng Việt")
        runner2 = FakeRunner(
            issues=[issue_row(7, body=render_body(changed))], edits=[edit_node(7)]
        )
        second = self.run_lane(runner2, students, previous=snapshot)
        self.assertIn("issue_body_edited", {issue.code for issue in second.findings})

    def test_the_listing_pass_reads_nothing_from_the_edit_history(self) -> None:
        runner = FakeRunner(issues=[STAFF_ISSUE])
        report = list_project_issues(
            BOARD, staff_logins=STAFF, runner=runner, sleeper=lambda _s: None, pace=0.0
        )
        self.assertEqual(report.updated, 0)
        self.assertFalse(any(call[2:3] == ["graphql"] for call in runner.calls))

    def test_every_call_carries_the_deadline(self) -> None:
        body = render_body(proposal_values())
        runner = FakeRunner(
            issues=[issue_row(7, body=body)],
            comments={7: [comment_row(11, FINAL_COMMENT)]},
            edits=[edit_node(7)],
            proposals={(REPO_AN, SHA_ONE): proposal_md()},
        )
        self.run_lane(runner, timeout=12.5, with_proposal_md=True)
        self.assertTrue(runner.timeouts)
        self.assertEqual(set(runner.timeouts), {12.5})

    def test_the_report_serialises_with_its_findings_attached(self) -> None:
        body = render_body(proposal_values())
        runner = FakeRunner(issues=[issue_row(7, body=body)], edits=[edit_node(7)])
        report = self.run_lane(runner, previous={7: "0" * 64})
        payload = report_to_dict(report)
        self.assertEqual(payload["board"], BOARD)
        self.assertTrue(payload["findings"])
        self.assertIn("student", payload["findings"][0])
        json.dumps(payload, ensure_ascii=False)
        self.assertIn(BOARD, format_project_issues(report))


# ---------------------------------------------------------------------------
# source hygiene
# ---------------------------------------------------------------------------


class TestHygiene(unittest.TestCase):
    """What the module is allowed to say, checked against its own syntax tree."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.path = REPO_ROOT / "course_hoanganhduc" / "project_issues.py"
        cls.source = cls.path.read_text(encoding="utf-8")
        cls.tree = ast.parse(cls.source)
        cls.constants = [
            node.value
            for node in ast.walk(cls.tree)
            if isinstance(node, ast.Constant) and isinstance(node.value, str)
        ]

    def test_the_module_builds_no_command_but_gh_api(self) -> None:
        for node in ast.walk(self.tree):
            if not isinstance(node, (ast.List, ast.Tuple)):
                continue
            first = node.elts[0] if node.elts else None
            if not isinstance(first, ast.Constant) or first.value != "gh":
                continue
            second = node.elts[1] if len(node.elts) > 1 else None
            self.assertIsInstance(second, ast.Constant)
            self.assertEqual(second.value, "api")

    def test_no_bulk_pagination_and_no_aggregating_filter(self) -> None:
        for value in self.constants:
            with self.subTest(value=value[:60]):
                self.assertNotIn("--paginate", value)
                self.assertNotIn("--jq", value)

    def test_the_module_never_asks_for_a_write(self) -> None:
        forbidden = (
            "--method",
            "-X PUT",
            "-X POST",
            "-X PATCH",
            "-X DELETE",
            "--add-label",
            "--remove-label",
            "gh issue",
            "gh teacher",
            "gh student",
        )
        for value in self.constants:
            for verb in forbidden:
                with self.subTest(value=value[:60], verb=verb):
                    self.assertNotIn(verb, value)

    def test_the_labels_are_read_and_never_set(self) -> None:
        # The three board labels appear as constants because the lane reports
        # them; none of them may appear beside a label-changing command.
        self.assertIn(LABEL_RECORDED, self.constants)
        for value in self.constants:
            if "labels" in value and value.startswith("repos/"):
                self.fail(f"the module addresses the labels endpoint: {value}")

    def test_the_module_stays_free_of_the_database_layer(self) -> None:
        for node in ast.walk(self.tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    self.assertNotEqual(alias.name.split(".")[0], "pandas")
            elif isinstance(node, ast.ImportFrom):
                self.assertNotIn("data", (node.module or "").split("."))
                self.assertNotEqual((node.module or "").split(".")[0], "pandas")
        self.assertNotIn("sys.modules", self.source)

    def test_importing_the_module_pulls_in_neither_pandas_nor_the_database(self) -> None:
        self.assertNotIn("pandas", sys.modules)
        self.assertNotIn("course_hoanganhduc.data", sys.modules)


if __name__ == "__main__":
    unittest.main(verbosity=2)
