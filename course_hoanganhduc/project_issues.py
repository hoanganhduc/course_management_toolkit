# -*- coding: utf-8 -*-
"""Read the mini-project topic board and record each group's chosen problem.

Students register a final-project topic by opening one Issue Form proposal on
their course's topic board.  This module reads those issues, decides which one
is canonical for each group repository, and writes the topic metadata onto the
students who are already in the database.  Nothing here writes to GitHub: no
label is changed, no comment is posted, no membership request is applied.

Three properties of the board shape almost every decision below.

* **Students can edit their own issues but cannot label them.**  On this board a
  student has ``read``; only staff have ``push``.  So a ``status:`` label is
  trustworthy and the title is not -- the ``[Proposal] `` prefix is student
  editable and is never used as a filter.  An issue is recognised by the shape
  of its body instead: how many of the form's own field headings it carries.
* **The form renders headings from labels, not from field ids.**  The labels
  live in ``.github/ISSUE_TEMPLATE/project-proposal.yml`` and change whenever
  staff edit the form, so the schema is read from the board repository at run
  time and :data:`PINNED_FIELDS` is only the fallback.
* **Every field is ``required: true``.**  There is no optional field on this
  form, so ``_No response_`` cannot occur through normal use; it is recorded as
  a sign the form was bypassed rather than as an ordinary empty answer.

The module is stdlib-only and imports neither :mod:`course_hoanganhduc.data`
nor pandas, so its tests run under a bare interpreter.  The paced, retrying
``gh`` helpers are imported from :mod:`course_hoanganhduc.c50_groups` rather
than copied: a third private implementation of the same backoff would drift
away from the two that already exist.
"""

from __future__ import annotations

import hashlib
import json
import re
import time
import unicodedata
from typing import (
    Any,
    Callable,
    Dict,
    FrozenSet,
    List,
    Mapping,
    NamedTuple,
    Optional,
    Sequence,
    Set,
    Tuple,
)

from .c50_cli import Classroom50Error, Runner
from .c50_groups import (
    DEFAULT_PACE,
    DEFAULT_TIMEOUT,
    FIELD_GROUP,
    MINI_PROJECT_SLUG,
    USERNAME_PATTERN,
    _call,
    _paged,
    timeout_runner,
)
from .c50_scores import _norm_login, _parse_json, _student_index
from .roster_audit import SEVERITY_ERROR, SEVERITY_WARNING, Issue, StudentRef, student_ref

# --------------------------------------------------------------------------
# database fields
# --------------------------------------------------------------------------

FIELD_NUMBER = "MiniProject_Issue_Number"
FIELD_URL = "MiniProject_Issue_URL"
FIELD_STATUS = "MiniProject_Issue_Status"
FIELD_TITLE = "MiniProject_Issue_Title"
FIELD_TOPIC_SOURCE = "MiniProject_Issue_Topic_Source"
FIELD_GROUP_NAME = "MiniProject_Issue_Group_Name"
FIELD_REPO = "MiniProject_Issue_Repo"
FIELD_COMMIT = "MiniProject_Issue_Proposal_Commit"
FIELD_MEMBERS = "MiniProject_Issue_Members"
FIELD_FINAL = "MiniProject_Issue_Final_Submitted"
FIELD_HASH = "MiniProject_Issue_Content_Hash"

#: Fields that describe *who is in the group*.  A quarantined student keeps the
#: value already on record for these instead of taking today's reading.
GROUP_FIELDS: Tuple[str, ...] = (FIELD_GROUP_NAME, FIELD_REPO, FIELD_MEMBERS)

#: Fields that describe *what the group is working on*.  A topic does not
#: depend on who is grouped with whom, so a membership conflict does not hold
#: these back.
TOPIC_FIELDS: Tuple[str, ...] = (
    FIELD_NUMBER,
    FIELD_URL,
    FIELD_STATUS,
    FIELD_TITLE,
    FIELD_TOPIC_SOURCE,
    FIELD_COMMIT,
    FIELD_FINAL,
    FIELD_HASH,
)

# --------------------------------------------------------------------------
# the issue form
# --------------------------------------------------------------------------

FORM_PATH = ".github/ISSUE_TEMPLATE/project-proposal.yml"
PROPOSAL_PATH = "proposal/proposal.md"

SCHEMA_REPO = "repo"
SCHEMA_PINNED = "pinned"

#: The form as read on 2026-09-06, kept only as a fallback for the run where
#: the board's own ``.yml`` cannot be read.  The live file wins whenever it is
#: available, because staff can rename a label at any time.
PINNED_FIELDS: Tuple[Tuple[str, str], ...] = (
    ("group_name", "Group name"),
    ("repository_url", "Exact Classroom50 group-repository URL"),
    ("founder", "Founder GitHub username"),
    ("members", "All member GitHub usernames"),
    ("project_title", "Project title"),
    ("topic_source", "Topic source"),
    ("selected_problem", "Exact selected problem"),
    ("public_summary", "Class-visible summary"),
    ("scope", "Student-defined scope and feasibility"),
    ("method_resources", "Planned method, data, tools, and sources"),
    ("expected_output", "Expected output or demonstration"),
    ("overlap", "Relationship to existing recorded problems"),
    ("proposal_commit", "Exact proposal commit URL or complete SHA"),
    ("confirmations", "Group confirmations"),
)

#: Written by hand from ``proposal/proposal.md`` and asserted by the tests.
#: The names do **not** follow from the field ids: ``scope`` is stored under
#: ``SCOPE_AND_FEASIBILITY``, and ``method_resources`` is split across two
#: markers.  Deriving a marker from ``field_id.upper()`` would silently read
#: the wrong block for four of the seven fields that have one.
MARKER_MAP: Tuple[Tuple[str, Tuple[str, ...]], ...] = (
    ("project_title", ("PROJECT_TITLE",)),
    ("topic_source", ("TOPIC_SOURCE",)),
    ("public_summary", ("PUBLIC_SUMMARY",)),
    ("selected_problem", ("SELECTED_PROBLEM",)),
    ("scope", ("SCOPE_AND_FEASIBILITY",)),
    ("method_resources", ("PLANNED_METHOD", "DATA_TOOLS_AND_SOURCES")),
    ("expected_output", ("EXPECTED_OUTPUT",)),
)

#: Form fields with no counterpart in ``proposal.md``.  ``overlap`` is asked on
#: the form only; the five identity fields live in ``team.json`` instead.
FIELDS_WITHOUT_MARKER: FrozenSet[str] = frozenset(
    {
        "group_name",
        "repository_url",
        "founder",
        "members",
        "overlap",
        "proposal_commit",
        "confirmations",
    }
)

#: Markers with no counterpart on the form.  Reading them would invent answers
#: to questions the proposal never asked.
MARKERS_WITHOUT_FIELD: FrozenSet[str] = frozenset(
    {"PROBLEM_AND_MOTIVATION", "MILESTONES", "REFERENCES", "CONSIDERATIONS"}
)

TOPIC_SOURCE_OPTIONS: Tuple[str, ...] = (
    "Based on one suggested idea",
    "Combines or adapts suggested ideas",
    "Self-proposed",
)

LABEL_SUBMITTED = "status: submitted"
LABEL_RECORDED = "status: recorded"
LABEL_DUPLICATE = "status: duplicate-problem"

NO_RESPONSE = "_no response_"

#: How many form headings a body must carry before it is read as a proposal.
#: Low enough that a half-filled proposal is reported as invalid rather than
#: skipped in silence, high enough that a staff note matches nothing.
PROPOSAL_MIN_HEADINGS = 4

# --------------------------------------------------------------------------
# structured comments
# --------------------------------------------------------------------------

KIND_PROPOSAL_UPDATE = "PROPOSAL UPDATE"
KIND_MEMBERSHIP_CHANGE = "MEMBERSHIP CHANGE REQUEST"
KIND_CUSTODIAN = "ISSUE CUSTODIAN HANDOVER"
KIND_FINAL = "FINAL SUBMISSION"
COMMENT_KINDS: Tuple[str, ...] = (
    KIND_PROPOSAL_UPDATE,
    KIND_MEMBERSHIP_CHANGE,
    KIND_CUSTODIAN,
    KIND_FINAL,
)

#: Every label a structured comment can open a section with, normalised.  The
#: guide puts some values on the same line as the label and some on the lines
#: after it, so one splitter has to accept both shapes.
COMMENT_LABELS: Tuple[str, ...] = (
    "previous proposal commit",
    "new proposal commit",
    "selected problem changed",
    "summary of changes",
    "reason for the changes",
    "current members",
    "proposed members",
    "reason",
    "commit updating team.json",
    "previous custodian",
    "new custodian",
    "all listed members agree",
    "final commit",
    "report",
    "slides",
)

SOURCE_ISSUE = "issue"
SOURCE_PROPOSAL_MD = "proposal.md"
SOURCE_COMMENT = "comment"

SOURCE_STAFF_RECORDED = "staff-recorded"
SOURCE_STAFF_DUPLICATE = "staff-duplicate"
SOURCE_COMPUTED = "computed"

#: Association values GitHub gives an account with write access, used only when
#: the caller cannot supply the staff logins outright.
STAFF_ASSOCIATIONS: FrozenSet[str] = frozenset({"OWNER", "COLLABORATOR"})

# --------------------------------------------------------------------------
# privacy
# --------------------------------------------------------------------------

#: Free-text fields, scrubbed before anything is stored.  The identity fields
#: are deliberately absent: a GitHub username on this board can be all digits
#: and look exactly like a student number, so scrubbing them would delete the
#: very data the lane exists to read.
REDACTED_FIELDS: FrozenSet[str] = frozenset(
    {
        "group_name",
        "project_title",
        "selected_problem",
        "public_summary",
        "scope",
        "method_resources",
        "expected_output",
        "overlap",
    }
)

REDACTION = "[lược bỏ]"

_STUDENT_ID_RE = re.compile(r"\b\d{8}\b")
_EMAIL_RE = re.compile(r"\b[\w.+-]+@[\w-]+\.[\w.-]+\b")

# --------------------------------------------------------------------------
# patterns
# --------------------------------------------------------------------------

_HEADING_RE = re.compile(r"^\s{0,3}(#{2,3})\s+(.+?)\s*#*\s*$")
_FENCE_RE = re.compile(r"^\s{0,3}(`{3,}|~{3,})\s*(\S*)")
_SHA_RE = re.compile(r"^[0-9a-f]{40}$")
_COMMIT_URL_RE = re.compile(
    r"github\.com/[^/\s]+/[^/\s]+/commit/([0-9a-fA-F]{40})", re.IGNORECASE
)
_REPO_URL_RE = re.compile(
    r"^(?:https?://)?(?:www\.)?github\.com/([^/\s]+)/([^/\s?#]+)", re.IGNORECASE
)
_BARE_REPO_RE = re.compile(r"^([A-Za-z0-9._-]+)/([A-Za-z0-9._-]+)$")
_DUPLICATE_RE = re.compile(r"duplicate\s+of\s+#(\d+)", re.IGNORECASE)
_BULLET_RE = re.compile(r"^\s*(?:[-*+]|\d+[.)])\s*")

PER_PAGE = 100
MAX_PAGES = 50

_YES_WORDS: FrozenSet[str] = frozenset({"yes", "y", "co", "true", "1", "dung", "roi"})
_NO_WORDS: FrozenSet[str] = frozenset({"no", "n", "khong", "false", "0", "chua"})

# --------------------------------------------------------------------------
# records
# --------------------------------------------------------------------------


class FormField(NamedTuple):
    """One field of the issue form: the id staff wrote and the label GitHub renders."""

    field_id: str
    label: str


class FormSchema(NamedTuple):
    """The board's current form, plus where it was read from."""

    repo: str
    fields: Tuple[FormField, ...]
    source: str


class RawComment(NamedTuple):
    """One comment as GitHub returned it."""

    comment_id: int
    author: str
    association: str
    body: str
    created_at: str
    updated_at: str


class RawIssue(NamedTuple):
    """One issue as GitHub returned it, before any interpretation."""

    number: int
    title: str
    body: str
    author: str
    association: str
    state: str
    labels: Tuple[str, ...]
    url: str
    created_at: str
    updated_at: str


class EditRecord(NamedTuple):
    """What GraphQL knows about edits to one issue and its comments.

    ``updated_at`` moves for a comment, a label, or a close, so it cannot tell
    an edited body from an untouched one.  ``lastEditedAt`` only moves when a
    body is rewritten, which is the signal this lane needs.
    """

    number: int
    last_edited_at: str
    edit_count: int
    edited_comments: Tuple[int, ...]


class CommentFacts(NamedTuple):
    """What the structured comments on one issue say."""

    kinds: Tuple[str, ...] = ()
    proposal_commit: str = ""
    problem_changed: Optional[bool] = None
    final_submitted: bool = False
    final_commit: str = ""
    membership_pending: bool = False
    membership_proposed: Tuple[str, ...] = ()
    custodian: str = ""
    duplicate_of: Optional[int] = None


class ParsedProposal(NamedTuple):
    """One issue read as a proposal.

    ``invalid`` empty means the proposal can be canonical; a non-empty
    ``invalid`` still keeps every field that did parse, so the report can tell
    the group exactly what to repair.
    """

    number: int
    url: str
    title: str
    state: str
    author: str
    created_at: str
    updated_at: str
    labels: Tuple[str, ...]
    values: Mapping[str, str]
    sources: Mapping[str, str]
    repo: str
    founder: str
    members: Tuple[str, ...]
    commit: str
    group_name: str
    topic_source: str
    project_title: str
    missing: Tuple[str, ...]
    bypass: Tuple[str, ...]
    invalid: Tuple[str, ...]
    redactions: Tuple[str, ...]
    content_hash: str
    comments: CommentFacts = CommentFacts()
    last_edited_at: str = ""
    edit_count: int = 0


class SkippedIssue(NamedTuple):
    """An issue that is not a proposal, and why it was read that way."""

    number: int
    title: str
    reason: str


class CanonicalChoice(NamedTuple):
    """Which issue speaks for one group repository."""

    repo: str
    number: int
    source: str
    superseded: Tuple[int, ...]


class MergeOutcome(NamedTuple):
    """What one write-back pass did, separated from what it read."""

    matched: Tuple[str, ...] = ()
    unmatched: Tuple[str, ...] = ()
    updated: int = 0
    quarantined: Tuple[str, ...] = ()
    warnings: Tuple[str, ...] = ()


class IssueReport(NamedTuple):
    """Everything one pass over the board found."""

    board: str
    schema_source: str
    total: int
    proposals: Tuple[ParsedProposal, ...]
    canonical: Tuple[CanonicalChoice, ...]
    skipped: Tuple[SkippedIssue, ...]
    matched: Tuple[str, ...]
    unmatched: Tuple[str, ...]
    updated: int
    quarantined: Tuple[str, ...]
    warnings: Tuple[str, ...]
    findings: Tuple[Issue, ...]


# --------------------------------------------------------------------------
# small helpers
# --------------------------------------------------------------------------


def _fold(value: str) -> str:
    """A string with its accents removed and its case dropped, for comparison."""
    decomposed = unicodedata.normalize("NFD", str(value or ""))
    stripped = "".join(ch for ch in decomposed if not unicodedata.combining(ch))
    return stripped.casefold().strip()


def normalize_label(value: str) -> str:
    """A form label as a lookup key: folded, unpunctuated, single-spaced."""
    text = _fold(value).rstrip(":").strip()
    return re.sub(r"\s+", " ", text)


def normalize_yes_no(value: str) -> Optional[bool]:
    """``True``/``False`` for an answer meant as yes or no, ``None`` otherwise.

    The guide asks for ``yes/no`` in prose, so the answer arrives as ``Yes``,
    ``YES``, ``có``, ``no.`` or a bare ``N``.  Comparing against the literal
    ``"yes"`` reads most of those as "not yes", which is a different claim.
    """
    text = _fold(value).strip().strip(".,;:!?()[]{}\"'")
    if not text:
        return None
    first = text.split()[0] if text.split() else text
    if first in _YES_WORDS:
        return True
    if first in _NO_WORDS:
        return False
    return None


def normalize_repo_url(value: str) -> str:
    """A repository as ``owner/name`` in lower case, or ``""`` when unusable.

    Students type this by hand, so it arrives with a trailing slash, a ``.git``
    suffix, no scheme, or in the owner's original casing.  All four are the same
    repository and have to land on the same key.
    """
    text = str(value or "").strip().strip("<>").strip()
    if not text:
        return ""
    text = text.split()[0]
    match = _REPO_URL_RE.match(text)
    if match is None:
        match = _BARE_REPO_RE.match(text.rstrip("/"))
    if match is None:
        return ""
    owner, name = match.group(1), match.group(2)
    name = name.rstrip("/")
    if name.lower().endswith(".git"):
        name = name[:-4]
    if not owner or not name:
        return ""
    return f"{owner.lower()}/{name.lower()}"


def normalize_commit(value: str) -> str:
    """The 40-character commit SHA behind an answer, or ``""``.

    The form accepts either a commit URL or a complete SHA.  A short SHA is
    rejected on purpose: it is ambiguous, and the whole point of the field is to
    pin one exact revision.
    """
    text = str(value or "").strip().strip("<>").strip()
    if not text:
        return ""
    match = _COMMIT_URL_RE.search(text)
    if match is not None:
        return match.group(1).lower()
    candidate = text.split()[0].lower()
    if _SHA_RE.match(candidate):
        return candidate
    return ""


def parse_members(value: str) -> Tuple[Tuple[str, ...], Tuple[str, ...]]:
    """The usernames in a members answer, and the tokens that are not usernames.

    The textarea asks for one username per line, so the answer comes back with
    bullets, ``@`` prefixes, commas, or all three.  Anything left that fails
    GitHub's own username rule is handed back separately rather than stored as
    if it were an account.
    """
    logins: List[str] = []
    rejected: List[str] = []
    seen: Set[str] = set()
    for chunk in re.split(r"[\n,;]+", str(value or "")):
        token = _BULLET_RE.sub("", chunk).strip().strip("`").strip()
        if token.startswith("@"):
            token = token[1:]
        token = token.strip()
        if not token:
            continue
        if not USERNAME_PATTERN.match(token):
            rejected.append(token)
            continue
        login = _norm_login(token)
        if login in seen:
            continue
        seen.add(login)
        logins.append(login)
    return tuple(logins), tuple(rejected)


def redact_personal(text: str) -> Tuple[str, int, int]:
    """``text`` with student numbers and addresses removed, and how many of each.

    The form forbids personal data outright, so anything matching is there by
    mistake.  The counts, not the values, go into the report: quoting back the
    student number would put it in the very file that is supposed not to hold
    it.
    """
    body = str(text or "")
    body, emails = _EMAIL_RE.subn(REDACTION, body)
    body, ids = _STUDENT_ID_RE.subn(REDACTION, body)
    return body, ids, emails


def content_hash(values: Mapping[str, str]) -> str:
    """A stable digest of one proposal's parsed answers.

    ``updated_at`` moves for reasons that have nothing to do with the text, so
    the digest is what tells a rewritten proposal from an untouched one across
    two runs.
    """
    payload = json.dumps(
        {key: values[key] for key in sorted(values)},
        ensure_ascii=False,
        sort_keys=True,
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _fence_mask(lines: Sequence[str]) -> List[bool]:
    """For each line, whether it sits inside a fenced code block.

    The board's own guide prints its comment templates inside ``text`` fences,
    so students copy the fences along with the text.  That gives the two parsers
    in this module opposite rules -- the body parser must not see a heading
    inside a fence, the comment parser must still read labels inside one -- and
    both need to know where the fences are.
    """
    mask: List[bool] = []
    marker = ""
    for line in lines:
        match = _FENCE_RE.match(line)
        if marker:
            mask.append(True)
            if (
                match is not None
                and match.group(1)[0] == marker[0]
                and len(match.group(1)) >= len(marker)
                and not match.group(2)
            ):
                marker = ""
            continue
        if match is not None:
            marker = match.group(1)
            mask.append(True)
            continue
        mask.append(False)
    return mask


# --------------------------------------------------------------------------
# the form schema
# --------------------------------------------------------------------------


def parse_form_schema(text: str, *, repo: str = "") -> Tuple[FormField, ...]:
    """The ``(id, label)`` pairs of an issue form, read without a YAML library.

    The file has one shape and this reader accepts only that shape: a body item
    opens at two spaces, its ``id`` sits at four, and its label at six.  A
    checkbox *option* label also spells ``label:`` but sits at eight behind a
    ``- ``, so requiring exactly six spaces and no dash is what keeps the five
    confirmation checkboxes from being read as five extra fields.
    """
    fields: List[FormField] = []
    field_id = ""
    label = ""
    in_attributes = False
    for line in text.splitlines():
        if re.match(r"^\s{0,2}- type:", line):
            if field_id and label:
                fields.append(FormField(field_id=field_id, label=label))
            field_id = ""
            label = ""
            in_attributes = False
            continue
        id_match = re.match(r"^ {4}id:\s*(\S+)\s*$", line)
        if id_match is not None:
            field_id = id_match.group(1).strip("\"'")
            continue
        if re.match(r"^ {4}attributes:\s*$", line):
            in_attributes = True
            continue
        if re.match(r"^ {4}\S", line):
            in_attributes = False
        if in_attributes:
            label_match = re.match(r"^ {6}label:\s*(.+?)\s*$", line)
            if label_match is not None and not label:
                label = label_match.group(1).strip().strip("\"'")
    if field_id and label:
        fields.append(FormField(field_id=field_id, label=label))
    return tuple(fields)


def load_form_schema(
    board: str,
    *,
    runner: Optional[Runner] = None,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> FormSchema:
    """The board's current form, falling back to the pinned copy.

    Labels are what the rendered issue carries, and staff can rename one at any
    time, so the live file is the authority.  A board that cannot be read still
    has to be parseable, hence the fallback -- but the report says which of the
    two was used, because a fallback run can silently miss a renamed field.
    """
    runner = runner or timeout_runner(timeout or DEFAULT_TIMEOUT)
    argv = [
        "gh",
        "api",
        "-H",
        "Accept: application/vnd.github.raw",
        f"repos/{board}/contents/{FORM_PATH}",
    ]
    try:
        text = _call(
            runner, argv, op="load_form_schema", timeout=timeout, sleeper=sleeper, pace=pace
        )
    except Classroom50Error:
        return FormSchema(
            repo=board,
            fields=tuple(FormField(field_id=fid, label=lab) for fid, lab in PINNED_FIELDS),
            source=SCHEMA_PINNED,
        )
    fields = parse_form_schema(text, repo=board)
    if not fields:
        return FormSchema(
            repo=board,
            fields=tuple(FormField(field_id=fid, label=lab) for fid, lab in PINNED_FIELDS),
            source=SCHEMA_PINNED,
        )
    return FormSchema(repo=board, fields=fields, source=SCHEMA_REPO)


# --------------------------------------------------------------------------
# reading the board
# --------------------------------------------------------------------------


def _issue_from_payload(row: Mapping[str, Any]) -> RawIssue:
    user = row.get("user") or {}
    labels: List[str] = []
    for label in row.get("labels") or ():
        if isinstance(label, Mapping):
            name = str(label.get("name") or "")
        else:
            name = str(label or "")
        if name:
            labels.append(name)
    return RawIssue(
        number=int(row.get("number") or 0),
        title=str(row.get("title") or ""),
        body=str(row.get("body") or ""),
        author=_norm_login(user.get("login") if isinstance(user, Mapping) else ""),
        association=str(row.get("author_association") or ""),
        state=str(row.get("state") or ""),
        labels=tuple(labels),
        url=str(row.get("html_url") or ""),
        created_at=str(row.get("created_at") or ""),
        updated_at=str(row.get("updated_at") or ""),
    )


def fetch_issues(
    board: str,
    *,
    runner: Optional[Runner] = None,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> List[RawIssue]:
    """Every issue on the board, closed ones included, pull requests excluded.

    Two defaults have to be overridden here.  ``/issues`` answers ``state=open``
    unless told otherwise, and a proposal stays a proposal after staff close it;
    and the same endpoint returns pull requests, which GitHub documents as
    "every pull request is an issue" and marks with a ``pull_request`` key.
    """
    runner = runner or timeout_runner(timeout or DEFAULT_TIMEOUT)
    rows = _paged(
        runner,
        f"repos/{board}/issues?state=all&sort=created&direction=asc",
        op="fetch_issues",
        timeout=timeout,
        sleeper=sleeper,
        pace=pace,
    )
    issues: List[RawIssue] = []
    for row in rows:
        if not isinstance(row, Mapping):
            continue
        if "pull_request" in row:
            continue
        issues.append(_issue_from_payload(row))
    return issues


def fetch_comments(
    board: str,
    number: int,
    *,
    runner: Optional[Runner] = None,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> List[RawComment]:
    """Every comment on one issue, oldest first."""
    runner = runner or timeout_runner(timeout or DEFAULT_TIMEOUT)
    rows = _paged(
        runner,
        f"repos/{board}/issues/{number}/comments",
        op="fetch_comments",
        timeout=timeout,
        sleeper=sleeper,
        pace=pace,
    )
    comments: List[RawComment] = []
    for row in rows:
        if not isinstance(row, Mapping):
            continue
        user = row.get("user") or {}
        comments.append(
            RawComment(
                comment_id=int(row.get("id") or 0),
                author=_norm_login(user.get("login") if isinstance(user, Mapping) else ""),
                association=str(row.get("author_association") or ""),
                body=str(row.get("body") or ""),
                created_at=str(row.get("created_at") or ""),
                updated_at=str(row.get("updated_at") or ""),
            )
        )
    return comments


EDIT_HISTORY_QUERY = """
query($owner: String!, $name: String!, $cursor: String) {
  repository(owner: $owner, name: $name) {
    issues(first: 50, after: $cursor, orderBy: {field: CREATED_AT, direction: ASC}) {
      pageInfo { hasNextPage endCursor }
      nodes {
        number
        lastEditedAt
        userContentEdits(first: 1) { totalCount }
        comments(first: 100) {
          nodes {
            databaseId
            lastEditedAt
          }
        }
      }
    }
  }
}
""".strip()


def fetch_edit_history(
    board: str,
    *,
    runner: Optional[Runner] = None,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> Dict[int, EditRecord]:
    """When each issue and comment was last rewritten, keyed by issue number.

    A content digest only sees changes between two of our own runs, so an edit
    made before the first run is invisible to it.  ``lastEditedAt`` covers that
    gap, and it is the only field on either API that answers "was this body
    rewritten" without being moved by comments, labels, or a close.
    """
    owner, _, name = board.partition("/")
    if not owner or not name:
        raise Classroom50Error(
            f"board must be OWNER/REPO, got {board!r}", code="bad_board"
        )
    runner = runner or timeout_runner(timeout or DEFAULT_TIMEOUT)
    records: Dict[int, EditRecord] = {}
    cursor = ""
    for _page in range(MAX_PAGES):
        argv = [
            "gh",
            "api",
            "graphql",
            "-f",
            f"query={EDIT_HISTORY_QUERY}",
            "-f",
            f"owner={owner}",
            "-f",
            f"name={name}",
        ]
        if cursor:
            argv.extend(["-f", f"cursor={cursor}"])
        text = _call(
            runner,
            argv,
            op="fetch_edit_history",
            timeout=timeout,
            sleeper=sleeper,
            pace=pace,
        )
        payload = _parse_json(text, op="fetch_edit_history", expect=dict)
        repository = ((payload.get("data") or {}).get("repository")) or {}
        block = repository.get("issues") or {}
        for node in block.get("nodes") or ():
            if not isinstance(node, Mapping):
                continue
            edited: List[int] = []
            comments = (node.get("comments") or {}).get("nodes") or ()
            for comment in comments:
                if isinstance(comment, Mapping) and comment.get("lastEditedAt"):
                    edited.append(int(comment.get("databaseId") or 0))
            number = int(node.get("number") or 0)
            records[number] = EditRecord(
                number=number,
                last_edited_at=str(node.get("lastEditedAt") or ""),
                edit_count=int((node.get("userContentEdits") or {}).get("totalCount") or 0),
                edited_comments=tuple(edited),
            )
        info = block.get("pageInfo") or {}
        if not info.get("hasNextPage"):
            break
        cursor = str(info.get("endCursor") or "")
        if not cursor:
            break
    return records


def read_proposal_markers(
    repo: str,
    sha: str,
    *,
    runner: Optional[Runner] = None,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> Dict[str, str]:
    """The marked blocks of ``proposal.md`` at one pinned commit.

    ``repo`` must already have been checked against the real repository list:
    the URL it comes from is student-typed, and this is the one place the lane
    would otherwise follow it.  Returns an empty mapping when the file is not
    there, which is the ordinary state before a group commits one.
    """
    if not _SHA_RE.match(str(sha or "").lower()):
        raise Classroom50Error(
            f"proposal commit must be a 40-character SHA, got {sha!r}",
            code="bad_commit",
        )
    runner = runner or timeout_runner(timeout or DEFAULT_TIMEOUT)
    argv = [
        "gh",
        "api",
        "-H",
        "Accept: application/vnd.github.raw",
        f"repos/{repo}/contents/{PROPOSAL_PATH}?ref={sha.lower()}",
    ]
    try:
        text = _call(
            runner,
            argv,
            op="read_proposal_markers",
            timeout=timeout,
            sleeper=sleeper,
            pace=pace,
        )
    except Classroom50Error as exc:
        lowered = str(exc).lower()
        if "404" in lowered or "not found" in lowered:
            return {}
        raise
    return parse_proposal_markers(text)


def parse_proposal_markers(text: str) -> Dict[str, str]:
    """The text between each ``BEGIN``/``END`` marker pair.

    A marker that appears more than once is dropped rather than guessed at: the
    file's own rule is one block per marker, and picking the first of two would
    quietly choose between two answers the group did not mean to give.
    """
    out: Dict[str, str] = {}
    names = set(MARKERS_WITHOUT_FIELD)
    for _field_id, markers in MARKER_MAP:
        names.update(markers)
    for name in sorted(names):
        pattern = re.compile(
            rf"<!--\s*BEGIN:{re.escape(name)}\s*-->(.*?)<!--\s*END:{re.escape(name)}\s*-->",
            re.DOTALL,
        )
        found = pattern.findall(text)
        if len(found) != 1:
            continue
        out[name] = found[0].strip()
    return out


def values_from_markers(markers: Mapping[str, str]) -> Dict[str, str]:
    """Marker blocks turned into form-field answers, using the hand-written map."""
    out: Dict[str, str] = {}
    for field_id, names in MARKER_MAP:
        parts = [markers[name].strip() for name in names if markers.get(name, "").strip()]
        if not parts:
            continue
        out[field_id] = "\n\n".join(parts)
    return out


# --------------------------------------------------------------------------
# reading one issue body
# --------------------------------------------------------------------------


def parse_issue_body(body: str, schema: FormSchema) -> Dict[str, str]:
    """The answers in a rendered issue body, keyed by form field id.

    Headings are matched by label because that is what GitHub renders; both
    ``##`` and ``###`` are accepted because the board's own example file uses
    the former while the real form produces the latter, and a parser pinned to
    one level matches nothing on half the inputs.  Headings inside a fenced
    block are ignored: no field on this form sets ``render:``, so a pasted code
    block containing ``###`` would otherwise cut a section in two.
    """
    by_label = {normalize_label(field.label): field.field_id for field in schema.fields}
    lines = str(body or "").splitlines()
    mask = _fence_mask(lines)
    sections: Dict[str, List[str]] = {}
    current = ""
    for line, fenced in zip(lines, mask):
        if not fenced:
            heading = _HEADING_RE.match(line)
            if heading is not None:
                field_id = by_label.get(normalize_label(heading.group(2)))
                if field_id is not None:
                    current = field_id
                    sections.setdefault(current, [])
                    continue
                current = ""
                continue
        if current:
            sections[current].append(line)
    return {field_id: "\n".join(body_lines).strip() for field_id, body_lines in sections.items()}


def is_proposal_issue(issue: RawIssue, schema: FormSchema) -> bool:
    """Whether an issue looks like a filled proposal form.

    Neither the title nor the labels can decide this.  A student can rewrite
    their own title, so the ``[Proposal] `` prefix proves nothing; and the
    ``status: submitted`` label is attached by a separate event just after
    creation, so a genuine proposal exists unlabelled for a moment.  The body's
    shape is the only signal that is both present from the start and outside
    the author's reach.
    """
    return len(parse_issue_body(issue.body, schema)) >= PROPOSAL_MIN_HEADINGS


# --------------------------------------------------------------------------
# reading comments
# --------------------------------------------------------------------------


def _label_sections(body: str) -> Dict[str, str]:
    """A structured comment split into its labelled parts.

    The guide writes some answers on the label's own line and some on the lines
    below it, so both are collected: a section runs from its label to the next
    one.  Unlike the body parser this one reads inside fences, because the guide
    prints its templates inside ``text`` fences and students copy them whole.
    """
    lines = str(body or "").splitlines()
    sections: Dict[str, List[str]] = {}
    current = ""
    for line in lines:
        stripped = line.strip().lstrip("*_# ").strip()
        head, sep, rest = stripped.partition(":")
        matched = ""
        if sep:
            key = normalize_label(head)
            for label in COMMENT_LABELS:
                if key == label or key.startswith(label):
                    matched = label
                    break
        if matched:
            current = matched
            sections.setdefault(current, [])
            if rest.strip():
                sections[current].append(rest.strip())
            continue
        if current:
            sections[current].append(line)
    return {key: "\n".join(value).strip() for key, value in sections.items()}


def _comment_kind(body: str) -> str:
    """Which structured template a comment follows, or ``""``."""
    upper = str(body or "").upper()
    for kind in COMMENT_KINDS:
        if kind in upper:
            return kind
    return ""


def parse_comments(
    comments: Sequence[RawComment],
    *,
    staff_logins: FrozenSet[str] = frozenset(),
) -> CommentFacts:
    """What the structured comments on one issue add to its proposal.

    Comments are read oldest first so a later ``PROPOSAL UPDATE`` wins, which is
    what the guide asks for.  A membership request counts as pending until staff
    answer it in the thread; this lane reports it and never applies it, because
    the guide puts that decision with staff.
    """
    staff = frozenset(_norm_login(login) for login in staff_logins if login)
    kinds: List[str] = []
    proposal_commit = ""
    problem_changed: Optional[bool] = None
    final_submitted = False
    final_commit = ""
    membership_pending = False
    membership_proposed: Tuple[str, ...] = ()
    custodian = ""
    duplicate_of: Optional[int] = None
    for comment in comments:
        is_staff = (
            _norm_login(comment.author) in staff
            if staff
            else comment.association.upper() in STAFF_ASSOCIATIONS
        )
        if is_staff:
            found = _DUPLICATE_RE.search(comment.body)
            if found is not None:
                duplicate_of = int(found.group(1))
        kind = _comment_kind(comment.body)
        if not kind:
            continue
        kinds.append(kind)
        sections = _label_sections(comment.body)
        if kind == KIND_PROPOSAL_UPDATE:
            new_commit = normalize_commit(sections.get("new proposal commit", ""))
            if new_commit:
                proposal_commit = new_commit
            changed = normalize_yes_no(sections.get("selected problem changed", ""))
            if changed is not None:
                problem_changed = changed
        elif kind == KIND_MEMBERSHIP_CHANGE:
            if is_staff:
                membership_pending = False
            else:
                membership_pending = True
                membership_proposed = parse_members(sections.get("proposed members", ""))[0]
        elif kind == KIND_CUSTODIAN:
            new_custodian = _norm_login(sections.get("new custodian", "").strip())
            if new_custodian:
                custodian = new_custodian
        elif kind == KIND_FINAL:
            agreed = normalize_yes_no(sections.get("all listed members agree", ""))
            final_commit = normalize_commit(sections.get("final commit", ""))
            final_submitted = agreed is not False
    return CommentFacts(
        kinds=tuple(kinds),
        proposal_commit=proposal_commit,
        problem_changed=problem_changed,
        final_submitted=final_submitted,
        final_commit=final_commit,
        membership_pending=membership_pending,
        membership_proposed=membership_proposed,
        custodian=custodian,
        duplicate_of=duplicate_of,
    )


# --------------------------------------------------------------------------
# reading one proposal
# --------------------------------------------------------------------------


def _ref_for(login: str, index: Mapping[str, Any]) -> StudentRef:
    """The record behind a login, or a placeholder naming the account."""
    student = index.get(_norm_login(login))
    if student is not None:
        return student_ref(student)
    return StudentRef(student_id="", name=f"@{login}", email="", google_id="")


def _issue(
    code: str,
    severity: str,
    login: str,
    index: Mapping[str, Any],
    *,
    field: str,
    found: str,
    detail: str,
    fix: str,
) -> Issue:
    return Issue(
        code=code,
        severity=severity,
        student=_ref_for(login, index),
        field=field,
        found=found,
        detail=detail,
        fix=fix,
    )


def parse_proposal(
    issue: RawIssue,
    schema: FormSchema,
    *,
    comments: Sequence[RawComment] = (),
    staff_logins: FrozenSet[str] = frozenset(),
    known_repos: FrozenSet[str] = frozenset(),
    expected_org: str = "",
    markers: Optional[Mapping[str, str]] = None,
    edit: Optional[EditRecord] = None,
) -> ParsedProposal:
    """One issue read as a proposal, with every problem recorded rather than raised.

    A proposal that fails validation is still returned in full.  The report has
    to be able to tell one group "your repository URL points nowhere" while the
    other twenty-nine proposals are processed normally, and an exception here
    would take the whole run down with the first typo.
    """
    values: Dict[str, str] = dict(parse_issue_body(issue.body, schema))
    sources: Dict[str, str] = {key: SOURCE_ISSUE for key in values}

    bypass = tuple(
        field_id for field_id, text in sorted(values.items()) if _fold(text) == NO_RESPONSE
    )

    if markers:
        for field_id, text in values_from_markers(markers).items():
            if text:
                values[field_id] = text
                sources[field_id] = SOURCE_PROPOSAL_MD

    ids_removed = 0
    emails_removed = 0
    for field_id in sorted(values):
        if field_id not in REDACTED_FIELDS:
            continue
        cleaned, ids, emails = redact_personal(values[field_id])
        values[field_id] = cleaned
        ids_removed += ids
        emails_removed += emails
    redactions: List[str] = []
    if ids_removed:
        redactions.append(f"{ids_removed} chuỗi giống mã sinh viên")
    if emails_removed:
        redactions.append(f"{emails_removed} địa chỉ email")

    facts = parse_comments(comments, staff_logins=staff_logins)

    repo = normalize_repo_url(values.get("repository_url", ""))
    founder = _norm_login(values.get("founder", "").strip().lstrip("@"))
    members, rejected = parse_members(values.get("members", ""))
    commit = normalize_commit(values.get("proposal_commit", ""))
    if facts.proposal_commit:
        commit = facts.proposal_commit

    field_ids = [field.field_id for field in schema.fields]
    missing = tuple(
        field_id
        for field_id in field_ids
        if not values.get(field_id, "").strip() or field_id in bypass
    )

    invalid: List[str] = []
    if missing:
        invalid.append("thiếu trường bắt buộc: " + ", ".join(missing))
    if not repo:
        invalid.append(f"repository_url không đọc được: {values.get('repository_url', '')!r}")
    else:
        owner = repo.split("/", 1)[0]
        if expected_org and owner != _norm_login(expected_org):
            invalid.append(f"repository_url nằm ngoài tổ chức {expected_org}: {repo}")
        elif known_repos and repo not in known_repos:
            invalid.append(f"repository_url không trỏ tới repo nhóm nào đang tồn tại: {repo}")
    if not commit:
        invalid.append(
            f"proposal_commit không phải URL commit hay SHA 40 ký tự: "
            f"{values.get('proposal_commit', '')!r}"
        )
    if founder and not USERNAME_PATTERN.match(founder):
        invalid.append(f"founder không đúng dạng username GitHub: {founder!r}")
    if rejected:
        invalid.append("members có mục không phải username GitHub: " + ", ".join(rejected))
    if not members:
        invalid.append("members không có username GitHub nào đọc được")
    elif founder and founder not in members:
        invalid.append(f"founder {founder!r} không nằm trong members")
    if values.get("topic_source", "").strip() and not any(
        _fold(option) == _fold(values["topic_source"]) for option in TOPIC_SOURCE_OPTIONS
    ):
        invalid.append(f"topic_source ngoài danh sách lựa chọn: {values['topic_source']!r}")

    return ParsedProposal(
        number=issue.number,
        url=issue.url,
        title=issue.title,
        state=issue.state,
        author=issue.author,
        created_at=issue.created_at,
        updated_at=issue.updated_at,
        labels=issue.labels,
        values=values,
        sources=sources,
        repo=repo,
        founder=founder,
        members=members,
        commit=commit,
        group_name=values.get("group_name", "").strip(),
        topic_source=values.get("topic_source", "").strip(),
        project_title=values.get("project_title", "").strip(),
        missing=missing,
        bypass=bypass,
        invalid=tuple(invalid),
        redactions=tuple(redactions),
        content_hash=content_hash(values),
        comments=facts,
        last_edited_at=edit.last_edited_at if edit is not None else "",
        edit_count=edit.edit_count if edit is not None else 0,
    )


# --------------------------------------------------------------------------
# deciding which issue counts
# --------------------------------------------------------------------------


def choose_canonical(proposals: Sequence[ParsedProposal]) -> Dict[str, CanonicalChoice]:
    """One canonical issue per group repository.

    Staff decide this, and their decision is sticky: the guide says that once
    staff record which issue is canonical, an earlier incomplete issue repaired
    later does not displace it.  Recomputing "earliest valid" from scratch on
    every run would do exactly that, so a ``status: recorded`` label or a staff
    ``Duplicate of #n`` comment outranks the ordering, and the ordering only
    decides repositories staff have not touched yet.
    """
    by_repo: Dict[str, List[ParsedProposal]] = {}
    for proposal in proposals:
        if not proposal.repo:
            continue
        by_repo.setdefault(proposal.repo, []).append(proposal)

    out: Dict[str, CanonicalChoice] = {}
    for repo, group in sorted(by_repo.items()):
        ordered = sorted(group, key=lambda item: (item.created_at, item.number))
        numbers = {item.number for item in ordered}
        marked = {item.number for item in ordered if item.comments.duplicate_of is not None}
        pointed = {
            item.comments.duplicate_of
            for item in ordered
            if item.comments.duplicate_of in numbers
        }
        chosen: Optional[ParsedProposal] = None
        source = SOURCE_COMPUTED
        recorded = [item for item in ordered if LABEL_RECORDED in item.labels]
        if recorded:
            chosen = recorded[0]
            source = SOURCE_STAFF_RECORDED
        elif pointed:
            chosen = next(item for item in ordered if item.number in pointed)
            source = SOURCE_STAFF_DUPLICATE
        else:
            for item in ordered:
                if not item.invalid and item.number not in marked:
                    chosen = item
                    break
        if chosen is None:
            continue
        out[repo] = CanonicalChoice(
            repo=repo,
            number=chosen.number,
            source=source,
            superseded=tuple(item.number for item in ordered if item.number != chosen.number),
        )
    return out


# --------------------------------------------------------------------------
# conflicts
# --------------------------------------------------------------------------


def proposal_findings(
    proposal: ParsedProposal, index: Mapping[str, Any], *, previous_hash: str = ""
) -> List[Issue]:
    """What is wrong with one proposal on its own terms.

    Cross-issue problems are not decided here; this is the pass that can run
    with nothing but the issue itself in hand.
    """
    findings: List[Issue] = []
    who = proposal.founder or (proposal.members[0] if proposal.members else "")
    if proposal.invalid and not proposal.repo:
        findings.append(
            _issue(
                "issue_repo_unknown",
                SEVERITY_ERROR,
                who,
                index,
                field="repository_url",
                found=proposal.values.get("repository_url", ""),
                detail=f"Issue #{proposal.number}: {'; '.join(proposal.invalid)}",
                fix=(
                    "Sửa trường 'Exact Classroom50 group-repository URL' trong issue thành "
                    "URL đầy đủ của repo nhóm final-project do Classroom50 tạo."
                ),
            )
        )
    elif any("repository_url" in reason for reason in proposal.invalid):
        findings.append(
            _issue(
                "issue_repo_unknown",
                SEVERITY_ERROR,
                who,
                index,
                field="repository_url",
                found=proposal.repo,
                detail=f"Issue #{proposal.number}: "
                + "; ".join(r for r in proposal.invalid if "repository_url" in r),
                fix=(
                    "Kiểm tra lại URL repo nhóm: nó phải thuộc tổ chức của lớp và là repo "
                    "final-project do Classroom50 tạo cho nhóm bạn."
                ),
            )
        )
    if proposal.bypass:
        findings.append(
            _issue(
                "issue_form_bypassed",
                SEVERITY_WARNING,
                who,
                index,
                field="body",
                found=", ".join(proposal.bypass),
                detail=(
                    f"Issue #{proposal.number} có '_No response_' ở {len(proposal.bypass)} "
                    "trường, trong khi mọi trường của form đều bắt buộc."
                ),
                fix=(
                    "Tạo lại issue bằng đúng mẫu 'Project topic proposal' và điền đủ mọi "
                    "trường, thay vì sửa tay hoặc tạo issue trống."
                ),
            )
        )
    if proposal.redactions:
        findings.append(
            _issue(
                "issue_personal_data",
                SEVERITY_WARNING,
                who,
                index,
                field="body",
                found=", ".join(proposal.redactions),
                detail=(
                    f"Issue #{proposal.number} chứa thông tin cá nhân; phần đó đã bị lược bỏ "
                    "trước khi lưu, nhưng vẫn còn nguyên trên GitHub."
                ),
                fix=(
                    "Sửa issue để bỏ mã sinh viên, email và các thông tin cá nhân khác; "
                    "chỉ dùng GitHub username."
                ),
            )
        )
    if proposal.comments.membership_pending:
        findings.append(
            _issue(
                "membership_change_pending",
                SEVERITY_WARNING,
                who,
                index,
                field="members",
                found=", ".join(proposal.comments.membership_proposed) or "(không đọc được)",
                detail=(
                    f"Issue #{proposal.number} có MEMBERSHIP CHANGE REQUEST chưa được giảng "
                    "viên trả lời; danh sách nhóm giữ nguyên cho tới khi có quyết định."
                ),
                fix="Chờ giảng viên trả lời trong chính issue đó; đừng tự đổi team.json.",
            )
        )
    edited = bool(previous_hash) and previous_hash != proposal.content_hash
    announced = KIND_PROPOSAL_UPDATE in proposal.comments.kinds
    if (edited or proposal.edit_count > 1 or proposal.last_edited_at) and not announced:
        detail = f"Issue #{proposal.number} đã được sửa"
        if proposal.last_edited_at:
            detail += f" (lần cuối {proposal.last_edited_at})"
        findings.append(
            _issue(
                "issue_body_edited",
                SEVERITY_WARNING,
                who,
                index,
                field="body",
                found=proposal.content_hash[:12],
                detail=detail + " mà không có comment PROPOSAL UPDATE kèm theo.",
                fix=(
                    "Mỗi lần đổi nội dung proposal phải kèm một comment PROPOSAL UPDATE nêu "
                    "commit cũ, commit mới và lý do."
                ),
            )
        )
    return findings


def detect_member_conflicts(
    canonical: Mapping[str, CanonicalChoice],
    proposals: Sequence[ParsedProposal],
    index: Mapping[str, Any],
    *,
    groups: Optional[Mapping[str, Tuple[str, Sequence[str]]]] = None,
) -> Tuple[List[Issue], Set[str]]:
    """Conflicts that only show up once every canonical proposal is in view.

    The board's own canonical rule keys on the repository URL, so it cannot see
    a student registered in two different groups: those are two repositories and
    each keeps its own canonical issue.  Inverting the table by member is what
    catches it.
    """
    findings: List[Issue] = []
    quarantine: Set[str] = set()
    by_number = {proposal.number: proposal for proposal in proposals}

    repos_by_login: Dict[str, List[str]] = {}
    for repo, choice in sorted(canonical.items()):
        proposal = by_number.get(choice.number)
        if proposal is None:
            continue
        for login in proposal.members:
            repos_by_login.setdefault(login, []).append(repo)

    for login, repos in sorted(repos_by_login.items()):
        if len(repos) < 2:
            continue
        findings.append(
            _issue(
                "issue_member_shared",
                SEVERITY_ERROR,
                login,
                index,
                field="members",
                found=", ".join(sorted(repos)),
                detail=(
                    f"@{login} có tên trong proposal của {len(repos)} nhóm khác nhau; "
                    "mỗi sinh viên chỉ được thuộc một nhóm."
                ),
                fix=(
                    "Các nhóm liên quan thống nhất xem bạn thuộc nhóm nào, rồi sửa issue và "
                    "team.json của nhóm còn lại cho khớp."
                ),
            )
        )
        quarantine.add(login)
        for repo in repos:
            choice = canonical.get(repo)
            proposal = by_number.get(choice.number) if choice is not None else None
            if proposal is not None:
                quarantine.update(proposal.members)

    if not groups:
        return findings, quarantine

    for repo, choice in sorted(canonical.items()):
        proposal = by_number.get(choice.number)
        if proposal is None:
            continue
        entry = groups.get(repo)
        if entry is None:
            continue
        group_founder, group_members = entry
        recorded = {_norm_login(login) for login in group_members}
        claimed = set(proposal.members)
        if recorded and recorded != claimed:
            extra = sorted(claimed - recorded)
            absent = sorted(recorded - claimed)
            detail = f"Issue #{proposal.number} khai danh sách khác với nhóm trên Classroom50"
            if extra:
                detail += "; thừa: " + ", ".join(extra)
            if absent:
                detail += "; thiếu: " + ", ".join(absent)
            findings.append(
                _issue(
                    "issue_group_mismatch",
                    SEVERITY_ERROR,
                    proposal.founder or repo,
                    index,
                    field="members",
                    found=", ".join(sorted(claimed)),
                    detail=detail + ".",
                    fix=(
                        "Sửa danh sách members trong issue cho khớp collaborator của repo "
                        "nhóm và team.json, hoặc gửi MEMBERSHIP CHANGE REQUEST."
                    ),
                )
            )
            quarantine.update(claimed | recorded)
        if group_founder and proposal.founder != _norm_login(group_founder):
            findings.append(
                _issue(
                    "issue_founder_mismatch",
                    SEVERITY_ERROR,
                    proposal.founder or repo,
                    index,
                    field="founder",
                    found=proposal.founder,
                    detail=(
                        f"Issue #{proposal.number} ghi founder là @{proposal.founder} nhưng "
                        f"repo nhóm thuộc về @{_norm_login(group_founder)}."
                    ),
                    fix="Sửa trường 'Founder GitHub username' trong issue cho đúng người tạo repo nhóm.",
                )
            )
            quarantine.update(claimed | recorded)

    return findings, quarantine


# --------------------------------------------------------------------------
# writing back
# --------------------------------------------------------------------------


def issue_snapshot(students: Sequence[Any]) -> Dict[int, str]:
    """The content digest already recorded for each issue, keyed by issue number.

    This is what turns a rewritten proposal into a finding on the *next* run:
    without a baseline the digest has nothing to compare against.
    """
    out: Dict[int, str] = {}
    for student in students:
        number = getattr(student, FIELD_NUMBER, None)
        digest = getattr(student, FIELD_HASH, "")
        if not number or not digest:
            continue
        try:
            out[int(number)] = str(digest)
        except (TypeError, ValueError):
            continue
    return out


def recorded_groups(
    students: Sequence[Any], *, slug: str = MINI_PROJECT_SLUG
) -> Dict[str, Tuple[str, Tuple[str, ...]]]:
    """The group membership the group lane already wrote, keyed by repository.

    Reading it off the records rather than off GitHub is deliberate.  The two
    lanes are meant to disagree sometimes -- that disagreement is what
    ``issue_group_mismatch`` reports -- and re-reading the collaborators here
    would compare the board against a second fresh reading instead of against
    what the database actually holds.  It also keeps this lane to one pass over
    the network.
    """
    out: Dict[str, Tuple[str, Tuple[str, ...]]] = {}
    for student in students:
        recorded = getattr(student, FIELD_GROUP, None)
        if not isinstance(recorded, dict):
            continue
        entry = recorded.get(slug)
        if not isinstance(entry, dict):
            continue
        repo = normalize_repo_url(str(entry.get("repo", "")))
        if not repo:
            continue
        members = entry.get("members")
        if not isinstance(members, (list, tuple)):
            members = []
        out[repo] = (
            _norm_login(str(entry.get("founder", ""))),
            tuple(_norm_login(str(login)) for login in members if str(login).strip()),
        )
    return out


def merge_into_students(
    students: Sequence[Any],
    canonical: Mapping[str, CanonicalChoice],
    proposals: Sequence[ParsedProposal],
    *,
    index: Optional[Mapping[str, Any]] = None,
    roster: Any = None,
    quarantine: FrozenSet[str] = frozenset(),
) -> MergeOutcome:
    """Write the canonical topic onto every student named in it.

    Matching is by GitHub username only.  The form forbids student numbers and
    email addresses outright, so a username is the only identifier a proposal is
    allowed to carry, and the lane must not invent a second way in.

    A quarantined student keeps the group fields already on record and still
    receives the topic fields: which problem a group chose does not depend on
    who is grouped with whom, and withholding it would lose data over a dispute
    it has nothing to do with.
    """
    warnings: List[str] = []
    if index is None:
        index, warnings = _student_index(students, roster)
    by_number = {proposal.number: proposal for proposal in proposals}
    matched: List[str] = []
    unmatched: List[str] = []
    quarantined: List[str] = []
    updated = 0

    for repo, choice in sorted(canonical.items()):
        proposal = by_number.get(choice.number)
        if proposal is None:
            continue
        for login in proposal.members:
            student = index.get(_norm_login(login))
            if student is None:
                unmatched.append(login)
                continue
            matched.append(login)
            setattr(student, FIELD_NUMBER, proposal.number)
            setattr(student, FIELD_URL, proposal.url)
            setattr(student, FIELD_STATUS, "; ".join(proposal.labels))
            setattr(student, FIELD_TITLE, proposal.project_title)
            setattr(student, FIELD_TOPIC_SOURCE, proposal.topic_source)
            setattr(student, FIELD_COMMIT, proposal.commit)
            setattr(
                student,
                FIELD_FINAL,
                "yes" if proposal.comments.final_submitted else "no",
            )
            setattr(student, FIELD_HASH, proposal.content_hash)
            if _norm_login(login) in quarantine:
                quarantined.append(login)
            else:
                setattr(student, FIELD_GROUP_NAME, proposal.group_name)
                setattr(student, FIELD_REPO, repo)
                setattr(student, FIELD_MEMBERS, ", ".join(proposal.members))
            updated += 1

    return MergeOutcome(
        matched=tuple(sorted(set(matched))),
        unmatched=tuple(sorted(set(unmatched))),
        updated=updated,
        quarantined=tuple(sorted(set(quarantined))),
        warnings=tuple(warnings),
    )


# --------------------------------------------------------------------------
# the lane
# --------------------------------------------------------------------------


def import_project_issues(
    students: Sequence[Any],
    *,
    board: str,
    groups: Optional[Mapping[str, Tuple[str, Sequence[str]]]] = None,
    known_repos: FrozenSet[str] = frozenset(),
    staff_logins: FrozenSet[str] = frozenset(),
    roster: Any = None,
    previous: Optional[Mapping[int, str]] = None,
    with_proposal_md: bool = False,
    with_edit_history: bool = True,
    cli: Any = None,
    runner: Optional[Runner] = None,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> IssueReport:
    """Read the board once and record what it says, changing nothing on GitHub."""
    if "/" not in board:
        raise Classroom50Error(
            f"board must be OWNER/REPO, got {board!r}", code="bad_board"
        )
    runner = runner or timeout_runner(timeout or DEFAULT_TIMEOUT)
    expected_org = board.split("/", 1)[0]
    warnings: List[str] = []

    schema = load_form_schema(
        board, runner=runner, timeout=timeout, sleeper=sleeper, pace=pace
    )
    if schema.source == SCHEMA_PINNED:
        warnings.append(
            f"không đọc được {FORM_PATH} của {board}; dùng bảng nhãn ghim sẵn, "
            "nhãn mới đổi sẽ không khớp"
        )

    issues = fetch_issues(
        board, runner=runner, timeout=timeout, sleeper=sleeper, pace=pace
    )

    edits: Dict[int, EditRecord] = {}
    if with_edit_history:
        try:
            edits = fetch_edit_history(
                board, runner=runner, timeout=timeout, sleeper=sleeper, pace=pace
            )
        except Classroom50Error as exc:
            warnings.append(f"không đọc được lịch sử sửa issue: {exc}")

    proposals: List[ParsedProposal] = []
    skipped: List[SkippedIssue] = []
    for issue in issues:
        if not is_proposal_issue(issue, schema):
            skipped.append(
                SkippedIssue(
                    number=issue.number,
                    title=issue.title,
                    reason=(
                        "thân bài không khớp mẫu proposal "
                        f"(dưới {PROPOSAL_MIN_HEADINGS} trường của form)"
                    ),
                )
            )
            continue
        comments = fetch_comments(
            board, issue.number, runner=runner, timeout=timeout, sleeper=sleeper, pace=pace
        )
        proposal = parse_proposal(
            issue,
            schema,
            comments=comments,
            staff_logins=staff_logins,
            known_repos=known_repos,
            expected_org=expected_org,
            edit=edits.get(issue.number),
        )
        if with_proposal_md and proposal.repo and proposal.commit and not proposal.invalid:
            try:
                markers = read_proposal_markers(
                    proposal.repo,
                    proposal.commit,
                    runner=runner,
                    timeout=timeout,
                    sleeper=sleeper,
                    pace=pace,
                )
            except Classroom50Error as exc:
                warnings.append(f"issue #{issue.number}: không đọc được proposal.md: {exc}")
                markers = {}
            if markers:
                proposal = parse_proposal(
                    issue,
                    schema,
                    comments=comments,
                    staff_logins=staff_logins,
                    known_repos=known_repos,
                    expected_org=expected_org,
                    markers=markers,
                    edit=edits.get(issue.number),
                )
        proposals.append(proposal)

    canonical = choose_canonical(proposals)
    index, index_warnings = _student_index(students, roster)
    warnings.extend(index_warnings)

    baseline = dict(previous or {})
    findings: List[Issue] = []
    for proposal in proposals:
        findings.extend(
            proposal_findings(
                proposal, index, previous_hash=baseline.get(proposal.number, "")
            )
        )
    cross, quarantine = detect_member_conflicts(canonical, proposals, index, groups=groups)
    findings.extend(cross)

    outcome = merge_into_students(
        students,
        canonical,
        proposals,
        index=index,
        quarantine=frozenset(quarantine),
    )
    warnings.extend(outcome.warnings)

    return IssueReport(
        board=board,
        schema_source=schema.source,
        total=len(issues),
        proposals=tuple(proposals),
        canonical=tuple(canonical[repo] for repo in sorted(canonical)),
        skipped=tuple(skipped),
        matched=outcome.matched,
        unmatched=outcome.unmatched,
        updated=outcome.updated,
        quarantined=outcome.quarantined,
        warnings=tuple(warnings),
        findings=tuple(findings),
    )


def list_project_issues(
    board: str,
    *,
    staff_logins: FrozenSet[str] = frozenset(),
    runner: Optional[Runner] = None,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> IssueReport:
    """The board read without touching the database, for a look before an import."""
    return import_project_issues(
        [],
        board=board,
        staff_logins=staff_logins,
        with_edit_history=False,
        runner=runner,
        timeout=timeout,
        sleeper=sleeper,
        pace=pace,
    )


# --------------------------------------------------------------------------
# reporting
# --------------------------------------------------------------------------


def report_to_dict(report: IssueReport) -> Dict[str, Any]:
    """The report as JSON, in the shape the group report already uses."""
    return {
        "board": report.board,
        "schema_source": report.schema_source,
        "total_issues": report.total,
        "proposals": [
            {
                "number": proposal.number,
                "url": proposal.url,
                "title": proposal.title,
                "state": proposal.state,
                "author": proposal.author,
                "created_at": proposal.created_at,
                "labels": list(proposal.labels),
                "repo": proposal.repo,
                "founder": proposal.founder,
                "members": list(proposal.members),
                "group_name": proposal.group_name,
                "project_title": proposal.project_title,
                "topic_source": proposal.topic_source,
                "proposal_commit": proposal.commit,
                "sources": dict(proposal.sources),
                "missing": list(proposal.missing),
                "bypass": list(proposal.bypass),
                "invalid": list(proposal.invalid),
                "redactions": list(proposal.redactions),
                "content_hash": proposal.content_hash,
                "last_edited_at": proposal.last_edited_at,
                "edit_count": proposal.edit_count,
                "comment_kinds": list(proposal.comments.kinds),
                "final_submitted": proposal.comments.final_submitted,
                "membership_pending": proposal.comments.membership_pending,
            }
            for proposal in report.proposals
        ],
        "canonical": [choice._asdict() for choice in report.canonical],
        "skipped": [item._asdict() for item in report.skipped],
        "matched": list(report.matched),
        "unmatched": list(report.unmatched),
        "updated": report.updated,
        "quarantined": list(report.quarantined),
        "warnings": list(report.warnings),
        "findings": [
            {**issue._asdict(), "student": issue.student._asdict()}
            for issue in report.findings
        ],
    }


def format_project_issues(report: IssueReport) -> str:
    """The report as text for the operator."""
    lines: List[str] = [
        f"Bảng chủ đề: {report.board}",
        f"  issue đọc được: {report.total}"
        f"  |  proposal: {len(report.proposals)}"
        f"  |  bỏ qua: {len(report.skipped)}",
        f"  nhóm có issue chuẩn: {len(report.canonical)}"
        f"  |  sinh viên khớp: {len(report.matched)}"
        f"  |  không khớp: {len(report.unmatched)}",
    ]
    if report.schema_source == SCHEMA_PINNED:
        lines.append("  (dùng bảng nhãn ghim sẵn — không đọc được form trên bảng)")
    for choice in report.canonical:
        lines.append(
            f"  {choice.repo}: #{choice.number} ({choice.source})"
            + (
                f", thay thế #{', #'.join(str(n) for n in choice.superseded)}"
                if choice.superseded
                else ""
            )
        )
    invalid = [proposal for proposal in report.proposals if proposal.invalid]
    if invalid:
        lines.append(f"  proposal chưa hợp lệ: {len(invalid)}")
        for proposal in invalid:
            lines.append(f"    #{proposal.number}: {'; '.join(proposal.invalid)}")
    if report.quarantined:
        lines.append("  cách ly trường nhóm: " + ", ".join(report.quarantined))
    for warning in report.warnings:
        lines.append(f"  cảnh báo: {warning}")
    for issue in report.findings:
        lines.append(
            f"  [{issue.severity}] {issue.code}: {issue.student.name} — {issue.detail}"
        )
    return "\n".join(lines)
