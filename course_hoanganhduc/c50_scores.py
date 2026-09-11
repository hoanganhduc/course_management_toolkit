# -*- coding: utf-8 -*-
"""Read a classroom's ``scores.json`` from the Classroom50 config repo into the database.

Classroom50 already publishes what it graded, in a shape its own collector
specifies, so this module reads that document rather than reconstructing grades
from downloaded submissions.  The whole lane is one authenticated read of a file
plus one read of that file's last commit date, which puts it in the same safety
class as ``list-roster``: no download, no write, nothing dispatched.

Two facts about the document govern almost every decision here.

``entries`` are keyed by *repository owner*, not by student.  On an individual
assignment the owner is the student; on a group assignment the owner is the
founder and everybody else appears only in ``member_usernames``.  Spreading one
entry across the students it credits is this module's job -- upstream does not
do it.

The file is only as fresh as the last time a teacher pressed *Sync now*.
``collect-scores.yaml`` has no schedule, so a slug missing from the document can
mean "nobody submitted" or "nobody has collected since the assignment was
registered", and those two must never be reported as the same thing.  See
:func:`submission_state`.

Nothing here writes to GitHub, and nothing imports ``data`` or pandas, so the
whole module is testable with a stub runner and a plain interpreter.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from typing import (
    Any,
    Dict,
    FrozenSet,
    List,
    Mapping,
    NamedTuple,
    Optional,
    Sequence,
    Tuple,
)

from .c50_cli import AgentCLI, Classroom50Error, Runner, _default_runner
from .c50_roster import local_github_id, local_github_username, parse_roster_payload
from .roster_audit import (
    SEVERITY_ERROR,
    Issue,
    StudentRef,
    format_issues,
    student_ref,
)

# The repository that holds every classroom's configuration and collected
# scores.  It is private, so an unauthenticated read answers 404 rather than
# 403 -- which is why the error text below says so out loud.
CONFIG_REPO = "classroom50"

SCORES_SCHEMA = "classroom50/scores/v1"
RESULT_SCHEMA = "classroom50/result/v1"

# ``collect_scores._REQUIRED_STR_FIELDS``.  ``datetime`` is the submission
# timestamp; ``release`` is a release *name*, not a date.
REQUIRED_STR_FIELDS: Tuple[str, ...] = (
    "submission",
    "commit",
    "release",
    "review",
    "datetime",
)

FIELD_GRADES = "Classroom50 Grades"
# Every repository that paid a student for one assignment, not just the one
# whose grade won.  ``FIELD_GRADES`` is keyed by slug alone, so a student on two
# group repositories keeps only the later submission and the other grade is gone
# with no record that it existed; this field is that record.
FIELD_GRADE_CANDIDATES = "Classroom50 Grade Candidates"
FIELD_SUBMISSIONS = "Classroom50 Submissions"
FIELD_DETAILS = "Classroom50 Submission Details"
FIELD_OVERRIDES = "Classroom50 Score Overrides"
FIELD_COLLECTED_AT = "Classroom50 Scores Collected At"

STATE_SUBMITTED = "submitted"
STATE_NOT_SUBMITTED = "not submitted"
STATE_NOT_COLLECTED = "not collected yet"
STATE_UNKNOWN_STALE = "unknown (stale)"

TYPE_INDIVIDUAL = "individual"
TYPE_GROUP = "group"

# The three staff roles ``gh teacher`` recognises, taken from the pinned
# v1.45.0 binary itself: it rejects anything else with
# ``invalid --role %q: must be "teacher", "hta", or "ta"``.  The role may also
# be empty, which the same help text spells out, and an empty role is a
# student here -- dropping a student from the credit list loses their grade,
# while keeping a teacher in it only produces one unmatched login.
STAFF_ROLES: FrozenSet[str] = frozenset({"teacher", "hta", "ta"})


class ScoreEntry(NamedTuple):
    """One repository's collected submissions, keyed the way upstream keys them."""

    owner: str
    member_usernames: List[str]
    submissions: List[Dict[str, Any]]
    override: bool


class AssignmentScores(NamedTuple):
    slug: str
    type: str
    entries: List[ScoreEntry]


class ScoresDocument(NamedTuple):
    """A parsed ``scores.json``; ``warnings`` records what was skipped and why."""

    schema: str
    assignments: Dict[str, AssignmentScores]
    warnings: List[str]


class AssignmentInfo(NamedTuple):
    """The bits of an assignment manifest row that decide a missing slug's meaning."""

    slug: str
    mode: str
    available_from: Optional[str]
    due: Optional[str]
    graded: bool


class StudentScore(NamedTuple):
    """One assignment's result as it applies to one credited student."""

    username: str
    slug: str
    grade: Optional[int]
    max_points: Optional[int]
    commit: str
    release: str
    datetime: str
    late: Optional[bool]
    group: bool
    override: bool
    owner: str
    submissions: List[Dict[str, Any]]


class MergeReport(NamedTuple):
    """What the merge did, in the terms a teacher would ask about."""

    collected_at: Optional[str]
    matched: List[str]
    unmatched: List[str]
    updated: int
    stale_skipped: List[str]
    group_slugs: List[str]
    credited: Dict[str, Dict[str, List[str]]]
    overrides: List[str]
    warnings: List[str]
    findings: List[Issue]


def _norm_login(value: Any) -> str:
    """A GitHub login as a comparison key: stripped, unprefixed, lower-cased."""
    text = str(value or "").strip()
    if text.startswith("@"):
        text = text[1:]
    return text.lower()


def _as_int(value: Any) -> Optional[int]:
    """``value`` when it is a genuine integer, else ``None``.

    ``bool`` is an ``int`` subclass and is rejected on purpose: upstream writes
    ``score`` as an integer, so a ``True`` here means the payload is wrong and
    should be reported, not silently graded as 1.
    """
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    return None


def _run(runner: Runner, argv: List[str], *, op: str) -> str:
    """Run one ``gh`` command and return stdout, or raise with the reason.

    The exit code is checked before the output is looked at.  ``gh`` prints an
    error object into the same stream it prints data into, so output that parses
    is not evidence that the call succeeded.
    """
    try:
        result = runner(argv)
    except Classroom50Error:
        raise
    except OSError as exc:
        raise Classroom50Error(f"{op} failed: {exc}", code=f"{op}_failed") from exc
    if getattr(result, "returncode", 1) != 0:
        stderr = (getattr(result, "stderr", "") or "").strip()
        raise Classroom50Error(
            f"{op} failed: {stderr or 'gh exited non-zero'}", code=f"{op}_failed"
        )
    return getattr(result, "stdout", "") or ""


def _parse_json(text: str, *, op: str, expect: type) -> Any:
    value = None
    try:
        value = json.loads(text)
    except ValueError as exc:
        raise Classroom50Error(f"{op}: invalid JSON: {exc}", code="bad_json") from exc
    if not isinstance(value, expect):
        raise Classroom50Error(
            f"{op}: expected {expect.__name__}, got {type(value).__name__}: "
            f"{str(value)[:200]}",
            code="bad_json",
        )
    return value


def fetch_scores(
    org: str,
    classroom: str,
    *,
    runner: Optional[Runner] = None,
) -> ScoresDocument:
    """Read ``<classroom>/scores.json`` out of the config repo."""
    runner = runner or _default_runner
    argv = [
        "gh",
        "api",
        "-H",
        "Accept: application/vnd.github.raw",
        f"repos/{org}/{CONFIG_REPO}/contents/{classroom}/scores.json",
    ]
    try:
        text = _run(runner, argv, op="fetch_scores")
    except Classroom50Error as exc:
        raise Classroom50Error(
            f"{exc}  (the {org}/{CONFIG_REPO} repository is private: an "
            f"unauthenticated read answers 404, not 403 -- check `gh auth status` "
            f"before concluding the classroom does not exist)",
            code=exc.code,
        ) from exc
    return parse_scores(_parse_json(text, op="fetch_scores", expect=dict))


def fetch_scores_collected_at(
    org: str,
    classroom: str,
    *,
    runner: Optional[Runner] = None,
) -> Optional[str]:
    """When the server last committed this classroom's ``scores.json``.

    This is the only honest answer to "how fresh are these grades".  The local
    clock says when the import ran, which is a different question.
    """
    runner = runner or _default_runner
    argv = [
        "gh",
        "api",
        f"repos/{org}/{CONFIG_REPO}/commits?path={classroom}/scores.json&per_page=1",
    ]
    text = _run(runner, argv, op="fetch_scores_collected_at")
    commits = _parse_json(text, op="fetch_scores_collected_at", expect=list)
    if not commits or not isinstance(commits[0], dict):
        return None
    commit = commits[0].get("commit")
    if not isinstance(commit, dict):
        return None
    committer = commit.get("committer")
    if not isinstance(committer, dict):
        return None
    date = str(committer.get("date") or "").strip()
    return date or None


def parse_scores(document: Any) -> ScoresDocument:
    """Validate the document's shape and normalize its entries.

    An unrecognised ``schema`` stops the read and quotes what was found: the
    field exists precisely so a consumer can refuse a document it was not
    written against, and guessing past it would write grades from a format
    nobody has checked.
    """
    if not isinstance(document, dict):
        raise Classroom50Error(
            f"scores.json is not an object: {type(document).__name__}",
            code="bad_scores_schema",
        )
    schema = str(document.get("schema") or "").strip()
    if schema and schema != SCORES_SCHEMA:
        raise Classroom50Error(
            f"unsupported scores schema {schema!r}; expected {SCORES_SCHEMA!r}",
            code="bad_scores_schema",
        )
    if "assignments" not in document:
        # Every document the collector writes carries this key, empty or not, so
        # a dict without it is something else that happens to be JSON -- most
        # often gh's own error object, which is a dict and so survives the type
        # check above.  Quote the keys rather than reporting an empty gradebook.
        raise Classroom50Error(
            "scores.json has no 'assignments' key; read "
            f"{sorted(str(k) for k in document)}",
            code="bad_scores_schema",
        )
    raw_assignments = document.get("assignments")
    if not isinstance(raw_assignments, dict):
        raise Classroom50Error(
            f"scores.json 'assignments' is not an object: "
            f"{type(raw_assignments).__name__}",
            code="bad_scores_schema",
        )

    warnings: List[str] = []
    assignments: Dict[str, AssignmentScores] = {}
    for slug, bucket in raw_assignments.items():
        slug_key = str(slug).strip()
        if not slug_key:
            warnings.append("assignment with an empty slug skipped")
            continue
        if not isinstance(bucket, dict):
            warnings.append(f"{slug_key}: bucket is not an object, skipped")
            continue
        kind = str(bucket.get("type") or TYPE_INDIVIDUAL).strip().lower()
        if kind not in (TYPE_INDIVIDUAL, TYPE_GROUP):
            warnings.append(f"{slug_key}: unknown assignment type {kind!r}, read as individual")
            kind = TYPE_INDIVIDUAL
        raw_entries = bucket.get("entries")
        entries: List[ScoreEntry] = []
        if raw_entries is None:
            raw_entries = []
        if not isinstance(raw_entries, list):
            warnings.append(f"{slug_key}: 'entries' is not a list, skipped")
            raw_entries = []
        for index, raw in enumerate(raw_entries):
            entry = _parse_entry(raw, slug=slug_key, index=index, warnings=warnings)
            if entry is not None:
                entries.append(entry)
        assignments[slug_key] = AssignmentScores(slug=slug_key, type=kind, entries=entries)
    return ScoresDocument(
        schema=schema or SCORES_SCHEMA, assignments=assignments, warnings=warnings
    )


def _parse_entry(
    raw: Any, *, slug: str, index: int, warnings: List[str]
) -> Optional[ScoreEntry]:
    """One ``entries`` element, or ``None`` when it cannot be attributed.

    An entry with no ``owner`` cannot be attached to anybody, so it is dropped
    with a warning rather than failing the import: one malformed record must not
    cost the class its other grades.
    """
    if not isinstance(raw, dict):
        warnings.append(f"{slug}: entry {index} is not an object, skipped")
        return None
    owner = _norm_login(raw.get("owner"))
    if not owner:
        warnings.append(f"{slug}: entry {index} has no owner, skipped")
        return None
    members: List[str] = []
    seen = set()
    raw_members = raw.get("member_usernames")
    if isinstance(raw_members, list):
        for value in raw_members:
            login = _norm_login(value)
            if login and login not in seen:
                seen.add(login)
                members.append(login)
    elif raw_members is not None:
        warnings.append(f"{slug}/{owner}: 'member_usernames' is not a list, ignored")
    submissions: List[Dict[str, Any]] = []
    raw_submissions = raw.get("submissions")
    if isinstance(raw_submissions, list):
        for payload in raw_submissions:
            if isinstance(payload, dict):
                submissions.append(payload)
            else:
                warnings.append(f"{slug}/{owner}: non-object submission skipped")
    elif raw_submissions is not None:
        warnings.append(f"{slug}/{owner}: 'submissions' is not a list, ignored")
    return ScoreEntry(
        owner=owner,
        member_usernames=members,
        submissions=submissions,
        override=bool(raw.get("override")),
    )


def _payload_warnings(payload: Mapping[str, Any], *, slug: str, owner: str) -> List[str]:
    """What is wrong with one collected payload, in the collector's own terms."""
    found: List[str] = []
    schema = str(payload.get("schema") or "").strip()
    if schema and schema != RESULT_SCHEMA:
        found.append(f"{slug}/{owner}: result schema {schema!r}, expected {RESULT_SCHEMA!r}")
    for field in REQUIRED_STR_FIELDS:
        value = payload.get(field)
        if not isinstance(value, str) or not value.strip():
            found.append(f"{slug}/{owner}: required field {field!r} is empty or not a string")
    score = _as_int(payload.get("score"))
    maximum = _as_int(payload.get("max-score"))
    if payload.get("score") is not None and score is None:
        found.append(f"{slug}/{owner}: 'score' is not an integer")
    if payload.get("max-score") is not None and maximum is None:
        found.append(f"{slug}/{owner}: 'max-score' is not an integer")
    if score is not None and maximum is not None and score > maximum:
        found.append(f"{slug}/{owner}: score {score} exceeds max-score {maximum}")
    return found


def credited_members(
    entry: ScoreEntry,
    *,
    kind: str,
    staff_logins: FrozenSet[str] = frozenset(),
) -> List[str]:
    """The students one entry pays out to.

    On a group assignment ``member_usernames`` is what Classroom50 actually
    credited; the staff logins are subtracted because the collector's roster is
    the union of the student team and every staff team, so a teacher who is a
    collaborator on the group repository is credited upstream by design.  That
    is upstream behaviour working correctly, not a fault to report.
    """
    if kind != TYPE_GROUP:
        return [entry.owner]
    members = [login for login in entry.member_usernames if login not in staff_logins]
    if not members and entry.owner not in staff_logins:
        return [entry.owner]
    return members


def expand_entries(
    doc: ScoresDocument,
    *,
    staff_logins: FrozenSet[str] = frozenset(),
) -> List[StudentScore]:
    """Spread every entry across the students it credits, newest submission first."""
    staff = frozenset(_norm_login(login) for login in staff_logins if _norm_login(login))
    scores: List[StudentScore] = []
    for slug, bucket in sorted(doc.assignments.items()):
        is_group = bucket.type == TYPE_GROUP
        for entry in bucket.entries:
            newest = entry.submissions[0] if entry.submissions else None
            if newest is not None:
                doc.warnings.extend(
                    _payload_warnings(newest, slug=slug, owner=entry.owner)
                )
            for username in credited_members(entry, kind=bucket.type, staff_logins=staff):
                scores.append(
                    StudentScore(
                        username=username,
                        slug=slug,
                        grade=_as_int((newest or {}).get("score")),
                        max_points=_as_int((newest or {}).get("max-score")),
                        commit=str((newest or {}).get("commit") or ""),
                        release=str((newest or {}).get("release") or ""),
                        datetime=str((newest or {}).get("datetime") or ""),
                        late=(newest or {}).get("late")
                        if isinstance((newest or {}).get("late"), bool)
                        else None,
                        group=is_group,
                        override=entry.override,
                        owner=entry.owner,
                        submissions=list(entry.submissions),
                    )
                )
    return scores


def credit_snapshot(
    doc: ScoresDocument,
    *,
    staff_logins: FrozenSet[str] = frozenset(),
) -> Dict[str, Dict[str, List[str]]]:
    """Who each group entry credited, keyed ``slug -> owner -> members``.

    Upstream keeps no history: a group entry whose member set changes between
    collects replaces its predecessor, so a student dropped from the set loses
    the shared grade with no record that they ever had it.  Saving this snapshot
    each run is how that becomes visible.
    """
    staff = frozenset(_norm_login(login) for login in staff_logins if _norm_login(login))
    snapshot: Dict[str, Dict[str, List[str]]] = {}
    for slug, bucket in doc.assignments.items():
        if bucket.type != TYPE_GROUP:
            continue
        per_owner: Dict[str, List[str]] = {}
        for entry in bucket.entries:
            per_owner[entry.owner] = credited_members(
                entry, kind=bucket.type, staff_logins=staff
            )
        if per_owner:
            snapshot[slug] = per_owner
    return snapshot


def previous_credit_snapshot(students: Sequence[Any]) -> Dict[str, Dict[str, List[str]]]:
    """Rebuild the last run's ``slug -> owner -> members`` from the database.

    The grades written by :func:`merge_into_students` carry the owner of the
    repository that paid out, which is enough to reconstruct who each group
    entry credited without keeping a file beside the database.

    One limit worth stating: a student is listed here under the login their
    record carries *now*.  If the founder renamed their GitHub account between
    two runs, the old and new owners do not line up and the whole group reads as
    dropped -- which errs towards reporting rather than towards silence.
    """
    snapshot: Dict[str, Dict[str, List[str]]] = {}
    for student in students:
        login = _norm_login(local_github_username(student) or "")
        if not login:
            continue
        grades = getattr(student, FIELD_GRADES, None)
        if not isinstance(grades, dict):
            continue
        for slug, record in grades.items():
            if not isinstance(record, dict) or not record.get("group"):
                continue
            owner = _norm_login(record.get("owner") or "")
            if not owner:
                continue
            members = snapshot.setdefault(str(slug), {}).setdefault(owner, [])
            if login not in members:
                members.append(login)
    for per_owner in snapshot.values():
        for members in per_owner.values():
            members.sort()
    return snapshot


def _ref_for(login: str, index: Mapping[str, Any]) -> StudentRef:
    """A ``StudentRef`` for a GitHub login, named by the database when it knows it."""
    student = index.get(login)
    if student is not None:
        return student_ref(student)
    return StudentRef(student_id="", name=login, email="", google_id="")


def credit_findings(
    doc: ScoresDocument,
    students: Sequence[Any],
    *,
    staff_logins: FrozenSet[str] = frozenset(),
    expected_members: Optional[Mapping[str, Mapping[str, Sequence[str]]]] = None,
    previous: Optional[Mapping[str, Mapping[str, Sequence[str]]]] = None,
) -> List[Issue]:
    """The two ways a student can be silently missing from a group grade.

    ``expected_members`` is who the group says it is -- collaborators or
    ``team.json`` -- and ``previous`` is :func:`credit_snapshot` from the last
    run.  Both are optional: on a first run there is nothing to compare against,
    and ``group_member_dropped`` correctly does not fire.
    """
    staff = frozenset(_norm_login(login) for login in staff_logins if _norm_login(login))
    index = {
        login: student
        for login, student in (
            (_norm_login(local_github_username(s) or ""), s) for s in students
        )
        if login
    }
    current = credit_snapshot(doc, staff_logins=staff)
    findings: List[Issue] = []

    for slug, per_owner in sorted(current.items()):
        for owner, members in sorted(per_owner.items()):
            credited = set(members)
            expected = {
                _norm_login(name)
                for name in ((expected_members or {}).get(slug, {}) or {}).get(owner, [])
                if _norm_login(name)
            } - staff
            uncredited = sorted(expected - credited)
            if uncredited and credited <= {owner}:
                for login in uncredited:
                    findings.append(
                        Issue(
                            code="group_credit_owner_only",
                            severity=SEVERITY_ERROR,
                            student=_ref_for(login, index),
                            field="GitHub Username",
                            found=login,
                            detail=(
                                f"bài {slug}: Classroom50 chỉ tính điểm cho {owner}; "
                                f"bạn không có tên trong danh sách được tính điểm của "
                                f"kho nhóm {owner}"
                            ),
                            fix=(
                                f"nhờ {owner} thêm bạn làm collaborator của kho nhóm, "
                                f"rồi báo giảng viên thu lại điểm"
                            ),
                        )
                    )
            before = {
                _norm_login(name)
                for name in ((previous or {}).get(slug, {}) or {}).get(owner, [])
                if _norm_login(name)
            } - staff
            for login in sorted(before - credited):
                findings.append(
                    Issue(
                        code="group_member_dropped",
                        severity=SEVERITY_ERROR,
                        student=_ref_for(login, index),
                        field="GitHub Username",
                        found=login,
                        detail=(
                            f"bài {slug}: lần thu trước bạn được tính điểm cùng kho nhóm "
                            f"{owner}, lần thu này thì không -- điểm chung đã mất"
                        ),
                        fix=(
                            f"nhờ {owner} thêm lại bạn làm collaborator của kho nhóm, "
                            f"rồi báo giảng viên thu lại điểm"
                        ),
                    )
                )
    return findings


def _points(record: Mapping[str, Any]) -> str:
    """One stored grade record as ``85/100``, for a message a student reads."""
    grade = record.get("grade")
    if grade is None:
        return "chưa chấm"
    maximum = record.get("max_points")
    return f"{grade}/{maximum}" if maximum is not None else f"{grade}"


def _parse_iso(value: Any) -> Optional[datetime]:
    """An ISO-8601 timestamp as an aware ``datetime``, or ``None``.

    Everything the collector writes carries a zone -- the manifest stores UTC
    and keeps the offset beside it -- so a naive value is read as UTC rather
    than as the machine's local time, which would shift every comparison by
    seven hours on this course's clock.
    """
    text = str(value or "").strip()
    if not text:
        return None
    if text.endswith(("Z", "z")):
        text = f"{text[:-1]}+00:00"
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=timezone.utc)
    return parsed


def _not_older(candidate: Any, existing: Any) -> bool:
    """Whether ``candidate`` is at least as recent as ``existing``.

    Unparseable or absent values answer ``True``: refusing to write because a
    timestamp could not be read would quietly drop real grades.
    """
    new = _parse_iso(candidate)
    old = _parse_iso(existing)
    if new is None or old is None:
        return True
    return new >= old


def submission_state(
    *,
    submitted: bool,
    collected_at: Optional[str],
    available_from: Optional[str] = None,
    due: Optional[str] = None,
) -> str:
    """What can honestly be said about one student and one assignment.

    A snapshot taken before the deadline cannot tell "has not submitted" from
    "submitted after the last collection", so it never claims the first.  A
    snapshot taken before the assignment even opened says only that nobody has
    collected yet.
    """
    if submitted:
        return STATE_SUBMITTED
    collected = _parse_iso(collected_at)
    if collected is None:
        return STATE_UNKNOWN_STALE
    opened = _parse_iso(available_from)
    if opened is not None and collected < opened:
        return STATE_NOT_COLLECTED
    deadline = _parse_iso(due)
    if deadline is None or collected < deadline:
        return STATE_UNKNOWN_STALE
    return STATE_NOT_SUBMITTED


def parse_assignments(payload: Any) -> Dict[str, AssignmentInfo]:
    """The manifest rows this module needs, from ``assignments.json`` or the CLI."""
    rows: Sequence[Any]
    if isinstance(payload, dict):
        candidate = payload.get("assignments")
        rows = candidate if isinstance(candidate, list) else []
    elif isinstance(payload, list):
        rows = payload
    else:
        rows = []
    out: Dict[str, AssignmentInfo] = {}
    for row in rows:
        if not isinstance(row, dict):
            continue
        slug = str(row.get("slug") or row.get("name") or "").strip()
        if not slug:
            continue
        grading = row.get("grading")
        mode = str(row.get("mode") or TYPE_INDIVIDUAL).strip().lower()
        graded = not bool(row.get("empty_repo"))
        if isinstance(grading, dict) and str(grading.get("mode") or "").strip().lower() == "off":
            graded = False
        out[slug] = AssignmentInfo(
            slug=slug,
            mode=mode,
            available_from=str(row.get("available_from") or "").strip() or None,
            due=str(row.get("due") or "").strip() or None,
            graded=graded,
        )
    return out


def _raw_roster_rows(roster: Any) -> List[Dict[str, Any]]:
    """The roster rows as the CLI emitted them, before :func:`parse_roster_payload`.

    That parser drops ``role`` on purpose -- re-emitting a row through its
    dialect would erase the role upstream -- so anything that needs the role has
    to read the payload itself.  The unwrapping mirrors the parser's so the two
    always see the same rows.
    """
    if isinstance(roster, list):
        rows: Sequence[Any] = roster
    elif isinstance(roster, dict):
        for key in ("students", "roster", "entries", "data"):
            value = roster.get(key)
            if isinstance(value, list):
                rows = value
                break
        else:
            rows = [roster]
    else:
        return []
    return [row for row in rows if isinstance(row, dict)]


def staff_logins(roster: Any) -> FrozenSet[str]:
    """The teacher, HTA and TA logins in a ``roster list --json`` payload.

    These are subtracted from a group entry's credited members.  The collector's
    roster is the union of the student team and every staff team, so a teacher
    who is a collaborator on a group repository is credited upstream by design;
    that is not a fault, but it is not a grade either.
    """
    out = set()
    for row in _raw_roster_rows(roster):
        role = str(row.get("role") or "").strip().lower()
        if role not in STAFF_ROLES:
            continue
        login = _norm_login(
            row.get("login")
            or row.get("username")
            or row.get("github_username")
            or ""
        )
        if login:
            out.add(login)
    return frozenset(out)


def _roster_ids(roster: Any) -> Dict[str, str]:
    """``login -> github_id`` from a Classroom50 roster listing."""
    out: Dict[str, str] = {}
    for row in parse_roster_payload(roster):
        login = _norm_login(row.get("username"))
        gid = str(row.get("github_id") or "").strip()
        if login and gid and gid not in ("0", "None"):
            out[login] = gid
    return out


def _student_index(
    students: Sequence[Any], roster: Any
) -> Tuple[Dict[str, Any], List[str]]:
    """Map every GitHub login to its database record, by numeric id where possible.

    A GitHub username can be changed by its owner; the numeric id cannot, so the
    id is the join key whenever both sides know it.  ``scores.json`` carries only
    usernames, which is why the roster is consulted to turn one into the other.
    """
    warnings: List[str] = []
    by_id: Dict[str, Any] = {}
    by_login: Dict[str, Any] = {}
    for student in students:
        gid = local_github_id(student)
        if gid:
            by_id[gid] = student
        login = _norm_login(local_github_username(student) or "")
        if login:
            by_login.setdefault(login, student)
    ids = _roster_ids(roster)
    index: Dict[str, Any] = dict(by_login)
    for login, gid in ids.items():
        student = by_id.get(gid)
        if student is not None:
            index[login] = student
    if roster is not None and not ids:
        warnings.append(
            "roster carries no github_id; matching by username only "
            "(a student who renamed their account will not be found)"
        )
    return index, warnings


def merge_into_students(
    students: Sequence[Any],
    scores: Sequence[StudentScore],
    *,
    collected_at: Optional[str] = None,
    roster: Any = None,
    assignments: Optional[Mapping[str, AssignmentInfo]] = None,
    warnings: Optional[Sequence[str]] = None,
) -> MergeReport:
    """Write the collected grades onto the records that already exist.

    No student is created: a login with no record is reported as unmatched,
    because inventing a row from a grade would put somebody in the class list
    who was never registered there.
    """
    index, notes = _student_index(students, roster)
    report_warnings: List[str] = list(warnings or []) + notes
    matched: set = set()
    unmatched: set = set()
    stale_skipped: List[str] = []
    overrides: List[str] = []
    group_slugs: set = set()
    # ``refreshed`` is the ``(record, slug)`` pairs this run has already rebuilt,
    # so the first entry replaces the stored candidates and the rest add to them.
    refreshed: set = set()
    contested: Dict[Tuple[str, str], Dict[str, Dict[str, Any]]] = {}
    updated = 0

    for score in scores:
        student = index.get(score.username)
        if student is None:
            unmatched.add(score.username)
            continue
        matched.add(score.username)
        if score.group:
            group_slugs.add(score.slug)
        if not _not_older(collected_at, getattr(student, FIELD_COLLECTED_AT, None)):
            stale_skipped.append(f"{score.slug}/{score.username}: older snapshot")
            continue
        record = {
            "grade": score.grade,
            "max_points": score.max_points,
            "commit": score.commit,
            "release": score.release,
            "datetime": score.datetime,
            "late": score.late,
            "group": score.group,
            # Which repository this grade came from.  Upstream keeps no history
            # of who a group entry credited, so this is what lets the next run
            # notice that somebody was dropped from the group.
            "owner": score.owner,
        }

        # Recorded before the staleness check below, because the entry that
        # check discards is exactly the one that would otherwise vanish: two
        # group repositories crediting the same student share one ``slug`` key,
        # so the later submission silently replaces the earlier grade.
        #
        # The slug's candidates are rebuilt from the snapshot in hand rather
        # than added to, so a repository that has stopped crediting the student
        # stops being reported instead of accusing them forever.
        #
        # Keyed by the record rather than by the login, because the field lives
        # on the record: two logins resolving to one student must not each
        # rebuild the same slug and wipe what the other just wrote.
        seen = (id(student), score.slug)
        candidates = dict(getattr(student, FIELD_GRADE_CANDIDATES, None) or {})
        paying = dict(candidates.get(score.slug) or {}) if seen in refreshed else {}
        paying[score.owner] = record
        candidates[score.slug] = paying
        setattr(student, FIELD_GRADE_CANDIDATES, candidates)
        refreshed.add(seen)
        contested[(score.username, score.slug)] = paying

        grades = dict(getattr(student, FIELD_GRADES, None) or {})
        previous = grades.get(score.slug)
        if isinstance(previous, dict) and not _not_older(
            score.datetime, previous.get("datetime")
        ):
            stale_skipped.append(f"{score.slug}/{score.username}: older submission")
            continue
        grades[score.slug] = dict(record)
        setattr(student, FIELD_GRADES, grades)

        details = dict(getattr(student, FIELD_DETAILS, None) or {})
        details[score.slug] = list(score.submissions)
        setattr(student, FIELD_DETAILS, details)

        submissions = dict(getattr(student, FIELD_SUBMISSIONS, None) or {})
        submissions[score.slug] = submission_state(
            submitted=bool(score.submissions), collected_at=collected_at
        )
        setattr(student, FIELD_SUBMISSIONS, submissions)

        if score.override:
            stored = dict(getattr(student, FIELD_OVERRIDES, None) or {})
            stored[score.slug] = True
            setattr(student, FIELD_OVERRIDES, stored)
            overrides.append(f"{score.slug}/{score.username}")
        updated += 1

    # Two repositories paying the same student for the same assignment is the
    # one conflict this lane resolves silently -- the later submission wins by
    # the same rule that keeps a re-collect from walking a grade back, which was
    # never meant to arbitrate between two repositories.  The rule stands; what
    # changes here is that it stops being silent.
    findings: List[Issue] = []
    for (login, slug), paying in sorted(contested.items()):
        if len(paying) < 2:
            continue
        in_effect = ""
        stored = (getattr(index[login], FIELD_GRADES, None) or {}).get(slug)
        if isinstance(stored, dict):
            in_effect = _norm_login(stored.get("owner") or "")
        shown = ", ".join(
            f"{owner}: {_points(paying[owner])}"
            + (" (đang dùng)" if owner == in_effect else "")
            for owner in sorted(paying)
        )
        findings.append(
            Issue(
                code="grade_multi_repo",
                severity=SEVERITY_ERROR,
                student=_ref_for(login, index),
                field="GitHub Username",
                found=login,
                detail=(
                    f"bài {slug}: {len(paying)} kho nhóm cùng tính điểm cho bạn "
                    f"-- {shown}; chỉ một điểm được ghi vào sổ"
                ),
                fix=(
                    "Mỗi bạn chỉ ở một nhóm cho mỗi bài. Rời khỏi kho của nhóm "
                    "không phải nhóm của mình, rồi báo giảng viên thu lại điểm."
                ),
            )
        )

    # A student with no entry for an assignment has not necessarily failed to
    # submit; what can be said depends on when the snapshot was taken relative to
    # the assignment's own dates, which is why the manifest is consulted here
    # rather than a blanket "not submitted" being written.
    for slug, info in sorted((assignments or {}).items()):
        if not info.graded:
            continue
        state = submission_state(
            submitted=False,
            collected_at=collected_at,
            available_from=info.available_from,
            due=info.due,
        )
        for student in index.values():
            submissions = dict(getattr(student, FIELD_SUBMISSIONS, None) or {})
            if slug in submissions:
                continue
            submissions[slug] = state
            setattr(student, FIELD_SUBMISSIONS, submissions)

    if collected_at:
        for student in index.values():
            if _not_older(collected_at, getattr(student, FIELD_COLLECTED_AT, None)):
                setattr(student, FIELD_COLLECTED_AT, collected_at)

    return MergeReport(
        collected_at=collected_at,
        matched=sorted(matched),
        unmatched=sorted(unmatched),
        updated=updated,
        stale_skipped=stale_skipped,
        group_slugs=sorted(group_slugs),
        credited={},
        overrides=sorted(overrides),
        warnings=report_warnings,
        findings=findings,
    )


def report_to_dict(report: MergeReport) -> Dict[str, Any]:
    """The merge report as JSON, matching ``roster_audit.report_to_dict``'s shape."""
    return {
        "collected_at": report.collected_at,
        "matched": list(report.matched),
        "unmatched": list(report.unmatched),
        "updated": report.updated,
        "stale_skipped": list(report.stale_skipped),
        "group_slugs": list(report.group_slugs),
        "credited": {
            slug: {owner: list(members) for owner, members in per_owner.items()}
            for slug, per_owner in report.credited.items()
        },
        "overrides": list(report.overrides),
        "warnings": list(report.warnings),
        # Same shape as ``roster_audit.report_to_dict`` writes its findings, so
        # one reader handles both reports.
        "findings": [
            {**issue._asdict(), "student": issue.student._asdict()}
            for issue in report.findings
        ],
    }


def format_merge(report: MergeReport) -> str:
    """A short human summary, in the shape the other lanes print."""
    lines = [
        f"Classroom50 scores: matched={len(report.matched)} "
        f"unmatched={len(report.unmatched)} updated={report.updated} "
        f"stale_skipped={len(report.stale_skipped)}"
    ]
    lines.append(f"  collected at: {report.collected_at or '(unknown)'}")
    if report.group_slugs:
        lines.append(f"  group assignments: {', '.join(report.group_slugs)}")
    if report.overrides:
        lines.append(f"  teacher overrides upstream: {', '.join(report.overrides)}")
    if report.unmatched:
        lines.append(
            f"  no database record for: {', '.join(report.unmatched)}"
        )
    for warning in report.warnings:
        lines.append(f"  warning: {warning}")
    if report.findings:
        lines.append("")
        lines.append(format_issues(report.findings))
    return "\n".join(lines)


def import_scores(
    students: Sequence[Any],
    *,
    org: str,
    classroom: str,
    assignment: Optional[str] = None,
    cli: Optional[AgentCLI] = None,
    runner: Optional[Runner] = None,
) -> MergeReport:
    """Read one classroom's collected scores and write them onto the records.

    Nothing is saved here: the caller decides whether the run was a dry one.
    The order matters -- the previous credit snapshot is taken from the database
    *before* the merge overwrites it, because that comparison is the only way a
    student dropped from a group repository becomes visible.
    """
    cli = cli or AgentCLI(runner=runner)

    roster = cli.list_roster(org, classroom)
    staff = staff_logins(roster)
    manifest = parse_assignments(cli.list_assignments(org, classroom))

    doc = fetch_scores(org, classroom, runner=runner)
    collected_at = fetch_scores_collected_at(org, classroom, runner=runner)

    if assignment:
        slug = str(assignment).strip()
        info = manifest.get(slug)
        if info is None:
            raise Classroom50Error(
                f"unknown assignment {slug!r}; this classroom has: "
                f"{', '.join(sorted(manifest)) or '(none)'}",
                code="unknown_assignment",
            )
        manifest = {slug: info}
        doc = doc._replace(
            assignments={
                key: bucket for key, bucket in doc.assignments.items() if key == slug
            }
        )
        if not info.graded:
            doc.warnings.append(
                f"{slug}: no automatic grade -- grading is off and the repository is "
                f"empty upstream, so this assignment never appears in scores.json; "
                f"enter these marks with --load-override-grades"
            )

    previous = previous_credit_snapshot(students)
    scores = expand_entries(doc, staff_logins=staff)
    findings = credit_findings(
        doc, students, staff_logins=staff, previous=previous
    )
    report = merge_into_students(
        students,
        scores,
        collected_at=collected_at,
        roster=roster,
        assignments=manifest,
        warnings=doc.warnings,
    )
    # The merge reports the conflicts it had to resolve; ``credit_findings``
    # reports who the document never paid at all.  Neither subsumes the other.
    return report._replace(
        credited=credit_snapshot(doc, staff_logins=staff),
        findings=list(findings) + list(report.findings),
    )
