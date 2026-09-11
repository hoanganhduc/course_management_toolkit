# -*- coding: utf-8 -*-
"""Read Classroom50 group membership for the ``mode=group`` assignments.

Both group assignments in these classrooms are ``mode=group``, which the pinned
``gh teacher`` binary describes as a "legacy shared repo through collaborators".
There is no GitHub Team behind them, so the authority on who is in a group is
the repository's collaborator list.  The upstream CLI refuses its team
subcommand on such an assignment, so that subcommand is deliberately not
wrapped here.

The module reproduces the collector's crediting rule exactly, in two separate
steps, because the two answer different questions:

* :func:`credit_members` answers "who does Classroom50 credit", and must match
  ``collect_scores.group_member_usernames`` line for line.  A list that differs
  from it describes a group whose grades went somewhere else.
* :func:`student_members` answers "who is a student in this group", and is ours:
  the collector's roster is the union of the student team and every staff team,
  so a teacher who is a collaborator is credited upstream by design.

Everything here is read-only and stdlib-only: no repository is modified, no
collaborator is added or removed, no ``team.json`` is rewritten, and the module
imports neither :mod:`course_hoanganhduc.data` nor pandas, so its tests run
under a bare interpreter.
"""

from __future__ import annotations

import json
import re
import subprocess
import time
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

from .c50_cli import AgentCLI, Classroom50Error, Runner, RunResult, _default_runner
from .c50_roster import parse_roster_payload
from .c50_scores import (
    CONFIG_REPO,
    TYPE_GROUP,
    AssignmentInfo,
    _norm_login,
    _parse_json,
    _run,
    _student_index,
    parse_assignments,
    staff_logins,
)
from .roster_audit import SEVERITY_ERROR, SEVERITY_WARNING, Issue, StudentRef, student_ref

# Database fields.  Every one of them starts with a prefix the merge step
# protects, so a second roster import cannot quietly overwrite them.
FIELD_GROUP = "Classroom50 Group"
FIELD_GROUP_CONFLICTS = "Classroom50 Group Conflicts"
FIELD_MP_REPO = "MiniProject_Group_Repo"
FIELD_MP_MEMBERS = "MiniProject_Group_Members_GitHub"
FIELD_MP_SOURCE = "MiniProject_Group_Source"
FIELD_MP_CONFLICT = "MiniProject_Group_Conflict"

# Where a group's member list came from.  ``final-project`` has autograding off
# and an empty repository, so the collector skips it entirely and it is never
# present in ``scores.json``; its members can only be computed.
SOURCE_COLLABORATORS = "c50-collaborators"
SOURCE_SCORES = "scores.json"

MINI_PROJECT_SLUG = "final-project"

# ``check_project_files.py``, pinned at blob
# 3af80db0206215cfd43314b5feba8060c632e417.  These are that file's constants,
# copied rather than reinvented: a group whose ``team.json`` passes the staff
# checker but fails ours would be told to fix a file that is already correct.
TEAM_KEYS: FrozenSet[str] = frozenset({"course", "group_name", "founder", "members"})
MEMBER_KEYS: FrozenSet[str] = frozenset({"full_name", "student_id", "github_username"})
KNOWN_COURSES: Tuple[str, ...] = ("MAT1206E", "MAT3508")
USERNAME_PATTERN = re.compile(
    r"(?=.{1,39}\Z)(?!-)[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?\Z"
)
PLACEHOLDER_PATTERNS: Tuple[str, ...] = ("REPLACE", "THAY ", "TODO")
MIN_MEMBERS = 1
MAX_MEMBERS = 5

# One request per repository adds up, and the secondary rate limit is shaped by
# burst rather than by total volume.  Calls are sequential, spaced, and retried
# with backoff; the limit is invisible to ``gh api rate_limit``, so there is
# nothing to check in advance and the only workable strategy is to be gentle and
# to read the refusal when it comes.
DEFAULT_TIMEOUT = 60.0
DEFAULT_PACE = 0.1
DEFAULT_RETRIES = 3
PER_PAGE = 100
MAX_PAGES = 50


class GroupRepo(NamedTuple):
    """One assignment repository, split back into the parts C50 built it from."""

    repo: str
    slug: str
    founder: str


class Collaborator(NamedTuple):
    """A repository collaborator: the login, and the id that survives a rename."""

    login: str
    github_id: str


class TeamMember(NamedTuple):
    """One row of ``team.json``.  This is the only place a student number is legitimate."""

    full_name: str
    student_id: str
    github_username: str


class TeamFault(NamedTuple):
    """One thing wrong with one ``team.json``, named so the fix can be specific."""

    code: str
    detail: str
    fix: str


class TeamFile(NamedTuple):
    """A parsed ``team.json``; ``faults`` empty means it passed every check."""

    repo: str
    course: str
    group_name: str
    founder: str
    members: List[TeamMember]
    faults: List[TeamFault]


class Group(NamedTuple):
    """One group as this lane resolved it, before conflicts are applied."""

    slug: str
    repo: str
    founder: str
    members: List[str]
    source: str
    credited: List[str]
    snapshot: List[str]
    outsiders: List[str]
    group_name: str
    # The parsed ``team.json`` rather than just the name taken out of it.  The
    # full names and student numbers a group declares there are the only
    # identity its repository carries, and dropping them means a reader that
    # wants to know who ``@someone`` is has to ask the database or GitHub.
    team: Optional[TeamFile] = None
    # What the collector published about this repository: the marks, the
    # submission history, whether the result was overridden.  It arrives in the
    # same entry as ``credited`` and was being discarded, so keeping it costs
    # nothing and saves a reader from fetching the whole gradebook again.
    result: Optional[Dict[str, Any]] = None


class GroupReport(NamedTuple):
    """What the group import found, in the terms a teacher would ask about."""

    slugs: List[str]
    groups: Dict[str, List[Dict[str, Any]]]
    quarantined: Dict[str, List[str]]
    matched: List[str]
    unmatched: List[str]
    updated: int
    skipped_team_json: List[str]
    warnings: List[str]
    findings: List[Issue]


def _group_order(group: Group) -> Tuple[str, str]:
    """Sort key for a :class:`Group`.

    A repository belongs to exactly one group, so this orders the same way the
    bare tuple used to.  It is spelled out because ``Group`` now carries a
    parsed ``team.json`` and a result payload, neither of which can be compared,
    and a bare ``sorted`` would reach them the day two rows tie on the fields
    above.
    """
    return (group.slug, group.repo)


# --------------------------------------------------------------------------
# gh plumbing
# --------------------------------------------------------------------------


def timeout_runner(seconds: float = DEFAULT_TIMEOUT) -> Runner:
    """A runner like ``_default_runner`` but with a deadline.

    The existing lanes issue one command per run, so a hung network call is
    visible as a hung command.  This lane issues one per repository, where the
    same hang stalls the whole pass with nothing on screen to say why.  The
    timeout is converted into a coded error the same way a missing binary is.
    """

    def run(argv: Sequence[str]) -> RunResult:
        try:
            completed = subprocess.run(
                list(argv),
                capture_output=True,
                text=True,
                check=False,
                timeout=seconds,
            )
        except FileNotFoundError as exc:
            raise Classroom50Error(f"missing binary: {argv[0]}", code="missing_binary") from exc
        except subprocess.TimeoutExpired as exc:
            raise Classroom50Error(
                f"{' '.join(argv[:3])} timed out after {seconds:g}s",
                code="gh_timeout",
            ) from exc
        return RunResult(completed.returncode, completed.stdout or "", completed.stderr or "")

    return run


def _is_rate_limited(message: str) -> bool:
    """Whether a 403 is a rate limit rather than a permission refusal.

    GitHub answers both with 403 and separates them only in the message, so the
    text is the whole signal.  Retrying a permission refusal wastes the budget
    that the real rate limit is about to need.
    """
    lowered = message.lower()
    return (
        "rate limit" in lowered
        or "secondary rate" in lowered
        or "abuse detection" in lowered
        or "too many requests" in lowered
    )


_RETRY_AFTER_RE = re.compile(r"retry[- ]after[:=]?\s*(\d+)", re.IGNORECASE)


def _retry_after(message: str) -> Optional[float]:
    """The server's own wait, when it stated one; documented as binding."""
    found = _RETRY_AFTER_RE.search(message)
    if not found:
        return None
    try:
        return float(found.group(1))
    except ValueError:
        return None


def _call(
    runner: Runner,
    argv: List[str],
    *,
    op: str,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
    retries: int = DEFAULT_RETRIES,
) -> str:
    """One paced, retried ``gh`` call.

    The timeout is offered to the runner as a keyword so an injected one can
    record it; a runner that does not take the argument is called without it,
    which is why :func:`timeout_runner` exists to carry the deadline instead.
    """
    attempt = 0
    while True:
        if pace:
            sleeper(pace)
        try:
            try:
                return _run(lambda a: runner(a, timeout=timeout), argv, op=op)
            except TypeError:
                return _run(runner, argv, op=op)
        except Classroom50Error as exc:
            message = str(exc)
            if not _is_rate_limited(message):
                raise
            if attempt >= retries:
                raise Classroom50Error(
                    f"{message}  (still rate limited after {attempt + 1} attempts; "
                    f"the secondary limit is not visible to `gh api rate_limit`, so "
                    f"wait a minute and run this again)",
                    code="gh_rate_limited",
                ) from exc
            sleeper(_retry_after(message) or (2.0 ** attempt))
            attempt += 1


def _paged(
    runner: Runner,
    path: str,
    *,
    op: str,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> List[Any]:
    """Every page of a list endpoint, one request at a time.

    Pages are walked by number rather than by handing the whole job to ``gh``.
    A single command that fails midway prints the pages it already had and then
    its error object into the same stream: the exit code is the only thing that
    says the list is short, and a short list here is a group with members
    missing.  One request per page keeps each answer checkable on its own.
    """
    out: List[Any] = []
    joiner = "&" if "?" in path else "?"
    for page in range(1, MAX_PAGES + 1):
        text = _call(
            runner,
            ["gh", "api", f"{path}{joiner}per_page={PER_PAGE}&page={page}"],
            op=op,
            timeout=timeout,
            sleeper=sleeper,
            pace=pace,
        )
        rows = _parse_json(text, op=op, expect=list)
        out.extend(rows)
        if len(rows) < PER_PAGE:
            break
    return out


# --------------------------------------------------------------------------
# reads
# --------------------------------------------------------------------------


def list_group_repos(
    org: str,
    classroom: str,
    slugs: Sequence[str],
    *,
    runner: Optional[Runner] = None,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> List[GroupRepo]:
    """Every repository in the org that belongs to one of these assignments.

    The whole org is listed rather than one repository guessed per roster row,
    because a student who has left the classroom still owns the repository they
    accepted, and a group whose founder left is exactly the group worth seeing.
    """
    runner = runner or timeout_runner(timeout or DEFAULT_TIMEOUT)
    rows = _paged(
        runner,
        f"/orgs/{org}/repos",
        op="list_group_repos",
        timeout=timeout,
        sleeper=sleeper,
        pace=pace,
    )
    out: List[GroupRepo] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        name = str(row.get("name") or "").strip()
        if not name:
            continue
        parsed = parse_group_repo_name(name, classroom, slugs)
        if parsed is not None:
            out.append(parsed)
    return sorted(out)


def list_repo_collaborators(
    org: str,
    repo: str,
    *,
    runner: Optional[Runner] = None,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> List[Collaborator]:
    """Every collaborator on one repository, at every permission level.

    The request carries no query beyond paging, and nothing below it drops a row
    for the rights that row holds.  Upstream's own note explains why: a teammate
    who is also an org owner is an admin on every repository, and a founder is
    often kept as an admin so they can invite the rest of the group.  Reading
    rights as a signal of "not a student" dropped both, crediting only the owner.
    """
    runner = runner or timeout_runner(timeout or DEFAULT_TIMEOUT)
    rows = _paged(
        runner,
        f"repos/{org}/{repo}/collaborators",
        op="list_repo_collaborators",
        timeout=timeout,
        sleeper=sleeper,
        pace=pace,
    )
    out: List[Collaborator] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        login = _norm_login(row.get("login"))
        if not login:
            continue
        out.append(Collaborator(login=login, github_id=str(row.get("id") or "").strip()))
    return out


def list_team_members(
    org: str,
    team_slug: str,
    *,
    runner: Optional[Runner] = None,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> List[Collaborator]:
    """The members of one classroom team, by slug."""
    runner = runner or timeout_runner(timeout or DEFAULT_TIMEOUT)
    rows = _paged(
        runner,
        f"/orgs/{org}/teams/{team_slug}/members",
        op="list_team_members",
        timeout=timeout,
        sleeper=sleeper,
        pace=pace,
    )
    out: List[Collaborator] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        login = _norm_login(row.get("login"))
        if login:
            out.append(Collaborator(login=login, github_id=str(row.get("id") or "").strip()))
    return out


def fetch_classroom_config(
    org: str,
    classroom: str,
    *,
    runner: Optional[Runner] = None,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> Dict[str, Any]:
    """Read ``<classroom>/classroom.json`` out of the config repo."""
    runner = runner or timeout_runner(timeout or DEFAULT_TIMEOUT)
    argv = [
        "gh",
        "api",
        "-H",
        "Accept: application/vnd.github.raw",
        f"repos/{org}/{CONFIG_REPO}/contents/{classroom}/classroom.json",
    ]
    text = _call(
        runner, argv, op="fetch_classroom_config", timeout=timeout, sleeper=sleeper, pace=pace
    )
    return _parse_json(text, op="fetch_classroom_config", expect=dict)


def staff_team_slugs(config: Any) -> List[str]:
    """The teacher, HTA and TA team slugs, as the configuration records them.

    The slug is read, never rebuilt: GitHub renames a team whose name collides,
    and a slug derived from the classroom short-name would then address a team
    that does not exist -- or, worse, someone else's.
    """
    out: List[str] = []
    if not isinstance(config, dict):
        return out
    teams = config.get("teams")
    if not isinstance(teams, dict):
        return out
    for role in ("teacher", "hta", "ta"):
        entry = teams.get(role)
        slug = ""
        if isinstance(entry, dict):
            slug = str(entry.get("slug") or "").strip()
        elif isinstance(entry, str):
            slug = entry.strip()
        if slug:
            out.append(slug)
    return out


def read_team_json(
    org: str,
    repo: str,
    *,
    runner: Optional[Runner] = None,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> Optional[TeamFile]:
    """One group's ``team.json``, or ``None`` when the group has not written one.

    A missing file is the ordinary state before the proposal deadline, so it is
    not an error and not a fault: the group simply has nothing to check yet.
    """
    runner = runner or timeout_runner(timeout or DEFAULT_TIMEOUT)
    argv = [
        "gh",
        "api",
        "-H",
        "Accept: application/vnd.github.raw",
        f"repos/{org}/{repo}/contents/team.json",
    ]
    try:
        text = _call(
            runner, argv, op="read_team_json", timeout=timeout, sleeper=sleeper, pace=pace
        )
    except Classroom50Error as exc:
        lowered = str(exc).lower()
        if "404" in lowered or "not found" in lowered:
            return None
        raise
    try:
        payload = json.loads(text)
    except ValueError as exc:
        return TeamFile(
            repo=repo,
            course="",
            group_name="",
            founder="",
            members=[],
            faults=[
                TeamFault(
                    code="team_json_invalid",
                    detail=f"team.json is not valid JSON: {exc}",
                    fix="Mở team.json, sửa lỗi cú pháp JSON rồi commit lại.",
                )
            ],
        )
    return parse_team_json(payload, repo=repo)


# --------------------------------------------------------------------------
# pure logic
# --------------------------------------------------------------------------


def parse_group_repo_name(
    repo: str, classroom: str, slugs: Sequence[str]
) -> Optional[GroupRepo]:
    """Split ``<classroom>-<slug>-<founder>`` back into its parts.

    Split from the front, never from the back.  A real repository name is
    ``vnu-hus-mat3508-winter-2026-w00-individual-onboarding-24007008-debug``:
    the classroom, the slug and the username all contain hyphens, so taking the
    last segment as the founder yields ``debug``.  When one slug is a prefix of
    another the longer one wins, otherwise ``w00-individual`` would claim a
    ``w00-individual-onboarding`` repository and read its founder as
    ``onboarding-24007008-debug``.
    """
    name = str(repo or "").strip().lower()
    short = str(classroom or "").strip().lower()
    if not name or not short:
        return None
    best: Optional[GroupRepo] = None
    for slug in slugs:
        clean = str(slug or "").strip().lower()
        if not clean:
            continue
        prefix = f"{short}-{clean}-"
        if not name.startswith(prefix):
            continue
        founder = name[len(prefix):]
        if not founder:
            continue
        if best is None or len(clean) > len(best.slug):
            best = GroupRepo(repo=name, slug=clean, founder=founder)
    return best


def all_roster_logins(roster: Any) -> FrozenSet[str]:
    """Every login on the classroom roster, staff included.

    Upstream intersects with exactly this set -- the union of the student team
    and all three staff teams -- so this is the set to intersect with if our
    membership list is to match the one the grades were credited against.
    """
    out: Set[str] = set()
    for row in parse_roster_payload(roster):
        login = _norm_login(row.get("username"))
        if login:
            out.add(login)
    return frozenset(out)


def credit_members(
    collaborators: Sequence[Any], *, roster_logins: FrozenSet[str], owner: str
) -> List[str]:
    """Step 3: who Classroom50 credits for this repository.

    Upstream, verbatim: the repository's collaborators at any permission level,
    intersected with the classroom team case-insensitively, sorted and deduped,
    with the owner guaranteed present.  Crediting is gated on team membership,
    not on collaborator permission.  Keep this identical to upstream even where
    a different answer would look tidier: a list that disagrees describes a
    group other than the one the marks went to.
    """
    owner_login = _norm_login(owner)
    found: Set[str] = set()
    for entry in collaborators:
        login = _norm_login(getattr(entry, "login", None) or entry)
        if login and login in roster_logins:
            found.add(login)
    if owner_login:
        found.add(owner_login)
    return sorted(found)


def student_members(credited: Sequence[str], *, staff_logins: FrozenSet[str]) -> List[str]:
    """Step 4: the students among the credited logins.

    This step is ours, not upstream's.  Every teacher who owns the organization
    appears as a collaborator on every repository and sits on the staff team, so
    step 3 keeps them by design; counting them as group members would make a
    five-student group look like a seven-member one and quarantine it for being
    over size.
    """
    return sorted(login for login in credited if _norm_login(login) not in staff_logins)


def course_for_classroom(classroom: str, courses: Sequence[str] = KNOWN_COURSES) -> str:
    """The course code named inside a classroom short-name, or ``""``."""
    lowered = str(classroom or "").lower()
    for course in courses:
        if course.lower() in lowered:
            return course
    return ""


def _placeholder_in(value: str) -> bool:
    upper = str(value or "").upper()
    return any(marker in upper for marker in PLACEHOLDER_PATTERNS)


def parse_team_json(payload: Any, *, repo: str) -> TeamFile:
    """Check one ``team.json`` against the staff checker's rules.

    Faults are collected rather than raised: one group's unfinished file must
    not stop the pass over the other groups, and a student needs to be told
    everything that is wrong at once rather than one item per run.
    """
    faults: List[TeamFault] = []
    if not isinstance(payload, dict):
        return TeamFile(
            repo=repo,
            course="",
            group_name="",
            founder="",
            members=[],
            faults=[
                TeamFault(
                    code="team_json_invalid",
                    detail=f"team.json holds a {type(payload).__name__}, not an object",
                    fix="Chép lại team.json từ template rồi điền thông tin nhóm.",
                )
            ],
        )

    keys = set(payload)
    missing = sorted(TEAM_KEYS - keys)
    extra = sorted(keys - TEAM_KEYS)
    if missing or extra:
        faults.append(
            TeamFault(
                code="team_json_keys",
                detail=(
                    f"team.json khoá thiếu {missing or '(không)'}, "
                    f"khoá thừa {extra or '(không)'}"
                ),
                fix=(
                    "team.json phải có đúng bốn khoá: course, group_name, founder, "
                    "members. Xoá khoá thừa và thêm khoá thiếu."
                ),
            )
        )

    course = str(payload.get("course") or "").strip()
    if course not in KNOWN_COURSES:
        faults.append(
            TeamFault(
                code="team_json_course",
                detail=f"course = {course!r}",
                fix=f"Điền course là một trong {', '.join(KNOWN_COURSES)}.",
            )
        )

    group_name = str(payload.get("group_name") or "").strip()
    if not group_name:
        faults.append(
            TeamFault(
                code="team_json_keys",
                detail="group_name để trống",
                fix="Điền tên nhóm vào group_name.",
            )
        )

    founder = str(payload.get("founder") or "").strip()
    if founder and not USERNAME_PATTERN.match(founder):
        faults.append(
            TeamFault(
                code="team_json_username",
                detail=f"founder = {founder!r} không phải tên đăng nhập GitHub hợp lệ",
                fix=(
                    "Điền tên đăng nhập GitHub, tức phần sau dấu gạch chéo trong "
                    "github.com/<tên đăng nhập>, không có dấu cách và không dấu tiếng Việt."
                ),
            )
        )

    raw_members = payload.get("members")
    members: List[TeamMember] = []
    if not isinstance(raw_members, list):
        faults.append(
            TeamFault(
                code="team_json_keys",
                detail=f"members là {type(raw_members).__name__}, phải là danh sách",
                fix="members phải là một danh sách các thành viên.",
            )
        )
        raw_members = []
    elif not (MIN_MEMBERS <= len(raw_members) <= MAX_MEMBERS):
        faults.append(
            TeamFault(
                code="team_json_size",
                detail=f"members có {len(raw_members)} người",
                fix=f"Nhóm phải có từ {MIN_MEMBERS} đến {MAX_MEMBERS} thành viên.",
            )
        )

    for position, row in enumerate(raw_members, start=1):
        if not isinstance(row, dict):
            faults.append(
                TeamFault(
                    code="team_json_member_keys",
                    detail=f"thành viên thứ {position} không phải một đối tượng",
                    fix="Mỗi thành viên là một đối tượng có full_name, student_id, github_username.",
                )
            )
            continue
        row_keys = set(row)
        row_missing = sorted(MEMBER_KEYS - row_keys)
        row_extra = sorted(row_keys - MEMBER_KEYS)
        if row_missing or row_extra:
            faults.append(
                TeamFault(
                    code="team_json_member_keys",
                    detail=(
                        f"thành viên thứ {position} thiếu {row_missing or '(không)'}, "
                        f"thừa {row_extra or '(không)'}"
                    ),
                    fix="Mỗi thành viên có đúng ba khoá: full_name, student_id, github_username.",
                )
            )
        full_name = str(row.get("full_name") or "").strip()
        student_id = str(row.get("student_id") or "").strip()
        username = str(row.get("github_username") or "").strip()
        if not full_name or not student_id or not username:
            faults.append(
                TeamFault(
                    code="team_json_member_keys",
                    detail=f"thành viên thứ {position} có ô để trống",
                    fix="Điền đủ full_name, student_id và github_username cho mọi thành viên.",
                )
            )
        if username and not USERNAME_PATTERN.match(username):
            faults.append(
                TeamFault(
                    code="team_json_username",
                    detail=f"github_username = {username!r} không hợp lệ",
                    fix=(
                        "Điền tên đăng nhập GitHub, tức phần sau dấu gạch chéo trong "
                        "github.com/<tên đăng nhập>."
                    ),
                )
            )
        for value in (full_name, student_id, username, group_name, founder, course):
            if _placeholder_in(value):
                faults.append(
                    TeamFault(
                        code="team_json_placeholder",
                        detail=f"còn sót chỗ điền mẫu: {value!r}",
                        fix="Thay mọi chỗ điền mẫu trong team.json bằng thông tin thật của nhóm.",
                    )
                )
                break
        members.append(
            TeamMember(full_name=full_name, student_id=student_id, github_username=username)
        )

    seen_logins: Set[str] = set()
    for member in members:
        login = _norm_login(member.github_username)
        if not login:
            continue
        if login in seen_logins:
            faults.append(
                TeamFault(
                    code="team_json_duplicate_member",
                    detail=f"{member.github_username} xuất hiện hai lần",
                    fix="Mỗi thành viên chỉ được liệt kê một lần trong team.json.",
                )
            )
        seen_logins.add(login)

    seen_ids: Set[str] = set()
    for member in members:
        key = member.student_id.strip().casefold()
        if not key:
            continue
        if key in seen_ids:
            faults.append(
                TeamFault(
                    code="team_json_duplicate_student_id",
                    detail=f"mã sinh viên {member.student_id} xuất hiện hai lần",
                    fix="Mỗi thành viên có một mã sinh viên riêng; kiểm lại xem gõ nhầm ở đâu.",
                )
            )
        seen_ids.add(key)

    if founder and _norm_login(founder) not in seen_logins:
        faults.append(
            TeamFault(
                code="team_json_founder_absent",
                detail=f"founder {founder!r} không có trong members",
                fix="Người tạo nhóm cũng là một thành viên; thêm chính mình vào members.",
            )
        )

    return TeamFile(
        repo=repo,
        course=course,
        group_name=group_name,
        founder=founder,
        members=members,
        faults=faults,
    )


def _ref_for(login: str, index: Mapping[str, Any]) -> StudentRef:
    """The record behind a login, or a placeholder naming the account.

    A finding about someone with no database record still has to be readable, so
    the login stands in for the name rather than the finding being dropped.
    """
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


def detect_group_conflicts(
    groups: Sequence[Group],
    *,
    index: Mapping[str, Any],
    team_files: Optional[Mapping[str, TeamFile]] = None,
    sheet_groups: Optional[Mapping[str, Sequence[str]]] = None,
    previous: Optional[Mapping[str, Mapping[str, Sequence[str]]]] = None,
    expected_course: str = "",
) -> List[Issue]:
    """Every disagreement between the sources, as findings.

    Nothing is repaired here.  A collaborator is never removed, a ``team.json``
    is never rewritten, and a group is never silently reshaped to match one
    source: the lists disagree because a person did something, and only a person
    can say which list is right.
    """
    team_files = team_files or {}
    sheet_groups = sheet_groups or {}
    previous = previous or {}
    findings: List[Issue] = []

    # One student, two group repositories in the same assignment.  Both groups
    # are affected, so both are named.
    by_login: Dict[Tuple[str, str], List[str]] = {}
    for group in groups:
        for login in group.members:
            by_login.setdefault((group.slug, login), []).append(group.repo)
    for (slug, login), repos in sorted(by_login.items()):
        if len(repos) < 2:
            continue
        findings.append(
            _issue(
                "group_multi_repo",
                SEVERITY_ERROR,
                login,
                index,
                field=f"{slug} nhóm",
                found=login,
                detail=f"là collaborator của {len(repos)} kho nhóm: {', '.join(sorted(repos))}",
                fix=(
                    "Mỗi bạn chỉ ở một nhóm cho mỗi bài. Rời khỏi kho của nhóm không "
                    "phải nhóm của mình, hoặc nhắn cho thầy biết nhóm đúng là nhóm nào."
                ),
            )
        )

    seen_group_names: Dict[Tuple[str, str], List[str]] = {}
    seen_student_ids: Dict[str, List[str]] = {}

    for group in sorted(groups, key=_group_order):
        if len(group.members) > MAX_MEMBERS:
            findings.append(
                _issue(
                    "group_over_size",
                    SEVERITY_ERROR,
                    group.founder,
                    index,
                    field=f"{group.slug} nhóm",
                    found=", ".join(group.members),
                    detail=(
                        f"kho {group.repo} có {len(group.members)} sinh viên, "
                        f"quá {MAX_MEMBERS}"
                    ),
                    fix=(
                        f"Nhóm tối đa {MAX_MEMBERS} bạn. Người tạo nhóm bỏ bớt "
                        "collaborator thừa khỏi kho rồi báo lại cho thầy."
                    ),
                )
            )

        for login in group.outsiders:
            findings.append(
                _issue(
                    "group_outsider",
                    SEVERITY_WARNING,
                    login,
                    index,
                    field=f"{group.slug} nhóm",
                    found=login,
                    detail=f"là collaborator của {group.repo} nhưng không có trong lớp",
                    fix=(
                        "Chỉ sinh viên trong lớp mới được tính là thành viên nhóm. "
                        "Nếu bạn này có trong lớp thì kiểm lại tên đăng nhập GitHub đã điền."
                    ),
                )
            )

        if group.source == SOURCE_SCORES:
            current = set(group.snapshot)
            credited = set(group.credited)
            added = sorted(current - credited)
            removed = sorted(credited - current)
            if added or removed:
                findings.append(
                    _issue(
                        "membership_changed_since_collect",
                        SEVERITY_WARNING,
                        group.founder,
                        index,
                        field=f"{group.slug} nhóm",
                        found=", ".join(sorted(credited)),
                        detail=(
                            f"{group.repo}: thêm {added or '(không)'}, "
                            f"bớt {removed or '(không)'} so với lúc thu điểm"
                        ),
                        fix=(
                            "Điểm đang chia theo danh sách lúc thu. Nhờ thầy bấm Sync now "
                            "để thu lại theo danh sách hiện tại."
                        ),
                    )
                )

        team = team_files.get(group.repo)
        if team is None:
            continue

        for fault in team.faults:
            findings.append(
                _issue(
                    fault.code,
                    SEVERITY_ERROR,
                    group.founder,
                    index,
                    field=f"{group.repo}/team.json",
                    found=fault.detail,
                    detail=f"team.json của {group.repo} chưa hợp lệ",
                    fix=fault.fix,
                )
            )

        declared = {_norm_login(m.github_username) for m in team.members if m.github_username}
        actual = set(group.members)
        for login in sorted(declared - actual):
            findings.append(
                _issue(
                    "team_json_extra_member",
                    SEVERITY_ERROR,
                    login,
                    index,
                    field=f"{group.slug} nhóm",
                    found=login,
                    detail=f"có trong team.json của {group.repo} nhưng không phải collaborator",
                    fix=(
                        "Người tạo nhóm vào Settings > Collaborators của kho nhóm và mời "
                        "bạn này; chưa là collaborator thì bạn ấy không được tính điểm nhóm."
                    ),
                )
            )
        for login in sorted(actual - declared):
            findings.append(
                _issue(
                    "team_json_missing_member",
                    SEVERITY_ERROR,
                    login,
                    index,
                    field=f"{group.slug} nhóm",
                    found=login,
                    detail=f"là collaborator của {group.repo} nhưng không có trong team.json",
                    fix="Thêm bạn này vào members trong team.json rồi commit lại.",
                )
            )

        if team.founder and _norm_login(team.founder) != _norm_login(group.founder):
            findings.append(
                _issue(
                    "founder_mismatch",
                    SEVERITY_ERROR,
                    group.founder,
                    index,
                    field=f"{group.repo}/team.json",
                    found=team.founder,
                    detail=f"tên kho nói người tạo nhóm là {group.founder}",
                    fix=(
                        "founder trong team.json phải là tên đăng nhập của người đã nhận "
                        "bài tập, tức phần đuôi trong tên kho nhóm."
                    ),
                )
            )

        if expected_course and team.course and team.course != expected_course:
            findings.append(
                _issue(
                    "course_mismatch",
                    SEVERITY_ERROR,
                    group.founder,
                    index,
                    field=f"{group.repo}/team.json",
                    found=team.course,
                    detail=f"kho này thuộc lớp {expected_course}",
                    fix=f"Sửa course trong team.json thành {expected_course}.",
                )
            )

        if team.group_name:
            seen_group_names.setdefault((group.slug, team.group_name.casefold()), []).append(
                group.repo
            )
        for member in team.members:
            key = member.student_id.strip().casefold()
            if key:
                seen_student_ids.setdefault(key, []).append(group.repo)

    for (slug, _name), repos in sorted(seen_group_names.items()):
        if len(set(repos)) < 2:
            continue
        findings.append(
            Issue(
                code="group_name_shared",
                severity=SEVERITY_WARNING,
                student=StudentRef(student_id="", name=f"@{sorted(repos)[0]}", email="", google_id=""),
                field=f"{slug} nhóm",
                found=", ".join(sorted(set(repos))),
                detail="hai nhóm đặt trùng tên nhóm",
                fix="Tên nhóm không phải định danh, nhưng đặt tên khác nhau thì dễ theo dõi hơn.",
            )
        )

    for student_id, repos in sorted(seen_student_ids.items()):
        if len(set(repos)) < 2:
            continue
        findings.append(
            Issue(
                code="student_id_shared_groups",
                severity=SEVERITY_ERROR,
                student=StudentRef(student_id=student_id, name="", email="", google_id=""),
                field="team.json",
                found=student_id,
                detail=f"cùng một mã sinh viên khai ở {len(set(repos))} nhóm: {', '.join(sorted(set(repos)))}",
                fix=(
                    "Một bạn chỉ ở một nhóm. Kiểm lại xem có gõ nhầm mã sinh viên trong "
                    "team.json của nhóm nào không."
                ),
            )
        )

    for group in sorted(groups, key=_group_order):
        expected = sheet_groups.get(group.repo)
        if expected is None:
            continue
        if sorted(_norm_login(login) for login in expected) != sorted(group.members):
            findings.append(
                _issue(
                    "sheet_group_mismatch",
                    SEVERITY_WARNING,
                    group.founder,
                    index,
                    field=f"{group.slug} nhóm",
                    found=", ".join(sorted(_norm_login(x) for x in expected)),
                    detail=f"khác danh sách trên kho {group.repo}: {', '.join(group.members)}",
                    fix="Hai nguồn ghi khác nhau; thầy đối chiếu và chọn danh sách đúng.",
                )
            )

    for group in sorted(groups, key=_group_order):
        before = previous.get(group.slug, {}).get(group.repo)
        if not before:
            continue
        dropped = sorted(set(_norm_login(x) for x in before) - set(group.members))
        for login in dropped:
            findings.append(
                _issue(
                    "group_member_dropped",
                    SEVERITY_ERROR,
                    login,
                    index,
                    field=f"{group.slug} nhóm",
                    found=login,
                    detail=f"không còn trong {group.repo}, lần trước thì có",
                    fix=(
                        "Bạn này mất phần điểm chung của nhóm. Người tạo nhóm mời lại vào "
                        "kho rồi báo thầy thu điểm lại."
                    ),
                )
            )

    return findings


def quarantined_units(
    groups: Sequence[Group], findings: Sequence[Issue]
) -> Dict[str, List[str]]:
    """Which repositories are held back, keyed by assignment slug.

    A finding of severity ``error`` quarantines the repository it names, and
    every other repository holding a student it names -- a person in two groups
    makes both groups unreliable, not just the one the finding happened to be
    filed against.  Warnings quarantine nothing.
    """
    logins: Set[str] = set()
    student_ids: Set[str] = set()
    repos: Set[str] = set()
    for issue in findings:
        if issue.severity != SEVERITY_ERROR:
            continue
        name = str(issue.student.name or "")
        if name.startswith("@"):
            logins.add(_norm_login(name[1:]))
        if issue.student.student_id:
            student_ids.add(issue.student.student_id.strip().casefold())
        for token in (issue.found or "").replace(",", " ").split():
            logins.add(_norm_login(token))
        for group in groups:
            if group.repo and group.repo in f"{issue.detail} {issue.found} {issue.field}":
                repos.add(group.repo)

    out: Dict[str, List[str]] = {}
    for group in groups:
        hit = group.repo in repos or _norm_login(group.founder) in logins
        if not hit:
            hit = any(_norm_login(login) in logins for login in group.members)
        if hit:
            out.setdefault(group.slug, [])
            if group.repo not in out[group.slug]:
                out[group.slug].append(group.repo)
    return {slug: sorted(values) for slug, values in out.items()}


def _team_to_dict(team: Optional[TeamFile]) -> Optional[Dict[str, Any]]:
    """A parsed ``team.json`` as plain data, for a report that has to be JSON."""
    if team is None:
        return None
    return {
        "course": team.course,
        "group_name": team.group_name,
        "founder": team.founder,
        "members": [member._asdict() for member in team.members],
        "faults": [fault._asdict() for fault in team.faults],
    }


def merge_groups_into_students(
    students: Sequence[Any],
    groups: Sequence[Group],
    findings: Sequence[Issue],
    *,
    roster: Any = None,
    index: Optional[Mapping[str, Any]] = None,
    warnings: Optional[Sequence[str]] = None,
) -> GroupReport:
    """Write the clean groups onto the records and quarantine the rest.

    Quarantine holds a write back; it never deletes.  A group that was recorded
    last week and is in conflict today keeps last week's value, because the new
    reading is the one that is not trusted, and erasing the old one would lose
    the only membership the database ever had for that group.
    """
    report_warnings: List[str] = list(warnings or [])
    if index is None:
        index, index_warnings = _student_index(students, roster)
        report_warnings.extend(index_warnings)

    held = quarantined_units(groups, findings)
    reasons: Dict[str, List[str]] = {}
    for issue in findings:
        for group in groups:
            if group.repo in f"{issue.detail} {issue.found} {issue.field}" or _norm_login(
                group.founder
            ) == _norm_login(str(issue.student.name or "").lstrip("@")):
                reasons.setdefault(group.repo, []).append(f"{issue.code}: {issue.detail}")

    matched: Set[str] = set()
    unmatched: Set[str] = set()
    updated = 0
    payload: Dict[str, List[Dict[str, Any]]] = {}

    for group in sorted(groups, key=_group_order):
        quarantined = group.repo in held.get(group.slug, [])
        payload.setdefault(group.slug, []).append(
            {
                "repo": group.repo,
                "founder": group.founder,
                "members": list(group.members),
                "source": group.source,
                "group_name": group.group_name,
                "quarantined": quarantined,
                # Everything the read resolved, not only the part the database
                # write needs: who the collector paid, who the repository has
                # as collaborators now, who is on it without being on the
                # roster, why it was held back, what the group says about
                # itself, and what it was marked.  A reader that wanted the
                # whole picture used to have to go back to GitHub for all of
                # this, although the fetch had already produced it.
                "credited": list(group.credited),
                "snapshot": list(group.snapshot),
                "outsiders": list(group.outsiders),
                "reasons": sorted(set(reasons.get(group.repo, []))),
                "team": _team_to_dict(group.team),
                "result": dict(group.result) if group.result else None,
            }
        )
        for login in group.members:
            student = index.get(_norm_login(login))
            if student is None:
                unmatched.add(login)
                continue
            matched.add(login)
            conflicts = getattr(student, FIELD_GROUP_CONFLICTS, None)
            if not isinstance(conflicts, dict):
                conflicts = {}
            group_reasons = reasons.get(group.repo, [])
            if group_reasons:
                conflicts[group.slug] = sorted(set(group_reasons))
                setattr(student, FIELD_GROUP_CONFLICTS, conflicts)
            if quarantined:
                if group.slug == MINI_PROJECT_SLUG and group_reasons:
                    setattr(student, FIELD_MP_CONFLICT, group_reasons[0])
                continue
            recorded = getattr(student, FIELD_GROUP, None)
            if not isinstance(recorded, dict):
                recorded = {}
            recorded[group.slug] = {
                "repo": group.repo,
                "founder": group.founder,
                "members": list(group.members),
                "source": group.source,
                "group_name": group.group_name,
            }
            setattr(student, FIELD_GROUP, recorded)
            updated += 1
            if group.slug == MINI_PROJECT_SLUG:
                setattr(student, FIELD_MP_REPO, group.repo)
                setattr(student, FIELD_MP_MEMBERS, ", ".join(group.members))
                setattr(student, FIELD_MP_SOURCE, group.source)
                setattr(student, FIELD_MP_CONFLICT, "")

    return GroupReport(
        slugs=sorted({group.slug for group in groups}),
        groups={slug: rows for slug, rows in sorted(payload.items())},
        quarantined=held,
        matched=sorted(matched),
        unmatched=sorted(unmatched),
        updated=updated,
        skipped_team_json=[],
        warnings=report_warnings,
        findings=list(findings),
    )


# --------------------------------------------------------------------------
# orchestration
# --------------------------------------------------------------------------


def _staff_for(
    org: str,
    classroom: str,
    roster: Any,
    *,
    runner: Runner,
    timeout: Optional[float],
    sleeper: Callable[[float], None],
    pace: float,
) -> Tuple[FrozenSet[str], List[str]]:
    """The staff logins, from the roster's role column or from the teams.

    The roster is asked first because it costs nothing: it has already been
    fetched.  The teams are only consulted when the roster answers with nobody,
    which is the case where the role column failed -- and where every teacher
    would otherwise be counted as a student and push their group over size.
    """
    warnings: List[str] = []
    found = staff_logins(roster)
    if found:
        return found, warnings
    warnings.append(
        "the roster names no teacher, HTA or TA: falling back to the classroom "
        "teams, because staff counted as students would push a full group over size"
    )
    try:
        config = fetch_classroom_config(
            org, classroom, runner=runner, timeout=timeout, sleeper=sleeper, pace=pace
        )
        out: Set[str] = set()
        for slug in staff_team_slugs(config):
            for member in list_team_members(
                org, slug, runner=runner, timeout=timeout, sleeper=sleeper, pace=pace
            ):
                out.add(member.login)
        return frozenset(out), warnings
    except Classroom50Error as exc:
        warnings.append(f"could not read the staff teams: {exc}")
        return frozenset(), warnings


def select_group_slugs(
    manifest: Mapping[str, AssignmentInfo], assignment: Optional[str] = None
) -> List[str]:
    """The group assignments to read, or the single one that was asked for."""
    group_slugs = sorted(
        slug for slug, info in manifest.items() if info.mode == TYPE_GROUP
    )
    if not assignment:
        return group_slugs
    wanted = assignment.strip()
    if wanted not in group_slugs:
        raise Classroom50Error(
            f"{wanted!r} is not a group assignment in this classroom; the group "
            f"assignments are: {', '.join(group_slugs) or '(none)'}",
            code="unknown_assignment",
        )
    return [wanted]


def import_groups(
    students: Sequence[Any],
    *,
    org: str,
    classroom: str,
    assignment: Optional[str] = None,
    with_team_json: bool = False,
    sheet_groups: Optional[Mapping[str, Sequence[str]]] = None,
    previous: Optional[Mapping[str, Mapping[str, Sequence[str]]]] = None,
    cli: Optional[AgentCLI] = None,
    runner: Optional[Runner] = None,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> GroupReport:
    """Read every group assignment's membership and write the clean ones down.

    The two group assignments are read by two different routes, and the reason is
    upstream's, not a preference: one is autograded, so the collector has already
    published the member list it credited, and reading anything else invites a
    list that disagrees with where the marks went.  The other has autograding off
    and an empty repository, so the collector skips it and never publishes a
    bucket for it at all -- its membership can only be computed from the
    collaborators.
    """
    cli = cli or AgentCLI(runner=runner)
    api = runner or timeout_runner(timeout or DEFAULT_TIMEOUT)

    roster = cli.list_roster(org, classroom)
    staff, warnings = _staff_for(
        org, classroom, roster, runner=api, timeout=timeout, sleeper=sleeper, pace=pace
    )
    roster_logins = all_roster_logins(roster)
    manifest = parse_assignments(cli.list_assignments(org, classroom))
    slugs = select_group_slugs(manifest, assignment)
    if not slugs:
        return GroupReport(
            slugs=[],
            groups={},
            quarantined={},
            matched=[],
            unmatched=[],
            updated=0,
            skipped_team_json=[],
            warnings=warnings + ["this classroom has no group assignment"],
            findings=[],
        )

    # A run over the ungraded assignment alone must not read the gradebook: the
    # collector never writes a bucket for it, so the read could only ever return
    # nothing, and a private repository answers a failed read with 404.
    credited_by_slug: Dict[str, Dict[str, List[str]]] = {}
    # ``slug -> owner -> what the collector marked that repository``.  Read from
    # the same entries as ``credited_by_slug`` and in the same pass, so it adds
    # no request; without it a reader wanting a group's marks has to fetch the
    # gradebook a second time and re-derive which repository each one came from.
    results_by_slug: Dict[str, Dict[str, Dict[str, Any]]] = {}
    if any(manifest[slug].graded for slug in slugs if slug in manifest):
        # fetch_scores already validates the schema and returns a parsed
        # document, so there is nothing left to parse here.
        from .c50_scores import fetch_scores

        try:
            doc = fetch_scores(org, classroom, runner=api)
        except Classroom50Error as exc:
            warnings.append(f"could not read the collected scores: {exc}")
        else:
            warnings.extend(doc.warnings)
            for slug, bucket in doc.assignments.items():
                if slug not in slugs or bucket.type != TYPE_GROUP:
                    continue
                for entry in bucket.entries:
                    owner = _norm_login(entry.owner)
                    if owner:
                        credited_by_slug.setdefault(slug, {})[owner] = sorted(
                            {_norm_login(name) for name in entry.member_usernames if name}
                            | {owner}
                        )
                        results_by_slug.setdefault(slug, {})[owner] = {
                            "override": entry.override,
                            "submissions": list(entry.submissions),
                        }

    repos = list_group_repos(
        org, classroom, slugs, runner=api, timeout=timeout, sleeper=sleeper, pace=pace
    )

    groups: List[Group] = []
    team_files: Dict[str, TeamFile] = {}
    skipped_team_json: List[str] = []
    for entry in repos:
        collaborators = list_repo_collaborators(
            org, entry.repo, runner=api, timeout=timeout, sleeper=sleeper, pace=pace
        )
        credited_now = credit_members(
            collaborators, roster_logins=roster_logins, owner=entry.founder
        )
        snapshot = student_members(credited_now, staff_logins=staff)
        outsiders = sorted(
            {c.login for c in collaborators} - set(roster_logins) - {_norm_login(entry.founder)}
        )
        published = credited_by_slug.get(entry.slug, {}).get(_norm_login(entry.founder))
        if published is None:
            members = snapshot
            credited = snapshot
            source = SOURCE_COLLABORATORS
        else:
            members = student_members(published, staff_logins=staff)
            credited = student_members(published, staff_logins=staff)
            source = SOURCE_SCORES

        group_name = ""
        team: Optional[TeamFile] = None
        if with_team_json:
            team = read_team_json(
                org, entry.repo, runner=api, timeout=timeout, sleeper=sleeper, pace=pace
            )
            if team is None:
                skipped_team_json.append(entry.repo)
            else:
                team_files[entry.repo] = team
                group_name = team.group_name

        groups.append(
            Group(
                slug=entry.slug,
                repo=entry.repo,
                founder=_norm_login(entry.founder),
                members=members,
                source=source,
                credited=credited,
                snapshot=snapshot,
                outsiders=outsiders,
                group_name=group_name,
                team=team,
                result=results_by_slug.get(entry.slug, {}).get(
                    _norm_login(entry.founder)
                ),
            )
        )

    index, index_warnings = _student_index(students, roster)
    warnings.extend(index_warnings)
    findings = detect_group_conflicts(
        groups,
        index=index,
        team_files=team_files,
        sheet_groups=sheet_groups,
        previous=previous,
        expected_course=course_for_classroom(classroom),
    )
    report = merge_groups_into_students(
        students, groups, findings, index=index, warnings=warnings
    )
    return report._replace(skipped_team_json=sorted(skipped_team_json))


def group_snapshot(students: Sequence[Any]) -> Dict[str, Dict[str, List[str]]]:
    """``{slug: {repo: members}}`` as the database currently records it.

    Taken before a merge overwrites it: the comparison against the next reading
    is the only way a student who was dropped from a group repository becomes
    visible, since nothing upstream keeps that history.
    """
    out: Dict[str, Dict[str, List[str]]] = {}
    for student in students:
        recorded = getattr(student, FIELD_GROUP, None)
        if not isinstance(recorded, dict):
            continue
        for slug, value in recorded.items():
            if not isinstance(value, dict):
                continue
            repo = str(value.get("repo") or "")
            members = value.get("members")
            if repo and isinstance(members, list):
                out.setdefault(str(slug), {})[repo] = [str(m) for m in members]
    return out


def report_to_dict(report: GroupReport) -> Dict[str, Any]:
    """The report as plain JSON-ready data, for the agent entrypoint."""
    return {
        "slugs": report.slugs,
        "groups": report.groups,
        "quarantined": report.quarantined,
        "matched": report.matched,
        "unmatched": report.unmatched,
        "updated": report.updated,
        "skipped_team_json": report.skipped_team_json,
        "warnings": report.warnings,
        # Same shape as ``roster_audit.report_to_dict`` writes its findings, so
        # one reader handles both reports.
        "findings": [
            {**issue._asdict(), "student": issue.student._asdict()}
            for issue in report.findings
        ],
    }


def format_groups(report: GroupReport) -> str:
    """The report as a teacher would read it."""
    from .roster_audit import format_issues

    held = sum(len(repos) for repos in report.quarantined.values())
    total = sum(len(rows) for rows in report.groups.values())
    lines = [
        f"Classroom50 groups: {total} group(s) across {len(report.slugs)} assignment(s), "
        f"{report.updated} record(s) updated, {held} group(s) quarantined"
    ]
    for slug in report.slugs:
        rows = report.groups.get(slug, [])
        lines.append(f"  {slug}: {len(rows)} group(s)")
        for row in rows:
            mark = " [quarantined]" if row.get("quarantined") else ""
            members = ", ".join(row.get("members") or []) or "(none)"
            lines.append(
                f"    {row.get('repo')} ({row.get('source')}): {members}{mark}"
            )
    if report.unmatched:
        lines.append(f"  no database record for: {', '.join(report.unmatched)}")
    if report.skipped_team_json:
        lines.append(
            f"  no team.json yet: {', '.join(report.skipped_team_json)}"
        )
    for warning in report.warnings:
        lines.append(f"  warning: {warning}")
    text = "\n".join(lines)
    if report.findings:
        text += "\n\n" + format_issues(report.findings)
    return text


# --------------------------------------------------------------------------
# the long-form export
# --------------------------------------------------------------------------


def _repo_url(org: str, repo: str) -> str:
    """A repository's browser address, for a teacher who wants to open it."""
    return f"https://github.com/{org}/{repo}" if org and repo else ""


def _mark(value: Any, yes: str = "có", no: str = "không") -> str:
    """A flag in words.  ``None`` is not the same answer as ``False``."""
    if value is None:
        return "không rõ"
    return yes if value else no


def _submission_points(payload: Mapping[str, Any]) -> str:
    """One submission's marks as ``85/100``, or why there are none."""
    score = payload.get("score")
    if score is None:
        return "chưa chấm"
    maximum = payload.get("max-score")
    return f"{score}/{maximum}" if maximum is not None else f"{score}"


def _member_lines(
    login: str,
    position: int,
    *,
    slug: str,
    founder: str,
    team: Optional[Mapping[str, Any]],
    index: Mapping[str, Any],
) -> List[str]:
    """One member: the login, who the database says that is, who they say they are.

    The three never come from the same place, and the point of printing them
    together is that they can disagree -- a student number in ``team.json`` that
    belongs to somebody else is invisible until the two sit on adjacent lines.
    """
    from .c50_scores import FIELD_GRADE_CANDIDATES, FIELD_GRADES, _points

    key = _norm_login(login)
    role = "  [người tạo kho]" if key and key == _norm_login(founder) else ""
    out = [f"    {position}. @{login}{role}"]

    student = index.get(key)
    if student is None:
        out.append("       (không có bản ghi trong database)")
    else:
        for label, field in (
            ("MSSV", "Student ID"),
            ("Họ tên", "Name"),
            ("Email", "Email"),
            ("Email HUS", "Emai VNU-HUS"),
            ("Lớp", "Class"),
            ("Nhóm lớp", "Section"),
        ):
            value = str(getattr(student, field, "") or "").strip()
            if value:
                out.append(f"       {label:<11}: {value}")

    if isinstance(team, dict):
        for member in team.get("members") or []:
            if not isinstance(member, dict):
                continue
            if _norm_login(member.get("github_username")) != key:
                continue
            told = " — ".join(
                part
                for part in (
                    str(member.get("full_name") or "").strip(),
                    str(member.get("student_id") or "").strip(),
                )
                if part
            )
            out.append(f"       {'Tự khai':<11}: {told or '(bỏ trống)'}  (theo team.json)")

    if student is not None:
        grades = getattr(student, FIELD_GRADES, None)
        stored = grades.get(slug) if isinstance(grades, dict) else None
        if isinstance(stored, dict):
            out.append(
                f"       {'Điểm trong sổ':<11}: {_points(stored)} "
                f"(ghi từ kho {stored.get('owner') or 'không rõ'})"
            )
        candidates = getattr(student, FIELD_GRADE_CANDIDATES, None)
        paying = candidates.get(slug) if isinstance(candidates, dict) else None
        if isinstance(paying, dict) and len(paying) > 1:
            shown = ", ".join(
                f"{owner}: {_points(paying[owner])}"
                for owner in sorted(paying)
                if isinstance(paying[owner], dict)
            )
            out.append(
                f"       {'CẢNH BÁO':<11}: {len(paying)} kho cùng tính điểm cho bạn "
                f"này — {shown}"
            )
    return out


def _group_block(
    row: Mapping[str, Any],
    *,
    slug: str,
    position: int,
    of: int,
    org: str,
    index: Mapping[str, Any],
    thin: str,
) -> List[str]:
    """One group, in the order a teacher would ask the questions."""
    repo = str(row.get("repo") or "")
    founder = str(row.get("founder") or "")
    declared = str(row.get("group_name") or "")
    members = [str(name) for name in (row.get("members") or [])]
    team = row.get("team") if isinstance(row.get("team"), dict) else None
    result = row.get("result") if isinstance(row.get("result"), dict) else None

    title = declared or (f"@{founder}" if founder else "(nhóm chưa rõ tên)")
    out = ["", thin, f"[{position}/{of}] {title}", thin]
    out.append(f"  Kho            : {repo or '(không rõ)'}")
    url = _repo_url(org, repo)
    if url:
        out.append(f"  Địa chỉ        : {url}")
    out.append(f"  Người tạo kho  : {'@' + founder if founder else '(không rõ)'}")
    out.append(f"  Tên nhóm       : {declared or '(chưa đặt)'}")

    out.append("")
    out.append(f"  Thành viên ({len(members)}):")
    if not members:
        out.append("    (không có)")
    for number, login in enumerate(members, start=1):
        out.extend(
            _member_lines(
                login, number, slug=slug, founder=founder, team=team, index=index
            )
        )

    out.append("")
    out.append("  Đối chiếu danh sách:")
    out.append(f"      {'Nguồn đang dùng':<24}: {row.get('source') or '(không rõ)'}")
    for label, field in (
        ("Classroom50 tính điểm", "credited"),
        ("Cộng tác viên hiện tại", "snapshot"),
        ("Ngoài danh sách lớp", "outsiders"),
    ):
        names = ", ".join(str(name) for name in (row.get(field) or []))
        out.append(f"      {label:<24}: {names or '(không có)'}")

    out.append("")
    if result is None:
        out.append("  Điểm của kho   : chưa có trong scores.json")
    else:
        submissions = [
            one for one in (result.get("submissions") or []) if isinstance(one, dict)
        ]
        newest = submissions[0] if submissions else {}
        out.append("  Điểm của kho:")
        out.append(f"      {'Kết quả':<24}: {_submission_points(newest)}")
        out.append(f"      {'Nộp lúc':<24}: {newest.get('datetime') or '(không rõ)'}")
        out.append(f"      {'Trễ hạn':<24}: {_mark(newest.get('late'))}")
        if newest.get("commit"):
            out.append(f"      {'Commit':<24}: {newest.get('commit')}")
        if newest.get("release"):
            out.append(f"      {'Release':<24}: {newest.get('release')}")
        out.append(f"      {'Chấm tay ghi đè':<24}: {_mark(result.get('override'))}")
        if len(submissions) > 1:
            out.append(f"      Các lần nộp ({len(submissions)}):")
            for number, one in enumerate(submissions, start=1):
                out.append(
                    f"        {number}. {one.get('datetime') or '(không rõ)'} — "
                    f"{_submission_points(one)} — trễ: {_mark(one.get('late'))}"
                )

    out.append("")
    if team is None:
        out.append("  team.json      : chưa đọc (cần --read-team-json) hoặc kho chưa có")
    else:
        out.append("  team.json:")
        out.append(f"      {'Môn':<24}: {team.get('course') or '(bỏ trống)'}")
        out.append(f"      {'Tên nhóm':<24}: {team.get('group_name') or '(bỏ trống)'}")
        out.append(f"      {'Trưởng nhóm khai báo':<24}: {team.get('founder') or '(bỏ trống)'}")
        declared_members = [
            one for one in (team.get("members") or []) if isinstance(one, dict)
        ]
        out.append(f"      {'Số thành viên khai báo':<24}: {len(declared_members)}")
        faults = [one for one in (team.get("faults") or []) if isinstance(one, dict)]
        if not faults:
            out.append(f"      {'Kiểm tra':<24}: hợp lệ")
        else:
            out.append(f"      Lỗi ({len(faults)}):")
            for fault in faults:
                out.append(f"        - {fault.get('code')}: {fault.get('detail')}")
                if fault.get("fix"):
                    out.append(f"          cách sửa: {fault.get('fix')}")

    out.append("")
    out.append(
        "  Trạng thái     : GIỮ LẠI — không ghi vào database"
        if row.get("quarantined")
        else "  Trạng thái     : đã ghi vào database"
    )
    for reason in row.get("reasons") or []:
        out.append(f"      - {reason}")
    return out


def format_groups_txt(
    report: GroupReport,
    *,
    org: str = "",
    classroom: str = "",
    students: Sequence[Any] = (),
    collected_at: str = "",
    generated_at: str = "",
) -> str:
    """The whole reading as a file to keep, one block per group.

    :func:`format_groups` answers "did the import go well".  This answers "who
    is in this group, where is their repository, and what were they marked",
    which needs the three places that knowledge lives -- the repositories, the
    ``team.json`` each group wrote about itself, and the local database -- none
    of which knows what the other two hold.

    ``students`` may be empty.  The file is then written without student
    numbers, names or recorded grades, and says so at the top rather than
    leaving a reader to conclude that the members are only GitHub logins.
    """
    from .roster_audit import format_issues

    index, _ = _student_index(students, None)
    rule = "=" * 78
    thin = "-" * 78
    total = sum(len(rows) for rows in report.groups.values())
    held = sum(len(repos) for repos in report.quarantined.values())

    lines = [rule, "THÔNG TIN NHÓM — Classroom50", rule]
    if org:
        lines.append(f"Tổ chức        : {org}")
    if classroom:
        lines.append(f"Lớp            : {classroom}")
    lines.append(f"Bài tập        : {', '.join(report.slugs) or '(không có)'}")
    if generated_at:
        lines.append(f"Xuất lúc       : {generated_at}")
    if collected_at:
        lines.append(f"Điểm thu lúc   : {collected_at}")
    lines.append(
        f"Tổng cộng      : {total} nhóm, {held} nhóm bị giữ lại, "
        f"{report.updated} bản ghi được cập nhật"
    )
    if not students:
        lines.append(
            "Ghi chú        : chạy không kèm database, nên không có MSSV, họ tên "
            "hay điểm trong sổ"
        )

    for slug in report.slugs:
        rows = report.groups.get(slug, [])
        lines.extend(["", rule, f"BÀI: {slug} — {len(rows)} nhóm", rule])
        for position, row in enumerate(rows, start=1):
            lines.extend(
                _group_block(
                    row,
                    slug=slug,
                    position=position,
                    of=len(rows),
                    org=org,
                    index=index,
                    thin=thin,
                )
            )

    lines.extend(["", rule, "PHẦN CÒN LẠI CỦA BÁO CÁO", rule])
    lines.append(
        f"Không có bản ghi trong database ({len(report.unmatched)}): "
        + (", ".join(report.unmatched) or "(không có)")
    )
    lines.append(
        f"Chưa có team.json ({len(report.skipped_team_json)}): "
        + (", ".join(report.skipped_team_json) or "(không có)")
    )
    for warning in report.warnings:
        lines.append(f"Cảnh báo: {warning}")

    text = "\n".join(lines)
    if report.findings:
        text += f"\n\n{rule}\nLỖI VÀ CẢNH BÁO THEO SINH VIÊN\n{rule}\n"
        text += format_issues(report.findings)
    return text + "\n"


def list_groups(
    *,
    org: str,
    classroom: str,
    assignment: Optional[str] = None,
    with_team_json: bool = False,
    cli: Optional[AgentCLI] = None,
    runner: Optional[Runner] = None,
    timeout: Optional[float] = DEFAULT_TIMEOUT,
    sleeper: Callable[[float], None] = time.sleep,
    pace: float = DEFAULT_PACE,
) -> Dict[str, Any]:
    """The same reading as :func:`import_groups`, with nothing written down.

    It runs the whole pipeline over an empty record list, so the conflict checks
    that only need the repositories still report.  The checks that need the
    database -- who a login belongs to, what was recorded last week -- have
    nothing to work with here, and the report says so by listing every login as
    unmatched.
    """
    report = import_groups(
        [],
        org=org,
        classroom=classroom,
        assignment=assignment,
        with_team_json=with_team_json,
        cli=cli,
        runner=runner,
        timeout=timeout,
        sleeper=sleeper,
        pace=pace,
    )
    return report_to_dict(report)
