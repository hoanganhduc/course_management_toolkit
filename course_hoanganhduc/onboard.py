# -*- coding: utf-8 -*-
"""Interactive onboarding: one registration CSV -> Google Classroom + Classroom50.

The three pieces this joins already existed and none of them talked to the
others: ``roster_csv`` loads a CSV into the database, ``gclass_invite`` invites
database students to a course, and ``c50_admin_cli roster-import`` upserts the
Classroom50 roster.  Running them by hand means knowing the course id and the
org/classroom pair up front, and it means trusting a registration form that
students fill in wrong -- a missing student id, a personal address where the
school one was asked for, a GitHub username copied from a friend.

So this module reads the form, says plainly which rows it cannot use and why,
shows what each platform would receive, and only then hands the work back to the
two lanes that already do it.  Nothing here re-implements an invitation.

Skipping a platform is a first-class answer: choosing neither still loads the
CSV into the database and prints the validation report, which is the useful
thing to run while a form is still collecting responses.

Imports: ``data``, ``gclass_invite`` and ``c50_admin_cli`` are imported lazily
inside the functions that need them -- ``data`` pulls pandas/openpyxl/sklearn at
module scope and ``gclass_invite`` pulls googleapiclient, and this module has to
stay importable, and testable, without either stack.
"""

from __future__ import annotations

import re
from typing import (
    Any,
    Callable,
    Dict,
    List,
    NamedTuple,
    Optional,
    Sequence,
    Set,
    Tuple,
    Union,
)

from .c50_cli import Classroom50Error, Runner, _default_runner
from .c50_flags import (
    SOURCE_CONFIG,
    SOURCE_FLAG,
    Resolved,
    resolve_classroom,
    resolve_org,
    resolve_setting,
)
from .roster_csv import ParseResult

# Pragmatic rather than RFC 5322: a form address that this rejects is far more
# likely to be a typo than a legal exotic address, and the fallback to the
# personal column means a rejection here is rarely fatal to the row.
_EMAIL_RE = re.compile(
    r"^[A-Za-z0-9._%+-]+"
    r"@[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?"
    r"(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?)+$"
)

# GitHub's own account rule: alphanumeric or single hyphens, no leading or
# trailing hyphen, 39 characters at most.  Deliberately *not* c50_cli_human's
# _SLUG_RE, which governs classroom slugs and allows a leading digit-only name
# this must reject.
_GITHUB_USERNAME_RE = re.compile(r"^[A-Za-z0-9](?:[A-Za-z0-9]|-(?=[A-Za-z0-9])){0,38}$")

_GITHUB_URL_RE = re.compile(r"^(?:https?://)?(?:www\.)?github\.com/", re.I)

GITHUB_EXISTS = "exists"
GITHUB_MISSING = "missing"
GITHUB_UNVERIFIED = "unverified"

# Order matters: the school address is what the two platforms should carry, and
# the personal one only stands in when the school column is empty or malformed.
EMAIL_SOURCE_SCHOOL = "vnu-hus"
EMAIL_SOURCE_PERSONAL = "personal"


class GithubCheck(NamedTuple):
    """What `gh api /users/<name>` said about one username.

    ``unverified`` is not a failure.  A rate limit or a dropped connection must
    not throw out a whole class, so an unverified row still goes through with an
    empty github_id and a line in the report.
    """

    username: str
    state: str
    github_id: str
    detail: str


class ResolvedRow(NamedTuple):
    """One form row that survived validation, in database attribute terms."""

    student_id: str
    name: str
    email: str
    email_source: str
    github_username: str
    github_id: str
    section: str
    class_name: str
    timestamp: str
    notes: List[str]


class RejectedRow(NamedTuple):
    """One form row that cannot be used, and the reason to tell the student."""

    student_id: str
    name: str
    reason: str


class ResolveResult(NamedTuple):
    ok: List[ResolvedRow]
    rejected: List[RejectedRow]
    duplicates: List[str]


def validate_email(raw: Any) -> Tuple[str, Optional[str]]:
    """Return ``(email, None)`` or ``("", reason)``.

    Lower-cased, because both platforms match addresses case-insensitively and
    a form entry capitalised by a phone keyboard should not read as a different
    person.
    """
    text = str(raw or "").strip()
    if not text:
        return "", "để trống"
    if text.lower().startswith("mailto:"):
        text = text[len("mailto:"):].strip()
    text = text.strip("<>").strip()
    if " " in text:
        return "", f"có khoảng trắng: {text!r}"
    if not _EMAIL_RE.match(text):
        return "", f"sai định dạng: {text!r}"
    tld = text.rsplit(".", 1)[-1]
    if len(tld) < 2 or not tld.isalpha():
        return "", f"đuôi tên miền không hợp lệ: {text!r}"
    return text.lower(), None


def validate_github_username(raw: Any) -> Tuple[str, Optional[str]]:
    """Return ``(username, None)`` or ``("", reason)``.

    Students paste a profile URL as often as they type a name, so the URL form
    and a leading ``@`` are normalised away before the rule is applied rather
    than being rejected.
    """
    text = str(raw or "").strip()
    if not text:
        return "", "để trống"
    text = _GITHUB_URL_RE.sub("", text)
    text = text.split("?")[0].split("#")[0]
    text = text.strip("/").split("/")[0]
    text = text.lstrip("@").strip()
    if not text:
        return "", "chỉ có ký tự phân cách"
    if len(text) > 39:
        return "", f"dài quá 39 ký tự: {text!r}"
    if not _GITHUB_USERNAME_RE.match(text):
        return "", f"sai cú pháp tài khoản GitHub: {text!r}"
    return text, None


def check_github_exists(
    names: Sequence[str],
    *,
    runner: Optional[Runner] = None,
    skip: bool = False,
    verbose: bool = False,
) -> Dict[str, GithubCheck]:
    """Ask GitHub whether each username exists, keyed by the lower-cased name.

    The same endpoint ``gh teacher roster import`` uses to re-resolve github_id,
    so the id captured here is the one the import would derive itself.  With
    ``skip`` the syntax check stands alone and every name comes back
    ``unverified``; that is the offline mode, not a silent pass.
    """
    runner = runner or _default_runner
    checks: Dict[str, GithubCheck] = {}
    for raw in names:
        key = str(raw or "").strip().lower()
        if not key or key in checks:
            continue
        if skip:
            checks[key] = GithubCheck(raw, GITHUB_UNVERIFIED, "", "bỏ qua kiểm tra API")
            continue
        argv = ["gh", "api", f"/users/{raw}", "--jq", ".id"]
        try:
            result = runner(argv)
        except Classroom50Error as exc:
            checks[key] = GithubCheck(raw, GITHUB_UNVERIFIED, "", str(exc))
            continue
        except OSError as exc:
            checks[key] = GithubCheck(raw, GITHUB_UNVERIFIED, "", str(exc))
            continue
        stdout = (getattr(result, "stdout", "") or "").strip()
        stderr = (getattr(result, "stderr", "") or "").strip()
        if getattr(result, "returncode", 1) == 0 and stdout:
            checks[key] = GithubCheck(raw, GITHUB_EXISTS, stdout.splitlines()[0].strip(), "")
        elif _looks_like_not_found(stderr):
            checks[key] = GithubCheck(raw, GITHUB_MISSING, "", "tài khoản không tồn tại")
        else:
            # Rate limit, no network, no auth: cannot tell, so do not decide.
            checks[key] = GithubCheck(raw, GITHUB_UNVERIFIED, "", stderr or "không rõ lỗi")
        if verbose:
            print(f"[Onboard] gh api /users/{raw} -> {checks[key].state}")
    return checks


def _looks_like_not_found(stderr: str) -> bool:
    lowered = (stderr or "").lower()
    return "404" in lowered or "not found" in lowered


def _cell(row: Dict[str, str], key: str) -> str:
    return str(row.get(key, "") or "").strip()


def _pick_email(row: Dict[str, str]) -> Tuple[str, str, Optional[str]]:
    """Choose the address to use, preferring the school one (decision 1).

    Returns ``(email, source, reason)``; ``reason`` is set only when neither
    column yields a usable address, and it names both failures so the student
    can be told which field to fix.
    """
    school, school_reason = validate_email(_cell(row, "Email"))
    if school:
        return school, EMAIL_SOURCE_SCHOOL, None
    personal, personal_reason = validate_email(_cell(row, "Personal Email"))
    if personal:
        return personal, EMAIL_SOURCE_PERSONAL, None
    return "", "", f"email trường {school_reason}; email cá nhân {personal_reason}"


def resolve_rows(
    parsed: ParseResult,
    checks: Optional[Dict[str, GithubCheck]] = None,
) -> ResolveResult:
    """Turn parsed CSV rows into rows the two lanes can act on.

    Every drop is recorded with the student's name and the field at fault, since
    the point of the report is a list of people to go and ask.  Two rows that
    declare the same GitHub username are *both* dropped: an import upserts by
    username, so keeping either one would silently overwrite the other, and
    there is no evidence here for which student typed it correctly.
    """
    checks = checks or {}
    rejected: List[RejectedRow] = []
    duplicates: List[str] = []

    # Google Forms appends, so a later row is a later submission by the same
    # student; the last one wins without having to guess a date format.
    by_key: Dict[str, ResolvedRow] = {}
    order: List[str] = []

    for row in parsed.rows:
        student_id = _cell(row, "Student ID")
        name = _cell(row, "Name")

        email, email_source, email_reason = _pick_email(row)
        raw_github = _cell(row, "GitHub Username")
        github, github_reason = validate_github_username(raw_github)

        notes: List[str] = []
        github_id = ""
        if github:
            check = checks.get(github.lower())
            if check is None:
                notes.append("GitHub chưa được kiểm tra")
            elif check.state == GITHUB_MISSING:
                rejected.append(RejectedRow(student_id, name, f"GitHub {github!r} không tồn tại"))
                continue
            elif check.state == GITHUB_UNVERIFIED:
                # Empty is safe, a stale id is not: roster import re-resolves the
                # id from the username and fails the line when the two disagree.
                notes.append(f"GitHub chưa kiểm được ({check.detail})")
            else:
                github_id = check.github_id
        elif raw_github:
            rejected.append(RejectedRow(student_id, name, f"GitHub Username {github_reason}"))
            continue

        if not email and not github:
            reason = email_reason or "thiếu cả email lẫn GitHub Username"
            rejected.append(RejectedRow(student_id, name, reason))
            continue
        if not email:
            notes.append(f"không có email dùng được ({email_reason})")

        key = student_id or github.lower() or email
        if not key:
            rejected.append(RejectedRow(student_id, name, "không có Mã Sinh Viên, GitHub, hay email để nhận dạng"))
            continue

        resolved = ResolvedRow(
            student_id=student_id,
            name=name,
            email=email,
            email_source=email_source,
            github_username=github,
            github_id=github_id,
            section=_cell(row, "Section"),
            class_name=_cell(row, "Class"),
            timestamp=_cell(row, "Timestamp"),
            notes=notes,
        )
        if key in by_key:
            previous = by_key[key]
            duplicates.append(
                f"{name or student_id}: nộp form nhiều lần, giữ bản sau"
                + (f" ({resolved.timestamp})" if resolved.timestamp else "")
                + (f", bỏ bản trước ({previous.timestamp})" if previous.timestamp else "")
            )
        else:
            order.append(key)
        by_key[key] = resolved

    ok = [by_key[key] for key in order]
    ok, clashes = _drop_github_clashes(ok)
    rejected.extend(clashes)
    return ResolveResult(ok, rejected, duplicates)


def _drop_github_clashes(rows: Sequence[ResolvedRow]) -> Tuple[List[ResolvedRow], List[RejectedRow]]:
    """Drop every row in a group that shares a GitHub username with another."""
    groups: Dict[str, List[ResolvedRow]] = {}
    for row in rows:
        if row.github_username:
            groups.setdefault(row.github_username.lower(), []).append(row)

    clashing: Set[int] = set()
    rejected: List[RejectedRow] = []
    for username, group in groups.items():
        if len(group) < 2:
            continue
        others = ", ".join(f"{r.name or '?'} ({r.student_id or 'không có MSSV'})" for r in group)
        for row in group:
            clashing.add(id(row))
            rejected.append(
                RejectedRow(
                    row.student_id,
                    row.name,
                    f"GitHub {username!r} bị khai trùng bởi: {others}",
                )
            )
    kept = [row for row in rows if id(row) not in clashing]
    return kept, rejected


# --- presets ---------------------------------------------------------------
#
# A course that has been configured once should not have to be re-picked every
# term, so the driver accepts the course id, the org and the classroom up front.
# Where each value came from decides how much it is trusted: a flag was typed
# for this run and is an instruction, while a config file or an environment
# variable was set once and applies to every run after that -- which is exactly
# how a course id quietly goes stale between semesters.  Those are shown and
# confirmed, and declining falls through to the picker.


class Presets(NamedTuple):
    """The three settings the driver can be handed before it asks anything."""

    google_course_id: Resolved = Resolved()
    c50_org: Resolved = Resolved()
    c50_classroom: Resolved = Resolved()


def resolve_presets(args: Any, config: Optional[Dict[str, Any]]) -> Presets:
    """Read all three settings from the flag, the config file, then the environment.

    The Google course id goes through the same resolver as the Classroom50 pair,
    which gives it the ``GOOGLE_CLASSROOM_COURSE_ID`` config tier the seven other
    Google handlers in ``core.py`` already read and this lane alone was missing.
    """
    cfg = config if isinstance(config, dict) else None
    return Presets(
        google_course_id=resolve_setting(
            args, cfg, attribute="google_course_id", key="GOOGLE_CLASSROOM_COURSE_ID"
        ),
        c50_org=resolve_org(args, cfg),
        c50_classroom=resolve_classroom(args, cfg),
    )


Setting = Union[str, Resolved, None]


def _as_resolved(value: Setting) -> Resolved:
    """Normalise a driver argument to a value plus the tier it came from.

    A bare string keeps the meaning it has always had -- something the caller
    chose for this run, used as given.  Only a :class:`Resolved` can carry the
    weaker claim that asks to be confirmed.
    """
    if isinstance(value, Resolved):
        return value
    text = str(value or "").strip()
    return Resolved(text, SOURCE_FLAG) if text else Resolved()


def confirm_preset(
    preset: Resolved,
    label: str,
    *,
    input_fn: Callable[[str], str] = input,
) -> Optional[str]:
    """Show a value that was not typed for this run, and ask before using it.

    Returns the value on yes and ``None`` on no, so the caller falls through to
    the picker as if nothing had been supplied.  Enter means yes, because a
    default that has to be retyped is not a default.
    """
    origin = "cấu hình" if preset.source == SOURCE_CONFIG else "biến môi trường"
    answer = str(
        input_fn(f"{label} từ {origin}: {preset.value} — dùng giá trị này? (Y/n): ") or ""
    ).strip().lower()
    return preset.value if answer in ("", "y", "yes") else None


def _accept_preset(
    preset: Resolved,
    label: str,
    *,
    ask: bool,
    input_fn: Callable[[str], str],
) -> Optional[str]:
    """The value to use, confirmed first unless it was typed for this run.

    ``ask`` is false on a dry run and off a terminal.  Both keep the value: a
    scheduled run that silently ignored its own configuration would be harder to
    explain than one that says which value it took.
    """
    if not preset.value or preset.source == SOURCE_FLAG:
        return preset.value
    if not ask:
        print(f"{label}: dùng {preset.value} theo cấu hình sẵn có.")
        return preset.value
    return confirm_preset(preset, label, input_fn=input_fn)


# --- pickers ---------------------------------------------------------------
#
# Only two live here.  The CSV picker is utils.input_with_completion, which
# already numbers the directory, takes 0 to cancel, completes on Tab and expands
# ~ and $VAR; core.py calls it in about sixty places and a second one would be a
# second thing to keep consistent.


def pick_csv_path(verbose: bool = False) -> str:
    """Ask for the registration CSV using the picker the rest of the tool uses."""
    from .utils import input_with_completion

    return input_with_completion(
        "File CSV đăng ký (Tab để hoàn thành, Enter để liệt kê, 0 để huỷ): ",
        select_file=True,
        file_filter=lambda f: f.lower().endswith(".csv"),
        verbose=verbose,
    )


def pick_google_course(
    credentials_path: str,
    token_path: str,
    verbose: bool = False,
    open_browser: Optional[bool] = None,
) -> Optional[str]:
    """Course id, or ``None`` to skip Google Classroom entirely.

    Thin on purpose: ``_select_course_interactively`` already lists the courses
    and already returns ``None`` on 'q', which is exactly "skip".  All this adds
    is saying so before the prompt, and turning a failure to reach Google into
    the same skip -- a machine with no credentials, or an API that is down, must
    not take the Classroom50 lane down with it.
    """
    from .gclass_invite import _select_course_interactively

    print("Chọn khoá Google Classroom (gõ 'q' để bỏ qua Google Classroom):")
    try:
        return _select_course_interactively(
            credentials_path, token_path, verbose, open_browser=open_browser
        )
    except Exception as exc:  # credentials missing, API down, token expired
        print(f"Không đọc được danh sách khoá Google Classroom ({exc}); bỏ qua Google Classroom.")
        return None


def _classroom_label(entry: Any) -> Tuple[str, str]:
    """(identifier, display name) from one `classroom list --json` entry."""
    if not isinstance(entry, dict):
        return str(entry or "").strip(), str(entry or "").strip()
    identifier = ""
    for key in ("slug", "short_name", "shortName", "name", "id"):
        value = str(entry.get(key, "") or "").strip()
        if value:
            identifier = value
            break
    display = ""
    for key in ("name", "title", "display_name", "displayName"):
        value = str(entry.get(key, "") or "").strip()
        if value:
            display = value
            break
    return identifier, display or identifier


def pick_c50_classroom(
    org: str,
    *,
    runner: Optional[Runner] = None,
    cli: Optional[Any] = None,
    input_fn: Callable[[str], str] = input,
    verbose: bool = False,
) -> Optional[str]:
    """Classroom short-name, or ``None`` to skip Classroom50 entirely."""
    from .c50_ops import list_classrooms

    try:
        payload = list_classrooms(org, cli=cli, runner=runner)
    except Classroom50Error as exc:
        print(f"Không đọc được danh sách lớp Classroom50 ({exc}); bỏ qua Classroom50.")
        return None

    entries = payload if isinstance(payload, list) else []
    rows = [_classroom_label(entry) for entry in entries]
    rows = [row for row in rows if row[0]]
    if not rows:
        print(f"Không có lớp Classroom50 nào trong org {org!r}; bỏ qua Classroom50.")
        return None

    print(f"Lớp Classroom50 trong org {org}:")
    for index, (identifier, display) in enumerate(rows, 1):
        suffix = f" ({identifier})" if display != identifier else ""
        print(f"{index}. {display}{suffix}")
    print("0. Bỏ qua Classroom50")
    while True:
        selection = str(input_fn("Chọn số lớp (0 để bỏ qua): ") or "").strip()
        if selection in ("0", "q", "Q"):
            return None
        if selection.isdigit() and 1 <= int(selection) <= len(rows):
            return rows[int(selection) - 1][0]
        print("Lựa chọn không hợp lệ.")


# --- platform state --------------------------------------------------------

GC_ENROLLED = "đã ghi danh"
GC_TEACHER = "đang là giáo viên"
GC_PENDING = "lời mời đang treo"
GC_WILL_INVITE = "SẼ MỜI"
GC_NO_EMAIL = "không có email dùng được"

C50_ON_ROSTER_MEMBER = "có roster, đã vào org"
C50_ON_ROSTER_NOT_MEMBER = "có roster, CHƯA vào org"
C50_ON_ROSTER_UNKNOWN_MEMBER = "có roster, chưa kiểm được org"
C50_UNLINKED_EMAIL = "đã có (dòng email chưa liên kết)"
C50_WILL_ADD = "SẼ THÊM + MỜI"
C50_NO_USERNAME = "bỏ qua: không có GitHub Username"


class GoogleState(NamedTuple):
    enrolled_emails: Set[str]
    enrolled_user_ids: Set[str]
    teacher_emails: Set[str]
    pending_emails: Set[str]
    pending_user_ids: Set[str]


class Classroom50State(NamedTuple):
    """What the two independent Classroom50 reads found.

    ``roster_usernames`` and ``members`` answer different questions and neither
    substitutes for the other: `roster import` commits the roster first and
    invites afterwards, so a run that died between the two leaves a matching
    roster and nobody invited.  ``unlinked_emails`` holds the roster rows that
    carry an address and no username -- an accepted email invitation that no
    `roster sync` has recorded yet.  Those rows are invisible to the importer's
    own diff, so a candidate matching one here would otherwise be imported as a
    second row for the same person.
    """

    roster_usernames: Set[str]
    unlinked_emails: Set[str]
    members: Set[str]
    members_known: bool
    sync_exit: Optional[int]
    sync_detail: str


def read_google_state(
    service: Any,
    course_id: str,
    *,
    known_google_ids: Optional[Set[str]] = None,
    profile_lookup_limit: int = 200,
    verbose: bool = False,
) -> GoogleState:
    """Read the three membership tiers without writing anything.

    Calls ``gclass_invite``'s own readers rather than its invite function: that
    function runs the numbered selection and the confirmation *before* its
    dry-run branch, so using it as a preview would ask the operator to confirm
    twice.  The readers are the same three API calls it makes.
    """
    from .gclass_invite import _read_course_membership, _read_pending_invitations

    membership = _read_course_membership(service, course_id, verbose)
    pending_emails, pending_user_ids = _read_pending_invitations(
        service, course_id, set(known_google_ids or ()), profile_lookup_limit, verbose
    )
    return GoogleState(
        enrolled_emails=set(membership.emails),
        enrolled_user_ids=set(membership.user_ids),
        teacher_emails=set(membership.teacher_emails),
        pending_emails=set(pending_emails),
        pending_user_ids=set(pending_user_ids),
    )


def confirm_pending_by_email(service: Any, course_id: str, emails: Sequence[str]) -> Set[str]:
    """Addresses that really do hold a pending invitation on this course.

    ``_read_pending_invitations`` resolves invitations to addresses through a
    per-invitation profile lookup capped at ``profile_lookup_limit``; past that
    cap a pending invitation is invisible to it, and the real run only discovers
    it as a 409 from ``invitations().create``.  That is fine for inviting and
    wrong for a preview, which would promise to invite someone already invited.
    ``invitations.list`` accepts an email address as ``userId`` alongside
    ``courseId``, so this asks the exact question for the shortlist only.
    """
    confirmed: Set[str] = set()
    for email in emails:
        if not email:
            continue
        try:
            response = service.invitations().list(
                courseId=str(course_id), userId=email
            ).execute() or {}
        except Exception:
            # A failed lookup must not be read as "no invitation pending"; the
            # 409 backstop in the invite path still covers it.
            continue
        if response.get("invitations"):
            confirmed.add(email.lower())
    return confirmed


def read_classroom50_state(
    org: str,
    classroom: str,
    *,
    runner: Optional[Runner] = None,
    cli: Optional[Any] = None,
    verbose: bool = False,
) -> Classroom50State:
    """Run the roster read, the membership read, and the report-only sync."""
    from .c50_cli_human import HumanCLI
    from .c50_ops import existing_member_logins
    from .c50_roster import parse_roster_payload

    cli = cli or HumanCLI(runner=runner)

    roster_usernames: Set[str] = set()
    unlinked_emails: Set[str] = set()
    for row in parse_roster_payload(cli.roster_list(org, classroom)):
        username = str(row.get("username") or "").strip().lstrip("@").lower()
        email = str(row.get("email") or "").strip().lower()
        if username:
            roster_usernames.add(username)
        elif email:
            unlinked_emails.add(email)

    members: Set[str] = set()
    members_known = True
    try:
        members = existing_member_logins(cli.member_list(org))
    except Classroom50Error as exc:
        members_known = False
        if verbose:
            print(f"[Onboard] không đọc được thành viên org: {exc}")

    result = cli.roster_sync(org, classroom)
    sync_exit = int(getattr(result, "returncode", 1))
    sync_detail = (getattr(result, "stdout", "") or getattr(result, "stderr", "") or "").strip()
    return Classroom50State(
        roster_usernames=roster_usernames,
        unlinked_emails=unlinked_emails,
        members=members,
        members_known=members_known,
        sync_exit=sync_exit,
        sync_detail=sync_detail,
    )


# --- comparison table ------------------------------------------------------


class RowPlan(NamedTuple):
    row: ResolvedRow
    google: str
    classroom50: str


def classify_rows(
    rows: Sequence[ResolvedRow],
    google: Optional[GoogleState],
    c50: Optional[Classroom50State],
    *,
    pending_confirmed: Optional[Set[str]] = None,
) -> List[RowPlan]:
    """Say, per student, what each platform would do -- before anything is done."""
    pending_confirmed = pending_confirmed or set()
    plans: List[RowPlan] = []
    for row in rows:
        plans.append(
            RowPlan(
                row,
                _google_label(row, google, pending_confirmed),
                _c50_label(row, c50),
            )
        )
    return plans


def _google_label(
    row: ResolvedRow, google: Optional[GoogleState], pending_confirmed: Set[str]
) -> str:
    if google is None:
        return "-"
    if not row.email:
        return GC_NO_EMAIL
    email = row.email.lower()
    if email in google.teacher_emails:
        return GC_TEACHER
    if email in google.enrolled_emails:
        return GC_ENROLLED
    if email in google.pending_emails or email in pending_confirmed:
        return GC_PENDING
    return GC_WILL_INVITE


def _c50_label(row: ResolvedRow, c50: Optional[Classroom50State]) -> str:
    if c50 is None:
        return "-"
    if not row.github_username:
        return C50_NO_USERNAME
    username = row.github_username.lower()
    if username in c50.roster_usernames:
        if not c50.members_known:
            return C50_ON_ROSTER_UNKNOWN_MEMBER
        return C50_ON_ROSTER_MEMBER if username in c50.members else C50_ON_ROSTER_NOT_MEMBER
    if row.email and row.email.lower() in c50.unlinked_emails:
        return C50_UNLINKED_EMAIL
    return C50_WILL_ADD


def render_table(plans: Sequence[RowPlan], rejected: Sequence[RejectedRow]) -> str:
    """One table showing every row, including the ones that will not be sent."""
    header = ("Sinh viên", "Google Classroom", "Classroom50")
    body: List[Tuple[str, str, str]] = [
        (plan.row.name or plan.row.student_id or "?", plan.google, plan.classroom50)
        for plan in plans
    ]
    body.extend(
        (item.name or item.student_id or "?", f"bỏ qua: {item.reason}", f"bỏ qua: {item.reason}")
        for item in rejected
    )
    widths = [
        max(len(header[column]), *(len(line[column]) for line in body)) if body else len(header[column])
        for column in range(3)
    ]
    rule = "-+-".join("-" * width for width in widths)
    lines = [" | ".join(header[i].ljust(widths[i]) for i in range(3)), rule]
    lines.extend(" | ".join(line[i].ljust(widths[i]) for i in range(3)) for line in body)
    return "\n".join(lines)


# --- driver ----------------------------------------------------------------


def _student_keys(student: Any) -> Tuple[str, str, str]:
    def field(name: str) -> str:
        return str(getattr(student, name, "") or "").strip()

    return (
        field("Student ID").lower(),
        field("GitHub Username").lstrip("@").lower(),
        field("Email").lower(),
    )


def merge_into_database(students: List[Any], rows: Sequence[ResolvedRow]) -> Dict[str, int]:
    """Apply the resolved rows to the student list in place.

    Deliberately not left to ``save_database``'s ``_dedup_students``: that merge
    is fill-only, so a ``GitHub ID`` already in the database survives even when
    it is stale.  ``gh teacher roster import`` re-resolves the id from the
    username and fails the line when the two disagree, which would drop exactly
    the student whose account was renamed.  A freshly verified id therefore
    overwrites; an unverified one is left empty, because empty is re-resolved
    and wrong is rejected.
    """
    from .models import Student

    added = 0
    updated = 0
    for row in rows:
        match = None
        for student in students:
            sid, username, email = _student_keys(student)
            if row.student_id and sid and sid == row.student_id.lower():
                match = student
                break
            if row.github_username and username and username == row.github_username.lower():
                match = student
                break
            if row.email and email and email == row.email:
                match = student
                break
        if match is None:
            match = Student(**{"Student ID": row.student_id})
            students.append(match)
            added += 1
        else:
            updated += 1
        for attribute, value in (
            ("Student ID", row.student_id),
            ("Name", row.name),
            ("Email", row.email),
            ("GitHub Username", row.github_username),
            ("Section", row.section),
            ("Class", row.class_name),
        ):
            if value:
                setattr(match, attribute, value)
        # Written unconditionally, empty included: see the docstring.
        setattr(match, "GitHub ID", row.github_id)
    return {"added": added, "updated": updated}


def run_onboarding(
    *,
    db_path: str,
    csv_path: Optional[str] = None,
    google_course_id: Setting = None,
    c50_org: Setting = None,
    c50_classroom: Setting = None,
    skip_github_check: bool = False,
    dry_run: bool = False,
    verbose: bool = False,
    report_path: Optional[str] = None,
    credentials_path: str = "gclassroom_credentials.json",
    token_path: str = "token.pickle",
    input_fn: Callable[[str], str] = input,
    runner: Optional[Runner] = None,
    service: Any = None,
    csv_picker: Optional[Callable[[], str]] = None,
    tty_check: Optional[Callable[[], bool]] = None,
    open_browser: Optional[bool] = None,
) -> Dict[str, Any]:
    """Pick the classes, read the CSV, show what would change, then change it.

    Nothing before the confirmation writes: the database save, the Google
    Classroom invitations and the Classroom50 import all happen after it, in
    that order.  ``dry_run`` stops before any of it *and* before the platform
    reads, so it works offline with no credentials -- the report it prints is
    the CSV validation, not a platform preview.

    ``google_course_id``, ``c50_org`` and ``c50_classroom`` take a plain string,
    used as given, or a :class:`Resolved` carrying the tier it came from, in
    which case anything but a command-line flag is confirmed before it is used.

    ``open_browser`` reaches the Google authorization flow: ``None`` opens one
    only if this machine has one, and ``False`` forces the printed-URL path, which
    is finished by pasting the failed redirect back at its own prompt.
    """
    from pathlib import Path

    from .roster_csv import RosterCsvError, parse_student_rows, read_csv_text

    report: Dict[str, Any] = {
        "csv": "",
        "dryRun": bool(dry_run),
        "googleCourseId": "",
        "classroom50": {"org": "", "classroom": ""},
        "read": 0,
        "resolved": 0,
        "rejected": [],
        "duplicates": [],
        "unverifiedGithub": [],
        "databaseSaved": False,
        "google": {"status": "skipped"},
        "c50": {"status": "skipped"},
    }

    # 1. Which classes, and from which file.  A preset that was not typed for
    #    this run is offered as a default and confirmed; declining it leads to
    #    the same picker as supplying nothing.
    if tty_check is None:
        from .c50_admin_cli import _default_tty_check

        tty_check = _default_tty_check
    ask = not dry_run and tty_check()

    google = _as_resolved(google_course_id)
    org = _as_resolved(c50_org)
    room = _as_resolved(c50_classroom)

    course_id = _accept_preset(google, "Khoá Google Classroom", ask=ask, input_fn=input_fn)
    if not course_id and not dry_run:
        course_id = pick_google_course(
            credentials_path, token_path, verbose, open_browser=open_browser
        )
    if not course_id:
        print("Bỏ qua Google Classroom.")
    report["googleCourseId"] = course_id or ""

    # There is no org picker -- a declined org has nothing to fall through to,
    # so it asks for a replacement and takes an empty answer as "skip".
    c50_org = _accept_preset(org, "Org Classroom50", ask=ask, input_fn=input_fn)
    if org.value and not c50_org and ask:
        c50_org = str(input_fn("Org GitHub khác (Enter để bỏ qua Classroom50): ") or "").strip()

    classroom = None
    if c50_org:
        # A classroom configured for one org is no default for another, so a
        # replaced org sends the classroom to the picker.  A classroom named on
        # the command line stays an instruction either way.
        if room.source == SOURCE_FLAG or c50_org == org.value:
            classroom = _accept_preset(room, "Lớp Classroom50", ask=ask, input_fn=input_fn)
        if not classroom and not dry_run:
            classroom = pick_c50_classroom(
                c50_org, runner=runner, input_fn=input_fn, verbose=verbose
            )
    if not (c50_org and classroom):
        print("Bỏ qua Classroom50.")
        classroom = None
    report["classroom50"]["org"] = c50_org or ""
    report["classroom50"]["classroom"] = classroom or ""

    path = csv_path or (csv_picker or pick_csv_path)()
    if not path:
        print("Không chọn file CSV; dừng.")
        return report
    report["csv"] = str(path)

    # 2. Read and resolve, before any platform is touched.
    try:
        text, encoding = read_csv_text(Path(path), verbose=verbose)
        parsed = parse_student_rows(text)
    except RosterCsvError as exc:
        print(f"Không đọc được CSV: {exc}")
        report["error"] = str(exc)
        return report
    report["read"] = len(parsed.rows)
    report["encoding"] = encoding
    report["skippedIncomplete"] = parsed.skipped

    candidates = []
    for row in parsed.rows:
        username, _ = validate_github_username(row.get("GitHub Username", ""))
        if username:
            candidates.append(username)
    checks = check_github_exists(
        candidates,
        runner=runner,
        skip=skip_github_check or dry_run,
        verbose=verbose,
    )
    resolved = resolve_rows(parsed, checks)
    report["resolved"] = len(resolved.ok)
    report["rejected"] = [
        {"studentId": r.student_id, "name": r.name, "reason": r.reason} for r in resolved.rejected
    ]
    report["duplicates"] = list(resolved.duplicates)
    report["unverifiedGithub"] = [
        check.username for check in checks.values() if check.state == GITHUB_UNVERIFIED
    ]

    if dry_run:
        print(_render_report(report, resolved))
        print("Dry run: không ghi cơ sở dữ liệu, không gọi nền tảng nào.")
        _write_report(report, report_path)
        return report

    from .data import load_database, save_database

    students = list(load_database(db_path, verbose=verbose) or [])

    # 3. Pull the Classroom50 roster into the database first, so a student
    #    already on the roster under an identifier the form did not repeat is
    #    matched rather than added again.
    c50_state: Optional[Classroom50State] = None
    if classroom:
        from .c50_ops import sync as c50_sync_pull

        try:
            students, sync_report, _ = c50_sync_pull(
                students, org=c50_org, classroom=classroom, runner=runner
            )
            report["c50"]["pull"] = sync_report
        except Classroom50Error as exc:
            print(f"Không kéo được roster Classroom50 ({exc}); dừng phía Classroom50.")
            report["c50"] = {"status": "read_failed", "detail": str(exc)}
            classroom = None

    # 4./5. Read both platforms and show the table.
    if classroom:
        try:
            c50_state = read_classroom50_state(
                c50_org, classroom, runner=runner, verbose=verbose
            )
        except Classroom50Error as exc:
            print(f"Không đọc được trạng thái Classroom50 ({exc}); dừng phía Classroom50.")
            report["c50"] = {"status": "read_failed", "detail": str(exc)}
            classroom = None

    if c50_state is not None and c50_state.sync_exit != 0:
        # terraform-plan exit codes: 2 is "changes pending", not a failure. Both
        # 1 and 2 stop the import, for different reasons -- 2 because importing
        # now would add a second roster row for a student whose accepted email
        # invitation has not been linked to their account yet.
        detail = _sync_message(c50_state)
        print(detail)
        report["c50"] = {
            "status": "sync_pending" if c50_state.sync_exit == 2 else "sync_failed",
            "exitCode": c50_state.sync_exit,
            "detail": detail,
        }
        classroom = None

    google_state: Optional[GoogleState] = None
    if course_id:
        if service is None:
            from googleapiclient.discovery import build

            from .gclass_auth import _get_google_classroom_credentials

            service = build(
                "classroom",
                "v1",
                credentials=_get_google_classroom_credentials(
                    credentials_path, token_path, verbose=verbose, open_browser=open_browser
                ),
            )
        known_ids = {
            str(getattr(s, "Google_ID", "") or "").strip()
            for s in students
            if str(getattr(s, "Google_ID", "") or "").strip()
        }
        google_state = read_google_state(
            service, course_id, known_google_ids=known_ids, verbose=verbose
        )

    plans = classify_rows(resolved.ok, google_state, c50_state)
    if google_state is not None:
        shortlist = [p.row.email for p in plans if p.google == GC_WILL_INVITE and p.row.email]
        confirmed = confirm_pending_by_email(service, course_id, shortlist)
        if confirmed:
            plans = classify_rows(
                resolved.ok, google_state, c50_state, pending_confirmed=confirmed
            )

    print(_render_report(report, resolved))
    print()
    print(render_table(plans, resolved.rejected))
    print()

    if not resolved.ok:
        print("Không có dòng nào dùng được; không ghi gì.")
        _write_report(report, report_path)
        return report

    answer = str(input_fn("Tiếp tục ghi cơ sở dữ liệu và thêm vào các nền tảng đã chọn? (y/n): ") or "").strip().lower()
    if answer not in ("y", "yes"):
        print("Đã huỷ; không ghi gì.")
        report["cancelled"] = True
        _write_report(report, report_path)
        return report

    # 6. Write: database, then Google Classroom, then Classroom50.
    report["database"] = merge_into_database(students, resolved.ok)
    save_database(students, db_path, verbose=verbose, audit_source="onboard")
    report["databaseSaved"] = True

    if course_id:
        from .gclass_invite import invite_students_to_google_classroom

        report["google"] = invite_students_to_google_classroom(
            course_id=course_id,
            db_path=db_path,
            credentials_path=credentials_path,
            token_path=token_path,
            service=service,
            verbose=verbose,
        )

    if classroom:
        from .c50_admin_cli import main as c50_admin_main

        code = c50_admin_main(
            [
                "roster-import",
                "--org",
                c50_org,
                "--classroom",
                classroom,
                "--db",
                db_path,
            ]
            + (["--verbose"] if verbose else []),
            runner=runner,
            input_fn=input_fn,
            tty_check=tty_check,
        )
        report["c50"] = {"status": "imported" if code == 0 else "failed", "exitCode": code}

    _write_report(report, report_path)
    return report


def _sync_message(state: Classroom50State) -> str:
    if state.sync_exit == 2:
        return (
            "`gh teacher roster sync` báo còn liên kết đang chờ (exit 2 — đây không phải lỗi). "
            "Sinh viên đã nhận lời mời qua email nhưng chưa được ghi nhận tài khoản vẫn nằm trên "
            "roster dưới dạng dòng chỉ có email; import bây giờ sẽ tạo dòng thứ hai cho cùng một "
            "người. Chạy `gh teacher roster sync <org> <classroom> --write` rồi chạy lại.\n"
            + (state.sync_detail or "")
        ).strip()
    return (
        f"`gh teacher roster sync` thất bại (exit {state.sync_exit}); dừng phía Classroom50.\n"
        + (state.sync_detail or "")
    ).strip()


def _render_report(report: Dict[str, Any], resolved: ResolveResult) -> str:
    lines = [
        f"Đọc {report['read']} dòng từ {report['csv']} "
        f"({report.get('encoding', '?')}; bỏ {report.get('skippedIncomplete', 0)} dòng thiếu cột bắt buộc).",
        f"Dùng được: {len(resolved.ok)}. Bị loại: {len(resolved.rejected)}.",
    ]
    for item in resolved.rejected:
        lines.append(f"  - loại: {item.name or '?'} ({item.student_id or 'không có MSSV'}): {item.reason}")
    for note in resolved.duplicates:
        lines.append(f"  - trùng: {note}")
    school = sum(1 for r in resolved.ok if r.email_source == EMAIL_SOURCE_SCHOOL)
    personal = sum(1 for r in resolved.ok if r.email_source == EMAIL_SOURCE_PERSONAL)
    lines.append(f"Email: {school} dùng cột VNU-HUS, {personal} lùi về cột cá nhân.")
    # Kept rows without an address: Classroom50 accepts an empty email per row,
    # so these still reach the roster, but Google Classroom invites by address
    # and will never see them.  Naming them is the only way the operator knows
    # whose form to fix.
    addressless = [row for row in resolved.ok if not row.email]
    if addressless:
        lines.append(
            f"Không có email dùng được ({len(addressless)}); "
            "vẫn lên roster Classroom50 nhưng không mời được lên Google Classroom:"
        )
        for row in addressless:
            lines.append(f"  - {row.name or '?'} ({row.student_id or 'không có MSSV'})")
    unverified = report.get("unverifiedGithub") or []
    if unverified:
        lines.append(f"GitHub chưa kiểm được ({len(unverified)}): {', '.join(unverified)}")
    return "\n".join(lines)


def _write_report(report: Dict[str, Any], path: Optional[str]) -> None:
    if not path:
        return
    import json
    from pathlib import Path

    Path(path).write_text(
        json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
