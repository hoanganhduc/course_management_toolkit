# -*- coding: utf-8 -*-
"""Audit the merged roster and say, per student, exactly what is wrong.

The database this reads is already the join of two sources: the registration
form supplies ``Student ID``, ``GitHub Username``, ``Section`` and ``Class``,
and the Google Classroom sync supplies ``Google_ID`` and the account address.
A student who appears with a ``Google_ID`` but no ``Student ID`` is therefore in
the course but has not answered the form, and that asymmetry is what makes the
whole audit readable off one table.

Every check answers a question a teacher actually asks before term starts: who
has not registered, who typed something that will break a platform operation,
and who looks like the same person twice.  Each finding names the field, quotes
what was entered, and says what to do about it, because the output is meant to
be pasted into a message to that student.

Nothing here writes.  ``data`` is imported lazily inside the one function that
loads a database, so the module stays importable -- and testable -- without
pandas.
"""

from __future__ import annotations

import re
import unicodedata
from collections import Counter, defaultdict
from typing import Any, Dict, List, Mapping, NamedTuple, Optional, Sequence, Tuple

from .c50_cli import Classroom50Error, Runner, _default_runner
from .onboard import (
    GITHUB_MISSING,
    GITHUB_UNVERIFIED,
    GithubCheck,
    _looks_like_not_found,
    validate_email,
    validate_github_username,
)

SEVERITY_ERROR = "error"
SEVERITY_WARNING = "warning"

# Where a GitHub account stands with the organization.  ``MEMBER_ABSENT`` is the
# only one that is work: an invitation was never sent, or was cancelled.
MEMBER_ACTIVE = "member"
MEMBER_INVITED = "invited"
MEMBER_ABSENT = "absent"
MEMBER_UNVERIFIED = "unverified"

# The findings a student has to act on, and the only ones the announcement
# repeats.  A code earns a place here by having a consequence: it stops the
# student being identified, or it stops their GitHub account reaching the
# roster, or it lets one student overwrite another.
#
# The rest stay in the audit for the teacher but are left out of the post.  An
# address that is a personal Gmail, or that does not match the student number,
# changes nothing for a student who is already on Google Classroom: the course
# reaches them either way, and the toolkit identifies them by student number and
# GitHub username, not by address.  ``github_unverified`` is a rate limit, which
# no student can fix.  ``class_format`` is the spelling of a class code, and the
# section is a roster column the teacher can fill in without asking anyone.
# The group codes below are consequential for the same reason: each one names a
# state in which somebody's mark does not reach them.  A student on two group
# repositories is credited on neither reliably, a group over size is not
# collected as one group, a name missing from ``member_usernames`` is a name the
# collector did not pay, and a ``team.json`` that does not parse is the only
# place a student number is recorded.  Every one of them is fixed by the students
# themselves, which is what puts them in a notice rather than in a staff report.
CONSEQUENTIAL_CODES = frozenset(
    {
        "student_id_format",
        "student_id_invalid",
        "student_id_shared",
        "github_missing",
        "github_syntax",
        "github_not_found",
        "github_shared",
        "group_multi_repo",
        "group_over_size",
        "group_member_dropped",
        "group_credit_owner_only",
        "student_id_shared_groups",
        "founder_mismatch",
        "course_mismatch",
        "team_json_invalid",
        "team_json_keys",
        "team_json_course",
        "team_json_member_keys",
        "team_json_size",
        "team_json_username",
        "team_json_placeholder",
        "team_json_duplicate_member",
        "team_json_duplicate_student_id",
        "team_json_founder_absent",
        "team_json_extra_member",
        "team_json_missing_member",
        "issue_member_shared",
        "issue_group_mismatch",
        "issue_founder_mismatch",
        "issue_repo_unknown",
    }
)

# The school address every VNU-HUS student is issued.  Kept as a default rather
# than a constant so the same audit runs for another institution.
DEFAULT_SCHOOL_DOMAIN = "hus.edu.vn"

# A VNU-HUS student number: eight digits, and the local part of the address the
# school issues.  Used only to *compare* the two when both look like a number;
# a name-based school address is normal and is never reported as a mismatch.
_STUDENT_ID_RE = re.compile(r"^\d{8}$")
_CLASS_CODE_RE = re.compile(r"\bK(\d{2})([A-Za-z])(\d)\b", re.IGNORECASE)


class StudentRef(NamedTuple):
    """Just enough of a student to identify them in a report or a message."""

    student_id: str
    name: str
    email: str
    google_id: str


class Issue(NamedTuple):
    """One thing wrong with one student's record.

    ``found`` quotes the value as entered so the student recognises it, and
    ``fix`` is the instruction to send them; a finding with no actionable fix is
    not worth reporting.
    """

    code: str
    severity: str
    student: StudentRef
    field: str
    found: str
    detail: str
    fix: str


class DuplicateGroup(NamedTuple):
    """Records that look like one person holding two Classroom accounts."""

    key: str
    reason: str
    members: List[StudentRef]


class MembershipCheck(NamedTuple):
    username: str
    state: str
    detail: str


class AuditReport(NamedTuple):
    """Counts are per student, not per record; see :func:`group_by_person`."""

    total: int
    with_form: int
    missing_form: List[StudentRef]
    issues: List[Issue]
    duplicate_accounts: List[DuplicateGroup]


def _field(student: Any, key: str) -> str:
    """Read one attribute as a stripped string, whatever the record type is."""
    if isinstance(student, dict):
        return str(student.get(key, "") or "").strip()
    return str(getattr(student, key, "") or "").strip()


def student_ref(student: Any) -> StudentRef:
    return StudentRef(
        student_id=_field(student, "Student ID"),
        name=_field(student, "Name"),
        email=_field(student, "Email"),
        google_id=_field(student, "Google_ID"),
    )


def fold_name(value: str) -> str:
    """Fold a name to a comparison key: no accents, no case, no word order.

    Google Classroom shows ``NGOC TRAN TRONG`` where the form says ``Trần Trọng Ngọc``;
    the two are the same student, so the key has to survive both the missing
    accents and the reversed order that the school's account names use.
    """
    text = unicodedata.normalize("NFD", str(value or ""))
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.replace("đ", "d").replace("Đ", "D")
    return " ".join(sorted(re.sub(r"[^a-z ]", " ", text.lower()).split()))


def has_form_response(student: Any) -> bool:
    """Whether this student answered the registration form.

    ``Student ID`` is the test because the form is its only source: the
    Classroom sync never supplies one.  A number the import rejected is kept
    under ``Invalid Student ID`` and counts just as much -- the student did
    answer, and telling them they never filled the form when the fault is one
    extra digit sends them looking for the wrong mistake.
    """
    return bool(_field(student, "Student ID") or _field(student, "Invalid Student ID"))


def identity_keys(student: Any) -> set:
    """The keys that identify the person behind one record.

    Two records belong to one student when they share any of these.  One
    ``Google_ID`` is one account and so one person, whatever the row says their
    name is; the sync writes it into the form reply it matched, so it links the
    two sources as soon as they have been joined once.  Before that the join
    rests on the columns both sources fill: the addresses, plus the student
    number, which the form supplies directly and a school address carries in its
    local part.
    """
    keys = set()
    google_id = _field(student, "Google_ID")
    if google_id:
        keys.add(f"gid:{google_id}")
    student_id = _field(student, "Student ID").lower()
    if student_id:
        keys.add(f"id:{student_id}")
    for attribute in ("Email", "Google_Classroom_Email"):
        address = _field(student, attribute).lower()
        if not address:
            continue
        keys.add(f"mail:{address}")
        local = address.split("@")[0]
        if _STUDENT_ID_RE.match(local):
            keys.add(f"id:{local}")
    return keys


def group_by_person(students: Sequence[Any]) -> List[List[Any]]:
    """Split the roster into one group per student, in first-seen order.

    Records are merged transitively: a shell that shares an address with a form
    reply, and a form reply that shares a student number with a third record,
    all land in the same group.  A record with no key at all stands alone, since
    nothing about it can be matched to anyone.
    """
    parent: Dict[int, int] = {index: index for index in range(len(students))}

    def root(index: int) -> int:
        while parent[index] != index:
            parent[index] = parent[parent[index]]
            index = parent[index]
        return index

    first_seen: Dict[str, int] = {}
    for index, student in enumerate(students):
        for key in identity_keys(student):
            if key in first_seen:
                left, right = root(first_seen[key]), root(index)
                if left != right:
                    parent[right] = left
            else:
                first_seen[key] = index

    groups: Dict[int, List[Any]] = defaultdict(list)
    for index, student in enumerate(students):
        groups[root(index)].append(student)
    return [groups[key] for key in sorted(groups)]


def list_missing_form(students: Sequence[Any]) -> List[StudentRef]:
    """Students enrolled in the course who have not answered the form.

    One student routinely holds two records -- the Classroom sync writes one
    under the unaccented display name with no student number, and the form
    writes another -- so counting records here would send a reminder to people
    who already answered.  The count is therefore per person.
    """
    return [
        student_ref(group[0])
        for group in group_by_person(students)
        if not any(has_form_response(s) for s in group)
    ]


def _class_code(raw: str) -> str:
    """The canonical class code inside a free-text answer, or ``""``."""
    match = _CLASS_CODE_RE.search(raw or "")
    if not match:
        return ""
    return f"K{match.group(1)}{match.group(2).upper()}{match.group(3)}"


def _invalid_id_detail(value: str) -> str:
    """Say what is wrong with a rejected student number, not merely that it is.

    A student who is told "sai mã sinh viên" reads it as a claim they mistyped
    the whole thing and often sends back the same number again; told they wrote
    nine digits where eight belong, they find the extra one themselves.
    """
    text = value.strip()
    digits = re.sub(r"\D", "", text)
    if len(digits) != len(text):
        return "mã sinh viên có ký tự không phải chữ số; mã đúng là 8 chữ số"
    if len(digits) > 8:
        return (
            f"mã sinh viên có {len(digits)} chữ số, thừa {len(digits) - 8}; "
            "mã đúng phải đủ 8 chữ số"
        )
    if len(digits) < 8:
        return (
            f"mã sinh viên chỉ có {len(digits)} chữ số, thiếu {8 - len(digits)}; "
            "mã đúng phải đủ 8 chữ số"
        )
    return "mã sinh viên không hợp lệ; mã đúng là 8 chữ số"


def _shared_values(
    students: Sequence[Any], key: str, *, casefold: bool = True
) -> Dict[str, List[Any]]:
    """Group records by one field, keeping only the values held by more than one."""
    buckets: Dict[str, List[Any]] = defaultdict(list)
    for student in students:
        value = _field(student, key)
        if value:
            buckets[value.lower() if casefold else value].append(student)
    return {value: rows for value, rows in buckets.items() if len(rows) > 1}


def list_invalid_info(
    students: Sequence[Any],
    *,
    checks: Optional[Dict[str, GithubCheck]] = None,
    school_domain: str = DEFAULT_SCHOOL_DOMAIN,
    expected_sections: Optional[Sequence[str]] = None,
) -> List[Issue]:
    """Every field-level problem in the roster, one Issue per problem.

    Only students who answered the form are checked for form fields; reporting a
    blank GitHub username for someone who never saw the form would bury the
    findings that a student can actually act on.  ``checks`` is the result of
    ``onboard.check_github_exists``; without it the GitHub rule is syntax only,
    and no row is ever accused of a username that merely could not be verified.
    """
    checks = checks or {}
    expected = {s.strip().lower() for s in (expected_sections or []) if s.strip()}
    domain = school_domain.lower().lstrip("@")
    issues: List[Issue] = []

    shared_ids = _shared_values(students, "Student ID")
    shared_github = _shared_values(students, "GitHub Username")

    dominant_class: Dict[str, str] = {}
    class_counts: Dict[str, Counter] = defaultdict(Counter)
    for student in students:
        raw_class = _field(student, "Class")
        code = _class_code(raw_class)
        if code:
            class_counts[code][raw_class] += 1
    for code, counter in class_counts.items():
        dominant_class[code] = counter.most_common(1)[0][0]

    for student in students:
        ref = student_ref(student)
        registered = has_form_response(student)

        def add(code: str, severity: str, field: str, found: str, detail: str, fix: str) -> None:
            issues.append(Issue(code, severity, ref, field, found, detail, fix))

        # -- email -------------------------------------------------------
        raw_email = _field(student, "Email")
        email, email_reason = validate_email(raw_email)
        if not raw_email:
            add(
                "email_missing", SEVERITY_ERROR, "Email", "",
                "không có địa chỉ email nào trong hồ sơ",
                f"điền lại form với email trường @{domain}",
            )
        elif not email:
            add(
                "email_syntax", SEVERITY_ERROR, "Email", raw_email,
                f"email {email_reason}",
                f"điền lại form với email trường @{domain}",
            )
        elif not email.endswith("@" + domain):
            add(
                "email_not_school", SEVERITY_WARNING, "Email", raw_email,
                f"không phải email trường (@{domain})",
                f"điền lại form bằng email @{domain} để nhận thông báo của lớp",
            )
        elif registered and ref.student_id:
            local = email.split("@")[0]
            if _STUDENT_ID_RE.match(local) and local != ref.student_id:
                add(
                    "email_id_mismatch", SEVERITY_WARNING, "Student ID", ref.student_id,
                    f"mã sinh viên khác phần đầu email trường ({local})",
                    "kiểm tra lại mã sinh viên hoặc email trường trong form",
                )

        google_email = _field(student, "Google_Classroom_Email") or (
            raw_email if ref.google_id else ""
        )
        if google_email and not google_email.lower().endswith("@" + domain):
            add(
                "google_account_personal", SEVERITY_WARNING, "Google Classroom",
                google_email,
                f"đang vào Google Classroom bằng tài khoản cá nhân, không phải @{domain}",
                f"rời lớp rồi vào lại bằng tài khoản @{domain}",
            )

        if not registered:
            continue

        # -- student id --------------------------------------------------
        # The import parks a number it could not accept under ``Invalid Student
        # ID`` rather than dropping the whole registration, and leaves the field
        # itself empty.  So the parked value describes the field only while the
        # field is still blank: report it in place of the blank it left, but
        # never over a number that has since arrived and parses, or a student
        # who answers the form again correctly is told to fix what they fixed.
        invalid_id = _field(student, "Invalid Student ID")
        if not _STUDENT_ID_RE.match(ref.student_id):
            if invalid_id:
                add(
                    "student_id_invalid", SEVERITY_ERROR, "Student ID", invalid_id,
                    _invalid_id_detail(invalid_id),
                    "điền lại form với đúng mã sinh viên 8 chữ số",
                )
            else:
                add(
                    "student_id_format", SEVERITY_WARNING, "Student ID", ref.student_id,
                    "mã sinh viên không phải 8 chữ số",
                    "điền lại form với đúng mã sinh viên 8 chữ số",
                )
        if ref.student_id.lower() in shared_ids:
            others = [
                _field(o, "Name")
                for o in shared_ids[ref.student_id.lower()]
                if _field(o, "Name") != ref.name
            ]
            add(
                "student_id_shared", SEVERITY_ERROR, "Student ID", ref.student_id,
                "mã sinh viên trùng với: " + ", ".join(others),
                "một trong hai bạn đã điền nhầm mã sinh viên; cả hai điền lại form",
            )

        # -- github ------------------------------------------------------
        raw_github = _field(student, "GitHub Username")
        github, _github_reason = validate_github_username(raw_github)
        if not raw_github:
            add(
                "github_missing", SEVERITY_ERROR, "GitHub Username", "",
                "chưa có tài khoản GitHub",
                "tạo tài khoản GitHub rồi điền lại form",
            )
        elif not github:
            add(
                "github_syntax", SEVERITY_ERROR, "GitHub Username", raw_github,
                "không phải tên đăng nhập GitHub hợp lệ (không được có dấu cách, "
                "dấu tiếng Việt hay ký tự lạ)",
                "điền đúng tên đăng nhập GitHub (không phải họ tên, không có dấu cách)",
            )
        else:
            if raw_github.lower() in shared_github:
                others = [
                    _field(o, "Name")
                    for o in shared_github[raw_github.lower()]
                    if _field(o, "Name") != ref.name
                ]
                add(
                    "github_shared", SEVERITY_ERROR, "GitHub Username", raw_github,
                    "tài khoản GitHub trùng với: " + ", ".join(others),
                    "mỗi bạn dùng một tài khoản GitHub riêng; điền lại form",
                )
            check = checks.get(github.lower())
            if check is not None and check.state == GITHUB_MISSING:
                add(
                    "github_not_found", SEVERITY_ERROR, "GitHub Username", raw_github,
                    "tài khoản GitHub này không tồn tại",
                    "kiểm tra lại tên đăng nhập tại github.com rồi điền lại form",
                )
            elif check is not None and check.state == GITHUB_UNVERIFIED:
                add(
                    "github_unverified", SEVERITY_WARNING, "GitHub Username", raw_github,
                    f"chưa kiểm tra được trên GitHub ({check.detail})",
                    "bạn không cần làm gì; mình kiểm tra lại sau",
                )

        # -- section and class -------------------------------------------
        section = _field(student, "Section")
        if not section:
            add(
                "section_missing", SEVERITY_ERROR, "Section", "",
                "chưa điền lớp học phần",
                "điền lại form và chọn đúng lớp học phần",
            )
        elif expected and section.lower() not in expected:
            add(
                "section_unknown", SEVERITY_ERROR, "Section", section,
                "lớp học phần không nằm trong danh sách của môn",
                "điền lại form và chọn đúng lớp học phần",
            )

        raw_class = _field(student, "Class")
        code = _class_code(raw_class)
        if code and raw_class != dominant_class.get(code, raw_class):
            add(
                "class_format", SEVERITY_WARNING, "Class", raw_class,
                f"ghi khác cách viết chung của lớp ({dominant_class[code]})",
                f"ghi lớp đúng dạng {code}",
            )

    return issues


def _person_ref(group: Sequence[Any]) -> StudentRef:
    """The record that identifies a person best, for one line of a report.

    The form reply carries the student's own spelling of their name and their
    number; the Classroom shell carries neither, so it is used only when there is
    nothing else.
    """
    for student in group:
        if has_form_response(student):
            return student_ref(student)
    return student_ref(group[0])


def list_duplicate_accounts(
    students: Sequence[Any], *, people: Optional[Sequence[Sequence[Any]]] = None
) -> List[DuplicateGroup]:
    """People who look like one student holding two Google Classroom accounts.

    The comparison is between *people*, never between records.  Two rows that
    already share a ``Google_ID``, a student number or an address are one person
    by construction -- that is what :func:`group_by_person` decided -- and
    reporting them here accuses a student of a second account when all that
    happened is that the sync wrote a second row.  ``people`` is the grouping the
    caller has already computed; without it the grouping is done here.

    Between people, two signals remain, both cheap and both grounded: the same
    name once accents and word order are folded away, and the same address local
    part under different domains -- which is how a student who joined with a
    personal Gmail and later with the school address shows up.
    """
    people = list(people if people is not None else group_by_person(students))
    groups: List[DuplicateGroup] = []
    seen: set = set()

    def members(indexes: Sequence[int]) -> List[StudentRef]:
        return [_person_ref(people[index]) for index in indexes]

    by_name: Dict[str, List[int]] = defaultdict(list)
    for index, person in enumerate(people):
        for key in {fold_name(_field(record, "Name")) for record in person}:
            if key:
                by_name[key].append(index)
    for key, indexes in sorted(by_name.items()):
        if len(indexes) > 1:
            groups.append(
                DuplicateGroup(key, "họ tên giống nhau khi bỏ dấu và không kể thứ tự",
                               members(indexes))
            )
            seen.add(frozenset(indexes))

    by_local: Dict[str, List[int]] = defaultdict(list)
    for index, person in enumerate(people):
        parts = set()
        for record in person:
            for attribute in ("Email", "Google_Classroom_Email"):
                value = _field(record, attribute).lower()
                if value and "@" in value:
                    parts.add(value.split("@")[0])
        for local in parts:
            by_local[local].append(index)
    for local, indexes in sorted(by_local.items()):
        # The same pair under a second signal is the same finding, not another one.
        if len(indexes) > 1 and frozenset(indexes) not in seen:
            groups.append(
                DuplicateGroup(local, "phần đầu địa chỉ email giống nhau",
                               members(indexes))
            )
    return groups


def audit_roster(
    students: Sequence[Any],
    *,
    checks: Optional[Dict[str, GithubCheck]] = None,
    school_domain: str = DEFAULT_SCHOOL_DOMAIN,
    expected_sections: Optional[Sequence[str]] = None,
) -> AuditReport:
    """Run every check and return the whole picture in one value."""
    people = group_by_person(students)
    return AuditReport(
        total=len(people),
        with_form=sum(1 for group in people if any(has_form_response(s) for s in group)),
        missing_form=list_missing_form(students),
        issues=list_invalid_info(
            students,
            checks=checks,
            school_domain=school_domain,
            expected_sections=expected_sections,
        ),
        duplicate_accounts=list_duplicate_accounts(students, people=people),
    )


def roster_github_usernames(students: Sequence[Any]) -> List[str]:
    """The syntactically valid GitHub usernames on the roster, cleaned.

    Both GitHub questions -- does the account exist, and is it in the org --
    ask about the same names, so they read them the same way.
    """
    names: List[str] = []
    for student in students:
        clean, _ = validate_github_username(_field(student, "GitHub Username"))
        if clean:
            names.append(clean)
    return names


def verify_github_accounts(
    students: Sequence[Any], *, skip: bool = False, verbose: bool = False
) -> Dict[str, GithubCheck]:
    """Ask GitHub about every syntactically valid username in the roster."""
    from .onboard import check_github_exists

    return check_github_exists(
        roster_github_usernames(students), skip=skip, verbose=verbose
    )


def _pending_invitations(
    org: str, *, runner: Runner, verbose: bool = False
) -> Optional[set]:
    """Every login with an open invitation to ``org``, or ``None`` if unreadable.

    ``None`` is not "nobody": it means the question was not answered, and the
    caller must not turn a 404 into an accusation on the strength of it.
    """
    argv = ["gh", "api", "--paginate", f"/orgs/{org}/invitations", "--jq", ".[].login"]
    try:
        result = runner(argv)
    except (Classroom50Error, OSError) as exc:
        if verbose:
            print(f"[Audit] invitation list unreadable: {exc}")
        return None
    if getattr(result, "returncode", 1) != 0:
        if verbose:
            stderr = (getattr(result, "stderr", "") or "").strip()
            print(f"[Audit] invitation list unreadable: {stderr}")
        return None
    logins = set()
    for line in (getattr(result, "stdout", "") or "").splitlines():
        login = line.strip().lower()
        # An invitation sent to an address has no account behind it yet, and jq
        # prints the missing field as the literal ``null``.
        if login and login != "null":
            logins.add(login)
    return logins


def verify_org_membership(
    usernames: Sequence[str],
    org: str,
    *,
    runner: Optional[Runner] = None,
    skip: bool = False,
    verbose: bool = False,
) -> Dict[str, MembershipCheck]:
    """Where each username stands with the organization, keyed by lower-cased name.

    ``GET /orgs/{org}/members/{username}`` answers 404 both for someone who was
    never invited and for someone whose invitation is still open, so that call
    alone cannot tell the two apart -- and reading it as absence is how a class
    that was invited correctly gets reported as a class nobody invited.  The
    pending list is read once and consulted for every 404; if it cannot be read,
    no name is called absent.
    """
    wanted: List[str] = []
    for raw in usernames:
        name = str(raw or "").strip()
        if name and name.lower() not in {w.lower() for w in wanted}:
            wanted.append(name)
    if not wanted:
        return {}
    if skip:
        return {
            name.lower(): MembershipCheck(name, MEMBER_UNVERIFIED, "bỏ qua kiểm tra API")
            for name in wanted
        }

    runner = runner or _default_runner
    pending = _pending_invitations(org, runner=runner, verbose=verbose)

    checks: Dict[str, MembershipCheck] = {}
    for name in wanted:
        key = name.lower()
        argv = ["gh", "api", f"/orgs/{org}/members/{name}", "--silent"]
        try:
            result = runner(argv)
        except (Classroom50Error, OSError) as exc:
            checks[key] = MembershipCheck(name, MEMBER_UNVERIFIED, str(exc))
        else:
            stderr = (getattr(result, "stderr", "") or "").strip()
            if getattr(result, "returncode", 1) == 0:
                checks[key] = MembershipCheck(name, MEMBER_ACTIVE, "đã vào org")
            elif pending is not None and key in pending:
                checks[key] = MembershipCheck(
                    name, MEMBER_INVITED, "đã mời, chờ sinh viên bấm nhận"
                )
            elif pending is None:
                # Without the invitation list a 404 proves nothing, so decide nothing.
                checks[key] = MembershipCheck(
                    name, MEMBER_UNVERIFIED, "chưa đọc được danh sách lời mời"
                )
            elif _looks_like_not_found(stderr):
                checks[key] = MembershipCheck(name, MEMBER_ABSENT, "chưa có lời mời nào")
            else:
                # Rate limit, no network, no auth: cannot tell, so do not decide.
                checks[key] = MembershipCheck(
                    name, MEMBER_UNVERIFIED, stderr or "không rõ lỗi"
                )
        if verbose:
            print(f"[Audit] /orgs/{org}/members/{name} -> {checks[key].state}")
    return checks


def load_students(db_path: str, verbose: bool = False) -> List[Any]:
    """Load the roster, importing the heavy data module only when actually used."""
    from .data import load_database

    return list(load_database(db_path, verbose=verbose) or [])


def report_to_dict(report: AuditReport) -> Dict[str, Any]:
    """The report as plain JSON-ready data, for the agent entrypoint."""
    return {
        "total": report.total,
        "with_form": report.with_form,
        "missing_form": [ref._asdict() for ref in report.missing_form],
        "issues": [
            {**issue._asdict(), "student": issue.student._asdict()}
            for issue in report.issues
        ],
        "duplicate_accounts": [
            {
                "key": group.key,
                "reason": group.reason,
                "members": [ref._asdict() for ref in group.members],
            }
            for group in report.duplicate_accounts
        ],
    }


def format_missing_form(refs: Sequence[StudentRef]) -> str:
    if not refs:
        return "Tất cả sinh viên trong lớp đã điền form."
    lines = [f"{len(refs)} sinh viên chưa điền form:"]
    for ref in sorted(refs, key=lambda r: r.name.lower()):
        lines.append(f"  - {ref.name or '(không tên)':<28} {ref.email}")
    return "\n".join(lines)


def format_issues(issues: Sequence[Issue]) -> str:
    """Findings grouped by student, because that is how they get sent out."""
    if not issues:
        return "Không có thông tin nào sai."
    by_student: Dict[Tuple[str, str], List[Issue]] = defaultdict(list)
    for issue in issues:
        by_student[(issue.student.student_id, issue.student.name)].append(issue)
    errors = sum(1 for i in issues if i.severity == SEVERITY_ERROR)
    lines = [
        f"{len(issues)} lỗi trên {len(by_student)} sinh viên "
        f"({errors} phải sửa, {len(issues) - errors} nên sửa):"
    ]
    for (student_id, name), found in sorted(by_student.items(), key=lambda kv: kv[0][1].lower()):
        lines.append(f"\n  {student_id or '--------'}  {name or '(không tên)'}  <{found[0].student.email}>")
        for issue in found:
            mark = "PHẢI SỬA" if issue.severity == SEVERITY_ERROR else "nên sửa"
            quoted = f" (đã điền: {issue.found!r})" if issue.found else ""
            lines.append(f"    [{mark}] {issue.field}: {issue.detail}{quoted}")
            lines.append(f"              -> {issue.fix}")
    return "\n".join(lines)


def format_duplicate_accounts(groups: Sequence[DuplicateGroup]) -> str:
    if not groups:
        return "Không có tài khoản nào nghi bị trùng."
    lines = [f"{len(groups)} nhóm nghi trùng tài khoản:"]
    for group in groups:
        lines.append(f"\n  {group.reason}:")
        for ref in group.members:
            lines.append(
                f"    - {ref.name or '(không tên)':<24} {ref.email:<34} "
                f"MSV={ref.student_id or '(chưa điền form)'}"
            )
    return "\n".join(lines)


_MEMBERSHIP_LABELS: Tuple[Tuple[str, str], ...] = (
    (MEMBER_ACTIVE, "đã vào org"),
    (MEMBER_INVITED, "đang chờ sinh viên nhận lời mời"),
    (MEMBER_ABSENT, "CHƯA ĐƯỢC MỜI"),
    (MEMBER_UNVERIFIED, "chưa kiểm được"),
)


def format_membership(
    checks: Mapping[str, MembershipCheck],
    *,
    not_on_roster: Sequence[str] = (),
    population: str = "roster",
) -> str:
    """Membership grouped by state, uninvited last and in capitals.

    ``checks`` answers one question -- where does this account stand with the
    organization -- and it can only answer it about accounts that reached the
    roster.  ``not_on_roster`` carries the rest: names the database holds that
    ``gh teacher roster import`` has never been given, which GitHub has no
    opinion about at all.  Keeping the two apart is the point: folded together
    they read as one pile of uninvited students, and the fix for each half is a
    different command.

    ``population`` says which list ``checks`` was built from, so the header does
    not claim to describe the live roster when the roster could not be read.
    """
    if not checks and not not_on_roster:
        return "Không có tài khoản GitHub nào trên roster để kiểm tra."
    by_state: Dict[str, List[MembershipCheck]] = defaultdict(list)
    for check in checks.values():
        by_state[check.state].append(check)
    summary = ", ".join(
        f"{len(by_state.get(state, []))} {label.lower()}"
        for state, label in _MEMBERSHIP_LABELS
        if by_state.get(state)
    )
    where = (
        "trên roster Classroom50"
        if population == "roster"
        else "trong database (chưa đọc được roster sống)"
    )
    lines = [f"{len(checks)} tài khoản GitHub {where}: {summary}."]
    for state, label in _MEMBERSHIP_LABELS:
        rows = by_state.get(state, [])
        if not rows:
            continue
        note = " -- cần mời" if state == MEMBER_ABSENT else ""
        lines.append(f"\n  {label} ({len(rows)}){note}:")
        for check in sorted(rows, key=lambda c: c.username.lower()):
            reason = f"  ({check.detail})" if state == MEMBER_UNVERIFIED else ""
            lines.append(f"    - {check.username}{reason}")
    if not_on_roster:
        lines.append(
            f"\n  CHƯA IMPORT LÊN ROSTER ({len(not_on_roster)}) -- cần roster import, "
            "chưa mời được:"
        )
        for username in sorted(set(not_on_roster), key=str.lower):
            lines.append(f"    - {username}")
    return "\n".join(lines)


def roster_usernames_from_payload(payload: Any) -> List[str]:
    """The logins on a live ``gh teacher roster list --json`` payload.

    Rows with no login are roster entries for a student who has not linked a
    GitHub account yet; they are dropped rather than counted, because there is no
    account to ask the organization about.
    """
    from .c50_roster import parse_roster_payload

    names: List[str] = []
    for row in parse_roster_payload(payload):
        username = str(row.get("username") or "").strip()
        if username:
            names.append(username)
    return names


def usernames_not_on_roster(
    db_usernames: Sequence[str], roster_usernames: Sequence[str]
) -> List[str]:
    """Database accounts the roster has never heard of, in database order.

    The database export lower-cases what the student typed and the live roster
    keeps GitHub's own capitalisation, so the comparison has to ignore case or
    every account looks missing.
    """
    known = {str(name).strip().lower() for name in roster_usernames if str(name).strip()}
    missing: List[str] = []
    for raw in db_usernames:
        name = str(raw or "").strip()
        if name and name.lower() not in known:
            missing.append(name)
    return missing


def format_audit(report: AuditReport) -> str:
    """The whole audit as one block of text for a terminal."""
    return "\n\n".join(
        [
            f"Sĩ số: {report.total} sinh viên; {report.with_form} đã điền form, "
            f"{report.total - report.with_form} chưa.",
            format_missing_form(report.missing_form),
            format_issues(report.issues),
            format_duplicate_accounts(report.duplicate_accounts),
        ]
    )


def format_announcement(
    report: AuditReport,
    *,
    course_name: str = "",
    form_url: str = "",
    deadline: str = "",
    school_domain: str = DEFAULT_SCHOOL_DOMAIN,
) -> str:
    """The text to post as a Google Classroom announcement.

    Every finding is named: a student who has to fix something can only act on
    it if they can find themselves in the post, and a teacher who has to chase
    twenty-two people one at a time gains nothing from a post that says only how
    many there are.  Sections with nothing in them are dropped rather than
    printed empty, so the numbering follows what the roster actually contains.

    Only ``CONSEQUENTIAL_CODES`` are repeated here.  The audit reports more than
    that, but a post naming a dozen harmless spelling differences buys nothing
    and costs attention: the few students whose GitHub account or student number
    is actually wrong are the ones who must notice themselves, and they notice a
    short list.  Suspected duplicate accounts are named without an instruction
    to leave the course, because two students can share a name and the wrong
    departure loses the work already submitted under that account.
    """
    domain = school_domain.lower().lstrip("@")
    heading = f"[{course_name}] " if course_name else ""
    missing = report.total - report.with_form
    lines = [
        f"{heading}RÀ SOÁT THÔNG TIN ĐĂNG KÝ",
        "",
        f"Lớp hiện có {report.total} bạn trên Google Classroom: {report.with_form} bạn "
        f"đã điền form đăng ký, {missing} bạn chưa điền.",
        "",
        "Mình đã đối chiếu form với danh sách lớp. Bên dưới chỉ liệt kê những chỗ thật "
        "sự gây ảnh hưởng, có ghi rõ tên. Bạn nào có tên thì điền lại form; bản điền sau "
        "sẽ thay bản điền trước.",
        "",
        "Bạn nào không có tên bên dưới thì không cần làm gì. Riêng chuyện email: bạn đang "
        "vào Google Classroom bằng địa chỉ nào cũng được, kể cả Gmail cá nhân - bạn vẫn "
        "nhận đủ thông báo và bài tập, nên mình không liệt kê ở đây.",
    ]
    if form_url:
        lines += ["", f"Link form: {form_url}"]

    section = 0

    if report.missing_form:
        section += 1
        lines += [
            "",
            f"{section}. CHƯA ĐIỀN FORM ({len(report.missing_form)} bạn)",
            "",
            "Mình chưa có mã sinh viên và tài khoản GitHub của các bạn dưới đây, nên chưa "
            "tạo được kho bài tập trên GitHub cho các bạn. Không có kho bài tập thì các bạn "
            "sẽ không nhận được bài tập.",
            "",
        ]
        for ref in sorted(report.missing_form, key=lambda r: r.name.lower()):
            lines.append(f"  - {ref.name} ({ref.email})")

    announced = [issue for issue in report.issues if issue.code in CONSEQUENTIAL_CODES]
    if announced:
        by_student: Dict[Tuple[str, str, str], List[Issue]] = defaultdict(list)
        for issue in announced:
            key = (issue.student.name, issue.student.student_id, issue.student.email)
            by_student[key].append(issue)
        section += 1
        lines += ["", f"{section}. THÔNG TIN ĐIỀN CHƯA ĐÚNG ({len(by_student)} bạn)"]
        for (name, student_id, email) in sorted(by_student, key=lambda k: k[0].lower()):
            label = f"{name} ({student_id})" if student_id else f"{name} ({email})"
            lines += ["", f"  {label}"]
            for issue in by_student[(name, student_id, email)]:
                entered = f' — bạn điền "{issue.found}"' if issue.found else ""
                lines.append(f"    - {issue.field}: {issue.detail}{entered}")
                lines.append(f"      => {issue.fix}")

    if report.duplicate_accounts:
        section += 1
        lines += [
            "",
            f"{section}. NGHI TRÙNG TÀI KHOẢN",
            "",
            "Các tài khoản dưới đây trông như cùng một người vào lớp hai lần. Chỗ này có "
            "ảnh hưởng thật: điểm và bài nộp bị tách làm hai hồ sơ, và khi mình tạo kho bài "
            "tập thì một trong hai tài khoản sẽ không có kho.",
            "",
            "Đừng tự rời lớp. Cũng có thể đây là hai bạn khác nhau trùng tên, mà rời nhầm "
            "thì mất luôn bài đã nộp. Bạn nào có tên dưới đây nhắn riêng cho mình, nói rõ "
            f"tài khoản nào là của bạn (nên giữ tài khoản @{domain}), mình xử lý phần còn "
            "lại.",
        ]
        for group in report.duplicate_accounts:
            lines += ["", f"  Nghi trùng vì {group.reason}:"]
            for ref in group.members:
                lines.append(f"    - {ref.name} ({ref.email})")

    section += 1
    lines += [
        "",
        f"{section}. CÁCH TỰ KIỂM TRA",
        "",
        "a) Tên đăng nhập GitHub. Phải là tên đăng nhập (username) trên github.com, "
        "không phải họ tên, không có dấu cách và không có dấu tiếng Việt. Cách kiểm tra: "
        "mở github.com, đăng nhập, vào trang cá nhân, địa chỉ trang có dạng "
        "github.com/<tên đăng nhập> - phần sau dấu gạch chéo chính là thứ cần điền. Điền "
        "sai thì mình không tạo được kho bài tập cho bạn.",
        "",
        "b) Mỗi bạn một tài khoản GitHub riêng. Hai bạn điền trùng một tên đăng nhập thì "
        "chỉ một bạn có kho bài tập, bạn còn lại không có.",
        "",
        "c) Mã sinh viên. Đủ 8 chữ số. Đây là thứ mình dùng để ghép điểm, nên sai một chữ "
        "số là ghép nhầm người.",
    ]

    if deadline:
        lines += [
            "",
            f"HẠN CHÓT rà soát và điền lại: {deadline}.",
            "",
            "Sau hạn trên mình sẽ dùng danh sách này để tạo kho bài tập trên GitHub. Bạn nào "
            "còn thiếu hoặc sai thông tin thì nhắn riêng cho mình để xử lý.",
        ]
    else:
        lines += [
            "",
            "Mình sẽ dùng danh sách này để tạo kho bài tập trên GitHub. Bạn nào còn thiếu "
            "hoặc sai thông tin thì nhắn riêng cho mình để xử lý.",
        ]
    return "\n".join(lines)
