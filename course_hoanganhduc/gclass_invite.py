# -*- coding: utf-8 -*-
"""Invite students from the local database to a Google Classroom course.

The flow mirrors :mod:`course_hoanganhduc.gclass_unenroll` (all-keyword
signature, interactive course picker, numbered selection, confirmation before
the dry-run check).  Three deliberate departures from that template:

* ``service`` and ``sleep_fn`` can be injected.  When ``service`` is ``None``
  this module builds its own exactly as the unenroll template does, so the
  injection is a strict superset that exists only to make the operation
  testable offline.
* A summary ``dict`` is returned instead of a bare ``int``.  Removing a student
  has one outcome; inviting one has six, and a caller cannot tell them apart
  from a count.  ``canvas_invites.invite_students_if_not_enrolled`` already
  returns a dict for the same operation.
* ``load_database``/``save_database`` are imported lazily inside the branches
  that need them.  ``data`` pulls pandas/openpyxl/sklearn at module scope, and
  importing it here would force every test of this module to stub them.

Retries: ``gclass_coursework`` pins ``num_retries=0`` because a coursework
create whose outcome is unknown must never be repeated.  Invitations are
different -- ``invitations().create`` is naturally deduplicated by its own 409
ALREADY_EXISTS -- so a bounded retry on transient statuses is safe here.  Do
not "fix" this to match the coursework module.

Dry-run placement: the check sits *after* the confirmation, as in the unenroll
template.  ``c50_admin_cli`` places it *before* its confirmation.  Both halves
are internally consistent with their own surface; unifying them would change a
shipped convention for no functional gain.
"""

import json
import time
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List, NamedTuple, Optional, Sequence, Set, Tuple

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from .gclass_auth import _get_google_classroom_credentials, list_google_classroom_courses
from .utils import get_input_with_quit, parse_selection

# Documented Classroom quota is 3,000 queries/minute/project (50 QPS) and 1,200
# queries/minute/user (20 QPS), checked on a 60-second moving average
# (developers.google.com/workspace/classroom/limits). The loop paces well under
# both rather than racing a whole class into a 429.
_WRITE_INTERVAL = 0.2
_TRANSIENT_STATUSES = (429, 500, 503)
_RETRY_ATTEMPTS = 4
_RETRY_BASE_SECONDS = 1.0

_SECTION_FIELDS = ("Section", "Canvas Section", "Additional Section")

STATUS_INVITED = "invited"
STATUS_SKIPPED_ENROLLED = "skipped_enrolled"
STATUS_SKIPPED_PENDING = "skipped_pending"
STATUS_SKIPPED_MISSING = "skipped_missing"
STATUS_SKIPPED_DUPLICATE = "skipped_duplicate"
STATUS_SKIPPED_NOT_IN_DB = "skipped_not_in_db"
STATUS_SKIPPED_ABORTED = "skipped_aborted"
STATUS_SKIPPED_PRECONDITION = "skipped_precondition"
STATUS_FAILED = "failed"

_COUNTER_KEYS = (
    STATUS_INVITED,
    STATUS_SKIPPED_ENROLLED,
    STATUS_SKIPPED_PENDING,
    STATUS_SKIPPED_MISSING,
    STATUS_SKIPPED_DUPLICATE,
    STATUS_SKIPPED_NOT_IN_DB,
    STATUS_SKIPPED_ABORTED,
    STATUS_SKIPPED_PRECONDITION,
    STATUS_FAILED,
)


class _Candidate(NamedTuple):
    """One database student considered for invitation."""

    name: str
    email: str
    google_id: str


def _normalize_values(value: Any) -> List[str]:
    """Split a comma string or sequence into stripped lowercase entries."""
    if not value:
        return []
    raw = value
    if not isinstance(raw, (list, tuple, set)):
        raw = [v.strip() for v in str(raw).split(",") if v.strip()]
    return [str(v).strip().lower() for v in raw if str(v).strip()]


def _normalize_domains(value: Any) -> List[str]:
    return [d[1:] if d.startswith("@") else d for d in _normalize_values(value)]


def _field(student: Any, name: str) -> str:
    """Read one Student attribute.  Field names contain spaces, so getattr."""
    return (getattr(student, name, "") or "").strip()


def _utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _http_status(exc: HttpError) -> Optional[int]:
    return getattr(getattr(exc, "resp", None), "status", None)


def _canonical_status(exc: HttpError) -> str:
    """Read the canonical error name (``FAILED_PRECONDITION``, ...) from the body.

    Several distinct Classroom conditions share HTTP 400, so the numeric status
    alone cannot classify them; the canonical name in ``error.status`` can.
    """
    content = getattr(exc, "content", None)
    if not content:
        return ""
    try:
        if isinstance(content, bytes):
            content = content.decode("utf-8", "replace")
        body = json.loads(content)
    except (ValueError, AttributeError):
        return ""
    if not isinstance(body, dict):
        return ""
    error = body.get("error")
    if not isinstance(error, dict):
        return ""
    return str(error.get("status") or "")


def _api_message(exc: HttpError) -> str:
    content = getattr(exc, "content", None)
    if not content:
        return ""
    try:
        if isinstance(content, bytes):
            content = content.decode("utf-8", "replace")
        body = json.loads(content)
    except (ValueError, AttributeError):
        return ""
    error = body.get("error") if isinstance(body, dict) else None
    if not isinstance(error, dict):
        return ""
    return str(error.get("message") or "")


def _paginate(list_fn: Callable[..., Any], key: str, page_size: int, **kwargs: Any) -> List[Dict[str, Any]]:
    """Collect every page of a Classroom list call."""
    items: List[Dict[str, Any]] = []
    token: Optional[str] = None
    while True:
        params = dict(kwargs)
        params["pageSize"] = page_size
        if token:
            params["pageToken"] = token
        resp = list_fn(**params).execute()
        items.extend(resp.get(key, []) or [])
        token = resp.get("nextPageToken")
        if not token:
            break
    return items


def _execute_with_retry(
    make_request: Callable[[], Any],
    sleep_fn: Callable[[float], None],
    attempts: int = _RETRY_ATTEMPTS,
    base: float = _RETRY_BASE_SECONDS,
) -> Any:
    """Run a request, retrying only on transient statuses."""
    for attempt in range(1, attempts + 1):
        try:
            return make_request().execute()
        except HttpError as exc:
            if _http_status(exc) not in _TRANSIENT_STATUSES or attempt == attempts:
                raise
            sleep_fn(base * (2 ** (attempt - 1)))


def _new_report(course_id: str, dry_run: bool) -> Dict[str, Any]:
    report: Dict[str, Any] = {"course_id": course_id, "results": [], "dry_run": bool(dry_run)}
    for key in _COUNTER_KEYS:
        report[key] = 0
    return report


def _record(
    report: Dict[str, Any],
    name: str,
    email: str,
    status: str,
    detail: str = "",
    user_id: str = "",
    invitation_id: str = "",
) -> None:
    report["results"].append({
        "name": name,
        "email": email,
        "status": status,
        "detail": detail,
        "user_id": user_id,
        "invitation_id": invitation_id,
    })
    report[status] += 1


def _select_course_interactively(
    credentials_path: str,
    token_path: str,
    verbose: bool,
    open_browser: Optional[bool] = None,
) -> Optional[str]:
    courses = list_google_classroom_courses(
        credentials_path, token_path, verbose=verbose, open_browser=open_browser
    )
    if not courses:
        print("No courses found.")
        return None
    print("Available Google Classroom courses:")
    for i, c in enumerate(courses, 1):
        print(f"{i}. {c.get('name')} (ID: {c.get('id')})")
    while True:
        sel = get_input_with_quit("Select course number (or 'q' to quit): ")
        if sel is None:
            return None
        try:
            idx = int(sel.strip()) - 1
        except (TypeError, ValueError):
            continue
        if 0 <= idx < len(courses):
            return courses[idx].get("id")


def _collect_candidates(
    students: Sequence[Any],
    report: Dict[str, Any],
    emails_list: Sequence[str],
    domains_list: Sequence[str],
    section: Optional[str],
    class_name: Optional[str],
) -> List[_Candidate]:
    """Filter database students down to the ones worth inviting."""
    section_wanted = (section or "").strip().lower()
    class_wanted = (class_name or "").strip().lower()
    email_filter = set(emails_list)

    candidates: List[_Candidate] = []
    seen: Set[str] = set()
    found_emails: Set[str] = set()

    for student in students:
        name = _field(student, "Name") or _field(student, "Full Name")
        email = _field(student, "Email")
        if not email:
            _record(report, name, "", STATUS_SKIPPED_MISSING, "no email address in the local database")
            continue
        email_lower = email.lower()

        if email_filter:
            if email_lower not in email_filter:
                continue
            found_emails.add(email_lower)
        if domains_list and not any(email_lower.endswith("@" + d) for d in domains_list):
            continue
        if section_wanted:
            values = [_field(student, f).lower() for f in _SECTION_FIELDS]
            if section_wanted not in [v for v in values if v]:
                continue
        if class_wanted and _field(student, "Class").lower() != class_wanted:
            continue

        if email_lower in seen:
            _record(report, name, email, STATUS_SKIPPED_DUPLICATE, "duplicate email in the local database")
            continue
        seen.add(email_lower)
        candidates.append(_Candidate(name=name, email=email, google_id=_field(student, "Google_ID")))

    for requested in email_filter - found_emails:
        _record(
            report,
            "",
            requested,
            STATUS_SKIPPED_NOT_IN_DB,
            f"{requested} is not in the local database -- run: course --add-google-sheet <URL>",
        )
    return candidates


class _Membership(NamedTuple):
    """Tier 1 membership.  The teacher sets are subsets of the first two.

    Both roles have to live in one set for the skip decision -- a teacher is as
    much "already on the course" as a student -- but the skip *reason* needs
    them apart, so the teacher halves are kept alongside rather than merged.
    """

    emails: Set[str]
    user_ids: Set[str]
    teacher_emails: Set[str]
    teacher_user_ids: Set[str]


def _read_course_membership(service: Any, course_id: str, verbose: bool) -> _Membership:
    """Tier 1: everyone already on the course, as emails and as user ids."""
    emails: Set[str] = set()
    user_ids: Set[str] = set()
    teacher_emails: Set[str] = set()
    teacher_user_ids: Set[str] = set()
    for key, entries in (
        ("students", _paginate(service.courses().students().list, "students", 200, courseId=course_id)),
        ("teachers", _paginate(service.courses().teachers().list, "teachers", 200, courseId=course_id)),
    ):
        is_teacher = key == "teachers"
        for entry in entries:
            profile = entry.get("profile", {}) or {}
            email = (profile.get("emailAddress") or "").strip().lower()
            if email:
                emails.add(email)
                if is_teacher:
                    teacher_emails.add(email)
            user_id = entry.get("userId")
            if user_id:
                user_ids.add(str(user_id))
                if is_teacher:
                    teacher_user_ids.add(str(user_id))
        if verbose:
            print(f"[GClassroom] Read {len(entries)} {key} already on the course.")
    return _Membership(emails, user_ids, teacher_emails, teacher_user_ids)


def _read_pending_invitations(
    service: Any,
    course_id: str,
    known_google_ids: Set[str],
    profile_lookup_limit: int,
    verbose: bool,
) -> Tuple[Set[str], Set[str]]:
    """Tier 2: pending invitations, as resolved emails and as user ids.

    An Invitation carries a numeric ``userId`` and never an email address, so
    matching by email needs a per-invitation profile lookup.  Local Google_ID
    values are checked first because they cost nothing, and the fan-out for the
    rest is capped -- the 409 backstop covers whatever this misses.
    """
    invitations = _paginate(service.invitations().list, "invitations", 500, courseId=course_id)
    user_ids: Set[str] = set()
    emails: Set[str] = set()
    lookups = 0
    limit_notice_shown = False
    for invitation in invitations:
        user_id = invitation.get("userId")
        if not user_id:
            continue
        user_id = str(user_id)
        user_ids.add(user_id)
        if user_id in known_google_ids:
            continue
        if lookups >= profile_lookup_limit:
            # continue, not break: the userId above is already recorded, and
            # matching a candidate by Google_ID against it costs nothing.
            if verbose and not limit_notice_shown:
                limit_notice_shown = True
                print(
                    f"[GClassroom] Profile lookup limit ({profile_lookup_limit}) reached; "
                    "remaining pending invitations rely on Google_ID and the 409 check."
                )
            continue
        lookups += 1
        try:
            profile = service.userProfiles().get(userId=user_id).execute()
        except HttpError as exc:
            if verbose:
                print(f"[GClassroom] Could not resolve pending invitation {user_id}: {exc}")
            continue
        email = (profile.get("emailAddress") or "").strip().lower()
        if email:
            emails.add(email)
    if verbose:
        print(f"[GClassroom] Read {len(invitations)} pending invitation(s), {lookups} profile lookup(s).")
    return emails, user_ids


def _write_back_google_ids(
    db_path: str,
    invited: Dict[str, Dict[str, str]],
    verbose: bool,
) -> None:
    """Fill Google_ID and invitation bookkeeping without ever overwriting."""
    from .data import load_database, save_database

    try:
        students = load_database(db_path, verbose=verbose)
    except Exception as exc:
        print(
            f"Warning: could not read {db_path} to record {len(invited)} sent invitation(s): {exc}. "
            "The invitations were sent; Google_ID was not saved, so the next run spends "
            "profile lookups again."
        )
        return
    if not students:
        print(
            f"Warning: {db_path} holds no student records, so {len(invited)} sent invitation(s) "
            "were not recorded. The next run spends profile lookups again."
        )
        return
    updated = 0
    for student in students:
        email = _field(student, "Email").lower()
        entry = invited.get(email)
        if not entry:
            continue
        if entry.get("user_id") and not _field(student, "Google_ID"):
            setattr(student, "Google_ID", entry["user_id"])
        if entry.get("invitation_id"):
            setattr(student, "Google_Invitation_ID", entry["invitation_id"])
        setattr(student, "Google_Classroom_Invited_At", entry.get("invited_at", ""))
        updated += 1
    if updated:
        save_database(students, db_path=db_path, verbose=verbose, audit_source="gclass-invite")
        print(f"Updated {updated} student record(s) in the local database.")


def invite_students_to_google_classroom(
    course_id=None,
    emails=None,
    domains=None,
    section=None,
    class_name=None,
    role="STUDENT",
    apply_all=False,
    credentials_path='gclassroom_credentials.json',
    token_path='token.pickle',
    db_path=None,
    update_local_db=True,
    dry_run=False,
    verbose=False,
    service=None,
    sleep_fn=None,
    profile_lookup_limit=200,
):
    """Invite local-database students to a Google Classroom course.

    Students already enrolled, already teaching, or already holding a pending
    invitation are skipped, so re-running the command is safe and quiet.
    Returns a summary dict; see the module docstring for the departures from
    the unenroll template.
    """
    sleep_fn = sleep_fn or time.sleep

    if service is None:
        creds = _get_google_classroom_credentials(credentials_path, token_path, verbose=verbose)
        service = build("classroom", "v1", credentials=creds)

    if not course_id:
        course_id = _select_course_interactively(credentials_path, token_path, verbose)
    if not course_id:
        print("No course selected.")
        return _new_report("", dry_run)

    report = _new_report(str(course_id), dry_run)

    if not db_path:
        print("Inviting students requires a local database. Provide --db or run from the DB folder.")
        return report
    from .data import load_database

    try:
        students = load_database(db_path, verbose=verbose)
    except Exception:
        students = []

    candidates = _collect_candidates(
        students or [],
        report,
        _normalize_values(emails),
        _normalize_domains(domains),
        section,
        class_name,
    )
    if not candidates:
        print("No matching students found in the local database.")
        return report

    membership = _read_course_membership(service, course_id, verbose)
    known_google_ids = {c.google_id for c in candidates if c.google_id}
    pending_emails, pending_user_ids = _read_pending_invitations(
        service, course_id, known_google_ids, profile_lookup_limit, verbose
    )

    to_invite: List[_Candidate] = []
    for candidate in candidates:
        email_lower = candidate.email.lower()
        on_course = email_lower in membership.emails or (
            candidate.google_id and candidate.google_id in membership.user_ids
        )
        if on_course:
            is_teacher = email_lower in membership.teacher_emails or (
                candidate.google_id and candidate.google_id in membership.teacher_user_ids
            )
            detail = "already a teacher on this course" if is_teacher else "already enrolled"
            _record(report, candidate.name, candidate.email, STATUS_SKIPPED_ENROLLED, detail,
                    user_id=candidate.google_id)
            continue
        if email_lower in pending_emails or (candidate.google_id and candidate.google_id in pending_user_ids):
            _record(report, candidate.name, candidate.email, STATUS_SKIPPED_PENDING,
                    "invitation already pending", user_id=candidate.google_id)
            continue
        to_invite.append(candidate)

    if not to_invite:
        print(
            f"Everyone is already on the course: {report[STATUS_SKIPPED_ENROLLED]} enrolled, "
            f"{report[STATUS_SKIPPED_PENDING]} pending. Nothing to invite."
        )
        return report

    to_invite.sort(key=lambda c: (c.name, c.email))
    print(f"{len(to_invite)} student(s) are not yet on the course:")
    for idx, candidate in enumerate(to_invite, 1):
        print(f"{idx}. {candidate.name} | {candidate.email}")

    if apply_all:
        selected_indices = list(range(1, len(to_invite) + 1))
    else:
        while True:
            sel = get_input_with_quit("Select students to invite (e.g. 1,3-5, 'a' for all, or 'q' to quit): ")
            if sel is None:
                return report
            selected_indices = parse_selection(sel, len(to_invite))
            if selected_indices:
                break
            print("Invalid selection. Please enter valid numbers, a range, 'a' for all, or 'q' to quit.")

    confirm = get_input_with_quit(
        f"Invite {len(selected_indices)} student(s) to the course as {role}? (y/n): ",
        default="n",
    )
    if confirm is None or confirm.lower() not in ("y", "yes"):
        print("Invite canceled.")
        return report

    selected = [to_invite[i - 1] for i in selected_indices]

    if dry_run:
        for candidate in selected:
            _record(report, candidate.name, candidate.email, STATUS_INVITED, "dry-run: not sent")
        print(f"Dry-run: would invite {len(selected)} student(s).")
        return report

    invited_for_db: Dict[str, Dict[str, str]] = {}
    aborted = False
    for position, candidate in enumerate(selected):
        if aborted:
            _record(report, candidate.name, candidate.email, STATUS_SKIPPED_ABORTED,
                    "skipped after an authorization failure")
            continue
        if position:
            sleep_fn(_WRITE_INTERVAL)
        body = {"courseId": str(course_id), "userId": candidate.email, "role": role}
        try:
            created = _execute_with_retry(
                lambda: service.invitations().create(body=body), sleep_fn
            ) or {}
        except HttpError as exc:
            status = _http_status(exc)
            if status == 409:
                _record(report, candidate.name, candidate.email, STATUS_SKIPPED_PENDING,
                        "invitation already pending")
                continue
            if status == 400 and _canonical_status(exc) == "FAILED_PRECONDITION":
                # Documented as "the requested user's account is disabled" or "the
                # user already has this role or a role with greater permissions".
                # The second is an enrolment the bulk read missed, so this is the
                # backstop 409 does not provide -- 409 covers only a pending
                # invitation. Neither cause is a write that failed, and neither is
                # fixed by retrying, so it is reported rather than counted failed.
                # The API message is passed through instead of matched on: the
                # two causes are documented, their wording is not.
                _record(report, candidate.name, candidate.email,
                        STATUS_SKIPPED_PRECONDITION,
                        _api_message(exc) or "precondition failed")
                print(
                    f"Skipped {candidate.name} ({candidate.email}): "
                    "already holds this role or greater, or the account is disabled."
                )
                continue
            if status == 404:
                # The API documents 404 as "the course or the user does not exist",
                # so a wrong course id lands here too.
                _record(report, candidate.name, candidate.email, STATUS_FAILED,
                        f"no Google account for {candidate.email}, the address is not "
                        f"visible to this Workspace domain, or course {course_id} does "
                        "not exist")
                print(
                    f"Failed to invite {candidate.name} ({candidate.email}): "
                    "no such user, or no such course."
                )
                continue
            if status == 403:
                _record(report, candidate.name, candidate.email, STATUS_FAILED,
                        "not authorized to invite on this course, or domain policy blocks the invitation")
                print(f"Failed to invite {candidate.name} ({candidate.email}): not authorized.")
                print("Stopping: a 403 applies to the whole course, not to one student.")
                aborted = True
                continue
            _record(report, candidate.name, candidate.email, STATUS_FAILED, str(exc))
            print(f"Failed to invite {candidate.name} ({candidate.email}).")
            continue
        except Exception as exc:
            _record(report, candidate.name, candidate.email, STATUS_FAILED, str(exc))
            print(f"Failed to invite {candidate.name} ({candidate.email}).")
            continue

        user_id = str(created.get("userId") or "")
        invitation_id = str(created.get("id") or "")
        _record(report, candidate.name, candidate.email, STATUS_INVITED,
                user_id=user_id, invitation_id=invitation_id)
        invited_for_db[candidate.email.lower()] = {
            "user_id": user_id,
            "invitation_id": invitation_id,
            "invited_at": _utc_now(),
        }
        if verbose:
            print(f"[GClassroom] Invited {candidate.name} ({candidate.email}).")

    print(
        f"Invited {report[STATUS_INVITED]} student(s); skipped "
        f"{report[STATUS_SKIPPED_ENROLLED]} already enrolled, "
        f"{report[STATUS_SKIPPED_PENDING]} already invited."
    )
    if report[STATUS_FAILED]:
        print(f"{report[STATUS_FAILED]} invitation(s) failed.")

    if invited_for_db and update_local_db and db_path:
        _write_back_google_ids(db_path, invited_for_db, verbose)

    return report
