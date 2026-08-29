# -*- coding: utf-8 -*-
"""Google Classroom assignment administration CLI.

The mutation surface is deliberately separate from the legacy ``course`` CLI and
the restricted agent entrypoint. Assignment JSON is converted into a canonical
request before any credential access or network call.
"""

from __future__ import annotations

import argparse
import copy
import errno
import getpass
import hashlib
import hmac
import http.client
import json
import os
import re
import stat
import sys
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Mapping, Optional, Sequence, TextIO, Tuple
from urllib.parse import parse_qsl, urlsplit

from .course_agent_common import (
    CourseAgentError,
    is_agent_mode,
    require_env_allowlist,
)
from .gclass_coursework import (
    GoogleClassroomAPIError,
    GoogleClassroomOutcomeUnknownError,
    GoogleClassroomPartialCreateError,
    build_google_classroom_assignment_body,
    build_google_classroom_rubric_body,
    create_google_classroom_assignment_with_rubric,
    google_classroom_drive_material,
    google_classroom_link_material,
    google_classroom_youtube_material,
    _execute_mutation,
)
from .gclass_coursework_auth import (
    CredentialSecurityError,
    CourseworkAuthPaths,
    TokenFileLock,
    get_google_classroom_auth_status,
    get_google_classroom_coursework_service,
    normalize_expected_account,
    resolve_coursework_auth_paths,
)


MAX_ASSIGNMENT_SPEC_BYTES = 1024 * 1024
MAX_LOOPBACK_CALLBACK_URL_CHARS = 16 * 1024
MAX_LOOPBACK_RESPONSE_BYTES = 8 * 1024
_COURSE_REFERENCE_RE = re.compile(r"^[A-Za-z0-9:_-]{1,256}$")
_CANONICAL_COURSE_ID_RE = re.compile(r"^[A-Za-z0-9_-]{1,256}$")
_COURSEWORK_ID_RE = re.compile(r"^[A-Za-z0-9_-]{1,256}$")
_FULL_DIGEST_RE = re.compile(r"^[0-9a-f]{64}$")
_CLIENT_FINGERPRINT_RE = re.compile(r"^[0-9a-f]{20}$")
_AGENT_SAFE_DRAFT_FIELDS = {
    "title",
    "description",
    "workType",
    "state",
    "assigneeMode",
    "submissionModificationMode",
}
_ALL_COURSEWORK_STATES = ["DRAFT", "PUBLISHED", "DELETED"]
MAX_AGENT_DRAFT_LIST_PAGES = 1000
_COURSE_LOOKUP_FIELDS = "id,name,courseState"
_AGENT_DRAFT_LIST_FIELDS = "nextPageToken,courseWork(id,title,state)"
_AGENT_DRAFT_CREATE_FIELDS = "id,courseId,state,associatedWithDeveloper"
_AGENT_DRAFT_READ_FIELDS = ",".join(
    (
        "id",
        "courseId",
        "title",
        "description",
        "workType",
        "state",
        "assigneeMode",
        "individualStudentsOptions",
        "submissionModificationMode",
        "scheduledTime",
        "dueDate",
        "dueTime",
        "maxPoints",
        "materials",
        "topicId",
        "gradingPeriodId",
        "associatedWithDeveloper",
    )
)
_TOP_LEVEL_FIELDS = {
    "title",
    "description",
    "state",
    "scheduled_at",
    "due_at",
    "max_points",
    "materials",
    "assignee",
    "submission_modification_mode",
    "topic_id",
    "grading_period",
    "rubric",
}


class AssignmentSpecError(ValueError):
    """The local assignment specification is malformed or unsupported."""


class GoogleClassroomAdminError(RuntimeError):
    """An assignment-admin precondition or confirmation failed."""


def _strict_pairs(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise AssignmentSpecError(f"Duplicate JSON key: {key}")
        result[key] = value
    return result


def _read_assignment_spec_json(path: Path) -> Dict[str, Any]:
    path = Path(path)
    flags = (
        os.O_RDONLY
        | getattr(os, "O_CLOEXEC", 0)
        | getattr(os, "O_NOFOLLOW", 0)
        | getattr(os, "O_NONBLOCK", 0)
    )
    descriptor: Optional[int] = None
    try:
        descriptor = os.open(path, flags)
        info = os.fstat(descriptor)
        if not stat.S_ISREG(info.st_mode):
            raise AssignmentSpecError("Assignment spec must be a regular file")
        if info.st_size > MAX_ASSIGNMENT_SPEC_BYTES:
            raise AssignmentSpecError("Assignment spec exceeds the 1 MiB size limit")
        chunks = []
        total = 0
        while True:
            chunk = os.read(
                descriptor,
                min(65536, MAX_ASSIGNMENT_SPEC_BYTES + 1 - total),
            )
            if not chunk:
                break
            chunks.append(chunk)
            total += len(chunk)
            if total > MAX_ASSIGNMENT_SPEC_BYTES:
                raise AssignmentSpecError("Assignment spec exceeds the 1 MiB size limit")
    except AssignmentSpecError:
        raise
    except OSError as exc:
        if exc.errno == errno.ELOOP:
            raise AssignmentSpecError("Assignment spec must not be a symlink") from None
        raise AssignmentSpecError("Assignment spec could not be opened safely") from None
    finally:
        if descriptor is not None:
            os.close(descriptor)

    try:
        text = b"".join(chunks).decode("utf-8")
        value = json.loads(text, object_pairs_hook=_strict_pairs)
    except AssignmentSpecError:
        raise
    except (UnicodeDecodeError, json.JSONDecodeError):
        raise AssignmentSpecError("Assignment spec must be strict UTF-8 JSON") from None
    if not isinstance(value, dict):
        raise AssignmentSpecError("Assignment spec must be a JSON object")
    return value


def _reject_unknown_fields(value: Mapping[str, Any], allowed: set, label: str) -> None:
    unknown = set(value) - allowed
    if unknown:
        raise AssignmentSpecError(f"Unsupported {label} fields: {sorted(unknown)}")


def _build_materials(value: Any) -> Optional[List[Dict[str, Any]]]:
    if value is None:
        return None
    if not isinstance(value, list):
        raise AssignmentSpecError("materials must be an array or null")
    materials: List[Dict[str, Any]] = []
    for item in value:
        if not isinstance(item, Mapping):
            raise AssignmentSpecError("Each material must be an object")
        kind = item.get("type")
        if kind == "drive_file":
            _reject_unknown_fields(
                item, {"type", "file_id", "share_mode"}, "Drive material"
            )
            if "file_id" not in item:
                raise AssignmentSpecError("Drive material requires file_id")
            materials.append(
                google_classroom_drive_material(
                    item["file_id"], item.get("share_mode", "VIEW")
                )
            )
        elif kind == "link":
            _reject_unknown_fields(item, {"type", "url"}, "link material")
            if "url" not in item:
                raise AssignmentSpecError("Link material requires url")
            materials.append(google_classroom_link_material(item["url"]))
        elif kind == "youtube":
            _reject_unknown_fields(
                item, {"type", "video_id"}, "YouTube material"
            )
            if "video_id" not in item:
                raise AssignmentSpecError("YouTube material requires video_id")
            materials.append(google_classroom_youtube_material(item["video_id"]))
        else:
            raise AssignmentSpecError(
                "Material type must be drive_file, link, or youtube"
            )
    return materials


def load_assignment_spec(
    path: Path,
    *,
    now: Optional[datetime] = None,
) -> Tuple[Dict[str, Any], Optional[Dict[str, Any]]]:
    """Load a strict JSON spec and return canonical assignment and rubric bodies."""

    raw = _read_assignment_spec_json(Path(path))
    _reject_unknown_fields(raw, _TOP_LEVEL_FIELDS, "assignment")
    if "title" not in raw:
        raise AssignmentSpecError("Assignment spec requires title")

    assignee = raw.get("assignee")
    if assignee is None:
        assignee = {}
    if not isinstance(assignee, Mapping):
        raise AssignmentSpecError("assignee must be an object or null")
    _reject_unknown_fields(assignee, {"mode", "student_ids"}, "assignee")

    grading_period = raw.get("grading_period")
    if grading_period is None:
        grading_period = {}
    if not isinstance(grading_period, Mapping):
        raise AssignmentSpecError("grading_period must be an object or null")
    _reject_unknown_fields(grading_period, {"mode", "id"}, "grading period")

    try:
        assignment = build_google_classroom_assignment_body(
            raw["title"],
            description=raw.get("description"),
            state=raw.get("state", "DRAFT"),
            scheduled_at=raw.get("scheduled_at"),
            due_at=raw.get("due_at"),
            max_points=raw.get("max_points"),
            materials=_build_materials(raw.get("materials")),
            assignee_mode=assignee.get("mode", "ALL_STUDENTS"),
            individual_student_ids=assignee.get("student_ids"),
            submission_modification_mode=raw.get(
                "submission_modification_mode", "MODIFIABLE_UNTIL_TURNED_IN"
            ),
            topic_id=raw.get("topic_id"),
            grading_period_mode=grading_period.get("mode", "AUTO"),
            grading_period_id=grading_period.get("id"),
            now=now,
        )
    except AssignmentSpecError:
        raise
    except (TypeError, ValueError) as exc:
        raise AssignmentSpecError(str(exc)) from None

    rubric_spec = raw.get("rubric")
    rubric: Optional[Dict[str, Any]] = None
    if rubric_spec is not None:
        if not isinstance(rubric_spec, Mapping):
            raise AssignmentSpecError("rubric must be an object or null")
        _reject_unknown_fields(
            rubric_spec, {"scoring_mode", "criteria"}, "rubric"
        )
        if "scoring_mode" not in rubric_spec or "criteria" not in rubric_spec:
            raise AssignmentSpecError("rubric requires scoring_mode and criteria")
        try:
            rubric = build_google_classroom_rubric_body(
                rubric_spec["criteria"],
                scoring_mode=rubric_spec["scoring_mode"],
            )
        except (TypeError, ValueError) as exc:
            raise AssignmentSpecError(str(exc)) from None
    return assignment, rubric


def _validate_course_reference(value: str) -> str:
    if not isinstance(value, str) or not _COURSE_REFERENCE_RE.fullmatch(value):
        raise AssignmentSpecError(
            "course ID must be a Classroom ID or alias without whitespace"
        )
    return value


def _release_status(assignment: Mapping[str, Any]) -> str:
    if assignment.get("scheduledTime"):
        return "SCHEDULED"
    return "PUBLISHED" if assignment.get("state") == "PUBLISHED" else "DRAFT"


def build_assignment_operation_plan(
    course_reference: str,
    assignment: Mapping[str, Any],
    rubric: Optional[Mapping[str, Any]],
    *,
    dry_run: bool = True,
) -> Dict[str, Any]:
    """Build the deterministic preview and confirmation digest for one mutation."""

    course_reference = _validate_course_reference(course_reference)
    canonical_payload: Dict[str, Any] = {
        "courseReference": course_reference,
        "assignment": copy.deepcopy(dict(assignment)),
        "rubric": copy.deepcopy(dict(rubric)) if rubric is not None else None,
    }

    release = _release_status(assignment)
    if rubric is None:
        operations = ["create-assignment"]
    else:
        operations = ["create-draft", "create-rubric"]
        if release == "PUBLISHED":
            operations.append("publish-assignment")
        elif release == "SCHEDULED":
            operations.append("schedule-assignment")

    drive_share_modes = sorted(
        {
            material["driveFile"]["shareMode"]
            for material in assignment.get("materials", [])
            if isinstance(material, Mapping)
            and isinstance(material.get("driveFile"), Mapping)
            and "shareMode" in material["driveFile"]
        }
    )
    plan = {
        "schemaVersion": 1,
        "executionMode": "DRY_RUN" if dry_run else "LIVE",
        "dryRun": bool(dry_run),
        **canonical_payload,
        "releaseStatus": release,
        "operations": operations,
        "requiresDriveSharingConfirmation": bool(drive_share_modes),
        "driveShareModes": drive_share_modes,
    }
    digest_input = json.dumps(
        plan,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    plan["operationDigest"] = hashlib.sha256(digest_input).hexdigest()
    return plan


def _mask_account(account: str) -> str:
    local, domain = account.split("@", 1)
    return f"{local[:1]}***@{domain}"


def _write_json(stream: TextIO, value: Mapping[str, Any]) -> None:
    stream.write(json.dumps(value, ensure_ascii=False, sort_keys=True) + "\n")


def _path_entry_exists(path: Optional[Path]) -> bool:
    if path is None:
        return False
    try:
        os.lstat(path)
        return True
    except FileNotFoundError:
        return False


def _require_interactive_human(tty_check: Callable[[], bool]) -> None:
    if is_agent_mode():
        raise CourseAgentError(
            "Google Classroom credential and mutation operations are forbidden in agent mode",
            code="agent_forbidden",
        )
    if not tty_check():
        raise GoogleClassroomAdminError(
            "Credential and mutation operations require an interactive terminal"
        )


def _terminal_text(value: Any) -> str:
    """Escape terminal control and formatting characters in API-derived text."""

    result = []
    for character in str(value):
        category = unicodedata.category(character)
        if category in {"Cc", "Cf", "Cs", "Zl", "Zp"}:
            result.append(f"\\u{ord(character):04x}")
        else:
            result.append(character)
    return "".join(result)


def _creation_summary(
    account: str,
    course: Mapping[str, Any],
    assignment: Mapping[str, Any],
    plan: Mapping[str, Any],
) -> str:
    materials = assignment.get("materials") or []
    lines = [
        "Google Classroom assignment",
        f"  Account: {_terminal_text(_mask_account(account))}",
        (
            "  Course: "
            f"{_terminal_text(course.get('name') or '(unnamed)')} "
            f"({_terminal_text(course['id'])})"
        ),
        f"  Title: {_terminal_text(assignment.get('title') or '')}",
        f"  Release: {_terminal_text(plan['releaseStatus'])}",
        f"  Materials: {len(materials)}",
        f"  Rubric: {'yes' if plan.get('rubric') is not None else 'no'}",
    ]
    if assignment.get("scheduledTime"):
        lines.append(
            f"  Scheduled: {_terminal_text(assignment['scheduledTime'])}"
        )
    if assignment.get("dueDate"):
        due = json.dumps(
            {
                "date": assignment["dueDate"],
                "time": assignment.get("dueTime"),
            },
            ensure_ascii=True,
            sort_keys=True,
            separators=(",", ":"),
        )
        lines.append(f"  Due: {due}")
    if "maxPoints" in assignment:
        lines.append(f"  Points: {assignment['maxPoints']}")
    drive_modes = plan.get("driveShareModes") or []
    lines.append(
        "  Drive sharing: "
        + (", ".join(_terminal_text(mode) for mode in drive_modes) or "none")
    )
    return "\n".join(lines)


def _confirm_creation(
    input_fn: Callable[[str], str],
    *,
    account: str,
    course: Mapping[str, Any],
    assignment: Mapping[str, Any],
    plan: Mapping[str, Any],
) -> None:
    prompt = _creation_summary(account, course, assignment, plan)
    try:
        received = input_fn(f"{prompt}\nCreate assignment? [y/N]: ")
    except (EOFError, StopIteration):
        raise GoogleClassroomAdminError("Assignment creation cancelled") from None
    if not isinstance(received, str) or received.strip().lower() not in {"y", "yes"}:
        raise GoogleClassroomAdminError("Assignment creation cancelled")


def _require_create_mode(
    *,
    yes: bool,
    tty_check: Callable[[], bool],
    plan: Mapping[str, Any],
) -> None:
    if is_agent_mode():
        raise CourseAgentError(
            "Google Classroom mutation operations are forbidden in agent mode",
            code="agent_forbidden",
        )
    if yes:
        if plan.get("releaseStatus") != "DRAFT":
            raise GoogleClassroomAdminError(
                "--yes is limited to DRAFT assignments; published and scheduled work requires interactive confirmation"
            )
        if plan.get("requiresDriveSharingConfirmation"):
            raise GoogleClassroomAdminError(
                "--yes cannot create assignments with Drive-sharing effects; use interactive confirmation"
            )
        return
    if not tty_check():
        raise GoogleClassroomAdminError(
            "Assignment creation requires an interactive terminal or --yes for a safe draft"
        )


def _validate_agent_safe_draft_shape(
    assignment: Mapping[str, Any],
    rubric: Optional[Mapping[str, Any]],
    plan: Mapping[str, Any],
) -> None:
    """Accept only the frozen minimal smoke-test shape."""

    unexpected = set(assignment) - _AGENT_SAFE_DRAFT_FIELDS
    if unexpected:
        raise GoogleClassroomAdminError(
            "Agent-safe draft contains disallowed fields: "
            + ", ".join(sorted(unexpected))
        )
    required = {
        "workType": "ASSIGNMENT",
        "state": "DRAFT",
        "assigneeMode": "ALL_STUDENTS",
        "submissionModificationMode": "MODIFIABLE_UNTIL_TURNED_IN",
    }
    mismatched = [
        field for field, expected in required.items() if assignment.get(field) != expected
    ]
    if mismatched:
        raise GoogleClassroomAdminError(
            "Agent-safe draft has unsupported fields: "
            + ", ".join(sorted(mismatched))
        )
    if rubric is not None:
        raise GoogleClassroomAdminError("Agent-safe draft cannot contain a rubric")
    if (
        plan.get("releaseStatus") != "DRAFT"
        or plan.get("operations") != ["create-assignment"]
        or plan.get("requiresDriveSharingConfirmation")
    ):
        raise GoogleClassroomAdminError(
            "Agent-safe draft plan is not a minimal no-sharing DRAFT"
        )


def _require_agent_safe_context(
    *,
    account: str,
    course_id: str,
    assignment: Mapping[str, Any],
    rubric: Optional[Mapping[str, Any]],
    plan: Mapping[str, Any],
    require_yes: bool,
    yes: bool = False,
    expected_approval_digest: Optional[str] = None,
) -> None:
    if not is_agent_mode():
        raise CourseAgentError(
            "--agent-safe-draft requires explicit agent mode",
            code="agent_mode_required",
        )
    if not _CANONICAL_COURSE_ID_RE.fullmatch(course_id):
        raise GoogleClassroomAdminError(
            "Agent-safe draft requires an exact canonical course ID; aliases are forbidden"
        )
    try:
        require_env_allowlist(
            "GCLASS_ACCOUNT_ALLOWLIST", account, label="google account"
        )
    except CourseAgentError as exc:
        raise CourseAgentError(
            "Approved Google account is not allowlisted in GCLASS_ACCOUNT_ALLOWLIST",
            code=exc.code,
        ) from None
    require_env_allowlist(
        "GCLASS_COURSE_ALLOWLIST", course_id, label="google classroom course id"
    )
    _validate_agent_safe_draft_shape(assignment, rubric, plan)
    if require_yes and not yes:
        raise GoogleClassroomAdminError(
            "Agent-safe draft creation requires --yes"
        )
    if require_yes and (
        not isinstance(expected_approval_digest, str)
        or not _FULL_DIGEST_RE.fullmatch(expected_approval_digest)
    ):
        raise GoogleClassroomAdminError(
            "Agent-safe draft creation requires a full --expect-approval-digest"
        )


def _token_path_fingerprint(path: Path) -> str:
    normalized = os.path.normcase(os.path.normpath(str(Path(path).absolute())))
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def _agent_safe_approval_digest(
    *,
    account: str,
    course_id: str,
    plan: Mapping[str, Any],
    paths: CourseworkAuthPaths,
    client_id_fingerprint: str,
) -> str:
    if not _CLIENT_FINGERPRINT_RE.fullmatch(client_id_fingerprint):
        raise GoogleClassroomAdminError(
            "Authenticated OAuth client fingerprint is invalid"
        )
    payload = {
        "schemaVersion": 1,
        "mode": "agent-safe-google-classroom-draft",
        "account": account,
        "accountFingerprint": paths.account_fingerprint,
        "canonicalCourseId": course_id,
        "operationDigest": plan["operationDigest"],
        "clientIdFingerprint": client_id_fingerprint,
        "authSource": paths.source,
        "tokenPathFingerprint": _token_path_fingerprint(paths.token_path),
    }
    encoded = json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _authenticate_existing_agent_token(
    authenticate: Callable[..., Any],
    paths: CourseworkAuthPaths,
    *,
    account: str,
) -> Any:
    if not _path_entry_exists(paths.token_path):
        raise GoogleClassroomAdminError(
            "Agent-safe draft requires an existing coursework token; run authorize as a human first"
        )
    session = authenticate(
        paths,
        expected_account=account,
        open_browser=False,
        require_existing_token=True,
    )
    if session.paths != paths:
        raise GoogleClassroomAdminError(
            "Authenticated session did not use the resolved credential source"
        )
    profile_account = str(session.profile.get("emailAddress") or "").strip().lower()
    if profile_account != account:
        raise GoogleClassroomAdminError(
            "Authenticated Google account did not match the approved account"
        )
    return session


def _prepare_agent_safe_draft_context(
    *,
    authenticate: Callable[..., Any],
    account: str,
    course_id: str,
    assignment: Mapping[str, Any],
    rubric: Optional[Mapping[str, Any]],
    plan: Mapping[str, Any],
    paths: CourseworkAuthPaths,
) -> Tuple[Any, Dict[str, Any], str]:
    session = _authenticate_existing_agent_token(
        authenticate, paths, account=account
    )
    course = _get_canonical_course(session.service, course_id)
    canonical_course_id = str(course["id"])
    if canonical_course_id != course_id:
        raise GoogleClassroomAdminError(
            "Google Classroom course lookup did not match the exact approved course ID"
        )
    approval_digest = _agent_safe_approval_digest(
        account=account,
        course_id=canonical_course_id,
        plan=plan,
        paths=paths,
        client_id_fingerprint=session.client_id_fingerprint,
    )
    return session, course, approval_digest


def _coursework_shape_mismatches(
    value: Mapping[str, Any],
    *,
    course_id: str,
    course_work_id: str,
    assignment: Mapping[str, Any],
) -> List[str]:
    """Compare canonical writable semantics without exposing response values."""

    mismatches = []
    exact = {
        "courseId": course_id,
        "id": course_work_id,
        "title": assignment["title"],
        "workType": "ASSIGNMENT",
        "state": "DRAFT",
        "assigneeMode": "ALL_STUDENTS",
        "submissionModificationMode": "MODIFIABLE_UNTIL_TURNED_IN",
        "associatedWithDeveloper": True,
    }
    for field, expected in exact.items():
        if value.get(field) != expected:
            mismatches.append(field)
    expected_description = assignment.get("description") or ""
    if (value.get("description") or "") != expected_description:
        mismatches.append("description")
    empty_equivalents = {
        "materials": (None, []),
        "individualStudentsOptions": (None, {}),
    }
    for field, allowed in empty_equivalents.items():
        if value.get(field) not in allowed:
            mismatches.append(field)
    for field in (
        "scheduledTime",
        "dueDate",
        "dueTime",
        "maxPoints",
        "topicId",
        "gradingPeriodId",
    ):
        if value.get(field) is not None:
            mismatches.append(field)
    return sorted(set(mismatches))


def _read_coursework(
    service: Any,
    *,
    course_id: str,
    course_work_id: str,
) -> Dict[str, Any]:
    try:
        response = (
            service.courses()
            .courseWork()
            .get(
                courseId=course_id,
                id=course_work_id,
                fields=_AGENT_DRAFT_READ_FIELDS,
            )
            .execute(num_retries=0)
        )
    except Exception:
        raise GoogleClassroomAdminError(
            "Google Classroom assignment read-back failed"
        ) from None
    if not isinstance(response, Mapping):
        raise GoogleClassroomAdminError(
            "Google Classroom assignment read-back was invalid"
        )
    return copy.deepcopy(dict(response))


def _find_same_title_coursework(
    service: Any,
    *,
    course_id: str,
    title: str,
) -> List[Dict[str, Any]]:
    matches: List[Dict[str, Any]] = []
    page_token: Optional[str] = None
    seen_tokens = set()
    for _page_number in range(MAX_AGENT_DRAFT_LIST_PAGES):
        kwargs: Dict[str, Any] = {
            "courseId": course_id,
            "courseWorkStates": list(_ALL_COURSEWORK_STATES),
            "pageSize": 100,
            "fields": _AGENT_DRAFT_LIST_FIELDS,
        }
        if page_token is not None:
            kwargs["pageToken"] = page_token
        try:
            response = (
                service.courses()
                .courseWork()
                .list(**kwargs)
                .execute(num_retries=0)
            )
        except Exception:
            raise GoogleClassroomAdminError(
                "Google Classroom duplicate preflight failed"
            ) from None
        if not isinstance(response, Mapping):
            raise GoogleClassroomAdminError(
                "Google Classroom duplicate preflight returned an invalid response"
            )
        values = response.get("courseWork", [])
        if not isinstance(values, list):
            raise GoogleClassroomAdminError(
                "Google Classroom duplicate preflight returned invalid coursework"
            )
        for value in values:
            if not isinstance(value, Mapping):
                raise GoogleClassroomAdminError(
                    "Google Classroom duplicate preflight returned invalid coursework"
                )
            if value.get("title") == title:
                matches.append(copy.deepcopy(dict(value)))
                if len(matches) > 1:
                    return matches
        raw_next = response.get("nextPageToken")
        if raw_next in (None, ""):
            return matches
        if (
            not isinstance(raw_next, str)
            or len(raw_next) > 4096
            or any(ord(character) < 32 or ord(character) == 127 for character in raw_next)
            or raw_next in seen_tokens
        ):
            raise GoogleClassroomAdminError(
                "Google Classroom duplicate preflight returned invalid pagination"
            )
        seen_tokens.add(raw_next)
        page_token = raw_next
    raise GoogleClassroomAdminError(
        "Google Classroom duplicate preflight exceeded the page limit"
    )


def _existing_agent_safe_draft(
    service: Any,
    *,
    course_id: str,
    assignment: Mapping[str, Any],
) -> Optional[Dict[str, Any]]:
    matches = _find_same_title_coursework(
        service,
        course_id=course_id,
        title=str(assignment["title"]),
    )
    if not matches:
        return None
    if len(matches) != 1:
        raise GoogleClassroomAdminError(
            "Multiple same-title coursework items exist; no assignment was created"
        )
    candidate = matches[0]
    if candidate.get("state") != "DRAFT":
        raise GoogleClassroomAdminError(
            "Same-title coursework exists outside DRAFT; no assignment was created"
        )
    course_work_id = candidate.get("id")
    if not isinstance(course_work_id, str) or not _COURSEWORK_ID_RE.fullmatch(
        course_work_id
    ):
        raise GoogleClassroomAdminError(
            "Same-title draft has an invalid coursework ID"
        )
    observed = _read_coursework(
        service,
        course_id=course_id,
        course_work_id=course_work_id,
    )
    mismatches = _coursework_shape_mismatches(
        observed,
        course_id=course_id,
        course_work_id=course_work_id,
        assignment=assignment,
    )
    if mismatches:
        raise GoogleClassroomAdminError(
            "Existing same-title draft does not match approved fields: "
            + ", ".join(mismatches)
        )
    return observed


def _create_and_verify_agent_safe_draft(
    service: Any,
    *,
    course_id: str,
    assignment: Mapping[str, Any],
) -> Dict[str, Any]:
    created = _execute_mutation(
        service.courses().courseWork().create(
            courseId=course_id,
            body=copy.deepcopy(dict(assignment)),
            fields=_AGENT_DRAFT_CREATE_FIELDS,
        ),
        "assignment create",
        course_id,
    )
    raw_course_work_id = created.get("id")
    if (
        not isinstance(raw_course_work_id, str)
        or not _COURSEWORK_ID_RE.fullmatch(raw_course_work_id)
    ):
        raise GoogleClassroomOutcomeUnknownError("assignment create", course_id)
    course_work_id = raw_course_work_id
    if (
        created.get("courseId") != course_id
        or created.get("state") != "DRAFT"
        or created.get("associatedWithDeveloper") is not True
    ):
        raise GoogleClassroomPartialCreateError(
            stage="create-response",
            course_id=course_id,
            course_work_id=course_work_id,
            rubric_created=False,
            intended_release="DRAFT",
        )
    try:
        observed = _read_coursework(
            service,
            course_id=course_id,
            course_work_id=course_work_id,
        )
    except GoogleClassroomAdminError:
        raise GoogleClassroomPartialCreateError(
            stage="read-back",
            course_id=course_id,
            course_work_id=course_work_id,
            rubric_created=False,
            intended_release="DRAFT",
        ) from None
    mismatches = _coursework_shape_mismatches(
        observed,
        course_id=course_id,
        course_work_id=course_work_id,
        assignment=assignment,
    )
    if mismatches:
        raise GoogleClassroomPartialCreateError(
            stage="read-back-fields-" + "-".join(mismatches),
            course_id=course_id,
            course_work_id=course_work_id,
            rubric_created=False,
            intended_release="DRAFT",
        )
    return observed


def _resolve_paths(args: argparse.Namespace) -> CourseworkAuthPaths:
    return resolve_coursework_auth_paths(
        args.account,
        credentials_path=getattr(args, "credentials", None),
        token_path=getattr(args, "token", None),
    )


def _get_canonical_course(service: Any, course_reference: str) -> Dict[str, Any]:
    try:
        course = (
            service.courses()
            .get(id=course_reference, fields=_COURSE_LOOKUP_FIELDS)
            .execute(num_retries=0)
        )
    except Exception:
        raise GoogleClassroomAdminError(
            "Google Classroom course lookup failed"
        ) from None
    if not isinstance(course, Mapping) or not course.get("id"):
        raise GoogleClassroomAdminError(
            "Google Classroom course lookup returned an invalid response"
        )
    canonical_id = str(course["id"])
    if not _COURSE_REFERENCE_RE.fullmatch(canonical_id):
        raise GoogleClassroomAdminError(
            "Google Classroom returned an invalid canonical course ID"
        )
    state = course.get("courseState")
    if state != "ACTIVE":
        displayed_state = "(missing)" if state is None else _terminal_text(state)
        raise GoogleClassroomAdminError(
            f"Target course is not active (state: {displayed_state})"
        )
    return copy.deepcopy(dict(course))


def _complete_loopback_callback(
    port: int,
    *,
    secret_input_fn: Callable[[str], str],
    connection_factory: Callable[..., Any],
) -> None:
    """Deliver a hidden, strictly validated OAuth callback to the local listener."""

    if isinstance(port, bool) or not isinstance(port, int) or not 1 <= port <= 65535:
        raise GoogleClassroomAdminError("Loopback port must be between 1 and 65535")
    try:
        callback_url = secret_input_fn(
            "Paste the new browser redirect URL (input hidden): "
        )
    except (EOFError, StopIteration):
        raise GoogleClassroomAdminError(
            "Loopback callback URL was not provided"
        ) from None
    if (
        not isinstance(callback_url, str)
        or not callback_url
        or len(callback_url) > MAX_LOOPBACK_CALLBACK_URL_CHARS
        or any(ord(char) < 32 or ord(char) == 127 for char in callback_url)
    ):
        raise GoogleClassroomAdminError("Loopback callback URL is invalid")

    try:
        parsed = urlsplit(callback_url)
        parsed_port = parsed.port
        query_pairs = parse_qsl(
            parsed.query,
            keep_blank_values=True,
            max_num_fields=64,
        )
    except ValueError:
        raise GoogleClassroomAdminError("Loopback callback URL is invalid") from None
    if (
        parsed.scheme != "http"
        or parsed.hostname != "127.0.0.1"
        or parsed_port != port
        or parsed.username is not None
        or parsed.password is not None
        or parsed.path != "/"
        or parsed.fragment
    ):
        raise GoogleClassroomAdminError(
            "Loopback callback must match the new http://127.0.0.1:PORT/ redirect"
        )

    values: Dict[str, List[str]] = {}
    for key, value in query_pairs:
        values.setdefault(key, []).append(value)
    if "error" in values:
        raise GoogleClassroomAdminError(
            "Google returned an OAuth error instead of an authorization code"
        )
    if (
        len(values.get("state", [])) != 1
        or len(values.get("code", [])) != 1
        or not values["state"][0]
        or not values["code"][0]
        or len(values["state"][0]) > 1024
        or len(values["code"][0]) > 8192
    ):
        raise GoogleClassroomAdminError(
            "Loopback callback must contain one nonempty state and code"
        )

    target = parsed.path + "?" + parsed.query
    connection = None
    try:
        connection = connection_factory("127.0.0.1", port, timeout=10)
        connection.request(
            "GET",
            target,
            headers={
                "Connection": "close",
                "Host": f"127.0.0.1:{port}",
            },
        )
        response = connection.getresponse()
        body = response.read(MAX_LOOPBACK_RESPONSE_BYTES + 1)
    except (OSError, http.client.HTTPException):
        raise GoogleClassroomAdminError(
            "Loopback callback could not be delivered; keep authorize running and check the port"
        ) from None
    finally:
        if connection is not None:
            connection.close()
    if response.status != 200 or len(body) > MAX_LOOPBACK_RESPONSE_BYTES:
        raise GoogleClassroomAdminError(
            "Loopback listener did not accept the callback"
        )


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="course-gclass-admin",
        description="Google Classroom assignment administration",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    status = subparsers.add_parser(
        "auth-status", help="Inspect isolated coursework credential files offline"
    )
    status.add_argument("--account", required=True)
    status.add_argument("--credentials")
    status.add_argument("--token")
    status.add_argument("--show-paths", action="store_true")

    authorize = subparsers.add_parser(
        "authorize", help="Authorize an isolated coursework OAuth token"
    )
    authorize.add_argument("--account", required=True)
    authorize.add_argument("--credentials")
    authorize.add_argument("--token")
    authorize.add_argument("--replace-token", action="store_true")
    authorize.add_argument("--no-open-browser", action="store_true")

    complete_loopback = subparsers.add_parser(
        "complete-loopback",
        help="Safely deliver a browser redirect to a remote loopback listener",
    )
    complete_loopback.add_argument("--port", required=True, type=int)

    prepare_agent_draft = subparsers.add_parser(
        "prepare-agent-safe-draft",
        help="Prepare a bound approval envelope without mutating Google Classroom",
    )
    prepare_agent_draft.add_argument("--course-id", required=True)
    prepare_agent_draft.add_argument("--spec", required=True)
    prepare_agent_draft.add_argument("--account", required=True)
    prepare_agent_draft.add_argument("--credentials")
    prepare_agent_draft.add_argument("--token")

    create = subparsers.add_parser(
        "create-assignment", help="Preview or create one Google Classroom assignment"
    )
    create.add_argument("--course-id", required=True)
    create.add_argument("--spec", required=True)
    create.add_argument("--account")
    create.add_argument("--credentials")
    create.add_argument("--token")
    create.add_argument("--dry-run", action="store_true")
    create.add_argument("--no-open-browser", action="store_true")
    create.add_argument(
        "-y",
        "--yes",
        action="store_true",
        help="use an existing token to create a DRAFT without prompting; agent mode additionally requires the dedicated safe-draft envelope",
    )
    create.add_argument(
        "--agent-safe-draft",
        action="store_true",
        help="use the separately allowlisted minimal-draft smoke-test path in explicit agent mode",
    )
    create.add_argument(
        "--expect-approval-digest",
        help="full digest emitted by prepare-agent-safe-draft for this exact envelope",
    )
    return parser


def _default_tty_check() -> bool:
    return sys.stdin.isatty() and sys.stdout.isatty()


def main(
    argv: Optional[Sequence[str]] = None,
    *,
    input_fn: Callable[[str], str] = input,
    stdout: Optional[TextIO] = None,
    stderr: Optional[TextIO] = None,
    tty_check: Optional[Callable[[], bool]] = None,
    auth_factory: Optional[Callable[..., Any]] = None,
    create_factory: Optional[Callable[..., Dict[str, Any]]] = None,
    secret_input_fn: Optional[Callable[[str], str]] = None,
    loopback_connection_factory: Optional[Callable[..., Any]] = None,
    now: Optional[datetime] = None,
) -> int:
    """Run the dedicated admin CLI; dependency injection keeps tests offline."""

    output = stdout or sys.stdout
    errors = stderr or sys.stderr
    check_tty = tty_check or _default_tty_check
    authenticate = auth_factory or get_google_classroom_coursework_service
    create_assignment = (
        create_factory or create_google_classroom_assignment_with_rubric
    )
    read_secret = secret_input_fn or getpass.getpass
    make_loopback_connection = loopback_connection_factory or http.client.HTTPConnection
    args = _build_parser().parse_args(argv)

    try:
        if args.command == "auth-status":
            paths = _resolve_paths(args)
            status = get_google_classroom_auth_status(
                paths, show_paths=args.show_paths
            )
            _write_json(output, status)
            ready = bool(
                status.get("token_safe")
                and status.get("token_usable")
                and status.get("scopes_match")
                and status.get("client_ids_match", True)
                and (
                    paths.credentials_path is None
                    or not status.get("credentials_exists")
                    or status.get("credentials_safe")
                )
            )
            return 0 if ready else 1

        if args.command == "authorize":
            _require_interactive_human(check_tty)
            account = normalize_expected_account(args.account)
            paths = _resolve_paths(args)
            token_exists = _path_entry_exists(paths.token_path)
            if token_exists and not args.replace_token:
                raise GoogleClassroomAdminError(
                    "A coursework token already exists; use --replace-token to replace it"
                )
            session = authenticate(
                paths,
                expected_account=account,
                open_browser=not args.no_open_browser,
                force_authorize=True,
                replace_token=args.replace_token,
            )
            _write_json(
                output,
                {
                    "authorized": True,
                    "account": _mask_account(account),
                    "accountFingerprint": paths.account_fingerprint,
                    "clientIdFingerprint": session.client_id_fingerprint,
                    "tokenName": paths.token_path.name,
                },
            )
            return 0

        if args.command == "complete-loopback":
            _require_interactive_human(check_tty)
            _complete_loopback_callback(
                args.port,
                secret_input_fn=read_secret,
                connection_factory=make_loopback_connection,
            )
            _write_json(
                output,
                {"callbackDelivered": True, "port": args.port},
            )
            return 0

        if args.command == "prepare-agent-safe-draft":
            assignment, rubric = load_assignment_spec(Path(args.spec), now=now)
            plan = build_assignment_operation_plan(
                args.course_id,
                assignment,
                rubric,
                dry_run=False,
            )
            account = normalize_expected_account(args.account)
            _require_agent_safe_context(
                account=account,
                course_id=args.course_id,
                assignment=assignment,
                rubric=rubric,
                plan=plan,
                require_yes=False,
            )
            paths = _resolve_paths(args)
            session, course, approval_digest = _prepare_agent_safe_draft_context(
                authenticate=authenticate,
                account=account,
                course_id=args.course_id,
                assignment=assignment,
                rubric=rubric,
                plan=plan,
                paths=paths,
            )
            _write_json(
                output,
                {
                    "prepared": True,
                    "classroomMutation": False,
                    "credentialStateMayChange": True,
                    "account": _mask_account(account),
                    "accountFingerprint": paths.account_fingerprint,
                    "courseId": str(course["id"]),
                    "courseName": _terminal_text(course.get("name") or ""),
                    "title": _terminal_text(assignment["title"]),
                    "releaseStatus": "DRAFT",
                    "materials": 0,
                    "rubric": False,
                    "clientIdFingerprint": session.client_id_fingerprint,
                    "authSource": paths.source,
                    "tokenPathFingerprint": _token_path_fingerprint(
                        paths.token_path
                    ),
                    "operationDigest": plan["operationDigest"],
                    "approvalDigest": approval_digest,
                },
            )
            return 0

        assignment, rubric = load_assignment_spec(Path(args.spec), now=now)
        plan = build_assignment_operation_plan(
            args.course_id,
            assignment,
            rubric,
            dry_run=bool(args.dry_run),
        )
        if args.dry_run:
            if args.agent_safe_draft or args.expect_approval_digest:
                raise GoogleClassroomAdminError(
                    "Use prepare-agent-safe-draft to build an agent approval envelope"
                )
            _write_json(output, plan)
            return 0

        if not args.account:
            raise GoogleClassroomAdminError(
                "--account is required unless --dry-run is used"
            )
        account = normalize_expected_account(args.account)
        if args.agent_safe_draft:
            _require_agent_safe_context(
                account=account,
                course_id=args.course_id,
                assignment=assignment,
                rubric=rubric,
                plan=plan,
                require_yes=True,
                yes=bool(args.yes),
                expected_approval_digest=args.expect_approval_digest,
            )
            paths = _resolve_paths(args)
            session, course, approval_digest = _prepare_agent_safe_draft_context(
                authenticate=authenticate,
                account=account,
                course_id=args.course_id,
                assignment=assignment,
                rubric=rubric,
                plan=plan,
                paths=paths,
            )
            if not hmac.compare_digest(
                approval_digest, str(args.expect_approval_digest)
            ):
                raise GoogleClassroomAdminError(
                    "Agent-safe draft approval digest does not match the authenticated envelope"
                )
            canonical_course_id = str(course["id"])
            with TokenFileLock(paths.token_path):
                if not _path_entry_exists(paths.token_path):
                    raise GoogleClassroomAdminError(
                        "Approved coursework token disappeared before mutation"
                    )
                course_work = _existing_agent_safe_draft(
                    session.service,
                    course_id=canonical_course_id,
                    assignment=assignment,
                )
                created = course_work is None
                if course_work is None:
                    course_work = _create_and_verify_agent_safe_draft(
                        session.service,
                        course_id=canonical_course_id,
                        assignment=assignment,
                    )
            receipt = {
                "created": created,
                "reusedExisting": not created,
                "readBackVerified": True,
                "accountFingerprint": paths.account_fingerprint,
                "courseId": canonical_course_id,
                "courseWorkId": course_work["id"],
                "state": "DRAFT",
                "releaseStatus": "DRAFT",
                "operationDigest": plan["operationDigest"],
                "approvalDigest": approval_digest,
            }
            _write_json(output, receipt)
            return 0

        if args.expect_approval_digest:
            raise GoogleClassroomAdminError(
                "--expect-approval-digest requires --agent-safe-draft"
            )
        _require_create_mode(
            yes=bool(args.yes),
            tty_check=check_tty,
            plan=plan,
        )
        paths = _resolve_paths(args)
        if args.yes and not _path_entry_exists(paths.token_path):
            raise GoogleClassroomAdminError(
                "--yes requires an existing coursework token; run authorize first"
            )
        session = authenticate(
            paths,
            expected_account=account,
            open_browser=False if args.yes else not args.no_open_browser,
            require_existing_token=bool(args.yes),
        )
        course = _get_canonical_course(session.service, args.course_id)
        canonical_course_id = str(course["id"])
        if not args.yes:
            _confirm_creation(
                input_fn,
                account=account,
                course=course,
                assignment=assignment,
                plan=plan,
            )

        result = create_assignment(
            session.service,
            canonical_course_id,
            assignment,
            rubric,
            now=now,
        )
        course_work = result.get("courseWork", {})
        created_course_id = str(course_work.get("courseId") or canonical_course_id)
        if created_course_id != canonical_course_id:
            raise GoogleClassroomAdminError(
                "Assignment response did not match the confirmed course"
            )
        receipt: Dict[str, Any] = {
            "created": True,
            "accountFingerprint": paths.account_fingerprint,
            "courseId": canonical_course_id,
            "courseWorkId": course_work.get("id"),
            "state": course_work.get("state"),
            "releaseStatus": result.get("releaseStatus"),
            "operationDigest": plan["operationDigest"],
        }
        if course_work.get("alternateLink"):
            receipt["alternateLink"] = course_work["alternateLink"]
        created_rubric = result.get("rubric")
        if isinstance(created_rubric, Mapping) and created_rubric.get("id"):
            receipt["rubricId"] = created_rubric["id"]
        _write_json(output, receipt)
        return 0

    except GoogleClassroomOutcomeUnknownError as exc:
        _write_json(
            errors,
            {
                "error": "outcome_unknown",
                "courseId": exc.course_id,
                "message": "The create outcome is unknown; inspect Classroom before rerunning.",
                "status": exc.status,
            },
        )
        return 3
    except GoogleClassroomPartialCreateError as exc:
        _write_json(
            errors,
            {
                "error": "partial_create",
                "stage": exc.stage,
                "courseId": exc.course_id,
                "courseWorkId": exc.course_work_id,
                "alternateLink": exc.alternate_link,
                "rubricCreated": exc.rubric_created,
                "intendedRelease": exc.intended_release,
                "status": exc.status,
            },
        )
        return 4
    except GoogleClassroomAPIError as exc:
        _write_json(
            errors,
            {
                "error": "google_api_error",
                "operation": exc.operation,
                "status": exc.status,
            },
        )
        return 3
    except (
        AssignmentSpecError,
        CourseAgentError,
        CredentialSecurityError,
        GoogleClassroomAdminError,
        ValueError,
    ) as exc:
        _write_json(errors, {"error": "validation", "message": str(exc)})
        return 2
    except Exception:
        _write_json(
            errors,
            {
                "error": "internal_error",
                "message": "Unexpected internal failure; no automatic retry was attempted.",
            },
        )
        return 2


if __name__ == "__main__":
    raise SystemExit(main())


__all__ = [
    "AssignmentSpecError",
    "GoogleClassroomAdminError",
    "MAX_ASSIGNMENT_SPEC_BYTES",
    "MAX_LOOPBACK_CALLBACK_URL_CHARS",
    "MAX_LOOPBACK_RESPONSE_BYTES",
    "_complete_loopback_callback",
    "build_assignment_operation_plan",
    "load_assignment_spec",
    "main",
]
