# -*- coding: utf-8 -*-
"""Google Classroom assignment builders and mutation orchestration.

The builders are dependency-free. API functions accept a discovery-style service
object so tests can run without Google packages or credentials.
"""

from __future__ import annotations

import copy
import json
import math
import re
from datetime import datetime, timedelta, timezone
from typing import Any, Callable, Dict, Mapping, Optional, Sequence, Tuple, Union
from urllib.parse import urlsplit

from .course_agent_common import CourseAgentError, is_agent_mode


ASSIGNMENT_STATES = {"DRAFT", "PUBLISHED"}
ASSIGNEE_MODES = {"ALL_STUDENTS", "INDIVIDUAL_STUDENTS"}
SUBMISSION_MODIFICATION_MODES = {"MODIFIABLE_UNTIL_TURNED_IN", "MODIFIABLE"}
DRIVE_SHARE_MODES = {"VIEW", "EDIT", "STUDENT_COPY"}
GRADING_PERIOD_MODES = {"AUTO", "NONE", "EXPLICIT"}
MAX_MATERIALS = 20
SCHEDULE_MIN_LEAD_SECONDS = 60

_OPAQUE_ID_RE = re.compile(r"^[A-Za-z0-9_-]{1,256}$")
_COURSE_REFERENCE_RE = re.compile(r"^(?:[dp]:)?[A-Za-z0-9_-]+$")
_OUTPUT_ONLY_COURSEWORK_FIELDS = {
    "courseId",
    "id",
    "alternateLink",
    "creationTime",
    "updateTime",
    "associatedWithDeveloper",
    "creatorUserId",
    "gradeCategory",
    "previewVersion",
    "learningGoals",
    "assignment",
    "multipleChoiceQuestion",
}


class GoogleClassroomError(RuntimeError):
    """Base error for the isolated coursework creation surface."""


class GoogleClassroomAPIError(GoogleClassroomError):
    """A definite, sanitized Google API failure."""

    def __init__(self, operation: str, *, status: Optional[int] = None):
        self.operation = operation
        self.status = status
        suffix = f" (HTTP {status})" if status is not None else ""
        super().__init__(f"Google Classroom {operation} failed{suffix}")


class GoogleClassroomOutcomeUnknownError(GoogleClassroomError):
    """A create request may have succeeded, but no resource ID is known."""

    def __init__(
        self,
        operation: str,
        course_id: str,
        *,
        status: Optional[int] = None,
    ):
        self.operation = operation
        self.course_id = course_id
        self.status = status
        super().__init__(
            f"Google Classroom {operation} outcome is unknown; verify the course before rerunning"
        )


class GoogleClassroomPartialCreateError(GoogleClassroomError):
    """Coursework exists but a rubric or release stage did not complete."""

    def __init__(
        self,
        *,
        stage: str,
        course_id: str,
        course_work_id: str,
        alternate_link: Optional[str] = None,
        rubric_created: Optional[bool] = None,
        intended_release: str = "DRAFT",
        observed_course_work: Optional[Mapping[str, Any]] = None,
        status: Optional[int] = None,
    ):
        self.stage = stage
        self.course_id = course_id
        self.course_work_id = course_work_id
        self.alternate_link = alternate_link
        self.rubric_created = rubric_created
        self.intended_release = intended_release
        self.observed_course_work = (
            copy.deepcopy(dict(observed_course_work)) if observed_course_work else None
        )
        self.status = status
        super().__init__(
            f"Google Classroom assignment {course_work_id} remains after {stage} failure"
        )


def _reject_agent_mode() -> None:
    if is_agent_mode():
        raise CourseAgentError(
            "create-assignment is not available on the agent surface",
            code="agent_forbidden",
        )


def _validate_utf8_text(
    value: Any,
    label: str,
    *,
    minimum: int = 0,
    maximum: Optional[int] = None,
    allow_newlines: bool = False,
    require_non_whitespace: bool = False,
) -> str:
    if not isinstance(value, str):
        raise ValueError(f"{label} must be a string")
    try:
        value.encode("utf-8")
    except UnicodeEncodeError as exc:
        raise ValueError(f"{label} must be valid UTF-8") from exc
    if len(value) < minimum or (maximum is not None and len(value) > maximum):
        bounds = f"between {minimum} and {maximum}" if maximum is not None else f"at least {minimum}"
        raise ValueError(f"{label} length must be {bounds}")
    if require_non_whitespace and not value.strip():
        raise ValueError(f"{label} must contain a non-whitespace character")
    allowed = {9, 10, 13} if allow_newlines else set()
    if any((ord(char) < 32 and ord(char) not in allowed) or ord(char) == 127 for char in value):
        raise ValueError(f"{label} contains unsafe control characters")
    return value


def _validate_opaque_id(value: Any, label: str) -> str:
    value = _validate_utf8_text(value, label, minimum=1, maximum=256)
    if not _OPAQUE_ID_RE.fullmatch(value):
        raise ValueError(f"{label} contains unsupported characters")
    return value


def google_classroom_drive_material(
    file_id: str,
    share_mode: str = "VIEW",
) -> Dict[str, Any]:
    file_id = _validate_opaque_id(file_id, "Drive file ID")
    share_mode = str(share_mode).upper()
    if share_mode not in DRIVE_SHARE_MODES:
        raise ValueError(f"Unsupported Drive share mode: {share_mode}")
    return {
        "driveFile": {
            "driveFile": {"id": file_id},
            "shareMode": share_mode,
        }
    }


def google_classroom_link_material(url: str) -> Dict[str, Any]:
    url = _validate_utf8_text(url, "Link URL", minimum=1, maximum=4096)
    parsed = urlsplit(url)
    if parsed.scheme.lower() != "https" or not parsed.hostname:
        raise ValueError("Link URL must be an absolute HTTPS URL")
    if parsed.username is not None or parsed.password is not None:
        raise ValueError("Link URL must not contain embedded credentials")
    return {"link": {"url": url}}


def google_classroom_youtube_material(video_id: str) -> Dict[str, Any]:
    return {"youtubeVideo": {"id": _validate_opaque_id(video_id, "YouTube video ID")}}


def _validate_material(material: Mapping[str, Any]) -> Dict[str, Any]:
    if not isinstance(material, Mapping):
        raise ValueError("Each material must be an object")
    if len(material) != 1:
        raise ValueError("Each material must contain exactly one material type")
    kind = next(iter(material))
    value = material[kind]
    if kind == "driveFile":
        if not isinstance(value, Mapping) or set(value) != {"driveFile", "shareMode"}:
            raise ValueError("Drive material must contain driveFile and shareMode only")
        details = value["driveFile"]
        if not isinstance(details, Mapping) or set(details) != {"id"}:
            raise ValueError("Drive material must provide only the Drive file ID")
        return google_classroom_drive_material(details["id"], value["shareMode"])
    if kind == "link":
        if not isinstance(value, Mapping) or set(value) != {"url"}:
            raise ValueError("Link material must provide only url")
        return google_classroom_link_material(value["url"])
    if kind == "youtubeVideo":
        if not isinstance(value, Mapping) or set(value) != {"id"}:
            raise ValueError("YouTube material must provide only id")
        return google_classroom_youtube_material(value["id"])
    raise ValueError(f"Unsupported or read-only material type: {kind}")


def _coerce_datetime(value: Union[str, datetime], label: str) -> datetime:
    if isinstance(value, str):
        text = value.strip()
        if text.endswith(("Z", "z")):
            text = text[:-1] + "+00:00"
        try:
            value = datetime.fromisoformat(text)
        except ValueError as exc:
            raise ValueError(f"{label} must be an RFC3339 timestamp") from exc
    if not isinstance(value, datetime) or value.tzinfo is None:
        raise ValueError(f"{label} must be timezone-aware")
    try:
        return value.astimezone(timezone.utc)
    except (OverflowError, ValueError) as exc:
        raise ValueError(f"{label} is outside the supported datetime range") from exc


def _utc_now(
    now: Optional[Union[datetime, Callable[[], datetime]]],
) -> datetime:
    if callable(now):
        now = now()
    return _coerce_datetime(now or datetime.now(timezone.utc), "now")


def _rfc3339_utc(value: datetime) -> str:
    value = value.astimezone(timezone.utc)
    text = value.isoformat(timespec="microseconds" if value.microsecond else "seconds")
    return text.replace("+00:00", "Z")


def _validate_scheduled_timestamp(
    value: Union[str, datetime],
    *,
    now: Optional[datetime] = None,
    minimum_lead_seconds: int = SCHEDULE_MIN_LEAD_SECONDS,
) -> datetime:
    scheduled = _coerce_datetime(value, "scheduled_at")
    if scheduled < _utc_now(now) + timedelta(seconds=minimum_lead_seconds):
        raise ValueError(
            f"scheduled_at must be at least {minimum_lead_seconds} seconds in the future"
        )
    return scheduled


def build_google_classroom_assignment_body(
    title: str,
    *,
    description: Optional[str] = None,
    state: str = "DRAFT",
    scheduled_at: Optional[Union[str, datetime]] = None,
    due_at: Optional[Union[str, datetime]] = None,
    max_points: Optional[int] = None,
    materials: Optional[Sequence[Mapping[str, Any]]] = None,
    assignee_mode: str = "ALL_STUDENTS",
    individual_student_ids: Optional[Sequence[str]] = None,
    submission_modification_mode: str = "MODIFIABLE_UNTIL_TURNED_IN",
    topic_id: Optional[str] = None,
    grading_period_mode: str = "AUTO",
    grading_period_id: Optional[str] = None,
    now: Optional[datetime] = None,
) -> Dict[str, Any]:
    """Build and validate a stable REST v1 ``ASSIGNMENT`` request body."""

    now_utc = _utc_now(now)
    title = _validate_utf8_text(
        title,
        "title",
        minimum=1,
        maximum=3000,
        require_non_whitespace=True,
    )
    state = str(state).upper()
    if state not in ASSIGNMENT_STATES:
        raise ValueError(f"Unsupported assignment state: {state}")
    assignee_mode = str(assignee_mode).upper()
    if assignee_mode not in ASSIGNEE_MODES:
        raise ValueError(f"Unsupported assignee mode: {assignee_mode}")
    submission_modification_mode = str(submission_modification_mode).upper()
    if submission_modification_mode not in SUBMISSION_MODIFICATION_MODES:
        raise ValueError(
            f"Unsupported submission modification mode: {submission_modification_mode}"
        )

    body: Dict[str, Any] = {
        "title": title,
        "workType": "ASSIGNMENT",
        "state": state,
        "assigneeMode": assignee_mode,
        "submissionModificationMode": submission_modification_mode,
    }
    if description is not None:
        body["description"] = _validate_utf8_text(
            description,
            "description",
            maximum=30000,
            allow_newlines=True,
        )

    scheduled: Optional[datetime] = None
    if scheduled_at is not None:
        if state != "DRAFT":
            raise ValueError("scheduled_at can only be used with requested state DRAFT")
        scheduled = _validate_scheduled_timestamp(scheduled_at, now=now_utc)
        body["scheduledTime"] = _rfc3339_utc(scheduled)

    if due_at is not None:
        due = _coerce_datetime(due_at, "due_at")
        if due <= now_utc:
            raise ValueError("due_at must be in the future")
        if scheduled is not None and due <= scheduled:
            raise ValueError("due_at must be later than scheduled_at")
        body["dueDate"] = {"year": due.year, "month": due.month, "day": due.day}
        due_time: Dict[str, int] = {
            "hours": due.hour,
            "minutes": due.minute,
            "seconds": due.second,
        }
        if due.microsecond:
            due_time["nanos"] = due.microsecond * 1000
        body["dueTime"] = due_time

    if max_points is not None:
        if isinstance(max_points, bool) or not isinstance(max_points, int) or max_points < 0:
            raise ValueError("max_points must be a nonnegative integer")
        body["maxPoints"] = max_points

    if materials is not None:
        if not isinstance(materials, Sequence) or isinstance(materials, (str, bytes)):
            raise ValueError("materials must be a sequence")
        if len(materials) > MAX_MATERIALS:
            raise ValueError(f"Coursework supports at most {MAX_MATERIALS} materials")
        validated = [_validate_material(item) for item in materials]
        canonical = [
            json.dumps(item, sort_keys=True, separators=(",", ":"), ensure_ascii=False)
            for item in validated
        ]
        if len(set(canonical)) != len(canonical):
            raise ValueError("Duplicate materials are not allowed")
        if validated:
            body["materials"] = validated

    if individual_student_ids is None:
        ids = []
    else:
        if not isinstance(individual_student_ids, Sequence) or isinstance(
            individual_student_ids, (str, bytes)
        ):
            raise ValueError("individual_student_ids must be a sequence")
        ids = list(individual_student_ids)
    if assignee_mode == "INDIVIDUAL_STUDENTS":
        if not ids:
            raise ValueError("individual_student_ids are required for individual assignment")
        validated_ids = [_validate_opaque_id(value, "student ID") for value in ids]
        if len(set(validated_ids)) != len(validated_ids):
            raise ValueError("Duplicate student IDs are not allowed")
        body["individualStudentsOptions"] = {"studentIds": validated_ids}
    elif ids:
        raise ValueError("individual_student_ids require INDIVIDUAL_STUDENTS mode")

    if topic_id is not None:
        body["topicId"] = _validate_opaque_id(topic_id, "topic ID")

    grading_period_mode = str(grading_period_mode).upper()
    if grading_period_mode not in GRADING_PERIOD_MODES:
        raise ValueError(f"Unsupported grading period mode: {grading_period_mode}")
    if grading_period_mode == "AUTO":
        if grading_period_id not in (None, ""):
            raise ValueError("AUTO grading period mode must not include an ID")
    elif grading_period_mode == "NONE":
        if grading_period_id not in (None, ""):
            raise ValueError("NONE grading period mode must not include an ID")
        body["gradingPeriodId"] = ""
    else:
        if grading_period_id is None:
            raise ValueError("EXPLICIT grading period mode requires an ID")
        body["gradingPeriodId"] = _validate_opaque_id(
            grading_period_id, "grading period ID"
        )
    return body


def _copy_optional_text(source: Mapping[str, Any], target: Dict[str, Any], key: str) -> None:
    if key in source and source[key] is not None:
        target[key] = _validate_utf8_text(
            source[key], key, maximum=30000, allow_newlines=True
        )


def build_google_classroom_rubric_body(
    criteria: Sequence[Mapping[str, Any]],
    *,
    scoring_mode: str,
) -> Dict[str, Any]:
    """Build an inline rubric body; spreadsheet-backed rubrics are not accepted."""

    if not isinstance(criteria, Sequence) or isinstance(criteria, (str, bytes)):
        raise ValueError("criteria must be a sequence")
    if not 1 <= len(criteria) <= 50:
        raise ValueError("A rubric must contain between 1 and 50 criteria")
    scoring_mode = str(scoring_mode).upper()
    if scoring_mode not in {"SCORED", "UNSCORED"}:
        raise ValueError("scoring_mode must be SCORED or UNSCORED")

    output: list = []
    for criterion in criteria:
        if not isinstance(criterion, Mapping):
            raise ValueError("Each criterion must be an object")
        allowed = {"title", "description", "levels"}
        unknown = set(criterion) - allowed
        if unknown:
            raise ValueError(f"Unsupported rubric criterion fields: {sorted(unknown)}")
        levels = criterion.get("levels")
        if not isinstance(levels, Sequence) or isinstance(levels, (str, bytes)):
            raise ValueError("Each criterion must contain levels")
        if not 1 <= len(levels) <= 10:
            raise ValueError("Each criterion must contain between 1 and 10 levels")
        built_criterion: Dict[str, Any] = {}
        _copy_optional_text(criterion, built_criterion, "title")
        _copy_optional_text(criterion, built_criterion, "description")
        built_levels = []
        points = []
        for level in levels:
            if not isinstance(level, Mapping):
                raise ValueError("Each rubric level must be an object")
            unknown_level = set(level) - {"title", "description", "points"}
            if unknown_level:
                raise ValueError(f"Unsupported rubric level fields: {sorted(unknown_level)}")
            built_level: Dict[str, Any] = {}
            _copy_optional_text(level, built_level, "title")
            _copy_optional_text(level, built_level, "description")
            if scoring_mode == "SCORED":
                if "points" not in level:
                    raise ValueError("Every scored rubric level must specify points")
                value = level["points"]
                if isinstance(value, bool) or not isinstance(value, (int, float)):
                    raise ValueError("Rubric points must be finite and nonnegative")
                try:
                    normalized_points = float(value)
                except (OverflowError, TypeError, ValueError):
                    raise ValueError(
                        "Rubric points must be finite and nonnegative"
                    ) from None
                if not math.isfinite(normalized_points) or value < 0:
                    raise ValueError("Rubric points must be finite and nonnegative")
                built_level["points"] = value
                points.append(normalized_points)
            else:
                if "points" in level:
                    raise ValueError("Unscored rubric levels must omit points")
                if not built_level.get("title", "").strip():
                    raise ValueError("Unscored rubric levels require a title")
            built_levels.append(built_level)
        if scoring_mode == "SCORED":
            if len(set(points)) != len(points):
                raise ValueError("Rubric points must be unique within each criterion")
            if len(points) > 2:
                increasing = all(a < b for a, b in zip(points, points[1:]))
                decreasing = all(a > b for a, b in zip(points, points[1:]))
                if not (increasing or decreasing):
                    raise ValueError("Rubric levels must be ordered by points")
        built_criterion["levels"] = built_levels
        output.append(built_criterion)
    if (
        scoring_mode == "SCORED"
        and len(output) == 1
        and len(output[0]["levels"]) == 1
        and output[0]["levels"][0].get("points") == 0
    ):
        raise ValueError(
            "A one-criterion, one-level scored rubric cannot have zero points"
        )
    return {"criteria": output}


def _status_from_exception(exc: BaseException) -> Optional[int]:
    response = getattr(exc, "resp", None)
    status = getattr(response, "status", None)
    if status is None:
        status = getattr(exc, "status_code", None)
    try:
        return int(status) if status is not None else None
    except (TypeError, ValueError):
        return None


def _is_ambiguous_exception(exc: BaseException) -> bool:
    status = _status_from_exception(exc)
    return status is None or status in {408, 429} or status >= 500


def _execute_mutation(request: Any, operation: str, course_id: str) -> Dict[str, Any]:
    try:
        response = request.execute(num_retries=0)
    except Exception as exc:
        status = _status_from_exception(exc)
        if _is_ambiguous_exception(exc):
            raise GoogleClassroomOutcomeUnknownError(
                operation, course_id, status=status
            ) from None
        raise GoogleClassroomAPIError(operation, status=status) from None
    if not isinstance(response, Mapping):
        raise GoogleClassroomOutcomeUnknownError(operation, course_id)
    return copy.deepcopy(dict(response))


def _resolve_canonical_course_id(service: Any, course_id: Any) -> str:
    """Resolve supported Classroom aliases before any mutation request."""

    reference = _validate_utf8_text(
        course_id,
        "course ID",
        minimum=1,
        maximum=256,
    )
    if not _COURSE_REFERENCE_RE.fullmatch(reference):
        raise ValueError("course ID must be a Classroom ID or d:/p: alias")
    if not reference.startswith(("d:", "p:")):
        return reference
    try:
        response = (
            service.courses()
            .get(id=reference, fields="id")
            .execute(num_retries=0)
        )
    except Exception as exc:
        raise GoogleClassroomAPIError(
            "course lookup", status=_status_from_exception(exc)
        ) from None
    if not isinstance(response, Mapping):
        raise GoogleClassroomAPIError("course lookup")
    canonical = response.get("id")
    if not isinstance(canonical, str) or not _OPAQUE_ID_RE.fullmatch(canonical):
        raise GoogleClassroomAPIError("course lookup")
    return canonical


def create_google_classroom_assignment(
    service: Any,
    course_id: str,
    body: Mapping[str, Any],
) -> Dict[str, Any]:
    """Create one assignment with no automatic retry."""

    _reject_agent_mode()
    if not isinstance(body, Mapping) or body.get("workType") != "ASSIGNMENT":
        raise ValueError("body must be a built ASSIGNMENT request")
    forbidden = set(body) & _OUTPUT_ONLY_COURSEWORK_FIELDS
    if forbidden:
        raise ValueError(f"Request contains output-only fields: {sorted(forbidden)}")
    canonical_course_id = _resolve_canonical_course_id(service, course_id)
    return _execute_mutation(
        service.courses().courseWork().create(
            courseId=canonical_course_id, body=copy.deepcopy(dict(body))
        ),
        "assignment create",
        canonical_course_id,
    )


def create_google_classroom_rubric(
    service: Any,
    course_id: str,
    course_work_id: str,
    body: Mapping[str, Any],
) -> Dict[str, Any]:
    _reject_agent_mode()
    canonical_course_id = _resolve_canonical_course_id(service, course_id)
    return _execute_mutation(
        service.courses().courseWork().rubrics().create(
            courseId=canonical_course_id,
            courseWorkId=str(course_work_id),
            body=copy.deepcopy(dict(body)),
        ),
        "rubric create",
        canonical_course_id,
    )


def _intended_release(body: Mapping[str, Any]) -> str:
    if body.get("scheduledTime"):
        return "SCHEDULED"
    return "PUBLISHED" if body.get("state") == "PUBLISHED" else "DRAFT"


def _scheduled_times_match(left: Any, right: Any) -> bool:
    try:
        return _coerce_datetime(left, "scheduledTime") == _coerce_datetime(
            right, "scheduledTime"
        )
    except (TypeError, ValueError):
        return False


def _release_response_matches(
    response: Mapping[str, Any],
    intended: str,
    desired: Mapping[str, Any],
) -> bool:
    if intended == "PUBLISHED":
        return response.get("state") == "PUBLISHED"
    if intended == "SCHEDULED":
        return _scheduled_times_match(
            response.get("scheduledTime"), desired.get("scheduledTime")
        )
    return response.get("state") == "DRAFT"


def _course_work_response_matches_parent(
    response: Mapping[str, Any],
    course_id: str,
    course_work_id: Optional[str] = None,
) -> bool:
    """Require output-only parent identifiers to match the requested resource."""

    if str(response.get("courseId") or "") != str(course_id):
        return False
    if course_work_id is not None and str(response.get("id") or "") != str(
        course_work_id
    ):
        return False
    return True


def _normalized_rubric_criteria(value: Any) -> Optional[list]:
    """Project a Rubric response onto the writable criterion/level fields."""

    if not isinstance(value, Sequence) or isinstance(value, (str, bytes)):
        return None
    criteria = []
    for criterion in value:
        if not isinstance(criterion, Mapping):
            return None
        levels = criterion.get("levels")
        if not isinstance(levels, Sequence) or isinstance(levels, (str, bytes)):
            return None
        normalized_criterion: Dict[str, Any] = {}
        for key in ("title", "description"):
            if criterion.get(key) not in (None, ""):
                normalized_criterion[key] = criterion[key]
        normalized_levels = []
        for level in levels:
            if not isinstance(level, Mapping):
                return None
            normalized_level: Dict[str, Any] = {}
            for key in ("title", "description"):
                if level.get(key) not in (None, ""):
                    normalized_level[key] = level[key]
            if "points" in level:
                normalized_level["points"] = level["points"]
            normalized_levels.append(normalized_level)
        normalized_criterion["levels"] = normalized_levels
        criteria.append(normalized_criterion)
    return criteria


def _rubric_response_matches(
    response: Mapping[str, Any],
    course_id: str,
    course_work_id: str,
    requested: Mapping[str, Any],
) -> bool:
    return bool(
        response.get("id")
        and str(response.get("courseId") or "") == str(course_id)
        and str(response.get("courseWorkId") or "") == str(course_work_id)
        and _normalized_rubric_criteria(response.get("criteria"))
        == _normalized_rubric_criteria(requested.get("criteria"))
    )


def _due_datetime_from_body(body: Mapping[str, Any]) -> Optional[datetime]:
    due_date = body.get("dueDate")
    due_time = body.get("dueTime")
    if due_date is None and due_time is None:
        return None
    if not isinstance(due_date, Mapping) or not isinstance(due_time, Mapping):
        raise ValueError("Built assignment must contain both dueDate and dueTime")
    try:
        nanos = int(due_time.get("nanos", 0))
        if nanos < 0 or nanos >= 1_000_000_000 or nanos % 1000:
            raise ValueError
        return datetime(
            int(due_date["year"]),
            int(due_date["month"]),
            int(due_date["day"]),
            int(due_time.get("hours", 0)),
            int(due_time.get("minutes", 0)),
            int(due_time.get("seconds", 0)),
            nanos // 1000,
            tzinfo=timezone.utc,
        )
    except (KeyError, TypeError, ValueError, OverflowError):
        raise ValueError("Built assignment contains an invalid UTC due date/time") from None


def _revalidate_assignment_timing(
    body: Mapping[str, Any],
    *,
    now: Optional[Union[datetime, Callable[[], datetime]]],
    schedule_min_lead_seconds: int,
) -> None:
    """Recheck time-sensitive constraints immediately before a write."""

    now_utc = _utc_now(now)
    scheduled = None
    if body.get("scheduledTime") is not None:
        scheduled = _validate_scheduled_timestamp(
            body["scheduledTime"],
            now=now_utc,
            minimum_lead_seconds=schedule_min_lead_seconds,
        )
    due = _due_datetime_from_body(body)
    if due is not None:
        if due <= now_utc:
            raise ValueError("due_at must still be in the future when creation starts")
        if scheduled is not None and due <= scheduled:
            raise ValueError("due_at must be later than scheduled_at")


def _partial_error(
    *,
    stage: str,
    course_work: Mapping[str, Any],
    course_id: str,
    intended_release: str,
    rubric_created: Optional[bool],
    observed: Optional[Mapping[str, Any]] = None,
    status: Optional[int] = None,
) -> GoogleClassroomPartialCreateError:
    return GoogleClassroomPartialCreateError(
        stage=stage,
        course_id=course_id,
        course_work_id=str(course_work.get("id", "unknown")),
        alternate_link=course_work.get("alternateLink"),
        rubric_created=rubric_created,
        intended_release=intended_release,
        observed_course_work=observed,
        status=status,
    )


def create_google_classroom_assignment_with_rubric(
    service: Any,
    course_id: str,
    assignment_body: Mapping[str, Any],
    rubric_body: Optional[Mapping[str, Any]] = None,
    *,
    schedule_min_lead_seconds: int = SCHEDULE_MIN_LEAD_SECONDS,
    now: Optional[Union[datetime, Callable[[], datetime]]] = None,
) -> Dict[str, Any]:
    """Create coursework, optionally attach a rubric, then publish or schedule it."""

    _reject_agent_mode()
    desired = copy.deepcopy(dict(assignment_body))
    canonical_course_id = _resolve_canonical_course_id(service, course_id)
    _revalidate_assignment_timing(
        desired,
        now=now,
        schedule_min_lead_seconds=schedule_min_lead_seconds,
    )
    intended = _intended_release(desired)
    if rubric_body is None:
        created = create_google_classroom_assignment(
            service, canonical_course_id, desired
        )
        if not created.get("id"):
            raise GoogleClassroomOutcomeUnknownError(
                "assignment create", canonical_course_id
            )
        if not _course_work_response_matches_parent(
            created, canonical_course_id
        ) or not _release_response_matches(created, intended, desired):
            raise _partial_error(
                stage="create-response",
                course_work=created,
                course_id=canonical_course_id,
                intended_release=intended,
                rubric_created=False,
                observed=created,
            )
        return {"courseWork": created, "rubric": None, "releaseStatus": intended}

    patch_body: Optional[Dict[str, Any]] = None
    update_mask: Optional[str] = None
    if intended == "PUBLISHED":
        patch_body = {"state": "PUBLISHED"}
        update_mask = "state"
    elif intended == "SCHEDULED":
        scheduled = _coerce_datetime(desired["scheduledTime"], "scheduledTime")
        patch_body = {"scheduledTime": _rfc3339_utc(scheduled)}
        update_mask = "scheduledTime"

    initial = copy.deepcopy(desired)
    initial["state"] = "DRAFT"
    initial.pop("scheduledTime", None)
    created = create_google_classroom_assignment(
        service, canonical_course_id, initial
    )
    course_work_id = created.get("id")
    if not course_work_id:
        raise GoogleClassroomOutcomeUnknownError(
            "assignment create", canonical_course_id
        )
    course_work_id = str(course_work_id)
    if (
        not _course_work_response_matches_parent(created, canonical_course_id)
        or created.get("state") != "DRAFT"
        or created.get("scheduledTime")
    ):
        raise _partial_error(
            stage="create-response",
            course_work=created,
            course_id=canonical_course_id,
            intended_release=intended,
            rubric_created=False,
            observed=created,
        )

    rubric: Optional[Dict[str, Any]] = None
    try:
        rubric = create_google_classroom_rubric(
            service,
            canonical_course_id,
            course_work_id,
            rubric_body,
        )
    except GoogleClassroomOutcomeUnknownError as exc:
        try:
            listed = (
                service.courses()
                .courseWork()
                .rubrics()
                .list(
                    courseId=canonical_course_id,
                    courseWorkId=course_work_id,
                    pageSize=1,
                )
                .execute(num_retries=0)
            )
            if not isinstance(listed, Mapping) or listed.get("nextPageToken"):
                raise ValueError("Rubric reconciliation response was not singular")
            rubrics = listed.get("rubrics", [])
        except Exception:
            raise _partial_error(
                stage="rubric",
                course_work=created,
                course_id=canonical_course_id,
                intended_release=intended,
                rubric_created=None,
                status=exc.status,
            ) from None
        if (
            not isinstance(rubrics, Sequence)
            or isinstance(rubrics, (str, bytes))
            or len(rubrics) != 1
        ):
            raise _partial_error(
                stage="rubric",
                course_work=created,
                course_id=canonical_course_id,
                intended_release=intended,
                rubric_created=None,
                status=exc.status,
            ) from None
        if not isinstance(rubrics[0], Mapping):
            raise _partial_error(
                stage="rubric",
                course_work=created,
                course_id=canonical_course_id,
                intended_release=intended,
                rubric_created=None,
                status=exc.status,
            ) from None
        rubric = copy.deepcopy(dict(rubrics[0]))
    except GoogleClassroomAPIError as exc:
        raise _partial_error(
            stage="rubric",
            course_work=created,
            course_id=canonical_course_id,
            intended_release=intended,
            rubric_created=False,
            status=exc.status,
        ) from None

    if not rubric or not _rubric_response_matches(
        rubric,
        canonical_course_id,
        course_work_id,
        rubric_body,
    ):
        raise _partial_error(
            stage="rubric",
            course_work=created,
            course_id=canonical_course_id,
            intended_release=intended,
            rubric_created=None,
        )

    latest = created
    if patch_body is not None and update_mask is not None:
        try:
            _revalidate_assignment_timing(
                desired,
                now=now,
                schedule_min_lead_seconds=schedule_min_lead_seconds,
            )
        except ValueError:
            raise _partial_error(
                stage="release-timing",
                course_work=created,
                course_id=canonical_course_id,
                intended_release=intended,
                rubric_created=True,
            ) from None
        try:
            latest = _execute_mutation(
                service.courses().courseWork().patch(
                    courseId=canonical_course_id,
                    id=course_work_id,
                    updateMask=update_mask,
                    body=patch_body,
                ),
                "assignment release",
                canonical_course_id,
            )
        except GoogleClassroomOutcomeUnknownError as exc:
            try:
                observed = (
                    service.courses()
                    .courseWork()
                    .get(courseId=canonical_course_id, id=course_work_id)
                    .execute(num_retries=0)
                )
            except Exception:
                raise _partial_error(
                    stage="release",
                    course_work=created,
                    course_id=canonical_course_id,
                    intended_release=intended,
                    rubric_created=True,
                    status=exc.status,
                ) from None
            if not isinstance(observed, Mapping):
                raise _partial_error(
                    stage="release",
                    course_work=created,
                    course_id=canonical_course_id,
                    intended_release=intended,
                    rubric_created=True,
                    status=exc.status,
                ) from None
            matched = _course_work_response_matches_parent(
                observed, canonical_course_id, course_work_id
            ) and (
                observed.get("state") == "PUBLISHED"
                if update_mask == "state"
                else _scheduled_times_match(
                    observed.get("scheduledTime"), patch_body["scheduledTime"]
                )
            )
            if not matched:
                raise _partial_error(
                    stage="release",
                    course_work=created,
                    course_id=canonical_course_id,
                    intended_release=intended,
                    rubric_created=True,
                    observed=observed,
                    status=exc.status,
                ) from None
            latest = copy.deepcopy(dict(observed))
        except GoogleClassroomAPIError as exc:
            raise _partial_error(
                stage="release",
                course_work=created,
                course_id=canonical_course_id,
                intended_release=intended,
                rubric_created=True,
                status=exc.status,
            ) from None

        if not _course_work_response_matches_parent(
            latest, canonical_course_id, course_work_id
        ) or not _release_response_matches(latest, intended, desired):
            raise _partial_error(
                stage="release",
                course_work=created,
                course_id=canonical_course_id,
                intended_release=intended,
                rubric_created=True,
                observed=latest,
            )

    return {"courseWork": latest, "rubric": rubric, "releaseStatus": intended}


__all__ = [
    "ASSIGNEE_MODES",
    "ASSIGNMENT_STATES",
    "DRIVE_SHARE_MODES",
    "GRADING_PERIOD_MODES",
    "GoogleClassroomAPIError",
    "GoogleClassroomError",
    "GoogleClassroomOutcomeUnknownError",
    "GoogleClassroomPartialCreateError",
    "SUBMISSION_MODIFICATION_MODES",
    "build_google_classroom_assignment_body",
    "build_google_classroom_rubric_body",
    "create_google_classroom_assignment",
    "create_google_classroom_assignment_with_rubric",
    "create_google_classroom_rubric",
    "google_classroom_drive_material",
    "google_classroom_link_material",
    "google_classroom_youtube_material",
]
