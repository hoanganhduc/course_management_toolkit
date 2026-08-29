# -*- coding: utf-8 -*-
# Google Classroom facade module.

from .gclass_auth import SCOPES, _get_google_classroom_credentials, list_google_classroom_courses, list_google_classroom_students
from .gclass_sync import sync_students_with_google_classroom
from .gclass_grading import grade_google_classroom_assignment_submissions
from .gclass_submissions import download_google_classroom_assignment_submissions
from .gclass_sheets import download_google_sheet_to_csv
from .gclass_unenroll import unenroll_google_classroom_students
from .gclass_coursework import (
    GoogleClassroomAPIError,
    GoogleClassroomError,
    GoogleClassroomOutcomeUnknownError,
    GoogleClassroomPartialCreateError,
    build_google_classroom_assignment_body,
    build_google_classroom_rubric_body,
    create_google_classroom_assignment,
    create_google_classroom_assignment_with_rubric,
    create_google_classroom_rubric,
    google_classroom_drive_material,
    google_classroom_link_material,
    google_classroom_youtube_material,
)
from .gclass_coursework_auth import (
    COURSEWORK_SCOPES,
    CredentialSecurityError,
    get_google_classroom_auth_status,
    get_google_classroom_coursework_service,
    resolve_coursework_auth_paths,
)

__all__ = [
    "SCOPES",
    "_get_google_classroom_credentials",
    "list_google_classroom_courses",
    "list_google_classroom_students",
    "sync_students_with_google_classroom",
    "grade_google_classroom_assignment_submissions",
    "download_google_classroom_assignment_submissions",
    "download_google_sheet_to_csv",
    "unenroll_google_classroom_students",
    "COURSEWORK_SCOPES",
    "CredentialSecurityError",
    "GoogleClassroomAPIError",
    "GoogleClassroomError",
    "GoogleClassroomOutcomeUnknownError",
    "GoogleClassroomPartialCreateError",
    "build_google_classroom_assignment_body",
    "build_google_classroom_rubric_body",
    "create_google_classroom_assignment",
    "create_google_classroom_assignment_with_rubric",
    "create_google_classroom_rubric",
    "get_google_classroom_auth_status",
    "get_google_classroom_coursework_service",
    "google_classroom_drive_material",
    "google_classroom_link_material",
    "google_classroom_youtube_material",
    "resolve_coursework_auth_paths",
]
