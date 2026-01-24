# -*- coding: utf-8 -*-
# Google Classroom facade module.

from .gclass_auth import SCOPES, _get_google_classroom_credentials, list_google_classroom_courses, list_google_classroom_students
from .gclass_sync import sync_students_with_google_classroom
from .gclass_grading import grade_google_classroom_assignment_submissions
from .gclass_submissions import download_google_classroom_assignment_submissions
from .gclass_sheets import download_google_sheet_to_csv
from .gclass_unenroll import unenroll_google_classroom_students

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
]
