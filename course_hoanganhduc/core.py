# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/
# Course Management Script

import argparse
import json
import os
import sys
import time

from .version import __version__
from .settings import *
from .config import *
from .models import *
from .utils import *
from .data import *
from .canvas import *
from .google_classroom import *


def _build_menu_sections():
    return [
        ("Student Database", [
            ("Import students from Excel or CSV file", "1"),
            ("Save current students to database", "2"),
            ("Load students from database", "3"),
            ("Search for students by keyword (name, student id, email, etc.)", "6"),
            ("Show details of a student", "7"),
            ("Show details of all students", "8"),
            ("Interactively modify the student database", "14"),
        ]),
        ("Student Exports", [
            ("Export student list to Excel file", "4"),
            ("Export all student emails to TXT file", "5"),
            ("Export all student details to TXT file", "9"),
            ("Export student names and emails to TXT file", "46"),
            ("Export student roster to CSV file", "48"),
            ("Update MAT*.xlsx files with grades from database", "37"),
            ("Export final grade distribution", "43"),
        ]),
        ("OCR and PDFs", [
            ("Extract and add blackboard counts from PDF to database", "10"),
            ("Extract handwriting text from PDF to TXT file", "11"),
            ("Print blackboard counts by date for all students", "12"),
            ("Export blackboard counts by date for all students to TXT/Markdown file", "13"),
        ]),
        ("Exams (Multichoice)", [
            ("Extract multiple-choice exam solutions from PDF", "33"),
            ("Extract student answers from scanned exam sheet PDF", "34"),
            ("Evaluate student answers for multiple-choice exam", "35"),
            ("Sync multichoice exam evaluations to Canvas assignment", "38"),
        ]),
        ("Canvas: People and Communication", [
            ("List all assignments on Canvas LMS", "15"),
            ("List all members of a Canvas course", "16"),
            ("Search for a user in Canvas by name or email", "17"),
            ("Download all submission files for a Canvas assignment", "18"),
            ("Add a comment to a Canvas assignment submission", "19"),
            ("Create a Canvas announcement", "20"),
            ("Invite a single user to Canvas course by email", "21"),
            ("Invite multiple users to Canvas course from a TXT file", "22"),
            ("Find and notify students who have not completed required peer reviews", "23"),
            ("Sync Canvas course members to local database", "24"),
            ("Grade Canvas assignment submissions", "25"),
            ("Fetch and reply to Canvas inbox messages", "26"),
            ("List and edit Canvas course pages", "27"),
            ("List students with multiple submissions and only the first submission on time", "28"),
        ]),
        ("Canvas: Rubrics and Grading", [
            ("List all unique rubrics used in Canvas course", "29"),
            ("Export Canvas rubrics to TXT/CSV file", "30"),
            ("Import rubrics to Canvas course from TXT/CSV file", "31"),
            ("Update rubrics for Canvas assignments", "32"),
            ("Export Canvas grading scheme(s) to JSON", "39"),
            ("Add grading scheme to Canvas course from JSON file", "40"),
            ("Check similarities between submissions of the same student for different assignments", "41"),
            ("Send final evaluations to students via Canvas", "42"),
        ]),
        ("Canvas: Admin Tools", [
            ("Change Canvas assignment deadlines", "44"),
            ("Change Canvas assignment lock dates", "47"),
            ("Create Canvas groups of students", "45"),
            ("Delete empty Canvas groups", "50"),
        ]),
        ("Configuration and Integrations", [
            ("Load config from JSON file and save to default location", "36"),
            ("Sync students with Google Classroom", "49"),
        ]),
    ]


def _flatten_menu_sections(sections):
    entries = []
    item_indices = []
    # Menu numbering is presentation-only; action codes map to legacy handlers below.
    display_to_action = {}
    code_to_index = {}
    display = 1
    for section_title, items in sections:
        entries.append({"type": "section", "label": section_title})
        for label, action in items:
            display_code = str(display)
            display += 1
            entries.append({
                "type": "item",
                "label": label,
                "code": display_code,
                "action": action,
            })
            item_indices.append(len(entries) - 1)
            display_to_action[display_code] = action
            code_to_index[display_code] = len(entries) - 1
    return entries, item_indices, display_to_action, code_to_index


def _enable_ansi():
    if os.name != "nt":
        return True
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        handle = kernel32.GetStdHandle(-11)
        mode = ctypes.c_uint()
        if kernel32.GetConsoleMode(handle, ctypes.byref(mode)) == 0:
            return False
        new_mode = mode.value | 0x0004
        if kernel32.SetConsoleMode(handle, new_mode) == 0:
            return False
        return True
    except Exception:
        return False


_ANSI_ENABLED = _enable_ansi()


def _style(text, *codes):
    if not _ANSI_ENABLED or not codes:
        return text
    return f"\x1b[{';'.join(codes)}m{text}\x1b[0m"


def _render_menu(entries, selected_index):
    os.system("cls" if os.name == "nt" else "clear")
    spinner = ["-", "\\", "|", "/"]
    indicator = spinner[int(time.time() * 4) % len(spinner)]
    header = f"Menu {indicator} (use arrow keys, Enter to select, q to quit)"
    if _ANSI_ENABLED:
        header = _style(header, "1", "36")
    print(f"{header}\n")
    for idx, entry in enumerate(entries):
        if entry["type"] == "section":
            label = entry["label"]
            if _ANSI_ENABLED:
                label = _style(label, "1", "36")
            print(label)
            continue
        prefix = ">" if idx == selected_index else " "
        line = f"{prefix} {entry['code']}. {entry['label']}"
        if idx == selected_index:
            line = _style(line, "7")
        print(line)
    print()


def _select_menu_option(sections):
    entries, item_indices, _, code_to_index = _flatten_menu_sections(sections)
    if not item_indices:
        return None
    selected_pos = 0
    selected_index = item_indices[selected_pos]
    # Track fast numeric entry (e.g., "10") to jump the selection.
    digit_buffer = ""
    last_digit_time = 0.0
    try:
        import msvcrt
    except Exception:
        return None

    while True:
        _render_menu(entries, selected_index)
        key = msvcrt.getch()
        now = time.time()
        if key in (b"q", b"Q", b"\x1b"):
            return None
        if key in (b"\x00", b"\xe0"):
            key = msvcrt.getch()
            if key == b"H":  # Up
                selected_pos = (selected_pos - 1) % len(item_indices)
                selected_index = item_indices[selected_pos]
            elif key == b"P":  # Down
                selected_pos = (selected_pos + 1) % len(item_indices)
                selected_index = item_indices[selected_pos]
            elif key == b"G":  # Home
                selected_pos = 0
                selected_index = item_indices[selected_pos]
            elif key == b"O":  # End
                selected_pos = len(item_indices) - 1
                selected_index = item_indices[selected_pos]
            continue
        if key in (b"w", b"W"):
            selected_pos = (selected_pos - 1) % len(item_indices)
            selected_index = item_indices[selected_pos]
            continue
        if key in (b"s", b"S"):
            selected_pos = (selected_pos + 1) % len(item_indices)
            selected_index = item_indices[selected_pos]
            continue
        if key in (b"\x08",):  # Backspace
            digit_buffer = digit_buffer[:-1]
            if digit_buffer in code_to_index:
                selected_index = code_to_index[digit_buffer]
                selected_pos = item_indices.index(selected_index)
            continue
        if b"0" <= key <= b"9":
            if now - last_digit_time > 0.8:
                digit_buffer = ""
            digit_buffer += key.decode("ascii")
            last_digit_time = now
            if digit_buffer in code_to_index:
                selected_index = code_to_index[digit_buffer]
                selected_pos = item_indices.index(selected_index)
            else:
                matches = [code for code in code_to_index if code.startswith(digit_buffer)]
                if not matches:
                    digit_buffer = key.decode("ascii")
                    if digit_buffer in code_to_index:
                        selected_index = code_to_index[digit_buffer]
                        selected_pos = item_indices.index(selected_index)
            continue
        if key in (b"\r", b"\n"):
            return entries[selected_index]["action"]


def _menu_choice_to_action(choice, sections):
    entries, _, display_to_action, _ = _flatten_menu_sections(sections)
    if choice in display_to_action:
        return display_to_action[choice]
    action_codes = {entry["action"] for entry in entries if entry["type"] == "item"}
    if choice in action_codes:
        return choice
    return None


def _print_menu_fallback(sections):
    entries, _, _, _ = _flatten_menu_sections(sections)
    print("\nMenu:")
    for entry in entries:
        if entry["type"] == "section":
            print(f"\n{entry['label']}")
            continue
        print(f"{entry['code']}. {entry['label']}")
    print("\n0. Exit\n")


def main():
    parser = argparse.ArgumentParser(
        description="Student database management. "
                    "Manage students, import/export data, analyze blackboard counts, and extract handwriting from PDFs."
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"course-hoanganhduc {__version__}",
        help="Show package name and version and exit.")

    general_group = parser.add_argument_group("General")
    general_group.add_argument('--verbose', '-v', action='store_true', help="Enable verbose output", dest="verbose")

    config_group = parser.add_argument_group("Configuration")
    config_group.add_argument('--config', '-cfg', '-c', type=str, help="Load config from JSON file and save to default location", dest="config", metavar="CONFIG")
    config_group.add_argument('--course-code', '-ccode', type=str, help="Course code for config folder (e.g., MAT3500)", dest="course_code", metavar="COURSE_CODE")
    config_group.add_argument('--clear-config', '-ccfg', action='store_true',
                              help="Delete stored config.json from the default location",
                              dest="clear_config")
    config_group.add_argument('--clear-credentials', '-ccred', action='store_true',
                              help="Delete stored credentials.json and token.pickle from the default location",
                              dest="clear_credentials")

    db_group = parser.add_argument_group("Student Database")
    db_group.add_argument('--db', '-db', '-D', type=str, help="Database file name (default: students.db, saved in script folder)", dest="db", metavar="DB")
    db_group.add_argument('--add-file', '-a', type=str, help="Import students from Excel or CSV file into the database", dest="add_file", metavar="FILE")
    db_group.add_argument('--save', '-s', action='store_true', help="Save current students to database file", dest="save")
    db_group.add_argument('--load', '-l', action='store_true', help="Load students from database file", dest="load")
    db_group.add_argument('--search', '-S', type=str, help="Search for students by keyword (name, student id, email, etc.)", dest="search", metavar="QUERY")
    db_group.add_argument('--details', '-d', type=str, help="Show details of a student by name, student id, or email", dest="details", metavar="IDENTIFIER")
    db_group.add_argument('--all-details', '-A', action='store_true', help="Show details of all students", dest="all_details")
    db_group.add_argument('--modify', '-m', action='store_true', help="Interactively modify the student database", dest="modify")
    db_group.add_argument('--export-excel', '-x', type=str, help="Export student list to Excel file", dest="export_excel", metavar="EXCEL")
    db_group.add_argument('--export-emails', '-e', type=str, help="Export all student emails to TXT file (avoids duplicates)", dest="export_emails", metavar="EMAILS")
    db_group.add_argument('--export-all-details', '-E', type=str, help="Export all student details to TXT file", dest="export_all_details", metavar="DETAILS")
    db_group.add_argument('--export-emails-and-names', '-en', nargs='?', const="emails_and_names.txt",
                          help="Export all student emails and names to TXT file (default: emails_and_names.txt)",
                          dest="export_emails_and_names", metavar="EMAILS_NAMES")
    db_group.add_argument('--export-final-grade-distribution', nargs='?', const=True,
                          help="Export final grade distribution to a TXT file. Optionally provide output path (default: ./final_grade_distribution.txt).",
                          dest='export_final_grade_distribution')
    db_group.add_argument('--update-mat-excel', '-ume', type=str, nargs='+',
                          help="Update MAT*.xlsx file(s) with grades from database (provide one or more file paths)",
                          dest="update_mat_excel", metavar="MAT_XLSX")
    db_group.add_argument('--export-roster', '-ero', nargs='?', const='classroom_roster.csv',
                          help="Export classroom roster to CSV file (default: classroom_roster.csv)",
                          dest="export_roster", metavar="CSV_FILE")

    ocr_group = parser.add_argument_group("OCR and PDFs")
    ocr_group.add_argument('--add-blackboard-counts', '-b', type=str,
                           help="Extract and add blackboard counts from PDF to database",
                           dest="add_blackboard_counts", metavar="PDF")
    ocr_group.add_argument('--extract-text', '-t', type=str,
                           help="Extract handwriting text from PDF and save to TXT file",
                           dest="extract_text", metavar="PDF")
    ocr_group.add_argument('--print-blackboard-counts', '-p', action='store_true',
                           help="Print blackboard counts by date for all students",
                           dest="print_blackboard_counts")
    ocr_group.add_argument('--export-blackboard-counts', '-B', type=str, nargs='?', const=True,
                           help="Export blackboard counts by date for all students to TXT/Markdown file (use .txt or .md extension, default: TXT)",
                           dest="export_blackboard_counts", metavar="TXT_OR_MD")
    ocr_group.add_argument('--ocr-service', '-O', type=str, choices=['ocrspace', 'tesseract', 'paddleocr'], default='ocrspace',
                           help="OCR service to use for PDF extraction (default: 'ocrspace'). The 'ocrspace' service uses the OCR.space API and works better for handwriting text. The other two services work better for printed text and require additional local installation.",
                           dest="ocr_service")
    ocr_group.add_argument('--ocr-lang', '-L', type=str, default='auto',
                           help="OCR language for PDF extraction (default: auto)",
                           dest="ocr_lang")
    ocr_group.add_argument('--simple-text', '-T', action='store_true',
                           help="Extract simple text (no layout) from PDF OCR",
                           dest="simple_text")
    ocr_group.add_argument('--refine', '-R', type=str, choices=['gemini', 'huggingface'], default=None,
                           help="Refine extracted text using AI ('gemini' or 'huggingface')",
                           dest="refine")

    exam_group = parser.add_argument_group("Exams (Multichoice)")
    exam_group.add_argument('--extract-multichoice-solutions', '-ems', type=str,
                            help="Extract multiple-choice exam solutions from PDF (each page is one sheet code)",
                            dest="extract_multichoice_solutions", metavar="PDF")
    exam_group.add_argument('--extract-multichoice-answers', '-ema', type=str,
                            help="Extract student answers from scanned multi-choice exam sheet PDF",
                            dest="extract_multichoice_answers", metavar="PDF")
    exam_group.add_argument('--evaluate-multichoice-answers', type=str, nargs='?', const=EXAM_TYPE,
                            help="Evaluate student answers for multiple-choice exam (provide exam type: midterm/final, default: global EXAM_TYPE)",
                            dest="evaluate_multichoice_answers", metavar="EXAM_TYPE")
    exam_group.add_argument('--sync-multichoice-evaluations', '-sme', type=str, nargs='?', const=EXAM_TYPE,
                            help="Sync multichoice exam evaluations to Canvas assignment (provide exam type: midterm/final, default: global EXAM_TYPE)",
                            dest="sync_multichoice_evaluations", metavar="EXAM_TYPE")

    canvas_group = parser.add_argument_group("Canvas: People and Communication")
    canvas_group.add_argument('--list-canvas-assignments', action='store_true', help="List all assignments on Canvas LMS", dest="list_canvas_assignments")
    canvas_group.add_argument('--canvas-assignment-category', '-cac', type=str, help="Assignment category (group) to filter when listing Canvas assignments", dest="canvas_assignment_category")
    canvas_group.add_argument('--list-canvas-members', '-cm', action='store_true', help="List all members (teachers, TAs, students) of a Canvas course", dest="list_canvas_members")
    canvas_group.add_argument('--canvas-course-id', '-cc', type=str, help="Canvas course ID (overrides default)", dest="canvas_course_id")
    canvas_group.add_argument('--search-canvas-user', '-cu', type=str, help="Search for a user in Canvas by name or email", dest="search_canvas_user")
    canvas_group.add_argument('--download-canvas-assignment', '-da', nargs='?', const=True, default=None,
                              help="Download all submission files for a Canvas assignment (optionally provide assignment ID)",
                              dest="download_canvas_assignment", metavar="ASSIGNMENT_ID")
    canvas_group.add_argument('--download-dest-dir', '-dd', type=str, help="Destination directory for downloaded Canvas assignment files", dest="download_dest_dir", metavar="DIR")
    canvas_group.add_argument('--comment-canvas-submission', '-cs', action='store_true', help="Add a comment to a Canvas assignment submission", dest="comment_canvas_submission")
    canvas_group.add_argument('--add-canvas-announcement', '-aa', action='store_true', help="Create a new announcement in Canvas course", dest="add_canvas_announcement")
    canvas_group.add_argument('--invite-canvas-email', '-ie', type=str, help="Invite a single user to Canvas course by email", dest="invite_canvas_email")
    canvas_group.add_argument('--invite-canvas-name', type=str, help="Name for Canvas invite (for single user)", dest="invite_canvas_name")
    canvas_group.add_argument('--invite-canvas-role', '-ir', type=str, default="student",
                              help="Role for Canvas invite (student/teacher/ta, default: student)",
                              dest="invite_canvas_role")
    canvas_group.add_argument('--invite-canvas-file', '-if', type=str, help="Invite multiple users to Canvas course from a TXT file or string of pairs/emails", dest="invite_canvas_file")
    canvas_group.add_argument('--notify-incomplete-reviews', '-nr', action='store_true',
                              help="Find and notify students who have not completed required peer reviews for a Canvas assignment",
                              dest="notify_incomplete_reviews")
    canvas_group.add_argument('--review-assignment-id', '-rai', type=str, help="Canvas assignment ID for peer review notification", dest="review_assignment_id")
    canvas_group.add_argument('--sync-canvas', '-sc', action='store_true', help="Sync Canvas course members to local database", dest="sync_canvas")
    canvas_group.add_argument('--grade-canvas-assignment', '-ga', action='store_true', help="Grade Canvas assignment submissions interactively", dest="grade_canvas_assignment")
    canvas_group.add_argument('--fetch-canvas-messages', '-fm', action='store_true', help="Fetch and reply to Canvas inbox messages", dest="fetch_canvas_messages")
    canvas_group.add_argument('--edit-canvas-pages', '-ep', action='store_true', help="List and edit Canvas course pages", dest="edit_canvas_pages")
    canvas_group.add_argument('--list-multiple-submissions-on-time', '-lm', nargs='?', const=None, type=str,
                              help="List students who submitted twice or more to an assignment and the first submission is on time (optionally provide assignment ID)",
                              dest="list_multiple_submissions_on_time", metavar="ASSIGNMENT_ID")

    canvas_rubric_group = parser.add_argument_group("Canvas: Rubrics and Grading")
    canvas_rubric_group.add_argument('--list-canvas-rubrics', '-lr', action='store_true', help="List all unique rubrics used in Canvas course", dest="list_canvas_rubrics")
    canvas_rubric_group.add_argument('--export-canvas-rubrics', '-er', type=str, help="Export Canvas rubrics to TXT/CSV file", dest="export_canvas_rubrics", metavar="RUBRICS_FILE")
    canvas_rubric_group.add_argument('--rubric-assignment-id', '-rid', type=str, help="Assignment ID to filter rubrics", dest="rubric_assignment_id")
    canvas_rubric_group.add_argument('--import-canvas-rubrics', '-imr', type=str, help="Import rubrics from TXT/CSV file to Canvas course", dest="import_canvas_rubrics", metavar="RUBRIC_FILE")
    canvas_rubric_group.add_argument('--update-canvas-rubrics', '-ur', type=str, nargs='*',
                                     help="Update rubric for one or more Canvas assignments (provide assignment IDs, or leave blank to select interactively)",
                                     dest="update_canvas_rubrics", metavar="ASSIGNMENT_IDS")
    canvas_rubric_group.add_argument('--update-canvas-rubric-id', '-uri', type=str,
                                     help="Rubric ID to associate with assignments (leave blank to select interactively)",
                                     dest="update_canvas_rubric_id", metavar="RUBRIC_ID")
    canvas_rubric_group.add_argument('--export-canvas-grading-scheme', '-egs', action='store_true',
                                     help="List and export Canvas grading schemes (grading standards) to JSON",
                                     dest="export_canvas_grading_scheme")
    canvas_rubric_group.add_argument('--add-canvas-grading-scheme', '-ags', type=str,
                                     help="Add a grading scheme to Canvas course from JSON file",
                                     dest="add_canvas_grading_scheme", metavar="GRADING_SCHEME_FILE")
    canvas_rubric_group.add_argument('--check-student-submission-similarity', '-css', nargs='?',
                                     help="Check similarities between submissions of the same student for different assignments. "
                                          "Optionally provide a Canvas student ID or a comma-separated list of IDs. "
                                          "If not provided, will prompt for selection interactively.",
                                     dest="check_student_submission_similarity")
    canvas_rubric_group.add_argument('--send-final-evaluations', '-sfe', nargs='?', const=True,
                                     help="Send final evaluation results to students via Canvas. Optionally provide directory with evaluation files (default: final_evaluations).",
                                     dest="send_final_evaluations", metavar="DIR")
    canvas_rubric_group.add_argument('--final-evals-course-id', '-fec', type=str,
                                     help="Canvas course ID to use when sending final evaluations (overrides default CANVAS_LMS_COURSE_ID).",
                                     dest="final_evals_course_id")
    canvas_rubric_group.add_argument('--final-evals-announce', '-fea', action='store_true',
                                     help="Also create a course announcement after sending final evaluations.",
                                     dest="final_evals_announce")

    canvas_admin_group = parser.add_argument_group("Canvas: Admin Tools")
    canvas_admin_group.add_argument('--no-restricted', '-nres', action='store_true',
                                    help="Disable restricted mode for grading Canvas assignments (list all assignments with submissions and all students who submitted)",
                                    dest="no_restricted")
    canvas_admin_group.add_argument('--change-canvas-deadlines', '-ccd', nargs='*',
                                    help="Change deadlines for one or more Canvas assignments (provide assignment IDs, or leave blank to select interactively)",
                                    dest="change_canvas_deadlines", metavar="ASSIGNMENT_IDS")
    canvas_admin_group.add_argument('--change-canvas-lock-dates', '-ccl', nargs='*',
                                    help="Change lock dates (lock_at) for one or more Canvas assignments (provide assignment IDs, or leave blank to select interactively)",
                                    dest="change_canvas_lock_dates", metavar="ASSIGNMENT_IDS")
    canvas_admin_group.add_argument('--new-canvas-lock-date', '-ncl', type=str,
                                    help="New lock date for Canvas assignments (format: YYYY-MM-DD HH:MM)",
                                    dest="new_canvas_lock_date", metavar="NEW_LOCK_DATE")
    canvas_admin_group.add_argument('--canvas-lock-category', '-clc', type=str,
                                    help="Assignment category (group) to filter when changing lock dates",
                                    dest="canvas_lock_category")
    canvas_admin_group.add_argument('--new-canvas-due-date', '-ncd', type=str,
                                    help="New due date for Canvas assignments (format: YYYY-MM-DD HH:MM)",
                                    dest="new_canvas_due_date", metavar="NEW_DUE_DATE")
    canvas_admin_group.add_argument('--canvas-deadline-category', '-cdc', type=str,
                                    help="Assignment category (group) to filter when changing deadlines",
                                    dest="canvas_deadline_category")
    canvas_admin_group.add_argument('--create-canvas-groups', action='store_true',
                                    help="Create groups in a Canvas course group set",
                                    dest="create_canvas_groups")
    canvas_admin_group.add_argument('--group-set-id', type=str,
                                    help="Canvas group set ID to create groups in (leave blank to select interactively)",
                                    dest="group_set_id")
    canvas_admin_group.add_argument('--num-groups', type=int, default=5,
                                    help="Number of groups to create (default: 5)",
                                    dest="num_groups")
    canvas_admin_group.add_argument('--group-name-pattern', type=str, default="Group {i}",
                                    help="Pattern for group names, e.g., 'Group {i}' (default: 'Group {i}')",
                                    dest="group_name_pattern")
    canvas_admin_group.add_argument('--delete-empty-canvas-groups', '-deg', action='store_true',
                                    help="Delete all empty groups (groups with no members) from a Canvas course group set",
                                    dest="delete_empty_canvas_groups")

    gclass_group = parser.add_argument_group("Google Classroom")
    gclass_group.add_argument('--sync-google-classroom', '-sgc', action='store_true',
                              help="Sync students in the local database with active students from Google Classroom course",
                              dest="sync_google_classroom")
    gclass_group.add_argument('--google-course-id', '-gci', type=str,
                              help="Google Classroom course ID (prompts if None)",
                              dest="google_course_id")
    gclass_group.add_argument('--google-credentials-path', '-gcp', type=str, default=None,
                              help="Path to Google Classroom credentials JSON file",
                              dest="google_credentials_path")
    gclass_group.add_argument('--google-token-path', '-gtp', type=str, default=None,
                              help="Path to Google Classroom token pickle file",
                              dest="google_token_path")

    args = parser.parse_args()

    # Persist course code early so config resolution is consistent for this run.
    if args.course_code:
        cache_course_code(args.course_code)

    if args.clear_config or args.clear_credentials:
        if args.clear_config:
            cleared = clear_config(course_code=args.course_code, verbose=args.verbose)
            if not args.verbose:
                msg = "Config file removed." if cleared else "Config file not found or could not be removed."
                print(msg)
        if args.clear_credentials:
            results = clear_credentials(course_code=args.course_code, verbose=args.verbose)
            if not args.verbose:
                any_removed = results.get("credentials") or results.get("token")
                msg = "Credentials cleared." if any_removed else "Credentials not found or could not be removed."
                print(msg)
        raise SystemExit(0)

    # Load config and set global variables for downstream modules.
    if args.config:
        config_path = args.config
        config = load_config(config_path, verbose=args.verbose)
        # Save to default location
        default_config_path = get_default_config_path(course_code=args.course_code, verbose=args.verbose)
        with open(default_config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        if args.verbose:
            print(f"[Config] Loaded config from {config_path}, saved to {default_config_path}")
    else:
        default_config_path = get_default_config_path(course_code=args.course_code, verbose=args.verbose)
        config = load_config(default_config_path, verbose=args.verbose)
        if args.verbose:
            print(f"[Config] Loaded config from default location: {default_config_path}")

    # Promote config values to module-level defaults (legacy behavior).
    if config:
        globals().update(config)

    if getattr(args, "no_restricted", False):
        DEFAULT_RESTRICTED = False
    else:
        DEFAULT_RESTRICTED = True

    # Database files are resolved from the current working directory.
    db_filename = args.db if args.db else "students.db"
    db_path = os.path.join(os.getcwd(), db_filename)

    students = []
    if os.path.exists(db_path):
        students = load_database(db_path, verbose=args.verbose)
    if args.load:
        print(f"Loaded {len(students)} students from database.")

    if args.add_file:
        new_students = read_students_from_excel_csv(args.add_file, db_path=db_path, verbose=args.verbose)
        students = load_database(db_path, verbose=args.verbose)
        print(f"Current number of students in database: {len(students)}.")

    if args.save:
        save_database(students, db_path=db_path, verbose=args.verbose)
        print("Database saved.")

    if args.export_excel:
        export_to_excel(students, args.export_excel, db_path=db_path, verbose=args.verbose)

    if args.export_emails:
        export_emails_to_txt(students, args.export_emails, db_path=db_path, verbose=args.verbose)

    if args.export_all_details:
        export_all_details_to_txt(students, args.export_all_details, db_path=db_path, verbose=args.verbose)

    if args.search:
        results = search_students(students, args.search, db_path=db_path, verbose=args.verbose)
        if results:
            print(f"Found {len(results)} student(s):")
            for idx, s in enumerate(results, 1):
                print(f"{idx}: {s.__dict__}")
        else:
            print("No student found matching your query.")

    if args.details:
        print_student_details(students, args.details, db_path=db_path, verbose=args.verbose)

    if args.all_details:
        print_all_student_details(students, db_path=db_path, verbose=args.verbose)

    if args.add_blackboard_counts:
        pdf_path = args.add_blackboard_counts
        add_blackboard_counts_from_pdf(
            pdf_path,
            db_path=db_path,
            lang=args.ocr_lang if hasattr(args, "ocr_lang") and args.ocr_lang else "auto",
            service=args.ocr_service if hasattr(args, "ocr_service") and args.ocr_service else "ocrspace",
            verbose=args.verbose
        )

    if args.extract_text:
        pdf_path = args.extract_text
        extract_text_from_scanned_pdf(
            pdf_path,
            service=args.ocr_service if hasattr(args, "ocr_service") and args.ocr_service else "ocrspace",
            lang=args.ocr_lang if hasattr(args, "ocr_lang") and args.ocr_lang else "auto",
            simple_text=args.simple_text if hasattr(args, "simple_text") else False,
            refine=args.refine if hasattr(args, "refine") else None,
            verbose=args.verbose
        )

    if args.print_blackboard_counts:
        print_all_blackboard_counts_by_date(students, db_path=db_path, verbose=args.verbose)

    # Updated export blackboard counts by date option
    if args.export_blackboard_counts is not None:
        file_path = None
        if isinstance(args.export_blackboard_counts, str) and args.export_blackboard_counts not in ("", "True"):
            file_path = args.export_blackboard_counts
        # If user did not specify a file, default to TXT
        if not file_path:
            file_path = os.path.join(os.getcwd(), "blackboard_counts_by_date.txt")
        if file_path.lower().endswith(".md"):
            export_all_blackboard_counts_by_date_to_markdown(students, file_path=file_path, db_path=db_path, verbose=args.verbose)
        else:
            export_all_blackboard_counts_by_date_to_txt(students, file_path=file_path, db_path=db_path, verbose=args.verbose)

    if args.modify:
        interactive_modify_database(students, db_path=db_path, verbose=args.verbose)

    if args.list_canvas_assignments:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        category = args.canvas_assignment_category if hasattr(args, "canvas_assignment_category") and args.canvas_assignment_category else None
        assignments_by_group = list_canvas_assignments(
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            category=category,
            verbose=args.verbose
        )
        if not assignments_by_group:
            print("No assignments found or failed to fetch assignments.")

    if args.list_canvas_members:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        people = list_canvas_people(
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            verbose=args.verbose
        )
        print_canvas_people(people, verbose=args.verbose)
        # Output is already printed in list_canvas_people

    if args.search_canvas_user:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        search_canvas_user(
            args.search_canvas_user,
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            verbose=args.verbose
        )

    if args.download_canvas_assignment:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        dest_dir = args.download_dest_dir if hasattr(args, "download_dest_dir") and args.download_dest_dir else None
        assignment_id = args.download_canvas_assignment if args.download_canvas_assignment is not True else None
        download_canvas_assignment_submissions(
            assignment_id=assignment_id,
            dest_dir=dest_dir,
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
            verbose=args.verbose
        )

    if args.comment_canvas_submission:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        add_comment_to_canvas_submission(
            assignment_id=None,
            student_canvas_id=None,
            comment_text=None,
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
            refine=args.refine if hasattr(args, "refine") else None,
            verbose=args.verbose
        )

    if args.add_canvas_announcement:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        add_canvas_announcement(
            title=None,
            message=None,
            course_id=course_id,
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            verbose=args.verbose
        )

    # Invite a single user to Canvas course by email, name, and role
    if args.invite_canvas_email:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        email = args.invite_canvas_email
        name = args.invite_canvas_name if hasattr(args, "invite_canvas_name") else None
        role = args.invite_canvas_role if hasattr(args, "invite_canvas_role") and args.invite_canvas_role else "student"
        result = invite_user_to_canvas_course(
            email=email,
            name=name,
            role=role,
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            verbose=args.verbose
        )
        print(result)

    # Invite multiple users to Canvas course from a TXT file or string of pairs/emails
    if args.invite_canvas_file:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        data_file = args.invite_canvas_file
        results = invite_users_to_canvas_course(
            data_file=data_file,
            name=None,
            role="student",
            course_id=course_id,
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            verbose=args.verbose
        )
        print(results)

    # New: Find and notify students who have not completed required peer reviews
    if args.notify_incomplete_reviews:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        assignment_id = args.review_assignment_id if hasattr(args, "review_assignment_id") and args.review_assignment_id else None
        notify_incomplete_canvas_peer_reviews(
            assignment_id=assignment_id,
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            refine=args.refine if hasattr(args, "refine") else None,
            category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
            verbose=args.verbose
        )
        
    # New: Sync Canvas students to local database
    if args.sync_canvas:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        added, updated = sync_students_with_canvas(
            students,
            db_path=db_path,
            course_id=course_id,
            api_key=CANVAS_LMS_API_KEY,
            api_url=CANVAS_LMS_API_URL,
            verbose=args.verbose
        )
        print(f"Sync completed: {added} students added, {updated} students updated.")

    # New: Grade Canvas assignment submissions interactively
    if args.grade_canvas_assignment:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        grade_canvas_assignment_submissions(
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
            verbose=args.verbose,
            restricted=DEFAULT_RESTRICTED
        )

    # New: Fetch and reply to Canvas messages
    if args.fetch_canvas_messages:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        refine = args.refine if hasattr(args, "refine") else None
        fetch_and_reply_canvas_messages(
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            only_unread=False,
            reply_text=None,
            refine=refine,
            max_messages=3,
            verbose=args.verbose
        )

    # New: Edit Canvas pages
    if getattr(args, "edit_canvas_pages", False):
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        refine = args.refine if hasattr(args, "refine") else None
        list_and_update_canvas_pages(
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            refine=refine,
            verbose=args.verbose
        )
        
    # Option: List students who submitted twice or more and the first submission is on time
    if getattr(args, "list_multiple_submissions_on_time", False):
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        assignment_id = args.list_multiple_submissions_on_time if hasattr(args, "list_multiple_submissions_on_time") and args.list_multiple_submissions_on_time is not None else None
        results = list_students_with_multiple_submissions_on_time(
            assignment_id=assignment_id,
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            verbose=args.verbose
        )
        if results:
            print(f"Students with >=2 submissions and first on time for assignment {assignment_id if assignment_id else '[selected assignments]'}:")
            for r in results:
                print(f"- {r['name']} (Canvas ID: {r['canvas_id']}): {len(r['submissions'])} submissions, first: {r['submissions'][0]}")
        else:
            print("No students found with >=2 submissions and first on time.")

    # New: List and export Canvas rubrics
    if getattr(args, "list_canvas_rubrics", False) or getattr(args, "export_canvas_rubrics", None):
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        assignment_id = args.rubric_assignment_id if hasattr(args, "rubric_assignment_id") and args.rubric_assignment_id else None
        export_path = args.export_canvas_rubrics if hasattr(args, "export_canvas_rubrics") and args.export_canvas_rubrics else None
        rubrics = list_and_export_canvas_rubrics(
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            assignment_id=assignment_id,
            export_path=export_path,
            verbose=args.verbose
        )
        if not export_path:
            print("Rubrics listed above.")
        else:
            print(f"Rubrics exported to {export_path}")

    # New: Import rubrics to Canvas course from TXT/CSV file
    if getattr(args, "import_canvas_rubrics", None):
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        rubric_file = args.import_canvas_rubrics
        results = import_canvas_rubrics(
            rubric_file=rubric_file,
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            verbose=args.verbose
        )
        print(results)
        
    # New: Update rubric for one or more Canvas assignments
    if getattr(args, "update_canvas_rubrics", None) is not None or getattr(args, "update_canvas_rubric_id", None) is not None:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        assignment_ids = args.update_canvas_rubrics if hasattr(args, "update_canvas_rubrics") and args.update_canvas_rubrics else None
        rubric_id = args.update_canvas_rubric_id if hasattr(args, "update_canvas_rubric_id") and args.update_canvas_rubric_id else None
        results = update_canvas_rubrics_for_assignments(
            assignment_ids=assignment_ids,
            rubric_id=rubric_id,
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
            verbose=args.verbose
        )
        print("Rubric update results:", results)
        
    if getattr(args, "extract_multichoice_solutions", None):
        pdf_path = args.extract_multichoice_solutions
        verbose = args.verbose if hasattr(args, "verbose") else False
        solutions = read_multichoice_exam_solutions_from_pdf(
            pdf_path,
            verbose=verbose
        )
        print("Extracted multichoice exam solutions:")
        for sheet_code, answers in solutions.items():
            print(f"Sheet code: {sheet_code}")
            for q_num in sorted(answers.keys()):
                print(f"  Question {q_num}: {answers[q_num]}")
                
    if getattr(args, "extract_multichoice_answers", None):
        pdf_path = args.extract_multichoice_answers
        ocr_service = args.ocr_service if hasattr(args, "ocr_service") and args.ocr_service else DEFAULT_OCR_METHOD
        lang = args.ocr_lang if hasattr(args, "ocr_lang") and args.ocr_lang else "auto"
        refine = args.refine if hasattr(args, "refine") else None
        verbose = args.verbose if hasattr(args, "verbose") else False
        answers = read_multichoice_answers_from_scanned_pdf(
            pdf_path,
            ocr_service=ocr_service,
            lang=lang,
            refine=refine,
            verbose=verbose
        )
        print("Extracted multichoice exam answers:")
        for entry in answers:
            print(f"Sheet code: {entry['sheet_code']}")
            print(f"Student ID: {entry['student_id']}")
            print(f"Student name: {entry['student_name']}")
            for q_num in sorted(entry['answers'].keys()):
                print(f"  Question {q_num}: {entry['answers'][q_num]}")

    # Option: Evaluate student answers for multiple-choice exam
    if getattr(args, "evaluate_multichoice_answers", None) is not None:
        exam_type = args.evaluate_multichoice_answers if args.evaluate_multichoice_answers else EXAM_TYPE
        db_path_eval = db_path if os.path.exists(db_path) else None
        verbose = args.verbose if hasattr(args, "verbose") else False
        results = evaluate_multichoice_exam_answers(
            exam_type=exam_type,
            db_path=db_path_eval,
            verbose=verbose
        )
        print(f"Evaluated {len(results)} students for exam type '{exam_type}':")
        for entry in results:
            print(f"Student ID: {entry['student_id']}, Name: {entry['student_name']}, Sheet: {entry['sheet_code']}, Mark: {entry['mark']:.2f} (Reward: {entry['reward_points']})")

    # New: Update MAT*.xlsx files with grades from database
    if getattr(args, "update_mat_excel", None):
        mat_files = args.update_mat_excel
        if not isinstance(mat_files, list):
            mat_files = [mat_files]
        for mat_file in mat_files:
            if not os.path.exists(mat_file):
                print(f"File not found: {mat_file}")
                continue
            print(f"Updating MAT Excel file: {mat_file}")
            updated_path = update_mat_excel_grades(mat_file, students, output_path=None, verbose=args.verbose)
            print(f"Updated MAT Excel file saved to: {updated_path}")

    # New: Sync multichoice exam evaluations to Canvas assignment
    if getattr(args, "sync_multichoice_evaluations", None) is not None:
        exam_type = args.sync_multichoice_evaluations if args.sync_multichoice_evaluations else EXAM_TYPE
        db_path_eval = db_path if os.path.exists(db_path) else None
        verbose = args.verbose if hasattr(args, "verbose") else False
        sync_multichoice_evaluations_to_canvas(
            exam_type=exam_type,
            db_path=db_path_eval,
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=CANVAS_LMS_COURSE_ID,
            verbose=verbose
        )
        
    # New: Export Canvas grading scheme(s) to JSON
    if getattr(args, "export_canvas_grading_scheme", False):
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        list_and_download_canvas_grading_standards(
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            verbose=args.verbose
        )

    # New: Add grading scheme to Canvas course from JSON file
    if getattr(args, "add_canvas_grading_scheme", None):
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        grading_scheme_file = args.add_canvas_grading_scheme
        add_canvas_grading_scheme(
            grading_scheme_file=grading_scheme_file,
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            verbose=args.verbose
        )
        
    # New: Check similarities between submissions of the same student for different assignments
    if hasattr(args, "check_student_submission_similarity") and args.check_student_submission_similarity is not None:
        # Accepts: no argument (interactive), a single id, or a comma-separated list of ids
        arg_val = args.check_student_submission_similarity
        if arg_val is None or arg_val is True or (isinstance(arg_val, str) and arg_val.strip() == ""):
            student_canvas_ids = None  # Interactive selection
        else:
            # Parse comma-separated list
            ids = [x.strip() for x in str(arg_val).split(",") if x.strip()]
            if len(ids) == 1:
                student_canvas_ids = ids[0]
            else:
                student_canvas_ids = ids
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        dest_dir = args.download_dest_dir if hasattr(args, "download_dest_dir") and args.download_dest_dir else os.path.join(DEFAULT_DOWNLOAD_FOLDER, "student_submissions")
        ocr_service = args.ocr_service if hasattr(args, "ocr_service") and args.ocr_service else DEFAULT_OCR_METHOD
        lang = args.ocr_lang if hasattr(args, "ocr_lang") and args.ocr_lang else "auto"
        refine = args.refine if hasattr(args, "refine") else None
        similarity_threshold = 0.85
        db_path_check = db_path if os.path.exists(db_path) else None
        download_and_check_student_submissions(
            student_canvas_id=student_canvas_ids,
            dest_dir=dest_dir,
            api_url=CANVAS_LMS_API_URL,
            api_key=CANVAS_LMS_API_KEY,
            course_id=course_id,
            ocr_service=ocr_service,
            lang=lang,
            refine=refine,
            similarity_threshold=similarity_threshold,
            db_path=db_path_check,
            verbose=args.verbose
        )
    
    if getattr(args, "send_final_evaluations", False):
        # Determine directory: if option provided without value, args.send_final_evaluations is True
        final_dir = args.send_final_evaluations if isinstance(args.send_final_evaluations, str) else "final_evaluations"
        course_id = args.final_evals_course_id if getattr(args, "final_evals_course_id", None) else CANVAS_LMS_COURSE_ID
        try:
            if args.verbose:
                print(f"[FinalEvals] Sending final evaluations from '{final_dir}' to Canvas course {course_id}...")
            else:
                print("Sending final evaluations...")
            result = send_final_evaluations_via_canvas(
                final_dir=final_dir,
                db_path=db_path,
                api_url=CANVAS_LMS_API_URL,
                api_key=CANVAS_LMS_API_KEY,
                course_id=course_id,
                verbose=args.verbose
            )
            if args.verbose:
                print(f"[FinalEvals] send_final_evaluations_via_canvas returned: {result}")
            else:
                print("Final evaluations processing complete.")

            # Optionally create a course announcement after sending evaluations
            if getattr(args, "final_evals_announce", False):
                title = "Kết quả đánh giá cuối kỳ đã được gửi"
                # Try to infer how many were sent from result
                sent_count = None
                try:
                    if isinstance(result, dict):
                        sent_count = result.get("sent") or result.get("count") or result.get("sent_count")
                    elif isinstance(result, (list, tuple, set)):
                        sent_count = len(result)
                    elif isinstance(result, int):
                        sent_count = result
                except Exception:
                    sent_count = None

                message = f"Thông báo: Kết quả đánh giá cuối kỳ đã được gửi từ thư mục '{final_dir}'."
                if sent_count is not None:
                    message += f" Số sinh viên được gửi: {sent_count}."
                message += "\n\nThông báo này được gửi tự động."

                try:
                    add_canvas_announcement(
                        title=title,
                        message=message,
                        api_url=CANVAS_LMS_API_URL,
                        api_key=CANVAS_LMS_API_KEY,
                        course_id=course_id,
                        verbose=args.verbose
                    )
                    if args.verbose:
                        print("[FinalEvals] Course announcement created.")
                except Exception as e:
                    if args.verbose:
                        print(f"[FinalEvals] Failed to create announcement: {e}")
                    else:
                        print("Failed to create announcement.")
        except Exception as e:
            if args.verbose:
                print(f"[FinalEvals] Error sending final evaluations: {e}")
                traceback.print_exc()
            else:
                print(f"Error sending final evaluations: {e}")

    # Handle export final grade distribution option
    if getattr(args, "export_final_grade_distribution", False):
        try:
            efgd_arg = args.export_final_grade_distribution
            # Determine requested output path (None -> use default inside function)
            out_path = None
            if isinstance(efgd_arg, str) and efgd_arg not in ("", "True"):
                out_path = efgd_arg

            # Ensure students loaded (try DB if needed)
            if not students and os.path.exists(db_path):
                students = load_database(db_path, verbose=args.verbose)

            # Call calculation (this writes default file ./final_grade_distribution.txt)
            result = calculate_and_print_final_grade_distribution(
                students,
                db_path=db_path,
                grade_field=None,
                verbose=args.verbose
            )

            # If user specified an explicit path, copy the generated default file to that path
            default_report = os.path.join(os.getcwd(), "final_grade_distribution.txt")
            if out_path:
                try:
                    if os.path.exists(default_report):
                        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
                        shutil.copyfile(default_report, out_path)
                        if args.verbose:
                            print(f"[GradeDist] Copied report to: {out_path}")
                        else:
                            print(f"Report exported to: {out_path}")
                    else:
                        # If default report not found but result exists, try to create simple report file at out_path
                        with open(out_path, "w", encoding="utf-8") as f:
                            f.write(f"Final grade distribution generated. Result summary: {json.dumps(result, ensure_ascii=False)}\n")
                        if args.verbose:
                            print(f"[GradeDist] Wrote fallback report to: {out_path}")
                        else:
                            print(f"Report exported to: {out_path}")
                except Exception as e:
                    if args.verbose:
                        print(f"[GradeDist] Failed to copy/write report to {out_path}: {e}")
                    else:
                        print(f"Failed to export report to: {out_path}")
            else:
                # No explicit path: notify user where default report was written
                default_report_path = os.path.join(os.getcwd(), "final_grade_distribution.txt")
                if os.path.exists(default_report_path):
                    if args.verbose:
                        print(f"[GradeDist] Report written to: {default_report_path}")
                    else:
                        print(f"Report written to: {default_report_path}")
                else:
                    if args.verbose:
                        print("[GradeDist] Report generation completed (no file found).")
                    else:
                        print("Report generation completed.")
        except Exception as e:
            if args.verbose:
                print(f"[GradeDist] Error exporting final grade distribution: {e}")
            else:
                print(f"Error exporting final grade distribution: {e}")
    
    # Handle change Canvas deadlines option (CLI)
    if getattr(args, "change_canvas_deadlines", None) is not None:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        raw_ids = args.change_canvas_deadlines  # nargs='*' -> list or []
        # If user passed -cd without values -> interactive selection inside function
        if raw_ids == []:
            assignment_ids = None
        else:
            # Normalize list: split any comma-separated items and strip
            ids = []
            for item in raw_ids:
                for part in str(item).split(","):
                    part = part.strip()
                    if part:
                        ids.append(part)
            assignment_ids = ids if ids else None

        new_due = args.new_canvas_due_date if hasattr(args, "new_canvas_due_date") and args.new_canvas_due_date else None
        category = args.canvas_deadline_category if hasattr(args, "canvas_deadline_category") and args.canvas_deadline_category else None

        try:
            if args.verbose:
                print(f"[ChangeDeadlines] Invoking change_canvas_deadlines(course_id={course_id}, assignment_ids={assignment_ids}, new_due_date={new_due}, category={category})")
            results = change_canvas_deadlines(
                assignment_ids=assignment_ids,
                new_due_date=new_due,
                api_url=CANVAS_LMS_API_URL,
                api_key=CANVAS_LMS_API_KEY,
                course_id=course_id,
                category=category,
                verbose=args.verbose
            )
            if args.verbose:
                print(f"[ChangeDeadlines] Results: {results}")
            else:
                print("Change deadlines operation completed.")
                if isinstance(results, dict) and results:
                    for aid, status in results.items():
                        print(f"Assignment {aid}: {status}")
        except Exception as e:
            if args.verbose:
                print(f"[ChangeDeadlines] Error: {e}")
            else:
                print(f"Error changing deadlines: {e}")
    
    if getattr(args, "create_canvas_groups", False):
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        group_set_id = args.group_set_id if hasattr(args, "group_set_id") and args.group_set_id else None
        num_groups = args.num_groups if hasattr(args, "num_groups") else 5
        group_name_pattern = args.group_name_pattern if hasattr(args, "group_name_pattern") else None
        try:
            created_groups = create_canvas_groups(
                api_url=CANVAS_LMS_API_URL,
                api_key=CANVAS_LMS_API_KEY,
                course_id=course_id,
                group_set_id=group_set_id,
                num_groups=num_groups,
                group_name_pattern=group_name_pattern,
                verbose=args.verbose
            )
            if created_groups:
                print(f"Successfully created {len(created_groups)} groups.")
            else:
                print("Failed to create groups.")
        except Exception as e:
            if args.verbose:
                print(f"[CreateGroups] Error: {e}")
            else:
                print(f"Error creating groups: {e}")
    
    # Handle export_emails_and_names
    if args.export_emails_and_names:
        export_emails_and_names_to_txt(students, args.export_emails_and_names, db_path=db_path, verbose=args.verbose)
        
    if getattr(args, "change_canvas_lock_dates", None) is not None:
        # Handle change canvas lock dates option (CLI)
        try:
            course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
            raw_ids = args.change_canvas_lock_dates  # nargs='*' -> list or []
            # If user passed -ccl without values -> interactive selection inside function
            if raw_ids == []:
                assignment_ids = None
            else:
                # Normalize list: split any comma-separated items and strip
                ids = []
                for item in raw_ids:
                    for part in str(item).split(","):
                        part = part.strip()
                        if part:
                            ids.append(part)
                assignment_ids = ids if ids else None

            new_lock = args.new_canvas_lock_date if hasattr(args, "new_canvas_lock_date") and args.new_canvas_lock_date else None
            category = args.canvas_lock_category if hasattr(args, "canvas_lock_category") and args.canvas_lock_category else None

            try:
                if args.verbose:
                    print(f"[ChangeLock] Invoking change_canvas_lock_dates(course_id={course_id}, assignment_ids={assignment_ids}, new_lock_date={new_lock}, category={category})")
                else:
                    print(f"Invoking change_canvas_lock_dates(course_id={course_id}, assignment_ids={assignment_ids}, new_lock_date={new_lock}, category={category})")
                results = change_canvas_lock_dates(
                    assignment_ids=assignment_ids,
                    new_lock_date=new_lock,
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    category=category,
                    verbose=args.verbose
                )
                if args.verbose:
                    print(f"[ChangeLock] Results: {results}")
                else:
                    print("Change lock dates operation completed.")
                    if isinstance(results, dict) and results:
                        for aid, status in results.items():
                            print(f"Assignment {aid}: {status}")
            except Exception as e:
                if args.verbose:
                    print(f"[ChangeLock] Error: {e}")
                else:
                    print(f"Error changing lock dates: {e}")
        except Exception as e:
            if args.verbose:
                print(f"[ChangeLock] Error: {e}")
            else:
                print(f"Error changing lock dates: {e}")
                
    if args.export_roster:
        export_roster_to_csv(students, file_path=args.export_roster, verbose=args.verbose)
        
    if args.sync_google_classroom:
        course_id = args.google_course_id if hasattr(args, "google_course_id") and args.google_course_id else None
        credentials_path = args.google_credentials_path if hasattr(args, "google_credentials_path") and args.google_credentials_path else None
        if not credentials_path:
            credentials_path = config.get("CREDENTIALS_PATH") if isinstance(config, dict) else None
        if not credentials_path:
            credentials_path = get_default_credentials_path(course_code=args.course_code)
        token_path = args.google_token_path if hasattr(args, "google_token_path") and args.google_token_path else None
        if not token_path:
            token_path = config.get("TOKEN_PATH") if isinstance(config, dict) else None
        if not token_path:
            token_path = get_default_token_path(course_code=args.course_code)
        added, updated = sync_students_with_google_classroom(
            students,
            db_path=db_path,
            course_id=course_id,
            credentials_path=credentials_path,
            token_path=token_path,
            verbose=args.verbose
        )
        print(f"Sync with Google Classroom completed: {added} students added, {updated} students updated.")
    
    if args.delete_empty_canvas_groups:
        course_id = args.canvas_course_id if hasattr(args, "canvas_course_id") and args.canvas_course_id else CANVAS_LMS_COURSE_ID
        group_set_id = args.group_set_id if hasattr(args, "group_set_id") and args.group_set_id else None
        try:
            deleted_count = delete_empty_canvas_groups(
                api_url=CANVAS_LMS_API_URL,
                api_key=CANVAS_LMS_API_KEY,
                course_id=course_id,
                group_set_id=group_set_id,
                verbose=args.verbose
            )
            if deleted_count > 0:
                print(f"Successfully deleted {deleted_count} empty groups.")
            else:
                print("No empty groups deleted.")
        except Exception as e:
            if args.verbose:
                print(f"[DeleteGroups] Error: {e}")
            else:
                print(f"Error deleting empty groups: {e}")

    # If no arguments are provided, show interactive menu
    if len(sys.argv) == 1:
        db_path = get_default_db_path()
        if os.path.exists(db_path):
            students = load_database(db_path, verbose=args.verbose)
        else:
            students = []
        menu_sections = _build_menu_sections()
        use_arrow_menu = True
        try:
            import msvcrt  # noqa: F401
        except Exception:
            use_arrow_menu = False

        while True:
            if use_arrow_menu:
                choice = _select_menu_option(menu_sections)
                if choice is None:
                    break
            else:
                _print_menu_fallback(menu_sections)
                choice = input("Choose an option (or 'q' to quit): ").strip()
                if choice.lower() in ('q', 'quit', '0'):
                    break
                mapped_choice = _menu_choice_to_action(choice, menu_sections)
                if not mapped_choice:
                    print("Invalid option.")
                    continue
                choice = mapped_choice

            if choice == '1':
                file_path = input_with_completion("Enter Excel/CSV file path (or 'q' to quit): ").strip()
                if file_path.lower() in ('q', 'quit', ''):
                    continue
                read_students_from_excel_csv(file_path, db_path=db_path, verbose=args.verbose)
                students = load_database(db_path, verbose=args.verbose)
                print(f"Current number of students in database: {len(students)}.")
            elif choice == '2':
                save_database(students, db_path=db_path, verbose=args.verbose)
                print("Database saved.")
            elif choice == '3':
                students = load_database(db_path, verbose=args.verbose)
                print(f"Loaded {len(students)} students from database.")
            elif choice == '4':
                export_path = input_with_completion("Enter export Excel file path (or 'q' to quit): ").strip()
                if export_path.lower() in ('q', 'quit', ''):
                    continue
                export_to_excel(students, export_path, verbose=args.verbose)
            elif choice == '5':
                export_path = input_with_completion("Enter export TXT file path (or 'q' to quit): ").strip()
                if export_path.lower() in ('q', 'quit'):
                    continue
                export_emails_to_txt(students, export_path, verbose=args.verbose)
            elif choice == '6':
                query = input("Enter name or student id to search (or 'q' to quit): ").strip()
                if query.lower() in ('q', 'quit', ''):
                    continue
                results = search_students(students, query, verbose=args.verbose)
                if results:
                    print(f"Found {len(results)} student(s):")
                    for idx, s in enumerate(results, 1):
                        print(f"{idx}: {s.__dict__}")
                else:
                    print("No student found matching your query.")
            elif choice == '7':
                identifier = input("Enter name, student id, or email (or 'q' to quit): ").strip()
                if identifier.lower() in ('q', 'quit', ''):
                    continue
                print_student_details(students, identifier, verbose=args.verbose)
            elif choice == '8':
                print_all_student_details(students, verbose=args.verbose)
            elif choice == '9':
                export_path = input_with_completion("Enter export TXT file path (or 'q' to quit): ").strip()
                if export_path.lower() in ('q', 'quit'):
                    continue
                export_all_details_to_txt(students, export_path, verbose=args.verbose)
            elif choice == '10':
                pdf_path = input_with_completion("Enter PDF file path (or 'q' to quit): ").strip()
                if pdf_path.lower() in ('q', 'quit'):
                    continue
                ocr_service = input("OCR service (ocrspace/tesseract/paddleocr) [ocrspace] (or 'q' to quit): ").strip().lower()
                if ocr_service in ('q', 'quit'):
                    continue
                ocr_service = ocr_service or "ocrspace"
                ocr_lang = input("OCR language [auto] (or 'q' to quit): ").strip()
                if ocr_lang in ('q', 'quit'):
                    continue
                ocr_lang = ocr_lang or "auto"
                add_blackboard_counts_from_pdf(
                    pdf_path,
                    db_path=db_path,
                    lang=ocr_lang,
                    service=ocr_service,
                    verbose=args.verbose
                )
            elif choice == '11':
                pdf_path = input_with_completion("Enter PDF file path (or 'q' to quit): ").strip()
                if pdf_path.lower() in ('q', 'quit', ''):
                    continue
                ocr_service = input("OCR service (ocrspace/tesseract/paddleocr) [ocrspace] (or 'q' to quit): ").strip().lower()
                if ocr_service in ('q', 'quit'):
                    continue
                ocr_service = ocr_service or "ocrspace"
                ocr_lang = input("OCR language [auto] (or 'q' to quit): ").strip()
                if ocr_lang in ('q', 'quit'):
                    continue
                ocr_lang = ocr_lang or "auto"
                simple_text = input("Simple text output? (y/n) [n] (or 'q' to quit): ").strip().lower()
                if simple_text in ('q', 'quit'):
                    continue
                simple_text = simple_text == "y"
                refine = input("Refine extracted text with AI? (none/gemini/huggingface) [none] (or 'q' to quit): ").strip().lower()
                if refine in ('q', 'quit'):
                    continue
                refine = refine if refine in ALL_AI_METHODS else None
                extract_text_from_scanned_pdf(
                    pdf_path,
                    service=ocr_service,
                    lang=ocr_lang,
                    simple_text=simple_text,
                    refine=refine,
                    verbose=args.verbose
                )
            elif choice == '12':
                print_all_blackboard_counts_by_date(students, db_path=db_path, verbose=args.verbose)
            elif choice == '13':
                export_path = input_with_completion("Enter export TXT/Markdown file path (or leave blank for TXT, 'q' to quit): ").strip()
                if export_path.lower() in ('q', 'quit'):
                    export_path = None
                # Default to TXT if not specified
                if not export_path:
                    export_path = os.path.join(os.getcwd(), "blackboard_counts_by_date.txt")
                if export_path.lower().endswith(".md"):
                    export_all_blackboard_counts_by_date_to_markdown(students, file_path=export_path, db_path=db_path, verbose=args.verbose)
                else:
                    export_all_blackboard_counts_by_date_to_txt(students, file_path=export_path, db_path=db_path, verbose=args.verbose)
            elif choice == '14':
                interactive_modify_database(students, db_path=db_path, verbose=args.verbose)
            elif choice == '15':
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                category = input("Enter assignment category (leave blank for all, or 'q' to quit): ").strip()
                if category.lower() in ('q', 'quit'):
                    continue
                category = category if category else None
                assignments_by_group = list_canvas_assignments(
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    category=category,
                    verbose=args.verbose
                )
                if not assignments_by_group:
                    print("No assignments found or failed to fetch assignments.")
            elif choice == '16':
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                people = list_canvas_people(
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    verbose=args.verbose
                )
                print_canvas_people(people, verbose=args.verbose)
            elif choice == '17':
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                query = input("Enter name or email to search in Canvas (or 'q' to quit): ").strip()
                if query.lower() in ('q', 'quit', ''):
                    continue
                search_canvas_user(
                    query,
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    verbose=args.verbose
                )
            elif choice == '18':
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                assignment_id = input("Enter Canvas assignment ID (leave blank to select, or 'q' to quit): ").strip()
                if assignment_id.lower() in ('q', 'quit'):
                    continue
                if not assignment_id:
                    assignment_id = None
                dest_dir = input_with_completion(
                    f"Enter destination directory for downloads (leave blank for default directory {DEFAULT_DOWNLOAD_FOLDER}, or 'q' to quit): ",
                    select_file=False
                ).strip()
                if dest_dir.lower() in ('q', 'quit'):
                    continue
                if not dest_dir:
                    dest_dir = DEFAULT_DOWNLOAD_FOLDER
                download_canvas_assignment_submissions(
                    assignment_id=assignment_id,
                    dest_dir=dest_dir,
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
                    verbose=args.verbose
                )
            elif choice == '19':
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                add_comment_to_canvas_submission(
                    assignment_id=None,
                    student_canvas_id=None,
                    comment_text=None,
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
                    refine=None,
                    verbose=args.verbose
                )
            elif choice == '20':
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                add_canvas_announcement(
                    title=None,
                    message=None,
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    verbose=args.verbose
                )
            elif choice == '21':
                # Invite a single user
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                email = input("Enter user email to invite (or 'q' to quit): ").strip()
                if email.lower() in ('q', 'quit', ''):
                    continue
                name = input("Enter user name (or 'q' to quit): ").strip()
                if name.lower() in ('q', 'quit', ''):
                    continue
                role = input("Enter role (student/teacher/ta, default: student): ").strip().lower()
                role = role if role in {"student", "teacher", "ta"} else "student"
                section = input("Enter section name to enroll in (leave blank for default section, or 'q' to quit): ").strip()
                if section.lower() in ('q', 'quit'):
                    continue
                section = section if section else None
                result = invite_user_to_canvas_course(
                    email=email,
                    name=name,
                    role=role,
                    section=section,
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    verbose=args.verbose
                )
                print(result)
            elif choice == '22':
                # Invite multiple users from file
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                data_file = input_with_completion("Enter path to TXT file or string of pairs/emails (or 'q' to quit): ").strip()
                if data_file.lower() in ('q', 'quit', ''):
                    continue
                results = invite_users_to_canvas_course(
                    data_file=data_file,
                    name=None,
                    role="student",
                    course_id=course_id,
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    verbose=args.verbose
                )
                print(results)
            elif choice == '23':
                # Find and notify students who have not completed required peer reviews
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                assignment_id = input("Enter Canvas assignment ID for peer review notification (or leave blank to select): ").strip()
                assignment_id = assignment_id if assignment_id else None
                refine = input("Refine reminder message with AI? (none/gemini/huggingface) [none] (or 'q' to quit): ").strip().lower()
                if refine in ('q', 'quit'):
                    continue
                refine = refine if refine in ALL_AI_METHODS else None
                notify_incomplete_canvas_peer_reviews(
                    assignment_id=assignment_id,
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    refine=refine,
                    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
                    verbose=args.verbose
                )
            elif choice == '24':
                # New option: Sync Canvas students to local database
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                added, updated = sync_students_with_canvas(
                    students,
                    db_path=db_path,
                    course_id=course_id,
                    api_key=CANVAS_LMS_API_KEY,
                    api_url=CANVAS_LMS_API_URL,
                    verbose=args.verbose
                )
                print(f"Sync completed: {added} students added, {updated} students updated.")
                # Refresh students list after syncing
                students = load_database(db_path, verbose=args.verbose)
            elif choice == '25':
                # Grade Canvas assignment submissions interactively
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                grade_canvas_assignment_submissions(
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
                    verbose=args.verbose,
                    restricted=DEFAULT_RESTRICTED
                )
            elif choice == '26':
                # Fetch and reply to Canvas inbox messages
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                refine = input("Refine reply with AI? (none/gemini/huggingface) [none] (or 'q' to quit): ").strip().lower()
                if refine in ('q', 'quit'):
                    continue
                refine = refine if refine in ALL_AI_METHODS else None
                fetch_and_reply_canvas_messages(
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    only_unread=False,
                    reply_text=None,
                    refine=refine,
                    max_messages=3,
                    verbose=args.verbose
                )
            elif choice == '27':
                # List and edit Canvas course pages
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                refine = input("Refine page content with AI? (none/gemini/huggingface) [none] (or 'q' to quit): ").strip().lower()
                if refine in ('q', 'quit'):
                    continue
                refine = refine if refine in ALL_AI_METHODS else None
                list_and_update_canvas_pages(
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    refine=refine,
                    verbose=args.verbose
                )
            elif choice == '28':
                # List students with multiple submissions and only the first submission on time
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                assignment_id = input("Enter Canvas assignment ID (leave blank to select, or 'q' to quit): ").strip()
                if assignment_id.lower() in ('q', 'quit'):
                    continue
                assignment_id = assignment_id if assignment_id else None
                results = list_students_with_multiple_submissions_on_time(
                    assignment_id=assignment_id,
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    verbose=args.verbose
                )
                if results:
                    print(f"Students with >=2 submissions and first on time for assignment {assignment_id if assignment_id else '[selected assignments]'}:")
                    for r in results:
                        print(f"- {r['name']} (Canvas ID: {r['canvas_id']}): {len(r['submissions'])} submissions, first: {r['submissions'][0]}")
                else:
                    print("No students found with >=2 submissions and first on time.")
            elif choice == '29':
                # List all unique rubrics used in Canvas course
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                assignment_id = input("Enter assignment ID to filter rubrics (leave blank for all, or 'q' to quit): ").strip()
                assignment_id = assignment_id if assignment_id else None
                rubrics = list_and_export_canvas_rubrics(
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    assignment_id=assignment_id,
                    export_path=None,
                    verbose=args.verbose
                )
                print("Rubrics listed above.")
            elif choice == '30':
                # Export Canvas rubrics to TXT/CSV file
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                assignment_id = input("Enter assignment ID to filter rubrics (leave blank for all, or 'q' to quit): ").strip()
                assignment_id = assignment_id if assignment_id else None
                export_path = input_with_completion("Enter export file path (TXT/CSV, leave blank for default, or 'q' to quit): ").strip()
                if export_path.lower() in ('q', 'quit'):
                    continue
                if not export_path:
                    export_path = os.path.join(os.getcwd(), "canvas_rubrics_export.csv")
                rubrics = list_and_export_canvas_rubrics(
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    assignment_id=assignment_id,
                    export_path=export_path,
                    verbose=args.verbose
                )
                print(f"Rubrics exported to {export_path}")
            elif choice == '31':
                # Import rubrics to Canvas course from TXT/CSV file
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                rubric_file = input_with_completion("Enter path to TXT/CSV rubrics file (or 'q' to quit): ").strip()
                if rubric_file.lower() in ('q', 'quit', ''):
                    continue
                results = import_canvas_rubrics(
                    rubric_file=rubric_file,
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    verbose=args.verbose
                )
                print(results)
            elif choice == '32':
                # Update rubricsfor Canvas assignments
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                assignment_ids_input = input("Enter assignment IDs separated by commas (leave blank to select interactively, or 'q' to quit): ").strip()
                if assignment_ids_input.lower() in ('q', 'quit'):
                    continue
                assignment_ids = [aid.strip() for aid in assignment_ids_input.split(",") if aid.strip()] if assignment_ids_input else None
                rubric_id_input = input("Enter rubric ID to associate (leave blank to select interactively, or 'q' to quit): ").strip()
                if rubric_id_input.lower() in ('q', 'quit'):
                    continue
                rubric_id = rubric_id_input if rubric_id_input else None
                results = update_canvas_rubrics_for_assignments(
                    assignment_ids=assignment_ids,
                    rubric_id=rubric_id,
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
                    verbose=args.verbose
                )
                print("Rubric update results:", results)
            elif choice == '33':
                # Extract multiple-choice exam solutions from PDF
                pdf_path = input_with_completion("Enter PDF file path (or 'q' to quit): ").strip()
                if pdf_path.lower() in ('q', 'quit', ''):
                    continue
                # Remove surrounding quotes if present
                if (pdf_path.startswith('"') and pdf_path.endswith('"')) or (pdf_path.startswith("'") and pdf_path.endswith("'")):
                    pdf_path = pdf_path[1:-1]
                verbose_opt = input("Enable verbose output? (y/n) [n]: ").strip().lower()
                verbose_flag = verbose_opt == "y"
                solutions = read_multichoice_exam_solutions_from_pdf(
                    pdf_path,
                    verbose=verbose_flag
                )
                print("Extracted multichoice exam solutions:")
                for sheet_code, answers in solutions.items():
                    print(f"Sheet code: {sheet_code}")
                    for q_num in sorted(answers.keys()):
                        print(f"  Question {q_num}: {answers[q_num]}")
            elif choice == '34':
                pdf_path = input_with_completion("Enter PDF file path (or 'q' to quit): ").strip()
                if pdf_path.lower() in ('q', 'quit', ''):
                    continue
                # Remove surrounding quotes if present
                if (pdf_path.startswith('"') and pdf_path.endswith('"')) or (pdf_path.startswith("'") and pdf_path.endswith("'")):
                    pdf_path = pdf_path[1:-1]
                ocr_service = input("OCR service (ocrspace/tesseract/paddleocr) [ocrspace] (or 'q' to quit): ").strip().lower()
                if ocr_service in ('q', 'quit'):
                    continue
                ocr_service = ocr_service or "ocrspace"
                ocr_lang = input("OCR language [auto] (or 'q' to quit): ").strip()
                if ocr_lang in ('q', 'quit'):
                    continue
                ocr_lang = ocr_lang or "auto"
                refine = input("Refine extracted text with AI? (none/gemini/huggingface) [none] (or 'q' to quit): ").strip().lower()
                if refine in ('q', 'quit'):
                    continue
                refine = refine if refine in ALL_AI_METHODS else None
                verbose_opt = input("Enable verbose output? (y/n) [n]: ").strip().lower()
                verbose_flag = verbose_opt == "y"
                answers = read_multichoice_answers_from_scanned_pdf(
                    pdf_path,
                    ocr_service=ocr_service,
                    lang=ocr_lang,
                    refine=refine,
                    verbose=verbose_flag
                )
                print("Extracted multichoice exam answers:")
                for entry in answers:
                    print(f"Sheet code: {entry['sheet_code']}")
                    print(f"Student ID: {entry['student_id']}")
                    print(f"Student name: {entry['student_name']}")
                    for q_num in sorted(entry['answers'].keys()):
                        print(f"  Question {q_num}: {entry['answers'][q_num]}")
            elif choice == '35':
                exam_type = input("Enter exam type (midterm/final, default: midterm): ").strip().lower()
                if not exam_type:
                    exam_type = "midterm"
                db_path_eval = db_path if os.path.exists(db_path) else None
                verbose_flag = input("Enable verbose output? (y/n) [n]: ").strip().lower() == "y"
                results = evaluate_multichoice_exam_answers(
                    exam_type=exam_type,
                    db_path=db_path_eval,
                    verbose=verbose_flag
                )
                print(f"Evaluated {len(results)} students for exam type '{exam_type}':")
                for entry in results:
                    print(f"Student ID: {entry['student_id']}, Name: {entry['student_name']}, Sheet: {entry['sheet_code']}, Mark: {entry['mark']:.2f} (Reward: {entry['reward_points']})")
            elif choice == '36':
                config_path = input_with_completion("Enter config JSON file path (or 'q' to quit): ").strip()
                if config_path.lower() in ('q', 'quit', ''):
                    continue
                config = load_config(config_path, verbose=args.verbose)
                default_config_path = get_default_config_path(verbose=args.verbose)
                with open(default_config_path, "w", encoding="utf-8") as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                print(f"Loaded config from {config_path}, saved to {default_config_path}")
            elif choice == '37':
                mat_files = input_with_completion("Enter MAT*.xlsx file path(s), separated by commas (or 'q' to quit): ").strip()
                if mat_files.lower() in ('q', 'quit', ''):
                    continue
                mat_file_list = [f.strip() for f in mat_files.split(",") if f.strip()]
                for mat_file in mat_file_list:
                    if not os.path.exists(mat_file):
                        print(f"File not found: {mat_file}")
                        continue
                    print(f"Updating MAT Excel file: {mat_file}")
                    updated_path = update_mat_excel_grades(mat_file, students, output_path=None, verbose=args.verbose)
                    print(f"Updated MAT Excel file saved to: {updated_path}")
            elif choice == '38':
                exam_type = input("Enter exam type (midterm/final, default: midterm): ").strip().lower()
                if not exam_type:
                    exam_type = EXAM_TYPE
                db_path_eval = db_path if os.path.exists(db_path) else None
                verbose_flag = input("Enable verbose output? (y/n) [n]: ").strip().lower() == "y"
                sync_multichoice_evaluations_to_canvas(
                    exam_type=exam_type,
                    db_path=db_path_eval,
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=CANVAS_LMS_COURSE_ID,
                    verbose=verbose_flag
                )
            elif choice == '39':
                # Export Canvas grading scheme(s) to JSON
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                list_and_download_canvas_grading_standards(
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    verbose=args.verbose
                )
            elif choice == '40':
                # Add grading scheme to Canvas course from JSON file
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                grading_scheme_file = input_with_completion("Enter grading scheme JSON file path (or 'q' to quit): ").strip()
                if grading_scheme_file.lower() in ('q', 'quit', ''):
                    continue
                add_canvas_grading_scheme(
                    grading_scheme_file=grading_scheme_file,
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    verbose=args.verbose
                )
            elif choice == '41':
                # Check similarities between submissions of the same student for different assignments
                # Accepts: no argument (interactive), a single id, or a comma-separated list of ids
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                arg_val = input("Enter Canvas student ID(s) (leave blank for interactive selection, or comma-separated list, or 'q' to quit): ").strip()
                if arg_val.lower() in ('q', 'quit'):
                    continue
                if not arg_val:
                    student_canvas_ids = None  # Interactive selection
                else:
                    ids = [x.strip() for x in arg_val.split(",") if x.strip()]
                    if len(ids) == 1:
                        student_canvas_ids = ids[0]
                    else:
                        student_canvas_ids = ids
                dest_dir = input_with_completion(
                    f"Enter destination directory for downloads (leave blank for default directory {DEFAULT_DOWNLOAD_FOLDER}, or 'q' to quit): ",
                    select_file=False
                ).strip()
                if dest_dir.lower() in ('q', 'quit'):
                    continue
                if not dest_dir:
                    dest_dir = DEFAULT_DOWNLOAD_FOLDER
                ocr_service = input("OCR service (ocrspace/tesseract/paddleocr) [ocrspace] (or 'q' to quit): ").strip().lower()
                if ocr_service in ('q', 'quit'):
                    continue
                ocr_service = ocr_service or "ocrspace"
                ocr_lang = input("OCR language [auto] (or 'q' to quit): ").strip()
                if ocr_lang in ('q', 'quit'):
                    continue
                ocr_lang = ocr_lang or "auto"
                refine = input("Refine extracted text with AI? (none/gemini/huggingface) [none] (or 'q' to quit): ").strip().lower()
                if refine in ('q', 'quit'):
                    continue
                refine = refine if refine in ALL_AI_METHODS else None
                similarity_threshold = 0.85
                db_path_check = db_path if os.path.exists(db_path) else None
                download_and_check_student_submissions(
                    student_canvas_id=student_canvas_ids,
                    dest_dir=dest_dir,
                    api_url=CANVAS_LMS_API_URL,
                    api_key=CANVAS_LMS_API_KEY,
                    course_id=course_id,
                    ocr_service=ocr_service,
                    lang=ocr_lang,
                    refine=refine,
                    similarity_threshold=similarity_threshold,
                    db_path=db_path_check,
                    verbose=args.verbose
                )
            elif choice == '42':
                final_dir = input_with_completion("Enter final evaluations directory (default: final_evaluations): ").strip()
                if not final_dir:
                    final_dir = "final_evaluations"
                course_id = input("Enter Canvas course ID (leave blank for default): ").strip() or CANVAS_LMS_COURSE_ID
                announce = input("Create a course announcement after sending? (y/N): ").strip().lower() == "y"
                try:
                    if args.verbose:
                        print(f"[FinalEvals] Sending final evaluations from '{final_dir}' to Canvas course {course_id}...")
                    else:
                        print("Sending final evaluations...")
                    result = send_final_evaluations_via_canvas(
                        final_dir=final_dir,
                        db_path=db_path,
                        api_url=CANVAS_LMS_API_URL,
                        api_key=CANVAS_LMS_API_KEY,
                        course_id=course_id,
                        verbose=args.verbose
                    )
                    print("Final evaluations processing complete.")
                    if announce:
                        # Try to infer sent count
                        sent_count = None
                        try:
                            if isinstance(result, dict):
                                sent_count = result.get("sent") or result.get("count") or result.get("sent_count")
                            elif isinstance(result, (list, tuple, set)):
                                sent_count = len(result)
                            elif isinstance(result, int):
                                sent_count = result
                        except Exception:
                            sent_count = None
                        title = "Kết quả đánh giá cuối kỳ đã được gửi"
                        message = f"Thông báo: Kết quả đánh giá cuối kỳ đã được gửi từ thư mục '{final_dir}'."
                        if sent_count is not None:
                            message += f" Số sinh viên được gửi: {sent_count}."
                        message += "\n\nThông báo này được gửi tự động."
                        try:
                            add_canvas_announcement(
                                title=title,
                                message=message,
                                api_url=CANVAS_LMS_API_URL,
                                api_key=CANVAS_LMS_API_KEY,
                                course_id=course_id,
                                verbose=args.verbose
                            )
                            print("Course announcement created.")
                        except Exception as e:
                            print(f"Failed to create announcement: {e}")
                except Exception as e:
                    print(f"Error sending final evaluations: {e}")
            elif choice == '43':
                try:
                    out_path = input_with_completion("Enter output path for final grade distribution (leave blank for default ./final_grade_distribution.txt, or 'q' to cancel): ").strip()
                    if out_path.lower() in ('q', 'quit'):
                        continue
                    if not out_path:
                        out_path = None

                    # Ensure students loaded
                    if not students and os.path.exists(db_path):
                        students = load_database(db_path, verbose=args.verbose)

                    result = calculate_and_print_final_grade_distribution(
                        students,
                        db_path=db_path,
                        grade_field=None,
                        verbose=args.verbose
                    )

                    default_report = os.path.join(os.getcwd(), "final_grade_distribution.txt")
                    if out_path:
                        try:
                            os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
                            if os.path.exists(default_report):
                                shutil.copyfile(default_report, out_path)
                                print(f"Report exported to: {out_path}")
                            else:
                                # Fallback: write a simple summary
                                with open(out_path, "w", encoding="utf-8") as f:
                                    f.write(f"Final grade distribution summary:\n{json.dumps(result, ensure_ascii=False, indent=2)}\n")
                                print(f"Report written to: {out_path} (fallback summary)")
                        except Exception as e:
                            print(f"Failed to export report to {out_path}: {e}")
                    else:
                        if os.path.exists(default_report):
                            print(f"Report written to: {default_report}")
                        else:
                            print("Report generation completed.")
                except Exception as e:
                    print(f"Error exporting final grade distribution: {e}")
            elif choice == '44':
                try:
                    course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                    if course_id.lower() in ('q', 'quit'):
                        continue
                    if not course_id:
                        course_id = CANVAS_LMS_COURSE_ID

                    raw_ids = input("Enter assignment IDs separated by commas (leave blank to select interactively, or 'q' to quit): ").strip()
                    if raw_ids.lower() in ('q', 'quit'):
                        continue
                    assignment_ids = None
                    if raw_ids:
                        assignment_ids = [part.strip() for part in raw_ids.split(",") if part.strip()]

                    new_due = input("Enter new due date to apply to all selected assignments (format: YYYY-MM-DD HH:MM) or leave blank to specify per-assignment (or 'q' to quit): ").strip()
                    if new_due.lower() in ('q', 'quit'):
                        continue
                    new_due_date = new_due if new_due else None

                    category = input("Enter assignment category (group) to filter when listing assignments (leave blank for all): ").strip()
                    if category.lower() in ('q', 'quit'):
                        continue
                    category = category if category else None

                    if args.verbose:
                        print(f"[ChangeDeadlines] Calling change_canvas_deadlines(course_id={course_id}, assignment_ids={assignment_ids}, new_due_date={new_due_date}, category={category})")
                    results = change_canvas_deadlines(
                        assignment_ids=assignment_ids,
                        new_due_date=new_due_date,
                        api_url=CANVAS_LMS_API_URL,
                        api_key=CANVAS_LMS_API_KEY,
                        course_id=course_id,
                        category=category,
                        verbose=args.verbose
                    )
                    if args.verbose:
                        print(f"[ChangeDeadlines] Results: {results}")
                    else:
                        print("Change deadlines operation completed.")
                        if isinstance(results, dict) and results:
                            for aid, status in results.items():
                                print(f"Assignment {aid}: {status}")
                except Exception as e:
                    if args.verbose:
                        print(f"[ChangeDeadlines] Error: {e}")
                    else:
                        print(f"Error changing deadlines: {e}")
            elif choice == '45':
                # Create Canvas groups
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                group_set_id = input("Enter group set ID (leave blank to select interactively, or 'q' to quit): ").strip()
                if group_set_id.lower() in ('q', 'quit'):
                    continue
                group_set_id = group_set_id if group_set_id else None
                num_groups_input = input("Enter number of groups to create (default: 5, or 'q' to quit): ").strip()
                if num_groups_input.lower() in ('q', 'quit'):
                    continue
                num_groups = int(num_groups_input) if num_groups_input.isdigit() else 5
                group_name_pattern = input("Enter group name pattern (default: 'Group {i}', or 'q' to quit): ").strip()
                if group_name_pattern.lower() in ('q', 'quit'):
                    continue
                if not group_name_pattern:
                    group_name_pattern = "Group {i}"
                try:
                    created_groups = create_canvas_groups(
                        api_url=CANVAS_LMS_API_URL,
                        api_key=CANVAS_LMS_API_KEY,
                        course_id=course_id,
                        group_set_id=group_set_id,
                        num_groups=num_groups,
                        group_name_pattern=group_name_pattern,
                        verbose=args.verbose
                    )
                    if created_groups:
                        print(f"Successfully created {len(created_groups)} groups.")
                    else:
                        print("Failed to create groups.")
                except Exception as e:
                    if args.verbose:
                        print(f"[CreateGroups] Error: {e}")
                    else:
                        print(f"Error creating groups: {e}")
            elif choice == '46':
                # Export student names and emails to TXT file
                export_emails_and_names_to_txt(students, args.export_emails_and_names, db_path=db_path, verbose=args.verbose)
            elif choice == '47':
                # Change assignment lock dates
                try:
                    course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                    if course_id.lower() in ('q', 'quit'):
                        continue
                    if not course_id:
                        course_id = CANVAS_LMS_COURSE_ID

                    raw_ids = input("Enter assignment IDs separated by commas (leave blank to select interactively, or 'q' to quit): ").strip()
                    if raw_ids.lower() in ('q', 'quit'):
                        continue
                    assignment_ids = None
                    if raw_ids:
                        assignment_ids = [part.strip() for part in raw_ids.split(",") if part.strip()]

                    new_lock = input("Enter new lock date to apply to all selected assignments (format: YYYY-MM-DD HH:MM) or leave blank to specify per-assignment (or 'q' to quit): ").strip()
                    if new_lock.lower() in ('q', 'quit'):
                        continue
                    new_lock_date = new_lock if new_lock else None

                    category = input("Enter assignment category (group) to filter when listing assignments (leave blank for all): ").strip()
                    if category.lower() in ('q', 'quit'):
                        continue
                    category = category if category else None

                    if args.verbose:
                        print(f"[ChangeLock] Calling change_canvas_lock_dates(course_id={course_id}, assignment_ids={assignment_ids}, new_lock_date={new_lock_date}, category={category})")
                    results = change_canvas_lock_dates(
                        assignment_ids=assignment_ids,
                        new_lock_date=new_lock_date,
                        api_url=CANVAS_LMS_API_URL,
                        api_key=CANVAS_LMS_API_KEY,
                        course_id=course_id,
                        category=category,
                        verbose=args.verbose
                    )
                    if args.verbose:
                        print(f"[ChangeLock] Results: {results}")
                    else:
                        print("Change lock dates operation completed.")
                        if isinstance(results, dict) and results:
                            for aid, status in results.items():
                                print(f"Assignment {aid}: {status}")
                except Exception as e:
                    if args.verbose:
                        print(f"[ChangeLock] Error: {e}")
                    else:
                        print(f"Error changing lock dates: {e}")
            elif choice == '48':
                # Export classroom roster to CSV file
                export_path = input_with_completion("Enter export file path (CSV, leave blank for default, or 'q' to quit): ").strip()
                if export_path.lower() in ('q', 'quit'):
                    continue
                if not export_path: 
                    export_path = os.path.join(os.getcwd(), "classroom_roster.csv")
                export_roster_to_csv(students, file_path=export_path, verbose=args.verbose)
            elif choice == '49':
                # Sync students with Google Classroom
                course_id = input("Enter Google Classroom course ID (leave blank to select interactively, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                course_id = course_id if course_id else None
                credentials_path = input_with_completion("Enter Google Classroom credentials JSON file path (default: gclassroom_credentials.json, or 'q' to quit): ").strip()
                if credentials_path.lower() in ('q', 'quit'):
                    continue
                if not credentials_path:
                    credentials_path = 'gclassroom_credentials.json'
                token_path = input_with_completion("Enter Google Classroom token pickle file path (default: token.pickle, or 'q' to quit): ").strip()
                if token_path.lower() in ('q', 'quit'):
                    continue
                if not token_path:
                    token_path = 'token.pickle'
                added, updated = sync_students_with_google_classroom(
                    students,
                    db_path=db_path,
                    course_id=course_id,
                    credentials_path=credentials_path,
                    token_path=token_path,
                    verbose=args.verbose
                )
                print(f"Sync with Google Classroom completed: {added} students added, {updated} students updated.")
                # Refresh students list after syncing
                students = load_database(db_path, verbose=args.verbose)
            elif choice == '50':
                # Delete empty Canvas groups
                course_id = input("Enter Canvas course ID (leave blank for default, or 'q' to quit): ").strip()
                if course_id.lower() in ('q', 'quit'):
                    continue
                if not course_id:
                    course_id = CANVAS_LMS_COURSE_ID
                group_set_id = input("Enter group set ID (leave blank to select interactively, or 'q' to quit): ").strip()
                if group_set_id.lower() in ('q', 'quit'):
                    continue
                group_set_id = group_set_id if group_set_id else None
                try:
                    deleted_count = delete_empty_canvas_groups(
                        api_url=CANVAS_LMS_API_URL,
                        api_key=CANVAS_LMS_API_KEY,
                        course_id=course_id,
                        group_set_id=group_set_id,
                        verbose=args.verbose
                    )
                    if deleted_count > 0:
                        print(f"Successfully deleted {deleted_count} empty groups.")
                    else:
                        print("No empty groups deleted.")
                except Exception as e:
                    if args.verbose:
                        print(f"[DeleteGroups] Error: {e}")
                    else:
                        print(f"Error deleting empty groups: {e}")
            else:
                print("Invalid option.")
