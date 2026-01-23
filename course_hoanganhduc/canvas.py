# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/
# Course Management Script

"""Canvas LMS helpers."""

from .canvas_calendar import import_canvas_calendar_from_ics
from .canvas_assignments import list_canvas_assignments
from .canvas_people import list_canvas_people, print_canvas_people, search_canvas_user, unenroll_canvas_students
from .canvas_announcements import add_canvas_announcement, send_final_evaluations_via_canvas
from .canvas_messages import (
    send_canvas_message_to_students,
    fetch_and_reply_canvas_messages,
    notify_incomplete_canvas_peer_reviews,
)
from .canvas_pages import list_and_update_canvas_pages
from .canvas_submissions import (
    download_canvas_assignment_submissions,
    download_canvas_assignment_submissions_auto,
    add_comment_to_canvas_submission,
)
from .canvas_checks import (
    compare_texts_from_pdfs_in_folder,
    detect_meaningful_level_and_notify_students,
    extract_canvas_id_from_filename,
)
from .canvas_weekly import (
    notify_missing_submissions_after_due,
    run_weekly_canvas_automation,
    list_closed_assignments_for_weekly_automation,
)
from .canvas_grading import (
    download_and_check_student_submissions,
    grade_canvas_assignment_submissions,
    list_students_with_multiple_submissions_on_time,
    grade_resubmissions,
)
from .canvas_invites import (
    invite_students_if_not_enrolled,
    invite_user_to_canvas_course,
    invite_users_to_canvas_course,
)
from .canvas_admin import (
    change_canvas_deadlines,
    change_canvas_lock_dates,
    create_canvas_groups,
    delete_empty_canvas_groups,
)
from .canvas_rubrics import (
    list_and_export_canvas_rubrics,
    import_canvas_rubrics,
    update_canvas_rubrics_for_assignments,
)
from .canvas_grading_schemes import (
    list_and_download_canvas_grading_standards,
    add_canvas_grading_scheme,
)
from .canvas_sync import sync_students_with_canvas

from .data import load_database, save_database
from .models import Student
from .utils import get_input_with_timeout, prefill_input_with_timeout

def interactive_modify_database(students, db_path=None, verbose=False):
    """
    Interactively modify student records in the database.
    Allows searching, editing, adding, and deleting students.
    When editing a field, pre-fill the old value for easy modification.
    After editing a field, ask if the user wants to continue editing other fields of the same student,
    edit another student, or quit. Always allow quitting at any step.
    If no response after 60 seconds from user then quit.
    If verbose is True, print more details; otherwise, print only important notice.
    """
    try:
        if db_path:
            students = load_database(db_path, verbose=verbose)
        if not students:
            if verbose:
                print("[ModifyDB] No students in the database.")
            else:
                print("No students in the database.")
            return

        def list_students():
            if verbose:
                print("[ModifyDB] List of students:")
            else:
                print("List of students:")
            for idx, s in enumerate(students, 1):
                name = getattr(s, "Name", "")
                sid = getattr(s, "Student ID", "")
                print(f"{idx}. {name} ({sid})")

        while True:
            if verbose:
                print("\n[ModifyDB] Modify Menu:")
            else:
                print("\nModify Menu:")
            print("1. List students")
            print("2. Edit a student")
            print("3. Add a new student")
            print("4. Delete a student")
            print("0. Exit modify menu")
            
            try:
                choice = get_input_with_timeout("Choose an option (or 'q' to quit): ").strip()
            except TimeoutError:
                return
            except KeyboardInterrupt:
                return
            
            if choice in ("0", "q", "Q"):
                break
            elif choice == "1":
                list_students()
            elif choice == "2":
                while True:
                    list_students()
                    try:
                        idx = get_input_with_timeout("Enter the number of the student to edit (or 'q' to quit): ").strip()
                    except TimeoutError:
                        return
                    except KeyboardInterrupt:
                        return
                    
                    if idx.lower() in ("q", "quit"):
                        break
                    if not idx.isdigit() or int(idx) < 1 or int(idx) > len(students):
                        if verbose:
                            print("[ModifyDB] Invalid index.")
                        else:
                            print("Invalid index.")
                        continue
                    s = students[int(idx) - 1]
                    while True:
                        if verbose:
                            print("[ModifyDB] Current fields:")
                        else:
                            print("Current fields:")
                        for k, v in s.__dict__.items():
                            print(f"{k}: {v}")
                        
                        try:
                            field = get_input_with_timeout("Enter field to edit (or leave blank to cancel, or 'q' to quit): ").strip()
                        except TimeoutError:
                            return
                        except KeyboardInterrupt:
                            return
                        
                        if not field or field.lower() == "q":
                            break
                        old_value = getattr(s, field, "")
                        # Pre-fill old value for editing
                        try:
                            value = prefill_input_with_timeout(f"Enter new value for '{field}' [{old_value}]: ", old_value)
                        except TimeoutError:
                            return
                        except KeyboardInterrupt:
                            return
                        except Exception:
                            try:
                                value = get_input_with_timeout(f"Enter new value for '{field}' (old: {old_value}): ").strip()
                                if not value:
                                    value = old_value
                            except TimeoutError:
                                return
                            except KeyboardInterrupt:
                                return
                        
                        if value.lower() == "q":
                            break
                        setattr(s, field, value)
                        if verbose:
                            print(f"[ModifyDB] Student updated: {field} = {value}")
                        else:
                            print("Student updated.")
                        if db_path:
                            save_database(students, db_path, verbose=verbose)
                        # Ask if continue editing this student, edit another student, or quit
                        try:
                            next_action = get_input_with_timeout("Continue editing other fields of this student (c), edit another student (a), or quit (q)? [c/a/q]: ").strip().lower()
                        except TimeoutError:
                            return
                        except KeyboardInterrupt:
                            return
                        
                        if next_action == "q":
                            return
                        elif next_action == "a":
                            break
                        # else continue editing this student
            elif choice == "3":
                fields = {}
                if verbose:
                    print("[ModifyDB] Enter new student information (leave blank to skip a field, or 'q' to quit):")
                else:
                    print("Enter new student information (leave blank to skip a field, or 'q' to quit):")
                for field in ["Name", "Student ID", "Email", "Class"]:
                    try:
                        value = get_input_with_timeout(f"{field}: ").strip()
                    except TimeoutError:
                        return
                    except KeyboardInterrupt:
                        return
                    
                    if value.lower() == "q":
                        if verbose:
                            print("[ModifyDB] Cancelled adding new student.")
                        else:
                            print("Cancelled adding new student.")
                        break
                    if value:
                        fields[field] = value
                else:
                    students.append(Student(**fields))
                    if verbose:
                        print(f"[ModifyDB] Student added: {fields}")
                    else:
                        print("Student added.")
                    if db_path:
                        save_database(students, db_path, verbose=verbose)
            elif choice == "4":
                while True:
                    list_students()
                    try:
                        idx = get_input_with_timeout("Enter the number of the student to delete (or 'q' to quit): ").strip()
                    except TimeoutError:
                        return
                    except KeyboardInterrupt:
                        return
                    
                    if idx.lower() in ("q", "quit"):
                        break
                    if not idx.isdigit() or int(idx) < 1 or int(idx) > len(students):
                        if verbose:
                            print("[ModifyDB] Invalid index.")
                        else:
                            print("Invalid index.")
                        continue
                    del students[int(idx) - 1]
                    if verbose:
                        print("[ModifyDB] Student deleted.")
                    else:
                        print("Student deleted.")
                    if db_path:
                        save_database(students, db_path, verbose=verbose)
                    break
            else:
                if verbose:
                    print("[ModifyDB] Invalid option.")
                else:
                    print("Invalid option.")
    
    except TimeoutError:
        if verbose:
            print("\n[ModifyDB] Timeout occurred. Exiting modify menu.")
        else:
            print("\nTimeout occurred. Exiting modify menu.")
        return
    except KeyboardInterrupt:
        if verbose:
            print("\n[ModifyDB] Operation cancelled by user. Exiting modify menu.")
        else:
            print("\nOperation cancelled by user. Exiting modify menu.")
        return

