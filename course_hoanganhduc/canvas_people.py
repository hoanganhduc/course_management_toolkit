# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas people helpers."""

import requests

from .canvas_auth import get_canvas_client

from .data import load_database
from .utils import get_input_with_quit, parse_selection
from .settings import (
    CANVAS_LMS_API_URL,
    CANVAS_LMS_API_KEY,
    CANVAS_LMS_COURSE_ID,
)

def search_canvas_user(query, api_url=CANVAS_LMS_API_URL, api_key=CANVAS_LMS_API_KEY, course_id=CANVAS_LMS_COURSE_ID, verbose=False):
    """
    Search for a user in Canvas by name or email address and print details, including only total grades and assignment category grades.

    Args:
        query (str): Name or email to search for.
        api_url (str): Canvas API base URL.
        api_key (str): Canvas API key.
        course_id (str or int): Canvas course ID.
        verbose (bool): If True, print more details; otherwise, print only important notice.
    """
    try:
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)
        enrollments = course.get_enrollments(type=['TeacherEnrollment', 'TaEnrollment', 'StudentEnrollment'])
        found = False
        query_lower = query.strip().lower()
        for enrollment in enrollments:
            user = enrollment.user
            name = user.get('name', '').lower()
            email = user.get('login_id', '').lower()
            if query_lower in name or query_lower in email:
                if verbose:
                    print("[CanvasUser] User found:")
                else:
                    print("User found:")
                print(f"  Name: {user.get('name', '')}")
                print(f"  Email: {user.get('login_id', '')}")
                print(f"  Canvas ID: {user.get('id', '')}")
                print(f"  Role: {enrollment.type}")
                print(f"  Enrollment State: {enrollment.enrollment_state}")
                # Only print total grades and assignment category grades if student
                if enrollment.type.lower() == "studentenrollment":
                    try:
                        user_id = user.get('id')
                        total_grade = getattr(enrollment, "grades", {}).get("final_grade", None)
                        total_score = getattr(enrollment, "grades", {}).get("final_score", None)
                        print("  Grades:")
                        print(f"    - Total Grade: {total_grade} (Score: {total_score})")
                        assignment_groups = course.get_assignment_groups(include=['assignments', 'score_statistics'])
                        for group in assignment_groups:
                            group_score = None
                            group_possible = None
                            if hasattr(group, "score_statistics") and group.score_statistics:
                                stats = group.score_statistics
                                user_stats = stats.get(str(user_id))
                                if user_stats:
                                    group_score = user_stats.get("score")
                                    group_possible = user_stats.get("possible")
                            if group_score is None:
                                group_score = 0
                                group_possible = 0
                                for assignment in group.assignments:
                                    try:
                                        submission = assignment.get_submission(user_id)
                                        score = submission.score if hasattr(submission, "score") else None
                                        if score is not None:
                                            group_score += float(score)
                                        group_possible += float(getattr(assignment, "points_possible", 0))
                                    except Exception:
                                        continue
                            print(f"    - {group.name}: {group_score}/{group_possible}")
                        if verbose:
                            print(f"[CanvasUser] Printed grades for user {user.get('name', '')}")
                    except Exception as e:
                        if verbose:
                            print(f"[CanvasUser] Could not fetch assignments/grades: {e}")
                        else:
                            print(f"  Could not fetch assignments/grades: {e}")
                found = True
        if not found:
            if verbose:
                print("[CanvasUser] No user found matching the query.")
            else:
                print("No user found matching the query.")
    except Exception as e:
        if verbose:
            print(f"[CanvasUser] Error searching for user: {e}")
        else:
            print(f"Error searching for user: {e}")


def list_canvas_people(api_url=CANVAS_LMS_API_URL, api_key=CANVAS_LMS_API_KEY, course_id=CANVAS_LMS_COURSE_ID, verbose=False):
    """
    List all teachers, TAs, and students in a Canvas course using the canvasapi library.
    For students, divides into active students and invited students.

    Args:
        api_url (str): The base URL for the Canvas instance.
        api_key (str): Your Canvas API access token.
        course_id (int or str): The Canvas course ID.
        verbose (bool): If True, print more details; otherwise, print only important notice.

    Returns:
        Dict with keys 'teachers', 'tas', 'active_students', 'invited_students', each a list of dicts with id, name, email, and canvas_id.
    """
    try:
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)
        enrollments = course.get_enrollments(type=['TeacherEnrollment', 'TaEnrollment', 'StudentEnrollment'])
        people = {'teachers': [], 'tas': [], 'active_students': [], 'invited_students': []}
        for enrollment in enrollments:
            user = enrollment.user
            role = enrollment.type.lower()
            person = {
                "id": user['id'],
                "canvas_id": user['id'],
                "name": user.get('name', ''),
                "email": user.get('login_id', '')
            }
            if role == 'teacherenrollment':
                people['teachers'].append(person)
            elif role == 'taenrollment':
                people['tas'].append(person)
            elif role == 'studentenrollment':
                state = enrollment.enrollment_state.lower()
                if state == 'active':
                    people['active_students'].append(person)
                elif state == 'invited':
                    people['invited_students'].append(person)
        return people
    except ImportError:
        return {'teachers': [], 'tas': [], 'active_students': [], 'invited_students': []}
    except Exception as e:
        return {'teachers': [], 'tas': [], 'active_students': [], 'invited_students': []}

def print_canvas_people(people, verbose=False):
    """
    Print the Canvas people dictionary returned by list_canvas_people.
    """
    if people.get('teachers'):
        if verbose:
            print(f"[CanvasPeople] Teachers ({len(people['teachers'])}):")
        else:
            print(f"\nTeachers ({len(people['teachers'])}):")
        for t in people['teachers']:
            if verbose:
                print(f"[CanvasPeople]   {t['name']} ({t['email']}), Canvas ID: {t['canvas_id']}")
            else:
                print(f"  {t['name']} ({t['email']}), Canvas ID: {t['canvas_id']}")
    if people.get('tas'):
        if verbose:
            print(f"[CanvasPeople] TAs ({len(people['tas'])}):")
        else:
            print(f"\nTAs ({len(people['tas'])}):")
        for ta in people['tas']:
            if verbose:
                print(f"[CanvasPeople]   {ta['name']} ({ta['email']}), Canvas ID: {ta['canvas_id']}")
            else:
                print(f"  {ta['name']} ({ta['email']}), Canvas ID: {ta['canvas_id']}")
    if people.get('active_students'):
        if verbose:
            print(f"[CanvasPeople] Active Students ({len(people['active_students'])}):")
        else:
            print(f"\nActive Students ({len(people['active_students'])}):")
        for s in people['active_students']:
            if verbose:
                print(f"[CanvasPeople]   {s['name']} ({s['email']}), Canvas ID: {s['canvas_id']}")
            else:
                print(f"  {s['name']} ({s['email']}), Canvas ID: {s['canvas_id']}")
    if people.get('invited_students'):
        if verbose:
            print(f"[CanvasPeople] Invited Students ({len(people['invited_students'])}):")
        else:
            print(f"\nInvited Students ({len(people['invited_students'])}):")
        for s in people['invited_students']:
            if verbose:
                print(f"[CanvasPeople]   {s['name']} ({s['email']}), Canvas ID: {s['canvas_id']}")
            else:
                print(f"  {s['name']} ({s['email']}), Canvas ID: {s['canvas_id']}")


def unenroll_canvas_students(
    course_id=None,
    domains=None,
    emails=None,
    select_mode=False,
    apply_all=False,
    missing_student_id=False,
    db_path=None,
    dry_run=False,
    update_local_db=True,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    verbose=False,
):
    """
    Unenroll students from a Canvas course by email domain, email list, selection, or missing Student ID in local DB.
    """
    def normalize_domains(value):
        if not value:
            return []
        raw = value
        if not isinstance(raw, (list, tuple, set)):
            raw = [v.strip() for v in str(raw).split(",") if v.strip()]
        domains_list = []
        for d in raw:
            d = str(d).strip().lower()
            if d.startswith("@"):
                d = d[1:]
            if d:
                domains_list.append(d)
        return domains_list

    def normalize_emails(value):
        if not value:
            return []
        raw = value
        if not isinstance(raw, (list, tuple, set)):
            raw = [v.strip() for v in str(raw).split(",") if v.strip()]
        return [str(v).strip().lower() for v in raw if str(v).strip()]

    canvas = get_canvas_client(api_url, api_key)
    course = canvas.get_course(course_id)

    domains_list = normalize_domains(domains)
    emails_list = normalize_emails(emails)

    missing_email_set = set()
    missing_canvas_id_set = set()
    if missing_student_id:
        if not db_path:
            print("Missing Student ID mode requires a local database. Provide --db or run from the DB folder.")
            return 0
        try:
            students = load_database(db_path, verbose=verbose)
        except Exception:
            students = []
        for student in students or []:
            sid = (getattr(student, "Student ID", "") or "").strip()
            if sid:
                continue
            email = (getattr(student, "Email", "") or "").strip().lower()
            canvas_id = (getattr(student, "Canvas ID", "") or "").strip()
            if email:
                missing_email_set.add(email)
            if canvas_id:
                missing_canvas_id_set.add(canvas_id)
        if not missing_email_set and not missing_canvas_id_set:
            print("No students missing Student ID in the local database.")
            return 0

    roster = []
    enrollments = course.get_enrollments(type=["StudentEnrollment"])
    for enrollment in enrollments:
        user = enrollment.user or {}
        email = (user.get("login_id") or "").strip()
        name = user.get("name") or ""
        user_id = user.get("id")
        state = getattr(enrollment, "enrollment_state", "")
        enrollment_id = getattr(enrollment, "id", None)
        if not enrollment_id or not user_id:
            continue
        roster.append({
            "enrollment_id": str(enrollment_id),
            "user_id": str(user_id),
            "name": name,
            "email": email,
            "state": state,
        })

    matches = []
    for entry in roster:
        email = entry.get("email", "")
        email_lower = email.lower()
        if select_mode:
            matches.append(entry)
            continue
        if missing_student_id:
            if (email_lower and email_lower in missing_email_set) or (entry["user_id"] in missing_canvas_id_set):
                matches.append(entry)
            continue
        if emails_list and email_lower in emails_list:
            matches.append(entry)
            continue
        if domains_list and any(email_lower.endswith("@" + d) for d in domains_list):
            matches.append(entry)

    if not matches:
        print("No matching students found.")
        return 0

    matches.sort(key=lambda m: (m["name"], m["email"]))
    if select_mode:
        print(f"Found {len(matches)} enrolled student(s).")
    else:
        label_parts = []
        if domains_list:
            label_parts.append("domain(s): " + ", ".join(domains_list))
        if emails_list:
            label_parts.append("email(s): " + ", ".join(emails_list))
        if missing_student_id:
            label_parts.append("missing student id")
        label = "; ".join(label_parts)
        print(f"Matched {len(matches)} student(s) for {label}.")
    for idx, m in enumerate(matches, 1):
        state = m.get("state") or ""
        print(f"{idx}. {m['name']} | {m['email']} | {state}")

    if apply_all:
        selected_indices = list(range(1, len(matches) + 1))
    else:
        while True:
            sel = get_input_with_quit("Select students to unenroll (e.g. 1,3-5, 'a' for all, or 'q' to quit): ")
            if sel is None:
                return 0
            selected_indices = parse_selection(sel, len(matches))
            if selected_indices:
                break
            print("Invalid selection. Please enter valid numbers, a range, 'a' for all, or 'q' to quit.")

    confirm = get_input_with_quit(
        f"Unenroll {len(selected_indices)} student(s) from the course? (y/n): ",
        default="n"
    )
    if confirm is None or confirm.lower() not in ("y", "yes"):
        print("Unenroll canceled.")
        return 0

    if dry_run:
        print(f"Dry-run: would unenroll {len(selected_indices)} student(s).")
        return len(selected_indices)

    removed = 0
    removed_emails = []
    removed_canvas_ids = []
    headers = {"Authorization": f"Bearer {api_key}"}
    for idx in selected_indices:
        entry = matches[idx - 1]
        enrollment_id = entry["enrollment_id"]
        url = f"{api_url}/api/v1/courses/{course_id}/enrollments/{enrollment_id}?task=delete"
        try:
            resp = requests.delete(url, headers=headers)
            if resp.status_code in (200, 204):
                removed += 1
                if entry.get("email"):
                    removed_emails.append(entry["email"].strip().lower())
                if entry.get("user_id"):
                    removed_canvas_ids.append(str(entry["user_id"]))
                if verbose:
                    print(f"[CanvasUnenroll] Unenrolled {entry['name']} ({entry['email']}).")
            else:
                if verbose:
                    print(f"[CanvasUnenroll] Failed to unenroll {entry['name']} ({entry['email']}): {resp.status_code} {resp.text}")
                else:
                    print(f"Failed to unenroll {entry['name']} ({entry['email']}).")
        except Exception as e:
            if verbose:
                print(f"[CanvasUnenroll] Failed to unenroll {entry['name']} ({entry['email']}): {e}")
            else:
                print(f"Failed to unenroll {entry['name']} ({entry['email']}).")

    print(f"Unenrolled {removed} student(s).")

    if removed and update_local_db and db_path:
        try:
            students = load_database(db_path, verbose=verbose)
        except Exception:
            students = []
        if students:
            kept = []
            removed_local = 0
            removed_email_set = set(removed_emails)
            removed_canvas_id_set = set(removed_canvas_ids)
            for student in students:
                email = (getattr(student, "Email", "") or "").strip().lower()
                canvas_id = (getattr(student, "Canvas ID", "") or "").strip()
                if (email and email in removed_email_set) or (canvas_id and canvas_id in removed_canvas_id_set):
                    removed_local += 1
                    continue
                kept.append(student)
            if removed_local:
                save_database(kept, db_path=db_path, verbose=verbose, audit_source="canvas-unenroll")
                print(f"Removed {removed_local} student(s) from local database.")
    return removed


