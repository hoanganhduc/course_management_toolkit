# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas grading helpers."""

import random
import requests
from datetime import datetime
import time

from .canvas_auth import get_canvas_client

from .canvas_utils import prompt_with_timeout, parse_selection, normalize_datetime_str
from .settings import (
    ALL_AI_METHODS,
    CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    CANVAS_LMS_API_KEY,
    CANVAS_LMS_API_URL,
    CANVAS_LMS_COURSE_ID,
    DEFAULT_AI_METHOD,
    DEFAULT_OCR_METHOD,
    DEFAULT_RESTRICTED,
)
from .data import refine_text_with_ai

def grade_canvas_assignment_submissions(
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    verbose=False,
    restricted=DEFAULT_RESTRICTED  # New option: if False, list all assignments with submissions and all students who submitted
):
    """
    List all assignments with at least one submission and whose lock date has passed (if restricted=True),
    or all assignments with at least one submission (if restricted=False).
    Allow user to select one or more (supports ranges).
    For each selected assignment, list all students who submitted but have not yet been given scores (if restricted=True),
    or all students who submitted (if restricted=False).
    Allow user to select one or more (supports ranges), then prompt for a score (optionally same for all), and set the score for those submissions.
    Allow quitting at any step.
    If verbose is True, print more details; otherwise, print only important notice.
    """
    def get_input_with_quit(prompt, default=None):
        try:
            val = input(prompt)
            if not val and default is not None:
                return default
            if val.strip().lower() in ("q", "quit"):
                return None
            return val
        except KeyboardInterrupt:
            return None

    try:
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)

        # Step 1: List assignments with at least one submission
        assignments = []
        assignment_groups = list(course.get_assignment_groups(include=['assignments']))
        now = datetime.utcnow()
        for group in assignment_groups:
            group_name = group.name
            if category and group_name.lower() != category.lower():
                continue
            for assignment in group.assignments:
                if assignment.get('has_submitted_submissions', False):
                    if restricted:
                        lock_at = assignment.get('lock_at')
                        if lock_at:
                            try:
                                lock_dt = datetime.strptime(lock_at, "%Y-%m-%dT%H:%M:%SZ")
                                if lock_dt > now:
                                    continue  # Skip if lock date is in the future
                            except Exception:
                                continue  # Skip if lock date is invalid
                        else:
                            continue  # Skip if no lock date
                    assignments.append({
                        "id": assignment['id'],
                        "name": assignment['name'],
                        "group": group_name,
                        "due_at": assignment.get('due_at'),
                        "lock_at": assignment.get('lock_at')
                    })
        if not assignments:
            msg = "No assignments found with submissions and lock date passed." if restricted else "No assignments found with submissions."
            if verbose:
                print(f"[GradeCanvas] {msg}")
            else:
                print(msg)
            return

        # Sort by due date
        def due_sort_key(a):
            raw = a.get("due_at")
            if raw:
                try:
                    return datetime.strptime(raw, "%Y-%m-%dT%H:%M:%SZ")
                except Exception:
                    return datetime.max
            return datetime.max
        assignments.sort(key=due_sort_key)

        if verbose:
            print("[GradeCanvas] Assignments with at least one submission{}:".format(" and lock date passed" if restricted else ""))
        else:
            print("Assignments with at least one submission{}:".format(" and lock date passed" if restricted else ""))
        for idx, a in enumerate(assignments, 1):
            due = a['due_at'] or "No due date"
            lock = a['lock_at'] or "No lock date"
            if verbose:
                print(f"[GradeCanvas] {idx}. [{a['group']}] {a['name']} (ID: {a['id']}, Due: {due}, Lock: {lock})")
            else:
                print(f"{idx}. [{a['group']}] {a['name']} (ID: {a['id']}, Due: {due}, Lock: {lock})")

        # Step 2: Select assignments to grade (supports ranges, comma-separated, 'a' for all)
        while True:
            sel = get_input_with_quit(
                "Enter the number(s) of the assignment(s) to grade (e.g. 1,3-5 or 'a' for all, or 'q' to quit): "
            )
            if sel is None:
                if verbose:
                    print("[GradeCanvas] Quitting.")
                else:
                    print("Quitting.")
                return
            if sel.lower() in ('a', 'all'):
                selected_assignments = list(range(1, len(assignments) + 1))
            else:
                selected_assignments = set()
                for part in sel.split(","):
                    part = part.strip()
                    if "-" in part:
                        try:
                            start, end = map(int, part.split("-"))
                            selected_assignments.update(range(start, end + 1))
                        except Exception:
                            continue
                    elif part.isdigit():
                        selected_assignments.add(int(part))
                selected_assignments = [i for i in selected_assignments if 1 <= i <= len(assignments)]
            if not selected_assignments:
                msg = "Invalid selection. Please enter valid numbers, a range, 'a' for all, or 'q' to quit."
                if verbose:
                    print(f"[GradeCanvas] {msg}")
                else:
                    print(msg)
                continue
            break

        # Step 3: For each selected assignment, process grading
        for assign_idx in selected_assignments:
            assignment_info = assignments[assign_idx - 1]
            assignment_id = assignment_info['id']
            assignment = course.get_assignment(assignment_id)
            if verbose:
                print(f"\n[GradeCanvas] Grading assignment: [{assignment_info['group']}] {assignment_info['name']} (ID: {assignment_id})")
            else:
                print(f"\n--- Grading assignment: [{assignment_info['group']}] {assignment_info['name']} (ID: {assignment_id}) ---")

            # List all students who submitted
            submissions = list(assignment.get_submissions(include=["user"]))
            students_to_grade = []
            for sub in submissions:
                user = getattr(sub, "user", {})
                student_name = user.get("name", "UnknownStudent")
                canvas_id = user.get("id", "unknown")
                submitted_at = getattr(sub, "submitted_at", None)
                score = getattr(sub, "score", None)
                # Only include if submitted (on time or late)
                if submitted_at:
                    if restricted:
                        # Only include if score is None or 0
                        if score is None or score == "0":
                            students_to_grade.append({
                                "canvas_id": canvas_id,
                                "name": student_name,
                                "submitted_at": submitted_at,
                                "submission": sub
                            })
                    else:
                        students_to_grade.append({
                            "canvas_id": canvas_id,
                            "name": student_name,
                            "submitted_at": submitted_at,
                            "score": score,
                            "submission": sub
                        })
            if not students_to_grade:
                msg = "No students have{} submissions for this assignment.".format(" ungraded" if restricted else "")
                if verbose:
                    print(f"[GradeCanvas] {msg}")
                else:
                    print(msg)
                continue

            students_to_grade.sort(key=lambda s: s["name"])
            if verbose:
                print("[GradeCanvas] Students who submitted{}:".format(" but have not yet been graded" if restricted else ""))
            else:
                print("Students who submitted{}:".format(" but have not yet been graded" if restricted else ""))
            for idx, s in enumerate(students_to_grade, 1):
                submitted = s["submitted_at"] or "No submission"
                score_str = f", Current Score: {s.get('score', '')}" if not restricted else ""
                if verbose:
                    print(f"[GradeCanvas] {idx}. {s['name']} (Canvas ID: {s['canvas_id']}), Submitted at: {submitted}{score_str}")
                else:
                    print(f"{idx}. {s['name']} (Canvas ID: {s['canvas_id']}), Submitted at: {submitted}{score_str}")

            # Select students to grade (supports ranges, comma-separated, 'a' for all)
            while True:
                sel = get_input_with_quit(
                    "Enter the number(s) of the student(s) to grade (e.g. 1,3-5 or 'a' for all, or 'q' to quit): "
                )
                if sel is None:
                    if verbose:
                        print("[GradeCanvas] Quitting grading for this assignment.")
                    else:
                        print("Quitting grading for this assignment.")
                    break
                if sel.lower() in ('a', 'all'):
                    selected = list(range(1, len(students_to_grade) + 1))
                else:
                    selected = set()
                    for part in sel.split(","):
                        part = part.strip()
                        if "-" in part:
                            try:
                                start, end = map(int, part.split("-"))
                                selected.update(range(start, end + 1))
                            except Exception:
                                continue
                        elif part.isdigit():
                            selected.add(int(part))
                    selected = [i for i in selected if 1 <= i <= len(students_to_grade)]
                if not selected:
                    msg = "Invalid selection. Please enter valid numbers, a range, 'a' for all, or 'q' to quit."
                    if verbose:
                        print(f"[GradeCanvas] {msg}")
                    else:
                        print(msg)
                    continue
                break
            if not selected:
                continue

            # Ask for score to give (optionally same for all)
            score = get_input_with_quit("Enter the score to assign to all selected submissions (or 'q' to quit): ")
            if score is None:
                if verbose:
                    print("[GradeCanvas] Quitting grading for this assignment.")
                else:
                    print("Quitting grading for this assignment.")
                continue
            try:
                score = float(score)
            except Exception:
                msg = "Invalid score. Skipping this assignment."
                if verbose:
                    print(f"[GradeCanvas] {msg}")
                else:
                    print(msg)
                continue

            # Confirm and assign scores
            msg = f"Assigning score {score} to {len(selected)} submission(s) for assignment '{assignment_info['name']}'."
            if verbose:
                print(f"[GradeCanvas] {msg}")
            else:
                print(msg)
            confirm = get_input_with_quit("Proceed? (y/n): ", default="y")
            if confirm is None or confirm.lower() not in ("y", "yes", ""):
                if verbose:
                    print("[GradeCanvas] Aborted grading for this assignment.")
                else:
                    print("Aborted grading for this assignment.")
                continue

            for idx in selected:
                s = students_to_grade[idx - 1]
                sub = s["submission"]
                try:
                    sub.edit(submission={'posted_grade': score})
                    if verbose:
                        print(f"[GradeCanvas] Assigned score {score} to {s['name']} (Canvas ID: {s['canvas_id']})")
                    else:
                        print(f"Assigned score {score} to {s['name']} (Canvas ID: {s['canvas_id']})")
                except Exception as e:
                    if verbose:
                        print(f"[GradeCanvas] Failed to assign score to {s['name']} (Canvas ID: {s['canvas_id']}): {e}")
                    else:
                        print(f"Failed to assign score to {s['name']} (Canvas ID: {s['canvas_id']}): {e}")

            if verbose:
                print(f"[GradeCanvas] Done grading selected submissions for assignment '{assignment_info['name']}'.")
            else:
                print(f"Done grading selected submissions for assignment '{assignment_info['name']}'.")

        if verbose:
            print("[GradeCanvas] Finished grading all selected assignments.")
        else:
            print("Finished grading all selected assignments.")

    except Exception as e:
        if verbose:
            print(f"[GradeCanvas] Error grading submissions: {e}")
        else:
            print(f"Error grading submissions: {e}")


def list_students_with_multiple_submissions_on_time(
    assignment_id=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    verbose=False
):
    """
    List all students who submitted at least twice to an assignment,
    where the first submission is on time and the second submission is late.
    If assignment_id is None, list all assignments in the specified category and ask user to select one or more.

    For students who have first submission on time and from second submission late,
    update the late_policy_status from the second submission to "none".

    Args:
        assignment_id (str or int or None): Canvas assignment ID. If None, prompt user to select.
        api_url, api_key, course_id: Canvas API info.
        category (str): Assignment group/category to filter.
        verbose (bool): Print more details.

    Returns:
        List of dicts: [{canvas_id, name, submissions: [submission_times], first_on_time: True/False, second_late: True/False}, ...]
    """

    def timeout_handler(signum, frame):
        # Arguments are required by signal handler signature but not used
        print("\nTimeout: No response after 60 seconds. Using default: all")
        raise TimeoutError("User input timeout")

    def _prompt(prompt, timeout=60, default=None):
        return prompt_with_timeout(prompt, timeout=timeout, default=default, verbose=verbose)

    try:
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)

        # If assignment_id is None, list all assignments in the specified category and prompt user
        assignment_ids = []
        if assignment_id is None:
            assignments = []
            assignment_groups = list(course.get_assignment_groups(include=['assignments']))
            for group in assignment_groups:
                group_name = group.name
                if category and group_name.lower() != category.lower():
                    continue
                for assignment in group.assignments:
                    if assignment.get('has_submitted_submissions', False):
                        assignments.append({
                            "id": assignment['id'],
                            "name": assignment['name'],
                            "group": group_name,
                            "due_at": assignment.get('due_at')
                        })
            if not assignments:
                msg = f"No assignments found with submissions in category '{category}'." if category else "No assignments found."
                if verbose:
                    print(f"[MultipleSubmissions] {msg}")
                else:
                    print(msg)
                return []
            # Sort by due date
            def due_sort_key(a):
                raw = a.get("due_at")
                if raw:
                    try:
                        return datetime.strptime(raw, "%Y-%m-%dT%H:%M:%SZ")
                    except Exception:
                        return datetime.max
                return datetime.max
            assignments.sort(key=due_sort_key)
            print("Assignments with at least one submission:")
            for idx, a in enumerate(assignments, 1):
                due = a['due_at'] or "No due date"
                print(f"{idx}. [{a['group']}] {a['name']} (ID: {a['id']}, Due: {due})")
            while True:
                try:
                    sel = _prompt(
                        "Enter the number(s) of the assignment(s) to check (e.g. 1,3-5 or 'a' for all, or 'q' to quit): ",
                        timeout=60,
                        default="a"
                    ).strip()
                except TimeoutError:
                    sel = "a"
                if sel.lower() in ('q', 'quit'):
                    return []
                if sel.lower() in ('a', 'all'):
                    assignment_ids = [a['id'] for a in assignments]
                    break
                selected = set()
                for part in sel.split(","):
                    part = part.strip()
                    if "-" in part:
                        try:
                            start, end = map(int, part.split("-"))
                            selected.update(range(start, end + 1))
                        except Exception:
                            continue
                    elif part.isdigit():
                        selected.add(int(part))
                selected = [i for i in selected if 1 <= i <= len(assignments)]
                if not selected:
                    print("Invalid selection. Please enter valid numbers, a range, 'a' for all, or 'q' to quit.")
                    continue
                assignment_ids = [assignments[i - 1]['id'] for i in selected]
                break
        else:
            assignment_ids = [assignment_id]

        all_results = []
        for aid in assignment_ids:
            assignment = course.get_assignment(aid)
            due_at = getattr(assignment, "due_at", None)
            if due_at:
                try:
                    due_dt = datetime.strptime(due_at, "%Y-%m-%dT%H:%M:%SZ")
                except Exception:
                    due_dt = None
            else:
                due_dt = None

            submissions = list(assignment.get_submissions(include=["user", "submission_history"]))
            results = []
            for sub in submissions:
                user = getattr(sub, "user", {})
                canvas_id = user.get("id", None)
                name = user.get("name", "")
                # Get all submission attempts (submission_history)
                history = getattr(sub, "submission_history", None)
                if not history:
                    # Fallback: treat as single submission
                    submitted_at = getattr(sub, "submitted_at", None)
                    if submitted_at:
                        times = [submitted_at]
                        histories = [sub]
                    else:
                        times = []
                        histories = []
                else:
                    times = [h.get("submitted_at") for h in history if h.get("submitted_at")]
                    histories = [h for h in history if h.get("submitted_at")]
                # Only consider students with 2 or more submissions
                if len(times) < 2:
                    continue
                # Sort submission times chronologically, keep mapping to histories
                times_histories = sorted(zip(times, histories), key=lambda x: x[0])
                times_sorted = [t for t, h in times_histories]
                histories_sorted = [h for t, h in times_histories]
                first_time = times_sorted[0]
                second_time = times_sorted[1]
                if not first_time or not second_time or not due_dt:
                    continue
                try:
                    first_dt = datetime.strptime(first_time, "%Y-%m-%dT%H:%M:%SZ")
                    second_dt = datetime.strptime(second_time, "%Y-%m-%dT%H:%M:%SZ")
                except Exception:
                    continue
                first_on_time = first_dt <= due_dt
                second_late = second_dt > due_dt
                if first_on_time and second_late:
                    # Update late_policy_status from second submission onward to "none"
                    for idx in range(1, len(histories_sorted)):
                        h = histories_sorted[idx]
                        # Only update if late_policy_status is not "none"
                        if h.get("late_policy_status") != "none":
                            try:
                                # Use the API to update the submission's late_policy_status
                                # Canvas API: PUT /api/v1/courses/:course_id/assignments/:assignment_id/submissions/:user_id
                                url = f"{api_url}/api/v1/courses/{course_id}/assignments/{aid}/submissions/{canvas_id}"
                                headers = {
                                    "Authorization": f"Bearer {api_key}",
                                    "Content-Type": "application/json"
                                }
                                data = {
                                    "submission": {
                                        "late_policy_status": "none"
                                    }
                                }
                                resp = requests.put(url, headers=headers, json=data)
                                if verbose:
                                    print(f"[MultipleSubmissions] Updated late_policy_status to 'none' for student {name} (ID: {canvas_id}), submission at {histories_sorted[idx].get('submitted_at')}, status code: {resp.status_code}")
                            except Exception as e:
                                if verbose:
                                    print(f"[MultipleSubmissions] Failed to update late_policy_status for student {name} (ID: {canvas_id}): {e}")
                    results.append({
                        "canvas_id": canvas_id,
                        "name": name,
                        "submissions": times_sorted,
                        "first_on_time": True,
                        "second_late": True,
                        "assignment_id": aid,
                        "assignment_name": getattr(assignment, "name", "")
                    })
            if verbose:
                print(f"[MultipleSubmissions] Assignment {aid}: Found {len(results)} students with first submission on time and second late:")
                for r in results:
                    print(f"  {r['name']} (ID: {r['canvas_id']}): {len(r['submissions'])} submissions, first: {r['submissions'][0]}, second: {r['submissions'][1]}")
            else:
                print(f"Assignment {aid}: Found {len(results)} students with first submission on time and second late.")
            all_results.extend(results)
        return all_results
    except Exception as e:
        if verbose:
            print(f"[MultipleSubmissions] Error: {e}")
        else:
            print(f"Error: {e}")
        return []


def grade_resubmissions(
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    assignment_ids=None,
    keep_old_grade=False,
    verbose=False
):
    """
    For each assignment, find resubmissions after a graded attempt and prompt to regrade them.

    If assignment_ids is None, prompt to select from assignments with submissions.
    """
    def _fetch_assignment_submissions(aid):
        submissions = []
        url = f"{api_url}/api/v1/courses/{course_id}/assignments/{aid}/submissions"
        headers = {"Authorization": f"Bearer {api_key}"}
        params = {
            "include[]": ["submission_history", "user"],
            "per_page": 100
        }
        while url:
            resp = requests.get(url, headers=headers, params=params)
            resp.raise_for_status()
            submissions.extend(resp.json() or [])
            next_url = None
            link_header = resp.headers.get("Link")
            if link_header:
                for part in link_header.split(","):
                    if 'rel="next"' in part:
                        next_url = part.split(";")[0].strip().strip("<>")
                        break
            url = next_url
            params = None
        return submissions

    def _parse_selection(selection, total):
        selected = set()
        for part in selection.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                start, end = part.split("-", 1)
                if start.isdigit() and end.isdigit():
                    for i in range(int(start), int(end) + 1):
                        if 1 <= i <= total:
                            selected.add(i)
            elif part.isdigit():
                i = int(part)
                if 1 <= i <= total:
                    selected.add(i)
        return sorted(selected)

    def _prompt_grade(old_grade, force_keep_old=False):
        if keep_old_grade or force_keep_old:
            return old_grade, False
        raw = input(f"New grade (blank=keep {old_grade}, k=keep, s=skip, q=quit): ").strip()
        if not raw:
            return old_grade, False
        if raw.lower() in ("k", "keep"):
            return old_grade, False
        if raw.lower() in ("s", "skip"):
            return None, False
        if raw.lower() in ("q", "quit"):
            return None, True
        try:
            return float(raw), False
        except Exception:
            return None, False

    def _get_entry_score(entry):
        if not isinstance(entry, dict):
            return None
        for key in ("score", "entered_score"):
            value = entry.get(key)
            if value is not None:
                return value
        return None

    def _get_entry_time(entry, key):
        raw = entry.get(key) if isinstance(entry, dict) else None
        if not raw:
            return None
        try:
            return datetime.strptime(raw, "%Y-%m-%dT%H:%M:%SZ")
        except Exception:
            return None

    try:
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)
        if assignment_ids:
            assignment_ids = {str(aid).strip() for aid in assignment_ids if str(aid).strip()}
        else:
            assignments = []
            for assignment in course.get_assignments():
                if not getattr(assignment, "has_submitted_submissions", False):
                    continue
                name = getattr(assignment, "name", "") or ""
                if name.strip().lower() == "roll call attendance":
                    continue
                assignments.append({
                    "id": getattr(assignment, "id", None),
                    "name": name,
                    "due_at": getattr(assignment, "due_at", None)
                })
            if not assignments:
                print("No assignments found with resubmissions.")
                return []
            filtered = []
            for assignment in assignments:
                aid = assignment.get("id")
                if not aid:
                    continue
                try:
                    submissions = _fetch_assignment_submissions(aid)
                except Exception:
                    continue
                need_regrade = 0
                for sub in submissions:
                    history = sub.get("submission_history") if isinstance(sub, dict) else None
                    if not isinstance(history, list) or len(history) < 2:
                        continue
                    sorted_hist = sorted(history, key=lambda item: _get_entry_time(item, "submitted_at") or datetime.min)
                    last_graded = None
                    for h in sorted_hist:
                        score = _get_entry_score(h)
                        submitted_at = _get_entry_time(h, "submitted_at")
                        graded_at = _get_entry_time(h, "graded_at")
                        if score is not None and submitted_at and graded_at and graded_at > submitted_at:
                            last_graded = h
                    if not last_graded:
                        continue
                    latest = sorted_hist[-1]
                    latest_submitted = _get_entry_time(latest, "submitted_at")
                    latest_graded = _get_entry_time(latest, "graded_at")
                    latest_score = _get_entry_score(latest)
                    if latest_score is not None and latest_graded and latest_submitted and latest_graded > latest_submitted:
                        continue
                    last_attempt = last_graded.get("attempt")
                    latest_attempt = latest.get("attempt")
                    if latest_attempt is not None and last_attempt is not None:
                        try:
                            if int(latest_attempt) <= int(last_attempt):
                                continue
                        except Exception:
                            pass
                    need_regrade += 1
                if need_regrade > 0:
                    assignment["need_regrade"] = need_regrade
                    filtered.append(assignment)
            assignments = filtered
            if not assignments:
                print("No assignments found with resubmissions.")
                return []
            def _due_key(a):
                raw = a.get("due_at")
                if raw:
                    try:
                        return datetime.strptime(raw, "%Y-%m-%dT%H:%M:%SZ")
                    except Exception:
                        return datetime.max
                return datetime.max
            assignments.sort(key=_due_key)
            print("Assignments with resubmissions:")
            for idx, a in enumerate(assignments, 1):
                due = a.get("due_at") or "No due date"
                regrade_count = a.get("need_regrade")
                count_info = f", Need regrade: {regrade_count}" if regrade_count is not None else ""
                print(f"{idx}. {a['name']} (ID: {a['id']}, Due: {due}{count_info})")
            sel = input("Select assignments (e.g. 1,3-5 or 'a' for all, 'q' to quit): ").strip().lower()
            if sel in ("q", "quit"):
                return []
            if sel in ("a", "all", ""):
                assignment_ids = {str(a["id"]) for a in assignments}
            else:
                indices = _parse_selection(sel, len(assignments))
                assignment_ids = {str(assignments[i - 1]["id"]) for i in indices}

        updated = []
        resub_candidates = 0
        graded_candidates = 0
        apply_keep_old_all = keep_old_grade
        if not apply_keep_old_all:
            keep_all_raw = input("Use latest graded score for ungraded resubmissions for all assignments? (y/n) [n]: ").strip().lower()
            apply_keep_old_all = keep_all_raw in ("y", "yes")
        for aid in assignment_ids:
            try:
                assignment = course.get_assignment(aid)
            except Exception as e:
                if verbose:
                    print(f"[Resubmissions] Warning: could not load assignment {aid}: {e}")
                continue
            try:
                submissions = _fetch_assignment_submissions(aid)
            except Exception as e:
                if verbose:
                    print(f"[Resubmissions] Warning: failed to list submissions for {aid}: {e}")
                continue
            resub_needed = 0
            resub_work_items = []
            if verbose:
                history_sizes = [len(sub.get("submission_history") or []) for sub in submissions if isinstance(sub, dict)]
                with_history = sum(1 for size in history_sizes if size > 1)
                print(f"[Resubmissions] Assignment {aid}: {len(submissions)} submissions, {with_history} with history > 1.")
            for sub in submissions:
                history = sub.get("submission_history") if isinstance(sub, dict) else None
                if not isinstance(history, list) or len(history) < 2:
                    continue
                if verbose:
                    student_id = sub.get("user_id") if isinstance(sub, dict) else None
                    student_name = None
                    if isinstance(sub, dict):
                        student_name = sub.get("user", {}).get("name")
                    print(f"[Resubmissions] Submission history for {student_name or student_id}:")
                    for idx, entry in enumerate(history, 1):
                        submitted_at = entry.get("submitted_at")
                        graded_at = entry.get("graded_at")
                        score = _get_entry_score(entry)
                        attempt = entry.get("attempt")
                        print(f"  {idx}. attempt={attempt} submitted_at={submitted_at} graded_at={graded_at} score={score}")
                def _sort_key(item):
                    ts = _get_entry_time(item, "submitted_at")
                    return ts or datetime.min
                sorted_hist = sorted(history, key=_sort_key)
                last_graded = None
                for h in sorted_hist:
                    score = _get_entry_score(h)
                    submitted_at = _get_entry_time(h, "submitted_at")
                    graded_at = _get_entry_time(h, "graded_at")
                    if score is not None and submitted_at and graded_at and graded_at > submitted_at:
                        last_graded = h
                if not last_graded:
                    continue
                graded_candidates += 1
                graded_at = _get_entry_time(last_graded, "graded_at")
                if not graded_at:
                    continue
                resub_entry = None
                latest = sorted_hist[-1]
                latest_submitted = _get_entry_time(latest, "submitted_at")
                latest_graded = _get_entry_time(latest, "graded_at")
                if _get_entry_score(latest) is not None and latest_graded and latest_submitted and latest_graded > latest_submitted:
                    continue
                last_attempt = last_graded.get("attempt")
                latest_attempt = latest.get("attempt")
                if latest_attempt is not None and last_attempt is not None:
                    try:
                        if int(latest_attempt) <= int(last_attempt):
                            continue
                    except Exception:
                        pass
                for h in reversed(sorted_hist):
                    submitted_at = _get_entry_time(h, "submitted_at")
                    if submitted_at and submitted_at > graded_at:
                        resub_entry = h
                        break
                if not resub_entry:
                    continue
                resub_needed += 1
                resub_candidates += 1
                old_grade = _get_entry_score(last_graded)
                if old_grade is None:
                    continue
                resub_work_items.append((sub, old_grade, resub_entry))
            if verbose:
                print(f"[Resubmissions] Assignment {aid}: {resub_needed} need regrade after filtering.")
            if not resub_work_items:
                continue
            assignment_keep_old = apply_keep_old_all
            if not apply_keep_old_all:
                keep_raw = input("Use latest graded score for ungraded resubmissions in this assignment? (y/n) [n]: ").strip().lower()
                assignment_keep_old = keep_raw in ("y", "yes")
            for sub, old_grade, resub_entry in resub_work_items:
                student_id = sub.get("user_id") if isinstance(sub, dict) else None
                student_name = None
                if isinstance(sub, dict):
                    student_name = sub.get("user", {}).get("name")
                if not student_name:
                    user_info = getattr(sub, "user", None)
                    if isinstance(user_info, dict):
                        student_name = user_info.get("name")
                if verbose:
                    resub_time = _get_entry_time(resub_entry, "submitted_at")
                    resub_str = resub_time.strftime("%Y-%m-%d %H:%M:%S") if resub_time else "unknown"
                    print(f"[Resubmissions] {student_name or student_id} resubmitted after grade {old_grade} at {resub_str}.")
                new_grade, should_quit = _prompt_grade(old_grade, force_keep_old=assignment_keep_old)
                if should_quit:
                    return updated
                if new_grade is None:
                    continue
                try:
                    url = f"{api_url}/api/v1/courses/{course_id}/assignments/{aid}/submissions/{student_id}"
                    headers = {"Authorization": f"Bearer {api_key}"}
                    payload = {"submission": {"posted_grade": new_grade}}
                    resp = requests.put(url, headers=headers, json=payload)
                    if resp.ok:
                        updated.append({"assignment_id": aid, "student_id": student_id, "grade": new_grade})
                        if verbose:
                            print(f"[Resubmissions] Updated {student_name or student_id} to {new_grade}.")
                    else:
                        if verbose:
                            print(f"[Resubmissions] Failed to update {student_name or student_id}: {resp.status_code} {resp.text}")
                except Exception as e:
                    if verbose:
                        print(f"[Resubmissions] Failed to update {student_name or student_id}: {e}")
        if verbose:
            print(f"[Resubmissions] Candidates with graded attempts: {graded_candidates}.")
            print(f"[Resubmissions] Candidates with resubmissions after grade: {resub_candidates}.")
            print(f"[Resubmissions] Updated {len(updated)} resubmission grade(s).")
        else:
            print(f"Updated {len(updated)} resubmission grade(s).")
        return updated
    except Exception as e:
        if verbose:
            print(f"[Resubmissions] Error: {e}")
        else:
            print(f"Error applying resubmission grades: {e}")
        return []

def download_and_check_student_submissions(
    student_canvas_id=None,
    dest_dir=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    ocr_service=DEFAULT_OCR_METHOD,
    lang="auto",
    similarity_threshold=0.85,
    db_path=None,
    verbose=False
):
    """
    Download the latest submission file of a student for each assignment, extract text, check for similarity,
    and if two files are highly similar, send a warning message to the student.
    If student_canvas_id is None, list all students and allow user to choose a range of one or more students to check,
    allow user to select all students, allow user to quit, if no response after 60 seconds then use the default option of selecting all students.
    Only the latest submission (by submission time) for each assignment is considered.
    Downloaded file is named: <student name>_<canvas id>_<assignment id>_<submitted time>_<status>.<ext>
    Also saves the similarity check results in the database entry for the student if db_path is provided.
    Now saves downloaded file to a subfolder named <student name>_<student id> in the dest_dir (create if not exist).
    """

    def timeout_handler(signum, frame):
        print("\nTimeout: No response after 60 seconds. Using default: all students.")
        raise TimeoutError("User input timeout")

    def _prompt(prompt, timeout=60, default=None):
        return prompt_with_timeout(prompt, timeout=timeout, default=default, verbose=verbose)

    # Load student database for Canvas ID -> Name mapping if available
    canvasid_to_name = {}
    canvasid_to_sid = {}
    if db_path and os.path.exists(db_path):
        try:
            students_db = load_database(db_path)
            for s in students_db:
                canvas_id = str(getattr(s, "Canvas ID", "")).strip()
                name = getattr(s, "Name", "")
                sid = str(getattr(s, "Student ID", "")).strip()
                if canvas_id and name:
                    canvasid_to_name[canvas_id] = name
                if canvas_id and sid:
                    canvasid_to_sid[canvas_id] = sid
        except Exception:
            canvasid_to_name = {}
            canvasid_to_sid = {}

    canvas = get_canvas_client(api_url, api_key)
    course = canvas.get_course(course_id)

    # If student_canvas_id is None, list all students and allow user to select
    student_canvas_ids = []
    if student_canvas_id is None:
        people = list_canvas_people(api_url, api_key, course_id, verbose=verbose)
        students = people.get("active_students", [])
        if not students:
            print("No active students found.")
            return
        print("Active students:")
        for idx, s in enumerate(students, 1):
            print(f"{idx}. {s['name']} ({s['email']}), Canvas ID: {s['canvas_id']}")
        while True:
            sel = _prompt(
                "Enter student numbers to check (e.g. 1-5,7,9 or 'a' for all, or 'q' to quit, default 'a' in 60s): ",
                timeout=60,
                default="a"
            ).strip()
            if sel.lower() in ("q", "quit"):
                print("Quitting.")
                return
            if sel.lower() in ("a", "all"):
                selected = list(range(1, len(students) + 1))
            else:
                selected = set()
                for part in sel.split(","):
                    part = part.strip()
                    if "-" in part:
                        try:
                            start, end = map(int, part.split("-"))
                            selected.update(range(start, end + 1))
                        except Exception:
                            continue
                    elif part.isdigit():
                        selected.add(int(part))
                selected = [i for i in selected if 1 <= i <= len(students)]
            if not selected:
                print("No valid selection. Try again or 'q' to quit.")
                continue
            student_canvas_ids = [students[i - 1]["canvas_id"] for i in selected]
            break
    else:
        student_canvas_ids = [student_canvas_id] if isinstance(student_canvas_id, (str, int)) else list(student_canvas_id)

    if not dest_dir:
        dest_dir = os.path.join(os.getcwd(), "student_submissions")
    os.makedirs(dest_dir, exist_ok=True)

    for student_canvas_id in student_canvas_ids:
        assignments = list(course.get_assignments())
        pdf_files = []
        file_info = []

        # Get student name and student id from database if possible, else from Canvas
        student_name = None
        student_sid = None
        canvas_id_str = str(student_canvas_id)
        if canvasid_to_name.get(canvas_id_str):
            student_name = canvasid_to_name[canvas_id_str]
        else:
            try:
                user = canvas.get_user(student_canvas_id)
                student_name = getattr(user, "name", f"student_{student_canvas_id}")
            except Exception:
                student_name = f"student_{student_canvas_id}"
        if canvasid_to_sid.get(canvas_id_str):
            student_sid = canvasid_to_sid[canvas_id_str]
        else:
            student_sid = "unknown"

        # Create subfolder for this student
        safe_student_name = re.sub(r"[^\w\s-]", "", student_name).strip().replace(" ", "_")
        safe_student_sid = re.sub(r"[^\w\s-]", "", student_sid).strip()
        student_subfolder = f"{safe_student_name}_{safe_student_sid}"
        student_dir = os.path.join(dest_dir, student_subfolder)
        os.makedirs(student_dir, exist_ok=True)

        # Download only the latest PDF submission for each assignment
        for assignment in tqdm(assignments, desc=f"Checking assignments for {student_name}"):
            try:
                sub = assignment.get_submission(student_canvas_id)
                # Find the latest submission attempt by submitted_at
                latest_attachment = None
                latest_time = None
                latest_status = None
                # Check submission_history if available
                if hasattr(sub, "submission_history") and sub.submission_history:
                    for h in sub.submission_history:
                        submitted_at = h.get("submitted_at")
                        workflow_state = h.get("workflow_state", "unknown")
                        if submitted_at:
                            try:
                                sub_time = datetime.strptime(submitted_at, "%Y-%m-%dT%H:%M:%SZ")
                            except Exception:
                                continue
                            attachments = h.get("attachments", [])
                            pdf_attachments = [att for att in attachments if getattr(att, "filename", "").lower().endswith(".pdf")]
                            if pdf_attachments:
                                if latest_time is None or sub_time > latest_time:
                                    latest_time = sub_time
                                    latest_attachment = pdf_attachments[0]
                                    latest_status = workflow_state
                # Fallback: check attachments directly on submission object
                elif hasattr(sub, "attachments") and sub.attachments:
                    attachments = sub.attachments
                    pdf_attachments = [att for att in attachments if getattr(att, "filename", "").lower().endswith(".pdf")]
                    if pdf_attachments:
                        submitted_at = getattr(sub, "submitted_at", None)
                        workflow_state = getattr(sub, "workflow_state", "unknown")
                        if submitted_at:
                            try:
                                sub_time = datetime.strptime(submitted_at, "%Y-%m-%dT%H:%M:%SZ")
                            except Exception:
                                sub_time = None
                        else:
                            sub_time = None
                        if latest_time is None or (sub_time and (latest_time is None or sub_time > latest_time)):
                            latest_time = sub_time
                            latest_attachment = pdf_attachments[0]
                            latest_status = workflow_state
                if not latest_attachment:
                    continue
                att = latest_attachment
                url = getattr(att, "url", None)
                orig_filename = getattr(att, "filename", None) or f"{assignment.id}_{student_canvas_id}.pdf"
                # Format: <student name>_<canvas id>_<assignment id>_<submitted time>_<status>.<ext>
                ext = os.path.splitext(orig_filename)[1]
                submitted_time_str = latest_time.strftime("%Y%m%d_%H%M") if latest_time else "unknown"
                status_str = latest_status or "unknown"
                out_filename = f"{safe_student_name}_{student_canvas_id}_{assignment.id}_{submitted_time_str}_{status_str}{ext}"
                out_path = os.path.join(student_dir, out_filename)
                if not os.path.exists(out_path):
                    r = requests.get(url)
                    with open(out_path, "wb") as f:
                        f.write(r.content)
                pdf_files.append(out_path)
                file_info.append({
                    "assignment_id": assignment.id,
                    "assignment_name": assignment.name,
                    "file_path": out_path
                })
            except Exception as e:
                if verbose:
                    print(f"[StudentSubmissions] Error downloading for assignment {assignment.id}: {e}")

        if not pdf_files:
            if verbose:
                print(f"[StudentSubmissions] No PDF submissions found for student {student_canvas_id}.")
            else:
                print("No PDF submissions found for this student.")
            continue

        # Extract text from all PDFs
        extracted_texts = {}
        for info in tqdm(file_info, desc=f"Extracting text for {student_name}"):
            pdf_path = info["file_path"]
            txt_path = pdf_path + f"_text_{ocr_service}.txt"
            if not os.path.exists(txt_path):
                txt_path = extract_text_from_scanned_pdf(
                    pdf_path,
                    txt_output_path=txt_path,
                    service=ocr_service,
                    lang=lang,
                    simple_text=True,
                    verbose=verbose
                )
            if txt_path and os.path.exists(txt_path):
                with open(txt_path, "r", encoding="utf-8") as f:
                    text = f.read()
                norm_text = re.sub(r"\s+", " ", text).strip().lower()
                extracted_texts[pdf_path] = norm_text
            else:
                extracted_texts[pdf_path] = ""

        # Compare all pairs for similarity
        pdf_list = list(extracted_texts.keys())
        texts = [extracted_texts[p] for p in pdf_list]
        if len(pdf_list) < 2:
            if verbose:
                print(f"[StudentSubmissions] Less than 2 submissions to compare for student {student_canvas_id}.")
            else:
                print("Less than 2 submissions to compare.")
            continue

        tfidf_vectorizer = TfidfVectorizer().fit(texts)
        tfidf_matrix = tfidf_vectorizer.transform(texts)
        similar_pairs = []
        for i, pdf1 in enumerate(pdf_list):
            for j in range(i + 1, len(pdf_list)):
                pdf2 = pdf_list[j]
                text1 = extracted_texts[pdf1]
                text2 = extracted_texts[pdf2]
                if not text1 or not text2:
                    continue
                cos_sim = cosine_similarity(tfidf_matrix[i], tfidf_matrix[j])[0, 0]
                seq_sim = difflib.SequenceMatcher(None, text1, text2).ratio()
                ratio = 0.7 * cos_sim + 0.3 * seq_sim
                if ratio >= similarity_threshold:
                    similar_pairs.append((pdf1, pdf2, ratio))

        # Only save similarity results to database if high similarity found
        if similar_pairs and db_path and os.path.exists(db_path):
            try:
                students = load_database(db_path)
                # Find the student entry by Canvas ID
                for s in students:
                    if str(getattr(s, "Canvas ID", "")) == str(student_canvas_id):
                        # Save the similarity results as a field
                        s.__dict__["Submission Similarity Results"] = [
                            {
                                "file1": os.path.basename(pdf1),
                                "file2": os.path.basename(pdf2),
                                "similarity": ratio
                            }
                            for pdf1, pdf2, ratio in similar_pairs
                        ]
                        save_database(students, db_path)
                        if verbose:
                            print(f"[StudentSubmissions] Saved similarity results to database for student {student_canvas_id}.")
                        break
            except Exception as e:
                if verbose:
                    print(f"[StudentSubmissions] Failed to save similarity results to database: {e}")

        if similar_pairs:
            # Compose warning message
            assignment_names = []
            for pdf1, pdf2, _ in similar_pairs:
                a1 = next((f["assignment_name"] for f in file_info if f["file_path"] == pdf1), pdf1)
                a2 = next((f["assignment_name"] for f in file_info if f["file_path"] == pdf2), pdf2)
                assignment_names.append((a1, a2))
            assignment_names_str = "\n".join([f"- {a1} <-> {a2}" for a1, a2 in assignment_names])
            message = (
                "Hệ thống phát hiện bạn đã nộp các file có nội dung rất giống nhau cho nhiều bài tập khác nhau:\n"
                f"{assignment_names_str}\n\n"
                "Việc nộp các file có nội dung rất giống nhau cho nhiều bài tập khác nhau bị coi là gian lận và vi phạm quy định của lớp học. "
                "Bạn cần nộp lại từng bài tập với nội dung phù hợp cho từng bài càng sớm càng tốt. "
                "Nếu có thắc mắc, hãy liên hệ với giảng viên. \n"
                "Thông báo này được gửi tự động từ hệ thống."
            )

            # Print message and ask user if want to send/refine/quit
            print("\n--- Prepared warning message to send ---")
            print(f"Subject: {subject}")
            print(f"Message:\n{message}\n")
            print("Options: [y] Send as is  [r] Refine with AI  [q] Quit/skip")
            try:
                choice = _prompt(
                    "Send this message? (y/r/q, default y in 60s): ",
                    timeout=60,
                    default="y"
                ).strip().lower()
            except TimeoutError:
                choice = "y"
            except KeyboardInterrupt:
                print("Operation cancelled by user.")
                return

            if choice in ("q", "quit"):
                if verbose:
                    print("[StudentSubmissions] Skipped sending warning message.")
                else:
                    print("Skipped sending warning message.")
                continue
            if choice == "r":
                # Refine with AI
                if refine in ALL_AI_METHODS:
                    prompt = (
                        "Bạn là trợ lý giáo dục chuyên nghiệp. Hãy viết lại thông báo sau bằng tiếng Việt, lịch sự, rõ ràng, "
                        "giải thích rằng việc nộp các file có nội dung rất giống nhau cho nhiều bài tập khác nhau bị coi là gian lận, "
                        "và nhắc sinh viên nộp lại từng bài tập với nội dung phù hợp. Đưa vào danh sách các bài bị phát hiện. "
                        "Chỉ trả về thông báo đã chỉnh sửa, không giải thích gì thêm.\n\n"
                        "Thông báo:\n{text}"
                    )
                    message = refine_text_with_ai(message, method=refine, user_prompt=prompt)
                    print("\n--- Refined message ---\n")
                    print(message)
                    try:
                        confirm = _prompt(
                            "Send this refined message? (y/n, default y in 60s): ",
                            timeout=60,
                            default="y"
                        ).strip().lower()
                    except TimeoutError:
                        confirm = "y"
                    except KeyboardInterrupt:
                        print("Operation cancelled by user.")
                        return
                    if confirm not in ("y", "yes", ""):
                        if verbose:
                            print("[StudentSubmissions] Skipped sending warning message.")
                        else:
                            print("Skipped sending warning message.")
                        continue
                else:
                    print("No AI method available for refinement. Sending original message.")

            # Send message
            try:
                canvas.create_conversation(
                    recipients=[str(student_canvas_id)],
                    subject=subject,
                    body=message,
                    force_new=True
                )
                if verbose:
                    print(f"[StudentSubmissions] Warning message sent to student {student_canvas_id}.")
                else:
                    print("Warning message sent to student.")
            except Exception as e:
                if verbose:
                    print(f"[StudentSubmissions] Failed to send warning message: {e}")
                else:
                    print(f"Failed to send warning message: {e}")
        else:
            if verbose:
                print(f"[StudentSubmissions] No highly similar submissions detected for student {student_canvas_id}.")
            else:
                print("No highly similar submissions detected.")

