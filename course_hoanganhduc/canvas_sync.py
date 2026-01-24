# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas roster sync helpers."""

import re
import traceback

import requests
from tqdm import tqdm

from .canvas_auth import get_canvas_client, require_canvas_config
from .canvas_people import list_canvas_people
from .data import save_database
from .models import Student
from .settings import (
    CANVAS_LMS_API_KEY,
    CANVAS_LMS_API_URL,
    CANVAS_LMS_COURSE_ID,
)

def sync_students_with_canvas(students, db_path=None, course_id=None, api_url=CANVAS_LMS_API_URL, api_key=CANVAS_LMS_API_KEY, verbose=False):
    """
    Sync students in the local database with active students from Canvas course.
    Adds new students from Canvas if not present, updates Canvas ID for existing students,
    and syncs scores of both total course grade and each assignment category.

    Args:
        students: List of Student objects
        db_path: Path to save the database
        course_id: Canvas course ID (uses default if None)
        api_url: Canvas API URL
        api_key: Canvas API key
        verbose: If True, print more details; otherwise, print only important notice

    Returns:
        (added_count, updated_count): Counts of students added and updated
    """
    try:
        def _scale_score_to_ten(score, max_points=None):
            try:
                if score is None:
                    return None
                score_val = float(score)
            except Exception:
                return score
            if max_points not in (None, 0):
                try:
                    max_val = float(max_points)
                    if max_val > 0:
                        return round(score_val / max_val * 10, 2)
                except Exception:
                    pass
            if score_val > 10:
                return round(score_val / 10, 2)
            return score_val

        if course_id is None:
            course_id = CANVAS_LMS_COURSE_ID
        api_url, api_key, course_id = require_canvas_config(api_url, api_key, course_id)
        if verbose:
            print(f"[SyncCanvas] Fetching students from Canvas course {course_id}...")
        else:
            print("Syncing students with Canvas course...")
        people = list_canvas_people(api_url=api_url, api_key=api_key, course_id=course_id)
        canvas_students = people.get("active_students", [])
        if not canvas_students:
            if verbose:
                print("[SyncCanvas] No active students found in Canvas course.")
            else:
                print("No active students found in Canvas course.")
            return 0, 0

        if verbose:
            print(f"[SyncCanvas] Found {len(canvas_students)} active students in Canvas course.")

        # Helper to normalize names for comparison
        def normalize_name(name):
            if not name:
                return ""
            name = str(name)
            name = re.sub(r"[^a-zA-Z0-9 ]", "", name)
            name = re.sub(r"\s+", " ", name)
            return name.strip().lower()

        # Build lookups for matching Canvas records to local students.
        existing_by_email = {}
        existing_by_name = {}
        existing_by_canvas_id = {}

        if verbose:
            print("[SyncCanvas] Building lookup tables for existing students...")
        for s in students:
            email = getattr(s, "Email", None)
            name = getattr(s, "Name", None)
            canvas_id = getattr(s, "Canvas ID", None)

            if email:
                existing_by_email[email.lower()] = s
            if name:
                norm_name = normalize_name(name)
                if norm_name:
                    existing_by_name[norm_name] = s
            if canvas_id:
                try:
                    existing_by_canvas_id[int(canvas_id)] = s
                except (ValueError, TypeError):
                    pass

        # Resolve duplicates by priority (Canvas ID > name > email). If multiple candidates
        # exist, prompt the operator to pick, create a new student, or skip.
        def _resolve_canvas_match(canvas_id_value, name_key, email_key):
            candidates = []
            seen = set()

            def add_candidate(label, student):
                if id(student) in seen:
                    return
                candidates.append((label, student))
                seen.add(id(student))

            if canvas_id_value:
                try:
                    cid = int(canvas_id_value)
                    if cid in existing_by_canvas_id:
                        add_candidate("canvas_id", existing_by_canvas_id[cid])
                except (ValueError, TypeError):
                    pass
            if name_key and name_key in existing_by_name:
                add_candidate("name", existing_by_name[name_key])
            if email_key and email_key in existing_by_email:
                add_candidate("email", existing_by_email[email_key])

            if not candidates:
                return None
            if len(candidates) == 1:
                return candidates[0][1]

            print("\n[SyncCanvas] Possible duplicate match detected:")
            for idx, (label, student) in enumerate(candidates, 1):
                s_name = getattr(student, "Name", "") or ""
                s_email = getattr(student, "Email", "") or ""
                s_cid = getattr(student, "Canvas ID", "") or ""
                print(f"{idx}. {s_name} | {s_email} | Canvas ID: {s_cid} (matched by {label})")
            print("n. Create new student")
            print("s. Skip this record")
            while True:
                choice = input("Choose a match (number), 'n' for new, or 's' to skip: ").strip().lower()
                if choice == "n":
                    return None
                if choice == "s":
                    return "__skip__"
                if choice.isdigit():
                    sel = int(choice) - 1
                    if 0 <= sel < len(candidates):
                        return candidates[sel][1]

        added_count = 0
        updated_count = 0

        # Prepare Canvas API for grades and scores
        if verbose:
            print("[SyncCanvas] Connecting to Canvas API...")
        canvas = get_canvas_client(api_url, api_key, verbose=verbose)
        course = canvas.get_course(course_id)

        # Fetch enrollments to get total grades
        if verbose:
            print("[SyncCanvas] Fetching enrollments for total grades...")
        enrollments = list(course.get_enrollments(type=['StudentEnrollment']))

        # Build a map: canvas_id -> total_grade info.
        total_grades_by_canvas_id = {}
        for enrollment in enrollments:
            user_id = getattr(enrollment, "user_id", None)
            if user_id and hasattr(enrollment, "grades"):
                grades = enrollment.grades
                total_grades_by_canvas_id[user_id] = {
                    "final_score": grades.get("final_score"),
                    "final_grade": grades.get("final_grade"),
                    "current_score": grades.get("current_score"),
                    "current_grade": grades.get("current_grade")
                }

        # Fetch assignment groups for category scores
        if verbose:
            print("[SyncCanvas] Fetching assignment groups and scores...")
        try:
            headers = {'Authorization': f'Bearer {api_key}'}
            group_scores_url = f"{api_url}/api/v1/courses/{course_id}/students/submissions?student_ids[]=all&include[]=assignment&include[]=user&include[]=score&grouped=true"
            response = requests.get(group_scores_url, headers=headers)
            response.raise_for_status()
            category_scores_data = response.json()
        except Exception as e:
            if verbose:
                print(f"[SyncCanvas] Warning: Could not fetch detailed category scores: {e}")
            else:
                print("Warning: Could not fetch detailed category scores.")
            category_scores_data = []

        assignment_groups = list(course.get_assignment_groups(include=['assignments', 'score_statistics']))

        # Build a map: canvas_id -> {group_name: {current_score, current_possible, final_score, final_possible}}.
        group_scores_by_canvas_id = {}

        # First, process score_statistics (current scores)
        for group in assignment_groups:
            group_name = group.name
            group_id = group.id
            stats = getattr(group, "score_statistics", {})
            for canvas_id_str, stat in stats.items():
                try:
                    canvas_id = int(canvas_id_str)
                except Exception:
                    continue
                if canvas_id not in group_scores_by_canvas_id:
                    group_scores_by_canvas_id[canvas_id] = {}
                if group_name not in group_scores_by_canvas_id[canvas_id]:
                    group_scores_by_canvas_id[canvas_id][group_name] = {
                        'current_score': 0,
                        'current_possible': 0,
                        'final_score': 0,
                        'final_possible': 0,
                        'group_id': group_id
                    }

                group_scores_by_canvas_id[canvas_id][group_name]['current_score'] = stat.get("score", 0)
                group_scores_by_canvas_id[canvas_id][group_name]['current_possible'] = stat.get("possible", 0)
                if group_scores_by_canvas_id[canvas_id][group_name]['final_score'] == 0:
                    group_scores_by_canvas_id[canvas_id][group_name]['final_score'] = stat.get("score", 0)
                if group_scores_by_canvas_id[canvas_id][group_name]['final_possible'] == 0:
                    group_scores_by_canvas_id[canvas_id][group_name]['final_possible'] = stat.get("possible", 0)

        # Then process submissions data to get final scores aggregated per assignment group.
        if verbose:
            print("[SyncCanvas] Processing category final scores...")
        try:
            assignments_by_group = {}
            for group in assignment_groups:
                assignments_by_group[group.id] = []
                for assignment in getattr(group, "assignments", []):
                    if assignment.get("published", False):
                        assignments_by_group[group.id].append(assignment['id'])

            for user_data in tqdm(category_scores_data, desc="Processing category final scores"):
                user_id = user_data.get("user_id")
                if not user_id:
                    continue

                if user_id not in group_scores_by_canvas_id:
                    group_scores_by_canvas_id[user_id] = {}

                for submission in user_data.get("submissions", []):
                    assignment = submission.get("assignment")
                    if not assignment:
                        continue

                    group_id = assignment.get("assignment_group_id")
                    group = next((g for g in assignment_groups if g.id == group_id), None)
                    if not group:
                        continue

                    group_name = group.name
                    if group_name not in group_scores_by_canvas_id[user_id]:
                        group_scores_by_canvas_id[user_id][group_name] = {
                            'current_score': 0,
                            'current_possible': 0,
                            'final_score': 0,
                            'final_possible': 0,
                            'group_id': group_id
                        }

                    score = submission.get("score", 0) or 0
                    points_possible = assignment.get("points_possible", 0) or 0

                    group_scores_by_canvas_id[user_id][group_name]['final_score'] += score
                    group_scores_by_canvas_id[user_id][group_name]['final_possible'] += points_possible
        except Exception as e:
            if verbose:
                print(f"[SyncCanvas] Warning: Error processing category final scores: {e}")
            else:
                print("Warning: Error processing category final scores.")

        if verbose:
            print("[SyncCanvas] Fetching individual assignments...")
        important_assignments = {}
        try:
            for group in assignment_groups:
                for assignment in getattr(group, "assignments", []):
                    if assignment.get("published", False):
                        important_assignments[assignment['id']] = {
                            "name": assignment['name'],
                            "points_possible": assignment['points_possible'],
                            "group_name": group.name
                        }
        except Exception as e:
            if verbose:
                print(f"[SyncCanvas] Warning: Could not fetch assignments: {e}")
            else:
                print("Warning: Could not fetch assignments.")

        assignment_scores_by_canvas_id = {}
        assignment_comments_by_canvas_id = {}
        assignment_rubrics_by_canvas_id = {}
        assignment_status_by_canvas_id = {}
        if important_assignments:
            if verbose:
                print(f"[SyncCanvas] Fetching submissions for {len(important_assignments)} assignments...")
            try:
                for assignment_id in list(important_assignments.keys()):
                    assignment = course.get_assignment(assignment_id)
                    assignment_info = important_assignments[assignment_id]

                    for submission in tqdm(assignment.get_submissions(include=["user", "submission_comments", "rubric_assessment"]),
                                         desc=f"Processing {assignment_info['name']} submissions"):
                        user_id = getattr(submission, "user_id", None)
                        score = getattr(submission, "score", None)
                        workflow_state = getattr(submission, "workflow_state", None)
                        if user_id and score is not None:
                            if user_id not in assignment_scores_by_canvas_id:
                                assignment_scores_by_canvas_id[user_id] = {}
                            assignment_scores_by_canvas_id[user_id][assignment_id] = {
                                "name": assignment_info["name"],
                                "group": assignment_info["group_name"],
                                "score": score,
                                "points_possible": assignment_info["points_possible"]
                            }
                        if user_id and workflow_state:
                            if user_id not in assignment_status_by_canvas_id:
                                assignment_status_by_canvas_id[user_id] = {}
                            assignment_status_by_canvas_id[user_id][assignment_id] = {
                                "name": assignment_info["name"],
                                "status": workflow_state,
                            }
                        if user_id:
                            comments = getattr(submission, "submission_comments", None)
                            if comments:
                                if user_id not in assignment_comments_by_canvas_id:
                                    assignment_comments_by_canvas_id[user_id] = {}
                                normalized_comments = []
                                for comment in comments:
                                    if isinstance(comment, dict):
                                        normalized_comments.append({
                                            "author_id": comment.get("author_id"),
                                            "author_name": comment.get("author_name"),
                                            "comment": comment.get("comment"),
                                            "posted_at": comment.get("posted_at"),
                                        })
                                assignment_comments_by_canvas_id[user_id][assignment_id] = normalized_comments
                            rubric = getattr(submission, "rubric_assessment", None)
                            if rubric:
                                if user_id not in assignment_rubrics_by_canvas_id:
                                    assignment_rubrics_by_canvas_id[user_id] = {}
                                assignment_rubrics_by_canvas_id[user_id][assignment_id] = rubric
            except Exception as e:
                if verbose:
                    print(f"[SyncCanvas] Warning: Error fetching assignment submissions: {e}")
                else:
                    print("Warning: Error fetching assignment submissions.")

        # Merge Canvas roster and grades into the local database.
        if verbose:
            print("[SyncCanvas] Syncing student data...")
        else:
            print("Syncing student data...")
        for canvas_student in tqdm(canvas_students, desc="Syncing students"):
            canvas_email = canvas_student.get("email", "").lower()
            canvas_name = canvas_student.get("name", "")
            canvas_id = canvas_student.get("canvas_id", "")

            norm_canvas_name = normalize_name(canvas_name)
            matched_student = _resolve_canvas_match(canvas_id, norm_canvas_name, canvas_email)
            if matched_student == "__skip__":
                continue

            if matched_student:
                changed = False

                if not hasattr(matched_student, "Canvas ID") or not getattr(matched_student, "Canvas ID"):
                    setattr(matched_student, "Canvas ID", canvas_id)
                    changed = True

                if canvas_email and (not hasattr(matched_student, "Email") or not getattr(matched_student, "Email")):
                    setattr(matched_student, "Email", canvas_email)
                    changed = True

                if canvas_id and total_grades_by_canvas_id.get(int(canvas_id)):
                    grade_data = total_grades_by_canvas_id[int(canvas_id)]
                    for grade_field, grade_value in grade_data.items():
                        if grade_value is not None:
                            field_name = {
                                "final_score": "Total Final Score",
                                "final_grade": "Total Final Grade",
                            }.get(grade_field)
                            if field_name:
                                if grade_field == "final_score":
                                    grade_value = _scale_score_to_ten(grade_value, 100)
                                if getattr(matched_student, field_name, None) != grade_value:
                                    setattr(matched_student, field_name, grade_value)
                                    changed = True

                if canvas_id and group_scores_by_canvas_id.get(int(canvas_id)):
                    for group_name, scores_data in group_scores_by_canvas_id[int(canvas_id)].items():
                        final_score = _scale_score_to_ten(
                            scores_data.get('final_score', 0),
                            scores_data.get('final_possible', 0),
                        )
                        final_field = f"{group_name} Final Score"
                        if getattr(matched_student, final_field, None) != final_score:
                            setattr(matched_student, final_field, final_score)
                            changed = True

                if canvas_id and assignment_scores_by_canvas_id.get(int(canvas_id)):
                    for assignment_id, assignment_data in assignment_scores_by_canvas_id[int(canvas_id)].items():
                        name = assignment_data["name"]
                        score = _scale_score_to_ten(
                            assignment_data["score"],
                            assignment_data.get("points_possible", 0),
                        )
                        field = f"Assignment: {name}"
                        if getattr(matched_student, field, None) != score:
                            setattr(matched_student, field, score)
                            changed = True

                if canvas_id and assignment_comments_by_canvas_id.get(int(canvas_id)):
                    existing_comments = getattr(matched_student, "Canvas Submission Comments", None) or {}
                    updated_comments = dict(existing_comments)
                    for assignment_id, comments in assignment_comments_by_canvas_id[int(canvas_id)].items():
                        name = important_assignments.get(assignment_id, {}).get("name", str(assignment_id))
                        updated_comments[name] = comments
                    if updated_comments != existing_comments:
                        setattr(matched_student, "Canvas Submission Comments", updated_comments)
                        changed = True

                if canvas_id and assignment_rubrics_by_canvas_id.get(int(canvas_id)):
                    existing_rubrics = getattr(matched_student, "Canvas Rubric Evaluations", None) or {}
                    updated_rubrics = dict(existing_rubrics)
                    for assignment_id, rubric in assignment_rubrics_by_canvas_id[int(canvas_id)].items():
                        name = important_assignments.get(assignment_id, {}).get("name", str(assignment_id))
                        updated_rubrics[name] = rubric
                    if updated_rubrics != existing_rubrics:
                        setattr(matched_student, "Canvas Rubric Evaluations", updated_rubrics)
                        changed = True
                if canvas_id and assignment_status_by_canvas_id.get(int(canvas_id)):
                    existing_statuses = getattr(matched_student, "Canvas Submissions", None) or {}
                    updated_statuses = dict(existing_statuses)
                    for assignment_id, status_info in assignment_status_by_canvas_id[int(canvas_id)].items():
                        name = important_assignments.get(assignment_id, {}).get("name", str(assignment_id))
                        status = status_info.get("status") if isinstance(status_info, dict) else status_info
                        if status:
                            updated_statuses[name] = status
                    if updated_statuses != existing_statuses:
                        setattr(matched_student, "Canvas Submissions", updated_statuses)
                        changed = True

                if changed:
                    updated_count += 1
            else:
                new_student_data = {
                    "Name": canvas_name,
                    "Email": canvas_email,
                    "Canvas ID": canvas_id
                }

                if canvas_id and total_grades_by_canvas_id.get(int(canvas_id)):
                    grade_data = total_grades_by_canvas_id[int(canvas_id)]
                    for grade_field, grade_value in grade_data.items():
                        if grade_value is not None:
                            field_name = {
                                "final_score": "Total Final Score",
                                "final_grade": "Total Final Grade",
                            }.get(grade_field)
                            if field_name:
                                if grade_field == "final_score":
                                    grade_value = _scale_score_to_ten(grade_value, 100)
                                new_student_data[field_name] = grade_value

                if canvas_id and group_scores_by_canvas_id.get(int(canvas_id)):
                    for group_name, scores_data in group_scores_by_canvas_id[int(canvas_id)].items():
                        new_student_data[f"{group_name} Final Score"] = _scale_score_to_ten(
                            scores_data.get('final_score', 0),
                            scores_data.get('final_possible', 0),
                        )

                if canvas_id and assignment_scores_by_canvas_id.get(int(canvas_id)):
                    for assignment_id, assignment_data in assignment_scores_by_canvas_id[int(canvas_id)].items():
                        name = assignment_data["name"]
                        score = _scale_score_to_ten(
                            assignment_data["score"],
                            assignment_data.get("points_possible", 0),
                        )
                        field = f"Assignment: {name}"
                        new_student_data[field] = score

                if canvas_id and assignment_comments_by_canvas_id.get(int(canvas_id)):
                    comments = {}
                    for assignment_id, items in assignment_comments_by_canvas_id[int(canvas_id)].items():
                        name = important_assignments.get(assignment_id, {}).get("name", str(assignment_id))
                        comments[name] = items
                    new_student_data["Canvas Submission Comments"] = comments

                if canvas_id and assignment_rubrics_by_canvas_id.get(int(canvas_id)):
                    rubrics = {}
                    for assignment_id, rubric in assignment_rubrics_by_canvas_id[int(canvas_id)].items():
                        name = important_assignments.get(assignment_id, {}).get("name", str(assignment_id))
                        rubrics[name] = rubric
                    new_student_data["Canvas Rubric Evaluations"] = rubrics
                if canvas_id and assignment_status_by_canvas_id.get(int(canvas_id)):
                    statuses = {}
                    for assignment_id, status_info in assignment_status_by_canvas_id[int(canvas_id)].items():
                        name = important_assignments.get(assignment_id, {}).get("name", str(assignment_id))
                        status = status_info.get("status") if isinstance(status_info, dict) else status_info
                        if status:
                            statuses[name] = status
                    if statuses:
                        new_student_data["Canvas Submissions"] = statuses

                students.append(Student(**new_student_data))
                if canvas_id:
                    try:
                        existing_by_canvas_id[int(canvas_id)] = students[-1]
                    except (ValueError, TypeError):
                        pass
                if canvas_email:
                    existing_by_email[canvas_email] = students[-1]
                if norm_canvas_name:
                    existing_by_name[norm_canvas_name] = students[-1]
                added_count += 1

        if added_count > 0 or updated_count > 0:
            if db_path:
                if verbose:
                    print(f"[SyncCanvas] Saving updated database with {added_count} new and {updated_count} modified students...")
                else:
                    print(f"Saving updated database with {added_count} new and {updated_count} modified students...")
                save_database(students, db_path)
                if verbose:
                    print("[SyncCanvas] Database saved successfully.")
                else:
                    print("Database saved successfully.")

        if verbose:
            print(f"[SyncCanvas] Sync completed: {added_count} students added, {updated_count} students updated.")
        else:
            print(f"Sync completed: {added_count} students added, {updated_count} students updated.")

        return added_count, updated_count

    except Exception as e:
        if verbose:
            print(f"[SyncCanvas] Error syncing with Canvas: {e}")
            traceback.print_exc()
        else:
            print(f"Error syncing with Canvas: {e}")
        return 0, 0

