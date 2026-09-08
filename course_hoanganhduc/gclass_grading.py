# -*- coding: utf-8 -*-

from googleapiclient.discovery import build
from .gclass_auth import _get_google_classroom_credentials, list_google_classroom_courses
from .utils import get_input_with_quit, parse_selection

def grade_google_classroom_assignment_submissions(
    course_id=None,
    credentials_path='gclassroom_credentials.json',
    token_path='token.pickle',
    coursework_ids=None,
    score=None,
    ungraded_only=True,
    apply_all=False,
    verbose=False,
):
    """
    Grade Google Classroom assignment submissions interactively or with provided IDs.
    - If coursework_ids is None, prompt for selection.
    - If score is None, prompt for a score and apply to selected submissions.
    - If ungraded_only is True, only list ungraded submissions.
    """
    def scale_score(score_value, max_points):
        try:
            score_val = float(score_value)
        except Exception:
            return None
        if max_points in (None, 0):
            return score_val
        try:
            max_val = float(max_points)
        except Exception:
            return score_val
        if score_val <= 10 and max_val > 10:
            return round(score_val / 10 * max_val, 2)
        return score_val

    creds = _get_google_classroom_credentials(credentials_path, token_path, verbose=verbose)
    service = build("classroom", "v1", credentials=creds)

    if not course_id:
        courses = list_google_classroom_courses(credentials_path, token_path, verbose=verbose)
        if not courses:
            print("No courses found.")
            return 0
        print("Available Google Classroom courses:")
        for i, c in enumerate(courses, 1):
            print(f"{i}. {c.get('name')} (ID: {c.get('id')})")
        while True:
            sel = input("Select course number: ").strip()
            if not sel:
                continue
            try:
                idx = int(sel) - 1
                if 0 <= idx < len(courses):
                    course_id = courses[idx].get("id")
                    break
            except Exception:
                continue
    if not course_id:
        print("No course selected.")
        return 0

    coursework = []
    next_token = None
    while True:
        req = service.courses().courseWork().list(courseId=course_id, pageToken=next_token, pageSize=200) if next_token else service.courses().courseWork().list(courseId=course_id, pageSize=200)
        resp = req.execute()
        coursework.extend(resp.get("courseWork", []) or [])
        next_token = resp.get("nextPageToken")
        if not next_token:
            break

    if not coursework:
        print("No assignments found in this course.")
        return 0

    selected_coursework = []
    if coursework_ids:
        ids = [str(cid).strip() for cid in (coursework_ids or []) if str(cid).strip()]
        for cw in coursework:
            if str(cw.get("id")) in ids:
                selected_coursework.append(cw)
        if not selected_coursework:
            print("No matching assignments found for provided coursework IDs.")
            return 0
    else:
        print("Assignments:")
        for i, cw in enumerate(coursework, 1):
            title = cw.get("title", "")
            max_points = cw.get("maxPoints")
            print(f"{i}. {title} (ID: {cw.get('id')}, max: {max_points})")
        while True:
            sel = get_input_with_quit("Select assignment numbers (e.g. 1,3-5, 'a' for all, or 'q' to quit): ")
            if sel is None:
                return 0
            indices = parse_selection(sel, len(coursework))
            if indices:
                selected_coursework = [coursework[i - 1] for i in indices]
                break
            print("Invalid selection. Please enter valid numbers, a range, 'a' for all, or 'q' to quit.")

    students_map = {}
    try:
        next_token = None
        while True:
            req = service.courses().students().list(courseId=course_id, pageToken=next_token, pageSize=200) if next_token else service.courses().students().list(courseId=course_id, pageSize=200)
            resp = req.execute()
            for entry in resp.get("students", []) or []:
                profile = entry.get("profile", {}) or {}
                user_id = entry.get("userId")
                full_name = (profile.get("name", {}) or {}).get("fullName") or ""
                email = profile.get("emailAddress") or ""
                if user_id:
                    students_map[str(user_id)] = {"name": full_name, "email": email}
            next_token = resp.get("nextPageToken")
            if not next_token:
                break
    except Exception:
        if verbose:
            print("[GClassroom] Warning: could not fetch roster; submissions will show user IDs only.")

    graded_count = 0
    for cw in selected_coursework:
        cw_id = cw.get("id")
        if not cw_id:
            continue
        title = cw.get("title", f"cw_{cw_id}")
        max_points = cw.get("maxPoints")
        submissions = []
        next_token = None
        while True:
            req = service.courses().courseWork().studentSubmissions().list(courseId=course_id, courseWorkId=cw_id, pageToken=next_token, pageSize=200) if next_token else service.courses().courseWork().studentSubmissions().list(courseId=course_id, courseWorkId=cw_id, pageSize=200)
            resp = req.execute()
            submissions.extend(resp.get("studentSubmissions", []) or [])
            next_token = resp.get("nextPageToken")
            if not next_token:
                break
        if not submissions:
            print(f"No submissions found for assignment '{title}'.")
            continue

        students_to_grade = []
        for sub in submissions:
            state = sub.get("state")
            assigned = sub.get("assignedGrade")
            user_id = str(sub.get("userId") or "")
            if state == "NEW":
                continue
            if ungraded_only and assigned is not None:
                continue
            student_info = students_map.get(user_id, {})
            students_to_grade.append({
                "submission": sub,
                "user_id": user_id,
                "name": student_info.get("name") or user_id,
                "email": student_info.get("email") or "",
                "state": state,
                "assigned": assigned,
            })

        if not students_to_grade:
            msg = "No ungraded submissions found." if ungraded_only else "No submissions available for grading."
            print(f"{msg} ({title})")
            continue

        students_to_grade.sort(key=lambda s: s["name"])
        print(f"\nAssignment: {title} (ID: {cw_id}, max: {max_points})")
        for idx, s in enumerate(students_to_grade, 1):
            assigned_display = "" if s["assigned"] is None else s["assigned"]
            print(f"{idx}. {s['name']} | {s['email']} | state={s['state']} | grade={assigned_display}")

        if apply_all:
            selected = list(range(1, len(students_to_grade) + 1))
        else:
            while True:
                sel = get_input_with_quit("Select students to grade (e.g. 1,3-5, 'a' for all, or 'q' to quit): ")
                if sel is None:
                    return graded_count
                selected = parse_selection(sel, len(students_to_grade))
                if selected:
                    break
                print("Invalid selection. Please enter valid numbers, a range, 'a' for all, or 'q' to quit.")

        score_value = score
        if score_value is None:
            raw_score = get_input_with_quit("Enter the score to assign to all selected submissions (or 'q' to quit): ")
            if raw_score is None:
                return graded_count
            try:
                score_value = float(raw_score)
            except Exception:
                print("Invalid score.")
                continue

        assigned_score = scale_score(score_value, max_points)
        if assigned_score is None:
            print("Invalid score; skipping.")
            continue
        confirm = get_input_with_quit(
            f"Assign score {assigned_score} to {len(selected)} submission(s) for '{title}'? (y/n): ",
            default="y"
        )
        if confirm is None:
            return graded_count
        if confirm.lower() not in ("y", "yes"):
            print("Aborted grading for this assignment.")
            continue

        for idx in selected:
            entry = students_to_grade[idx - 1]
            sub = entry["submission"]
            sub_id = sub.get("id")
            user_id = entry.get("user_id")
            if not sub_id or not user_id:
                continue
            try:
                service.courses().courseWork().studentSubmissions().patch(
                    courseId=course_id,
                    courseWorkId=cw_id,
                    id=sub_id,
                    updateMask="assignedGrade,draftGrade",
                    body={
                        "assignedGrade": assigned_score,
                        "draftGrade": assigned_score,
                    },
                ).execute()
                graded_count += 1
                if verbose:
                    print(f"[GClassroom] Graded {entry['name']} ({user_id}) with {assigned_score}.")
            except Exception as e:
                if verbose:
                    print(f"[GClassroom] Failed to grade {entry['name']} ({user_id}): {e}")
                else:
                    print(f"Failed to grade {entry['name']} ({user_id}).")

        print(f"Done grading selected submissions for '{title}'.")

    if verbose:
        print(f"[GClassroom] Graded {graded_count} submission(s).")
    else:
        print(f"Graded {graded_count} submission(s).")
    return graded_count
