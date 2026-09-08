# -*- coding: utf-8 -*-

import os
from types import SimpleNamespace
from googleapiclient.discovery import build
from .gclass_auth import _get_google_classroom_credentials, list_google_classroom_courses
from .settings import GOOGLE_CLASSROOM_GRADE_CATEGORY_METHOD
from .models import *
from .data import load_database, save_database


def _display_name_of(student):
    """The Classroom display name, under either spelling it has been stored as.

    ``normalize_columns`` maps the form column to "Google Classroom Display
    Name" while this module writes ``Google_Classroom_Display_Name``.  A reader
    that knows only one of the two misses every record written by the other
    lane, which is why a name this function had just stored never matched on
    the following run.
    """
    for attribute in ("Google_Classroom_Display_Name", "Google Classroom Display Name"):
        value = getattr(student, attribute, "") or ""
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def sync_students_with_google_classroom(students, db_path=None, course_id=None, credentials_path='gclassroom_credentials.json', token_path='token.pickle', fetch_grades=False, verbose=False):
    """
    Sync students in the local database with active students from Google Classroom.
    For each student fetched from Google Classroom:
        - match by Google_ID first, then by display name, name or email (case-insensitive)
        - if matched: fill missing local fields from Google data (Google_ID, Email, Google_Classroom_Display_Name)
        - if not matched: create a new student entry with Name, Email, Google_ID and Google_Classroom_Display_Name
    Optionally fetch grades/submission state when fetch_grades=True.

    Rows the course no longer contains are reported, never removed: only the
    teacher can tell a student who has left from one who filled the form before
    being enrolled.

    Returns (added_count, updated_count).
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
        def _coerce_float(value):
            try:
                return float(value)
            except Exception:
                return None

        def _normalize_gc_grade_category_method(value):
            if value is None:
                return "average"
            method = str(value).strip().lower()
            if method in ("avg", "average", "mean"):
                return "average"
            if method in ("weighted", "weight", "ratio"):
                return "weighted"
            if method in ("sum", "total"):
                return "sum"
            return "average"

        # load current DB list only if no students were passed in
        if (students is None or students == []) and db_path and os.path.exists(db_path):
            try:
                students = load_database(db_path, verbose=verbose)
            except Exception:
                if verbose:
                    print("[GClassroom] Failed to load DB, proceeding with provided students list.")

        # ensure students is a list
        if students is None:
            students = []

        # helper to create Student object if no Student class defined
        try:
            Student  # type: ignore
            StudentClass = Student  # use existing Student class if defined
        except Exception:
            StudentClass = lambda **kw: SimpleNamespace(**kw)

        # authenticate
        creds = _get_google_classroom_credentials(credentials_path, token_path, verbose=verbose)

        service = build("classroom", "v1", credentials=creds)

        # if course_id not provided, list and ask user to select
        if not course_id:
            resp = service.courses().list(pageSize=50).execute()
            courses = resp.get("courses", []) or []
            if not courses:
                if verbose:
                    print("[GClassroom] No courses available for this account.")
                else:
                    print("No courses found.")
                return 0, 0
            print("Available Google Classroom courses:")
            for i, c in enumerate(courses, 1):
                print(f"{i}. {c.get('name')} (ID: {c.get('id')})")
            while True:
                try:
                    sel = input("Select course number: ").strip()
                except EOFError:
                    print("No terminal to select a course on. Pass --google-course-id to run without one.")
                    return 0, 0
                if not sel:
                    continue
                try:
                    idx = int(sel) - 1
                    if 0 <= idx < len(courses):
                        course_id = courses[idx]["id"]
                        break
                except Exception:
                    continue

        # fetch all students (handle pagination)
        classroom_students = []
        next_token = None
        while True:
            req = service.courses().students().list(courseId=course_id, pageToken=next_token, pageSize=200) if next_token else service.courses().students().list(courseId=course_id, pageSize=200)
            resp = req.execute()
            classroom_students.extend(resp.get("students", []) or [])
            next_token = resp.get("nextPageToken")
            if not next_token:
                break

        # Build lookup maps from local students for matching incoming records.
        local_by_google_id = {}
        local_by_google_name = {}
        local_by_name = {}
        local_by_email = {}
        # ensure keys are lowercased for robust matching
        for s in students:
            gname = _display_name_of(s)
            name = getattr(s, "Name", "") or ""
            email = getattr(s, "Email", "") or ""
            gid = getattr(s, "Google_ID", "") or ""
            if isinstance(gid, str) and gid.strip():
                local_by_google_id[gid.strip()] = s
            if isinstance(gname, str) and gname.strip():
                local_by_google_name[gname.strip().lower()] = s
            if isinstance(name, str) and name.strip():
                local_by_name[name.strip().lower()] = s
            if isinstance(email, str) and email.strip():
                local_by_email[email.strip().lower()] = s

        # Resolve by Google_ID first, then display name > name > email. A shared
        # Google_ID is the same account and so the same person, whatever the two
        # records say their name or address is; matching on the strings alone is
        # what used to create a second row for a student who joined Classroom
        # with a personal address and filled the form with the school one.
        # If several candidates remain, prompt the operator to pick, create a new
        # student, or skip.
        def _resolve_gc_match(google_id, name_key, email_key):
            if google_id and google_id in local_by_google_id:
                return local_by_google_id[google_id]

            candidates = []
            seen = set()

            def add_candidate(label, student):
                if id(student) in seen:
                    return
                candidates.append((label, student))
                seen.add(id(student))

            if name_key and name_key in local_by_google_name:
                add_candidate("google_name", local_by_google_name[name_key])
            if name_key and name_key in local_by_name:
                add_candidate("name", local_by_name[name_key])
            if email_key and email_key in local_by_email:
                add_candidate("email", local_by_email[email_key])

            if not candidates:
                return None
            if len(candidates) == 1:
                return candidates[0][1]

            print("\n[GClassroom] Possible duplicate match detected:")
            for idx, (label, student) in enumerate(candidates, 1):
                s_name = getattr(student, "Name", "") or ""
                s_email = getattr(student, "Email", "") or ""
                s_gid = getattr(student, "Google_ID", "") or ""
                print(f"{idx}. {s_name} | {s_email} | Google ID: {s_gid} (matched by {label})")
            print("n. Create new student")
            print("s. Skip this record")
            while True:
                try:
                    choice = input("Choose a match (number), 'n' for new, or 's' to skip: ").strip().lower()
                except EOFError:
                    # Nobody is there to answer. Creating a row on a guess is the
                    # one outcome that cannot be undone by re-running, so the
                    # record is skipped and reported instead.
                    print("[GClassroom] No terminal to ask on; skipping this record.")
                    return "__skip__"
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
        unresolved = []
        seen_google_ids = set()

        for cs in classroom_students:
            profile = cs.get("profile", {}) or {}
            email = (profile.get("emailAddress") or "").strip()
            full_name = (profile.get("name", {}).get("fullName") or "").strip()
            google_id = cs.get("userId", "") or ""

            if not full_name:
                # skip unknown entries
                continue

            google_id = google_id.strip()
            if google_id:
                seen_google_ids.add(google_id)

            key_name = full_name.lower()
            matched = _resolve_gc_match(google_id, key_name, email.lower() if email else None)
            match_type = None
            if matched == "__skip__":
                unresolved.append((full_name, email or "(no email)"))
                continue
            if matched:
                if google_id and local_by_google_id.get(google_id) is matched:
                    match_type = "google_id"
                elif key_name in local_by_google_name and local_by_google_name[key_name] is matched:
                    match_type = "google_name"
                elif key_name in local_by_name and local_by_name[key_name] is matched:
                    match_type = "name"
                elif email and email.lower() in local_by_email and local_by_email[email.lower()] is matched:
                    match_type = "email"

            if matched:
                changed = False
                # Do not overwrite existing local Name if present; only fill missing fields
                if not getattr(matched, "Google_ID", "") and google_id:
                    matched.Google_ID = google_id
                    local_by_google_id[google_id] = matched
                    changed = True
                if not getattr(matched, "Email", "") and email:
                    matched.Email = email
                    changed = True
                if not _display_name_of(matched) and full_name:
                    matched.Google_Classroom_Display_Name = full_name
                    changed = True
                if changed:
                    updated_count += 1
                    if verbose:
                        print(f"[GClassroom] Updated local student from GC: {full_name} ({match_type})")
            else:
                # create new student entry
                new_student = StudentClass(
                    Name=full_name,
                    Email=email,
                    Google_ID=google_id,
                    Google_Classroom_Display_Name=full_name
                )
                # Append to local list and update maps to prevent duplicate additions.
                students.append(new_student)
                if google_id:
                    local_by_google_id[google_id] = new_student
                local_by_google_name[key_name] = new_student
                local_by_name[key_name] = new_student
                if email:
                    local_by_email[email.lower()] = new_student
                added_count += 1
                if verbose:
                    print(f"[GClassroom] Added new student: {full_name} ({email})")

        # Read-only reconciliation. Nothing is deleted or marked here: a row can
        # be stale, or it can belong to a student who filled the form before
        # being enrolled, and only the teacher can tell the two apart.
        left_classroom = []
        never_matched = []
        for s in students:
            gid = str(getattr(s, "Google_ID", "") or "").strip()
            label = f"{str(getattr(s, 'Student ID', '') or '').strip() or '--------'}  {getattr(s, 'Name', '') or '(no name)'}"
            email_label = str(getattr(s, "Email", "") or "").strip() or "(no email)"
            if gid:
                if gid not in seen_google_ids:
                    left_classroom.append(f"{label}  <{email_label}>")
            else:
                never_matched.append(f"{label}  <{email_label}>")

        print(f"Google Classroom returned {len(classroom_students)} student(s); the database holds {len(students)}.")
        if left_classroom:
            print(
                f"Notice: {len(left_classroom)} row(s) carry a Google ID that is no longer "
                "in this course. Remove them by hand if the students have left:"
            )
            for row in left_classroom:
                print(f"  {row}")
        if never_matched:
            print(f"Notice: {len(never_matched)} row(s) have no Google ID and were never matched to a Classroom account:")
            for row in never_matched:
                print(f"  {row}")
        if unresolved:
            print(f"Notice: {len(unresolved)} Classroom record(s) were skipped as ambiguous and are not in the database:")
            for gc_name, gc_email in unresolved:
                print(f"  {gc_name} <{gc_email}>")

        # optionally fetch coursework and grades (if requested)
        if fetch_grades:
            if verbose:
                print("[GClassroom] Fetching coursework and student submission data...")
            topic_map = {}
            try:
                topics = []
                next_token = None
                while True:
                    req = service.courses().topics().list(courseId=course_id, pageToken=next_token, pageSize=100) if next_token else service.courses().topics().list(courseId=course_id, pageSize=100)
                    resp = req.execute()
                    topics.extend(resp.get("topic", []) or resp.get("topics", []) or [])
                    next_token = resp.get("nextPageToken")
                    if not next_token:
                        break
                for topic in topics:
                    topic_id = topic.get("topicId") or topic.get("id")
                    if not topic_id:
                        continue
                    topic_map[topic_id] = topic.get("name") or f"Topic {topic_id}"
            except Exception:
                if verbose:
                    print("[GClassroom] Warning: could not fetch topics; topic summaries will be uncategorized.")
            # fetch all coursework (paginated)
            coursework = []
            next_token = None
            while True:
                req = service.courses().courseWork().list(courseId=course_id, pageToken=next_token, pageSize=200) if next_token else service.courses().courseWork().list(courseId=course_id, pageSize=200)
                resp = req.execute()
                coursework.extend(resp.get("courseWork", []) or [])
                next_token = resp.get("nextPageToken")
                if not next_token:
                    break
            category_method = _normalize_gc_grade_category_method(GOOGLE_CLASSROOM_GRADE_CATEGORY_METHOD)

            # for each local student with Google_ID, fetch submissions per coursework
            for s in students:
                gid = getattr(s, "Google_ID", "") or ""
                if not gid:
                    continue
                grades = getattr(s, "Grades", {}) or {}
                submissions = getattr(s, "Submissions", {}) or {}
                submission_details = getattr(s, "Google_Classroom_Submission_Details", {}) or {}
                topic_stats = {}
                for cw in coursework:
                    cw_id = cw.get("id")
                    if not cw_id:
                        continue
                    title = cw.get("title", f"cw_{cw_id}")
                    try:
                        # list studentSubmissions filtered by userId
                        resp = service.courses().courseWork().studentSubmissions().list(
                            courseId=course_id, courseWorkId=cw_id, userId=gid, pageSize=50
                        ).execute()
                        subs = resp.get("studentSubmissions", []) or []
                        if subs:
                            sub = subs[0]
                            state = sub.get("state")
                            # assignedGrade may be under "assignedGrade" or in assignedGrade field
                            grade = sub.get("assignedGrade")
                            # maxPoints often available on coursework object
                            max_points = cw.get("maxPoints")
                            if grade is not None:
                                scaled = _scale_score_to_ten(grade, max_points)
                                scaled_max_points = None
                                if max_points not in (None, 0) or (isinstance(grade, (int, float)) and grade > 10):
                                    scaled_max_points = 10
                                    grades[title] = {"grade": scaled, "max_points": scaled_max_points}
                                else:
                                    scaled_max_points = max_points
                                    grades[title] = {"grade": scaled, "max_points": scaled_max_points}
                                topic_id = cw.get("topicId")
                                topic_name = topic_map.get(topic_id) if topic_id else None
                                if not topic_name:
                                    topic_name = "Uncategorized"
                                stats = topic_stats.get(topic_name)
                                if stats is None:
                                    stats = {
                                        "scaled_total": 0.0,
                                        "scaled_max_total": 0.0,
                                        "raw_total": 0.0,
                                        "raw_max_total": 0.0,
                                        "count": 0,
                                    }
                                    topic_stats[topic_name] = stats
                                scaled_val = _coerce_float(scaled)
                                raw_val = _coerce_float(grade)
                                raw_max_val = _coerce_float(max_points)
                                scaled_max_val = _coerce_float(scaled_max_points)
                                if scaled_val is not None:
                                    stats["scaled_total"] += scaled_val
                                    if scaled_max_val is not None:
                                        stats["scaled_max_total"] += scaled_max_val
                                    if raw_val is not None:
                                        stats["raw_total"] += raw_val
                                    if raw_max_val is not None:
                                        stats["raw_max_total"] += raw_max_val
                                    stats["count"] += 1
                                    stats["count"] += 1
                            submissions[title] = state
                            
                            # Capture attachment details
                            attachments = sub.get("assignmentSubmission", {}).get("attachments") or []
                            submitted_at = sub.get("updateTime") or sub.get("creationTime")
                            if attachments or submitted_at:
                                files_info = []
                                for att in attachments:
                                    name = "Unknown Attachment"
                                    ext = ""

                                    if "driveFile" in att:
                                        drive_file = att["driveFile"]
                                        file_title = drive_file.get("title", "Unknown Drive File")
                                        name = file_title
                                        ext = os.path.splitext(file_title)[1] if file_title else ""
                                    elif "link" in att:
                                        link = att["link"]
                                        name = link.get("title", link.get("url", "External Link"))
                                        ext = ".url"
                                    elif "form" in att:
                                        form = att["form"]
                                        name = form.get("title", form.get("formUrl", "Google Form"))
                                        ext = ".form"
                                    elif "youTubeVideo" in att:
                                        video = att["youTubeVideo"]
                                        name = video.get("title", f"YouTube Video ({video.get('id', 'Unknown')})")
                                        ext = ".video"

                                    files_info.append({
                                        "name": name,
                                        "ext": ext
                                    })
                                
                                submission_details[title] = {
                                    "attachment_count": len(attachments),
                                    "submitted_at": submitted_at,
                                    "files": files_info
                                }
                    except Exception:
                        # ignore per-assignment errors but continue
                        if verbose:
                            print(f"[GClassroom] Warning: could not fetch submission for student {getattr(s,'Name','?')} cw:{title}")
                s.Grades = grades
                s.Submissions = submissions
                s.Google_Classroom_Submission_Details = submission_details
                topic_grades = {}
                for topic_name, stats in topic_stats.items():
                    count = stats.get("count", 0)
                    if count <= 0:
                        continue
                    if category_method == "average":
                        grade_val = round(stats["scaled_total"] / count, 2)
                        max_val = 10
                    elif category_method == "weighted":
                        raw_max_total = stats.get("raw_max_total", 0.0)
                        if raw_max_total > 0:
                            grade_val = round(stats.get("raw_total", 0.0) / raw_max_total * 10, 2)
                        else:
                            grade_val = round(stats["scaled_total"] / count, 2)
                        max_val = 10
                    else:
                        grade_val = round(stats["scaled_total"], 2)
                        max_val = round(stats["scaled_max_total"], 2) if stats.get("scaled_max_total") else None
                    topic_grades[topic_name] = {
                        "grade": grade_val,
                        "max_points": max_val,
                        "count": count,
                        "method": category_method,
                    }
                s.Google_Classroom_Topic_Grades = topic_grades
            if verbose:
                print("[GClassroom] Grades/submissions fetch complete.")

        # save back to db if requested
        if db_path:
            try:
                save_database(students, db_path, verbose=verbose)
            except Exception as e:
                if verbose:
                    print(f"[GClassroom] Warning: failed to save DB: {e}")

        if verbose:
            print(f"[GClassroom] Sync finished. Added: {added_count}, Updated: {updated_count}")
        else:
            print(f"Sync completed: {added_count} added, {updated_count} updated")

        return added_count, updated_count

    except Exception as e:
        # handle auth errors specially: if token exists and auth failure, remove token to force reauth
        try:
            err_status = None
            if hasattr(e, "resp"):
                err_status = getattr(e.resp, "status", None)
            if err_status == 401 and os.path.exists(token_path):
                try:
                    os.remove(token_path)
                except Exception:
                    pass
                if verbose:
                    print("[GClassroom] Authentication failed; token removed. Re-run to re-authenticate.")
                else:
                    print("Authentication failed. Please re-run to re-authenticate.")
            else:
                if verbose:
                    print(f"[GClassroom] Error during sync: {e}")
                else:
                    print(f"Error syncing with Google Classroom: {e}")
        except Exception:
            if verbose:
                print(f"[GClassroom] Error during exception handling: {e}")
            else:
                print("Error syncing with Google Classroom.")
        return 0, 0
