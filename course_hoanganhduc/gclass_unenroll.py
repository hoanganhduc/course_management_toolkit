# -*- coding: utf-8 -*-

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from .gclass_auth import _get_google_classroom_credentials, list_google_classroom_courses
from .utils import get_input_with_quit, parse_selection
from .data import load_database, save_database

def unenroll_google_classroom_students(
    course_id=None,
    domains=None,
    emails=None,
    select_all=False,
    select_mode=False,
    credentials_path='gclassroom_credentials.json',
    token_path='token.pickle',
    apply_all=False,
    missing_student_id=False,
    db_path=None,
    update_local_db=True,
    dry_run=False,
    verbose=False,
):
    """
    Unenroll students from a Google Classroom course by email domain.
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

    domains_list = normalize_domains(domains)
    emails_list = normalize_emails(emails)

    missing_email_set = set()
    missing_google_id_set = set()
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
            gid = (getattr(student, "Google_ID", "") or "").strip()
            if email:
                missing_email_set.add(email)
            if gid:
                missing_google_id_set.add(gid)
        if not missing_email_set and not missing_google_id_set:
            print("No students missing Student ID in the local database.")
            return 0

    roster = []
    next_token = None
    while True:
        req = service.courses().students().list(courseId=course_id, pageToken=next_token, pageSize=200) if next_token else service.courses().students().list(courseId=course_id, pageSize=200)
        resp = req.execute()
        roster.extend(resp.get("students", []) or [])
        next_token = resp.get("nextPageToken")
        if not next_token:
            break

    matches = []
    for entry in roster:
        profile = entry.get("profile", {}) or {}
        email = (profile.get("emailAddress") or "").strip()
        user_id = entry.get("userId")
        name = (profile.get("name", {}) or {}).get("fullName") or ""
        if not email or not user_id:
            continue
        email_lower = email.lower()
        if select_mode:
            matches.append({
                "user_id": str(user_id),
                "name": name,
                "email": email,
            })
            continue
        if missing_student_id:
            if (email_lower and email_lower in missing_email_set) or (str(user_id) in missing_google_id_set):
                matches.append({
                    "user_id": str(user_id),
                    "name": name,
                    "email": email,
                })
            continue
        if emails_list and email_lower in emails_list:
            matches.append({
                "user_id": str(user_id),
                "name": name,
                "email": email,
            })
            continue
        if domains_list and any(email_lower.endswith("@" + d) for d in domains_list):
            matches.append({
                "user_id": str(user_id),
                "name": name,
                "email": email,
            })

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
        label = "; ".join(label_parts)
        print(f"Matched {len(matches)} student(s) for {label}.")
    for idx, m in enumerate(matches, 1):
        print(f"{idx}. {m['name']} | {m['email']}")

    if apply_all or select_all:
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
    removed_user_ids = []
    for idx in selected_indices:
        entry = matches[idx - 1]
        try:
            service.courses().students().delete(courseId=course_id, userId=entry["user_id"]).execute()
            try:
                service.courses().students().get(courseId=course_id, userId=entry["user_id"]).execute()
                print(f"Failed to unenroll {entry['name']} ({entry['email']}): still enrolled.")
                continue
            except HttpError as exc:
                status = getattr(getattr(exc, "resp", None), "status", None)
                if status not in (403, 404):
                    if verbose:
                        print(f"[GClassroom] Verification failed for {entry['name']} ({entry['email']}): {exc}")
                else:
                    removed += 1
                    removed_emails.append(entry["email"].strip().lower())
                    removed_user_ids.append(str(entry["user_id"]))
                    if verbose:
                        print(f"[GClassroom] Unenrolled {entry['name']} ({entry['email']}).")
        except Exception as e:
            if verbose:
                print(f"[GClassroom] Failed to unenroll {entry['name']} ({entry['email']}): {e}")
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
            removed_user_id_set = set(removed_user_ids)
            for student in students:
                email = (getattr(student, "Email", "") or "").strip().lower()
                google_id = (getattr(student, "Google_ID", "") or "").strip()
                if (email and email in removed_email_set) or (google_id and google_id in removed_user_id_set):
                    removed_local += 1
                    continue
                kept.append(student)
            if removed_local:
                save_database(kept, db_path=db_path, verbose=verbose, audit_source="gclass-unenroll")
                print(f"Removed {removed_local} student(s) from local database.")
    return removed
