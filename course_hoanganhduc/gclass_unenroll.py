# -*- coding: utf-8 -*-

from googleapiclient.discovery import build
from .gclass_auth import _get_google_classroom_credentials, list_google_classroom_courses
from .utils import get_input_with_quit, parse_selection

def unenroll_google_classroom_students(
    course_id=None,
    domains=None,
    emails=None,
    select_all=False,
    select_mode=False,
    credentials_path='gclassroom_credentials.json',
    token_path='token.pickle',
    apply_all=False,
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

    removed = 0
    for idx in selected_indices:
        entry = matches[idx - 1]
        try:
            service.courses().students().delete(courseId=course_id, userId=entry["user_id"]).execute()
            removed += 1
            if verbose:
                print(f"[GClassroom] Unenrolled {entry['name']} ({entry['email']}).")
        except Exception as e:
            if verbose:
                print(f"[GClassroom] Failed to unenroll {entry['name']} ({entry['email']}): {e}")
            else:
                print(f"Failed to unenroll {entry['name']} ({entry['email']}).")

    print(f"Unenrolled {removed} student(s).")
    return removed
