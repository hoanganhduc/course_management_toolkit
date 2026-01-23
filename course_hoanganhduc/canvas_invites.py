# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas invite helpers."""

import re
import time

import requests

from .canvas_auth import get_canvas_client
from .canvas_people import list_canvas_people

from .settings import (
    CANVAS_LMS_API_KEY,
    CANVAS_LMS_API_URL,
    CANVAS_LMS_COURSE_ID,
)

def invite_students_if_not_enrolled(
    students,
    course_id=CANVAS_LMS_COURSE_ID,
    role="student",
    section=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    verbose=False,
):
    """
    Invite students to Canvas if they are not already active or invited.

    Returns:
        dict: summary with results and skip counts.
    """
    if not students:
        if verbose:
            print("[CanvasInvite] No students provided for invitation.")
        else:
            print("No students provided for invitation.")
        return {"results": [], "invited": 0, "skipped_enrolled": 0, "skipped_missing": 0, "skipped_duplicate": 0}

    people = list_canvas_people(api_url=api_url, api_key=api_key, course_id=course_id, verbose=verbose)
    enrolled_emails = set()
    for group in ("active_students", "invited_students"):
        for entry in people.get(group, []) or []:
            email = (entry.get("email") or "").strip().lower()
            if email:
                enrolled_emails.add(email)

    def _get_student_email(student):
        email_val = getattr(student, "Email", None)
        if not email_val:
            for key, value in getattr(student, "__dict__", {}).items():
                if "email" in key.lower() and value:
                    email_val = value
                    break
        if email_val:
            return str(email_val).strip()
        sid = _normalize_student_id(getattr(student, "Student ID", "") or "")
        if sid:
            return f"{sid}@hus.edu.vn"
        return None

    def _get_student_name(student, fallback_email=None):
        name_val = getattr(student, "Name", None)
        if not name_val:
            for key, value in getattr(student, "__dict__", {}).items():
                if "name" in key.lower() and value:
                    name_val = value
                    break
        if name_val:
            return str(name_val).strip()
        if fallback_email and "@" in fallback_email:
            return fallback_email.split("@")[0]
        return "Student"

    def _get_student_section(student):
        for key in ("Canvas Section", "Section", "Class"):
            value = getattr(student, key, None)
            if value:
                return str(value).strip()
        return None

    results = []
    invited = 0
    skipped_enrolled = 0
    skipped_missing = 0
    skipped_duplicate = 0
    seen_emails = set()

    for student in students:
        email = _get_student_email(student)
        if not email:
            skipped_missing += 1
            if verbose:
                print("[CanvasInvite] Skipping student without email.")
            continue
        email_key = email.strip().lower()
        if email_key in seen_emails:
            skipped_duplicate += 1
            continue
        seen_emails.add(email_key)
        if email_key in enrolled_emails:
            skipped_enrolled += 1
            continue

        name = _get_student_name(student, fallback_email=email)
        student_section = _get_student_section(student) or section
        result = invite_user_to_canvas_course(
            email=email,
            name=name,
            role=role,
            section=student_section,
            api_url=api_url,
            api_key=api_key,
            course_id=course_id,
            verbose=verbose,
        )
        results.append(result)
        if result.get("success"):
            invited += 1

    if verbose:
        print(
            "[CanvasInvite] Summary: "
            f"{invited} invited, {skipped_enrolled} already enrolled, "
            f"{skipped_missing} missing email, {skipped_duplicate} duplicates."
        )
    else:
        print(
            f"Invited {invited}. "
            f"Skipped: {skipped_enrolled} enrolled, "
            f"{skipped_missing} missing email, {skipped_duplicate} duplicates."
        )

    return {
        "results": results,
        "invited": invited,
        "skipped_enrolled": skipped_enrolled,
        "skipped_missing": skipped_missing,
        "skipped_duplicate": skipped_duplicate,
    }

def invite_user_to_canvas_course(
    email,
    name,
    role="student",
    section=None,  # Add section parameter
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    verbose=False
):
    """
    Invite a single user to a Canvas course via email by creating an account and enrolling them.
    Supports roles: "student", "teacher", "ta", "observer", "designer".
    If the user already exists in the course, does not invite again.
    If section is specified, enrolls the user into that specific section.

    Args:
        email (str): User email address.
        name (str): Full name of the user.
        role (str): Role to enroll ("student", "teacher", "ta", "observer", "designer"). Default: "student".
        section (str): Optional section name to enroll user into.
        api_url (str): Canvas API base URL.
        api_key (str): Canvas API key.
        course_id (str or int): Canvas course ID.
        verbose (bool): If True, print more details; otherwise, print only important notice.

    Returns:
        dict: Success/error information about the invitation process.
    """
    try:
        if not email or not name:
            if verbose:
                print("[CanvasInvite] Email and name are required.")
            else:
                print("Email and name are required.")
            return {"error": "Email and name are required."}

        role = role.lower()
        role_map = {
            "student": "StudentEnrollment",
            "teacher": "TeacherEnrollment",
            "ta": "TaEnrollment",
            "observer": "ObserverEnrollment",
            "designer": "DesignerEnrollment"
        }
        if role not in role_map:
            if verbose:
                print(f"[CanvasInvite] Invalid role: {role}. Must be one of: {list(role_map.keys())}")
            else:
                print(f"Invalid role: {role}. Must be one of: {list(role_map.keys())}")
            return {"error": f"Invalid role: {role}"}

        canvas = get_canvas_client(api_url, api_key, verbose=verbose)
        course = canvas.get_course(course_id)

        # Check if already enrolled
        enrollments = list(course.get_enrollments(type=[role_map[role]], include=['user']))
        for enrollment in enrollments:
            user = getattr(enrollment, "user", {})
            user_email = user.get("login_id", "").lower()
            if user_email == email.lower():
                if verbose:
                    print(f"[CanvasInvite] User with email {email} is already enrolled in course {course_id} as {role}.")
                else:
                    print(f"User with email {email} is already enrolled in course {course_id} as {role}.")
                return {"already_enrolled": True, "email": email, "name": name, "role": role}

        # Find section ID if section is specified
        section_id = None
        if section:
            try:
                course_sections = list(course.get_sections())
                for s in course_sections:
                    if s.name.lower() == section.lower():
                        section_id = s.id
                        if verbose:
                            print(f"[CanvasInvite] Found section '{section}' with ID {section_id}")
                        else:
                            print(f"Found section '{section}' with ID {section_id}")
                        break
                if not section_id:
                    if verbose:
                        print(f"[CanvasInvite] Warning: Section '{section}' not found in course {course_id}")
                    else:
                        print(f"Warning: Section '{section}' not found in course {course_id}")
            except Exception as section_error:
                if verbose:
                    print(f"[CanvasInvite] Error finding section: {section_error}")
                else:
                    print(f"Error finding section: {section_error}")

        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }

        account_id = 10 if "canvas.instructure.com" in api_url else 1
        account_url = f"{api_url}/api/v1/accounts/{account_id}/users"

        name_parts = name.strip().split()
        if len(name_parts) >= 2:
            first_name = name_parts[0]
            last_name = name_parts[-1]
            short_name = f"{first_name} {last_name}"
            sortable_name = f"{last_name}, {first_name}"
        else:
            short_name = name
            sortable_name = name

        account_data = {
            'user': {
                'name': name,
                'short_name': short_name,
                'sortable_name': sortable_name,
                'terms_of_use': True
            },
            'pseudonym': {
                'unique_id': email,
                'force_self_registration': True
            }
        }

        account_response = requests.post(account_url, headers=headers, json=account_data)
        user_id = None

        if account_response.status_code == 400:
            response_data = account_response.json()
            error_message = str(response_data).lower()
            if 'already belongs to a user' in error_message or 'id already in use' in error_message:
                if verbose:
                    print(f"[CanvasInvite] Email {email} already has a Canvas account. Finding existing user...")
                else:
                    print(f"Email {email} already has a Canvas account. Finding existing user...")
                
                # Find user ID by querying /api/v1/accounts/self/users with search_term=<email>
                search_url = f"{api_url}/api/v1/accounts/self/users?search_term={email}"
                search_response = requests.get(search_url, headers=headers)
                if search_response.status_code == 200:
                    users = search_response.json()
                    for user in users:
                        if user.get('login_id', '').lower() == email.lower():
                            user_id = user.get('id')
                            if verbose:
                                print(f"[CanvasInvite] Found existing user with ID {user_id} via search API")
                            else:
                                print(f"Found existing user with ID {user_id} via search API")
                            break
                if not user_id:
                    if verbose:
                        print(f"[CanvasInvite] Could not find existing user for email {email}")
                    else:
                        print(f"Could not find existing user for email {email}")
                    return {"error": f"Could not find existing user for email {email}"}
            else:
                if verbose:
                    print(f"[CanvasInvite] Account creation failed: {response_data}")
                else:
                    print(f"Account creation failed: {response_data}")
                return {"error": f"Account creation failed: {response_data}"}
        elif account_response.status_code != 200:
            if verbose:
                print(f"[CanvasInvite] Account creation failed: {account_response.status_code} - {account_response.text}")
            else:
                print(f"Account creation failed: {account_response.status_code} - {account_response.text}")
            return {"error": f"Account creation failed: {account_response.status_code} - {account_response.text}"}
        else:
            new_user_data = account_response.json()
            user_id = new_user_data.get('id')
            if verbose:
                print(f"[CanvasInvite] Canvas account created for {email} (User ID: {user_id})")
            else:
                print(f"Canvas account created for {email} (User ID: {user_id})")

        # Enroll user using POST to /api/v1/courses/:course_id/enrollments
        enroll_url = f'{api_url}/api/v1/courses/{course_id}/enrollments'
        enroll_data = {
            'enrollment[user_id]': user_id,
            'enrollment[type]': role_map[role],
            'enrollment[enrollment_state]': 'active',
            'enrollment[notify]': True
        }
        
        # Add section_id to enrollment if specified and found
        if section_id:
            enroll_data['enrollment[course_section_id]'] = section_id
            enroll_data['enrollment[limit_privileges_to_course_section]'] = True
            if verbose:
                print(f"[CanvasInvite] Enrolling user in section ID {section_id}")
            else:
                print(f"Enrolling user in section ID {section_id}")
                
        headers_form = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        enroll_response = requests.post(enroll_url, headers=headers_form, data=enroll_data)
        if enroll_response.status_code in [200, 201]:
            enrollment_data = enroll_response.json()
            section_info = f" in section '{section}'" if section_id else ""
            if verbose:
                print(f"[CanvasInvite] Successfully enrolled {email} in course {course_id}{section_info} as {role}")
            else:
                print(f"Successfully enrolled {email} in course {course_id}{section_info} as {role}")
            return {
                "success": True,
                "email": email,
                "name": name,
                "role": role,
                "user_id": user_id,
                "section_id": section_id,
                "enrollment_state": enrollment_data.get('enrollment_state', 'unknown')
            }
        else:
            if verbose:
                print(f"[CanvasInvite] Enrollment failed: {enroll_response.status_code} - {enroll_response.text}")
            else:
                print(f"Enrollment failed: {enroll_response.status_code} - {enroll_response.text}")
            return {
                "error": f"Enrollment failed: {enroll_response.status_code} - {enroll_response.text}",
                "email": email,
                "name": name,
                "role": role
            }
    except Exception as e:
        if verbose:
            print(f"[CanvasInvite] Error inviting user: {e}")
        else:
            print(f"Error inviting user: {e}")
        return {"error": str(e), "email": email, "name": name, "role": role}

def invite_users_to_canvas_course(
    data_file,
    name=None,
    role="student",
    section=None,
    course_id=CANVAS_LMS_COURSE_ID,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    verbose=False
):
    """
    Invite one or more users (student, teacher, or TA) to a Canvas course from a TXT file or a string.

    Supported input formats for `data_file`:
      - Path to a TXT file (each line: "name,email,role,section" or "name,email,role" or "name,email" or just "email"; also supports semicolon-separated pairs in file)
      - A string of pairs: "<name>,<email>,<role>,<section>;<name>,<email>,<role>,<section>;..." (semicolon-separated)
      - A string of emails separated by commas/semicolons/spaces/newlines (if `name` is provided separately)
      - A string of alternating "name,email,role,section,name,email,role,section,..." (if no semicolons, will try to group as quads)
      - If only email is provided, the part before @ is used as the name, and `role` and `section` arguments are used.

    Args:
        data_file (str): Path to TXT file or string with user info.
            - If a file path exists, reads from file.
            - Otherwise, parses as string input.
        name (str or None): Optional name(s) if data_file is only emails.
            - If provided, should be comma/semicolon/space/newline separated, matching the number of emails.
        role (str): Role to enroll ("student", "teacher", "ta"). Default: "student".
            - Can be overridden per user if specified in the input.
        section (str): Optional Canvas section name to enroll users into.
            - Can be overridden per user if specified in the input.
        course_id (str): Canvas course ID.
        api_url (str): Canvas API URL.
        api_key (str): Canvas API key.
        verbose (bool): If True, print more details; otherwise, print only important notice.

    Returns:
        List of dicts with invitation results for each user.
        Each dict contains success/error info for that user.

    Details:
        - If a line or entry is "name,email,role,section", uses all four.
        - If "name,email,role", uses provided section argument.
        - If "name,email", uses provided role and section arguments.
        - If only email is provided, uses the part before "@" as the name and provided role and section.
        - Skips invalid emails and prints a warning.
        - Calls invite_user_to_canvas_course() for each user.
        - Prints a summary of how many users were invited.
        - Handles both file and string input robustly.
        - Returns a list of results (one per user processed).
    """
    results = []
    email_pattern = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")
    valid_roles = {"student", "teacher", "ta"}

    # Helper to parse a line or entry
    def parse_entry(line, default_role, default_section):
        # Try "name,email,role,section"
        m = re.match(r'^(.*?),(.*?),(.*?),(.*)$', line)
        if m:
            n = m.group(1).strip()
            e = m.group(2).strip()
            r = m.group(3).strip().lower()
            s = m.group(4).strip()
            if not email_pattern.match(e):
                return None, None, None, None
            if r not in valid_roles:
                r = default_role
            return n, e, r, s
        # Try "name,email,role"
        m = re.match(r'^(.*?),(.*?),(.*)$', line)
        if m:
            n = m.group(1).strip()
            e = m.group(2).strip()
            r = m.group(3).strip().lower()
            if not email_pattern.match(e):
                return None, None, None, None
            if r not in valid_roles:
                r = default_role
            return n, e, r, default_section
        # Try "name,email"
        m = re.match(r'^(.*?),(.*)$', line)
        if m:
            n = m.group(1).strip()
            e = m.group(2).strip()
            if not email_pattern.match(e):
                return None, None, None, None
            return n, e, default_role, default_section
        # Try just email
        if "@" in line and email_pattern.match(line.strip()):
            e = line.strip()
            n = e.split("@")[0]
            return n, e, default_role, default_section
        return None, None, None, None

    # Read and normalize input lines
    lines = []
    if os.path.exists(data_file):
        with open(data_file, "r", encoding="utf-8") as f:
            file_content = f.read()
        # Support both one-user-per-line and multi-pair string format in file
        # If file contains semicolons, treat as multi-pair string
        if ";" in file_content:
            raw_lines = [x.strip() for x in file_content.split(";") if x.strip()]
        else:
            # Otherwise, treat each non-empty line as a record
            raw_lines = [line.strip() for line in file_content.splitlines() if line.strip()]
        lines = raw_lines
    else:
        # Otherwise, treat as string of pairs/triples/quads or emails
        # Try to split by semicolon first (for "name,email,role,section;name,email,role,section;...")
        if ";" in data_file:
            lines = [x.strip() for x in data_file.split(";") if x.strip()]
        elif "," in data_file and "@" in data_file:
            # Could be "name,email,role,section,name,email,role,section" or "name,email,role,name,email,role" etc
            # If name is provided separately, use it
            if name:
                emails = [x.strip() for x in re.split(r"[;, \n]", data_file) if x.strip()]
                names = [x.strip() for x in re.split(r"[;, \n]", name) if x.strip()]
                if len(names) == len(emails):
                    lines = [f"{n},{e}" for n, e in zip(names, emails)]
                else:
                    # Fallback: treat as emails only
                    lines = emails
            else:
                # Try to group as quads (name,email,role,section)
                parts = [x.strip() for x in data_file.split(",") if x.strip()]
                i = 0
                while i < len(parts):
                    if i + 3 < len(parts) and "@" in parts[i+1]:
                        # name,email,role,section
                        n, e, r, s = parts[i], parts[i+1], parts[i+2].lower(), parts[i+3]
                        if r not in valid_roles:
                            r = role
                        lines.append(f"{n},{e},{r},{s}")
                        i += 4
                    elif i + 2 < len(parts) and "@" in parts[i+1]:
                        # name,email,role
                        n, e, r = parts[i], parts[i+1], parts[i+2].lower()
                        if r not in valid_roles:
                            r = role
                        lines.append(f"{n},{e},{r}")
                        i += 3
                    elif i + 1 < len(parts) and "@" in parts[i+1]:
                        # name,email
                        lines.append(f"{parts[i]},{parts[i+1]}")
                        i += 2
                    else:
                        lines.append(parts[i])
                        i += 1
        else:
            # Space/comma/semicolon/newline separated emails
            lines = [x.strip() for x in re.split(r"[;, \n]", data_file) if x.strip()]

    for idx, line in enumerate(lines):
        n, e, r, s = parse_entry(line, role, section)
        if not e or not email_pattern.match(e):
            if verbose:
                print(f"[CanvasInvite] Invalid email format: {e} (line: {line})")
            else:
                print(f"Invalid email: {e}")
            continue
        result = invite_user_to_canvas_course(
            email=e,
            name=n,
            role=r,
            section=s,
            api_url=api_url,
            api_key=api_key,
            course_id=course_id,
            verbose=verbose
        )
        results.append(result)
    if verbose:
        print(f"[CanvasInvite] Invited {len(results)} users from input.")
    else:
        print(f"Invited {len(results)} users from input.")
    return results

