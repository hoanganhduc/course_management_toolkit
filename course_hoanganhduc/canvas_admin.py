# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas admin helpers."""

import signal
from datetime import datetime

from .canvas_auth import get_canvas_client

from .canvas_utils import prompt_with_timeout, parse_selection, normalize_datetime_str
from .settings import (
    CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    CANVAS_LMS_API_KEY,
    CANVAS_LMS_API_URL,
    CANVAS_LMS_COURSE_ID,
)

def change_canvas_deadlines(
    assignment_ids=None,
    new_due_date=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    verbose=False
):
    """
    Change deadlines for one or more Canvas assignments.
    - If assignment_ids is None, list assignments (optionally filtered by category) and allow interactive selection (supports ranges, comma-separated, 'a' for all).
    - Ask whether to apply a single deadline to all selected assignments or specify different dates per assignment.
    - Prompts time out after 60 seconds and the operation quits (default).
    - new_due_date (if provided) should be a string in "YYYY-MM-DD HH:MM" or "YYYY-MM-DDTHH:MM" format (local time). If provided and multiple assignments selected, it will be used as the single deadline option.
    Returns a dict mapping assignment_id -> status string.
    Note: this implementation assumes the local time provided by the user is in the same timezone as Canvas (no timezone conversions are performed).
    """
    def timeout_handler(_signum, _frame):
        print("\nTimeout: No response after 60 seconds. Quitting...")
        raise TimeoutError("User input timeout")

    def _prompt(prompt, timeout=60, default=None):
        return prompt_with_timeout(prompt, timeout=timeout, default=default, verbose=verbose)

    def parse_selection(sel, n):
        sel = sel.strip().lower()
        if sel in ("a", "all"):
            return list(range(1, n + 1))
        parts = set()
        for part in sel.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                try:
                    start, end = map(int, part.split("-", 1))
                    parts.update(range(start, end + 1))
                except Exception:
                    continue
            else:
                if part.isdigit():
                    parts.add(int(part))
        return sorted(i for i in parts if 1 <= i <= n)

    def normalize_datetime_str(s):
        # Accept "YYYY-MM-DD HH:MM" or "YYYY-MM-DDTHH:MM" or "YYYY-MM-DD"
        # Also accept trailing GMT/UTC or trailing 'Z'. Treat inputs as GMT (UTC).
        if s is None:
            return None
        s = s.strip()
        if not s:
            return None
        # Remove trailing "GMT" or "UTC" (case-insensitive) and normalize trailing Z
        s_clean = re.sub(r'\s*(?:gmt|utc)$', '', s, flags=re.IGNORECASE)
        # If ends with Z keep it (ISO)
        try_formats = ("%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M", "%Y-%m-%d")
        for fmt in try_formats:
            try:
                dt = datetime.strptime(s_clean, fmt)
                # Treat user-provided time as GMT (UTC) and output in Canvas-friendly Z format
                if fmt == "%Y-%m-%d":
                    return dt.strftime("%Y-%m-%dT00:00:00Z")
                return dt.strftime("%Y-%m-%dT%H:%M:%SZ")
            except Exception:
                continue
        # Try more flexible parsing (ISO)
        try:
            # fromisoformat doesn't accept trailing 'Z', so replace if present
            iso_candidate = s_clean
            if iso_candidate.endswith("Z"):
                iso_candidate = iso_candidate[:-1]
            dt = datetime.fromisoformat(iso_candidate)
            return dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        except Exception:
            return None

    results = {}

    try:
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)
    except Exception as e:
        if verbose:
            print(f"[ChangeDeadlines] Failed to connect to Canvas: {e}")
        else:
            print("Failed to connect to Canvas.")
        return results

    # Gather assignments (filtered by category if provided)
    try:
        groups = list(course.get_assignment_groups(include=["assignments"]))
        assignments = []
        for g in groups:
            group_name = g.name
            if category and group_name.lower() != category.lower():
                continue
            for a in g.assignments:
                assignments.append({
                    "id": a["id"] if isinstance(a, dict) else getattr(a, "id", None),
                    "name": a["name"] if isinstance(a, dict) else getattr(a, "name", ""),
                    "group": group_name,
                    "due_at": a.get("due_at") if isinstance(a, dict) else getattr(a, "due_at", None)
                })
    except Exception as e:
        if verbose:
            print(f"[ChangeDeadlines] Failed to list assignments: {e}")
        else:
            print("Failed to list assignments.")
        return results

    if not assignments:
        if verbose:
            print("[ChangeDeadlines] No assignments found (after applying category filter).")
        else:
            print("No assignments found.")
        return results

    # If assignment_ids provided, normalize to list of assignment ids (strings/ints)
    selected_assignments = []
    if assignment_ids:
        if isinstance(assignment_ids, (str, int)):
            # single id or comma-separated
            s = str(assignment_ids)
            if "," in s:
                ids = [x.strip() for x in s.split(",") if x.strip()]
            else:
                ids = [s]
            for ident in ids:
                # accept numeric id or assignment name
                if ident.isdigit():
                    for a in assignments:
                        if str(a["id"]) == ident:
                            selected_assignments.append(a)
                            break
                else:
                    for a in assignments:
                        if ident.lower() in a["name"].lower():
                            selected_assignments.append(a)
        elif isinstance(assignment_ids, (list, tuple)):
            for ident in assignment_ids:
                sid = str(ident)
                for a in assignments:
                    if sid.isdigit() and str(a["id"]) == sid:
                        selected_assignments.append(a)
                        break
                    if isinstance(ident, str) and ident.lower() in a["name"].lower():
                        selected_assignments.append(a)
    else:
        # Interactive selection
        if verbose:
            print("[ChangeDeadlines] Listing assignments for selection:")
        else:
            print("Listing assignments:")
        for idx, a in enumerate(assignments, 1):
            due = a["due_at"] or "No due date"
            print(f"{idx}. [{a['group']}] {a['name']} (ID: {a['id']}, Due: {due})")
        sel = _prompt(
            "Enter assignment numbers to change (e.g. 1,3-5 or 'a' for all, or 'q' to quit) [q in 60s]: ",
            timeout=60,
            default="q"
        )
        if sel is None or sel.strip().lower() in ("q", "quit"):
            if verbose:
                print("[ChangeDeadlines] Operation cancelled due to timeout or user quit.")
            return results
        sel_indices = parse_selection(sel, len(assignments))
        if not sel_indices:
            print("No valid selection. Aborting.")
            return results
        for i in sel_indices:
            selected_assignments.append(assignments[i - 1])

    if not selected_assignments:
        print("No assignments selected. Aborting.")
        return results

    # Decide single deadline or different per assignment
    choice = None
    if new_due_date:
        # If new_due_date provided via argument, confirm with user to use single deadline
        confirmed = _prompt(
            f"Use provided date '{new_due_date}' for all selected assignments? (y/n, default y in 60s): ",
            timeout=60,
            default="y"
        )
        if confirmed is None or confirmed.strip().lower() in ("q", "quit"):
            print("Operation cancelled.")
            return results
        if confirmed.strip().lower() in ("y", "yes", ""):
            choice = "single"
        else:
            choice = "multiple"
    else:
        c = _prompt(
            "Apply one deadline to all selected assignments or specify different dates for each? (one/multiple, default one) [60s]: ",
            timeout=60,
            default="one"
        )
        if c is None:
            print("Operation cancelled.")
            return results
        c = c.strip().lower()
        if c in ("one", "single", "s", "1"):
            choice = "single"
        else:
            choice = "multiple"

    # If single: determine date
    single_iso = None
    if choice == "single":
        # Alert user about timezone expectation: treat input as GMT/UTC
        print("Note: Please enter dates in GMT (UTC) timezone. Examples: '2025-08-28 14:00 GMT' or '2025-08-28T14:00Z'.")
        if new_due_date:
            single_iso = normalize_datetime_str(new_due_date)
            if not single_iso:
                print(f"Could not parse provided new_due_date: {new_due_date}")
                return results
        else:
            date_input = _prompt(
                "Enter new due date/time for all selected assignments (format: YYYY-MM-DD HH:MM) [60s]: ",
                timeout=60,
                default=None
            )
            if date_input is None:
                print("Operation cancelled.")
                return results
            single_iso = normalize_datetime_str(date_input)
            if not single_iso:
                print("Could not parse the date you entered. Aborting.")
                return results

        # Apply to each selected assignment: first clear due_at (set to None), then set to new date
        for a in selected_assignments:
            aid = a["id"]
            try:
                # Clear existing due date first
                try:
                    assignment = course.get_assignment(aid)
                    # attempt clear via canvasapi
                    assignment.edit(due_at=None)
                except Exception:
                    # fallback via REST
                    url = f"{api_url}/api/v1/courses/{course_id}/assignments/{aid}"
                    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
                    payload_clear = {"assignment": {"due_at": None}}
                    try:
                        requests.put(url, headers=headers, json=payload_clear, timeout=15)
                    except Exception:
                        pass

                # Now set the new due date
                try:
                    assignment = course.get_assignment(aid)
                    assignment.edit(due_at=single_iso)
                    results[str(aid)] = "updated (via canvasapi)"
                except Exception:
                    url = f"{api_url}/api/v1/courses/{course_id}/assignments/{aid}"
                    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
                    payload = {"assignment": {"due_at": single_iso}}
                    resp = requests.put(url, headers=headers, json=payload)
                    if resp.status_code in (200, 201):
                        results[str(aid)] = "updated (via REST)"
                    else:
                        results[str(aid)] = f"failed ({resp.status_code})"
                if verbose:
                    print(f"[ChangeDeadlines] Assignment {aid} -> due_at set to {single_iso} (cleared first)")
            except Exception as e:
                results[str(aid)] = f"error: {e}"
                if verbose:
                    print(f"[ChangeDeadlines] Failed to update assignment {aid}: {e}")

        return results

    # If multiple: prompt for each assignment individually
    print("Note: Please enter dates in GMT (UTC) timezone when specifying per-assignment due dates.")
    for a in selected_assignments:
        aid = a["id"]
        prompt = f"Enter new due date for assignment '{a['name']}' (ID: {aid}) (format: YYYY-MM-DD HH:MM), or leave blank to skip: "
        date_input = _prompt(prompt, timeout=60, default=None)
        if date_input is None:
            print("Operation cancelled.")
            return results
        if not date_input.strip():
            results[str(aid)] = "skipped"
            continue
        # Parse user input and treat it as GMT (UTC) time.
        iso = None
        if date_input and date_input.strip():
            iso = normalize_datetime_str(date_input.strip())
        if not iso:
            results[str(aid)] = "invalid date format"
            continue
        try:
            # Clear existing due date first
            try:
                assignment = course.get_assignment(aid)
                assignment.edit(due_at=None)
            except Exception:
                url_clear = f"{api_url}/api/v1/courses/{course_id}/assignments/{aid}"
                headers_clear = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
                payload_clear = {"assignment": {"due_at": None}}
                try:
                    requests.put(url_clear, headers=headers_clear, json=payload_clear, timeout=15)
                except Exception:
                    pass

            # Now set the new due date
            try:
                assignment = course.get_assignment(aid)
                assignment.edit(due_at=iso)
                results[str(aid)] = "updated (via canvasapi)"
            except Exception:
                url = f"{api_url}/api/v1/courses/{course_id}/assignments/{aid}"
                headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
                payload = {"assignment": {"due_at": iso}}
                resp = requests.put(url, headers=headers, json=payload)
                if resp.status_code in (200, 201):
                    results[str(aid)] = "updated (via REST)"
                else:
                    results[str(aid)] = f"failed ({resp.status_code})"
            if verbose:
                print(f"[ChangeDeadlines] Assignment {aid} -> due_at set to {iso} (cleared first)")
        except Exception as e:
            results[str(aid)] = f"error: {e}"
            if verbose:
                print(f"[ChangeDeadlines] Failed to update assignment {aid}: {e}")

    return results

def create_canvas_groups(
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    group_set_id=None,
    num_groups=5,
    group_name_pattern=None,
    verbose=False
):
    """
    Create several groups with names like 'Group 1', 'Group 2', etc., in a specified group set.
    If group_set_id is None, list all group sets and allow user to select one.
    If no group sets are available, ask user to create one.
    Always allow user to quit at any step.
    If no response from user within 60 seconds, quit.
    Skip creating a group if a group with the same name already exists in the group set.

    Args:
        api_url (str): Canvas API base URL.
        api_key (str): Canvas API key.
        course_id (str or int): Canvas course ID.
        group_set_id (str or int or None): ID of the group set to create groups in. If None, prompt user.
        num_groups (int): Number of groups to create (default: 5).
        group_name_pattern (str or None): Pattern for group names, e.g., "Group {i}". If None, prompt user.
        verbose (bool): Print more details.

    Returns:
        List of created group dicts (from API response) or None if cancelled.
    """
    def timeout_handler(_signum, _frame):
        if verbose:
            print("\n[CreateGroups] Timeout: No response after 60 seconds. Quitting...")
        else:
            print("\nTimeout: No response after 60 seconds. Quitting...")
        raise TimeoutError("User input timeout")

    def _prompt(prompt, timeout=60, default=None):
        return prompt_with_timeout(prompt, timeout=timeout, default=default, verbose=verbose)

    try:
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)
        if verbose:
            print(f"[CreateGroups] Connected to Canvas course {course_id}.")
        else:
            print(f"Connected to Canvas course {course_id}.")

        # Step 1: Get or select group set
        if group_set_id is None:
            group_sets = list(course.get_group_categories())
            if not group_sets:
                if verbose:
                    print("[CreateGroups] No group sets found. Creating a new one...")
                else:
                    print("No group sets found. Creating a new one...")
                name = _prompt(
                    "Enter name for new group set (or 'q' to quit): ",
                    timeout=60,
                    default="Default Group Set"
                )
                if name is None or name.strip().lower() in ("q", "quit"):
                    if verbose:
                        print("[CreateGroups] Quitting.")
                    else:
                        print("Quitting.")
                    return None
                group_set = course.create_group_category(name=name.strip())
                group_set_id = group_set.id
                if verbose:
                    print(f"[CreateGroups] Created group set '{name}' with ID {group_set_id}.")
                else:
                    print(f"Created group set '{name}' with ID {group_set_id}.")
            else:
                if verbose:
                    print("[CreateGroups] Available group sets:")
                else:
                    print("Available group sets:")
                for idx, gs in enumerate(group_sets, 1):
                    if verbose:
                        print(f"[CreateGroups] {idx}. {gs.name} (ID: {gs.id})")
                    else:
                        print(f"{idx}. {gs.name} (ID: {gs.id})")
                sel = _prompt(
                    "Select group set number (or 'q' to quit): ",
                    timeout=60,
                    default="q"
                )
                if sel is None or sel.strip().lower() in ("q", "quit"):
                    if verbose:
                        print("[CreateGroups] Quitting.")
                    else:
                        print("Quitting.")
                    return None
                try:
                    idx = int(sel) - 1
                    if 0 <= idx < len(group_sets):
                        group_set_id = group_sets[idx].id
                        if verbose:
                            print(f"[CreateGroups] Selected group set: {group_sets[idx].name}")
                        else:
                            print(f"Selected group set: {group_sets[idx].name}")
                    else:
                        if verbose:
                            print("[CreateGroups] Invalid selection.")
                        else:
                            print("Invalid selection.")
                        return None
                except ValueError:
                    if verbose:
                        print("[CreateGroups] Invalid input.")
                    else:
                        print("Invalid input.")
                    return None

        # Fetch existing groups in the selected group set (handle pagination)
        headers = {"Authorization": f"Bearer {api_key}"}
        groups_url = f"{api_url}/api/v1/group_categories/{group_set_id}/groups"
        existing_groups = []
        page = 1
        per_page = 100  # Fetch up to 100 groups per page to minimize requests
        try:
            while True:
                paginated_url = f"{groups_url}?page={page}&per_page={per_page}"
                response = requests.get(paginated_url, headers=headers)
                if response.status_code != 200:
                    if verbose:
                        print(f"[CreateGroups] Failed to fetch page {page} of existing groups: {response.status_code}")
                    break
                data = response.json()
                if not data:
                    break
                existing_groups.extend(data)
                page += 1
            existing_names = {group['name'] for group in existing_groups}
            if verbose:
                print(f"[CreateGroups] Found {len(existing_groups)} existing groups in group set {group_set_id}.")
        except Exception as e:
            if verbose:
                print(f"[CreateGroups] Error fetching existing groups: {e}")
            else:
                print("Error fetching existing groups.")
            existing_names = set()

        # Step 2: Confirm number of groups
        num_str = _prompt(
            f"Enter number of groups to create (default: {num_groups}, or 'q' to quit): ",
            timeout=60,
            default=str(num_groups)
        )
        if num_str is None or num_str.strip().lower() in ("q", "quit"):
            if verbose:
                print("[CreateGroups] Quitting.")
            else:
                print("Quitting.")
            return None
        try:
            num_groups = int(num_str.strip())
            if num_groups <= 0:
                if verbose:
                    print("[CreateGroups] Number of groups must be positive.")
                else:
                    print("Number of groups must be positive.")
                return None
        except ValueError:
            if verbose:
                print("[CreateGroups] Invalid number.")
            else:
                print("Invalid number.")
            return None

        # Step 3: Get group name pattern
        if group_name_pattern is None:
            pattern = _prompt(
                "Enter group name pattern (e.g., 'Group {i}', default: 'Group {i}', or 'q' to quit): ",
                timeout=60,
                default="Group {i}"
            )
            if pattern is None or pattern.strip().lower() in ("q", "quit"):
                if verbose:
                    print("[CreateGroups] Quitting.")
                else:
                    print("Quitting.")
                return None
            group_name_pattern = pattern.strip()

        # Step 4: Create groups using REST API, skipping if name exists
        created_groups = []
        skipped_groups = []
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        for i in range(1, num_groups + 1):
            group_name = group_name_pattern.format(i=i)
            if group_name in existing_names:
                skipped_groups.append(group_name)
                if verbose:
                    print(f"[CreateGroups] Skipped creating group '{group_name}' (already exists).")
                else:
                    print(f"Skipped creating group '{group_name}' (already exists).")
                continue
            url = f"{api_url}/api/v1/group_categories/{group_set_id}/groups"
            data = {"name": group_name}
            try:
                response = requests.post(url, headers=headers, json=data)
                if response.status_code in (200, 201):
                    group = response.json()
                    created_groups.append(group)
                    if verbose:
                        print(f"[CreateGroups] Created group '{group_name}' in group set {group_set_id}.")
                    else:
                        print(f"Created group '{group_name}'.")
                else:
                    if verbose:
                        print(f"[CreateGroups] Failed to create group '{group_name}': {response.status_code} - {response.text}")
                    else:
                        print(f"Failed to create group '{group_name}': {response.status_code}")
            except Exception as e:
                if verbose:
                    print(f"[CreateGroups] Failed to create group '{group_name}': {e}")
                else:
                    print(f"Failed to create group '{group_name}': {e}")
                # Continue creating others

        if skipped_groups:
            if verbose:
                print(f"[CreateGroups] Skipped {len(skipped_groups)} groups that already exist: {', '.join(skipped_groups)}")
            else:
                print(f"Skipped {len(skipped_groups)} groups that already exist.")

        if verbose:
            print(f"[CreateGroups] Successfully created {len(created_groups)} groups.")
        else:
            print(f"Successfully created {len(created_groups)} groups.")
        return created_groups

    except Exception as e:
        if verbose:
            print(f"[CreateGroups] Error: {e}")
        else:
            print(f"Error: {e}")
        return None

def change_canvas_lock_dates(
    assignment_ids=None,
    new_lock_date=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    verbose=False
):
    """
    Change lock dates (lock_at) for one or more Canvas assignments.
    - If assignment_ids is None, list assignments (optionally filtered by category) and allow interactive selection (supports ranges, comma-separated, 'a' for all).
    - If new_lock_date is provided, use it as the single lock date for all selected assignments.
    - If new_lock_date is None, prompt the user for a date; if no response within 60 seconds, default to the assignment's due_at + 4 days (if due_at exists).
    - Supports setting a single date for all or different dates per assignment.
    - Prompts time out after 60 seconds and the operation quits (default).
    - If verbose is True, print more details; otherwise, print only important notice.
    Returns a dict mapping assignment_id -> status string.
    Note: this implementation assumes the local time provided by the user is in the same timezone as Canvas (no timezone conversions are performed).
    """
    def timeout_handler(_signum, _frame):
        print("\nTimeout: No response after 60 seconds. Quitting...")
        raise TimeoutError("User input timeout")

    def _prompt(prompt, timeout=60, default=None):
        return prompt_with_timeout(prompt, timeout=timeout, default=default, verbose=verbose)

    def parse_selection(sel, n):
        sel = sel.strip().lower()
        if sel in ("a", "all"):
            return list(range(1, n + 1))
        parts = set()
        for part in sel.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                try:
                    start, end = map(int, part.split("-", 1))
                    parts.update(range(start, end + 1))
                except Exception:
                    continue
            else:
                if part.isdigit():
                    parts.add(int(part))
        return sorted(i for i in parts if 1 <= i <= n)

    def normalize_datetime_str(s):
        # Accept "YYYY-MM-DD HH:MM" or "YYYY-MM-DDTHH:MM" or "YYYY-MM-DD"
        # Also accept trailing GMT/UTC or trailing 'Z'. Treat inputs as GMT (UTC).
        if s is None:
            return None
        s = s.strip()
        if not s:
            return None
        # Remove trailing "GMT" or "UTC" (case-insensitive) and normalize trailing Z
        s_clean = re.sub(r'\s*(?:gmt|utc)$', '', s, flags=re.IGNORECASE)
        # If ends with Z keep it (ISO)
        try_formats = ("%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M", "%Y-%m-%d")
        for fmt in try_formats:
            try:
                dt = datetime.strptime(s_clean, fmt)
                # Treat user-provided time as GMT (UTC) and output in Canvas-friendly Z format
                if fmt == "%Y-%m-%d":
                    return dt.strftime("%Y-%m-%dT00:00:00Z")
                return dt.strftime("%Y-%m-%dT%H:%M:%SZ")
            except Exception:
                continue
        # Try more flexible parsing (ISO)
        try:
            # fromisoformat doesn't accept trailing 'Z', so replace if present
            iso_candidate = s_clean
            if iso_candidate.endswith("Z"):
                iso_candidate = iso_candidate[:-1]
            dt = datetime.fromisoformat(iso_candidate)
            return dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        except Exception:
            return None

    results = {}

    try:
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)
    except Exception as e:
        if verbose:
            print(f"[ChangeLock] Failed to connect to Canvas: {e}")
        else:
            print("Failed to connect to Canvas.")
        return results

    # Gather assignments (filtered by category if provided)
    try:
        groups = list(course.get_assignment_groups(include=["assignments"]))
        assignments = []
        for g in groups:
            group_name = g.name
            if category and group_name.lower() != category.lower():
                continue
            for a in g.assignments:
                assignments.append({
                    "id": a["id"] if isinstance(a, dict) else getattr(a, "id", None),
                    "name": a["name"] if isinstance(a, dict) else getattr(a, "name", ""),
                    "group": group_name,
                    "due_at": a.get("due_at") if isinstance(a, dict) else getattr(a, "due_at", None)
                })
    except Exception as e:
        if verbose:
            print(f"[ChangeLock] Failed to list assignments: {e}")
        else:
            print("Failed to list assignments.")
        return results

    if not assignments:
        if verbose:
            print("[ChangeLock] No assignments found (after applying category filter).")
        else:
            print("No assignments found.")
        return results

    # If assignment_ids provided, normalize to list of assignment ids (strings/ints)
    selected_assignments = []
    if assignment_ids:
        if isinstance(assignment_ids, (str, int)):
            # single id or comma-separated
            s = str(assignment_ids)
            if "," in s:
                ids = [x.strip() for x in s.split(",") if x.strip()]
            else:
                ids = [s]
            for ident in ids:
                # accept numeric id or assignment name
                if ident.isdigit():
                    for a in assignments:
                        if str(a["id"]) == ident:
                            selected_assignments.append(a)
                            break
                else:
                    for a in assignments:
                        if ident.lower() in a["name"].lower():
                            selected_assignments.append(a)
        elif isinstance(assignment_ids, (list, tuple)):
            for ident in assignment_ids:
                sid = str(ident)
                for a in assignments:
                    if sid.isdigit() and str(a["id"]) == sid:
                        selected_assignments.append(a)
                        break
                    if isinstance(ident, str) and ident.lower() in a["name"].lower():
                        selected_assignments.append(a)
    else:
        # Interactive selection
        if verbose:
            print("[ChangeLock] Listing assignments for selection:")
        else:
            print("Listing assignments:")
        for idx, a in enumerate(assignments, 1):
            due = a["due_at"] or "No due date"
            print(f"{idx}. [{a['group']}] {a['name']} (ID: {a['id']}, Due: {due})")
        sel = _prompt(
            "Enter assignment numbers to change (e.g. 1,3-5 or 'a' for all, or 'q' to quit) [q in 60s]: ",
            timeout=60,
            default="q"
        )
        if sel is None or sel.strip().lower() in ("q", "quit"):
            if verbose:
                print("[ChangeLock] Operation cancelled due to timeout or user quit.")
            return results
        sel_indices = parse_selection(sel, len(assignments))
        if not sel_indices:
            print("No valid selection. Aborting.")
            return results
        for i in sel_indices:
            selected_assignments.append(assignments[i - 1])

    if not selected_assignments:
        print("No assignments selected. Aborting.")
        return results

    # Decide single lock date or different per assignment
    choice = None
    if new_lock_date:
        # If new_lock_date provided via argument, confirm with user to use single lock date
        confirmed = _prompt(
            f"Use provided date '{new_lock_date}' for all selected assignments? (y/n, default y in 60s): ",
            timeout=60,
            default="y"
        )
        if confirmed is None or confirmed.strip().lower() in ("q", "quit"):
            print("Operation cancelled.")
            return results
        if confirmed.strip().lower() in ("y", "yes", ""):
            choice = "single"
        else:
            choice = "multiple"
    else:
        c = _prompt(
            "Apply one lock date to all selected assignments or specify different dates for each? (one/multiple, default one) [60s]: ",
            timeout=60,
            default="one"
        )
        if c is None:
            print("Operation cancelled.")
            return results
        c = c.strip().lower()
        if c in ("one", "single", "s", "1"):
            choice = "single"
        else:
            choice = "multiple"

    # If single: determine date
    single_iso = None
    if choice == "single":
        # Alert user about timezone expectation: treat input as GMT/UTC
        print("Note: Please enter dates in GMT (UTC) timezone. Examples: '2025-08-28 14:00 GMT' or '2025-08-28T14:00Z'.")
        if new_lock_date:
            single_iso = normalize_datetime_str(new_lock_date)
            if not single_iso:
                print(f"Could not parse provided new_lock_date: {new_lock_date}")
                return results
        else:
            date_input = _prompt(
                "Enter new lock date/time for all selected assignments (format: YYYY-MM-DD HH:MM) [60s]: ",
                timeout=60,
                default=None
            )
            if date_input is None:
                print("Operation cancelled.")
                return results
            single_iso = normalize_datetime_str(date_input)
            if not single_iso:
                print("Could not parse the date you entered. Aborting.")
                return results

        # Apply to each selected assignment: set lock_at to the date
        for a in selected_assignments:
            aid = a["id"]
            try:
                # Set the new lock date
                try:
                    assignment = course.get_assignment(aid)
                    assignment.edit(lock_at=single_iso)
                    results[str(aid)] = "updated (via canvasapi)"
                except Exception:
                    url = f"{api_url}/api/v1/courses/{course_id}/assignments/{aid}"
                    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
                    payload = {"assignment": {"lock_at": single_iso}}
                    resp = requests.put(url, headers=headers, json=payload)
                    if resp.status_code in (200, 201):
                        results[str(aid)] = "updated (via REST)"
                    else:
                        results[str(aid)] = f"failed ({resp.status_code})"
                if verbose:
                    print(f"[ChangeLock] Assignment {aid} -> lock_at set to {single_iso}")
            except Exception as e:
                results[str(aid)] = f"error: {e}"
                if verbose:
                    print(f"[ChangeLock] Failed to update assignment {aid}: {e}")

        return results

    # If multiple: prompt for each assignment individually
    print("Note: Please enter dates in GMT (UTC) timezone when specifying per-assignment lock dates.")
    for a in selected_assignments:
        aid = a["id"]
        due_at = a["due_at"]
        default_date = None
        if due_at:
            try:
                due_dt = datetime.strptime(due_at, "%Y-%m-%dT%H:%M:%SZ")
                default_dt = due_dt + timedelta(days=4)
                default_date = default_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
            except Exception:
                pass
        prompt = f"Enter new lock date for assignment '{a['name']}' (ID: {aid}) (format: YYYY-MM-DD HH:MM), or leave blank to use default (due_at + 4 days: {default_date or 'N/A'}) [60s]: "
        date_input = _prompt(prompt, timeout=60, default=None)
        if date_input is None:
            print("Operation cancelled.")
            return results
        if not date_input.strip():
            if default_date:
                iso = default_date
            else:
                results[str(aid)] = "skipped (no due_at to calculate default)"
                continue
        else:
            # Parse user input and treat it as GMT (UTC) time.
            iso = normalize_datetime_str(date_input.strip())
            if not iso:
                results[str(aid)] = "invalid date format"
                continue
        try:
            # Set the new lock date
            try:
                assignment = course.get_assignment(aid)
                assignment.edit(lock_at=iso)
                results[str(aid)] = "updated (via canvasapi)"
            except Exception:
                url = f"{api_url}/api/v1/courses/{course_id}/assignments/{aid}"
                headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
                payload = {"assignment": {"lock_at": iso}}
                resp = requests.put(url, headers=headers, json=payload)
                if resp.status_code in (200, 201):
                    results[str(aid)] = "updated (via REST)"
                else:
                    results[str(aid)] = f"failed ({resp.status_code})"
            if verbose:
                print(f"[ChangeLock] Assignment {aid} -> lock_at set to {iso}")
        except Exception as e:
            results[str(aid)] = f"error: {e}"
            if verbose:
                print(f"[ChangeLock] Failed to update assignment {aid}: {e}")

    return results

def delete_empty_canvas_groups(
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    group_set_id=None,
    verbose=False
):
    """
    Delete all empty groups (groups with no members) from a Canvas course group set.
    If group_set_id is None, list all group sets and allow user to select one.
    Always allow user to quit at any step.
    If no response from user within 60 seconds, quit.

    Args:
        api_url (str): Canvas API base URL.
        api_key (str): Canvas API key.
        course_id (str or int): Canvas course ID.
        group_set_id (str or int or None): ID of the group set to delete empty groups from. If None, prompt user.
        verbose (bool): Print more details.

    Returns:
        int: Number of empty groups deleted.
    """
    def timeout_handler(_signum, _frame):
        if verbose:
            print("\n[DeleteGroups] Timeout: No response after 60 seconds. Quitting...")
        else:
            print("\nTimeout: No response after 60 seconds. Quitting...")
        raise TimeoutError("User input timeout")

    def _prompt(prompt, timeout=60, default=None):
        return prompt_with_timeout(prompt, timeout=timeout, default=default, verbose=verbose)

    try:
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)
        if verbose:
            print(f"[DeleteGroups] Connected to Canvas course {course_id}.")
        else:
            print(f"Connected to Canvas course {course_id}.")

        # Step 1: Get or select group set
        if group_set_id is None:
            group_sets = list(course.get_group_categories())
            if not group_sets:
                if verbose:
                    print("[DeleteGroups] No group sets found in this course.")
                else:
                    print("No group sets found in this course.")
                return 0
            if verbose:
                print("[DeleteGroups] Available group sets:")
            else:
                print("Available group sets:")
            for idx, gs in enumerate(group_sets, 1):
                if verbose:
                    print(f"[DeleteGroups] {idx}. {gs.name} (ID: {gs.id})")
                else:
                    print(f"{idx}. {gs.name} (ID: {gs.id})")
            sel = _prompt(
                "Select group set number (or 'q' to quit): ",
                timeout=60,
                default="q"
            )
            if sel is None or sel.strip().lower() in ("q", "quit"):
                if verbose:
                    print("[DeleteGroups] Quitting.")
                else:
                    print("Quitting.")
                return 0
            try:
                idx = int(sel) - 1
                if 0 <= idx < len(group_sets):
                    group_set_id = group_sets[idx].id
                    if verbose:
                        print(f"[DeleteGroups] Selected group set: {group_sets[idx].name}")
                    else:
                        print(f"Selected group set: {group_sets[idx].name}")
                else:
                    if verbose:
                        print("[DeleteGroups] Invalid selection.")
                    else:
                        print("Invalid selection.")
                    return 0
            except ValueError:
                if verbose:
                    print("[DeleteGroups] Invalid input.")
                else:
                    print("Invalid input.")
                return 0

        # Fetch all groups in the selected group set (handle pagination)
        headers = {"Authorization": f"Bearer {api_key}"}
        groups_url = f"{api_url}/api/v1/group_categories/{group_set_id}/groups"
        all_groups = []
        page = 1
        per_page = 100
        try:
            while True:
                paginated_url = f"{groups_url}?page={page}&per_page={per_page}"
                response = requests.get(paginated_url, headers=headers)
                if response.status_code != 200:
                    if verbose:
                        print(f"[DeleteGroups] Failed to fetch page {page} of groups: {response.status_code}")
                    break
                data = response.json()
                if not data:
                    break
                all_groups.extend(data)
                page += 1
            if verbose:
                print(f"[DeleteGroups] Found {len(all_groups)} groups in group set {group_set_id}.")
        except Exception as e:
            if verbose:
                print(f"[DeleteGroups] Error fetching groups: {e}")
            else:
                print("Error fetching groups.")
            return 0

        # Find empty groups (no members)
        empty_groups = []
        for group in all_groups:
            group_id = group['id']
            members_url = f"{api_url}/api/v1/groups/{group_id}/users"
            try:
                members_response = requests.get(members_url, headers=headers)
                if members_response.status_code == 200:
                    members = members_response.json()
                    if not members:
                        empty_groups.append(group)
            except Exception as e:
                if verbose:
                    print(f"[DeleteGroups] Error fetching members for group {group_id}: {e}")

        if not empty_groups:
            if verbose:
                print("[DeleteGroups] No empty groups found.")
            else:
                print("No empty groups found.")
            return 0

        if verbose:
            print(f"[DeleteGroups] Found {len(empty_groups)} empty groups:")
            for g in empty_groups:
                print(f"[DeleteGroups]   - {g['name']} (ID: {g['id']})")
        else:
            print(f"Found {len(empty_groups)} empty groups:")
            for g in empty_groups:
                print(f"  - {g['name']} (ID: {g['id']})")

        # Confirm deletion
        confirm = _prompt(
            f"Delete all {len(empty_groups)} empty groups? (y/n, default 'n' in 60s): ",
            timeout=60,
            default="n"
        ).strip().lower()
        if confirm not in ("y", "yes"):
            if verbose:
                print("[DeleteGroups] Deletion cancelled.")
            else:
                print("Deletion cancelled.")
            return 0

        # Delete empty groups using REST API
        deleted_count = 0
        for group in empty_groups:
            group_id = group['id']
            delete_url = f"{api_url}/api/v1/groups/{group_id}"
            try:
                response = requests.delete(delete_url, headers=headers)
                if response.status_code in (200, 204):
                    deleted_count += 1
                    if verbose:
                        print(f"[DeleteGroups] Deleted group '{group['name']}' (ID: {group_id}).")
                    else:
                        print(f"Deleted group '{group['name']}'.")
                else:
                    if verbose:
                        print(f"[DeleteGroups] Failed to delete group '{group['name']}': {response.status_code}")
                    else:
                        print(f"Failed to delete group '{group['name']}': {response.status_code}")
            except Exception as e:
                if verbose:
                    print(f"[DeleteGroups] Failed to delete group '{group['name']}': {e}")
                else:
                    print(f"Failed to delete group '{group['name']}': {e}")

        if verbose:
            print(f"[DeleteGroups] Successfully deleted {deleted_count} empty groups.")
        else:
            print(f"Successfully deleted {deleted_count} empty groups.")
        return deleted_count

    except Exception as e:
        if verbose:
            print(f"[DeleteGroups] Error: {e}")
        else:
            print(f"Error: {e}")
        return 0
