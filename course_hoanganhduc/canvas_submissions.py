# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas submission helpers."""

import os
import signal
from datetime import datetime, timezone

import requests
from .canvas_auth import get_canvas_client
from tqdm import tqdm

from .canvas_checks import compare_texts_from_pdfs_in_folder, detect_meaningful_level_and_notify_students
from .data import load_database
from .canvas_utils import prompt_with_timeout, parse_selection, normalize_datetime_str
from .settings import (
    ALL_AI_METHODS,
    CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    CANVAS_LMS_API_KEY,
    CANVAS_LMS_API_URL,
    CANVAS_LMS_COURSE_ID,
    DEFAULT_OCR_METHOD,
    DEFAULT_AI_METHOD,
)
from .config import DEFAULT_DOWNLOAD_FOLDER

def download_canvas_assignment_submissions(
    assignment_id=None,
    dest_dir=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    verbose=False
):
    """
    Download only the latest version of each student's submission file for one or more Canvas assignments.
    If an old version is already downloaded, delete the old version and keep the new one.
    If assignment_id is not specified, list all assignments (optionally only in one category) that have at least one submission and prompt user to select one or more (supports ranges, comma-separated, or 'a' for all).
    Files are renamed as: <student name>_<canvas id>_<assignment id>_<submitted time>_<status>.<ext>
    All files are saved in a folder named <assignment group>_<assignment_name>.
    After downloading, check PDF content similarity and meaningful level in the downloaded folder.
    At every user prompt, allow quitting by entering 'q' or 'quit'.
    Added options:
      - Download all latest submissions (delete old versions, skip already downloaded latest)
      - Perform all tasks (download, check meaningfulness, check similarity)
      - These options are used as default if user does not respond in 60 seconds.
    """

    def timeout_handler(*_):
        if verbose:
            print("\n[DownloadCanvas] Timeout: No response after 60 seconds. Using default option.")
        else:
            print("\nTimeout: No response after 60 seconds. Using default option.")
        raise TimeoutError("User input timeout")

    def _prompt(prompt, timeout=60, default=None):
        return prompt_with_timeout(prompt, timeout=timeout, default=default, verbose=verbose)

    try:
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)

        # If assignment_id is not specified, list all assignments with at least one submission and prompt user
        assignment_ids = []
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
            msg = "No assignments found with submissions." if category else "No assignments found."
            if verbose:
                print(f"[DownloadCanvas] {msg}")
            else:
                print(msg)
            return

        if not assignment_id:
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
                print("[DownloadCanvas] Assignments with at least one submission:")
            else:
                print("Assignments with at least one submission:")
            for idx, a in enumerate(assignments, 1):
                due = a['due_at'] or "No due date"
                print(f"{idx}. [{a['group']}] {a['name']} (ID: {a['id']}, Due: {due})")
            while True:
                sel = _prompt(
                    "Enter the number(s) of the assignment(s) to download submissions from (e.g. 1,3-5 or 'a' for all, or 'q' to quit): ",
                    timeout=60,
                    default="a"
                ).strip()
                if sel.lower() in ('q', 'quit'):
                    if verbose:
                        print("[DownloadCanvas] Quitting download.")
                    else:
                        print("Quitting download.")
                    return
                if sel.lower() in ('a', 'all'):
                    assignment_ids = [a['id'] for a in assignments]
                    break
                # Parse comma-separated and range input
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
                    if verbose:
                        print("[DownloadCanvas] Invalid selection. Please enter valid numbers, a range, 'a' for all, or 'q' to quit.")
                    else:
                        print("Invalid selection. Please enter valid numbers, a range, 'a' for all, or 'q' to quit.")
                    continue
                assignment_ids = [assignments[i - 1]['id'] for i in selected]
                break
        else:
            assignment_ids = [assignment_id]

        for assignment_id in assignment_ids:
            assignment = course.get_assignment(assignment_id)
            assignment_name = assignment.name.replace("/", "_").replace("\\", "_")
            # Get assignment group name
            group_name = "UnknownGroup"
            try:
                group_id = assignment.assignment_group_id
                for ag in course.get_assignment_groups():
                    if ag.id == group_id:
                        group_name = ag.name.replace("/", "_").replace("\\", "_")
                        break
            except Exception:
                pass
            folder_name = f"{group_name}_{assignment_name}".replace(" ", "_")
            if dest_dir:
                out_dir = os.path.join(dest_dir, folder_name)
            else:
                out_dir = os.path.join(DEFAULT_DOWNLOAD_FOLDER, folder_name)
            os.makedirs(out_dir, exist_ok=True)
            if verbose:
                print(f"[DownloadCanvas] Downloading submissions to: {out_dir}")
            else:
                print(f"Downloading submissions to: {out_dir}")

            # Use include=["user", "attachments"] for fast batch fetch
            submissions = list(assignment.get_submissions(include=["user", "attachments"]))
            count = 0
            skipped = 0

            # Precompute due date for all submissions
            due_at = getattr(assignment, "due_at", None)
            if due_at:
                try:
                    due_dt = datetime.strptime(due_at, "%Y-%m-%dT%H:%M:%SZ")
                except Exception:
                    due_dt = None
            else:
                due_dt = None

            # Track latest submission info per student (canvas_id)
            latest_submissions = {}
            for sub in submissions:
                user = getattr(sub, "user", {})
                canvas_id = user.get("id", "unknown")
                submitted_at = getattr(sub, "submitted_at", None)
                if not submitted_at:
                    continue
                try:
                    sub_dt = datetime.strptime(submitted_at, "%Y-%m-%dT%H:%M:%SZ")
                except Exception:
                    sub_dt = None
                # Only keep the latest submission for each student
                if canvas_id not in latest_submissions or (sub_dt and latest_submissions[canvas_id]["sub_dt"] and sub_dt > latest_submissions[canvas_id]["sub_dt"]):
                    latest_submissions[canvas_id] = {
                        "submission": sub,
                        "sub_dt": sub_dt,
                        "submitted_at": submitted_at
                    }

            # Gather all existing files in the output directory
            existing_files = set(os.listdir(out_dir))

            # Ask user if they want to download all latest submissions (delete old, skip already downloaded latest)
            try:
                download_all_choice = _prompt(
                    f"\nDo you want to download all latest submissions for '{assignment_name}' (delete old versions, skip already downloaded latest)? (y/n, default 'y' in 60s): ",
                    timeout=60,
                    default="y"
                ).strip().lower()
            except TimeoutError:
                download_all_choice = "y"
            if download_all_choice in ("q", "quit"):
                if verbose:
                    print("[DownloadCanvas] Quitting download.")
                else:
                    print("Quitting download.")
                return
            download_all = (download_all_choice in ("y", "yes", ""))

            if download_all:
                # For each latest submission, download the latest file and remove old versions
                for canvas_id, info in tqdm(latest_submissions.items(), desc=f"Downloading latest submissions for {assignment_name}"):
                    sub = info["submission"]
                    user = getattr(sub, "user", {})
                    student_name = user.get("name", "UnknownStudent").replace("/", "_").replace("\\", "_").replace(" ", "_")
                    submitted_at = info["submitted_at"]
                    if submitted_at:
                        try:
                            dt = datetime.strptime(submitted_at, "%Y%m%d_%H%M%S")
                        except Exception:
                            try:
                                dt = datetime.strptime(submitted_at, "%Y-%m-%dT%H:%M:%SZ")
                                dt_str = dt.strftime("%Y%m%d_%H%M%S")
                            except Exception:
                                dt_str = "no_time"
                        else:
                            dt_str = dt.strftime("%Y%m%d_%H%M%S")
                    else:
                        dt_str = "no_time"
                    # Determine on time or late
                    status = "on_time"
                    if due_dt and submitted_at:
                        try:
                            sub_dt = datetime.strptime(submitted_at, "%Y-%m-%dT%H:%M:%SZ")
                            status = "on_time" if sub_dt <= due_dt else "late"
                        except Exception:
                            status = "on_time"
                    elif getattr(sub, "late", False):
                        status = "late"
                    attachments = getattr(sub, "attachments", [])
                    if not attachments:
                        continue
                    for att in attachments:
                        url = getattr(att, "url", None)
                        filename = getattr(att, "filename", "file")
                        ext = os.path.splitext(filename)[1]
                        # Add assignment_id to the filename
                        new_filename = f"{student_name}_{canvas_id}_{assignment_id}_{dt_str}_{status}{ext}"
                        out_path = os.path.join(out_dir, new_filename)
                        # Remove all old versions for this student (same canvas_id, same assignment_id, same ext, but different dt_str)
                        for f in list(os.listdir(out_dir)):
                            if (
                                f.startswith(f"{student_name}_{str(canvas_id)}_{str(assignment_id)}_")
                                and f.endswith(ext)
                                and f != new_filename
                            ):
                                try:
                                    os.remove(os.path.join(out_dir, f))
                                except Exception:
                                    pass
                        # Download if not exists or file is empty
                        if os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                            skipped += 1
                            continue
                        try:
                            resp = requests.get(url, headers={"Authorization": f"Bearer {api_key}"}, stream=False, timeout=60)
                            resp.raise_for_status()
                            with open(out_path, "wb") as f:
                                f.write(resp.content)
                            count += 1
                        except Exception as e:
                            if verbose:
                                print(f"[DownloadCanvas] Failed to download {filename} for {student_name}: {e}")
                            else:
                                print(f"Failed to download {filename} for {student_name}: {e}")
                if verbose:
                    print(f"[DownloadCanvas] Downloaded {count} files to {out_dir} (skipped {skipped} already existing latest files)")
                else:
                    print(f"Downloaded {count} files to {out_dir} (skipped {skipped} already existing latest files)")
            else:
                if verbose:
                    print("[DownloadCanvas] Skipped downloading all latest submissions.")
                else:
                    print("Skipped downloading all latest submissions.")

            # After downloading, check PDF content similarity and meaningful level in the downloaded folder
            pdf_files = [os.path.join(out_dir, f) for f in os.listdir(out_dir) if f.lower().endswith(".pdf")]
            if len(pdf_files) >= 1:
                if verbose:
                    print("[DownloadCanvas] Downloaded PDFs found. You can now check PDF content quality and/or similarity.")
                else:
                    print("Downloaded PDFs found. You can now check PDF content quality and/or similarity.")

                # Ask user what to do next, default is "3" (perform all tasks)
                try:
                    print("\nWhat would you like to do?")
                    print("1. Check PDF meaningful level (quality)")
                    print("2. Check PDF content similarity")
                    print("3. Check both meaningful level and similarity")
                    print("4. Skip checks and exit")
                    choice = _prompt(
                        "Enter your choice (1-4, or 'q' to quit): ",
                        timeout=60,
                        default="3"
                    ).strip()
                except TimeoutError:
                    choice = "3"

                if choice.lower() in ("q", "quit"):
                    if verbose:
                        print("[DownloadCanvas] Quitting after download.")
                    else:
                        print("Quitting after download.")
                    return
                elif choice == "1":
                    if verbose:
                        print("[DownloadCanvas] Checking meaningful level of PDF submissions...")
                    else:
                        print("Checking meaningful level of PDF submissions...")
                    detect_meaningful_level_and_notify_students(
                        out_dir,
                        assignment_id=assignment_id,
                        api_url=api_url,
                        api_key=api_key,
                        course_id=course_id,
                        verbose=verbose,
                        ocr_service=DEFAULT_OCR_METHOD
                    )
                elif choice == "2":
                    if len(pdf_files) >= 2:
                        if verbose:
                            print("[DownloadCanvas] Checking PDF content similarity...")
                        else:
                            print("Checking PDF content similarity...")
                        compare_texts_from_pdfs_in_folder(
                            out_dir,
                            api_url=api_url,
                            api_key=api_key,
                            course_id=course_id,
                            ocr_service=DEFAULT_OCR_METHOD
                        )
                    else:
                        if verbose:
                            print("[DownloadCanvas] Not enough PDF files to compare similarity (need at least 2).")
                        else:
                            print("Not enough PDF files to compare similarity (need at least 2).")
                elif choice == "3":
                    if verbose:
                        print("[DownloadCanvas] Checking meaningful level of PDF submissions...")
                    else:
                        print("Checking meaningful level of PDF submissions...")
                    detect_meaningful_level_and_notify_students(
                        out_dir,
                        assignment_id=assignment_id,
                        api_url=api_url,
                        api_key=api_key,
                        course_id=course_id,
                        verbose=verbose,
                        ocr_service=DEFAULT_OCR_METHOD
                    )
                    if len(pdf_files) >= 2:
                        if verbose:
                            print("[DownloadCanvas] Checking PDF content similarity...")
                        else:
                            print("\nChecking PDF content similarity...")
                        compare_texts_from_pdfs_in_folder(
                            out_dir,
                            api_url=api_url,
                            api_key=api_key,
                            course_id=course_id,
                            ocr_service=DEFAULT_OCR_METHOD
                        )
                    else:
                        if verbose:
                            print("[DownloadCanvas] Not enough PDF files to compare similarity (need at least 2).")
                        else:
                            print("Not enough PDF files to compare similarity (need at least 2).")
                elif choice == "4":
                    if verbose:
                        print("[DownloadCanvas] Skipped PDF quality and similarity checks.")
                    else:
                        print("Skipped PDF quality and similarity checks.")
                else:
                    if verbose:
                        print("[DownloadCanvas] Invalid choice. Please enter 1, 2, 3, 4, or 'q' to quit.")
                    else:
                        print("Invalid choice. Please enter 1, 2, 3, 4, or 'q' to quit.")
            else:
                if verbose:
                    print("[DownloadCanvas] No PDF files found in the downloaded folder to check.")
                else:
                    print("No PDF files found in the downloaded folder to check.")

    except Exception as e:
        if verbose:
            print(f"[DownloadCanvas] Error downloading submissions: {e}")
        else:
            print(f"Error downloading submissions: {e}")

def download_canvas_assignment_submissions_auto(
    assignment_id,
    dest_dir=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    verbose=False
):
    """
    Download latest submissions for a single Canvas assignment without prompts.
    Returns (output_dir, downloaded_files).
    """
    canvas = get_canvas_client(api_url, api_key)
    course = canvas.get_course(course_id)
    assignment = course.get_assignment(assignment_id)
    lock_at = getattr(assignment, "lock_at", None)
    if lock_at:
        try:
            lock_dt = datetime.strptime(lock_at, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
            if lock_dt > datetime.now(timezone.utc):
                if verbose:
                    print(f"[WeeklyAuto] Warning: assignment {assignment_id} is not locked yet (lock_at={lock_at}).")
        except Exception:
            pass
    assignment_name = assignment.name.replace("/", "_").replace("\\", "_")
    group_name = "UnknownGroup"
    try:
        group_id = assignment.assignment_group_id
        for ag in course.get_assignment_groups():
            if ag.id == group_id:
                group_name = ag.name.replace("/", "_").replace("\\", "_")
                break
    except Exception:
        pass
    folder_name = f"{group_name}_{assignment_name}".replace(" ", "_")
    if dest_dir:
        out_dir = os.path.join(dest_dir, folder_name)
    else:
        out_dir = os.path.join(DEFAULT_DOWNLOAD_FOLDER, folder_name)
    os.makedirs(out_dir, exist_ok=True)
    if verbose:
        print(f"[DownloadCanvasAuto] Downloading submissions to: {out_dir}")
    else:
        print(f"Downloading submissions to: {out_dir}")

    submissions = list(assignment.get_submissions(include=["user", "attachments"]))
    due_at = getattr(assignment, "due_at", None)
    due_dt = None
    if due_at:
        try:
            due_dt = datetime.strptime(due_at, "%Y-%m-%dT%H:%M:%SZ")
        except Exception:
            due_dt = None

    latest_submissions = {}
    for sub in submissions:
        user = getattr(sub, "user", {})
        canvas_id = user.get("id", "unknown")
        submitted_at = getattr(sub, "submitted_at", None)
        if not submitted_at:
            continue
        try:
            sub_dt = datetime.strptime(submitted_at, "%Y-%m-%dT%H:%M:%SZ")
        except Exception:
            sub_dt = None
        if canvas_id not in latest_submissions or (sub_dt and latest_submissions[canvas_id]["sub_dt"] and sub_dt > latest_submissions[canvas_id]["sub_dt"]):
            latest_submissions[canvas_id] = {
                "submission": sub,
                "sub_dt": sub_dt,
                "submitted_at": submitted_at
            }

    downloaded_files = []
    for canvas_id, info in tqdm(latest_submissions.items(), desc=f"Downloading latest submissions for {assignment_name}"):
        sub = info["submission"]
        user = getattr(sub, "user", {})
        student_name = user.get("name", "UnknownStudent").replace("/", "_").replace("\\", "_").replace(" ", "_")
        submitted_at = info["submitted_at"]
        if submitted_at:
            try:
                dt = datetime.strptime(submitted_at, "%Y-%m-%dT%H:%M:%SZ")
                dt_str = dt.strftime("%Y%m%d_%H%M%S")
            except Exception:
                dt_str = "no_time"
        else:
            dt_str = "no_time"
        status = "on_time"
        if due_dt and submitted_at:
            try:
                sub_dt = datetime.strptime(submitted_at, "%Y-%m-%dT%H:%M:%SZ")
                status = "on_time" if sub_dt <= due_dt else "late"
            except Exception:
                status = "on_time"
        elif getattr(sub, "late", False):
            status = "late"
        attachments = getattr(sub, "attachments", [])
        if not attachments:
            continue
        for att in attachments:
            url = getattr(att, "url", None)
            filename = getattr(att, "filename", "file")
            ext = os.path.splitext(filename)[1]
            new_filename = f"{student_name}_{canvas_id}_{assignment_id}_{dt_str}_{status}{ext}"
            out_path = os.path.join(out_dir, new_filename)
            if os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                downloaded_files.append(out_path)
                continue
            try:
                resp = requests.get(url, headers={"Authorization": f"Bearer {api_key}"}, stream=False, timeout=60)
                resp.raise_for_status()
                with open(out_path, "wb") as f:
                    f.write(resp.content)
                downloaded_files.append(out_path)
            except Exception as e:
                if verbose:
                    print(f"[DownloadCanvasAuto] Failed to download {filename} for {student_name}: {e}")
                else:
                    print(f"Failed to download {filename} for {student_name}: {e}")

    return out_dir, downloaded_files


def add_comment_to_canvas_submission(
    assignment_id=None,
    student_canvas_id=None,
    comment_text=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    refine=DEFAULT_AI_METHOD,
    verbose=False
):
    """
    Add a comment to a student's submission for a given Canvas assignment.
    If assignment_id or student_canvas_id is None, interactively prompt the user to select.
    If comment_text is None, prompt the user to input the comment (multi-line supported).
    Only list assignments having at least one submission (submitted_at is not None for any student).
    If category is specified, only list assignments in that category.
    If refine is 'gemini', 'huggingface', or 'local', refine the comment via AI before posting.
    Returns True if successful, False otherwise.
    At every user prompt, allow quitting by entering 'q' or 'quit'.
    If no response after 60 seconds, execute the default option (quit if available).
    If verbose is True, print more details; otherwise, print only important notice.
    """

    def timeout_handler(signum, frame):
        if verbose:
            print("\n[CanvasComment] Timeout: No response after 60 seconds. Using default option.")
        else:
            print("\nTimeout: No response after 60 seconds. Using default option.")
        raise TimeoutError("User input timeout")

    def _prompt(prompt, timeout=60, default=None):
        return prompt_with_timeout(prompt, timeout=timeout, default=default, verbose=verbose)

    try:
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)

        # Select assignment if not provided
        if not assignment_id:
            if verbose:
                print("[CanvasComment] No assignment ID specified. Listing assignments with at least one actual submission:")
            else:
                print("No assignment ID specified. Listing assignments with at least one actual submission:")
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
                msg = "No assignments found with submissions." if not category else f"No assignments found with submissions in category '{category}'."
                if verbose:
                    print(f"[CanvasComment] {msg}")
                else:
                    print(msg)
                return False
            def due_sort_key(a):
                raw = a.get("due_at")
                if raw:
                    try:
                        return datetime.strptime(raw, "%Y-%m-%dT%H:%M:%SZ")
                    except Exception:
                        return datetime.max
                return datetime.max
            assignments.sort(key=due_sort_key)
            for idx, a in enumerate(assignments, 1):
                due = a['due_at'] or "No due date"
                if verbose:
                    print(f"[CanvasComment] {idx}. [{a['group']}] {a['name']} (ID: {a['id']}, Due: {due})")
                else:
                    print(f"{idx}. [{a['group']}] {a['name']} (ID: {a['id']}, Due: {due})")
            while True:
                try:
                    sel = _prompt(
                        "Enter the number of the assignment to comment on (or 'q' to quit): ",
                        timeout=60,
                        default="q"
                    ).strip()
                except TimeoutError:
                    sel = "q"
                if sel.lower() in ('q', 'quit'):
                    if verbose:
                        print("[CanvasComment] Quitting.")
                    else:
                        print("Quitting.")
                    return False
                if sel.isdigit() and 1 <= int(sel) <= len(assignments):
                    assignment_id = assignments[int(sel)-1]['id']
                    break
                else:
                    if verbose:
                        print("[CanvasComment] Invalid selection. Please enter a valid number or 'q' to quit.")
                    else:
                        print("Invalid selection. Please enter a valid number or 'q' to quit.")

        assignment = course.get_assignment(assignment_id)

        # Select student if not provided
        if not student_canvas_id:
            if verbose:
                print("[CanvasComment] Fetching submissions for the assignment...")
            else:
                print("Fetching submissions for the assignment...")
            students = []
            for sub in assignment.get_submissions(include=["user"]):
                user = getattr(sub, "user", {})
                student_name = user.get("name", "UnknownStudent")
                canvas_id = user.get("id", "unknown")
                submitted_at = getattr(sub, "submitted_at", None)
                if submitted_at:
                    students.append({
                        "canvas_id": canvas_id,
                        "name": student_name,
                        "submitted_at": submitted_at
                    })
            if not students:
                if verbose:
                    print("[CanvasComment] No students found with submissions for this assignment.")
                else:
                    print("No students found with submissions for this assignment.")
                return False
            students.sort(key=lambda s: s["name"])
            for idx, s in enumerate(students, 1):
                submitted = s["submitted_at"] or "No submission"
                if verbose:
                    print(f"[CanvasComment] {idx}. {s['name']} (Canvas ID: {s['canvas_id']}), Submitted at: {submitted}")
                else:
                    print(f"{idx}. {s['name']} (Canvas ID: {s['canvas_id']}), Submitted at: {submitted}")
            while True:
                try:
                    sel = _prompt(
                        "Enter the number of the student to comment on (or 'q' to quit): ",
                        timeout=60,
                        default="q"
                    ).strip()
                except TimeoutError:
                    sel = "q"
                if sel.lower() in ('q', 'quit'):
                    if verbose:
                        print("[CanvasComment] Quitting.")
                    else:
                        print("Quitting.")
                    return False
                if sel.isdigit() and 1 <= int(sel) <= len(students):
                    student_canvas_id = students[int(sel)-1]['canvas_id']
                    break
                else:
                    if verbose:
                        print("[CanvasComment] Invalid selection. Please enter a valid number or 'q' to quit.")
                    else:
                        print("Invalid selection. Please enter a valid number or 'q' to quit.")

        if not comment_text:
            if verbose:
                print("[CanvasComment] Enter the comment to add (multi-line, end with an empty line). Enter 'q' or 'quit' on a new line to quit.")
            else:
                print("Enter the comment to add (multi-line, end with an empty line). Enter 'q' or 'quit' on a new line to quit.")
            lines = []
            while True:
                try:
                    line = _prompt("", timeout=60, default="q")
                except TimeoutError:
                    if verbose:
                        print("[CanvasComment] No input after 60 seconds. Quitting.")
                    else:
                        print("No input after 60 seconds. Quitting.")
                    return False
                if line.strip().lower() in ('q', 'quit'):
                    if verbose:
                        print("[CanvasComment] Quitting.")
                    else:
                        print("Quitting.")
                    return False
                if line == "":
                    break
                lines.append(line)
            comment_text = "\n".join(lines)
            if not comment_text.strip():
                if verbose:
                    print("[CanvasComment] No comment entered. Aborting.")
                else:
                    print("No comment entered. Aborting.")
                return False

        # Optionally refine the comment using AI
        if refine in ALL_AI_METHODS:
            if verbose:
                print(f"[CanvasComment] Refining comment using AI service: {refine} ...")
            else:
                print(f"Refining comment using AI service: {refine} ...")
            prompt = (
                "You are an expert assistant. Please refine the following comment for clarity, "
                "correct any spelling or grammar mistakes, and improve readability. "
                "Keep the meaning and important information unchanged. "
                "Return ONLY the improved comment, no explanations.\n\n"
                "Comment:\n{text}"
            )
            refined_comment = refine_text_with_ai(comment_text, method=refine, user_prompt=prompt)
            if verbose:
                print("\n[CanvasComment] Refined comment:\n")
                print(refined_comment)
            else:
                print("\nRefined comment:\n")
                print(refined_comment)
            while True:
                try:
                    confirm = _prompt(
                        "\nDo you want to use the refined comment? (y/n, or 'q' to quit): ",
                        timeout=60,
                        default="q"
                    ).strip().lower()
                except TimeoutError:
                    confirm = "q"
                if confirm in ("q", "quit"):
                    if verbose:
                        print("[CanvasComment] Quitting.")
                    else:
                        print("Quitting.")
                    return False
                if confirm == "y":
                    comment_text = refined_comment
                    break
                elif confirm == "n":
                    if verbose:
                        print("[CanvasComment] Using the original comment.")
                    else:
                        print("Using the original comment.")
                    break
                else:
                    if verbose:
                        print("[CanvasComment] Please enter 'y', 'n', or 'q' to quit.")
                    else:
                        print("Please enter 'y', 'n', or 'q' to quit.")

        submission = assignment.get_submission(student_canvas_id)
        submission.edit(comment={'text_comment': comment_text})
        if verbose:
            print(f"[CanvasComment] Comment added to submission of student {student_canvas_id} for assignment {assignment_id}.")
        else:
            print(f"Comment added to submission of student {student_canvas_id} for assignment {assignment_id}.")
        return True
    except Exception as e:
        if verbose:
            print(f"[CanvasComment] Failed to add comment: {e}")
        else:
            print(f"Failed to add comment: {e}")
        return False

