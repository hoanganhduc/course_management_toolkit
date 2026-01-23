# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas announcement and final evaluation helpers."""

import os
import re
import signal
from datetime import datetime

from .canvas_auth import get_canvas_client

from .config import get_cached_course_code, get_default_db_path
from .data import load_database, refine_text_with_ai
from .canvas_utils import prompt_with_timeout, parse_selection, normalize_datetime_str
from .settings import (
    CANVAS_LMS_API_URL,
    CANVAS_LMS_API_KEY,
    CANVAS_LMS_COURSE_ID,
    COURSE_CODE,
    COURSE_NAME,
    DRY_RUN,
)

def send_final_evaluations_via_canvas(
    final_dir="final_evaluations",
    db_path=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    dry_run=False,
    verbose=False
):
    """
    Read each "<student id>_<student name>_results.txt" in final_dir and send its content
    to the corresponding student via Canvas conversation. Student's Canvas ID is looked up
    from the local database (db_path). Message is in Vietnamese and mentions it was sent
    automatically and to contact the lecturer ASAP if any problem arises.

    Enhancement: if the student's name is missing in the database, extract it from the
    text filename (the part between the first underscore after the 8-digit id and
    "_results.txt"), replacing underscores with spaces.

    dry_run: if True, do not send any messages and only report what would be sent.
    """
    if db_path is None:
        db_path = get_default_db_path()

    if not os.path.isdir(final_dir):
        if verbose:
            print(f"[SendFinals] Folder not found: {final_dir}")
        else:
            print(f"Folder not found: {final_dir}")
        return {"sent": 0, "skipped": 0, "errors": 0}

    # Load database
    students = load_database(db_path) if os.path.exists(db_path) else []
    sid_map = {}
    for s in students:
        sid = getattr(s, "Student ID", None)
        if sid:
            sid_map[str(sid).strip()] = s

    # Prepare Canvas client
    try:
        canvas = get_canvas_client(api_url, api_key)
    except Exception as e:
        if verbose:
            print(f"[SendFinals] Failed to initialize Canvas client: {e}")
        else:
            print("Failed to initialize Canvas client.")
        return {"sent": 0, "skipped": 0, "errors": 0}

    sent = 0
    skipped = 0
    errors = 0
    dry_run_count = 0

    # Capture student id and name part from filename
    pattern = re.compile(r'^(\d{8})_(.+?)_results\.txt$', re.IGNORECASE)
    for fname in sorted(os.listdir(final_dir)):
        if not fname.lower().endswith("_results.txt"):
            continue
        m = pattern.match(fname)
        if not m:
            if verbose:
                print(f"[SendFinals] Skipping file with unexpected name format: {fname}")
            else:
                print(f"Skipping file: {fname}")
            skipped += 1
            continue

        sid = m.group(1)
        raw_name_part = m.group(2) or ""
        # Clean up extracted name: replace underscores and multiple separators with spaces
        extracted_name = re.sub(r'[_\-\.\s]+', ' ', raw_name_part).strip()
        # Further sanitize: remove any trailing/leading non-letter characters
        extracted_name = re.sub(r'^[^A-Za-z0-9À-ỹà-ỹ]+|[^A-Za-z0-9À-ỹà-ỹ]+$', '', extracted_name).strip()

        student = sid_map.get(sid)
        file_path = os.path.join(final_dir, fname)
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read().strip()
        except Exception as e:
            if verbose:
                print(f"[SendFinals] Failed to read {file_path}: {e}")
            else:
                print(f"Failed to read {file_path}.")
            errors += 1
            continue

        if not student:
            if verbose:
                print(f"[SendFinals] Student ID {sid} not found in database. Skipping file: {fname}")
            else:
                print(f"Student ID {sid} not found, skipped.")
            skipped += 1
            continue

        # New preference: prefer name extracted from filename; if it fails/empty, use database name
        db_name = getattr(student, "Name", "")
        if isinstance(db_name, str):
            db_name = db_name.strip()
        else:
            db_name = str(db_name) if db_name is not None else ""

        if extracted_name:
            student_name = extracted_name
        else:
            student_name = db_name

        # As a last resort, if both are empty, derive a readable name from filename base (without id)
        if not student_name:
            student_name = extracted_name or db_name or ""

        canvas_id = getattr(student, "Canvas ID", None) or student.__dict__.get("Canvas ID")
        if not canvas_id:
            if verbose:
                print(f"[SendFinals] Canvas ID missing for student {sid}. Skipping.")
            else:
                print(f"Canvas ID missing for {sid}, skipped.")
            skipped += 1
            continue

        # Ensure recipient id is a string
        recipient = str(canvas_id)

        greeting = f"Chào {student_name},\n\n" if student_name else "Chào bạn,\n\n"

        sent_time = datetime.now().strftime("%d/%m/%Y %H:%M")
        greeting += f"Đây là thông báo tự động gửi kết quả đánh giá học phần cho sinh viên (MSSV: {sid}).\n"
        greeting += f"Thời điểm gửi: {sent_time}.\n\n"
        greeting += "Vui lòng kiểm tra kỹ nội dung bên dưới. Nếu có thắc mắc, sai sót về thông tin hoặc điểm số, bạn hãy phản hồi trực tiếp cho giảng viên càng sớm càng tốt.\n"
        greeting += "Chúc bạn một học kỳ tốt và kết quả học tập tiến bộ.\n"

        course_code = (COURSE_CODE or get_cached_course_code() or "").strip()
        course_name = (COURSE_NAME or "").strip()
        if course_code:
            course_code = course_code.upper()
        if course_name:
            course_name = course_name.title()
        course_title = " - ".join(part for part in (course_code, course_name) if part)
        subject = "Thông báo kết quả đánh giá học phần"
        if course_title:
            subject = f"{subject} ({course_title})"

        body = (
            f"{greeting}\n\n"
            f"{content}\n\n"
        )

        if dry_run:
            dry_run_count += 1
            if verbose:
                print(f"[SendFinals] Dry run: would send results to {sid} (Canvas ID: {recipient}) from file {fname}")
            continue

        try:
            canvas.create_conversation(
                recipients=[recipient],
                subject=subject,
                body=body,
                force_new=True
            )
            sent += 1
            if verbose:
                print(f"[SendFinals] Sent results to {sid} (Canvas ID: {recipient}) from file {fname}")
        except Exception as e:
            errors += 1
            if verbose:
                print(f"[SendFinals] Failed to send to {sid} (Canvas ID: {recipient}): {e}")
            else:
                print(f"Failed to send to {sid}.")
            continue

    summary = {"sent": sent, "skipped": skipped, "errors": errors}
    if dry_run:
        summary["dry_run"] = dry_run_count
    if verbose:
        if dry_run:
            print(f"[SendFinals] Completed. Sent: {sent}, Skipped: {skipped}, Errors: {errors}, Dry-run: {dry_run_count}")
        else:
            print(f"[SendFinals] Completed. Sent: {sent}, Skipped: {skipped}, Errors: {errors}")
    else:
        if dry_run:
            print(f"Done. Sent: {sent}, Skipped: {skipped}, Errors: {errors}, Dry-run: {dry_run_count}")
        else:
            print(f"Done. Sent: {sent}, Skipped: {skipped}, Errors: {errors}")
    return summary


def _parse_announcement_file(message_path):
    with open(message_path, "r", encoding="utf-8") as f:
        lines = f.readlines()
    title = None
    message_lines = []
    in_message = False
    for line in lines:
        if line.strip().lower().startswith("title:"):
            title = line.strip()[6:].strip()
        elif line.strip().lower().startswith("message:"):
            in_message = True
        elif in_message:
            message_lines.append(line.rstrip("\n"))
    if title or in_message:
        return title, "\n".join(message_lines).strip()
    return None, "".join(lines).strip()


def add_canvas_announcement(
    title,
    message,
    message_path=None,
    refine=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    dry_run=False,
    verbose=False
):
    """
    Add an announcement to a Canvas course.
    Accepts a short message (manual input or TXT file), optionally refines with AI,
    confirms with the user, and posts to Canvas.
    """

    def timeout_handler(signum, frame):
        if verbose:
            print("\n[CanvasAnnouncement] Timeout: No response after 60 seconds. Using default option.")
        else:
            print("\nTimeout: No response after 60 seconds. Using default option.")
        raise TimeoutError("User input timeout")

    try:
        if message_path:
            file_title, file_message = _parse_announcement_file(message_path)
            title = title or file_title
            message = message or file_message

        if not title or not title.strip():
            title = input("Announcement title: ").strip()
        if not message or not message.strip():
            message = input("Short announcement message: ").strip()

        if refine and refine != "none":
            try:
                user_prompt = (
                    "You are a formal announcement editor. Rewrite the following message in a more formal tone. "
                    "Keep all details accurate (dates, times, deadlines, numbers, and names must not change). "
                    "The final message must start with a greeting such as 'Dear all,' or 'Chào các bạn,'. "
                    "Choose the greeting to match the language of the input. "
                    "Return ONLY the rewritten announcement text.\n\n"
                    "Message:\n{text}"
                )
                refined = refine_text_with_ai(
                    message,
                    method=refine,
                    verbose=verbose,
                    user_prompt=user_prompt,
                )
                if refined and refined.strip():
                    message = refined.strip()
            except Exception as exc:
                if verbose:
                    print(f"[CanvasAnnouncement] Failed to refine message: {exc}")

        # Confirm submission
        if verbose:
            print("\n[CanvasAnnouncement] Preview of announcement to submit:")
            print(f"Title: {title}")
            print("Message:")
            print(message[:500] + ("..." if len(message) > 500 else ""))
        else:
            print("\nPreview of announcement to submit:")
            print(f"Title: {title}")
            print("Message:")
            print(message[:500] + ("..." if len(message) > 500 else ""))
        try:
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(60)
            confirm = input(
            "\nFinish editing and submit this announcement? (y/n, default y): "
            ).strip().lower() or "y"
            signal.alarm(0)
        except TimeoutError:
            if verbose:
                print("[CanvasAnnouncement] No response after 60 seconds. Using default: y")
            else:
                print("No response after 60 seconds. Using default: y")
            confirm = "y"
        except KeyboardInterrupt:
            signal.alarm(0)
            if verbose:
                print("\n[CanvasAnnouncement] Operation cancelled by user.")
            else:
                print("\nOperation cancelled by user.")
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
            return {"error": "User cancelled."}
        if confirm in ("n", "no"):
            if verbose:
                print("[CanvasAnnouncement] Cancelled.")
            else:
                print("Cancelled.")
            return {"error": "User cancelled."}

        if dry_run or DRY_RUN:
            if verbose:
                print(f"[CanvasAnnouncement] Dry run: '{title}' not posted.")
            else:
                print(f"Dry run: '{title}' not posted.")
            return {"dry_run": True, "title": title, "message": message}

        # Post to Canvas
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)
        discussion = course.create_discussion_topic(
            title=title,
            message=message,
            is_announcement=True
        )
        if verbose:
            print(f"[CanvasAnnouncement] Announcement '{title}' created for course {course_id}.")
        else:
            print(f"Announcement '{title}' created for course {course_id}.")
        return discussion
    except Exception as e:
        if verbose:
            print(f"[CanvasAnnouncement] Failed to create announcement: {e}")
        else:
            print(f"Failed to create announcement: {e}")
        return {"error": str(e)}

