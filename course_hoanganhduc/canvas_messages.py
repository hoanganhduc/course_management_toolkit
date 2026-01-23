# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas messaging helpers."""

import signal
from datetime import datetime, timezone

import requests
from .canvas_auth import get_canvas_client

from .canvas_people import list_canvas_people
from .data import refine_text_with_ai
from .canvas_utils import prompt_with_timeout, parse_selection, normalize_datetime_str
from .settings import (
    ALL_AI_METHODS,
    CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    CANVAS_LMS_API_KEY,
    CANVAS_LMS_API_URL,
    CANVAS_LMS_COURSE_ID,
    DEFAULT_AI_METHOD,
)

def send_canvas_message_to_students(
    subject=None,
    message=None,
    student_canvas_ids=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    refine=DEFAULT_AI_METHOD,
    verbose=False
):
    """
    Send a message via Canvas to one or more students.
    If student_canvas_ids is not specified, list all active students and allow user to select (supports ranges, e.g., 1-5,7,9).
    Args:
        subject (str): Subject of the message.
        message (str): Message body (plain text or HTML).
        student_canvas_ids (list): List of Canvas user IDs to send to.
        api_url, api_key, course_id: Canvas API info.
        refine (str): Refine message using AI ('gemini', 'huggingface', or None).
        verbose (bool): If True, print more details; otherwise, print only important notice.
    Returns:
        True if sent, False otherwise.
    """

    def timeout_handler(signum, frame):
        if verbose:
            print("\n[CanvasMessage] Timeout: No response after 60 seconds. Using default option.")
        else:
            print("\nTimeout: No response after 60 seconds. Using default option.")
        raise TimeoutError("User input timeout")

    def _prompt(prompt, timeout=60, default=None):
        return prompt_with_timeout(prompt, timeout=timeout, default=default, verbose=verbose)

    try:
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)
        # Get all active students if not provided
        if not student_canvas_ids:
            people = list_canvas_people(api_url, api_key, course_id, verbose=verbose)
            students = people.get("active_students", [])
            if not students:
                if verbose:
                    print("[CanvasMessage] No active students found.")
                else:
                    print("No active students found.")
                return False
            if verbose:
                print("[CanvasMessage] Active students:")
            else:
                print("Active students:")
            for idx, s in enumerate(students, 1):
                if verbose:
                    print(f"[CanvasMessage] {idx}. {s['name']} ({s['email']}), Canvas ID: {s['canvas_id']}")
                else:
                    print(f"{idx}. {s['name']} ({s['email']}), Canvas ID: {s['canvas_id']}")
            if verbose:
                print("[CanvasMessage] Select students to send message to (e.g., 1-5,7,9 or 'a' for all, or 'q' to quit):")
            else:
                print("Select students to send message to (e.g., 1-5,7,9 or 'a' for all, or 'q' to quit):")
            while True:
                sel = _prompt("Enter selection: ", timeout=60, default="q").strip()
                if sel.lower() in ("q", "quit"):
                    if verbose:
                        print("[CanvasMessage] Quitting.")
                    else:
                        print("Quitting.")
                    return False
                if sel.lower() in ("a", "all"):
                    selected = list(range(1, len(students) + 1))
                else:
                    selected = []
                    for part in sel.split(","):
                        part = part.strip()
                        if "-" in part:
                            try:
                                start, end = map(int, part.split("-"))
                                selected.extend(range(start, end + 1))
                            except Exception:
                                continue
                        elif part.isdigit():
                            selected.append(int(part))
                    selected = [i for i in selected if 1 <= i <= len(students)]
                if not selected:
                    if verbose:
                        print("[CanvasMessage] No valid selection. Try again or 'q' to quit.")
                    else:
                        print("No valid selection. Try again or 'q' to quit.")
                    continue
                student_canvas_ids = [students[i - 1]["canvas_id"] for i in selected]
                break

        # Prompt for subject if not provided
        while not subject:
            if subject.lower() in ("q", "quit"):
                if verbose:
                    print("[CanvasMessage] Quitting.")
                else:
                    print("Quitting.")
                return False
        # Prompt for message if not provided
        if not message:
            if verbose:
                print("[CanvasMessage] Enter message body (multi-line, end with an empty line, or 'q' to quit):")
            else:
                print("Enter message body (multi-line, end with an empty line, or 'q' to quit):")
            lines = []
            while True:
                line = _prompt("", timeout=60, default="q")
                if line.strip().lower() in ("q", "quit"):
                    if verbose:
                        print("[CanvasMessage] Quitting.")
                    else:
                        print("Quitting.")
                    return False
                if line == "":
                    break
                lines.append(line)
            message = "\n".join(lines)
        if not message:
            if verbose:
                print("[CanvasMessage] Message body is required.")
            else:
                print("Message body is required.")
            return False

        # Optionally refine the message using AI
        if refine in ALL_AI_METHODS:
            if verbose:
                print(f"[CanvasMessage] Refining message using AI service: {refine} ...")
            else:
                print(f"Refining message using AI service: {refine} ...")
            prompt = (
                "You are an expert assistant. Please refine the following message for clarity, "
                "correct any spelling or grammar mistakes, and improve readability. "
                "Keep the meaning and important information unchanged. "
                "Return ONLY the improved message, no explanations.\n\n"
                "Message:\n{text}"
            )
            refined_message = refine_text_with_ai(message, method=refine, user_prompt=prompt)
            if verbose:
                print("\n[CanvasMessage] Refined message:\n")
                print(refined_message)
            else:
                print("\nRefined message:\n")
                print(refined_message)
            while True:
                confirm = _prompt(
                    "\nDo you want to use the refined message? (y/n, or 'q' to quit): ",
                    timeout=60,
                    default="q"
                ).strip().lower()
                if confirm in ("q", "quit"):
                    if verbose:
                        print("[CanvasMessage] Quitting.")
                    else:
                        print("Quitting.")
                    return False
                if confirm == "y":
                    message = refined_message
                    break
                elif confirm == "n":
                    if verbose:
                        print("[CanvasMessage] Using the original message.")
                    else:
                        print("Using the original message.")
                    break
                else:
                    if verbose:
                        print("[CanvasMessage] Please enter 'y', 'n', or 'q' to quit.")
                    else:
                        print("Please enter 'y', 'n', or 'q' to quit.")

        # Send message via Canvas Conversations API
        recipient_ids = [str(uid) for uid in student_canvas_ids]
        user = canvas.get_current_user()
        if verbose:
            print(f"[CanvasMessage] Sending message from {user.name} to {len(recipient_ids)} student(s)...")
        else:
            print(f"Sending message from {user.name} to {len(recipient_ids)} student(s)...")
        canvas.create_conversation(
            recipients=recipient_ids,
            subject=subject,
            body=message,
            force_new=True
        )
        if verbose:
            print("[CanvasMessage] Message sent successfully.")
        else:
            print("Message sent successfully.")
        return True
    except Exception as e:
        if verbose:
            print(f"[CanvasMessage] Failed to send message: {e}")
        else:
            print(f"Failed to send message: {e}")
        return False

def notify_incomplete_canvas_peer_reviews(
    assignment_id=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    verbose=False,
    refine=None,
    confirm_each=True  # New option: if False, send to all without confirmation
):
    """
    For each submission in a Canvas assignment, list students who are assigned to review that submission.
    For each reviewer with pending reviews, send a reminder message to that user (optionally refined via AI),
    including the list of pending reviews and the assignment title. Reviews are considered incomplete if rubric points are not given.
    Only list assignments whose lock date has passed.
    Allows user to select one or more assignments (supports ranges, comma-separated, or 'a' for all).
    Optionally allows sending messages to all students without confirming one by one.
    Does not return anything.
    """

    def timeout_handler(_signum=None, _frame=None):
        print("\nTimeout: No response after 60 seconds. Quitting...")
        raise TimeoutError("User input timeout")

    def _prompt(prompt, timeout=60, default=None):
        return prompt_with_timeout(prompt, timeout=timeout, default=default, verbose=verbose)

    try:
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)
        assignments = []
        assignment_groups = list(course.get_assignment_groups(include=['assignments']))
        now = datetime.now(timezone.utc)
        for group in assignment_groups:
            group_name = group.name
            if category and group_name.lower() != category.lower():
                continue
            for assignment in group.assignments:
                # Only include assignments whose lock date has passed
                lock_at = assignment.get('lock_at')
                if lock_at:
                    try:
                        lock_dt = datetime.strptime(lock_at, "%Y-%m-%dT%H:%M:%SZ")
                        if lock_dt > now:
                            continue
                    except Exception:
                        pass
                else:
                    # If no lock date, skip (do not include assignments without lock date)
                    continue
                if assignment.get('has_submitted_submissions', False):
                    assignments.append({
                        "id": assignment['id'],
                        "name": assignment['name'],
                        "group": group_name,
                        "due_at": assignment.get('due_at'),
                        "lock_at": lock_at
                    })
        if not assignments:
            print("No assignments found with submissions and lock date passed.")
            return

        def due_sort_key(a):
            raw = a.get("due_at")
            if raw:
                try:
                    return datetime.strptime(raw, "%Y-%m-%dT%H:%M:%SZ")
                except Exception:
                    return datetime.max
            return datetime.max
        assignments.sort(key=due_sort_key)
        print("Assignments with submissions and lock date passed:")
        for idx, a in enumerate(assignments, 1):
            due = a['due_at'] or "No due date"
            lock = a['lock_at'] or "No lock date"
            print(f"{idx}. [{a['group']}] {a['name']} (ID: {a['id']}, Due: {due}, Lock: {lock})")

        selected_assignment_ids = []
        if not assignment_id:
            while True:
                sel = _prompt(
                    "Enter the number(s) of the assignment(s) to view reviews (e.g. 1,3-5 or 'a' for all, or 'q' to quit): ",
                    timeout=60,
                    default="q"
                ).strip()
                if sel.lower() in ('q', 'quit'):
                    return
                if sel.lower() in ('a', 'all'):
                    selected_assignment_ids = [a['id'] for a in assignments]
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
                selected_assignment_ids = [assignments[i - 1]['id'] for i in selected]
                break
        else:
            selected_assignment_ids = [assignment_id]

        for assignment_id in selected_assignment_ids:
            assignment = course.get_assignment(assignment_id)
            assignment_title = assignment.name
            submissions = list(assignment.get_submissions(include=["user"]))
            canvas_id_to_name = {}
            for sub in submissions:
                user = getattr(sub, "user", {})
                canvas_id = user.get("id", None)
                name = user.get("name", "")
                if canvas_id:
                    canvas_id_to_name[canvas_id] = name

            headers = {"Authorization": f"Bearer {api_key}"}
            url = f"{api_url}/api/v1/courses/{course_id}/assignments/{assignment_id}/peer_reviews?include[]=user"
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            peer_reviews = response.json()

            submission_map = {}
            for sub in submissions:
                user = getattr(sub, "user", {})
                canvas_id = user.get("id", None)
                name = user.get("name", "")
                submitted = getattr(sub, "submitted_at", None) is not None
                submission_map[canvas_id] = {
                    "canvas_id": canvas_id,
                    "name": name,
                    "submitted": submitted,
                    "reviewers": []
                }

            for pr in peer_reviews:
                asset_user_id = pr.get("user_id")
                assessor_id = pr.get("assessor_id")
                workflow_state = pr.get("workflow_state")
                reviewed_at = None
                if workflow_state == "completed":
                    reviewed_at = pr.get("reviewed_at", "DONE")
                reviewer_name = ""
                if "assessor" in pr and isinstance(pr["assessor"], dict):
                    reviewer_name = pr["assessor"].get("name", "")
                elif assessor_id in canvas_id_to_name:
                    reviewer_name = canvas_id_to_name.get(assessor_id, "")
                reviewer_info = {
                    "canvas_id": assessor_id,
                    "name": reviewer_name,
                    "reviewed_at": reviewed_at
                }
                if asset_user_id in submission_map:
                    submission_map[asset_user_id]["reviewers"].append(reviewer_info)

            submission_reviewers = list(submission_map.values())
            reviewer_pending_map = {}
            for sub_info in submission_reviewers:
                for r in sub_info["reviewers"]:
                    if not r["reviewed_at"]:
                        reviewer_id = r["canvas_id"]
                        if reviewer_id not in reviewer_pending_map:
                            reviewer_pending_map[reviewer_id] = []
                        reviewer_pending_map[reviewer_id].append(sub_info)

            if not reviewer_pending_map:
                print(f"No reviewers with pending reviews for assignment ID {assignment_id}.")
                continue

            print(f"\nReviewers with pending reviews for assignment '{assignment_title}':")
            for reviewer_id, pending_subs in reviewer_pending_map.items():
                reviewer_name = canvas_id_to_name.get(reviewer_id, "")
                print(f"- {reviewer_name} (Canvas ID: {reviewer_id}): {len(pending_subs)} pending reviews")
                for sub in pending_subs:
                    print(f"    * {sub['name']} (Canvas ID: {sub['canvas_id']})")

            # Ask if user wants to send all messages without confirmation
            send_all = False
            send_messages = _prompt(
                "\nDo you want to send reminder messages to these reviewers? (y/n, 'a' for all without confirmation, or 'q' to quit, default 'a' in 60s): ",
                timeout=60,
                default="a"
            ).strip().lower()
            if send_messages in ("q", "quit"):
                return
            if send_messages == "a":
                send_all = True
            elif send_messages not in ("y", "yes", ""):
                print("Messages not sent.")
                continue

            for reviewer_id, pending_subs in reviewer_pending_map.items():
                reviewer_name = canvas_id_to_name.get(reviewer_id, "")
                pending_list_str = "\n".join(
                    [f"- {sub['name']} (Canvas ID: {sub['canvas_id']})" for sub in pending_subs]
                )
                note = (
                    "Lưu ý: Một bài đánh giá được coi là chưa hoàn thành nếu bạn chưa cho điểm các tiêu chí (rubrics) trong phần đánh giá."
                )
                base_message = (
                    f"Chào {reviewer_name},\n\n"
                    f"Hệ thống phát hiện bạn còn {len(pending_subs)} bài đánh giá chưa hoàn thành cho bài tập \"{assignment_title}\":\n"
                    f"{pending_list_str}\n\n"
                    f"{note}\n\n"
                    "Vui lòng hoàn thành các đánh giá càng sớm càng tốt để đảm bảo tiến độ học tập.\n"
                    "Thông báo này được gửi tự động bởi hệ thống."
                )
                message = base_message
                if refine in ALL_AI_METHODS:
                    prompt = (
                        "Bạn là trợ lý giáo dục chuyên nghiệp. Hãy viết lại thông báo sau bằng tiếng Việt, lịch sự, rõ ràng, "
                        "giải thích rằng các bài đánh giá được coi là chưa hoàn thành nếu chưa cho điểm các tiêu chí (rubrics), "
                        "và nhắc sinh viên hoàn thành các đánh giá còn thiếu. Đưa vào danh sách các bài cần đánh giá và tên bài tập. "
                        "Chỉ trả về thông báo đã chỉnh sửa, không giải thích gì thêm.\n\n"
                        "Thông báo:\n{text}"
                    )
                    message = refine_text_with_ai(base_message, method=refine, user_prompt=prompt)
                    print("\nRefined message preview:\n")
                    print(message[:500] + ("..." if len(message) > 500 else ""))
                    if not send_all and confirm_each:
                        confirm = _prompt(
                            "Send this refined message? (y/n, or 'q' to skip): ",
                            timeout=60,
                            default="y"
                        ).strip().lower()
                        if confirm in ("q", "quit", "n", "no"):
                            print("Skipped sending message to this reviewer.")
                            continue
                else:
                    print("\nMessage preview:\n")
                    print(message[:500] + ("..." if len(message) > 500 else ""))
                    if not send_all and confirm_each:
                        confirm = _prompt(
                            "Send this message? (y/n, or 'q' to skip): ",
                            timeout=60,
                            default="y"
                        ).strip().lower()
                        if confirm in ("q", "quit", "n", "no"):
                            print("Skipped sending message to this reviewer.")
                            continue

                try:
                    canvas.create_conversation(
                        recipients=[str(reviewer_id)],
                        subject=subject,
                        body=message,
                        force_new=True
                    )
                    print(f"Reminder sent to reviewer {reviewer_name} ({reviewer_id}) for {len(pending_subs)} pending reviews.")
                except Exception as e:
                    print(f"Failed to send reminder to reviewer {reviewer_name} ({reviewer_id}): {e}")

            print(f"All reminders sent to reviewers with pending reviews for assignment ID {assignment_id}.")

    except Exception as e:
        if verbose:
            print(f"[CanvasReviews] Error listing reviews or sending reminders: {e}")
        else:
            print(f"Error listing reviews or sending reminders: {e}")


def fetch_and_reply_canvas_messages(
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    only_unread=False,
    reply_text=None,
    refine=DEFAULT_AI_METHOD,
    max_messages=3,
    verbose=False
):
    """
    Fetches Canvas inbox conversations (messages), prints them, and allows replying.
    If only_unread is True, only fetches unread conversations.
    If reply_text is provided, replies with that text to all selected conversations.
    If refine is 'gemini', 'huggingface', or 'local', refines the reply via AI before sending.
    Messages are shown in order of latest first.
    max_messages: maximum number of conversations to fetch and display (default: 3)
    Also fetches and prints the full content of each message in the conversation.
    If verbose is True, print more details; otherwise, print only important notice.
    """
    try:
        canvas = get_canvas_client(api_url, api_key)
        user = canvas.get_current_user()
        if verbose:
            print(f"[CanvasMsg] Fetching Canvas inbox for user: {user.name} ({user.id})")
        else:
            print(f"Fetching Canvas inbox for user: {user.name} ({user.id})")
        inbox = canvas.get_conversations(scope="unread" if only_unread else "unread,read")
        conversations = list(inbox)
        if not conversations:
            msg = "No messages found."
            if verbose:
                print(f"[CanvasMsg] {msg}")
            else:
                print(msg)
            return
        # Sort conversations by last_message_at, latest first
        def parse_time(t):
            try:
                return datetime.strptime(t, "%Y-%m-%dT%H:%M:%SZ")
            except Exception:
                return datetime.min
        conversations.sort(key=lambda c: parse_time(getattr(c, "last_message_at", "")), reverse=True)
        # Limit number of conversations
        if max_messages is not None and max_messages > 0:
            conversations = conversations[:max_messages]
        if verbose:
            print(f"[CanvasMsg] Found {len(conversations)} conversation(s) (showing up to {max_messages}):")
        else:
            print(f"Found {len(conversations)} conversation(s) (showing up to {max_messages}):")
        for idx, conv in enumerate(conversations, 1):
            if verbose:
                print(f"\n[CanvasMsg] --- Conversation {idx} ---")
                print(f"[CanvasMsg] Subject: {conv.subject}")
                print(f"[CanvasMsg] ID: {conv.id}")
                print(f"[CanvasMsg] Last message at: {conv.last_message_at}")
                print(f"[CanvasMsg] Participants: {[p['name'] for p in conv.participants]}")
            else:
                print(f"\n--- Conversation {idx} ---")
                print(f"Subject: {conv.subject}")
                print(f"ID: {conv.id}")
                print(f"Last message at: {conv.last_message_at}")
                print(f"Participants: {[p['name'] for p in conv.participants]}")
            # Fetch full content of messages in this conversation
            try:
                full_conv = canvas.get_conversation(conv.id)
                messages = getattr(full_conv, "messages", [])
            except Exception:
                messages = getattr(conv, "messages", [])
            # Sort messages by created_at, latest first
            def msg_time(m):
                t = m.get('created_at', '')
                try:
                    return datetime.strptime(t, "%Y-%m-%dT%H:%M:%SZ")
                except Exception:
                    return datetime.min
            messages_sorted = sorted(messages, key=msg_time, reverse=True)
            for m in messages_sorted[:3]:
                sender = m.get('author', {}).get('display_name', 'Unknown')
                body = m.get('body', '')
                if verbose:
                    print(f"[CanvasMsg]   [{sender}]: {body}")
                else:
                    print(f"  [{sender}]: {body}")
            # Print all message contents (full history)
            if verbose:
                print("[CanvasMsg] Full message history (oldest first):")
            else:
                print("Full message history (oldest first):")
            messages_sorted_oldest = sorted(messages, key=msg_time)
            for m in messages_sorted_oldest:
                sender = m.get('author', {}).get('display_name', 'Unknown')
                body = m.get('body', '')
                created_at = m.get('created_at', '')
                if verbose:
                    print(f"[CanvasMsg]   [{created_at}] {sender}: {body}")
                else:
                    print(f"  [{created_at}] {sender}: {body}")
        # Ask user which conversations to reply to
        def _prompt(prompt, timeout=60, default=None):
            return prompt_with_timeout(prompt, timeout=timeout, default=default, verbose=verbose)

        sel = _prompt("Enter conversation numbers to reply (comma/range, 'a' for all, or 'q' to quit): ", timeout=60, default="q").strip()
        if sel.lower() in ('q', 'quit'):
            return
        if sel.lower() in ('a', 'all'):
            selected = list(range(1, len(conversations) + 1))
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
            selected = [i for i in selected if 1 <= i <= len(conversations)]
        if not selected:
            msg = "No valid selection."
            if verbose:
                print(f"[CanvasMsg] {msg}")
            else:
                print(msg)
            return
        for idx in selected:
            conv = conversations[idx - 1]
            if not reply_text:
                if verbose:
                    print(f"\n[CanvasMsg] Replying to conversation {idx} (Subject: {conv.subject})")
                    print("[CanvasMsg] Enter reply message (multi-line, end with empty line, or 'q' to skip):")
                else:
                    print(f"\nReplying to conversation {idx} (Subject: {conv.subject})")
                    print("Enter reply message (multi-line, end with empty line, or 'q' to skip):")
                lines = []
                while True:
                    line = input()
                    if line.strip().lower() in ('q', 'quit'):
                        lines = []
                        break
                    if line == "":
                        break
                    lines.append(line)
                if not lines:
                    if verbose:
                        print("[CanvasMsg] Skipped.")
                    else:
                        print("Skipped.")
                    continue
                msg = "\n".join(lines)
            else:
                msg = reply_text
            # Optionally refine
            if refine in ALL_AI_METHODS:
                if verbose:
                    print(f"[CanvasMsg] Refining reply using AI service: {refine} ...")
                else:
                    print(f"Refining reply using AI service: {refine} ...")
                prompt = (
                    "You are an expert assistant. Please refine the following reply for clarity, "
                    "correct any spelling or grammar mistakes, and improve readability. "
                    "Keep the meaning and important information unchanged. "
                    "Return ONLY the improved reply, no explanations.\n\n"
                    "Reply:\n{text}"
                )
                msg = refine_text_with_ai(msg, method=refine, user_prompt=prompt)
                if verbose:
                    print("\n[CanvasMsg] Refined reply:\n")
                    print(msg)
                else:
                    print("\nRefined reply:\n")
                    print(msg)
                confirm = input("Send this reply? (y/n): ").strip().lower()
                if confirm not in ("y", "yes", ""):
                    if verbose:
                        print("[CanvasMsg] Skipped.")
                    else:
                        print("Skipped.")
                    continue
            # Use the Conversation.add_message method to reply
            try:
                conv.add_message(body=msg)
                if verbose:
                    print("[CanvasMsg] Reply sent.")
                else:
                    print("Reply sent.")
            except Exception as e:
                if verbose:
                    print(f"[CanvasMsg] Failed to send reply: {e}")
                else:
                    print(f"Failed to send reply: {e}")
    except Exception as e:
        if verbose:
            print(f"[CanvasMsg] Error fetching or replying to messages: {e}")
        else:
            print(f"Error fetching or replying to messages: {e}")

