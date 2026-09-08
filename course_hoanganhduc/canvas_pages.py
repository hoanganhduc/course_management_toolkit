# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas page helpers."""

import os
import signal
import tempfile

import requests
from .canvas_auth import get_canvas_client

from .data import refine_text_with_ai
from .canvas_utils import prompt_with_timeout, parse_selection, normalize_datetime_str
from .settings import (
    ALL_AI_METHODS,
    CANVAS_LMS_API_KEY,
    CANVAS_LMS_API_URL,
    CANVAS_LMS_COURSE_ID,
    DEFAULT_AI_METHOD,
)

def list_and_update_canvas_pages(
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    refine=DEFAULT_AI_METHOD,
    verbose=False
):
    """
    List all pages in a Canvas course and allow the user to update or delete one or more pages.
    For each selected page, fetch the content to a temporary file in the current folder, open that file for editing,
    and after editing, re-upload the content to Canvas using the official API (PUT /v1/courses/:course_id/pages/:url_or_id).
    Also allows deleting a page. Optionally refine the new content using AI before updating.
    If no response from user within 60 seconds, use default option.
    If verbose is True, print more details; otherwise, print only important notice.
    """

    def timeout_handler(_signum, _frame):
        if verbose:
            print("\n[CanvasPages] Timeout: No response after 60 seconds. Using default option.")
        else:
            print("\nTimeout: No response after 60 seconds. Using default option.")
        raise TimeoutError("User input timeout")

    def _prompt(prompt, timeout=60, default=None):
        return prompt_with_timeout(prompt, timeout=timeout, default=default, verbose=verbose)

    try:
        # Ensure api_url is valid
        if not api_url or not isinstance(api_url, str) or not api_url.startswith(("http://", "https://")):
            if verbose:
                print(f"[CanvasPages] Error: Canvas API URL is missing or invalid. Please provide a valid HTTP or HTTPS URL for the Canvas instance. {api_url}")
            else:
                print(f"Error: Canvas API URL is missing or invalid.")
            return
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)
        pages = list(course.get_pages())
        if not pages:
            if verbose:
                print("[CanvasPages] No pages found in this course.")
            else:
                print("No pages found in this course.")
            return
        # Sort pages by title
        pages.sort(key=lambda p: p.title.lower())
        if verbose:
            print("[CanvasPages] Pages in this course:")
        else:
            print("Pages in this course:")
        for idx, page in enumerate(pages, 1):
            if verbose:
                print(f"[CanvasPages] {idx}. {page.title} (url: {page.url})")
            else:
                print(f"{idx}. {page.title} (url: {page.url})")
        # Prompt user to select pages to update/delete
        sel = _prompt(
            "Enter page numbers to update/delete (e.g. 1,3-5 or 'a' for all, or 'q' to quit): ",
            timeout=60,
            default="q"
        ).strip()
        if sel.lower() in ('q', 'quit'):
            return
        if sel.lower() in ('a', 'all'):
            selected = list(range(1, len(pages) + 1))
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
            selected = [i for i in selected if 1 <= i <= len(pages)]
        if not selected:
            if verbose:
                print("[CanvasPages] No valid selection.")
            else:
                print("No valid selection.")
            return
        for idx in selected:
            page = pages[idx - 1]
            if verbose:
                print(f"\n[CanvasPages] --- Page {idx}: {page.title} ---")
            else:
                print(f"\n--- Page {idx}: {page.title} ---")
            full_page = course.get_page(page.url)
            old_body = full_page.body or ""
            if verbose:
                print("[CanvasPages] Current content (truncated to 500 chars):")
                print(old_body[:500])
            else:
                print("Current content (truncated to 500 chars):")
                print(old_body[:500])
            print("\nOptions: [e]dit, [d]elete, [s]kip")
            action = _prompt(
                "What do you want to do with this page? (e/d/s): ",
                timeout=60,
                default="s"
            ).strip().lower()
            if action == "d":
                confirm = _prompt(
                    f"Are you sure you want to delete page '{page.title}'? (y/n): ",
                    timeout=60,
                    default="n"
                ).strip().lower()
                if confirm == "y":
                    try:
                        full_page.delete()
                        if verbose:
                            print(f"[CanvasPages] Page '{page.title}' deleted successfully.")
                        else:
                            print(f"Page '{page.title}' deleted successfully.")
                    except Exception as e:
                        if verbose:
                            print(f"[CanvasPages] Failed to delete page '{page.title}': {e}")
                        else:
                            print(f"Failed to delete page '{page.title}': {e}")
                else:
                    if verbose:
                        print("[CanvasPages] Skipped deleting this page.")
                    else:
                        print("Skipped deleting this page.")
                continue
            elif action == "s":
                if verbose:
                    print("[CanvasPages] Skipped this page.")
                else:
                    print("Skipped this page.")
                continue
            elif action != "e":
                if verbose:
                    print("[CanvasPages] Unknown action, skipping this page.")
                else:
                    print("Unknown action, skipping this page.")
                continue

            # Write old content to a temp file in the current folder
            safe_title = "".join(c if c.isalnum() or c in "-_." else "_" for c in page.title)[:40]
            temp_filename = f"canvas_page_{safe_title}_{page.url}.html"
            temp_path = os.path.join(os.getcwd(), temp_filename)
            with open(temp_path, "w", encoding="utf-8") as f:
                f.write(old_body)
            if verbose:
                print(f"\n[CanvasPages] Edit the content for this page in your editor: {temp_path}")
                print("[CanvasPages] After saving and closing the editor, the content will be uploaded to Canvas.")
            else:
                print(f"\nEdit the content for this page in your editor: {temp_path}")
                print("After saving and closing the editor, the content will be uploaded to Canvas.")
            default_editor = "notepad" if os.name == "nt" else ("nano" if shutil.which("nano") else "vi")
            editor = os.environ.get("EDITOR", default_editor)
            try:
                subprocess.call([editor, temp_path])
            except Exception as e:
                if verbose:
                    print(f"[CanvasPages] Could not open editor: {e}")
                else:
                    print(f"Could not open editor: {e}")
                continue

            # Read the edited content
            try:
                if verbose:
                    print(f"[CanvasPages] Reading edited content from: {temp_path}")
                else:
                    print(f"Reading edited content from: {temp_path}")
                with open(temp_path, "r", encoding="utf-8") as f:
                    new_body = f.read()
            except Exception as e:
                if verbose:
                    print(f"[CanvasPages] Could not read edited file: {e}")
                else:
                    print(f"Could not read edited file: {e}")
                continue

            # Ask user if they want to refine the new content using AI
            refine_choice = None
            if refine is None:
                try:
                    refine_choice = _prompt(
                        "Do you want to refine the page content using AI? (none/gemini/huggingface/local) [none]: ",
                        timeout=60,
                        default="none"
                    ).strip().lower()
                except TimeoutError:
                    if verbose:
                        print("[CanvasPages] No response after 60 seconds. Using default: none")
                    else:
                        print("No response after 60 seconds. Using default: none")
                    refine_choice = "none"
                if refine_choice in ALL_AI_METHODS:
                    refine = refine_choice
                else:
                    refine = None
            if refine in ALL_AI_METHODS:
                if verbose:
                    print(f"[CanvasPages] Refining page content using AI service: {refine} ...")
                else:
                    print(f"Refining page content using AI service: {refine} ...")
                prompt = (
                    "You are an expert assistant. Please refine the following Canvas page content for clarity, "
                    "correct any spelling or grammar mistakes, and improve readability. "
                    "Keep the meaning and important information unchanged. "
                    "Return ONLY the improved content, no explanations.\n\n"
                    "Content:\n{text}"
                )
                refined_body = refine_text_with_ai(new_body, method=refine, user_prompt=prompt)
                if verbose:
                    print("\n[CanvasPages] Refined content preview (truncated to 500 chars):\n")
                    print(refined_body[:500])
                else:
                    print("\nRefined content preview (truncated to 500 chars):\n")
                    print(refined_body[:500])
                confirm = _prompt(
                    "Use refined content? (y/n): ",
                    timeout=60,
                    default="y"
                ).strip().lower()
                if confirm == "y":
                    new_body = refined_body

            # Ask if user wants to update the title
            update_title = _prompt(
                f"Do you want to update the page title? (current: '{page.title}') (y/n): ",
                timeout=60,
                default="n"
            ).strip().lower()
            new_title = page.title
            if update_title == "y":
                t = _prompt(
                    "Enter new title (leave blank to keep current): ",
                    timeout=60,
                    default=""
                ).strip()
                if t:
                    new_title = t

            # Ask for optional parameters
            editing_roles = _prompt(
                "Enter editing roles (comma-separated, e.g. teachers,students) or leave blank to keep unchanged: ",
                timeout=60,
                default=""
            ).strip()
            notify_of_update = _prompt(
                "Notify participants of update? (y/n, default n): ",
                timeout=60,
                default="n"
            ).strip().lower()
            published = _prompt(
                "Publish page? (y/n, default y): ",
                timeout=60,
                default="y"
            ).strip().lower()
            front_page = _prompt(
                "Set as front page? (y/n, default n): ",
                timeout=60,
                default="n"
            ).strip().lower()

            # Prepare API request
            url = f"{api_url}/api/v1/courses/{course_id}/pages/{page.url}"
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            }
            data = {
                "wiki_page": {
                    "title": new_title,
                    "body": new_body
                }
            }
            if editing_roles:
                data["wiki_page"]["editing_roles"] = editing_roles
            if notify_of_update == "y":
                data["wiki_page"]["notify_of_update"] = True
            if published == "n":
                data["wiki_page"]["published"] = False
            else:
                data["wiki_page"]["published"] = True
            if front_page == "y":
                data["wiki_page"]["front_page"] = True

            # Send PUT request to update the page
            try:
                resp = requests.put(url, headers=headers, json=data)
                if resp.status_code in (200, 201):
                    if verbose:
                        print(f"[CanvasPages] Page '{new_title}' updated successfully.")
                    else:
                        print(f"Page '{new_title}' updated successfully.")
                else:
                    if verbose:
                        print(f"[CanvasPages] Failed to update page '{new_title}': {resp.status_code} {resp.text}")
                    else:
                        print(f"Failed to update page '{new_title}': {resp.status_code} {resp.text}")
            except Exception as e:
                if verbose:
                    print(f"[CanvasPages] Failed to update page '{new_title}': {e}")
                else:
                    print(f"Failed to update page '{new_title}': {e}")

            # Remove the temporary file after updating
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                    if verbose:
                        print(f"[CanvasPages] Temporary file {temp_path} removed.")
                    else:
                        print(f"Temporary file {temp_path} removed.")
            except Exception as e:
                if verbose:
                    print(f"[CanvasPages] Could not remove temporary file {temp_path}: {e}")
                else:
                    print(f"Could not remove temporary file {temp_path}: {e}")

    except Exception as e:
        if verbose:
            print(f"[CanvasPages] Error listing or updating Canvas pages: {e}")
        else:
            print(f"Error listing or updating Canvas pages: {e}")

