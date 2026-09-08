# -*- coding: utf-8 -*-

import getpass
import os
import pickle
import webbrowser
from urllib.parse import parse_qsl, urlsplit, urlunsplit

from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = [
    "https://www.googleapis.com/auth/classroom.courses",
    "https://www.googleapis.com/auth/classroom.rosters",
    "https://www.googleapis.com/auth/classroom.coursework.students",
    "https://www.googleapis.com/auth/classroom.topics.readonly",
    "https://www.googleapis.com/auth/classroom.profile.emails",
    "https://www.googleapis.com/auth/classroom.profile.photos",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]


def _browser_available():
    """Whether this machine has a browser ``run_local_server`` could open.

    It calls ``webbrowser.get(browser).open(...)`` before it prints the
    authorization URL, and ``webbrowser.get`` raises when nothing is registered,
    which is the normal state on a headless server. Asking first turns a crash
    that hides the URL into a printed URL the operator can open elsewhere.
    """
    try:
        webbrowser.get()
    except Exception:
        return False
    return True


# Nothing listens on this port. The browser is meant to fail to reach it so that
# the operator can copy the failed URL out of the address bar; Google accepts any
# loopback port for an installed-app client, so the number itself is arbitrary.
PASTE_REDIRECT_URI = "http://127.0.0.1:8910/"

_MAX_REDIRECT_URL_CHARS = 8192
_PASTE_ATTEMPTS = 3


def _authorization_response(raw):
    """Turn a pasted browser redirect into the response ``fetch_token`` accepts.

    Returns ``(url, None)`` when the paste is usable, or ``(None, reason)``. The
    scheme is rewritten to https because oauthlib refuses to read an OAuth
    response over http; only the query is read from it, and the code is still
    exchanged against Google's own token endpoint.
    """
    if not isinstance(raw, str):
        return None, "nothing was pasted"
    url = raw.strip()
    if not url:
        return None, "nothing was pasted"
    if len(url) > _MAX_REDIRECT_URL_CHARS:
        return None, "the pasted URL is too long"
    if any(ord(char) < 32 or ord(char) == 127 for char in url):
        return None, "the pasted URL contains control characters"
    try:
        parsed = urlsplit(url)
        pairs = parse_qsl(parsed.query, keep_blank_values=True, max_num_fields=64)
    except ValueError:
        return None, "the pasted URL could not be read"
    if (
        parsed.scheme not in ("http", "https")
        or parsed.hostname not in ("127.0.0.1", "localhost")
        or parsed.username is not None
        or parsed.password is not None
    ):
        return None, f"that is not the {PASTE_REDIRECT_URI} redirect"
    values = {}
    for key, value in pairs:
        values.setdefault(key, []).append(value)
    if values.get("error"):
        return None, f"Google reported '{values['error'][0]}' instead of a code"
    if len(values.get("code", [])) != 1 or not values["code"][0]:
        return None, "the pasted URL carries no authorization code"
    if len(values.get("state", [])) != 1 or not values["state"][0]:
        return None, "the pasted URL carries no state value"
    return urlunsplit(("https", parsed.netloc, parsed.path or "/", parsed.query, "")), None


def _authorize_by_paste(flow, input_fn=None, print_fn=None):
    """Complete consent in this one terminal, on a machine with no browser.

    ``run_local_server`` blocks inside ``handle_request()`` until a browser
    reaches its loopback port, so the process that prints the URL cannot also ask
    for the redirect, and finishing it needed a second terminal. Setting the
    redirect URI here and exchanging a pasted response instead keeps the whole
    flow in the terminal the operator already has open.
    """
    input_fn = getpass.getpass if input_fn is None else input_fn
    print_fn = print if print_fn is None else print_fn

    flow.redirect_uri = PASTE_REDIRECT_URI
    auth_url, _ = flow.authorization_url(access_type="offline")
    print_fn("[GClassroom] No browser on this machine. Finish authorization like this:")
    print_fn("[GClassroom]   1. Open the URL below on a machine that has a browser.")
    print_fn("[GClassroom]   2. Sign in and grant access.")
    print_fn(f"[GClassroom]   3. The browser then fails to load {PASTE_REDIRECT_URI}...")
    print_fn("[GClassroom]      That failure is expected; the URL is what matters.")
    print_fn("[GClassroom]   4. Copy that failed URL from the address bar and paste it here.")
    print_fn("")
    print_fn(auth_url)
    print_fn("")

    for remaining in range(_PASTE_ATTEMPTS - 1, -1, -1):
        try:
            raw = input_fn("Paste the redirect URL here (input hidden): ")
        except (EOFError, KeyboardInterrupt, StopIteration):
            raise RuntimeError(
                "Google authorization was cancelled before a redirect URL was pasted"
            ) from None
        response, reason = _authorization_response(raw)
        if response:
            try:
                flow.fetch_token(authorization_response=response)
            except Exception as exc:
                # A rejected exchange is not a typo the operator can fix by
                # pasting again: the code is single-use and short-lived, so the
                # way back is a fresh authorization URL.
                raise RuntimeError(f"Google rejected the pasted redirect: {exc}") from exc
            return flow.credentials
        if remaining:
            print_fn(f"[GClassroom] Cannot use that paste: {reason}. Try again.")
    raise RuntimeError(
        "Google authorization did not complete: no usable redirect URL was pasted"
    )


def _get_google_classroom_credentials(credentials_path, token_path, verbose=False, open_browser=None):
    creds = None
    if os.path.exists(token_path):
        try:
            with open(token_path, "rb") as f:
                creds = pickle.load(f)
        except Exception:
            creds = None
    if creds and getattr(creds, "scopes", None):
        try:
            if set(creds.scopes) != set(SCOPES):
                creds = None
                if os.path.exists(token_path):
                    try:
                        os.remove(token_path)
                    except Exception:
                        pass
                if verbose:
                    print("[GClassroom] Stored token scopes do not match required scopes; token removed.")
        except Exception:
            creds = None

    if not creds or not getattr(creds, "valid", False):
        if creds and getattr(creds, "expired", False) and getattr(creds, "refresh_token", None):
            try:
                creds.refresh(Request())
            except Exception:
                creds = None
                if os.path.exists(token_path):
                    try:
                        os.remove(token_path)
                    except Exception:
                        pass
        if not creds:
            if not os.path.exists(credentials_path):
                raise FileNotFoundError(f"Google credentials not found: {credentials_path}")
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            use_browser = _browser_available() if open_browser is None else bool(open_browser)
            if use_browser:
                creds = flow.run_local_server(host="127.0.0.1", port=0, open_browser=True)
            else:
                creds = _authorize_by_paste(flow)
            with open(token_path, "wb") as f:
                pickle.dump(creds, f)
    return creds


def list_google_classroom_courses(credentials_path, token_path, verbose=False, open_browser=None):
    creds = _get_google_classroom_credentials(
        credentials_path, token_path, verbose=verbose, open_browser=open_browser
    )
    service = build("classroom", "v1", credentials=creds)
    courses = []
    next_token = None
    while True:
        req = service.courses().list(pageToken=next_token, pageSize=50) if next_token else service.courses().list(pageSize=50)
        resp = req.execute()
        courses.extend(resp.get("courses", []) or [])
        next_token = resp.get("nextPageToken")
        if not next_token:
            break
    return courses


def list_google_classroom_students(credentials_path, token_path, course_id=None, verbose=False, open_browser=None):
    creds = _get_google_classroom_credentials(
        credentials_path, token_path, verbose=verbose, open_browser=open_browser
    )
    service = build("classroom", "v1", credentials=creds)

    if not course_id:
        courses = list_google_classroom_courses(
            credentials_path, token_path, verbose=verbose, open_browser=open_browser
        )
        if not courses:
            print("No courses found.")
            return []
        print("Available Google Classroom courses:")
        for i, c in enumerate(courses, 1):
            print(f"{i}. {c.get('name')} (ID: {c.get('id')})")
        while True:
            sel = input("Select course number (or 'q' to quit): ").strip().lower()
            if sel in ("q", "quit"):
                return []
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
        return []

    students = []
    next_token = None
    while True:
        req = service.courses().students().list(courseId=course_id, pageToken=next_token, pageSize=200) if next_token else service.courses().students().list(courseId=course_id, pageSize=200)
        resp = req.execute()
        students.extend(resp.get("students", []) or [])
        next_token = resp.get("nextPageToken")
        if not next_token:
            break
    return students
