# -*- coding: utf-8 -*-

import os
import pickle

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


def _get_google_classroom_credentials(credentials_path, token_path, verbose=False):
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
            creds = flow.run_local_server(port=0)
            with open(token_path, "wb") as f:
                pickle.dump(creds, f)
    return creds


def list_google_classroom_courses(credentials_path, token_path, verbose=False):
    creds = _get_google_classroom_credentials(credentials_path, token_path, verbose=verbose)
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
