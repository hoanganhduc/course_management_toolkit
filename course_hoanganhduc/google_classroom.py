# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/
# Course Management Script

import pandas as pd
import os
import pickle
import readline
import sys
import argparse
import re
import glob
from datetime import datetime
import openpyxl
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import requests
import tempfile
import base64
import PyPDF2
import io
from collections import defaultdict
import difflib
from openpyxl.styles import Alignment
import unicodedata
import numpy as np
from paddleocr import PaddleOCR
import json
import shutil
from canvasapi import Canvas
from datetime import datetime, timedelta, timezone
from itertools import combinations
import signal
import platform
from pathlib import Path
import time
import traceback
import cv2
import subprocess
import csv
import copy
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import hashlib
import math
import statistics
from collections import Counter, OrderedDict
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from types import SimpleNamespace

from tqdm import tqdm  # <-- Add tqdm for progress bars

from .version import __version__

from .settings import *
from .config import *
from .models import *
from .utils import *
from .data import *

def sync_students_with_google_classroom(students, db_path=None, course_id=None, credentials_path='gclassroom_credentials.json', token_path='token.pickle', fetch_grades=False, verbose=False):
    """
    Sync students in the local database with active students from Google Classroom.
    For each student fetched from Google Classroom:
        - match by the local field 'Google Classroom Display Name' (case-insensitive)
        - if matched: fill missing local fields from Google data (Google_ID, Email, Google_Classroom_Display_Name)
        - if not matched: create a new student entry with Name, Email, Google_ID and Google_Classroom_Display_Name
    Optionally fetch grades/submission state when fetch_grades=True.

    Returns (added_count, updated_count).
    """
    SCOPES = [
        "https://www.googleapis.com/auth/classroom.courses.readonly",
        "https://www.googleapis.com/auth/classroom.rosters.readonly",
        "https://www.googleapis.com/auth/classroom.coursework.me.readonly",
        "https://www.googleapis.com/auth/classroom.student-submissions.students.readonly"
    ]

    try:
        # load current DB list if db_path provided
        if db_path and os.path.exists(db_path):
            try:
                students = load_database(db_path, verbose=verbose)
            except Exception:
                if verbose:
                    print("[GClassroom] Failed to load DB, proceeding with provided students list.")

        # ensure students is a list
        if students is None:
            students = []

        # helper to create Student object if no Student class defined
        try:
            Student  # type: ignore
            StudentClass = Student  # use existing Student class if defined
        except Exception:
            StudentClass = lambda **kw: SimpleNamespace(**kw)

        # authenticate
        creds = None
        if os.path.exists(token_path):
            try:
                with open(token_path, "rb") as f:
                    creds = pickle.load(f)
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
                # save token
                with open(token_path, "wb") as f:
                    pickle.dump(creds, f)

        service = build("classroom", "v1", credentials=creds)

        # if course_id not provided, list and ask user to select
        if not course_id:
            resp = service.courses().list(pageSize=50).execute()
            courses = resp.get("courses", []) or []
            if not courses:
                if verbose:
                    print("[GClassroom] No courses available for this account.")
                else:
                    print("No courses found.")
                return 0, 0
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
                        course_id = courses[idx]["id"]
                        break
                except Exception:
                    continue

        # fetch all students (handle pagination)
        classroom_students = []
        next_token = None
        while True:
            req = service.courses().students().list(courseId=course_id, pageToken=next_token, pageSize=200) if next_token else service.courses().students().list(courseId=course_id, pageSize=200)
            resp = req.execute()
            classroom_students.extend(resp.get("students", []) or [])
            next_token = resp.get("nextPageToken")
            if not next_token:
                break

        # Build lookup maps from local students for matching incoming records.
        local_by_google_name = {}
        local_by_name = {}
        local_by_email = {}
        # ensure keys are lowercased for robust matching
        for s in students:
            gname = getattr(s, "Google Classroom Display Name", "") or ""
            name = getattr(s, "Name", "") or ""
            email = getattr(s, "Email", "") or ""
            if isinstance(gname, str) and gname.strip():
                local_by_google_name[gname.strip().lower()] = s
            if isinstance(name, str) and name.strip():
                local_by_name[name.strip().lower()] = s
            if isinstance(email, str) and email.strip():
                local_by_email[email.strip().lower()] = s

        # Resolve duplicates by priority (display name > name > email). If multiple
        # candidates exist, prompt the operator to pick, create a new student, or skip.
        def _resolve_gc_match(name_key, email_key):
            candidates = []
            seen = set()

            def add_candidate(label, student):
                if id(student) in seen:
                    return
                candidates.append((label, student))
                seen.add(id(student))

            if name_key and name_key in local_by_google_name:
                add_candidate("google_name", local_by_google_name[name_key])
            if name_key and name_key in local_by_name:
                add_candidate("name", local_by_name[name_key])
            if email_key and email_key in local_by_email:
                add_candidate("email", local_by_email[email_key])

            if not candidates:
                return None
            if len(candidates) == 1:
                return candidates[0][1]

            print("\n[GClassroom] Possible duplicate match detected:")
            for idx, (label, student) in enumerate(candidates, 1):
                s_name = getattr(student, "Name", "") or ""
                s_email = getattr(student, "Email", "") or ""
                s_gid = getattr(student, "Google_ID", "") or ""
                print(f"{idx}. {s_name} | {s_email} | Google ID: {s_gid} (matched by {label})")
            print("n. Create new student")
            print("s. Skip this record")
            while True:
                choice = input("Choose a match (number), 'n' for new, or 's' to skip: ").strip().lower()
                if choice == "n":
                    return None
                if choice == "s":
                    return "__skip__"
                if choice.isdigit():
                    sel = int(choice) - 1
                    if 0 <= sel < len(candidates):
                        return candidates[sel][1]

        added_count = 0
        updated_count = 0

        for cs in classroom_students:
            profile = cs.get("profile", {}) or {}
            email = (profile.get("emailAddress") or "").strip()
            full_name = (profile.get("name", {}).get("fullName") or "").strip()
            google_id = cs.get("userId", "") or ""

            if not full_name:
                # skip unknown entries
                continue

            key_name = full_name.lower()
            matched = _resolve_gc_match(key_name, email.lower() if email else None)
            match_type = None
            if matched == "__skip__":
                continue
            if matched:
                if key_name in local_by_google_name and local_by_google_name[key_name] is matched:
                    match_type = "google_name"
                elif key_name in local_by_name and local_by_name[key_name] is matched:
                    match_type = "name"
                elif email and email.lower() in local_by_email and local_by_email[email.lower()] is matched:
                    match_type = "email"

            if matched:
                changed = False
                # Do not overwrite existing local Name if present; only fill missing fields
                if not getattr(matched, "Google_ID", "") and google_id:
                    matched.Google_ID = google_id
                    changed = True
                if not getattr(matched, "Email", "") and email:
                    matched.Email = email
                    changed = True
                if not getattr(matched, "Google Classroom Display Name", "") and full_name:
                    matched.Google_Classroom_Display_Name = full_name
                    changed = True
                if changed:
                    updated_count += 1
                    if verbose:
                        print(f"[GClassroom] Updated local student from GC: {full_name} ({match_type})")
            else:
                # create new student entry
                new_student = StudentClass(
                    Name=full_name,
                    Email=email,
                    Google_ID=google_id,
                    Google_Classroom_Display_Name=full_name
                )
                # Append to local list and update maps to prevent duplicate additions.
                students.append(new_student)
                local_by_google_name[key_name] = new_student
                local_by_name[key_name] = new_student
                if email:
                    local_by_email[email.lower()] = new_student
                added_count += 1
                if verbose:
                    print(f"[GClassroom] Added new student: {full_name} ({email})")

        # optionally fetch coursework and grades (if requested)
        if fetch_grades:
            if verbose:
                print("[GClassroom] Fetching coursework and student submission data...")
            # fetch all coursework (paginated)
            coursework = []
            next_token = None
            while True:
                req = service.courses().courseWork().list(courseId=course_id, pageToken=next_token, pageSize=200) if next_token else service.courses().courseWork().list(courseId=course_id, pageSize=200)
                resp = req.execute()
                coursework.extend(resp.get("courseWork", []) or [])
                next_token = resp.get("nextPageToken")
                if not next_token:
                    break

            # for each local student with Google_ID, fetch submissions per coursework
            for s in students:
                gid = getattr(s, "Google_ID", "") or ""
                if not gid:
                    continue
                grades = getattr(s, "Grades", {}) or {}
                submissions = getattr(s, "Submissions", {}) or {}
                for cw in coursework:
                    cw_id = cw.get("id")
                    if not cw_id:
                        continue
                    title = cw.get("title", f"cw_{cw_id}")
                    try:
                        # list studentSubmissions filtered by userId
                        resp = service.courses().courseWork().studentSubmissions().list(
                            courseId=course_id, courseWorkId=cw_id, userId=gid, pageSize=50
                        ).execute()
                        subs = resp.get("studentSubmissions", []) or []
                        if subs:
                            sub = subs[0]
                            state = sub.get("state")
                            # assignedGrade may be under "assignedGrade" or in assignedGrade field
                            grade = sub.get("assignedGrade")
                            # maxPoints often available on coursework object
                            max_points = cw.get("maxPoints")
                            if grade is not None:
                                grades[title] = {"grade": grade, "max_points": max_points}
                            submissions[title] = state
                    except Exception:
                        # ignore per-assignment errors but continue
                        if verbose:
                            print(f"[GClassroom] Warning: could not fetch submission for student {getattr(s,'Name','?')} cw:{title}")
                s.Grades = grades
                s.Submissions = submissions
            if verbose:
                print("[GClassroom] Grades/submissions fetch complete.")

        # save back to db if requested
        if db_path:
            try:
                save_database(students, db_path, verbose=verbose)
            except Exception as e:
                if verbose:
                    print(f"[GClassroom] Warning: failed to save DB: {e}")

        if verbose:
            print(f"[GClassroom] Sync finished. Added: {added_count}, Updated: {updated_count}")
        else:
            print(f"Sync completed: {added_count} added, {updated_count} updated")

        return added_count, updated_count

    except Exception as e:
        # handle auth errors specially: if token exists and auth failure, remove token to force reauth
        try:
            err_status = None
            if hasattr(e, "resp"):
                err_status = getattr(e.resp, "status", None)
            if err_status == 401 and os.path.exists(token_path):
                try:
                    os.remove(token_path)
                except Exception:
                    pass
                if verbose:
                    print("[GClassroom] Authentication failed; token removed. Re-run to re-authenticate.")
                else:
                    print("Authentication failed. Please re-run to re-authenticate.")
            else:
                if verbose:
                    print(f"[GClassroom] Error during sync: {e}")
                else:
                    print(f"Error syncing with Google Classroom: {e}")
        except Exception:
            if verbose:
                print(f"[GClassroom] Error during exception handling: {e}")
            else:
                print("Error syncing with Google Classroom.")
        return 0, 0
