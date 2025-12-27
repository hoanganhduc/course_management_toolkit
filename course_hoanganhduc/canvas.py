# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/
# Course Management Script

"""Canvas LMS helpers."""

import pandas as pd
import os
import pickle
try:
    import readline
except ImportError:  # Windows without pyreadline installed
    readline = None
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

def send_final_evaluations_via_canvas(
    final_dir="final_evaluations",
    db_path=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
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
        canvas = Canvas(api_url, api_key)
    except Exception as e:
        if verbose:
            print(f"[SendFinals] Failed to initialize Canvas client: {e}")
        else:
            print("Failed to initialize Canvas client.")
        return {"sent": 0, "skipped": 0, "errors": 0}

    sent = 0
    skipped = 0
    errors = 0

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

        body = (
            f"{greeting}\n\n"
            f"{content}\n\n"
        )

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
    if verbose:
        print(f"[SendFinals] Completed. Sent: {sent}, Skipped: {skipped}, Errors: {errors}")
    else:
        print(f"Done. Sent: {sent}, Skipped: {skipped}, Errors: {errors}")
    return summary

def sync_students_with_canvas(students, db_path=None, course_id=None, api_url=CANVAS_LMS_API_URL, api_key=CANVAS_LMS_API_KEY, verbose=False):
    """
    Sync students in the local database with active students from Canvas course.
    Adds new students from Canvas if not present, updates Canvas ID for existing students,
    and syncs scores of both total course grade and each assignment category.

    Args:
        students: List of Student objects
        db_path: Path to save the database
        course_id: Canvas course ID (uses default if None)
        api_url: Canvas API URL
        api_key: Canvas API key
        verbose: If True, print more details; otherwise, print only important notice

    Returns:
        (added_count, updated_count): Counts of students added and updated
    """
    try:
        if course_id is None:
            course_id = CANVAS_LMS_COURSE_ID
        if verbose:
            print(f"[SyncCanvas] Fetching students from Canvas course {course_id}...")
        else:
            print("Syncing students with Canvas course...")
        people = list_canvas_people(api_url=api_url, api_key=api_key, course_id=course_id)
        canvas_students = people.get("active_students", [])
        if not canvas_students:
            if verbose:
                print("[SyncCanvas] No active students found in Canvas course.")
            else:
                print("No active students found in Canvas course.")
            return 0, 0

        if verbose:
            print(f"[SyncCanvas] Found {len(canvas_students)} active students in Canvas course.")

        # Helper to normalize names for comparison
        def normalize_name(name):
            if not name:
                return ""
            name = str(name)
            name = re.sub(r"[^a-zA-Z0-9 ]", "", name)
            name = re.sub(r"\s+", " ", name)
            return name.strip().lower()

        # Build lookups for matching Canvas records to local students.
        existing_by_email = {}
        existing_by_name = {}
        existing_by_canvas_id = {}

        if verbose:
            print("[SyncCanvas] Building lookup tables for existing students...")
        for s in students:
            email = getattr(s, "Email", None)
            name = getattr(s, "Name", None)
            canvas_id = getattr(s, "Canvas ID", None)

            if email:
                existing_by_email[email.lower()] = s
            if name:
                norm_name = normalize_name(name)
                if norm_name:
                    existing_by_name[norm_name] = s
            if canvas_id:
                try:
                    existing_by_canvas_id[int(canvas_id)] = s
                except (ValueError, TypeError):
                    pass

        # Resolve duplicates by priority (Canvas ID > name > email). If multiple candidates
        # exist, prompt the operator to pick, create a new student, or skip.
        def _resolve_canvas_match(canvas_id_value, name_key, email_key):
            candidates = []
            seen = set()

            def add_candidate(label, student):
                if id(student) in seen:
                    return
                candidates.append((label, student))
                seen.add(id(student))

            if canvas_id_value:
                try:
                    cid = int(canvas_id_value)
                    if cid in existing_by_canvas_id:
                        add_candidate("canvas_id", existing_by_canvas_id[cid])
                except (ValueError, TypeError):
                    pass
            if name_key and name_key in existing_by_name:
                add_candidate("name", existing_by_name[name_key])
            if email_key and email_key in existing_by_email:
                add_candidate("email", existing_by_email[email_key])

            if not candidates:
                return None
            if len(candidates) == 1:
                return candidates[0][1]

            print("\n[SyncCanvas] Possible duplicate match detected:")
            for idx, (label, student) in enumerate(candidates, 1):
                s_name = getattr(student, "Name", "") or ""
                s_email = getattr(student, "Email", "") or ""
                s_cid = getattr(student, "Canvas ID", "") or ""
                print(f"{idx}. {s_name} | {s_email} | Canvas ID: {s_cid} (matched by {label})")
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

        # Prepare Canvas API for grades and scores
        if verbose:
            print("[SyncCanvas] Connecting to Canvas API...")
        canvas = Canvas(api_url, api_key)
        course = canvas.get_course(course_id)

        # Fetch enrollments to get total grades
        if verbose:
            print("[SyncCanvas] Fetching enrollments for total grades...")
        enrollments = list(course.get_enrollments(type=['StudentEnrollment']))

        # Build a map: canvas_id -> total_grade info.
        total_grades_by_canvas_id = {}
        for enrollment in enrollments:
            user_id = getattr(enrollment, "user_id", None)
            if user_id and hasattr(enrollment, "grades"):
                grades = enrollment.grades
                total_grades_by_canvas_id[user_id] = {
                    "final_score": grades.get("final_score"),
                    "final_grade": grades.get("final_grade"),
                    "current_score": grades.get("current_score"),
                    "current_grade": grades.get("current_grade")
                }

        # Fetch assignment groups for category scores
        if verbose:
            print("[SyncCanvas] Fetching assignment groups and scores...")
        try:
            headers = {'Authorization': f'Bearer {api_key}'}
            group_scores_url = f"{api_url}/api/v1/courses/{course_id}/students/submissions?student_ids[]=all&include[]=assignment&include[]=user&include[]=score&grouped=true"
            response = requests.get(group_scores_url, headers=headers)
            response.raise_for_status()
            category_scores_data = response.json()
        except Exception as e:
            if verbose:
                print(f"[SyncCanvas] Warning: Could not fetch detailed category scores: {e}")
            else:
                print("Warning: Could not fetch detailed category scores.")
            category_scores_data = []

        assignment_groups = list(course.get_assignment_groups(include=['assignments', 'score_statistics']))

        # Build a map: canvas_id -> {group_name: {current_score, current_possible, final_score, final_possible}}.
        group_scores_by_canvas_id = {}

        # First, process score_statistics (current scores)
        for group in assignment_groups:
            group_name = group.name
            group_id = group.id
            stats = getattr(group, "score_statistics", {})
            for canvas_id_str, stat in stats.items():
                try:
                    canvas_id = int(canvas_id_str)
                except Exception:
                    continue
                if canvas_id not in group_scores_by_canvas_id:
                    group_scores_by_canvas_id[canvas_id] = {}
                if group_name not in group_scores_by_canvas_id[canvas_id]:
                    group_scores_by_canvas_id[canvas_id][group_name] = {
                        'current_score': 0,
                        'current_possible': 0,
                        'final_score': 0,
                        'final_possible': 0,
                        'group_id': group_id
                    }

                group_scores_by_canvas_id[canvas_id][group_name]['current_score'] = stat.get("score", 0)
                group_scores_by_canvas_id[canvas_id][group_name]['current_possible'] = stat.get("possible", 0)
                if group_scores_by_canvas_id[canvas_id][group_name]['final_score'] == 0:
                    group_scores_by_canvas_id[canvas_id][group_name]['final_score'] = stat.get("score", 0)
                if group_scores_by_canvas_id[canvas_id][group_name]['final_possible'] == 0:
                    group_scores_by_canvas_id[canvas_id][group_name]['final_possible'] = stat.get("possible", 0)

        # Then process submissions data to get final scores aggregated per assignment group.
        if verbose:
            print("[SyncCanvas] Processing category final scores...")
        try:
            assignments_by_group = {}
            for group in assignment_groups:
                assignments_by_group[group.id] = []
                for assignment in getattr(group, "assignments", []):
                    if assignment.get("published", False):
                        assignments_by_group[group.id].append(assignment['id'])

            for user_data in tqdm(category_scores_data, desc="Processing category final scores"):
                user_id = user_data.get("user_id")
                if not user_id:
                    continue

                if user_id not in group_scores_by_canvas_id:
                    group_scores_by_canvas_id[user_id] = {}

                for submission in user_data.get("submissions", []):
                    assignment = submission.get("assignment")
                    if not assignment:
                        continue

                    group_id = assignment.get("assignment_group_id")
                    group = next((g for g in assignment_groups if g.id == group_id), None)
                    if not group:
                        continue

                    group_name = group.name
                    if group_name not in group_scores_by_canvas_id[user_id]:
                        group_scores_by_canvas_id[user_id][group_name] = {
                            'current_score': 0,
                            'current_possible': 0,
                            'final_score': 0,
                            'final_possible': 0,
                            'group_id': group_id
                        }

                    score = submission.get("score", 0) or 0
                    points_possible = assignment.get("points_possible", 0) or 0

                    group_scores_by_canvas_id[user_id][group_name]['final_score'] += score
                    group_scores_by_canvas_id[user_id][group_name]['final_possible'] += points_possible
        except Exception as e:
            if verbose:
                print(f"[SyncCanvas] Warning: Error processing category final scores: {e}")
            else:
                print("Warning: Error processing category final scores.")

        if verbose:
            print("[SyncCanvas] Fetching individual assignments...")
        important_assignments = {}
        try:
            for group in assignment_groups:
                for assignment in getattr(group, "assignments", []):
                    if assignment.get("published", False):
                        important_assignments[assignment['id']] = {
                            "name": assignment['name'],
                            "points_possible": assignment['points_possible'],
                            "group_name": group.name
                        }
        except Exception as e:
            if verbose:
                print(f"[SyncCanvas] Warning: Could not fetch assignments: {e}")
            else:
                print("Warning: Could not fetch assignments.")

        assignment_scores_by_canvas_id = {}
        if important_assignments:
            if verbose:
                print(f"[SyncCanvas] Fetching submissions for {len(important_assignments)} assignments...")
            try:
                for assignment_id in list(important_assignments.keys()):
                    assignment = course.get_assignment(assignment_id)
                    assignment_info = important_assignments[assignment_id]

                    for submission in tqdm(assignment.get_submissions(include=["user"]),
                                         desc=f"Processing {assignment_info['name']} submissions"):
                        user_id = getattr(submission, "user_id", None)
                        score = getattr(submission, "score", None)
                        if user_id and score is not None:
                            if user_id not in assignment_scores_by_canvas_id:
                                assignment_scores_by_canvas_id[user_id] = {}
                            assignment_scores_by_canvas_id[user_id][assignment_id] = {
                                "name": assignment_info["name"],
                                "group": assignment_info["group_name"],
                                "score": score,
                                "points_possible": assignment_info["points_possible"]
                            }
            except Exception as e:
                if verbose:
                    print(f"[SyncCanvas] Warning: Error fetching assignment submissions: {e}")
                else:
                    print("Warning: Error fetching assignment submissions.")

        # Merge Canvas roster and grades into the local database.
        if verbose:
            print("[SyncCanvas] Syncing student data...")
        else:
            print("Syncing student data...")
        for canvas_student in tqdm(canvas_students, desc="Syncing students"):
            canvas_email = canvas_student.get("email", "").lower()
            canvas_name = canvas_student.get("name", "")
            canvas_id = canvas_student.get("canvas_id", "")

            norm_canvas_name = normalize_name(canvas_name)
            matched_student = _resolve_canvas_match(canvas_id, norm_canvas_name, canvas_email)
            if matched_student == "__skip__":
                continue

            if matched_student:
                changed = False

                if not hasattr(matched_student, "Canvas ID") or not getattr(matched_student, "Canvas ID"):
                    setattr(matched_student, "Canvas ID", canvas_id)
                    changed = True

                if canvas_email and (not hasattr(matched_student, "Email") or not getattr(matched_student, "Email")):
                    setattr(matched_student, "Email", canvas_email)
                    changed = True

                if canvas_id and total_grades_by_canvas_id.get(int(canvas_id)):
                    grade_data = total_grades_by_canvas_id[int(canvas_id)]
                    for grade_field, grade_value in grade_data.items():
                        if grade_value is not None:
                            field_name = {
                                "final_score": "Total Final Score",
                                "final_grade": "Total Final Grade",
                            }.get(grade_field)
                            if field_name:
                                if getattr(matched_student, field_name, None) != grade_value:
                                    setattr(matched_student, field_name, grade_value)
                                    changed = True

                if canvas_id and group_scores_by_canvas_id.get(int(canvas_id)):
                    for group_name, scores_data in group_scores_by_canvas_id[int(canvas_id)].items():
                        final_score = scores_data.get('final_score', 0)
                        final_field = f"{group_name} Final Score"
                        if getattr(matched_student, final_field, None) != final_score:
                            setattr(matched_student, final_field, final_score)
                            changed = True

                if canvas_id and assignment_scores_by_canvas_id.get(int(canvas_id)):
                    for assignment_id, assignment_data in assignment_scores_by_canvas_id[int(canvas_id)].items():
                        name = assignment_data["name"]
                        score = assignment_data["score"]
                        field = f"Assignment: {name}"
                        if getattr(matched_student, field, None) != score:
                            setattr(matched_student, field, score)
                            changed = True

                if changed:
                    updated_count += 1
            else:
                new_student_data = {
                    "Name": canvas_name,
                    "Email": canvas_email,
                    "Canvas ID": canvas_id
                }

                if canvas_id and total_grades_by_canvas_id.get(int(canvas_id)):
                    grade_data = total_grades_by_canvas_id[int(canvas_id)]
                    for grade_field, grade_value in grade_data.items():
                        if grade_value is not None:
                            field_name = {
                                "final_score": "Total Final Score",
                                "final_grade": "Total Final Grade",
                            }.get(grade_field)
                            if field_name:
                                new_student_data[field_name] = grade_value

                if canvas_id and group_scores_by_canvas_id.get(int(canvas_id)):
                    for group_name, scores_data in group_scores_by_canvas_id[int(canvas_id)].items():
                        new_student_data[f"{group_name} Final Score"] = scores_data.get('final_score', 0)

                if canvas_id and assignment_scores_by_canvas_id.get(int(canvas_id)):
                    for assignment_id, assignment_data in assignment_scores_by_canvas_id[int(canvas_id)].items():
                        name = assignment_data["name"]
                        score = assignment_data["score"]
                        field = f"Assignment: {name}"
                        new_student_data[field] = score

                students.append(Student(**new_student_data))
                if canvas_id:
                    try:
                        existing_by_canvas_id[int(canvas_id)] = students[-1]
                    except (ValueError, TypeError):
                        pass
                if canvas_email:
                    existing_by_email[canvas_email] = students[-1]
                if norm_canvas_name:
                    existing_by_name[norm_canvas_name] = students[-1]
                added_count += 1

        if added_count > 0 or updated_count > 0:
            if db_path:
                if verbose:
                    print(f"[SyncCanvas] Saving updated database with {added_count} new and {updated_count} modified students...")
                else:
                    print(f"Saving updated database with {added_count} new and {updated_count} modified students...")
                save_database(students, db_path)
                if verbose:
                    print("[SyncCanvas] Database saved successfully.")
                else:
                    print("Database saved successfully.")

        if verbose:
            print(f"[SyncCanvas] Sync completed: {added_count} students added, {updated_count} students updated.")
        else:
            print(f"Sync completed: {added_count} students added, {updated_count} students updated.")

        return added_count, updated_count

    except Exception as e:
        if verbose:
            print(f"[SyncCanvas] Error syncing with Canvas: {e}")
            traceback.print_exc()
        else:
            print(f"Error syncing with Canvas: {e}")
        return 0, 0

def interactive_modify_database(students, db_path=None, verbose=False):
    """
    Interactively modify student records in the database.
    Allows searching, editing, adding, and deleting students.
    When editing a field, pre-fill the old value for easy modification.
    After editing a field, ask if the user wants to continue editing other fields of the same student,
    edit another student, or quit. Always allow quitting at any step.
    If no response after 60 seconds from user then quit.
    If verbose is True, print more details; otherwise, print only important notice.
    """
    try:
        if db_path:
            students = load_database(db_path, verbose=verbose)
        if not students:
            if verbose:
                print("[ModifyDB] No students in the database.")
            else:
                print("No students in the database.")
            return

        def list_students():
            if verbose:
                print("[ModifyDB] List of students:")
            else:
                print("List of students:")
            for idx, s in enumerate(students, 1):
                name = getattr(s, "Name", "")
                sid = getattr(s, "Student ID", "")
                print(f"{idx}. {name} ({sid})")

        while True:
            if verbose:
                print("\n[ModifyDB] Modify Menu:")
            else:
                print("\nModify Menu:")
            print("1. List students")
            print("2. Edit a student")
            print("3. Add a new student")
            print("4. Delete a student")
            print("0. Exit modify menu")
            
            try:
                choice = get_input_with_timeout("Choose an option (or 'q' to quit): ").strip()
            except TimeoutError:
                return
            except KeyboardInterrupt:
                return
            
            if choice in ("0", "q", "Q"):
                break
            elif choice == "1":
                list_students()
            elif choice == "2":
                while True:
                    list_students()
                    try:
                        idx = get_input_with_timeout("Enter the number of the student to edit (or 'q' to quit): ").strip()
                    except TimeoutError:
                        return
                    except KeyboardInterrupt:
                        return
                    
                    if idx.lower() in ("q", "quit"):
                        break
                    if not idx.isdigit() or int(idx) < 1 or int(idx) > len(students):
                        if verbose:
                            print("[ModifyDB] Invalid index.")
                        else:
                            print("Invalid index.")
                        continue
                    s = students[int(idx) - 1]
                    while True:
                        if verbose:
                            print("[ModifyDB] Current fields:")
                        else:
                            print("Current fields:")
                        for k, v in s.__dict__.items():
                            print(f"{k}: {v}")
                        
                        try:
                            field = get_input_with_timeout("Enter field to edit (or leave blank to cancel, or 'q' to quit): ").strip()
                        except TimeoutError:
                            return
                        except KeyboardInterrupt:
                            return
                        
                        if not field or field.lower() == "q":
                            break
                        old_value = getattr(s, field, "")
                        # Pre-fill old value for editing
                        try:
                            value = prefill_input_with_timeout(f"Enter new value for '{field}' [{old_value}]: ", old_value)
                        except TimeoutError:
                            return
                        except KeyboardInterrupt:
                            return
                        except Exception:
                            try:
                                value = get_input_with_timeout(f"Enter new value for '{field}' (old: {old_value}): ").strip()
                                if not value:
                                    value = old_value
                            except TimeoutError:
                                return
                            except KeyboardInterrupt:
                                return
                        
                        if value.lower() == "q":
                            break
                        setattr(s, field, value)
                        if verbose:
                            print(f"[ModifyDB] Student updated: {field} = {value}")
                        else:
                            print("Student updated.")
                        if db_path:
                            save_database(students, db_path, verbose=verbose)
                        # Ask if continue editing this student, edit another student, or quit
                        try:
                            next_action = get_input_with_timeout("Continue editing other fields of this student (c), edit another student (a), or quit (q)? [c/a/q]: ").strip().lower()
                        except TimeoutError:
                            return
                        except KeyboardInterrupt:
                            return
                        
                        if next_action == "q":
                            return
                        elif next_action == "a":
                            break
                        # else continue editing this student
            elif choice == "3":
                fields = {}
                if verbose:
                    print("[ModifyDB] Enter new student information (leave blank to skip a field, or 'q' to quit):")
                else:
                    print("Enter new student information (leave blank to skip a field, or 'q' to quit):")
                for field in ["Name", "Student ID", "Email", "Class"]:
                    try:
                        value = get_input_with_timeout(f"{field}: ").strip()
                    except TimeoutError:
                        return
                    except KeyboardInterrupt:
                        return
                    
                    if value.lower() == "q":
                        if verbose:
                            print("[ModifyDB] Cancelled adding new student.")
                        else:
                            print("Cancelled adding new student.")
                        break
                    if value:
                        fields[field] = value
                else:
                    students.append(Student(**fields))
                    if verbose:
                        print(f"[ModifyDB] Student added: {fields}")
                    else:
                        print("Student added.")
                    if db_path:
                        save_database(students, db_path, verbose=verbose)
            elif choice == "4":
                while True:
                    list_students()
                    try:
                        idx = get_input_with_timeout("Enter the number of the student to delete (or 'q' to quit): ").strip()
                    except TimeoutError:
                        return
                    except KeyboardInterrupt:
                        return
                    
                    if idx.lower() in ("q", "quit"):
                        break
                    if not idx.isdigit() or int(idx) < 1 or int(idx) > len(students):
                        if verbose:
                            print("[ModifyDB] Invalid index.")
                        else:
                            print("Invalid index.")
                        continue
                    del students[int(idx) - 1]
                    if verbose:
                        print("[ModifyDB] Student deleted.")
                    else:
                        print("Student deleted.")
                    if db_path:
                        save_database(students, db_path, verbose=verbose)
                    break
            else:
                if verbose:
                    print("[ModifyDB] Invalid option.")
                else:
                    print("Invalid option.")
    
    except TimeoutError:
        if verbose:
            print("\n[ModifyDB] Timeout occurred. Exiting modify menu.")
        else:
            print("\nTimeout occurred. Exiting modify menu.")
        return
    except KeyboardInterrupt:
        if verbose:
            print("\n[ModifyDB] Operation cancelled by user. Exiting modify menu.")
        else:
            print("\nOperation cancelled by user. Exiting modify menu.")
        return

def list_canvas_assignments(
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    category=None,
    verbose=False
):
    """
    List all assignments in a Canvas course, grouped by assignment category (assignment group), using the canvasapi library.
    If category is specified (case-insensitive), only list assignments in that category.
    Assignments are sorted by due date (earliest first; undated last).

    Args:
        api_url (str): The base URL for the Canvas instance (e.g., "https://canvas.instructure.com")
        api_key (str): Your Canvas API access token
        course_id (int or str): The Canvas course ID
        category (str, optional): Assignment group name to filter (case-insensitive)
        verbose (bool): If True, print more details; otherwise, print only important notice.

    Returns:
        Dict mapping assignment group names to lists of assignment dicts with id, name, due_at, and points_possible
    """
    try:
        canvas = Canvas(api_url, api_key)
        course = canvas.get_course(course_id)
        if verbose:
            print(f"[CanvasAssignments] Listing assignments for course: \"{course.name} (ID: {course.id})\"")
        else:
            print(f"Listing assignments for course: \"{course.name} (ID: {course.id})\"")
        # Get all assignment groups with assignments included
        assignment_groups = course.get_assignment_groups(include=['assignments'])
        group_assignments = {}
        for group in assignment_groups:
            group_name = group.name
            assignments = []
            for assignment in group.assignments:
                assignments.append({
                    "id": assignment['id'],
                    "name": assignment['name'],
                    "due_at": format_time(assignment.get('due_at')),
                    "due_at_raw": assignment.get('due_at'),
                    "points_possible": assignment.get('points_possible')
                })
            # Sort assignments by due date (None last)
            def due_sort_key(a):
                raw = a.get("due_at_raw")
                if raw:
                    try:
                        return datetime.strptime(raw, "%Y-%m-%dT%H:%M:%SZ")
                    except Exception:
                        return datetime.max
                return datetime.max
            assignments.sort(key=due_sort_key)
            # Remove 'due_at_raw' from output
            for a in assignments:
                a.pop("due_at_raw", None)
            group_assignments[group_name] = assignments
        # If category is specified, filter to only that group (case-insensitive)
        if category:
            matched = None
            for group_name in group_assignments:
                if group_name.lower() == category.lower():
                    matched = group_name
                    break
            if matched:
                group_assignments = {matched: group_assignments[matched]}
            else:
                if verbose:
                    print(f"[CanvasAssignments] No assignment group found matching category '{category}'.")
                else:
                    print(f"No assignment group found matching category '{category}'.")
                return {}
        # Print assignments grouped by category
        for group_name, assignments in group_assignments.items():
            if verbose:
                print(f"[CanvasAssignments] Category: {group_name} ({len(assignments)} assignments)")
            else:
                print(f"\nCategory: {group_name} ({len(assignments)} assignments)")
            for a in assignments:
                if verbose:
                    print(f"[CanvasAssignments]   ID: {a['id']}, Name: {a['name']}, Due: {a['due_at']}, Points: {a['points_possible']}")
                else:
                    print(f"  ID: {a['id']}, Name: {a['name']}, Due: {a['due_at']}, Points: {a['points_possible']}")
        return group_assignments
    except ImportError:
        if verbose:
            print("[CanvasAssignments] canvasapi library is not installed. Please install it with 'pip install canvasapi'.")
        else:
            print("canvasapi library is not installed. Please install it with 'pip install canvasapi'.")
        return {}
    except Exception as e:
        if verbose:
            print(f"[CanvasAssignments] Error listing assignments: {e}")
        else:
            print(f"Error listing assignments: {e}")
        return {}

def list_canvas_people(api_url=CANVAS_LMS_API_URL, api_key=CANVAS_LMS_API_KEY, course_id=CANVAS_LMS_COURSE_ID, verbose=False):
    """
    List all teachers, TAs, and students in a Canvas course using the canvasapi library.
    For students, divides into active students and invited students.

    Args:
        api_url (str): The base URL for the Canvas instance.
        api_key (str): Your Canvas API access token.
        course_id (int or str): The Canvas course ID.
        verbose (bool): If True, print more details; otherwise, print only important notice.

    Returns:
        Dict with keys 'teachers', 'tas', 'active_students', 'invited_students', each a list of dicts with id, name, email, and canvas_id.
    """
    try:
        canvas = Canvas(api_url, api_key)
        course = canvas.get_course(course_id)
        enrollments = course.get_enrollments(type=['TeacherEnrollment', 'TaEnrollment', 'StudentEnrollment'])
        people = {'teachers': [], 'tas': [], 'active_students': [], 'invited_students': []}
        for enrollment in enrollments:
            user = enrollment.user
            role = enrollment.type.lower()
            person = {
                "id": user['id'],
                "canvas_id": user['id'],
                "name": user.get('name', ''),
                "email": user.get('login_id', '')
            }
            if role == 'teacherenrollment':
                people['teachers'].append(person)
            elif role == 'taenrollment':
                people['tas'].append(person)
            elif role == 'studentenrollment':
                state = enrollment.enrollment_state.lower()
                if state == 'active':
                    people['active_students'].append(person)
                elif state == 'invited':
                    people['invited_students'].append(person)
        return people
    except ImportError:
        return {'teachers': [], 'tas': [], 'active_students': [], 'invited_students': []}
    except Exception as e:
        return {'teachers': [], 'tas': [], 'active_students': [], 'invited_students': []}

def print_canvas_people(people, verbose=False):
    """
    Print the Canvas people dictionary returned by list_canvas_people.
    """
    if people.get('teachers'):
        if verbose:
            print(f"[CanvasPeople] Teachers ({len(people['teachers'])}):")
        else:
            print(f"\nTeachers ({len(people['teachers'])}):")
        for t in people['teachers']:
            if verbose:
                print(f"[CanvasPeople]   {t['name']} ({t['email']}), Canvas ID: {t['canvas_id']}")
            else:
                print(f"  {t['name']} ({t['email']}), Canvas ID: {t['canvas_id']}")
    if people.get('tas'):
        if verbose:
            print(f"[CanvasPeople] TAs ({len(people['tas'])}):")
        else:
            print(f"\nTAs ({len(people['tas'])}):")
        for ta in people['tas']:
            if verbose:
                print(f"[CanvasPeople]   {ta['name']} ({ta['email']}), Canvas ID: {ta['canvas_id']}")
            else:
                print(f"  {ta['name']} ({ta['email']}), Canvas ID: {ta['canvas_id']}")
    if people.get('active_students'):
        if verbose:
            print(f"[CanvasPeople] Active Students ({len(people['active_students'])}):")
        else:
            print(f"\nActive Students ({len(people['active_students'])}):")
        for s in people['active_students']:
            if verbose:
                print(f"[CanvasPeople]   {s['name']} ({s['email']}), Canvas ID: {s['canvas_id']}")
            else:
                print(f"  {s['name']} ({s['email']}), Canvas ID: {s['canvas_id']}")
    if people.get('invited_students'):
        if verbose:
            print(f"[CanvasPeople] Invited Students ({len(people['invited_students'])}):")
        else:
            print(f"\nInvited Students ({len(people['invited_students'])}):")
        for s in people['invited_students']:
            if verbose:
                print(f"[CanvasPeople]   {s['name']} ({s['email']}), Canvas ID: {s['canvas_id']}")
            else:
                print(f"  {s['name']} ({s['email']}), Canvas ID: {s['canvas_id']}")

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

    def get_input_with_timeout(prompt, timeout=60, default=None):
        # Use signal.SIGALRM only if available (not on Windows)
        if hasattr(signal, "SIGALRM"):
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(timeout)
            try:
                result = input(prompt)
                signal.alarm(0)
                if not result and default is not None:
                    if verbose:
                        print(f"[DownloadCanvas] Using default: {default}")
                    else:
                        print(f"Using default: {default}")
                    return default
                return result
            except TimeoutError:
                signal.alarm(0)
                if default is not None:
                    if verbose:
                        print(f"[DownloadCanvas] Using default: {default}")
                    else:
                        print(f"Using default: {default}")
                    return default
                raise
            except KeyboardInterrupt:
                signal.alarm(0)
                if verbose:
                    print("\n[DownloadCanvas] Operation cancelled by user.")
                else:
                    print("\nOperation cancelled by user.")
                raise
        else:
            # Fallback for platforms without SIGALRM (e.g., Windows)
            try:
                result = input(prompt)
                if not result and default is not None:
                    if verbose:
                        print(f"[DownloadCanvas] Using default: {default}")
                    else:
                        print(f"Using default: {default}")
                    return default
                return result
            except KeyboardInterrupt:
                if verbose:
                    print("\n[DownloadCanvas] Operation cancelled by user.")
                else:
                    print("\nOperation cancelled by user.")
                raise

    try:
        canvas = Canvas(api_url, api_key)
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
                sel = get_input_with_timeout(
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
                download_all_choice = get_input_with_timeout(
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
                    choice = get_input_with_timeout(
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
    canvas = Canvas(api_url, api_key)
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

def compare_texts_from_pdfs_in_folder(
    folder_path,
    ocr_service=DEFAULT_OCR_METHOD,
    lang="auto",
    simple_text=False,
    refine=DEFAULT_AI_METHOD,
    similarity_threshold=0.85,
    db_path=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    auto_send=False,
    verbose=False
):
    """
    Extract texts from all PDFs in a folder, compare the extracted texts,
    and output the names of the corresponding PDFs which have high similarity in contents.
    Also saves the comparison results to a TXT file in the same folder.
    If two or more PDFs are highly similar, send a message to all corresponding students
    asking them to resubmit, indicating that this is cheating and not allowed.

    Additionally, save the status of message sent for each pair to the TXT file.
    On later runs, do not send message for the same pair of PDFs again.

    Args:
        folder_path (str): Path to the folder containing PDF files.
        ocr_service (str): OCR service to use ("ocrspace", "tesseract", "paddleocr").
        lang (str): OCR language.
        simple_text (bool): If True, extract simple text.
        refine (str): AI refinement ("gemini", "huggingface", or None).
        similarity_threshold (float): Threshold for considering two PDFs as similar (0-1).
        api_url (str): Canvas API base URL.
        api_key (str): Canvas API key.
        course_id (str): Canvas course ID.
        verbose (bool): If True, print more details; otherwise, print only important notice.

    Returns:
        List of tuples: [(pdf1, pdf2, similarity), ...] for pairs above threshold.
    """

    assignment_name_guess = os.path.basename(folder_path).replace("_", " ").strip()
    image_phash_threshold = 0.95
    image_ssim_threshold = 0.9
    layout_threshold = 0.9
    shingle_threshold = 0.6
    embedding_threshold = 0.8

    def _extract_metadata_from_filename(filename):
        # Expected pattern: <name>_<canvas_id>_<assignment_id>_<time>_<status>.pdf
        base = os.path.basename(filename)
        meta = {
            "file": base,
            "student_name": None,
            "canvas_id": None,
            "assignment_id": None,
            "submitted_at": None,
            "status": None,
        }
        match = re.match(r"^(?P<name>.+)_(?P<canvas_id>\\d+)_(?P<assignment_id>\\d+)_(?P<submitted>[^_]+)_(?P<status>[^_]+)\\.pdf$", base)
        if match:
            meta["student_name"] = match.group("name").replace("_", " ").strip()
            meta["canvas_id"] = match.group("canvas_id")
            meta["assignment_id"] = match.group("assignment_id")
            meta["submitted_at"] = match.group("submitted")
            meta["status"] = match.group("status")
            return meta
        # Fallback: try to infer a canvas id from numeric tokens
        parts = base.replace(".pdf", "").split("_")
        numeric = [p for p in parts if p.isdigit()]
        if numeric:
            meta["canvas_id"] = numeric[0]
        meta["student_name"] = parts[0].replace("_", " ").strip() if parts else None
        return meta

    def _file_md5(path, block_size=1 << 20):
        digest = hashlib.md5()
        try:
            with open(path, "rb") as f:
                while True:
                    data = f.read(block_size)
                    if not data:
                        break
                    digest.update(data)
            return digest.hexdigest()
        except OSError:
            return None

    def _phash(gray_image):
        # Perceptual hash via DCT on a 32x32 grayscale image.
        resized = cv2.resize(gray_image, (32, 32))
        dct = cv2.dct(resized.astype(np.float32))
        dct_low = dct[:8, :8].flatten()
        if len(dct_low) <= 1:
            return None
        median = np.median(dct_low[1:])
        bits = dct_low > median
        return "".join("1" if b else "0" for b in bits)

    def _phash_similarity(h1, h2):
        if not h1 or not h2 or len(h1) != len(h2):
            return None
        dist = sum(c1 != c2 for c1, c2 in zip(h1, h2))
        return 1.0 - (dist / float(len(h1)))

    def _ssim(img1, img2):
        # Basic SSIM on full-image statistics (fast, no sliding window).
        img1 = img1.astype(np.float64)
        img2 = img2.astype(np.float64)
        c1 = (0.01 * 255) ** 2
        c2 = (0.03 * 255) ** 2
        mu1 = img1.mean()
        mu2 = img2.mean()
        sigma1 = img1.var()
        sigma2 = img2.var()
        sigma12 = ((img1 - mu1) * (img2 - mu2)).mean()
        numerator = (2 * mu1 * mu2 + c1) * (2 * sigma12 + c2)
        denominator = (mu1 ** 2 + mu2 ** 2 + c1) * (sigma1 + sigma2 + c2)
        if denominator == 0:
            return 0.0
        return float(numerator / denominator)

    def _psnr(img1, img2):
        try:
            return float(cv2.PSNR(img1, img2))
        except Exception:
            mse = np.mean((img1.astype(np.float64) - img2.astype(np.float64)) ** 2)
            if mse == 0:
                return float("inf")
            return 20 * math.log10(255.0 / math.sqrt(mse))

    def _layout_signature(pil_image, grid_size=4):
        # Layout signature based on OCR bounding boxes bucketed into a grid.
        try:
            data = pytesseract.image_to_data(pil_image, output_type=pytesseract.Output.DICT)
        except Exception:
            return None
        w, h = pil_image.size
        if not w or not h:
            return None
        grid = np.zeros((grid_size, grid_size), dtype=np.float64)
        for left, top, width, height, text in zip(
            data.get("left", []),
            data.get("top", []),
            data.get("width", []),
            data.get("height", []),
            data.get("text", []),
        ):
            if not text or not str(text).strip():
                continue
            cx = left + width / 2.0
            cy = top + height / 2.0
            gx = min(grid_size - 1, max(0, int((cx / w) * grid_size)))
            gy = min(grid_size - 1, max(0, int((cy / h) * grid_size)))
            grid[gy, gx] += 1.0
        vec = grid.flatten()
        norm = np.linalg.norm(vec)
        if norm == 0:
            return None
        return vec / norm

    def _cosine_similarity_vector(vec1, vec2):
        if vec1 is None or vec2 is None:
            return None
        denom = (np.linalg.norm(vec1) * np.linalg.norm(vec2))
        if denom == 0:
            return None
        return float(np.dot(vec1, vec2) / denom)

    def _make_shingles(text, size=5):
        tokens = [t for t in text.split() if t]
        if len(tokens) < size:
            return set()
        return {" ".join(tokens[i:i + size]) for i in range(len(tokens) - size + 1)}

    pdf_files = sorted([
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.lower().endswith(".pdf")
    ])
    if not pdf_files:
        if verbose:
            print("[PDFSimilarity] No PDF files found in the folder.")
        else:
            print("No PDF files found in the folder.")
        return []

    # Collect file metadata and extract texts for all PDFs
    file_metadata = {}
    extracted_texts = {}
    for pdf_path in tqdm(pdf_files, desc="Extracting texts from PDFs"):
        meta = _extract_metadata_from_filename(pdf_path)
        meta["size_bytes"] = os.path.getsize(pdf_path) if os.path.exists(pdf_path) else None
        meta["md5"] = _file_md5(pdf_path)
        try:
            reader = PyPDF2.PdfReader(pdf_path)
            meta["page_count"] = len(reader.pages)
            info = getattr(reader, "metadata", None) or {}
            meta["producer"] = info.get("/Producer") or info.get("Producer")
            meta["creator"] = info.get("/Creator") or info.get("Creator")
        except Exception:
            meta.setdefault("page_count", None)
        file_metadata[pdf_path] = meta
        base = os.path.splitext(pdf_path)[0]
        txt_path = base + f"_text_{ocr_service}.txt"
        if not os.path.exists(txt_path):
            txt_path = extract_text_from_scanned_pdf(
                pdf_path,
                txt_output_path=txt_path,
                service=ocr_service,
                lang=lang,
                simple_text=simple_text,
                refine=refine,
                verbose=verbose
            )
        if txt_path and os.path.exists(txt_path):
            with open(txt_path, "r", encoding="utf-8") as f:
                text = f.read()
            # Normalize text: remove whitespace, lowercase
            norm_text = re.sub(r"\s+", " ", text).strip().lower()
            extracted_texts[pdf_path] = norm_text
        else:
            extracted_texts[pdf_path] = ""

    # Prepare image-based features and layout signatures (first page only)
    image_features = {}
    for pdf_path in tqdm(pdf_files, desc="Extracting image features"):
        features = {}
        try:
            images = convert_from_path(pdf_path, first_page=1, last_page=1, dpi=100)
            if images:
                pil_image = images[0]
                gray = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2GRAY)
                resized = cv2.resize(gray, (256, 256))
                features["phash"] = _phash(gray)
                features["layout"] = _layout_signature(pil_image)
                features["image"] = resized
        except Exception:
            features = {}
        image_features[pdf_path] = features

    # Build shingle sets for text similarity
    shingle_sets = {pdf_path: _make_shingles(text) for pdf_path, text in extracted_texts.items()}

    # Compare all pairs using multiple similarity metrics for better accuracy

    similar_pairs = []
    all_pairs = []
    pdf_list = list(extracted_texts.keys())
    similarity_matrix = {}  # (pdf1, pdf2) -> ratio
    all_pairs_detail = []

    # Prepare TF-IDF vectors for all texts
    texts = [extracted_texts[p] for p in pdf_list]
    tfidf_matrix = None
    if any(texts):
        tfidf_vectorizer = TfidfVectorizer().fit(texts)
        tfidf_matrix = tfidf_vectorizer.transform(texts)

    # Optional sentence-embedding similarity (if sentence-transformers is installed)
    embedding_vectors = None
    embedding_method = None
    try:
        from sentence_transformers import SentenceTransformer

        model = SentenceTransformer("all-MiniLM-L6-v2")
        embedding_vectors = model.encode(texts, convert_to_numpy=True, normalize_embeddings=True)
        embedding_method = "sentence_transformers/all-MiniLM-L6-v2"
    except Exception:
        embedding_vectors = None

    # Helper to compute Jaccard similarity
    def jaccard_similarity(a, b):
        set_a = set(a.split())
        set_b = set(b.split())
        intersection = set_a & set_b
        union = set_a | set_b
        return len(intersection) / len(union) if union else 0.0

    # Helper to compute Euclidean distance similarity (1 / (1 + distance))
    def euclidean_similarity(vec1, vec2):
        dist = np.linalg.norm(vec1 - vec2)
        return 1.0 / (1.0 + dist)

    for i, pdf1 in enumerate(pdf_list):
        for j in range(i + 1, len(pdf_list)):
            pdf2 = pdf_list[j]
            text1 = extracted_texts[pdf1]
            text2 = extracted_texts[pdf2]

            meta1 = file_metadata.get(pdf1, {})
            meta2 = file_metadata.get(pdf2, {})
            exact_hash = meta1.get("md5") and meta1.get("md5") == meta2.get("md5")

            # Cosine similarity (TF-IDF)
            if tfidf_matrix is not None:
                cos_sim = cosine_similarity(tfidf_matrix[i], tfidf_matrix[j])[0, 0]
                euc_sim = euclidean_similarity(tfidf_matrix[i].toarray(), tfidf_matrix[j].toarray())
            else:
                cos_sim = 0.0
                euc_sim = 0.0

            # Jaccard similarity
            jac_sim = jaccard_similarity(text1, text2) if text1 and text2 else 0.0

            # SequenceMatcher similarity
            seq_sim = difflib.SequenceMatcher(None, text1, text2).ratio() if text1 and text2 else 0.0

            # Weighted average (can adjust weights as needed)
            ratio = (0.4 * cos_sim) + (0.25 * jac_sim) + (0.2 * seq_sim) + (0.15 * euc_sim)
            if exact_hash:
                ratio = 1.0

            # Image-based similarity (first page)
            img1 = image_features.get(pdf1, {}).get("image")
            img2 = image_features.get(pdf2, {}).get("image")
            phash_sim = _phash_similarity(
                image_features.get(pdf1, {}).get("phash"),
                image_features.get(pdf2, {}).get("phash")
            )
            ssim_value = _ssim(img1, img2) if img1 is not None and img2 is not None else None
            psnr_value = _psnr(img1, img2) if img1 is not None and img2 is not None else None

            # Layout-aware similarity
            layout_sim = _cosine_similarity_vector(
                image_features.get(pdf1, {}).get("layout"),
                image_features.get(pdf2, {}).get("layout")
            )

            # N-gram shingle similarity
            shingle_sim = 0.0
            shingles1 = shingle_sets.get(pdf1, set())
            shingles2 = shingle_sets.get(pdf2, set())
            if shingles1 and shingles2:
                shingle_sim = len(shingles1 & shingles2) / float(len(shingles1 | shingles2))

            # Embedding similarity (optional)
            embed_sim = None
            if embedding_vectors is not None:
                embed_sim = float(np.dot(embedding_vectors[i], embedding_vectors[j]))

            # Metadata match (does not trigger by itself)
            meta_match = False
            if meta1.get("producer") and meta2.get("producer") and meta1.get("producer") == meta2.get("producer"):
                if meta1.get("creator") == meta2.get("creator") and meta1.get("page_count") == meta2.get("page_count"):
                    meta_match = True

            all_pairs.append((os.path.basename(pdf1), os.path.basename(pdf2), ratio))
            similarity_matrix[(pdf1, pdf2)] = ratio
            similarity_matrix[(pdf2, pdf1)] = ratio
            text_flag = ratio >= similarity_threshold
            image_phash_flag = phash_sim is not None and phash_sim >= image_phash_threshold
            image_ssim_flag = ssim_value is not None and ssim_value >= image_ssim_threshold
            layout_flag = layout_sim is not None and layout_sim >= layout_threshold
            shingle_flag = shingle_sim >= shingle_threshold
            embedding_flag = embed_sim is not None and embed_sim >= embedding_threshold

            # Balanced rule: exact hash OR (text similarity + at least one image/layout signal).
            flagged = exact_hash or (text_flag and (image_phash_flag or image_ssim_flag or layout_flag))

            reasons = []
            if flagged:
                if exact_hash:
                    reasons.append("exact_hash")
                if text_flag:
                    reasons.append("text_similarity")
                if image_phash_flag:
                    reasons.append("image_phash")
                if image_ssim_flag:
                    reasons.append("image_ssim")
                if layout_flag:
                    reasons.append("layout_similarity")
                if shingle_flag:
                    reasons.append("ngram_shingles")
                if embedding_flag:
                    reasons.append("embedding_similarity")
                if meta_match:
                    reasons.append("metadata_match")
                similar_pairs.append((os.path.basename(pdf1), os.path.basename(pdf2), ratio))
            all_pairs_detail.append({
                "pdf1": os.path.basename(pdf1),
                "pdf2": os.path.basename(pdf2),
                "ratio": ratio,
                "metrics": {
                    "cosine": cos_sim,
                    "jaccard": jac_sim,
                    "sequence": seq_sim,
                    "euclidean": euc_sim,
                    "phash_similarity": phash_sim,
                    "ssim": ssim_value,
                    "psnr": psnr_value,
                    "layout_similarity": layout_sim,
                    "shingle_jaccard": shingle_sim,
                    "embedding_cosine": embed_sim,
                },
                "exact_hash": bool(exact_hash),
                "metadata_match": meta_match,
                "reasons": reasons,
            })

    # Save results to file in the same folder
    result_path = os.path.join(folder_path, "pdf_similarity_results.txt")
    status_path = os.path.join(folder_path, "pdf_similarity_status.json")
    report_path = os.path.join(folder_path, "pdf_similarity_report.json")

    # Load previous status if exists
    sent_status = {}
    if os.path.exists(status_path):
        try:
            with open(status_path, "r", encoding="utf-8") as f:
                sent_status = json.load(f)
        except Exception:
            sent_status = {}

    # Helper to create a unique key for a pair (order-independent)
    def pair_key(pdf1, pdf2):
        return "||".join(sorted([pdf1.lower(), pdf2.lower()]))

    # Save results to txt
    with open(result_path, "w", encoding="utf-8") as f:
        f.write("PDF similarity comparison results:\n")
        if all_pairs:
            for pdf1, pdf2, ratio in sorted(all_pairs, key=lambda x: -x[2]):
                mark = " <== HIGH SIMILARITY" if ratio >= similarity_threshold else ""
                key = pair_key(pdf1, pdf2)
                msg_status = sent_status.get(key, "NOT_SENT")
                f.write(f"{pdf1} <-> {pdf2}: similarity = {ratio:.2f}{mark} [Message: {msg_status}]\n")
        else:
            f.write("No PDF pairs to compare.\n")
        if similar_pairs:
            f.write("\nPDF pairs with high similarity (>= {:.2f}):\n".format(similarity_threshold))
            for pdf1, pdf2, ratio in sorted(similar_pairs, key=lambda x: -x[2]):
                key = pair_key(pdf1, pdf2)
                msg_status = sent_status.get(key, "NOT_SENT")
                f.write(f"{pdf1} <-> {pdf2}: similarity = {ratio:.2f} [Message: {msg_status}]\n")
        else:
            f.write("\nNo highly similar PDF pairs found.\n")
    if verbose:
        print(f"[PDFSimilarity] Comparison results saved to {result_path}")
    else:
        print(f"Comparison results saved to {result_path}")

    pair_details_by_key = {
        pair_key(p["pdf1"], p["pdf2"]): p for p in all_pairs_detail
    }

    # Cluster detection for groups of similar submissions
    def _build_clusters(details):
        parent = {}

        def find(x):
            parent.setdefault(x, x)
            if parent[x] != x:
                parent[x] = find(parent[x])
            return parent[x]

        def union(a, b):
            ra, rb = find(a), find(b)
            if ra != rb:
                parent[rb] = ra

        for detail in details:
            if detail.get("reasons"):
                union(detail["pdf1"], detail["pdf2"])

        clusters = {}
        for detail in details:
            for f in (detail["pdf1"], detail["pdf2"]):
                root = find(f)
                clusters.setdefault(root, set()).add(f)
        return [sorted(list(members)) for members in clusters.values() if len(members) > 1]

    clusters = _build_clusters(all_pairs_detail)

    # Save a structured report for downstream processing
    try:
        report_payload = {
            "generated_at": datetime.now().isoformat(),
            "threshold": similarity_threshold,
            "assignment_name": assignment_name_guess,
            "methods": {
                "text_similarity_threshold": similarity_threshold,
                "image_phash_threshold": image_phash_threshold,
                "image_ssim_threshold": image_ssim_threshold,
                "layout_threshold": layout_threshold,
                "shingle_threshold": shingle_threshold,
                "embedding_threshold": embedding_threshold,
                "embedding_method": embedding_method,
                "flag_rule": "exact_hash OR (text_similarity AND (image_phash OR image_ssim OR layout_similarity))",
            },
            "files": {os.path.basename(k): v for k, v in file_metadata.items()},
            "pairs": all_pairs_detail,
            "high_similarity": [
                p for p in all_pairs_detail
                if p["reasons"]
            ],
            "clusters": clusters,
        }
        with open(report_path, "w", encoding="utf-8") as f:
            json.dump(report_payload, f, ensure_ascii=False, indent=2)
        if verbose:
            print(f"[PDFSimilarity] Report saved to {report_path}")
    except Exception as e:
        if verbose:
            print(f"[PDFSimilarity] Failed to save report JSON: {e}")

    if similar_pairs:
        if verbose:
            print("[PDFSimilarity] PDF pairs with high similarity:")
            for pdf1, pdf2, ratio in sorted(similar_pairs, key=lambda x: -x[2]):
                key = pair_key(pdf1, pdf2)
                msg_status = sent_status.get(key, "NOT_SENT")
                print(f"  {pdf1} <-> {pdf2}: similarity = {ratio:.2f} [Message: {msg_status}]")
        else:
            print("PDF pairs with high similarity found.")
    else:
        if verbose:
            print("[PDFSimilarity] No highly similar PDF pairs found.")
        else:
            print("No highly similar PDF pairs found.")

    # Update local database with similarity flags if possible
    if db_path is None:
        db_path = get_default_db_path()
    if db_path and os.path.exists(db_path) and similar_pairs:
        try:
            students = load_database(db_path, verbose=verbose)
            sid_map = {}
            name_map = {}

            def _norm_name(value):
                return re.sub(r"\\s+", " ", str(value or "")).strip().lower()

            for s in students:
                canvas_id = getattr(s, "Canvas ID", None)
                if canvas_id is not None:
                    sid_map[str(canvas_id)] = s
                name = getattr(s, "Name", None)
                if name:
                    name_map[_norm_name(name)] = s

            def _resolve_student(meta):
                sid = str(meta.get("canvas_id") or "")
                if sid and sid in sid_map:
                    return sid_map[sid]
                name_key = _norm_name(meta.get("student_name"))
                return name_map.get(name_key)

            def _pair_key(a, b):
                return "||".join(sorted([a, b]))

            for pdf1, pdf2, ratio in similar_pairs:
                meta1 = file_metadata.get(os.path.join(folder_path, pdf1), file_metadata.get(pdf1, {}))
                meta2 = file_metadata.get(os.path.join(folder_path, pdf2), file_metadata.get(pdf2, {}))
                student1 = _resolve_student(meta1)
                student2 = _resolve_student(meta2)
                if not student1 and not student2:
                    continue
                pair_key = _pair_key(pdf1, pdf2)
                detail = pair_details_by_key.get(pair_key, {})
                reasons = detail.get("reasons", [])

                entry = {
                    "pair_key": pair_key,
                    "other_file": pdf2,
                    "similarity": round(ratio, 4),
                    "reasons": reasons,
                    "report_path": report_path,
                    "assignment_id": meta1.get("assignment_id") or meta2.get("assignment_id"),
                    "assignment_name": assignment_name_guess,
                }
                entry_other = {
                    "pair_key": pair_key,
                    "other_file": pdf1,
                    "similarity": round(ratio, 4),
                    "reasons": reasons,
                    "report_path": report_path,
                    "assignment_id": meta1.get("assignment_id") or meta2.get("assignment_id"),
                    "assignment_name": assignment_name_guess,
                }

                for student, payload in ((student1, entry), (student2, entry_other)):
                    if not student:
                        continue
                    existing = getattr(student, "Plagiarism Matches", [])
                    if not isinstance(existing, list):
                        existing = []
                    if not any(item.get("pair_key") == pair_key for item in existing if isinstance(item, dict)):
                        existing.append(payload)
                        setattr(student, "Plagiarism Matches", existing)

            save_database(students, db_path, verbose=verbose)
            if verbose:
                print(f"[PDFSimilarity] Saved plagiarism flags to database: {db_path}")
        except Exception as e:
            if verbose:
                print(f"[PDFSimilarity] Failed to update database: {e}")

    # Only send messages for pairs that have not been sent before
    pairs_to_notify = []
    for pdf1, pdf2, ratio in similar_pairs:
        key = pair_key(pdf1, pdf2)
        if sent_status.get(key, "NOT_SENT") != "SENT":
            pairs_to_notify.append((pdf1, pdf2, ratio))

    # If there are highly similar pairs to notify, send a separate message for each pair
    if pairs_to_notify:
        if verbose:
            print("[PDFSimilarity] Sending messages to students involved in newly detected highly similar submissions (one message per pair)...")
        else:
            print("Sending messages to students involved in highly similar submissions...")
        canvas = Canvas(api_url, api_key)
        course = canvas.get_course(course_id)

        # Helper to extract Canvas ID from filename
        def extract_canvas_id_from_filename(filename):
            parts = filename.split('_')
            for part in parts:
                if part.isdigit():
                    return int(part)
            return None

        # Ask user if they want to refine the message via AI if refine is None
        if refine is None and not auto_send:
            def get_input_with_timeout(prompt, timeout=60, default=None):
                # Use signal.SIGALRM only if available (not on Windows)
                if hasattr(signal, "SIGALRM"):
                    signal.signal(signal.SIGALRM, timeout_handler)
                    signal.alarm(timeout)
                    try:
                        result = input(prompt)
                        signal.alarm(0)
                        if not result and default is not None:
                            return default
                        return result
                    except TimeoutError:
                        signal.alarm(0)
                        if default is not None:
                            return default
                        raise
                    except KeyboardInterrupt:
                        signal.alarm(0)
                        raise
                else:
                    # Fallback for platforms without SIGALRM (e.g., Windows)
                    try:
                        result = input(prompt)
                        if not result and default is not None:
                            return default
                        return result
                    except KeyboardInterrupt:
                        raise

            try:
                refine_choice = get_input_with_timeout(
                    "Do you want to refine the message via AI? (none/gemini/huggingface/local) [none]: ",
                    timeout=60,
                    default="none"
                ).strip().lower()
                if refine_choice in ("none", "gemini", "huggingface"):
                    refine = refine_choice if refine_choice != "none" else None
                else:
                    if verbose:
                        print("[PDFSimilarity] Invalid choice. Using default 'none'.")
                    else:
                        print("Invalid choice. Using default 'none'.")
                    refine = None
            except TimeoutError:
                if verbose:
                    print("[PDFSimilarity] No response after 60 seconds. Using default 'none'.")
                else:
                    print("No response after 60 seconds. Using default 'none'.")
                refine = None
        if refine is None and auto_send:
            refine = None

        for pdf1, pdf2, ratio in pairs_to_notify:
            canvas_id1 = extract_canvas_id_from_filename(pdf1)
            canvas_id2 = extract_canvas_id_from_filename(pdf2)
            recipients = []
            if canvas_id1:
                recipients.append(str(canvas_id1))
            if canvas_id2 and canvas_id2 != canvas_id1:
                recipients.append(str(canvas_id2))
            if not recipients:
                if verbose:
                    print(f"[PDFSimilarity] Could not extract Canvas IDs from {pdf1} and {pdf2}. Skipping message.")
                else:
                    print(f"Could not extract Canvas IDs from {pdf1} and {pdf2}. Skipping message.")
                sent_status[pair_key(pdf1, pdf2)] = "FAILED"
                continue

            detail = pair_details_by_key.get(pair_key(pdf1, pdf2), {})
            reasons = detail.get("reasons", [])
            metrics = detail.get("metrics", {})

            method_descriptions = {
                "exact_hash": "Exact file hash match",
                "text_similarity": "Text similarity (TF-IDF/sequence)",
                "image_phash": "Perceptual image hash match",
                "image_ssim": "Image structural similarity",
                "layout_similarity": "Layout similarity from OCR bounding boxes",
                "ngram_shingles": "N-gram shingle overlap",
                "embedding_similarity": "Embedding similarity (if available)",
                "metadata_match": "PDF metadata match (producer/creator/pages)",
            }
            methods_used = [method_descriptions.get(r, r) for r in reasons]
            if not methods_used:
                methods_used = ["Text similarity (TF-IDF/sequence)"]
            method_block = "\n".join([f"- {m}" for m in methods_used])

            # Generate message for this pair
            similarity_results = f"{pdf1} <-> {pdf2}: similarity = {ratio:.2f}"
            metrics_summary = (
                f"TF-IDF cosine: {metrics.get('cosine')}, "
                f"Jaccard: {metrics.get('jaccard')}, "
                f"Sequence: {metrics.get('sequence')}, "
                f"Image pHash: {metrics.get('phash_similarity')}, "
                f"SSIM: {metrics.get('ssim')}, "
                f"Layout: {metrics.get('layout_similarity')}, "
                f"Shingles: {metrics.get('shingle_jaccard')}, "
                f"Embedding: {metrics.get('embedding_cosine')}"
            )
            if refine in ALL_AI_METHODS:
                if verbose:
                    print(f"[PDFSimilarity] Generating message using AI service: {refine} for pair {pdf1} <-> {pdf2} ...")
                else:
                    print(f"Generating message using AI service: {refine} for pair {pdf1} <-> {pdf2} ...")
                prompt = (
                    "You are an expert assistant. Compose a clear, formal, and professional message in Vietnamese to notify students "
                    "about potential similarity detected in their submissions by the system. The message should include the following points:\n\n"
                    "1. Provide the similarity results, listing the pair of submissions with their similarity score.\n"
                    "2. Explain the methods used (text similarity, image similarity, layout similarity, n-gram overlap, metadata match, optional embedding similarity).\n"
                    "3. Emphasize that automated detection can produce false positives and the case will be reviewed by lecturers and TAs before any final decision.\n"
                    "4. Ask students to wait for review or respond if they believe the detection is incorrect.\n\n"
                    "Ensure the message is complete, concise, and does not require any additional edits or replacements.\n\n"
                    "Similarity result:\n{text}\n"
                    "Methods:\n{methods}\n"
                    "Metrics:\n{metrics}"
                )
                message = refine_text_with_ai(
                    similarity_results,
                    method=refine,
                    user_prompt=prompt.format(text=similarity_results, methods=method_block, metrics=metrics_summary)
                )
            else:
                message = (
                    "Potential similarity detected by automated checks for the following submissions:\n"
                    + similarity_results +
                    "\n\nAssignment: "
                    + assignment_name_guess +
                    "\n\nMethods used:\n"
                    + method_block +
                    "\n\nMetrics:\n"
                    + metrics_summary +
                    "\n\nNote: Automated detection can produce false positives (OCR errors, formatting differences, or similar templates). "
                    "This case will be reviewed by the lecturers and TAs before any final decision. "
                    "If you believe this detection is incorrect, you may respond with clarification."
                )

            subject = "Notice: Potential similarity detected in submissions"

            if verbose:
                print(f"[PDFSimilarity] Subject:\n{subject}")
                print(f"[PDFSimilarity] Message for {pdf1} <-> {pdf2}:\n{message}")
            else:
                print(f"Prepared message for {pdf1} <-> {pdf2}.")

            if not auto_send:
                while True:
                    try:
                        # Only use SIGALRM if available (not on Windows)
                        if hasattr(signal, "SIGALRM"):
                            signal.signal(signal.SIGALRM, timeout_handler)
                            signal.alarm(60)  # 60 second timeout
                            confirm = input(f"\nDo you want to send this message for {pdf1} <-> {pdf2}? (y/n, or 'r' to regenerate, default 'y' in 60s): ").strip().lower()
                            signal.alarm(0)  # Cancel the alarm
                        else:
                            # Fallback for Windows (no timeout)
                            confirm = input(f"\nDo you want to send this message for {pdf1} <-> {pdf2}? (y/n, or 'r' to regenerate): ").strip().lower()
                            if not confirm:
                                confirm = "y"  # Default value

                        if confirm == "y" or confirm == "":
                            break
                        elif confirm == "n":
                            if verbose:
                                print(f"[PDFSimilarity] Message sending canceled for this pair.")
                            else:
                                print("Message sending canceled for this pair.")
                            sent_status[pair_key(pdf1, pdf2)] = "SKIPPED"
                            break
                        elif confirm == "r":
                            if verbose:
                                print(f"[PDFSimilarity] Regenerating message...")
                            else:
                                print("Regenerating message...")
                            if refine in ALL_AI_METHODS:
                                message = refine_text_with_ai(similarity_results, method=refine, user_prompt=prompt)
                            else:
                                message = (
                                    "Potential similarity detected by automated checks for the following submissions:\n"
                                    + similarity_results +
                                    "\n\nAssignment: "
                                    + assignment_name_guess +
                                    "\n\nMethods used:\n"
                                    + method_block +
                                    "\n\nMetrics:\n"
                                    + metrics_summary +
                                    "\n\nNote: Automated detection can produce false positives (OCR errors, formatting differences, or similar templates). "
                                    "This case will be reviewed by the lecturers and TAs before any final decision. "
                                    "If you believe this detection is incorrect, you may respond with clarification."
                                )
                            if verbose:
                                print(f"[PDFSimilarity] Regenerated Message:\n{message}")
                            else:
                                print("Regenerated message.")
                        else:
                            if verbose:
                                print("[PDFSimilarity] Invalid input. Please enter 'y', 'n', or 'r'.")
                            else:
                                print("Invalid input. Please enter 'y', 'n', or 'r'.")
                    except TimeoutError:
                        if verbose:
                            print("[PDFSimilarity] No response after 60 seconds, using default 'y'.")
                        else:
                            print("No response after 60 seconds, using default 'y'.")
                        break  # Use default 'y' option and break the loop

            if sent_status.get(pair_key(pdf1, pdf2)) == "SKIPPED":
                continue

            try:
                canvas.create_conversation(
                    recipients=recipients,
                    subject=subject,
                    body=message,
                    force_new=True
                )
                if verbose:
                    print(f"[PDFSimilarity] Message sent successfully to students for {pdf1} <-> {pdf2}.")
                else:
                    print(f"Message sent for {pdf1} <-> {pdf2}.")
                sent_status[pair_key(pdf1, pdf2)] = "SENT"
            except Exception as e:
                if verbose:
                    print(f"[PDFSimilarity] Failed to send message for {pdf1} <-> {pdf2}: {e}")
                else:
                    print(f"Failed to send message for {pdf1} <-> {pdf2}: {e}")
                sent_status[pair_key(pdf1, pdf2)] = "FAILED"

        # Save updated status to JSON
        try:
            with open(status_path, "w", encoding="utf-8") as f:
                json.dump(sent_status, f, ensure_ascii=False, indent=2)
            if verbose:
                print(f"[PDFSimilarity] Message status saved to {status_path}")
        except Exception as e:
            if verbose:
                print(f"[PDFSimilarity] Failed to save message status: {e}")
            else:
                print(f"Failed to save message status: {e}")

    else:
        # Save status file even if nothing to send, to keep track
        try:
            with open(status_path, "w", encoding="utf-8") as f:
                json.dump(sent_status, f, ensure_ascii=False, indent=2)
            if verbose:
                print(f"[PDFSimilarity] Message status saved to {status_path}")
        except Exception as e:
            if verbose:
                print(f"[PDFSimilarity] Failed to save message status: {e}")
            else:
                print(f"Failed to save message status: {e}")

    return similar_pairs

def detect_meaningful_level_and_notify_students(
    folder_path,
    assignment_id=None,
    ocr_service=DEFAULT_OCR_METHOD,
    lang="auto",
    simple_text=False,
    refine=DEFAULT_AI_METHOD,
    meaningfulness_threshold=0.4,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    auto_send=False,
    verbose=False
):
    """
    Detect the meaningful level of extracted texts from PDFs in a folder using AI agent.
    If the meaningful level is too low, generate a message via AI and send it to the student
    asking them to reformat and resubmit their submissions.

    Args:
        folder_path (str): Path to the folder containing PDF files.
        assignment_id (str): Canvas assignment ID for sending messages.
        ocr_service (str): OCR service to use ("ocrspace", "tesseract", "paddleocr").
        lang (str): OCR language.
        simple_text (bool): If True, extract simple text.
        refine (str): AI refinement ("gemini", "huggingface", or None).
        meaningfulness_threshold (float): Threshold for considering text as meaningful (0-1).
        api_url, api_key, course_id: Canvas API configuration.
        verbose (bool): If True, print more details; otherwise, print only important notice.

    Returns:
        Dict with results: {filename: {"meaningful_score": score, "message_sent": bool}, ...}
    """

    def get_input_with_timeout_default(prompt, timeout=60, default=None):
        # Use signal.SIGALRM only if available (not on Windows)
        if hasattr(signal, "SIGALRM"):
            try:
                signal.signal(signal.SIGALRM, timeout_handler)
                signal.alarm(timeout)
                result = input(prompt)
                signal.alarm(0)
                if not result and default is not None:
                    if verbose:
                        print(f"[Meaningfulness] Using default: {default}")
                    else:
                        print(f"Using default: {default}")
                    return default
                return result
            except TimeoutError:
                signal.alarm(0)
                if default is not None:
                    if verbose:
                        print(f"\n[Meaningfulness] No response after {timeout} seconds, using default '{default}'")
                    else:
                        print(f"\nNo response after {timeout} seconds, using default '{default}'")
                    return default
                raise
            except KeyboardInterrupt:
                signal.alarm(0)
                if verbose:
                    print("\n[Meaningfulness] Operation cancelled by user.")
                else:
                    print("\nOperation cancelled by user.")
                raise
        else:
            # Fallback for platforms without SIGALRM (e.g., Windows)
            try:
                result = input(prompt)
                if not result and default is not None:
                    if verbose:
                        print(f"[Meaningfulness] Using default: {default}")
                    else:
                        print(f"Using default: {default}")
                    return default
                return result
            except KeyboardInterrupt:
                if verbose:
                    print("\n[Meaningfulness] Operation cancelled by user.")
                else:
                    print("\nOperation cancelled by user.")
                raise

    pdf_files = sorted([
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.lower().endswith(".pdf")
    ])
    
    if not pdf_files:
        if verbose:
            print("[Meaningfulness] No PDF files found in the folder.")
        else:
            print("No PDF files found in the folder.")
        return {}

    status_path = os.path.join(folder_path, "meaningfulness_status.json")
    if os.path.exists(status_path):
        try:
            with open(status_path, "r", encoding="utf-8") as f:
                sent_status = json.load(f)
        except Exception:
            sent_status = {}
    else:
        sent_status = {}

    if refine is None and not auto_send:
        refine = get_input_with_timeout_default(
            "Which AI model do you want to use for meaningfulness analysis? (gemini/huggingface/local, default 'gemini' in 60s): ",
            timeout=60,
            default="gemini"
        ).strip().lower()
        if refine not in ALL_AI_METHODS:
            refine = "gemini"
    elif refine is None and auto_send:
        refine = DEFAULT_AI_METHOD or None

    results = {}
    low_quality_files = []
    text_lengths = []

    extracted_texts = {}
    for pdf_path in tqdm(pdf_files, desc="Extracting texts from PDFs"):
        filename = os.path.basename(pdf_path)
        base = os.path.splitext(pdf_path)[0]
        txt_path = base + f"_text_{ocr_service}.txt"

        if filename in sent_status and "meaningful_score" in sent_status[filename]:
            results[filename] = {
                "meaningful_score": sent_status[filename].get("meaningful_score", 0.0),
                "message_sent": sent_status[filename].get("message_sent", False),
                "error": sent_status[filename].get("error", "")
            }
            if "text_length" in sent_status[filename]:
                text_lengths.append(sent_status[filename]["text_length"])
            else:
                if os.path.exists(txt_path):
                    with open(txt_path, "r", encoding="utf-8") as f:
                        text = f.read()
                    text_lengths.append(len(text.strip()))
            continue

        if not os.path.exists(txt_path):
            txt_path = extract_text_from_scanned_pdf(
                pdf_path,
                txt_output_path=txt_path,
                service=ocr_service,
                lang=lang,
                simple_text=simple_text,
                refine=refine
            )
        
        if not txt_path or not os.path.exists(txt_path):
            results[filename] = {"meaningful_score": 0.0, "message_sent": False, "error": "Failed to extract text"}
            sent_status[filename] = results[filename]
            continue
        
        with open(txt_path, "r", encoding="utf-8") as f:
            text = f.read()
        
        extracted_texts[filename] = text
        text_lengths.append(len(text.strip()))

    average_length = sum(text_lengths) / len(text_lengths) if text_lengths else 0
    if verbose:
        print(f"[Meaningfulness] Average text length: {average_length:.0f} characters")
    else:
        print(f"Average text length: {average_length:.0f} characters")

    for filename, text in tqdm(extracted_texts.items(), desc="Analyzing PDF meaningfulness"):
        meaningful_score = analyze_text_meaningfulness(text, refine, average_length)
        # Store diagnostic metrics so the report can show why a file was flagged.
        metrics = _compute_text_quality_metrics(text)
        issues = _summarize_quality_issues(metrics, average_length=average_length)
        already_sent = sent_status.get(filename, {}).get("message_sent", False)
        results[filename] = {
            "meaningful_score": meaningful_score,
            "message_sent": already_sent,
            "text_length": len(text.strip()),
            "issues": issues,
            "metrics": {
                "vn_char_ratio": metrics["vn_char_ratio"],
                "alnum_ratio": metrics["alnum_ratio"],
                "symbol_ratio": metrics["symbol_ratio"],
                "unique_char_ratio": metrics["unique_char_ratio"],
                "repeat_char_ratio": metrics["repeat_char_ratio"],
                "line_empty_ratio": metrics["line_empty_ratio"],
                "likely_math": metrics["likely_math"],
            },
        }
        sent_status[filename] = results[filename]
        if meaningful_score < meaningfulness_threshold and not already_sent:
            low_quality_files.append((filename, meaningful_score, text))
    
    result_path = os.path.join(folder_path, "meaningfulness_analysis.txt")
    with open(result_path, "w", encoding="utf-8") as f:
        f.write("PDF meaningfulness analysis results:\n")
        f.write(f"Threshold: {meaningfulness_threshold}\n")
        f.write(f"Average text length: {average_length:.0f} characters\n")
        f.write(f"Analysis date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        total_files = len(results)
        low_quality_count = sum(1 for v in results.values() if v["meaningful_score"] < meaningfulness_threshold)
        acceptable_count = total_files - low_quality_count
        f.write(f"Summary:\n")
        f.write(f"Total files analyzed: {total_files}\n")
        f.write(f"Acceptable quality: {acceptable_count}\n")
        f.write(f"Low quality: {low_quality_count}\n\n")
        f.write("Detailed results:\n")
        f.write("-" * 80 + "\n")
        for filename, result in sorted(results.items()):
            score = result["meaningful_score"]
            status = "LOW QUALITY" if score < meaningfulness_threshold else "ACCEPTABLE"
            error = result.get("error", "")
            text_length = result.get("text_length", 0)
            length_ratio = text_length / average_length if average_length > 0 else 0
            msg_status = "SENT" if result.get("message_sent") else "NOT_SENT"
            issues = result.get("issues", [])
            metrics = result.get("metrics", {})
            issues_text = "; ".join(issues) if issues else "None"
            f.write(
                f"{filename}: score = {score:.2f} ({status}), length = {text_length} chars ({length_ratio:.2f}x avg), message: {msg_status}"
            )
            if error:
                f.write(f" - ERROR: {error}")
            f.write(f"\n  Issues: {issues_text}\n")
            if metrics:
                # Keep metrics on one line to simplify manual scanning.
                f.write("  Metrics: ")
                f.write(
                    f"vn_ratio={metrics.get('vn_char_ratio', 0):.2f}, alnum={metrics.get('alnum_ratio', 0):.2f}, "
                    f"symbol={metrics.get('symbol_ratio', 0):.2f}, unique={metrics.get('unique_char_ratio', 0):.2f}, "
                    f"repeat={metrics.get('repeat_char_ratio', 0):.2f}, empty_lines={metrics.get('line_empty_ratio', 0):.2f}, "
                    f"likely_math={metrics.get('likely_math', False)}\n"
                )
        if low_quality_files:
            f.write(f"\nLow quality files requiring attention (< {meaningfulness_threshold}):\n")
            f.write("-" * 80 + "\n")
            for filename, score, _ in low_quality_files:
                text_length = results.get(filename, {}).get("text_length", 0)
                length_ratio = text_length / average_length if average_length > 0 else 0
                msg_status = "SENT" if sent_status.get(filename, {}).get("message_sent") else "NOT_SENT"
                issues = results.get(filename, {}).get("issues", [])
                issues_text = "; ".join(issues) if issues else "None"
                f.write(
                    f"{filename}: score = {score:.2f}, length = {text_length} chars ({length_ratio:.2f}x avg), message: {msg_status}\n"
                )
                f.write(f"  Issues: {issues_text}\n")
    if verbose:
        print(f"[Meaningfulness] Analysis results saved to {result_path}")
    else:
        print(f"Analysis results saved to {result_path}")
    
    with open(status_path, "w", encoding="utf-8") as f:
        json.dump(sent_status, f, ensure_ascii=False, indent=2)

    if low_quality_files:
        if verbose:
            print(f"[Meaningfulness] Found {len(low_quality_files)} low quality submissions (not yet notified):")
            for filename, score, _ in low_quality_files:
                text_length = results.get(filename, {}).get("text_length", 0)
                length_ratio = text_length / average_length if average_length > 0 else 0
                print(f"  {filename}: score = {score:.2f}, length = {text_length} chars ({length_ratio:.2f}x avg)")
        else:
            print(f"Found {len(low_quality_files)} low quality submissions (not yet notified).")
        
        if not auto_send:
            send_messages = get_input_with_timeout_default(
                "\nDo you want to send messages to students with low quality submissions? (y/n, or 'q' to quit, default 'y' in 60s): ",
                timeout=60,
                default="y"
            ).strip().lower()
            if send_messages in ("q", "quit"):
                with open(status_path, "w", encoding="utf-8") as f:
                    json.dump(sent_status, f, ensure_ascii=False, indent=2)
                return results
            if send_messages not in ("y", "yes", ""):
                if verbose:
                    print("[Meaningfulness] Messages not sent.")
                else:
                    print("Messages not sent.")
                with open(status_path, "w", encoding="utf-8") as f:
                    json.dump(sent_status, f, ensure_ascii=False, indent=2)
                return results
        
        try:
            canvas = Canvas(api_url, api_key)
            course = canvas.get_course(course_id)
            for filename, score, text in low_quality_files:
                canvas_id = extract_canvas_id_from_filename(filename)
                if not canvas_id:
                    if verbose:
                        print(f"[Meaningfulness] Could not extract Canvas ID from {filename}")
                    else:
                        print(f"Could not extract Canvas ID from {filename}")
                    continue
                message = generate_low_quality_message(filename, score, text, refine)
                if verbose:
                    print(f"\n[Meaningfulness] Subject: {subject}")
                    print(f"[Meaningfulness] Message to {filename}:")
                    print("-" * 50)
                    print(message)
                    print("-" * 50)
                else:
                    print(f"\nPrepared message for {filename}.")
                if not auto_send:
                    while True:
                        action = get_input_with_timeout_default(
                            "\nWhat would you like to do? (s)end, (r)egenerate, or (q)uit [default: s in 60s]: ",
                            timeout=60,
                            default="s"
                        ).strip().lower()
                        if action in ('q', 'quit'):
                            if verbose:
                                print("[Meaningfulness] Quitting message sending.")
                            else:
                                print("Quitting message sending.")
                            with open(status_path, "w", encoding="utf-8") as f:
                                json.dump(sent_status, f, ensure_ascii=False, indent=2)
                            return results
                        elif action in ('s', 'send', ''):
                            break
                        elif action in ('r', 'regenerate'):
                            if verbose:
                                print("[Meaningfulness] Regenerating message...")
                            else:
                                print("Regenerating message...")
                            message = generate_low_quality_message(filename, score, text, refine)
                            if verbose:
                                print(f"\n[Meaningfulness] Regenerated message:")
                                print("-" * 50)
                                print(message)
                                print("-" * 50)
                            else:
                                print("Regenerated message.")
                        else:
                            if verbose:
                                print("[Meaningfulness] Please enter 's' to send, 'r' to regenerate, or 'q' to quit.")
                            else:
                                print("Please enter 's' to send, 'r' to regenerate, or 'q' to quit.")
                try:
                    canvas.create_conversation(
                        recipients=[str(canvas_id)],
                        subject="Yêu cầu định dạng lại và nộp lại bài tập",
                        body=message,
                        force_new=True
                    )
                    results[filename]["message_sent"] = True
                    sent_status[filename]["message_sent"] = True
                    if verbose:
                        print(f"[Meaningfulness] Message sent to student {canvas_id} for {filename}")
                    else:
                        print(f"Message sent to student {canvas_id} for {filename}")
                except Exception as e:
                    if verbose:
                        print(f"[Meaningfulness] Failed to send message for {filename}: {e}")
                    else:
                        print(f"Failed to send message for {filename}: {e}")
                    results[filename]["message_sent"] = False
                    sent_status[filename]["message_sent"] = False
        except Exception as e:
            if verbose:
                print(f"[Meaningfulness] Error setting up Canvas connection: {e}")
            else:
                print(f"Error setting up Canvas connection: {e}")
        
        with open(result_path, "a", encoding="utf-8") as f:
            f.write(f"\nMessage sending results:\n")
            f.write("-" * 80 + "\n")
            for filename, result in results.items():
                if result["meaningful_score"] < meaningfulness_threshold:
                    message_status = "SENT" if result["message_sent"] else "FAILED"
                    f.write(f"{filename}: Message {message_status}\n")
        with open(status_path, "w", encoding="utf-8") as f:
            json.dump(sent_status, f, ensure_ascii=False, indent=2)
    else:
        if verbose:
            print("[Meaningfulness] All submissions have acceptable meaningfulness scores.")
        else:
            print("All submissions have acceptable meaningfulness scores.")
        with open(status_path, "w", encoding="utf-8") as f:
            json.dump(sent_status, f, ensure_ascii=False, indent=2)
    
    return results

def extract_canvas_id_from_filename(filename):
    """
    Extract Canvas ID from filename with format: <student name>_<canvas id>_<time>_<status>.<ext>
    
    Args:
        filename (str): The PDF filename.
    
    Returns:
        int or None: Canvas ID if found, None otherwise.
    """
    parts = filename.split('_')
    for part in parts:
        if part.isdigit():
            return int(part)
    return None


def notify_missing_submissions_after_due(
    assignment_ids=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    refine=None,
    auto_send=True,
    verbose=False
):
    """
    Notify students who have not submitted for assignments whose due date has passed but are not locked.
    Returns a summary dict mapping assignment_id to missing student info.
    """
    canvas = Canvas(api_url, api_key)
    course = canvas.get_course(course_id)
    now = datetime.now(timezone.utc)

    people = list_canvas_people(api_url, api_key, course_id, verbose=verbose)
    active_students = people.get("active_students", []) if isinstance(people, dict) else []
    active_map = {str(s["canvas_id"]): s for s in active_students if s.get("canvas_id")}

    assignments = []
    group_map = {}
    try:
        groups = list(course.get_assignment_groups(include=["assignments"]))
        for g in groups:
            group_map[g.id] = g.name
            for a in g.assignments:
                assignments.append(a)
    except Exception:
        assignments = list(course.get_assignments())

    selected_assignments = []
    if assignment_ids:
        if isinstance(assignment_ids, (str, int)):
            assignment_ids = [assignment_ids]
        for a in assignments:
            if str(a.get("id") if isinstance(a, dict) else getattr(a, "id", "")) in {str(x) for x in assignment_ids}:
                selected_assignments.append(a)
    else:
        for a in assignments:
            group_name = group_map.get(a.get("assignment_group_id")) if isinstance(a, dict) else group_map.get(getattr(a, "assignment_group_id", None))
            if category and group_name and group_name.lower() != category.lower():
                continue
            selected_assignments.append(a)

    results = {}
    for assignment in selected_assignments:
        aid = assignment.get("id") if isinstance(assignment, dict) else getattr(assignment, "id", None)
        name = assignment.get("name") if isinstance(assignment, dict) else getattr(assignment, "name", "")
        due_at = assignment.get("due_at") if isinstance(assignment, dict) else getattr(assignment, "due_at", None)
        lock_at = assignment.get("lock_at") if isinstance(assignment, dict) else getattr(assignment, "lock_at", None)
        if not due_at:
            continue
        try:
            due_dt = datetime.strptime(due_at, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
        except Exception:
            continue
        if due_dt > now:
            continue
        if lock_at:
            try:
                lock_dt = datetime.strptime(lock_at, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
                if lock_dt <= now:
                    continue
            except Exception:
                pass

        assignment_obj = course.get_assignment(aid)
        submissions = list(assignment_obj.get_submissions(include=["user"]))
        submitted_ids = set()
        for sub in submissions:
            submitted_at = getattr(sub, "submitted_at", None)
            if submitted_at:
                user = getattr(sub, "user", {})
                canvas_id = user.get("id")
                if canvas_id:
                    submitted_ids.add(str(canvas_id))

        missing = []
        for canvas_id, info in active_map.items():
            if canvas_id not in submitted_ids:
                missing.append(info)

        if not missing:
            continue

        results[str(aid)] = {
            "assignment_name": name,
            "missing": missing,
        }

        for student in missing:
            canvas_id = student.get("canvas_id")
            student_name = student.get("name", "")
            if not canvas_id:
                continue
            message = (
                f"Chào {student_name},\n\n"
                f"Hệ thống ghi nhận bạn chưa nộp bài cho \"{name}\" mặc dù đã quá hạn nộp. "
                "Vui lòng nộp bài sớm nhất có thể hoặc liên hệ giảng viên nếu gặp khó khăn.\n\n"
                "Thông báo này được gửi tự động từ hệ thống."
            )
            if refine in ALL_AI_METHODS:
                prompt = (
                    "Bạn là trợ giảng. Hãy viết lại thông báo sau bằng tiếng Việt, lịch sự, rõ ràng, "
                    "nhắc sinh viên chưa nộp bài sau hạn nhưng chưa bị khóa. Trả về nội dung hoàn chỉnh.\n\n"
                    "Thông báo:\n{text}"
                )
                message = refine_text_with_ai(message, method=refine, user_prompt=prompt)
            if not auto_send:
                print(message)
                continue
            try:
                canvas.create_conversation(
                    recipients=[str(canvas_id)],
                    subject="Nhắc nhở: Chưa nộp bài quá hạn",
                    body=message,
                    force_new=True
                )
            except Exception as e:
                if verbose:
                    print(f"[MissingSubmissions] Failed to notify {student_name} ({canvas_id}): {e}")

    return results


def run_weekly_canvas_automation(
    assignment_id,
    dest_dir=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    teacher_canvas_id=None,
    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    ocr_service=DEFAULT_OCR_METHOD,
    lang="auto",
    refine=None,
    similarity_threshold=0.85,
    meaningfulness_threshold=0.4,
    auto_grade_score=10,
    notify_missing=True,
    verbose=False
):
    """
    Run weekly automation for a closed assignment:
    - Download submissions
    - Check meaningfulness + similarity and notify students
    - Assign score to clean submissions
    - Notify missing submissions for overdue assignments
    - Send summary to teacher
    """
    if not assignment_id:
        raise ValueError("assignment_id is required")

    out_dir, files = download_canvas_assignment_submissions_auto(
        assignment_id=assignment_id,
        dest_dir=dest_dir,
        api_url=api_url,
        api_key=api_key,
        course_id=course_id,
        verbose=verbose,
    )

    meaningful_results = {}
    similarity_pairs = []
    if files:
        meaningful_results = detect_meaningful_level_and_notify_students(
            out_dir,
            assignment_id=assignment_id,
            ocr_service=ocr_service,
            lang=lang,
            refine=refine,
            meaningfulness_threshold=meaningfulness_threshold,
            api_url=api_url,
            api_key=api_key,
            course_id=course_id,
            auto_send=True,
            verbose=verbose,
        )
        similarity_pairs = compare_texts_from_pdfs_in_folder(
            out_dir,
            ocr_service=ocr_service,
            lang=lang,
            refine=refine,
            similarity_threshold=similarity_threshold,
            api_url=api_url,
            api_key=api_key,
            course_id=course_id,
            auto_send=True,
            verbose=verbose,
        )

    flagged_ids = set()
    for filename, result in (meaningful_results or {}).items():
        if result.get("meaningful_score", 1.0) < meaningfulness_threshold:
            cid = extract_canvas_id_from_filename(filename)
            if cid:
                flagged_ids.add(str(cid))
    for pdf1, pdf2, _ in (similarity_pairs or []):
        for fname in (pdf1, pdf2):
            cid = extract_canvas_id_from_filename(fname)
            if cid:
                flagged_ids.add(str(cid))

    canvas = Canvas(api_url, api_key)
    course = canvas.get_course(course_id)
    assignment = course.get_assignment(assignment_id)
    assignment_name = getattr(assignment, "name", "") or f"assignment_{assignment_id}"
    safe_assignment = re.sub(r"[^A-Za-z0-9._-]+", "_", assignment_name)

    evidence_dir = None
    if flagged_ids and out_dir and os.path.isdir(out_dir):
        evidence_dir = os.path.join(
            os.getcwd(),
            f"flagged_submissions_{safe_assignment}_{assignment_id}"
        )
        os.makedirs(evidence_dir, exist_ok=True)
        for entry in os.listdir(out_dir):
            src_path = os.path.join(out_dir, entry)
            if not os.path.isfile(src_path):
                continue
            cid = extract_canvas_id_from_filename(entry)
            if cid and str(cid) in flagged_ids:
                try:
                    shutil.copy2(src_path, os.path.join(evidence_dir, entry))
                except Exception as e:
                    if verbose:
                        print(f"[WeeklyAuto] Failed to save evidence for {entry}: {e}")
    submissions = list(assignment.get_submissions(include=["user"]))
    graded = []
    skipped = []
    for sub in submissions:
        user = getattr(sub, "user", {})
        canvas_id = user.get("id")
        submitted_at = getattr(sub, "submitted_at", None)
        if not submitted_at:
            continue
        if canvas_id and str(canvas_id) in flagged_ids:
            skipped.append(str(canvas_id))
            continue
        current_score = getattr(sub, "score", None)
        if current_score is None or current_score < auto_grade_score:
            try:
                sub.edit(submission={"posted_grade": auto_grade_score})
                graded.append(str(canvas_id))
            except Exception as e:
                if verbose:
                    print(f"[WeeklyAuto] Failed to grade {canvas_id}: {e}")

    missing_summary = {}
    if notify_missing:
        missing_summary = notify_missing_submissions_after_due(
            assignment_ids=None,
            api_url=api_url,
            api_key=api_key,
            course_id=course_id,
            category=category,
            refine=refine,
            auto_send=True,
            verbose=verbose,
        )

    summary_lines = []
    summary_lines.append(f"Weekly automation summary for assignment {assignment_id}")
    summary_lines.append(f"Downloaded files: {len(files)}")
    summary_lines.append(f"Flagged (issues): {len(flagged_ids)}")
    summary_lines.append(f"Graded with {auto_grade_score}: {len(graded)}")
    summary_lines.append(f"Skipped (issues): {len(skipped)}")
    if similarity_pairs:
        summary_lines.append(f"Similarity pairs: {len(similarity_pairs)}")
    low_quality = [f for f, r in (meaningful_results or {}).items() if r.get("meaningful_score", 1.0) < meaningfulness_threshold]
    if low_quality:
        summary_lines.append(f"Low quality submissions: {len(low_quality)}")
    if missing_summary:
        summary_lines.append(f"Assignments with missing submissions: {len(missing_summary)}")

    summary_text = "\n".join(summary_lines)
    if teacher_canvas_id:
        try:
            canvas.create_conversation(
                recipients=[str(teacher_canvas_id)],
                subject="Weekly submission checks summary",
                body=summary_text,
                force_new=True,
            )
        except Exception as e:
            if verbose:
                print(f"[WeeklyAuto] Failed to send summary: {e}")

    summary_payload = {
        "assignment_id": str(assignment_id),
        "assignment_name": assignment_name,
        "generated_at": datetime.now().isoformat(),
        "flagged_count": len(flagged_ids),
        "graded_count": len(graded),
        "skipped_count": len(skipped),
    }
    try:
        with open("weekly_automation_summary.json", "w", encoding="utf-8") as f:
            json.dump(summary_payload, f, ensure_ascii=False, indent=2)
    except Exception as e:
        if verbose:
            print(f"[WeeklyAuto] Failed to write weekly summary: {e}")

    return {
        "out_dir": out_dir,
        "flagged_ids": sorted(flagged_ids),
        "graded": graded,
        "skipped": skipped,
        "missing_summary": missing_summary,
        "similarity_pairs": similarity_pairs,
        "low_quality_files": low_quality,
        "summary": summary_text,
        "evidence_dir": evidence_dir,
    }


def list_closed_assignments_for_weekly_automation(
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    category=None,
    require_submissions=True,
    verbose=False,
):
    """
    List assignments whose lock date has passed (closed) for weekly automation.
    """
    canvas = Canvas(api_url, api_key)
    course = canvas.get_course(course_id)
    try:
        groups = list(course.get_assignment_groups(include=["assignments"]))
    except Exception as e:
        if verbose:
            print(f"[WeeklyAuto] Failed to list assignments: {e}")
        return []

    group_map = {}
    for g in groups:
        group_map[g.id] = getattr(g, "name", "")

    now = datetime.now(timezone.utc)
    closed = []
    for g in groups:
        if category and getattr(g, "name", "").lower() != category.lower():
            continue
        for assignment in g.assignments:
            lock_at = assignment.get("lock_at")
            if not lock_at:
                continue
            lock_at_clean = lock_at[:-1] if lock_at.endswith("Z") else lock_at
            try:
                lock_dt = datetime.fromisoformat(lock_at_clean)
            except Exception:
                continue
            if lock_at.endswith("Z"):
                lock_dt = lock_dt.replace(tzinfo=timezone.utc)
            if lock_dt.tzinfo is None:
                lock_dt = lock_dt.replace(tzinfo=timezone.utc)
            if lock_dt > now:
                continue
            if require_submissions and not assignment.get("has_submitted_submissions", False):
                continue
            closed.append({
                "id": str(assignment.get("id")),
                "name": assignment.get("name", ""),
                "lock_at": lock_at,
                "group": group_map.get(assignment.get("assignment_group_id")),
            })

    closed.sort(key=lambda a: a.get("lock_at") or "")
    return closed

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

    def get_input_with_timeout(prompt, timeout=60, default=None):
        signal.signal(signal.SIGALRM, timeout_handler)
        signal.alarm(timeout)
        try:
            result = input(prompt)
            signal.alarm(0)
            if not result and default is not None:
                if verbose:
                    print(f"[CanvasComment] Using default: {default}")
                else:
                    print(f"Using default: {default}")
                return default
            return result
        except TimeoutError:
            signal.alarm(0)
            if default is not None:
                if verbose:
                    print(f"[CanvasComment] Using default: {default}")
                else:
                    print(f"Using default: {default}")
                return default
            raise
        except KeyboardInterrupt:
            signal.alarm(0)
            if verbose:
                print("\n[CanvasComment] Operation cancelled by user.")
            else:
                print("\nOperation cancelled by user.")
            raise

    try:
        canvas = Canvas(api_url, api_key)
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
                    sel = get_input_with_timeout(
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
                    sel = get_input_with_timeout(
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
                    line = get_input_with_timeout("", timeout=60, default="q")
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
                    confirm = get_input_with_timeout(
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

def search_canvas_user(query, api_url=CANVAS_LMS_API_URL, api_key=CANVAS_LMS_API_KEY, course_id=CANVAS_LMS_COURSE_ID, verbose=False):
    """
    Search for a user in Canvas by name or email address and print details, including only total grades and assignment category grades.

    Args:
        query (str): Name or email to search for.
        api_url (str): Canvas API base URL.
        api_key (str): Canvas API key.
        course_id (str or int): Canvas course ID.
        verbose (bool): If True, print more details; otherwise, print only important notice.
    """
    try:
        canvas = Canvas(api_url, api_key)
        course = canvas.get_course(course_id)
        enrollments = course.get_enrollments(type=['TeacherEnrollment', 'TaEnrollment', 'StudentEnrollment'])
        found = False
        query_lower = query.strip().lower()
        for enrollment in enrollments:
            user = enrollment.user
            name = user.get('name', '').lower()
            email = user.get('login_id', '').lower()
            if query_lower in name or query_lower in email:
                if verbose:
                    print("[CanvasUser] User found:")
                else:
                    print("User found:")
                print(f"  Name: {user.get('name', '')}")
                print(f"  Email: {user.get('login_id', '')}")
                print(f"  Canvas ID: {user.get('id', '')}")
                print(f"  Role: {enrollment.type}")
                print(f"  Enrollment State: {enrollment.enrollment_state}")
                # Only print total grades and assignment category grades if student
                if enrollment.type.lower() == "studentenrollment":
                    try:
                        user_id = user.get('id')
                        total_grade = getattr(enrollment, "grades", {}).get("final_grade", None)
                        total_score = getattr(enrollment, "grades", {}).get("final_score", None)
                        print("  Grades:")
                        print(f"    - Total Grade: {total_grade} (Score: {total_score})")
                        assignment_groups = course.get_assignment_groups(include=['assignments', 'score_statistics'])
                        for group in assignment_groups:
                            group_score = None
                            group_possible = None
                            if hasattr(group, "score_statistics") and group.score_statistics:
                                stats = group.score_statistics
                                user_stats = stats.get(str(user_id))
                                if user_stats:
                                    group_score = user_stats.get("score")
                                    group_possible = user_stats.get("possible")
                            if group_score is None:
                                group_score = 0
                                group_possible = 0
                                for assignment in group.assignments:
                                    try:
                                        submission = assignment.get_submission(user_id)
                                        score = submission.score if hasattr(submission, "score") else None
                                        if score is not None:
                                            group_score += float(score)
                                        group_possible += float(getattr(assignment, "points_possible", 0))
                                    except Exception:
                                        continue
                            print(f"    - {group.name}: {group_score}/{group_possible}")
                        if verbose:
                            print(f"[CanvasUser] Printed grades for user {user.get('name', '')}")
                    except Exception as e:
                        if verbose:
                            print(f"[CanvasUser] Could not fetch assignments/grades: {e}")
                        else:
                            print(f"  Could not fetch assignments/grades: {e}")
                found = True
        if not found:
            if verbose:
                print("[CanvasUser] No user found matching the query.")
            else:
                print("No user found matching the query.")
    except Exception as e:
        if verbose:
            print(f"[CanvasUser] Error searching for user: {e}")
        else:
            print(f"Error searching for user: {e}")

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

        canvas = Canvas(api_url, api_key)
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

def add_canvas_announcement(
    title,
    message,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    verbose=False
):
    """
    Add an announcement to a Canvas course.
    Opens a temporary file for editing title and message, then posts to Canvas.
    Deletes the temporary file after posting.
    If both title and message are None or empty, prompt the user to edit the temporary file with placeholder text.
    If verbose is True, print more details; otherwise, print only important notice.
    """

    def timeout_handler(signum, frame):
        if verbose:
            print("\n[CanvasAnnouncement] Timeout: No response after 60 seconds. Using default option.")
        else:
            print("\nTimeout: No response after 60 seconds. Using default option.")
        raise TimeoutError("User input timeout")

    def get_input_with_timeout(prompt, timeout=60, default=None):
        signal.signal(signal.SIGALRM, timeout_handler)
        signal.alarm(timeout)
        try:
            result = input(prompt)
            signal.alarm(0)
            if not result and default is not None:
                if verbose:
                    print(f"[CanvasAnnouncement] Using default: {default}")
                else:
                    print(f"Using default: {default}")
                return default
            return result
        except TimeoutError:
            signal.alarm(0)
            if default is not None:
                if verbose:
                    print(f"[CanvasAnnouncement] Using default: {default}")
                else:
                    print(f"Using default: {default}")
                return default
            raise
        except KeyboardInterrupt:
            signal.alarm(0)
            if verbose:
                print("\n[CanvasAnnouncement] Operation cancelled by user.")
            else:
                print("\nOperation cancelled by user.")
            raise

    temp_path = None
    try:
        # If both title and message are None or empty, use placeholder text
        if not (title and title.strip()) and not (message and message.strip()):
            initial_title = "Announcement Title (edit this line)"
            initial_message = "Enter your announcement message here (edit below)."
        else:
            initial_title = title or "Announcement Title"
            initial_message = message or "Enter your announcement message here."

        temp_content = f"Title: {initial_title}\n\nMessage:\n{initial_message}\n"

        # Create a temporary file for editing in the current directory
        temp_dir = os.getcwd()
        with tempfile.NamedTemporaryFile(mode="w+", suffix=".txt", delete=False, encoding="utf-8", dir=temp_dir) as tmpf:
            tmpf.write(temp_content)
            tmpf.flush()
            temp_path = tmpf.name

        if verbose:
            print(f"[CanvasAnnouncement] Edit the announcement in your editor: {temp_path}")
            print("Format: The first line starting with 'Title:' is the title. The rest after 'Message:' is the message body.")
        else:
            print(f"Edit the announcement in your editor: {temp_path}")
            print("Format: The first line starting with 'Title:' is the title. The rest after 'Message:' is the message body.")
        default_editor = "notepad" if os.name == "nt" else ("nano" if shutil.which("nano") else "vi")
        editor = os.environ.get("EDITOR", default_editor)
        try:
            subprocess.call([editor, temp_path])
        except Exception as e:
            if verbose:
                print(f"[CanvasAnnouncement] Could not open editor: {e}")
            else:
                print(f"Could not open editor: {e}")
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
            return {"error": str(e)}

        # Wait for the user to finish editing before reading the file
        get_input_with_timeout("Press Enter after you have finished editing and saved the file...", timeout=60, default="")

        # Read edited content
        with open(temp_path, "r", encoding="utf-8") as f:
            lines = f.readlines()

        # Parse title and message
        new_title = ""
        new_message = []
        in_message = False
        for line in lines:
            if line.strip().lower().startswith("title:"):
                new_title = line.strip()[6:].strip()
            elif line.strip().lower().startswith("message:"):
                in_message = True
            elif in_message:
                # Keep all lines including empty lines and newlines
                new_message.append(line.rstrip('\n'))
        # Join with '\n' to preserve original formatting, including empty lines
        new_message = "\n".join(new_message)

        # Confirm submission
        if verbose:
            print("\n[CanvasAnnouncement] Preview of announcement to submit:")
            print(f"Title: {new_title}")
            print("Message:")
            print(new_message[:500] + ("..." if len(new_message) > 500 else ""))
        else:
            print("\nPreview of announcement to submit:")
            print(f"Title: {new_title}")
            print("Message:")
            print(new_message[:500] + ("..." if len(new_message) > 500 else ""))
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
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
            return {"error": "User cancelled."}

        # Post to Canvas
        canvas = Canvas(api_url, api_key)
        course = canvas.get_course(course_id)
        discussion = course.create_discussion_topic(
            title=new_title,
            message=new_message,
            is_announcement=True
        )
        if verbose:
            print(f"[CanvasAnnouncement] Announcement '{new_title}' created for course {course_id}.")
        else:
            print(f"Announcement '{new_title}' created for course {course_id}.")
        return discussion
    except Exception as e:
        if verbose:
            print(f"[CanvasAnnouncement] Failed to create announcement: {e}")
        else:
            print(f"Failed to create announcement: {e}")
        return {"error": str(e)}
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

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

    def get_input_with_timeout(prompt, timeout=60, default=None):
        signal.signal(signal.SIGALRM, timeout_handler)
        signal.alarm(timeout)
        try:
            result = input(prompt)
            signal.alarm(0)
            if not result and default is not None:
                if verbose:
                    print(f"[CanvasMessage] Using default: {default}")
                else:
                    print(f"Using default: {default}")
                return default
            return result
        except TimeoutError:
            signal.alarm(0)
            if default is not None:
                if verbose:
                    print(f"[CanvasMessage] Using default: {default}")
                else:
                    print(f"Using default: {default}")
                return default
            raise
        except KeyboardInterrupt:
            signal.alarm(0)
            if verbose:
                print("\n[CanvasMessage] Operation cancelled by user.")
            else:
                print("\nOperation cancelled by user.")
            raise

    try:
        canvas = Canvas(api_url, api_key)
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
                sel = get_input_with_timeout("Enter selection: ", timeout=60, default="q").strip()
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
                line = get_input_with_timeout("", timeout=60, default="q")
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
                confirm = get_input_with_timeout(
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

    def get_input_with_timeout(prompt, timeout=60, default=None):
        # Use signal.SIGALRM only if available (not on Windows)
        if hasattr(signal, "SIGALRM"):
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(timeout)
            try:
                result = input(prompt)
                signal.alarm(0)
                if not result and default is not None:
                    print(f"Using default: {default}")
                    return default
                return result
            except TimeoutError:
                signal.alarm(0)
                print(f"Using default: {default}")
                return default
            except KeyboardInterrupt:
                signal.alarm(0)
                print("\nOperation cancelled by user.")
                raise
        else:
            # Fallback for platforms without SIGALRM (e.g., Windows)
            result = input(prompt)
            if not result and default is not None:
                print(f"Using default: {default}")
                return default
            return result

    try:
        canvas = Canvas(api_url, api_key)
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
                sel = get_input_with_timeout(
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
            send_messages = get_input_with_timeout(
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
                        confirm = get_input_with_timeout(
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
                        confirm = get_input_with_timeout(
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

def grade_canvas_assignment_submissions(
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    verbose=False,
    restricted=DEFAULT_RESTRICTED  # New option: if False, list all assignments with submissions and all students who submitted
):
    """
    List all assignments with at least one submission and whose lock date has passed (if restricted=True),
    or all assignments with at least one submission (if restricted=False).
    Allow user to select one or more (supports ranges).
    For each selected assignment, list all students who submitted but have not yet been given scores (if restricted=True),
    or all students who submitted (if restricted=False).
    Allow user to select one or more (supports ranges), then prompt for a score (optionally same for all), and set the score for those submissions.
    Allow quitting at any step.
    If verbose is True, print more details; otherwise, print only important notice.
    """
    def get_input_with_quit(prompt, default=None):
        try:
            val = input(prompt)
            if not val and default is not None:
                return default
            if val.strip().lower() in ("q", "quit"):
                return None
            return val
        except KeyboardInterrupt:
            return None

    try:
        canvas = Canvas(api_url, api_key)
        course = canvas.get_course(course_id)

        # Step 1: List assignments with at least one submission
        assignments = []
        assignment_groups = list(course.get_assignment_groups(include=['assignments']))
        now = datetime.utcnow()
        for group in assignment_groups:
            group_name = group.name
            if category and group_name.lower() != category.lower():
                continue
            for assignment in group.assignments:
                if assignment.get('has_submitted_submissions', False):
                    if restricted:
                        lock_at = assignment.get('lock_at')
                        if lock_at:
                            try:
                                lock_dt = datetime.strptime(lock_at, "%Y-%m-%dT%H:%M:%SZ")
                                if lock_dt > now:
                                    continue  # Skip if lock date is in the future
                            except Exception:
                                continue  # Skip if lock date is invalid
                        else:
                            continue  # Skip if no lock date
                    assignments.append({
                        "id": assignment['id'],
                        "name": assignment['name'],
                        "group": group_name,
                        "due_at": assignment.get('due_at'),
                        "lock_at": assignment.get('lock_at')
                    })
        if not assignments:
            msg = "No assignments found with submissions and lock date passed." if restricted else "No assignments found with submissions."
            if verbose:
                print(f"[GradeCanvas] {msg}")
            else:
                print(msg)
            return

        # Sort by due date
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
            print("[GradeCanvas] Assignments with at least one submission{}:".format(" and lock date passed" if restricted else ""))
        else:
            print("Assignments with at least one submission{}:".format(" and lock date passed" if restricted else ""))
        for idx, a in enumerate(assignments, 1):
            due = a['due_at'] or "No due date"
            lock = a['lock_at'] or "No lock date"
            if verbose:
                print(f"[GradeCanvas] {idx}. [{a['group']}] {a['name']} (ID: {a['id']}, Due: {due}, Lock: {lock})")
            else:
                print(f"{idx}. [{a['group']}] {a['name']} (ID: {a['id']}, Due: {due}, Lock: {lock})")

        # Step 2: Select assignments to grade (supports ranges, comma-separated, 'a' for all)
        while True:
            sel = get_input_with_quit(
                "Enter the number(s) of the assignment(s) to grade (e.g. 1,3-5 or 'a' for all, or 'q' to quit): "
            )
            if sel is None:
                if verbose:
                    print("[GradeCanvas] Quitting.")
                else:
                    print("Quitting.")
                return
            if sel.lower() in ('a', 'all'):
                selected_assignments = list(range(1, len(assignments) + 1))
            else:
                selected_assignments = set()
                for part in sel.split(","):
                    part = part.strip()
                    if "-" in part:
                        try:
                            start, end = map(int, part.split("-"))
                            selected_assignments.update(range(start, end + 1))
                        except Exception:
                            continue
                    elif part.isdigit():
                        selected_assignments.add(int(part))
                selected_assignments = [i for i in selected_assignments if 1 <= i <= len(assignments)]
            if not selected_assignments:
                msg = "Invalid selection. Please enter valid numbers, a range, 'a' for all, or 'q' to quit."
                if verbose:
                    print(f"[GradeCanvas] {msg}")
                else:
                    print(msg)
                continue
            break

        # Step 3: For each selected assignment, process grading
        for assign_idx in selected_assignments:
            assignment_info = assignments[assign_idx - 1]
            assignment_id = assignment_info['id']
            assignment = course.get_assignment(assignment_id)
            if verbose:
                print(f"\n[GradeCanvas] Grading assignment: [{assignment_info['group']}] {assignment_info['name']} (ID: {assignment_id})")
            else:
                print(f"\n--- Grading assignment: [{assignment_info['group']}] {assignment_info['name']} (ID: {assignment_id}) ---")

            # List all students who submitted
            submissions = list(assignment.get_submissions(include=["user"]))
            students_to_grade = []
            for sub in submissions:
                user = getattr(sub, "user", {})
                student_name = user.get("name", "UnknownStudent")
                canvas_id = user.get("id", "unknown")
                submitted_at = getattr(sub, "submitted_at", None)
                score = getattr(sub, "score", None)
                # Only include if submitted (on time or late)
                if submitted_at:
                    if restricted:
                        # Only include if score is None or 0
                        if score is None or score == "0":
                            students_to_grade.append({
                                "canvas_id": canvas_id,
                                "name": student_name,
                                "submitted_at": submitted_at,
                                "submission": sub
                            })
                    else:
                        students_to_grade.append({
                            "canvas_id": canvas_id,
                            "name": student_name,
                            "submitted_at": submitted_at,
                            "score": score,
                            "submission": sub
                        })
            if not students_to_grade:
                msg = "No students have{} submissions for this assignment.".format(" ungraded" if restricted else "")
                if verbose:
                    print(f"[GradeCanvas] {msg}")
                else:
                    print(msg)
                continue

            students_to_grade.sort(key=lambda s: s["name"])
            if verbose:
                print("[GradeCanvas] Students who submitted{}:".format(" but have not yet been graded" if restricted else ""))
            else:
                print("Students who submitted{}:".format(" but have not yet been graded" if restricted else ""))
            for idx, s in enumerate(students_to_grade, 1):
                submitted = s["submitted_at"] or "No submission"
                score_str = f", Current Score: {s.get('score', '')}" if not restricted else ""
                if verbose:
                    print(f"[GradeCanvas] {idx}. {s['name']} (Canvas ID: {s['canvas_id']}), Submitted at: {submitted}{score_str}")
                else:
                    print(f"{idx}. {s['name']} (Canvas ID: {s['canvas_id']}), Submitted at: {submitted}{score_str}")

            # Select students to grade (supports ranges, comma-separated, 'a' for all)
            while True:
                sel = get_input_with_quit(
                    "Enter the number(s) of the student(s) to grade (e.g. 1,3-5 or 'a' for all, or 'q' to quit): "
                )
                if sel is None:
                    if verbose:
                        print("[GradeCanvas] Quitting grading for this assignment.")
                    else:
                        print("Quitting grading for this assignment.")
                    break
                if sel.lower() in ('a', 'all'):
                    selected = list(range(1, len(students_to_grade) + 1))
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
                    selected = [i for i in selected if 1 <= i <= len(students_to_grade)]
                if not selected:
                    msg = "Invalid selection. Please enter valid numbers, a range, 'a' for all, or 'q' to quit."
                    if verbose:
                        print(f"[GradeCanvas] {msg}")
                    else:
                        print(msg)
                    continue
                break
            if not selected:
                continue

            # Ask for score to give (optionally same for all)
            score = get_input_with_quit("Enter the score to assign to all selected submissions (or 'q' to quit): ")
            if score is None:
                if verbose:
                    print("[GradeCanvas] Quitting grading for this assignment.")
                else:
                    print("Quitting grading for this assignment.")
                continue
            try:
                score = float(score)
            except Exception:
                msg = "Invalid score. Skipping this assignment."
                if verbose:
                    print(f"[GradeCanvas] {msg}")
                else:
                    print(msg)
                continue

            # Confirm and assign scores
            msg = f"Assigning score {score} to {len(selected)} submission(s) for assignment '{assignment_info['name']}'."
            if verbose:
                print(f"[GradeCanvas] {msg}")
            else:
                print(msg)
            confirm = get_input_with_quit("Proceed? (y/n): ", default="y")
            if confirm is None or confirm.lower() not in ("y", "yes", ""):
                if verbose:
                    print("[GradeCanvas] Aborted grading for this assignment.")
                else:
                    print("Aborted grading for this assignment.")
                continue

            for idx in selected:
                s = students_to_grade[idx - 1]
                sub = s["submission"]
                try:
                    sub.edit(submission={'posted_grade': score})
                    if verbose:
                        print(f"[GradeCanvas] Assigned score {score} to {s['name']} (Canvas ID: {s['canvas_id']})")
                    else:
                        print(f"Assigned score {score} to {s['name']} (Canvas ID: {s['canvas_id']})")
                except Exception as e:
                    if verbose:
                        print(f"[GradeCanvas] Failed to assign score to {s['name']} (Canvas ID: {s['canvas_id']}): {e}")
                    else:
                        print(f"Failed to assign score to {s['name']} (Canvas ID: {s['canvas_id']}): {e}")

            if verbose:
                print(f"[GradeCanvas] Done grading selected submissions for assignment '{assignment_info['name']}'.")
            else:
                print(f"Done grading selected submissions for assignment '{assignment_info['name']}'.")

        if verbose:
            print("[GradeCanvas] Finished grading all selected assignments.")
        else:
            print("Finished grading all selected assignments.")

    except Exception as e:
        if verbose:
            print(f"[GradeCanvas] Error grading submissions: {e}")
        else:
            print(f"Error grading submissions: {e}")

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
        canvas = Canvas(api_url, api_key)
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
        # Use input with timeout if possible, fallback to normal input on platforms without SIGALRM
        def get_input_with_timeout(prompt, timeout=60, default=None):
            if hasattr(signal, "SIGALRM"):
                def timeout_handler(signum, frame):
                    print("\nTimeout: No response after 60 seconds. Using default option.")
                    raise TimeoutError("User input timeout")
                signal.signal(signal.SIGALRM, timeout_handler)
                signal.alarm(timeout)
                try:
                    result = input(prompt)
                    signal.alarm(0)
                    if not result and default is not None:
                        print(f"Using default: {default}")
                        return default
                    return result
                except TimeoutError:
                    signal.alarm(0)
                    if default is not None:
                        print(f"Using default: {default}")
                        return default
                    raise
                except KeyboardInterrupt:
                    signal.alarm(0)
                    print("\nOperation cancelled by user.")
                    raise
            else:
                # On platforms without SIGALRM (e.g., Windows), just use input without timeout
                result = input(prompt)
                if not result and default is not None:
                    print(f"Using default: {default}")
                    return default
                return result

        sel = get_input_with_timeout("Enter conversation numbers to reply (comma/range, 'a' for all, or 'q' to quit): ", timeout=60, default="q").strip()
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

    def get_input_with_timeout(prompt, timeout=60, default=None):
        # Use signal.SIGALRM only if available (not on Windows)
        if hasattr(signal, "SIGALRM"):
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(timeout)
            try:
                result = input(prompt)
                signal.alarm(0)
                if not result and default is not None:
                    if verbose:
                        print(f"[CanvasPages] Using default: {default}")
                    else:
                        print(f"Using default: {default}")
                    return default
                return result
            except TimeoutError:
                signal.alarm(0)
                if default is not None:
                    if verbose:
                        print(f"[CanvasPages] Using default: {default}")
                    else:
                        print(f"Using default: {default}")
                    return default
                raise
            except KeyboardInterrupt:
                signal.alarm(0)
                if verbose:
                    print("\n[CanvasPages] Operation cancelled by user.")
                else:
                    print("\nOperation cancelled by user.")
                raise
        else:
            # On platforms without SIGALRM (e.g., Windows), just use input without timeout
            result = input(prompt)
            if not result and default is not None:
                if verbose:
                    print(f"[CanvasPages] Using default: {default}")
                else:
                    print(f"Using default: {default}")
                return default
            return result

    try:
        # Ensure api_url is valid
        if not api_url or not isinstance(api_url, str) or not api_url.startswith(("http://", "https://")):
            if verbose:
                print(f"[CanvasPages] Error: Canvas API URL is missing or invalid. Please provide a valid HTTP or HTTPS URL for the Canvas instance. {api_url}")
            else:
                print(f"Error: Canvas API URL is missing or invalid.")
            return
        canvas = Canvas(api_url, api_key)
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
        sel = get_input_with_timeout(
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
            action = get_input_with_timeout(
                "What do you want to do with this page? (e/d/s): ",
                timeout=60,
                default="s"
            ).strip().lower()
            if action == "d":
                confirm = get_input_with_timeout(
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
                    refine_choice = get_input_with_timeout(
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
                confirm = get_input_with_timeout(
                    "Use refined content? (y/n): ",
                    timeout=60,
                    default="y"
                ).strip().lower()
                if confirm == "y":
                    new_body = refined_body

            # Ask if user wants to update the title
            update_title = get_input_with_timeout(
                f"Do you want to update the page title? (current: '{page.title}') (y/n): ",
                timeout=60,
                default="n"
            ).strip().lower()
            new_title = page.title
            if update_title == "y":
                t = get_input_with_timeout(
                    "Enter new title (leave blank to keep current): ",
                    timeout=60,
                    default=""
                ).strip()
                if t:
                    new_title = t

            # Ask for optional parameters
            editing_roles = get_input_with_timeout(
                "Enter editing roles (comma-separated, e.g. teachers,students) or leave blank to keep unchanged: ",
                timeout=60,
                default=""
            ).strip()
            notify_of_update = get_input_with_timeout(
                "Notify participants of update? (y/n, default n): ",
                timeout=60,
                default="n"
            ).strip().lower()
            published = get_input_with_timeout(
                "Publish page? (y/n, default y): ",
                timeout=60,
                default="y"
            ).strip().lower()
            front_page = get_input_with_timeout(
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

def list_students_with_multiple_submissions_on_time(
    assignment_id=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    verbose=False
):
    """
    List all students who submitted at least twice to an assignment,
    where the first submission is on time and the second submission is late.
    If assignment_id is None, list all assignments in the specified category and ask user to select one or more.

    For students who have first submission on time and from second submission late,
    update the late_policy_status from the second submission to "none".

    Args:
        assignment_id (str or int or None): Canvas assignment ID. If None, prompt user to select.
        api_url, api_key, course_id: Canvas API info.
        category (str): Assignment group/category to filter.
        verbose (bool): Print more details.

    Returns:
        List of dicts: [{canvas_id, name, submissions: [submission_times], first_on_time: True/False, second_late: True/False}, ...]
    """

    def timeout_handler(signum, frame):
        # Arguments are required by signal handler signature but not used
        print("\nTimeout: No response after 60 seconds. Using default: all")
        raise TimeoutError("User input timeout")

    def get_input_with_timeout(prompt, timeout=60, default=None):
        # Use signal.SIGALRM only if available (not on Windows)
        if hasattr(signal, "SIGALRM"):
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(timeout)
            try:
                result = input(prompt)
                signal.alarm(0)
                if not result and default is not None:
                    print(f"Using default: {default}")
                    return default
                return result
            except TimeoutError:
                signal.alarm(0)
                if default is not None:
                    print(f"Using default: {default}")
                    return default
                raise
            except KeyboardInterrupt:
                signal.alarm(0)
                print("\nOperation cancelled by user.")
                raise
        else:
            # On platforms without SIGALRM (e.g., Windows), just use input without timeout
            result = input(prompt)
            if not result and default is not None:
                print(f"Using default: {default}")
                return default
            return result

    try:
        canvas = Canvas(api_url, api_key)
        course = canvas.get_course(course_id)

        # If assignment_id is None, list all assignments in the specified category and prompt user
        assignment_ids = []
        if assignment_id is None:
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
                msg = f"No assignments found with submissions in category '{category}'." if category else "No assignments found."
                if verbose:
                    print(f"[MultipleSubmissions] {msg}")
                else:
                    print(msg)
                return []
            # Sort by due date
            def due_sort_key(a):
                raw = a.get("due_at")
                if raw:
                    try:
                        return datetime.strptime(raw, "%Y-%m-%dT%H:%M:%SZ")
                    except Exception:
                        return datetime.max
                return datetime.max
            assignments.sort(key=due_sort_key)
            print("Assignments with at least one submission:")
            for idx, a in enumerate(assignments, 1):
                due = a['due_at'] or "No due date"
                print(f"{idx}. [{a['group']}] {a['name']} (ID: {a['id']}, Due: {due})")
            while True:
                try:
                    sel = get_input_with_timeout(
                        "Enter the number(s) of the assignment(s) to check (e.g. 1,3-5 or 'a' for all, or 'q' to quit): ",
                        timeout=60,
                        default="a"
                    ).strip()
                except TimeoutError:
                    sel = "a"
                if sel.lower() in ('q', 'quit'):
                    return []
                if sel.lower() in ('a', 'all'):
                    assignment_ids = [a['id'] for a in assignments]
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
                assignment_ids = [assignments[i - 1]['id'] for i in selected]
                break
        else:
            assignment_ids = [assignment_id]

        all_results = []
        for aid in assignment_ids:
            assignment = course.get_assignment(aid)
            due_at = getattr(assignment, "due_at", None)
            if due_at:
                try:
                    due_dt = datetime.strptime(due_at, "%Y-%m-%dT%H:%M:%SZ")
                except Exception:
                    due_dt = None
            else:
                due_dt = None

            submissions = list(assignment.get_submissions(include=["user", "submission_history"]))
            results = []
            for sub in submissions:
                user = getattr(sub, "user", {})
                canvas_id = user.get("id", None)
                name = user.get("name", "")
                # Get all submission attempts (submission_history)
                history = getattr(sub, "submission_history", None)
                if not history:
                    # Fallback: treat as single submission
                    submitted_at = getattr(sub, "submitted_at", None)
                    if submitted_at:
                        times = [submitted_at]
                        histories = [sub]
                    else:
                        times = []
                        histories = []
                else:
                    times = [h.get("submitted_at") for h in history if h.get("submitted_at")]
                    histories = [h for h in history if h.get("submitted_at")]
                # Only consider students with 2 or more submissions
                if len(times) < 2:
                    continue
                # Sort submission times chronologically, keep mapping to histories
                times_histories = sorted(zip(times, histories), key=lambda x: x[0])
                times_sorted = [t for t, h in times_histories]
                histories_sorted = [h for t, h in times_histories]
                first_time = times_sorted[0]
                second_time = times_sorted[1]
                if not first_time or not second_time or not due_dt:
                    continue
                try:
                    first_dt = datetime.strptime(first_time, "%Y-%m-%dT%H:%M:%SZ")
                    second_dt = datetime.strptime(second_time, "%Y-%m-%dT%H:%M:%SZ")
                except Exception:
                    continue
                first_on_time = first_dt <= due_dt
                second_late = second_dt > due_dt
                if first_on_time and second_late:
                    # Update late_policy_status from second submission onward to "none"
                    for idx in range(1, len(histories_sorted)):
                        h = histories_sorted[idx]
                        # Only update if late_policy_status is not "none"
                        if h.get("late_policy_status") != "none":
                            try:
                                # Use the API to update the submission's late_policy_status
                                # Canvas API: PUT /api/v1/courses/:course_id/assignments/:assignment_id/submissions/:user_id
                                url = f"{api_url}/api/v1/courses/{course_id}/assignments/{aid}/submissions/{canvas_id}"
                                headers = {
                                    "Authorization": f"Bearer {api_key}",
                                    "Content-Type": "application/json"
                                }
                                data = {
                                    "submission": {
                                        "late_policy_status": "none"
                                    }
                                }
                                resp = requests.put(url, headers=headers, json=data)
                                if verbose:
                                    print(f"[MultipleSubmissions] Updated late_policy_status to 'none' for student {name} (ID: {canvas_id}), submission at {histories_sorted[idx].get('submitted_at')}, status code: {resp.status_code}")
                            except Exception as e:
                                if verbose:
                                    print(f"[MultipleSubmissions] Failed to update late_policy_status for student {name} (ID: {canvas_id}): {e}")
                    results.append({
                        "canvas_id": canvas_id,
                        "name": name,
                        "submissions": times_sorted,
                        "first_on_time": True,
                        "second_late": True,
                        "assignment_id": aid,
                        "assignment_name": getattr(assignment, "name", "")
                    })
            if verbose:
                print(f"[MultipleSubmissions] Assignment {aid}: Found {len(results)} students with first submission on time and second late:")
                for r in results:
                    print(f"  {r['name']} (ID: {r['canvas_id']}): {len(r['submissions'])} submissions, first: {r['submissions'][0]}, second: {r['submissions'][1]}")
            else:
                print(f"Assignment {aid}: Found {len(results)} students with first submission on time and second late.")
            all_results.extend(results)
        return all_results
    except Exception as e:
        if verbose:
            print(f"[MultipleSubmissions] Error: {e}")
        else:
            print(f"Error: {e}")
        return []

def list_and_export_canvas_rubrics(
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    assignment_id=None,
    export_path=None,
    verbose=False
):
    """
    List all unique rubrics used in a Canvas course (optionally for a specific assignment) and export to TXT/CSV file.
    Uses Canvas API GET /api/v1/courses/:course_id/rubrics/:id for full rubric info.
    If export_path is provided and ends with .csv, downloads the official Canvas rubric CSV template and fills it with rubric info.
    The exported CSV matches the Canvas rubric import template (GET /api/v1/rubrics/upload_template).
    """
    try:
        canvas = Canvas(api_url, api_key)
        course = canvas.get_course(course_id)
        rubric_ids = set()
        rubric_details = {}

        # Helper to fetch rubric details via API
        def fetch_rubric(rubric_id):
            url = f"{api_url}/api/v1/courses/{course_id}/rubrics/{rubric_id}"
            headers = {"Authorization": f"Bearer {api_key}"}
            params = {"include[]": "associations"}
            resp = requests.get(url, headers=headers, params=params)
            if resp.status_code == 200:
                return resp.json()
            return None

        if assignment_id:
            assignment = course.get_assignment(assignment_id)
            rubric_settings = getattr(assignment, "rubric_settings", {})
            rubric_id = rubric_settings.get("id") or getattr(assignment, "rubric_id", None)
            if rubric_id:
                rubric_ids.add(rubric_id)
        else:
            for assignment in course.get_assignments():
                rubric_settings = getattr(assignment, "rubric_settings", {})
                rubric_id = rubric_settings.get("id") or getattr(assignment, "rubric_id", None)
                if rubric_id:
                    rubric_ids.add(rubric_id)

        if not rubric_ids:
            print("[Rubrics] No rubrics found in this course." if not verbose else "[Rubrics] No rubrics found in this course.")
            return []

        # Fetch rubric details for each rubric_id
        for rid in rubric_ids:
            rubric = fetch_rubric(rid)
            if rubric:
                rubric_details[rid] = rubric

        if not rubric_details:
            print("[Rubrics] No rubric details found." if not verbose else "[Rubrics] No rubric details found.")
            return []

        # Print rubric info in Canvas API assessment format
        for rid, rubric in rubric_details.items():
            title = rubric.get("title", "")
            print(f"\nRubric ID: {rid} | Title: {title}")
            print("Format for assessmentRubricAssessmentsController#create:")
            print(f"POST /api/v1/courses/{course_id}/rubric_associations/{rid}/rubric_assessments")
            print("rubric_assessment[user_id]: <user_id>")
            print("rubric_assessment[assessment_type]: grading|peer_review|provisional_grade")
            criteria = rubric.get("data", [])
            for idx, criterion in enumerate(criteria, 1):
                crit_id = criterion.get("id", "")
                desc = criterion.get("description", "")
                long_desc = criterion.get("long_description", "")
                points = criterion.get("points", "")
                print(f"  criterion_{crit_id}[points]: <points_awarded>  # {desc} ({points} pts)")
                print(f"  criterion_{crit_id}[comments]: <comments>      # {long_desc}")
                ratings = criterion.get("ratings", [])
                for rating in ratings or []:
                    print(f"     - Rating: {rating.get('description', '')}: {rating.get('points', '')} pts")

        # Export to file if requested
        if export_path:
            ext = os.path.splitext(export_path)[1].lower()
            if ext == ".csv":
                # Download Canvas rubric template
                template_url = f"{api_url}/api/v1/rubrics/upload_template"
                headers = {"Authorization": f"Bearer {api_key}"}
                resp = requests.get(template_url, headers=headers)
                if resp.status_code != 200:
                    print("[Rubrics] Failed to download rubric template.")
                    return []
                template_csv = resp.text
                template_reader = csv.reader(io.StringIO(template_csv))
                template_rows = list(template_reader)
                header = template_rows[0]
                # Find correct column names for rubric title and criterion description
                rubric_title_col = None
                criterion_desc_col = None
                points_col = None
                long_desc_col = None
                rating_desc_col = None
                rating_points_col = None
                for idx, col in enumerate(header):
                    col_lower = col.strip().lower()
                    if "rubric name" in col_lower or "rubric title" in col_lower:
                        rubric_title_col = idx
                    elif "criterion description" in col_lower:
                        criterion_desc_col = idx
                    elif col_lower == "points":
                        points_col = idx
                    elif "long description" in col_lower:
                        long_desc_col = idx
                    elif "rating description" in col_lower:
                        rating_desc_col = idx
                    elif "rating points" in col_lower:
                        rating_points_col = idx
                rubric_rows = []
                for rid, rubric in rubric_details.items():
                    title = rubric.get("title", "")
                    for criterion in rubric.get("data", []):
                        desc = criterion.get("description", "")
                        long_desc = criterion.get("long_description", "")
                        points = criterion.get("points", "")
                        ratings = criterion.get("ratings", [])
                        if ratings:
                            for rating in ratings:
                                row = ["" for _ in header]
                                if rubric_title_col is not None:
                                    row[rubric_title_col] = title
                                if criterion_desc_col is not None:
                                    row[criterion_desc_col] = desc
                                if points_col is not None:
                                    row[points_col] = rating.get("points", points)
                                if long_desc_col is not None:
                                    row[long_desc_col] = long_desc
                                if rating_desc_col is not None:
                                    row[rating_desc_col] = rating.get("description", "")
                                if rating_points_col is not None:
                                    row[rating_points_col] = rating.get("points", "")
                                rubric_rows.append(row)
                        else:
                            row = ["" for _ in header]
                            if rubric_title_col is not None:
                                row[rubric_title_col] = title
                            if criterion_desc_col is not None:
                                row[criterion_desc_col] = desc
                            if points_col is not None:
                                row[points_col] = points
                            if long_desc_col is not None:
                                row[long_desc_col] = long_desc
                            rubric_rows.append(row)
                with open(export_path, "w", encoding="utf-8", newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(header)
                    for row in rubric_rows:
                        writer.writerow(row)
                print(f"Rubrics exported to {export_path}" if not verbose else f"[Rubrics] Rubrics exported to {export_path}")
            else:
                with open(export_path, "w", encoding="utf-8") as f:
                    for rid, rubric in rubric_details.items():
                        title = rubric.get("title", "")
                        f.write(f"Rubric ID: {rid} | Title: {title}\n")
                        f.write(f"Format for assessmentRubricAssessmentsController#create:\n")
                        f.write(f"POST /api/v1/courses/{course_id}/rubric_associations/{rid}/rubric_assessments\n")
                        f.write("rubric_assessment[user_id]: <user_id>\n")
                        f.write("rubric_assessment[assessment_type]: grading|peer_review|provisional_grade\n")
                        for criterion in rubric.get("data", []):
                            crit_id = criterion.get("id", "")
                            desc = criterion.get("description", "")
                            long_desc = criterion.get("long_description", "")
                            points = criterion.get("points", "")
                            f.write(f"  criterion_{crit_id}[points]: <points_awarded>  # {desc} ({points} pts)\n")
                            f.write(f"  criterion_{crit_id}[comments]: <comments>      # {long_desc}\n")
                            ratings = criterion.get("ratings", [])
                            for rating in ratings or []:
                                f.write(f"     - Rating: {rating.get('description', '')}: {rating.get('points', '')} pts\n")
                        f.write("\n")
                print(f"Rubrics exported to {export_path}" if not verbose else f"[Rubrics] Rubrics exported to {export_path}")

        # Return rubric details as a list
        return [rubric for rubric in rubric_details.values()]
    except Exception as e:
        print(f"Error listing or exporting rubrics: {e}" if not verbose else f"[Rubrics] Error listing or exporting rubrics: {e}")
        return []

def import_canvas_rubrics(
    rubric_file,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    verbose=False
):
    """
    Import rubrics to Canvas course using the official rubric CSV upload API.
    If rubric_file is a CSV, directly upload it and print status.
    If rubric_file is TXT, download the rubric template, fill it, and upload.
    Uses POST /api/v1/courses/:course_id/rubrics/upload.
    After import, fetches the status of each rubric import via GET /api/v1/courses/:course_id/rubrics/upload/:id.
    Returns a list of results for each rubric imported, including status details.
    Format matches the export function: uses the Canvas rubric template for CSV, and TXT for plain text.
    """
    results = []
    try:
        ext = os.path.splitext(rubric_file)[1].lower()
        if ext == ".csv":
            # Directly upload the CSV file
            upload_url = f"{api_url}/api/v1/courses/{course_id}/rubrics/upload"
            files = {'attachment': open(rubric_file, 'rb')}
            upload_headers = {"Authorization": f"Bearer {api_key}"}
            upload_resp = requests.post(upload_url, headers=upload_headers, files=files)
            files['attachment'].close()
            import_id = None
            status_detail = None
            if upload_resp.status_code in (200, 201):
                try:
                    resp_json = upload_resp.json()
                    import_id = resp_json.get("id") or resp_json.get("import_id")
                except Exception:
                    import_id = None
                if verbose:
                    print(f"[RubricImport] Imported rubric CSV to course {course_id}.")
                else:
                    print(f"Imported rubric CSV to course.")
            else:
                if verbose:
                    print(f"[RubricImport] Failed to import rubric CSV: {upload_resp.text}")
                else:
                    print(f"Failed to import rubric CSV.")
                results.append({"status": f"failed ({upload_resp.status_code})"})
                return results
            # After import, get status of import if import_id is available
            if import_id:
                status_url = f"{api_url}/api/v1/courses/{course_id}/rubrics/upload/{import_id}"
                status_headers = {"Authorization": f"Bearer {api_key}"}
                status_resp = requests.get(status_url, headers=status_headers)
                if status_resp.status_code == 200:
                    try:
                        status_json = status_resp.json()
                        status_detail = status_json
                        status_str = status_json.get("workflow_state", "unknown")
                        results.append({"status": "imported", "import_id": import_id, "detail": status_detail})
                        if verbose:
                            print(f"[RubricImport] Import status: {status_str}")
                        else:
                            print(f"Import status: {status_str}")
                    except Exception:
                        results.append({"status": "imported", "import_id": import_id, "detail": None})
                        if verbose:
                            print(f"[RubricImport] Imported rubric CSV, but could not parse status detail.")
                else:
                    results.append({"status": "imported", "import_id": import_id, "detail": None})
                    if verbose:
                        print(f"[RubricImport] Imported rubric CSV, but failed to fetch import status ({status_resp.status_code}).")
            else:
                results.append({"status": "imported", "import_id": None, "detail": None})
            return results

        # TXT format: download template, fill, and upload
        # Download the rubric CSV template
        template_url = f"{api_url}/api/v1/rubrics/upload_template"
        headers = {"Authorization": f"Bearer {api_key}"}
        resp = requests.get(template_url, headers=headers)
        if resp.status_code != 200:
            if verbose:
                print(f"[RubricImport] Failed to download rubric template: {resp.text}")
            else:
                print("Failed to download rubric template.")
            return []
        template_csv = resp.text

        # Parse rubric_file (TXT) and fill template rows
        rubrics_to_import = []
        with open(rubric_file, "r", encoding="utf-8") as f:
            lines = f.readlines()
        current_title = None
        criteria = []
        for line in lines:
            line = line.strip()
            if line.lower().startswith("rubric:"):
                if current_title and criteria:
                    rubrics_to_import.append({"title": current_title, "criteria": criteria})
                current_title = line[7:].strip()
                criteria = []
            elif "|" in line:
                parts = line.split("|")
                if len(parts) >= 2:
                    desc = parts[0].strip()
                    pts = float(parts[1].strip())
                    long_desc = parts[2].strip() if len(parts) > 2 else ""
                    criteria.append({
                        "description": desc,
                        "points": pts,
                        "long_description": long_desc
                    })
        if current_title and criteria:
            rubrics_to_import.append({"title": current_title, "criteria": criteria})

        # For each rubric, fill the template and upload
        for rubric in rubrics_to_import:
            title = rubric["title"]
            criteria = rubric["criteria"]
            # Fill template CSV rows
            template_reader = csv.reader(io.StringIO(template_csv))
            template_rows = list(template_reader)
            header = template_rows[0]
            rubric_rows = []
            for idx, crit in enumerate(criteria, 1):
                row = ["" for _ in header]
                for col_idx, col in enumerate(header):
                    col_lower = col.strip().lower()
                    if "rubric title" in col_lower:
                        row[col_idx] = title
                    elif "criterion description" in col_lower:
                        row[col_idx] = crit["description"]
                    elif "points" == col_lower:
                        row[col_idx] = crit["points"]
                    elif "long description" in col_lower:
                        row[col_idx] = crit.get("long_description", "")
                rubric_rows.append(row)
            # Write filled CSV to temp file
            with tempfile.NamedTemporaryFile(mode="w+", suffix=".csv", delete=False, encoding="utf-8") as tmpf:
                writer = csv.writer(tmpf)
                writer.writerow(header)
                for row in rubric_rows:
                    writer.writerow(row)
                tmpf.flush()
                temp_csv_path = tmpf.name

            # Upload rubric CSV to Canvas
            upload_url = f"{api_url}/api/v1/courses/{course_id}/rubrics/upload"
            files = {'attachment': open(temp_csv_path, 'rb')}
            upload_headers = {"Authorization": f"Bearer {api_key}"}
            upload_resp = requests.post(upload_url, headers=upload_headers, files=files)
            files['attachment'].close()
            os.remove(temp_csv_path)

            import_id = None
            status_detail = None
            if upload_resp.status_code in (200, 201):
                try:
                    resp_json = upload_resp.json()
                    import_id = resp_json.get("id") or resp_json.get("import_id")
                except Exception:
                    import_id = None
                if verbose:
                    print(f"[RubricImport] Imported rubric '{title}' to course {course_id}.")
                else:
                    print(f"Imported rubric '{title}' to course.")
            else:
                if verbose:
                    print(f"[RubricImport] Failed to import rubric '{title}': {upload_resp.text}")
                else:
                    print(f"Failed to import rubric '{title}'.")
                results.append({"title": title, "status": f"failed ({upload_resp.status_code})"})
                continue

            # After import, get status of import if import_id is available
            if import_id:
                status_url = f"{api_url}/api/v1/courses/{course_id}/rubrics/upload/{import_id}"
                status_headers = {"Authorization": f"Bearer {api_key}"}
                status_resp = requests.get(status_url, headers=status_headers)
                if status_resp.status_code == 200:
                    try:
                        status_json = status_resp.json()
                        status_detail = status_json
                        status_str = status_json.get("workflow_state", "unknown")
                        results.append({"title": title, "status": "imported", "import_id": import_id, "detail": status_detail})
                        if verbose:
                            print(f"[RubricImport] Import status for '{title}': {status_str}")
                        else:
                            print(f"Import status for '{title}': {status_str}")
                    except Exception:
                        results.append({"title": title, "status": "imported", "import_id": import_id, "detail": None})
                        if verbose:
                            print(f"[RubricImport] Imported rubric '{title}', but could not parse status detail.")
                else:
                    results.append({"title": title, "status": "imported", "import_id": import_id, "detail": None})
                    if verbose:
                        print(f"[RubricImport] Imported rubric '{title}', but failed to fetch import status ({status_resp.status_code}).")
            else:
                results.append({"title": title, "status": "imported", "import_id": None, "detail": None})

        return results
    except Exception as e:
        if verbose:
            print(f"[RubricImport] Error importing rubrics: {e}")
        else:
            print(f"Error importing rubrics: {e}")
        return []

def update_canvas_rubrics_for_assignments(
    assignment_ids=None,
    rubric_id=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    category=CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    verbose=False
):
    """
    Update the rubric for a range of assignments in a Canvas course.
    If assignment_ids is None, list all assignments in the category and prompt user to select.
    If rubric_id is None, list all rubrics in the course and prompt user to select.

    Uses Canvas API GET /api/v1/courses/:course_id/rubrics to list all possible rubrics.

    Args:
        assignment_ids (list or None): List of assignment IDs to update, or None to prompt user.
        rubric_id (str or int or None): The rubric ID to associate, or None to prompt user.
        api_url (str): Canvas API base URL.
        api_key (str): Canvas API key.
        course_id (str or int): Canvas course ID.
        category (str): Assignment group/category to filter.
        verbose (bool): Print more details.

    Returns:
        dict: {assignment_id: status} for each assignment updated.
    """

    def timeout_handler(signum, frame):
        print("\nTimeout: No response after 60 seconds. Quitting...")
        raise TimeoutError("User input timeout")

    def get_input_with_timeout(prompt, timeout=60, default=None):
        # Use signal.SIGALRM only if available (not on Windows)
        if hasattr(signal, "SIGALRM"):
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(timeout)
            try:
                result = input(prompt)
                signal.alarm(0)
                if not result and default is not None:
                    print(f"Using default: {default}")
                    return default
                return result
            except TimeoutError:
                signal.alarm(0)
                print("Quitting due to timeout.")
                return None
            except KeyboardInterrupt:
                signal.alarm(0)
                print("\nOperation cancelled by user.")
                return None
        else:
            # On platforms without SIGALRM (e.g., Windows), just use input without timeout
            try:
                result = input(prompt)
                if not result and default is not None:
                    print(f"Using default: {default}")
                    return default
                return result
            except KeyboardInterrupt:
                print("\nOperation cancelled by user.")
                return None

    try:
        canvas = Canvas(api_url, api_key)
        course = canvas.get_course(course_id)

        # If assignment_ids is None, list all assignments in category and prompt user
        if assignment_ids is None:
            assignments = []
            assignment_groups = list(course.get_assignment_groups(include=['assignments']))
            for group in assignment_groups:
                group_name = group.name
                if category and group_name.lower() != category.lower():
                    continue
                for assignment in group.assignments:
                    assignments.append({
                        "id": assignment['id'],
                        "name": assignment['name'],
                        "group": group_name,
                        "due_at": assignment.get('due_at')
                    })
            if not assignments:
                print("No assignments found in category.")
                return {}
            # Sort by due date
            def due_sort_key(a):
                raw = a.get("due_at")
                if raw:
                    try:
                        return datetime.strptime(raw, "%Y-%m-%dT%H:%M:%SZ")
                    except Exception:
                        return datetime.max
                return datetime.max
            assignments.sort(key=due_sort_key)
            print("Assignments in category:")
            for idx, a in enumerate(assignments, 1):
                due = a['due_at'] or "No due date"
                print(f"{idx}. [{a['group']}] {a['name']} (ID: {a['id']}, Due: {due})")
            sel = get_input_with_timeout(
                "Enter the number(s) of the assignment(s) to update (e.g. 1,3-5 or 'a' for all, or 'q' to quit): ",
                timeout=60,
                default="q"
            )
            if sel is None or sel.lower() in ('q', 'quit'):
                return {}
            if sel.lower() in ('a', 'all'):
                assignment_ids = [a['id'] for a in assignments]
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
                selected = [i for i in selected if 1 <= i <= len(assignments)]
                if not selected:
                    print("Invalid selection.")
                    return {}
                assignment_ids = [assignments[i - 1]['id'] for i in selected]

        # If rubric_id is None, list all rubrics using GET /api/v1/courses/:course_id/rubrics
        if rubric_id is None:
            url = f"{api_url}/api/v1/courses/{course_id}/rubrics"
            headers = {"Authorization": f"Bearer {api_key}"}
            resp = requests.get(url, headers=headers)
            rubrics = []
            if resp.status_code == 200:
                rubrics = resp.json()
                print(f"Found {len(rubrics)} rubrics in this course." if not verbose else f"[RubricUpdate] Found {len(rubrics)} rubrics in this course.")
            else:
                print("Failed to fetch rubrics from Canvas course.")
                return {}
            if not rubrics:
                print("No rubrics found in this course.")
                return {}
            print("Rubrics in this course:")
            for idx, r in enumerate(rubrics, 1):
                rid = r.get("id")
                title = r.get("title", "")
                print(f"{idx}. Rubric ID: {rid} | Title: {title}")
            sel = get_input_with_timeout(
                "Enter the number of the rubric to use (or 'q' to quit): ",
                timeout=60,
                default="q"
            )
            if sel is None or sel.lower() in ('q', 'quit'):
                return {}
            if sel.isdigit() and 1 <= int(sel) <= len(rubrics):
                rubric_id = rubrics[int(sel) - 1].get("id")
            else:
                print("Invalid selection.")
                return {}

        # Update rubric for each assignment
        results = {}
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        for aid in assignment_ids:
            url = f"{api_url}/api/v1/courses/{course_id}/assignments/{aid}"
            data = {
                "assignment": {
                    "rubric_settings": {
                        "id": rubric_id
                    },
                    "rubric_id": rubric_id
                }
            }
            try:
                resp = requests.put(url, headers=headers, json=data)
                if resp.status_code in (200, 201):
                    results[aid] = "updated"
                    if verbose:
                        print(f"[RubricUpdate] Assignment {aid}: rubric updated to {rubric_id}.")
                else:
                    results[aid] = f"failed ({resp.status_code})"
                    if verbose:
                        print(f"[RubricUpdate] Assignment {aid}: failed to update rubric ({resp.status_code}).")
            except Exception as e:
                results[aid] = f"error: {e}"
                if verbose:
                    print(f"[RubricUpdate] Assignment {aid}: error {e}")
        return results
    except Exception as e:
        print(f"Error updating rubrics: {e}")
        return {}

def list_and_download_canvas_grading_standards(
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    verbose=False
):
    """
    List all grading standards (grading schemes) for a Canvas course and allow user to select one to download as JSON to the current folder.

    Args:
        api_url (str): Canvas API base URL.
        api_key (str): Canvas API key.
        course_id (str or int): Canvas course ID.
        verbose (bool): Print more details.

    Returns:
        dict: The selected grading standard object, or None if cancelled.
    """

    headers = {"Authorization": f"Bearer {api_key}"}
    url = f"{api_url}/api/v1/courses/{course_id}/grading_standards"
    try:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        standards = resp.json()
        if not standards:
            print("No grading standards found for this course.")
            return None
        print("Grading standards available in this course:")
        for idx, std in enumerate(standards, 1):
            print(f"{idx}. {std.get('title', '')} (ID: {std.get('id')}, Context: {std.get('context_type')})")
            if verbose:
                print(f"   - Grading scheme: {std.get('grading_scheme', [])}")
        while True:
            sel = input("Enter the number of the grading standard to download (or 'q' to quit): ").strip()
            if sel.lower() in ("q", "quit"):
                print("Cancelled.")
                return None
            if sel.isdigit() and 1 <= int(sel) <= len(standards):
                idx = int(sel) - 1
                selected = standards[idx]
                std_id = selected.get("id")
                # Fetch full grading standard details
                detail_url = f"{api_url}/api/v1/courses/{course_id}/grading_standards/{std_id}"
                detail_resp = requests.get(detail_url, headers=headers)
                detail_resp.raise_for_status()
                detail = detail_resp.json()
                fname = f"grading_standard_{detail.get('id')}_{detail.get('title','').replace(' ','_')}.json"
                with open(fname, "w", encoding="utf-8") as f:
                    json.dump(detail, f, ensure_ascii=False, indent=2)
                print(f"Downloaded grading standard '{detail.get('title','')}' to {fname}")
                return detail
            else:
                print("Invalid selection. Please enter a valid number or 'q' to quit.")
    except Exception as e:
        print(f"Error listing or downloading grading standards: {e}")

def add_canvas_grading_scheme(
    grading_scheme_file,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    verbose=False
):
    """
    Add a grading scheme (grading standard) to a Canvas course from a JSON file.
    The JSON file should contain:
    {
        "title": "Scheme Name",
        "grading_scheme": [
            {"name": "A", "value": 0.9},
            {"name": "B", "value": 0.8},
            ...
        ]
    }
    Returns the created grading standard object or None on failure.
    """
    try:
        with open(grading_scheme_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        title = data.get("title")
        grading_scheme = data.get("grading_scheme")
        if not title or not grading_scheme:
            if verbose:
                print("[GradingScheme] JSON must contain 'title' and 'grading_scheme' fields.")
            else:
                print("JSON must contain 'title' and 'grading_scheme' fields.")
            return None
        url = f"{api_url}/api/v1/courses/{course_id}/grading_standards"
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        payload = {
            "title": title,
            "grading_scheme_entry": grading_scheme
        }
        resp = requests.post(url, headers=headers, json=payload)
        if resp.status_code in (200, 201):
            result = resp.json()
            if verbose:
                print(f"[GradingScheme] Grading scheme '{title}' added to course {course_id}.")
            else:
                print(f"Grading scheme '{title}' added to course.")
            return result
        else:
            if verbose:
                print(f"[GradingScheme] Failed to add grading scheme: {resp.status_code} {resp.text}")
            else:
                print(f"Failed to add grading scheme: {resp.status_code}")
            return None
    except Exception as e:
        if verbose:
            print(f"[GradingScheme] Error adding grading scheme: {e}")
        else:
            print(f"Error adding grading scheme: {e}")
        return None

def download_and_check_student_submissions(
    student_canvas_id=None,
    dest_dir=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    ocr_service=DEFAULT_OCR_METHOD,
    lang="auto",
    refine=DEFAULT_AI_METHOD,
    similarity_threshold=0.85,
    db_path=None,
    verbose=False
):
    """
    Download the latest submission file of a student for each assignment, extract text, check for similarity,
    and if two files are highly similar, send a warning message to the student.
    If student_canvas_id is None, list all students and allow user to choose a range of one or more students to check,
    allow user to select all students, allow user to quit, if no response after 60 seconds then use the default option of selecting all students.
    Only the latest submission (by submission time) for each assignment is considered.
    Downloaded file is named: <student name>_<canvas id>_<assignment id>_<submitted time>_<status>.<ext>
    Also saves the similarity check results in the database entry for the student if db_path is provided.
    Now saves downloaded file to a subfolder named <student name>_<student id> in the dest_dir (create if not exist).
    """

    def timeout_handler(signum, frame):
        print("\nTimeout: No response after 60 seconds. Using default: all students.")
        raise TimeoutError("User input timeout")

    def get_input_with_timeout(prompt, timeout=60, default=None):
        # Use signal.SIGALRM only if available (not on Windows)
        if hasattr(signal, "SIGALRM"):
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(timeout)
            try:
                result = input(prompt)
                signal.alarm(0)
                if not result and default is not None:
                    print(f"Using default: {default}")
                    return default
                return result
            except TimeoutError:
                signal.alarm(0)
                if default is not None:
                    print(f"Using default: {default}")
                    return default
                raise
            except KeyboardInterrupt:
                signal.alarm(0)
                print("\nOperation cancelled by user.")
                raise
        else:
            # On platforms without SIGALRM (e.g., Windows), just use input without timeout
            result = input(prompt)
            if not result and default is not None:
                print(f"Using default: {default}")
                return default
            return result

    # Load student database for Canvas ID -> Name mapping if available
    canvasid_to_name = {}
    canvasid_to_sid = {}
    if db_path and os.path.exists(db_path):
        try:
            students_db = load_database(db_path)
            for s in students_db:
                canvas_id = str(getattr(s, "Canvas ID", "")).strip()
                name = getattr(s, "Name", "")
                sid = str(getattr(s, "Student ID", "")).strip()
                if canvas_id and name:
                    canvasid_to_name[canvas_id] = name
                if canvas_id and sid:
                    canvasid_to_sid[canvas_id] = sid
        except Exception:
            canvasid_to_name = {}
            canvasid_to_sid = {}

    canvas = Canvas(api_url, api_key)
    course = canvas.get_course(course_id)

    # If student_canvas_id is None, list all students and allow user to select
    student_canvas_ids = []
    if student_canvas_id is None:
        people = list_canvas_people(api_url, api_key, course_id, verbose=verbose)
        students = people.get("active_students", [])
        if not students:
            print("No active students found.")
            return
        print("Active students:")
        for idx, s in enumerate(students, 1):
            print(f"{idx}. {s['name']} ({s['email']}), Canvas ID: {s['canvas_id']}")
        while True:
            sel = get_input_with_timeout(
                "Enter student numbers to check (e.g. 1-5,7,9 or 'a' for all, or 'q' to quit, default 'a' in 60s): ",
                timeout=60,
                default="a"
            ).strip()
            if sel.lower() in ("q", "quit"):
                print("Quitting.")
                return
            if sel.lower() in ("a", "all"):
                selected = list(range(1, len(students) + 1))
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
                selected = [i for i in selected if 1 <= i <= len(students)]
            if not selected:
                print("No valid selection. Try again or 'q' to quit.")
                continue
            student_canvas_ids = [students[i - 1]["canvas_id"] for i in selected]
            break
    else:
        student_canvas_ids = [student_canvas_id] if isinstance(student_canvas_id, (str, int)) else list(student_canvas_id)

    if not dest_dir:
        dest_dir = os.path.join(os.getcwd(), "student_submissions")
    os.makedirs(dest_dir, exist_ok=True)

    for student_canvas_id in student_canvas_ids:
        assignments = list(course.get_assignments())
        pdf_files = []
        file_info = []

        # Get student name and student id from database if possible, else from Canvas
        student_name = None
        student_sid = None
        canvas_id_str = str(student_canvas_id)
        if canvasid_to_name.get(canvas_id_str):
            student_name = canvasid_to_name[canvas_id_str]
        else:
            try:
                user = canvas.get_user(student_canvas_id)
                student_name = getattr(user, "name", f"student_{student_canvas_id}")
            except Exception:
                student_name = f"student_{student_canvas_id}"
        if canvasid_to_sid.get(canvas_id_str):
            student_sid = canvasid_to_sid[canvas_id_str]
        else:
            student_sid = "unknown"

        # Create subfolder for this student
        safe_student_name = re.sub(r"[^\w\s-]", "", student_name).strip().replace(" ", "_")
        safe_student_sid = re.sub(r"[^\w\s-]", "", student_sid).strip()
        student_subfolder = f"{safe_student_name}_{safe_student_sid}"
        student_dir = os.path.join(dest_dir, student_subfolder)
        os.makedirs(student_dir, exist_ok=True)

        # Download only the latest PDF submission for each assignment
        for assignment in tqdm(assignments, desc=f"Checking assignments for {student_name}"):
            try:
                sub = assignment.get_submission(student_canvas_id)
                # Find the latest submission attempt by submitted_at
                latest_attachment = None
                latest_time = None
                latest_status = None
                # Check submission_history if available
                if hasattr(sub, "submission_history") and sub.submission_history:
                    for h in sub.submission_history:
                        submitted_at = h.get("submitted_at")
                        workflow_state = h.get("workflow_state", "unknown")
                        if submitted_at:
                            try:
                                sub_time = datetime.strptime(submitted_at, "%Y-%m-%dT%H:%M:%SZ")
                            except Exception:
                                continue
                            attachments = h.get("attachments", [])
                            pdf_attachments = [att for att in attachments if getattr(att, "filename", "").lower().endswith(".pdf")]
                            if pdf_attachments:
                                if latest_time is None or sub_time > latest_time:
                                    latest_time = sub_time
                                    latest_attachment = pdf_attachments[0]
                                    latest_status = workflow_state
                # Fallback: check attachments directly on submission object
                elif hasattr(sub, "attachments") and sub.attachments:
                    attachments = sub.attachments
                    pdf_attachments = [att for att in attachments if getattr(att, "filename", "").lower().endswith(".pdf")]
                    if pdf_attachments:
                        submitted_at = getattr(sub, "submitted_at", None)
                        workflow_state = getattr(sub, "workflow_state", "unknown")
                        if submitted_at:
                            try:
                                sub_time = datetime.strptime(submitted_at, "%Y-%m-%dT%H:%M:%SZ")
                            except Exception:
                                sub_time = None
                        else:
                            sub_time = None
                        if latest_time is None or (sub_time and (latest_time is None or sub_time > latest_time)):
                            latest_time = sub_time
                            latest_attachment = pdf_attachments[0]
                            latest_status = workflow_state
                if not latest_attachment:
                    continue
                att = latest_attachment
                url = getattr(att, "url", None)
                orig_filename = getattr(att, "filename", None) or f"{assignment.id}_{student_canvas_id}.pdf"
                # Format: <student name>_<canvas id>_<assignment id>_<submitted time>_<status>.<ext>
                ext = os.path.splitext(orig_filename)[1]
                submitted_time_str = latest_time.strftime("%Y%m%d_%H%M") if latest_time else "unknown"
                status_str = latest_status or "unknown"
                out_filename = f"{safe_student_name}_{student_canvas_id}_{assignment.id}_{submitted_time_str}_{status_str}{ext}"
                out_path = os.path.join(student_dir, out_filename)
                if not os.path.exists(out_path):
                    r = requests.get(url)
                    with open(out_path, "wb") as f:
                        f.write(r.content)
                pdf_files.append(out_path)
                file_info.append({
                    "assignment_id": assignment.id,
                    "assignment_name": assignment.name,
                    "file_path": out_path
                })
            except Exception as e:
                if verbose:
                    print(f"[StudentSubmissions] Error downloading for assignment {assignment.id}: {e}")

        if not pdf_files:
            if verbose:
                print(f"[StudentSubmissions] No PDF submissions found for student {student_canvas_id}.")
            else:
                print("No PDF submissions found for this student.")
            continue

        # Extract text from all PDFs
        extracted_texts = {}
        for info in tqdm(file_info, desc=f"Extracting text for {student_name}"):
            pdf_path = info["file_path"]
            txt_path = pdf_path + f"_text_{ocr_service}.txt"
            if not os.path.exists(txt_path):
                txt_path = extract_text_from_scanned_pdf(
                    pdf_path,
                    txt_output_path=txt_path,
                    service=ocr_service,
                    lang=lang,
                    simple_text=True,
                    refine=refine,
                    verbose=verbose
                )
            if txt_path and os.path.exists(txt_path):
                with open(txt_path, "r", encoding="utf-8") as f:
                    text = f.read()
                norm_text = re.sub(r"\s+", " ", text).strip().lower()
                extracted_texts[pdf_path] = norm_text
            else:
                extracted_texts[pdf_path] = ""

        # Compare all pairs for similarity
        pdf_list = list(extracted_texts.keys())
        texts = [extracted_texts[p] for p in pdf_list]
        if len(pdf_list) < 2:
            if verbose:
                print(f"[StudentSubmissions] Less than 2 submissions to compare for student {student_canvas_id}.")
            else:
                print("Less than 2 submissions to compare.")
            continue

        tfidf_vectorizer = TfidfVectorizer().fit(texts)
        tfidf_matrix = tfidf_vectorizer.transform(texts)
        similar_pairs = []
        for i, pdf1 in enumerate(pdf_list):
            for j in range(i + 1, len(pdf_list)):
                pdf2 = pdf_list[j]
                text1 = extracted_texts[pdf1]
                text2 = extracted_texts[pdf2]
                if not text1 or not text2:
                    continue
                cos_sim = cosine_similarity(tfidf_matrix[i], tfidf_matrix[j])[0, 0]
                seq_sim = difflib.SequenceMatcher(None, text1, text2).ratio()
                ratio = 0.7 * cos_sim + 0.3 * seq_sim
                if ratio >= similarity_threshold:
                    similar_pairs.append((pdf1, pdf2, ratio))

        # Only save similarity results to database if high similarity found
        if similar_pairs and db_path and os.path.exists(db_path):
            try:
                students = load_database(db_path)
                # Find the student entry by Canvas ID
                for s in students:
                    if str(getattr(s, "Canvas ID", "")) == str(student_canvas_id):
                        # Save the similarity results as a field
                        s.__dict__["Submission Similarity Results"] = [
                            {
                                "file1": os.path.basename(pdf1),
                                "file2": os.path.basename(pdf2),
                                "similarity": ratio
                            }
                            for pdf1, pdf2, ratio in similar_pairs
                        ]
                        save_database(students, db_path)
                        if verbose:
                            print(f"[StudentSubmissions] Saved similarity results to database for student {student_canvas_id}.")
                        break
            except Exception as e:
                if verbose:
                    print(f"[StudentSubmissions] Failed to save similarity results to database: {e}")

        if similar_pairs:
            # Compose warning message
            assignment_names = []
            for pdf1, pdf2, _ in similar_pairs:
                a1 = next((f["assignment_name"] for f in file_info if f["file_path"] == pdf1), pdf1)
                a2 = next((f["assignment_name"] for f in file_info if f["file_path"] == pdf2), pdf2)
                assignment_names.append((a1, a2))
            assignment_names_str = "\n".join([f"- {a1} <-> {a2}" for a1, a2 in assignment_names])
            message = (
                "Hệ thống phát hiện bạn đã nộp các file có nội dung rất giống nhau cho nhiều bài tập khác nhau:\n"
                f"{assignment_names_str}\n\n"
                "Việc nộp các file có nội dung rất giống nhau cho nhiều bài tập khác nhau bị coi là gian lận và vi phạm quy định của lớp học. "
                "Bạn cần nộp lại từng bài tập với nội dung phù hợp cho từng bài càng sớm càng tốt. "
                "Nếu có thắc mắc, hãy liên hệ với giảng viên. \n"
                "Thông báo này được gửi tự động từ hệ thống."
            )

            # Print message and ask user if want to send/refine/quit
            print("\n--- Prepared warning message to send ---")
            print(f"Subject: {subject}")
            print(f"Message:\n{message}\n")
            print("Options: [y] Send as is  [r] Refine with AI  [q] Quit/skip")
            try:
                choice = get_input_with_timeout(
                    "Send this message? (y/r/q, default y in 60s): ",
                    timeout=60,
                    default="y"
                ).strip().lower()
            except TimeoutError:
                choice = "y"
            except KeyboardInterrupt:
                print("Operation cancelled by user.")
                return

            if choice in ("q", "quit"):
                if verbose:
                    print("[StudentSubmissions] Skipped sending warning message.")
                else:
                    print("Skipped sending warning message.")
                continue
            if choice == "r":
                # Refine with AI
                if refine in ALL_AI_METHODS:
                    prompt = (
                        "Bạn là trợ lý giáo dục chuyên nghiệp. Hãy viết lại thông báo sau bằng tiếng Việt, lịch sự, rõ ràng, "
                        "giải thích rằng việc nộp các file có nội dung rất giống nhau cho nhiều bài tập khác nhau bị coi là gian lận, "
                        "và nhắc sinh viên nộp lại từng bài tập với nội dung phù hợp. Đưa vào danh sách các bài bị phát hiện. "
                        "Chỉ trả về thông báo đã chỉnh sửa, không giải thích gì thêm.\n\n"
                        "Thông báo:\n{text}"
                    )
                    message = refine_text_with_ai(message, method=refine, user_prompt=prompt)
                    print("\n--- Refined message ---\n")
                    print(message)
                    try:
                        confirm = get_input_with_timeout(
                            "Send this refined message? (y/n, default y in 60s): ",
                            timeout=60,
                            default="y"
                        ).strip().lower()
                    except TimeoutError:
                        confirm = "y"
                    except KeyboardInterrupt:
                        print("Operation cancelled by user.")
                        return
                    if confirm not in ("y", "yes", ""):
                        if verbose:
                            print("[StudentSubmissions] Skipped sending warning message.")
                        else:
                            print("Skipped sending warning message.")
                        continue
                else:
                    print("No AI method available for refinement. Sending original message.")

            # Send message
            try:
                canvas.create_conversation(
                    recipients=[str(student_canvas_id)],
                    subject=subject,
                    body=message,
                    force_new=True
                )
                if verbose:
                    print(f"[StudentSubmissions] Warning message sent to student {student_canvas_id}.")
                else:
                    print("Warning message sent to student.")
            except Exception as e:
                if verbose:
                    print(f"[StudentSubmissions] Failed to send warning message: {e}")
                else:
                    print(f"Failed to send warning message: {e}")
        else:
            if verbose:
                print(f"[StudentSubmissions] No highly similar submissions detected for student {student_canvas_id}.")
            else:
                print("No highly similar submissions detected.")

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

    def get_input_with_timeout(prompt, timeout=60, default=None):
        # Use SIGALRM when available (not on Windows)
        if hasattr(signal, "SIGALRM"):
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(timeout)
            try:
                res = input(prompt)
                signal.alarm(0)
                if not res and default is not None:
                    return default
                return res
            except TimeoutError:
                signal.alarm(0)
                return None
            except KeyboardInterrupt:
                signal.alarm(0)
                raise
        else:
            try:
                res = input(prompt)
                if not res and default is not None:
                    return default
                return res
            except KeyboardInterrupt:
                raise

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
        canvas = Canvas(api_url, api_key)
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
        sel = get_input_with_timeout(
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
        confirmed = get_input_with_timeout(
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
        c = get_input_with_timeout(
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
            date_input = get_input_with_timeout(
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
        date_input = get_input_with_timeout(prompt, timeout=60, default=None)
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

    def get_input_with_timeout(prompt, timeout=60, default=None):
        # Use signal.SIGALRM only if available (not on Windows)
        if hasattr(signal, "SIGALRM"):
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(timeout)
            try:
                result = input(prompt)
                signal.alarm(0)
                if not result and default is not None:
                    if verbose:
                        print(f"[CreateGroups] Using default: {default}")
                    else:
                        print(f"Using default: {default}")
                    return default
                return result
            except TimeoutError:
                signal.alarm(0)
                if default is not None:
                    if verbose:
                        print(f"[CreateGroups] Using default: {default}")
                    else:
                        print(f"Using default: {default}")
                    return default
                raise
            except KeyboardInterrupt:
                signal.alarm(0)
                if verbose:
                    print("\n[CreateGroups] Operation cancelled by user.")
                else:
                    print("\nOperation cancelled by user.")
                raise
        else:
            # Fallback for platforms without SIGALRM (e.g., Windows)
            result = input(prompt)
            if not result and default is not None:
                if verbose:
                    print(f"[CreateGroups] Using default: {default}")
                else:
                    print(f"Using default: {default}")
                return default
            return result

    try:
        canvas = Canvas(api_url, api_key)
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
                name = get_input_with_timeout(
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
                sel = get_input_with_timeout(
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
        num_str = get_input_with_timeout(
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
            pattern = get_input_with_timeout(
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

    def get_input_with_timeout(prompt, timeout=60, default=None):
        # Use signal.SIGALRM only if available (not on Windows)
        if hasattr(signal, "SIGALRM"):
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(timeout)
            try:
                res = input(prompt)
                signal.alarm(0)
                if not res and default is not None:
                    print(f"Using default: {default}")
                    return default
                return res
            except TimeoutError:
                signal.alarm(0)
                if default is not None:
                    print(f"Using default: {default}")
                    return default
                raise
            except KeyboardInterrupt:
                signal.alarm(0)
                raise
        else:
            try:
                res = input(prompt)
                if not res and default is not None:
                    print(f"Using default: {default}")
                    return default
                return res
            except KeyboardInterrupt:
                raise

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
        canvas = Canvas(api_url, api_key)
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
        sel = get_input_with_timeout(
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
        confirmed = get_input_with_timeout(
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
        c = get_input_with_timeout(
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
            date_input = get_input_with_timeout(
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
        date_input = get_input_with_timeout(prompt, timeout=60, default=None)
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

    def get_input_with_timeout(prompt, timeout=60, default=None):
        # Use signal.SIGALRM only if available (not on Windows)
        if hasattr(signal, "SIGALRM"):
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(timeout)
            try:
                result = input(prompt)
                signal.alarm(0)
                if not result and default is not None:
                    if verbose:
                        print(f"[DeleteGroups] Using default: {default}")
                    else:
                        print(f"Using default: {default}")
                    return default
                return result
            except TimeoutError:
                signal.alarm(0)
                if default is not None:
                    if verbose:
                        print(f"[DeleteGroups] Using default: {default}")
                    else:
                        print(f"Using default: {default}")
                    return default
                raise
            except KeyboardInterrupt:
                signal.alarm(0)
                if verbose:
                    print("\n[DeleteGroups] Operation cancelled by user.")
                else:
                    print("\nOperation cancelled by user.")
                raise
        else:
            # Fallback for platforms without SIGALRM (e.g., Windows)
            result = input(prompt)
            if not result and default is not None:
                if verbose:
                    print(f"[DeleteGroups] Using default: {default}")
                else:
                    print(f"Using default: {default}")
                return default
            return result

    try:
        canvas = Canvas(api_url, api_key)
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
            sel = get_input_with_timeout(
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
        confirm = get_input_with_timeout(
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
