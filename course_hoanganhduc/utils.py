# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/
# Course Management Script

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
import logging
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
from logging.handlers import RotatingFileHandler
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from types import SimpleNamespace
from datetime import datetime

from tqdm import tqdm  # <-- Add tqdm for progress bars

from .version import __version__

from .settings import *

def complete_path(text, state):
    """
    Tab-completion for file and directory paths, mimicking real terminal behavior.
    - Expands ~ and $HOME.
    - Handles quoted paths.
    - Supports directories (adds trailing slash).
    - Ignores hidden files unless prefix starts with '.'.
    - Handles spaces in filenames.
    - Supports file name completion and selection.
    """

    # Get the current line buffer and cursor position
    line = readline.get_line_buffer()
    if not line:
        line = ''
    else:
        line = line.strip()

    # Remove quotes for completion, but remember if quoted
    quote = ''
    if line and line[0] in ('"', "'"):
        quote = line[0]
        line = line[1:]
        if line and line[-1] == quote:
            line = line[:-1]

    # Expand ~ and $VARS
    expanded_line = os.path.expandvars(os.path.expanduser(line))
    # If the line ends with a space, show all files in that dir
    if expanded_line.endswith(os.sep):
        dirname = expanded_line
        prefix = ''
    else:
        dirname = os.path.dirname(expanded_line) or '.'
        prefix = os.path.basename(expanded_line)

    try:
        files = os.listdir(dirname)
    except Exception:
        files = []

    # Show hidden files only if prefix starts with '.'
    show_hidden = prefix.startswith('.')
    matches = []
    for f in files:
        if not show_hidden and f.startswith('.'):
            continue
        if f.startswith(prefix):
            full_path = os.path.join(dirname, f)
            # Add trailing slash for directories
            if os.path.isdir(full_path):
                f_out = f + os.sep
            else:
                f_out = f
            # Quote if needed
            if ' ' in f_out or any(c in f_out for c in ('"', "'")):
                if not quote:
                    f_out = '"' + f_out.replace('"', '\\"') + '"'
                else:
                    f_out = quote + f_out.replace(quote, '\\'+quote) + quote
            matches.append(os.path.join(dirname, f_out) if dirname not in ('.', '') else f_out)

    # Also support glob-style completion (e.g., *.pdf)
    if '*' in prefix or '?' in prefix or '[' in prefix:
        pattern = os.path.join(dirname, prefix)
        globbed = glob.glob(pattern)
        for g in globbed:
            g_base = os.path.basename(g)
            if not show_hidden and g_base.startswith('.'):
                continue
            if os.path.isdir(g):
                g_out = g_base + os.sep
            else:
                g_out = g_base
            if ' ' in g_out or any(c in g_out for c in ('"', "'")):
                if not quote:
                    g_out = '"' + g_out.replace('"', '\\"') + '"'
                else:
                    g_out = quote + g_out.replace(quote, '\\'+quote) + quote
            matches.append(os.path.join(dirname, g_out) if dirname not in ('.', '') else g_out)

    matches = list(sorted(set(matches)))
    try:
        return matches[state]
    except IndexError:
        return None

def input_with_completion(prompt, select_file=False, file_filter=None, verbose=False):
    """
    Input with tab-completion for file paths.
    If select_file=True, shows a numbered list of files in the directory for selection.
    file_filter: optional function that takes a filename and returns True/False.
    If verbose is True, print more details; otherwise, print only important notice.
    """
    if readline:
        readline.set_completer_delims(' \t\n;')
        readline.parse_and_bind("tab: complete")
        readline.set_completer(complete_path)
    try:
        user_input = input(prompt)
        user_input = user_input.strip()
        # Expand ~ and $HOME in the returned path
        expanded = os.path.expandvars(os.path.expanduser(user_input))
        if select_file:
            # If a directory, show files for selection
            path = expanded
            if not os.path.isdir(path):
                path = os.path.dirname(path) or '.'
            try:
                files = os.listdir(path)
                if file_filter:
                    files = [f for f in files if file_filter(f)]
                files = sorted(files)
                if not files:
                    if verbose:
                        print("[input_with_completion] No files found in directory:", path)
                    else:
                        print("No files found.")
                    return ""
                if verbose:
                    print(f"[input_with_completion] Files in '{path}':")
                print("Select a file:")
                for idx, f in enumerate(files, 1):
                    print(f"{idx}. {f}")
                while True:
                    sel = input("Enter file number or 0 to cancel: ").strip()
                    if sel == "0":
                        if verbose:
                            print("[input_with_completion] File selection cancelled by user.")
                        return ""
                    if sel.isdigit() and 1 <= int(sel) <= len(files):
                        selected_file = os.path.join(path, files[int(sel)-1])
                        if verbose:
                            print(f"[input_with_completion] Selected file: {selected_file}")
                        return selected_file
                    print("Invalid selection.")
            except Exception as e:
                if verbose:
                    print(f"[input_with_completion] Error listing files in '{path}': {e}")
                else:
                    print(f"Error listing files: {e}")
                return ""
        if verbose:
            print(f"[input_with_completion] Expanded input: {expanded}")
        return expanded
    finally:
        if readline:
            readline.set_completer(None)

def timeout_handler(signum, frame):
    """Handle timeout for user input operations."""
    print("\nTimeout: No response after 60 seconds. Quitting...")
    raise TimeoutError("User input timeout")

def get_input_with_timeout(prompt, timeout=60):
    """Get input with a timeout. Raises TimeoutError if no response."""
    # signal.SIGALRM is not available on some platforms (notably Windows).
    # Use it only when present; otherwise fall back to a blocking input() without timeout.
    if hasattr(signal, "SIGALRM"):
        signal.signal(signal.SIGALRM, timeout_handler)
        signal.alarm(timeout)
        try:
            result = input(prompt)
            signal.alarm(0)  # Cancel the alarm
            return result
        except KeyboardInterrupt:
            signal.alarm(0)
            print("\nOperation cancelled by user.")
            raise
        except TimeoutError:
            signal.alarm(0)
            raise
    else:
        # Fallback for platforms without SIGALRM: blocking input without timeout
        try:
            return input(prompt)
        except KeyboardInterrupt:
            print("\nOperation cancelled by user.")
            raise

def prefill_input_with_timeout(prompt, text, timeout=60):
    """Get input with pre-filled text and timeout."""
    def hook():
        readline.insert_text(str(text))
        readline.redisplay()
    readline.set_pre_input_hook(hook)
    try:
        signal.signal(signal.SIGALRM, timeout_handler)
        signal.alarm(timeout)
        result = input(prompt)
        signal.alarm(0)
        return result
    except (KeyboardInterrupt, TimeoutError):
        signal.alarm(0)
        raise
    finally:
        readline.set_pre_input_hook()

def format_time(unformatted_time):
    if not unformatted_time:
        return ""
    try:
        # Canvas returns ISO 8601 UTC time, e.g., "2024-07-15T16:59:00Z"
        dt = datetime.strptime(unformatted_time, "%Y-%m-%dT%H:%M:%SZ")
        # Convert to Vietnam time (GMT+7)
        vn_tz = timezone(timedelta(hours=7))
        dt_vn = dt.replace(tzinfo=timezone.utc).astimezone(vn_tz)
        # Format as "dd/mm/yyyy HH:MM (GMT+7)"
        return dt_vn.strftime("%d/%m/%Y %H:%M (GMT+7)")
    except Exception:
        return unformatted_time

def multiline_input(prompt):
    print(prompt + " (Enter an empty line to finish):")
    lines = []
    while True:
        line = input()
        if line == "":
            break
        lines.append(line)
    return "\n".join(lines)


def setup_logging(log_dir=None, log_level="INFO", max_bytes=5_000_000, backup_count=3, verbose=False):
    """
    Configure rotating file logging for the CLI. Returns the logger instance.
    """
    logger = logging.getLogger("course")
    if any(isinstance(handler, RotatingFileHandler) for handler in logger.handlers):
        return logger
    level = getattr(logging, str(log_level).upper(), logging.INFO)
    logger.setLevel(level)
    log_dir = log_dir or os.getcwd()
    try:
        os.makedirs(log_dir, exist_ok=True)
    except Exception as e:
        if verbose:
            print(f"[Logging] Failed to create log dir {log_dir}: {e}")
        log_dir = os.getcwd()
    log_path = os.path.join(log_dir, "course.log")
    handler = RotatingFileHandler(log_path, maxBytes=max_bytes, backupCount=backup_count, encoding="utf-8")
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    if verbose:
        print(f"[Logging] Writing logs to {log_path}")
    return logger


def append_run_report(action, details=None, outputs=None, verbose=False):
    """
    Append a one-line summary of a completed action to run_report.txt.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    outputs_list = []
    if outputs:
        if isinstance(outputs, (list, tuple, set)):
            outputs_list = [str(x) for x in outputs if x]
        else:
            outputs_list = [str(outputs)]
    line = f"{timestamp} | {action}"
    if details:
        line += f" | {details}"
    if outputs_list:
        line += f" | outputs: {', '.join(outputs_list)}"
    if DRY_RUN:
        if verbose:
            print(f"[RunReport] Dry run: would append: {line}")
        return
    report_path = os.path.join(os.getcwd(), "run_report.txt")
    try:
        with open(report_path, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception as e:
        if verbose:
            print(f"[RunReport] Failed to write report: {e}")


def generate_weekly_github_workflow(
    output_path=".github/workflows/weekly-course-tasks.yml",
    toolkit_repo_url="https://github.com/hoanganhduc/course_management_toolkit.git",
    toolkit_repo_branch="main",
    assignment_id="ASSIGNMENT_ID",
    course_code="MAT3500",
    course_id="COURSE_ID",
    teacher_canvas_id="TEACHER_CANVAS_ID",
    category="",
):
    """
    Generate a sample GitHub Actions workflow YAML for weekly automation.
    """
    category_line = f"          \"CANVAS_DEFAULT_ASSIGNMENT_CATEGORY\": \"{category}\"," if category else ""
    content = f"""name: Weekly Course Automation

on:
  schedule:
    - cron: "59 23 * * 0"
  workflow_dispatch:

jobs:
  weekly:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout course repo
        uses: actions/checkout@v4
      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.10"
      - name: Clone course toolkit
        run: |
          git clone --branch {toolkit_repo_branch} {toolkit_repo_url} course_toolkit
      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          pip install -r course_toolkit/requirements.txt
          pip install -e course_toolkit
      - name: Write config
        run: |
          cat > weekly_config.json <<'EOF'
          {{
            "CANVAS_LMS_API_URL": "${{ secrets.CANVAS_API_URL }}",
            "CANVAS_LMS_API_KEY": "${{ secrets.CANVAS_API_KEY }}",
            "CANVAS_LMS_COURSE_ID": "{course_id}",
            "DEFAULT_OCR_METHOD": "ocrspace",
            "OCRSPACE_API_KEY": "${{ secrets.OCRSPACE_API_KEY }}"
            {category_line}
          }}
          EOF
      - name: Run weekly checks
        env:
          COURSE_CODE: "{course_code}"
        run: |
          course --config weekly_config.json --course-code "{course_code}"
          course --run-weekly-automation \\
            --weekly-assignment-id "{assignment_id}" \\
            --weekly-teacher-canvas-id "{teacher_canvas_id}" \\
            --db students.db
      - name: Archive reports and artifacts
        run: |
          RUN_TS=$(date +"%Y%m%d %H:%M %Z")
          ARCHIVE_DIR="weekly_reports/$RUN_TS"
          mkdir -p "$ARCHIVE_DIR"
          cp students.db "$ARCHIVE_DIR/students.db.bak"
          for f in run_report.txt data_validation_report.txt grade_diff.csv weekly_automation_summary.json; do
            if [ -f "$f" ]; then mv "$f" "$ARCHIVE_DIR/"; fi
          done
          for d in final_evaluations student_submissions; do
            if [ -d "$d" ]; then
              base=$(basename "$d")
              mv "$d" "$ARCHIVE_DIR/$base"
            fi
          done
          for d in flagged_submissions_*; do
            if [ -d "$d" ]; then
              base=$(basename "$d")
              mv "$d" "$ARCHIVE_DIR/$base"
            fi
          done
          patterns=("pdf_similarity_results.txt" "pdf_similarity_status.json" "pdf_similarity_report.json" "meaningfulness_analysis.txt" "meaningfulness_status.json")
          for pattern in "${patterns[@]}"; do
            find . -path "./students_repo/weekly_reports" -prune -o -name "$pattern" -print0 | while IFS= read -r -d '' file; do
              rel="${file#./}"
              dest="$ARCHIVE_DIR/$(dirname "$rel")"
              mkdir -p "$dest"
              mv "$file" "$dest/"
            done
          done
          rm -f weekly_config.json
          rm -rf course_toolkit
        shell: bash
      - name: Commit updated students.db
        run: |
          if [ -n "$(git status --porcelain)" ]; then
            git add students.db weekly_reports
            git commit -m "Weekly automation: update students.db ($RUN_TS)"
            git push
          else
            echo "No changes to commit."
          fi
        shell: bash
"""
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)
    return output_path


def archive_weekly_artifacts_local(report_root="weekly_reports", students_db_path="students.db", base_dir=None, verbose=False):
    """
    Move weekly automation artifacts into a timestamped folder and back up students.db.
    """
    base_dir = base_dir or os.getcwd()
    timestamp = time.strftime("%Y%m%d %H:%M %Z", time.localtime())
    report_dir = os.path.join(base_dir, report_root, timestamp)
    os.makedirs(report_dir, exist_ok=True)

    db_path = os.path.join(base_dir, students_db_path)
    if os.path.isfile(db_path):
        try:
            shutil.copy2(db_path, os.path.join(report_dir, "students.db.bak"))
        except Exception as e:
            if verbose:
                print(f"[WeeklyLocal] Failed to back up students.db: {e}")

    files_to_move = [
        "run_report.txt",
        "data_validation_report.txt",
        "grade_diff.csv",
        "weekly_automation_summary.json",
    ]
    for filename in files_to_move:
        src = os.path.join(base_dir, filename)
        if os.path.isfile(src):
            try:
                shutil.move(src, os.path.join(report_dir, filename))
            except Exception as e:
                if verbose:
                    print(f"[WeeklyLocal] Failed to move {filename}: {e}")

    dirs_to_move = ["final_evaluations", "student_submissions"]
    for dirname in dirs_to_move:
        src = os.path.join(base_dir, dirname)
        if os.path.isdir(src):
            dest = os.path.join(report_dir, dirname)
            try:
                shutil.move(src, dest)
            except Exception as e:
                if verbose:
                    print(f"[WeeklyLocal] Failed to move {dirname}: {e}")

    for entry in os.listdir(base_dir):
        src = os.path.join(base_dir, entry)
        if entry.startswith("flagged_submissions_") and os.path.isdir(src):
            dest = os.path.join(report_dir, entry)
            try:
                shutil.move(src, dest)
            except Exception as e:
                if verbose:
                    print(f"[WeeklyLocal] Failed to move {entry}: {e}")

    patterns = {
        "pdf_similarity_results.txt",
        "pdf_similarity_status.json",
        "pdf_similarity_report.json",
        "meaningfulness_analysis.txt",
        "meaningfulness_status.json",
    }
    report_root_path = os.path.join(base_dir, report_root)
    for root, dirs, files in os.walk(base_dir):
        if root.startswith(report_root_path):
            continue
        for filename in files:
            if filename not in patterns:
                continue
            src = os.path.join(root, filename)
            rel = os.path.relpath(root, base_dir)
            dest_dir = os.path.join(report_dir, rel)
            os.makedirs(dest_dir, exist_ok=True)
            try:
                shutil.move(src, os.path.join(dest_dir, filename))
            except Exception as e:
                if verbose:
                    print(f"[WeeklyLocal] Failed to move {src}: {e}")

    return report_dir


def load_weekly_automation_history(report_root="weekly_reports", base_dir=None, verbose=False):
    """
    Load weekly automation history from archived summary files.
    """
    base_dir = base_dir or os.getcwd()
    report_root_path = os.path.join(base_dir, report_root)
    history = {}
    if not os.path.isdir(report_root_path):
        return history

    for root, _, files in os.walk(report_root_path):
        if "weekly_automation_summary.json" not in files:
            continue
        summary_path = os.path.join(root, "weekly_automation_summary.json")
        try:
            with open(summary_path, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except Exception as e:
            if verbose:
                print(f"[WeeklyLocal] Failed to read {summary_path}: {e}")
            continue
        assignment_id = str(payload.get("assignment_id", "")).strip()
        if not assignment_id:
            continue
        entry = {
            "assignment_id": assignment_id,
            "assignment_name": payload.get("assignment_name", ""),
            "generated_at": payload.get("generated_at", ""),
            "path": summary_path,
        }
        existing = history.get(assignment_id)
        if not existing or entry["generated_at"] > existing.get("generated_at", ""):
            history[assignment_id] = entry

    return history
