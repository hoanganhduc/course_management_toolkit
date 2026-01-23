# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas grading scheme helpers."""

import json
import requests

from .settings import CANVAS_LMS_API_KEY, CANVAS_LMS_API_URL, CANVAS_LMS_COURSE_ID


def list_and_download_canvas_grading_standards(
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    verbose=False
):
    """
    List all grading standards (grading schemes) for a Canvas course and allow user to select one to download as JSON to the current folder.
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
        payload = {
            "grading_standard": {
                "title": title,
                "grading_scheme": grading_scheme
            }
        }
        headers = {"Authorization": f"Bearer {api_key}"}
        url = f"{api_url}/api/v1/courses/{course_id}/grading_standards"
        resp = requests.post(url, headers=headers, json=payload)
        if resp.status_code in (200, 201):
            if verbose:
                print(f"[GradingScheme] Added grading scheme '{title}' to course {course_id}.")
            else:
                print(f"Added grading scheme '{title}'.")
            return resp.json()
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
