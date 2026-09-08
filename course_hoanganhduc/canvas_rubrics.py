# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas rubric helpers."""

import csv
import io
import os
import tempfile
from datetime import datetime
import signal

import requests

from .canvas_auth import get_canvas_client

from .canvas_utils import prompt_with_timeout, parse_selection, normalize_datetime_str
from .settings import (
    CANVAS_LMS_API_KEY,
    CANVAS_LMS_API_URL,
    CANVAS_LMS_COURSE_ID,
    CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
)


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
        canvas = get_canvas_client(api_url, api_key, verbose=verbose)
        course = canvas.get_course(course_id)
        rubric_ids = set()
        rubric_details = {}

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

        for rid in rubric_ids:
            rubric = fetch_rubric(rid)
            if rubric:
                rubric_details[rid] = rubric

        if not rubric_details:
            print("[Rubrics] No rubric details found." if not verbose else "[Rubrics] No rubric details found.")
            return []

        for rid, rubric in rubric_details.items():
            title = rubric.get("title", "")
            print(f"\nRubric ID: {rid} | Title: {title}")
            print("Format for assessmentRubricAssessmentsController#create:")
            print(f"POST /api/v1/courses/{course_id}/rubric_associations/{rid}/rubric_assessments")
            print("rubric_assessment[user_id]: <user_id>")
            print("rubric_assessment[assessment_type]: grading|peer_review|provisional_grade")
            criteria = rubric.get("data", [])
            for criterion in criteria:
                crit_id = criterion.get("id", "")
                desc = criterion.get("description", "")
                long_desc = criterion.get("long_description", "")
                points = criterion.get("points", "")
                print(f"  criterion_{crit_id}[points]: <points_awarded>  # {desc} ({points} pts)")
                print(f"  criterion_{crit_id}[comments]: <comments>      # {long_desc}")
                ratings = criterion.get("ratings", [])
                for rating in ratings or []:
                    print(f"     - Rating: {rating.get('description', '')}: {rating.get('points', '')} pts")

        if export_path:
            ext = os.path.splitext(export_path)[1].lower()
            if ext == ".csv":
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
                for rubric in rubric_details.values():
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
                        f.write("Format for assessmentRubricAssessmentsController#create:\n")
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
                    print("Imported rubric CSV to course.")
            else:
                if verbose:
                    print(f"[RubricImport] Failed to import rubric CSV: {upload_resp.text}")
                else:
                    print("Failed to import rubric CSV.")
                results.append({"status": f"failed ({upload_resp.status_code})"})
                return results
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
                            print("[RubricImport] Imported rubric CSV, but could not parse status detail.")
                else:
                    results.append({"status": "imported", "import_id": import_id, "detail": None})
                    if verbose:
                        print(f"[RubricImport] Imported rubric CSV, but failed to fetch import status ({status_resp.status_code}).")
            else:
                results.append({"status": "imported", "import_id": None, "detail": None})
            return results

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

        rubrics_to_import = []
        with open(rubric_file, "r", encoding="utf-8") as f:
            lines = f.readlines()
        current_title = None
        criteria = []
        for line in lines:
            line = line.strip()
            if not line:
                continue
            if line.lower().startswith("rubric id:"):
                continue
            if line.lower().startswith("rubric title:"):
                if current_title and criteria:
                    rubrics_to_import.append({"title": current_title, "criteria": criteria})
                current_title = line.split(":", 1)[1].strip()
                criteria = []
                continue
            if line.lower().startswith("criterion_"):
                parts = line.split(":", 1)
                if len(parts) != 2:
                    continue
                crit_key = parts[0].strip()
                crit_value = parts[1].strip()
                if crit_key.endswith("[points]"):
                    criteria.append({
                        "description": crit_key.split("[")[0].replace("criterion_", ""),
                        "points": crit_value,
                        "long_description": "",
                    })
                elif crit_key.endswith("[comments]") and criteria:
                    criteria[-1]["long_description"] = crit_value
        if current_title and criteria:
            rubrics_to_import.append({"title": current_title, "criteria": criteria})

        for rubric in rubrics_to_import:
            title = rubric["title"]
            criteria = rubric["criteria"]
            template_reader = csv.reader(io.StringIO(template_csv))
            template_rows = list(template_reader)
            header = template_rows[0]
            rubric_rows = []
            for crit in criteria:
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
            with tempfile.NamedTemporaryFile(mode="w+", suffix=".csv", delete=False, encoding="utf-8") as tmpf:
                writer = csv.writer(tmpf)
                writer.writerow(header)
                for row in rubric_rows:
                    writer.writerow(row)
                tmpf.flush()
                temp_csv_path = tmpf.name

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
    """

    def timeout_handler(signum, frame):
        print("\nTimeout: No response after 60 seconds. Quitting...")
        raise TimeoutError("User input timeout")

    def _prompt(prompt, timeout=60, default=None):
        return prompt_with_timeout(prompt, timeout=timeout, default=default, verbose=verbose)

    try:
        canvas = get_canvas_client(api_url, api_key, verbose=verbose)
        course = canvas.get_course(course_id)

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
            sel = _prompt(
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
            sel = _prompt(
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
