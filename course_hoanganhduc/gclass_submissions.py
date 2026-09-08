# -*- coding: utf-8 -*-

import os
import re
import io
import json
import unicodedata
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from .gclass_auth import _get_google_classroom_credentials, list_google_classroom_courses
from .settings import DEFAULT_OCR_METHOD, DEFAULT_AI_METHOD
from .submission_checks import compare_texts_from_pdfs_in_folder, analyze_meaningfulness_in_folder
from .utils import get_input_with_quit, parse_selection

def download_google_classroom_assignment_submissions(
    course_id=None,
    credentials_path='gclassroom_credentials.json',
    token_path='token.pickle',
    coursework_ids=None,
    dest_dir=None,
    ocr_service=DEFAULT_OCR_METHOD,
    lang="auto",
    meaningfulness_threshold=0.4,
    similarity_threshold=0.85,
    verbose=False,
):
    """
    Download latest Google Classroom submissions for selected coursework,
    then run meaningfulness + similarity checks on PDFs.
    """
    def safe_filename(value):
        text = str(value or "").strip()
        if not text:
            return "unknown"
        text = unicodedata.normalize('NFD', text)
        text = ''.join(c for c in text if not unicodedata.combining(c))
        text = re.sub(r"[^A-Za-z0-9._-]+", "_", text)
        return text.strip("_") or "unknown"

    def normalize_timestamp(value):
        if not value:
            return "unknown"
        text = str(value).replace(":", "").replace("-", "").replace("T", "_").replace("Z", "")
        text = re.sub(r"[^0-9_]+", "", text)
        return text or "unknown"

    def download_drive_file(drive_service, file_id, suggested_name, out_dir):
        try:
            meta = drive_service.files().get(fileId=file_id, fields="name,mimeType").execute()
        except Exception:
            meta = {}
        file_name = suggested_name or meta.get("name") or file_id
        mime_type = meta.get("mimeType")
        export_pdf = False
        if mime_type and mime_type.startswith("application/vnd.google-apps"):
            export_pdf = True
        base_name = os.path.splitext(file_name)[0] if file_name else file_id
        if export_pdf:
            file_name = base_name + ".pdf"
        safe_name = safe_filename(file_name)
        dest_path = os.path.join(out_dir, safe_name)
        if os.path.exists(dest_path):
            stem, ext = os.path.splitext(safe_name)
            dest_path = os.path.join(out_dir, f"{stem}_{file_id}{ext}")
        fh = io.BytesIO()
        if export_pdf:
            request = drive_service.files().export_media(fileId=file_id, mimeType="application/pdf")
        else:
            request = drive_service.files().get_media(fileId=file_id)
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        with open(dest_path, "wb") as f:
            f.write(fh.getvalue())
        return dest_path

    creds = _get_google_classroom_credentials(credentials_path, token_path, verbose=verbose)
    service = build("classroom", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    if not course_id:
        courses = list_google_classroom_courses(credentials_path, token_path, verbose=verbose)
        if not courses:
            print("No courses found.")
            return None
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
                    course_id = courses[idx].get("id")
                    break
            except Exception:
                continue
    if not course_id:
        print("No course selected.")
        return None

    coursework = []
    next_token = None
    while True:
        req = service.courses().courseWork().list(courseId=course_id, pageToken=next_token, pageSize=200) if next_token else service.courses().courseWork().list(courseId=course_id, pageSize=200)
        resp = req.execute()
        coursework.extend(resp.get("courseWork", []) or [])
        next_token = resp.get("nextPageToken")
        if not next_token:
            break
    if not coursework:
        print("No assignments found in this course.")
        return None

    selected_coursework = []
    if coursework_ids:
        ids = [str(cid).strip() for cid in (coursework_ids or []) if str(cid).strip()]
        for cw in coursework:
            if str(cw.get("id")) in ids:
                selected_coursework.append(cw)
        if not selected_coursework:
            print("No matching assignments found for provided coursework IDs.")
            return None
    else:
        print("Assignments:")
        for i, cw in enumerate(coursework, 1):
            title = cw.get("title", "")
            due_date = cw.get("dueDate")
            print(f"{i}. {title} (ID: {cw.get('id')}, due: {due_date})")
        while True:
            sel = get_input_with_quit("Select assignment numbers (e.g. 1,3-5, 'a' for all, or 'q' to quit): ")
            if sel is None:
                return None
            indices = parse_selection(sel, len(coursework))
            if indices:
                selected_coursework = [coursework[i - 1] for i in indices]
                break
            print("Invalid selection. Please enter valid numbers, a range, 'a' for all, or 'q' to quit.")

    students_map = {}
    try:
        next_token = None
        while True:
            req = service.courses().students().list(courseId=course_id, pageToken=next_token, pageSize=200) if next_token else service.courses().students().list(courseId=course_id, pageSize=200)
            resp = req.execute()
            for entry in resp.get("students", []) or []:
                profile = entry.get("profile", {}) or {}
                user_id = entry.get("userId")
                full_name = (profile.get("name", {}) or {}).get("fullName") or ""
                email = profile.get("emailAddress") or ""
                if user_id:
                    students_map[str(user_id)] = {"name": full_name, "email": email}
            next_token = resp.get("nextPageToken")
            if not next_token:
                break
    except Exception:
        if verbose:
            print("[GClassroom] Warning: could not fetch roster; submissions will show user IDs only.")

    base_dir = dest_dir or os.path.join(os.getcwd(), "gclassroom_submissions")
    os.makedirs(base_dir, exist_ok=True)

    summary = []
    for cw in selected_coursework:
        cw_id = cw.get("id")
        if not cw_id:
            continue
        title = cw.get("title", f"cw_{cw_id}")
        assignment_dir = os.path.join(base_dir, f"{safe_filename(title)}_{cw_id}")
        os.makedirs(assignment_dir, exist_ok=True)

        submissions = []
        next_token = None
        while True:
            req = service.courses().courseWork().studentSubmissions().list(courseId=course_id, courseWorkId=cw_id, pageToken=next_token, pageSize=200) if next_token else service.courses().courseWork().studentSubmissions().list(courseId=course_id, courseWorkId=cw_id, pageSize=200)
            resp = req.execute()
            submissions.extend(resp.get("studentSubmissions", []) or [])
            next_token = resp.get("nextPageToken")
            if not next_token:
                break

        submission_index = {}
        downloaded_files = []
        for sub in submissions:
            assignment_submission = sub.get("assignmentSubmission") or {}
            attachments = assignment_submission.get("attachments") or []
            if not attachments:
                continue
            user_id = str(sub.get("userId") or "")
            student_info = students_map.get(user_id, {})
            student_name = student_info.get("name") or user_id
            submitted_at = normalize_timestamp(sub.get("updateTime") or sub.get("creationTime") or "")
            for attach in attachments:
                drive_file = attach.get("driveFile")
                if not drive_file:
                    continue
                file_id = drive_file.get("id")
                file_title = drive_file.get("title") or file_id
                if not file_id:
                    continue
                prefix = f"{safe_filename(student_name)}_{safe_filename(user_id)}_{cw_id}_{submitted_at}"
                out_path = download_drive_file(drive_service, file_id, f"{prefix}_{file_title}", assignment_dir)
                filename = os.path.basename(out_path)
                submission_index[filename] = {
                    "user_id": user_id,
                    "name": student_name,
                    "email": student_info.get("email") or "",
                    "coursework_id": cw_id,
                    "submitted_at": submitted_at,
                    "file_id": file_id,
                }
                downloaded_files.append(filename)

        index_path = os.path.join(assignment_dir, "submission_index.json")
        with open(index_path, "w", encoding="utf-8") as f:
            json.dump(submission_index, f, ensure_ascii=False, indent=2)

        if not downloaded_files:
            print(f"No downloadable submissions found for '{title}'.")
            continue

        print(f"Downloaded {len(downloaded_files)} file(s) for '{title}' to {assignment_dir}")

        meaningful_results, low_quality, _, _ = analyze_meaningfulness_in_folder(
            assignment_dir,
            ocr_service=ocr_service,
            lang=lang,
            meaningfulness_threshold=meaningfulness_threshold,
            refine_method=DEFAULT_AI_METHOD,
            return_texts=False,
            write_report=True,
            verbose=verbose,
        )
        similarity_pairs = compare_texts_from_pdfs_in_folder(
            assignment_dir,
            ocr_service=ocr_service,
            lang=lang,
            refine=None,
            similarity_threshold=similarity_threshold,
            auto_send=False,
            notify_students=False,
            verbose=verbose,
        )

        notify_choice = input("Notify students about flagged submissions? (y/n) [n]: ").strip().lower()
        if notify_choice in ("y", "yes"):
            draft_path = os.path.join(assignment_dir, "gclassroom_notification_drafts.txt")
            with open(draft_path, "w", encoding="utf-8") as f:
                f.write("Google Classroom notification drafts (manual send required)\n")
                f.write(f"Assignment: {title} ({cw_id})\n\n")
                if low_quality:
                    f.write("Low quality submissions:\n")
                    for filename in low_quality:
                        meta = submission_index.get(filename, {})
                        f.write(f"- {filename} | {meta.get('name')} | {meta.get('email')}\n")
                    f.write("\n")
                if similarity_pairs:
                    f.write("Similarity pairs:\n")
                    for pdf1, pdf2, ratio in similarity_pairs:
                        m1 = submission_index.get(pdf1, {})
                        m2 = submission_index.get(pdf2, {})
                        f.write(f"- {pdf1} ({m1.get('name')}) <-> {pdf2} ({m2.get('name')}): {ratio:.2f}\n")
                f.write("\nNote: Google Classroom API does not support direct messaging.\n")
            print(f"Draft notifications saved to {draft_path}")
        else:
            print("Skipping notifications.")

        summary.append({
            "coursework_id": cw_id,
            "title": title,
            "downloaded": len(downloaded_files),
            "low_quality": len(low_quality),
            "similarity_pairs": len(similarity_pairs or []),
            "folder": assignment_dir,
        })

    return summary
