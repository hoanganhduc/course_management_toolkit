# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas weekly automation helpers."""

import json
import os
import re
import shutil
from datetime import datetime, timezone

from .canvas_auth import get_canvas_client

from .canvas_checks import (
    compare_texts_from_pdfs_in_folder,
    detect_meaningful_level_and_notify_students,
    extract_canvas_id_from_filename,
)
from .canvas_people import list_canvas_people
from .canvas_submissions import download_canvas_assignment_submissions_auto
from .data import refine_text_with_ai
from .settings import (
    ALL_AI_METHODS,
    CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
    CANVAS_LMS_API_KEY,
    CANVAS_LMS_API_URL,
    CANVAS_LMS_COURSE_ID,
    DEFAULT_OCR_METHOD,
)

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
    canvas = get_canvas_client(api_url, api_key)
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

    canvas = get_canvas_client(api_url, api_key)
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
    canvas = get_canvas_client(api_url, api_key)
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

