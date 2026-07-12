# -*- coding: utf-8 -*-
"""Agent-safe Google Classroom entrypoint (course-google-classroom skill).

Does not expose unenroll, grade, or submission download.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from typing import Any, List, Optional

from .course_agent_common import (
    CourseAgentError,
    force_agent_mode,
    refuse,
    require_env_allowlist,
)


def main(argv: Optional[List[str]] = None) -> int:
    force_agent_mode()
    parser = argparse.ArgumentParser(
        prog="python -m course_hoanganhduc.gclass_agent",
        description="Agent-safe Google Classroom operations",
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("preflight", help="Check credentials/token paths exist")
    p_lc = sub.add_parser("list-courses", help="List Google Classroom courses")
    p_lc.add_argument("--credentials", default=None)
    p_lc.add_argument("--token", default=None)
    p_ls = sub.add_parser("list-students", help="List students for a course")
    p_ls.add_argument("--course-id", required=True)
    p_ls.add_argument("--credentials", default=None)
    p_ls.add_argument("--token", default=None)
    p_sync = sub.add_parser("sync", help="Sync GC students into local DB")
    p_sync.add_argument("--course-id", required=True)
    p_sync.add_argument("--db", default="students.db")
    p_sync.add_argument("--credentials", default=None)
    p_sync.add_argument("--token", default=None)
    for bad in ("unenroll", "grade", "download"):
        sub.add_parser(bad, help="(refused) not agent-safe")

    args = parser.parse_args(argv)

    try:
        if args.cmd in {"unenroll", "grade", "download"}:
            refuse(args.cmd)

        cred = (
            getattr(args, "credentials", None)
            or os.environ.get("GOOGLE_CLASSROOM_CREDENTIALS")
            or "gclassroom_credentials.json"
        )
        token = (
            getattr(args, "token", None)
            or os.environ.get("GOOGLE_CLASSROOM_TOKEN")
            or "token.pickle"
        )

        if args.cmd == "preflight":
            print(
                json.dumps(
                    {
                        "ok": os.path.exists(cred) or os.path.exists(token),
                        "credentials_path": cred,
                        "credentials_exists": os.path.exists(cred),
                        "token_path": token,
                        "token_exists": os.path.exists(token),
                    }
                )
            )
            return 0 if (os.path.exists(cred) or os.path.exists(token)) else 1

        if getattr(args, "course_id", None):
            require_env_allowlist(
                "GCLASS_COURSE_ALLOWLIST",
                str(args.course_id),
                label="google classroom course id",
            )

        if args.cmd == "list-courses":
            from .gclass_auth import list_google_classroom_courses

            courses = list_google_classroom_courses(cred, token, verbose=False)
            print(json.dumps(courses, indent=2, default=str, ensure_ascii=False))
            return 0

        if args.cmd == "list-students":
            from .gclass_auth import list_google_classroom_students

            students = list_google_classroom_students(
                cred, token, course_id=args.course_id, verbose=False
            )
            print(json.dumps(students, indent=2, default=str, ensure_ascii=False))
            return 0

        if args.cmd == "sync":
            from .data import load_database, save_database
            from .gclass_sync import sync_students_with_google_classroom

            students: List[Any] = []
            if os.path.exists(args.db):
                students = load_database(args.db, verbose=False) or []
            added, updated = sync_students_with_google_classroom(
                students,
                db_path=args.db,
                course_id=args.course_id,
                credentials_path=cred,
                token_path=token,
                fetch_grades=False,
                verbose=False,
            )
            save_database(
                students, args.db, verbose=False, audit_source="gclass_agent_sync"
            )
            print(json.dumps({"added": added, "updated": updated}))
            return 0

    except CourseAgentError as exc:
        print(f"gclass_agent error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # pragma: no cover
        print(f"gclass_agent error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
