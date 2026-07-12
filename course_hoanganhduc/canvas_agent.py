# -*- coding: utf-8 -*-
"""Agent-safe Canvas entrypoint (ADR split skill course-canvas).

Does not expose unenroll, grade, invite, announce, messages, or bulk download.
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


def _course_id(args_course: Optional[str]) -> Optional[str]:
    return args_course or os.environ.get("CANVAS_LMS_COURSE_ID") or None


def main(argv: Optional[List[str]] = None) -> int:
    force_agent_mode()
    parser = argparse.ArgumentParser(
        prog="python -m course_hoanganhduc.canvas_agent",
        description="Agent-safe Canvas operations via course_hoanganhduc",
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("preflight", help="Check Canvas config presence (no secrets printed)")
    p_la = sub.add_parser("list-assignments", help="List Canvas assignments")
    p_la.add_argument("--course-id", default=None)
    p_la.add_argument("--category", default=None)
    p_lm = sub.add_parser("list-members", help="List Canvas course members")
    p_lm.add_argument("--course-id", default=None)
    p_su = sub.add_parser("search-user", help="Search Canvas user by name/email")
    p_su.add_argument("query")
    p_su.add_argument("--course-id", default=None)
    p_sync = sub.add_parser("sync", help="Sync Canvas members into local DB")
    p_sync.add_argument("--course-id", default=None)
    p_sync.add_argument("--db", default="students.db")
    # forbidden stubs for clear errors
    for bad in (
        "unenroll",
        "grade",
        "invite",
        "announce",
        "download",
        "messages",
        "pages",
    ):
        sub.add_parser(bad, help="(refused) not agent-safe")

    args = parser.parse_args(argv)

    try:
        if args.cmd in {
            "unenroll",
            "grade",
            "invite",
            "announce",
            "download",
            "messages",
            "pages",
        }:
            refuse(args.cmd)

        if args.cmd == "preflight":
            from . import settings

            url = getattr(settings, "CANVAS_LMS_API_URL", "") or ""
            key = getattr(settings, "CANVAS_LMS_API_KEY", "") or ""
            cid = getattr(settings, "CANVAS_LMS_COURSE_ID", "") or ""
            print(
                json.dumps(
                    {
                        "ok": bool(url and key),
                        "api_url_set": bool(url),
                        "api_key_set": bool(key),
                        "course_id_set": bool(cid),
                        # never print secrets
                    }
                )
            )
            return 0 if (url and key) else 1

        cid = _course_id(getattr(args, "course_id", None))
        if cid:
            require_env_allowlist(
                "CANVAS_COURSE_ALLOWLIST",
                str(cid),
                label="canvas course id",
            )

        if args.cmd == "list-assignments":
            from .canvas_assignments import list_canvas_assignments

            rows = list_canvas_assignments(
                course_id=cid,
                category=getattr(args, "category", None),
            )
            print(json.dumps(rows, indent=2, default=str, ensure_ascii=False))
            return 0

        if args.cmd == "list-members":
            from .canvas_people import list_canvas_people

            people = list_canvas_people(course_id=cid)
            print(json.dumps(people, indent=2, default=str, ensure_ascii=False))
            return 0

        if args.cmd == "search-user":
            from .canvas_people import search_canvas_user

            hits = search_canvas_user(args.query, course_id=cid)
            print(json.dumps(hits, indent=2, default=str, ensure_ascii=False))
            return 0

        if args.cmd == "sync":
            from .canvas_sync import sync_students_with_canvas
            from .data import load_database, save_database

            students: List[Any] = []
            if os.path.exists(args.db):
                students = load_database(args.db, verbose=False) or []
            added, updated = sync_students_with_canvas(
                students,
                db_path=args.db,
                course_id=cid,
                verbose=False,
            )
            save_database(
                students, args.db, verbose=False, audit_source="canvas_agent_sync"
            )
            print(json.dumps({"added": added, "updated": updated}))
            return 0

    except CourseAgentError as exc:
        print(f"canvas_agent error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # pragma: no cover - live API errors
        print(f"canvas_agent error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
