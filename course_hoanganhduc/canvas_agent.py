# -*- coding: utf-8 -*-
"""Agent-safe Canvas entrypoint (ADR split skill course-canvas).

Does not expose unenroll, grade, invite, announce, messages, or bulk download.
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import stat
import sys
from typing import Any, List, Optional

from .course_agent_common import (
    CourseAgentError,
    force_agent_mode,
    refuse,
    require_env_allowlist,
)


CANVAS_CONFIG_KEYS = (
    "CANVAS_LMS_API_URL",
    "CANVAS_LMS_API_KEY",
    "CANVAS_LMS_COURSE_ID",
)
MAX_CANVAS_CONFIG_BYTES = 1024 * 1024


def _configure_canvas() -> None:
    """Load the explicit private Canvas config before operation modules import."""

    from . import settings

    selected: dict[str, str] = {}
    configured_path = os.environ.get("CANVAS_CONFIG_PATH")
    if configured_path:
        path = Path(configured_path).expanduser()
        if not path.is_absolute():
            raise CourseAgentError("CANVAS_CONFIG_PATH must be absolute")
        flags = os.O_RDONLY | getattr(os, "O_CLOEXEC", 0) | getattr(os, "O_NOFOLLOW", 0)
        try:
            descriptor = os.open(path, flags)
        except OSError as exc:
            raise CourseAgentError("Canvas config is unavailable or unsafe") from exc
        try:
            before = os.fstat(descriptor)
            if (
                not stat.S_ISREG(before.st_mode)
                or before.st_uid != os.getuid()
                or before.st_nlink != 1
                or stat.S_IMODE(before.st_mode) & 0o077
                or not 0 < before.st_size <= MAX_CANVAS_CONFIG_BYTES
            ):
                raise CourseAgentError("Canvas config metadata is unsafe")
            remaining = before.st_size
            chunks: list[bytes] = []
            while remaining:
                chunk = os.read(descriptor, min(remaining, 65536))
                if not chunk:
                    raise CourseAgentError("Canvas config was truncated")
                chunks.append(chunk)
                remaining -= len(chunk)
            if os.read(descriptor, 1):
                raise CourseAgentError("Canvas config grew while being read")
            after = os.fstat(descriptor)
            identity = lambda info: (
                info.st_dev,
                info.st_ino,
                info.st_mode,
                info.st_uid,
                info.st_nlink,
                info.st_size,
                info.st_mtime_ns,
                info.st_ctime_ns,
            )
            if identity(before) != identity(after):
                raise CourseAgentError("Canvas config changed while being read")
        finally:
            os.close(descriptor)
        try:
            value = json.loads(b"".join(chunks).decode("utf-8"))
        except (UnicodeError, json.JSONDecodeError) as exc:
            raise CourseAgentError("Canvas config is not valid JSON") from exc
        if not isinstance(value, dict):
            raise CourseAgentError("Canvas config root must be an object")
        for key in CANVAS_CONFIG_KEYS:
            candidate = value.get(key)
            if isinstance(candidate, str) and candidate.strip():
                selected[key] = candidate.strip()

    for key in CANVAS_CONFIG_KEYS:
        candidate = os.environ.get(key)
        if candidate:
            selected[key] = candidate.strip()
        if key in selected:
            setattr(settings, key, selected[key])


def _course_id(args_course: Optional[str]) -> Optional[str]:
    from . import settings

    return (
        args_course
        or os.environ.get("CANVAS_LMS_COURSE_ID")
        or getattr(settings, "CANVAS_LMS_COURSE_ID", "")
        or None
    )


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
        _configure_canvas()
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
