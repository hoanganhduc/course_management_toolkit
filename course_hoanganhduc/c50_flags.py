# -*- coding: utf-8 -*-
"""Register and dispatch Classroom50 CLI flags (ADR D3/D4)."""

from __future__ import annotations

import json
import os
from typing import Any, Dict, List, Optional

from .c50_auth import preflight
from .c50_cli import Classroom50Error
from .c50_ops import (
    agent_refuse_download,
    export_csv,
    is_agent_mode,
    list_assignments,
    list_classrooms,
    list_roster,
    sync,
)


def register_classroom50_flags(parser) -> None:
    group = parser.add_argument_group("Classroom50")
    group.add_argument(
        "--sync-classroom50",
        "-sc50",
        action="store_true",
        help="Sync Classroom50 roster into local student database",
        dest="sync_classroom50",
    )
    group.add_argument(
        "--classroom50-org",
        type=str,
        default=None,
        help="Classroom50 GitHub organization",
        dest="classroom50_org",
    )
    group.add_argument(
        "--classroom50-classroom",
        type=str,
        default=None,
        help="Classroom50 classroom short-name",
        dest="classroom50_classroom",
    )
    group.add_argument(
        "--classroom50-assignment",
        type=str,
        default=None,
        help="Classroom50 assignment slug (human download)",
        dest="classroom50_assignment",
    )
    group.add_argument(
        "--list-classroom50-classrooms",
        action="store_true",
        help="List Classroom50 classrooms",
        dest="list_classroom50_classrooms",
    )
    group.add_argument(
        "--list-classroom50-roster",
        action="store_true",
        help="List Classroom50 roster",
        dest="list_classroom50_roster",
    )
    group.add_argument(
        "--list-classroom50-assignments",
        action="store_true",
        help="List Classroom50 assignments",
        dest="list_classroom50_assignments",
    )
    group.add_argument(
        "--export-classroom50-roster",
        type=str,
        nargs="?",
        const="classroom50_roster.csv",
        default=None,
        help="Export local roster to Classroom50 CSV (default: classroom50_roster.csv)",
        dest="export_classroom50_roster",
        metavar="PATH",
    )
    group.add_argument(
        "--classroom50-report",
        type=str,
        default=None,
        help="Write Classroom50 JSON report to PATH",
        dest="classroom50_report",
        metavar="PATH",
    )
    group.add_argument(
        "--download-classroom50",
        action="store_true",
        help="Download student submissions (human CLI only)",
        dest="download_classroom50",
    )
    group.add_argument(
        "--classroom50-download-dest",
        type=str,
        default=None,
        help="Destination directory for Classroom50 downloads",
        dest="classroom50_download_dest",
        metavar="DIR",
    )
    group.add_argument(
        "--classroom50-preflight",
        action="store_true",
        help="Run Classroom50 auth preflight (whoami)",
        dest="classroom50_preflight",
    )


def _c50_action_requested(args) -> bool:
    return bool(
        getattr(args, "sync_classroom50", False)
        or getattr(args, "list_classroom50_classrooms", False)
        or getattr(args, "list_classroom50_roster", False)
        or getattr(args, "list_classroom50_assignments", False)
        or getattr(args, "export_classroom50_roster", None) is not None
        or getattr(args, "download_classroom50", False)
        or getattr(args, "classroom50_preflight", False)
    )


def _resolve_org(args, config: Optional[Dict]) -> Optional[str]:
    org = getattr(args, "classroom50_org", None)
    if org:
        return org
    if config:
        return config.get("CLASSROOM50_ORG") or config.get("classroom50_org")
    return os.environ.get("CLASSROOM50_ORG")


def _resolve_classroom(args, config: Optional[Dict]) -> Optional[str]:
    c = getattr(args, "classroom50_classroom", None)
    if c:
        return c
    if config:
        return config.get("CLASSROOM50_CLASSROOM") or config.get("classroom50_classroom")
    return os.environ.get("CLASSROOM50_CLASSROOM")


def dispatch_classroom50(
    args,
    students: List[Any],
    db_path: str,
    config: Optional[Dict] = None,
) -> Optional[int]:
    """
    Handle Classroom50 actions. Return exit code if handled, else None.
    Must be called after DB load when students list is available.
    """
    if not _c50_action_requested(args):
        return None

    config = config or {}
    verbose = bool(getattr(args, "verbose", False))

    try:
        if getattr(args, "classroom50_preflight", False):
            login = preflight()
            print(login)
            return 0

        if getattr(args, "download_classroom50", False):
            if is_agent_mode():
                agent_refuse_download()
            # human path
            org = _resolve_org(args, config)
            classroom = _resolve_classroom(args, config)
            assignment = getattr(args, "classroom50_assignment", None)
            dest = getattr(args, "classroom50_download_dest", None)
            if not org or not classroom or not assignment or not dest:
                print(
                    "download requires --classroom50-org, --classroom50-classroom, "
                    "--classroom50-assignment, and --classroom50-download-dest"
                )
                return 2
            from .c50_cli_human import HumanCLI

            HumanCLI().download(org, classroom, assignment, dest)
            print(f"Downloaded to {dest}")
            return 0

        org = _resolve_org(args, config)
        classroom = _resolve_classroom(args, config)

        if getattr(args, "list_classroom50_classrooms", False):
            if not org:
                print("--classroom50-org required (or CLASSROOM50_ORG)")
                return 2
            data = list_classrooms(org)
            print(json.dumps(data, indent=2, ensure_ascii=False))
            return 0

        if getattr(args, "list_classroom50_roster", False):
            if not org or not classroom:
                print("org and classroom required for roster list")
                return 2
            data = list_roster(org, classroom)
            print(json.dumps(data, indent=2, ensure_ascii=False))
            return 0

        if getattr(args, "list_classroom50_assignments", False):
            if not org or not classroom:
                print("org and classroom required for assignment list")
                return 2
            data = list_assignments(org, classroom)
            print(json.dumps(data, indent=2, ensure_ascii=False))
            return 0

        if getattr(args, "sync_classroom50", False):
            if not org or not classroom:
                print("org and classroom required for sync")
                return 2
            students_out, report, text = sync(
                list(students),
                org=org,
                classroom=classroom,
                report_path=getattr(args, "classroom50_report", None),
            )
            students.clear()
            students.extend(students_out)
            try:
                from .data import save_database

                save_database(
                    students,
                    db_path,
                    verbose=verbose,
                    audit_source="classroom50_sync",
                )
            except Exception as exc:
                if verbose:
                    print(f"[Classroom50] save failed: {exc}")
                print(f"Sync done but save failed: {exc}")
                return 1
            # human summary
            print(
                f"Classroom50 sync: matched={len(report.get('matched', []))} "
                f"remote_only={len(report.get('remote_only', []))} "
                f"local_only={len(report.get('local_only', []))} "
                f"multi_match={len(report.get('multi_match', []))} "
                f"conflicts={len(report.get('conflicts', []))}"
            )
            if not getattr(args, "classroom50_report", None):
                # short JSON to stdout when no path
                pass
            return int(report.get("exit_hint") or 0)

        if getattr(args, "export_classroom50_roster", None) is not None:
            path = args.export_classroom50_roster or "classroom50_roster.csv"
            csv_text, mode = export_csv(students)
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(csv_text)
            print(f"Exported Classroom50 roster to {path} (name_split={mode})")
            return 0

    except Classroom50Error as exc:
        print(f"Classroom50 error: {exc}")
        return 1

    return None
