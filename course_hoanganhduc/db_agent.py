# -*- coding: utf-8 -*-
"""Agent-safe local student DB operations (course-db skill).

Read/search/export only. No interactive modify, restore, or bulk import apply.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import sys
from typing import Any, List, Optional


from .course_agent_common import CourseAgentError, force_agent_mode, refuse


def _load(db: str) -> List[Any]:
    from .data import load_database

    if not os.path.exists(db):
        return []
    return load_database(db, verbose=False) or []


def _student_dict(s: Any) -> dict:
    return {k: v for k, v in getattr(s, "__dict__", {}).items() if v not in (None, "")}


def main(argv: Optional[List[str]] = None) -> int:
    force_agent_mode()
    parser = argparse.ArgumentParser(
        prog="python -m course_hoanganhduc.db_agent",
        description="Agent-safe local student database operations",
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_search = sub.add_parser("search", help="Search students by keyword")
    p_search.add_argument("query")
    p_search.add_argument("--db", default="students.db")
    p_det = sub.add_parser("details", help="Show one student by id/email/name")
    p_det.add_argument("identifier")
    p_det.add_argument("--db", default="students.db")
    p_led = sub.add_parser("list-email-domain", help="List students by email domain")
    p_led.add_argument("domain")
    p_led.add_argument("--db", default="students.db")
    p_dup = sub.add_parser("list-duplicate-names", help="List duplicate full names")
    p_dup.add_argument("--db", default="students.db")
    p_miss = sub.add_parser("list-missing-ids", help="List missing Google/Canvas/Student IDs")
    p_miss.add_argument("--which", default="all", help="google,canvas,student,all")
    p_miss.add_argument("--db", default="students.db")
    p_er = sub.add_parser("export-roster", help="Export local classroom roster CSV")
    p_er.add_argument("--db", default="students.db")
    p_er.add_argument("--out", default="classroom_roster.csv")
    p_ee = sub.add_parser("export-emails", help="Export unique emails")
    p_ee.add_argument("--db", default="students.db")
    p_ee.add_argument("--out", default="emails.txt")
    p_val = sub.add_parser("count", help="Count students in DB")
    p_val.add_argument("--db", default="students.db")
    for bad in ("modify", "restore-db", "import-apply", "delete"):
        sub.add_parser(bad, help="(refused) not agent-safe")

    args = parser.parse_args(argv)

    try:
        if args.cmd in {"modify", "restore-db", "import-apply", "delete"}:
            refuse(args.cmd)

        students = _load(getattr(args, "db", "students.db"))

        if args.cmd == "count":
            print(json.dumps({"count": len(students), "db": args.db}))
            return 0

        if args.cmd == "search":
            q = args.query.lower()
            hits = []
            for s in students:
                blob = " ".join(str(v) for v in _student_dict(s).values()).lower()
                if q in blob:
                    hits.append(_student_dict(s))
            print(json.dumps(hits, indent=2, ensure_ascii=False, default=str))
            return 0

        if args.cmd == "details":
            ident = args.identifier.lower()
            for s in students:
                d = _student_dict(s)
                blob = " ".join(str(v) for v in d.values()).lower()
                if ident in blob:
                    print(json.dumps(d, indent=2, ensure_ascii=False, default=str))
                    return 0
            print(json.dumps({"found": False}))
            return 1

        if args.cmd == "list-email-domain":
            dom = args.domain.lower().lstrip("@")
            hits = []
            for s in students:
                email = str(getattr(s, "Email", "") or "").lower()
                if email.endswith("@" + dom) or email.endswith(dom):
                    hits.append(_student_dict(s))
            print(json.dumps(hits, indent=2, ensure_ascii=False, default=str))
            return 0

        if args.cmd == "list-duplicate-names":
            from collections import defaultdict

            buckets = defaultdict(list)
            for s in students:
                name = str(getattr(s, "Name", "") or "").strip().lower()
                if name:
                    buckets[name].append(_student_dict(s))
            dups = {k: v for k, v in buckets.items() if len(v) > 1}
            print(json.dumps(dups, indent=2, ensure_ascii=False, default=str))
            return 0

        if args.cmd == "list-missing-ids":
            which = (args.which or "all").lower()
            missing = []
            for s in students:
                d = _student_dict(s)
                issues = []
                if which in ("all", "student") and not d.get("Student ID"):
                    issues.append("student")
                if which in ("all", "canvas") and not (
                    d.get("Canvas ID") or d.get("Canvas_ID")
                ):
                    issues.append("canvas")
                if which in ("all", "google") and not (
                    d.get("Google_ID")
                    or d.get("Google ID")
                    or d.get("Google Classroom Display Name")
                ):
                    issues.append("google")
                if issues:
                    missing.append({"student": d, "missing": issues})
            print(json.dumps(missing, indent=2, ensure_ascii=False, default=str))
            return 0

        if args.cmd == "export-roster":
            # Prefer toolkit helper when available
            try:
                from .data import export_roster_to_csv

                export_roster_to_csv(students, file_path=args.out, verbose=False)
            except Exception:
                with open(args.out, "w", encoding="utf-8", newline="") as fh:
                    w = csv.writer(fh)
                    w.writerow(["Name", "Student ID", "Email", "Section"])
                    for s in students:
                        w.writerow(
                            [
                                getattr(s, "Name", ""),
                                getattr(s, "Student ID", ""),
                                getattr(s, "Email", ""),
                                getattr(s, "Section", ""),
                            ]
                        )
            print(json.dumps({"wrote": args.out, "count": len(students)}))
            return 0

        if args.cmd == "export-emails":
            seen = set()
            lines = []
            for s in students:
                email = str(getattr(s, "Email", "") or "").strip()
                if email and email not in seen:
                    seen.add(email)
                    lines.append(email)
            with open(args.out, "w", encoding="utf-8") as fh:
                fh.write("\n".join(lines) + ("\n" if lines else ""))
            print(json.dumps({"wrote": args.out, "count": len(lines)}))
            return 0

    except CourseAgentError as exc:
        print(f"db_agent error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # pragma: no cover
        print(f"db_agent error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
