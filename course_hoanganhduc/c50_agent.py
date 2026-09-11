# -*- coding: utf-8 -*-
"""Dedicated agent entrypoint — always sets COURSE_C50_AGENT_MODE=1 (ADR D6.A)."""

from __future__ import annotations

import argparse
import json
import os
import sys


# Mutating verbs live on course-c50-admin, for a human operator. They are named here
# so an agent gets a structured refusal instead of an argparse usage error, whatever
# arguments follow.
REFUSED_VERBS = (
    "assignment-add",
    "assignment-remove",
    "invite",
    "roster-import",
    "roster-add",
)


def main(argv=None) -> int:
    os.environ["COURSE_C50_AGENT_MODE"] = "1"

    raw = list(sys.argv[1:] if argv is None else argv)
    if raw and raw[0] in REFUSED_VERBS:
        from .c50_cli import Classroom50Error
        from .c50_ops import agent_refuse

        try:
            agent_refuse(raw[0].replace("-", " "))
        except Classroom50Error as exc:
            print(f"Classroom50 error: {exc}", file=sys.stderr)
            return 1

    parser = argparse.ArgumentParser(
        prog="python -m course_hoanganhduc.c50_agent",
        description="Classroom50 agent-safe operations",
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_pre = sub.add_parser("preflight", help="Auth preflight / whoami")
    p_lc = sub.add_parser("list-classrooms", help="List classrooms")
    p_lc.add_argument("--org", required=True)
    p_lr = sub.add_parser("list-roster", help="List roster")
    p_lr.add_argument("--org", required=True)
    p_lr.add_argument("--classroom", required=True)
    p_la = sub.add_parser("list-assignments", help="List assignments")
    p_la.add_argument("--org", required=True)
    p_la.add_argument("--classroom", required=True)
    p_sync = sub.add_parser("sync", help="Pull roster into local DB")
    p_sync.add_argument("--org", required=True)
    p_sync.add_argument("--classroom", required=True)
    p_sync.add_argument("--db", default="students.db")
    p_sync.add_argument("--report", default=None)
    p_is = sub.add_parser(
        "import-scores", help="Read collected scores into the local DB"
    )
    p_is.add_argument("--org", required=True)
    p_is.add_argument("--classroom", required=True)
    p_is.add_argument("--assignment", default=None)
    p_is.add_argument("--db", default="students.db")
    p_is.add_argument("--report", default=None)
    p_is.add_argument("--dry-run", action="store_true")
    p_lg = sub.add_parser(
        "list-groups", help="Read group membership of the mode=group assignments"
    )
    p_lg.add_argument("--org", required=True)
    p_lg.add_argument("--classroom", required=True)
    p_lg.add_argument("--assignment", default=None)
    p_lg.add_argument("--read-team-json", action="store_true")
    p_eg = sub.add_parser(
        "export-groups", help="Write everything known about the groups to a text file"
    )
    p_eg.add_argument("--org", required=True)
    p_eg.add_argument("--classroom", required=True)
    p_eg.add_argument("--assignment", default=None)
    p_eg.add_argument("--db", default="students.db")
    p_eg.add_argument("--read-team-json", action="store_true")
    p_eg.add_argument("--out", default="classroom50_groups.txt")
    p_ex = sub.add_parser("export", help="Export local roster as C50 CSV")
    p_ex.add_argument("--db", default="students.db")
    p_ex.add_argument("--out", default="classroom50_roster.csv")
    p_dl = sub.add_parser("download", help="(refused) download not agent-safe")
    p_dl.add_argument("--org", default="")
    p_dl.add_argument("--classroom", default="")
    p_dl.add_argument("--assignment", default="")
    p_dl.add_argument("--dest", default="")
    for name in REFUSED_VERBS:
        sub.add_parser(name, help=f"(refused) {name} is not agent-safe")

    args = parser.parse_args(argv)

    from .c50_auth import preflight
    from .c50_cli import Classroom50Error
    from .c50_ops import (
        agent_refuse_download,
        export_csv,
        export_groups_txt,
        import_scores,
        list_assignments,
        list_classrooms,
        list_groups,
        list_roster,
        sync,
    )

    try:
        if args.cmd == "preflight":
            print(preflight())
            return 0
        if args.cmd == "list-classrooms":
            print(json.dumps(list_classrooms(args.org), indent=2, ensure_ascii=False))
            return 0
        if args.cmd == "list-roster":
            print(
                json.dumps(
                    list_roster(args.org, args.classroom),
                    indent=2,
                    ensure_ascii=False,
                )
            )
            return 0
        if args.cmd == "list-assignments":
            print(
                json.dumps(
                    list_assignments(args.org, args.classroom),
                    indent=2,
                    ensure_ascii=False,
                )
            )
            return 0
        if args.cmd == "sync":
            from .data import load_database, save_database

            students = []
            if os.path.exists(args.db):
                students = load_database(args.db, verbose=False) or []
            students, report, _ = sync(
                list(students),
                org=args.org,
                classroom=args.classroom,
                report_path=args.report,
            )
            save_database(
                students, args.db, verbose=False, audit_source="c50_agent_sync"
            )
            print(
                f"matched={len(report.get('matched', []))} "
                f"remote_only={len(report.get('remote_only', []))} "
                f"multi_match={len(report.get('multi_match', []))}"
            )
            return int(report.get("exit_hint") or 0)
        if args.cmd == "import-scores":
            from .c50_scores import format_merge
            from .data import load_database, save_database

            students = []
            if os.path.exists(args.db):
                students = load_database(args.db, verbose=False) or []
            report, _ = import_scores(
                students,
                org=args.org,
                classroom=args.classroom,
                assignment=args.assignment,
                report_path=args.report,
            )
            if not args.dry_run:
                save_database(
                    students,
                    args.db,
                    verbose=False,
                    audit_source="c50_agent_import_scores",
                )
            print(format_merge(report))
            return 0
        if args.cmd == "list-groups":
            print(
                json.dumps(
                    list_groups(
                        args.org,
                        args.classroom,
                        assignment=args.assignment,
                        with_team_json=args.read_team_json,
                    ),
                    indent=2,
                    ensure_ascii=False,
                )
            )
            return 0
        if args.cmd == "export-groups":
            from .data import load_database

            students = []
            if os.path.exists(args.db):
                students = load_database(args.db, verbose=False) or []
            report, text = export_groups_txt(
                students,
                org=args.org,
                classroom=args.classroom,
                assignment=args.assignment,
                with_team_json=args.read_team_json,
            )
            with open(args.out, "w", encoding="utf-8") as fh:
                fh.write(text)
            # The read updates the records in memory, as every read in this
            # module does; an export is not the place to decide that a group is
            # settled, so nothing is saved and the line below says so.
            total = sum(len(rows) for rows in report.groups.values())
            print(
                f"wrote {args.out} groups={total} "
                f"students={len(students)} database=not saved"
            )
            return 0
        if args.cmd == "export":
            from .data import load_database

            students = []
            if os.path.exists(args.db):
                students = load_database(args.db, verbose=False) or []
            text, mode = export_csv(students)
            with open(args.out, "w", encoding="utf-8") as fh:
                fh.write(text)
            print(f"wrote {args.out} name_split={mode}")
            return 0
        if args.cmd == "download":
            agent_refuse_download()
    except Classroom50Error as exc:
        print(f"Classroom50 error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
