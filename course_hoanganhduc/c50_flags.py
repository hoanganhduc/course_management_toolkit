# -*- coding: utf-8 -*-
"""Register and dispatch Classroom50 CLI flags (ADR D3/D4)."""

from __future__ import annotations

import json
import os
from typing import Any, Dict, List, NamedTuple, Optional

from .c50_auth import preflight
from .c50_cli import Classroom50Error
from .c50_ops import (
    agent_refuse_download,
    export_csv,
    import_groups,
    import_project_issues,
    import_scores,
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
        "--list-classroom50-membership",
        action="store_true",
        help="Check each roster GitHub account against the org: member, invited, or never invited",
        dest="list_classroom50_membership",
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
        "--import-classroom50-scores",
        action="store_true",
        help=(
            "Read the classroom's collected scores into the local database "
            "(final-project has autograding off and never appears there: enter "
            "those marks with --load-override-grades)"
        ),
        dest="import_classroom50_scores",
    )
    group.add_argument(
        "--classroom50-scores-report",
        type=str,
        default=None,
        help="Write the score-import JSON report to PATH",
        dest="classroom50_scores_report",
        metavar="PATH",
    )
    group.add_argument(
        "--import-classroom50-groups",
        action="store_true",
        help=(
            "Read group membership of the mode=group assignments into the local "
            "database (a group in conflict keeps its recorded value and is "
            "reported instead of overwritten)"
        ),
        dest="import_classroom50_groups",
    )
    group.add_argument(
        "--classroom50-group-assignment",
        type=str,
        default=None,
        help="Limit the group import to one assignment slug (default: every mode=group assignment)",
        dest="classroom50_group_assignment",
        metavar="SLUG",
    )
    group.add_argument(
        "--classroom50-read-team-json",
        action="store_true",
        help="Also read team.json from each group repository (one extra read per repository)",
        dest="classroom50_read_team_json",
    )
    group.add_argument(
        "--classroom50-groups-report",
        type=str,
        default=None,
        help="Write the group-import JSON report to PATH",
        dest="classroom50_groups_report",
        metavar="PATH",
    )
    group.add_argument(
        "--import-project-issues",
        action="store_true",
        help=(
            "Read the mini-project topic board's issues into the local database "
            "(read-only: no label, comment or issue on the board is changed)"
        ),
        dest="import_project_issues",
    )
    group.add_argument(
        "--project-issues-repo",
        type=str,
        default=None,
        help=(
            "Topic board as OWNER/REPO (or CLASSROOM50_PROJECT_ISSUES_REPO)"
        ),
        dest="project_issues_repo",
        metavar="SLUG",
    )
    group.add_argument(
        "--project-issues-report",
        type=str,
        default=None,
        help="Write the issue-import JSON report to PATH",
        dest="project_issues_report",
        metavar="PATH",
    )
    group.add_argument(
        "--project-issues-read-proposal",
        action="store_true",
        help=(
            "Also read proposal/proposal.md at the pinned commit of each "
            "proposal (one extra read per group repository)"
        ),
        dest="project_issues_read_proposal",
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
        or getattr(args, "list_classroom50_membership", False)
        or getattr(args, "export_classroom50_roster", None) is not None
        or getattr(args, "import_classroom50_scores", False)
        or getattr(args, "import_classroom50_groups", False)
        or getattr(args, "import_project_issues", False)
        or getattr(args, "download_classroom50", False)
        or getattr(args, "classroom50_preflight", False)
    )


SOURCE_FLAG = "flag"
SOURCE_CONFIG = "config"
SOURCE_ENV = "env"


class Resolved(NamedTuple):
    """A setting and the tier it came from; ``value`` is ``None`` if no tier had one.

    The tier matters to callers that treat the tiers differently.  A value typed
    on the command line was chosen for this run, while one from the config file
    or the environment was chosen once and then applies to every run after that,
    which is the difference between an instruction and a default.
    """

    value: Optional[str] = None
    source: str = SOURCE_FLAG


def resolve_setting(args, config: Optional[Dict], *, attribute: str, key: str) -> Resolved:
    """Read one setting from the flag, then the config file, then the environment.

    Each tier falls through when it holds no value, so a config file that says
    nothing about a setting does not shadow the matching environment variable.
    The config tier only works for a key ``load_config`` whitelists; it drops
    any key outside ``known_keys``, lower-case spellings included.

    Nothing here is specific to Classroom50 -- it lives in this module because
    this is where the tier order was first written, and the onboarding lane
    reads the Google course id through it so the two cannot drift apart.
    """
    flag = getattr(args, attribute, None)
    if flag:
        return Resolved(str(flag), SOURCE_FLAG)
    if config and config.get(key):
        return Resolved(str(config.get(key)), SOURCE_CONFIG)
    environment = os.environ.get(key)
    if environment:
        return Resolved(environment, SOURCE_ENV)
    return Resolved()


def resolve_org(args, config: Optional[Dict]) -> Resolved:
    """The Classroom50 org, with the tier it came from."""
    return resolve_setting(args, config, attribute="classroom50_org", key="CLASSROOM50_ORG")


def resolve_classroom(args, config: Optional[Dict]) -> Resolved:
    """The Classroom50 classroom, with the tier it came from."""
    return resolve_setting(
        args, config, attribute="classroom50_classroom", key="CLASSROOM50_CLASSROOM"
    )


def resolve_project_issues_repo(args, config: Optional[Dict]) -> Resolved:
    """The mini-project topic board, with the tier it came from.

    Read here rather than in the issue module so the key passes through
    ``resolve_setting`` and stays inside the guard that checks every key this
    module reads against the ``load_config`` whitelist.  A key read elsewhere
    would be dropped silently by ``load_config`` and nothing would notice.
    """
    return resolve_setting(
        args,
        config,
        attribute="project_issues_repo",
        key="CLASSROOM50_PROJECT_ISSUES_REPO",
    )


def _resolve_org(args, config: Optional[Dict]) -> Optional[str]:
    """The org alone, for the callers that do not care which tier supplied it."""
    return resolve_org(args, config).value


def _resolve_classroom(args, config: Optional[Dict]) -> Optional[str]:
    """The classroom alone, for the callers that do not care which tier supplied it."""
    return resolve_classroom(args, config).value


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

        if getattr(args, "list_classroom50_membership", False):
            if not org:
                print("--classroom50-org required (or CLASSROOM50_ORG)")
                return 2
            from .roster_audit import (
                format_membership,
                roster_github_usernames,
                roster_usernames_from_payload,
                usernames_not_on_roster,
                verify_org_membership,
            )

            # The organization is asked about the accounts that actually reached
            # the roster, not about the database: a name that was never imported
            # is not an uninvited student, and reporting it as one sends the
            # teacher looking for an invitation that could not have been sent.
            db_usernames = roster_github_usernames(students)
            live: Optional[List[str]] = None
            if classroom:
                try:
                    live = roster_usernames_from_payload(list_roster(org, classroom))
                except Classroom50Error as exc:
                    if verbose:
                        print(f"[Classroom50] roster list unreadable: {exc}")
            if live is None:
                # No classroom given, or the roster could not be read.  Fall back
                # to the database and let the header say that is what happened.
                population, wanted, not_on_roster = "database", db_usernames, []
            else:
                population = "roster"
                wanted = live
                not_on_roster = usernames_not_on_roster(db_usernames, live)
            print(
                format_membership(
                    verify_org_membership(wanted, org, verbose=verbose),
                    not_on_roster=not_on_roster,
                    population=population,
                )
            )
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

        if getattr(args, "import_classroom50_scores", False):
            if not org or not classroom:
                print("org and classroom required for score import")
                return 2
            # Deferred: c50_scores reaches roster_audit -> onboard, and onboard
            # imports this module, so a module-scope import closes a cycle.
            from .c50_scores import format_merge

            report, _text = import_scores(
                students,
                org=org,
                classroom=classroom,
                assignment=getattr(args, "classroom50_assignment", None),
                report_path=getattr(args, "classroom50_scores_report", None),
            )
            print(format_merge(report))
            if getattr(args, "dry_run", False):
                print("Dry run: database not saved.")
                return 0
            try:
                from .data import save_database

                save_database(
                    students,
                    db_path,
                    verbose=verbose,
                    audit_source="classroom50_scores",
                )
            except Exception as exc:
                if verbose:
                    print(f"[Classroom50] save failed: {exc}")
                print(f"Score import done but save failed: {exc}")
                return 1
            return 0

        if getattr(args, "import_classroom50_groups", False):
            if not org or not classroom:
                print("org and classroom required for group import")
                return 2
            # Deferred for the same cycle as the score import above.
            from .c50_groups import format_groups, group_snapshot

            report, _text = import_groups(
                students,
                org=org,
                classroom=classroom,
                assignment=getattr(args, "classroom50_group_assignment", None),
                with_team_json=bool(
                    getattr(args, "classroom50_read_team_json", False)
                ),
                previous=group_snapshot(students),
                report_path=getattr(args, "classroom50_groups_report", None),
            )
            print(format_groups(report))
            if getattr(args, "dry_run", False):
                print("Dry run: database not saved.")
                return 0
            try:
                from .data import save_database

                save_database(
                    students,
                    db_path,
                    verbose=verbose,
                    audit_source="classroom50_groups",
                )
            except Exception as exc:
                if verbose:
                    print(f"[Classroom50] save failed: {exc}")
                print(f"Group import done but save failed: {exc}")
                return 1
            return 0

        if getattr(args, "import_project_issues", False):
            board = resolve_project_issues_repo(args, config).value
            if not board:
                print(
                    "--project-issues-repo OWNER/REPO required "
                    "(or CLASSROOM50_PROJECT_ISSUES_REPO)"
                )
                return 2
            # Deferred for the same cycle as the score import above.
            from .project_issues import (
                format_project_issues,
                issue_snapshot,
                recorded_groups,
            )

            groups = recorded_groups(students)
            report, _text = import_project_issues(
                students,
                board=board,
                groups=groups,
                known_repos=sorted(groups),
                previous=issue_snapshot(students),
                with_proposal_md=bool(
                    getattr(args, "project_issues_read_proposal", False)
                ),
                report_path=getattr(args, "project_issues_report", None),
            )
            print(format_project_issues(report))
            if getattr(args, "dry_run", False):
                print("Dry run: database not saved.")
                return 0
            try:
                from .data import save_database

                save_database(
                    students,
                    db_path,
                    verbose=verbose,
                    audit_source="project_issues",
                )
            except Exception as exc:
                if verbose:
                    print(f"[Classroom50] save failed: {exc}")
                print(f"Issue import done but save failed: {exc}")
                return 1
            return 0

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
