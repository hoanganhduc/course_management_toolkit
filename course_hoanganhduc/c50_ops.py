# -*- coding: utf-8 -*-
"""Classroom50 operations with agent-mode gates (ADR D6)."""

from __future__ import annotations

import json
import os
from typing import Any, Dict, List, Optional, Sequence

from .c50_cli import AgentCLI, Classroom50Error, Runner
from .c50_roster import export_roster_csv, parse_roster_payload
from .c50_sync import sync_pull, write_report


def is_agent_mode() -> bool:
    return os.environ.get("COURSE_C50_AGENT_MODE", "").strip() in (
        "1",
        "true",
        "TRUE",
        "yes",
        "YES",
    )


def require_org_allowlist(org: str) -> None:
    if not is_agent_mode():
        return
    raw = os.environ.get("CLASSROOM50_ORG_ALLOWLIST", "").strip()
    if not raw:
        raise Classroom50Error(
            "CLASSROOM50_ORG_ALLOWLIST required in agent mode",
            code="allowlist_required",
        )
    allowed = {x.strip() for x in raw.split(",") if x.strip()}
    if org not in allowed:
        raise Classroom50Error(
            f"org {org!r} not in CLASSROOM50_ORG_ALLOWLIST",
            code="org_not_allowlisted",
        )


def ensure_agent_safe() -> None:
    """Agent entry always forces mode."""
    os.environ["COURSE_C50_AGENT_MODE"] = "1"


def list_classrooms(
    org: str,
    *,
    cli: Optional[AgentCLI] = None,
    runner: Optional[Runner] = None,
) -> Any:
    if is_agent_mode():
        require_org_allowlist(org)
    cli = cli or AgentCLI(runner=runner)
    return cli.list_classrooms(org)


def list_roster(
    org: str,
    classroom: str,
    *,
    cli: Optional[AgentCLI] = None,
    runner: Optional[Runner] = None,
) -> Any:
    if is_agent_mode():
        require_org_allowlist(org)
        if not classroom:
            raise Classroom50Error("classroom required", code="missing_classroom")
    cli = cli or AgentCLI(runner=runner)
    return cli.list_roster(org, classroom)


def list_assignments(
    org: str,
    classroom: str,
    *,
    cli: Optional[AgentCLI] = None,
    runner: Optional[Runner] = None,
) -> Any:
    if is_agent_mode():
        require_org_allowlist(org)
        if not classroom:
            raise Classroom50Error("classroom required", code="missing_classroom")
    cli = cli or AgentCLI(runner=runner)
    return cli.list_assignments(org, classroom)


def sync(
    students: List[Any],
    *,
    org: str,
    classroom: str,
    report_path: Optional[str] = None,
    cli: Optional[AgentCLI] = None,
    runner: Optional[Runner] = None,
) -> tuple:
    if is_agent_mode():
        require_org_allowlist(org)
        if not classroom:
            raise Classroom50Error("classroom required", code="missing_classroom")
    students, report = sync_pull(
        students,
        org=org,
        classroom=classroom,
        cli=cli,
        runner=runner,
        agent_mode=is_agent_mode(),
    )
    text = write_report(report, report_path)
    return students, report, text


def export_csv(students: Sequence[Any]) -> tuple:
    return export_roster_csv(students)


def agent_refuse_download() -> None:
    raise Classroom50Error(
        "download is not available in agent mode",
        code="agent_download_forbidden",
    )


def agent_refuse(op: str) -> None:
    raise Classroom50Error(
        f"{op} is not available in agent mode",
        code="agent_forbidden",
    )


def _human_cli(cli: Optional[Any], runner: Optional[Runner]) -> Any:
    from .c50_cli_human import HumanCLI

    return cli or HumanCLI(runner=runner)


def find_assignment(records: Any, slug: str) -> Optional[Dict[str, Any]]:
    """Return the `assignments.json` entry with this slug, or None."""
    if records is None:
        records = []
    if not isinstance(records, list):
        raise Classroom50Error(
            "assignment manifest is not a JSON array", code="bad_manifest"
        )
    for item in records:
        if isinstance(item, dict) and str(item.get("slug", "")) == slug:
            return item
    return None


def assignment_allows_by_pattern(record: Optional[Dict[str, Any]]) -> bool:
    """`--by-pattern` fits only empty-repository assignments.

    It skips the team lookup, so it fetches no `result.json` and writes no
    `scores.csv`. An empty-repository assignment has neither to collect; an
    autograded one does, and would silently lose them.
    """
    if not isinstance(record, dict):
        return False
    return record.get("empty_repo") is True


def existing_member_logins(rows: Any) -> set:
    """Logins that are already org members or hold a pending invitation."""
    if not isinstance(rows, list):
        raise Classroom50Error(
            "member list is not a JSON array", code="bad_member_list"
        )
    logins = set()
    for row in rows:
        if isinstance(row, dict):
            login = str(row.get("login", "")).strip()
            if login:
                logins.add(login.lower())
    return logins


def preflight_assignment_add(
    org: str,
    classroom: str,
    slug: str,
    *,
    allow_overwrite: bool = False,
    cli: Optional[Any] = None,
    runner: Optional[Runner] = None,
) -> Optional[Dict[str, Any]]:
    """Return the existing entry for this slug, refusing an unintended overwrite."""
    if is_agent_mode():
        agent_refuse("assignment add")
    cli = _human_cli(cli, runner)
    record = find_assignment(cli.list_assignments(org, classroom), slug)
    if record is not None and not allow_overwrite:
        raise Classroom50Error(
            f"assignment {slug!r} already exists. `assignment add` replaces the entry "
            "in place, and the pinned CLI has no way to set submission mode, so a "
            "re-run restores the default every-push submission mode and discards a "
            "tagged-commit setting made in the web form. Pass --allow-overwrite only "
            "if that is intended.",
            code="assignment_exists",
        )
    return record


def preflight_assignment_remove(
    org: str,
    classroom: str,
    slug: str,
    *,
    cli: Optional[Any] = None,
    runner: Optional[Runner] = None,
) -> Dict[str, Any]:
    """Return the entry that would be removed, refusing a no-op removal."""
    if is_agent_mode():
        agent_refuse("assignment remove")
    cli = _human_cli(cli, runner)
    record = find_assignment(cli.list_assignments(org, classroom), slug)
    if record is None:
        raise Classroom50Error(
            f"assignment {slug!r} is not registered in {org}/{classroom}. "
            "`assignment remove` is idempotent and would exit 0 without changing "
            "anything.",
            code="assignment_absent",
        )
    return record


def download_submissions(
    org: str,
    classroom: str,
    assignment: str,
    dest: str,
    *,
    by_pattern: bool = False,
    cli: Optional[Any] = None,
    runner: Optional[Runner] = None,
) -> Any:
    """Download submissions, gating `--by-pattern` on the recorded assignment mode."""
    if is_agent_mode():
        agent_refuse("download")
    cli = _human_cli(cli, runner)
    verified = False
    if by_pattern:
        record = find_assignment(cli.list_assignments(org, classroom), assignment)
        if record is None:
            raise Classroom50Error(
                f"assignment {assignment!r} is not registered in {org}/{classroom}, so "
                "its empty-repository mode cannot be verified",
                code="assignment_not_found",
            )
        if not assignment_allows_by_pattern(record):
            raise Classroom50Error(
                f"assignment {assignment!r} is not an empty-repository assignment; "
                "--by-pattern would skip its result.json and scores.csv summary",
                code="by_pattern_not_permitted",
            )
        verified = True
    return cli.download(
        org,
        classroom,
        assignment,
        dest,
        by_pattern=by_pattern,
        empty_repo_verified=verified,
    )


def invite_users(
    target: str,
    usernames: Sequence[str],
    *,
    admin: bool = False,
    permission: Optional[str] = None,
    cli: Optional[Any] = None,
    runner: Optional[Runner] = None,
) -> Dict[str, Any]:
    """Invite users to an org or repo; report invited, skipped, and failed separately.

    Repository invitations are idempotent, so they are sent directly. Organization
    invitations are not: GitHub rejects a re-invite to a pending or existing member,
    so actual membership is read first and those logins are skipped.
    """
    if is_agent_mode():
        agent_refuse("invite")
    from .c50_cli_human import validate_target

    target = validate_target(target)
    cli = _human_cli(cli, runner)

    wanted: List[str] = []
    seen = set()
    for raw in usernames:
        name = str(raw or "").strip()
        if not name or name.lower() in seen:
            continue
        seen.add(name.lower())
        wanted.append(name)

    report: Dict[str, Any] = {
        "target": target,
        "invited": [],
        "skipped": [],
        "failed": [],
    }
    existing = set()
    if "/" not in target:
        existing = existing_member_logins(cli.member_list(target))

    for username in wanted:
        if username.lower() in existing:
            report["skipped"].append(username)
            continue
        try:
            cli.invite(target, username, admin=admin, permission=permission)
        except Classroom50Error as exc:
            report["failed"].append(
                {"username": username, "code": exc.code, "message": str(exc)}
            )
            continue
        report["invited"].append(username)
    return report
