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
