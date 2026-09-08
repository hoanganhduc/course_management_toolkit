# -*- coding: utf-8 -*-
"""Pull Classroom50 roster into local student list (ADR D5 merge)."""

from __future__ import annotations

import json
from typing import Any, Dict, List, Optional, Sequence, Tuple

from .c50_cli import AgentCLI, Classroom50Error, Runner
from .c50_roster import apply_fill_only, join_roster, parse_roster_payload


def sync_pull(
    students: List[Any],
    *,
    org: str,
    classroom: str,
    cli: Optional[AgentCLI] = None,
    runner: Optional[Runner] = None,
    agent_mode: bool = False,
) -> Tuple[List[Any], Dict[str, Any]]:
    """
    Fetch remote roster and fill-only merge into students.
    Never auto-creates local students in agent mode.
    """
    cli = cli or AgentCLI(runner=runner)
    payload = cli.list_roster(org, classroom)
    remote = parse_roster_payload(payload)
    report = join_roster(students, remote)

    for remote_row, idx in report.get("pairs", []):
        apply_fill_only(students[idx], remote_row, report)

    # agent/non-interactive: multi_match => nonzero exit_hint already set
    if agent_mode and report.get("multi_match"):
        report["exit_hint"] = 1
    elif not agent_mode and report.get("multi_match"):
        report["exit_hint"] = 1  # ADR: agent and non-interactive same

    # strip internal pairs for JSON export
    public = {k: v for k, v in report.items() if k != "pairs"}
    return students, public


def write_report(report: Dict[str, Any], path: Optional[str]) -> str:
    text = json.dumps(report, indent=2, ensure_ascii=False) + "\n"
    if path:
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(text)
    return text
