# -*- coding: utf-8 -*-
"""Classroom50 facade — agent-safe ops only (no download)."""

from .c50_auth import gh_on_path, gh_teacher_available, preflight
from .c50_cli import AgentCLI, Classroom50Error, agent_cli, redact_secrets
from .c50_ops import (
    export_csv,
    is_agent_mode,
    list_assignments,
    list_classrooms,
    list_roster,
    sync,
)
from .c50_roster import (
    CANONICAL_COLUMNS,
    export_roster_csv,
    join_roster,
    parse_csv_text,
    parse_roster_payload,
)

__all__ = [
    "AgentCLI",
    "Classroom50Error",
    "agent_cli",
    "redact_secrets",
    "preflight",
    "gh_on_path",
    "gh_teacher_available",
    "is_agent_mode",
    "list_classrooms",
    "list_roster",
    "list_assignments",
    "sync",
    "export_csv",
    "export_roster_csv",
    "join_roster",
    "parse_csv_text",
    "parse_roster_payload",
    "CANONICAL_COLUMNS",
]
