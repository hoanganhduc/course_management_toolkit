# -*- coding: utf-8 -*-
"""Shared agent-mode helpers for course LMS / DB agent entrypoints."""

from __future__ import annotations

import os
from typing import Optional, Sequence


class CourseAgentError(Exception):
    def __init__(self, message: str, *, code: str = "course_agent_error"):
        super().__init__(message)
        self.code = code


def force_agent_mode() -> None:
    os.environ["COURSE_AGENT_MODE"] = "1"


def is_agent_mode() -> bool:
    return os.environ.get("COURSE_AGENT_MODE", "").strip().lower() in {
        "1",
        "true",
        "yes",
    }


def require_env_allowlist(env_name: str, value: str, *, label: str) -> None:
    """Fail closed when agent mode is on and allowlist is empty or value missing."""
    if not is_agent_mode():
        return
    raw = os.environ.get(env_name, "").strip()
    if not raw:
        raise CourseAgentError(
            f"{env_name} required in agent mode",
            code="allowlist_required",
        )
    allowed = {x.strip() for x in raw.split(",") if x.strip()}
    if value not in allowed:
        raise CourseAgentError(
            f"{label} {value!r} not in {env_name}",
            code="not_allowlisted",
        )


def refuse(op: str) -> None:
    raise CourseAgentError(
        f"{op} is not available on the agent surface",
        code="agent_forbidden",
    )
