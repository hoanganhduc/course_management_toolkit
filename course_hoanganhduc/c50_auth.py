# -*- coding: utf-8 -*-
"""Classroom50 auth / preflight (ADR D3 / D7 A4)."""

from __future__ import annotations

import shutil
import subprocess
from typing import Optional, Tuple

from .c50_cli import AgentCLI, Classroom50Error, Runner


def gh_on_path() -> bool:
    return shutil.which("gh") is not None


def gh_teacher_available() -> Tuple[bool, str]:
    """Return (ok, message). Checks extension or help."""
    if not gh_on_path():
        return False, "gh not found on PATH; install GitHub CLI"
    try:
        proc = subprocess.run(
            ["gh", "teacher", "--help"],
            capture_output=True,
            text=True,
            check=False,
            timeout=30,
        )
    except Exception as exc:
        return False, f"gh teacher probe failed: {exc}"
    if proc.returncode != 0:
        # fallback: extension list
        try:
            ext = subprocess.run(
                ["gh", "extension", "list"],
                capture_output=True,
                text=True,
                check=False,
                timeout=30,
            )
            text = (ext.stdout or "") + (ext.stderr or "")
            if "teacher" in text.lower():
                return True, "gh teacher extension listed"
        except Exception:
            pass
        return (
            False,
            "gh teacher not available; install Classroom50 teacher extension "
            "(see foundation50/classroom50 cli docs)",
        )
    return True, "gh teacher --help ok"


def preflight(*, runner: Optional[Runner] = None) -> str:
    """Raise Classroom50Error if preflight fails; return whoami login on success."""
    ok, msg = gh_teacher_available()
    if not ok:
        raise Classroom50Error(msg, code="preflight_failed")
    cli = AgentCLI(runner=runner)
    return cli.whoami()
