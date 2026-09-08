# -*- coding: utf-8 -*-
"""Agent-safe closed method-map adapter over `gh teacher` (ADR v2.1 D1.A)."""

from __future__ import annotations

import json
import os
import re
import subprocess
from dataclasses import dataclass
from typing import Any, Callable, Optional

# Injected for tests: (argv: list[str]) -> CompletedProcess-like
Runner = Callable[[list[str]], "RunResult"]


@dataclass
class RunResult:
    returncode: int
    stdout: str = ""
    stderr: str = ""


class Classroom50Error(Exception):
    """Structured failure from the Classroom50 adapter."""

    def __init__(self, message: str, *, code: str = "c50_error"):
        super().__init__(message)
        self.code = code


# Full-token redaction (ADR D1.C / R2-M15)
_REDACT_PATTERNS = [
    re.compile(r"ghp_[A-Za-z0-9_]{20,}", re.I),
    re.compile(r"gho_[A-Za-z0-9_]{20,}", re.I),
    re.compile(r"github_pat_[A-Za-z0-9_]{20,}", re.I),
    re.compile(r"Bearer\s+\S+", re.I),
    re.compile(r"CLASSROOM50_SERVICE_TOKEN=\S+", re.I),
    re.compile(r"(password|secret|token|api[_-]?key)\s*[:=]\s*\S+", re.I),
]


def redact_secrets(text: str) -> str:
    if not text:
        return text
    out = text
    for pat in _REDACT_PATTERNS:
        out = pat.sub("<redacted>", out)
    return out


def _default_runner(argv: list[str]) -> RunResult:
    try:
        proc = subprocess.run(
            argv,
            capture_output=True,
            text=True,
            check=False,
        )
    except FileNotFoundError as exc:
        raise Classroom50Error(
            f"command not found: {argv[0]}", code="missing_binary"
        ) from exc
    return RunResult(
        returncode=proc.returncode,
        stdout=proc.stdout or "",
        stderr=proc.stderr or "",
    )


class AgentCLI:
    """Closed method map — no free-form argv, no download (ADR D1.A)."""

    ALLOWED = frozenset(
        {"whoami", "list_classrooms", "list_roster", "list_assignments"}
    )

    def __init__(
        self,
        *,
        runner: Optional[Runner] = None,
        gh_teacher: str = "gh",
    ):
        self._runner = runner or _default_runner
        self._gh = gh_teacher

    def _run(self, parts: list[str]) -> RunResult:
        argv = [self._gh, "teacher", *parts]
        result = self._runner(argv)
        # Log sinks only (parse uses raw stdout)
        _ = redact_secrets(result.stderr)
        _ = redact_secrets(" ".join(argv))
        return result

    def __getattr__(self, name: str):
        if name.startswith("_"):
            raise AttributeError(name)
        if name not in self.ALLOWED:
            raise Classroom50Error(
                f"unknown or forbidden agent method: {name}",
                code="unknown_method",
            )
        raise AttributeError(name)

    def whoami(self) -> str:
        result = self._run(["whoami"])
        if result.returncode != 0:
            raise Classroom50Error(
                redact_secrets(result.stderr or "whoami failed"),
                code="whoami_failed",
            )
        login = (result.stdout or "").strip().splitlines()
        if not login or not login[0].strip():
            raise Classroom50Error("whoami returned empty login", code="whoami_empty")
        return login[0].strip()

    def list_classrooms(self, org: str) -> Any:
        if not org:
            raise Classroom50Error("org required", code="missing_org")
        result = self._run(["classroom", "list", org, "--json"])
        return self._parse_json(result, op="list_classrooms")

    def list_roster(self, org: str, classroom: str) -> Any:
        if not org or not classroom:
            raise Classroom50Error("org and classroom required", code="missing_params")
        result = self._run(["roster", "list", org, classroom, "--json"])
        return self._parse_json(result, op="list_roster")

    def list_assignments(self, org: str, classroom: str) -> Any:
        if not org or not classroom:
            raise Classroom50Error("org and classroom required", code="missing_params")
        result = self._run(["assignment", "list", org, classroom, "--json"])
        return self._parse_json(result, op="list_assignments")

    def download(self, *args, **kwargs):
        raise Classroom50Error(
            "download is not available on the agent adapter",
            code="agent_download_forbidden",
        )

    @staticmethod
    def _parse_json(result: RunResult, *, op: str) -> Any:
        if result.returncode != 0:
            raise Classroom50Error(
                redact_secrets(result.stderr or f"{op} failed"),
                code=f"{op}_failed",
            )
        raw = result.stdout or ""
        try:
            return json.loads(raw)
        except json.JSONDecodeError as exc:
            raise Classroom50Error(
                f"{op}: unparseable JSON", code="bad_json"
            ) from exc


def agent_cli(**kwargs) -> AgentCLI:
    return AgentCLI(**kwargs)
