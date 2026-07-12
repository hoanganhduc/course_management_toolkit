# -*- coding: utf-8 -*-
"""Human-only Classroom50 adapter methods (ADR v2.1 D1.B). Not imported by agent entry."""

from __future__ import annotations

from typing import Optional

from .c50_cli import Classroom50Error, RunResult, Runner, _default_runner, redact_secrets


class HumanCLI:
    """Closed map for human download only — never pass --by-pattern."""

    def __init__(
        self,
        *,
        runner: Optional[Runner] = None,
        gh_teacher: str = "gh",
    ):
        self._runner = runner or _default_runner
        self._gh = gh_teacher

    def download(
        self,
        org: str,
        classroom: str,
        assignment: str,
        dest: str,
    ) -> RunResult:
        if not org or not classroom or not assignment:
            raise Classroom50Error(
                "download requires org, classroom, and assignment",
                code="missing_download_params",
            )
        if not dest:
            raise Classroom50Error(
                "download requires destination directory",
                code="missing_dest",
            )
        argv = [
            self._gh,
            "teacher",
            "download",
            org,
            classroom,
            assignment,
            "-d",
            dest,
        ]
        # Never pass --by-pattern (ADR R2-m4)
        result = self._runner(argv)
        _ = redact_secrets(result.stderr)
        if result.returncode != 0:
            raise Classroom50Error(
                redact_secrets(result.stderr or "download failed"),
                code="download_failed",
            )
        return result
