# -*- coding: utf-8 -*-
"""Human-only Classroom50 adapter methods (ADR v2.1 D1.B). Not imported by agent entry."""

from __future__ import annotations

import json
import os
import re
from typing import Any, List, Optional

from .c50_cli import Classroom50Error, RunResult, Runner, _default_runner, redact_secrets

# Slug shape enforced by `gh teacher assignment add` v1.25.1.
_SLUG_RE = re.compile(r"^[a-z0-9][a-z0-9-]{1,38}$")
_TARGET_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*(/[A-Za-z0-9][A-Za-z0-9._-]*)?$")
_USERNAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9-]{0,38}$")
_PERMISSIONS = ("pull", "triage", "push", "maintain", "admin")
_MODES = ("individual", "group")


def _operand(value: Any, label: str, *, code: str = "missing_params") -> str:
    """Reject empty and flag-like operands so no value can become a `gh` flag."""
    text = "" if value is None else str(value)
    if not text.strip():
        raise Classroom50Error(f"{label} required", code=code)
    if text.startswith("-"):
        raise Classroom50Error(
            f"{label} must not begin with '-'", code="flaglike_operand"
        )
    return text


def _tests_file(path: str) -> str:
    """Validate the `--tests` payload the pinned CLI expects: a readable JSON file
    holding a bare array of test specs. The CLI also accepts `-` for stdin, which this
    adapter cannot offer because every operand is rejected when it begins with '-'."""
    if not os.path.isfile(path):
        raise Classroom50Error(
            f"tests file not found: {path}", code="invalid_tests_file"
        )
    try:
        with open(path, "r", encoding="utf-8") as handle:
            parsed = json.load(handle)
    except (OSError, ValueError) as exc:
        raise Classroom50Error(
            f"tests file is not readable JSON: {exc}", code="invalid_tests_file"
        )
    if not isinstance(parsed, list):
        raise Classroom50Error(
            "tests file must hold a bare JSON array of test specs",
            code="invalid_tests_file",
        )
    return path


def validate_target(target: str) -> str:
    """Accept `<org>` or `<org>/<repo>` only."""
    target = _operand(target, "target")
    if not _TARGET_RE.match(target):
        raise Classroom50Error(
            f"target {target!r} must be <org> or <org>/<repo>", code="invalid_target"
        )
    return target


class HumanCLI:
    """Closed map of reviewed teacher-side commands. One fixed argv per method."""

    def __init__(
        self,
        *,
        runner: Optional[Runner] = None,
        gh_teacher: str = "gh",
    ):
        self._runner = runner or _default_runner
        self._gh = gh_teacher

    # -- argv builders (pure; used by --dry-run and by the methods below) ------

    def assignment_add_argv(
        self,
        org: str,
        classroom: str,
        slug: str,
        *,
        name: str,
        description: Optional[str] = None,
        template: Optional[str] = None,
        mode: Optional[str] = None,
        max_group_size: Optional[int] = None,
        available_from: Optional[str] = None,
        due: Optional[str] = None,
        empty_repo: bool = False,
        tests: Optional[str] = None,
        feedback_pr: Optional[bool] = None,
        pass_threshold: Optional[int] = None,
        allowed_files: Optional[List[str]] = None,
        student_permission: Optional[str] = None,
    ) -> List[str]:
        org = _operand(org, "org")
        classroom = _operand(classroom, "classroom")
        slug = _operand(slug, "slug")
        if not _SLUG_RE.match(slug):
            raise Classroom50Error(
                f"slug {slug!r} must match ^[a-z0-9][a-z0-9-]{{1,38}}$",
                code="invalid_slug",
            )
        name = _operand(name, "name", code="missing_name")
        if mode is not None and mode not in _MODES:
            raise Classroom50Error(
                f"mode must be one of {', '.join(_MODES)}", code="invalid_mode"
            )
        if mode == "group":
            if max_group_size is None:
                raise Classroom50Error(
                    "group mode requires --max-group-size", code="missing_max_group_size"
                )
            if int(max_group_size) < 2:
                raise Classroom50Error(
                    "--max-group-size must be at least 2", code="invalid_max_group_size"
                )
        elif max_group_size is not None:
            raise Classroom50Error(
                "--max-group-size applies only to group mode",
                code="max_group_size_without_group",
            )

        # v1.25.1 rejects these five alongside --empty-repo, and the setting is
        # immutable after creation, so catch the combination before the network call.
        conflicting = [
            flag
            for flag, given in (
                ("--template", bool(template)),
                ("--tests", tests is not None),
                ("--feedback-pr", feedback_pr is not None),
                ("--allowed-files", bool(allowed_files)),
                ("--pass-threshold", pass_threshold is not None),
            )
            if given
        ]
        if empty_repo and conflicting:
            raise Classroom50Error(
                "--empty-repo is mutually exclusive with " + ", ".join(conflicting),
                code="mutually_exclusive",
            )
        if tests is not None:
            tests = _tests_file(_operand(tests, "tests"))
        if pass_threshold is not None:
            try:
                pass_threshold = int(pass_threshold)
            except (TypeError, ValueError):
                raise Classroom50Error(
                    "--pass-threshold must be an integer percentage",
                    code="invalid_pass_threshold",
                )
            if not 0 <= pass_threshold <= 100:
                raise Classroom50Error(
                    "--pass-threshold must be between 0 and 100",
                    code="invalid_pass_threshold",
                )
        if student_permission is not None and student_permission not in _PERMISSIONS:
            raise Classroom50Error(
                f"--student-permission must be one of {', '.join(_PERMISSIONS)}",
                code="invalid_permission",
            )
        patterns = [
            _operand(pattern, "allowed-files") for pattern in (allowed_files or [])
        ]

        argv = [
            self._gh,
            "teacher",
            "assignment",
            "add",
            org,
            classroom,
            slug,
            "--name",
            name,
        ]
        if description:
            argv += ["--description", _operand(description, "description")]
        if template:
            argv += ["--template", _operand(template, "template")]
        if tests is not None:
            argv += ["--tests", tests]
        if mode:
            argv += ["--mode", mode]
        if max_group_size is not None:
            argv += ["--max-group-size", str(int(max_group_size))]
        if available_from:
            argv += ["--available-from", _operand(available_from, "available-from")]
        if due:
            argv += ["--due", _operand(due, "due")]
        if feedback_pr is not None:
            # Default is true in v1.25.1; disabling needs the `=false` form.
            argv.append("--feedback-pr" if feedback_pr else "--feedback-pr=false")
        if pass_threshold is not None:
            argv += ["--pass-threshold", str(pass_threshold)]
        for pattern in patterns:
            argv += ["--allowed-files", pattern]
        if student_permission:
            argv += ["--student-permission", student_permission]
        if empty_repo:
            argv.append("--empty-repo")
        return argv

    def assignment_remove_argv(self, org: str, classroom: str, slug: str) -> List[str]:
        return [
            self._gh,
            "teacher",
            "assignment",
            "remove",
            _operand(org, "org"),
            _operand(classroom, "classroom"),
            _operand(slug, "slug"),
        ]

    def assignment_list_argv(self, org: str, classroom: str) -> List[str]:
        return [
            self._gh,
            "teacher",
            "assignment",
            "list",
            _operand(org, "org"),
            _operand(classroom, "classroom"),
            "--json",
        ]

    def member_list_argv(self, target: str) -> List[str]:
        return [
            self._gh,
            "teacher",
            "member",
            "list",
            self._target(target),
            "--json",
        ]

    def invite_argv(
        self,
        target: str,
        username: str,
        *,
        admin: bool = False,
        permission: Optional[str] = None,
    ) -> List[str]:
        target = self._target(target)
        username = _operand(username, "username")
        if not _USERNAME_RE.match(username):
            raise Classroom50Error(
                f"username {username!r} is not a GitHub login", code="invalid_username"
            )
        is_repo = "/" in target
        if admin and is_repo:
            raise Classroom50Error(
                "--admin applies only to organization targets", code="admin_requires_org"
            )
        if permission is not None:
            if not is_repo:
                raise Classroom50Error(
                    "--permission applies only to <org>/<repo> targets",
                    code="permission_requires_repo",
                )
            if permission not in _PERMISSIONS:
                raise Classroom50Error(
                    f"permission must be one of {', '.join(_PERMISSIONS)}",
                    code="invalid_permission",
                )

        argv = [self._gh, "teacher", "invite"]
        if admin:
            argv.append("--admin")
        if permission is not None:
            argv += ["-p", permission]
        argv += [target, username]
        return argv

    def download_argv(
        self,
        org: str,
        classroom: str,
        assignment: str,
        dest: str,
        *,
        by_pattern: bool = False,
    ) -> List[str]:
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
        argv = [self._gh, "teacher", "download"]
        if by_pattern:
            argv.append("--by-pattern")
        argv += [
            _operand(org, "org"),
            _operand(classroom, "classroom"),
            _operand(assignment, "assignment"),
            "-d",
            _operand(dest, "dest"),
        ]
        return argv

    # -- commands -------------------------------------------------------------

    def list_assignments(self, org: str, classroom: str) -> Any:
        return self._json(self.assignment_list_argv(org, classroom), op="list_assignments")

    def member_list(self, target: str) -> Any:
        return self._json(self.member_list_argv(target), op="member_list")

    def assignment_add(self, org: str, classroom: str, slug: str, **kwargs) -> RunResult:
        argv = self.assignment_add_argv(org, classroom, slug, **kwargs)
        return self._checked(argv, code="assignment_add_failed")

    def assignment_remove(self, org: str, classroom: str, slug: str) -> RunResult:
        argv = self.assignment_remove_argv(org, classroom, slug)
        return self._checked(argv, code="assignment_remove_failed")

    def invite(self, target: str, username: str, **kwargs) -> RunResult:
        argv = self.invite_argv(target, username, **kwargs)
        return self._checked(argv, code="invite_failed")

    def download(
        self,
        org: str,
        classroom: str,
        assignment: str,
        dest: str,
        *,
        by_pattern: bool = False,
        empty_repo_verified: bool = False,
    ) -> RunResult:
        # ADR R2-m4, narrowed: --by-pattern skips the team lookup, so it collects no
        # result.json and writes no scores.csv. That is wrong for an autograded
        # assignment and required for an empty-repository one, which has neither to
        # collect. The caller must have read the assignment record and confirmed
        # empty_repo; c50_ops.download_submissions is the checked entry point.
        if by_pattern and not empty_repo_verified:
            raise Classroom50Error(
                "--by-pattern requires a verified empty-repository assignment; "
                "use c50_ops.download_submissions",
                code="by_pattern_not_permitted",
            )
        argv = self.download_argv(org, classroom, assignment, dest, by_pattern=by_pattern)
        return self._checked(argv, code="download_failed")

    # -- internals ------------------------------------------------------------

    def _target(self, target: str) -> str:
        return validate_target(target)

    def _checked(self, argv: List[str], *, code: str) -> RunResult:
        result = self._runner(argv)
        if result.returncode != 0:
            raise Classroom50Error(
                redact_secrets(result.stderr or f"{argv[2]} failed"), code=code
            )
        return result

    def _json(self, argv: List[str], *, op: str) -> Any:
        result = self._runner(argv)
        if result.returncode != 0:
            raise Classroom50Error(
                redact_secrets(result.stderr or f"{op} failed"), code=f"{op}_failed"
            )
        try:
            return json.loads(result.stdout or "")
        except json.JSONDecodeError as exc:
            raise Classroom50Error(f"{op}: unparseable JSON", code="bad_json") from exc
