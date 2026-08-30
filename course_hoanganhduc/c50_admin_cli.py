# -*- coding: utf-8 -*-
"""Human-only Classroom50 operator CLI (ADR v2.1 D1.B / D6).

Mutating `gh teacher` operations live here, never on the agent entrypoint. Every
verb refuses in agent mode, requires an interactive terminal unless ``--dry-run``,
and builds a fixed argv through :class:`course_hoanganhduc.c50_cli_human.HumanCLI`.

Remote writes remain gated by the course runbook. Reaching an operation through
this CLI does not authorize running it.
"""

from __future__ import annotations

import argparse
import json
import sys
import unicodedata
from typing import Any, Callable, Dict, List, Mapping, Optional, Sequence, TextIO

from .c50_cli import Classroom50Error, Runner
from .c50_cli_human import HumanCLI
from .c50_ops import (
    download_submissions,
    invite_users,
    is_agent_mode,
    preflight_assignment_add,
    preflight_assignment_remove,
)

# A failed `gh` invocation leaves the remote state unverified.
_OUTCOME_UNKNOWN_CODES = frozenset(
    {
        "assignment_add_failed",
        "assignment_remove_failed",
        "invite_failed",
        "download_failed",
        "list_assignments_failed",
        "member_list_failed",
        "missing_binary",
    }
)


def _write_json(stream: TextIO, value: Mapping[str, Any]) -> None:
    stream.write(json.dumps(value, indent=2, ensure_ascii=False, sort_keys=True) + "\n")


def _terminal_text(value: Any) -> str:
    """Escape terminal control and formatting characters in CLI-derived text."""
    out = []
    for character in str(value):
        if unicodedata.category(character) in {"Cc", "Cf", "Cs", "Zl", "Zp"}:
            out.append(f"\\u{ord(character):04x}")
        else:
            out.append(character)
    return "".join(out)


def _default_tty_check() -> bool:
    try:
        return bool(sys.stdin.isatty())
    except Exception:
        return False


def _confirm(input_fn: Callable[[str], str], prompt: str) -> None:
    try:
        received = input_fn(f"{prompt}\nProceed? [y/N]: ")
    except (EOFError, StopIteration):
        raise Classroom50Error("operation cancelled", code="cancelled") from None
    if not isinstance(received, str) or received.strip().lower() not in {"y", "yes"}:
        raise Classroom50Error("operation cancelled", code="cancelled")


def _require_human(*, dry_run: bool, tty_check: Callable[[], bool]) -> None:
    if is_agent_mode():
        raise Classroom50Error(
            "Classroom50 mutation and download operations are forbidden in agent mode",
            code="agent_forbidden",
        )
    if dry_run:
        return
    if not tty_check():
        raise Classroom50Error(
            "this operation requires an interactive terminal; use --dry-run to print "
            "the command instead",
            code="not_interactive",
        )


def _add_kwargs(args: argparse.Namespace) -> Dict[str, Any]:
    return {
        "name": args.name,
        "description": args.description,
        "template": args.template,
        "mode": args.mode,
        "max_group_size": args.max_group_size,
        "available_from": args.available_from,
        "due": args.due,
        "empty_repo": args.empty_repo,
        "tests": args.tests,
        "feedback_pr": args.feedback_pr,
        "pass_threshold": args.pass_threshold,
        "allowed_files": args.allowed_files,
        "student_permission": args.student_permission,
    }


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="course-c50-admin",
        description=(
            "Human-only Classroom50 teacher operations. Remote writes still require "
            "the authorization gates in the course runbook."
        ),
        allow_abbrev=False,
    )
    sub = parser.add_subparsers(dest="command", required=True)

    add = sub.add_parser(
        "assignment-add", help="Register or replace an assignment entry"
    )
    add.add_argument("--org", required=True)
    add.add_argument("--classroom", required=True)
    add.add_argument("--slug", required=True)
    add.add_argument("--name", required=True, help="Display name for the assignment")
    add.add_argument("--description", default=None)
    add.add_argument("--template", default=None, help="<owner>/<repo>[@<branch>]")
    add.add_argument("--mode", default=None, choices=["individual", "group"])
    add.add_argument("--max-group-size", type=int, default=None)
    add.add_argument("--available-from", default=None)
    add.add_argument("--due", default=None)
    add.add_argument("--empty-repo", action="store_true")
    add.add_argument(
        "--tests",
        default=None,
        help="Path to a JSON file holding a bare array of declarative test specs",
    )
    feedback = add.add_mutually_exclusive_group()
    feedback.add_argument(
        "--feedback-pr",
        dest="feedback_pr",
        action="store_true",
        default=None,
        help="Open the long-lived Feedback pull request (the pinned CLI default)",
    )
    feedback.add_argument(
        "--no-feedback-pr", dest="feedback_pr", action="store_false"
    )
    add.add_argument(
        "--pass-threshold",
        type=int,
        default=None,
        help="Advisory passing bar as a percentage (0-100)",
    )
    add.add_argument(
        "--allowed-files",
        action="append",
        default=None,
        metavar="PATTERN",
        help=(
            "Ordered .gitignore-style pattern; repeatable, order preserved, "
            "'!' re-includes"
        ),
    )
    add.add_argument(
        "--student-permission",
        default=None,
        choices=["pull", "triage", "push", "maintain", "admin"],
        help="Collaborator role each student gets on their own repository",
    )
    add.add_argument(
        "--allow-overwrite",
        action="store_true",
        help=(
            "Permit replacing an existing slug. The pinned CLI cannot set submission "
            "mode, so this restores the default every-push mode."
        ),
    )
    add.add_argument("--dry-run", action="store_true")

    remove = sub.add_parser(
        "assignment-remove", help="Drop an assignment entry from the manifest"
    )
    remove.add_argument("--org", required=True)
    remove.add_argument("--classroom", required=True)
    remove.add_argument("--slug", required=True)
    remove.add_argument("--dry-run", action="store_true")

    invite = sub.add_parser("invite", help="Invite users to an organization or repository")
    invite.add_argument("--target", required=True, help="<org> or <org>/<repo>")
    invite.add_argument(
        "--username", action="append", required=True, help="Repeatable GitHub login"
    )
    invite.add_argument("--admin", action="store_true", help="Organization targets only")
    invite.add_argument(
        "--permission",
        default=None,
        choices=["pull", "triage", "push", "maintain", "admin"],
        help="Repository targets only",
    )
    invite.add_argument("--dry-run", action="store_true")

    download = sub.add_parser("download", help="Clone student submission repositories")
    download.add_argument("--org", required=True)
    download.add_argument("--classroom", required=True)
    download.add_argument("--assignment", required=True)
    download.add_argument("--dest", required=True)
    download.add_argument(
        "--by-pattern",
        action="store_true",
        help=(
            "Clone every matching repository without the team lookup. Permitted only "
            "for empty-repository assignments, which have no result.json to collect."
        ),
    )
    download.add_argument("--dry-run", action="store_true")

    return parser


def _dry_run(
    output: TextIO, command: str, preconditions: Sequence[str], **payload: Any
) -> int:
    body: Dict[str, Any] = {
        "dryRun": True,
        "command": command,
        "preconditions": list(preconditions),
    }
    body.update(payload)
    _write_json(output, body)
    return 0


def main(
    argv: Optional[Sequence[str]] = None,
    *,
    runner: Optional[Runner] = None,
    input_fn: Callable[[str], str] = input,
    tty_check: Optional[Callable[[], bool]] = None,
    stdout: Optional[TextIO] = None,
    stderr: Optional[TextIO] = None,
) -> int:
    """Run the operator CLI; dependency injection keeps the tests offline."""

    output = stdout or sys.stdout
    errors = stderr or sys.stderr
    check_tty = tty_check or _default_tty_check
    args = _build_parser().parse_args(argv)
    cli = HumanCLI(runner=runner)

    try:
        _require_human(dry_run=args.dry_run, tty_check=check_tty)

        if args.command == "assignment-add":
            argv_out = cli.assignment_add_argv(
                args.org, args.classroom, args.slug, **_add_kwargs(args)
            )
            if args.dry_run:
                return _dry_run(
                    output,
                    args.command,
                    [
                        "the slug must be absent from assignments.json, "
                        "unless --allow-overwrite is given",
                    ],
                    argv=argv_out,
                )
            existing = preflight_assignment_add(
                args.org,
                args.classroom,
                args.slug,
                allow_overwrite=args.allow_overwrite,
                cli=cli,
            )
            if existing is None:
                _confirm(
                    input_fn,
                    "Register a new Classroom50 assignment\n"
                    f"  Org:        {_terminal_text(args.org)}\n"
                    f"  Classroom:  {_terminal_text(args.classroom)}\n"
                    f"  Slug:       {_terminal_text(args.slug)}\n"
                    f"  Name:       {_terminal_text(args.name)}\n"
                    f"  Empty repo: {'yes' if args.empty_repo else 'no'}",
                )
            else:
                _confirm(
                    input_fn,
                    "Replace an existing Classroom50 assignment entry\n"
                    f"  Org:        {_terminal_text(args.org)}\n"
                    f"  Classroom:  {_terminal_text(args.classroom)}\n"
                    f"  Slug:       {_terminal_text(args.slug)}\n"
                    f"  Current:    {_terminal_text(existing.get('name') or '(unnamed)')}\n"
                    f"  New name:   {_terminal_text(args.name)}\n"
                    "  This rewrites the entry in place. The pinned CLI cannot set "
                    "submission mode, so a tagged-commit submission mode set in the "
                    "web form is restored to the default every-push mode and cannot "
                    "be put back from the CLI.",
                )
            cli.assignment_add(args.org, args.classroom, args.slug, **_add_kwargs(args))
            _write_json(
                output,
                {
                    "org": args.org,
                    "classroom": args.classroom,
                    "slug": args.slug,
                    "action": "replaced" if existing else "registered",
                },
            )
            return 0

        if args.command == "assignment-remove":
            argv_out = cli.assignment_remove_argv(args.org, args.classroom, args.slug)
            if args.dry_run:
                return _dry_run(
                    output,
                    args.command,
                    ["the slug must be present in assignments.json"],
                    argv=argv_out,
                )
            record = preflight_assignment_remove(
                args.org, args.classroom, args.slug, cli=cli
            )
            _confirm(
                input_fn,
                "Remove a Classroom50 assignment entry\n"
                f"  Org:       {_terminal_text(args.org)}\n"
                f"  Classroom: {_terminal_text(args.classroom)}\n"
                f"  Slug:      {_terminal_text(args.slug)}\n"
                f"  Name:      {_terminal_text(record.get('name') or '(unnamed)')}\n"
                "  This edits assignments.json only. It does not delete any student "
                "repository, and submission history stays intact; only new "
                "`gh student accept` calls stop finding the slug.\n"
                "  Re-adding the same slug afterwards is not a clean reset: already "
                "accepted repositories keep the old empty-repository behaviour. To "
                "change that setting, add under a new slug.",
            )
            cli.assignment_remove(args.org, args.classroom, args.slug)
            _write_json(
                output,
                {
                    "org": args.org,
                    "classroom": args.classroom,
                    "slug": args.slug,
                    "action": "removed",
                },
            )
            return 0

        if args.command == "invite":
            argvs: List[List[str]] = [
                cli.invite_argv(
                    args.target, username, admin=args.admin, permission=args.permission
                )
                for username in args.username
            ]
            if args.dry_run:
                return _dry_run(
                    output,
                    args.command,
                    [
                        "organization targets: logins already members or holding a "
                        "pending invitation are skipped",
                        "repository targets: invitations are idempotent and are sent "
                        "directly",
                    ],
                    argvs=argvs,
                )
            _confirm(
                input_fn,
                "Invite users to a GitHub target\n"
                f"  Target: {_terminal_text(args.target)}\n"
                f"  Users:  {_terminal_text(', '.join(args.username))}\n"
                f"  Role:   {'org admin' if args.admin else (args.permission or 'default')}",
            )
            report = invite_users(
                args.target,
                args.username,
                admin=args.admin,
                permission=args.permission,
                cli=cli,
            )
            _write_json(output, report)
            return 4 if report["failed"] else 0

        if args.command == "download":
            argv_out = cli.download_argv(
                args.org,
                args.classroom,
                args.assignment,
                args.dest,
                by_pattern=args.by_pattern,
            )
            if args.dry_run:
                return _dry_run(
                    output,
                    args.command,
                    [
                        "--by-pattern requires the assignment to be registered with "
                        "empty_repo true",
                    ]
                    if args.by_pattern
                    else [],
                    argv=argv_out,
                )
            download_submissions(
                args.org,
                args.classroom,
                args.assignment,
                args.dest,
                by_pattern=args.by_pattern,
                cli=cli,
            )
            _write_json(
                output,
                {
                    "org": args.org,
                    "classroom": args.classroom,
                    "assignment": args.assignment,
                    "dest": args.dest,
                    "byPattern": bool(args.by_pattern),
                    "action": "downloaded",
                },
            )
            return 0

    except Classroom50Error as exc:
        _write_json(
            errors,
            {
                "error": (
                    "outcome_unknown"
                    if exc.code in _OUTCOME_UNKNOWN_CODES
                    else "validation"
                ),
                "code": exc.code,
                "message": str(exc),
            },
        )
        return 3 if exc.code in _OUTCOME_UNKNOWN_CODES else 2
    except ValueError as exc:
        _write_json(errors, {"error": "validation", "code": "value_error", "message": str(exc)})
        return 2
    except Exception:
        _write_json(
            errors,
            {
                "error": "internal_error",
                "code": "internal_error",
                "message": "Unexpected internal failure; no automatic retry was attempted.",
            },
        )
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


__all__ = ["main"]
