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
import contextlib
import json
import shutil
import sys
import tempfile
import unicodedata
from pathlib import Path
from typing import Any, Callable, Dict, List, Mapping, Optional, Sequence, TextIO, Tuple

from .c50_cli import Classroom50Error, Runner
from .c50_cli_human import HumanCLI
from .c50_ops import (
    download_submissions,
    existing_member_logins,
    invite_users,
    is_agent_mode,
    preflight_assignment_add,
    preflight_assignment_remove,
)
from .c50_roster import (
    export_roster_csv,
    parse_csv_text,
    parse_roster_payload,
    partition_roster_candidates,
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
        "roster_import_failed",
        "roster_add_failed",
        "roster_list_failed",
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


_ROSTER_PREVIEW = 5
_CONFIRM_LIST_LIMIT = 10

# Fields this CLI actually writes. github_id is excluded on purpose: `gh teacher
# roster import` ignores the incoming value and re-resolves it from
# GET /users/{username}, so a difference there is not drift we can act on.
_ROSTER_COMPARED_FIELDS = ("first_name", "last_name", "email", "section")


def _student_label(student: Any) -> str:
    for key in ("Name", "Full Name", "Email", "Student ID"):
        value = getattr(student, key, None)
        if value and str(value).strip():
            return str(value).strip()
    return "(unnamed)"


def _roster_key(username: Any) -> str:
    return str(username or "").strip().lstrip("@").lower()


def _csv_rows_from_db(
    db_path: str, verbose: bool, diagnostics: TextIO
) -> Tuple[str, List[Dict[str, str]], List[str]]:
    """Render the roster CSV from the local database.

    ``load_database`` is imported here, not at module scope: it pulls pandas and
    friends, and this CLI must stay importable without them. It also prints
    progress on stdout, so its output is redirected onto the diagnostics stream:
    stdout carries the JSON report and nothing else.
    """
    from .data import load_database

    with contextlib.redirect_stdout(diagnostics):
        students = load_database(db_path, verbose=verbose)
    exportable, missing = partition_roster_candidates(students or [])
    csv_text, _ = export_roster_csv(exportable)
    return csv_text, parse_csv_text(csv_text), [_student_label(s) for s in missing]


def _csv_rows_from_file(csv_path: str) -> Tuple[str, List[Dict[str, str]]]:
    try:
        csv_text = Path(csv_path).read_text(encoding="utf-8")
    except OSError as exc:
        raise Classroom50Error(f"cannot read {csv_path}: {exc}", code="invalid_csv_file")
    return csv_text, parse_csv_text(csv_text)


def _roster_diff(
    rows: Sequence[Mapping[str, Any]], remote: Sequence[Mapping[str, Any]]
) -> Dict[str, List[str]]:
    """Classify local rows against the remote roster.

    ``roster import`` is an upsert: rows absent from the CSV stay on the roster
    untouched, so ``remoteOnly`` is reported, never removed.
    """
    remote_by_key = {_roster_key(r.get("username")): r for r in remote if r.get("username")}
    new: List[str] = []
    updated: List[str] = []
    unchanged: List[str] = []
    for row in rows:
        key = _roster_key(row.get("username"))
        if not key:
            continue
        current = remote_by_key.get(key)
        if current is None:
            new.append(row["username"])
        elif any(
            str(row.get(field) or "").strip() != str(current.get(field) or "").strip()
            for field in _ROSTER_COMPARED_FIELDS
        ):
            updated.append(row["username"])
        else:
            # Reported so a row needing no work is still visible. Without it an
            # unchanged student lands in no bucket, and a 41-row CSV that says
            # "40 to add" reads like a lost row rather than a repeated import.
            unchanged.append(row["username"])
    local_keys = {_roster_key(r.get("username")) for r in rows}
    remote_only = [
        r["username"] for key, r in remote_by_key.items() if key not in local_keys
    ]
    return {
        "new": new,
        "updated": updated,
        "unchanged": unchanged,
        "remoteOnly": sorted(remote_only),
    }


def _confirm_list(values: Sequence[str]) -> str:
    if not values:
        return "(none)"
    shown = [_terminal_text(v) for v in values[:_CONFIRM_LIST_LIMIT]]
    extra = len(values) - len(shown)
    return ", ".join(shown) + (f", ... and {extra} more" if extra else "")


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

    roster_import = sub.add_parser(
        "roster-import",
        help="Upsert the classroom roster from the local database or a CSV file",
        description=(
            "Bulk-upsert roster.csv. Rows already on the roster are updated in place "
            "and rows absent from the CSV are left untouched. After the commit lands, "
            "any listed student who is not already in the organization (and has no "
            "pending invitation) is invited."
        ),
    )
    roster_import.add_argument("--org", required=True)
    roster_import.add_argument("--classroom", required=True)
    roster_source = roster_import.add_mutually_exclusive_group(required=True)
    roster_source.add_argument(
        "--db", default=None, help="Local student database to render the CSV from"
    )
    roster_source.add_argument(
        "--csv",
        default=None,
        help="Existing CSV, e.g. the output of `course --export-classroom50-roster`",
    )
    roster_import.add_argument(
        "--keep-csv",
        action="store_true",
        help="Keep the generated CSV instead of deleting it (--db only)",
    )
    roster_import.add_argument(
        "--force",
        action="store_true",
        help=(
            "Import even when the roster already matches and everyone is already "
            "an organization member; use after a run that failed mid-way"
        ),
    )
    roster_import.add_argument("--verbose", action="store_true")
    roster_import.add_argument("--dry-run", action="store_true")

    roster_add = sub.add_parser(
        "roster-add",
        help="Add or correct one student on the classroom roster",
        description=(
            "Upsert a single roster row. Like roster-import, this invites the student "
            "to the organization when they are not already a member."
        ),
    )
    roster_add.add_argument("--org", required=True)
    roster_add.add_argument("--classroom", required=True)
    roster_add.add_argument("--username", required=True, help="GitHub login")
    roster_add.add_argument("--first-name", default=None)
    roster_add.add_argument("--last-name", default=None)
    roster_add.add_argument("--email", default=None)
    roster_add.add_argument("--section", default=None)
    roster_add.add_argument("--dry-run", action="store_true")

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
                    f"  Empty repo: {'yes' if args.empty_repo else 'no'}\n"
                    "  This adds the slug to assignments.json. Students can accept "
                    "the assignment as soon as the entry lands.",
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
                    "  This rewrites the entry in place.\n"
                    "  Warning: if this assignment was set to tagged-commit "
                    "submission mode in the web form, rewriting it resets the "
                    "submission mode to the default, which grades every push. The "
                    "pinned CLI cannot set submission mode, so restoring it means "
                    "going back to the web form.",
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
                "repository, and submission history stays intact; the only effect is "
                "that `gh student accept` stops finding this slug.\n"
                "  Re-adding the same slug later is not a clean reset: repositories "
                "already accepted keep the old empty-repository setting. To change "
                "that setting, add the assignment under a new slug.",
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
                "Invite users to a GitHub organization or repository\n"
                f"  Target: {_terminal_text(args.target)}\n"
                f"  Users:  {_terminal_text(', '.join(args.username))}\n"
                f"  Role:   {'org admin' if args.admin else (args.permission or 'default')}\n"
                "  Each user listed is sent an invitation. On an organization "
                "target, anyone already a member or already holding a pending "
                "invitation is skipped.",
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

        if args.command == "roster-import":
            verbose = bool(getattr(args, "verbose", False))
            if args.csv and args.keep_csv:
                # --csv generates no temporary file, so there is nothing to keep.
                # Silently ignoring the flag would hide a mistaken invocation.
                raise Classroom50Error(
                    "--keep-csv applies to --db only; the file passed to --csv is "
                    "never deleted",
                    code="invalid_flag_combination",
                )
            if args.csv:
                csv_text, rows = _csv_rows_from_file(args.csv)
                skipped_no_username: List[str] = []
                csv_source = "file"
            else:
                csv_text, rows, skipped_no_username = _csv_rows_from_db(
                    args.db, verbose, errors
                )
                csv_source = "generated-from-db"

            usernames = [row["username"] for row in rows]
            if args.dry_run:
                # The other verbs return before their `gh` preflight too, so --dry-run
                # stays a zero-call mode usable without credentials or a terminal. The
                # cost is that the roster diff cannot be shown here.
                plan: Dict[str, Any] = {
                    "csvSource": csv_source,
                    "rowCount": len(rows),
                    "firstUsernames": usernames[:_ROSTER_PREVIEW],
                    "skippedNoUsername": skipped_no_username,
                    "rosterDiff": "not computed: --dry-run makes no gh calls",
                }
                preconditions = [
                    "import upserts: rows already on the roster are updated, rows "
                    "absent from the CSV are left untouched",
                    "a listed student who is not already in the organization (and has "
                    "no pending invitation) is invited once the commit lands",
                ]
                if args.csv:
                    # Only the --csv form has a real path to print. Inventing one for
                    # a CSV that does not exist yet would not be truthful.
                    return _dry_run(
                        output,
                        args.command,
                        preconditions,
                        argv=cli.roster_import_argv(args.org, args.classroom, args.csv),
                        plan=plan,
                    )
                return _dry_run(output, args.command, preconditions, plan=plan)

            remote = parse_roster_payload(cli.roster_list(args.org, args.classroom))
            diff = _roster_diff(rows, remote)
            report: Dict[str, Any] = {
                "org": args.org,
                "classroom": args.classroom,
                "csvSource": csv_source,
                "rowCount": len(rows),
                "skippedNoUsername": skipped_no_username,
                "new": diff["new"],
                "updated": diff["updated"],
                "remoteOnly": diff["remoteOnly"],
            }
            not_yet_members: List[str] = []
            if not diff["new"] and not diff["updated"] and not args.force:
                # A matching roster is not proof that the work is done. `roster
                # import` commits the CSV first and invites afterwards, so a run
                # that died between the two leaves the roster matching and nobody
                # invited -- exactly the exit-3 case the operator is told to re-run.
                # This extra read only happens on the "nothing to do" path; when
                # there are rows to import the import itself does the inviting.
                try:
                    members = existing_member_logins(cli.member_list(args.org))
                except Classroom50Error as exc:
                    # Nothing was written, so this must not surface as exit 3.
                    report["action"] = "unchanged"
                    report["membershipNotVerified"] = True
                    print(
                        "nothing to import: every row already matches the roster",
                        file=errors,
                    )
                    print(
                        f"could not check organization membership ({exc}), so this run "
                        "cannot tell whether those students were ever invited; usually "
                        "the token lacks the admin:org scope. If an earlier run stopped "
                        "after writing the roster but before inviting, re-run with "
                        "--force to send the invitations.",
                        file=errors,
                    )
                    _write_json(output, report)
                    return 0
                not_yet_members = [u for u in usernames if u.lower() not in members]
                if not not_yet_members:
                    report["action"] = "unchanged"
                    print(
                        "nothing to import: every row already matches the roster, and "
                        "every student listed is already in the organization",
                        file=errors,
                    )
                    _write_json(output, report)
                    return 0
                report["notYetMembers"] = not_yet_members

            # A run reaching here with `not_yet_members` has nothing to add or
            # update, so it must not claim that it writes anybody to the roster.
            membership_line = ""
            if not_yet_members:
                membership_line = (
                    f"  Not in org yet:   {len(not_yet_members)} -- "
                    f"{_confirm_list(not_yet_members)}\n"
                )
                effect_line = (
                    "  Nothing on the roster changes. This run exists only to send "
                    "an organization invitation to each student listed as not in "
                    "the org yet.\n"
                )
            else:
                effect_line = (
                    "  The students under Adding and Updating are written to "
                    "roster.csv. Each of them who is not yet in the organization is "
                    "then sent an invitation.\n"
                )
            _confirm(
                input_fn,
                "Import a Classroom50 roster\n"
                f"  Org:              {_terminal_text(args.org)}\n"
                f"  Classroom:        {_terminal_text(args.classroom)}\n"
                f"  Adding:           {len(diff['new'])} -- "
                f"{_confirm_list(diff['new'])}  (not on the roster yet)\n"
                f"  Updating:         {len(diff['updated'])} -- "
                f"{_confirm_list(diff['updated'])}  (name, email or section differs)\n"
                f"  Already correct:  {len(diff['unchanged'])} -- "
                f"{_confirm_list(diff['unchanged'])}  (on the roster, nothing to change)\n"
                f"  Cannot import:    {len(skipped_no_username)} -- "
                f"{_confirm_list(skipped_no_username)}  (the CSV row has no username)\n"
                f"  Leaving alone:    {len(diff['remoteOnly'])} -- "
                f"{_confirm_list(diff['remoteOnly'])}  (on the roster, not in this CSV)\n"
                f"{membership_line}"
                f"{effect_line}"
                "  An import never removes anyone from the roster.",
            )

            workdir: Optional[str] = None
            if args.csv:
                csv_path = args.csv
            else:
                workdir = tempfile.mkdtemp(prefix="c50-roster-")
                csv_path = str(Path(workdir) / "roster.csv")
                # No BOM: ParseImportCSV wants a bare `username,...` header.
                Path(csv_path).write_text(csv_text, encoding="utf-8")
            try:
                result = cli.roster_import(args.org, args.classroom, csv_path)
            finally:
                if workdir and not args.keep_csv:
                    # Delete on failure too: the CSV is reproducible from the database,
                    # and student emails should not be left behind in a temp directory.
                    shutil.rmtree(workdir, ignore_errors=True)
            if workdir and args.keep_csv:
                report["csvPath"] = csv_path
            report["action"] = "imported"
            stderr_text = getattr(result, "stderr", "") or ""
            if stderr_text.strip():
                # The CLI warns about partial onboarding here. Passed through verbatim
                # rather than parsed, because the wording is version-specific.
                report["ghStderr"] = _terminal_text(stderr_text.strip())
            _write_json(output, report)
            return 0

        if args.command == "roster-add":
            argv_out = cli.roster_add_argv(
                args.org,
                args.classroom,
                args.username,
                first_name=args.first_name,
                last_name=args.last_name,
                email=args.email,
                section=args.section,
            )
            if args.dry_run:
                return _dry_run(
                    output,
                    args.command,
                    [
                        "an existing row for this username is updated in place",
                        "the student is invited to the organization unless they are "
                        "already a member or hold a pending invitation",
                    ],
                    argv=argv_out,
                )
            _confirm(
                input_fn,
                "Add one student to a Classroom50 roster\n"
                f"  Org:       {_terminal_text(args.org)}\n"
                f"  Classroom: {_terminal_text(args.classroom)}\n"
                f"  Username:  {_terminal_text(args.username)}\n"
                "  This writes one row to roster.csv, then sends an organization "
                "invitation if the student is not already a member.",
            )
            cli.roster_add(
                args.org,
                args.classroom,
                args.username,
                first_name=args.first_name,
                last_name=args.last_name,
                email=args.email,
                section=args.section,
            )
            _write_json(
                output,
                {
                    "org": args.org,
                    "classroom": args.classroom,
                    "username": args.username,
                    "action": "upserted",
                },
            )
            return 0

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
