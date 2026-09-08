#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Offline tests for the Classroom50 operator surface (assignment/invite/download)."""

from __future__ import annotations

import contextlib
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import types
import unittest
from pathlib import Path

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc import c50_admin_cli
from course_hoanganhduc.c50_roster import CANONICAL_COLUMNS
from course_hoanganhduc.models import Student
from course_hoanganhduc.c50_cli import Classroom50Error, RunResult
from course_hoanganhduc.c50_cli_human import HumanCLI
from course_hoanganhduc.c50_ops import (
    assignment_allows_by_pattern,
    download_submissions,
    existing_member_logins,
    find_assignment,
    invite_users,
    preflight_assignment_add,
    preflight_assignment_remove,
)

MANIFEST = [
    {"slug": "week-0a", "name": "Week 0A", "empty_repo": False, "mode": "individual"},
    {"slug": "final-project", "name": "Final", "empty_repo": True, "mode": "group"},
]
MEMBERS = [
    {"login": "Alice", "kind": "member", "role": "admin", "github_id": 1},
    {"login": "bob", "kind": "invitation", "role": "direct_member", "github_id": 2},
]


def _runner(mapping, calls=None):
    """Injected runner: exact argv tuple -> RunResult; records every call."""

    def run(argv):
        if calls is not None:
            calls.append(list(argv))
        return mapping.get(
            tuple(argv), RunResult(returncode=1, stderr=f"unexpected argv: {argv}")
        )

    return run


def _list_call(org, classroom):
    return ("gh", "teacher", "assignment", "list", org, classroom, "--json")


def _member_call(target):
    return ("gh", "teacher", "member", "list", target, "--json")


_CSV_HEADER = ",".join(CANONICAL_COLUMNS) + "\n"
_CSV_TEXT = _CSV_HEADER + "alice,Alice,Andersson,alice@x.edu,S1,\n"


def _roster_list_call(org, classroom):
    return ("gh", "teacher", "roster", "list", org, classroom, "--json")


def _roster_import_call(org, classroom, path):
    return ("gh", "teacher", "roster", "import", org, classroom, path)


class TestPreflightHelpers(unittest.TestCase):
    def test_find_assignment_hit_and_miss(self):
        self.assertEqual(find_assignment(MANIFEST, "week-0a")["name"], "Week 0A")
        self.assertIsNone(find_assignment(MANIFEST, "week-9z"))

    def test_find_assignment_rejects_non_array(self):
        with self.assertRaises(Classroom50Error) as ctx:
            find_assignment({"slug": "x"}, "x")
        self.assertEqual(ctx.exception.code, "bad_manifest")

    def test_by_pattern_allowed_only_for_empty_repo(self):
        self.assertTrue(assignment_allows_by_pattern(MANIFEST[1]))
        self.assertFalse(assignment_allows_by_pattern(MANIFEST[0]))
        self.assertFalse(assignment_allows_by_pattern(None))
        # a truthy non-True value must not open the gate
        self.assertFalse(assignment_allows_by_pattern({"empty_repo": "yes"}))

    def test_existing_member_logins_lowercases_members_and_invitations(self):
        self.assertEqual(existing_member_logins(MEMBERS), {"alice", "bob"})

    def test_existing_member_logins_rejects_non_array(self):
        with self.assertRaises(Classroom50Error) as ctx:
            existing_member_logins("nope")
        self.assertEqual(ctx.exception.code, "bad_member_list")


class TestArgvBuilders(unittest.TestCase):
    def setUp(self):
        self.cli = HumanCLI(runner=lambda a: RunResult(0))

    def test_assignment_add_argv_matches_reviewed_command(self):
        argv = self.cli.assignment_add_argv(
            "VNU-HUS",
            "introai",
            "final-project",
            name="Final Examination Mini-Project",
            mode="group",
            max_group_size=5,
            available_from="2026-09-11T13:00:00+07:00",
            due="2026-11-04T23:59:00+07:00",
            empty_repo=True,
        )
        self.assertEqual(
            argv,
            [
                "gh", "teacher", "assignment", "add",
                "VNU-HUS", "introai", "final-project",
                "--name", "Final Examination Mini-Project",
                "--mode", "group",
                "--max-group-size", "5",
                "--available-from", "2026-09-11T13:00:00+07:00",
                "--due", "2026-11-04T23:59:00+07:00",
                "--empty-repo",
            ],
        )

    def test_assignment_add_rejects_bad_slug(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.assignment_add_argv("O", "C", "Week_0A", name="x")
        self.assertEqual(ctx.exception.code, "invalid_slug")

    def test_assignment_add_rejects_empty_repo_with_template(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.assignment_add_argv(
                "O", "C", "w0", name="x", empty_repo=True, template="o/t"
            )
        self.assertEqual(ctx.exception.code, "mutually_exclusive")

    def test_assignment_add_group_requires_max_group_size(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.assignment_add_argv("O", "C", "w0", name="x", mode="group")
        self.assertEqual(ctx.exception.code, "missing_max_group_size")

    def test_assignment_add_rejects_max_group_size_below_two(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.assignment_add_argv(
                "O", "C", "w0", name="x", mode="group", max_group_size=1
            )
        self.assertEqual(ctx.exception.code, "invalid_max_group_size")

    def test_assignment_add_rejects_max_group_size_without_group(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.assignment_add_argv("O", "C", "w0", name="x", max_group_size=5)
        self.assertEqual(ctx.exception.code, "max_group_size_without_group")

    def test_flaglike_operand_never_reaches_argv(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.assignment_remove_argv("--force", "C", "w0")
        self.assertEqual(ctx.exception.code, "flaglike_operand")

    def test_assignment_remove_argv(self):
        self.assertEqual(
            self.cli.assignment_remove_argv("VNU-HUS", "pilot", "tmp-slug"),
            ["gh", "teacher", "assignment", "remove", "VNU-HUS", "pilot", "tmp-slug"],
        )

    def test_invite_argv_org_and_repo_forms(self):
        self.assertEqual(
            self.cli.invite_argv("VNU-HUS", "alice"),
            ["gh", "teacher", "invite", "VNU-HUS", "alice"],
        )
        self.assertEqual(
            self.cli.invite_argv("VNU-HUS", "alice", admin=True),
            ["gh", "teacher", "invite", "--admin", "VNU-HUS", "alice"],
        )
        self.assertEqual(
            self.cli.invite_argv("VNU-HUS/introai-w0", "alice", permission="maintain"),
            ["gh", "teacher", "invite", "-p", "maintain", "VNU-HUS/introai-w0", "alice"],
        )

    def test_invite_rejects_admin_on_repo_target(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.invite_argv("O/r", "alice", admin=True)
        self.assertEqual(ctx.exception.code, "admin_requires_org")

    def test_invite_rejects_permission_on_org_target(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.invite_argv("O", "alice", permission="push")
        self.assertEqual(ctx.exception.code, "permission_requires_repo")

    def test_invite_rejects_unknown_permission(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.invite_argv("O/r", "alice", permission="owner")
        self.assertEqual(ctx.exception.code, "invalid_permission")

    def test_invite_rejects_bad_username(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.invite_argv("O", "not a login")
        self.assertEqual(ctx.exception.code, "invalid_username")

    def test_download_argv_with_and_without_pattern(self):
        self.assertEqual(
            self.cli.download_argv("O", "C", "a", "dest"),
            ["gh", "teacher", "download", "O", "C", "a", "-d", "dest"],
        )
        self.assertEqual(
            self.cli.download_argv("O", "C", "a", "dest", by_pattern=True),
            ["gh", "teacher", "download", "--by-pattern", "O", "C", "a", "-d", "dest"],
        )

    def test_download_requires_two_keys_for_by_pattern(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.download("O", "C", "a", "dest", by_pattern=True)
        self.assertEqual(ctx.exception.code, "by_pattern_not_permitted")

    def test_download_redacts_secret_in_failure(self):
        token = "ghp_abcdefghijklmnopqrstuvwxyz0123456789"
        run = _runner(
            {
                ("gh", "teacher", "download", "O", "C", "a", "-d", "d"): RunResult(
                    1, "", f"failed with {token}"
                )
            }
        )
        with self.assertRaises(Classroom50Error) as ctx:
            HumanCLI(runner=run).download("O", "C", "a", "d")
        self.assertNotIn(token, str(ctx.exception))
        self.assertIn("<redacted>", str(ctx.exception))


    # -- roster builders ------------------------------------------------------

    def test_roster_import_argv_matches_probed_command(self):
        argv = self.cli.roster_import_argv("VNU-HUS", "introai", "/tmp/r/roster.csv")
        self.assertEqual(
            argv,
            ["gh", "teacher", "roster", "import", "VNU-HUS", "introai", "/tmp/r/roster.csv"],
        )

    def test_roster_import_rejects_flaglike_csv_path(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.roster_import_argv("O", "C", "--config=/etc/gh")
        self.assertEqual(ctx.exception.code, "flaglike_operand")

    def test_roster_import_rejects_empty_operands(self):
        for org, classroom, path in (
            ("", "C", "r.csv"), ("O", "", "r.csv"), ("O", "C", ""),
        ):
            with self.subTest(org=org, classroom=classroom, path=path):
                with self.assertRaises(Classroom50Error):
                    self.cli.roster_import_argv(org, classroom, path)

    def test_roster_list_argv_ends_with_json(self):
        argv = self.cli.roster_list_argv("O", "C")
        self.assertEqual(argv, ["gh", "teacher", "roster", "list", "O", "C", "--json"])

    def test_roster_add_argv_pins_flag_order(self):
        argv = self.cli.roster_add_argv(
            "VNU-HUS", "introai", "alice",
            first_name="Alice", last_name="Andersson",
            email="alice@x.edu", section="S1",
        )
        self.assertEqual(
            argv,
            [
                "gh", "teacher", "roster", "add",
                "--first-name", "Alice",
                "--last-name", "Andersson",
                "--email", "alice@x.edu",
                "--section", "S1",
                "VNU-HUS", "introai", "alice",
            ],
        )

    def test_roster_add_omits_unset_flags(self):
        argv = self.cli.roster_add_argv("O", "C", "alice", email="a@x.edu", section="  ")
        self.assertEqual(
            argv,
            ["gh", "teacher", "roster", "add", "--email", "a@x.edu", "O", "C", "alice"],
        )

    def test_roster_add_rejects_non_login_username(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.roster_add_argv("O", "C", "not a login")
        self.assertEqual(ctx.exception.code, "invalid_username")


class TestAssignmentPreflight(unittest.TestCase):
    def test_add_refuses_existing_slug(self):
        run = _runner({_list_call("O", "C"): RunResult(0, json.dumps(MANIFEST))})
        with self.assertRaises(Classroom50Error) as ctx:
            preflight_assignment_add("O", "C", "final-project", runner=run)
        self.assertEqual(ctx.exception.code, "assignment_exists")
        self.assertIn("submission mode", str(ctx.exception))

    def test_add_allows_new_slug(self):
        run = _runner({_list_call("O", "C"): RunResult(0, json.dumps(MANIFEST))})
        self.assertIsNone(preflight_assignment_add("O", "C", "week-1", runner=run))

    def test_add_returns_record_when_overwrite_allowed(self):
        run = _runner({_list_call("O", "C"): RunResult(0, json.dumps(MANIFEST))})
        record = preflight_assignment_add(
            "O", "C", "final-project", allow_overwrite=True, runner=run
        )
        self.assertEqual(record["slug"], "final-project")

    def test_remove_refuses_absent_slug(self):
        run = _runner({_list_call("O", "C"): RunResult(0, json.dumps(MANIFEST))})
        with self.assertRaises(Classroom50Error) as ctx:
            preflight_assignment_remove("O", "C", "week-9z", runner=run)
        self.assertEqual(ctx.exception.code, "assignment_absent")

    def test_remove_returns_record_for_present_slug(self):
        run = _runner({_list_call("O", "C"): RunResult(0, json.dumps(MANIFEST))})
        record = preflight_assignment_remove("O", "C", "week-0a", runner=run)
        self.assertEqual(record["slug"], "week-0a")


class TestDownloadGate(unittest.TestCase):
    def test_by_pattern_refused_when_not_empty_repo(self):
        run = _runner({_list_call("O", "C"): RunResult(0, json.dumps(MANIFEST))})
        with self.assertRaises(Classroom50Error) as ctx:
            download_submissions("O", "C", "week-0a", "dest", by_pattern=True, runner=run)
        self.assertEqual(ctx.exception.code, "by_pattern_not_permitted")

    def test_by_pattern_refused_when_assignment_absent(self):
        run = _runner({_list_call("O", "C"): RunResult(0, json.dumps(MANIFEST))})
        with self.assertRaises(Classroom50Error) as ctx:
            download_submissions("O", "C", "ghost", "dest", by_pattern=True, runner=run)
        self.assertEqual(ctx.exception.code, "assignment_not_found")

    def test_by_pattern_allowed_for_empty_repo_assignment(self):
        calls = []
        run = _runner(
            {
                _list_call("O", "C"): RunResult(0, json.dumps(MANIFEST)),
                (
                    "gh", "teacher", "download", "--by-pattern",
                    "O", "C", "final-project", "-d", "dest",
                ): RunResult(0, "cloned"),
            },
            calls,
        )
        result = download_submissions(
            "O", "C", "final-project", "dest", by_pattern=True, runner=run
        )
        self.assertEqual(result.returncode, 0)
        self.assertIn("--by-pattern", calls[-1])

    def test_plain_download_never_reads_the_manifest(self):
        calls = []
        run = _runner(
            {
                ("gh", "teacher", "download", "O", "C", "week-0a", "-d", "dest"): RunResult(
                    0, "cloned"
                )
            },
            calls,
        )
        download_submissions("O", "C", "week-0a", "dest", runner=run)
        self.assertEqual(len(calls), 1)
        self.assertNotIn("--by-pattern", calls[0])


class TestInvite(unittest.TestCase):
    def test_org_invite_skips_existing_member_and_pending_invitation(self):
        calls = []
        run = _runner(
            {
                _member_call("VNU-HUS"): RunResult(0, json.dumps(MEMBERS)),
                ("gh", "teacher", "invite", "VNU-HUS", "carol"): RunResult(0, "invited"),
            },
            calls,
        )
        report = invite_users("VNU-HUS", ["alice", "bob", "carol"], runner=run)
        self.assertEqual(report["invited"], ["carol"])
        self.assertEqual(sorted(report["skipped"]), ["alice", "bob"])
        self.assertEqual(report["failed"], [])
        self.assertEqual(len(calls), 2)

    def test_repo_invite_is_not_preflighted(self):
        calls = []
        run = _runner(
            {
                ("gh", "teacher", "invite", "VNU-HUS/w0", "alice"): RunResult(0, "ok"),
            },
            calls,
        )
        report = invite_users("VNU-HUS/w0", ["alice"], runner=run)
        self.assertEqual(report["invited"], ["alice"])
        self.assertEqual(len(calls), 1)

    def test_one_failure_does_not_abort_the_batch(self):
        run = _runner(
            {
                _member_call("O"): RunResult(0, "[]"),
                ("gh", "teacher", "invite", "O", "alice"): RunResult(0, "ok"),
                ("gh", "teacher", "invite", "O", "bob"): RunResult(1, "", "denied"),
            }
        )
        report = invite_users("O", ["alice", "bob"], runner=run)
        self.assertEqual(report["invited"], ["alice"])
        self.assertEqual([f["username"] for f in report["failed"]], ["bob"])

    def test_duplicate_usernames_are_invited_once(self):
        calls = []
        run = _runner(
            {
                _member_call("O"): RunResult(0, "[]"),
                ("gh", "teacher", "invite", "O", "alice"): RunResult(0, "ok"),
            },
            calls,
        )
        report = invite_users("O", ["alice", "Alice"], runner=run)
        self.assertEqual(report["invited"], ["alice"])
        self.assertEqual(len(calls), 2)


class _CLICase(unittest.TestCase):
    def setUp(self):
        os.environ.pop("COURSE_C50_AGENT_MODE", None)
        self.out = io.StringIO()
        self.err = io.StringIO()
        self.prompts = []

    def tearDown(self):
        os.environ.pop("COURSE_C50_AGENT_MODE", None)

    def run_cli(self, argv, *, runner=None, answers=None, tty=True):
        replies = iter(answers or [])
        self.prompts = []

        def input_fn(prompt):
            self.prompts.append(prompt)
            return next(replies)

        calls = []
        return (
            c50_admin_cli.main(
                argv,
                runner=runner or _runner({}, calls),
                input_fn=input_fn,
                tty_check=lambda: tty,
                stdout=self.out,
                stderr=self.err,
            ),
            calls,
        )


class TestAdminCLIGates(_CLICase):
    ADD = [
        "assignment-add", "--org", "O", "--classroom", "C",
        "--slug", "w0", "--name", "W0",
    ]
    def setUp(self):
        super().setUp()
        # roster-import reads the CSV before it decides anything, so the dry-run
        # gate needs a file that parses.
        workdir = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, workdir, True)
        csv_path = str(Path(workdir) / "roster.csv")
        Path(csv_path).write_text(_CSV_TEXT, encoding="utf-8")
        # Every mutating verb. A verb missing from this list has no gate coverage.
        self.verbs = [
            self.ADD,
            ["assignment-remove", "--org", "O", "--classroom", "C", "--slug", "w0"],
            ["invite", "--target", "O", "--username", "alice"],
            ["download", "--org", "O", "--classroom", "C", "--assignment", "w0",
             "--dest", "d"],
            ["roster-import", "--org", "O", "--classroom", "C", "--csv", csv_path],
            ["roster-add", "--org", "O", "--classroom", "C", "--username", "alice"],
        ]

    def test_dry_run_prints_argv_and_calls_nothing(self):
        calls = []
        code, _ = self.run_cli(self.ADD + ["--dry-run"], runner=_runner({}, calls))
        self.assertEqual(code, 0)
        payload = json.loads(self.out.getvalue())
        self.assertTrue(payload["dryRun"])
        self.assertEqual(
            payload["argv"],
            ["gh", "teacher", "assignment", "add", "O", "C", "w0", "--name", "W0"],
        )
        self.assertEqual(calls, [])

    def test_dry_run_still_validates(self):
        code, _ = self.run_cli(
            [
                "assignment-add", "--org", "O", "--classroom", "C",
                "--slug", "BAD_SLUG", "--name", "W0", "--dry-run",
            ]
        )
        self.assertEqual(code, 2)
        self.assertEqual(json.loads(self.err.getvalue())["code"], "invalid_slug")

    def test_agent_mode_refuses_every_mutating_verb(self):
        os.environ["COURSE_C50_AGENT_MODE"] = "1"
        for argv in self.verbs:
            with self.subTest(argv=argv[0]):
                self.out.truncate(0), self.out.seek(0)
                self.err.truncate(0), self.err.seek(0)
                code, calls = self.run_cli(argv)
                self.assertEqual(code, 2)
                self.assertEqual(
                    json.loads(self.err.getvalue())["code"], "agent_forbidden"
                )
                self.assertEqual(calls, [])

    def test_agent_mode_refuses_even_dry_run(self):
        os.environ["COURSE_C50_AGENT_MODE"] = "1"
        for argv in self.verbs:
            with self.subTest(argv=argv[0]):
                self.err.truncate(0), self.err.seek(0)
                code, _ = self.run_cli(argv + ["--dry-run"])
                self.assertEqual(code, 2)
                self.assertEqual(
                    json.loads(self.err.getvalue())["code"], "agent_forbidden"
                )

    def test_non_tty_refused_unless_dry_run(self):
        for argv in self.verbs:
            with self.subTest(argv=argv[0]):
                self.err.truncate(0), self.err.seek(0)
                code, calls = self.run_cli(argv, tty=False)
                self.assertEqual(code, 2)
                self.assertEqual(
                    json.loads(self.err.getvalue())["code"], "not_interactive"
                )
                self.assertEqual(calls, [])

    def test_non_tty_dry_run_allowed(self):
        for argv in self.verbs:
            with self.subTest(argv=argv[0]):
                self.out.truncate(0), self.out.seek(0)
                code, _ = self.run_cli(argv + ["--dry-run"], tty=False)
                self.assertEqual(code, 0)


class TestAdminCLIAssignment(_CLICase):
    def test_add_refuses_existing_slug_without_overwrite(self):
        run = _runner({_list_call("O", "C"): RunResult(0, json.dumps(MANIFEST))})
        code, _ = self.run_cli(
            [
                "assignment-add", "--org", "O", "--classroom", "C",
                "--slug", "final-project", "--name", "F",
            ],
            runner=run,
            answers=["y"],
        )
        self.assertEqual(code, 2)
        self.assertEqual(json.loads(self.err.getvalue())["code"], "assignment_exists")

    def test_overwrite_requires_confirmation(self):
        run = _runner({_list_call("O", "C"): RunResult(0, json.dumps(MANIFEST))})
        code, _ = self.run_cli(
            [
                "assignment-add", "--org", "O", "--classroom", "C",
                "--slug", "final-project", "--name", "F", "--allow-overwrite",
            ],
            runner=run,
            answers=["n"],
        )
        self.assertEqual(code, 2)
        self.assertEqual(json.loads(self.err.getvalue())["code"], "cancelled")

    def test_overwrite_confirmation_names_the_submission_mode_hazard(self):
        run = _runner({_list_call("O", "C"): RunResult(0, json.dumps(MANIFEST))})
        prompts = []

        def input_fn(prompt):
            prompts.append(prompt)
            return "n"

        c50_admin_cli.main(
            [
                "assignment-add", "--org", "O", "--classroom", "C",
                "--slug", "final-project", "--name", "F", "--allow-overwrite",
            ],
            runner=run,
            input_fn=input_fn,
            tty_check=lambda: True,
            stdout=self.out,
            stderr=self.err,
        )
        self.assertIn("submission mode", prompts[0])

    def test_add_succeeds_for_new_slug_after_confirmation(self):
        add = ("gh", "teacher", "assignment", "add", "O", "C", "week-1", "--name", "W1")
        run = _runner(
            {
                _list_call("O", "C"): RunResult(0, json.dumps(MANIFEST)),
                add: RunResult(0, "registered"),
            }
        )
        code, _ = self.run_cli(
            [
                "assignment-add", "--org", "O", "--classroom", "C",
                "--slug", "week-1", "--name", "W1",
            ],
            runner=run,
            answers=["yes"],
        )
        self.assertEqual(code, 0)
        self.assertEqual(json.loads(self.out.getvalue())["slug"], "week-1")

    def test_remove_refuses_absent_slug(self):
        run = _runner({_list_call("O", "C"): RunResult(0, json.dumps(MANIFEST))})
        code, _ = self.run_cli(
            ["assignment-remove", "--org", "O", "--classroom", "C", "--slug", "ghost"],
            runner=run,
            answers=["y"],
        )
        self.assertEqual(code, 2)
        self.assertEqual(json.loads(self.err.getvalue())["code"], "assignment_absent")

    def test_remove_confirmation_states_repos_survive(self):
        run = _runner({_list_call("O", "C"): RunResult(0, json.dumps(MANIFEST))})
        prompts = []

        def input_fn(prompt):
            prompts.append(prompt)
            return "n"

        c50_admin_cli.main(
            ["assignment-remove", "--org", "O", "--classroom", "C", "--slug", "week-0a"],
            runner=run,
            input_fn=input_fn,
            tty_check=lambda: True,
            stdout=self.out,
            stderr=self.err,
        )
        self.assertIn("does not delete", prompts[0])
        self.assertIn("not a clean reset", prompts[0])

    def test_failed_mutation_exits_outcome_unknown(self):
        add = ("gh", "teacher", "assignment", "add", "O", "C", "week-1", "--name", "W1")
        run = _runner(
            {
                _list_call("O", "C"): RunResult(0, json.dumps(MANIFEST)),
                add: RunResult(1, "", "server said no"),
            }
        )
        code, _ = self.run_cli(
            [
                "assignment-add", "--org", "O", "--classroom", "C",
                "--slug", "week-1", "--name", "W1",
            ],
            runner=run,
            answers=["y"],
        )
        self.assertEqual(code, 3)


class TestAdminCLIInviteAndDownload(_CLICase):
    def test_invite_partial_failure_exits_4(self):
        run = _runner(
            {
                _member_call("O"): RunResult(0, "[]"),
                ("gh", "teacher", "invite", "O", "alice"): RunResult(0, "ok"),
                ("gh", "teacher", "invite", "O", "bob"): RunResult(1, "", "denied"),
            }
        )
        code, _ = self.run_cli(
            ["invite", "--target", "O", "--username", "alice", "--username", "bob"],
            runner=run,
            answers=["y"],
        )
        self.assertEqual(code, 4)
        self.assertEqual(json.loads(self.out.getvalue())["invited"], ["alice"])

    def test_invite_all_skipped_exits_zero(self):
        run = _runner({_member_call("O"): RunResult(0, json.dumps(MEMBERS))})
        code, _ = self.run_cli(
            ["invite", "--target", "O", "--username", "alice"],
            runner=run,
            answers=["y"],
        )
        self.assertEqual(code, 0)
        self.assertEqual(json.loads(self.out.getvalue())["skipped"], ["alice"])

    def test_download_by_pattern_refused_for_autograded_assignment(self):
        run = _runner({_list_call("O", "C"): RunResult(0, json.dumps(MANIFEST))})
        code, _ = self.run_cli(
            [
                "download", "--org", "O", "--classroom", "C",
                "--assignment", "week-0a", "--dest", "d", "--by-pattern",
            ],
            runner=run,
        )
        self.assertEqual(code, 2)
        self.assertEqual(
            json.loads(self.err.getvalue())["code"], "by_pattern_not_permitted"
        )

    def test_download_by_pattern_allowed_for_empty_repo_assignment(self):
        run = _runner(
            {
                _list_call("O", "C"): RunResult(0, json.dumps(MANIFEST)),
                (
                    "gh", "teacher", "download", "--by-pattern",
                    "O", "C", "final-project", "-d", "d",
                ): RunResult(0, "cloned"),
            }
        )
        code, _ = self.run_cli(
            [
                "download", "--org", "O", "--classroom", "C",
                "--assignment", "final-project", "--dest", "d", "--by-pattern",
            ],
            runner=run,
        )
        self.assertEqual(code, 0)

    def test_download_needs_no_confirmation(self):
        run = _runner(
            {
                ("gh", "teacher", "download", "O", "C", "week-0a", "-d", "d"): RunResult(
                    0, "cloned"
                )
            }
        )
        code, _ = self.run_cli(
            [
                "download", "--org", "O", "--classroom", "C",
                "--assignment", "week-0a", "--dest", "d",
            ],
            runner=run,
        )
        self.assertEqual(code, 0)


class _TestsFileCase(unittest.TestCase):
    """Weekly registration needs a real `--tests` file, so give each case one."""

    def setUp(self):
        self.cli = HumanCLI(runner=_runner({}))
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.tests_path = os.path.join(self.tmp.name, "classroom50-tests.json")
        with open(self.tests_path, "w", encoding="utf-8") as handle:
            json.dump([{"type": "run", "cmd": "python3 hello.py"}], handle)

    def write(self, name, text):
        path = os.path.join(self.tmp.name, name)
        with open(path, "w", encoding="utf-8") as handle:
            handle.write(text)
        return path


class TestWeeklyAssignmentFlags(_TestsFileCase):
    def test_argv_matches_the_runbook_weekly_command(self):
        argv = self.cli.assignment_add_argv(
            "VNU-HUS",
            "classroom50-pilot-2026",
            "w00-individual-onboarding",
            name="Week 0A",
            template="VNU-HUS/introai-w00-individual-template@main",
            tests=self.tests_path,
            available_from="2026-09-04T13:00:00+07:00",
            due="2026-09-11T23:59:00+07:00",
            feedback_pr=True,
            pass_threshold=100,
            mode="individual",
        )
        self.assertEqual(
            argv,
            [
                "gh", "teacher", "assignment", "add",
                "VNU-HUS", "classroom50-pilot-2026", "w00-individual-onboarding",
                "--name", "Week 0A",
                "--template", "VNU-HUS/introai-w00-individual-template@main",
                "--tests", self.tests_path,
                "--mode", "individual",
                "--available-from", "2026-09-04T13:00:00+07:00",
                "--due", "2026-09-11T23:59:00+07:00",
                "--feedback-pr",
                "--pass-threshold", "100",
            ],
        )

    def test_feedback_pr_disabled_uses_the_equals_form(self):
        argv = self.cli.assignment_add_argv("O", "C", "w0", name="x", feedback_pr=False)
        self.assertIn("--feedback-pr=false", argv)
        self.assertNotIn("--feedback-pr", argv)

    def test_feedback_pr_omitted_when_unset(self):
        argv = self.cli.assignment_add_argv("O", "C", "w0", name="x")
        self.assertFalse([a for a in argv if a.startswith("--feedback-pr")])

    def test_missing_tests_file_is_refused(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.assignment_add_argv(
                "O", "C", "w0", name="x", tests=os.path.join(self.tmp.name, "nope.json")
            )
        self.assertEqual(ctx.exception.code, "invalid_tests_file")

    def test_unparsable_tests_file_is_refused(self):
        path = self.write("bad.json", "{not json")
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.assignment_add_argv("O", "C", "w0", name="x", tests=path)
        self.assertEqual(ctx.exception.code, "invalid_tests_file")

    def test_tests_file_must_hold_an_array(self):
        path = self.write("obj.json", '{"tests": []}')
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.assignment_add_argv("O", "C", "w0", name="x", tests=path)
        self.assertEqual(ctx.exception.code, "invalid_tests_file")

    def test_empty_repo_conflicts_with_every_autograded_flag(self):
        cases = {
            "tests": self.tests_path,
            "feedback_pr": True,
            "pass_threshold": 100,
            "allowed_files": ["*"],
            "template": "o/t",
        }
        for keyword, value in cases.items():
            with self.subTest(keyword=keyword):
                with self.assertRaises(Classroom50Error) as ctx:
                    self.cli.assignment_add_argv(
                        "O", "C", "w0", name="x", empty_repo=True, **{keyword: value}
                    )
                self.assertEqual(ctx.exception.code, "mutually_exclusive")

    def test_pass_threshold_outside_percentage_range_is_refused(self):
        for value in (-1, 101):
            with self.subTest(value=value):
                with self.assertRaises(Classroom50Error) as ctx:
                    self.cli.assignment_add_argv(
                        "O", "C", "w0", name="x", pass_threshold=value
                    )
                self.assertEqual(ctx.exception.code, "invalid_pass_threshold")

    def test_pass_threshold_zero_is_emitted(self):
        argv = self.cli.assignment_add_argv("O", "C", "w0", name="x", pass_threshold=0)
        self.assertEqual(argv[-2:], ["--pass-threshold", "0"])

    def test_allowed_files_preserve_their_order(self):
        argv = self.cli.assignment_add_argv(
            "O", "C", "w0", name="x", allowed_files=["*", "!hello.py"]
        )
        self.assertEqual(
            argv[-4:],
            ["--allowed-files", "*", "--allowed-files", "!hello.py"],
        )

    def test_allowed_files_reject_a_flaglike_pattern(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.assignment_add_argv(
                "O", "C", "w0", name="x", allowed_files=["--force"]
            )
        self.assertEqual(ctx.exception.code, "flaglike_operand")

    def test_student_permission_is_validated(self):
        with self.assertRaises(Classroom50Error) as ctx:
            self.cli.assignment_add_argv(
                "O", "C", "w0", name="x", student_permission="owner"
            )
        self.assertEqual(ctx.exception.code, "invalid_permission")
        argv = self.cli.assignment_add_argv(
            "O", "C", "w0", name="x", student_permission="admin"
        )
        self.assertEqual(argv[-2:], ["--student-permission", "admin"])


class TestAdminCLIWeeklyRegistration(_CLICase):
    def setUp(self):
        super().setUp()
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.tests_path = os.path.join(self.tmp.name, "classroom50-tests.json")
        with open(self.tests_path, "w", encoding="utf-8") as handle:
            json.dump([], handle)

    def test_dry_run_reproduces_the_weekly_command(self):
        calls = []
        code, _ = self.run_cli(
            [
                "assignment-add", "--org", "VNU-HUS",
                "--classroom", "classroom50-pilot-2026",
                "--slug", "w00-group-collaboration", "--name", "Week 0B",
                "--template", "VNU-HUS/introai-w00-group-template@main",
                "--tests", self.tests_path,
                "--available-from", "2026-09-04T13:00:00+07:00",
                "--due", "2026-09-11T23:59:00+07:00",
                "--feedback-pr", "--pass-threshold", "100",
                "--mode", "group", "--max-group-size", "5",
                "--dry-run",
            ],
            runner=_runner({}, calls),
        )
        self.assertEqual(code, 0)
        payload = json.loads(self.out.getvalue())
        self.assertEqual(calls, [])
        self.assertEqual(
            payload["argv"],
            [
                "gh", "teacher", "assignment", "add",
                "VNU-HUS", "classroom50-pilot-2026", "w00-group-collaboration",
                "--name", "Week 0B",
                "--template", "VNU-HUS/introai-w00-group-template@main",
                "--tests", self.tests_path,
                "--mode", "group",
                "--max-group-size", "5",
                "--available-from", "2026-09-04T13:00:00+07:00",
                "--due", "2026-09-11T23:59:00+07:00",
                "--feedback-pr",
                "--pass-threshold", "100",
            ],
        )

    def test_no_feedback_pr_flag_reaches_the_adapter(self):
        code, _ = self.run_cli(
            [
                "assignment-add", "--org", "O", "--classroom", "C",
                "--slug", "w0", "--name", "W0", "--no-feedback-pr", "--dry-run",
            ]
        )
        self.assertEqual(code, 0)
        self.assertIn("--feedback-pr=false", json.loads(self.out.getvalue())["argv"])

    def test_missing_tests_file_exits_two(self):
        code, _ = self.run_cli(
            [
                "assignment-add", "--org", "O", "--classroom", "C",
                "--slug", "w0", "--name", "W0",
                "--tests", os.path.join(self.tmp.name, "absent.json"), "--dry-run",
            ]
        )
        self.assertEqual(code, 2)
        self.assertEqual(json.loads(self.err.getvalue())["code"], "invalid_tests_file")

def _roster_runner(remote_rows, calls, *, import_result=None, on_import=None, members=None):
    """Runner for the roster verbs: `roster list` replies, `roster import` records.

    ``members`` is the org membership `roster import` reads before short-circuiting.
    Left at ``None`` the call fails, which is what every test that never reaches
    that path wants.
    """

    def run(argv):
        calls.append(list(argv))
        if argv[:4] == ["gh", "teacher", "roster", "list"]:
            return RunResult(0, json.dumps(remote_rows))
        if argv[:4] == ["gh", "teacher", "member", "list"] and members is not None:
            return RunResult(0, json.dumps([{"login": m} for m in members]))
        if argv[:4] == ["gh", "teacher", "roster", "import"]:
            if on_import is not None:
                on_import(argv)
            return import_result if import_result is not None else RunResult(0, "", "")
        return RunResult(1, "", f"unexpected argv: {argv}")

    return run


def _remote(username, *, first="", last="", email="", section=""):
    return {
        "username": username,
        "first_name": first,
        "last_name": last,
        "email": email,
        "section": section,
        "github_id": "1",
        "role": "student",
    }


def _student(name, username, email, section="S1"):
    first, _, last = name.partition(" ")
    return Student(
        **{
            "Name": name,
            "First Name": first,
            "Last Name": last,
            "Email": email,
            "Section": section,
            "GitHub Username": username,
        }
    )


ALICE = ("Alice Andersson", "alice", "alice@x.edu")
BOB = ("Bob Brown", "bob", "bob@x.edu")


class TestAdminCLIRosterImport(_CLICase):
    def setUp(self):
        super().setUp()
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.students = [_student(*ALICE), _student(*BOB)]
        # course_hoanganhduc.data pulls pandas; the CLI imports load_database lazily
        # exactly so this stub is enough (pinned by test_import_stays_lazy).
        module = types.ModuleType("course_hoanganhduc.data")
        module.load_database = lambda path, verbose=False: list(self.students)
        self._saved = sys.modules.get("course_hoanganhduc.data")
        sys.modules["course_hoanganhduc.data"] = module
        self.addCleanup(self._restore_data_module)

    def _restore_data_module(self):
        if self._saved is None:
            sys.modules.pop("course_hoanganhduc.data", None)
        else:
            sys.modules["course_hoanganhduc.data"] = self._saved

    def _csv_file(self, text=_CSV_TEXT):
        path = str(Path(self.tmp.name) / "given.csv")
        Path(path).write_text(text, encoding="utf-8")
        return path

    def _import_argv(self, base=("roster-import", "--org", "O", "--classroom", "C")):
        return list(base)

    @staticmethod
    def _workdirs():
        return sorted(Path(tempfile.gettempdir()).glob("c50-roster-*"))

    # -- dry run --------------------------------------------------------------

    def test_dry_run_csv_reports_real_argv_and_calls_nothing(self):
        path = self._csv_file()
        code, calls = self.run_cli(self._import_argv() + ["--csv", path, "--dry-run"])
        self.assertEqual(code, 0)
        payload = json.loads(self.out.getvalue())
        self.assertEqual(payload["argv"], list(_roster_import_call("O", "C", path)))
        self.assertEqual(payload["plan"]["csvSource"], "file")
        self.assertEqual(calls, [])

    def test_dry_run_db_reports_a_plan_and_no_argv(self):
        code, calls = self.run_cli(self._import_argv() + ["--db", "students.db", "--dry-run"])
        self.assertEqual(code, 0)
        payload = json.loads(self.out.getvalue())
        self.assertNotIn("argv", payload)
        self.assertEqual(payload["plan"]["csvSource"], "generated-from-db")
        self.assertEqual(payload["plan"]["rowCount"], 2)
        self.assertEqual(payload["plan"]["firstUsernames"], ["alice", "bob"])
        self.assertEqual(calls, [])

    def test_dry_run_writes_no_temporary_csv(self):
        before = self._workdirs()
        self.run_cli(self._import_argv() + ["--db", "students.db", "--dry-run"])
        self.assertEqual(self._workdirs(), before)

    def test_database_progress_output_stays_off_stdout(self):
        # The real load_database prints "Loading database from ..." unconditionally.
        # stdout carries the JSON report, so that chatter belongs on the error stream.
        def loud(path, verbose=False):
            print(f"Loading database from {path}...")
            return list(self.students)

        sys.modules["course_hoanganhduc.data"].load_database = loud
        code, _ = self.run_cli(self._import_argv() + ["--db", "students.db", "--dry-run"])
        self.assertEqual(code, 0)
        payload = json.loads(self.out.getvalue())
        self.assertEqual(payload["plan"]["rowCount"], 2)
        self.assertIn("Loading database from students.db", self.err.getvalue())

    # -- temporary CSV lifecycle ---------------------------------------------

    def test_temp_csv_exists_during_the_call_and_is_removed_after(self):
        seen = {}

        def on_import(argv):
            seen["path"] = argv[-1]
            seen["exists"] = os.path.exists(argv[-1])

        calls = []
        code = self._run_import(
            ["--db", "students.db"],
            _roster_runner([], calls, on_import=on_import),
            calls,
            answers=["y"],
        )
        self.assertEqual(code, 0)
        self.assertTrue(seen["exists"])
        self.assertFalse(os.path.exists(seen["path"]))

    def test_keep_csv_preserves_the_file_and_reports_its_path(self):
        calls = []
        code = self._run_import(
            ["--db", "students.db", "--keep-csv"],
            _roster_runner([], calls),
            calls,
            answers=["y"],
        )
        self.assertEqual(code, 0)
        csv_path = json.loads(self.out.getvalue())["csvPath"]
        self.addCleanup(shutil.rmtree, str(Path(csv_path).parent), True)
        self.assertTrue(os.path.exists(csv_path))

    def test_temp_csv_is_removed_even_when_gh_fails(self):
        seen = {}
        calls = []
        code = self._run_import(
            ["--db", "students.db"],
            _roster_runner(
                [],
                calls,
                import_result=RunResult(1, "", "boom"),
                on_import=lambda argv: seen.__setitem__("path", argv[-1]),
            ),
            calls,
            answers=["y"],
        )
        self.assertEqual(code, 3)
        self.assertFalse(os.path.exists(seen["path"]))

    def test_gh_failure_is_outcome_unknown(self):
        calls = []
        code = self._run_import(
            ["--db", "students.db"],
            _roster_runner([], calls, import_result=RunResult(1, "", "boom")),
            calls,
            answers=["y"],
        )
        self.assertEqual(code, 3)
        payload = json.loads(self.err.getvalue())
        self.assertEqual(payload["error"], "outcome_unknown")
        self.assertEqual(payload["code"], "roster_import_failed")

    # -- reporting ------------------------------------------------------------

    def test_students_without_a_github_username_are_reported_and_exit_zero(self):
        self.students.append(
            Student(**{"Name": "Cara Cruz", "Email": "cara@x.edu", "Section": "S1"})
        )
        calls = []
        code = self._run_import(
            ["--db", "students.db"], _roster_runner([], calls), calls, answers=["y"]
        )
        self.assertEqual(code, 0)
        self.assertEqual(
            json.loads(self.out.getvalue())["skippedNoUsername"], ["Cara Cruz"]
        )

    def test_confirmation_truncates_a_long_list(self):
        self.students = [
            _student(f"S{i:02d} Student", f"s{i:02d}", f"s{i:02d}@x.edu")
            for i in range(1, 26)
        ]
        calls = []
        self._run_import(
            ["--db", "students.db"], _roster_runner([], calls), calls, answers=["y"]
        )
        prompt = self.prompts[0]
        self.assertIn("... and 15 more", prompt)
        self.assertIn("s10", prompt)
        self.assertNotIn("s11", prompt)

    def test_cancelling_writes_nothing(self):
        calls = []
        code = self._run_import(
            ["--db", "students.db"], _roster_runner([], calls), calls, answers=["n"]
        )
        self.assertEqual(code, 2)
        self.assertEqual(json.loads(self.err.getvalue())["code"], "cancelled")
        self.assertEqual(calls, [list(_roster_list_call("O", "C"))])

    def test_rows_only_on_the_remote_roster_are_reported_never_sent(self):
        sent = {}
        calls = []
        remote = [
            _remote("alice", first="Alice", last="Andersson", email="alice@x.edu", section="S1"),
            _remote("dave-ta", first="Dave", last="Dinh", email="dave@x.edu", section="S1"),
        ]
        code = self._run_import(
            ["--db", "students.db"],
            _roster_runner(
                remote,
                calls,
                on_import=lambda argv: sent.__setitem__(
                    "text", Path(argv[-1]).read_text(encoding="utf-8")
                ),
            ),
            calls,
            answers=["y"],
        )
        self.assertEqual(code, 0)
        report = json.loads(self.out.getvalue())
        self.assertEqual(report["remoteOnly"], ["dave-ta"])
        self.assertEqual(report["new"], ["bob"])
        self.assertNotIn("dave-ta", sent["text"])

    def _matching_remote(self):
        return [
            _remote("alice", first="Alice", last="Andersson", email="alice@x.edu", section="S1"),
            _remote("bob", first="Bob", last="Brown", email="bob@x.edu", section="S1"),
        ]

    def test_second_run_short_circuits_without_importing(self):
        calls = []
        code = self._run_import(
            ["--db", "students.db"],
            _roster_runner(self._matching_remote(), calls, members=["alice", "bob"]),
            calls,
        )
        self.assertEqual(code, 0)
        report = json.loads(self.out.getvalue())
        self.assertEqual(report["action"], "unchanged")
        self.assertNotIn("notYetMembers", report)
        self.assertEqual(
            calls, [list(_roster_list_call("O", "C")), list(_member_call("O"))]
        )
        self.assertIn("nothing to import", self.err.getvalue())

    def test_matching_roster_still_imports_when_a_student_is_not_yet_a_member(self):
        # `roster import` commits the CSV and invites afterwards, so a run that
        # died between the two leaves this exact state. Short-circuiting here
        # would make the documented "just re-run it" recovery a no-op.
        calls = []
        code = self._run_import(
            ["--db", "students.db"],
            _roster_runner(self._matching_remote(), calls, members=["alice"]),
            calls,
            answers=["y"],
        )
        self.assertEqual(code, 0)
        report = json.loads(self.out.getvalue())
        self.assertEqual(report["action"], "imported")
        self.assertEqual(report["notYetMembers"], ["bob"])
        self.assertIn(["gh", "teacher", "roster", "import", "O", "C"], [c[:6] for c in calls])
        self.assertIn("Not in org yet:   1", self.prompts[0])

    def test_unreadable_member_list_degrades_instead_of_reporting_unknown(self):
        calls = []
        code = self._run_import(
            ["--db", "students.db"], _roster_runner(self._matching_remote(), calls), calls
        )
        # Nothing was written, so this is not outcome-unknown (exit 3).
        self.assertEqual(code, 0)
        report = json.loads(self.out.getvalue())
        self.assertEqual(report["action"], "unchanged")
        self.assertTrue(report["membershipNotVerified"])
        self.assertIn("--force", self.err.getvalue())

    def test_force_imports_a_matching_roster_without_reading_membership(self):
        calls = []
        code = self._run_import(
            ["--db", "students.db", "--force"],
            _roster_runner(self._matching_remote(), calls),
            calls,
            answers=["y"],
        )
        self.assertEqual(code, 0)
        self.assertEqual(json.loads(self.out.getvalue())["action"], "imported")
        self.assertNotIn(list(_member_call("O")), calls)

    def test_keep_csv_with_csv_source_is_refused(self):
        calls = []
        code = self._run_import(
            ["--csv", self._csv_file(), "--keep-csv"], _roster_runner([], calls), calls
        )
        self.assertEqual(code, 2)
        payload = json.loads(self.err.getvalue())
        self.assertEqual(payload["code"], "invalid_flag_combination")
        self.assertEqual(calls, [])

    def test_generated_csv_is_utf8_without_bom_and_canonically_headed(self):
        self.students = [_student("Đặng Vân Anh", "van-anh", "vananh@x.edu")]
        sent = {}
        calls = []
        self._run_import(
            ["--db", "students.db"],
            _roster_runner(
                [], calls, on_import=lambda argv: sent.__setitem__(
                    "raw", Path(argv[-1]).read_bytes()
                )
            ),
            calls,
            answers=["y"],
        )
        raw = sent["raw"]
        self.assertFalse(raw.startswith(b"\xef\xbb\xbf"))
        text = raw.decode("utf-8")
        self.assertEqual(text.splitlines()[0], ",".join(CANONICAL_COLUMNS))
        self.assertIn("Đặng", text)

    def test_db_and_csv_are_a_required_exclusive_pair(self):
        for extra in ([], ["--db", "students.db", "--csv", "r.csv"]):
            with self.subTest(extra=extra):
                with contextlib.redirect_stderr(io.StringIO()):
                    with self.assertRaises(SystemExit) as ctx:
                        c50_admin_cli.main(self._import_argv() + extra)
                self.assertEqual(ctx.exception.code, 2)

    def test_import_stays_lazy(self):
        probe = (
            "import course_hoanganhduc.c50_admin_cli, sys; "
            "raise SystemExit('course_hoanganhduc.data' in sys.modules)"
        )
        result = subprocess.run(
            [sys.executable, "-c", probe], cwd=REPO_ROOT, capture_output=True
        )
        self.assertEqual(result.returncode, 0, result.stderr.decode("utf-8", "replace"))

    def _run_import(self, extra, runner, calls, *, answers=None):
        replies = iter(answers or [])

        def input_fn(prompt):
            self.prompts.append(prompt)
            return next(replies)

        return c50_admin_cli.main(
            self._import_argv() + list(extra),
            runner=runner,
            input_fn=input_fn,
            tty_check=lambda: True,
            stdout=self.out,
            stderr=self.err,
        )


class TestAdminCLIRosterAdd(_CLICase):
    ARGV = ["roster-add", "--org", "O", "--classroom", "C", "--username", "alice"]

    def test_dry_run_reports_argv_and_both_effects(self):
        code, calls = self.run_cli(self.ARGV + ["--first-name", "Alice", "--dry-run"])
        self.assertEqual(code, 0)
        payload = json.loads(self.out.getvalue())
        self.assertEqual(
            payload["argv"],
            ["gh", "teacher", "roster", "add", "--first-name", "Alice", "O", "C", "alice"],
        )
        self.assertTrue(any("invited" in p for p in payload["preconditions"]))
        self.assertEqual(calls, [])

    def test_confirmed_add_runs_one_command(self):
        argv = ("gh", "teacher", "roster", "add", "O", "C", "alice")
        calls = []
        code, _ = self.run_cli(
            self.ARGV, runner=_runner({argv: RunResult(0)}, calls), answers=["y"]
        )
        self.assertEqual(code, 0)
        self.assertEqual(calls, [list(argv)])
        self.assertEqual(json.loads(self.out.getvalue())["action"], "upserted")

    def test_cancelled_add_runs_nothing(self):
        code, calls = self.run_cli(self.ARGV, answers=["n"])
        self.assertEqual(code, 2)
        self.assertEqual(json.loads(self.err.getvalue())["code"], "cancelled")
        self.assertEqual(calls, [])


if __name__ == "__main__":
    unittest.main()
