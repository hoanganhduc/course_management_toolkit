#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Offline tests for the Classroom50 operator surface (assignment/invite/download)."""

from __future__ import annotations

import io
import json
import os
import sys
import tempfile
import unittest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc import c50_admin_cli
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

    def tearDown(self):
        os.environ.pop("COURSE_C50_AGENT_MODE", None)

    def run_cli(self, argv, *, runner=None, answers=None, tty=True):
        replies = iter(answers or [])

        def input_fn(_prompt):
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
        for argv in (
            self.ADD,
            ["assignment-remove", "--org", "O", "--classroom", "C", "--slug", "w0"],
            ["invite", "--target", "O", "--username", "alice"],
            ["download", "--org", "O", "--classroom", "C", "--assignment", "w0",
             "--dest", "d"],
        ):
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
        code, _ = self.run_cli(self.ADD + ["--dry-run"])
        self.assertEqual(code, 2)
        self.assertEqual(json.loads(self.err.getvalue())["code"], "agent_forbidden")

    def test_non_tty_refused_unless_dry_run(self):
        code, calls = self.run_cli(self.ADD, tty=False)
        self.assertEqual(code, 2)
        self.assertEqual(json.loads(self.err.getvalue())["code"], "not_interactive")
        self.assertEqual(calls, [])

    def test_non_tty_dry_run_allowed(self):
        code, _ = self.run_cli(self.ADD + ["--dry-run"], tty=False)
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


if __name__ == "__main__":
    unittest.main()
