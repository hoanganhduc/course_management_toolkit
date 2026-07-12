#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Offline Classroom50 unit tests (ADR v2.1 T1–T15 subset)."""

from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from types import SimpleNamespace

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc.c50_cli import AgentCLI, Classroom50Error, RunResult, redact_secrets
from course_hoanganhduc.c50_cli_human import HumanCLI
from course_hoanganhduc.c50_ops import agent_refuse_download, is_agent_mode, list_classrooms, sync
from course_hoanganhduc.c50_roster import (
    CANONICAL_COLUMNS,
    export_roster_csv,
    join_roster,
    parse_csv_text,
    parse_roster_payload,
)
from course_hoanganhduc.models import Student


def _mock_runner(mapping):
    def runner(argv):
        key = tuple(argv)
        if key in mapping:
            return mapping[key]
        # flexible match on last tokens
        for k, v in mapping.items():
            if list(argv[-len(k) :]) == list(k):
                return v
        return RunResult(returncode=1, stderr=f"unexpected argv: {argv}")

    return runner


class TestAgentCLI(unittest.TestCase):
    def test_t1_unknown_method(self):
        cli = AgentCLI(runner=lambda a: RunResult(0, "x"))
        with self.assertRaises(Classroom50Error) as ctx:
            cli.teardown("org")  # type: ignore[attr-defined]
        self.assertEqual(ctx.exception.code, "unknown_method")

    def test_t1_no_download_on_agent(self):
        cli = AgentCLI(runner=lambda a: RunResult(0))
        with self.assertRaises(Classroom50Error) as ctx:
            cli.download("o", "c", "a", "d")
        self.assertEqual(ctx.exception.code, "agent_download_forbidden")

    def test_t2_list_roster_json(self):
        payload = [{"username": "alice", "github_id": 1, "email": "a@x.com"}]
        runner = _mock_runner(
            {
                ("gh", "teacher", "roster", "list", "ORG", "CLS", "--json"): RunResult(
                    0, json.dumps(payload)
                )
            }
        )
        cli = AgentCLI(runner=runner)
        data = cli.list_roster("ORG", "CLS")
        self.assertEqual(data[0]["username"], "alice")

    def test_t3_bad_json(self):
        runner = _mock_runner(
            {
                ("gh", "teacher", "classroom", "list", "ORG", "--json"): RunResult(
                    0, "not-json"
                )
            }
        )
        cli = AgentCLI(runner=runner)
        with self.assertRaises(Classroom50Error) as ctx:
            cli.list_classrooms("ORG")
        self.assertEqual(ctx.exception.code, "bad_json")

    def test_t3_nonzero(self):
        runner = _mock_runner(
            {
                ("gh", "teacher", "classroom", "list", "ORG", "--json"): RunResult(
                    2, "", "boom"
                )
            }
        )
        cli = AgentCLI(runner=runner)
        with self.assertRaises(Classroom50Error):
            cli.list_classrooms("ORG")

    def test_t9_redact_full_token(self):
        raw = "Authorization: Bearer ghp_abcdefghijklmnopqrstuvwxyz0123456789"
        red = redact_secrets(raw)
        self.assertNotIn("ghp_abcdefghijklmnopqrstuvwxyz0123456789", red)
        self.assertIn("<redacted>", red)

    def test_t14_whoami_empty(self):
        cli = AgentCLI(runner=lambda a: RunResult(0, "\n"))
        with self.assertRaises(Classroom50Error) as ctx:
            cli.whoami()
        self.assertEqual(ctx.exception.code, "whoami_empty")

    def test_t15_token_word_in_json_still_parses(self):
        # redaction must not be applied before parse
        body = json.dumps([{"username": "tokenuser", "github_id": 2}])
        runner = _mock_runner(
            {
                ("gh", "teacher", "roster", "list", "O", "C", "--json"): RunResult(
                    0, body
                )
            }
        )
        data = AgentCLI(runner=runner).list_roster("O", "C")
        self.assertEqual(data[0]["username"], "tokenuser")


class TestJoinAndCSV(unittest.TestCase):
    def test_t4_join_matrix(self):
        students = [
            Student(**{"GitHub ID": "10", "Name": "A"}),
            Student(**{"GitHub Username": "bob", "Name": "B"}),
            Student(**{"Email": "c@x.com", "Name": "C"}),
            Student(**{"Name": "Only Local"}),
        ]
        remote = [
            {"username": "alice", "github_id": 10, "email": ""},
            {"username": "Bob", "github_id": 0, "email": ""},
            {"username": "carol", "github_id": 0, "email": "c@x.com"},
            {"username": "nobody", "github_id": 0, "email": ""},
            {"username": "zero", "github_id": "0", "email": ""},
        ]
        report = join_roster(students, remote)
        self.assertEqual(len(report["matched"]), 3)
        self.assertTrue(any(r["username"] == "nobody" for r in report["remote_only"]))
        # id 0 skipped to username path for bob
        self.assertTrue(any(m["via"] == "username" for m in report["matched"]))

    def test_t5_fill_only_and_t5b_distinct(self):
        students = [Student(**{"Name": "A", "Student ID": "MSSV1", "Canvas ID": "99"})]
        remote = [
            {
                "username": "alice",
                "github_id": 42,
                "email": "a@x.com",
                "first_name": "A",
                "last_name": "L",
                "section": "",
            }
        ]
        # match by forcing username after empty - need match: put username empty, use email
        students[0].Email = "a@x.com"
        report = join_roster(students, remote)
        from course_hoanganhduc.c50_roster import apply_fill_only

        for row, idx in report["pairs"]:
            apply_fill_only(students[idx], row, report)
        self.assertEqual(getattr(students[0], "GitHub Username"), "alice")
        self.assertEqual(str(getattr(students[0], "GitHub ID")), "42")
        self.assertEqual(getattr(students[0], "Student ID"), "MSSV1")
        self.assertEqual(getattr(students[0], "Canvas ID"), "99")
        # conflict email keep local if both set differ
        students[0].Email = "local@x.com"
        report2 = {"conflicts": []}
        apply_fill_only(students[0], {"username": "alice", "email": "remote@x.com", "github_id": 42}, report2)
        self.assertEqual(students[0].Email, "local@x.com")
        self.assertTrue(report2["conflicts"])

    def test_t6_csv_roundtrip(self):
        students = [
            Student(
                **{
                    "GitHub Username": "alice",
                    "GitHub ID": "7",
                    "Email": "a@x.com",
                    "Name": "Alice Nguyen",
                }
            )
        ]
        text, mode = export_roster_csv(students)
        self.assertTrue(text.startswith(",".join(CANONICAL_COLUMNS)))
        rows = parse_csv_text("\ufeff" + text)
        self.assertEqual(rows[0]["username"], "alice")
        self.assertEqual(rows[0]["github_id"], "7")
        with self.assertRaises(ValueError):
            parse_csv_text("username,first_name,last_name,email,section,github_id\n,,x,y,z,1\n")

    def test_cross_rule_multi_match(self):
        students = [
            Student(**{"GitHub ID": "1", "Name": "A"}),
            Student(**{"Email": "shared@x.com", "Name": "B"}),
        ]
        remote = [{"username": "x", "github_id": 1, "email": "shared@x.com"}]
        report = join_roster(students, remote)
        self.assertTrue(report["multi_match"])
        self.assertEqual(report["exit_hint"], 1)


class TestAgentMode(unittest.TestCase):
    def tearDown(self):
        os.environ.pop("COURSE_C50_AGENT_MODE", None)
        os.environ.pop("CLASSROOM50_ORG_ALLOWLIST", None)

    def test_t11_empty_allowlist(self):
        os.environ["COURSE_C50_AGENT_MODE"] = "1"
        os.environ["CLASSROOM50_ORG_ALLOWLIST"] = ""
        with self.assertRaises(Classroom50Error) as ctx:
            list_classrooms(
                "ORG",
                runner=lambda a: RunResult(0, "[]"),
            )
        self.assertEqual(ctx.exception.code, "allowlist_required")

    def test_t12_agent_download_refused(self):
        os.environ["COURSE_C50_AGENT_MODE"] = "1"
        with self.assertRaises(Classroom50Error):
            agent_refuse_download()

    def test_t13_org_not_allowlisted(self):
        os.environ["COURSE_C50_AGENT_MODE"] = "1"
        os.environ["CLASSROOM50_ORG_ALLOWLIST"] = "other-org"
        with self.assertRaises(Classroom50Error) as ctx:
            list_classrooms("ORG", runner=lambda a: RunResult(0, "[]"))
        self.assertEqual(ctx.exception.code, "org_not_allowlisted")

    def test_t12_human_download_requires_params(self):
        human = HumanCLI(runner=lambda a: RunResult(0))
        with self.assertRaises(Classroom50Error):
            human.download("o", "c", "", "dest")


class TestFlags(unittest.TestCase):
    def test_t8_flags_registered(self):
        # lightweight: ensure register adds destinations
        import argparse

        from course_hoanganhduc.c50_flags import register_classroom50_flags

        p = argparse.ArgumentParser(allow_abbrev=False)
        register_classroom50_flags(p)
        dests = {a.dest for a in p._actions}
        for d in (
            "sync_classroom50",
            "classroom50_org",
            "classroom50_classroom",
            "classroom50_assignment",
            "list_classroom50_classrooms",
            "list_classroom50_roster",
            "list_classroom50_assignments",
            "export_classroom50_roster",
            "classroom50_report",
            "download_classroom50",
            "classroom50_download_dest",
        ):
            self.assertIn(d, dests)
        self.assertNotIn("push_classroom50_roster", dests)


class TestSyncOffline(unittest.TestCase):
    def test_t10_pipeline(self):
        remote = [
            {
                "username": "alice",
                "github_id": 9,
                "email": "a@x.com",
                "first_name": "A",
                "last_name": "L",
                "section": "1",
            }
        ]
        runner = _mock_runner(
            {
                ("gh", "teacher", "roster", "list", "ORG", "CLS", "--json"): RunResult(
                    0, json.dumps(remote)
                )
            }
        )
        students = [Student(**{"Email": "a@x.com", "Name": "A L"})]
        out, report = __import__(
            "course_hoanganhduc.c50_sync", fromlist=["sync_pull"]
        ).sync_pull(students, org="ORG", classroom="CLS", runner=runner, agent_mode=False)
        self.assertEqual(getattr(out[0], "GitHub Username"), "alice")
        self.assertEqual(str(getattr(out[0], "GitHub ID")), "9")
        csv_text, _ = export_roster_csv(out)
        self.assertIn("alice", csv_text)


class TestDataMapping(unittest.TestCase):
    def test_github_id_column_maps_to_id_not_username(self):
        path = os.path.join(REPO_ROOT, "course_hoanganhduc", "data.py")
        with open(path, encoding="utf-8") as fh:
            src = fh.read()
        self.assertIn('"GitHub id": "GitHub ID"', src)
        self.assertIn('"Github id": "GitHub ID"', src)
        self.assertNotIn('"GitHub id": "GitHub Username"', src)



if __name__ == "__main__":
    unittest.main()
