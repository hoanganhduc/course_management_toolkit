#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Offline Classroom50 unit tests (ADR v2.1 T1–T15 subset)."""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
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
    partition_roster_candidates,
    parse_csv_text,
    parse_roster_payload,
)
from course_hoanganhduc.models import Student



def _known_config_keys():
    """The ``known_keys`` whitelist from ``load_config``, read as source.

    Importing ``config`` would pull in pandas, openpyxl and the OCR stack, none
    of which this offline suite carries.  ``config.py`` also carries an invalid
    ``\\c`` escape in a Windows-path docstring, present since the first commit,
    so parsing it emits a SyntaxWarning that is not this test's business.
    """
    import ast
    import warnings

    source = os.path.join(REPO_ROOT, "course_hoanganhduc", "config.py")
    with open(source, encoding="utf-8") as handle:
        text = handle.read()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", SyntaxWarning)
        tree = ast.parse(text, filename=source)
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id == "known_keys" for t in node.targets
        ):
            return set(ast.literal_eval(node.value))
    raise AssertionError("known_keys literal not found in config.py")


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

    ROSTER_FIXTURE = [
        Student(
            **{
                "GitHub Username": "@alice",
                "GitHub ID": "7",
                "Email": "a@x.edu",
                "Name": "Alice Nguyen",
                "Section": "S1",
            }
        ),
        Student(**{"Name": "Bao Tran", "Email": "bao@x.edu"}),
        Student(
            **{
                "GitHub Username": "cara",
                "First Name": "Cara",
                "Last Name": "Cruz",
                "Email": "cara@x.edu",
                "Section": "S2",
            }
        ),
    ]

    def test_partition_splits_on_username_and_keeps_order(self):
        exportable, missing = partition_roster_candidates(self.ROSTER_FIXTURE)
        self.assertEqual([s.Email for s in exportable], ["a@x.edu", "cara@x.edu"])
        self.assertEqual([s.Email for s in missing], ["bao@x.edu"])
        self.assertEqual(
            len(exportable) + len(missing), len(self.ROSTER_FIXTURE)
        )

    def test_export_roster_csv_output_is_unchanged(self):
        # Pinned byte for byte: partition_roster_candidates was factored out of this
        # function, and --export-classroom50-roster must keep emitting the same file.
        text, mode = export_roster_csv(self.ROSTER_FIXTURE)
        self.assertEqual(
            text,
            "username,first_name,last_name,email,section,github_id\n"
            "alice,Alice,Nguyen,a@x.edu,S1,7\n"
            "cara,Cara,Cruz,cara@x.edu,S2,\n",
        )
        self.assertEqual(mode, "heuristic")

    def test_export_header_matches_the_pinned_cli_contract(self):
        # gh teacher ParseImportCSV wants exactly these six columns, in this order, with
        # no BOM and no surrounding whitespace. Accepted on both the v1.25.1 this was
        # written against and the v1.45.0 now installed, where it is the canonical
        # seven-column header without the trailing `role`.
        text, _ = export_roster_csv(self.ROSTER_FIXTURE)
        header = text.splitlines()[0]
        self.assertEqual(
            header, "username,first_name,last_name,email,section,github_id"
        )
        self.assertEqual(header.split(","), CANONICAL_COLUMNS)
        self.assertFalse(text.encode("utf-8").startswith(b"\xef\xbb\xbf"))

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

    def test_t8b_org_and_classroom_resolve_flag_then_config_then_env(self):
        """Each tier falls through, so a config file cannot shadow the environment.

        The original code returned ``config.get(...)`` as soon as ``config`` was
        a non-empty dict, which killed the environment tier for anyone who had a
        config file at all -- and the config tier itself never fired, because
        ``load_config`` was dropping the key.
        """
        import argparse

        from course_hoanganhduc.c50_flags import _resolve_classroom, _resolve_org

        blank = argparse.Namespace(classroom50_org=None, classroom50_classroom=None)
        flagged = argparse.Namespace(
            classroom50_org="from-flag", classroom50_classroom="from-flag-class"
        )
        cfg = {"CLASSROOM50_ORG": "from-cfg", "CLASSROOM50_CLASSROOM": "from-cfg-class"}
        # A config carrying unrelated keys must not stand in for a C50 answer.
        unrelated = {"COURSE_CODE": "mat3508"}

        saved = {k: os.environ.get(k) for k in ("CLASSROOM50_ORG", "CLASSROOM50_CLASSROOM")}
        os.environ["CLASSROOM50_ORG"] = "from-env"
        os.environ["CLASSROOM50_CLASSROOM"] = "from-env-class"
        try:
            self.assertEqual(_resolve_org(flagged, cfg), "from-flag")
            self.assertEqual(_resolve_classroom(flagged, cfg), "from-flag-class")

            self.assertEqual(_resolve_org(blank, cfg), "from-cfg")
            self.assertEqual(_resolve_classroom(blank, cfg), "from-cfg-class")

            self.assertEqual(_resolve_org(blank, unrelated), "from-env")
            self.assertEqual(_resolve_classroom(blank, unrelated), "from-env-class")
            self.assertEqual(_resolve_org(blank, {}), "from-env")
            self.assertEqual(_resolve_org(blank, None), "from-env")
        finally:
            for key, value in saved.items():
                if value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = value

    def test_t8c_the_classroom50_keys_survive_the_load_config_whitelist(self):
        """``load_config`` keeps only ``known_keys``; the C50 pair must be in it."""
        known = _known_config_keys()
        self.assertIn("CLASSROOM50_ORG", known)
        self.assertIn("CLASSROOM50_CLASSROOM", known)

    def test_t8d_no_resolver_reads_a_key_load_config_would_drop(self):
        """Every config key the C50 lane reads has to be whitelisted, not just these two."""
        flags = os.path.join(REPO_ROOT, "course_hoanganhduc", "c50_flags.py")
        with open(flags, encoding="utf-8") as handle:
            read = set(re.findall(r'config\.get\(\s*"([A-Za-z0-9_]+)"', handle.read()))
        self.assertEqual(
            sorted(read - _known_config_keys()),
            [],
            "config keys read but not whitelisted",
        )


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


class TestOperatorSurfaceRefusals(unittest.TestCase):
    """The mutating verbs stay off the agent entrypoint (ADR D6)."""

    def tearDown(self):
        os.environ.pop("COURSE_C50_AGENT_MODE", None)

    def test_agent_entrypoint_refuses_mutating_verbs(self):
        import contextlib
        import io

        from course_hoanganhduc import c50_agent

        for argv in (
            ["assignment-add", "--org", "O"],
            ["assignment-remove", "--org", "O"],
            ["invite", "--target", "O"],
            ["download", "--org", "O"],
        ):
            with self.subTest(verb=argv[0]):
                err = io.StringIO()
                with contextlib.redirect_stderr(err):
                    code = c50_agent.main(argv)
                self.assertEqual(code, 1)
                self.assertIn("Classroom50 error", err.getvalue())

    def test_agent_refuse_code(self):
        from course_hoanganhduc.c50_ops import agent_refuse

        with self.assertRaises(Classroom50Error) as ctx:
            agent_refuse("invite")
        self.assertEqual(ctx.exception.code, "agent_forbidden")

    def test_human_download_still_refuses_unverified_by_pattern(self):
        human = HumanCLI(runner=lambda a: RunResult(0))
        with self.assertRaises(Classroom50Error) as ctx:
            human.download("o", "c", "a", "dest", by_pattern=True)
        self.assertEqual(ctx.exception.code, "by_pattern_not_permitted")

    def test_plain_human_download_argv_unchanged(self):
        human = HumanCLI(runner=lambda a: RunResult(0))
        self.assertEqual(
            human.download_argv("o", "c", "a", "dest"),
            ["gh", "teacher", "download", "o", "c", "a", "-d", "dest"],
        )


class TestRosterImportContract(unittest.TestCase):
    """`gh teacher roster import --help`, read verbatim.

    The note at c50_roster.py:94-98 says re-emitting a remote row would erase
    its role.  The pinned v1.45.0 says the opposite, so this reads the help
    text -- which changes nothing -- rather than trusting either comment.
    """

    @classmethod
    def setUpClass(cls):
        if shutil.which("gh") is None:
            raise unittest.SkipTest("gh is not installed")
        try:
            proc = subprocess.run(
                ["gh", "teacher", "roster", "import", "--help"],
                capture_output=True,
                text=True,
                timeout=30,
            )
        except (OSError, subprocess.SubprocessError) as exc:  # pragma: no cover
            raise unittest.SkipTest(f"gh teacher is not available: {exc}")
        if proc.returncode != 0:
            raise unittest.SkipTest("gh teacher extension is not installed")
        # The help text is hard-wrapped; compare on a single line.
        cls.help = re.sub(r"\s+", " ", proc.stdout)

    def test_seven_six_and_five_column_headers_are_all_valid(self):
        self.assertIn(
            "`" + ",".join(CANONICAL_COLUMNS) + ",role`",
            self.help,
        )
        self.assertIn("the same without `role`", self.help)
        self.assertIn("just the first five columns", self.help)

    def test_the_email_column_may_be_empty_per_row(self):
        self.assertIn("The `email` column may be empty per row.", self.help)

    def test_github_id_is_re_resolved_from_the_username(self):
        self.assertIn(
            "github_id is re-resolved from `GET /users/{username}`", self.help
        )
        self.assertIn("names a different account than the username fails that line", self.help)

    def test_role_is_carried_and_never_overwritten(self):
        self.assertIn("`role` is carried, never applied", self.help)
        self.assertIn("an already-recorded role is never overwritten", self.help)

    def test_a_row_without_a_username_is_skipped(self):
        self.assertIn("A row with a github_id but no username is skipped", self.help)


if __name__ == "__main__":
    unittest.main()
