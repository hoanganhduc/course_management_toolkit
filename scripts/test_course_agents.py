#!/usr/bin/env python3
"""Offline tests for canvas/gclass/db agent entrypoints."""

from __future__ import annotations

import argparse
import json
import os
import pickle
import subprocess
import sys
import tempfile
import types
import unittest
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)


def _stub_heavy_deps():
    def ensure(name, attrs=None):
        if name in sys.modules:
            return
        m = types.ModuleType(name)
        if attrs:
            for k, v in attrs.items():
                setattr(m, k, v)
        sys.modules[name] = m

    ensure("openpyxl")
    ensure("openpyxl.styles", {"Alignment": object})
    ensure("pandas")
    ensure("canvasapi", {"Canvas": object})
    ensure("googleapiclient")
    ensure("googleapiclient.discovery", {"build": lambda *a, **k: None})
    ensure("googleapiclient.http", {"MediaIoBaseDownload": object})
    ensure("googleapiclient.errors", {"HttpError": Exception})
    ensure("google")
    ensure("google.oauth2")
    ensure("google.oauth2.credentials", {"Credentials": object})
    ensure("google_auth_oauthlib")
    ensure("google_auth_oauthlib.flow", {"InstalledAppFlow": object})
    ensure("google.auth")
    ensure("google.auth.transport")
    ensure("google.auth.transport.requests", {"Request": object})
    ensure("paddleocr", {"PaddleOCR": object})
    ensure("pytesseract")
    ensure("pdf2image", {"convert_from_path": lambda *a, **k: []})
    ensure("PIL", {"Image": object, "ImageOps": object, "ImageFilter": object})
    ensure("PyPDF2")
    ensure("numpy")
    ensure("cv2")
    ensure("sklearn")
    ensure("sklearn.feature_extraction")
    ensure("sklearn.feature_extraction.text", {"TfidfVectorizer": object})
    ensure("sklearn.metrics")
    ensure("sklearn.metrics.pairwise", {"cosine_similarity": lambda *a, **k: None})
    ensure("tqdm", {"tqdm": lambda x, **k: x})


_stub_heavy_deps()

from course_hoanganhduc.c50_agent import main as c50_main  # noqa: E402
from course_hoanganhduc.canvas_agent import main as canvas_main  # noqa: E402
from course_hoanganhduc.course_agent_common import (  # noqa: E402
    CourseAgentError,
    require_env_allowlist,
)
from course_hoanganhduc.db_agent import main as db_main  # noqa: E402
from course_hoanganhduc.gclass_agent import main as gclass_main  # noqa: E402
from course_hoanganhduc.models import Student  # noqa: E402


def _refused_verbs(main_fn):
    """Subcommands whose help marks them refused, read off the real parser.

    The parsers are built inside main(), so they are captured on the way past
    rather than rebuilt here -- a copy would drift from what ships.
    """

    parsers = []
    original_init = argparse.ArgumentParser.__init__

    def capture_init(self, *args, **kwargs):
        original_init(self, *args, **kwargs)
        parsers.append(self)

    argparse.ArgumentParser.__init__ = capture_init
    try:
        with redirect_stderr(StringIO()):
            try:
                main_fn([])  # no subcommand: exits after the parser is built
            except SystemExit:
                pass
    finally:
        argparse.ArgumentParser.__init__ = original_init
    if not parsers:
        raise RuntimeError(f"no parser captured from {main_fn}")
    refused = []
    for action in parsers[0]._actions:
        if isinstance(action, argparse._SubParsersAction):
            for choice in action._choices_actions:
                if "(refused)" in (choice.help or ""):
                    refused.append(choice.dest)
    return refused


class TestRefusalsAreComplete(unittest.TestCase):
    """Every verb advertised as refused has to actually refuse.

    c50_agent refuses `download` on its own branch, outside REFUSED_VERBS, so the
    two lists can drift; reading the advertised set off the parser catches that.
    """

    def test_every_refused_subcommand_refuses(self):
        for name, main_fn in (
            ("c50_agent", c50_main),
            ("gclass_agent", gclass_main),
            ("db_agent", db_main),
        ):
            verbs = _refused_verbs(main_fn)
            self.assertTrue(verbs, f"{name} advertises no refused verbs")
            for verb in verbs:
                with self.subTest(agent=name, verb=verb):
                    err = StringIO()
                    with redirect_stderr(err), redirect_stdout(StringIO()):
                        code = main_fn([verb])
                    self.assertEqual(code, 1)
                    self.assertIn("is not available", err.getvalue())


def _c50_subcommands():
    """Every c50_agent subcommand, with its options, read off the real parser."""
    parsers = []
    original_init = argparse.ArgumentParser.__init__

    def capture_init(self, *args, **kwargs):
        original_init(self, *args, **kwargs)
        parsers.append(self)

    argparse.ArgumentParser.__init__ = capture_init
    try:
        with redirect_stderr(StringIO()):
            try:
                c50_main([])
            except SystemExit:
                pass
    finally:
        argparse.ArgumentParser.__init__ = original_init
    found = {}
    for action in parsers[0]._actions:
        if isinstance(action, argparse._SubParsersAction):
            for name, sub in action.choices.items():
                options = set()
                for opt_action in sub._actions:
                    options.update(opt_action.option_strings)
                help_text = ""
                for choice in action._choices_actions:
                    if choice.dest == name:
                        help_text = choice.help or ""
                found[name] = (options, help_text)
    return found


class TestClassroom50AgentSurface(unittest.TestCase):
    """The Classroom50 agent entrypoint reads; it does not write.

    The three new lanes each added a read verb here, so what these tests pin is
    the boundary itself rather than the three names: a verb that reaches the
    network has to pass the org allowlist first, and the refusals that were
    advertised before are still advertised.
    """

    def setUp(self):
        self._saved = {
            key: os.environ.get(key)
            for key in ("CLASSROOM50_ORG_ALLOWLIST", "COURSE_C50_AGENT_MODE")
        }

    def tearDown(self):
        for key, value in self._saved.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value

    def test_every_org_taking_verb_stops_without_an_allowlist(self):
        """require_org_allowlist is the one gate keeping an agent in its org.

        A new verb that reaches gh without passing through it is a way around
        that gate, so the check is derived from the parser: any subcommand
        asking for an --org is required to refuse when no org is allowed.

        subprocess.run is replaced for the duration, which does two jobs.  It
        keeps the test offline, and it turns a bypassed gate into a loud
        failure: a verb that gets as far as calling gh raises here instead of
        failing later for some unrelated reason -- a 404 from an org that does
        not exist would otherwise look exactly like the gate working.
        """
        os.environ["CLASSROOM50_ORG_ALLOWLIST"] = ""

        def forbidden(*args, **kwargs):
            raise AssertionError("a gh call escaped the org allowlist gate")

        saved_run = subprocess.run
        subprocess.run = forbidden
        try:
            for name, (options, help_text) in sorted(_c50_subcommands().items()):
                if "--org" not in options or "(refused)" in help_text:
                    continue
                with self.subTest(verb=name):
                    argv = [name, "--org", "some-org"]
                    if "--classroom" in options:
                        argv += ["--classroom", "c"]
                    err = StringIO()
                    with redirect_stderr(err), redirect_stdout(StringIO()):
                        code = c50_main(argv)
                    self.assertEqual(code, 1, f"{name} ran with no org allowlist")
                    self.assertIn("CLASSROOM50_ORG_ALLOWLIST", err.getvalue())
        finally:
            subprocess.run = saved_run

    def test_the_new_read_verbs_reach_their_operation(self):
        """Wired end to end: the verb parses, and the read layer is called.

        The operations are stubbed because the real ones call gh; what is under
        test is the dispatch, not the network.
        """
        from course_hoanganhduc import c50_ops, c50_scores

        os.environ["CLASSROOM50_ORG_ALLOWLIST"] = "some-org"
        calls = []

        def fake_import_scores(students, **kwargs):
            calls.append(("import-scores", kwargs))
            return {"matched": []}, ""

        def fake_list_groups(org, classroom, **kwargs):
            calls.append(("list-groups", kwargs))
            return {"assignments": {}}

        saved = (c50_ops.import_scores, c50_ops.list_groups, c50_scores.format_merge)
        c50_ops.import_scores = fake_import_scores
        c50_ops.list_groups = fake_list_groups
        c50_scores.format_merge = lambda report: "stub report"
        try:
            with tempfile.TemporaryDirectory() as temporary:
                missing_db = os.path.join(temporary, "absent.db")
                with redirect_stdout(StringIO()), redirect_stderr(StringIO()):
                    scores_code = c50_main(
                        [
                            "import-scores",
                            "--org",
                            "some-org",
                            "--classroom",
                            "c",
                            "--db",
                            missing_db,
                            "--dry-run",
                        ]
                    )
                    groups_code = c50_main(
                        ["list-groups", "--org", "some-org", "--classroom", "c"]
                    )
        finally:
            c50_ops.import_scores, c50_ops.list_groups, c50_scores.format_merge = saved

        self.assertEqual(scores_code, 0)
        self.assertEqual(groups_code, 0)
        self.assertEqual([name for name, _ in calls], ["import-scores", "list-groups"])

    def test_dry_run_import_scores_does_not_save(self):
        """--dry-run reports and stops; the database is the operator's to change."""
        from course_hoanganhduc import c50_ops, c50_scores, data

        os.environ["CLASSROOM50_ORG_ALLOWLIST"] = "some-org"
        saves = []
        saved = (c50_ops.import_scores, c50_scores.format_merge, data.save_database)
        c50_ops.import_scores = lambda students, **kwargs: ({"matched": []}, "")
        c50_scores.format_merge = lambda report: "stub report"
        data.save_database = lambda *a, **k: saves.append(a)
        try:
            with tempfile.TemporaryDirectory() as temporary:
                missing_db = os.path.join(temporary, "absent.db")
                with redirect_stdout(StringIO()), redirect_stderr(StringIO()):
                    code = c50_main(
                        [
                            "import-scores",
                            "--org",
                            "some-org",
                            "--classroom",
                            "c",
                            "--db",
                            missing_db,
                            "--dry-run",
                        ]
                    )
        finally:
            c50_ops.import_scores, c50_scores.format_merge, data.save_database = saved
        self.assertEqual(code, 0)
        self.assertEqual(saves, [])

    def test_the_advertised_refusals_did_not_shrink(self):
        """Adding read verbs must not quietly drop a refusal.

        The names are written out here rather than read from REFUSED_VERBS on
        purpose.  Dropping a name from that tuple unregisters its subparser, so
        the verb stops being advertised as refused and TestRefusalsAreComplete
        -- which reads the advertised set off the parser -- simply checks one
        thing less.  Deriving the expected set from the module under test would
        make this assertion agree with whatever the module currently says.
        Adding a refusal is fine and does not fail here; losing one does.
        """
        expected = {
            "download",
            "assignment-add",
            "assignment-remove",
            "invite",
            "roster-import",
            "roster-add",
        }
        advertised = set(_refused_verbs(c50_main))
        missing = sorted(expected - advertised)
        self.assertEqual(missing, [], f"no longer advertised as refused: {missing}")


class TestDbAgent(unittest.TestCase):
    def setUp(self):
        self.td = tempfile.TemporaryDirectory()
        self.db = os.path.join(self.td.name, "students.db")
        students = [
            Student(**{"Name": "Alice", "Email": "a@gmail.com", "Student ID": "1"}),
            Student(**{"Name": "Bob", "Email": "b@uni.edu", "Student ID": "2"}),
            Student(**{"Name": "Alice", "Email": "a2@x.com", "Student ID": "3"}),
        ]
        with open(self.db, "wb") as fh:
            pickle.dump(students, fh)

    def tearDown(self):
        self.td.cleanup()

    def test_count_search_export(self):
        self.assertEqual(db_main(["count", "--db", self.db]), 0)
        self.assertEqual(db_main(["search", "alice", "--db", self.db]), 0)
        self.assertEqual(db_main(["list-email-domain", "gmail.com", "--db", self.db]), 0)
        self.assertEqual(db_main(["list-duplicate-names", "--db", self.db]), 0)
        out = os.path.join(self.td.name, "e.txt")
        self.assertEqual(db_main(["export-emails", "--db", self.db, "--out", out]), 0)
        self.assertTrue(os.path.exists(out))

    def test_refuse_modify(self):
        self.assertEqual(db_main(["modify"]), 1)


class TestCanvasGclassAgent(unittest.TestCase):
    def test_refuse_destructive(self):
        self.assertEqual(canvas_main(["unenroll"]), 1)
        self.assertEqual(canvas_main(["grade"]), 1)
        self.assertEqual(canvas_main(["download"]), 1)
        self.assertEqual(gclass_main(["unenroll"]), 1)
        self.assertEqual(gclass_main(["download"]), 1)
        self.assertEqual(gclass_main(["grade"]), 1)
        self.assertEqual(gclass_main(["create-assignment"]), 1)
        self.assertEqual(gclass_main(["create-coursework"]), 1)
        self.assertEqual(gclass_main(["create"]), 1)

    def test_allowlist_required(self):
        os.environ["COURSE_AGENT_MODE"] = "1"
        os.environ["CANVAS_COURSE_ALLOWLIST"] = ""
        with self.assertRaises(CourseAgentError):
            require_env_allowlist("CANVAS_COURSE_ALLOWLIST", "1", label="c")
        os.environ.pop("COURSE_AGENT_MODE", None)
        os.environ.pop("CANVAS_COURSE_ALLOWLIST", None)

    def test_canvas_preflight_loads_explicit_private_config(self):
        with tempfile.TemporaryDirectory() as temporary:
            config = os.path.join(temporary, "canvas.json")
            with open(config, "w", encoding="utf-8") as stream:
                json.dump(
                    {
                        "CANVAS_LMS_API_URL": "https://canvas.example.invalid",
                        "CANVAS_LMS_API_KEY": "fixture-secret",
                        "CANVAS_LMS_COURSE_ID": "42",
                    },
                    stream,
                )
            os.chmod(config, 0o600)
            previous = os.environ.get("CANVAS_CONFIG_PATH")
            os.environ["CANVAS_CONFIG_PATH"] = config
            output = StringIO()
            try:
                with redirect_stdout(output):
                    result = canvas_main(["preflight"])
            finally:
                if previous is None:
                    os.environ.pop("CANVAS_CONFIG_PATH", None)
                else:
                    os.environ["CANVAS_CONFIG_PATH"] = previous
            self.assertEqual(result, 0)
            status = json.loads(output.getvalue())
            self.assertEqual(
                status,
                {
                    "ok": True,
                    "api_url_set": True,
                    "api_key_set": True,
                    "course_id_set": True,
                },
            )
            self.assertNotIn("fixture-secret", output.getvalue())

    def test_canvas_config_rejects_public_mode_and_symlink(self):
        with tempfile.TemporaryDirectory() as temporary:
            config = os.path.join(temporary, "canvas.json")
            with open(config, "w", encoding="utf-8") as stream:
                stream.write('{"CANVAS_LMS_API_KEY":"fixture"}\n')
            previous = os.environ.get("CANVAS_CONFIG_PATH")
            try:
                os.chmod(config, 0o644)
                os.environ["CANVAS_CONFIG_PATH"] = config
                self.assertEqual(canvas_main(["preflight"]), 1)
                link = os.path.join(temporary, "canvas-link.json")
                os.symlink(config, link)
                os.chmod(config, 0o600)
                os.environ["CANVAS_CONFIG_PATH"] = link
                self.assertEqual(canvas_main(["preflight"]), 1)
            finally:
                if previous is None:
                    os.environ.pop("CANVAS_CONFIG_PATH", None)
                else:
                    os.environ["CANVAS_CONFIG_PATH"] = previous


if __name__ == "__main__":
    unittest.main()
