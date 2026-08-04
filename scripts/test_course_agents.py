#!/usr/bin/env python3
"""Offline tests for canvas/gclass/db agent entrypoints."""

from __future__ import annotations

import json
import os
import pickle
import sys
import tempfile
import types
import unittest
from contextlib import redirect_stdout
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

from course_hoanganhduc.canvas_agent import main as canvas_main  # noqa: E402
from course_hoanganhduc.course_agent_common import (  # noqa: E402
    CourseAgentError,
    require_env_allowlist,
)
from course_hoanganhduc.db_agent import main as db_main  # noqa: E402
from course_hoanganhduc.gclass_agent import main as gclass_main  # noqa: E402
from course_hoanganhduc.models import Student  # noqa: E402


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
