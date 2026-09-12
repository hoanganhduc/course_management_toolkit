#!/usr/bin/env python3
"""Offline tests for the ``known_keys`` allowlist in ``config.load_config``.

``load_config`` copies a config file into a dict one key at a time, and only for keys
named in ``known_keys``.  Anything absent from that list is dropped without a word, so a
value the operator put in ``config.json`` can silently never reach the code that reads it.
These tests pin the allowlist to its two sources of truth: the docstring that says which
keys the function handles, and the module globals that ``_apply_config_overrides`` can
assign to.

Config files here are written into a ``tempfile`` tree and passed to ``load_config`` by
path, so nothing reads the operator's real ``%APPDATA%``.
"""

from __future__ import annotations

import ast
import json
import os
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from typing import List, Set

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc import data, settings  # noqa: E402
from course_hoanganhduc.config import load_config  # noqa: E402
from course_hoanganhduc.core import _apply_config_overrides  # noqa: E402

CONFIG_PY = os.path.join(REPO_ROOT, "course_hoanganhduc", "config.py")
SETTINGS_PY = os.path.join(REPO_ROOT, "course_hoanganhduc", "settings.py")

# settings.py holds the fallback course code, not a value a course config overrides.
NOT_CONFIGURABLE = {"DEFAULT_COURSE_CODE"}


def read_source(path: str) -> str:
    with open(path, encoding="utf-8") as f:
        return f.read()


def read_known_keys() -> Set[str]:
    """The allowlist as written, parsed from the source rather than re-typed here."""
    body = read_source(CONFIG_PY).split("known_keys = [", 1)[1].split("]", 1)[0]
    return set(ast.literal_eval("[" + body + "]"))


def read_docstring_keys() -> List[str]:
    """The keys ``load_config``'s docstring claims the function handles."""
    tree = ast.parse(read_source(CONFIG_PY))
    for node in ast.walk(tree):
        if isinstance(node, ast.FunctionDef) and node.name == "load_config":
            doc = ast.get_docstring(node) or ""
            break
    else:
        raise AssertionError("load_config not found in config.py")
    return [
        line.strip()
        for line in doc.splitlines()
        if line.strip().isidentifier() and line.strip().isupper()
    ]


def read_settings_globals() -> Set[str]:
    """Uppercase module-level names in settings.py, the ones overrides can assign to."""
    tree = ast.parse(read_source(SETTINGS_PY))
    names = set()
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id.isupper():
                    names.add(target.id)
    return names


def load_quietly(config_path):
    """``load_config`` narrates to stdout; the transcript is not what is under test."""
    with redirect_stdout(StringIO()):
        return load_config(config_path)


class AllowlistTests(unittest.TestCase):
    def setUp(self):
        self.known = read_known_keys()

    def test_every_key_the_docstring_promises_is_actually_copied(self):
        promised = read_docstring_keys()
        self.assertGreater(len(promised), 20, "docstring key list failed to parse")
        missing = sorted(k for k in promised if k not in self.known)
        self.assertEqual(missing, [], "documented but dropped by load_config: {}".format(missing))

    def test_every_overridable_setting_is_copied(self):
        # _apply_config_overrides assigns a key onto a module only when the module already
        # has that attribute, so a settings.py global absent from known_keys can never be
        # set from a config file.
        missing = sorted(
            k for k in read_settings_globals() if k not in self.known and k not in NOT_CONFIGURABLE
        )
        self.assertEqual(missing, [], "settable in settings.py but dropped: {}".format(missing))

    def test_the_three_canvas_assignment_ids_are_treated_alike(self):
        # The midterm and CC ids were allowlisted and the final one was not, which is how
        # the omission stayed invisible: two thirds of the family worked.
        for key in ("CANVAS_CC_ASSIGNMENT_ID", "CANVAS_MIDTERM_ASSIGNMENT_ID",
                    "CANVAS_FINAL_ASSIGNMENT_ID"):
            self.assertIn(key, self.known)


class RoundTripTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.path = os.path.join(self._tmp.name, "config.json")

    def write(self, values):
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(values, f)
        return self.path

    def test_final_assignment_id_survives_the_load(self):
        loaded = load_quietly(self.write({
            "COURSE_CODE": "MAT3397",
            "CANVAS_MIDTERM_ASSIGNMENT_ID": "111",
            "CANVAS_FINAL_ASSIGNMENT_ID": "222",
        }))
        self.assertEqual(loaded.get("CANVAS_MIDTERM_ASSIGNMENT_ID"), "111")
        self.assertEqual(loaded.get("CANVAS_FINAL_ASSIGNMENT_ID"), "222")

    def test_an_unknown_key_is_still_dropped(self):
        # The allowlist is doing its job for real typos; this change widens it by one key,
        # it does not turn load_config into a pass-through.
        loaded = load_quietly(self.write({"COURSE_CODE": "MAT3397", "CANVAS_FINAL_ASSIGNMNET_ID": "222"}))
        self.assertNotIn("CANVAS_FINAL_ASSIGNMNET_ID", loaded)


class OverrideTests(unittest.TestCase):
    """The whole chain: config file -> load_config -> module global that data.py reads."""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        for module in (settings, data):
            self.addCleanup(
                setattr, module, "CANVAS_FINAL_ASSIGNMENT_ID",
                getattr(module, "CANVAS_FINAL_ASSIGNMENT_ID"),
            )

    def test_the_value_reaches_the_module_that_grades_with_it(self):
        self.assertEqual(data.CANVAS_FINAL_ASSIGNMENT_ID, "",
                         "test needs to start from the unset default")
        path = os.path.join(self._tmp.name, "config.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"COURSE_CODE": "MAT3397", "CANVAS_FINAL_ASSIGNMENT_ID": "222"}, f)

        _apply_config_overrides(load_quietly(path))

        # data.py:2291 reads this global to fall back to the Canvas final assignment when a
        # student has no CK score; at "" that branch is dead.
        self.assertEqual(data.CANVAS_FINAL_ASSIGNMENT_ID, "222")


if __name__ == "__main__":
    unittest.main(verbosity=2)
