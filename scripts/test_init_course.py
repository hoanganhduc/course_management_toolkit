#!/usr/bin/env python3
"""Offline tests for course initialization.

Every outward call is injected: the three platform listings go through a fake
``Discovery``, the wizard reads from a scripted list instead of stdin, and the config
directory is a ``tempfile`` tree passed in as ``base_dir``.  So nothing here reaches the
network, the real ``%APPDATA%``, or the operator's terminal.
"""

from __future__ import annotations

import json
import os
import pickle
import sys
import tempfile
import unittest
from typing import Any, Dict, List, Optional

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc import init_course  # noqa: E402
from course_hoanganhduc.init_course import (  # noqa: E402
    ACCOUNT_WIDE_KEYS,
    InitCourseError,
    course_id_from_classroom_url,
    find_source_courses,
    inherit_values,
    normalize_course_code,
    normalize_sheet_url,
    read_config_file,
    redact,
    run_init_course,
    scaffold_course_folder,
    summarize,
)

CANVAS_KEY = "canvas-secret-0123456789"
GEMINI_KEY = "gemini-secret-abcdefghij"

RICH_CONFIG = {
    "CANVAS_LMS_API_URL": "https://canvas.example.edu",
    "CANVAS_LMS_API_KEY": CANVAS_KEY,
    "GEMINI_API_KEY": GEMINI_KEY,
    "GEMINI_DEFAULT_MODEL": "gemini-2.5-flash",
    "DEFAULT_AI_METHOD": "gemini",
    "ALL_AI_METHODS": ["gemini", "huggingface"],
    "DEFAULT_OCR_METHOD": "ocrspace",
    "OCRSPACE_API_KEY": "ocr-secret-123",
    "UNIVERSITY_NAME": "HUS",
    # Course-specific: must not be inherited.
    "COURSE_CODE": "MAT3397",
    "GOOGLE_CLASSROOM_COURSE_ID": "111111111111",
    "GOOGLE_SHEET_URL": "https://docs.google.com/spreadsheets/d/OLDOLDOLDOLDOLDOLDOLD/edit",
    "CLASSROOM50_ORG": "VNU-HUS",
}

POOR_CONFIG = {
    "COURSE_CODE": "MAT3508",
    "GOOGLE_CLASSROOM_COURSE_ID": "872017288939",
    "CLASSROOM50_ORG": "VNU-HUS",
}


class FakeDiscovery:
    """Stand-in for the live platform lookups; same four method names."""

    def __init__(self, google=None, canvas=None, c50=None, fail=()):
        self.google = google if google is not None else [
            {"id": "872017288939", "name": "MAT3508 - Toán rời rạc"},
            {"id": "872018493510", "name": "MAT3500 - Đồ thị"},
        ]
        self.canvas = canvas if canvas is not None else [
            {"id": "55997", "name": "Canvas MAT3508"},
        ]
        self.c50 = c50 if c50 is not None else [
            {"id": "vnu-hus-mat3508-winter-2026", "name": "MAT3508 Winter 2026"},
        ]
        self.fail = set(fail)
        self.calls: List[str] = []

    def google_courses(self, credentials_path, token_path, verbose=False):
        self.calls.append("google")
        if "google" in self.fail:
            raise RuntimeError("token hết hạn")
        return list(self.google)

    def canvas_courses(self, api_url, api_key, verbose=False):
        self.calls.append("canvas")
        if "canvas" in self.fail:
            raise RuntimeError("401 không xác thực được")
        return list(self.canvas)

    def c50_classrooms(self, org):
        self.calls.append("c50")
        if "c50" in self.fail:
            raise RuntimeError("gh teacher chưa cài")
        return list(self.c50)


class Recorder:
    """Captures the wizard transcript and feeds it scripted answers."""

    def __init__(self, answers: Optional[List[str]] = None):
        self.answers = list(answers or [])
        self.lines: List[str] = []
        self.prompts: List[str] = []

    def out(self, *parts):
        self.lines.append(" ".join(str(p) for p in parts))

    def input(self, prompt=""):
        self.prompts.append(prompt)
        if not self.answers:
            raise AssertionError("wizard asked for more input than the script supplies: "
                                 + prompt)
        return self.answers.pop(0)

    @property
    def transcript(self) -> str:
        return "\n".join(self.lines + self.prompts)


def refusing_input(prompt=""):
    raise AssertionError("stdin was read: " + prompt)


def write_source(base: str, code: str, values: Dict[str, Any],
                 credentials: bool = False) -> str:
    folder = os.path.join(base, code)
    os.makedirs(folder, exist_ok=True)
    with open(os.path.join(folder, "config.json"), "w", encoding="utf-8") as f:
        json.dump(values, f)
    if credentials:
        with open(os.path.join(folder, "credentials.json"), "w", encoding="utf-8") as f:
            f.write('{"installed": {"client_id": "fake"}}')
        with open(os.path.join(folder, "token.pickle"), "wb") as f:
            f.write(b"not a real token")
    return folder


class TempTreeTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.root = self._tmp.name
        self.base = os.path.join(self.root, "config")
        self.course_dir = os.path.join(self.root, "MAT9999")
        os.makedirs(self.base)
        os.makedirs(self.course_dir)
        self.addCleanup(self._tmp.cleanup)


class NormalizeTests(unittest.TestCase):
    def test_course_code_is_trimmed_and_upper_cased(self):
        self.assertEqual(normalize_course_code("  mat3508 "), "MAT3508")

    def test_empty_course_code_is_none(self):
        self.assertIsNone(normalize_course_code("   "))
        self.assertIsNone(normalize_course_code(None))

    def test_course_code_rejects_path_characters(self):
        for bad in ("mat 3508", "mat/3508", "..\\evil", "a:b"):
            with self.assertRaises(InitCourseError):
                normalize_course_code(bad)


class ClassroomUrlTests(unittest.TestCase):
    def test_plain_numeric_id_passes_through(self):
        self.assertEqual(course_id_from_classroom_url("872017884599"), "872017884599")

    def test_url_decodes_to_numeric_id(self):
        url = "https://classroom.google.com/c/ODcyMDE3ODg0NTk5"
        self.assertEqual(course_id_from_classroom_url(url), "872017884599")

    def test_multi_account_url_form_decodes(self):
        url = "https://classroom.google.com/u/1/c/NzI0OTU4MTk5ODM5?hl=vi"
        self.assertEqual(course_id_from_classroom_url(url), "724958199839")

    def test_non_numeric_decode_is_rejected(self):
        # Decodes cleanly to "hello world", which is not a course id.
        url = "https://classroom.google.com/c/aGVsbG8gd29ybGQ="
        self.assertIsNone(course_id_from_classroom_url(url))

    def test_garbage_is_none(self):
        self.assertIsNone(course_id_from_classroom_url("khong phai duong dan"))
        self.assertIsNone(course_id_from_classroom_url(""))
        self.assertIsNone(course_id_from_classroom_url(None))


class SheetUrlTests(unittest.TestCase):
    def test_full_url_is_kept_verbatim_including_gid(self):
        url = ("https://docs.google.com/spreadsheets/d/"
               "12eXE3ZokDlsuvx5inOyCqnEx_4ykUeW2CXrAkeVGdSI/edit?gid=351955757#gid=351955757")
        self.assertEqual(normalize_sheet_url(url), url)

    def test_bare_id_is_expanded(self):
        self.assertEqual(
            normalize_sheet_url("12eXE3ZokDlsuvx5inOyCqnEx_4ykUeW2CXrAkeVGdSI"),
            "https://docs.google.com/spreadsheets/d/"
            "12eXE3ZokDlsuvx5inOyCqnEx_4ykUeW2CXrAkeVGdSI/edit",
        )

    def test_unrelated_string_is_none(self):
        self.assertIsNone(normalize_sheet_url("https://example.com/not-a-sheet"))
        self.assertIsNone(normalize_sheet_url("short"))


class RedactTests(unittest.TestCase):
    def test_secret_keys_are_hidden(self):
        self.assertNotIn(CANVAS_KEY, str(redact("CANVAS_LMS_API_KEY", CANVAS_KEY)))
        self.assertNotIn("abc", str(redact("GEMINI_API_KEY", "abc123")))

    def test_plain_keys_are_shown(self):
        self.assertEqual(redact("COURSE_CODE", "MAT3508"), "MAT3508")
        self.assertEqual(redact("ALL_AI_METHODS", ["gemini"]), ["gemini"])


class SourceDiscoveryTests(TempTreeTest):
    def test_richest_source_comes_first(self):
        write_source(self.base, "mat3508", POOR_CONFIG)
        write_source(self.base, "mat3397", RICH_CONFIG)
        sources = find_source_courses(self.base)
        self.assertEqual([s.code for s in sources], ["mat3397", "mat3508"])
        self.assertGreater(len(sources[0].inherited_keys), len(sources[1].inherited_keys))

    def test_target_course_is_excluded(self):
        write_source(self.base, "mat3397", RICH_CONFIG)
        write_source(self.base, "mat9999", POOR_CONFIG)
        self.assertEqual([s.code for s in find_source_courses(self.base, exclude="MAT9999")],
                         ["mat3397"])

    def test_folders_without_config_are_ignored(self):
        os.makedirs(os.path.join(self.base, "google-classroom"))
        write_source(self.base, "mat3397", RICH_CONFIG)
        self.assertEqual([s.code for s in find_source_courses(self.base)], ["mat3397"])

    def test_unreadable_config_is_ignored(self):
        folder = os.path.join(self.base, "broken")
        os.makedirs(folder)
        with open(os.path.join(folder, "config.json"), "w", encoding="utf-8") as f:
            f.write("{ not json")
        self.assertEqual(find_source_courses(self.base), [])

    def test_missing_base_dir_is_empty(self):
        self.assertEqual(find_source_courses(os.path.join(self.root, "nope")), [])


class InheritTests(unittest.TestCase):
    def test_only_account_wide_keys_are_taken(self):
        values = inherit_values(RICH_CONFIG)
        self.assertIn("CANVAS_LMS_API_KEY", values)
        self.assertIn("ALL_AI_METHODS", values)
        for key in ("COURSE_CODE", "GOOGLE_CLASSROOM_COURSE_ID", "GOOGLE_SHEET_URL",
                    "CLASSROOM50_ORG"):
            self.assertNotIn(key, values)

    def test_blank_values_are_not_inherited(self):
        values = inherit_values({"CANVAS_LMS_API_KEY": "", "GEMINI_API_KEY": "  ",
                                 "UNIVERSITY_NAME": "HUS"})
        self.assertEqual(values, {"UNIVERSITY_NAME": "HUS"})

    def test_every_account_wide_key_is_course_independent(self):
        # Guards the taxonomy itself: nothing course-specific may leak into inheritance.
        for key in ("COURSE_CODE", "COURSE_NAME", "GOOGLE_CLASSROOM_COURSE_ID",
                    "CANVAS_LMS_COURSE_ID", "CLASSROOM50_CLASSROOM", "GOOGLE_SHEET_URL",
                    "MIDTERM_DATE"):
            self.assertNotIn(key, ACCOUNT_WIDE_KEYS)


class ScaffoldTests(TempTreeTest):
    EXPECTED = {
        ".course_code", "students.db",
        "course.sh", "course.bat",
        "course-c50-admin.sh", "course-c50-admin.bat",
        "course-gclass-admin.sh", "course-gclass-admin.bat",
    }

    def test_creates_the_whole_folder(self):
        result = scaffold_course_folder(self.course_dir, "MAT9999")
        self.assertEqual(set(result.created), self.EXPECTED)
        self.assertEqual(result.skipped, [])
        self.assertEqual(set(os.listdir(self.course_dir)), self.EXPECTED)

    def test_course_code_marker_is_lower_case(self):
        scaffold_course_folder(self.course_dir, "MAT9999")
        with open(os.path.join(self.course_dir, ".course_code"), encoding="utf-8") as f:
            self.assertEqual(f.read().strip(), "mat9999")

    def test_empty_database_loads_as_an_empty_roster(self):
        scaffold_course_folder(self.course_dir, "MAT9999")
        with open(os.path.join(self.course_dir, "students.db"), "rb") as f:
            self.assertEqual(pickle.load(f), [])

    def test_second_run_keeps_existing_files(self):
        scaffold_course_folder(self.course_dir, "MAT9999")
        with open(os.path.join(self.course_dir, "course.sh"), "w", encoding="utf-8") as f:
            f.write("# tuned by hand\n")
        result = scaffold_course_folder(self.course_dir, "MAT9999")
        self.assertEqual(result.created, [])
        self.assertEqual(set(result.skipped), self.EXPECTED)
        with open(os.path.join(self.course_dir, "course.sh"), encoding="utf-8") as f:
            self.assertEqual(f.read(), "# tuned by hand\n")

    def test_dry_run_writes_nothing_but_reports_everything(self):
        empty = os.path.join(self.root, "MAT0000")
        result = scaffold_course_folder(empty, "MAT0000", dry_run=True)
        self.assertEqual(set(result.created), self.EXPECTED)
        self.assertFalse(os.path.exists(empty))

    def test_wrappers_probe_both_venv_layouts(self):
        scaffold_course_folder(self.course_dir, "MAT9999")
        with open(os.path.join(self.course_dir, "course.sh"), encoding="utf-8") as f:
            body = f.read()
        self.assertIn(".course_venv/bin/course", body)
        self.assertIn(".course_venv/Scripts/course.exe", body)
        with open(os.path.join(self.course_dir, "course-c50-admin.bat"), encoding="utf-8") as f:
            self.assertIn("course-c50-admin.exe", f.read())

    def test_no_db_leaves_the_database_out(self):
        result = scaffold_course_folder(self.course_dir, "MAT9999", make_db=False)
        self.assertNotIn("students.db", result.created)


class WizardTests(TempTreeTest):
    """The interactive path: discovery-first, with guidance when discovery fails."""

    SHEET = ("https://docs.google.com/spreadsheets/d/"
             "1AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/edit?gid=7#gid=7")

    def setUp(self):
        super().setUp()
        write_source(self.base, "mat3397", RICH_CONFIG, credentials=True)
        write_source(self.base, "mat3508", POOR_CONFIG)
        self.discovery = FakeDiscovery()

    def run_wizard(self, answers, **kwargs):
        recorder = Recorder(answers)
        params = dict(
            course_code="MAT9999", course_name="Môn thử", directory=self.course_dir,
            base_dir=self.base, discovery=self.discovery,
            input_fn=recorder.input, out=recorder.out,
        )
        params.update(kwargs)
        return run_init_course(**params), recorder

    # Answers for the full happy path, in the order the wizard asks:
    #   source pick, Classroom pick, sheet URL, "use Canvas?", Canvas pick,
    #   "use Classroom50?", org (blank = suggested), classroom pick.
    HAPPY = ["1", "1", SHEET, "c", "1", "c", "", "1"]

    def test_happy_path_writes_every_field(self):
        result, _ = self.run_wizard(self.HAPPY)
        self.assertEqual(result.course_code, "MAT9999")
        self.assertEqual(result.source, "mat3397")
        written = read_config_file(result.config_path)
        self.assertEqual(written["COURSE_CODE"], "MAT9999")
        self.assertEqual(written["COURSE_NAME"], "Môn thử")
        self.assertEqual(written["GOOGLE_CLASSROOM_COURSE_ID"], "872017288939")
        self.assertEqual(written["GOOGLE_SHEET_URL"], self.SHEET)
        self.assertEqual(written["CANVAS_LMS_COURSE_ID"], "55997")
        self.assertEqual(written["CLASSROOM50_ORG"], "VNU-HUS")
        self.assertEqual(written["CLASSROOM50_CLASSROOM"], "vnu-hus-mat3508-winter-2026")

    def test_account_wide_settings_are_inherited_course_ones_are_not(self):
        result, _ = self.run_wizard(self.HAPPY)
        written = read_config_file(result.config_path)
        self.assertEqual(written["CANVAS_LMS_API_KEY"], CANVAS_KEY)
        self.assertEqual(written["GEMINI_API_KEY"], GEMINI_KEY)
        self.assertNotEqual(written["GOOGLE_CLASSROOM_COURSE_ID"],
                            RICH_CONFIG["GOOGLE_CLASSROOM_COURSE_ID"])
        self.assertNotEqual(written["GOOGLE_SHEET_URL"], RICH_CONFIG["GOOGLE_SHEET_URL"])

    def test_richest_source_is_offered_first(self):
        _, recorder = self.run_wizard(self.HAPPY)
        offered = [line for line in recorder.lines if line.startswith("  1. mat")]
        self.assertEqual(offered[0], "  1. mat3397 (9 mục dùng chung)")

    def test_credentials_are_copied_but_the_token_is_not(self):
        result, _ = self.run_wizard(self.HAPPY)
        folder = os.path.dirname(result.config_path)
        self.assertTrue(os.path.isfile(os.path.join(folder, "credentials.json")))
        self.assertFalse(os.path.exists(os.path.join(folder, "token.pickle")))

    def test_folder_is_scaffolded(self):
        result, _ = self.run_wizard(self.HAPPY)
        self.assertIn(".course_code", result.scaffold.created)
        self.assertTrue(os.path.isfile(os.path.join(self.course_dir, "students.db")))

    def test_no_secret_reaches_the_transcript(self):
        result, recorder = self.run_wizard(self.HAPPY)
        summarize(result, out=recorder.out)
        self.assertNotIn(CANVAS_KEY, recorder.transcript)
        self.assertNotIn(GEMINI_KEY, recorder.transcript)
        self.assertNotIn(RICH_CONFIG["OCRSPACE_API_KEY"], recorder.transcript)
        # The summary did run and did mention the key by name.
        self.assertIn("CANVAS_LMS_API_KEY", recorder.transcript)

    def test_question_mark_prints_guidance_and_asks_again(self):
        answers = ["1", "?", "1", self.SHEET, "k", "k"]
        _, recorder = self.run_wizard(answers)
        self.assertIn(init_course.GUIDANCE["google_classroom"][0], recorder.lines)

    def test_skipping_a_platform_is_recorded_not_fatal(self):
        answers = ["1", "0", "", "k", "k"]
        result, _ = self.run_wizard(answers)
        states = {o.key: o.state for o in result.outcomes}
        self.assertEqual(states["google_classroom"], "skipped")
        self.assertEqual(states["google_sheet"], "skipped")
        self.assertEqual(states["canvas"], "skipped")
        self.assertNotIn("GOOGLE_CLASSROOM_COURSE_ID", read_config_file(result.config_path))

    def test_pasted_classroom_url_is_decoded_and_confirmed(self):
        answers = ["1", "m", "https://classroom.google.com/c/ODcyMDE3ODg0NTk5", "c",
                   "", "k", "k"]
        result, recorder = self.run_wizard(answers)
        self.assertEqual(read_config_file(result.config_path)["GOOGLE_CLASSROOM_COURSE_ID"],
                         "872017884599")
        self.assertIn("Giải mã được ID: 872017884599", recorder.lines)

    def test_declining_the_decoded_id_asks_again(self):
        answers = ["1", "m", "https://classroom.google.com/c/ODcyMDE3ODg0NTk5", "k",
                   "872018493510", "", "k", "k"]
        result, _ = self.run_wizard(answers)
        self.assertEqual(read_config_file(result.config_path)["GOOGLE_CLASSROOM_COURSE_ID"],
                         "872018493510")

    def test_failed_listing_offers_setup_then_continues(self):
        self.discovery = FakeDiscovery(fail=["google"])
        # Classroom listing fails -> "set it up?" yes -> guidance -> "enter id now?" no.
        answers = ["1", "c", "k", "", "k", "k"]
        result, recorder = self.run_wizard(answers)
        states = {o.key: o.state for o in result.outcomes}
        self.assertEqual(states["google_classroom"], "unavailable")
        self.assertIn("token hết hạn", " ".join(recorder.lines))
        self.assertIn(init_course.GUIDANCE["google_classroom"][0], recorder.lines)
        # The rest of the run still finished and wrote a config.
        self.assertTrue(os.path.isfile(result.config_path))

    def test_empty_listing_offers_setup(self):
        self.discovery = FakeDiscovery(google=[])
        answers = ["1", "k", "", "k", "k"]
        result, recorder = self.run_wizard(answers)
        states = {o.key: o.state for o in result.outcomes}
        self.assertEqual(states["google_classroom"], "skipped")
        self.assertIn("Chưa thấy Google Classroom nào dùng được: tài khoản chưa có khoá nào",
                      recorder.lines)

    def test_bad_sheet_url_is_rejected_with_guidance(self):
        answers = ["1", "0", "https://example.com/nope", self.SHEET, "k", "k"]
        result, recorder = self.run_wizard(answers)
        self.assertIn("Chuỗi đó không giống đường dẫn Google Sheet.", recorder.lines)
        self.assertEqual(read_config_file(result.config_path)["GOOGLE_SHEET_URL"], self.SHEET)

    def test_choosing_no_source_starts_from_defaults(self):
        # No source means no credentials.json either, so Classroom cannot be listed and
        # the wizard falls through to offering setup.
        answers = ["0", "k", "", "k", "k"]
        result, recorder = self.run_wizard(answers)
        self.assertIsNone(result.source)
        self.assertNotIn("CANVAS_LMS_API_KEY", read_config_file(result.config_path))
        self.assertIn("Chưa thấy Google Classroom nào dùng được: chưa có credentials.json",
                      recorder.lines)


class FlagDrivenTests(TempTreeTest):
    def setUp(self):
        super().setUp()
        write_source(self.base, "mat3397", RICH_CONFIG, credentials=True)
        self.discovery = FakeDiscovery()

    def init(self, **kwargs):
        params = dict(
            course_code="MAT9999", directory=self.course_dir, base_dir=self.base,
            discovery=self.discovery, interactive=False, input_fn=refusing_input,
            out=lambda *a: None,
        )
        params.update(kwargs)
        return run_init_course(**params)

    def already_authorized(self):
        """Plant a token in the target config dir, as a re-run of an authorized course has."""
        folder = os.path.join(self.base, "mat9999")
        os.makedirs(folder, exist_ok=True)
        with open(os.path.join(folder, "token.pickle"), "wb") as f:
            f.write(b"not a real token")

    def test_non_interactive_never_reads_stdin(self):
        result = self.init()
        self.assertEqual(result.course_code, "MAT9999")
        self.assertTrue(os.path.isfile(result.config_path))

    def test_flags_set_every_platform_without_prompting(self):
        result = self.init(
            course_name="Môn cờ",
            google_course_id="https://classroom.google.com/c/ODcyMDE3Mjg4OTM5",
            canvas_course_id="55997",
            c50_org="VNU-HUS", c50_classroom="vnu-hus-mat3508-winter-2026",
            sheet_url="1AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
        )
        written = read_config_file(result.config_path)
        self.assertEqual(written["GOOGLE_CLASSROOM_COURSE_ID"], "872017288939")
        self.assertEqual(written["CANVAS_LMS_COURSE_ID"], "55997")
        self.assertEqual(written["CLASSROOM50_CLASSROOM"], "vnu-hus-mat3508-winter-2026")
        self.assertEqual(written["COURSE_NAME"], "Môn cờ")
        self.assertTrue(written["GOOGLE_SHEET_URL"].startswith("https://docs.google.com/"))

    def test_flag_id_absent_from_the_listing_is_reported_not_rejected(self):
        self.already_authorized()
        result = self.init(google_course_id="999999999999")
        outcome = {o.key: o for o in result.outcomes}["google_classroom"]
        self.assertEqual(outcome.state, "chosen")
        self.assertIn("không có trong danh sách", outcome.note)
        self.assertEqual(read_config_file(result.config_path)["GOOGLE_CLASSROOM_COURSE_ID"],
                         "999999999999")

    def test_flag_id_is_kept_unverified_when_google_is_not_authorized_yet(self):
        # Listing a brand-new course would have to run the OAuth flow, which blocks on a
        # browser or on a pasted redirect read from stdin. Neither is allowed here.
        result = self.init(google_course_id="872017288939")
        outcome = {o.key: o for o in result.outcomes}["google_classroom"]
        self.assertEqual(outcome.state, "chosen")
        self.assertIn("chưa uỷ quyền", outcome.note)
        self.assertNotIn("google", self.discovery.calls)
        self.assertEqual(read_config_file(result.config_path)["GOOGLE_CLASSROOM_COURSE_ID"],
                         "872017288939")

    def test_unverifiable_flag_id_is_unknown_not_failed(self):
        self.discovery = FakeDiscovery(fail=["google"])
        result = self.init(google_course_id="872017288939")
        outcome = {o.key: o for o in result.outcomes}["google_classroom"]
        self.assertEqual(outcome.state, "chosen")
        self.assertIn("không kiểm tra được", outcome.note)

    def test_explicit_inherit_from_is_honoured(self):
        write_source(self.base, "mat3508", POOR_CONFIG)
        result = self.init(inherit_from="mat3508")
        self.assertEqual(result.source, "mat3508")
        self.assertNotIn("CANVAS_LMS_API_KEY", read_config_file(result.config_path))

    def test_unknown_inherit_from_is_an_error(self):
        with self.assertRaises(InitCourseError) as caught:
            self.init(inherit_from="mat0000")
        self.assertIn("mat0000", str(caught.exception))

    def test_second_run_refuses_and_names_the_override(self):
        self.init()
        with self.assertRaises(InitCourseError) as caught:
            self.init()
        self.assertIn("--init-force", str(caught.exception))

    def test_force_overwrites(self):
        first = self.init()
        with open(first.config_path, encoding="utf-8") as f:
            self.assertNotIn("CANVAS_LMS_COURSE_ID", json.load(f))
        second = self.init(force=True, canvas_course_id="55997")
        self.assertEqual(read_config_file(second.config_path)["CANVAS_LMS_COURSE_ID"], "55997")

    def test_dry_run_writes_absolutely_nothing(self):
        before_course = sorted(os.listdir(self.course_dir))
        before_base = sorted(os.listdir(self.base))
        result = self.init(dry_run=True, canvas_course_id="55997")
        self.assertTrue(result.dry_run)
        self.assertEqual(sorted(os.listdir(self.course_dir)), before_course)
        self.assertEqual(sorted(os.listdir(self.base)), before_base)
        self.assertFalse(os.path.exists(result.config_path))
        # ...while still reporting what it would have done.
        self.assertEqual(result.values["CANVAS_LMS_COURSE_ID"], "55997")
        self.assertIn("students.db", result.scaffold.created)

    def test_no_scaffold_leaves_the_folder_alone(self):
        before = sorted(os.listdir(self.course_dir))
        result = self.init(scaffold=False)
        self.assertEqual(sorted(os.listdir(self.course_dir)), before)
        self.assertEqual(result.scaffold.created, [])
        self.assertTrue(os.path.isfile(result.config_path))

    def test_missing_course_code_is_an_error(self):
        with self.assertRaises(InitCourseError):
            self.init(course_code=None)

    def test_classroom50_org_is_suggested_from_existing_courses(self):
        result = self.init()
        self.assertEqual(read_config_file(result.config_path)["CLASSROOM50_ORG"], "VNU-HUS")

    def test_summary_lists_pending_platforms(self):
        result = self.init()
        recorder = Recorder()
        summarize(result, out=recorder.out)
        self.assertIn("Bước tiếp theo:", recorder.lines)
        self.assertTrue(any("Google Classroom chưa đặt" in line for line in recorder.lines))


class CliFlagTests(unittest.TestCase):
    def test_flags_are_registered_and_dispatch_before_config_load(self):
        source = open(os.path.join(REPO_ROOT, "course_hoanganhduc", "core.py"),
                      encoding="utf-8").read()
        for flag in ("--init-course", "--init-dir", "--init-name", "--inherit-from",
                     "--init-google-id", "--init-canvas-id", "--init-c50-org",
                     "--init-c50-classroom", "--init-sheet-url", "--init-no-scaffold",
                     "--init-no-db", "--init-non-interactive", "--init-force"):
            self.assertIn("'{}'".format(flag), source)
        # The dispatch must sit before the config load, or an empty folder hits the
        # course-code prompt in get_default_config_path instead of being initialized.
        self.assertLess(source.index("if args.init_course:"),
                        source.index("# Load config and set global variables"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
