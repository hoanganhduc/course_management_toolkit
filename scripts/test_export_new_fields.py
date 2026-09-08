#!/usr/bin/env python3
"""Tests that "export everything" keeps meaning everything.

Three lanes were added to the toolkit — Classroom50 scores, Classroom50 groups
and mini-project issues — and each writes new attributes onto ``Student``.  The
two export paths do not pick those up on their own: the Excel exporter lists
its scalar columns by hand, and the two detail views each hold a hand-written
tuple of container keys back from the plain key loop.  A field added to a lane
and forgotten here vanishes from the exports without any error.

So the suite is written as rules over what the lanes *declare*, not over a list
of field names copied into the test.  ``declared_fields()`` reads the
``FIELD_*`` constants out of the three modules, and the coverage test demands
that every one of them reach both export paths.  Add a field to a lane, forget
the exporter, and this suite fails naming it.

Unlike most suites in this directory, this one imports the real database layer,
so it needs the course interpreter rather than a bare python3:

    ~/.course_venv/bin/python scripts/test_export_new_fields.py -v
"""

from __future__ import annotations

import ast
import contextlib
import io
import os
import re
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any, Dict, List, Set, Tuple

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

import pandas as pd

from course_hoanganhduc import data as data_module
from course_hoanganhduc import c50_groups, c50_scores, project_issues, settings
from course_hoanganhduc.data import (
    _C50_CONTAINER_FIELDS,
    en_to_vn_field,
    export_all_details_to_txt,
    export_anonymized_roster,
    export_to_excel,
    print_all_student_details,
)
from course_hoanganhduc.models import Student

DATA_SOURCE = Path(REPO_ROOT) / "course_hoanganhduc" / "data.py"

# A real Vietnamese name, not "Test Student": save_database runs refine_database,
# which deletes students whose Name matches its test/sample patterns.
REAL_NAME = "Nguyễn Văn An"
REAL_ID = "24007008"
SLUG_ONE = "ch01-introduction"
SLUG_TWO = "w00-group-collaboration"
COLLECTED_AT = "2026-09-05T04:52:10Z"


# --------------------------------------------------------------------------
# What the lanes declare, read from the lanes themselves
# --------------------------------------------------------------------------
def declared_fields() -> Set[str]:
    """Every database field name the three new lanes name as a constant.

    Reading the constants rather than restating them is the point: a field
    added to a lane joins this set on its own, and the coverage test below then
    demands the exporters carry it.
    """
    found: Set[str] = set()
    for module in (c50_scores, c50_groups, project_issues):
        for name, value in vars(module).items():
            if name.startswith("FIELD_") and isinstance(value, str):
                found.add(value)
    return found


def hand_graded_fields() -> Set[str]:
    """The final-project grade slot and its free-text note.

    These belong to the pre-existing ``Override *`` lane rather than to the
    three new modules, and the audit list is where that lane declares them.
    """
    return {f for f in settings.GRADE_AUDIT_FIELDS if f.startswith("Override Final Project")}


def new_scalar_fields() -> Set[str]:
    return (declared_fields() | hand_graded_fields()) - set(_C50_CONTAINER_FIELDS)


# --------------------------------------------------------------------------
# Reading the two hand-written lists out of the source
# --------------------------------------------------------------------------
def _source_tree() -> ast.Module:
    return ast.parse(DATA_SOURCE.read_text(encoding="utf-8"))


def _function(name: str) -> ast.FunctionDef:
    for node in ast.walk(_source_tree()):
        if isinstance(node, ast.FunctionDef) and node.name == name:
            return node
    raise AssertionError(f"{name} is gone from data.py")


def _assigned_in(func: ast.FunctionDef, target: str) -> ast.expr:
    for node in ast.walk(func):
        if isinstance(node, ast.Assign):
            first = node.targets[0]
            if isinstance(first, ast.Name) and first.id == target:
                return node.value
    raise AssertionError(f"{target} is gone from {func.name}")


def vietnamese_labels() -> Dict[str, str]:
    return ast.literal_eval(_assigned_in(_function("en_to_vn_field"), "field_vn_map"))


def export_field_names() -> List[str]:
    node = _assigned_in(_function("export_to_excel"), "export_fields")
    return [entry.elts[0].value for entry in node.elts]


def excluded_keys(function_name: str) -> Set[str]:
    """The container keys one detail view holds back from its plain key loop.

    Each view spells the tuple inline inside a comprehension condition, so the
    keys are collected from every ``k not in (...)`` test in the function plus
    the shared tuple it now also references by name.
    """
    keys: Set[str] = set()
    for node in ast.walk(_function(function_name)):
        if not isinstance(node, ast.Compare):
            continue
        if not node.ops or not isinstance(node.ops[0], ast.NotIn):
            continue
        right = node.comparators[0]
        if isinstance(right, ast.Tuple):
            keys.update(el.value for el in right.elts if isinstance(el, ast.Constant))
        elif isinstance(right, ast.Name) and right.id == "_C50_CONTAINER_FIELDS":
            keys.update(_C50_CONTAINER_FIELDS)
    return keys


# --------------------------------------------------------------------------
# Fixtures
# --------------------------------------------------------------------------
def loaded_student() -> Student:
    """One student carrying every field the new lanes write.

    The scalars are filled from ``new_scalar_fields()`` so that a field added
    later is exercised without editing this fixture; the containers get shapes
    matching what the lanes actually store.
    """
    student = Student(
        **{
            "Name": REAL_NAME,
            "Student ID": REAL_ID,
            "Email": "an.nguyen@example.edu.vn",
            "Class": "K61A6",
            "Section": "01",
            "GitHub Username": "an-nguyen",
        }
    )
    for index, field in enumerate(sorted(new_scalar_fields()), 1):
        setattr(student, field, f"value-{index}")
    setattr(student, c50_scores.FIELD_COLLECTED_AT, COLLECTED_AT)
    setattr(student, "Override Final Project", 78)
    setattr(student, "Override Final Project Reason", "slide: 3; Q&A 5")
    setattr(
        student,
        c50_scores.FIELD_GRADES,
        {
            SLUG_ONE: {
                "grade": 85,
                "max_points": 100,
                "commit": "3af80db0206215cfd43314b5feba8060c632e417",
                "release": "v1",
                "datetime": "2026-09-05T10:00:00Z",
                "late": False,
                "group": False,
                "owner": "an-nguyen",
            },
            SLUG_TWO: {
                "grade": 40,
                "max_points": 50,
                "commit": "cbb330fa1d2e3f405162738495a6b7c8d9e0f102",
                "release": "v1",
                "datetime": "2026-09-11T10:00:00Z",
                "late": True,
                "group": True,
                "owner": "binh-tran",
            },
        },
    )
    setattr(
        student,
        c50_scores.FIELD_SUBMISSIONS,
        {SLUG_ONE: "submitted", SLUG_TWO: "not collected yet"},
    )
    setattr(
        student,
        c50_scores.FIELD_DETAILS,
        {
            SLUG_ONE: [
                {
                    "datetime": "2026-09-05T10:00:00Z",
                    "score": 85,
                    "max-score": 100,
                    "commit": "3af80db0206215cfd43314b5feba8060c632e417",
                    "tests": [
                        {"test-name": "t1", "passed": True, "score": 1, "max-score": 1},
                        {"test-name": "t2", "passed": False, "score": 0, "max-score": 1},
                    ],
                }
            ]
        },
    )
    setattr(student, c50_scores.FIELD_OVERRIDES, {SLUG_ONE: True})
    setattr(
        student,
        c50_groups.FIELD_GROUP,
        {
            SLUG_TWO: {
                "repo": "vnu-hus-mat1206e-winter-2026-w00-group-collaboration-binh-tran",
                "founder": "binh-tran",
                "members": ["binh-tran", "an-nguyen"],
                "source": "c50-collaborators",
                "group_name": "Nhóm 1",
            }
        },
    )
    setattr(
        student,
        c50_groups.FIELD_GROUP_CONFLICTS,
        {SLUG_TWO: ["group_over_size: 6 student members"]},
    )
    return student


def bare_student() -> Student:
    """A student none of the new lanes has touched — the common case today."""
    return Student(
        **{
            "Name": "Trần Thị Bình",
            "Student ID": "24001100",
            "Email": "binh.tran@example.edu.vn",
            "Class": "K61A6",
            "Section": "01",
        }
    )


@contextlib.contextmanager
def workdir():
    """Run inside a temp directory: the exporters append to ./run_report.txt."""
    previous = os.getcwd()
    with tempfile.TemporaryDirectory() as tmp:
        os.chdir(tmp)
        try:
            yield Path(tmp)
        finally:
            os.chdir(previous)


def screen_text(students: List[Student]) -> str:
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer):
        print_all_student_details(list(students))
    return buffer.getvalue()


def txt_text(students: List[Student], tmp: Path) -> str:
    target = tmp / "details.txt"
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer):
        export_all_details_to_txt(list(students), file_path=str(target))
    return target.read_text(encoding="utf-8")


def excel_frame(students: List[Student], tmp: Path) -> pd.DataFrame:
    """Export to Excel with every offered column selected.

    ``_select_export_columns`` always prompts, by design; the patch here stands
    in for a user who ticks everything, which is what the picker pre-selects.
    """
    target = tmp / "export.xlsx"
    original = data_module._select_export_columns
    data_module._select_export_columns = lambda available, pre_selected=None: list(available)
    try:
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            export_to_excel(list(students), file_path=str(target))
    finally:
        data_module._select_export_columns = original
    return pd.read_excel(target)


def anonymized_frame(students: List[Student], tmp: Path) -> pd.DataFrame:
    target = tmp / "anon.csv"
    original = data_module._select_export_columns
    data_module._select_export_columns = lambda available, pre_selected=None: list(available)
    try:
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            export_anonymized_roster(list(students), file_path=str(target))
    finally:
        data_module._select_export_columns = original
    return pd.read_csv(target, dtype=str).fillna("")


# --------------------------------------------------------------------------


class TestCoverage(unittest.TestCase):
    """Every field a lane declares must reach both export paths."""

    def test_every_declared_field_is_a_container_or_an_excel_column(self) -> None:
        columns = set(export_field_names())
        for field in sorted(new_scalar_fields()):
            self.assertIn(
                field,
                columns,
                f"{field} is written by a lane but has no column in export_to_excel",
            )

    def test_the_container_tuple_names_only_fields_a_lane_writes(self) -> None:
        """No stale entry: a held-back key that nothing writes hides nothing."""
        for field in _C50_CONTAINER_FIELDS:
            self.assertIn(field, declared_fields(), f"{field} is held back but unwritten")

    def test_both_hand_grade_fields_are_declared_and_exported(self) -> None:
        self.assertEqual(
            hand_graded_fields(),
            {"Override Final Project", "Override Final Project Reason"},
        )
        self.assertLessEqual(hand_graded_fields(), set(export_field_names()))

    def test_every_new_scalar_reaches_the_excel_sheet(self) -> None:
        with workdir() as tmp:
            frame = excel_frame([loaded_student()], tmp)
        for field in sorted(new_scalar_fields()):
            self.assertIn(en_to_vn_field(field), frame.columns, field)


class TestContainersAreRendered(unittest.TestCase):
    """The six containers are summarised, never dumped as raw dicts."""

    def test_both_detail_views_hold_back_every_container(self) -> None:
        for view in ("print_all_student_details", "export_all_details_to_txt"):
            excluded = excluded_keys(view)
            for field in _C50_CONTAINER_FIELDS:
                self.assertIn(field, excluded, f"{field} escapes the key loop in {view}")

    def test_neither_view_prints_a_raw_container_dict(self) -> None:
        student = loaded_student()
        with workdir() as tmp:
            outputs = {"screen": screen_text([student]), "txt": txt_text([student], tmp)}
        for name, text in outputs.items():
            for field in _C50_CONTAINER_FIELDS:
                value = getattr(student, field, None)
                if not value:
                    continue
                self.assertNotIn(repr(value), text, f"{field} dumped raw in {name}")
                self.assertNotIn(f"{field}: {{", text, f"{field} dumped raw in {name}")

    def test_both_views_summarise_the_same_facts(self) -> None:
        student = loaded_student()
        with workdir() as tmp:
            screen = screen_text([student])
            txt = txt_text([student], tmp)
        for text in (screen, txt):
            self.assertIn("Classroom50 grades", text)
            self.assertIn(f"{SLUG_ONE}: 85/100", text)
            self.assertIn("giảng viên đã sửa", text)   # the override marker
            self.assertIn("nộp muộn", text)            # the late marker
            self.assertIn("nhóm", text)                # the group marker
            self.assertIn("Total score (Classroom50) (Tổng điểm Classroom50): 125.0/150.0", text)
            self.assertIn("Classroom50 submissions", text)
            self.assertIn("not collected yet", text)
            self.assertIn("Classroom50 submission details", text)
            self.assertIn("tests: 1/2 passed", text)
            self.assertIn("Classroom50 group", text)
            self.assertIn("Nhóm 1", text)
            self.assertIn("Classroom50 group conflicts", text)
            self.assertIn("group_over_size", text)

    def test_the_classroom50_total_is_not_folded_into_the_classroom_total(self) -> None:
        """Two lanes, two totals.

        The Google Classroom block already prints "Total score (Classroom)".
        Adding Classroom50 marks to that sum would silently change a number the
        lecturer already reads, so the new line is separately labelled.
        """
        student = loaded_student()
        setattr(student, "Grades", {"Bài 1": {"grade": 9, "max_points": 10}})
        with workdir() as tmp:
            for text in (screen_text([student]), txt_text([student], tmp)):
                self.assertIn("Total score (Classroom) (Tổng điểm Classroom): 9.0/10.0", text)
                self.assertIn("Total score (Classroom50) (Tổng điểm Classroom50): 125.0/150.0", text)


class TestScalarsStayVisible(unittest.TestCase):
    """A scalar held back with the containers would disappear from both views."""

    def test_the_collection_timestamp_is_in_neither_exclusion_tuple(self) -> None:
        field = c50_scores.FIELD_COLLECTED_AT
        for view in ("print_all_student_details", "export_all_details_to_txt"):
            self.assertNotIn(field, excluded_keys(view), f"{field} held back in {view}")

    def test_the_collection_timestamp_shows_in_both_views(self) -> None:
        with workdir() as tmp:
            student = loaded_student()
            for text in (screen_text([student]), txt_text([student], tmp)):
                self.assertIn(COLLECTED_AT, text)

    def test_the_hand_grade_and_its_note_show_in_the_txt_export(self) -> None:
        with workdir() as tmp:
            text = txt_text([loaded_student()], tmp)
        self.assertIn("78", text)
        # The note is free text a lecturer typed; it is stored and shown as
        # typed, never parsed into fields.
        self.assertIn("slide: 3; Q&A 5", text)

    def test_the_hand_grade_survives_as_a_number_and_as_a_string(self) -> None:
        numeric = loaded_student()
        setattr(numeric, "Override Final Project", 78)
        textual = bare_student()
        setattr(textual, "Override Final Project", "78")
        column = en_to_vn_field("Override Final Project")
        with workdir() as tmp:
            frame = excel_frame([numeric, textual], tmp)
        self.assertEqual({str(v) for v in frame[column].tolist()}, {"78"})


class TestVietnameseLabels(unittest.TestCase):
    def test_every_new_field_has_a_label(self) -> None:
        labels = vietnamese_labels()
        for field in sorted(declared_fields() | hand_graded_fields()):
            self.assertIn(field, labels, f"{field} has no Vietnamese label")

    def test_no_notice_line_names_a_new_field(self) -> None:
        """A missing label prints a notice once per key per student.

        The txt exporter labels every key of every student, so one unmapped
        field turns into one junk line per student in the file the lecturer
        reads.
        """
        new_fields = declared_fields() | hand_graded_fields()
        with workdir() as tmp:
            buffer = io.StringIO()
            with contextlib.redirect_stdout(buffer):
                export_all_details_to_txt(
                    [loaded_student()], file_path=str(tmp / "details.txt"), verbose=True
                )
            printed = buffer.getvalue()
        notices = [line for line in printed.splitlines() if "No Vietnamese mapping" in line]
        offenders = [line for line in notices if any(f"'{f}'" in line for f in new_fields)]
        self.assertEqual(offenders, [], "\n".join(offenders))

    def test_new_labels_are_pairwise_distinct_and_unlike_the_old_ones(self) -> None:
        """Two fields sharing a label silently lose a column.

        ``export_to_excel`` keys its header map by the Vietnamese label, so a
        collision drops one of the two fields with no error.  The codebase
        already has such collisions — Section, Canvas Section, Registration ID
        and Registered Class ID all read "Lớp học phần" — which is what makes
        this worth pinning for the new fields.
        """
        labels = vietnamese_labels()
        new_fields = sorted(declared_fields() | hand_graded_fields())
        seen: Dict[str, str] = {}
        for field in new_fields:
            label = labels[field]
            self.assertNotIn(
                label, seen, f"{field} and {seen.get(label)} share the label {label!r}"
            )
            seen[label] = field
        for field, label in labels.items():
            if field in seen.values() or field in new_fields:
                continue
            self.assertNotIn(
                label,
                seen,
                f"{field} already uses the label {label!r} claimed by {seen.get(label)}",
            )


class TestDynamicColumns(unittest.TestCase):
    def test_each_classroom50_slug_becomes_its_own_column(self) -> None:
        with workdir() as tmp:
            frame = excel_frame([loaded_student()], tmp)
        self.assertIn(f"C50: {SLUG_ONE}", frame.columns)
        self.assertIn(f"C50: {SLUG_TWO}", frame.columns)
        self.assertEqual(frame[f"C50: {SLUG_ONE}"].tolist(), ["submitted"])
        self.assertEqual(frame[f"C50: {SLUG_TWO}"].tolist(), ["not collected yet"])

    def test_the_slug_columns_are_separate_from_the_other_two_lanes(self) -> None:
        """Same slug, three lanes, three columns — no lane overwrites another."""
        student = loaded_student()
        setattr(student, "Submissions", {SLUG_ONE: "turned in"})
        setattr(student, "Canvas Submissions", {SLUG_ONE: "graded"})
        with workdir() as tmp:
            frame = excel_frame([student], tmp)
        self.assertEqual(frame[f"GC Submission: {SLUG_ONE}"].tolist(), ["turned in"])
        self.assertEqual(frame[f"Canvas Submission: {SLUG_ONE}"].tolist(), ["graded"])
        self.assertEqual(frame[f"C50: {SLUG_ONE}"].tolist(), ["submitted"])

    def test_a_student_without_classroom50_data_gets_an_empty_cell(self) -> None:
        with workdir() as tmp:
            frame = excel_frame([loaded_student(), bare_student()], tmp)
        column = frame[f"C50: {SLUG_ONE}"].fillna("").tolist()
        self.assertIn("submitted", column)
        self.assertIn("", column)

    def test_a_roster_with_no_classroom50_data_exports_no_slug_columns(self) -> None:
        with workdir() as tmp:
            frame = excel_frame([bare_student()], tmp)
        self.assertEqual([c for c in frame.columns if str(c).startswith("C50: ")], [])


class TestUntouchedStudents(unittest.TestCase):
    """The lanes are prospective; today almost every student has none of this."""

    def test_all_three_exporters_run_clean_on_a_bare_roster(self) -> None:
        with workdir() as tmp:
            screen = screen_text([bare_student()])
            txt = txt_text([bare_student()], tmp)
            frame = excel_frame([bare_student()], tmp)
        for text in (screen, txt):
            self.assertIn("Trần Thị Bình", text)
            self.assertNotIn("Classroom50 grades", text)
            self.assertNotIn("Classroom50 group", text)
        for field in sorted(new_scalar_fields()):
            column = en_to_vn_field(field)
            self.assertIn(column, frame.columns)
            self.assertEqual(frame[column].fillna("").tolist(), [""])

    def test_a_container_holding_the_wrong_shape_does_not_crash_the_views(self) -> None:
        """The lanes write dicts; a hand edit or an older database may not."""
        student = bare_student()
        for field in _C50_CONTAINER_FIELDS:
            setattr(student, field, "not a dict")
        with workdir() as tmp:
            screen_text([student])
            txt_text([student], tmp)


class TestAnonymizedRosterStaysAnonymous(unittest.TestCase):
    """The anonymised roster builds its columns from every attribute present."""

    def test_no_lane_declares_a_second_copy_of_the_name_or_the_student_id(self) -> None:
        """``team.json`` is read for cross-checking, never for storage.

        The anonymiser masks exactly two columns, ``Name`` and ``Student ID``,
        and takes every other attribute through untouched.  A lane that stored
        the full name or student id from ``team.json`` under any other field
        name would therefore publish it in the file called anonymised.
        """
        forbidden = re.compile(r"full[_ ]?name|student[_ ]?id", re.IGNORECASE)
        for field in sorted(declared_fields() | hand_graded_fields()):
            self.assertIsNone(
                forbidden.search(field),
                f"{field} reads as a second copy of masked identity data",
            )

    def test_a_fully_loaded_student_leaks_neither_name_nor_id(self) -> None:
        with workdir() as tmp:
            frame = anonymized_frame([loaded_student()], tmp)
        rendered = frame.to_csv(index=False)
        self.assertNotIn(REAL_NAME, rendered)
        self.assertNotIn(REAL_ID, rendered)

    def test_every_new_field_still_reaches_the_anonymised_file(self) -> None:
        """Masking identity is not a reason to drop the rest of the record."""
        with workdir() as tmp:
            frame = anonymized_frame([loaded_student()], tmp)
        for field in sorted(new_scalar_fields()):
            self.assertIn(field, frame.columns, field)


if __name__ == "__main__":
    unittest.main(verbosity=2)
