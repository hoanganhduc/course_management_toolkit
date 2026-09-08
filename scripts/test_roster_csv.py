#!/usr/bin/env python3
"""Offline tests for the fixed-contract roster CSV importer.

The module under test is stdlib-only and imports the database layer lazily, so
these tests run under a bare interpreter with no pandas installed.
"""

from __future__ import annotations

import io
import os
import sys
import tempfile
import types
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from typing import Any, Dict, List

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc.roster_csv import (  # noqa: E402
    COLUMN_ALIASES,
    RosterCsvError,
    import_students_from_csv,
    _map_headers,
    parse_student_rows,
    read_csv_text,
)

VIETNAMESE_HEADER = "Mã sinh viên,Họ và tên,Email,GitHub username\n"
ENGLISH_HEADER = "Student ID,Full Name,Email,GitHub\n"
ROW = "22001,Nguyễn Văn A,a@hus.edu.vn,nva\n"


class _CsvCase(unittest.TestCase):
    def setUp(self) -> None:
        self.td = tempfile.TemporaryDirectory()
        self.addCleanup(self.td.cleanup)

    def write(self, text: str, encoding: str = "utf-8", name: str = "roster.csv") -> Path:
        path = Path(self.td.name) / name
        path.write_bytes(text.encode(encoding))
        return path

    def read(self, path: Path):
        buffer = io.StringIO()
        with redirect_stdout(buffer):
            text, encoding = read_csv_text(path)
        return text, encoding, buffer.getvalue()


class TestHeaders(_CsvCase):
    def test_vietnamese_headers_map_to_database_attributes(self):
        parsed = parse_student_rows(VIETNAMESE_HEADER + ROW)
        self.assertEqual(
            parsed.rows,
            [
                {
                    "Student ID": "22001",
                    "Name": "Nguyễn Văn A",
                    "Email": "a@hus.edu.vn",
                    "GitHub Username": "nva",
                }
            ],
        )
        self.assertEqual(parsed.skipped, 0)

    def test_english_aliases_give_the_same_rows(self):
        self.assertEqual(
            parse_student_rows(ENGLISH_HEADER + ROW).rows,
            parse_student_rows(VIETNAMESE_HEADER + ROW).rows,
        )

    def test_missing_name_column_names_what_it_read(self):
        with self.assertRaises(RosterCsvError) as caught:
            parse_student_rows("Mã sinh viên,Email\n22001,a@hus.edu.vn\n")
        message = str(caught.exception)
        self.assertIn("Name", message)
        self.assertIn("Mã sinh viên", message)

    def test_neither_email_nor_github_is_refused(self):
        with self.assertRaises(RosterCsvError) as caught:
            parse_student_rows("Mã sinh viên,Họ và tên,Lớp\n22001,Nguyễn Văn A,K68\n")
        self.assertIn("Email or GitHub Username", str(caught.exception))

    def test_row_without_a_student_id_is_skipped_not_fatal(self):
        parsed = parse_student_rows(
            VIETNAMESE_HEADER + ROW + ",Trần Thị B,b@hus.edu.vn,ttb\n"
        )
        self.assertEqual(len(parsed.rows), 1)
        self.assertEqual(parsed.skipped, 1)


class TestEncoding(_CsvCase):
    def test_utf8_bom_is_swallowed(self):
        path = self.write(VIETNAMESE_HEADER + ROW, encoding="utf-8-sig")
        text, encoding, printed = self.read(path)
        self.assertEqual(encoding, "utf-8-sig")
        self.assertEqual(printed, "")
        self.assertEqual(parse_student_rows(text).rows[0]["Student ID"], "22001")

    def test_windows_1252_bytes_are_read_with_a_warning(self):
        path = self.write("Student ID,Full Name,Email\n1,René,r@x.edu\n", encoding="cp1252")
        text, encoding, printed = self.read(path)
        self.assertEqual(encoding, "cp1252")
        self.assertIn("cp1252", printed)
        self.assertEqual(parse_student_rows(text).rows[0]["Name"], "René")


class TestImport(_CsvCase):
    """The database layer is stubbed; nothing here reads or writes a real DB."""

    def setUp(self) -> None:
        super().setUp()
        self.saved: List[Dict[str, Any]] = []
        self.existing = [types.SimpleNamespace(Name="Old Student")]
        self._real = sys.modules.get("course_hoanganhduc.data")
        module = types.ModuleType("course_hoanganhduc.data")
        module.load_database = lambda db_path, verbose=False: list(self.existing)
        module.save_database = lambda rows, db_path, verbose=False, audit_source=None: (
            self.saved.append(
                {"rows": rows, "db_path": db_path, "audit_source": audit_source}
            )
        )
        sys.modules["course_hoanganhduc.data"] = module

    def tearDown(self) -> None:
        if self._real is None:
            sys.modules.pop("course_hoanganhduc.data", None)
        else:
            sys.modules["course_hoanganhduc.data"] = self._real

    def run_import(self, path: Path, **kwargs: Any):
        buffer = io.StringIO()
        with redirect_stdout(buffer):
            report = import_students_from_csv(str(path), db_path="students.db", **kwargs)
        return report, buffer.getvalue()

    def test_merge_goes_through_the_existing_save_path(self):
        path = self.write(VIETNAMESE_HEADER + ROW)
        report, _ = self.run_import(path)
        self.assertTrue(report["saved"])
        self.assertEqual(len(self.saved), 1)
        write = self.saved[0]
        self.assertEqual(write["db_path"], "students.db")
        self.assertEqual(write["audit_source"], "import-csv")
        self.assertEqual(len(write["rows"]), len(self.existing) + 1)
        self.assertEqual(getattr(write["rows"][-1], "Name"), "Nguyễn Văn A")

    def test_dry_run_writes_nothing_and_previews(self):
        rows = "".join(f"2200{i},Sinh Viên {i},s{i}@hus.edu.vn,sv{i}\n" for i in range(4))
        path = self.write(VIETNAMESE_HEADER + rows)
        report, printed = self.run_import(path, dry_run=True, preview_rows=2)
        self.assertFalse(report["saved"])
        self.assertEqual(self.saved, [])
        self.assertEqual(printed.count("Student ID:"), 2)
        self.assertIn("and 2 more", printed)

    def test_a_bad_header_never_reaches_the_database(self):
        path = self.write("Mã sinh viên,Email\n22001,a@hus.edu.vn\n")
        with self.assertRaises(RosterCsvError):
            self.run_import(path)
        self.assertEqual(self.saved, [])

    def test_every_mapped_column_reaches_the_database(self):
        """Answering a question about the form must not need the sheet again."""
        path = self.write(
            "Timestamp,Email Address,Họ và tên,Mã sinh viên,Ngày sinh\n"
            "9/5/2026 8:00:00,a@gmail.com,Nguyễn Văn A,22001,1/2/2004\n"
        )
        report, _ = self.run_import(path)
        self.assertTrue(report["saved"])
        written = self.saved[0]["rows"][-1]
        # One address column, so it fills Email and Personal Email stays empty.
        self.assertEqual(getattr(written, "Email"), "a@gmail.com")
        self.assertFalse(hasattr(written, "Personal Email"))
        self.assertEqual(getattr(written, "Dob"), "1/2/2004")
        self.assertEqual(getattr(written, "Timestamp"), "9/5/2026 8:00:00")

    def test_both_address_columns_are_kept_apart(self):
        """The form carries a typed school address and one Google collected."""
        path = self.write(
            "Email Address,Họ và tên,Emai VNU-HUS,Mã sinh viên\n"
            "a@gmail.com,Nguyễn Văn A,22001@hus.edu.vn,22001\n"
        )
        self.run_import(path)
        written = self.saved[0]["rows"][-1]
        self.assertEqual(getattr(written, "Email"), "22001@hus.edu.vn")
        self.assertEqual(getattr(written, "Personal Email"), "a@gmail.com")

    def test_a_column_with_no_alias_is_stored_under_its_header(self):
        """A question added to the form later needs no alias to be kept."""
        path = self.write(
            "Mã sinh viên,Họ và tên,Email,Số điện thoại\n"
            "22001,Nguyễn Văn A,a@hus.edu.vn,0912345678\n"
        )
        report, _ = self.run_import(path)
        written = self.saved[0]["rows"][-1]
        self.assertEqual(getattr(written, "Số điện thoại"), "0912345678")
        self.assertEqual(report["extra_columns"], ["Số điện thoại"])

    def test_the_real_form_header_stores_all_nine_columns(self):
        path = self.write(
            "Timestamp,Email Address,Họ và Tên (Full Name),Emai VNU-HUS,"
            "GitHub Username,Mã Sinh Viên (Student ID),Ngày sinh (Date of Birth),"
            "Lớp (Class) *,Lớp học phần (Course Section)\n"
            "9/5/2026 7:45:03,24001208@hus.edu.vn,Trần Quang Sơn,"
            "24001208@hus.edu.vn,SonHocCode100,24001208,3/4/2000,"
            "K68A4,MAT3508 4\n"
        )
        self.run_import(path)
        written = self.saved[0]["rows"][-1]
        self.assertEqual(
            {k: v for k, v in written.__dict__.items()},
            {
                "Timestamp": "9/5/2026 7:45:03",
                "Personal Email": "24001208@hus.edu.vn",
                "Name": "Trần Quang Sơn",
                "Email": "24001208@hus.edu.vn",
                "GitHub Username": "SonHocCode100",
                "Student ID": "24001208",
                "Dob": "3/4/2000",
                "Class": "K68A4",
                "Section": "MAT3508 4",
            },
        )


class TestAliasRegression(unittest.TestCase):
    """Every alias the module carried before the onboarding work still lands
    on the same attribute; the new aliases are additions, not replacements."""

    ALIASES_BEFORE_ONBOARDING = {
        "Student ID": ("mã sinh viên", "mssv", "mã sv", "student id", "studentid", "id"),
        "Name": ("họ và tên", "họ tên", "tên", "name", "full name"),
        "Email": ("email", "e-mail", "email address", "địa chỉ email", "gmail"),
        "GitHub Username": (
            "github username",
            "github",
            "github account",
            "github handle",
            "tài khoản github",
        ),
        "GitHub ID": ("github id",),
        "Section": ("lớp học phần", "course section", "section", "nhóm", "lớp"),
        "Class": ("class", "lớp khoá học", "lớp khóa học"),
    }

    def test_every_pre_existing_alias_maps_where_it_always_did(self):
        for attribute, aliases in self.ALIASES_BEFORE_ONBOARDING.items():
            for alias in aliases:
                with self.subTest(attribute=attribute, alias=alias):
                    self.assertEqual(_map_headers([alias]), {0: attribute})

    def test_the_alias_targets_are_still_spelled_the_same(self):
        for attribute, aliases in self.ALIASES_BEFORE_ONBOARDING.items():
            with self.subTest(attribute=attribute):
                self.assertEqual(
                    COLUMN_ALIASES[attribute][-len(aliases) :], tuple(aliases)
                )


if __name__ == "__main__":
    unittest.main(verbosity=2)
