# -*- coding: utf-8 -*-
"""Import a roster CSV with a fixed column contract into the local database.

This is the twin of ``--add-google-sheet`` for people who already have a CSV and
do not want to go through Google OAuth.  It is deliberately *not* the same code
path as ``--add-file``: that one auto-detects headers and carries quirks for the
``MAT*.xlsx`` files, while this one states which columns it needs and says so
plainly when they are missing.

Nothing here imports pandas.  ``load_database``/``save_database`` are imported
lazily inside the function that needs them because ``course_hoanganhduc.data``
pulls pandas, openpyxl and sklearn at module scope -- the same reason recorded in
``gclass_invite``.  Keeping this module stdlib-only also keeps its tests runnable
without the optional dependency stack.
"""

from __future__ import annotations

import csv
import io
import re
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, NamedTuple, Sequence, Tuple

# Tried in order.  utf-8-sig swallows the BOM that Excel's "CSV UTF-8" always
# writes, and latin-1 decodes any byte at all, so this chain never fails -- a
# deliberate difference from the Classroom50 lane, where a mis-decoded roster
# would be pushed to GitHub and mojibake is worse than a refusal.
ENCODINGS: Tuple[str, ...] = ("utf-8-sig", "cp1252", "latin-1")

# Aliases are written the way a person types them; they are normalized before
# comparison, so accents and case do not matter here.  The targets are the
# database attribute names the two lanes actually read: gclass_invite reads
# Name/Email/Section/Class, c50_roster reads GitHub Username/GitHub ID/Student ID.
#
# Order matters *within* each tuple: it is the alias's rank.  When two columns
# both claim an attribute, the lower rank wins, which is how a Google Form that
# carries both a school address and a personal one lands the school address in
# Email.  See _map_headers.
COLUMN_ALIASES: Dict[str, Tuple[str, ...]] = {
    "Student ID": ("mã sinh viên", "mssv", "mã sv", "student id", "studentid", "id"),
    "Name": ("họ và tên", "họ tên", "tên", "name", "full name"),
    "Email": (
        "emai vnu-hus",
        "email vnu-hus",
        "email vnu hus",
        "email trường",
        "email nhà trường",
        "email",
        "e-mail",
        "email address",
        "địa chỉ email",
        "gmail",
    ),
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
    "Personal Email": ("email address", "địa chỉ email", "email cá nhân"),
    # "Dob", not "Date of Birth": read_students_from_excel_csv maps both
    # spellings to Dob, so the two import lanes leave one name in the
    # database and the existing exporters find it.
    "Dob": ("ngày sinh", "date of birth"),
    "Timestamp": ("timestamp", "dấu thời gian"),
}

# Which attribute gets first pick of the columns claiming it.  The columns the
# two lanes read come first and the rest take what is left, so a CSV whose only
# address column is "Email Address" still fills Email rather than stranding it
# in Personal Email.
ATTRIBUTE_PRIORITY: Tuple[str, ...] = (
    "Student ID",
    "Name",
    "Email",
    "GitHub Username",
    "GitHub ID",
    "Section",
    "Class",
    "Personal Email",
    "Dob",
    "Timestamp",
)

# An alias has to match a whole header; a rule only has to appear inside one.
# That is the difference that catches "Emai VNU-HUS", where the form itself
# dropped a letter: "emai" still prefixes both the correct and the typo'd
# spelling, and "vnu" pins it to the school address rather than a personal one.
HEADER_RULES: Tuple[Tuple[str, Tuple[str, ...]], ...] = (("Email", ("emai", "vnu")),)

REQUIRED_COLUMNS: Tuple[str, ...] = ("Student ID", "Name")
# Without one of these the CSV feeds neither lane: Google Classroom invites by
# email, Classroom50 by GitHub username.
ONE_OF_COLUMNS: Tuple[str, ...] = ("Email", "Personal Email", "GitHub Username")


class RosterCsvError(Exception):
    """The CSV cannot be used as a roster; raised before the database is touched."""


class ParseResult(NamedTuple):
    """Rows keyed by database attribute name, plus what was dropped and why.

    ``skipped`` counts rows missing a Student ID or a Name; those are dropped one
    by one rather than failing the whole run, because a trailing summary line is
    a normal thing to find at the bottom of an exported sheet.
    """

    rows: List[Dict[str, str]]
    skipped: int
    headers_seen: List[str]


def normalize_header(value: str) -> str:
    """Fold a header to its comparison form: NFKC, lower, no accents, one space.

    Same recipe as the ``normalize_header`` closure in ``data.py``; that one is
    local to its function and cannot be imported.
    """
    raw = unicodedata.normalize("NFKC", str(value)).strip().lower()
    raw = unicodedata.normalize("NFD", raw)
    raw = "".join(c for c in raw if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", raw)


def header_variants(value: str) -> List[str]:
    """Fold a header into every form worth matching against an alias.

    A Google Form header is bilingual and marked required, so the raw text is
    ``"Mã Sinh Viên (Student ID)"`` or ``"Lớp (Class) *"`` rather than the bare
    name an alias lists.  Matching only the whole string drops seven of the nine
    columns on the form this was built for, so each header also yields the text
    with the parenthetical removed, the text inside the parentheses, and the
    text without the form's trailing ``*``.  Duplicates and empties are dropped,
    and the full folded header always comes first.
    """
    raw = str(value).strip()
    unmarked = re.sub(r"[*\s]+$", "", raw)
    candidates = [raw, unmarked, re.sub(r"\([^)]*\)", " ", unmarked)]
    candidates.extend(re.findall(r"\(([^)]*)\)", unmarked))

    variants: List[str] = []
    for candidate in candidates:
        folded = normalize_header(candidate)
        if folded and folded not in variants:
            variants.append(folded)
    return variants


def _build_alias_lookup() -> Dict[str, List[Tuple[str, int]]]:
    """Folded alias -> [(attribute, rank)].

    An alias's rank is its position in its own tuple, which is what lets Email
    prefer the school address over the personal one.  A single alias can serve
    two attributes -- "email address" is both a lower-ranked Email and the first
    choice for Personal Email -- so the value is a list, not a single pair.
    """
    lookup: Dict[str, List[Tuple[str, int]]] = {}
    for attribute, aliases in COLUMN_ALIASES.items():
        for rank, alias in enumerate(aliases):
            lookup.setdefault(normalize_header(alias), []).append((attribute, rank))
    return lookup


_ALIAS_LOOKUP: Dict[str, List[Tuple[str, int]]] = _build_alias_lookup()


def read_csv_text(path: Path, verbose: bool = False) -> Tuple[str, str]:
    """Return the file's text and the encoding that decoded it.

    Falling back past utf-8 is reported, since a wrong guess shows up later as
    mangled Vietnamese names rather than as an error.
    """
    data = Path(path).read_bytes()
    for encoding in ENCODINGS:
        try:
            text = data.decode(encoding)
        except UnicodeDecodeError:
            continue
        if encoding != ENCODINGS[0]:
            print(
                f"Warning: {path} is not valid UTF-8; read it as {encoding}. "
                "Re-save the file as UTF-8 if any name looks wrong."
            )
        elif verbose:
            print(f"[ImportCSV] Read {path} as {encoding}.")
        return text, encoding
    raise RosterCsvError(f"Could not decode {path} with any of: {', '.join(ENCODINGS)}")


def _column_claims(header: str) -> Dict[str, int]:
    """Attribute -> best rank this one header can claim it at."""
    claims: Dict[str, int] = {}
    variants = header_variants(header)
    for variant in variants:
        for attribute, rank in _ALIAS_LOOKUP.get(variant, ()):
            if rank < claims.get(attribute, rank + 1):
                claims[attribute] = rank
    for attribute, needles in HEADER_RULES:
        if any(all(n in variant for n in needles) for variant in variants):
            claims[attribute] = 0
    return claims


def _map_headers(header_row: Sequence[str]) -> Dict[int, str]:
    """Column index -> database attribute, for the columns this module knows.

    Each attribute takes the best column that claims it, in the order of
    ``ATTRIBUTE_PRIORITY``, rather than each column taking the first attribute it
    matches.  The difference shows on a form carrying two addresses: first-wins
    let ``Email Address`` take the Email slot and dropped the school address
    entirely, while here Email picks the higher-ranked school column and
    ``Email Address`` falls through to Personal Email.  A form with only one
    address column still fills Email, because Email gets to choose first.
    """
    claims = [_column_claims(raw) for raw in header_row]
    mapping: Dict[int, str] = {}
    for attribute in ATTRIBUTE_PRIORITY:
        best: Tuple[int, int] = (-1, -1)
        for index, column in enumerate(claims):
            if index in mapping or attribute not in column:
                continue
            if best[0] < 0 or column[attribute] < best[1]:
                best = (index, column[attribute])
        if best[0] >= 0:
            mapping[best[0]] = attribute
    return dict(sorted(mapping.items()))


def parse_student_rows(text: str) -> ParseResult:
    """Turn CSV text into rows keyed by database attribute name.

    Columns the module has no alias for are carried under their own header
    rather than dropped, so a question added to the form later still reaches the
    database without anyone editing ``COLUMN_ALIASES`` first.

    Raises ``RosterCsvError`` when the header does not carry the columns the two
    lanes need, naming both what is missing and what was actually read -- a
    header that is merely spelled differently is the usual cause.
    """
    reader = csv.reader(io.StringIO(text))
    header_row: List[str] = []
    for row in reader:
        if any(cell.strip() for cell in row):
            header_row = row
            break
    if not header_row:
        raise RosterCsvError("The CSV is empty; expected a header row.")

    headers_seen = [cell.strip() for cell in header_row]
    mapping = _map_headers(header_row)
    found = set(mapping.values())

    missing = [name for name in REQUIRED_COLUMNS if name not in found]
    if not any(name in found for name in ONE_OF_COLUMNS):
        missing.append(" or ".join(ONE_OF_COLUMNS))
    if missing:
        raise RosterCsvError(
            "The CSV is missing required column(s): "
            + ", ".join(missing)
            + ". Columns read: "
            + (", ".join(headers_seen) or "(none)")
        )

    rows: List[Dict[str, str]] = []
    skipped = 0
    for row in reader:
        if not any(cell.strip() for cell in row):
            continue
        record: Dict[str, str] = {}
        for index, attribute in mapping.items():
            if index < len(row):
                value = row[index].strip()
                if value:
                    record[attribute] = value
        for index, header in enumerate(headers_seen):
            if index in mapping or not header or index >= len(row):
                continue
            value = row[index].strip()
            if value and header not in record:
                record[header] = value
        if not all(record.get(name) for name in REQUIRED_COLUMNS):
            skipped += 1
            continue
        rows.append(record)
    return ParseResult(rows, skipped, headers_seen)


def import_students_from_csv(
    csv_path: str,
    db_path: str,
    verbose: bool = False,
    dry_run: bool = False,
    preview_rows: int = 5,
) -> Dict[str, Any]:
    """Read ``csv_path`` and merge its rows into the database at ``db_path``.

    Merging goes through ``save_database``, whose ``_dedup_students`` already
    collapses repeats by Student ID and name; no second merge rule is invented
    here.  ``dry_run`` prints a preview and writes nothing.
    """
    from .data import load_database, save_database
    from .models import Student

    text, encoding = read_csv_text(Path(csv_path), verbose=verbose)
    parsed = parse_student_rows(text)

    print(
        f"Read {len(parsed.rows)} student row(s) from {csv_path} "
        f"({encoding}; {parsed.skipped} row(s) skipped for a missing "
        f"{' or '.join(REQUIRED_COLUMNS)})."
    )
    for record in parsed.rows[:preview_rows]:
        print("  " + " | ".join(f"{k}: {v}" for k, v in record.items()))
    if len(parsed.rows) > preview_rows:
        print(f"  ... and {len(parsed.rows) - preview_rows} more.")

    report: Dict[str, Any] = {
        "csv": str(csv_path),
        "encoding": encoding,
        "read": len(parsed.rows),
        "skipped": parsed.skipped,
        "columns": sorted(set(_map_headers(parsed.headers_seen).values())),
        "extra_columns": sorted(
            {key for record in parsed.rows for key in record} - set(ATTRIBUTE_PRIORITY)
        ),
        "dry_run": bool(dry_run),
        "saved": False,
    }
    if dry_run:
        print("Dry run: the database was not modified.")
        return report
    if not parsed.rows:
        print("Nothing to import.")
        return report

    existing = load_database(db_path, verbose=verbose) or []
    # Every column that was read is written.  Student takes arbitrary keyword
    # arguments and the database is a pickle of those objects, so the answer to
    # a later question about a form column is in the database rather than back
    # in the sheet the CSV came from.
    new_students = [Student(**record) for record in parsed.rows]
    save_database(
        list(existing) + new_students,
        db_path,
        verbose=verbose,
        audit_source="import-csv",
    )
    report["saved"] = True
    report["existing"] = len(existing)
    return report
