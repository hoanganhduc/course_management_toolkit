# -*- coding: utf-8 -*-
"""Classroom50 CSV dialect + identity join + export (ADR D5)."""

from __future__ import annotations

import csv
import io
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from .models import Student

CANONICAL_COLUMNS = [
    "username",
    "first_name",
    "last_name",
    "email",
    "section",
    "github_id",
]


def _norm_id(value: Any) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip()
    if s in ("", "0", "None"):
        return None
    try:
        if int(float(s)) == 0:
            return None
    except Exception:
        pass
    # prefer integer string when numeric
    try:
        return str(int(float(s)))
    except Exception:
        return s


def _norm_username(value: Any) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    if s.startswith("@"):
        s = s[1:]
    return s.lower() or None


def _norm_email(value: Any) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip().lower()
    return s or None


def parse_roster_payload(data: Any) -> List[Dict[str, Any]]:
    """Normalize gh teacher roster list --json into list of dict rows."""
    if data is None:
        return []
    if isinstance(data, list):
        rows = data
    elif isinstance(data, dict):
        for key in ("students", "roster", "entries", "data"):
            if isinstance(data.get(key), list):
                rows = data[key]
                break
        else:
            rows = [data]
    else:
        return []
    out = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        # support both snake and common JSON keys
        username = row.get("username") or row.get("login") or row.get("github_username")
        out.append(
            {
                "username": username,
                "first_name": row.get("first_name") or row.get("firstName") or "",
                "last_name": row.get("last_name") or row.get("lastName") or "",
                "email": row.get("email") or "",
                "section": row.get("section") or "",
                "github_id": row.get("github_id") if "github_id" in row else row.get("githubId"),
            }
        )
    return out


def local_github_id(student: Any) -> Optional[str]:
    for key in ("GitHub ID", "github_id", "Github ID"):
        v = _norm_id(getattr(student, key, None))
        if v:
            return v
    return None


def local_github_username(student: Any) -> Optional[str]:
    for key in ("GitHub Username", "GitHub Account", "GitHub Handle", "GitHub"):
        v = _norm_username(getattr(student, key, None))
        if v:
            return v
    # do NOT use GitHub ID as username
    return None


def local_email(student: Any) -> Optional[str]:
    return _norm_email(getattr(student, "Email", None) or getattr(student, "email", None))


def join_roster(
    students: Sequence[Any],
    remote_rows: Sequence[Dict[str, Any]],
) -> Dict[str, Any]:
    """
    Match remote C50 rows to local students per ADR D5.
    Returns report dict + pairs for merge application.
    """
    report = {
        "schema": "classroom50-toolkit-report.v1",
        "op": "sync",
        "matched": [],
        "remote_only": [],
        "local_only": [],
        "multi_match": [],
        "conflicts": [],
        "name_split": "none",
        "exit_hint": 0,
        "pairs": [],  # (remote_row, local_student_index)
    }

    # index locals
    by_id: Dict[str, List[int]] = {}
    by_user: Dict[str, List[int]] = {}
    by_email: Dict[str, List[int]] = {}
    for i, s in enumerate(students):
        gid = local_github_id(s)
        if gid:
            by_id.setdefault(gid, []).append(i)
        gu = local_github_username(s)
        if gu:
            by_user.setdefault(gu, []).append(i)
        em = local_email(s)
        if em:
            by_email.setdefault(em, []).append(i)

    matched_local = set()

    for row in remote_rows:
        rid = _norm_id(row.get("github_id"))
        ruser = _norm_username(row.get("username"))
        remail = _norm_email(row.get("email"))

        candidates: Optional[List[int]] = None
        via = None
        if rid and rid in by_id:
            candidates = list(by_id[rid])
            via = "github_id"
        elif ruser and ruser in by_user:
            candidates = list(by_user[ruser])
            via = "username"
        elif remail and remail in by_email:
            candidates = list(by_email[remail])
            via = "email"

        # cross-rule multi-match: if later rule points elsewhere, already handled
        # by exclusive first-hit order; if same rule multi, multi_match
        if not candidates:
            report["remote_only"].append(
                {"username": row.get("username"), "github_id": row.get("github_id")}
            )
            continue
        if len(candidates) != 1:
            report["multi_match"].append(
                {
                    "username": row.get("username"),
                    "via": via,
                    "local_indexes": candidates,
                }
            )
            report["exit_hint"] = 1
            continue

        # also detect cross-rule conflict: id says A, email says B
        if rid and remail:
            id_hits = by_id.get(rid, [])
            em_hits = by_email.get(remail, [])
            if id_hits and em_hits and set(id_hits) != set(em_hits):
                report["multi_match"].append(
                    {
                        "username": row.get("username"),
                        "via": "cross_rule",
                        "id_hits": id_hits,
                        "email_hits": em_hits,
                    }
                )
                report["exit_hint"] = 1
                continue

        idx = candidates[0]
        matched_local.add(idx)
        report["matched"].append(
            {"username": row.get("username"), "local_index": idx, "via": via}
        )
        report["pairs"].append((row, idx))

    for i, s in enumerate(students):
        if i not in matched_local:
            report["local_only"].append(
                {
                    "index": i,
                    "name": getattr(s, "Name", None),
                    "email": getattr(s, "Email", None),
                }
            )

    return report


def apply_fill_only(student: Any, remote: Dict[str, Any], report: Dict[str, Any]) -> None:
    """Fill missing local GitHub Username/ID/empty email; never clobber protected fields."""
    ruser = remote.get("username")
    rid = _norm_id(remote.get("github_id"))
    remail = remote.get("email")

    if ruser and not local_github_username(student):
        setattr(student, "GitHub Username", str(ruser).lstrip("@"))
    if rid and not local_github_id(student):
        setattr(student, "GitHub ID", rid)

    local_em = getattr(student, "Email", None)
    if remail and (local_em is None or str(local_em).strip() == ""):
        setattr(student, "Email", remail)
    elif (
        remail
        and local_em
        and _norm_email(local_em) != _norm_email(remail)
    ):
        report["conflicts"].append(
            {
                "field": "Email",
                "username": ruser,
                "local": local_em,
                "remote": remail,
            }
        )

    # never clobber Student ID / Canvas / Google IDs — we simply never write them


def split_name(full: str) -> Tuple[str, str, str]:
    """Return (first_name, last_name, name_split_mode)."""
    full = (full or "").strip()
    if not full:
        return "", "", "none"
    parts = full.split()
    if len(parts) == 1:
        return parts[0], "", "heuristic"
    return " ".join(parts[:-1]), parts[-1], "heuristic"


def export_roster_csv(students: Sequence[Any]) -> Tuple[str, str]:
    """
    Export local students to C50 6-col CSV text.
    Returns (csv_text, name_split_mode).
    """
    buf = io.StringIO()
    writer = csv.writer(buf, lineterminator="\n")
    writer.writerow(CANONICAL_COLUMNS)
    mode = "none"
    for s in students:
        username = local_github_username(s) or ""
        if not username:
            # try raw GitHub Username with original case
            raw = getattr(s, "GitHub Username", None) or getattr(s, "GitHub", None)
            username = str(raw).lstrip("@").strip() if raw else ""
        if not username:
            continue  # reject empty username rows
        gid = local_github_id(s) or ""
        email = getattr(s, "Email", None) or ""
        section = getattr(s, "Section", None) or ""
        first = getattr(s, "First Name", None) or getattr(s, "first_name", None)
        last = getattr(s, "Last Name", None) or getattr(s, "last_name", None)
        if not first and not last:
            first, last, mode = split_name(getattr(s, "Name", None) or "")
            if mode == "heuristic":
                pass
        else:
            first = first or ""
            last = last or ""
            mode = "explicit" if mode == "none" else mode
        writer.writerow(
            [
                username if not username.startswith("@") else username[1:],
                first or "",
                last or "",
                email or "",
                section or "",
                gid or "",
            ]
        )
    return buf.getvalue(), mode


def parse_csv_text(text: str) -> List[Dict[str, str]]:
    """Parse 5- or 6-col C50 CSV (BOM stripped)."""
    if text.startswith("\ufeff"):
        text = text[1:]
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)
    if not rows:
        return []
    header = [h.strip() for h in rows[0]]
    if header[:5] != CANONICAL_COLUMNS[:5]:
        raise ValueError(f"unexpected header: {header}")
    if len(header) >= 6 and header[5] != "github_id":
        raise ValueError(f"unexpected github_id column: {header}")
    out = []
    for row in rows[1:]:
        if not row or not any(row):
            continue
        while len(row) < 6:
            row.append("")
        username = (row[0] or "").strip()
        if not username:
            raise ValueError("empty username row rejected")
        out.append(
            {
                "username": username,
                "first_name": row[1],
                "last_name": row[2],
                "email": row[3],
                "section": row[4],
                "github_id": row[5] if len(row) > 5 else "",
            }
        )
    return out
