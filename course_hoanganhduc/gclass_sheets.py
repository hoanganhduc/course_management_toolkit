# -*- coding: utf-8 -*-

import os
import re
import csv
from urllib.parse import urlparse, parse_qs
from googleapiclient.discovery import build
from .gclass_auth import _get_google_classroom_credentials

def download_google_sheet_to_csv(
    sheet_url,
    output_path=None,
    credentials_path='gclassroom_credentials.json',
    token_path='token.pickle',
    verbose=False,
):
    """
    Download a Google Sheet as CSV using the first sheet or gid in the URL.
    Returns the output CSV path.
    """
    if not sheet_url:
        raise ValueError("Google Sheet URL is required.")

    def _extract_sheet_id(value):
        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", value or "")
        return match.group(1) if match else None

    def _extract_gid(value):
        parsed = urlparse(value)
        if parsed.fragment:
            frag_qs = parse_qs(parsed.fragment)
            if "gid" in frag_qs:
                return frag_qs["gid"][0]
            if parsed.fragment.startswith("gid="):
                return parsed.fragment.split("gid=", 1)[-1]
        query = parse_qs(parsed.query)
        if "gid" in query:
            return query["gid"][0]
        return None

    sheet_id = _extract_sheet_id(sheet_url)
    if not sheet_id:
        raise ValueError("Could not parse spreadsheet ID from URL.")
    gid = _extract_gid(sheet_url)

    creds = _get_google_classroom_credentials(credentials_path, token_path, verbose=verbose)
    service = build("sheets", "v4", credentials=creds)

    sheet_title = None
    try:
        meta = service.spreadsheets().get(
            spreadsheetId=sheet_id,
            fields="sheets(properties(sheetId,title))"
        ).execute()
        sheets = meta.get("sheets", []) or []
        if gid:
            try:
                gid_int = int(gid)
            except Exception:
                gid_int = None
            if gid_int is not None:
                for s in sheets:
                    props = s.get("properties", {}) or {}
                    if props.get("sheetId") == gid_int:
                        sheet_title = props.get("title")
                        break
        if not sheet_title and sheets:
            sheet_title = (sheets[0].get("properties") or {}).get("title")
    except Exception as e:
        if verbose:
            print(f"[GSheets] Failed to resolve sheet title: {e}")

    range_name = sheet_title or "Sheet1"
    values = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=range_name,
    ).execute().get("values", []) or []
    if not values:
        raise ValueError("No data found in Google Sheet.")

    max_cols = max(len(row) for row in values)
    normalized_rows = [row + [""] * (max_cols - len(row)) for row in values]

    if not output_path:
        output_path = os.path.join(os.getcwd(), "google_sheet.csv")
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(normalized_rows)
    if verbose:
        print(f"[GSheets] Saved CSV to {output_path}")
    return output_path
