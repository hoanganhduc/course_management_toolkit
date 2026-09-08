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
    sheet_name=None,
    sheet_selection='first',
):
    """
    Download a Google Sheet as CSV using the first sheet, gid in URL, or specified sheet.
    Returns the output CSV path.
    
    Args:
        sheet_name: Specific sheet name to download
        sheet_selection: Mode for sheet selection - 'first' (default), 'select' (interactive), 'all', or 'merge'
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

    # Get all sheets for selection
    all_sheets = list_google_sheets(sheet_url, credentials_path, token_path, verbose=False)
    
    if not all_sheets:
        raise ValueError("Could not retrieve sheet information from Google Sheets.")
    
    # Determine which sheet(s) to download
    selected_sheet_titles = []
    
    if sheet_name:
        # Specific sheet requested
        for s in all_sheets:
            if s['title'] == sheet_name:
                selected_sheet_titles = [sheet_name]
                break
        if not selected_sheet_titles:
            available = ', '.join([s['title'] for s in all_sheets])
            raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {available}")
    elif sheet_selection == 'select':
        # Interactive selection
        print(f"\nAvailable sheets in Google Sheets document:")
        for i, s in enumerate(all_sheets, 1):
            print(f"  {i}. {s['title']}")
        
        choice = input(f"\nSelect sheet number(s) (1-{len(all_sheets)}, comma-separated for multiple, or 'all'): ").strip().lower()
        
        if choice == 'all':
            selected_sheet_titles = [s['title'] for s in all_sheets]
            print(f"[GSheets] Selected all {len(all_sheets)} sheets for merging")
        else:
            try:
                sheet_indices = [int(x.strip()) - 1 for x in choice.split(',')]
                for idx in sheet_indices:
                    if 0 <= idx < len(all_sheets):
                        selected_sheet_titles.append(all_sheets[idx]['title'])
                    else:
                        print(f"Warning: Sheet number {idx + 1} out of range, skipping")
                
                if not selected_sheet_titles:
                    print("No valid sheets selected. Using first sheet.")
                    selected_sheet_titles = [all_sheets[0]['title']]
            except ValueError:
                print("Invalid input. Using first sheet.")
                selected_sheet_titles = [all_sheets[0]['title']]
    elif sheet_selection in ['all', 'merge']:
        # Import all sheets
        selected_sheet_titles = [s['title'] for s in all_sheets]
        if verbose:
            print(f"[GSheets] Importing all {len(all_sheets)} sheets for merging")
    elif gid:
        # GID specified in URL - try to find matching sheet
        try:
            gid_int = int(gid)
            for s in all_sheets:
                if s['sheetId'] == gid_int:
                    selected_sheet_titles = [s['title']]
                    break
        except ValueError:
            pass
        if not selected_sheet_titles:
            selected_sheet_titles = [all_sheets[0]['title']]
    else:
        # Default: first sheet
        selected_sheet_titles = [all_sheets[0]['title']]
    
    if verbose or len(selected_sheet_titles) > 1:
        print(f"[GSheets] Selected sheet(s): {', '.join(selected_sheet_titles)}")

    # Download and merge sheets
    all_values = []
    for sheet_title in selected_sheet_titles:
        if len(selected_sheet_titles) > 1 and verbose:
            print(f"[GSheets] Downloading sheet: {sheet_title}")
        
        # Quote the sheet name to prevent interpretation as cell reference
        # (e.g., "DN2022" would be interpreted as column DN, row 2022 without quotes)
        quoted_sheet = f"'{sheet_title}'"
        
        values = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=quoted_sheet,
        ).execute().get("values", []) or []
        
        if values:
            all_values.append(values)
    
    if not all_values:
        raise ValueError("No data found in selected Google Sheet(s).")
    
    # Merge all sheets
    if len(all_values) > 1:
        if verbose:
            print(f"[GSheets] Merging {len(all_values)} sheets...")
        # Concatenate all rows from all sheets
        merged_values = []
        for values in all_values:
            merged_values.extend(values)
        values = merged_values
    else:
        values = all_values[0]

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


def list_google_sheets(sheet_url, credentials_path='gclassroom_credentials.json', token_path='token.pickle', verbose=False):
    """
    List all sheets in a Google Sheets document.
    Returns a list of dicts with 'title' and 'sheetId' for each sheet.
    """
    if not sheet_url:
        raise ValueError("Google Sheet URL is required.")
    
    def _extract_sheet_id(value):
        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", value or "")
        return match.group(1) if match else None
    
    sheet_id = _extract_sheet_id(sheet_url)
    if not sheet_id:
        raise ValueError("Could not parse spreadsheet ID from URL.")
    
    try:
        creds = _get_google_classroom_credentials(credentials_path, token_path, verbose=verbose)
        service = build("sheets", "v4", credentials=creds)
        
        meta = service.spreadsheets().get(
            spreadsheetId=sheet_id,
            fields="sheets(properties(sheetId,title,index))"
        ).execute()
        
        sheets = meta.get("sheets", []) or []
        sheet_list = []
        for s in sheets:
            props = s.get("properties", {}) or {}
            sheet_list.append({
                "title": props.get("title", ""),
                "sheetId": props.get("sheetId", 0),
                "index": props.get("index", 0)
            })
        
        # Sort by index
        sheet_list.sort(key=lambda x: x["index"])
        
        if verbose:
            print(f"[GSheets] Found {len(sheet_list)} sheets:")
            for s in sheet_list:
                print(f"  - {s['title']} (ID: {s['sheetId']})")
        
        return sheet_list
    except Exception as e:
        if verbose:
            print(f"[GSheets] Error listing sheets: {e}")
        return []


def get_google_file_metadata(file_id, credentials_path='gclassroom_credentials.json', token_path='token.pickle', verbose=False):
    """
    Get metadata for a file from Google Drive API.
    Returns a dict with 'id', 'name', 'modifiedTime', 'webViewLink', etc.
    """
    if not file_id:
        return None
    
    try:
        creds = _get_google_classroom_credentials(credentials_path, token_path, verbose=verbose)
        if not creds:
            if verbose:
                print("[GDrive] No valid credentials found.")
            return None
            
        service = build("drive", "v3", credentials=creds)
        file_meta = service.files().get(
            fileId=file_id, 
            fields="id, name, modifiedTime, webViewLink, owners"
        ).execute()
        return file_meta
    except Exception as e:
        if verbose:
            print(f"[GDrive] Error fetching metadata for {file_id}: {e}")
        return None
