# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas calendar helpers."""

from datetime import datetime, timedelta, timezone, date
import requests

from .settings import CANVAS_LMS_API_KEY, CANVAS_LMS_API_URL, CANVAS_LMS_COURSE_ID


def _unfold_ics_lines(lines):
    unfolded = []
    for line in lines:
        raw = line.rstrip("\r\n")
        if not raw:
            continue
        if raw.startswith((" ", "\t")) and unfolded:
            unfolded[-1] += raw.lstrip()
        else:
            unfolded.append(raw)
    return unfolded


def _parse_ics_datetime(raw_value):
    value = raw_value.strip()
    if not value:
        return None, False
    if "T" in value:
        if value.endswith("Z"):
            dt = datetime.strptime(value, "%Y%m%dT%H%M%SZ").replace(tzinfo=timezone.utc)
            return dt, False
        dt = datetime.strptime(value, "%Y%m%dT%H%M%S")
        return dt, False
    date_val = datetime.strptime(value, "%Y%m%d").date()
    return date_val, True


def _parse_ics_events(ics_path, verbose=False):
    try:
        with open(ics_path, "r", encoding="utf-8") as f:
            lines = _unfold_ics_lines(f.readlines())
    except Exception as exc:
        raise ValueError(f"Failed to read iCal file: {exc}") from exc

    events = []
    current = None
    for line in lines:
        if line == "BEGIN:VEVENT":
            current = {}
            continue
        if line == "END:VEVENT":
            if current:
                events.append(current)
            current = None
            continue
        if current is None:
            continue
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        key = key.split(";", 1)[0].upper()
        current[key] = value.strip()
    if verbose:
        print(f"[CanvasCalendar] Parsed {len(events)} event(s) from {ics_path}")
    return events


def _canvas_datetime_key(value):
    if not value:
        return ""
    try:
        if value.endswith("Z"):
            return value.replace("Z", "+00:00")
        return value
    except Exception:
        return str(value)


def _calendar_event_key(title, start_at, end_at, location, all_day):
    return (
        (title or "").strip().lower(),
        _canvas_datetime_key(start_at),
        _canvas_datetime_key(end_at),
        (location or "").strip().lower(),
        bool(all_day),
    )


def _fetch_canvas_calendar_events(api_url, api_key, course_id, start_at=None, end_at=None, verbose=False):
    headers = {"Authorization": f"Bearer {api_key}"}
    url = f"{api_url}/api/v1/calendar_events"
    params = {
        "context_codes[]": f"course_{course_id}",
        "per_page": 100,
    }
    if start_at:
        params["start_date"] = start_at
    if end_at:
        params["end_date"] = end_at
    events = []
    while url:
        resp = requests.get(url, headers=headers, params=params, timeout=30)
        resp.raise_for_status()
        batch = resp.json()
        if isinstance(batch, dict) and batch.get("calendar_events"):
            batch = batch["calendar_events"]
        if isinstance(batch, list):
            events.extend(batch)
        link = resp.headers.get("Link", "")
        next_url = None
        for part in link.split(","):
            if 'rel="next"' in part:
                next_url = part.split(";")[0].strip().strip("<>")
                break
        url = next_url
        params = None
    if verbose:
        print(f"[CanvasCalendar] Loaded {len(events)} existing event(s) from Canvas.")
    return events


def import_canvas_calendar_from_ics(
    ics_path,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    dry_run=False,
    skip_duplicates=True,
    verbose=False,
):
    """
    Import an iCal (.ics) file and create Canvas course calendar events.
    """
    if not api_url or not api_key or not course_id:
        raise ValueError("Canvas API URL, API key, and course ID are required.")
    events = _parse_ics_events(ics_path, verbose=verbose)
    if not events:
        print("[CanvasCalendar] No events found in iCal file.")
        return {"created": 0, "failed": 0}

    created = 0
    skipped = 0
    failed = 0
    headers = {"Authorization": f"Bearer {api_key}"}
    url = f"{api_url}/api/v1/calendar_events"
    context_code = f"course_{course_id}"
    existing_keys = set()
    if skip_duplicates:
        starts = []
        ends = []
        for event in events:
            start_raw = event.get("DTSTART")
            end_raw = event.get("DTEND")
            start_value, _ = _parse_ics_datetime(start_raw or "")
            end_value, _ = _parse_ics_datetime(end_raw or "")
            if isinstance(start_value, datetime):
                starts.append(start_value.isoformat())
            elif isinstance(start_value, date):
                starts.append(start_value.isoformat())
            if isinstance(end_value, datetime):
                ends.append(end_value.isoformat())
            elif isinstance(end_value, date):
                ends.append(end_value.isoformat())
        start_at = min(starts) if starts else None
        end_at = max(ends) if ends else None
        existing = _fetch_canvas_calendar_events(
            api_url,
            api_key,
            course_id,
            start_at=start_at,
            end_at=end_at,
            verbose=verbose,
        )
        for item in existing:
            existing_keys.add(
                _calendar_event_key(
                    item.get("title"),
                    item.get("start_at"),
                    item.get("end_at"),
                    item.get("location_name") or item.get("location"),
                    item.get("all_day"),
                )
            )

    for event in events:
        start_raw = event.get("DTSTART")
        end_raw = event.get("DTEND")
        start_value, start_all_day = _parse_ics_datetime(start_raw or "")
        end_value, end_all_day = _parse_ics_datetime(end_raw or "")
        if start_value is None:
            if verbose:
                print("[CanvasCalendar] Skipping event without DTSTART.")
            failed += 1
            continue
        all_day = start_all_day or end_all_day
        if isinstance(start_value, datetime):
            start_dt = start_value
        else:
            start_dt = datetime.combine(start_value, datetime.min.time())
        if isinstance(end_value, datetime):
            end_dt = end_value
        elif isinstance(end_value, date):
            end_dt = datetime.combine(end_value, datetime.max.time().replace(microsecond=0))
        else:
            end_dt = start_dt + timedelta(hours=1)
        event_key = _calendar_event_key(
            event.get("SUMMARY") or "Class session",
            start_dt.isoformat(),
            end_dt.isoformat(),
            event.get("LOCATION") or "",
            all_day,
        )
        if skip_duplicates and event_key in existing_keys:
            skipped += 1
            if verbose:
                print(f"[CanvasCalendar] Skipping duplicate: {event.get('SUMMARY') or 'Class session'}")
            continue
        payload = {
            "calendar_event": {
                "context_code": context_code,
                "title": event.get("SUMMARY") or "Class session",
                "start_at": start_dt.isoformat(),
                "end_at": end_dt.isoformat(),
                "location_name": event.get("LOCATION") or "",
                "description": event.get("DESCRIPTION") or "",
            }
        }
        if all_day:
            payload["calendar_event"]["all_day"] = True
        try:
            if dry_run:
                created += 1
                if verbose:
                    print(f"[CanvasCalendar] Dry run: {payload['calendar_event']['title']}")
            else:
                resp = requests.post(url, headers=headers, json=payload, timeout=30)
                resp.raise_for_status()
                created += 1
                if verbose:
                    print(f"[CanvasCalendar] Created: {payload['calendar_event']['title']}")
        except Exception as exc:
            failed += 1
            if verbose:
                print(f"[CanvasCalendar] Failed to create event: {exc}")
    if verbose:
        print(f"[CanvasCalendar] Created {created} event(s), skipped {skipped}, failed {failed}.")
    return {"created": created, "skipped": skipped, "failed": failed}
