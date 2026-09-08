# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Shared Canvas utility helpers."""

import os
import signal
from datetime import datetime


def ensure_dir(path):
    if not path:
        return None
    os.makedirs(path, exist_ok=True)
    return path


def normalize_datetime_str(value):
    if not value:
        return None
    value = str(value).strip()
    if not value:
        return None
    # Accept "YYYY-MM-DD HH:MM" and convert to Canvas ISO with Z.
    try:
        dt = datetime.strptime(value, "%Y-%m-%d %H:%M")
        return dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    except Exception:
        return value


def parse_selection(selection, total):
    if not selection:
        return []
    sel = selection.strip().lower()
    if sel in ("a", "all"):
        return list(range(1, total + 1))
    indices = set()
    for part in sel.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            try:
                start, end = part.split("-", 1)
                start = int(start)
                end = int(end)
                for i in range(start, end + 1):
                    if 1 <= i <= total:
                        indices.add(i)
            except Exception:
                continue
        else:
            try:
                idx = int(part)
                if 1 <= idx <= total:
                    indices.add(idx)
            except Exception:
                continue
    return sorted(indices)


def prompt_with_timeout(prompt, timeout=60, default=None, verbose=False):
    def timeout_handler(*_):
        if verbose:
            print("\n[Canvas] Timeout: No response after 60 seconds.")
        raise TimeoutError("User input timeout")

    if hasattr(signal, "SIGALRM"):
        signal.signal(signal.SIGALRM, timeout_handler)
        signal.alarm(timeout)
        try:
            result = input(prompt)
            signal.alarm(0)
            if not result and default is not None:
                return default
            return result
        except TimeoutError:
            signal.alarm(0)
            if default is not None:
                return default
            raise
        except KeyboardInterrupt:
            signal.alarm(0)
            raise
    else:
        result = input(prompt)
        if not result and default is not None:
            return default
        return result
