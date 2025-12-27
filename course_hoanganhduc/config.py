"""Configuration helpers."""

import base64
import glob
import json
import os
import platform
import shutil
import sys
import threading
from datetime import datetime
from pathlib import Path

from .settings import *
from .utils import append_run_report


def _normalize_course_code(value):
    if value is None:
        return None
    value = str(value).strip()
    if not value:
        return None
    return value.lower()


def _prompt_course_code(timeout=60):
    """Prompt the user for a course code with a timeout."""
    result = {"value": None}

    def _read():
        try:
            result["value"] = input("Enter course code (e.g., MAT3500): ").strip()
        except (EOFError, KeyboardInterrupt):
            result["value"] = None

    thread = threading.Thread(target=_read, daemon=True)
    thread.start()
    thread.join(timeout)
    if thread.is_alive():
        return None
    return result["value"]

def _course_code_marker_path():
    return os.path.join(os.getcwd(), ".course_code")


def _load_cached_course_code():
    path = _course_code_marker_path()
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read().strip()
    except OSError:
        return None


def _save_cached_course_code(course_code):
    path = _course_code_marker_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(course_code)
    except OSError:
        pass


def cache_course_code(course_code):
    """
    Persist a normalized course code in the local .course_code cache.
    """
    course_code = _normalize_course_code(course_code) or _normalize_course_code(os.environ.get("COURSE_CODE"))
    if course_code:
        _save_cached_course_code(course_code)


def get_default_config_path(course_code=None, verbose=False):
    """
    Get the default config file path for the current operating system.
    The course_code controls the subfolder name under the base config directory.
    - Windows: %APPDATA%\course\<course_code>\config.json
    - macOS: ~/Library/Application Support/course/<course_code>/config.json
    - Linux: ~/.config/course/<course_code>/config.json
    If verbose is True, print details about the chosen path.
    Otherwise, print only an important notice if the config file does not exist.
    """
    course_code = _normalize_course_code(course_code)
    if not course_code:
        course_code = _normalize_course_code(_load_cached_course_code())
    if not course_code:
        course_code = _normalize_course_code(_prompt_course_code())
        if course_code:
            _save_cached_course_code(course_code)
    if not course_code:
        print("[Config] A course code is required to run this script. Exiting.")
        raise SystemExit(2)
    system = platform.system().lower()
    if system == "windows":
        appdata = os.environ.get("APPDATA", str(Path.home()))
        config_dir = os.path.join(appdata, "course", course_code)
    elif system == "darwin":  # macOS
        config_dir = os.path.join(str(Path.home()), "Library", "Application Support", "course", course_code)
    else:  # Linux and others
        config_dir = os.path.join(str(Path.home()), ".config", "course", course_code)
    os.makedirs(config_dir, exist_ok=True)
    config_path = os.path.join(config_dir, "config.json")
    if verbose:
        print(f"[Config] OS detected: {system}")
        print(f"[Config] Config directory: {config_dir}")
        print(f"[Config] Config file path: {config_path}")
        if not os.path.exists(config_path):
            print(f"[Config] Notice: Config file does not exist yet. You may need to create {config_path}")
    else:
        if not os.path.exists(config_path):
            print(f"Notice: Config file not found at {config_path}")
    return config_path


def get_default_credentials_path(course_code=None, verbose=False):
    """
    Get the default credentials file path alongside the config file.
    """
    config_path = get_default_config_path(course_code=course_code, verbose=verbose)
    return os.path.join(os.path.dirname(config_path), "credentials.json")


def get_default_token_path(course_code=None, verbose=False):
    """
    Get the default token file path alongside the config file.
    """
    config_path = get_default_config_path(course_code=course_code, verbose=verbose)
    return os.path.join(os.path.dirname(config_path), "token.pickle")


def _safe_remove(path, label=None, verbose=False):
    if not path:
        if verbose:
            print(f"[Config] {label or 'path'} not set; nothing to remove.")
        return False
    if os.path.exists(path):
        try:
            os.remove(path)
        except OSError as e:
            print(f"[Config] Failed to remove {label or 'file'} at {path}: {e}")
            return False
        if verbose:
            print(f"[Config] Removed {label or 'file'} at {path}")
        return True
    if verbose:
        print(f"[Config] {label or 'file'} not found at {path}")
    return False


def clear_config(config_path=None, course_code=None, verbose=False):
    """
    Remove the stored config file.
    """
    if not config_path:
        config_path = get_default_config_path(course_code=course_code, verbose=verbose)
    return _safe_remove(config_path, label="config file", verbose=verbose)


def clear_credentials(credentials_path=None, token_path=None, course_code=None, verbose=False):
    """
    Remove stored credentials and token files.
    """
    if not credentials_path:
        credentials_path = get_default_credentials_path(course_code=course_code, verbose=verbose)
    if not token_path:
        token_path = get_default_token_path(course_code=course_code, verbose=verbose)
    return {
        "credentials": _safe_remove(credentials_path, label="credentials file", verbose=verbose),
        "token": _safe_remove(token_path, label="token file", verbose=verbose),
    }


def _list_backups(backup_dir, base_name, ext):
    pattern = os.path.join(backup_dir, f"{base_name}_backup_*{ext}")
    return sorted(glob.glob(pattern), key=lambda p: os.path.getmtime(p))


def _cleanup_backups(backup_dir, base_name, ext, keep=5, verbose=False):
    if keep is None:
        return []
    try:
        keep = int(keep)
    except (TypeError, ValueError):
        return []
    backups = _list_backups(backup_dir, base_name, ext)
    if keep < 0 or len(backups) <= keep:
        return []
    to_remove = backups[:len(backups) - keep]
    removed = []
    for path in to_remove:
        try:
            os.remove(path)
            removed.append(path)
        except OSError as e:
            if verbose:
                print(f"[ConfigBackup] Failed to remove old backup {path}: {e}")
    return removed


def backup_config(config_path=None, backup_dir=None, keep=None, course_code=None, verbose=False):
    if not config_path:
        config_path = get_default_config_path(course_code=course_code, verbose=verbose)
    if not os.path.exists(config_path):
        if verbose:
            print(f"[ConfigBackup] Config not found at {config_path}")
        else:
            print(f"Config not found at {config_path}")
        return None
    backup_dir = backup_dir or os.path.dirname(config_path) or "."
    os.makedirs(backup_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(config_path))[0]
    ext = os.path.splitext(config_path)[1]
    now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(backup_dir, f"{base}_backup_{now_str}{ext}")
    if DRY_RUN:
        print(f"[ConfigBackup] Dry run: would back up config to {backup_path}")
        return backup_path
    try:
        shutil.copy2(config_path, backup_path)
        if verbose:
            print(f"[ConfigBackup] Backed up config to {backup_path}")
        else:
            print(f"Config backup created at {backup_path}")
        append_run_report("backup-config", outputs=backup_path, verbose=verbose)
    except OSError as e:
        print(f"[ConfigBackup] Failed to back up config: {e}")
        return None
    _cleanup_backups(backup_dir, base, ext, keep=keep if keep is not None else CONFIG_BACKUP_KEEP, verbose=verbose)
    return backup_path


def restore_config(config_path=None, backup_path=None, course_code=None, verbose=False):
    if not config_path:
        config_path = get_default_config_path(course_code=course_code, verbose=verbose)
    backup_dir = os.path.dirname(config_path) or "."
    base = os.path.splitext(os.path.basename(config_path))[0]
    ext = os.path.splitext(config_path)[1]
    if not backup_path or backup_path == "latest":
        backups = _list_backups(backup_dir, base, ext)
        if not backups:
            print("No config backups found.")
            return None
        backup_path = backups[-1]
    if not os.path.exists(backup_path):
        print(f"Config backup not found: {backup_path}")
        return None
    if DRY_RUN:
        print(f"[ConfigBackup] Dry run: would restore config from {backup_path}")
        return backup_path
    os.makedirs(os.path.dirname(config_path) or ".", exist_ok=True)
    try:
        shutil.copy2(backup_path, config_path)
    except OSError as e:
        print(f"[ConfigBackup] Failed to restore config: {e}")
        return None
    if verbose:
        print(f"[ConfigBackup] Restored config from {backup_path}")
    else:
        print(f"Config restored from {backup_path}")
    append_run_report("restore-config", outputs=backup_path, verbose=verbose)
    return backup_path


def load_config(config_path=None, verbose=False):
    """
    Load configuration from a JSON or base64-encoded JSON file at the default location and return config values as a dict.
    If config_path is not provided, loads from the OS-specific default location.
    config_path can be a file path (JSON or base64-encoded JSON) or a base64 string.
    Returns a dict of the loaded config values (does NOT set global variables).
    If verbose is True, print more details; otherwise, print only important notice.

    NOTE: If you want to update global variables, you must set:
        DEFAULT_AI_METHOD
        ALL_AI_METHODS
        GEMINI_API_KEY
        HUGGINGFACE_API_KEY
        GEMINI_DEFAULT_MODEL
        DEFAULT_OCR_METHOD
        ALL_OCR_METHODS
        OCRSPACE_API_KEY
        OCRSPACE_API_URL
        LOCAL_LLM_COMMAND
        LOCAL_LLM_MODEL
        LOCAL_LLM_ARGS
        LOCAL_LLM_TIMEOUT
        CANVAS_LMS_API_URL
        CANVAS_LMS_API_KEY
        CANVAS_LMS_COURSE_ID
        MIDTERM_DATE
        EXAM_TYPE
        CANVAS_MIDTERM_ASSIGNMENT_ID
        CANVAS_FINAL_ASSIGNMENT_ID
    """
    if config_path is None:
        config_path = get_default_config_path(verbose=verbose)
    config_data = None
    config = None

    if isinstance(config_path, str):
        # Try to treat as a file path first
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    config_data = f.read()
                # Try to parse as JSON first
                try:
                    config = json.loads(config_data)
                    if verbose:
                        print(f"[Config] Parsed as JSON from file: {config_path}")
                except Exception:
                    # If not valid JSON, try base64 decode then JSON
                    try:
                        json_str = base64.b64decode(config_data.encode("utf-8")).decode("utf-8")
                        config = json.loads(json_str)
                        if verbose:
                            print(f"[Config] Parsed as base64 JSON from file: {config_path}")
                    except Exception as e:
                        print(f"Failed to parse config file as JSON or base64: {e}")
                        return None
            except Exception as e:
                print(f"Failed to read config file: {e}")
                return None
        else:
            # Check if it looks like a file path (contains path separators)
            if os.sep in config_path or (os.name == 'nt' and '\\' in config_path):
                # It's a file path that doesn't exist, return None
                if verbose:
                    print(f"[Config] Config file not found at {config_path}. Using defaults.")
                else:
                    print(f"Notice: Config file not found at {config_path}. Using defaults.")
                return None
            else:
                # Treat as base64 string
                try:
                    json_str = base64.b64decode(config_path.encode("utf-8")).decode("utf-8")
                    config = json.loads(json_str)
                    if verbose:
                        print(f"[Config] Parsed as base64 JSON from string input.")
                except Exception as e:
                    print(f"Failed to parse config_path as base64 JSON: {e}")
                    return None
    else:
        print("Invalid config_path type. Must be str.")
        return None

    if not config:
        if verbose:
            print(f"[Config] No config data found at {config_path}. Using defaults.")
        else:
            print(f"Notice: No config data found at {config_path}. Using defaults.")
        return None

    # Only return config values, do not set globals
    result = {}
    known_keys = [
        "CONFIG_VERSION",
        "CREDENTIALS_PATH",
        "TOKEN_PATH",
        "DEFAULT_AI_METHOD",
        "ALL_AI_METHODS",
        "GEMINI_API_KEY",
        "HUGGINGFACE_API_KEY",
        "GEMINI_DEFAULT_MODEL",
        "REPORT_REFINE_METHOD",
        "LOCAL_LLM_COMMAND",
        "LOCAL_LLM_MODEL",
        "LOCAL_LLM_ARGS",
        "LOCAL_LLM_TIMEOUT",
        "LOCAL_LLM_GGUF_DIR",
        "DRY_RUN",
        "LOG_DIR",
        "LOG_LEVEL",
        "LOG_MAX_BYTES",
        "LOG_BACKUP_COUNT",
        "DB_BACKUP_KEEP",
        "CONFIG_BACKUP_KEEP",
        "GRADE_AUDIT_ENABLED",
        "GRADE_AUDIT_FIELDS",
        "DEFAULT_OCR_METHOD",
        "ALL_OCR_METHODS",
        "OCRSPACE_API_KEY",
        "OCRSPACE_API_URL",
        "CANVAS_LMS_API_URL",
        "CANVAS_LMS_API_KEY",
        "CANVAS_LMS_COURSE_ID",
        "CANVAS_DEFAULT_ASSIGNMENT_CATEGORY",
        "MIDTERM_DATE",
        "EXAM_TYPE",
        "CANVAS_MIDTERM_ASSIGNMENT_ID",
        "CANVAS_FINAL_ASSIGNMENT_ID",
    ]
    for key in known_keys:
        if key in config:
            result[key] = config.get(key)
    if verbose:
        print(f"[Config] Configuration loaded from {config_path}")
        for k, v in result.items():
            print(f"[Config] {k}: {v}")
    else:
        print(f"Configuration loaded from {config_path}")
    return result

def get_default_download_folder(verbose=False):
    """
    Get the default download folder for the current operating system.
    Returns the Downloads folder path appropriate for Windows, Mac, or Linux.
    If verbose is True, print details about the chosen path.
    Otherwise, print only an important notice if the folder does not exist.
    """
    system = platform.system().lower()
    downloads_path = Path.home() / "Downloads"

    if system == "windows":
        downloads_path = Path.home() / "Downloads"
    elif system == "darwin":  # macOS
        downloads_path = Path.home() / "Downloads"
    elif system == "linux":
        downloads_path = Path.home() / "Downloads"
        if not downloads_path.exists():
            downloads_path = Path.home() / "downloads"
            if not downloads_path.exists():
                downloads_path = Path.home()
    else:
        downloads_path = Path.home() / "Downloads"

    # Create the folder if it doesn't exist
    try:
        downloads_path.mkdir(exist_ok=True)
        if verbose:
            print(f"[DownloadFolder] OS detected: {system}")
            print(f"[DownloadFolder] Download folder: {downloads_path}")
    except Exception as e:
        if verbose:
            print(f"[DownloadFolder] Could not create Downloads folder: {e}")
            print(f"[DownloadFolder] Falling back to home directory: {Path.home()}")
        downloads_path = Path.home()
    if not downloads_path.exists():
        if verbose:
            print(f"[DownloadFolder] Notice: Download folder does not exist at {downloads_path}")
        else:
            print(f"Notice: Download folder not found at {downloads_path}")
    return str(downloads_path)

def get_default_db_path():
    return os.path.join(os.getcwd(), "students.db")

DEFAULT_DOWNLOAD_FOLDER = get_default_download_folder()
