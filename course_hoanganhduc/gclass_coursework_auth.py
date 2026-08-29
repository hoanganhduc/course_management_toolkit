# -*- coding: utf-8 -*-
"""Isolated credentials for human Google Classroom coursework mutations.

This module intentionally does not reuse the legacy pickle token or course-scoped
credential resolver. Google dependencies are imported only when online auth starts.
"""

from __future__ import annotations

import errno
import hashlib
import inspect
import json
import logging
import os
import platform
import re
import secrets
import stat
import threading
from contextlib import AbstractContextManager, contextmanager
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, Mapping, MutableMapping, Optional, Tuple
from urllib.parse import urlsplit


COURSEWORK_SCOPES: Tuple[str, ...] = (
    "https://www.googleapis.com/auth/classroom.coursework.students",
    "https://www.googleapis.com/auth/classroom.courses.readonly",
    "https://www.googleapis.com/auth/userinfo.email",
)
_REQUIRED_COURSEWORK_SCOPES = frozenset(COURSEWORK_SCOPES)
_GOOGLE_ADDITIONAL_IDENTITY_SCOPES: Tuple[str, ...] = ("email", "openid")
_GOOGLE_GRANT_SCOPE_ORDER = COURSEWORK_SCOPES + ("email", "openid")
GOOGLE_AUTH_URI = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_TOKEN_URI = "https://oauth2.googleapis.com/token"
GOOGLE_USERINFO_URI = "https://www.googleapis.com/oauth2/v2/userinfo"
MAX_CREDENTIAL_BYTES = 256 * 1024
MAX_TOKEN_BYTES = 64 * 1024
MAX_GOOGLE_API_RESPONSE_BYTES = 8 * 1024 * 1024
GOOGLE_API_TIMEOUT_SECONDS = (10, 60)
_OAUTH_SECRET_LOGGER_NAMES = (
    "google_auth_oauthlib.flow",
    "requests_oauthlib.oauth2_session",
)
_OAUTH_SECRET_LOGGING_LOCK = threading.Lock()

_ACCOUNT_RE = re.compile(r"^[^\s@]+@[^\s@]+$")
_CLIENT_ID_RE = re.compile(r"^[A-Za-z0-9._-]+\.apps\.googleusercontent\.com$")


class CredentialSecurityError(RuntimeError):
    """Credential location, metadata, or content is unsafe."""


class UnsupportedPlatformCredentialSecurityError(CredentialSecurityError):
    """Native credential guarantees are not implemented for this platform."""


class _GoogleAPIResponse(dict):
    """Minimal httplib2-style response consumed by google-api-python-client."""

    def __init__(self, response: Any):
        super().__init__(
            (str(key).lower(), str(value))
            for key, value in response.headers.items()
        )
        self.status = int(response.status_code)
        self.reason = str(response.reason or "")


class _OneShotAuthorizedHttp:
    """httplib2-compatible adapter with no response replay or redirect following."""

    _ALLOWED_HOSTS = frozenset(
        {
            "classroom.googleapis.com",
            "www.googleapis.com",
        }
    )
    _ALLOWED_METHODS = frozenset({"GET", "POST", "PATCH"})

    def __init__(self, session: Any):
        self._session = session

    def request(
        self,
        uri: str,
        method: str = "GET",
        body: Any = None,
        headers: Optional[Mapping[str, str]] = None,
        redirections: int = 0,
        connection_type: Any = None,
        **kwargs: Any,
    ) -> Tuple[_GoogleAPIResponse, bytes]:
        del redirections, connection_type
        if kwargs:
            raise CredentialSecurityError(
                "Google API transport received unsupported request options"
            )
        parsed = urlsplit(uri)
        try:
            port = parsed.port
        except ValueError:
            raise CredentialSecurityError("Google API request URL is invalid") from None
        if (
            parsed.scheme != "https"
            or parsed.hostname not in self._ALLOWED_HOSTS
            or port not in (None, 443)
            or parsed.username is not None
            or parsed.password is not None
            or parsed.fragment
        ):
            raise CredentialSecurityError("Google API request URL is not allowlisted")
        normalized_method = str(method).upper()
        if normalized_method not in self._ALLOWED_METHODS:
            raise CredentialSecurityError("Google API request method is not allowlisted")
        response = self._session.request(
            normalized_method,
            uri,
            data=body,
            headers=dict(headers or {}),
            timeout=GOOGLE_API_TIMEOUT_SECONDS,
            allow_redirects=False,
        )
        content = bytes(response.content)
        if len(content) > MAX_GOOGLE_API_RESPONSE_BYTES:
            raise CredentialSecurityError("Google API response exceeded the size limit")
        return _GoogleAPIResponse(response), content

    def close(self) -> None:
        self._session.close()


def _build_one_shot_authorized_http(
    credentials: Any,
    *,
    authorized_session_cls: Any,
) -> _OneShotAuthorizedHttp:
    """Build a transport that sends each Classroom HTTP request at most once."""

    try:
        import requests
        from urllib3.util import Retry
    except ImportError as exc:
        raise CredentialSecurityError(
            "Installed requests/urllib3 cannot provide the one-shot Google transport"
        ) from exc
    signature = inspect.signature(authorized_session_cls)
    if not {"refresh_status_codes", "max_refresh_attempts"}.issubset(
        signature.parameters
    ):
        raise CredentialSecurityError(
            "Installed google-auth cannot disable response-triggered request replay"
        )
    session = authorized_session_cls(
        credentials,
        refresh_status_codes=(),
        max_refresh_attempts=0,
    )
    retry_policy = Retry(
        total=0,
        connect=0,
        read=0,
        redirect=0,
        status=0,
        other=0,
    )
    for prefix in ("https://", "http://"):
        previous_adapter = session.adapters.get(prefix)
        session.mount(
            prefix,
            requests.adapters.HTTPAdapter(max_retries=retry_policy),
        )
        if previous_adapter is not None:
            previous_adapter.close()
    return _OneShotAuthorizedHttp(session)


@contextmanager
def _suppress_oauth_secret_logging():
    """Prevent dependency loggers from emitting callbacks or bearer tokens."""

    with _OAUTH_SECRET_LOGGING_LOCK:
        loggers = [logging.getLogger(name) for name in _OAUTH_SECRET_LOGGER_NAMES]
        previous = [logger.disabled for logger in loggers]
        try:
            for logger in loggers:
                logger.disabled = True
            yield
        finally:
            for logger, disabled in zip(loggers, previous):
                logger.disabled = disabled


@dataclass(frozen=True)
class CourseworkAuthPaths:
    account: str
    account_fingerprint: str
    credentials_path: Optional[Path]
    token_path: Path
    metadata_path: Path
    source: str
    token_only: bool


@dataclass(frozen=True)
class CourseworkAuthSession:
    service: Any
    profile: Mapping[str, Any]
    paths: CourseworkAuthPaths
    client_id_fingerprint: str


def normalize_expected_account(account: str) -> str:
    if not isinstance(account, str):
        raise ValueError("Expected Google account must be a primary email address")
    normalized = account.strip().lower()
    if not _ACCOUNT_RE.fullmatch(normalized):
        raise ValueError("Expected Google account must be a primary email address")
    if any(ord(char) < 33 or ord(char) == 127 for char in normalized):
        raise ValueError("Expected Google account contains unsafe characters")
    return normalized


def account_fingerprint(account: str) -> str:
    normalized = normalize_expected_account(account)
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:20]


def _fingerprint(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()[:20]


def _expand_absolute_path(value: str, *, home: Path) -> Path:
    if not isinstance(value, str) or not value.strip():
        raise ValueError("Credential paths must be nonempty")
    value = value.strip()
    if value == "~":
        value = str(home)
    elif value.startswith("~/") or value.startswith("~\\"):
        value = str(home / value[2:])
    elif value.startswith("~"):
        raise ValueError("Only the current user's ~ path is supported")
    path = Path(value)
    if not path.is_absolute():
        raise ValueError("Credential paths must be absolute")
    return Path(os.path.normpath(str(path)))


def _path_is_in_git_worktree(path: Path) -> bool:
    """Detect a conventional repository/worktree without invoking Git."""

    current = Path(path).parent
    while True:
        try:
            os.lstat(current / ".git")
            return True
        except FileNotFoundError:
            pass
        parent = current.parent
        if parent == current:
            return False
        current = parent


def _default_auth_root(
    *,
    env: Mapping[str, str],
    platform_system: str,
    home: Path,
) -> Path:
    system = platform_system.lower()
    if system == "windows":
        appdata = (env.get("APPDATA") or "").strip()
        if not appdata:
            raise ValueError("APPDATA is required for the Windows credential location")
        root = _expand_absolute_path(appdata, home=home)
        return root / "course" / "google-classroom"
    if system == "darwin":
        return home / "Library" / "Application Support" / "course" / "google-classroom"
    xdg = (env.get("XDG_CONFIG_HOME") or "").strip()
    if xdg:
        root = _expand_absolute_path(xdg, home=home)
    else:
        root = home / ".config"
    return root / "course" / "google-classroom"


def resolve_coursework_auth_paths(
    account: str,
    *,
    credentials_path: Optional[str] = None,
    token_path: Optional[str] = None,
    env: Optional[Mapping[str, str]] = None,
    platform_system: Optional[str] = None,
    home: Optional[Path] = None,
) -> CourseworkAuthPaths:
    """Resolve one deterministic CLI, environment, or default credential pair."""

    normalized_account = normalize_expected_account(account)
    fingerprint = account_fingerprint(normalized_account)
    environment = dict(os.environ if env is None else env)
    home_path = Path.home() if home is None else Path(home)
    if not home_path.is_absolute():
        raise ValueError("The home directory used for credential resolution must be absolute")
    system = (platform_system or platform.system()).lower()

    cli_credentials = credentials_path.strip() if isinstance(credentials_path, str) else None
    cli_token = token_path.strip() if isinstance(token_path, str) else None
    if cli_credentials or cli_token:
        source = "cli"
        selected_credentials = cli_credentials
        selected_token = cli_token
    else:
        env_credentials = (environment.get("COURSE_GCLASS_CREDENTIALS") or "").strip()
        env_token = (environment.get("COURSE_GCLASS_COURSEWORK_TOKEN") or "").strip()
        if env_credentials or env_token:
            source = "environment"
            selected_credentials = env_credentials or None
            selected_token = env_token or None
        else:
            source = "default"
            root = _default_auth_root(
                env=environment,
                platform_system=system,
                home=home_path,
            )
            selected_credentials = str(root / "credentials.json")
            selected_token = str(root / "tokens" / f"{fingerprint}.json")

    if selected_credentials and not selected_token:
        raise ValueError(
            "A custom credentials path requires an explicit token path; token-only mode is allowed"
        )
    if not selected_token:
        raise ValueError("A token path is required")
    credential = (
        _expand_absolute_path(selected_credentials, home=home_path)
        if selected_credentials
        else None
    )
    token = _expand_absolute_path(selected_token, home=home_path)
    metadata = token.with_suffix(".meta.json")
    if not token.name:
        raise ValueError("The token path must name a file")
    if credential is not None and not credential.name:
        raise ValueError("The credentials path must name a file")
    lock_path = token.parent / f".{token.name}.lock"
    protected_outputs = {os.path.normcase(str(token)), os.path.normcase(str(metadata))}
    if credential is not None:
        normalized_credential = os.path.normcase(str(credential))
        if normalized_credential in protected_outputs or normalized_credential == os.path.normcase(
            str(lock_path)
        ):
            raise ValueError(
                "Credentials, token, metadata, and lock paths must be distinct"
            )
    if _path_is_in_git_worktree(token) or (
        credential is not None and _path_is_in_git_worktree(credential)
    ):
        raise ValueError(
            "Coursework OAuth credentials and tokens must be stored outside Git worktrees"
        )
    return CourseworkAuthPaths(
        account=normalized_account,
        account_fingerprint=fingerprint,
        credentials_path=credential,
        token_path=token,
        metadata_path=metadata,
        source=source,
        token_only=credential is None,
    )


def _require_posix_security() -> None:
    if os.name == "nt":
        raise UnsupportedPlatformCredentialSecurityError(
            "Secure Google Classroom mutation credentials are not supported on native Windows yet"
        )


def _safe_pairs(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise CredentialSecurityError(f"Duplicate JSON key: {key}")
        result[key] = value
    return result


def _validate_directory_stat(info: os.stat_result, label: str) -> None:
    uid = os.getuid()
    if not stat.S_ISDIR(info.st_mode):
        raise CredentialSecurityError(f"{label} is not a directory")
    if info.st_uid not in {0, uid}:
        raise CredentialSecurityError(f"{label} is not owned by the current user or root")
    writable = info.st_mode & 0o022
    root_sticky = info.st_uid == 0 and bool(info.st_mode & stat.S_ISVTX)
    if writable and not root_sticky:
        raise CredentialSecurityError(f"{label} is writable by another user")


def _open_directory_descriptor(path: Path, *, create: bool = False) -> int:
    _require_posix_security()
    path = Path(path)
    if not path.is_absolute():
        raise CredentialSecurityError("Secure paths must be absolute")
    flags = os.O_RDONLY | getattr(os, "O_DIRECTORY", 0) | getattr(os, "O_CLOEXEC", 0)
    nofollow = getattr(os, "O_NOFOLLOW", 0)
    descriptor = os.open(path.anchor or "/", flags)
    try:
        _validate_directory_stat(os.fstat(descriptor), path.anchor or "/")
        for part in path.parts[1:]:
            try:
                next_descriptor = os.open(part, flags | nofollow, dir_fd=descriptor)
            except FileNotFoundError:
                if not create:
                    raise CredentialSecurityError(f"Credential directory does not exist: {part}") from None
                try:
                    os.mkdir(part, mode=0o700, dir_fd=descriptor)
                except FileExistsError:
                    pass
                next_descriptor = os.open(part, flags | nofollow, dir_fd=descriptor)
            except OSError as exc:
                if exc.errno in {errno.ELOOP, errno.ENOTDIR}:
                    raise CredentialSecurityError("Credential path contains a symlink or non-directory") from None
                raise
            os.close(descriptor)
            descriptor = next_descriptor
            _validate_directory_stat(os.fstat(descriptor), str(path))
        return descriptor
    except BaseException:
        os.close(descriptor)
        raise


def _read_regular_file(path: Path, *, maximum_bytes: int) -> bytes:
    _require_posix_security()
    path = Path(path)
    parent = _open_directory_descriptor(path.parent, create=False)
    descriptor = None
    try:
        flags = (
            os.O_RDONLY
            | getattr(os, "O_CLOEXEC", 0)
            | getattr(os, "O_NOFOLLOW", 0)
            | getattr(os, "O_NONBLOCK", 0)
        )
        try:
            descriptor = os.open(path.name, flags, dir_fd=parent)
        except OSError as exc:
            if exc.errno in {errno.ELOOP, errno.ENOTDIR}:
                raise CredentialSecurityError("Credential file must not be a symlink") from None
            raise CredentialSecurityError("Credential file could not be opened safely") from None
        info = os.fstat(descriptor)
        if not stat.S_ISREG(info.st_mode):
            raise CredentialSecurityError("Credential input must be a regular file")
        if info.st_uid != os.getuid():
            raise CredentialSecurityError("Credential file must be owned by the current user")
        if info.st_mode & 0o077:
            raise CredentialSecurityError("Credential file must not grant group or other permissions")
        if info.st_size > maximum_bytes:
            raise CredentialSecurityError("Credential file exceeds the size limit")
        chunks = []
        total = 0
        while True:
            chunk = os.read(descriptor, min(65536, maximum_bytes + 1 - total))
            if not chunk:
                break
            chunks.append(chunk)
            total += len(chunk)
            if total > maximum_bytes:
                raise CredentialSecurityError("Credential file exceeds the size limit")
        return b"".join(chunks)
    finally:
        if descriptor is not None:
            os.close(descriptor)
        os.close(parent)


def secure_read_json_file(
    path: Path,
    *,
    maximum_bytes: int = MAX_TOKEN_BYTES,
) -> Dict[str, Any]:
    try:
        payload = _read_regular_file(Path(path), maximum_bytes=maximum_bytes)
        text = payload.decode("utf-8")
        value = json.loads(text, object_pairs_hook=_safe_pairs)
    except CredentialSecurityError:
        raise
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise CredentialSecurityError("Credential file is not valid strict UTF-8 JSON") from exc
    if not isinstance(value, dict):
        raise CredentialSecurityError("Credential JSON must be an object")
    return value


def _write_all(descriptor: int, data: bytes) -> None:
    view = memoryview(data)
    while view:
        written = os.write(descriptor, view)
        view = view[written:]


def write_private_json_atomic(path: Path, value: Mapping[str, Any]) -> None:
    _require_posix_security()
    path = Path(path)
    payload = (
        json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n"
    ).encode("utf-8")
    if len(payload) > MAX_TOKEN_BYTES:
        raise CredentialSecurityError("Credential JSON exceeds the size limit")
    parent = _open_directory_descriptor(path.parent, create=True)
    temporary_name = f".{path.name}.{secrets.token_hex(8)}.tmp"
    descriptor = None
    try:
        flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, "O_CLOEXEC", 0)
        descriptor = os.open(temporary_name, flags, 0o600, dir_fd=parent)
        _write_all(descriptor, payload)
        os.fsync(descriptor)
        os.close(descriptor)
        descriptor = None
        os.replace(temporary_name, path.name, src_dir_fd=parent, dst_dir_fd=parent)
        os.fsync(parent)
    except BaseException:
        if descriptor is not None:
            os.close(descriptor)
        try:
            os.unlink(temporary_name, dir_fd=parent)
        except OSError:
            pass
        raise
    finally:
        os.close(parent)


class TokenFileLock(AbstractContextManager):
    """Nonblocking POSIX lock preventing concurrent token refresh/write."""

    def __init__(self, token_path: Path):
        self.token_path = Path(token_path)
        self._directory_fd: Optional[int] = None
        self._lock_fd: Optional[int] = None

    def __enter__(self):
        _require_posix_security()
        import fcntl

        self._directory_fd = _open_directory_descriptor(self.token_path.parent, create=True)
        flags = os.O_RDWR | os.O_CREAT | getattr(os, "O_CLOEXEC", 0) | getattr(os, "O_NOFOLLOW", 0)
        lock_name = f".{self.token_path.name}.lock"
        try:
            self._lock_fd = os.open(
                lock_name, flags, 0o600, dir_fd=self._directory_fd
            )
        except OSError as exc:
            if exc.errno in {errno.ELOOP, errno.ENOTDIR}:
                self.__exit__(None, None, None)
                raise CredentialSecurityError(
                    "Credential lock must be a regular private file"
                ) from None
            self.__exit__(None, None, None)
            raise
        lock_info = os.fstat(self._lock_fd)
        if (
            not stat.S_ISREG(lock_info.st_mode)
            or lock_info.st_uid != os.getuid()
            or lock_info.st_mode & 0o077
        ):
            self.__exit__(None, None, None)
            raise CredentialSecurityError(
                "Credential lock must be a regular private file owned by the current user"
            )
        try:
            fcntl.flock(self._lock_fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            self.__exit__(None, None, None)
            raise CredentialSecurityError("Credential profile is already in use") from None
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._lock_fd is not None:
            try:
                import fcntl

                fcntl.flock(self._lock_fd, fcntl.LOCK_UN)
            finally:
                os.close(self._lock_fd)
                self._lock_fd = None
        if self._directory_fd is not None:
            os.close(self._directory_fd)
            self._directory_fd = None
        return False


def load_sanitized_client_config(path: Path) -> Dict[str, Any]:
    raw = secure_read_json_file(path, maximum_bytes=MAX_CREDENTIAL_BYTES)
    if set(raw) != {"installed"} or not isinstance(raw.get("installed"), Mapping):
        raise CredentialSecurityError(
            "credentials.json must contain exactly one installed desktop OAuth client"
        )
    installed = raw["installed"]
    client_id = installed.get("client_id")
    client_secret = installed.get("client_secret")
    if not isinstance(client_id, str) or not _CLIENT_ID_RE.fullmatch(client_id):
        raise CredentialSecurityError("Installed OAuth client ID is invalid")
    if not isinstance(client_secret, str) or not client_secret or len(client_secret) > 4096:
        raise CredentialSecurityError("Installed OAuth client secret is invalid")
    return {
        "installed": {
            "client_id": client_id,
            "client_secret": client_secret,
            "auth_uri": GOOGLE_AUTH_URI,
            "token_uri": GOOGLE_TOKEN_URI,
        }
    }


def _normalize_token_expiry(value: Any) -> str:
    if not isinstance(value, str) or not value or len(value) > 128:
        raise CredentialSecurityError("Stored token expiry is invalid")
    candidate = value.strip()
    if candidate.endswith("Z"):
        candidate = candidate[:-1]
        if not re.search(r"[+-]\d{2}:\d{2}$", candidate):
            candidate += "+00:00"
    try:
        parsed = datetime.fromisoformat(candidate)
    except ValueError:
        raise CredentialSecurityError("Stored token expiry is invalid") from None
    if parsed.tzinfo is None:
        raise CredentialSecurityError("Stored token expiry must include a timezone")
    try:
        utc = parsed.astimezone(timezone.utc).replace(tzinfo=None)
    except (OverflowError, ValueError):
        raise CredentialSecurityError(
            "Stored token expiry is outside the supported datetime range"
        ) from None
    timespec = "microseconds" if utc.microsecond else "seconds"
    return utc.isoformat(timespec=timespec) + "Z"


def _sanitize_token_mapping(raw: Mapping[str, Any]) -> Dict[str, Any]:
    scopes = raw.get("scopes")
    if (
        not isinstance(scopes, list)
        or len(scopes) != len(COURSEWORK_SCOPES)
        or any(not isinstance(scope, str) for scope in scopes)
        or set(scopes) != set(COURSEWORK_SCOPES)
    ):
        raise CredentialSecurityError("Stored token scopes do not match coursework scopes")
    client_id = raw.get("client_id")
    client_secret = raw.get("client_secret")
    if not isinstance(client_id, str) or not _CLIENT_ID_RE.fullmatch(client_id):
        raise CredentialSecurityError("Stored token client ID is invalid")
    if (
        not isinstance(client_secret, str)
        or not client_secret
        or len(client_secret) > 4096
    ):
        raise CredentialSecurityError("Stored token client secret is missing")
    token = raw.get("token")
    refresh_token = raw.get("refresh_token")
    for value in (token, refresh_token):
        if value is not None and (not isinstance(value, str) or not value):
            raise CredentialSecurityError("Stored token credentials are invalid")
    if not token and not refresh_token:
        raise CredentialSecurityError("Stored token contains no access or refresh token")
    granted_scopes = _normalize_granted_scopes(raw.get("granted_scopes"))
    if granted_scopes is None:
        raise CredentialSecurityError(
            "Stored token lacks verified coursework grants; authorize a replacement"
        )
    clean = {
        "token": token,
        "refresh_token": refresh_token,
        "token_uri": GOOGLE_TOKEN_URI,
        "client_id": client_id,
        "client_secret": client_secret,
        "scopes": list(COURSEWORK_SCOPES),
        "granted_scopes": granted_scopes,
    }
    expiry = raw.get("expiry")
    if expiry is not None:
        clean["expiry"] = _normalize_token_expiry(expiry)
    return clean


def load_sanitized_token_info(path: Path) -> Dict[str, Any]:
    raw = secure_read_json_file(path, maximum_bytes=MAX_TOKEN_BYTES)
    return _sanitize_token_mapping(raw)


def _file_exists_without_following(path: Optional[Path]) -> bool:
    if path is None:
        return False
    try:
        os.lstat(path)
        return True
    except FileNotFoundError:
        return False


def get_google_classroom_auth_status(
    paths: CourseworkAuthPaths,
    *,
    show_paths: bool = False,
) -> Dict[str, Any]:
    """Inspect credential metadata without refresh, OAuth, network, or writes."""

    status: Dict[str, Any] = {
        "source": paths.source,
        "account_fingerprint": paths.account_fingerprint,
        "principal_status": "offline/unverified",
        "credentials_exists": _file_exists_without_following(paths.credentials_path),
        "token_exists": _file_exists_without_following(paths.token_path),
        "credentials_safe": None,
        "token_safe": None,
        "token_usable": None,
        "scopes_match": None,
        "declared_scopes_match": None,
        "granted_scopes_verified": None,
    }
    if show_paths:
        status["credentials_path"] = str(paths.credentials_path) if paths.credentials_path else None
        status["token_path"] = str(paths.token_path)
    else:
        status["credentials_name"] = paths.credentials_path.name if paths.credentials_path else None
        status["token_name"] = paths.token_path.name

    if status["credentials_exists"] and paths.credentials_path is not None:
        try:
            client = load_sanitized_client_config(paths.credentials_path)
            status["credentials_safe"] = True
            status["client_id_fingerprint"] = _fingerprint(client["installed"]["client_id"])
        except CredentialSecurityError:
            status["credentials_safe"] = False
    if status["token_exists"]:
        try:
            token = load_sanitized_token_info(paths.token_path)
            status["token_safe"] = True
            status["scopes_match"] = True
            status["declared_scopes_match"] = True
            status["granted_scopes_verified"] = True
            status["token_client_id_fingerprint"] = _fingerprint(token["client_id"])
            status["expiry"] = token.get("expiry")
            refreshable = bool(token.get("refresh_token"))
            expired = None
            expires_soon = None
            if token.get("expiry"):
                expiry = datetime.fromisoformat(
                    token["expiry"].replace("Z", "+00:00")
                )
                status_now = datetime.now(timezone.utc)
                expired = expiry <= status_now
                expires_soon = expiry <= status_now + timedelta(minutes=5)
            access_usable = bool(token.get("token")) and expires_soon is False
            refresh_usable = refreshable and (
                token.get("expiry") is None or expired is True
            )
            status["token_refreshable"] = refreshable
            status["token_expired"] = expired
            status["token_expiring_soon"] = expires_soon
            status["token_usable"] = bool(access_usable or refresh_usable)
        except CredentialSecurityError:
            status["token_safe"] = False
            status["scopes_match"] = False
            status["declared_scopes_match"] = False
            status["granted_scopes_verified"] = False
            status["token_usable"] = False
    if status.get("client_id_fingerprint") and status.get(
        "token_client_id_fingerprint"
    ):
        status["client_ids_match"] = (
            status["client_id_fingerprint"]
            == status["token_client_id_fingerprint"]
        )
    if _file_exists_without_following(paths.metadata_path):
        try:
            metadata = secure_read_json_file(paths.metadata_path)
            status["metadata_safe"] = True
            status["last_verified_at"] = metadata.get("verified_at")
            status["verified_user_id_fingerprint"] = metadata.get("user_id_fingerprint")
        except CredentialSecurityError:
            status["metadata_safe"] = False
    return status


def _normalize_granted_scopes(
    value: Any,
    *,
    allow_omitted_equal: bool = False,
) -> Optional[list]:
    if value is None or value == [] or value == ():
        return list(COURSEWORK_SCOPES) if allow_omitted_equal else None
    if isinstance(value, str):
        value = value.split()
    if not isinstance(value, (list, tuple, set, frozenset)) or any(
        not isinstance(scope, str) for scope in value
    ):
        raise CredentialSecurityError("Authorized grant scope evidence is invalid")
    scopes = list(value)
    scope_set = set(scopes)
    allowed_scopes = _REQUIRED_COURSEWORK_SCOPES | frozenset(
        _GOOGLE_ADDITIONAL_IDENTITY_SCOPES
    )
    if (
        len(scopes) != len(scope_set)
        or not _REQUIRED_COURSEWORK_SCOPES.issubset(scope_set)
        or not scope_set.issubset(allowed_scopes)
    ):
        raise CredentialSecurityError(
            "Authorized token grants do not match approved coursework permissions"
        )
    return [scope for scope in _GOOGLE_GRANT_SCOPE_ORDER if scope in scope_set]


def _register_google_scope_compatibility_hook(
    oauth_session: Any,
    scope_evidence: MutableMapping[str, Any],
) -> None:
    """Handle Google's additional identity scopes without relaxing other grants."""

    register = getattr(oauth_session, "register_compliance_hook", None)
    if not callable(register):
        raise CredentialSecurityError(
            "Installed OAuth library cannot validate Google's granted scopes"
        )

    def normalize_access_token_response(response: Any) -> Any:
        try:
            token_payload = response.json()
        except (AttributeError, ValueError):
            return response
        if not isinstance(token_payload, Mapping):
            return response
        if "scope" not in token_payload:
            scope_evidence["scope_omitted"] = True
            return response
        raw_scopes = token_payload.get("scope")
        if not isinstance(raw_scopes, str) or not raw_scopes.strip():
            raise CredentialSecurityError(
                "Google returned malformed OAuth grant scope evidence"
            )
        actual_scopes = _normalize_granted_scopes(raw_scopes)
        if actual_scopes is None:
            raise CredentialSecurityError(
                "Google returned malformed OAuth grant scope evidence"
            )
        scope_evidence["granted_scopes"] = actual_scopes
        oauth_session.scope = list(actual_scopes)
        return response

    register("access_token_response", normalize_access_token_response)


def _clean_token_for_write(
    credentials: Any,
    *,
    verified_granted_scopes: Optional[Any] = None,
    allow_omitted_equal: bool = False,
    require_refresh_token: bool = False,
) -> Dict[str, Any]:
    try:
        raw = json.loads(credentials.to_json())
    except Exception as exc:
        raise CredentialSecurityError("Google credentials could not be serialized safely") from exc
    has_scopes = getattr(credentials, "has_scopes", None)
    if callable(has_scopes) and not has_scopes(list(COURSEWORK_SCOPES)):
        raise CredentialSecurityError(
            "Authorized token did not grant every coursework scope"
        )
    actual_grants = verified_granted_scopes
    if actual_grants is None:
        actual_grants = getattr(credentials, "granted_scopes", None)
    normalized_grants = _normalize_granted_scopes(
        actual_grants,
        allow_omitted_equal=allow_omitted_equal,
    )
    if normalized_grants is None:
        raise CredentialSecurityError(
            "Google did not provide verifiable grant scope evidence; authorization was not stored"
        )
    clean: Dict[str, Any] = {
        "token": raw.get("token"),
        "refresh_token": raw.get("refresh_token"),
        "token_uri": GOOGLE_TOKEN_URI,
        "client_id": raw.get("client_id"),
        "client_secret": raw.get("client_secret"),
        "scopes": list(COURSEWORK_SCOPES),
        "granted_scopes": normalized_grants,
    }
    if raw.get("expiry"):
        clean["expiry"] = raw["expiry"]
    sanitized = _sanitize_token_mapping(clean)
    if require_refresh_token and not sanitized.get("refresh_token"):
        raise CredentialSecurityError(
            "Google did not return a durable refresh token; existing token storage was not changed"
        )
    return sanitized


def get_google_classroom_coursework_service(
    paths: CourseworkAuthPaths,
    *,
    expected_account: str,
    open_browser: bool = True,
    force_authorize: bool = False,
    replace_token: bool = False,
    require_existing_token: bool = False,
    timeout_seconds: int = 300,
    profile_timeout_seconds: int = 30,
    service_builder: Optional[Any] = None,
    profile_fetcher: Optional[Any] = None,
) -> CourseworkAuthSession:
    """Load/authorize a token, verify its principal, and build one Classroom service."""

    _require_posix_security()
    if require_existing_token and force_authorize:
        raise ValueError(
            "require_existing_token cannot be combined with force_authorize"
        )
    account = normalize_expected_account(expected_account)
    if account != paths.account:
        raise CredentialSecurityError("Expected account does not match resolved auth paths")
    try:
        from google.auth.transport.requests import AuthorizedSession, Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise CredentialSecurityError(
            "Google API dependencies are not installed; install the package dependencies"
        ) from exc
    if "granted_scopes" not in inspect.signature(Credentials).parameters:
        raise CredentialSecurityError(
            "Installed google-auth is too old to verify granted scopes"
        )

    with TokenFileLock(paths.token_path):
        token_exists = _file_exists_without_following(paths.token_path)
        if require_existing_token and not token_exists:
            raise CredentialSecurityError(
                "An existing coursework token is required; automatic authorization is disabled"
            )
        client_config = None
        if paths.credentials_path is not None and _file_exists_without_following(paths.credentials_path):
            client_config = load_sanitized_client_config(paths.credentials_path)

        credentials = None
        verified_granted_scopes = None
        fresh_authorization = False
        should_write_token = False
        if force_authorize:
            if token_exists and not replace_token:
                raise CredentialSecurityError(
                    "A token already exists; explicit replacement confirmation is required"
                )
            if client_config is None:
                raise CredentialSecurityError("Installed OAuth client credentials are required")
        elif token_exists:
            token_info = load_sanitized_token_info(paths.token_path)
            verified_granted_scopes = token_info["granted_scopes"]
            if client_config is not None:
                if token_info["client_id"] != client_config["installed"]["client_id"]:
                    raise CredentialSecurityError("Token and OAuth client IDs do not match")
            try:
                credentials = Credentials.from_authorized_user_info(
                    token_info, scopes=list(COURSEWORK_SCOPES)
                )
            except Exception:
                raise CredentialSecurityError(
                    "Stored token could not be loaded by google-auth"
                ) from None
            if not credentials.valid:
                if credentials.expired and credentials.refresh_token:
                    try:
                        credentials.refresh(Request())
                    except Exception as exc:
                        raise CredentialSecurityError(
                            "Stored token refresh failed; authorize a replacement explicitly"
                        ) from exc
                    refreshed_grants = getattr(credentials, "granted_scopes", None)
                    if refreshed_grants is not None:
                        verified_granted_scopes = refreshed_grants
                    should_write_token = True
                else:
                    raise CredentialSecurityError(
                        "Stored token is invalid; authorize a replacement explicitly"
                    )

        if credentials is None:
            if paths.token_only or client_config is None:
                raise CredentialSecurityError(
                    "Token-only mode cannot perform interactive authorization"
                )
            signature = inspect.signature(InstalledAppFlow.run_local_server)
            if "timeout_seconds" not in signature.parameters:
                raise CredentialSecurityError(
                    "Installed google-auth-oauthlib is too old; upgrade it before authorizing"
                )
            try:
                flow = InstalledAppFlow.from_client_config(
                    client_config,
                    scopes=list(COURSEWORK_SCOPES),
                    autogenerate_code_verifier=True,
                )
                scope_evidence: Dict[str, Any] = {}
                _register_google_scope_compatibility_hook(
                    flow.oauth2session,
                    scope_evidence,
                )
                with _suppress_oauth_secret_logging():
                    credentials = flow.run_local_server(
                        host="127.0.0.1",
                        port=0,
                        open_browser=open_browser,
                        timeout_seconds=timeout_seconds,
                        access_type="offline",
                        prompt="consent",
                    )
                if "granted_scopes" in scope_evidence:
                    verified_granted_scopes = scope_evidence["granted_scopes"]
                elif scope_evidence.get("scope_omitted") is True:
                    verified_granted_scopes = list(COURSEWORK_SCOPES)
                else:
                    raise CredentialSecurityError(
                        "Google OAuth token response scopes were not validated"
                    )
            except CredentialSecurityError:
                raise
            except Warning as exc:
                raise CredentialSecurityError(
                    "Google returned an unexpected OAuth grant scope set; no token was stored"
                ) from exc
            except Exception as exc:
                raise CredentialSecurityError("Google OAuth authorization did not complete") from exc
            should_write_token = True
            fresh_authorization = True

        grant_evidence = verified_granted_scopes
        if grant_evidence is None:
            grant_evidence = getattr(credentials, "granted_scopes", None)
        normalized_grants = _normalize_granted_scopes(
            grant_evidence,
        )
        if normalized_grants is None:
            raise CredentialSecurityError(
                "Google did not provide verifiable grant scope evidence"
            )
        verified_granted_scopes = normalized_grants

        builder = service_builder or build
        api_http = None
        try:
            api_http = _build_one_shot_authorized_http(
                credentials,
                authorized_session_cls=AuthorizedSession,
            )
            service = builder(
                "classroom",
                "v1",
                http=api_http,
                cache_discovery=False,
            )
            if profile_fetcher is not None:
                profile = profile_fetcher(credentials)
            else:
                userinfo_session = AuthorizedSession(credentials)
                try:
                    response = userinfo_session.get(
                        GOOGLE_USERINFO_URI,
                        timeout=profile_timeout_seconds,
                    )
                    if response.status_code != 200:
                        raise RuntimeError("userinfo request was rejected")
                    profile = response.json()
                finally:
                    userinfo_session.close()
        except CredentialSecurityError:
            if api_http is not None:
                api_http.close()
            raise
        except Exception as exc:
            if api_http is not None:
                api_http.close()
            raise CredentialSecurityError("Authenticated Google profile could not be verified") from exc
        if not isinstance(profile, Mapping):
            raise CredentialSecurityError(
                "Authenticated Google profile returned an invalid response"
            )
        returned_email = str(profile.get("email") or "").strip().lower()
        if returned_email != account:
            raise CredentialSecurityError("Authenticated Google account does not match --account")
        email_verified = profile.get("email_verified")
        if email_verified is None:
            email_verified = profile.get("verified_email")
        if email_verified is not True:
            raise CredentialSecurityError(
                "Authenticated Google account email is not verified"
            )
        user_id = str(profile.get("sub") or profile.get("id") or "")
        if not user_id:
            raise CredentialSecurityError("Authenticated Google profile has no stable ID")

        client_id = getattr(credentials, "client_id", None)
        if not isinstance(client_id, str) or not client_id:
            raise CredentialSecurityError("Authorized credentials have no client ID")
        if should_write_token:
            write_private_json_atomic(
                paths.token_path,
                _clean_token_for_write(
                    credentials,
                    verified_granted_scopes=verified_granted_scopes,
                    require_refresh_token=fresh_authorization,
                ),
            )
        metadata = {
            "account_hash": hashlib.sha256(account.encode("utf-8")).hexdigest(),
            "user_id_fingerprint": _fingerprint(user_id),
            "client_id_fingerprint": _fingerprint(client_id),
            "scopes": list(COURSEWORK_SCOPES),
            "granted_scopes_verified": True,
            "verified_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        }
        write_private_json_atomic(paths.metadata_path, metadata)
        return CourseworkAuthSession(
            service=service,
            profile={
                "id": user_id,
                "emailAddress": returned_email,
                "emailVerified": True,
            },
            paths=paths,
            client_id_fingerprint=_fingerprint(client_id),
        )


__all__ = [
    "COURSEWORK_SCOPES",
    "CourseworkAuthPaths",
    "CourseworkAuthSession",
    "CredentialSecurityError",
    "TokenFileLock",
    "UnsupportedPlatformCredentialSecurityError",
    "account_fingerprint",
    "get_google_classroom_auth_status",
    "get_google_classroom_coursework_service",
    "load_sanitized_client_config",
    "load_sanitized_token_info",
    "normalize_expected_account",
    "resolve_coursework_auth_paths",
    "secure_read_json_file",
    "write_private_json_atomic",
]
