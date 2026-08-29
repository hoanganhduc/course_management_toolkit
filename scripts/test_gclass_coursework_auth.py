#!/usr/bin/env python3
"""Offline tests for isolated Google Classroom coursework credentials."""

from __future__ import annotations

import json
import importlib.util
import inspect
import logging
import os
import stat
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from course_hoanganhduc.gclass_coursework_auth import (  # noqa: E402
    COURSEWORK_SCOPES,
    GOOGLE_TOKEN_URI,
    CredentialSecurityError,
    TokenFileLock,
    _register_google_scope_compatibility_hook,
    account_fingerprint,
    get_google_classroom_auth_status,
    load_sanitized_client_config,
    load_sanitized_token_info,
    resolve_coursework_auth_paths,
    secure_read_json_file,
    write_private_json_atomic,
    get_google_classroom_coursework_service,
    _clean_token_for_write,
)


def _has_git_ancestor(path):
    current = Path(path).absolute()
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


def _select_test_temp_parent():
    candidates = [Path(tempfile.gettempdir()), Path.home()]
    if os.name != "nt":
        candidates.append(Path("/var/tmp"))
    seen = set()
    for candidate in candidates:
        normalized = os.path.normcase(str(candidate.absolute()))
        if normalized in seen:
            continue
        seen.add(normalized)
        if (
            candidate.is_dir()
            and os.access(candidate, os.W_OK | os.X_OK)
            and not _has_git_ancestor(candidate)
        ):
            return candidate
    raise RuntimeError("No writable test temporary directory exists outside a Git worktree")


TEST_TEMP_PARENT = _select_test_temp_parent()


def google_dependencies_available():
    try:
        return all(
            importlib.util.find_spec(name) is not None
            for name in (
                "google.oauth2.credentials",
                "google_auth_oauthlib.flow",
                "googleapiclient.discovery",
            )
        )
    except (ImportError, ModuleNotFoundError, ValueError):
        return False


class TestPathResolution(unittest.TestCase):
    def test_linux_default_is_account_scoped(self):
        paths = resolve_coursework_auth_paths(
            "Teacher@Example.edu",
            env={"XDG_CONFIG_HOME": "/safe/config"},
            platform_system="linux",
            home=Path("/home/tester"),
        )
        self.assertEqual(paths.source, "default")
        self.assertEqual(paths.credentials_path, Path("/safe/config/course/google-classroom/credentials.json"))
        self.assertEqual(
            paths.token_path.name,
            f"{account_fingerprint('teacher@example.edu')}.json",
        )

    def test_cli_pair_and_token_only(self):
        pair = resolve_coursework_auth_paths(
            "teacher@example.edu",
            credentials_path="/secrets/client.json",
            token_path="/secrets/token.json",
            env={
                "COURSE_GCLASS_CREDENTIALS": "/ignored/client.json",
                "COURSE_GCLASS_COURSEWORK_TOKEN": "/ignored/token.json",
            },
        )
        self.assertEqual(pair.source, "cli")
        self.assertEqual(pair.credentials_path, Path("/secrets/client.json"))
        token_only = resolve_coursework_auth_paths(
            "teacher@example.edu", token_path="/secrets/token.json", env={}
        )
        self.assertIsNone(token_only.credentials_path)
        with self.assertRaises(ValueError):
            resolve_coursework_auth_paths(
                "teacher@example.edu", credentials_path="/secrets/client.json", env={}
            )
        with self.assertRaises(ValueError):
            resolve_coursework_auth_paths(
                "teacher@example.edu",
                credentials_path="/secrets/same.json",
                token_path="/secrets/same.json",
                env={},
            )
        with self.assertRaises(ValueError):
            resolve_coursework_auth_paths(
                "teacher@example.edu",
                credentials_path="/secrets/token.meta.json",
                token_path="/secrets/token.json",
                env={},
            )
        with self.assertRaises(ValueError):
            resolve_coursework_auth_paths(
                "teacher@example.edu", token_path="/", env={}
            )

    def test_environment_pair_and_invalid_relative_values(self):
        paths = resolve_coursework_auth_paths(
            "teacher@example.edu",
            env={
                "COURSE_GCLASS_CREDENTIALS": "/env/client.json",
                "COURSE_GCLASS_COURSEWORK_TOKEN": "/env/token.json",
            },
        )
        self.assertEqual(paths.source, "environment")
        with self.assertRaises(ValueError):
            resolve_coursework_auth_paths(
                "teacher@example.edu",
                env={"COURSE_GCLASS_COURSEWORK_TOKEN": "relative.json"},
            )
        with self.assertRaises(ValueError):
            resolve_coursework_auth_paths(
                "teacher@example.edu",
                env={"XDG_CONFIG_HOME": "relative"},
                platform_system="linux",
            )

    def test_invalid_account(self):
        for value in ("", "not-an-email", "a b@example.edu", "x\x1b@example.edu"):
            with self.subTest(value=value):
                with self.assertRaises(ValueError):
                    resolve_coursework_auth_paths(value, env={})

    def test_fresh_oauth_scope_omission_and_google_identity_additions_are_accepted(self):
        class FakeCredentials:
            def __init__(self, granted_scopes):
                self.granted_scopes = granted_scopes

            def has_scopes(self, scopes):
                return set(scopes) == set(COURSEWORK_SCOPES)

            def to_json(self):
                return json.dumps(
                    {
                        "token": "access",
                        "refresh_token": "refresh",
                        "client_id": "abc.apps.googleusercontent.com",
                        "client_secret": "secret",
                        "scopes": list(COURSEWORK_SCOPES),
                        "expiry": "2099-01-01T00:00:00Z",
                    }
                )

        clean = _clean_token_for_write(
            FakeCredentials([]), allow_omitted_equal=True
        )
        self.assertEqual(clean["granted_scopes"], list(COURSEWORK_SCOPES))
        google_identity_additions = list(COURSEWORK_SCOPES) + ["email", "openid"]
        clean = _clean_token_for_write(
            FakeCredentials(google_identity_additions),
            allow_omitted_equal=True,
        )
        self.assertEqual(
            set(clean["granted_scopes"]), set(google_identity_additions)
        )
        oidc_identity_scopes = list(COURSEWORK_SCOPES[:2]) + ["email", "openid"]
        with self.assertRaises(CredentialSecurityError):
            _clean_token_for_write(
                FakeCredentials(oidc_identity_scopes),
                allow_omitted_equal=True,
            )
        no_refresh = FakeCredentials(list(COURSEWORK_SCOPES))
        original_to_json = no_refresh.to_json
        no_refresh.to_json = lambda: json.dumps(
            dict(json.loads(original_to_json()), refresh_token=None)
        )
        with self.assertRaisesRegex(CredentialSecurityError, "durable refresh token"):
            _clean_token_for_write(
                no_refresh,
                allow_omitted_equal=True,
                require_refresh_token=True,
            )
        with self.assertRaises(CredentialSecurityError):
            _clean_token_for_write(FakeCredentials([]))
        with self.assertRaises(CredentialSecurityError):
            _clean_token_for_write(
                FakeCredentials(
                    list(COURSEWORK_SCOPES)
                    + ["https://www.googleapis.com/auth/drive"]
                ),
                allow_omitted_equal=True,
            )
        with self.assertRaises(CredentialSecurityError):
            _clean_token_for_write(
                FakeCredentials(
                    list(COURSEWORK_SCOPES[:2]) + ["openid"]
                ),
                allow_omitted_equal=True,
            )

    def test_custom_secret_paths_inside_git_worktree_are_rejected(self):
        with tempfile.TemporaryDirectory(dir=str(TEST_TEMP_PARENT)) as temporary:
            root = Path(temporary)
            (root / ".git").mkdir()
            with self.assertRaises(ValueError):
                resolve_coursework_auth_paths(
                    "teacher@example.edu",
                    credentials_path=str(root / "credentials.json"),
                    token_path=str(root / "token.json"),
                    env={},
                )


@unittest.skipIf(os.name == "nt", "POSIX permission tests")
class TestSecureFiles(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(dir=str(TEST_TEMP_PARENT))
        self.root = Path(self.temp.name)
        os.chmod(self.root, 0o700)

    def tearDown(self):
        self.temp.cleanup()

    def write_json(self, name, value, mode=0o600):
        path = self.root / name
        path.write_text(json.dumps(value), encoding="utf-8")
        os.chmod(path, mode)
        return path

    def test_secure_read_rejects_public_file_and_symlink(self):
        path = self.write_json("token.json", {"a": 1}, mode=0o644)
        with self.assertRaises(CredentialSecurityError):
            secure_read_json_file(path)
        os.chmod(path, 0o600)
        link = self.root / "link.json"
        link.symlink_to(path)
        with self.assertRaises(CredentialSecurityError):
            secure_read_json_file(link)

    def test_secure_read_rejects_duplicate_keys(self):
        path = self.root / "duplicate.json"
        path.write_text('{"a":1,"a":2}', encoding="utf-8")
        os.chmod(path, 0o600)
        with self.assertRaises(CredentialSecurityError):
            secure_read_json_file(path)

    def test_secure_read_rejects_fifo_without_blocking(self):
        fifo = self.root / "credential.fifo"
        os.mkfifo(fifo)
        with self.assertRaises(CredentialSecurityError):
            secure_read_json_file(fifo)

    def test_atomic_write_and_lock(self):
        path = self.root / "tokens" / "token.json"
        write_private_json_atomic(path, {"token": "secret"})
        self.assertEqual(secure_read_json_file(path)["token"], "secret")
        self.assertEqual(stat.S_IMODE(path.stat().st_mode), 0o600)
        self.assertEqual(stat.S_IMODE(path.parent.stat().st_mode), 0o700)
        with TokenFileLock(path):
            with self.assertRaises(CredentialSecurityError):
                with TokenFileLock(path):
                    pass

    def test_lock_rejects_public_or_symlink_lock_file(self):
        token = self.root / "tokens" / "token.json"
        token.parent.mkdir(mode=0o700)
        lock = token.parent / ".token.json.lock"
        lock.write_text("", encoding="utf-8")
        os.chmod(lock, 0o644)
        with self.assertRaises(CredentialSecurityError):
            with TokenFileLock(token):
                pass
        lock.unlink()
        target = self.root / "lock-target"
        target.write_text("", encoding="utf-8")
        os.chmod(target, 0o600)
        lock.symlink_to(target)
        with self.assertRaises(CredentialSecurityError):
            with TokenFileLock(token):
                pass

    def test_client_config_ignores_hostile_endpoints(self):
        path = self.write_json(
            "credentials.json",
            {
                "installed": {
                    "client_id": "abc.apps.googleusercontent.com",
                    "client_secret": "secret",
                    "auth_uri": "https://attacker.invalid/auth",
                    "token_uri": "https://attacker.invalid/token",
                }
            },
        )
        config = load_sanitized_client_config(path)
        installed = config["installed"]
        self.assertEqual(installed["auth_uri"], "https://accounts.google.com/o/oauth2/v2/auth")
        self.assertEqual(installed["token_uri"], "https://oauth2.googleapis.com/token")
        self.assertNotIn("attacker.invalid", json.dumps(config))
        service = self.write_json("service.json", {"type": "service_account"})
        with self.assertRaises(CredentialSecurityError):
            load_sanitized_client_config(service)

    def test_token_info_uses_canonical_endpoint_and_exact_scopes(self):
        path = self.write_json(
            "token.json",
            {
                "token": "access",
                "refresh_token": "refresh",
                "token_uri": "https://attacker.invalid/token",
                "client_id": "abc.apps.googleusercontent.com",
                "client_secret": "secret",
                "scopes": list(COURSEWORK_SCOPES),
                "granted_scopes": list(COURSEWORK_SCOPES),
                "expiry": "2099-01-01T00:00:00Z",
            },
        )
        info = load_sanitized_token_info(path)
        self.assertEqual(info["token_uri"], "https://oauth2.googleapis.com/token")
        offset = dict(
            secure_read_json_file(path),
            expiry="2027-01-01T07:00:00+07:00Z",
        )
        offset_path = self.write_json("offset.json", offset)
        self.assertEqual(
            load_sanitized_token_info(offset_path)["expiry"],
            "2027-01-01T00:00:00Z",
        )
        for label, expiry in (
            ("underflow", "0001-01-01T00:00:00+14:00"),
            ("overflow", "9999-12-31T23:59:59-14:00"),
        ):
            with self.subTest(label=label):
                extreme = dict(secure_read_json_file(path), expiry=expiry)
                extreme_path = self.write_json(f"{label}-expiry.json", extreme)
                with self.assertRaises(CredentialSecurityError):
                    load_sanitized_token_info(extreme_path)
        wrong = dict(secure_read_json_file(path), scopes=[COURSEWORK_SCOPES[0]])
        wrong_path = self.write_json("wrong.json", wrong)
        with self.assertRaises(CredentialSecurityError):
            load_sanitized_token_info(wrong_path)
        duplicated = dict(
            secure_read_json_file(path),
            scopes=list(COURSEWORK_SCOPES) + [COURSEWORK_SCOPES[0]],
        )
        duplicate_path = self.write_json("duplicate-scopes.json", duplicated)
        with self.assertRaises(CredentialSecurityError):
            load_sanitized_token_info(duplicate_path)
        missing_grants = dict(secure_read_json_file(path))
        missing_grants.pop("granted_scopes")
        missing_path = self.write_json("missing-grants.json", missing_grants)
        with self.assertRaises(CredentialSecurityError):
            load_sanitized_token_info(missing_path)
        broader_grants = dict(
            secure_read_json_file(path),
            granted_scopes=list(COURSEWORK_SCOPES)
            + ["https://www.googleapis.com/auth/drive"],
        )
        broader_path = self.write_json("broader-grants.json", broader_grants)
        with self.assertRaises(CredentialSecurityError):
            load_sanitized_token_info(broader_path)
        identity_additions = dict(
            secure_read_json_file(path),
            granted_scopes=list(COURSEWORK_SCOPES) + ["email", "openid"],
        )
        additions_path = self.write_json(
            "google-identity-additions.json", identity_additions
        )
        self.assertEqual(
            set(load_sanitized_token_info(additions_path)["granted_scopes"]),
            set(identity_additions["granted_scopes"]),
        )

    def test_auth_status_is_offline_and_redacted(self):
        client = self.write_json(
            "credentials.json",
            {"installed": {"client_id": "abc.apps.googleusercontent.com", "client_secret": "secret"}},
        )
        token = self.write_json(
            "token.json",
            {
                "token": "access-secret",
                "refresh_token": "refresh-secret",
                "client_id": "abc.apps.googleusercontent.com",
                "client_secret": "secret",
                "scopes": list(COURSEWORK_SCOPES),
                "granted_scopes": list(COURSEWORK_SCOPES),
                "expiry": "2027-01-01T00:00:00Z",
            },
        )
        paths = resolve_coursework_auth_paths(
            "teacher@example.edu",
            credentials_path=str(client),
            token_path=str(token),
            env={},
        )
        status = get_google_classroom_auth_status(paths, show_paths=False)
        rendered = json.dumps(status)
        self.assertTrue(status["token_exists"])
        self.assertEqual(status["principal_status"], "offline/unverified")
        self.assertTrue(status["client_ids_match"])
        self.assertNotIn("access-secret", rendered)
        self.assertNotIn(str(token), rendered)

    def test_auth_status_marks_expired_access_only_token_unusable(self):
        token = self.write_json(
            "expired.json",
            {
                "token": "expired-access",
                "refresh_token": None,
                "client_id": "abc.apps.googleusercontent.com",
                "client_secret": "secret",
                "scopes": list(COURSEWORK_SCOPES),
                "granted_scopes": list(COURSEWORK_SCOPES),
                "expiry": "2020-01-01T00:00:00Z",
            },
        )
        paths = resolve_coursework_auth_paths(
            "teacher@example.edu", token_path=str(token), env={}
        )
        status = get_google_classroom_auth_status(paths)
        self.assertTrue(status["token_safe"])
        self.assertTrue(status["token_expired"])
        self.assertFalse(status["token_refreshable"])
        self.assertFalse(status["token_usable"])

    def test_auth_status_matches_google_loader_for_unusual_token_shapes(self):
        common = {
            "client_id": "abc.apps.googleusercontent.com",
            "client_secret": "secret",
            "scopes": list(COURSEWORK_SCOPES),
            "granted_scopes": list(COURSEWORK_SCOPES),
        }
        cases = {
            "access-only-without-expiry": (
                {**common, "token": "access", "refresh_token": None},
                False,
            ),
            "refresh-only-with-future-expiry": (
                {
                    **common,
                    "token": None,
                    "refresh_token": "refresh",
                    "expiry": "2099-01-01T00:00:00Z",
                },
                False,
            ),
            "refresh-only-without-expiry": (
                {**common, "token": None, "refresh_token": "refresh"},
                True,
            ),
        }
        for label, (payload, expected_usable) in cases.items():
            with self.subTest(label=label):
                token = self.write_json(f"{label}.json", payload)
                paths = resolve_coursework_auth_paths(
                    "teacher@example.edu", token_path=str(token), env={}
                )
                status = get_google_classroom_auth_status(paths)
                self.assertTrue(status["token_safe"])
                self.assertEqual(status["token_usable"], expected_usable)


@unittest.skipUnless(
    google_dependencies_available(),
    "Google API dependencies are not installed",
)
@unittest.skipIf(os.name == "nt", "POSIX credential integration test")
class TestGoogleDependencyIntegration(unittest.TestCase):
    def test_require_existing_token_never_falls_back_to_authorization(self):
        with tempfile.TemporaryDirectory(dir=str(TEST_TEMP_PARENT)) as temporary:
            root = Path(temporary)
            os.chmod(root, 0o700)
            paths = resolve_coursework_auth_paths(
                "teacher@example.edu",
                credentials_path=str(root / "credentials.json"),
                token_path=str(root / "missing-token.json"),
                env={},
            )
            builder_calls = []
            with self.assertRaisesRegex(
                CredentialSecurityError, "existing coursework token"
            ):
                get_google_classroom_coursework_service(
                    paths,
                    expected_account="teacher@example.edu",
                    open_browser=False,
                    require_existing_token=True,
                    service_builder=lambda *args, **kwargs: builder_calls.append(
                        (args, kwargs)
                    ),
                )
            self.assertEqual(builder_calls, [])
            self.assertFalse(paths.token_path.exists())

            with self.assertRaisesRegex(ValueError, "cannot be combined"):
                get_google_classroom_coursework_service(
                    paths,
                    expected_account="teacher@example.edu",
                    open_browser=False,
                    force_authorize=True,
                    require_existing_token=True,
                )

    def test_google_additional_identity_scopes_are_narrowly_accepted(self):
        from google_auth_oauthlib.flow import InstalledAppFlow
        from requests import Request, Response
        from requests_oauthlib import OAuth2Session

        def token_response(scopes):
            response = Response()
            response.status_code = 200
            response.headers["Content-Type"] = "application/json"
            response._content = json.dumps(
                {
                    "access_token": "synthetic-access-token",
                    "refresh_token": "synthetic-refresh-token",
                    "expires_in": 3600,
                    "token_type": "Bearer",
                    "scope": " ".join(scopes),
                }
            ).encode("utf-8")
            response.request = Request("POST", GOOGLE_TOKEN_URI).prepare()
            return response

        actual_scopes = list(COURSEWORK_SCOPES) + ["email", "openid"]
        evidence = {}
        session = OAuth2Session(
            client_id="abc.apps.googleusercontent.com",
            scope=list(COURSEWORK_SCOPES),
            redirect_uri="http://127.0.0.1:12345/",
        )
        _register_google_scope_compatibility_hook(session, evidence)
        response = token_response(actual_scopes)
        response_content = response.content
        session.request = lambda *args, **kwargs: response
        token = session.fetch_token(
            GOOGLE_TOKEN_URI,
            code="synthetic-code",
            client_secret="synthetic-secret",
            include_client_id=True,
        )
        self.assertEqual(set(token["scope"]), set(actual_scopes))
        self.assertEqual(set(session.scope), set(actual_scopes))
        self.assertEqual(set(evidence["granted_scopes"]), set(actual_scopes))
        self.assertEqual(response.content, response_content)

        flow = InstalledAppFlow.from_client_config(
            {
                "installed": {
                    "client_id": "abc.apps.googleusercontent.com",
                    "client_secret": "synthetic-secret",
                    "auth_uri": "https://accounts.google.com/o/oauth2/v2/auth",
                    "token_uri": GOOGLE_TOKEN_URI,
                }
            },
            scopes=list(COURSEWORK_SCOPES),
            autogenerate_code_verifier=True,
        )
        flow.redirect_uri = "http://127.0.0.1:12345/"
        flow_evidence = {}
        _register_google_scope_compatibility_hook(
            flow.oauth2session,
            flow_evidence,
        )
        flow_response = token_response(actual_scopes)
        flow_content = flow_response.content
        flow.oauth2session.request = lambda *args, **kwargs: flow_response
        flow.fetch_token(code="synthetic-code")
        google_credentials = flow.credentials
        self.assertEqual(set(google_credentials.scopes), set(actual_scopes))
        self.assertEqual(set(google_credentials.granted_scopes), set(actual_scopes))
        clean = _clean_token_for_write(
            google_credentials,
            verified_granted_scopes=flow_evidence["granted_scopes"],
            require_refresh_token=True,
        )
        self.assertEqual(clean["scopes"], list(COURSEWORK_SCOPES))
        self.assertEqual(set(clean["granted_scopes"]), set(actual_scopes))
        self.assertEqual(flow_response.content, flow_content)

        broader_scopes = actual_scopes + [
            "https://www.googleapis.com/auth/drive"
        ]
        evidence = {}
        session = OAuth2Session(
            client_id="abc.apps.googleusercontent.com",
            scope=list(COURSEWORK_SCOPES),
            redirect_uri="http://127.0.0.1:12345/",
        )
        _register_google_scope_compatibility_hook(session, evidence)
        session.request = lambda *args, **kwargs: token_response(broader_scopes)
        with self.assertRaisesRegex(
            CredentialSecurityError, "do not match approved coursework permissions"
        ):
            session.fetch_token(
                GOOGLE_TOKEN_URI,
                code="synthetic-code",
                client_secret="synthetic-secret",
                include_client_id=True,
            )

    def test_google_scope_hook_distinguishes_omitted_from_malformed(self):
        from requests import Response
        from requests_oauthlib import OAuth2Session

        def response_with(payload):
            response = Response()
            response.status_code = 200
            response.headers["Content-Type"] = "application/json"
            response._content = json.dumps(payload).encode("utf-8")
            return response

        session = OAuth2Session(
            client_id="abc.apps.googleusercontent.com",
            scope=list(COURSEWORK_SCOPES),
        )
        evidence = {}
        _register_google_scope_compatibility_hook(session, evidence)
        hook = next(iter(session.compliance_hook["access_token_response"]))
        omitted = response_with({"access_token": "synthetic-access-token"})
        self.assertIs(hook(omitted), omitted)
        self.assertEqual(evidence, {"scope_omitted": True})

        for label, value in (
            ("null", None),
            ("empty", ""),
            ("list", list(COURSEWORK_SCOPES)),
            ("number", 7),
        ):
            with self.subTest(label=label):
                session = OAuth2Session(
                    client_id="abc.apps.googleusercontent.com",
                    scope=list(COURSEWORK_SCOPES),
                )
                evidence = {}
                _register_google_scope_compatibility_hook(session, evidence)
                hook = next(iter(session.compliance_hook["access_token_response"]))
                with self.assertRaisesRegex(
                    CredentialSecurityError, "malformed OAuth grant scope evidence"
                ):
                    hook(response_with({"scope": value}))
                self.assertEqual(evidence, {})

    def test_current_google_libraries_load_token_verify_profile_and_write_metadata(self):
        from datetime import datetime, timezone

        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow

        self.assertIn(
            "timeout_seconds",
            inspect.signature(InstalledAppFlow.run_local_server).parameters,
        )

        with tempfile.TemporaryDirectory(dir=str(TEST_TEMP_PARENT)) as temporary:
            root = Path(temporary)
            os.chmod(root, 0o700)
            credentials_path = root / "credentials.json"
            token_path = root / "token.json"
            client_id = "abc.apps.googleusercontent.com"
            client_secret = "client-secret"
            write_private_json_atomic(
                credentials_path,
                {
                    "installed": {
                        "client_id": client_id,
                        "client_secret": client_secret,
                        "auth_uri": "https://attacker.invalid/auth",
                        "token_uri": "https://attacker.invalid/token",
                    }
                },
            )
            sanitized_config = load_sanitized_client_config(credentials_path)
            flow = InstalledAppFlow.from_client_config(
                sanitized_config,
                scopes=list(COURSEWORK_SCOPES),
                autogenerate_code_verifier=True,
            )
            self.assertTrue(flow.autogenerate_code_verifier)
            google_credentials = Credentials(
                token="access-token",
                refresh_token="refresh-token",
                token_uri="https://oauth2.googleapis.com/token",
                client_id=client_id,
                client_secret=client_secret,
                scopes=list(COURSEWORK_SCOPES),
                granted_scopes=list(COURSEWORK_SCOPES),
                expiry=datetime(2099, 1, 1, tzinfo=timezone.utc),
            )
            token_payload = json.loads(google_credentials.to_json())
            token_payload["granted_scopes"] = list(COURSEWORK_SCOPES)
            write_private_json_atomic(
                token_path, token_payload
            )
            paths = resolve_coursework_auth_paths(
                "teacher@example.edu",
                credentials_path=str(credentials_path),
                token_path=str(token_path),
                env={},
            )
            call_log = []
            authorized_transports = []

            class Service:
                pass

            def service_builder(*args, **kwargs):
                call_log.append(("build", args, kwargs))
                self.assertNotIn("credentials", kwargs)
                authorized_http = kwargs["http"]
                self.assertEqual(
                    authorized_http._session._max_refresh_attempts, 0
                )
                self.assertEqual(
                    tuple(authorized_http._session._refresh_status_codes), ()
                )
                self.assertEqual(
                    authorized_http._session.adapters[
                        "https://"
                    ].max_retries.total,
                    0,
                )
                authorized_transports.append(authorized_http)
                return Service()

            def profile_fetcher(_credentials):
                call_log.append(("userinfo",))
                return {
                    "id": "google-user-id",
                    "email": "teacher@example.edu",
                    "verified_email": True,
                }

            session = get_google_classroom_coursework_service(
                paths,
                expected_account="teacher@example.edu",
                service_builder=service_builder,
                profile_fetcher=profile_fetcher,
            )
            self.assertEqual(
                session.profile["emailAddress"], "teacher@example.edu"
            )
            self.assertEqual(len(authorized_transports), 1)

            import socket
            import threading
            import time
            import requests

            listener = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            listener.bind(("127.0.0.1", 0))
            listener.listen(4)
            listener.settimeout(0.1)
            accepted_connections = []

            def close_stale_connections():
                deadline = time.monotonic() + 0.7
                while time.monotonic() < deadline:
                    try:
                        connection, _address = listener.accept()
                    except socket.timeout:
                        continue
                    accepted_connections.append(connection)
                    try:
                        connection.recv(4096)
                    finally:
                        connection.close()
                listener.close()

            stale_server = threading.Thread(
                target=close_stale_connections,
                daemon=True,
            )
            stale_server.start()
            stale_url = "http://127.0.0.1:{}/mutation".format(
                listener.getsockname()[1]
            )
            with self.assertRaises(requests.exceptions.RequestException):
                authorized_transports[0]._session.request(
                    "POST",
                    stale_url,
                    data=b"{}",
                    allow_redirects=False,
                    timeout=(1, 1),
                )
            stale_server.join(timeout=2)
            self.assertFalse(stale_server.is_alive())
            self.assertEqual(len(accepted_connections), 1)

            from requests import Response

            wire_calls = []

            def one_shot_send(request, **kwargs):
                wire_calls.append((request, kwargs))
                response = Response()
                response.status_code = 401
                response.reason = "Unauthorized"
                response._content = b"unauthorized"
                response.headers["content-type"] = "application/json"
                response.request = request
                return response

            authorized_transports[0]._session.send = one_shot_send
            response, _content = authorized_transports[0].request(
                "https://classroom.googleapis.com/v1/courses/1/courseWork",
                method="POST",
                body=b"{}",
                headers={"content-type": "application/json"},
            )
            self.assertEqual(response.status, 401)
            self.assertEqual(len(wire_calls), 1)
            self.assertFalse(wire_calls[0][1]["allow_redirects"])
            self.assertTrue(paths.metadata_path.exists())
            self.assertEqual(call_log[-1], ("userinfo",))

            token_before = token_path.read_bytes()
            metadata_before = paths.metadata_path.read_bytes()
            with self.assertRaises(CredentialSecurityError):
                get_google_classroom_coursework_service(
                    paths,
                    expected_account="teacher@example.edu",
                    service_builder=service_builder,
                    profile_fetcher=lambda _credentials: {
                        "id": "other-user",
                        "email": "other@example.edu",
                        "verified_email": True,
                    },
                )
            self.assertEqual(token_path.read_bytes(), token_before)
            self.assertEqual(paths.metadata_path.read_bytes(), metadata_before)

    def test_reauthorization_without_refresh_token_preserves_existing_token(self):
        from google_auth_oauthlib.flow import InstalledAppFlow

        with tempfile.TemporaryDirectory(dir=str(TEST_TEMP_PARENT)) as temporary:
            root = Path(temporary)
            os.chmod(root, 0o700)
            credentials_path = root / "credentials.json"
            token_path = root / "token.json"
            client_id = "abc.apps.googleusercontent.com"
            client_secret = "client-secret"
            write_private_json_atomic(
                credentials_path,
                {
                    "installed": {
                        "client_id": client_id,
                        "client_secret": client_secret,
                    }
                },
            )
            write_private_json_atomic(
                token_path,
                {
                    "token": "existing-access",
                    "refresh_token": "existing-refresh",
                    "client_id": client_id,
                    "client_secret": client_secret,
                    "scopes": list(COURSEWORK_SCOPES),
                    "granted_scopes": list(COURSEWORK_SCOPES),
                    "expiry": "2099-01-01T00:00:00Z",
                },
            )
            paths = resolve_coursework_auth_paths(
                "teacher@example.edu",
                credentials_path=str(credentials_path),
                token_path=str(token_path),
                env={},
            )
            token_before = token_path.read_bytes()
            flow_calls = []
            logger_names = (
                "google_auth_oauthlib.flow",
                "requests_oauthlib.oauth2_session",
            )
            logger_before = {
                name: logging.getLogger(name).disabled for name in logger_names
            }
            logger_states = []

            class AccessOnlyCredentials:
                granted_scopes = list(COURSEWORK_SCOPES)

                def __init__(self):
                    self.client_id = client_id

                def has_scopes(self, scopes):
                    return set(scopes) == set(COURSEWORK_SCOPES)

                def to_json(self):
                    return json.dumps(
                        {
                            "token": "replacement-access",
                            "refresh_token": None,
                            "client_id": client_id,
                            "client_secret": client_secret,
                            "scopes": list(COURSEWORK_SCOPES),
                            "expiry": "2099-01-01T00:00:00Z",
                        }
                    )

            class AccessOnlyFlow:
                class OAuthSession:
                    def __init__(self):
                        self.hooks = []
                        self.scope = list(COURSEWORK_SCOPES)

                    def register_compliance_hook(self, hook_type, hook):
                        self.hooks.append((hook_type, hook))

                def __init__(self, invoke_hook):
                    self.oauth2session = self.OAuthSession()
                    self.invoke_hook = invoke_hook

                def run_local_server(self, **kwargs):
                    flow_calls.append(kwargs)
                    logger_states.append(
                        {
                            name: logging.getLogger(name).disabled
                            for name in logger_names
                        }
                    )
                    if self.invoke_hook:
                        class ScopeOmittedResponse:
                            @staticmethod
                            def json():
                                return {"access_token": "synthetic-access-token"}

                        for _hook_type, hook in self.oauth2session.hooks:
                            hook(ScopeOmittedResponse())
                    return AccessOnlyCredentials()

            no_hook_flow = AccessOnlyFlow(invoke_hook=False)
            with mock.patch.object(
                InstalledAppFlow,
                "from_client_config",
                return_value=no_hook_flow,
            ):
                with self.assertRaisesRegex(
                    CredentialSecurityError, "scopes were not validated"
                ):
                    get_google_classroom_coursework_service(
                        paths,
                        expected_account="teacher@example.edu",
                        force_authorize=True,
                        replace_token=True,
                        service_builder=lambda *args, **kwargs: object(),
                        profile_fetcher=lambda _credentials: {
                            "id": "google-user-id",
                            "email": "teacher@example.edu",
                            "verified_email": True,
                        },
                    )
            self.assertEqual(token_path.read_bytes(), token_before)

            access_only_flow = AccessOnlyFlow(invoke_hook=True)
            with mock.patch.object(
                InstalledAppFlow,
                "from_client_config",
                return_value=access_only_flow,
            ):
                with self.assertRaisesRegex(
                    CredentialSecurityError, "durable refresh token"
                ):
                    get_google_classroom_coursework_service(
                        paths,
                        expected_account="teacher@example.edu",
                        force_authorize=True,
                        replace_token=True,
                        service_builder=lambda *args, **kwargs: object(),
                        profile_fetcher=lambda _credentials: {
                            "id": "google-user-id",
                            "email": "teacher@example.edu",
                            "verified_email": True,
                        },
                    )

            self.assertEqual(token_path.read_bytes(), token_before)
            self.assertFalse(paths.metadata_path.exists())
            self.assertEqual(len(flow_calls), 2)
            self.assertEqual(flow_calls[-1]["host"], "127.0.0.1")
            self.assertEqual(flow_calls[-1]["port"], 0)
            self.assertEqual(flow_calls[-1]["access_type"], "offline")
            self.assertEqual(flow_calls[-1]["prompt"], "consent")
            self.assertEqual(
                [hook_type for hook_type, _hook in access_only_flow.oauth2session.hooks],
                ["access_token_response"],
            )
            self.assertTrue(
                all(
                    state[name]
                    for state in logger_states
                    for name in logger_names
                )
            )
            self.assertEqual(
                {
                    name: logging.getLogger(name).disabled
                    for name in logger_names
                },
                logger_before,
            )


if __name__ == "__main__":
    unittest.main()
