#!/usr/bin/env python3
"""Offline tests for the legacy Google Classroom credential flow.

The point of this file is the headless path. ``run_local_server`` blocks inside
``handle_request()`` until a browser reaches its loopback port, so it can never
also ask the operator for the redirect; a machine with no browser therefore
finishes consent by pasting the failed redirect back into the same terminal,
with no listener and no second process involved.

Nothing touches the network: the OAuth flow, the browser lookup and the
Classroom service are all replaced, and the third-party stack is stubbed at
import time.
"""

from __future__ import annotations

import os
import pickle
import sys
import tempfile
import types
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
from typing import Any, Dict, List, Optional
from unittest import mock

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)


def _stub_heavy_deps() -> None:
    """Stub the third-party stack gclass_auth imports at module scope."""

    def ensure(name: str, attrs: Optional[Dict[str, Any]] = None) -> None:
        if name in sys.modules:
            return
        module = types.ModuleType(name)
        for key, value in (attrs or {}).items():
            setattr(module, key, value)
        sys.modules[name] = module

    ensure("googleapiclient")
    ensure("googleapiclient.discovery", {"build": lambda *a, **k: None})
    ensure("googleapiclient.errors", {"HttpError": Exception})
    ensure("google")
    ensure("google.oauth2")
    ensure("google.oauth2.credentials", {"Credentials": object})
    ensure("google.auth")
    ensure("google.auth.transport")
    ensure("google.auth.transport.requests", {"Request": object})
    ensure("google_auth_oauthlib")
    ensure("google_auth_oauthlib.flow", {"InstalledAppFlow": object})


_stub_heavy_deps()

from course_hoanganhduc import gclass_auth  # noqa: E402

AUTH_URL = "https://accounts.google.com/o/oauth2/auth?client_id=stub"
GOOD_PASTE = f"{gclass_auth.PASTE_REDIRECT_URI}?state=st-1&code=cd-1"


class FakeCredentials:
    """A credential object with the attributes the module reads."""

    def __init__(self, scopes: Optional[List[str]] = None, valid: bool = True) -> None:
        self.scopes = list(scopes if scopes is not None else gclass_auth.SCOPES)
        self.valid = valid
        self.expired = False
        self.refresh_token = None


class FakeFlow:
    """Records both authorization strategies and hands back a credential."""

    def __init__(self) -> None:
        self.server_calls: List[Dict[str, Any]] = []
        self.auth_url_calls: List[Dict[str, Any]] = []
        self.token_calls: List[Dict[str, Any]] = []
        self.redirect_uri: Optional[str] = None
        self.fetch_error: Optional[Exception] = None
        self.credentials = FakeCredentials()

    def run_local_server(self, **kwargs: Any) -> FakeCredentials:
        self.server_calls.append(dict(kwargs))
        return self.credentials

    def authorization_url(self, **kwargs: Any):
        self.auth_url_calls.append(dict(kwargs))
        return AUTH_URL, "st-1"

    def fetch_token(self, **kwargs: Any) -> None:
        self.token_calls.append(dict(kwargs))
        if self.fetch_error is not None:
            raise self.fetch_error


class _AuthCase(unittest.TestCase):
    """A temporary credential/token pair plus a stand-in OAuth flow."""

    def setUp(self) -> None:
        workspace = tempfile.TemporaryDirectory()
        self.addCleanup(workspace.cleanup)
        self.directory = Path(workspace.name)
        self.credentials_path = self.directory / "credentials.json"
        self.credentials_path.write_text("{}", encoding="utf-8")
        self.token_path = self.directory / "token.pickle"
        self.flow = FakeFlow()

    def authorize(
        self,
        *,
        browser: bool,
        open_browser: Optional[bool] = None,
        pastes: Optional[List[str]] = None,
    ) -> str:
        """Run the first-authorization branch and return what it printed.

        ``pastes`` feeds the hidden prompt one answer at a time; running out of
        them is exactly what an operator pressing Ctrl-D looks like.
        """
        buffer = StringIO()
        answers = iter(pastes or [])
        flow_factory = types.SimpleNamespace(
            from_client_secrets_file=lambda *a, **k: self.flow
        )
        with mock.patch.object(gclass_auth, "InstalledAppFlow", flow_factory), \
                mock.patch.object(gclass_auth, "_browser_available", lambda: browser), \
                mock.patch.object(
                    gclass_auth.getpass, "getpass", lambda _prompt="": next(answers)
                ), \
                redirect_stdout(buffer):
            gclass_auth._get_google_classroom_credentials(
                str(self.credentials_path),
                str(self.token_path),
                open_browser=open_browser,
            )
        return buffer.getvalue()


class TestPasteFlow(_AuthCase):
    def test_01_headless_authorization_never_starts_a_local_server(self) -> None:
        """No listener at all, so there is nothing for a second terminal to feed.

        The old shape bound a port and blocked on it; the redirect URI is now set
        by hand, which is what lets this same process go on to read the paste.
        """
        self.authorize(browser=False, pastes=[GOOD_PASTE])
        self.assertEqual(self.flow.server_calls, [])
        self.assertEqual(len(self.flow.auth_url_calls), 1)
        self.assertEqual(self.flow.redirect_uri, gclass_auth.PASTE_REDIRECT_URI)

    def test_02_the_operator_is_told_to_paste_here_not_elsewhere(self) -> None:
        output = self.authorize(browser=False, pastes=[GOOD_PASTE])
        self.assertIn(AUTH_URL, output)
        self.assertIn("paste it here", output)
        self.assertNotIn("complete-loopback", output)

    def test_03_the_paste_is_exchanged_over_https(self) -> None:
        """oauthlib refuses to read an OAuth response over http.

        Only the scheme is rewritten; the query has to survive untouched, since
        that is where the code and state actually live.
        """
        self.authorize(browser=False, pastes=[GOOD_PASTE])
        self.assertEqual(len(self.flow.token_calls), 1)
        sent = self.flow.token_calls[0]["authorization_response"]
        self.assertEqual(sent, "https://127.0.0.1:8910/?state=st-1&code=cd-1")

    def test_04_the_advertised_redirect_survives_its_own_validator(self) -> None:
        """The drift guard.

        The URI printed to the operator and the URI the validator accepts are two
        constants that must agree; this fails if either one moves.
        """
        response, reason = gclass_auth._authorization_response(GOOD_PASTE)
        self.assertIsNone(reason)
        self.assertTrue(response.endswith("?state=st-1&code=cd-1"))

    def test_05_the_token_lands_in_the_token_file(self) -> None:
        self.authorize(browser=False, pastes=[GOOD_PASTE])
        with open(self.token_path, "rb") as handle:
            stored = pickle.load(handle)
        self.assertEqual(set(stored.scopes), set(gclass_auth.SCOPES))

    def test_06_a_bad_paste_asks_again_instead_of_giving_up(self) -> None:
        """The authorization code expires in minutes, so a typo must not end it."""
        output = self.authorize(
            browser=False, pastes=["that is not a url", GOOD_PASTE]
        )
        self.assertEqual(len(self.flow.token_calls), 1)
        self.assertIn("Try again", output)

    def test_07_a_denied_consent_is_named_and_never_becomes_a_token(self) -> None:
        denied = f"{gclass_auth.PASTE_REDIRECT_URI}?error=access_denied&state=st-1"
        with self.assertRaises(RuntimeError):
            self.authorize(browser=False, pastes=[denied, denied, denied])
        self.assertEqual(self.flow.token_calls, [])
        self.assertFalse(self.token_path.exists())

    def test_08_nothing_pasted_leaves_no_token(self) -> None:
        with self.assertRaises(RuntimeError):
            self.authorize(browser=False, pastes=[])
        self.assertEqual(self.flow.token_calls, [])
        self.assertFalse(self.token_path.exists())

    def test_09_a_rejected_exchange_stops_instead_of_re_prompting(self) -> None:
        """Google refusing the code is not something a second paste can fix.

        The code is single-use and short-lived, so the way back is a fresh
        authorization URL, not another prompt against the same dead code.
        """
        self.flow.fetch_error = ValueError("invalid_grant")
        with self.assertRaises(RuntimeError) as caught:
            self.authorize(browser=False, pastes=[GOOD_PASTE, GOOD_PASTE])
        self.assertIn("invalid_grant", str(caught.exception))
        self.assertEqual(len(self.flow.token_calls), 1)
        self.assertFalse(self.token_path.exists())


class TestPasteValidator(unittest.TestCase):
    """The pasted URL carries a live authorization code, so it is checked first."""

    def accepted(self, url: str) -> bool:
        response, _reason = gclass_auth._authorization_response(url)
        return response is not None

    def test_10_only_a_loopback_redirect_is_accepted(self) -> None:
        self.assertTrue(self.accepted(GOOD_PASTE))
        self.assertTrue(self.accepted("http://localhost:8910/?state=a&code=b"))
        self.assertFalse(self.accepted("https://evil.example/?state=a&code=b"))
        self.assertFalse(self.accepted("http://u:p@127.0.0.1:8910/?state=a&code=b"))
        self.assertFalse(self.accepted("http://127.0.0.1.evil.example/?state=a&code=b"))

    def test_11_exactly_one_code_and_one_state_are_required(self) -> None:
        base = gclass_auth.PASTE_REDIRECT_URI
        self.assertFalse(self.accepted(f"{base}?state=a"))
        self.assertFalse(self.accepted(f"{base}?code=b"))
        self.assertFalse(self.accepted(f"{base}?state=a&code=b&code=c"))
        self.assertFalse(self.accepted(f"{base}?state=&code=b"))

    def test_12_junk_input_is_refused_with_a_reason(self) -> None:
        for junk in ("", "   ", None, "http://127.0.0.1:8910/?code=b\nstate=a"):
            response, reason = gclass_auth._authorization_response(junk)
            self.assertIsNone(response)
            self.assertTrue(reason)
        long_url = f"{GOOD_PASTE}&pad=" + "x" * gclass_auth._MAX_REDIRECT_URL_CHARS
        self.assertFalse(self.accepted(long_url))


class TestBrowserChoice(_AuthCase):
    def test_13_a_machine_with_a_browser_still_uses_the_local_server(self) -> None:
        """The good path is untouched: a desktop finishes without any pasting."""
        self.authorize(browser=True)
        self.assertEqual(len(self.flow.server_calls), 1)
        self.assertEqual(self.flow.server_calls[0].get("host"), "127.0.0.1")
        self.assertEqual(self.flow.server_calls[0].get("port"), 0)
        self.assertIs(self.flow.server_calls[0].get("open_browser"), True)
        self.assertEqual(self.flow.auth_url_calls, [])

    def test_14_the_flag_forces_the_paste_path_on_a_desktop(self) -> None:
        """Authorizing here for a consent that happens somewhere else."""
        self.authorize(browser=True, open_browser=False, pastes=[GOOD_PASTE])
        self.assertEqual(self.flow.server_calls, [])
        self.assertEqual(len(self.flow.token_calls), 1)

    def test_15_true_is_obeyed_even_where_detection_says_no(self) -> None:
        self.authorize(browser=False, open_browser=True)
        self.assertEqual(len(self.flow.server_calls), 1)

    def test_16_detection_reports_no_browser_when_the_lookup_raises(self) -> None:
        """The real detector, against the real ``webbrowser`` failure mode."""
        import webbrowser

        with mock.patch.object(
            webbrowser, "get", side_effect=webbrowser.Error("no browser")
        ):
            self.assertFalse(gclass_auth._browser_available())
        with mock.patch.object(webbrowser, "get", lambda *a, **k: object()):
            self.assertTrue(gclass_auth._browser_available())


class TestUnchangedPaths(_AuthCase):
    def test_17_a_valid_stored_token_authorizes_nothing(self) -> None:
        """The authorization branch is the only one that changed."""
        with open(self.token_path, "wb") as handle:
            pickle.dump(FakeCredentials(), handle)
        self.authorize(browser=False)
        self.assertEqual(self.flow.server_calls, [])
        self.assertEqual(self.flow.auth_url_calls, [])

    def test_18_the_scope_set_is_untouched(self) -> None:
        """Eight scopes, and a stored token that differs is still discarded.

        Widening this set would silently invalidate every operator's token, so it
        is pinned alongside the change that shares the file.
        """
        self.assertEqual(len(gclass_auth.SCOPES), 8)
        with open(self.token_path, "wb") as handle:
            pickle.dump(FakeCredentials(scopes=gclass_auth.SCOPES[:3]), handle)
        self.authorize(browser=True)
        self.assertEqual(len(self.flow.server_calls), 1)


class TestThreading(_AuthCase):
    def test_19_the_listing_helpers_pass_the_choice_down(self) -> None:
        """The pickers authorize too, so the choice has to reach them."""
        seen: List[Optional[bool]] = []

        def fake_credentials(_c: str, _t: str, verbose: bool = False, open_browser: Any = None):
            seen.append(open_browser)
            return FakeCredentials()

        empty = types.SimpleNamespace(
            list=lambda **k: types.SimpleNamespace(execute=lambda: {})
        )

        class FakeService:
            def courses(self) -> Any:
                return types.SimpleNamespace(
                    list=empty.list, students=lambda: empty
                )

        with mock.patch.object(
            gclass_auth, "_get_google_classroom_credentials", fake_credentials
        ), mock.patch.object(gclass_auth, "build", lambda *a, **k: FakeService()):
            gclass_auth.list_google_classroom_courses("c", "t", open_browser=False)
            gclass_auth.list_google_classroom_students(
                "c", "t", course_id="1", open_browser=False
            )
        self.assertEqual(seen, [False, False])


if __name__ == "__main__":
    unittest.main(verbosity=2)
