import argparse
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

# Provide lightweight stubs for optional dependencies to allow parser import.
import types


def _ensure_module(name, attrs=None):
    if name in sys.modules:
        return
    module = types.ModuleType(name)
    if attrs:
        for key, value in attrs.items():
            setattr(module, key, value)
    sys.modules[name] = module


_ensure_module("paddleocr", {"PaddleOCR": object})
_ensure_module("canvasapi", {"Canvas": object})
_ensure_module("googleapiclient")
_ensure_module("googleapiclient.discovery", {"build": lambda *a, **k: None})
_ensure_module("googleapiclient.http", {"MediaIoBaseDownload": object})
_ensure_module("googleapiclient.errors", {"HttpError": Exception})
_ensure_module("google")
_ensure_module("google.oauth2")
_ensure_module("google.oauth2.credentials", {"Credentials": object})
_ensure_module("google_auth_oauthlib")
_ensure_module("google_auth_oauthlib.flow", {"InstalledAppFlow": object})
_ensure_module("google.auth")
_ensure_module("google.auth.transport")
_ensure_module("google.auth.transport.requests", {"Request": object})
_ensure_module("pandas")
_ensure_module("openpyxl")
_ensure_module("openpyxl.styles", {"Alignment": object})
_ensure_module("pytesseract")
_ensure_module("pdf2image", {"convert_from_path": lambda *a, **k: []})
_ensure_module("PIL", {"Image": object, "ImageOps": object, "ImageFilter": object})
_ensure_module("PyPDF2")
_ensure_module("numpy")
_ensure_module("cv2")
_ensure_module("sklearn")
_ensure_module("sklearn.feature_extraction")
_ensure_module("sklearn.feature_extraction.text", {"TfidfVectorizer": object})
_ensure_module("sklearn.metrics")
_ensure_module("sklearn.metrics.pairwise", {"cosine_similarity": lambda *a, **k: None})
_ensure_module("tqdm", {"tqdm": lambda x, **k: x})


def _capture_parser():
    parsers = []
    original_init = argparse.ArgumentParser.__init__

    def capture_init(self, *args, **kwargs):
        original_init(self, *args, **kwargs)
        parsers.append(self)

    argparse.ArgumentParser.__init__ = capture_init
    os.environ["COURSE_PARSE_ONLY"] = "1"
    sys.argv = ["course"]
    try:
        import course_hoanganhduc.core as core
        core.main()
    finally:
        argparse.ArgumentParser.__init__ = original_init
    if not parsers:
        raise RuntimeError("Failed to capture parser from core.main()")
    return parsers[-1]


def _dummy_value(action):
    if action.choices:
        return str(next(iter(action.choices)))
    if action.type is int:
        return "1"
    if action.type is float:
        return "1.0"
    metavar = action.metavar or ""
    upper = str(metavar).upper()
    if any(token in upper for token in ("URL",)):
        return "https://example.com"
    if any(token in upper for token in ("DIR", "FOLDER")):
        return "C:\\temp"
    if any(token in upper for token in ("FILE", "PATH", "CSV", "XLSX", "TXT", "PDF", "ICS", "JSON")):
        return "dummy.txt"
    if "DATE" in upper:
        return "2025-01-01 00:00"
    if "LANG" in upper:
        return "en"
    return "dummy"


def _is_noarg_action(action):
    return action.nargs == 0 or isinstance(
        action,
        (
            argparse._StoreTrueAction,
            argparse._StoreFalseAction,
            argparse._CountAction,
        ),
    )


def main():
    parser = _capture_parser()
    failures = []
    tested = 0

    for action in parser._actions:
        if isinstance(action, (argparse._HelpAction, argparse._VersionAction)):
            continue
        if not action.option_strings:
            continue
        long_opt = next((opt for opt in action.option_strings if opt.startswith("--")), None)
        flag = long_opt or action.option_strings[0]
        args = ["--dry-run"]
        if flag != "--dry-run":
            args.append(flag)
        if not _is_noarg_action(action) and flag != "--dry-run":
            if action.nargs in (None, 1, "?"):
                args.append(_dummy_value(action))
            elif action.nargs in ("*", "+"):
                args.append(_dummy_value(action))
            elif isinstance(action.nargs, int):
                for _ in range(action.nargs):
                    args.append(_dummy_value(action))
        try:
            parser.parse_args(args)
            tested += 1
        except SystemExit as exc:
            failures.append((flag, args, exc.code))

    if failures:
        print("CLI parse failures:")
        for flag, args, code in failures:
            print(f"- {flag} (exit {code}): {args}")
        return 1

    print(f"CLI parse OK for {tested} flags (with --dry-run).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
