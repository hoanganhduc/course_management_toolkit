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


# Short options, pinned.
#
# Most short options are not written anywhere: _generate_short_aliases (core.py)
# invents one for every long flag that lacks one, walking parser._actions in
# registration order and taking the initials of the words in the flag name.  A
# flag registered in the middle of the list can therefore take the candidate a
# later flag would have had, and that later flag silently moves to its fallback
# -- a short option a user has in a script stops meaning what it meant.
#
# The two tables below are that mapping as it stands.  Adding a flag adds an
# entry, which is fine; changing an entry is the failure this guards against.
DECLARED_SHORT_OPTIONS = {
    "-A": "--all-details",
    "-B": "--export-blackboard-counts",
    "-D": "--db",
    "-E": "--export-all-details",
    "-L": "--ocr-lang",
    "-O": "--ocr-service",
    "-S": "--search",
    "-T": "--simple-text",
    "-a": "--add-file",
    "-aa": "--add-canvas-announcement",
    "-ags": "--add-canvas-grading-scheme",
    "-b": "--add-blackboard-counts",
    "-c": "--config",
    "-cac": "--canvas-assignment-category",
    "-cc": "--canvas-course-id",
    "-ccd": "--change-canvas-deadlines",
    "-ccfg": "--clear-config",
    "-ccl": "--change-canvas-lock-dates",
    "-ccode": "--course-code",
    "-ccred": "--clear-credentials",
    "-cdc": "--canvas-deadline-category",
    "-cfg": "--config",
    "-clc": "--canvas-lock-category",
    "-cm": "--list-canvas-members",
    "-cs": "--comment-canvas-submission",
    "-css": "--check-student-submission-similarity",
    "-cu": "--search-canvas-user",
    "-d": "--details",
    "-da": "--download-canvas-assignment",
    "-db": "--db",
    "-dd": "--download-dest-dir",
    "-deg": "--delete-empty-canvas-groups",
    "-dgcs": "--download-google-classroom-submissions",
    "-e": "--export-emails",
    "-ean": "--export-anonymized",
    "-egs": "--export-canvas-grading-scheme",
    "-ema": "--extract-multichoice-answers",
    "-ems": "--extract-multichoice-solutions",
    "-en": "--export-emails-and-names",
    "-ep": "--edit-canvas-pages",
    "-er": "--export-canvas-rubrics",
    "-ero": "--export-roster",
    "-et": "--export-type",
    "-fea": "--final-evals-announce",
    "-fec": "--final-evals-course-id",
    "-ff": "--filter-file",
    "-fm": "--fetch-canvas-messages",
    "-ga": "--grade-canvas-assignment",
    "-gci": "--google-course-id",
    "-gcp": "--google-credentials-path",
    "-gfe": "--generate-final-evaluations",
    "-ggc": "--grade-google-classroom",
    "-grs": "--grade-resubmission",
    "-gsh": "--add-google-sheet",
    "-gtp": "--google-token-path",
    "-h": "--help",
    "-ie": "--invite-canvas-email",
    "-if": "--invite-canvas-file",
    "-igc": "--invite-google-classroom",
    "-imr": "--import-canvas-rubrics",
    "-ir": "--invite-canvas-role",
    "-l": "--load",
    "-lam": "--list-ai-models",
    "-led": "--list-email-domain",
    "-lgc": "--list-google-courses",
    "-lm": "--list-multiple-submissions-on-time",
    "-log": "--load-override-grades",
    "-lr": "--list-canvas-rubrics",
    "-m": "--modify",
    "-ncd": "--new-canvas-due-date",
    "-ncl": "--new-canvas-lock-date",
    "-nr": "--notify-incomplete-reviews",
    "-nres": "--no-restricted",
    "-p": "--print-blackboard-counts",
    "-rai": "--review-assignment-id",
    "-rid": "--rubric-assignment-id",
    "-s": "--save",
    "-sc": "--sync-canvas",
    "-sc50": "--sync-classroom50",
    "-sfe": "--send-final-evaluations",
    "-sgc": "--sync-google-classroom",
    "-sme": "--sync-multichoice-evaluations",
    "-sn": "--sheet-name",
    "-ss": "--sheet-selection",
    "-t": "--extract-text",
    "-tai": "--test-ai",
    "-tam": "--test-ai-model",
    "-type": "--export-type",
    "-ugc": "--unenroll-google-classroom",
    "-ume": "--update-mat-excel",
    "-ur": "--update-canvas-rubrics",
    "-uri": "--update-canvas-rubric-id",
    "-v": "--verbose",
    "-vcf": "--export-vcf",
    "-x": "--export-excel",
}

GENERATED_SHORT_ALIASES = {
    "-ve": "--version",
    "-dr": "--dry-run",
    "-ld": "--log-dir",
    "-ll": "--log-level",
    "-lmb": "--log-max-bytes",
    "-lb": "--log-backups",
    "-r": "--refine",
    "-bc": "--backup-config",
    "-rc": "--restore-config",
    "-cbk": "--config-backup-keep",
    "-tagm": "--test-ai-gemini-model",
    "-tahm": "--test-ai-huggingface-model",
    "-ssm": "--student-sort-method",
    "-llc": "--local-llm-command",
    "-llm": "--local-llm-model",
    "-lla": "--local-llm-args",
    "-llgd": "--local-llm-gguf-dir",
    "-dla": "--detect-local-ai",
    "-drr": "--dry-run-rows",
    "-efgd": "--export-final-grade-distribution",
    "-bd": "--backup-db",
    "-rd": "--restore-db",
    "-dbk": "--db-backup-keep",
    "-vd": "--validate-data",
    "-smc": "--sync-mat-canvas",
    "-smt": "--sync-mat-types",
    "-egd": "--export-grade-diff",
    "-ldn": "--list-duplicate-names",
    "-lmi": "--list-missing-ids",
    "-lmf": "--list-missing-form",
    "-lii": "--list-invalid-info",
    "-lda": "--list-duplicate-accounts",
    "-ar": "--audit-roster",
    "-ca": "--classroom-announcement",
    "-acg": "--audit-check-github",
    "-asd": "--audit-school-domain",
    "-as": "--audit-section",
    "-afu": "--announcement-form-url",
    "-ad": "--announcement-deadline",
    "-lss": "--list-submission-status",
    "-mif": "--missing-ids-format",
    "-mio": "--missing-ids-output",
    "-dnf": "--duplicate-name-field",
    "-dnfo": "--duplicate-name-format",
    "-dno": "--duplicate-name-output",
    "-ac": "--add-csv",
    "-o": "--onboard",
    "-oc": "--onboard-csv",
    "-osgc": "--onboard-skip-github-check",
    "-nob": "--no-open-browser",
    "-ii": "--import-internships",
    "-ire": "--import-registrations",
    "-ipr": "--import-progress-reports",
    "-imp": "--import-mini-projects",
    "-mpls": "--mini-project-lecturer-sheet",
    "-mprs": "--mini-project-registration-sheet",
    "-ic": "--import-companies",
    "-ec": "--export-companies",
    "-eman": "--evaluate-multichoice-answers",
    "-lca": "--list-canvas-assignments",
    "-uc": "--unenroll-canvas",
    "-cud": "--canvas-unenroll-domain",
    "-cua": "--canvas-unenroll-all",
    "-cue": "--canvas-unenroll-email",
    "-cus": "--canvas-unenroll-select",
    "-cumsi": "--canvas-unenroll-missing-student-id",
    "-at": "--announcement-title",
    "-am": "--announcement-message",
    "-af": "--announcement-file",
    "-icn": "--invite-canvas-name",
    "-i": "--invite",
    "-iro": "--invite-role",
    "-is": "--invite-section",
    "-ccg": "--create-canvas-groups",
    "-gsi": "--group-set-id",
    "-ng": "--num-groups",
    "-gnp": "--group-name-pattern",
    "-kog": "--keep-old-grade",
    "-lgs": "--list-google-students",
    "-gcid": "--gc-coursework-id",
    "-ggs": "--gc-grade-score",
    "-gig": "--gc-include-graded",
    "-gaa": "--gc-apply-all",
    "-gdci": "--gc-download-coursework-id",
    "-gddd": "--gc-download-dest-dir",
    "-gmt": "--gc-meaningful-threshold",
    "-gst": "--gc-similarity-threshold",
    "-gos": "--gc-ocr-service",
    "-gol": "--gc-ocr-lang",
    "-gie": "--gc-invite-email",
    "-gid": "--gc-invite-domain",
    "-gis": "--gc-invite-section",
    "-gic": "--gc-invite-class",
    "-gir": "--gc-invite-role",
    "-gia": "--gc-invite-all",
    "-gud": "--gc-unenroll-domain",
    "-gua": "--gc-unenroll-all",
    "-gue": "--gc-unenroll-email",
    "-gus": "--gc-unenroll-select",
    "-gumsi": "--gc-unenroll-missing-student-id",
    "-co": "--classroom50-org",
    "-cca": "--classroom50-classroom",
    "-cas": "--classroom50-assignment",
    "-lcc": "--list-classroom50-classrooms",
    "-lcr": "--list-classroom50-roster",
    "-lcas": "--list-classroom50-assignments",
    "-lcm": "--list-classroom50-membership",
    "-ecr": "--export-classroom50-roster",
    "-cr": "--classroom50-report",
    "-dc": "--download-classroom50",
    "-cdd": "--classroom50-download-dest",
    "-ics": "--import-classroom50-scores",
    "-csr": "--classroom50-scores-report",
    "-icg": "--import-classroom50-groups",
    "-cga": "--classroom50-group-assignment",
    "-crtj": "--classroom50-read-team-json",
    "-cgr": "--classroom50-groups-report",
    "-ipi": "--import-project-issues",
    "-pir": "--project-issues-repo",
    "-pire": "--project-issues-report",
    "-pirp": "--project-issues-read-proposal",
    "-cp": "--classroom50-preflight",
    "-rwa": "--run-weekly-automation",
    "-wai": "--weekly-assignment-id",
    "-wdd": "--weekly-dest-dir",
    "-wtci": "--weekly-teacher-canvas-id",
    "-wc": "--weekly-category",
    "-wmt": "--weekly-meaningful-threshold",
    "-wst": "--weekly-similarity-threshold",
    "-ws": "--weekly-score",
    "-wr": "--weekly-refine",
    "-wos": "--weekly-ocr-service",
    "-wol": "--weekly-ocr-lang",
    "-wnm": "--weekly-notify-missing",
    "-rwl": "--run-weekly-local",
    "-wlr": "--weekly-local-root",
    "-gww": "--generate-weekly-workflow",
    "-wtr": "--workflow-toolkit-repo",
    "-wtb": "--workflow-toolkit-branch",
    "-wsr": "--workflow-students-repo",
    "-wsb": "--workflow-students-branch",
    "-waid": "--workflow-assignment-id",
    "-wcc": "--workflow-course-code",
    "-wci": "--workflow-course-id",
    "-wtcid": "--workflow-teacher-canvas-id",
    "-bcc": "--build-course-calendar",
    "-ci": "--calendar-input",
    "-cw": "--calendar-weeks",
    "-cew": "--calendar-extra-week",
    "-ccc": "--calendar-course-code",
    "-ccn": "--calendar-course-name",
    "-cod": "--calendar-output-dir",
    "-cob": "--calendar-output-base",
    "-ceh": "--calendar-extra-holidays",
    "-icci": "--import-canvas-calendar-ics",
    "-sd": "--skip-duplicates",
    "-f": "--force",
    "-lcal": "--list-cli-aliases",
}


def _short_options(parser):
    """Both kinds of short option: the ones declared, and the ones invented."""
    declared = {}
    for action in parser._actions:
        long_opts = [opt for opt in action.option_strings if opt.startswith("--")]
        for opt in action.option_strings:
            if opt.startswith("-") and not opt.startswith("--"):
                declared[opt] = long_opts[0] if long_opts else opt

    import course_hoanganhduc.core as core

    return declared, core._generate_short_aliases(parser)


def _check_short_options(parser):
    """No pinned short option may come to mean a different flag."""
    failures = []
    declared, generated = _short_options(parser)
    for pinned, live, kind in (
        (DECLARED_SHORT_OPTIONS, declared, "declared"),
        (GENERATED_SHORT_ALIASES, generated, "generated"),
    ):
        for short, expected in pinned.items():
            actual = live.get(short)
            if actual is None:
                failures.append(f"{kind} short option {short} ({expected}) is gone")
            elif actual != expected:
                failures.append(
                    f"{kind} short option {short} moved from {expected} to {actual}"
                )
    return failures


# Three Classroom50 flags predate the metavar rule below and are left as they
# are; the rule applies to everything registered after them.
_CLASSROOM50_METAVAR_EXEMPT = frozenset(
    {"--classroom50-org", "--classroom50-classroom", "--classroom50-assignment"}
)


def _check_classroom50_metavars(parser):
    """Every Classroom50 flag that takes a value has to name what it takes.

    _dummy_value reads the metavar to decide what to feed a flag, so a flag with
    no metavar is smoke-tested with the literal string "dummy" -- it parses, and
    proves nothing about a flag that wants a path or a slug.  The help output has
    the same problem for the reader.
    """
    failures = []
    groups = [g for g in parser._action_groups if g.title == "Classroom50"]
    if not groups:
        return ["the Classroom50 argument group is gone"]
    for action in groups[0]._group_actions:
        if _is_noarg_action(action) or action.choices:
            continue
        flag = next(
            (opt for opt in action.option_strings if opt.startswith("--")),
            action.option_strings[0],
        )
        if flag in _CLASSROOM50_METAVAR_EXEMPT:
            continue
        if not action.metavar:
            failures.append(f"{flag} takes a value but declares no metavar")
    return failures


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

    problems = _check_short_options(parser) + _check_classroom50_metavars(parser)

    if failures or problems:
        if failures:
            print("CLI parse failures:")
            for flag, args, code in failures:
                print(f"- {flag} (exit {code}): {args}")
        for problem in problems:
            print(f"- {problem}")
        return 1

    print(f"CLI parse OK for {tested} flags (with --dry-run).")
    print(
        f"Short options OK: {len(DECLARED_SHORT_OPTIONS)} declared, "
        f"{len(GENERATED_SHORT_ALIASES)} generated, none reassigned."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
