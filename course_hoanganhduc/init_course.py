"""Course initialization: from an empty folder to a course the toolkit can run.

Nothing here re-implements a platform listing. Google Classroom courses come from
``gclass_auth.list_google_classroom_courses``, Classroom50 classrooms from
``c50_ops.list_classrooms``, and Canvas courses from the client that ``canvas_auth``
builds. All three are reached through :class:`Discovery`, whose methods are the only
place in this module that import those packages -- so the module stays importable, and
testable, without googleapiclient, canvasapi, or pandas.

The wizard prefers showing a list over asking for an identifier, because finding a
Google Classroom course id by hand means reading a base64 fragment out of a browser URL.
When a platform cannot be listed, or the account has nothing on it yet, the wizard asks
whether that platform should be set up and otherwise prints the steps for finding the id
(see ``GUIDANCE``). A platform that is missing is recorded and skipped; it never fails
the initialization.

Config is written here rather than through ``config.update_config_values`` because that
helper echoes the whole update dict under ``--verbose`` and under ``DRY_RUN``, and the
inherited values include API keys. The config *location* still comes from
``config.get_config_base_dir`` so the per-platform layout is defined in one place; that
helper is used rather than ``get_default_config_path`` because the latter creates the
course folder as a side effect, which a dry run must not do.
"""

import json
import os
import pickle
import re
import shutil
from collections import namedtuple

from .config import get_config_base_dir


CONFIG_VERSION = "0.1.2"

# Keys that are the same for every course this account runs. Inherited wholesale from a
# source course so a new course never has to re-enter an API key.
ACCOUNT_WIDE_KEYS = (
    "DEFAULT_AI_METHOD",
    "ALL_AI_METHODS",
    "GEMINI_API_KEY",
    # GEMINI_API_KEY2 is deliberately absent. Three existing course configs hold it, but
    # nothing in the toolkit reads it: it has no global in settings.py, so even adding it
    # to config.load_config's allowlist would not make it reach any module. Copying a
    # live API key into every new course for no reader is exposure without a use.
    "HUGGINGFACE_API_KEY",
    "GEMINI_DEFAULT_MODEL",
    "DEFAULT_OCR_METHOD",
    "ALL_OCR_METHODS",
    "OCRSPACE_API_KEY",
    "OCRSPACE_API_URL",
    "LOCAL_LLM_COMMAND",
    "LOCAL_LLM_MODEL",
    "LOCAL_LLM_ARGS",
    "LOCAL_LLM_GGUF_DIR",
    "CANVAS_LMS_API_URL",
    "CANVAS_LMS_API_KEY",
    "UNIVERSITY_NAME",
    "STUDENT_SORT_METHOD",
    "DB_BACKUP_KEEP",
    "CONFIG_BACKUP_KEEP",
)

# Keys this wizard can set for the new course. Everything else falls back to settings.py.
COURSE_SPECIFIC_KEYS = (
    "COURSE_CODE",
    "COURSE_NAME",
    "GOOGLE_CLASSROOM_COURSE_ID",
    "GOOGLE_SHEET_URL",
    "CANVAS_LMS_COURSE_ID",
    "CLASSROOM50_ORG",
    "CLASSROOM50_CLASSROOM",
)

# Printed when a platform cannot be listed, or when the operator asks for help at a
# picker. Kept as data so a test can assert guidance was offered without matching console
# output.
GUIDANCE = {
    "google_classroom": [
        "Tìm ID khoá Google Classroom:",
        "  1. Mở https://classroom.google.com và vào khoá cần dùng.",
        "  2. Nhìn thanh địa chỉ: .../c/NjE4OTQ1MzE5NTAx  <- đoạn sau /c/ là ID đã mã hoá.",
        "  3. Dán nguyên cả đường dẫn vào đây, công cụ sẽ tự giải mã ra ID dạng số.",
        "Nếu khoá chưa tồn tại: bấm dấu + ở góc phải trên Google Classroom để tạo khoá,",
        "sau đó chạy lại lệnh này (hoặc chọn 0 để thiết lập sau).",
        "Nếu danh sách khoá trống dù khoá đã tồn tại: tài khoản đăng nhập có thể không phải",
        "tài khoản giáo viên của khoá, hoặc credentials.json thuộc dự án Google khác.",
        "Nếu báo thiếu credentials.json: kế thừa từ một môn đã có (chạy lại và chọn môn nguồn),",
        "hoặc tải OAuth client dạng Desktop app từ Google Cloud Console rồi chép vào thư mục",
        "cấu hình của môn với tên credentials.json.",
    ],
    "canvas": [
        "Tìm ID khoá Canvas:",
        "  1. Mở khoá trên Canvas.",
        "  2. Đường dẫn có dạng https://<canvas>/courses/12345 -- 12345 là ID.",
        "Nếu không liệt kê được khoá nào: kiểm tra CANVAS_LMS_API_URL và CANVAS_LMS_API_KEY",
        "(kế thừa từ môn nguồn). Token Canvas tạo ở Account -> Settings -> New Access Token.",
    ],
    "classroom50": [
        "Tìm tên lớp Classroom50:",
        "  1. Cần có GitHub CLI và extension: gh extension install github/gh-teacher",
        "  2. Xem danh sách lớp: gh teacher classroom list --org <TỔ-CHỨC>",
        "  3. Tên lớp là dạng slug, ví dụ vnu-hus-mat3508-winter-2026.",
        "Nếu lớp chưa có: tạo trên giao diện Classroom50 trước, rồi chạy lại lệnh này.",
    ],
    "google_sheet": [
        "Lấy đường dẫn Google Sheet đăng ký:",
        "  1. Mở Google Form đăng ký -> tab Responses -> biểu tượng Sheets để mở bảng trả lời.",
        "  2. Sao chép nguyên đường dẫn trên thanh địa chỉ, kể cả phần #gid=... ở cuối,",
        "     vì phần đó chọn đúng sheet con chứa câu trả lời.",
        "Nếu chưa có form: tạo form đăng ký trước, rồi chạy lại lệnh này.",
    ],
}

# Wrapper scripts dropped into the course folder. Generated rather than copied: the
# copies already sitting in older course folders hard-code the POSIX venv layout, which
# does not exist on Windows.
_CONSOLE_SCRIPTS = ("course", "course-c50-admin", "course-gclass-admin")

_WRAPPER_SH = """#!/usr/bin/env sh
# Gọi lệnh {name} cho môn học trong thư mục này.
# Sinh bởi `course --init-course`.
for candidate in "$HOME/.course_venv/bin/{name}" "$HOME/.course_venv/Scripts/{name}.exe"; do
  [ -x "$candidate" ] && exec "$candidate" "$@"
done
exec {name} "$@"
"""

_WRAPPER_BAT = """@echo off
rem Goi lenh {name} cho mon hoc trong thu muc nay.
rem Sinh boi `course --init-course`.
if exist "%USERPROFILE%\\.course_venv\\Scripts\\{name}.exe" (
  "%USERPROFILE%\\.course_venv\\Scripts\\{name}.exe" %*
) else (
  {name} %*
)
"""

_SECRET_KEY_RE = re.compile(r"(secret|token|key|password|api)", re.IGNORECASE)

# The /c/<token> segment of a Classroom URL; the token is base64 of the decimal course id.
_CLASSROOM_URL_RE = re.compile(r"classroom\.google\.com/(?:u/\d+/)?c/([A-Za-z0-9_\-=]+)")
_SHEET_URL_RE = re.compile(r"/spreadsheets/d/([A-Za-z0-9_-]{20,})")
_SHEET_ID_RE = re.compile(r"^[A-Za-z0-9_-]{20,}$")


SourceCourse = namedtuple("SourceCourse", "code config_path config_dir inherited_keys mtime")
PlatformOutcome = namedtuple("PlatformOutcome", "key label state values note")
ScaffoldResult = namedtuple("ScaffoldResult", "created skipped")
InitResult = namedtuple(
    "InitResult",
    "course_code course_dir config_path values outcomes scaffold source dry_run",
)


class InitCourseError(Exception):
    """Initialization cannot continue: bad course code, or a config already there."""


### Pure helpers

def _is_set(value):
    """True when a config value carries something worth copying."""
    if value is None:
        return False
    if isinstance(value, str):
        return bool(value.strip())
    if isinstance(value, (list, tuple, dict, set)):
        return bool(value)
    return True


def redact(key, value):
    """Render a config value for display, hiding anything that looks like a secret."""
    if _SECRET_KEY_RE.search(str(key)) and isinstance(value, str) and value.strip():
        return "<đã ẩn, {} ký tự>".format(len(value))
    return value


def normalize_course_code(value):
    """Return the display form of a course code (upper case, trimmed), or None."""
    if value is None:
        return None
    value = str(value).strip()
    if not value:
        return None
    if re.search(r"[\\/:*?\"<>|\s]", value):
        raise InitCourseError(
            "Mã môn không được chứa khoảng trắng hay ký tự đường dẫn: {!r}".format(value)
        )
    return value.upper()


def course_id_from_classroom_url(value):
    """Best effort: turn a Google Classroom URL into the numeric course id.

    The ``/c/<token>`` segment decodes as base64 of the decimal id, so a pasted URL can be
    resolved without a network call. This is not a documented guarantee, so the caller
    must confirm the result before writing it. ``None`` means "could not decode; ask".
    """
    if value is None:
        return None
    value = str(value).strip()
    if not value:
        return None
    if value.isdigit():
        return value
    match = _CLASSROOM_URL_RE.search(value)
    if not match:
        return None
    token = match.group(1)
    import base64

    padded = token + "=" * (-len(token) % 4)
    try:
        decoded = base64.urlsafe_b64decode(padded.encode("ascii")).decode("ascii", "ignore")
    except Exception:
        return None
    decoded = decoded.strip()
    return decoded if decoded.isdigit() else None


def normalize_sheet_url(value):
    """Return a usable Google Sheet URL, or None.

    A pasted URL is kept verbatim -- the trailing ``#gid=`` selects the sheet tab holding
    the responses, so trimming it would silently read the wrong tab. Only a bare
    spreadsheet id is expanded into a URL.
    """
    if value is None:
        return None
    value = str(value).strip()
    if not value:
        return None
    if _SHEET_URL_RE.search(value):
        return value
    if _SHEET_ID_RE.match(value):
        return "https://docs.google.com/spreadsheets/d/{}/edit".format(value)
    return None


### Filesystem layer

def read_config_file(path):
    """Load a config.json, or None when it is missing or unreadable."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return None
    return data if isinstance(data, dict) else None


def config_base_dir():
    """The directory that holds one config folder per course.

    ``get_config_base_dir`` rather than ``get_default_config_path``: the latter creates
    the course folder as a side effect, which a dry run must not do.
    """
    return get_config_base_dir()


def find_source_courses(base_dir, exclude=None):
    """Course config folders that could seed a new course, richest first.

    Ranked by how many inheritable keys the config actually holds, not by how recent it
    is: the newest configs in this account are the sparsest ones.
    """
    found = []
    try:
        names = sorted(os.listdir(base_dir))
    except OSError:
        return found
    for name in names:
        if exclude and name.lower() == str(exclude).lower():
            continue
        config_dir = os.path.join(base_dir, name)
        config_path = os.path.join(config_dir, "config.json")
        if not os.path.isfile(config_path):
            continue
        data = read_config_file(config_path)
        if data is None:
            continue
        inherited = tuple(k for k in ACCOUNT_WIDE_KEYS if _is_set(data.get(k)))
        try:
            mtime = os.path.getmtime(config_path)
        except OSError:
            mtime = 0.0
        found.append(SourceCourse(name, config_path, config_dir, inherited, mtime))
    found.sort(key=lambda s: (-len(s.inherited_keys), -s.mtime, s.code))
    return found


def inherit_values(source_config):
    """Pick out the keys that are the same for every course this account runs."""
    if not isinstance(source_config, dict):
        return {}
    return {k: source_config[k] for k in ACCOUNT_WIDE_KEYS if _is_set(source_config.get(k))}


def write_course_config(config_path, values, force=False, dry_run=False):
    """Write config.json, refusing to clobber an existing one unless forced."""
    if os.path.exists(config_path) and not force:
        raise InitCourseError(
            "Đã có config tại {}. Dùng --init-force để ghi đè.".format(config_path)
        )
    if dry_run:
        return config_path
    os.makedirs(os.path.dirname(config_path) or ".", exist_ok=True)
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(values, f, ensure_ascii=False, indent=2, sort_keys=True)
    return config_path


def scaffold_course_folder(course_dir, course_code, make_db=True, dry_run=False):
    """Write .course_code, the wrapper scripts, and an empty database.

    Existing files are reported as skipped, never rewritten: a course folder that already
    has a tuned wrapper keeps it.
    """
    created = []
    skipped = []

    def _emit(name, content, executable=False):
        path = os.path.join(course_dir, name)
        if os.path.exists(path):
            skipped.append(name)
            return
        created.append(name)
        if dry_run:
            return
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write(content)
        if executable:
            try:
                os.chmod(path, 0o755)
            except OSError:
                pass

    if not dry_run:
        os.makedirs(course_dir, exist_ok=True)

    _emit(".course_code", course_code.lower() + "\n")
    for name in _CONSOLE_SCRIPTS:
        _emit(name + ".sh", _WRAPPER_SH.format(name=name), executable=True)
        _emit(name + ".bat", _WRAPPER_BAT.format(name=name))

    if make_db:
        db_path = os.path.join(course_dir, "students.db")
        if os.path.exists(db_path):
            skipped.append("students.db")
        else:
            created.append("students.db")
            if not dry_run:
                # Same shape data.save_database writes: a pickled list of students. Written
                # directly so scaffolding does not need pandas.
                with open(db_path, "wb") as f:
                    pickle.dump([], f)

    return ScaffoldResult(created, skipped)


def copy_credentials(source_dir, target_dir, dry_run=False):
    """Copy credentials.json from a source course. Returns the path, or None.

    token.pickle is deliberately not copied: a token is bound to one authorization and
    scope set, and a stale copy makes the refresh path fail in a way that reads like a
    credentials problem.
    """
    if not source_dir:
        return None
    src = os.path.join(source_dir, "credentials.json")
    if not os.path.isfile(src):
        return None
    dest = os.path.join(target_dir, "credentials.json")
    if os.path.exists(dest):
        return dest
    if dry_run:
        return dest
    os.makedirs(target_dir, exist_ok=True)
    try:
        shutil.copy2(src, dest)
    except OSError:
        return None
    return dest


### Platform layer

class Discovery:
    """The live platform lookups, in one place so tests can swap the whole object out.

    Every import sits inside a method: this module must stay importable on a machine with
    no googleapiclient and no canvasapi.
    """

    def google_courses(self, credentials_path, token_path, verbose=False):
        from .gclass_auth import list_google_classroom_courses

        courses = list_google_classroom_courses(
            credentials_path, token_path, verbose=verbose
        ) or []
        return [
            {"id": str(c.get("id")), "name": c.get("name") or "(không tên)"}
            for c in courses
            if c.get("id")
        ]

    def canvas_courses(self, api_url, api_key, verbose=False):
        from .canvas_auth import get_canvas_client

        canvas = get_canvas_client(api_url, api_key, verbose=verbose)
        out = []
        for course in canvas.get_courses():
            course_id = getattr(course, "id", None)
            if course_id is None:
                continue
            name = getattr(course, "name", None) or "(không tên)"
            out.append({"id": str(course_id), "name": name})
        return out

    def c50_classrooms(self, org):
        from .c50_ops import list_classrooms
        from .onboard import _classroom_label

        payload = list_classrooms(org)
        # Same reading of the payload as onboard.pick_c50_classroom.
        entries = payload if isinstance(payload, list) else []
        out = []
        for entry in entries:
            identifier, display = _classroom_label(entry)
            if identifier:
                out.append({"id": identifier, "name": display or identifier})
        return out


def _verify_in_list(items, wanted):
    """(ok, matched_name) for an id checked against a listing."""
    for item in items:
        if str(item.get("id")) == str(wanted):
            return True, item.get("name")
    return False, None


### Wizard

class _Ctx(object):
    """Everything the wizard steps share: I/O seams and the discovery object."""

    def __init__(self, input_fn, out, interactive, verbose, discovery):
        self.input_fn = input_fn
        self.out = out
        self.interactive = interactive
        self.verbose = verbose
        self.discovery = discovery

    def guide(self, key):
        for line in GUIDANCE.get(key, []):
            self.out(line)


def _ask_yes(ctx, question, default=True):
    suffix = " [C/k]: " if default else " [c/K]: "
    while True:
        answer = (ctx.input_fn(question + suffix) or "").strip().lower()
        if not answer:
            return default
        if answer in ("c", "co", "có", "y", "yes"):
            return True
        if answer in ("k", "khong", "không", "n", "no"):
            return False
        ctx.out("Trả lời c (có) hoặc k (không).")


def _ask_text(ctx, question, default=None):
    hint = " [{}]".format(default) if default else ""
    answer = (ctx.input_fn(question + hint + ": ") or "").strip()
    if not answer:
        return default
    if answer.lower() in ("q", "quit"):
        return None
    return answer


def choose_from_list(items, title, input_fn=input, out=print,
                     manual_hint="tự nhập ID hoặc dán đường dẫn",
                     skip_hint="bỏ qua, thiết lập sau"):
    """Numbered picker shared by every platform step.

    Returns ``("chosen", item)``, ``("manual", None)``, ``("help", None)`` or
    ``("skip", None)``. Writing the contract once is what keeps the four platforms
    behaving the same way.
    """
    out(title)
    for index, item in enumerate(items, 1):
        out("  {}. {} (ID: {})".format(index, item.get("name") or "?", item.get("id")))
    out("  m. " + manual_hint)
    out("  ?. hướng dẫn tìm ID")
    out("  0. " + skip_hint)
    prompt = "Chọn [1-{}/m/?/0]: ".format(len(items))
    while True:
        answer = (input_fn(prompt) or "").strip().lower()
        if answer in ("0", "q", "quit"):
            return ("skip", None)
        if answer == "m":
            return ("manual", None)
        if answer in ("?", "h", "help"):
            return ("help", None)
        try:
            index = int(answer)
        except ValueError:
            out("Không hiểu lựa chọn, nhập lại.")
            continue
        if 1 <= index <= len(items):
            return ("chosen", items[index - 1])
        out("Số ngoài khoảng, nhập lại.")


def _offer_setup(ctx, key, label, reason):
    """Nothing to list. Ask whether to set the platform up; guide, then offer manual entry.

    Returns True when the caller should fall through to asking for an id by hand.
    """
    ctx.out("Chưa thấy {} nào dùng được: {}".format(label, reason))
    if not ctx.interactive:
        return False
    if not _ask_yes(ctx, "Bạn có muốn thiết lập {} cho môn này không?".format(label)):
        return False
    ctx.guide(key)
    return _ask_yes(ctx, "Nhập ID {} bây giờ?".format(label), default=False)


def _pick(ctx, key, label, items, title):
    """Run the picker, looping on the guidance answer. Returns the chosen item or a tag."""
    while True:
        action, item = choose_from_list(items, title, input_fn=ctx.input_fn, out=ctx.out)
        if action == "help":
            ctx.guide(key)
            continue
        return action, item


def resolve_google_classroom(ctx, preset, credentials_path, token_path):
    key, label = "google_classroom", "Google Classroom"
    field = "GOOGLE_CLASSROOM_COURSE_ID"

    def _listing():
        if not credentials_path or not os.path.isfile(credentials_path):
            raise InitCourseError("chưa có credentials.json")
        if not ctx.interactive and not os.path.isfile(token_path):
            # The first listing has to authorize, and authorizing blocks on a browser or on
            # a pasted redirect read from stdin. Neither belongs in a non-interactive run.
            raise InitCourseError("chưa uỷ quyền Google, thiếu token.pickle")
        return ctx.discovery.google_courses(credentials_path, token_path, verbose=ctx.verbose)

    if preset:
        course_id = course_id_from_classroom_url(preset) or str(preset).strip()
        try:
            ok, name = _verify_in_list(_listing(), course_id)
        except Exception as exc:
            return PlatformOutcome(key, label, "chosen", {field: course_id},
                                   "không kiểm tra được ({})".format(exc))
        if ok:
            return PlatformOutcome(key, label, "chosen", {field: course_id}, name or "")
        return PlatformOutcome(key, label, "chosen", {field: course_id},
                               "ID không có trong danh sách khoá của tài khoản")

    if not ctx.interactive:
        return PlatformOutcome(key, label, "skipped", {}, "thiếu --init-google-id")

    try:
        items = _listing()
    except Exception as exc:
        items = None
        reason = str(exc)
    if items:
        action, item = _pick(ctx, key, label, items,
                             "Chọn khoá Google Classroom cho môn này:")
        if action == "chosen":
            return PlatformOutcome(key, label, "chosen", {field: item["id"]}, item["name"])
        if action == "skip":
            return PlatformOutcome(key, label, "skipped", {}, "bỏ qua theo lựa chọn")
    elif items is not None:
        if not _offer_setup(ctx, key, label, "tài khoản chưa có khoá nào"):
            return PlatformOutcome(key, label, "skipped", {}, "chưa có khoá nào")
    else:
        if not _offer_setup(ctx, key, label, reason):
            return PlatformOutcome(key, label, "unavailable", {}, reason)

    while True:
        raw = _ask_text(ctx, "Dán đường dẫn khoá Google Classroom hoặc ID dạng số (Enter để bỏ qua)")
        if not raw:
            return PlatformOutcome(key, label, "skipped", {}, "bỏ qua theo lựa chọn")
        course_id = course_id_from_classroom_url(raw)
        if course_id is None:
            ctx.out("Không đọc được ID từ chuỗi đó.")
            ctx.guide(key)
            continue
        if course_id != raw.strip():
            ctx.out("Giải mã được ID: {}".format(course_id))
            if not _ask_yes(ctx, "Dùng ID này?"):
                continue
        return PlatformOutcome(key, label, "manual", {field: course_id}, "nhập tay")


def resolve_canvas(ctx, preset, api_url, api_key):
    key, label = "canvas", "Canvas"
    field = "CANVAS_LMS_COURSE_ID"

    def _listing():
        if not api_url or not api_key:
            raise InitCourseError("chưa có CANVAS_LMS_API_URL / CANVAS_LMS_API_KEY")
        return ctx.discovery.canvas_courses(api_url, api_key, verbose=ctx.verbose)

    if preset:
        course_id = str(preset).strip()
        try:
            ok, name = _verify_in_list(_listing(), course_id)
        except Exception as exc:
            return PlatformOutcome(key, label, "chosen", {field: course_id},
                                   "không kiểm tra được ({})".format(exc))
        if ok:
            return PlatformOutcome(key, label, "chosen", {field: course_id}, name or "")
        return PlatformOutcome(key, label, "chosen", {field: course_id},
                               "ID không có trong danh sách khoá Canvas của tài khoản")

    if not ctx.interactive:
        return PlatformOutcome(key, label, "skipped", {}, "thiếu --init-canvas-id")

    if not _ask_yes(ctx, "Môn này có dùng Canvas không?", default=False):
        return PlatformOutcome(key, label, "skipped", {}, "không dùng Canvas")

    try:
        items = _listing()
    except Exception as exc:
        items = None
        reason = str(exc)
    if items:
        action, item = _pick(ctx, key, label, items, "Chọn khoá Canvas cho môn này:")
        if action == "chosen":
            return PlatformOutcome(key, label, "chosen", {field: item["id"]}, item["name"])
        if action == "skip":
            return PlatformOutcome(key, label, "skipped", {}, "bỏ qua theo lựa chọn")
    elif items is not None:
        if not _offer_setup(ctx, key, label, "tài khoản chưa có khoá Canvas nào"):
            return PlatformOutcome(key, label, "skipped", {}, "chưa có khoá Canvas nào")
    else:
        if not _offer_setup(ctx, key, label, reason):
            return PlatformOutcome(key, label, "unavailable", {}, reason)

    raw = _ask_text(ctx, "Nhập ID khoá Canvas dạng số (Enter để bỏ qua)")
    if not raw:
        return PlatformOutcome(key, label, "skipped", {}, "bỏ qua theo lựa chọn")
    return PlatformOutcome(key, label, "manual", {field: raw.strip()}, "nhập tay")


def resolve_classroom50(ctx, preset_org, preset_classroom, default_org=None):
    key, label = "classroom50", "Classroom50"
    org_field, room_field = "CLASSROOM50_ORG", "CLASSROOM50_CLASSROOM"

    org = preset_org or default_org
    if preset_classroom:
        if not org:
            return PlatformOutcome(key, label, "skipped", {},
                                   "có --init-c50-classroom nhưng thiếu --init-c50-org")
        return PlatformOutcome(key, label, "chosen",
                               {org_field: org, room_field: preset_classroom},
                               "theo tham số dòng lệnh")

    if not ctx.interactive:
        if org:
            return PlatformOutcome(key, label, "chosen", {org_field: org}, "chỉ đặt tổ chức")
        return PlatformOutcome(key, label, "skipped", {}, "thiếu --init-c50-classroom")

    if not _ask_yes(ctx, "Môn này có dùng Classroom50 không?", default=bool(org)):
        return PlatformOutcome(key, label, "skipped", {}, "không dùng Classroom50")

    org = _ask_text(ctx, "Tổ chức GitHub của Classroom50", default=org)
    if not org:
        return PlatformOutcome(key, label, "skipped", {}, "chưa có tổ chức")

    try:
        items = ctx.discovery.c50_classrooms(org)
    except Exception as exc:
        items = None
        reason = str(exc)
    if items:
        action, item = _pick(ctx, key, label, items,
                             "Chọn lớp Classroom50 trong tổ chức {}:".format(org))
        if action == "chosen":
            return PlatformOutcome(key, label, "chosen",
                                   {org_field: org, room_field: item["id"]}, item["name"])
        if action == "skip":
            return PlatformOutcome(key, label, "chosen", {org_field: org},
                                   "chỉ đặt tổ chức, chọn lớp sau")
    elif items is not None:
        if not _offer_setup(ctx, key, label, "tổ chức {} chưa có lớp nào".format(org)):
            return PlatformOutcome(key, label, "chosen", {org_field: org}, "chưa có lớp nào")
    else:
        if not _offer_setup(ctx, key, label, reason):
            return PlatformOutcome(key, label, "chosen", {org_field: org}, reason)

    raw = _ask_text(ctx, "Nhập tên lớp Classroom50 (Enter để bỏ qua)")
    if not raw:
        return PlatformOutcome(key, label, "chosen", {org_field: org}, "chỉ đặt tổ chức")
    return PlatformOutcome(key, label, "manual", {org_field: org, room_field: raw}, "nhập tay")


def resolve_google_sheet(ctx, preset):
    key, label = "google_sheet", "Google Sheet đăng ký"
    field = "GOOGLE_SHEET_URL"

    if preset:
        url = normalize_sheet_url(preset)
        if not url:
            return PlatformOutcome(key, label, "skipped", {},
                                   "--init-sheet-url không phải đường dẫn Google Sheet")
        return PlatformOutcome(key, label, "chosen", {field: url}, "theo tham số dòng lệnh")

    if not ctx.interactive:
        return PlatformOutcome(key, label, "skipped", {}, "thiếu --init-sheet-url")

    while True:
        raw = _ask_text(ctx, "Dán đường dẫn Google Sheet trả lời form đăng ký "
                             "(Enter để bỏ qua, ? để xem hướng dẫn)")
        if not raw:
            return PlatformOutcome(key, label, "skipped", {}, "bỏ qua theo lựa chọn")
        if raw.strip() in ("?", "h", "help"):
            ctx.guide(key)
            continue
        url = normalize_sheet_url(raw)
        if not url:
            ctx.out("Chuỗi đó không giống đường dẫn Google Sheet.")
            ctx.guide(key)
            continue
        return PlatformOutcome(key, label, "manual", {field: url}, "nhập tay")


### Driver

def _choose_source(ctx, sources, requested):
    """Pick the course to inherit account-wide settings from."""
    if requested:
        for source in sources:
            if source.code.lower() == str(requested).lower():
                return source
        raise InitCourseError(
            "Không tìm thấy môn nguồn {!r} để kế thừa cấu hình.".format(requested)
        )
    if not sources:
        return None
    if not ctx.interactive:
        return sources[0]
    ctx.out("Kế thừa cấu hình chung (API key, OCR, AI, Canvas) từ môn nào?")
    for index, source in enumerate(sources, 1):
        ctx.out("  {}. {} ({} mục dùng chung)".format(
            index, source.code, len(source.inherited_keys)))
    ctx.out("  0. không kế thừa, bắt đầu từ mặc định")
    while True:
        answer = (ctx.input_fn("Chọn [1-{}/0]: ".format(len(sources))) or "").strip()
        if answer in ("0", "q", "quit"):
            return None
        if not answer:
            return sources[0]
        try:
            index = int(answer)
        except ValueError:
            ctx.out("Không hiểu lựa chọn, nhập lại.")
            continue
        if 1 <= index <= len(sources):
            return sources[index - 1]
        ctx.out("Số ngoài khoảng, nhập lại.")


def run_init_course(course_code=None, course_name=None, directory=None, inherit_from=None,
                    google_course_id=None, canvas_course_id=None, c50_org=None,
                    c50_classroom=None, sheet_url=None, scaffold=True, make_db=True,
                    interactive=True, force=False, dry_run=False, verbose=False,
                    input_fn=input, out=print, discovery=None, base_dir=None):
    """Initialize a course: config, credentials, and the folder that runs it.

    Returns an :class:`InitResult`. Raises :class:`InitCourseError` for the two things the
    operator has to fix: an unusable course code, and a config that is already there.
    """
    ctx = _Ctx(input_fn, out, interactive, verbose, discovery or Discovery())

    course_dir = os.path.abspath(directory or os.getcwd())
    if not course_code and interactive:
        course_code = _ask_text(ctx, "Mã môn học (ví dụ MAT3508)",
                                default=os.path.basename(course_dir))
    code = normalize_course_code(course_code)
    if not code:
        raise InitCourseError("Cần mã môn học. Dùng --course-code hoặc trả lời câu hỏi.")
    slug = code.lower()

    if base_dir is None:
        base_dir = config_base_dir()
    config_dir = os.path.join(base_dir, slug)
    config_path = os.path.join(config_dir, "config.json")
    if os.path.exists(config_path) and not force:
        raise InitCourseError(
            "Đã có config tại {}. Dùng --init-force để ghi đè.".format(config_path)
        )

    if not course_name and interactive:
        course_name = _ask_text(ctx, "Tên môn học (Enter để bỏ trống)")

    sources = find_source_courses(base_dir, exclude=slug)
    source = _choose_source(ctx, sources, inherit_from)
    values = inherit_values(read_config_file(source.config_path) or {}) if source else {}
    values["CONFIG_VERSION"] = CONFIG_VERSION
    values["COURSE_CODE"] = code
    if course_name:
        values["COURSE_NAME"] = course_name

    credentials_path = copy_credentials(
        source.config_dir if source else None, config_dir, dry_run=dry_run)
    if credentials_path is None:
        existing = os.path.join(config_dir, "credentials.json")
        credentials_path = existing if os.path.isfile(existing) else None
    token_path = os.path.join(config_dir, "token.pickle")

    outcomes = [
        resolve_google_classroom(ctx, google_course_id, credentials_path, token_path),
        resolve_google_sheet(ctx, sheet_url),
        resolve_canvas(ctx, canvas_course_id,
                       values.get("CANVAS_LMS_API_URL"), values.get("CANVAS_LMS_API_KEY")),
        resolve_classroom50(ctx, c50_org, c50_classroom,
                            default_org=_default_c50_org(sources)),
    ]
    for outcome in outcomes:
        values.update(outcome.values)

    write_course_config(config_path, values, force=force, dry_run=dry_run)
    scaffold_result = (
        scaffold_course_folder(course_dir, code, make_db=make_db, dry_run=dry_run)
        if scaffold else ScaffoldResult([], [])
    )

    return InitResult(code, course_dir, config_path, values, outcomes, scaffold_result,
                      source.code if source else None, dry_run)


def _default_c50_org(sources):
    """The organization every existing course already uses, when they agree on one."""
    orgs = set()
    for source in sources:
        data = read_config_file(source.config_path) or {}
        org = data.get("CLASSROOM50_ORG")
        if _is_set(org):
            orgs.add(org)
    return orgs.pop() if len(orgs) == 1 else None


_STATE_LABELS = {
    "chosen": "đã đặt",
    "manual": "đã đặt (nhập tay)",
    "skipped": "bỏ qua",
    "unavailable": "không truy cập được",
}


def summarize(result, out=print):
    """Print what was done. Every value goes through the redactor on the way out."""
    prefix = "[thử] " if result.dry_run else ""
    out("")
    out("=" * 60)
    out("{}Khởi tạo môn {}".format(prefix, result.course_code))
    out("=" * 60)
    out("Thư mục môn : {}".format(result.course_dir))
    out("Cấu hình    : {}".format(result.config_path))
    out("Kế thừa từ  : {}".format(result.source or "(không)"))
    out("")
    out("Nền tảng:")
    for outcome in result.outcomes:
        note = " -- {}".format(outcome.note) if outcome.note else ""
        out("  {:<22} {}{}".format(
            outcome.label, _STATE_LABELS.get(outcome.state, outcome.state), note))
    out("")
    out("Giá trị đã ghi:")
    for key in sorted(result.values):
        out("  {:<36} {}".format(key, redact(key, result.values[key])))
    if result.scaffold.created or result.scaffold.skipped:
        out("")
        out("Tệp trong thư mục môn:")
        for name in result.scaffold.created:
            out("  + {}".format(name))
        for name in result.scaffold.skipped:
            out("  = {} (đã có, giữ nguyên)".format(name))

    pending = [o for o in result.outcomes if o.state in ("skipped", "unavailable")]
    out("")
    out("Bước tiếp theo:")
    if result.dry_run:
        out("  - Bỏ --dry-run để ghi thật.")
    for outcome in pending:
        out("  - {} chưa đặt: chạy lại `course --init-course --init-force` sau khi có ID,"
            .format(outcome.label))
        out("    hoặc sửa trực tiếp {}".format(result.config_path))
    out("  - Nạp danh sách đăng ký:  course --add-google-sheet")
    out("  - Kéo sĩ số lớp:          course --sync-google-classroom")
