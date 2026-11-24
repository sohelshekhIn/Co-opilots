"""Desktop actions across every attached monitor.

ponytail: a screenshot is one monitor unless the work is on another. Two monitors
side by side are one image only when both are needed. Upgrade path is stitching
monitors that are not a single rectangle.
"""

import ctypes
import os
import re
import subprocess
import time
import urllib.parse
import webbrowser
from ctypes import wintypes
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

if not ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        pass

import pyautogui
import pygetwindow
import pyperclip
from PIL import Image, ImageGrab

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05


def dbg(hypothesis, location, message, **data):
    # #region agent log
    try:
        import json
        payload = {"sessionId": "44cfa9", "hypothesisId": hypothesis, "location": location,
                   "message": message, "data": data, "timestamp": int(time.time() * 1000)}
        with open(r"c:\Users\ShekhSo\OneDrive - Government of Ontario\Documents\Scripts\Co-opilots\.cursor\debug-44cfa9.log", "a", encoding="utf-8") as handle:
            handle.write(json.dumps(payload, default=str) + "\n")
    except Exception:
        pass
    # #endregion


_UIA_KIND = {
    50000: "Button", 50002: "CheckBox", 50003: "ComboBox", 50004: "Edit", 50005: "Hyperlink",
    50007: "ListItem", 50008: "List", 50009: "Menu", 50011: "MenuItem", 50013: "RadioButton",
    50018: "Tab", 50019: "TabItem", 50020: "Text", 50021: "ToolBar", 50025: "Custom",
    50026: "Group", 50030: "Document", 50031: "SplitButton", 50032: "Window", 50033: "Pane",
    50036: "Table", 50037: "TitleBar",
}


class _Guid(ctypes.Structure):
    _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD), ("Data3", wintypes.WORD), ("Data4", wintypes.BYTE * 8)]


class _Variant(ctypes.Structure):
    _fields_ = [("vt", ctypes.c_ushort), ("r1", ctypes.c_ushort), ("r2", ctypes.c_ushort),
                ("r3", ctypes.c_ushort), ("ll", ctypes.c_ulonglong)]


def _guid(text):
    value = _Guid()
    ctypes.windll.ole32.CLSIDFromString(text, ctypes.byref(value))
    return value


def _release(pointer):
    if not pointer:
        return
    table = ctypes.cast(pointer, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)))[0]
    ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(table[2])(pointer)


def _uia_prop(element, prop):
    table = ctypes.cast(element, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)))[0]
    value = _Variant()
    hr = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, ctypes.c_int, ctypes.POINTER(_Variant))(table[10])(
        element, prop, ctypes.byref(value))
    if hr < 0:
        raise OSError(f"UI Automation property {prop} failed ({hr & 0xFFFFFFFF:08x})")
    if value.vt == 8:
        text = ctypes.wstring_at(value.ll) if value.ll else ""
        if value.ll:
            ctypes.windll.oleaut32.SysFreeString.argtypes = [ctypes.c_void_p]
            ctypes.windll.oleaut32.SysFreeString(value.ll)
        return text
    return value.ll if value.vt == 3 else ""


def _focused_uia():
    """Name, kind, and automation id of the focused control. Windows UI Automation, no extra package."""
    ole32 = ctypes.windll.ole32
    started = ole32.CoInitialize(None)
    clsid, iid = _guid("{ff48dba4-60ef-4201-aa87-54103eef594e}"), _guid("{30cbe57d-d9d0-452a-ab13-7ac5ac4825ee}")
    automation = ctypes.c_void_p()
    hr = ole32.CoCreateInstance(ctypes.byref(clsid), None, 1, ctypes.byref(iid), ctypes.byref(automation))
    if hr < 0:
        if started >= 0:
            ole32.CoUninitialize()
        raise OSError(f"UI Automation unavailable ({hr & 0xFFFFFFFF:08x})")
    try:
        table = ctypes.cast(automation, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)))[0]
        element = ctypes.c_void_p()
        hr = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p))(table[8])(
            automation, ctypes.byref(element))
        if hr < 0 or not element:
            raise OSError(f"no focused control ({hr & 0xFFFFFFFF:08x})")
        try:
            kind = _UIA_KIND.get(_uia_prop(element, 30003), "control")
            name = _uia_prop(element, 30005) or "(unnamed)"
            return kind, name, _uia_prop(element, 30011) or ""
        finally:
            _release(element)
    finally:
        _release(automation)
        if started >= 0:
            ole32.CoUninitialize()


def focused_control():
    """Read the current Windows UI Automation focus without changing it."""
    try:
        kind, name, identifier = _focused_uia()
    except Exception as exc:
        return f"Focused control unavailable: {exc}"
    return f"Active window: {foreground_title() or '(untitled)'}; focused: {kind} {name}; automation id: {identifier}"

# Long edge of the JPEG sent to Copilot, per monitor side by side, up to _MAX_TOTAL.
_MAX_EDGE = 1280
_MAX_TOTAL = 2560
# Opening these would run a program, which open_app already covers through the Start menu.
_RUNNABLE = {".exe", ".bat", ".cmd", ".com", ".ps1", ".vbs", ".vbe", ".js", ".jse", ".wsf", ".wsh",
             ".msi", ".msp", ".scr", ".lnk", ".pif", ".reg", ".hta", ".cpl", ".url", ".appref-ms"}
_BAD_APP = re.compile(r'[&|;<>^`\n\r"%]')
# Start menu and the search box. Typing while either is in front lands in the search field.
_SEARCH_PROCS = {"searchhost.exe", "startmenuexperiencehost.exe", "shellexperiencehost.exe"}
_SEARCH_TITLES = {"search", "start", "searching"}
_KEY_NAMES = {
    "enter": "enter",
    "return": "enter",
    "tab": "tab",
    "escape": "esc",
    "esc": "esc",
    "up": "up",
    "down": "down",
    "left": "left",
    "right": "right",
    "win": "win",
    "ctrl": "ctrl",
    "alt": "alt",
    "shift": "shift",
    "space": "space",
    "backspace": "backspace",
    "delete": "delete",
    "home": "home",
    "end": "end",
    "pageup": "pageup",
    "pagedown": "pagedown",
    "volumemute": "volumemute",
    "volumeup": "volumeup",
    "volumedown": "volumedown",
    "playpause": "playpause",
    "nexttrack": "nexttrack",
    "prevtrack": "prevtrack",
    "printscreen": "printscreen",
}
# Link types open_url accepts: web pages, a new email, Windows Settings pages, and Teams deep links.
_URL_SCHEMES = {"http", "https", "mailto", "ms-settings", "msteams"}


@dataclass
class Shot:
    jpeg: bytes
    screen_w: int
    screen_h: int
    sent_w: int
    sent_h: int
    left: int = 0
    top: int = 0
    monitors: list = field(default_factory=list)

    def layout(self):
        """Monitors in screenshot pixels, e.g. 'monitor 1 (main): x 0-959, y 0-539'."""
        sx, sy = self.sent_w / self.screen_w, self.sent_h / self.screen_h
        parts = []
        for index, mon in enumerate(self.monitors, 1):
            left, top, right, bottom, primary = mon[:5]
            number = mon[5] if len(mon) > 5 else index
            parts.append(f"monitor {number}{' (main)' if primary else ''}: x {round((left - self.left) * sx)}-"
                         f"{round((right - self.left) * sx) - 1}, y {round((top - self.top) * sy)}-{round((bottom - self.top) * sy) - 1}")
        return "; ".join(parts)


def monitors():
    """(left, top, right, bottom, is_primary) per monitor in physical pixels, main first."""
    found = []
    proc = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p, ctypes.POINTER(wintypes.RECT), ctypes.c_double)

    def visit(_monitor, _dc, rect, _data):
        r = rect.contents
        found.append((r.left, r.top, r.right, r.bottom, r.left == 0 and r.top == 0))
        return 1

    ctypes.windll.user32.EnumDisplayMonitors(None, None, proc(visit), 0)
    found.sort(key=lambda m: (not m[4], m[0], m[1]))
    if not found:
        width, height = pyautogui.size()
        found = [(0, 0, int(width), int(height), True)]
    return found


def screen_size():
    width, height = pyautogui.size()
    return int(width), int(height)


_SHOT_SKIP = {"co-opilot", "developer log", "windows input experience"}
_TASK_SKIP = {"about", "after", "again", "also", "anything", "because", "before", "cant", "cannot", "could",
              "does", "from", "have", "hear", "here", "just", "need", "open", "please", "still", "that",
              "them", "then", "there", "they", "this", "want", "wearing", "what", "when", "where", "with",
              "would", "your", "screen", "monitor", "display", "window"}


def _skip_title(title):
    low = (title or "").strip().lower()
    return not low or low in _SHELL_TITLES or low in _SHOT_SKIP or low.startswith("co-opilot ")


def _rect_monitor(rect, screens):
    left, top, right, bottom = rect
    cx, cy = (left + right) // 2, (top + bottom) // 2
    for index, (l, t, r, b, _primary) in enumerate(screens, 1):
        if l <= cx < r and t <= cy < b:
            return index
    return 0


def _mentioned_monitors(task, count):
    found = set()
    for match in re.finditer(r"\b(?:monitor|display)\s*(\d+)\b", task or "", re.I):
        number = int(match.group(1))
        if 1 <= number <= count:
            found.add(number)
    if count >= 2 and re.search(r"\b(?:second|other|secondary)\s+(?:screen|monitor|display)\b", task or "", re.I):
        found.add(2)
    return found


def _task_words(task):
    return [word for word in re.findall(r"[a-z0-9]+", (task or "").lower()) if len(word) > 4 and word not in _TASK_SKIP]


def choose_monitors(screens, windows, task="", title=""):
    """Which monitors to shoot. A secondary monitor is included only when the work is on it.

    windows items are (hwnd, title, process, left, top, right, bottom, foreground).
    """
    if len(screens) < 2:
        return list(screens)
    if title and title.strip():
        rows = [(w[0], w[1], w[2]) for w in windows if not _skip_title(w[1])]
        hwnd = match_window(title, rows)
        hit = next((w for w in windows if w[0] == hwnd), None) if hwnd else None
        if hit is None:
            wanted = title.strip().lower()
            hit = next((w for w in windows if wanted in w[1].lower() and not _skip_title(w[1])), None)
        if hit is not None:
            number = _rect_monitor(hit[3:7], screens)
            if number:
                return [screens[number - 1]]
    chosen = _mentioned_monitors(task, len(screens))
    words = _task_words(task)
    for window in windows:
        if _skip_title(window[1]):
            continue
        low = window[1].lower()
        if words and any(word in low for word in words):
            number = _rect_monitor(window[3:7], screens)
            if number:
                chosen.add(number)
    if not chosen:
        for window in windows:
            if len(window) > 7 and window[7] and not _skip_title(window[1]):
                number = _rect_monitor(window[3:7], screens)
                if number:
                    chosen.add(number)
                    break
    if not chosen:
        chosen.add(1)
    return [screens[i - 1] for i in sorted(chosen)]


def screens_for(task="", title=""):
    """Live monitors for the next screenshot. Main only, unless the work is elsewhere."""
    screens = monitors()
    if len(screens) < 2:
        return screens
    foreground = _foreground_hwnd()
    windows = []
    for hwnd, wtitle, proc in taskbar_windows():
        rect = wintypes.RECT()
        if not ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            continue
        windows.append((hwnd, wtitle, proc, rect.left, rect.top, rect.right, rect.bottom, hwnd == foreground))
    return choose_monitors(screens, windows, task, title)


def screenshot(picked=None):
    all_screens = monitors()
    chosen = list(picked) if picked else all_screens
    if not chosen:
        chosen = all_screens
    left, top = min(m[0] for m in chosen), min(m[1] for m in chosen)
    right, bottom = max(m[2] for m in chosen), max(m[3] for m in chosen)
    image = ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)
    # A grab that ignored the box comes back as the whole desktop; crop it down.
    if image.size != (right - left, bottom - top):
        whole_left, whole_top = min(m[0] for m in all_screens), min(m[1] for m in all_screens)
        image = image.crop((left - whole_left, top - whole_top, right - whole_left, bottom - whole_top))
    screen_w, screen_h = image.size
    sent = image
    cap = min(_MAX_TOTAL, _MAX_EDGE * max(1, len(chosen)))
    long_edge = max(image.size)
    if long_edge > cap:
        scale = cap / long_edge
        sent = image.resize(
            (max(1, round(image.width * scale)), max(1, round(image.height * scale))),
            Image.Resampling.LANCZOS,
        )
    numbered = []
    for mon in chosen:
        number = next((i for i, screen in enumerate(all_screens, 1) if screen[:4] == mon[:4]), len(numbered) + 1)
        numbered.append((*mon[:5], number))
    buffer = BytesIO()
    sent.convert("RGB").save(buffer, format="JPEG", quality=70)
    return Shot(buffer.getvalue(), screen_w, screen_h, sent.width, sent.height, left, top, numbered)


def shot_caption(shot):
    """What Copilot is told about this screenshot, including monitors that were left out."""
    text = f"Screenshot attached: {shot.sent_w} x {shot.sent_h} pixels. Monitors: {shot.layout()}."
    shown = {mon[5] if len(mon) > 5 else index for index, mon in enumerate(shot.monitors, 1)}
    missing = [index for index in range(1, len(monitors()) + 1) if index not in shown]
    if missing:
        text += " Left out: " + ", ".join(f"monitor {index}" for index in missing) + ". The work is not on them; focus a window there to include that monitor."
    return text


def run(action, shot, on_slow=None):
    """Do one action. Returns text for Copilot when the action reads something, else None."""
    kind = action["action"]
    if kind in ("click", "double_click", "right_click"):
        x, y = _screen_xy(action["x"], action["y"], shot)
        if kind == "right_click":
            pyautogui.click(x, y, button="right")
        else:
            pyautogui.click(x, y, clicks=2 if kind == "double_click" else 1, interval=0.08)
    elif kind == "move":
        pyautogui.moveTo(*_screen_xy(action["x"], action["y"], shot), duration=0.15)
    elif kind == "drag":
        start = _screen_xy(action["x"], action["y"], shot)
        end = _screen_xy(action["to_x"], action["to_y"], shot)
        pyautogui.moveTo(*start)
        pyautogui.dragTo(*end, duration=0.5, button="left")
    elif kind == "focus":
        focus_window(action["title"])
    elif kind == "window":
        return window_op(action["title"], action["op"])
    elif kind == "open_url":
        open_url(action["url"])
    elif kind == "open_path":
        open_path(action["path"])
    elif kind == "read_clipboard":
        text = pyperclip.paste() or ""
        return text if len(text) <= 3000 else text[:3000] + "\n...(truncated)"
    elif kind == "type":
        _type_text(action["text"])
    elif kind == "key":
        _press(action["key"])
    elif kind == "scroll":
        _scroll(action, shot)
    elif kind == "open_app":
        return open_app(action["name"], on_slow)
    elif kind == "wait":
        time.sleep(min(int(action["ms"]), 10_000) / 1000)
    elif kind == "done":
        return
    else:
        raise ValueError(f"unknown action: {kind}")


def app_name_ok(name):
    if not isinstance(name, str):
        return False
    name = name.strip()
    if not name or len(name) > 200 or _BAD_APP.search(name):
        return False
    return True


def open_app(name, on_slow=None):
    """Bring the app forward if it is already on the taskbar; otherwise launch it and wait while it opens."""
    if not app_name_ok(name):
        raise ValueError(f"rejected app name: {name!r}")
    hit = match_window(name, taskbar_windows())
    if hit and _bring(hit):
        return f"{name} was already open, so it was brought to the front."
    # Start menu search resolves display names ("File Explorer") that `start` cannot.
    before = _foreground_hwnd()
    pyautogui.press("win")
    if not _wait_until(_search_open, 2.5):
        pyautogui.press("escape")
        raise ValueError("Start search did not open, so the app name was not typed")
    time.sleep(0.35)  # the search box is up, but its text field is not focused yet
    _paste(name.strip())
    time.sleep(0.7)
    pyautogui.press("enter")
    # ponytail: up to 20s for slow starters like Outlook. Leaving Search is not enough; focus often falls back to this app.
    if _wait_for_app(before, name, 20, on_slow):
        return None
    hit = match_window(name, taskbar_windows())
    if hit and _bring(hit):
        return f"{name} opened in the background, so it was brought to the front."
    raise ValueError(f"{name} did not open within 20 seconds, so nothing was typed into it")


# Taskbar names that differ from the process. Titles cover the rest ("Inbox - Outlook", "Untitled - Notepad").
_APP_PROCS = {
    "outlook": {"outlook.exe", "olk.exe"}, "teams": {"ms-teams.exe", "teams.exe"}, "edge": {"msedge.exe"},
    "word": {"winword.exe"}, "excel": {"excel.exe"}, "powerpoint": {"powerpnt.exe"}, "onenote": {"onenote.exe"},
    "file explorer": {"explorer.exe"}, "explorer": {"explorer.exe"}, "notepad": {"notepad.exe"},
    "chrome": {"chrome.exe"}, "google chrome": {"chrome.exe"}, "settings": {"systemsettings.exe"},
    "calculator": {"calculatorapp.exe"}, "visual studio code": {"code.exe"}, "vs code": {"code.exe"},
}
_SHELL_TITLES = {"program manager", "copilot assistant"}


def _app_words(name):
    return [w for w in re.findall(r"[a-z0-9]+", name.lower()) if w not in {"the", "app", "microsoft", "new", "classic"} and len(w) > 2]


def match_window(name, windows):
    """First (hwnd, title, process) window that is this app, by process alias or every name word in the title."""
    key = re.sub(r"^(microsoft|new|classic)\s+", "", name.strip().lower())
    procs, words = _APP_PROCS.get(key, set()), _app_words(name)
    for hwnd, title, proc in windows:
        low = title.strip().lower()
        if not low or low in _SHELL_TITLES or _searchish(proc, title):
            continue
        if proc in procs or (words and all(w in low for w in words)):
            return hwnd
    return None


def _starting(name):
    """True when a known process for this app is running, even before it has a window."""
    procs = _APP_PROCS.get(re.sub(r"^(microsoft|new|classic)\s+", "", name.strip().lower()), set())
    if not procs:
        return False
    try:
        out = subprocess.run(["tasklist", "/fo", "csv", "/nh"], capture_output=True, text=True, timeout=5,
                             creationflags=subprocess.CREATE_NO_WINDOW).stdout.lower()
    except (OSError, subprocess.TimeoutExpired):
        return False
    return any(f'"{p}"' in out for p in procs)


def taskbar_windows():
    """Windows that have a taskbar button: visible, titled, not owned, not tool windows, not cloaked."""
    user32, found = ctypes.windll.user32, []

    def visit(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd) or user32.GetWindow(hwnd, 4) or user32.GetWindowLongW(hwnd, -20) & 0x80:
            return True
        cloaked = wintypes.DWORD()
        ctypes.windll.dwmapi.DwmGetWindowAttribute(hwnd, 14, ctypes.byref(cloaked), ctypes.sizeof(cloaked))
        length = user32.GetWindowTextLengthW(hwnd)
        rect = wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        # The Copilot chat browser is parked at x -32000 while hidden; it is not the user's app.
        if cloaked.value or not length or (rect.left <= -30000 and not user32.IsIconic(hwnd)):
            return True
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buffer, length + 1)
        if own_console(buffer.value, hwnd):
            return True
        found.append((hwnd, buffer.value, _process_name(hwnd)))
        return True

    user32.EnumWindows(ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(visit), 0)
    return found


def _bring(hwnd):
    """Restore if minimised and bring to the front. True once it is the foreground window."""
    user32 = ctypes.windll.user32
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, 9)
    pyautogui.press("alt")  # Windows only lets a process take the foreground right after a key press.
    user32.SetForegroundWindow(hwnd)
    return _wait_until(lambda: _foreground_hwnd() == hwnd, 2)


def url_ok(url):
    if not isinstance(url, str) or re.search(r"\s", url.strip()):
        return False
    parts = urllib.parse.urlsplit(url.strip())
    if parts.scheme not in _URL_SCHEMES:
        return False
    # Windows URI handlers such as ms-settings: are valid without a host or
    # path (the bare scheme opens the Settings home page).
    if parts.scheme == "ms-settings" and not parts.netloc and not parts.path:
        return True
    return bool(parts.netloc) if parts.scheme in ("http", "https") else bool(parts.netloc or parts.path)


def open_url(url):
    if not url_ok(url):
        raise ValueError(f"rejected link (allowed: {', '.join(sorted(_URL_SCHEMES))}; no spaces): {url!r}")
    previous = _foreground_hwnd()
    webbrowser.open(url.strip())
    _wait_foreground_change(previous, 6)


def open_path(path):
    target = Path(os.path.expandvars(os.path.expanduser(path.strip())))
    if not target.exists():
        raise ValueError(f"path does not exist: {target}")
    if target.is_file() and target.suffix.lower() in _RUNNABLE:
        raise ValueError(f"refusing to run {target.suffix} files; use open_app instead")
    previous = _foreground_hwnd()
    os.startfile(str(target))
    _wait_foreground_change(previous, 6)


def window_op(title, op):
    window = _find_window(title)
    if op == "minimize":
        window.minimize()
    elif op == "maximize":
        window.maximize()
    elif op == "restore":
        window.restore()
    elif op == "move_to_main":
        if window.isMaximized:
            window.restore()
        window.moveTo(40, 40)
    else:
        raise ValueError(f"rejected window op: {op}")
    time.sleep(0.3)
    return None


def window_monitor(window, screens=None):
    """1-based monitor number holding the centre of a pygetwindow window."""
    screens = screens or monitors()
    cx, cy = window.left + window.width // 2, window.top + window.height // 2
    for index, (l, t, r, b, _primary) in enumerate(screens, 1):
        if l <= cx < r and t <= cy < b:
            return index
    return 0


def _find_window(title):
    wanted = title.strip().lower()
    matches = [w for w in pygetwindow.getAllWindows() if wanted in w.title.lower() and w.title.strip()]
    if not matches:
        raise ValueError(f"no open window title contains {title!r}")
    return matches[0]


def foreground_title():
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    buffer = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(hwnd, buffer, 512)
    return buffer.value


def focus_window(title):
    window = _find_window(title)
    if window.isMinimized:
        window.restore()
    # Windows only lets a process take the foreground right after a key press.
    pyautogui.press("alt")
    try:
        window.activate()
    except Exception as exc:
        window.minimize()
        window.restore()
    wanted = title.strip().lower()
    deadline = time.time() + 2
    while time.time() < deadline:
        if wanted in foreground_title().lower() and not _search_open():
            return
        time.sleep(0.1)
    raise RuntimeError(f"could not focus {title!r}; active window is {foreground_title()!r}")


def _screen_xy(x, y, shot):
    sx = shot.left + int(round(float(x) * shot.screen_w / shot.sent_w))
    sy = shot.top + int(round(float(y) * shot.screen_h / shot.sent_h))
    return sx, sy


def _type_text(text):
    if own_console():
        raise ValueError("Co-opilot's own window is in front, not the app, so nothing was typed")
    _ready_to_type()
    if own_console():
        raise ValueError("Co-opilot's own window is in front, not the app, so nothing was typed")
    # Key-by-key typing drops letters in busy apps (Notepad lost characters after commas and
    # line breaks). Paste arrives whole. Tabs stay real Tab presses so form fields still advance.
    for index, part in enumerate(text.replace("\r\n", "\n").split("\t")):
        if index:
            pyautogui.press("tab")
        if not part:
            continue
        if len(part) <= 3 and part.isascii() and "\n" not in part:
            pyautogui.write(part, interval=0.02)
        else:
            _paste(part)


def _paste(text):
    try:
        previous = pyperclip.paste()
    except Exception:
        previous = None
    pyperclip.copy(text)
    try:
        pyautogui.hotkey("ctrl", "v")
        # The app reads the clipboard after the keystroke; restoring too soon pastes the old contents.
        time.sleep(0.15 + min(len(text), 4000) / 20000)
    finally:
        if previous is not None:
            pyperclip.copy(previous)


def _searchish(process, title):
    return process in _SEARCH_PROCS or (title or "").strip().lower() in _SEARCH_TITLES


def _foreground_hwnd():
    return ctypes.windll.user32.GetForegroundWindow()


def _process_name(hwnd):
    pid = wintypes.DWORD()
    ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if not pid.value:
        return ""
    handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid.value)
    if not handle:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(260)
        size = wintypes.DWORD(260)
        if not ctypes.windll.kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
            return ""
        return Path(buf.value).name.lower()
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def own_console(title=None, hwnd=None):
    """True for the window this app was launched from, which must never receive typed text."""
    title = foreground_title() if title is None else title
    if re.search(r"app\.py", title or "", re.I):
        return True
    hwnd = _foreground_hwnd() if hwnd is None else hwnd
    if not hwnd:
        return False
    pid = wintypes.DWORD()
    ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value == os.getpid()


def _wait_for_app(previous, name, timeout, on_slow=None):
    """Wait until a new window is in front. Prefer one whose title matches the app name.
    After 5 seconds, on_slow(name) is called once if the app has a window but isn't in front yet."""
    words = _app_words(name)
    deadline = time.time() + timeout
    slow_at = time.time() + 5
    last, hits = None, 0
    while time.time() < deadline:
        if on_slow and time.time() > slow_at:
            slow_at = float("inf")
            if match_window(name, taskbar_windows()) or _starting(name):
                on_slow(name)
        hwnd = _foreground_hwnd()
        title = foreground_title()
        if not hwnd or hwnd == previous or _searchish(_process_name(hwnd), title) or own_console(title, hwnd):
            last, hits = None, 0
            time.sleep(0.1)
            continue
        named = not words or any(word in title.lower() for word in words)
        if hwnd == last:
            hits += 1
            if hits >= (3 if named else 12):
                return True
        else:
            last, hits = hwnd, 0
        time.sleep(0.1)
    return False


def _search_open():
    hwnd = _foreground_hwnd()
    return bool(hwnd) and _searchish(_process_name(hwnd), foreground_title())


def _wait_until(pred, timeout):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if pred():
            return True
        time.sleep(0.08)
    return False


def _wait_not_search(timeout):
    """Return once the foreground window has stayed off Start/Search for a moment."""
    deadline = time.time() + timeout
    last, hits = None, 0
    while time.time() < deadline:
        hwnd = _foreground_hwnd()
        if not hwnd or _searchish(_process_name(hwnd), foreground_title()):
            last, hits = None, 0
            time.sleep(0.1)
            continue
        if hwnd == last:
            hits += 1
            if hits >= 3:
                return
        else:
            last, hits = hwnd, 0
        time.sleep(0.1)
    if _search_open():
        pyautogui.press("escape")
        raise ValueError("Start search still had focus, so nothing was typed")


def _wait_foreground_change(previous, timeout):
    deadline = time.time() + timeout
    while time.time() < deadline:
        hwnd = _foreground_hwnd()
        if hwnd and hwnd != previous and not _search_open():
            _wait_not_search(2)
            return
        time.sleep(0.1)


def _ready_to_type():
    """Don't type into Start search, or into a window that is still taking focus."""
    if _search_open():
        _wait_not_search(5)
        return
    first = _foreground_hwnd()
    time.sleep(0.2)
    if _search_open() or _foreground_hwnd() != first:
        _wait_not_search(4)


def _press(combo):
    if str(combo).strip().lower() != "escape":
        _ready_to_type()
    parts = [_map_key(part) for part in str(combo).split("+")]
    if len(parts) == 1:
        pyautogui.press(parts[0])
    else:
        pyautogui.hotkey(*parts)


def _map_key(part):
    name = part.strip().lower()
    if name in _KEY_NAMES:
        return _KEY_NAMES[name]
    if len(name) == 1 and name.isascii() and name.isprintable():
        return name
    if re.fullmatch(r"f([1-9]|1[0-2])", name):
        return name
    raise ValueError(f"rejected key: {part}")


def _scroll(action, shot):
    x, y = _screen_xy(action["x"], action["y"], shot)
    amount = max(1, min(20, int(action["amount"])))
    direction = action["direction"]
    if direction == "up":
        pyautogui.scroll(amount, x=x, y=y)
    elif direction == "down":
        pyautogui.scroll(-amount, x=x, y=y)
    elif direction == "right":
        pyautogui.hscroll(amount, x=x, y=y)
    elif direction == "left":
        pyautogui.hscroll(-amount, x=x, y=y)
    else:
        raise ValueError(f"rejected scroll direction: {direction}")


def _check():
    assert app_name_ok("notepad")
    assert app_name_ok("Microsoft Edge")
    assert not app_name_ok("notepad & calc")
    assert not app_name_ok("a|b")
    assert not app_name_ok('say "hi"')
    assert not app_name_ok("")
    assert url_ok("https://example.com/a?b=1") and not url_ok("file:///C:/x") and not url_ok("https://a b")
    assert not url_ok("javascript:alert(1)")
    assert url_ok("mailto:a@b.ca?subject=Hi%20there") and url_ok("ms-settings:sound") and not url_ok("mailto:a@b.ca?subject=Hi there")
    assert url_ok("msteams:/l/chat/0/0?users=a@b.ca") and url_ok("ms-settings:") and not url_ok("shell:startup")
    assert _searchish("searchhost.exe", "") and _searchish("notepad.exe", "Search")
    assert not _searchish("notepad.exe", "Untitled - Notepad")
    assert own_console(r"C:\Windows\System32\cmd.exe - py app.py", 0)
    assert not own_console("Untitled - Notepad", 0)
    wins = [(1, "Copilot Assistant", "python.exe"), (2, "Search", "searchhost.exe"), (3, "Program Manager", "explorer.exe"),
            (4, "Mail - Sohel - Outlook", "olk.exe"), (5, "Inbox - Outlook", "outlook.exe"), (6, "Untitled - Notepad", "notepad.exe")]
    assert match_window("Outlook", wins) == 4 and match_window("Microsoft Outlook", wins[4:]) == 5
    assert match_window("Notepad", wins) == 6 and match_window("Excel", wins) is None
    assert match_window("Copilot Assistant", wins) is None and match_window("Search", wins) is None
    assert match_window("File Explorer", wins[:3]) is None
    assert all(len(w) == 3 for w in taskbar_windows())
    # A second monitor left of the main one: screenshot x 0 is screen x -1920.
    fake = Shot(b"", 3840, 1080, 1920, 540, -1920, 0, [(0, 0, 1920, 1080, True), (-1920, 0, 0, 1080, False)])
    assert _screen_xy(0, 0, fake) == (-1920, 0) and _screen_xy(960, 270, fake) == (0, 540), _screen_xy(960, 270, fake)
    assert fake.layout() == "monitor 1 (main): x 960-1919, y 0-539; monitor 2: x 0-959, y 0-539", fake.layout()
    pair = [(0, 0, 1920, 1080, True), (1920, 0, 3840, 1080, False)]
    wins = [
        (1, "Inbox - Outlook", "olk.exe", 40, 40, 900, 700, False),
        (2, "Settings", "systemsettings.exe", 2100, 40, 3000, 800, False),
        (3, "Co-opilot", "python.exe", 40, 40, 500, 700, True),
    ]
    assert choose_monitors(pair, wins, "what time is it") == [pair[0]]
    assert choose_monitors(pair, wins, "turn the volume up", "Settings") == [pair[1]]
    assert choose_monitors(pair, wins, "on monitor 2") == [pair[1]]
    assert choose_monitors(pair, wins, "put it on the second screen") == [pair[1]]
    assert choose_monitors(pair, wins, "read my outlook and the settings window") == pair
    assert choose_monitors(pair[:1], wins, "on monitor 2") == pair[:1]
    focus = focused_control()
    assert focus.startswith("Active window:") and "focused:" in focus and "unavailable" not in focus, focus
    shot = screenshot()
    assert shot.jpeg[:2] == b"\xff\xd8", "screenshot is not a jpeg"
    assert shot.screen_w > 0 and shot.screen_h > 0
    assert max(shot.sent_w, shot.sent_h) <= _MAX_TOTAL
    one = screenshot([monitors()[0]])
    assert len(one.monitors) == 1 and one.jpeg[:2] == b"\xff\xd8"
    assert _screen_xy(0, 0, one) == (one.left, one.top)
    if len(shot.monitors) > 1:
        assert one.screen_w < shot.screen_w or one.screen_h < shot.screen_h
        assert "Left out: monitor" in shot_caption(one)
    print(f"desktop ok {len(shot.monitors)} monitor(s) {shot.screen_w}x{shot.screen_h} -> {shot.sent_w}x{shot.sent_h}")
    print(" ", shot.layout())
    print(" ", one.layout())


if __name__ == "__main__":
    _check()
