"""Desktop actions across every attached monitor.

ponytail: one screenshot of the whole virtual desktop, so on very wide setups each
monitor gets fewer pixels. Upgrade path is a per-monitor screenshot action.
"""

import ctypes
import os
import re
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

# Long edge of the JPEG sent to Copilot, per monitor side by side, up to _MAX_TOTAL.
_MAX_EDGE = 1280
_MAX_TOTAL = 2560
# Opening these would run a program, which open_app already covers through the Start menu.
_RUNNABLE = {".exe", ".bat", ".cmd", ".com", ".ps1", ".vbs", ".vbe", ".js", ".jse", ".wsf", ".wsh",
             ".msi", ".msp", ".scr", ".lnk", ".pif", ".reg", ".hta", ".cpl", ".url", ".appref-ms"}
_BAD_APP = re.compile(r'[&|;<>^`\n\r"%]')
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
        for index, (l, t, r, b, primary) in enumerate(self.monitors, 1):
            parts.append(f"monitor {index}{' (main)' if primary else ''}: x {round((l - self.left) * sx)}-"
                         f"{round((r - self.left) * sx) - 1}, y {round((t - self.top) * sy)}-{round((b - self.top) * sy) - 1}")
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


def screenshot():
    screens = monitors()
    left, top = min(m[0] for m in screens), min(m[1] for m in screens)
    image = ImageGrab.grab(all_screens=True)
    screen_w, screen_h = image.size
    sent = image
    cap = min(_MAX_TOTAL, _MAX_EDGE * len(screens))
    long_edge = max(image.size)
    if long_edge > cap:
        scale = cap / long_edge
        sent = image.resize(
            (max(1, round(image.width * scale)), max(1, round(image.height * scale))),
            Image.Resampling.LANCZOS,
        )
    buffer = BytesIO()
    sent.convert("RGB").save(buffer, format="JPEG", quality=70)
    return Shot(buffer.getvalue(), screen_w, screen_h, sent.width, sent.height, left, top, screens)


def run(action, shot):
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
        open_app(action["name"])
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


def open_app(name):
    if not app_name_ok(name):
        raise ValueError(f"rejected app name: {name!r}")
    # Start menu search resolves display names ("File Explorer") that `start` cannot.
    pyautogui.press("win")
    time.sleep(0.7)
    _paste(name.strip())
    time.sleep(1.0)
    pyautogui.press("enter")


def url_ok(url):
    if not isinstance(url, str) or re.search(r"\s", url.strip()):
        return False
    parts = urllib.parse.urlsplit(url.strip())
    if parts.scheme not in _URL_SCHEMES:
        return False
    return bool(parts.netloc) if parts.scheme in ("http", "https") else bool(parts.netloc or parts.path)


def open_url(url):
    if not url_ok(url):
        raise ValueError(f"rejected link (allowed: {', '.join(sorted(_URL_SCHEMES))}; no spaces): {url!r}")
    webbrowser.open(url.strip())


def open_path(path):
    target = Path(os.path.expandvars(os.path.expanduser(path.strip())))
    if not target.exists():
        raise ValueError(f"path does not exist: {target}")
    if target.is_file() and target.suffix.lower() in _RUNNABLE:
        raise ValueError(f"refusing to run {target.suffix} files; use open_app instead")
    os.startfile(str(target))


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


def focus_window(title):
    window = _find_window(title)
    if window.isMinimized:
        window.restore()
    # Windows only lets a process take the foreground right after a key press.
    pyautogui.press("alt")
    try:
        window.activate()
    except Exception:
        window.minimize()
        window.restore()
    time.sleep(0.3)


def _screen_xy(x, y, shot):
    sx = shot.left + int(round(float(x) * shot.screen_w / shot.sent_w))
    sy = shot.top + int(round(float(y) * shot.screen_h / shot.sent_h))
    return sx, sy


def _type_text(text):
    if not text.isascii():
        _paste(text)
        return
    lines = text.split("\n")
    for index, line in enumerate(lines):
        if index:
            pyautogui.press("enter")
        parts = line.split("\t")
        for part_index, part in enumerate(parts):
            if part_index:
                pyautogui.press("tab")
            if part:
                pyautogui.write(part, interval=0.01)


def _paste(text):
    try:
        previous = pyperclip.paste()
    except Exception:
        previous = None
    pyperclip.copy(text)
    try:
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.05)
    finally:
        if previous is not None:
            pyperclip.copy(previous)


def _press(combo):
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
    assert url_ok("msteams:/l/chat/0/0?users=a@b.ca") and not url_ok("ms-settings:") and not url_ok("shell:startup")
    # A second monitor left of the main one: screenshot x 0 is screen x -1920.
    fake = Shot(b"", 3840, 1080, 1920, 540, -1920, 0, [(0, 0, 1920, 1080, True), (-1920, 0, 0, 1080, False)])
    assert _screen_xy(0, 0, fake) == (-1920, 0) and _screen_xy(960, 270, fake) == (0, 540), _screen_xy(960, 270, fake)
    assert fake.layout() == "monitor 1 (main): x 960-1919, y 0-539; monitor 2: x 0-959, y 0-539", fake.layout()
    shot = screenshot()
    assert shot.jpeg[:2] == b"\xff\xd8", "screenshot is not a jpeg"
    assert shot.screen_w > 0 and shot.screen_h > 0
    assert max(shot.sent_w, shot.sent_h) <= _MAX_TOTAL
    print(f"desktop ok {len(shot.monitors)} monitor(s) {shot.screen_w}x{shot.screen_h} -> {shot.sent_w}x{shot.sent_h}")
    print(" ", shot.layout())


if __name__ == "__main__":
    _check()
