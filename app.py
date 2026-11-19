"""Co-opilot: computer use through work Copilot. Run: python app.py  (Ctrl+Shift+D opens the developer log)"""

import ctypes
import ctypes.wintypes
import html
import itertools
import json
import os
import queue
import re
import subprocess
import threading
import webbrowser
import time
import tkinter as tk
import winreg
from pathlib import Path

import webview

import account
import commute
import desktop
import files
import loop

APP_NAME = "Copilot Assistant"
SIDE_W = 520
UI = Path(__file__).with_name("ui.html")
LOGO = Path(__file__).resolve().parent / "public" / "assistant-logo-2.png"
# The folder name stays, so a rename of the window does not hide settings already saved.
SETTINGS_FILE = Path(os.environ.get("APPDATA") or Path.home()) / "Co-opilot" / "settings.json"
_HEX = re.compile(r"^#[0-9a-fA-F]{6}$")
CHOICES = {
    "theme": ("system", "light", "dark"),
    "text": ("default", "large", "larger"),
    "hide": (True, False),
    "pinned": (True, False),
    "glance": ("first", "collapse", "stay", "off"),
    "suggestions": (True, False),
    "notify": (True, False),
    "workday_end": tuple(range(14, 22)),
    "refresh": (5, 15, 30),
    "home_mode": ("off", "drive", "train"),
    "go_line": commute.LINES,
}
DEFAULTS = {"theme": "system", "text": "default", "hide": True, "pinned": False, "glance": "first", "suggestions": True,
            "notify": False, "workday_end": 17, "refresh": 5, "accent": "#3c6e71",
            "home_mode": "off", "home_address": "", "go_line": "Kitchener", "go_stop": ""}


def clean_settings(raw):
    """Known keys with allowed values only; anything else falls back to the default."""
    raw = dict(raw) if isinstance(raw, dict) else {}
    if raw.get("glance") is True:
        raw["glance"] = "first"
    elif raw.get("glance") is False:
        raw["glance"] = "off"
    out = {}
    for key, value in DEFAULTS.items():
        got = raw.get(key, value)
        if key == "accent":
            out[key] = got if isinstance(got, str) and _HEX.match(got) else value
        elif key in ("home_address", "go_stop"):
            out[key] = " ".join(got.split())[:80] if isinstance(got, str) else ""
        elif got in CHOICES.get(key, ()) and type(got) is type(value):
            out[key] = got
        else:
            out[key] = value
    # A pinned window must remain visible while the assistant is working.
    if out["pinned"]:
        out["hide"] = False
    return out


def load_settings():
    try:
        return clean_settings(json.loads(SETTINGS_FILE.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        return dict(DEFAULTS)


def save_settings(settings):
    SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
    SETTINGS_FILE.write_text(json.dumps(settings, indent=1), encoding="utf-8")


def _reduce_motion():
    enabled = ctypes.c_int(1)
    ok = ctypes.windll.user32.SystemParametersInfoW(0x1042, 0, ctypes.byref(enabled), 0)
    return bool(ok) and not enabled.value


def _toast(text):
    """One Windows notification. Shows under PowerShell because this app is not an installed package."""
    body = html.escape(_plain(text)[:160] or "Your answer is ready.").replace("'", "&apos;")
    script = (
        "$xml = '<toast><visual><binding template=\"ToastGeneric\"><text>" + html.escape(APP_NAME) + "</text><text>"
        + body + "</text></binding></visual></toast>';"
        "$doc = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]::New();"
        "$doc.LoadXml($xml);"
        "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null;"
        "$id = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\\WindowsPowerShell\\v1.0\\powershell.exe';"
        "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($id).Show("
        "[Windows.UI.Notifications.ToastNotification]::New($doc))"
    )
    subprocess.run(["powershell", "-NoProfile", "-Command", script], capture_output=True, timeout=20)


def _dark_mode():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
    except OSError:
        return False


_INLINE = re.compile(r"(\*\*[^*]+\*\*|`[^`]+`)")


def markdown_runs(text):
    """Split light markdown into (text, tags) runs: ## headings, - / * / 1. bullets, **bold**, `code`."""
    runs = []
    text = text.replace("\r", "")
    # ponytail: Copilot sometimes inlines lists as "...: - a - b". Only " - " after : ) ' " or . starts
    # a bullet, so hyphens inside titles stay; a stray "x. - y" sentence will also split.
    text = re.sub(r"(?<=[:)'\".])\s+-\s+(?=\S)", "\n- ", text)
    for line in text.split("\n"):
        stripped = line.strip()
        line_tags = ()
        if stripped.startswith("#"):
            stripped, line_tags = stripped.lstrip("#").strip(), ("h",)
        elif re.match(r"^[-*•]\s+", stripped):
            stripped, line_tags = re.sub(r"^[-*•]\s+", "", stripped), ("bullet",)
            runs.append(("•  ", ("bullet", "dot")))
        elif re.match(r"^\d+[.)]\s+", stripped):
            line_tags = ("bullet",)
        for part in _INLINE.split(stripped):
            if not part:
                continue
            if part.startswith("**") and part.endswith("**") and len(part) > 4:
                runs.append((part[2:-2], line_tags + ("b",)))
            elif part.startswith("`") and part.endswith("`") and len(part) > 2:
                runs.append((part[1:-1], line_tags + ("code",)))
            else:
                runs.append((part, line_tags))
        runs.append(("\n", line_tags))
    return runs


def _linkify(safe):
    """Turn [label](https://...) and bare http(s) addresses into links. `safe` is already escaped."""
    def linked(label, url):
        return f'<a href="{html.escape(url, quote=True)}">{label}</a>'

    safe = re.sub(r"\[([^\]]+)\]\((https?://[^\s)]+)\)", lambda m: linked(m.group(1), m.group(2)), safe)
    parts = re.split(r"(<a href=.*?</a>)", safe)
    for i, part in enumerate(parts):
        if i % 2:
            continue
        def bare(match):
            url = match.group(0).rstrip(".,)")
            return linked(url, url)

        parts[i] = re.sub(r"https?://[^\s<]+", bare, part)
    return "".join(parts)


def _plain(text):
    text = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", text or "")
    return " ".join(re.sub(r"[*`#]", "", text).split())


def markdown_html(text):
    """Escaped HTML paragraphs for the transcript, built from markdown_runs."""
    out, line = [], []
    for part, tags in markdown_runs((text or "").strip()):
        if part != "\n":
            if "dot" not in tags:
                safe = _linkify(html.escape(part))
                line.append(f"<strong>{safe}</strong>" if "b" in tags else f"<code>{safe}</code>" if "code" in tags else safe)
            continue
        body = "".join(line).strip()
        line = []
        if body:
            cls = "h" if "h" in tags else ("ol" if re.match(r"\d+[.)]\s", body) else "li") if "bullet" in tags else ""
            out.append(f'<p class="{cls}">{body}</p>' if cls else f"<p>{body}</p>")
    return "".join(out)


def describe(action):
    """Plain step label for the transcript, or None for actions that aren't worth a row."""
    kind = action.get("action")
    labels = {
        "open_app": f"Opening {action.get('name', 'an app')}",
        "focus": f"Switching to {action.get('title', 'a window')}",
        "window": f"Arranging {action.get('title', 'a window')}",
        "type": "Typing", "key": f"Pressing {action.get('key', 'a key')}", "scroll": "Scrolling",
        "click": "Clicking", "double_click": "Clicking", "right_click": "Clicking", "drag": "Dragging",
        "open_url": "Opening a web page", "open_path": "Opening a folder", "look": "Checking the screen",
        "inspect": f"Checking {action.get('topic', 'this PC')}", "account": "Reading your work account",
        "ado": "Working in Azure DevOps", "list_dir": "Looking through files", "read_file": "Reading a file",
        "write_file": "Saving a file", "read_clipboard": "Reading the clipboard",
    }
    if kind == "account" and action.get("op") == "recordings":
        return "Finding the meeting chat and transcript"
    return labels.get(kind)


class Overlay:
    """Click-through 'hands off' layer shown while Co-opilot drives the PC, with a clickable Stop
    button per monitor. All of it is hidden from screenshots."""

    KEY = "#010203"
    ACCENT = "#ff453a"
    TINT = 0.10

    def __init__(self, root, on_stop):
        self.root = root
        self.on_stop = on_stop
        self.stops = []
        screens = desktop.monitors()
        vx, vy = min(m[0] for m in screens), min(m[1] for m in screens)
        width, height = max(m[2] for m in screens) - vx, max(m[3] for m in screens) - vy
        # Tk reads "-1920" as "from the right edge"; "+-1920" is a negative position.
        geometry = f"{width}x{height}+{vx}+{vy}"
        self.tint = self._layer(geometry, alpha=self.TINT, bg="#000000")
        self.chrome = self._layer(geometry, alpha=0.96, bg=self.KEY)
        self.chrome.attributes("-transparentcolor", self.KEY)
        canvas = tk.Canvas(self.chrome, bg=self.KEY, highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        for l, t, r, b, _primary in screens:
            l, t, r, b = l - vx, t - vy, r - vx, b - vy
            canvas.create_rectangle(l + 2, t + 2, r - 2, b - 2, outline=self.ACCENT, width=3)
            pill_w, pill_h, top = min(620, r - l - 40), 60, t + 28
            left = l + (r - l - pill_w) // 2
            canvas.create_rectangle(left, top, left + pill_w, top + pill_h, fill="#111111", outline="#2b2b2b")
            canvas.create_rectangle(left + 20, top + 26, left + 27, top + 33, fill=self.ACCENT, outline="")
            canvas.create_text(left + 42, top + 20, anchor="w", fill="#ececec",
                               text=f"{APP_NAME} is using your computer", font=("Segoe UI Semibold", 12))
            canvas.create_text(left + 42, top + 41, anchor="w", fill="#a0a0a0",
                               text="Select Stop at any time to take back control.",
                               font=("Segoe UI", 9))
            # Separate small window: the overlay itself must stay click-through.
            stop = tk.Toplevel(root, bg=self.ACCENT)
            stop.withdraw()
            stop.overrideredirect(True)
            stop.attributes("-topmost", True)
            stop.geometry(f"88x34+{left + pill_w - 104 + vx}+{top + 13 + vy}")
            label = tk.Label(stop, text="■  Stop", bg=self.ACCENT, fg="#160303", activebackground="#e0352b",
                             font=("Segoe UI Semibold", 10), cursor="hand2")
            label.pack(fill="both", expand=True)
            for widget in (stop, label):
                widget.bind("<Button-1>", lambda _e: self._stop_clicked())
            label.bind("<Enter>", lambda _e, w=label: w.configure(bg="#e0352b"))
            label.bind("<Leave>", lambda _e, w=label: w.configure(bg=self.ACCENT))
            self.stops.append((stop, label))
        self._shown = False
        self._fade = 1.0

    def _layer(self, geometry, alpha, bg):
        layer = tk.Toplevel(self.root, bg=bg)
        layer.withdraw()
        layer.overrideredirect(True)
        layer.geometry(geometry)
        layer.attributes("-topmost", True, "-alpha", alpha)
        return layer

    @staticmethod
    def _seal(layer, click_through=True):
        """Never focused, no taskbar button, excluded from screen capture, and optionally click-through."""
        user32 = ctypes.windll.user32
        user32.GetWindowLongPtrW.restype = ctypes.c_ssize_t
        user32.SetWindowLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_ssize_t]
        hwnd = int(layer.wm_frame(), 16)
        ex_style = user32.GetWindowLongPtrW(ctypes.c_void_p(hwnd), -20) | 0x00000080 | 0x08000000
        if click_through:
            ex_style |= 0x00080000 | 0x00000020
        user32.SetWindowLongPtrW(hwnd, -20, ex_style)
        user32.SetWindowDisplayAffinity(ctypes.c_void_p(hwnd), 0x11)

    def _stop_clicked(self):
        for _stop, label in self.stops:
            label.configure(text="Stopping…")
        self.on_stop()

    def show(self):
        if self._shown:
            return
        self._shown = True
        for layer in (self.tint, self.chrome):
            layer.deiconify()
            layer.update_idletasks()
            self._seal(layer)
            layer.lift()
        for stop, label in self.stops:
            label.configure(text="■  Stop")
            stop.deiconify()
            stop.update_idletasks()
            self._seal(stop, click_through=False)
            stop.lift()
        self._fade = 1.0 if _reduce_motion() else 0.0
        self.tint.attributes("-alpha", self.TINT if self._fade else 0.04)
        self._pulse()

    def hide(self):
        self._shown = False
        self.tint.withdraw()
        self.chrome.withdraw()
        for stop, _label in self.stops:
            stop.withdraw()

    def _pulse(self):
        if not self._shown:
            return
        if self._fade < 1:
            self._fade = min(1.0, self._fade + 0.125)
            eased = 1 - (1 - self._fade) ** 3
            self.tint.attributes("-alpha", 0.04 + (self.TINT - 0.04) * eased)
            self.root.after(50, self._pulse)
        else:
            self.tint.attributes("-alpha", self.TINT)


class OverlayThread:
    """Runs the tk overlay on its own thread; the webview owns the main thread."""

    def __init__(self, on_stop):
        self.commands = queue.Queue()
        threading.Thread(target=self._run, args=(on_stop,), daemon=True).start()

    def _run(self, on_stop):
        root = tk.Tk()
        root.withdraw()
        overlay = Overlay(root, on_stop)

        def poll():
            try:
                while True:
                    command = self.commands.get_nowait()
                    if command is None:
                        root.destroy()
                        return
                    getattr(overlay, command)()
            except queue.Empty:
                pass
            root.after(50, poll)

        poll()
        root.mainloop()

    def show(self):
        self.commands.put("show")

    def hide(self):
        self.commands.put("hide")

    def close(self):
        self.commands.put(None)


class Api:
    """What the page may call. Only these methods are exposed to JavaScript."""

    def __init__(self, app):
        self._app = app

    def send(self, text):
        return self._app.start(str(text))

    def stop(self):
        self._app.on_stop()

    def new_chat(self):
        if not self._app.busy:
            self._app.tasks.put(loop.RESET)

    def reply(self, card_id, answer):
        self._app.reply(card_id, answer)

    def get_settings(self):
        return self._app.settings

    def save_settings(self, raw):
        return self._app.apply_settings(raw)

    def refresh_glance(self):
        self._app.glance_now.set()

    def set_glance_live(self, live):
        self._app.set_glance_live(live)

    def file_op(self, op, path):
        """Only files this app wrote in this session; the page can't point these at anything else."""
        path = str(path)
        if path not in self._app._made or not Path(path).exists():
            return "That file is no longer there."
        if op == "open":
            os.startfile(path)
        elif op == "open_with":
            subprocess.Popen(["rundll32.exe", "shell32.dll,OpenAs_RunDLL", path])
        elif op == "reveal":
            subprocess.Popen(f'explorer.exe /select,"{path}"')
        elif op == "copy_file":
            subprocess.run(["powershell", "-NoProfile", "-Command", "Set-Clipboard -LiteralPath $args[0]", path],
                           creationflags=subprocess.CREATE_NO_WINDOW, timeout=15)
        return ""

    def open_link(self, url):
        url = str(url)
        if url.startswith(("https://", "http://")):
            webbrowser.open(url)


class App:
    def __init__(self):
        # Windows reads the app identity when the first native window is created.
        # Versioning it also avoids a stale taskbar icon cached under the old identity.
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Copilot.Assistant.v3")
        self.messages = queue.Queue()
        self.tasks = queue.Queue()
        self.stop = threading.Event()
        self.replies = {}
        self._ids = itertools.count(1)
        self.busy = False
        self.settings = load_settings()
        self.glance_now = threading.Event()
        self._glance_live = True
        self._glance_force = False
        self._hands_on = False
        self._parked = False
        self._answered = False
        self._made = set()
        self._last_error = ""
        self._last_answer = ""
        self._status = ""
        self._started_at = 0.0
        self._loaded = False
        self.overlay = OverlayThread(self.on_stop)
        self.window = webview.create_window(
            APP_NAME, UI.as_uri(), js_api=Api(self), width=SIDE_W, height=820, min_size=(420, 560),
            background_color="#0e0e0e" if self._dark() else "#f3f3f1")
        self.window.events.loaded += self._on_loaded
        self.window.events.shown += self._brand_icon
        self.window.events.closing += self._on_close

    def run(self):
        webview.start()

    @property
    def hide(self):
        return self.settings["hide"]

    def _dark(self):
        theme = self.settings["theme"]
        return _dark_mode() if theme == "system" else theme == "dark"

    def apply_settings(self, raw):
        old = self.settings
        self.settings = clean_settings(raw)
        try:
            save_settings(self.settings)
        except OSError as exc:
            self._dev_log(f"Error: settings not saved: {exc}")
        if old.get("pinned") != self.settings["pinned"]:
            self._set_pinned(self.settings["pinned"])
        watched = ("workday_end", "glance", "refresh", "home_mode", "home_address", "go_line", "go_stop")
        if tuple(old[k] for k in watched) != tuple(self.settings[k] for k in watched):
            self._glance_force = True
            self.glance_now.set()
        return self.settings

    def _js(self, fn, *args):
        self.window.run_js(f"ui.{fn}({', '.join(json.dumps(a) for a in args)})")

    def _on_loaded(self):
        if self._loaded:
            return
        self._loaded = True
        self._brand_icon()
        self._set_pinned(self.settings["pinned"])
        threading.Thread(target=self._dispatch, daemon=True).start()
        threading.Thread(target=self._glance, daemon=True).start()
        self.worker = threading.Thread(target=self._session, daemon=False)
        self.worker.start()

    def _session(self):
        try:
            loop.session(
                self.tasks,
                lambda message: self.messages.put(("log", str(message))),
                lambda message: self.messages.put(("status", str(message))),
                self.stop,
                on_ready=lambda: self.messages.put(("idle", None)),
                on_idle=lambda: self.messages.put(("idle", None)),
                answer=lambda text: self.messages.put(("answer", str(text))),
                say=lambda text: self.messages.put(("say", str(text))),
                ask=lambda payload, reply: self.messages.put(("ask", (payload, reply))),
                hands_on=lambda: self.messages.put(("hands_on", None)),
                on_end=lambda: self.messages.put(("end", None)),
                on_files=lambda paths: self.messages.put(("files", paths)),
            )
        finally:
            self.messages.put(("closed", None))

    def set_glance_live(self, live):
        """The page says whether the strip is actually on screen. A rising edge refreshes at once."""
        live = bool(live)
        if live and not self._glance_live:
            self._glance_force = True
            self.glance_now.set()
        self._glance_live = live

    def _glance(self):
        """Read Outlook only while the strip is visible and this window is in front."""
        last, shown = 0.0, False
        while True:
            self.glance_now.wait(1.5)
            self.glance_now.clear()
            force, self._glance_force = self._glance_force, False
            visible = self.settings["glance"] != "off" and self._glance_live and self._focused()
            if not visible:
                shown = False
                continue
            if shown and not force and time.time() - last < self.settings["refresh"] * 60:
                continue
            try:
                data = account.glance(end_hour=self.settings["workday_end"])
                home = commute.card(self.settings)
                if home:
                    data["home"] = home
                self._js("glance", data)
                last = time.time()
            except Exception as exc:
                self._dev_log(f"Glance failed: {exc}")
            shown = True

    # ---- page calls ----

    def start(self, task):
        task = task.strip()
        if not task or self.busy:
            return False
        self.busy, self._answered, self._last_error, self._hands_on = True, False, "", False
        self._last_answer = ""
        self._started_at = time.time()
        self._dev_log(f"▶ {task}")
        self.tasks.put(task)
        return True

    def on_stop(self):
        self.stop.set()
        self.messages.put(("status", "Stopping"))

    def reply(self, card_id, answer):
        reply = self.replies.pop(int(card_id), None)
        if reply is None:
            return
        if not isinstance(answer, dict):
            reply.put(None)
            return
        choices = [str(c) for c in answer.get("choices") or []]
        reply.put({"choices": choices, "text": str(answer.get("text") or "").strip()})
        if self._hands_on:
            self.overlay.show()
            time.sleep(0.3)
            self._park()

    # ---- worker messages ----

    def _dev_log(self, message):
        is_error = message.startswith(("Error", "Action failed", "No action after"))
        if is_error:
            self._last_error = message
        self._js("log", time.strftime("%H:%M:%S  ") + message, is_error)

    def _brand_icon(self):
        """Title-bar and taskbar icon from the Copilot logo. Handles are kept so Windows does not drop them."""
        hwnd = self._hwnd()
        if not hwnd or not LOGO.is_file():
            return
        try:
            from PIL import Image
            # A versioned filename prevents Windows from reusing the previous logo's cached icon.
            ico = Path(os.environ.get("TEMP") or ".") / "CopilotAssistant-v3.ico"
            if not ico.exists() or ico.stat().st_mtime < LOGO.stat().st_mtime:
                Image.open(LOGO).save(ico, format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (256, 256)])
        except OSError:
            return
        user32 = ctypes.windll.user32
        big = user32.LoadImageW(None, str(ico), 1, 32, 32, 0x10)
        small = user32.LoadImageW(None, str(ico), 1, 16, 16, 0x10)
        if big:
            user32.SendMessageW(hwnd, 0x80, 1, big)
            self._icon_big = big
        if small:
            user32.SendMessageW(hwnd, 0x80, 0, small)
            self._icon_small = small

    def _set_pinned(self, pinned):
        """Keep the main window above other apps without activating it."""
        hwnd = self._hwnd()
        if hwnd:
            insert_after = ctypes.c_void_p(-1 if pinned else -2)  # HWND_TOPMOST / HWND_NOTOPMOST
            ctypes.windll.user32.SetWindowPos(hwnd, insert_after, 0, 0, 0, 0, 0x0013)

    def _hwnd(self):
        return ctypes.windll.user32.FindWindowW(None, APP_NAME)

    def _park(self):
        """Minimise while working, remembering the size the user chose."""
        if not self.hide or self._parked:
            return
        user32 = ctypes.windll.user32
        hwnd = self._hwnd()
        self._was_maximized = bool(hwnd and user32.IsZoomed(hwnd))
        if hwnd and not self._was_maximized:
            rect = ctypes.wintypes.RECT()
            user32.GetWindowRect(hwnd, ctypes.byref(rect))
            self._bounds = (rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)
        self._parked = True
        self.window.minimize()

    def _bring_back(self):
        """Show the window again. Only touch it if we minimised it, and put back the exact size."""
        self.overlay.hide()
        if not self._parked:
            return
        self._parked = False
        if self._was_maximized:
            self.window.maximize()
            return
        self.window.restore()
        hwnd, bounds = self._hwnd(), getattr(self, "_bounds", None)
        if hwnd and bounds:
            ctypes.windll.user32.SetWindowPos(hwnd, None, *bounds, 0x0014)

    def _dispatch(self):
        while True:
            kind, message = self.messages.get()
            try:
                self._handle(kind, message)
            except Exception as exc:
                print(f"UI update failed ({kind}): {exc}")

    def _handle(self, kind, message):
        if kind == "status":
            self._status = message
            self._js("status", message)
            if message.startswith("Sign in"):
                self._bring_back()
        elif kind == "hands_on":
            self._hands_on = True
            self._js("step", "Using the mouse and keyboard")
            self.overlay.show()
            self._park()
        elif kind == "log":
            self._dev_log(message)
            if message.startswith("Action: "):
                try:
                    label = describe(json.loads(message[8:]))
                except ValueError:
                    label = None
                if label:
                    self._js("step", label)
        elif kind == "say":
            self._js("say", markdown_html(message))
            self._dev_log("Say: " + message)
        elif kind == "ask":
            payload, reply = message
            card_id = next(self._ids)
            self.replies[card_id] = reply
            self._bring_back()
            self._dev_log("Ask: " + payload["question"])
            self._js("ask", card_id, payload)
        elif kind == "end":
            self._answered = True
            self._js("end")
        elif kind == "answer":
            self._answered = True
            self._last_answer = message
            self._js("answer", markdown_html(message))
            self._dev_log("Answer: " + message)
        elif kind == "files":
            self._made.update(message)
            self._js("files", [{"path": p, "name": Path(p).name, "folder": str(Path(p).parent),
                                "size": files.size_text(p)} for p in message])
        elif kind == "idle":
            away = not self._focused()
            self._bring_back()
            if self.busy:
                self._finish_task(away)
            for reply in self.replies.values():
                reply.put(None)
            self.replies.clear()
            self._js("idle")
        elif kind == "closed":
            self.overlay.hide()
            if self._last_error and self._status == "Error":
                self._js("error", markdown_html(f"{APP_NAME} couldn't start. Close any other copy of this app and "
                                                f"reopen it.\n`{self._last_error[:300]}`"))
            self._js("closed")

    def _focused(self):
        hwnd = self._hwnd()
        return bool(hwnd and ctypes.windll.user32.GetForegroundWindow() == hwnd)

    def _finish_task(self, away=False):
        if away and self.settings["notify"] and self._answered and self._last_answer:
            threading.Thread(target=_toast, args=(self._last_answer,), daemon=True).start()
        if not self._answered:
            if self._status.lower().startswith("stopped"):
                self._js("note", markdown_html("Stopped. Nothing else will be done for this request."))
            else:
                reason = self._last_error or f"{APP_NAME} didn't finish this request."
                self._js("error", markdown_html(f"That didn't work.\n{reason}"))
        self._js("took", round(time.time() - self._started_at))
        self.busy = False

    def _on_close(self):
        self.stop.set()
        self.tasks.put(None)
        self.overlay.close()


def _check():
    runs = markdown_runs("## Top\n- **msedge** uses `940 MB`\nplain")
    assert runs[0] == ("Top", ("h",)), runs
    assert ("msedge", ("bullet", "b")) in runs and ("940 MB", ("bullet", "code")) in runs, runs
    assert runs[1:3] == [("\n", ("h",)), ("•  ", ("bullet", "dot"))], runs
    assert runs[-2:] == [("plain", ()), ("\n", ())], runs
    inline = "".join(t for t, _ in markdown_runs("Two windows: - 'Intranet - Work - Edge' (8 tabs) - 'Dashboard - Edge'"))
    assert inline == "Two windows:\n•  'Intranet - Work - Edge' (8 tabs)\n•  'Dashboard - Edge'\n", inline
    page = markdown_html("## Top\n- **msedge** uses `940 MB`\n1. first\n<b>x</b>")
    assert page == ('<p class="h">Top</p><p class="li"><strong>msedge</strong> uses <code>940 MB</code></p>'
                    '<p class="ol">1. first</p><p>&lt;b&gt;x&lt;/b&gt;</p>'), page
    assert describe({"action": "open_app", "name": "Notepad"}) == "Opening Notepad"
    assert describe({"action": "done", "summary": "x"}) is None
    assert clean_settings({"theme": "dark", "hide": "yes", "workday_end": 99, "refresh": True, "x": 1}) == dict(
        DEFAULTS, theme="dark")
    assert clean_settings(None) == DEFAULTS
    assert clean_settings({})["notify"] is False
    assert clean_settings({"glance": True})["glance"] == "first"
    assert clean_settings({"glance": False})["glance"] == "off"
    assert clean_settings({"accent": "red", "home_address": "  1 Main  St \n"})["accent"] == "#3c6e71"
    assert clean_settings({"home_address": "  1 Main  St "})["home_address"] == "1 Main St"
    linked = markdown_html("The next train is 3:34 PM.\nSource: [GO Transit](https://www.gotransit.com/en).")
    assert 'href="https://www.gotransit.com/en"' in linked and ">GO Transit</a>" in linked, linked
    assert _plain("**3:34** from [GO Transit](https://www.gotransit.com).") == "3:34 from GO Transit."
    print("app ok")


if __name__ == "__main__":
    App().run()
