"""Copilot Assistant: computer use through work Copilot. Run: python app.py."""

import ctypes
import queue
import re
import threading
import time
import tkinter as tk
import tkinter.font as tkfont
import winreg
from pathlib import Path
from tkinter import ttk

import sv_ttk
from PIL import Image, ImageDraw, ImageTk
import desktop
import loop

APP_NAME = "Copilot Assistant"
SIDE_W = 520
HINT = "Ask or describe a task, e.g. open Notepad and type hello, or why is my PC slow?"
WELCOME = ("What can I help you do on this PC?\n"
           "- **Do things:** open apps, click, type\n"
           "- **Diagnose:** slow PC, disk space, network, failing services, recent errors\n"
           "Press **Enter** to send, **Shift+Enter** for a new line. Move the mouse to a screen corner to stop.")


def app_mark(size=64):
    """Load the shared Copilot logo for the titlebar and chat header."""
    logo = Path(__file__).resolve().parent / "public" / "copilot-logo.png"
    with Image.open(logo) as source:
        image = source.convert("RGBA")
        image.thumbnail((size, size), Image.Resampling.LANCZOS)
    return ImageTk.PhotoImage(image)


class RoundedButton(tk.Button):
    """A native keyboard-accessible button with rounded, stateful artwork."""

    def __init__(self, parent, text, command, width=92, height=36, kind="quiet", dark=False, surface=None):
        self._dark = dark
        self._kind = kind
        self._width = width
        self._height = height
        self._surface = surface or ("#1c1c1c" if dark else "#fafafa")
        self._hover = False
        self._pressed = False
        self._focused = False
        self._images = {state: self._art(state) for state in ("normal", "hover", "pressed", "disabled", "focus")}
        background = self._surface
        super().__init__(parent, text=text, command=command, image=self._images["normal"], compound="center",
                         font=("Segoe UI Semibold", 10), fg="#ffffff" if kind in ("primary", "danger") else
                         ("#f4f4f4" if dark else "#24252b"), bg=background, activebackground=background,
                         activeforeground="#ffffff" if kind in ("primary", "danger") else
                         ("#f4f4f4" if dark else "#24252b"), bd=0, relief="flat", highlightthickness=0,
                         padx=0, pady=0, cursor="hand2", takefocus=True)
        self.bind("<Enter>", lambda _e: self._set_state(hover=True))
        self.bind("<Leave>", lambda _e: self._set_state(hover=False, pressed=False))
        self.bind("<ButtonPress-1>", lambda _e: self._set_state(pressed=True))
        self.bind("<ButtonRelease-1>", lambda _e: self._set_state(pressed=False))
        self.bind("<FocusIn>", lambda _e: self._set_state(focused=True))
        self.bind("<FocusOut>", lambda _e: self._set_state(focused=False))

    def _art(self, state):
        palette = {
            "primary": ("#464feb", "#3942d9", "#3038bd", "#8d92cc"),
            "danger": ("#c42b1c", "#ad261a", "#922117", "#b98680"),
            "ghost": (self._surface, "#30333d", "#252832", self._surface) if self._dark else
                     (self._surface, "#eceef4", "#e2e5ee", self._surface),
            "quiet": ("#30333d", "#3e4350", "#252832", "#484a50") if self._dark else
                     ("#f0f1f5", "#e6e8ef", "#d9dce7", "#ebebee"),
        }[self._kind]
        fill = palette[{"normal": 0, "hover": 1, "pressed": 2, "disabled": 3, "focus": 0}[state]]
        scale = 2
        image = Image.new("RGBA", (self._width * scale, self._height * scale), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        radius = min(13, self._height // 2) * scale
        edge = "#8c93ff" if state == "focus" else ("#565b69" if self._dark else "#d7d9e1")
        if state == "focus":
            draw.rounded_rectangle((1, 1, self._width * scale - 2, self._height * scale - 2),
                                   radius=radius, outline=edge, width=3)
            inset = 5
        else:
            inset = 1
        draw.rounded_rectangle((inset, inset, self._width * scale - inset - 1,
                                self._height * scale - inset - 1), radius=max(2, radius - inset),
                               fill=fill, outline=edge if self._kind == "quiet" else None, width=1)
        return ImageTk.PhotoImage(image.resize((self._width, self._height), Image.Resampling.LANCZOS))

    def _set_state(self, hover=None, pressed=None, focused=None):
        if hover is not None:
            self._hover = hover
        if pressed is not None:
            self._pressed = pressed
        if focused is not None:
            self._focused = focused
        self._refresh()

    def _refresh(self):
        state = "disabled" if str(self["state"]) == "disabled" else (
            "pressed" if self._pressed else "hover" if self._hover else "focus" if self._focused else "normal")
        self.configure(image=self._images[state], cursor="arrow" if state == "disabled" else "hand2")

    def configure(self, cnf=None, **kw):
        result = super().configure(cnf, **kw)
        if "state" in kw and hasattr(self, "_images"):
            self._refresh()
        return result


class MessageBubble(tk.Canvas):
    """Rounded message surface with selectable, read-only text."""

    def __init__(self, parent, runs, *, user, dark, on_scroll):
        self.runs = runs
        self.user = user
        self.surface = "#343944" if dark and user else "#272b33" if dark else "#e9ebf3" if user else "#f0f2f6"
        self.outer = "#1c1c1c" if dark else "#fafafa"
        self.foreground = "#f3f3f3" if dark else "#202124"
        self.font = tkfont.Font(family="Segoe UI", size=11)
        super().__init__(parent, bg=self.outer, bd=0, highlightthickness=0, height=40)
        self.body = tk.Text(self, wrap="word", relief="flat", bd=0, highlightthickness=0,
                            bg=self.surface, fg=self.foreground, insertwidth=0, cursor="xterm",
                            font=self.font, padx=0, pady=0, exportselection=True, height=1)
        self.body.tag_configure("b", font=("Segoe UI Semibold", 11))
        self.body.tag_configure("h", font=("Segoe UI Semibold", 12))
        self.body.tag_configure("code", font=("Cascadia Mono", 10))
        self.body.tag_configure("error", foreground="#ff817f" if dark else "#b3261e")
        self.body.tag_configure("note", foreground="#b3b3b3" if dark else "#656b75")
        for part, tags in runs:
            self.body.insert("end", part, tuple(tag for tag in tags if tag in ("b", "h", "code", "error", "note")))
        self.body.bind("<Key>", self._read_only)
        self.body.bind("<MouseWheel>", on_scroll)
        self.body.bind("<Button-4>", on_scroll)
        self.body.bind("<Button-5>", on_scroll)
        self.bind("<MouseWheel>", on_scroll)
        self.bind("<Button-4>", on_scroll)
        self.bind("<Button-5>", on_scroll)
        self.body.bind("<<Paste>>", lambda _event: "break")
        self.body.bind("<<Cut>>", lambda _event: "break")
        self.body.bind("<Map>", lambda _event: self.after_idle(self._finish_layout))
        self._image = None
        self._last_width = None

    @staticmethod
    def _read_only(event):
        if event.keysym in {"Left", "Right", "Up", "Down", "Home", "End", "Prior", "Next"}:
            return None
        if event.state & 0x4 and event.keysym.lower() in ("c", "a"):
            return None
        return "break"

    def resize(self, width):
        width = max(180, int(width))
        if width == self._last_width:
            return
        self._last_width = width
        self.body.place(x=14, y=11, width=width - 28)
        self.after_idle(self._finish_layout)

    def _finish_layout(self):
        if not self.winfo_exists() or self._last_width is None:
            return
        width = self._last_width
        self.body.update_idletasks()
        try:
            # count() measures line transitions; the last display line is not
            # included, so add one. Measure after the widget has been mapped.
            lines = self.body.count("1.0", "end-1c", "displaylines")[0] + 1
        except (tk.TclError, TypeError):
            lines = len("".join(part for part, _tags in self.runs).splitlines())
        self.body.configure(height=max(1, lines))
        self.body.update_idletasks()
        height = max(40, self.body.winfo_reqheight() + 22)
        self.configure(width=width, height=height)
        self.master.configure(height=height)
        image = Image.new("RGBA", (width * 2, height * 2), (0, 0, 0, 0))
        ImageDraw.Draw(image).rounded_rectangle((1, 1, width * 2 - 2, height * 2 - 2),
                                                radius=28, fill=self.surface)
        self._image = ImageTk.PhotoImage(image.resize((width, height), Image.Resampling.LANCZOS))
        self.delete("surface")
        self.create_image(0, 0, image=self._image, anchor="nw", tags="surface")
        self.tag_lower("surface")


def _dark_mode():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
    except OSError:
        return False


def _dot_colour(status):
    s = status.lower()
    if s.startswith(("error", "stopped")):
        return "#d13438"
    if s.startswith("step") or s == "stopping":
        return "#0078d4"
    if s.startswith(("ready", "done", "chat ready")):
        return "#107c10"
    return "#ca5010"


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


class Overlay:
    """Click-through 'hands off' layer shown while Copilot Assistant drives the PC, with a clickable Stop
    button per monitor. All of it is hidden from screenshots."""

    KEY = "#010203"
    ACCENT = "#3b8eea"
    TINT = 0.05

    def __init__(self, root, on_stop):
        self.root = root
        self.on_stop = on_stop
        self.stops = []
        screens = desktop.monitors()
        vx, vy = min(m[0] for m in screens), min(m[1] for m in screens)
        width, height = max(m[2] for m in screens) - vx, max(m[3] for m in screens) - vy
        # Tk reads "-1920" as "from the right edge"; "+-1920" is a negative position.
        geometry = f"{width}x{height}+{vx}+{vy}"
        self.tint = self._layer(geometry, alpha=self.TINT, bg="#0a2a5e")
        self.chrome = self._layer(geometry, alpha=0.95, bg=self.KEY)
        self.chrome.attributes("-transparentcolor", self.KEY)
        canvas = tk.Canvas(self.chrome, bg=self.KEY, highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        for l, t, r, b, _primary in screens:
            l, t, r, b = l - vx, t - vy, r - vx, b - vy
            for inset, line, colour in ((2, 6, "#1f5fa8"), (5, 3, self.ACCENT), (7, 1, "#9cc7ff")):
                canvas.create_rectangle(l + inset, t + inset, r - inset, b - inset, outline=colour, width=line)
            pill_w, pill_h, top = min(640, r - l - 32), 72, t + 28
            left = l + (r - l - pill_w) // 2
            self._pill(canvas, left, top, left + pill_w, top + pill_h, "#1b1d22")
            canvas.create_oval(left + 22, top + 29, left + 34, top + 41, fill=self.ACCENT, outline="")
            canvas.create_text(left + 48, top + 22, anchor="w", fill="#ffffff",
                               text=f"{APP_NAME} is using your computer", font=("Segoe UI Semibold", 13))
            guidance = ("Please pause your mouse and keyboard for a moment"
                        if pill_w >= 600 else "One moment while I finish this step")
            canvas.create_text(left + 48, top + 49, anchor="w", fill="#a9b4c2",
                               text=guidance,
                               font=("Segoe UI", 9))
            # Separate small window: the overlay itself must stay click-through.
            stop = tk.Toplevel(root, bg="#1b1d22")
            stop.withdraw()
            stop.overrideredirect(True)
            stop.attributes("-topmost", True)
            stop.geometry(f"96x36+{left + pill_w - 110 + vx}+{top + 18 + vy}")
            label = RoundedButton(stop, text="Stop", command=self._stop_clicked,
                                  width=96, height=36, kind="danger", dark=True, surface="#1b1d22")
            label.pack(fill="both", expand=True)
            self.stops.append((stop, label))
        self._shown = False

    def _layer(self, geometry, alpha, bg):
        layer = tk.Toplevel(self.root, bg=bg)
        layer.withdraw()
        layer.overrideredirect(True)
        layer.geometry(geometry)
        layer.attributes("-topmost", True, "-alpha", alpha)
        return layer

    @staticmethod
    def _pill(canvas, x1, y1, x2, y2, colour):
        radius = (y2 - y1) // 2
        canvas.create_oval(x1, y1, x1 + 2 * radius, y2, fill=colour, outline="")
        canvas.create_oval(x2 - 2 * radius, y1, x2, y2, fill=colour, outline="")
        canvas.create_rectangle(x1 + radius, y1, x2 - radius, y2, fill=colour, outline="")

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
            label.configure(text="Stopping")
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
            label.configure(text="Stop")
            stop.deiconify()
            stop.update_idletasks()
            self._seal(stop, click_through=False)
            stop.lift()

    def hide(self):
        self._shown = False
        self.tint.withdraw()
        self.chrome.withdraw()
        for stop, _label in self.stops:
            stop.withdraw()

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(APP_NAME)
        self.app_icon = app_mark()
        self.root.iconphoto(True, self.app_icon)
        self.root.geometry(f"{SIDE_W}x760")
        self.root.minsize(440, 520)
        self.dark = _dark_mode()
        sv_ttk.set_theme("dark" if self.dark else "light")
        self.messages = queue.Queue()
        self.tasks = queue.Queue()
        self.stop = threading.Event()
        self._busy = False
        self._next_task_id = 0
        self._active_task_id = None
        self._hands_on = False
        self._answered = False
        self._last_error = ""
        self._started_at = 0.0
        bg = self._bg()
        self.text_bg, self.text_fg = bg, ("#f3f3f3" if self.dark else "#1b1b1b")
        card = "#333842" if self.dark else "#ffffff"
        reply_bg = "#262a30" if self.dark else "#eef1f5"
        self.muted = "#9a9a9a" if self.dark else "#6b6b6b"
        code_bg = "#363b44" if self.dark else "#dfe4ea"
        accent = "#4ca3ff" if self.dark else "#0063b1"
        self.card, self.accent = card, accent
        self._open_card = None

        side = ttk.Frame(self.root, padding=(18, 16, 18, 14), width=SIDE_W)
        side.pack(side="left", fill="both", expand=True)

        header = ttk.Frame(side)
        header.pack(fill="x")
        self.header_icon = app_mark(32)
        ttk.Label(header, image=self.header_icon).pack(side="left", padx=(0, 10))
        ttk.Label(header, text=APP_NAME, font=("Segoe UI Semibold", 14)).pack(side="left")
        self.pinned = tk.BooleanVar(value=False)
        ttk.Checkbutton(header, text="Pin", variable=self.pinned, command=self._set_pin).pack(side="right", padx=(8, 0))

        status_row = ttk.Frame(side)
        status_row.pack(fill="x", pady=(10, 0))
        RoundedButton(status_row, text="＋  New chat", command=self._new_conversation,
                      width=120, height=36, kind="ghost", dark=self.dark, surface=bg).pack(side="right")
        self.dot = tk.Canvas(status_row, width=12, height=12, highlightthickness=0, bg=bg)
        self.dot.pack(side="left", padx=(0, 8))
        self.dot_id = self.dot.create_oval(2, 2, 11, 11, fill=_dot_colour("starting"), outline="")
        self.status = tk.StringVar(value="Starting")
        ttk.Label(status_row, textvariable=self.status, foreground=self.muted,
                  font=("Segoe UI", 10)).pack(side="left")

        # Composer at the bottom, transcript fills the rest.
        bottom = self.bottom = ttk.Frame(side)
        bottom.pack(side="bottom", fill="x", pady=(10, 0))
        self.task = tk.Text(bottom, height=3, wrap="word", font=("Segoe UI", 11), relief="flat", padx=14, pady=10,
                            highlightthickness=2)
        self.task.pack(fill="x")
        # sv_ttk's tk_setPalette repaints classic widgets after the theme loads, so colour it afterwards too.
        composer = dict(bg=card, fg=self.text_fg, insertbackground=self.text_fg,
                        highlightbackground="#4a6a8f" if self.dark else "#9cbbe0",
                        highlightcolor="#4ca3ff" if self.dark else "#0078d4")
        self.task.configure(**composer)
        self.task.bind("<<ThemeChanged>>", lambda _e: self.task.configure(**composer))
        self.root.after(300, lambda: self.task.configure(**composer))
        self.task.tag_configure("hint", foreground=self.muted)
        self._show_hint()
        self.task.bind("<FocusIn>", lambda _e: self._clear_hint())
        self.task.bind("<FocusOut>", lambda _e: self._show_hint())
        self.task.bind("<Return>", lambda _e: (self._start(), "break")[1])
        self.task.bind("<Shift-Return>", lambda _e: None)

        actions = ttk.Frame(bottom)
        actions.pack(fill="x", pady=(8, 0))
        self.start_btn = RoundedButton(actions, text="Send  ↑", command=self._start,
                                       width=90, kind="primary", dark=self.dark)
        self.start_btn.configure(state="disabled")
        self.start_btn.pack(side="right")
        self.stop_btn = RoundedButton(actions, text="Stop", command=self._on_stop,
                                      width=76, kind="quiet", dark=self.dark)
        self.stop_btn.configure(state="disabled")
        self.stop_btn.pack(side="right", padx=(0, 8))
        self.minimize = tk.BooleanVar(value=True)
        self.hide_toggle = ttk.Checkbutton(actions, text="Hide while working", variable=self.minimize)
        self.hide_toggle.pack(side="left")
        ttk.Label(bottom, text="Enter to send  ·  Shift+Enter for a new line", foreground=self.muted,
                  font=("Segoe UI", 9)).pack(anchor="w", pady=(6, 0))

        body = ttk.Frame(side)
        body.pack(fill="both", expand=True, pady=(12, 0))
        self.chat = tk.Text(body, wrap="word", font=("Segoe UI", 11), relief="flat", insertwidth=0,
                            bg=self.text_bg, fg=self.text_fg, padx=14, pady=12, highlightthickness=0,
                            cursor="xterm", spacing1=2, spacing3=2, exportselection=True,
                            selectbackground="#264f78" if self.dark else "#b3d4fc",
                            inactiveselectbackground="#264f78" if self.dark else "#b3d4fc")
        scroll = ttk.Scrollbar(body, orient="vertical", command=self.chat.yview)
        self.chat.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.chat.pack(side="left", fill="both", expand=True)
        self.chat.bind("<Configure>", self._fit_column)
        self.chat.bind("<Configure>", self._resize_bubbles, add="+")
        self.bubbles = []
        # Created first so h/bullet margins (later tags win) still indent inside the card.
        self.chat.tag_configure("reply", background=reply_bg, lmargin1=14, lmargin2=14, rmargin=14,
                                lmargincolor=reply_bg, rmargincolor=reply_bg)
        self.chat.tag_configure("pad", font=("Segoe UI", 4))
        self.chat.tag_configure("who", font=("Segoe UI Semibold", 9), foreground=self.muted, spacing1=22,
                                spacing3=4)
        self.chat.tag_configure("dot", foreground=accent)
        self.chat.tag_configure("gap", font=("Segoe UI", 8))
        self.chat.tag_configure("user_line", justify="right", lmargin1=120, lmargin2=120, spacing1=0,
                                spacing3=0, font=("Segoe UI", 11))
        self.chat.tag_configure("user", foreground=self.text_fg, font=("Segoe UI", 11))
        self.chat.tag_configure("h", font=("Segoe UI Semibold", 12), spacing1=12, spacing3=2,
                                lmargin1=14, lmargin2=14)
        self.chat.tag_configure("bullet", lmargin1=24, lmargin2=42, spacing1=3)
        self.chat.tag_configure("b", font=("Segoe UI Semibold", 11))
        self.chat.tag_configure("code", font=("Cascadia Mono", 10), background=code_bg)
        self.chat.tag_configure("note", foreground=self.muted, font=("Segoe UI", 10, "italic"))
        self.chat.tag_configure("error", foreground="#ff6b6b" if self.dark else "#c42b1c")
        self.chat.tag_configure("working", foreground=accent, font=("Segoe UI", 10, "italic"))
        # Read-only Text never takes focus on its own, so Ctrl+C would not reach it.
        # Normal state so the mouse can select (a disabled Text will not), with every edit blocked here.
        self.chat.bind("<Button-1>", lambda _e: self.chat.focus_set(), add="+")
        self.chat.bind("<Key>", self._chat_key)
        for event in ("<<Paste>>", "<<Cut>>", "<<Clear>>", "<<PasteSelection>>", "<Button-2>"):
            self.chat.bind(event, lambda _e: "break")
        self.chat.bind("<Control-c>", lambda _e: (self._copy_chat(), "break")[1])
        self.chat.bind("<Control-a>", lambda _e: (self.chat.tag_add("sel", "1.0", "end-1c"), "break")[1])
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Copy", command=self._copy_chat)
        menu.add_command(label="Select all", command=lambda: self.chat.tag_add("sel", "1.0", "end-1c"))
        self.chat.bind("<Button-3>", lambda e: menu.tk_popup(e.x_root, e.y_root))
        self._say_copilot(WELCOME)

        # Developer log: everything the worker reports, hidden until Ctrl+Shift+D.
        self.dev = tk.Toplevel(self.root)
        self.dev.title("Developer log")
        self.dev.geometry("720x480")
        self.dev.withdraw()
        self.dev.protocol("WM_DELETE_WINDOW", self.dev.withdraw)
        self.log = tk.Text(self.dev, wrap="word", font=("Cascadia Mono", 9), relief="flat",
                           bg=self.text_bg, fg=self.text_fg, padx=10, pady=8)
        self.log.pack(fill="both", expand=True)
        self.log.tag_configure("error", foreground="#ff6b6b" if self.dark else "#c42b1c")
        self.root.bind_all("<Control-Shift-D>", lambda _e: self._toggle_dev())
        self.root.bind_all("<Control-Shift-d>", lambda _e: self._toggle_dev())

        self.overlay = Overlay(self.root, self._on_stop)
        self.worker = threading.Thread(target=self._session, daemon=False)
        self.worker.start()
        self.root.after(100, self._drain)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def mainloop(self):
        self.root.mainloop()

    def _bg(self):
        return ttk.Style().lookup("TFrame", "background") or ("#1c1c1c" if self.dark else "#fafafa")

    def _session(self):
        try:
            loop.session(
                self.tasks,
                lambda message: self.messages.put(("log", str(message))),
                lambda message: self.messages.put(("status", str(message))),
                self.stop,
                on_ready=lambda: self.messages.put(("ready", None)),
                on_idle=lambda task_id: self.messages.put(("idle", task_id)),
                answer=lambda task_id, text: self.messages.put(("answer", (task_id, str(text)))),
                say=lambda text: self.messages.put(("say", str(text))),
                ask=lambda payload, reply: self.messages.put(("ask", (payload, reply))),
                hands_on=lambda: self.messages.put(("hands_on", None)),
            )
        finally:
            self.messages.put(("closed", None))

    # ---- transcript ----

    @staticmethod
    def _at_bottom(widget):
        return widget.yview()[1] >= 0.98

    def _write(self, runs):
        follow = self._at_bottom(self.chat)
        top = self.chat.yview()[0]
        for text, tags in runs:
            self.chat.insert("end", text, tags)
        if follow:
            self.chat.see("end")
        else:
            self.chat.yview_moveto(top)

    def _fit_column(self, event):
        """Keep text in a centred ~780px reading column however wide the window is."""
        pad = max(16, (event.width - 780) // 2)
        if int(self.chat.cget("padx")) != pad:
            self.chat.configure(padx=pad)
            self.bottom.pack_configure(padx=(max(0, pad - 16), max(0, pad - 16) + 18))

    def _resize_bubbles(self, _event=None):
        available = max(260, self.chat.winfo_width() - 2 * int(self.chat.cget("padx")) - 24)
        for outer, bubble in self.bubbles:
            if not outer.winfo_exists():
                continue
            width = min(available, max(190, int(available * .76))) if bubble.user else available
            bubble.resize(width)
            outer.configure(width=available, height=bubble.winfo_reqheight())

    def _bubble_wheel(self, event):
        if getattr(event, "num", None) == 4:
            units = -3
        elif getattr(event, "num", None) == 5:
            units = 3
        elif event.delta:
            units = -max(1, round(abs(event.delta) / 40)) if event.delta > 0 else max(1, round(abs(event.delta) / 40))
        else:
            return "break"
        self.chat.yview_scroll(units, "units")
        return "break"

    _NAV_KEYS = {"Left", "Right", "Up", "Down", "Home", "End", "Prior", "Next", "Shift_L", "Shift_R",
                 "Control_L", "Control_R"}

    def _chat_key(self, event):
        """Read-only transcript: allow copy, select-all, and moving or extending the selection."""
        ctrl = event.state & 0x4
        if event.keysym in self._NAV_KEYS or (ctrl and event.keysym.lower() in ("c", "a", "insert")):
            return None
        return "break"

    def _copy_chat(self):
        try:
            text = self.chat.get("sel.first", "sel.last")
        except tk.TclError:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(line.strip() for line in text.replace("\u200b", "").split("\n")).strip())

    def _say_user(self, text):
        self._message_bubble([(text.strip(), ())], user=True)

    def _say_copilot(self, text, tag=None):
        runs = [("● ", ("who", "dot")), (f"{APP_NAME}\n", ("who",))]
        self._write(runs)
        body = [(t, tags + ((tag,) if tag else ())) for t, tags in markdown_runs(text.strip())]
        if body and body[-1][0] == "\n":
            body.pop()
        self._message_bubble(body, user=False)

    def _message_bubble(self, runs, user):
        follow = self._at_bottom(self.chat)
        top = self.chat.yview()[0]
        outer = tk.Frame(self.chat, bg=self.text_bg)
        outer.pack_propagate(False)
        bubble = MessageBubble(outer, runs, user=user, dark=self.dark, on_scroll=self._bubble_wheel)
        outer.bind("<MouseWheel>", self._bubble_wheel)
        outer.bind("<Button-4>", self._bubble_wheel)
        outer.bind("<Button-5>", self._bubble_wheel)
        bubble.pack(side="right" if user else "left")
        self.bubbles.append((outer, bubble))
        self.chat.window_create("end", window=outer, stretch=True, pady=4)
        self.chat.insert("end", "\n")
        self._resize_bubbles()
        if follow:
            self.chat.see("end")
        else:
            self.chat.yview_moveto(top)

    def _working(self, label):
        follow = self._at_bottom(self.chat)
        top = self.chat.yview()[0]
        ranges = self.chat.tag_ranges("working")
        if ranges:
            self.chat.delete(ranges[0], ranges[-1])
        if label:
            self.chat.insert("end", label + "\n", ("working",))
        if follow:
            self.chat.see("end")
        else:
            self.chat.yview_moveto(top)

    def _new_conversation(self):
        if self._busy:
            return
        self.chat.delete("1.0", "end")
        self.bubbles.clear()
        self._say_copilot(WELCOME)
        self.start_btn.configure(state="disabled")
        self._set_status("Starting new chat")
        self.tasks.put(loop.NEW_CHAT)

    # ---- developer log ----

    def _toggle_dev(self):
        if self.dev.state() == "withdrawn":
            self.dev.deiconify()
            self.dev.lift()
        else:
            self.dev.withdraw()

    def _dev_log(self, message):
        is_error = message.startswith(("Error", "Action failed", "No action after"))
        if is_error:
            self._last_error = message
        follow = self._at_bottom(self.log)
        top = self.log.yview()[0]
        self.log.insert("end", time.strftime("%H:%M:%S  ") + message + "\n", ("error",) if is_error else ())
        if follow:
            self.log.see("end")
        else:
            self.log.yview_moveto(top)

    # ---- controls ----

    def _show_hint(self):
        if not self.task.get("1.0", "end").strip():
            self.task.insert("1.0", HINT, "hint")

    def _clear_hint(self):
        if self.task.tag_ranges("hint"):
            self.task.delete("1.0", "end")

    def _task_text(self):
        return "" if self.task.tag_ranges("hint") else self.task.get("1.0", "end").strip()

    def _start(self):
        task = self._task_text()
        if not task or str(self.start_btn["state"]) == "disabled":
            return
        self.task.delete("1.0", "end")
        self.chat.tag_remove("sel", "1.0", "end")
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self._busy, self._answered, self._last_error = True, False, ""
        self._next_task_id += 1
        self._active_task_id = self._next_task_id
        self._started_at = time.time()
        self._say_user(task)
        self._working("Working on it…")
        self._dev_log(f"▶ {task}")
        self._hands_on = False
        self.tasks.put((self._active_task_id, task))

    def _on_stop(self):
        self.stop.set()
        self._set_status("Stopping")

    def _set_status(self, message):
        self.status.set(message)
        self.dot.itemconfigure(self.dot_id, fill=_dot_colour(message))
        if self._busy and message.startswith("Step"):
            self._working(f"Working on it… ({message.lower()})")

    def _set_pin(self):
        self.root.attributes("-topmost", self.pinned.get())
        # A pinned chat should stay visible to the user without becoming part
        # of the screenshots sent back to Copilot.
        try:
            hwnd = int(self.root.wm_frame(), 16)
            ctypes.windll.user32.SetWindowDisplayAffinity(ctypes.c_void_p(hwnd),
                                                            0x11 if self.pinned.get() else 0)
        except (ValueError, tk.TclError):
            pass
        if self.pinned.get():
            self.minimize.set(False)
            self.hide_toggle.configure(state="disabled")
            self.root.deiconify()
            self.root.lift()
        else:
            self.hide_toggle.configure(state="normal")

    def _ask_card(self, payload, reply):
        """Question card in the transcript: radio buttons or checkboxes, an optional details box, Submit / Skip."""
        self.overlay.hide()
        self.root.deiconify()
        self.root.lift()
        self._working(None)
        self._dev_log("Ask: " + payload["question"])
        self.root.update_idletasks()
        room = max(260, self.chat.winfo_width() - 2 * int(self.chat.cget("padx")) - 40)
        wrap = max(220, min(680, room - 48))
        frame = tk.Frame(self.chat, bg=self.card, padx=16, pady=14, highlightthickness=1,
                         highlightbackground=self.accent)
        body = tk.Frame(frame, bg=self.card)
        body.pack(fill="x")
        canvas = tk.Canvas(body, bg=self.card, highlightthickness=0, bd=0, width=room, height=90)
        canvas.pack(side="left", fill="x", expand=True)
        card_scroll = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=card_scroll.set)
        content = tk.Frame(canvas, bg=self.card)
        content_id = canvas.create_window((0, 0), window=content, anchor="nw")
        tk.Label(content, text=payload["question"], bg=self.card, fg=self.text_fg, font=("Segoe UI Semibold", 11),
                 wraplength=wrap, justify="left").pack(anchor="w", pady=(0, 8))
        choice, checks = tk.StringVar(value=""), []
        for option in payload["options"]:
            if payload["multi"]:
                var = tk.BooleanVar(value=False)
                checks.append((option, var))
                widget = tk.Checkbutton(content, text=option, variable=var)
            else:
                widget = tk.Radiobutton(content, text=option, value=option, variable=choice)
            widget.configure(bg=self.card, fg=self.text_fg, selectcolor=self.text_bg, activebackground=self.card,
                             activeforeground=self.text_fg, font=("Segoe UI", 10), wraplength=wrap, justify="left",
                             anchor="w", highlightthickness=0, cursor="hand2")
            widget.pack(anchor="w", fill="x", pady=1)
        details = None
        if payload["allow_text"]:
            tk.Label(content, text="Describe more (optional)" if payload["options"] else "Your answer",
                     bg=self.card, fg=self.muted, font=("Segoe UI", 9)).pack(anchor="w", pady=(8, 2))
            details = tk.Text(content, height=2, width=1, wrap="word", font=("Segoe UI", 10), relief="flat",
                              bg=self.text_bg, fg=self.text_fg, insertbackground=self.text_fg, padx=8, pady=6,
                              highlightthickness=1, highlightbackground="#4a6a8f", highlightcolor=self.accent)
            details.pack(anchor="w", fill="x")
        buttons = ttk.Frame(frame)
        buttons.pack(anchor="e", pady=(10, 0))
        result = tk.Label(frame, bg=self.card, fg=self.muted, font=("Segoe UI", 9, "italic"), wraplength=wrap,
                          justify="left")

        def resize_card(_event=None):
            if not frame.winfo_exists():
                return
            width = max(220, self.chat.winfo_width() - 2 * int(self.chat.cget("padx")) - 42)
            canvas.configure(width=width)
            canvas.itemconfigure(content_id, width=width)
            for child in content.winfo_children():
                if isinstance(child, (tk.Label, tk.Radiobutton, tk.Checkbutton)):
                    child.configure(wraplength=max(180, width - 24))
            content.update_idletasks()
            required = content.winfo_reqheight()
            available = max(88, min(380, int(self.chat.winfo_height() * .55)))
            canvas.configure(height=min(required, available))
            canvas.configure(scrollregion=(0, 0, width, required))
            if required > available:
                card_scroll.pack(side="right", fill="y")
            else:
                card_scroll.pack_forget()

        resize_binding = self.chat.bind("<Configure>", resize_card, add="+")
        frame._body = body
        frame._buttons = buttons
        frame._content = content
        frame._resize_binding = resize_binding

        def close(answer):
            if self._open_card is not frame:
                return
            self._open_card = None
            self.chat.unbind("<Configure>", resize_binding)
            picked = [o for o, v in checks if v.get()] if payload["multi"] else ([choice.get()] if choice.get() else [])
            for child in content.winfo_children():
                if isinstance(child, (tk.Radiobutton, tk.Checkbutton, tk.Text)):
                    child.configure(state="disabled")
            body.pack_forget()
            buttons.pack_forget()
            if answer is None:
                result.configure(text="Skipped")
                reply.put(None)
            else:
                text = details.get("1.0", "end").strip() if details else ""
                result.configure(text="You answered: " + "; ".join(picked + ([text] if text else [])))
                reply.put({"choices": picked, "text": text})
            result.pack(anchor="w", pady=(8, 0))
            self._working("Working on it…")
            if self._hands_on:
                self.overlay.show()
                if self.minimize.get() and not self.pinned.get():
                    self.root.after(300, self.root.iconify)

        def submit():
            picked = any(v.get() for _o, v in checks) or choice.get()
            typed = details is not None and details.get("1.0", "end").strip()
            if payload["options"] and not picked and not typed:
                result.configure(text="Pick an option or describe what you want.")
                result.pack(anchor="w", pady=(8, 0))
                return
            close({})

        RoundedButton(buttons, text="Skip", command=lambda: close(None),
                      width=72, dark=self.dark).pack(side="left", padx=(0, 8))
        RoundedButton(buttons, text="Continue", command=submit, width=96,
                      kind="primary", dark=self.dark).pack(side="left")
        self._bind_card_scroll(frame, canvas)
        self._open_card = frame
        self._write([("● ", ("who", "dot")), (f"{APP_NAME} needs your input\n", ("who",))])
        self.chat.window_create("end", window=frame, stretch=True)
        self.chat.insert("end", "\n")
        resize_card()
        self.chat.see("end")
        if details is not None and not payload["options"]:
            details.focus_set()

    def _bind_card_scroll(self, widget, canvas):
        """Scroll choices when tall; otherwise scroll the transcript."""
        def scroll(units):
            bounds = canvas.bbox("all")
            if bounds and canvas.winfo_height() < bounds[3]:
                before = canvas.yview()
                canvas.yview_scroll(units, "units")
                if canvas.yview() != before:
                    return "break"
            if self.chat.winfo_exists():
                self.chat.yview_scroll(units, "units")
            return "break"

        def wheel(event):
            if not event.delta:
                return "break"
            return scroll(-max(1, round(abs(event.delta) / 40)) if event.delta > 0 else
                          max(1, round(abs(event.delta) / 40)))

        widget.bind("<MouseWheel>", wheel)
        widget.bind("<Button-4>", lambda _event: scroll(-3))
        widget.bind("<Button-5>", lambda _event: scroll(3))
        for child in widget.winfo_children():
            self._bind_card_scroll(child, canvas)

    def _finish_task(self):
        if self._open_card is not None:
            card = self._open_card
            self.chat.unbind("<Configure>", card._resize_binding)
            for child in card._content.winfo_children():
                if isinstance(child, (tk.Radiobutton, tk.Checkbutton, tk.Text)):
                    child.configure(state="disabled")
            card._body.pack_forget()
            card._buttons.pack_forget()
            self._open_card = None
        self._working(None)
        if not self._answered:
            status = self.status.get().lower()
            if status.startswith("stopped"):
                self._say_copilot("Stopped. Nothing else will be done for this request.", "note")
            else:
                reason = self._last_error or f"{APP_NAME} didn't finish this request."
                self._say_copilot(f"Sorry, that didn't work.\n{reason}", "error")
        self._write([(f"{time.time() - self._started_at:.0f}s\n", ("note",))])
        self._busy = False
        self._active_task_id = None

    def _drain(self):
        try:
            while True:
                kind, message = self.messages.get_nowait()
                if kind == "status":
                    self._set_status(message)
                    if message.startswith("Sign in"):
                        self.overlay.hide()
                        self.root.deiconify()
                elif kind == "hands_on":
                    self._hands_on = True
                    self._working("Using the mouse and keyboard…")
                    self.overlay.show()
                    if self.minimize.get() and not self.pinned.get():
                        self.root.iconify()
                elif kind == "log":
                    self._dev_log(message)
                elif kind == "say":
                    self._working(None)
                    self._say_copilot(message)
                    self._working("Working on it…")
                    self._dev_log("Say: " + message)
                elif kind == "ask":
                    self._ask_card(*message)
                elif kind == "answer":
                    task_id, message = message
                    if task_id != self._active_task_id:
                        continue
                    self._answered = True
                    self._working(None)
                    self._say_copilot(message)
                    self._dev_log("Answer: " + message)
                elif kind == "idle":
                    if message != self._active_task_id:
                        continue
                    self.overlay.hide()
                    if self._busy:
                        self._finish_task()
                    self.start_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")
                    self.root.deiconify()
                elif kind == "ready":
                    if not self._busy:
                        self.start_btn.configure(state="normal")
                elif kind == "closed":
                    self.overlay.hide()
                    if self._last_error and self.status.get() == "Error":
                        self._say_copilot(f"{APP_NAME} couldn't start. Close any other copy of this app and reopen it.\n"
                                          f"`{self._last_error[:300]}`", "error")
                    self.start_btn.configure(state="disabled")
                    self.stop_btn.configure(state="disabled")
        except queue.Empty:
            pass
        if self.root.winfo_exists():
            self.root.after(100, self._drain)

    def _on_close(self):
        self.stop.set()
        self.tasks.put(None)
        self._set_status("Closing")
        self._finish_close()

    def _finish_close(self):
        if self.worker.is_alive():
            self.root.after(200, self._finish_close)
            return
        self.root.destroy()


def _check():
    runs = markdown_runs("## Top\n- **msedge** uses `940 MB`\nplain")
    assert runs[0] == ("Top", ("h",)), runs
    assert ("msedge", ("bullet", "b")) in runs and ("940 MB", ("bullet", "code")) in runs, runs
    assert runs[1:3] == [("\n", ("h",)), ("•  ", ("bullet", "dot"))], runs
    assert runs[-2:] == [("plain", ()), ("\n", ())], runs
    inline = "".join(t for t, _ in markdown_runs("Two windows: - 'Intranet - Work - Edge' (8 tabs) - 'Dashboard - Edge'"))
    assert inline == "Two windows:\n•  'Intranet - Work - Edge' (8 tabs)\n•  'Dashboard - Edge'\n", inline
    print("app markdown ok")


if __name__ == "__main__":
    App().mainloop()
