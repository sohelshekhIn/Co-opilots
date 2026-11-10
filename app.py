"""Co-opilot: computer use through work Copilot. Run: python app.py  (Ctrl+Shift+D opens the developer log)"""

import ctypes
import queue
import re
import threading
import time
import tkinter as tk
import winreg
from tkinter import ttk

import sv_ttk
import desktop
import loop

APP_NAME = "Co-opilot"
SIDE_W = 520
HINT = "Ask or describe a task, e.g. open Notepad and type hello, or why is my PC slow?"
WELCOME = ("Hi! I can use this PC for you or check what's going on with it.\n"
           "- **Do things:** open apps, click, type\n"
           "- **Diagnose:** slow PC, disk space, network, failing services, recent errors\n"
           "Press **Enter** to send, **Shift+Enter** for a new line. Move the mouse to a screen corner to stop.")


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
    """Click-through 'hands off' layer shown while Co-opilot drives the PC, with a clickable Stop
    button per monitor. All of it is hidden from screenshots."""

    KEY = "#010203"
    ACCENT = "#3b8eea"
    TINT = 0.08

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
            pill_w, pill_h, top = 600, 64, t + 28
            left = l + (r - l - pill_w) // 2
            self._pill(canvas, left, top, left + pill_w, top + pill_h, "#1b1d22")
            canvas.create_oval(left + 22, top + 26, left + 34, top + 38, fill=self.ACCENT, outline="")
            canvas.create_text(left + 48, top + 22, anchor="w", fill="#ffffff",
                               text=f"{APP_NAME} is using your computer", font=("Segoe UI Semibold", 13))
            canvas.create_text(left + 48, top + 44, anchor="w", fill="#a9b4c2",
                               text="Please don't touch the mouse or keyboard · press Stop to take over",
                               font=("Segoe UI", 9))
            # Separate small window: the overlay itself must stay click-through.
            stop = tk.Toplevel(root, bg="#c42b1c")
            stop.withdraw()
            stop.overrideredirect(True)
            stop.attributes("-topmost", True)
            stop.geometry(f"96x36+{left + pill_w - 116 + vx}+{top + 14 + vy}")
            label = tk.Label(stop, text="■  Stop", bg="#c42b1c", fg="#ffffff", activebackground="#a4262c",
                             font=("Segoe UI Semibold", 11), cursor="hand2")
            label.pack(fill="both", expand=True)
            for widget in (stop, label):
                widget.bind("<Button-1>", lambda _e: self._stop_clicked())
            label.bind("<Enter>", lambda _e, w=label: w.configure(bg="#a4262c"))
            label.bind("<Leave>", lambda _e, w=label: w.configure(bg="#c42b1c"))
            self.stops.append((stop, label))
        self._shown = False
        self._phase = 0

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
        self._phase = (self._phase + 1) % 80
        wave = abs(40 - self._phase) / 40
        self.tint.attributes("-alpha", self.TINT * (0.6 + 0.8 * wave))
        self.tint.attributes("-topmost", True)
        self.chrome.attributes("-topmost", True)
        for stop, _label in self.stops:
            stop.attributes("-topmost", True)
            stop.lift()
        self.root.after(50, self._pulse)


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(APP_NAME)
        self.root.geometry(f"{SIDE_W}x760")
        self.root.minsize(400, 520)
        self.dark = _dark_mode()
        sv_ttk.set_theme("dark" if self.dark else "light")
        self.messages = queue.Queue()
        self.tasks = queue.Queue()
        self.stop = threading.Event()
        self._busy = False
        self._answered = False
        self._last_error = ""
        self._started_at = 0.0
        bg = self._bg()
        self.text_bg, self.text_fg = bg, ("#f3f3f3" if self.dark else "#1b1b1b")
        card = "#333842" if self.dark else "#ffffff"
        reply_bg = "#262a30" if self.dark else "#eef1f5"
        self.muted = "#9a9a9a" if self.dark else "#6b6b6b"
        bubble = "#0f6cbd"
        code_bg = "#363b44" if self.dark else "#dfe4ea"
        accent = "#4ca3ff" if self.dark else "#0063b1"
        self.card, self.accent = card, accent
        self._open_card = None

        side = ttk.Frame(self.root, padding=(16, 14, 16, 12), width=SIDE_W)
        side.pack(side="left", fill="both", expand=True)

        header = ttk.Frame(side)
        header.pack(fill="x")
        self.dot = tk.Canvas(header, width=12, height=12, highlightthickness=0, bg=bg)
        self.dot.pack(side="left", padx=(0, 8))
        self.dot_id = self.dot.create_oval(2, 2, 11, 11, fill=_dot_colour("starting"), outline="")
        self.status = tk.StringVar(value="Starting")
        ttk.Label(header, textvariable=self.status, font=("Segoe UI Semibold", 11)).pack(side="left")
        ttk.Button(header, text="New chat", command=self._new_conversation).pack(side="right")

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
        self.start_btn = ttk.Button(actions, text="Send", style="Accent.TButton", command=self._start, state="disabled")
        self.start_btn.pack(side="right")
        self.stop_btn = ttk.Button(actions, text="Stop", command=self._on_stop, state="disabled")
        self.stop_btn.pack(side="right", padx=(0, 8))
        self.minimize = tk.BooleanVar(value=True)
        ttk.Checkbutton(actions, text="Hide while working", variable=self.minimize).pack(side="left")
        ttk.Label(actions, text="Enter to send · Shift+Enter for a new line", foreground=self.muted,
                  font=("Segoe UI", 9)).pack(side="right", padx=(0, 14))

        body = ttk.Frame(side)
        body.pack(fill="both", expand=True, pady=(12, 0))
        self.chat = tk.Text(body, wrap="word", state="disabled", font=("Segoe UI", 11), relief="flat",
                            bg=self.text_bg, fg=self.text_fg, padx=14, pady=12, highlightthickness=0,
                            cursor="xterm", spacing1=2, spacing3=2, exportselection=True,
                            selectbackground="#264f78" if self.dark else "#b3d4fc",
                            inactiveselectbackground="#264f78" if self.dark else "#b3d4fc")
        scroll = ttk.Scrollbar(body, orient="vertical", command=self.chat.yview)
        self.chat.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.chat.pack(side="left", fill="both", expand=True)
        self.chat.bind("<Configure>", self._fit_column)
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
        self.chat.tag_configure("user", background=bubble, foreground="#ffffff", font=("Segoe UI", 11))
        self.chat.tag_configure("h", font=("Segoe UI Semibold", 12), spacing1=12, spacing3=2,
                                lmargin1=14, lmargin2=14)
        self.chat.tag_configure("bullet", lmargin1=24, lmargin2=42, spacing1=3)
        self.chat.tag_configure("b", font=("Segoe UI Semibold", 11))
        self.chat.tag_configure("code", font=("Cascadia Mono", 10), background=code_bg)
        self.chat.tag_configure("note", foreground=self.muted, font=("Segoe UI", 10, "italic"))
        self.chat.tag_configure("error", foreground="#ff6b6b" if self.dark else "#c42b1c")
        self.chat.tag_configure("working", foreground=accent, font=("Segoe UI", 10, "italic"))
        # Read-only Text never takes focus on its own, so Ctrl+C would not reach it.
        self.chat.bind("<Button-1>", lambda _e: self.chat.focus_set(), add="+")
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
                on_ready=lambda: self.messages.put(("idle", None)),
                on_idle=lambda: self.messages.put(("idle", None)),
                answer=lambda text: self.messages.put(("answer", str(text))),
                say=lambda text: self.messages.put(("say", str(text))),
                ask=lambda payload, reply: self.messages.put(("ask", (payload, reply))),
            )
        finally:
            self.messages.put(("closed", None))

    # ---- transcript ----

    def _write(self, runs):
        self.chat.configure(state="normal")
        for text, tags in runs:
            self.chat.insert("end", text, tags)
        self.chat.configure(state="disabled")
        self.chat.see("end")

    def _fit_column(self, event):
        """Keep text in a centred ~780px reading column however wide the window is."""
        pad = max(16, (event.width - 780) // 2)
        if int(self.chat.cget("padx")) != pad:
            self.chat.configure(padx=pad)
            self.bottom.pack_configure(padx=(max(0, pad - 16), max(0, pad - 16) + 18))

    def _copy_chat(self):
        try:
            text = self.chat.get("sel.first", "sel.last")
        except tk.TclError:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(line.strip() for line in text.replace("\u200b", "").split("\n")))

    def _say_user(self, text):
        # Tk paints a tag's background out to the margin when it covers a line's first character
        # and through line spacing, so each line starts with an untagged zero-width space and the gap
        # above is its own small line.
        runs = [("\n", ("gap",))]
        for line in text.strip().split("\n"):
            runs += [("\u200b", ("user_line",)), (f"  {line}  ", ("user_line", "user")), ("\n", ("user_line",))]
        self._write(runs)

    def _say_copilot(self, text, tag=None):
        runs = [("● ", ("who", "dot")), (f"{APP_NAME}\n", ("who",))]
        body = [(t, ("reply",) + tags + ((tag,) if tag else ())) for t, tags in markdown_runs(text.strip())]
        self._write(runs + [("\n", ("reply", "pad"))] + body + [("\n", ("reply", "pad"))])

    def _working(self, label):
        self.chat.configure(state="normal")
        ranges = self.chat.tag_ranges("working")
        if ranges:
            self.chat.delete(ranges[0], ranges[-1])
        if label:
            self.chat.insert("end", label + "\n", ("working",))
        self.chat.configure(state="disabled")
        self.chat.see("end")

    def _new_conversation(self):
        if self._busy:
            return
        self.chat.configure(state="normal")
        self.chat.delete("1.0", "end")
        self.chat.configure(state="disabled")
        self._say_copilot(WELCOME)

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
        self.log.insert("end", time.strftime("%H:%M:%S  ") + message + "\n", ("error",) if is_error else ())
        self.log.see("end")

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
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self._busy, self._answered, self._last_error = True, False, ""
        self._started_at = time.time()
        self._say_user(task)
        self._working("Working on it…")
        self._dev_log(f"▶ {task}")
        self.tasks.put(task)
        if self.minimize.get():
            self.root.after(150, self.root.iconify)

    def _on_stop(self):
        self.stop.set()
        self._set_status("Stopping")

    def _set_status(self, message):
        self.status.set(message)
        self.dot.itemconfigure(self.dot_id, fill=_dot_colour(message))
        if self._busy and message.startswith("Step"):
            self._working(f"Working on it… ({message.lower()})")

    def _ask_card(self, payload, reply):
        """Question card in the transcript: radio buttons or checkboxes, an optional details box, Submit / Skip."""
        self.overlay.hide()
        self.root.deiconify()
        self.root.lift()
        self._working(None)
        self._dev_log("Ask: " + payload["question"])
        self.root.update_idletasks()
        room = self.chat.winfo_width() - 2 * int(self.chat.cget("padx")) - 40
        wrap = min(740, room) if room > 300 else 640
        frame = tk.Frame(self.chat, bg=self.card, padx=16, pady=14, highlightthickness=1,
                         highlightbackground=self.accent)
        tk.Label(frame, text=payload["question"], bg=self.card, fg=self.text_fg, font=("Segoe UI Semibold", 11),
                 wraplength=wrap, justify="left").pack(anchor="w", pady=(0, 8))
        choice, checks = tk.StringVar(value=""), []
        for option in payload["options"]:
            if payload["multi"]:
                var = tk.BooleanVar(value=False)
                checks.append((option, var))
                widget = tk.Checkbutton(frame, text=option, variable=var)
            else:
                widget = tk.Radiobutton(frame, text=option, value=option, variable=choice)
            widget.configure(bg=self.card, fg=self.text_fg, selectcolor=self.text_bg, activebackground=self.card,
                             activeforeground=self.text_fg, font=("Segoe UI", 10), wraplength=wrap, justify="left",
                             anchor="w", highlightthickness=0, cursor="hand2")
            widget.pack(anchor="w", fill="x", pady=1)
        details = None
        if payload["allow_text"]:
            tk.Label(frame, text="Describe more (optional)" if payload["options"] else "Your answer",
                     bg=self.card, fg=self.muted, font=("Segoe UI", 9)).pack(anchor="w", pady=(8, 2))
            details = tk.Text(frame, height=2, width=60, wrap="word", font=("Segoe UI", 10), relief="flat",
                              bg=self.text_bg, fg=self.text_fg, insertbackground=self.text_fg, padx=8, pady=6,
                              highlightthickness=1, highlightbackground="#4a6a8f", highlightcolor=self.accent)
            details.pack(anchor="w", fill="x")
        buttons = ttk.Frame(frame)
        buttons.pack(anchor="e", pady=(10, 0))
        result = tk.Label(frame, bg=self.card, fg=self.muted, font=("Segoe UI", 9, "italic"), wraplength=wrap,
                          justify="left")

        def close(answer):
            if self._open_card is not frame:
                return
            self._open_card = None
            picked = [o for o, v in checks if v.get()] if payload["multi"] else ([choice.get()] if choice.get() else [])
            for child in frame.winfo_children():
                if isinstance(child, (tk.Radiobutton, tk.Checkbutton, tk.Text)):
                    child.configure(state="disabled")
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
            if self.minimize.get():
                self.root.after(300, self.root.iconify)

        def submit():
            picked = any(v.get() for _o, v in checks) or choice.get()
            typed = details is not None and details.get("1.0", "end").strip()
            if payload["options"] and not picked and not typed:
                result.configure(text="Pick an option or describe what you want.")
                result.pack(anchor="w", pady=(8, 0))
                return
            close({})

        ttk.Button(buttons, text="Skip", command=lambda: close(None)).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="Submit", style="Accent.TButton", command=submit).pack(side="left")
        self._open_card = frame
        self._write([("● ", ("who", "dot")), (f"{APP_NAME} needs your input\n", ("who",))])
        self.chat.configure(state="normal")
        self.chat.window_create("end", window=frame)
        self.chat.insert("end", "\n")
        self.chat.configure(state="disabled")
        self.chat.see("end")
        if details is not None and not payload["options"]:
            details.focus_set()

    def _finish_task(self):
        if self._open_card is not None:
            for child in self._open_card.winfo_children():
                if isinstance(child, (tk.Radiobutton, tk.Checkbutton, tk.Text)):
                    child.configure(state="disabled")
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

    def _drain(self):
        try:
            while True:
                kind, message = self.messages.get_nowait()
                if kind == "status":
                    self._set_status(message)
                    if message.startswith("Sign in"):
                        self.overlay.hide()
                        self.root.deiconify()
                    elif message.startswith("Step"):
                        self.overlay.show()
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
                    self._answered = True
                    self._working(None)
                    self._say_copilot(message)
                    self._dev_log("Answer: " + message)
                elif kind == "idle":
                    self.overlay.hide()
                    if self._busy:
                        self._finish_task()
                    self.start_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")
                    self.root.deiconify()
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
