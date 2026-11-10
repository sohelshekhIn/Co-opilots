"""Hidden browser window for work Copilot chat.

Managed Edge on this PC disallows remote debugging, so Playwright cannot attach
to it. This uses Playwright's Chromium instead. Sign-in is still the normal
Microsoft page. Selectors live in this file. Microsoft changes them. Role names
are the first try, then a file input. The upgrade path is to adjust the name
patterns below.
"""

import base64
import ctypes
import os
import re
import subprocess
import tempfile
import threading
import time
from ctypes import wintypes
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlparse

import desktop

CHAT_URL = "https://m365.cloud.microsoft/chat"
CHAT_HOSTS = {"copilot.cloud.microsoft", "m365.cloud.microsoft"}
LOGIN_HOSTS = {
    "login.microsoftonline.com",
    "login.microsoft.com",
    "login.live.com",
    "login.windows.net",
    "account.live.com",
    "account.activedirectory.windowsazure.com",
    "mysignins.microsoft.com",
    "passwordreset.microsoftonline.com",
}
_READY_TIMEOUT = 300
_REPLY_TIMEOUT = 120
_NAME = re.IGNORECASE

_EXTRACT = r"""
() => {
  const deep = (root, selector) => {
    const out = [];
    const visit = (node) => {
      if (!node || !node.querySelectorAll) return;
      out.push(...node.querySelectorAll(selector));
      for (const el of node.querySelectorAll('*')) {
        if (el.shadowRoot) visit(el.shadowRoot);
      }
    };
    visit(root);
    return out;
  };
  const className = (node) => typeof node.className === 'string' ? node.className : '';
  const composers = deep(document, '[role="textbox"], textarea, [contenteditable="true"]');
  const composer = composers.filter(el => {
    const box = el.getBoundingClientRect();
    return box.width > 40 && box.height > 10;
  }).pop() || null;
  const main = deep(document, '[data-testid="markdown-reply"]').length
    ? document
    : (deep(document, 'main')[0] || document.body);
  const selectors = [
    '[data-testid="markdown-reply"]',
    '[data-testid*="message"]',
    '[data-automationid*="message"]',
    '[class*="markdown"]',
    '[class*="Markdown"]',
    'article',
    '[role="article"]'
  ];
  let nodes = [];
  for (const selector of selectors) {
    nodes = deep(main, selector).filter(el => {
      const text = (el.innerText || '').trim();
      if (!text) return false;
      if (composer && (el.contains(composer) || composer.contains(el))) return false;
      return true;
    });
    if (!nodes.length) continue;
    nodes = nodes.filter(el => !nodes.some(other => other !== el && el.contains(other)));
    while (nodes.length > 1 && (nodes[nodes.length - 1].innerText || '').trim().length < 15) {
      nodes.pop();
    }
    break;
  }
  const node = nodes.length ? nodes[nodes.length - 1] : null;
  const root = node || null;
  const pres = root ? deep(root, 'pre') : [];
  const codes = pres.length
    ? pres.map(el => {
        const code = el.querySelector('code');
        return [(code && className(code)) || '', el.innerText || ''];
      })
    : (root ? deep(root, 'code').map(el => [className(el), el.innerText || '']) : []);
  const imgs = root ? deep(root, 'img').filter(el => {
    const src = el.currentSrc || el.src || '';
    if (!src || src.startsWith('data:image/svg')) return false;
    if (el.naturalWidth && el.naturalWidth < 48) return false;
    return true;
  }).map(el => el.currentSrc || el.src) : [];
  return {
    text: root ? (root.innerText || '') : '',
    count: nodes.length,
    codes,
    imgs
  };
}
"""


# Microsoft sign-in pages, old and new layouts. Never touches password or code fields.
_SIGN_IN = r"""
(account) => {
  const visible = el => el && el.offsetParent !== null && el.getBoundingClientRect().width > 0;
  const text = el => (el.innerText || el.value || el.getAttribute('aria-label') || '').trim().toLowerCase();
  const press = el => { el.click(); return true; };
  const buttons = [...document.querySelectorAll('input[type=submit], button, [role=button], [data-test-id], div[tabindex]')].filter(visible);
  const byText = (...words) => buttons.find(el => words.some(w => text(el) === w));
  const body = (document.body.innerText || '').toLowerCase();

  if ([...document.querySelectorAll('input[type=password], input[name=otc], input[autocomplete=one-time-code]')].some(visible)) {
    return 'needs-user';
  }
  if (body.includes('stay signed in')) {
    const box = document.querySelector('#KmsiCheckboxField, input[name=DontShowAgain]');
    if (box && !box.checked) box.click();
    const yes = document.querySelector('#idSIButton9') || byText('yes');
    if (yes && press(yes)) return 'stay signed in';
  }
  if (account) {
    const tile = buttons.find(el => (el.getAttribute('data-test-id') || '').toLowerCase() === account)
      || buttons.find(el => text(el).includes(account) && text(el).length < account.length + 120);
    if (tile && press(tile)) return 'picked ' + account;
    const email = [...document.querySelectorAll('input[type=email], input[name=loginfmt]')].find(visible);
    if (email && !email.value) {
      email.focus();
      email.value = account;
      email.dispatchEvent(new Event('input', {bubbles: true}));
      email.dispatchEvent(new Event('change', {bubbles: true}));
      const next = document.querySelector('#idSIButton9') || byText('next');
      if (next && press(next)) return 'entered ' + account;
    }
  }
  return '';
}
"""


def proxy_args():
    """Skip the proxy when Windows' PAC server is unreachable (off the office network).

    Chromium otherwise waits ~20s for the PAC fetch before falling back to direct.
    A manual proxy, or a PAC server that answers, is left for Chromium to use as usual.
    """
    import socket
    import urllib.parse
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Internet Settings") as key:
            def value(name):
                try:
                    return winreg.QueryValueEx(key, name)[0]
                except OSError:
                    return None
            manual, pac, detect = value("ProxyEnable"), value("AutoConfigURL"), value("AutoDetect")
    except OSError:
        return []
    if manual:
        return []
    if pac:
        url = urllib.parse.urlsplit(pac)
        try:
            socket.create_connection((url.hostname, url.port or 80), timeout=1.5).close()
            return []
        except OSError:
            return ["--no-proxy-server"]
    return [] if detect else ["--no-proxy-server"]


_ACCOUNT = []


def work_account():
    """Work email from Windows (whoami /upn), lower-case, or '' when not domain joined."""
    if not _ACCOUNT:
        try:
            out = subprocess.run(["whoami.exe", "/upn"], capture_output=True, text=True, timeout=5,
                                 creationflags=subprocess.CREATE_NO_WINDOW).stdout.strip().lower()
        except Exception:
            out = ""
        _ACCOUNT.append(out if "@" in out else "")
    return _ACCOUNT[0]


class Stopped(Exception):
    pass


@dataclass
class ChatReply:
    text: str
    code_blocks: list
    images: list


def profile_dir():
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    path = Path(base) / "CoPilotComputerUse" / "edge-profile"
    path.mkdir(parents=True, exist_ok=True)
    return path


def is_login_url(url):
    host = (urlparse(url or "").hostname or "").lower()
    return host in LOGIN_HOSTS or host.endswith(".microsoftonline.com")


def is_chat_url(url):
    host = (urlparse(url or "").hostname or "").lower()
    return host in CHAT_HOSTS


def is_echo(text, outbound):
    text = (text or "").strip()
    outbound = (outbound or "").strip()
    if not text or not outbound:
        return False
    # Replies may quote the instruction's examples, so only a copy of our own message counts.
    return text == outbound or (len(outbound) > 120 and text.startswith(outbound[:120]))


_user32 = ctypes.windll.user32
_kernel32 = ctypes.windll.kernel32
_WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


def _exe_path(pid):
    handle = _kernel32.OpenProcess(0x1000, False, pid)
    if not handle:
        return ""
    try:
        buffer = ctypes.create_unicode_buffer(1024)
        size = wintypes.DWORD(1024)
        if _kernel32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
            return buffer.value
        return ""
    finally:
        _kernel32.CloseHandle(handle)


def browser_hwnd():
    """Top-level window of Playwright's Chromium (its install path contains 'chromium-')."""
    found = []

    def visit(hwnd, _param):
        if not _user32.IsWindowVisible(hwnd) or _user32.GetWindowTextLengthW(hwnd) == 0:
            return True
        name = ctypes.create_unicode_buffer(64)
        _user32.GetClassNameW(hwnd, name, 64)
        if name.value != "Chrome_WidgetWin_1":
            return True
        pid = wintypes.DWORD()
        _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if "chromium-" in _exe_path(pid.value).lower():
            found.append(hwnd)
        return True

    _user32.EnumWindows(_WNDENUMPROC(visit), 0)
    return found[0] if found else None


class _GUID(ctypes.Structure):
    _fields_ = [("data", ctypes.c_byte * 16)]


def drop_taskbar_button(hwnd):
    """Remove the browser's taskbar button. It stays gone until the window is hidden and shown again."""
    ole32 = ctypes.windll.ole32
    ole32.CoInitialize(None)
    clsid, iid = _GUID(), _GUID()
    ole32.CLSIDFromString("{56FDF344-FD6D-11d0-958A-006097C9A090}", ctypes.byref(clsid))
    ole32.CLSIDFromString("{56FDF342-FD6D-11d0-958A-006097C9A090}", ctypes.byref(iid))
    taskbar = ctypes.c_void_p()
    if ole32.CoCreateInstance(ctypes.byref(clsid), None, 1, ctypes.byref(iid), ctypes.byref(taskbar)) != 0:
        return
    vtable = ctypes.cast(taskbar, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
    call = lambda index, *types: ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, *types)(vtable[index])
    call(3)(taskbar)
    call(5, wintypes.HWND)(taskbar, hwnd)
    call(2)(taskbar)


class CopilotChat:
    def __init__(self, status, log, stop, mirror=None, mirror_wanted=None):
        self.status = status
        self.log = log
        self.stop = stop
        self.mirror = mirror
        self.mirror_wanted = mirror_wanted or threading.Event()
        if mirror_wanted is None:
            self.mirror_wanted.set()
        self._mirrored_at = 0.0
        self._playwright = None
        self.context = None
        self.page = None
        self._session = None
        self._window_id = None
        self._visible = False
        self._asked_sign_in = False
        self._signin_acted_at = 0.0
        self._returned_to_chat = False
        self._window_warned = False

    def open(self):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError as exc:
            raise RuntimeError("Install dependencies: pip install -r requirements.txt") from exc
        self._playwright = sync_playwright().start()
        try:
            self.context = self._playwright.chromium.launch_persistent_context(
                user_data_dir=str(profile_dir()),
                headless=False,
                viewport={"width": 960, "height": 800},
                args=[
                    "--disable-background-timer-throttling",
                    "--disable-renderer-backgrounding",
                    "--disable-backgrounding-occluded-windows",
                    "--disable-blink-features=AutomationControlled",
                    "--no-first-run",
                    "--no-default-browser-check",
                    "--disable-fre",
                    "--hide-crash-restore-bubble",
                    "--window-position=-32000,0",
                    "--window-size=1100,800",
                    *proxy_args(),
                ],
                ignore_default_args=["--enable-automation"],
            )
        except Exception as exc:
            self.close()
            low = str(exc).lower()
            if "executable doesn't exist" in low or "please run the following command" in low:
                raise RuntimeError(
                    "Playwright's Chromium is not installed. Run: python -m playwright install chromium"
                ) from exc
            if "remote debugging is disallowed" in low:
                raise RuntimeError(
                    "This browser disallows remote debugging, so Copilot cannot be driven."
                ) from exc
            if "processsingleton" in low or "already in use" in low:
                raise RuntimeError(
                    "The Copilot browser profile is already open. Close the other window and try again."
                ) from exc
            raise
        self.context.on("page", self._on_page)
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        self.page.set_default_timeout(10_000)
        self._hide()
        self._drop_taskbar_button()
        try:
            self.page.goto(CHAT_URL, wait_until="domcontentloaded", timeout=60_000)
        except Exception as exc:
            self.log(f"Navigation: {exc}")
        self._wait_until_ready()
        self._new_chat()
        self._go_temporary()
        self._hide()

    def _drop_taskbar_button(self):
        deadline = time.time() + 10
        while time.time() < deadline:
            hwnd = browser_hwnd()
            if hwnd:
                drop_taskbar_button(hwnd)
                return
            time.sleep(0.2)

    def idle(self, seconds):
        """Keep the mirror fresh while waiting for the next task."""
        deadline = time.time() + seconds
        while time.time() < deadline:
            self._mirror()
            time.sleep(0.25)

    def _mirror(self):
        if not self.mirror or not self.mirror_wanted.is_set() or time.time() - self._mirrored_at < 1.0:
            return
        self._mirrored_at = time.time()
        try:
            self.mirror(self._page().screenshot(type="jpeg", quality=60, timeout=3000))
        except Exception:
            pass

    def close(self):
        context, playwright = self.context, self._playwright
        self.context = None
        self.page = None
        self._session = None
        self._window_id = None
        self._playwright = None
        try:
            if context:
                context.close()
        except Exception:
            pass
        try:
            if playwright:
                playwright.stop()
        except Exception:
            pass

    def send(self, text, jpeg=None, ready=None):
        self._check_stop()
        self._wait_until_ready()
        before = self._extract()
        if jpeg is not None:
            self._attach(jpeg)
        self._insert_text(text)
        deadline = time.time() + 20
        while not self._posted():
            self._check_stop()
            if time.time() > deadline:
                raise RuntimeError("Copilot did not accept the message")
            if not self._click_send():
                self._press_enter()
            for _ in range(8):
                time.sleep(0.25)
                if self._posted():
                    break
        return self._wait_reply(before, text, ready)

    def _posted(self):
        """An empty composer means the message went out."""
        try:
            composer = self._composer()
            return composer is not None and not composer.inner_text().strip()
        except Exception:
            return False

    def _attachment_count(self):
        return self._page().get_by_role("button", name=re.compile(r"^remove attachment", _NAME)).count()

    def side_page(self):
        """A second tab in the same signed-in browser for other sites (Azure DevOps); the chat ignores it."""
        side = getattr(self, "_side", None)
        if side and not side.is_closed():
            return side
        self._opening_side = True
        try:
            self._side = self.context.new_page()
        finally:
            self._opening_side = False
        self._side.set_default_timeout(30_000)
        return self._side

    def front(self):
        try:
            self._page().bring_to_front()
        except Exception:
            pass

    def _on_page(self, page):
        if getattr(self, "_opening_side", False) or page is getattr(self, "_side", None):
            return
        self.page = page
        self._session = None
        self._window_id = None
        self._visible = False

    def _check_stop(self):
        if self.stop.is_set():
            raise Stopped()

    def _page(self):
        if self.page and not self.page.is_closed():
            return self.page
        if self.context:
            for page in self.context.pages:
                if not page.is_closed():
                    self._on_page(page)
                    return page
        raise RuntimeError("Copilot window closed")

    def _wait_until_ready(self, timeout=_READY_TIMEOUT):
        deadline = time.time() + timeout
        stuck_since = None
        login_since = None
        while time.time() < deadline:
            self._check_stop()
            try:
                url = self._page().url
                if not is_login_url(url):
                    login_since = None
                self._dismiss_cookies()
                if is_login_url(url):
                    stuck_since = None
                    if login_since is None:
                        login_since = time.time()
                    step = self._auto_sign_in()
                    if step:
                        login_since = time.time()
                        self.status(f"Signing in automatically ({step})")
                    elif step is None or time.time() - login_since > 6:
                        self.status("Sign in with your work account. This window hides when chat is ready.")
                        self._show()
                elif self._composer() is not None:
                    self._hide()
                    self.status("Chat ready")
                    return
                elif self._signed_out():
                    stuck_since = None
                    self.status("Sign in with your work account. This window hides when chat is ready.")
                    self._show()
                    self._start_sign_in()
                else:
                    if stuck_since is None:
                        stuck_since = time.time()
                    elif not self._returned_to_chat and time.time() - stuck_since > 3 and is_chat_url(url):
                        self._returned_to_chat = True
                        self._page().goto(CHAT_URL, wait_until="domcontentloaded", timeout=60_000)
                    elif time.time() - stuck_since > 8:
                        self.status("Waiting for the Copilot chat. Use this window if a prompt is in the way.")
                        self._show()
            except Stopped:
                raise
            except Exception:
                pass
            self._mirror()
            time.sleep(0.5)
        raise TimeoutError("Copilot chat did not become ready. Sign-in may still be required.")

    def _composer_present(self):
        try:
            return self._composer() is not None
        except Exception:
            return False

    def _composer(self):
        try:
            page = self._page()
            named = page.get_by_role(
                "textbox",
                name=re.compile(r"message copilot|ask copilot|ask me anything", _NAME),
            )
            if named.count():
                return named.last
            if self._signed_out():
                return None
            loc = page.get_by_role("textbox")
            if loc.count():
                return loc.last
        except Exception:
            return None
        return None

    def _signed_out(self):
        try:
            page = self._page()
        except Exception:
            return False
        for role in ("link", "button"):
            try:
                loc = page.get_by_role(role, name=re.compile(r"^sign in$", _NAME))
                count = min(loc.count(), 5)
            except Exception:
                continue
            for index in range(count):
                try:
                    if loc.nth(index).is_visible():
                        return True
                except Exception:
                    continue
        return False

    def _auto_sign_in(self):
        """Handle Microsoft sign-in pages that need no secret.

        Returns the step taken, "" when nothing applied yet, or None when a
        password, code, or other prompt needs the person.
        """
        account = work_account()
        if time.time() - self._signin_acted_at < 2.5:
            return ""
        try:
            step = self._page().evaluate(_SIGN_IN, account)
        except Exception:
            return ""
        if step == "needs-user":
            return None
        if step:
            self._signin_acted_at = time.time()
        return step or ""

    def _dismiss_cookies(self):
        self._click_role("button", r"^reject$|^refuser$")

    def _start_sign_in(self):
        if self._asked_sign_in or is_login_url(self._page().url):
            return
        self._asked_sign_in = True
        try:
            href = self._page().evaluate(
                """() => {
                    const link = [...document.querySelectorAll('a')].find(el => {
                        const label = (el.getAttribute('aria-label') || el.innerText || '').trim();
                        return /^sign in$/i.test(label) && el.href && el.getBoundingClientRect().width > 20;
                    });
                    return link ? link.href : '';
                }"""
            )
        except Exception:
            href = ""
        if href:
            try:
                self._page().goto(href, wait_until="domcontentloaded", timeout=60_000)
                return
            except Exception as exc:
                self.log(f"Sign-in: {exc}")
        self._click_role("link", r"^sign in$") or self._click_role("button", r"^sign in$")

    def _new_chat(self):
        page = self._page()
        loc = page.get_by_role("link", name=re.compile(r"^new chat$|^nouvelle conversation$|^nouveau clavardage$", _NAME))
        if loc.count() == 0:
            loc = page.get_by_role("button", name=re.compile(r"new chat|nouvelle conversation|nouveau clavardage", _NAME))
        if loc.count() == 0:
            # Narrow layouts hide the sidebar; loading CHAT_URL already opened a fresh chat.
            return
        loc.first.click()
        time.sleep(0.6)
        self._asked_sign_in = False
        self._wait_until_ready(timeout=60)

    def _go_temporary(self):
        """Temporary chats stay out of Recent chats and Copilot memory."""
        page = self._page()
        temporary = page.get_by_role("textbox", name=re.compile(r"temporarily message", _NAME))
        for _ in range(2):
            if temporary.count() or not self._click_role("button", r"^temporary chat$"):
                break
            try:
                temporary.wait_for(timeout=4000)
            except Exception:
                continue
        if temporary.count():
            self.log("Using a temporary chat (not saved to Recent chats)")
        else:
            self.log("Temporary chat unavailable; using a normal chat")

    def _click_role(self, role, pattern):
        try:
            loc = self._page().get_by_role(role, name=re.compile(pattern, _NAME))
            if loc.count() == 0:
                return False
            loc.first.evaluate("el => el.click()")
            return True
        except Exception:
            return False

    def _insert_text(self, text):
        composer = self._composer()
        if composer is None:
            raise RuntimeError("Copilot composer not found")
        composer.evaluate(
            """(el, value) => {
                el.focus();
                const selection = window.getSelection();
                if (selection) {
                    const range = document.createRange();
                    range.selectNodeContents(el);
                    selection.removeAllRanges();
                    selection.addRange(range);
                }
                document.execCommand('delete', false, null);
                const inserted = document.execCommand('insertText', false, value);
                if (!inserted) {
                    el.textContent = value;
                    el.dispatchEvent(new InputEvent('input', {
                        bubbles: true,
                        inputType: 'insertText',
                        data: value
                    }));
                }
            }""",
            text,
        )

    def _click_send(self):
        page = self._page()
        locators = (
            page.locator('[data-testid="chat-input-send-button"]'),
            page.get_by_role("button", name=re.compile(r"^(send|submit|envoyer)$", _NAME)),
        )
        for loc in locators:
            try:
                count = loc.count()
            except Exception:
                continue
            for index in range(count):
                button = loc.nth(index)
                try:
                    if button.is_disabled():
                        continue
                    button.click()
                    return True
                except Exception:
                    continue
        return False

    def _press_enter(self):
        composer = self._composer()
        if composer is None:
            raise RuntimeError("Copilot send control not found")
        composer.evaluate(
            """(el) => {
                el.dispatchEvent(new KeyboardEvent('keydown', {
                    key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true
                }));
            }"""
        )

    def _attach(self, jpeg):
        handle, path = tempfile.mkstemp(suffix=".jpg")
        os.close(handle)
        try:
            Path(path).write_bytes(jpeg)
            if self._set_files(path):
                deadline = time.time() + 15
                while self._attachment_count() == 0 and time.time() < deadline:
                    self._check_stop()
                    time.sleep(0.2)
                time.sleep(0.3)
                return
            page = self._page()
            attach = page.locator('[data-testid="chat-input-attach-button"]')
            if attach.count() == 0:
                attach = page.get_by_role("button", name=re.compile(r"^add$", _NAME))
            if attach.count():
                opened = False
                try:
                    with page.expect_file_chooser(timeout=2500) as waiting:
                        attach.first.click()
                    waiting.value.set_files(path)
                    opened = True
                except Exception:
                    opened = False
                if opened:
                    time.sleep(0.4)
                    return
                if self._choose_files("menuitem", r"this device|upload|browse|photo|file|image|cet appareil", path):
                    time.sleep(0.4)
                    return
                files_tab = page.locator('[data-testid="peek-pivot-files"]')
                if files_tab.count():
                    files_tab.first.evaluate("el => el.click()")
                    time.sleep(0.3)
                if self._set_files(path) or self._choose_files("button", r"upload|browse|this device|add file", path):
                    time.sleep(0.4)
                    return
            raise RuntimeError("Copilot has no image upload control")
        finally:
            try:
                os.remove(path)
            except OSError:
                pass

    def _set_files(self, path):
        target = self._pick_file_input()
        if target is None:
            return False
        target.set_input_files(path)
        return True

    def _choose_files(self, role, pattern, path):
        # Trusted click so the file chooser opens. It stays inside Edge and does not move the desktop mouse.
        page = self._page()
        try:
            loc = page.get_by_role(role, name=re.compile(pattern, _NAME))
            if loc.count() == 0:
                return False
            with page.expect_file_chooser(timeout=2500) as waiting:
                loc.first.click()
            waiting.value.set_files(path)
            return True
        except Exception:
            return self._set_files(path)

    def _pick_file_input(self):
        try:
            locator = self._page().locator("input[type=file]")
            count = locator.count()
        except Exception:
            return None
        blank = None
        for index in range(count):
            item = locator.nth(index)
            accept = (item.get_attribute("accept") or "").lower()
            if "image" in accept or any(ext in accept for ext in (".png", ".jpg", ".jpeg", ".gif", ".webp")):
                return item
            if accept == "" and blank is None:
                blank = item
        return blank

    def _wait_reply(self, before, outbound, ready=None):
        deadline = time.time() + _REPLY_TIMEOUT
        started = time.time()
        seen_generating = False
        last_sig = None
        stable = 0
        while time.time() < deadline:
            self._check_stop()
            try:
                if is_login_url(self._page().url):
                    self._wait_until_ready()
                if self._generating():
                    seen_generating = True
                data = self._extract()
            except Stopped:
                raise
            except Exception:
                time.sleep(0.5)
                continue
            signature = (data["text"], tuple(body for _lang, body in data["codes"]))
            grown = data["count"] > before["count"] or (
                data["text"] and data["text"] != before["text"]
            )
            if ready and grown and not is_echo(data["text"], outbound):
                quick = self._to_reply(data, images=False)
                if ready(quick):
                    return quick
            settled = (
                signature == last_sig
                and not self._generating()
                and grown
                and not is_echo(data["text"], outbound)
                and (seen_generating or time.time() > started + 4)
            )
            if settled:
                stable += 1
                if stable >= 2:
                    return self._to_reply(data)
            else:
                stable = 0
            last_sig = signature
            self._mirror()
            time.sleep(0.25)
        data = self._extract()
        if is_echo(data["text"], outbound):
            return ChatReply("", [], [])
        return self._to_reply(data)

    def _generating(self):
        try:
            loc = self._page().get_by_role(
                "button",
                name=re.compile(r"^stop( generating| responding)?$|^arr[eê]ter", _NAME),
            )
            return loc.count() > 0
        except Exception:
            return False

    def _extract(self):
        try:
            data = self._page().evaluate(_EXTRACT)
        except Exception:
            return {"text": "", "count": 0, "codes": [], "imgs": []}
        codes = []
        for item in data.get("codes") or []:
            if not item or len(item) < 2:
                continue
            body = str(item[1] or "").strip()
            if body:
                codes.append((str(item[0] or ""), body))
        return {
            "text": str(data.get("text") or ""),
            "count": int(data.get("count") or 0),
            "codes": codes,
            "imgs": [str(src) for src in (data.get("imgs") or []) if src],
        }

    def _to_reply(self, data, images=True):
        blocks = []
        for lang, body in data["codes"]:
            match = re.search(r"language-([\w-]+)", lang or "")
            blocks.append((match.group(1) if match else "", body))
        found = self._download_images(data["imgs"]) if images else []
        return ChatReply(data["text"].strip(), blocks, found)

    def _download_images(self, urls):
        images = []
        for url in urls[:4]:
            if len(images) >= 4:
                break
            try:
                data = self._read_image(url)
            except Exception:
                continue
            if data and 200 < len(data) <= 8_000_000:
                images.append(data)
        return images

    def _read_image(self, url):
        if url.startswith("data:"):
            _header, _sep, payload = url.partition(",")
            return base64.b64decode(payload)
        if url.startswith("blob:"):
            encoded = self._page().evaluate(
                """async (src) => {
                    const response = await fetch(src);
                    const buffer = await response.arrayBuffer();
                    const bytes = new Uint8Array(buffer);
                    let binary = '';
                    const step = 0x8000;
                    for (let i = 0; i < bytes.length; i += step) {
                        binary += String.fromCharCode(...bytes.subarray(i, i + step));
                    }
                    return btoa(binary);
                }""",
                url,
            )
            return base64.b64decode(encoded)
        if not url.startswith(("http://", "https://")):
            return b""
        response = self.context.request.get(url, timeout=15_000)
        if not response.ok:
            return b""
        return response.body()

    def _hide(self):
        if not self._visible and self._window_id is not None:
            return
        self._place(-32000, 0, 1100, 800)
        self._visible = False

    def _show(self):
        if self._visible:
            return
        width, height = desktop.screen_size()
        box_w, box_h = 1100, 800
        self._place(max(0, (width - box_w) // 2), max(0, (height - box_h) // 2), box_w, box_h)
        try:
            self._page().bring_to_front()
        except Exception:
            pass
        self._visible = True

    def _place(self, left, top, width, height):
        try:
            session = self._cdp()
            window_id = self._browser_window(session)
            session.send("Browser.setWindowBounds", {
                "windowId": window_id,
                "bounds": {"windowState": "normal"},
            })
            session.send("Browser.setWindowBounds", {
                "windowId": window_id,
                "bounds": {
                    "left": int(left),
                    "top": int(top),
                    "width": int(width),
                    "height": int(height),
                    "windowState": "normal",
                },
            })
        except Exception as exc:
            if not self._window_warned:
                self._window_warned = True
                self.log(f"Could not move the Copilot window: {exc}")

    def _cdp(self):
        if self._session is None:
            self._session = self.context.new_cdp_session(self._page())
        return self._session

    def _browser_window(self, session):
        if self._window_id is None:
            self._window_id = session.send("Browser.getWindowForTarget")["windowId"]
        return self._window_id


def _check():
    assert is_login_url("https://login.microsoftonline.com/common/oauth2/v2.0/authorize")
    assert is_login_url("https://login.microsoftonline.com/common/SAS/ProcessAuth")
    assert not is_login_url("https://copilot.cloud.microsoft/")
    assert not is_login_url("https://m365.cloud.microsoft/chat")
    assert is_chat_url("https://copilot.cloud.microsoft/chats/abc")
    assert is_echo("open notepad", "open notepad")
    assert not is_echo('{"action":"done","summary":"ok"}', "Task:\nopen notepad")
    long_out = "You control this Windows PC. " + ("x" * 200)
    assert is_echo(long_out, long_out)
    assert not is_echo('{"action":"click","x":4,"y":5}', long_out)
    # A reply that quotes an example from the instruction is still a reply.
    assert not is_echo('{"action":"open_app","name":"notepad"}', 'Actions:\n{"action":"open_app","name":"notepad"}\n')
    print("copilot helpers ok")


if __name__ == "__main__":
    _check()
