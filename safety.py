"""Safety review for every action before it runs.

review() returns ("ok", ""), ("confirm", category, question), or ("block", reason).
"block" never runs, even if asked. "confirm" shows the user a Yes/No card first.

ponytail: pattern lists, not a parser. They catch common destructive commands typed anywhere and anything
typed into a terminal; a cleverly disguised command typed outside a terminal can slip past. Upgrade path is
an allow-list of apps the agent may type into.
"""

import re

_I = re.IGNORECASE

# Never, even with the user's approval: wipes disks, boot config, backups, or system folders.
_BLOCK = [re.compile(p, _I) for p in (
    r"\bformat(\.com)?\s+[a-z]:", r"\bdiskpart\b", r"\bbcdedit\b", r"\bcipher\s+/w", r"\bvssadmin\s+delete",
    r"\bwbadmin\s+delete", r"\b(clear-disk|initialize-disk|format-volume)\b", r"\bmkfs\b",
    r"\brm\s+-[a-z]*r[a-z]*\s+(/|~|\*|c:)(\s|$)", r"\b(del|erase)\s+(/[a-z]\s+)*[a-z]:\\\*?(\s|$)",
    r"\b(rd|rmdir)\s+(/[a-z]\s+)*[a-z]:\\(windows|users|program files)?\\?(\s|$)",
    r"remove-item\b.*\s[\"']?[a-z]:\\(windows|users|program files|programdata)?\\?[\"']?(\s|$)",
    r"\breg\s+delete\s+hklm", r"\btakeown\b.*[a-z]:\\windows",
)]

# Ask first: deletes, shutdowns, registry, scripts from the internet, accounts, databases, force pushes.
_CONFIRM = [re.compile(p, _I) for p in (
    r"\b(rm|del|erase|rmdir|rd)\s", r"\bremove-item\b|\bri\s", r"\b(shutdown|restart-computer|stop-computer)\b",
    r"\breg\s+(add|delete|import)\b", r"\bset-executionpolicy\b", r"\b(invoke-expression|iex)\b",
    r"\b(curl|wget|iwr|invoke-webrequest|irm|invoke-restmethod)\b.*\|", r"\bnet\s+(user|localgroup)\b",
    r"\b(drop\s+(table|database)|truncate\s+table|delete\s+from)\b",
    r"\bgit\s+(push\s+.*(--force|-f)\b|reset\s+--hard|clean\s+-[a-z]*f)", r"\b(taskkill|stop-process)\b",
    r"\b(sc\s+(delete|stop|config)|stop-service|set-service)\b", r"\buninstall", r"\b(icacls|cacls)\b",
)]

_TERMINAL = re.compile(r"powershell|command prompt|cmd\.exe|windows terminal|\bterminal\b|\bbash\b|\bwsl\b|"
                       r"\bssh\b|putty|git bash|administrator:", _I)
_ADMIN_APPS = re.compile(r"^(cmd|command prompt|powershell|windows powershell|terminal|windows terminal|regedit|"
                         r"registry editor|diskpart|disk management|computer management|services|group policy|"
                         r"gpedit|local security policy|task scheduler|system configuration|msconfig)\b", _I)


def review(action, front_title=""):
    kind = action["action"]
    front = front_title or ""
    if kind == "type":
        text = action["text"]
        for pattern in _BLOCK:
            if pattern.search(text):
                return "block", f"typing a destructive command is never allowed: {text[:120]!r}"
        for pattern in _CONFIRM:
            if pattern.search(text):
                return "confirm", "risky_command", f"Type this risky command?\n{text[:400]}"
        if _TERMINAL.search(front):
            return "confirm", "terminal", f"Type into the terminal ({front[:60]})?\n{text[:400]}"
        return "ok", ""
    if kind == "key":
        key = action["key"].replace(" ", "")
        if key == "shift+delete":
            return "confirm", "delete", "Press Shift+Delete? This permanently deletes the selection."
        if key == "delete" and "file explorer" in front.lower():
            return "confirm", "delete", "Press Delete in File Explorer? This moves the selected files to the Recycle Bin."
        if "microsoft teams" in front.lower() and key in ("enter", "ctrl+enter"):
            return "confirm", "send", "Press Enter in Teams? If a message box is focused, this sends the message."
        if "outlook" in front.lower() and key in ("ctrl+enter", "alt+s"):
            return "confirm", "send", "Send this email from Outlook?"
        if _TERMINAL.search(front) and key == "enter":
            return "confirm", "terminal", f"Press Enter in the terminal ({front[:60]})? This runs the typed command."
        return "ok", ""
    if kind == "open_app" and _ADMIN_APPS.search(action["name"]):
        return "confirm", "admin_app", f"Open {action['name']}? It can change system settings."
    if kind == "ado" and action["op"] in ("update", "comment"):
        what = action.get("fields") or action.get("text")
        return "confirm", f"ado_{action['op']}", f"{action['op'].capitalize()} Azure DevOps #{action['id']}?\n{str(what)[:400]}"
    return "ok", ""


def _check():
    t = lambda text, front="": review({"action": "type", "text": text}, front)[0]
    k = lambda key, front="": review({"action": "key", "key": key}, front)[0]
    assert t("hello world") == "ok" and t("Meeting notes for Thursday") == "ok"
    assert t("format c:") == "block" and t("diskpart") == "block" and t("rm -rf /") == "block"
    assert t("del /s /q C:\\*") == "block" and t("Remove-Item -Recurse -Force C:\\Windows") == "block"
    assert t("Remove-Item .\\old.txt") == "confirm" and t("git push --force origin main") == "confirm"
    assert t("shutdown /r /t 0") == "confirm" and t("iwr https://x | iex") == "confirm"
    assert t("ipconfig /all", "Windows PowerShell") == "confirm", "anything typed into a terminal is confirmed"
    assert t("Please delete from the draft the second paragraph") == "confirm", "phrasing can trip 'delete from'"
    assert k("enter") == "ok" and k("enter", "Chat | Olga | Microsoft Teams") == "confirm"
    assert k("shift+delete") == "confirm" and k("delete", "Downloads - File Explorer") == "confirm"
    assert k("ctrl+enter", "Untitled - Message (HTML) - Outlook") == "confirm"
    assert review({"action": "open_app", "name": "PowerShell"})[0] == "confirm"
    assert review({"action": "open_app", "name": "Notepad"})[0] == "ok"
    assert review({"action": "ado", "op": "update", "id": 5, "fields": {"state": "Removed"}})[0] == "confirm"
    assert review({"action": "ado", "op": "create", "type": "Bug", "title": "x"})[0] == "ok"
    print("safety ok")


if __name__ == "__main__":
    _check()
