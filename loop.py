"""One Copilot session for the app's lifetime. Each reply is a batch of up to 5 actions; up to 40 rounds per task."""

import json
import queue
import re
import threading
import time

import pyautogui

import account
import commute
import ado
import meetings
import copilot_chat
import desktop
import files
import insideops
import protocol
import safety
import sysinfo

MAX_ROUNDS = 40
_SETTLE = 0.35


class Skipped(Exception):
    """The user pressed Skip. The task ends; nothing is sent back to Copilot."""


# Queued by New chat. Not a task.
RESET = object()

_CANT = ("I can't do that from here. I can use this PC, check your work account "
         "(calendar, mail, tasks, files), or answer a question.")
# A whole-message shrug, or a promise to act. A finished sentence is shown as-is.
_PLAN = re.compile(
    r"^(?:sure|ok|okay|got it|on it|absolutely|of course)[.!]?\s*$"
    r"|(?:\bi'?ll\b|\blet me\b|\bi(?:'| a)?m going to\b)",
    re.I,
)


# Only the opening of a reply: "I can't access…". An explanation that mentions "cannot" later is fine.
_REFUSE = re.compile(
    r"^\W*(?:sorry[,.]?\s*)?i(?:\s+|'| a)?(?:m |am )?(?:can[’']?t|cannot|unable|not able|don[’']?t have (?:access|the ability)|"
    r"do not have (?:access|the ability))",
    re.I,
)
_NUDGE = ("You refused without trying a tool. account reads calendar, mail, people, tasks, OneDrive files, and Teams "
          "meeting recordings and transcripts (op recordings, query = meeting name or person). The screen (look, then "
          "mouse and keyboard) can use any app on this PC. Try the right one now. Only if none can do it, reply with "
          "done and one polite sentence.")
_VERIFY = "Look at this screenshot and confirm the task really worked before done. If it did not, fix it."


class Memory:
    """This chat's recent turns, plus the last account read so a short follow-up stays local."""

    def __init__(self):
        self.turns = []
        self.last_account = None

    def add(self, role, text):
        text = " ".join(str(text or "").split())
        if not text:
            return
        if len(text) > 500:
            text = text[:500] + "..."
        self.turns.append((role, text))
        del self.turns[:-8]  # ponytail: last 4 exchanges, 500 chars each. Older Copilot turns stay in the browser thread.

    def block(self):
        if not self.turns:
            return ""
        lines = ["User: " + text if role == "user" else "Co-opilot: " + text for role, text in self.turns]
        return "Conversation so far (the new message may refer to this):\n" + "\n".join(lines) + "\n"

    def remember(self, action):
        if not action:
            return
        slim = {key: action[key] for key in ("op", "when", "query", "focus") if action.get(key)}
        if slim.get("op"):
            self.last_account = slim

    def clear(self):
        self.turns.clear()
        self.last_account = None
        meetings.clear()


def user_reply(text):
    """Prose worth showing, or None when it is empty, a status line, JSON, or a plan to act."""
    raw = text or ""
    if "```" in raw or raw.lstrip()[:1] in "{[":
        return None
    clean = " ".join(raw.split())
    if len(clean) < 16 or account.status_only(clean) or _PLAN.search(clean):
        return None
    return clean[:1200]


class Tools:
    """What skills need besides the action itself: the chat's browser, and ways to talk to the user."""

    def __init__(self, chat, log, stop, say, ask, hands_on=None):
        self.chat, self.log, self.stop = chat, log, stop
        self.say = say or (lambda text: log("Note: " + text))
        self._ask = ask
        self.hands_on = hands_on or (lambda: None)
        self.ado = ado.Ado(chat)
        self.allowed = set()
        self.retried = False
        self.nudged = False
        self.verified = False
        self.memory = Memory()

    def ask(self, question, options=(), multi=False, allow_text=True):
        """Show a question card and wait. Skip ends the task instead of telling Copilot."""
        if self._ask is None:
            return None
        reply = queue.Queue()
        self._ask({"question": question, "options": list(options), "multi": multi, "allow_text": allow_text}, reply)
        while True:
            if self.stop.is_set():
                raise copilot_chat.Stopped()
            try:
                got = reply.get(timeout=0.2)
            except queue.Empty:
                self.chat.idle(0.2)
                continue
            if got is None:
                raise Skipped()
            return got

    def confirm(self, question, category=None):
        """Yes/No card. With a category, "Yes, for the rest of this task" skips later asks of the same kind."""
        if category and category in self.allowed:
            return True
        options = ["Yes, go ahead"] + (["Yes, for the rest of this task"] if category else []) + ["No, cancel"]
        answer = self.ask(question, options, allow_text=False)
        picked = (answer or {}).get("choices") or []
        if picked == ["Yes, for the rest of this task"]:
            self.allowed.add(category)
        return bool(picked) and picked[0].startswith("Yes")


def session(tasks, log, status, stop, on_ready, on_idle, mirror=None, mirror_wanted=None, answer=None,
            say=None, ask=None, hands_on=None, on_end=None, on_files=None):
    """Open the chat once, then run tasks from the queue until it yields None."""
    facts = {}
    threading.Thread(target=lambda: facts.setdefault("system", sysinfo.collect("system")), daemon=True).start()
    chat = copilot_chat.CopilotChat(status, log, stop, mirror, mirror_wanted)
    tools = Tools(chat, log, stop, say, ask, hands_on)
    try:
        status("Preparing chat")
        chat.open()
        status("Ready")
        on_ready()
        # ponytail: one chat for every task; the instruction goes in once. New chat clears memory
        # and opens a fresh browser thread, which sends the instruction again.
        instructed = False
        while True:
            try:
                task = tasks.get(timeout=0.2)
            except queue.Empty:
                chat.idle(1.0)
                continue
            if task is None:
                return
            if task is RESET:
                tools.memory.clear()
                instructed = False
                try:
                    status("New chat")
                    chat._new_chat()
                    chat._go_temporary()
                    status("Selecting GPT model")
                    if not chat._select_model():
                        log(f"Could not verify {copilot_chat.PREFERRED_MODEL}; keeping the current model")
                except Exception as exc:
                    log(f"New chat: {exc}")
                status("Ready")
                continue
            stop.clear()
            files.touched.clear()
            tools.allowed.clear()
            tools.retried = tools.nudged = tools.verified = False
            started = time.time()
            outcome = {"answer": "Not finished"}

            def answered(text):
                outcome["answer"] = text
                (answer or log)(text)

            try:
                if _local_open(task, log, status, answered):
                    pass
                elif _drive(task, chat, log, status, stop, instructed, facts, answered, tools) is not False:
                    instructed = True
            except Skipped:
                log("Skipped")
                status("Done")
                outcome["answer"] = "Skipped"
                if on_end:
                    on_end()
            except copilot_chat.Stopped:
                log("Stopped")
                status("Stopped")
                outcome["answer"] = "Stopped by the user"
            except pyautogui.FailSafeException:
                log("Stopped: mouse is in a screen corner")
                status("Stopped")
                outcome["answer"] = "Stopped by the user"
            except Exception as exc:
                log(f"Error: {exc}")
                status("Error")
                outcome["answer"] = f"Failed: {exc}"
            if files.touched and on_files:
                on_files([str(p) for p in files.touched if p.exists()])
            if outcome["answer"] not in ("Skipped", "Stopped by the user", "Not finished"):
                tools.memory.add("user", task)
                tools.memory.add("assistant", outcome["answer"])
            sysinfo.record_history(task, outcome["answer"], time.time() - started)
            log(f"Task took {time.time() - started:.0f}s")
            on_idle()
    except copilot_chat.Stopped:
        status("Stopped")
    except Exception as exc:
        log(f"Error: {exc}")
        status("Error")
    finally:
        chat.close()


def _drive(task, chat, log, status, stop, instructed, facts, answer, tools):
    prior = tools.memory.turns[-1][1] if tools.memory.turns and tools.memory.turns[-1][0] == "assistant" else ""
    try:
        go = commute.ask(task)
    except Exception as exc:
        go = f"I couldn't read the GO timetable just now. {exc}"
    if go:
        log("GO timetable (not asking Copilot)")
        answer(go)
        status("Done")
        return False
    if insideops.wants(task):
        log("InsideOPS (background browser, no screen)")
        try:
            text = insideops.answer(task, chat, log, status)
        except copilot_chat.Stopped:
            raise
        except Exception as exc:
            text = f"I couldn't read InsideOPS just now. {exc}"
        answer(text)
        status("Done")
        return False
    found = account.match(task, tools.memory.last_account, prior)
    if found:
        log(f"Account {found['op']} (not asking Copilot)")
        status("Reading your account")
        tools.memory.remember(found)
        try:
            text = account.run(found, chat)
        except Exception as exc:
            text = f"I couldn't read that from your account.\n{exc}"
        answer(text)
        status("Done")
        return False
    report = None
    shot = None          # last screenshot Copilot saw; clicks are mapped against it
    want_shot = False    # set by look, or after mouse/keyboard actions
    hands_on = False
    tried = False        # any tool ran in this task; a refusal before that gets one nudge
    keyboard_navigation = 0
    last_mouse_target = None
    keyboard_window = None
    display_target = re.search(r"monitor\s*(\d+)", task, re.I)
    display_focus = None
    for round_no in range(1, MAX_ROUNDS + 1):
        if stop.is_set():
            raise copilot_chat.Stopped()
        status(f"Step {round_no}")
        jpeg = None
        if want_shot:
            shot = desktop.screenshot()
            jpeg = shot.jpeg
            seen = f"Screenshot attached: {shot.sent_w} x {shot.sent_h} pixels. Monitors: {shot.layout()}."
        else:
            seen = "No screenshot attached (use look if you need to see the screen)."
        want_shot = False
        if report is None:
            head = protocol.NEXT_TASK if instructed else protocol.INSTRUCTION
            known = ""
            if not instructed:
                system = _facts_text(facts, stop)
                known = f"\nThis PC:\n{system}\n" if system else ""
                past = sysinfo.recent_history(5)
                known += f"\nRecent Co-opilot tasks (oldest first; inspect history for more):\n{past}\n" if past else ""
            transcript = meetings.for_question(task, tools.memory.last_account)
            mode = ("\nFor desktop tasks, try Tab, Shift+Tab, arrows, and Enter before mouse coordinates. "
                    "After each Tab, use focus_info. If a click doesn't work, try another control or the keyboard. "
                    "Do every step yourself; use ask only for a decision, never to hand a step to the user.\n")
            now = time.strftime("%A, %Y-%m-%d %H:%M")
            text = f"{head}{known}\n{tools.memory.block()}{transcript}{mode}Now: {now}\nTask:\n{task}\n{seen}"
        else:
            report["screen"] = seen
            report["task"] = task[:300]
            report["reply_with"] = "only the fenced json actions for the next step of the task; do not describe the screenshot"
            text = _fit(report)
        log(f"Step {round_no}: sending{' with screenshot' if jpeg else ''}")
        reply = chat.send(text, jpeg, ready=_complete)
        _log_reply(log, reply)
        actions = protocol.parse_actions(reply.text, reply.code_blocks)
        if actions is None:
            # A finished sentence is the answer. A promise to act gets one reminder, then we say so.
            # Prose is only a final answer to the first message, before any screenshot or tool.
            # Mid-task, Copilot sometimes describes the screenshot instead of acting; that is not an answer.
            first_reply = report is None and jpeg is None and not tried
            prose = user_reply(reply.text) if first_reply else None
            if prose and _REFUSE.search(prose) and not tools.nudged:
                tools.nudged = True
                log("Refused without trying a tool. Nudging once.")
                report = {"ran": [], "ok": False, "error": _NUDGE, "results": []}
                continue
            if prose:
                answer(prose)
                status("Done")
                return
            log("No action in the reply. Asking once for JSON.")
            first = reply.text
            reply = chat.send(protocol.REMINDER + "\nTask:\n" + task, ready=_complete)
            _log_reply(log, reply)
            actions = protocol.parse_actions(reply.text, reply.code_blocks)
            if actions is None:
                log("No action after the reminder.")
                if first_reply:
                    answer(user_reply(reply.text) or user_reply(first) or _CANT)
                else:
                    answer("I got stuck partway through that and stopped. Nothing else was done; try asking again.")
                status("Done")
                return
        report = {"ran": [], "ok": True, "error": None, "results": []}
        acted = False    # a mouse or keyboard action ran in this batch
        for action in actions:
            if stop.is_set():
                raise copilot_chat.Stopped()
            log("Action: " + json.dumps(action))
            if action["action"] == "done":
                summary = action["summary"] or "Done."
                if _REFUSE.search(summary) and not tried and not report["ran"] and not tools.nudged:
                    tools.nudged = True
                    report.update(ok=False, error=_NUDGE)
                    log("Refused without trying a tool. Nudging once.")
                    break
                if acted and not tools.verified:
                    # ponytail: one extra round per screen task, so "typed it" is checked against the screen.
                    tools.verified = True
                    want_shot = True
                    report["note"] = _VERIFY
                    log("Checking the screen before done")
                    break
                if account.status_only(summary):
                    fetched = any(item.get("topic") in ("account", "inspect") for item in report["results"])
                    if not tools.retried:
                        tools.retried = True
                        if not fetched:
                            hit = account.match(task, tools.memory.last_account, prior)
                            if hit:
                                tools.memory.remember(hit)
                                answer(account.run(hit, chat))
                                status("Done")
                                return
                        report.update(ok=False, error=(
                            "That is not an answer. You already have the result in this report; put those facts in done."
                            if fetched else
                            "That is not an answer. Use account for meetings, mail, people, tasks, or files, then done "
                            "with the actual result. Do not say you are checking or looking something up."
                        ))
                        log("Rejected a status-only answer")
                        break
                answer(summary)
                status("Done")
                return
            if action["action"] == "look":
                want_shot = True
                report["ran"].append(action)
                tried = True
                continue
            if action["action"] in _NEEDS_COORDS and keyboard_navigation < 2:
                report.update(ok=False, error=(
                    "Keyboard navigation is required before mouse coordinates. Use key Tab or Shift+Tab "
                    "and Enter or Space, then focus_info. Mouse actions are available after two keyboard steps."
                ))
                log("Blocked mouse action: keyboard navigation has not been tried")
                break
            if action["action"] in _NEEDS_COORDS and display_target and display_focus and display_focus != display_target.group(1):
                report.update(ok=False, error=(
                    f"Keyboard focus is on Display {display_focus}; the requested target is Display {display_target.group(1)}. "
                    "Use Tab or Shift+Tab, call focus_info, then press Enter on the target."
                ))
                log("Blocked mouse action while the target display is reachable by keyboard")
                break
            if action["action"] in _NEEDS_COORDS and shot is not None:
                target = desktop._screen_xy(action["x"], action["y"], shot)
                if last_mouse_target and abs(target[0] - last_mouse_target[0]) < 24 and abs(target[1] - last_mouse_target[1]) < 24:
                    report.update(ok=False, error="That click was already tried. Try another control or the keyboard; do the step yourself.")
                    log("Blocked a repeated mouse target")
                    break
            verdict = safety.review(action, desktop.foreground_title())
            if verdict[0] == "block":
                report.update(ok=False, error=f"Blocked by the safety check, not run: {verdict[1]}")
                log(f"Blocked: {verdict[1]}")
                break
            if verdict[0] == "confirm" and not tools.confirm(verdict[2], verdict[1]):
                report.update(ok=False, error=f"The user declined: {verdict[2].splitlines()[0]} Nothing was done.")
                log(f"Declined: {verdict[2].splitlines()[0]}")
                break
            visible = action["action"] in protocol.VISIBLE
            if visible and shot is None and action["action"] in _NEEDS_COORDS:
                report.update(ok=False, error=f"{action['action']} needs coordinates from a screenshot; use look first.")
                break
            if visible:
                want_shot = True
                if action["action"] == "key" and action["key"] in ("tab", "shift+tab", "enter", "space", "up", "down", "left", "right"):
                    keyboard_navigation += 1
                if action["action"] in _NEEDS_COORDS and shot is not None:
                    last_mouse_target = desktop._screen_xy(action["x"], action["y"], shot)
                if not hands_on:
                    hands_on = True
                    tools.hands_on()
                    time.sleep(0.4)
            try:
                if action["action"] == "focus_info":
                    output = desktop.focused_control()
                    focused_display = re.search(r"focused:\s*ListItem\s+Display\s+(\d+)", output, re.I)
                    display_focus = focused_display.group(1) if focused_display else display_focus
                    report["results"].append({"topic": "focus_info", "output": output})
                    log(output)
                elif action["action"] == "inspect":
                    output = sysinfo.collect(action["topic"])
                    report["results"].append({"topic": action["topic"], "output": output})
                    log(f"inspect {action['topic']}: {len(output)} chars")
                elif action["action"] in _SKILLS:
                    status(f"Step {round_no}: {action['action']}")
                    output = _skill(action, tools)
                    if output is not None:
                        report["results"].append({"topic": action["action"], "output": output})
                        log(f"{action['action']}: {len(str(output))} chars")
                else:
                    if action["action"] == "key" and keyboard_window:
                        if keyboard_window.lower() not in desktop.foreground_title().lower():
                            desktop.focus_window(keyboard_window)
                            log(f"Focused {keyboard_window} before the key")
                    output = desktop.run(action, shot)
                    if action["action"] == "open_url" and str(action.get("url", "")).startswith("ms-settings:"):
                        keyboard_window = "Settings"
                    if output is not None:
                        report["results"].append({"topic": action["action"], "output": output})
                    time.sleep(_SETTLE)
                report["ran"].append(action)
                tried = True
                acted = acted or visible
            except pyautogui.FailSafeException:
                raise
            except Exception as exc:
                report.update(ok=False, error=f"{action['action']}: {exc}")
                log(f"Action failed: {exc}")
                break
    log("Round cap reached")
    answer("I wasn't able to finish that. Ask for a smaller step and I'll try again.")
    status("Done")


_SKILLS = {"list_dir", "read_file", "write_file", "ado", "say", "ask", "account"}
_NEEDS_COORDS = {"click", "double_click", "right_click", "move", "drag", "scroll"}


def _local_open(task, log, status, answer):
    """Open Settings directly when the request names nothing else."""
    request = re.sub(r"[?.!]+$", "", task.strip().lower())
    if re.fullmatch(r"(?:can you |please )?open (?:windows )?settings(?: please)?", request):
        uri, label = "ms-settings:", "Settings"
    elif re.fullmatch(r"(?:can you |please )?open (?:windows )?display settings(?: please)?", request):
        uri, label = "ms-settings:display", "Display Settings"
    else:
        return False
    log(f"Local action: opening {uri}")
    desktop.open_url(uri)
    answer(f"Opened {label}.")
    status("Done")
    return True


def _skill(action, tools):
    """Non-desktop skills. Returns text for Copilot, or None."""
    kind = action["action"]
    if kind == "say":
        if account.status_only(action["text"]):
            return None
        tools.say(action["text"])
        return None
    if kind == "ask":
        answer = tools.ask(action["question"], action["options"], action["multi"], action["allow_text"])
        return json.dumps({"question": action["question"], "answer": answer or "no answer (skipped)"})
    if kind == "account":
        tools.memory.remember(action)
        return account.run(action, tools.chat)
    if kind == "list_dir":
        return files.list_dir(action["path"])
    if kind == "read_file":
        return files.read_file(action["path"], action.get("sheet"))
    if kind == "write_file":
        target = files.resolve(action["path"])
        if action["mode"] == "overwrite" and target.exists() and not tools.confirm(f"Overwrite {target}?"):
            return "The user declined to overwrite the file. Nothing was written."
        return files.write_file(target, action.get("content"), action.get("rows"), action["mode"])
    if kind == "ado":
        if action["op"] == "import_excel":
            rows, preview = ado.import_preview(action)
            if not tools.confirm(f"Create these in Azure DevOps?\n{preview}"):
                return "The user declined the import. Nothing was created.\n" + preview
            action = dict(action, rows=rows)
        return tools.ado.run(action)
    raise ValueError(f"unknown skill: {kind}")


def _facts_text(facts, stop):
    """The PC summary, waiting briefly if the background collect has not finished.

    ponytail: at most 0.6s, and only on the first Copilot prompt. A late summary is still
    available through inspect system.
    """
    if not isinstance(facts, dict):
        return facts or ""
    for _ in range(4):
        got = facts.get("system") or ""
        if got or stop.is_set():
            return got
        time.sleep(0.15)
    return facts.get("system") or ""


def _fit(report, limit=12000):
    """Halve the longest inspect output until the report fits one chat message."""
    text = json.dumps(report)
    while len(text) > limit and report["results"]:
        longest = max(report["results"], key=lambda item: len(item["output"]))
        if len(longest["output"]) < 200:
            break
        longest["output"] = longest["output"][: len(longest["output"]) // 2] + "\n...(cut)"
        text = json.dumps(report)
    return text


def _complete(reply):
    return protocol.parse_actions(reply.text, reply.code_blocks, complete=True) is not None


def _check():
    mem = Memory()
    assert mem.block() == ""
    mem.add("user", "whens my next meeting")
    mem.add("assistant", "Your next meeting is Budget tomorrow at 08:00.")
    assert "User: whens my next meeting" in mem.block()
    assert "Co-opilot: Your next meeting is Budget tomorrow at 08:00." in mem.block()
    mem.remember({"op": "next"})
    mem.clear()
    assert mem.block() == "" and mem.last_account is None
    assert user_reply("I can't book a flight from here.") == "I can't book a flight from here."
    assert user_reply("I'll open Notepad now.") is None
    assert user_reply("Checking your next meeting.") is None
    assert user_reply("Sure.") is None
    assert _REFUSE.search("I can’t access your meeting transcripts from the information available in this chat.")
    assert _REFUSE.search("Sorry, I don't have access to your calendar.")
    assert not _REFUSE.search("Your PC cannot reach the VPN because the adapter is disabled.")
    assert _facts_text({"system": "PC"}, type("S", (), {"is_set": lambda self: False})()) == "PC"
    assert _local_open("what time is it", lambda *_: None, lambda *_: None, lambda *_: None) is False
    print("loop ok")


def _log_reply(log, reply):
    text = " ".join(reply.text.split())
    if len(text) > 500:
        text = text[:500] + "..."
    log("Copilot: " + (text or "(no text)"))
    if reply.images:
        size = sum(len(item) for item in reply.images)
        log(f"image received ({len(reply.images)}, {size} bytes)")


if __name__ == "__main__":
    _check()
