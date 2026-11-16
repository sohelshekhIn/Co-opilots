"""One Copilot session for the app's lifetime. Each reply is a batch of up to 5 actions; up to 40 rounds per task."""

import json
import queue
import re
import threading
import time

import pyautogui

import ado
import copilot_chat
import desktop
import files
import protocol
import safety
import sysinfo

MAX_ROUNDS = 40
_SETTLE = 0.35
NEW_CHAT = object()


class Tools:
    """What skills need besides the action itself: the chat's browser, and ways to talk to the user."""

    def __init__(self, chat, log, stop, say, ask, hands_on=None):
        self.chat, self.log, self.stop = chat, log, stop
        self.say = say or (lambda text: log("Note: " + text))
        self._ask = ask
        self.hands_on = hands_on or (lambda: None)
        self.ado = ado.Ado(chat)
        self.allowed = set()

    def ask(self, question, options=(), multi=False, allow_text=True):
        """Show a question card and wait for the answer. Returns {"choices": [...], "text": "..."} or None."""
        if self._ask is None:
            return None
        reply = queue.Queue()
        self._ask({"question": question, "options": list(options), "multi": multi, "allow_text": allow_text}, reply)
        while True:
            if self.stop.is_set():
                raise copilot_chat.Stopped()
            try:
                return reply.get(timeout=0.2)
            except queue.Empty:
                self.chat.idle(0.2)

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
            say=None, ask=None, hands_on=None):
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
        # ponytail: one chat for every task; the instruction goes in once. Upgrade path is a
        # New chat button that sets instructed back to False if the thread gets too long.
        instructed = False
        while True:
            try:
                task = tasks.get(timeout=0.2)
            except queue.Empty:
                chat.idle(1.0)
                continue
            if task is None:
                return
            if task is NEW_CHAT:
                log("Starting a new Copilot chat")
                chat._new_chat()
                chat._go_temporary()
                status("Selecting GPT model")
                if not chat._select_model():
                    raise RuntimeError(f"Could not verify {copilot_chat.PREFERRED_MODEL} in Microsoft Copilot")
                chat._wait_until_ready(timeout=60)
                instructed = False
                status("Ready")
                on_ready()
                continue
            task_id, task = task
            stop.clear()
            tools.allowed.clear()
            started = time.time()
            outcome = {"answer": "Not finished"}

            def answered(text):
                outcome["answer"] = text
                if answer:
                    answer(task_id, text)
                else:
                    log(text)

            try:
                if _local_open(task, log, status, answered):
                    pass
                else:
                    _drive(task, chat, log, status, stop, instructed, facts.get("system", ""), answered, tools)
                instructed = True
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
            sysinfo.record_history(task, outcome["answer"], time.time() - started)
            log(f"Task took {time.time() - started:.0f}s")
            on_idle(task_id)
    except copilot_chat.Stopped:
        status("Stopped")
    except Exception as exc:
        log(f"Error: {exc}")
        status("Error")
    finally:
        chat.close()


def _drive(task, chat, log, status, stop, instructed, system_facts, answer, tools):
    report = None
    shot = None          # last screenshot Copilot saw; clicks are mapped against it
    want_shot = False    # set by look, or after mouse/keyboard actions
    hands_on = False
    needs_verification = False
    mouse_actions = 0
    mouse_limit = 3
    handoff_requested = False
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
            needs_verification = False
            seen = f"Screenshot attached: {shot.sent_w} x {shot.sent_h} pixels. Monitors: {shot.layout()}."
        else:
            seen = "No screenshot attached (use look if you need to see the screen)."
        want_shot = False
        if report is None:
            head = protocol.NEXT_TASK if instructed else protocol.INSTRUCTION
            facts = ""
            if not instructed:
                facts = f"\nThis PC:\n{system_facts}\n" if system_facts else ""
                past = sysinfo.recent_history(5)
                facts += f"\nRecent Copilot Assistant tasks (oldest first; inspect history for more):\n{past}\n" if past else ""
            mode = ("\nFor desktop tasks, try direct actions and keyboard shortcuts first. Use screenshot coordinates only "
                    "when keyboard navigation cannot reach the control. After at most three mouse actions, ask the user "
                    "to perform one specific visible step, then look again to verify. Never repeat uncertain clicks.\n")
            text = f"{head}{mode}{facts}\nTask:\n{task}\n{seen}"
        else:
            report["screen"] = seen
            text = _fit(report)
        log(f"Step {round_no}: sending{' with screenshot' if jpeg else ''}")
        reply = chat.send(text, jpeg, ready=_complete)
        _log_reply(log, reply)
        actions = protocol.parse_actions(reply.text, reply.code_blocks)
        if actions is None and report is None and _is_greeting(task) and not _false_capability(reply.text):
            answer(reply.text.strip())
            status("Done")
            return
        # Microsoft Copilot sometimes answers as a standalone chat assistant,
        # ignoring the action bridge. Retry at most twice, including JSON
        # refusals, before returning a clear error to the user.
        for attempt in range(2):
            refusal = _false_capability(reply.text) or (
                actions and len(actions) == 1 and actions[0]["action"] == "done"
                and _false_capability(actions[0]["summary"])
            )
            if actions is not None and not refusal:
                break
            reminder = protocol.CAPABILITY_REMINDER if refusal else protocol.REMINDER
            log("Copilot refused the PC action; retrying action protocol" if refusal else
                "No action in the reply; retrying JSON protocol")
            reply = chat.send(reminder, ready=_complete)
            _log_reply(log, reply)
            actions = protocol.parse_actions(reply.text, reply.code_blocks)
        if actions is None or (len(actions) == 1 and actions[0]["action"] == "done"
                               and _false_capability(actions[0]["summary"])):
            log("Copilot did not provide a usable PC action after retries")
            status("Error")
            raise RuntimeError("Microsoft Copilot declined to provide a PC action after two retries")
        report = {"ran": [], "ok": True, "error": None, "results": []}
        for action in actions:
            if stop.is_set():
                raise copilot_chat.Stopped()
            log("Action: " + json.dumps(action))
            if action["action"] == "done":
                if needs_verification:
                    report.update(ok=True, error=None)
                    report["results"].append({"topic": "verification", "output":
                        "The screen changed after the last screenshot. Inspect the next screenshot before claiming completion."})
                    log("Deferring done until the changed screen is verified")
                    want_shot = True
                    break
                answer(action["summary"] or "Done.")
                status("Done")
                return
            if action["action"] == "look":
                want_shot = True
                report["ran"].append(action)
                continue
            if action["action"] in _NEEDS_COORDS and keyboard_navigation < 2:
                report.update(ok=False, error=(
                    "Keyboard navigation is required before mouse coordinates. Use key Tab/Shift+Tab "
                    "and Enter/Space, then look to check focus. Mouse actions are available only after "
                    "at least two keyboard navigation steps."
                ))
                log("Blocked mouse action: keyboard navigation has not been tried")
                break
            if action["action"] in _NEEDS_COORDS and display_target and display_focus and display_focus != display_target.group(1):
                report.update(ok=False, error=(
                    f"Keyboard focus is on Display {display_focus}; the requested target is Display {display_target.group(1)}. "
                    "Use Tab or Shift+Tab through the display list, call focus_info, then press Enter on the target."
                ))
                log("Blocked mouse action while target display is keyboard reachable")
                break
            if action["action"] in _NEEDS_COORDS and shot is not None:
                target = desktop._screen_xy(action["x"], action["y"], shot)
                if last_mouse_target and abs(target[0] - last_mouse_target[0]) < 24 and abs(target[1] - last_mouse_target[1]) < 24:
                    report.update(ok=False, error="Repeated mouse target did not work. Use keyboard navigation or ask the user for the specific click.")
                    log("Blocked repeated mouse target")
                    break
            if action["action"] in _NEEDS_COORDS and mouse_actions >= mouse_limit:
                report.update(ok=False, error=(
                    "Mouse action limit reached. Do not request more coordinates. "
                    "If the last click focused the intended control, try key Enter or Space now, then look to verify. "
                    "If focus is uncertain or the keyboard action fails, use ask to tell the user exactly which "
                    "visible control to click. After their response, use look to verify."
                ))
                log("Mouse limit reached; requesting keyboard activation or user handoff")
                if handoff_requested:
                    answer("Stopped after the mouse action limit. Please complete the remaining step in the open window.")
                    status("Stopped")
                    return
                handoff_requested = True
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
                needs_verification = True
                if action["action"] == "key" and action["key"] in ("tab", "shift+tab", "enter", "space", "up", "down", "left", "right"):
                    keyboard_navigation += 1
                if action["action"] in _NEEDS_COORDS and shot is not None:
                    mouse_actions += 1
                    last_mouse_target = desktop._screen_xy(action["x"], action["y"], shot)
                    log(f"Desktop target: screenshot ({action['x']}, {action['y']}) -> screen {desktop._screen_xy(action['x'], action['y'], shot)}")
                if not hands_on:
                    hands_on = True
                    tools.hands_on()
                    time.sleep(0.4)
            try:
                if action["action"] == "focus_info":
                    output = desktop.focused_control()
                    focused_display = re.search(r"focused:\s*ListItem\s+Display\s+(\d+)", output, re.I)
                    display_focus = focused_display.group(1) if focused_display else None
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
                            log(f"Focused {keyboard_window} before keyboard action")
                    output = desktop.run(action, shot)
                    if action["action"] == "open_url" and action["url"].startswith("ms-settings:"):
                        keyboard_window = "Settings"
                    if output is not None:
                        report["results"].append({"topic": action["action"], "output": output})
                    time.sleep(_SETTLE)
                report["ran"].append(action)
            except pyautogui.FailSafeException:
                raise
            except Exception as exc:
                report.update(ok=False, error=f"{action['action']}: {exc}")
                log(f"Action failed: {exc}")
                break
    log("Round cap reached")
    status("Stopped")


_SKILLS = {"list_dir", "read_file", "write_file", "ado", "say", "ask"}
_NEEDS_COORDS = {"click", "double_click", "right_click", "move", "drag", "scroll"}


def _local_open(task, log, status, answer):
    """Handle unambiguous Settings launch requests without a model round trip."""
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


def _false_capability(summary):
    text = summary.lower()
    no_access = ("can't", "cannot", "can’t", "unable", "don't have", "do not have")
    pc_action = ("control", "operate", "execute", "open apps", "change your display",
                 "change display", "pc-control", "windows pc", "your pc")
    return any(word in text for word in no_access) and any(word in text for word in pc_action)


def _is_greeting(task):
    text = task.strip().lower().rstrip("?!. ")
    return bool(re.fullmatch(r"(?:hi|hello|hey|good morning|good afternoon|good evening)"
                             r"(?:,? how are you)?", text))


def _skill(action, tools):
    """Non-desktop skills. Returns text for Copilot, or None."""
    kind = action["action"]
    if kind == "say":
        tools.say(action["text"])
        return None
    if kind == "ask":
        answer = tools.ask(action["question"], action["options"], action["multi"], action["allow_text"])
        return json.dumps({"question": action["question"], "answer": answer or "no answer (skipped)"})
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


def _log_reply(log, reply):
    text = " ".join(reply.text.split())
    if len(text) > 500:
        text = text[:500] + "..."
    log("Copilot: " + (text or "(no text)"))
    if reply.images:
        size = sum(len(item) for item in reply.images)
        log(f"image received ({len(reply.images)}, {size} bytes)")
