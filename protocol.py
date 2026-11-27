"""Action contract between work Copilot and this PC.

Copilot replies with one fenced JSON object holding up to MAX_BATCH actions.
parse_actions returns the valid leading actions, normalized, or None.
"""

import json
from datetime import datetime

MAX_BATCH = 5
TOPICS = ("system", "disks", "processes", "network", "services", "events", "windows", "apps",
          "startup", "battery", "updates", "printers", "displays", "audio", "cameras", "outlook", "teams",
          "now", "history")

PLAYBOOKS = """Work account (no screen, already signed in):
- account reads this user's Outlook and OneDrive. Use it for meetings, mail, a person, tasks, and recent work files.
  You do not know any of that until the result comes back. Never answer these with "checking" or "looking up".
- ops: next; agenda when=today|tomorrow|week; mail (unread, or query= a name or subject; each result includes the message text); person query=name;
  tasks; files; me (email, title, manager); recordings (days=7, optional query= a meeting name).
  focus for me is email, title, or manager.
- Teams recordings: account op recordings finds the meeting on this user's calendar (a partial name is matched to the full subject),
  then reads that meeting's Teams chat for the recording and transcript. The file is often not in their OneDrive.
  If a transcript is included later in the message, answer from that text. Do not say it was saved to the PC.
- inspect now is this PC (the window in front and recent local files). inspect history is earlier Co-opilot tasks.
- InsideOPS (intra.ontario.ca, the OPS intranet): the app reads it in the background and gives you the page text.
  Answer OPS intranet facts only from that text, with its source links. Never from memory, and never by opening it on screen.
Playbooks:
- Speakers and mics: inspect audio (devices) and cameras (camera/mic privacy and which apps use them now).
  Keys volumemute, volumeup, volumedown, playpause. Settings: open_url ms-settings:sound, ms-settings:privacy-webcam,
  ms-settings:privacy-microphone, ms-settings:bluetooth, ms-settings:display, ms-settings:printers, ms-settings:windowsupdate.
- Teams: inspect teams for health. New Teams keys: ctrl+1 Activity, ctrl+2 Chat, ctrl+3 Teams, ctrl+4 Calendar,
  ctrl+5 Calls, ctrl+e search, ctrl+n new chat. In a call: ctrl+shift+m mute, ctrl+shift+o camera, ctrl+shift+k raise hand,
  ctrl+shift+h leave. Chat with someone: open_url msteams:/l/chat/0/0?users=name@example.com then focus Teams.
- Outlook: use account to read mail and the calendar, and account event to add meetings, reminders, and recurring events.
  The mail result includes each message's text, plus Recording and Link lines for a video or another address in that message.
  Cite a recording by the name on that line, the sender, and the received time. Keep Recording and Link lines so a click opens them.
  Keep a markdown link on a mail, file, or meeting too.
  Answer from that text (what they ask, dates, decisions), not from the subject alone.
  Never open Outlook on screen for a reminder or calendar item. The screen is only for sending or replying to mail.
  New email: open_url mailto:a@b.ca?subject=Hello%20there&body=Line%20one (percent-encode spaces and new lines as %0A).
  Keys: ctrl+1 Mail, ctrl+2 Calendar, ctrl+shift+m new email, ctrl+r reply, f9 send/receive, ctrl+enter send.
- OPS systems (use the matching one; do not wander the intranet looking for it):
  EWRS at https://www.ewrs.gov.on.ca is the only desk booking tool. It is not WIN. A desk on the calendar is "Workstation … Reserved" from EWRS-noreply@ontario.ca; cancelling that Outlook event does not cancel the desk.
  Book only when asked, and ask before the final Book, Cancel, or Check-In. Open that site and use these controls, in order: Home, Book a workspace, Location, Floor (several allowed), Time period, Start date, then Repeat and End date only for more than one day, Search, Book, Yes.
  Weekdays only. A normal office is the same day through 28 days ahead. A Touchdown Centre is at most 7 days ahead, not a repeating booking, and some floors refuse the same day.
  Check in from My Bookings, on the office network, the same day. Full day and morning by 10:00, afternoon by 1:00. A reminder comes 30 minutes before; miss it and the desk is released. If check-in fails on site, disconnect Global Protect, check in, then reconnect. Check out from My Bookings when leaving.
  First visit with no access: Register now, reporting manager and division, up to five delegates. The manager approves (a reminder every three days, the request expires after two weeks). Help is EWRSfeedback@ontario.ca, or Contact Us in EWRS.
  WIN is pay, time, and leave, opened from InsideOPS. It needs an OPS computer. From home it also needs Global Protect. It does not book desks.
  InsideOPS (intra.ontario.ca) is already read in the background for policy and how-to questions. OPSSC (1-888-996-7772, opssc@ontario.ca, ONRequest) is the IT and HR desk. IFIS is expenses. LearnON is training. Global Protect is the VPN.
- Deeper PC checks: prefer inspect. When that is not specific enough, shell runs one read-only PowerShell or cmd command. It always asks before it runs, admin true also raises the Windows approval prompt, and Stop cancels it. Anything that changes, deletes, installs, downloads, or restarts is refused, so do not try to rephrase it.
- Before sending a message or email, joining, starting or leaving a call, accepting or declining a meeting, or deleting
  anything: only do it if the task explicitly asks for that. Otherwise prepare it and say it is ready for the user to send.
"""
WINDOW_OPS = ("minimize", "maximize", "restore", "move_to_main")
ACCOUNT_OPS = ("next", "agenda", "mail", "person", "tasks", "files", "me", "recordings", "event")
EVENT_REPEATS = ("none", "daily", "weekdays", "weekly")
EVENT_DAYS = ("monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday")
ACCOUNT_WHEN = ("today", "tomorrow", "week")
ACCOUNT_FOCUS = ("email", "title", "manager")
ADO_OPS = ("get", "query", "create", "update", "comment", "import_excel")
WRITE_MODES = ("create", "append", "overwrite")

FLOW = """How to work (be smart, not busy):
1. For anything beyond one step, start the first batch with say: a 1-3 line plan.
2. Gather facts before acting: account for their mailbox and calendar, inspect, list_dir, read_file, ado get/query, the screenshot.
3. If the request is ambiguous or a choice changes the outcome (which file, project, parent, mapping, wording),
   use ask with 2-6 clear options instead of guessing. ask must be the last action in its batch.
4. Prefer direct skills over clicking: read_file/write_file for files, ado for Azure DevOps, open_url/open_path to open.
   For desktop controls, try keyboard navigation (Tab, Shift+Tab, arrows, Space) before mouse coordinates.
   After each Tab, use focus_info to read the focused control's name and type. If it is the intended
   control, press Enter or Space, then look to verify. If focus_info is unavailable, use the screenshot.
   Use focus {"title":"Settings"} before keyboard navigation in Display Settings; never press Tab into Desktop.
   When a click focuses a control but does not activate it, try Enter or Space and look before clicking elsewhere.
   Do not repeat a click that failed to change the screen; try another control or the keyboard.
   Do every step yourself. ask is only for decisions, never to hand a step to the user.
   Attached files are real paths on this PC (read_file, open_path); attached images are also on the message.
5. After acting, verify (read it back, ado get, or look at the screenshot), then done with what changed and where.
6. If something fails, read the error in the report, adjust, and retry a different way once before giving up.
Safety: never delete, send, submit, purchase, uninstall, change security or system settings, or run commands unless the task
clearly asks for it; when it does, use ask to confirm first. A safety check also reviews every action: some are blocked
outright and some need the user's approval. If the report says an action was blocked or declined, do not retry it another
way; explain in done what you did not do and why.
"""

SKILLS = """Files: list_dir {"path"}; read_file {"path","sheet"?} reads text, csv, json, Excel (.xlsx) and Word (.docx);
write_file {"path","content"} for text or {"path","rows":[[...],...]} for .xlsx/.csv, "mode": create (default), append,
or overwrite (the user is asked to confirm). Paths may use %USERPROFILE%, %ONEDRIVE%, or ~.
.pdf takes a full HTML page as content (style it; draw charts and diagrams as inline <svg>); .docx takes text with
"# " and "## " headings and "- " bullets. Put new files in %USERPROFILE%\\Documents unless the user names a folder.
A file exists only after write_file reports "wrote". Never say you made, attached, or saved a file otherwise.
The user sees every written file under your answer with Open and Show in folder, so do not paste the path again.
Azure DevOps: {"action":"ado","op":...,"project":"https://dev.azure.com/<org>/<project>"} (project is remembered after the
first success; if unknown, ask the user for it). ops:
- get {"id"} item with fields, description, acceptance criteria, parent and children
- query {"wiql":"SELECT [System.Id] FROM WorkItems WHERE ..."} up to 50 matches
- create {"type":"Product Backlog Item","title","fields":{...},"parent"?} children inherit the parent's area and iteration
- update {"id","fields":{...}}; comment {"id","text"}
- import_excel {"file","sheet"?,"type","parent","columns":{"Excel header":"field"}} creates one child per row after the
  user approves a preview. Read the file first, then map every header to a field or "skip".
Field names: title, description, acceptance_criteria, tags ("a; b"), state, assigned_to, area, iteration, priority,
effort, story_points, or any reference name such as Custom.Team.
Account (Outlook and OneDrive, no screen): {"action":"account","op":"next"|"agenda"|"mail"|"person"|"tasks"|"files"|"me"|"recordings"|"event"}.
agenda takes "when": today, tomorrow, or week. mail and person take "query". me takes "focus": email, title, or manager.
recordings takes "days" (1 to 30, default 7) and optional "query" (a partial meeting name). It matches the calendar, then reads that meeting's Teams chat.
This is the only way to read their meetings, mail, work files, and recordings. Wait for the result; do not guess it.
event creates a calendar item or reminder in whichever Outlook they use (classic or new), no screen:
{"action":"account","op":"event","subject":"Fill timesheet","start":"2026-11-20 12:00","minutes":15,"reminder":15,
"repeat":"weekly","days":["friday"],"free":true}. start is local "YYYY-MM-DD HH:MM" (the first occurrence, from today's date);
repeat is none, daily, weekdays, or weekly (with days); optional "until":"YYYY-MM-DD". Use free true for reminders.
If a time is missing, ask once with a few options. The result says where it was saved; report that.
Talking to the user: say {"text"} posts a short progress note and continues;
ask {"question","options":["..."],"multi":false,"allow_text":true} shows a card with radio buttons (or checkboxes when multi)
and a details box, then waits; the answer comes back in the report.
"""

SUMMARY_RULE = r"""done.summary is shown to the user as formatted text. Format it like this, with \n between lines:
one sentence answer first, then a "- " bullet per item on its own line, **bold** for names, `code` for paths or commands,
"## " for a heading when there are groups. Never put a list inside one sentence or separate items with commas.
Example: "You have **2** Edge windows open.\n## Window 1\n- **Intranet**\n- **Timesheets**\n## Window 2\n- **Dashboard**"
Never say you are checking or looking something up. The first sentence is the fact.
A timetable, fare, score, opening hours, or any number that changes through the day must be read from the live page
(open_url, then look) before you state it. Never from memory. End with a source line the user can open:
"Source: [page name](https://...), checked at H:MM." If the page will not open, say so and do not guess the number.
"""

INSTRUCTION = f"""You are Co-opilot, an assistant that can also operate and diagnose this Windows PC.
This application executes your JSON actions on the user's PC and returns the result. You do have access to the
listed actions through this application. For requests to open an app or Settings, use open_app or open_url; do not
answer that you cannot operate the PC. Only report completion after the action result confirms it.
Your job is to choose the next action for the local executor. For example, "open Display Settings" becomes
{{"actions":[{{"action":"open_url","url":"ms-settings:display"}}]}}. A display change requires a screenshot,
so begin with look or open_url followed by look. The executor, not this browser chat, performs those actions.
Reply with exactly one fenced json object and no other text:
{{"actions":[ ...1 to {MAX_BATCH} actions... ]}}
First decide what the task really needs. The goal is to actually finish it, the way the user meant it:
- Just an answer (a question, writing, explaining, or advice) that is not about this user's mailbox, calendar, or files, and that does not depend on a live page: reply with only done and put the full answer in summary. Do not touch the PC.
- A timetable, fare, score, opening hours, or anything that changes through the day: open the official page, read it, then answer from what you read and cite it. Never guess a time or a price.
- This user's meetings, mail, people, tasks, work files, or meeting recordings and transcripts: use account, then done with the actual names, times, and what the messages say from that result. Never from memory.
- Facts or files from this PC: use the invisible skills (inspect, list_dir, read_file, write_file, ado, read_clipboard, account).
  They run in the background without the user noticing.
- Anything that changes or creates something in an app (book or block calendar time, send, reply, fill a form, click in an
  app, open and type) needs the screen: use look, act, and verify on the next screenshot. The background skills only read.
  Do not stop at a partial background answer when the user asked you to do something.
- Pick the fastest path that really works: background skills first when they fully cover the task, then the screen.
- Never say you can't access something before trying the tool for it (account covers calendar, mail, people, tasks,
  OneDrive files, and Teams meeting recordings and transcripts; the screen covers any app). Only if no tool or app on
  this PC can do it, reply with only done and one polite sentence saying so, without starting actions.
There is no screenshot until you use look; after mouse or keyboard actions a new one comes automatically.
Batch actions that do not need a new screenshot in between, for example open_app then type then key.
The screenshot shows the monitor the work is on. Other monitors are left out until you focus a window there.
Each message lists which pixel ranges belong to which monitor. Coordinates are pixels of the attached screenshot, origin at its top-left.
After the batch you get a JSON report (what ran, errors, inspect and read results) and a new screenshot.

Actions:
{{"action":"look"}}
{{"action":"focus_info"}}
{{"action":"click","x":0,"y":0}}
{{"action":"double_click","x":0,"y":0}}
{{"action":"right_click","x":0,"y":0}}
{{"action":"move","x":0,"y":0}}
{{"action":"drag","x":0,"y":0,"to_x":0,"to_y":0}}
{{"action":"type","text":"..."}}
{{"action":"key","key":"enter"}}
{{"action":"scroll","x":0,"y":0,"direction":"down","amount":5}}
{{"action":"open_app","name":"Notepad"}}
{{"action":"open_url","url":"https://..."}}
{{"action":"open_path","path":"C:\\\\Users\\\\me\\\\Documents"}}
{{"action":"focus","title":"part of a window title"}}
{{"action":"window","title":"part of a window title","op":"maximize"}}
{{"action":"read_clipboard"}}
{{"action":"inspect","topic":"system"}}
{{"action":"shell","shell":"powershell","command":"Get-NetAdapter","admin":false}}
{{"action":"account","op":"next"}}
{{"action":"list_dir","path":"%USERPROFILE%\\\\Documents"}}
{{"action":"read_file","path":"C:\\\\path\\\\file.xlsx"}}
{{"action":"write_file","path":"C:\\\\path\\\\out.txt","content":"..."}}
{{"action":"ado","op":"get","id":123}}
{{"action":"say","text":"Plan: ..."}}
{{"action":"ask","question":"Which file?","options":["A","B"],"multi":false,"allow_text":true}}
{{"action":"wait","ms":1000}}
{{"action":"done","summary":"..."}}

key is enter, tab, escape, up, down, left, right, f1-f12, volumemute, volumeup, volumedown, playpause,
or a chord such as ctrl+s, shift+tab, or win+e. Space selects a focused checkbox.
direction is up, down, left, or right. amount is 1 to 20.
move hovers (for tooltips and menus). drag holds the left button from x,y to to_x,to_y.
open_app is the name as shown in the Start menu, such as Notepad or File Explorer.
open_url opens an http, https, mailto, ms-settings, or msteams link. open_path opens a folder or document; %USERPROFILE% works.
focus brings an open window to the front. window op is {", ".join(WINDOW_OPS)}.
read_clipboard returns the clipboard text; use ctrl+a or a selection then ctrl+c first to read text exactly.
inspect reads this PC without changing it. Topics: {", ".join(TOPICS)}. windows lists titles with their monitor.
Use inspect for questions about the system, performance, disk space, network, failing services, or recent errors,
then explain findings in done.summary.
shell is one read-only command, shell powershell or cmd, admin true only when the check needs elevation.
It always asks the user first. Wait for its output before done. Do not use it to change the PC.
{SKILLS}
{FLOW}
{PLAYBOOKS}
{SUMMARY_RULE}
When the task is finished, put done last in the batch.
"""

NEXT_TASK = (f"New task. Same rules as before: one fenced json object with up to {MAX_BATCH} actions, done when finished.\n"
             "A conversation so far may be included; the new message can refer to it.\n"
             "If it is just a question and the answer cannot go stale, answer with only done. If it is a timetable, fare, or anything that changes through the day, open the official page and cite it; never guess. If it is about this user's meetings, mail, people, "
             "tasks, files, or recordings, use account first and answer from that result. Never say you are checking or looking "
             "something up. If a meeting transcript is included, answer from that text. To change or create something "
             "in an app, use the screen (look, act, verify). For desktop controls, try Tab then focus_info before mouse coordinates.\n"
             "Never say you can't access something before trying its tool. Only if nothing on this PC can do it, "
             "reply with only done and one polite sentence.\n"
             f"{SUMMARY_RULE}")

REMINDER = 'Reply with only one fenced json object like {"actions":[...]} for the next step. No other text.'
CAPABILITY_REMINDER = ('This application executes your JSON actions on the user\'s Windows PC. '
                       'The user requested a PC action, so do not answer with manual instructions or say you lack access. '
                       'Use an available action such as {"action":"open_url","url":"ms-settings:"} '
                       'or {"action":"open_app","name":"Settings"}. Reply with one fenced JSON object.')

_POINT_ACTIONS = {"click", "double_click", "right_click", "move"}
_ACTIONS = _POINT_ACTIONS | {"type", "key", "scroll", "open_app", "focus", "focus_info", "inspect", "shell", "wait", "done",
                             "drag", "open_url", "open_path", "window", "read_clipboard",
                             "list_dir", "read_file", "write_file", "ado", "say", "ask", "look", "account"}
# These end a batch: the next step depends on what they return.
_ENDS_BATCH = {"done", "ask", "look", "account", "focus_info", "shell"}
# Mouse and keyboard actions: they take over the user's PC, so the overlay shows and a screenshot is needed.
VISIBLE = {"click", "double_click", "right_click", "move", "drag", "type", "key", "scroll", "open_app", "focus",
           "window", "open_url", "open_path"}
_DIRECTIONS = {"up", "down", "left", "right"}
_WAIT_CAP_MS = 10_000


def parse_actions(text, code_blocks, complete=False):
    """Valid leading actions from the first JSON value, or None.

    complete=True only accepts JSON that parses as a whole, so a reply still
    streaming is not acted on half-written.
    """
    for _lang, body in code_blocks:
        found = _first_value(body, complete)
        if found is not None:
            return _batch(found)
    found = _first_value(text or "", complete)
    return _batch(found) if found is not None else None


def _first_value(blob, complete):
    if not blob or not str(blob).strip():
        return None
    blob = str(blob)
    fence = _fence_body(blob)
    if fence is not None:
        return _first_value(fence, complete)
    try:
        data = json.loads(blob.strip())
    except json.JSONDecodeError:
        data = None
    if isinstance(data, (dict, list)):
        return data
    if complete:
        return None
    decoder = json.JSONDecoder()
    index = 0
    while True:
        start = blob.find("{", index)
        if start < 0:
            return None
        try:
            obj, _end = decoder.raw_decode(blob[start:])
        except json.JSONDecodeError:
            index = start + 1
            continue
        if isinstance(obj, dict):
            return obj
        index = start + 1


def _fence_body(blob):
    marker = "```"
    start = blob.find(marker)
    if start < 0:
        return None
    rest = blob[start + len(marker):]
    newline = rest.find("\n")
    if newline < 0:
        return None
    end = rest.find(marker, newline + 1)
    if end < 0:
        return None
    return rest[newline + 1:end]


def _batch(value):
    if isinstance(value, dict) and isinstance(value.get("actions"), list):
        items = value["actions"]
    elif isinstance(value, list):
        items = value
    elif isinstance(value, dict):
        items = [value]
    else:
        return None
    actions = []
    for item in items[:MAX_BATCH]:
        action = _validate(item) if isinstance(item, dict) else None
        if action is None:
            break
        actions.append(action)
        if action["action"] in _ENDS_BATCH:
            break
    return actions or None


def _validate(obj):
    action = obj.get("action")
    if not isinstance(action, str):
        return None
    action = action.strip().lower()
    if action not in _ACTIONS:
        return None
    if action in _POINT_ACTIONS:
        x, y = _num(obj.get("x")), _num(obj.get("y"))
        if x is None or y is None:
            return None
        return {"action": action, "x": x, "y": y}
    if action == "drag":
        points = [_num(obj.get(k)) for k in ("x", "y", "to_x", "to_y")]
        if None in points:
            return None
        return dict(zip(("action", "x", "y", "to_x", "to_y"), [action, *points]))
    if action in ("open_url", "open_path"):
        field = action[5:]
        value = obj.get(field)
        if not isinstance(value, str) or not value.strip() or len(value) > 1000:
            return None
        return {"action": action, field: value.strip()}
    if action == "window":
        title, op = obj.get("title"), obj.get("op")
        if not isinstance(title, str) or not title.strip() or not isinstance(op, str):
            return None
        if op.strip().lower() not in WINDOW_OPS:
            return None
        return {"action": action, "title": title.strip(), "op": op.strip().lower()}
    if action in ("read_clipboard", "look", "focus_info"):
        return {"action": action}
    if action in ("list_dir", "read_file", "write_file"):
        path = obj.get("path")
        if not _text(path, 1000):
            return None
        out = {"action": action, "path": path.strip()}
        if action == "read_file" and _text(obj.get("sheet"), 100):
            out["sheet"] = obj["sheet"].strip()
        if action == "write_file":
            mode = str(obj.get("mode") or "create").strip().lower()
            rows = obj.get("rows")
            if mode not in WRITE_MODES:
                return None
            if isinstance(rows, list) and all(isinstance(r, list) for r in rows):
                out["rows"] = rows
            elif isinstance(obj.get("content"), str):
                out["content"] = obj["content"]
            else:
                return None
            out["mode"] = mode
        return out
    if action == "say":
        return {"action": action, "text": obj["text"].strip()} if _text(obj.get("text"), 4000) else None
    if action == "ask":
        options = obj.get("options") or []
        if not _text(obj.get("question"), 1000) or not isinstance(options, list):
            return None
        options = [str(o).strip() for o in options if str(o).strip()][:8]
        allow_text = obj.get("allow_text", True) is not False
        if not options and not allow_text:
            return None
        return {"action": action, "question": obj["question"].strip(), "options": options,
                "multi": obj.get("multi") is True, "allow_text": allow_text}
    if action == "account":
        return _validate_account(obj)
    if action == "ado":
        return _validate_ado(obj)
    if action == "type":
        text = obj.get("text")
        if not isinstance(text, str) or text == "":
            return None
        return {"action": action, "text": text}
    if action == "key":
        key = obj.get("key")
        if not isinstance(key, str) or not key.strip():
            return None
        return {"action": action, "key": key.strip().lower()}
    if action == "scroll":
        x, y = _num(obj.get("x")), _num(obj.get("y"))
        direction = obj.get("direction")
        amount = _num(obj.get("amount"))
        if x is None or y is None or amount is None or not isinstance(direction, str):
            return None
        direction = direction.strip().lower()
        if direction not in _DIRECTIONS:
            return None
        return {"action": action, "x": x, "y": y, "direction": direction, "amount": max(1, min(20, amount))}
    if action in ("open_app", "focus"):
        field = "name" if action == "open_app" else "title"
        value = obj.get(field)
        if not isinstance(value, str) or not value.strip() or len(value) > 200:
            return None
        return {"action": action, field: value.strip()}
    if action == "inspect":
        topic = obj.get("topic")
        if not isinstance(topic, str) or topic.strip().lower() not in TOPICS:
            return None
        return {"action": action, "topic": topic.strip().lower()}
    if action == "shell":
        import sysinfo
        shell = str(obj.get("shell") or "powershell").strip().lower()
        command = sysinfo.command_ok(shell, obj.get("command"))
        if not command:
            return None
        return {"action": "shell", "shell": shell, "command": command, "admin": obj.get("admin") is True}
    if action == "wait":
        ms = _num(obj.get("ms"))
        if ms is None or ms < 0:
            return None
        return {"action": action, "ms": min(ms, _WAIT_CAP_MS)}
    summary = obj.get("summary", "")
    if not isinstance(summary, str):
        return None
    return {"action": "done", "summary": summary}


def _validate_account(obj):
    op = str(obj.get("op") or "").strip().lower()
    if op not in ACCOUNT_OPS:
        return None
    out = {"action": "account", "op": op}
    if op == "agenda":
        when = str(obj.get("when") or "today").strip().lower()
        if when not in ACCOUNT_WHEN:
            return None
        out["when"] = when
    if op == "me":
        focus = str(obj.get("focus") or "").strip().lower()
        if focus and focus not in ACCOUNT_FOCUS:
            return None
        if focus:
            out["focus"] = focus
    if op in ("mail", "person") and _text(obj.get("query"), 80):
        out["query"] = " ".join(obj["query"].split())
    if op == "person" and "query" not in out:
        return None
    if op == "event":
        return _validate_event(obj, out)
    if op == "recordings":
        if obj.get("days") is not None:
            days = _num(obj.get("days"))
            if days is None or not 1 <= days <= 30:
                return None
            out["days"] = days
        if _text(obj.get("query"), 80):
            out["query"] = " ".join(obj["query"].split())
    return out


def _validate_ado(obj):
    op = str(obj.get("op") or "").strip().lower()
    if op not in ADO_OPS:
        return None
    out = {"action": "ado", "op": op}
    if _text(obj.get("project"), 500):
        if not obj["project"].strip().startswith("https://"):
            return None
        out["project"] = obj["project"].strip()
    need = {"get": ("id",), "query": ("wiql",), "create": ("type", "title"), "update": ("id", "fields"),
            "comment": ("id", "text"), "import_excel": ("file", "type", "parent")}[op]
    for key in ("id", "parent"):
        if obj.get(key) is not None:
            value = _num(obj[key])
            if value is None or value <= 0:
                return None
            out[key] = value
    for key in ("wiql", "type", "title", "text", "file", "sheet"):
        if _text(obj.get(key), 20000):
            out[key] = obj[key].strip()
    for key in ("fields", "columns"):
        if isinstance(obj.get(key), dict):
            out[key] = obj[key]
    if any(key not in out for key in need):
        return None
    return out


def _validate_event(obj, out):
    subject, start = obj.get("subject"), obj.get("start")
    if not _text(subject, 120) or not _text(start, 16):
        return None
    try:
        datetime.strptime(start.strip(), "%Y-%m-%d %H:%M")
        if obj.get("until"):
            datetime.strptime(str(obj["until"]).strip(), "%Y-%m-%d")
    except ValueError:
        return None
    minutes = _num(obj.get("minutes", 30))
    reminder = _num(obj.get("reminder", 15))
    repeat = str(obj.get("repeat") or "none").strip().lower()
    days = [str(d).strip().lower() for d in obj.get("days") or []] if isinstance(obj.get("days") or [], list) else None
    if minutes is None or not 1 <= minutes <= 1440 or reminder is None or not 0 <= reminder <= 10080:
        return None
    if repeat not in EVENT_REPEATS or days is None or any(d not in EVENT_DAYS for d in days):
        return None
    out.update(subject=" ".join(subject.split()), start=start.strip(), minutes=minutes, reminder=reminder, repeat=repeat)
    if days:
        out["days"] = days
    if obj.get("until"):
        out["until"] = str(obj["until"]).strip()
    if obj.get("free") is True:
        out["free"] = True
    return out


def _text(value, limit):
    return isinstance(value, str) and bool(value.strip()) and len(value) <= limit


def _num(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return int(round(value))


def _check():
    assert "Source:" in INSTRUCTION and "open the official page" in NEXT_TASK
    click = parse_actions("Next step.", [("json", '{"action":"click","x":1,"y":2}')])
    assert click == [{"action": "click", "x": 1, "y": 2}], click
    assert parse_actions("Just prose, no json.", []) is None
    assert parse_actions("", [("json", '{"action":"explode"}')]) is None
    batch = parse_actions("", [("json", '{"actions":[{"action":"open_app","name":"Notepad"},'
                                         '{"action":"type","text":"hi"},{"action":"key","key":"ctrl+s"}]}')])
    assert [a["action"] for a in batch] == ["open_app", "type", "key"], batch
    # A bad entry ends the batch; done ends it too.
    cut = parse_actions("", [("json", '[{"action":"wait","ms":99999},{"action":"done","summary":"x"},{"action":"click","x":1,"y":1}]')])
    assert cut == [{"action": "wait", "ms": 10000}, {"action": "done", "summary": "x"}], cut
    assert parse_actions("", [("json", '[{"action":"nope"},{"action":"wait","ms":1}]')]) is None
    # Half-streamed batch: lenient parsing grabs the first object, complete parsing waits.
    partial = '{"actions":[{"action":"click","x":1,"y":2},{"action":"ty'
    assert parse_actions("", [("json", partial)]) == [{"action": "click", "x": 1, "y": 2}]
    assert parse_actions("", [("json", partial)], complete=True) is None
    fenced = parse_actions('Sure.\n```json\n{"actions":[{"action":"done","summary":"Saved"}]}\n```\n', [])
    assert fenced == [{"action": "done", "summary": "Saved"}], fenced
    assert parse_actions("", [("json", '{"action":"inspect","topic":"Disks"}')]) == [{"action": "inspect", "topic": "disks"}]
    assert parse_actions("", [("json", '{"action":"inspect","topic":"registry"}')]) is None
    assert parse_actions("", [("json", '{"action":"shell","command":"Get-Service Spooler"}')]) == [
        {"action": "shell", "shell": "powershell", "command": "Get-Service Spooler", "admin": False}]
    assert parse_actions("", [("json", '{"action":"shell","command":"Remove-Item C:\\\\Windows","admin":true}')]) is None
    new = parse_actions("", [("json", '{"actions":[{"action":"drag","x":1,"y":2,"to_x":3,"to_y":4},'
                                       '{"action":"window","title":"Edge","op":"Maximize"},{"action":"read_clipboard"},'
                                       '{"action":"open_url","url":"https://example.com"},{"action":"move","x":5,"y":6}]}')])
    assert new == [{"action": "drag", "x": 1, "y": 2, "to_x": 3, "to_y": 4},
                   {"action": "window", "title": "Edge", "op": "maximize"}, {"action": "read_clipboard"},
                   {"action": "open_url", "url": "https://example.com"}, {"action": "move", "x": 5, "y": 6}], new
    assert parse_actions("", [("json", '{"action":"window","title":"x","op":"close"}')]) is None
    smart = parse_actions("", [("json", '{"actions":[{"action":"say","text":"Plan: read then ask"},'
                                         '{"action":"read_file","path":"~/a.xlsx"},'
                                         '{"action":"ask","question":"Which parent?","options":["#123","#124"]},'
                                         '{"action":"click","x":1,"y":1}]}')])
    assert [a["action"] for a in smart] == ["say", "read_file", "ask"], smart
    assert smart[2] == {"action": "ask", "question": "Which parent?", "options": ["#123", "#124"], "multi": False,
                        "allow_text": True}, smart[2]
    imp = parse_actions("", [("json", '{"action":"ado","op":"import_excel","file":"x.xlsx","type":"Product Backlog Item",'
                                       '"parent":"123"}')])
    assert imp is None, "parent must be a number"
    imp = parse_actions("", [("json", '{"action":"ado","op":"import_excel","file":"x.xlsx","type":"Product Backlog Item",'
                                       '"parent":123,"columns":{"Name":"title"}}')])
    assert imp == [{"action": "ado", "op": "import_excel", "parent": 123, "type": "Product Backlog Item",
                    "file": "x.xlsx", "columns": {"Name": "title"}}], imp
    assert parse_actions("", [("json", '{"action":"ado","op":"delete","id":5}')]) is None
    fetched = parse_actions("", [("json", '{"actions":[{"action":"account","op":"next"},{"action":"done","summary":"Checking"}]}')])
    assert fetched == [{"action": "account", "op": "next"}], fetched
    assert parse_actions("", [("json", '{"action":"account","op":"agenda","when":"tomorrow"}')]) == [
        {"action": "account", "op": "agenda", "when": "tomorrow"}]
    assert parse_actions("", [("json", '{"action":"account","op":"person"}')]) is None
    assert parse_actions("", [("json", '{"action":"account","op":"me","focus":"manager"}')]) == [
        {"action": "account", "op": "me", "focus": "manager"}]
    assert parse_actions("", [("json", '{"action":"account","op":"recordings","days":7}')]) == [
        {"action": "account", "op": "recordings", "days": 7}]
    assert parse_actions("", [("json", '{"action":"account","op":"recordings","days":90}')]) is None
    weekly = ('{"action":"account","op":"event","subject":"Fill timesheet","start":"2026-11-20 12:00",'
              '"repeat":"weekly","days":["Friday"],"free":true}')
    assert parse_actions("", [("json", weekly)]) == [{"action": "account", "op": "event", "subject": "Fill timesheet",
        "start": "2026-11-20 12:00", "minutes": 30, "reminder": 15, "repeat": "weekly", "days": ["friday"], "free": True}]
    assert parse_actions("", [("json", '{"action":"account","op":"event","subject":"X","start":"Friday noon"}')]) is None
    assert parse_actions("", [("json", '{"action":"account","op":"event","subject":"X","start":"2026-11-20 12:00","days":["fri"]}')]) is None
    assert parse_actions("", [("json", '{"action":"write_file","path":"a.txt","content":"x","mode":"nuke"}')]) is None
    print("protocol ok")


if __name__ == "__main__":
    _check()
