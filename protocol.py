"""Action contract between work Copilot and this PC.

Copilot replies with one fenced JSON object holding up to MAX_BATCH actions.
parse_actions returns the valid leading actions, normalized, or None.
"""

import json

MAX_BATCH = 5
TOPICS = ("system", "disks", "processes", "network", "services", "events", "windows", "apps",
          "startup", "battery", "updates", "printers", "displays", "audio", "cameras", "outlook", "teams",
          "now", "history")

PLAYBOOKS = """Work context:
- You can read this user's Microsoft 365 work data (email, Teams chats and meetings, calendar, OneDrive and
  SharePoint files, people). For questions about their work, status, what was asked or decided, what they did,
  or who said what, search that data yourself and answer with done; no desktop actions are needed.
- inspect now shows what the user is doing on this PC (window in front, meeting now and next, recent files).
  inspect history lists earlier Co-opilot tasks and their results. Combine both with the work data when useful.
Playbooks:
- Speakers and mics: inspect audio (devices) and cameras (camera/mic privacy and which apps use them now).
  Keys volumemute, volumeup, volumedown, playpause. Settings: open_url ms-settings:sound, ms-settings:privacy-webcam,
  ms-settings:privacy-microphone, ms-settings:bluetooth, ms-settings:display, ms-settings:printers, ms-settings:windowsupdate.
- Teams: inspect teams for health. New Teams keys: ctrl+1 Activity, ctrl+2 Chat, ctrl+3 Teams, ctrl+4 Calendar,
  ctrl+5 Calls, ctrl+e search, ctrl+n new chat. In a call: ctrl+shift+m mute, ctrl+shift+o camera, ctrl+shift+k raise hand,
  ctrl+shift+h leave. Chat with someone: open_url msteams:/l/chat/0/0?users=name@example.com then focus Teams.
- Outlook: inspect outlook for unread mail and today's meetings (classic Outlook only; otherwise read the screen).
  New email: open_url mailto:a@b.ca?subject=Hello%20there&body=Line%20one (percent-encode spaces and new lines as %0A).
  Keys: ctrl+1 Mail, ctrl+2 Calendar, ctrl+shift+m new email, ctrl+r reply, f9 send/receive, ctrl+enter send.
- Before sending a message or email, joining, starting or leaving a call, accepting or declining a meeting, or deleting
  anything: only do it if the task explicitly asks for that. Otherwise prepare it and say it is ready for the user to send.
"""
WINDOW_OPS = ("minimize", "maximize", "restore", "move_to_main")
ADO_OPS = ("get", "query", "create", "update", "comment", "import_excel")
WRITE_MODES = ("create", "append", "overwrite")

FLOW = """How to work (be smart, not busy):
1. For anything beyond one step, start the first batch with say: a 1-3 line plan.
2. Gather facts before acting: inspect, list_dir, read_file, ado get/query, your Microsoft 365 data, the screenshot.
3. If the request is ambiguous or a choice changes the outcome (which file, project, parent, mapping, wording),
   use ask with 2-6 clear options instead of guessing. ask must be the last action in its batch.
4. Prefer direct skills over clicking: read_file/write_file for files, ado for Azure DevOps, open_url/open_path to open.
5. After acting, verify (read it back, ado get, or look at the screenshot), then done with what changed and where.
6. If something fails, read the error in the report, adjust, and retry a different way once before giving up.
"""

SKILLS = """Files: list_dir {"path"}; read_file {"path","sheet"?} reads text, csv, json, Excel (.xlsx) and Word (.docx);
write_file {"path","content"} for text or {"path","rows":[[...],...]} for .xlsx/.csv, "mode": create (default), append,
or overwrite (the user is asked to confirm). Paths may use %USERPROFILE%, %ONEDRIVE%, or ~.
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
Talking to the user: say {"text"} posts a short progress note and continues;
ask {"question","options":["..."],"multi":false,"allow_text":true} shows a card with radio buttons (or checkboxes when multi)
and a details box, then waits; the answer comes back in the report.
"""

SUMMARY_RULE = r"""done.summary is shown to the user as formatted text. Format it like this, with \n between lines:
one sentence answer first, then a "- " bullet per item on its own line, **bold** for names, `code` for paths or commands,
"## " for a heading when there are groups. Never put a list inside one sentence or separate items with commas.
Example: "You have **2** Edge windows open.\n## Window 1\n- **Intranet**\n- **Timesheets**\n## Window 2\n- **Dashboard**"
"""

INSTRUCTION = f"""You control this Windows PC and help diagnose it. Each message includes a screenshot.
Reply with exactly one fenced json object and no other text:
{{"actions":[ ...1 to {MAX_BATCH} actions... ]}}
Batch actions that do not need a new screenshot in between, for example open_app then type then key.
The screenshot shows every monitor side by side; each message lists which pixel ranges belong to which monitor.
Coordinates are pixels of the attached screenshot, origin at its top-left, and can point at any monitor.
After the batch you get a JSON report (what ran, errors, inspect and read results) and a new screenshot.

Actions:
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
{{"action":"list_dir","path":"%USERPROFILE%\\\\Documents"}}
{{"action":"read_file","path":"C:\\\\path\\\\file.xlsx"}}
{{"action":"write_file","path":"C:\\\\path\\\\out.txt","content":"..."}}
{{"action":"ado","op":"get","id":123}}
{{"action":"say","text":"Plan: ..."}}
{{"action":"ask","question":"Which file?","options":["A","B"],"multi":false,"allow_text":true}}
{{"action":"wait","ms":1000}}
{{"action":"done","summary":"..."}}

key is enter, tab, escape, up, down, left, right, f1-f12, volumemute, volumeup, volumedown, playpause,
or a chord such as ctrl+s or win+e.
direction is up, down, left, or right. amount is 1 to 20.
move hovers (for tooltips and menus). drag holds the left button from x,y to to_x,to_y.
open_app is the name as shown in the Start menu, such as Notepad or File Explorer.
open_url opens an http, https, mailto, ms-settings, or msteams link. open_path opens a folder or document; %USERPROFILE% works.
focus brings an open window to the front. window op is {", ".join(WINDOW_OPS)}.
read_clipboard returns the clipboard text; use ctrl+a or a selection then ctrl+c first to read text exactly.
inspect reads this PC without changing it. Topics: {", ".join(TOPICS)}. windows lists titles with their monitor.
Use inspect for questions about the system, performance, disk space, network, failing services, or recent errors,
then explain findings in done.summary.
{SKILLS}
{FLOW}
{PLAYBOOKS}
{SUMMARY_RULE}
When the task is finished, put done last in the batch.
"""

NEXT_TASK = (f"New task. Same rules as before: one fenced json object with up to {MAX_BATCH} actions, done when finished.\n"
             f"{SUMMARY_RULE}")

REMINDER = 'Reply with only one fenced json object like {"actions":[...]} for the next step. No other text.'

_POINT_ACTIONS = {"click", "double_click", "right_click", "move"}
_ACTIONS = _POINT_ACTIONS | {"type", "key", "scroll", "open_app", "focus", "inspect", "wait", "done",
                             "drag", "open_url", "open_path", "window", "read_clipboard",
                             "list_dir", "read_file", "write_file", "ado", "say", "ask"}
# These end a batch: the next step depends on what they return.
_ENDS_BATCH = {"done", "ask"}
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
    if action == "read_clipboard":
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
    if action == "wait":
        ms = _num(obj.get("ms"))
        if ms is None or ms < 0:
            return None
        return {"action": action, "ms": min(ms, _WAIT_CAP_MS)}
    summary = obj.get("summary", "")
    if not isinstance(summary, str):
        return None
    return {"action": "done", "summary": summary}


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


def _text(value, limit):
    return isinstance(value, str) and bool(value.strip()) and len(value) <= limit


def _num(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return int(round(value))


def _check():
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
    assert parse_actions("", [("json", '{"action":"write_file","path":"a.txt","content":"x","mode":"nuke"}')]) is None
    print("protocol ok")


if __name__ == "__main__":
    _check()
