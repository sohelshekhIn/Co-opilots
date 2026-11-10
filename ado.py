"""Azure DevOps work items through the REST API, called from a tab of the app's signed-in browser.

The browser's work sign-in authenticates the calls, so no personal access token is needed.
"""

import html
import json
import os
import re
import time
from pathlib import Path

import files

SETTINGS = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local") / "CoPilotComputerUse" / "settings.json"
API = "api-version=7.0"
FIELDS = {
    "title": "System.Title",
    "description": "System.Description",
    "acceptance_criteria": "Microsoft.VSTS.Common.AcceptanceCriteria",
    "tags": "System.Tags",
    "state": "System.State",
    "assigned_to": "System.AssignedTo",
    "area": "System.AreaPath",
    "iteration": "System.IterationPath",
    "priority": "Microsoft.VSTS.Common.Priority",
    "effort": "Microsoft.VSTS.Scheduling.Effort",
    "story_points": "Microsoft.VSTS.Scheduling.StoryPoints",
    "value_area": "Microsoft.VSTS.Common.ValueArea",
    "repro_steps": "Microsoft.VSTS.TCM.ReproSteps",
}
_HTML_TAG = re.compile(r"</?(p|br|div|ul|ol|li|b|strong|i|em|u|table|tr|td|th|h[1-6]|span|a)\b[^>]*>", re.I)
_HTML_FIELDS = {"System.Description", "Microsoft.VSTS.Common.AcceptanceCriteria", "Microsoft.VSTS.TCM.ReproSteps"}
_SHOW = ["System.WorkItemType", "System.Title", "System.State", "System.AssignedTo", "System.AreaPath",
         "System.IterationPath", "System.Tags", "Microsoft.VSTS.Common.Priority", "Microsoft.VSTS.Scheduling.Effort",
         "System.Description", "Microsoft.VSTS.Common.AcceptanceCriteria"]

_FETCH = r"""
async ([method, url, body, type]) => {
  const response = await fetch(url, {method, credentials: 'include',
    headers: {'Content-Type': type, 'Accept': 'application/json', 'X-TFS-FedAuthRedirect': 'Suppress'},
    body: body === null ? undefined : body});
  return {status: response.status, text: await response.text()};
}
"""


def project_url(given=None):
    """The project URL from the action, else the last one that worked."""
    url = (given or _settings().get("ado_project") or "").strip().rstrip("/")
    if not re.match(r"^https://[^\s/]+/\S+$", url):
        raise ValueError("No Azure DevOps project URL known yet. Ask the user for it, "
                         "e.g. https://dev.azure.com/<org>/<project>, and pass it as project.")
    return url


def field_name(name):
    key = str(name).strip()
    return FIELDS.get(key.lower().replace(" ", "_"), key) if "." not in key else key


class Ado:
    def __init__(self, chat):
        self.chat = chat
        self._origin = None

    def run(self, action):
        project = project_url(action.get("project"))
        op = action["op"]
        try:
            if op == "get":
                result = self.describe(project, action["id"])
            elif op == "query":
                result = self.query(project, action["wiql"])
            elif op == "create":
                fields = dict(action.get("fields") or {}, title=action["title"])
                result = self.create(project, action["type"], fields, action.get("parent"))
            elif op == "update":
                result = self.update(project, action["id"], action["fields"])
            elif op == "comment":
                result = self.comment(project, action["id"], action["text"])
            elif op == "import_excel":
                result = self.import_rows(project, action["rows"], action["type"], action["parent"])
            else:
                raise ValueError(f"unknown ado op: {op}")
        finally:
            self.chat.front()
        _save_settings(ado_project=project)
        return result

    # ---- operations ----

    def describe(self, project, item_id):
        item = self._call("GET", f"{_org(project)}/_apis/wit/workitems/{int(item_id)}?$expand=relations&{API}", project)
        return _summary(item, full=True)

    def query(self, project, wiql):
        found = self._call("POST", f"{project}/_apis/wit/wiql?$top=50&{API}", project, json.dumps({"query": wiql}))
        ids = [w["id"] for w in found.get("workItems", [])][:50]
        if not ids:
            return "No work items match."
        batch = self._call("POST", f"{_org(project)}/_apis/wit/workitemsbatch?{API}", project,
                           json.dumps({"ids": ids, "fields": _SHOW[:9]}))
        return "\n".join(_summary(item) for item in batch.get("value", []))

    def create(self, project, item_type, fields, parent=None, inherit=True):
        patch = _patch(fields, "add")
        if parent:
            patch.append({"op": "add", "path": "/relations/-", "value": {
                "rel": "System.LinkTypes.Hierarchy-Reverse", "url": f"{_org(project)}/_apis/wit/workItems/{int(parent)}"}})
            if inherit:
                given = {p["path"] for p in patch}
                parent_fields = self._call("GET", f"{_org(project)}/_apis/wit/workitems/{int(parent)}?{API}", project)["fields"]
                for name in ("System.AreaPath", "System.IterationPath"):
                    if f"/fields/{name}" not in given and parent_fields.get(name):
                        patch.append({"op": "add", "path": f"/fields/{name}", "value": parent_fields[name]})
        kind = item_type.replace(" ", "%20")
        item = self._call("POST", f"{project}/_apis/wit/workitems/${kind}?{API}", project, json.dumps(patch),
                          "application/json-patch+json")
        return f"Created #{item['id']}: {item['fields'].get('System.Title')}  {_link(project, item['id'])}"

    def update(self, project, item_id, fields):
        item = self._call("PATCH", f"{_org(project)}/_apis/wit/workitems/{int(item_id)}?{API}", project,
                          json.dumps(_patch(fields, "add")), "application/json-patch+json")
        return f"Updated #{item['id']}: {item['fields'].get('System.Title')}  {_link(project, item['id'])}"

    def comment(self, project, item_id, text):
        self._call("POST", f"{project}/_apis/wit/workItems/{int(item_id)}/comments?api-version=7.0-preview.3", project,
                   json.dumps({"text": _html(text)}))
        return f"Commented on #{int(item_id)}"

    def import_rows(self, project, rows, item_type, parent):
        done, failed = [], []
        for number, fields in enumerate(rows, 1):
            try:
                done.append(self.create(project, item_type, fields, parent))
            except Exception as exc:
                failed.append(f"row {number} ({fields.get('title', '?')}): {exc}")
        lines = [f"Created {len(done)} of {len(rows)} child items under #{parent}:"] + done
        if failed:
            lines += ["Failed:"] + failed
        return "\n".join(lines)

    # ---- transport ----

    def _call(self, method, url, project, body=None, content_type="application/json"):
        page = self._ready(project)
        response = page.evaluate(_FETCH, [method, url, body, content_type])
        status, text = response["status"], response["text"]
        if status in (203, 302, 401) or text.lstrip().startswith("<"):
            raise RuntimeError("Azure DevOps did not accept the browser sign-in. Open the project in the "
                               "Co-opilot browser once (Show sign-in) or check the URL.")
        data = json.loads(text) if text else {}
        if status >= 400:
            raise RuntimeError(f"Azure DevOps {status}: {data.get('message', text[:300])}")
        return data

    def _ready(self, project):
        """The side tab, sitting on the project's site so fetch carries its cookies."""
        page = self.chat.side_page()
        origin = re.match(r"^https://[^/]+", project).group(0)
        if self._origin != origin or not page.url.startswith(origin):
            page.goto(project, wait_until="domcontentloaded", timeout=60_000)
            deadline = time.time() + 45
            while not page.url.startswith(origin) and time.time() < deadline:
                time.sleep(0.5)
            if not page.url.startswith(origin):
                raise RuntimeError(f"Azure DevOps sign-in did not finish (stuck at {page.url[:80]})")
            self._origin = origin
        return page


def import_preview(action):
    """Map an Excel sheet to work item fields. Returns (rows, preview text) without creating anything."""
    raw = files.excel_rows(action["file"], action.get("sheet"))
    columns = action.get("columns") or {}
    rows = []
    for record in raw:
        fields = {}
        for header, value in record.items():
            target = columns.get(header) or header
            if value and str(target).strip().lower() not in ("skip", "ignore", "-"):
                fields[str(target).strip().lower().replace(" ", "_")] = value
        if fields.get("title"):
            rows.append(fields)
    if not rows:
        raise ValueError("No rows with a title. Map the title column with columns, e.g. {\"Name\": \"title\"}.")
    used = sorted({k for row in rows for k in row})
    lines = [f"{len(rows)} {action['type']} items under #{action['parent']} from {Path(action['file']).name}",
             f"Fields: {', '.join(used)}"]
    lines += [f"- {row['title']}" for row in rows[:12]]
    if len(rows) > 12:
        lines.append(f"- ...and {len(rows) - 12} more")
    return rows, "\n".join(lines)


def _patch(fields, op):
    patch = []
    for name, value in fields.items():
        reference = field_name(name)
        if reference == "System.Tags" and isinstance(value, (list, tuple)):
            value = "; ".join(map(str, value))
        if reference in _HTML_FIELDS and isinstance(value, str) and not _HTML_TAG.search(value):
            value = _html(value)
        patch.append({"op": op, "path": f"/fields/{reference}", "value": value})
    return patch


def _html(text):
    return "<br>".join(html.escape(line) for line in str(text).split("\n"))


def _summary(item, full=False):
    f = item.get("fields", {})
    who = f.get("System.AssignedTo")
    who = who.get("displayName") if isinstance(who, dict) else who
    line = (f"#{item.get('id')} [{f.get('System.WorkItemType')}] {f.get('System.Title')} | {f.get('System.State')}"
            f" | {who or 'unassigned'} | {f.get('System.IterationPath', '')} | tags: {f.get('System.Tags', '')}")
    if not full:
        return line
    parts = [line, f"Area: {f.get('System.AreaPath')}  Priority: {f.get('Microsoft.VSTS.Common.Priority')}  "
                   f"Effort: {f.get('Microsoft.VSTS.Scheduling.Effort')}"]
    for name, label in (("System.Description", "Description"), ("Microsoft.VSTS.Common.AcceptanceCriteria", "Acceptance criteria")):
        if f.get(name):
            parts.append(f"{label}: {_text(f[name])[:1500]}")
    relations = item.get("relations") or []
    parents = [r["url"].rsplit("/", 1)[-1] for r in relations if r.get("rel") == "System.LinkTypes.Hierarchy-Reverse"]
    children = [r["url"].rsplit("/", 1)[-1] for r in relations if r.get("rel") == "System.LinkTypes.Hierarchy-Forward"]
    parts.append(f"Parent: {', '.join(parents) or 'none'}  Children: {', '.join(children) or 'none'}")
    return "\n".join(parts)


def _text(value):
    value = re.sub(r"<br\s*/?>|</p>|</li>|</div>", "\n", str(value), flags=re.I)
    return html.unescape(re.sub(r"<[^>]+>", "", value)).strip()


def _org(project):
    return project.rsplit("/", 1)[0]


def _link(project, item_id):
    return f"{project}/_workitems/edit/{item_id}"


def _settings():
    try:
        return json.loads(SETTINGS.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _save_settings(**values):
    data = _settings()
    if all(data.get(k) == v for k, v in values.items()):
        return
    data.update(values)
    try:
        SETTINGS.parent.mkdir(parents=True, exist_ok=True)
        SETTINGS.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except OSError:
        pass


def _check():
    assert field_name("Acceptance Criteria") == "Microsoft.VSTS.Common.AcceptanceCriteria"
    assert field_name("Custom.Team") == "Custom.Team" and field_name("tags") == "System.Tags"
    patch = _patch({"title": "A & B", "acceptance_criteria": "Given x\nThen <y>", "tags": ["ui", "auth"]}, "add")
    assert patch[0] == {"op": "add", "path": "/fields/System.Title", "value": "A & B"}, patch
    assert patch[1]["value"] == "Given x<br>Then &lt;y&gt;", patch[1]
    assert patch[2]["value"] == "ui; auth", patch[2]
    assert _org("https://dev.azure.com/org/My%20Project") == "https://dev.azure.com/org"
    import tempfile
    sheet = Path(tempfile.mkdtemp()) / "pbis.xlsx"
    files.write_file(sheet, rows=[["Name", "AC", "Tags", "Notes"], ["Login", "Given a user", "auth", "x"], ["", "", "", ""]])
    rows, preview = import_preview({"file": str(sheet), "type": "Product Backlog Item", "parent": 123,
                                    "columns": {"Name": "title", "AC": "acceptance_criteria", "Notes": "skip"}})
    assert rows == [{"title": "Login", "acceptance_criteria": "Given a user", "tags": "auth"}], rows
    assert preview.startswith("1 Product Backlog Item items under #123"), preview
    print("ado ok")


if __name__ == "__main__":
    _check()
