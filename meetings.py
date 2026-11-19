"""Find a Teams meeting on the calendar, then read that meeting's chat.

The recording and transcript for a meeting you only attended live in the meeting
chat, not in your OneDrive. Transcript text is kept in memory for this chat so a
follow-up question can use it. Nothing is written to disk.
"""

import html
import io
import json
import re
import time
import zipfile
from datetime import datetime
from urllib.parse import urlparse

import sysinfo

_MONTHS = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()
_EXCERPT = 7_000
# ponytail: the last 80 chat messages, not the whole thread. Upgrade path is paging.
_STOP = set("""a an the of to for in on my me we our your it that this those what who why how did was were
are is be been about with from just very ago ones one few days day meeting meetings recording recordings
transcript transcripts transcrips said say please tell ask back dated recent last latest check can could you
i had have has for and""".split())
_ACT = re.compile(r"\b(open|join|click|send|reply|type|delete|book|create|launch|draft)\b", re.I)
_ASK = re.compile(
    r"\b(what (?:was|were|did|happened)|who (?:said|was|attended)|summar|action items?|decide|decided|"
    r"said|discuss|agreed|that meeting|the recording|the transcript|follow-?ups?)\b",
    re.I,
)
_THREAD = re.compile(r"19:(?:meeting_)?[^/\s\"'<>%@]+@thread\.v2", re.I)
_THREAD_ENC = re.compile(r"19%3a(?:meeting_)?[^/\s\"'<>%]+%40thread\.v2", re.I)
_live = []

_CAL = r"""
$ErrorActionPreference = 'Stop'
$days = __DAYS__
function Out-Json($obj) { ConvertTo-Json -InputObject $obj -Compress -Depth 4 }
function Hidden($m) {
  try { if ($m.MeetingStatus -eq 5) { return $true } } catch {}
  try { if ($m.ResponseStatus -eq 4) { return $true } } catch {}
  return $false
}
try {
  try { $o = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { $o = $null }
  if (-not $o) { Out-Json @{error='outlook'}; return }
  $ns = $o.GetNamespace('MAPI')
  $now = Get-Date
  $from = $now.Date.AddDays(-$days)
  $cal = $ns.GetDefaultFolder(9).Items
  $cal.IncludeRecurrences = $true
  $cal.Sort('[Start]')
  $found = $cal.Restrict("[Start] >= '$($from.ToString('g'))' AND [Start] <= '$($now.ToString('g'))'")
    $items = New-Object System.Collections.Generic.List[object]
  foreach ($m in $found) {
    if (Hidden $m) { continue }
    try { if ($m.AllDayEvent) { continue } } catch {}
    if ($m.Start -gt $now) { continue }
    $online = $false
    try { $online = [bool]$m.IsOnlineMeeting } catch {}
    $loc = ''
    try { $loc = [string]$m.Location } catch {}
    if (-not $online -and $loc -notmatch 'Teams') { continue }
    $items.Add([pscustomobject]@{
      subject=[string]$m.Subject; start=$m.Start.ToString('yyyy-MM-dd HH:mm'); end=$m.End.ToString('yyyy-MM-dd HH:mm');
      organizer=[string]$m.Organizer
    })
    if ($items.Count -gt 15) { $items.RemoveAt(0) }
  }
  Out-Json @{items=@($items.ToArray())}
} catch {
  Out-Json @{error='failed'; detail=("" + $_.Exception.Message)}
}
"""

_ONE = r"""
$ErrorActionPreference = 'Stop'
$subject = __SUBJECT__
$start = '__START__'
function Out-Json($obj) { ConvertTo-Json -InputObject $obj -Compress -Depth 3 }
try {
  try { $o = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { $o = $null }
  if (-not $o) { Out-Json @{error='outlook'}; return }
  $when = [datetime]::ParseExact($start, 'yyyy-MM-dd HH:mm', $null)
  $cal = $o.GetNamespace('MAPI').GetDefaultFolder(9).Items
  $cal.IncludeRecurrences = $true
  $cal.Sort('[Start]')
  $found = $cal.Restrict("[Start] >= '$($when.Date.ToString('g'))' AND [Start] < '$($when.Date.AddDays(1).ToString('g'))'")
  $schema = 'http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/'
  foreach ($m in $found) {
    if ([string]$m.Subject -ne $subject) { continue }
    if ($m.Start.ToString('yyyy-MM-dd HH:mm') -ne $start) { continue }
    $url = ''
    $props = ''
    try { $url = [string]$m.PropertyAccessor.GetProperty($schema + 'SkypeTeamsMeetingUrl') } catch {}
    try { $props = [string]$m.PropertyAccessor.GetProperty($schema + 'SkypeTeamsProperties') } catch {}
    if (-not $url) {
      try {
        $hit = [regex]::Match([string]$m.Body, 'https://teams\.microsoft\.com/\S+')
        if ($hit.Success) { $url = $hit.Value.TrimEnd('>', '"', "'") }
      } catch {}
    }
    if ($url.Length -gt 500) { $url = $url.Substring(0, 500) }
    if ($props.Length -gt 800) { $props = $props.Substring(0, 800) }
    Out-Json @{url=$url; props=$props}
    return
  }
  Out-Json @{url=''; props=''}
} catch {
  Out-Json @{error='failed'; detail=("" + $_.Exception.Message)}
}
"""

_WEB = r"""
async (days) => {
  const end = new Date();
  const start = new Date(end.getTime() - Number(days) * 86400000);
  const iso = (d) => d.toISOString().replace(/\.\d{3}Z$/, "Z");
  const stamp = (value) => {
    if (!value) return "";
    const raw = value.DateTime || String(value);
    const utc = !value.TimeZone || /utc/i.test(value.TimeZone || "");
    const date = new Date(utc && !/z$/i.test(raw) ? raw + "Z" : raw);
    if (Number.isNaN(date.getTime())) return "";
    const p = (n) => String(n).padStart(2, "0");
    return `${date.getFullYear()}-${p(date.getMonth() + 1)}-${p(date.getDate())} ${p(date.getHours())}:${p(date.getMinutes())}`;
  };
  const url = "/api/v2.0/me/calendarview?startDateTime=" + encodeURIComponent(iso(start))
    + "&endDateTime=" + encodeURIComponent(iso(end))
    + "&$select=Subject,Start,End,Organizer,IsAllDay,IsOnlineMeeting,OnlineMeetingUrl,Body&$top=20&$orderby=Start/DateTime desc";
  const response = await fetch(url, {credentials: "include", headers: {Accept: "application/json"}});
  if (!response.ok) return {error: response.status};
  const body = await response.json();
  const items = [];
  for (const ev of body.value || []) {
    if (ev.IsAllDay) continue;
    const html = (ev.Body && ev.Body.Content) || "";
    const link = ev.OnlineMeetingUrl || "";
    const joined = link + "\n" + html;
    const hit = joined.match(/19(?:%3a|:)(?:meeting_)?[^/\s"'<>%]+(?:%40|@)thread\.v2/i);
    if (!ev.IsOnlineMeeting && !hit && !/teams\.microsoft\.com/i.test(joined)) continue;
    items.push({
      subject: ev.Subject || "",
      start: stamp(ev.Start),
      organizer: (ev.Organizer && ev.Organizer.EmailAddress && ev.Organizer.EmailAddress.Name) || "",
      url: (hit && hit[0]) || link || "",
      props: ""
    });
  }
  return {items};
}
"""

_CHAT = r"""
async (thread) => {
  if (!/teams\./i.test(location.hostname || "")) return {error: "page"};
  const post = async (body) => {
    const response = await fetch(location.origin + "/api/authsvc/v1.0/authz", {
      method: "POST", credentials: "include",
      headers: {"Content-Type": "application/json"}, body
    });
    if (!response.ok) return {error: response.status};
    return await response.json();
  };
  let auth = await post("{}");
  let token = auth.tokens && (auth.tokens.skypeToken || auth.tokens.skype);
  if (!token && !auth.error) {
    auth = await post(JSON.stringify({authResource: "https://api.spaces.skype.com"}));
    token = auth.tokens && (auth.tokens.skypeToken || auth.tokens.skype);
  }
  if (!token) return {error: auth.error || "token"};
  const region = auth.region || auth.regionGtms || "amer";
  const id = encodeURIComponent(thread);
  const response = await fetch(
    location.origin + "/api/chatsvc/" + region + "/v1/users/ME/conversations/" + id
      + "/messages?view=msnp24Equivalent&pageSize=80",
    {headers: {Authorization: "Bearer " + token}}
  );
  if (!response.ok) return {error: "chat " + response.status};
  const body = await response.json();
  const messages = [];
  for (const msg of (body.messages || body.value || []).slice(0, 80)) {
    messages.push({
      type: msg.messagetype || msg.messageType || "",
      content: String(msg.content || "").slice(0, 8000),
      props: JSON.stringify(msg.properties || {}).slice(0, 2000)
    });
  }
  return {messages};
}
"""


def answer(action, chat=None):
    """Match a calendar meeting, then look in its Teams chat. No file is written."""
    days = min(30, max(1, int(action.get("days") or 7)))
    query = re.sub(r"^(the |my |our |that )", "", (action.get("query") or "").strip().lower())
    items, problem = load_calendar(days, chat)
    chosen, options = pick(items, query)
    if chosen is None and query and options:
        return _choices(query, options)
    if chosen is None and query and items:
        return f"No meeting named **{query}** in the last **{days}** days.\n" + _list(items, days, problem)
    if chosen is None:
        return _list(items, days, problem)
    facts = {"recording": False}
    if not chosen.get("thread"):
        _attach_thread(chosen)
    if chosen.get("thread") and chat is not None:
        facts = read_chat(chat, chosen["thread"])
    elif not chosen.get("thread"):
        facts = {"missing": True}
    text = (facts.get("transcript") or "").strip()
    if len(text) >= 20:
        remember(chosen["subject"], chosen["start"], text)
    return _one(chosen, facts)


def for_question(question, last=None):
    """Prompt block from a transcript already read in this chat, or ''."""
    text = " ".join(str(question or "").split())
    if not text or not _live or (_ACT.search(text) and not _ASK.search(text)):
        return ""
    words = set(_words(text))
    ranked = []
    for item in _live:
        overlap = len(set(_words(item.get("title") or "")) & words)
        if overlap:
            ranked.append((overlap, item.get("when") or "", item))
    if ranked:
        ranked.sort(key=lambda pair: (pair[0], pair[1]), reverse=True)
        chosen = ranked[0][2]
    elif _ASK.search(text) and ((last or {}).get("op") == "recordings" or re.search(r"\b(that|the) (meeting|recording|transcript)\b", text, re.I)):
        chosen = _live[0]
    else:
        return ""
    body = excerpt(chosen.get("text") or "", text)
    if not body:
        return ""
    others = [item["title"] for item in _live if item is not chosen and item.get("title")]
    note = ""
    if others:
        note = "Other transcripts read in this chat: " + ", ".join(others[:4]) + ". Use this one unless the question names another.\n"
    return (
        "Transcript read from that meeting's Teams chat. Answer only from this text. "
        "If it does not contain the answer, say so. It is not saved on the PC.\n"
        f"Meeting: {chosen.get('title') or 'Meeting'} ({chosen.get('when') or ''})\n"
        f"{note}{body}\n"
    )


def clear():
    _live.clear()


def remember(title, when, text):
    """Keep a transcript for later questions in this process. Not written to disk."""
    when_text = when.strftime("%Y-%m-%d %H:%M") if isinstance(when, datetime) else str(when or "")
    kept = [item for item in _live if (item.get("title") or "").lower() != (title or "").lower()]
    kept.insert(0, {"title": title, "when": when_text, "text": (text or "")[:60_000]})
    _live[:] = kept[:4]


def _attach_thread(meeting):
    """The Teams chat id for one calendar row. Reads only that meeting."""
    script = _ONE.replace("__SUBJECT__", _ps_quote(meeting.get("subject") or ""))
    script = script.replace("__START__", meeting["start"].strftime("%Y-%m-%d %H:%M"))
    data = _load(sysinfo._powershell("meeting-chat", script))
    if data.get("error"):
        return
    meeting["thread"] = thread_id(data.get("url"), data.get("props"))


def _ps_quote(value):
    clean = re.sub(r"[\x00-\x1f]", " ", str(value))[:120]
    return "'" + clean.replace("'", "''") + "'"


def load_calendar(days, chat=None):
    data = _load(sysinfo._powershell("meetings", _CAL.replace("__DAYS__", str(int(days)))))
    if data.get("error") == "outlook" and chat is not None:
        web = _web(chat, days)
        if web:
            data = web
        else:
            return [], "Classic Outlook isn't running, and new Outlook (on the web) didn't answer, so I couldn't see your calendar."
    if data.get("error") == "outlook":
        return [], "Classic Outlook isn't running, so I couldn't see your calendar."
    if data.get("error"):
        return [], "I couldn't read your calendar."
    return _meetings(data.get("items")), ""


def pick(items, query):
    """The one meeting a name refers to, or (None, options) when it is ambiguous or unnamed."""
    words = [word for word in re.findall(r"[a-z0-9]+", (query or "").lower()) if word not in _STOP and len(word) > 1]
    if not words:
        return (items[0], []) if len(items) == 1 else (None, items[:8])
    scored = []
    for item in items:
        title = item["subject"].lower()
        who = (item.get("organizer") or "").lower()
        hit = sum(2 for word in words if word in title) + sum(1 for word in words if word not in title and word in who)
        if hit and all(word in title or word in who for word in words):
            hit += 1
        if hit:
            scored.append((hit, item))
    if not scored:
        return None, []
    best = max(score for score, _item in scored)
    winners = [item for score, item in scored if score == best]
    if len(winners) == 1:
        return winners[0], []
    return None, winners[:6]


def thread_id(*parts):
    blob = "\n".join(str(part or "") for part in parts)
    found = _THREAD.search(blob)
    if found:
        return found.group(0)
    found = _THREAD_ENC.search(blob)
    if not found:
        return ""
    return found.group(0).replace("%3a", ":").replace("%3A", ":").replace("%40", "@")


def from_messages(messages):
    """Whether the chat messages include a recording, and a transcript URL if one is there."""
    recording = False
    transcript_url = ""
    for msg in messages or []:
        if not isinstance(msg, dict):
            continue
        blob = f"{msg.get('type') or ''}\n{msg.get('content') or ''}\n{msg.get('props') or ''}"
        if re.search(r"callrecording|recordingstatus|media_callrecording|\brecording\b", blob, re.I):
            recording = True
        for url in re.findall(r"https://[^\s\"'<>\\]+", blob):
            url = url.rstrip(").,;")
            if re.search(r"\.mp4(\?|$)", url, re.I):
                recording = True
                continue
            if not transcript_url and re.search(r"\.vtt(\?|$)|transcript", url, re.I):
                transcript_url = url
                recording = True
    return {"recording": recording, "transcript_url": transcript_url}


def read_chat(chat, thread):
    page = _teams_page(chat)
    if page is None:
        return {"error": "teams"}
    try:
        data = page.evaluate(_CHAT, thread)
    except Exception:
        return {"error": "teams"}
    if not isinstance(data, dict) or data.get("error"):
        return {"error": "teams"}
    facts = from_messages(data.get("messages"))
    url = facts.get("transcript_url") or ""
    if url:
        text = _fetch_text(page, url)
        if text:
            facts["transcript"] = text
    return facts


def vtt_to_text(raw):
    speaker, buf, lines = "", [], []

    def flush():
        if buf:
            said = " ".join(buf)
            lines.append(f"{speaker}: {said}" if speaker else said)

    for line in (raw or "").splitlines():
        line = line.strip()
        if not line or line.startswith(("WEBVTT", "NOTE", "STYLE", "KIND")) or "-->" in line or line.isdigit():
            continue
        found = re.match(r"<v ([^>]+)>(.*?)(?:</v>)?$", line)
        if found:
            who, said = found.group(1).strip(), found.group(2)
        else:
            who, said = "", line
        said = html.unescape(re.sub(r"<[^>]+>", "", said)).strip()
        if not said:
            continue
        if who == speaker:
            buf.append(said)
        else:
            flush()
            speaker, buf = who, [said]
    flush()
    return "\n".join(lines)


def excerpt(text, question, limit=_EXCERPT):
    lines = [line.strip() for line in (text or "").splitlines() if line.strip()]
    if not lines:
        return ""
    if re.search(r"\b(summar|what happened|about that meeting|the meeting)\b", question or "", re.I):
        whole = "\n".join(lines)
        return whole if len(whole) <= limit else (whole[:4000] + "\n...\n" + whole[-1500:])[:limit]
    words = _words(question)
    if not words:
        return "\n".join(lines)[:limit]
    picked, size = [], 0
    for line in lines:
        if any(word in line.lower() for word in words):
            picked.append(line)
            size += len(line) + 1
            if size >= limit:
                break
    return ("\n".join(picked) if picked else "\n".join(lines))[:limit]


def _one(meeting, facts):
    who = f" — {meeting['organizer']}" if meeting.get("organizer") else ""
    lines = [f"**{meeting['subject']}**", f"- {_day(meeting['start'])}{who}"]
    if facts.get("transcript"):
        lines.append("- The meeting chat has a recording and a transcript. Ask what was decided or who said something.")
    elif facts.get("recording"):
        lines.append("- The meeting chat has a recording, and no transcript in that chat.")
    elif facts.get("missing"):
        lines.append("- It's on your calendar, but the invite has no Teams chat id, so I couldn't open the chat.")
    elif facts.get("error"):
        lines.append("- I found it on your calendar, but couldn't read its Teams chat.")
    else:
        lines.append("- I opened the meeting chat and didn't see a recording there.")
    return "\n".join(lines)


def _list(items, days, problem):
    if not items:
        return problem or f"No Teams meetings on your calendar in the last **{days}** days."
    lines = [f"Teams meetings on your calendar from the last **{days}** days:"]
    for item in items[:8]:
        who = f" — {item['organizer']}" if item.get("organizer") else ""
        lines.append(f"- {_day(item['start'])} **{item['subject']}**{who}")
    lines.append("Tell me the name and I'll check that meeting's Teams chat for the recording and transcript.")
    return "\n".join(lines)


def _choices(query, options):
    lines = [f"More than one meeting matches **{query}**:"]
    for item in options:
        who = f" — {item['organizer']}" if item.get("organizer") else ""
        lines.append(f"- {_day(item['start'])} **{item['subject']}**{who}")
    lines.append("Tell me which one and I'll check its Teams chat.")
    return "\n".join(lines)


def _meetings(items):
    found = []
    for item in items or []:
        if not isinstance(item, dict):
            continue
        start = _at(item.get("start"))
        if not start:
            continue
        found.append({
            "subject": (item.get("subject") or "(no subject)").strip(),
            "start": start,
            "organizer": item.get("organizer") or "",
            "thread": thread_id(item.get("url"), item.get("props")),
        })
    found.sort(key=lambda row: row["start"], reverse=True)
    return found


def _web(chat, days):
    import account  # account imports this module
    try:
        page = account.web_page(chat)
        data = page and page.evaluate(_WEB, days)
    except Exception:
        return None
    return data if isinstance(data, dict) and not data.get("error") else None


def _teams_page(chat):
    page = chat.side_page()
    try:
        if not _on_teams(getattr(page, "url", "")):
            page.goto("https://teams.microsoft.com/v2/", wait_until="domcontentloaded", timeout=25_000)
            deadline = time.time() + 15
            while time.time() < deadline and not _on_teams(page.url):
                if "login" in (page.url or ""):
                    return None
                time.sleep(0.4)
    except Exception:
        return None
    return page if _on_teams(getattr(page, "url", "")) else None


def _fetch_text(page, url):
    try:
        response = page.context.request.get(url, timeout=20_000)
    except Exception:
        return ""
    if not response.ok:
        return ""
    body = response.body()
    kind = (response.headers.get("content-type") or "").lower()
    if url.lower().split("?")[0].endswith(".docx") or "wordprocessingml" in kind:
        try:
            return _docx_text(body)
        except Exception:
            return ""
    if "html" in kind and b"WEBVTT" not in body[:300] and b"<v " not in body[:800]:
        return ""
    raw = body.decode("utf-8-sig", "replace")
    if "WEBVTT" in raw[:300] or "<v " in raw[:800]:
        return vtt_to_text(raw)
    if "transcript" in url.lower() and "<html" not in raw[:200].lower():
        return (vtt_to_text(raw) or raw.strip())[:60_000]
    return ""


def _on_teams(url):
    host = urlparse(url or "").hostname or ""
    return host.startswith("teams.") and "login" not in host


def _docx_text(data):
    xml = zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml").decode("utf-8", "replace")
    xml = re.sub(r"</w:p>", "\n", xml)
    return html.unescape(re.sub(r"<[^>]+>", "", xml)).strip()


def _words(text):
    return [word for word in re.findall(r"[a-z0-9']{4,}", (text or "").lower()) if word not in _STOP]


def _day(when):
    return f"{_MONTHS[when.month - 1]} {when.day} {when.strftime('%H:%M')}"


def _at(text):
    text = (text or "").strip()
    if not text:
        return None
    if "T" in text or text.endswith("Z"):
        return _stamp(text)
    try:
        return datetime.strptime(text[:16], "%Y-%m-%d %H:%M")
    except ValueError:
        return None


def _stamp(value):
    raw = str(value or "").replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(raw)
    except ValueError:
        return None
    if parsed.tzinfo:
        parsed = parsed.astimezone().replace(tzinfo=None)
    return parsed


def _load(text):
    start = (text or "").find("{")
    if start < 0:
        return {"error": "failed"}
    try:
        data = json.loads(text[start:])
    except json.JSONDecodeError:
        return {"error": "failed"}
    return data if isinstance(data, dict) else {"error": "failed"}


def _check():
    encoded = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc123%40thread.v2/0?context=1"
    assert thread_id(encoded) == "19:meeting_abc123@thread.v2", thread_id(encoded)
    assert thread_id('{"cid":"19:meeting_abc123@thread.v2"}') == "19:meeting_abc123@thread.v2"
    budget = {"subject": "FY27 Budget review", "start": datetime(2026, 11, 14, 9, 0), "organizer": "Ada", "thread": "19:meeting_a@thread.v2"}
    other = {"subject": "Staff stand-up", "start": datetime(2026, 11, 13, 10, 0), "organizer": "", "thread": ""}
    another = {"subject": "Budget office hours", "start": datetime(2026, 11, 12, 15, 0), "organizer": "Sam", "thread": "19:meeting_b@thread.v2"}
    assert pick([budget, other], "budget")[0]["subject"] == "FY27 Budget review"
    chosen, options = pick([budget, another], "budget")
    assert chosen is None and len(options) == 2
    assert pick([other], "")[0]["subject"] == "Staff stand-up"
    mandate = {"subject": "Mandate Review - Requirements", "start": datetime(2026, 11, 13, 13, 0), "organizer": "Hasan, Nabeel (MPBSDP)", "thread": ""}
    workshop = {"subject": "Mandate Review - Workshop 1", "start": datetime(2026, 11, 3, 13, 0), "organizer": "Sam", "thread": ""}
    assert pick([workshop, mandate], "nabeel for mandate review requirements")[0] is mandate
    facts = from_messages([
        {"type": "RichText/Media_CallRecording", "content": "https://contoso.sharepoint.com/rec.mp4", "props": ""},
        {"type": "RichText/Html", "content": "Transcript https://contoso.sharepoint.com/note.vtt", "props": ""},
    ])
    assert facts["recording"] and facts["transcript_url"].endswith("note.vtt"), facts
    text = vtt_to_text("WEBVTT\n\n00:00:01.000 --> 00:00:03.000\n<v Ada>We decided to ship on Friday.</v>\n")
    clear()
    remember("FY27 Budget review", budget["start"], text)
    block = for_question("what did we decide in the budget review", {"op": "recordings"})
    assert "ship on Friday" in block and "not saved on the PC" in block, block
    assert for_question("open notepad") == ""
    clear()
    assert for_question("what did we decide") == ""
    listed = _list([budget], 7, "")
    assert "FY27 Budget review" in listed and "OneDrive" not in listed and "Teams chat" in listed
    print("meetings ok")


if __name__ == "__main__":
    _check()
