"""Read this user's work account without touching the screen.

Classic Outlook (COM) answers calendar, mail, people, and tasks. New Outlook, or no
classic Outlook running, uses Outlook on the web in the signed-in browser instead.
event creates calendar items the same way. OneDrive files and the work email address are local.
"""

import json
import os
import re
import subprocess
import time
from datetime import datetime, timedelta
from urllib.parse import urlparse

import meetings
import sysinfo

OPS = ("next", "agenda", "mail", "person", "tasks", "files", "me", "recordings", "event")
_DO = re.compile(
    r"\b(open|join|click|send|reply|forward|accept|decline|delete|book|create|cancel|reschedule|"
    r"share|mute|launch|draft|compose)\b",
    re.I,
)
_STATUS = re.compile(
    r"\b(checking|looking up|let me (look|check|find)|i'?ll (look|check|find)|searching (your|the|for))\b",
    re.I,
)
_NEXT = re.compile(
    r"^(?:when'?s?|what'?s?|when is|what is)(?: my| the)? next (?:teams )?meeting$"
    r"|^(?:my )?next meeting$"
    r"|^(?:do i have (?:a |any )?)?(?:an? )?meeting(?:s)? (?:soon|coming up|next)$"
    r"|^(?:what(?:'s| is) my )?upcoming meeting$"
)
_AGENDA = re.compile(
    r"^(?:what(?:'s| is|s) on )?(?:my )?(calendar|schedule|agenda)(?: (today|tomorrow|this week))?$"
    r"|^(?:my )?(meetings|agenda)(?: (today|tomorrow|this week))?$"
    r"|^(?:what meetings do i have|what do i have)(?: (?:on )?(today|tomorrow|this week))?$"
    r"|^(?:what(?:'s| is) my )(schedule|calendar)(?: (today|tomorrow|this week))?$"
)
_UNREAD = re.compile(
    r"^(?:any |do i have (?:any )?|(?:summari[sz]e|show|list|read)(?: me)? (?:my )?)?(?:unread|new) (?:email|emails|mail|messages)$"
    r"|^check my (?:email|inbox|mail)$"
    r"|^do i have (?:any )?mail$"
)
_FROM = re.compile(r"^(?:emails?|mail) from ([a-z][\w .'-]{1,40})$")
_TASKS = re.compile(r"^(?:my )?(?:tasks|to-?dos?|flagged (?:email|emails|mail))$|^(?:what(?:'s| is) on my (?:to-?do|task) list)$")
_FILES = re.compile(
    r"^(?:my )?(?:recent|latest) (?:files|documents)$"
    r"|^(?:what(?: files)? (?:was|have) i (?:been )?(?:working on|opened))$"
    r"|^(?:my onedrive files)$"
)
_ME = (
    (re.compile(r"^who(?:'s| is) my manager$"), "manager"),
    (re.compile(r"^what(?:'s| is) my (?:job )?title$"), "title"),
    (re.compile(r"^(?:what(?:'s| is) my (?:email|email address)|my email(?: address)?)$"), "email"),
)
# "what about tomorrow" only means the calendar when the last answer was one.
_PREFIX = re.compile(r"^(?:what about |how about |and |also )")
_WHEN_ONLY = re.compile(r"^(?:(?:the )?same (?:thing )?for |for )?(today|tomorrow|this week)$")
_REC_ONE = re.compile(r"^(?:(?:read|get|save|load|show) )?(?:the )?(?:transcript|recording) (?:of|for|from) (.+)$")
# Loose on purpose: "transcrips", "transcripts", "recorded".
_REC_LIST = re.compile(r"\b(?:recordings?|recorded|transcri(?:pts?|ps?|ptions?))\b")
_REC_TOPIC = re.compile(r"\b(?:for|about|on|called|named|of|with)\s+(.+)$")
_RECENTISH = re.compile(r"\b(few days|last few days|past few days|recent ones|back ?dated)\b")
_VAGUE_REC = re.compile(r"^(the |my )?(recent ones?|few days( ago)?|last few days|that one|it|same)$")

# ponytail: mail search looks at the 50 newest inbox items, not the whole mailbox.
_PS = r"""
$ErrorActionPreference = 'Stop'
$op = '__OP__'
$when = '__WHEN__'
$q = __QUERY__
function Hidden($m) {
  try { if ($m.MeetingStatus -eq 5) { return $true } } catch {}
  try { if ($m.ResponseStatus -eq 4) { return $true } } catch {}
  return $false
}
function Out-Json($obj) { ConvertTo-Json -InputObject $obj -Compress -Depth 4 }
try {
  try { $o = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { $o = $null }
  if (-not $o) { Out-Json @{error='outlook'}; return }
  $ns = $o.GetNamespace('MAPI')
  if ($op -eq 'me') {
    $u = $ns.CurrentUser.AddressEntry.GetExchangeUser()
    if (-not $u) { Out-Json @{error='none'}; return }
    $mgr = ''
    try { if ($u.Manager) { $mgr = [string]$u.Manager.Name } } catch {}
    Out-Json @{name=[string]$u.Name; email=[string]$u.PrimarySmtpAddress; title=[string]$u.JobTitle;
      department=[string]$u.Department; manager=$mgr}
    return
  }
  if ($op -eq 'person') {
    $r = $ns.CreateRecipient($q)
    $null = $r.Resolve()
    if (-not $r.Resolved) { Out-Json @{error='none'}; return }
    $u = $r.AddressEntry.GetExchangeUser()
    if (-not $u) { Out-Json @{name=[string]$r.Name; email=[string]$r.Address}; return }
    $mgr = ''
    try { if ($u.Manager) { $mgr = [string]$u.Manager.Name } } catch {}
    Out-Json @{name=[string]$u.Name; email=[string]$u.PrimarySmtpAddress; title=[string]$u.JobTitle;
      department=[string]$u.Department; manager=$mgr}
    return
  }
  if ($op -eq 'tasks') {
    $items = New-Object System.Collections.Generic.List[object]
    try {
      $flagged = $ns.GetDefaultFolder(6).Items
      $flagged.Sort('[ReceivedTime]', $true)
      foreach ($m in $flagged.Restrict('[FlagStatus] = 2')) {
        $items.Add([pscustomobject]@{kind='flagged'; title=[string]$m.Subject; who=[string]$m.SenderName; due=''})
        if ($items.Count -ge 8) { break }
      }
    } catch {}
    try {
      foreach ($t in $ns.GetDefaultFolder(13).Items) {
        if ($items.Count -ge 16) { break }
        try { if ($t.Complete) { continue } } catch {}
        $due = ''
        try { if ($t.DueDate -and $t.DueDate.Year -gt 2000 -and $t.DueDate.Year -lt 2100) { $due = $t.DueDate.ToString('MM-dd') } } catch {}
        $items.Add([pscustomobject]@{kind='task'; title=[string]$t.Subject; who=''; due=$due})
      }
    } catch {}
    Out-Json @{items=@($items.ToArray())}
    return
  }
  if ($op -eq 'mail') {
    $inbox = $ns.GetDefaultFolder(6)
    $all = $inbox.Items
    $all.Sort('[ReceivedTime]', $true)
    $items = New-Object System.Collections.Generic.List[object]
    if ($q) {
      $like = $q.Replace('[','[[]').Replace('*','[*]').Replace('?','[?]')
      $n = 0
      foreach ($m in $all) {
        $n++
        if ($n -gt 50) { break }
        $from = [string]$m.SenderName
        if ($m.Subject -like "*$like*" -or $from -like "*$like*") {
          $items.Add([pscustomobject]@{from=$from; subject=[string]$m.Subject; received=$m.ReceivedTime.ToString('MM-dd HH:mm')})
          if ($items.Count -ge 8) { break }
        }
      }
      Out-Json @{items=@($items.ToArray()); query=$q}
      return
    }
    $unread = $all.Restrict('[UnRead] = True')
    for ($i = 1; $i -le [math]::Min(8, $unread.Count); $i++) {
      $m = $unread.Item($i)
      $items.Add([pscustomobject]@{from=[string]$m.SenderName; subject=[string]$m.Subject; received=$m.ReceivedTime.ToString('MM-dd HH:mm')})
    }
    Out-Json @{items=@($items.ToArray()); unread=[int]$inbox.UnReadItemCount}
    return
  }
  $now = Get-Date
  $skipAllDay = $false
  $cap = 12
  if ($op -eq 'next') { $from = $now.AddHours(-6); $to = $now.Date.AddDays(14); $skipAllDay = $true; $cap = 15 }
  elseif ($when -eq 'tomorrow') { $from = $now.Date.AddDays(1); $to = $from.AddDays(1) }
  elseif ($when -eq 'week') { $from = $now.Date; $to = $from.AddDays(7); $cap = 20 }
  else { $from = $now.Date; $to = $from.AddDays(1) }
  $cal = $ns.GetDefaultFolder(9).Items
  $cal.IncludeRecurrences = $true
  $cal.Sort('[Start]')
  $found = $cal.Restrict("[End] > '$($from.ToString('g'))' AND [Start] < '$($to.ToString('g'))'")
  $items = New-Object System.Collections.Generic.List[object]
  foreach ($m in $found) {
    if (Hidden $m) { continue }
    $allday = $false
    try { $allday = [bool]$m.AllDayEvent } catch {}
    if ($skipAllDay -and $allday) { continue }
    $online = $false
    try { $online = [bool]$m.IsOnlineMeeting } catch {}
    $tentative = $false
    try { $tentative = ($m.ResponseStatus -eq 2) } catch {}
    $items.Add([pscustomobject]@{
      subject=[string]$m.Subject; start=$m.Start.ToString('yyyy-MM-dd HH:mm'); end=$m.End.ToString('yyyy-MM-dd HH:mm');
      location=[string]$m.Location; organizer=[string]$m.Organizer; online=$online; allday=$allday; tentative=$tentative
    })
    if ($items.Count -ge $cap) { break }
  }
  Out-Json @{items=@($items.ToArray())}
} catch {
  Out-Json @{error='failed'; detail=("" + $_.Exception.Message)}
}
"""

_FILES_PS = r"""
$roots = @($env:OneDriveCommercial, $env:OneDrive, $env:OneDriveConsumer) | Where-Object { $_ -and (Test-Path $_) }
if (-not $roots) { ConvertTo-Json @{error='onedrive'} -Compress; return }
$shell = New-Object -ComObject WScript.Shell
$items = New-Object System.Collections.Generic.List[object]
Get-ChildItem "$env:APPDATA\Microsoft\Windows\Recent\*.lnk" -ErrorAction SilentlyContinue |
  Sort-Object LastWriteTime -Descending | Select-Object -First 40 | ForEach-Object {
    if ($items.Count -ge 10) { return }
    $target = $shell.CreateShortcut($_.FullName).TargetPath
    if (-not $target) { return }
    foreach ($root in $roots) {
      if ($target.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
        $items.Add([pscustomobject]@{name=[IO.Path]::GetFileName($target); path=$target; when=$_.LastWriteTime.ToString('MM-dd HH:mm')})
        break
      }
    }
  }
ConvertTo-Json -InputObject @{items=@($items.ToArray())} -Compress -Depth 3
"""

_WEB = r"""
async ([op, when, query]) => {
  const headers = {Accept: "application/json", Prefer: 'outlook.timezone="UTC"'};
  const stamp = (value) => {
    if (!value) return "";
    const raw = value.DateTime || String(value);
    const utc = !value.TimeZone || /utc/i.test(value.TimeZone || "");
    const date = new Date(utc && !/z$/i.test(raw) ? raw + "Z" : raw);
    if (Number.isNaN(date.getTime())) return "";
    const p = (n) => String(n).padStart(2, "0");
    return `${date.getFullYear()}-${p(date.getMonth() + 1)}-${p(date.getDate())} ${p(date.getHours())}:${p(date.getMinutes())}`;
  };
  if (op === "mail") {
    const url = query
      ? "/api/v2.0/me/mailfolders/inbox/messages?$top=30&$select=Subject,From,ReceivedDateTime&$orderby=ReceivedDateTime desc"
      : "/api/v2.0/me/mailfolders/inbox/messages?$top=8&$filter=IsRead eq false&$select=Subject,From,ReceivedDateTime&$orderby=ReceivedDateTime desc";
    const response = await fetch(url, {credentials: "include", headers});
    if (!response.ok) return {error: "web", detail: String(response.status)};
    const body = await response.json();
    const needle = (query || "").toLowerCase();
    const items = [];
    for (const m of body.value || []) {
      const from = (m.From && m.From.EmailAddress && m.From.EmailAddress.Name) || "";
      const subject = m.Subject || "";
      if (needle && !from.toLowerCase().includes(needle) && !subject.toLowerCase().includes(needle)) continue;
      items.push({from, subject, received: stamp({DateTime: m.ReceivedDateTime, TimeZone: "UTC"}).slice(5)});
      if (items.length >= 8) break;
    }
    return query ? {items, query} : {items, unread: items.length};
  }
  const now = new Date();
  let start = new Date(now.getTime() - 6 * 3600 * 1000);
  let end = new Date(now.getTime() + 14 * 86400000);
  if (op !== "next") {
    start = new Date(now); start.setHours(0, 0, 0, 0);
    if (when === "tomorrow") start.setDate(start.getDate() + 1);
    end = new Date(start);
    end.setDate(end.getDate() + (when === "week" ? 7 : 1));
  }
  const iso = (d) => d.toISOString().replace(/\.\d{3}Z$/, "Z");
  const url = "/api/v2.0/me/calendarview?startDateTime=" + encodeURIComponent(iso(start))
    + "&endDateTime=" + encodeURIComponent(iso(end))
    + "&$select=Subject,Start,End,Location,Organizer,IsAllDay,IsOnlineMeeting,ResponseStatus&$orderby=Start/DateTime&$top=20";
  const response = await fetch(url, {credentials: "include", headers});
  if (!response.ok) return {error: "web", detail: String(response.status)};
  const body = await response.json();
  const items = [];
  for (const m of body.value || []) {
    const subject = m.Subject || "";
    if (/^canceled:/i.test(subject)) continue;
    const loc = m.Location && (m.Location.DisplayName || m.Location) || "";
    items.push({
      subject, start: stamp(m.Start), end: stamp(m.End),
      location: typeof loc === "string" ? loc : "",
      organizer: (m.Organizer && m.Organizer.EmailAddress && m.Organizer.EmailAddress.Name) || "",
      online: Boolean(m.IsOnlineMeeting), allday: Boolean(m.IsAllDay),
      tentative: (m.ResponseStatus && m.ResponseStatus.Response) === "TentativelyAccepted",
    });
  }
  return {items};
}
"""


def _known(text):
    """An account op this sentence asks for on its own, or None."""
    if _NEXT.match(text):
        return {"op": "next"}
    found = _AGENDA.match(text)
    if found:
        when = next((g for g in found.groups() if g in ("today", "tomorrow", "this week")), None)
        return {"op": "agenda", "when": {"this week": "week"}.get(when, when or "today")}
    if _UNREAD.match(text):
        return {"op": "mail"}
    found = _FROM.match(text)
    if found:
        return {"op": "mail", "query": found.group(1).strip()}
    if _TASKS.match(text):
        return {"op": "tasks"}
    if _FILES.match(text):
        return {"op": "files"}
    for pattern, focus in _ME:
        if pattern.match(text):
            return {"op": "me", "focus": focus}
    return None


def _span(text):
    found = re.search(r"\blast (\d+) days\b", text)
    if found:
        return max(1, min(30, int(found.group(1))))
    if re.search(r"\btoday\b", text):
        return 1
    if re.search(r"\byesterday\b", text):
        return 2
    if re.search(r"\b(latest|last|recent|most recent)\b", text):
        return 14
    return 7


def _recording(text, last, prior):
    """A recordings read, including 'the recent ones' after a recording answer."""
    found = _REC_ONE.match(text)
    if found:
        query = re.sub(r"^(the |my |our |that )", "", found.group(1).strip())
        out = {"op": "recordings", "days": _span(text)}
        if query and not _VAGUE_REC.match(query):
            out["query"] = query[:80]
        return out
    talked = re.search(r"\b(record\w*|transcript\w*)\b", prior or "", re.I) or (last or {}).get("op") == "recordings"
    if _REC_LIST.search(text) or (_RECENTISH.search(text) and talked):
        out = {"op": "recordings", "days": _span(text)}
        topic = _REC_TOPIC.search(text)
        if topic and not _VAGUE_REC.match(topic.group(1)):
            # Names and title words together; meetings.pick scores them against subject and organizer.
            out["query"] = topic.group(1).strip()[:80]
        return out
    return None


def match(task, last=None, prior=""):
    """A direct account read for a short question, or None when the agent should handle it.

    `last` is the previous account op, so "tomorrow" after a meeting answer stays local.
    `prior` is the last Co-opilot reply, so "the recent ones" after a recording answer stays local.
    ponytail: only a single short question, with no "and" and no request to click or send.
    Anything else goes to the agent, which can still call account. "next week" is not a when-word.
    """
    text = " ".join(str(task).lower().replace("\u2019", "'").split()).rstrip("?.!")
    if not text or len(text) > 140 or ". " in text or " and " in text or _DO.search(text):
        return None
    found = _known(text)
    if found:
        return found
    stripped = _PREFIX.sub("", text, count=1)
    if stripped != text:
        found = _known(stripped)
        if found:
            return found
        text = stripped
    found = _recording(text, last, prior)
    if found:
        return found
    if not last or last.get("op") not in ("next", "agenda"):
        return None
    found = _WHEN_ONLY.match(text)
    if not found:
        return None
    word = found.group(1)
    return {"op": "agenda", "when": "week" if word == "this week" else word}


def status_only(summary):
    """True when a reply only says it is checking, and never gives a fact."""
    text = " ".join((summary or "").split())
    return bool(text) and len(text) <= 280 and not re.search(r"\d", text) and bool(_STATUS.search(text))


def run(action, chat=None):
    """User-facing answer for one account op. `chat` enables the Outlook on the web fallback."""
    op = action["op"]
    if op == "recordings":
        return meetings.answer(action, chat)
    if op == "files":
        return present(action, _load(_files_raw()))
    if op == "me" and action.get("focus") == "email":
        return present(action, {"email": _upn()})
    if op == "event":
        return event(action, chat)
    if chat is not None and op in ("next", "agenda", "mail") and outlook_order(_running(), True)[0] == "web":
        web = _web(chat, action)
        if web:
            return present(action, web)
    data = _load(_outlook(action))
    if data.get("error") == "outlook" and chat is not None and op in ("next", "agenda", "mail"):
        web = _web(chat, action)
        if web:
            data = web
        else:
            data["web"] = "tried"
    if op == "me" and data.get("error") and action.get("focus") in (None, "", "email"):
        email = _upn()
        if email:
            data = {"email": email}
    return present(action, data)


def present(action, data, now=None):
    """Format one account payload. `now` is injectable for the self-check."""
    now = now or datetime.now()
    op = action.get("op")
    if data.get("error") == "outlook":
        if data.get("web") == "tried":
            return ("I couldn't read that from your work account. Classic Outlook isn't running, and Outlook "
                    "on the web didn't answer. Open Outlook, or ask me to look at it on the screen.")
        return ("Classic Outlook isn't running, so I can't read that from your account. "
                "Open Outlook and ask again, or ask me to look at it on the screen.")
    if data.get("error") == "onedrive":
        return "Your OneDrive folder isn't on this PC, so I can't list work files from here."
    if data.get("error") == "none" and op == "person":
        return f"No one in the directory matched **{action.get('query', '').strip()}**."
    if data.get("error") == "none":
        return "Outlook doesn't have that on your account."
    if data.get("error"):
        detail = str(data.get("detail") or data["error"])[:200]
        return f"I couldn't read that from your account. {detail}"
    if op == "me":
        return _present_me(data, action.get("focus") or "")
    if op == "person":
        return _present_me(data, "")
    if op in ("next", "agenda"):
        return _present_cal(op, action.get("when") or "today", _items(data), now)
    if op == "mail":
        return _present_mail(action, data)
    if op == "tasks":
        return _present_tasks(_items(data))
    if op == "files":
        return _present_files(_items(data))
    return "I couldn't read that from your account."


def _booking(item):
    """A desk or room reservation, not a meeting with people."""
    subject = str(item.get("subject") or "")
    return bool(re.search(r"\breserved\b|^workstation\b|^desk\b", subject, re.I)
                or re.search(r"no-?reply", str(item.get("organizer") or ""), re.I))


_CANCELLED = re.compile(r"^\s*(?:\[?cancell?ed\]?:?)\s", re.I)


def _cancelled(item):
    return bool(_CANCELLED.match(str(item.get("subject") or "") + " "))


def _present_cal(op, when, items, now):
    clean, desks = [], []
    for item in items:
        if _cancelled(item):
            continue
        if _booking(item):
            if op != "next" and when != "week":
                desks.append(re.sub(r"\s*reserved\b.*$", "", _name(item), flags=re.I))
            continue
        try:
            _at(item["start"])
            _at(item.get("end") or item["start"])
        except (KeyError, TypeError, ValueError):
            continue
        if not item.get("end"):
            item["end"] = item["start"]
        clean.append(item)
    items = _dedupe(clean)
    if op == "next":
        items = [i for i in items if not i.get("allday")]
        current = [i for i in items if _at(i["start"]) <= now < _at(i["end"])]
        upcoming = [i for i in items if _at(i["start"]) > now]
        lines = []
        if current:
            lines.append(f"You're in **{_name(current[-1])}** until {_at(current[-1]['end']).strftime('%H:%M')}.")
        if not upcoming:
            lines.append("You have no more meetings in the next 14 days." if not current
                         else "Nothing else is scheduled after it in the next 14 days.")
            return "\n".join(lines)
        nxt = upcoming[0]
        lines.append(f"Your next meeting is **{_name(nxt)}** {_phrase(nxt, now)}.")
        lines.extend(_where(nxt))
        return "\n".join(lines)
    desk = f"Desk booked: {', '.join(dict.fromkeys(desks))}." if desks else ""
    if not items:
        return "\n".join(filter(None, [f"No meetings {_range_word(when)}.", desk]))
    title = {"today": "today", "tomorrow": "tomorrow", "week": "this week"}.get(when, "today")
    lines = [f"You have **{len(items)}** meeting{'' if len(items) == 1 else 's'} {title}."] + ([desk] if desk else [])
    day = None
    for item in items:
        start = _at(item["start"])
        if when == "week" and start.date() != day:
            day = start.date()
            lines.append(f"## {_day_label(day, now.date())}")
        if item.get("allday"):
            lines.append(f"- All day **{_name(item)}**")
        else:
            extra = " (tentative)" if item.get("tentative") else ""
            loc = _short(item.get("location"))
            lines.append(f"- {start.strftime('%H:%M')}–{_at(item['end']).strftime('%H:%M')} **{_name(item)}**{extra}"
                         + (f" ({loc})" if loc else ""))
    return "\n".join(lines)


def _present_mail(action, data):
    items = _items(data)
    query = action.get("query") or data.get("query")
    if query:
        if not items:
            return f"No recent inbox mail matches **{query}**."
        lines = [f"Recent mail matching **{query}**:"]
    else:
        count = data.get("unread")
        if count is None:
            count = len(items)
        if not count:
            return "You have no unread email."
        lines = [f"You have **{count}** unread email{'' if count == 1 else 's'}."]
    for item in items:
        lines.append(f"- {item.get('received', '')} **{item.get('from') or 'Unknown'}**: {item.get('subject') or '(no subject)'}")
    return "\n".join(lines)


def _present_tasks(items):
    if not items:
        return "You have no flagged mail and no open Outlook tasks."
    lines = [f"You have **{len(items)}** open item{'' if len(items) == 1 else 's'}."]
    for item in items:
        due = f" (due {item['due']})" if item.get("due") else ""
        who = f" — {item['who']}" if item.get("who") else ""
        kind = "Flagged" if item.get("kind") == "flagged" else "Task"
        lines.append(f"- {kind}: **{item.get('title') or '(no subject)'}**{who}{due}")
    return "\n".join(lines)


def _present_files(items):
    if not items:
        return "No recently opened OneDrive files."
    lines = ["Recent files in your OneDrive:"]
    for item in items:
        lines.append(f"- {item.get('when', '')} **{item.get('name') or 'file'}**")
        if item.get("path"):
            lines.append(f"  `{item['path']}`")
    return "\n".join(lines)


def _present_me(data, focus):
    name = data.get("name") or "You"
    email = data.get("email") or ""
    title = data.get("title") or ""
    dept = data.get("department") or ""
    manager = data.get("manager") or ""
    if focus == "manager":
        return f"Your manager is **{manager}**." if manager else "Outlook doesn't list a manager on your account."
    if focus == "title":
        role = ", ".join(p for p in (title, dept) if p)
        return f"Your title is **{role}**." if role else "Outlook doesn't list a title on your account."
    if focus == "email":
        return f"Your work email is **{email}**." if email else "I couldn't read your work email from this PC."
    bits = [p for p in (title, dept) if p]
    lines = [f"**{name}**" + (f" ({email})" if email else "")]
    if bits:
        lines.append(f"- {', '.join(bits)}")
    if manager:
        lines.append(f"- Reports to {manager}")
    return "\n".join(lines)


def _phrase(item, now):
    start = _at(item["start"])
    clock = start.strftime("%H:%M")
    mins = int((start - now).total_seconds() // 60)
    tentative = " (tentative)" if item.get("tentative") else ""
    if start.date() == now.date():
        soon = f" (in {mins} min)" if 0 <= mins < 180 else ""
        return f"today at {clock}{soon}{tentative}"
    if start.date() == now.date() + timedelta(days=1):
        return f"tomorrow at {clock}{tentative}"
    return start.strftime("%a %b ") + str(start.day) + start.strftime(" at %H:%M") + tentative


def _where(item):
    lines = []
    loc = _short(item.get("location"))
    if loc:
        lines.append(f"- {loc}")
    elif item.get("online"):
        lines.append("- Teams meeting")
    if item.get("organizer"):
        lines.append(f"- With {item['organizer']}")
    return lines


def _day_label(day, today):
    if day == today:
        return "Today"
    if day == today + timedelta(days=1):
        return "Tomorrow"
    return day.strftime("%A ") + str(day.day) + day.strftime(" %b")


def _range_word(when):
    return {"today": "today", "tomorrow": "tomorrow", "week": "this week"}.get(when, "then")


def _name(item):
    return item.get("subject") or "(no subject)"


def _short(value):
    value = " ".join(str(value or "").split())
    return value if len(value) <= 80 else value[:77] + "..."


def _at(text):
    return datetime.strptime(text, "%Y-%m-%d %H:%M")


def _items(data):
    items = data.get("items") or []
    if isinstance(items, dict):
        items = [items]
    return [i for i in items if isinstance(i, dict)]


def _dedupe(items):
    seen, out = set(), []
    for item in items:
        key = (item.get("subject"), item.get("start"))
        if key in seen:
            continue
        seen.add(key)
        out.append(item)
    return out


def _outlook(action):
    op = action["op"] if action["op"] in OPS else "next"
    when = action.get("when") if action.get("when") in ("today", "tomorrow", "week") else "today"
    script = _PS.replace("__OP__", op).replace("__WHEN__", when).replace("__QUERY__", _ps_quote(action.get("query") or ""))
    return sysinfo._powershell("account", script)


def _files_raw():
    return sysinfo._powershell("files", _FILES_PS)


def _load(text):
    start = (text or "").find("{")
    if start < 0:
        return {"error": "failed", "detail": (text or "no output")[:200]}
    try:
        data = json.loads(text[start:])
    except json.JSONDecodeError:
        return {"error": "failed", "detail": text[:200]}
    return data if isinstance(data, dict) else {"error": "failed", "detail": "bad result"}


def _ps_quote(value):
    clean = re.sub(r"[\x00-\x1f]", " ", str(value))[:80]
    return "'" + clean.replace("'", "''") + "'"


def _upn():
    try:
        out = subprocess.run(["whoami.exe", "/upn"], capture_output=True, text=True, timeout=5,
                             creationflags=subprocess.CREATE_NO_WINDOW).stdout.strip()
    except (OSError, subprocess.TimeoutExpired):
        return ""
    return out.lower() if "@" in out else ""


def _web(chat, action):
    """Outlook on the web, using the browser the user is already signed in with. None if it can't."""
    try:
        page = web_page(chat)
        data = page and page.evaluate(_WEB, [action["op"], action.get("when") or "today", action.get("query") or ""])
    except Exception:
        return None
    finally:
        _front(chat)
    if not isinstance(data, dict) or data.get("error"):
        return None
    return data


def _front(chat):
    try:
        chat.front()
    except Exception:
        pass


_OUTLOOK_HOSTS = ("outlook.office.com", "outlook.office365.com", "outlook.cloud.microsoft")


def _on_outlook(url):
    host = (urlparse(url or "").hostname or "").lower()
    return any(host == h or host.endswith("." + h) for h in _OUTLOOK_HOSTS)


# Outlook on the web (and new Outlook, which is the same service) rejects cookie-only API calls with 401.
# Its own requests carry a bearer token; we borrow it and add it to fetch calls for /api/v2.0 on that page.
_token = {"value": "", "at": 0.0}
_TOKEN_LIFE = 30 * 60
_ARM = """token => {
  window.__coAuth = token;
  if (window.__coArmed) return;
  window.__coArmed = true;
  const original = window.fetch;
  window.fetch = (url, opts = {}) => {
    if (String(url).includes("/api/v2.0/")) {
      opts = {...opts, headers: {...(opts.headers || {}), Authorization: window.__coAuth}};
    }
    return original(url, opts);
  };
}"""


def web_page(chat):
    """Outlook on the web in the background tab, with fetch carrying Outlook's own token. None if it can't."""
    page = chat.side_page()
    seen = []

    def grab(request):
        auth = request.headers.get("authorization", "")
        if auth.lower().startswith("bearer ") and _on_outlook(request.url):
            seen.append(auth)

    fresh = bool(_token["value"]) and time.time() - _token["at"] < _TOKEN_LIFE
    page.on("request", grab)
    try:
        if not fresh or not _on_outlook(page.url):
            page.goto("https://outlook.cloud.microsoft/calendar/", wait_until="domcontentloaded", timeout=30_000)
        deadline = time.time() + 25
        while time.time() < deadline and not seen and not (fresh and _on_outlook(page.url)):
            page.wait_for_timeout(400)  # not time.sleep: request events only fire while Playwright waits
    finally:
        page.remove_listener("request", grab)
    if seen:
        _token.update(value=seen[-1], at=time.time())
    if not _on_outlook(page.url) or not _token["value"]:
        return None
    page.evaluate(_ARM, _token["value"])
    return page


# Creating events: classic Outlook through COM, new Outlook through the web API, an .ics file as the last resort.
DAYS = ("monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday")
REPEATS = ("none", "daily", "weekdays", "weekly")

_EVENT_PS = r"""
$ErrorActionPreference = 'Stop'
function Out-Json($obj) { ConvertTo-Json -InputObject $obj -Compress }
try {
  try { $o = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
  catch { $o = New-Object -ComObject Outlook.Application }
  $start = [datetime]::ParseExact('__START__', 'yyyy-MM-dd HH:mm', $null)
  $a = $o.CreateItem(1)
  $a.Subject = __SUBJECT__
  $a.Start = $start
  $a.Duration = __MINUTES__
  $a.ReminderSet = $true
  $a.ReminderMinutesBeforeStart = __REMINDER__
  if (__FREE__) { $a.BusyStatus = 0 }
  if (__TYPE__ -ge 0) {
    $p = $a.GetRecurrencePattern()
    $p.RecurrenceType = __TYPE__
    if (__TYPE__ -eq 1) { $p.DayOfWeekMask = __MASK__ }
    $p.PatternStartDate = $start.Date
    $p.StartTime = $start
    $p.Duration = __MINUTES__
    if ('__UNTIL__') { $p.PatternEndDate = [datetime]::ParseExact('__UNTIL__', 'yyyy-MM-dd', $null) } else { $p.NoEndDate = $true }
  }
  $a.Save()
  Out-Json @{ok=$true; subject=[string]$a.Subject; start=$a.Start.ToString('yyyy-MM-dd HH:mm')}
} catch {
  Out-Json @{error='failed'; detail=("" + $_.Exception.Message)}
}
"""

_EVENT_WEB = """async body => {
  const headers = {"Content-Type": "application/json", Accept: "application/json"};
  const made = await fetch("/api/v2.0/me/events", {method: "POST", credentials: "include", headers, body: JSON.stringify(body)});
  if (!made.ok) return {error: "web", detail: made.status + " " + (await made.text()).slice(0, 200)};
  const item = await made.json();
  const back = await fetch("/api/v2.0/me/events/" + encodeURIComponent(item.Id) + "?$select=Subject", {credentials: "include", headers});
  return back.ok ? {ok: true, subject: (await back.json()).Subject} : {error: "verify", detail: String(back.status)};
}"""


def _days(action):
    if action.get("repeat") == "weekdays":
        return list(DAYS[:5])
    return action.get("days") or [DAYS[datetime.strptime(action["start"], "%Y-%m-%d %H:%M").weekday()]]


def _event_ps(action):
    repeat = action.get("repeat", "none")
    mask = sum(1 << ((DAYS.index(d) + 1) % 7) for d in _days(action))  # Outlook: Sunday 1, Monday 2 ... Saturday 64
    kind = {"none": -1, "daily": 0}.get(repeat, 1)
    return (_EVENT_PS.replace("__SUBJECT__", _ps_quote(action["subject"])).replace("__START__", action["start"])
            .replace("__MINUTES__", str(int(action.get("minutes", 30)))).replace("__REMINDER__", str(int(action.get("reminder", 15))))
            .replace("__FREE__", "$true" if action.get("free") else "$false").replace("__TYPE__", str(kind))
            .replace("__MASK__", str(mask)).replace("__UNTIL__", action.get("until") or ""))


def _event_body(action, zone):
    start = datetime.strptime(action["start"], "%Y-%m-%d %H:%M")
    end = start + timedelta(minutes=int(action.get("minutes", 30)))
    body = {"Subject": action["subject"],
            "Start": {"DateTime": start.isoformat(), "TimeZone": zone},
            "End": {"DateTime": end.isoformat(), "TimeZone": zone},
            "IsReminderOn": True, "ReminderMinutesBeforeStart": int(action.get("reminder", 15)),
            "ShowAs": "Free" if action.get("free") else "Busy"}
    repeat = action.get("repeat", "none")
    if repeat != "none":
        pattern = ({"Type": "Daily", "Interval": 1} if repeat == "daily" else
                   {"Type": "Weekly", "Interval": 1, "DaysOfWeek": [d.title() for d in _days(action)], "FirstDayOfWeek": "Sunday"})
        span = ({"Type": "EndDate", "StartDate": start.date().isoformat(), "EndDate": action["until"]} if action.get("until")
                else {"Type": "NoEnd", "StartDate": start.date().isoformat()})
        body["Recurrence"] = {"Pattern": pattern, "Range": span}
    return body


def _ics(action, stamp=None):
    start = datetime.strptime(action["start"], "%Y-%m-%d %H:%M")
    end = start + timedelta(minutes=int(action.get("minutes", 30)))
    summary = re.sub(r"([,;\\])", r"\\\1", action["subject"])
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Copilot Assistant//EN", "BEGIN:VEVENT",
             f"UID:{start:%Y%m%dT%H%M}-{abs(hash(action['subject']))}@copilot-assistant",
             f"DTSTAMP:{(stamp or datetime.utcnow()):%Y%m%dT%H%M%SZ}",
             f"DTSTART:{start:%Y%m%dT%H%M%S}", f"DTEND:{end:%Y%m%dT%H%M%S}", f"SUMMARY:{summary}"]
    repeat = action.get("repeat", "none")
    if repeat != "none":
        rule = "FREQ=DAILY" if repeat == "daily" else "FREQ=WEEKLY;BYDAY=" + ",".join(d[:2].upper() for d in _days(action))
        if action.get("until"):
            rule += f";UNTIL={action['until'].replace('-', '')}T235959"
        lines.append("RRULE:" + rule)
    if action.get("free"):
        lines.append("TRANSP:TRANSPARENT")
    lines += ["BEGIN:VALARM", f"TRIGGER:-PT{int(action.get('reminder', 15))}M", "ACTION:DISPLAY",
              f"DESCRIPTION:{summary}", "END:VALARM", "END:VEVENT", "END:VCALENDAR"]
    return "\r\n".join(lines) + "\r\n"


def _running():
    try:
        out = subprocess.run(["tasklist", "/fo", "csv", "/nh"], capture_output=True, text=True, timeout=10,
                             creationflags=subprocess.CREATE_NO_WINDOW).stdout
    except (OSError, subprocess.TimeoutExpired):
        return set()
    return {line.split('","')[0].strip('"').lower() for line in out.splitlines() if line}


def _classic_installed():
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Outlook.Application\CLSID"))
        return True
    except OSError:
        return False


def outlook_order(running, classic):
    """Which Outlook to use first: the one that is open, else classic if installed and new Outlook isn't open."""
    if "outlook.exe" in running:
        return ["classic", "web"]
    if "olk.exe" in running or not classic:
        return ["web"] + (["classic"] if classic else [])
    return ["classic", "web"]


def _zone():
    try:
        return subprocess.run(["tzutil", "/g"], capture_output=True, text=True, timeout=5,
                              creationflags=subprocess.CREATE_NO_WINDOW).stdout.strip() or "Eastern Standard Time"
    except (OSError, subprocess.TimeoutExpired):
        return "Eastern Standard Time"


def event(action, chat=None):
    """Create a calendar event or reminder in whichever Outlook works. User-facing text."""
    order = outlook_order(_running(), _classic_installed())
    for way in order:
        if way == "classic":
            data = _load(sysinfo._powershell("event", _event_ps(action)))
        elif chat is not None:
            try:
                page = web_page(chat)
                data = page.evaluate(_EVENT_WEB, _event_body(action, _zone())) if page else {"error": "web"}
            except Exception as exc:
                data = {"error": "web", "detail": str(exc)}
            finally:
                _front(chat)
        else:
            continue
        if data.get("ok"):
            where = "classic Outlook" if way == "classic" else "Outlook"
            return f"Added **{data.get('subject') or action['subject']}** to your {where} calendar.\n{_when_text(action)}"
    import files
    path = files.resolve(os.path.join("%USERPROFILE%", "Documents", (re.sub(r"[^\w -]", "", action["subject"])[:60] or "Event") + ".ics"))
    path.write_text(_ics(action), encoding="utf-8")
    files.touched.append(path)
    os.startfile(str(path))
    return (f"Outlook didn't let me add **{action['subject']}** directly, so I opened it in your calendar app. "
            f"Choose **Save** there to keep it.\n{_when_text(action)}")


def _when_text(action):
    start = datetime.strptime(action["start"], "%Y-%m-%d %H:%M")
    repeat = action.get("repeat", "none")
    if repeat == "none":
        when = f"{start:%A, %B} {start.day} at {start:%H:%M}"
    elif repeat == "daily":
        when = f"Every day at {start:%H:%M}, from {start:%B} {start.day}"
    else:
        names = [d.title() for d in _days(action)]
        when = f"Every {', '.join(names)} at {start:%H:%M}, from {start:%B} {start.day}"
    if action.get("until"):
        when += f" until {action['until']}"
    return f"- {when}\n- Reminder {int(action.get('reminder', 15))} min before"


def glance(now=None, end_hour=17):
    """Start-screen facts: next meeting, unread mail, free time left today. Empty parts when Outlook is closed."""
    now = now or datetime.now()
    cal = _load(_outlook({"op": "next"}))
    mail = _load(_outlook({"op": "mail"}))
    out = {"outlook": cal.get("error") != "outlook"}
    items = [i for i in _items(cal) if not _booking(i) and not _cancelled(i)]
    out.update(_glance_cal(items, now, end_hour))
    if "unread" in mail:
        latest = _items(mail)[:1]
        out["mail"] = {"unread": int(mail["unread"]), "from": latest[0].get("from", "") if latest else ""}
    return out


def _glance_cal(items, now, end_hour=17):
    """Next meeting plus free minutes and the first 30+ minute gap between now and end of workday."""
    spans = []
    for item in items:
        try:
            spans.append((_at(item["start"]), _at(item["end"]), item))
        except (KeyError, ValueError):
            continue
    spans.sort(key=lambda s: s[0])
    out = {}
    upcoming = [s for s in spans if s[1] > now]
    if upcoming:
        start, end, item = upcoming[0]
        mins = int((start - now).total_seconds() // 60)
        if start <= now:
            when = f"now, until {end:%H:%M}"
        elif start.date() == now.date():
            when = f"in {mins} min" if mins < 60 else f"in {mins // 60} h {mins % 60:02d} min"
        else:
            when = _day_label(start.date(), now.date())
        out["next"] = {"subject": str(item.get("subject") or "Meeting"), "time": f"{start:%H:%M}", "when": when,
                       "soon": start.date() == now.date() and mins <= 10}
    day_end = now.replace(hour=end_hour, minute=0, second=0, microsecond=0)
    free, gap, cursor = 0, None, now
    for start, end, _item in spans + [(day_end, day_end, None)]:
        start, end = min(max(start, now), day_end), min(end, day_end)
        if end <= cursor and start < day_end:
            continue
        if start > cursor:
            mins = int((start - cursor).total_seconds() // 60)
            free += mins
            if gap is None and mins >= 30:
                gap = (cursor, start)
        cursor = max(cursor, end)
    out["free"] = {"minutes": max(0, free), "gap": f"{gap[0]:%H:%M} to {gap[1]:%H:%M}" if gap else "",
                   "now": bool(gap) and gap[0] == now}
    return out


def _check():
    day = datetime(2026, 11, 16, 13, 52)
    meets = [{"subject": "Backlog", "start": "2026-11-16 14:00", "end": "2026-11-16 14:30"},
             {"subject": "Sync", "start": "2026-11-16 16:00", "end": "2026-11-16 16:30"}]
    g = _glance_cal(meets, day)
    assert g["next"] == {"subject": "Backlog", "time": "14:00", "when": "in 8 min", "soon": True}, g
    assert g["free"] == {"minutes": 128, "gap": "14:30 to 16:00", "now": False}, g
    assert _glance_cal([], datetime(2026, 11, 16, 18, 0))["free"]["minutes"] == 0
    assert _glance_cal([], datetime(2026, 11, 16, 16, 0))["free"] == {"minutes": 60, "gap": "16:00 to 17:00", "now": True}
    assert match("whens my next meeting") == {"op": "next"}
    assert match("When's my next meeting?") == {"op": "next"}
    assert match("what's my next teams meeting") == {"op": "next"}
    assert match("open my next meeting") is None
    assert match("join the next meeting") is None
    assert match("what's on my calendar tomorrow") == {"op": "agenda", "when": "tomorrow"}
    assert match("meetings this week") == {"op": "agenda", "when": "week"}
    assert match("my calendar") == {"op": "agenda", "when": "today"}
    assert match("any unread email") == {"op": "mail"}
    assert match("Summarise my unread mail") == {"op": "mail"}
    assert match("emails from Sam") == {"op": "mail", "query": "sam"}
    assert match("my tasks") == {"op": "tasks"}
    assert match("what files was i working on") == {"op": "files"}
    assert match("who is my manager") == {"op": "me", "focus": "manager"}
    assert match("what's my email") == {"op": "me", "focus": "email"}
    assert match("whens my next meeting and open it") is None
    cal = {"op": "next"}
    assert match("what about tomorrow", cal) == {"op": "agenda", "when": "tomorrow"}
    assert match("tomorrow", cal) == {"op": "agenda", "when": "tomorrow"}
    assert match("same for this week", cal) == {"op": "agenda", "when": "week"}
    assert match("tomorrow") is None
    assert match("what about tomorrow", {"op": "mail"}) is None
    assert match("what about my email") == {"op": "me", "focus": "email"}
    assert match("and my tasks") == {"op": "tasks"}
    assert match("open it", cal) is None
    assert match("recent recordings") == {"op": "recordings", "days": 14}
    assert match("transcribe this paragraph into french") is None
    nabeel = match("can you check the latest meeting transcrips i had with nabeel for mandate review requirements")
    assert nabeel["op"] == "recordings" and "nabeel" in nabeel["query"] and "mandate" in nabeel["query"], nabeel
    assert _booking({"subject": "Workstation WS-2B089 Reserved - Full Day", "organizer": "EWRS-noreply@ontario.ca"})
    assert not _booking({"subject": "Daily EDCP Scrum", "organizer": "Ada"})
    assert match("transcript of the weekly sync") == {"op": "recordings", "days": 7, "query": "weekly sync"}
    follow = "what about the recent ones just few days ago"
    assert match(follow) is None
    assert match(follow, prior="I can't determine the most recent recorded meeting") == {"op": "recordings", "days": 14}
    shot = "Checking your next meeting.\n- Looking up upcoming calendar events and current schedule"
    assert status_only(shot)
    assert not status_only("Your next meeting is **Budget** today at 15:30 (in 30 min).")
    assert not status_only("Disk C has 42 GB free.")
    now = datetime(2026, 11, 16, 15, 0)
    items = [{"subject": "Budget", "start": "2026-11-16 15:30", "end": "2026-11-16 16:00",
              "location": "Teams", "organizer": "Ada Lovelace", "online": True, "allday": False}]
    text = present({"op": "next"}, {"items": items}, now)
    assert "Your next meeting is **Budget** today at 15:30 (in 30 min)." in text, text
    assert "- Teams" in text and "- With Ada Lovelace" in text
    current = [{"subject": "Standup", "start": "2026-11-16 14:30", "end": "2026-11-16 15:30",
                "location": "", "organizer": "", "online": True, "allday": False}]
    assert present({"op": "next"}, {"items": current}, now).startswith("You're in **Standup** until 15:30.")
    assert present({"op": "next"}, {"items": []}, now) == "You have no more meetings in the next 14 days."
    later = [{"subject": "Review", "start": "2026-11-17 09:00", "end": "2026-11-17 09:30",
              "location": "", "organizer": "", "online": False, "allday": False}]
    assert "tomorrow at 09:00" in present({"op": "next"}, {"items": later}, now)
    agenda = present({"op": "agenda", "when": "today"}, {"items": items + [{
        "subject": "Holiday", "start": "2026-11-16 00:00", "end": "2026-11-17 00:00", "allday": True}]}, now)
    assert agenda.startswith("You have **2** meetings today.") and "- All day **Holiday**" in agenda
    desk = {"subject": "Workstation WS-2B034 Reserved - Full Day", "start": "2026-11-16 08:00", "end": "2026-11-16 08:00"}
    gone = dict(desk, subject="[CANCELLED] Workstation WS-2B089 Reserved - Full Day")
    scrum = {"subject": "Daily EDCP Scrum", "start": "2026-11-16 09:30", "end": "2026-11-16 10:00"}
    day = present({"op": "agenda", "when": "today"}, {"items": [gone, desk, desk, scrum]}, now)
    assert day.startswith("You have **1** meeting today.\nDesk booked: Workstation WS-2B034.") and "CANCELLED" not in day, day
    assert _cancelled({"subject": "Canceled: Sync"}) and not _cancelled({"subject": "Cancellation policy review"})
    assert present({"op": "mail"}, {"unread": 0, "items": []}) == "You have no unread email."
    mail = present({"op": "mail"}, {"unread": 4, "items": [{"from": "Ada", "subject": "Notes", "received": "11-16 14:02"}]})
    assert mail.startswith("You have **4** unread emails.") and "**Ada**: Notes" in mail
    assert "No recent inbox mail matches **sam**." in present({"op": "mail", "query": "sam"}, {"items": [], "query": "sam"})
    assert _ps_quote("a'; Remove-Item") == "'a''; Remove-Item'"
    assert _load("noise\n{\"error\": \"outlook\"}")["error"] == "outlook"
    fri = {"op": "event", "subject": "Fill timesheet; WIN", "start": "2026-11-20 12:00", "minutes": 15, "reminder": 10,
           "repeat": "weekly", "days": ["friday"], "free": True}
    ics = _ics(fri, datetime(2026, 11, 19, 19, 0))
    assert "RRULE:FREQ=WEEKLY;BYDAY=FR\r\n" in ics and "SUMMARY:Fill timesheet\\; WIN" in ics and "TRIGGER:-PT10M" in ics
    assert "DTSTART:20261120T120000" in ics and "DTEND:20261120T121500" in ics and "TRANSP:TRANSPARENT" in ics
    ps = _event_ps(dict(fri, subject="Tim's sheet"))
    assert "$a.Subject = 'Tim''s sheet'" in ps and "DayOfWeekMask = 32" in ps and "RecurrenceType = 1" in ps
    assert "if ($true)" in ps and "if ('')" in ps
    assert "RecurrenceType = 0" in _event_ps(dict(fri, repeat="daily")) and "(-1 -ge 0)" in _event_ps(dict(fri, repeat="none"))
    assert "DayOfWeekMask = 62" in _event_ps(dict(fri, repeat="weekdays", days=None))
    body = _event_body(dict(fri, until="2027-06-30"), "Eastern Standard Time")
    assert body["Recurrence"] == {"Pattern": {"Type": "Weekly", "Interval": 1, "DaysOfWeek": ["Friday"], "FirstDayOfWeek": "Sunday"},
                                  "Range": {"Type": "EndDate", "StartDate": "2026-11-20", "EndDate": "2027-06-30"}}
    assert body["ShowAs"] == "Free" and body["End"]["DateTime"] == "2026-11-20T12:15:00"
    assert "Recurrence" not in _event_body(dict(fri, repeat="none"), "UTC")
    assert outlook_order({"outlook.exe"}, True) == ["classic", "web"]
    assert outlook_order({"olk.exe"}, True) == ["web", "classic"]
    assert outlook_order(set(), False) == ["web"] and outlook_order(set(), True) == ["classic", "web"]
    assert _when_text(fri).startswith("- Every Friday at 12:00, from November 20")
    assert _on_outlook("https://outlook.cloud.microsoft/calendar/") and _on_outlook("https://outlook.office.com/mail/")
    assert not _on_outlook("https://login.microsoftonline.com/x") and not _on_outlook("https://evil-outlook.office.com.example/")
    print("account ok")


if __name__ == "__main__":
    _check()
