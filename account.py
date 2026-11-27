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
from pathlib import Path
import html as html_lib
from urllib.parse import parse_qs, quote, unquote, urlparse

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
$light = __LIGHT__
function Hidden($m) {
  try { if ($m.MeetingStatus -eq 5) { return $true } } catch {}
  try { if ($m.ResponseStatus -eq 4) { return $true } } catch {}
  return $false
}
function MailBody($m) {
  # ponytail: first 6000 characters of plain text. A longer quoted thread is cut here; the whole mailbox is not.
  $text = ''
  try { $text = [string]$m.Body } catch {}
  $text = ($text -replace '(\r?\n){3,}', "`n`n").Trim()
  if ($text.Length -gt 6000) { $text = $text.Substring(0, 6000).TrimEnd() + '...' }
  return $text
}
function MailLinks($m) {
  # The recording card is an HTML link. Plain Body drops the address and often the title.
  $html = ''
  try { $html = [string]$m.HTMLBody } catch {}
  $found = New-Object System.Collections.Generic.List[object]
  if (-not $html) { return @() }
  $hits = [regex]::Matches($html, '(?is)<a\b[^>]*href\s*=\s*"([^"]+)"[^>]*>(.*?)</a>')
  foreach ($hit in $hits) {
    if ($found.Count -ge 12) { break }
    $url = [System.Net.WebUtility]::HtmlDecode($hit.Groups[1].Value).Trim()
    if ($url -notmatch '^https?://') { continue }
    $label = [regex]::Replace($hit.Groups[2].Value, '(?is)<[^>]+>', ' ')
    $label = [System.Net.WebUtility]::HtmlDecode(($label -replace '\s+', ' ').Trim())
    if (-not $label) {
      $alt = [regex]::Match($hit.Groups[2].Value, '(?i)alt\s*=\s*"([^"]+)"')
      if ($alt.Success) { $label = [System.Net.WebUtility]::HtmlDecode($alt.Groups[1].Value).Trim() }
    }
    if ($label.Length -gt 120) { $label = $label.Substring(0, 120).TrimEnd() }
    $found.Add([pscustomobject]@{label=$label; url=$url})
  }
  return @($found.ToArray())
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
        $items.Add([pscustomobject]@{kind='flagged'; title=[string]$m.Subject; who=[string]$m.SenderName; due=''; id=[string]$m.EntryID})
        if ($items.Count -ge 8) { break }
      }
    } catch {}
    try {
      foreach ($t in $ns.GetDefaultFolder(13).Items) {
        if ($items.Count -ge 16) { break }
        try { if ($t.Complete) { continue } } catch {}
        $due = ''
        try { if ($t.DueDate -and $t.DueDate.Year -gt 2000 -and $t.DueDate.Year -lt 2100) { $due = $t.DueDate.ToString('MM-dd') } } catch {}
        $items.Add([pscustomobject]@{kind='task'; title=[string]$t.Subject; who=''; due=$due; id=[string]$t.EntryID})
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
          $items.Add([pscustomobject]@{from=$from; subject=[string]$m.Subject; received=$m.ReceivedTime.ToString('MM-dd HH:mm'); body=(MailBody $m); id=[string]$m.EntryID; links=(MailLinks $m)})
          if ($items.Count -ge 8) { break }
        }
      }
      Out-Json @{items=@($items.ToArray()); query=$q}
      return
    }
    $unread = $all.Restrict('[UnRead] = True')
    if ($light) {
      $from = ''
      try { if ($unread.Count -gt 0) { $from = [string]$unread.Item(1).SenderName } } catch {}
      Out-Json @{unread=[int]$inbox.UnReadItemCount; from=$from; items=@()}
      return
    }
    for ($i = 1; $i -le [math]::Min(8, $unread.Count); $i++) {
      $m = $unread.Item($i)
      $items.Add([pscustomobject]@{from=[string]$m.SenderName; subject=[string]$m.Subject; received=$m.ReceivedTime.ToString('MM-dd HH:mm'); body=(MailBody $m); id=[string]$m.EntryID; links=(MailLinks $m)})
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
      location=[string]$m.Location; organizer=[string]$m.Organizer; online=$online; allday=$allday; tentative=$tentative;
      id=[string]$m.EntryID
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
    const fields = "Subject,From,ReceivedDateTime,BodyPreview,Body,WebLink";
    const url = query
      ? "/api/v2.0/me/mailfolders/inbox/messages?$top=20&$select=" + fields + "&$orderby=ReceivedDateTime desc"
      : "/api/v2.0/me/mailfolders/inbox/messages?$top=8&$filter=IsRead eq false&$select=" + fields + "&$orderby=ReceivedDateTime desc";
    const response = await fetch(url, {credentials: "include", headers});
    if (!response.ok) return {error: "web", detail: String(response.status)};
    const body = await response.json();
    const linksOf = (raw) => {
      const links = [];
      const re = /<a\b[^>]*href\s*=\s*["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
      let hit;
      while ((hit = re.exec(raw)) && links.length < 12) {
        let label = hit[2].replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
        if (!label) {
          const alt = /alt\s*=\s*["']([^"']+)["']/i.exec(hit[2]);
          label = alt ? alt[1] : "";
        }
        const href = hit[1].replace(/&amp;/g, "&").trim();
        if (/^https?:\/\//i.test(href)) links.push({label: label.slice(0, 120), url: href});
      }
      return links;
    };
    const textOf = (m) => {
      const raw = (m.Body && m.Body.Content) || m.BodyPreview || "";
      const text = raw.replace(/<style[\s\S]*?<\/style>/gi, " ").replace(/<script[\s\S]*?<\/script>/gi, " ")
        .replace(/<br\s*\/?>/gi, "\n").replace(/<\/p>/gi, "\n").replace(/<[^>]+>/g, " ")
        .replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
        .replace(/&#39;|&apos;/g, "'").replace(/&quot;/g, '"')
        .replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").replace(/[ \t]{2,}/g, " ").trim();
      return text.length > 6000 ? text.slice(0, 6000).trimEnd() + "..." : text;
    };
    const needle = (query || "").toLowerCase();
    const items = [];
    for (const m of body.value || []) {
      const from = (m.From && m.From.EmailAddress && m.From.EmailAddress.Name) || "";
      const subject = m.Subject || "";
      if (needle && !from.toLowerCase().includes(needle) && !subject.toLowerCase().includes(needle)) continue;
      items.push({from, subject, received: stamp({DateTime: m.ReceivedDateTime, TimeZone: "UTC"}).slice(5), body: textOf(m), link: m.WebLink || "", links: linksOf((m.Body && m.Body.Content) || "")});
      if (items.length >= 8) break;
    }
    const folder = await fetch("/api/v2.0/me/mailfolders/inbox?$select=UnreadItemCount", {credentials: "include", headers});
    let unread = items.length;
    if (folder.ok) {
      const info = await folder.json();
      if (typeof info.UnreadItemCount === "number") unread = info.UnreadItemCount;
    }
    return query ? {items, query} : {items, unread};
  }
  if (op === "tasks") {
    const response = await fetch("/api/v2.0/me/messages?$top=8&$filter=Flag/FlagStatus eq 'Flagged'&$select=Subject,From", {credentials: "include", headers});
    if (!response.ok) return {error: "web", detail: String(response.status)};
    const body = await response.json();
    return {items: (body.value || []).map(m => ({
      kind: "flagged", title: m.Subject || "(no subject)",
      who: (m.From && m.From.EmailAddress && m.From.EmailAddress.Name) || "", due: ""
    }))};
  }
  if (op === "me") {
    const response = await fetch("/api/v2.0/me?$select=DisplayName,EmailAddress,JobTitle,Department", {credentials: "include", headers});
    if (!response.ok) return {error: "web", detail: String(response.status)};
    const me = await response.json();
    let manager = "";
    try {
      const mgr = await fetch("/api/v2.0/me/manager?$select=DisplayName", {credentials: "include", headers});
      if (mgr.ok) manager = ((await mgr.json()).DisplayName) || "";
    } catch (e) {}
    return {name: me.DisplayName || "", email: me.EmailAddress || "", title: me.JobTitle || "", department: me.Department || "", manager};
  }
  if (op !== "next" && op !== "agenda") return {error: "web"};
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
    + "&$select=Subject,Start,End,Location,Organizer,IsAllDay,IsOnlineMeeting,ResponseStatus,WebLink&$orderby=Start/DateTime&$top=20";
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
      link: m.WebLink || "",
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
        found = {"op": "mail"}
        if re.search(r"summari[sz]", text):
            found["summarize"] = True
        return found
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
    """User-facing answer for one account op. `chat` enables the other Outlook, which is Outlook on the web."""
    op = action["op"]
    if op == "recordings":
        return meetings.answer(action, chat)
    if op == "files":
        return present(action, _load(_files_raw()))
    if op == "me" and action.get("focus") == "email":
        return present(action, {"email": _upn()})
    if op == "event":
        return event(action, chat)
    return present(action, _read(action, chat))


def _read(action, chat=None):
    """Outlook payload for one op. `chat` is how the other Outlook app is read."""
    op = action["op"]
    web_ops = ("next", "agenda", "mail", "tasks", "me")
    order = outlook_order(_running(), True)
    if chat is not None and op in web_ops and order[0] == "web":
        web = _web(chat, action)
        if web:
            return web
    data = _load(_outlook(action))
    if data.get("error") == "outlook" and chat is not None and op in web_ops:
        web = _web(chat, action)
        if web:
            data = web
        else:
            data["web"] = "tried"
    if op == "me" and data.get("error") and action.get("focus") in (None, "", "email"):
        email = _upn()
        if email:
            data = {"email": email}
    return data


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
        detail = str(data.get("detail") or "")[:200]
        if detail and detail != "bad result":
            return f"I couldn't read that from your account. {detail}"
        return "I couldn't read that from your account."
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
        link = _item_link(item)
        tip = _mail_tip(item) if link.startswith("mailitem:") else ""
        lines.append(f"- {item.get('received', '')} **{item.get('from') or 'Unknown'}**: {_linked(item.get('subject') or '(no subject)', link, tip)}")
        if not action.get("summarize"):
            body = str(item.get("body") or "").strip()
            if body:
                lines.append(body)
            lines.extend(_cite_lines(item))
    return "\n".join(lines)


_SUMMARY = (
    "For each message, in order, write exactly two lines and nothing else:\n"
    "N. <sender> — <date>\n"
    "<one or two sentences. Say what it is and if the reader should act. "
    "Do not write a web address, a mailitem id, or the words Open this message.>\n"
    "After all of the messages, only when something needs a reply, a decision, or a deadline, add:\n"
    "## Important\n"
    "- <one short line: what to do, and why. Name the sender. Do not write a web address.>\n"
    "Leave Important out when nothing qualifies."
)


def summarize(action, chat, on_text=None):
    """One or two sentences per unread message. Links are added here so they stay clickable."""
    data = _read(action, chat)
    if data.get("error") or not _items(data):
        return present(action, data)
    if data.get("unread") == 0:
        return present(action, data)
    brief = _mail_brief(data)

    def emit(partial):
        if not on_text:
            return
        shown = "\n".join(line for line in partial.splitlines() if "mailitem:" not in line.lower())
        if shown.strip():
            on_text(shown)

    reply = chat.send(_SUMMARY + "\n\n" + brief, on_text=emit)
    text = re.sub(r"^```[a-z]*\n?|```$", "", (reply.text or "").strip()).strip()
    if len(text) < 40 or text.lstrip()[:1] in "{[":
        text = present({**action, "summarize": True}, data)
    else:
        text = _shape_summary(text, data)
    return text


def _mail_brief(data):
    items = _items(data)
    count = data.get("unread")
    if count is None:
        count = len(items)
    lines = [f"Unread count: {count}. The newest {len(items)} follow."]
    for number, item in enumerate(items, 1):
        lines.append(f"\n{number}. {item.get('received', '')} | {item.get('from') or 'Unknown'} | {item.get('subject') or '(no subject)'}")
        excerpt = _excerpt(item.get("body"))
        if excerpt:
            lines.append(excerpt)
        names = [label for label, _url in list(_mail_urls(item))[:3] if label and not label.lower().startswith("http")]
        if names:
            lines.append("Link names: " + "; ".join(names))
    return "\n".join(lines)


def _excerpt(body):
    text = _BARE_URL.sub(" ", str(body or ""))
    text = " ".join(text.split())
    return text[:280]


def _two_sentences(text):
    bits = re.split(r"(?<=[.!?])\s+", " ".join(str(text or "").split()))
    return " ".join(bit for bit in bits[:2] if bit)[:360]


def _nice_label(label, limit=64):
    """A name fit to show. A host or a raw address stays off the line; the hover still has the full address."""
    label = re.sub(r"\s+", " ", str(label or "")).strip(" .")
    if not label or re.search(r"https?://|safelinks|protection\.outlook", label, re.I):
        return ""
    if " " not in label and re.search(r"\w+\.\w{2,}", label):
        return ""
    if len(label) <= limit:
        return label
    cut = label[:limit].rsplit(" ", 1)[0]
    return (cut or label[:limit]).rstrip(" -–—") + "…"


def _shape_summary(text, data):
    """One block per message: who and when, the subject that opens it, one or two sentences, then named links."""
    items = _items(data)
    body, _, important = text.partition("\n##")
    if re.match(r"\s*important\b", important, re.I):
        important = re.sub(r"^\s*important\b[:\s]*", "", important, count=1, flags=re.I).strip()
    else:
        important = ""
    pieces = [piece.strip() for piece in re.split(r"\n(?=\d+[.)]\s)", body) if piece.strip()]
    if len(pieces) < 2:
        pieces = [block.strip() for block in re.split(r"\n\s*\n", body) if block.strip()]
    sentences = []
    for block in pieces:
        lines = [line.strip() for line in block.splitlines()
                 if line.strip() and "mailitem:" not in line.lower() and not line.lower().startswith("open this message")]
        if not lines:
            continue
        if re.match(r"\d+[.)]\s", lines[0]):
            lines[0] = re.sub(r"^\d+[.)]\s*", "", lines[0])
        if len(lines) > 1 and re.search(r"—| - ", lines[0]) and len(lines[0]) < 90:
            lines = lines[1:]
        sentence = _two_sentences(" ".join(lines))
        if sentence:
            sentences.append(sentence)
    count = data.get("unread")
    if count is None:
        count = len(items)
    if count and count != len(items):
        lines = [f"You have **{count}** unread emails. These are the newest {len(items)}."]
    else:
        lines = [f"You have **{count}** unread email{'' if count == 1 else 's'}."]
    for index, item in enumerate(items):
        who = item.get("from") or "Unknown"
        when = item.get("received") or ""
        url = _item_link(item)
        tip = _mail_tip(item) if url.startswith("mailitem:") else ""
        subject = _nice_label(item.get("subject"), 72) or "(no subject)"
        lines.append("")
        lines.append(f"## {who}" + (f" · {when}" if when else ""))
        lines.append(_linked(subject, url, tip))
        if index < len(sentences):
            lines.append(sentences[index])
        named = []
        for label, link in _mail_urls(item):
            show = _nice_label(label, 48)
            if show and link not in url:
                named.append((show, link))
            if len(named) == 2:
                break
        lines.extend(f"- {_linked(show, link)}" for show, link in named)
    notes = []
    for line in important.splitlines():
        note = _BARE_URL.sub("", line).strip().lstrip("-").strip()
        note = re.sub(r"mailitem:\S+", "", note).strip()
        if not note:
            continue
        for item in items:
            who = str(item.get("from") or "")
            token = who.split("(")[0].split(",")[0].strip()
            url = _item_link(item)
            if token and len(token) > 2 and token.lower() in note.lower() and url and url not in note:
                tip = _mail_tip(item) if url.startswith("mailitem:") else ""
                subject = _nice_label(item.get("subject"), 48) or "Open"
                note += " " + _linked(subject, url, tip)
                break
        notes.append(f"- {note}")
    if notes:
        lines.extend(["", "## Important", *notes])
    return "\n".join(lines).strip()


def _named_links(text, data):
    """A bare address becomes a named link. The address stays for hover, not as the words."""
    pairs = [pair for item in _items(data) for pair in _mail_urls(item)]
    for label, url in sorted(pairs, key=lambda pair: len(pair[1]), reverse=True):
        if url not in text or f"]({url})" in text or f"]({url} " in text:
            continue
        name = label if label and not label.lower().startswith("http") else "Link"
        wrapped = f"[{name}]({url})"
        held = text.replace(f"<{url}>", "\ue000")
        text = re.sub(r"(?<!\]\()" + re.escape(url), wrapped, held).replace("\ue000", wrapped)
    return text


def _ensure_open(text, data):
    extra = []
    for item in _items(data):
        url = _item_link(item)
        if not url or url in text:
            continue
        tip = _mail_tip(item) if url.startswith("mailitem:") else ""
        extra.append("- " + _linked(item.get("subject") or "Open", url, tip))
    if not extra:
        return text
    return text.rstrip() + "\n\nOpen:\n" + "\n".join(extra[:8])


def _present_tasks(items):
    if not items:
        return "You have no flagged mail and no open Outlook tasks."
    lines = [f"You have **{len(items)}** open item{'' if len(items) == 1 else 's'}."]
    for item in items:
        due = f" (due {item['due']})" if item.get("due") else ""
        who = f" — {item['who']}" if item.get("who") else ""
        kind = "Flagged" if item.get("kind") == "flagged" else "Task"
        link = _item_link(item)
        tip = _mail_tip({"from": item.get("who"), "received": item.get("received")}) if item.get("kind") == "flagged" and link.startswith("mailitem:") else ""
        title = _linked(item.get("title") or "(no subject)", link, tip)
        lines.append(f"- {kind}: **{title}**{who}{due}")
    return "\n".join(lines)


def _present_files(items):
    if not items:
        return "No recently opened OneDrive files."
    lines = ["Recent files in your OneDrive:"]
    for item in items:
        lines.append(f"- {item.get('when', '')} **{_linked(item.get('name') or 'file', _item_link(item))}**")
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
    subject = item.get("subject") or "(no subject)"
    link = _item_link(item)
    return _linked(subject, link) if link else subject


def _item_link(item):
    """https WebLink, mailitem: Outlook id, or file: path. Empty when there is nothing to open."""
    link = str(item.get("link") or "").strip()
    if link.startswith(("https://", "http://", "mailitem:", "file:")):
        return link.replace(")", "%29")
    item_id = str(item.get("id") or "").strip()
    if item_id:
        return "mailitem:" + quote(item_id, safe="")
    path = str(item.get("path") or "").strip()
    if not path:
        return ""
    target = Path(os.path.expandvars(os.path.expanduser(path)))
    if not target.is_absolute():
        return ""
    try:
        return target.as_uri().replace(")", "%29")
    except ValueError:
        return ""


def _mail_tip(item):
    """Hover text for a mail link: Mail (date time) - sender. The raw id stays the address."""
    when = str(item.get("received") or "").strip()
    who = re.sub(r'[\r\n"]', "", str(item.get("from") or "").strip())
    if when and who:
        return f"Mail ({when}) - {who}"
    if when:
        return f"Mail ({when})"
    if who:
        return f"Mail - {who}"
    return ""


def _linked(label, url, tip=""):
    """Markdown link, or the plain label when the address would break the link syntax."""
    label = re.sub(r"[\[\]*`]", "", str(label or "").strip()) or "Open"
    if not url or re.search(r"[\s)]", url):
        return label
    tip = re.sub(r'[\r\n"]', "", str(tip or "").strip())
    if tip:
        return f'[{label}]({url} "{tip}")'
    return f"[{label}]({url})"


def open_mail(entry_id):
    """Open one Outlook item by EntryID. Ids are hex or a quoted REST id; anything else is ignored."""
    entry_id = str(entry_id or "").strip()
    if not re.fullmatch(r"[A-Za-z0-9+/=_-]{8,500}", entry_id):
        raise ValueError("That mail link is not an Outlook item.")
    script = (
        "$ErrorActionPreference = 'Stop'\n"
        f"$id = '{entry_id}'\n"
        "try { $o = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }\n"
        "catch { $o = New-Object -ComObject Outlook.Application }\n"
        "$o.GetNamespace('MAPI').GetItemFromID($id).Display()\n"
    )
    sysinfo._powershell("open mail", script)


_BARE_URL = re.compile(r"https?://[^\s<>'\"]+")
_REC_HOST = re.compile(r"(?:^|\.)(?:stream\.microsoft\.com|microsoftstream\.com|clipchamp\.com)$", re.I)
_REC_HINT = re.compile(r"record|clipchamp|/video|stream\.aspx|:v:|/embed", re.I)


def _unwrap(url):
    """Outlook wraps links in safelinks. The real address is the url= query."""
    url = html_lib.unescape(str(url or "").strip()).rstrip(".,)")
    parts = urlparse(url)
    if parts.netloc.lower().endswith("safelinks.protection.outlook.com"):
        query = parse_qs(parts.query)
        inner = (query.get("url") or query.get("URL") or [""])[0]
        if inner:
            return unquote(inner).rstrip(".,)")
    return url


def is_recording(url):
    """A video or meeting recording, not every SharePoint file."""
    parts = urlparse(_unwrap(url))
    if parts.scheme not in ("http", "https"):
        return False
    host = parts.netloc.lower()
    if _REC_HOST.search(host):
        return True
    blob = host + parts.path
    if "sharepoint.com" in host or host.endswith("teams.microsoft.com") or host.endswith("onedrive.live.com"):
        return bool(_REC_HINT.search(blob))
    return False


def _link_label(url, label):
    label = re.sub(r"\s+", " ", str(label or "")).strip()
    if label and label.lower() not in ("here", "link", "click", "click here", "open"):
        return label[:120]
    name = unquote(urlparse(url).path).rstrip("/").split("/")[-1]
    return (name or urlparse(url).netloc)[:120]


def _mail_urls(item):
    """(label, url) from the HTML card and from addresses written in the text."""
    own = {_unwrap(item.get("link") or "")}
    seen = set()
    rows = item.get("links") or []
    if isinstance(rows, dict):
        rows = [rows]
    pairs = [(row.get("label"), row.get("url")) for row in rows if isinstance(row, dict)]
    pairs += [(None, match) for match in _BARE_URL.findall(str(item.get("body") or ""))]
    for label, url in pairs:
        url = _unwrap(url)
        if not url.startswith(("https://", "http://")) or url in seen or url in own:
            continue
        seen.add(url)
        yield _link_label(url, label), url


def _cite_lines(item):
    """Recording and other links, with who sent the mail and when, so the answer can cite them."""
    who = item.get("from") or "the sender"
    when = f", received {item['received']}" if item.get("received") else ""
    recordings, others = [], []
    for label, url in _mail_urls(item):
        if is_recording(url) and len(recordings) < 4:
            recordings.append(f"Recording: {_linked(label, url)} — {who}{when}")
        elif not is_recording(url) and len(others) < 4:
            others.append(f"Link: {_linked(label, url)}")
    return recordings + others


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
    script = script.replace("__LIGHT__", "$true" if action.get("light") else "$false")
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
        return {"error": "failed", "detail": "bad result"}
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


def _glance_mail(mail):
    """Unread count and the newest sender, or None when this payload has no mail."""
    if not isinstance(mail, dict) or "unread" not in mail:
        return None
    latest = _items(mail)[:1]
    who = str(mail.get("from") or (latest[0].get("from") if latest else "") or "")
    return {"unread": int(mail["unread"]), "from": who}


def glance(now=None, end_hour=17):
    """Start-screen facts: next meeting, unread mail, free time left today. Empty parts when Outlook is closed."""
    now = now or datetime.now()
    cal = _load(_outlook({"op": "next"}))
    mail = _load(_outlook({"op": "mail", "light": True}))
    out = {"outlook": cal.get("error") != "outlook"}
    items = [i for i in _items(cal) if not _booking(i) and not _cancelled(i)]
    out.update(_glance_cal(items, now, end_hour))
    viewed = _glance_mail(mail)
    if viewed:
        out["mail"] = viewed
    # Classic Outlook answers directly. The other Outlook app is read on the web by whoever holds the browser.
    out["need_web"] = mail.get("error") == "outlook" or cal.get("error") == "outlook"
    return out


def glance_fill(out, extra, now=None, end_hour=17):
    """Fill whatever classic Outlook missed, from the other Outlook app."""
    now = now or datetime.now()
    extra = extra or {}
    viewed = _glance_mail(extra.get("mail") or {})
    if viewed:
        out["mail"] = viewed
    cal = extra.get("cal") or {}
    if not out.get("outlook") and isinstance(cal, dict) and cal.get("items") is not None and not cal.get("error"):
        items = [i for i in _items(cal) if not _booking(i) and not _cancelled(i)]
        out.update(_glance_cal(items, now, end_hour))
        out["outlook"] = True
    return out


def web_glance(chat):
    """Mail count and the next meetings through the signed-in Outlook on the web."""
    return {"mail": _web(chat, {"op": "mail"}) or {}, "cal": _web(chat, {"op": "next"}) or {}}


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
    assert _glance_mail({"unread": 3, "from": "Ada"}) == {"unread": 3, "from": "Ada"}
    assert _glance_mail({"error": "outlook"}) is None
    filled = glance_fill({"outlook": False}, {"mail": {"unread": 2, "items": [{"from": "Sam"}]}, "cal": {"items": []}},
                         now=datetime(2026, 11, 16, 15, 0))
    assert filled["mail"] == {"unread": 2, "from": "Sam"} and filled["outlook"] is True, filled
    assert match("whens my next meeting") == {"op": "next"}
    assert match("When's my next meeting?") == {"op": "next"}
    assert match("what's my next teams meeting") == {"op": "next"}
    assert match("open my next meeting") is None
    assert match("join the next meeting") is None
    assert match("what's on my calendar tomorrow") == {"op": "agenda", "when": "tomorrow"}
    assert match("meetings this week") == {"op": "agenda", "when": "week"}
    assert match("my calendar") == {"op": "agenda", "when": "today"}
    assert match("any unread email") == {"op": "mail"}
    assert match("Summarise my unread mail") == {"op": "mail", "summarize": True}
    sample = {"items": [{"from": "Ada", "subject": "Hi", "received": "11-26 12:19", "id": "ABC",
                         "body": "See https://example.com/a please",
                         "links": [{"label": "Agenda", "url": "https://example.com/a"}]}]}
    assert _named_links("See <https://example.com/a> today", sample) == "See [Agenda](https://example.com/a) today"
    assert "https://example.com/a" not in _excerpt(sample["items"][0]["body"])
    opened = _ensure_open("A short summary of the mail.", sample)
    assert "Open:" in opened and "mailitem:ABC" in opened
    assert _nice_label("can01.safelinks.protection.outlook.com") == ""
    assert _nice_label("Manage my booking page") == "Manage my booking page"
    assert "ontario.ca" in _nice_label("Invitation for sohel.shekh@ontario.ca tomorrow")
    shaped = _shape_summary(
        '1. Ada — 11-26 12:19\nYour booking page is ready. Share it when you want.\n\n'
        'Open this message: mailitem:ABC "Mail (11-26 12:19 - Ada)"\n\n'
        '## Important\n- Reply to Ada about the page, because people are waiting.',
        {"unread": 3, "items": sample["items"]})
    assert "](mailitem:ABC" in shaped and "Open this message" not in shaped, shaped
    assert "## Important" in shaped and "booking page is ready" in shaped
    assert _two_sentences("One. Two. Three.") == "One. Two."
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
    mail = present({"op": "mail"}, {"unread": 4, "items": [
        {"from": "Ada", "subject": "Notes", "received": "11-16 14:02", "body": "Please review the draft by Friday."}]})
    assert mail.startswith("You have **4** unread emails.") and "**Ada**: Notes" in mail
    assert "Please review the draft by Friday." in mail
    opened = present({"op": "mail"}, {"unread": 1, "items": [
        {"from": "Ada", "subject": "Notes", "received": "11-16 14:02", "id": "ABCDEF0123456789", "body": "Hi"}]})
    assert '[Notes](mailitem:ABCDEF0123456789 "Mail (11-16 14:02) - Ada")' in opened, opened
    web = present({"op": "mail"}, {"unread": 1, "items": [
        {"from": "Ada", "subject": "Notes", "link": "https://outlook.office.com/mail/id/1"}]})
    assert "[Notes](https://outlook.office.com/mail/id/1)" in web
    clip = "https://www.clipchamp.com/watch/abc"
    safe = "https://nam.safelinks.protection.outlook.com/?url=" + quote(clip)
    doc = "https://contoso.sharepoint.com/sites/a/Shared%20Documents/plan.docx"
    video = "https://contoso.sharepoint.com/sites/a/video/SCO.mp4"
    uat = "https://stage.supplychaindata.mgcs.gov.on.ca/en-CA/"
    shown = present({"op": "mail"}, {"unread": 1, "items": [{
        "from": "Chen, Yi-Leng", "subject": "SCO recordings", "received": "03-04 14:10",
        "id": "ABCDEF0123456789", "body": "UAT " + uat,
        "links": [{"label": "SCO Demo", "url": safe}, {"label": "plan", "url": doc}, {"label": "", "url": video}]}]})
    assert f"Recording: [SCO Demo]({clip}) — Chen, Yi-Leng, received 03-04 14:10" in shown, shown
    assert f"Recording: [SCO.mp4]({video})" in shown and f"Link: [plan]({doc})" in shown, shown
    assert f"Link: [en-CA]({uat})" in shown, shown
    assert is_recording(safe) and is_recording(video) and not is_recording(doc) and not is_recording(uat)
    files = present({"op": "files"}, {"items": [{"name": "Plan.docx", "path": r"C:\Users\me\Plan.docx", "when": "11-16"}]})
    assert "file:///C:/Users/me/Plan.docx" in files, files
    meet = present({"op": "next"}, {"items": [dict(items[0], id="ABCDEF0123456789")]}, now)
    assert "[Budget](mailitem:ABCDEF0123456789)" in meet, meet
    try:
        open_mail("nope")
        raise AssertionError("short id should be rejected")
    except ValueError:
        pass
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
