"""InsideOPS (intra.ontario.ca) questions, answered from the site's own JSON API with the signed-in Copilot browser profile.

How the site works (mapped from the live site, Nov 2026):
- Every URL returns the same empty app shell; the app fills it from JSON, so pages are never rendered here.
- Search:   /api/search?q=...  (Elasticsearch; matches ALL words, so send keywords, not the question)
- Post:     /YYYY/MM/DD/<slug> -> /api/posts/slug/<slug>             -> [0].content.rendered
- Page:     /<section>/<slug>  -> /api/pages/slug/<slug>             -> [i].id -> /wp-json/wp/v2/pages/content/<id>
- Section:  /<section>         -> /api/section/slug/<section>        -> section_admin.section_home_page -> same content call
- /wp-json/wp/v2/search ignores the query; don't use it.
Auth is the site cookie in the Edge profile; nothing here touches the screen.
"""

import json
import re
from datetime import datetime
from html import unescape
from urllib.parse import quote, urlparse

import protocol

HOME = "https://intra.ontario.ca"
_SEARCH = ("/api/search?q={q}&date_type=all&date_month=01&date_range=1w&date_year={y}"
           "&page=1&per_page=5&sort=relevancy_desc&text_exclude=&type=all&within=")
_PAGES = 3
_PAGE_CHARS = 6000

# Topic map: what people ask -> where InsideOPS keeps it. query wins over path when both exist ({year} is filled in).
TOPICS = [
    (r"pay ?days?|pay ?dates?|pay calendar|when .*(get )?paid|(stat(utory)?|public|next|upcoming) holidays?|holiday (dates|list|calendar)|days? off",
     "{year} pay and holiday dates", None),
    (r"\bwin\b|pay ?stub|pay ?cheque|direct deposit|t4|tax slip|absence|timesheet|vacation (balance|request)", None, "/pay"),
    (r"benefits?|sun ?life|dental|health (plan|coverage)", "benefits", "/pay"),
    (r"flexible work|telework|hybrid work|work from home|wfh|office attendance", None, "/flexible-work"),
    (r"\beap\b|employee (and|&) family assistance|counsell?ing", None, "/ops/employee-family-and-assistance-program"),
    (r"learning hub|learnon|training|course", None, "/learning-and-development"),
    (r"health and wellness|wellness|mental health", None, "/cehsw"),
    (r"office space|workspace|desk booking|book a desk|ewrs", None, "/office-space"),
    (r"service centre|opssc|it help|onrequest", None, "/opssc"),
    (r"events? calendar|ops events", None, "/ops-events-calendar"),
    (r"\bifis\b|expense claim|travel expense|financial services", None, "/financial-services"),
    (r"cyber ?security|phishing", None, "/cyber"),
    (r"\bai hub\b|copilot (policy|guidance)|responsible ai", None, "/govtechon/aihub"),
    (r"executive announcements?", None, "/tbs/executive-announcement"),
    (r"collective agreements?|amapceo|opseu|pay directives?", None, "/tbs/cpslrc-collective-agreements"),
    (r"translation", None, "/translations"),
    (r"accessibility (in the )?ops|ops accessibility", None, "/tbs/accessibility-in-the-ontario-public-service"),
]
_SITE = re.compile(r"\b(inside\s*ops|insideops|intranet|intra\.ontario\.ca)\b", re.I)
_STOP = set("""a an and are about any at be by can could do does for from get give how i in is it its me my next
of on or our please show tell than that the their there this to up us we what whats when where which who why will
with would you your insideops inside ops intranet intra ontario ca""".split())


def topics(text):
    return [(q, p) for pattern, q, p in TOPICS if re.search(pattern, text or "", re.I)]


_ASKS = re.compile(r"\s*(what|when|where|which|who|why|how|is|are|does|do i|tell me|show me|list|find)\b", re.I)
_DOES = re.compile(r"\b(make|create|set ?up|set|remind|reminder|schedule|send|open|add|fill|submit|write|draft|email|"
                   r"start|launch|delete|move|copy|save|type|click)\b", re.I)


def wants(text):
    """Named InsideOPS, or a topic question. A topic word inside a request to do something stays a PC task."""
    text = text or ""
    if _SITE.search(text):
        return True
    return bool(topics(text)) and bool(_ASKS.match(text) or not _DOES.search(text))


def keywords(text):
    return [w for w in re.findall(r"[\w'-]+", (text or "").lower()) if w not in _STOP and len(w) > 1]


def api_paths(link):
    """The JSON calls that hold a page's text, given its InsideOPS link."""
    parts = [p for p in urlparse(link).path.split("/") if p]
    if len(parts) >= 4 and all(p.isdigit() for p in parts[-4:-1]):
        return "post", parts[-1]
    if len(parts) == 1:
        return "section", parts[0]
    return ("page", parts[-1]) if parts else (None, None)


def to_text(html):
    html = re.sub(r"(?is)<(script|style)\b.*?</\1>", "", html or "")
    html = re.sub(r"(?i)<li\b[^>]*>", "\n- ", html)
    html = re.sub(r"(?i)<br\s*/?>|</(p|div|li|h\d|tr|table|ul|ol|summary|details)>", "\n", html)
    html = re.sub(r"(?i)</t[dh]>", " | ", html)
    text = unescape(re.sub(r"<[^>]+>", "", html)).replace("\xa0", " ")
    text = re.sub(r"[ \t\r]+", " ", text)
    return re.sub(r"\n\s*\n+", "\n\n", text).strip()


class _Site:
    def __init__(self, chat, log):
        self.chat, self.log, self._woke = chat, log, False

    def get(self, path):
        for attempt in (0, 1):
            r = self.chat.context.request.get(HOME + path, timeout=20_000)
            if r.ok and "json" in (r.headers.get("content-type") or ""):
                return r.json()
            if attempt or self._woke:
                break
            self._wake()
        raise RuntimeError("InsideOPS didn't answer. Open intra.ontario.ca once in the Copilot browser to sign in.")

    def _wake(self):
        """Load the home page once in the background tab so the site's sign-in cookie is refreshed."""
        self._woke = True
        self.log("InsideOPS: refreshing sign-in in the background tab")
        page = self.chat.side_page()
        page.goto(HOME + "/", wait_until="domcontentloaded", timeout=45_000)
        page.wait_for_timeout(4000)

    def search(self, words):
        """Search with keywords, dropping the last word until something matches (search needs every word)."""
        words = list(words)
        while words:
            q = " ".join(words)
            hits = self.get(_SEARCH.format(q=quote(q), y=datetime.now().year))["hits"]["hits"]
            if hits:
                self.log(f"InsideOPS search '{q}': {len(hits)} results")
                return [h["_source"] for h in hits]
            words.pop()
        return []

    def read(self, link):
        kind, slug = api_paths(link)
        if kind == "post":
            item = (self.get(f"/api/posts/slug/{slug}") or [{}])[0]
            return item.get("title", {}).get("rendered", ""), item.get("modified", ""), to_text(item.get("content", {}).get("rendered"))
        if kind == "section":
            section = self.get(f"/api/section/slug/{slug}")
            page_id, title, modified = section.get("section_admin", {}).get("section_home_page"), section.get("name", slug), ""
        elif kind == "page":
            items = self.get(f"/api/pages/slug/{slug}") or []
            item = next((i for i in items if urlparse(i.get("link", "")).path.rstrip("/") == urlparse(link).path.rstrip("/")),
                        items[0] if items else {})
            page_id, title, modified = item.get("id"), item.get("title", {}).get("rendered", slug), item.get("modified", "")
        else:
            return "", "", ""
        if not page_id:
            return title, modified, ""
        return unescape(title), modified, to_text(self.get(f"/wp-json/wp/v2/pages/content/{page_id}").get("content"))


def answer(task, chat, log, status):
    """Answer from InsideOPS pages: mapped pages first, then search. Returns the reply text."""
    status("Checking InsideOPS")
    site = _Site(chat, log)
    year = datetime.now().year
    found = topics(task)
    link = lambda hit: hit.get("external_url") or HOME + hit["post_link"]
    links = [link(h) for q, _p in found if q for h in site.search(q.format(year=year).split())[:1]]
    links += [HOME + p for _q, p in found if p]
    links += [link(h) for h in site.search(keywords(task))[:2]]
    pages = []
    for link in dict.fromkeys(links):
        if len(pages) == _PAGES or not link.startswith(HOME):
            continue
        chat._check_stop()
        status("Reading InsideOPS")
        try:
            title, modified, text = site.read(link)
        except Exception as exc:
            log(f"InsideOPS read failed: {link}: {exc}")
            continue
        if text:
            pages.append({"title": title or link, "url": link, "modified": modified[:10], "text": text[:_PAGE_CHARS]})
    if not pages:
        return f"I couldn't find this on InsideOPS. Try different words, or open [InsideOPS]({HOME})."
    status("Writing the answer")
    return _write(task, pages, chat, log)


_FORMAT = """Write the answer in Markdown, shaped by the question:
- First line: the direct answer to exactly what was asked, with the key fact in **bold** (a date, amount, name, yes/no).
- Then only what helps: a short bullet list for several items (e.g. upcoming dates, steps, contacts), a numbered list for steps,
  a small table only when comparing. Use the weekday with dates. Link to InsideOPS pages inline where useful.
- Answer only what was asked; no page summaries, no headings, no preamble, no sources list (the app adds it).
- Plain Markdown only: no JSON, no code blocks."""


def _write(task, pages, chat, log):
    """Copilot writes the answer from the page text only, as plain Markdown."""
    now = datetime.now()
    sources = "\n".join(f"- [{p['title']}]({p['url']})" + (f", updated {p['modified']}" if p["modified"] else "") for p in pages)
    material = "\n\n".join(f"### {p['title']}\n{p['url']}\n{p['text']}" for p in pages)
    prompt = ("Answer using ONLY the InsideOPS page text below. If it doesn't answer the question, say so in one line. "
              "No other knowledge, no PC actions.\n"
              f"{_FORMAT}\nToday is {now:%A, %B %d, %Y}, {now:%H:%M}.\nQuestion:\n{task}\n\nInsideOPS pages:\n{material}")
    try:
        summary = _plain(chat.send(prompt))
    except Exception as exc:
        log(f"InsideOPS answer failed: {exc}")
        summary = ""
    if not summary:
        summary = "I found these InsideOPS pages but couldn't write the answer. They should cover it:"
    return f"{summary}\n\n**Source:** InsideOPS, read at {now:%H:%M}\n{sources}"


def _plain(reply):
    """Copilot's text, unwrapped if it answered in the action JSON anyway."""
    actions = protocol.parse_actions(reply.text, reply.code_blocks) or []
    done = next((a.get("summary") for a in actions if a.get("action") == "done"), "")
    if done:
        return done.strip()
    text = reply.text.strip()
    if re.match(r"^\s*(json\s*)?\{", text, re.I):
        found = re.search(r'"summary"\s*:\s*"((?:[^"\\]|\\.)*)', text)
        try:
            return json.loads('"' + found.group(1).rstrip("\\") + '"').strip() if found else ""
        except ValueError:
            return ""
    return text


def _check():
    assert wants("can you tell me whats the next pay day and next holiday")
    assert wants("what's the vacation policy on InsideOPS?") and wants("how do I book a desk")
    assert not wants("summarise my unread mail") and not wants("open notepad")
    assert not wants("Can you make me reminder for every friday to fillout timesheet")
    assert not wants("open WIN and fill my timesheet") and wants("when is the timesheet deadline")
    assert topics("next pay day and next holiday")[0][0] == "{year} pay and holiday dates"
    assert keywords("What does InsideOPS say about flexible work?") == ["say", "flexible", "work"]
    assert api_paths(HOME + "/topical/2025/11/12/2026-pay-and-holiday-dates-accessible-version") == \
        ("post", "2026-pay-and-holiday-dates-accessible-version")
    assert api_paths(HOME + "/ops/how-to-read-your-paystub") == ("page", "how-to-read-your-paystub")
    assert api_paths(HOME + "/pay") == ("section", "pay")
    assert to_text("<p>Pay Days</p><ul><li>November 19</li></ul><p>A&amp;B</p>") == "Pay Days\n\n- November 19\n\nA&B"
    class Reply:
        def __init__(self, text):
            self.text, self.code_blocks = text, []
    cut = 'JSON {"actions":[{"action":"done","summary":"Your next pay day is **December 3, 2026**.\\n- Today is a pay day.\\nSource: [2026 pay'
    assert _plain(Reply(cut)).startswith("Your next pay day is **December 3, 2026**.\n- Today is a pay day.")
    assert _plain(Reply("Your next pay day is **Thursday, December 3**.")) == "Your next pay day is **Thursday, December 3**."
    print("insideops ok")


if __name__ == "__main__":
    _check()
