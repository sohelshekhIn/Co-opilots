"""Ride home for the today strip: an approximate drive, or the next GO train.

Shown only from 45 minutes before the workday ends. The schedule is the published
GO timetable for today, kept in memory. Nothing is written to disk.
"""

import csv
import difflib
import io
import json
import os
import re
import zipfile
from datetime import datetime, timedelta
from urllib.request import Request, urlopen

LINES = ("Lakeshore West", "Lakeshore East", "Kitchener", "Milton", "Barrie", "Richmond Hill", "Stouffville")
_GTFS = "https://assets.metrolinx.com/raw/upload/Documents/Metrolinx/Open%20Data/GO-GTFS.zip"
_UA = "Co-opilot/1.0 (desktop commute estimate)"
_zip = None
_tables = {}
_trip_cache = {}
_places = {}
_here = None


def due(now, end_hour):
    """True from 45 minutes before the workday ends until two hours after."""
    end = now.replace(hour=end_hour, minute=0, second=0, microsecond=0)
    return end - timedelta(minutes=45) <= now <= end + timedelta(hours=2)


def card(settings, now=None):
    """The home cell, or None when it should not be on screen."""
    now = now or datetime.now()
    mode = settings.get("home_mode")
    if mode not in ("drive", "train") or not due(now, int(settings.get("workday_end") or 17)):
        return None
    try:
        if mode == "drive":
            return _drive(settings.get("home_address") or "")
        return _train(settings.get("go_line") or "", settings.get("go_stop") or "", now)
    except Exception:
        return {"k": "Home", "value": "–", "detail": "Couldn't check the trip just now", "ask": ""}


def _get(url, timeout=20):
    return urlopen(Request(url, headers={"User-Agent": _UA}), timeout=timeout)


def _drive(address):
    address = " ".join(str(address).split())
    if not address:
        return {"k": "Drive home", "value": "–", "detail": "Add a home address in Settings", "ask": "settings"}
    dest = _geocode(address)
    origin = _origin()
    if not dest or not origin:
        return {"k": "Drive home", "value": "–", "detail": "Couldn't place that address", "ask": "settings"}
    route = json.load(_get(
        "https://router.project-osrm.org/route/v1/driving/"
        f"{origin[1]},{origin[0]};{dest[1]},{dest[0]}?overview=false"))
    if route.get("code") != "Ok":
        return {"k": "Drive home", "value": "–", "detail": "No driving route found", "ask": ""}
    leg = route["routes"][0]
    mins = max(1, round(leg["duration"] / 60))
    km = max(1, round(leg["distance"] / 1000))
    return {"k": "Drive home", "value": _span(mins), "detail": f"About {km} km, approximate",
            "ask": f"How long to drive to {address}?"}


def _geocode(address):
    if address not in _places:
        found = json.load(_get(
            "https://nominatim.openstreetmap.org/search?format=json&limit=1&q=" + _quote(address)))
        _places[address] = (float(found[0]["lat"]), float(found[0]["lon"])) if found else None
    return _places[address]


def _origin():
    """Where this network appears to be. Approximate, and cached for the session."""
    global _here
    if _here is None:
        found = json.load(_get("https://ipwho.is/"))
        _here = (float(found["latitude"]), float(found["longitude"])) if found.get("success", True) else None
    return _here


def _quote(text):
    from urllib.parse import quote
    return quote(text)


def _train(line, stop, now):
    """Next train from Union, where the office is, to the home station."""
    if line not in LINES or not stop.strip():
        return {"k": "GO train", "value": "–", "detail": "Choose a line and a station in Settings", "ask": "settings"}
    trips = _trips(line, now.strftime("%Y%m%d"))
    if not trips:
        return {"k": line, "value": "–", "detail": "No GO schedule published for today", "ask": ""}
    names = {stop_name for trip in trips for _seq, stop_name, _minute in trip["stops"]}
    dest = _pick(names, stop)
    origin = _pick(names, "Union Station")
    if not dest:
        return {"k": line, "value": "–", "detail": f"No station matching {stop.strip()}", "ask": "settings"}
    if not origin:
        return {"k": line, "value": "–", "detail": "Union Station is not on today's timetable", "ask": ""}
    return _toward(trips, origin, dest, line, now)["card"]


def _from_times(times, line, name, now):
    """times: (minute, direction, headsign). Shows the next train each way, since home could be either direction."""
    now_m = now.hour * 60 + now.minute
    nxt = {}
    for m, way, sign in sorted(times):
        if m >= now_m and sign != _sign(name):
            nxt.setdefault(way, (m, sign))
    if not nxt:
        return {"k": line, "value": "None", "detail": f"No more trains from {name} today", "ask": ""}
    first = min(m for m, _ in nxt.values())
    detail = " · ".join(f"{m // 60 % 24:02d}:{m % 60:02d} to {sign}" for m, sign in sorted(nxt.values()))
    return {"k": f"{line} from {name}", "value": _span(first - now_m), "detail": detail,
            "ask": f"When is the next {line} train from {name}?"}


def _span(mins):
    return f"{mins}m" if mins < 60 else f"{mins // 60}h {mins % 60:02d}m"


def _pick(table, query):
    q = " ".join(query.casefold().replace("station", "").split())
    q = q[:-3] if q.endswith(" go") else q
    bare = {name.casefold().removesuffix(" go").removesuffix(" station"): name for name in table}
    if q in bare:
        return bare[q]
    # Whole word first, so "union" is Union Station and not Unionville.
    words = [key for key in bare if q in key.split()]
    if words:
        return bare[min(words, key=len)]
    hits = [key for key in bare if q and q in key]
    if hits:
        return bare[min(hits, key=len)]
    close = difflib.get_close_matches(q, list(bare), n=1, cutoff=0.75)
    return bare[close[0]] if close else None


def _toward(trips, origin, dest, line, now):
    """Trips that stop at origin and later at dest. The office is at Union, so that is the usual origin."""
    now_m = now.hour * 60 + now.minute
    hits = []
    for trip in trips:
        stops = trip["stops"]
        at_origin = next((stop for stop in stops if stop[1] == origin), None)
        at_dest = next((stop for stop in stops if at_origin and stop[1] == dest and stop[0] > at_origin[0]), None)
        if at_origin and at_dest and at_origin[2] >= now_m:
            hits.append((at_origin[2], at_dest[2], trip["line"], trip["sign"]))
    hits.sort()
    where = _sign(dest)
    if not hits:
        card = {"k": line, "value": "None", "detail": f"No more trains from {_sign(origin)} to {where} today", "ask": ""}
        return {"card": card, "hits": []}
    leave, arrive, _line, _head = hits[0]
    later = f" · then {_clock(hits[1][0])}" if len(hits) > 1 else ""
    card = {"k": f"{_sign(origin)} to {where}", "value": _span(leave - now_m),
            "detail": f"{_clock(leave)} from {_sign(origin)}, {where} {_clock(arrive)}{later}",
            "ask": f"When is the next train from {_sign(origin)} to {where}?"}
    return {"card": card, "hits": hits}


def _clock(minute):
    return f"{minute // 60 % 24:02d}:{minute % 60:02d}"


def _day(line, day):
    """Departure minutes per station for one line on one date. None when that date is not published."""
    key = (line, day)
    if key in _tables:
        return _tables[key]
    trips = _trips(line, day)
    if not trips:
        _tables[key] = None
        return None
    table = {}
    for trip in trips:
        for _seq, name, minute in trip["stops"]:
            table.setdefault(name, set()).add((minute, trip["way"], trip["sign"]))
    table = {name: sorted(times) for name, times in table.items()}
    _tables[key] = table
    return table


def _trips(line, day):
    """One line's trains that day: stops in order, so a destination can be told from a passing station."""
    key = (line, day)
    if key in _trip_cache:
        return _trip_cache[key]
    z = _feed()
    dates = _rows(z, "calendar_dates.txt")
    services = {row["service_id"] for row in dates if row["date"] == day and row["exception_type"] == "1"}
    if not services:
        _trip_cache[key] = None
        return None
    routes = {row["route_id"] for row in _rows(z, "routes.txt")
              if row["route_long_name"] == line and row["route_type"] == "2"}
    info = {row["trip_id"]: (row["direction_id"], _sign(row["trip_headsign"])) for row in _rows(z, "trips.txt")
            if row["route_id"] in routes and row["service_id"] in services}
    names = {row["stop_id"]: row["stop_name"] for row in _rows(z, "stops.txt")}
    built = {}
    for row in _rows(z, "stop_times.txt"):
        if row["trip_id"] not in info:
            continue
        name = names.get(row["stop_id"])
        if name:
            built.setdefault(row["trip_id"], []).append(
                (int(row["stop_sequence"]), name, _minutes(row["departure_time"] or row["arrival_time"])))
    out = [{"line": line, "way": info[trip_id][0], "sign": info[trip_id][1], "stops": sorted(stops)}
           for trip_id, stops in built.items()]
    _trip_cache[key] = out or None
    return _trip_cache[key]


_ASK = re.compile(r"\b(go\s+)?(train|trains|departures?|go transit)\b", re.I)


def _saved_stop():
    try:
        path = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "Co-opilot", "settings.json")
        with open(path, encoding="utf-8") as f:
            return json.load(f).get("go_stop") or ""
    except (OSError, ValueError):
        return ""


def _line_named(text):
    hits = [line for line in LINES if line.casefold() in text.casefold()]
    return hits[0] if len(hits) == 1 else None


def _place(text, word):
    flat = " ".join(re.sub(r"[^\w\s]", " ", text.casefold()).split())
    found = re.search(
        rf"\b{word}\s+([a-z][a-z .'-]{{1,30}}?)(?:\s+go\b|\s+station\b|$|\s+(?:to|from|today|tonight|now|train|trains)\b)",
        flat)
    return found.group(1).strip() if found else ""


def ask(text, now=None):
    """Answer a GO train question from the official timetable, or None when it isn't one.

    The office is at Union, so a named station is the destination and the train leaves from Union.
    """
    if not _ASK.search(text):
        return None
    now = now or datetime.now()
    day = now.strftime("%Y%m%d")
    chosen = _line_named(text)
    trips = [trip for line in ([chosen] if chosen else LINES) for trip in (_trips(line, day) or [])]
    if not trips:
        return "GO hasn't published a timetable for today, so I can't check trains."
    names = {stop[1] for trip in trips for stop in trip["stops"]}
    origin, dest = _intent(names, text)
    if dest is None and origin:
        return _board(_departures(trips, origin), origin, now)
    if not dest:
        return "Which station are you going to? Name it, or set your home station in Settings under Way home."
    origin = origin or _pick(names, "Union Station")
    if not origin:
        return "Union Station is not on today's timetable."
    return _ride(_toward(trips, origin, dest, chosen or "GO", now), origin, dest, now)


def _intent(names, text):
    """Destination first. Origin stays Union unless the question names both ends."""
    to = _pick(names, _place(text, "to")) if _place(text, "to") else None
    frm = _pick(names, _place(text, "from")) if _place(text, "from") else None
    if frm is None and _place(text, "at"):
        frm = _pick(names, _place(text, "at"))
    if to and frm:
        return frm, to
    if to:
        return None, to
    stripped = text
    for line in sorted(LINES, key=len, reverse=True):
        stripped = re.sub(re.escape(line), " ", stripped, flags=re.I)
    named = frm or _station_in(names, stripped)
    if named and _sign(named) == "Union":
        return named, None
    if named:
        return None, named
    saved = _pick(names, _saved_stop())
    return (None, saved) if saved else (None, None)


def _station_in(names, text):
    flat = " " + " ".join(re.sub(r"[^\w\s]", " ", text.casefold()).split()) + " "
    bare = {}
    for name in names:
        bare.setdefault(name.casefold().removesuffix(" go").removesuffix(" station"), name)
    hits = [key for key in bare if f" {key} " in flat]
    if hits:
        return bare[max(hits, key=len)]
    for token in sorted(flat.split(), key=len, reverse=True):
        if len(token) >= 5 and _pick(names, token):
            return _pick(names, token)
    return None


def _departures(trips, origin):
    times = []
    for trip in trips:
        for _seq, name, minute in trip["stops"]:
            if name == origin:
                times.append((minute, trip["line"], trip["way"], trip["sign"]))
                break
    return times


def _ride(found, origin, dest, now):
    hits = found["hits"]
    if not hits:
        return f"No more trains from **{_sign(origin)}** to **{_sign(dest)}** today."
    now_m = now.hour * 60 + now.minute
    out = [f"Next trains from **{_sign(origin)}** to **{_sign(dest)}** (it's {now:%H:%M}):", ""]
    by_line = {}
    for leave, arrive, line, _head in hits:
        by_line.setdefault(line, []).append((leave, arrive))
    for line, rows in sorted(by_line.items(), key=lambda kv: kv[1][0][0]):
        bits = ", ".join(f"{_clock(leave)} (in {_span(leave - now_m)}), arrives {_clock(arrive)}" for leave, arrive in rows[:3])
        out.append(f"- {line}: {bits}")
    out += ["", _cite(now)]
    return "\n".join(out)


def _cite(now):
    return ("Scheduled times, not live tracking. "
            "Source: [GO Transit published timetable](https://www.gotransit.com/en/open-data), "
            f"checked at {now:%H:%M}. Live platform and delays: [GO departures](https://www.gotransit.com/en/trip-planning/departures).")


def _board(times, name, now):
    now_m = now.hour * 60 + now.minute
    short = _sign(name)
    later = sorted(t for t in times if t[0] >= now_m and t[3] != short)
    if not later:
        return f"No more GO trains leave **{name}** today."
    ways = {}
    for m, line, way, sign in later:
        ways.setdefault((line, way), []).append((m, sign))
    out = [f"Next GO trains from **{name}** (it's {now:%H:%M}):", ""]
    for (line, _way), rows in sorted(ways.items(), key=lambda kv: kv[1][0][0]):
        trips = ", ".join(f"{m // 60 % 24:02d}:{m % 60:02d} to {sign} (in {_span(m - now_m)})" for m, sign in rows[:3])
        out.append(f"- {line}: {trips}")
    out += ["", _cite(now)]
    return "\n".join(out)


def _sign(headsign):
    """'KI - Kitchener GO' -> 'Kitchener'."""
    return headsign.split(" - ", 1)[-1].removesuffix(" GO").removesuffix(" Station").strip() or "?"


def _feed():
    global _zip
    if _zip is None:
        _zip = zipfile.ZipFile(io.BytesIO(_get(_GTFS, timeout=60).read()))
    return _zip


def _rows(z, name):
    return csv.DictReader(io.TextIOWrapper(z.open(name), encoding="utf-8-sig"))


def _minutes(text):
    hour, minute, _second = (int(part) for part in text.split(":"))
    return hour * 60 + minute


def _check():
    assert due(datetime(2026, 11, 17, 16, 15), 17)
    assert not due(datetime(2026, 11, 17, 16, 14), 17)
    assert not due(datetime(2026, 11, 17, 19, 1), 17)
    stops = {"Bramalea GO": [(16 * 60 + 4, "1", "Union"), (17 * 60 + 34, "0", "Kitchener"),
                             (17 * 60 + 40, "0", "Georgetown"), (17 * 60 + 50, "1", "Union")],
             "Union Station GO": [(17 * 60, "0", "Kitchener")]}
    for typed in ("bramalea", "Bramlea GO", " BRAMALEA go ", "Bramalea Station"):
        assert _pick(stops, typed) == "Bramalea GO", typed
    names = set(stops) | {"Unionville GO", "Kitchener GO"}
    assert _pick(names, "Union") == "Union Station GO"
    assert _pick(stops, "Oshawa") is None
    got = _from_times(stops["Bramalea GO"], "Kitchener", "Bramalea", datetime(2026, 11, 17, 16, 20))
    assert got["value"] == "1h 14m" and got["detail"] == "17:34 to Kitchener · 17:50 to Union", got
    assert _from_times([(25 * 60 + 5, "1", "Union")], "Kitchener", "X", datetime(2026, 11, 17, 23, 0))["value"] == "2h 05m"
    assert _from_times([(10 * 60, "1", "Union"), (19 * 60, "0", "X")], "Kitchener", "X", datetime(2026, 11, 17, 18, 0))["value"] == "None"
    board = {n: [(m, "Kitchener", w, sg) for m, w, sg in t] for n, t in stops.items()}
    txt = _board(board["Bramalea GO"], "Bramalea GO", datetime(2026, 11, 17, 16, 6))
    assert "17:34 to Kitchener (in 1h 28m)" in txt and "16:04" not in txt and "Source:" in txt, txt
    assert _intent(names, "next KI kitchener go train from bramalea go") == (None, "Bramalea GO")
    assert _intent(names, "next train to bramlea") == (None, "Bramalea GO")
    assert _intent(names, "from bramalea to kitchener") == ("Bramalea GO", "Kitchener GO")
    assert _intent(names, "next train at union") == ("Union Station GO", None)
    trips = [
        {"line": "Kitchener", "way": "0", "sign": "Kitchener", "stops": [
            (1, "Union Station GO", 16 * 60 + 34), (4, "Bramalea GO", 17 * 60 + 12)]},
        {"line": "Kitchener", "way": "1", "sign": "Union", "stops": [
            (1, "Bramalea GO", 16 * 60 + 40), (6, "Union Station GO", 17 * 60 + 20)]},
    ]
    ride = _toward(trips, "Union Station GO", "Bramalea GO", "Kitchener", datetime(2026, 11, 17, 16, 20))
    assert ride["card"]["value"] == "14m" and ride["card"]["detail"].startswith("16:34 from Union, Bramalea 17:12"), ride
    assert "17:20" not in _ride(ride, "Union Station GO", "Bramalea GO", datetime(2026, 11, 17, 16, 20))
    assert ask("summarise my unread mail") is None
    assert _sign("KI - Kitchener GO") == "Kitchener" and _sign("MI - Union Station GO") == "Union"
    assert card({"home_mode": "off", "workday_end": 17}, datetime(2026, 11, 17, 16, 30)) is None
    assert card({"home_mode": "train", "workday_end": 17, "go_line": "Kitchener", "go_stop": ""},
                datetime(2026, 11, 17, 12, 0)) is None
    print("commute ok")


if __name__ == "__main__":
    _check()
