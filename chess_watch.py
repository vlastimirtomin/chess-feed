#!/usr/bin/env python3
"""
chess_watch.py - RSS feed of chess-results.com tournaments for a fixed list of FIDE IDs.
chess_watch.py - RSS фид турнира са chess-results.com за унапред задату листу FIDE ID бројева.

Normal run : python chess_watch.py
Other modes: --discover (print the search form's fields), --export-env (settings for the workflow),
             --websub-ping (notify WebSub hubs), --from-file "dir/{id}.xlsx" (offline test).
All settings live in config.ini, the player list in players.txt. See README.md / README.sr.md.
"""
import argparse
import configparser
import io
import json
import os
import random
import re
import sys
import time
from datetime import date, datetime, timedelta, timezone
from datetime import time as dtime
from email.utils import format_datetime
from pathlib import Path
from urllib.parse import urljoin
from xml.sax.saxutils import escape as xe

BASE = "https://chess-results.com"
SEARCH_PATH = "/SpielerSuche.aspx"
LAN = "14"  # Serbian: used for the search page and for every generated link
EXCEL_RE = re.compile(r"excel|xls|екцел|ексел", re.I)
SEX_TOKENS = {"M": "M", "F": "F", "М": "M", "Ж": "F"}  # Latin M/F and Cyrillic М/Ж
STATUS_TOKENS = {"A": "A", "I": "I", "А": "A", "И": "I"}  # Active / Inactive (Latin or Cyrillic)
DUE_SLACK = timedelta(minutes=45)  # a run may start a little earlier than "every N hours" after the last check

try:
    from zoneinfo import ZoneInfo

    TZ = ZoneInfo("Europe/Belgrade")
except Exception:  # tzdata missing
    TZ = timezone.utc


class Blocked(Exception):
    """Server answered 403/429: stop the whole run, do not retry."""


class SearchError(Exception):
    """One player's search failed."""


class NotExcel(SearchError):
    """The site answered, but not with an Excel file."""


# ---------------------------------------------------------------- config
DEFAULTS = {
    "requests": {"delay_seconds": "45", "jitter_seconds": "20", "max_players_per_run": "5", "max_failures": "3",
                 "use_date_filter": "yes", "date_format": "%Y-%m-%d", "newest_first_sort": "yes"},
    "window": {"finished_days": "33", "upcoming_days": "366"},
    "feed": {"player_prefix": "fide", "player_title": "Турнири: {label} - {fide}",
             "withdraw_after_days": "3", "max_items": "300", "announce_initial": "yes"},
    "groups": {"default": ""},
    "results": {"enabled": "no", "wait_days": "2", "stable_hours": "24", "retry_hours": "4",
                "retry_other_hours": "24", "give_up_days": "7", "max_per_run": "3", "min_games": "1"},
    "websub": {"enabled": "yes", "hubs": "https://pubsubhubbub.appspot.com/, https://websubhub.com/hub",
               "wait_seconds": "90"},
    "schedule": {"timezone": "Europe/Belgrade", "active_from": "05:00", "active_until": "01:00",
                 "inactive_from": "01:00", "inactive_until": "05:00",
                 "active_every_hours": "4", "inactive_every_hours": "20"},
    "publish": {"feed_repo": ""},
    "log": {"hours": "48"},
}


def load_config(path):
    cp = configparser.ConfigParser(interpolation=None)
    cp.read_dict(DEFAULTS)
    if Path(path).exists():
        cp.read(path, encoding="utf-8")
    return cp


def site_url(cfg, override=""):
    if override:
        return override if override.endswith("/") else override + "/"
    repo = cfg.get("publish", "feed_repo").strip()
    if not re.fullmatch(r"[\w.-]+/[\w.-]+", repo) or "YOUR-GITHUB-NAME" in repo:
        sys.exit("config.ini: set [publish] feed_repo to yourname/reponame (the PUBLIC feed repository).")
    owner, name = repo.split("/")
    host = f"{owner.lower()}.github.io"
    return f"https://{host}/" if name.lower() == host else f"https://{host}/{name}/"


def file_part(text, what):
    text = text.strip()
    if not re.fullmatch(r"[A-Za-z0-9_-]{1,40}", text):
        sys.exit(f"config.ini: {what} must be 1-40 letters, digits, - or _ (it becomes part of a file name).")
    return text


def load_groups(cfg):
    """[groups]  1 = kikinda | Title   ->  {"1": {"file": "kikinda", "title": "Title"}}"""
    groups, used = {}, set()
    for key, value in cfg.items("groups"):
        if not key.isdigit():
            continue
        name, sep, title = value.partition("|")
        if not sep or not title.strip():
            sys.exit(f"config.ini: group {key} must look like:  {key} = file_name | Feed title")
        fname = file_part(name, f"the file name of group {key}")
        if fname in used:
            sys.exit(f"config.ini: two groups use the file name {fname}")
        used.add(fname)
        groups[key] = {"file": fname, "title": title.strip()}
    if not groups:
        sys.exit("config.ini: define at least one group in the [groups] section, e.g.  1 = rss | My feed")
    return groups


def hubs(cfg):
    if not cfg.getboolean("websub", "enabled"):
        return []
    return [h for h in re.split(r"[,\s]+", cfg.get("websub", "hubs")) if h]


# ---------------------------------------------------------------- request log (last N hours)
class RequestLog:
    def __init__(self, path):
        self.path = Path(path)
        self.lines = []
        if self.path.exists():
            self.lines = [l for l in self.path.read_text(encoding="utf-8").splitlines() if l.strip()]

    def add(self, what, ident, ok, detail, secs):
        ts = datetime.now(TZ).isoformat(timespec="seconds")
        self.lines.append(f"{ts} | {what} | {ident} | {'OK' if ok else 'FAIL'} | {detail} | {secs:.1f}s")

    def note(self, text):
        self.lines.append(f"{datetime.now(TZ).isoformat(timespec='seconds')} | {text}")

    def save(self, hours):
        cutoff = datetime.now(TZ) - timedelta(hours=hours)
        keep = []
        for l in self.lines:
            try:
                if datetime.fromisoformat(l.split(" | ", 1)[0]) < cutoff:
                    continue
            except ValueError:
                pass
            keep.append(l)
        self.lines = keep
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text("\n".join(keep) + ("\n" if keep else ""), encoding="utf-8")


# ---------------------------------------------------------------- searching
def list_fields(form, newest_first=True):
    out = []
    for el in form.find_all(["input", "select", "textarea", "button"]):
        name = el.get("name")
        if not name:
            continue
        if el.name == "input":
            kind = (el.get("type") or "text").lower()
            value = el.get("value", "")
            if kind in ("checkbox", "radio") and el.has_attr("checked"):
                value = el.get("value", "on")
        elif el.name == "select":
            kind = "select"
            opt = el.find("option", selected=True) or el.find("option")
            value = opt.get("value", opt.text) if opt else ""
            for o in el.find_all("option"):  # "maximum number of lines": take the biggest
                if o.get("value", o.text).strip() == "2000" or o.text.strip() == "2000":
                    value = o.get("value", o.text)
            if newest_first and "sort" in name.lower():  # "FIDE-Id, Enddate of tournament descending"
                for o in el.find_all("option"):
                    if re.search(r"fide.*desc", o.text, re.I):
                        value = o.get("value", o.text)
        elif el.name == "textarea":
            kind, value = "textarea", el.text
        else:
            kind, value = (el.get("type") or "button").lower(), el.get("value", "")
        out.append((name, kind, value))
    return out


class Searcher:
    def __init__(self, session, base, lan, log, fide_field=None, excel_button=None, newest_first=True):
        self.session, self.base, self.lan, self.log = session, base, lan, log
        self.fide_field, self.excel_button = fide_field, excel_button
        self.newest_first = newest_first
        self.names = {}           # FIDE ID -> label, only used to make the private log readable
        self.date_fields = None   # (from-field, to-field) of the "tournament end between" filter
        self.last_filtered = False
        self.action = None
        self.template = {}
        self.button_payload = {}
        self.ctx = ("form", "-")

    def _request(self, method, url, **kw):
        import requests

        what, ident = self.ctx
        for attempt in range(3):
            t0 = time.time()
            try:
                r = self.session.request(method, url, timeout=60, **kw)
            except requests.RequestException as e:
                self.log.add(f"{method} {what}", ident, False, f"network error: {type(e).__name__}", time.time() - t0)
                if attempt == 2:
                    raise SearchError(f"network error: {e}")
                time.sleep(30 * (attempt + 1))
                continue
            self.log.add(f"{method} {what}", ident, r.status_code < 400, f"HTTP {r.status_code}", time.time() - t0)
            if r.status_code in (403, 429):
                raise Blocked(f"HTTP {r.status_code} from {url}")
            if r.status_code >= 500:
                if attempt == 2:
                    raise SearchError(f"HTTP {r.status_code}")
                time.sleep(30 * (attempt + 1))
                continue
            if r.status_code >= 400:
                raise SearchError(f"HTTP {r.status_code}")
            return r

    def form(self):
        from bs4 import BeautifulSoup

        self.ctx = ("form", "-")
        r = self._request("GET", f"{self.base}{SEARCH_PATH}?lan={self.lan}")
        form = BeautifulSoup(r.text, "html.parser").find("form")
        if form is None:
            raise SearchError("no <form> on the search page")
        return urljoin(r.url, form.get("action") or r.url), form

    def prepare(self):
        self.action, form = self.form()
        fields = list_fields(form, self.newest_first)
        von = [n for n, k, _ in fields if re.search(r"von_tag$", n, re.I)]
        bis = [n for n, k, _ in fields if re.search(r"bis_tag$", n, re.I)]
        self.date_fields = (von[0], bis[0]) if len(von) == 1 and len(bis) == 1 else None
        self.template = {}
        for name, kind, value in fields:
            if kind in ("submit", "button", "image", "reset", "file"):
                continue
            if kind in ("checkbox", "radio") and not value:
                continue
            self.template[name] = value

        if not self.fide_field:
            c = [n for n, k, _ in fields if k == "text" and re.search("fide", n, re.I)]
            if len(c) != 1:
                raise SearchError(f"cannot pick the FIDE field (candidates {c}); use --fide-field")
            self.fide_field = c[0]
        if not self.excel_button:
            c = [n for n, k, v in fields
                 if k in ("submit", "button", "image") and EXCEL_RE.search(n + " " + v)]
            if len(c) != 1:
                raise SearchError(f"cannot pick the Excel button (candidates {c}); use --excel-button")
            self.excel_button = c[0]
        kind = next(k for n, k, _ in fields if n == self.excel_button)
        value = next(v for n, k, v in fields if n == self.excel_button)
        if kind == "image":
            self.button_payload = {self.excel_button + ".x": "5", self.excel_button + ".y": "5"}
        else:
            self.button_payload = {self.excel_button: value}

    def search(self, fide_id, window=None):
        """window = (from_date, to_date) already formatted; filters by tournament end date."""
        r = None
        for _ in range(2):  # second try reloads the form in case its state went stale
            if self.action is None:
                self.prepare()
            payload = dict(self.template)
            payload[self.fide_field] = str(fide_id)
            payload.update(self.button_payload)
            self.last_filtered = bool(window and self.date_fields)
            if self.last_filtered:
                payload[self.date_fields[0]], payload[self.date_fields[1]] = window
            self.ctx = ("search", f"{fide_id} {self.names.get(fide_id, '')}".strip())
            r = self._request("POST", self.action, data=payload)
            if r.content.startswith(b"PK"):  # xlsx = zip
                return r.content
            self.log.add("POST search", fide_id, False, "answer was not an Excel file", 0)
            self.action = None
            time.sleep(3)
        Path("debug_response.html").write_bytes(r.content)
        raise NotExcel("response was not an Excel file (saved as debug_response.html)")


def _fetch_card(self, tnr, snr, ident):
    """Plain GET of one player's page in one tournament (art=9)."""
    self.ctx = ("card", ident)
    r = self._request("GET", f"{self.base}/tnr{tnr}.aspx?lan={self.lan}&art=9&snr={snr}")
    return r.content.decode("utf-8", errors="replace")


Searcher.fetch_card = _fetch_card


def discover(searcher):
    action, form = searcher.form()
    print(f"Form action: {action}\n")
    print(f"{'NAME':60} {'KIND':10} VALUE (as the script would send it)")
    for name, kind, value in list_fields(form, searcher.newest_first):
        shown = value if len(value) < 40 else value[:37] + "..."
        print(f"{name:60} {kind:10} {shown}")
    for sel in form.find_all("select"):
        print(f"\nOptions of {sel.get('name')}  (text -> value; * = selected on the page):")
        for o in sel.find_all("option"):
            mark = "*" if o.has_attr("selected") else " "
            print(f"  {mark} {o.text.strip()!r} -> {o.get('value', o.text)!r}")


# ---------------------------------------------------------------- parsing
def parse_sheet(data):
    import pandas as pd

    raw = pd.read_excel(io.BytesIO(data), header=None, dtype=str)
    hdr = next((i for i, row in raw.iterrows()
                if "tnr" in [str(v).strip().lower() for v in row]), None)
    if hdr is None:
        raise SearchError("header row with 'tnr' not found in the Excel file")
    df = raw.iloc[hdr + 1:].copy()
    df.columns = [str(c).strip().lower() for c in raw.iloc[hdr]]
    return df.dropna(subset=["tnr"])


def iso(v):
    s = str(v).strip()
    return f"{s[0:4]}-{s[4:6]}-{s[6:8]}" if len(s) == 8 and s.isdigit() else s


def in_window(df, today, finished_days, upcoming_days):
    """Tournaments whose END date is in [today - finished_days, today + upcoming_days]."""
    lo = int((today - timedelta(days=finished_days)).strftime("%Y%m%d"))
    hi = int((today + timedelta(days=upcoming_days)).strftime("%Y%m%d"))
    rows = []
    for _, r in df.iterrows():
        try:
            end = int(str(r["bis"]).strip())
        except ValueError:
            continue
        if not lo <= end <= hi:
            continue
        rows.append({
            "tnr": str(r["tnr"]).strip(),
            "name": str(r["turname"]).strip(),
            "start": iso(r["von"]),
            "end": iso(r["bis"]),
            "rounds": str(r.get("rundenanzahl", "")).strip(),
            "round_now": str(r.get("rundeaktuell", "")).strip(),
            "snr": str(r.get("snr", "")).strip(),
            "rank": str(r.get("rang", "")).strip(),
        })
    return rows


def status(t, today):
    """finished / ongoing / upcoming, judged by dates."""
    t_iso = today.isoformat()
    if t["end"] < t_iso:
        return "finished"
    if t["start"] <= t_iso:
        return "ongoing"
    return "upcoming"


def tournament_url(tnr):
    return f"{BASE}/tnr{tnr}.aspx?lan={LAN}"


# ---------------------------------------------------------------- players list
GROUPS_RE = re.compile(r"[Gg]?\d+(,\d+)*")


def parse_players(text, source, default_groups=()):
    """Lines: FIDE_ID [A|I] [M|F] [groups] label  (flags in any order; # starts a comment).
    A = active (default), I = inactive; M = male (default), F = female; groups = numbers like 1 or 1,3
    (or G1,3); without them the player joins the groups listed as 'default' in config.ini."""
    out, seen = [], set()
    for line in text.splitlines():
        line = line.split("#", 1)[0].strip()
        if not line:
            continue
        tokens = line.split()
        if not tokens[0].isdigit():
            sys.exit(f"{source}: a line must start with a FIDE ID, got: {line!r}")
        fide_id, rest, sex, state, groups, got = tokens[0], tokens[1:], "M", "A", None, set()
        while rest and len(got) < 3:
            t = rest[0].upper()
            if t in SEX_TOKENS and "sex" not in got:
                sex, got = SEX_TOKENS[t], got | {"sex"}
            elif t in STATUS_TOKENS and "state" not in got:
                state, got = STATUS_TOKENS[t], got | {"state"}
            elif GROUPS_RE.fullmatch(rest[0]) and "groups" not in got:
                groups, got = rest[0].lstrip("Gg").split(","), got | {"groups"}
            else:
                break
            rest = rest[1:]
        if fide_id not in seen:
            seen.add(fide_id)
            out.append({"fide": fide_id, "sex": sex, "state": state, "label": " ".join(rest), "n": len(out) + 1,
                        "groups": groups if groups is not None else list(default_groups)})
    return out


# ---------------------------------------------------------------- state + feed items
def load_state(path):
    p = Path(path)
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {"players": {}}


def add_feed_item(data, kind, player, pair, now, extra=None):
    item_id = f"{kind}:{player['fide']}:{pair['tnr']}"
    if any(i["id"] == item_id for i in data["feed"]):  # never announce the same thing twice
        return
    data["feed"].append({
        "id": item_id, "kind": kind, "fide": player["fide"], "player": player["label"] or f"FIDE {player['fide']}",
        "sex": player["sex"], "tnr": pair["tnr"], "tournament": pair["name"], "start": pair["start"],
        "end": pair["end"], "rounds": pair.get("rounds", ""),
        "finished": pair["end"] < now.date().isoformat(), "published": now.isoformat(),
        **(extra or {}),
    })


def track_pairs(data, player, rows, all_tnrs, today, now, announce, suspicious, withdraw_after, results_on=False):
    """Follow (player, tournament) pairs.

    - A pair seen for the first time becomes a "join" item (silent baseline when `announce` is False).
    - A pair of a not-yet-finished tournament that is missing from successful checks on
      `withdraw_after` different days becomes a "withdraw" item (once).
    - Failed/blocked checks never get here; an empty answer for a player who had rows before
      (`suspicious`) is not counted as a miss.
    """
    fide_id = player["fide"]
    pairs = data.setdefault("pairs", {})
    data.setdefault("feed", [])
    t_iso = today.isoformat()

    for t in rows:
        key = f"{fide_id}:{t['tnr']}"
        pair = pairs.get(key)
        if pair is None:
            pair = pairs[key] = {"fide": fide_id, "tnr": t["tnr"]}
            pair.update(name=t["name"], start=t["start"], end=t["end"], rounds=t["rounds"])
            # With result checks on, a tournament found only after it ended is not announced as "played"
            # here: it goes through the verification below and is announced only if the player really played.
            if announce and not (results_on and t["end"] < t_iso):
                add_feed_item(data, "join", player, pair, now)
        pair.update(name=t["name"], start=t["start"], end=t["end"], rounds=t["rounds"],
                    round_now=t["round_now"], snr=t["snr"], rank=t["rank"], missing=0, last_miss=None)

    if suspicious:
        return
    for pair in pairs.values():
        if pair["fide"] != fide_id or pair["tnr"] in all_tnrs or pair["end"] < t_iso:
            continue
        if pair.get("last_miss") != t_iso:  # at most one miss per calendar day
            pair["last_miss"] = t_iso
            pair["missing"] = pair.get("missing", 0) + 1
        if pair["missing"] >= withdraw_after and not pair.get("withdraw_announced"):
            pair["withdraw_announced"] = True
            add_feed_item(data, "withdraw", player, pair, now)


def prune(data, order, today, finished_days, max_items):
    """Forget pairs once their tournament left the collection window (it can never be re-announced)."""
    lo = (today - timedelta(days=finished_days)).isoformat()
    data["pairs"] = {k: p for k, p in data.get("pairs", {}).items() if p["fide"] in order and p["end"] >= lo}
    data["feed"] = sorted(data.get("feed", []), key=lambda i: (i["published"], i["id"]))[-max_items:]


# ---------------------------------------------------------------- feed (Serbian Cyrillic)
def sr_date(s):
    return f"{s[8:10]}.{s[5:7]}.{s[0:4]}." if len(s) == 10 and s[4] == "-" else s


def sr_range(a, b):
    return sr_date(a) if a == b else f"{sr_date(a)} – {sr_date(b)}"


def sr_rounds(n):
    return "1 коло" if n == "1" else f"{n} кола"


def item_title(it):
    f = it.get("sex") == "F"
    if it["kind"] == "withdraw":
        return f"{it['player']} више није на стартној листи: {it['tournament']}"
    if it["kind"] == "result":
        bits = []
        if it.get("points"):
            bits.append(f"{it['points']}/{it['rounds']}" if it.get("rounds") else it["points"])
        if str(it.get("place", "")).isdigit():  # team events show "-" instead of a place
            bits.append(f"{it['place']}. место")
        tail = f" – {', '.join(bits)}" if bits else ""
        return f"{it['player']} {'је играла на' if f else 'је играо на'}: {it['tournament']}{tail}"
    if it.get("finished"):
        verb = "је учествовала на" if f else "је учествовао на"
    else:
        verb = "је пријављена за" if f else "је пријављен за"
    return f"{it['player']} {verb}: {it['tournament']}"


def render_feed(items, title, description, base, filename, hub_list):
    out = []
    for it in sorted(items, key=lambda i: (i["published"], i["id"]), reverse=True):
        body = sr_range(it["start"], it["end"])
        if it.get("rounds"):
            body += f" · {sr_rounds(it['rounds'])}"
        out.append(
            f"<item><title>{xe(item_title(it))}</title><link>{xe(it.get('link') or tournament_url(it['tnr']))}</link>"
            f"<guid isPermaLink=\"false\">{xe(it['id'])}</guid>"
            f"<pubDate>{format_datetime(datetime.fromisoformat(it['published']))}</pubDate>"
            f"<description>{xe(body)}</description></item>")
    last = max((i["published"] for i in items), default=None)
    build = f"<lastBuildDate>{format_datetime(datetime.fromisoformat(last))}</lastBuildDate>" if last else ""
    links = "".join(f"<atom:link rel=\"hub\" href=\"{xe(h)}\"/>" for h in hub_list)
    links += f"<atom:link rel=\"self\" type=\"application/rss+xml\" href=\"{xe(base + filename)}\"/>"
    return (f"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            f"<rss version=\"2.0\" xmlns:atom=\"http://www.w3.org/2005/Atom\"><channel>"
            f"<title>{xe(title)}</title><link>{xe(base)}</link>"
            f"<description>{xe(description)}</description><language>sr-Cyrl</language>"
            f"{build}{links}"
            + "".join(out) + "</channel></rss>\n")


BLANK_PAGE = ("<!doctype html><html><head><meta charset=\"utf-8\">"
              "<meta name=\"robots\" content=\"noindex, nofollow, noarchive\"><title></title></head><body></body></html>\n")


# ---------------------------------------------------------------- WebSub
def websub_ping(cfg, state_dir, log, session):
    hub_list = hubs(cfg)
    ping_file = Path(state_dir) / "ping_urls.txt"
    urls = [u.strip() for u in ping_file.read_text(encoding="utf-8").splitlines() if u.strip()] if ping_file.exists() else []
    if not hub_list or not urls:
        print("WebSub: nothing to ping.")
        return
    wait = cfg.getint("websub", "wait_seconds")
    print(f"WebSub: waiting {wait}s for the feed to go live, then pinging {len(hub_list)} hub(s) for {len(urls)} feed(s).")
    time.sleep(wait)
    import requests

    for hub in hub_list:
        for url in urls:
            t0 = time.time()
            try:
                r = session.post(hub, data={"hub.mode": "publish", "hub.url": url}, timeout=30)
                ok, detail = r.status_code in (200, 202, 204), f"HTTP {r.status_code}"
            except requests.RequestException as e:
                ok, detail = False, f"network error: {type(e).__name__}"
            log.add("POST websub", hub, ok, detail, time.time() - t0)
            time.sleep(2)
    ping_file.write_text("", encoding="utf-8")


# ---------------------------------------------------------------- results of finished tournaments
PLAYED = {"1", "0", "½", "1/2", "0,5", "0.5"}  # a game that was really played (not a bye or a forfeit)
OPPONENT_LINK = re.compile(r"art=9.*snr=\d+")


def card_url(tnr, snr):
    return f"{BASE}/tnr{tnr}.aspx?lan={LAN}&art=9&snr={snr}"


def parse_card(html, fide_id):
    """Read the page of one player in one tournament. Labels are the Serbian (lan=14) ones, with English fallbacks."""
    from bs4 import BeautifulSoup

    clean = lambda t: re.sub(r"\s+", " ", t.replace("\xa0", " ")).strip()
    soup = BeautifulSoup(html, "html.parser")
    info = {}
    for tr in soup.find_all("tr"):
        cells = tr.find_all(["td", "th"], recursive=False)
        if len(cells) == 2:
            k, v = clean(cells[0].get_text(" ")), clean(cells[1].get_text(" "))
            if k and v and k not in info:
                info[k] = v
    games = []
    for tr in soup.find_all("tr"):
        if not tr.find("a", href=OPPONENT_LINK):
            continue
        cells = tr.find_all("td", recursive=False)
        if len(cells) < 4 or not clean(cells[0].get_text(" ")).isdigit():
            continue
        games.append({"round": int(clean(cells[0].get_text(" "))), "result": clean(cells[-1].get_text(" "))})
    last = None
    m = re.search(r"(?:Последња измена|Last update)\s*(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2}):(\d{2})",
                  clean(soup.get_text(" ")))
    if m:
        d, mo, y, hh, mi, ss = map(int, m.groups())
        last = datetime(y, mo, d, hh, mi, ss, tzinfo=TZ)
    return {"identity_ok": str(fide_id) in info.values(), "games": games, "last_update": last,
            "points": info.get("Бодови") or info.get("Points") or "",
            "place": info.get("Место") or info.get("Rank") or ""}


def classify_card(card, now, min_games, stable_hours):
    """competed / not_played / pending (a game has no result yet) / recent (tournament page changed lately) /
    mismatch (the page is about someone else) / unparsed (no games found)."""
    if not card["identity_ok"]:
        return "mismatch"
    if not card["games"]:
        return "unparsed"
    if card["last_update"] and now - card["last_update"] < timedelta(hours=stable_hours):
        return "recent"
    if any(g["result"] == "" for g in card["games"]):
        return "pending"
    played = sum(1 for g in card["games"] if g["result"] in PLAYED)
    return "competed" if played >= min_games else "not_played"


def result_candidates(data, now, wait_days, give_up_days):
    """Pairs whose tournament ended more than wait_days ago and shows all rounds as played.
    Returns (to_try_now, gave_up): a pair is skipped until its next allowed try, and given up on
    when its tournament ended more than give_up_days ago."""
    today = now.date()
    limit = (today - timedelta(days=wait_days)).isoformat()
    todo, gave_up = [], []
    for pair in data.get("pairs", {}).values():
        if pair.get("result") or not pair.get("snr") or pair["end"] >= limit:
            continue
        if not pair.get("rounds") or pair.get("round_now") != pair.get("rounds"):
            continue
        if (today - date.fromisoformat(pair["end"])).days > give_up_days:
            pair["result"] = "gave_up"
            gave_up.append(pair)
            continue
        nxt = pair.get("result_next")
        if nxt and now < datetime.fromisoformat(nxt) - DUE_SLACK:
            continue
        todo.append(pair)
    todo.sort(key=lambda p: (p.get("result_tries", 0), p["end"], p["fide"], p["tnr"]))
    return todo, gave_up


def check_results(players, searcher, cfg, today, now, finished_days, upcoming_days, window, pause, count=3):
    """Diagnostic: for the first players find a finished tournament, fetch the player's page and show what the
    script reads from it, so that it can be compared with the page in a browser. Changes nothing."""
    shown = 0
    for n, p in enumerate(players):
        if shown >= count:
            break
        if n:
            pause()
        df = parse_sheet(searcher.search(p["fide"], window))
        done = [r for r in in_window(df, today, finished_days, upcoming_days)
                if status(r, today) == "finished" and r["rounds"] and r["round_now"] == r["rounds"] and r["snr"]]
        if not done:
            print(f"player #{p['n']}: no finished tournament in the window")
            continue
        r = max(done, key=lambda x: (x["end"], x["tnr"]))
        pause()
        card = parse_card(searcher.fetch_card(r["tnr"], r["snr"], f"{p['fide']} {p['label']}".strip()), p["fide"])
        verdict = classify_card(card, now, cfg.getint("results", "min_games"), cfg.getfloat("results", "stable_hours"))
        played = sum(1 for g in card["games"] if g["result"] in PLAYED)
        print(f"player #{p['n']}: {r['name']} (ended {r['end']}, {r['rounds']} rounds)\n"
              f"   games found: {len(card['games'])}, really played: {played}, results: "
              f"{[g['result'] for g in card['games']]}\n"
              f"   points on the page: {card['points'] or '?'}, place: {card['place'] or '?'}, "
              f"FIDE ID found on the page: {card['identity_ok']}, page last changed: {card['last_update'] or '?'}\n"
              f"   verdict: {verdict}\n   compare with: {card_url(r['tnr'], r['snr'])}")
        shown += 1
    if not shown:
        print("RESULT: nothing to check.")


def explain_player(p, searcher, cfg, today, now, finished_days, upcoming_days, window, pause, data, count=3):
    """Diagnostic for ONE player: every tournament in the window, whether the result check would pick it up,
    and if not, why. Opens at most `count` player pages. Changes nothing."""
    give_up = cfg.getint("results", "give_up_days")
    df = parse_sheet(searcher.search(p["fide"], window))
    rows = sorted(in_window(df, today, finished_days, upcoming_days), key=lambda r: (r["end"], r["tnr"]), reverse=True)
    print(f"{len(df)} rows in the search table, {len(rows)} in the window "
          f"(finished up to {finished_days} days ago, ending within {upcoming_days} days)\n")
    fetched = 0
    for r in rows:
        st = status(r, today)
        pair = (data or {}).get("pairs", {}).get(f"{p['fide']}:{r['tnr']}", {})
        print(f"- {r['name']}\n    ended {r['end']} | rounds {r['rounds'] or '?'}, current round {r['round_now'] or '?'} | "
              f"starting rank (snr) {r['snr'] or '?'} | place in table {r['rank'] or '?'} | dates say: {st}")
        if pair:
            print(f"    saved state: result={pair.get('result') or 'not decided'}, tries={pair.get('result_tries', 0)}, "
                  f"next try not before {pair.get('result_next') or '-'}")
        if st != "finished":
            print("    why not: the tournament is not finished according to its dates")
        elif not r["snr"] or r["snr"] in ("nan", "-"):
            print("    why not: the search table has no starting rank (snr), so the player's page cannot be opened")
        elif not r["rounds"] or r["round_now"] != r["rounds"]:
            print(f"    why not: the table shows round {r['round_now'] or '?'} of {r['rounds'] or '?'}, "
                  "which is not treated as 'all rounds played'")
        elif (today - date.fromisoformat(r["end"])).days > give_up:
            print(f"    why not: ended more than {give_up} days ago (give_up_days)")
        elif fetched < count:
            pause()
            card = parse_card(searcher.fetch_card(r["tnr"], r["snr"], f"{p['fide']} {p['label']}".strip()), p["fide"])
            fetched += 1
            verdict = classify_card(card, now, cfg.getint("results", "min_games"), cfg.getfloat("results", "stable_hours"))
            played = sum(1 for g in card["games"] if g["result"] in PLAYED)
            print(f"    opened the player's page: games found {len(card['games'])}, really played {played}, results "
                  f"{[g['result'] for g in card['games']]}\n"
                  f"    points on the page: {card['points'] or '?'}, place: {card['place'] or '?'}, FIDE ID on the page: "
                  f"{card['identity_ok']}, page last changed: {card['last_update'] or '?'}\n"
                  f"    verdict: {verdict}   link: {card_url(r['tnr'], r['snr'])}")
        else:
            print("    eligible (page not opened here: limit of pages for this check)")
    if not rows:
        print("No tournament of this player in the window.")


# ---------------------------------------------------------------- schedule
def hhmm(text):
    h, m = text.strip().split(":")
    return dtime(int(h), int(m))


def in_time_window(t, start, end):
    """True if time-of-day t lies in [start, end); the window may wrap past midnight (05:00-01:00)."""
    return start <= t < end if start <= end else (t >= start or t < end)


def is_due(player, entry, now, cfg):
    """Active players: inside the active window and active_every_hours after their last check.
    Inactive players: inside the inactive window and inactive_every_hours after their last check.
    A player that was never checked is due at any time."""
    if not entry or "checked" not in entry:
        return True
    if entry.get("checked_at"):
        last = datetime.fromisoformat(entry["checked_at"])
    else:
        last = datetime.combine(date.fromisoformat(entry["checked"]), dtime(0, 0), tzinfo=TZ)
    kind = "active" if player["state"] == "A" else "inactive"
    sec = cfg["schedule"]
    if not in_time_window(now.time(), hhmm(sec[f"{kind}_from"]), hhmm(sec[f"{kind}_until"])):
        return False
    return now - last >= timedelta(hours=float(sec[f"{kind}_every_hours"])) - DUE_SLACK


# ---------------------------------------------------------------- main
def load_players(path, groups, default_groups):
    if not Path(path).exists():
        sys.exit(f"{path} not found")
    players = parse_players(Path(path).read_text(encoding="utf-8"), path, default_groups)
    for p in players:
        for g in p["groups"]:
            if g not in groups:
                sys.exit(f"{path}: player #{p['n']} is in group {g}, which is not defined in config.ini [groups]")
    return players


def player_title(cfg, p):
    try:
        return cfg.get("feed", "player_title").format(label=p["label"] or "FIDE", fide=p["fide"])
    except (KeyError, IndexError, ValueError):
        sys.exit("config.ini: [feed] player_title may only use {label} and {fide}.")


def check_filter(players, searcher, today, finished_days, upcoming_days, window, pause, count=3):
    """Compare a filtered with an unfiltered search for the first players (3 requests per player pair)."""
    same = True
    for n, p in enumerate(players[:count]):
        if n:
            pause()
        full = parse_sheet(searcher.search(p["fide"]))
        pause()
        filt = parse_sheet(searcher.search(p["fide"], window))
        used = searcher.last_filtered
        a = {r["tnr"] for r in in_window(full, today, finished_days, upcoming_days)}
        b = {r["tnr"] for r in in_window(filt, today, finished_days, upcoming_days)}
        print(f"player #{p['n']}: unfiltered {len(full)} rows, {len(a)} in window | "
              f"filtered {len(filt)} rows, {len(b)} in window | date fields used: {used}")
        if a != b or not used:
            same = False
            print(f"  DIFFERENT. only without filter: {sorted(a - b)}  only with filter: {sorted(b - a)}")
    print("RESULT: the filter gives the same tournaments as the unfiltered search." if same
          else "RESULT: problems found - set use_date_filter = no in config.ini.")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default="config.ini")
    ap.add_argument("--players", default="players.txt")
    ap.add_argument("--state-dir", default="state")
    ap.add_argument("--out-dir", default="site")
    ap.add_argument("--discover", action="store_true")
    ap.add_argument("--check-filter", action="store_true",
                    help="compare filtered and unfiltered searches for the first 3 players")
    ap.add_argument("--export-env", action="store_true", help="print FEED_REPO and SITE_URL for the workflow")
    ap.add_argument("--websub-ping", action="store_true")
    ap.add_argument("--check-results", action="store_true",
                    help="show what the script reads from the player pages of up to 3 finished tournaments")
    ap.add_argument("--only-fide", default="", help="with --check-results: explain this one player in detail")
    ap.add_argument("--cards-dir", help="testing: read player pages from DIR/TNR_SNR.html instead of the site")
    ap.add_argument("--rebuild", action="store_true",
                    help="regenerate the feeds from the saved state without checking any player")
    ap.add_argument("--base", default=BASE)
    ap.add_argument("--fide-field")
    ap.add_argument("--excel-button")
    ap.add_argument("--from-file", help='parse local Excel files instead of fetching, e.g. "tests/{id}.xlsx"')
    ap.add_argument("--today", help="YYYY-MM-DD (testing)")
    ap.add_argument("--at", help="HH:MM local time (testing)")
    ap.add_argument("--site-url", default="", help="public address of the feeds, e.g. https://NAME.github.io/REPO/")
    ap.add_argument("--contact", default="", help="URL or email added to the User-Agent")
    args = ap.parse_args()

    cfg = load_config(args.config)
    state_dir = Path(args.state_dir)
    global TZ
    try:
        from zoneinfo import ZoneInfo as _ZI

        TZ = _ZI(cfg.get("schedule", "timezone").strip())
    except Exception:
        pass  # keep the default

    if args.export_env:
        print(f"FEED_REPO={cfg.get('publish', 'feed_repo').strip()}")
        print(f"SITE_URL={site_url(cfg)}")
        return

    log = RequestLog(state_dir / "request_log.txt")
    log_hours = cfg.getint("log", "hours")

    session = searcher = None
    if not args.from_file:
        import requests

        session = requests.Session()
        ua = "chess-watch/1.0 (personal tournament watcher"
        session.headers["User-Agent"] = ua + (f"; {args.contact})" if args.contact else ")")
        searcher = Searcher(session, args.base, LAN, log, args.fide_field, args.excel_button,
                            newest_first=cfg.getboolean("requests", "newest_first_sort"))

    if args.discover:
        if searcher is None:
            sys.exit("--discover needs network access")
        discover(searcher)
        return

    if args.websub_ping:
        websub_ping(cfg, state_dir, log, session)
        log.save(log_hours)
        return

    if args.today or args.at:  # simulated clock (testing)
        today = date.fromisoformat(args.today) if args.today else datetime.now(TZ).date()
        hh, mm = (args.at or "06:00").split(":")
        now = datetime.combine(today, dtime(int(hh), int(mm)), tzinfo=TZ)
        clock = lambda: now
    else:
        now = datetime.now(TZ)
        today = now.date()
        clock = lambda: datetime.now(TZ)
    delay, jitter = cfg.getfloat("requests", "delay_seconds"), cfg.getfloat("requests", "jitter_seconds")
    finished_days, upcoming_days = cfg.getint("window", "finished_days"), cfg.getint("window", "upcoming_days")
    fmt = cfg.get("requests", "date_format")
    window = ((today - timedelta(days=finished_days)).strftime(fmt),
              (today + timedelta(days=upcoming_days)).strftime(fmt))

    def pause():
        time.sleep(delay + random.uniform(0, jitter))

    default_groups = [g for g in re.split(r"[,\s]+", cfg.get("groups", "default")) if g]

    if args.check_filter:
        if searcher is None:
            sys.exit("--check-filter needs network access")
        check_players = load_players(args.players, load_groups(cfg), default_groups)
        searcher.names = {p["fide"]: p["label"] for p in check_players}
        check_filter(check_players, searcher, today, finished_days, upcoming_days, window, pause)
        log.save(log_hours)
        return

    if args.check_results:
        if searcher is None:
            sys.exit("--check-results needs network access")
        cr_players = load_players(args.players, load_groups(cfg), default_groups)
        searcher.names = {p["fide"]: p["label"] for p in cr_players}
        only = args.only_fide.strip()
        if only:
            target = next((p for p in cr_players if p["fide"] == only),
                          {"fide": only, "label": "", "n": 0, "sex": "M", "state": "A", "groups": []})
            explain_player(target, searcher, cfg, today, now, finished_days, upcoming_days, window, pause,
                           load_state(state_dir / "data.json"))
        else:
            check_results(cr_players, searcher, cfg, today, now, finished_days, upcoming_days, window, pause)
        log.save(log_hours)
        return

    # ------------------------------------------------ normal run
    groups = load_groups(cfg)
    prefix = file_part(cfg.get("feed", "player_prefix"), "[feed] player_prefix")
    base = site_url(cfg, args.site_url)
    hub_list = hubs(cfg)
    players = load_players(args.players, groups, default_groups)
    order = [p["fide"] for p in players]
    if searcher is not None:
        searcher.names = {p["fide"]: p["label"] for p in players}

    data = load_state(state_dir / "data.json")
    had_pairs = "pairs" in data  # false on the very first run: everything is a silent baseline
    data.setdefault("feed", [])
    data["players"] = {k: v for k, v in data["players"].items() if k in order}
    before_ids = {i["id"] for i in data["feed"]}

    results_on = cfg.getboolean("results", "enabled")
    wait_days, give_up_days = cfg.getint("results", "wait_days"), cfg.getint("results", "give_up_days")
    retry_hours, retry_other = cfg.getfloat("results", "retry_hours"), cfg.getfloat("results", "retry_other_hours")
    min_games, stable_hours = cfg.getint("results", "min_games"), cfg.getfloat("results", "stable_hours")

    # Who is due now? Active players inside the active window every active_every_hours, inactive players inside the
    # inactive window every inactive_every_hours, new players immediately. Longest-waiting first, at most max_players_per_run.
    due = [p for p in players if is_due(p, data["players"].get(p["fide"]), now, cfg)]
    due.sort(key=lambda p: (data["players"].get(p["fide"], {}).get("checked_at")
                            or data["players"].get(p["fide"], {}).get("checked", ""), p["n"]))
    max_run = cfg.getint("requests", "max_players_per_run")
    selected = [] if args.rebuild else (due[:max_run] if max_run > 0 else due)
    n_active = sum(1 for p in players if p["state"] == "A")
    print(f"{len(players)} players ({n_active} active, {len(players) - n_active} inactive); "
          f"{len(due)} due at {now.strftime('%H:%M')}, checking {len(selected)} in this run.")
    waiting = result_candidates(data, now, wait_days, give_up_days)[0] if results_on and not args.rebuild else []
    if not selected and not waiting and not args.rebuild:
        print("Nothing to do now.")
        return
    if args.rebuild:
        print("Rebuild: regenerating the feeds from the saved state, no player is checked.")

    # Server-side date filter ("tournament end between"): smaller downloads. It is verified on the fly:
    # if the site rejects it, or answers "nothing" while an unfiltered search finds tournaments, it is switched off.
    use_filter = cfg.getboolean("requests", "use_date_filter")
    filter_state = {"ok": None}  # None = not verified yet, True = verified, False = switched off

    def warn(msg):
        print(f"WARNING: {msg}", file=sys.stderr)
        log.note(f"WARNING | {msg}")

    def fetch_df(fide_id):
        """Returns (table, filtered?)."""
        if use_filter and filter_state["ok"] is not False:
            try:
                df = parse_sheet(searcher.search(fide_id, window))
            except NotExcel as e:
                if filter_state["ok"] is True:
                    raise
                filter_state["ok"] = False
                warn(f"the site did not accept the date filter ({e}); continuing without it")
                return parse_sheet(searcher.search(fide_id)), False
            if not searcher.last_filtered:
                filter_state["ok"] = False
                warn("date fields not found on the form; continuing without the filter")
                return df, False
            if len(df) > 0 or filter_state["ok"] is True:
                filter_state["ok"] = True
                return df, True
            pause()  # empty answer while still unverified: compare once with an unfiltered search
            full = parse_sheet(searcher.search(fide_id))
            if in_window(full, today, finished_days, upcoming_days):
                filter_state["ok"] = False
                warn("the date filter returned nothing although tournaments exist; continuing without it")
                return full, False
            return df, True
        return parse_sheet(searcher.search(fide_id)), False

    max_fail = cfg.getint("requests", "max_failures")
    withdraw_after = cfg.getint("feed", "withdraw_after_days")
    announce_initial = cfg.getboolean("feed", "announce_initial")

    ok = failures = consecutive = 0
    aborted = None
    for n, player in enumerate(selected):
        fide_id = player["fide"]
        try:
            if args.from_file:
                src = args.from_file.format(id=fide_id) if "{id}" in args.from_file else args.from_file
                df, filtered = parse_sheet(Path(src).read_bytes()), False
            else:
                if n:
                    pause()
                df, filtered = fetch_df(fide_id)
            rows = in_window(df, today, finished_days, upcoming_days)
            before = data["players"].get(fide_id, {})
            known = "checked" in before
            track_pairs(data, player, rows, set(df["tnr"].astype(str).str.strip()), today, now,
                        announce=(known and had_pairs) or announce_initial,
                        # an empty UNFILTERED answer for a player who had rows before is suspicious
                        suspicious=(not filtered and len(df) == 0 and before.get("rows", 0) > 0),
                        withdraw_after=withdraw_after, results_on=results_on)
            data["players"][fide_id] = {"checked": today.isoformat(),
                                        "checked_at": clock().isoformat(timespec="seconds"), "rows": len(df)}
            ok += 1
            consecutive = 0
            counts = {k: sum(1 for r in rows if status(r, today) == k) for k in ("ongoing", "upcoming", "finished")}
            result = f"ok - {counts['ongoing']} ongoing, {counts['upcoming']} upcoming, {counts['finished']} finished"
            print(f"player #{player['n']}: {result}")
            log.note(f"PLAYER | {fide_id} | {player['label'] or '-'} | {result}")
        except Blocked as e:
            aborted = f"blocked ({e}); skipping the rest of this run"
            print(f"player #{player['n']}: {aborted}", file=sys.stderr)
            log.note(f"PLAYER | {fide_id} | {player['label'] or '-'} | BLOCKED - {e}")
            break
        except SearchError as e:
            failures += 1
            consecutive += 1
            print(f"player #{player['n']}: FAILED - {e}", file=sys.stderr)
            log.note(f"PLAYER | {fide_id} | {player['label'] or '-'} | FAILED - {e}")
            if consecutive >= max_fail:
                aborted = f"{consecutive} failures in a row; skipping the rest of this run"
                print(aborted, file=sys.stderr)
                break

    # Results of finished tournaments: fetch the player's page once the tournament is over and final, check that the
    # player really played, and only then announce. A few pages per run. A tournament whose results are not final yet
    # is tried again after retry_hours; a page that cannot be read after retry_other_hours; after give_up_days: no more.
    if results_on and not args.rebuild and not aborted:
        by_fide = {p["fide"]: p for p in players}
        todo, gave_up = result_candidates(data, now, wait_days, give_up_days)
        for pair in gave_up:
            if pair["fide"] in by_fide:
                log.note(f"RESULT | {pair['fide']} | {by_fide[pair['fide']]['label'] or '-'} | tnr {pair['tnr']} | "
                         f"gave up (ended more than {give_up_days} days ago)")
        todo = [pr for pr in todo if pr["fide"] in by_fide]
        max_pages, k = cfg.getint("results", "max_per_run"), 0
        for pair in todo:
            nxt = pair.get("result_next")
            if nxt and now < datetime.fromisoformat(nxt) - DUE_SLACK:
                continue  # another player's page of the same tournament showed that it is not final yet
            if k >= max_pages:
                break
            player = by_fide[pair["fide"]]
            pair["result_tries"] = pair.get("result_tries", 0) + 1
            try:
                if args.cards_dir:
                    html = (Path(args.cards_dir) / f"{pair['tnr']}_{pair['snr']}.html").read_text(encoding="utf-8")
                else:
                    if k or selected:
                        pause()
                    html = searcher.fetch_card(pair["tnr"], pair["snr"], f"{pair['fide']} {player['label']}".strip())
                card = parse_card(html, pair["fide"])
                verdict = classify_card(card, now, min_games, stable_hours)
            except Blocked as e:
                aborted = f"blocked ({e}); skipping the rest of this run"
                print(f"result check: {aborted}", file=sys.stderr)
                break
            except (SearchError, OSError) as e:
                verdict = "fetch_failed"
                print(f"result check #{k + 1}: FAILED - {e}", file=sys.stderr)
            if verdict == "competed":
                place = card["place"] if card["place"].isdigit() else (pair["rank"] if str(pair.get("rank", "")).isdigit() else "")
                add_feed_item(data, "result", player, pair, now, extra={
                    "points": card["points"], "place": place, "link": card_url(pair["tnr"], pair["snr"]),
                    "finished": True})
                pair["result"] = "announced"
            elif verdict == "not_played":
                pair["result"] = "not_played"
            else:
                hours = retry_hours if verdict in ("pending", "recent") else retry_other
                pair["result_next"] = (clock() + timedelta(hours=hours)).isoformat(timespec="seconds")
                if verdict == "recent":  # the tournament page changed lately: true for every player in it
                    for other in data["pairs"].values():
                        if other["tnr"] == pair["tnr"] and not other.get("result") and other is not pair:
                            other["result_next"] = pair["result_next"]
            k += 1
            print(f"result check #{k}: {verdict}")
            log.note(f"RESULT | {pair['fide']} | {player['label'] or '-'} | tnr {pair['tnr']} snr {pair['snr']} | "
                     f"{verdict} (try {pair['result_tries']})")

    # Items found in this run: give them publication times one second apart, ordered by tournament date,
    # so that readers show the latest tournament on top.
    new_items = sorted((i for i in data["feed"] if i["id"] not in before_ids),
                       key=lambda i: (i["start"], i["end"], i["tnr"], i["fide"]))
    for idx, it in enumerate(new_items):
        it["published"] = (now - timedelta(seconds=len(new_items) - 1 - idx)).isoformat()

    prune(data, order, today, finished_days, cfg.getint("feed", "max_items"))

    # ------------------------------------------------ outputs
    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    (out / "index.html").write_text(BLANK_PAGE, encoding="utf-8")
    items = [i for i in data["feed"] if i["fide"] in order]
    new_files = set()
    for gid, g in groups.items():  # one feed per group
        members = {p["fide"] for p in players if gid in p["groups"]}
        mine = [i for i in items if i["fide"] in members]
        fname = f"{g['file']}.xml"
        (out / fname).write_text(render_feed(
            mine, g["title"], f"Нови уписи играча на турнире: {g['title']} (chess-results.com)",
            base, fname, hub_list), encoding="utf-8")
        if any(i["id"] not in before_ids for i in mine):
            new_files.add(fname)
    for p in players:  # one feed per player
        mine = [i for i in items if i["fide"] == p["fide"]]
        fname = f"{prefix}-{p['fide']}.xml"
        ptitle = player_title(cfg, p)
        (out / fname).write_text(render_feed(
            mine, ptitle, f"Нови уписи: {ptitle} (chess-results.com)", base, fname, hub_list), encoding="utf-8")
        if any(i["id"] not in before_ids for i in mine):
            new_files.add(fname)
    ping = [base + f for f in sorted(new_files)] if new_items else []

    state_dir.mkdir(parents=True, exist_ok=True)
    (state_dir / "data.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    (state_dir / "ping_urls.txt").write_text("\n".join(ping) + ("\n" if ping else ""), encoding="utf-8")
    log.note(f"RUN | {ok} ok, {failures} failed, {len(new_items)} new feed items"
             + (f", ABORTED: {aborted}" if aborted else ""))
    log.save(log_hours)

    print(f"Done: {ok} ok, {failures} failed, {len(new_items)} new feed items{', ABORTED' if aborted else ''}.")
    print("Feeds: " + ", ".join(f"{base}{g['file']}.xml" for g in groups.values()))
    if aborted or (ok == 0 and not args.rebuild):
        sys.exit(2)


if __name__ == "__main__":
    main()
