"""
Term-IV timetable (Google Drive .xlsx) -> one personal .ics calendar per person.

Run by GitHub Actions every 3 hours; output goes to ./public and is published
on GitHub Pages. Anyone subscribes with:
    webcal://<username>.github.io/<repo>/<slug>.ics

Edit config.json to add people or change courses. No Google login needed:
the timetable is shared "anyone with the link".

Local test with a file you already have:
    python build_ics.py path/to/timetable.xlsx
"""
from __future__ import annotations

import hashlib
import io
import json
import re
import sys
import urllib.request
from datetime import date, datetime, timezone
from pathlib import Path

import openpyxl

ROOT = Path(__file__).parent
OUT = ROOT / "public"
MONTHS = {m: i + 1 for i, m in enumerate(
    "jan feb mar apr may jun jul aug sep oct nov dec".split())}
SESSION_RE = re.compile(
    r"^([A-Za-z][A-Za-z0-9]*?)\s*[- ]\s*S\s*-?\s*(\d+)\s*(?:\(([^)]+)\))?")
RANGE_RE = re.compile(
    r"(\d{1,2})[.:](\d{2})\s*(am|pm)?\s*-\s*(\d{1,2})[.:](\d{2})\s*(am|pm)?", re.I)


# ───────────────────────── fetch ─────────────────────────
def download_xlsx(file_id: str) -> bytes:
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        data = r.read()
    if not data.startswith(b"PK"):
        raise RuntimeError("Download did not return an .xlsx file - is the "
                           "timetable still shared 'anyone with the link'?")
    return data


def read_rows(xlsx: bytes) -> list[list[str]]:
    ws = openpyxl.load_workbook(io.BytesIO(xlsx), data_only=True).worksheets[0]
    rows = []
    for row in ws.iter_rows(values_only=True):
        cells = []
        for v in row:
            if v is None:
                cells.append("")
            elif isinstance(v, datetime):
                cells.append(v)          # keep real dates as dates
            elif isinstance(v, float) and v.is_integer():
                cells.append(str(int(v)))
            else:
                cells.append(str(v).strip())
        rows.append(cells)
    return rows


# ───────────────────────── parse ─────────────────────────
def parse_range(text: str, afternoon_default: bool = False):
    m = RANGE_RE.search(text or "")
    if not m:
        return None
    ap2 = (m.group(6) or "").lower()
    ap1 = (m.group(3) or ap2).lower()

    def to24(h, mm, ap):
        h = int(h)
        if ap == "pm" and h < 12:
            h += 12
        if ap == "am" and h == 12:
            h = 0
        if not ap and afternoon_default and h < 8:
            h += 12
        return f"{h:02d}:{mm}"
    return to24(m.group(1), m.group(2), ap1), to24(m.group(4), m.group(5), ap2)


def parse_timetable(rows, person, cfg):
    s = lambda v: v if isinstance(v, str) else ""
    venue, start = "IIM Udaipur", cfg["fallback_start"]
    for r in rows[:6]:
        for c in map(s, r):
            v = re.search(r"Venue\s*:\s*(.+)$", c, re.I)
            if v:
                venue = "IIM Udaipur, " + v.group(1).strip()
            d = re.search(r"Timetable\s+([A-Za-z]{3})[a-z]*\s+(\d{1,2}),\s*(\d{4})", c, re.I)
            if d:
                start = f"{d.group(3)}-{MONTHS[d.group(1).lower()]:02d}-{int(d.group(2)):02d}"

    hdr = next(i for i, r in enumerate(rows)
               if any(s(c) == "Date" for c in r) and any(s(c).startswith("Session") for c in r))
    date_col = [s(c) for c in rows[hdr]].index("Date")
    slots = []
    for i, c in enumerate(rows[hdr]):
        if s(c).startswith("Session"):
            t = parse_range(s(rows[hdr + 1][i]))
            if t:
                slots.append({"col": i, "name": s(c), "start": t[0], "end": t[1]})

    courses = {}
    for r_i, r in enumerate(rows):
        cells = [s(c) for c in r]
        if "Course Name" in cells:
            ci, code_c, fac_c = cells.index("Course Name"), cells.index("Code"), cells.index("Faculty")
            for k in range(r_i + 1, len(rows)):
                row = [s(c) for c in rows[k]]
                if not row[ci]:
                    break
                courses[row[code_c].upper()] = {"name": row[ci].strip(), "faculty": row[fac_c].strip()}
            break

    mine = {c.upper() for c in person["courses"]}
    section = person["section"].upper()
    events, cur, prev_day = [], None, None

    for r in rows[hdr + 2:]:
        raw = r[date_col] if date_col < len(r) else ""
        if isinstance(raw, datetime):
            d = raw.date()
        else:
            if not s(raw).isdigit():
                continue
            day = int(raw)
            if cur is None:
                y, m, dd = map(int, start.split("-"))
                cur = date(y, m, dd)
            elif day < prev_day:
                cur = date(cur.year + (cur.month == 12), cur.month % 12 + 1, day)
            else:
                cur = cur.replace(day=day)
            prev_day, d = day, cur

        filled = [sl for sl in slots if s(r[sl["col"]]) and s(r[sl["col"]]) != "--"]
        for sl in filled:
            cell = s(r[sl["col"]])
            m = SESSION_RE.match(cell)
            code = m.group(1).upper() if m else None

            if code and code in courses:
                if code not in mine:
                    continue
                sec = (m.group(3) or "").upper()
                if sec in ("DEM", "GSCM") and sec != section:
                    continue
                c = courses[code]
                events.append({
                    "key": f"{d}|{sl['start']}|{code}|S{m.group(2)}",
                    "date": d, "start": sl["start"], "end": sl["end"], "all_day": False,
                    "title": f"{m.group(1)} S{m.group(2)} · {c['name']}",
                    "location": venue,
                    "description": f"{m.group(1)} Session {m.group(2)}\nFaculty: {c['faculty']}\n"
                                   f"Slot: {sl['name']} ({sl['start']}–{sl['end']})",
                })
                continue
            if code and re.match(r"^(CS|ENTR|B2BM|DSPAE|EAT|NIS|OE|OME|SCF)", code):
                continue
            if not cfg.get("include_events", True):
                continue
            explicit = parse_range(cell, afternoon_default=True)
            name = re.sub(r"\s*\d{1,2}[.:]\d{2}\s*(am|pm)?\s*-\s*\d{1,2}[.:]\d{2}\s*(am|pm)?\s*$",
                          "", cell, flags=re.I).strip()
            all_day = not explicit and len(filled) == 1 and sl is slots[0]
            st, en = explicit if explicit else (sl["start"], sl["end"])
            events.append({
                "key": f"{d}|{'allday' if all_day else st}|EVT|{name}",
                "date": d, "start": st, "end": en, "all_day": all_day,
                "title": name, "location": venue,
                "description": "From the Term-IV timetable",
            })
    return events


# ───────────────────────── ics ─────────────────────────
def esc(t: str) -> str:
    return t.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def fold(line: str) -> str:
    out, b = [], line.encode()
    while len(b) > 75:
        cut = 75 if not out else 74
        while (b[cut] & 0xC0) == 0x80:      # don't split a UTF-8 character
            cut -= 1
        out.append(b[:cut].decode())
        b = b[cut:]
    out.append(b.decode())
    return "\r\n ".join(out)


def to_ics(events, person, cfg) -> str:
    tz = cfg["timezone"]
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    L = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//term4-calendar//IIMU//EN",
         "CALSCALE:GREGORIAN", "METHOD:PUBLISH",
         f"X-WR-CALNAME:{esc(person['calendar_name'])}", f"X-WR-TIMEZONE:{tz}",
         "REFRESH-INTERVAL;VALUE=DURATION:PT3H", "X-PUBLISHED-TTL:PT3H",
         "BEGIN:VTIMEZONE", f"TZID:{tz}", "BEGIN:STANDARD", "DTSTART:19700101T000000",
         "TZOFFSETFROM:+0530", "TZOFFSETTO:+0530", "TZNAME:IST", "END:STANDARD", "END:VTIMEZONE"]
    for e in events:
        uid = hashlib.sha1(f"{person['slug']}|{e['key']}".encode()).hexdigest()[:20]
        ymd = e["date"].strftime("%Y%m%d")
        L += ["BEGIN:VEVENT", f"UID:{uid}@term4-calendar", f"DTSTAMP:{stamp}"]
        if e["all_day"]:
            nxt = date.fromordinal(e["date"].toordinal() + 1).strftime("%Y%m%d")
            L += [f"DTSTART;VALUE=DATE:{ymd}", f"DTEND;VALUE=DATE:{nxt}", "TRANSP:TRANSPARENT"]
        else:
            L += [f"DTSTART;TZID={tz}:{ymd}T{e['start'].replace(':', '')}00",
                  f"DTEND;TZID={tz}:{ymd}T{e['end'].replace(':', '')}00"]
        L += [f"SUMMARY:{esc(e['title'])}", f"LOCATION:{esc(e['location'])}",
              f"DESCRIPTION:{esc(e['description'])}"]
        if not e["all_day"] and cfg.get("reminder_minutes") is not None:
            L += ["BEGIN:VALARM", "ACTION:DISPLAY", f"DESCRIPTION:{esc(e['title'])}",
                  f"TRIGGER:-PT{cfg['reminder_minutes']}M", "END:VALARM"]
        L.append("END:VEVENT")
    L.append("END:VCALENDAR")
    return "\r\n".join(fold(x) for x in L) + "\r\n"


def index_html(people, counts) -> str:
    items = "\n".join(
        f'<li><b>{p["calendar_name"]}</b> ({counts[p["slug"]]} events) — '
        f'<a class="w" data-f="{p["slug"]}.ics" href="{p["slug"]}.ics">{p["slug"]}.ics</a></li>'
        for p in people)
    updated = datetime.now(timezone.utc).astimezone().strftime("%d %b %Y %H:%M UTC")
    return f"""<!doctype html><meta charset="utf-8"><title>IIMU Term-IV calendars</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>body{{font:16px system-ui;max-width:680px;margin:40px auto;padding:0 16px;line-height:1.6}}
code{{background:#f2f2f2;padding:2px 6px;border-radius:4px;word-break:break-all}}</style>
<h1>IIMU Term-IV calendars</h1>
<p>Google Calendar → <b>Other calendars +</b> → <b>From URL</b> → paste your link.</p>
<ul>{items}</ul>
<p><small>Last built {updated}. Rebuilt every 3 hours from the official timetable.</small></p>
<script>document.querySelectorAll('a.w').forEach(a=>{{const u=new URL(a.dataset.f,location.href).href;
const c=document.createElement('code');c.textContent=u;a.after(' ',c);}});</script>
"""


def main():
    cfg = json.loads((ROOT / "config.json").read_text())
    xlsx = Path(sys.argv[1]).read_bytes() if len(sys.argv) > 1 else download_xlsx(cfg["source_file_id"])
    rows = read_rows(xlsx)
    OUT.mkdir(exist_ok=True)
    counts = {}
    for p in cfg["people"]:
        events = parse_timetable(rows, p, cfg)
        if len(events) < 10:
            raise RuntimeError(f"Only {len(events)} events for {p['slug']} - timetable layout may have changed.")
        (OUT / f"{p['slug']}.ics").write_text(to_ics(events, p, cfg), encoding="utf-8", newline="")
        counts[p["slug"]] = len(events)
        print(f"{p['slug']}: {len(events)} events")
    (OUT / "index.html").write_text(index_html(cfg["people"], counts), encoding="utf-8")


if __name__ == "__main__":
    main()
