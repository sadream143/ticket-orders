"""
Ticket order screenshot -> Excel (orders.xlsx)

Run:  open this file in VS Code and click the Run (play) button at the top right,
      or in a terminal:  python app.py

Everything runs on your PC. No data is sent to the cloud.
"""

import datetime
import difflib
import hashlib
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import urllib.request
from urllib.parse import urlparse

# ---------------------------------------------------------------- settings
APP_NAME = "Ticket Orders"
APP_VERSION = "1.0.0"  # change this number for every new release
GITHUB_REPO = "sadream143/ticket-orders"  # where updates are published
UPDATE_MANIFEST_URL = "https://github.com/%s/releases/latest/download/version.json" % GITHUB_REPO

FROZEN = bool(getattr(sys, "frozen", False))  # True when running as the installed .exe
APP_DIR = os.path.dirname(sys.executable) if FROZEN else os.path.dirname(os.path.abspath(__file__))
BASE_DIR = APP_DIR


def _documents_dir():
    if os.name == "nt":
        try:
            import ctypes

            buf = ctypes.create_unicode_buffer(260)
            ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf)  # 5 = My Documents
            if buf.value:
                return buf.value
        except Exception:  # noqa: BLE001
            pass
    return os.path.join(os.path.expanduser("~"), "Documents")


# Installed app: data lives in Documents\Ticket Orders (never touched by updates/uninstall).
# Running from source: everything stays next to this file.
DATA_DIR = os.path.join(_documents_dir(), "Ticket Orders") if FROZEN else APP_DIR
EXCEL_FILE = os.path.join(DATA_DIR, "orders.xlsx")
LOG_DIR = os.path.join(DATA_DIR, "logs")
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
UPDATE_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(), "TicketOrders", "updates"
)

_DEV_TESSERACT = r"C:\Users\noexp\AppData\Local\Tesseract-OCR\tesseract.exe"
_BUNDLED_TESSERACT = os.path.join(APP_DIR, "tesseract", "tesseract.exe")
_BUNDLED_TESSDATA = os.path.join(APP_DIR, "tesseract", "tessdata")


def _find_tesseract():
    if os.path.exists(_BUNDLED_TESSERACT):
        return _BUNDLED_TESSERACT
    if os.path.exists(_DEV_TESSERACT):
        return _DEV_TESSERACT
    return shutil.which("tesseract") or _DEV_TESSERACT


TESSERACT_EXE = _find_tesseract()
LOGGER = logging.getLogger("ticket_orders")

COLUMNS = [
    "Event Name",
    "Event Date & Time",
    "Month/Date/Year",
    "Time",
    "Venue (Location)",
    "Quantity",
    "Order Number",
    "Email",
    "Total Cost",
    "Seat details",
    "Sec",
    "Row",
    "Low seat",
    "Last 4 digit of card",
    "Ticket Information",
]

MONTHS = ["jan", "feb", "mar", "apr", "may", "jun",
          "jul", "aug", "sep", "oct", "nov", "dec"]

# ---------------------------------------------------------------- regex
DATE_RE = re.compile(
    r"([A-Za-z]{3,9})\.?\s+(\d{1,2})[\s,._]*(\d{4})"
    r"[\s,.\-\u2013\u2014\u00b7\u2022|~]*?"
    r"(\d{1,2}\s*[:.;]?\s*\d{2})\s*([AaPpDd9])\.?\s*[Mm]"
)
SEAT_RE = re.compile(
    r"\bS[eo][ce]\s+([A-Za-z0-9]+)[\s,.\-\u2013\u2014:;\u00b7\u2022|+*~]+Row\s+([A-Za-z0-9]+)"
    r"[\s,.\-\u2013\u2014:;\u00b7\u2022|+*~]+Se[ao]ts?\s+(\d+)(?:\s*[-\u2013\u2014~]\s*(\d+))?",
    re.I,
)
SEC_ONLY_RE = re.compile(r"\bSec\s+([A-Za-z0-9]{1,8})\b")
DATE_ONLY_RE = re.compile(r"([A-Za-z]{3,9})\.?\s+(\d{1,2})[\s,._]*(20\d{2})\b")
QTY_RE = re.compile(r"(?<![A-Za-z0-9])[xX]\s*(\d{1,3})(?!\d)")
CARD_RE = re.compile(
    r"(American\s*Express|Am[a-z]*\s*Ex[a-z]*|Vi[a-z]{1,2}|Master\s*card|Mastercard"
    r"|Discover|Diners)\s*[-\u2013\u2014:]?\s*(\d{4})\b",
    re.I,
)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
ORDER_RE = re.compile(
    r"Order\s*(?:#|No\.?|Number)?\s*[:=]?\s*#?\s*(\d{2}-\d{4,6}/[A-Z0-9]{3})", re.I
)
ORDER_OLD_RE = re.compile(
    r"Order\s*(?:#|No\.?|Number)?\s*[:=]?\s*#?\s*(\d{2}-\d{4,6}/[A-Z][A-Z0-9]{1,3})", re.I
)
ORDER_ANY_RE = re.compile(r"\b(\d{2}-\d{4,6}/[A-Z0-9]{3})\b")
INFO_LINE_RE = re.compile(
    r"((?:Seats?\s+\d+(?:\s*[-\u2013\u2014]\s*\d+)?|Sec\s+[A-Za-z0-9]+)\s*:\s*[A-Za-z].+)$"
)

# text from the right-hand side of the page that sometimes gets mixed in
NOISE = [
    r"The\s*Perfect\s*Tic\w*\s*Prot\w*",
    r"Learn\s*mo\w*\.?",
    r"Life\s*hap\w*.*",
    r"peace\s*of\s*mi\w*\.?",
    r"Need to Spl.*",
    r"Spl\w*\s*on\b.*",
    r"\bvenmo\b.*",
    r"Use the Venmo.*",
    r"Order\s*det\w*\s*em\w*.*",
    r"\bticketmaster\b.*",
    r"THINGS\s*TO\s*DO.*",
    r"Find\s*tickets.*",
    r"Transfer Your Tickets.*",
    r"\bMy Tickets\b.*",
    r"Manage.*",
]
KNOWN_DOMAINS = ["gmail.com", "hotmail.com", "icloud.com", "yahoo.com",
                 "outlook.com", "aol.com", "me.com", "live.com", "msn.com"]

# ---------------------------------------------------------------- currency
CODE_MAP = {
    "USD": "USD", "US": "USD", "CAD": "CAD", "CA": "CAD", "MXN": "MXN", "MX": "MXN",
    "MEX": "MXN", "EUR": "EUR", "GBP": "GBP", "GB": "GBP", "UK": "GBP", "AUD": "AUD",
    "AU": "AUD",
}
CODE_WORDS = "USD|CAD|MXN|EUR|GBP|AUD|MEX|CA|US|MX|GB|UK|AU"
CUR_SHOW = {"USD": "$", "CAD": "CA$", "MXN": "MX$", "EUR": "\u20ac", "GBP": "\u00a3", "AUD": "AU$"}
CUR_FORMAT = {
    "USD": '"$"#,##0.00', "CAD": '"CA$"#,##0.00', "MXN": '"MX$"#,##0.00',
    "EUR": '"\u20ac"#,##0.00', "GBP": '"\u00a3"#,##0.00', "AUD": '"AU$"#,##0.00',
}
AMOUNT_RE = re.compile(r"(\d{1,3}(?:[.,\s]\d{3})+[.,]\d{2}|\d+[.,]\d{2})")
LABEL_RE = re.compile(
    r"(?<![A-Za-z])(?:TOTAL(?:E)?(?:\s+PAID)?|GESAMT\w*|IMPORTE\s+TOTAL)(?![A-Za-z])", re.I
)


def _to_amount(txt):
    """'1,251.68' or '1.251,68' or '1 251,68' -> 1251.68"""
    m = re.search(r"[.,](\d{2})$", txt)
    if not m:
        return None
    ip = re.sub(r"\D", "", txt[:m.start()])
    return float(ip + "." + m.group(1)) if ip else None


def money_from_text(rest, whole=""):
    """Extract (currency, amount) from text. Defaults to USD when no currency is found."""
    m = AMOUNT_RE.search(rest)
    if not m:
        return None
    amt = _to_amount(m.group(1))
    if amt is None:
        return None
    before, after = rest[:m.start()], rest[m.end():m.end() + 8]
    cur = None
    for seg in (before, after):
        cm = re.search(r"(?<![A-Za-z])(%s)(?![A-Za-z])" % CODE_WORDS, seg, re.I)
        if cm:
            cur = CODE_MAP[cm.group(1).upper()]
            break
    if not cur:
        both = before + after
        if "\u20ac" in both:
            cur = "EUR"
        elif "\u00a3" in both:
            cur = "GBP"
    if not cur and whole:  # no currency on the total line: infer it from other lines
        found = set()
        for cm in re.finditer(r"(?<![A-Za-z])(CAD|MXN|EUR|GBP|AUD|MEX)(?![A-Za-z])", whole):
            found.add(CODE_MAP[cm.group(1)])
        for cm in re.finditer(r"(?<![A-Za-z])(CA|MX|AU)\s*\$", whole):
            found.add(CODE_MAP[cm.group(1)])
        if len(found) == 1:
            cur = found.pop()
    return (cur or "USD", amt)


def format_money(cur, amt):
    return CUR_SHOW[cur] + "{:,.2f}".format(amt)


def _find_total(lines, text):
    for i, ln in enumerate(lines):
        m = LABEL_RE.search(ln)
        if not m:
            continue
        mon = money_from_text(ln[m.end():], text)
        if not mon and i + 1 < len(lines):
            mon = money_from_text(lines[i + 1], text)
        if mon:
            return format_money(*mon)
    return ""




# ---------------------------------------------------------------- helpers
def _fix_words(t):
    """Fix common OCR mistakes (i <-> l) in a few known words."""
    t = re.sub(r"Admi[s5]{1,2}[il1]{1,2}on", "Admission", t, flags=re.I)
    t = re.sub(r"Restr[il1]ct[il1]on", "Restriction", t, flags=re.I)
    t = re.sub(r"Stad[il1]um", "Stadium", t, flags=re.I)
    return t


def _strip_noise_tail(c):
    for p in NOISE:
        c = re.sub(r"\s*" + p, "", c, flags=re.I)
    return c.strip(" \t|*-")


def _clean_noise(line):
    for p in NOISE:
        line = re.sub(p, "", line, flags=re.I)
    line = re.sub(r"^[^A-Za-z0-9]+", "", line)
    return line.strip(" \t|*-")


def _strip_lead(line):
    """OCR often reads a small icon as 1 stray character: drop it."""
    toks = line.split()
    dropped = 0
    while toks and dropped < 2 and (
        len(toks[0]) == 1 or not any(c.isalnum() for c in toks[0])
    ):
        toks.pop(0)
        dropped += 1
    return " ".join(toks)


def _clean_title(t):
    t = re.sub(r"\s+\S*\s*venmo\W*$", "", t, flags=re.I)
    t = re.sub(r"\s+Spl\w*\s*on\W*$", "", t, flags=re.I)
    t = _strip_lead(t).strip()
    # Event titles are always uppercase; lowercase text after them belongs to another column
    toks = t.split()
    cut = len(toks)
    for i, tok in enumerate(toks):
        if re.search(r"[a-z]", tok):
            cut = i
            break
    if 2 <= cut < len(toks):
        if cut > 2 and len(toks[cut - 1]) <= 2 and (
            len(toks[cut - 1]) == 1 or any(ch.isdigit() for ch in toks[cut - 1])
        ):
            cut -= 1
        t = " ".join(toks[:cut])
    toks = t.split()
    while toks and not any(ch.isalpha() for ch in toks[-1]):
        toks.pop()
    return " ".join(toks)


def _fix_email(text):
    m = re.search(r"emailed\s+to\s+(.+)", text, re.I)
    tail = m.group(1).strip() if m else ""
    em = EMAIL_RE.search(tail) or EMAIL_RE.search(text)
    sp = re.search(r"(\S.*?)@([A-Za-z0-9.\-]+\.[A-Za-z]{2,})", tail)
    if sp and not (em and em.group(0) == sp.group(0)):
        email = sp.group(1).replace(" ", "") + "@" + sp.group(2)
    elif em:
        email = em.group(0).strip(".,;")
    else:
        email = ""
    if email:
        local, _, dom = email.partition("@")
        dom = dom.lower()
        close = difflib.get_close_matches(dom, KNOWN_DOMAINS, n=1, cutoff=0.75)
        if close and dom not in KNOWN_DOMAINS:
            dom = close[0]
        return local.lower() + "@" + dom
    if tail:  # no "@" was read: try to repair the domain
        tail = tail.split("  ")[0]
        fixed = re.sub(
            r"\s*\S{0,2}(icloud|gmail|hotmail|yahoo|outlook|aol)\.com.*$",
            r"@\1.com", tail, flags=re.I)
        return fixed.replace(" ", "")
    return ""


MONTH_ALIASES = {
    "ene": 1, "abr": 4, "ago": 8, "dic": 12, "okt": 10, "dez": 12, "janv": 1, "fev": 2,
    "fevr": 2, "avr": 4, "juil": 7, "juin": 6, "mai": 5, "mag": 5, "giu": 6, "lug": 7,
    "gen": 1, "set": 9, "ott": 10, "mrt": 3, "mei": 5, "marz": 3,
}
DAYMONTH_RE = re.compile(
    r"(?<!\d)(\d{1,2})(?:st|nd|rd|th|er)?\.?\s+(?:de\s+)?([A-Za-z\u00c0-\u00ff]{3,10})\.?,?"
    r"\s+(?:de\s+)?(20\d{2})\b"
)
TIME_RE = re.compile(r"(\d{1,2})\s*[:.h]\s*(\d{2})\s*([AaPpDd9])?\.?\s*([Mm])?\b")


def _month_number(tok):
    import unicodedata

    low = "".join(
        c for c in unicodedata.normalize("NFKD", tok.lower()) if not unicodedata.combining(c)
    )
    for k in (low[:4], low[:3]):
        if k in MONTH_ALIASES:
            return MONTH_ALIASES[k]
    key = low[:3]
    if key in MONTHS:
        return MONTHS.index(key) + 1
    m = difflib.get_close_matches(key, MONTHS, n=1, cutoff=0.6)
    return MONTHS.index(m[0]) + 1 if m else None


def _time_12h(hh, mm, ap, m_flag):
    """Convert 12-hour or 24-hour time to 'h:mm AM/PM'."""
    hh = int(hh)
    if ap and m_flag:
        return "%d:%s %sM" % (hh, mm, "A" if ap.upper() == "A" else "P")
    if hh == 0:
        return "12:%s AM" % mm
    if hh < 12:
        return "%d:%s AM" % (hh, mm)
    return "%d:%s PM" % (hh - 12 if hh > 12 else 12, mm)


def _parse_daymonth(text_lines):
    for i, ln in enumerate(text_lines):
        for dm in DAYMONTH_RE.finditer(ln):
            mon = _month_number(dm.group(2))
            day, year = int(dm.group(1)), int(dm.group(3))
            if not mon or not 1 <= day <= 31:
                continue
            tm = TIME_RE.search(ln[dm.end():dm.end() + 25])
            time_s = _time_12h(*tm.groups()[:2], tm.group(3), tm.group(4)) if tm else ""
            mdy = "%02d/%02d/%d" % (mon, day, year)
            pretty = "%s %d, %d" % (MONTHS[mon - 1].title(), day, year)
            try:
                wd = datetime.date(year, mon, day).strftime("%a")
                pretty = "%s \u00b7 %s" % (wd, pretty)
            except ValueError:
                pass
            combo = pretty + (" \u00b7 " + time_s if time_s else "")
            return i, mdy, time_s, combo
    return None


def _parse_date(text_lines, full_text):
    """returns (line_index or None, month/date/year, time, date&time)"""
    found = None
    for i, ln in enumerate(text_lines):
        m = DATE_RE.search(ln)
        if m and _month_number(m.group(1)):
            found = (i, m)
            break
    if not found:
        m = DATE_RE.search(full_text)
        if m and _month_number(m.group(1)):
            found = (None, m)
    if not found:  # day-month order (UK/Europe/Mexico), 24h time, other languages
        r = _parse_daymonth(text_lines)
        if r:
            return r
    if not found:  # time could not be read: keep the date only
        for i, ln in enumerate(text_lines):
            for dm in DATE_ONLY_RE.finditer(ln):
                mon = _month_number(dm.group(1))
                if mon and 1 <= int(dm.group(2)) <= 31:
                    day, year = int(dm.group(2)), int(dm.group(3))
                    mdy = "%02d/%02d/%d" % (mon, day, year)
                    pretty = "%s %d, %d" % (MONTHS[mon - 1].title(), day, year)
                    try:
                        wd = datetime.date(year, mon, day).strftime("%a")
                        pretty = "%s \u00b7 %s" % (wd, pretty)
                    except ValueError:
                        pass
                    return i, mdy, "", pretty
        return None, "", "", ""
    idx, m = found
    mon = _month_number(m.group(1))
    day, year = int(m.group(2)), int(m.group(3))
    digits = re.sub(r"\D", "", m.group(4))
    if len(digits) == 3:
        hh, mm = int(digits[0]), digits[1:]
    else:
        hh, mm = int(digits[:2]), digits[2:]
    ampm = "A" if m.group(5).upper() == "A" else "P"  # D / 9 are misreads of PM
    time_s = "%d:%s %sM" % (hh, mm, ampm)
    pretty = "%s %d, %d" % (MONTHS[mon - 1].title(), day, year)  # used for the Event Date & Time column
    mdy = "%02d/%02d/%d" % (mon, day, year)  # Month/Date/Year column: MM/DD/YYYY
    try:
        wd = datetime.date(year, mon, day).strftime("%a")
        combo = "%s \u00b7 %s \u00b7 %s" % (wd, pretty, time_s)
    except ValueError:
        combo = "%s \u00b7 %s" % (pretty, time_s)
    return idx, mdy, time_s, combo


def _find_venue(lines, layout2):
    for i, ln in enumerate(lines):
        c = _strip_lead(ln)
        c = _strip_noise_tail(c)
        if not c or any(ch.isdigit() for ch in c):
            continue
        if re.search(r"venmo|split|express|visa|admission|presale", c, re.I):
            continue
        st = re.search(r",\s*[A-Z]{2,5}\b", c)
        has_state = bool(st) and (re.search(r"\s[-=~\u2013\u2014]\s", c) or layout2)
        has_long_dash = (("\u2014" in c) or ("\u2013" in c)) and \
            sum(ch.isalpha() for ch in c) >= 12
        if has_state and layout2 and st:
            c = c[:st.end()]  # drop text from other columns after the state
        if has_state or (has_long_dash and not layout2):
            c = re.sub(r"\s=\s", " - ", c)
            return i, _fix_words(c)
    return None, ""


def _prev_nonempty(lines, idx):
    j = idx - 1
    while j >= 0:
        if lines[j].strip():
            return j
        j -= 1
    return None


# ---------------------------------------------------------------- parser
def parse_order(text):
    """OCR text -> list of row dicts (one row per seat group)."""
    lines = [l.strip() for l in text.splitlines()]
    layout2 = bool(re.search(r"total\s*paid|is\s*confirmed|order\s*summary", text, re.I))

    # --- simple fields
    m = ORDER_RE.search(text) or ORDER_ANY_RE.search(text) or ORDER_OLD_RE.search(text)
    order_no = m.group(1).upper() if m else ""
    if order_no:  # OCR: S -> 5, O -> 0 (first 2 characters of the suffix)
        head, tail = order_no.rsplit("/", 1)
        fix = {"5": "S", "0": "O", "8": "B"}
        tail = "".join(fix.get(c, c) if i < 2 else c for i, c in enumerate(tail))
        order_no = head + "/" + tail

    email = _fix_email(text)

    total = _find_total(lines, text)

    m = CARD_RE.search(text)
    card4 = m.group(2) if m else ""

    date_idx, mdy, time_s, date_time = _parse_date(lines, text)

    # --- seat groups
    groups = []  # (line_idx, sec, row, lo, hi)
    for i, ln in enumerate(lines):
        for sm in SEAT_RE.finditer(ln):
            lo = int(sm.group(3))
            hi = int(sm.group(4)) if sm.group(4) else None
            groups.append((i, sm.group(1).upper(), sm.group(2).upper(), lo, hi))
    if not groups:
        seen = []
        for i, ln in enumerate(lines):
            for sm in SEC_ONLY_RE.finditer(ln):
                sec = sm.group(1).upper()
                if sec not in seen and sec != "ROW":
                    seen.append(sec)
                    groups.append((i, sec, "", None, None))

    has_assigned = any(g[3] is not None for g in groups)
    ga_flag = (not has_assigned) and bool(
        re.search(r"general\s*admi\w*", text, re.I)
        or any(re.fullmatch(r"GA\d*", g[1]) for g in groups)
    )
    if ga_flag:
        groups = [(groups[0][0] if groups else 0, "GA", "", None, None)]

    # --- venue
    venue_idx, venue = _find_venue(lines, layout2)

    if not venue and not layout2 and date_idx is not None:
        j = date_idx + 1
        while j < len(lines) and not lines[j].strip():
            j += 1
        if j < len(lines):
            c = _strip_noise_tail(_strip_lead(lines[j]))
            if c and not any(ch.isdigit() for ch in c) and len(c) > 6:
                venue_idx, venue = j, _fix_words(c)

    # --- event name
    title = ""
    if layout2 and venue_idx is not None:
        j = _prev_nonempty(lines, venue_idx)
        for _ in range(5):  # look a few lines up for the uppercase title line
            if j is None:
                break
            cand = _clean_title(lines[j])
            up = sum(ch.isupper() for ch in cand)
            if up >= 6 and len(cand.split()) >= 2 and not DATE_RE.search(lines[j]):
                title = cand
                break
            j = _prev_nonempty(lines, j)
    elif groups:
        j = _prev_nonempty(lines, groups[0][0])
        if j is not None:
            title = _clean_title(lines[j])
    if not title:
        cands = []
        for ln in lines:
            letters = [c for c in ln if c.isalpha()]
            if (len(ln) > 10 and letters
                    and sum(c.isupper() for c in letters) / len(letters) > 0.7
                    and not re.search(r"YOU GOT|YOU.?RE IN|WHAT YOU ORDERED", ln.upper())
                    and not re.search(r"LEVEL|SEATING|SECTION|MEZZANINE|FLOOR|ADMISSION|"
                                      r"ORCHESTRA|CLUB|GENERAL", ln.upper())):
                cands.append(_clean_title(ln))
        cands = [c for c in cands if c]
        if cands:
            title = next((c for c in cands if ":" in c or "TOUR" in c.upper()), cands[0])

    # --- ticket information
    info = ""
    if layout2:
        for ln in lines:
            im = INFO_LINE_RE.search(ln)
            if im:
                info = im.group(1).strip()
                break
    elif title:
        start = None
        if groups and groups[0][3] is not None:
            start = groups[-1][0] + 1
        else:
            for i, ln in enumerate(lines):
                if re.search(r"Ticket\s*Inf\w*|\bInformat\w*", ln, re.I):
                    start = i + 1
                    break
            if start is None:
                for i, ln in enumerate(lines):
                    if title and title[:12] in ln:
                        start = i + 1
                        break
        parts = []
        for ln in (lines[start:] if start is not None else []):
            if DATE_RE.search(ln) or re.match(r"\W*(Tickets|Order|Delivery)\b", ln, re.I):
                break
            if venue and venue[:10] in ln:
                break
            c = _clean_noise(ln).strip(" :;|,.-")
            if re.search(r"\b20\d{2}\b", ln):  # date line (even when OCR misread it)
                break
            if re.match(r"ticket\s*inf\w*\W*$", c, re.I):
                continue
            if re.search(r"inform", c, re.I):
                continue  # "Ticket Information ..." heading line
            first = next((ch for ch in c if ch.isalpha()), "")
            if first and not first.isupper():
                continue  # starts with lowercase = junk from another column
            if sum(ch.isalpha() for ch in c) < 3:
                continue
            if max((len(re.sub(r"[^A-Za-z]", "", w)) for w in c.split()), default=0) < 4:
                continue  # junk (scrambled OCR letters)
            if parts and parts[-1] == c:
                continue
            parts.append(_fix_words(c))
        info = "; ".join(parts)

    # --- fallback quantity (when seat numbers are not shown)
    fallback_qty = None
    for ln in lines:
        if re.search(r"venmo|split", ln, re.I):
            continue
        qm = QTY_RE.search(ln)
        if qm:
            n = qm.group(1)
            if len(n) == 2 and n.endswith("0"):  # "x 2G)" is sometimes read as "x 20"
                n = n[0]
            fallback_qty = int(n)
            break

    # --- build rows
    base = {
        "Event Name": title,
        "Event Date & Time": date_time,
        "Month/Date/Year": mdy,
        "Time": time_s,
        "Venue (Location)": venue,
        "Order Number": order_no,
        "Email": email,
        "Last 4 digit of card": card4,
        "Ticket Information": info,
    }
    rows = []
    if not groups:
        groups = [(0, "", "", None, None)]
    for n, (_, sec, row, lo, hi) in enumerate(groups):
        r = dict(base)
        r["Sec"] = sec
        r["Row"] = row
        r["Low seat"] = "" if lo is None else str(lo)
        if lo is not None and hi is not None and hi >= lo:
            qty = hi - lo + 1
        elif lo is not None and hi is None:
            qty = 1
        else:
            qty = fallback_qty if len(groups) == 1 else None
        r["Quantity"] = "" if qty is None else str(qty)
        parts = []
        if sec:
            parts.append("Sec " + sec)
        if row:
            parts.append("Row " + row)
        if lo is not None:
            parts.append("Seats %d - %d" % (lo, hi) if hi else "Seat %d" % lo)
        r["Seat details"] = "General Admission" if ga_flag else ", ".join(parts)
        r["Total Cost"] = total if n == 0 else ""
        rows.append({c: r.get(c, "") for c in COLUMNS})
    return rows


# ---------------------------------------------------------------- OCR
def _tess_env():
    env = dict(os.environ)
    if os.path.isdir(_BUNDLED_TESSDATA):
        env["TESSDATA_PREFIX"] = _BUNDLED_TESSDATA
    return env


def _tesseract(path, psm):
    flags = 0x08000000 if os.name == "nt" else 0  # hide the console window
    res = subprocess.run(
        [TESSERACT_EXE, path, "stdout", "--psm", str(psm)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=flags,
        env=_tess_env(),
    )
    if res.returncode != 0 and not res.stdout.strip():
        raise RuntimeError(res.stderr.strip() or "Could not run Tesseract")
    return res.stdout


_RAPID = {"engine": None, "failed": False}


def _rapid_text(path):
    """Read with RapidOCR (offline). Silently skipped if it is not installed."""
    if _RAPID["failed"]:
        return ""
    try:
        if _RAPID["engine"] is None:
            import logging

            from rapidocr import RapidOCR

            logging.getLogger("RapidOCR").setLevel(logging.ERROR)
            _RAPID["engine"] = RapidOCR()
        out = _RAPID["engine"](path)
        if out is None or out.boxes is None or out.txts is None:
            return ""
        items = []
        for box, txt in zip(out.boxes, out.txts):
            ys = [float(p[1]) for p in box]
            xs = [float(p[0]) for p in box]
            items.append((sum(ys) / len(ys), min(xs), max(ys) - min(ys), txt))
        items.sort(key=lambda t: (t[0], t[1]))
        lines, cur, cy, ch = [], [], None, None
        for yc, x, h, txt in items:
            if cur and abs(yc - cy) <= 0.6 * max(h, ch):
                cur.append((x, txt))
            else:
                if cur:
                    lines.append(" ".join(t for _, t in sorted(cur)))
                cur, cy, ch = [(x, txt)], yc, h
        if cur:
            lines.append(" ".join(t for _, t in sorted(cur)))
        return "\n".join(lines)
    except Exception:  # noqa: BLE001
        _RAPID["failed"] = True
        return ""


def run_ocr(image_path):
    """Returns a LIST of OCR texts (several passes): RapidOCR first, then Tesseract."""
    from PIL import Image, ImageFilter, ImageOps

    first = []
    rt = _rapid_text(image_path)
    if rt:
        first.append(rt)

    img = Image.open(image_path)
    if img.width >= 1400:  # large screenshot: one pass is enough
        return first + [_tesseract(image_path, 6)]

    scale = 3 if img.width < 1000 else 2
    gray = ImageOps.grayscale(img.convert("RGB"))
    gray = gray.resize((gray.width * scale, gray.height * scale), Image.LANCZOS)
    sharp = ImageOps.autocontrast(gray).filter(ImageFilter.SHARPEN)
    tmp = os.path.join(tempfile.gettempdir(), "ticket_ocr_tmp.png")

    texts = list(first)
    for variant, psm in ((gray, 6), (sharp, 6), (sharp, 4)):
        variant.save(tmp)
        texts.append(_tesseract(tmp, psm))
    try:
        crop_email = _email_crop_pass(gray, scale, tmp)
        if crop_email:
            texts.append(EMAIL_MARK + crop_email)
    except Exception:  # noqa: BLE001  (the main result still works if this extra pass fails)
        pass
    return texts


EMAIL_MARK = "@@EMAIL@@"


def _email_crop_pass(gray, scale, tmp):
    """Crop the email line, enlarge it and read it separately (emails are the most error-prone)."""
    import csv
    import io
    from PIL import ImageOps

    gray.save(tmp)
    flags = 0x08000000 if os.name == "nt" else 0
    out = subprocess.run(
        [TESSERACT_EXE, tmp, "stdout", "--psm", "6", "tsv"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        creationflags=flags, env=_tess_env(),
    ).stdout
    words = list(csv.DictReader(io.StringIO(out), delimiter="\t", quoting=csv.QUOTE_NONE))
    hit = None
    for w in words:
        t = (w.get("text") or "").lower()
        if t and difflib.SequenceMatcher(None, t, "emailed").ratio() >= 0.7:
            hit = w
            break
    if not hit:
        return ""
    same = [w for w in words
            if w["block_num"] == hit["block_num"] and w["par_num"] == hit["par_num"]
            and w["line_num"] == hit["line_num"]]
    k = same.index(hit)
    after = same[k + 2:]
    if not after:
        return ""
    x0 = int(after[0]["left"]) - 6
    x1 = max(int(w["left"]) + int(w["width"]) for w in same) + 8
    top, h = int(hit["top"]), int(hit["height"])
    crop = gray.crop((max(0, x0), max(0, top - 8), min(gray.width, x1), top + h + 8))
    f = max(2, int(round(9.0 / scale)))
    crop = ImageOps.autocontrast(crop.resize((crop.width * f, crop.height * f)))
    crop = ImageOps.expand(crop, border=30, fill=255)
    crop.save(tmp)
    res = subprocess.run(
        [TESSERACT_EXE, tmp, "stdout", "--psm", "7", "-c",
         "tessedit_char_whitelist=abcdefghijklmnopqrstuvwxyz0123456789@._-"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        creationflags=flags, env=_tess_env(),
    ).stdout.strip()
    for tok in res.split():
        if "@" in tok:
            return tok.strip(".,;")
    return ""


EVENT_COLS = [
    "Event Name", "Event Date & Time", "Month/Date/Year", "Time",
    "Venue (Location)", "Order Number", "Email", "Total Cost",
    "Last 4 digit of card", "Ticket Information",
]


def _valid(col, val):
    if col == "Email":
        return bool(EMAIL_RE.fullmatch(val or ""))
    if col == "Order Number":
        return bool(re.fullmatch(r"\d{2}-\d{4,6}/[A-Z0-9]{3}", val or ""))
    return bool(val)


def _score(rows):
    r = rows[0]
    s = sum(2 for c in ("Email", "Order Number") if _valid(c, r[c]))
    s += sum(1 for c in ("Total Cost", "Last 4 digit of card", "Month/Date/Year",
                         "Sec", "Venue (Location)", "Event Name") if r[c])
    return s


def _vote_title(results, preferred=""):
    titles = [r[0]["Event Name"] for r in results if r[0]["Event Name"]]
    if not titles:
        return ""
    counts = [
        sum(1 for u in titles if difflib.SequenceMatcher(None, t, u).ratio() >= 0.85)
        for t in titles
    ]
    top = max(counts)
    leaders = [t for t, c in zip(titles, counts) if c == top]
    if preferred in leaders:
        return preferred
    return max(leaders, key=len)


def parse_best(texts):
    """Pick the best result from several OCR texts."""
    if isinstance(texts, str):
        texts = [texts]
    crop_email = ""
    plain = []
    for t in texts:
        if t.startswith(EMAIL_MARK):
            crop_email = _fix_email("emailed to " + t[len(EMAIL_MARK):])
        else:
            plain.append(t)
    results = [parse_order(t) for t in plain]
    best = max(results, key=_score)  # on a tie the first pass wins

    # seat not readable? use another pass if it has one
    if all(not r["Sec"] for r in best):
        for other in results:
            if other is not best and other[0]["Sec"]:
                keep = {c: best[0][c] for c in EVENT_COLS if _valid(c, best[0][c])}
                best = [dict(r) for r in other]
                for r in best:
                    r.update(keep)
                break

    # date + time: majority vote across passes (first pass wins ties)
    combos = [
        (o[0]["Month/Date/Year"], o[0]["Time"], o[0]["Event Date & Time"])
        for o in results if o[0]["Month/Date/Year"] and o[0]["Time"]
    ]
    if combos:
        keys = [(a, b) for a, b, _ in combos]
        win = max(keys, key=lambda k: keys.count(k))
        pick = next(c for c in combos if (c[0], c[1]) == win)
        for r in best:
            r["Month/Date/Year"], r["Time"], r["Event Date & Time"] = pick
    elif not best[0]["Month/Date/Year"]:
        for o in results:
            if o[0]["Month/Date/Year"]:
                for r in best:
                    r["Month/Date/Year"] = o[0]["Month/Date/Year"]
                    r["Event Date & Time"] = o[0]["Event Date & Time"]
                break

    for col in EVENT_COLS:
        if _valid(col, best[0][col]):
            continue
        for other in results:
            if _valid(col, other[0][col]):
                for r in best:
                    r[col] = other[0][col]
                break

    title = _vote_title(results, best[0]["Event Name"])  # the title most passes agree on
    if title:
        for r in best:
            r["Event Name"] = title

    mails = [o[0]["Email"] for o in results if EMAIL_RE.fullmatch(o[0]["Email"] or "")]
    if results and EMAIL_RE.fullmatch(results[0][0]["Email"] or ""):
        mails.append(results[0][0]["Email"])  # the first pass (RapidOCR) counts twice
    if crop_email and EMAIL_RE.fullmatch(crop_email):
        mails.append(crop_email)
    if mails:  # treat look-alike characters (i/l/1, s/z) as equal and pick the majority
        def look(e):
            e = e.lower()
            e = re.sub(r"[l1|!]", "i", e)
            return e.replace("0", "o").replace("z", "s")

        keys = [look(m) for m in mails]
        top = max(keys, key=lambda k: keys.count(k))
        group = [m for m, k in zip(mails, keys) if k == top]
        pick = max(group, key=lambda m: group.count(m))
        for r in best:
            r["Email"] = pick
    return best


# ---------------------------------------------------------------- Excel
def _convert(col, val):
    val = (val or "").strip()
    if val == "":
        return None
    if col == "Total Cost":
        mon = money_from_text(val)
        return mon[1] if mon else val
    if col in ("Quantity", "Low seat") and val.isdigit():
        return int(val)
    return val


def _backup_excel(keep=20):
    """Copy the current orders.xlsx to the backups folder before every save."""
    if not os.path.exists(EXCEL_FILE):
        return
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    shutil.copy2(EXCEL_FILE, os.path.join(BACKUP_DIR, "orders-%s.xlsx" % stamp))
    old = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith("orders-"))
    for f in old[:-keep]:
        try:
            os.remove(os.path.join(BACKUP_DIR, f))
        except OSError:
            pass


def save_to_excel(rows, ask_duplicate):
    """rows: list of dict. ask_duplicate(list_of_keys)->bool. returns added count"""
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font

    os.makedirs(DATA_DIR, exist_ok=True)
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Orders"
        ws.append(COLUMNS)
        for c in ws[1]:
            c.font = Font(bold=True)
        widths = [38, 30, 16, 10, 36, 9, 16, 30, 12, 30, 8, 7, 9, 12, 50]
        for i, w in enumerate(widths):
            ws.column_dimensions[chr(65 + i)].width = w
        ws.freeze_panes = "A2"

    oi, si, ri = (COLUMNS.index(n) for n in ("Order Number", "Sec", "Row"))
    existing = set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        existing.add((str(r[oi] or ""), str(r[si] or ""), str(r[ri] or "")))

    dups = [
        (r["Order Number"], r["Sec"], r["Row"])
        for r in rows
        if (r["Order Number"], r["Sec"], r["Row"]) in existing
    ]
    skip_dups = bool(dups) and not ask_duplicate(dups)

    added = 0
    for r in rows:
        key = (r["Order Number"], r["Sec"], r["Row"])
        if skip_dups and key in existing:
            continue
        ws.append([_convert(c, r.get(c, "")) for c in COLUMNS])
        mon = money_from_text(r.get("Total Cost", "") or "")
        if mon:
            tc = COLUMNS.index("Total Cost") + 1
            ws.cell(row=ws.max_row, column=tc).number_format = CUR_FORMAT[mon[0]]
        added += 1
    _backup_excel()
    wb.save(EXCEL_FILE)
    LOGGER.info("Saved %d row(s) to %s", added, EXCEL_FILE)
    return added


# ---------------------------------------------------------------- logging
def setup_logging():
    """Write a log file to Documents\\Ticket Orders\\logs (no order data is logged)."""
    from logging.handlers import RotatingFileHandler

    os.makedirs(LOG_DIR, exist_ok=True)
    handler = RotatingFileHandler(
        os.path.join(LOG_DIR, "app.log"), maxBytes=1_000_000, backupCount=3, encoding="utf-8"
    )
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    LOGGER.setLevel(logging.INFO)
    if not LOGGER.handlers:
        LOGGER.addHandler(handler)
    if FROZEN:  # a windowed .exe has no console: send stdout/stderr to a file
        try:
            console = open(os.path.join(LOG_DIR, "console.log"), "a", encoding="utf-8", buffering=1)
            sys.stdout = sys.stderr = console
        except OSError:
            pass

    def _hook(exc_type, exc, tb):
        LOGGER.error("Uncaught exception", exc_info=(exc_type, exc, tb))

    sys.excepthook = _hook
    LOGGER.info("Starting %s v%s (frozen=%s)", APP_NAME, APP_VERSION, FROZEN)


_MUTEX = []  # keeps the handle alive


def acquire_single_instance():
    """Only one copy of the app may run at a time (Windows)."""
    if os.name != "nt":
        return True
    import ctypes

    handle = ctypes.windll.kernel32.CreateMutexW(None, False, "TicketOrdersAppMutex")
    _MUTEX.append(handle)
    return ctypes.windll.kernel32.GetLastError() != 183  # 183 = already exists


# ---------------------------------------------------------------- updates
def _ver_tuple(v):
    return tuple(int(x) for x in re.findall(r"\d+", str(v))[:4])


def _valid_manifest(m):
    try:
        url = m["url"]
        ok_url = url.startswith("https://github.com/%s/releases/download/" % GITHUB_REPO)
        ok_name = re.fullmatch(r"TicketOrders-Setup-[\d.]+\.exe", os.path.basename(urlparse(url).path))
        return bool(_ver_tuple(m["version"])) and ok_url and bool(ok_name)
    except (KeyError, TypeError, AttributeError):
        return False


def check_for_update(manifest_url=None, timeout=6):
    """Returns the manifest dict if a newer version exists, otherwise None.
    Raises OSError/ValueError when the check itself fails (offline etc.)."""
    req = urllib.request.Request(
        manifest_url or UPDATE_MANIFEST_URL, headers={"User-Agent": "TicketOrders/" + APP_VERSION}
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        manifest = json.loads(resp.read(200_000).decode("utf-8"))
    if not _valid_manifest(manifest):
        raise ValueError("Invalid update manifest")
    if _ver_tuple(manifest["version"]) > _ver_tuple(APP_VERSION):
        return manifest
    return None


def download_update(manifest, progress=None):
    """Download the installer, verify its SHA-256 and return the file path."""
    os.makedirs(UPDATE_DIR, exist_ok=True)
    dest = os.path.join(UPDATE_DIR, os.path.basename(urlparse(manifest["url"]).path))
    req = urllib.request.Request(manifest["url"], headers={"User-Agent": "TicketOrders/" + APP_VERSION})
    digest = hashlib.sha256()
    with urllib.request.urlopen(req, timeout=30) as resp, open(dest, "wb") as out:
        total = int(resp.headers.get("Content-Length") or 0)
        done = 0
        while True:
            chunk = resp.read(65536)
            if not chunk:
                break
            out.write(chunk)
            digest.update(chunk)
            done += len(chunk)
            if progress:
                progress(done, total)
    expected = (manifest.get("sha256") or "").lower()
    if expected and digest.hexdigest() != expected:
        os.remove(dest)
        raise RuntimeError("Downloaded file is damaged (checksum mismatch)")
    return dest


def launch_installer(path):
    """Run the installer silently; it closes this app, updates it and starts it again."""
    flags = 0x00000008 if os.name == "nt" else 0  # DETACHED_PROCESS
    subprocess.Popen([path, "/SILENT", "/CLOSEAPPLICATIONS"], close_fds=True, creationflags=flags)


# ---------------------------------------------------------------- self-test (used by the build)
def run_selftest():
    """Checks that the packaged OCR engines really work. Exit code 0 = OK."""
    from PIL import Image, ImageDraw, ImageFont

    result_file = os.path.join(tempfile.gettempdir(), "ticket_selftest_result.txt")
    try:
        setup_logging()
        img = Image.new("RGB", (1500, 520), "white")
        d = ImageDraw.Draw(img)
        font = ImageFont.load_default(size=52)
        for i, line in enumerate(["Order #: 39-57687/NEV", "TOTAL PAID $238.00", "Visa - 4528"]):
            d.text((60, 40 + i * 150), line, fill="black", font=font)
        path = os.path.join(tempfile.gettempdir(), "ticket_selftest.png")
        img.save(path)
        texts = run_ocr(path)
        rapid_ok = bool(_RAPID["engine"]) and not _RAPID["failed"]
        tess_ok = bool(_tesseract(path, 6).strip())
        row = parse_best(texts)[0]
        parse_ok = row["Order Number"] == "39-57687/NEV" and row["Total Cost"] == "$238.00"
        ok = rapid_ok and tess_ok and parse_ok
        msg = "SELFTEST %s | RapidOCR=%s Tesseract=%s Parse=%s | order=%r total=%r" % (
            "PASS" if ok else "FAIL", rapid_ok, tess_ok, parse_ok,
            row["Order Number"], row["Total Cost"])
    except Exception as e:  # noqa: BLE001
        LOGGER.exception("Self-test crashed")
        ok, msg = False, "SELFTEST CRASH: %r" % (e,)
    LOGGER.info(msg)
    with open(result_file, "w", encoding="utf-8") as f:
        f.write(msg + "\n")
    return 0 if ok else 1


# ---------------------------------------------------------------- window
def _draw_icon(kind, color, size=96):
    """Draw small icons in code (no separate image files needed)."""
    from PIL import Image, ImageDraw

    S = size * 4
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    w = int(S * 0.075)
    p = lambda *v: [int(S * x) for x in v]  # noqa: E731

    def line(*pts):
        d.line(p(*pts), fill=color, width=w, joint="curve")
        for i in range(0, len(pts), 2):
            x, y = int(S * pts[i]), int(S * pts[i + 1])
            d.ellipse([x - w // 2, y - w // 2, x + w // 2, y + w // 2], fill=color)

    def rrect(x0, y0, x1, y1, r=0.1, fill=None, outline=True):
        d.rounded_rectangle(p(x0, y0, x1, y1), radius=int(S * r), fill=fill,
                            outline=color if outline else None, width=w)

    if kind == "folder":
        d.polygon(p(.1, .28, .38, .28, .46, .38, .9, .38, .9, .8, .1, .8), fill=color)
    elif kind == "clipboard":
        rrect(.2, .16, .8, .9, .1)
        rrect(.36, .08, .64, .24, .05, fill=color)
        line(.34, .5, .66, .5)
        line(.34, .68, .6, .68)
    elif kind == "trash":
        line(.14, .24, .86, .24)
        line(.38, .24, .4, .12, .6, .12, .62, .24)
        rrect(.22, .3, .78, .9, .08)
        line(.42, .46, .42, .74)
        line(.58, .46, .58, .74)
    elif kind == "excel":
        rrect(.1, .1, .9, .9, .16, fill=color, outline=False)
        d.line(p(.32, .3, .68, .7), fill="white", width=w)
        d.line(p(.68, .3, .32, .7), fill="white", width=w)
    elif kind == "external":
        rrect(.12, .26, .74, .88, .1)
        d.rectangle(p(.5, .06, .96, .5), fill=(255, 255, 255, 0))
        line(.5, .5, .9, .1)
        line(.62, .1, .9, .1, .9, .38)
    elif kind in ("calendar", "calgrid"):
        rrect(.14, .2, .86, .88, .1)
        line(.14, .42, .86, .42)
        line(.32, .1, .32, .28)
        line(.68, .1, .68, .28)
        if kind == "calgrid":
            for x in (.32, .5, .68):
                for y in (.56, .72):
                    d.ellipse(p(x - .035, y - .035, x + .035, y + .035), fill=color)
    elif kind == "hash":
        line(.36, .16, .28, .84)
        line(.7, .16, .62, .84)
        line(.16, .38, .86, .38)
        line(.12, .64, .82, .64)
    elif kind == "clock":
        d.ellipse(p(.12, .12, .88, .88), outline=color, width=w)
        line(.5, .28, .5, .52, .68, .6)
    elif kind == "pin":
        d.ellipse(p(.22, .1, .78, .66), outline=color, width=w)
        line(.28, .58, .5, .9, .72, .58)
        d.ellipse(p(.42, .28, .58, .44), fill=color)
    elif kind == "cube":
        line(.5, .08, .88, .28, .88, .72, .5, .92, .12, .72, .12, .28, .5, .08)
        line(.12, .28, .5, .48, .88, .28)
        line(.5, .48, .5, .92)
    elif kind == "tag":
        line(.1, .12, .52, .12, .9, .5, .5, .9, .1, .5, .1, .12)
        d.ellipse(p(.24, .24, .36, .36), fill=color)
    elif kind == "mail":
        rrect(.1, .22, .9, .78, .08)
        line(.12, .28, .5, .56, .88, .28)
    elif kind == "coin":
        d.ellipse(p(.12, .12, .88, .88), outline=color, width=w)
        line(.5, .28, .5, .72)
        line(.62, .38, .4, .38, .4, .5, .6, .5, .6, .62, .38, .62)
    elif kind == "card":
        rrect(.08, .22, .92, .78, .08)
        line(.08, .42, .92, .42)
        line(.2, .62, .4, .62)
    elif kind == "info":
        d.ellipse(p(.12, .12, .88, .88), outline=color, width=w)
        line(.5, .46, .5, .7)
        d.ellipse(p(.45, .27, .55, .37), fill=color)
    elif kind == "seat":
        rrect(.2, .1, .8, .56, .12)
        rrect(.12, .6, .88, .78, .08)
        line(.26, .78, .26, .92)
        line(.74, .78, .74, .92)
    elif kind == "list":
        for y in (.28, .5, .72):
            d.ellipse(p(.12, y - .04, .2, y + .04), fill=color)
            line(.32, y, .88, y)
    elif kind == "download":
        line(.5, .1, .5, .62)
        line(.26, .42, .5, .66, .74, .42)
        line(.14, .84, .86, .84)
    elif kind == "up":
        line(.2, .62, .5, .34, .8, .62)
    elif kind == "down":
        line(.2, .38, .5, .66, .8, .38)
    elif kind == "logo":
        rrect(.04, .04, .96, .96, .26, fill=color, outline=False)
        d.rounded_rectangle(p(.24, .3, .76, .7), radius=int(S * .06), outline="white", width=w)
        d.ellipse(p(.19, .44, .29, .56), fill=color)
        d.ellipse(p(.71, .44, .81, .56), fill=color)
        d.line(p(.55, .34, .55, .66), fill="white", width=max(2, w // 2))
    return im.resize((size, size), Image.LANCZOS)


def _empty_illustration():
    from PIL import Image, ImageDraw

    S = 4
    W, H = 300 * S, 170 * S
    im = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    cloud = (232, 238, 255, 255)
    for box in ((20, 70, 140, 160), (90, 30, 210, 150), (160, 60, 285, 160), (40, 110, 270, 162)):
        d.ellipse([v * S for v in box], fill=cloud)
    d.rounded_rectangle([105 * S, 28 * S, 195 * S, 128 * S], radius=10 * S,
                        fill=(250, 251, 255, 255), outline=(150, 160, 190, 255), width=3 * S)
    d.line([105 * S, 54 * S, 195 * S, 54 * S], fill=(150, 160, 190, 255), width=3 * S)
    d.line([105 * S, 80 * S, 195 * S, 80 * S], fill=(190, 198, 222, 255), width=2 * S)
    d.line([105 * S, 104 * S, 195 * S, 104 * S], fill=(190, 198, 222, 255), width=2 * S)
    d.line([135 * S, 54 * S, 135 * S, 128 * S], fill=(190, 198, 222, 255), width=2 * S)
    d.line([165 * S, 54 * S, 165 * S, 128 * S], fill=(190, 198, 222, 255), width=2 * S)
    d.ellipse([172 * S, 100 * S, 212 * S, 140 * S], fill=(37, 99, 235, 255))
    d.line([192 * S, 112 * S, 192 * S, 128 * S], fill="white", width=4 * S)
    d.line([184 * S, 120 * S, 200 * S, 120 * S], fill="white", width=4 * S)
    for a, b in (((206, 24), (212, 32)), ((218, 34), (226, 38)), ((196, 20), (198, 12))):
        d.line([a[0] * S, a[1] * S, b[0] * S, b[1] * S], fill=(96, 125, 240, 255), width=3 * S)
    return im.resize((300, 170), Image.LANCZOS)


def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    import customtkinter as ctk
    from PIL import ImageTk

    # ---- colors (light lavender SaaS look)
    BG = "#EEF2FF"
    CARD = "#FFFFFF"
    LINE = "#DDE3F5"
    INK = "#0F172A"
    MUTED = "#64748B"
    BLUE = "#2563EB"
    INDIGO = "#4F46E5"
    GREEN = "#16A34A"
    RED = "#DC2626"
    FONT = "Segoe UI"

    ctk.set_appearance_mode("light")
    root = ctk.CTk()
    root.report_callback_exception = lambda *e: LOGGER.error("UI error", exc_info=e)
    root.title("Ticket Orders - Excel")
    root.configure(fg_color=BG)
    root.geometry("1400x760")
    root.minsize(1280, 600)
    sc = max(1.0, root.winfo_fpixels("1i") / 96.0)  # DPI scale for the ttk table

    def cimg(kind, color, size=20):
        return ctk.CTkImage(_draw_icon(kind, color), size=(size, size))

    # ---------------- header
    head = ctk.CTkFrame(root, fg_color="transparent")
    head.pack(fill="x", padx=22, pady=(18, 4))
    ctk.CTkLabel(head, text="", image=cimg("logo", INDIGO, 40)).pack(side="left")
    tbox = ctk.CTkFrame(head, fg_color="transparent")
    tbox.pack(side="left", padx=12)
    ctk.CTkLabel(head, text="v" + APP_VERSION, font=ctk.CTkFont(FONT, 12),
                 text_color=MUTED).pack(side="right", anchor="n", pady=4)
    ctk.CTkLabel(tbox, text="Ticket Orders - Excel", font=ctk.CTkFont(FONT, 22, "bold"),
                 text_color=INK).pack(anchor="w")
    ctk.CTkLabel(tbox, text="From Screenshot to Excel \u2014 Everything runs on your PC.",
                 font=ctk.CTkFont(FONT, 12), text_color=MUTED).pack(anchor="w")

    # ---------------- toolbar
    bar = ctk.CTkFrame(root, fg_color="transparent")
    bar.pack(fill="x", padx=22, pady=(10, 12))

    update_btn = ctk.CTkButton(
        bar, text="Updates", image=cimg("download", MUTED, 18), compound="left",
        width=140, height=44, corner_radius=22, font=ctk.CTkFont(FONT, 12), fg_color=CARD,
        hover_color="#E8EDFF", text_color=INK, border_width=1, border_color=LINE,
    )
    update_btn.pack(side="right")  # packed first so it always keeps its space

    def pill(text, kind, color, cmd, primary=False):
        b = ctk.CTkButton(
            bar, text=text, image=cimg(kind, "#FFFFFF" if primary else color, 20),
            compound="left", command=cmd, height=44, corner_radius=22,
            font=ctk.CTkFont(FONT, 12, "bold" if primary else "normal"),
            fg_color=INDIGO if primary else CARD,
            hover_color="#4338CA" if primary else "#E8EDFF",
            text_color="#FFFFFF" if primary else INK,
            border_width=0 if primary else 1, border_color=LINE,
        )
        b.pack(side="left", padx=(0, 8))
        return b

    # ---------------- table card
    card = ctk.CTkFrame(root, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
    card.pack(fill="both", expand=True, padx=22)
    card.grid_rowconfigure(0, weight=1)
    card.grid_columnconfigure(0, weight=1)

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure("Modern.Treeview", background=CARD, fieldbackground=CARD, foreground=INK,
                    rowheight=int(36 * sc), borderwidth=0, font=(FONT, 10),
                    bordercolor=CARD, lightcolor=CARD, darkcolor=CARD)
    style.configure("Modern.Treeview.Heading", background="#F5F7FF", foreground=INK,
                    font=(FONT, 10, "bold"), relief="solid", borderwidth=1,
                    bordercolor="#C7D2FE", lightcolor="#C7D2FE", darkcolor="#C7D2FE",
                    padding=(int(8 * sc), int(10 * sc)))
    style.map("Modern.Treeview", background=[("selected", "#E0E7FF")],
              foreground=[("selected", INK)])
    style.map("Modern.Treeview.Heading", background=[("active", "#EAEFFF")])
    try:
        style.layout("Modern.Treeview.Heading", [
            ("Treeheading.cell", {"sticky": "nswe"}),
            ("Treeheading.border", {"sticky": "nswe", "children": [
                ("Treeheading.padding", {"sticky": "nswe", "children": [
                    ("Treeheading.image", {"side": "left", "sticky": ""}),
                    ("Treeheading.text", {"sticky": "we"}),
                ]})]}),
        ])
    except tk.TclError:
        pass

    tree = ttk.Treeview(card, columns=COLUMNS, show="headings", selectmode="extended",
                        style="Modern.Treeview")
    xs = ctk.CTkScrollbar(card, orientation="horizontal", command=tree.xview, height=12,
                           fg_color="transparent", button_color="#C7D2FE",
                           button_hover_color="#A5B4FC")
    ys = ctk.CTkScrollbar(card, orientation="vertical", command=tree.yview, width=12,
                           fg_color="transparent", button_color="#C7D2FE",
                           button_hover_color="#A5B4FC")
    tree.configure(xscrollcommand=xs.set, yscrollcommand=ys.set)
    tree.configure(takefocus=0)
    tree.grid(row=0, column=0, sticky="nsew", padx=(6, 0), pady=(6, 0))
    ys.grid(row=0, column=1, sticky="ns", padx=(2, 6), pady=(8, 0))
    xs.grid(row=1, column=0, sticky="ew", padx=8, pady=(2, 8))

    head_icons = {
        "Event Name": "calendar", "Event Date & Time": "calgrid", "Month/Date/Year": "hash",
        "Time": "clock", "Venue (Location)": "pin", "Quantity": "cube",
        "Order Number": "tag", "Email": "mail", "Total Cost": "coin",
        "Seat details": "seat", "Sec": "hash", "Row": "hash", "Low seat": "seat",
        "Last 4 digit of card": "card", "Ticket Information": "info",
    }
    keep = []  # keep image references so they are not garbage collected
    widths = [260, 230, 150, 100, 270, 100, 140, 240, 110, 230, 80, 70, 90, 170, 340]
    for c, w in zip(COLUMNS, widths):
        ph = ImageTk.PhotoImage(_draw_icon(head_icons[c], INK, int(18 * sc)))
        keep.append(ph)
        tree.heading(c, text=" " + c, image=ph, anchor="w")
        tree.column(c, width=int(w * sc), minwidth=int(60 * sc), stretch=False, anchor="w")
    tree.tag_configure("odd", background="#F8FAFF")
    tree.tag_configure("even", background=CARD)

    # ---------------- empty state
    empty = tk.Frame(card, bg=CARD)
    art = ctk.CTkImage(_empty_illustration(), size=(300, 170))
    ctk.CTkLabel(empty, text="", image=art).pack(pady=(0, 6))
    ctk.CTkLabel(empty, text="No Data Loaded", font=ctk.CTkFont(FONT, 20, "bold"),
                 text_color=INK).pack()
    ctk.CTkLabel(empty, text='Click "1. Select Screenshot" to get started',
                 font=ctk.CTkFont(FONT, 13), text_color=MUTED).pack(pady=(2, 0))

    def refresh_state():
        for n, it in enumerate(tree.get_children()):
            tree.item(it, tags=("odd" if n % 2 else "even",))
        if tree.get_children():
            empty.place_forget()
        else:
            empty.place(in_=tree, x=0, y=int(40 * sc), relwidth=1, relheight=1,
                        height=-int(40 * sc))
            empty.lift()

    # ---------------- hint + log
    ctk.CTkLabel(root, text="Double-click a cell to edit (Enter = Save, Esc = Cancel). Please review the Email, "
                 "Venue, and Date.   Tip: press Win+Shift+S to capture, then click Paste from Clipboard.", font=ctk.CTkFont(FONT, 12),
                 text_color=MUTED).pack(anchor="w", padx=26, pady=(8, 4))

    logcard = ctk.CTkFrame(root, fg_color=CARD, corner_radius=16, border_width=1, border_color=LINE)
    logcard.pack(fill="x", padx=22, pady=(0, 18))
    lhead = ctk.CTkFrame(logcard, fg_color="transparent")
    lhead.pack(fill="x", padx=14, pady=(10, 4))
    ctk.CTkLabel(lhead, text=" Log / Status", image=cimg("list", INDIGO, 20), compound="left",
                 font=ctk.CTkFont(FONT, 14, "bold"), text_color=INK).pack(side="left")
    log = ctk.CTkTextbox(logcard, height=110, fg_color="#F8FAFF", text_color="#1E293B",
                         font=ctk.CTkFont("Consolas", 12), corner_radius=12,
                         border_width=1, border_color=LINE)
    log.configure(state="disabled")
    state = {"open": True}
    arrow_up, arrow_dn = cimg("up", MUTED, 20), cimg("down", MUTED, 20)

    def toggle_log():
        if state["open"]:
            log.pack_forget()
            toggler.configure(image=arrow_dn)
        else:
            log.pack(fill="x", padx=14, pady=(0, 12))
            toggler.configure(image=arrow_up)
        state["open"] = not state["open"]

    toggler = ctk.CTkButton(lhead, text="", image=arrow_up, width=32, height=28,
                            fg_color="transparent", hover_color="#EEF2FF", command=toggle_log)
    toggler.pack(side="right")
    log.pack(fill="x", padx=14, pady=(0, 12))

    def write_log(msg):
        log.configure(state="normal")
        log.insert("end", msg + "\n")
        log.see("end")
        log.configure(state="disabled")
        root.update_idletasks()

    # ---------------- double click to edit a cell
    editor = {"widget": None}

    def finish_edit(item, col_idx, entry, commit):
        if commit:
            vals = list(tree.item(item, "values"))
            vals[col_idx] = entry.get()
            tree.item(item, values=vals)
        entry.destroy()
        editor["widget"] = None

    def on_double_click(event):
        if tree.identify_region(event.x, event.y) != "cell":
            return
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item:
            return
        col_idx = int(col[1:]) - 1
        x, y, w, h = tree.bbox(item, col)
        if editor["widget"]:
            editor["widget"].destroy()
        entry = tk.Entry(tree, relief="flat", highlightthickness=2,
                         highlightbackground=INDIGO, highlightcolor=INDIGO,
                         font=(FONT, 10))
        entry.insert(0, tree.item(item, "values")[col_idx])
        entry.place(x=x, y=y, width=max(w, 140), height=h)
        entry.focus()
        entry.bind("<Return>", lambda e: finish_edit(item, col_idx, entry, True))
        entry.bind("<Escape>", lambda e: finish_edit(item, col_idx, entry, False))
        entry.bind("<FocusOut>", lambda e: finish_edit(item, col_idx, entry, True))
        editor["widget"] = entry

    tree.bind("<Double-1>", on_double_click)

    # ---------------- actions
    def process_files(paths):
        for p in paths:
            name = os.path.basename(p)
            write_log("Processing: " + name + " ...")
            try:
                rows = parse_best(run_ocr(p))
            except Exception as e:  # noqa: BLE001
                LOGGER.exception("Could not read %s", name)
                write_log("  ERROR: %s" % e)
                continue
            LOGGER.info("Read %s: %d row(s)", name, len(rows))
            for r in rows:
                tree.insert("", "end", values=[r[c] for c in COLUMNS])
            write_log("  %d row(s) found. Please review them before saving to the spreadsheet." % len(rows))
        refresh_state()

    def choose_files():
        paths = filedialog.askopenfilenames(
            title="Select Screenshot",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.webp *.bmp"), ("All", "*.*")],
        )
        if paths:
            process_files(paths)

    def from_clipboard():
        from PIL import ImageGrab

        data = ImageGrab.grabclipboard()
        if data is None:
            messagebox.showinfo(
                "Clipboard", "There is no screenshot on the clipboard.\n"
                "Press Windows + Shift + S to take a screenshot, then click this button."
            )
            return
        if isinstance(data, list):  # copied files
            process_files(data)
            return
        tmp = os.path.join(tempfile.gettempdir(), "ticket_clipboard.png")
        data.save(tmp)
        process_files([tmp])

    def delete_selected():
        for item in tree.selection():
            tree.delete(item)
        refresh_state()

    def clear_all():
        if tree.get_children() and messagebox.askyesno("Clear", "Remove all rows?"):
            tree.delete(*tree.get_children())
            refresh_state()

    def save_rows():
        items = tree.get_children()
        if not items:
            messagebox.showinfo("Save", "There are no rows to save.")
            return
        rows = []
        for it in items:
            vals = tree.item(it, "values")
            rows.append({c: str(v) for c, v in zip(COLUMNS, vals)})

        def ask_dup(keys):
            return messagebox.askyesno(
                "Duplicate",
                "%d row(s) (Order Number + Sec + Row) already exist in the sheet.\n\n"
                "Add them again anyway?" % len(keys),
            )

        try:
            added = save_to_excel(rows, ask_dup)
        except PermissionError:
            messagebox.showerror(
                "Excel", "Please close orders.xlsx (it is open in Excel) and try again."
            )
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))
            return
        write_log("%d row(s) saved to orders.xlsx: %s" % (added, EXCEL_FILE))
        tree.delete(*tree.get_children())
        refresh_state()
        messagebox.showinfo("Done", "%d row(s) saved.\n%s" % (added, EXCEL_FILE))

    def open_excel():
        if os.path.exists(EXCEL_FILE):
            os.startfile(EXCEL_FILE)  # Windows
        else:
            messagebox.showinfo("Excel", "orders.xlsx has not been created yet.")

    pill("1. Select Screenshot", "folder", BLUE, choose_files, primary=True)
    pill("Paste from Clipboard", "clipboard", INDIGO, from_clipboard)
    pill("Delete Row", "trash", RED, delete_selected)
    pill("Clear All", "trash", MUTED, clear_all)
    pill("2. Save to Excel", "excel", GREEN, save_rows)
    pill("Open Excel", "external", BLUE, open_excel)

    # ---------------- updates (button turns orange when a new version exists)
    ORANGE, ORANGE_HOVER = "#F59E0B", "#D97706"
    upd = {"manifest": None, "checked": False, "error": False, "manual": False,
           "pct": None, "ready": None, "failed": None}

    def highlight():
        m = upd["manifest"]
        update_btn.configure(
            text="Update v%s" % m["version"], image=cimg("download", "#FFFFFF", 18),
            fg_color=ORANGE, hover_color=ORANGE_HOVER, text_color="#FFFFFF",
            border_width=0, font=ctk.CTkFont(FONT, 12, "bold"),
        )

    def bg_check():
        try:
            upd["manifest"] = check_for_update()
            upd["error"] = False
        except Exception:  # noqa: BLE001
            LOGGER.info("Update check failed", exc_info=True)
            upd["error"] = True
        upd["checked"] = True

    def poll():
        if upd["failed"]:
            err, upd["failed"] = upd["failed"], None
            update_btn.configure(state="normal", text="Update v%s" % upd["manifest"]["version"])
            messagebox.showerror("Update", "The update could not be installed:\n%s" % err)
        elif upd["ready"]:
            path, upd["ready"] = upd["ready"], None
            LOGGER.info("Update downloaded, starting installer")
            launch_installer(path)
            root.after(800, root.destroy)
            return
        elif upd["pct"] is not None:
            update_btn.configure(text="Downloading... %d%%" % upd["pct"])
        elif upd["checked"]:
            upd["checked"] = False
            if upd["manifest"]:
                highlight()
                write_log("Update available: v%s" % upd["manifest"]["version"])
            elif upd["manual"]:
                if upd["error"]:
                    messagebox.showinfo("Update", "Could not check for updates.\n"
                                        "Please check your internet connection.")
                else:
                    messagebox.showinfo("Update", "You are up to date (v%s)." % APP_VERSION)
            upd["manual"] = False
            update_btn.configure(state="normal")
        root.after(700, poll)

    def start_download():
        m = upd["manifest"]

        def work():
            try:
                def prog(done, total):
                    upd["pct"] = int(done * 100 / total) if total else 0

                upd["pct"] = 0
                path = download_update(m, prog)
                upd["pct"] = None
                upd["ready"] = path
            except Exception as e:  # noqa: BLE001
                LOGGER.exception("Update download failed")
                upd["pct"] = None
                upd["failed"] = str(e)

        update_btn.configure(state="disabled", text="Downloading... 0%")
        threading.Thread(target=work, daemon=True).start()

    def on_update_click():
        m = upd["manifest"]
        if m:
            notes = (m.get("notes") or "").strip()
            msg = "Version %s is available (you have %s).\n\n%sDownload and install now?\n" \
                  "The app will close and reopen automatically." % (
                      m["version"], APP_VERSION, (notes + "\n\n") if notes else "")
            if messagebox.askyesno("Update available", msg):
                start_download()
        else:
            upd["manual"] = True
            update_btn.configure(state="disabled")
            threading.Thread(target=bg_check, daemon=True).start()

    update_btn.configure(command=on_update_click)
    threading.Thread(target=bg_check, daemon=True).start()
    root.after(700, poll)

    refresh_state()
    write_log("Ready. Excel file: " + EXCEL_FILE)
    root.mainloop()


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        sys.exit(run_selftest())
    setup_logging()
    if not acquire_single_instance():
        import tkinter as _tk
        from tkinter import messagebox as _mb

        _r = _tk.Tk()
        _r.withdraw()
        _mb.showinfo(APP_NAME, "Ticket Orders is already running.")
        sys.exit(0)
    run_gui()
