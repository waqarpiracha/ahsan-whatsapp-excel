"""
Ahsan Mushtaq — WhatsApp chat export parser (importer v1).

Reads a WhatsApp .txt export (plain non-bracketed format:
  24/02/2026, 5:14 pm - Sender Name: message body
possibly spanning multiple lines) and extracts EV-charger survey rows.

Message templates were derived from the REAL export
(ahsan-sample-export.txt, 16,804 messages, EV charging solutions by Comet).

Row types:
  survey  — "*Survey Report* / *Resurvey Report* / Survey done" with cable
            measurements  -> Status=Pending
  assign  — "Survey Required" with client details        -> Status=Pending
  install — "EV charger installation complete"           -> Status=Closed
  cable   — power/grounding cable lengths with a client  -> Status=Pending

Anything that looks data-bearing but cannot be turned into a row goes to
the review list — NEVER guessed. Pure chatter is ignored silently.

Never-guess rule is absolute: no invented names, numbers, cities or brands.
"""

import re
from datetime import datetime

# ---------------------------------------------------------------------------
# Export line parsing
# ---------------------------------------------------------------------------

LINE_RE = re.compile(
    r"^(\d{1,2}/\d{1,2}/\d{4}),\s+(\d{1,2}:\d{2})\s*(am|pm)\s*-\s*([^:]+?):\s?(.*)$"
)

SYSTEM_RE = re.compile(
    r"(created group|were added|was added|changed the group|"
    r"Messages and calls are end-to-end encrypted|"
    r"Waiting for this message|changed their phone number)",
    re.I,
)


def parse_export(path, date_format="DMY"):
    """Parse the raw export into message dicts.

    date_format: "DMY" (default) or "MDY" — used only when the date is
    ambiguous (both components <= 12). Never silently defaults: caller must
    pass the user's explicit choice.
    """
    messages = []
    with open(path, encoding="utf-8") as f:
        for raw in f:
            line = raw.rstrip("\n")
            m = LINE_RE.match(line)
            if m:
                d, t, ap, sender, body = m.groups()
                messages.append({
                    "date_s": d, "time_s": f"{t} {ap}",
                    "sender": sender.strip(), "body": body,
                })
            elif messages:
                messages[-1]["body"] += "\n" + line
    for msg in messages:
        msg["date"] = _parse_date(msg["date_s"], date_format)
    return messages


def _parse_date(s, date_format):
    a, b, y = (int(x) for x in s.split("/"))
    if date_format == "MDY":
        # export is day-first; only honour MDY when unambiguous
        if a > 12:  # cannot be a month -> treat as day
            day, month = a, b
        else:
            month, day = a, b
    else:  # DMY
        if b > 12:  # cannot be a month -> first component is the day
            day, month = a, b
        elif a > 12:
            day, month = a, b
        else:
            day, month = a, b  # ambiguous -> explicit DMY choice
    try:
        return datetime(y, month, day).date()
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Field extractors
# ---------------------------------------------------------------------------

BRAND_CANON = {
    "byd": "BYD", "byd shark": "BYD Shark",
    "deepal": "Deepal",
    "tank phev": "Tank PHEV", "tank p": "Tank P", "tank 500": "Tank 500",
    "tank": "Tank",
    "haval": "Haval", "haval h6": "Haval H6", "h6 phev": "H6 PHEV",
    "gac": "GAC", "j7": "J7",
    "jaecoo": "Jaecoo", "jaecoo j7": "Jaecoo J7",
    "changan": "Changan", "mg": "MG", "audi": "Audi",
}
# longest-first so "tank phev" wins over "tank"
BRAND_KEYS = sorted(BRAND_CANON, key=len, reverse=True)
NOISE_PARENS = {"file attached", "whatsapp", "edited"}

PHONE_RE = re.compile(r"(?:\+92[\s\-]?|0)?3\d{2}[\s\-]?\d{7}")


def extract_phones(body):
    """Return list of normalized 10-digit strings (leading 0/92 stripped)."""
    out = []
    for m in PHONE_RE.finditer(body):
        digits = re.sub(r"\D", "", m.group(0))
        if digits.startswith("92") and len(digits) == 12:
            digits = digits[2:]
        elif digits.startswith("0") and len(digits) == 11:
            digits = digits[1:]
        if len(digits) == 10 and digits.startswith("3") and digits not in out:
            out.append(digits)
    return out


def extract_brand(body):
    low = body.lower()
    # 0) "EV charger Jaecoo J7 installed" form
    m = re.search(r"ev\s+charger\s+([a-z0-9 .&/-]{2,20}?)\s+install", low)
    if m:
        cand = m.group(1).strip()
        for key in BRAND_KEYS:
            if key in cand:
                return BRAND_CANON[key]
    # 1) header forms: "Survey Required (BYD)" / "*Survey Required BYD*"
    m = re.search(
        r"survey\s+(?:re)?required\s+\(?\*?([a-z0-9 .&/-]{2,20})\*?\)?", low)
    if m:
        cand = m.group(1).strip().strip("*")
        if cand in BRAND_CANON:
            return BRAND_CANON[cand]
    m = re.search(
        r"survey\s+report\s+\(?\*?([a-z0-9 .&/-]{2,20})\*?\)?", low)
    if m:
        cand = m.group(1).strip().strip("*")
        if cand in BRAND_CANON:
            return BRAND_CANON[cand]
    # 2) parenthesised brand anywhere: "(Tank P)"
    for m in re.finditer(r"\(([a-z0-9 .&/-]{2,20})\)", low):
        cand = m.group(1).strip()
        if cand in NOISE_PARENS:
            continue
        if cand in BRAND_CANON:
            return BRAND_CANON[cand]
    # 3) bare brand word as its own line (or brand-led line)
    for line in low.split("\n"):
        s = line.strip().strip("*")
        if s in BRAND_CANON:
            return BRAND_CANON[s]
        if len(s) <= 30:
            for key in BRAND_KEYS:
                if s.startswith(key):
                    return BRAND_CANON[key]
    return ""


def _field_block(body, *names):
    """Grab '*Name* : value' (possibly multi-line) for any of names."""
    for name in names:
        pat = re.compile(
            r"\*?" + re.escape(name) + r"\*?\s*:+\s*(.+?)"
            r"(?=\n\s*\*?[A-Za-z][\w ]{0,20}\*?\s*:|\n\s*\n|\Z)",
            re.I | re.S,
        )
        m = pat.search(body)
        if m:
            val = m.group(1).strip()
            val = re.sub(r"\s*\n\s*", ", ", val).strip(" ,")
            if val:
                return val
    return ""


def extract_client(body):
    return _field_block(body, "Client")


def extract_address(body):
    addr = _field_block(body, "Location", "Address")
    if addr:
        # a brand line glued to the address ("... road adda 11 L, BYD SHARK
        # 6- 2026") belongs to the vehicle, not the address
        parts = [p.strip() for p in addr.split(",")]
        while parts and any(parts[-1].lower().startswith(k)
                            for k in BRAND_KEYS):
            parts.pop()
        addr = ", ".join(parts)
    return addr


CITIES = [
    # major + towns seen in the export (word-boundary matched, last wins)
    "lahore", "karachi", "islamabad", "rawalpindi", "multan", "faisalabad",
    "jhelum", "kasur", "narowal", "sheikhupura", "khanewal", "bahawalnagar",
    "vehari", "layyah", "okara", "depalpur", "chishtian", "sharqpur",
    "burewala", "chichawatni", "sahiwal", "sialkot", "gujranwala", "gujrat",
    "peshawar", "quetta", "hyderabad", "sukkur", "sargodha",
    "dera ghazi khan", "bahawalpur", "rahim yar khan", "abbottabad",
    "mardan", "jhang", "toba tek singh", "chiniot", "kamalia", "arifwala",
    "pakpattan", "mailsi", "dunyapur", "gojra", "samundri", "muridke",
    "kamoke", "daska", "pasrur", "fateh pur", "manawala", "abdul hakeem",
    "kharian", "wazirabad", "hafizabad", "nankana", "shahkot", "sangla",
    "phalia", "mandi bahauddin", "chakwal", "talagang", "attock", "taxila",
    "wah cantt", "haripur", "mansehra", "kohat", "bannu", "dera ismail khan",
    "sukkur", "larkana", "nawabshah", "mirpurkhas", "jacobabad", "shikarpur",
]


def extract_city(address):
    if not address:
        return ""
    low = " " + address.lower().replace(",", " ") + " "
    # "dha lahore" / "dha karachi" style is the most reliable signal
    m = re.search(r"\bdha\s+([a-z ]+?)(?:\s|$)", low)
    if m and m.group(1).strip() in ("lahore", "karachi", "islamabad",
                                    "multan", "quetta", "peshawar"):
        return m.group(1).strip().title()
    best, pos = "", -1
    for city in CITIES:
        for m in re.finditer(r"\b" + re.escape(city) + r"\b", low):
            if m.start() >= pos:
                pos, best = m.start(), city
    if best:
        return " ".join(w.title() for w in best.split())
    if "cantt" in low:
        # e.g. "jhelum cantt" — city is the word before cantt
        m = re.search(r"([a-z ]+?)\s+cantt", low)
        if m:
            return m.group(1).strip().split()[-1].title()
    return ""


LEN_UNIT = r"(ft|feet|meter|meters|m)\b"
# NOTE: "10mm" (cable thickness) must never match as a length.


def _to_meters(value, unit):
    v = float(value)
    if unit.lower().startswith("f"):
        return round(v * 0.3048, 1)
    return round(v, 1)


def extract_cables(body):
    """Return (power_m, ground_m, original_text). None where not found."""
    power = ground = None
    notes = []
    # "Power and grounding cable 50ft each" -> both
    m = re.search(
        r"power\s+and\s+ground(?:ing)?\s+cable\s+"
        r"(\d+(?:\.\d+)?)\s*" + LEN_UNIT + r"\s*each", body, re.I)
    if m:
        v = _to_meters(m.group(1), m.group(2))
        power = ground = v
        notes.append(f"Power and grounding cable {m.group(1)}{m.group(2)} each")
    else:
        m = re.search(
            r"power\s*(?:cable)?\s*(?:\d+\s*core\s*)?"
            r"(\d+(?:\.\d+)?)\s*" + LEN_UNIT, body, re.I)
        if not m:  # "300ft power cable" (number first)
            m = re.search(
                r"(\d+(?:\.\d+)?)\s*" + LEN_UNIT + r"\s*power\s*cable",
                body, re.I)
        if m:
            power = _to_meters(m.group(1), m.group(2))
            notes.append(f"Power cable {m.group(1)} {m.group(2)}")
        m = re.search(
            r"(?:ground(?:ing)?|earth(?:ing)?)\s*cable\s*"
            r"(?:\d+\s*core\s*)?(\d+(?:\.\d+)?)\s*" + LEN_UNIT, body, re.I)
        if m:
            ground = _to_meters(m.group(1), m.group(2))
            notes.append(f"Grounding cable {m.group(1)} {m.group(2)}")
    return power, ground, "; ".join(notes)


def extract_bore(body):
    low = body.lower()
    if re.search(r"no\s+bore\s+requir", low):
        return "Not required"
    if "already done" in low and "bore" in low:
        return "Done"
    if re.search(r"(earth\s+bore|bore).{0,12}requir", low):
        return "Required"
    if "earthing" in low and "availab" in low:
        return "Available"
    return ""


def extract_remote_local(body):
    if re.search(r"remote\s+area\s+site", body, re.I):
        return "Remote"
    return ""


# ---------------------------------------------------------------------------
# Row building
# ---------------------------------------------------------------------------

CHATTER_RE = re.compile(r"^(@.+|ok|👍+|yes|no|done|noted|seen)$", re.I)


def _is_chatter(body):
    b = body.strip()
    if not b or b == "<Media omitted>" or b == "This message was deleted":
        return True
    if len(b) <= 20 and not re.search(
            r"survey|client|power|ground|earth|bore|install|meter|\bft\b",
            b, re.I):
        return True
    return False


def _base_row(msg):
    return {
        "date": msg["date"],
        "time": msg["time_s"],
        "client_name": "",
        "address": "",
        "city": "",
        "client_cell": "",
        "multiline_poc": "",
        "vehicle": "",
        "priority": "",
        "status": "Pending",
        "charger_detail": "",
        "surveyor_name": "",
        "power_m": None,
        "ground_m": None,
        "bore": "",
        "surveyor_exp": "",
        "remarks": "",
        "remote_local": "",
        "full_message": "",
    }


def _fill_positional_assign(row, body):
    """Positional format used in 'Survey Required' messages:

        *Survey Required BYD*
        Dr Habib Ur Rehman
        52/12, Naveed Rasul Mirza Road, Jhelum Cantt
        03224801229
    first non-phone line = name, following non-phone lines = address.
    """
    lines = []
    for ln in body.split("\n"):
        s = ln.strip().strip("*").strip()
        if not s or "survey required" in s.lower() or s.startswith("@"):
            continue
        lines.append(s)
    name, addr_lines = "", []
    for ln in lines:
        if PHONE_RE.search(ln) and not re.search(r"[a-zA-Z]{3,}", ln):
            continue  # pure phone line -> handled by extract_phones
        # "Muhammad Shahid - +92 333 4330751" -> name before the dash
        m = re.match(r"^(.{2,40}?)\s*[-–]\s*(?:\+92[\s\-]?|0)?3\d{2}", ln)
        if m and _looks_like_name(m.group(1)) and not name:
            name = m.group(1).strip()
            continue
        if not name and _looks_like_name(ln):
            name = ln
        else:
            addr_lines.append(ln)
    if name and not row["client_name"]:
        row["client_name"] = name
    if addr_lines and not row["address"]:
        row["address"] = ", ".join(addr_lines)
    return row


ADDR_WORDS = re.compile(
    r"house|street|road|town|block|phase|chowk|colony|plaza|hospital|"
    r"shop|sector|mohalla|gali|no\.|#|society|avenue|multiline|box|"
    r"wiring|breaker|charger|check", re.I)


def _looks_like_name(s):
    """A person name: 1-4 words, letters/dots only, not an address line."""
    if len(s) > 40 or ADDR_WORDS.search(s):
        return False
    return bool(re.match(r"^[A-Za-z. ]{2,40}$", s)
                and 1 <= len(s.split()) <= 4)


def _fill_common(row, body, positional=False):
    row["client_name"] = extract_client(body)
    row["address"] = extract_address(body)
    if positional:
        row = _fill_positional_assign(row, body)
    row["city"] = extract_city(row["address"])
    row["vehicle"] = extract_brand(body)
    phones = extract_phones(body)
    if phones:
        row["client_cell"] = phones[0]
        if len(phones) > 1:
            row["remarks"] = "Alt numbers: " + ", ".join(phones[1:])
    row["power_m"], row["ground_m"], cable_note = extract_cables(body)
    if cable_note:
        row["remarks"] = (row["remarks"] + "; " + cable_note).strip("; ")
    row["bore"] = extract_bore(body)
    row["remote_local"] = extract_remote_local(body)
    return row


def _testing_note(body):
    low = body.lower()
    if "testing not done" in low or "testing not  done" in low:
        m = re.search(r"charger testing[^.👆\n]*", body, re.I)
        return m.group(0).strip() if m else "charger testing not done"
    if "testing done" in low:
        return "charger testing done"
    return ""


def extract_rows(messages):
    """Return (rows, review, stats). rows are dicts ready for the writer."""
    rows, review = [], []
    stats = {"total": len(messages), "ignored": 0,
             "survey": 0, "assign": 0, "install": 0, "cable": 0,
             "review": 0}
    for msg in messages:
        body = msg["body"] or ""
        if SYSTEM_RE.search(body) or _is_chatter(body):
            stats["ignored"] += 1
            continue
        low = body.lower()
        is_survey = ("survey report" in low or "resurvey" in low
                     or "survey done" in low)
        is_assign = "survey required" in low
        is_install = ("installation complete" in low
                      or ("ev charger" in low
                          and ("installed" in low or "complete" in low)))
        has_cable = bool(re.search(
            r"(power|ground(?:ing)?|earth(?:ing)?)\s*(cable)?\s*"
            r"(\d+\s*core\s*)?\d+(\.\d+)?\s*(ft|feet|meter|meters|m)\b",
            body, re.I)) or "cable" in low and re.search(
            r"\d+\s*(ft|feet|meter|meters)\b", body, re.I)

        row = None
        rtype = None
        if is_survey:
            row = _fill_common(_base_row(msg), body)
            row["status"] = "Pending"
            rtype = "survey"
        elif is_install:
            row = _fill_common(_base_row(msg), body)
            row["status"] = "Closed"
            note = _testing_note(body)
            if note:
                row["remarks"] = (row["remarks"] + "; " + note).strip("; ")
            rtype = "install"
        elif is_assign:
            row = _fill_common(_base_row(msg), body, positional=True)
            row["status"] = "Pending"
            rtype = "assign"
        elif has_cable:
            row = _fill_common(_base_row(msg), body)
            row["status"] = "Pending"
            rtype = "cable"

        if row is None:
            # data-ish but unclassifiable -> review, never guessed.
            # asterisked *Client*/*Location* markers always qualify;
            # bare keywords need substance (>30 chars) so one-word chatter
            # like "Survey??" stays ignored.
            if ("*client*" in low or "*location*" in low
                    or (len(body) > 30 and re.search(
                        r"survey|installation|power\s+cable|"
                        r"ground(?:ing)?\s+cable|earth\s+bore|earthing|"
                        r"ev\s+charger", low))):
                review.append({
                    "date": msg["date_s"], "time": msg["time_s"],
                    "sender": msg["sender"],
                    "message": body[:300].replace("\n", " "),
                    "reason": "data-bearing but no parseable template",
                })
                stats["review"] += 1
            else:
                stats["ignored"] += 1
            continue

        # a row with neither client nor cable nor address is not a row.
        # an "address" that is really work-order prose does not count.
        addr_ok = (row["address"] and len(row["address"]) <= 150
                   and not re.search(r"wiring|breaker|take a picture|"
                                     r"just need to go", row["address"], re.I))
        if not (row["client_name"] or addr_ok
                or row["power_m"] is not None or row["ground_m"] is not None
                or row["client_cell"]):
            review.append({
                "date": msg["date_s"], "time": msg["time_s"],
                "sender": msg["sender"],
                "message": body[:300].replace("\n", " "),
                "reason": "survey/install keyword but no client, address, "
                          "phone or cable found",
            })
            stats["review"] += 1
            continue
        stats[rtype] += 1
        row["_type"] = rtype
        row["full_message"] = body
        rows.append(row)
    return rows, review, stats
