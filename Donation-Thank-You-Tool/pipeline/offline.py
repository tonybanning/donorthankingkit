"""FREE MODE - the whole tool, with no API key and no internet.

Everything Claude was doing is done here in plain Python instead:

  * clean_rows()        - parses and tidies the spreadsheet deterministically
  * personal_note()     - picks a warm sentence written for that donor's
                          situation (first-time vs recurring, small vs large)

The trade-off is honest: the AI version writes a fresher, more individual
sentence and copes with weirder spreadsheets. This version costs nothing, needs
no account, and never sends your donor data anywhere. For most small nonprofits
that is the better deal, so it is the default.
"""

import re
from datetime import datetime

# Amounts a volunteer might type as words.
_WORD_NUMBERS = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
    "sixteen": 16, "seventeen": 17, "eighteen": 18, "nineteen": 19,
    "twenty": 20, "thirty": 30, "forty": 40, "fifty": 50, "sixty": 60,
    "seventy": 70, "eighty": 80, "ninety": 90,
}

# Date formats people actually type.
_DATE_FORMATS = (
    "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d/%m/%Y", "%Y/%m/%d",
    "%B %d, %Y", "%B %d %Y", "%b %d, %Y", "%b %d %Y", "%d %B %Y", "%m-%d-%Y",
)


def parse_amount(text):
    """'$1,250.00' / '50' / 'fifty' -> float, or None if it isn't a number."""
    if text is None or str(text).strip() == "":
        return None
    if isinstance(text, (int, float)):
        return float(text)

    # Digits first: strip currency symbols, commas, spaces.
    stripped = re.sub(r"[^0-9.\-]", "", str(text))
    if stripped not in ("", "-", ".", "-."):
        try:
            return float(stripped)
        except ValueError:
            pass

    # Then words: "fifty", "one hundred", "two hundred fifty".
    words = re.findall(r"[a-z]+", str(text).lower())
    if not words:
        return None
    total, current, seen = 0, 0, False
    for word in words:
        if word in _WORD_NUMBERS:
            current += _WORD_NUMBERS[word]
            seen = True
        elif word == "hundred" and seen:
            current = (current or 1) * 100
        elif word == "thousand" and seen:
            total += (current or 1) * 1000
            current = 0
        elif word in ("and", "dollars", "dollar", "usd", "bucks"):
            continue
        else:
            return None  # a word we don't understand -> don't guess
    return float(total + current) if seen else None


def parse_date(text):
    """Any common date format -> 'YYYY-MM-DD'. Returns None if unparseable."""
    text = str(text or "").strip()
    if not text:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def tidy_name(text):
    """'  SMITH, ROBERT ' -> 'Robert Smith'.  'bob smith' -> 'Bob Smith'."""
    name = " ".join(str(text or "").split())
    if not name:
        return ""
    if "," in name:  # "Last, First" -> "First Last"
        last, _, first = name.partition(",")
        name = f"{first.strip()} {last.strip()}".strip()
    # Only re-case if the volunteer shouted or whispered; leave "McDonald" alone.
    if name.isupper() or name.islower():
        name = name.title()
    return name


def _find_column(header, keywords):
    """Index of the first column whose name contains one of `keywords`."""
    for index, cell in enumerate(header):
        lowered = str(cell).lower()
        if any(keyword in lowered for keyword in keywords):
            return index
    return None


def _cell(row, index):
    if index is None or index >= len(row):
        return ""
    return str(row[index]).strip()


def clean_rows(header, rows):
    """Deterministic replacement for the Claude cleaning pass.

    Returns the same shape the AI version returns:
    {"records": [...], "change_log": [...]}
    """
    name_at = _find_column(header, ("name", "donor", "supporter"))
    amount_at = _find_column(header, ("amount", "donation", "gift", "sum", "total", "$"))
    email_at = _find_column(header, ("email", "e-mail", "mail"))
    date_at = _find_column(header, ("date", "received", "when"))

    # Fall back to position if the headers were renamed to something odd.
    if name_at is None:
        name_at = 0
    if email_at is None and rows:
        for index in range(len(rows[0])):
            if any("@" in _cell(row, index) for row in rows):
                email_at = index
                break

    dates_fixed = 0
    gifts = []  # one entry per spreadsheet row that has anything in it

    for row in rows:
        name = tidy_name(_cell(row, name_at))
        email = _cell(row, email_at).lower()
        raw_amount = _cell(row, amount_at)
        raw_date = _cell(row, date_at)

        amount = parse_amount(raw_amount)
        date = parse_date(raw_date)
        if raw_date and date and raw_date != date:
            dates_fixed += 1

        notes = []
        if raw_amount and amount is None:
            notes.append(f"could not read the donation amount '{raw_amount}'")
        if raw_date and date is None:
            notes.append(f"could not read the date '{raw_date}'")

        gifts.append({
            "donor_name": name,
            "email": email or None,
            "amount": amount,
            "date": date,
            "notes": notes,
        })

    # --- Group the gifts into one record per donor -------------------------
    groups = {}
    order = []
    for gift in gifts:
        # Same email = same person. No email? Fall back to the name.
        key = gift["email"] or f"name:{gift['donor_name'].lower()}"
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(gift)

    merged = 0
    records = []
    for key in order:
        group = groups[key]
        if len(group) > 1:
            merged += len(group) - 1

        # The most recent gift represents the donor (undated gifts sort last).
        latest = sorted(group, key=lambda g: g["date"] or "")[-1]
        amounts = [g["amount"] for g in group if g["amount"] is not None]

        notes = []
        for gift in group:
            notes.extend(gift["notes"])
        if len(group) > 1:
            total = sum(amounts)
            notes.insert(0, f"recurring donor: {len(group)} gifts totalling ${total:,.2f}")

        records.append({
            "donor_name": latest["donor_name"],
            "email": latest["email"],
            "amount": latest["amount"],
            "date": latest["date"],
            "donor_type": "recurring" if len(group) > 1 else "first-time",
            "status": "ready",  # clean.py re-checks this and flags what's missing
            "notes": "; ".join(notes),
        })

    change_log = [f"read {len(gifts)} row(s) from the spreadsheet"]
    if merged:
        change_log.append(f"merged {merged} duplicate row(s) into existing donors")
    if dates_fixed:
        change_log.append(f"normalised {dates_fixed} date(s) to YYYY-MM-DD")
    change_log.append("cleaned offline - no API key used, no data sent anywhere")

    return {"records": records, "change_log": change_log}


# --- The personal sentence, without an AI ----------------------------------
#
# A form letter that says the same thing to a $10 first-timer and a $500
# regular is exactly what we're trying to avoid. So we pick from sentences
# written for each situation, and vary them so consecutive letters differ.

_NOTES = {
    # --- Under $25: a token gift. Humble, sincere, no fireworks. ------------
    ("first-time", "small"): [
        "It means a great deal that you chose to give at all, and we don't take that lightly.",
        "Thank you for taking that first step with us. Gifts this size keep the lights on, quietly and reliably.",
        "We're simply glad you're here. Every gift counts, and yours is no exception.",
    ],
    ("recurring", "small"): [
        "You keep coming back, and we notice. That steadiness is worth more to us than any single large gift.",
        "Thank you for giving again. Faithful friends like you are the quiet backbone of everything we manage to do.",
        "Your support arrives without fail, and we've come to depend on it more than you probably realize.",
    ],
    # --- $25-$249: a real contribution. Warm and specific. ------------------
    ("first-time", "medium"): [
        "Your first gift is a real encouragement, and we're grateful you decided to trust us with it.",
        "Thank you for choosing to support this work. A first gift of this size opens real doors for us.",
        "You've arrived at exactly the right moment, and we're glad to welcome you into this community.",
    ],
    ("recurring", "medium"): [
        "Your continued generosity is something we genuinely plan around, and we don't take it for granted.",
        "Year after year, you keep showing up for this work. Thank you for being someone we can count on.",
        "Supporters who stay the course are what make long-term work possible - and you have stayed the course.",
    ],
    # --- $250-$999: significant. Name what it unlocks. ----------------------
    ("first-time", "large"): [
        "We're genuinely moved that your very first gift was such a generous one, and we will put every dollar of it to work with care.",
        "A first gift of this size is a serious vote of confidence, and we fully intend to live up to it.",
        "Thank you for beginning your support so generously. This will make a difference people will actually feel.",
    ],
    ("recurring", "large"): [
        "Your sustained and substantial support has shaped what this organization is able to do, and we are deeply grateful to have you in our corner.",
        "Few people give as faithfully, or as generously, as you have. Thank you for continuing to believe in this work.",
        "Your ongoing commitment is felt in everything we accomplish. We would not be who we are without supporters like you.",
    ],
    # --- $1,000+: transformational. They changed a life. --------------------
    ("first-time", "major"): [
        "It is hard to overstate what a gift like this does. Somewhere out there is a person whose life will be permanently different because you decided to give - and you did it without ever having given to us before. Thank you. Truly.",
        "A gift of this magnitude changes lives, and we mean that literally, not as a figure of speech. You have altered someone's future for good, and we are honored that you chose us to make it happen.",
        "You have done something extraordinary. This gift will reach a real person, at a real moment of need, and change the direction of their life - and you did it on your very first gift. We will never forget it.",
    ],
    ("recurring", "major"): [
        "Year after year you have given at a level that changes lives, and lives have in fact been changed - permanently, and for the better. You are one of the reasons this organization exists at all. Thank you, from all of us.",
        "There are people walking a different path today because of you. Your continued generosity at this level is not simply support; it is the thing itself, the reason the work is possible. We are profoundly grateful.",
        "Your faithfulness at this magnitude has transformed lives, and it has transformed us. We do not have adequate words - only the certainty that what you have made possible will outlast all of us.",
    ],
}

_FALLBACK = (
    "Your generosity means a great deal to us, and we are grateful for your support."
)


def _tier(amount):
    """Gift size band. Drives how warm the personal sentence gets.

    The humble register is reserved for genuinely token gifts (under $25).
    Anything from $25 up is a real, considered contribution and is thanked
    warmly - underplaying it would be its own kind of insult.
    """
    try:
        amount = float(amount)
    except (TypeError, ValueError):
        return "medium"
    if amount < 25:
        return "small"       # token gift - quiet, humble thanks
    if amount < 250:
        return "medium"      # a real contribution - warm thanks
    if amount < 1000:
        return "large"       # significant - name what it unlocks
    return "major"           # transformational - you changed a life


def personal_note(record, index=0):
    """Return a warm sentence suited to this donor. Free, instant, offline."""
    donor_type = record.get("donor_type") or "first-time"
    if donor_type not in ("first-time", "recurring"):
        donor_type = "first-time"
    options = _NOTES.get((donor_type, _tier(record.get("amount"))))
    if not options:
        return _FALLBACK
    # Vary the wording between letters so a batch doesn't read like a mail merge.
    return options[index % len(options)]
