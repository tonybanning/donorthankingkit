"""Stage 1 - Clean.

Read the messy .xlsx, ask Claude to normalise it, then *re-validate every
record in code*. We never trust the model's output blindly: required fields
are enforced here, so a half-blank row can never slip through as "ready".
"""

import re

import openpyxl

import config
from pipeline import ai, offline

_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_SIMPLE_EMAIL = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def read_xlsx(path):
    """Return (header, rows) from the first worksheet, ignoring blank rows."""
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    rows = []
    for excel_row in sheet.iter_rows(values_only=True):
        # Normalise every cell to a trimmed string ("" for blanks).
        cells = ["" if cell is None else str(cell).strip() for cell in excel_row]
        if any(cells):  # drop the empty trailing rows Excel loves to leave behind
            rows.append(cells)
    workbook.close()
    if not rows:
        return [], []
    return rows[0], rows[1:]


def looks_unfilled(rows):
    """True if the sheet is still just the 'Insert ... here' starter boxes.

    Lets us stop with a friendly message instead of spending an API call
    thanking a donor named "Insert name here".
    """
    if not rows:
        return True
    for row in rows:
        for cell in row:
            text = str(cell).strip().lower()
            if text and not text.startswith("insert"):
                return False  # found something the user actually typed
    return True


def _coerce_amount(value):
    """Turn '$1,250.00' / '250' into a float, or return None if not numeric."""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = re.sub(r"[^0-9.\-]", "", str(value))
    try:
        return float(cleaned) if cleaned not in ("", "-", ".") else None
    except ValueError:
        return None


def _validate(record):
    """Coerce types and enforce required fields. Returns the fixed record.

    This is the fail-safe: whatever the model returned, a record only stays
    'ready' if it genuinely has a name, a valid-looking email and an amount.
    """
    notes = []
    original_note = (record.get("notes") or "").strip()
    if original_note:
        notes.append(original_note)

    record["donor_name"] = (record.get("donor_name") or "").strip()

    email = (record.get("email") or "").strip().lower()
    record["email"] = email or None

    record["amount"] = _coerce_amount(record.get("amount"))

    date = (record.get("date") or "").strip()
    if date and not _ISO_DATE.match(date):
        notes.append(f"could not parse date '{date}'")
        date = ""
    record["date"] = date or None

    record["donor_type"] = record.get("donor_type") or "first-time"

    # Leave genuinely empty/junk rows as the model marked them.
    if record.get("status") == "skipped":
        record["notes"] = "; ".join(n for n in notes if n)
        return record

    # Determine required-field failures ourselves.
    missing = []
    if not record["donor_name"]:
        missing.append("name")
    if not record["email"] or not _SIMPLE_EMAIL.match(record["email"]):
        missing.append("email")
    if record["amount"] is None:
        missing.append("amount")

    if missing:
        record["status"] = "flagged"
        notes.append("held back - missing/invalid " + ", ".join(missing))
    elif record.get("status") not in ("ready", "flagged", "skipped"):
        record["status"] = "ready"

    record["notes"] = "; ".join(n for n in notes if n)
    return record


def _dedupe_ready(records, change_log):
    """Safety net: if two 'ready' records share an email, keep the first."""
    seen = {}
    result = []
    for record in records:
        if record["status"] == "ready" and record["email"]:
            key = record["email"]
            if key in seen:
                change_log.append(f"merged duplicate ready record for {key}")
                continue
            seen[key] = True
        result.append(record)
    return result


def run(input_path):
    """Execute Stage 1. Returns (records, change_log)."""
    header, rows = read_xlsx(input_path)
    if not rows:
        return [], ["input file had no data rows"]

    if ai.enabled():
        try:
            data = ai.clean_rows(header, rows)
        except Exception as exc:  # noqa: BLE001 - a failed API call must never
            print(f"  ! Claude cleaning failed ({exc}); falling back to free mode.")
            data = offline.clean_rows(header, rows)  # ...cost you the whole run
    else:
        data = offline.clean_rows(header, rows)

    raw_records = data.get("records", []) if isinstance(data, dict) else []
    change_log = list(data.get("change_log", [])) if isinstance(data, dict) else []

    records = [_validate(dict(r)) for r in raw_records if isinstance(r, dict)]
    records = _dedupe_ready(records, change_log)

    ready = sum(1 for r in records if r["status"] == "ready")
    flagged = sum(1 for r in records if r["status"] == "flagged")
    skipped = sum(1 for r in records if r["status"] == "skipped")
    change_log.append(
        f"validated in code: {ready} ready, {flagged} flagged, {skipped} skipped"
    )
    return records, change_log
