"""All calls to Claude live here.

Two jobs:
  * clean_rows()          -> Stage 1: turn a messy table into structured records
  * draft_personal_note() -> Stage 2: write one warm sentence tuned to the donor

The API key is read from the environment (ANTHROPIC_API_KEY). It is never
hardcoded, logged, or written to disk.
"""

import json
import os
import re

from anthropic import Anthropic

import config

_client = None


def enabled():
    """True if we should call the API at all.

    The tool is free by default: with no key set, this returns False and the
    pipeline uses pipeline/offline.py instead. Nothing breaks, nothing costs.
    """
    if config.USE_AI == "off":
        return False
    if config.USE_AI == "on":
        return True
    return bool(os.environ.get("ANTHROPIC_API_KEY"))


def client():
    """Return a lazily-created Anthropic client, or raise a clear error."""
    global _client
    if _client is None:
        key = os.environ.get("ANTHROPIC_API_KEY")
        if not key:
            raise RuntimeError(
                "USE_AI is 'on' but ANTHROPIC_API_KEY is not set.\n"
                "Either add a key, or set USE_AI = 'off' in config.py to run free."
            )
        _client = Anthropic(api_key=key)
    return _client


def _text_of(response):
    """Concatenate the text blocks of a Claude response."""
    return "".join(block.text for block in response.content if block.type == "text")


def _extract_json(text):
    """Pull a JSON object out of a model response, tolerating code fences.

    We never trust the model blindly, so callers still validate the parsed
    result field by field. This just gets us from raw text to a Python dict.
    """
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\n?", "", text)
        text = re.sub(r"\n?```$", "", text.strip())
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        text = text[start:end + 1]
    return json.loads(text)


# --- Stage 1: cleaning ------------------------------------------------------

_CLEAN_SYSTEM = (
    "You clean messy donation logs for a small nonprofit. You map arbitrary "
    "column names onto a fixed schema, normalise values, merge rows that are "
    "clearly the same donor, and flag rows missing required information. "
    "You never invent data. You reply with JSON only."
)


def clean_rows(header, rows):
    """Send the raw table to Claude and return a parsed dict.

    Returns {"records": [...], "change_log": [...]}. The caller re-validates
    every record in code afterwards -- this is a starting point, not gospel.
    """
    payload = {"columns": header, "rows": rows}
    instructions = f"""\
Here is a raw donation log as JSON (column headers plus rows):

{json.dumps(payload, ensure_ascii=False)}

Produce ONE record per unique donor using this exact schema:

  donor_name : full name, tidy capitalisation (e.g. "Robert Smith")
  email      : lowercase email, or null if genuinely absent
  amount     : the gift amount as a NUMBER (convert "$50", "50.00", "fifty" -> 50).
               If a donor gave more than once, use their MOST RECENT gift here.
  date       : most recent gift date as "YYYY-MM-DD", or null if absent/unparseable
  donor_type : "recurring" if the donor appears more than once (by email or name),
               otherwise "first-time"
  status     : "ready"   if donor_name, email and amount are all present
               "flagged" if any of those required fields is missing
               "skipped" for empty or junk rows
  notes      : short human-readable explanation of anything you changed or flagged
               (e.g. "merged 2 rows", "missing email", "recurring: 2 gifts totalling $150")

Rules:
  - Merge rows that are obviously the same donor (same email, or the same name in
    different formats like "Bob Smith" / "Robert Smith"). Do not merge different people.
  - Never fabricate a missing email or amount. Flag instead.
  - Trim whitespace and fix casing on names and emails.

Reply with JSON only, in this shape:
{{"records": [ ...one object per donor... ],
  "change_log": [ "short sentence", "short sentence", ... ]}}
"""
    response = client().messages.create(
        model=config.CLEAN_MODEL,
        max_tokens=8000,
        system=_CLEAN_SYSTEM,
        messages=[{"role": "user", "content": instructions}],
    )
    return _extract_json(_text_of(response))


# --- Stage 2: tone drafting -------------------------------------------------

_NOTE_SYSTEM = (
    "You write the personal note in a nonprofit thank-you letter. "
    "Scale the emotional register to the gift:\n"
    "- Under $25: ONE sentence. Humble, quiet, sincere gratitude. Never gush - "
    "overselling a token gift reads as insincere.\n"
    "- $25-$249: one or two sentences. Warm and specific. This is a real, "
    "considered contribution, not spare change - do not underplay it.\n"
    "- $250-$999: two sentences. Significant. Name what it makes possible.\n"
    "- $1,000+: two to four sentences. Transformational. Tell them plainly that "
    "they changed a life - that a real person's future is permanently different "
    "because of them - and mean it. This is the one case where real emotional "
    "weight is warranted.\n"
    "Recurring donors should hear that their faithfulness is noticed. Never "
    "restate the dollar amount mechanically, never make promises you can't keep, "
    "never add a greeting or sign-off. Output only the note itself, as plain text."
)


def draft_personal_note(record):
    """Return one personalised sentence tuned to this donor, or None on failure.

    Returning None lets the caller fall back to the free offline sentence, so a
    single API hiccup never derails a run.
    """
    donor_type = record.get("donor_type") or "first-time"
    amount = record.get("amount")
    prompt = (
        f"Donor: {record.get('donor_name')}\n"
        f"Relationship: {donor_type}\n"
        f"Most recent gift: {amount}\n"
        f"Notes: {record.get('notes') or ''}\n\n"
        "Write the personalised sentence now."
    )
    try:
        response = client().messages.create(
            model=config.DRAFT_MODEL,
            max_tokens=256,
            # One short sentence - no need for extended reasoning, and leaving
            # thinking on would eat into the small max_tokens budget.
            thinking={"type": "disabled"},
            system=_NOTE_SYSTEM,
            messages=[{"role": "user", "content": prompt}],
        )
        return _text_of(response).strip() or None
    except Exception as exc:  # noqa: BLE001 - degrade gracefully, never crash a run
        print(f"    ! tone drafting failed for {record.get('donor_name')}: {exc}")
        return None
