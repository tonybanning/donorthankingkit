"""Central configuration for the Donation Acknowledgment Pipeline.

Everything a non-technical maintainer might want to change lives here:
the organisation's name, which Claude models to call, where files live,
and the placeholder tokens used in the Word template.
"""

import os

# --- Organisation details (safe to edit) ------------------------------------
# These fill the [ORGANIZATION] and [YOUR SLOGAN HERE] boxes in the letter.
# No email address or password is ever needed here.
ORG_NAME = "Your Organization Name"
ORG_SLOGAN = ""  # optional tagline, e.g. "Feeding our neighbours since 1998"

# Drop an image with one of these names beside the tool and it will be inserted
# wherever [YOUR LOGO HERE] appears in the template. Optional.
LOGO_FILENAMES = ("logo.png", "logo.jpg", "logo.jpeg")
LOGO_WIDTH_INCHES = 1.5

# Subject line for the prepared emails. {name} and {org} are filled in.
SUBJECT_TEMPLATE = "Thank you for your donation to {org}"

# --- Free mode vs AI mode ---------------------------------------------------
# The tool runs completely FREE by default: no API key, no internet, no cost.
# Cleaning and the personal sentence are done in plain Python (pipeline/offline.py).
#
# If an ANTHROPIC_API_KEY happens to be present, the tool automatically uses
# Claude instead, which writes a fresher sentence and copes with messier
# spreadsheets. That is the only difference.
#
#   "auto" - use Claude only if an API key is set (recommended)
#   "off"  - never use the API, even if a key is set
#   "on"   - require the API (errors if no key)
USE_AI = "auto"

# --- Claude models (only used when USE_AI applies) ---------------------------
# Kept as constants so they are easy to change in one place.
# Verify current IDs at: https://docs.claude.com/en/docs/about-claude/models/overview
#
# Cleaning/dedup is a cheap, structured pass  -> Haiku (fast, low cost).
# Letter tone drafting benefits from stronger writing -> Sonnet.
CLEAN_MODEL = "claude-haiku-4-5"
DRAFT_MODEL = "claude-sonnet-5"

# --- Word-template placeholder tokens ---------------------------------------
# The organisation writes its letter normally and drops these tokens in where
# it wants each value to appear. Documented in the README.
#
# Filled differently for every donor:
PLACEHOLDER_NAME = "[NAME]"
PLACEHOLDER_AMOUNT = "[AMOUNT]"
PLACEHOLDER_DATE = "[DATE]"
PLACEHOLDER_NOTE = "[PERSONAL_NOTE]"  # where Claude's personalised sentence goes

# The same on every letter (your letterhead). Filled from the settings above,
# or just type over them directly in Word:
PLACEHOLDER_ORG = "[ORGANIZATION]"
PLACEHOLDER_SLOGAN = "[YOUR SLOGAN HERE]"
PLACEHOLDER_LOGO = "[YOUR LOGO HERE]"  # replaced by your logo image, if you add one


def find_logo(base_dir=None):
    """Return the path to the org's logo image, or None if they didn't add one."""
    base_dir = base_dir or BASE_DIR
    for filename in LOGO_FILENAMES:
        candidate = os.path.join(base_dir, filename)
        if os.path.exists(candidate):
            return candidate
    return None

# --- Paths ------------------------------------------------------------------
# Resolve everything relative to this file so the tool works from any folder.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SAMPLES_DIR = os.path.join(BASE_DIR, "samples")
INPUT_XLSX = os.path.join(SAMPLES_DIR, "donations.xlsx")
TEMPLATE_DOCX = os.path.join(SAMPLES_DIR, "letter_template.docx")

OUTPUT_DIR = os.path.join(BASE_DIR, "output")
LETTERS_DIR = os.path.join(OUTPUT_DIR, "letters")
OUTBOX_DIR = os.path.join(OUTPUT_DIR, "outbox")
CLEANED_CSV = os.path.join(OUTPUT_DIR, "cleaned.csv")
REVIEW_MD = os.path.join(OUTPUT_DIR, "REVIEW.md")

# Fields a row MUST have to become a ready-to-send acknowledgment.
REQUIRED_FIELDS = ("donor_name", "email", "amount")


def load_dotenv(path=None):
    """Minimal .env loader (avoids an extra dependency).

    Reads KEY=VALUE lines from a .env file in the project root and puts them
    into the environment *without* overriding anything already set. Lines that
    are blank or start with '#' are ignored.
    """
    path = path or os.path.join(BASE_DIR, ".env")
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value
