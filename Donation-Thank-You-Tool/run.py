"""Donation Acknowledgment Pipeline - single entry point.

Usage:
    python run.py                       # use the bundled sample files
    python run.py mylog.xlsx template.docx

What it does, in order:
    1. Clean   - read the .xlsx, normalise it with Claude, validate in code
    2. Personalize - fill the Word template + a Claude-written note, one .docx per donor
    3. Prepare - write a ready-to-send .eml per donor + a REVIEW.md manifest

It never sends email. A human reviews the outbox and clicks send.
"""

import csv
import os
import sys

import config
from pipeline import ai, clean, personalize, prepare


def _ensure_samples():
    """Create the blank starter files on first run if they're missing.

    These are fill-in-the-blank templates, not simulated data: the donation
    sheet holds three labelled boxes for the coordinator to type over.
    """
    if os.path.exists(config.INPUT_XLSX) and os.path.exists(config.TEMPLATE_DOCX):
        return
    from samples import make_samples

    os.makedirs(config.SAMPLES_DIR, exist_ok=True)
    if not os.path.exists(config.INPUT_XLSX):
        make_samples.generate_donations(config.INPUT_XLSX)
        print(f"  Created a blank donation sheet: {config.INPUT_XLSX}")
    if not os.path.exists(config.TEMPLATE_DOCX):
        make_samples.generate_template(config.TEMPLATE_DOCX)
        print(f"  Created a letter template: {config.TEMPLATE_DOCX}")


def _write_cleaned_csv(records, path):
    fields = ["donor_name", "email", "amount", "date", "donor_type", "status", "notes"]
    with open(path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for record in records:
            writer.writerow({key: record.get(key, "") for key in fields})


def main(argv):
    config.load_dotenv()

    input_xlsx = argv[0] if len(argv) >= 1 else config.INPUT_XLSX
    template_docx = argv[1] if len(argv) >= 2 else config.TEMPLATE_DOCX

    print("Donation Acknowledgment Pipeline")
    print("=" * 40)
    if ai.enabled():
        print("Mode: Claude AI (an API key was found - this run will cost a few cents)")
    else:
        print("Mode: FREE (no API key, no internet, no cost - nothing is sent anywhere)")
    print("Planned steps:")
    print("  1. Clean       messy .xlsx  -> normalised records")
    print("  2. Personalize records      -> one .docx letter per donor")
    print("  3. Prepare     letters      -> .eml outbox + REVIEW.md (you send)")
    print()

    # Only auto-generate when we're using the bundled sample paths.
    if input_xlsx == config.INPUT_XLSX and template_docx == config.TEMPLATE_DOCX:
        _ensure_samples()

    for path, label in ((input_xlsx, "donation log"), (template_docx, "letter template")):
        if not os.path.exists(path):
            print(f"ERROR: {label} not found: {path}")
            return 1

    # Stop before spending an API call if the sheet is still the blank starter.
    _header, _rows = clean.read_xlsx(input_xlsx)
    if clean.looks_unfilled(_rows):
        print("Your donation sheet hasn't been filled in yet.")
        print(f"  Open: {input_xlsx}")
        print("  Replace the three boxes (name / donation / email) with a real donor,")
        print("  add one row per additional donor, save the file, and run this again.")
        return 0

    os.makedirs(config.OUTPUT_DIR, exist_ok=True)

    # --- Stage 1 ---
    print("[1/3] Cleaning donation log ...")
    records, change_log = clean.run(input_xlsx)
    _write_cleaned_csv(records, config.CLEANED_CSV)
    print("  Change log:")
    for entry in change_log:
        print(f"    - {entry}")
    print(f"  Cleaned table written to {config.CLEANED_CSV}")
    print()

    # --- Stage 2 ---
    print("[2/3] Writing personalised letters ...")
    letters = personalize.run(records, template_docx, config.LETTERS_DIR)
    print(f"  {len(letters)} letter(s) in {config.LETTERS_DIR}")
    print()

    # --- Stage 3 ---
    print("[3/3] Preparing outbox ...")
    prepared = prepare.run(
        records, letters, config.OUTBOX_DIR, config.REVIEW_MD, change_log
    )
    print()

    ready = sum(1 for r in records if r["status"] == "ready")
    flagged = sum(1 for r in records if r["status"] == "flagged")
    skipped = sum(1 for r in records if r["status"] == "skipped")
    print("Done.")
    print(f"  {ready} ready, {flagged} flagged, {skipped} skipped.")
    print(f"  {len(prepared)} email(s) prepared - 0 sent by this tool. Sending is yours.")
    print(f"  Review everything in: {config.REVIEW_MD}")
    return 0


def _explain_locked_file(exc):
    """A PermissionError on Windows nearly always means Excel/Word has the file."""
    name = getattr(exc, "filename", None) or "one of your files"
    print("\n" + "-" * 60)
    print(f"Windows won't let the tool open:  {name}")
    print("\nThis almost always means the file is still open in another program.")
    print("\n  1. Close it in Excel / Word  (save your changes first).")
    print("  2. Close the File Explorer preview pane if it's showing the file.")
    print("  3. If it's in OneDrive, wait for the sync icon to stop spinning.")
    print("  4. Then run it again.")
    print("-" * 60)


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except PermissionError as exc:
        _explain_locked_file(exc)
        sys.exit(1)
