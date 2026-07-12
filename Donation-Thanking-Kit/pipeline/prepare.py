"""Stage 3 - Prepare to send (a human sends).

For each ready donor we write a local .eml file: recipient, subject, the letter
as the body, and the .docx attached. A coordinator double-clicks it, it opens
in their normal mail client fully pre-filled, and *they* click send.

The tool never transmits anything. There is deliberately no SMTP here.
"""

import os
from email.generator import BytesGenerator
from email.message import EmailMessage

import config

_DOCX_MIME = (
    "application",
    "vnd.openxmlformats-officedocument.wordprocessingml.document",
)


def _write_eml(letter, eml_path):
    record = letter["record"]
    message = EmailMessage()
    message["To"] = record["email"]
    message["From"] = config.ORG_NAME  # a label, not an address - no credentials involved
    message["Subject"] = config.SUBJECT_TEMPLATE.format(
        name=record["donor_name"], org=config.ORG_NAME
    )
    message.set_content(letter["body_text"])

    with open(letter["letter_path"], "rb") as handle:
        message.add_attachment(
            handle.read(),
            maintype=_DOCX_MIME[0],
            subtype=_DOCX_MIME[1],
            filename=os.path.basename(letter["letter_path"]),
        )

    with open(eml_path, "wb") as handle:
        BytesGenerator(handle).flatten(message)


def _write_review(records, prepared, review_path, change_log):
    ready = [r for r in records if r["status"] == "ready"]
    flagged = [r for r in records if r["status"] == "flagged"]
    skipped = [r for r in records if r["status"] == "skipped"]

    lines = []
    lines.append("# Review before sending\n")
    lines.append(
        f"**{len(prepared)} ready, {len(flagged)} flagged, {len(skipped)} skipped "
        "- 0 sent by this tool. Sending is yours.**\n"
    )
    lines.append(
        "Each prepared email is a `.eml` file in `output/outbox/`. Double-click one "
        "to open it in your mail program with everything pre-filled, review it, then "
        "click Send. Nothing leaves your computer until you do.\n"
    )

    lines.append("## Prepared emails\n")
    if prepared:
        lines.append("| Donor | Email | Amount | File |")
        lines.append("| --- | --- | --- | --- |")
        for item in prepared:
            record = item["record"]
            lines.append(
                f"| {record['donor_name']} | {record['email']} | "
                f"${float(record['amount']):,.2f} | {os.path.basename(item['eml_path'])} |"
            )
    else:
        lines.append("_None._")
    lines.append("")

    lines.append("## Held back (review by hand)\n")
    held = flagged + skipped
    if held:
        lines.append("| Donor | Status | Reason |")
        lines.append("| --- | --- | --- |")
        for record in held:
            name = record.get("donor_name") or "(no name)"
            lines.append(f"| {name} | {record['status']} | {record.get('notes') or ''} |")
    else:
        lines.append("_None - every row was ready._")
    lines.append("")

    lines.append("## What the cleaning pass did\n")
    for entry in change_log:
        lines.append(f"- {entry}")
    lines.append("")

    with open(review_path, "w", encoding="utf-8") as handle:
        handle.write("\n".join(lines))


def run(records, letters, outbox_dir, review_path, change_log):
    """Write .eml files and the REVIEW.md manifest. Returns the prepared list."""
    os.makedirs(outbox_dir, exist_ok=True)
    prepared = []
    for letter in letters:
        eml_path = os.path.join(
            outbox_dir,
            os.path.splitext(os.path.basename(letter["letter_path"]))[0] + ".eml",
        )
        _write_eml(letter, eml_path)
        letter["eml_path"] = eml_path
        prepared.append(letter)
        print(f"    - outbox: {os.path.basename(eml_path)}")

    _write_review(records, prepared, review_path, change_log)
    return prepared
