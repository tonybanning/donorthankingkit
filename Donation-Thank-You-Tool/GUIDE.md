# Guide

The normal walkthrough. Assumes you're comfortable installing a program and
editing a spreadsheet.

*(Never run a program before? Read [GUIDE-SIMPLE.md](GUIDE-SIMPLE.md) — same
steps, slower, nothing assumed.)*

---

## What you need

**Thonny** — free, from [thonny.org](https://thonny.org). It bundles Python, so
it's the only install. Nothing else: no API key, no account, no internet after
setup, no cost.

## 1. Open and run

In Thonny: **File → Open →** `donation_tool_standalone.py`, then press **▶** (F5).

First run installs two packages (`openpyxl`, `python-docx`) — a minute, once.
Then it creates two files beside the program and **stops**:

```
Mode: FREE - no key, no internet, no cost.
Created a blank donation sheet for you: donations.xlsx
Created a letter template for you: letter_template.docx
------------------------------------------------------------
Your donation sheet is empty - nothing to thank anyone for yet.
```

That's expected — there's no data yet.

## 2. Fill in the donors

Open **`donations.xlsx`**. It has three columns and one row of placeholder boxes:

| Donor Name | Donation Amount | Email |
| --- | --- | --- |
| Insert name here | Insert donation here | Insert email here |

Type over them, one row per donor. Don't bother tidying — the tool handles it:

- `$50`, `50.00`, `50`, `fifty` all parse to the same number
- `SMITH, ROBERT` → `Robert Smith`; `bob smith` → `Bob Smith`
- the same donor entered twice is **merged** into one letter and marked
  **recurring** (which changes the wording they get)
- add a **Date** column if you want dates in the letters — optional

**Save, then close Excel.** Windows locks open spreadsheets; if Excel still has
the file, the tool can't read it and you'll get a `PermissionError`.

> **Testing?** Use your own email address as the donor email. The prepared email
> comes to you, so you can see exactly what a donor receives.

## 3. Make it yours (optional)

| What | How |
| --- | --- |
| **Logo** | Save it as `logo.png` (or `.jpg`) beside the program. It's inserted into every letter automatically. |
| **Name / slogan** | Near the top of `donation_tool_standalone.py`: `ORG_NAME = "..."` and `ORG_SLOGAN = "..."`. |
| **Wording & styling** | `letter_template.docx` is a plain Word file. Reword it, restyle it, move things around. |

The template's fill-in boxes:

| Box | Filled with | Same on every letter? |
| --- | --- | --- |
| `[NAME]` | donor's name | no |
| `[AMOUNT]` | e.g. `$50.00` | no |
| `[DATE]` | e.g. `January 2, 2026` | no |
| `[PERSONAL_NOTE]` | a sentence matched to the donor | no |
| `[ORGANIZATION]` | your `ORG_NAME` | yes |
| `[YOUR SLOGAN HERE]` | your `ORG_SLOGAN` | yes |
| `[YOUR LOGO HERE]` | your `logo.png` | yes |

**Any box you leave blank is deleted from the letter** — a donor never receives
a literal `[YOUR SLOGAN HERE]`.

## 4. Run it again

Press **▶**. It runs three stages and finishes in seconds:

```
[1/3] Cleaning the donation log...
[2/3] Writing personalised letters...
[3/3] Preparing the outbox...

Done.
  4 ready, 1 flagged, 0 skipped.
  4 email(s) prepared - 0 sent by this tool. Sending is yours.
```

## 5. Send

**`output/SEND.html` opens in your browser automatically.** One row per donor,
with a button.

1. Click **Copy letter & open email** — it copies the formatted letter *and*
   opens your email with the address and subject already filled in.
2. Click into the message box and press **Ctrl+V**. The letter pastes in with
   your logo, fonts, bold and colors intact.
3. Read it, then press **Send**.

Rows tick off as you go and the page remembers your progress if you close it.

> **Why the paste?** A browser compose link can only carry plain text — it
> physically cannot hold a logo or formatting. So the letter travels via the
> clipboard. That one keystroke is what buys you the formatting.

**Using a desktop mail client instead?** `output/outbox/` holds the same emails
as `.eml` files — double-click one and it opens fully formatted with the Word
letter attached. No pasting needed.

## What lands in `output/`

| File | What it is |
| --- | --- |
| `SEND.html` | the click-to-send page — **start here** |
| `REVIEW.md` | who was prepared, who was **held back and why**, totals |
| `letters/` | the finished `.docx` letters, one per donor |
| `outbox/` | the same emails as `.eml` files |
| `cleaned.csv` | the tidied version of your spreadsheet |

**Anyone missing a name, amount, or email is held back and flagged** — never
half-written, never guessed at. `REVIEW.md` tells you who and why; fix the sheet
and run again.

---

## Troubleshooting

**`PermissionError` / "Windows won't let the tool open donations.xlsx"**
Excel still has the file open — the most common problem by far. Close Excel. Also
close the File Explorer **preview pane** if it's showing the file; that alone
locks it. If the folder is in OneDrive, let the sync icon settle. Then run again.

**"Your donation sheet is empty" but you filled it in**
It wasn't saved, or it isn't in the same folder as the program.

**Ctrl+V pastes nothing**
Click *inside* the message box first so the cursor is there. Failing that, open
"Read the letter" on the send page, select it, Ctrl+C, and paste.

**The logo doesn't appear**
Must be named exactly `logo.png` / `logo.jpg` and sit beside the program. It's
embedded in the letters and in the pasted email — a plain-text fallback won't
show it.

**Gmail opened under the wrong account**
It composes from your *default* Google account. Sign out of the others, or make
the nonprofit's account default, before starting a batch.

**Red traceback in the Shell**
Copy it and send it on — that's a bug, not you.

---

## Optional: Claude AI for the wording (costs money)

Free mode picks the personal sentence from a set written for each donor
situation. If you'd rather have each one freshly written:

1. Get an API key at [console.anthropic.com](https://console.anthropic.com).
2. Near the top of the program, set `USE_CLAUDE_AI = True`.
3. Run. It asks for the key and the run costs a few cents.

Everything else is identical. **Leave it `False` and the tool stays free forever.**
