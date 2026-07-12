# Donation Thank-You Tool
> (This same read me file will also be included with the download btw)

Turns a messy donation spreadsheet into a personal thank-you email for every
donor — and then stops, and lets **you** press Send.

**It is free.** No API key, no account, no internet, no cost. It runs on your own
computer and your donor list never leaves it.

**[▶ See it in action] (https://donorthankingkit.netlify.app/)** — a short animated
walkthrough. (In the event the link doesn't work, Open `how-it-works.html`)

> **Two guides — pick one:**
> - **[GUIDE.md](GUIDE.md)** — the normal walkthrough. Assumes you can install a
>   program and edit a spreadsheet.
> - **[GUIDE-SIMPLE.md](GUIDE-SIMPLE.md)** — for anyone who has never run a
>   program before. Same steps, slower, nothing assumed.
>
> They both lead you to the same destination, but I recommend the ez steps if you're
>   handing this off to someone who's not versatile with electronics.

---

## Quick start

1. **Install Thonny** — free, from [thonny.org](https://thonny.org). It's the only
   thing you install; it includes everything else.
2. **Open the program.** In Thonny: **File → Open →** `donation_tool_standalone.py`.
3. **Press the green ▶ button.** It installs what it needs (a minute, once), makes
   you a blank donor sheet and a letter template, then stops.
4. **Fill in `donations.xlsx`** — three columns: *Donor Name*, *Donation Amount*,
   *Email*. One row per donor. **Save it, then close Excel.**
5. **Press ▶ again.** A page opens in your browser with a button beside each donor.
   Click it, press **Ctrl+V** in your email, read it, press **Send**.

Everything lands in a new `output` folder. Start with `output/REVIEW.md`.

---

## Make it yours

| What | How |
| --- | --- |
| **Your logo** | Save it as `logo.png` in this folder. It appears on every letter. |
| **Your name & tagline** | Near the top of `donation_tool_standalone.py`, set `ORG_NAME` and `ORG_SLOGAN`. |
| **Your wording** | `letter_template.docx` is a normal Word file. Reword and restyle it however you like. |

The letter has fill-in boxes: `[NAME]`, `[AMOUNT]`, `[DATE]`, `[PERSONAL_NOTE]`,
`[ORGANIZATION]`, `[YOUR SLOGAN HERE]`, `[YOUR LOGO HERE]`. Any box you leave
blank is **deleted** from the letter, not mailed to a donor as literal
`[YOUR SLOGAN HERE]` text.

---

## What it will never do

- **Never send an email by itself.** You press Send, every time.
- **Never ask for your email password.** It has no mail account and nowhere to put one.
- **Never guess at missing details.** A donor with no name, amount, or email is
  held back and flagged for you — never half-written, never sent.
- **Never upload your donors.** It works offline. The list stays on your machine.

It also doesn't give legal or tax advice — if your acknowledgments have legal
requirements, check them yourself.

---

## Works with any email

Gmail, Outlook, Yahoo — it doesn't matter. The tool fills in the message and
hands it to whatever email you already use. Browser users get a click-to-open
page; desktop mail users can double-click the `.eml` files in `output/outbox/`,
which arrive fully formatted with the Word letter attached.

---

## For developers

```
donation_tool_standalone.py   # the whole tool in one file - this is what people run
how-it-works.html             # the animated explainer page
GUIDE.md                      # normal step-by-step walkthrough
GUIDE-SIMPLE.md               # the same walkthrough, plain language, nothing assumed

# A modular version of the same pipeline, for reading and modifying:
run.py            # entry point:  python run.py
config.py         # models, paths, placeholder tokens, org name
pipeline/
  offline.py      # FREE MODE: cleaning + the personal sentence, no API at all
  ai.py           # optional Claude calls, key read from the environment
  clean.py        # Stage 1: xlsx -> validated records
  personalize.py  # Stage 2: records + template -> per-donor .docx
  prepare.py      # Stage 3: letters -> .eml outbox, SEND page, REVIEW manifest
samples/
  make_samples.py # generates the blank starter sheet and template
```

**Free by default.** The spreadsheet cleaning and the personalized sentence are
both plain Python (`pipeline/offline.py`). Claude is an *optional* upgrade: set
`USE_CLAUDE_AI = True` (or supply an `ANTHROPIC_API_KEY` for the modular version)
and it writes a fresher sentence per donor and untangles messier spreadsheets.
That is the only difference — and it costs money, so it's off.

**The tone scales with the gift.** Under $25 gets quiet, humble thanks; $1,000+
is told plainly that they changed a life. Recurring donors are recognized
separately from first-time ones. See `_NOTES` in `pipeline/offline.py`.

**Model output is never trusted.** In AI mode, Claude's cleaning result is
re-validated in code (`pipeline/clean.py`) — required fields are enforced there,
so a donor missing an email can never be marked ready to send.

Setup for the modular version: `pip install -r requirements.txt`, then
`python run.py`.

---

Found a bug? Open an issue — I'll do my best to get to it.
