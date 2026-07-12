# Simple Guide

**For anyone who has never run a program before. Nothing is assumed.**

*(If you're comfortable with computers, read [GUIDE.md](GUIDE.md) instead —
it's the same thing, shorter.)*

---

This tool writes thank-you letters to your donors for you.

You type your donors into a list. The tool writes each person a letter. Then it
opens each email for you, ready to go. **You just press Send.**

**It never sends anything by itself. You are always the one who presses Send.**

It is free. You do not need to pay for anything, and you do not need an account.

---

# Part 1 — Get it ready (you only do this once)

## Step 1. Get the free program

1. Open your web browser.
2. Go to **thonny.org**
3. Click the big **Windows** button to download it.
4. Open the file you just downloaded.
5. Keep clicking **Next** until it says it is done.

This is a free program called Thonny. It is what runs the tool. **It is the only
thing you install.**

## Step 2. Open the tool

1. Open **Thonny** (look for it in your Start menu).
2. At the top, click **File**, then **Open**.
3. Find the folder this guide is in.
4. Click the file named **donation_tool_standalone.py**
5. Click **Open**.

Some writing appears on the screen. **Do not change it.** That is fine.

## Step 3. Press the Play button

1. Look at the top of Thonny for a round green **▶** button.
2. Click it.
3. **Wait.** The first time, it takes a minute or two. Do not close anything.
4. When it stops, look at the bottom of the screen. It will say something like:

   > Your donation sheet is empty.

**This is good. Nothing is broken.** It just means you have not typed in your
donors yet. That is Part 2.

---

# Part 2 — Type in your donors

## Step 4. Open the donor list

The tool just made you a new file, in the same folder.

1. Find the file called **donations.xlsx**
2. Double-click it. It opens in Excel.

You will see three boxes:

| Donor Name | Donation Amount | Email |
| --- | --- | --- |
| Insert name here | Insert donation here | Insert email here |

## Step 5. Type over the boxes

Type over each box with a real person, like this:

| Donor Name | Donation Amount | Email |
| --- | --- | --- |
| Jane Doe | 50 | jane@example.org |

Then keep going down the page. **One line for each person who gave money.**

**Do not worry about being neat.** The tool fixes messy typing by itself:

- `$50` or `50.00` or `50` or even `fifty` — all fine.
- `BOB SMITH` or `bob smith` — it fixes the capital letters.
- Put the same person in twice? It notices, sends them only one letter, and
  treats them as a repeat donor.

> **Testing it for the first time?** Put **your own email address** in as the
> donor email. Then the email comes to you, and you can see exactly what a donor
> would get — with no risk at all.

## Step 6. Save it, then CLOSE Excel

1. Press **Ctrl + S** to save.
2. **Now close Excel completely.**

**This part matters.** If you leave the file open in Excel, the tool cannot read
it, and you will get an error. Close it.

---

# Part 3 — Make the letters and send them

## Step 7. Press the Play button again

Go back to Thonny and click the round green **▶** button. Wait a few seconds.

## Step 8. Send the emails

**A page pops up in your web browser by itself.** It shows your donors, one per
line, with a blue button next to each one.

For each person, do these **three things**:

**1.** Click the blue **Copy letter & open email** button next to their name.

Your email opens in a new tab. Their address is already filled in. The subject is
already filled in.

**2.** Click inside the **big empty message box**, and press **Ctrl + V**.

The letter appears — with your logo, your fonts, and your colors.

**3. Read it.** Then press **Send**.

Then go back to the page and do the next person. That is the whole thing.

> **What is Ctrl + V?**
> Hold down the **Ctrl** key. While still holding it, tap the **V** key.
> That pastes the letter in. The blue button already did the copying for you —
> you just have to paste.

- Want to see the letter first? Click **Read the letter** under their name.
- The page ticks off each person as you go, so you do not lose your place.
- Closed the page by accident? Open the **output** folder and double-click
  **SEND.html**.

---

# Making the letters your own (optional)

## Your logo

Put a picture file named exactly **logo.png** in the same folder as the tool.
It will appear at the top of every letter. Press ▶ again to rebuild them.

## Your organization's name and slogan

In Thonny, near the top of the program, find these two lines and type between
the quote marks:

```
ORG_NAME = "Your Organization Name"
ORG_SLOGAN = ""
```

Leave the slogan empty if you do not have one — that line is simply removed from
the letter.

## Your wording

**letter_template.docx** is a normal Word document. Open it and change the
wording, the font, the colors — anything. Just leave the boxes in square brackets
where you want the tool to fill something in:

| Box | Becomes |
| --- | --- |
| `[NAME]` | the donor's name |
| `[AMOUNT]` | the donation, e.g. `$50.00` |
| `[PERSONAL_NOTE]` | a warm sentence matched to that donor |
| `[ORGANIZATION]` | your organization's name |
| `[YOUR SLOGAN HERE]` | your slogan |
| `[YOUR LOGO HERE]` | your logo picture |

If you leave a box blank, that line is **deleted** — a donor will never receive a
letter that says "[YOUR SLOGAN HERE]".

---

# If something goes wrong

## "Permission denied" / "Windows won't let the tool open donations.xlsx"

**Excel still has the file open.** This is the most common problem.

1. Close Excel.
2. If File Explorer is showing a **preview panel** on the right with the file in
   it, close that too — it can hold the file open.
3. Press ▶ again.

## It says my donor list is empty, but I filled it in

You did not save it. Open **donations.xlsx**, press **Ctrl + S**, close Excel,
and press ▶ again.

## I pressed Ctrl + V and nothing appeared

Click **inside the big message box** first, so the cursor is blinking there. Then
press **Ctrl + V**.

Still nothing? Go back to the page, click **Read the letter** under that person,
select the letter with your mouse, press **Ctrl + C**, then paste it in your
email.

## Someone did not get a letter

The tool skips anyone missing their **name**, **donation amount**, or **email**.
It will not guess.

To see who was skipped and why: open the **output** folder and read
**REVIEW.md**. Fix that person in your list, save, close Excel, press ▶ again.

## The logo did not appear

The file must be named exactly **logo.png** and sit in the same folder as the
tool.

## My email opened under the wrong account

You are signed in to more than one account. It uses whichever is your main one.
Sign out of the others first, then try again.

## There is red writing at the bottom of Thonny

Something went wrong. Select the red writing with your mouse, press **Ctrl + C**
to copy it, and send it to whoever set this up for you.

---

# Things to remember

- **The tool never sends an email by itself.** You press Send, every time.
- **It never asks for your email password.** It does not need one.
- **Always read a letter before you send it.**
- **Close Excel before you press ▶.** This is the most common mistake.
- It is free, and it works offline. Your donor list never leaves your computer.
