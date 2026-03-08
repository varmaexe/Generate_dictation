# How to Run — Dictation Generator

Hey Shravya! 👋 Follow these steps just once and you're all set forever.

---

## Step 1 — Open Terminal

Press **Command (⌘) + Space**, type **Terminal**, and press Enter.

A black/white window will open. Don't worry, we'll only use it for a few minutes!

---

## Step 2 — Install Python

Copy and paste this into Terminal, then press Enter:

```
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

It will ask for your Mac password — type it (nothing will show while typing, that's normal) and press Enter.

This may take a few minutes. Wait until you see the `$` prompt again.

Then paste this and press Enter:

```
brew install python
```

---

## Step 3 — Install the required library

Paste this and press Enter:

```
pip3 install openpyxl
```

---

## Step 4 — Run the tool

Paste this and press Enter (update the path if you saved the file somewhere else):

```
python3 ~/Desktop/generate_dictation.py
```

The tool will guide you through everything with simple questions.
Just press **Enter** to accept the defaults, or type a number and press Enter.

When it's done, your Excel file will open automatically! 🎉

---

## Every time after that

You only need Step 4 from now on. Just open Terminal and run:

```
python3 ~/Desktop/generate_dictation.py
```

---

## Trouble?

Ask Sai 😄
