# ✨ What's New — Dictation Generator

A full rundown of every improvement made to the tool, written so anyone can understand it.

---

## 📋 New Modes

| # | Mode | What it generates |
|---|------|-------------------|
| 1 | All single digit | Numbers from 1–9 (or your custom range) |
| 2 | All double digit | Numbers from 10–99 |
| 3 | All triple digit | Numbers from 100–999 |
| 4 | Alternating (single + double) | Odd rows → 1–9 · Even rows → 10–99 |
| 5 | **New!** Alternating (double + triple) | Odd rows → 10–99 · Even rows → 100–999 |

> Modes are now ordered logically — all "solid" modes first, then "alternating" modes.

---

## 📁 Smarter File Naming

- **Timestamp** — every file now includes the date and time it was made
  e.g. `dictation_2026-03-11_1435_seed42156.xlsx`

- **Seed in filename** — the random seed is baked into the filename so you never lose it, even after closing the terminal

- **Sanitization** — if you type a filename with spaces or special characters, they're automatically cleaned up and you're warned

- **Collision warning** — if a file with that name already exists, the tool asks before overwriting (defaults to *No*)

---

## 📊 Excel Sheet Improvements

### Structure & Layout
- **Block numbers** — headers now show "Block 1", "Block 2", etc. instead of the generic "Column"
- **Row heights** — data rows are taller (18 px), summary rows slightly taller (16 px), header tallest (20 px) — easier to read at a glance
- **Sheet names** — tabs are now labelled "Dictation 1", "Dictation 2" instead of "Sheet1", "Sheet2"

### Visual Polish
- **Borders** — every data cell, label cell, and summary cell now has a clean thin border
- **Alternating row shading** — in modes 4 and 5, even rows use a slightly deeper shade of their group colour so you can instantly spot which rows are which digit type
- **Gap column** — the separator column between blocks is now filled with a soft grey instead of plain white

### Print Ready
- **Landscape orientation** — set automatically, no manual adjustment needed
- **Fit to width** — all blocks fit on one page width when printing
- **Print area** — only the actual data is included in the print range, no blank columns

### Seed Preserved in File
- A small grey note is placed below the summary rows on every sheet:
  `Seed: 42156  ·  Generated: 2026-03-11_1435`
  So even if you rename the file, the seed travels with it.

---

## 💻 Terminal Experience

### While Answering Questions
- **Coloured prompts** — question text is bright, the default value is highlighted in yellow, hints are dimmed
- **Coloured warnings** — ⚠ messages appear in yellow so they stand out
- **Section dividers** — inputs are grouped into clear sections: Setup / Mode / Range / Output
- **Mode confirmation** — after picking a mode, a green `✓` echoes back exactly what you selected
- **Ctrl+C hint** — a subtle reminder is shown in the header

### During Generation
- **Per-sheet progress** — shows `→ Writing sheet 1 of 2 ...  ✓` live as each sheet is written

### After Generation
- **Full file path** — the done box shows the complete path to the file, so you always know exactly where it was saved
- **Styled done box** — a green bordered box with the filename, full path, and a summary (blocks · sheets · seed)

### Loop
- **"Generate another?" prompt** — after finishing, asks if you want to make another file without restarting the tool. The banner only prints once.
- **Clean exit** — pressing Ctrl+C at any point exits gracefully with a tidy message instead of a Python error

---

## 🛠 Fixes

- **File open removed** — the tool no longer auto-opens the Excel file after saving
- **Seed reproducibility** — use the same seed number to generate the exact same file again at any time

---

*Made with love for Shravya 💙*
