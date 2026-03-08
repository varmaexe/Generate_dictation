# generate_dictation.py

A Python script that generates Excel files with repeated Dictation-style data blocks — identical structure and formulas, different random values every run.

---

## Requirements

```bash
pip install openpyxl
```

---

## Quick Start

```bash
python generate_dictation.py --blocks 4
```

---

## Flags

| Flag | Type | Default | Description |
|------|------|---------|-------------|
| `--blocks` | int | `4` | Total number of data blocks to generate |
| `--sheets` | int | `1` | Number of sheets to spread blocks across |
| `--min` | int | `1` | Minimum value for single-digit random integers |
| `--max` | int | `9` | Maximum value for single-digit random integers |
| `--double` | flag | off | Enable alternating row mode (see below) |
| `--output` | str | `dictation_output.xlsx` | Output filename |

---

## Examples

### Basic

```bash
# 4 blocks on 1 sheet
python generate_dictation.py --blocks 4

# 10 blocks on 1 sheet, custom filename
python generate_dictation.py --blocks 10 --output my_dictation.xlsx
```

### Multiple Sheets

```bash
# 15 blocks spread across 3 sheets (5 per sheet)
python generate_dictation.py --blocks 15 --sheets 3

# 8 blocks across 4 sheets (2 per sheet)
python generate_dictation.py --blocks 8 --sheets 4
```

### Custom Value Range

```bash
# Single-digit values between 3 and 7 only
python generate_dictation.py --blocks 4 --min 3 --max 7
```

### Double Mode

Enables alternating row values within each group (A, B, C):

- **Odd rows** (row 1, 3, 5) → single digit using `--min` / `--max`
- **Even rows** (row 2, 4) → double digit, always in range `[10, 99]`

The pattern resets at the start of each group.

```bash
# Default single-digit range (1-9) for odd rows
python generate_dictation.py --blocks 4 --double

# Custom single-digit range for odd rows
python generate_dictation.py --blocks 4 --double --min 3 --max 7
```

**Example output (Group A, one block):**

| | Col 1 | Col 2 | Col 3 | Col 4 | Col 5 |
|---|---|---|---|---|---|
| Row 1 (odd) | 5 | 2 | 8 | 1 | 6 |
| Row 2 (even) | 43 | 71 | 28 | 56 | 19 |
| Row 3 (odd) | 3 | 7 | 4 | 9 | 2 |
| Row 4 (even) | 82 | 35 | 64 | 47 | 91 |
| Row 5 (odd) | 6 | 1 | 5 | 3 | 8 |

### All Flags Together

```bash
python generate_dictation.py --blocks 15 --sheets 3 --double --min 2 --max 8 --output final.xlsx
```

---

## Block Structure

Each block follows this layout:

```
Row  1        : Header — "Column" label + column numbers 1–5
Rows 2–6      : Group A — 5 rows of random data
Rows 7–11     : Group B — 5 rows of random data
Rows 12–16    : Group C — 5 rows of random data
Rows 17–30    : Summary formulas (14 rows)
```

### Summary Rows

| Row | Label | Formula Range | Meaning |
|-----|-------|--------------|---------|
| 17 | A | rows 2–6 | Group A total |
| 18 | B | rows 7–11 | Group B total |
| 19 | C | rows 12–16 | Group C total |
| 20 | A-C | rows 2–16 | All groups total |
| 21 | A-6 | rows 2–7 | A + 1 row of B |
| 22 | A-7 | rows 2–8 | A + 2 rows of B |
| 23 | A-8 | rows 2–9 | A + 3 rows of B |
| 24 | A-9 | rows 2–10 | A + 4 rows of B |
| 25 | AB | rows 2–11 | All of A + all of B |
| 26 | BC | rows 7–16 | All of B + all of C |
| 27 | B-6 | rows 7–12 | B + 1 row of C |
| 28 | B-7 | rows 7–13 | B + 2 rows of C |
| 29 | B-8 | rows 7–14 | B + 3 rows of C |
| 30 | B-9 | rows 7–15 | B + 4 rows of C |

All summary rows use `=SUM(...)` formulas referencing their own block's columns — no hardcoded values.

---

## Notes

- Every run generates **new random values** while keeping the structure identical
- `--min` / `--max` must be in single-digit range (`1–9`) when using `--double`
- When using `--sheets`, blocks are distributed evenly using ceiling division
- All formula cells are verified to have **zero errors** (`#REF!`, `#DIV/0!`, etc.)
