"""
generate_dictation.py
Generates an Excel file with N blocks of the Dictation structure.

Usage:
    python generate_dictation.py --blocks 4
    python generate_dictation.py --blocks 10 --output my_file.xlsx
    python generate_dictation.py --blocks 4 --min 1 --max 9
    python generate_dictation.py --blocks 4 --sheets 3       # spread across multiple sheets
    python generate_dictation.py --blocks 4 --double         # odd rows=single digit, even rows=double digit
    python generate_dictation.py --blocks 4 --double --min 2 --max 8  # custom single-digit range
    python generate_dictation.py --blocks 4 --all-double     # all rows=double digit (10–99)
    python generate_dictation.py --blocks 4 --double-triple  # odd rows=double digit, even rows=triple digit
"""

import random
import subprocess
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Config ──────────────────────────────────────────────────────────────────

DATA_ROWS    = 5          # rows per group (A, B, C)
NUM_GROUPS   = 3          # groups per block (A, B, C)
NUM_COLS     = 5          # data columns (1–5)
BLOCK_GAP    = 1          # empty columns between blocks
HEADER_ROWS  = 1          # header row at top of block

# Row labels for each group
GROUP_LABELS = ["A", "B", "C "]

# Summary row definitions: (label, start_group_offset, end_row_offset_from_data_start)
# Each entry: (label, start_row_in_block, end_row_in_block)  -- 1-indexed within block data area
SUMMARY_DEFS = [
    ("A",    1,  5),    # Group A only
    ("B",    6,  10),   # Group B only
    ("C ",   11, 15),   # Group C only
    ("A-C",  1,  15),   # All groups
    ("A-6",  1,  6),    # A + 1 row of B
    ("A-7",  1,  7),    # A + 2 rows of B
    ("A-8",  1,  8),    # A + 3 rows of B
    ("A-9",  1,  9),    # A + 4 rows of B
    ("AB",   1,  10),   # A + all of B
    ("BC",   6,  15),   # B + all of C
    ("B-6",  6,  11),   # B + 1 row of C
    ("B-7",  6,  12),   # B + 2 rows of C
    ("B-8",  6,  13),   # B + 3 rows of C
    ("B-9",  6,  14),   # B + 4 rows of C
]

# Colors
COLOR_HEADER_BG  = "4472C4"   # blue
COLOR_HEADER_FG  = "FFFFFF"   # white
COLOR_GROUP_A    = "DDEEFF"
COLOR_GROUP_B    = "EEFFDD"
COLOR_GROUP_C    = "FFEECC"
COLOR_SUMMARY_BG = "F2F2F2"
COLOR_LABEL_BG   = "D9D9D9"


# ── Helpers ─────────────────────────────────────────────────────────────────

def col_letter(col: int) -> str:
    return get_column_letter(col)


def make_fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", start_color=hex_color, fgColor=hex_color)


def thin_border() -> Border:
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)


def write_block(ws, block_index: int, start_col: int, val_min: int, val_max: int, double_mode: bool = False, all_double: bool = False, all_triple: bool = False, double_triple: bool = False):
    """
    Write one full block at the given starting column.
    Block layout (rows):
      Row 1        : Header  ("Column", 1, 2, 3, 4, 5)
      Rows 2–6     : Group A data
      Rows 7–11    : Group B data
      Rows 12–16   : Group C data
      Rows 17–30   : Summary formulas
    """

    DATA_START_ROW  = 2                                   # first data row (Excel row)
    SUMMARY_START_R = DATA_START_ROW + NUM_GROUPS * DATA_ROWS  # row 17

    label_col = start_col
    data_col_start = start_col + 1   # columns for the 5 data columns

    # ── Header row ──────────────────────────────────────────────────────────
    hdr_fill  = make_fill(COLOR_HEADER_BG)
    hdr_font  = Font(bold=True, color=COLOR_HEADER_FG, name="Arial", size=10)
    hdr_align = Alignment(horizontal="center")

    cell = ws.cell(row=1, column=label_col, value="Column")
    cell.font = hdr_font; cell.fill = hdr_fill; cell.alignment = hdr_align

    for i, col_num in enumerate(range(1, NUM_COLS + 1)):
        cell = ws.cell(row=1, column=data_col_start + i, value=col_num)
        cell.font = hdr_font; cell.fill = hdr_fill; cell.alignment = hdr_align

    # ── Data rows (Groups A, B, C) ───────────────────────────────────────────
    group_colors = [COLOR_GROUP_A, COLOR_GROUP_B, COLOR_GROUP_C]

    for g_idx, (g_label, g_color) in enumerate(zip(GROUP_LABELS, group_colors)):
        g_fill      = make_fill(g_color)
        label_fill  = make_fill(g_color)
        label_font  = Font(bold=True, name="Arial", size=10)
        data_font   = Font(name="Arial", size=10)
        center      = Alignment(horizontal="center")

        for r in range(DATA_ROWS):
            excel_row = DATA_START_ROW + g_idx * DATA_ROWS + r

            # Label column (only on first row of each group)
            lbl_cell = ws.cell(row=excel_row, column=label_col)
            if r == 0:
                lbl_cell.value = g_label
                lbl_cell.font  = label_font
                lbl_cell.fill  = label_fill
                lbl_cell.alignment = Alignment(horizontal="left")
            else:
                lbl_cell.fill = label_fill

            # Data columns
            for c in range(NUM_COLS):
                # Determine value based on mode
                if all_triple:
                    val = random.randint(100, 999)
                elif all_double:
                    val = random.randint(10, 99)
                elif double_triple and (r + 1) % 2 == 0:
                    # alternating: odd rows=double digit, even rows=triple digit
                    val = random.randint(100, 999)
                elif double_triple:
                    val = random.randint(10, 99)
                elif double_mode and (r + 1) % 2 == 0:
                    # alternating: odd rows=single digit, even rows=double digit
                    val = random.randint(10, 99)
                else:
                    val = random.randint(val_min, val_max)
                cell = ws.cell(row=excel_row, column=data_col_start + c, value=val)
                cell.font      = data_font
                cell.fill      = g_fill
                cell.alignment = center

    # ── Summary rows ─────────────────────────────────────────────────────────
    sum_label_font  = Font(bold=True, name="Arial", size=10)
    sum_data_font   = Font(name="Arial", size=10)
    sum_fill        = make_fill(COLOR_SUMMARY_BG)
    lbl_fill        = make_fill(COLOR_LABEL_BG)
    center          = Alignment(horizontal="center")

    for s_idx, (s_label, rel_start, rel_end) in enumerate(SUMMARY_DEFS):
        excel_row   = SUMMARY_START_R + s_idx
        abs_start   = DATA_START_ROW + rel_start - 1     # Excel absolute start row
        abs_end     = DATA_START_ROW + rel_end   - 1     # Excel absolute end row

        # Label cell
        lbl = ws.cell(row=excel_row, column=label_col, value=s_label)
        lbl.font  = sum_label_font
        lbl.fill  = lbl_fill
        lbl.alignment = Alignment(horizontal="left")

        # Formula cells
        for c in range(NUM_COLS):
            excel_col  = data_col_start + c
            col_ltr    = col_letter(excel_col)
            formula    = f"=SUM({col_ltr}{abs_start}:{col_ltr}{abs_end})"
            cell       = ws.cell(row=excel_row, column=excel_col, value=formula)
            cell.font  = sum_data_font
            cell.fill  = sum_fill
            cell.alignment = center

    # ── Column widths ─────────────────────────────────────────────────────────
    ws.column_dimensions[col_letter(label_col)].width = 7
    for c in range(NUM_COLS):
        ws.column_dimensions[col_letter(data_col_start + c)].width = 6


# ── Interactive prompt helpers ───────────────────────────────────────────────

def ask_int(question: str, default: int, min_val: int = None, max_val: int = None) -> int:
    while True:
        raw = input(f"  {question} (default: {default}): ").strip()
        if raw == "":
            return default
        try:
            val = int(raw)
            if min_val is not None and val < min_val:
                print(f"    ⚠  Please enter a number >= {min_val}")
                continue
            if max_val is not None and val > max_val:
                print(f"    ⚠  Please enter a number <= {max_val}")
                continue
            return val
        except ValueError:
            print(f"    ⚠  Please enter a whole number")


def ask_yes_no(question: str, default: bool) -> bool:
    default_hint = "Y/n" if default else "y/N"
    while True:
        raw = input(f"  {question} ({default_hint}): ").strip().lower()
        if raw == "":
            return default
        if raw in ("y", "yes"):
            return True
        if raw in ("n", "no"):
            return False
        print("    ⚠  Please enter y or n")


def ask_str(question: str, default: str) -> str:
    raw = input(f"  {question} (default: {default}): ").strip()
    return raw if raw else default


# ── Main ────────────────────────────────────────────────────────────────────

def main():
    # colour helpers
    P  = "\033[35m"   # magenta/purple
    PL = "\033[95m"   # light magenta
    C  = "\033[36m"   # cyan
    Y  = "\033[93m"   # yellow
    W  = "\033[97m"   # bright white
    R  = "\033[0m"    # reset

    print()
    print(f"{P}  ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿{R}")
    print(f"{PL}  ♡  Dedicated with love to                 ♡{R}")
    print()
    print(f"  {C} ███████╗{Y}██╗  ██╗{C}██████╗  {Y} █████╗ {C}██╗   ██╗{Y}██╗   ██╗{C} █████╗ {R}")
    print(f"  {C} ██╔════╝{Y}██║  ██║{C}██╔══██╗ {Y}██╔══██╗{C}██║   ██║{Y}╚██╗ ██╔╝{C}██╔══██╗{R}")
    print(f"  {C} ███████╗{Y}███████║{C}██████╔╝ {Y}███████║{C}██║   ██║{Y} ╚████╔╝ {C}███████║{R}")
    print(f"  {C} ╚════██║{Y}██╔══██║{C}██╔══██╗ {Y}██╔══██║{C}╚██╗ ██╔╝{Y}  ╚██╔╝  {C}██╔══██║{R}")
    print(f"  {C} ███████║{Y}██║  ██║{C}██║  ██║ {Y}██║  ██║{C} ╚████╔╝ {Y}   ██║   {C}██║  ██║{R}")
    print(f"  {C} ╚══════╝{Y}╚═╝  ╚═╝{C}╚═╝  ╚═╝ {Y}╚═╝  ╚═╝{C}  ╚═══╝  {Y}   ╚═╝   {C}╚═╝  ╚═╝{R}")
    print()
    print(f"{P}                    {W}✨ Hope you enjoy it! ✨{R}")
    print(f"{P}  ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿ ✿{R}")
    print()
    print("╔══════════════════════════════════════════╗")
    print("║      Dictation Excel Generator           ║")
    print("╚══════════════════════════════════════════╝")
    print("  Press Enter to accept the default value.\n")

    # ── Collect inputs ───────────────────────────────────────────────────────
    blocks = ask_int("How many blocks do you want?", default=4, min_val=1)
    sheets = ask_int("How many sheets to spread blocks across?", default=1, min_val=1, max_val=blocks)

    print("  Mode options:")
    print("    1 = All single digit (1–9)")
    print("    2 = Alternating (odd rows = single digit, even rows = double digit)")
    print("    3 = All double digit (10–99)")
    print("    4 = All triple digit (100–999)")
    print("    5 = Alternating (odd rows = double digit, even rows = triple digit)")
    mode_choice = ask_int("Select mode", default=1, min_val=1, max_val=5)

    double        = (mode_choice == 2)
    all_double    = (mode_choice == 3)
    all_triple    = (mode_choice == 4)
    double_triple = (mode_choice == 5)

    if double:
        print("  (Single-digit range applies to odd rows only. Double-digit is always 10–99.)")
        val_min = ask_int("  Min value for single-digit rows", default=1, min_val=1, max_val=9)
        val_max = ask_int("  Max value for single-digit rows", default=9, min_val=val_min, max_val=9)
    elif all_double:
        val_min = 10
        val_max = 99
    elif all_triple:
        val_min = 100
        val_max = 999
    elif double_triple:
        val_min = 10
        val_max = 999
    else:
        val_min = ask_int("Min random value", default=1, min_val=1)
        val_max = ask_int("Max random value", default=9, min_val=val_min)

    today = date.today().strftime("%Y-%m-%d")
    output = ask_str("Output filename", default=f"dictation_{today}.xlsx")
    if not output.endswith(".xlsx"):
        output += ".xlsx"

    # ── Confirm ──────────────────────────────────────────────────────────────
    print()
    print("  ┌─────────────────────────────────────────┐")
    print("  │  Summary                                │")
    print("  ├─────────────────────────────────────────┤")
    print(f"  │  Blocks       : {blocks:<24}│")
    print(f"  │  Sheets       : {sheets:<24}│")
    if double:
        print(f"  │  Mode         : {'Alternating (single + double)':<24}│")
        print(f"  │  Single range : {f'{val_min}–{val_max}':<24}│")
        print(f"  │  Double range : {'10–99':<24}│")
    elif all_double:
        print(f"  │  Mode         : {'All double digit':<24}│")
        print(f"  │  Value range  : {'10–99':<24}│")
    elif all_triple:
        print(f"  │  Mode         : {'All triple digit':<24}│")
        print(f"  │  Value range  : {'100–999':<24}│")
    elif double_triple:
        print(f"  │  Mode         : {'Alternating (double + triple)':<24}│")
        print(f"  │  Double range : {'10–99':<24}│")
        print(f"  │  Triple range : {'100–999':<24}│")
    else:
        print(f"  │  Mode         : {'All single digit':<24}│")
        print(f"  │  Value range  : {f'{val_min}–{val_max}':<24}│")
    print(f"  │  Output file  : {output:<24}│")
    print("  └─────────────────────────────────────────┘")
    print()

    confirm = ask_yes_no("Generate file with these settings?", default=True)
    if not confirm:
        print("\n  Cancelled. No file was created.\n")
        return

    # ── Generate ─────────────────────────────────────────────────────────────
    print("\n  Generating...")

    wb = Workbook()
    wb.remove(wb.active)

    blocks_per_sheet = (blocks + sheets - 1) // sheets

    block_index = 0
    for sheet_num in range(sheets):
        ws = wb.create_sheet(title=f"Sheet{sheet_num + 1}")
        blocks_this_sheet = min(blocks_per_sheet, blocks - block_index)

        for b in range(blocks_this_sheet):
            block_width = 1 + NUM_COLS + BLOCK_GAP
            start_col   = 1 + b * block_width
            write_block(ws, block_index, start_col, val_min, val_max, double, all_double, all_triple, double_triple)
            block_index += 1

        ws.freeze_panes = "A2"

    wb.save(output)

    print(f"\n  ✅ Done! File saved as: {output}")
    print(f"     {sheets} sheet(s), {blocks} block(s) total, {blocks_per_sheet} per sheet\n")
    subprocess.run(["open", output])


if __name__ == "__main__":
    main()