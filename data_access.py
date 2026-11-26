"""data_access.py - reading and saving the Excel workbook.

This is the only module that uses openpyxl. It returns plain Python
lists of dictionaries, so the other modules never need to know about Excel.
A missing value in Excel is stored as None. For mice and volumes it is
NEVER turned into 0. (Only the antibody stock in tubes is treated as 0 when
empty, as agreed by the group.)

Every save: 1) makes a backup copy, 2) opens the workbook with formulas
(data_only=False) so formulas and formatting are kept, 3) writes only the
cells that belong to the changed record, 4) saves to a temporary file and
then replaces the real file, so a failed save cannot damage the workbook.
"""

import copy
import datetime
import os
import re
import shutil

import openpyxl
from openpyxl.utils import get_column_letter

# --- Layout of the workbook (change here if the sheets change) ---
ANTIBODY_SHEET = "2nd Antibody@-20 degree"
ANTIBODY_FIRST_ROW = 3      # row 1 = title, row 2 = column headers
ANTIBODY_HEADER_ROW = 2
ANTIBODY_STOCK_COL = 1      # A  (number of tubes)
ANTIBODY_INITIAL_COL = 13   # M  (new column: initial volume in uL)
ANTIBODY_REMAINING_COL = 14 # N  (new column: remaining volume in uL)
INITIAL_HEADER = "Initial volume (µL)"
REMAINING_HEADER = "Remaining volume (µL)"

MOUSE_SHEET = "Mouse database"
MOUSE_FIRST_ROW = 11        # row 10 = column headers
MOUSE_COLUMNS = {           # record field -> column number in the sheet
    "number": 1,        # A
    "cage": 2,          # B
    "background": 3,    # C
    "males": 4,         # D
    "females": 5,       # E
    "birth_date": 6,    # F
    "label": 7,         # G
    "owner": 8,         # H
    "experiment": 9,    # I
    "comment": 10,      # J
}

BACKUP_FOLDER = "backups"   # created next to the workbook (relative path)

# Header row of each worksheet, used to show column names in the search.
# None = the sheet has no simple header row (column letters are shown).
HEADER_ROWS = {
    "2nd Antibody@-20 degree": 2,
    "Antibody@4 degree": 2,
    "Antibody@-20 degree": 1,
    "Plasmids": 1,
    "-80C": None,
    "Cell Lines": 1,
    "Mouse database": 10,
    "Chemicals and Reagents": 1,
    "Paraffin blocks": 1,
    "Primer database": 2,
    "Formalin fixed samples": None,
    "Trangenic animal": 1,
    "Pigs (drugs+medical appliances)": 1,
    "10L N2 tank": 1,
}
# Some first columns have no header in Excel; give them a name here.
EXTRA_HEADERS = {
    ("Antibody@-20 degree", 1): "Products",
    ("Cell Lines", 1): "Cell line",
    ("Chemicals and Reagents", 1): "Code",
    ("Chemicals and Reagents", 2): "Product",
    ("Paraffin blocks", 1): "Sample",
    ("Pigs (drugs+medical appliances)", 1): "Category",
    ("Pigs (drugs+medical appliances)", 2): "Item",
    ("10L N2 tank", 1): "Canister",
    ("Formalin fixed samples", 1): "Owner",
    ("Formalin fixed samples", 2): "Content",
}


# ---------------------------------------------------------------- cleaning

def clean_value(value):
    """Return None for empty cells and '-' placeholders.
    Whole numbers like 2.0 become 2. Text is stripped of extra spaces."""
    if value is None:
        return None
    if isinstance(value, float) and value == int(value):
        return int(value)
    if isinstance(value, str):
        value = value.strip()
        if value == "" or value == "-":
            return None
    return value


def clean_count(value):
    """Like clean_value, but for numbers of animals or vials.
    Some cells contain numbers stored as text (e.g. '1'), so convert them.
    Anything that is missing or not a number stays None (unknown)."""
    value = clean_value(value)
    if value is None or isinstance(value, int):
        return value
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return None


def clean_volume(value):
    """Volume in uL as a number (decimals allowed).
    Missing, non-numeric or negative cells stay None (unknown)."""
    value = clean_value(value)
    if value is None or isinstance(value, bool):
        return None
    try:
        number = float(value)
    except (ValueError, TypeError):
        return None
    if number != number or number < 0 or number == float("inf"):
        return None
    return number


def display_text(value):
    """Text of a cell for the all-worksheets search; '' if the cell is empty."""
    if value is None:
        return ""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return " ".join(str(value).split())


# ----------------------------------------------------------------- reading

def open_workbook(path):
    """Open the workbook for reading only (values, not formulas)."""
    return openpyxl.load_workbook(path, data_only=True)


def load_antibodies(path):
    """Read secondary antibodies. Columns: A stock (tubes), B product,
    C isotype, D company, E catalogue number, F batch, G price, J note,
    M initial volume, N remaining volume (uL).
    An empty stock cell is 0 tubes. An empty volume stays None (unknown).
    'row' is the Excel row number, used when saving."""
    workbook = open_workbook(path)
    sheet = workbook[ANTIBODY_SHEET]
    antibodies = []

    first = ANTIBODY_FIRST_ROW
    for row_number, row in enumerate(sheet.iter_rows(min_row=first, values_only=True), start=first):
        product = clean_value(row[1])
        if product is None:
            continue    # blank or stray row: skip it and keep reading

        stock = clean_count(row[0])
        if stock is None:
            stock = 0   # agreed by the group: empty stock means 0 tubes

        antibody = {
            "row": row_number,
            "stock": stock,
            "product": product,
            "isotype": clean_value(row[2]),
            "company": clean_value(row[3]),
            "catalog": clean_value(row[4]),
            "batch": clean_value(row[5]),
            "price": clean_value(row[6]),      # kept as text, e.g. '5700(4000)'
            "note": clean_value(row[9]),
            "initial_volume": clean_volume(row[ANTIBODY_INITIAL_COL - 1]),    # None = unknown
            "remaining_volume": clean_volume(row[ANTIBODY_REMAINING_COL - 1]),  # None = unknown
        }
        antibodies.append(antibody)

    workbook.close()
    return antibodies


def load_mice(path):
    """Read mouse cages. Columns: A number, B cage, C genetic background,
    D males, E females, F birth date, G label, H owner, I experiment, J comment.
    Rows without a cage number (blank rows, totals at the bottom) are skipped.
    'row' is the Excel row number, used when saving."""
    workbook = open_workbook(path)
    sheet = workbook[MOUSE_SHEET]
    mice = []

    first = MOUSE_FIRST_ROW
    for row_number, row in enumerate(sheet.iter_rows(min_row=first, values_only=True), start=first):
        cage = clean_value(row[1])
        if cage is None:
            continue

        cage_record = {
            "row": row_number,
            "number": clean_count(row[0]),
            "cage": str(cage),
            "background": clean_value(row[2]),
            "males": clean_count(row[3]),      # None = unknown
            "females": clean_count(row[4]),    # None = unknown
            "birth_date": clean_value(row[5]), # kept as text, e.g. '2026.03.11'
            "label": clean_value(row[6]),
            "owner": clean_value(row[7]),
            "experiment": clean_value(row[8]),
            "comment": clean_value(row[9]),
        }
        mice.append(cage_record)

    workbook.close()
    return mice


def load_all_rows(path):
    """Read every non-empty row of all worksheets (for the search).
    Returns a list of dictionaries:
      {"sheet": name, "row": Excel row number, "cells": [(column name, text), ...]}
    Header rows are not included."""
    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    results = []

    for sheet in workbook.worksheets:
        header_row = HEADER_ROWS.get(sheet.title)
        headers = {}    # column number -> column name

        for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
            if row_number == header_row:
                for col, value in enumerate(row, start=1):
                    if display_text(value) != "":
                        headers[col] = display_text(value)
                continue

            cells = []
            for col, value in enumerate(row, start=1):
                text = display_text(value)
                if text == "":
                    continue
                name = EXTRA_HEADERS.get((sheet.title, col))
                if name is None:
                    name = headers.get(col)
                if name is None:
                    name = "Column " + get_column_letter(col)
                cells.append((name, text))
            if len(cells) > 0:
                results.append({"sheet": sheet.title, "row": row_number, "cells": cells})

    workbook.close()
    return results


# ------------------------------------------------------------------ saving

def make_backup(path):
    """Copy the workbook to the 'backups' folder next to it before changing it.
    Returns the path of the backup file."""
    folder = os.path.join(os.path.dirname(os.path.abspath(path)), BACKUP_FOLDER)
    os.makedirs(folder, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    name = os.path.splitext(os.path.basename(path))[0] + "_backup_" + stamp + ".xlsx"
    backup_path = os.path.join(folder, name)
    shutil.copy2(path, backup_path)
    return backup_path


def open_for_saving(path):
    """Make a backup, then open the workbook WITH formulas (data_only=False)."""
    make_backup(path)
    return openpyxl.load_workbook(path)


def finish_saving(workbook, path):
    """Save to a temporary file first, then replace the real file."""
    temp_path = path + ".tmp"
    try:
        workbook.save(temp_path)
        os.replace(temp_path, path)
    finally:
        workbook.close()
        if os.path.exists(temp_path):
            os.remove(temp_path)


def check_row(sheet, row, column, expected_text):
    """Safety check: make sure the row still holds the record we loaded.
    Protects against writing into the wrong row if the file was edited."""
    found = clean_value(sheet.cell(row=row, column=column).value)
    if found is None or str(found) != str(expected_text):
        raise ValueError(f"Row {row} of '{sheet.title}' has changed since the "
                         f"program loaded the file. Please restart the program.")


def copy_header_style(sheet, source_cell, target_cell):
    """Give a new header cell the same look as an existing one."""
    target_cell._style = copy.copy(source_cell._style)
    target_cell.number_format = "General"


def ensure_volume_headers(sheet):
    """Create the two new volume column headers if they do not exist yet."""
    source = sheet.cell(row=ANTIBODY_HEADER_ROW, column=ANTIBODY_INITIAL_COL - 2)  # column K
    for col, title in ((ANTIBODY_INITIAL_COL, INITIAL_HEADER),
                       (ANTIBODY_REMAINING_COL, REMAINING_HEADER)):
        cell = sheet.cell(row=ANTIBODY_HEADER_ROW, column=col)
        if cell.value is None:
            cell.value = title
            copy_header_style(sheet, source, cell)
        letter = get_column_letter(col)
        if sheet.column_dimensions[letter].width is None or sheet.column_dimensions[letter].width < 22:
            sheet.column_dimensions[letter].width = 22


def save_antibody(path, antibody):
    """Write stock, initial volume and remaining volume of ONE antibody record."""
    workbook = open_for_saving(path)
    try:
        sheet = workbook[ANTIBODY_SHEET]
        check_row(sheet, antibody["row"], 2, antibody["product"])
        ensure_volume_headers(sheet)
        row = antibody["row"]
        sheet.cell(row=row, column=ANTIBODY_STOCK_COL).value = antibody["stock"]
        sheet.cell(row=row, column=ANTIBODY_INITIAL_COL).value = antibody["initial_volume"]
        sheet.cell(row=row, column=ANTIBODY_REMAINING_COL).value = antibody["remaining_volume"]
    except Exception:
        workbook.close()
        raise
    finish_saving(workbook, path)


def save_mouse(path, record, old_cage):
    """Write all columns of ONE existing mouse record.
    old_cage is the cage number the record had when it was loaded (safety check)."""
    workbook = open_for_saving(path)
    try:
        sheet = workbook[MOUSE_SHEET]
        check_row(sheet, record["row"], MOUSE_COLUMNS["cage"], old_cage)
        for field, col in MOUSE_COLUMNS.items():
            sheet.cell(row=record["row"], column=col).value = record[field]
    except Exception:
        workbook.close()
        raise
    finish_saving(workbook, path)


def add_mouse(path, record):
    """Add a new cage directly after the last cage of the table.
    The rows below (totals, notes) are moved down by one row and the totals
    that sum the cage table are extended to include the new row.
    Returns the Excel row number of the new cage."""
    workbook = open_for_saving(path)
    try:
        sheet = workbook[MOUSE_SHEET]

        last = MOUSE_FIRST_ROW - 1      # last row that holds a cage
        for r in range(MOUSE_FIRST_ROW, sheet.max_row + 1):
            if clean_value(sheet.cell(row=r, column=MOUSE_COLUMNS["cage"]).value) is not None:
                last = r
        new_row = last + 1
        max_row = sheet.max_row
        max_col = sheet.max_column

        # 1. move everything below the table down by one row (formulas are
        #    moved as text, not changed)
        if max_row >= new_row:
            sheet.move_range(f"A{new_row}:{get_column_letter(max_col)}{max_row}", rows=1)
            # 2. extend sums like =SUM(D11:D45) that end on the old last row
            pattern = re.compile(rf"([A-Z]+{MOUSE_FIRST_ROW}:[A-Z]+){last}(?![0-9])")
            for row_cells in sheet.iter_rows(min_row=new_row + 1, max_row=max_row + 1):
                for cell in row_cells:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        cell.value = pattern.sub(lambda m: m.group(1) + str(new_row), cell.value)

        # 3. the new row looks like the last cage row
        for col in range(1, max_col + 1):
            sheet.cell(row=new_row, column=col)._style = copy.copy(sheet.cell(row=last, column=col)._style)
        if last in sheet.row_dimensions and sheet.row_dimensions[last].height is not None:
            sheet.row_dimensions[new_row].height = sheet.row_dimensions[last].height

        # 4. write the values
        for field, col in MOUSE_COLUMNS.items():
            sheet.cell(row=new_row, column=col).value = record[field]
    except Exception:
        workbook.close()
        raise
    finish_saving(workbook, path)
    return new_row
