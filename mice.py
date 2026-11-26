"""mice.py - mouse inventory functions.

Works on a list of cage dictionaries (see data_access.load_mice).
This module does not import any other project module.
"""

import datetime


def text_or_dash(value):
    """Text for a normal field: a dash if the value is missing."""
    if value is None:
        return "-"
    return str(value)


def count_text(count):
    """Missing number of mice is 'unknown', never 0."""
    if count is None:
        return "unknown"
    return str(count)


def find_duplicate_cages(mice):
    """Return a list of cage numbers that appear more than once."""
    seen = []
    duplicates = []
    for m in mice:
        if m["cage"] in seen and m["cage"] not in duplicates:
            duplicates.append(m["cage"])
        seen.append(m["cage"])
    return duplicates


def show_mice(mice, all_mice=None):
    """Print cages as a table. Duplicate cage numbers get a warning.
    all_mice is the full list, used to detect duplicates when only
    search results are shown (default: the list itself)."""
    if len(mice) == 0:
        print("No mouse cages to show.")
        return
    if all_mice is None:
        all_mice = mice
    duplicates = find_duplicate_cages(all_mice)

    print(f"{'Cage':<22}{'Background':<12}{'Male':<9}{'Female':<9}"
          f"{'Birth date':<12}{'Label':<30}{'Owner':<9}")
    print("-" * 103)
    for m in mice:
        line = (f"{m['cage']:<22}"
                f"{text_or_dash(m['background']):<12}"
                f"{count_text(m['males']):<9}"
                f"{count_text(m['females']):<9}"
                f"{text_or_dash(m['birth_date']):<12}"
                f"{text_or_dash(m['label']):<30}"
                f"{text_or_dash(m['owner']):<9}")
        if m["cage"] in duplicates:
            line += "  WARNING: duplicate cage number"
        print(line)
    print(f"\n{len(mice)} cage(s).")


def short(value, width):
    """Text of a field, cut to a maximum width (for tables)."""
    text = text_or_dash(value)
    if len(text) > width:
        text = text[:width - 3] + "..."
    return text


def show_mice_choices(mice, all_mice):
    """Print cages with a number in front, so the user can choose one.
    The Excel row and more columns are shown, so that cages with the same
    cage number can be told apart."""
    duplicates = find_duplicate_cages(all_mice)
    print(f"{'No.':<5}{'Row':<6}{'Cage':<16}{'Backgr.':<9}{'M':<9}{'F':<9}"
          f"{'Birth date':<12}{'Label':<12}{'Owner':<10}{'Experiment':<14}{'Comment':<22}")
    print("-" * 124)
    for position, m in enumerate(mice, start=1):
        line = (f"{position:<5}{m['row']:<6}{short(m['cage'], 15):<16}"
                f"{short(m['background'], 8):<9}"
                f"{count_text(m['males']):<9}{count_text(m['females']):<9}"
                f"{short(m['birth_date'], 11):<12}{short(m['label'], 11):<12}"
                f"{short(m['owner'], 9):<10}{short(m['experiment'], 13):<14}"
                f"{short(m['comment'], 21):<22}")
        if m["cage"] in duplicates:
            line += " (duplicate cage no.)"
        print(line)
    print(f"\n{len(mice)} cage(s).")


# Fields that can be edited: (field name, label shown to the user)
EDIT_FIELDS = [
    ("number", "Number"),
    ("cage", "Cage number"),
    ("background", "Genetic background"),
    ("males", "Males"),
    ("females", "Females"),
    ("birth_date", "Birth date (YYYY.MM.DD)"),
    ("label", "Label"),
    ("owner", "Owner"),
    ("experiment", "Experiment"),
    ("comment", "Comment"),
]


def next_number(mice):
    """The next free 'number' (one more than the highest number used)."""
    highest = 0
    for m in mice:
        if m["number"] is not None and m["number"] > highest:
            highest = m["number"]
    return highest + 1


def other_cages_with_number(mice, cage, record=None):
    """Return the records that already use this cage number
    (not counting 'record' itself)."""
    found = []
    for m in mice:
        if m is not record and m["cage"] == cage:
            found.append(m)
    return found


def parse_birth_date(text):
    """Accept YYYY.MM.DD (also with - or /) or YYYY.MM. (month only, which
    already exists in the data). Returns (text in YYYY.MM.DD form, None)
    or (None, error message)."""
    cleaned = text.strip().replace("-", ".").replace("/", ".")
    parts = cleaned.rstrip(".").split(".")
    ok = len(parts) in (2, 3) and all(p.isdecimal() for p in parts)
    if ok and len(parts[0]) == 4 and len(parts[1]) in (1, 2):
        try:
            if len(parts) == 3:
                datetime.date(int(parts[0]), int(parts[1]), int(parts[2]))
                return f"{parts[0]}.{int(parts[1]):02d}.{int(parts[2]):02d}", None
            datetime.date(int(parts[0]), int(parts[1]), 1)
            return f"{parts[0]}.{int(parts[1]):02d}.", None
        except ValueError:
            pass    # impossible date such as 2026.02.30
    return None, f"'{text.strip()}' is not a valid date. Use YYYY.MM.DD, e.g. 2026.03.11."


def validate_field(field, text):
    """Check one typed value for a mouse field.
    Returns (value, None) if valid or (None, error message) if not.
    '-' means 'no value' (the field is cleared and stays unknown/empty)."""
    text = text.strip()
    empty = (text == "" or text == "-")

    if text.startswith("="):
        return None, "Text cannot start with '=' (Excel would treat it as a formula)."
    if field == "cage":
        if empty:
            return None, "The cage number cannot be empty."
        return text, None
    if field == "number":
        if not text.isdecimal() or int(text) < 1:
            return None, "The number must be a whole number of 1 or more."
        return int(text), None
    if empty:
        return None, None
    if field in ("males", "females"):
        if text.lower() == "unknown":
            return None, None
        if not text.isdecimal():
            return None, "The number of mice must be a whole number (0, 1, 2, ...); '-' = unknown."
        return int(text), None
    if field == "birth_date":
        return parse_birth_date(text)
    return text, None       # background, label, owner, experiment, comment


def search_by_cage(mice, text):
    """Return cages whose cage number contains the text."""
    text = text.strip().lower()
    results = []
    for m in mice:
        if text in m["cage"].lower():
            results.append(m)
    return results


def search_by_background(mice, text):
    """Return cages whose genetic background contains the text."""
    text = text.strip().lower()
    results = []
    for m in mice:
        if m["background"] is not None and text in m["background"].lower():
            results.append(m)
    return results


def count_known_mice(mice):
    """Return (males, females, cages_with_unknown_counts).
    Unknown counts are NOT added as zero; they are counted separately."""
    males = 0
    females = 0
    unknown_cages = 0
    for m in mice:
        if m["males"] is None and m["females"] is None:
            unknown_cages += 1
            continue
        if m["males"] is not None:
            males += m["males"]
        if m["females"] is not None:
            females += m["females"]
    return males, females, unknown_cages

