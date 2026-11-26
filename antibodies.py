"""antibodies.py - antibody inventory functions.

Works on a list of antibody dictionaries (see data_access.load_antibodies).
This module does not import any other project module.

Stock = number of tubes (empty in Excel = 0).
Volumes are in microliters (uL), tracked per product record (not per tube).
A volume that was never recorded is None = unknown (never assumed).
"""

LOW_VOLUME_LIMIT = 5    # warn when the known remaining volume is 5 uL or less


def text_or_dash(value):
    """Text for a normal field: a dash if the value is missing."""
    if value is None:
        return "-"
    return str(value)


def stock_text(stock):
    """Number of tubes as text (empty stock is already 0)."""
    return str(stock)


def volume_text(volume):
    """Volume in uL as text. Unknown stays 'unknown'; decimals are kept."""
    if volume is None:
        return "unknown"
    text = format(round(volume, 4), "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def is_low_stock(antibody):
    """True if the remaining volume is KNOWN and is 5 uL or less.
    An unknown volume never gives a warning."""
    volume = antibody["remaining_volume"]
    return volume is not None and volume <= LOW_VOLUME_LIMIT


def low_stock_message(antibody):
    """Warning text for one antibody ('' if its volume is not low)."""
    if not is_low_stock(antibody):
        return ""
    return (f"WARNING: low stock! {antibody['product']} has only "
            f"{volume_text(antibody['remaining_volume'])} uL left "
            f"(limit: {LOW_VOLUME_LIMIT} uL).")


def show_antibodies(antibodies, numbered=False):
    """Print a list of antibodies as a table.
    numbered=True adds a number in front of each line, so the user can pick one."""
    if len(antibodies) == 0:
        print("No antibodies to show.")
        return

    prefix = "No.  " if numbered else ""
    print(f"{prefix}{'Tubes':<7}{'Product':<38}{'Isotype':<9}{'Company':<24}"
          f"{'Cat.#':<14}{'Initial uL':<12}{'Remaining uL':<14}")
    print("-" * (111 + len(prefix)))
    for position, a in enumerate(antibodies, start=1):
        product = a["product"]
        if len(product) > 36:
            product = product[:33] + "..."
        start = f"{position:<5}" if numbered else ""
        line = (f"{start}{stock_text(a['stock']):<7}"
                f"{product:<38}"
                f"{text_or_dash(a['isotype']):<9}"
                f"{text_or_dash(a['company']):<24}"
                f"{text_or_dash(a['catalog']):<14}"
                f"{volume_text(a['initial_volume']):<12}"
                f"{volume_text(a['remaining_volume']):<14}")
        if is_low_stock(a):
            line += "LOW STOCK"
        print(line)
    print(f"\n{len(antibodies)} antibody record(s).")


def search_antibodies(antibodies, text):
    """Return the antibodies where the text is found in the product,
    isotype, company or catalogue number (upper/lower case is ignored)."""
    text = text.strip().lower()
    results = []
    for a in antibodies:
        for field in ("product", "isotype", "company", "catalog"):
            value = a[field]
            if value is not None and text in str(value).lower():
                results.append(a)
                break   # found once, do not add the same antibody twice
    return results


def low_stock_antibodies(antibodies):
    """Return the antibodies with a known remaining volume of 5 uL or less."""
    results = []
    for a in antibodies:
        if is_low_stock(a):
            results.append(a)
    return results


def count_unknown_volume(antibodies):
    """How many antibodies have no remaining volume recorded (unknown)."""
    count = 0
    for a in antibodies:
        if a["remaining_volume"] is None:
            count += 1
    return count


# ------------------------------------------------------------ validation

def parse_volume(text):
    """Convert text to a volume in uL (decimals allowed, e.g. '2.5').
    Returns (number, None) if valid or (None, error message) if not.
    Negative numbers, text, 'nan' and 'inf' are rejected. Zero is allowed here."""
    text = text.strip()
    try:
        if "_" in text:
            raise ValueError    # Python would accept '1_0' as 10
        number = float(text)
    except ValueError:
        return None, f"'{text}' is not a number. Use digits, e.g. 2 or 2.5."
    if number != number or number in (float("inf"), float("-inf")):
        return None, f"'{text}' is not a valid number."
    if number < 0:
        return None, "The volume cannot be negative."
    return round(number, 4), None


def parse_amount_used(text):
    """Like parse_volume, but the amount used must be greater than 0."""
    number, error = parse_volume(text)
    if error is None and number == 0:
        return None, "The amount used must be greater than 0."
    return number, error


def parse_tubes(text):
    """Convert text to a whole number of tubes (0 or more).
    Returns (number, None) if valid or (None, error message) if not."""
    text = text.strip()
    if not text.isdecimal():
        return None, f"'{text}' is not a whole number of tubes (0, 1, 2, ...)."
    return int(text), None


# ------------------------------------------------------------- changes
# These functions change the dictionary only if everything is valid.
# They return an error message, or None when the change was made.

def use_volume(antibody, amount):
    """Register an amount (uL) used in an experiment: subtract it from the
    remaining volume."""
    if amount is None or amount <= 0:
        return "The amount used must be greater than 0."
    remaining = antibody["remaining_volume"]
    if remaining is None:
        return ("The remaining volume of this antibody is unknown, so the "
                "amount used cannot be subtracted. Set the volume first.")
    if amount > remaining:
        return (f"Not enough antibody: only {volume_text(remaining)} uL "
                f"available, but {volume_text(amount)} uL requested.")
    antibody["remaining_volume"] = round(remaining - amount, 4)
    return None


def restock(antibody, tubes, volume):
    """Register restocking: add tubes to the stock and, if given (not None),
    add the volume (uL) to the remaining volume (and to the initial volume,
    if that is known).
    If the remaining volume was unknown it stays unknown (we do not guess
    how much was already there)."""
    if tubes is None or tubes < 0 or (volume is not None and volume < 0):
        return "Restock quantities cannot be negative."
    if tubes == 0 and (volume is None or volume == 0):
        return "Nothing to restock: enter at least one tube or a volume."
    antibody["stock"] = antibody["stock"] + tubes
    if volume is not None and antibody["remaining_volume"] is not None:
        antibody["remaining_volume"] = round(antibody["remaining_volume"] + volume, 4)
        if antibody["initial_volume"] is not None:
            antibody["initial_volume"] = round(antibody["initial_volume"] + volume, 4)
    return None


def set_volumes(antibody, initial, remaining):
    """Enter volumes manually. None = leave unchanged.
    The remaining volume may not be larger than the initial volume."""
    new_initial = antibody["initial_volume"] if initial is None else initial
    new_remaining = antibody["remaining_volume"] if remaining is None else remaining
    if new_initial is not None and new_remaining is not None and new_remaining > new_initial:
        return (f"The remaining volume ({volume_text(new_remaining)} uL) cannot be "
                f"larger than the initial volume ({volume_text(new_initial)} uL).")
    antibody["initial_volume"] = new_initial
    antibody["remaining_volume"] = new_remaining
    return None
