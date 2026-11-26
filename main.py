"""main.py - terminal menu for the lab inventory.

Run with:  python main.py
Needs:     pip install openpyxl

Optional: python main.py path\\to\\other_workbook.xlsx   (used for testing)
Changes are saved to the workbook automatically (a backup is made first).
"""

import os
import sys

import data_access
import antibodies
import mice
import search

# The workbook is expected in the same folder as this file (relative path).
EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "lab_inventory.xlsx")


def print_menu():
    print("\n=== Lab Inventory ===")
    print("1. Show all secondary antibodies (-20 C)")
    print("2. Search antibodies")
    print("3. Show all mouse cages")
    print("4. Search mice by cage number")
    print("5. Search mice by genetic background")
    print("6. Search ALL worksheets")
    print("7. Register antibody use (subtract volume)")
    print("8. Register antibody restock")
    print("9. Enter antibody volumes manually")
    print("10. Show low-stock antibodies (5 uL or less)")
    print("11. Edit a mouse cage")
    print("12. Add a new mouse cage")
    print("0. Quit")


# ------------------------------------------------------------ small helpers

def ask_text(question):
    """Ask for a search text; return None if the user types nothing."""
    text = input(question).strip()
    if text == "":
        print("Nothing entered.")
        return None
    return text


def ask_yes_no(question):
    """Ask a yes/no question; anything except y/yes means no."""
    answer = input(question + " (y/n): ").strip().lower()
    return answer == "y" or answer == "yes"


def choose_from(records, show_function, what):
    """Show a numbered list and let the user pick one record.
    Returns the record, or None if the user cancels (empty input)."""
    show_function(records)
    while True:
        text = input(f"Enter the number of the {what} (empty = cancel): ").strip()
        if text == "":
            print("Cancelled.")
            return None
        if text.isdecimal() and 1 <= int(text) <= len(records):
            return records[int(text) - 1]
        print(f"Please enter a number from 1 to {len(records)}.")


def ask_number(question, parser, allow_empty=False):
    """Ask until the input is valid. parser returns (value, error).
    Empty input: cancel (None), or 'no value' if allow_empty is True.
    Returns (ok, value)."""
    while True:
        text = input(question).strip()
        if text == "":
            return allow_empty, None
        value, error = parser(text)
        if error is None:
            return True, value
        print("Invalid input:", error)


def save_or_undo(save_function, record, old_values):
    """Run the save. If saving fails, restore the old values in memory so
    that the screen never shows data that is not in the file."""
    try:
        save_function()
    except Exception as error:
        record.clear()
        record.update(old_values)
        print("Could not save to Excel, the change was NOT made:", error)
        print("(If the file is open in Excel, close it and try again.)")
        return False
    return True


# ------------------------------------------------------ antibody functions

def pick_antibody(antibody_list):
    """Search for an antibody (empty = list all) and let the user pick one."""
    text = input("Search the antibody (product, isotype, company, cat.#; "
                 "empty = show all): ").strip()
    if text == "":
        matches = antibody_list
    else:
        matches = antibodies.search_antibodies(antibody_list, text)
    if len(matches) == 0:
        print("No antibodies found.")
        return None
    return choose_from(matches, lambda r: antibodies.show_antibodies(r, numbered=True),
                       "antibody")


def print_antibody_state(a):
    print(f"{a['product']} ({antibodies.text_or_dash(a['company'])}, "
          f"cat.# {antibodies.text_or_dash(a['catalog'])})")
    print(f"  Tubes: {a['stock']}   "
          f"Initial: {antibodies.volume_text(a['initial_volume'])} uL   "
          f"Remaining: {antibodies.volume_text(a['remaining_volume'])} uL")


def use_antibody(path, antibody_list):
    a = pick_antibody(antibody_list)
    if a is None:
        return
    print_antibody_state(a)
    if a["remaining_volume"] is None:
        print("The remaining volume is unknown, so use cannot be registered.")
        print("Use menu option 9 to enter the volume first.")
        return
    ok, amount = ask_number("Amount used in uL (decimals allowed, empty = cancel): ",
                            antibodies.parse_amount_used)
    if not ok:
        print("Cancelled.")
        return

    old = dict(a)
    error = antibodies.use_volume(a, amount)
    if error is not None:
        print("Not registered:", error)
        return
    if save_or_undo(lambda: data_access.save_antibody(path, a), a, old):
        print(f"Saved. Used {antibodies.volume_text(amount)} uL, "
              f"{antibodies.volume_text(a['remaining_volume'])} uL remaining.")
        if antibodies.is_low_stock(a):
            print(antibodies.low_stock_message(a))


def restock_antibody(path, antibody_list):
    a = pick_antibody(antibody_list)
    if a is None:
        return
    print_antibody_state(a)
    ok, tubes = ask_number("Number of tubes added (whole number, empty = 0): ",
                           antibodies.parse_tubes, allow_empty=True)
    if not ok:
        return
    if tubes is None:
        tubes = 0
    ok, volume = ask_number("Volume added in uL (decimals allowed, empty = not recorded): ",
                            antibodies.parse_volume, allow_empty=True)
    if not ok:
        return

    old = dict(a)
    error = antibodies.restock(a, tubes, volume)
    if error is not None:
        print("Not registered:", error)
        return
    if save_or_undo(lambda: data_access.save_antibody(path, a), a, old):
        print(f"Saved. Tubes: {a['stock']}.")
        if volume is not None and old["remaining_volume"] is None:
            print("The old remaining volume was unknown, so the volume stays unknown. "
                  "Use menu option 9 to enter it.")
        elif volume is None:
            print("No volume was entered, so the volumes were not changed.")
        else:
            print(f"Remaining volume: {antibodies.volume_text(a['remaining_volume'])} uL.")


def enter_volumes(path, antibody_list):
    a = pick_antibody(antibody_list)
    if a is None:
        return
    print_antibody_state(a)
    print("Press Enter to keep a value unchanged.")
    ok, initial = ask_number("New initial volume in uL: ", antibodies.parse_volume,
                             allow_empty=True)
    if not ok:
        return
    ok, remaining = ask_number("New remaining volume in uL: ", antibodies.parse_volume,
                               allow_empty=True)
    if not ok:
        return
    if initial is None and remaining is None:
        print("Nothing changed.")
        return

    old = dict(a)
    error = antibodies.set_volumes(a, initial, remaining)
    if error is not None:
        print("Not saved:", error)
        return
    if save_or_undo(lambda: data_access.save_antibody(path, a), a, old):
        print("Saved.")
        print_antibody_state(a)
        if antibodies.is_low_stock(a):
            print(antibodies.low_stock_message(a))


# --------------------------------------------------------- mouse functions

def field_text(field, value):
    """How a value is shown in a prompt."""
    if field in ("males", "females"):
        return mice.count_text(value)
    return mice.text_or_dash(value)


def confirm_duplicate(mouse_list, cage, record=None):
    """If the cage number is already used, warn and ask. True = go on."""
    others = mice.other_cages_with_number(mouse_list, cage, record)
    if len(others) == 0:
        return True
    rows = ", ".join(str(m["row"]) for m in others)
    print(f"WARNING: cage number {cage} already exists (Excel row {rows}).")
    return ask_yes_no("Use this cage number anyway?")


def ask_mouse_field(field, label, current=None, adding=False):
    """Ask for one field until valid. Returns (ok, value).
    Editing: Enter keeps the current value, '-' clears it.
    Adding: Enter / '-' means no value (cage number is required)."""
    while True:
        if adding:
            question = f"{label}: "
        else:
            question = f"{label} [{field_text(field, current)}] (Enter = keep, - = clear): "
        text = input(question)
        if text.strip() == "" and not adding:
            return True, current        # keep
        if text.strip() == "" and adding and field == "number":
            return True, None           # the caller fills in the default number
        if text.strip() == "" and adding and field == "cage":
            return False, None          # the cage number is required: cancel
        value, error = mice.validate_field(field, text)
        if error is None:
            return True, value
        print("Invalid input:", error)


def edit_mouse(path, mouse_list):
    text = ask_text("Cage number to edit (or part of it): ")
    if text is None:
        return
    matches = mice.search_by_cage(mouse_list, text)
    if len(matches) == 0:
        print("No cages found.")
        return
    record = choose_from(matches, lambda r: mice.show_mice_choices(r, mouse_list), "cage")
    if record is None:
        return

    print("\nEditing the cage in Excel row", record["row"])
    old = dict(record)
    changes = {}
    for field, label in mice.EDIT_FIELDS:
        ok, value = ask_mouse_field(field, label, record[field])
        if value != record[field]:
            if field == "cage" and not confirm_duplicate(mouse_list, value, record):
                print("Cage number not changed.")
                continue
            changes[field] = value

    if len(changes) == 0:
        print("Nothing changed.")
        return
    print("\nChanges:")
    for field, label in mice.EDIT_FIELDS:
        if field in changes:
            print(f"  {label}: {field_text(field, record[field])} -> "
                  f"{field_text(field, changes[field])}")
    if not ask_yes_no("Save these changes?"):
        print("Cancelled. Nothing was saved.")
        return

    record.update(changes)
    if save_or_undo(lambda: data_access.save_mouse(path, record, old["cage"]), record, old):
        print("Saved.")


def add_mouse(path, mouse_list):
    print("Add a new cage. Empty/'-' = no value (the number of mice stays unknown).")
    record = {}
    for field, label in mice.EDIT_FIELDS:
        if field == "number":
            label = f"Number (Enter = {mice.next_number(mouse_list)})"
        ok, value = ask_mouse_field(field, label, adding=True)
        if not ok:
            print("The cage number is required. Cancelled. Nothing was saved.")
            return
        if field == "number" and value is None:
            value = mice.next_number(mouse_list)
        if field == "cage" and not confirm_duplicate(mouse_list, value):
            print("Cancelled. Nothing was saved.")
            return
        record[field] = value

    print("\nNew cage:")
    for field, label in mice.EDIT_FIELDS:
        print(f"  {label}: {field_text(field, record[field])}")
    if not ask_yes_no("Save this new cage?"):
        print("Cancelled. Nothing was saved.")
        return

    try:
        record["row"] = data_access.add_mouse(path, record)
    except Exception as error:
        print("Could not save to Excel, the cage was NOT added:", error)
        print("(If the file is open in Excel, close it and try again.)")
        return
    mouse_list.append(record)
    print(f"Saved in Excel row {record['row']}.")


# -------------------------------------------------------------------- main

def run_menu(path, antibody_list, mouse_list):
    running = True
    while running:
        print_menu()
        choice = input("Choose an option: ").strip()

        if choice == "1":
            antibodies.show_antibodies(antibody_list)
            unknown = antibodies.count_unknown_volume(antibody_list)
            print(f"Note: the remaining volume is unknown (not zero) for {unknown} antibody record(s).")
        elif choice == "2":
            text = ask_text("Search (product, isotype, company or cat.#): ")
            if text is not None:
                antibodies.show_antibodies(antibodies.search_antibodies(antibody_list, text))
        elif choice == "3":
            mice.show_mice(mouse_list)
            males, females, unknown = mice.count_known_mice(mouse_list)
            print(f"Known mice: {males} male, {females} female. "
                  f"{unknown} cage(s) have unknown counts (not included).")
        elif choice == "4":
            text = ask_text("Cage number (or part of it): ")
            if text is not None:
                mice.show_mice(mice.search_by_cage(mouse_list, text), mouse_list)
        elif choice == "5":
            text = ask_text("Genetic background (e.g. C57BL/6 or ICR): ")
            if text is not None:
                mice.show_mice(mice.search_by_background(mouse_list, text), mouse_list)
        elif choice == "6":
            text = ask_text("Search all worksheets for: ")
            if text is not None:
                try:
                    rows = data_access.load_all_rows(path)
                    search.show_results(search.search_rows(rows, text))
                except Exception as error:
                    print("Could not read the Excel file:", error)
        elif choice == "7":
            use_antibody(path, antibody_list)
        elif choice == "8":
            restock_antibody(path, antibody_list)
        elif choice == "9":
            enter_volumes(path, antibody_list)
        elif choice == "10":
            low = antibodies.low_stock_antibodies(antibody_list)
            if len(low) == 0:
                print("No antibody has a known remaining volume of 5 uL or less.")
            else:
                antibodies.show_antibodies(low)
        elif choice == "11":
            edit_mouse(path, mouse_list)
        elif choice == "12":
            add_mouse(path, mouse_list)
        elif choice == "0":
            print("Goodbye!")
            running = False
        else:
            print("Invalid choice, please enter 0-12.")


def main():
    path = EXCEL_FILE
    if len(sys.argv) > 1:
        path = sys.argv[1]

    if not os.path.exists(path):
        print("Could not find the Excel file:", path)
        return

    try:
        antibody_list = data_access.load_antibodies(path)
        mouse_list = data_access.load_mice(path)
    except Exception as error:
        print("Could not read the Excel file:", error)
        return

    print(f"Loaded {len(antibody_list)} antibodies and {len(mouse_list)} mouse cages.")
    low = antibodies.low_stock_antibodies(antibody_list)
    if len(low) > 0:
        print(f"WARNING: {len(low)} antibody record(s) have 5 uL or less left "
              f"(menu option 10).")

    try:
        run_menu(path, antibody_list, mouse_list)
    except (EOFError, KeyboardInterrupt):
        print("\nInput ended. Goodbye!")


if __name__ == "__main__":
    main()
