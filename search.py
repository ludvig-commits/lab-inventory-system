"""search.py - search through all worksheets of the workbook.

Works on the list of rows from data_access.load_all_rows.
This module does not import any other project module.
"""

MAX_SHOWN = 50      # show at most this many matches, to keep the screen readable


def search_rows(rows, text):
    """Return the rows where the text is found in any cell
    (upper/lower case is ignored)."""
    text = text.strip().lower()
    results = []
    for row in rows:
        for name, value in row["cells"]:
            if text in value.lower():
                results.append(row)
                break   # found once, do not add the same row twice
    return results


def count_by_sheet(results):
    """Return a dictionary: worksheet name -> number of matching rows."""
    counts = {}
    for row in results:
        counts[row["sheet"]] = counts.get(row["sheet"], 0) + 1
    return counts


def show_results(results):
    """Print the matches with their worksheet, Excel row and information."""
    if len(results) == 0:
        print("No matches found in any worksheet.")
        return

    counts = count_by_sheet(results)
    print(f"Found {len(results)} matching row(s) in {len(counts)} worksheet(s):")
    for sheet_name, number in counts.items():
        print(f"  {sheet_name}: {number}")

    for row in results[:MAX_SHOWN]:
        print(f"\n[{row['sheet']}]  row {row['row']}")
        for name, value in row["cells"]:
            if len(value) > 70:
                value = value[:67] + "..."
            print(f"    {name}: {value}")

    if len(results) > MAX_SHOWN:
        print(f"\nShowing the first {MAX_SHOWN} of {len(results)} matches. "
              f"Type a more specific search to see the rest.")
