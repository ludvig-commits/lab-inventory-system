# Lab Inventory System

A terminal (menu-driven) program for managing a laboratory inventory kept in an Excel workbook: secondary antibodies (stored at -20 C) and mouse cages. It supports searching, stock tracking and editing, and automatically makes a backup of the workbook before every save.

## Features

- List and search secondary antibodies
- Register antibody use (subtract volume), restocks, and manual volume entry
- Low-stock report (5 uL or less)
- List mouse cages; search by cage number or genetic background
- Edit an existing mouse cage or add a new one (with input validation and duplicate-cage detection)
- Search across all worksheets at once
- Automatic timestamped backup in `backups/` before any change is saved

## Modules

| File | Purpose |
|------|---------|
| `main.py` | Terminal menu and user interaction; program entry point |
| `data_access.py` | Reading and writing the Excel workbook, cleaning values, creating backups |
| `antibodies.py` | Antibody display, search, volume parsing and stock/low-stock logic |
| `mice.py` | Mouse-cage display, search, validation and numbering logic |
| `search.py` | Search across all worksheets |
| `tests/test_inventory.py` | Automated test suite (`unittest`) |

## Installation

Requires Python 3.10+ (developed with 3.13).

```
git clone https://github.com/ludvig-commits/lab-inventory-system.git
cd lab-inventory-system
pip install -r requirements.txt
```

## Data file

The inventory workbook is **not included in this repository**. Place your own `lab_inventory.xlsx` in the project folder (next to `main.py`).

## Running

```
python main.py
```

Optionally use a different workbook:

```
python main.py path\to\other_workbook.xlsx
```

## Tests

```
python -m unittest discover tests
```

The tests work on a fresh copy of the workbook (created in `tests/test_data/`) and verify the original file is never modified. Note that the current test suite expects `lab_inventory.xlsx` to be present in the project folder.
