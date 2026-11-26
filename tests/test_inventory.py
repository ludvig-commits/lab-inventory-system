"""test_inventory.py - tests for the lab inventory system.

Run from the project folder with:   python -m unittest discover -s tests -v

The tests NEVER touch lab_inventory.xlsx. Each test works on a fresh copy
(tests/test_data/lab_inventory_TEST.xlsx). The antibody volumes below are
SAMPLE VALUES FOR TESTING ONLY - they are not real lab measurements. They are
written into the test copy, and marked in column O of that copy.
"""

import hashlib
import os
import shutil
import subprocess
import sys
import unittest
from unittest import mock

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(HERE)
sys.path.insert(0, PROJECT)

import data_access      # noqa: E402
import antibodies       # noqa: E402
import mice             # noqa: E402
import search           # noqa: E402

REAL_FILE = os.path.join(PROJECT, "lab_inventory.xlsx")
TEST_DIR = os.path.join(HERE, "test_data")
TEST_FILE = os.path.join(TEST_DIR, "lab_inventory_TEST.xlsx")
MAIN = os.path.join(PROJECT, "main.py")

# --- SAMPLE TEST DATA (not real measurements): catalogue number -> (initial, remaining) uL
SAMPLE_VOLUMES = {
    "115-545-003": (100.0, 100.0),   # plenty
    "715-545-151": (50.5, 7.25),     # decimals
    "711-545-152": (20.0, 5.0),      # exactly at the 5 uL limit
    "705-545-003": (10.0, 5.1),      # just above the limit
}
# Records left WITHOUT volumes on purpose:  712-547-003 (stock 'unknown'/empty is not used,
# stock 1) and 115-585-003 (stock cell is EMPTY in Excel -> 0 tubes).
UNKNOWN_VOLUME_CAT = "712-547-003"
EMPTY_STOCK_CAT = "115-585-003"

DUPLICATE_CAGE = "113-00149-46"


def file_hash(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


REAL_HASH_AT_START = file_hash(REAL_FILE)


def tearDownModule():
    assert file_hash(REAL_FILE) == REAL_HASH_AT_START, "lab_inventory.xlsx was modified!"


def find_antibody(items, catalog):
    for a in items:
        if a["catalog"] == catalog:
            return a
    raise AssertionError("antibody not found: " + catalog)


def make_test_copy():
    """Fresh copy of the workbook + clearly marked sample volumes."""
    if os.path.exists(TEST_DIR):
        shutil.rmtree(TEST_DIR)
    os.makedirs(TEST_DIR)
    shutil.copy2(REAL_FILE, TEST_FILE)
    workbook = openpyxl.load_workbook(TEST_FILE)
    sheet = workbook[data_access.ANTIBODY_SHEET]
    sheet.cell(row=2, column=15).value = "TEST COPY"
    for row in sheet.iter_rows(min_row=3, max_col=5):
        catalog = row[4].value
        if catalog in SAMPLE_VOLUMES:
            initial, remaining = SAMPLE_VOLUMES[catalog]
            sheet.cell(row=row[0].row, column=data_access.ANTIBODY_INITIAL_COL).value = initial
            sheet.cell(row=row[0].row, column=data_access.ANTIBODY_REMAINING_COL).value = remaining
            sheet.cell(row=row[0].row, column=15).value = "SAMPLE VOLUME - TEST DATA ONLY"
    data_access.ensure_volume_headers(sheet)
    workbook.save(TEST_FILE)
    workbook.close()


def run_program(lines, path=TEST_FILE):
    """Run the real program (python main.py <test file>) with typed input.
    Each call is a separate process = a program restart."""
    result = subprocess.run([sys.executable, "-X", "utf8", MAIN, path],
                            input="\n".join(lines) + "\n", capture_output=True,
                            text=True, encoding="utf-8", cwd=PROJECT, timeout=180)
    assert result.returncode == 0, result.stderr
    assert "Traceback" not in result.stdout + result.stderr, result.stderr
    return result.stdout


def backups():
    folder = os.path.join(TEST_DIR, data_access.BACKUP_FOLDER)
    if not os.path.isdir(folder):
        return []
    return sorted(os.path.join(folder, n) for n in os.listdir(folder))


def snapshot(path):
    """Everything about every cell that we want to stay the same."""
    workbook = openpyxl.load_workbook(path)        # formulas, not values
    result = {"sheets": workbook.sheetnames}
    for sheet in workbook.worksheets:
        style_cache = {}
        cells = {}
        for (row, col), cell in sheet._cells.items():
            sid = cell.style_id
            if sid not in style_cache:
                style_cache[sid] = (repr(cell.font), repr(cell.fill), repr(cell.border),
                                    repr(cell.alignment), cell.number_format)
            cells[(row, col)] = (cell.value, style_cache[sid])
        result[sheet.title] = {
            "cells": cells,
            "merged": sorted(str(r) for r in sheet.merged_cells.ranges),
            "cf": len(sheet.conditional_formatting),
            "tables": sorted(sheet.tables.keys()),
            "autofilter": sheet.auto_filter.ref,
            "widths": {k: v.width for k, v in sheet.column_dimensions.items()},
        }
    workbook.close()
    return result


class Base(unittest.TestCase):
    def setUp(self):
        make_test_copy()


# ====================================================== 1. all-sheet search

class TestSearchAllSheets(Base):
    # (search text, worksheet where it must be found)
    EXPECTED = [
        ("Goat anti-Mouse Alexa", "2nd Antibody@-20 degree"),
        ("Anti-Aquaporin 2", "Antibody@4 degree"),
        ("Anti-53BP1", "Antibody@-20 degree"),
        ("ampicillin", "Plasmids"),
        ("hypochlorite", "-80C"),
        ("AIIB2 hybridoma", "Cell Lines"),
        ("112-00111-03", "Mouse database"),
        ("Acetic Acid", "Chemicals and Reagents"),
        ("Kiss1 KO mouse testis", "Paraffin blocks"),
        ("rActb", "Primer database"),
        ("Formalin fixed", "Formalin fixed samples"),
        ("Sox Cre", "Trangenic animal"),
        ("Xylazine", "Pigs (drugs+medical appliances)"),
        ("pangolin", "10L N2 tank"),
    ]

    def setUp(self):
        super().setUp()
        self.rows = data_access.load_all_rows(TEST_FILE)

    def test_workbook_has_14_sheets_and_all_are_searchable(self):
        sheets_in_rows = {r["sheet"] for r in self.rows}
        self.assertEqual(len(openpyxl.load_workbook(TEST_FILE).sheetnames), 14)
        self.assertEqual(sheets_in_rows, set(data_access.HEADER_ROWS.keys()))

    def test_each_sheet_finds_its_item(self):
        for text, sheet_name in self.EXPECTED:
            with self.subTest(sheet=sheet_name, text=text):
                hits = search.search_rows(self.rows, text)
                self.assertIn(sheet_name, search.count_by_sheet(hits))

    def test_covers_all_14_sheets(self):
        self.assertEqual(len({s for _, s in self.EXPECTED}), 14)

    def test_result_shows_sheet_row_and_information(self):
        hit = search.search_rows(self.rows, "ampicillin")[0]
        self.assertEqual(hit["sheet"], "Plasmids")
        self.assertEqual(hit["row"], 2)
        labels = dict(hit["cells"])
        self.assertEqual(labels["Plasmid Name"], "Aquaporin 2")
        self.assertEqual(labels["Vector"], "pcDNA3")

    def test_case_insensitive_and_no_match(self):
        self.assertEqual(len(search.search_rows(self.rows, "XYLAZINE")),
                         len(search.search_rows(self.rows, "xylazine")))
        self.assertGreater(len(search.search_rows(self.rows, "xylazine")), 0)
        self.assertEqual(search.search_rows(self.rows, "zzzz-not-in-lab-qqq"), [])

    def test_menu_option_6_prints_worksheet_names(self):
        out = run_program(["6", "ampicillin", "6", "Canister 4", "6", "nothing-here-xyz", "0"])
        self.assertIn("[Plasmids]  row 2", out)
        self.assertIn("[10L N2 tank]", out)
        self.assertIn("No matches found in any worksheet.", out)

    def test_many_matches_are_limited(self):
        out = run_program(["6", "a", "0"])
        self.assertIn(f"Showing the first {search.MAX_SHOWN} of", out)

    def test_existing_antibody_and_mouse_searches_still_work(self):
        out = run_program(["2", "Alexa", "4", "113-00149-46", "5", "ICR", "0"])
        self.assertIn("Goat anti-Mouse Alexa", out)
        self.assertIn("duplicate cage number", out)
        self.assertIn("ICR", out)


# ===================================================== 2. antibody inventory

class TestAntibodyLoading(Base):
    def setUp(self):
        super().setUp()
        self.items = data_access.load_antibodies(TEST_FILE)

    def test_empty_stock_is_zero(self):
        a = find_antibody(self.items, EMPTY_STOCK_CAT)
        self.assertEqual(a["stock"], 0)
        self.assertTrue(all(isinstance(x["stock"], int) for x in self.items))

    def test_real_zero_stock_stays_zero(self):
        self.assertEqual(find_antibody(self.items, "715-585-151")["stock"], 0)

    def test_stock_counts_tubes(self):
        self.assertEqual(find_antibody(self.items, "711-545-152")["stock"], 2)

    def test_volume_unknown_unless_recorded(self):
        a = find_antibody(self.items, UNKNOWN_VOLUME_CAT)
        self.assertIsNone(a["initial_volume"])
        self.assertIsNone(a["remaining_volume"])
        self.assertEqual(antibodies.volume_text(None), "unknown")

    def test_real_workbook_has_no_invented_volumes(self):
        for a in data_access.load_antibodies(REAL_FILE):
            self.assertIsNone(a["initial_volume"])
            self.assertIsNone(a["remaining_volume"])

    def test_sample_volumes_with_decimals_are_read(self):
        a = find_antibody(self.items, "715-545-151")
        self.assertEqual(a["initial_volume"], 50.5)
        self.assertEqual(a["remaining_volume"], 7.25)

    def test_all_records_loaded(self):
        self.assertEqual(len(self.items), len(data_access.load_antibodies(REAL_FILE)))
        self.assertEqual(len(self.items), 21)


class TestAntibodyLogic(Base):
    def setUp(self):
        super().setUp()
        self.items = data_access.load_antibodies(TEST_FILE)
        self.a = find_antibody(self.items, "115-545-003")     # 100 / 100

    def test_use_subtracts(self):
        self.assertIsNone(antibodies.use_volume(self.a, 30))
        self.assertEqual(self.a["remaining_volume"], 70)
        self.assertEqual(self.a["initial_volume"], 100)       # initial never changes

    def test_use_decimal_has_no_float_noise(self):
        antibodies.use_volume(self.a, 0.1)
        antibodies.use_volume(self.a, 0.2)
        self.assertEqual(self.a["remaining_volume"], 99.7)

    def test_use_everything_is_allowed(self):
        self.assertIsNone(antibodies.use_volume(self.a, 100))
        self.assertEqual(self.a["remaining_volume"], 0)
        self.assertTrue(antibodies.is_low_stock(self.a))

    def test_use_more_than_available_is_refused_and_nothing_changes(self):
        error = antibodies.use_volume(self.a, 100.01)
        self.assertIn("Not enough antibody", error)
        self.assertEqual(self.a["remaining_volume"], 100)

    def test_use_zero_or_negative_refused(self):
        for amount in (0, -1, -0.5, None):
            with self.subTest(amount=amount):
                self.assertIsNotNone(antibodies.use_volume(self.a, amount))
                self.assertEqual(self.a["remaining_volume"], 100)

    def test_use_with_unknown_volume_refused_and_stays_unknown(self):
        unknown = find_antibody(self.items, UNKNOWN_VOLUME_CAT)
        error = antibodies.use_volume(unknown, 1)
        self.assertIn("unknown", error)
        self.assertIsNone(unknown["remaining_volume"])

    def test_parse_volume_valid(self):
        for text, expected in (("5", 5), ("2.5", 2.5), ("0", 0), (" 0.25 ", 0.25), ("1e1", 10)):
            with self.subTest(text=text):
                self.assertEqual(antibodies.parse_volume(text), (expected, None))

    def test_parse_volume_invalid(self):
        for text in ("-1", "-0.1", "abc", "", "nan", "inf", "-inf", "1,5", "1_0", "5 uL", "--2"):
            with self.subTest(text=text):
                value, error = antibodies.parse_volume(text)
                self.assertIsNone(value)
                self.assertIsNotNone(error)

    def test_parse_tubes(self):
        self.assertEqual(antibodies.parse_tubes("3"), (3, None))
        self.assertEqual(antibodies.parse_tubes("0"), (0, None))
        for text in ("-1", "1.5", "abc", "", "²"):
            with self.subTest(text=text):
                self.assertIsNone(antibodies.parse_tubes(text)[0])

    def test_low_stock_boundaries(self):
        by_cat = {c: find_antibody(self.items, c) for c in SAMPLE_VOLUMES}
        self.assertTrue(antibodies.is_low_stock(by_cat["711-545-152"]))      # exactly 5
        self.assertFalse(antibodies.is_low_stock(by_cat["705-545-003"]))     # 5.1
        self.assertFalse(antibodies.is_low_stock(by_cat["115-545-003"]))     # 100
        self.assertFalse(antibodies.is_low_stock(by_cat["715-545-151"]))     # 7.25
        unknown = find_antibody(self.items, UNKNOWN_VOLUME_CAT)
        self.assertFalse(antibodies.is_low_stock(unknown))                   # unknown: no warning
        self.assertEqual(antibodies.low_stock_message(unknown), "")
        self.assertIn("low stock", antibodies.low_stock_message(by_cat["711-545-152"]))

    def test_using_volume_can_trigger_warning(self):
        b = find_antibody(self.items, "715-545-151")        # 7.25
        antibodies.use_volume(b, 2.25)                      # -> exactly 5
        self.assertEqual(b["remaining_volume"], 5)
        self.assertTrue(antibodies.is_low_stock(b))
        self.assertEqual(len(antibodies.low_stock_antibodies(self.items)), 2)

    def test_restock_adds_tubes_and_volume(self):
        self.assertIsNone(antibodies.restock(self.a, 2, 40.5))
        self.assertEqual(self.a["stock"], 3)
        self.assertEqual(self.a["remaining_volume"], 140.5)
        self.assertEqual(self.a["initial_volume"], 140.5)

    def test_restock_tubes_only_keeps_volumes(self):
        self.assertIsNone(antibodies.restock(self.a, 1, None))
        self.assertEqual(self.a["stock"], 2)
        self.assertEqual(self.a["remaining_volume"], 100)

    def test_restock_from_empty_stock_starts_at_zero(self):
        e = find_antibody(self.items, EMPTY_STOCK_CAT)
        antibodies.restock(e, 2, None)
        self.assertEqual(e["stock"], 2)

    def test_restock_with_unknown_volume_stays_unknown(self):
        unknown = find_antibody(self.items, UNKNOWN_VOLUME_CAT)
        antibodies.restock(unknown, 1, 50)
        self.assertEqual(unknown["stock"], 2)
        self.assertIsNone(unknown["remaining_volume"])

    def test_restock_can_clear_low_stock(self):
        low = find_antibody(self.items, "711-545-152")
        antibodies.restock(low, 1, 20)
        self.assertEqual(low["remaining_volume"], 25)
        self.assertFalse(antibodies.is_low_stock(low))

    def test_restock_invalid(self):
        before = dict(self.a)
        for tubes, volume in ((-1, 5), (1, -5), (0, None), (0, 0), (None, 1)):
            with self.subTest(tubes=tubes, volume=volume):
                self.assertIsNotNone(antibodies.restock(self.a, tubes, volume))
                self.assertEqual(self.a, before)

    def test_set_volumes(self):
        unknown = find_antibody(self.items, UNKNOWN_VOLUME_CAT)
        self.assertIsNone(antibodies.set_volumes(unknown, 25.5, 20.25))
        self.assertEqual((unknown["initial_volume"], unknown["remaining_volume"]), (25.5, 20.25))
        before = dict(unknown)
        self.assertIsNotNone(antibodies.set_volumes(unknown, None, 30))   # remaining > initial
        self.assertIsNotNone(antibodies.set_volumes(unknown, 10, None))   # initial < remaining
        self.assertEqual(unknown, before)

    def test_search_antibodies_matches_cat_and_company(self):
        self.assertEqual(len(antibodies.search_antibodies(self.items, "115-545-003")), 1)
        self.assertEqual(len(antibodies.search_antibodies(self.items, "jackson")),
                         len([a for a in self.items if "Jackson" in str(a["company"])]))
        self.assertEqual(len(self.items), 21)


# ============================================= 3. Excel saving and restarting

class TestAntibodyPersistence(Base):
    def test_saved_use_survives_restart(self):
        out = run_program(["7", "115-545-003", "1", "12.5", "0"])
        self.assertIn("Saved. Used 12.5 uL, 87.5 uL remaining.", out)
        # second process = restart; reads the file again
        items = data_access.load_antibodies(TEST_FILE)
        self.assertEqual(find_antibody(items, "115-545-003")["remaining_volume"], 87.5)
        out = run_program(["1", "0"])
        self.assertIn("87.5", out)

    def test_use_that_triggers_warning_prints_it_and_persists(self):
        out = run_program(["7", "715-545-151", "1", "2.5", "0"])
        self.assertIn("WARNING: low stock!", out)
        out = run_program(["10", "0"])        # restarted program
        self.assertIn("LOW STOCK", out)
        self.assertIn("4.75", out)
        self.assertIn("have 5 uL or less left", out)    # start-up warning

    def test_insufficient_stock_is_refused_and_file_unchanged(self):
        before = file_hash(TEST_FILE)
        out = run_program(["7", "715-545-151", "1", "100", "0"])
        self.assertIn("Not enough antibody: only 7.25 uL available", out)
        self.assertEqual(file_hash(TEST_FILE), before)
        self.assertEqual(backups(), [])        # nothing modified -> no backup needed

    def test_invalid_amounts_are_refused_then_valid_amount_accepted(self):
        out = run_program(["7", "115-545-003", "1", "-5", "abc", "nan", "1,5", "0", "2", "0"])
        self.assertEqual(out.count("Invalid input"), 5)
        # 0 is also refused (and asked again), so 5 refusals in total
        self.assertIn("The amount used must be greater than 0.", out)
        self.assertIn("Saved. Used 2 uL, 98 uL remaining.", out)

    def test_use_with_unknown_volume_in_menu(self):
        before = file_hash(TEST_FILE)
        out = run_program(["7", UNKNOWN_VOLUME_CAT, "1", "0"])
        self.assertIn("remaining volume is unknown", out)
        self.assertEqual(file_hash(TEST_FILE), before)

    def test_restock_survives_restart(self):
        run_program(["8", "115-585-003", "1", "3", "40.5", "0"])   # empty stock -> 3 tubes
        a = find_antibody(data_access.load_antibodies(TEST_FILE), EMPTY_STOCK_CAT)
        self.assertEqual(a["stock"], 3)
        self.assertIsNone(a["remaining_volume"])                   # was unknown -> stays unknown
        run_program(["8", "115-545-003", "1", "1", "20.25", "0"])
        a = find_antibody(data_access.load_antibodies(TEST_FILE), "115-545-003")
        self.assertEqual((a["stock"], a["remaining_volume"], a["initial_volume"]), (2, 120.25, 120.25))

    def test_restock_invalid_inputs(self):
        before = file_hash(TEST_FILE)
        out = run_program(["8", "115-545-003", "1", "-1", "1.5", "abc", "", "-3", "x", "0"])
        self.assertGreaterEqual(out.count("Invalid input"), 4)
        out = run_program(["8", "115-545-003", "1", "", "", "0"])     # nothing to restock
        self.assertIn("Nothing to restock", out)
        self.assertEqual(file_hash(TEST_FILE), before)

    def test_enter_volumes_manually_and_restart(self):
        run_program(["9", UNKNOWN_VOLUME_CAT, "1", "30.5", "12.75", "0"])
        a = find_antibody(data_access.load_antibodies(TEST_FILE), UNKNOWN_VOLUME_CAT)
        self.assertEqual((a["initial_volume"], a["remaining_volume"]), (30.5, 12.75))
        # a different record is still unknown (no volume was assumed for it)
        other = find_antibody(data_access.load_antibodies(TEST_FILE), EMPTY_STOCK_CAT)
        self.assertIsNone(other["remaining_volume"])

    def test_enter_volumes_remaining_larger_than_initial_refused(self):
        before = file_hash(TEST_FILE)
        out = run_program(["9", UNKNOWN_VOLUME_CAT, "1", "10", "20", "0"])
        self.assertIn("cannot be larger than the initial volume", out)
        self.assertEqual(file_hash(TEST_FILE), before)

    def test_selecting_invalid_antibody_number(self):
        out = run_program(["7", "Alexa", "99", "0", "x", "", "0"])
        self.assertIn("Please enter a number from 1 to", out)

    def test_antibody_search_with_no_match_in_use_menu(self):
        out = run_program(["7", "zzz-nothing", "0"])
        self.assertIn("No antibodies found.", out)

    def test_stock_cell_and_volume_columns_written_in_excel(self):
        run_program(["8", "115-585-003", "1", "2", "", "0"])
        workbook = openpyxl.load_workbook(TEST_FILE)
        sheet = workbook[data_access.ANTIBODY_SHEET]
        row = find_antibody(data_access.load_antibodies(TEST_FILE), EMPTY_STOCK_CAT)["row"]
        self.assertEqual(sheet.cell(row=row, column=1).value, 2)
        self.assertEqual(sheet["M2"].value, data_access.INITIAL_HEADER)
        self.assertEqual(sheet["N2"].value, data_access.REMAINING_HEADER)
        self.assertTrue(sheet["M2"].font.b)         # header looks like the others


class TestVolumeHeadersOnRealLayout(Base):
    def test_headers_are_created_when_missing(self):
        shutil.copy2(REAL_FILE, TEST_FILE)           # copy WITHOUT our headers
        sheet = openpyxl.load_workbook(TEST_FILE)[data_access.ANTIBODY_SHEET]
        self.assertIsNone(sheet["M2"].value)
        a = data_access.load_antibodies(TEST_FILE)[0]
        a["remaining_volume"] = 1.5
        data_access.save_antibody(TEST_FILE, a)
        sheet = openpyxl.load_workbook(TEST_FILE)[data_access.ANTIBODY_SHEET]
        self.assertEqual(sheet["M2"].value, data_access.INITIAL_HEADER)
        self.assertEqual(sheet["N2"].value, data_access.REMAINING_HEADER)
        self.assertEqual(sheet.cell(row=a["row"], column=14).value, 1.5)
        self.assertIsNone(sheet.cell(row=a["row"], column=13).value)    # initial left empty


# ============================================================ 4. mouse inventory

class TestMouseValidation(unittest.TestCase):
    def test_cage_required(self):
        self.assertIsNotNone(mice.validate_field("cage", "")[1])
        self.assertIsNotNone(mice.validate_field("cage", "-")[1])
        self.assertEqual(mice.validate_field("cage", " 113-00149-99 "), ("113-00149-99", None))

    def test_counts(self):
        self.assertEqual(mice.validate_field("males", "3"), (3, None))
        self.assertEqual(mice.validate_field("males", "0"), (0, None))
        self.assertEqual(mice.validate_field("females", ""), (None, None))     # unknown
        self.assertEqual(mice.validate_field("females", "-"), (None, None))
        self.assertEqual(mice.validate_field("females", "unknown"), (None, None))
        for text in ("-2", "1.5", "abc", "2 mice", "²"):
            with self.subTest(text=text):
                self.assertIsNotNone(mice.validate_field("males", text)[1])

    def test_number(self):
        self.assertEqual(mice.validate_field("number", "12"), (12, None))
        for text in ("0", "-1", "", "x", "1.5"):
            with self.subTest(text=text):
                self.assertIsNotNone(mice.validate_field("number", text)[1])

    def test_birth_date(self):
        for text, expected in (("2026.03.11", "2026.03.11"), ("2026-3-1", "2026.03.01"),
                               ("2026/12/31", "2026.12.31"), ("2026.12.", "2026.12."),
                               ("2026.12", "2026.12.")):
            with self.subTest(text=text):
                self.assertEqual(mice.validate_field("birth_date", text), (expected, None))
        for text in ("2026.02.30", "2026.13.01", "yesterday", "11.03.2026", "2026", "20260311", "2026..03"):
            with self.subTest(text=text):
                self.assertIsNotNone(mice.validate_field("birth_date", text)[1])
        self.assertEqual(mice.validate_field("birth_date", "-"), (None, None))

    def test_text_fields_and_formula_protection(self):
        self.assertEqual(mice.validate_field("label", "WT"), ("WT", None))
        self.assertEqual(mice.validate_field("comment", ""), (None, None))
        for field in ("cage", "label", "comment", "owner", "background", "experiment"):
            self.assertIsNotNone(mice.validate_field(field, "=1+1")[1])


class TestMouseLoading(Base):
    def test_unknown_counts_stay_unknown(self):
        items = data_access.load_mice(TEST_FILE)
        first = items[0]            # 112-00111-03: males empty, females 4
        self.assertIsNone(first["males"])
        self.assertEqual(first["females"], 4)
        self.assertEqual(len(items), 35)
        self.assertTrue(any(m["males"] is None and m["females"] is None for m in items))

    def test_duplicate_cage_is_found(self):
        items = data_access.load_mice(TEST_FILE)
        self.assertEqual(mice.find_duplicate_cages(items), [DUPLICATE_CAGE])
        same = [m for m in items if m["cage"] == DUPLICATE_CAGE]
        self.assertEqual(len(same), 2)
        self.assertNotEqual(same[0]["row"], same[1]["row"])


class TestMouseEditing(Base):
    EDIT_KEEP = [""] * 10

    def test_edit_all_columns_and_restart(self):
        items = data_access.load_mice(TEST_FILE)
        target = items[0]           # 112-00111-03
        out = run_program(["11", "112-00111-03", "1",
                           "50",                # number
                           "112-00111-03-X",    # cage
                           "ICR",               # background
                           "2",                 # males
                           "5",                 # females
                           "2026.10.01",        # birth date
                           "KO",                # label
                           "TestOwner",         # owner
                           "TestExperiment",    # experiment
                           "Test comment",      # comment
                           "y", "0"])
        self.assertIn("Saved.", out)
        changed = [m for m in data_access.load_mice(TEST_FILE) if m["row"] == target["row"]][0]
        self.assertEqual(changed, {"row": target["row"], "number": 50, "cage": "112-00111-03-X",
                                   "background": "ICR", "males": 2, "females": 5,
                                   "birth_date": "2026.10.01", "label": "KO", "owner": "TestOwner",
                                   "experiment": "TestExperiment", "comment": "Test comment"})
        # all other cages are untouched
        others_before = [m for m in items if m["row"] != target["row"]]
        others_after = [m for m in data_access.load_mice(TEST_FILE) if m["row"] != target["row"]]
        self.assertEqual(others_before, others_after)

    def test_edit_only_selected_duplicate(self):
        before = data_access.load_mice(TEST_FILE)
        same = [m for m in before if m["cage"] == DUPLICATE_CAGE]
        # choose the SECOND listed record; change its label only
        out = run_program(["11", DUPLICATE_CAGE, "2", "", "", "", "", "", "", "EDITED", "", "", "", "y", "0"])
        self.assertIn("(duplicate cage no.)", out)
        after = {m["row"]: m for m in data_access.load_mice(TEST_FILE)}
        self.assertEqual(after[same[0]["row"]], same[0])                       # first unchanged
        self.assertEqual(after[same[1]["row"]]["label"], "EDITED")             # second changed
        self.assertEqual({k: v for k, v in after[same[1]["row"]].items() if k != "label"},
                         {k: v for k, v in same[1].items() if k != "label"})

    def test_edit_first_duplicate_leaves_second(self):
        before = data_access.load_mice(TEST_FILE)
        same = [m for m in before if m["cage"] == DUPLICATE_CAGE]
        run_program(["11", DUPLICATE_CAGE, "1", "", "", "", "", "", "", "", "", "", "FIRST ONLY", "y", "0"])
        after = {m["row"]: m for m in data_access.load_mice(TEST_FILE)}
        self.assertEqual(after[same[0]["row"]]["comment"], "FIRST ONLY")
        self.assertEqual(after[same[1]["row"]], same[1])

    def test_missing_counts_stay_unknown_when_editing_other_fields(self):
        before = data_access.load_mice(TEST_FILE)
        target = before[0]          # males unknown
        run_program(["11", "112-00111-03", "1", "", "", "", "", "", "", "NEWLABEL", "", "", "", "y", "0"])
        after = [m for m in data_access.load_mice(TEST_FILE) if m["row"] == target["row"]][0]
        self.assertIsNone(after["males"])
        self.assertEqual(after["label"], "NEWLABEL")
        sheet = openpyxl.load_workbook(TEST_FILE)[data_access.MOUSE_SHEET]
        self.assertIsNone(sheet.cell(row=target["row"], column=4).value)    # still empty, not 0

    def test_clear_count_with_dash_makes_it_unknown(self):
        run_program(["11", "112-00111-03", "1", "", "", "", "", "-", "", "", "", "", "", "y", "0"])
        first = data_access.load_mice(TEST_FILE)[0]
        self.assertIsNone(first["females"])      # was 4
        run_program(["11", "112-00111-03", "1", "", "", "", "0", "", "", "", "", "", "", "y", "0"])
        first = data_access.load_mice(TEST_FILE)[0]
        self.assertEqual(first["males"], 0)      # a real 0 is kept as 0
        self.assertIsNone(first["females"])

    def test_invalid_input_is_rejected_and_nothing_saved(self):
        before_hash = file_hash(TEST_FILE)
        out = run_program(["11", "112-00111-03", "1",
                           "abc", "0", "50",            # number: bad, bad, ok
                           "=SUM(1)", "112-00111-03",   # cage: formula text rejected, then ok (unchanged)
                           "", "-3", "x", "2",          # background keep; males: bad, bad, ok
                           "", "2026.02.30", "2026.03.11",   # females keep; birth date bad then ok
                           "", "", "", "", "n", "0"])   # keep rest; DO NOT save
        self.assertGreaterEqual(out.count("Invalid input"), 6)
        self.assertIn("Cancelled. Nothing was saved.", out)
        self.assertEqual(file_hash(TEST_FILE), before_hash)

    def test_invalid_then_saved_values_are_the_valid_ones(self):
        run_program(["11", "112-00111-03", "1", "50", "", "", "-1", "3", "", "", "", "", "", "", "y", "0"])
        # number=50 ok, males '-1' invalid -> re-asked -> '3'
        first = data_access.load_mice(TEST_FILE)[0]
        self.assertEqual((first["number"], first["males"]), (50, 3))

    def test_duplicate_cage_number_warning_on_edit(self):
        items = data_access.load_mice(TEST_FILE)
        other = items[1]["cage"]
        # try to give record 1 the cage number of record 2; answer 'n' -> not changed
        out = run_program(["11", "112-00111-03", "1", "", other, "", "", "", "", "", "", "", "", "n", "0"])
        self.assertIn("already exists (Excel row", out)
        self.assertEqual(data_access.load_mice(TEST_FILE)[0]["cage"], "112-00111-03")
        # answer 'y' -> allowed
        run_program(["11", "112-00111-03", "1", "", other, "y", "", "", "", "", "", "", "", "", "y", "0"])
        after = data_access.load_mice(TEST_FILE)
        self.assertEqual(after[0]["cage"], other)
        self.assertIn(other, mice.find_duplicate_cages(after))

    def test_edit_selection_errors_and_cancel(self):
        before_hash = file_hash(TEST_FILE)
        out = run_program(["11", "nonexistent-cage", "11", "112-00111", "99", "0", "x", "", "0"])
        self.assertIn("No cages found.", out)
        self.assertIn("Please enter a number from 1 to", out)
        self.assertEqual(file_hash(TEST_FILE), before_hash)

    def test_edit_nothing_changed(self):
        before_hash = file_hash(TEST_FILE)
        out = run_program(["11", "112-00111-03", "1"] + self.EDIT_KEEP + ["0"])
        self.assertIn("Nothing changed.", out)
        self.assertEqual(file_hash(TEST_FILE), before_hash)

    def test_stale_row_check_refuses_to_write(self):
        record = dict(data_access.load_mice(TEST_FILE)[0])
        record["row"] += 1          # pretend the file changed under us
        with self.assertRaises(ValueError):
            data_access.save_mouse(TEST_FILE, record, record["cage"])


class TestMouseAdding(Base):
    def test_add_cage_survives_restart_and_extends_totals(self):
        before = data_access.load_mice(TEST_FILE)
        old_last_row = max(m["row"] for m in before)
        run_program(["12", "", "113-00200-01", "C57BL/6", "3", "", "2026.11.20", "WT",
                     "Tester", "Exp A", "added by test", "y", "0"])
        after = data_access.load_mice(TEST_FILE)           # fresh process-independent read
        self.assertEqual(len(after), len(before) + 1)
        new = after[-1]
        self.assertEqual(new, {"row": old_last_row + 1, "number": 38, "cage": "113-00200-01",
                               "background": "C57BL/6", "males": 3, "females": None,
                               "birth_date": "2026.11.20", "label": "WT", "owner": "Tester",
                               "experiment": "Exp A", "comment": "added by test"})
        self.assertEqual(after[:-1], before)               # old records unchanged
        sheet = openpyxl.load_workbook(TEST_FILE)[data_access.MOUSE_SHEET]
        # totals moved down one row and now include the new row
        self.assertEqual(sheet["D48"].value, "=SUM(D11:D46)")
        self.assertEqual(sheet["E48"].value, "=SUM(E11:E46)")
        self.assertEqual(sheet["C48"].value, "總量")
        self.assertEqual(sheet["D50"].value, "=SUM(D11:E21)")       # unrelated formulas unchanged
        self.assertEqual(sheet["D51"].value, "=sum(D22:E35)")
        self.assertEqual(sheet["C50"].value, "C57 (檢查前尚有8隻會犧牲)")
        self.assertIsNone(sheet["C47"].value)                         # blank spacer row kept
        self.assertEqual(sheet["K9"].value, "=SUM(K3:K7)+80")         # top block untouched
        # new row has the formatting of the last cage row
        self.assertEqual(sheet["B46"].alignment.horizontal, sheet["B45"].alignment.horizontal)
        self.assertEqual(sheet["B46"].font.name, sheet["B45"].font.name)
        # the program shows it after restart
        out = run_program(["4", "113-00200-01", "0"])
        self.assertIn("113-00200-01", out)

    def test_add_twice_extends_totals_twice(self):
        for cage in ("113-00300-01", "113-00300-02"):
            run_program(["12", "", cage, "", "", "", "", "", "", "", "", "y", "0"])
        sheet = openpyxl.load_workbook(TEST_FILE)[data_access.MOUSE_SHEET]
        self.assertEqual(sheet["D49"].value, "=SUM(D11:D47)")
        self.assertEqual(sheet["B46"].value, "113-00300-01")
        self.assertEqual(sheet["B47"].value, "113-00300-02")
        self.assertEqual(sheet["A47"].value, 39)
        self.assertEqual(len(data_access.load_mice(TEST_FILE)), 37)

    def test_missing_counts_stay_unknown_when_adding(self):
        run_program(["12", "", "113-00400-01", "", "", "", "", "", "", "", "", "y", "0"])
        new = data_access.load_mice(TEST_FILE)[-1]
        self.assertIsNone(new["males"])
        self.assertIsNone(new["females"])
        self.assertIsNone(new["birth_date"])
        sheet = openpyxl.load_workbook(TEST_FILE)[data_access.MOUSE_SHEET]
        self.assertIsNone(sheet["D46"].value)
        self.assertIsNone(sheet["E46"].value)

    def test_add_duplicate_cage_number_asks_first(self):
        n = len(data_access.load_mice(TEST_FILE))
        out = run_program(["12", "", DUPLICATE_CAGE, "", "", "", "", "", "", "", "", "n", "0"])
        self.assertIn("already exists (Excel row", out)
        self.assertEqual(len(data_access.load_mice(TEST_FILE)), n)
        run_program(["12", "", DUPLICATE_CAGE, "y", "", "", "", "", "", "", "", "", "y", "0"])
        items = data_access.load_mice(TEST_FILE)
        self.assertEqual(len(items), n + 1)
        self.assertEqual(len([m for m in items if m["cage"] == DUPLICATE_CAGE]), 3)

    def test_add_invalid_input_rejected(self):
        before_hash = file_hash(TEST_FILE)
        out = run_program(["12", "abc", "-4", "7",              # number: bad, bad, ok
                           "=1+1", "113-00500-01",         # cage: formula text bad, ok
                           "", "-1", "x", "1",             # background; males: bad, bad, ok
                           "", "13.45.99", "2026.01.01",   # females; birth date: bad, ok
                           "", "", "", "", "n", "0"])      # label, owner, experiment, comment; DO NOT save
        self.assertGreaterEqual(out.count("Invalid input"), 6)
        self.assertEqual(file_hash(TEST_FILE), before_hash)     # declined -> nothing saved

    def test_add_without_cage_number_is_cancelled(self):
        before_hash = file_hash(TEST_FILE)
        out = run_program(["12", "", "", "0"])
        self.assertIn("The cage number is required", out)
        self.assertEqual(file_hash(TEST_FILE), before_hash)


# ============================================ 5. workbook safety and integrity

class TestWorkbookSafety(Base):
    def test_backup_is_created_before_change_and_equals_old_file(self):
        original = file_hash(TEST_FILE)
        self.assertEqual(backups(), [])
        run_program(["7", "115-545-003", "1", "1", "0"])
        files = backups()
        self.assertEqual(len(files), 1)
        self.assertEqual(file_hash(files[0]), original)         # backup = state BEFORE the change
        self.assertNotEqual(file_hash(TEST_FILE), original)
        run_program(["7", "115-545-003", "1", "1", "0"])
        self.assertEqual(len(backups()), 2)
        self.assertEqual(len(set(backups())), 2)

    def test_backup_folder_is_next_to_workbook_and_relative(self):
        run_program(["7", "115-545-003", "1", "1", "0"])
        self.assertTrue(os.path.isdir(os.path.join(TEST_DIR, "backups")))
        self.assertFalse(os.path.exists(os.path.join(PROJECT, "backups")))
        self.assertFalse(os.path.exists(TEST_FILE + ".tmp"))

    def test_other_content_and_formatting_preserved_after_antibody_and_mouse_changes(self):
        before = snapshot(TEST_FILE)
        run_program(["7", "115-545-003", "1", "10.5",
                     "8", "115-585-003", "1", "2", "",
                     "11", "112-00111-03", "1", "", "", "", "9", "", "", "NEWLABEL", "", "", "", "y",
                     "0"])
        after = snapshot(TEST_FILE)
        self.assertEqual(before["sheets"], after["sheets"])
        self.assertEqual(len(after["sheets"]), 14)

        anti = data_access.ANTIBODY_SHEET
        mouse = data_access.MOUSE_SHEET
        a_row = find_antibody(data_access.load_antibodies(TEST_FILE), "115-545-003")["row"]
        b_row = find_antibody(data_access.load_antibodies(TEST_FILE), "115-585-003")["row"]
        m_row = data_access.load_mice(TEST_FILE)[0]["row"]
        expected_changes = {anti: {(a_row, 14), (b_row, 1)}, mouse: {(m_row, 4), (m_row, 7)}}

        for name in before["sheets"]:
            b, a = before[name], after[name]
            self.assertEqual(b["merged"], a["merged"], name)
            self.assertEqual(b["cf"], a["cf"], name)
            self.assertEqual(b["tables"], a["tables"], name)
            self.assertEqual(b["autofilter"], a["autofilter"], name)
            self.assertEqual(b["widths"], a["widths"], name)
            for key in set(b["cells"]) | set(a["cells"]):
                if key in expected_changes.get(name, set()):
                    continue
                self.assertEqual(b["cells"].get(key), a["cells"].get(key), f"{name} {key}")
        # the changes themselves are present
        self.assertEqual(after[anti]["cells"][(a_row, 14)][0], 89.5)
        self.assertEqual(after[anti]["cells"][(b_row, 1)][0], 2)
        self.assertEqual(after[mouse]["cells"][(m_row, 4)][0], 9)
        self.assertEqual(after[mouse]["cells"][(m_row, 7)][0], "NEWLABEL")

    def test_formulas_still_formulas_and_other_sheets_byte_for_byte_equal_in_values(self):
        run_program(["7", "115-545-003", "1", "1", "0"])
        sheet = openpyxl.load_workbook(TEST_FILE)[data_access.MOUSE_SHEET]
        for coord, formula in (("K9", "=SUM(K3:K7)+80"), ("L9", "=SUM(L3:L7)+65"),
                               ("D47", "=SUM(D11:D45)"), ("E47", "=SUM(E11:E45)"),
                               ("D49", "=SUM(D11:E21)"), ("D50", "=sum(D22:E35)")):
            self.assertEqual(sheet[coord].value, formula)

    def test_save_failure_leaves_file_and_memory_unchanged(self):
        import main
        items = data_access.load_antibodies(TEST_FILE)
        a = find_antibody(items, "115-545-003")
        old = dict(a)
        before = file_hash(TEST_FILE)
        antibodies.use_volume(a, 10)
        with mock.patch.object(data_access, "finish_saving", side_effect=PermissionError("file is open")):
            ok = main.save_or_undo(lambda: data_access.save_antibody(TEST_FILE, a), a, old)
        self.assertFalse(ok)
        self.assertEqual(a, old)                         # screen = file
        self.assertEqual(file_hash(TEST_FILE), before)   # file not damaged
        self.assertFalse(os.path.exists(TEST_FILE + ".tmp"))

    def test_real_workbook_untouched_by_test_runs(self):
        run_program(["7", "115-545-003", "1", "1", "0"])
        self.assertEqual(file_hash(REAL_FILE), REAL_HASH_AT_START)

    def test_missing_file_message(self):
        out = run_program([], path=os.path.join(TEST_DIR, "does_not_exist.xlsx"))
        self.assertIn("Could not find the Excel file", out)

    def test_menu_invalid_choice_and_clean_quit(self):
        out = run_program(["99", "abc", "", "0"])
        self.assertEqual(out.count("Invalid choice"), 3)
        self.assertIn("Goodbye!", out)

    def test_input_ending_does_not_crash(self):
        out = run_program(["1"])        # no '0': input ends
        self.assertIn("Goodbye!", out)


if __name__ == "__main__":
    unittest.main()
