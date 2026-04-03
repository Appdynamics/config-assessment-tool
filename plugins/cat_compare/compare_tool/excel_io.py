"""
excel_io.py
-----------
This module provides utility functions for working with Excel files.

Purpose:
- Handles operations such as saving workbooks and validating data consistency.
- Ensures formulas in Excel files are recalculated before processing.

Key Features:
- `save_workbook`: Opens and saves an Excel workbook using `xlwings` to ensure formulas are recalculated.
- `check_controllers_match`: Validates that two Excel files have matching controller values.
"""

import logging
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import load_workbook
import xlwings as xw


def save_workbook(filepath: str) -> None:
    """
    Open and save the workbook in Excel so formulas are recalculated
    before we read it with pandas/openpyxl.
    """
    path = Path(filepath).resolve()
    logging.info("Saving workbook via Excel: %s", path)

    app = xw.App(visible=False)
    try:
        wb = app.books.open(str(path))
        wb.save()
    finally:
        # Always try to close/quit even if something goes wrong
        try:
            wb.close()
        except Exception:
            pass
        app.quit()


def check_controllers_match(previous_file_path: str, current_file_path: str) -> bool:
    """
    Ensure both workbooks have a single, matching controller value
    in the Analysis sheet's 'controller' column.
    """
    try:
        prev_df = pd.read_excel(previous_file_path, sheet_name="Analysis")
        curr_df = pd.read_excel(current_file_path, sheet_name="Analysis")
    except Exception as e:
        logging.error("Failed to read 'Analysis' sheet from one of the files: %s", e)
        return False

    if "controller" not in prev_df.columns or "controller" not in curr_df.columns:
        logging.error("Missing 'controller' column in one of the Analysis sheets.")
        return False

    prev_ctrls = prev_df["controller"].dropna().astype(str).str.strip().unique()
    curr_ctrls = curr_df["controller"].dropna().astype(str).str.strip().unique()

    logging.debug(f"Previous controller(s): {prev_ctrls}")
    logging.debug(f"Current controller(s): {curr_ctrls}")

    if len(prev_ctrls) != 1 or len(curr_ctrls) != 1:
        logging.error(
            "Controller column does not contain exactly one unique value in each file."
        )
        return False

    if prev_ctrls[0] != curr_ctrls[0]:
        logging.error(
            f"Controllers do not match: {prev_ctrls[0]} vs {curr_ctrls[0]}"
        )
        return False

    return True


def get_key_column(worksheet, header_name: str) -> Optional[int]:
    """
    Find the 1-based column index for a header in an openpyxl worksheet.
    Used by comparison routines to align columns.
    """
    try:
        header_row = next(worksheet.iter_rows(min_row=1, max_row=1))
    except StopIteration:
        return None

    for cell in header_row:
        if str(cell.value or "").strip() == header_name:
            return cell.column
    return None
