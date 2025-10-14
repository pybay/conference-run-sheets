"""
utility to read conference info from Sessionize and create Conference Run Sheets for the Organizer and Volunteer Staff
"""
from pathlib import Path

from core_utils.create_run_sheets import RunSheetCollection
from core_utils.save_run_sheets_excel import ExcelRunSheetWriter

if __name__ == "__main__":
    input_filename = "pybay2025 flattened accepted sessions - exported 2025-10-13.xlsx"
    output_filename = "pybay2025_run_sheets.xlsx"
    input_path = Path(".", input_filename)
    output_path = Path(".", output_filename)

    results = RunSheetCollection(sessionize_input_path=input_path).organize_data()

    ExcelRunSheetWriter(results=results, sessionize_output_path=output_path).create_sheets()
