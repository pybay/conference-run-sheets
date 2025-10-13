"""
utility to read conference info from Sessionize and create Conference Run Sheets for the Organizer and Volunteer Staff
"""
from core_utils.create_run_sheets import RunSheetCollection
from pathlib import Path

if __name__ == "__main__":
    filename = "pybay2025 flattened accepted sessions - exported 2025-10-13.xlsx"
    input_path = Path(".", filename)

    run_sheets = RunSheetCollection(sessionize_input_path=input_path)

    results = run_sheets.organize_data()
