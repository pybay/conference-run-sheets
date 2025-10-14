"""
utility for saving run sheets to an Excel workbook, with formatting

Intended to deliver final, formatted run sheets ready for use
"""
from abc import ABC, abstractmethod
from pathlib import Path
from core_utils.create_run_sheets import COLUMN_ORDER_DETAIL, COLUMN_ORDER_SUMMARY
from pandas import DataFrame


class RunSheetSaveManager(ABC):
    """
    Base Class for managing creation of run sheets into spreadsheet/workbook with multiple tabs.  Handles validation, sheet filters and naming.
    Separate subclasses handle actual creating and formatting of Excel and Google Sheets.

    Why two methods?
    - Excel - for users who have not configured their Google Sheets API to reach the PyBay Google Drive sheet for current year; workaround is for them to save result as Excel and manually upload.  With a volunteer team having high turnover, plus Google Sheets API setup complexity we need a low friction way to get results in hand so team's don't revert to "fully manual" processing of Sessionize speaker output (bad!).
    - Google Sheet - for users who  have correctly configured their Google Sheets API, and can write directly to Google Sheets
    """

    # we only want tabs needed for final Run Sheets; and use specific patterns
    VALID_SHEET_PATTERNS = ["_summary", "_detail"]
    EXCLUDE_SHEET_PATTERNS = ["df_"]

    def __init__(self, results: dict[str, DataFrame], sessionize_output_path: Path):
        """
        Args:
            results: Dict of DataFrames with sheet data (must include 'conference_year' key)
            sessionize_output_path: Path (Excel) or Sheet ID/name (Google Sheets)
        """
        self.results = results
        self.sessionize_output_path = sessionize_output_path
        self.conference_year = results.get('conference_year', 'YYYY')  # Extract year for headers
        self.sheet_keys = self._validate_sheet_keys()
        self.COLUMN_ORDER_DETAIL = COLUMN_ORDER_DETAIL
        self.COLUMN_ORDER_SUMMARY = COLUMN_ORDER_SUMMARY

    @abstractmethod
    def _setup(self):
        """Initialize output (Excel workbook, Google Sheet connection, etc.)"""
        pass

    @abstractmethod
    def _write_sheet(self, df: DataFrame, sheet_name: str, sheet_type: str):
        """Write a single sheet with appropriate formatting."""
        pass

    @abstractmethod
    def _finalize(self):
        """Cleanup and close (save Excel/Google workbook, close API, notify user, etc.)"""
        pass

    @staticmethod
    def _get_sheet_type(key: str) -> str:
        """Determine if sheet is summary or detail - which have different content and formatting"""
        if 'summary' in key:
            return 'summary'
        elif 'detail' in key:
            return 'detail'
        else:
            raise ValueError(f"Unknown sheet type for key: {key}")

    def _validate_sheet_keys(self) -> list[str]:
        valid_keys = []

        for key in self.results.keys():
            # Skip metadata keys
            if key == 'conference_year':
                continue

            # Check if key starts with any exclude pattern
            if any(key.startswith(pattern) for pattern in self.EXCLUDE_SHEET_PATTERNS):
                continue  # Skip this key

            if any(pattern in key for pattern in self.VALID_SHEET_PATTERNS):
                valid_keys.append(key)
            else:
                print(f"Warning: Ignoring unexpected key '{key}'")

        return sorted(valid_keys, key=lambda x: (x.split('_')[0], 'detail' in x))

    def create_sheets(self):
        """Main entry point - creates all sheets with appropriate formatting."""
        self._setup()

        for sheet_key in self.sheet_keys:
            df = self.results[sheet_key]
            sheet_name = sheet_key
            sheet_type = self._get_sheet_type(sheet_key)

            self._write_sheet(df, sheet_name, sheet_type)

        self._finalize()
