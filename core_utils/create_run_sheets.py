"""
utility to organize the data and create the run sheets in an Excel file
"""
from pathlib import Path
import re
from typing import Any

import pandas as pd
from pandas import DataFrame

from core_utils.get_input import get_input_from_sessionize

# shortlist of columns we expect for the runsheet
# NOTE talks with more than one speaker, will have multiple entries - once for each speaker
ALL_EXPECTED_COLUMN_SUBSET = [
    "First name - pronunciation",
    "Last name - pronunciation",
    "Mobile # with Country Code (not shared publicly)",
    "Owner",
    "Profile Picture",
    "Pronouns",
    "Room",
    "Scheduled Duration",
    "Session format",
    "Session Id",
    "Speaker introduction - bullet 1",
    "Speaker introduction - bullet 2",
    "Speaker introduction - bullet 3",
    "This would be my first Conference Talk",
    "Scheduled At",
    "Title",
    "What will attendees learn?"
]

COLUMNS_DETAIL = ALL_EXPECTED_COLUMN_SUBSET
COLUMN_ORDER_SUMMARY = ["Room", "Time", "Title", "Speaker"]
COLUMN_ORDER_DETAIL_ORIGINAL = [
    "Room",
    "Time",
    "Title",
    "Scheduled Duration",
    "What will attendees learn?",
    "Speaker",
    "Profile Picture",
    "First name - pronunciation",
    "Last name - pronunciation",
    "Mobile # with Country Code (not shared publicly)",
    "Pronouns",
    "This would be my first Conference Talk",
    "Speaker introduction - bullet 1",
    "Speaker introduction - bullet 2",
    "Speaker introduction - bullet 3"
]

COLUMN_RENAME_MAP = {
    "Owner": "Speaker",
    "What will attendees learn?": "Attendees Learn",
    "Mobile # with Country Code (not shared publicly)": "MOBILE # - PRIVATE!",
    "This would be my first Conference Talk": "First Conf Talk",
    "Profile Picture": "Profile Photo",
    "Scheduled Duration": "Duration",
    "Speaker introduction - bullet 1": "Speaker intro #1",
    "Speaker introduction - bullet 2": "Speaker intro #2",
    "Speaker introduction - bullet 3": "Speaker intro #3",
}

COLUMN_ORDER_DETAIL = [
    "Room",
    "Time",
    "Duration",
    "Title",
    "Speaker",
    "First name - pronunciation",
    "Last name - pronunciation",
    "MOBILE # - PRIVATE!",
    "Pronouns",
    "First Conf Talk",
    "Profile Photo",
    "Attendees Learn",
    "Speaker intro #1",
    "Speaker intro #2",
    "Speaker intro #3"
]


def format_phone_number(phone):
    """Format phone number to XXX.XXX.XXXX"""
    # Handle empty/invalid values
    if pd.isna(phone) or phone == "Not Provided":
        result = "Not Provided"
    else:
        # Remove all non-numeric characters
        digits = re.sub(r'\D', '', str(phone))

        # Strip leading country code for US ONLY if present (end user readability)
        if digits.startswith('1') and len(digits) == 11:
            digits = digits[1:]

        # Format based on digit count (end user readability)
        match len(digits):
            case 10:
                result = f"{digits[:3]}.{digits[3:6]}.{digits[6:]}"
            case _:
                result = digits or "Not Provided"

    return result

class RunSheetCollection:
    """Container for organized run sheet DataFrames."""

    def __init__(self, sessionize_input_path: Path):
        self.sessionize_input_path = Path(sessionize_input_path)
        self.df_core: DataFrame = get_input_from_sessionize(sessionize_input_path)

        results = self.organize_data()
        # IMPORTANT!  `_summary` and `_detail` are critical keywords for df and sheet names
        self.df_core_sorted = results.get('df_core_sorted')
        self.robertson_summary = results['robertson_summary']
        self.robertson_detail = results['robertson_detail']
        self.fisher_summary = results['fisher_summary']
        self.fisher_detail = results['fisher_detail']
        self.workshop_summary = results['workshop_summary']
        self.workshop_detail = results['workshop_detail']

    def to_dict(self) -> dict[str, DataFrame]:
        """Convert to dictionary, excluding None values."""
        return {k: v for k, v in self.__dict__.items() if v is not None}

    def organize_data(self) -> dict[str, Any]:
        """
        Filters and organizes the data for each runsheet into separate dataframes
        :param self: this class
        :return: Dict mapping run sheet names to DataFrames
                - 'robertson_summary': Robertson room summary view (time, title, speaker)
                - 'robertson_detail': Robertson room detailed view (includes contact info, intro notes)
                - 'fisher_summary': Fisher room summary view
                - 'fisher_detail': Fisher room detailed view
                - 'workshop_summary': Workshop room summary view (if applicable)
                - 'workshop_detail': Workshop room detailed view (if applicable)
        """
        df_core = self.df_core[ALL_EXPECTED_COLUMN_SUBSET].copy().sort_values(
            by=["Room", "Scheduled At", "Session Id", "Owner"], ascending=True
        )

        # Data Cleanup
        df_core = df_core.rename(columns=COLUMN_RENAME_MAP)

        df_core["MOBILE # - PRIVATE!"] = df_core["MOBILE # - PRIVATE!"].apply(format_phone_number)
        df_core["Pronouns"] = df_core["Pronouns"].str.title()

        # Replace "Not Provided" with 7pm (using a dummy date since datetime needs both date and time)
        df_core.loc[df_core["Scheduled At"] == "Not Provided", "Scheduled At"] = "2025-10-18 19:00:00"
        df_core["Scheduled At"] = pd.to_datetime(df_core["Scheduled At"])
        df_core["Time"] = df_core["Scheduled At"].dt.strftime("%I:%M %p")  # time only, more readable in run sheets

        # Extract conference year from scheduled dates (most common year in the data)
        conference_year = int(df_core["Scheduled At"].dt.year.mode()[0])
        df_core["Session format"] = df_core["Session format"].astype(str).str[:2]
        # alternate speakers don't have assigned/scheduled rooms - update to "Alternate" so they can go to any room
        df_core.loc[df_core["Room"] == "Not Provided", "Room"] = "Alternate Speaker - ANY room"
        fallback_duration_mask = df_core["Duration"] == "Not Provided"
        df_core.loc[fallback_duration_mask, "Duration"] = df_core.loc[
            fallback_duration_mask, "Session format"]
        df_core["Duration"] = df_core["Duration"].astype(int)

        df_core = df_core[COLUMN_ORDER_DETAIL]  # helpfully order columns AFTER adding usefully named 'TIME' column

        # Formatting
        # df_core["Scheduled Duration"] = df_core["Scheduled Duration"].astype(int)

        df_core_sorted = df_core.copy()

        # dataframes for each room
        df_robertson_summary = df_core_sorted[df_core_sorted["Room"].str.contains("Robertson")][COLUMN_ORDER_SUMMARY]
        df_robertson_detail = df_core_sorted[
            df_core_sorted["Room"].str.contains("Robertson")
        ][COLUMN_ORDER_DETAIL]
        df_fisher_detail = df_core_sorted[
            df_core_sorted["Room"].str.contains("Fisher")
        ][COLUMN_ORDER_DETAIL]

        df_fisher_summary = df_core_sorted[df_core_sorted["Room"].str.contains("Fisher")][COLUMN_ORDER_SUMMARY]

        df_workshop_summary = df_core_sorted[df_core_sorted["Room"].str.contains("Workshop")][COLUMN_ORDER_SUMMARY]
        df_workshop_detail = df_core_sorted[
            df_core_sorted["Room"].str.contains("Workshop")
        ][COLUMN_ORDER_DETAIL]

        return {
            "conference_year": conference_year,
            "df_core_sorted": df_core_sorted,
            "robertson_summary": df_robertson_summary,
            "robertson_detail": df_robertson_detail,
            "fisher_summary": df_fisher_summary,
            "fisher_detail": df_fisher_detail,
            "workshop_summary": df_workshop_summary,
            "workshop_detail": df_workshop_detail,
        }
