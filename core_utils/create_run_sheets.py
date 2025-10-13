"""
utility to organize the data and create the run sheets in an Excel file
"""
from dataclasses import dataclass
import pandas as pd
from pathlib import Path
from typing import Any

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
COLUMN_ORDER_DETAIL = [
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


class RunSheetCollection:
    """Container for organized run sheet DataFrames."""

    def __init__(self, sessionize_input_path: Path):
        self.sessionize_input_path = Path(sessionize_input_path)
        self.df_core: DataFrame = get_input_from_sessionize(sessionize_input_path)

        results = self.organize_data()
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
        # Replace "Not Provided" with 7pm (using a dummy date since datetime needs both date and time)
        df_core.loc[df_core["Scheduled At"] == "Not Provided", "Scheduled At"] = "2025-10-18 19:00:00"
        df_core["Scheduled At"] = pd.to_datetime(df_core["Scheduled At"])
        df_core["Time"] = df_core["Scheduled At"].dt.strftime("%I:%M %p")  # time only, more readable in run sheets
        df_core["Speaker"] = df_core["Owner"]
        df_core = df_core.drop(["Owner", "Scheduled At"], axis=1)
        df_core["Session format"] = df_core["Session format"].astype(str).str[:2]
        # alternate speakers don't have assigned/scheduled rooms - update to "Alternate" so they can go to any room
        df_core.loc[df_core["Room"] == "Not Provided", "Room"] = "Alternate Speaker - ANY room"

        fallback_duration_mask = df_core["Scheduled Duration"] == "Not Provided"
        df_core.loc[fallback_duration_mask, "Scheduled Duration"] = df_core.loc[
            fallback_duration_mask, "Session format"]
        df_core["Scheduled Duration"] = df_core["Scheduled Duration"].astype(int)

        df_core = df_core[COLUMN_ORDER_DETAIL]  # helpfully order columns AFTER adding usefully named 'TIME' column

        # Formatting
        # df_core["Scheduled Duration"] = df_core["Scheduled Duration"].astype(int)

        df_core_sorted = df_core.copy()

        # dataframes for each room
        df_robertson_summary = df_core_sorted[df_core_sorted["Room"].str.contains("Robertson")][COLUMN_ORDER_SUMMARY]
        df_robertson_detail = df_core_sorted[
            df_core_sorted["Room"].str.contains("Robertson")
        ][COLUMN_ORDER_DETAIL]

        df_fisher_summary = df_core_sorted[df_core_sorted["Room"].str.contains("Fisher")][COLUMN_ORDER_SUMMARY]
        df_fisher_detail = df_core_sorted[
            df_core_sorted["Room"].str.contains("Fisher")
        ][COLUMN_ORDER_DETAIL]

        df_workshop_summary = df_core_sorted[df_core_sorted["Room"].str.contains("Workshop")][COLUMN_ORDER_SUMMARY]
        df_workshop_detail = df_core_sorted[
            df_core_sorted["Room"].str.contains("Workshop")
        ][COLUMN_ORDER_DETAIL]

        return {
            "df_core_sorted": df_core_sorted,
            "robertson_summary": df_robertson_summary,
            "robertson_detail": df_robertson_detail,
            "fisher_summary": df_fisher_summary,
            "fisher_detail": df_fisher_detail,
            "workshop_summary": df_workshop_summary,
            "workshop_detail": df_workshop_detail,
        }
