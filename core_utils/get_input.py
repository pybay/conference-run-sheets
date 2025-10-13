"""
utility to get input from Sessionize - initially focused on local file download
TODO: consider using Sessionize API to save manual steps
"""

from pathlib import Path
import pandas as pd
from pandas import DataFrame


def get_input_from_sessionize(sessionize_input_path: Path) -> DataFrame:
    """
    Reads the first sheet of a validated input Excel file into a pandas DataFrame
    Expects all info to be on a single sheet
    :param sessionize_input_path:
    :return: pandas DataFrame
    """
    if sessionize_input_path is None:
        raise ValueError(f"Input path {sessionize_input_path} not provided")

    try:
        if sessionize_input_path and input_file_exists(sessionize_input_path):
            return pd.read_excel(sessionize_input_path,
                                 engine="calamine",
                                 header=0,
                                 keep_default_na=True,
                                 parse_dates=True).fillna("Not Provided")
    except FileNotFoundError as e:
        raise FileNotFoundError(f"File {sessionize_input_path} not found") from e
    except PermissionError as e:
        raise PermissionError(f"File {sessionize_input_path} not readable") from e


def input_file_exists(sessionize_path: Path) -> bool:
    """
    Validates that expected input Excel file exists 
    :param sessionize_path: 
    :return: bool 
    """
    
    return bool(Path(sessionize_path).is_file())
