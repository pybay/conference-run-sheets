"""
Concrete implementation for Excel output
"""
from pathlib import Path

from pandas import DataFrame

from core_utils.save_run_sheet_manager import RunSheetSaveManager
from core_utils.pybay_standard_theme import PYBAY_PRIMARY_BLUE


class ExcelRunSheetWriter(RunSheetSaveManager):
    """Writes run sheets to Excel files using xlsxwriter."""

    def __init__(self, results: dict[str, DataFrame], sessionize_output_path: Path):
        super().__init__(results, sessionize_output_path)
        self.sessionize_output_path = sessionize_output_path
        self.workbook = None
        self.formats = {}

    def _setup(self):
        """Initialize Excel workbook and formats."""
        import xlsxwriter

        self.workbook = xlsxwriter.Workbook(str(self.sessionize_output_path))

        # Create formats - all standardized to top vertical alignment for consistency
        self.formats['header'] = self.workbook.add_format({
            'bold': True,
            'bg_color': f'#{PYBAY_PRIMARY_BLUE}',
            'font_color': 'white',
            'align': 'center',
            'valign': 'top',
            'border': 1,
            'text_wrap': True
        })

        self.formats['cell_wrap'] = self.workbook.add_format({
            'text_wrap': True,
            'valign': 'top',
            'border': 1
        })

        self.formats['cell_normal'] = self.workbook.add_format({
            'valign': 'top',
            'border': 1
        })

        self.formats['time'] = self.workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'top',
            'border': 1
        })

        self.formats['title'] = self.workbook.add_format({
            'bold': True,
            'font_size': 11,
            'valign': 'top',
            'border': 1
        })

        self.formats['duration'] = self.workbook.add_format({
            'align': 'center',
            'valign': 'top',
            'border': 1
        })

    def _write_sheet(self, df: DataFrame, sheet_name: str, sheet_type: str):
        """Write sheet based on type (summary or detail)."""
        if sheet_name:
            try:
                worksheet = self.workbook.add_worksheet(sheet_name)
            except ValueError as e:
                raise RuntimeError(f"Worksheet '{sheet_name}' not found.") from e

            if sheet_type == 'summary':
                self._write_summary_sheet(df, worksheet)
            elif sheet_type == 'detail':
                self._write_detail_sheet(df, worksheet)

    def _write_summary_sheet(self, df: DataFrame, worksheet):
        """Write and format summary sheet."""
        # Set column widths
        worksheet.set_column('A:A', 15)  # Room
        worksheet.set_column('B:B', 12)  # Time
        worksheet.set_column('C:C', 60)  # Title
        worksheet.set_column('D:D', 30)  # Speaker

        # Write headers
        for col_num, column_name in enumerate(self.COLUMN_ORDER_SUMMARY):
            worksheet.write(0, col_num, column_name, self.formats['header'])

        # Create column index mapping
        col_idx = {col: idx for idx, col in enumerate(self.COLUMN_ORDER_SUMMARY)}

        # Write data
        for row_num, row_values in enumerate(df.values, start=1):
            worksheet.write(row_num, col_idx['Room'], row_values[col_idx['Room']], self.formats['cell_normal'])
            worksheet.write(row_num, col_idx['Time'], row_values[col_idx['Time']], self.formats['time'])
            worksheet.write(row_num, col_idx['Title'], row_values[col_idx['Title']], self.formats['title'])
            worksheet.write(row_num, col_idx['Speaker'], row_values[col_idx['Speaker']], self.formats['cell_normal'])

        worksheet.freeze_panes(1, 0)
        worksheet.set_row(0, 30)

    def _write_detail_sheet(self, df: DataFrame, worksheet):
        """Write and format detail sheet with images."""
        import pandas as pd

        # Column widths
        column_widths = {
            'Room': 15,
            'Time': 12,
            'Duration': 12,
            'Title': 50,
            'Speaker': 25,
            'First name - pronunciation': 20,
            'Last name - pronunciation': 20,
            'Mobile # (NOT PUBLIC)': 20,
            'Pronouns': 12,
            'First Conf Talk': 15,
            'Profile Photo': 15,
            'Attendees Learn': 50,
            'Speaker intro #1': 50,
            'Speaker intro #2': 50,
            'Speaker intro #3': 50
        }

        # Set column widths
        for col_num, column_name in enumerate(self.COLUMN_ORDER_DETAIL):
            width = column_widths.get(column_name, 15)
            worksheet.set_column(col_num, col_num, width)

        # Write headers
        for col_num, column_name in enumerate(self.COLUMN_ORDER_DETAIL):
            worksheet.write(0, col_num, column_name, self.formats['header'])

        worksheet.set_row(0, 40)

        # Create column index mapping
        col_idx = {col: idx for idx, col in enumerate(self.COLUMN_ORDER_DETAIL)}

        # Write data rows
        for row_num, row_values in enumerate(df.values, start=1):
            row_height = 60

            # Handle Profile Photo with image insertion
            # TODO: add check cache for image in images_cache/ dir, add concurrent parallel retrieval,
            #       and add image size check before insertion
            profile_photo = row_values[col_idx['Profile Photo']]
            if profile_photo and not pd.isna(profile_photo):
                # Temporarily writing URL as text until parallel image download is implemented
                worksheet.write(row_num, col_idx['Profile Photo'], profile_photo, self.formats['cell_normal'])
            else:
                worksheet.write(row_num, col_idx['Profile Photo'], '', self.formats['cell_normal'])

            # Long text fields with wrapping
            for col_name in ['Attendees Learn', 'Speaker intro #1', 'Speaker intro #2', 'Speaker intro #3']:
                value = row_values[col_idx[col_name]]
                display_value = '' if pd.isna(value) else str(value)
                worksheet.write(row_num, col_idx[col_name], display_value, self.formats['cell_wrap'])
                if display_value and len(display_value) > 100:
                    row_height = max(row_height, 80)

            # Time column with special formatting
            time_value = row_values[col_idx['Time']]
            worksheet.write(row_num, col_idx['Time'],
                            '' if pd.isna(time_value) else time_value,
                            self.formats['time'])

            # Title column with special formatting
            title_value = row_values[col_idx['Title']]
            worksheet.write(row_num, col_idx['Title'],
                            '' if pd.isna(title_value) else title_value,
                            self.formats['title'])

            # Duration column with center alignment
            duration_value = row_values[col_idx['Duration']]
            worksheet.write(row_num, col_idx['Duration'],
                           '' if pd.isna(duration_value) else duration_value,
                           self.formats['duration'])

            # All other columns with normal formatting
            for col_name in ['Room', 'Speaker', 'First name - pronunciation',
                             'Last name - pronunciation', 'Mobile # (NOT PUBLIC)',
                             'Pronouns', 'First Conf Talk']:
                value = row_values[col_idx[col_name]]
                worksheet.write(row_num, col_idx[col_name],
                                '' if pd.isna(value) else value,
                                self.formats['cell_normal'])

            worksheet.set_row(row_num, row_height)

        worksheet.freeze_panes(1, 2)

    def _insert_image_excel(self, worksheet, row: int, col: int, url: str):
        """Download and insert image into Excel cell."""
        import requests
        from io import BytesIO

        try:
            print(f"Downloading image for row {row}...")
            response = requests.get(url, timeout=5)  # Reduced timeout from 10 to 5
            response.raise_for_status()
            image_data = BytesIO(response.content)

            worksheet.insert_image(
                row, col, url,
                {
                    'image_data': image_data,
                    'x_scale': 0.5,
                    'y_scale': 0.5,
                    'x_offset': 5,
                    'y_offset': 5,
                    'positioning': 1
                }
            )
            print(f"✓ Image inserted for row {row}")
        except requests.exceptions.Timeout:
            print(f"⚠ Timeout downloading image for row {row}, skipping...")
            worksheet.write(row, col, "Image timeout", self.formats['cell_normal'])
        except Exception as e:
            print(f"⚠ Failed to insert image for row {row}: {e}")
            worksheet.write(row, col, url, self.formats['cell_normal'])

    def _finalize(self):
        """Close workbook."""
        self.workbook.close()
        print(f"Excel workbook saved to: {self.sessionize_output_path}")
