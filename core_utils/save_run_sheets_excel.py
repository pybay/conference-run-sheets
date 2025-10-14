"""
Concrete implementation for Excel output
"""
from pathlib import Path

from pandas import DataFrame

from core_utils.save_run_sheet_manager import RunSheetSaveManager
from core_utils.pybay_standard_theme import PYBAY_PRIMARY_BLUE, PYBAY_SECONDARY_YELLOW, TAB_COLOR_PALETTE


class ExcelRunSheetWriter(RunSheetSaveManager):
    """
    Writes run sheets to Excel files using xlsxwriter.

    All sheets are formatted for printing on 8.5" x 11" portrait paper.
    """

    def __init__(self, results: dict[str, DataFrame], sessionize_output_path: Path):
        super().__init__(results, sessionize_output_path)
        self.sessionize_output_path = sessionize_output_path
        self.workbook = None
        self.formats = {}
        # Build room color mapping - assign colors sequentially to unique rooms
        unique_rooms = sorted(set(key.split('_')[0] for key in self.sheet_keys))
        self.room_colors = {room: TAB_COLOR_PALETTE[i % len(TAB_COLOR_PALETTE)]
                            for i, room in enumerate(unique_rooms)}

    def _setup(self):
        """Initialize Excel workbook and formats."""
        import xlsxwriter

        self.workbook = xlsxwriter.Workbook(str(self.sessionize_output_path))

        # Create formats - all standardized to top vertical alignment for consistency
        # No borders by default for cleaner print appearance
        self.formats['header'] = self.workbook.add_format({
            'bold': True,
            'bg_color': f'#{PYBAY_PRIMARY_BLUE}',
            'font_color': 'white',
            'align': 'center',
            'valign': 'top',
            'text_wrap': True
        })

        self.formats['cell_wrap'] = self.workbook.add_format({
            'text_wrap': True,
            'valign': 'top'
        })

        self.formats['cell_normal'] = self.workbook.add_format({
            'valign': 'top'
        })

        self.formats['time'] = self.workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'top'
        })

        self.formats['title'] = self.workbook.add_format({
            'bold': True,
            'font_size': 11,
            'valign': 'top',
            'text_wrap': True
        })

        self.formats['duration'] = self.workbook.add_format({
            'align': 'center',
            'valign': 'top'
        })

        self.formats['label'] = self.workbook.add_format({
            'bold': True,
            'align': 'right',
            'valign': 'top',
            'text_wrap': True
        })

        self.formats['url_visible'] = self.workbook.add_format({
            'valign': 'top',
            'font_color': 'blue',
            'underline': True
        })

        self.formats['url_visible_right'] = self.workbook.add_format({
            'valign': 'top',
            'align': 'right',
            'font_color': 'blue',
            'underline': True
        })

        self.formats['cell_bold'] = self.workbook.add_format({
            'bold': True,
            'valign': 'top'
        })

    def _write_sheet(self, df: DataFrame, sheet_name: str, sheet_type: str):
        """Write sheet based on type (summary or detail)."""
        if sheet_name:
            # Get tab color for this room from pre-built mapping
            room_name = sheet_name.split('_')[0]  # Extract room (fisher, robertson, workshop)
            tab_color = self.room_colors.get(room_name, TAB_COLOR_PALETTE[0])

            if sheet_type == 'summary':
                try:
                    worksheet = self.workbook.add_worksheet(sheet_name)
                    worksheet.set_tab_color(f'#{tab_color}')
                    self._write_summary_sheet(df, worksheet, sheet_name)
                except ValueError as e:
                    raise RuntimeError(f"Worksheet '{sheet_name}' not found.") from e
            elif sheet_type == 'detail':
                # Create mobile version (for phone viewing)
                try:
                    mobile_sheet_name = f"{sheet_name}_mobile"
                    worksheet_mobile = self.workbook.add_worksheet(mobile_sheet_name)
                    worksheet_mobile.set_tab_color(f'#{tab_color}')
                    self._write_detail_sheet_mobile(df, worksheet_mobile, mobile_sheet_name)
                except ValueError as e:
                    raise RuntimeError(f"Worksheet '{mobile_sheet_name}' creation failed.") from e

                # Create print version (for clipboard/printing)
                try:
                    print_sheet_name = f"{sheet_name}_print"
                    worksheet_print = self.workbook.add_worksheet(print_sheet_name)
                    worksheet_print.set_tab_color(f'#{tab_color}')
                    self._write_detail_sheet_print(df, worksheet_print, print_sheet_name)
                except ValueError as e:
                    raise RuntimeError(f"Worksheet '{print_sheet_name}' creation failed.") from e

    def _write_detail_sheet_mobile(self, df: DataFrame, worksheet, sheet_name: str):
        """
        Write detail sheet with vertical layout optimized for mobile viewing.

        Simple 2-column layout: Label | Data
        Designed to fit iPhone 15 Pro portrait mode with no horizontal scrolling.

        Column Width Calculation Rationale:
        - iPhone 15 Pro screen width: 393 CSS points (portrait mode)
        - Excel row number gutter: ~30-40 points
        - Available content width: 393 - 40 = 353 points
        - Excel column width units: 1 unit ≈ 7 points at default zoom
        - Maximum usable units: 353 / 7 ≈ 50 units

        Actual column allocation:
        - Column A (Labels): 18 units × 7 = 126 points
        - Column B (Data):   30 units × 7 = 210 points
        - Total content:     48 units × 7 = 336 points
        - With row numbers:  336 + 40 = 376 points (17pt buffer, fits comfortably)

        This ensures Room Captains can view all information on their phone
        without horizontal scrolling while maintaining readability.
        """
        import pandas as pd

        # Set default row height for all rows
        worksheet.set_default_row(15.0)

        # Mobile-optimized 2-column layout (fits iPhone 15 Pro portrait: 393pt screen)
        # Row numbers take ~30-40pt, leaving ~350pt for content
        # Excel units: 350pt / 7pt per unit ≈ 50 units max
        worksheet.set_column(0, 0, 18)  # Column A: Labels (18 units ≈ 126pt)
        worksheet.set_column(1, 1, 30)  # Column B: Data (30 units ≈ 210pt)
        # Total: 48 units ≈ 336pt (fits comfortably within 350pt available)

        # Create column index mapping from DataFrame
        col_idx = {col: idx for idx, col in enumerate(self.COLUMN_ORDER_DETAIL)}

        current_row = 0

        for record_idx, row_values in enumerate(df.values):
            # === Room NAME as header (only for first record - will be frozen/repeated) ===
            if record_idx == 0:
                room_name = row_values[col_idx['Room']] or 'Room'
                # Merge both columns, use primary blue background
                worksheet.merge_range(current_row, 0, current_row, 1, room_name, self.formats['header'])
                worksheet.set_row(current_row, 20)
                current_row += 1

            # === Time and Duration (in separate cells with yellow background) ===
            time_val = row_values[col_idx['Time']] or ''
            duration_val = row_values[col_idx['Duration']] or ''
            # Create yellow background formats matching print tabs
            time_yellow = self.workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'top',
                'bg_color': f'#{PYBAY_SECONDARY_YELLOW}'
            })
            duration_yellow = self.workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'top',
                'bg_color': f'#{PYBAY_SECONDARY_YELLOW}'
            })
            worksheet.write(current_row, 0, time_val, time_yellow)
            worksheet.write(current_row, 1, duration_val, duration_yellow)
            current_row += 1

            # === Title (merged both columns) ===
            worksheet.merge_range(current_row, 0, current_row, 1,
                                 row_values[col_idx['Title']] or '', self.formats['title'])
            worksheet.set_row(current_row, 30)
            current_row += 1

            # === Speaker name (merged both columns) ===
            worksheet.merge_range(current_row, 0, current_row, 1,
                                 row_values[col_idx['Speaker']] or '', self.formats['cell_bold'])
            current_row += 1

            # === Photo URL (clickable link, left-justified in column B) ===
            profile_photo = row_values[col_idx['Profile Photo']]
            worksheet.write(current_row, 0, '', self.formats['cell_normal'])
            if profile_photo and not pd.isna(profile_photo) and str(profile_photo) != 'Not Provided':
                worksheet.write_url(current_row, 1, str(profile_photo), self.formats['url_visible'], string='Photo_URL')
            else:
                worksheet.write(current_row, 1, '', self.formats['cell_normal'])
            current_row += 1

            # === Image insertion area (single tall row: 149pt = 46+46+57) ===
            from core_utils.image_helper import insert_image_to_worksheet

            if profile_photo and not pd.isna(profile_photo) and str(profile_photo) != 'Not Provided':
                # Insert image in column B (normalized to 144x144 = 1.5 inches)
                insert_image_to_worksheet(
                    worksheet, current_row, 1, str(profile_photo),
                    self.formats['cell_normal'],
                    target_size=(144, 144)
                )
                # Write empty cells for image area
                worksheet.write(current_row, 0, '', self.formats['cell_normal'])
                worksheet.write(current_row, 1, '', self.formats['cell_normal'])
            else:
                # No image, just write empty cells
                worksheet.write(current_row, 0, '', self.formats['cell_normal'])
                worksheet.write(current_row, 1, '', self.formats['cell_normal'])

            # Set single tall row for image (150pt to contain 144px image)
            worksheet.set_row(current_row, 150)
            current_row += 1

            # === pronunciation section (merged both columns) ===
            worksheet.merge_range(current_row, 0, current_row, 1, 'pronunciation', self.formats['cell_bold'])
            current_row += 1

            # First name
            worksheet.write(current_row, 0, 'First name:', self.formats['label'])
            worksheet.write(current_row, 1, row_values[col_idx['First name - pronunciation']] or '', self.formats['cell_normal'])
            current_row += 1

            # Last name
            worksheet.write(current_row, 0, 'Last name:', self.formats['label'])
            worksheet.write(current_row, 1, row_values[col_idx['Last name - pronunciation']] or '', self.formats['cell_normal'])
            current_row += 1

            # Blank row
            worksheet.write(current_row, 0, '', self.formats['cell_normal'])
            worksheet.write(current_row, 1, '', self.formats['cell_normal'])
            current_row += 1

            # Pronouns
            worksheet.write(current_row, 0, 'Pronouns:', self.formats['label'])
            worksheet.write(current_row, 1, row_values[col_idx['Pronouns']] or '', self.formats['cell_normal'])
            current_row += 1

            # First Conf Talk
            worksheet.write(current_row, 0, 'First Conf Talk:', self.formats['label'])
            worksheet.write(current_row, 1, row_values[col_idx['First Conf Talk']] or '', self.formats['cell_normal'])
            current_row += 1

            # Mobile # - PRIVATE!
            worksheet.write(current_row, 0, 'Mobile # PRIVATE!', self.formats['label'])
            worksheet.write(current_row, 1, row_values[col_idx['MOBILE # - PRIVATE!']] or '', self.formats['cell_normal'])
            current_row += 1

            # Blank row
            worksheet.write(current_row, 0, '', self.formats['cell_normal'])
            worksheet.write(current_row, 1, '', self.formats['cell_normal'])
            current_row += 1

            # === Attendees Learn (full width) ===
            worksheet.write(current_row, 0, 'Attendees Learn:', self.formats['cell_bold'])
            worksheet.write(current_row, 1, '', self.formats['cell_normal'])
            current_row += 1

            attendees_learn = row_values[col_idx['Attendees Learn']] or ''
            display_value = '' if pd.isna(attendees_learn) else str(attendees_learn)
            # Merge both columns for content
            worksheet.merge_range(current_row, 0, current_row, 1, display_value, self.formats['cell_wrap'])
            row_height = 15.0 if len(display_value) < 50 else 30 if len(display_value) < 100 else 50
            worksheet.set_row(current_row, row_height)
            current_row += 1

            # Blank row
            worksheet.write(current_row, 0, '', self.formats['cell_normal'])
            worksheet.write(current_row, 1, '', self.formats['cell_normal'])
            current_row += 1

            # === Speaker Bullets ===
            worksheet.write(current_row, 0, 'Speaker Bullets:', self.formats['cell_bold'])
            worksheet.write(current_row, 1, '', self.formats['cell_normal'])
            current_row += 1

            # Speaker intro bullets (merged for easier reading)
            for bullet_num in ['#1', '#2', '#3']:
                content = row_values[col_idx[f'Speaker intro {bullet_num}']] or ''
                display_value = '' if pd.isna(content) else str(content)
                worksheet.merge_range(current_row, 0, current_row, 1, display_value, self.formats['cell_wrap'])
                row_height = 15.0 if len(display_value) < 50 else 30 if len(display_value) < 100 else 50
                worksheet.set_row(current_row, row_height)
                current_row += 1

            # === Blank separator rows between records (more space for vertical layout) ===
            # First separator row
            worksheet.write(current_row, 0, '', self.formats['cell_normal'])
            worksheet.write(current_row, 1, '', self.formats['cell_normal'])
            worksheet.set_row(current_row, 20.0)
            current_row += 1

            # Second separator row for extra spacing
            worksheet.write(current_row, 0, '', self.formats['cell_normal'])
            worksheet.write(current_row, 1, '', self.formats['cell_normal'])
            worksheet.set_row(current_row, 20.0)
            current_row += 1

        # === Worksheet formatting (consistent order) ===
        # 1. Row sizing (first row with room name set to 20 in loop above)

        # 2. Screen display settings
        worksheet.freeze_panes(1, 0)  # Freeze first row (room name)

        # 3. Page setup - optimized for mobile/vertical scrolling
        worksheet.set_portrait()
        worksheet.set_paper(1)  # Letter size (8.5 x 11")
        worksheet.set_margins(left=0.25, right=0.25, top=0.75, bottom=0.75)

        # 4. Print layout
        worksheet.fit_to_pages(1, 0)  # Fit to 1 page wide, unlimited pages tall

        # 5. Repeating elements
        worksheet.repeat_rows(0)  # Repeat row 1 (room name) on every printed page
        worksheet.set_header(f'&C&BPyBay {self.conference_year}')  # Centered header with year
        worksheet.set_footer(f'&L&B{sheet_name}&R&Bpage &P of &N')

    def _write_summary_sheet(self, df: DataFrame, worksheet, sheet_name: str):
        """Write and format summary sheet."""
        # Set default row height for all rows
        worksheet.set_default_row(15.0)

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

        # === Worksheet formatting (consistent order) ===
        # 1. Row sizing
        worksheet.set_row(0, 30)

        # 2. Screen display settings
        worksheet.freeze_panes(1, 0)

        # 3. Page setup
        worksheet.set_portrait()
        worksheet.set_paper(1)  # Letter size (8.5 x 11")
        worksheet.set_margins(left=0.25, right=0.25, top=0.75, bottom=0.75)

        # 4. Print layout
        worksheet.fit_to_pages(1, 0)  # Fit to 1 page wide, unlimited pages tall

        # 5. Repeating elements
        worksheet.repeat_rows(0)  # Repeat row 1 (header) on every printed page
        worksheet.set_header(f'&C&BPyBay {self.conference_year}')  # Centered header with year
        worksheet.set_footer(f'&L&B{sheet_name}&R&Bpage &P of &N')

    def _write_detail_sheet_print(self, df: DataFrame, worksheet, sheet_name: str):
        """
        Write detail sheet with card-based layout optimized for portrait printing.

        Portrait layout: 8.5" wide x 11" tall
        With 0.25" side margins: 8.0" usable width
        Grid: 12 columns @ ~0.67" each = 8" total
        Each record uses ~11 rows in a card format
        """
        import pandas as pd

        # Set default row height for all rows
        worksheet.set_default_row(15.0)

        # Define fixed column grid for portrait (8.0" usable width)
        # Adjust column widths: labels wider (A-C), data starts at D
        worksheet.set_column(0, 2, 11)   # A-C: Label columns (slightly wider)
        worksheet.set_column(3, 9, 10)   # D-J: Data and content columns
        worksheet.set_column(10, 11, 10) # K-L: Photo columns

        # Define layout zones (column spans) - using 0-indexed columns
        ZONES = {
            'header_room': (0, 1),      # A-B: Room (2 cols)
            'header_time': (2, 3),      # C-D: Time (2 cols)
            'header_duration': (4, 4),  # E: Duration (1 col)
            'header_title': (5, 9),     # F-J: Title (5 cols wide)
            'header_speaker': (10, 11), # K-L: Speaker (2 cols)
            'label': (0, 2),            # A-C: Field labels (3 cols) - increased from 2
            'data_short': (3, 5),       # D-F: Short data (3 cols) - adjusted
            'data_wide': (3, 9),        # D-J: Wide content (7 cols) - adjusted
            'photo': (10, 11)           # K-L: Photo area (2 cols)
        }

        # Create column index mapping from DataFrame
        col_idx = {col: idx for idx, col in enumerate(self.COLUMN_ORDER_DETAIL)}

        current_row = 0

        for record_idx, row_values in enumerate(df.values):
            # === ROW 1: Header labels (blue background) ===
            # Only write headers for the first record
            if record_idx == 0:
                # Get the room name from the first record's data
                room_name = row_values[col_idx['Room']] or 'Room'

                worksheet.merge_range(
                    current_row, ZONES['header_room'][0],
                    current_row, ZONES['header_room'][1],
                    room_name, self.formats['header']
                )
                worksheet.merge_range(
                    current_row, ZONES['header_time'][0],
                    current_row, ZONES['header_time'][1],
                    'Time', self.formats['header']
                )
                worksheet.write(
                    current_row, ZONES['header_duration'][0],
                    'Duration', self.formats['header']
                )
                worksheet.merge_range(
                    current_row, ZONES['header_title'][0],
                    current_row, ZONES['header_title'][1],
                    'Title', self.formats['header']
                )
                worksheet.merge_range(
                    current_row, ZONES['header_speaker'][0],
                    current_row, ZONES['header_speaker'][1],
                    'Speaker', self.formats['header']
                )
                worksheet.set_row(current_row, 20)
                current_row += 1

            # === ROW 2 (or ROW 1 for subsequent records): Data row with main session info ===
            # Create formats with yellow background for this row
            # Room (bold, yellow bg)
            room_format = self.workbook.add_format({
                'bold': True,
                'valign': 'top',
                'bg_color': f'#{PYBAY_SECONDARY_YELLOW}'
            })
            worksheet.merge_range(
                current_row, ZONES['header_room'][0],
                current_row, ZONES['header_room'][1],
                row_values[col_idx['Room']] or '', room_format
            )
            # Time (bold, centered, yellow bg)
            time_format = self.workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'top',
                'bg_color': f'#{PYBAY_SECONDARY_YELLOW}'
            })
            worksheet.merge_range(
                current_row, ZONES['header_time'][0],
                current_row, ZONES['header_time'][1],
                row_values[col_idx['Time']] or '', time_format
            )
            # Duration (bold, centered, yellow bg)
            duration_format = self.workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'top',
                'bg_color': f'#{PYBAY_SECONDARY_YELLOW}'
            })
            worksheet.write(
                current_row, ZONES['header_duration'][0],
                row_values[col_idx['Duration']] or '', duration_format
            )
            # Title (bold, wide, yellow bg, WRAP TEXT)
            title_format = self.workbook.add_format({
                'bold': True,
                'font_size': 11,
                'valign': 'top',
                'bg_color': f'#{PYBAY_SECONDARY_YELLOW}',
                'text_wrap': True
            })
            worksheet.merge_range(
                current_row, ZONES['header_title'][0],
                current_row, ZONES['header_title'][1],
                row_values[col_idx['Title']] or '', title_format
            )
            # Speaker (bold, yellow bg, wrap text)
            speaker_format = self.workbook.add_format({
                'bold': True,
                'valign': 'top',
                'bg_color': f'#{PYBAY_SECONDARY_YELLOW}',
                'text_wrap': True
            })
            worksheet.merge_range(
                current_row, ZONES['header_speaker'][0],
                current_row, ZONES['header_speaker'][1],
                row_values[col_idx['Speaker']] or '', speaker_format
            )
            worksheet.set_row(current_row, 30)  # Taller data row
            current_row += 1

            # === Add "pronunciation:" label row ===
            # Merge first three columns (A-C) for the pronunciation label
            worksheet.merge_range(
                current_row, 0,
                current_row, 2,
                'pronunciation:', self.formats['label']
            )
            # Merge data area (D-I) for pronunciation section
            worksheet.merge_range(
                current_row, 3,
                current_row, 9,
                '', self.formats['cell_normal']
            )
            # Columns K-L merged - Photo URL link (left-justified)
            profile_photo = row_values[col_idx['Profile Photo']]
            if profile_photo and not pd.isna(profile_photo) and str(profile_photo) != 'Not Provided':
                worksheet.merge_range(
                    current_row, 10,  # Column K
                    current_row, 11,  # Column L
                    '', self.formats['cell_normal']
                )
                worksheet.write_url(
                    current_row, 10,  # Write to merged K-L, starting at K
                    str(profile_photo),
                    self.formats['url_visible'],  # Left-aligned by default
                    string='Photo URL'
                )
            else:
                worksheet.merge_range(current_row, 10, current_row, 11, '', self.formats['cell_normal'])
            worksheet.set_row(current_row, 15.0)
            current_row += 1

            # === Detail fields with labels - extended data area ===
            detail_fields = [
                ('First name:', 'First name - pronunciation'),
                ('Last name:', 'Last name - pronunciation'),
                # Blank row after Last name
                None,
                ('Pronouns:', 'Pronouns'),
                ('First Conf Talk:', 'First Conf Talk'),
                ('MOBILE # - PRIVATE!', 'MOBILE # - PRIVATE!')
            ]

            photo_start_row = current_row  # Remember where photo area starts

            # FIRST: Merge the entire photo area (K-L) for all 6 detail field rows
            # Calculate end row (6 fields including 1 blank = 6 rows)
            photo_end_row = current_row + 5  # 0-indexed: current + 5 = 6 rows total
            worksheet.merge_range(
                photo_start_row, ZONES['photo'][0],
                photo_end_row, ZONES['photo'][1],
                '', self.formats['cell_normal']
            )

            # SECOND: Insert image into the merged area (top-right corner)
            profile_photo_url = row_values[col_idx['Profile Photo']]
            if profile_photo_url and not pd.isna(profile_photo_url) and str(profile_photo_url) != 'Not Provided':
                from core_utils.image_helper import insert_image_to_worksheet
                insert_image_to_worksheet(
                    worksheet,
                    photo_start_row,
                    ZONES['photo'][0],
                    str(profile_photo_url),
                    self.formats['cell_normal'],
                    target_size=(144, 144)  # Normalized to 144x144 = 1.5 inches @ 96 DPI
                )

            # THIRD: Write detail fields (labels and data only - photo area already handled)
            for field in detail_fields:
                if field is None:
                    # Insert blank row - only label and data areas (photo already merged above)
                    # Label area (A-C)
                    worksheet.merge_range(
                        current_row, 0,
                        current_row, 2,
                        '', self.formats['cell_normal']
                    )
                    # Data area (D-I)
                    worksheet.merge_range(
                        current_row, 3,
                        current_row, 9,
                        '', self.formats['cell_normal']
                    )
                    # Photo area (K-L) already merged, don't write
                    worksheet.set_row(current_row, 22.0)
                    current_row += 1
                else:
                    label_text, data_field = field
                    # Label (right-aligned, bold) - columns A-C
                    worksheet.merge_range(
                        current_row, ZONES['label'][0],
                        current_row, ZONES['label'][1],
                        label_text, self.formats['label']
                    )
                    # Data - extended to columns D-I
                    worksheet.merge_range(
                        current_row, 3,
                        current_row, 9,
                        row_values[col_idx[data_field]] or '', self.formats['cell_normal']
                    )
                    # Photo area (K-L) already merged, don't write
                    worksheet.set_row(current_row, 22.0)
                    current_row += 1

            # === Add blank separator row between detail fields and long text sections ===
            # Use same merge pattern as data rows
            # Label area (A-C)
            worksheet.merge_range(
                current_row, 0,
                current_row, 2,
                '', self.formats['cell_normal']
            )
            # Data area extends to full width (D-L)
            worksheet.merge_range(
                current_row, 3,
                current_row, 11,
                '', self.formats['cell_normal']
            )
            worksheet.set_row(current_row, 15.0)
            current_row += 1

            # === ROWS 7-10: Long text fields with labels ===
            long_text_fields = [
                ('Attendees Learn:', 'Attendees Learn'),
                # Blank row with "Speaker Bullets:" label
                'SPEAKER_BULLETS_HEADER',
                ('#1:', 'Speaker intro #1'),
                ('#2:', 'Speaker intro #2'),
                ('#3:', 'Speaker intro #3')
            ]

            for field in long_text_fields:
                if field == 'SPEAKER_BULLETS_HEADER':
                    # Insert blank row with "Speaker Bullets:" label
                    # Label area (A-C) with section header
                    worksheet.merge_range(
                        current_row, 0,
                        current_row, 2,
                        'Speaker Bullets:', self.formats['label']
                    )
                    # Data area extends to full width (D-L)
                    worksheet.merge_range(
                        current_row, 3,
                        current_row, 11,
                        '', self.formats['cell_normal']
                    )
                    worksheet.set_row(current_row, 15.0)
                    current_row += 1
                else:
                    label_text, data_field = field
                    # Label (left side)
                    worksheet.merge_range(
                        current_row, ZONES['label'][0],
                        current_row, ZONES['label'][1],
                        label_text, self.formats['label']
                    )
                    # Wide content area spanning all the way to column L (rightmost)
                    content = row_values[col_idx[data_field]] or ''
                    display_value = '' if pd.isna(content) else str(content)
                    worksheet.merge_range(
                        current_row, ZONES['data_wide'][0],
                        current_row, 11,  # Column L is index 11 (rightmost column)
                        display_value, self.formats['cell_wrap']
                    )

                    # Set row height based on content length - minimum 15.0
                    row_height = 15.0 if len(display_value) < 50 else 30 if len(display_value) < 100 else 50
                    worksheet.set_row(current_row, row_height)
                    current_row += 1

            # === Add blank separator row between records ===
            for col in range(12):
                worksheet.write(current_row, col, '', self.formats['cell_normal'])
            worksheet.set_row(current_row, 15.0)  # Standard separator height
            current_row += 1

        # === Worksheet formatting (consistent order) ===
        # 1. Row sizing (header row set to 20 in loop above)

        # 2. Screen display settings
        worksheet.freeze_panes(1, 0)

        # 3. Page setup
        worksheet.set_portrait()
        worksheet.set_paper(1)  # Letter size (8.5 x 11")
        worksheet.set_margins(left=0.25, right=0.25, top=0.75, bottom=0.75)

        # 4. Print layout
        worksheet.fit_to_pages(1, 0)  # Fit to 1 page wide, unlimited pages tall

        # 5. Repeating elements
        worksheet.repeat_rows(0)  # Repeat row 1 (blue header: Room name|Time|Duration|Title|Speaker)
        worksheet.set_header(f'&C&BPyBay {self.conference_year}')  # Centered header with year
        worksheet.set_footer(f'&L&B{sheet_name}&R&Bpage &P of &N')


    def _finalize(self):
        """Close workbook."""
        self.workbook.close()
        print(f"Excel workbook saved to: {self.sessionize_output_path}")
