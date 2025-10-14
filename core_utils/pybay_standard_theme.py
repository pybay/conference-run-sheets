"""
PyBay 2025 Brand Theme Configuration

Contains official PyBay color palette and typography constants for consistent
formatting across all generated run sheets and conference materials.
"""

# PyBay Primary Colors
PYBAY_PRIMARY_BLUE = "2E648E"
PYBAY_PRIMARY_YELLOW = "FDC13C"
PYBAY_BLACK = "000000"

# PyBay Secondary Colors
PYBAY_SECONDARY_BLUE = "D9E3EA"
PYBAY_SECONDARY_YELLOW = "FCD582"

# Typography
PYBAY_DEFAULT_FONT = "Urbanist"
PYBAY_DEFAULT_FONT_URL = "https://fonts.google.com/specimen/Urbanist?preview.text=PyBay%202025"

# Color hex values must be without the '#' prefix for xlsxwriter compatibility
# Example usage in xlsxwriter:
#   format = workbook.add_format({'bg_color': f'#{PYBAY_PRIMARY_BLUE}'})

# Tab colors - rotate through PyBay brand colors
# Order: Primary colors first, then secondary colors
TAB_COLOR_PALETTE = [
    PYBAY_PRIMARY_BLUE,      # 2E648E - Deep blue (1st room)
    PYBAY_PRIMARY_YELLOW,    # FDC13C - Golden yellow (2nd room)
    PYBAY_SECONDARY_BLUE,    # D9E3EA - Very light blue (3rd room)
    PYBAY_SECONDARY_YELLOW,  # FCD582 - Light yellow (4th room, if needed)
]
