#! 
DEBUG = False
WARNINGS = True

#
################ [ SECTION OF CONSTANTS ] ##################################
#

TRY_OPTIMIZE = False  # If True - tries to use optimized modes in some places, if possible
# MAX_COLUMN_WIDTH - width in pixels and characters. Because xlsxwriter autofit() uses pixels, and we use the number
# of characters in compute_column_widths().
MAX_COLUMN_WIDTH = {'pixels': 500, 'chars': 70}
# Extra character reserve added to computed column widths to reduce accidental wrapping in Excel.
COLUMN_WIDTH_RESERVE_CHARS = 2
