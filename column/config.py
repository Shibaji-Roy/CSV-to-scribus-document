#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
config.py - Configuration file for the PDF generation script
All parameters are centralized here for easier management and customization.
"""

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── FILE PATHS SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Default input and output file paths
# Set to None to show file dialog for input or use derived filename for output
DEFAULT_INPUT_FILE = None  # Set to a specific path like "/path/to/input.json" to skip dialog
DEFAULT_OUTPUT_FILE = None  # Set to None to use the input filename with .pdf extension
                           # Or specify a path like "/path/to/output.pdf"

# Default path for images folder
# Can be absolute (e.g., "D:/images") or relative to JSON file (e.g., "Pictures")
DEFAULT_IMAGES_PATH = "Pictures"
 
# Whether to always show file dialog for input regardless of DEFAULT_INPUT_FILE setting
ALWAYS_SHOW_INPUT_DIALOG = True

# Whether to always show file dialog for output regardless of DEFAULT_OUTPUT_FILE setting
ALWAYS_SHOW_OUTPUT_DIALOG = False

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── PAGE LAYOUT SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Page dimensions and margins
PAGE_WIDTH, PAGE_HEIGHT = 480, 595  # Wider than A5 size in points
MARGINS = (28, 28, 32, 28)   # left, top, right, bottom (in points) - shifted 2px left

# Two-column layout settings
USE_TWO_COLUMN_LAYOUT = True  # Enable/disable two-column layout
COLUMN_COUNT = 2  # Number of columns
COLUMN_GAP = 20  # Gap between columns (in points) - increased to prevent overlap
BALANCE_COLUMNS = True  # Balance text height across columns for equal fill
AGGRESSIVE_BALANCING = True  # Force aggressive column balancing to minimize wasted space

# Calculate column width based on full page width (like in the example image)
COLUMN_WIDTH = (PAGE_WIDTH - MARGINS[0] - MARGINS[2] - ((COLUMN_COUNT - 1) * COLUMN_GAP)) / COLUMN_COUNT

# Additional layout parameters
PAGE_LENGTH = PAGE_WIDTH  # Total width of the page (same as PAGE_WIDTH, now 480)
PAGE_WIDTH_VERT = PAGE_HEIGHT  # Total height of the page (same as PAGE_HEIGHT)
TAB_WIDTH = 50  # Width of the tab column
FIELD3_WIDTH = 150  # Width of Field 3 column
SIGN_WIDTH = 80  # Width of the sign/image area

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── SPACING SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Spacing constants
BLOCK_SPACING = 10          # Space between content blocks to prevent overlap (in points)
MIN_SPACE_THRESHOLD = 4     # Minimum space required before forcing a new page
TEMPLATE_PADDING = 0        # No padding between templates
MIN_BOTTOM_SPACE = 5        # Minimal space at bottom of page

# Hierarchical spacing system based on JSON structure
HEADER_TO_DESC_SPACING = 2      # Very tight spacing: header → its description
MODULE_TO_TEMPLATE_SPACING = 1  # Minimal spacing: module name → template content
TEMPLATE_TO_TEMPLATE_SPACING = 3 # Small spacing: template → next template
SECTION_TO_SECTION_SPACING = 8  # Medium spacing: topic → next topic, chapter → chapter

# Text padding (left, right, top, bottom) in points
TEMPLATE_TEXT_PADDING = (0, 0, 0, 0)    # No padding for maximum space
REGULAR_TEXT_PADDING = (0, 0, 0, 0)     # No padding for maximum space

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── FONT SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Default font to use when no specific font is mentioned
DEFAULT_FONT = "Arial"

# Font options - tried in order until one is found
FONT_CANDIDATES = [
    "Myriad Pro Cond","Arial", "Verdana"
]

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── HEADER SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Module header settings
MODULE_FONT_SIZE = 11       # Further reduced font size for module headers
MODULE_BOLD = True          # Whether module headers should be bold
MODULE_BG_COLOR = "None"    # Background color for module headers (None = transparent)
MODULE_TEXT_COLOR = "Black" # Text color for module headers
MODULE_PADDING = 0          # No padding around module headers

# Topic header settings
TOPIC_FONT_SIZE = 12        # Increased font size for topic headers to maintain hierarchy
TOPIC_BOLD = True           # Whether topic headers should be bold  
TOPIC_BG_COLOR = "None"     # Background color for topic headers (None = transparent)
TOPIC_TEXT_COLOR = "Black"  # Text color for topic headers
TOPIC_PADDING = 0           # No padding around topic headers

# Description text settings
AREA_DESC_FONT_SIZE = 7     # Font size for area descriptions - compact content text
TOPIC_DESC_FONT_SIZE = 7    # Font size for topic descriptions - compact content text
MODULE_TEXT_FONT_SIZE = 7   # Font size for module/template text content

# Text container sizing settings to prevent overflow
TEXT_HEIGHT_BUFFER_FACTOR = 0.5   # Much smaller buffer for very tight containers
OVERFLOW_INCREMENT_SIZE = 5        # Very small increments when resizing for overflow
MAX_OVERFLOW_ITERATIONS = 10       # Allow more iterations to fit full text
TEXT_HEIGHT_REDUCTION = 0.95      # Reduce measured height by 5% for tighter fit

# Header font settings
AREA_HEADER_FONT_SIZE = 12  # Font size for area headers
VIDEO_TEXT_FONT_SIZE = 6    # Font size for video text in templates - compact content text
SUBTITLE_FONT_SIZE = 11     # Font size for subtitles in two-column layouts

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── TEMPLATE SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Template appearance and limits
GLOBAL_TEMPLATE_LIMIT = 60 # Maximum number of templates to process
BACKGROUND_COLORS = ["Red", "Green", "Yellow", "Blue", "Cyan", "Magenta"]

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── QUIZ SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Quiz display settings
QUIZ_SHOW_QUESTIONS = True      # Show questions in quiz sections
QUIZ_SHOW_ANSWERS = True        # Show answers in quiz sections
QUIZ_FILTER_MODE = "all"        # Filter mode: "all", "true_only", "false_only"

# Quiz font settings
QUIZ_FONT_FAMILY = "Myriad Pro Cond"  # Primary font for quiz elements

# Quiz heading settings
QUIZ_HEADING_HEIGHT = 20        # Height of the quiz heading
QUIZ_HEADING_TEXT = "QUIZ"      # Text to display in the quiz heading
QUIZ_HEADING_FONT_SIZE = 14     # Font size for quiz heading

# Quiz layout settings
QUIZ_LENGTH = PAGE_WIDTH - MARGINS[0] - MARGINS[2] - 20  # Width of quiz section (adjusted for wider page)
QUIZ_CARD_SPACING = 0           # No spacing between question cards for compact layout
QUIZ_QUESTION_PADDING = 1       # Minimal internal padding for question text
QUIZ_HORIZONTAL_SPACING = 4     # Reduced space between question text and answer box
MIN_QUIZ_SPACING = 1            # Reduced space between quiz items

# Answer box settings
QUIZ_ANSWER_BOX_WIDTH = 60      # Width of the answer box
QUIZ_ANSWER_BOX_HEIGHT = 40     # Height of the answer box
MIN_ANSWER_WIDTH = 150          # Minimum width for answer column

# Font sizes
QUIZ_QUESTION_FONT_SIZE = 9     # Quiz questions at readable size
QUIZ_ANSWER_FONT_SIZE = 14      # Font size for answers (V/F boxes)

# Answer indicators
QUIZ_TRUE_TEXT = "V"            # Text for true answers
QUIZ_FALSE_TEXT = "F"           # Text for false answers
QUIZ_TRUE_BORDER_COLOR = "Green"  # Border color for true answers
QUIZ_FALSE_BORDER_COLOR = "Red"   # Border color for false answers

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── VERTICAL BANNER SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Vertical topic banner settings
TOPIC_BANNER_WIDTH = 20         # Width of the vertical topic banner (in points)
TOPIC_BANNER_COLOR = "DarkGrey" # Default color of the vertical topic banner

# Banner margin offset settings  
LEFT_BANNER_MARGIN_OFFSET = 2   # Gap between left banner and left margin (in points)
RIGHT_BANNER_MARGIN_OFFSET = 6   # Gap between right banner and right margin (in points)

# Left banner text settings (odd pages)
LEFT_BANNER_HORIZONTAL_OFFSET = 4    # Distance from left edge of banner (in points)
LEFT_BANNER_VERTICAL_OFFSET = 400    # Vertical offset from center position (in points)
LEFT_BANNER_ROTATION = 90            # Rotation for left banner text (degrees)

# Right banner text settings (even pages)
RIGHT_BANNER_HORIZONTAL_OFFSET = 4   # Distance from right edge of banner (in points)
RIGHT_BANNER_VERTICAL_OFFSET = 0     # Vertical offset from center position (in points)
RIGHT_BANNER_ROTATION = 270          # Rotation for right banner text (degrees)

# Banner text appearance
BANNER_TEXT_FONT_SIZE = 10         # Reduced font size for banner text
BANNER_TEXT_COLOR = "White"        # Text color for banner text
BANNER_TEXT_HEIGHT_PERCENT = 0.75  # Percentage of banner height to use for text height 

# Font aliases - alternative names for the same fonts
FONT_ALIASES = {
    "Myriad Pro Condensed": "Myriad Pro Cond",  # Condensed version is known as Cond
    "SansSerifCollection": "sans-serif",        # Alias for sans-serif
    "Myriad Pro": "Myriad Pro Cond"                 # Simplified name
} 
