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

# Default image size for integrated text-image layout
DEFAULT_IMAGE_SIZE = 80  # Configurable default image size (can be increased without text overflow)
 
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
AGGRESSIVE_BALANCING = False  # Disable aggressive column balancing to prevent overlap

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
BLOCK_SPACING = 3           # Space between content blocks (in points) - reduced for tighter layout
MIN_SPACE_THRESHOLD = 4     # Minimum space required before forcing a new page
TEMPLATE_PADDING = 0        # No padding between templates
MIN_BOTTOM_SPACE = 5        # Minimal space at bottom of page

# Hierarchical spacing system based on JSON structure
HEADER_TO_DESC_SPACING = 5      # Spacing: header → its description
MODULE_TO_TEMPLATE_SPACING = 2  # Spacing: module name → template content
TEMPLATE_TO_TEMPLATE_SPACING = 12 # Spacing: template → next template (reduced from 36)

# Hierarchy-based spacing (larger gaps for higher-level elements)
AREA_TO_AREA_SPACING = 10       # Large spacing: area → next area (reduced from 12)
AREA_TO_CHAPTER_SPACING = 6     # Medium spacing: area → first chapter (reduced from 8)
CHAPTER_TO_CHAPTER_SPACING = 6  # Medium spacing: chapter → next chapter (reduced from 8)
TOPIC_TO_TOPIC_SPACING = 2      # Smaller spacing: topic → next topic (reduced from 5)
TOPIC_TO_MODULE_SPACING = 3     # Minimal spacing: topic → first module (reduced from 4)

# Legacy constant for backward compatibility
SECTION_TO_SECTION_SPACING = 2  # Default section spacing (reduced from 5)

# Text padding (left, right, top, bottom) in points
TEMPLATE_TEXT_PADDING = (0, 0, 0, 0)    # No padding for maximum space
REGULAR_TEXT_PADDING = (0, 0, 0, 0)     # No padding for maximum space

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── FONT & TYPOGRAPHY SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Default font to use when no specific font is mentioned
DEFAULT_FONT = "Arial"

# Font options - tried in order until one is found
FONT_CANDIDATES = [
    "Myriad Pro Cond", "Arial", "Verdana"
]

# Font aliases - alternative names for the same fonts
FONT_ALIASES = {
    "Myriad Pro Condensed": "Myriad Pro Cond",  # Condensed version is known as Cond
    "SansSerifCollection": "sans-serif",        # Alias for sans-serif
    "Myriad Pro": "Myriad Pro Cond"             # Simplified name
}

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── AREA HEADER SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Area header text styling
AREA_HEADER_FONT_FAMILY = "Myriad Pro Cond"  # Font family for area headers
AREA_HEADER_FONT_SIZE = 12                    # Font size for area headers
AREA_HEADER_BOLD = True                       # Whether area headers should be bold
AREA_HEADER_ITALIC = False                    # Whether area headers should be italic
AREA_HEADER_TEXT_COLOR = "Black"              # Text color for area headers
AREA_HEADER_BG_COLOR = "None"                 # Background color (None = transparent)

# Area description text styling
AREA_DESC_FONT_FAMILY = "Myriad Pro Cond"    # Font family for area descriptions
AREA_DESC_FONT_SIZE = 9                       # Font size for area descriptions - compact text
AREA_DESC_BOLD = False                        # Whether area descriptions should be bold
AREA_DESC_ITALIC = False                      # Whether area descriptions should be italic
AREA_DESC_TEXT_COLOR = "Black"                # Text color for area descriptions

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── CHAPTER HEADER SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Chapter header text styling
CHAPTER_HEADER_FONT_FAMILY = "Myriad Pro Cond"  # Font family for chapter headers
CHAPTER_HEADER_FONT_SIZE = 11                    # Font size for chapter headers
CHAPTER_HEADER_BOLD = True                       # Whether chapter headers should be bold
CHAPTER_HEADER_ITALIC = False                    # Whether chapter headers should be italic
CHAPTER_HEADER_TEXT_COLOR = "Black"              # Text color for chapter headers
CHAPTER_HEADER_BG_COLOR = "None"                 # Background color (None = transparent)

# Chapter description text styling
CHAPTER_DESC_FONT_FAMILY = "Myriad Pro Cond"    # Font family for chapter descriptions
CHAPTER_DESC_FONT_SIZE = 9                       # Font size for chapter descriptions
CHAPTER_DESC_BOLD = False                        # Whether chapter descriptions should be bold
CHAPTER_DESC_ITALIC = False                      # Whether chapter descriptions should be italic
CHAPTER_DESC_TEXT_COLOR = "Black"                # Text color for chapter descriptions

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── TOPIC HEADER SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Topic header text styling
TOPIC_HEADER_FONT_FAMILY = "Myriad Pro Cond"    # Font family for topic headers
TOPIC_FONT_SIZE = 12                             # Font size for topic headers (maintains hierarchy)
TOPIC_BOLD = True                                # Whether topic headers should be bold
TOPIC_ITALIC = False                             # Whether topic headers should be italic
TOPIC_BG_COLOR = "None"                          # Background color (None = transparent)
TOPIC_TEXT_COLOR = "Black"                       # Text color for topic headers
TOPIC_PADDING = 0                                # No padding around topic headers

# Topic description text styling
TOPIC_DESC_FONT_FAMILY = "Myriad Pro Cond"      # Font family for topic descriptions
TOPIC_DESC_FONT_SIZE = 9                         # Font size for topic descriptions - compact text
TOPIC_DESC_BOLD = False                          # Whether topic descriptions should be bold
TOPIC_DESC_ITALIC = False                        # Whether topic descriptions should be italic
TOPIC_DESC_TEXT_COLOR = "Black"                  # Text color for topic descriptions

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── MODULE HEADER SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Module header text styling
MODULE_HEADER_FONT_FAMILY = "Myriad Pro Cond"   # Font family for module headers
MODULE_FONT_SIZE = 11                            # Font size for module headers
MODULE_BOLD = True                               # Whether module headers should be bold
MODULE_ITALIC = False                            # Whether module headers should be italic
MODULE_BG_COLOR = "None"                         # Background color (None = transparent)
MODULE_TEXT_COLOR = "Black"                      # Text color for module headers
MODULE_PADDING = 0                               # No padding around module headers

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── TEMPLATE TEXT CONTENT SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Template main text content styling
TEMPLATE_TEXT_FONT_FAMILY = "Myriad Pro Cond"   # Font family for template text content
MODULE_TEXT_FONT_SIZE = 9                        # Font size for template text content
TEMPLATE_TEXT_BOLD = False                       # Whether template text should be bold
TEMPLATE_TEXT_ITALIC = False                     # Whether template text should be italic
TEMPLATE_TEXT_COLOR = "Black"                    # Text color for template content

# Video text within templates
VIDEO_TEXT_FONT_FAMILY = "Myriad Pro Cond"      # Font family for video text
VIDEO_TEXT_FONT_SIZE = 6                         # Font size for video text - compact text
VIDEO_TEXT_BOLD = False                          # Whether video text should be bold
VIDEO_TEXT_ITALIC = False                        # Whether video text should be italic
VIDEO_TEXT_COLOR = "Black"                       # Text color for video text

# Text container sizing settings to prevent overflow
TEXT_HEIGHT_BUFFER_FACTOR = 0.5                  # Much smaller buffer for very tight containers
OVERFLOW_INCREMENT_SIZE = 5                      # Very small increments when resizing for overflow
MAX_OVERFLOW_ITERATIONS = 10                     # Allow more iterations to fit full text
TEXT_HEIGHT_REDUCTION = 0.95                     # Reduce measured height by 5% for tighter fit

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── SUBTITLE & MISCELLANEOUS TEXT SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Subtitle styling (used in two-column layouts)
SUBTITLE_FONT_FAMILY = "Myriad Pro Cond"        # Font family for subtitles
SUBTITLE_FONT_SIZE = 11                          # Font size for subtitles
SUBTITLE_BOLD = True                             # Whether subtitles should be bold
SUBTITLE_ITALIC = False                          # Whether subtitles should be italic
SUBTITLE_TEXT_COLOR = "Black"                    # Text color for subtitles

# Roadsign label text styling
ROADSIGN_TEXT_FONT_FAMILY = "Myriad Pro Cond"   # Font family for roadsign labels
ROADSIGN_TEXT_FONT_SIZE = 6                      # Font size for roadsign labels
ROADSIGN_TEXT_BOLD = False                       # Whether roadsign text should be bold
ROADSIGN_TEXT_ITALIC = False                     # Whether roadsign text should be italic
ROADSIGN_TEXT_COLOR = "Black"                    # Text color for roadsign labels

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

# Quiz heading text styling
QUIZ_HEADING_FONT_FAMILY = "Myriad Pro Cond"  # Font family for quiz heading
QUIZ_HEADING_HEIGHT = 20                       # Height of the quiz heading
QUIZ_HEADING_TEXT = "QUIZ"                     # Text to display in the quiz heading
QUIZ_HEADING_FONT_SIZE = 14                    # Font size for quiz heading
QUIZ_HEADING_BOLD = True                       # Whether quiz heading should be bold
QUIZ_HEADING_ITALIC = False                    # Whether quiz heading should be italic
QUIZ_HEADING_TEXT_COLOR = "Black"              # Text color for quiz heading
QUIZ_HEADING_BG_COLOR = "None"                 # Background color for quiz heading

# Quiz question text styling
QUIZ_FONT_FAMILY = "Myriad Pro Cond"          # Primary font family for quiz questions
QUIZ_QUESTION_FONT_SIZE = 9                    # Font size for quiz questions (readable size)
QUIZ_QUESTION_BOLD = False                     # Whether quiz questions should be bold
QUIZ_QUESTION_ITALIC = False                   # Whether quiz questions should be italic
QUIZ_QUESTION_TEXT_COLOR = "Black"             # Text color for quiz questions

# Quiz answer text styling
QUIZ_ANSWER_FONT_FAMILY = "Myriad Pro Cond"   # Font family for quiz answers
QUIZ_ANSWER_FONT_SIZE = 14                     # Font size for answers (V/F boxes)
QUIZ_ANSWER_BOLD = True                        # Whether quiz answers should be bold
QUIZ_ANSWER_ITALIC = False                     # Whether quiz answers should be italic

# Quiz layout settings
QUIZ_LENGTH = PAGE_WIDTH - MARGINS[0] - MARGINS[2] - 20  # Width of quiz section (adjusted for wider page)
QUIZ_CARD_SPACING = 0                          # No spacing between question cards for compact layout
QUIZ_QUESTION_PADDING = 1                      # Minimal internal padding for question text
QUIZ_HORIZONTAL_SPACING = 4                    # Reduced space between question text and answer box
MIN_QUIZ_SPACING = 1                           # Reduced space between quiz items

# Answer box settings
QUIZ_ANSWER_BOX_WIDTH = 60                     # Width of the answer box
QUIZ_ANSWER_BOX_HEIGHT = 40                    # Height of the answer box
MIN_ANSWER_WIDTH = 150                         # Minimum width for answer column

# Answer indicators
QUIZ_TRUE_TEXT = "V"                           # Text for true answers
QUIZ_FALSE_TEXT = "F"                          # Text for false answers
QUIZ_TRUE_BORDER_COLOR = "Green"               # Border color for true answers
QUIZ_FALSE_BORDER_COLOR = "Red"                # Border color for false answers

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── VERTICAL BANNER SETTINGS ─────────────
# ────────────────────────────────────────────────────────────────────────────────

# Vertical topic banner layout
TOPIC_BANNER_WIDTH = 20         # Width of the vertical topic banner (in points)
TOPIC_BANNER_COLOR = "DarkGrey" # Default color of the vertical topic banner

# Banner margin offset settings
LEFT_BANNER_MARGIN_OFFSET = 2   # Gap between left banner and left margin (in points)
RIGHT_BANNER_MARGIN_OFFSET = 6   # Gap between right banner and right margin (in points)

# Left banner text positioning (odd pages)
LEFT_BANNER_HORIZONTAL_OFFSET = 4    # Distance from left edge of banner (in points)
LEFT_BANNER_VERTICAL_OFFSET = 400    # Vertical offset from center position (in points)
LEFT_BANNER_ROTATION = 90            # Rotation for left banner text (degrees)

# Right banner text positioning (even pages)
RIGHT_BANNER_HORIZONTAL_OFFSET = 4   # Distance from right edge of banner (in points)
RIGHT_BANNER_VERTICAL_OFFSET = 0     # Vertical offset from center position (in points)
RIGHT_BANNER_ROTATION = 270          # Rotation for right banner text (degrees)

# Banner text styling
BANNER_TEXT_FONT_FAMILY = "Myriad Pro Cond"  # Font family for banner text
BANNER_TEXT_FONT_SIZE = 10                    # Font size for banner text
BANNER_TEXT_BOLD = True                       # Whether banner text should be bold
BANNER_TEXT_ITALIC = False                    # Whether banner text should be italic
BANNER_TEXT_COLOR = "White"                   # Text color for banner text
BANNER_TEXT_HEIGHT_PERCENT = 0.75             # Percentage of banner height to use for text height 
