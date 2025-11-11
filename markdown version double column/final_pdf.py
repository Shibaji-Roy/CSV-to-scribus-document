#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
generate_pdf.py  – 2025-04-23
All template and quiz text are now chunk-paginated and each frame is trimmed
to its content so nothing ever overruns the bottom margin.
"""

import os
import re
import json
import scribus
# Removed BeautifulSoup - now using Markdown parser
from datetime import datetime

# Position tracking log files
LOG_DIR = "pdf_generation_logs"
DEBUG_LOG_FILE = None  # Will be set dynamically
OVERLAP_WARNINGS = []
BOUNDARY_VIOLATIONS = []

# Performance optimization: Disable detailed logging to improve speed
# Set to True only when debugging specific issues
ENABLE_DETAILED_LOGGING = False

def init_logging():
    """Initialize logging system with timestamped files"""
    global DEBUG_LOG_FILE, LOG_DIR
    try:
        # Get script directory - use absolute path
        script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
        LOG_DIR = os.path.join(script_dir, "pdf_generation_logs")

        # Create logs directory if it doesn't exist
        if not os.path.exists(LOG_DIR):
            os.makedirs(LOG_DIR)

        # Create timestamped log file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        DEBUG_LOG_FILE = os.path.join(LOG_DIR, f"positions_{timestamp}.log")

        with open(DEBUG_LOG_FILE, 'w', encoding='utf-8') as f:
            f.write("="*140 + "\n")
            f.write(f"PDF Generation Position Log - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Log file: {DEBUG_LOG_FILE}\n")
            f.write("="*140 + "\n")
            f.write("Format: Time | Page | Element Type | ID | X Y Width Height | Bottom | Space Remaining | Status | Notes\n")
            f.write("="*140 + "\n")
            f.write("\n")

        # Write a test message to confirm logging works
        log_position("INIT", "test", 0, 0, 0, 0, 0, f"Logging initialized at {LOG_DIR}")

    except Exception as e:
        # Fallback to current directory
        DEBUG_LOG_FILE = os.path.join(os.getcwd(), "element_positions.log")
        LOG_DIR = os.getcwd()
        try:
            with open(DEBUG_LOG_FILE, 'w', encoding='utf-8') as f:
                f.write(f"Logging fallback activated due to error: {str(e)}\n")
        except:
            pass

def log_position(element_type, element_id, x, y, width, height, page_num, notes=""):
    """Log element position for debugging"""
    global OVERLAP_WARNINGS, BOUNDARY_VIOLATIONS

    # Skip logging if disabled for performance
    if not ENABLE_DETAILED_LOGGING or not DEBUG_LOG_FILE:
        return

    try:
        safe_boundary = 595 - 28 - 22  # PAGE_HEIGHT - MARGINS[3] - buffer
        bottom = y + height
        space_remaining = safe_boundary - bottom

        # Detect issues
        status = "OK"
        if space_remaining < 0:
            status = "OVERLAP!"
            warning = f"Page {page_num}: {element_type} ID:{element_id} exceeds boundary by {abs(space_remaining):.1f}pt"
            BOUNDARY_VIOLATIONS.append(warning)
        elif space_remaining < 20:
            status = "WARNING"
            warning = f"Page {page_num}: {element_type} ID:{element_id} too close to boundary ({space_remaining:.1f}pt remaining)"
            OVERLAP_WARNINGS.append(warning)

        with open(DEBUG_LOG_FILE, 'a', encoding='utf-8') as f:
            timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
            f.write(f"{timestamp} | Page {page_num:2d} | {element_type:22s} | ID:{element_id:12s} | "
                   f"X:{x:6.1f} Y:{y:6.1f} W:{width:6.1f} H:{height:6.1f} | "
                   f"Bottom:{bottom:6.1f} | Space:{space_remaining:6.1f}pt | {status:8s} | {notes}\n")
    except:
        pass  # Silently ignore logging errors

def generate_summary_report():
    """Generate a summary report of issues found"""
    if not DEBUG_LOG_FILE:
        return

    try:
        summary_file = DEBUG_LOG_FILE.replace("positions_", "SUMMARY_")

        with open(summary_file, 'w', encoding='utf-8') as f:
            f.write("="*100 + "\n")
            f.write(f"PDF GENERATION SUMMARY REPORT\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("="*100 + "\n\n")

            # Boundary violations (critical issues)
            f.write(f"CRITICAL ISSUES - BOUNDARY VIOLATIONS ({len(BOUNDARY_VIOLATIONS)})\n")
            f.write("-"*100 + "\n")
            if BOUNDARY_VIOLATIONS:
                for i, violation in enumerate(BOUNDARY_VIOLATIONS, 1):
                    f.write(f"{i}. {violation}\n")
            else:
                f.write("None - all elements within page boundaries ✓\n")
            f.write("\n")

            # Warnings (potential issues)
            f.write(f"WARNINGS - CLOSE TO BOUNDARY ({len(OVERLAP_WARNINGS)})\n")
            f.write("-"*100 + "\n")
            if OVERLAP_WARNINGS:
                for i, warning in enumerate(OVERLAP_WARNINGS, 1):
                    f.write(f"{i}. {warning}\n")
            else:
                f.write("None - adequate spacing maintained ✓\n")
            f.write("\n")

            # Instructions
            f.write("="*100 + "\n")
            f.write("INSTRUCTIONS FOR ANALYSIS:\n")
            f.write("-"*100 + "\n")
            f.write(f"1. Review this summary file for critical issues and warnings\n")
            f.write(f"2. Open the detailed log file: {os.path.basename(DEBUG_LOG_FILE)}\n")
            f.write(f"3. Search for 'OVERLAP!' to find elements exceeding page boundaries\n")
            f.write(f"4. Search for 'WARNING' to find elements too close to boundaries\n")
            f.write(f"5. Look for TEMPLATE_START/TEMPLATE_END pairs to track template spacing\n")
            f.write(f"6. Check QUIZ_START positions relative to previous TEMPLATE_END\n")
            f.write(f"7. Share this summary and the detailed log for further analysis\n")
            f.write("="*100 + "\n")

        return summary_file
    except:
        return None

def clear_log():
    """Deprecated - use init_logging() instead"""
    init_logging()

# Import all configuration settings
try:
    from config import *
except ImportError:
    # Fallback values if config.py is missing
    import sys
    scribus.messageBox("Config Error", "Could not import config.py. Using default values.", scribus.ICON_WARNING)

# Ensure spacing constants are defined (for IDE/linting support)
try:
    HEADER_TO_DESC_SPACING
except NameError:
    HEADER_TO_DESC_SPACING = 2
try:
    MODULE_TO_TEMPLATE_SPACING
except NameError:
    MODULE_TO_TEMPLATE_SPACING = 1
try:
    TEMPLATE_TO_TEMPLATE_SPACING
except NameError:
    TEMPLATE_TO_TEMPLATE_SPACING = 3
try:
    SECTION_TO_SECTION_SPACING
except NameError:
    SECTION_TO_SECTION_SPACING = 8

try:
    from PIL import Image
except ImportError:
    Image = None

# Add image size cache and helper function
IMAGE_SIZE_CACHE = {}

def get_image_size(img_path):
    """Return (width, height) of image, using cache to avoid repeated disk I/O."""
    if img_path in IMAGE_SIZE_CACHE:
        return IMAGE_SIZE_CACHE[img_path]
    if Image:
        try:
            with Image.open(img_path) as im:
                size = im.size
                IMAGE_SIZE_CACHE[img_path] = size
                return size
        except:
            pass
    # Fallback if PIL is not available or image can't be opened
    size = (300, 200)
    IMAGE_SIZE_CACHE[img_path] = size
    return size

# Define common Myriad font variants to use across all headers and banners
MYRIAD_VARIANTS = ["Myriad Pro", "MyriadPro", "Myriad Pro Condensed", "MyriadPro-Cond", "Myriad"]

# Pre-register the Myriad Pro Cond font for quiz sections
QUIZ_ACTUAL_FONT = QUIZ_FONT_FAMILY  # Default to the configured font
try:
    # Check if the font exists in the Scribus available fonts
    available_fonts = scribus.getFontNames()
    
    # First check if the configured font exists directly
    if QUIZ_FONT_FAMILY in available_fonts:
        QUIZ_ACTUAL_FONT = QUIZ_FONT_FAMILY
    else:
        # Check if any of the aliases exist
        found_alias = False
        
        # Forward check: if alias in available fonts
        for alias, target in FONT_ALIASES.items():
            if alias in available_fonts and target == QUIZ_FONT_FAMILY:
                QUIZ_ACTUAL_FONT = alias
                found_alias = True
                scribus.messageBox("Font Alias Found", 
                                f"Using font '{alias}' as an alias for '{QUIZ_FONT_FAMILY}'.",
                                scribus.ICON_INFORMATION)
                break
                
        # Reverse check: if our configured font is an alias for an available font
        if not found_alias:
            for alias, target in FONT_ALIASES.items():
                if target in available_fonts and alias == QUIZ_FONT_FAMILY:
                    QUIZ_ACTUAL_FONT = target
                    found_alias = True
                    scribus.messageBox("Font Alias Found", 
                                    f"Using font '{target}' as a substitute for '{QUIZ_FONT_FAMILY}'.",
                                    scribus.ICON_INFORMATION)
                    break
                
        # If neither the font nor its aliases are found, use fallback
        if not found_alias:
            QUIZ_ACTUAL_FONT = FONT_CANDIDATES[0]  # Use first available font as fallback
            scribus.messageBox("Font Warning", 
                            f"{QUIZ_FONT_FAMILY} font not found. Using {QUIZ_ACTUAL_FONT} for quizzes instead.",
                            scribus.ICON_WARNING)
except:
    # If we can't check available fonts, use a safe default
    QUIZ_ACTUAL_FONT = FONT_CANDIDATES[0]  # Use first available font as fallback

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── RUNTIME STATE VARIABLES ────────────
# ────────────────────────────────────────────────────────────────────────────────
y_offset              = MARGINS[1]
global_template_count = 0
limit_reached         = False
CURRENT_COLOR         = BACKGROUND_COLORS[0]
current_topic_text    = None    # Currently active topic text
current_topic_color   = None    # Color for the current topic banner

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── PROPER SCRIBUS COLUMN MANAGEMENT ────────────
# ────────────────────────────────────────────────────────────────────────────────

class ScribusColumnManager:
    """Manages text layout using proper Scribus column API"""

    def __init__(self):
        self.text_frame_chain = []
        self.use_columns = USE_TWO_COLUMN_LAYOUT if 'USE_TWO_COLUMN_LAYOUT' in globals() else False
        self.quiz_mode = False
        self.current_column = 0  # For compatibility with test scripts

    def create_text_frame(self, text, font_size=6, use_columns=None, in_template=False):
        """Create a text frame with proper Scribus column support for clean column flow"""
        global y_offset

        if not text or not text.strip():
            return None

        if use_columns is None:
            use_columns = self.use_columns and not self.quiz_mode

        # Calculate frame dimensions for full page width (like in example image)
        frame_width = PAGE_WIDTH - MARGINS[0] - MARGINS[2]  # Use full page width for text
        available_height = PAGE_HEIGHT - y_offset - MARGINS[3] - 20

        # Only force new page if truly insufficient space (less than 1 line of text)
        if available_height < 10:
            new_page()
            available_height = PAGE_HEIGHT - y_offset - MARGINS[3] - 20

        # Measure the actual text height needed for proper frame sizing
        # This prevents overflow indicators by creating correctly sized frames
        estimated_text_height = measure_text_height(text, frame_width / 2 if use_columns else frame_width, in_template, font_size)

        # Use much larger safety margin to prevent overflow crosses
        safety_margin = max(estimated_text_height * 0.5, 30)  # 50% safety margin or 30pt minimum
        frame_height = min(max(estimated_text_height + safety_margin, 50), available_height)

        # Create text frame with proper calculated height
        frame_name = f"textframe_{len(self.text_frame_chain)}"
        frame = scribus.createText(MARGINS[0], y_offset, frame_width, frame_height, frame_name)

        # Configure columns for perfectly balanced text flow
        if use_columns:
            try:
                scribus.setColumns(2, frame)
                scribus.setColumnGap(COLUMN_GAP, frame)

                # Enable all column balancing features
                try:
                    # Block alignment for even text distribution
                    scribus.setTextAlignment(scribus.ALIGN_BLOCK, frame)
                except:
                    pass

                try:
                    # Enable column balancing mode
                    scribus.setColumnFillMode(scribus.COLUMN_FILL_BALANCE, frame)
                except:
                    pass

                try:
                    # Force equal column heights
                    scribus.setFirstLineOffsetPolicy(scribus.FLO_REALGLYPHHEIGHT, frame)
                except:
                    pass

            except:
                pass

        # Set text and properties
        scribus.setText(text, frame)

        # Set font and size - this is critical for templates to use 6px
        try:
            scribus.setFont(DEFAULT_FONT, frame)
            scribus.setFontSize(font_size, frame)
        except:
            for f in FONT_CANDIDATES:
                try:
                    scribus.setFont(f, frame)
                    scribus.setFontSize(font_size, frame)
                    break
                except:
                    pass

        # Force font size again for templates (ensuring correct font size)
        if in_template:
            try:
                scribus.setFontSize(MODULE_TEXT_FONT_SIZE, frame)
            except:
                pass

        # Apply universal overflow handling before other formatting
        frame = self._handle_text_overflow(frame, font_size)

        # Template text will be handled by the universal overflow handler

        # Set zero text padding for compact layout
        try:
            scribus.setTextDistances(0, 0, 0, 0, frame)  # No padding for compact text
        except:
            pass

        # Set compact line spacing
        try:
            scribus.setLineSpacingMode(scribus.FIXED_LINESPACING, frame)
            scribus.setLineSpacing(font_size * 1.0, frame)  # Tight line spacing for compact text
        except:
            pass

        # Layout the text properly
        try:
            scribus.layoutText(frame)
        except:
            pass

        # Handle overflow with column balancing in mind
        max_attempts = 20
        attempts = 0

        # First pass - expand frame to fit all text
        while scribus.textOverflows(frame) and attempts < max_attempts:
            current_height = scribus.getSize(frame)[1]
            # Smaller increments for precise fitting
            new_height = min(current_height + 10, available_height)
            if new_height <= current_height:
                break
            scribus.sizeObject(frame_width, new_height, frame)
            try:
                scribus.layoutText(frame)
            except:
                pass
            attempts += 1

        # Enhanced column balancing - ensure ALL text frames have perfectly equal columns
        if use_columns and BALANCE_COLUMNS:
            try:
                # Force layout first
                scribus.layoutText(frame)
                current_height = scribus.getSize(frame)[1]

                # Apply balancing to ALL frames with columns, even single line text
                num_lines = None
                try:
                    num_lines = scribus.getTextLines(frame)
                except:
                    # Fallback estimation if getTextLines fails
                    text_length = len(text)
                    chars_per_line = int(frame_width / (font_size * 0.6))  # Rough estimation
                    num_lines = max(1, text_length // chars_per_line)

                if num_lines and num_lines >= 1:  # Balance even single line text
                    line_height = font_size * 1.0

                    # Calculate optimal height for perfect column balance
                    lines_per_column = num_lines / 2.0
                    ideal_height = max(int(lines_per_column * line_height) + 5, font_size * 2)

                    # Enhanced aggressive balancing with more test points
                    if AGGRESSIVE_BALANCING:
                        test_heights = [
                            ideal_height,
                            ideal_height + line_height * 0.1,
                            ideal_height + line_height * 0.2,
                            ideal_height + line_height * 0.3,
                            ideal_height + line_height * 0.4,
                            ideal_height + line_height * 0.5,
                            ideal_height + line_height * 0.6,
                            ideal_height + line_height * 0.7,
                            ideal_height + line_height * 0.8,
                            ideal_height + line_height * 0.9,
                            ideal_height + line_height,
                            ideal_height + line_height * 1.1,
                            ideal_height + line_height * 1.2
                        ]
                    else:
                        test_heights = [
                            ideal_height,
                            ideal_height + line_height * 0.3,
                            ideal_height + line_height * 0.6,
                            ideal_height + line_height
                        ]

                    best_height = current_height
                    best_balanced = False

                    # Test each height for perfect balance
                    for test_height in test_heights:
                        if test_height < current_height and test_height >= font_size * 2:
                            try:
                                # Test this height
                                scribus.sizeObject(frame_width, test_height, frame)
                                scribus.layoutText(frame)

                                # Check if text fits without overflow
                                if not scribus.textOverflows(frame):
                                    best_height = test_height
                                    best_balanced = True
                                    # For aggressive mode, take the first working height (smallest)
                                    if AGGRESSIVE_BALANCING:
                                        break

                            except:
                                continue

                    # Apply the best balanced height
                    if best_balanced and best_height != current_height:
                        try:
                            scribus.sizeObject(frame_width, best_height, frame)
                            scribus.layoutText(frame)

                            # Double-check for overflow
                            if scribus.textOverflows(frame):
                                # Restore original height if balancing caused overflow
                                scribus.sizeObject(frame_width, current_height, frame)
                                scribus.layoutText(frame)
                        except:
                            # Restore on any error
                            try:
                                scribus.sizeObject(frame_width, current_height, frame)
                                scribus.layoutText(frame)
                            except:
                                pass

                    # Force final layout to ensure proper rendering
                    try:
                        scribus.layoutText(frame)
                    except:
                        pass

            except Exception as e:
                # Fallback - ensure frame is still functional
                try:
                    scribus.layoutText(frame)
                except:
                    pass

        # Get final frame height and position (Y position already updated by overflow handler)
        try:
            frame_height = scribus.getSize(frame)[1]
            frame_y = scribus.getPosition(frame)[1]

            # Only update Y if overflow handler didn't already update it
            expected_y = frame_y + frame_height + BLOCK_SPACING
            if y_offset < expected_y:
                y_offset = expected_y
        except:
            # Fallback if frame operations fail
            pass

        self.text_frame_chain.append(frame)
        return frame

    def set_quiz_mode(self, enabled):
        """Enable or disable quiz mode"""
        self.quiz_mode = enabled

    def get_current_y(self):
        """Get current Y position"""
        global y_offset
        return y_offset

    def set_current_y(self, new_y):
        """Set current Y position"""
        global y_offset
        y_offset = new_y

    def reset_for_new_page(self):
        """Reset column manager for a new page"""
        global y_offset
        y_offset = MARGINS[1]
        self.current_column = 0
        self.text_frame_chain = []

    def get_column_width(self):
        """Get the width of a column"""
        if self.quiz_mode:
            # Single column mode - use full page width (quizzes span entire width)
            return PAGE_WIDTH - MARGINS[0] - MARGINS[2]
        else:
            # Two column mode - use column width (full page width divided by 2)
            return COLUMN_WIDTH

    def get_column_x(self):
        """Get the X position for current column"""
        if self.quiz_mode:
            # Single column mode - start at left margin
            return MARGINS[0]
        else:
            # Two column mode - calculate based on current column
            if self.current_column == 0:
                return MARGINS[0]  # Left column
            else:
                return MARGINS[0] + COLUMN_WIDTH + COLUMN_GAP  # Right column

    def get_available_height(self):
        """Get available height from current Y position to bottom margin"""
        global y_offset
        return PAGE_HEIGHT - y_offset - MARGINS[3] - 20  # 20pt buffer

    def switch_column(self):
        """Switch to the next column (for compatibility with existing code)"""
        if not self.quiz_mode and self.use_columns:
            self.current_column = (self.current_column + 1) % 2

    @property
    def enabled(self):
        """Check if column layout is enabled (for compatibility)"""
        return self.use_columns and not self.quiz_mode

    def _handle_text_overflow(self, frame, font_size):
        """Handle text overflow using proper Scribus API documentation methods"""
        global y_offset

        try:
            # Step 1: Layout text first (required by Scribus API)
            scribus.layoutText(frame)

            # Step 2: Check if text overflows (returns 1 if overflow, 0 if not)
            if scribus.textOverflows(frame) == 0:
                return frame  # No overflow

            # Step 3: Get current frame dimensions
            frame_w, frame_h = scribus.getSize(frame)
            frame_x, frame_y = scribus.getPosition(frame)

            # Step 4: Calculate maximum allowed height
            bottom_boundary = PAGE_HEIGHT - MARGINS[3] - 40  # 20pt buffer
            max_height = bottom_boundary - frame_y

            # Step 5: Expand frame incrementally (following Scribus documentation pattern)
            h = frame_h

            # Coarse adjustment (10pt increments)
            while (scribus.textOverflows(frame) > 0) and (h < max_height):
                h += 10
                scribus.sizeObject(frame_w, h, frame)

            # Fine adjustment (1pt increments)
            while (scribus.textOverflows(frame) > 0) and (h < max_height):
                h += 1
                scribus.sizeObject(frame_w, h, frame)

            # Step 6: If still overflowing, force new page
            if scribus.textOverflows(frame) > 0:
                new_page()

                # Get text content before deleting frame
                text_content = scribus.getText(frame)
                scribus.deleteObject(frame)

                # Create new frame on new page
                new_frame = scribus.createText(MARGINS[0], y_offset, frame_w, h)
                scribus.setText(text_content, new_frame)

                # Apply same formatting as original
                try:
                    scribus.setFont(DEFAULT_FONT, new_frame)
                    scribus.setFontSize(font_size, new_frame)
                    scribus.setTextDistances(0, 0, 0, 0, new_frame)
                    scribus.layoutText(new_frame)
                except:
                    pass

                frame = new_frame

            # Step 7: Update Y position
            final_frame_y = scribus.getPosition(frame)[1]
            final_frame_h = scribus.getSize(frame)[1]
            y_offset = max(y_offset, final_frame_y + final_frame_h + BLOCK_SPACING)

            return frame

        except:
            return frame

    def ensure_consistent_balancing(self):
        """Ensure all text frames in chain have consistent column balancing"""
        if not self.use_columns or self.quiz_mode:
            return

        # Apply consistent balancing to all frames in the chain
        for frame in self.text_frame_chain:
            try:
                # Ensure frame has columns enabled
                scribus.setColumns(2, frame)
                scribus.setColumnGap(COLUMN_GAP, frame)

                # Force column balancing mode
                try:
                    scribus.setColumnFillMode(scribus.COLUMN_FILL_BALANCE, frame)
                except:
                    pass

                # Re-layout to ensure proper balance
                scribus.layoutText(frame)

                # Apply overflow handling after balancing
                self._handle_text_overflow(frame, 6)  # Use default content font size
            except:
                continue

# Global column manager instance
column_mgr = ScribusColumnManager()
PRINT_QUIZZES         = False   # Will be set based on user choice

def ensure_no_overlaps():
    """Ensure proper spacing and no overlaps between elements"""
    global y_offset
    # Add minimal spacing after each element
    y_offset += BLOCK_SPACING

    # Ensure we're not too close to bottom margin
    if y_offset > PAGE_HEIGHT - MARGINS[3] - 40:  # 50pt buffer
        new_page()

def safe_create_element(element_height, force_new_page_threshold=30):
    """Safely create an element with proper spacing checks"""
    global y_offset

    # Check if element would fit on current page
    available_space = PAGE_HEIGHT - MARGINS[3] - y_offset - 20  # 20pt buffer

    if element_height > available_space and available_space < force_new_page_threshold:
        new_page()

    return y_offset

def validate_template_text_frame(frame, text, font_size):
    """Validate template text frame using proper Scribus API methods"""
    if not frame:
        return frame

    try:
        # Layout text first (required by API)
        scribus.layoutText(frame)

        # Check for overflow using API method
        if scribus.textOverflows(frame) == 0:
            return frame  # No overflow

        # Apply standard overflow handling
        frame = column_mgr._handle_text_overflow(frame, font_size)

        return frame

    except:
        return frame

# Scribus constants (with backwards compatibility)
try:
    TEXT_FLOW_OBJECTBOUNDINGBOX = scribus.TEXT_FLOW_OBJECTBOUNDINGBOX
    TEXT_FLOW_INTERACTIVE       = scribus.TEXT_FLOW_INTERACTIVE
except AttributeError:
    TEXT_FLOW_OBJECTBOUNDINGBOX = 1
    TEXT_FLOW_INTERACTIVE       = 2

# Add this global variable near other runtime state variables
global quiz_heading_placed_on_page
quiz_heading_placed_on_page = False

def get_font_with_style(font_family, font_weight=None, font_style=None):
    """
    Find the best matching font based on family, weight and style.
    Returns the font name that most closely matches the requested attributes.
    """
    if not font_family:
        return DEFAULT_FONT
    
    # Clean and normalize input
    font_family = font_family.strip()
    if ',' in font_family:
        # Handle font stacks (comma-separated alternatives)
        font_options = [f.strip().strip('\'"') for f in font_family.split(',')]
    else:
        font_options = [font_family.strip().strip('\'"')]
    
    # Get available fonts from Scribus
    try:
        available_fonts = scribus.getFontNames()
    except:
        available_fonts = FONT_CANDIDATES
    
    is_bold = font_weight in ('bold', 'bolder', '700', '800', '900')
    is_italic = font_style == 'italic'
    
    for base_font in font_options:
        # Add Myriad Pro font variants to try
        base_fonts_to_try = [base_font]
        if "Myriad" in base_font:
            # Add common Myriad Pro variations
            base_fonts_to_try.extend([
                "Myriad Pro Condensed",
                "MyriadPro-Cond",
                "Myriad Pro Cond",
                "Myriad Pro",
                "MyriadPro",
                "Myriad"
            ])
            # Remove duplicates
            seen = set()
            base_fonts_to_try = [x for x in base_fonts_to_try if not (x in seen or seen.add(x))]

        for base_font_variant in base_fonts_to_try:
            # Build styled options for this variant
            styled_options = []
            if is_bold and is_italic:
                styled_options = [
                    f"{base_font_variant} Bold Italic",
                    f"{base_font_variant} Italic Bold",
                    f"{base_font_variant} BoldItalic",
                    f"{base_font_variant} Bold-Italic",
                    f"{base_font_variant}-Bold-Italic"
                ]
            elif is_bold:
                styled_options = [
                    f"{base_font_variant} Bold",
                    f"{base_font_variant}Bold",
                    f"{base_font_variant}-Bold"
                ]
            elif is_italic:
                styled_options = [
                    f"{base_font_variant} Italic",
                    f"{base_font_variant}Italic",
                    f"{base_font_variant}-Italic"
                ]
            else:
                styled_options = [
                    f"{base_font_variant} Regular",
                    f"{base_font_variant}Regular",
                    f"{base_font_variant}-Regular",
                    base_font_variant
                ]
            # Try styled options (case-sensitive, then case-insensitive)
            for styled_font in styled_options:
                if styled_font in available_fonts:
                    return styled_font
                for font in available_fonts:
                    if styled_font.lower() == font.lower():
                        return font
            # Try base font variant (case-sensitive, then case-insensitive, then substring)
            if base_font_variant in available_fonts:
                return base_font_variant
            for font in available_fonts:
                if base_font_variant.lower() == font.lower() or base_font_variant.lower() in font.lower():
                    return font
    # Fallback to default font
    return DEFAULT_FONT

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── HTML PARSING UTILITIES ────────────
# ────────────────────────────────────────────────────────────────────────────────
def parse_style_attribute(style_str):
    """Parse a CSS style attribute string into a dictionary of style properties."""
    styles = {}
    for part in style_str.split(";"):
        if ":" in part:
            k, v = part.split(":", 1)
            styles[k.strip().lower()] = v.strip().lower()
    return styles

def parse_markdown_to_segments(markdown_text):
    """
    Parse markdown text and convert to styled segments.
    Handles: **bold**, *italic*, {tip=N}text{end}, newlines
    Returns: List of (text, style_dict) tuples
    """
    if not markdown_text:
        return []
    
    markdown_text = markdown_text.strip()
    segments = []
    
    # Split by newlines to handle paragraphs
    lines = markdown_text.split('\n')
    
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            # Empty line = skip (no paragraph breaks to avoid large gaps)
            continue
        
        # Parse inline formatting in the line
        pos = 0
        while pos < len(line):
            # Check for bold **text**
            bold_match = re.match(r'\*\*(.+?)\*\*', line[pos:])
            if bold_match:
                text = bold_match.group(1)
                segments.append((text, {"bold": True, "font-weight": "bold"}))
                pos += len(bold_match.group(0))
                continue
            
            # Check for italic *text*
            italic_match = re.match(r'\*(.+?)\*', line[pos:])
            if italic_match:
                text = italic_match.group(1)
                segments.append((text, {"italic": True, "font-style": "italic"}))
                pos += len(italic_match.group(0))
                continue
            
            # Check for {tip=N}text{end}
            tip_match = re.match(r'\{tip=(\d+)\}(.+?)\{end\}', line[pos:])
            if tip_match:
                tip_id = tip_match.group(1)
                text = tip_match.group(2)
                segments.append((text, {"color": "#00ae00", "tip_id": tip_id}))
                pos += len(tip_match.group(0))
                continue
            
            # Regular text until next special character
            next_special = None
            for pattern in [r'\*\*', r'\*', r'\{tip=']:
                match = re.search(pattern, line[pos:])
                if match:
                    if next_special is None:
                        next_special = pos + match.start()
                    else:
                        next_special = min(next_special, pos + match.start())
            
            # If no special chars found, take rest of line
            if next_special is None:
                text = line[pos:]
                if text:
                    segments.append((text, {}))
                break  # Done with this line
            elif next_special > pos:
                # Text before next special char
                text = line[pos:next_special]
                if text:
                    segments.append((text, {}))
                pos = next_special
            else:
                # Should not happen, but safety
                pos += 1
        
        # Add space (not newline) after each line except the last one
        # This prevents large gaps between text segments
        if i < len(lines) - 1:
            segments.append((' ', {}))
    
    return segments
# ────────────────────────────────────────────────────────────────────────────────
# ─────────── LAYOUT UTILITY FUNCTIONS ────────────
# ────────────────────────────────────────────────────────────────────────────────
def find_fit(text, frame):
    """Find the maximum amount of text that fits in a frame."""
    if not text:
        return 0

    # Get current frame size - use full height, no artificial reduction
    frame_width, frame_height = scribus.getSize(frame)

    # Create a temporary invisible frame with the full height
    temp_frame = scribus.createText(0, 0, frame_width, frame_height)
    try:
        # Copy the text distances (padding) from the original frame
        try:
            left, right, top, bottom = scribus.getTextDistances(frame)
            scribus.setTextDistances(left, right, top, bottom, temp_frame)
        except:
            pass
            
        # Copy the font from the original frame
        try:
            font = scribus.getFont(frame)
            scribus.setFont(font, temp_frame)
        except:
            # Set a default font if we can't get the original
            try:
                scribus.setFont(DEFAULT_FONT, temp_frame)
            except:
                for f in FONT_CANDIDATES:
                    try:
                        scribus.setFont(f, temp_frame)
                        break
                    except:
                        continue
        
        # Apply the same font size if possible
        try:
            scribus.setFontSize(font_size, temp_frame)
        except:
            pass
            
        # Standard binary search to find how much text fits
        lo, hi, best = 0, len(text), 0
        while lo <= hi:
            mid = (lo + hi) // 2
            scribus.setText(text[:mid], temp_frame)
            if scribus.textOverflows(temp_frame):
                hi = mid - 1
            else:
                best = mid
                lo = mid + 1
                
        result = best
    finally:
        # Clean up the temporary frame
        scribus.deleteObject(temp_frame)
        
    return result

def space_left_on_page():
    """Calculate remaining vertical space on the current page with page number buffer."""
    global column_mgr
    # Leave 20 points buffer for page number (15pt height + 3px below + clearance)
    current_y = column_mgr.get_current_y()
    return PAGE_HEIGHT - current_y - MARGINS[3] - 20

def enforce_margin_boundary():
    """Ensure y_offset never exceeds the bottom margin boundary with buffer for page number."""
    global y_offset, column_mgr
    # Leave 20 points buffer for page number (15pt height + 3px below + clearance)
    safe_boundary = PAGE_HEIGHT - MARGINS[3] - 40

    current_y = column_mgr.get_current_y()
    if current_y > safe_boundary:
        column_mgr.set_current_y(safe_boundary)

    # Keep legacy y_offset in sync
    if y_offset > safe_boundary:
        y_offset = safe_boundary

def simple_boundary_check(element_height):
    """
    Simple boundary check - create new page if element won't fit.
    Uses y_offset directly since column_mgr may not be synchronized.
    """
    global y_offset, column_mgr
    safe_boundary = PAGE_HEIGHT - MARGINS[3] - 40  # Page number buffer

    # Use y_offset directly instead of column_mgr for accurate check
    would_exceed = y_offset + element_height > safe_boundary

    # Log large elements that might cause issues
    if element_height > 200:
        current_page = scribus.currentPage()
        log_position("BOUNDARY_CHECK_LARGE", "element", MARGINS[0], y_offset, 0, element_height,
                    current_page, f"Height={element_height:.1f}pt, y={y_offset:.1f}, bottom would be {y_offset + element_height:.1f}, boundary={safe_boundary:.1f}, exceeds={would_exceed}")

    if would_exceed:
        new_page()
        return True  # New page was created
    return False  # Element fits

def precise_fit_image_frame(img_frame):
    """
    Use official Scribus function to adjust frame to image size, eliminating gaps.
    Simple approach using setScaleFrameToImage which is designed for this purpose.
    """
    try:
        # Use the official Scribus function to adjust frame to image size
        scribus.setScaleFrameToImage(img_frame)

        # Ensure image is positioned properly within the adjusted frame
        scribus.setImageOffset(0, 0, img_frame)

        # Force image to fill the frame after adjustment
        scribus.setScaleImageToFrame(True, True, img_frame)

    except Exception:
        # Fallback to standard image scaling if setScaleFrameToImage doesn't work
        try:
            scribus.setScaleImageToFrame(True, True, img_frame)
        except:
            pass

def simple_constrain_element(element_obj):
    """
    Simple constraint - just prevent element from exceeding boundary.
    """
    try:
        element_x, element_y = scribus.getPosition(element_obj)
        element_w, element_h = scribus.getSize(element_obj)
        max_allowed_bottom = PAGE_HEIGHT - MARGINS[3] - 40  # Page number buffer
        
        if element_y + element_h > max_allowed_bottom:
            new_height = max_allowed_bottom - element_y
            if new_height > 10:  # Only if there's some space
                scribus.sizeObject(element_w, new_height, element_obj)
        return True
    except:
        return False

def can_place_element_safely(element_height, min_safe_margin=22):
    """Check if an element can be placed without exceeding bottom margin with page number buffer."""
    global y_offset
    # Use same safe boundary as enforce_margin_boundary
    bottom_boundary = PAGE_HEIGHT - MARGINS[3] - min_safe_margin
    return (y_offset + element_height) <= bottom_boundary

def force_new_page_if_needed(element_height, min_safe_margin=22):
    """Force new page if element would exceed bottom margin with page number buffer."""
    global y_offset
    if not can_place_element_safely(element_height, min_safe_margin):
        new_page()
        return True
    return False

def can_fit_content(content_height, min_required=MIN_SPACE_THRESHOLD):
    """
    Determines if content can fit in remaining page space.
    More flexible than just checking against MIN_SPACE_THRESHOLD.
    
    Args:
        content_height: Height of the content to check
        min_required: Minimum space required (defaults to MIN_SPACE_THRESHOLD)
    
    Returns:
        True if content can fit, False otherwise
    """
    space_available = space_left_on_page()
    
    # If plenty of space, return True
    if space_available >= content_height + BLOCK_SPACING:
        return True
    
    # If very small space left, return False
    if space_available < min_required:
        return False
    
    # Special case: If we have at least 60% of the height needed,
    # and the content is a modest size (not a huge block),
    # allow it to fit to better utilize page space
    if (space_available >= content_height * 0.6 and 
        content_height < 100 and 
        space_available >= 40):
        return True
        
    return False

def measure_text_height(text, width, in_template=False, font_size=8):
    """
    More accurate text height measurement using actual Scribus measurement.
    """
    if not text:
        return font_size * 2

    # Create a temporary probe frame with generous height for accurate measurement
    probe_height = max(PAGE_HEIGHT * 0.5, 200)  # Half page or 200pt minimum
    probe = scribus.createText(MARGINS[0], y_offset, width, probe_height)

    try:
        # Remove border
        scribus.setLineColor("None", probe)

        # Apply padding
        try:
            if in_template:
                scribus.setTextDistances(*TEMPLATE_TEXT_PADDING, probe)
            else:
                scribus.setTextDistances(*REGULAR_TEXT_PADDING, probe)
        except:
            pass

        # Set font
        try:
            scribus.setFont(DEFAULT_FONT, probe)
            scribus.setFontSize(font_size, probe)
        except:
            for f in FONT_CANDIDATES:
                try:
                    scribus.setFont(f, probe)
                    scribus.setFontSize(font_size, probe)
                    break
                except:
                    continue

        # Set fixed line spacing for consistent measurement (documentation-based approach)
        try:
            scribus.setLineSpacing(font_size * 1.0, probe)
        except:
            pass

        # Set text
        scribus.setText(text, probe)

        # Refresh layout
        try:
            scribus.layoutText(probe)
        except:
            pass

        # If there's overflow in this generous frame, we need to expand
        if scribus.textOverflows(probe):
            # Text is very long, estimate more aggressively
            words = len(text.split())
            lines_estimate = max(words // 8, len(text) // 80)  # Rough estimates
            needed_height = lines_estimate * font_size * 1.3
        else:
            # Try to measure more precisely
            try:
                # Get number of actual lines displayed
                num_lines = scribus.getTextLines(probe)
                if num_lines > 0:
                    try:
                        line_spacing = scribus.getLineSpacing(probe)
                    except:
                        line_spacing = font_size * 1.2

                    # Calculate height based on actual lines
                    text_height = num_lines * line_spacing

                    # Add padding
                    try:
                        left, right, top, bottom = scribus.getTextDistances(probe)
                        needed_height = text_height + top + bottom
                    except:
                        needed_height = text_height
                else:
                    # Fallback to estimation
                    needed_height = font_size * 3
            except:
                # Final fallback
                needed_height = len(text) * 0.3  # Very rough estimate

        scribus.deleteObject(probe)

        # Return precise measurement with minimal safety margin
        return max(needed_height + font_size * 0.5, font_size * 2)

    except:
        # If probe creation fails, fallback to simple estimation
        try:
            scribus.deleteObject(probe)
        except:
            pass

        # Simple fallback estimation
        estimated_lines = max(len(text) // 50, 1)
        return estimated_lines * font_size * 1.5

def new_page():
    """Create a new page in the document and reset the y position."""
    global y_offset, column_mgr
    global quiz_heading_placed_on_page
    scribus.newPage(-1)
    y_offset = MARGINS[1]

    # Reset column manager for new page
    column_mgr.reset_for_new_page()

    # Force refresh to show live page creation updates
    scribus.redrawAll()
    # Reset quiz header flag for new page - each page gets its own quiz header
    quiz_heading_placed_on_page = False
    # Add vertical topic banner to the new page if we have an active topic
    # This will automatically position it based on whether the new page is odd or even
    if current_topic_text:
        create_vertical_topic_banner()

    # Add page number to new page - exactly like copy 6
    add_page_number()

def add_page_number():
    """Add page number outside bottom margin - exactly like copy 6"""
    global PAGE_WIDTH, PAGE_HEIGHT, MARGINS
    
    # Get current page number
    page_num = scribus.pageCount()
    
    # Create page number text box outside the margin area
    page_num_width = 30
    page_num_height = 15
    x_pos = (PAGE_WIDTH - page_num_width) / 2  # Center horizontally
    # Position 3 pixels below the margin area (outside printable area)
    y_pos = PAGE_HEIGHT - MARGINS[3] + 3  # Outside margin by 3 pixels
    
    try:
        page_num_box = scribus.createText(x_pos, y_pos, page_num_width, page_num_height)
        scribus.setText(str(page_num), page_num_box)
        
        # Set font and formatting
        font_applied = False
        for font in FONT_CANDIDATES:
            try:
                scribus.setFont(font, page_num_box)
                font_applied = True
                break
            except:
                pass

        # Final fallback if FONT_CANDIDATES didn't work
        if not font_applied:
            try:
                scribus.setFont(DEFAULT_FONT, page_num_box)
            except:
                pass
        
        try:
            scribus.setFontSize(8, page_num_box)  # Small font for page number
            scribus.setTextAlignment(1, page_num_box)  # Center align
            scribus.setTextColor("Black", page_num_box)
        except:
            pass
    except Exception as e:
        # Ignore errors in page number creation
        pass


def place_roadsigns_grid(images, base_path, start_y, max_height=None, frame_x=None, frame_width=None):
    """
    Place roadsigns in a grid pattern where:
    - Maximum of 2 images per row (unlike regular images which allow 3)
    - After the first row of 2, the next row has images 3 and 4
    - Then row with images 5 and 6, and so on
    
    This is specifically for roadsigns with a different row limit.
    All images in a group will have the same height for visual consistency.
    """
    global y_offset
    
    if not images:
        return start_y
    
    # Set default target height, much smaller for roadsigns
    target_height = min(25, max_height) if max_height else 25

    # Detect attention signs and make them smaller
    attention_signs = []
    for img in images:
        # Check if this might be an attention sign (typically larger/different aspect ratio)
        if img and Image:
            try:
                img_path = os.path.join(base_path, img)
                with Image.open(img_path) as im:
                    w, h = im.size
                    aspect_ratio = w / h if h > 0 else 1
                    # Attention signs are often square or have specific aspect ratios
                    # and may be larger than typical road signs
                    if (0.8 <= aspect_ratio <= 1.2) or w > 200 or h > 200:
                        attention_signs.append(img)
            except:
                pass

    # Available width for images - use frame width if provided, otherwise column width
    if frame_width is not None:
        available_width = frame_width
    else:
        available_width = column_mgr.get_column_width()

    # Adapt roadsigns per row based on available width
    if available_width < 120:
        images_per_row = 1  # Single image per row for very narrow columns
    elif available_width < 200:
        images_per_row = 2  # Standard for narrow columns
    else:
        images_per_row = 3  # Allow more for wider areas
    
    # Process images in adaptive rows
    current_y = start_y
    row_start_idx = 0
    
    while row_start_idx < len(images):
        # Get images for this row (up to 2)
        row_end_idx = min(row_start_idx + images_per_row, len(images))
        row_images = images[row_start_idx:row_end_idx]
        
        # Calculate dimensions for all images based on aspect ratios
        widths = []
        for rel in row_images:
            img_path = os.path.join(base_path, rel)
            if Image:
                try:
                    with Image.open(img_path) as im:
                        orig_w, orig_h = im.size
                except:
                    orig_w, orig_h = (300, 200)
            else:
                orig_w, orig_h = (300, 200)
            scaled_w = (orig_w / orig_h) * target_height if orig_h else 150
            widths.append(scaled_w)

        # Adjust if total width exceeds available space
        total_width = sum(widths) + (len(widths) - 1) * BLOCK_SPACING
        if total_width > available_width:
            ratio = available_width / total_width
            adjusted_height = target_height * ratio
            widths = [w * ratio for w in widths]
            total_width = available_width
        else:
            adjusted_height = target_height

        # Right align for roadsigns within the available space
        if frame_x is not None:
            # Right align within the provided frame
            x = frame_x + available_width - total_width
        elif column_mgr.enabled:
            # Right align within the column
            x = column_mgr.get_column_x() + available_width - total_width
        else:
            # Right align for roadsigns (standard for roadsigns) - like final_pdf copy
            x = PAGE_WIDTH - MARGINS[2] - total_width

        # Place images in this row
        for idx, rel in enumerate(row_images):
            img_path = os.path.join(base_path, rel)
            w_i = widths[idx]

            # Use smaller height for attention signs
            if rel in attention_signs:
                h_i = adjusted_height * 0.6  # Make attention signs 40% smaller
            else:
                h_i = adjusted_height
            img_frame = scribus.createImage(x, current_y, w_i, h_i)
            scribus.loadImage(img_path, img_frame)
            scribus.setScaleImageToFrame(True, True, img_frame)
            scribus.setLineColor("None", img_frame)
            # Simple boundary enforcement for image
            simple_constrain_element(img_frame)
            
            # Enable shaped text wrap for roadsigns
            if scribus.getObjectType(img_frame) == "ImageFrame":
                try:
                    scribus.setItemShapeSetting(img_frame, scribus.ITEM_BOUNDED_TEXTFLOW)
                except:
                    pass
                try:
                    scribus.setTextFlowMode(img_frame, TEXT_FLOW_OBJECTBOUNDINGBOX)
                except:
                    pass
            
            x += w_i + BLOCK_SPACING

        # Check if next position would exceed margin before updating
        next_y = current_y + adjusted_height + BLOCK_SPACING
        safe_boundary = PAGE_HEIGHT - MARGINS[3] - 40  # Page number buffer
        if next_y > safe_boundary:
            # Force to safe boundary if would exceed
            current_y = safe_boundary
        else:
            # Update y position after this row
            current_y += adjusted_height + BLOCK_SPACING
        
        # Move to next row
        row_start_idx = row_end_idx
    
    # Return the final y position
    return current_y

# ─────────── TEXT PLACEMENT FUNCTIONS ────────────
def place_text_block_flow(markdown_text, font_size=8, bold=False, in_template=False, no_bottom_gap=False, is_heading=False, balanced_columns=False, custom_spacing=None, font_family=None, has_roadsigns=False):
    """
    Place an HTML text block with flowing text and formatting.
    Uses the proven working approach with optional two-column layout for descriptions.
    Headings use single column, descriptions can use two-column layout.
    balanced_columns: If True, split text equally between two column frames instead of flowing.
    font_family: Optional font family to use (defaults to DEFAULT_FONT if not specified).
    """
    global y_offset

    if not markdown_text:  # Skip empty text blocks
        return

    # Format HTML with BeautifulSoup to handle malformed HTML better
    segments = parse_markdown_to_segments(markdown_text)

    # Normalize segments for better handling of bold tags and newlines
    normalized_segments = []
    plain_parts = []

    # Process segments to handle newlines consistently
    for t, sty in segments:
        if "\n" in t:
            # If the segment contains newlines, split into multiple segments
            parts = t.split("\n")
            for i, part in enumerate(parts):
                if part:  # Only add non-empty parts
                    normalized_segments.append((part, sty))
                # Add newline after all parts except the last one
                if i < len(parts) - 1:
                    normalized_segments.append(("\n", {}))
        else:
            normalized_segments.append((t, sty))

    # Collect plain text parts to measure frame height
    for t, _ in normalized_segments:
        if t != "\n" or (plain_parts and plain_parts[-1] != "\n"):
            plain_parts.append(t)

    plain = "".join(plain_parts)
    if not plain:
        return

    # Use standard margins for frame position
    frame_w = PAGE_WIDTH - MARGINS[0] - MARGINS[2]

    # Check if we should use balanced columns (split text into two equal frames)
    if balanced_columns and USE_TWO_COLUMN_LAYOUT and not is_heading:
        # For balanced columns, create two separate frames side by side
        col_width = (frame_w - COLUMN_GAP) / 2

        # Split the SEGMENTS (with styles) for proper style preservation
        # Strategy: Split at word boundaries within segments for fine-grained balance

        # First, expand segments into word-level segments for better splitting
        word_segments = []
        for txt, sty in normalized_segments:
            if txt == "\n":
                # Keep newlines as separate segments
                word_segments.append((txt, sty))
            else:
                # Split text into words while preserving styles
                words = txt.split()
                for i, word in enumerate(words):
                    # Add space after each word except the last one
                    if i < len(words) - 1:
                        word_segments.append((word + " ", sty))
                    else:
                        word_segments.append((word, sty))

        # Calculate total length
        total_words = len(word_segments)
        if total_words == 0:
            # Fallback to original segments if no words
            word_segments = normalized_segments
            total_words = len(word_segments)

        # Find the best split point by trying different positions
        mid_word = total_words // 2
        best_split_idx = mid_word
        best_height_diff = float('inf')

        # Search range: +/- 10% of total words for optimal balance (reduced from 30% for performance)
        search_range = max(5, min(int(total_words * 0.1), 30))  # Cap at 30 iterations max

        for test_split_idx in range(max(1, mid_word - search_range), min(total_words, mid_word + search_range + 1)):
            # Build plain text for height measurement
            left_test = ''.join(seg[0] for seg in word_segments[:test_split_idx])
            right_test = ''.join(seg[0] for seg in word_segments[test_split_idx:])

            # Skip if one side is empty
            if not left_test.strip() or not right_test.strip():
                continue

            # Measure heights
            left_h = measure_text_height(left_test, col_width, in_template, font_size)
            right_h = measure_text_height(right_test, col_width, in_template, font_size)

            height_diff = abs(left_h - right_h)

            if height_diff < best_height_diff:
                best_height_diff = height_diff
                best_split_idx = test_split_idx

                # Early exit if columns are already well-balanced (within 5pt)
                if best_height_diff < 5:
                    break

        # Split word segments at the best split point
        left_segments = word_segments[:best_split_idx]
        right_segments = word_segments[best_split_idx:]

        # Build plain text for each column (for height measurement)
        left_plain = ''.join(seg[0] for seg in left_segments)
        right_plain = ''.join(seg[0] for seg in right_segments)

        # Measure heights
        left_h = measure_text_height(left_plain, col_width, in_template, font_size)
        right_h = measure_text_height(right_plain, col_width, in_template, font_size)
        max_h = max(left_h, right_h)

        # Simple boundary check - create new page if needed
        simple_boundary_check(max_h)

        # Create minimal frames and let overflow handling expand to exact size
        minimal_height = font_size * 2  # Start with minimal height
        left_frame = scribus.createText(MARGINS[0], y_offset, col_width, minimal_height)
        right_frame = scribus.createText(MARGINS[0] + col_width + COLUMN_GAP, y_offset, col_width, minimal_height)

        # Use the maximum height for y_offset calculation
        text_h = max_h

        # Store segments with frames for style application
        frame = left_frame
        frames_to_setup = [(left_frame, left_segments), (right_frame, right_segments)]

    else:
        # Standard single frame or flowing columns
        text_h = measure_text_height(plain, frame_w, in_template, font_size)

        # Simple boundary check - create new page if needed
        simple_boundary_check(text_h)

        frame = scribus.createText(MARGINS[0], y_offset, frame_w, text_h)

        # Add two-column layout for descriptions AND templates (not headings)
        # Templates need two-column layout for equal distribution
        if USE_TWO_COLUMN_LAYOUT and not is_heading:
            try:
                scribus.setColumns(2, frame)
                scribus.setColumnGap(COLUMN_GAP, frame)
            except:
                pass  # If columns fail, continue with single column

        frames_to_setup = [(frame, plain)]

    # Process each frame (either single frame or two balanced column frames)
    for current_frame, frame_data in frames_to_setup:
        # Remove the frame border
        try:
            scribus.setLineColor("None", current_frame)
        except:
            pass

        # Set internal padding for all text frames
        try:
            if in_template:
                scribus.setTextDistances(*TEMPLATE_TEXT_PADDING, current_frame)
            else:
                scribus.setTextDistances(*REGULAR_TEXT_PADDING, current_frame)
        except:
            pass

        # frame_data is either a string (plain text) or a list of segments (text, style) tuples
        if isinstance(frame_data, str):
            # Single frame with plain text
            frame_text = frame_data
            scribus.setText(frame_text, current_frame)
        else:
            # Balanced columns with segments - extract plain text
            frame_text = ''.join(seg[0] for seg in frame_data)
            scribus.setText(frame_text, current_frame)

        # For balanced columns, apply formatting to BOTH frames
        if balanced_columns:
            # Apply basic formatting for ALL balanced column frames
            # Set the font - use provided font_family or DEFAULT_FONT
            font_set = False

            # Determine which font to use
            target_font = font_family if font_family else DEFAULT_FONT

            # Build list of font variants to try
            font_variants_to_try = [target_font]

            # Add common variations if it's Myriad
            if "Myriad" in target_font:
                font_variants_to_try.extend([
                    "Myriad Pro Condensed",
                    "MyriadPro-Cond",
                    "Myriad Pro Cond",
                    "Myriad Pro",
                    "MyriadPro",
                    "Myriad"
                ])
            else:
                # Generic variations
                font_variants_to_try.extend([
                    target_font.replace(" ", ""),  # Remove spaces
                    target_font.replace(" ", "-")  # Spaces to dashes
                ])

            # Remove duplicates while preserving order
            seen = set()
            font_variants_to_try = [x for x in font_variants_to_try if not (x in seen or seen.add(x))]

            # Try all font variants
            for variant in font_variants_to_try:
                try:
                    scribus.setFont(variant, current_frame)
                    font_set = True
                    break
                except:
                    pass

            # Fallback to FONT_CANDIDATES if nothing worked
            if not font_set:
                for f in FONT_CANDIDATES:
                    try:
                        scribus.setFont(f, current_frame)
                        font_set = True
                        break
                    except:
                        continue

            # Apply font size and bold to all text
            try:
                # Select all text and set font size
                scribus.selectText(0, len(frame_text), current_frame)
                scribus.setFontSize(font_size, current_frame)
                if bold:
                    scribus.setFontFeatures("Bold", current_frame)
                # Deselect to apply
                scribus.deselectAll()
            except:
                pass

            # Set very tight line spacing for more compact layout
            try:
                # For very small fonts (descriptions), use even tighter spacing
                if font_size <= 7:
                    scribus.setLineSpacing(font_size * 1.0, current_frame)  # No extra spacing for small text
                else:
                    scribus.setLineSpacing(font_size * 1.0, current_frame)  # Reduced from 1.2 to 1.1
            except:
                pass

            # Remove paragraph spacing
            try:
                scribus.selectText(0, scribus.getTextLength(current_frame), current_frame)
                scribus.setParagraphGap(0, current_frame)
            except:
                pass

            # Set vertical alignment
            try:
                scribus.setTextVerticalAlignment(scribus.ALIGNV_TOP, current_frame)
            except:
                pass

            # Handle overflow for each column frame with adaptive expansion
            overflow_iterations = 0
            max_overflow_iterations = 200  # Increased safety limit
            while scribus.textOverflows(current_frame) and overflow_iterations < max_overflow_iterations:
                current_w, current_h = scribus.getSize(current_frame)
                # Adaptive increment for faster convergence
                if overflow_iterations < 50:
                    increment = 10
                elif overflow_iterations < 100:
                    increment = 5
                else:
                    increment = 2
                scribus.sizeObject(current_w, current_h + increment, current_frame)
                scribus.layoutText(current_frame)
                overflow_iterations += 1

    # Apply styles for both balanced and non-balanced frames
    if balanced_columns and len(frames_to_setup) > 1:
        # For balanced columns, apply styles to each frame based on its segments
        for current_frame, frame_segments in frames_to_setup:
            # Convert segments to style_segments format for handle_text_styles
            frame_style_segments = []
            pos = 0
            for txt, sty in frame_segments:
                if len(txt) > 0:
                    # Ensure text is visible against the current background
                    if in_template and "color" not in sty:
                        # Templates have white/transparent backgrounds, so always use black text
                        sty = dict(sty)  # Make a copy to avoid modifying original
                        sty["color"] = "Black"

                    frame_style_segments.append((pos, len(txt), sty))
                    pos += len(txt)

            # Apply styles to this frame
            handle_text_styles(current_frame, frame_style_segments, font_size, font_family)

            # CRITICAL: Re-check overflow AFTER applying styles, as styles can cause text reflow
            # Use adaptive expansion for faster and more reliable convergence
            overflow_iterations = 0
            max_overflow_iterations = 200  # Increased limit
            while scribus.textOverflows(current_frame) and overflow_iterations < max_overflow_iterations:
                current_w, current_h = scribus.getSize(current_frame)
                # Adaptive increment: start large, then get smaller for fine-tuning
                if overflow_iterations < 50:
                    increment = 10
                elif overflow_iterations < 100:
                    increment = 5
                else:
                    increment = 2
                scribus.sizeObject(current_w, current_h + increment, current_frame)
                scribus.layoutText(current_frame)
                overflow_iterations += 1

                # Force redraw every 20 iterations to ensure accurate overflow detection
                if overflow_iterations % 20 == 0:
                    try:
                        scribus.redrawAll()
                    except:
                        pass

    else:
        # Single frame - create style segments from normalized_segments and apply
        style_segments = []
        pos = 0
        for txt, sty in normalized_segments:
            if len(txt) > 0:
                # Ensure text is visible against the current background
                if in_template and "color" not in sty:
                    # Templates have white/transparent backgrounds, so always use black text
                    sty = dict(sty)  # Make a copy
                    sty["color"] = "Black"

                style_segments.append((pos, len(txt), sty))
                pos += len(txt)

        # Apply styles to single frame
        handle_text_styles(frame, style_segments, font_size, font_family)

        # Force font size for area and topic descriptions
        if not in_template and font_size <= max(AREA_DESC_FONT_SIZE, TOPIC_DESC_FONT_SIZE):
            try:
                scribus.selectText(0, len(plain), frame)
                scribus.setFontSize(font_size, frame)
            except:
                pass

        # Set very tight line spacing for ALL text (both template and non-template) for consistent measurement
        try:
            # CRITICAL: Use FIXED line spacing mode to override any paragraph-level spacing
            # This prevents extra gaps when text wraps around images
            scribus.setLineSpacingMode(scribus.FIXED_LINESPACING, frame)

            # For very small fonts (descriptions), use even tighter spacing
            if font_size <= 7:
                scribus.setLineSpacing(font_size * 1.0, frame)  # No extra spacing for small text
            else:
                scribus.setLineSpacing(font_size * 1.0, frame)  # Reduced from 1.2 to 1.1
        except:
            pass

        # CRITICAL: Remove all paragraph spacing to eliminate inline gaps
        try:
            # Select all text in the frame
            scribus.selectText(0, scribus.getTextLength(frame), frame)
            # Set paragraph spacing before and after to 0
            scribus.setParagraphGap(0, frame)  # Gap after paragraph
        except:
            pass

        # Apply vertical justification to distribute text evenly
        try:
            # Use TOP alignment to keep text at top and measure exact height needed
            scribus.setTextVerticalAlignment(scribus.ALIGNV_TOP, frame)
        except:
            pass

        # Apply bold to entire frame if requested
        if bold:
            try:
                scribus.selectText(0, len(plain), frame)
                # Just make it bold, don't increase size
                scribus.setFontFeatures("Bold", frame)
            except:
                pass

        # PROVEN OVERFLOW HANDLING from working file
        try:
            # Force text refresh to ensure all styling is applied
            scribus.redrawAll()

            # CRITICAL: Reapply line spacing after all HTML styles to ensure consistency
            # This fixes the issue where left side (wrapping around image) has different spacing than right side
            try:
                scribus.selectText(0, scribus.getTextLength(frame), frame)
                scribus.setLineSpacingMode(scribus.FIXED_LINESPACING, frame)
                if font_size <= 7:
                    scribus.setLineSpacing(font_size * 1.0, frame)
                else:
                    scribus.setLineSpacing(font_size * 1.0, frame)
                # Remove paragraph spacing
                scribus.setParagraphGap(0, frame)
                scribus.deselectAll()
            except:
                pass

            # Step 1: Ensure no overflow first with minimal expansion
            scribus.layoutText(frame)

            # Expand minimally to ensure all text is visible - THIS IS THE KEY
            overflow_iterations = 0
            max_overflow_iterations = 100  # Safety limit
            while scribus.textOverflows(frame) and overflow_iterations < max_overflow_iterations:
                current_w, current_h = scribus.getSize(frame)
                scribus.sizeObject(current_w, current_h + 3, frame)
                scribus.layoutText(frame)
                overflow_iterations += 1

            # Step 2: Calculate exact height using official Scribus methods
            try:
                # Force fixed line spacing for accurate calculation
                scribus.setLineSpacing(font_size * 1.0, frame)
                scribus.layoutText(frame)

                # Get actual text metrics
                num_lines = scribus.getTextLines(frame)
                line_spacing = scribus.getLineSpacing(frame)
                left, right, top, bottom = scribus.getTextDistances(frame)

                if num_lines > 0:
                    # Calculate exact height: (lines × spacing) + padding
                    exact_text_height = num_lines * line_spacing
                    exact_frame_height = exact_text_height + top + bottom

                    # Resize to exact height
                    current_w, current_h = scribus.getSize(frame)
                    scribus.sizeObject(current_w, exact_frame_height, frame)
                    scribus.layoutText(frame)

                    # Verify no overflow after exact sizing
                    if scribus.textOverflows(frame):
                        # Add minimal space if needed
                        scribus.sizeObject(current_w, exact_frame_height + line_spacing * 0.1, frame)
                        scribus.layoutText(frame)

            except:
                # If official method fails, minimal fallback
                pass

        except:
            # Basic fallback overflow handling
            overflow_count = 0
            while scribus.textOverflows(frame) and overflow_count < 15:
                try:
                    current_w, current_h = scribus.getSize(frame)
                    scribus.sizeObject(current_w, current_h + 2, frame)
                    overflow_count += 1
                except:
                    break

    # Update y_offset to position after this frame
    # Use hierarchical spacing if provided, otherwise use standard spacing logic
    if custom_spacing is not None:
        spacing = custom_spacing
    else:
        spacing = 0 if no_bottom_gap else BLOCK_SPACING

    try:
        if balanced_columns and len(frames_to_setup) > 1:
            # For balanced columns, use the maximum height of both frames after overflow handling
            left_frame = frames_to_setup[0][0]
            right_frame = frames_to_setup[1][0]

            left_frame_height = scribus.getSize(left_frame)[1]
            right_frame_height = scribus.getSize(right_frame)[1]
            max_frame_height = max(left_frame_height, right_frame_height)

            # IMPORTANT: Resize both columns to the same height for visual balance
            left_frame_width = scribus.getSize(left_frame)[0]
            right_frame_width = scribus.getSize(right_frame)[0]
            scribus.sizeObject(left_frame_width, max_frame_height, left_frame)
            scribus.sizeObject(right_frame_width, max_frame_height, right_frame)

            # CRITICAL FIX: Final overflow check after resizing to same height
            # Check BOTH columns one more time and expand if needed
            for col_frame in [left_frame, right_frame]:
                overflow_count = 0
                # Use larger increments for faster convergence
                while scribus.textOverflows(col_frame) and overflow_count < 200:
                    col_w, col_h = scribus.getSize(col_frame)
                    # Use adaptive expansion - larger increments initially, then smaller
                    if overflow_count < 50:
                        increment = 10  # Fast expansion initially
                    elif overflow_count < 100:
                        increment = 5   # Medium expansion
                    else:
                        increment = 2   # Fine-tuning at the end
                    scribus.sizeObject(col_w, col_h + increment, col_frame)
                    scribus.layoutText(col_frame)
                    overflow_count += 1

                    # Force redraw every 20 iterations for accurate detection
                    if overflow_count % 20 == 0:
                        try:
                            scribus.redrawAll()
                        except:
                            pass

            # Recalculate max height after final overflow fixes
            left_frame_height = scribus.getSize(left_frame)[1]
            right_frame_height = scribus.getSize(right_frame)[1]
            max_frame_height = max(left_frame_height, right_frame_height)

            # Add safety buffer ONLY for templates with roadsigns
            # Roadsigns take up space on the right side, so text needs extra vertical space
            if in_template and has_roadsigns:
                # Templates with roadsigns need extra space to prevent text overflow
                safety_buffer = 40  # Extra space for roadsign displacement
            else:
                # No roadsigns or not template text - minimal/no buffer
                safety_buffer = 0

            max_frame_height += safety_buffer

            # Resize both to match the final max height with buffer
            scribus.sizeObject(left_frame_width, max_frame_height, left_frame)
            scribus.sizeObject(right_frame_width, max_frame_height, right_frame)

            frame_y = scribus.getPosition(frame)[1]
            y_offset = frame_y + max_frame_height + spacing

            # Log balanced columns
            frame_x = scribus.getPosition(frame)[1]
            current_page = scribus.currentPage()
            frame_type = "TEMPLATE_TEXT" if in_template else "DESC_TEXT"
            log_position(f"{frame_type}_BALANCED", "frame", frame_x, frame_y, frame_w, max_frame_height,
                        current_page, f"Balanced cols: L={left_frame_height:.1f}, R={right_frame_height:.1f}, Max={max_frame_height:.1f}")
        else:
            # Standard single frame
            frame_height = scribus.getSize(frame)[1]
            frame_y = scribus.getPosition(frame)[1]
            frame_x = scribus.getPosition(frame)[0]
            y_offset = frame_y + frame_height + spacing

            # Log text frame position
            current_page = scribus.currentPage()
            frame_type = "TEMPLATE_TEXT" if in_template else "DESC_TEXT"
            overflow = scribus.textOverflows(frame)
            log_position(frame_type, "frame", frame_x, frame_y, frame_w, frame_height,
                        current_page, f"Overflow={overflow}, spacing={spacing}")
    except:
        y_offset += text_h + spacing

    return frame

# Helper function to determine if a color is dark or light
def is_dark_color(color_name):
    """Determine if a named color is dark (needing white text) or light (needing black text)"""
    # Known dark colors that need white text
    dark_colors = ["Black", "Blue", "Red", "DarkRed", "Green", "DarkGreen", 
                  "DarkBlue", "Purple", "Magenta", "DarkGrey", "Brown"]
                  
    # Known bright/light colors that need black text
    light_colors = ["White", "Yellow", "Cyan", "LightGrey", "Lime", "Orange", "Pink"]
    
    # If it's a known dark color
    if color_name in dark_colors:
        return True
        
    # If it's a known light color  
    if color_name in light_colors:
        return False
        
    # Add special case for the special bright green in the template
    if color_name in BACKGROUND_COLORS and "Green" in color_name:
        # For the bright green templates, use black text
        return False
    
    # Default to assuming it's dark enough for white text
    return True

# ───────── TEXT + IMAGE COMBINED LAYOUT ─────────
def place_roadsigns_on_right(roadsign_list, base_path, text_start_y=None):
    """Places road signs horizontally on the right side with bounding boxes and text flow."""
    global y_offset

    if not roadsign_list:
        return

    # Road sign dimensions - compact sizing
    sign_width = 30  # Road sign width (reduced from 35)
    sign_height = 30  # Road sign height (reduced from 35)
    sign_spacing = 2  # Spacing between multiple signs (reduced from 4)
    box_padding = 1  # Padding inside bounding box (reduced from 2)

    # Calculate total width needed for all road signs
    valid_signs = [rs for rs in roadsign_list if rs]
    num_signs = len(valid_signs)

    # DEBUG: Commented out for performance
    # print(f"DEBUG place_roadsigns_on_right: {num_signs} valid signs, text_start_y={text_start_y}, y_offset={y_offset}")

    if num_signs == 0:
        # DEBUG: Commented out for performance
        # print("DEBUG: No valid signs, returning")
        return

    # Calculate layout for road signs on the right
    total_signs_width = (num_signs * sign_width) + ((num_signs - 1) * sign_spacing) + (box_padding * 2)
    signs_area_x = PAGE_WIDTH - MARGINS[2] - total_signs_width  # Right-aligned
    signs_height = sign_height + (box_padding * 2)

    # Use provided text start position, or current y_offset if not provided
    # Position road signs at the same level as the template text started
    block_y = text_start_y if text_start_y is not None else y_offset

    # Check space for road signs
    force_new_page_if_needed(signs_height)

    # CRITICAL: After potential new page, update block_y to use current y_offset
    # This prevents roadsigns from being placed at old positions from previous page
    if text_start_y is not None:
        # If we forced a new page, use top of new page instead of old text_start_y
        safe_boundary = PAGE_HEIGHT - MARGINS[3] - 40
        if block_y > safe_boundary or block_y < MARGINS[1]:
            # Either beyond boundary or before top margin (wrapped around) - use current position
            # DEBUG: Commented out for performance
            # print(f"DEBUG: Adjusting block_y from {block_y} to {y_offset} (outside valid range)")
            block_y = y_offset

    # Draw overall bounding box for all road signs with text flow properties
    signs_bounding_box = scribus.createRect(signs_area_x, block_y, total_signs_width, signs_height)
    scribus.setLineColor("None", signs_bounding_box)  # Invisible border
    scribus.setLineWidth(0.5, signs_bounding_box)
    scribus.setFillColor("None", signs_bounding_box)

    # Apply text flow properties to make template text flow around the road signs
    try:
        # Set the text flow mode so text wraps around this shape
        scribus.setTextFlowMode(signs_bounding_box, 1)  # 1 = text flows around frame
        scribus.setTextFlowUsesFrame(signs_bounding_box, True)  # Use frame boundaries
        scribus.setTextFlowUsesBoundingBox(signs_bounding_box, True)  # Use bounding box for flow
    except Exception as e:
        # Fallback if text flow functions are not available
        pass

    # Place road signs horizontally inside the bounding box
    current_sign_x = signs_area_x + box_padding
    for roadsign in valid_signs:
        try:
            img_path = os.path.join(base_path, roadsign)
            if os.path.exists(img_path):
                sign_frame = scribus.createImage(
                    current_sign_x,
                    block_y + box_padding,
                    sign_width,
                    sign_height
                )
                scribus.loadImage(img_path, sign_frame)
                scribus.setScaleImageToFrame(True, True, sign_frame)
                scribus.setImageScaleMode(1, sign_frame)  # Scale to frame proportionally
        except Exception as e:
            print(f"Could not load roadsign: {roadsign}")

        # Move to next sign position
        current_sign_x += sign_width + sign_spacing

    return


def place_wrapped_text_and_images(text_arr, image_list, base_path,
                                default_font_size=6, in_template=False, alignment="C", is_continuation=False, font_family=None, roadsigns=None):
    """Places text with images integrated inside the same container with text flow."""
    global y_offset

    # Handle roadsigns parameter
    rs = roadsigns if roadsigns else []

    text_arr = [str(t) for t in (text_arr or [])]

    # Clean text items
    cleaned_text_arr = []
    for t in text_arr:
        t = handle_superscripts(t)
        t = t.strip()
        # Check if text is empty after stripping HTML tags (to filter out <p><br></p> etc)
        test_text = strip_markdown_formatting(t).strip()
        if test_text:  # Only add if there's actual text content
            cleaned_text_arr.append(t)

    # Store image information for integrated placement with text
    # Only prepare images if this is template text
    # Changed to handle multiple images in a column
    image_info_list = []
    if image_list and in_template:
        # Calculate column width for image sizing
        frame_w = PAGE_WIDTH - MARGINS[0] - MARGINS[2]
        col_width = (frame_w - COLUMN_GAP) / 2 if USE_TWO_COLUMN_LAYOUT else frame_w

        # Prepare ALL images for stacked column layout
        for image in image_list:
            if image:
                img_path = os.path.join(base_path, image)
                if os.path.exists(img_path):
                    # Use original image sizing logic - standard height with dynamic width
                    standard_image_height = DEFAULT_IMAGE_SIZE  # Use configurable image size from config

                    # Get original image dimensions for aspect ratio calculation
                    if Image:
                        try:
                            with Image.open(img_path) as im:
                                orig_w, orig_h = im.size
                        except:
                            orig_w, orig_h = (300, 200)
                    else:
                        orig_w, orig_h = (300, 200)

                    # Calculate width maintaining aspect ratio with the standard height
                    scale = float(standard_image_height) / float(orig_h) if orig_h else 1.0
                    img_width = orig_w * scale
                    img_height = standard_image_height

                    # Ensure image fits within column width (with padding)
                    max_img_width = col_width * 0.9  # Use 90% of column width
                    if img_width > max_img_width:
                        # Resize to fit column
                        scale_factor = max_img_width / img_width
                        img_width = max_img_width
                        img_height = img_height * scale_factor

                    image_info_list.append({
                        'path': img_path,
                        'width': img_width,
                        'height': img_height
                    })
                    # Continue to process all images (removed break)

    # Place text and integrate image if available
    if cleaned_text_arr:
        # Store the starting Y position BEFORE creating anything
        saved_y_offset = y_offset
        img_bounding_box = None
        img_frame = None

        # Prepare text first
        if len(cleaned_text_arr) == 1:
            text = cleaned_text_arr[0]
        else:
            text = "\n".join(cleaned_text_arr)
        text = text.rstrip()

        # Create text frame FIRST (so it's at the back)
        frame_start_y = y_offset

        # If this is a continuation template (same ID), suppress bottom gap to eliminate spacing
        # Use balanced_columns=True for even text distribution in both columns
        # Pass roadsigns info so text frames can reserve extra space if needed
        # DEBUG: Commented out for performance
        # print(f"DEBUG: Creating template text with balanced_columns=True, has_roadsigns={len(rs) > 0}")
        frame = place_text_block_flow(text, default_font_size, False, in_template, no_bottom_gap=is_continuation, balanced_columns=True, font_family=font_family, has_roadsigns=(len(rs) > 0))

        # If we have images AND this is template text, extend frame to prevent text overlap
        # NOTE: With balanced columns, 'frame' is only the LEFT frame, but we need to check BOTH frames for overflow
        if image_info_list and in_template and frame:
            try:
                # Calculate total height for all images stacked vertically
                total_img_height = 0
                max_img_width = 0
                image_gap = 1  # Gap between stacked images
                
                for img_info in image_info_list:
                    total_img_height += img_info['height'] + image_gap
                    max_img_width = max(max_img_width, img_info['width'])
                
                # Remove last gap
                if total_img_height > 0:
                    total_img_height -= image_gap

                # Calculate how much extra height is needed for text displacement
                frame_w = PAGE_WIDTH - MARGINS[0] - MARGINS[2]

                # When images take up width in left column, text flows around them
                # Don't extend frame - let it size naturally based on text
                # The images will have text flow enabled, so text wraps around them
                col_width = (frame_w - COLUMN_GAP) / 2 if USE_TWO_COLUMN_LAYOUT else frame_w
                
                # Don't artificially extend the frame - let text flow naturally
                # Just ensure frame height is at least as tall as the image stack
                current_w, current_h = scribus.getSize(frame)
                min_height = max(current_h, total_img_height + 10)  # Just ensure minimum height
                scribus.sizeObject(current_w, min_height, frame)

                # CRITICAL: Update y_offset to account for the extended frame
                # Always use the new extended height since we have images
                scribus.layoutText(frame)

                # CRITICAL FIX: For balanced columns with images, check and fix overflow in RIGHT column
                # The images are on the left, but text can overflow on the right
                try:
                    # Get all text frames on current page that might be the right column
                    all_frames = scribus.getAllObjects()
                    current_page = scribus.currentPage()
                    frame_y = scribus.getPosition(frame)[1]

                    # Find potential right column frame (same Y position, different X)
                    frame_x = scribus.getPosition(frame)[0]
                    for obj_name in all_frames:
                        try:
                            if scribus.getObjectType(obj_name) == "TextFrame":
                                obj_x, obj_y = scribus.getPosition(obj_name)
                                # Check if this is roughly at the same Y but different X (right column)
                                if abs(obj_y - frame_y) < 5 and obj_x > frame_x + 50:  # Right of main frame
                                    # This is the right column - check for overflow with aggressive expansion
                                    overflow_count = 0
                                    while scribus.textOverflows(obj_name) and overflow_count < 100:
                                        obj_w, obj_h = scribus.getSize(obj_name)
                                        scribus.sizeObject(obj_w, obj_h + 10, obj_name)  # Increased from 5 to 10
                                        scribus.layoutText(obj_name)
                                        overflow_count += 1
                        except:
                            continue
                except:
                    pass

                # IMPORTANT: Get actual frame Y position and height after layout
                actual_frame_y = scribus.getPosition(frame)[1]
                actual_frame_h = scribus.getSize(frame)[1]

                # Calculate the actual bottom of content (text or images, whichever is lower)
                text_bottom = actual_frame_y + actual_frame_h
                images_bottom = actual_frame_y + total_img_height
                
                # Use the lower of the two (text or images)
                content_bottom = max(text_bottom, images_bottom)
                y_offset = content_bottom + 2  # Minimal spacing after content
            except:
                pass

        # Then place images on top if available - use ACTUAL FRAME position
        # Only place images if this is template text
        # Stack multiple images vertically in the left column
        if image_info_list and in_template:
            box_padding = 1
            image_gap = 1  # Vertical gap between stacked images
            
            # Get starting Y position from frame
            if frame:
                current_y = scribus.getPosition(frame)[1]  # Get actual Y position of text frame
            else:
                current_y = saved_y_offset  # Fallback to saved position if no frame

            # Track if we created any new pages during image placement
            first_image_page = scribus.currentPage()

            # Stack images vertically
            for img_idx, image_info in enumerate(image_info_list):
                img_width = image_info['width']
                img_height = image_info['height']

                # Calculate image container with padding
                container_width = img_width + (box_padding * 2)
                container_height = img_height + (box_padding * 2)

                # PAGINATION CHECK: Check if image fits on current page
                safe_boundary = PAGE_HEIGHT - MARGINS[3] - 40  # Page number buffer
                if current_y + container_height > safe_boundary:
                    # Image won't fit on current page - create new page
                    new_page()
                    # Reset current_y to top margin of new page
                    current_y = MARGINS[1]
                    # Update y_offset to match
                    y_offset = current_y
                    
                    # If this is not the first image, we need to handle text continuation
                    # The text frame on previous page needs to be finalized
                    if img_idx > 0 and frame:
                        try:
                            # Ensure the previous page's text frame is properly laid out
                            scribus.gotoPage(scribus.currentPage() - 1)
                            scribus.layoutText(frame)
                            scribus.gotoPage(scribus.currentPage() + 1)
                        except:
                            pass

                # Position at left margin for column placement
                img_x = MARGINS[0]

                # Create bounding box container for the image
                img_bounding_box = scribus.createRect(img_x, current_y, container_width, container_height)
                scribus.setLineColor("None", img_bounding_box)  # Invisible border
                scribus.setLineWidth(0.5, img_bounding_box)
                scribus.setFillColor("None", img_bounding_box)

                # Enable text flow around the bounding box - mode 2 = around bounding box
                try:
                    scribus.setTextFlowMode(img_bounding_box, 2)
                except:
                    pass

                # Create and place the actual image inside the bounding box
                img_frame = scribus.createImage(
                    img_x + box_padding,
                    current_y + box_padding,
                    img_width,
                    img_height
                )
                scribus.loadImage(image_info['path'], img_frame)
                scribus.setScaleImageToFrame(True, True, img_frame)
                scribus.setLineColor("None", img_frame)  # Invisible border for image

                # Enable text flow on the image frame itself as well - mode 1 = around frame shape
                try:
                    scribus.setTextFlowMode(img_frame, 1)
                except:
                    pass

                # Make frame fit perfectly
                try:
                    scribus.setScaleFrameToImage(img_frame)
                    # Get the actual frame size after scaling
                    actual_width, actual_height = scribus.getSize(img_frame)
                    actual_x, actual_y = scribus.getPosition(img_frame)

                    # Add proportional padding around the image for orange box
                    # Padding scales with image size - larger images get more padding
                    padding = max(2, min(actual_width, actual_height) * 0.01)  # 1% of smaller dimension, minimum 2 points
                    box_width = actual_width + (padding * 2)
                    box_height = actual_height + (padding * 2)
                    box_x = actual_x - padding
                    box_y = actual_y - padding

                    # Resize orange bounding box with padding
                    scribus.sizeObject(box_width, box_height, img_bounding_box)
                    scribus.moveObject(box_x, box_y, img_bounding_box)

                    # Ensure orange box is visible by bringing it to front
                    scribus.moveObjectAbs(box_x, box_y, img_bounding_box)

                    # Re-enable text flow after moving/resizing
                    try:
                        scribus.setTextFlowMode(img_bounding_box, 2)  # around bounding box
                        scribus.setTextFlowMode(img_frame, 1)  # around frame shape
                    except:
                        pass
                    
                    # Move to next image position (stack vertically)
                    current_y = box_y + box_height + image_gap
                except:
                    # If scaling fails, use original dimensions and move down
                    current_y += container_height + image_gap
            
            # After placing all images, update y_offset to continue after the last image
            # This ensures subsequent content starts after the images
            if image_info_list:
                y_offset = max(y_offset, current_y)

        # Additional template-specific overflow safety check
        if in_template and frame:
            try:
                # Validate and fix template text frame
                frame = validate_template_text_frame(frame, text, default_font_size)
            except:
                pass

    return

# Removed duplicate/broken functions - using working versions below

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── QUIZ PLACEMENT FUNCTIONS ────────────
def strip_markdown_formatting(text):
    """Remove all HTML tags from a string."""
    if not text or not isinstance(text, str):
        return text
    return re.sub(r'<[^>]+>', '', text)

def place_quiz(arr, in_template=True, group_image=None, base_path=None):
    global y_offset, CURRENT_COLOR
    global quiz_heading_placed_on_page

    # Simple quiz placement without column manager - like dopy.py

    if not PRINT_QUIZZES:
        return
    if not arr or not isinstance(arr, list):
        return
    # Filter quiz items based on QUIZ_FILTER_MODE
    filtered_arr = []
    for qa in arr:
        if not isinstance(qa, dict):
            continue
        # Remove HTML tags from question and answer
        if 'que' in qa:
            qa['que'] = strip_markdown_formatting(qa['que'])
        if 'ans' in qa:
            qa['ans'] = strip_markdown_formatting(qa['ans'])
        # Get the answer text to check if it's true (V) or false (F)
        answer_text = qa.get('ans', '').strip()
        is_true = qa.get('is_true', None)
        if is_true is None:
            if answer_text.endswith(QUIZ_TRUE_TEXT) or f'A: {QUIZ_TRUE_TEXT}' in answer_text:
                is_true = True
            elif answer_text.endswith(QUIZ_FALSE_TEXT) or f'A: {QUIZ_FALSE_TEXT}' in answer_text:
                is_true = False
            else:
                is_true = True
        if QUIZ_FILTER_MODE == "true_only" and not is_true:
            continue
        elif QUIZ_FILTER_MODE == "false_only" and is_true:
            continue
        qa['is_true'] = is_true
        filtered_arr.append(qa)
    if not filtered_arr:
        return
    # Copy 6 layout style from quiz_from_csv.py
    quiz_width = PAGE_WIDTH - MARGINS[0] - MARGINS[2]
    header_height = 24  # Reduced blue header height (matching dopy.py)
    row_height = 16  # Increased row height for larger 9pt font (matching dopy.py)
    answer_box_width = 18  # V/F box width
    answer_box_height = 16  # V/F box height increased for 10pt font (matching dopy.py)
    answer_box_gap = 3  # Gap between V and F boxes

    # CRITICAL: Calculate total height needed for entire quiz upfront
    # This prevents quiz rows from exceeding page boundaries
    def calculate_quiz_height(quiz_array):
        """Calculate total height needed for quiz header and all rows"""
        total = header_height  # Start with header
        text_width = quiz_width - 42
        for qa in quiz_array:
            question = qa.get('que', '')
            formatted = re.sub(r'<[^>]+>', '', question)
            display = re.sub(r'cm(\d)', lambda m: 'cm' + {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'}[m.group(1)], formatted)
            # Estimate row height conservatively
            chars_per_line = text_width / (8 * 0.42)
            if len(display) <= chars_per_line * 0.85:
                row_h = row_height
            else:
                lines = int(len(display) / chars_per_line) + 1
                row_h = max(row_height, lines * 8 * 1.2 * 1.1)
            total += row_h
        return total

    # Smart quiz pagination: only move to new page if not enough space for header + minimum rows
    # This allows quizzes to start on current page and naturally flow to next page
    safe_boundary = PAGE_HEIGHT - MARGINS[3] - 40
    available_space = safe_boundary - y_offset

    # Minimum space needed: header (24pt) + at least 2-3 rows (16pt each) = ~70pt
    min_quiz_start_space = header_height + (row_height * 2)  # Header + 2 rows minimum

    # Log quiz start position
    current_page = scribus.currentPage()
    log_position("QUIZ_START", "quiz", MARGINS[0], y_offset, quiz_width, 0, current_page,
                f"Available: {available_space:.1f}pt, min needed: {min_quiz_start_space:.1f}pt for {len(filtered_arr)} rows")

    # Only move to new page if we can't fit even the header + 2 rows
    if available_space < min_quiz_start_space:
        log_position("QUIZ_NEW_PAGE", "quiz", 0, 0, 0, 0, current_page,
                    f"Insufficient space ({available_space:.1f}pt < {min_quiz_start_space:.1f}pt), moving to new page")
        new_page()
    # Otherwise, start quiz on current page and let rows flow naturally to next page when needed

    # Draw blue header like copy 6 (from quiz_from_csv.py)
    if not quiz_heading_placed_on_page:
        # Define colors if not exists
        try:
            scribus.defineColor("Cyan", 0, 160, 224)  # Blue color
            scribus.defineColor("Yellow", 255, 255, 0)
            scribus.defineColor("NumBoxBlue", 210, 235, 255)
        except:
            pass

        # Create blue header background (full width - no gap for yellow box)
        header_bg = scribus.createRect(MARGINS[0], y_offset, quiz_width, header_height)
        scribus.setFillColor("Cyan", header_bg)
        scribus.setLineColor("Cyan", header_bg)

        # Quiz header text (full width with padding)
        header_text = scribus.createText(MARGINS[0] + 3, y_offset + 3, quiz_width - 6, header_height - 6)
        scribus.setText("Quiz", header_text)
        try:
            scribus.setFont(DEFAULT_FONT, header_text)
        except:
            try:
                available_fonts = scribus.getFontNames()
                if available_fonts:
                    scribus.setFont(available_fonts[0], header_text)
            except:
                pass
        scribus.setFontSize(10, header_text)
        scribus.setTextColor("White", header_text)
        scribus.setTextAlignment(0, header_text)

        # Configure header text wrapping and overflow control like copy6.py
        try:
            scribus.setTextDistances(1, 1, 1, 1, header_text)  # Minimal padding
            scribus.setLineSpacing(10, header_text)  # Adjusted line spacing for 10pt font

            # Force text to fit within frame bounds - prevent header overflow
            try:
                scribus.setTextBehaviour(header_text, 0)  # Force text to stay in frame
            except:
                try:
                    scribus.setTextToFrameOverflow(header_text, False)  # Disable overflow
                except:
                    pass

            # Enable text wrapping for headers
            try:
                scribus.setTextFlowMode(header_text, 0)  # Enable text flow
            except:
                pass

        except:
            pass

        # Yellow page number box removed per user request

        quiz_heading_placed_on_page = True
        y_offset += header_height  # No gap after header - like dopy.py
        enforce_margin_boundary()
    # Add the exact text overflow function from quiz_from_csv.py
    def check_text_overflow(text, width, font_size, default_height, is_header=False):
        """Check if text will overflow and calculate required height if needed"""
        # More conservative estimation to ensure no overflow
        if is_header:
            # Headers: more conservative estimate
            chars_per_line = width / (font_size * 0.45)
        else:
            # Regular text: use copy6.py original values
            chars_per_line = width / (font_size * 0.42)  # Copy6.py original value

        # Check if text truly needs multiple lines
        # Use copy6.py original threshold
        if len(text) <= chars_per_line * 0.85:  # Copy6.py original 85% threshold
            return default_height  # Single line - use compact height

        # Calculate actual lines needed
        lines_needed = int(len(text) / chars_per_line) + 1

        # Special handling for borderline cases (text near one line) - safe
        if len(text) > chars_per_line * 0.85 and len(text) <= chars_per_line * 1.2:
            # Text is close to or slightly over one line - safe version
            return default_height * 1.3  # Safe 30% more height to prevent overflow

        # Calculate required height for true multi-line content - safe spacing
        line_height = font_size * 1.2  # Safe line height to prevent overflow
        calculated_height = lines_needed * line_height + 3  # Safe padding for line spacing

        # Add buffer for safety - prevent overflow
        calculated_height = calculated_height * 1.1  # 10% buffer to prevent overflow

        # Return calculated height
        return max(default_height, calculated_height)

    # Process each question as a table row (from quiz_from_csv.py)
    for idx, qa in enumerate(filtered_arr):
        question = qa.get('que', '')
        is_true = qa.get('is_true', False)

        # Apply superscript conversion BEFORE removing HTML tags
        question = handle_superscripts(question)

        # Keep it very simple - just remove HTML tags and show plain text
        formatted_question = re.sub(r'<[^>]+>', '', question)
        cleaned_question_for_calc = formatted_question

        # Apply superscript conversion for height calculation too
        display_question_for_calc = re.sub(r'cm(\d)', lambda m: 'cm' + {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'}[m.group(1)], cleaned_question_for_calc)

        # Calculate row height using corrected text width (matches the fixed positioning)
        text_width = quiz_width - 42  # Adjusted to match the new text box positioning
        current_row_height = check_text_overflow(display_question_for_calc, text_width, 8, row_height, is_header=False)

        # Check if this row would exceed boundary - if so, create new page
        safe_boundary = PAGE_HEIGHT - MARGINS[3] - 40
        row_bottom = y_offset + current_row_height

        if row_bottom > safe_boundary:
            # This row won't fit - create new page and continue placing rows
            current_page = scribus.currentPage()
            log_position("QUIZ_ROW_NEWPAGE", f"row_{idx}", MARGINS[0], y_offset, 0, 0, current_page,
                        f"Row {idx} needs {current_row_height:.1f}pt but only {safe_boundary - y_offset:.1f}pt left, creating new page")
            new_page()
            # After new page, recalculate row bottom
            row_bottom = y_offset + current_row_height

        # IMPORTANT: Use current y_offset AFTER boundary check (not before)
        # This ensures we use the correct position if a new page was created
        current_quiz_y = y_offset

        # Draw answer row exactly like copy 6 (from quiz_from_csv.py)
        # Remove number box background - no numbering needed
        # num_box_bg = scribus.createRect(MARGINS[0] + 2, y_offset, 12, current_row_height - 1)
        # scribus.setFillColor("NumBoxBlue", num_box_bg)
        # scribus.setLineColor("Cyan", num_box_bg)
        # scribus.setLineWidth(0.5, num_box_bg)

        # Remove the number box completely - no numbering needed
        # num_box_height = current_row_height - 2
        # text_y_offset = 1
        # num_box = scribus.createText(MARGINS[0] + 2, y_offset + text_y_offset, 12, num_box_height)

        # Answer text box - ensure it stays within margins
        text_start_x = MARGINS[0] + 2
        text_end_x = MARGINS[0] + quiz_width - 40  # Leave space for V/F boxes (38 + 2 margin)
        text_width = text_end_x - text_start_x
        text_box_bg = scribus.createRect(text_start_x, current_quiz_y, text_width, current_row_height - 1)

        # Alternate row colors like copy6.py
        if idx % 2 == 0:
            scribus.setFillColor("White", text_box_bg)
        else:
            try:
                scribus.defineColor("VeryLightCyan", 245, 252, 255)
                scribus.setFillColor("VeryLightCyan", text_box_bg)
            except:
                scribus.setFillColor("LightGray", text_box_bg)
        scribus.setLineColor("Cyan", text_box_bg)
        scribus.setLineWidth(0.5, text_box_bg)

        # Question text frame - adjusted positioning to remove numbering space
        q_frame = scribus.createText(text_start_x + 2, current_quiz_y + 1, text_width - 4, current_row_height - 2)
        scribus.setText(formatted_question, q_frame)
        try:
            scribus.setFont(QUIZ_ACTUAL_FONT, q_frame)
        except:
            try:
                scribus.setFont(DEFAULT_FONT, q_frame)
            except:
                # If both fail, try first available font
                try:
                    available_fonts = scribus.getFontNames()
                    if available_fonts:
                        scribus.setFont(available_fonts[0], q_frame)
                except:
                    pass
        # Set the correct font size to match actual quiz text (matching dopy.py)
        try:
            scribus.setFontSize(9, q_frame)
        except:
            pass
        scribus.setTextColor("Black", q_frame)

        # Enable proper text alignment and centering like dopy.py
        try:
            scribus.setTextDistances(0, 0, 0, 0, q_frame)  # No padding for perfect centering like V/F boxes
            # Set line spacing based on row height - increased for better readability
            if current_row_height > 14:
                scribus.setLineSpacing(9, q_frame)  # Increased spacing for multi-line
            else:
                scribus.setLineSpacing(8, q_frame)  # Increased spacing for single line

            # Set horizontal alignment
            scribus.setTextAlignment(0, q_frame)  # Left align

            # Try to set vertical alignment to middle like dopy.py
            try:
                scribus.setTextVerticalAlignment(1, q_frame)  # 1 = middle alignment
            except:
                try:
                    # Alternative method for vertical centering from dopy.py
                    scribus.setTextBehaviour(q_frame, 1)  # Try different behavior
                except:
                    pass

            # Force text to stay within bounds - comprehensive dopy.py approach
            try:
                scribus.setTextBehaviour(q_frame, 0)  # Force text in frame
            except:
                pass

            # Additional overflow protection from dopy.py
            try:
                scribus.setTextToFrameOverflow(q_frame, False)  # Disable overflow
            except:
                pass

            # Enable text wrapping like dopy.py
            try:
                scribus.setTextFlowMode(q_frame, 0)  # Enable text flow
            except:
                pass
        except:
            pass

        # Define checkbox colors if they don't exist
        try:
            scribus.defineColor("CheckBoxColor", 240, 255, 240)  # Light green tint for V
            scribus.defineColor("CheckBoxColor2", 255, 240, 240)  # Light red tint for F
        except:
            pass

        # V checkbox box - adjusted position since no number box
        v_box_bg = scribus.createRect(MARGINS[0] + quiz_width - 38, current_quiz_y, 18, current_row_height - 1)
        try:
            scribus.setFillColor("CheckBoxColor", v_box_bg)
        except:
            scribus.setFillColor("White", v_box_bg)
        scribus.setLineColor("Cyan", v_box_bg)
        scribus.setLineWidth(0.5, v_box_bg)

        # V checkbox text - adjusted position
        checkbox_box_height = current_row_height - 2  # Use almost full row height with 1pt padding
        checkbox_y_offset = 1  # Minimal top padding
        v_box = scribus.createText(MARGINS[0] + quiz_width - 38, current_quiz_y + checkbox_y_offset, 18, checkbox_box_height)
        scribus.setText("V", v_box)
        try:
            scribus.setFont(QUIZ_ACTUAL_FONT, v_box)
        except:
            try:
                scribus.setFont(DEFAULT_FONT, v_box)
            except:
                try:
                    available_fonts = scribus.getFontNames()
                    if available_fonts:
                        scribus.setFont(available_fonts[0], v_box)
                except:
                    pass
        scribus.setFontSize(10, v_box)  # Increased V/F box font size (matching dopy.py)
        scribus.setTextAlignment(1, v_box)  # Center align horizontally

        # Try to set vertical alignment to middle like dopy.py
        try:
            scribus.setTextVerticalAlignment(1, v_box)  # 1 = middle alignment
        except:
            pass
        # Set proper text distances for centering
        scribus.setTextDistances(0, 0, 0, 0, v_box)  # No padding for perfect centering

        # Set V box color based on correctness - cyan if true answer
        if is_true:
            scribus.setTextColor("Cyan", v_box)
        else:
            scribus.setTextColor("Black", v_box)

        # F checkbox box - adjusted position
        f_box_bg = scribus.createRect(MARGINS[0] + quiz_width - 18, current_quiz_y, 18, current_row_height - 1)
        try:
            scribus.setFillColor("CheckBoxColor2", f_box_bg)
        except:
            scribus.setFillColor("White", f_box_bg)
        scribus.setLineColor("Cyan", f_box_bg)
        scribus.setLineWidth(0.5, f_box_bg)

        # F checkbox text - adjusted position
        f_box = scribus.createText(MARGINS[0] + quiz_width - 18, current_quiz_y + checkbox_y_offset, 18, checkbox_box_height)
        scribus.setText("F", f_box)
        try:
            scribus.setFont(QUIZ_ACTUAL_FONT, f_box)
        except:
            try:
                scribus.setFont(DEFAULT_FONT, f_box)
            except:
                try:
                    available_fonts = scribus.getFontNames()
                    if available_fonts:
                        scribus.setFont(available_fonts[0], f_box)
                except:
                    pass
        scribus.setFontSize(10, f_box)  # Increased V/F box font size (matching dopy.py)
        scribus.setTextAlignment(1, f_box)  # Center align horizontally

        # Try to set vertical alignment to middle like dopy.py
        try:
            scribus.setTextVerticalAlignment(1, f_box)  # 1 = middle alignment
        except:
            pass
        # Set proper text distances for centering
        scribus.setTextDistances(0, 0, 0, 0, f_box)  # No padding for perfect centering

        # Set F box color based on correctness - cyan if false answer
        if not is_true:
            scribus.setTextColor("Cyan", f_box)
        else:
            scribus.setTextColor("Black", f_box)

        # Simple boundary enforcement for quiz elements
        simple_constrain_element(q_frame)
        simple_constrain_element(text_box_bg)

        # Move to next row - simple approach like dopy.py
        # Log quiz row placement showing position where row was placed
        safe_boundary = PAGE_HEIGHT - MARGINS[3] - 40
        if current_quiz_y > safe_boundary - 50:  # Log when within 50pts of boundary
            current_page = scribus.currentPage()
            row_end = current_quiz_y + current_row_height
            log_position("QUIZ_ROW_PLACED", f"row_{idx}", MARGINS[0], current_quiz_y, quiz_width, current_row_height,
                        current_page, f"Row {idx} placed at Y:{current_quiz_y:.1f}, ends at {row_end:.1f}, boundary={safe_boundary:.1f}")

        y_offset += current_row_height

    # No additional spacing after quiz section - questions should be compact
    # No need to cap y_offset - pagination is now handled properly during row placement

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── OTHER FUNCTIONS ────────────
# ────────────────────────────────────────────────────────────────────────────────

def measure_quiz_group_height(group, group_image, base_path):
    # This function mimics the height calculation logic in place_quiz, but does not place anything.
    global y_offset  # Declare global variable
    quiz_header_height = 22  # Match the safe header height
    card_spacing = 1  # Minimal spacing between cards
    answer_box_width = 25
    answer_box_height = 16
    quiz_width = PAGE_WIDTH - MARGINS[0] - MARGINS[2]
    answer_box_gap = 4
    quiz_bar_height = 14  # Match the safe row height
    padding_top = 1  # Minimal safe padding
    padding_bottom = 1  # Minimal safe padding
    image_height = 28  # Safe image height
    img_margin = 2
    image_gap = 2
    img_w = 0
    img_h = 0
    img_path = None
    if group_image:
        img_path = os.path.join(base_path, group_image) if base_path else group_image
        orig_w, orig_h = get_image_size(img_path)
        scale = float(image_height) / float(orig_h) if orig_h else 1.0
        img_w = orig_w * scale
        img_h = image_height
    answer_boxes_total_width = (2 * answer_box_width) + answer_box_gap
    question_width = quiz_width - (img_margin + (img_w + image_gap if img_path else 0)) - answer_boxes_total_width - 10
    question_heights = []
    for qa in group:
        question = qa.get('que', '')
        formatted_question = handle_superscripts(question)
        temp_frame = scribus.createText(0, 0, question_width, 40)
        scribus.setText(formatted_question, temp_frame)
        try:
            scribus.setFont(QUIZ_ACTUAL_FONT, temp_frame)
        except:
            try:
                scribus.setFont(DEFAULT_FONT, temp_frame)
            except:
                # If both fail, try first available font
                try:
                    available_fonts = scribus.getFontNames()
                    if available_fonts:
                        scribus.setFont(available_fonts[0], temp_frame)
                except:
                    pass
        # Set the correct font size to match actual quiz text
        try:
            scribus.setFontSize(8, temp_frame)
        except:
            pass
        scribus.setTextColor("Black", temp_frame)
        # Measure the required height
        try:
            required_height = max(scribus.getFrameText(temp_frame).count('\n') + 1, 1) * 9  # 9pt line height
        except:
            required_height = quiz_bar_height
        question_heights.append(max(required_height, quiz_bar_height))
        scribus.deleteObject(temp_frame)

    # Calculate total height
    total_question_height = sum(question_heights) + (len(question_heights) - 1) * card_spacing
    image_height_contribution = img_h if img_path else 0
    card_height = max(total_question_height, image_height_contribution) + padding_top + padding_bottom
    card_top_margin = 1

    # Add space for the heading (QUIZ bar)
    total_height = (quiz_bar_height + 1) + card_height + card_top_margin + 1
    return total_height

def process_template(tmpl, base_path, is_continuation=False, next_is_continuation=False):
    global y_offset, CURRENT_COLOR, global_template_count

    tid = tmpl.get("id","0")
    template_name = tmpl.get("text_md", [""])[0][:50] if tmpl.get("text_md") else f"Template {tid}"

    # Smart pagination: Check if there's enough space for template content
    # Use same pagination logic for all templates (including continuations under same module)
    safe_boundary = PAGE_HEIGHT - MARGINS[3] - 40  # Page number buffer
    available_space = safe_boundary - y_offset

    # Same threshold for all templates: aggressive but prevents bad placement
    min_template_space = 35  # header + 2 text lines minimum

    if available_space < min_template_space:
        current_page = scribus.currentPage()
        template_type = "continuation" if is_continuation else "new"
        log_position("TEMPLATE_NEW_PAGE", tid, 0, 0, 0, 0, current_page,
                    f"Before {template_type} template: insufficient space ({available_space:.1f}pt < {min_template_space}pt), y_offset={y_offset:.1f}")
        new_page()

    # Log template start position AFTER pagination check
    current_page = scribus.currentPage()
    log_position("TEMPLATE_START", tid, MARGINS[0], y_offset, 0, 0, current_page,
                f"Start: {template_name}")

    # Set color based on template ID (for text styling, but no background)
    tid = tmpl.get("id","0")
    try:
        idx = int(tid) % len(BACKGROUND_COLORS)
    except:
        idx = hash(tid) % len(BACKGROUND_COLORS)
    CURRENT_COLOR = BACKGROUND_COLORS[idx]
    
    # Get text content
    txt = tmpl.get("text_md", [])
    txt = txt if isinstance(txt, list) else ([txt] if txt else [])
    
    # Aggressively clean text items
    cleaned_txt = []
    for item in txt:
        if item:  # Skip None or empty items
            # Convert to string and strip all whitespace
            item_str = str(item).strip()
            if item_str:  # Only add non-empty strings
                # Apply superscript/subscript conversion
                item_str = handle_superscripts(item_str)
                cleaned_txt.append(item_str)
    
    # Get roadsigns
    rs = tmpl.get("roadsigns", [])
    rs = rs if isinstance(rs, list) else ([rs] if rs else [])
    
    # Determine alignment for images based on template ID for variety
    # Alternate between center, top, and bottom alignments
    alignments = ["C", "TC", "BC"]
    try:
        aid = int(tid) % len(alignments)
    except:
        aid = 0
    template_alignment = alignments[aid]
    
    # Get and normalize images early for integrated placement
    imgs = tmpl.get("images", [])
    imgs = imgs if isinstance(imgs, list) else ([imgs] if imgs else [])

    # Save y_offset and page number before placing template text (for roadsigns positioning)
    template_text_start_y = y_offset
    template_start_page = scribus.currentPage()

    # Place text and images together in integrated layout (like the correct.png example)
    if cleaned_txt:
        # Save position before text placement
        text_start_y = y_offset
        # Pass images to be integrated with text placement
        # Always suppress BLOCK_SPACING for templates - we handle spacing explicitly at template end
        # This prevents double-spacing (BLOCK_SPACING + template-to-template spacing)
        suppress_spacing = True  # Always suppress for templates

        # DEBUG: Commented out for performance
        # print(f"DEBUG: Using MODULE_TEXT_FONT_SIZE = {MODULE_TEXT_FONT_SIZE}")

        place_wrapped_text_and_images(cleaned_txt, imgs, base_path, MODULE_TEXT_FONT_SIZE, True, template_alignment, is_continuation=suppress_spacing, font_family=TEMPLATE_TEXT_FONT_FAMILY, roadsigns=rs)

        # Synchronize column_mgr with y_offset after template text placement
        # No forced minimum spacing - let templates use only the space they need
        column_mgr.set_current_y(y_offset)

    # If we have road signs, place them on the right side with bounding boxes
    if rs and len(rs) > 0:
        # DEBUG: Commented out for performance
        # template_name_short = tmpl.get("text_md", [""])[0][:30] if tmpl.get("text_md") else f"Template {tid}"
        # print(f"DEBUG: Placing {len(rs)} roadsigns for template: {template_name_short}...")

        # CRITICAL: Check if we're still on the same page as when template started
        current_page = scribus.currentPage()
        if current_page != template_start_page:
            # Template text caused a page break - place roadsigns on the NEW page at top
            # Use the current page's top margin position instead of old page position
            # DEBUG: Commented out for performance
            # print(f"DEBUG: Template crossed pages ({template_start_page} -> {current_page}), placing roadsigns at top of new page")
            # Use top of current page as roadsign position
            roadsign_y_position = MARGINS[1]  # Top margin of new page
            place_roadsigns_on_right(rs, base_path, roadsign_y_position)
        else:
            # Same page - safe to place roadsigns at original template start position
            place_roadsigns_on_right(rs, base_path, template_text_start_y)

        # Synchronize column_mgr with y_offset after roadsigns placement
        column_mgr.set_current_y(y_offset)

    # Place videos
    for v in tmpl.get("videos", []):
        # Make video text always visible
        video_text = f"[video: {v}]"
        place_text_block_flow(video_text, VIDEO_TEXT_FONT_SIZE, False, True)

    # Synchronize column_mgr with y_offset after videos placement
    if tmpl.get("videos", []):
        column_mgr.set_current_y(y_offset)

    # Images are now processed integrated with text above - no separate processing needed
    # Place quiz section using global constants - only if quizzes are enabled
    if "quiz" in tmpl and PRINT_QUIZZES:
        # Reset quiz header flag for this template - allow ONE header per template (matching dopy.py)
        global quiz_heading_placed_on_page
        quiz_heading_placed_on_page = False

        # CRITICAL: Sync y_offset from column_mgr to ensure we're past all template content
        # This prevents quiz from overlapping with template text/images
        y_offset = max(y_offset, column_mgr.get_current_y())

        # Add ultra-minimal spacing between template content and quiz section (seamless flow)
        y_offset += 1
        enforce_margin_boundary()

        # Get all quiz entries for this template and place under single header
        quiz_entries = tmpl.get("quiz", [])
        if quiz_entries:
            # Group quiz entries by image but place all under single header
            template_images = tmpl.get("images", [])
            if template_images:
                # Template has images - group quiz with first image
                place_quiz(quiz_entries, True, template_images[0], base_path)
            else:
                # Template has no images - group quiz entries together without image
                place_quiz(quiz_entries, True, None, base_path)

        # Synchronize column_mgr with y_offset after quiz placement
        column_mgr.set_current_y(y_offset)

    # Template end spacing - different for templates with/without quiz
    # Special case: if next template has same ID (continuation), use NO spacing
    safe_boundary = PAGE_HEIGHT - MARGINS[3] - 40

    # Cap y_offset if it exceeds boundary (prevents overlap)
    if y_offset > safe_boundary:
        y_offset = safe_boundary

    # Check if template has quiz to determine spacing
    has_quiz = "quiz" in tmpl and PRINT_QUIZZES and tmpl.get("quiz")

    if next_is_continuation:
        # Next template is continuation (same ID): 1px spacing for all
        spacing_to_add = 1
    elif has_quiz:
        # Templates with quiz: minimal spacing
        spacing_to_add = 2
    else:
        # Templates without quiz: no spacing (seamless)
        spacing_to_add = 0

    # Add the appropriate spacing
    y_offset += spacing_to_add

    # Final cap to ensure we don't exceed boundary
    if y_offset > safe_boundary:
        y_offset = safe_boundary

    global_template_count += 1

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── VERTICAL TOPIC BANNER FUNCTIONS ────────────
# ────────────────────────────────────────────────────────────────────────────────
def add_vertical_topic_banner(topic_name, banner_color=None):
    """
    Creates a vertical rectangle with topic text on the side of the page.

    Parameters:
    - topic_name: The name of the topic to display vertically
    - banner_color: Optional color name from JSON (e.g., "Cyan", "Blue", "Red")
    """
    global current_topic_text, current_topic_color

    # Store the current topic for use when creating new pages
    current_topic_text = topic_name

    # Use provided color if available, otherwise use hash-based color assignment
    if banner_color:
        # Use color from JSON
        current_topic_color = banner_color
    else:
        # Fallback: Assign a consistent color based on topic name hash
        # This ensures the same topic always gets the same color
        try:
            # Use hash of topic name to get a consistent color index
            color_idx = hash(topic_name) % len(BACKGROUND_COLORS)
            current_topic_color = BACKGROUND_COLORS[color_idx]
        except:
            # Fallback to first color if there's an issue
            current_topic_color = BACKGROUND_COLORS[0]

    # Add the vertical topic banner to the current page
    create_vertical_topic_banner()
    
def create_vertical_topic_banner():
    """Creates the vertical topic banner on the current page."""
    if not current_topic_text:
        return

    # Get current page number and determine side
    try:
        current_page = scribus.currentPage()
        is_odd_page = current_page % 2 == 1
    except:
        is_odd_page = True

    rect_height = PAGE_HEIGHT - MARGINS[1] - MARGINS[3]
    # Position banners with different offsets for left and right sides
    if is_odd_page:
        # Left side - banner positioned with smaller offset (closer to content)
        banner_x = MARGINS[0] - LEFT_BANNER_MARGIN_OFFSET - TOPIC_BANNER_WIDTH
    else:
        # Right side - banner positioned with larger offset (further from content)
        banner_x = PAGE_WIDTH - MARGINS[2] + RIGHT_BANNER_MARGIN_OFFSET
    banner_y = MARGINS[1]

    # Create rectangle for the banner background
    rect = scribus.createRect(banner_x, banner_y, TOPIC_BANNER_WIDTH, rect_height)
    if current_topic_color:
        scribus.setFillColor(current_topic_color, rect)
    scribus.setLineColor("None", rect)

    # Prepare and place banner text
    display_text = " ".join(current_topic_text.upper())
    banner_center_y = banner_y + (rect_height / 2)
    text_height = rect_height * BANNER_TEXT_HEIGHT_PERCENT

    try:
        text_frame = scribus.createText(0, 0, text_height, TOPIC_BANNER_WIDTH - 4)
        scribus.setText(display_text, text_frame)

        # Apply font settings with variant support
        font_applied = False

        # Build list of font variants to try
        font_variants_to_try = [BANNER_TEXT_FONT_FAMILY]

        # Add common variations of the font name
        if "Myriad" in BANNER_TEXT_FONT_FAMILY:
            # Special handling for Myriad Pro fonts
            font_variants_to_try.extend([
                "Myriad Pro Condensed",
                "MyriadPro-Cond",
                "Myriad Pro Cond",
                "Myriad Pro",
                "MyriadPro",
                "Myriad"
            ])
        else:
            # Generic variations
            font_variants_to_try.extend([
                BANNER_TEXT_FONT_FAMILY.replace(" ", ""),  # Remove spaces
                BANNER_TEXT_FONT_FAMILY.replace(" ", "-")  # Spaces to dashes
            ])

        # Remove duplicates while preserving order
        seen = set()
        font_variants_to_try = [x for x in font_variants_to_try if not (x in seen or seen.add(x))]

        # Try all variants
        for variant in font_variants_to_try:
            try:
                scribus.setFont(variant, text_frame)
                scribus.setFontSize(BANNER_TEXT_FONT_SIZE, text_frame)
                font_applied = True
                break
            except:
                pass

        # Fallback to FONT_CANDIDATES if specified font fails
        if not font_applied:
            for font in FONT_CANDIDATES:
                try:
                    scribus.setFont(font, text_frame)
                    scribus.setFontSize(BANNER_TEXT_FONT_SIZE, text_frame)
                    font_applied = True
                    break
                except:
                    continue

        # Final fallback to DEFAULT_FONT
        if not font_applied:
            try:
                scribus.setFont(DEFAULT_FONT, text_frame)
                scribus.setFontSize(BANNER_TEXT_FONT_SIZE, text_frame)
            except:
                pass

        # Apply text properties
        scribus.setTextAlignment(scribus.ALIGN_CENTERED, text_frame)
        scribus.setTextColor(BANNER_TEXT_COLOR, text_frame)
        scribus.setLineColor("None", text_frame)
        
        if is_odd_page:
            scribus.rotateObject(LEFT_BANNER_ROTATION, text_frame)
            final_x = banner_x + LEFT_BANNER_HORIZONTAL_OFFSET
            final_y = banner_center_y - (text_height / 2) + LEFT_BANNER_VERTICAL_OFFSET
        else:
            scribus.rotateObject(RIGHT_BANNER_ROTATION, text_frame)
            final_x = banner_x + TOPIC_BANNER_WIDTH - RIGHT_BANNER_HORIZONTAL_OFFSET
            final_y = banner_center_y - (text_height / 2) + RIGHT_BANNER_VERTICAL_OFFSET
        scribus.moveObject(final_x, final_y, text_frame)
    except Exception as e:
        scribus.messageBox("Error in banner text", str(e))

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── SUPERSCRIPT MAP & NORMALIZATION ───────────
# ────────────────────────────────────────────────────────────────────────────────
_SUP_MAP = {
    '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
    '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'
}

# Unicode map for subscript digits
_SUB_MAP = {
    '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
    '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉'
}

def apply_quiz_superscripts(text_frame, original_text, cleaned_text):
    """
    Apply superscript/subscript formatting to quiz questions.
    Uses a simpler approach focusing on the most common patterns.
    """
    if not original_text or not text_frame or not cleaned_text:
        return
        
    try:
        # Find all superscript and subscript patterns
        sup_pattern = r'<sup>(\d+)</sup>'
        sub_pattern = r'<sub>(\d+)</sub>'
        
        sup_matches = list(re.finditer(sup_pattern, original_text))
        sub_matches = list(re.finditer(sub_pattern, original_text))
        
        # Create a list of all formatting to apply
        format_list = []
        
        # Process superscripts
        offset = 0
        for match in sup_matches:
            digits = match.group(1)
            # Find where these digits appear in the cleaned text
            start_pos = cleaned_text.find(digits, offset)
            if start_pos != -1:
                format_list.append((start_pos, len(digits), 'sup'))
                offset = start_pos + len(digits)
        
        # Process subscripts
        offset = 0
        for match in sub_matches:
            digits = match.group(1)
            # Find where these digits appear in the cleaned text
            start_pos = cleaned_text.find(digits, offset)
            if start_pos != -1:
                format_list.append((start_pos, len(digits), 'sub'))
                offset = start_pos + len(digits)
        
        # Apply formatting
        for start_pos, length, format_type in format_list:
            try:
                scribus.selectText(start_pos, length, text_frame)
                
                # Get current font size for the selection
                current_size = scribus.getFontSize(text_frame)
                
                if format_type == 'sup':
                    # Superscript: smaller font size
                    scribus.setFontSize(int(current_size * 0.75), text_frame)
                elif format_type == 'sub':
                    # Subscript: smaller font size  
                    scribus.setFontSize(int(current_size * 0.75), text_frame)
                    
            except:
                # If this particular formatting fails, continue with others
                continue
                
    except:
        # If the entire function fails, just continue - text will display without formatting
        pass

def handle_superscripts(text):
    """
    Superscript and subscript handling function for HTML and Markdown tagged content.
    Handles:
    - ^digits^ (markdown superscript)
    - ~digits~ (markdown subscript)
    - <span class="S-T...">digits</span> (superscripts)
    - <sup>digits</sup> (superscripts)
    - <sub>digits</sub> (subscripts)
    - Complex HTML span patterns with style attributes
    - Removes empty paragraphs with <br> tags that cause excessive spacing
    """
    if not text or not isinstance(text, str):
        return text

    original_text = text

    # SPECIAL CASE: Handle common unit patterns like "cm3" → "cm³", "m2" → "m²", "km2" → "km²"
    # This handles cases where JSON doesn't have markup
    # Match unit followed by digit, with optional space before digit
    text = re.sub(
        r'([ckm]?m)\s?(\d)(?=\s|$|[^\d])',
        lambda m: m.group(1) + _SUP_MAP.get(m.group(2), m.group(2)),
        text
    )

    # SPECIAL CASE: Handle "L4e" pattern for vehicle classes (shouldn't be converted)
    # Revert any accidental conversions in codes like L4e
    text = re.sub(r'L⁴e', 'L4e', text)

    # MARKDOWN SYNTAX: Handle ^digits^ for superscripts (e.g., m^2^ → m²)
    text = re.sub(
        r'\^(\d+)\^',
        lambda m: ''.join(_SUP_MAP.get(ch, ch) for ch in m.group(1)),
        text
    )

    # MARKDOWN SYNTAX: Handle ~digits~ for subscripts (e.g., H~2~O → H₂O)
    text = re.sub(
        r'~(\d+)~',
        lambda m: ''.join(_SUB_MAP.get(ch, ch) for ch in m.group(1)),
        text
    )

    # FIRST: Remove empty paragraphs that only contain <br> tags or are completely empty
    # These create excessive spacing between text blocks - remove them completely
    # Handle multiple <br> tags: <p><br><br></p>, <p><br /></p>, etc.
    text = re.sub(r'<p[^>]*>(\s*<br\s*/?\s*>)+\s*</p>', '', text)  # Remove paragraphs with only <br> tags
    text = re.sub(r'<p[^>]*>\s*</p>', '', text)  # Remove completely empty paragraphs

    # Handle special HTML span pattern for digit superscripts with any unit
    text = re.sub(
        r'([a-zA-Z]+)</span><span\s+class=["\']S-T\d+["\']\s+style=["\'][^"\']*vertical-align[^"\']*["\']>\s*(\d+)\s*</span>',
        lambda m: m.group(1) + ''.join(_SUP_MAP.get(ch, ch) for ch in m.group(2)),
        text
    )
    
    # Handle text followed by S-T span (e.g., "cm<span class="S-T18">3</span>") - FIRST
    # For quiz questions, convert to simple format that apply_quiz_superscripts can handle
    text = re.sub(
        r'([a-zA-Z]+)<span\s+class=(?:["\']|\")S-T[^"\'>]*(?:["\']|\")(?:\s*[^>]*)?>(\d+)</span>',
        lambda m: m.group(1) + '<sup>' + m.group(2) + '</sup>',
        text
    )
    
    # Also handle the simple case - convert to <sup> format
    text = re.sub(
        r'<span\s+class=(?:["\']|\")S-T[^"\'>]*(?:["\']|\")(?:\s*[^>]*)?>(\d+)</span>',
        lambda m: '<sup>' + m.group(1) + '</sup>',
        text
    )
    
    # Handle complex S-T spans with style attributes (vertical-align for superscripts)
    text = re.sub(
        r'([a-zA-Z]+)<span\s+class=["\\\']S-T[^"\\\']*["\\\'][^>]*vertical-align[^>]*>(\d+)</span>',
        lambda m: m.group(1) + ''.join(_SUP_MAP.get(ch, ch) for ch in m.group(2)),
        text,
        flags=re.IGNORECASE
    )
    
    # Replace any remaining standalone <span class="S-T...">digits</span> (superscripts)
    text = re.sub(
        r'<span\s+class=["\\\']S-T[^"\\\']*["\\\'](?:\s*[^>]*)?>(\d+)</span>',
        lambda m: ''.join(_SUP_MAP.get(ch, ch) for ch in m.group(1)),
        text
    )
    
    # Replace <sup>digits</sup> (superscripts)
    text = re.sub(
        r'<sup>(\d+)</sup>',
        lambda m: ''.join(_SUP_MAP.get(ch, ch) for ch in m.group(1)),
        text
    )
    
    # Replace <sub>digits</sub> (subscripts)
    text = re.sub(
        r'<sub>(\d+)</sub>',
        lambda m: ''.join(_SUB_MAP.get(ch, ch) for ch in m.group(1)),
        text
    )
    
    # Handle span elements with vertical-align style for superscripts
    text = re.sub(
        r'<span[^>]*vertical-align:\s*super[^>]*>(\d+)</span>',
        lambda m: ''.join(_SUP_MAP.get(ch, ch) for ch in m.group(1)),
        text,
        flags=re.IGNORECASE
    )
    
    # Handle span elements with vertical-align style for subscripts
    text = re.sub(
        r'<span[^>]*vertical-align:\s*sub[^>]*>(\d+)</span>',
        lambda m: ''.join(_SUB_MAP.get(ch, ch) for ch in m.group(1)),
        text,
        flags=re.IGNORECASE
    )
    
    return text

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── MAIN DRIVER FUNCTION ────────────
# ────────────────────────────────────────────────────────────────────────────────
def create_styled_header(text, font_size, bold, bg_color, text_color, padding, font_family=None):
    global y_offset
    # Skip empty text
    if not text:
        return
    # Aggressive trimming of whitespace and newlines
    text = text.strip()
    # Skip if text is empty after trimming
    if not text:
        return
    # Check if this is a template header by looking for the [template= pattern
    is_template_header = text.startswith("[template=")
    # Check if this is a module header (smaller font, used before templates)
    is_module_header = (font_size == MODULE_FONT_SIZE)
    # Use reduced padding for template headers
    actual_padding = 5 if is_template_header else padding
    # Check page overflow with strict margin enforcement
    frame_w = PAGE_WIDTH - MARGINS[0] - MARGINS[2]
    # Add extra vertical padding to text height
    text_h = measure_text_height(text, frame_w, False, font_size) + (actual_padding * 2)
    # Simple boundary check
    simple_boundary_check(text_h)
    # Only create a background rectangle if a real color is specified
    if bg_color and bg_color.lower() != "none":
        bg_rect = scribus.createRect(MARGINS[0], y_offset, frame_w, text_h)
        scribus.setFillColor(bg_color, bg_rect)
        try:
            scribus.setLineColor("None", bg_rect)
        except:
            pass
        # Simple boundary enforcement for background rectangle
        simple_constrain_element(bg_rect)
    else:
        bg_rect = None
    # Create text frame with padding
    text_frame = scribus.createText(
        MARGINS[0] + actual_padding,
        y_offset + actual_padding,
        frame_w - (actual_padding * 2),
        text_h - (actual_padding * 2)
    )
    scribus.setText(text, text_frame)
    try:
        scribus.setLineColor("None", text_frame)
    except:
        pass

    # Use provided font_family or fall back to MYRIAD_VARIANTS for backward compatibility
    if font_family:
        # Build list of font variants to try
        font_variants_to_try = [font_family]

        # Add common variations of the font name
        base_name = font_family.replace(" Cond", "").replace("-Cond", "")
        if "Myriad" in font_family:
            # Special handling for Myriad Pro fonts
            font_variants_to_try.extend([
                "Myriad Pro Condensed",
                "MyriadPro-Cond",
                "Myriad Pro Cond",
                "Myriad Pro",
                "MyriadPro",
                "Myriad"
            ])
        else:
            # Generic variations
            font_variants_to_try.extend([
                font_family.replace(" ", ""),  # Remove spaces
                font_family.replace(" ", "-"), # Spaces to dashes
                base_name
            ])

        # Remove duplicates while preserving order
        seen = set()
        font_variants_to_try = [x for x in font_variants_to_try if not (x in seen or seen.add(x))]

        # Try all variants
        font_applied = False
        for variant in font_variants_to_try:
            try:
                scribus.setFont(variant, text_frame)
                scribus.setFontSize(font_size, text_frame)
                scribus.textColor(text_color, text_frame)
                font_applied = True
                break
            except:
                pass

        # If specified font fails, try DEFAULT_FONT
        if not font_applied:
            try:
                scribus.setFont(DEFAULT_FONT, text_frame)
                scribus.setFontSize(font_size, text_frame)
                scribus.textColor(text_color, text_frame)
                font_applied = True
            except:
                pass

        # Final fallback to FONT_CANDIDATES
        if not font_applied:
            for f in FONT_CANDIDATES:
                try:
                    scribus.setFont(f, text_frame)
                    scribus.setFontSize(font_size, text_frame)
                    scribus.textColor(text_color, text_frame)
                    font_applied = True
                    break
                except:
                    pass

        # Last resort fallback - try basic system fonts
        if not font_applied:
            try:
                scribus.setFont("Arial", text_frame)
                scribus.setFontSize(font_size, text_frame)
                scribus.textColor(text_color, text_frame)
            except:
                pass
    else:
        # Original behavior - try MYRIAD_VARIANTS first (backward compatibility)
        font_applied = False
        for myriad_font in MYRIAD_VARIANTS:
            try:
                scribus.setFont(myriad_font, text_frame)
                scribus.setFontSize(font_size, text_frame)
                scribus.textColor(text_color, text_frame)
                font_applied = True
                break
            except:
                continue
        if not font_applied:
            try:
                scribus.setFont(DEFAULT_FONT, text_frame)
                scribus.setFontSize(font_size, text_frame)
                scribus.textColor(text_color, text_frame)
            except:
                success = False
                for f in FONT_CANDIDATES:
                    try:
                        scribus.setFont(f, text_frame)
                        scribus.setFontSize(font_size, text_frame)
                        scribus.textColor(text_color, text_frame)
                        success = True
                        break
                    except:
                        pass
                if not success:
                    try:
                        scribus.setFont("Arial", text_frame)
                        scribus.setFontSize(font_size, text_frame)
                        scribus.textColor(text_color, text_frame)
                    except:
                        pass
    if bold:
        try:
            scribus.setFontSize(font_size + 2, text_frame)
        except:
            pass
    if is_template_header:
        try:
            # Use fixed line spacing for consistent measurement (documentation-based approach)
            template_font_size = MODULE_TEXT_FONT_SIZE  # Use config template font size
            scribus.setLineSpacing(template_font_size * 1.0, text_frame)
        except:
            pass
    # Only auto-resize text frame if there is no background rectangle
    if not bg_rect:
        # Official Scribus method for precise text fitting based on documentation
        try:
            # Force text refresh to ensure all styling is applied and show live updates
            scribus.redrawAll()

            # Step 1: Ensure no overflow first with minimal expansion
            scribus.layoutText(text_frame)

            # Expand minimally to ensure all text is visible
            overflow_iterations = 0
            max_overflow_iterations = 100  # Safety limit
            while scribus.textOverflows(text_frame) and overflow_iterations < max_overflow_iterations:
                current_w, current_h = scribus.getSize(text_frame)
                # Check if expanding would exceed bottom margin (with minimal buffer)
                minimal_buffer = 10  # Just enough for page numbers
                max_allowed_height = (PAGE_HEIGHT - MARGINS[3] - minimal_buffer) - y_offset  # Minimal buffer
                if current_h + 3 > max_allowed_height:
                    break  # Don't expand if it would exceed margins
                scribus.sizeObject(current_w, current_h + 3, text_frame)
                scribus.layoutText(text_frame)
                overflow_iterations += 1

            # Step 2: Calculate exact height using official Scribus methods
            try:
                # Get actual text metrics
                num_lines = scribus.getTextLines(text_frame)
                line_spacing = scribus.getLineSpacing(text_frame)
                left, right, top, bottom = scribus.getTextDistances(text_frame)

                if num_lines > 0:
                    # Calculate exact height: (lines × spacing) + padding
                    exact_text_height = num_lines * line_spacing
                    exact_frame_height = exact_text_height + top + bottom

                    # Resize to exact height
                    current_w, current_h = scribus.getSize(text_frame)
                    scribus.sizeObject(current_w, exact_frame_height, text_frame)
                    scribus.layoutText(text_frame)

                    # Verify no overflow after exact sizing
                    if scribus.textOverflows(text_frame):
                        # Add minimal space if needed
                        scribus.sizeObject(current_w, exact_frame_height + line_spacing * 0.1, text_frame)
                        scribus.layoutText(text_frame)
            except:
                # If official method fails, minimal fallback
                pass
        except:
            # Basic fallback if documentation approach fails
            pass

        # Adjust width if needed
        text_width = 0
        try:
            text_width = scribus.getTextWidth(text_frame)
            if text_width > 0:
                text_width += 20
                text_width = max(text_width, 200)
                text_width = min(text_width, frame_w)
                scribus.sizeObject(text_width, scribus.getSize(text_frame)[1], text_frame)
        except:
            pass

    # CRITICAL FIX: Resize background rectangle to match actual text frame size
    # This prevents text from overflowing the background rectangle
    if bg_rect:
        try:
            # Get actual text frame dimensions after all resizing
            actual_frame_width, actual_frame_height = scribus.getSize(text_frame)
            frame_x, frame_y = scribus.getPosition(text_frame)

            # Calculate background rectangle size including padding
            bg_width = actual_frame_width + (actual_padding * 2)
            bg_height = actual_frame_height + (actual_padding * 2)

            # Resize and reposition background rectangle to match
            scribus.sizeObject(bg_width, bg_height, bg_rect)
            scribus.moveObject(frame_x - actual_padding, frame_y - actual_padding, bg_rect)

            # Ensure background is behind text
            try:
                scribus.moveObjectAbs(frame_x - actual_padding, frame_y - actual_padding, bg_rect)
            except:
                pass
        except Exception as e:
            # If resize fails, continue without error
            pass

    # Module headers have no spacing after (templates appear directly below)
    # Topic/chapter headers also have no spacing after (description handles its own spacing)
    # This prevents double-spacing between headers and descriptions
    spacing_after = 0
    # Use actual frame height instead of calculated text_h to minimize gap
    actual_frame_height = scribus.getSize(text_frame)[1]
    y_offset += actual_frame_height + spacing_after

    # Simple final constraint
    simple_constrain_element(text_frame)

    # Also constrain background rectangle if it exists
    if bg_rect:
        try:
            simple_constrain_element(bg_rect)
        except:
            pass
    
    # Text overflow is already handled by the documentation-based approach above
    
    # Enforce strict margin boundary
    enforce_margin_boundary()

    # Ensure no overlaps after header creation
    ensure_no_overlaps()

    return text_frame

def create_module_header(module_name):
    """Create a module header with the configured style."""
    return create_styled_header(
        module_name,
        MODULE_FONT_SIZE,
        MODULE_BOLD,
        MODULE_BG_COLOR,
        MODULE_TEXT_COLOR,
        MODULE_PADDING,
        MODULE_HEADER_FONT_FAMILY
    )

def create_topic_header(topic_name):
    """Create a topic header with the configured style."""
    return create_styled_header(
        topic_name,
        TOPIC_FONT_SIZE,
        TOPIC_BOLD,
        TOPIC_BG_COLOR,
        TOPIC_TEXT_COLOR,
        TOPIC_PADDING,
        TOPIC_HEADER_FONT_FAMILY
    )

def handle_text_styles(frame, style_segments, default_size, base_font_family=None):
    """Apply text styles based on parsed style segments."""
    # Skip empty text or if frame is not valid
    if not style_segments or not frame:
        return

    # Determine which font to use
    target_font = base_font_family if base_font_family else DEFAULT_FONT

    # Build list of font variants to try
    font_variants_to_try = [target_font]

    # Add common variations if it's Myriad
    if "Myriad" in target_font:
        font_variants_to_try.extend([
            "Myriad Pro Condensed",
            "MyriadPro-Cond",
            "Myriad Pro Cond",
            "Myriad Pro",
            "MyriadPro",
            "Myriad"
        ])
    else:
        # Generic variations
        font_variants_to_try.extend([
            target_font.replace(" ", ""),  # Remove spaces
            target_font.replace(" ", "-")  # Spaces to dashes
        ])

    # Remove duplicates while preserving order
    seen = set()
    font_variants_to_try = [x for x in font_variants_to_try if not (x in seen or seen.add(x))]

    # Set default font and size first before applying specific styles
    font_set = False
    for variant in font_variants_to_try:
        try:
            scribus.setFont(variant, frame)
            scribus.setFontSize(default_size, frame)
            font_set = True
            break
        except:
            pass

    # Fallback to FONT_CANDIDATES if nothing worked
    if not font_set:
        for f in FONT_CANDIDATES:
            try:
                scribus.setFont(f, frame)
                scribus.setFontSize(default_size, frame)
                break
            except:
                continue
    
    for start, length, style_dict in style_segments:
        if length <= 0:
            continue
            
        try:
            scribus.selectText(start, length, frame)
            
            # Determine which font family to use as base
            # Priority: base_font_family (from config) > HTML font > DEFAULT_FONT
            font_to_use = base_font_family if base_font_family else None

            # If no base font configured, try to get font from HTML style
            if not font_to_use and "font" in style_dict:
                font_to_use = style_dict.get("font-family", style_dict["font"])

            # Final fallback
            if not font_to_use:
                font_to_use = DEFAULT_FONT

            # Get style attributes (bold, italic)
            font_weight = style_dict.get("font-weight", None)
            font_style_attr = style_dict.get("font-style", None)

            # Apply bold/italic flags from style_dict
            if style_dict.get("bold", False) and not font_weight:
                font_weight = "bold"
            if style_dict.get("italic", False) and not font_style_attr:
                font_style_attr = "italic"

            # Get the best matching font with styles applied
            try:
                best_font = get_font_with_style(font_to_use, font_weight, font_style_attr)
                scribus.setFont(best_font, frame)
            except Exception as e:
                # If styled font fails, try without styles
                try:
                    scribus.setFont(font_to_use, frame)
                except:
                    # Last resort - try DEFAULT_FONT
                    try:
                        scribus.setFont(DEFAULT_FONT, frame)
                    except:
                        pass
            
            # Apply font size if specified
            if "font_size" in style_dict:
                size_str = style_dict["font_size"]
                try:
                    # Handle various formats: ##pt, ##px, ##em, ##%
                    if "pt" in size_str:
                        size = float(size_str.replace("pt", "").strip())
                        # Scale down to maintain hierarchy
                        size = size * 0.9
                    elif "px" in size_str:
                        # Approximate px to pt (0.75 factor) then scale down
                        size = float(size_str.replace("px", "").strip()) * 0.75 * 0.9
                    elif "em" in size_str:
                        size = float(size_str.replace("em", "").strip()) * default_size
                    elif "%" in size_str:
                        pct = float(re.sub(r'[^\d.]','', size_str)) / 100.0
                        size = default_size * pct
                    else:
                        size = float(re.sub(r'[^\d.]','', size_str))
                        # Don't scale down - use the size as specified in HTML or config

                    # Ensure minimum readable size (changed to match MODULE_TEXT_FONT_SIZE)
                    size = max(size, MODULE_TEXT_FONT_SIZE)
                    scribus.setFontSize(size, frame)
                except:
                    pass
            
            # Apply color if specified
            if "color" in style_dict:
                try:
                    color_value = style_dict["color"]
                    
                    # Handle various formats: named colors or hex values
                    if color_value.startswith("#"):
                        color_name = f"Color_{color_value.replace('#', '')}"
                        
                        # Check if color exists or needs to be defined
                        color_exists = False
                        try:
                            if color_name in scribus.getColorNames():
                                color_exists = True
                        except:
                            pass
                            
                        if not color_exists:
                            # Convert hex to RGB
                            if len(color_value) == 4:  # Short hex: #RGB
                                r = int(color_value[1] + color_value[1], 16)
                                g = int(color_value[2] + color_value[2], 16)
                                b = int(color_value[3] + color_value[3], 16)
                            else:  # Normal hex: #RRGGBB
                                r = int(color_value[1:3], 16)
                                g = int(color_value[3:5], 16)
                                b = int(color_value[5:7], 16)
                                
                            # Define the color
                            try:
                                scribus.defineColor(color_name, r, g, b)
                            except:
                                pass
                        
                        # Apply the color
                        scribus.setTextColor(color_name, frame)
                    else:
                        # For named colors
                        scribus.setTextColor(color_value, frame)
                except:
                    # If color application fails, try with capitalized variant
                    try:
                        capitalized = style_dict["color"].capitalize()
                        scribus.setTextColor(capitalized, frame)
                    except:
                        pass
            
            # Apply bold if specified but no styled font was applied
            # Some Scribus versions don't have true bold, so we increase font size
            if style_dict.get("bold", False) and not font_applied:
                try:
                    current_size = scribus.getFontSize(frame)
                    scribus.setFontSize(current_size + 1, frame)
                except:
                    pass
                    
            # Handle superscript/subscript vertical alignment
            if "vertical_align" in style_dict:
                try:
                    v_align = style_dict["vertical_align"]
                    
                    # Get current size
                    curr_size = scribus.getFontSize(frame)
                    
                    # Try to determine if this is a superscript or subscript
                    is_super = False
                    is_sub = False
                    
                    # Check text values like "super", "sup", "subscript", etc.
                    if any(x in v_align.lower() for x in ["super", "sup"]):
                        is_super = True
                    elif any(x in v_align.lower() for x in ["sub", "subscript"]):
                        is_sub = True
                    # Check percentage values
                    elif "%" in v_align:
                        try:
                            pct_value = float(re.sub(r'[^\d.-]', '', v_align))
                            is_super = pct_value > 0
                            is_sub = pct_value < 0
                        except:
                            pass
                    
                    # Apply the appropriate styling
                    if is_super:
                        # Make superscript smaller and try to raise it
                        reduced_size = curr_size * 0.7
                        scribus.setFontSize(reduced_size, frame)
                        
                        # Try different methods to raise the text
                        try:
                            # First try with baseline offset if available
                            scribus.setBaseline(-curr_size * 0.4, frame)
                        except:
                            try:
                                # Try with text distance instead
                                l, r, t, b = scribus.getTextDistances(frame)
                                scribus.setTextDistances(l, r, t - curr_size * 0.4, b, frame)
                            except:
                                pass
                    
                    elif is_sub:
                        # Make subscript smaller and try to lower it
                        reduced_size = curr_size * 0.7
                        scribus.setFontSize(reduced_size, frame)
                        
                        # Try different methods to lower the text
                        try:
                            # First try with baseline offset if available
                            scribus.setBaseline(curr_size * 0.2, frame)
                        except:
                            try:
                                # Try with text distance instead
                                l, r, t, b = scribus.getTextDistances(frame)
                                scribus.setTextDistances(l, r, t, b + curr_size * 0.2, frame)
                            except:
                                pass
                except Exception as e:
                    pass
                
        except Exception as e:
            # If any issue with applying styles to this segment, continue with the next
            pass

def create_pages_from_json(json_path=None, include_quizzes=True, filter_mode="all"):
    """
    Create a PDF document from a JSON file containing the content structure.

    Parameters:
        json_path (str): Path to the JSON file
        include_quizzes (bool): Whether to include quizzes in the PDF
        filter_mode (str): Filter mode for quizzes - "all", "true_only", or "false_only"
    """
    global y_offset, global_template_count, limit_reached, current_topic_text, current_topic_color, PRINT_QUIZZES, QUIZ_FILTER_MODE, OVERLAP_WARNINGS, BOUNDARY_VIOLATIONS

    # Initialize position tracking log
    OVERLAP_WARNINGS = []
    BOUNDARY_VIOLATIONS = []
    init_logging()

    # Set the global quiz printing flag and filter mode
    PRINT_QUIZZES = include_quizzes
    QUIZ_FILTER_MODE = filter_mode
    
    # Debug message to confirm filter settings
    filter_msg = "ALL quizzes"
    if filter_mode == "true_only":
        filter_msg = "ONLY TRUE quizzes"
    elif filter_mode == "false_only":
        filter_msg = "ONLY FALSE quizzes"
    scribus.messageBox("Quiz Filter", f"PDF will include {filter_msg}", scribus.ICON_INFORMATION)
    
    # Use default path if none provided
    if json_path is None:
        # Check if we have a default input file configured
        if DEFAULT_INPUT_FILE is not None and not ALWAYS_SHOW_INPUT_DIALOG:
            json_path = DEFAULT_INPUT_FILE
        else:
            # Show file dialog if no path provided or dialog always required
            json_path = scribus.fileDialog("Open JSON", "JSON Files (*.json)")
    
    if not json_path:
        return
    
    # Load and validate JSON
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        scribus.messageBox("Error", str(e), scribus.ICON_WARNING)
        return
        
    if not isinstance(data.get("areas"), list):
        scribus.messageBox("Bad JSON", "'areas' missing/invalid", scribus.ICON_WARNING)
        return
    
    # Set up paths and create new document
    # Use the configured image path if it's an absolute path, otherwise make it relative to JSON file
    if os.path.isabs(DEFAULT_IMAGES_PATH):
        base_pics = DEFAULT_IMAGES_PATH
    else:
        base_pics = os.path.join(os.path.dirname(json_path), DEFAULT_IMAGES_PATH)
        
    scribus.newDocument((PAGE_WIDTH, PAGE_HEIGHT), MARGINS,
                        scribus.PORTRAIT, 1, scribus.UNIT_POINTS,
                        scribus.PAGE_1, 0, 1)
    
    # Add page number to first page - exactly like copy 6
    add_page_number()
    
    # Initialize state
    y_offset = MARGINS[1]
    global_template_count = 0
    current_topic_text = None
    current_topic_color = None
    
    # Process content hierarchy
    for area in data["areas"]:
        if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
            break

        create_styled_header(
            f"{area.get('name','Unnamed')}",
            AREA_HEADER_FONT_SIZE,
            AREA_HEADER_BOLD,
            AREA_HEADER_BG_COLOR,
            AREA_HEADER_TEXT_COLOR,
            0,  # padding
            AREA_HEADER_FONT_FAMILY
        )
        # Area description
        desc_text = area.get("desc","")
        if desc_text:
            desc_text = handle_superscripts(desc_text)
            place_text_block_flow(desc_text, AREA_DESC_FONT_SIZE, balanced_columns=True, custom_spacing=HEADER_TO_DESC_SPACING, font_family=AREA_DESC_FONT_FAMILY)

        chapter_list = area.get("chapters", [])
        for chap_index, chap in enumerate(chapter_list):
            if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
                break

            # Add spacing between chapters (but not before first chapter)
            if chap_index > 0:
                y_offset += SECTION_TO_SECTION_SPACING

            create_styled_header(
                f"{chap.get('name','Unnamed')}",
                CHAPTER_HEADER_FONT_SIZE,
                CHAPTER_HEADER_BOLD,
                CHAPTER_HEADER_BG_COLOR,
                CHAPTER_HEADER_TEXT_COLOR,
                0,  # padding
                CHAPTER_HEADER_FONT_FAMILY
            )

            # Chapter description
            desc_text = chap.get("desc","")
            if desc_text:
                desc_text = handle_superscripts(desc_text)
                place_text_block_flow(desc_text, CHAPTER_DESC_FONT_SIZE, balanced_columns=True, custom_spacing=HEADER_TO_DESC_SPACING, font_family=CHAPTER_DESC_FONT_FAMILY)

            topic_list = chap.get("topics", [])
            for topic_index, topic in enumerate(topic_list):
                if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
                    break

                # Add spacing between topics (but not before first topic)
                if topic_index > 0:
                    y_offset += SECTION_TO_SECTION_SPACING

                # Clear previous topic banner and set new one
                current_topic_text = topic.get('name', 'Unnamed')
                banner_color = topic.get('banner_color', None)  # Read banner color from JSON
                create_topic_header(current_topic_text)
                add_vertical_topic_banner(current_topic_text, banner_color)

                # Topic description
                desc_text = topic.get("desc","")
                if desc_text:
                    desc_text = handle_superscripts(desc_text)
                    place_text_block_flow(desc_text, TOPIC_DESC_FONT_SIZE, balanced_columns=True, custom_spacing=HEADER_TO_DESC_SPACING, font_family=TOPIC_DESC_FONT_FAMILY)
                    # Spacing is already handled by custom_spacing parameter above

                # Handle templates directly in topic (without modules)
                if "templates" in topic and topic.get("templates"):
                    templates_list = topic.get("templates", [])
                    prev_id = None

                    for idx, tmpl in enumerate(templates_list):
                        if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
                            break

                        current_id = str(tmpl.get("id", ""))  # Convert to string

                        # Check if THIS template is a continuation of the PREVIOUS one
                        is_continuation_of_prev = False
                        if prev_id and current_id and prev_id == current_id:
                            is_continuation_of_prev = True

                        # Check if NEXT template continues from THIS one
                        next_tmpl_same_id = False
                        if idx + 1 < len(templates_list):
                            next_id = str(templates_list[idx + 1].get("id", ""))  # Convert to string
                            if current_id and next_id and current_id == next_id:
                                next_tmpl_same_id = True

                        # Pass both flags to process_template
                        process_template(tmpl, base_pics,
                                       is_continuation=is_continuation_of_prev,
                                       next_is_continuation=next_tmpl_same_id)

                        prev_id = current_id

                for mod in topic.get("modules", []):
                    if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
                        break

                    create_module_header(mod.get('name','Unnamed'))
                    # No spacing between module name and template content - removed to eliminate gap

                    # Track template IDs to detect continuation templates (same ID)
                    templates_list = mod.get("templates", [])
                    prev_id = None

                    for idx, tmpl in enumerate(templates_list):
                        if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
                            break

                        current_id = str(tmpl.get("id", ""))  # Convert to string for reliable comparison

                        # Check if THIS template is a continuation of the PREVIOUS one
                        is_continuation_of_prev = False
                        if prev_id and current_id and prev_id == current_id:
                            is_continuation_of_prev = True

                        # Check if NEXT template continues from THIS one
                        next_tmpl_same_id = False
                        if idx + 1 < len(templates_list):
                            next_id = str(templates_list[idx + 1].get("id", ""))  # Convert to string
                            if current_id and next_id and current_id == next_id:
                                next_tmpl_same_id = True

                        # Pass both flags to process_template
                        process_template(tmpl, base_pics,
                                       is_continuation=is_continuation_of_prev,
                                       next_is_continuation=next_tmpl_same_id)

                        prev_id = current_id

                    # Apply balancing after each module to ensure consistency
                    if column_mgr.use_columns and not column_mgr.quiz_mode:
                        column_mgr.ensure_consistent_balancing()

    # Apply final uniform balancing to all text frames
    column_mgr.ensure_consistent_balancing()

    # Force final refresh to ensure all content is displayed
    scribus.redrawAll()

    # Determine output PDF path
    pdf_out = None
    
    # Check if we need to show output dialog
    if ALWAYS_SHOW_OUTPUT_DIALOG:
        pdf_out = scribus.fileDialog("Save PDF", "PDF (*.pdf)",
                                 os.path.splitext(json_path)[0] + ".pdf", True)
    else:
        # Use default output file if specified
        if DEFAULT_OUTPUT_FILE is not None:
            pdf_out = DEFAULT_OUTPUT_FILE
        else:
            # Otherwise derive from input filename
            pdf_out = os.path.splitext(json_path)[0] + ".pdf"
    
    if pdf_out:
        try:
            pdf = scribus.PDFfile()
            pdf.file = pdf_out
            pdf.save()

            # Generate summary report
            summary_file = generate_summary_report()

            # Show completion message with log info
            msg = f"PDF saved: {pdf_out}\n\n"
            if DEBUG_LOG_FILE and summary_file:
                msg += f"=== POSITION LOGS ===\n"
                msg += f"Folder: {LOG_DIR}\n\n"
                msg += f"Summary file:\n{summary_file}\n\n"
                msg += f"Detailed log:\n{DEBUG_LOG_FILE}\n\n"
                if BOUNDARY_VIOLATIONS:
                    msg += f"⚠ {len(BOUNDARY_VIOLATIONS)} CRITICAL boundary violations detected!\n"
                if OVERLAP_WARNINGS:
                    msg += f"⚠ {len(OVERLAP_WARNINGS)} warnings (elements close to boundary)\n"
                if not BOUNDARY_VIOLATIONS and not OVERLAP_WARNINGS:
                    msg += "✓ No issues detected - all elements within boundaries\n"
            else:
                msg += "Note: Logging may have failed. Check script directory for logs.\n"

            scribus.messageBox("Done", msg, scribus.ICON_INFORMATION)
        except Exception as e:
            scribus.messageBox("Error", str(e), scribus.ICON_WARNING)

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── SCRIPT ENTRY POINT ────────────
# ────────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if scribus.haveDoc():
        scribus.messageBox("Close the current document first.", "",
                           scribus.ICON_WARNING)
    else:
        # First ask whether to include quizzes at all
        include_quizzes = True
        
        # User decides whether to include quizzes
        msg = "Quiz Option Selection\n\n"
        msg += "Press Y key, then OK: to INCLUDE quizzes\n"
        msg += "Press N key, then OK: to EXCLUDE quizzes\n\n"
        msg += "Your choice (Y/N): "
        
        user_input = scribus.valueDialog("Include Quizzes?", msg, "Y")
        
        # Check the first letter of their input (case insensitive)
        if user_input and user_input.strip().upper().startswith("N"):
            include_quizzes = False
        else:
            include_quizzes = True
        
        # Display confirmation of the quiz inclusion choice
        status_msg = "Quizzes will be included" if include_quizzes else "Quizzes will NOT be included"
        scribus.messageBox("Quiz Status", status_msg, scribus.ICON_INFORMATION)
        
        # If including quizzes, ask which filter mode to use
        quiz_filter_mode = "all"  # Default mode
        if include_quizzes:
            # Ask user to choose a filter mode
            filter_msg = "Quiz Filter Mode Selection\n\n"
            filter_msg += "Enter a number to choose which quizzes to include:\n"
            filter_msg += "1 - Include ALL quizzes (true and false)\n"
            filter_msg += "2 - Include ONLY TRUE quizzes\n"
            filter_msg += "3 - Include ONLY FALSE quizzes\n\n"
            filter_msg += "Your choice (1-3): "
            
            filter_input = scribus.valueDialog("Quiz Filter Mode", filter_msg, "1")
            
            # Parse the filter choice with better validation
            valid_choice = False
            try:
                choice = int(filter_input.strip())
                if choice == 2:
                    quiz_filter_mode = "true_only"
                    filter_status = "Only TRUE quizzes will be included"
                    valid_choice = True
                elif choice == 3:
                    quiz_filter_mode = "false_only"
                    filter_status = "Only FALSE quizzes will be included"
                    valid_choice = True
                elif choice == 1:
                    quiz_filter_mode = "all"
                    filter_status = "ALL quizzes (true and false) will be included"
                    valid_choice = True
                else:
                    quiz_filter_mode = "all"  # Default for invalid number input
                    filter_status = "Invalid option. ALL quizzes will be included."
            except:
                quiz_filter_mode = "all"  # Default for non-numeric input
                filter_status = "Invalid input. ALL quizzes will be included."
            
            # Display the selected filter mode with clearer messaging
            icon = scribus.ICON_INFORMATION if valid_choice else scribus.ICON_WARNING
            scribus.messageBox("Quiz Filter Mode", filter_status, icon)
        
        # Call the main function with the specified input file (None will trigger file dialog)
        create_pages_from_json(
            json_path=DEFAULT_INPUT_FILE if not ALWAYS_SHOW_INPUT_DIALOG else None,
            include_quizzes=include_quizzes, 
            filter_mode=quiz_filter_mode
        )
