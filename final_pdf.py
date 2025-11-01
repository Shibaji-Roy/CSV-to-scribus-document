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
from bs4 import BeautifulSoup

# Import all configuration settings
try:
    from config import *
except ImportError:
    # Fallback values if config.py is missing
    import sys
    scribus.messageBox("Config Error", "Could not import config.py. Using default values.", scribus.ICON_WARNING)

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
PRINT_QUIZZES         = False   # Will be set based on user choice

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
        # Build styled options
        styled_options = []
        if is_bold and is_italic:
            styled_options = [
                f"{base_font} Bold Italic",
                f"{base_font} Italic Bold",
                f"{base_font} BoldItalic",
                f"{base_font} Bold-Italic",
                f"{base_font}-Bold-Italic"
            ]
        elif is_bold:
            styled_options = [
                f"{base_font} Bold",
                f"{base_font}Bold",
                f"{base_font}-Bold"
            ]
        elif is_italic:
            styled_options = [
                f"{base_font} Italic",
                f"{base_font}Italic",
                f"{base_font}-Italic"
            ]
        else:
            styled_options = [
                f"{base_font} Regular",
                f"{base_font}Regular",
                f"{base_font}-Regular",
                base_font
            ]
        # Try styled options (case-sensitive, then case-insensitive)
        for styled_font in styled_options:
            if styled_font in available_fonts:
                return styled_font
            for font in available_fonts:
                if styled_font.lower() == font.lower():
                    return font
        # Try base font (case-sensitive, then case-insensitive, then substring)
        if base_font in available_fonts:
            return base_font
        for font in available_fonts:
            if base_font.lower() == font.lower() or base_font.lower() in font.lower():
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

def parse_html_to_segments(html):
    """
    Parse HTML text into a list of (text, style) segments.
    Handles nested tags, styles, and ensures proper paragraph breaks.
    """
    # First trim any whitespace from the input HTML
    html = html.strip() if html else ""
    
    soup = BeautifulSoup(html, "html.parser")
    segments = []
    
    def walk(node, cur_style):
        if node.name is None:
            txt = re.sub(r"\{[^}]+\}", "", str(node))
            # Trim excess whitespace within text nodes - more aggressive
            txt = re.sub(r'\s+', ' ', txt)
            if txt:
                segments.append((txt, cur_style.copy()))
            return
            
        style = cur_style.copy()
        tag = node.name.lower()
        
        # Handle formatting tags
        if tag in ("b", "strong"):
            style["bold"] = True
            style["font-weight"] = "bold"
        if tag in ("i", "em"):
            style["italic"] = True
            style["font-style"] = "italic"
        if tag == "u":
            style["underline"] = True
            
        # Handle font tag attributes
        if tag == "font":
            if node.has_attr("color"):    style["color"]     = node["color"]
            if node.has_attr("size"):     style["font_size"] = node["size"]
            if node.has_attr("face"):     
                style["font"] = node["face"]
                style["font-family"] = node["face"]
            
        # Handle CSS style attributes
        if node.has_attr("style"):
            css = parse_style_attribute(node["style"])
            
            # Extract font-related properties
            if "font-family" in css:
                style["font"] = css["font-family"]
                style["font-family"] = css["font-family"]
                
                # If we also have weight or style, determine the correct font name
                if "font-weight" in css or "font-style" in css:
                    font_family = css["font-family"]
                    font_weight = css.get("font-weight")
                    font_style = css.get("font-style")
                    
                    # Store individual style attributes
                    if font_weight:
                        style["font-weight"] = font_weight
                    if font_style:
                        style["font-style"] = font_style
                    
                    # Get the font with correct styling
                    full_font_name = get_font_with_style(font_family, font_weight, font_style)
                    style["font"] = full_font_name
                    
                    # Set style flags for backup styling
                    if font_weight in ["bold", "bolder", "700", "800", "900"]:
                        style["bold"] = True
                    if font_style == "italic":
                        style["italic"] = True
            else:
                # Handle individual style properties
                if "font-weight" in css:  
                    style["font-weight"] = css["font-weight"]
                    style["bold"] = css["font-weight"] in ["bold", "bolder", "700", "800", "900"]
                if "font-style" in css:   
                    style["font-style"] = css["font-style"]
                    style["italic"] = css["font-style"] == "italic"
            
            # Handle other properties
            if "color" in css:            style["color"] = css["color"]
            if "font-size" in css:        style["font_size"] = css["font-size"]
            if "text-decoration" in css:  style["underline"] = "underline" in css["text-decoration"]
            if "vertical-align" in css:   style["vertical_align"] = css["vertical-align"]
            
        # Handle paragraph and line break tags - simplified to prevent doubled newlines
        if tag == "p":
            if segments and segments[-1][0] != "\n":
                segments.append(("\n", style.copy()))
            for c in node.children: 
                walk(c, style)
            if segments and segments[-1][0] != "\n":
                segments.append(("\n", style.copy()))
            return
            
        if tag == "br":
            if segments and segments[-1][0] != "\n":
                segments.append(("\n", style.copy()))
            return
            
        # Process child nodes with the updated style
        for c in node.children:
            walk(c, style)
            
    walk(soup, {})
    
    # Remove trailing newlines
    while segments and segments[-1][0] == "\n":
        segments.pop()
        
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
    # Leave 20 points buffer for page number (15pt height + 3px below + clearance)
    return PAGE_HEIGHT - y_offset - MARGINS[3] - 20

def enforce_margin_boundary():
    """Ensure y_offset never exceeds the bottom margin boundary with buffer for page number."""
    global y_offset
    # Leave 20 points buffer for page number (15pt height + 3px below + clearance)
    safe_boundary = PAGE_HEIGHT - MARGINS[3] - 20
    if y_offset > safe_boundary:
        y_offset = safe_boundary

def simple_boundary_check(element_height):
    """
    Simple boundary check - create new page if element won't fit.
    """
    global y_offset
    safe_boundary = PAGE_HEIGHT - MARGINS[3] - 20  # Page number buffer
    
    if y_offset + element_height > safe_boundary:
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
        max_allowed_bottom = PAGE_HEIGHT - MARGINS[3] - 20  # Page number buffer
        
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
            scribus.setLineSpacing(font_size * 1.2, probe)
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
    global y_offset
    global quiz_heading_placed_on_page
    scribus.newPage(-1)
    y_offset = MARGINS[1]

    # Force refresh to show live page creation updates
    scribus.redrawAll()
    # DON'T reset quiz_heading_placed_on_page - preserve quiz header state across pages
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
        try:
            scribus.setFont("Liberation Sans", page_num_box)
        except:
            font_candidates = ["DejaVu Sans", "Arial", "Helvetica"]
            for font in font_candidates:
                try:
                    scribus.setFont(font, page_num_box)
                    break
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


def place_roadsigns_grid(images, base_path, start_y, max_height=None):
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
    
    # Available width for images
    available_width = PAGE_WIDTH - MARGINS[0] - MARGINS[2]
    
    # For roadsigns, we use 2 images per row instead of 3
    images_per_row = 2
    
    # Process images in rows of 2
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

        # Left align for roadsigns - ALL images on the same side (left)
        x = MARGINS[0]

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
        safe_boundary = PAGE_HEIGHT - MARGINS[3] - 22  # Page number buffer
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
def place_text_block_flow(html_text, font_size=8, bold=False, in_template=False, no_bottom_gap=False):
    """
    Place an HTML text block with flowing text and formatting.
    Handles text wrapping, styling, and creates multiple frames if needed.
    If no_bottom_gap is True, do not add the extra line at the bottom.
    """
    global y_offset
    
    if not html_text:  # Skip empty text blocks
        return
    
    # Format HTML with BeautifulSoup to handle malformed HTML better
    segments = parse_html_to_segments(html_text)
    
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
    text_h = measure_text_height(plain, frame_w, in_template, font_size)
    
    # No buffer needed - start with generous size and expand as needed
    
    # Simple boundary check - create new page if needed
    simple_boundary_check(text_h)

    frame = scribus.createText(MARGINS[0], y_offset, frame_w, text_h)
    
    # Remove the frame border
    try:
        scribus.setLineColor("None", frame)
    except:
        pass
    
    # Set internal padding for all text frames with more padding for templates
    try:
        # Set internal padding (left, right, top, bottom)
        if in_template:
            scribus.setTextDistances(*TEMPLATE_TEXT_PADDING, frame)
        else:
            scribus.setTextDistances(*REGULAR_TEXT_PADDING, frame)
    except:
        pass
            
    scribus.setText(plain, frame)
    
    # Create style segments in the format expected by handle_text_styles
    style_segments = []
    pos = 0
    for txt, sty in normalized_segments:
        if len(txt) > 0:
            # Ensure text is visible against the current background
            if in_template and "color" not in sty:
                # For template text with no specified color, force high contrast
                # Determine if background is light or dark
                bg_color = CURRENT_COLOR
                is_dark_bg = is_dark_color(bg_color)
                
                # Set text color to white for dark backgrounds, black for light
                if is_dark_bg:
                    sty["color"] = "White"
                else:
                    sty["color"] = "Black"
            
            style_segments.append((pos, len(txt), sty))
            pos += len(txt)
    
    # Apply all text styles using the central function
    handle_text_styles(frame, style_segments, font_size)

    # Force font size for area and topic descriptions
    if not in_template and font_size <= max(AREA_DESC_FONT_SIZE, TOPIC_DESC_FONT_SIZE):
        try:
            scribus.selectText(0, len(plain), frame)
            scribus.setFontSize(font_size, frame)
        except:
            pass

    # Set fixed line spacing for ALL text (both template and non-template) for consistent measurement
    try:
        scribus.setLineSpacing(font_size * 1.2, frame)
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

    # Remove complex initial overflow handling - let the main handler do it

    # Trim the frame height to fit the content exactly - DISABLED for consistency
    # This section was causing inconsistent extra space at the bottom of containers
    # The measure_text_height and overflow handling already size frames correctly

    # Simple final constraint
    simple_constrain_element(frame)

    # Official Scribus method for precise text fitting based on documentation
    try:
        # Force text refresh to ensure all styling is applied and show live updates
        scribus.redrawAll()

        # Step 1: Ensure no overflow first with minimal expansion
        scribus.layoutText(frame)

        # Expand minimally to ensure all text is visible
        while scribus.textOverflows(frame):
            current_w, current_h = scribus.getSize(frame)
            scribus.sizeObject(current_w, current_h + 3, frame)
            scribus.layoutText(frame)

        # Step 2: Calculate exact height using official Scribus methods
        try:
            # Force fixed line spacing for accurate calculation (as per docs)
            scribus.setLineSpacing(font_size * 1.2, frame)
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
        # Basic fallback
        overflow_count = 0
        while scribus.textOverflows(frame) and overflow_count < 15:
            try:
                current_w, current_h = scribus.getSize(frame)
                scribus.sizeObject(current_w, current_h + 2, frame)
                overflow_count += 1
            except:
                break

    # Get the actual final height of the frame after ALL adjustments
    try:
        final_frame_height = scribus.getSize(frame)[1]
    except:
        final_frame_height = text_h  # Fallback to original height

    # Update y offset for next element using actual frame height
    y_offset += final_frame_height + BLOCK_SPACING

    # Enforce strict margin boundary
    enforce_margin_boundary()

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
def place_wrapped_text_and_images(text_arr, image_list, base_path,
                                default_font_size=8, in_template=False, alignment="C"):
    """Places text with images with text wrapping."""
    global y_offset
    text_arr = [str(t) for t in (text_arr or [])]
    
    # More aggressively clean text items
    cleaned_text_arr = []
    for t in text_arr:
        # Apply superscripts before trimming
        t = handle_superscripts(t)
        # Trim all whitespace including newlines
        t = t.strip()
        if t:  # Only add non-empty text
            cleaned_text_arr.append(t)
    
    if not cleaned_text_arr and not image_list:
        return
        
    if len(cleaned_text_arr) == 1:
        text = cleaned_text_arr[0]
    else:
        # Join with single newlines between items
        text = "\n".join(cleaned_text_arr)
    
    # Ensure text ends without trailing whitespace
    text = text.rstrip()
    
    segs = parse_html_to_segments(text)
    
    # Double check no trailing newlines in segments
    while segs and segs[-1][0] == "\n":
        segs.pop()
    
    plain = "".join(t for t, _ in segs)
    if not plain and not image_list:
        return
    
    seg_index = []
    cursor = 0
    
    for txt, sty in segs:
        # Ensure text is visible against the current background
        if in_template and "color" not in sty:
            # For template text with no specified color, force high contrast
            # Determine if background is light or dark
            bg_color = CURRENT_COLOR
            is_dark_bg = is_dark_color(bg_color)
            
            # Set text color to white for dark backgrounds, black for light
            if is_dark_bg:
                sty["color"] = "White" 
            else:
                sty["color"] = "Black"
        
        seg_index.append((cursor, txt, sty))
        cursor += len(txt)

    frame_w = PAGE_WIDTH - MARGINS[0] - MARGINS[2]
    remaining = plain
    base_offset = 0
    placed_images = False

    while remaining:
        # Calculate available space with minimal buffer for page numbers
        # Use just 10 points to ensure page numbers are visible
        minimal_buffer = 10
        max_available_height = PAGE_HEIGHT - MARGINS[3] - y_offset - minimal_buffer

        if max_available_height < MIN_SPACE_THRESHOLD:
            new_page()
            minimal_buffer = 10
            max_available_height = PAGE_HEIGHT - MARGINS[3] - y_offset - minimal_buffer

        # Create a frame that uses ALL available space down to the margin
        frame = scribus.createText(MARGINS[0], y_offset, frame_w, max_available_height)
        try: scribus.setLineColor("None", frame)
        except: pass
        try:
            if in_template:
                scribus.setTextDistances(*TEMPLATE_TEXT_PADDING, frame)
            else:
                scribus.setTextDistances(*REGULAR_TEXT_PADDING, frame)
        except: pass

        # Set fixed line spacing for consistent measurement (documentation-based approach)
        try:
            scribus.setLineSpacing(default_font_size * 1.2, frame)
        except: pass

        # Get frame position for proper image placement
        frame_x, frame_y = scribus.getPosition(frame)
        frame_w_actual, frame_h_actual = scribus.getSize(frame)
        # Get text padding to position images inside the frame
        try:
            left_pad, right_pad, top_pad, bottom_pad = scribus.getTextDistances(frame)
        except:
            left_pad, right_pad, top_pad, bottom_pad = (0, 0, 0, 0)


        # Enable text flow mode for wrapping around images - EXACTLY like final_pdf copy
        try:
            scribus.setTextFlowMode(frame, TEXT_FLOW_INTERACTIVE)
        except:
            pass

        scribus.setText(remaining, frame)

        # Initially assume all text will be processed
        # We'll adjust this later if there's overflow that can't be handled
        chunk = remaining
        chunk_len = len(remaining)

        # Apply DEFAULT_FONT first, then try font candidates if it fails
        try:
            scribus.setFont(DEFAULT_FONT, frame)
            scribus.setFontSize(default_font_size, frame)
        except:
            # Fall back to font candidates if DEFAULT_FONT fails
            for f in FONT_CANDIDATES:
                try:
                    scribus.setFont(f, frame)
                    scribus.setFontSize(default_font_size, frame)
                    break
                except:
                    pass

        for start, txt, seg_sty in seg_index:
            end = start + len(txt)
            ov_s = max(start, base_offset)
            ov_e = min(end, base_offset + chunk_len)
            if ov_e <= ov_s:
                continue
            rel_off = ov_s - base_offset
            rel_len = ov_e - ov_s
            try:
                scribus.selectText(rel_off, rel_len, frame)
                if "font_size" in seg_sty:
                    size_str = seg_sty["font_size"]
                    if "%" in size_str:
                        pct = float(re.sub(r'[^\d.]','', size_str)) / 100.0
                        sz = int(default_font_size * pct)
                    else:
                        # Scale down JSON-specified sizes to maintain hierarchy
                        # If JSON says 10pt, reduce it proportionally
                        original_size = float(re.sub(r'[^\d.]','', size_str))
                        # Apply a scaling factor (0.9 = 90% of original)
                        sz = int(original_size * 0.9)
                        # Ensure minimum size
                        sz = max(sz, 7)
                    scribus.setFontSize(sz, frame)
                if "font" in seg_sty:
                    scribus.setFont(seg_sty["font"], frame)
                if "color" in seg_sty:
                    scribus.setTextColor(seg_sty["color"].capitalize(), frame)
                
                # Handle vertical alignment for superscript/subscript
                if "vertical_align" in seg_sty:
                    try:
                        v_align = seg_sty["vertical_align"]
                        # Get current position and size
                        x, y = scribus.getTextDistances(frame)
                        curr_size = scribus.getFontSize(frame)
                        
                        # Apply superscript/subscript
                        if "%" in v_align and int(re.sub(r'[^\d.]', '', v_align)) > 0:
                            # Positive percentage = superscript
                            offset = int(curr_size * 0.4)
                            scribus.setTextOffset(frame, x, offset)  # Positive for up in Scribus
                            # Superscripts are typically smaller
                            if "font_size" not in seg_sty:
                                scribus.setFontSize(int(curr_size * 0.7), frame)
                        elif "sub" in v_align or (("%" in v_align) and int(re.sub(r'[^\d.]', '', v_align)) < 0):
                            # Negative percentage or "sub" = subscript
                            offset = int(curr_size * 0.2)
                            scribus.setTextOffset(frame, x, offset)  # Positive for down
                            # Subscripts are typically smaller
                            if "font_size" not in seg_sty:
                                scribus.setFontSize(int(curr_size * 0.7), frame)
                    except: 
                        pass
                
                # Apply bold
                if seg_sty.get("bold"):
                    try: 
                        scribus.setFontSize(scribus.getFontSize(frame)+2, frame)
                    except: 
                        pass
            except:
                pass

        # Handle roadsigns/images placement within the text frame
        if not placed_images and image_list:
            # Calculate total height needed for all images in single column
            # Calculate maximum available width inside frame (accounting for padding)
            safety_margin = 5  # Extra margin to ensure image stays within bounds
            max_available_width = frame_w_actual - left_pad - right_pad - safety_margin
            # Use smaller percentage (15%) and enforce maximum width of 80 points
            col_width = min(frame_w_actual * 0.15, max_available_width, 80)  # Use 15% of frame width, max 80pt
            total_image_height = 0
            
            # Calculate maximum available height inside frame
            max_available_height = frame_h_actual - top_pad - bottom_pad - safety_margin

            for img_name in image_list:
                img_path = os.path.join(base_path, img_name)
                # Get real image dimensions
                if Image:
                    try:
                        with Image.open(img_path) as im:
                            orig_w, orig_h = im.size
                    except:
                        orig_w, orig_h = (300, 200)
                else:
                    orig_w, orig_h = (300, 200)

                # Scale to fit width while maintaining aspect ratio
                scale = float(col_width) / float(orig_w) if orig_w else 1.0
                new_h = orig_h * scale

                # Add height for this image with spacing
                total_image_height += new_h + BLOCK_SPACING

            # Get current frame height
            _, frame_h = scribus.getSize(frame)

            # Ensure frame is tall enough for all images
            if frame_h < total_image_height:
                # Resize the frame to fit all images properly
                scribus.sizeObject(frame_w, total_image_height, frame)
                frame_h = total_image_height

            if len(image_list) == 1:
                # Single image with text wrap - place on LEFT side with border box
                img_path = os.path.join(base_path, image_list[0])

                # Get real image dimensions
                if Image:
                    try:
                        with Image.open(img_path) as im:
                            orig_w, orig_h = im.size
                    except:
                        orig_w, orig_h = (300, 200)
                else:
                    orig_w, orig_h = (300, 200)

                # Scale to fit width while maintaining aspect ratio
                # Respect maximum available dimensions
                scale = float(col_width) / float(orig_w) if orig_w else 1.0
                new_h = orig_h * scale
                new_w = col_width
                
                # If height exceeds available space, scale down
                if new_h > max_available_height:
                    scale = float(max_available_height) / float(orig_h) if orig_h else 1.0
                    new_h = max_available_height
                    new_w = orig_w * scale
                    
                # Ensure width doesn't exceed available space
                if new_w > max_available_width:
                    scale = float(max_available_width) / float(new_w) if new_w else 1.0
                    new_w = max_available_width
                    new_h = new_h * scale

                # Position image inside text frame - LEFT aligned
                x = frame_x + left_pad  # Use frame position with padding
                y = frame_y + top_pad  # Use frame position with padding

                # Create image frame directly (no separate border box)
                img_frame = scribus.createImage(x, y, new_w, new_h)
                scribus.loadImage(img_path, img_frame)
                scribus.setScaleImageToFrame(True, True, img_frame)
                scribus.setLineColor("None", img_frame)  # No border on image
                
                # Resize frame to match image exactly (removes white space)
                try:
                    scribus.setScaleFrameToImage(img_frame)
                except:
                    pass

                # AGGRESSIVE: Force image to stay within text frame boundaries
                try:
                    img_x, img_y = scribus.getPosition(img_frame)
                    img_w, img_h = scribus.getSize(img_frame)
                    
                    # Calculate absolute boundaries
                    frame_right = frame_x + frame_w_actual - right_pad
                    frame_bottom = frame_y + frame_h_actual - bottom_pad
                    
                    # Calculate how much image exceeds boundaries
                    img_right = img_x + img_w
                    img_bottom = img_y + img_h
                    
                    # Clip width if exceeds
                    if img_right > frame_right:
                        new_img_w = frame_right - img_x
                        if new_img_w < 10:  # Minimum width
                            new_img_w = 10
                        scribus.sizeObject(new_img_w, img_h, img_frame)
                        scribus.setScaleImageToFrame(True, True, img_frame)
                        img_w = new_img_w
                    
                    # Clip height if exceeds
                    img_bottom = img_y + img_h  # Recalculate after potential width change
                    if img_bottom > frame_bottom:
                        new_img_h = frame_bottom - img_y
                        if new_img_h < 10:  # Minimum height
                            new_img_h = 10
                        scribus.sizeObject(img_w, new_img_h, img_frame)
                        scribus.setScaleImageToFrame(True, True, img_frame)
                except:
                    pass

                # Strict boundary enforcement
                simple_constrain_element(img_frame)

                # Enable text wrap around the image
                try:
                    scribus.setTextFlowMode(img_frame, TEXT_FLOW_OBJECTBOUNDINGBOX)
                except:
                    pass
            else:
                # Multiple images - place them in a single column within text frame
                x = frame_x + left_pad  # Use frame position with padding
                current_y = frame_y + top_pad

                for img_name in image_list:
                    img_path = os.path.join(base_path, img_name)

                    # Get real image dimensions
                    if Image:
                        try:
                            with Image.open(img_path) as im:
                                orig_w, orig_h = im.size
                        except:
                            orig_w, orig_h = (300, 200)
                    else:
                        orig_w, orig_h = (300, 200)

                    # Scale to fit width while maintaining aspect ratio
                    scale = float(col_width) / float(orig_w) if orig_w else 1.0
                    new_h = orig_h * scale
                    new_w = col_width
                    
                    # Ensure dimensions don't exceed available space
                    if new_w > max_available_width:
                        scale = float(max_available_width) / float(new_w) if new_w else 1.0
                        new_w = max_available_width
                        new_h = new_h * scale
                    
                    # Check if we have enough vertical space for this image
                    space_left = max_available_height - (current_y - (frame_y + top_pad))
                    if new_h > space_left:
                        # Scale down to fit remaining space
                        scale = float(space_left) / float(new_h) if new_h else 1.0
                        new_h = space_left
                        new_w = new_w * scale

                    # Create image frame directly (no separate border box)
                    img_frame = scribus.createImage(x, current_y, new_w, new_h)
                    scribus.loadImage(img_path, img_frame)
                    scribus.setScaleImageToFrame(True, True, img_frame)
                    scribus.setLineColor("None", img_frame)  # No border on image
                    
                    # Resize frame to match image exactly (removes white space)
                    try:
                        scribus.setScaleFrameToImage(img_frame)
                    except:
                        pass

                    # AGGRESSIVE: Force image to stay within text frame boundaries
                    try:
                        img_x, img_y = scribus.getPosition(img_frame)
                        img_w, img_h = scribus.getSize(img_frame)
                        
                        # Calculate absolute boundaries
                        frame_right = frame_x + frame_w_actual - right_pad
                        frame_bottom = frame_y + frame_h_actual - bottom_pad
                        
                        # Calculate how much image exceeds boundaries
                        img_right = img_x + img_w
                        img_bottom = img_y + img_h
                        
                        # Clip width if exceeds
                        if img_right > frame_right:
                            new_img_w = frame_right - img_x
                            if new_img_w < 10:  # Minimum width
                                new_img_w = 10
                            scribus.sizeObject(new_img_w, img_h, img_frame)
                            scribus.setScaleImageToFrame(True, True, img_frame)
                            img_w = new_img_w
                        
                        # Clip height if exceeds
                        img_bottom = img_y + img_h  # Recalculate after potential width change
                        if img_bottom > frame_bottom:
                            new_img_h = frame_bottom - img_y
                            if new_img_h < 10:  # Minimum height
                                new_img_h = 10
                            scribus.sizeObject(img_w, new_img_h, img_frame)
                            scribus.setScaleImageToFrame(True, True, img_frame)
                            img_h = new_img_h
                    except:
                        pass

                    # Strict boundary enforcement
                    simple_constrain_element(img_frame)

                    # Enable text wrap around the image
                    try:
                        scribus.setTextFlowMode(img_frame, TEXT_FLOW_OBJECTBOUNDINGBOX)
                    except:
                        pass

                    # Get actual frame height after scaling
                    try:
                        actual_w, actual_h = scribus.getSize(img_frame)
                    except:
                        actual_h = new_h
                    
                    # Move down for next image (add spacing between images)
                    current_y += actual_h + BLOCK_SPACING

            placed_images = True

        # Calculate the height of content for this frame
        used_h = measure_text_height(chunk, frame_w, in_template, default_font_size)
        
        # Add one line of space at the bottom to prevent text from being cut off
        line_height = default_font_size * 1.2  # Approximate line height
        used_h += line_height
        
        # If we have images, ensure frame is tall enough for both images AND text flowing around them
        if placed_images:
            # Calculate total height for images in single column
            total_image_height = 0
            safety_margin = 5  # Extra margin to ensure image stays within bounds
            max_available_width = frame_w_actual - left_pad - right_pad - safety_margin
            # Use smaller percentage (15%) and enforce maximum width of 80 points
            col_width = min(frame_w_actual * 0.15, max_available_width, 80)  # Use 15% of frame width, max 80pt

            for img_name in image_list:
                img_path = os.path.join(base_path, img_name)
                # Get real image dimensions
                if Image:
                    try:
                        with Image.open(img_path) as im:
                            orig_w, orig_h = im.size
                    except:
                        orig_w, orig_h = (300, 200)
                else:
                    orig_w, orig_h = (300, 200)

                # Scale to fit width while maintaining aspect ratio
                scale = float(col_width) / float(orig_w) if orig_w else 1.0
                new_h = orig_h * scale

                # Add height for this image with spacing
                total_image_height += new_h + BLOCK_SPACING

            # When text flows around images, it needs extra height
            # Calculate additional height needed based on the amount of horizontal space taken by images
            # Images take up 25% of width, so text needs roughly 33% more height (1/3 more)
            text_height_adjustment = used_h * 0.33
            adjusted_text_height = used_h + text_height_adjustment

            # Use the maximum of total image height or adjusted text height
            if adjusted_text_height > total_image_height:
                used_h = adjusted_text_height
            else:
                used_h = total_image_height
            
        # Final overflow check using documentation-based approach
        try:
            # Force text refresh to ensure all styling is applied and show live updates
            scribus.redrawAll()

            # Official Scribus method for precise text fitting based on documentation
            scribus.layoutText(frame)

            # Calculate boundary limits before expansion
            frame_x, frame_y = scribus.getPosition(frame)
            minimal_buffer = 10
            max_allowed = PAGE_HEIGHT - MARGINS[3] - frame_y - minimal_buffer

            # Expand minimally to ensure all text is visible, but respect boundaries
            while scribus.textOverflows(frame):
                current_w, current_h = scribus.getSize(frame)
                new_height = current_h + 3

                # Stop if we'd exceed the boundary
                if new_height > max_allowed:
                    break

                scribus.sizeObject(current_w, new_height, frame)
                scribus.layoutText(frame)

            # Calculate exact height using official Scribus methods
            try:
                # Get actual text metrics
                num_lines = scribus.getTextLines(frame)
                line_spacing = scribus.getLineSpacing(frame)
                left, right, top, bottom = scribus.getTextDistances(frame)

                if num_lines > 0:
                    # Calculate exact height: (lines × spacing) + padding
                    exact_text_height = num_lines * line_spacing
                    exact_frame_height = exact_text_height + top + bottom

                    # For frames with images, add extra space for text flow
                    if placed_images:
                        # Add more space for complex template image layouts
                        exact_frame_height += default_font_size * 2  # Add two lines for better image flow

                    # Calculate maximum allowed height with minimal buffer
                    frame_x, frame_y = scribus.getPosition(frame)
                    minimal_buffer = 10
                    max_allowed = PAGE_HEIGHT - MARGINS[3] - frame_y - minimal_buffer

                    # Resize to exact height, but respect page boundaries
                    current_w, current_h = scribus.getSize(frame)
                    safe_height = min(exact_frame_height, max_allowed)
                    scribus.sizeObject(current_w, safe_height, frame)
                    scribus.layoutText(frame)

                    # Verify no overflow after exact sizing (only expand if within boundaries)
                    if scribus.textOverflows(frame) and safe_height < max_allowed:
                        # For frames with images, may need more aggressive expansion
                        if placed_images:
                            # Add more space for complex image flow
                            expanded_height = min(exact_frame_height + line_spacing * 2, max_allowed)
                            scribus.sizeObject(current_w, expanded_height, frame)
                        else:
                            # Add minimal space for regular text
                            expanded_height = min(exact_frame_height + line_spacing * 0.1, max_allowed)
                            scribus.sizeObject(current_w, expanded_height, frame)
                        scribus.layoutText(frame)

                        # Final check - if still overflowing, expand more (but stay within boundaries)
                        if scribus.textOverflows(frame) and expanded_height < max_allowed:
                            final_height = min(exact_frame_height + line_spacing * 3, max_allowed)
                            scribus.sizeObject(current_w, final_height, frame)
                            scribus.layoutText(frame)
            except:
                # If official method fails, minimal fallback
                pass
        except:
            # Basic fallback overflow handling if documentation approach fails completely
            fallback_iterations = 0
            frame_x, frame_y = scribus.getPosition(frame)
            minimal_buffer = 10
            max_allowed = PAGE_HEIGHT - MARGINS[3] - frame_y - minimal_buffer

            while scribus.textOverflows(frame) and fallback_iterations < 10:
                try:
                    current_w, current_h = scribus.getSize(frame)
                    new_height = current_h + default_font_size

                    # Stop expanding if we'd exceed boundaries
                    if new_height > max_allowed:
                        break

                    scribus.sizeObject(current_w, new_height, frame)
                    fallback_iterations += 1
                except:
                    break

        # Final check - if text still overflows after all attempts to fit it
        # This only happens when frame is at maximum allowed size
        if scribus.textOverflows(frame):
            # Frame is at maximum size but text still overflows - need to split
            # Check if we've reached the page boundary
            frame_x, frame_y = scribus.getPosition(frame)
            current_w, current_h = scribus.getSize(frame)
            minimal_buffer = 10
            max_allowed = PAGE_HEIGHT - MARGINS[3] - frame_y - minimal_buffer

            # Only split if frame is at or near maximum height
            if current_h >= max_allowed - 5:  # Within 5 points of max
                # Use find_fit to determine how much text actually fits
                actual_fit = find_fit(remaining, frame)
                if actual_fit < len(remaining):
                    # Split the text - keep what fits, send rest to next page
                    chunk = remaining[:actual_fit]
                    chunk_len = actual_fit

                    # Re-apply the text for the truncated portion
                    scribus.setText(chunk, frame)

                    # Re-apply font settings after text change
                    try:
                        scribus.setFont(DEFAULT_FONT, frame)
                        scribus.setFontSize(default_font_size, frame)
                    except:
                        for f in FONT_CANDIDATES:
                            try:
                                scribus.setFont(f, frame)
                                scribus.setFontSize(default_font_size, frame)
                                break
                            except:
                                pass

                    # Re-apply text segments for the truncated text
                    for start, txt, seg_sty in seg_index:
                        end = start + len(txt)
                        ov_s = max(start, base_offset)
                        ov_e = min(end, base_offset + chunk_len)
                        if ov_e <= ov_s:
                            continue
                        rel_off = ov_s - base_offset
                        rel_len = ov_e - ov_s
                        try:
                            scribus.selectText(rel_off, rel_len, frame)
                            if "font_size" in seg_sty:
                                size_str = seg_sty["font_size"]
                                if "%" in size_str:
                                    pct = float(re.sub(r'[^\d.]','', size_str)) / 100.0
                                    sz = int(default_font_size * pct)
                                else:
                                    original_size = float(re.sub(r'[^\d.]','', size_str))
                                    sz = int(original_size * 0.9)
                                    sz = max(sz, 7)
                                scribus.setFontSize(sz, frame)
                        except:
                            pass

        # Use actual frame height for y_offset calculation
        try:
            actual_frame_height = scribus.getSize(frame)[1]
            y_offset += actual_frame_height + BLOCK_SPACING
        except:
            # Fallback to estimated height if getting actual height fails
            y_offset += used_h + BLOCK_SPACING
        enforce_margin_boundary()
        remaining = remaining[chunk_len:]
        base_offset += chunk_len
        
        if remaining:
            # Store if we placed images on the first page
            had_images = placed_images
            # Create new page for continued content
            new_page()

            # Optimize page break by using minimal top spacing for continued content
            # This eliminates large gaps when text flows to the next page
            y_offset = MARGINS[1] + BLOCK_SPACING  # Use standard block spacing at top of new page

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── QUIZ PLACEMENT FUNCTIONS ────────────
# ────────────────────────────────────────────────────────────────────────────────
def strip_html_tags(text):
    """Remove all HTML tags from a string."""
    if not text or not isinstance(text, str):
        return text
    return re.sub(r'<[^>]+>', '', text)

def place_quiz(arr, in_template=True, group_image=None, base_path=None):
    global y_offset, CURRENT_COLOR
    global quiz_heading_placed_on_page
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
            qa['que'] = strip_html_tags(qa['que'])
        if 'ans' in qa:
            qa['ans'] = strip_html_tags(qa['ans'])
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
    header_height = 24  # Reduced blue header height
    row_height = 16  # Increased row height for larger 8pt font
    answer_box_width = 18  # V/F box width
    answer_box_height = 16  # V/F box height increased for 10pt font
    answer_box_gap = 3  # Gap between V and F boxes
    # Draw blue header like copy 6 (from quiz_from_csv.py)
    if not quiz_heading_placed_on_page:
        # Define colors if not exists
        try:
            scribus.defineColor("Cyan", 0, 160, 224)  # Blue color
            scribus.defineColor("Yellow", 255, 255, 0)
            scribus.defineColor("NumBoxBlue", 210, 235, 255)
        except:
            pass
        
        # Create blue header background
        header_bg = scribus.createRect(MARGINS[0], y_offset, quiz_width, header_height)
        scribus.setFillColor("Cyan", header_bg)
        scribus.setLineColor("Cyan", header_bg)
        
        # Quiz header text - extended to fill the gap (removed yellow box)
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

        # Yellow page number box removed - quiz header now fills the full width

        quiz_heading_placed_on_page = True
        y_offset += header_height + 1  # Minimal gap after header
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
        
        # Special handling for borderline cases (text near one line) - copy6.py original
        if len(text) > chars_per_line * 0.85 and len(text) <= chars_per_line * 1.2:
            # Text is close to or slightly over one line - copy6.py original
            return default_height * 1.5  # Copy6.py original 50% more height
        
        # Calculate required height for true multi-line content - reduced for compact spacing
        line_height = font_size * 1.2  # Reduced line height for more compact layout
        calculated_height = lines_needed * line_height + 4  # Reduced padding for line spacing

        # Add minimal buffer for safety - reduced from original
        calculated_height = calculated_height * 1.05  # Reduced from 10% to 5% buffer
        
        # Return calculated height
        return max(default_height, calculated_height)
    
    # Process each question as a table row (from quiz_from_csv.py)
    for idx, qa in enumerate(filtered_arr):
        question = qa.get('que', '')
        is_true = qa.get('is_true', False)
        
        # Keep it very simple - just remove HTML tags and show plain text
        formatted_question = re.sub(r'<[^>]+>', '', question)
        cleaned_question_for_calc = formatted_question
        
        # Apply superscript conversion for height calculation too
        display_question_for_calc = re.sub(r'cm(\d)', lambda m: 'cm' + {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'}[m.group(1)], cleaned_question_for_calc)
        
        # Calculate row height using corrected text width (matches the fixed positioning)
        text_width = quiz_width - 42  # Adjusted to match the new text box positioning
        current_row_height = check_text_overflow(display_question_for_calc, text_width, 8, row_height, is_header=False)
        
        # Simple boundary check
        simple_boundary_check(current_row_height)
        
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
        # scribus.setText("", num_box)  # Empty - no numbers
        # No number box needed - removed completely
        # try:
        #     scribus.setFont(DEFAULT_FONT, num_box)
        # except:
        #     pass
        # scribus.setFontSize(6, num_box)
        # scribus.setTextAlignment(1, num_box)
        # scribus.setTextVerticalAlignment(1, num_box)
        # scribus.setTextDistances(0, 0, 0, 0, num_box)
        # scribus.setTextColor("Black", num_box)
        
        # Answer text box - ensure it stays within margins
        text_start_x = MARGINS[0] + 2
        text_end_x = MARGINS[0] + quiz_width - 40  # Leave space for V/F boxes (38 + 2 margin)
        text_width = text_end_x - text_start_x
        text_box_bg = scribus.createRect(text_start_x, y_offset, text_width, current_row_height - 1)
        
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
        
        # Answer text - properly centered vertically and horizontally like copy6.py
        text_box_height = current_row_height - 2  # Use almost full row height with 1pt padding top/bottom
        text_y_offset_answer = 1  # Minimal top padding
        q_frame = scribus.createText(text_start_x + 1, y_offset + text_y_offset_answer, text_width - 2, text_box_height)
        
        # Use the same converted text for display as we used for height calculation
        display_question = display_question_for_calc
        scribus.setText(display_question, q_frame)
        
        # Use quiz font
        try:
            scribus.setFont(QUIZ_ACTUAL_FONT, q_frame)
            scribus.setFontSize(9, q_frame)
        except:
            try:
                scribus.setFont(DEFAULT_FONT, q_frame)
                scribus.setFontSize(9, q_frame)
            except:
                pass
        scribus.setTextColor("Black", q_frame)
        # Enable proper text alignment and centering like copy6.py
        try:
            scribus.setTextDistances(0, 0, 0, 0, q_frame)  # No padding for perfect centering like V/F boxes
            # Set line spacing based on row height - increased for better readability
            if current_row_height > 14:
                scribus.setLineSpacing(9, q_frame)  # Increased spacing for multi-line
            else:
                scribus.setLineSpacing(8, q_frame)  # Increased spacing for single line
            
            # Set horizontal alignment
            scribus.setTextAlignment(0, q_frame)  # Left align
            
            # Try to set vertical alignment to middle like copy6.py
            try:
                scribus.setTextVerticalAlignment(1, q_frame)  # 1 = middle alignment
            except:
                try:
                    # Alternative method for vertical centering from copy6.py
                    scribus.setTextBehaviour(q_frame, 1)  # Try different behavior
                except:
                    pass
            
            # Force text to stay within bounds - comprehensive copy6.py approach
            try:
                scribus.setTextBehaviour(q_frame, 0)  # Force text in frame
            except:
                pass
            
            # Additional overflow protection from copy6.py
            try:
                scribus.setTextToFrameOverflow(q_frame, False)  # Disable overflow
            except:
                pass
            
            # Enable text wrapping like copy6.py
            try:
                scribus.setTextFlowMode(q_frame, 0)  # Enable text flow
            except:
                pass
        except:
            pass
        # V/F checkboxes exactly from copy6.py
        # Define checkbox colors from copy6.py
        try:
            scribus.defineColor("CheckBoxColor", 240, 255, 240)  # Light green tint for V
            scribus.defineColor("CheckBoxColor2", 255, 240, 240)  # Light red tint for F
        except:
            pass
        
        # V checkbox box - adjusted position since no number box
        v_box_bg = scribus.createRect(MARGINS[0] + quiz_width - 38, y_offset, 18, current_row_height - 1)
        try:
            scribus.setFillColor("CheckBoxColor", v_box_bg)
        except:
            scribus.setFillColor("White", v_box_bg)
        scribus.setLineColor("Cyan", v_box_bg)
        scribus.setLineWidth(0.5, v_box_bg)
        
        # V checkbox text - adjusted position
        checkbox_box_height = current_row_height - 2  # Use almost full row height with 1pt padding
        checkbox_y_offset = 1  # Minimal top padding
        v_box = scribus.createText(MARGINS[0] + quiz_width - 38, y_offset + checkbox_y_offset, 18, checkbox_box_height)
        scribus.setText("V", v_box)
        try:
            scribus.setFontSize(10, v_box)  # Increased V/F box font size
            scribus.setTextAlignment(1, v_box)  # Center align horizontally
            # Try to set vertical alignment to middle
            try:
                scribus.setTextVerticalAlignment(1, v_box)  # 1 = middle alignment
            except:
                pass
            # Set proper text distances for centering
            scribus.setTextDistances(0, 0, 0, 0, v_box)  # No padding for perfect centering
        except:
            pass
        # Color V with header color if it's the correct answer (from copy6.py)
        if is_true:
            scribus.setTextColor("Cyan", v_box)
        else:
            scribus.setTextColor("Black", v_box)
        
        # F checkbox box - adjusted position
        f_box_bg = scribus.createRect(MARGINS[0] + quiz_width - 18, y_offset, 18, current_row_height - 1)
        try:
            scribus.setFillColor("CheckBoxColor2", f_box_bg)
        except:
            scribus.setFillColor("White", f_box_bg)
        scribus.setLineColor("Cyan", f_box_bg)
        scribus.setLineWidth(0.5, f_box_bg)
        
        # F checkbox text - adjusted position
        f_box = scribus.createText(MARGINS[0] + quiz_width - 18, y_offset + checkbox_y_offset, 18, checkbox_box_height)
        scribus.setText("F", f_box)
        try:
            scribus.setFontSize(10, f_box)  # Increased V/F box font size
            scribus.setTextAlignment(1, f_box)  # Center align horizontally
            # Try to set vertical alignment to middle
            try:
                scribus.setTextVerticalAlignment(1, f_box)  # 1 = middle alignment
            except:
                pass
            # Set proper text distances for centering
            scribus.setTextDistances(0, 0, 0, 0, f_box)  # No padding for perfect centering
        except:
            pass
        # Color F with header color if it's the correct answer (from copy6.py)
        if not is_true:  # F is correct when is_true is False
            scribus.setTextColor("Cyan", f_box)
        else:
            scribus.setTextColor("Black", f_box)
        
        # Simple boundary enforcement for quiz elements
        simple_constrain_element(q_frame)
        simple_constrain_element(text_box_bg)
        
        # Move to next row
        y_offset += current_row_height
        enforce_margin_boundary()
    
    # Add minimal spacing after quiz section only if space permits (with page number buffer)
    safe_boundary = PAGE_HEIGHT - MARGINS[3] - 22  # Page number buffer
    if y_offset + 1 < safe_boundary:
        y_offset += 1  # Minimal gap after quiz section

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── TEMPLATE PROCESSING ────────────
# ────────────────────────────────────────────────────────────────────────────────
def place_images_grid(images, base_path, start_y, max_height=None, max_images=None):
    """
    Place images in a grid pattern where:
    - First 3 images are in a row at the top
    - 4th and 5th images are placed below centered
    - Pattern repeats for 6th image onwards
    
    This only applies to regular images, not roadsigns.
    All images in a group will have the same height for visual consistency.
    
    Args:
        images: List of image paths
        base_path: Base path for images
        start_y: Starting Y coordinate
        max_height: Maximum height constraint (optional)
        max_images: Maximum number of images to place (optional, for partial placement)
    
    Returns:
        The final Y position after placing images
    """
    global y_offset
    
    if not images:
        return start_y
    
    # If max_images is specified, only place that many images
    if max_images and max_images > 0:
        images = images[:max_images]
    
    # Set default target height
    target_height = min(150, max_height) if max_height else 150
    
    # Available width for images
    available_width = PAGE_WIDTH - MARGINS[0] - MARGINS[2]
    
    # Divide images into groups of 5 for the grid pattern
    image_groups = []
    for i in range(0, len(images), 5):
        group = images[i:i+5]
        image_groups.append(group)
    
    current_y = start_y
    
    # Process each group
    for group in image_groups:
        # First gather all image aspect ratios and determine one consistent height
        all_widths = []
        all_aspect_ratios = []
        
        # Calculate aspect ratio for all images in the group
        for rel in group:
            img_path = os.path.join(base_path, rel)
            if Image:
                try:
                    with Image.open(img_path) as im:
                        orig_w, orig_h = im.size
                except:
                    orig_w, orig_h = (300, 200)
            else:
                orig_w, orig_h = (300, 200)
            
            aspect_ratio = float(orig_w) / float(orig_h) if orig_h else 1.5
            all_aspect_ratios.append(aspect_ratio)
        
        # Calculate initial widths based on target height
        for ratio in all_aspect_ratios:
            width = ratio * target_height
            all_widths.append(width)
        
        # First row - up to 3 images side by side
        first_row = group[:min(3, len(group))]
        first_row_widths = all_widths[:min(3, len(group))]
        
        # Check if first row fits, adjust height if needed
        if first_row:
            first_row_total_width = sum(first_row_widths) + (len(first_row_widths) - 1) * BLOCK_SPACING
            
            # Second row - up to 2 images
            second_row = group[3:min(5, len(group))] if len(group) > 3 else []
            second_row_widths = all_widths[3:min(5, len(group))] if len(all_widths) > 3 else []
            
            if second_row:
                second_row_total_width = sum(second_row_widths) + (len(second_row_widths) - 1) * BLOCK_SPACING
                # Use the row that requires the most scaling as the limiting factor
                row1_ratio = available_width / first_row_total_width if first_row_total_width > available_width else 1.0
                row2_ratio = available_width / second_row_total_width if second_row_total_width > available_width else 1.0
                scaling_ratio = min(row1_ratio, row2_ratio)
            else:
                # Only first row needs to be considered
                scaling_ratio = available_width / first_row_total_width if first_row_total_width > available_width else 1.0
            
            # Apply scaling to all images if needed
            adjusted_height = target_height * scaling_ratio
            all_widths = [width * scaling_ratio for width in all_widths]
        else:
            adjusted_height = target_height
        
        # First row - place images side by side
        if first_row:
            # Calculate starting position for the first row (left-aligned)
            x = MARGINS[0]
            
            # Place first row images
            for idx, rel in enumerate(first_row):
                img_path = os.path.join(base_path, rel)
                w_i = all_widths[idx]
                h_i = adjusted_height
                img_frame = scribus.createImage(x, current_y, w_i, h_i)
                scribus.loadImage(img_path, img_frame)
                scribus.setScaleImageToFrame(True, True, img_frame)
                scribus.setLineColor("None", img_frame)

                # Try to eliminate gaps - test without custom function
                try:
                    scribus.setScaleFrameToImage(img_frame)
                except:
                    pass

                # Strict boundary enforcement for image
                simple_constrain_element(img_frame)
                x += w_i + BLOCK_SPACING
            
            # Update y position after first row
            current_y += adjusted_height + BLOCK_SPACING
        
        # Second row - up to 2 images centered
        second_row = group[3:min(5, len(group))] if len(group) > 3 else []
        if second_row:
            # Use the same height as the first row for consistency
            second_row_widths = all_widths[3:min(5, len(group))]
            second_row_total_width = sum(second_row_widths) + (len(second_row_widths) - 1) * BLOCK_SPACING
            
            # Calculate starting position for the second row (centered)
            x = MARGINS[0] + (available_width - second_row_total_width) / 2
            
            # Place second row images
            for idx, rel in enumerate(second_row):
                img_path = os.path.join(base_path, rel)
                w_i = all_widths[idx + 3]  # Offset by 3 for second row
                h_i = adjusted_height
                img_frame = scribus.createImage(x, current_y, w_i, h_i)
                scribus.loadImage(img_path, img_frame)
                scribus.setScaleImageToFrame(True, True, img_frame)
                scribus.setLineColor("None", img_frame)

                # Try to eliminate gaps - test without custom function
                try:
                    scribus.setScaleFrameToImage(img_frame)
                except:
                    pass

                # Strict boundary enforcement for image
                simple_constrain_element(img_frame)
                x += w_i + BLOCK_SPACING
            
            # Update y position after second row
            current_y += adjusted_height + BLOCK_SPACING
    
    # Return the final y position
    return current_y

# Place these two functions before process_template
def measure_quiz_group_height(group, group_image, base_path):
    # This function mimics the height calculation logic in place_quiz, but does not place anything.
    global y_offset  # Declare global variable
    quiz_header_height = 34
    card_spacing = 1  # Minimal spacing between cards
    answer_box_width = 25
    answer_box_height = 16
    quiz_width = PAGE_WIDTH - MARGINS[0] - MARGINS[2]
    answer_box_gap = 4
    quiz_bar_height = 26
    padding_top = 1  # Minimal padding
    padding_bottom = 1  # Minimal padding
    image_height = 32
    img_margin = 6
    image_gap = 8
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
        while scribus.textOverflows(temp_frame):
            w, h = scribus.getSize(temp_frame)
            scribus.sizeObject(w, h + 10, temp_frame)
        q_height = scribus.getSize(temp_frame)[1]
        scribus.deleteObject(temp_frame)
        question_heights.append(q_height)
    # Match copy 6's content padding
    card_top_margin = 2  # Minimal space before each quiz item
    padding_top = 1  # Minimal padding
    padding_bottom = 1  # Minimal padding
    
    # Instead of modifying y_offset directly, include its value in the calculation
    
    # Card height calculation
    total_questions_height = sum([h + padding_top + padding_bottom for h in question_heights]) + (card_spacing * (len(group)-1))
    card_height = max(answer_box_height + padding_top + padding_bottom, total_questions_height, img_h + padding_top + padding_bottom)
    
    # Add space for the heading (QUIZ bar)
    total_height = (quiz_bar_height + 1) + card_height + card_top_margin + 1
    return total_height

def group_quiz_by_image(quiz_entries):
    """
    Groups quiz entries by their associated images.
    Since quiz entries don't directly contain image info, we group all entries together under None.
    This maintains compatibility with the existing place_quiz_group_paginated function.
    """
    if not quiz_entries:
        return {}
    
    # For now, group all quiz entries under a single key (None)
    # This maintains the expected structure while allowing future image-based grouping
    return {None: quiz_entries}

def place_quiz_group_paginated(group, group_image, base_path):
    global y_offset
    if not group:
        return
    idx = 0
    n = len(group)
    while idx < n:
        # Try to fit as many questions as possible on this page
        best_end = idx + 1
        for end in range(n, idx, -1):
            height_needed = measure_quiz_group_height(group[idx:end], group_image, base_path)
            reduced_gap = 30
            available = PAGE_HEIGHT - y_offset - MARGINS[3] - 25  # Page number buffer
            if y_offset > MARGINS[1]:
                height_needed -= reduced_gap
            if height_needed <= available:
                best_end = end
                break
        # Avoid leaving a single question alone at the top of a page (unless group is size 1)
        if best_end == idx + 1 and (n - idx) > 1:
            # Not enough space for more than one, so move at least two to next page
            new_page()
            continue
        place_quiz(group[idx:best_end], True, group_image, base_path)
        idx = best_end
        if idx < n:
            new_page()

def process_template(tmpl, base_path):
    global y_offset, CURRENT_COLOR, global_template_count
    # Check if we have enough space for at least the template header (with page number buffer)
    # If not, start a new page
    safe_boundary = PAGE_HEIGHT - MARGINS[3] - 22  # Page number buffer
    if y_offset + 10 > safe_boundary:
        new_page()
    else:
        # Add padding before the template only if we're not at the top of a page
        if y_offset > MARGINS[1]:
            y_offset += TEMPLATE_PADDING / 2  # Use half the normal padding
            enforce_margin_boundary()
    
    # Set color based on template ID (for text styling, but no background)
    tid = tmpl.get("id","0")
    try:
        idx = int(tid) % len(BACKGROUND_COLORS)
    except:
        idx = hash(tid) % len(BACKGROUND_COLORS)
    CURRENT_COLOR = BACKGROUND_COLORS[idx]
    
    # Get text content
    txt = tmpl.get("text", [])
    txt = txt if isinstance(txt, list) else ([txt] if txt else [])
    
    # Aggressively clean text items
    cleaned_txt = []
    for item in txt:
        if item:  # Skip None or empty items
            # Convert to string and strip all whitespace
            item_str = str(item).strip()
            if item_str:  # Only add non-empty strings
                cleaned_txt.append(item_str)
    
    # Get template images
    imgs = tmpl.get("images", [])
    imgs = imgs if isinstance(imgs, list) else ([imgs] if imgs else [])

    # Get roadsigns
    rs = tmpl.get("roadsigns", [])
    rs = rs if isinstance(rs, list) else ([rs] if rs else [])

    # Combine images and roadsigns for placement within text frame
    # Place template images first (for left side), then roadsigns (for right side)
    all_images = imgs + rs if (imgs and rs) else (imgs if imgs else rs)

    # Determine alignment for images based on template ID for variety
    # Alternate between center, top, and bottom alignments
    alignments = ["C", "TC", "BC"]
    try:
        aid = int(tid) % len(alignments)
    except:
        aid = 0
    template_alignment = alignments[aid]

    # Place text with images/roadsigns integrated in single column
    if cleaned_txt or all_images:
        place_wrapped_text_and_images(cleaned_txt, all_images, base_path, 8, True, template_alignment)
    
    # Place videos
    for v in tmpl.get("videos", []):
        # Make video text always visible
        video_text = f"[video: {v}]"
        place_text_block_flow(video_text, 8, False, True)
    
    # Images are now placed within text frames by place_wrapped_text_and_images()
    # Old separate image placement code removed to prevent duplicates
    
    # Place quiz section using global constants - only if quizzes are enabled
    if "quiz" in tmpl and PRINT_QUIZZES:
        # Reset quiz header flag for this template - allow ONE header per template
        global quiz_heading_placed_on_page
        quiz_heading_placed_on_page = False
        
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
    
    # Ultra-minimal spacing between templates (with page number buffer)
    remaining_space = PAGE_HEIGHT - MARGINS[3] - 22 - y_offset  # Page number buffer
    if remaining_space > 2:
        y_offset += 1  # Ultra-minimal 1-point padding between templates
        enforce_margin_boundary()
    
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
        try:
            # Colors are pre-defined in define_banner_colors()
            scribus.setFillColor(current_topic_color, rect)
        except:
            # Fallback to Red if color not available
            try:
                scribus.setFillColor("Red", rect)
            except:
                pass
    scribus.setLineColor("None", rect)

    # Prepare and place banner text
    display_text = " ".join(current_topic_text.upper())
    banner_center_y = banner_y + (rect_height / 2)
    text_height = rect_height * BANNER_TEXT_HEIGHT_PERCENT

    try:
        text_frame = scribus.createText(0, 0, text_height, TOPIC_BANNER_WIDTH - 4)
        scribus.setText(display_text, text_frame)
        
        # Apply font settings
        font_applied = False
        for myriad_font in MYRIAD_VARIANTS:
            try:
                scribus.setFont(myriad_font, text_frame)
                scribus.setFontSize(BANNER_TEXT_FONT_SIZE)
                font_applied = True
                break
            except:
                continue
                
        # Apply text properties
        scribus.setTextAlignment(scribus.ALIGN_CENTERED, text_frame)
        scribus.setTextColor("White", text_frame)
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
    Superscript and subscript handling function for HTML tagged content and special patterns.
    Handles:
    - <span class="S-T...">digits</span> (superscripts)
    - <sup>digits</sup> (superscripts)  
    - <sub>digits</sub> (subscripts)
    - Complex HTML span patterns with style attributes
    """
    if not text or not isinstance(text, str):
        return text
    
    original_text = text
    
    # Handle special HTML span pattern for digit superscripts with any unit
    text = re.sub(
        r'([a-zA-Z]+)</span><span\s+class=["\']S-T\d+["\']\s+style=["\'][^"\']*vertical-align[^"\']*["\']>\s*(\d+)\s*</span>',
        lambda m: m.group(1) + ''.join(_SUP_MAP.get(ch, ch) for ch in m.group(2)),
        text
    )
    
    # Handle text followed by S-T span (e.g., "cm<span class="S-T18">3</span>") - FIRST
    # For quiz questions, convert to simple format that apply_quiz_superscripts can handle
    text = re.sub(
        r'([a-zA-Z]+)<span\s+class=(?:["\']|\\")S-T[^"\'>]*(?:["\']|\\")(?:\s*[^>]*)?>(\d+)</span>',
        lambda m: m.group(1) + '<sup>' + m.group(2) + '</sup>',
        text
    )
    
    # Also handle the simple case - convert to <sup> format
    text = re.sub(
        r'<span\s+class=(?:["\']|\\")S-T[^"\'>]*(?:["\']|\\")(?:\s*[^>]*)?>(\d+)</span>',
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
def create_styled_header(text, font_size, bold, bg_color, text_color, padding):
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
    # Use reduced padding for template headers
    actual_padding = 5 if is_template_header else padding
    # Add spacing BEFORE chapter headers (non-template headers)
    if not is_template_header:
        y_offset += BLOCK_SPACING * 6  # Add 18 points spacing before chapter headers
        
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
    
    # Create text frame - ALL headers use SAME width (frame_w) for consistent alignment
    # All headers: position at MARGINS[0] with full frame_w width
    text_x = MARGINS[0]
    text_y = y_offset
    text_w = frame_w
    text_h_frame = text_h
    
    text_frame = scribus.createText(text_x, text_y, text_w, text_h_frame)
    scribus.setText(text, text_frame)
    
    # No internal padding - text starts right at the frame edge (at MARGINS[0])
    try:
        scribus.setTextDistances(0, 0, 0, 0, text_frame)
    except:
        pass
    
    try:
        scribus.setLineColor("None", text_frame)
    except:
        pass
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
                for f in ["Arial", "Times", "Courier"]:
                    try:
                        scribus.setFont(f, text_frame)
                        scribus.setFontSize(font_size, text_frame)
                        scribus.textColor(text_color, text_frame)
                        break
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
            template_font_size = 7  # Default template font size
            scribus.setLineSpacing(template_font_size * 1.2, text_frame)
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
            while scribus.textOverflows(text_frame):
                current_w, current_h = scribus.getSize(text_frame)
                # Check if expanding would exceed bottom margin (with minimal buffer)
                minimal_buffer = 10  # Just enough for page numbers
                max_allowed_height = (PAGE_HEIGHT - MARGINS[3] - minimal_buffer) - y_offset  # Minimal buffer
                if current_h + 3 > max_allowed_height:
                    break  # Don't expand if it would exceed margins
                scribus.sizeObject(current_w, current_h + 3, text_frame)
                scribus.layoutText(text_frame)

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

        # Keep all header/title text frames at full width for consistent alignment
        # Only adjust width for non-header content (identified by bold=False and no special formatting)
        # For now, keep all text frames at their created width (full width)
        if bg_rect:
            # With background: text frame is at frame_w (full width)
            # Keep it as is
            pass
        else:
            # Without background: text frame is at frame_w
            # Keep full width for consistent header alignment
            pass
    spacing_after = BLOCK_SPACING * 0.25 if is_template_header else BLOCK_SPACING  # Normal spacing after headers
    y_offset += text_h + spacing_after
    
    # Simple final constraint
    simple_constrain_element(text_frame)
    
    # Text overflow is already handled by the documentation-based approach above
    
    # Enforce strict margin boundary
    enforce_margin_boundary()
    
    return text_frame

def create_module_header(module_name):
    """Create a module header with the configured style."""
    return create_styled_header(
        module_name, 
        MODULE_FONT_SIZE, 
        MODULE_BOLD, 
        MODULE_BG_COLOR, 
        MODULE_TEXT_COLOR, 
        MODULE_PADDING
    )

def create_topic_header(topic_name):
    """Create a topic header with the configured style."""
    return create_styled_header(
        topic_name, 
        TOPIC_FONT_SIZE,
        TOPIC_BOLD,
        TOPIC_BG_COLOR,
        TOPIC_TEXT_COLOR,
        TOPIC_PADDING
    )

def handle_text_styles(frame, style_segments, default_size):
    """Apply text styles based on parsed style segments."""
    # Skip empty text or if frame is not valid
    if not style_segments or not frame:
        return
    
    # Set default font and size first before applying specific styles
    try:
        scribus.setFont(DEFAULT_FONT, frame)
        scribus.setFontSize(default_size, frame)
    except:
        # Try font candidates if DEFAULT_FONT fails
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
            
            # Apply font family if specified in style
            font_applied = False
            if "font" in style_dict:
                # First try the detected full font name
                try:
                    scribus.setFont(style_dict["font"], frame)
                    font_applied = True
                except:
                    # If that fails, use the font detection logic with any available style info
                    try:
                        font_family = style_dict.get("font-family", style_dict["font"])
                        font_weight = style_dict.get("font-weight", None)
                        font_style = style_dict.get("font-style", None)
                        
                        # Apply bold/italic flags
                        if style_dict.get("bold", False) and not font_weight:
                            font_weight = "bold"
                        if style_dict.get("italic", False) and not font_style:
                            font_style = "italic"
                            
                        # Get best matching font = get_font_with_style(font_family, font_weight, font_style)
                        best_font = get_font_with_style(font_family, font_weight, font_style)
                        
                        # Apply the font
                        scribus.setFont(best_font, frame)
                        font_applied = True
                    except Exception as e:
                        # If all else fails, try the DEFAULT_FONT
                        try:
                            scribus.setFont(DEFAULT_FONT, frame)
                            font_applied = True
                        except:
                            pass
            
            # If no font was specified or applied yet, try to set best font based on style attributes
            if not font_applied and (style_dict.get("bold", False) or style_dict.get("italic", False)):
                try:
                    current_font = None
                    try:
                        # Try to get current font
                        current_font = scribus.getFont(frame)
                    except:
                        current_font = DEFAULT_FONT
                        
                    # Apply styling to current font
                    font_weight = "bold" if style_dict.get("bold", False) else None
                    font_style = "italic" if style_dict.get("italic", False) else None
                    
                    best_font = get_font_with_style(current_font, font_weight, font_style)
                    scribus.setFont(best_font, frame)
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
                        # Scale down non-percentage sizes
                        size = size * 0.9
                    
                    # Ensure minimum readable size
                    size = max(size, 7)
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

def define_banner_colors():
    """Define custom colors for topic banners in the Scribus document."""
    try:
        # Define common colors that might be used in JSON banner_color field
        # Scribus uses CMYK color model
        custom_colors = {
            "Cyan": (100, 0, 0, 0),      # Cyan CMYK
            "Magenta": (0, 100, 0, 0),   # Magenta CMYK
            "Yellow": (0, 0, 100, 0),    # Yellow CMYK
            "Blue": (100, 100, 0, 0),    # Blue CMYK
            "Green": (100, 0, 100, 0),   # Green CMYK
            "Red": (0, 100, 100, 0),     # Red CMYK
        }

        for color_name, (c, m, y, k) in custom_colors.items():
            try:
                # Check if color already exists
                scribus.defineColor(color_name, c, m, y, k)
            except:
                # Color might already exist, skip
                pass
    except:
        # If color definition fails, continue anyway
        pass

def create_pages_from_json(json_path=None, include_quizzes=True, filter_mode="all"):
    """
    Create a PDF document from a JSON file containing the content structure.

    Parameters:
        json_path (str): Path to the JSON file
        include_quizzes (bool): Whether to include quizzes in the PDF
        filter_mode (str): Filter mode for quizzes - "all", "true_only", or "false_only"
    """
    global y_offset, global_template_count, limit_reached, current_topic_text, current_topic_color, PRINT_QUIZZES, QUIZ_FILTER_MODE

    # Define custom colors for banners
    define_banner_colors()

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
            
        create_styled_header(f"{area.get('name','Unnamed')}", 11, True, "None", "Black", 5)
        place_text_block_flow(area.get("desc",""), AREA_DESC_FONT_SIZE, no_bottom_gap=True)

        for chap in area.get("chapters", []):
            if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
                break

            create_styled_header(f"{chap.get('name','Unnamed')}", 11, True, "None", "Black", 5)

            for topic in chap.get("topics", []):
                if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
                    break

                # Clear previous topic banner and set new one
                current_topic_text = topic.get('name', 'Unnamed')
                banner_color = topic.get('banner_color', None)  # Read banner color from JSON
                create_topic_header(current_topic_text)
                add_vertical_topic_banner(current_topic_text, banner_color)

                # Display topic description
                place_text_block_flow(topic.get("desc",""), TOPIC_DESC_FONT_SIZE, no_bottom_gap=True)
                
                for mod in topic.get("modules", []):
                    if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
                        break
                        
                    create_module_header(mod.get('name','Unnamed'))
                    
                    for tmpl in mod.get("templates", []):
                        if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
                            break
                        process_template(tmpl, base_pics)

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
            scribus.messageBox("Done", f"Saved: {pdf_out}", scribus.ICON_INFORMATION)
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
