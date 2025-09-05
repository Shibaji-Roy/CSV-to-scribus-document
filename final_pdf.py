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
    """Find the maximum amount of text that fits in a frame, leaving one line of space at the bottom."""
    if not text:
        return 0
        
    # First, get the line height to know how much space to reserve at bottom
    one_line_height = 0
    try:
        # Try to get current font size as an approximation of line height
        font_size = scribus.getFontSize(frame)
        # Line height is typically ~120% of font size
        one_line_height = font_size * 1.2
    except:
        # Default to a reasonable line height if we can't get font size
        one_line_height = 12
    
    # Get current frame size
    frame_width, frame_height = scribus.getSize(frame)
    
    # Adjust frame height to leave one line of space at bottom
    adjusted_height = frame_height - one_line_height
    
    # If the adjusted height is too small, return empty fit
    if adjusted_height <= 0:
        return 0
    
    # Create a temporary invisible frame with the adjusted height
    temp_frame = scribus.createText(0, 0, frame_width, adjusted_height)
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
    """Calculate remaining vertical space on the current page."""
    return PAGE_HEIGHT - y_offset - MARGINS[3]

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

def measure_text_height(text, width, in_template=False):
    """
    Create a temporary text frame to measure the height needed for the given text.
    Returns the height in points.
    """
    probe = scribus.createText(MARGINS[0], y_offset, width, 20)
    
    # Remove the frame border
    try:
        scribus.setLineColor("None", probe)
    except:
        pass
    
    # Apply text distance padding for all text frames
    try:
        # Set internal padding (left, right, top, bottom)
        if in_template:
            scribus.setTextDistances(*TEMPLATE_TEXT_PADDING, probe)
        else:
            scribus.setTextDistances(*REGULAR_TEXT_PADDING, probe)
    except:
        pass
            
    # Set default font for measurement
    try:
        scribus.setFont(DEFAULT_FONT, probe)
    except:
        # If DEFAULT_FONT fails, try FONT_CANDIDATES as fallback
        for f in FONT_CANDIDATES:
            try:
                scribus.setFont(f, probe)
                break
            except:
                pass
                
    scribus.setText(text, probe)
    while scribus.textOverflows(probe):
        w, h = scribus.getSize(probe)
        scribus.sizeObject(w, h + 20, probe)
    final_h = scribus.getSize(probe)[1]
    scribus.deleteObject(probe)
    return final_h

def new_page():
    """Create a new page in the document and reset the y position."""
    global y_offset
    global quiz_heading_placed_on_page
    scribus.newPage(-1)
    y_offset = MARGINS[1]
    quiz_heading_placed_on_page = False
    # Add vertical topic banner to the new page if we have an active topic
    # This will automatically position it based on whether the new page is odd or even
    if current_topic_text:
        create_vertical_topic_banner()



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
    
    # Set default target height, smaller for roadsigns
    target_height = min(40, max_height) if max_height else 40
    
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
        
        # Right align for roadsigns (standard for roadsigns)
        x = PAGE_WIDTH - MARGINS[2] - total_width
        
        # Place images in this row
        for idx, rel in enumerate(row_images):
            img_path = os.path.join(base_path, rel)
            w_i = widths[idx]
            h_i = adjusted_height
            img_frame = scribus.createImage(x, current_y, w_i, h_i)
            scribus.loadImage(img_path, img_frame)
            scribus.setScaleImageToFrame(True, True, img_frame)
            scribus.setLineColor("None", img_frame)
            
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
        
        # Update y position after this row
        current_y += adjusted_height + BLOCK_SPACING
        
        # Move to next row
        row_start_idx = row_end_idx
    
    # Return the final y position
    return current_y

# ─────────── TEXT PLACEMENT FUNCTIONS ────────────
def place_text_block_flow(html_text, font_size=10, bold=False, in_template=False, no_bottom_gap=False):
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
    text_h = measure_text_height(plain, frame_w, in_template)
    
    # Add a small buffer for one line of text to ensure nothing gets cut off, unless no_bottom_gap is True
    if not no_bottom_gap:
        text_h += font_size * 1.2  # Add approximately one line of height
    
    # Use the more flexible content fitting logic
    if not can_fit_content(text_h):
        new_page()

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
    
    # Apply bold to entire frame if requested
    if bold:
        try: 
            current_size = scribus.getFontSize(frame)
            scribus.selectText(0, len(plain), frame)
            scribus.setFontSize(current_size + 2, frame)
        except: 
            pass

    # Trim the frame height to fit the content exactly
    try:
        actual_lines = scribus.getTextLines(frame)
        if actual_lines:
            # Calculate actual text height based on lines
            actual_height = sum(line_info[1] for line_info in actual_lines)
            # Add a little extra space to prevent cutoff
            actual_height += 2
            
            # Get current padding
            left, right, top, bottom = scribus.getTextDistances(frame)
            # Include padding in final height
            final_height = actual_height + top + bottom
            
            # Add one line of space at the bottom to ensure text isn't cut off, unless no_bottom_gap is True
            if not no_bottom_gap:
                line_height = font_size * 1.2  # Approximate line height
                final_height += line_height
            
            # Resize the frame to fit content exactly with the extra line buffer
            scribus.sizeObject(frame_w, final_height, frame)
            
            # Update text_h to the new height
            text_h = final_height
    except: 
        # If getting text lines fails, keep original height
        pass

    # Update y offset for next element
    y_offset += text_h + BLOCK_SPACING
    
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
                                default_font_size=10, in_template=False, alignment="C"):
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
        free_h = PAGE_HEIGHT - MARGINS[3] - y_offset
        if free_h < MIN_SPACE_THRESHOLD:
            new_page()
            free_h = PAGE_HEIGHT - MARGINS[3] - y_offset

        # Create a frame with room for text and images
        frame = scribus.createText(MARGINS[0], y_offset, frame_w, free_h)
        try: scribus.setLineColor("None", frame)
        except: pass
        try:
            if in_template:
                scribus.setTextDistances(*TEMPLATE_TEXT_PADDING, frame)
            else:
                scribus.setTextDistances(*REGULAR_TEXT_PADDING, frame)
        except: pass
        
        # Enable text flow mode for wrapping around images - EXACTLY like test2.py
        try:
            scribus.setTextFlowMode(frame, TEXT_FLOW_INTERACTIVE)
        except:
            pass

        scribus.setText(remaining, frame)
        if scribus.textOverflows(frame):
            cut = find_fit(remaining, frame)
            chunk = remaining[:cut]
            scribus.setText(chunk, frame)
            chunk_len = cut
        else:
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
                        sz = int(float(re.sub(r'[^\d.]','', size_str)))
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
            # Get minimum height needed for roadsigns (40pt from config)
            roadsign_height = 40  # This matches the target_height in place_roadsigns_grid
            
            # Calculate total height needed for all roadsigns
            # For every 2 roadsigns (1 row), we need roadsign_height + BLOCK_SPACING
            # Math.ceil to handle odd number of roadsigns
            rows_needed = (len(image_list) + 1) // 2  # Integer division rounded up
            total_roadsign_height = (roadsign_height * rows_needed) + (BLOCK_SPACING * (rows_needed - 1)) if rows_needed > 0 else 0
            
            # Add some padding to ensure proper spacing
            total_roadsign_height += 4
            
            # Get current frame height
            _, frame_h = scribus.getSize(frame)
            
            # Ensure frame is tall enough for all roadsigns
            if frame_h < total_roadsign_height:
                # Resize the frame to fit all roadsigns properly
                scribus.sizeObject(frame_w, total_roadsign_height, frame)
                frame_h = total_roadsign_height
            
            col_width = frame_w * 0.4  # Use 40% of frame width for images
            
            if len(image_list) == 1:
                # Single image with text wrap
                img_path = os.path.join(base_path, image_list[0])
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
                if new_h > frame_h:
                    scale = float(frame_h) / float(orig_h)
                    new_h = frame_h
                    new_w = orig_w * scale
                else:
                    new_w = col_width
                
                # Position in text frame - right aligned
                x = MARGINS[0] + frame_w - new_w
                
                img_frame = scribus.createImage(x, y_offset, new_w, new_h)
                scribus.loadImage(img_path, img_frame)
                scribus.setScaleImageToFrame(True, True, img_frame)
                scribus.setLineColor("None", img_frame)
                
                # Enable shaped text wrap
                if scribus.getObjectType(img_frame) == "ImageFrame":
                    try:
                        scribus.setItemShapeSetting(img_frame, scribus.ITEM_BOUNDED_TEXTFLOW)
                    except:
                        pass
                    try:
                        scribus.setTextFlowMode(img_frame, TEXT_FLOW_OBJECTBOUNDINGBOX)
                    except:
                        pass
            else:
                # Multiple images - use roadsigns grid layout for roadsigns
                place_roadsigns_grid(
                    image_list, base_path,
                    start_y=y_offset,
                    max_height=frame_h
                )
                
            placed_images = True

        # Calculate the height of content for this frame
        used_h = measure_text_height(chunk, frame_w, in_template)
        
        # Add one line of space at the bottom to prevent text from being cut off
        line_height = default_font_size * 1.2  # Approximate line height
        used_h += line_height
        
        # If we have images, make sure frame height is exactly equal to roadsign height
        if placed_images:
            rows_needed = (len(image_list) + 1) // 2  # Integer division rounded up
            total_roadsign_height = (roadsign_height * rows_needed) + (BLOCK_SPACING * (rows_needed - 1)) if rows_needed > 0 else 0
            total_roadsign_height += 4  # Add padding
            
            if used_h < total_roadsign_height:
                used_h = total_roadsign_height
            
        scribus.sizeObject(frame_w, used_h, frame)

        y_offset += used_h + BLOCK_SPACING
        remaining = remaining[chunk_len:]
        base_offset += chunk_len
        
        if remaining:
            # Store if we placed images on the first page
            had_images = placed_images
            # Create new page for continued content
            new_page()
            
            # If we had images on first page and have more content,
            # reduce the top margin for the continued text to minimize gap
            if had_images and not placed_images:
                # Only adjust y_offset if we didn't place images yet
                # This reduces the excessive spacing when content flows to new page
                y_offset = MARGINS[1] + 2  # Minimal spacing at top of continued page

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
    quiz_header_height = 34
    card_spacing = 1  # Minimal spacing between cards
    answer_box_width = 25
    answer_box_height = 16
    quiz_width = PAGE_WIDTH - MARGINS[0] - MARGINS[2]
    horizontal_spacing = QUIZ_HORIZONTAL_SPACING
    answer_box_gap = 4
    answer_box_fill = "White"
    answer_box_border = "Black"
    answer_box_radius = 3
    answer_box_font_size = 14
    answer_box_font_bold = False
    quiz_bar_height = 26
    quiz_bar_y = y_offset + 3
    # Only place the heading if it hasn't been placed on this page
    if not quiz_heading_placed_on_page:
        quiz_label_frame = scribus.createText(MARGINS[0] + 6, quiz_bar_y, 100, quiz_bar_height)
        scribus.setText("QUIZ", quiz_label_frame)
        try:
            scribus.setFont(QUIZ_ACTUAL_FONT, quiz_label_frame)
        except:
            scribus.setFont(DEFAULT_FONT, quiz_label_frame)
        scribus.setFontSize(QUIZ_HEADING_FONT_SIZE, quiz_label_frame)
        try:
            scribus.setLineColor("None", quiz_label_frame)
        except:
            pass
        try:
            scribus.setFontSize(QUIZ_HEADING_FONT_SIZE + 2, quiz_label_frame)
        except:
            pass
        quiz_heading_placed_on_page = True
    y_offset = quiz_bar_y + quiz_bar_height + 2
    # --- IMAGE HANDLING ---
    img_w = 0
    img_h = 0
    img_frame = None
    img_margin = 6
    image_height = 32  # Fixed height for quiz images
    image_gap = 7      # Gap between image and question text
    img_path = None
    if group_image:
        img_path = os.path.join(base_path, group_image) if base_path else group_image
        orig_w, orig_h = get_image_size(img_path)
        scale = float(image_height) / float(orig_h) if orig_h else 1.0
        img_w = orig_w * scale
        img_h = image_height
    # Calculate total height for all questions
    question_heights = []
    question_frames = []
    answer_boxes_total_width = (2 * answer_box_width) + answer_box_gap
    question_width = quiz_width - (img_margin + (img_w + image_gap if img_path else 0)) - answer_boxes_total_width - 10
    for qa in filtered_arr:
        question = qa.get('que', '')
        formatted_question = handle_superscripts(question)
        temp_frame = scribus.createText(0, 0, question_width, 40)
        scribus.setText(formatted_question, temp_frame)
        try:
            scribus.setFont(QUIZ_ACTUAL_FONT, temp_frame)
        except:
            scribus.setFont(DEFAULT_FONT, temp_frame)
        while scribus.textOverflows(temp_frame):
            w, h = scribus.getSize(temp_frame)
            scribus.sizeObject(w, h + 10, temp_frame)
        q_height = scribus.getSize(temp_frame)[1]
        scribus.deleteObject(temp_frame)
        question_heights.append(q_height)
    # Add minimal margin before quiz cards
    card_top_margin = 1  # Minimal space above quiz cards
    padding_top = 1
    padding_bottom = 1
    
    # Adjust y_offset to add space before the quiz card
    y_offset += card_top_margin
    
    # Card height calculation
    total_questions_height = sum([h + padding_top + padding_bottom for h in question_heights]) + (card_spacing * (len(filtered_arr)-1))
    card_height = max(answer_box_height + padding_top + padding_bottom, total_questions_height, img_h + padding_top + padding_bottom)
    
    # Create quiz card with the white background
    card = scribus.createRect(MARGINS[0], y_offset, quiz_width, card_height)
    scribus.setFillColor("White", card)
    scribus.setLineColor("Black", card)
    scribus.setLineWidth(0.5, card)
    # Place image if present
    if img_path:
        img_frame = scribus.createImage(MARGINS[0] + img_margin, y_offset + padding_top + (card_height - img_h - padding_top - padding_bottom) / 2, img_w, img_h)
        scribus.loadImage(img_path, img_frame)
        scribus.setScaleImageToFrame(True, True, img_frame)
        scribus.setLineColor("None", img_frame)
    # Place all questions and answer boxes stacked to the right of the image
    current_y = y_offset + padding_top
    for idx, qa in enumerate(filtered_arr):
        question = qa.get('que', '')
        formatted_question = handle_superscripts(question)
        q_height = question_heights[idx]
        question_x = MARGINS[0] + img_margin + (img_w + image_gap if img_path else 0)
        q_frame = scribus.createText(question_x, current_y, question_width, q_height)
        scribus.setText(formatted_question, q_frame)
        try:
            scribus.setFont(QUIZ_ACTUAL_FONT, q_frame)
        except:
            scribus.setFont(DEFAULT_FONT, q_frame)
        scribus.setFontSize(QUIZ_QUESTION_FONT_SIZE, q_frame)
        try:
            scribus.setLineColor("None", q_frame)
        except:
            pass
        # Place answer boxes (V/F)
        ans_x_f = MARGINS[0] + quiz_width - answer_box_width - QUIZ_QUESTION_PADDING
        ans_x_v = ans_x_f - answer_box_width - answer_box_gap
        ans_y = current_y + (q_height - answer_box_height) / 2
        ans_box_v = scribus.createRect(ans_x_v, ans_y, answer_box_width, answer_box_height)
        try:
            scribus.setFillColor(answer_box_fill, ans_box_v)
        except:
            scribus.setFillColor("White", ans_box_v)
        try:
            scribus.setLineColor(answer_box_border, ans_box_v)
        except:
            scribus.setLineColor("Black", ans_box_v)
        scribus.setLineWidth(0.7, ans_box_v)
        try:
            scribus.setCornerRadius(ans_box_v, answer_box_radius)
        except:
            pass
        ans_txt_v = scribus.createText(ans_x_v, ans_y, answer_box_width, answer_box_height)
        scribus.setText(QUIZ_TRUE_TEXT, ans_txt_v)
        try:
            scribus.setFont(QUIZ_ACTUAL_FONT, ans_txt_v)
        except:
            scribus.setFont(DEFAULT_FONT, ans_txt_v)
        scribus.setFontSize(answer_box_font_size, ans_txt_v)
        if answer_box_font_bold:
            try:
                scribus.setFontSize(answer_box_font_size + 2, ans_txt_v)
            except:
                pass
        scribus.setTextAlignment(scribus.ALIGN_CENTERED, ans_txt_v)
        try:
            scribus.setLineColor("None", ans_txt_v)
        except:
            pass
        ans_box_f = scribus.createRect(ans_x_f, ans_y, answer_box_width, answer_box_height)
        try:
            scribus.setFillColor(answer_box_fill, ans_box_f)
        except:
            scribus.setFillColor("White", ans_box_f)
        try:
            scribus.setLineColor(answer_box_border, ans_box_f)
        except:
            scribus.setLineColor("Black", ans_box_f)
        scribus.setLineWidth(0.7, ans_box_f)
        try:
            scribus.setCornerRadius(ans_box_f, answer_box_radius)
        except:
            pass
        ans_txt_f = scribus.createText(ans_x_f, ans_y, answer_box_width, answer_box_height)
        scribus.setText(QUIZ_FALSE_TEXT, ans_txt_f)
        try:
            scribus.setFont(QUIZ_ACTUAL_FONT, ans_txt_f)
        except:
            scribus.setFont(DEFAULT_FONT, ans_txt_f)
        scribus.setFontSize(answer_box_font_size, ans_txt_f)
        if answer_box_font_bold:
            try:
                scribus.setFontSize(answer_box_font_size + 2, ans_txt_f)
            except:
                pass
        scribus.setTextAlignment(scribus.ALIGN_CENTERED, ans_txt_f)
        try:
            scribus.setLineColor("None", ans_txt_f)
        except:
            pass
        current_y += q_height + card_spacing
    # Add minimal spacing after the quiz card
    quiz_bottom_margin = 1  # Minimal space after quiz cards
    
    # Update y_offset to include card height and bottom margin
    y_offset += card_height + quiz_bottom_margin

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
            scribus.setFont(DEFAULT_FONT, temp_frame)
        while scribus.textOverflows(temp_frame):
            w, h = scribus.getSize(temp_frame)
            scribus.sizeObject(w, h + 10, temp_frame)
        q_height = scribus.getSize(temp_frame)[1]
        scribus.deleteObject(temp_frame)
        question_heights.append(q_height)
    # Add minimal margin before quiz cards
    card_top_margin = 1  # Minimal space above quiz cards
    padding_top = 1
    padding_bottom = 1
    
    # Instead of modifying y_offset directly, include its value in the calculation
    
    # Card height calculation
    total_questions_height = sum([h + padding_top + padding_bottom for h in question_heights]) + (card_spacing * (len(group)-1))
    card_height = max(answer_box_height + padding_top + padding_bottom, total_questions_height, img_h + padding_top + padding_bottom)
    
    # Add space for the heading (QUIZ bar)
    total_height = (quiz_bar_height + 3) + card_height + card_top_margin + 2
    return total_height

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
            available = PAGE_HEIGHT - y_offset - MARGINS[3]
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
    # Check if we have enough space for at least the template header
    # If not, start a new page
    if y_offset + 10 > PAGE_HEIGHT - MARGINS[3]:  # Reduced from TEMPLATE_PADDING + 20
        new_page()
    else:
        # Add padding before the template only if we're not at the top of a page
        if y_offset > MARGINS[1]:
            y_offset += TEMPLATE_PADDING / 2  # Use half the normal padding
    
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
    
    # Place text and road signs
    if cleaned_txt or rs:
        place_wrapped_text_and_images(cleaned_txt, rs, base_path, 10, True, template_alignment)
    
    # Place videos
    for v in tmpl.get("videos", []):
        # Make video text always visible
        video_text = f"[video: {v}]"
        place_text_block_flow(video_text, 9, False, True)
    
    # Get and normalize images
    imgs = tmpl.get("images", [])
    imgs = imgs if isinstance(imgs, list) else ([imgs] if imgs else [])
    
    # Check if we have images to place
    if imgs:
        # Get available space
        available_space = space_left_on_page()
            
        # Normal height for images - this will be consistent across pages
        standard_image_height = 150
        
        # Special case: For single image templates, always use the standard size
        if len(imgs) == 1:
            # For single images, be more flexible with available space
            # Allow fitting if at least 60% of the standard height is available
            min_height_for_single = standard_image_height * 0.6  # 60% of standard height
            
            if available_space >= min_height_for_single:
                # Single image - place at left margin with adjusted height to fit available space
                img_path = os.path.join(base_path, imgs[0])
                if Image:
                    try:
                        with Image.open(img_path) as im:
                            orig_w, orig_h = im.size
                    except:
                        orig_w, orig_h = (300, 200)
                else:
                    orig_w, orig_h = (300, 200)
                
                # Use either standard height or adjusted to fit available space
                actual_height = min(standard_image_height, available_space - 5) # Leave minimal margin
                
                # Scale image maintaining aspect ratio with the chosen height
                scale = float(actual_height) / float(orig_h) if orig_h else 1.0
                new_w = orig_w * scale
                
                # Position at left margin
                x = MARGINS[0]
                
                img_frame = scribus.createImage(x, y_offset, new_w, actual_height)
                scribus.loadImage(img_path, img_frame)
                scribus.setScaleImageToFrame(True, True, img_frame)
                scribus.setLineColor("None", img_frame)
                y_offset += actual_height + BLOCK_SPACING
            else:
                # Not enough space, move to next page
                new_page()
                # Single image - place at left margin with standard height
                img_path = os.path.join(base_path, imgs[0])
                if Image:
                    try:
                        with Image.open(img_path) as im:
                            orig_w, orig_h = im.size
                    except:
                        orig_w, orig_h = (300, 200)
                else:
                    orig_w, orig_h = (300, 200)
                
                # Scale image maintaining aspect ratio with standard height
                scale = float(standard_image_height) / float(orig_h) if orig_h else 1.0
                new_w = orig_w * scale
                
                # Position at left margin
                x = MARGINS[0]
                
                img_frame = scribus.createImage(x, y_offset, new_w, standard_image_height)
                scribus.loadImage(img_path, img_frame)
                scribus.setScaleImageToFrame(True, True, img_frame)
                scribus.setLineColor("None", img_frame)
                y_offset += standard_image_height + BLOCK_SPACING
        else:
            # For multiple images, we need to decide if we can fit at least one row
            # Calculate how many images we can fit in the first row
            first_row_count = min(3, len(imgs))
            
            # If we have enough space for at least one row at a reasonable size
            # We'll use a minimum height of 80 points to ensure image quality
            min_acceptable_height = 80
            
            if available_space >= min_acceptable_height + 10:
                # We can fit at least one row
                # Place first row of images with adjusted height to fit available space
                # but not smaller than minimum acceptable height
                adjusted_height = max(min_acceptable_height, min(standard_image_height, available_space - 10))
                
                # Place just the first row
                y_offset = place_images_grid(imgs[:first_row_count], base_path, y_offset, adjusted_height)
                
                # If more images, continue on next page
                if len(imgs) > first_row_count:
                    new_page()
                    # Critical: use the SAME adjusted_height for consistency
                    # This ensures images on the next page match the size of those on the previous page
                    y_offset = place_images_grid(imgs[first_row_count:], base_path, y_offset, adjusted_height)
            else:
                # Not enough space for even one row at acceptable size, move to next page
                new_page()
                # Place all images with standard height
                y_offset = place_images_grid(imgs, base_path, y_offset, standard_image_height)
    
    # Place quiz section using global constants - only if quizzes are enabled
    if "quiz" in tmpl and PRINT_QUIZZES:
        # Group quizzes by their first quiz_images entry
        quiz_entries = tmpl.get("quiz", [])
        groups = {}
        for qa in quiz_entries:
            img_key = None
            quiz_images = qa.get('quiz_images', [])
            if isinstance(quiz_images, list) and quiz_images:
                img_key = quiz_images[0]
            elif isinstance(quiz_images, str) and quiz_images:
                img_key = quiz_images
            # Use None as key for quizzes with no image
            group_key = img_key if img_key else None
            if group_key not in groups:
                groups[group_key] = []
            groups[group_key].append(qa)
        # For each group, call the paginated placement function
        for img_key, group in groups.items():
            group_image = img_key
            place_quiz_group_paginated(group, group_image, base_path)
    
    # # Add padding after the template (only if we're not going to immediately add another template)
    # if y_offset < PAGE_HEIGHT - MARGINS[3] - TEMPLATE_PADDING:
    #     y_offset += TEMPLATE_PADDING
    
    global_template_count += 1

# ────────────────────────────────────────────────────────────────────────────────
# ─────────── VERTICAL TOPIC BANNER FUNCTIONS ────────────
# ────────────────────────────────────────────────────────────────────────────────
def add_vertical_topic_banner(topic_name):
    """
    Creates a vertical rectangle with topic text on the side of the page.
    
    Parameters:
    - topic_name: The name of the topic to display vertically
    """
    global current_topic_text, current_topic_color
    
    # Store the current topic for use when creating new pages
    current_topic_text = topic_name
    
    # Assign a consistent color to this topic based on topic name
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
    banner_x = (MARGINS[0] - TOPIC_BANNER_WIDTH - 2) if is_odd_page else (PAGE_WIDTH - MARGINS[2] + 2)
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

def handle_superscripts(text):
    """
    Superscript handling function for HTML tagged content and special patterns.
    Handles:
    - <span class="S-T...">digits</span>
    - <sup>digits</sup>
    - Complex HTML span patterns with style attributes
    """
    if not text or not isinstance(text, str):
        return text
    
    # Handle special HTML span pattern for digit superscripts with any unit
    text = re.sub(
        r'([a-zA-Z]+)</span><span\s+class=["\']S-T\d+["\']\s+style=["\'][^"\']*vertical-align[^"\']*["\']>\s*(\d+)\s*</span>',
        lambda m: m.group(1) + ''.join(_SUP_MAP.get(ch, ch) for ch in m.group(2)),
        text
    )
    
    # Replace any <span class="S-T...">digits</span>
    text = re.sub(
        r'<span\s+class=["\']S-T[^"\']*["\']\s*[^>]*>(\d+)</span>',
        lambda m: ''.join(_SUP_MAP.get(ch, ch) for ch in m.group(1)),
        text
    )
    
    # Replace <sup>digits</sup>
    text = re.sub(
        r'<sup>(\d+)</sup>',
        lambda m: ''.join(_SUP_MAP.get(ch, ch) for ch in m.group(1)),
        text
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
    # Check page overflow
    frame_w = PAGE_WIDTH - MARGINS[0] - MARGINS[2]
    # Add extra vertical padding to text height
    text_h = measure_text_height(text, frame_w) + (actual_padding * 2)
    if y_offset + text_h > PAGE_HEIGHT - MARGINS[3]:
        new_page()
    # Only create a background rectangle if a real color is specified
    if bg_color and bg_color.lower() != "none":
        bg_rect = scribus.createRect(MARGINS[0], y_offset, frame_w, text_h)
        scribus.setFillColor(bg_color, bg_rect)
        try:
            scribus.setLineColor("None", bg_rect)
        except:
            pass
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
            scribus.setLineSpacing(scribus.LINESP_PERCENT, 85, text_frame)
        except:
            pass
    # Only auto-resize text frame if there is no background rectangle
    if not bg_rect:
        try:
            while scribus.textOverflows(text_frame):
                current_h = scribus.getSize(text_frame)[1]
                scribus.sizeObject(frame_w - (actual_padding * 2), current_h + 5, text_frame)
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
        except:
            pass
    spacing_after = BLOCK_SPACING * 0.25 if is_template_header else BLOCK_SPACING
    y_offset += text_h + spacing_after
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
                    elif "px" in size_str:
                        # Approximate px to pt (0.75 factor)
                        size = float(size_str.replace("px", "").strip()) * 0.75
                    elif "em" in size_str:
                        size = float(size_str.replace("em", "").strip()) * default_size
                    elif "%" in size_str:
                        pct = float(re.sub(r'[^\d.]','', size_str)) / 100.0
                        size = default_size * pct
                    else:
                        size = float(re.sub(r'[^\d.]','', size_str))
                        
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
    global y_offset, global_template_count, limit_reached, current_topic_text, current_topic_color, PRINT_QUIZZES, QUIZ_FILTER_MODE
    
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
    
    # Initialize state
    y_offset = MARGINS[1]
    global_template_count = 0
    current_topic_text = None
    current_topic_color = None
    
    # Process content hierarchy
    for area in data["areas"]:
        if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
            break
            
        create_styled_header(f"{area.get('name','Unnamed')}", 20, True, "None", "Black", 5)
        place_text_block_flow(area.get("desc",""), 10)
        
        for chap in area.get("chapters", []):
            if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
                break
                
            create_styled_header(f"{chap.get('name','Unnamed')}", 20, True, "None", "Black", 5)
            
            for topic in chap.get("topics", []):
                if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
                    break
                    
                # Clear previous topic banner and set new one
                current_topic_text = topic.get('name', 'Unnamed')
                create_topic_header(current_topic_text)
                add_vertical_topic_banner(current_topic_text)
                
                # Display topic description
                place_text_block_flow(topic.get("desc",""), 10)
                
                for mod in topic.get("modules", []):
                    if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
                        break
                        
                    create_module_header(mod.get('name','Unnamed'))
                    
                    for tmpl in mod.get("templates", []):
                        if global_template_count >= GLOBAL_TEMPLATE_LIMIT:
                            break
                        process_template(tmpl, base_pics)
    
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