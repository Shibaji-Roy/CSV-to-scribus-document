#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Scribus Quiz Template Generator from CSV
Creates a multi-page quiz document from CSV data with same layout as template
Each question starts on a new page with blue header and V/F checkboxes
"""

import scribus
import sys
import csv
import os

# Global variables for page dimensions and state (like final_pdf.py)
y_offset = 0
PAGE_WIDTH = 0
PAGE_HEIGHT = 0
MARGINS = (0, 0, 0, 0)

def read_csv_data(filepath):
    """Read and parse the CSV file - process all questions"""
    questions = {}
    
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            
            for row in reader:
                # Process ALL chapters - no filtering
                question_id = row['QuestionID']
                
                if question_id not in questions:
                    questions[question_id] = {
                        'chapter': row['Chapter'],
                        'question_id': question_id,
                        'question_text': row['QuestionText'],
                        'answers': []
                    }
                
                questions[question_id]['answers'].append({
                    'number': row['AnswerNumber'],
                    'text': row['AnswerText'],
                    'correct': row['CorrectFlag']
                })
    
    except Exception as e:
        scribus.messageBox('Error', 'Failed to read CSV file: ' + str(e))
        return None
    
    return questions

def create_quiz_document(csv_filepath):
    """Main function to create the quiz document from CSV - ONE document with all chapters"""
    global y_offset  # Global state like final_pdf.py uses
    global PAGE_WIDTH, PAGE_HEIGHT, MARGINS
    
    # Read CSV data
    all_questions = read_csv_data(csv_filepath)
    if not all_questions:
        return
    
    # Document settings - A5 in points 
    PAGE_WIDTH = 420.94  # A5 width in points  
    PAGE_HEIGHT = 595.28  # A5 height in points
    MARGINS = (30, 30, 30, 30)  # Reduced margins for more content area
    
    # Create ONE document for all chapters
    scribus.newDocument((PAGE_WIDTH, PAGE_HEIGHT), MARGINS,
                        scribus.PORTRAIT, 1, scribus.UNIT_POINTS,
                        scribus.PAGE_1, 0, 1)
    
    # Initialize y_offset
    y_offset = MARGINS[0]  # Use top margin
    
    # Define colors
    define_colors()
    
    # Add page number to first page
    add_page_number(1)
    
    # Get all unique chapters and sort them with natural ordering
    chapters = set()
    for question_data in all_questions.values():
        chapters.add(question_data['chapter'])
    
    # Natural sorting function for chapters like 1a, 1b, 2a, 2b, etc.
    def natural_sort_key(chapter):
        import re
        # Split chapter into numeric and alphabetic parts
        parts = re.match(r'(\d+)([a-z]*)', chapter.lower())
        if parts:
            num_part = int(parts.group(1))
            alpha_part = parts.group(2) if parts.group(2) else ''
            return (num_part, alpha_part)
        return (999, chapter)  # Put non-matching chapters at the end
    
    chapters = sorted(list(chapters), key=natural_sort_key)
    
    # Show chapters being processed
    chapter_list = ', '.join(chapters)
    scribus.messageBox('Processing Chapters', f'Creating single document with chapters: {chapter_list}')
    
    # Track current page and total questions
    current_page = 1
    total_questions_processed = 0
    
    # Process each chapter in the SAME document
    for chapter in chapters:
        # Filter questions for this chapter
        chapter_questions = {qid: qdata for qid, qdata in all_questions.items() 
                           if qdata['chapter'] == chapter}
        
        if not chapter_questions:
            continue
        
        # Sort questions by ID
        sorted_questions = sorted(chapter_questions.items(), key=lambda x: x[0])
        
        # Process all questions in this chapter
        for i, (question_id, question_data) in enumerate(sorted_questions):
            # Calculate space needed for this question with smart overflow detection
            
            # Check if header needs more height
            question_text = f"Capitolo {question_data['chapter']} - {question_data['question_text']}"
            header_text_width = PAGE_WIDTH - MARGINS[1] - MARGINS[3] - 35  # Quiz width minus ID badge
            header_height = check_text_overflow(question_text, header_text_width, 7, 26, is_header=True)
            
            # Check each answer for overflow and calculate individual heights
            total_answer_height = 0
            answer_heights = []
            for answer in question_data['answers']:
                text_width = PAGE_WIDTH - MARGINS[1] - MARGINS[3] - 58  # Quiz width minus boxes
                answer_height = check_text_overflow(answer['text'], text_width, 6, 14, is_header=False)
                answer_heights.append(answer_height)
                total_answer_height += answer_height
            
            spacing_after_header = 3  # Small gap after header to separate from answers  
            spacing_after_question = 2  # 2px gap between questions as requested
            
            # Total question height with dynamic sizes only where needed
            total_question_height = header_height + spacing_after_header + total_answer_height + spacing_after_question
            
            # Store heights for this question to pass to create function
            question_data['_header_height'] = header_height
            question_data['_answer_heights'] = answer_heights
            
            # Check available space
            available_space = PAGE_HEIGHT - y_offset - MARGINS[2]
            
            # If question would exceed bottom margin, move to next page
            safety_buffer = 10  # Smaller buffer to use more page space
            if total_question_height > (available_space - safety_buffer):
                # Go to new page
                scribus.newPage(-1)
                current_page += 1
                # Reset y position
                y_offset = MARGINS[0]
                # Make sure we're on the new page
                scribus.gotoPage(current_page)
                # Add page number to new page
                add_page_number(current_page)
            
            # Create the question block
            new_y = create_quiz_question_block(question_data, y_offset)
            
            # Update y_offset for next question
            y_offset = new_y
            total_questions_processed += 1
    
    # Add footer to the last page
    add_page_footer(current_page, total_questions_processed, "All Chapters")
    
    # Final message
    scribus.messageBox('Complete', f'Single document created with {total_questions_processed} questions from {len(chapters)} chapters!')

def check_text_overflow(text, width, font_size, default_height, is_header=False):
    """Check if text will overflow and calculate required height if needed"""
    # More conservative estimation to ensure no overflow
    if is_header:
        # Headers: more conservative estimate
        chars_per_line = width / (font_size * 0.45)
    else:
        # Regular text: much more conservative for answers
        chars_per_line = width / (font_size * 0.42)  # More space per character
    
    # Check if text truly needs multiple lines
    # Lower threshold to catch more potential overflows
    if len(text) <= chars_per_line * 0.85:  # 85% threshold - more conservative
        return default_height  # Single line - use compact height
    
    # Calculate actual lines needed
    lines_needed = int(len(text) / chars_per_line) + 1
    
    # Special handling for borderline cases (text near one line)
    if len(text) > chars_per_line * 0.85 and len(text) <= chars_per_line * 1.2:
        # Text is close to or slightly over one line - give more space
        return default_height * 1.5  # 50% more height for borderline cases
    
    # Calculate required height for true multi-line content
    line_height = font_size * 1.35  # More generous line spacing
    calculated_height = lines_needed * line_height + 6  # More padding
    
    # Add extra buffer for safety
    calculated_height = calculated_height * 1.1  # 10% extra buffer
    
    # Return calculated height
    return max(default_height, calculated_height)

def define_colors():
    """Define colors using Scribus default colors like final_pdf.py"""
    # Use Scribus built-in colors - these always work
    pass  # No need to define colors, use built-ins
    
    try:
        # Natural light gray for ID badges  
        scribus.defineColor("BadgeGray", 192, 192, 192)  # Silver gray
    except:
        pass
    
    try:
        # Very light gray for alternating rows
        scribus.defineColor("LightGray", 245, 245, 245)  # White smoke
    except:
        pass
    
    try:
        # Natural black text
        scribus.defineColor("Black", 0, 0, 0)  # Pure black
    except:
        pass
    
    try:
        # Soft gray for secondary text
        scribus.defineColor("Gray", 128, 128, 128)  # Standard gray
    except:
        pass
    
    try:
        # White for backgrounds
        scribus.defineColor("White", 255, 255, 255)  # Pure white
    except:
        pass

def create_quiz_question_block(question_data, start_y):
    """Create a single question block - matching scrib.jpg layout"""
    global PAGE_WIDTH, PAGE_HEIGHT, MARGINS
    
    # Use global dimensions set in create_quiz_document
    
    # Calculate working width like final_pdf.py
    quiz_width = PAGE_WIDTH - MARGINS[1] - MARGINS[3]  # left margin + right margin
    current_y = start_y  # Use provided Y position
    
    # Use cyan blue for headers like before
    header_color = "Cyan"
    
    # Use the pre-calculated dynamic heights (only expanded where needed)
    header_height = question_data.get('_header_height', 26)  # Use calculated height or default
    answer_heights = question_data.get('_answer_heights', [14] * len(question_data['answers']))  # Use calculated heights
    
    spacing_after_header = 3  # Small gap after header to separate from answers
    total_block_height = header_height + spacing_after_header + sum(answer_heights)
    
    # Create light blue background box for entire question block (like scrib.jpg)
    block_bg = scribus.createRect(MARGINS[1], current_y, quiz_width, total_block_height)  # Use left margin
    try:
        # Very light cyan background for the whole block
        scribus.defineColor("VeryLightCyan", 230, 250, 255)  # Light cyan tint
        scribus.setFillColor("VeryLightCyan", block_bg)
    except:
        scribus.setFillColor("LightGray", block_bg)  # Fallback to light gray
    scribus.setLineColor("Cyan", block_bg)  # Cyan border for the block
    
    # Create colorful question header section - like final_pdf.py banners
    # header_height already defined above for consistency
    header_bg = scribus.createRect(MARGINS[1], current_y, quiz_width, header_height)  # Use left margin
    scribus.setFillColor(header_color, header_bg)
    scribus.setLineColor(header_color, header_bg)
    
    # Question text in header with chapter - more width available due to smaller ID box
    header = scribus.createText(MARGINS[1] + 3, current_y + 3, quiz_width - 35, header_height - 6)  # More space and better vertical positioning
    chapter_text = f"Capitolo {question_data['chapter']} - {question_data['question_text']}"
    
    scribus.setText(chapter_text, header)
    try:
        scribus.setFont("Liberation Sans", header)
    except:
        # Use fallback fonts like final_pdf.py
        font_candidates = ["DejaVu Sans", "Bitstream Vera Sans", "Arial", "Helvetica"]
        for f in font_candidates:
            try:
                scribus.setFont(f, header)
                break
            except:
                pass
    try:
        scribus.setFontSize(7, header)  # Consistent 7pt font for headers
    except:
        pass
    scribus.setTextColor("White", header)
    scribus.setTextAlignment(0, header)
    
    # Configure header text wrapping and overflow control
    try:
        scribus.setTextDistances(1, 1, 1, 1, header)  # Minimal padding
        scribus.setLineSpacing(7, header)  # Very tight line spacing for 7pt font
        
        # Force text to fit within frame bounds - prevent header overflow  
        try:
            scribus.setTextBehaviour(header, 0)  # Force text to stay in frame
        except:
            try:
                scribus.setTextToFrameOverflow(header, False)  # Disable overflow
            except:
                pass
        
        # Enable text wrapping for headers
        try:
            scribus.setTextFlowMode(header, 0)  # Enable text flow
        except:
            pass
            
    except:
        pass
    
    # Question ID badge in header (right side) - reduced width with yellow background
    id_width = 30  # Reduced from 40 to 30
    id_bg = scribus.createRect(MARGINS[1] + quiz_width - id_width, current_y, id_width, header_height)  # Use left margin
    scribus.setFillColor("Yellow", id_bg)  # Yellow background for ID badge
    scribus.setLineColor("Black", id_bg)
    
    # Center the ID text within the reduced width box
    id_box = scribus.createText(MARGINS[1] + quiz_width - id_width + 1, current_y + 6, id_width - 2, header_height - 12)  # Centered with minimal padding
    scribus.setText(question_data['question_id'], id_box)
    try:
        scribus.setFont("Liberation Sans", id_box)
    except:
        font_candidates = ["DejaVu Sans", "Bitstream Vera Sans", "Arial", "Helvetica"]
        for f in font_candidates:
            try:
                scribus.setFont(f, id_box)
                break
            except:
                pass
    try:
        scribus.setFontSize(8, id_box)  # Smaller ID text for A5 format
    except:
        pass
    scribus.setTextColor("Black", id_box)  # Black text on yellow badge
    scribus.setTextAlignment(1, id_box)  # Center alignment for better centering
    
    current_y += header_height + spacing_after_header  # Add gap after header
    
    # Sort answers by number
    sorted_answers = sorted(question_data['answers'], key=lambda x: int(x['number']))
    
    # Create answer rows with selective dynamic heights
    for idx, answer in enumerate(sorted_answers):
        # Use the pre-calculated height for this specific answer (14pt default or expanded if needed)
        row_height = answer_heights[idx]
        
        # Create separate colored boxes for each element (like scrib.jpg)
        
        # 1. Answer number box with blue tint - dynamic height
        num_box_bg = scribus.createRect(MARGINS[1] + 2, current_y, 12, row_height - 1)  # Further reduced width
        try:
            scribus.defineColor("NumBoxBlue", 210, 235, 255)  # Light blue tint for number box
            scribus.setFillColor("NumBoxBlue", num_box_bg)
        except:
            scribus.setFillColor("LightGray", num_box_bg)
        scribus.setLineColor("Cyan", num_box_bg)
        scribus.setLineWidth(0.5, num_box_bg)
        
        # Answer number text - properly centered in 12pt number box
        num_box_height = row_height - 2  # Use almost full row height with 1pt padding
        text_y_offset = 1  # Minimal top padding
        num_box = scribus.createText(MARGINS[1] + 2, current_y + text_y_offset, 12, num_box_height)  # Full height number box
        scribus.setText(answer['number'], num_box)
        try:
            scribus.setFont("Liberation Sans", num_box)
        except:
            pass
        try:
            scribus.setFontSize(6, num_box)  # 6pt font for numbers
            scribus.setTextAlignment(1, num_box)  # Center horizontally
            # Try to set vertical alignment to middle
            try:
                scribus.setTextVerticalAlignment(1, num_box)  # 1 = middle alignment
            except:
                pass
            # Set proper text distances for centering
            scribus.setTextDistances(0, 0, 0, 0, num_box)  # No padding for perfect centering
        except:
            pass
        scribus.setTextColor("Black", num_box)
        
        # 2. Answer text box - widened to use space from reduced number box
        text_width = quiz_width - 58  # Adjust for 12pt number box + V/F boxes (wider than before)
        text_box_bg = scribus.createRect(MARGINS[1] + 16, current_y, text_width, row_height - 1)  # Start right after number box
        if idx % 2 == 0:
            # Even rows - use white
            scribus.setFillColor("White", text_box_bg)
        else:
            # Odd rows - try different approaches for visible alternation
            try:
                # Try using built-in Cyan with transparency (testing different values)
                scribus.setFillColor("Cyan", text_box_bg)
                try:
                    # Try transparency value between 0-1 (0.3 = 30% opacity)
                    scribus.setFillTransparency(0.7, text_box_bg)  # 70% transparent = light tint
                except:
                    try:
                        # Try percentage scale 0-100 (30 = 30% opacity)  
                        scribus.setFillTransparency(70, text_box_bg)  # 70% transparent = light tint
                    except:
                        # If transparency doesn't work, just use the full color
                        pass
            except:
                # If Cyan fails, try Yellow as alternative
                try:
                    scribus.setFillColor("Yellow", text_box_bg)
                    try:
                        scribus.setFillTransparency(0.8, text_box_bg)  # Very light yellow
                    except:
                        try:
                            scribus.setFillTransparency(80, text_box_bg)  # Very light yellow
                        except:
                            pass
                except:
                    # Final fallback - use LightGray if it exists
                    scribus.setFillColor("LightGray", text_box_bg)
        scribus.setLineColor("Cyan", text_box_bg)
        scribus.setLineWidth(0.5, text_box_bg)
        
        # Answer text - properly centered vertically and horizontally
        text_box_height = row_height - 2  # Use almost full row height with 1pt padding top/bottom
        text_y_offset_answer = 1  # Minimal top padding
        text_box = scribus.createText(MARGINS[1] + 17, current_y + text_y_offset_answer, text_width - 2, text_box_height)
        
        # Use full answer text - no truncation
        answer_text = answer['text']
        scribus.setText(answer_text, text_box)
        try:
            scribus.setFont("Liberation Sans", text_box)
        except:
            pass
        try:
            scribus.setFontSize(6, text_box)  # Standard size for all
        except:
            pass
        scribus.setTextColor("Black", text_box)
        # Enable proper text alignment and centering
        try:
            scribus.setTextDistances(1, 1, 1, 1, text_box)  # Small padding on all sides
            # Set line spacing based on row height
            if row_height > 14:
                scribus.setLineSpacing(7, text_box)  # More spacing for multi-line
            else:
                scribus.setLineSpacing(6, text_box)  # Tight for single line
            
            # Set horizontal alignment
            scribus.setTextAlignment(0, text_box)  # Left align
            
            # Try to set vertical alignment to middle
            try:
                scribus.setTextVerticalAlignment(1, text_box)  # 1 = middle alignment
            except:
                try:
                    # Alternative method for vertical centering
                    scribus.setTextBehaviour(text_box, 1)  # Try different behavior
                except:
                    pass
            
            # Force text to stay within bounds
            try:
                scribus.setTextBehaviour(text_box, 0)  # Force text in frame
            except:
                pass
        except:
            pass
        
        # 3. V checkbox box - ultra compact for 14pt rows
        v_box_bg = scribus.createRect(MARGINS[1] + quiz_width - 40, current_y, 18, row_height - 1)  # Almost full height
        try:
            scribus.defineColor("CheckBoxColor", 240, 255, 240)  # Light green tint
            scribus.setFillColor("CheckBoxColor", v_box_bg)
        except:
            scribus.setFillColor("White", v_box_bg)
        scribus.setLineColor("Cyan", v_box_bg)
        scribus.setLineWidth(0.5, v_box_bg)
        
        # V checkbox text - properly centered in 18pt checkbox area
        checkbox_box_height = row_height - 2  # Use almost full row height with 1pt padding
        checkbox_y_offset = 1  # Minimal top padding
        v_box = scribus.createText(MARGINS[1] + quiz_width - 40, current_y + checkbox_y_offset, 18, checkbox_box_height)  # Full 18pt width
        scribus.setText("V", v_box)
        try:
            scribus.setFontSize(8, v_box)
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
        # Color V with header color if it's the correct answer
        if answer['correct'] == 'V':
            scribus.setTextColor(header_color, v_box)
        else:
            scribus.setTextColor("Black", v_box)
        
        # 4. F checkbox box - ultra compact for 14pt rows
        f_box_bg = scribus.createRect(MARGINS[1] + quiz_width - 20, current_y, 18, row_height - 1)  # Almost full height
        try:
            scribus.defineColor("CheckBoxColor2", 255, 240, 240)  # Light red tint
            scribus.setFillColor("CheckBoxColor2", f_box_bg)
        except:
            scribus.setFillColor("White", f_box_bg)
        scribus.setLineColor("Cyan", f_box_bg)
        scribus.setLineWidth(0.5, f_box_bg)
        
        # F checkbox text - properly centered in 18pt checkbox area
        f_box = scribus.createText(MARGINS[1] + quiz_width - 20, current_y + checkbox_y_offset, 18, checkbox_box_height)  # Full 18pt width
        scribus.setText("F", f_box)
        try:
            scribus.setFontSize(8, f_box)
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
        # Color F with header color if it's the correct answer
        if answer['correct'] == 'F':
            scribus.setTextColor(header_color, f_box)
        else:
            scribus.setTextColor("Black", f_box)
        
        current_y += row_height
    
    # Add spacing after this question block - 2px gap between questions
    current_y += 3
    
    # Return the new Y position for the next question block
    return current_y

def add_page_number(page_num):
    """Add page number at bottom middle"""
    global PAGE_WIDTH, PAGE_HEIGHT, MARGINS
    
    # Create page number text box at bottom middle
    page_num_width = 30
    page_num_height = 15
    x_pos = (PAGE_WIDTH - page_num_width) / 2  # Center horizontally
    y_pos = PAGE_HEIGHT - MARGINS[2] - page_num_height  # Bottom margin minus box height
    
    try:
        page_num_box = scribus.createText(x_pos, y_pos, page_num_width, page_num_height)
        scribus.setText(str(page_num), page_num_box)
        
        # Set font and formatting
        try:
            scribus.setFont("Liberation Sans", page_num_box)
        except:
            font_candidates = ["DejaVu Sans", "Arial", "Helvetica"]
            for f in font_candidates:
                try:
                    scribus.setFont(f, page_num_box)
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

def add_page_footer(page_num, total_pages, chapter):
    """Add footer - removed all footer texts"""
    # Footer texts removed as requested
    pass

def main():
    """Main entry point - following final_pdf.py pattern"""
    # CSV file path - Windows path
    csv_filepath = r'C:\Users\rctbr\Pictures\questions_output_final.csv'
    
    # Check if file exists
    if not os.path.exists(csv_filepath):
        scribus.messageBox("Error", "CSV file not found at: " + csv_filepath)
        return
    
    # Check if running in Scribus
    try:
        scribus.statusMessage("Creating quiz template from CSV...")
        create_quiz_document(csv_filepath)
        scribus.statusMessage("Quiz template created!")
    except NameError:
        print("This script must be run from within Scribus!")

# Main execution - exactly like final_pdf.py
if __name__ == "__main__":
    if scribus.haveDoc():
        scribus.messageBox("Close the current document first.", "",
                           scribus.ICON_WARNING)
    else:
        main()