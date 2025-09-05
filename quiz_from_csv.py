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
                # TEMPORARY: Process only chapter 1a for testing
                if row['Chapter'] != '1a':
                    continue
                    
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
        
        # Process all questions in this chapter with improved table splitting
        for i, (question_id, question_data) in enumerate(sorted_questions):
            # Use the new table splitting function that handles row-by-row overflow
            result = create_quiz_question_with_table_splitting(question_data, y_offset, current_page)
            y_offset = result['new_y']
            current_page = result['current_page']
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

def create_quiz_question_with_table_splitting(question_data, start_y, current_page):
    """
    Create a question block with intelligent table splitting.
    If the answers table has too many rows to fit on a page, it splits at row boundaries.
    Returns dict with new_y position and current_page number.
    """
    global PAGE_WIDTH, PAGE_HEIGHT, MARGINS
    
    quiz_width = PAGE_WIDTH - MARGINS[1] - MARGINS[3]
    current_y = start_y
    
    # Calculate header dimensions
    question_text = f"Capitolo {question_data['chapter']} - {question_data['question_text']}"
    header_text_width = quiz_width - 35  # Quiz width minus ID badge
    header_height = check_text_overflow(question_text, header_text_width, 7, 26, is_header=True)
    
    # Check if we need to start with a new page for the header
    available_space = PAGE_HEIGHT - current_y - MARGINS[2]
    if header_height + 20 > available_space:  # Need at least header + one row
        scribus.newPage(-1)
        current_page += 1
        current_y = MARGINS[0]
        scribus.gotoPage(current_page)
        add_page_number(current_page)
    
    # Draw the question header
    draw_question_header(question_data, current_y, header_height, quiz_width)
    current_y += header_height + 3  # Small gap after header
    
    # Sort answers by number
    sorted_answers = sorted(question_data['answers'], key=lambda x: int(x['number']))
    
    # Process answers with row-by-row overflow checking
    answer_index = 0
    total_rows = len(sorted_answers)
    rows_on_current_page = 0
    first_page_of_question = True
    
    while answer_index < total_rows:
        answer = sorted_answers[answer_index]
        
        # Calculate this row's height
        text_width = quiz_width - 58  # Quiz width minus boxes
        row_height = check_text_overflow(answer['text'], text_width, 6, 14, is_header=False)
        
        # Check available space for this row
        available_space = PAGE_HEIGHT - current_y - MARGINS[2]
        
        # If this row won't fit, start a new page
        if row_height + 10 > available_space:  # 10pt safety buffer
            # Go to new page
            scribus.newPage(-1)
            current_page += 1
            current_y = MARGINS[0]
            scribus.gotoPage(current_page)
            add_page_number(current_page)
            
            # Add a continuation header for split tables
            if not first_page_of_question and rows_on_current_page > 0:
                cont_height = draw_continuation_header(question_data, current_y)
                current_y += cont_height + 3
            
            rows_on_current_page = 0
            first_page_of_question = False
        
        # Draw this answer row
        draw_answer_row(answer, answer_index, current_y, row_height, quiz_width)
        current_y += row_height
        rows_on_current_page += 1
        answer_index += 1
    
    # Add spacing after the complete question
    current_y += 3
    
    return {'new_y': current_y, 'current_page': current_page}

def draw_question_header(question_data, y_pos, header_height, quiz_width):
    """Draw the question header with blue background"""
    header_color = "Cyan"
    
    # Create colorful header background
    header_bg = scribus.createRect(MARGINS[1], y_pos, quiz_width, header_height)
    scribus.setFillColor(header_color, header_bg)
    scribus.setLineColor(header_color, header_bg)
    
    # Question text in header with chapter - more width available due to smaller ID box
    header = scribus.createText(MARGINS[1] + 3, y_pos + 3, quiz_width - 35, header_height - 6)  # More space and better vertical positioning
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
    
    # Question ID badge (right side)
    id_width = 30
    id_bg = scribus.createRect(MARGINS[1] + quiz_width - id_width, y_pos, id_width, header_height)
    scribus.setFillColor("Yellow", id_bg)
    scribus.setLineColor("Black", id_bg)
    
    id_box = scribus.createText(MARGINS[1] + quiz_width - id_width + 1, y_pos + 6, 
                                id_width - 2, header_height - 12)
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
        scribus.setFontSize(8, id_box)
    except:
        pass
    scribus.setTextColor("Black", id_box)
    scribus.setTextAlignment(1, id_box)

def draw_continuation_header(question_data, y_pos):
    """Draw a smaller header for continued questions on new pages"""
    header_height = 16
    quiz_width = PAGE_WIDTH - MARGINS[1] - MARGINS[3]
    
    # Create a subtle continuation header
    cont_bg = scribus.createRect(MARGINS[1], y_pos, quiz_width, header_height)
    try:
        scribus.defineColor("LightCyan", 200, 240, 255)
        scribus.setFillColor("LightCyan", cont_bg)
    except:
        scribus.setFillColor("LightGray", cont_bg)
    scribus.setLineColor("Cyan", cont_bg)
    
    # Continuation text
    cont_text = scribus.createText(MARGINS[1] + 3, y_pos + 2, quiz_width - 6, header_height - 4)
    scribus.setText(f"Question {question_data['question_id']} (continued)", cont_text)
    try:
        scribus.setFont("Liberation Sans", cont_text)
        scribus.setFontSize(6, cont_text)  # Keep 6pt for continuation header
    except:
        pass
    scribus.setTextColor("Black", cont_text)
    scribus.setTextAlignment(0, cont_text)
    
    return header_height

def draw_answer_row(answer, row_index, y_pos, row_height, quiz_width):
    """Draw a single answer row"""
    # Answer number box
    num_box_bg = scribus.createRect(MARGINS[1] + 2, y_pos, 12, row_height - 1)
    try:
        scribus.defineColor("NumBoxBlue", 210, 235, 255)
        scribus.setFillColor("NumBoxBlue", num_box_bg)
    except:
        scribus.setFillColor("LightGray", num_box_bg)
    scribus.setLineColor("Cyan", num_box_bg)
    scribus.setLineWidth(0.5, num_box_bg)
    
    # Answer number text - properly centered in 12pt number box
    num_box_height = row_height - 2  # Use almost full row height with 1pt padding
    text_y_offset = 1  # Minimal top padding
    num_box = scribus.createText(MARGINS[1] + 2, y_pos + text_y_offset, 12, num_box_height)  # Full height number box
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
    
    # Answer text box
    text_width = quiz_width - 58
    text_box_bg = scribus.createRect(MARGINS[1] + 16, y_pos, text_width, row_height - 1)
    
    # Alternate row colors
    if row_index % 2 == 0:
        scribus.setFillColor("White", text_box_bg)
    else:
        try:
            scribus.defineColor("VeryLightCyan", 245, 252, 255)
            scribus.setFillColor("VeryLightCyan", text_box_bg)
        except:
            scribus.setFillColor("LightGray", text_box_bg)
    
    scribus.setLineColor("Cyan", text_box_bg)
    scribus.setLineWidth(0.5, text_box_bg)
    
    # Answer text - properly centered vertically and horizontally
    text_box_height = row_height - 2  # Use almost full row height with 1pt padding top/bottom
    text_y_offset_answer = 1  # Minimal top padding
    text_box = scribus.createText(MARGINS[1] + 17, y_pos + text_y_offset_answer, text_width - 2, text_box_height)
    
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
    v_box_bg = scribus.createRect(MARGINS[1] + quiz_width - 40, y_pos, 18, row_height - 1)  # Almost full height
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
    v_box = scribus.createText(MARGINS[1] + quiz_width - 40, y_pos + checkbox_y_offset, 18, checkbox_box_height)  # Full 18pt width
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
        scribus.setTextColor("Cyan", v_box)
    else:
        scribus.setTextColor("Black", v_box)
    
    # 4. F checkbox box - ultra compact for 14pt rows
    f_box_bg = scribus.createRect(MARGINS[1] + quiz_width - 20, y_pos, 18, row_height - 1)  # Almost full height
    try:
        scribus.defineColor("CheckBoxColor2", 255, 240, 240)  # Light red tint
        scribus.setFillColor("CheckBoxColor2", f_box_bg)
    except:
        scribus.setFillColor("White", f_box_bg)
    scribus.setLineColor("Cyan", f_box_bg)
    scribus.setLineWidth(0.5, f_box_bg)
    
    # F checkbox text - properly centered in 18pt checkbox area
    f_box = scribus.createText(MARGINS[1] + quiz_width - 20, y_pos + checkbox_y_offset, 18, checkbox_box_height)  # Full 18pt width
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
        scribus.setTextColor("Cyan", f_box)
    else:
        scribus.setTextColor("Black", f_box)


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