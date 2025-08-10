import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from datetime import datetime
import os

# --- Helper Functions ---

def get_year_semester_string(roman_numeral):
    """Converts a Roman numeral semester into the required string format."""
    mapping = {
        'I': 'I Year/ I semester', 'II': 'I Year/ II semester',
        'III': 'II Year/ III semester', 'IV': 'II Year/ IV semester',
        'V': 'III Year/ V semester', 'VI': 'III Year/ VI semester',
        'VII': 'IV Year/ VII semester', 'VIII': 'IV Year/ VIII semester'
    }
    return mapping.get(str(roman_numeral).strip().upper(), str(roman_numeral))

def calculate_display_weightage(midterm_marks, cgpa):
    """Calculates the two-part weightage for display purposes only."""
    try:
        # As per the template image (e.g., 5% for a score of 30)
        weightage1 = (float(midterm_marks) / 30) * 5 
        # As per the template image (e.g., 3.64 for a CGPA of 10)
        weightage2 = float(cgpa) * 0.364 
        return weightage1, weightage2
    except (ValueError, TypeError):
        return 0, 0

def calculate_midterm_percentage(midterm_marks):
    """Calculates the midterm score as a percentage for filtering."""
    try:
        return (float(midterm_marks) / 30) * 100
    except (ValueError, TypeError):
        return 0

def set_cell_properties(cell, text, bold=False, font_size=10, align='LEFT', valign='TOP'):
    """Helper to set text and alignment in a table cell."""
    cell.text = ''
    p = cell.add_paragraph()
    run = p.add_run(str(text))
    run.font.size = Pt(font_size)
    run.bold = bold
    p.alignment = getattr(WD_ALIGN_PARAGRAPH, align)
    cell.vertical_alignment = getattr(WD_ALIGN_VERTICAL, valign)

def add_signature_line(doc_or_cell):
    """Adds a formatted signature line to a document or a cell."""
    p = doc_or_cell.add_paragraph()
    p.add_run("\n\n" + "_" * 40 + "\n")
    p.add_run("Signature of the\nsubject teacher / class coordinator")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

# --- Main Report Generation Logic ---

def generate_word_report(excel_path, format_choice, learner_type, slow_threshold, fast_threshold):
    """
    Reads student data, filters for fast/slow learners, and generates a Word document.
    """
    try:
        df = pd.read_excel(excel_path)
        df.columns = df.columns.str.strip()
    except FileNotFoundError:
        print(f"Error: The file '{excel_path}' was not found.")
        return
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return

    # Calculate midterm percentage for filtering
    df['MidtermPercentage'] = df['Midterm Exam Marks (Out of 30)'].apply(calculate_midterm_percentage)

    # Filter based on the midterm percentage
    if learner_type == 'slow':
        filtered_df = df[df['MidtermPercentage'] <= slow_threshold].copy()
        report_prefix = "Slow_Learners"
    elif learner_type == 'fast':
        filtered_df = df[df['MidtermPercentage'] >= fast_threshold].copy()
        report_prefix = "Fast_Learners"
    else:
        print("Invalid learner type specified.")
        return

    if filtered_df.empty:
        print(f"\nNo students found for the '{learner_type}' category with the given threshold.")
        return

    print(f"\nFound {len(filtered_df)} {learner_type} learners. Generating report...")
    
    doc = Document()
    
    if format_choice == '1':
        for index, row in filtered_df.iterrows():
            # --- Start Formatting for Format 1 ---
            title = doc.add_heading('Format 1.   Assessment of the learning levels of the students:', level=2)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            container_table = doc.add_table(rows=5, cols=1)
            container_table.style = 'Table Grid'

            header_cell = container_table.cell(0, 0)
            header_cell.text = ''
            p1 = header_cell.add_paragraph(); p1.add_run('Manipal Institute of Technology').bold = True; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p2 = header_cell.add_paragraph(); p2.add_run('MAHE Manipal').bold = True; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p3 = header_cell.add_paragraph(); p3.add_run('Computer Science and Engineering Department').bold = True; p3.alignment = WD_ALIGN_PARAGRAPH.CENTER

            student_info_cell = container_table.cell(1, 0)
            student_info_cell.text = ''
            student_info_table = student_info_cell.add_table(rows=4, cols=2)
            set_cell_properties(student_info_table.cell(0, 0), 'Name of the Student:')
            set_cell_properties(student_info_table.cell(0, 1), str(row['Student Name']))
            set_cell_properties(student_info_table.cell(1, 0), 'Registration Number:')
            set_cell_properties(student_info_table.cell(1, 1), str(row['Register Number of the Student']))
            set_cell_properties(student_info_table.cell(2, 0), 'Course:')
            set_cell_properties(student_info_table.cell(2, 1), str(row['Subject Name']))
            set_cell_properties(student_info_table.cell(3, 0), 'Year /semester:')
            set_cell_properties(student_info_table.cell(3, 1), get_year_semester_string(row['Semester']))

            params_cell = container_table.cell(2, 0)
            params_cell.text = ''
            params_table = params_cell.add_table(rows=3, cols=3)
            params_table.style = 'Table Grid'
            set_cell_properties(params_table.cell(0, 0), 'Sr. No.', bold=True, align='CENTER')
            set_cell_properties(params_table.cell(0, 1), 'Parameter', bold=True, align='CENTER')
            set_cell_properties(params_table.cell(0, 2), 'Weightage in\nPercentage', bold=True, align='CENTER')
            
            # Calculate display weightages
            w1, w2 = calculate_display_weightage(row['Midterm Exam Marks (Out of 30)'], row['CGPA (up to previous semester)'])

            set_cell_properties(params_table.cell(1, 0), '1', align='CENTER')
            set_cell_properties(params_table.cell(1, 1), 'Scores obtained by student class test / internal examination...\nConsidered Midterm exam conducted for 30M:')
            set_cell_properties(params_table.cell(1, 2), f"{w1:.2f}     > %", align='CENTER')
            
            set_cell_properties(params_table.cell(2, 0), '2', align='CENTER')
            set_cell_properties(params_table.cell(2, 1), 'Performance of students in preceding university examination')
            set_cell_properties(params_table.cell(2, 2), f"{w2:.2f}     > %", align='CENTER')

            total_cell = container_table.cell(3, 0)
            total_cell.text = "Total Weightage" # Keep the label but no value

            footer_cell = container_table.cell(4, 0)
            footer_cell.text = ''
            footer_cell.add_paragraph(f"1. Midterm score less than {slow_threshold}% considered as a slow learner")
            footer_cell.add_paragraph(f"2. Midterm score more than {fast_threshold}% considered as an advanced learner **")
            footer_cell.add_paragraph(f"Date: {datetime.now().strftime('%d-%m-%Y')}")
            add_signature_line(footer_cell)
            # --- End Formatting for Format 1 ---
            
            if index != filtered_df.index.tolist()[-1]:
                doc.add_page_break()
        output_filename = f'{report_prefix}_Format1_Report.docx'

    elif format_choice == '2':
        for index, row in filtered_df.iterrows():
            # --- Start Formatting for Format 2 ---
            doc.add_heading('Format -2   Report of performance/ improvement for slow and advanced learners', level=2).alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            header_table = doc.add_table(rows=3, cols=1)
            p1 = header_table.cell(0, 0).paragraphs[0]; p1.add_run('Manipal Institute of Technology').bold = True; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p2 = header_table.cell(1, 0).paragraphs[0]; p2.add_run('MAHE Manipal').bold = True; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p3 = header_table.cell(2, 0).paragraphs[0]; p3.add_run('Computer Science and Engineering Department').bold = True; p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            content_table = doc.add_table(rows=8, cols=2)
            content_table.style = 'Table Grid'
            set_cell_properties(content_table.cell(0, 0), '1. Registration Number')
            set_cell_properties(content_table.cell(0, 1), row['Register Number of the Student'])
            set_cell_properties(content_table.cell(1, 0), '2. Name of the student')
            set_cell_properties(content_table.cell(1, 1), row['Student Name'])
            set_cell_properties(content_table.cell(2, 0), '3. Course')
            set_cell_properties(content_table.cell(2, 1), row['Subject Name'])
            set_cell_properties(content_table.cell(3, 0), '4. Year/Semester')
            set_cell_properties(content_table.cell(3, 1), get_year_semester_string(row['Semester']))
            set_cell_properties(content_table.cell(4, 0), '5. Midterm Percentage')
            set_cell_properties(content_table.cell(4, 1), f"{row['MidtermPercentage']:.2f}%")
            set_cell_properties(content_table.cell(5, 0), '6. Activities/ Measure/special programs\ntaken to improve the performance')
            set_cell_properties(content_table.cell(5, 1), str(row['Actions taken to improve performance']).replace(';', '\n'))
            set_cell_properties(content_table.cell(6, 0), '7. Progress')
            set_cell_properties(content_table.cell(6, 1), str(row['Outcome (Based on clearance in end-semester or makeup exam)']))
            set_cell_properties(content_table.cell(7, 0), 'Comments/remarks')
            set_cell_properties(content_table.cell(7, 1), str(row.get('Remarks if any', '')))

            doc.add_paragraph(f"\nDate:{datetime.now().strftime('%d-%m-%Y')}")
            add_signature_line(doc)
            # --- End Formatting for Format 2 ---

            if index != filtered_df.index.tolist()[-1]:
                doc.add_page_break()
        output_filename = f'{report_prefix}_Format2_Report.docx'

    elif format_choice == '3':
        # --- Start Formatting for Format 3 ---
        doc.add_heading('Format -3   Report of performance/ improvement for slow and advanced learners', level=2).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        grouped = filtered_df.groupby(['Subject Name', 'Semester'])
        
        for (subject, semester), group in grouped:
            doc.add_paragraph(f"Course: {subject}", style='Heading 3')
            doc.add_paragraph(f"Year /Semester: {get_year_semester_string(semester)}", style='Heading 3')
            
            summary_cols = ['Sl. No', 'Reg Number', 'Name of the student', 'Midterm Percentage', 'Progress']
            table = doc.add_table(rows=1, cols=len(summary_cols))
            table.style = 'Table Grid'
            
            hdr_cells = table.rows[0].cells
            for i, col_name in enumerate(summary_cols):
                hdr_cells[i].text = col_name

            for index, row in group.reset_index(drop=True).iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = str(index + 1)
                row_cells[1].text = str(row['Register Number of the Student'])
                row_cells[2].text = str(row['Student Name'])
                row_cells[3].text = f"{row['MidtermPercentage']:.2f}"
                row_cells[4].text = str(row['Outcome (Based on clearance in end-semester or makeup exam)'])
            
            doc.add_paragraph() # Add space between tables
        # --- End Formatting for Format 3 ---
        output_filename = f'{report_prefix}_Format3_Report.docx'

    else:
        print("Invalid format choice.")
        return

    try:
        doc.save(output_filename)
        print(f"\nSuccess! Report generated as '{output_filename}' ✨")
    except Exception as e:
        print(f"\nError: Could not save the file. Details: {e}")


if __name__ == "__main__":
    input_excel_file = input("Enter the file name (if in same folder) or the full file path: ")

    if not os.path.exists(input_excel_file):
        print(f"Error: The file '{input_excel_file}' does not exist.")
    else:
        learner_choice = ''
        while learner_choice not in ['fast', 'slow']:
            learner_choice = input("Generate report for 'fast' or 'slow' learners? ").lower()

        try:
            slow_thresh = float(input("Enter the percentage threshold for SLOW learners (e.g., 40): "))
            fast_thresh = float(input("Enter the percentage threshold for FAST learners (e.g., 80): "))
        except ValueError:
            print("Invalid input. Please enter a number for the thresholds.")
            exit()

        print("\nPlease choose a report format:")
        print("  1: Format 1 - Assessment of learning levels")
        print("  2: Format 2 - Report of performance/improvement")
        print("  3: Format 3 - Tabular Summary Report")
        
        format_choice = ''
        while format_choice not in ['1', '2', '3']:
            format_choice = input("Enter your choice (1, 2, or 3): ")
        
        generate_word_report(input_excel_file, format_choice, learner_choice, slow_thresh, fast_thresh)
