import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_LAYOUT
from datetime import datetime
import os

# ==============================================================================
# 1. READER: Responsible for reading data from the source.
# ==============================================================================
class DataReader:
    """Reads data from an Excel file and returns it in a neutral format."""

    def read_excel(self, file_path):
        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            # Convert DataFrame to a list of dictionaries for easy use
            return df.to_dict('records')
        except FileNotFoundError:
            print(f"Error: The file '{file_path}' was not found.")
            return None
        except Exception as e:
            print(f"An error occurred while reading the Excel file: {e}")
            return None

# ==============================================================================
# 2. WRITER: Responsible for writing the final output.
# ==============================================================================
class DocxWriter:
    """Takes a Document object and saves it to a .docx file."""

    def write(self, doc, output_filename):
        try:
            doc.save(output_filename)
            print(f"\nSuccess! Report generated as '{output_filename}' ✨")
        except Exception as e:
            print(f"\nError: Could not save the file. Details: {e}")

# ==============================================================================
# 3. FORMATTERS: Responsible for creating the specific Word document layouts.
# ==============================================================================
class BaseFormatter:
    """Base class for all formatters with shared helper methods."""
    
    def get_year_semester_string(self, roman_numeral):
        mapping = {
            'I': 'I Year/ I semester', 'II': 'I Year/ II semester',
            'III': 'II Year/ III semester', 'IV': 'II Year/ IV semester',
            'V': 'III Year/ V semester', 'VI': 'III Year/ VI semester',
            'VII': 'IV Year/ VII semester', 'VIII': 'IV Year/ VIII semester'
        }
        return mapping.get(str(roman_numeral).strip().upper(), str(roman_numeral))

    def set_cell_properties(self, cell, text, bold=False, font_size=10, align='LEFT', valign='TOP'):
        cell.text = ''
        p = cell.add_paragraph()
        run = p.add_run(str(text))
        run.font.size = Pt(font_size)
        run.bold = bold
        p.alignment = getattr(WD_ALIGN_PARAGRAPH, align)
        cell.vertical_alignment = getattr(WD_ALIGN_VERTICAL, valign)

    def add_signature_line(self, doc_or_cell):
        p = doc_or_cell.add_paragraph()
        p.add_run("\n\n" + "_" * 40 + "\n")
        p.add_run("Signature of the\nsubject teacher / class coordinator")
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

class Format1Formatter(BaseFormatter):
    """Creates the Word document for Format 1."""

    def format(self, students, slow_threshold, fast_threshold):
        doc = Document()
        for i, student in enumerate(students):
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
            self.set_cell_properties(student_info_table.cell(0, 0), 'Name of the Student:')
            self.set_cell_properties(student_info_table.cell(0, 1), str(student['Student Name']))
            self.set_cell_properties(student_info_table.cell(1, 0), 'Registration Number:')
            self.set_cell_properties(student_info_table.cell(1, 1), str(student['Register Number of the Student']))
            self.set_cell_properties(student_info_table.cell(2, 0), 'Course:')
            self.set_cell_properties(student_info_table.cell(2, 1), str(student['Subject Name']))
            self.set_cell_properties(student_info_table.cell(3, 0), 'Year /semester:')
            self.set_cell_properties(student_info_table.cell(3, 1), self.get_year_semester_string(student['Semester']))

            params_cell = container_table.cell(2, 0)
            params_cell.text = ''
            params_table = params_cell.add_table(rows=3, cols=4)
            try:
                params_table.layout_algorithm = WD_TABLE_LAYOUT.FIXED
            except NameError:
                pass
            params_table.style = 'Table Grid'

            hdr_cell1 = params_table.cell(0, 2); hdr_cell2 = params_table.cell(0, 3); hdr_cell1.merge(hdr_cell2)
            self.set_cell_properties(params_table.cell(0, 0), 'Sr. No.', bold=True, align='CENTER')
            self.set_cell_properties(params_table.cell(0, 1), 'Parameter', bold=True, align='CENTER')
            self.set_cell_properties(params_table.cell(0, 2), 'Weightage in Percentage', bold=True, align='CENTER')
            
            self.set_cell_properties(params_table.cell(1, 0), '1', align='CENTER')
            self.set_cell_properties(params_table.cell(1, 1), 'Scores obtained by student class test / internal examination...\nConsidered Midterm exam conducted for 30M:')
            self.set_cell_properties(params_table.cell(1, 2), f"{student['MidtermPercentage']:.2f}", align='CENTER')
            self.set_cell_properties(params_table.cell(1, 3), "> %", align='CENTER')
            
            self.set_cell_properties(params_table.cell(2, 0), '2', align='CENTER')
            self.set_cell_properties(params_table.cell(2, 1), 'Performance of students in preceding university examination')
            self.set_cell_properties(params_table.cell(2, 2), str(student['CGPA (up to previous semester)']), align='CENTER')
            self.set_cell_properties(params_table.cell(2, 3), "> %", align='CENTER')

            params_table.columns[0].width = Inches(0.5)
            params_table.columns[1].width = Inches(4.0)
            params_table.columns[2].width = Inches(1.0)
            params_table.columns[3].width = Inches(0.5)

            total_cell = container_table.cell(3, 0)
            total_cell.text = "Total Weightage"

            footer_cell = container_table.cell(4, 0)
            footer_cell.text = ''
            footer_cell.add_paragraph(f"1. Midterm score less than {slow_threshold}% considered as a slow learner")
            footer_cell.add_paragraph(f"2. Midterm score more than {fast_threshold}% considered as an advanced learner **")
            footer_cell.add_paragraph(f"Date: {datetime.now().strftime('%d-%m-%Y')}")
            self.add_signature_line(footer_cell)

            if i < len(students) - 1:
                doc.add_page_break()
        return doc

class Format2Formatter(BaseFormatter):
    """Creates the Word document for Format 2."""
    def format(self, students, slow_threshold, fast_threshold):
        doc = Document()
        for i, student in enumerate(students):
            doc.add_heading('Format -2   Report of performance/ improvement for slow and advanced learners', level=2).alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            header_table = doc.add_table(rows=3, cols=1)
            p1 = header_table.cell(0, 0).paragraphs[0]; p1.add_run('Manipal Institute of Technology').bold = True; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p2 = header_table.cell(1, 0).paragraphs[0]; p2.add_run('MAHE Manipal').bold = True; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p3 = header_table.cell(2, 0).paragraphs[0]; p3.add_run('Computer Science and Engineering Department').bold = True; p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            content_table = doc.add_table(rows=8, cols=2)
            content_table.style = 'Table Grid'
            self.set_cell_properties(content_table.cell(0, 0), '1. Registration Number')
            self.set_cell_properties(content_table.cell(0, 1), student['Register Number of the Student'])
            self.set_cell_properties(content_table.cell(1, 0), '2. Name of the student')
            self.set_cell_properties(content_table.cell(1, 1), student['Student Name'])
            self.set_cell_properties(content_table.cell(2, 0), '3. Course')
            self.set_cell_properties(content_table.cell(2, 1), student['Subject Name'])
            self.set_cell_properties(content_table.cell(3, 0), '4. Year/Semester')
            self.set_cell_properties(content_table.cell(3, 1), self.get_year_semester_string(student['Semester']))
            self.set_cell_properties(content_table.cell(4, 0), '5. Midterm Percentage')
            self.set_cell_properties(content_table.cell(4, 1), f"{student['MidtermPercentage']:.2f}%")
            self.set_cell_properties(content_table.cell(5, 0), '6. Activities/ Measure/special programs\ntaken to improve the performance')
            self.set_cell_properties(content_table.cell(5, 1), str(student['Actions taken to improve performance']).replace(';', '\n'))
            self.set_cell_properties(content_table.cell(6, 0), '7. Progress')
            self.set_cell_properties(content_table.cell(6, 1), str(student['Outcome (Based on clearance in end-semester or makeup exam)']))
            self.set_cell_properties(content_table.cell(7, 0), 'Comments/remarks')
            self.set_cell_properties(content_table.cell(7, 1), str(student.get('Remarks if any', '')))

            doc.add_paragraph(f"\nDate:{datetime.now().strftime('%d-%m-%Y')}")
            self.add_signature_line(doc)

            if i < len(students) - 1:
                doc.add_page_break()
        return doc

class Format3Formatter(BaseFormatter):
    """Creates the Word document for Format 3."""
    def format(self, students, slow_threshold, fast_threshold):
        doc = Document()
        # Group students by subject and semester to create separate tables
        df = pd.DataFrame(students)
        grouped = df.groupby(['Subject Name', 'Semester'])
        
        for i, ((subject, semester), group) in enumerate(grouped):
            doc.add_paragraph(f"Course: {subject}", style='Heading 3')
            doc.add_paragraph(f"Year /Semester: {self.get_year_semester_string(semester)}", style='Heading 3')
            
            summary_cols = ['Sl. No', 'Reg Number', 'Name of the student', 'Midterm Percentage', 'Progress']
            table = doc.add_table(rows=1, cols=len(summary_cols))
            table.style = 'Table Grid'
            
            hdr_cells = table.rows[0].cells
            for j, col_name in enumerate(summary_cols):
                hdr_cells[j].text = col_name

            for index, row in group.reset_index(drop=True).iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = str(index + 1)
                row_cells[1].text = str(row['Register Number of the Student'])
                row_cells[2].text = str(row['Student Name'])
                row_cells[3].text = f"{row['MidtermPercentage']:.2f}"
                row_cells[4].text = str(row['Outcome (Based on clearance in end-semester or makeup exam)'])
            
            if i < len(grouped) - 1:
                doc.add_page_break()
        return doc

# ==============================================================================
# 4. CONTROLLER: The "brain" that orchestrates the process.
# ==============================================================================
class ReportController:
    """Controls the report generation workflow."""

    def __init__(self, excel_path, format_choice, learner_type, slow_threshold, fast_threshold):
        self.excel_path = excel_path
        self.format_choice = format_choice
        self.learner_type = learner_type
        self.slow_threshold = slow_threshold
        self.fast_threshold = fast_threshold
        self.reader = DataReader()
        self.writer = DocxWriter()
        self.formatters = {
            '1': Format1Formatter(),
            '2': Format2Formatter(),
            '3': Format3Formatter(),
        }

    def _calculate_midterm_percentage(self, marks):
        try:
            return (float(marks) / 30) * 100
        except (ValueError, TypeError):
            return 0

    def run(self):
        # 1. Read data
        student_data = self.reader.read_excel(self.excel_path)
        if not student_data:
            return

        # 2. Perform business logic (calculate and filter)
        for student in student_data:
            student['MidtermPercentage'] = self._calculate_midterm_percentage(student['Midterm Exam Marks (Out of 30)'])

        if self.learner_type == 'slow':
            filtered_students = [s for s in student_data if s['MidtermPercentage'] <= self.slow_threshold]
            report_prefix = "Slow_Learners"
        elif self.learner_type == 'fast':
            filtered_students = [s for s in student_data if s['MidtermPercentage'] >= self.fast_threshold]
            report_prefix = "Fast_Learners"
        else:
            print("Invalid learner type.")
            return

        if not filtered_students:
            print(f"\nNo students found for the '{self.learner_type}' category.")
            return
        
        print(f"\nFound {len(filtered_students)} {self.learner_type} learners. Generating report...")

        # 3. Select the correct formatter
        formatter = self.formatters.get(self.format_choice)
        if not formatter:
            print("Invalid format choice.")
            return
        
        # 4. Format the document
        doc = formatter.format(filtered_students, self.slow_threshold, self.fast_threshold)

        # 5. Write the document
        output_filename = f'{report_prefix}_Format{self.format_choice}_Report.docx'
        self.writer.write(doc, output_filename)

# ==============================================================================
# 5. MAIN EXECUTION BLOCK: Gathers user input and starts the controller.
# ==============================================================================
if __name__ == "__main__":
    # Get user input
    excel_file = input("Enter the file name or full file path: ")
    if not os.path.exists(excel_file):
        print(f"Error: The file '{excel_file}' does not exist.")
    else:
        learner = input("Generate report for 'fast' or 'slow' learners? ").lower()
        slow_thresh = float(input("Enter percentage threshold for SLOW learners (e.g., 40): "))
        fast_thresh = float(input("Enter percentage threshold for FAST learners (e.g., 80): "))
        
        print("\nPlease choose a report format:")
        print("  1: Format 1 - Assessment of learning levels")
        print("  2: Format 2 - Report of performance/improvement")
        print("  3: Format 3 - Tabular Summary Report")
        
        format_num = input("Enter your choice (1, 2, or 3): ")

        # Create and run the controller
        controller = ReportController(excel_file, format_num, learner, slow_thresh, fast_thresh)
        controller.run()
