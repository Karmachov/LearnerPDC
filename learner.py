import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from fpdf import FPDF 
from datetime import datetime
import os

# ==============================================================================
# 1. READER
# ==============================================================================
class DataReader:
    """Reads data from an Excel file and returns it in a neutral format."""
    def read_excel(self, file_path):
        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            return df.to_dict('records')
        except FileNotFoundError:
            print(f"Error: The file '{file_path}' was not found.")
            return None
        except Exception as e:
            print(f"An error occurred while reading the Excel file: {e}")
            return None

# ==============================================================================
# 2. WRITERS
# ==============================================================================
class DocxWriter:
    """Takes a Document object and saves it to a .docx file."""
    def write(self, doc, output_filename):
        try:
            doc.save(output_filename)
            print(f"\nSuccess! Report generated as '{output_filename}' ✨")
        except Exception as e:
            print(f"\nError: Could not save the file. Details: {e}")

class PdfWriter:
    """Takes an FPDF object and saves it to a .pdf file."""
    def write(self, pdf, output_filename):
        try:
            pdf.output(output_filename)
            print(f"\nSuccess! Report generated as '{output_filename}' ✨")
        except Exception as e:
            print(f"\nError: Could not save the file. Details: {e}")

# ==============================================================================
# 3. FORMATTERS
# ==============================================================================
class BaseFormatter:
    """Base class for all formatters with shared helper methods."""
    def get_year_semester_string(self, roman_numeral):
        mapping = {'I': 'I Year/ I semester', 'II': 'I Year/ II semester', 'III': 'II Year/ III semester'}
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

class Format1DocxFormatter(BaseFormatter):
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
            self.set_cell_properties(student_info_table.cell(2, 1), str(student['Subject Name']).title())
            self.set_cell_properties(student_info_table.cell(3, 0), 'Year /semester:')
            self.set_cell_properties(student_info_table.cell(3, 1), self.get_year_semester_string(student['Semester']))
            params_cell = container_table.cell(2, 0)
            params_cell.text = ''
            params_table = params_cell.add_table(rows=3, cols=4)
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

class Format2DocxFormatter(BaseFormatter):
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
            self.set_cell_properties(content_table.cell(2, 1), str(student['Subject Name']).title())
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

class Format3DocxFormatter(BaseFormatter):
    """Creates the Word document for Format 3."""
    def format(self, students, slow_threshold, fast_threshold):
        doc = Document()
        df = pd.DataFrame(students)
        grouped = df.groupby(['Subject Name', 'Semester'])
        for i, ((subject, semester), group) in enumerate(grouped):
            doc.add_paragraph(f"Course: {str(subject).title()}", style='Heading 3')
            doc.add_paragraph(f"Year /Semester: {self.get_year_semester_string(semester)}", style='Heading 3')
            summary_cols = ['Sl. No', 'Reg Number', 'Name of the student', 'Midterm Percentage', 'Progress']
            table = doc.add_table(rows=1, cols=len(summary_cols))
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            for j, col_name in enumerate(summary_cols):
                hdr_cells[j].text = col_name
            for index, row_data in group.reset_index(drop=True).iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = str(index + 1)
                row_cells[1].text = str(row_data['Register Number of the Student'])
                row_cells[2].text = str(row_data['Student Name'])
                row_cells[3].text = f"{row_data['MidtermPercentage']:.2f}"
                row_cells[4].text = str(row_data['Outcome (Based on clearance in end-semester or makeup exam)'])
            
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.paragraph_format.keep_with_next = True

            if i < len(grouped) - 1:
                doc.add_page_break()
        return doc

class Format1And2DocxFormatter(BaseFormatter):
    """Creates a combined report with Format 1 and 2 for each student."""
    def format(self, students, slow_threshold, fast_threshold):
        doc = Document()
        for i, student in enumerate(students):
            # --- Build Format 1 Page ---
            title1 = doc.add_heading('Format 1.   Assessment of the learning levels of the students:', level=2)
            title1.alignment = WD_ALIGN_PARAGRAPH.CENTER
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
            self.set_cell_properties(student_info_table.cell(2, 1), str(student['Subject Name']).title())
            self.set_cell_properties(student_info_table.cell(3, 0), 'Year /semester:')
            self.set_cell_properties(student_info_table.cell(3, 1), self.get_year_semester_string(student['Semester']))
            params_cell = container_table.cell(2, 0)
            params_cell.text = ''
            params_table = params_cell.add_table(rows=3, cols=4)
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
            
            doc.add_page_break()

            # --- Build Format 2 Page ---
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
            self.set_cell_properties(content_table.cell(2, 1), str(student['Subject Name']).title())
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

class PdfFormatter(BaseFormatter):
    """Creates a PDF document."""
    def format(self, students, slow_threshold, fast_threshold):
        pdf = FPDF()
        # ... (PDF formatting logic would go here) ...
        return pdf

# ==============================================================================
# 4. CONTROLLER
# ==============================================================================
class ReportController:
    """Controls the report generation workflow."""
    def __init__(self, excel_path, format_choice, learner_type, slow_thresh, fast_thresh, output_type, subject, semester):
        self.excel_path = excel_path
        self.format_choice = format_choice
        self.learner_type = learner_type
        self.slow_threshold = slow_thresh
        self.fast_threshold = fast_thresh
        self.output_type = output_type
        self.subject = subject.lower().strip() # Standardize user input
        self.semester = semester.lower().strip() # Standardize user input
        self.reader = DataReader()
        
        if self.output_type == 'word':
            self.writer = DocxWriter()
            self.formatters = {
                '1': Format1DocxFormatter(),
                '2': Format2DocxFormatter(),
                '3': Format3DocxFormatter(),
                '4': Format1And2DocxFormatter(),
            }
        elif self.output_type == 'pdf':
            self.writer = PdfWriter()
            self.formatters = {'1': PdfFormatter()}
        else:
            self.writer = None
            self.formatters = {}

    def _calculate_midterm_percentage(self, marks):
        try:
            return (float(marks) / 30) * 100
        except (ValueError, TypeError):
            return 0

    def run(self):
        if not self.writer:
            print("Invalid output type selected.")
            return

        # 1. Read data
        all_student_data = self.reader.read_excel(self.excel_path)
        if not all_student_data: return

        # 2. Perform business logic
        for student in all_student_data:
            student['MidtermPercentage'] = self._calculate_midterm_percentage(student['Midterm Exam Marks (Out of 30)'])
            student['Subject Name'] = str(student.get('Subject Name', '')).strip().lower()
            student['Semester'] = str(student.get('Semester', '')).strip().lower()
        
        filtered_by_course = [s for s in all_student_data if (self.semester == 'all' or s['Semester'] == self.semester) and (self.subject == 'all' or s['Subject Name'] == self.subject)]

        if self.learner_type == 'slow':
            final_filtered_students = [s for s in filtered_by_course if s['MidtermPercentage'] <= self.slow_threshold]
        else:
            final_filtered_students = [s for s in filtered_by_course if s['MidtermPercentage'] >= self.fast_threshold]

        if not final_filtered_students:
            print(f"\nNo students found for the selected criteria.")
            return
            
        final_filtered_students.sort(key=lambda s: s.get('Subject Name', ''))
        
        if self.format_choice == '5':
            print(f"\nFound {len(final_filtered_students)} {self.learner_type} learners. Generating combined and summary reports...")
            f1_and_2_formatter = Format1And2DocxFormatter()
            doc1 = f1_and_2_formatter.format(final_filtered_students, self.slow_threshold, self.fast_threshold)
            self.writer.write(doc1, f'{self.learner_type.title()}_Learners_Combined_Report.docx')

            f3_formatter = Format3DocxFormatter()
            doc2 = f3_formatter.format(final_filtered_students, self.slow_threshold, self.fast_threshold)
            self.writer.write(doc2, f'{self.learner_type.title()}_Learners_Summary_Report.docx')
            return

        # 3. Format the document (for options 1-4)
        formatter = self.formatters.get(self.format_choice)
        if not formatter:
            print("Invalid format choice.")
            return
        output_object = formatter.format(final_filtered_students, self.slow_threshold, self.fast_threshold)

        # 4. Write the document
        file_extension = 'docx' if self.output_type == 'word' else 'pdf'
        report_name_map = {'1': 'Format1_Report', '2': 'Format2_Report', '3': 'Summary_Report', '4': 'Combined_Report'}
        report_name = report_name_map.get(self.format_choice, "Report")
        output_filename = f'{self.learner_type.title()}_Learners_{report_name}.{file_extension}'
        self.writer.write(output_object, output_filename)

# ==============================================================================
# 5. MAIN EXECUTION BLOCK
# ==============================================================================
if __name__ == "__main__":
    excel_file = input("Enter the file name or full file path: ")
    if not os.path.exists(excel_file):
        print(f"Error: The file '{excel_file}' does not exist.")
    else:
        subject_filter = input("Enter Subject Name to filter by (or type 'all'): ").strip()
        semester_filter = input("Enter Semester to filter by (e.g., 'III' or 'all'): ").strip()

        output_format = ''
        while output_format not in ['word', 'pdf']:
            output_format = input("Choose output format ('word' or 'pdf'): ").lower()
        
        learner = input("Generate report for 'fast' or 'slow' learners? ").lower()
        slow_thresh = float(input("Enter percentage threshold for SLOW learners (e.g., 40): "))
        fast_thresh = float(input("Enter percentage threshold for FAST learners (e.g., 80): "))
        
        print("\nPlease choose a report format:")
        print("  1: Format 1 - Assessment of learning levels")
        print("  2: Format 2 - Report of performance/improvement")
        print("  3: Format 3 - Tabular Summary Report")
        print("  4: Combined Format 1 & 2")
        print("  5: All Formats (Generates 2 separate files)")
        
        format_num = ''
        while format_num not in ['1', '2', '3', '4', '5']:
            format_num = input("Enter your choice (1, 2, 3, 4, or 5): ")

        controller = ReportController(excel_file, format_num, learner, slow_thresh, fast_thresh, output_format, subject_filter, semester_filter)
        controller.run()
        print("\nReport generation completed.")