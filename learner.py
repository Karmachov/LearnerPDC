import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from datetime import datetime, timedelta, timezone
import os
import tempfile
import platform
import time
import getpass
import traceback
from PyPDF2 import PdfReader
import shutil
import warnings

# Suppress the PyPDF2 warning about using PdfReader instead of PdfFileReader
warnings.filterwarnings("ignore", category=DeprecationWarning)

try:
    from docx2pdf import convert
except ImportError:
    print("Warning: 'docx2pdf' module not found. PDF output will not be available.")
    convert = None

try:
    from endesive import pdf
    from cryptography.hazmat.primitives.serialization import load_pem_private_key
    from cryptography.x509 import load_pem_x509_certificate
except ImportError:
    print("Warning: Required crypto libraries not found. PDF signing will not be available.")
    print("Please run: pip install endesive cryptography")
    pdf = None

# ==============================================================================
# 0. CONSTANTS & CONFIGURATION
# ==============================================================================
MIDTERM_TOTAL_MARKS = 30
SEMESTER_MAPPING = {
    'i': 'I Year/ I semester',
    'ii': 'I Year/ II semester',
    'iii': 'II Year/ III semester'
}
CERTIFICATE_INFO = {
    "country": "IN",
    "state": "Karnataka",
    "locality": "Manipal",
    "org": "Manipal Institute of Technology",
}


def sign_pdf(pdf_path, key_path, cert_path, image_path, password):
    """Signs a PDF with a visible signature on every page, one page at a time."""
    if not all([pdf, key_path, cert_path, image_path, password]):
        print("Skipping signing due to missing information.")
        return

    try:
        single_page_box = (435, 72, 540, 105)
        date = datetime.now().strftime('D:%Y%m%d%H%M%S+05\'30\'')

        with open(key_path, 'rb') as f:
            private_key = load_pem_private_key(f.read(), password=password.encode('utf-8'))
        with open(cert_path, 'rb') as f:
            certificate = load_pem_x509_certificate(f.read())

        with open(pdf_path, 'rb') as f:
            pdf_data = f.read()

        reader = PdfReader(pdf_path)
        page_count = len(reader.pages)

        for i in range(page_count):
            print(f"Signing page {i + 1}/{page_count}...")
            signdata = {
                'sigflags': 3,
                'contact': 'faculty.email@example.com',
                'location': 'Manipal, India',
                'reason': 'I am the author of this document',
                'signaturebox': single_page_box,
                'signature_img': image_path,
                'signingdate': date,
                'page': i
            }

            signed_data_obj = pdf.cms.sign(
                pdf_data,
                signdata,
                key=private_key,
                cert=certificate,
                othercerts=()
            )

            pdf_data += signed_data_obj

        with open(pdf_path, 'wb') as f:
            f.write(pdf_data)

        print(f"\nSuccess! Successfully signed all {page_count} pages of '{pdf_path}' ✨")

    except Exception:
        print(f"\nCRITICAL ERROR: Failed to sign the PDF.")
        print("Please check that the key/certificate paths and password are correct.")
        print("----- Full Error Details -----")
        traceback.print_exc()
        print("----------------------------")

# ==============================================================================
# 1. READER
# ==============================================================================
class DataReader:
    """Reads data from an Excel or CSV file and returns it in a neutral format."""
    def read_data(self, file_path):
        try:
            if file_path.lower().endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            return df.to_dict('records')
        except FileNotFoundError:
            print(f"Error: The file '{file_path}' was not found. Please check the path and try again.")
            return None
        except Exception as e:
            print(f"An error occurred while reading the data file: {e}")
            return None

# ==============================================================================
# 2. WRITERS
# ==============================================================================
class DocxWriter:
    """Takes a Document object and saves it to a .docx file."""
    def write(self, doc, output_filename, **kwargs):
        try:
            doc.save(output_filename)
            print(f"\nSuccess! Report generated as '{output_filename}' ✨")
        except Exception as e:
            print(f"\nError: Could not save the file. Details: {e}")

class PdfWriter:
    """Creates a DOCX, converts it to PDF, and optionally signs it."""
    def write(self, doc, output_filename, sign_info=None, format_choice=None):
        if convert is None:
            print("\nError: 'docx2pdf' not installed. Cannot generate PDF.")
            return

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_docx = os.path.join(temp_dir, "temp_report.docx")
                doc.save(temp_docx)
                convert(temp_docx, output_filename)

            time.sleep(2)

            if platform.system() == "Darwin":
                time.sleep(1)

            print(f"\nSuccess! PDF generated as '{output_filename}' ✨")

            if sign_info and sign_info.get('should_sign'):
                if format_choice in ['1', '2', '4', '5']:
                    print("Proceeding to sign the PDF...")
                    sign_pdf(
                        pdf_path=output_filename,
                        key_path=sign_info['key_path'],
                        cert_path=sign_info['cert_path'],
                        image_path=sign_info['image_path'],
                        password=sign_info['password']
                    )
                else:
                    print("Skipping signature for Format 3 (Summary Report).")

        except Exception as e:
            print(f"\nError: Could not save or convert the file. Details: {e}")
            print("Please ensure Microsoft Word (on Windows) or LibreOffice (on macOS/Linux) is installed and accessible.")

# ==============================================================================
# 3. FORMATTERS
# ==============================================================================
class BaseFormatter:
    """Base class for all formatters with shared helper methods."""
    COMIC_SANS = "Brush Script MT Italic"
    def get_year_semester_string(self, roman_numeral):
        return SEMESTER_MAPPING.get(str(roman_numeral).strip().lower(), str(roman_numeral))

    def set_cell_properties(self, cell, text, bold=False, font_size=10, align='LEFT', valign='TOP', font_name=None):
        cell.text = ''
        p = cell.add_paragraph()
        run = p.add_run(str(text))
        run.font.size = Pt(font_size)
        run.bold = bold
        if font_name:
            run.font.name = font_name
        try:
            p.alignment = getattr(WD_ALIGN_PARAGRAPH, str(align).upper())
        except AttributeError:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        try:
            cell.vertical_alignment = getattr(WD_ALIGN_VERTICAL, str(valign).upper())
        except AttributeError:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    def add_signature_line(self, doc_or_cell):
        p = doc_or_cell.add_paragraph()
        p.add_run("\n\n" + "_" * 40 + "\n")
        p.add_run("Signature of the\nsubject teacher / class coordinator")
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    def _create_format1_content(self, doc, student, slow_threshold, fast_threshold):
        doc.add_heading('Format 1. Assessment of the learning levels of the students:', level=2).alignment = WD_ALIGN_PARAGRAPH.CENTER
        container_table = doc.add_table(rows=5, cols=1)
        container_table.style = 'Table Grid'
        header_cell = container_table.cell(0, 0)
        p1 = header_cell.add_paragraph(); p1.add_run('Manipal Institute of Technology').bold = True; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2 = header_cell.add_paragraph(); p2.add_run('MAHE Manipal').bold = True; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p3 = header_cell.add_paragraph(); p3.add_run('Computer Science and Engineering Department').bold = True; p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        student_info_table = container_table.cell(1, 0).add_table(rows=4, cols=2)
        self.set_cell_properties(student_info_table.cell(0, 0), 'Name of the Student:')
        self.set_cell_properties(student_info_table.cell(0, 1), str(student.get('Student Name', '')), font_name=self.COMIC_SANS)
        self.set_cell_properties(student_info_table.cell(1, 0), 'Registration Number:')
        self.set_cell_properties(student_info_table.cell(1, 1), str(student.get('Register Number of the Student', '')), font_name=self.COMIC_SANS)
        self.set_cell_properties(student_info_table.cell(2, 0), 'Course:')
        self.set_cell_properties(student_info_table.cell(2, 1), str(student.get('Subject Name', '')).title(), font_name=self.COMIC_SANS)
        self.set_cell_properties(student_info_table.cell(3, 0), 'Year /semester:')
        self.set_cell_properties(student_info_table.cell(3, 1), self.get_year_semester_string(student.get('Semester', '')), font_name=self.COMIC_SANS)
        params_table = container_table.cell(2, 0).add_table(rows=3, cols=4)
        params_table.style = 'Table Grid'
        hdr_cell1 = params_table.cell(0, 2); hdr_cell2 = params_table.cell(0, 3); hdr_cell1.merge(hdr_cell2)
        self.set_cell_properties(params_table.cell(0, 0), 'Sr. No.', bold=True, align='CENTER')
        self.set_cell_properties(params_table.cell(0, 1), 'Parameter', bold=True, align='CENTER')
        self.set_cell_properties(params_table.cell(0, 2), 'Weightage in Percentage', bold=True, align='CENTER')
        self.set_cell_properties(params_table.cell(1, 0), '1', align='CENTER')
        self.set_cell_properties(params_table.cell(1, 1), f"Scores obtained by student class test / internal examination...\nConsidered Midterm exam conducted for {MIDTERM_TOTAL_MARKS}M:")
        self.set_cell_properties(params_table.cell(1, 2), f"{student.get('MidtermPercentage', 0):.2f}", align='CENTER', font_name=self.COMIC_SANS)
        self.set_cell_properties(params_table.cell(1, 3), "> %", align='CENTER')
        self.set_cell_properties(params_table.cell(2, 0), '2', align='CENTER')
        self.set_cell_properties(params_table.cell(2, 1), 'Performance of students in preceding university examination')
        self.set_cell_properties(params_table.cell(2, 2), str(student.get('CGPA (up to previous semester)', 'N/A')), align='CENTER', font_name=self.COMIC_SANS)
        self.set_cell_properties(params_table.cell(2, 3), "> %", align='CENTER')
        params_table.columns[0].width = Inches(0.5); params_table.columns[1].width = Inches(4.0); params_table.columns[2].width = Inches(1.0); params_table.columns[3].width = Inches(0.5)
        container_table.cell(3, 0).text = "Total Weightage"
        footer_cell = container_table.cell(4, 0)
        p_footer_1 = footer_cell.add_paragraph(f"1. Midterm score less than {slow_threshold}% considered as a slow learner")
        p_footer_2 = footer_cell.add_paragraph(f"2. Midterm score more than {fast_threshold}% considered as an advanced learner **")
        p_date = footer_cell.add_paragraph()
        run_date = p_date.add_run(f"Date: {datetime.now().strftime('%d-%m-%Y')}")
        run_date.font.name = self.COMIC_SANS
        self.add_signature_line(footer_cell)

    def _create_format2_content(self, doc, student):
        doc.add_heading('Format -2 Report of performance/ improvement for slow and advanced learners', level=2).alignment = WD_ALIGN_PARAGRAPH.CENTER
        header_table = doc.add_table(rows=3, cols=1)
        p1 = header_table.cell(0, 0).paragraphs[0]; p1.add_run('Manipal Institute of Technology').bold = True; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2 = header_table.cell(1, 0).paragraphs[0]; p2.add_run('MAHE Manipal').bold = True; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p3 = header_table.cell(2, 0).paragraphs[0]; p3.add_run('Computer Science and Engineering Department').bold = True; p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        content_table = doc.add_table(rows=8, cols=2)
        content_table.style = 'Table Grid'
        self.set_cell_properties(content_table.cell(0, 0), '1. Registration Number')
        self.set_cell_properties(content_table.cell(0, 1), str(student.get('Register Number of the Student', '')), font_name=self.COMIC_SANS)
        self.set_cell_properties(content_table.cell(1, 0), '2. Name of the student')
        self.set_cell_properties(content_table.cell(1, 1), str(student.get('Student Name', '')), font_name=self.COMIC_SANS)
        self.set_cell_properties(content_table.cell(2, 0), '3. Course')
        self.set_cell_properties(content_table.cell(2, 1), str(student.get('Subject Name', '')).title(), font_name=self.COMIC_SANS)
        self.set_cell_properties(content_table.cell(3, 0), '4. Year/Semester')
        self.set_cell_properties(content_table.cell(3, 1), self.get_year_semester_string(student.get('Semester', '')), font_name=self.COMIC_SANS)
        self.set_cell_properties(content_table.cell(4, 0), '5. Midterm Percentage')
        self.set_cell_properties(content_table.cell(4, 1), f"{student.get('MidtermPercentage', 0):.2f}%", font_name=self.COMIC_SANS)
        self.set_cell_properties(content_table.cell(5, 0), '6. Activities/ Measure/special programs\ntaken to improve the performance')
        self.set_cell_properties(content_table.cell(5, 1), str(student.get('Actions taken to improve performance', '')).replace(';', '\n'), font_name=self.COMIC_SANS)
        self.set_cell_properties(content_table.cell(6, 0), '7. Progress')
        self.set_cell_properties(content_table.cell(6, 1), str(student.get('Outcome (Based on clearance in end-semester or makeup exam)', '')), font_name=self.COMIC_SANS)
        self.set_cell_properties(content_table.cell(7, 0), 'Comments/remarks')
        self.set_cell_properties(content_table.cell(7, 1), str(student.get('Remarks if any', '')), font_name=self.COMIC_SANS)
        p_date = doc.add_paragraph()
        run_date = p_date.add_run(f"\nDate:{datetime.now().strftime('%d-%m-%Y')}")
        run_date.font.name = self.COMIC_SANS
        self.add_signature_line(doc)

    def _generate_pages(self, doc, students, content_method, *args):
        for i, student in enumerate(students):
            content_method(doc, student, *args)
            if i < len(students) - 1:
                doc.add_page_break()
        return doc
class Format1DocxFormatter(BaseFormatter):
    def format(self, students, slow_threshold, fast_threshold):
        doc = Document()
        return self._generate_pages(doc, students, self._create_format1_content, slow_threshold, fast_threshold)
class Format2DocxFormatter(BaseFormatter):
    def format(self, students, slow_threshold, fast_threshold):
        doc = Document()
        return self._generate_pages(doc, students, self._create_format2_content)
class Format3DocxFormatter(BaseFormatter):
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
            for j, col_name in enumerate(summary_cols):
                self.set_cell_properties(table.cell(0, j), col_name, bold=True)
            for index, row_data in group.reset_index(drop=True).iterrows():
                row_cells = table.add_row().cells
                self.set_cell_properties(row_cells[0], str(index + 1), font_name=self.COMIC_SANS)
                self.set_cell_properties(row_cells[1], str(row_data.get('Register Number of the Student', '')), font_name=self.COMIC_SANS)
                self.set_cell_properties(row_cells[2], str(row_data.get('Student Name', '')), font_name=self.COMIC_SANS)
                self.set_cell_properties(row_cells[3], f"{row_data.get('MidtermPercentage', 0):.2f}", font_name=self.COMIC_SANS)
                self.set_cell_properties(row_cells[4], str(row_data.get('Outcome (Based on clearance in end-semester or makeup exam)', '')), font_name=self.COMIC_SANS)
            if i < len(grouped) - 1:
                doc.add_page_break()
        return doc
class Format1And2DocxFormatter(BaseFormatter):
    def format(self, students, slow_threshold, fast_threshold):
        doc = Document()
        for i, student in enumerate(students):
            self._create_format1_content(doc, student, slow_threshold, fast_threshold)
            doc.add_page_break()
            self._create_format2_content(doc, student)
            if i < len(students) - 1:
                doc.add_page_break()
        return doc

# ==============================================================================
# 5. FACTORIES
# ==============================================================================
def get_writer(output_type):
    if output_type == 'word':
        return DocxWriter()
    elif output_type == 'pdf':
        if convert is None:
            raise ValueError("PDF output is not available. 'docx2pdf' module not found.")
        return PdfWriter()
    raise ValueError(f"Unsupported output type: {output_type}")
def get_formatter(format_choice):
    formatters = {
        '1': Format1DocxFormatter(),
        '2': Format2DocxFormatter(),
        '3': Format3DocxFormatter(),
        '4': Format1And2DocxFormatter(),
    }
    formatter = formatters.get(format_choice)
    if formatter:
        return formatter
    raise ValueError(f"Unsupported format choice '{format_choice}'")

# ==============================================================================
# 6. DATA PROCESSOR
# ==============================================================================
class StudentDataProcessor:
    def _calculate_midterm_percentage(self, marks):
        try:
            return (float(marks) / MIDTERM_TOTAL_MARKS) * 100
        except (ValueError, TypeError):
            return 0
    def process_data(self, all_student_data):
        for student in all_student_data:
            student['MidtermPercentage'] = self._calculate_midterm_percentage(student.get('Midterm Exam Marks (Out of 30)'))
            student['Subject Name'] = str(student.get('Subject Name', '')).strip().lower()
            student['Semester'] = str(student.get('Semester', '')).strip().lower()
        return all_student_data
    def filter_students(self, students, semester, subject, learner_type, slow_thresh, fast_thresh):
        filtered_by_course = [
            s for s in students if
            (semester == 'all' or s['Semester'] == semester) and
            (subject == 'all' or s['Subject Name'] == subject)
        ]
        if learner_type == 'slow':
            final_filtered = [s for s in filtered_by_course if s['MidtermPercentage'] <= slow_thresh]
        else:
            final_filtered = [s for s in filtered_by_course if s['MidtermPercentage'] >= fast_thresh]
        if subject == 'all':
            final_filtered.sort(key=lambda s: (s.get('Subject Name', ''), s.get('Register Number of the Student', '')))
        else:
            final_filtered.sort(key=lambda s: s.get('Register Number of the Student', ''))
        return final_filtered

# ==============================================================================
# 6. CONTROLLER
# ==============================================================================
class ReportController:
    def __init__(self, excel_path, format_choice, learner_type, slow_thresh, fast_thresh, output_type, subject, semester, sign_info):
        self.excel_path = excel_path
        self.format_choice = format_choice
        self.learner_type = learner_type
        self.slow_threshold = slow_thresh
        self.fast_threshold = fast_thresh
        self.output_type = output_type
        self.subject = subject.lower().strip()
        self.semester = semester.lower().strip()
        self.sign_info = sign_info
        self.reader = DataReader()
        self.processor = StudentDataProcessor()
        try:
            self.writer = get_writer(output_type)
        except ValueError as e:
            print(e)
            self.writer = None
    def run(self):
        if not self.writer:
            print("Report generation halted due to invalid writer configuration.")
            return
        all_student_data = self.reader.read_data(self.excel_path)
        if not all_student_data: return
        processed_students = self.processor.process_data(all_student_data)
        final_filtered_students = self.processor.filter_students(
            processed_students, self.semester, self.subject, self.learner_type,
            self.slow_threshold, self.fast_threshold
        )
        if not final_filtered_students:
            print(f"\nNo students found for the selected criteria (Subject: '{self.subject}', Semester: '{self.semester}').")
            return
        print(f"\nFound {len(final_filtered_students)} {self.learner_type} learners.")
        date_str = datetime.now().strftime('%d_%m_%y')
        base_dir = "Learner Monitor Reports"
        learner_folder = f"{self.learner_type.title()} Learners"
        semester_name_for_file = "AllSemesters" if self.semester == 'all' else self.semester.upper()
        subject_name_for_file = "AllSubjects" if self.subject == 'all' else self.subject.replace(' ', '_').title()
        semester_folder = f"Semester_{semester_name_for_file}" if self.semester != 'all' else semester_name_for_file
        subject_folder = subject_name_for_file
        output_dir = os.path.join(base_dir, learner_folder, semester_folder, subject_folder)
        try:
            os.makedirs(output_dir, exist_ok=True)
        except Exception as e:
            print(f"Error creating output directory: {e}")
            return
        if self.format_choice == '5':
            self._generate_all_formats(final_filtered_students, output_dir, date_str, semester_name_for_file, subject_name_for_file)
            return
        try:
            formatter = get_formatter(self.format_choice)
            output_object = formatter.format(final_filtered_students, self.slow_threshold, self.fast_threshold)
            file_extension = 'docx' if self.output_type == 'word' else 'pdf'
            report_name_map = {'1': 'Format1', '2': 'Format2', '3': 'Summary', '4': 'Combined'}
            report_name = report_name_map.get(self.format_choice, "Report")
            output_filename = f'{subject_name_for_file}_{semester_name_for_file}_{self.learner_type.title()}Learner_{report_name}_{date_str}.{file_extension}'
            full_output_path = os.path.join(output_dir, output_filename)
            self.writer.write(output_object, full_output_path, sign_info=self.sign_info, format_choice=self.format_choice)
        except ValueError as e:
            print(e)
    def _generate_all_formats(self, students, output_dir, date_str, semester_name_for_file, subject_name_for_file):
        file_extension = 'docx' if self.output_type == 'word' else 'pdf'
        combined_filename = f'{subject_name_for_file}_{semester_name_for_file}_{self.learner_type.title()}Learner_Combined_Report_{date_str}.{file_extension}'
        summary_filename = f'{subject_name_for_file}_{semester_name_for_file}_{self.learner_type.title()}Learner_Summary_Report_{date_str}.{file_extension}'
        f1_and_2_formatter = Format1And2DocxFormatter()
        doc1 = f1_and_2_formatter.format(students, self.slow_threshold, self.fast_threshold)
        self.writer.write(doc1, os.path.join(output_dir, combined_filename), sign_info=self.sign_info, format_choice='5')
        f3_formatter = Format3DocxFormatter()
        doc2 = f3_formatter.format(students, self.slow_threshold, self.fast_threshold)
        self.writer.write(doc2, os.path.join(output_dir, summary_filename), sign_info=self.sign_info, format_choice='3')

# ==============================================================================
# 7. MAIN EXECUTION BLOCK
# ==============================================================================
def get_valid_input(prompt, valid_options=None, input_type=str):
    """Helper function to get and validate user input."""
    while True:
        user_input = input(prompt)
        try:
            converted_input = input_type(user_input)
            if valid_options and converted_input not in valid_options:
                print(f"Invalid input. Please choose from {valid_options}.")
                continue
            return converted_input
        except ValueError:
            print("Invalid input. Please enter a valid number.")

if __name__ == "__main__":
    print("Welcome to the Student Learner Report Generator.")
    excel_file = input("Enter the Excel file name or full file path: ")
    if not os.path.exists(excel_file):
        print(f"Error: The file '{excel_file}' does not exist.")
        exit()
    subject_filter = input("Enter Subject Name to filter by (or 'all' for all subjects): ").strip()
    semester_filter = input("Enter Semester to filter by (e.g., 'III' or 'all'): ").strip()
    output_format = get_valid_input("Choose output format ('word' or 'pdf'): ", ['word', 'pdf'])
    learner = get_valid_input("Generate report for 'fast' or 'slow' learners? ", ['fast', 'slow'])
    slow_thresh = get_valid_input("Enter percentage threshold for SLOW learners (e.g., 40): ", input_type=float)
    fast_thresh = get_valid_input("Enter percentage threshold for FAST learners (e.g., 80): ", input_type=float)
    print("\nPlease choose a report format:")
    print(" 1: Format 1 - Assessment of learning levels")
    print(" 2: Format 2 - Report of performance/improvement")
    print(" 3: Format 3 - Tabular Summary Report")
    print(" 4: Combined Format 1 & 2")
    print(" 5: All Formats (Generates 2 separate files)")
    format_num = get_valid_input("Enter your choice (1, 2, 3, 4, or 5): ", ['1', '2', '3', '4', '5'])
    
    signing_data = {'should_sign': False}
    if output_format == 'pdf':
        sign_choice = get_valid_input("\nDo you want to digitally sign the PDF report? (y/n): ", ['y', 'n'])
        if sign_choice == 'y':
            if pdf is None:
                print("\nCannot sign PDF because the required libraries are not installed.")
                print("Please install them with: pip install endesive cryptography")
            else:
                print("\nPlease provide the paths to your signing assets.")
                key_file = input("Path to your private key file (e.g., private_key.pem): ")
                cert_file = input("Path to your certificate file (e.g., certificate.pem): ")
                image_file = input("Path to your signature image file (e.g., signature.png): ")
                password = getpass.getpass("Enter the password for your private key: ")
                
                signing_data = {
                    'should_sign': True,
                    'key_path': key_file,
                    'cert_path': cert_file,
                    'image_path': image_file,
                    'password': password
                }
    controller = ReportController(excel_file, format_num, learner, slow_thresh, fast_thresh, output_format, subject_filter, semester_filter, signing_data)
    controller.run()
    print("\nReport generation completed.")