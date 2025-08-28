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
import shutil

# Optional imports for PDF conversion/signing.
try:
    from docx2pdf import convert

    # Cryptography library components
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography import x509
    from cryptography.x509.oid import NameOID

    # PyHanko library components for PDF signing
    from pyhanko import stamp
    from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
    from pyhanko.sign import fields, signers
    from pyhanko.sign.fields import SigFieldSpec
    
    LIBS_AVAILABLE = True
except ImportError as e:
    print(f"Warning: A required module is missing. PDF generation or signing may fail. Error: {e}")
    print("Please run: pip install pandas openpyxl python-docx docx2pdf cryptography pyhanko")
    convert = None
    LIBS_AVAILABLE = False

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
# 2. PDF SIGNER
# ==============================================================================
class PdfSigner:
    """
    Auto-generates and reuses a key and self-signed certificate for a given faculty
    name, and signs PDF documents using pyHanko.
    """
    def __init__(self, faculty_name: str):
        self.faculty_name = faculty_name
        safe_name = str(faculty_name).replace(' ', '_').lower()
        self.private_key_path = f"{safe_name}_private_key.pem"
        self.cert_path = f"{safe_name}_cert.pem"

        if not LIBS_AVAILABLE:
            self.available = False
            print("PDF signing libraries not available. Install 'cryptography' and 'pyhanko'.")
            return

        self.available = True
        try:
            self.private_key, self.certificate = self._load_or_create_credentials()
            
            self.signer = signers.SimpleSigner(
                signing_cert=self.certificate,
                signing_key=self.private_key,
                other_certs_to_embed=(self.certificate,)
            )
        except Exception as e:
            print(f"Error initializing signer for {self.faculty_name}: {e}")
            self.available = False
            
    def _load_or_create_credentials(self):
        """Loads existing key/cert files or generates new ones."""
        if os.path.exists(self.private_key_path) and os.path.exists(self.cert_path):
            with open(self.private_key_path, "rb") as f:
                private_key = serialization.load_pem_private_key(f.read(), password=None)
            with open(self.cert_path, "rb") as f:
                cert = x509.load_pem_x509_certificate(f.read())
            return private_key, cert

        print(f"Generating new key and certificate for '{self.faculty_name}'...")
        private_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        
        with open(self.private_key_path, "wb") as f:
            f.write(private_key.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.PKCS8,
                encryption_algorithm=serialization.NoEncryption()
            ))

        subject = issuer = x509.Name([
            x509.NameAttribute(NameOID.COUNTRY_NAME, CERTIFICATE_INFO["country"]),
            x509.NameAttribute(NameOID.STATE_OR_PROVINCE_NAME, CERTIFICATE_INFO["state"]),
            x509.NameAttribute(NameOID.LOCALITY_NAME, CERTIFICATE_INFO["locality"]),
            x509.NameAttribute(NameOID.ORGANIZATION_NAME, CERTIFICATE_INFO["org"]),
            x509.NameAttribute(NameOID.COMMON_NAME, self.faculty_name),
        ])

        cert = (x509.CertificateBuilder()
            .subject_name(subject).issuer_name(issuer)
            .public_key(private_key.public_key())
            .serial_number(x509.random_serial_number())
            .not_valid_before(datetime.now(timezone.utc))
            .not_valid_after(datetime.now(timezone.utc) + timedelta(days=365 * 2)) # 2-year validity
            .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
            .sign(private_key, hashes.SHA256())
        )

        with open(self.cert_path, "wb") as f:
            f.write(cert.public_bytes(serialization.Encoding.PEM))

        print(f"✅ Saved key and certificate to '{self.private_key_path}' and '{self.cert_path}'")
        return private_key, cert

    def sign_pdf(self, input_pdf_path, output_pdf_path):
        """Signs a PDF and returns True on success."""
        if not self.available:
            return False
        try:
            with open(input_pdf_path, "rb") as inf:
                w = IncrementalPdfFileWriter(inf)
                pdf_signer = signers.PdfSigner(
                    signers.PdfSignatureMetadata(field_name='Signature', hash_algorithm='sha256'),
                    signer=self.signer,
                    stamp_style=stamp.TextStampStyle(
                        stamp_text=f"Digitally Signed by:\n{self.faculty_name}\nDate: %(ts)s"
                    )
                )
                fields.append_signature_field(w, sig_field_spec=SigFieldSpec('Signature', box=(20, 20, 220, 70)))
                with open(output_pdf_path, "wb") as outf:
                    pdf_signer.sign_pdf(w, output=outf)
            return True
        except Exception as e:
            print(f"An error occurred during PDF signing: {e}")
            return False

# ==============================================================================
# 3. WRITERS
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
    """Creates a DOCX, converts it to a PDF, and then digitally signs it."""
    def __init__(self, signer=None):
        self.signer = signer

    def write(self, doc, output_filename):
        if convert is None:
            print("\nError: A required module for PDF conversion is missing. Cannot generate PDF.")
            return
            
        if platform.system() == "Darwin":
            if not os.path.exists("/Applications/LibreOffice.app"):
                print("\n" + "="*60)
                print("CRITICAL ERROR: LibreOffice is not installed correctly.")
                print("The application must be in your main /Applications folder to work.")
                print("Please download and install LibreOffice from https://www.libreoffice.org/")
                print("="*60)
                return

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_docx = os.path.join(temp_dir, "temp_report.docx")
                unsigned_pdf = os.path.join(temp_dir, "unsigned_report.pdf")
                
                doc.save(temp_docx)
                convert(temp_docx, unsigned_pdf)

                if self.signer and self.signer.available:
                    print(f"Signing document with {self.signer.faculty_name}'s key...")
                    if self.signer.sign_pdf(unsigned_pdf, output_filename):
                        print(f"\nSuccess! Signed report generated as '{output_filename}' ✨")
                    else:
                        print("\nFailed to generate signed PDF.")
                else:
                    shutil.copy(unsigned_pdf, output_filename)
                    print(f"\nSuccess! Unsigned report generated as '{output_filename}' ✨")

        except Exception as e:
            print(f"\nError: Could not save or convert the file. Details: {e}")
            print("Please ensure Microsoft Word (on Windows) or LibreOffice (on macOS/Linux) is installed.")

# ==============================================================================
# 4. FORMATTERS
# ==============================================================================
class BaseFormatter:
    """Base class for all formatters with shared helper methods."""
    def get_year_semester_string(self, roman_numeral):
        return SEMESTER_MAPPING.get(str(roman_numeral).strip().lower(), str(roman_numeral))

    def set_cell_properties(self, cell, text, bold=False, font_size=10, align='LEFT', valign='TOP'):
        cell.text = ''
        p = cell.add_paragraph()
        run = p.add_run(str(text))
        run.font.size = Pt(font_size)
        run.bold = bold
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
        # Implementation for creating Format 1 content
        doc.add_heading('Format 1. Assessment of the learning levels of the students:', level=2).alignment = WD_ALIGN_PARAGRAPH.CENTER
        container_table = doc.add_table(rows=5, cols=1)
        container_table.style = 'Table Grid'

        header_cell = container_table.cell(0, 0)
        p1 = header_cell.add_paragraph(); p1.add_run('Manipal Institute of Technology').bold = True; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2 = header_cell.add_paragraph(); p2.add_run('MAHE Manipal').bold = True; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p3 = header_cell.add_paragraph(); p3.add_run('Computer Science and Engineering Department').bold = True; p3.alignment = WD_ALIGN_PARAGRAPH.CENTER

        student_info_table = container_table.cell(1, 0).add_table(rows=4, cols=2)
        self.set_cell_properties(student_info_table.cell(0, 0), 'Name of the Student:')
        self.set_cell_properties(student_info_table.cell(0, 1), str(student.get('Student Name', '')))
        self.set_cell_properties(student_info_table.cell(1, 0), 'Registration Number:')
        self.set_cell_properties(student_info_table.cell(1, 1), str(student.get('Register Number of the Student', '')))
        self.set_cell_properties(student_info_table.cell(2, 0), 'Course:')
        self.set_cell_properties(student_info_table.cell(2, 1), str(student.get('Subject Name', '')).title())
        self.set_cell_properties(student_info_table.cell(3, 0), 'Year /semester:')
        self.set_cell_properties(student_info_table.cell(3, 1), self.get_year_semester_string(student.get('Semester', '')))

        params_table = container_table.cell(2, 0).add_table(rows=3, cols=4)
        params_table.style = 'Table Grid'
        hdr_cell1 = params_table.cell(0, 2); hdr_cell2 = params_table.cell(0, 3); hdr_cell1.merge(hdr_cell2)
        self.set_cell_properties(params_table.cell(0, 0), 'Sr. No.', bold=True, align='CENTER')
        self.set_cell_properties(params_table.cell(0, 1), 'Parameter', bold=True, align='CENTER')
        self.set_cell_properties(params_table.cell(0, 2), 'Weightage in Percentage', bold=True, align='CENTER')
        self.set_cell_properties(params_table.cell(1, 0), '1', align='CENTER')
        self.set_cell_properties(params_table.cell(1, 1), f"Scores obtained by student class test / internal examination...\nConsidered Midterm exam conducted for {MIDTERM_TOTAL_MARKS}M:")
        self.set_cell_properties(params_table.cell(1, 2), f"{student.get('MidtermPercentage', 0):.2f}", align='CENTER')
        self.set_cell_properties(params_table.cell(1, 3), "> %", align='CENTER')
        self.set_cell_properties(params_table.cell(2, 0), '2', align='CENTER')
        self.set_cell_properties(params_table.cell(2, 1), 'Performance of students in preceding university examination')
        self.set_cell_properties(params_table.cell(2, 2), str(student.get('CGPA (up to previous semester)', 'N/A')), align='CENTER')
        self.set_cell_properties(params_table.cell(2, 3), "> %", align='CENTER')
        params_table.columns[0].width = Inches(0.5); params_table.columns[1].width = Inches(4.0); params_table.columns[2].width = Inches(1.0); params_table.columns[3].width = Inches(0.5)

        container_table.cell(3, 0).text = "Total Weightage"
        footer_cell = container_table.cell(4, 0)
        footer_cell.add_paragraph(f"1. Midterm score less than {slow_threshold}% considered as a slow learner")
        footer_cell.add_paragraph(f"2. Midterm score more than {fast_threshold}% considered as an advanced learner **")
        footer_cell.add_paragraph(f"Date: {datetime.now().strftime('%d-%m-%Y')}")
        self.add_signature_line(footer_cell)

    def _create_format2_content(self, doc, student):
        # Implementation for creating Format 2 content
        doc.add_heading('Format -2 Report of performance/ improvement for slow and advanced learners', level=2).alignment = WD_ALIGN_PARAGRAPH.CENTER
        header_table = doc.add_table(rows=3, cols=1)
        p1 = header_table.cell(0, 0).paragraphs[0]; p1.add_run('Manipal Institute of Technology').bold = True; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2 = header_table.cell(1, 0).paragraphs[0]; p2.add_run('MAHE Manipal').bold = True; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p3 = header_table.cell(2, 0).paragraphs[0]; p3.add_run('Computer Science and Engineering Department').bold = True; p3.alignment = WD_ALIGN_PARAGRAPH.CENTER

        content_table = doc.add_table(rows=8, cols=2)
        content_table.style = 'Table Grid'
        self.set_cell_properties(content_table.cell(0, 0), '1. Registration Number')
        self.set_cell_properties(content_table.cell(0, 1), student.get('Register Number of the Student', ''))
        self.set_cell_properties(content_table.cell(1, 0), '2. Name of the student')
        self.set_cell_properties(content_table.cell(1, 1), student.get('Student Name', ''))
        self.set_cell_properties(content_table.cell(2, 0), '3. Course')
        self.set_cell_properties(content_table.cell(2, 1), str(student.get('Subject Name', '')).title())
        self.set_cell_properties(content_table.cell(3, 0), '4. Year/Semester')
        self.set_cell_properties(content_table.cell(3, 1), self.get_year_semester_string(student.get('Semester', '')))
        self.set_cell_properties(content_table.cell(4, 0), '5. Midterm Percentage')
        self.set_cell_properties(content_table.cell(4, 1), f"{student.get('MidtermPercentage', 0):.2f}%")
        self.set_cell_properties(content_table.cell(5, 0), '6. Activities/ Measure/special programs\ntaken to improve the performance')
        self.set_cell_properties(content_table.cell(5, 1), str(student.get('Actions taken to improve performance', '')).replace(';', '\n'))
        self.set_cell_properties(content_table.cell(6, 0), '7. Progress')
        self.set_cell_properties(content_table.cell(6, 1), str(student.get('Outcome (Based on clearance in end-semester or makeup exam)', '')))
        self.set_cell_properties(content_table.cell(7, 0), 'Comments/remarks')
        self.set_cell_properties(content_table.cell(7, 1), str(student.get('Remarks if any', '')))

        doc.add_paragraph(f"\nDate:{datetime.now().strftime('%d-%m-%Y')}")
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
                row_cells[0].text = str(index + 1)
                row_cells[1].text = str(row_data.get('Register Number of the Student', ''))
                row_cells[2].text = str(row_data.get('Student Name', ''))
                row_cells[3].text = f"{row_data.get('MidtermPercentage', 0):.2f}"
                row_cells[4].text = str(row_data.get('Outcome (Based on clearance in end-semester or makeup exam)', ''))
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
def get_writer(output_type, faculty_name=None):
    """Factory function to get the appropriate file writer."""
    if output_type == 'word':
        return DocxWriter()
    elif output_type == 'pdf':
        if convert is None:
            raise ValueError("PDF output is not available. Required modules are missing.")
        
        if faculty_name and faculty_name.lower() != 'none':
            signer = PdfSigner(faculty_name)
            return PdfWriter(signer=signer)
        else:
            return PdfWriter(signer=None) 
            
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
        
        final_filtered.sort(key=lambda s: (s.get('Subject Name', ''), s.get('Semester',''), s.get('Register Number of the Student', '')))
        return final_filtered

# ==============================================================================
# 7. CONTROLLER
# ==============================================================================
class ReportController:
    """Controls the report generation workflow."""
    def __init__(self, excel_path, format_choice, learner_type, slow_thresh, fast_thresh, output_type, subject, semester, faculty_name):
        self.excel_path = excel_path
        self.format_choice = format_choice
        self.learner_type = learner_type
        self.slow_threshold = slow_thresh
        self.fast_threshold = fast_thresh
        self.output_type = output_type
        self.subject = subject.lower().strip()
        self.semester = semester.lower().strip()
        self.faculty_name = faculty_name.strip()
        self.reader = DataReader()
        self.processor = StudentDataProcessor()
        try:
            self.writer = get_writer(output_type, self.faculty_name)
        except ValueError as e:
            print(e)
            self.writer = None

    def _get_output_path(self, filename_base):
        date_str = datetime.now().strftime('%d_%m_%y')
        
        base_dir = "Learner Monitor Reports"
        learner_folder = f"{self.learner_type.title()} Learners"
        
        semester_name_for_file = "AllSemesters" if self.semester == 'all' else self.semester.upper()
        subject_name_for_file = "AllSubjects" if self.subject == 'all' else self.subject.replace(' ', '_').title()
        
        semester_folder = f"Semester_{semester_name_for_file}" if self.semester != 'all' else semester_name_for_file
        subject_folder = subject_name_for_file
        
        output_dir = os.path.join(base_dir, learner_folder, semester_folder, subject_folder)
        os.makedirs(output_dir, exist_ok=True)
        
        filename = f"{subject_name_for_file}_{semester_name_for_file}_{self.learner_type.title()}Learner_{filename_base}_{date_str}.{self.output_type}"
        return os.path.join(output_dir, filename)


    def run(self):
        if not self.writer:
            print("Report generation halted due to invalid writer configuration.")
            return

        all_student_data = self.reader.read_data(self.excel_path)
        if not all_student_data:
            return

        processed_students = self.processor.process_data(all_student_data)
        final_filtered_students = self.processor.filter_students(
            processed_students, self.semester, self.subject, self.learner_type,
            self.slow_threshold, self.fast_threshold
        )

        if not final_filtered_students:
            print(f"\nNo students found for the selected criteria (Subject: '{self.subject}', Semester: '{self.semester}').")
            return

        print(f"\nFound {len(final_filtered_students)} {self.learner_type} learners.")

        if self.format_choice == '5':
            self._generate_all_formats(final_filtered_students)
            return

        try:
            formatter = get_formatter(self.format_choice)
            output_object = formatter.format(final_filtered_students, self.slow_threshold, self.fast_threshold)

            report_name_map = {'1': 'Format1', '2': 'Format2', '3': 'Summary', '4': 'Combined'}
            report_name = report_name_map.get(self.format_choice, "Report")
            
            output_path = self._get_output_path(report_name)
            self.writer.write(output_object, output_path)

        except ValueError as e:
            print(e)

    def _generate_all_formats(self, students):
        """Handles the logic for generating reports for format choice '5'."""
        
        combined_path = self._get_output_path("Combined_Report")
        summary_path = self._get_output_path("Summary_Report")

        f1_and_2_formatter = Format1And2DocxFormatter()
        doc1 = f1_and_2_formatter.format(students, self.slow_threshold, self.fast_threshold)
        self.writer.write(doc1, combined_path)

        f3_formatter = Format3DocxFormatter()
        doc2 = f3_formatter.format(students, self.slow_threshold, self.fast_threshold)
        self.writer.write(doc2, summary_path)

# ==============================================================================
# 8. MAIN EXECUTION BLOCK
# ==============================================================================
def get_valid_input(prompt, valid_options=None, input_type=str):
    """Helper function to get and validate user input."""
    while True:
        user_input = input(prompt).strip()
        if not user_input:
            print("Input cannot be empty. Please try again.")
            continue
        try:
            converted_input = input_type(user_input)
            if valid_options and converted_input not in valid_options:
                print(f"Invalid input. Please choose from {valid_options}.")
                continue
            return converted_input
        except ValueError:
            print(f"Invalid input type. Please enter a valid {input_type.__name__}.")

if __name__ == "__main__":
    print("="*50)
    print(" Welcome to the Student Learner Report Generator ")
    print("="*50)
    
    excel_file = input("Enter the data file name (e.g., students.xlsx or data.csv): ")
    if not os.path.exists(excel_file):
        print(f"Error: The file '{excel_file}' does not exist.")
        exit()

    subject_filter = input("Enter Subject Name to filter by (or 'all'): ").strip()
    semester_filter = input("Enter Semester as a Roman numeral (e.g., 'III' or 'all'): ").strip()

    output_format = get_valid_input("Choose output format ('word' or 'pdf'): ", ['word', 'pdf'])
    
    faculty_name_input = 'none'
    if output_format == 'pdf':
        sign_choice = get_valid_input("Do you want to digitally sign this PDF? (yes/no): ", ['yes', 'no'])
        if sign_choice == 'yes':
            faculty_name_input = input("Enter the full name of the signing faculty (e.g., 'Dr. Jane Doe'): ").strip()
        
    learner = get_valid_input("Generate report for 'fast' or 'slow' learners? ", ['fast', 'slow'])
    slow_thresh = get_valid_input("Enter percentage threshold for SLOW learners (e.g., 40): ", input_type=float)
    fast_thresh = get_valid_input("Enter percentage threshold for FAST learners (e.g., 80): ", input_type=float)

    print("\nPlease choose a report format:")
    print(" 1: Format 1 - Assessment of learning levels")
    print(" 2: Format 2 - Report of performance/improvement")
    print(" 3: Format 3 - Tabular Summary Report")
    print(" 4: Combined Format 1 & 2")
    print(" 5: All Formats (Generates 2 separate files)")

    format_num = get_valid_input("Enter your choice (1-5): ", ['1', '2', '3', '4', '5'])

    controller = ReportController(
        excel_path=excel_file, 
        format_choice=format_num, 
        learner_type=learner, 
        slow_thresh=slow_thresh, 
        fast_thresh=fast_thresh, 
        output_type=output_format, 
        subject=subject_filter, 
        semester=semester_filter,
        faculty_name=faculty_name_input
    )
    controller.run()
    print("\nReport generation complete")
