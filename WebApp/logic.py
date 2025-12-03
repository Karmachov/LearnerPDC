# ==============================================================================
# REPORT GENERATION LOGIC LIBRARY
# ==============================================================================

import os
import tempfile
import platform
import time
import re
import traceback
import warnings
from datetime import datetime

import pandas as pd
import openpyxl

import fitz 
from docx2pdf import convert
from endesive import pdf
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL

warnings.filterwarnings("ignore", category=DeprecationWarning)

# --- CONFIGURATION ---
MIDTERM_TOTAL_MARKS = 30
SEMESTER_MAPPING = {
    'i': 'I Year/ I semester', 'ii': 'I Year/ II semester', 'iii': 'II Year/ III semester',
    'iv': 'II Year/ IV semester', 'v': 'III Year/ V semester', 'vi': 'III Year/ VI semester',
    'vii': 'IV Year/ VII semester', 'viii': 'IV Year/ VIII semester',
}



# --- PDF SIGNING UTILITY ---
def sign_pdf(pdf_path, key_path, cert_path, image_path, password):
    try:
        date = datetime.now().strftime('D:%Y%m%d%H%M%S+05\'30\'')

        with open(key_path, 'rb') as f: 
            private_key = load_pem_private_key(f.read(), password=password.encode('utf-8') if password else None)
        with open(cert_path, 'rb') as f: 
            certificate = load_pem_x509_certificate(f.read())
        with open(pdf_path, 'rb') as f_in: 
            pdf_data = f_in.read()

        # Invisible Crypto Signature
        signdata = {
            'sigflags': 3,
            'contact': 'faculty@example.com',
            'location': 'Manipal',
            'reason': 'Verified Report',
            'signingdate': date,
            'page': 0
        }

        signed_pdf_bytes = pdf.cms.sign(pdf_data, signdata, key=private_key, cert=certificate, othercerts=())

        with open(pdf_path, 'wb') as f_out: 
            f_out.write(pdf_data + signed_pdf_bytes)
            
        print("PDF digitally signed (Invisible).")
        return True
    except Exception:
        print("Failed to sign PDF.")
        traceback.print_exc()
        return False

# --- DATA READER ---
class DataReader:
    COLUMN_MAPPING = {
        'Roll Number': 'Register Number of the Student', 'Student Name': 'Student Name',
        'Total (30) *': 'Midterm Exam Marks (Out of 30)', 'Student Viewed': 'Did student view the paper'
    }

    def _extract_subject_from_header(self, file_path):
        try:
            workbook = openpyxl.load_workbook(file_path, read_only=True)
            sheet = workbook.active
            for row in range(1, 6):
                cell_value = sheet.cell(row=row, column=1).value
                if cell_value and isinstance(cell_value, str) and "Exam:" in cell_value:
                    last_slash_index = cell_value.rfind('/')
                    first_bracket_index = cell_value.find('[')
                    if last_slash_index != -1 and first_bracket_index != -1 and last_slash_index < first_bracket_index:
                        return cell_value[last_slash_index + 1 : first_bracket_index].strip()
            return None
        except Exception:
            return None

    def read_data(self, file_path):
        subject_name = self._extract_subject_from_header(file_path)
        if not subject_name: raise ValueError("Could not auto-detect subject name.")

        try:
            df = pd.read_excel(file_path, skiprows=2)
            df.columns = df.columns.str.strip()
            df.rename(columns=self.COLUMN_MAPPING, inplace=True)
            reg_col = 'Register Number of the Student'
            if reg_col in df.columns:
                df[reg_col] = df[reg_col].astype(str).str.replace(r'\.0$', '', regex=True)
            return df.to_dict('records'), subject_name
        except Exception:
            traceback.print_exc()
            raise

    def read_cgpa_map(self, file_path):
        if not file_path or not os.path.exists(file_path):
            return {}
        try:
            print(f"Reading CGPA file: {os.path.basename(file_path)}")
            is_csv = file_path.lower().endswith('.csv')
            if is_csv: df = pd.read_csv(file_path)
            else: df = pd.read_excel(file_path)
            
            def find_cols(dataframe):
                cols = dataframe.columns.astype(str).str.strip().str.lower()
                r = next((c for c in dataframe.columns if any(x in str(c).lower() for x in ['roll', 'reg', 'usn'])), None)
                c = next((c for c in dataframe.columns if any(x in str(c).lower() for x in ['cgpa', 'sgpa', 'gpa', 'mark'])), None)
                return r, c

            roll_col, cgpa_col = find_cols(df)
            if not roll_col or not cgpa_col:
                print("Columns not found in first row, trying header=1...")
                if is_csv: df = pd.read_csv(file_path, header=1)
                else: df = pd.read_excel(file_path, header=1)
                roll_col, cgpa_col = find_cols(df)

            if not roll_col or not cgpa_col:
                print(f"Warning: Columns not found in {file_path}")
                return {}

            df[roll_col] = df[roll_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            return pd.Series(df[cgpa_col].values, index=df[roll_col]).to_dict()
        except Exception as e:
            print(f"Error reading CGPA file: {e}")
            return {}

# --- DATA PROCESSOR ---
class StudentDataProcessor:
    def _calculate_midterm_percentage(self, marks):
        try: return (float(marks) / MIDTERM_TOTAL_MARKS) * 100
        except (ValueError, TypeError): return 0

    def process_data(self, all_student_data, subject_name, semester, common_comment, cgpa_map=None):
        cgpa_map = cgpa_map or {}
        for student in all_student_data:
            student['MidtermPercentage'] = self._calculate_midterm_percentage(student.get('Midterm Exam Marks (Out of 30)'))
            student['Subject Name'] = str(subject_name).strip().lower()
            student['Semester'] = str(semester).strip().lower()
            roll_no = str(student.get('Register Number of the Student', '')).strip()
            student['CGPA (up to previous semester)'] = cgpa_map.get(roll_no, '') 
            student['Actions taken to improve performance'] = common_comment
        return all_student_data

    def filter_students(self, students, learner_type, slow_thresh, fast_thresh):
        if learner_type == 'slow':
            final_filtered = [s for s in students if s['MidtermPercentage'] <= slow_thresh]
        else:
            final_filtered = [s for s in students if s['MidtermPercentage'] >= fast_thresh]
        final_filtered.sort(key=lambda s: s.get('Register Number of the Student', ''))
        return final_filtered

# --- REPORT FORMATTERS ---
class BaseFormatter:
    BODY_FONT = "Times New Roman"
    
    def __init__(self):
        self.signature_image_path = None

    def get_year_semester_string(self, roman_numeral): 
        return SEMESTER_MAPPING.get(str(roman_numeral).strip().lower(), str(roman_numeral))
    
    def set_cell_properties(self, cell, text, bold=False, font_size=10, align='LEFT', valign='TOP', font_name=None):
        cell.text = ''
        p = cell.add_paragraph()
        run = p.add_run(str(text))
        run.font.size = Pt(font_size)
        run.bold = bold
        if font_name: run.font.name = font_name
        p.alignment = getattr(WD_ALIGN_PARAGRAPH, str(align).upper(), WD_ALIGN_PARAGRAPH.LEFT)
        cell.vertical_alignment = getattr(WD_ALIGN_VERTICAL, str(valign).upper(), WD_ALIGN_VERTICAL.TOP)
    
    def add_signature_line(self, doc_or_cell, print_faculty=True):
        """
        Adds the signature (image + underline + label).
        If print_faculty is True and self.faculty_name exists, prints the faculty name
        below the signature. For formats that need custom placement, pass print_faculty=False
        and handle the faculty name separately.
        """
        p = doc_or_cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Try to insert signature image if provided
        if getattr(self, "signature_image_path", None) and os.path.exists(self.signature_image_path):
            try:
                run = p.add_run()
                run.add_picture(self.signature_image_path, width=Inches(0.8))
                run.add_break()
            except Exception as e:
                print(f"Error inserting signature image: {e}")

    # Add signature underline and label
        p.add_run("\n" + "_" * 40 + "\n")
        label_run = p.add_run("Signature of the\nsubject teacher / class coordinator")
        try:
            label_run.font.name = self.BODY_FONT
            label_run.font.size = Pt(10)
        except Exception:
            pass

    # If requested, print the faculty name (right-aligned) below the signature
        if print_faculty:
            faculty = getattr(self, "faculty_name", None)
            if faculty:
                fn_para = doc_or_cell.add_paragraph()
                fn_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                fn_run = fn_para.add_run(str(faculty).strip())
                try:
                    fn_run.font.name = self.BODY_FONT
                    fn_run.font.size = Pt(10)
                    fn_run.font.bold = True
                except Exception:
                    pass


    
    def _add_document_header(self, cell):
        p1 = cell.add_paragraph(); p1.add_run('Manipal Institute of Technology').bold = True; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2 = cell.add_paragraph(); p2.add_run('MAHE Manipal').bold = True; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p3 = cell.add_paragraph(); p3.add_run('Computer Science and Engineering Department').bold = True; p3.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def _create_format1_content(self, doc, student, slow_threshold, fast_threshold):
        doc.add_heading('Format 1. Assessment of the learning levels of the students:', level=2).alignment = WD_ALIGN_PARAGRAPH.CENTER
        ct = doc.add_table(rows=5, cols=1); ct.style = 'Table Grid'; self._add_document_header(ct.cell(0, 0))
        st = ct.cell(1, 0).add_table(rows=4, cols=2)
        self.set_cell_properties(st.cell(0, 0), 'Name of the Student:'); self.set_cell_properties(st.cell(0, 1), student.get('Student Name', ''), font_name=self.BODY_FONT)
        self.set_cell_properties(st.cell(1, 0), 'Registration Number:'); self.set_cell_properties(st.cell(1, 1), student.get('Register Number of the Student', ''), font_name=self.BODY_FONT)
        self.set_cell_properties(st.cell(2, 0), 'Course:'); self.set_cell_properties(st.cell(2, 1), str(student.get('Subject Name', '')).title(), font_name=self.BODY_FONT)
        self.set_cell_properties(st.cell(3, 0), 'Year /semester:'); self.set_cell_properties(st.cell(3, 1), self.get_year_semester_string(student.get('Semester', '')), font_name=self.BODY_FONT)
        pt = ct.cell(2, 0).add_table(rows=3, cols=4); pt.style = 'Table Grid'; pt.cell(0, 2).merge(pt.cell(0, 3))
        self.set_cell_properties(pt.cell(0, 0), 'Sr. No.', bold=True, align='CENTER'); self.set_cell_properties(pt.cell(0, 1), 'Parameter', bold=True, align='CENTER'); self.set_cell_properties(pt.cell(0, 2), 'Weightage in Percentage', bold=True, align='CENTER')
        self.set_cell_properties(pt.cell(1, 0), '1', align='CENTER'); self.set_cell_properties(pt.cell(1, 1), f"Scores obtained by student class test / internal examination...\nConsidered Midterm exam conducted for {MIDTERM_TOTAL_MARKS}M:"); self.set_cell_properties(pt.cell(1, 2), f"{student.get('MidtermPercentage', 0):.2f}", align='CENTER', font_name=self.BODY_FONT); self.set_cell_properties(pt.cell(1, 3), "> %", align='CENTER')
        self.set_cell_properties(pt.cell(2, 0), '2', align='CENTER'); self.set_cell_properties(pt.cell(2, 1), 'Performance of students in preceding university examination'); self.set_cell_properties(pt.cell(2, 2), str(student.get('CGPA (up to previous semester)', '')), align='CENTER', font_name=self.BODY_FONT); self.set_cell_properties(pt.cell(2, 3), "> %", align='CENTER')
        ct.cell(3, 0).text = "Total Weightage"; fc = ct.cell(4, 0); fc.add_paragraph(f"1. Midterm score less than {slow_threshold}% considered as a slow learner"); fc.add_paragraph(f"2. Midterm score more than {fast_threshold}% considered as an advanced learner **")
        pd_ = fc.add_paragraph(); pd_.add_run(f"Date: {datetime.now().strftime('%d-%m-%Y')}").font.name = self.BODY_FONT; self.add_signature_line(fc)

    def _create_format2_content(self, doc, student, slow_threshold, fast_threshold):
        # --- Dynamic Heading Based on Learner Type (slow vs advanced) ---
# use a paragraph with Heading style so Word renders it like a heading
        heading = doc.add_paragraph(style='Heading 2')
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

        r1 = heading.add_run("Format -2 Report of performance/ improvement for ")
        r2 = heading.add_run("slow")
        r3 = heading.add_run(" and ")
        r4 = heading.add_run("advanced")
        r5 = heading.add_run(" learners")

# make sure font properties are explicitly set (helps ensure character formatting will show)
        for r in (r1, r2, r3, r4, r5):
            r.font.name = self.BODY_FONT
            r.font.size = Pt(14)

# Ensure thresholds and mid are numeric and compute membership
        try:
            mid = float(student.get("MidtermPercentage", 0))
        except Exception:
            mid = 0.0

        try:
            slow_num = float(slow_threshold)
            fast_num = float(fast_threshold)
        except Exception:
            slow_num = slow_threshold
            fast_num = fast_threshold

        is_slow = mid <= slow_num

# Set underline on the selected run (and explicitly disable underline on the other)
        if is_slow:
            r2.font.underline = True   # underline "slow"
            r4.font.underline = False
        else:
            r4.font.underline = True   # underline "advanced"
            r2.font.underline = False



        ht = doc.add_table(rows=1, cols=1); self._add_document_header(ht.cell(0,0))
        ct = doc.add_table(rows=8, cols=2); ct.style = 'Table Grid'
        self.set_cell_properties(ct.cell(0, 0), '1. Registration Number'); self.set_cell_properties(ct.cell(0, 1), student.get('Register Number of the Student', ''), font_name=self.BODY_FONT)
        self.set_cell_properties(ct.cell(1, 0), '2. Name of the student'); self.set_cell_properties(ct.cell(1, 1), student.get('Student Name', ''), font_name=self.BODY_FONT)
        self.set_cell_properties(ct.cell(2, 0), '3. Course'); self.set_cell_properties(ct.cell(2, 1), str(student.get('Subject Name', '')).title(), font_name=self.BODY_FONT)
        self.set_cell_properties(ct.cell(3, 0), '4. Year/Semester'); self.set_cell_properties(ct.cell(3, 1), self.get_year_semester_string(student.get('Semester', '')), font_name=self.BODY_FONT)
        self.set_cell_properties(ct.cell(4, 0), '5. Midterm Percentage'); self.set_cell_properties(ct.cell(4, 1), f"{student.get('MidtermPercentage', 0):.2f}%", font_name=self.BODY_FONT)
        self.set_cell_properties(ct.cell(5, 0), '6. Activities/ Measure/special programs\ntaken to improve the performance'); self.set_cell_properties(ct.cell(5, 1), str(student.get('Actions taken to improve performance', '')).replace(';', '\n'), font_name=self.BODY_FONT)
        self.set_cell_properties(ct.cell(6, 0), '7. Progress'); self.set_cell_properties(ct.cell(6, 1), str(student.get('Outcome (Based on clearance in end-semester or makeup exam)', '')), font_name=self.BODY_FONT)
        self.set_cell_properties(ct.cell(7, 0), 'Comments/remarks'); self.set_cell_properties(ct.cell(7, 1), str(student.get('Remarks if any', '')), font_name=self.BODY_FONT)
        pd_ = doc.add_paragraph(); pd_.add_run(f"\nDate:{datetime.now().strftime('%d-%m-%Y')}").font.name = self.BODY_FONT; self.add_signature_line(doc)

    def _generate_pages(self, doc, students, content_method, *args):
        for i, student in enumerate(students): content_method(doc, student, *args); (i < len(students) - 1) and doc.add_page_break()
        return doc

class Format1DocxFormatter(BaseFormatter):
    def format(self, s, st, ft): doc = Document(); return self._generate_pages(doc, s, self._create_format1_content, st, ft)
class Format2DocxFormatter(BaseFormatter):
    def format(self, s, st, ft): doc = Document(); return self._generate_pages(doc, s, self._create_format2_content, st, ft)


# --- FORMAT 3: TABLE + DATE + SIGNATURE (NO CRITERIA TEXT) ---
class Format3DocxFormatter(BaseFormatter):
    def format(self, students, st, ft):
        doc = Document(); df = pd.DataFrame(students); grouped = df.groupby(['Subject Name', 'Semester'])
        for i, ((subject, semester), group) in enumerate(grouped):
            doc.add_paragraph(f"Course: {str(subject).title()}", style='Heading 3'); doc.add_paragraph(f"Year /Semester: {self.get_year_semester_string(semester)}", style='Heading 3')
            sc = ['Sl. No', 'Reg Number', 'Name of the student', 'Midterm Percentage', 'Progress']; t = doc.add_table(rows=1, cols=len(sc)); t.style = 'Table Grid'
            for j, col_name in enumerate(sc): self.set_cell_properties(t.cell(0, j), col_name, bold=True)
            for index, row_data in group.reset_index(drop=True).iterrows():
                rc = t.add_row().cells
                self.set_cell_properties(rc[0], str(index + 1), font_name=self.BODY_FONT); self.set_cell_properties(rc[1], str(row_data.get('Register Number of the Student', '')), font_name=self.BODY_FONT)
                self.set_cell_properties(rc[2], str(row_data.get('Student Name', '')), font_name=self.BODY_FONT); self.set_cell_properties(rc[3], f"{row_data.get('MidtermPercentage', 0):.2f}", font_name=self.BODY_FONT)
                self.set_cell_properties(rc[4], str(row_data.get('Outcome (Based on clearance in end-semester or makeup exam)', '')), font_name=self.BODY_FONT)
            
            # --- FOOTER: Date & Signature ONLY ---
            # --- FOOTER: Date & Signature ONLY ---
            p = doc.add_paragraph()
            p.add_run(f"\nDate: {datetime.now().strftime('%d-%m-%Y')}").font.name = self.BODY_FONT
            self.add_signature_line(doc,print_faculty=False)

# --- ensure faculty name is printed right-aligned under the signature for Format 3 ---
            faculty = getattr(self, "faculty_name", None)
            if faculty:
                fn_para = doc.add_paragraph()
                fn_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                fn_run = fn_para.add_run(str(faculty).strip())
                try:
                    fn_run.font.name = self.BODY_FONT
                    fn_run.font.size = Pt(10)
                    fn_run.font.bold = True
                except Exception:
                    pass

            if i < len(grouped) - 1: doc.add_page_break()

        return doc

class Format1And2DocxFormatter(BaseFormatter):
    def format(self, students, st, ft):
        doc = Document()
        for i, student in enumerate(students):
            # Format 1 page
            self._create_format1_content(doc, student, st, ft)
            doc.add_page_break()

            # Format 2 page (with thresholds)
            self._create_format2_content(doc, student, st, ft)

            # Page break between different students
            if i < len(students) - 1:
                doc.add_page_break()

        return doc


# --- FILE WRITERS ---
class DocxWriter:
    def write(self, doc, output_filename, **kwargs):
        doc.save(output_filename)
        print(f"Report generated: {output_filename}")

class PdfWriter:
    def write(self, doc, output_filename, sign_info=None, format_choice=None):
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_docx = os.path.join(temp_dir, "temp.docx")
                doc.save(temp_docx)
                convert(temp_docx, output_filename)
            
            time.sleep(1)
            if platform.system() == "Darwin": time.sleep(1)
            print(f"PDF generated: {output_filename}")

            if sign_info and sign_info.get('should_sign') and format_choice in ['1', '2', '4', '5']:
                sign_pdf(
                    pdf_path=output_filename, 
                    key_path=sign_info['key_path'], 
                    cert_path=sign_info['cert_path'], 
                    image_path=sign_info['image_path'], 
                    password=sign_info['password']
                )

        except Exception as e:
            print(f"Error in PDF generation: {e}")
            traceback.print_exc()

# --- FACTORIES ---
def get_writer(output_type):
    if output_type == 'word': return DocxWriter()
    if output_type == 'pdf': return PdfWriter()
    raise ValueError(f"Unsupported output type: {output_type}")

def get_formatter(format_choice):
    formatters = {
        '1': Format1DocxFormatter(), '2': Format2DocxFormatter(),
        '3': Format3DocxFormatter(), '4': Format1And2DocxFormatter()
    }
    formatter = formatters.get(format_choice)
    if not formatter: raise ValueError(f"Invalid format: {format_choice}")
    return formatter

# --- CONTROLLER ---
class ReportController:
    def __init__(self, excel_path, cgpa_path, format_choice, learner_type, slow_thresh, fast_thresh, output_type, semester, sign_info, common_comment, faculty_name=None):
        self.excel_path = excel_path
        self.cgpa_path = cgpa_path
        self.format_choice = format_choice; self.learner_type = learner_type
        self.slow_threshold = slow_thresh; self.fast_threshold = fast_thresh; self.output_type = output_type
        self.subject = ""; self.semester = semester.lower().strip(); self.sign_info = sign_info
        self.faculty_name= faculty_name
        self.common_comment = common_comment; self.reader = DataReader(); self.processor = StudentDataProcessor()
        self.writer = get_writer(output_type)
        
    def run(self):
        all_student_data, detected_subject = self.reader.read_data(self.excel_path)
        if not all_student_data: return None
        
        cgpa_map = self.reader.read_cgpa_map(self.cgpa_path)
        self.subject = detected_subject.lower().strip()
        processed_students = self.processor.process_data(all_student_data, self.subject, self.semester, self.common_comment, cgpa_map=cgpa_map)
        final_filtered_students = self.processor.filter_students(processed_students, self.learner_type, self.slow_threshold, self.fast_threshold)
        
        if not final_filtered_students:
            print("No students found.")
            return None
        
        print(f"Found {len(final_filtered_students)} {self.learner_type} learners.")
        
        date_str = datetime.now().strftime('%d_%m_%y'); base_dir = "Learner_Monitor_Reports"
        learner_folder = f"{self.learner_type.title()}_Learners"; sem_name = self.semester.upper()
        subj_name = re.sub(r'[\\/*?:"<>|]', "", self.subject.replace(' ', '_').title())
        output_dir = os.path.join(base_dir, learner_folder, f"Semester_{sem_name}", subj_name)
        os.makedirs(output_dir, exist_ok=True)
        
        if self.format_choice == '5':
            return self._generate_all_formats(final_filtered_students, output_dir, date_str, sem_name, subj_name)
        
        formatter = get_formatter(self.format_choice)
        
        if self.sign_info and self.sign_info.get('image_path'):
            formatter.signature_image_path = self.sign_info.get('image_path')

        formatter.faculty_name = self.faculty_name if hasattr(self, 'faculty_name') else None

        output_object = formatter.format(final_filtered_students, self.slow_threshold, self.fast_threshold)
        ext = 'docx' if self.output_type == 'word' else 'pdf'
        report_name = {'1': 'Format1', '2': 'Format2', '3': 'Summary', '4': 'Combined'}.get(self.format_choice, "Report")
        output_filename = f'{subj_name}_{sem_name}_{self.learner_type.title()}Learner_{report_name}_{date_str}.{ext}'
        full_output_path = os.path.join(output_dir, output_filename)
        
        self.writer.write(output_object, full_output_path, sign_info=self.sign_info, format_choice=self.format_choice)
        return full_output_path
    
    def _generate_all_formats(self, students, output_dir, date_str, sem_name, subj_name):
        ext = 'docx' if self.output_type == 'word' else 'pdf'
        
        sig_path = None
        if self.sign_info and self.sign_info.get('image_path'):
            sig_path = self.sign_info.get('image_path')

        comb_fname = f'{subj_name}_{sem_name}_{self.learner_type.title()}Learner_Combined_Report_{date_str}.{ext}'
        f1_2_formatter = Format1And2DocxFormatter()
        f1_2_formatter.signature_image_path = sig_path
        f1_2_formatter.faculty_name = self.faculty_name
        doc1 = f1_2_formatter.format(students, self.slow_threshold, self.fast_threshold)
        full_path1 = os.path.join(output_dir, comb_fname)
        self.writer.write(doc1, full_path1, sign_info=self.sign_info, format_choice='4')
        
        summ_fname = f'{subj_name}_{sem_name}_{self.learner_type.title()}Learner_Summary_Report_{date_str}.{ext}'
        f3_formatter = Format3DocxFormatter()
        f3_formatter.signature_image_path = sig_path
        f3_formatter.faculty_name = self.faculty_name
        doc2 = f3_formatter.format(students, self.slow_threshold, self.fast_threshold)
        full_path2 = os.path.join(output_dir, summ_fname)
        self.writer.write(doc2, full_path2, sign_info=self.sign_info, format_choice='3')

        return [full_path1, full_path2]