# ==============================================================================
# logic.py — Report Generation Logic Library (Worker-side)
# Ported from WebApp/logic.py with the following changes:
#   1. sign_pdf() now accepts raw bytes instead of file paths (in-memory signing)
#   2. PdfWriter always uses LibreOffice subprocess (no docx2pdf dependency)
#   3. add_image_to_all_pages_fitz() added back for post-sign image overlay
# ==============================================================================

# Version: 2.0.1 - Fixed digital signing with empty image bytes
import os
import re
import io
import tempfile
import platform
import subprocess
import traceback
import warnings
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd
from endesive import pdf as endesive_pdf
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate

from docx import Document
from docx.shared import Pt, Inches, RGBColor
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


def get_libreoffice_command() -> str:
    system = platform.system()
    if system == 'Darwin':
        return '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    elif system == 'Windows':
        return r'C:\Program Files\LibreOffice\program\soffice.exe'
    else:
        return 'libreoffice'  # Linux (Docker worker)


# ==============================================================================
# PDF SIGNING — in-memory version
# Accepts key_bytes / cert_bytes / image_bytes directly, never touches disk for secrets.
# ==============================================================================

def add_image_to_all_pages_fitz(pdf_path: str, image_bytes: bytes, x=435, y=72, width=100, height=40):
    """Overlay signature image on every page except the first (which has the digital sig widget)."""
    if not image_bytes:
        return
    doc = fitz.open(pdf_path)
    for page_index in range(1, len(doc)):
        page = doc[page_index]
        rect = fitz.Rect(x, y, x + width, y + height)
        page.insert_image(rect, stream=image_bytes)
    doc.saveIncr()


def sign_pdf(pdf_path: str, key_bytes: bytes, cert_bytes: bytes, image_bytes: bytes, password: str) -> None:
    """
    Digitally sign the PDF on the FIRST page.
    All secret material is passed as bytes — never written to disk.

    Raises RuntimeError with a user-safe message on failure (propagates to Celery FAILURE).
    """
    password = (password or "").strip()
    print(f"DEBUG: sign_pdf called. image_bytes type: {type(image_bytes)}, length: {len(image_bytes) if image_bytes else 0}")
    try:
        date = datetime.now().strftime("D:%Y%m%d%H%M%S+05'30'")
        private_key = load_pem_private_key(
            key_bytes,
            password=password.encode("utf-8") if password else None,
        )
        certificate = load_pem_x509_certificate(cert_bytes)

        with open(pdf_path, 'rb') as f:
            pdf_data = f.read()

        # Robust check: only include signature_img if it looks like a valid PNG/JPEG
        is_valid_image = False
        img_obj = None
        if image_bytes and isinstance(image_bytes, bytes) and len(image_bytes) > 0:
            try:
                from PIL import Image
                img_obj = Image.open(io.BytesIO(image_bytes))
                img_obj.verify()  # Check if it's a valid image
                # Re-open because verify() closes the file or makes it unusable for further processing
                img_obj = Image.open(io.BytesIO(image_bytes))
                is_valid_image = True
            except Exception as e:
                print(f"DEBUG: Image validation failed: {e}")
                is_valid_image = False

        signdata = {
            'sigflags': 1,  # Invisible signature to avoid visual clutter/spam
            'contact': 'faculty@manipal.edu',
            'location': 'Manipal, India',
            'reason': 'Verified Learner Report',
            'signingdate': date,
            'page': 0,
        }
        # We no longer add signature_img or signaturebox here because the 
        # BaseFormatter already adds the signature image above the signature lines
        # in the Word document, which is more accurate.

        try:
            signed_bytes = endesive_pdf.cms.sign(
                pdf_data, signdata, key=private_key, cert=certificate, othercerts=()
            )
        except Exception as e:
            # If signing fails, we propagate the error
            raise

        with open(pdf_path, 'wb') as f:
            f.write(pdf_data + signed_bytes)

        # We no longer call add_image_to_all_pages_fitz here to avoid "signature spam"
        # on every page. The signatures in the document body are sufficient.

    except ValueError as exc:
        traceback.print_exc()
        msg = str(exc).lower()
        if "password" in msg or "decrypt" in msg:
            raise RuntimeError(
                "Incorrect key passphrase. Open Profile Settings, re-upload your private key "
                ".pem and certificate, and enter the passphrase that unlocks the key file."
            ) from exc
        raise RuntimeError(f"Digital signing failed: {exc}") from exc
    except ModuleNotFoundError as exc:
        traceback.print_exc()
        raise RuntimeError(
            f"Digital signing is misconfigured on the server (missing dependency: {exc.name}). "
            "Contact the administrator or retry after the worker image is updated."
        ) from exc
    except Exception as exc:
        traceback.print_exc()
        raise RuntimeError(
            "Digital signing failed. Verify your key, certificate, and passphrase in Profile Settings."
        ) from exc


# ==============================================================================
# Normalization helper
# ==============================================================================

def normalize_registration_number(reg_num) -> str:
    if pd.isna(reg_num) or reg_num is None:
        return ''
    normalized = str(reg_num).strip().replace('.0', '').replace(' ', '').upper()
    return normalized


# ==============================================================================
# DATA READER
# ==============================================================================

class DataReader:
    COLUMN_MAPPING = {
        'Roll Number': 'Register Number of the Student',
        'Student Name': 'Student Name',
        'Total (30) *': 'Midterm Exam Marks (Out of 30)',
        'Student Viewed': 'Did student view the paper',
    }

    def _extract_subject_from_header(self, file_path: str):
        try:
            engine = 'xlrd' if file_path.lower().endswith('.xls') else 'openpyxl'
            df_header = pd.read_excel(file_path, engine=engine, nrows=5, header=None)
            for val in df_header.iloc[:, 0]:
                if val and isinstance(val, str) and "Exam:" in val:
                    last_slash = val.rfind('/')
                    first_bracket = val.find('[')
                    last_bracket = val.find(']')
                    if last_slash != -1 and first_bracket != -1:
                        name = val[last_slash + 1: first_bracket].strip()
                        code = val[first_bracket + 1: last_bracket].strip() if last_bracket != -1 else ""
                        return f"{name} ({code})" if code else name
            return None
        except Exception:
            return None

    def read_data(self, file_path: str):
        subject_name = self._extract_subject_from_header(file_path)
        if not subject_name:
            raise ValueError("Could not auto-detect subject name from the Excel file header.")
        engine = 'xlrd' if file_path.lower().endswith('.xls') else 'openpyxl'
        df = pd.read_excel(file_path, skiprows=2, engine=engine)
        df.columns = df.columns.str.strip()
        df.rename(columns=self.COLUMN_MAPPING, inplace=True)
        reg_col = 'Register Number of the Student'
        if reg_col in df.columns:
            df[reg_col] = df[reg_col].apply(normalize_registration_number)
        return df.to_dict('records'), subject_name

    def read_cgpa_map(self, file_path: str) -> dict:
        if not file_path or not os.path.exists(file_path):
            return {}
        try:
            engine = 'xlrd' if file_path.lower().endswith('.xls') else 'openpyxl'
            if file_path.lower().endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                xl = pd.ExcelFile(file_path)
                best = xl.sheet_names[0]
                if len(xl.sheet_names) > 1:
                    max_rows = 0
                    for sheet in xl.sheet_names:
                        tmp = pd.read_excel(file_path, sheet_name=sheet, engine=engine)
                        if len(tmp) > max_rows:
                            max_rows = len(tmp); best = sheet
                df = pd.read_excel(file_path, sheet_name=best, engine=engine)
            if len(df.columns) >= 2:
                roll_col, cgpa_col = df.columns[0], df.columns[1]
                df[roll_col] = df[roll_col].apply(normalize_registration_number)
                return pd.Series(df[cgpa_col].values, index=df[roll_col]).to_dict()
            return {}
        except Exception:
            return {}

    def read_grade_map(self, file_path: str, course_code: str = None) -> dict:
        if not file_path or not os.path.exists(file_path):
            return {}
        try:
            engine = 'xlrd' if file_path.lower().endswith('.xls') else 'openpyxl'
            df = pd.read_csv(file_path) if file_path.lower().endswith('.csv') else pd.read_excel(file_path, engine=engine)
            if len(df.columns) >= 3:
                enroll_col, course_col, grade_col = df.columns[0], df.columns[1], df.columns[2]
                if course_code:
                    target = course_code.split('(')[-1].replace(')', '').strip() if '(' in course_code else course_code
                    df = df[df[course_col].astype(str).str.contains(target, case=False, na=False)]
                df[enroll_col] = df[enroll_col].apply(normalize_registration_number)
                return pd.Series(df[grade_col].values, index=df[enroll_col]).to_dict()
            return {}
        except Exception:
            return {}


# ==============================================================================
# DATA PROCESSOR
# ==============================================================================

class StudentDataProcessor:
    def _calc_pct(self, marks) -> float:
        try:
            return (float(marks) / MIDTERM_TOTAL_MARKS) * 100
        except Exception:
            return 0.0

    def process_data(self, all_data, subject_name, semester, proofs_text, common_comment, cgpa_map=None, grade_map=None):
        cgpa_map = cgpa_map or {}
        grade_map = grade_map or {}
        for student in all_data:
            student['MidtermPercentage'] = self._calc_pct(student.get('Midterm Exam Marks (Out of 30)'))
            student['Subject Name'] = str(subject_name).strip()
            student['Semester'] = str(semester).strip().lower()
            roll = normalize_registration_number(student.get('Register Number of the Student', ''))
            student['CGPA (up to previous semester)'] = cgpa_map.get(roll, '')
            student['Actions taken to improve performance'] = proofs_text
            student['Remarks if any'] = common_comment
            grade = str(grade_map.get(roll, '')).strip().upper()
            if not grade or grade in ['NAN', 'NONE']:
                student['Outcome (Based on clearance in end-semester or makeup exam)'] = ''
            elif grade in {'A+', 'A', 'B', 'C', 'D', 'E', 'S'}:
                student['Outcome (Based on clearance in end-semester or makeup exam)'] = 'Improved'
            else:
                student['Outcome (Based on clearance in end-semester or makeup exam)'] = 'Not Improved'
        return all_data

    def filter_students(self, students, learner_type, slow_thresh, advanced_thresh):
        if learner_type == 'slow':
            filtered = [s for s in students if s['MidtermPercentage'] < slow_thresh]
        else:
            filtered = [s for s in students if s['MidtermPercentage'] > advanced_thresh]
        filtered.sort(key=lambda s: s.get('Register Number of the Student', ''))
        return filtered


# ==============================================================================
# FORMATTERS
# ==============================================================================

class BaseFormatter:
    BODY_FONT = "Times New Roman"

    def __init__(self):
        self.signature_image_bytes: bytes | None = None  # raw image bytes (decrypted from DB)
        self.faculty_name: str | None = None
        self.learner_type: str | None = None
        self.subject: str | None = None
        self.semester: str | None = None

    def get_year_semester_string(self, roman):
        return SEMESTER_MAPPING.get(str(roman).strip().lower(), str(roman))

    def set_cell_properties(self, cell, text, bold=False, font_size=10, align='LEFT', valign='TOP', font_name=None):
        cell.text = ''
        p = cell.add_paragraph()
        run = p.add_run(str(text))
        run.font.size = Pt(font_size)
        run.bold = bold
        if font_name:
            run.font.name = font_name
        p.alignment = getattr(WD_ALIGN_PARAGRAPH, str(align).upper(), WD_ALIGN_PARAGRAPH.LEFT)
        cell.vertical_alignment = getattr(WD_ALIGN_VERTICAL, str(valign).upper(), WD_ALIGN_VERTICAL.TOP)

    def add_signature_line(self, doc_or_cell):
        if self.signature_image_bytes:
            p_img = doc_or_cell.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            try:
                run = p_img.add_run()
                run.add_picture(io.BytesIO(self.signature_image_bytes), width=Inches(0.8))
            except Exception:
                pass
                
        p_text = doc_or_cell.add_paragraph()
        p_text.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_text.add_run("_" * 40 + "\n")
        if self.faculty_name:
            p_text.add_run(f"{self.faculty_name}\n")
        p_text.add_run("Signature of the\nsubject teacher / class coordinator")

    def _add_document_header(self, cell):
        for line in ['Manipal Institute of Technology', 'MAHE Manipal', 'Computer Science and Engineering Department']:
            p = cell.add_paragraph()
            p.add_run(line).bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def _create_format1_content(self, doc, student, slow_threshold, fast_threshold, page_break_before=False):
        heading = doc.add_heading(
            'Format 1. Assessment of the learning levels of the students:', level=2
        )
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if page_break_before:
            heading.paragraph_format.page_break_before = True
        ct = doc.add_table(rows=5, cols=1)
        ct.style = 'Table Grid'
        self._add_document_header(ct.cell(0, 0))

        st_table = ct.cell(1, 0).add_table(rows=4, cols=2)
        mapping = [
            ('Name of the Student:', 'Student Name'),
            ('Registration Number:', 'Register Number of the Student'),
            ('Course:', 'Subject Name'),
            ('Year /semester:', 'Semester'),
        ]
        for i, (label, key) in enumerate(mapping):
            self.set_cell_properties(st_table.cell(i, 0), label)
            val = student.get(key, '') if key != 'Semester' else self.get_year_semester_string(student.get(key, ''))
            self.set_cell_properties(
                st_table.cell(i, 1),
                str(val).upper() if key == 'Subject Name' else val,
                font_name=self.BODY_FONT
            )

        pt = ct.cell(2, 0).add_table(rows=3, cols=4)
        pt.style = 'Table Grid'
        pt.cell(0, 2).merge(pt.cell(0, 3))
        for text, col in [('Sr. No.', 0), ('Parameter', 1), ('Weightage in Percentage', 2)]:
            self.set_cell_properties(pt.cell(0, col), text, bold=True, align='CENTER')

        self.set_cell_properties(pt.cell(1, 0), '1', align='CENTER')
        self.set_cell_properties(pt.cell(1, 1), f"Scores obtained by student class test / internal examination...\nConsidered Midterm exam conducted for {MIDTERM_TOTAL_MARKS}M:")
        self.set_cell_properties(pt.cell(1, 2), f"{student.get('MidtermPercentage', 0):.2f}", align='CENTER', font_name=self.BODY_FONT)
        self.set_cell_properties(pt.cell(1, 3), "> %", align='CENTER')
        self.set_cell_properties(pt.cell(2, 0), '2', align='CENTER')
        self.set_cell_properties(pt.cell(2, 1), 'Performance of students in preceding university examination')
        self.set_cell_properties(pt.cell(2, 2), str(student.get('CGPA (up to previous semester)', '')), align='CENTER', font_name=self.BODY_FONT)
        self.set_cell_properties(pt.cell(2, 3), "> %", align='CENTER')

        ct.cell(3, 0).text = "Total Weightage"
        fc = ct.cell(4, 0)

        p1 = fc.add_paragraph()
        p1.add_run(f"1. Midterm score less than {slow_threshold}% considered as a ")
        r1 = p1.add_run("slow learner")
        if self.learner_type == 'slow':
            r1.font.underline = True; r1.font.color.rgb = RGBColor(255, 0, 0)

        p2 = fc.add_paragraph()
        p2.add_run(f"2. Midterm score more than {fast_threshold}% considered as an ")
        r2 = p2.add_run("advanced learner")
        if self.learner_type == 'advanced':
            r2.font.underline = True; r2.font.color.rgb = RGBColor(255, 0, 0)
        p2.add_run(" **")

        pd_ = fc.add_paragraph()
        pd_.add_run(f"Date: {datetime.now().strftime('%d-%m-%Y')}").font.name = self.BODY_FONT
        self.add_signature_line(fc)

    def _create_format2_content(self, doc, student, page_break_before=False):
        h = doc.add_paragraph()
        h.style = 'Heading 2'
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if page_break_before:
            h.paragraph_format.page_break_before = True
        h.add_run('Format -2 Report of performance/ improvement for ')
        r1 = h.add_run('slow')
        if self.learner_type == 'slow':
            r1.font.underline = True; r1.font.color.rgb = RGBColor(255, 0, 0)
        h.add_run(' and ')
        r2 = h.add_run('advanced')
        if self.learner_type == 'advanced':
            r2.font.underline = True; r2.font.color.rgb = RGBColor(255, 0, 0)
        h.add_run(' learners')

        ht = doc.add_table(rows=1, cols=1)
        self._add_document_header(ht.cell(0, 0))
        ct = doc.add_table(rows=8, cols=2)
        ct.style = 'Table Grid'

        fields = [
            ('1. Registration Number', 'Register Number of the Student'),
            ('2. Name of the student', 'Student Name'),
            ('3. Course', 'Subject Name'),
            ('4. Year/Semester', 'Semester'),
            ('5. Midterm Percentage', 'MidtermPercentage'),
            ('6. Activities/ Measure/special programs\ntaken to improve the performance', 'Actions taken to improve performance'),
            ('7. Progress', 'Outcome (Based on clearance in end-semester or makeup exam)'),
            ('Comments/remarks', 'Remarks if any'),
        ]
        for i, (label, key) in enumerate(fields):
            self.set_cell_properties(ct.cell(i, 0), label)
            val = student.get(key, '')
            if key == 'Semester': val = self.get_year_semester_string(val)
            elif key == 'MidtermPercentage': val = f"{val:.2f}%"
            elif key == 'Subject Name': val = str(val).upper()
            self.set_cell_properties(ct.cell(i, 1), str(val).replace(';', '\n'), font_name=self.BODY_FONT)
        pd_ = doc.add_paragraph()
        pd_.add_run(f"\nDate: {datetime.now().strftime('%d-%m-%Y')}").font.name = self.BODY_FONT
        self.add_signature_line(doc)


class Format1DocxFormatter(BaseFormatter):
    def format(self, students, slow_threshold, fast_threshold):
        doc = Document()
        for sec in doc.sections: sec.top_margin = Inches(0.5)
        for i, s in enumerate(students):
            self._create_format1_content(
                doc, s, slow_threshold, fast_threshold, page_break_before=(i > 0)
            )
        return doc


class Format2DocxFormatter(BaseFormatter):
    def format(self, students, slow_threshold, fast_threshold):
        doc = Document()
        for sec in doc.sections: sec.top_margin = Inches(0.5)
        for i, s in enumerate(students):
            self._create_format2_content(doc, s, page_break_before=(i > 0))
        return doc


class Format3DocxFormatter(BaseFormatter):
    def format(self, students, slow_threshold, fast_threshold):
        doc = Document()
        for sec in doc.sections: sec.top_margin = Inches(0.5)
        cols = ['Sl. No', 'Reg Number', 'Name of the student', 'Midterm Percentage', 'Progress']

        if not students:
            if self.subject:
                doc.add_paragraph(f"Course: {str(self.subject).upper()}", style='Heading 3')
            t = doc.add_table(rows=1, cols=len(cols)); t.style = 'Table Grid'
            for j, c in enumerate(cols): self.set_cell_properties(t.cell(0, j), c, bold=True)
            for i in range(8):
                rc = t.add_row().cells
                self.set_cell_properties(rc[0], str(i + 1), font_name=self.BODY_FONT)
                if i == 3: self.set_cell_properties(rc[2], "NIL", font_name=self.BODY_FONT, align='CENTER')
            self.add_signature_line(doc)
            return doc

        df = pd.DataFrame(students)
        doc.add_paragraph(f"Course: {str(df['Subject Name'].iloc[0]).upper()}", style='Heading 3')
        t = doc.add_table(rows=1, cols=len(cols)); t.style = 'Table Grid'
        for j, c in enumerate(cols): self.set_cell_properties(t.cell(0, j), c, bold=True)
        for idx, row in df.reset_index(drop=True).iterrows():
            rc = t.add_row().cells
            self.set_cell_properties(rc[0], str(idx + 1))
            self.set_cell_properties(rc[1], row['Register Number of the Student'])
            self.set_cell_properties(rc[2], row['Student Name'])
            self.set_cell_properties(rc[3], f"{row['MidtermPercentage']:.2f}")
            self.set_cell_properties(rc[4], row['Outcome (Based on clearance in end-semester or makeup exam)'])
        self.add_signature_line(doc)
        return doc


class Format1And2DocxFormatter(BaseFormatter):
    def format(self, students, slow_threshold, fast_threshold):
        doc = Document()
        for sec in doc.sections: sec.top_margin = Inches(0.5)
        for i, s in enumerate(students):
            # Use paragraph page_break_before instead of doc.add_page_break() — the latter
            # often produces a blank page in LibreOffice PDF when the prior section is full.
            self._create_format1_content(
                doc, s, slow_threshold, fast_threshold, page_break_before=(i > 0)
            )
            self._create_format2_content(doc, s, page_break_before=True)
        return doc


# ==============================================================================
# WRITERS
# ==============================================================================

class DocxWriter:
    def write(self, doc, out_path: str, **kwargs):
        doc.save(out_path)


class PdfWriter:
    def write(self, doc, out_path: str, sign_info: dict = None, format_choice: str = None):
        import shutil

        with tempfile.TemporaryDirectory() as td:
            temp_docx = os.path.join(td, "temp.docx")
            doc.save(temp_docx)
            try:
                subprocess.run(
                    [get_libreoffice_command(), '--headless', '--convert-to', 'pdf', '--outdir', td, temp_docx],
                    check=True,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    timeout=120,
                )
            except subprocess.TimeoutExpired as exc:
                raise RuntimeError(
                    "PDF conversion timed out. Try a smaller document or reduce the number of students."
                ) from exc
            except subprocess.CalledProcessError as exc:
                raise RuntimeError(
                    "PDF conversion failed. Ensure LibreOffice is available in the worker container."
                ) from exc

            temp_pdf = os.path.join(td, "temp.pdf")
            if not os.path.exists(temp_pdf):
                raise RuntimeError(
                    "PDF conversion failed: LibreOffice did not produce an output file."
                )

            shutil.move(temp_pdf, out_path)

            if sign_info and sign_info.get('should_sign') and format_choice in ['1', '2', '4', '5']:
                if not sign_info.get('key_bytes') or not sign_info.get('cert_bytes'):
                    raise RuntimeError(
                        "Digital signing is enabled but signing credentials are missing."
                    )
                sign_pdf(
                    out_path,
                    key_bytes=sign_info['key_bytes'],
                    cert_bytes=sign_info['cert_bytes'],
                    image_bytes=sign_info.get('image_bytes', b''),
                    password=sign_info.get('password', ''),
                )


# ==============================================================================
# CONTROLLER
# ==============================================================================

class ReportController:
    """
    Worker-side controller. Accepts fully resolved paths and decrypted bytes.
    sign_info keys: should_sign, key_bytes, cert_bytes, image_bytes, password
    """

    def __init__(self, excel_path, cgpa_path, format_choice, learner_type, slow_thresh,
                 advanced_thresh, output_type, semester, sign_info, common_comment, proofs_text,
                 grade_path=None, faculty_name=None, output_dir=None):
        self.excel_path = excel_path
        self.cgpa_path = cgpa_path
        self.grade_path = grade_path
        self.format_choice = format_choice
        self.learner_type = learner_type
        self.slow_threshold = slow_thresh
        self.advanced_threshold = advanced_thresh
        self.output_type = output_type
        self.semester = semester.lower().strip()
        self.sign_info = sign_info or {}
        self.common_comment = common_comment
        self.proofs_text = proofs_text
        self.faculty_name = faculty_name
        self.output_dir = output_dir  # worker writes into the task-specific subdir
        self.reader = DataReader()
        self.processor = StudentDataProcessor()
        self.writer = DocxWriter() if output_type == 'word' else PdfWriter()
        self.subject = ""

    def _configure_formatter(self, fmt):
        fmt.signature_image_bytes = self.sign_info.get('image_bytes')
        fmt.faculty_name = self.faculty_name
        fmt.learner_type = self.learner_type
        fmt.subject = self.subject
        fmt.semester = self.semester
        return fmt

    def run(self):
        all_data, self.subject = self.reader.read_data(self.excel_path)
        if not all_data:
            return None

        cgpa_map = self.reader.read_cgpa_map(self.cgpa_path)
        grade_map = self.reader.read_grade_map(self.grade_path, course_code=self.subject)
        students_all = self.processor.process_data(
            all_data, self.subject, self.semester, self.proofs_text, self.common_comment, cgpa_map, grade_map
        )
        filtered = self.processor.filter_students(
            students_all, self.learner_type, self.slow_threshold, self.advanced_threshold
        )

        act_f = '3' if not filtered else self.format_choice
        ds = datetime.now().strftime('%d_%m_%y')
        sn = self.semester.upper()
        sub_dir = re.sub(r'[\\/*?:"<>|]', "", self.subject.replace(' ', '_'))

        od = self.output_dir or os.path.join(
            "Learner_Monitor_Reports", f"{self.learner_type.title()}_Learners",
            f"Semester_{sn}", sub_dir
        )
        os.makedirs(od, exist_ok=True)

        if act_f == '5' and filtered:
            return self._generate_all_formats(filtered, od, ds, sn, sub_dir)

        fmt_map = {
            '1': Format1DocxFormatter, '2': Format2DocxFormatter,
            '3': Format3DocxFormatter, '4': Format1And2DocxFormatter,
        }
        fmt = self._configure_formatter(fmt_map[act_f]())
        ext = 'docx' if self.output_type == 'word' else 'pdf'
        lbl = {'1': 'Format1', '2': 'Format2', '3': 'Summary', '4': 'Combined'}.get(act_f, "Report") if filtered else "Empty_Summary"
        out_p = os.path.join(od, f'{sub_dir}_{sn}_{self.learner_type.title()}Learner_{lbl}_{ds}.{ext}')
        self.writer.write(
            fmt.format(filtered, self.slow_threshold, self.advanced_threshold),
            out_p, sign_info=self.sign_info, format_choice=act_f
        )
        return out_p

    def _generate_all_formats(self, students, od, ds, sn, sub_dir):
        ext = 'docx' if self.output_type == 'word' else 'pdf'
        results = []
        for cls, lbl, fc in [(Format1And2DocxFormatter, 'Combined', '4'), (Format3DocxFormatter, 'Summary', '3')]:
            fmt = self._configure_formatter(cls())
            p = os.path.join(od, f'{sub_dir}_{sn}_{self.learner_type.title()}Learner_{lbl}_{ds}.{ext}')
            self.writer.write(
                fmt.format(students, self.slow_threshold, self.advanced_threshold),
                p, sign_info=self.sign_info, format_choice=fc
            )
            results.append(p)
        return results
