---
name: learner-classifier
description: Classify students as slow, average or advanced learners from an institutional mid-term marks sheet (.xlsx/.xls, total out of 30), and generate the LearnerPDC-style outputs — a classified spreadsheet, a chat summary, and Format 1 / Format 2 / Summary Word and PDF reports. Optionally joins a CGPA file and an end-semester grade file to fill CGPA and Improved/Not Improved progress. Use this whenever the user uploads a midterm marks sheet and mentions slow learners, fast or advanced learners, learner categorisation, learning-level assessment, remedial lists, Format 1/Format 2 reports, or the LearnerPDC app — even if they only say "find the weak students" or "who scored below 40%".
---

# Learner Classifier

Ports the core logic of the LearnerPDC app (`farm/worker/logic.py`) so the same
classification and reports can be produced straight from a marks sheet, without
running the Docker stack. All the logic lives in `scripts/classify_learners.py`;
use it rather than re-implementing, because the reports must match what the app
produces byte-for-byte in their rules (strict thresholds, reg-number cleanup,
grade mapping, empty-summary fallback).

## The rules

- **Percentage** = Total(30) ÷ 30 × 100. Blank, `AB` or any non-numeric total becomes **0%**, the same as in the app, so those students land in the slow list.
- **Slow learner**: percentage **strictly below** the slow threshold (default 40).
- **Advanced learner**: percentage **strictly above** the advanced threshold (default 90).
- Everyone else is **Average** (spreadsheet only; the app has no report for them).
- Exactly 40% or exactly 90% is Average. The older `learner.py` used ≤ / ≥ and 80; the active `farm/` code uses strict < / > with 40 / 90, and that is what this skill follows.
- Lists are sorted by registration number.

## Input: the institutional marks sheet

Only this layout is supported, because column positions, not headers, carry the meaning:

| Location | Content |
|---|---|
| Cell C2 | Subject as `[CODE] Subject Name` → shown as `Subject Name (CODE)` |
| Rows 1–12 | Header block (question numbers, marks, CO, Bloom's level) — skipped |
| Column D | Student name |
| Column E | Registration number (its header says "Question"; that's expected) |
| Column Y | Total out of 30 |

Rows without a registration number are skipped. Registration numbers are
normalised (spaces and non-breaking spaces stripped, trailing `.0` removed,
uppercased) so they join cleanly with the CGPA and grade files.

`assets/sample_marks.xlsx` is a small synthetic sheet in this layout, useful for a test run or for showing the user what's expected.

If the script reports fewer than 25 columns or no student rows, the file isn't in
this layout. Tell the user what was expected instead of guessing columns.

**Optional files**
- `--cgpa`: column A = enrollment number, column B = CGPA (first visible sheet; .xlsx/.xls/.csv).
- `--grades`: either (reg, grade) or (reg, course, grade). With three columns, rows are filtered to the marks sheet's course by code or name. Grades A+, A, B, C, D, E, S → **Improved**; any other grade → **Not Improved**; no grade → blank.

## Workflow

1. **Collect inputs.** You need the marks file and the semester (Roman numeral i–viii; it becomes e.g. "II Year/ III semester"). If the user hasn't given the semester, ask, since it is printed on every report. Also accept, if offered: thresholds, which learner type, report format, Word/PDF, the "actions taken to improve performance" text, remarks, faculty name, and a signature image. Don't block on the optional ones.

2. **Run the script** (Python 3 with pandas, openpyxl, python-docx; `xlrd` for `.xls`; LibreOffice `soffice` for PDF):

   ```bash
   python scripts/classify_learners.py MARKS.xlsx --semester iii \
     [--cgpa CGPA.xlsx] [--grades GRADES.csv] \
     [--slow 40] [--advanced 90] [--learner-type slow|advanced|both] \
     [--format 1|2|3|4|5|none] [--output word|pdf|both] \
     [--actions "Remedial classes;Extra assignments"] [--remarks "..."] \
     [--faculty-name "Dr. X"] [--signature-image sig.png] \
     [--sheet-format xlsx|csv] --outdir learner_reports
   ```

   Defaults: both learner types, format 5, Word + PDF, xlsx sheet. In `--actions`, a `;` becomes a new line in Format 2.

   The script prints a JSON summary (counts, slow/advanced lists, non-numeric marks, CGPA/grade match counts, output paths). Read it; don't re-open the spreadsheet to count.

3. **Report back** in chat:
   - Subject, semester and thresholds used.
   - Counts of Slow / Average / Advanced out of the total.
   - The slow and advanced learners (reg no., name, %). For long lists, give the count and point to the spreadsheet.
   - Any **non-numeric marks**, named explicitly. They were scored as 0% and sit in the slow list, but they may be absentees rather than weak students, and the user should decide.
   - If CGPA or grade files were given and few rows matched, say so. A low match count usually means a reg-number format mismatch.
   - Then deliver the spreadsheet and reports as files.

## Report formats

| `--format` | Produces (per learner type) |
|---|---|
| `1` | Format 1: assessment of learning levels, one page per student (name, reg, course, semester, midterm %, CGPA, threshold notes with the relevant category underlined in red, date, signature) |
| `2` | Format 2: performance/improvement report, one page per student (midterm %, actions taken, progress, remarks) |
| `3` | Summary table: Sl. No, Reg Number, Name, Midterm %, Progress |
| `4` | Format 1 + Format 2 interleaved per student |
| `5` | Format 4 combined **and** the Summary as separate files (the app's "all formats") |
| `none` | Spreadsheet only |

If a learner type has no students, the app's behaviour is kept: a single
"Empty_Summary" report with an 8-row table and NIL in row 4.

File names follow the app: `<Subject>_<SEM>_<Slow|Advanced>Learner_<Label>_<dd_mm_yy>.<ext>`.
The spreadsheet is `<Subject>_<SEM>_Classified_<dd_mm_yy>.xlsx`, with sheets *All Students*, *Slow (<N%)* and *Advanced (>N%)*.

The institution header (Manipal Institute of Technology / MAHE Manipal /
Computer Science and Engineering Department) is hard-coded as in the app. If the
user is from another department, edit `HEADER_LINES` at the top of the script.

## What is deliberately not included

The app's **cryptographic PDF signing** (PEM key and certificate via endesive)
is not ported. It needs private key material that the app keeps encrypted in
its vault, and handling that in a chat would be the wrong place for it. A
signature *image* and faculty name can still be placed on the reports. If the
user needs digitally signed PDFs, point them to the LearnerPDC app.
