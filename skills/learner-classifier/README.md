# Learner Classifier — Claude Skill (mini project)

A standalone [Claude Agent Skill](https://docs.claude.com/en/docs/agents-and-tools/agent-skills/overview) that brings LearnerPDC's classification and report engine into a chat. A faculty member uploads a mid-term marks sheet, and Claude classifies the students and produces the same Format 1 / Format 2 / Summary reports as the app. Docker, MongoDB and Celery aren't needed.

The logic is a port of `farm/worker/logic.py`, so the outputs follow the app's rules exactly.

## What it does

| Step | Detail |
|---|---|
| Parse | Institutional marks sheet: subject from `C2`, name / reg no. / total from columns D / E / Y, 12-row header skipped |
| Classify | `% = total / 30 × 100`. **Slow** if strictly < 40%, **Advanced** if strictly > 90%, otherwise Average. `AB` or blank → 0% (flagged for review) |
| Join (optional) | CGPA file (reg → CGPA) and end-sem grade file (A+…E, S → *Improved*, else *Not Improved*), with reg numbers normalised |
| Report | Classified `.xlsx` (All / Slow / Advanced sheets) and Word + PDF reports: Format 1, Format 2, Format 1+2 combined, Summary, or the empty-summary fallback |

## Layout

```
skills/learner-classifier/
├── SKILL.md                      # instructions Claude follows (frontmatter + workflow)
├── scripts/classify_learners.py  # all logic; prints a JSON summary
├── assets/sample_marks.xlsx      # synthetic sheet in the institutional layout
└── requirements.txt
```

## Use it

**As a Claude skill:** zip the `learner-classifier/` folder and upload it under Settings → Capabilities → Skills on claude.ai, or drop it into `~/.claude/skills/` for Claude Code. Then upload a marks sheet and ask Claude to classify the slow learners.

**From the command line:**

```bash
pip install -r requirements.txt          # LibreOffice (soffice) needed for PDF output
python scripts/classify_learners.py assets/sample_marks.xlsx --semester iii \
  --format 5 --output both --faculty-name "Dr. X" --outdir out
```

Main options: `--cgpa`, `--grades`, `--slow 40`, `--advanced 90`, `--learner-type slow|advanced|both`, `--format 1|2|3|4|5|none`, `--output word|pdf|both`, `--actions "Remedial classes;Extra assignments"`, `--remarks`, `--signature-image`.

## Deliberately not included

The app's cryptographic PDF signing (PEM key and certificate via endesive) is left out. Private keys belong in the app's Fernet-encrypted vault, not in a chat session. The reports can still carry a signature image and the faculty name.

## Data privacy

Only the synthetic `sample_marks.xlsx` is committed. Real marks sheets and generated reports contain student personal data and stay out of git (`*.xlsx` / `*.docx` are ignored repo-wide).
