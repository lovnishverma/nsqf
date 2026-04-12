# NSQF Exam Document Generator — NIELIT Ropar

Production-ready Flask application that auto-generates exam documents
(Theory/Practical Attendance, Internal Assessment, Compiled Result HQ)
for all 12 NSQF courses via Template-Based Injection.

---

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Add your formatted .xlsx templates into templates/xlsx/
#    (see "Template File Naming" below)

# 3. Run
python app.py

# 4. Open http://localhost:5000
```

---

## Project Structure

```
nsqf_app/
├── app.py                  ← Flask application (main entry point)
├── requirements.txt
├── README.md
└── templates/
    ├── index.html          ← Upload UI (auto-served by Flask)
    └── xlsx/               ← ⚠️  Place ALL your .xlsx templates here
        ├── I091_theory_attendance.xlsx
        ├── I091_practical_attendance.xlsx
        ├── I091_internal_assessment.xlsx
        ├── I091_compiled_result_hq.xlsx
        ├── I092_*.xlsx
        ├── ... (see full list below)
```

---

## Template File Naming Convention

All templates go in `templates/xlsx/`. Filenames are **case-sensitive**.

### Single-module courses (1 Theory + 1 Practical)

| Course | Files needed |
|--------|-------------|
| I091 — Essentials of AI | `I091_theory_attendance.xlsx`, `I091_practical_attendance.xlsx`, `I091_internal_assessment.xlsx`, `I091_compiled_result_hq.xlsx` |
| I092 — Cloud Computing | `I092_*.xlsx` (same pattern) |
| I081 — Data Curation (Python) | `I081_*.xlsx` |
| I079 — Data Annotation (Python) | `I079_*.xlsx` |
| I060 — Comp. App, Accounting & Publishing | `I060_*.xlsx` |

### I059 — Certified Data Entry (has Typing sheet)

```
I059_theory_attendance.xlsx
I059_practical_attendance.xlsx
I059_typing_attendance.xlsx       ← extra typing test sheet
I059_internal_assessment.xlsx
I059_compiled_result_hq.xlsx
```

### Two-theory-module courses (M1 + M2 theory sheets)

| Course | Files needed |
|--------|-------------|
| I090 — Full Stack Dev | `I090_theory1_attendance.xlsx`, `I090_theory2_attendance.xlsx`, `I090_practical_attendance.xlsx`, `I090_internal_assessment.xlsx`, `I090_compiled_result_hq.xlsx` |
| I077 — Multimedia Dev | `I077_*.xlsx` (same pattern) |
| I078 — ITeS BPO Voice | `I078_*.xlsx` |
| E063 — Asst. Computer Technician | `E063_*.xlsx` |
| E058 — IoT Assistant | `E058_*.xlsx` |
| E059 — IoT Associate | `E059_*.xlsx` |

---

## Cell Injection Mapping (per template type)

### Theory / Practical Attendance
| Cell | Value |
|------|-------|
| B1   | Institute Name |
| B2   | Course Name |
| H3   | Exam Cycle (first Exam Date in batch) |
| Row 5+ | Student data (Theory) |
| Row 6+ | Student data (Practical) |

Row columns: A=S.No, B=Roll, C=RegNo, D=Name, E=Gender(initial), F=Father, G=Mother, H=DOB, I=ExamDate, J=ExamTime  
Practical adds blank K–N columns (physical signature/marks).

### Typing Attendance (I059 only)
Same as Theory but: G=DOB (no Mother column), H=ExamDate, I=ExamTime. J–L blank.

### Internal Assessment
| Cell | Value |
|------|-------|
| B1 | Institute Name |
| B2 | Course Name |
| Row 5+ | A=SNo, B=Roll, C=RegNo, D=Name, E=Gender, F=Father, G=DOB |

### Compiled Result HQ
| Cell | Value |
|------|-------|
| D2 | Institute Name |
| D4 | Course Name |
| Row 9+ | A=SNo, B=Roll, C=RegNo, D=Name, E=Gender, F=Father, G=Mother, H=DOB |

---

## Input CSV Schema

All 12 columns are required:

| Column | Description |
|--------|-------------|
| `Course_name` | Course ID (`I091`) or full label |
| `Module_Short_Name` | `M1` / `M2` — for multi-module courses |
| `Candidate_registration_no` | NIELIT reg. number |
| `Candidate_Name` | Full candidate name |
| `Candidate_Gender` | Male/Female or M/F |
| `Father_Name` | Father's full name |
| `Mother_Name` | Mother's full name |
| `Candidate_Birth_date` | Date of birth |
| `Institute_Name` | Affiliated institute |
| `Roll_Number` | Exam roll number |
| `Exam Date` | Scheduled date |
| `Exam Time` | Scheduled time |

### Multi-course CSVs
A single CSV can contain **multiple courses and institutes**. The app
groups by `(Institute_Name, Course_name)` and generates separate
document sets for each group. Files are organised in the ZIP by institute folder.

---

## Output ZIP Structure

```
NSQF_Documents.zip
├── NIELIT_Ropar/
│   ├── NIELIT_Ropar_I091_Theory_Attendance.xlsx
│   ├── NIELIT_Ropar_I091_Practical_Attendance.xlsx
│   ├── NIELIT_Ropar_I091_Internal_Assessment.xlsx
│   └── NIELIT_Ropar_I091_Compiled_Result_HQ.xlsx
├── Another_Institute/
│   └── ...
└── UNRECOGNISED_COURSES.txt   ← only present if unknown course names found
```

---

## API Endpoints

| Method | URL | Description |
|--------|-----|-------------|
| GET | `/` | Upload UI |
| POST | `/generate` | Process CSV → return ZIP |
| GET | `/courses` | JSON list of all registered courses |

### POST /generate — error responses

```json
// 400 — missing columns
{ "error": "CSV is missing required columns.", "missing_columns": ["Roll_Number"] }

// 400 — no recognised courses
{ "error": "No recognised courses found in CSV.", "unrecognised_courses": ["XYZ999"] }

// 500 — no files generated (templates missing)
{ "error": "No documents were generated. Verify that templates exist..." }
```

---

## Adding a New Course

1. Add an entry to `COURSE_REGISTRY` in `app.py`
2. Add corresponding entries to `COURSE_META` in `templates/index.html`
3. Place the template `.xlsx` files in `templates/xlsx/`
4. No other changes needed.
