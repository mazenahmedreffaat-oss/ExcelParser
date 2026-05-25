# Benha University — Transcript Parser

A lightweight Python script that parses Benha University Excel transcripts and returns all student data as clean, valid JSON. No GUI, no database — just input an `.xlsx` file and get structured JSON out.

---

## Requirements

- Python 3.8+
- pip packages:

```
pandas
openpyxl
```

Install dependencies:

```bash
pip install pandas openpyxl
```

---

## File Structure

```
project/
│
├── parser.py          # The main parser script
├── README.md          # This file
└── transcript.xlsx    # Your input Excel file (not included)
```

---

## How to Run

### Option 1 — Command Line (prints JSON to terminal)

```bash
python parser.py transcript.xlsx
```

### Option 2 — Save output to a JSON file

```bash
python parser.py transcript.xlsx > output.json
```

### Option 3 — Import as a module in your backend

```python
from parser import parse_transcript

result = parse_transcript("transcript.xlsx")

student = result["student_info"]
courses = result["courses"]

print(student["name"])        # e.g. "ابانوب مينا محسن"
print(len(courses))           # e.g. 56
```

---

## JSON Output Structure

The function returns a single dict with two keys:

```json
{
  "student_info": { ... },
  "courses": [ ... ]
}
```

### `student_info` fields

| Field | Type | Description |
|---|---|---|
| `name` | string | Full student name (Arabic) |
| `student_id` | string | University student code |
| `current_level` | string | Current academic level (raw Arabic string) |
| `faculty` | string | Faculty name |
| `university` | string | University name |
| `overall_grade` | string | Overall grade letter (e.g. `D`, `B+`) |
| `cumulative_gpa` | float | Final cumulative GPA |
| `passed_hours` | int | Total passed credit hours |
| `national_id` | string | National ID number |
| `date_of_birth` | string | Date of birth in `YYYY-MM-DD` format |
| `nationality` | string | Nationality (Arabic) |

### `courses` — each course object

| Field | Type | Description |
|---|---|---|
| `level` | int | Academic level (1, 2, 3, 4) |
| `academic_year` | string | e.g. `"2023-2024"` |
| `semester` | string | `"S1"`, `"S2"`, or `"Summer"` |
| `course_name` | string | Course name in Arabic |
| `credits` | int | Credit hours |
| `score` | float or null | Numeric score (null if absent or `*`) |
| `grade` | string | Grade string (e.g. `A+`, `D`, `F`, `راسب لائحة`, `غياب`) |
| `grade_points` | float | Grade points value |
| `pts_x_hrs` | float | Grade points × credit hours |
| `status` | string | `Pass`, `Fail`, `Fail (Regulation)`, or `Absent` |
| `cumulative_gpa` | float | Cumulative GPA at end of this semester |
| `semester_gpa` | float | GPA for this semester |

---

## Status Values Explained

| Status | Meaning |
|---|---|
| `Pass` | Student passed the course |
| `Fail` | Grade is `F` or score below 50 |
| `Fail (Regulation)` | Failed due to university regulation (`راسب لائحة`) |
| `Absent` | Student was absent (`غياب`) |

---

## Important Notes

- The parser only supports `.xlsx` format. If your file is `.xls`, open it in Excel and save as **Excel Workbook (.xlsx)** first.
- The parser is built around Benha University's specific Excel layout (hardcoded column positions). It will not work on other university formats without modification.
- A student can appear at the same level across multiple academic years (e.g. Level 3 repeated across 2023-2024, 2024-2025, 2025-2026). This is normal and handled correctly.
- Courses with `score: null` had a `*` in the score column in the original transcript, meaning the score was not yet recorded.

---

## Example Full Output (shortened)

```json
{
  "student_info": {
    "name": "ابانوب مينا محسن وديع صبرى رزق الله",
    "student_id": "26120210199637",
    "current_level": "المستوى الحالى :المستوي الرابع / الحسابات العلمية",
    "faculty": "كلية الحاسبات والذكاء الاصطناعى",
    "university": "جامعة بنها",
    "overall_grade": "D",
    "cumulative_gpa": 2.086,
    "passed_hours": 101,
    "national_id": "30212151201255",
    "date_of_birth": "2002-12-15",
    "nationality": "مصرى"
  },
  "courses": [
    {
      "level": 1,
      "academic_year": "2021-2022",
      "semester": "S1",
      "course_name": "أساسيات علوم الحاسب",
      "credits": 3,
      "score": 66.0,
      "grade": "C",
      "grade_points": 2.4,
      "pts_x_hrs": 7.2,
      "status": "Pass",
      "cumulative_gpa": 2.368,
      "semester_gpa": 2.369
    }
  ]
}
```

---

## Dependencies

| Package | Version | Purpose |
|---|---|---|
| `pandas` | >= 1.3 | DataFrame handling and deduplication |
| `openpyxl` | >= 3.0 | Reading `.xlsx` files |

---

## Contact / Issues

If the parser returns missing or incorrect data, the most likely cause is a change in the Excel layout from the university. Check the column constants at the top of `parser.py` and adjust if needed:

```python
COL_GRADE   = 22   # Column for grade letter
COL_SCORE   = 28   # Column for numeric score
COL_CREDITS = 32   # Column for credit hours
COL_COURSE  = 40   # Column for course name
```
