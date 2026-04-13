# Student Marks Manager (Python + Excel)

A simple CLI tool to manage student marks stored in an Excel sheet using Python and `openpyxl`.  
You can recalculate totals, averages, grades, ranks, add new students, and export failed students.

## Features

- Validate header row to ensure the Excel template structure is correct
- Recalculate total, average, grade, and remarks for all students
- Assign ranks using Excel's `RANK.EQ` function (failed students get `FAIL`)
- Add new students interactively from the terminal with basic marks validation
- Export all failed students (Grade = F) to a separate Excel file

## Requirements

- Python 3.x
- `openpyxl` library

Install dependency:

```bash
pip install openpyxl
```

## Project Structure

- `main.py` – main script with all functions and CLI menu
- `students_template.xlsx` – Excel template with the expected headers

Expected headers (row 1):

```text
Name | D.O.B | Blood group | Age | Tamil | English | Maths | Physics | Chemistry | Computer science | Total | Average | Rank | Grade | Remarks
```

## How to Use

1. Make sure `students_template.xlsx` is in the same folder as `main.py`.
2. Run the script:

```bash
python main.py
```

3. Use the menu options:
   - `1` – Recalculate marks, grades, and ranks
   - `2` – Add a new student (interactive input)
   - `3` – Export failed students to a new Excel file
   - `4` – Save and exit

## Notes

- Grades are assigned based on the average:
  - A1 (≥ 90), A2 (≥ 80), B1 (≥ 70), B2 (≥ 60), C1 (≥ 50), C2 (≥ 40), F (< 40)
- Failed students get `"FAIL"` in the Rank column
- Exported failed students file keeps the same columns as the original sheet
