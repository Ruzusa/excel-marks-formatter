# Excel Marks Formatter

This Python script automates the process of calculating total marks and assigning grades to students using data from an Excel sheet. It's a simple and efficient solution for teachers, educators, or anyone working with student performance data.

---

## Features

-Reads student marks from an Excel file (`marks.xlsx`)

-Calculates total marks for each student

-Assigns grades based on score thresholds

-Saves the updated data to a new Excel file (`formatted_marks.xlsx`)

-Built using Python and `openpyxl`

---

## How It Works

1. Reads each student's Math, Physics, and Chemistry marks
2. Calculates total marks
3. Assigns a grade:
   - **A**: 240 and above
   - **B**: 180–239
   - **C**: 179 or below
4. Writes results (Total + Grade) back to the Excel sheet

---

## Requirements

-Python 3.x
-`openpyxl` library

Install with:

`pip install openpyxl`
