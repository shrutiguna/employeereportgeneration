# Employee Report Generator Bot

A UiPath RPA bot that reads employee data from an Excel file, categorizes employees based on salary thresholds, generates a formatted report with summary metrics, and logs execution details.

---

## What It Does

1. Reads employee data from `employees.xlsx` (Emp ID, Name, Department, City, Salary)
2. Validates salary data and categorizes each employee as **High Earner**, **Standard**, or **Invalid Salary**
3. Adds a "Category" column
4. Generates `output_report.xlsx` with two sheets:
   - **Report** — Full employee data with category column
   - **Summary** — Metrics including totals and run details
5. Applies professional Excel formatting via VBA
6. Creates a timestamped execution log file
7. Displays a summary message box on completion

---

## Prerequisites

| Requirement | Version |
|------------|--------|
| UiPath Studio | Community Edition 2023.10+ |
| Project Type | Windows (Modern Design) |
| UiPath.Excel.Activities | 2.24.2 |
| UiPath.System.Activities | 26.x |
| Expression Language | VB.NET |

---

## Project Structure

```text
EmployeeReportGenerator/
|-- Main.xaml
|-- employees.xlsx
|-- formate_employee_report.vbs
|-- output_report.xlsx
|-- execution_log_DD-MM-YYYY.txt
|-- project.json
|-- project.uiproj
`-- entry-points.json
```

---

## Input File Format

`employees.xlsx` must have a sheet named **Employees**

| Column | Header | Example |
|-------|--------|--------|
| A | Emp ID | EMP001 |
| B | Name | Aarav Sharma |
| C | Department | IT |
| D | City | Chennai |
| E | Salary (₹) | 72000 |

---

## Configuration

| Variable | Default | Description |
|---------|--------|-------------|
| `strInputFile` | employees.xlsx | Input Excel file |
| `strOutputFile` | output_report.xlsx | Output Excel file |
| `strInputSheet` | Employees | Input sheet name |
| `strOutputSheet` | Report | Output sheet name |
| `dblThreshold` | 50000 | Salary threshold |
| `intSalaryColIndex` | 4 | Salary column index (0-based) |

---

## Workflow Overview

```text
Main Sequence
|-- Start Time Capture
|-- Log Start
|-- Try-Catch Block
    |-- Validate Input File
    |-- Read Excel Data
    |-- Validate Data Rows
    |-- Add Category Column
    |-- Categorize Employees
    |-- Write Report Sheet
    |-- Generate Summary Sheet
    |-- Apply Excel Formatting (VBA)
    |-- Log Completion
    |-- Show Message Box
|-- Catch Errors
|-- Finally Write Logs
```

---

## Categorization Logic

- Salary empty → **Invalid Salary**
- Salary not numeric → **Invalid Salary**
- Salary > 50,000 → **High Earner**
- Salary ≤ 50,000 → **Standard**

---

## Output

### Report Sheet
- Full employee data with Category column
- Styled formatting with headers and highlights

### Summary Sheet

| Metric | Value |
|-------|------|
| Total Employees | Example: 15 |
| High Earners | Example: 9 |
| Standard | Example: 6 |
| Invalid Salary | Example: 0 |
| Threshold Used | ₹50000 |
| Run Date | Auto-generated |

---

## Logging

- UiPath Output Panel → Real-time logs
- Execution Log File → Stored in project folder

---

## Error Handling

Uses Try-Catch-Finally:

- Try → Main execution
- Catch → Error handling
- Finally → Always writes log file

---

## How to Run

1. Clone or download the project
2. Open in UiPath Studio
3. Ensure `employees.xlsx` is present
4. Click Run (F5)

---

