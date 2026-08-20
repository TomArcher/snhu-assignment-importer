# SNHU Assignment Importer

SNHU Assignment Importer is a simple tool for SNHU students who want an easy way to keep track of their coursework and make sure they don't fall behind or forget an assignment.

SNHU courses can contain dozens of assignments spread across multiple modules. Rather than manually entering each assignment into a spreadsheet, this application reads the assignments directly from an SNHU Brightspace Grades page and adds them to an Excel grade-tracking workbook.

The resulting spreadsheet provides a single place to see your courses, assignments, modules, available points, and grades throughout your degree program.

![Example of the SNHU assignment tracking spreadsheet](example.png)

## Features

- Imports assignments directly from an SNHU Brightspace Grades page
- Records the module number, assignment name, and maximum point value
- Eliminates the need to manually enter every assignment into Excel
- Provides a central place to keep track of assignments across multiple courses
- Helps identify upcoming or incomplete assignments so they aren't overlooked
- Tracks grades and available points as the course progresses
- Preserves the formulas and formatting in the provided Excel workbook
- Prevents the same course from being imported more than once
- Includes a blank Excel workbook template for creating your own grade tracker
- Uses a persistent browser profile so your SNHU login session can be reused

## Requirements

-   Python 3.12 or later
-   Microsoft Excel
-   Chromium browser installed through Playwright
-   An SNHU student account

## Installation

Clone the repository:

``` bash
git clone https://github.com/<your-username>/snhu-assignment-importer.git
cd snhu-assignment-importer
```

Create a Python virtual environment:

``` bash
py -3.14 -m venv .venv
```

Activate the virtual environment in PowerShell:

``` powershell
.venv\Scripts\Activate.ps1
```

Install the required Python packages:

``` bash
python -m pip install -r requirements.txt
```

Install Chromium for Playwright:

``` bash
python -m playwright install chromium
```

## Workbook Setup

The repository includes `snhu_template.xlsx`, which contains the worksheets,
tables, formulas, and formatting required by the importer.

Before importing assignments, you must add your courses to the **Data** worksheet.

For each course you want to track, add a row containing the course information.
The **Name** field is especially important because the importer uses the course
name (for example, `CS-350`) to locate the course when assignments are imported.

The Data worksheet contains the following columns:

- **Name** - Course number, such as `CS-350`
- **Term** - SNHU term in which the course is taken
- **University** - University offering the course
- **Status** - Course status
- **Credits** - Number of course credits
- **Description** - Course title
- **Instructor** - Instructor's name
- **Instructor email** - Instructor's email address
- **Notes** - Optional notes about the course

Once a course has been added to the Data worksheet, the importer can retrieve
that course information and add its assignments to the **Grades** worksheet.

You only need to enter the course information manually. The assignment rows are
created automatically by the importer.

## Usage

Activate the virtual environment and run:

``` bash
python import_course_assignments.py
```

The application will open a browser window.

1.  Log in to SNHU.
2.  Navigate to the course whose assignments you want to import.
3.  Open the course's **Grades** page.
4.  Return to the terminal.
5.  Enter the same course identifier used in the Name column of the **Data** worksheet.
6.  Review the assignments found by the application.
7.  Confirm the import when prompted.

The application extracts assignment information from the displayed
Grades page and adds the assignments to the Excel Grades table.

## How It Works

SNHU Brightspace displays each assignment with a name and maximum point
value. The application reads the text from the Grades page and
identifies assignment entries using the format of the Points field.

For example:

``` text
3-2 Journal: Peripheral Interfaces in Embedded Systems
- / 30
```

From this information, the importer determines:

``` text
Module:      3
Assignment:  3-2 Journal: Peripheral Interfaces in Embedded Systems
Max Points:  30
```

Course information such as the term and course description is obtained
from the Data worksheet in the Excel workbook.

## Privacy

The `browser-profile` directory created by Playwright may contain
browser session and authentication information. It is excluded from the
repository through `.gitignore` and should never be committed to source
control.

Personal Excel workbooks containing course or grade information should
likewise not be committed to a public repository.

## Limitations

-   The application depends on the current structure and text format of
    the SNHU Brightspace Grades page. Changes to Brightspace may require
    changes to the parser.
-   The course must exist in the workbook's Data worksheet before its
    assignments can be imported.
-   The application is intended for the supplied workbook schema.
-   SNHU authentication is performed interactively through the browser.
    The application does not collect or store SNHU usernames or
    passwords itself.

## License

This project is licensed under the MIT License. See the `LICENSE` file
for details.

## Disclaimer

This project is an independent utility and is not affiliated with,
endorsed by, or supported by Southern New Hampshire University (SNHU) or
D2L Corporation.
