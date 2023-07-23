# Readme

## Overview

This Python program allows users to select courses from an Excel file and checks if there are any time conflicts between the selected courses. If there are conflicts, they are written to a `conflicts.xlsx` file. If there are no conflicts, the selected courses are written to an `output.xlsx` file.

## Prerequisites

The program requires the following Python packages, which are listed in the `requirements.txt` file:

- `altgraph==0.17.3`
- `et-xmlfile==1.1.0`
- `numpy==1.25.1`
- `openpyxl==3.1.2`
- `pandas==2.0.3`
- `pefile==2023.2.7`
- `pyinstaller==5.13.0`
- `pyinstaller-hooks-contrib==2023.6`
- `python-dateutil==2.8.2`
- `pytz==2023.3`
- `pywin32-ctypes==0.2.2`
- `rethinkdb==2.4.9`
- `six==1.16.0`
- `thinker==1.1.1`
- `tkinterdnd2==0.3.0`
- `tzdata==2023.3`
- `xlrd==2.0.1`
- `xlwt==1.3.0`

You can install these packages using pip, the Python package installer, with the command:

```
pip install -r requirements.txt
```

Additionally, the program requires an Excel file with the following columns in the first sheet:

- Semester
- Day
- Course Start Time
- Course End Time
- ...
- Course Name (6th column)

## Usage

1. Start the program by running `start.bat`. This batch file will:
   - Activate the virtual environment
   - Install the required Python packages
   - Run the `main.py` Python script
   - Deactivate the virtual environment
   - Pause the command prompt window

2. A window will appear asking you to drag and drop the Excel file. Do so, and the window will close automatically.

3. A new window will appear displaying a list of all the courses in the Excel file. Use the Ctrl key to select multiple courses, or the Shift key to select a range of courses. Click the "Select" button when you're done.

4. The program will check if there are any time conflicts between your selected courses. If there are, it will write them to a file named `conflicts.xlsx` and print a message in the console. If there are no conflicts, it will write the selected courses to a file named `output.xlsx` and print a different message in the console.

Note: The `output.xlsx` and `conflicts.xlsx` files are overwritten each time the program is run. If you want to keep the results of a particular run, you should rename or move these files before running the program again.

## Troubleshooting

If you encounter errors while running the program, check the following:

- Are all the required packages installed?
- Is the Excel file in the correct format?
- Are the column numbers correct in the code? If your Excel file has a different structure, you might need to change the column numbers in the `check_conflicts_and_write` function.
- Is the Excel file closed while running the program? The program cannot access the file if it is open in Excel or another program.

If none of these suggestions help, you might want to check the error message and stack trace for clues about what went wrong.