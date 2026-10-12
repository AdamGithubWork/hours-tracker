# Hours Tracker

A simple Python hours tracker that stores tracked time in an Excel spreadsheet.

## Requirements

* Windows
* Python 3
* Microsoft Excel

## Setup

1. Download this repository from GitHub.
2. Extract the folder if it downloaded as a ZIP.
3. Make sure `Spreadsheet Jun.xlsx` is in the same folder as `main.py`.
4. Double-click **Run Tracker.bat**.

The launcher will automatically install the required Python package and start the tracker.

### Manual setup

If you prefer to run it from Command Prompt:

```text
python -m pip install -r requirements.txt
python main.py
```

## Using the Tracker

When the program starts, you'll see a list of tasks:

```text
=== Tasks ===
1. Task 2 (0.00 / 2.00 hours)
2. Task 4 (0.01 / 45.00 hours)
3. Task 6 (0.00 / 1.00 hours)
...
0. Exit
```

Enter the number of the task you want to work on.

Press **Enter** to start the timer.

Press **Enter** again to stop the timer.

The time will automatically be added to that task's **Hours Tracked** column in `Spreadsheet Jun.xlsx`.

You can return to the task menu and start another timer without closing the program.

## Excel Spreadsheet

The spreadsheet uses three columns:

| Column         | Description                        |
| -------------- | ---------------------------------- |
| Task           | Name of the task                   |
| Hours Budgeted | Total hours allocated to the task  |
| Hours Tracked  | Total time recorded by the tracker |

The program does not change the **Hours Budgeted** values.

## Files

```text
hours-tracker/
├── main.py
├── Run Tracker.bat
├── Spreadsheet Jun.xlsx
├── requirements.txt
└── README.md
```

## Notes

* Keep `Spreadsheet Jun.xlsx` in the same folder as the program.
* Close the Excel file before stopping/saving tracked time.
* The tracker saves time automatically after each completed timer session.
* Tracked hours are stored with their full precision, while the program displays them rounded to two decimal places.
