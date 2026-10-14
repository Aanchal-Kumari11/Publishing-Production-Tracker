# Publishing Production Tracker

A Python-based tool designed to help organize and monitor tasks in a book publishing workflow.

## Features

- Reads publishing task data from a CSV file.
- Cleans and standardizes task information.
- Groups tasks by book title while preserving repeated and duplicate rows.
- Prioritizes tasks based on deadlines and priority levels.
- Generates reminder labels for overdue and upcoming tasks.
- Creates a formatted Excel workbook with color-coded deadlines, priorities, and statuses.
- Generates a text summary of production progress and tasks requiring follow-up.

## Technologies Used

- Python
- CSV file processing
- Excel report generation with OpenPyXL

## How to Run

Install the required package:

```bash
python -m pip install openpyxl
```

Run the tracker with your input CSV:

```bash
python publishing_production_tracker.py sample_publishing_tasks.csv
```

The script generates an output folder containing the formatted Excel report and a text summary.

## Input Data

The sample CSV contains book titles, tasks, production stages, assigned team members, deadlines, priorities, statuses, dependencies, and notes.

You can replace the sample data with your own CSV file using the supported column names.

## Project Purpose

This project demonstrates Python automation for production task tracking, deadline monitoring, data organization, and report generation.
