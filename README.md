# Automated-TimeTable
“Automated timetable scheduling project for IIIT Dharwad (Aug-Dec 2025). Generates conflict-free timetables from input data.”
Project Overview

The Automated Timetable Generator is a Python-based tool that automatically generates weekly class timetables for schools or colleges. It takes into account subjects, teachers, periods per week, and constraints to create an optimized timetable.

This project helps reduce manual scheduling errors and saves time for administrators while ensuring all teaching requirements are met.

Features

Generate a weekly timetable automatically.

Handles multiple subjects and teachers.

Supports constraints, such as unavailable periods for teachers.

Avoids teacher clashes (same teacher in two places at once).

Outputs timetable in a readable format (console or CSV).

Project Structure
Automated_Timetable/
│
├── data/                   # CSV files for subjects, teachers, and constraints
│   ├── subjects.csv
│   ├── teachers.csv
│   └── constraints.csv
│
├── src/                    # Source code
│   ├── __init__.py
│   ├── timetable_generator.py
│   └── utils.py
│
├── main.py                 # Entry point to generate timetable
├── requirements.txt        # Python dependencies
└── README.md               # Project explanation

How It Works

Data Input:

Subjects and weekly hours are listed in subjects.csv.

Teachers and their subjects are listed in teachers.csv.

Any unavailable periods or constraints are listed in constraints.csv.

Timetable Generation:

The program uses Google OR-Tools (Constraint Programming) to generate the timetable.

Ensures each subject gets its required hours per week.

Makes sure a teacher is not scheduled in two periods at the same time.

Output:

Prints the weekly timetable in the console.

Can be extended to export timetable to CSV or Excel.

Future Improvements

Add GUI interface for easier input and viewing.

Allow multiple classes/sections scheduling.

Export timetable to Excel or PDF.

Automatically resolve conflicts when constraints are unsatisfiable.
