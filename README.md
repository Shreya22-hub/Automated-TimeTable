:: Description

The Automated Timetable Generator is a Python-based web application designed to simplify and automate the process of creating academic timetables for colleges and universities. It efficiently generates timetables for different branches, semesters, and faculties by analyzing course details, subject codes, and faculty availability.

The system minimizes manual effort and errors by allowing administrators to upload subject and faculty data, process it, and automatically produce optimized timetables. It ensures that no faculty is double-booked and that classroom and slot conflicts are avoided, resulting in a perfectly balanced schedule.

:: Key Features

:: Automated Timetable Generation: Generates timetables for all branches and faculties with a single click.

:: Faculty Timetable View: Each faculty can view their individual schedule.

::  Branch-wise Timetables: Supports multiple departments and semesters simultaneously.

:: CSV/Excel Upload Support: Admins can upload subject and faculty details directly from .csv or .xlsx files.

:: Conflict-Free Scheduling: Automatically detects and avoids overlapping class or faculty slots.

:: Data Storage: Saves generated timetables for easy retrieval and updates.

:: Web-Based Interface: Simple and intuitive dashboard for admins and faculty members.

:: Tech Stack

:: Frontend: HTML, CSS, JavaScript

:: Backend: Python (Flask Framework)

:: Database:  CSV-based Data Management

 :: Libraries Used: Pandas, Flask 

 :: How It Works

Admin uploads faculty and subject data files.

The system processes the uploaded data and maps subjects to faculties.

The algorithm automatically assigns slots and rooms while avoiding clashes.

Generated timetables can be viewed or downloaded branch-wise and faculty-wise.

:: Objective

To automate the tedious manual process of timetable creation, saving time and effort for academic administrators, while ensuring accuracy and consistency across all departments and faculty schedules.
