This project is a Timetable Generator Software designed for IIIT Dharwad.
It automates the process of scheduling lectures and labs for various departments, batches, and faculty members while ensuring that no clashes occur between rooms, teachers, or students.
The application also allows administrators to add, view, and modify timetable data through a simple web interface.


Objectives

To eliminate manual timetable creation errors and save time.

To ensure no overlapping of classes for any batch or faculty.

To generate department-wise and faculty-wise timetables automatically.

To store timetable data in a structured, easily accessible format.

To allow future extensions like exam scheduling and attendance management.

Features
Automatic timetable generation using a greedy scheduling algorithm
 Separate views for faculty, students, and administrators
 Room and batch conflict checking
Configuration of slots, working days, and lab/theory hours
Export timetable to Excel / CSV
Web-based interface built with Flask, HTML, CSS, and JavaScript
Easy integration with other campus systems

Tech Stack
Component	Technology
Frontend	HTML5, CSS3, JavaScript
Backend	Python (Flask Framework)
Database	SQLite (via SQLAlchemy ORM)
Tools	Pandas, OpenPyXL
Version Control	Git & GitHub
