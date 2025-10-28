# 🕒 Automated Time Table Generator

This is a **Flask-based web application** that allows you to generate both **Exam** and **Class Time Tables** seamlessly in one place.  
You can run both modules together for smooth and automated scheduling.

---

## 🚀 Features

- Generate **Exam Time Table** and **Class Time Table** at once  
- Simple and interactive **web interface**  
- Built using **Flask (Python)**  
- Upload input data (like subjects, teachers, and schedules)  
- Automatically generates optimized timetables  
- Option to **download or view** generated timetables  

---

## 🛠️ Tech Stack

| Component | Technology |
|------------|-------------|
| Backend | Python (Flask) |
| Frontend | HTML, CSS, Bootstrap |
| Database | Using Excel as DataBase|
| Deployment | Localhost (can be deployed on Heroku, Render, etc.) |

---

## 📂 Project Structure

TimeTable-Generator/
│
├── app.py # Flask main file
├── app2.py # Flask Exam file
├── run_both.py # Runs both Exam and Class modules together
├── templates/ # HTML templates
├── static/ # CSS, JS, and assets
├── uploads/ # Uploaded input files
├── uploadsExam/ # Uploaded input files for exam time table
└── README.md # Project documentation

Run both modules (Exam and Class) together:
--> python run_both.py
You will find 127.0.0.1:5000 --> Class Time Table
127.0.0.1:5001 --> Exam Time Table
Your ip:5001 --> Exam Time Table which can be accessed through Phone and Tablet which are in the same LAN Network area


