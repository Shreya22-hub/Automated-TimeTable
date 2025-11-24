from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash, jsonify
import os
import pandas as pd
from datetime import datetime, timedelta
from datetime import datetime
import json
import re
from werkzeug.utils import secure_filename
from itertools import cycle
import io

app = Flask(__name__, template_folder="app3")
app.secret_key = 'exam_schedule_secret_key'

# Hardcoded credentials
ADMIN_CREDENTIALS = {'username': 'admin', 'password': 'admin123'}
VIEWER_CREDENTIALS = {'username': 'view', 'password': 'view123'}

# Upload folder configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Create uploads directory if it doesn't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def generate_schedule(courses_file, rooms_file, faculty_file, start_date, courses_per_room=2):
    """
    Generate exam schedule with dynamic room divisions
    """
    SLOTS = ["Morning", "Evening"]

    def next_date(date_str, days=1):
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        return (dt + timedelta(days=days)).strftime("%Y-%m-%d")
    
    def is_special_room(room_name):
        """Check if room is in C403-C408 range"""
        match = re.match(r'C40([3-8])', room_name.strip())
        return match is not None

    # -----------------------------------------------------------
    # CREATE CONFIGURATION FILE
    # -----------------------------------------------------------
    config_dir = UPLOAD_FOLDER
    config_file = os.path.join(config_dir, "configurations.json")
    
    # Calculate max students per slot
    rooms_df = pd.read_excel(rooms_file)
    total_capacity = 0
    
    for _, r in rooms_df.iterrows():
        cap = int(r["Capacity"])
        room_name = str(r["Room"])
        is_special = is_special_room(room_name)
        
        if is_special:
            total_capacity += cap
        else:
            total_capacity += cap * courses_per_room
    
    # Create configuration data
    config = {
        "max_students_per_slot": total_capacity,
        "courses_per_room": courses_per_room,
        "start_date": start_date,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    # Save configuration file
    with open(config_file, 'w') as f:
        json.dump(config, f, indent=4)
    
    # -----------------------------------------------------------
    # READ COURSES FILE (multiple sheets: 1st year, 2nd year etc)
    # -----------------------------------------------------------
    xls = pd.ExcelFile(courses_file)
    students_data = []

    for year_index, sheet_name in enumerate(xls.sheet_names, start=1):
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)

        for col in range(df.shape[1]):
            course = df.iloc[0, col]
            students = df.iloc[1:, col].dropna().astype(str).tolist()

            if str(course).strip():  # valid course name
                students_data.append({
                    "Year": year_index,
                    "Course": str(course).strip(),
                    "Students": students
                })

    # -----------------------------------------------------------
    # READ ROOMS AND PREPARE ROOM INFORMATION
    # -----------------------------------------------------------
    rooms = []

    for _, r in rooms_df.iterrows():
        cap = int(r["Capacity"])
        room_name = str(r["Room"])
        is_special = is_special_room(room_name)
        
        # Calculate capacity per division
        if is_special:
            capacity_per_division = cap // courses_per_room
        else:
            capacity_per_division = cap
        
        rooms.append({
            "Room": room_name,
            "TotalCapacity": cap,
            "CapacityPerDivision": capacity_per_division,
            "IsSpecial": is_special
        })

    # -----------------------------------------------------------
    # READ FACULTY AND ASSIGN TO ROOMS
    # -----------------------------------------------------------
    faculty_df = pd.read_csv(faculty_file)
    faculty_list = faculty_df.iloc[:, 0].dropna().astype(str).tolist()
    faculty_cycle = cycle(faculty_list)

    room_faculty_map = {room["Room"]: next(faculty_cycle) for room in rooms}

    # -----------------------------------------------------------
    # SCHEDULING ALGORITHM WITH PROPER SLOT TRACKING
    # -----------------------------------------------------------
    schedule = []
    current_date = start_date
    slot_index = 0
    
    # Track room usage per slot: {(date, slot): {room_name: divisions_used}}
    slot_room_usage = {}
    
    # Track unscheduled courses and students
    unscheduled_courses = []
    
    for course_info in students_data:
        course = course_info["Course"]
        students = course_info["Students"]
        needed = len(students)
        year = f"{course_info['Year']} Year"
        
        course_scheduled = False
        
        # Keep trying slots until we find one that fits the ENTIRE course
        while not course_scheduled:
            slot_key = (current_date, SLOTS[slot_index])
            
            # Initialize room usage for this slot if not exists
            if slot_key not in slot_room_usage:
                slot_room_usage[slot_key] = {room["Room"]: 0 for room in rooms}
            
            room_usage = slot_room_usage[slot_key]
            
            # Calculate remaining capacity in current slot
            total_capacity = 0
            for room in rooms:
                divisions_used = room_usage[room["Room"]]
                divisions_available = courses_per_room - divisions_used
                
                if divisions_available > 0:
                    total_capacity += room["CapacityPerDivision"]
            
            # Check if entire course can fit in this slot
            if total_capacity < needed:
                # Not enough space - move to next slot
                slot_index += 1
                if slot_index >= len(SLOTS):
                    slot_index = 0
                    current_date = next_date(current_date)
                continue
            
            # ✅ Enough capacity! Now allocate the course
            assigned = 0
            course_allocations = []
            
            # Allocate students to rooms
            for room in rooms:
                if assigned >= needed:
                    break
                
                # Check how many divisions are still available in this room
                divisions_used = room_usage[room["Room"]]
                
                if divisions_used >= courses_per_room:
                    # This room is full for this slot
                    continue
                
                # Use the next available division
                capacity = room["CapacityPerDivision"]
                take = min(capacity, needed - assigned)
                
                roll_slice = students[assigned:assigned + take]
                faculty = room_faculty_map[room["Room"]]
                
                # Division number is divisions_used + 1
                division_number = divisions_used + 1
                
                # Format room display
                if room["IsSpecial"]:
                    room_display = f"{room['Room']} (Division {division_number}/{courses_per_room})"
                else:
                    room_display = f"{room['Room']} (Section {division_number}/{courses_per_room})"
                
                course_allocations.append({
                    "Date": current_date,
                    "Slot": SLOTS[slot_index],
                    "Room": room_display,
                    "Course": course,
                    "Year": year,
                    "Faculty": faculty,
                    "Student Count": len(roll_slice),
                    "Roll Numbers": ", ".join(roll_slice)
                })
                
                # Mark this division as used IN THE PERSISTENT TRACKER
                room_usage[room["Room"]] += 1
                assigned += take
            
            # Verify entire course was allocated
            if assigned == needed:
                schedule.extend(course_allocations)
                course_scheduled = True
            else:
                # Safety fallback - shouldn't happen
                slot_index += 1
                if slot_index >= len(SLOTS):
                    slot_index = 0
                    current_date = next_date(current_date)
        
        # If we couldn't schedule the course after trying all slots
        if not course_scheduled:
            unscheduled_courses.append({
                "Course": course,
                "Year": year,
                "Student Count": needed,
                "Students": students
            })

    # -----------------------------------------------------------
    # EXPORT TO EXCEL
    # -----------------------------------------------------------
    output_file = os.path.join(UPLOAD_FOLDER, "exam_schedule.xlsx")
    
    # Sort by Date, Slot, Room for better readability
    schedule_df = pd.DataFrame(schedule)
    schedule_df = schedule_df.sort_values(['Date', 'Slot', 'Room'])
    schedule_df.to_excel(output_file, index=False)
    
    # Create unscheduled report if needed
    if unscheduled_courses:
        unscheduled_file = os.path.join(UPLOAD_FOLDER, "unscheduled_courses.xlsx")
        unscheduled_df = pd.DataFrame(unscheduled_courses)
        unscheduled_df.to_excel(unscheduled_file, index=False)
    
    return {
        "schedule": schedule_df.to_dict('records'),
        "unscheduled": unscheduled_courses,
        "config": config,
        "output_file": output_file
    }

# Routes
@app.route('/')
def login():
    if 'username' in session:
        if session['username'] == 'admin':
            return redirect(url_for('index'))
        else:
            return redirect(url_for('view'))
    return render_template('login.html')

@app.route('/login', methods=['POST'])
def do_login():
    username = request.form['username']
    password = request.form['password']
    
    if username == ADMIN_CREDENTIALS['username'] and password == ADMIN_CREDENTIALS['password']:
        session['username'] = username
        session['role'] = 'admin'
        return redirect(url_for('index'))
    elif username == VIEWER_CREDENTIALS['username'] and password == VIEWER_CREDENTIALS['password']:
        session['username'] = username
        session['role'] = 'viewer'
        return redirect(url_for('view'))
    else:
        flash('Invalid credentials', 'error')
        return redirect(url_for('login'))

@app.route('/logout')
def logout():
    session.pop('username', None)
    session.pop('role', None)
    return redirect(url_for('login'))

@app.route('/index')
def index():
    if 'username' not in session or session['role'] != 'admin':
        return redirect(url_for('login'))
    return render_template('index.html')

@app.route('/edit')
def edit():
    if 'username' not in session or session['role'] != 'admin':
        return redirect(url_for('login'))
    
    # Check if schedule exists
    schedule_file = os.path.join(UPLOAD_FOLDER, "exam_schedule.xlsx")
    if os.path.exists(schedule_file):
        schedule_df = pd.read_excel(schedule_file)
        schedule = schedule_df.to_dict('records')
        return render_template('edit.html', schedule=schedule)
    else:
        flash('No schedule found. Please generate a schedule first.', 'error')
        return redirect(url_for('index'))

@app.route('/view')
def view():
    if 'username' not in session:
        return redirect(url_for('login'))
    
    # Check if schedule exists
    schedule_file = os.path.join(UPLOAD_FOLDER, "exam_schedule.xlsx")
    if os.path.exists(schedule_file):
        schedule_df = pd.read_excel(schedule_file)
        schedule = schedule_df.to_dict('records')
        return render_template('view.html', schedule=schedule)
    else:
        flash('No schedule available. Please contact the administrator.', 'error')
        return redirect(url_for('login'))

@app.route('/generate', methods=['POST'])
def generate():
    if 'username' not in session or session['role'] != 'admin':
        return jsonify({"success": False, "message": "Unauthorized"})
    
    # Check if files were uploaded
    if 'courses_file' not in request.files or 'rooms_file' not in request.files or 'faculty_file' not in request.files:
        return jsonify({"success": False, "message": "Missing required files"})
    
    courses_file = request.files['courses_file']
    rooms_file = request.files['rooms_file']
    faculty_file = request.files['faculty_file']
    
    start_date = request.form['start_date']
    courses_per_room = int(request.form['courses_per_room'])
    
    # Save uploaded files
    courses_filename = os.path.join(UPLOAD_FOLDER, secure_filename(courses_file.filename))
    rooms_filename = os.path.join(UPLOAD_FOLDER, secure_filename(rooms_file.filename))
    faculty_filename = os.path.join(UPLOAD_FOLDER, secure_filename(faculty_file.filename))
    
    courses_file.save(courses_filename)
    rooms_file.save(rooms_filename)
    faculty_file.save(faculty_filename)
    
    try:
        # Generate schedule
        result = generate_schedule(courses_filename, rooms_filename, faculty_filename, start_date, courses_per_room)
        
        return jsonify({
            "success": True,
            "message": "Schedule generated successfully",
            "schedule": result["schedule"],
            "unscheduled": result["unscheduled"],
            "config": result["config"]
        })
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})

@app.route('/download/<filename>')
def download(filename):
    if 'username' not in session:
        return redirect(url_for('login'))
    
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        flash('File not found', 'error')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)