import os
import json
import pandas as pd
import math
import string
from datetime import datetime, timedelta
from itertools import cycle
from flask import Flask, render_template, request, send_file, jsonify, session
from werkzeug.utils import secure_filename

app = Flask(__name__, template_folder="app3")
app.secret_key = 'exam_scheduler_secret_key_v2'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# -----------------------------------------------------------
# SEATING LOGIC (Mixed: Bench = All Groups, Individual = Alternating)
# -----------------------------------------------------------

def get_seating_layout(rows, cols, seat_type, num_groups, room_name="Standard"):
    """
    Generates a 2D grid.
    - If BENCH: Every bench contains all groups ("A1 B1 C1").
    - If INDIVIDUAL: Seats alternate groups ("A1", "B1", "A2", "B2").
    """
    
    num_groups = min(max(1, num_groups), 4) # Limit to 4 groups A,B,C,D
    group_chars = list(string.ascii_uppercase[:num_groups])
    seat_type = seat_type.lower().strip()
    
    grid = []
    label_counters = {char: 1 for char in group_chars}
    
    if seat_type == "bench":
        # BENCH MODE: Each cell gets all group labels
        for r in range(rows):
            row_data = []
            for c in range(cols):
                sub_labels = []
                for group_char in group_chars:
                    label = f"{group_char}{label_counters[group_char]}"
                    sub_labels.append(label)
                    label_counters[group_char] += 1
                row_data.append(" ".join(sub_labels))
            grid.append(row_data)
            
    else: 
        # INDIVIDUAL MODE: Cells alternate groups
        for r in range(rows):
            row_data = []
            for c in range(cols):
                # Determine group based on linear position in grid
                idx = r * cols + c
                group_char = group_chars[idx % num_groups]
                
                label = f"{group_char}{label_counters[group_char]}"
                label_counters[group_char] += 1
                row_data.append(label)
            grid.append(row_data)

    return {
        "grid": grid,
        "type": seat_type, 
        "rows": rows, "cols": cols,
        "room_name": room_name
    }

# -----------------------------------------------------------
# SCHEDULING LOGIC
# -----------------------------------------------------------

SLOTS = ["Morning", "Evening"]

def next_date(date_str, days=1):
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return (dt + timedelta(days=days)).strftime("%Y-%m-%d")

def generate_schedule_logic(courses_path, rooms_path, faculty_path, output_folder, start_date, courses_per_room):
    
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    excel_filename = f"exam_schedule_{start_date}_{timestamp}.xlsx"
    excel_path = os.path.join(output_folder, excel_filename)
    
    config_filename = "configurations.json"
    config_path = os.path.join(output_folder, config_filename)
    
    snapshot_filename = f"verification_snapshot_{timestamp}.json"
    snapshot_path = os.path.join(output_folder, snapshot_filename)

    try:
        rooms_df = pd.read_excel(rooms_path)
    except Exception as e:
        raise ValueError(f"Error reading Rooms file: {str(e)}")

    total_capacity = 0
    room_map = {} 
    room_layouts_map = {} 

    for _, r in rooms_df.iterrows():
        try:
            cap = int(r["Capacity"])
        except ValueError:
            continue
        room_name = str(r["Room"])
        
        parsed_layout = {"rows": 0, "cols": 0, "type": "individual"}
        
        if "layout" in r and pd.notna(r["layout"]):
            try:
                parts = [p.strip() for p in str(r["layout"]).split(',')]
                if len(parts) >= 3:
                    parsed_layout["rows"] = int(parts[0])
                    parsed_layout["cols"] = int(parts[1])
                    parsed_layout["type"] = parts[2].lower().strip()
            except Exception:
                pass 
        
        room_layouts_map[room_name] = parsed_layout
        room_map[room_name] = cap
        total_capacity += cap * courses_per_room
    
    room_layouts_file = os.path.join(output_folder, "room_layouts.json")
    with open(room_layouts_file, 'w') as f:
        json.dump(room_layouts_map, f, indent=4)
    
    config = {
        "max_students_per_slot": total_capacity,
        "courses_per_room": courses_per_room,
        "start_date": start_date,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "schedule_file": excel_filename
    }
    with open(config_path, 'w') as f:
        json.dump(config, f, indent=4)

    try:
        xls = pd.ExcelFile(courses_path)
    except Exception as e:
        raise ValueError(f"Error reading Courses file: {str(e)}")

    students_data = []
    input_courses_set = set()
    input_students_set = set()

    for year_index, sheet_name in enumerate(xls.sheet_names, start=1):
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        except Exception:
            continue
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        if df.empty: continue

        for col in range(df.shape[1]):
            course = df.iloc[0, col]
            students = df.iloc[1:, col].dropna().astype(str).tolist()

            if str(course).strip():
                course_name = str(course).strip()
                students_data.append({
                    "Year": year_index,
                    "Course": course_name,
                    "Students": students
                })
                input_courses_set.add(course_name)
                input_students_set.update(students)

    if not students_data:
        raise ValueError("No valid student/course data found.")

    verification_snapshot = {
        "total_input_courses": len(input_courses_set),
        "total_input_students": len(input_students_set)
    }
    with open(snapshot_path, 'w') as f:
        json.dump(verification_snapshot, f, indent=4)

    rooms = []
    for _, r in rooms_df.iterrows():
        try:
            cap = int(r["Capacity"])
        except ValueError: continue
        room_name = str(r["Room"])
        rooms.append({
            "Room": room_name, "TotalCapacity": cap,
            "CapacityPerDivision": cap, 
            "IsSpecial": False 
        })

    try:
        faculty_df = pd.read_csv(faculty_path)
        faculty_list = faculty_df.iloc[:, 0].dropna().astype(str).tolist()
    except Exception as e:
        raise ValueError(f"Error reading Faculty: {str(e)}")
    if not faculty_list: raise ValueError("No faculty data.")
    
    faculty_cycle = cycle(faculty_list)
    room_faculty_map = {room["Room"]: next(faculty_cycle) for room in rooms}

    schedule = []
    current_date = start_date
    slot_index = 0
    slot_room_usage = {}
    
    for course_info in students_data:
        course = course_info["Course"]
        students = course_info["Students"]
        needed = len(students)
        year = f"{course_info['Year']} Year"
        course_scheduled = False
        
        while not course_scheduled:
            slot_key = (current_date, SLOTS[slot_index])
            if slot_key not in slot_room_usage:
                slot_room_usage[slot_key] = {room["Room"]: 0 for room in rooms}
            room_usage = slot_room_usage[slot_key]
            
            total_capacity = 0
            for room in rooms:
                divisions_used = room_usage[room["Room"]]
                divisions_available = courses_per_room - divisions_used
                if divisions_available > 0:
                    total_capacity += room["CapacityPerDivision"]
            
            if total_capacity < needed:
                slot_index += 1
                if slot_index >= len(SLOTS):
                    slot_index = 0
                    current_date = next_date(current_date)
                continue
            
            assigned = 0
            course_allocations = []
            for room in rooms:
                if assigned >= needed: break
                divisions_used = room_usage[room["Room"]]
                if divisions_used >= courses_per_room: continue
                
                capacity = room["CapacityPerDivision"]
                take = min(capacity, needed - assigned)
                roll_slice = students[assigned:assigned + take]
                faculty = room_faculty_map[room["Room"]]
                
                division_number = divisions_used + 1
                room_display = f"{room['Room']} (Sec {division_number}/{courses_per_room})"
                
                base_room_name = room["Room"]
                total_room_cap = room["TotalCapacity"]
                
                course_allocations.append({
                    "Date": current_date, "Slot": SLOTS[slot_index],
                    "Room": room_display, "BaseRoom": base_room_name, "RoomCapacity": total_room_cap,
                    "Course": course, "Year": year,
                    "Faculty": faculty, "Student Count": len(roll_slice),
                    "Roll Numbers": ", ".join(roll_slice)
                })
                room_usage[room["Room"]] += 1
                assigned += take
            
            if assigned == needed:
                schedule.extend(course_allocations)
                course_scheduled = True
            else:
                slot_index += 1
                if slot_index >= len(SLOTS):
                    slot_index = 0
                    current_date = next_date(current_date)

    if not schedule: raise ValueError("No schedule generated.")

    schedule_df = pd.DataFrame(schedule)
    schedule_df = schedule_df.sort_values(['Date', 'Slot', 'Room'])
    schedule_df.to_excel(excel_path, index=False)
    
    return excel_path, snapshot_path

# -----------------------------------------------------------
# ROUTES
# -----------------------------------------------------------

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    if 'courses' not in request.files or 'rooms' not in request.files or 'faculty' not in request.files:
        return jsonify({"error": "Missing files."}), 400
    
    files = {k: request.files[k] for k in ['courses', 'rooms', 'faculty']}
    if any(f.filename == '' for f in files.values()):
        return jsonify({"error": "No file selected."}), 400

    start_date = request.form.get('start_date')
    if not start_date: return jsonify({"error": "Date required."}), 400

    try:
        courses_per_room = int(request.form.get('courses_per_room', 2))
        datetime.strptime(start_date, "%Y-%m-%d")
    except ValueError:
        return jsonify({"error": "Invalid format."}), 400

    base_path = app.config['UPLOAD_FOLDER']
    paths = {k: os.path.join(base_path, secure_filename(v.filename)) for k, v in files.items()}

    try:
        for k, f in files.items(): f.save(paths[k])
        
        result_excel, result_snapshot = generate_schedule_logic(
            paths['courses'], paths['rooms'], paths['faculty'], 
            base_path, start_date, courses_per_room
        )
        
        session['last_schedule'] = os.path.basename(result_excel)
        session['last_snapshot'] = os.path.basename(result_snapshot)
        
        return jsonify({"success": True, "message": "Schedule generated successfully"})
        
    except ValueError as ve:
        return jsonify({"error": str(ve)}), 400
    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

@app.route('/view', methods=['GET'])
def view_schedule():
    if 'last_schedule' not in session or 'last_snapshot' not in session:
        return jsonify({"error": "No active schedule. Please generate a schedule first."}), 400
    
    excel_name = session['last_schedule']
    snapshot_name = session['last_snapshot']
    
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_name)
    snapshot_path = os.path.join(app.config['UPLOAD_FOLDER'], snapshot_name)
    
    if not os.path.exists(excel_path):
        return jsonify({"error": "Excel file missing from server."}), 400
    
    try:
        df = pd.read_excel(excel_path)
        schedule_data = df.to_dict(orient='records')
    except Exception as e:
        return jsonify({"error": f"Error reading Excel: {str(e)}"}), 500

    config_data = {}
    config_path = os.path.join(app.config['UPLOAD_FOLDER'], "configurations.json")
    if os.path.exists(config_path):
        with open(config_path, 'r') as f:
            config_data = json.load(f)

    verification = {
        "status": "Unknown", "input_courses": 0, "scheduled_courses": 0,
        "input_students": 0, "scheduled_students": 0, "message": ""
    }

    if os.path.exists(snapshot_path):
        with open(snapshot_path, 'r') as f:
            snapshot = json.load(f)
        verification['input_courses'] = snapshot['total_input_courses']
        verification['input_students'] = snapshot['total_input_students']

        scheduled_course_names = set()
        scheduled_student_rolls = set()
        for row in schedule_data:
            if 'Course' in row: scheduled_course_names.add(row['Course'])
            if 'Roll Numbers' in row:
                rolls = [str(r).strip() for r in str(row['Roll Numbers']).split(',') if str(r).strip()]
                scheduled_student_rolls.update(rolls)
        
        verification['scheduled_courses'] = len(scheduled_course_names)
        verification['scheduled_students'] = len(scheduled_student_rolls)
        
        if verification['scheduled_courses'] == verification['input_courses'] and \
           verification['scheduled_students'] == verification['input_students']:
            verification['status'] = "Success"
        else:
            verification['status'] = "Partial"
            missing = verification['input_students'] - verification['scheduled_students']
            if missing > 0:
                 verification['message'] = f"Warning: {missing} students missing."
    
    return jsonify({
        "schedule": schedule_data, 
        "config": config_data, 
        "verification": verification
    })

@app.route('/api/seating', methods=['POST'])
def api_seating():
    data = request.json
    groups = int(data.get('groups', 2))
    room_name = data.get('room_name', "Standard")
    
    layout_path = os.path.join(app.config['UPLOAD_FOLDER'], "room_layouts.json")
    rows, cols, seat_type = 10, 10, "individual"
    
    if os.path.exists(layout_path):
        with open(layout_path, 'r') as f:
            layouts = json.load(f)
            if room_name in layouts:
                info = layouts[room_name]
                if info.get('rows', 0) > 0:
                    rows = info['rows']
                    cols = info['cols']
                    seat_type = info['type']
                else:
                    capacity = int(data.get('capacity', 60))
                    cols = int(math.ceil(math.sqrt(capacity)))
                    rows = int(math.ceil(capacity / cols))
                    seat_type = "individual"
    else:
        capacity = int(data.get('capacity', 60))
        cols = int(math.ceil(math.sqrt(capacity)))
        rows = int(math.ceil(capacity / cols))
        seat_type = "individual"

    layout = get_seating_layout(rows, cols, seat_type, groups, room_name)
    return jsonify(layout)

@app.route('/save_changes', methods=['POST'])
def save_changes():
    if 'last_schedule' not in session:
        return jsonify({"error": "Session expired."}), 400
    
    filename = session['last_schedule']
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    try:
        updated_rows = request.json.get('rows', [])
        if not updated_rows: return jsonify({"error": "No data"}), 400
        df_new = pd.DataFrame(updated_rows)
        df_new.to_excel(file_path, index=False)
        return jsonify({"success": True})
    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename))

if __name__ == '__main__':
    app.run(debug=True, port=5002) 