import pandas as pd
from itertools import cycle
from datetime import datetime, timedelta
import os
import re
import json
import argparse
import sys

def generate_schedule(courses_file, rooms_file, faculty_file, output_file, start_date, courses_per_room=2):
    """
    Generate exam schedule with dynamic room divisions
    
    Logic:
    - Regular rooms: Each course gets FULL room capacity
    - Special rooms (C403-C408): Room capacity divided by courses_per_room
    - Entire course must fit in ONE slot (never split across slots)
    - Each room can accommodate exactly courses_per_room courses simultaneously
    
    Args:
        courses_per_room: Number of courses to be conducted simultaneously in each room
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
    config_dir = os.path.dirname(output_file)
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
    
    print(f"✓ Configuration file created: {config_file}")
    print(f"  Max students per slot: {config['max_students_per_slot']}")
    print(f"  Courses per room: {config['courses_per_room']}")

    # -----------------------------------------------------------
    # READ COURSES FILE (multiple sheets: 1st year, 2nd year etc)
    # -----------------------------------------------------------
    xls = pd.ExcelFile(courses_file)
    students_data = []

    for year_index, sheet_name in enumerate(xls.sheet_names, start=1):
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        except Exception:
            # If sheet cannot be read, skip it silently
            continue
            
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)

        # --- NEW LOGIC: Skip if sheet is empty ---
        if df.empty:
            continue
        # ----------------------------------------

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
    
    # --- VERIFICATION SETUP ---
    # Create a master set of all students from input for final verification
    all_students_from_input = set()
    for course_info in students_data:
        all_students_from_input.update(course_info["Students"])
    
    # Create a set to track students as they are actually scheduled
    actually_scheduled_students = set()
    
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
                
                divisions_used = room_usage[room["Room"]]
                if divisions_used >= courses_per_room:
                    continue
                
                capacity = room["CapacityPerDivision"]
                take = min(capacity, needed - assigned)
                
                roll_slice = students[assigned:assigned + take]
                faculty = room_faculty_map[room["Room"]]
                
                # --- VERIFICATION STEP ---
                # Add these actually assigned students to our tracking set
                actually_scheduled_students.update(roll_slice)
                
                division_number = divisions_used + 1
                
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
                
                room_usage[room["Room"]] += 1
                assigned += take
            
            # Verify entire course was allocated
            if assigned == needed:
                schedule.extend(course_allocations)
                course_scheduled = True
            else:
                # This should not happen if capacity check is correct, but as a fallback:
                slot_index += 1
                if slot_index >= len(SLOTS):
                    slot_index = 0
                    current_date = next_date(current_date)

    # -----------------------------------------------------------
    # FINAL VERIFICATION REPORT
    # -----------------------------------------------------------
    print("\n" + "="*50)
    print("          FINAL VERIFICATION REPORT")
    print("="*50)

    # 1. Course Verification
    schedule_df = pd.DataFrame(schedule)
    scheduled_course_names = set(schedule_df['Course'].unique()) if not schedule_df.empty else set()
    all_course_names = {item['Course'] for item in students_data}
    truly_unscheduled_courses = all_course_names - scheduled_course_names

    print("\n--- COURSES ---")
    print(f"Total courses in input: {len(all_course_names)}")
    print(f"Courses successfully scheduled: {len(scheduled_course_names)}")
    print(f"Courses left unscheduled: {len(truly_unscheduled_courses)}")

    if truly_unscheduled_courses:
        print("\n❌ ERROR: The following courses were not scheduled:")
        for course in truly_unscheduled_courses:
            print(f"  - {course}")
        
        # Save a detailed report for unscheduled courses
        unscheduled_courses_report = []
        for course_info in students_data:
            if course_info['Course'] in truly_unscheduled_courses:
                unscheduled_courses_report.append({
                    "Course": course_info['Course'],
                    "Year": f"{course_info['Year']} Year",
                    "Student Count": len(course_info['Students']),
                    "Students": ", ".join(course_info['Students'])
                })
        if unscheduled_courses_report:
            report_df = pd.DataFrame(unscheduled_courses_report)
            report_file = os.path.join(config_dir, "unscheduled_courses_report.xlsx")
            report_df.to_excel(report_file, index=False)
            print(f"\n  📄 Detailed report saved to: {report_file}")
    else:
        print("\n✅ SUCCESS: All courses from the input have been scheduled.")

    # 2. Student Verification
    unscheduled_students = all_students_from_input - actually_scheduled_students

    print("\n--- STUDENTS ---")
    print(f"Total unique students in input: {len(all_students_from_input)}")
    print(f"Students successfully assigned: {len(actually_scheduled_students)}")
    print(f"Students left unassigned: {len(unscheduled_students)}")

    if unscheduled_students:
        print("\n❌ ERROR: The following students were not assigned to any exam slot:")
        # Create a map to find which course each student belongs to
        student_to_course_map = {}
        for course_info in students_data:
            for student in course_info['Students']:
                student_to_course_map[student] = course_info['Course']

        for student in sorted(list(unscheduled_students)):
            course = student_to_course_map.get(student, "Unknown Course")
            print(f"  - {student} (from course: {course})")

        # Save a report for unscheduled students
        unscheduled_students_report = []
        for student in sorted(list(unscheduled_students)):
            unscheduled_students_report.append({
                "Roll Number": student,
                "Course": student_to_course_map.get(student, "Unknown Course")
            })
        if unscheduled_students_report:
            report_df = pd.DataFrame(unscheduled_students_report)
            report_file = os.path.join(config_dir, "unscheduled_students_report.xlsx")
            report_df.to_excel(report_file, index=False)
            print(f"\n  📄 Detailed report saved to: {report_file}")
    else:
        print("\n✅ SUCCESS: All students from the input have been assigned to an exam slot.")
    
    print("="*50)

    # -----------------------------------------------------------
    # EXPORT TO EXCEL
    # -----------------------------------------------------------
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    # Sort by Date, Slot, Room for better readability
    if not schedule_df.empty:
        schedule_df = schedule_df.sort_values(['Date', 'Slot', 'Room'])
        schedule_df.to_excel(output_file, index=False)
        
        print(f"\n✓ Schedule generated successfully!")
        print(f"  Total entries: {len(schedule_df)}")
        print(f"  Unique courses: {schedule_df['Course'].nunique()}")
        print(f"  Date range: {schedule_df['Date'].min()} to {schedule_df['Date'].max()}")
        
        # Show room usage verification
        print(f"\n  Verifying room usage per slot:")
        schedule_df['Base_Room'] = schedule_df['Room'].apply(lambda x: x.split(' (')[0])
        grouped = schedule_df.groupby(['Date', 'Slot', 'Base_Room']).size()
        
        violations = []
        for (date, slot, room), count in grouped.items():
            if count != courses_per_room:
                violations.append(f"{date} {slot} {room}: {count} courses (expected {courses_per_room})")
        
        if not violations:
            print(f"    ✓ All rooms have exactly {courses_per_room} courses per slot!")
        else:
            print(f"\n  ❌ Found {len(violations)} room usage violations:")
            for v in violations:
                print(f"    - {v}")
    else:
        print("\n⚠️ No schedule generated (no valid courses found in input).")
        # Create an empty file to prevent errors downstream
        pd.DataFrame().to_excel(output_file)
    
    return schedule

def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Generate exam schedule from input files.')
    parser.add_argument('date', type=str, help='Start date for the schedule (YYYY-MM-DD format)')
    parser.add_argument('--courses_per_room', type=int, default=2, 
                        help='Number of courses per room (default: 2)')
    args = parser.parse_args()
    
    # Validate date format
    try:
        datetime.strptime(args.date, "%Y-%m-%d")
    except ValueError:
        print("Error: Date must be in YYYY-MM-DD format")
        sys.exit(1)
    
    # Get the directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Define input and output folders relative to the script's location
    input_folder = os.path.join(script_dir, "uploadsElectiveExam")
    output_folder = os.path.join(script_dir, "output")
    
    # Check if input folder exists
    if not os.path.exists(input_folder):
        print(f"Error: Input folder '{input_folder}' not found")
        sys.exit(1)
    
    # Define input files
    courses_file = os.path.join(input_folder, "courses.xlsx")
    rooms_file = os.path.join(input_folder, "rooms.xlsx")
    faculty_file = os.path.join(input_folder, "faculty.csv")
    
    # Check if input files exist
    for file_path in [courses_file, rooms_file, faculty_file]:
        if not os.path.exists(file_path):
            print(f"Error: Input file '{file_path}' not found")
            sys.exit(1)
    
    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)
    output_file = os.path.join(output_folder, f"exam_schedule_{args.date}.xlsx")
    
    # Generate the schedule
    print(f"Generating exam schedule starting from {args.date}...")
    print(f"Using input files from: {input_folder}")
    print(f"Output will be saved to: {output_folder}")
    
    generate_schedule(
        courses_file=courses_file,
        rooms_file=rooms_file,
        faculty_file=faculty_file,
        output_file=output_file,
        start_date=args.date,
        courses_per_room=args.courses_per_room
    )

if __name__ == "__main__":
    main()