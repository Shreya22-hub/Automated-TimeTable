import pandas as pd
from itertools import cycle
from datetime import datetime, timedelta
import os
import re
import json

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
    scheduled_students = set()
    total_students = 0
    
    for course_info in students_data:
        course = course_info["Course"]
        students = course_info["Students"]
        needed = len(students)
        year = f"{course_info['Year']} Year"
        
        # Track all students for later comparison
        total_students += len(students)
        for student in students:
            scheduled_students.add(student)

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
    # CHECK FOR UNSCHEDULED COURSES/STUDENTS
    # -----------------------------------------------------------
    print(f"\n✓ Schedule verification:")
    print(f"  Total courses: {len(students_data)}")
    print(f"  Scheduled courses: {len(students_data) - len(unscheduled_courses)}")
    print(f"  Unscheduled courses: {len(unscheduled_courses)}")
    
    if len(unscheduled_courses) == 0:
        print(f"  ✓ All courses scheduled successfully!")
    else:
        print(f"  ❌ The following courses could not be scheduled:")
        for course in unscheduled_courses:
            print(f"    - {course['Course']} ({course['Year']}): {course['Student Count']} students")
    
    # Create unscheduled report if needed
    if unscheduled_courses:
        unscheduled_file = os.path.join(config_dir, "unscheduled_courses.xlsx")
        unscheduled_df = pd.DataFrame(unscheduled_courses)
        unscheduled_df.to_excel(unscheduled_file, index=False)
        print(f"  ✓ Unscheduled courses report saved to: {unscheduled_file}")

    # -----------------------------------------------------------
    # EXPORT TO EXCEL
    # -----------------------------------------------------------
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    # Sort by Date, Slot, Room for better readability
    schedule_df = pd.DataFrame(schedule)
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
        status = "✓" if count == courses_per_room else "❌"
        if count != courses_per_room:
            violations.append(f"{date} {slot} {room}: {count} courses (expected {courses_per_room})")
            print(f"    {status} {date} {slot} {room}: {count} courses")
    
    if not violations:
        print(f"    ✓ All rooms have exactly {courses_per_room} courses per slot!")
    else:
        print(f"\n  ❌ Found {len(violations)} violations:")
        for v in violations:
            print(f"    - {v}")
    
    return schedule