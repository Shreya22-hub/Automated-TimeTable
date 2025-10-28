import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font
import math
from datetime import datetime, timedelta
import copy
from collections import defaultdict

def generate_timetable(start_date, end_date, branch_slot_allocation, max_credits_per_day=5):
    """
    Generate exam timetable with given parameters
    
    Args:
        start_date (str): Start date in 'YYYY-MM-DD' format
        end_date (str): End date in 'YYYY-MM-DD' format
        branch_slot_allocation (dict): Branch allocation per year and slot
        max_credits_per_day (int): Maximum credits allowed per day per branch (default: 5)
    
    Returns:
        str: Path to generated Excel file
    """
    
    # Convert string dates to datetime
    start_date = datetime.strptime(start_date, "%Y-%m-%d")
    end_date = datetime.strptime(end_date, "%Y-%m-%d")
    
    # Public holidays in India (example, update as needed)
    public_holidays = [
        "2025-01-26", "2025-03-14", "2025-03-31", "2025-04-06", "2025-04-18",
        "2025-05-12", "2025-07-06", "2025-08-15", "2025-10-02", "2025-10-14", "2025-10-20"
    ]
    public_holidays = [datetime.strptime(d, "%Y-%m-%d") for d in public_holidays]
    
    # Generate list of valid exam dates (exclude Sundays & holidays)
    exam_dates = []
    current_date = start_date
    while current_date <= end_date:
        if current_date.weekday() != 6 and current_date not in public_holidays:
            exam_dates.append(current_date)
        current_date += timedelta(days=1)
    
    # -----------------------------
    # Step 1: Read Inputs
    # -----------------------------
    df_strength = pd.read_excel("BranchStrength.xlsx")
    df_courses = pd.read_excel("CoursesPerYear.xlsx")
    df_common = pd.read_excel("CommonCourse.xlsx")
    df_settings = pd.read_excel("Settings.xlsx")
    df_faculty = pd.read_csv("FACULTY.csv")
    df_room = pd.read_excel("rooms.xlsx")
    course_book = pd.read_excel("courselist.xlsx", sheet_name=None)
    
    # -----------------------------
    # Step 2: Clean Inputs
    # -----------------------------
    df_strength["Year"] = df_strength["Year"].astype(str).str.strip().str.title()
    df_strength["Branch"] = df_strength["Branch"].astype(str).str.strip().str.upper()
    df_courses["Year"] = df_courses["Year"].astype(str).str.strip().str.title()
    faculty_list = df_faculty["Name"].astype(str).tolist()
    
    # -----------------------------
    # Step 2: Map Sheets to Years Automatically
    # -----------------------------
    sheet_year_map = {}
    year_names = ["1St Year", "2Nd Year", "3Rd Year", "4Th Year", "5Th Year"]
    
    for i, sheet_name in enumerate(course_book.keys()):
        if i < len(year_names):
            sheet_year_map[sheet_name] = year_names[i]
        else:
            sheet_year_map[sheet_name] = f"{i+1}Th Year"
    
    # -----------------------------
    # Step 3: Prepare Branch Courses
    # -----------------------------
    branch_courses = {}
    for sheet_name, df in course_book.items():
        year = sheet_year_map[sheet_name]
        branch_courses[year] = {}
        df.columns = [str(c).strip().upper() for c in df.columns]
        for branch in ["CSE", "DSAI", "ECE"]:
            branch_courses[year][branch] = []
            for val in df.get(branch, pd.Series()).dropna().astype(str):
                try:
                    course_code, credits = val.split(",")
                    branch_courses[year][branch].append({
                        "course_code": course_code.strip(),
                        "credits": int(credits.strip())
                    })
                except:
                    print(f"Skipping invalid course entry: {val}")
    
    # -----------------------------
    # Step 3: Extract Settings
    # -----------------------------
    settings = dict(zip(df_settings["SettingName"], df_settings["Value"]))
    credits_per_course = int(settings["CreditsPerCourse"])
    max_students_per_slot = int(settings["MaxStudentsPerSlot"])
    max_courses_per_slot = int(settings["MaxCoursesPerSlot"])
    total_rooms = int(settings["TotalRooms"])
    room_capacity_per_course = int(settings["RoomCapacityPerCourse"])
    
    # -----------------------------
    # Step 4: Python Structures
    # -----------------------------
    branches = df_strength["Branch"].unique().tolist()
    years = df_strength["Year"].unique().tolist()
    
    branch_strength = {
        year: dict(zip(
            df_strength[df_strength["Year"]==year]["Branch"],
            df_strength[df_strength["Year"]==year]["Strength"]
        ))
        for year in years
    }
    
    courses_per_year = dict(zip(df_courses["Year"], df_courses["CoursesPerYear"]))
    
    common_course = {
        "course_code": df_common.loc[0, "CourseCode"],
        "credits": int(df_common.loc[0, "Credits"])
    }
    
    # -----------------------------
    # Step 5: Initialize Schedule
    # -----------------------------
    days = [d.strftime("%Y-%m-%d") for d in exam_dates]
    slots = ["Morning", "Evening"]
    schedule = {day: {slot: {year: [] for year in years} for slot in slots} for day in days}
    day_year_credits = {day: {year: {branch: 0 for branch in branches} for year in years} for day in days}
    day_slot_total_students = {day: {slot:0 for slot in slots} for day in days}
    day_slot_total_courses = {day: {slot:0 for slot in slots} for day in days}
    remaining_courses = {year: {branch: courses_per_year[year] for branch in branches} for year in years}
    
    def normalize_year(year_text):
        text = str(year_text).lower().replace(" ", "")
        if "1st" in text or "first" in text or text == "1":
            return "1St Year"
        elif "2nd" in text or "second" in text or text == "2":
            return "2Nd Year"
        elif "3rd" in text or "third" in text or text == "3":
            return "3Rd Year"
        else:
            return year_text
    
    slot_max_students = total_rooms * 48
    common_course_map = {}
    for _, row in df_common.iterrows():
        year_norm = normalize_year(str(row['Year']))
        branches_cell = row['Branches']
        if pd.isna(branches_cell):
            branches_for_course = []
        else:
            branches_for_course = [b.strip() for b in str(branches_cell).split(",")]
        
        common_course_map[row['CourseCode']] = {
            "credits": row['Credits'],
            "Year": year_norm,
            "Branches": branches_for_course
        }
    
    common_assigned = {code: False for code in common_course_map}
    
    # -----------------------------
    # Step 6b: Assign Common Courses
    # -----------------------------
    for day in days:
        for course_code, info in common_course_map.items():
            year_norm = info['Year']
            branches_to_block = info['Branches']
            if common_assigned[course_code]:
                continue
            
            for slot in slots:
                if year_norm not in schedule[day][slot]:
                    schedule[day][slot][year_norm] = []
                if year_norm not in day_year_credits[day]:
                    day_year_credits[day][year_norm] = {b: 0 for b in branches}
                
                conflict = False
                for branch in branches_to_block:
                    if day_year_credits[day][year_norm].get(branch, 0) + info['credits'] > max_credits_per_day:
                        conflict = True
                        break
                    if any(c.get('branch') == branch for c in schedule[day][slot][year_norm]):
                        conflict = True
                        break
                if conflict:
                    continue
                
                total_students = sum(branch_strength[year_norm][b] for b in branches_to_block)
                if day_slot_total_students[day][slot] + total_students > slot_max_students:
                    continue
                
                for branch in branches_to_block:
                    schedule[day][slot][year_norm].append({
                        "course_code": course_code,
                        "credits": info['credits'],
                        "branch": branch,
                        "year": year_norm,
                        "students": branch_strength[year_norm][branch],
                        "type": "Common"
                    })
                    day_year_credits[day][year_norm][branch] += info['credits']
                    day_slot_total_students[day][slot] += branch_strength[year_norm][branch]
                
                day_slot_total_courses[day][slot] += 1
                common_assigned[course_code] = True
                break
    
    # -----------------------------
    # Step 7: Assign Main Courses
    # -----------------------------
    default_branch_slot_allocation = branch_slot_allocation
    
    # Validate against slot capacity
    for year in default_branch_slot_allocation:
        normalized_year = normalize_year(year)
        for slot in default_branch_slot_allocation[year]:
            branches_in_slot = default_branch_slot_allocation[year][slot]
            if normalized_year not in branch_strength:
                continue
            total_strength = sum(branch_strength[normalized_year][b] for b in branches_in_slot if b in branch_strength[normalized_year])
            if total_strength > slot_max_students:
                print(f"⚠️ Warning: {normalized_year} {slot} total ({total_strength}) exceeds slot capacity ({slot_max_students}).")
    
    branch_slot_allocation_day = {day: copy.deepcopy(default_branch_slot_allocation) for day in days}
    
    for day in days:
        for slot in slots:
            for year in years:
                normalized_year = normalize_year(year)
                allowed_branches = branch_slot_allocation_day[day][normalized_year][slot]
                
                running_total = day_slot_total_students[day][slot]
                
                final_branches = []
                for b in allowed_branches:
                    b_strength = branch_strength[normalized_year].get(b, 0)
                    if running_total + b_strength <= slot_max_students:
                        final_branches.append(b)
                        running_total += b_strength
                    else:
                        print(f"Skipping branch {b} for {normalized_year} {slot} on {day} — would exceed slot capacity")
                
                branch_slot_allocation_day[day][normalized_year][slot] = final_branches
    
    def total_remaining():
        return sum(remaining_courses[y][b] for y in years for b in branches)
    
    # Find ENV day
    env_day = None
    env_code = common_course["course_code"]
    for d in days:
        if any(
            c.get("course_code") == env_code
            for s in slots
            for y in years
            for c in schedule[d][s][y]
        ):
            env_day = d
            break
    
    after_env_days = days[days.index(env_day)+1:] if env_day in days else days[:]
    
    for day in days:
        day_allocation = branch_slot_allocation_day.get(day, default_branch_slot_allocation)
        
        for slot in slots:
            blocked_branches = set()
            for year in years:
                for c in schedule[day][slot][year]:
                    if c.get("type") == "Common":
                        if isinstance(c["branch"], str):
                            for b in c["branch"].split(","):
                                blocked_branches.add(b.strip())
            
            for year in years:
                normalized_year = normalize_year(year)
                
                if any(c.get("type") == "Common" for c in schedule[day][slot][year]):
                    continue
                
                allowed_branches = day_allocation.get(normalized_year, {}).get(slot, [])
                
                candidates = [
                    b for b in allowed_branches
                    if remaining_courses[year].get(b, 0) > 0
                    and b not in blocked_branches
                ]
                
                candidates.sort(key=lambda b: branch_strength[year][b], reverse=True)
                
                for b in candidates:
                    if day_year_credits[day][year][b] + credits_per_course > max_credits_per_day:
                        continue
                    if day_slot_total_students[day][slot] + branch_strength[year][b] > slot_max_students:
                        continue
                    
                    course_index = courses_per_year[year] - remaining_courses[year][b]
                    year_key = normalize_year(year)
                    if course_index < len(branch_courses.get(year_key, {}).get(b, [])):
                        course_info = branch_courses[year_key][b][course_index]
                    else:
                        course_info = {"course_code": f"{b}{year[0]}X", "credits": credits_per_course}
                    
                    schedule[day][slot][year].append({
                        "course_code": course_info["course_code"],
                        "credits": course_info["credits"],
                        "branch": b,
                        "year": year,
                        "students": branch_strength[year][b],
                        "type": "Main"
                    })
                    
                    remaining_courses[year][b] -= 1
                    day_year_credits[day][year][b] += course_info["credits"]
                    day_slot_total_students[day][slot] += branch_strength[year][b]
                    day_slot_total_courses[day][slot] += 1
            
            # Fill empty slots after ENV day
            if day in after_env_days and total_remaining() > 0:
                slot_has_main = any(
                    len([c for c in schedule[day][slot][y] if c.get("type") != "Common"]) > 0
                    for y in years
                )
                if not slot_has_main:
                    placed = False
                    year_order = sorted(
                        years, key=lambda Y: sum(remaining_courses[Y][b] for b in branches), reverse=True
                    )
                    for year in year_order:
                        normalized_year = normalize_year(year)
                        if any(c.get("type") == "Common" for c in schedule[day][slot][year]):
                            continue
                        allowed_branches = day_allocation.get(normalized_year, {}).get(slot, [])
                        for b in allowed_branches:
                            if remaining_courses[year].get(b, 0) <= 0:
                                continue
                            if day_year_credits[day][year][b] + credits_per_course > max_credits_per_day:
                                continue
                            if day_slot_total_students[day][slot] + branch_strength[year][b] > slot_max_students:
                                continue
                            
                            course_index = courses_per_year[year] - remaining_courses[year][b]
                            if course_index < len(branch_courses.get(normalized_year, {}).get(b, [])):
                                course_info = branch_courses[normalized_year][b][course_index]
                            else:
                                course_info = {"course_code": f"{b}{year[0]}X", "credits": credits_per_course}
                            
                            schedule[day][slot][year].append({
                                "course_code": course_info["course_code"],
                                "credits": course_info["credits"],
                                "branch": b,
                                "year": year,
                                "students": branch_strength[year][b],
                                "type": "Main"
                            })
                            
                            remaining_courses[year][b] -= 1
                            day_year_credits[day][year][b] += course_info["credits"]
                            day_slot_total_students[day][slot] += branch_strength[year][b]
                            day_slot_total_courses[day][slot] += 1
                            placed = True
                            break
                        if placed:
                            break
    
    # -----------------------------
    # Step 8: Prepare Exam Schedule DataFrame
    # -----------------------------
    rows = []
    for day in days:
        row = {"Day": day}
        for slot in slots:
            for year in years:
                course_list = [f"{c['course_code']} ({c['students']} students)" for c in schedule[day][slot][year]]
                row[f"{slot} - {year}"] = ", ".join(course_list) if course_list else "Empty"
        for year in years:
            for branch in branches:
                row[f"Credits {year}-{branch}"] = day_year_credits[day][year][branch]
        for slot in slots:
            row[f"Total Students - {slot}"] = day_slot_total_students[day][slot]
        rows.append(row)
    
    df_schedule = pd.DataFrame(rows)
    
    # -----------------------------
    # Step 9 & 10: Room Allocation, Faculty Assignment
    # -----------------------------
    student_book = pd.read_excel("students.xlsx", sheet_name=None)
    student_ids = {}
    
    for sheet_name, df in student_book.items():
        year_key = normalize_year(sheet_name.replace(" Year", ""))
        for branch in ["CSE", "DSAI", "ECE"]:
            ids = df[branch].dropna().astype(str).tolist()
            student_ids[(year_key, branch)] = ids
    
    room_names = df_room["Room"].dropna().astype(str).tolist()
    faculty_count = len(faculty_list)
    faculty_index = 0
    df_rooms_rows = []
    
    for day in days:
        used_faculty_today = set()
        for slot in slots:
            assigned_faculty = []
            for _ in range(total_rooms):
                while faculty_list[faculty_index % faculty_count] in used_faculty_today:
                    faculty_index += 1
                assigned_faculty.append(faculty_list[faculty_index % faculty_count])
                used_faculty_today.add(faculty_list[faculty_index % faculty_count])
                faculty_index += 1
            room_faculty_mapping = {room_names[i]: assigned_faculty[i] for i in range(total_rooms)}
            
            room_courses = {room: [] for room in room_names}
            room_students = {room: [] for room in room_names}
            
            courses_in_slot = []
            for year in years:
                courses_in_slot.extend(schedule[day][slot][year])
            
            for course in courses_in_slot:
                branch = course["branch"]
                year_label = normalize_year(course["year"])
                ids = student_ids.get((year_label, branch), [])
                batch_size = room_capacity_per_course
                
                if not ids:
                    student_ranges = [f"{course['students']} students"]
                else:
                    student_ranges = [
                        f"{ids[i]}–{ids[min(i + batch_size - 1, len(ids) - 1)]}"
                        for i in range(0, len(ids), batch_size)
                    ]
                
                for sr in student_ranges:
                    target_room = None
                    for r in room_names:
                        if len(room_courses[r]) < 2 and course["course_code"] not in room_courses[r]:
                            target_room = r
                            break
                    if not target_room:
                        continue
                    
                    room_courses[target_room].append(course["course_code"])
                    room_students[target_room].append(sr)
            
            for r in room_names:
                for i, c in enumerate(room_courses[r]):
                    df_rooms_rows.append({
                        "Day": day,
                        "Slot": slot,
                        "Course": c,
                        "Branch": "".join([x for x in c if not x.isdigit()])[:4],
                        "Students": room_students[r][i],
                        "Rooms Assigned": r,
                        "Faculty": room_faculty_mapping[r]
                    })
    
    df_rooms = pd.DataFrame(df_rooms_rows)
    
    # Merge courses in same room
    merged_rows = []
    room_dict = defaultdict(list)
    
    for row in df_rooms_rows:
        key = (row["Day"], row["Slot"], row["Rooms Assigned"], row["Faculty"])
        course_info = f"{row['Course']} ({row['Students']})"
        room_dict[key].append(course_info)
    
    for key, course_list in room_dict.items():
        day, slot, room, faculty = key
        merged_rows.append({
            "Day": day,
            "Slot": slot,
            "Rooms Assigned": room,
            "Faculty": faculty,
            "Courses + Students": ", ".join(course_list)
        })
    
    df_rooms_merged = pd.DataFrame(merged_rows)
    
    # -----------------------------
    # Step 11: Save to Excel
    # -----------------------------
    output_file = "exam_schedule_with_rooms_faculty.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_schedule.to_excel(writer, sheet_name="Exam Schedule", index=False)
        for day in days:
            df_day = df_rooms[df_rooms["Day"]==day]
            df_day.to_excel(writer, sheet_name=f"Rooms-{day}", index=False)
    
    # -----------------------------
    # Step 12: Excel Formatting
    # -----------------------------
    wb = load_workbook(output_file)
    ws_main = wb["Exam Schedule"]
    
    for i, col in enumerate(ws_main.columns, start=1):
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws_main.column_dimensions[get_column_letter(i)].width = max_length + 5
    
    for row in ws_main.iter_rows(min_row=2, max_row=ws_main.max_row):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws_main.row_dimensions[cell.row].height = 30
    
    branch_colors = {
        "CSE": "FFC7CE",
        "DSAI": "C6EFCE",
        "ECE": "FFEB9C",
        "All": "BDD7EE"
    }
    
    for row in ws_main.iter_rows(min_row=2, max_row=ws_main.max_row):
        for cell in row:
            for branch, color in branch_colors.items():
                if branch in str(cell.value):
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                    cell.font = Font(bold=True)
    
    wb.save(output_file)
    
    # -----------------------------
    # Step 13: Create Free Slot Sheet
    # -----------------------------
    free_rows = []
    
    for day in days:
        for slot in slots:
            rooms_filled = df_rooms[(df_rooms["Day"]==day) & (df_rooms["Slot"]==slot)]["Rooms Assigned"].unique().tolist()
            available_rooms = [r for r in room_names if r not in rooms_filled]
            
            total_students_in_slot = day_slot_total_students[day][slot]
            remaining_capacity = max_students_per_slot - total_students_in_slot
            
            for year in years:
                for branch in branches:
                    courses_assigned = schedule[day][slot][year]
                    assigned_branches = [c["branch"] for c in courses_assigned]
                    has_common_course = any(c["branch"] == "All" for c in courses_assigned)
                    is_free = branch not in assigned_branches and not has_common_course
                    
                    free_rows.append({
                        "Day": day,
                        "Slot": slot,
                        "Year": year,
                        "Branch": branch,
                        "Status": "Free" if is_free else "Engaged",
                        "Available Rooms": ", ".join(available_rooms) if year==years[0] and branch==branches[0] else "",
                        "Remaining Capacity": remaining_capacity if year==years[0] and branch==branches[0] else ""
                    })
    
    df_free_slots = pd.DataFrame(free_rows)
    
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        df_free_slots.to_excel(writer, sheet_name="Free Slots", index=False)
    
    # Final formatting for all sheets
    wb = load_workbook(output_file)
    branch_colors = {
        "CSE": "FFC7CE",
        "DSAI": "C6EFCE",
        "ECE": "FFEB9C",
        "ALL": "BDD7EE"
    }
    
    sheet_names = wb.sheetnames
    
    for sheet_name in sheet_names:
        ws = wb[sheet_name]
        
        for i, col in enumerate(ws.columns, start=1):
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
            ws.column_dimensions[get_column_letter(i)].width = max_length + 5
        
        for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row_cells:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws.row_dimensions[cell.row].height = 30
        
        for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row_cells:
                cell_value_upper = str(cell.value).upper()
                for branch, color in branch_colors.items():
                    if branch in cell_value_upper:
                        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                        cell.font = Font(bold=True)
    
    wb.save(output_file)
    print(f"Exam timetable generated successfully: {output_file}")
    
    return output_file


# Test function to verify module can be imported
if __name__ == "__main__":
    print("ExamTimeTable module loaded successfully!")
    print("Available function: generate_timetable(start_date, end_date, branch_slot_allocation)")
    
    # Example usage (commented out):
    # result = generate_timetable(
    #     start_date='2025-09-26',
    #     end_date='2025-10-06',
    #     branch_slot_allocation={
    #         "1St Year": {"Morning": ["CSE", "DSAI", "ECE"], "Evening": []},
    #         "2Nd Year": {"Morning": ["DSAI"], "Evening": ["ECE", "CSE"]},
    #         "3Rd Year": {"Morning": ["CSE", "DSAI"], "Evening": ["ECE"]},
    #         "4Th Year": {"Morning": [], "Evening": []}
    #     }
    # )
    # print(f"Generated file: {result}")