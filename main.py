import pandas as pd
import math
from collections import defaultdict

# Input Excel file with sheets
input_excel_path = "input_excel_worksheet.xlsx"

# Output file paths
op_1_path_csv = "op_1.csv"
op_2_path_csv = "op_2.csv"
op_excel_path = "output.xlsx"

# Parameters
buffer_size = 5  # Adjustable buffer per room
dense_mode = True  # Dense if True, Sparse otherwise

# Load Excel sheets
excel_data = pd.ExcelFile(input_excel_path)
ip_1 = excel_data.parse("ip_1", skiprows=1)
ip_2 = excel_data.parse("ip_2", skiprows=1)
ip_3 = excel_data.parse("ip_3")
ip_4 = excel_data.parse("ip_4")

# Prepare input data
ip_1["rollno"] = ip_1["rollno"].astype(str)
course_students = ip_1.groupby("course_code")["rollno"].apply(list).to_dict()
ip_3_sorted = ip_3.sort_values(by=["Block", "Room No."])  # Sort rooms by Block and Room

# Initialize output structures
op_1 = []
op_2 = ip_3.copy()
op_2["Vacant"] = op_2["Exam Capacity"]

# Track remaining capacity for rooms by date and slot
room_usage = defaultdict(lambda: defaultdict(dict))  # {date: {slot: {room_no: remaining_capacity}}}

# Initialize room capacities for all dates and slots
for _, room in ip_3.iterrows():
    for date in ip_2["Date"].unique():
        for slot in ["Morning", "Evening"]:
            room_usage[date][slot][room["Room No."]] = room["Exam Capacity"] - buffer_size

# Helper function for room allocation
def allocate_students(course_code, students, date, slot, room_data, dense, buffer):
    allocated = []
    remaining = students[:]
    for _, room in room_data.iterrows():
        room_no = room["Room No."]
        remaining_capacity = room_usage[date][slot][room_no]

        if not dense:
            remaining_capacity = math.floor(remaining_capacity / 2)  # Sparse allocation

        num_to_allocate = min(len(remaining), remaining_capacity)
        allocated_students = remaining[:num_to_allocate]
        remaining = remaining[num_to_allocate:]

        if num_to_allocate > 0:
            allocated.append({
                "Room": room_no,
                "Allocated_students_count": len(allocated_students),
                "Roll_list": ";".join(allocated_students),
            })
            # Update remaining capacity
            room_usage[date][slot][room_no] -= len(allocated_students)

        if not remaining:
            break

    return allocated, remaining

# Allocate students based on timetable
for _, row in ip_2.iterrows():
    date, day = row["Date"], row["Day"]
    for slot in ["Morning", "Evening"]:
        courses = row[slot].split("; ") if pd.notna(row[slot]) else ["NO EXAM"]
        for course_code in courses:
            if course_code not in course_students:
                continue

            students = course_students[course_code]
            block_9_rooms = ip_3_sorted[ip_3_sorted["Block"] == 9]
            lt_rooms = ip_3_sorted[ip_3_sorted["Block"] == "LT"]

            # First, allocate in Block 9 rooms
            allocations, remaining = allocate_students(course_code, students, date, slot, block_9_rooms, dense_mode, buffer_size)

            # If students remain, allocate in LT rooms
            if remaining:
                additional_allocations, remaining = allocate_students(course_code, remaining, date, slot, lt_rooms, dense_mode, buffer_size)
                allocations.extend(additional_allocations)

            # If students still remain, raise an error
            if remaining:
                raise ValueError(f"Could not allocate all students for course {course_code} on {date} in {slot} slot.")

            for alloc in allocations:
                op_1.append({
                    "Date": date,
                    "Day": day,
                    "Morning": course_code if slot == "Morning" else "NO EXAM",
                    "Evening": course_code if slot == "Evening" else "NO EXAM",
                    **alloc,
                })

# Prepare op_1 DataFrame
op_1_df = pd.DataFrame(op_1)

# Reorder columns to move Roll_list to the last column
columns = ["Date", "Day", "Room", "Allocated_students_count", "Morning", "Evening", "Roll_list"]
op_1_df = op_1_df[columns]

# Save outputs as CSV
op_1_df.to_csv(op_1_path_csv, sep=";", index=False)
op_2.to_csv(op_2_path_csv, sep=";", index=False)

# Save outputs as Excel file with separate sheets
with pd.ExcelWriter(op_excel_path, engine="openpyxl") as writer:
    op_1_df.to_excel(writer, sheet_name="op_1", index=False)
    op_2.to_excel(writer, sheet_name="op_2", index=False)

print("Seating arrangement completed and outputs generated.")
