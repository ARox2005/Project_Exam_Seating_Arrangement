# **Exam Seating Arrangement and Attendance Sheet Generator**

## **Project Overview**
This project automates the process of generating seating arrangements for exams and attendance sheets for invigilators, ensuring optimal usage of room capacities while adhering to constraints like building locality, buffer seats, and student distribution. The program processes data from input Excel sheets, allocates rooms dynamically, and outputs the results in both CSV and Excel formats.

---

## **Features**
### **Exam Seating Arrangement**
- **Dynamic Room Allocation**:
  - Prioritizes **Block 9** rooms before using **LT rooms** for seating allocation.
  - Supports user-defined **buffer size** per room to leave some seats vacant.
  - Offers two modes for allocation:
    - **Dense Mode**: Fills rooms as much as possible.
    - **Sparse Mode**: Limits allocation to 50% of room capacity for better student distribution.
  - Ensures no overlap or exceeding of room capacity during allocation.
  
- **Optimized Room Usage**:
  - Allocates larger courses first to minimize the number of rooms used.
  - Prefers rooms closer to each other for exams scheduled on the same date.

- **Error Handling**:
  - Raises an alert if the number of students exceeds the total available capacity.

### **Outputs**
1. **Seating Arrangement (op_1)**:
   - Contains columns: `Date`, `Day`, `Room`, `Allocated_students_count`, `Morning`, `Evening`, and `Roll_list` (with roll numbers of students separated by semicolons).

2. **Room Utilization Summary (op_2)**:
   - Includes all columns from the room capacity sheet (`Room No.`, `Exam Capacity`, `Block`) and an additional column, `Vacant`, showing remaining seats after allocation.

3. **Formats**:
   - Both outputs are generated in **CSV** and **Excel** formats, with the Excel file containing separate sheets for `op_1` and `op_2`.

---

## **Input Files**
The program processes an Excel file with the following sheets:

1. **ip_1**:
   - Columns: `rollno` (student roll numbers), `register_sem` (semester of registration), `schedule_sem` (exam semester), `course_code` (course enrolled).

2. **ip_2**:
   - Columns: `Date`, `Day`, `Morning` (course codes of morning exams separated by semicolons), `Evening` (course codes of evening exams separated by semicolons).

3. **ip_3**:
   - Columns: `Room No.` (room numbers), `Exam Capacity` (capacity of each room), `Block` (building of the room, e.g., `9` or `LT`).

4. **ip_4**:
   - Columns: `Roll` (student roll numbers), `Name` (student names).

---

## **Output Files**
1. **op_1.csv / Excel Sheet (Seating Arrangement)**:
   - **Columns**:
     - `Date`: Date of the exam.
     - `Day`: Day of the exam (e.g., Monday, Tuesday).
     - `Room`: Room number where the students are seated.
     - `Allocated_students_count`: Number of students seated in the room.
     - `Morning`: Course code for the morning slot (or `NO EXAM` if no exam is scheduled).
     - `Evening`: Course code for the evening slot (or `NO EXAM` if no exam is scheduled).
     - `Roll_list`: Semicolon-separated list of roll numbers allocated to the room.

2. **op_2.csv / Excel Sheet (Room Utilization Summary)**:
   - **Columns**:
     - `Room No.`: Room number.
     - `Exam Capacity`: Total capacity of the room.
     - `Block`: Building of the room (e.g., `9`, `LT`).
     - `Vacant`: Remaining seats in the room after allocation.


---

## **How It Works**
1. **Input Parsing**:
   - The program reads the input Excel file, processes the sheets, and converts them to CSV for easier manipulation.

2. **Seating Allocation**:
   - Dynamically allocates students to rooms based on the exam timetable (`ip_2`) and room availability (`ip_3`).
   - Ensures constraints like buffer size, building locality, and room capacity are followed.

3. **Output Generation**:
   - Generates seating arrangements and room utilization summaries.
   - Saves outputs in both CSV and Excel formats for convenience.

---

## **Setup**
### **Prerequisites**
- Python 3.7 or later
- Required libraries:
  - `pandas`
  - `openpyxl`
  - `collections` (built-in)

### **Installation**
1. Clone the repository:
   ```bash
   git clone https://github.com/<your-username>/exam-seating-arrangement.git
   cd exam-seating-arrangement
   ```

2. Install dependencies:
   ```bash
   pip install pandas openpyxl
   ```

### **Usage**
1. Place your input Excel file in the project directory.
2. Update the `input_excel_path` variable in the code with the path to your input file.
3. Run the program:
   ```bash
   python seating_arrangement.py
   ```
4. Find the outputs (`op_1.csv`, `op_2.csv`, `output.xlsx`) in the project directory.

---

## **Customization**
1. **Buffer Size**:
   - Adjust the `buffer_size` variable to control the number of unoccupied seats in each room.

2. **Dense vs Sparse Allocation**:
   - Set the `dense_mode` variable:
     - `True` for filling rooms completely.
     - `False` for limiting allocation to 50% of room capacity.

---

**End of description**
