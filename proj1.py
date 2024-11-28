import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Border, Side

# Helper functions
def compute_max_capacity(room_size, arrangement_mode, seat_margin):
    """
    Calculate the maximum number of students that can be seated in a room.
    """
    if arrangement_mode == 'dense':
        return max(0, room_size - seat_margin)
    elif arrangement_mode == 'sparse':
        return max(1, (room_size // 2) - seat_margin)

def create_attendance_file(attendance_details, exam_day, course_id, room_id, session, roll_name_mapping):
    """
    Generate an attendance sheet as an Excel file for a specific exam session.
    """
    attendance_details["Student Name"] = attendance_details["Roll"].map(roll_name_mapping).fillna("Unknown Name")
    attendance_details["Signature"] = ""

    # Add blank rows for invigilator and TA signatures
    blank_rows = pd.DataFrame({"Roll": [""] * 5, "Student Name": [""] * 5, "Signature": [""] * 5})
    attendance_details = pd.concat([attendance_details, blank_rows], ignore_index=True)

    # Create an Excel workbook
    file_name = f"{exam_day.strftime('%d_%m_%Y')}_{course_id}_{room_id}_{session.lower()}.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = f"{course_id} Room {room_id}"

    # Write data to the sheet
    for record in dataframe_to_rows(attendance_details, index=False, header=True):
        worksheet.append(record)

    # Adjust column widths
    worksheet.column_dimensions["A"].width = 15  # Roll Number column
    worksheet.column_dimensions["B"].width = max(15, attendance_details["Student Name"].str.len().max() + 2)  # Name column
    worksheet.column_dimensions["C"].width = 20  # Signature column

    # Apply formatting
    border_style = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )
    for row in worksheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_style

    workbook.save(file_name)
    print(f"Attendance sheet created: {file_name}")

# Main function
def main():
    # Load input data
    input_file = 'proj1.xlsx'
    excel_data = pd.ExcelFile(input_file)

    # Get user input for seat margin and arrangement mode
    seat_margin = int(input("Enter margin seats per room: "))
    arrangement_mode = input("Choose seating arrangement ('dense' or 'sparse'): ").strip().lower()
    if arrangement_mode not in ["dense", "sparse"]:
        print("Invalid input. Defaulting to 'dense'.")
        arrangement_mode = "dense"

    # Load data sheets
    student_records = pd.read_excel(excel_data, sheet_name='ip_1', skiprows=1)
    schedule_details = pd.read_excel(excel_data, sheet_name='ip_2', skiprows=1)
    room_details = pd.read_excel(excel_data, sheet_name='ip_3')
    roll_name_mapping_data = pd.read_excel(excel_data, sheet_name='ip_4')

    # Roll number to name mapping
    roll_name_mapping = dict(zip(roll_name_mapping_data["Roll"].astype(str), roll_name_mapping_data["Name"]))

    # Prepare course details
    course_enrollment = student_records.groupby('course_code')['rollno'].count().reset_index()
    course_enrollment.columns = ['course_code', 'student_count']
    course_to_students = student_records.groupby("course_code")["rollno"].apply(list).to_dict()

    # Convert schedule dates to datetime
    schedule_details['Date'] = pd.to_datetime(schedule_details['Date'], dayfirst=True)
    block_9_rooms = room_details[room_details['Block'] == 9].sort_values(by=["Room No."])
    lt_block_rooms = room_details[room_details['Block'] == 'LT'].sort_values(by='Exam Capacity', ascending=False)

    # Prepare final seating plan
    seating_plan = pd.DataFrame()

    # Allocate rooms for each session
    for _, session_info in schedule_details.iterrows():
        exam_date = session_info['Date']
        for session in ['Morning', 'Evening']:
            if pd.isna(session_info[session]):
                continue

            session_courses = session_info[session].split('; ')
            course_student_counts = {
                course: course_enrollment[course_enrollment['course_code'] == course]['student_count'].values[0]
                if course in course_enrollment['course_code'].values else 0
                for course in session_courses
            }
            sorted_courses = sorted(course_student_counts.items(), key=lambda x: x[1], reverse=True)

            allocation_plan = []
            for course, num_students in sorted_courses:
                assigned_rooms = []
                remaining_students = num_students

                # Assign Block 9 rooms
                for _, room_info in block_9_rooms.iterrows():
                    if remaining_students <= 0:
                        break
                    room_capacity = room_info['Exam Capacity']
                    max_capacity = compute_max_capacity(room_capacity, arrangement_mode, seat_margin)
                    allocation_count = min(max_capacity, remaining_students)
                    remaining_students -= allocation_count
                    assigned_rooms.append({'course_code': course, 'room': room_info['Room No.'], 'allocated': allocation_count})

                # Assign LT block rooms if needed
                for _, room_info in lt_block_rooms.iterrows():
                    if remaining_students <= 0:
                        break
                    room_capacity = room_info['Exam Capacity']
                    max_capacity = compute_max_capacity(room_capacity, arrangement_mode, seat_margin)
                    allocation_count = min(max_capacity, remaining_students)
                    remaining_students -= allocation_count
                    assigned_rooms.append({'course_code': course, 'room': room_info['Room No.'], 'allocated': allocation_count})

                allocation_plan.extend(assigned_rooms)

            # Update seating plan
            for allocation in allocation_plan:
                course_id = allocation['course_code']
                room_id = allocation['room']
                allocated_count = allocation['allocated']

                if course_id in course_to_students:
                    allocated_students = course_to_students[course_id][:allocated_count]
                    course_to_students[course_id] = course_to_students[course_id][allocated_count:]

                    student_rolls = "; ".join(allocated_students)

                    seating_plan = pd.concat([seating_plan, pd.DataFrame({
                        'Date': [exam_date],
                        'Session': [session],
                        'Course_Code': [course_id],
                        'Room': [room_id],
                        'Allocated_Students': [allocated_count],
                        'Students': [student_rolls]
                    })], ignore_index=True)

    # Save seating arrangement
    seating_plan.to_excel('py_project.xlsx', index=False)
    print("Seating arrangement saved as 'final_seating_arrangement.xlsx'.")

    # Create attendance sheets
    for _, seating_row in seating_plan.iterrows():
        exam_day = seating_row["Date"]
        session = seating_row["Session"]
        course_id = seating_row["Course_Code"]
        room_id = seating_row["Room"]
        student_rolls = seating_row["Students"].split("; ")

        attendance_details = pd.DataFrame({"Roll": student_rolls})
        create_attendance_file(attendance_details, exam_day, course_id, room_id, session, roll_name_mapping)

    print("All attendance sheets created successfully.")

if __name__ == "__main__":
    main()

