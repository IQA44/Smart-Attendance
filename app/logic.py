
from datetime import datetime
from app.storage import (
    load_stage_department_attendance,
    save_stage_department_attendance,
    load_stage_department_students,
    save_stage_department_students
)

# ==================================================
# تسجيل الحضور / الانصراف
# ==================================================

def record_attendance(student, operation=None):
    stage = student["stage"]
    department = student["department"]

    stage_department_attendance = load_stage_department_attendance(stage, department)

    student_key = f"{student['name']}|{stage}|{department}"
    current_time = datetime.now().strftime("%I:%M %p")
    current_date = datetime.now().strftime("%Y-%m-%d")

    if student_key not in stage_department_attendance:
        stage_department_attendance[student_key] = {}

    if current_date not in stage_department_attendance[student_key]:
        stage_department_attendance[student_key][current_date] = []

    if operation is None:
        if not stage_department_attendance[student_key][current_date]:
            op = "حضور"
        else:
            last_action = stage_department_attendance[student_key][current_date][-1]['type']
            op = "انصراف" if last_action == "حضور" else "حضور"
    else:
        op = operation

    stage_department_attendance[student_key][current_date].append({
        "type": op,
        "time": current_time
    })

    save_stage_department_attendance(stage, department, stage_department_attendance)

# ==================================================
# البحث عن طالب داخل مرحلة وقسم
# ==================================================

def search_students(stage, department, search_text):
    students = load_stage_department_students(stage, department)

    if not search_text:
        return students

    search_text = search_text.lower()
    return [s for s in students if search_text in s["name"].lower()]


def move_students(
    source_stage,
    source_department,
    target_stage,
    target_department,
    student_names,
    card_students
):
    if source_stage == target_stage and source_department == target_department:
        return False, "لا يمكن النقل إلى نفس المرحلة والقسم"

    source_students = load_stage_department_students(source_stage, source_department)
    target_students = load_stage_department_students(target_stage, target_department)

    moved = []
    skipped_duplicates = []
    target_names = {s["name"] for s in target_students}

    for name in student_names:
        student_obj = None
        for s in source_students:
            if s["name"] == name:
                student_obj = s
                break

        if student_obj:
            if name in target_names:
                skipped_duplicates.append(name)
                continue

            target_students.append({
                "name": name,
                "stage": target_stage,
                "department": target_department
            })
            moved.append(name)
            target_names.add(name)

            
            for uid, data in card_students.items():
                if (
                    data["name"] == name
                    and data["stage"] == source_stage
                    and data["department"] == source_department
                ):
                    card_students[uid]["stage"] = target_stage
                    card_students[uid]["department"] = target_department

    if not moved:
        if skipped_duplicates:
            return False, "جميع الطلاب المحددين موجودون مسبقًا في المرحلة/القسم الهدف"
        return False, "لم يتم نقل أي طالب"

    source_students = [s for s in source_students if s["name"] not in moved]

    save_stage_department_students(source_stage, source_department, source_students)
    save_stage_department_students(target_stage, target_department, target_students)

    return True, f"تم نقل {len(moved)} طالب بنجاح"

# ==================================================
# حساب أيام الغياب          
# ==================================================

def calculate_absence_days(stage, department, student_name, records_folder):
    import os
    import pandas as pd

    class_folder = os.path.join(records_folder, stage, department)

    if not os.path.exists(class_folder):
        return 0, 0

    excel_files = [f for f in os.listdir(class_folder) if f.endswith(".xlsx")]
    total_days = len(excel_files)

    absence_count = 0

    for file in excel_files:
        try:
            df = pd.read_excel(os.path.join(class_folder, file))
            row = df[df["اسم"] == student_name]
            if not row.empty:
                if row["الحضور"].values[0] == "غياب":
                    absence_count += 1
        except Exception:
            pass

    return absence_count, total_days

